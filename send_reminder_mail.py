"""Author-registration reminder for INDIS 2026.

One short reminder to the authors of every accepted paper, read from the
Consolidated workbook. Separate preview folder and sent log from the other two
campaigns, so none of them can mask another.

The letter carries no paper-specific text, so it is addressed per recipient
rather than per paper: an author with five accepted papers gets one reminder,
not five identical ones.
"""
import csv
import html
import os
import smtplib
import time
from collections import defaultdict
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from send_mail import (
    EMAIL_PASS,
    EMAIL_PATTERN,
    EMAIL_USER,
    SMTP_PORT,
    SMTP_SERVER,
    find_column,
    note,
    paper_sort_key,
    read_worksheets,
    require_column,
    split_header,
)

# Nothing is sent while this is True: every mail is written to PREVIEW_DIR
# instead and SMTP is never contacted. Flip to False only for the real run.
DRY_RUN = False

# Previews, the sent log and every report all live together under here.
REMINDER_LOG_DIR = "reminder_logs"
PREVIEW_DIR = REMINDER_LOG_DIR
SENT_LOG = os.path.join(REMINDER_LOG_DIR, "sent_log_reminder.csv")
REQUIRE_CONFIRMATION = True

CONSOLIDATED_FILE = "official/Consolidated.xlsx"

# Only short papers get this reminder. Matched case-insensitively against the
# "Paper Type" column. An accepted paper whose type is neither this nor a full
# paper is reported rather than dropped in silence.
SHORT_PAPER_TYPES = {"short paper"}
FULL_PAPER_TYPES = {"full paper"}

# Only these recommendations get a reminder. Compared case-insensitively after
# stripping, so "ACCEPTED" and "accepted" both match.
ACCEPTED_DECISIONS = {
    "accept",
    "accept with major revisions",
    "accept with minor revisions",
    "accepted",
}

REGISTRATION_DEADLINE = "25 August 2026"
CONFERENCE_DATES = "28–30 September 2026"

# The reminder text has no subject line of its own; this one is written to
# match it. Change it here if you would rather word it differently.
SUBJECT = f"INDIS 2026 – Reminder: Author Registration by {REGISTRATION_DEADLINE}"

BODY = f"""
<html>
<body style="margin:0; padding:0; font-family:Aptos, Calibri, Arial, sans-serif; font-size:12pt; color:#000000; line-height:1.5;">

<p>Dear Authors,</p>

<p>
A gentle reminder to kindly complete your author registration by
{REGISTRATION_DEADLINE} to confirm your participation and secure your
paper&rsquo;s presentation slot at INDIS 2026.
</p>

<p>
Please note that this deadline applies to short papers, and
{REGISTRATION_DEADLINE} is a hard deadline for completing author registration.
No extensions may be possible beyond this date.
</p>

<p>
We would also like to emphasize that the accepted papers need to be submitted
for the final proceedings within the stipulated timeline. If you are unable to
complete the registration and submit the required paper within the deadline, we
may have to reallocate your presentation slot and finalize the conference
programme and proceedings submission.
</p>

<p>
We therefore request you to complete the registration and submission process at
the earliest to avoid any inconvenience.
</p>

<p>
We look forward to welcoming you to IIT Guwahati from {CONFERENCE_DATES}.
</p>

<p>
Warm regards,<br>
<strong>INDIS 2026 Organising Committee</strong><br>
Department of Design, IIT Guwahati
</p>

</body>
</html>
"""


# ---------------------------------------------------------------------------
# Reading the consolidated workbook
# ---------------------------------------------------------------------------

def consolidated_rows():
    """Yield (columns, data rows) for the one sheet that holds the paper list.

    The workbook has several sheets and three of them start with a 'Paper ID'
    header, so the sheet is chosen by its full column signature rather than by
    position or name - 'Single Review' even has a 'Q3 (Recommendation)' column
    that a looser test would happily accept.
    """
    matches = []
    for rows in read_worksheets(CONSOLIDATED_FILE):
        _, columns, data = split_header(rows)
        if not columns:
            continue
        if "Paper ID" in columns and any(
            "primary contact author email" in name.lower() for name in columns
        ):
            matches.append((columns, data))

    if not matches:
        raise KeyError(
            f"no sheet in {CONSOLIDATED_FILE!r} has both a 'Paper ID' and a "
            "'Primary Contact Author Email' column"
        )
    if len(matches) > 1:
        raise KeyError(
            f"{len(matches)} sheets in {CONSOLIDATED_FILE!r} look like the paper "
            "list; expected exactly one"
        )
    return matches[0]


def normalise_address(email):
    """Return (address, change) with any international domain punycoded.

    SMTP needs an ASCII domain, so a non-ASCII one is converted via IDNA -
    "example.संगठन" becomes "example.xn--i1b6b1a6a2e". A non-ASCII local part
    cannot be converted without SMTPUTF8, so it is passed through and recorded;
    if the server refuses it, the per-message error handling catches it rather
    than losing the whole batch.
    """
    if email.isascii() or "@" not in email:
        return email, None

    local, _, domain = email.rpartition("@")
    change = None
    if not domain.isascii():
        try:
            domain = domain.encode("idna").decode("ascii")
            change = f"international domain converted for SMTP - {domain}"
        except (UnicodeError, UnicodeDecodeError):
            return email, "international domain could not be converted - may be rejected"
    if not local.isascii():
        change = "non-ASCII characters before the @ - may be rejected by the mail server"
    return f"{local}@{domain}", change


def load_papers():
    """-> (accepted paper records, skipped, warnings, decision counts)."""
    if not os.path.exists(CONSOLIDATED_FILE):
        directory = os.path.dirname(CONSOLIDATED_FILE) or "."
        # Only spreadsheets, and only a handful - listing a whole Downloads
        # folder back at someone helps nobody.
        nearby = sorted(
            name for name in (os.listdir(directory) if os.path.isdir(directory) else [])
            if name.lower().endswith((".xlsx", ".xls", ".xlsm", ".csv"))
        )
        hint = ", ".join(nearby[:10]) + (" ..." if len(nearby) > 10 else "")
        raise FileNotFoundError(
            f"consolidated workbook not found at {CONSOLIDATED_FILE!r}.\n"
            f"  spreadsheets in {directory}: {hint or '(none)'}\n"
            f"  Fix the name, or update CONSOLIDATED_FILE at the top of this file."
        )

    columns, data = consolidated_rows()
    paper_id_col = require_column(columns, "Paper ID")
    primary_col = find_column(columns, "Primary Contact Author Email")
    decision_col = find_column(columns, "Recommendation")
    track_col = columns.get("Track")
    title_col = columns.get("Paper Title")
    type_col = columns.get("Paper Type")
    if type_col is None:
        # Without it every accepted paper would look like a short paper.
        raise KeyError(
            f"no 'Paper Type' column in {CONSOLIDATED_FILE!r} - cannot tell short "
            "papers from full papers, so refusing to guess who to remind"
        )
    # Who the decision mail went to. Blank for papers not yet mailed, in which
    # case the reminder goes to the primary contact alone.
    sent_to_col = next(
        (index for name, index in columns.items() if "sent to" in name.lower()), None
    )

    papers, skipped, warnings = [], [], []
    counts = defaultdict(int)
    seen_ids = set()
    # Every address that was dropped, merged or otherwise not used verbatim.
    address_log = []
    # Accept-like decisions that do not match ACCEPTED_DECISIONS exactly.
    unrecognised_accepts = defaultdict(list)
    # Paper types among accepted papers, and the ones we could not classify.
    type_counts = defaultdict(int)
    unrecognised_types = defaultdict(list)
    # lower-cased address -> the one spelling actually mailed.
    spellings = {}

    for row in data:
        paper_id = row.get(paper_id_col, "")
        if not paper_id:
            continue

        decision = row.get(decision_col, "").strip()
        counts[decision or "(blank)"] += 1
        if decision.lower() not in ACCEPTED_DECISIONS:
            # A new accept-like wording would otherwise be dropped in silence,
            # and those authors would never learn they had to register.
            if "accept" in decision.lower():
                unrecognised_accepts[decision].append(row.get(paper_id_col, ""))
            continue

        # Accepted - now keep only the short papers.
        paper_type = (row.get(type_col, "") if type_col else "").strip()
        type_counts[paper_type or "(blank)"] += 1
        if paper_type.lower() not in SHORT_PAPER_TYPES:
            if paper_type.lower() not in FULL_PAPER_TYPES:
                # Neither "Short Paper" nor "Full Paper" - could be a short
                # paper under another name, so say so instead of dropping it.
                unrecognised_types[paper_type or "(blank)"].append(
                    row.get(paper_id_col, "")
                )
            continue

        track = row.get(track_col, "") if track_col else ""
        if paper_id in seen_ids:
            warnings.append(note(
                paper_id, track, "duplicate Paper ID in the consolidated sheet",
            ))
        seen_ids.add(paper_id)

        # The column is comma-separated, but semicolons also show up. Splitting
        # and trimming is parsing, not correction, so neither that nor the '*'
        # primary marker is reported as a change to the address.
        raw_addresses = [row.get(primary_col, "")]
        if sent_to_col is not None:
            raw_addresses += row.get(sent_to_col, "").replace(";", ",").split(",")

        # The primary contact is also listed inside "Email(s) Sent To" by
        # design, so that repeat is expected and not worth reporting.
        primary_key = normalise_address(
            raw_addresses[0].strip().rstrip("*").strip()
        )[0].lower()

        seen, ordered, invalid = set(), [], []
        for position, raw in enumerate(raw_addresses):
            email = raw.strip().rstrip("*").strip()
            if not email:
                continue

            original = email
            email, idn_change = normalise_address(email)
            key = email.lower()

            if key in seen:
                # Already handled on this paper. The primary contact also
                # appearing inside "Email(s) Sent To" is how the export is
                # built, so that repeat is expected; anything else is noted.
                if not (position > 0 and key == primary_key):
                    address_log.append({
                        "paper_id": paper_id, "raw": email, "used": email,
                        "action": "listed more than once on this paper - used once",
                    })
                continue
            seen.add(key)

            # Logged only once the address is actually taken, so the primary
            # reappearing in "Email(s) Sent To" does not report it twice.
            if idn_change:
                address_log.append({
                    "paper_id": paper_id, "raw": original, "used": email,
                    "action": idn_change,
                })

            if not EMAIL_PATTERN.match(email):
                invalid.append(email)
                address_log.append({
                    "paper_id": paper_id, "raw": email, "used": "",
                    "action": "dropped - not a valid email address",
                })
                continue

            # Keep one spelling per person so they are not mailed twice. Record
            # it whenever the spelling on this row is not the one that is used.
            canonical = spellings.setdefault(key, email)
            if email != canonical:
                address_log.append({
                    "paper_id": paper_id, "raw": email, "used": canonical,
                    "action": "same address in different case - merged",
                })
            ordered.append(canonical)

        if invalid:
            warnings.append(note(
                paper_id, track, f"dropped malformed address(es): {', '.join(invalid)}",
            ))
        if not ordered:
            skipped.append(note(paper_id, track, "no usable author email address"))
            address_log.append({
                "paper_id": paper_id, "raw": "(none on this paper)", "used": "",
                "action": "paper has no usable address - nobody mailed for it",
            })
            continue

        papers.append({
            "paper_id": paper_id,
            "title": row.get(title_col, "") if title_col else "",
            "track": track,
            "decision": decision,
            "to": ordered[:1],
            "cc": ordered[1:],
        })

    return (papers, skipped, warnings, counts, address_log,
            type_counts, unrecognised_types)


def build_records():
    """One reminder per unique person, addressed individually.

    Two requirements pull against each other: nobody may be mailed twice, and
    no author or co-author of an accepted paper may be missed. Grouping by
    paper fails the first (an author with five accepted papers gets five
    copies); grouping by lead author fails it too, because the same person can
    be primary contact on one paper and a co-author on another.

    The only arrangement that satisfies both is to flatten every address across
    every accepted paper into a set and send each address its own mail. That
    also means no Cc, so nobody's address is exposed to anyone else.
    """
    (papers, skipped, warnings, counts, address_log,
     type_counts, unrecognised_types) = load_papers()

    people = {}
    for paper in papers:
        for address in paper["to"] + paper["cc"]:
            key = address.lower()
            if key not in people:
                # Keep the spelling as first seen; the key does the matching.
                people[key] = {"to": [address], "cc": [], "papers": {}}
            # Keyed by paper id so one person listed twice on a paper still
            # produces a single row per paper in the log.
            people[key]["papers"][paper["paper_id"]] = {
                "paper_id": paper["paper_id"],
                "track": paper["track"],
                "title": paper["title"],
                "decision": paper["decision"],
            }

    records = []
    for person in people.values():
        person["papers"] = sorted(
            person["papers"].values(), key=lambda p: paper_sort_key(p["paper_id"])
        )
        person["paper_ids"] = [p["paper_id"] for p in person["papers"]]
        person["tracks"] = {p["track"] for p in person["papers"] if p["track"]}
        records.append(person)
    records.sort(key=lambda record: record["to"][0].lower())

    return (records, papers, skipped, warnings, counts, address_log,
            type_counts, unrecognised_types)


def write_address_log(address_log, records, sent_addresses=None):
    """Write 'error in mail id.csv': every address not used exactly as given.

    Says for each one whether the person still receives the reminder, and if
    not, why. sent_addresses is None on a dry run - nothing has been sent, so
    the column reports what would happen instead of what did.
    """
    mailed = {record["to"][0].lower() for record in records}
    rows = []
    for entry in sorted(address_log, key=lambda e: paper_sort_key(e["paper_id"])):
        target = (entry["used"] or entry["raw"]).lower()
        reaches_someone = bool(entry["used"]) and target in mailed

        if not reaches_someone:
            status, reason = "No", entry["action"]
        elif sent_addresses is None:
            status, reason = "Not yet - dry run", f"would be mailed as {entry['used']}"
        elif target in sent_addresses:
            status, reason = "Yes", f"mailed as {entry['used']}"
        else:
            status, reason = "No", "send failed or skipped - see sent_log_reminder.csv"

        rows.append([entry["paper_id"], entry["raw"], entry["action"],
                     entry["used"] or "-", status, reason])

    return write_report(
        "error in mail id.csv", rows,
        ["Paper ID", "Original Mail ID", "What Happened", "Mail ID Used",
         "Mail Sent", "Reason"],
    )


# ---------------------------------------------------------------------------
# Building the mail
# ---------------------------------------------------------------------------

def create_message(record):
    msg = MIMEMultipart("mixed")
    msg["From"] = EMAIL_USER
    msg["To"] = ", ".join(record["to"])
    if record["cc"]:
        msg["Cc"] = ", ".join(record["cc"])
    msg["Subject"] = SUBJECT
    msg.attach(MIMEText(BODY, "html"))
    return msg


# ---------------------------------------------------------------------------
# Dry run / sending
# ---------------------------------------------------------------------------

def write_report(name, rows, fieldnames):
    os.makedirs(PREVIEW_DIR, exist_ok=True)
    path = os.path.join(PREVIEW_DIR, name)
    with open(path, "w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(fieldnames)
        writer.writerows(rows)
    return path


def write_previews(records):
    os.makedirs(PREVIEW_DIR, exist_ok=True)
    for index, record in enumerate(records, start=1):
        msg = create_message(record)
        body = next(
            part for part in msg.walk() if part.get_content_type() == "text/html"
        ).get_payload(decode=True).decode("utf-8")
        header = (
            f"<div style='font-family:monospace; font-size:10pt; background:#f4f4f4;"
            f" padding:10px; margin-bottom:16px;'>"
            f"<b>To:</b> {html.escape(msg['To'])}<br>"
            f"<b>Cc:</b> {html.escape(msg['Cc'] or '-')}<br>"
            f"<b>Subject:</b> {html.escape(msg['Subject'])}<br>"
            f"<b>Covers papers:</b> {', '.join(record['paper_ids'])}</div>"
        )
        name = f"{index:03d}_{record['to'][0].replace('@', '_at_')}.html"
        with open(os.path.join(PREVIEW_DIR, name), "w", encoding="utf-8") as preview:
            preview.write(header + body)

    # Same shape as the real sent log, so a dry run shows exactly what will be
    # recorded: one row per (recipient, paper).
    return write_report(
        "_summary.csv",
        [[r["to"][0], p["paper_id"], p["decision"], p["track"], p["title"],
          len(r["papers"])]
         for r in records for p in r["papers"]],
        ["Recipient", "Paper ID", "Decision", "Track", "Paper Title",
         "Papers For This Recipient"],
    )


def report(records, papers, skipped, warnings, counts,
           type_counts=None, unrecognised_types=None):
    print("Decisions in the consolidated sheet:")
    for decision, count in sorted(counts.items(), key=lambda item: -item[1]):
        if decision.lower() in ACCEPTED_DECISIONS:
            mark = "-> reminder"
        elif "accept" in decision.lower():
            mark = "*** LOOKS ACCEPTED BUT IS NOT IN ACCEPTED_DECISIONS - NO REMINDER"
        else:
            mark = ""
        print(f"  {count:4d}  {decision!r:32s} {mark}")

    if type_counts:
        print("\nPaper types among accepted papers:")
        for paper_type, count in sorted(type_counts.items(), key=lambda i: -i[1]):
            if paper_type.lower() in SHORT_PAPER_TYPES:
                mark = "-> reminder"
            elif paper_type.lower() in FULL_PAPER_TYPES:
                mark = "(full paper - no reminder)"
            else:
                mark = "*** UNRECOGNISED TYPE - NO REMINDER, CHECK THIS"
            print(f"  {count:4d}  {paper_type!r:28s} {mark}")

    if unrecognised_types:
        print("\nAccepted papers whose type is neither Short nor Full - not mailed:")
        for paper_type, ids in unrecognised_types.items():
            print(f"  {paper_type!r}: papers {', '.join(ids)}")

    print(f"\nAccepted SHORT papers  : {len(papers)}")
    print(f"Reminders to send      : {len(records)} (one per person)")

    # Prove both guarantees rather than assuming them.
    addresses = [r["to"][0].lower() for r in records]
    repeats = sorted({a for a in addresses if addresses.count(a) > 1})
    everyone = {e.lower() for p in papers for e in p["to"] + p["cc"]}
    missed = sorted(everyone - set(addresses))
    uncovered = sorted(
        {p["paper_id"] for p in papers}
        - {pid for r in records for pid in r["paper_ids"]},
        key=paper_sort_key,
    )
    print(f"  anyone mailed twice  : {repeats or 'no'}")
    print(f"  anyone missed        : {missed or 'no'}")
    print(f"  accepted papers with no recipient: {uncovered or 'none'}")

    multi = [r for r in records if len(r["paper_ids"]) > 1]
    if multi:
        print(f"\nPeople on more than one accepted paper ({len(multi)}) - "
              f"one reminder each, covering all their papers:")
        for record in sorted(multi, key=lambda r: -len(r["paper_ids"]))[:15]:
            print(f"  {record['to'][0]:48s} papers {', '.join(record['paper_ids'])}")
        if len(multi) > 15:
            print(f"  ... and {len(multi) - 15} more (see {PREVIEW_DIR}/_summary.csv)")

    if skipped:
        print(f"\nACCEPTED BUT NOT MAILABLE ({len(skipped)}):")
        for entry in skipped:
            print(f"  paper {entry['paper_id']:>6s} | {entry['track']} | {entry['reason']}")

    if warnings:
        print(f"\nWARNINGS ({len(warnings)}):")
        for entry in warnings:
            print(f"  paper {entry['paper_id']:>6s} | {entry['track']} | {entry['reason']}")

    paths = []
    if skipped or warnings:
        paths.append(write_report(
            "_issues.csv",
            [[kind, e["paper_id"], e["track"], e["reason"]]
             for kind, group in (("skipped", skipped), ("warning", warnings))
             for e in group],
            ["Kind", "Paper ID", "Track", "Detail"],
        ))
    return paths


def load_sent_log():
    """Reminders already sent, keyed by the To address."""
    if not os.path.exists(SENT_LOG):
        return {}
    with open(SENT_LOG, encoding="utf-8", newline="") as handle:
        return {row["To"].lower(): row["Sent At"] for row in csv.DictReader(handle)}


def append_sent_log(record, msg):
    """One row per (recipient, paper) - a person with four accepted papers
    appears four times, so the log can be read paper-by-paper."""
    os.makedirs(os.path.dirname(SENT_LOG) or ".", exist_ok=True)
    is_new = not os.path.exists(SENT_LOG)
    timestamp = datetime.now().isoformat(timespec="seconds")
    with open(SENT_LOG, "a", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        if is_new:
            writer.writerow(["To", "Paper ID", "Sent At", "Decision", "Track",
                             "Paper Title", "Papers For This Recipient"])
        for paper in record["papers"]:
            writer.writerow([
                record["to"][0],
                paper["paper_id"],
                timestamp,
                paper["decision"],
                paper["track"],
                paper["title"],
                len(record["papers"]),
            ])


def send_emails():
    (records, papers, skipped, warnings, counts, address_log,
     type_counts, unrecognised_types) = build_records()
    report_paths = report(records, papers, skipped, warnings, counts,
                          type_counts, unrecognised_types)

    if address_log:
        dropped = sum(1 for e in address_log if not e["used"])
        print(f"\nMAIL ID PROBLEMS ({len(address_log)}; {dropped} address(es) dropped) "
              f"- full list in {PREVIEW_DIR}/error in mail id.csv:")
        for entry in sorted(address_log, key=lambda e: paper_sort_key(e["paper_id"])):
            print(f"  paper {entry['paper_id']:>6s} | {entry['raw'][:44]:46s} | {entry['action']}")

    if DRY_RUN:
        summary_path = write_previews(records)
        report_paths.append(write_address_log(address_log, records))
        print(f"\nDRY RUN - no mail sent. {len(records)} previews in {PREVIEW_DIR}/")
        for path in [summary_path] + report_paths:
            print(f"  {path}")
        return

    already_sent = load_sent_log()
    pending = [r for r in records if r["to"][0].lower() not in already_sent]
    if already_sent:
        print(f"\n{len(already_sent)} recipient(s) already reminded per {SENT_LOG} - skipping.")
    if not pending:
        print("Nothing left to send.")
        return

    pending_recipients = {e.lower() for r in pending for e in r["to"] + r["cc"]}
    print(f"\nAbout to send {len(pending)} REAL reminder emails to "
          f"{len(pending_recipients)} recipients, as {EMAIL_USER}.")
    if REQUIRE_CONFIRMATION:
        try:
            answer = input("Type SEND to proceed (anything else aborts): ").strip()
        except EOFError:
            answer = ""
        if answer != "SEND":
            print("Aborted - nothing sent.")
            return

    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASS)
    except Exception as exc:
        print(f"SMTP connection/login failed, nothing sent: {exc}")
        return

    sent, failed, sent_addresses = 0, [], set()
    for record in pending:
        try:
            msg = create_message(record)
            server.send_message(msg)
            append_sent_log(record, msg)
            sent_addresses.add(record["to"][0].lower())
            sent += 1
            print(f"Sent to {msg['To']} (cc: {msg['Cc'] or '-'}) "
                  f"for paper(s) {', '.join(record['paper_ids'])}")
        except Exception as exc:
            failed.append(note(record["paper_ids"][0], record["to"][0], str(exc)))
            print(f"FAILED {record['to'][0]}: {exc}")
        time.sleep(2)

    try:
        server.quit()
    except Exception:
        pass

    # Anything already sent in an earlier run still counts as delivered.
    sent_addresses |= set(already_sent)
    print(f"\nSent {sent}, failed {len(failed)}.")
    if address_log:
        print(write_address_log(address_log, records, sent_addresses))
    if failed:
        for entry in failed:
            print(f"  {entry['track']} | {entry['reason']}")
        print(write_report(
            "_failed.csv",
            [[e["paper_id"], e["track"], e["reason"]] for e in failed],
            ["First Paper ID", "To", "Error"],
        ))
        print(f"Re-running will retry only these - {SENT_LOG} records what succeeded.")


if __name__ == "__main__":
    send_emails()
