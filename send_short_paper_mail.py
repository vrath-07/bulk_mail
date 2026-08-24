"""Short-paper decision mails for INDIS 2026.

Kept separate from send_mail.py on purpose: different source files, different
templates, and a different decision vocabulary (ACCEPTED / REJECTED rather than
the meta-reviewer wording). The reading, recipient and sending machinery is
imported from send_mail.py so both flows behave identically, but the preview
folder and the sent log are separate so one run can never disturb the other.

Data:
  Short_Paper_ Submissions.xlsx - one worksheet per track, author emails
  Short_Paper_Summary.xlsx      - one flat sheet: Paper ID, Decision, Reviews

A paper is mailed only when it appears in BOTH files.
"""
import csv
import html
import os
import re
import smtplib
import time
from collections import defaultdict
from datetime import datetime
from difflib import SequenceMatcher
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from send_mail import (
    EMAIL_PASS,
    EMAIL_USER,
    SMTP_PORT,
    SMTP_SERVER,
    find_column,
    note,
    paper_sort_key,
    read_worksheets,
    require_column,
    resolve_recipients,
    split_header,
)

# Nothing is sent while this is True: every mail is written to PREVIEW_DIR
# instead and SMTP is never contacted. Flip to False only for the real run.
DRY_RUN = False
PREVIEW_DIR = "preview_short"

# Separate from the full-paper log, so the two campaigns never mask each other.
SENT_LOG = "sent_log_short.csv"
REQUIRE_CONFIRMATION = True

# Note the space in the submissions filename - it is spelled that way in the
# export. Rename the file and this constant together if you want it tidied.
SUBMISSIONS_FILE = "official/Short_Paper_Submissions.xlsx"
SUMMARY_FILE = "official/Short_Paper_Summary.xlsx"

FINAL_PAPER_LINK = "https://forms.gle/f3yuPCC6mgLenCBX9"
REGISTRATION_DEADLINE = "25 August 2026"
CAMERA_READY_DEADLINE = "25 August 2026"

SIGN_OFF = """
<p>
Warm regards,<br><br>
<strong>INDIS 2026 Organising Committee</strong><br>
Department of Design<br>
Indian Institute of Technology Guwahati
</p>
"""

# Decision value (Short_Paper_Summary.xlsx, "Decision") -> the letter to send.
# Keys are upper-cased, and the lookup upper-cases too, so "Accepted" matches.
# Each template is split around the reviewer comments, which the docx places
# mid-letter in the acceptance mail and near the end in the rejection.
DECISIONS = {
    "ACCEPTED": {
        "subject": ("INDIS 2026 – Short Paper Acceptance, Author Registration "
                    "& Poster Presentation"),
        "before_reviews": f"""
<p>
We are pleased to inform you that your short paper submitted to
<strong>INDIS 2026 – International Conference on Design and Innovation
Studies</strong> has been <strong>accepted for presentation and
publication</strong>, subject to the final submission requirements outlined
below.
</p>

<p>
The conference will be held from
<strong>28–30 September 2026 at IIT Guwahati</strong>, under the theme
<strong>&ldquo;Design Meets Technology and Business.&rdquo;</strong>
</p>

<p><strong>Important Requirements for Accepted Short Papers</strong></p>

<p><strong>1. Author Registration – by {REGISTRATION_DEADLINE}</strong></p>

<p>
At least <strong>one author of the paper must register as an Author by
{REGISTRATION_DEADLINE}</strong>. The registered author will be required to
present the paper in the form of a <strong>poster during INDIS 2026</strong>.
</p>

<p><strong>2. Final Camera-Ready Paper – by {CAMERA_READY_DEADLINE}</strong></p>

<p>
Please upload the revised <strong>camera-ready version</strong> of your short
paper through the link provided below:
</p>

<p>
<strong>Final Paper Submission:</strong>
<a href="{FINAL_PAPER_LINK}">{FINAL_PAPER_LINK}</a>
</p>

<p>
While preparing the final version, authors must carefully incorporate the
<strong>reviewer comments provided below</strong>. Please also ensure that all
figures, photographs, diagrams, and other visual elements are of
<strong>sufficiently high resolution and quality for print
reproduction</strong>.
</p>

<p><strong>Reviewer Comments:</strong></p>
""",
        "after_reviews": """
<p><strong>3. Conference Publication</strong></p>

<p>
The accepted short papers will be included in the
<strong>INDIS 2026 Conference Book</strong>, to be published by
<strong>Bloomsbury</strong> with an <strong>ISBN</strong>.
</p>

<p><strong>4. Poster Presentation</strong></p>

<p>
All accepted short papers will be presented as
<strong>posters during INDIS 2026</strong>. The posters will be displayed at the
conference venue, providing authors an opportunity to discuss their work with
researchers, designers, academics, and industry professionals.
</p>

<p>
The <strong>poster template, specifications, dimensions, and submission
guidelines will be shared with the authors shortly</strong>.
</p>

<p>
Please note that <strong>registration of at least one author and presentation of
the poster during the conference are mandatory requirements</strong> for
inclusion of the work in the conference publication.
</p>

<p>
We look forward to welcoming you to <strong>INDIS 2026 at IIT Guwahati</strong>
and to engaging with your work during the conference.
</p>
""",
    },
    "REJECTED": {
        # The docx gives no subject for the rejection letter; this one is
        # written to match the house style. Change it here if you prefer.
        "subject": "INDIS 2026 – Decision on Your Short Paper Submission",
        "before_reviews": """
<p>
Thank you for submitting your short paper to
<strong>INDIS 2026 – International Conference on Design and Innovation
Studies</strong>.
</p>

<p>
After careful review and consideration by the Conference Committee, we regret to
inform you that your submission <strong>has not been accepted for presentation at
INDIS 2026</strong>.
</p>

<p>
The decision was based on the reviewers&rsquo; assessments and the overall
evaluation of the submission. The reviewer comments are provided below for your
reference.
</p>

<p><strong>Reviewer Comments:</strong></p>
""",
        "after_reviews": """
<p>
We sincerely appreciate your interest in INDIS 2026 and the effort invested in
your submission.
</p>
""",
    },
}


# ---------------------------------------------------------------------------
# Reading the two short-paper files
# ---------------------------------------------------------------------------

def _title_key(text):
    """Letters and digits only, so case, spacing and punctuation drop out."""
    return re.sub(r"[^a-z0-9]+", "", text.lower())


def title_similarity(left, right):
    return SequenceMatcher(None, _title_key(left), _title_key(right)).ratio()


def titles_agree(left, right, threshold=0.9):
    """True when two titles are the same paper allowing for formatting drift."""
    a, b = _title_key(left), _title_key(right)
    if a == b or a.startswith(b) or b.startswith(a):
        return True
    return SequenceMatcher(None, a, b).ratio() >= threshold


def find_any_column(columns, *needles):
    """First column matching any needle, most specific needle first.

    Guards against a future export gaining e.g. a 'Reviewer Email' column that
    a bare 'Review' substring search would grab instead of the comments.
    """
    for needle in needles:
        try:
            return find_column(columns, needle)
        except KeyError:
            continue
    raise KeyError(f"no column matching any of {needles} in {sorted(columns)}")


def resolve_path(path):
    """Find the file even if the stray space in the export name is missing.

    The CMT download is called "Short_Paper_ Submissions.xlsx" - with a space
    before "Submissions" - which anyone re-creating the file by hand naturally
    drops. Match on the name ignoring spaces and case so either spelling works.
    """
    if os.path.exists(path):
        return path
    directory, name = os.path.split(path)
    directory = directory or "."
    if not os.path.isdir(directory):
        return path
    wanted = name.replace(" ", "").lower()
    for candidate in os.listdir(directory):
        if candidate.replace(" ", "").lower() == wanted:
            return os.path.join(directory, candidate)
    return path


def require_files():
    """Fail with a readable message, naming what is actually in the folder."""
    for label, path in (("submissions", SUBMISSIONS_FILE), ("summary", SUMMARY_FILE)):
        resolved = resolve_path(path)
        if not os.path.exists(resolved):
            directory = os.path.dirname(path) or "."
            present = (
                sorted(os.listdir(directory)) if os.path.isdir(directory)
                else [f"(no such directory: {directory})"]
            )
            raise FileNotFoundError(
                f"short-paper {label} file not found at {path!r}.\n"
                f"  {directory} contains: {present}\n"
                f"  Fix the name, or update SUBMISSIONS_FILE / SUMMARY_FILE "
                f"at the top of this file."
            )


def load_submissions():
    """-> (Paper ID -> {primary, authors, title}, Paper ID -> track, warnings).

    One worksheet per track, same shape as the other CMT exports: a track-title
    row, a blank row, then the header.
    """
    submissions, tracks, warnings = {}, {}, []
    for rows in read_worksheets(resolve_path(SUBMISSIONS_FILE)):
        track, columns, data = split_header(rows)
        if not columns:
            continue
        paper_id_col = require_column(columns, "Paper ID")
        primary_col = find_column(columns, "Primary Contact Author Email")
        authors_col = find_column(columns, "Author Emails")
        title_col = columns.get("Paper Title")
        for row in data:
            paper_id = row.get(paper_id_col, "")
            if not paper_id:
                continue
            if paper_id in submissions:
                warnings.append(note(
                    paper_id, track,
                    "duplicate Paper ID in Submissions - the last row's authors are used",
                ))
            submissions[paper_id] = {
                "primary": row.get(primary_col, ""),
                "authors": [
                    email.strip().rstrip("*").strip()
                    for email in row.get(authors_col, "").split(";")
                    if email.strip()
                ],
                "title": row.get(title_col, "") if title_col else "",
            }
            tracks[paper_id] = track
    return submissions, tracks, warnings


def load_summary():
    """-> (Paper ID -> {decision, review, title}, warnings).

    A single flat sheet covering every track, header on the first row.
    """
    summary, warnings = {}, []
    for rows in read_worksheets(resolve_path(SUMMARY_FILE)):
        _, columns, data = split_header(rows)
        if not columns:
            continue
        paper_id_col = require_column(columns, "Paper ID")
        decision_col = find_any_column(columns, "Decision")
        review_col = find_any_column(columns, "Reviews", "Review", "Comments")
        title_col = columns.get("Paper Title")
        for row in data:
            paper_id = row.get(paper_id_col, "")
            if not paper_id:
                continue
            decision = row.get(decision_col, "")
            if paper_id in summary:
                previous = summary[paper_id]["decision"]
                # Compare normalised, so 'Accepted' vs 'ACCEPTED' is not a clash.
                clash = (
                    f"CONFLICTING decisions {previous!r} vs {decision!r}"
                    if previous.strip().upper() != decision.strip().upper()
                    else f"same decision {decision!r}"
                )
                warnings.append(note(
                    paper_id, "",
                    f"duplicate Paper ID in Summary - {clash}; the last row is used",
                ))
            summary[paper_id] = {
                "decision": decision,
                "review": row.get(review_col, ""),
                "title": row.get(title_col, "") if title_col else "",
            }
    return summary, warnings


def build_records():
    """Join the two files.

    Returns (records, skipped, incomplete, warnings):
      records    - in both files and mailable
      skipped    - in both files, but unmailable (unknown decision, no address)
      incomplete - in only one file, so never mailed
      warnings   - mailable, but worth a look
    """
    require_files()
    submissions, tracks, warnings = load_submissions()
    summary, summary_warnings = load_summary()
    warnings = warnings + summary_warnings

    def track_of(paper_id):
        return tracks.get(paper_id) or "(unknown track)"

    # load_summary() has no track to work with, so fill them in now.
    for warning in warnings:
        if not warning["track"]:
            warning["track"] = track_of(warning["paper_id"])

    records, skipped, title_mismatches = [], [], []
    common = set(submissions) & set(summary)

    for paper_id in sorted(common, key=paper_sort_key):
        track = track_of(paper_id)
        entry = summary[paper_id]
        decision = DECISIONS.get(entry["decision"].strip().upper())

        if decision is None:
            reason = (
                "no decision recorded" if not entry["decision"].strip()
                else f"decision {entry['decision']!r} has no template"
            )
            skipped.append(note(paper_id, track, reason))
            continue

        to, cc, invalid = resolve_recipients(submissions[paper_id])
        if invalid:
            warnings.append(note(
                paper_id, track, f"dropped malformed address(es): {', '.join(invalid)}",
            ))
        if not to:
            skipped.append(note(paper_id, track, "no usable author email address"))
            continue

        if not entry["review"].strip():
            warnings.append(note(
                paper_id, track,
                'no reviewer comments - the "Reviewer Comments" section would be empty',
            ))

        # Every difference between the two titles is recorded, because a title
        # that disagrees may mean the id no longer points at the same paper in
        # both files - which would mail the wrong decision. Differences are
        # graded so formatting drift does not hide genuine divergence.
        submitted_title = submissions[paper_id]["title"].strip()
        summary_title = entry["title"].strip()
        if submitted_title and summary_title and submitted_title != summary_title:
            title_mismatches.append({
                "paper_id": paper_id,
                "track": track,
                "submissions_title": submitted_title,
                "summary_title": summary_title,
                "similarity": title_similarity(submitted_title, summary_title),
                "substantive": not titles_agree(submitted_title, summary_title),
            })

        # Submissions is the authoritative CMT export of what the author typed;
        # Summary is a working sheet that carries placeholders like "NO TITLE".
        title = submitted_title or summary_title
        if not title:
            warnings.append(note(paper_id, track, "no Paper Title in either file"))

        records.append({
            "paper_id": paper_id,
            "title": title,
            "track": track,
            "decision_value": entry["decision"].strip().upper(),
            "decision": decision,
            "review": entry["review"],
            "to": to,
            "cc": cc,
        })

    present_in = {"Submissions": set(submissions), "Summary": set(summary)}
    incomplete = []
    for paper_id in sorted((set(submissions) | set(summary)) - common, key=paper_sort_key):
        found = [name for name, ids in present_in.items() if paper_id in ids]
        missing = [name for name in present_in if name not in found]
        incomplete.append({
            "paper_id": paper_id,
            "track": track_of(paper_id),
            "found_in": found,
            "missing_from": missing,
            "reason": "missing from " + ", ".join(missing),
        })

    return records, skipped, incomplete, warnings, title_mismatches


# ---------------------------------------------------------------------------
# Building the mail
# ---------------------------------------------------------------------------

def format_review(text):
    """Escape the verbatim reviewer comments and keep their line breaks."""
    return html.escape(text).replace("\n", "<br>")


def create_message(record):
    decision = record["decision"]

    msg = MIMEMultipart("mixed")
    msg["From"] = EMAIL_USER
    msg["To"] = ", ".join(record["to"])
    if record["cc"]:
        msg["Cc"] = ", ".join(record["cc"])
    # One author has five short papers, so the id has to be on the letter.
    msg["Subject"] = f"{decision['subject']} (Paper ID {record['paper_id']})"

    review_html = (
        f"""
<div style="border-left:3px solid #cfcfcf; padding-left:12px; margin-bottom:18px;">
{format_review(record["review"])}
</div>
"""
        if record["review"].strip() else ""
    )

    # Never render a dangling "Paper Title:" label with nothing after it.
    identifiers = f'<strong>Paper ID:</strong> {html.escape(record["paper_id"])}'
    if record["title"].strip():
        identifiers += f'<br>\n<strong>Paper Title:</strong> {html.escape(record["title"])}'

    body = f"""
<html>
<body style="margin:0; padding:0; font-family:Aptos, Calibri, Arial, sans-serif; font-size:12pt; color:#000000; line-height:1.5;">

<p>Dear Author(s),</p>

<p>
{identifiers}
</p>
{decision["before_reviews"]}
{review_html}
{decision["after_reviews"]}
{SIGN_OFF}
</body>
</html>
"""

    msg.attach(MIMEText(body, "html"))
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
    for record in records:
        msg = create_message(record)
        body = next(
            part for part in msg.walk() if part.get_content_type() == "text/html"
        ).get_payload(decode=True).decode("utf-8")
        header = (
            f"<div style='font-family:monospace; font-size:10pt; background:#f4f4f4;"
            f" padding:10px; margin-bottom:16px;'>"
            f"<b>To:</b> {html.escape(msg['To'])}<br>"
            f"<b>Cc:</b> {html.escape(msg['Cc'] or '-')}<br>"
            f"<b>Subject:</b> {html.escape(msg['Subject'])}</div>"
        )
        path = os.path.join(PREVIEW_DIR, f"paper_{record['paper_id']}.html")
        with open(path, "w", encoding="utf-8") as preview:
            preview.write(header + body)

    return write_report(
        "_summary.csv",
        [[r["paper_id"], r["decision_value"], len(r["review"]),
          "; ".join(r["to"]), "; ".join(r["cc"]), r["track"], r["title"]]
         for r in records],
        ["Paper ID", "Decision", "Review Chars", "To", "Cc", "Track", "Title"],
    )


def report_title_mismatches(mismatches):
    """List every paper whose title differs between the two files.

    The mailed title always comes from Submissions; this is purely a check that
    the same Paper ID still means the same paper in both files.
    """
    if not mismatches:
        print("\nTITLE CHECK: every Paper ID has matching titles in both files.")
        return []

    substantive = [m for m in mismatches if m["substantive"]]
    print(f"\nTITLE MISMATCHES ({len(mismatches)} paper(s); "
          f"{len(substantive)} substantially different):")
    print("  (the letter always uses the Submissions title)")
    for entry in sorted(mismatches, key=lambda m: m["similarity"]):
        flag = "DIFFERENT" if entry["substantive"] else "formatting"
        print(f"\n  paper {entry['paper_id']:>6s} | {entry['track']} "
              f"| {flag} (similarity {entry['similarity']:.2f})")
        print(f"     Submissions (used): {entry['submissions_title'][:96]}")
        print(f"     Summary           : {entry['summary_title'][:96]}")

    return [write_report(
        "_title_mismatches.csv",
        [[m["paper_id"], m["track"],
          "DIFFERENT" if m["substantive"] else "formatting",
          f"{m['similarity']:.2f}", m["submissions_title"], m["summary_title"]]
         for m in sorted(mismatches, key=lambda m: m["similarity"])],
        ["Paper ID", "Track", "Severity", "Similarity",
         "Submissions Title (used)", "Summary Title"],
    )]


def report_data_issues(skipped, incomplete, warnings):
    if incomplete:
        print(f"\nNOT IN BOTH FILES - not mailed ({len(incomplete)}):")
        grouped = defaultdict(list)
        for entry in incomplete:
            grouped[", ".join(entry["found_in"])].append(entry)
        for found, entries in grouped.items():
            print(f"  found only in {found} ({len(entries)}) -"
                  f" missing {', '.join(entries[0]['missing_from'])}")
            for entry in entries:
                print(f"     paper {entry['paper_id']:>6s} | {entry['track']}")

    if skipped:
        print(f"\nIN BOTH BUT NOT MAILABLE ({len(skipped)}):")
        for entry in skipped:
            print(f"  paper {entry['paper_id']:>6s} | {entry['track']} | {entry['reason']}")

    if warnings:
        print(f"\nWARNINGS - mailed, but check these ({len(warnings)}):")
        for entry in warnings:
            print(f"  paper {entry['paper_id']:>6s} | {entry['track']} | {entry['reason']}")

    paths = []
    if incomplete or skipped or warnings:
        paths.append(write_report(
            "_issues.csv",
            [[kind, e["paper_id"], e["track"], e["reason"]]
             for kind, group in (("not-in-both", incomplete), ("skipped", skipped),
                                 ("warning", warnings))
             for e in group],
            ["Kind", "Paper ID", "Track", "Detail"],
        ))
    return paths


def load_sent_log():
    if not os.path.exists(SENT_LOG):
        return {}
    with open(SENT_LOG, encoding="utf-8", newline="") as handle:
        return {row["Paper ID"]: row["Sent At"] for row in csv.DictReader(handle)}


def append_sent_log(record, msg):
    is_new = not os.path.exists(SENT_LOG)
    with open(SENT_LOG, "a", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        if is_new:
            writer.writerow(["Paper ID", "Sent At", "Decision", "To", "Cc"])
        writer.writerow([
            record["paper_id"],
            datetime.now().isoformat(timespec="seconds"),
            record["decision_value"],
            msg["To"],
            msg["Cc"] or "",
        ])


def send_emails():
    records, skipped, incomplete, warnings, title_mismatches = build_records()

    counts = defaultdict(int)
    for record in records:
        counts[record["decision_value"]] += 1
    recipients = {email.lower() for r in records for email in r["to"] + r["cc"]}

    print(f"Short papers in both files : {len(records)}")
    print(f"Unique recipients          : {len(recipients)}")
    for decision, count in sorted(counts.items()):
        print(f"  {decision:12s} {count:4d}")

    report_paths = report_data_issues(skipped, incomplete, warnings)
    report_paths += report_title_mismatches(title_mismatches)

    if DRY_RUN:
        summary_path = write_previews(records)
        print(f"\nDRY RUN - no mail sent. {len(records)} previews in {PREVIEW_DIR}/")
        for path in [summary_path] + report_paths:
            print(f"  {path}")
        return

    already_sent = load_sent_log()
    pending = [r for r in records if r["paper_id"] not in already_sent]
    if already_sent:
        print(f"\n{len(already_sent)} paper(s) already sent per {SENT_LOG} - skipping them.")
    if not pending:
        print("Nothing left to send.")
        return

    pending_recipients = {e.lower() for r in pending for e in r["to"] + r["cc"]}
    print(f"\nAbout to send {len(pending)} REAL short-paper emails to "
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

    sent, failed = 0, []
    for record in pending:
        try:
            msg = create_message(record)
            server.send_message(msg)
            append_sent_log(record, msg)
            sent += 1
            print(f"Sent paper {record['paper_id']} to {msg['To']} (cc: {msg['Cc'] or '-'})")
        except Exception as exc:
            failed.append(note(record["paper_id"], record["track"], str(exc)))
            print(f"FAILED paper {record['paper_id']}: {exc}")
        time.sleep(2)

    try:
        server.quit()
    except Exception:
        pass

    print(f"\nSent {sent}, failed {len(failed)}.")
    if failed:
        for entry in failed:
            print(f"  paper {entry['paper_id']:>6s} | {entry['track']} | {entry['reason']}")
        print(write_report(
            "_failed.csv",
            [[e["paper_id"], e["track"], e["reason"]] for e in failed],
            ["Paper ID", "Track", "Error"],
        ))
        print(f"Re-running will retry only these - {SENT_LOG} records what succeeded.")


if __name__ == "__main__":
    send_emails()
