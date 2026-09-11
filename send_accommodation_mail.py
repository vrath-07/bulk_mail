"""Accommodation notice for INDIS 2026 participants.

A separate script because the source data is a different shape entirely: a plain
Full Name / Email Address list with no Paper ID, decision or paper type. The
reminder script hard-errors without those columns, and relaxing that guard would
weaken the protection that stops the paper campaigns mass-mailing the wrong set.

This one also carries a PDF attachment, which none of the others do.
"""
import csv
import html
import os
import smtplib
import time
from collections import defaultdict
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from send_mail import (
    EMAIL_PASS,
    EMAIL_PATTERN,
    EMAIL_USER,
    SMTP_PORT,
    SMTP_SERVER,
    read_worksheets,
    split_header,
)
from send_reminder_mail import normalise_address

# Nothing is sent while this is True: every mail is written to PREVIEW_DIR
# instead and SMTP is never contacted. Flip to False only for the real run.
DRY_RUN = False

LOG_DIR = "accommodation_logs"
PREVIEW_DIR = LOG_DIR
SENT_LOG = os.path.join(LOG_DIR, "sent_log_accommodation.csv")
REQUIRE_CONFIRMATION = True

SOURCE_FILE = "official/accomodation.xlsx"

# The letter says the photographs are attached, so a missing file is a hard
# error rather than a mail that promises an attachment it does not carry.
ATTACHMENTS = ["official/Ashroy Hostel.pdf"]

ACCOMMODATION_FORM_LINK = ("https://docs.google.com/forms/d/e/"
                           "1FAIpQLSfvcmBhAwZSh-Kw4PXkTL8vhlEI0i01qFmEOYqbfkqMfpHmvg/viewform")

# The draft shows "Aashroy Hostel - IIT Guwahati" styled as a link but carries
# no URL behind it. Put the Maps URL here and it becomes a real link; left
# empty, the text is rendered plain and the run reports it as missing.
MAPS_LINK = "https://maps.app.goo.gl/396Yq585wnLWEzBo8"
MAPS_TEXT = "https://maps.app.goo.gl/396Yq585wnLWEzBo8"

SUBJECT = "INDIS 2026 – On-Campus Accommodation at Aashroy Hostel, IIT Guwahati"

RED = "#FF0000"
BLUE = "#0000FF"

HEADING = 'style="font-size:13pt; margin:18px 0 6px;"'

_maps = (f'Aashroy Hostel – IIT Guwahati <a href="{MAPS_LINK}">{MAPS_TEXT}</a>' if MAPS_LINK
         else f'<span style="color:{BLUE};">{MAPS_TEXT}</span>')

BODY = f"""
<html>
<body style="margin:0; padding:0; font-family:Aptos, Calibri, Arial, sans-serif; font-size:12pt; color:#000000; line-height:1.5;">

<p>Dear INDIS 2026 Participant,</p>

<p>Greetings from the INDIS 2026 Organising Committee.</p>

<p>
We are writing to provide an important update regarding your accommodation at
IIT Guwahati during <strong>INDIS 2026</strong>.
</p>

<h3 {HEADING}><span style="color:{RED};">On-Campus Accommodation Confirmed</span></h3>

<p>
We are pleased to inform you that <strong>on-campus accommodation at IIT
Guwahati has been secured for INDIS 2026 participants</strong>.
</p>

<p>
At the time of registration, it was communicated that accommodation in the IIT
Guwahati student hostels would be available at a rate of
<strong>&#8377;350 per person per day</strong>, excluding food. However, owing to
accommodation constraints and the unavailability of the originally proposed
hostel facilities for the conference dates, IIT Guwahati has provided an
alternative accommodation facility, <strong>Aashroy Hostel</strong>, for the
participants.
</p>

<p>The applicable rate for Aashroy Hostel is:</p>

<p>
<strong>&#8377;700 per person per day</strong><br>
<strong>Food charges: Not included</strong>
</p>

<p>
We regret the increase from the originally communicated rate of &#8377;350 per
day. This change is due to the change in the accommodation facility made
available by IIT Guwahati and is not a revision introduced by the INDIS 2026
Organising Committee.
</p>

<h3 {HEADING}>Aashroy Hostel &ndash; Photographs and Location</h3>

<p>
To help you make an informed decision before proceeding with the accommodation
payment, <strong>photographs of the Aashroy Hostel rooms and the hostel premises
are attached to this email</strong> for your reference.
</p>

<p>The location of Aashroy Hostel can also be viewed through Google Maps:</p>

<p><strong>Google Maps Location:</strong> {_maps}</p>

<p>
We strongly recommend that participants review the attached photographs and the
hostel location before making the payment, particularly because
<strong>the accommodation payment is non-refundable</strong>.
</p>

<h3 {HEADING}>Action Required for Participants</h3>

<p>
Participants who wish to avail themselves of the
<strong>Aashroy Hostel accommodation</strong> are required to:
</p>

<ul>
<li><strong>Fill in the accommodation form</strong> provided by the INDIS 2026 Organising Committee.</li>
<li><strong>Make the applicable accommodation payment</strong> for the required
    number of days.</li>
<li>Ensure that the information provided in the form, particularly the dates of
    stay and personal details, is accurate.</li>
</ul>

<strong>On-Campus Accommodation Form : </strong><a href="{ACCOMMODATION_FORM_LINK}">{ACCOMMODATION_FORM_LINK}</a>

<p>
<span style="color:{RED};"><strong>Accommodation will be confirmed only after
completion of the required form and payment.</strong></span>
</p>

<p>
Since the availability of on-campus accommodation is limited, participants who
intend to stay at Aashroy Hostel are requested to complete the process at the
earliest.
</p>

<h3 {HEADING}>Food Arrangements</h3>

<p>
The accommodation charge of <strong>&#8377;700 per day does not include
food</strong>.
</p>

<p>
Participants may independently arrange their meals through the various food
outlets available on the IIT Guwahati campus. The regular hostel messes are also
available, subject to their operating arrangements and access provisions.
</p>

<p>The currently applicable indicative mess rates are:</p>

<table style="border-collapse:collapse; margin:8px 0 16px; font-size:12pt;">
<tr>
  <th style="border:1px solid #cfcfcf; padding:6px 16px; text-align:left;">Meal</th>
  <th style="border:1px solid #cfcfcf; padding:6px 16px; text-align:left;">Rate</th>
</tr>
<tr>
  <td style="border:1px solid #cfcfcf; padding:6px 16px;">Breakfast</td>
  <td style="border:1px solid #cfcfcf; padding:6px 16px;">&#8377;60</td>
</tr>
<tr>
  <td style="border:1px solid #cfcfcf; padding:6px 16px;">Lunch</td>
  <td style="border:1px solid #cfcfcf; padding:6px 16px;">&#8377;75</td>
</tr>
<tr>
  <td style="border:1px solid #cfcfcf; padding:6px 16px;">Dinner</td>
  <td style="border:1px solid #cfcfcf; padding:6px 16px;">&#8377;100</td>
</tr>
</table>

<p>
Participants may therefore budget approximately <strong>&#8377;235 per
day</strong> if all three meals are taken at the regular hostel mess.
</p>

<p>
In addition to hostel messes, several other food outlets and eateries are
available on and around the IIT Guwahati campus, including commercial food
outlets such as <strong>Domino&rsquo;s and KFC</strong>, as well as other nearby
eating options.
</p>

<p>
For information regarding campus restaurants, food outlets, transportation,
campus navigation, bus services and other facilities, participants are
encouraged to install the <strong>IITG
<span style="color:{RED};">OneStop</span></strong> application.
</p>

<h3 {HEADING}>Important Hostel Rules and Regulations</h3>

<p>
Participants staying at Aashroy Hostel will be required to comply with the
<strong>rules, regulations, discipline requirements and instructions of IIT
Guwahati and the hostel authorities</strong> throughout their stay.
</p>

<p>Participants are specifically advised to observe the following:</p>

<ul>
<li>The accommodation is strictly for the participant to whom it has been
    allotted. Rooms/beds may not be transferred, exchanged, sublet or handed
    over to another person.</li>
<li>Unauthorized persons are not permitted to stay in the allotted
    accommodation.</li>
<li>Visitors must comply with the applicable visitor and security procedures of
    the hostel.</li>
<li>Participants must carry and produce valid identification when required and
    comply with all security and verification procedures.</li>
<li><strong>Smoking, consumption or possession of alcohol, narcotic substances
    or other prohibited substances within the hostel/residential premises is
    strictly prohibited.</strong></li>
<li>Drunken, disorderly, threatening or abusive behaviour will not be
    tolerated.</li>
<li>Excessive noise, loud music, shouting or any activity that causes
    disturbance to other residents is prohibited.</li>
<li>Any form of harassment, intimidation, physical altercation or threatening
    behaviour is strictly prohibited.</li>
<li>Damage to hostel property or Institute facilities is strictly prohibited.
    Participants may be held responsible for any damage caused.</li>
<li>Furniture and other hostel facilities must not be moved, removed or altered
    without permission.</li>
<li>Participants must maintain cleanliness and hygiene in their rooms and common
    areas.</li>
<li>Cooking or use of unauthorized heating/electrical appliances inside rooms is
    not permitted where prohibited by the hostel authorities.</li>
<li>No commercial activity, solicitation or unauthorized selling is permitted
    inside the hostel premises.</li>
<li>Pets or animals are not permitted inside the residential hostel
    facilities.</li>
<li>Participants must not enter restricted areas of the hostel.</li>
<li>Participants must follow all instructions issued by the hostel authorities,
    wardens, security personnel and Institute officials.</li>
<li>Participants must not interfere with the functioning of hostel staff or
    security personnel.</li>
<li>Participants must respect the privacy, safety and rights of other
    residents.</li>
<li>Participants must observe all applicable gender-specific access and
    visitation regulations of the residential facility.</li>
<li>Participants must vacate the accommodation within the period allotted to
    them. Extension of stay will be subject to approval by the competent
    authority.</li>
</ul>

<h3 {HEADING}>Disciplinary Action</h3>

<p>Please take the above regulations seriously.</p>

<p>
A participant found violating hostel rules or engaging in misconduct may be
subject to <strong>disciplinary action by the hostel/Institute
authorities</strong>. Depending on the nature and seriousness of the violation,
the participant may be required to <strong>vacate the accommodation immediately
and leave the hostel/campus</strong>.
</p>

<p>
In such circumstances, <strong>the INDIS 2026 Organising Committee will not be
responsible for arranging alternative accommodation</strong>, and
<strong>no refund of the accommodation amount will be provided</strong>.
</p>

<h3 {HEADING}>No-Refund Policy</h3>

<p>Please note the following very carefully before making payment:</p>

<p>
<strong>Once the accommodation payment has been made, the amount will be strictly
non-refundable under any circumstances.</strong>
</p>

<p>This includes, but is not limited to:</p>

<ul>
<li>Change of travel plans;</li>
<li>Cancellation of participation;</li>
<li>Failure to travel to IIT Guwahati;</li>
<li>Late arrival or early departure;</li>
<li>Change in dates of stay;</li>
<li>Personal or medical reasons;</li>
<li>Visa/travel-related issues;</li>
<li>Failure to attend the conference; or</li>
<li>Removal from the hostel due to violation of hostel rules.</li>
</ul>

<p>
Participants are therefore requested to <strong>carefully determine their
required dates of stay before making the payment</strong>.
</p>

<h3 {HEADING}>Important</h3>

<p>
The Aashroy Hostel facility is being provided to facilitate convenient on-campus
accommodation for INDIS 2026 participants. However, accommodation remains
subject to the rules and operational requirements of IIT Guwahati and the hostel
administration.
</p>

<p>
By submitting the accommodation form and making the payment, the participant will
be deemed to have <strong>accepted the applicable hostel rules, disciplinary
requirements and the above no-refund policy</strong>.
</p>

<p>
We strongly recommend that participants review the <strong>attached photographs
and Google Maps location of Aashroy Hostel</strong> before making the payment.
</p>

<p>
We look forward to welcoming you to <strong>IIT Guwahati for INDIS 2026</strong>
and hope that your stay on campus will be comfortable, safe and academically
enriching.
</p>

<p>
With regards,<br>
<strong>INDIS 2026 Organising Committee</strong><br>
Indian Institute of Technology Guwahati<br>
Guwahati, Assam, India
</p>

</body>
</html>
"""


# ---------------------------------------------------------------------------
# Reading the participant list
# ---------------------------------------------------------------------------

def split_header_by_email(rows):
    """Locate the header row by the email column; the sheet has no Paper ID."""
    for position, row in enumerate(rows):
        if any("email" in str(value).lower() for value in row.values()):
            columns = {value: index for index, value in row.items() if value}
            return columns, rows[position + 1:]
    return {}, []


def require_inputs():
    """Fail with a readable message rather than a bare traceback."""
    if not os.path.exists(SOURCE_FILE):
        directory = os.path.dirname(SOURCE_FILE) or "."
        nearby = sorted(
            name for name in (os.listdir(directory) if os.path.isdir(directory) else [])
            if name.lower().endswith((".xlsx", ".xls", ".xlsm", ".csv"))
        )
        raise FileNotFoundError(
            f"participant list not found at {SOURCE_FILE!r}.\n"
            f"  spreadsheets in {directory}: {', '.join(nearby[:10]) or '(none)'}"
        )
    for path in ATTACHMENTS:
        if not os.path.exists(path):
            raise FileNotFoundError(
                f"attachment not found at {path!r} - the letter tells participants "
                "the photographs are attached, so refusing to send without it"
            )


def load_participants():
    """-> (records, skipped, address_log).

    One mail per unique address. The list has the same person more than once,
    sometimes under slightly different names, so it is deduplicated on the
    address and the first spelling of the name is kept.
    """
    require_inputs()

    matches = []
    for rows in read_worksheets(SOURCE_FILE):
        columns, data = split_header_by_email(rows)
        if columns:
            matches.append((columns, data))
    if not matches:
        raise KeyError(f"no sheet in {SOURCE_FILE!r} has an email column")
    if len(matches) > 1:
        raise KeyError(
            f"{len(matches)} sheets in {SOURCE_FILE!r} have an email column; "
            "expected exactly one"
        )
    columns, data = matches[0]

    email_col = next(index for name, index in columns.items()
                     if "email" in name.lower())
    name_col = next((index for name, index in columns.items()
                     if "name" in name.lower()), None)

    people, skipped, address_log = {}, [], []
    for position, row in enumerate(data, start=1):
        raw = row.get(email_col, "").strip()
        name = row.get(name_col, "").strip() if name_col is not None else ""
        if not raw:
            if name:
                skipped.append({"row": position, "name": name, "email": "",
                                "reason": "no email address in this row"})
            continue

        email, change = normalise_address(raw)
        if change:
            address_log.append({"row": position, "raw": raw, "used": email,
                                "action": change})

        if not EMAIL_PATTERN.match(email):
            skipped.append({"row": position, "name": name, "email": raw,
                            "reason": "not a valid email address"})
            address_log.append({"row": position, "raw": raw, "used": "",
                                "action": "dropped - not a valid email address"})
            continue

        key = email.lower()
        if key in people:
            previous = people[key]["name"]
            detail = (f"listed again as {name!r}" if name and name != previous
                      else "listed more than once")
            address_log.append({"row": position, "raw": raw, "used": email,
                                "action": f"duplicate - {detail}; mailed once"})
            continue

        people[key] = {"name": name, "to": [email], "rows": [position]}

    records = sorted(people.values(), key=lambda p: p["to"][0].lower())
    return records, skipped, address_log


# ---------------------------------------------------------------------------
# Building the mail
# ---------------------------------------------------------------------------

def create_message(record):
    msg = MIMEMultipart("mixed")
    msg["From"] = EMAIL_USER
    msg["To"] = record["to"][0]
    msg["Subject"] = SUBJECT
    msg.attach(MIMEText(BODY, "html"))

    for path in ATTACHMENTS:
        with open(path, "rb") as handle:
            part = MIMEApplication(handle.read(), _subtype="pdf")
        part.add_header("Content-Disposition", "attachment",
                        filename=os.path.basename(path))
        msg.attach(part)
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
    attachments = ", ".join(os.path.basename(p) for p in ATTACHMENTS)
    for index, record in enumerate(records, start=1):
        header = (
            f"<div style='font-family:monospace; font-size:10pt; background:#f4f4f4;"
            f" padding:10px; margin-bottom:16px;'>"
            f"<b>To:</b> {html.escape(record['to'][0])}<br>"
            f"<b>Name:</b> {html.escape(record['name'] or '-')}<br>"
            f"<b>Subject:</b> {html.escape(SUBJECT)}<br>"
            f"<b>Attachments:</b> {html.escape(attachments)}</div>"
        )
        name = f"{index:03d}_{record['to'][0].replace('@', '_at_')}.html"
        with open(os.path.join(PREVIEW_DIR, name), "w", encoding="utf-8") as preview:
            preview.write(header + BODY)

    return write_report(
        "_summary.csv",
        [[r["to"][0], r["name"], ", ".join(str(x) for x in r["rows"])]
         for r in records],
        ["Email Address", "Full Name", "Source Row"],
    )


def report(records, skipped, address_log):
    print(f"Participants to mail : {len(records)} (one per person)")
    addresses = [r["to"][0].lower() for r in records]
    repeats = sorted({a for a in addresses if addresses.count(a) > 1})
    print(f"  anyone mailed twice: {repeats or 'no'}")
    print(f"Attachments          : "
          f"{', '.join(f'{os.path.basename(p)} ({os.path.getsize(p):,} bytes)' for p in ATTACHMENTS)}")
    if not MAPS_LINK:
        print("  NOTE: MAPS_LINK is empty - the Google Maps line is plain text, "
              "not a clickable link")

    if skipped:
        print(f"\nNOT MAILED ({len(skipped)}):")
        for entry in skipped:
            print(f"  row {entry['row']:>3} | {entry['name'][:28]:30s} | "
                  f"{entry['email'][:32]:34s} | {entry['reason']}")

    if address_log:
        print(f"\nMAIL ID NOTES ({len(address_log)}):")
        for entry in address_log:
            print(f"  row {entry['row']:>3} | {entry['raw'][:34]:36s} | {entry['action']}")

    paths = []
    if skipped or address_log:
        paths.append(write_report(
            "error in mail id.csv",
            [[e["row"], e["raw"], e["action"], e["used"] or "-",
              "No" if not e["used"] else "Yes (see summary)"]
             for e in address_log]
            + [[e["row"], e["email"], e["reason"], "-", "No"] for e in skipped],
            ["Source Row", "Original Mail ID", "What Happened", "Mail ID Used",
             "Mailed"],
        ))
    return paths


def load_sent_log():
    if not os.path.exists(SENT_LOG):
        return {}
    with open(SENT_LOG, encoding="utf-8", newline="") as handle:
        return {row["To"].lower(): row["Sent At"] for row in csv.DictReader(handle)}


def append_sent_log(record):
    os.makedirs(os.path.dirname(SENT_LOG) or ".", exist_ok=True)
    is_new = not os.path.exists(SENT_LOG)
    with open(SENT_LOG, "a", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        if is_new:
            writer.writerow(["To", "Full Name", "Sent At", "Attachments"])
        writer.writerow([
            record["to"][0],
            record["name"],
            datetime.now().isoformat(timespec="seconds"),
            "; ".join(os.path.basename(p) for p in ATTACHMENTS),
        ])


def send_emails():
    records, skipped, address_log = load_participants()
    report_paths = report(records, skipped, address_log)

    if DRY_RUN:
        summary_path = write_previews(records)
        print(f"\nDRY RUN - no mail sent. {len(records)} previews in {PREVIEW_DIR}/")
        for path in [summary_path] + report_paths:
            print(f"  {path}")
        return

    already_sent = load_sent_log()
    pending = [r for r in records if r["to"][0].lower() not in already_sent]
    if already_sent:
        print(f"\n{len(already_sent)} already mailed per {SENT_LOG} - skipping.")
    if not pending:
        print("Nothing left to send.")
        return

    total_mb = sum(os.path.getsize(p) for p in ATTACHMENTS) / 1_048_576
    print(f"\nAbout to send {len(pending)} REAL emails as {EMAIL_USER}, "
          f"each with {total_mb:.1f} MB of attachments.")
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
            server.send_message(create_message(record))
            append_sent_log(record)
            sent += 1
            print(f"Sent to {record['to'][0]}")
        except Exception as exc:
            failed.append({"to": record["to"][0], "error": str(exc)})
            print(f"FAILED {record['to'][0]}: {exc}")
        time.sleep(2)

    try:
        server.quit()
    except Exception:
        pass

    print(f"\nSent {sent}, failed {len(failed)}.")
    if failed:
        for entry in failed:
            print(f"  {entry['to']} | {entry['error']}")
        print(write_report("_failed.csv",
                           [[e["to"], e["error"]] for e in failed],
                           ["To", "Error"]))
        print(f"Re-running will retry only these - {SENT_LOG} records what succeeded.")


if __name__ == "__main__":
    send_emails()
