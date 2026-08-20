import csv
import html
import os
import re
import smtplib
import time
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import datetime

import openpyxl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")

SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

# Nothing is sent while this is True: every mail is written to PREVIEW_DIR
# instead and SMTP is never contacted. Flip to False only for the real run.
DRY_RUN = False
PREVIEW_DIR = "preview"

# Every successful send is appended here, and papers already listed are skipped
# on the next run. This is what makes a re-run after a crash safe: delete the
# file only if you genuinely intend to mail everyone again.
SENT_LOG = "sent_log.csv"

# A live run asks for typed confirmation before contacting SMTP.
REQUIRE_CONFIRMATION = True

META_REVIEWS_FILE = "official/MetaReviews.xlsx"
REVIEWS_FILE = "official/Reviews.xlsx"
PAPERS_FILE = "official/Papers.xlsx"

# Short papers / not-full-papers: one flat list of Paper IDs covering every
# track. Anything listed here is never mailed, whatever its meta-review says.
# Set to "" to run with no exclusion list at all.
EXCLUDED_PAPERS_FILE = "official/shortPaper.xlsx"

EMAIL_PATTERN = re.compile(r"^[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+$")

# Dates and links quoted in the letters.
CAMERA_READY_DEADLINE = "5 September 2026"
REGISTRATION_DEADLINE = "25 August 2026"
REGISTRATION_OPENS = "08 August 2026"
CAMERA_READY_LINK = "https://forms.gle/9hUW9otNMLedHRZY6"
PAPER_TEMPLATE_LINK = "https://event.iitg.ac.in/indis2026/Paper_Template.docx"

GREETING = """
<p>Dear Author(s),</p>

<p>
Greetings from the Organizing Committee of the
<strong>International Conference on Design and Innovation Studies (INDIS 2026).</strong>
</p>
"""

# Shared tail of the three acceptance letters - identical in all of them, so that
# every acceptance mail carries exactly the same conditions.
ACCEPTANCE_BLOCK = f"""
<p><strong>Camera-Ready Submission</strong></p>

<ul>
<li><strong>Deadline: {CAMERA_READY_DEADLINE}</strong></li>
<li><strong>Submission Link:</strong>
    <a href="{CAMERA_READY_LINK}">{CAMERA_READY_LINK}</a></li>
<li>Prepare the manuscript using the
    <strong>Taylor &amp; Francis Single-Column Template</strong>, available here:
    <a href="{PAPER_TEMPLATE_LINK}">{PAPER_TEMPLATE_LINK}</a></li>
<li><strong>All revisions made in response to the reviewers&rsquo; comments must be
    clearly highlighted</strong> in the submitted manuscript.</li>
</ul>

<p><strong>Publication Conditions</strong></p>

<p>
Publication in the <strong>Taylor &amp; Francis Conference Proceedings</strong> is
subject to <strong>all</strong> of the following conditions:
</p>

<ul>
<li>Submission of a compliant Camera-Ready manuscript before the deadline.</li>
<li>Adequate incorporation of the reviewers&rsquo; comments.</li>
<li>Compliance with the prescribed formatting, citation standards, publication
    ethics, and conference guidelines.</li>
<li>Registration of <strong>at least one author</strong> on or before
    <strong>{REGISTRATION_DEADLINE}</strong>.</li>
<li>Presentation of the paper during INDIS 2026 by a registered author.</li>
<li>To be considered for publication, the manuscript must have a Similarity Index
    below 15% and an AI-generated content score below 15%. Papers that do not
    comply with these requirements will be rejected.</li>
<li>The registration portal will open on <strong>{REGISTRATION_OPENS}</strong>
    through the INDIS 2026 website.</li>
</ul>

<p>
Failure to satisfy <strong>any</strong> of the above conditions will result in the
paper being withdrawn from the conference proceedings.
</p>

<p><strong>Page Limit</strong></p>

<p>
The maximum permitted length of a full paper is <strong>10 pages</strong>.
</p>

<p>
For manuscripts exceeding this limit, an <strong>additional publication charge of
INR 2,000 per extra page</strong> will be applicable during registration.
</p>

<p>
The Editorial Committee reserves the right to decline publication of the
Camera-Ready manuscript if the reviewers&rsquo; comments have not been adequately
addressed or if the manuscript fails to comply with the prescribed formatting
requirements, citation standards, publication ethics, or other conference
guidelines.
</p>

<p>
Congratulations on your acceptance. We look forward to welcoming you to IIT
Guwahati for INDIS 2026.
</p>
"""

SIGN_OFF = """
<p>
With best wishes,<br><br>
<strong>Dr. Debayan Dhar</strong><br>
On behalf of the Editorial Committee<br>
International Conference on Design and Innovation Studies (INDIS 2026)<br>
Department of Design<br>
Indian Institute of Technology Guwahati
</p>
"""

# Meta-reviewer recommendation (MetaReviews.xls, Q3) -> the letter the authors get.
# "acceptance" decides whether ACCEPTANCE_BLOCK follows the opening paragraphs.
DECISIONS = {
    "Accept": {
        "subject": "INDIS 2026 – Acceptance of Your Paper",
        "acceptance": True,
        "opening": """
<p>
We are pleased to inform you that, following a rigorous peer-review process, your
paper has been <strong>accepted</strong> for presentation at INDIS 2026 and for
publication in the conference proceedings, <strong>subject to successful completion
of the requirements outlined below.</strong>
</p>

<p>
The reviewers have provided valuable comments and suggestions for improving your
manuscript. Please carefully review the comments from <strong>Reviewers</strong>,
and incorporate all appropriate revisions while preparing the Camera-Ready version
of your paper.
</p>
""",
    },
    "Accept with minor revisions": {
        "subject": "INDIS 2026 – Acceptance with Minor Revisions",
        "acceptance": True,
        "opening": """
<p>
Following the peer-review process, your paper has been
<strong>accepted with Minor Revisions</strong>.
</p>

<p>
Before your paper can be included in the conference proceedings, please carefully
address the comments provided by Reviewers while preparing the Camera-Ready
manuscript.
</p>
""",
    },
    "Accept with major revisions": {
        "subject": "INDIS 2026 – Provisional Acceptance Subject to Major Revisions",
        "acceptance": True,
        "opening": """
<p>
Following a rigorous peer-review process, your paper has been provisionally
accepted, subject to satisfactory completion of major revisions.
</p>

<p>
The reviewers have identified important concerns that must be addressed before the
paper can be considered for publication in the conference proceedings. We therefore
request that you carefully study the comments from Reviewers, revise your
manuscript accordingly, and submit a substantially improved Camera-Ready version.
</p>
""",
    },
    "Reject": {
        "subject": "INDIS 2026 – Editorial Decision on Your Submission",
        "acceptance": False,
        "opening": f"""
<p>
Thank you for submitting your work to INDIS 2026.
</p>

<p>
This year, the conference received an overwhelming number of submissions, resulting
in a highly competitive review and selection process. Every paper underwent peer
review, and editorial decisions were made after careful consideration of the
reviewers&rsquo; recommendations and the overall publication capacity of the
conference proceedings.
</p>

<p>
After careful evaluation, we regret to inform you that your paper has not been
accepted for inclusion in INDIS 2026.
</p>

<p>
We encourage you to carefully read the comments provided by Reviewers. Their
observations are intended to help strengthen your work and may prove valuable for
future submissions or journal publications.
</p>

<p>
Please note that this decision should not be interpreted as an indication that your
research lacks merit or quality. In many cases, the reviewers identified issues
requiring further development, additional validation, stronger analysis, improved
presentation, or clearer articulation of the paper&rsquo;s contribution. Given the
production schedule for the conference proceedings, we believe that these revisions
cannot reasonably be completed within the remaining timeline.
</p>

<p>
In addition, under our publication agreement with Taylor &amp; Francis, the
conference proceedings are limited to approximately 100 papers (or approximately
1,000 published pages). Consequently, the Editorial Committee was required to make
difficult selection decisions even among papers with considerable potential.
</p>

<p>
Following the author registration deadline ({REGISTRATION_DEADLINE}), publication
space may become available if accepted papers are withdrawn or authors fail to
register. Should such vacancies arise, the Editorial Committee may review a limited
number of previously declined submissions for possible inclusion in the
proceedings. Any such consideration will be based on the original review reports,
editorial priorities, and the availability of publication space. This should not be
interpreted as an offer of acceptance or a guarantee that your paper will be
reconsidered.
</p>

<p>
We sincerely appreciate your interest in INDIS 2026 and thank you for giving us the
opportunity to review your work. We hope to receive your future submissions and
wish you continued success in your research.
</p>
""",
    },
}


# ---------------------------------------------------------------------------
# Reading the CMT exports
#
# CMT exports arrive in two shapes, and both are in use:
#   .xls  - SpreadsheetML 2003, which is XML text despite the extension, so
#           pandas/xlrd cannot open it;
#   .xlsx - real Excel (a ZIP), which the XML parser cannot open.
# Either way, each file has one worksheet per track, and each worksheet starts
# with a track-title row followed by the header row.
# ---------------------------------------------------------------------------

SS = "{urn:schemas-microsoft-com:office:spreadsheet}"


def read_worksheets(path):
    """Yield [row, ...] per worksheet, where each row maps 1-based column -> text."""
    if path.lower().endswith((".xlsx", ".xlsm")):
        yield from read_xlsx_worksheets(path)
    else:
        yield from read_spreadsheetml_worksheets(path)


def read_spreadsheetml_worksheets(path):
    """Read the SpreadsheetML 2003 (.xls) flavour of the export."""
    for worksheet in ET.parse(path).getroot().findall(f"{SS}Worksheet"):
        table = worksheet.find(f"{SS}Table")
        rows = []
        for row in table.findall(f"{SS}Row") if table is not None else []:
            cells, index = {}, 0
            for cell in row.findall(f"{SS}Cell"):
                # A skipped column carries an explicit ss:Index on the next cell.
                index = int(cell.get(f"{SS}Index", index + 1))
                data = cell.find(f"{SS}Data")
                cells[index] = "" if data is None else "".join(data.itertext()).strip()
            rows.append(cells)
        yield rows


def read_xlsx_worksheets(path):
    """Read the real-Excel (.xlsx) flavour of the export.

    Values are forced to str: Excel hands back Paper IDs as int, which would
    never match the string ids the other files yield, so the three-way join
    would silently come out empty.
    """
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        for worksheet in workbook.worksheets:
            rows = []
            for row in worksheet.iter_rows(values_only=True):
                cells = {}
                for index, value in enumerate(row, start=1):
                    if value is not None and str(value).strip():
                        cells[index] = str(value).strip()
                rows.append(cells)
            yield rows
    finally:
        workbook.close()


def split_header(rows):
    """Return (track, column-name -> index, data rows), locating the header by 'Paper ID'.

    The rows above the header hold the full track name. The worksheet name is not
    used for this: Excel caps it at 31 characters, so CMT truncates it there
    ("Product - Service - System...") while the title row keeps it intact.
    """
    for position, row in enumerate(rows):
        if any(value == "Paper ID" for value in row.values()):
            track = next(
                (value for above in rows[:position] for value in above.values() if value),
                "",
            )
            columns = {value: index for index, value in row.items() if value}
            return track, columns, rows[position + 1:]
    return "", {}, []


def find_column(columns, needle):
    """Look up a column index by substring, so CMT question numbers can shift."""
    for name, index in columns.items():
        if needle.lower() in name.lower():
            return index
    raise KeyError(f"no column matching {needle!r} in {sorted(columns)}")


def require_column(columns, name):
    """Exact-name lookup that says what was actually there when it fails."""
    if name in columns:
        return columns[name]
    raise KeyError(f"missing required column {name!r}; found {sorted(columns)}")


def paper_sort_key(paper_id):
    """Numeric ids in numeric order, anything else after them alphabetically."""
    return (0, int(paper_id), "") if paper_id.isdigit() else (1, 0, paper_id)


def note(paper_id, track, message):
    return {"paper_id": paper_id, "track": track, "reason": message}


def load_meta_reviews():
    """-> (Paper ID -> {title, recommendation}, Paper ID -> track, warnings).

    Only the Q3 recommendation is read. The meta-reviewer's name and email and
    the internal Q1/Q2 assessments stay out of the author-facing mail.
    """
    meta_reviews, tracks, warnings = {}, {}, []
    for rows in read_worksheets(META_REVIEWS_FILE):
        track, columns, data = split_header(rows)
        if not columns:
            continue
        paper_id_col = require_column(columns, "Paper ID")
        title_col = require_column(columns, "Paper Title")
        recommendation_col = find_column(columns, "Recommendation")
        for row in data:
            paper_id = row.get(paper_id_col, "")
            if not paper_id:
                continue
            recommendation = row.get(recommendation_col, "")
            if paper_id in meta_reviews:
                # Last row wins, so an unnoticed duplicate can mail the wrong
                # decision. Always surface it, and say so loudly when they differ.
                previous = meta_reviews[paper_id]["recommendation"]
                clash = (
                    f"CONFLICTING decisions {previous!r} vs {recommendation!r}"
                    if previous != recommendation
                    else f"same decision {recommendation!r}"
                )
                warnings.append(note(
                    paper_id, track,
                    f"duplicate Paper ID in MetaReviews - {clash}; the last row is used",
                ))
            meta_reviews[paper_id] = {
                "title": row.get(title_col, ""),
                "recommendation": recommendation,
            }
            tracks[paper_id] = track
    return meta_reviews, tracks, warnings


def load_reviews():
    """-> (Paper ID -> [review text, ...] in Reviewer 1..N order, tracks, warnings)."""
    reviews, tracks, warnings = defaultdict(list), {}, []
    for rows in read_worksheets(REVIEWS_FILE):
        track, columns, data = split_header(rows)
        if not columns:
            continue
        paper_id_col = require_column(columns, "Paper ID")
        review_col = find_column(columns, "detailed review")
        for row in data:
            paper_id = row.get(paper_id_col, "")
            if not paper_id:
                continue
            tracks[paper_id] = track
            review = row.get(review_col, "")
            if review:
                reviews[paper_id].append(review)
            else:
                warnings.append(note(
                    paper_id, track, "review row with empty text - not included",
                ))
    return reviews, tracks, warnings


def load_papers():
    """-> (Paper ID -> {primary, authors}, tracks, warnings).

    Author emails are ';'-separated with the primary contact marked '*'.
    """
    papers, tracks, warnings = {}, {}, []
    for rows in read_worksheets(PAPERS_FILE):
        track, columns, data = split_header(rows)
        if not columns:
            continue
        paper_id_col = require_column(columns, "Paper ID")
        primary_col = find_column(columns, "Primary Contact Author Email")
        authors_col = find_column(columns, "Author Emails")
        for row in data:
            paper_id = row.get(paper_id_col, "")
            if not paper_id:
                continue
            authors = [
                email.strip().rstrip("*").strip()
                for email in row.get(authors_col, "").split(";")
                if email.strip()
            ]
            if paper_id in papers:
                warnings.append(note(
                    paper_id, track,
                    "duplicate Paper ID in Papers - the last row's authors are used",
                ))
            papers[paper_id] = {
                "primary": row.get(primary_col, ""),
                "authors": authors,
            }
            tracks[paper_id] = track
    return papers, tracks, warnings


def load_excluded_papers():
    """Paper IDs that must never be mailed, from the short-paper list.

    This one fails closed. A missing file or a sheet without a 'Paper ID'
    column raises instead of yielding an empty set: silently mailing an author
    who was meant to be excluded is far worse than refusing to run.
    """
    if not EXCLUDED_PAPERS_FILE:
        return set()
    if not os.path.exists(EXCLUDED_PAPERS_FILE):
        raise FileNotFoundError(
            f"exclusion list {EXCLUDED_PAPERS_FILE!r} not found - "
            'set EXCLUDED_PAPERS_FILE = "" to run without one'
        )

    excluded, found_header = set(), False
    for rows in read_worksheets(EXCLUDED_PAPERS_FILE):
        _, columns, data = split_header(rows)
        if not columns:
            continue
        found_header = True
        paper_id_col = require_column(columns, "Paper ID")
        for row in data:
            paper_id = row.get(paper_id_col, "")
            if paper_id:
                excluded.add(paper_id)

    if not found_header:
        raise KeyError(
            f"no 'Paper ID' column found in {EXCLUDED_PAPERS_FILE!r} - "
            "refusing to run with an empty exclusion list"
        )
    return excluded


def build_records():
    """Join the three files.

    Returns (records, skipped, incomplete, warnings, excluded):
      records    - papers present in all three files that can be mailed
      skipped    - in all three, but unmailable (no template, no usable address)
      incomplete - present in only one or two files, so never mailed
      warnings   - mailable, but something about the data is worth a look
      excluded   - listed in the short-paper file, so deliberately not mailed
    """
    meta_reviews, meta_tracks, warnings = load_meta_reviews()
    reviews, review_tracks, review_warnings = load_reviews()
    papers, paper_tracks, paper_warnings = load_papers()
    warnings = warnings + review_warnings + paper_warnings
    excluded_ids = load_excluded_papers()

    def track_of(paper_id):
        for source in (meta_tracks, review_tracks, paper_tracks):
            if source.get(paper_id):
                return source[paper_id]
        return "(unknown track)"

    records, skipped = [], []
    common = set(meta_reviews) & set(reviews) & set(papers)

    for paper_id in sorted(common, key=paper_sort_key):
        # The exclusion list outranks everything else - check it before the
        # decision, so a listed paper is never mailed whatever its outcome.
        if paper_id in excluded_ids:
            continue

        meta_review = meta_reviews[paper_id]
        track = track_of(paper_id)
        decision = DECISIONS.get(meta_review["recommendation"])

        if decision is None:
            reason = (
                "no recommendation recorded"
                if not meta_review["recommendation"]
                else f"recommendation {meta_review['recommendation']!r} has no template"
            )
            skipped.append(note(paper_id, track, reason))
            continue

        to, cc, invalid = resolve_recipients(papers[paper_id])
        if invalid:
            warnings.append(note(
                paper_id, track, f"dropped malformed address(es): {', '.join(invalid)}",
            ))
        if not to:
            skipped.append(note(paper_id, track, "no usable author email address"))
            continue

        if len(reviews[paper_id]) < 2:
            warnings.append(note(
                paper_id, track,
                f"only {len(reviews[paper_id])} review(s); the letter says \"Reviewers\"",
            ))

        records.append({
            "paper_id": paper_id,
            "title": meta_review["title"],
            "track": track,
            "recommendation": meta_review["recommendation"],
            "decision": decision,
            "reviews": reviews[paper_id],
            "to": to,
            "cc": cc,
        })

    # Anything present in only one or two files is never mailed. Report every
    # such id in both directions - a paper with reviews but no meta-review is
    # just as much a gap as one with a decision but no reviews.
    present_in = {
        "MetaReviews": set(meta_reviews),
        "Reviews": set(reviews),
        "Papers": set(papers),
    }
    incomplete = []
    for paper_id in sorted(set().union(*present_in.values()) - common, key=paper_sort_key):
        # Excluded papers are reported under the exclusion list only, so they
        # are never counted twice.
        if paper_id in excluded_ids:
            continue
        found = [name for name, ids in present_in.items() if paper_id in ids]
        missing = [name for name in present_in if name not in found]
        incomplete.append({
            "paper_id": paper_id,
            "track": track_of(paper_id),
            "found_in": found,
            "missing_from": missing,
            "reason": "missing from " + ", ".join(missing),
        })

    # Account for every listed id, so a typo in the exclusion file is visible
    # rather than looking like a paper that was quietly dropped.
    everywhere = set().union(*present_in.values())
    excluded = []
    for paper_id in sorted(excluded_ids, key=paper_sort_key):
        if paper_id in common:
            state = "blocked - would otherwise have been mailed"
        elif paper_id in everywhere:
            state = "listed, but not mailable anyway (missing from " + ", ".join(
                name for name, ids in present_in.items() if paper_id not in ids) + ")"
        else:
            state = "listed, but not present in any of the three files"
        excluded.append(note(paper_id, track_of(paper_id), state))

    return records, skipped, incomplete, warnings, excluded


def resolve_recipients(paper):
    """Return (to, cc, invalid): primary contact in To, co-authors in Cc.

    Malformed addresses are dropped rather than passed to SMTP, where one bad
    address would fail the send. A paper is only lost if nothing valid remains.
    """
    seen, ordered, invalid = set(), [], []
    for email in [paper["primary"]] + paper["authors"]:
        if not email or email.lower() in seen:
            continue
        seen.add(email.lower())
        if EMAIL_PATTERN.match(email):
            ordered.append(email)
        else:
            invalid.append(email)
    if not ordered:
        return [], [], invalid
    return ordered[:1], ordered[1:], invalid


# ---------------------------------------------------------------------------
# Building the mail
# ---------------------------------------------------------------------------

def format_review(text):
    """Escape the verbatim review text and keep its line breaks."""
    return html.escape(text).replace("\n", "<br>")


def create_message(record):
    decision = record["decision"]

    msg = MIMEMultipart("mixed")
    msg["From"] = EMAIL_USER
    msg["To"] = ", ".join(record["to"])
    if record["cc"]:
        msg["Cc"] = ", ".join(record["cc"])
    # The Paper ID goes in both the subject and the body: several authors have
    # more than one submission, and without it two letters - sometimes an
    # acceptance and a rejection - are indistinguishable.
    msg["Subject"] = f"{decision['subject']} (Paper ID {record['paper_id']})"

    reviews_html = "".join(
        f"""
<p style="margin-bottom:4px;"><strong>Reviewer {number}:</strong></p>
<div style="border-left:3px solid #cfcfcf; padding-left:12px; margin-bottom:18px; white-space:normal;">
{format_review(review)}
</div>
"""
        for number, review in enumerate(record["reviews"], start=1)
    )

    body = f"""
<html>
<body style="margin:0; padding:0; font-family:Aptos, Calibri, Arial, sans-serif; font-size:12pt; color:#000000; line-height:1.5;">
{GREETING}
<p><strong>Paper ID: {html.escape(record["paper_id"])}</strong></p>
{decision["opening"]}
{ACCEPTANCE_BLOCK if decision["acceptance"] else ""}
{SIGN_OFF}
<hr style="border:none; border-top:1px solid #cfcfcf; margin:20px 0;">

<p><strong>Reviewer Comments</strong></p>

{reviews_html}

</body>
</html>
"""

    msg.attach(MIMEText(body, "html"))

    return msg


# ---------------------------------------------------------------------------
# Dry run / sending
# ---------------------------------------------------------------------------

def write_previews(records):
    """Write one .html per paper plus a summary sheet, without contacting SMTP."""
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

    summary_path = os.path.join(PREVIEW_DIR, "_summary.csv")
    with open(summary_path, "w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Paper ID", "Decision", "Reviews", "To", "Cc", "Track", "Title"])
        for record in records:
            writer.writerow([
                record["paper_id"],
                record["recommendation"],
                len(record["reviews"]),
                "; ".join(record["to"]),
                "; ".join(record["cc"]),
                record["track"],
                record["title"],
            ])

    return summary_path


def write_report(name, rows, fieldnames):
    """Write one of the diagnostic CSVs next to the previews."""
    os.makedirs(PREVIEW_DIR, exist_ok=True)
    path = os.path.join(PREVIEW_DIR, name)
    with open(path, "w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(fieldnames)
        writer.writerows(rows)
    return path


def report_data_issues(skipped, incomplete, warnings, excluded=()):
    """Print every paper that will not be mailed, with its track, and save it."""
    if excluded:
        blocked = [e for e in excluded if e["reason"].startswith("blocked")]
        print(f"\nSHORT-PAPER EXCLUSION LIST ({len(excluded)} id(s) listed, "
              f"{len(blocked)} actually blocked):")
        for entry in blocked:
            print(f"  paper {entry['paper_id']:>6s} | {entry['track']} | BLOCKED")
        others = [e for e in excluded if e not in blocked]
        if others:
            print(f"  {len(others)} listed id(s) were not mailable anyway "
                  f"or are not in this batch - see {PREVIEW_DIR}/_excluded.csv")

    if incomplete:
        print(f"\nNOT IN ALL THREE FILES - not mailed ({len(incomplete)}):")
        grouped = defaultdict(list)
        for entry in incomplete:
            grouped[", ".join(entry["found_in"])].append(entry)
        for found in sorted(grouped, key=lambda k: (-len(k), k)):
            entries = grouped[found]
            print(f"  found only in {found}  ({len(entries)} paper(s)) -"
                  f" missing {', '.join(entries[0]['missing_from'])}")
            for entry in entries:
                print(f"     paper {entry['paper_id']:>6s} | {entry['track']}")

    if skipped:
        print(f"\nIN ALL THREE BUT NOT MAILABLE ({len(skipped)}):")
        for entry in skipped:
            print(f"  paper {entry['paper_id']:>6s} | {entry['track']} | {entry['reason']}")

    if warnings:
        print(f"\nWARNINGS - mailed, but check these ({len(warnings)}):")
        for entry in warnings:
            print(f"  paper {entry['paper_id']:>6s} | {entry['track']} | {entry['reason']}")

    paths = []
    if excluded:
        paths.append(write_report(
            "_excluded.csv",
            [[e["paper_id"], e["track"], e["reason"]] for e in excluded],
            ["Paper ID", "Track", "Status"],
        ))
    if incomplete:
        paths.append(write_report(
            "_incomplete.csv",
            [[e["paper_id"], e["track"], ", ".join(e["found_in"]),
              ", ".join(e["missing_from"])] for e in incomplete],
            ["Paper ID", "Track", "Found In", "Missing From"],
        ))
    if skipped or warnings:
        paths.append(write_report(
            "_issues.csv",
            [[k, e["paper_id"], e["track"], e["reason"]]
             for k, group in (("skipped", skipped), ("warning", warnings))
             for e in group],
            ["Kind", "Paper ID", "Track", "Detail"],
        ))
    return paths


def load_sent_log():
    """Paper IDs already mailed in an earlier run, so a re-run never duplicates."""
    if not os.path.exists(SENT_LOG):
        return {}
    with open(SENT_LOG, encoding="utf-8", newline="") as handle:
        return {row["Paper ID"]: row["Sent At"] for row in csv.DictReader(handle)}


def append_sent_log(record, msg):
    """Record a send immediately, so a crash cannot lose what already went out."""
    is_new = not os.path.exists(SENT_LOG)
    with open(SENT_LOG, "a", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        if is_new:
            writer.writerow(["Paper ID", "Sent At", "Decision", "To", "Cc"])
        writer.writerow([
            record["paper_id"],
            datetime.now().isoformat(timespec="seconds"),
            record["recommendation"],
            msg["To"],
            msg["Cc"] or "",
        ])


def send_emails():
    records, skipped, incomplete, warnings, excluded = build_records()

    counts = defaultdict(int)
    for record in records:
        counts[record["recommendation"]] += 1
    recipients = {email.lower() for r in records for email in r["to"] + r["cc"]}

    print(f"Papers in all three files : {len(records)}")
    print(f"Unique recipients         : {len(recipients)}")
    for recommendation, count in sorted(counts.items()):
        print(f"  {recommendation:30s} {count:3d}")

    report_paths = report_data_issues(skipped, incomplete, warnings, excluded)

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
    print(f"\nAbout to send {len(pending)} REAL emails to "
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
        # One bad recipient must not abandon the rest of the batch.
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
