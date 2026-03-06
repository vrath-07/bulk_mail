import smtplib
import os
import pandas as pd
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")

SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

df = pd.read_excel("contacts.xlsx")


def create_message(first_name, last_name, recipient_email,
                   designation, committee, institution, theme, track):

    msg = MIMEMultipart("related")
    msg["From"] = EMAIL_USER
    msg["To"] = recipient_email
    msg["Subject"] = f"Invitation to join as the {str(designation).title()} – INDIS 2026"

    # Handle possible empty values safely
    designation = str(designation).strip()
    committee = str(committee).strip()
    theme = str(theme).strip()
    track = str(track).strip()

    body = f"""
<html>
<body style="margin:0; padding:0; font-family:Aptos, Calibri, Arial, sans-serif; font-size:12pt; color:#000000; line-height:1.5;">

<p>Dear {str(first_name).title()} {str(last_name).title()},</p>

"""

    # Conditional section for Theme / Track
    if designation.lower() == "theme chair" and theme:
        body += f"""
<p>
I am writing to invite you to serve as the <strong>{designation.title()}</strong>
for the theme <strong>{theme}</strong> of the <strong>International Conference on Design and Innovation Studies (INDIS 2026)</strong>,
to be hosted by the Department of Design, Indian Institute of Technology Guwahati.
</p>
"""
    elif designation.lower() == "track chair" and track:
        body += f"""
<p>
I am writing to invite you to serve as the <strong>{designation.title()}</strong>
for <strong>{track}</strong> for the <strong>International Conference on Design and Innovation Studies (INDIS 2026)</strong>,
to be hosted by the Department of Design, Indian Institute of Technology Guwahati.
</p>
"""

    body += f"""
<p>
INDIS 2026 is being established as a focused international platform for scholarly
work at the intersection of design, innovation, technology, and society.
The conference will bring together academic researchers and industry practitioners
through keynote sessions, peer-reviewed paper presentations, and curated academic discussions.
The event will be organized alongside <strong>Asia Design Week</strong>, providing a broader
international context for academic and professional exchange.
</p>

<p>
The conference will feature peer-reviewed contributions in the form of full papers,
case studies and industry papers, research-through-design contributions, and posters.
Accepted papers will be published in the Springer conference proceedings (ISBN),
with selected papers considered for journal publication.
</p>

<p>
Your presence as the {designation.title()} would play an important role in guiding the
academic direction and long-term development of the conference. We would greatly value
your support in strengthening the scholarly profile and international reach of INDIS.
</p>

<p>
I am attaching the conference poster for your reference. I would appreciate it if you
could kindly circulate it within your academic and professional networks to help us
reach a wider community of scholars and practitioners.
</p>

<p>
Further details are available at: <br>
<a href="https://event.iitg.ac.in/indis2026/" style="color:#1155cc; text-decoration:underline;">
https://event.iitg.ac.in/indis2026/
</a>
</p>

<p>
I hope you will be willing to support this initiative.
</p>

<p>
Warm regards,<br>
Prof. Pratul Chandra Kalita
</p>

<p style="margin-bottom:5px;">
<em>On behalf of</em><br>
</p>

<hr style="border:none; border-top:1px solid #cfcfcf; margin:8px 0;">

<img src="cid:signature_image" style="display:block; max-width:40%; height:auto;">

</body>
</html>
"""

    msg.attach(MIMEText(body, "html"))

    # Attach inline signature image
    with open("signature.png", "rb") as img:
        mime_img = MIMEImage(img.read())
        mime_img.add_header("Content-ID", "<signature_image>")
        mime_img.add_header("Content-Disposition", "inline", filename="signature.png")
        msg.attach(mime_img)

    # Attach poster file (PNG properly named)
    with open("poster.png", "rb") as file:
        part = MIMEApplication(file.read(), Name="INDIS_2026_Poster.png")
        part["Content-Disposition"] = 'attachment; filename="INDIS_2026_Poster.png"'
        msg.attach(part)

    return msg


def send_emails():
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()
    server.login(EMAIL_USER, EMAIL_PASS)

    for _, row in df.iterrows():
        msg = create_message(
            row["First Name"],
            row["Last Name"],
            row["Email"],
            row["Designation"],
            row["Committee"],
            row["Institution"],
            row["Theme"],
            row["Track"]
        )
        server.send_message(msg)
        print(f"Sent to {row['Email']}")
        time.sleep(2)

    server.quit()
    print("All emails sent successfully.")


if __name__ == "__main__":
    send_emails()