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

df = pd.read_excel("indis_cmt.xlsx")


def create_message(recipient_email):

    msg = MIMEMultipart("related")
    msg["From"] = EMAIL_USER
    msg["To"] = recipient_email
    msg["Subject"] = "Extension of Paper Submission Deadline – INDIS 2026"

    body = f"""
<html>
<body style="margin:0; padding:0; font-family:Aptos, Calibri, Arial, sans-serif; font-size:12pt; color:#000000; line-height:1.5;">

<p>Dear Authors and Participants,</p>

<p>
Greetings from the INDIS 2026 Secretariat.
</p>

<p>
We are pleased to inform you that, in response to multiple requests from prospective contributors, the deadline for paper submission to the International Conference on Design and Innovation Studies (INDIS 2026) has been extended.
</p>

<p>
<strong>The revised submission deadline is: 15 May 2026, 1700 hrs IST </strong>
</p>
<p>
We encourage you to make use of this extended window to prepare and submit your work.
</p>
<p>
For all further updates, announcements, and detailed information regarding the conference, we request you to regularly visit our official webpage and follow our LinkedIn page.
</p>
<p>
INDIS 2026 LinkedIn Page:
<a href="https://www.linkedin.com/company/indis-2026-international-conference-on-design-and-innovation-studies/" target="_blank">
Visit our LinkedIn page
</a>
</p>

<p>We look forward to your valuable contributions and to welcoming you at INDIS 2026.
</p>

<p>
Warm regards,<br><br>
Dr. Debayan Dhar
<br>
Vice-Chair INDIS 2026
<br>
Email: <a href="mailto:indis2026@iitg.ac.in">indis2026@iitg.ac.in</a><br>
Conference Website: 
<a href="https://event.iitg.ac.in/indis2026/">https://event.iitg.ac.in/indis2026/</a>
</p>
<br>
On behalf of
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

    # Attach poster
    # with open("poster.pdf", "rb") as file:
    #     part = MIMEApplication(file.read(), Name="INDIS_2026_Poster.pdf")
    #     part["Content-Disposition"] = 'attachment; filename="INDIS_2026_Poster.pdf"'
    #     msg.attach(part)

    return msg


def send_emails():
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()
    server.login(EMAIL_USER, EMAIL_PASS)

    for _, row in df.iterrows():
        msg = create_message(
            row["Email"]
        )
        server.send_message(msg)
        print(f"Sent to {row['Email']}")
        time.sleep(2)

    server.quit()
    print("All emails sent successfully.")


if __name__ == "__main__":
    send_emails()