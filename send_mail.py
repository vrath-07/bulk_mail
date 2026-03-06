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


def create_message(name, recipient_email):

    msg = MIMEMultipart("related")
    msg["From"] = EMAIL_USER
    msg["To"] = recipient_email
    msg["Subject"] = "Invitation to Participate – INDIS 2026 | IIT Guwahati"

    body = f"""
<html>
<body style="margin:0; padding:0; font-family:Aptos, Calibri, Arial, sans-serif; font-size:12pt; color:#000000; line-height:1.5;">

<p>Dear Prof. {str(name).title()},</p>

<p>
Greetings from the Department of Design, Indian Institute of Technology Guwahati.
</p>

<p>
We are pleased to invite you and your colleagues to participate in 
<strong>INDIS 2026 – International Conference on Design and Innovation Studies</strong>, 
to be held at IIT Guwahati from <strong>28–30 September 2026</strong>. 
The conference aims to bring together researchers, scholars, practitioners, 
and industry experts to explore <strong>design-led innovation for sustainable and inclusive futures</strong>.
</p>

<p>
INDIS 2026 will feature research presentations, student exhibitions, and industry engagement, 
and will be held alongside <strong>Asia Design Week</strong>, creating a vibrant platform for 
international collaboration and knowledge exchange.
</p>

<p>
We particularly encourage <strong>faculty members, researchers, and doctoral scholars</strong> 
to submit <strong>original research contributions</strong>. The conference invites submissions 
in the following categories:
</p>

<ul>
<li>Full Research Papers</li>
<li>Case Studies & Industry Papers</li>
<li>Research-through-Design (RtD)</li>
<li>Posters / Short Papers</li>
</ul>

<p>
All submissions will undergo a <strong>double-blind peer review process</strong> and accepted 
papers will be published in the <strong>Springer Proceedings</strong>, with selected papers 
invited for journal publication.
</p>

<p><strong>Important Dates</strong><br>
Portal Opens: 5 April 2026<br>
Submission Deadline: 15 April 2026
</p>

<p>
We would greatly appreciate it if you could:
</p>

<ol>
<li>Consider submitting your research work to the conference.</li>
<li>Share this call for papers with colleagues and research scholars in your department and network.</li>
<li>Follow and engage with the conference updates on LinkedIn:<br>
<a href="https://www.linkedin.com/company/indis-2026-international-conference-on-design-and-innovation-studies/">
https://www.linkedin.com/company/indis-2026-international-conference-on-design-and-innovation-studies/
</a>
</li>
</ol>

<p>
The conference website with full details can be accessed here:<br>
<a href="https://event.iitg.ac.in/indis2026/" style="color:#1155cc; text-decoration:underline;">
https://event.iitg.ac.in/indis2026/
</a>
</p>

<p>
Please find the <strong>conference poster attached</strong> for your reference and circulation.
</p>

<p>
We look forward to your participation and to welcoming you to IIT Guwahati for INDIS 2026.
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
    with open("poster.pdf", "rb") as file:
        part = MIMEApplication(file.read(), Name="INDIS_2026_Poster.pdf")
        part["Content-Disposition"] = 'attachment; filename="INDIS_2026_Poster.pdf"'
        msg.attach(part)

    return msg


def send_emails():
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()
    server.login(EMAIL_USER, EMAIL_PASS)

    for _, row in df.iterrows():
        msg = create_message(
            row["Name"],
            row["Email"]
        )
        server.send_message(msg)
        print(f"Sent to {row['Email']}")
        time.sleep(2)

    server.quit()
    print("All emails sent successfully.")


if __name__ == "__main__":
    send_emails()