import imaplib
import email
from email.header import decode_header
from datetime import datetime
import re

# Set up IMAP connection
mail = imaplib.IMAP4_SSL("imap.gmail.com")
mail.login("<your-mail-id>", "<app_password>")

today = datetime.today().strftime("%d-%b-%Y")

# Select the mailbox you want to check (inbox by default)
mail.select("inbox")


status, messages = mail.search(None,  f'(ON "{today}")') # Search for today emails in the mailbox
# status, messages = mail.search(None,  "ALL") # Search for all emails in the mailbox

# Convert messages to a list of email IDs
messages = messages[0].split()

for mail_id in messages:
    status, msg_data = mail.fetch(mail_id, "(RFC822)")
    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                # If it's a bytes object, decode to string
                subject = subject.decode(encoding if encoding else "utf-8")
            from_ = msg.get("From")
            # print("From:", from_)
            # print("Subject:", subject)
            
            # Extract the email body
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))

                    if "attachment" not in content_disposition:
                        if content_type == "text/plain" or content_type == "text/html":
                            body = part.get_payload(decode=True).decode(part.get_content_charset())
                            # print("Body:", body)
                            break
            else:
                # Not multipart - i.e., plain text or HTML email
                body = msg.get_payload(decode=True).decode(msg.get_content_charset())
                # print("Body:", body)
            
            # Process the email to check if it's a bounce
            if "(failure)" in subject.lower() or "address not found" in body.lower() :
                print("This is a bounced email.")
                email_pattern = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
                print(re.findall(email_pattern, body)[0])
                # Handle the bounce (e.g., log it, update database, etc.)

# Logout and close the connection
mail.logout()
