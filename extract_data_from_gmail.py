import imaplib
import email
from email.header import decode_header
import openpyxl
import re
from tqdm import tqdm


# Access Data to Gmail
username = 'youremail@gmail.com'
password = 'yourpassword'

# Connect IMAP Gmail server
mail = imaplib.IMAP4_SSL("imap.gmail.com")
mail.login(username, password)
mail.select("inbox")  # Select "Inbox" folder

# Search for specific emails
result, data = mail.search(None, '(FROM "test@email.com")')

# Extract IDs from found emails
email_ids = data[0].split()

# Create Excel File
wb = openpyxl.Workbook()
ws = wb.active
ws.append(["Name", "Email"])

# Configure Progress Bar
total_emails = len(email_ids)
progress_bar = tqdm(total=total_emails, desc="Processing Emails")

# Process every email
for email_id in email_ids:
    result, data = mail.fetch(email_id, "(RFC822)")
    raw_email = data[0][1]
    msg = email.message_from_bytes(raw_email)

    # Decode Sender
    sender = msg["From"]

    # Extract body message content
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))
            try:
                body = part.get_payload(decode=True).decode()
            except:
                pass
            if content_type == "text/plain" and "attachment" not in content_disposition:
                break
    else:
        body = msg.get_payload(decode=True).decode()

    # Do analysis of the extracted data, e.g. extract names and emails.
    # This will depend on the structure of the email and your specific requirements.
    # Here is a basic example of how you might search for email addresses and names.
    emails_found = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', body)
    name_match = re.search(r'Name:\s*(\w+)', body)

    
    # If a name is found, use it; otherwise, use an empty string
    if name_match:
        name = name_match.group(1).strip()  # Extract the name of the found pattern
    else:
        name = ""

    # If at least one email is found, add them to the Excel file.
    if emails_found:
        for email_found in emails_found:
            ws.append([name, email_found])

    # Update Progress Bar
    progress_bar.update(1)

# Close Progress Bar
progress_bar.close()

# Save Excel File
wb.save("data.xlsx")

# Cerrar la conexión
mail.close()
mail.logout()
