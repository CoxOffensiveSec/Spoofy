import argparse
import sys
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase  # Add import for MIMEBase
from email import encoders  # Add import for encoders
import smtplib
import random
import time

def extract_first_name(email):
    # Extract first name before dot (.) if exists, otherwise before @
    match = re.match(r'([^@.]+)\.', email)
    if match:
        return match.group(1).capitalize()
    return email.split('@')[0].capitalize()

def read_html_file(html_file):
    try:
        with open(html_file, 'r') as file:
            html_content = file.read()
        return html_content
    except FileNotFoundError:
        print(f"Error: {html_file} not found.")
        sys.exit(1)

def send_emails(sender_email, recipient_file, delay_range, rhost, attachments=None, html_file=None):
    # Fixed values or variables
    static_rhost = ".mail.protection.outlook.com"
    rport = 25  # or 489, 587 for alternative ports

    # Read recipient email addresses from the specified file
    try:
        with open(recipient_file, 'r') as file:
            recipient_emails = file.read().splitlines()
    except FileNotFoundError:
        print(f"Error: {recipient_file} not found.")
        return

    # Iterate over each recipient email and send the email
    for recipient_email in recipient_emails:
        # create message object instance for each recipient
        msg = MIMEMultipart()

        # setup the parameters of the message
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = "Hi"

        # Extract and capitalize first name from recipient's email
        first_name = extract_first_name(recipient_email)

        # Determine HTML content source (from file or default)
        if html_file:
            html_content = read_html_file(html_file)
            # Replace placeholder {first_name} in HTML content with extracted first name
            html_content = html_content.replace('{first_name}', first_name)
        else:
            # Default HTML content for the message body with dynamic capitalized first name
            html_content = f"""
                   <html>
              <head></head>
              <body>
                <p>{first_name},</p>
                <p>Cold reboots with the 0-day Linux roots<br>
                I got a job - but I still spam. </p>
                <p><br><br>Cheers!</p>
              </body>
            </html>
            """


        # Attach HTML content as part of the message
        msg.attach(MIMEText(html_content, 'html'))

        # Attach files if provided
        if attachments:
            for attachment in attachments:
                attach_file(msg, attachment)

        print(f"[*] Payload is generated : HTML Email for {recipient_email}")

        server = smtplib.SMTP(host=rhost + static_rhost, port=rport)

        if server.noop()[0] != 250:
            print("[-] Connection Error")
            exit()

        server.starttls()

        # Uncomment if log-in with authentication
        # password = ""
        # server.login(msg['From'], password)

        server.sendmail(msg['From'], msg['To'], msg.as_string())
        print(f"[***] Successfully sent HTML email to {msg['To']}")

        server.quit()

        # Generate random jitter delay within the specified range
        jitter_delay = random.uniform(delay_range[0], delay_range[1])
        print(f"[*] Jitter delay: {jitter_delay:.2f} seconds")
        time.sleep(jitter_delay)

    print("[***] All emails have been sent.")

def attach_file(msg, filepath):
    # Open the file to be sent
    with open(filepath, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email    
    encoders.encode_base64(part)

    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filepath}",
    )

    # Add attachment to message and convert message to string
    msg.attach(part)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Send HTML emails to multiple recipients with a jitter delay and optional attachments.')
    parser.add_argument('--sender', required=True, help='Sender email address')
    parser.add_argument('--recipient-file', required=True, help='Path to file containing recipient email addresses')
    parser.add_argument('--delay-min', type=float, default=1.0, help='Minimum delay in seconds between sending each email (default: 1.0)')
    parser.add_argument('--delay-max', type=float, default=3.0, help='Maximum delay in seconds between sending each email (default: 3.0)')
    parser.add_argument('--rhost', required=True, help='Hostname part before .mail.protection.outlook.com')
    parser.add_argument('--attachments', nargs='+', default=None, help='List of file paths to attach to the email (optional)')
    parser.add_argument('--html-file', help='Path to HTML file to use as email body (optional)')

    args = parser.parse_args()

    sender_email = args.sender
    recipient_file = args.recipient_file
    delay_min = args.delay_min
    delay_max = args.delay_max
    rhost = args.rhost
    attachments = args.attachments
    html_file = args.html_file

    # Ensure delay_min is less than delay_max
    if delay_min >= delay_max:
        print("Error: --delay-min must be less than --delay-max")
        sys.exit(1)

    delay_range = (delay_min, delay_max)

    send_emails(sender_email, recipient_file, delay_range, rhost, attachments, html_file)