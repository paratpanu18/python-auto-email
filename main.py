import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import json
import getpass
import os

cc = []

def create_config() -> None:
    """Creates a configuration file with Gmail credentials."""
    gmail_user = input('Enter your Gmail email: ')
    gmail_password = getpass.getpass('Enter your Gmail password: ')
    try:
        print('Creating config file...')
        with open('config.json', 'w') as file:
            json.dump({'gmail_user': gmail_user, 'gmail_password': gmail_password}, file)
        print('Config file created successfully!')
    except Exception as e:
        print(f"Error creating config file: {e}")

def load_config() -> tuple[str, str] | None:
    """
    Loads Gmail credentials from the configuration file or creates a new one.

    :return: tuple[str, str]: Gmail username and password or None if an error occurred.
    :rtype: tuple[str, str] | None
    """
    if not os.path.isfile('config.json'):
        to_create = input('Config file not found. Create a new one? (Y/n): ').strip().lower()
        if to_create not in ['n', 'no']:
            create_config()

    try:
        print('Reading config file...')
        with open('config.json', 'r') as file:
            config = json.load(file)
        print('Config file read successfully!')
        return config['gmail_user'], config['gmail_password']
    except (FileNotFoundError, KeyError, json.JSONDecodeError):
        print("Error: Config file is missing or corrupted. Please recreate it.")
        exit()

def load_email_template(template_file) -> str:
    """
    Loads and returns the HTML email template.
    
    :return: HTML email template.
    :rtype: str
    """
    try:
        with open(template_file, 'r', encoding='utf-8') as file:
            return file.read()
    except FileNotFoundError:
        print(f"Error: Template file '{template_file}' not found.")
        exit()

def load_contacts(excel_file) -> pd.DataFrame:
    """Loads and returns contacts from the Excel file."""
    try:
        return pd.read_excel(excel_file)
    except FileNotFoundError:
        print(f"Error: Contacts file '{excel_file}' not found.")
        exit()

def setup_smtp_server(smtp_server, smtp_port, gmail_user, gmail_password):
    """Sets up and returns the SMTP server."""
    try:
        print('Setting up the SMTP server...')
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        print('Logging in...')
        server.login(gmail_user, gmail_password)
        print('Logged in successfully!')
        return server
    except smtplib.SMTPAuthenticationError:
        print('Error: Authentication failed. Please check your email and password.')
        print('You may need to enable "Less secure app access" in your Google account settings.')
        exit()
    except Exception as e:
        print(f"Error setting up the SMTP server: {e}")
        exit()

def send_emails(df, gmail_user, subject, html_body, cc, server):
    """Sends personalized emails to each contact."""
    total = len(df)
    for i, (index, row) in enumerate(df.iterrows(), start=1):
        name = row['Name']
        email = row['Email']
        
        # Create email message
        msg = MIMEMultipart()
        msg['From'] = gmail_user
        msg['Cc'] = ", ".join(cc)
        msg['To'] = email
        msg['Subject'] = subject

        # Replace {name} in the HTML body with the contact's name
        personalized_body = html_body.replace('{name}', name)
        msg.attach(MIMEText(personalized_body, 'html'))
        
        try:
            # Send the email
            server.sendmail(gmail_user, email, msg.as_string())
            print(f'[{i}/{total}] Email successfully sent to {name} ({email})')
        except Exception as e:
            print(f"Error sending email to {name} ({email}): {e}")

def main():
    # Load config or create if it doesn't exist
    gmail_user, gmail_password = load_config()

    # Email subject
    subject = input('Enter the email subject: ')

    # Load HTML email template
    template_file = input('Enter the HTML email template file (default: email_template.html): ') or 'email_template.html'
    html_body = load_email_template(template_file)

    # Load Excel contacts file
    excel_file = input('Enter the Excel file with contacts (default: contacts.xlsx): ') or 'contacts.xlsx'
    df = load_contacts(excel_file)

    # Setup SMTP server
    smtp_server = input('Enter the SMTP server (default: smtp.gmail.com): ') or 'smtp.gmail.com'
    smtp_port = int(input('Enter the SMTP port (default: 587): ') or 587)
    server = setup_smtp_server(smtp_server, smtp_port, gmail_user, gmail_password)

    # Send emails
    print(f'Sending emails to {len(df)} recipients...')
    send_emails(df, gmail_user, subject, html_body, cc, server)

    # Close the server connection
    server.quit()
    print('All emails sent successfully!')

if __name__ == '__main__':
    main()
