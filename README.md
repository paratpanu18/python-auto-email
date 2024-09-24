# Python Auto Email (Bulk Email Sender)

This Python script allows you to send personalized emails in bulk to a list of recipients from an Excel file. It supports Gmail accounts, HTML email templates, and Cc functionality for specific addresses.

## Features
- Sends personalized bulk emails to recipients listed in an Excel file.
- Allows HTML email templates for custom styling and personalization.
- Option to create a configuration file for storing Gmail credentials.
- Uses Gmail's SMTP server to send emails.
- Adds a Cc list to every email sent.


## Getting Started
### 1. Clone the repository
```bash
git clone https://github.com/paratpanu18/python-auto-email.git
cd python-auto-email
```

### 2. Install Python's dependencies
```bash
pip install -r requirements.txt
```

### 3. Configure Gmail Credential
If the config.json file is not found, the script will prompt you to create one. It securely stores your Gmail email and password.
> If the config.json file is not present when the code is executed, a new configuration file will be automatically generated.

Here is the format of `config.json` file:
```json
{
    "gmail_user": "john.doe@example.com", 
    "gmail_password": "https://youtu.be/elsh3J5lJ6g"
}
```

### 4. Prepare Your Contacts
Prepare an Excel file (`contacts.xlsx` by default) containing your recipients. The file should have at least two columns: `Name` and `Email`. Here's an example:
| Name        | Email                   |
|-------------|--------------------------|
| John Doe    | john.doe@example.com     |
| Jane Smith  | jane.smith@example.com   |
| Alice Brown | alice.brown@example.com  |
| Bob White   | bob.white@example.com    |
| Charlie Lee | charlie.lee@example.com  |

### 5. Create an HTML Email Template
Create an HTML file for your email body (e.g., email_template.html). You can use {name} as a placeholder for the recipient's name. Here's an example template:
```html
<!DOCTYPE html>
<html>
<body>
    <h1>Hello {name}!</h1>
    <p>We are excited to announce our new service.</p>
</body>
</html>
```

### 6. Run it !!!
```bash
python main.py
```
The script will prompt you to enter the following:
- Email subject
- HTML email template file (default: email_template.html)
- Excel contacts file (default: contacts.xlsx)
- SMTP server (default: smtp.gmail.com)
- SMTP port (default: 587)

### 7. Log Output
As the script runs, it will provide real-time feedback, indicating whether the emails were successfully sent.

```less
[1/100] Email successfully sent to John Doe (john.doe@example.com)
[2/100] Email successfully sent to Jane Smith (jane.smith@example.com)
```

## Customization
- **Cc List**: You can modify the cc list in the script to include email addresses that should receive a copy of every email.
- **Email Template**: Modify the HTML template to create a more personalized or branded email layout.
- **SMTP Settings**: The default SMTP server is Gmail. You can modify it if you're using a different email provider by changing the SMTP server and port.
