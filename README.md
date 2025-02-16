# Email-Sender
This is an automated email sender using the Gmail API. It personalizes emails, generates dynamic PDF attachments from HTML, and supports bulk email sending with multiple Gmail accounts.

## Features
- Send bulk emails via Gmail API.
- Personalize email content with dynamic placeholders.
- Convert HTML invoices to PDF attachments.
- Logs all sent emails for tracking.

## Requirements
Ensure you have the following installed:

- Python 3.10
- Required dependencies in the dependencies section.

## Setting Up Gmail API Credentials
1. **Go to Google Cloud Console**
   - Visit: [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project.

2. **Enable Gmail API**
   - Navigate to `API & Services > Library`
   - Search for `Gmail API` and enable it.

3. **Create OAuth Credentials**
   - Go to `API & Services > Credentials`
   - Click `Create Credentials > OAuth client ID`
   - Configure the consent screen and create credentials.
   - Download the `credentials.json` file and place it in the project directory.

## Configuring Sender and Receiver Emails
1. **Adding Senders (Gmail Accounts)**
   - Open `gmail.csv` and add sender email addresses with their app passwords.
   
   Format:
   ```csv
   email,password
   sender1@gmail.com,app_password
   sender2@gmail.com,app_password
   ```

2. **Adding Receivers (Contacts)**
   - Open `contacts.csv` and add recipient details.
   
   Format:
   ```csv
   name,email
   John Doe,johndoe@example.com
   Jane Doe,janedoe@example.com
   ```

## Customizing Email Content
1. **Email Body**
   - Edit `body.txt` to change the email content.

2. **Email Subject**
   - Modify `subjects.csv` to customize subjects.

3. **HTML Invoice for PDF Attachment**
   - Update `html_code.html` to change the invoice content before conversion to PDF.

## Running the Script
Run the script with:

```bash
python send_email.py
```

## Dependencies
Ensure all required Python libraries are installed:

```bash
pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client pandas imgkit pillow
```

## Notes
- Ensure `wkhtmltopdf` is installed and correctly configured.
- The script logs email activity in `mail.log`.

This project is for educational purposes. Use responsibly.

