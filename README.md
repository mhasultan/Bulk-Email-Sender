# Email-Sender
This is an automated email sender using the Gmail API. It personalizes emails, and supports bulk email sending with multiple Gmail accounts.

## Features
- Send bulk emails via Gmail API.
- Personalize email content with dynamic placeholders.
- Graphical User Interface (GUI)
- Logs all sent emails for tracking.
- Send emails in primary inbox
- Add attachments to the emails
  
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

## Running the Script
Run the script with:

```bash
python gui_send_email.py
```

## Dependencies
Ensure all required Python libraries are installed:

```bash
pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client pandas imgkit pillow tkinter
```

## Notes
- The script logs email activity in `mail.log`.

This project is for educational purposes. Use responsibly.

