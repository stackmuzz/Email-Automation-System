# Email-Automation-System
Custom built email automation service.

# Email Automation Project

This project automates the process of sending emails to a list of recipients stored in a Google Sheet. The script reads the email addresses from the sheet, composes a predefined email, and sends it via the Gmail API. 

## Features

- Read email addresses from a specified range in a Google Sheet.
- Send personalized emails to each recipient.
- Uses Google Sheets API and Gmail API.
- Logs the status of email sending for debugging purposes.

## Prerequisites

- Python 3.7+
- Google Cloud Platform account
- Google Sheets with email addresses
- Gmail account for sending emails

## Setup

### Step 1: Google Cloud Platform Configuration

1. **Create a Project:**
   - Go to the [Google Cloud Console](https://console.cloud.google.com/).
   - Create a new project or select an existing project.

2. **Enable APIs:**
   - Enable the Google Sheets API and Gmail API for your project.

3. **Create OAuth Consent Screen:**
   - Navigate to `APIs & Services` > `OAuth consent screen`.
   - Set up the OAuth consent screen. Select `External` for the user type.
   - Add your email and any other information required.
   - Add scopes: `https://www.googleapis.com/auth/spreadsheets.readonly` and `https://www.googleapis.com/auth/gmail.send`.

4. **Create OAuth Client ID:**
   - Navigate to `APIs & Services` > `Credentials`.
   - Create credentials > OAuth 2.0 Client IDs.
   - Select `Desktop app` for the application type.
   - Download the JSON file. This will be your `credentials.json`.

5. **Add Test Users:**
   - Navigate back to the OAuth consent screen.
   - Add any email addresses that will use this application under the `Test users` section.

### Step 2: Local Environment Setup

1. **Install Python Packages:**
   - Use the following pip command to install required packages:
     ```sh
     pip install gspread google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client
     ```

2. **Save `credentials.json`:**
   - Save the downloaded `credentials.json` file to your project directory.

### Step 3: Google Sheets Configuration

1. **Create a Google Sheet:**
   - Create a new Google Sheet or use an existing one.
   - Ensure email addresses are stored in a column, for example, column B from row 50 to row 350.

2. **Share the Sheet:**
   - Share the Google Sheet with the service account email found in your `credentials.json`.

## Usage

1. **Run the Script:**
   - Run the Python script to start the email automation process.
   - The script will authenticate using OAuth, read the email addresses from the Google Sheet, and send emails.

2. **Check Logs:**
   - Check the logs for the status of each email sent. This will help in debugging any issues.

## Customization

- **Email Template:**
   - Customize the email template within the script as needed.
   
- **Email Range:**
   - Modify the range from which email addresses are read in the Google Sheet.

## Troubleshooting

- Ensure that the `credentials.json` file is correctly referenced in your script.
- Verify that the Google Sheet is shared with the service account email.
- Make sure to follow the OAuth consent screen setup correctly to avoid authorization issues.
- If encountering errors, check the logs for detailed error messages and troubleshoot accordingly.


