import os
import base64
import json
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Constants
SERVICE_ACCOUNT_FILE = r'Path to serviceaccount.json '
SPREADSHEET_ID = 'Sheet id '  # Replace with your Google Sheets ID
SHEET_NAME = 'Sheet1'  # Replace with your sheet name
EMAIL_TEMPLATE = """
Greetings Professor,

Type your email here!!!!!

"""
SUBJECT = "Mail Subject "

# Define the scopes
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly', 'https://www.googleapis.com/auth/gmail.modify']


# Create the Sheets API service
def create_sheets_service():
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=credentials)
    return service


# Create the Gmail API service with OAuth 2.0
def create_gmail_service():
    flow = InstalledAppFlow.from_client_secrets_file(
        r'path to client oauth.json', SCOPES)
    creds = flow.run_local_server(port=0)
    service = build('gmail', 'v1', credentials=creds)
    return service


# Read data from Google Sheets
def read_sheet(service):
    range_name = f'{SHEET_NAME}!B334:B385'  # Define the range
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=range_name).execute()
    values = result.get('values', [])
    return values


# Create and send an email
def send_email(service, to, subject, body):
    message = MIMEMultipart()
    message['to'] = to
    message['subject'] = subject
    message.attach(MIMEText(body, 'html'))

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')

    try:
        message = service.users().messages().send(userId="me", body={'raw': raw_message}).execute()
        print(f"Message sent to {to} with Message ID: {message['id']}")
    except HttpError as error:
        print(f"An error occurred: {error}")


# Main function
def main():
    # Create services
    sheets_service = create_sheets_service()
    gmail_service = create_gmail_service()

    # Read data from Google Sheets
    email_list = read_sheet(sheets_service)

    # Send emails
    for row in email_list:
        if row:
            email = row[0]  # Assuming the email addresses are in column B
            body = EMAIL_TEMPLATE  # Use the new email template
            send_email(gmail_service, email, SUBJECT, body)


if __name__ == '__main__':
    main()
