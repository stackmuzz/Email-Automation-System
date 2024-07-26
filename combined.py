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
SERVICE_ACCOUNT_FILE = r'C:\Users\Amrendra pratapsingh\Downloads\Email automation\service account key\mail-automation-430509-97a9a75b8baf.json'
SPREADSHEET_ID = '1f1GxVdhHC1ezoEal3pcTitTdDvaxzHa6R90WWgl5wNg'  # Replace with your Google Sheets ID
SHEET_NAME = 'Sheet1'  # Replace with your sheet name
EMAIL_TEMPLATE = """
Greetings Professor,

I am Amrendra Pratap Singh, a final-year B.Tech student majoring in Computer Science and Information Technology at MJP Rohilkhand University, Bareilly. I am writing to inquire about any potential internship opportunities in your research group. I am keen to apply my knowledge in ML and deep learning to practical projects under your guidance. Your research aligns closely with my academic interests, and I am eager to contribute to and learn from your esteemed team.

Could we possibly schedule a brief meeting or call at your convenience to discuss any internship openings or projects where I could contribute? I would greatly appreciate your guidance on the application process or any further steps required.

I have a strong passion for ML and deep learning and have been actively involved in self-study and practical applications in this field. I am eager to contribute to your ongoing projects and learn from your expertise. I am flexible regarding the internship duration and schedule. I would take care of accommodation and other related stuff on my own. Attached is my CV [here](https://drive.google.com/file/d/1ydmBYKDYyM8vvuyV9JxTH79dy7Wn-j47/view?usp=sharing) for your consideration.

Thank you for considering my application. I look forward to the opportunity to work with you.

Best Regards,

Amrendra Pratap Singh

MJP Rohilkhand University, Bareilly

+91 7078045398
"""
SUBJECT = "Inquiry regarding internship opportunity"

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
        r'C:\Users\Amrendra pratapsingh\Downloads\Email automation\oAuth client json\client_secret_509924110112-labu54u7eq2bg0ipgtupl0vg85vb6o10.apps.googleusercontent.com.json', SCOPES)
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
