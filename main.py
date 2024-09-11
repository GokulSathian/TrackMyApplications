from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
import base64
import openpyxl
from datetime import datetime
import os.path

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def get_gmail_service():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return build('gmail', 'v1', credentials=creds)

def read_emails_and_filter():
    # Get Gmail API service
    service = get_gmail_service()

    # Create a new Excel workbook and add a worksheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Filtered Emails"
    sheet.append(["Date", "Subject"])

    # Request a list of all messages
    results = service.users().messages().list(userId='me').execute()
    messages = results.get('messages', [])

    for message in messages:
        # Get the message details
        msg = service.users().messages().get(userId='me', id=message['id']).execute()
        
        # Get email subject
        subject = ''
        for header in msg['payload']['headers']:
            if header['name'] == 'Subject':
                subject = header['value']
                break

        # Get email date
        date_str = ''
        for header in msg['payload']['headers']:
            if header['name'] == 'Date':
                date_str = header['value']
                break
        
        # Parse and format the date
        date = datetime.strptime(date_str.split(' (')[0], "%a, %d %b %Y %H:%M:%S %z")
        formatted_date = date.strftime("%Y-%m-%d %H:%M:%S")

        # Get email body
        if 'parts' in msg['payload']:
            parts = msg['payload']['parts']
            data = parts[0]['body']['data']
        else:
            data = msg['payload']['body']['data']
        
        body = base64.urlsafe_b64decode(data).decode('utf-8')

        # Check if the email contains the specific content
        if "application was received" in body.lower():
            sheet.append([formatted_date, subject])

    # Save the Excel file
    workbook.save("filtered_emails.xlsx")

if __name__ == "__main__":
    read_emails_and_filter()
    print("Emails have been filtered and saved to 'filtered_emails.xlsx'")