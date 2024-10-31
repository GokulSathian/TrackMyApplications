# TRACK MY APPLICATION

## 📧 Gmail Email Filter & Export to Excel
This project retrieves emails from Gmail using the Google API, searches for specific content, and exports relevant details to an Excel file for easy tracking and filtering.

## ✨ What does it do?

The script connects to your Gmail account, filters out emails based on keywords, and exports the date and subject line of relevant emails to an Excel file (filtered_emails.xlsx). It’s especially useful if you're looking to monitor email responses automatically and keep records organized!

## 🚀 Key Features
- **Gmail Access**: Connects to your Gmail account through Google API for authorized read-only access.
- **Keyword Filtering**: Filters emails containing the phrase "application was received" (customizable).
- **Excel Export**: Saves filtered email details (date and subject) to an organized Excel sheet.
- **Automatic Authorization Handling**: Uses token.json for handling repeated authorizations smoothly.

## 🛠️ How it Works
- **Authorization**: The script authorizes with Gmail using credentials stored in token.json. If this file doesn't exist or is expired, the script refreshes the authorization flow and saves new credentials.
- **Email Fetching**: Connects to Gmail, retrieves recent emails, and checks each email's subject and date.
- **Filtering & Saving**: If the email's body contains the target keyword, the date and subject are saved to the Excel file.
  
## 📋 Requirements
- Python 3.x
- The following Python libraries:
1. google-auth and google-auth-oauthlib
2. google-auth-httplib2
3. google-api-python-client
4. openpyxl
You can install the required libraries using:

bash
Copy code
```
pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client openpyxl
```

## 🔑 Setup & Run Instructions
1. Get Credentials:
2. Create a new project in the Google Cloud Console.
3. Enable the Gmail API.
4. Download your credentials.json file and place it in the project folder.
5. Run the Script:
6. Run the script with:
bash
Copy code
```
python your_script.py
```
7. The first run will open a browser window for authorization. Afterward, token.json will be used for repeated authorizations without additional logins.
8. View Results:
The output Excel file, filtered_emails.xlsx, will contain the date and subject of all filtered emails.

## 🏗️ Code Structure
- get_gmail_service(): Handles Gmail API authorization and returns a service instance.
- read_emails_and_filter(): Fetches emails, filters by keyword, and exports matching emails to Excel.
  
## 📂 Project Files
- credentials.json: Required for Google OAuth.
- token.json: Stores access tokens for Gmail access.
- filtered_emails.xlsx: Output file containing filtered emails.

## 🤝 Contributions
Got ideas or want to contribute? Feel free to open issues or submit pull requests!

## 🌟 Enjoy Filtering!
This is a simple yet powerful tool to automate email filtering and export. Feel free to customize it as per your needs!

