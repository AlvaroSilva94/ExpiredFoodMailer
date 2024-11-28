import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from google.oauth2 import service_account
from datetime import date, datetime, timedelta
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv
load_dotenv()

def get_sheet():
     # Use credentials to authenticate and access Google Sheet
    scope = ['https://spreadsheets.google.com/feeds',
              'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('expirationdates-0f9604e3480e.json', scope)
    client = gspread.authorize(creds)
    sheet_url = 'https://docs.google.com/spreadsheets/d/1DVLac3F7hx1gdTB_y3Dh83YlKYEhIxMMvOx2M7Qxd5g/edit#gid=0'
    sheet = client.open_by_url(sheet_url).worksheet('Prazos')

    #---------------------------------------------------
    service = build('sheets', 'v4', credentials=creds)
    #----------------------------------------------------

    # Get all values from the sheet and convert to a dictionary
    data_dict = {}
    for row in sheet.get_all_values()[1:]:
         key = row[0]
         value = tuple(row[1:])
         data_dict[key] = value

    # Compare dates and store matching dates in a dict to send by email
    matching_dict = {}
    for key, value in data_dict.items():
         date_str = value[0] # the date is in the second column of the tuple
         date_obj = datetime.strptime(date_str, '%d-%m-%Y').date()
         current_date = date.today()
         this_week_start = current_date - timedelta(days=current_date.weekday())
         this_week_end = this_week_start + timedelta(days=6)
         passed_date = []

         if date_obj < current_date:
                passed_date.append({
                    'deleteDimension': {
                         'range': {
                             'sheetId': 0,
                            'dimension': 'ROWS',
                            'startIndex': key,
                            'endIndex': key + 1
                        }
                    }
                })
        
         # If there are any rows to delete, update the sheet
         if passed_date:
          body = {
             'requests': passed_date
         }
         service.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body=body).execute()
        #-----------------------------------------------

         if date_obj == current_date or (this_week_start <= date_obj <= this_week_end):
             matching_dict[key] = date_str

    # Set up email details
    email_sender = os.environ.get("EMAIL_SENDER")
    email_receiver = os.environ.get("EMAIL_RECEIVER")
    email_subject = "Comida que vai expirar esta semana!"
    email_body = "Os seguintes aliementos vão ficar fora do prazo esta semana: \n\n"
    for item, item_info in matching_dict.items():
         email_body += item + " - " + item_info + "\n\n"

    # Set up SMTP server details
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    smtp_username = email_sender
    smtp_password = os.environ.get("PASS_SENDER")

    # Create message object
    message = MIMEMultipart()
    message["From"] = email_sender
    message["To"] = email_receiver
    message["Subject"] = email_subject
    message.attach(MIMEText(email_body))

    # Send message using SMTP server
    smtp_connection = smtplib.SMTP(smtp_server, smtp_port)
    smtp_connection.starttls()
    smtp_connection.login(smtp_username, smtp_password)
    smtp_connection.sendmail(email_sender, email_receiver, message.as_string())
    smtp_connection.quit()

    print(matching_dict)
    return matching_dict

get_sheet()
