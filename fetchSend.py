import ssl
import smtplib
from email.message import EmailMessage

email_sender = "ENTER EMAIL ADDRESS TO SEND FROM"
email_password = "GIVE PASSWORD (FROM APP PASSWORD SECTION)"
#ADD A "token.json" file in the same repository

import ssl
import smtplib
from email.message import EmailMessage
import random
import string
import os

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

spreadsheet_id = 'GIVE SPREADSHEED ID'

class SheetObject:
    def __init__(self, **kwargs):
        for key, value in kwargs.items():
            setattr(self, key, value)
def mail(student_info):
    # Customize the confirmation message using the student information
    registration_confirmation = f"""
Dear {student_info.name},

I hope this message finds you well. I would like to take this opportunity to extend my warmest greetings to you on the occasion of Eid ul-Adha, and I hope that this festive season brings immense joy and happiness to your life.

In the session of "{student_info.session}", from "{student_info.union}" union of Chatkhil Upazila, Noakhali, you got admitted into University of Dhaka in "{student_info.dept}" dapartment, making "{student_info.hall}" your secoond home . The journey to this achievement was undoubtedly challenging, and we commend your dedication and hard work. You have become an integral part of our esteemed community of Dhaka University students from Chatkhil, and your contributions have been highly appreciated.

I would like to draw your attention to an upcoming event that we are organizing in order to provide guidance and support to new admissions. The Admission Guideline Program Season 4 will be held at the Chatkhil Upazila Auditorium on July 3rd, 2023. We cordially invite you to attend this program, as your presence and participation will greatly contribute to our shared goal of increasing the number of students from Chatkhil who gain admission to various public universities.

We eagerly anticipate your presence and support at this event. Your valuable insights and efforts will play a vital role in achieving our objectives. Please consider this as a formal invitation, and we sincerely hope that you will join us.

Thank you for your attention and cooperation.

Event Details:
Date: 3rd July, 2023
Time: 10:00 AM
Venue: Chatkhil Upazilla Parishad Auditorium, Chatkhil, Noakhali

Yours faithfully,
Abdullah Ibne Hanif Arean
Department of Computer Science and Engineering
University of Dhaka
Joint General Secretary
Dhaka University Student's Association of Chatkhil (DUSACH)
    """


    em = EmailMessage()
    em['From'] = email_sender
    em['To'] = student_info.email
    em['Subject'] = "Invitation: Dhaka University Admission Test Guideline Program Season 4"
    em.set_content(registration_confirmation)

    context = ssl.create_default_context()

    with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
        smtp.starttls(context=context)
        smtp.login(email_sender, email_password)
        smtp.send_message(em)

def main():
    credentials = None
    if os.path.exists("token.json"):
        credentials = Credentials.from_authorized_user_file("token.json", SCOPES)
        print("Credentials Done!")
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("/Users/abdullaharean/Desktop/django/mailSender/token.json", SCOPES)
            credentials = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(credentials.to_json())
    try:
        service =build("sheets", "v4", credentials=credentials)
        sheets = service.spreadsheets()
        result = sheets.values().get(spreadsheetId = spreadsheet_id, range = "Sheet1!A1:L38").execute()
        values = result.get("values", [])
        header_row = values[0]
        objects = []

        for row in values[1:]:
            data = {}
            for i, value in enumerate(row):
                data[header_row[i]] = value

            obj = SheetObject(**data)
            objects.append(obj)

        for obj in objects:
            mail(obj)
            # if obj.name == "Abdullah Ibne Hanif Arean":
            #     print(obj.__dict__)
                

        
            
    except HttpError as error:
        print(error)
        
if __name__ == "__main__":
    main()