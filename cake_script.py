import imaplib
import email
import re
import os
import datetime
import pytz
from datetime import timedelta
from email.header import decode_header
import pickle
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from bs4 import BeautifulSoup


# Securely access environment variables
EMAIL_ACCOUNT = os.getenv("EMAIL_ACCOUNT")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

# Google Calendar API Credentials
SCOPES = ["https://www.googleapis.com/auth/calendar"]
CREDENTIALS_FILE = "credentials.json"  # Ensure this file is in the same directory

def authenticate_google_calendar():
    """Authenticate using OAuth 2.0 and return the Google Calendar service object."""
    
    creds = None
    # Token file stores user's access & refresh tokens
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials, authenticate via OAuth
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())  # Refresh token if expired
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)

        # Save the credentials for future use
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)

    service = build("calendar", "v3", credentials=creds)
    return service

def add_event_to_calendar(service, event_details):
    """Create an event in Google Calendar."""
    event = {
        "summary": f"Cake Pickup for {event_details['customer_name']}",
        "description": f"Cake Order: {event_details['cake_type']}",
        "start": {
            "dateTime": event_details["pickup_datetime"].isoformat(),
            "timeZone": "America/New_York",
        },
        "end": {
            "dateTime": (event_details["pickup_datetime"] + datetime.timedelta(hours=1)).isoformat(),
            "timeZone": "America/New_York",
        },
        "reminders": {
            "useDefault": False,
            "overrides": [
                {"method": "popup", "minutes": 60},  # Reminder 1 hour before
                {"method": "popup", "minutes": 15},  # Reminder 15 minutes before
            ],
        },
    }

    event = service.events().insert(calendarId="primary", body=event).execute()
    print(f"Event created: {event.get('htmlLink')}")

def extract_cake_orders():
    print("Extracting cake orders from Gmail...")
    
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
        mail.select("inbox")
        print("Logged in and inbox selected.")
        
        subject = "Hawk Delights LLC CAKE ORDER FORM Completed"
        status, messages = mail.search(None, f'SUBJECT "{subject}"')
        
        # Check if any emails were found
        num_emails = len(messages[0].split())
        if num_emails == 0:
            print(f"No emails found with the subject: {subject}")
            return []

        print(f"Found {num_emails} emails matching the subject.")
        
        orders = []
        for email_id in messages[0].split():
            try:
                status, msg_data = mail.fetch(email_id, "(RFC822)")
                if status != "OK":
                    print(f"Failed to fetch email {email_id}")
                    continue  # Skip to next email if fetching fails

                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])

                        # Extract email content
                        email_body = ""
                        if msg.is_multipart():
                            for part in msg.walk():
                                if part.get_content_type() == "text/plain":
                                    email_body = part.get_payload(decode=True).decode(errors='ignore')
                            
                        else:
                            email_body = msg.get_payload(decode=True).decode(errors='ignore')

                        # Extract date using BeautifulSoup to remove any HTML content
                        soup = BeautifulSoup(email_body, 'html.parser')
                        text_content = soup.get_text()

                        # Extract pickup date/time using regex
                        pickup_match = re.search(r"Pick Up Date/Time\s*([\s\S]*?)\s*(\d{1,2}/\d{1,2}/\d{2,4})", text_content)
                        customer_match = re.search(r"Customer Name\s*([\s\S]*?)\s*(\d{1,2}/\d{1,2}/\d{2,4})", text_content)
                        cake_match = re.search(r"Cake Type\s*([\s\S]*?)\s*(\d{1,2}/\d{1,2}/\d{2,4})", text_content)

                        if pickup_match and customer_match and cake_match:
                            pickup_datetime_str = pickup_match.group(1).strip()
                            customer_name = customer_match.group(1).strip()
                            customer_name = customer_name.split('\n\r\n')[0].strip()
                            cake_type = cake_match.group(1).strip()
                            parts = re.split(r'\s{2,}', cake_type) 
                            cake_type = parts[0].strip()

                        
                            # Clean pickup_datetime_str by removing HTML tags and extra spaces
                            pickup_datetime_str = pickup_datetime_str.replace('<td>', '').replace('</td>', '').strip()
                            pickup_datetime_str = pickup_datetime_str.replace(" @ ", " ")  # Remove '@' for clean datetime 
                        
                            # Extract the correct date and customer_name
                            pickup_datetime_str, customer_nme = (lambda s: (s.split("\n")[0].strip(), s.split("\n")[1].strip()))(pickup_datetime_str)
                            
                            try:
                        # Convert to datetime format
                                pickup_datetime = datetime.datetime.strptime(pickup_datetime_str, "%a %b %d, %Y %I:%M %p")
                                pickup_datetime = pytz.timezone("America/New_York").localize(pickup_datetime)
        
                                # Append order to orders list
                                orders.append({
                                    "pickup_datetime": pickup_datetime,
                                    "customer_name": customer_name,
                                    "cake_type": cake_type
                                })
                            except ValueError:
                                print(f"Skipping order due to invalid date format: {pickup_datetime_str}")
                                continue
        
            except Exception as e:
                print(f"Error processing email {email_id}: {e}")
                    
        return orders

    except Exception as e:
        print(f"An error occurred: {e}")
        return []

    finally:
        try:
            mail.logout()  # Ensure logout even if error occurs
            print("Logged out successfully.")
        except Exception as e:
            print(f"Error logging out: {e}")
        
def main():
    """Main function to extract cake orders and add them to Google Calendar."""
    print("Extracting cake orders from Gmail...")
    orders = extract_cake_orders()

    if not orders:
        print("No new cake orders found.")
        return
    for order in orders:
        print(order)
    print(f"Found {len(orders)} orders. Adding to Google Calendar...")
    service = authenticate_google_calendar()

    for order in orders:
        add_event_to_calendar(service, order)

    print("All orders have been added to Google Calendar!")
    
def main2():
    print(count_cake_orders())
    

if __name__ == "__main__":
    main()
