import imaplib
import email
import re
import os
import datetime
import pytz
from email.header import decode_header
from google.oauth2 import service_account
from googleapiclient.discovery import build
from bs4 import BeautifulSoup


# Securely access environment variables
EMAIL_ACCOUNT = os.getenv("EMAIL_ACCOUNT")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

# Google Calendar API Credentials
SCOPES = ["https://www.googleapis.com/auth/calendar"]
CREDENTIALS_FILE = "credentials.json"  # Ensure this file is in the same directory

def authenticate_google_calendar():
    """Authenticate and return the Google Calendar service object."""
    creds = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    service = build("calendar", "v3", credentials=creds)
    return service

def add_event_to_calendar(service, event_details):
    """Create an event in Google Calendar."""
    event = {
        "summary": f"Cake Pickup for {event_details['customer_name']}",
        "location": "Hawk Delights LLC",
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

import re
from bs4 import BeautifulSoup
import datetime
import pytz

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
                        
                        # Debugging: Check the extracted email body
                        print(f"Email Body:\n{email_body}")
                        
                        # Extract text using BeautifulSoup
                        soup = BeautifulSoup(email_body, 'html.parser')
                        text_content = soup.get_text()
                        
                        # Debugging: Check the text content after BeautifulSoup processing
                        print(f"Text Content:\n{text_content}")
                        
                        # Extract relevant details
                        pickup_match = re.search(r"Pick Up Date/Time\s*([\w\s]+? \d{1,2}, \d{4} \d{1,2}:\d{2} [APM]{2})", text_content)
                        customer_match = re.search(r"Customer Name\s*([\s\S]*?)\s*(\d{1,2}/\d{1,2}/\d{2,4})", text_content)
                        cake_match = re.search(r"Cake Type\s*([\s\S]*?)\s*(\d{1,2}/\d{1,2}/\d{2,4})", text_content)
                        
                        # Debugging: Check what was matched
                        if pickup_match:
                            print(f"Pickup Date: {pickup_match.group(1)}")
                        else:
                            print("Pickup Date not found.")
                            
                        if customer_match:
                            print(f"Customer Name: {customer_match.group(1)}")
                        else:
                            print("Customer Name not found.")
                            
                        if cake_match:
                            print(f"Cake Type: {cake_match.group(1)}")
                        else:
                            print("Cake Type not found.")
                        if pickup_match and customer_match and cake_match:
                            pickup_datetime_str = pickup_match.group(1).strip()
                            customer_name = customer_match.group(1).strip()
                            cake_type = cake_match.group(1).strip()

                            # Clean up pickup_datetime_str and remove unnecessary parts
                            pickup_datetime_str = pickup_datetime_str.replace(" @ ", " ").replace(",", "")  # Clean format

                            # Convert to datetime format
                            try:
                                pickup_datetime = datetime.datetime.strptime(pickup_datetime_str, "%a %b %d %Y %I:%M %p")
                                pickup_datetime = pytz.timezone("America/New_York").localize(pickup_datetime)
                            except ValueError:
                                print(f"Skipping order due to invalid date format: {pickup_datetime_str}")
                                continue

                            orders.append({
                                "pickup_datetime": pickup_datetime,
                                "customer_name": customer_name,
                                "cake_type": cake_type
                            })

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
        
def count_cake_orders():
    """Count the number of cake orders with the specified subject."""
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
        mail.select("inbox")

        # Search for emails with the specific subject
        subject = "Hawk Delights LLC CAKE ORDER FORM Completed"
        status, messages = mail.search(None, f'SUBJECT "{subject}"')

        # Count the number of matching emails
        email_count = len(messages[0].split()) if status == "OK" else 0

        return email_count

    except Exception as e:
        print(f"An error occurred: {e}")
        return 0

    finally:
        mail.logout()  # Ensure logout even if error occurs
        
def main():
    """Main function to extract cake orders and add them to Google Calendar."""
    print("Extracting cake orders from Gmail...")
    orders = extract_cake_orders()

    if not orders:
        print("No new cake orders found.")
        return

    print(f"Found {len(orders)} orders. Adding to Google Calendar...")
    service = authenticate_google_calendar()

    for order in orders:
        add_event_to_calendar(service, order)

    print("All orders have been added to Google Calendar!")
def main2():
    print(count_cake_orders())
    

if __name__ == "__main__":
    main()
