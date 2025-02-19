import imaplib
import email
import re
import os
import datetime
import pytz
from email.header import decode_header
from google.oauth2 import service_account
from googleapiclient.discovery import build

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

def extract_cake_orders():
    """Extract cake orders from Gmail and return them as a list of dictionaries."""
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
        mail.select("inbox")

        # Search for emails from Jolt (info@joltup.com)
        status, messages = mail.search(None, 'FROM "info@joltup.com"')
        orders = []

        for email_id in messages[0].split():
            status, msg_data = mail.fetch(email_id, "(RFC822)")

            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])

                    # Extract email content
                    email_body = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            if part.get_content_type() == "text/plain":
                                email_body = part.get_payload(decode=True).decode(errors='ignore')
                                break
                    else:
                        email_body = msg.get_payload(decode=True).decode(errors='ignore')

                    # Extract relevant details
                    pickup_match = re.search(r"Pick Up Date/Time\s*(.*?)\s*\w+\s*\d{1,2}/\d{1,2}/\d{2,4}", email_body)
                    customer_match = re.search(r"Customer Name\s*(.*?)\s*\w+\s*\d{1,2}/\d{1,2}/\d{2,4}", email_body)
                    cake_match = re.search(r"Cake Type\s*(.*?)\s*\w+\s*\d{1,2}/\d{1,2}/\d{2,4}", email_body)

                    if pickup_match and customer_match and cake_match:
                        pickup_datetime_str = pickup_match.group(1).strip()
                        customer_name = customer_match.group(1).strip()
                        cake_type = cake_match.group(1).strip()

                        # Convert to datetime format
                        try:
                            pickup_datetime = datetime.datetime.strptime(pickup_datetime_str, "%a %b %d, %Y @ %I:%M %p")
                            pickup_datetime = pytz.timezone("America/New_York").localize(pickup_datetime)
                        except ValueError:
                            print(f"Skipping order due to invalid date format: {pickup_datetime_str}")
                            continue

                        orders.append({
                            "pickup_datetime": pickup_datetime,
                            "customer_name": customer_name,
                            "cake_type": cake_type
                        })

        return orders

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        mail.logout()

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

if __name__ == "__main__":
    main()
