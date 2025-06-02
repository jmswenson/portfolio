#!/usr/bin/env python
# coding: utf-8

"""
Gmail and Google Calendar Integration Script

This script automates the process of fetching specific confirmation emails from Gmail 
based on a search query and creating corresponding events in Google Calendar. It uses 
Google's APIs for Gmail and Calendar to perform authentication, retrieve email data, 
extract event details, and schedule calendar events.

Features:
- OAuth 2.0 authentication for Gmail and Google Calendar
- Fetches emails matching a user-defined query
- Extracts event details from email subject lines
- Creates calendar events with attendee information
- Prevents duplicate event creation by checking existing events

Usage:
    python LVLUP_final-gmail-gcal-interface-cli.py --email1 <email1> --email2 <email2> 
        --query <gmail_search_query> --max_messages <number_of_emails>

Arguments:
- email1: The first email address to invite to the event (default: 'REDACTED').
- email2: The second email address to invite to the event (default: 'REDACTED').
- query: Gmail search query to filter emails (default: 'subject:"Registration Confirmation: Beginners (White/Orng/Yellow)"').
- max_messages: Maximum number of email messages to process (default: 6).

Requirements:
- Google OAuth credentials file ('credentials.json') in the same directory.
- Python packages: google-auth, google-auth-oauthlib, google-api-python-client, pytz.

Author: Joel Swenson
Date: November 25, 2024
"""


import os.path
import argparse
import sys
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime, timedelta
import pytz

# If modifying these SCOPES, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly',
          'https://www.googleapis.com/auth/calendar']

def authenticate_gmail_and_calendar():
    """
    Authenticate the user for Gmail and Google Calendar services.
    
    This function handles the OAuth 2.0 flow, including token refresh and new authentication if necessary.
    It saves the credentials to a file for future use.

    Returns:
        tuple: Gmail service object and Calendar service object
    """
    creds = None
    token_path = 'token.json'
    
    # Load existing credentials if available
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f'Error refreshing token: {e}')
                handle_refresh_error(token_path)
                creds = None
        if not creds:
            try:
                print("No valid credentials found. Starting new authentication flow.")
                print("A browser window should open automatically. If it doesn't, please check your console for a URL to visit.")
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
                # Save the credentials for the next run
                with open(token_path, 'w') as token:
                    token.write(creds.to_json())
                print("Authentication successful. Credentials saved for future use.")
            except FileNotFoundError:
                print("Error: 'credentials.json' file not found. Please make sure it's in the same directory as the script.")
                sys.exit(1)
    else:
        print("Existing credentials are valid.")

    try:
        gmail_service = build('gmail', 'v1', credentials=creds)
        calendar_service = build('calendar', 'v3', credentials=creds)
        return gmail_service, calendar_service
    except HttpError as error:
        print(f'An error occurred while building the service: {error}')
        sys.exit(1)

def handle_refresh_error(token_path):
    """
    Handle token refresh errors by deleting the invalid token file.

    Args:
        token_path (str): Path to the token file
    """
    if os.path.exists(token_path):
        os.remove(token_path)
        print(f"Deleted the expired or invalid token file: {token_path}")
    else:
        print(f"Token file {token_path} does not exist.")

def get_confirmation_emails(gmail_service, query, max_messages):
    """
    Retrieve confirmation emails from Gmail based on the given query.

    Args:
        gmail_service: Authenticated Gmail service object
        query (str): Gmail search query
        max_messages (int): Maximum number of messages to retrieve

    Returns:
        list: List of message objects
    """
    try:
        results = gmail_service.users().messages().list(userId='me', q=query, maxResults=max_messages).execute()
        messages = results.get('messages', [])
        
        if not messages:
            print(f'No messages found matching the query: "{query}"')
            print('Please check if the subject line in your Gmail search query is correct.')
            return []
        else:
            print(f'Found {len(messages)} messages.')
            return messages
    except HttpError as error:
        print(f'An error occurred while fetching emails: {error}')
        return []

def parse_event_time(event_time_str):
    """
    Parse the event time string into a datetime object.

    Args:
        event_time_str (str): Event time string

    Returns:
        datetime: Parsed datetime object or None if parsing fails
    """
    formats = [
        "%A, %B %d, %Y %I:%M %p",
        "%A, %b %d, %Y %I:%M %p",
        "%a, %B %d, %Y %I:%M %p",
        "%a, %b %d, %Y %I:%M %p"
    ]
    
    for fmt in formats:
        try:
            return datetime.strptime(event_time_str, fmt)
        except ValueError:
            continue
    
    print(f"Error: Unable to parse time data '{event_time_str}'")
    print(f"Expected formats are: {', '.join(formats)}")
    return None

def extract_event_details(msg):
    """
    Extract event details from an email message.

    Args:
        msg (dict): Email message object

    Returns:
        dict: Event details including name and time, or None if extraction fails
    """
    subject_line = None

    for header in msg['payload']['headers']:
        if header['name'] == 'Subject':
            subject_line = header['value']
            break

    if not subject_line:
        print("Error: Unable to find subject line in the email.")
        return None

    try:
        print("Subject Line:")
        print(subject_line)
        split_subject = subject_line.split("Registration Confirmation: ")
        if len(split_subject) > 1:
            event_details = split_subject[1].split(" on ")
            if len(event_details) != 2:
                print("Error: Subject line does not contain expected 'on' separator.")
                return None
            gcal_event_name = event_details[0].strip()
            gcal_event_time = event_details[1].strip()
            
            print(f"Extracted Event Name: {gcal_event_name}")
            print(f"Extracted Event Time: {gcal_event_time}")

            event_time_obj = parse_event_time(gcal_event_time)
            if event_time_obj is None:
                return None
            event_time_iso = event_time_obj.isoformat()

            return {
                'event_name': gcal_event_name,
                'event_time': event_time_iso
            }
        else:
            print("Error: The subject line does not contain the expected 'Registration Confirmation: ' format.")
            return None
    except Exception as e:
        print(f"Error while extracting event details: {e}")
        return None

def check_existing_event(calendar_service, event_name, start_time_iso, end_time_iso):
    """
    Check if an event already exists in the calendar.

    Args:
        calendar_service: Authenticated Calendar service object
        event_name (str): Name of the event
        start_time_iso (str): Start time in ISO format
        end_time_iso (str): End time in ISO format

    Returns:
        list: List of existing events matching the criteria
    """
    events_result = calendar_service.events().list(
        calendarId='primary',
        timeMin=start_time_iso,
        timeMax=end_time_iso,
        q=event_name,
        singleEvents=True,
        orderBy='startTime'
    ).execute()
    
    existing_events = events_result.get('items', [])
    
    return existing_events

def create_calendar_event(calendar_service, event_details, email1, email2):
    """
    Create a new calendar event.

    Args:
        calendar_service: Authenticated Calendar service object
        event_details (dict): Event details including name and time
        email1 (str): Email of the first attendee
        email2 (str): Email of the second attendee

    Returns:
        dict: Created event object or None if creation fails
    """
    chicago_tz = pytz.timezone('America/Chicago')
    
    start_time_obj = datetime.fromisoformat(event_details['event_time'])
    start_time_obj = start_time_obj.astimezone(chicago_tz)
    
    end_time_obj = start_time_obj + timedelta(hours=1)
    
    start_time_iso = start_time_obj.isoformat()
    end_time_iso = end_time_obj.isoformat()
    
    existing_events = check_existing_event(calendar_service, event_details['event_name'], start_time_iso, end_time_iso)
    
    if existing_events:
        print(f'Event "{event_details["event_name"]}" already exists at this time.')
        return None

    event = {
        'summary': event_details['event_name'],
        'start': {
            'dateTime': start_time_iso,
            'timeZone': 'America/Chicago',
        },
        'end': {
            'dateTime': end_time_iso,
            'timeZone': 'America/Chicago',
        },
        'attendees': [
            {'email': email1},
            {'email': email2},
        ],
    }

    try:
        event = calendar_service.events().insert(calendarId='primary', body=event).execute()
        print(f'Event created: {event.get("htmlLink")}')
        return event
    except HttpError as error:
        print(f'An error occurred while creating the event: {error}')
        return None

def main(email1, email2, query, max_messages):
    """
    Main function to orchestrate the email fetching and calendar event creation process.

    Args:
        email1 (str): Email of the first attendee
        email2 (str): Email of the second attendee
        query (str): Gmail search query
        max_messages (int): Maximum number of messages to process
    """
    try:
        gmail_service, calendar_service = authenticate_gmail_and_calendar()
        messages = get_confirmation_emails(gmail_service, query, max_messages)

        if not messages:
            print("No messages found. Exiting program.")
            return

        events_created = 0
        for message in messages:
            try:
                msg = gmail_service.users().messages().get(userId='me', id=message['id']).execute()
                event_details = extract_event_details(msg)
                if event_details:
                    event = create_calendar_event(calendar_service, event_details, email1, email2)
                    if event:
                        events_created += 1
            except HttpError as error:
                print(f'An error occurred while processing a message: {error}')

        print(f"Process completed. {events_created} events were created.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Gmail and Google Calendar Integration Script")
    parser.add_argument("--email1", default="REDACTED", help="First email address to invite to the event")
    parser.add_argument("--email2", default="REDACTED", help="Second email address to invite to the event")
    parser.add_argument("--query", 
                        default='subject:"Registration Confirmation: Beginners (White/Orng/Yellow)"', 
                        help="Gmail search query to find specific confirmation emails. "
                             "This should match the subject line of the emails you're looking for. "
                             "Example: 'subject:\"Registration Confirmation: Beginners (White/Orng/Yellow)\"'")
    parser.add_argument("--max_messages", type=int, default=6, 
                        help="Maximum number of email messages to process (default: 6)")

    args = parser.parse_args()

    main(args.email1, args.email2, args.query, args.max_messages)


