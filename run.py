from __future__ import print_function
from pprint import pprint

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from datetime import datetime
from datetime import timedelta
from collections import defaultdict
import json
# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly',
          'https://www.googleapis.com/auth/calendar']

# The ID and range of a sample spreadsheet.
RANGE_NAME = 'Week 5!A:E'

with open("./emails.json") as f:
    emails = json.loads(f.read())

with open("./config.json") as f:
    config = json.loads(f.read())

SPREADSHEET_ID = config['spreadsheet_id']
CALENDAR_ID = config['calendar_id']


def process(data):
    # map each name to list of (day, start_time, end_time, walkup_line) tuples
    sessions = defaultdict(list)
    most_recent_date = data[0][0]
    for row in data:
        if len(row) == 0:
            continue
        date = row[0]
        shift_length = row[1]
        shift = row[2]
        tenters = row[3:]
        if date != "":
            most_recent_date = date.split(" ")[0]

        if shift == "":
            continue
        for i, tenter in enumerate(tenters):
            start, stop = shift.split("-")
            if i > 1:
                sessions[tenter].append((most_recent_date, start, stop, True))
            else:
                sessions[tenter].append((most_recent_date, start, stop, False))

    return sessions


def create_cal_events(results, creds):
    service = build('calendar', 'v3', credentials=creds)
    # service.calendars().clear(calendarId=CALENDAR_ID).execute()
    VALID_DATES = ["2/23", "2/24"]
    for person in results:
        if person not in emails:
            continue
        for date, start_time, end_time, walkup in results[person]:
            month, day = date.split('/')
            month = int(month)
            day = int(day)
            if date not in VALID_DATES:
                continue
            # print(person, start_time, end_time)
            start_dt = datetime.strptime(start_time, "%I:%M%p")
            end_dt = datetime.strptime(end_time, "%I:%M%p")

            # if moving to next day:
            if 'am' in start_time and 'am' in end_time and start_time != '9:15am' and start_time != '7:00am':
                day += 1
            start_date = datetime(2022, month, day)
            start_date += timedelta(hours=start_dt.hour,
                                    minutes=start_dt.minute)
            if 'am' in end_time and 'am' not in start_time:
                day += 1
            end_date = datetime(2022, month, day)
            end_date += timedelta(hours=end_dt.hour, minutes=end_dt.minute)
            # print(start_date, end_date)
            if end_time == '9:15am':
                shift_type = "Night Tenting"
            elif walkup:
                shift_type = "Tenting WUL Day"
            else:
                shift_type = "Tenting Day"
            event = {
                'summary': f'{person} {shift_type} Shift',
                'location': 'Krzyzewskiville, 330 Towerview Rd, Durham, NC 27705, USA',
                'description': f'Tenting shift from {start_dt.strftime("%H:%M")}-{end_dt.strftime("%H:%M")}',
                'start': {
                    'dateTime': start_date.strftime('%Y-%m-%dT%H:%M:%S.%f%z'),
                    'timeZone': 'America/New_York',
                },
                'end': {
                    'dateTime': end_date.strftime('%Y-%m-%dT%H:%M:%S.%f%z'),
                    'timeZone': 'America/New_York'
                },
                'reminders': {
                    'useDefault': False,
                    'overrides': [
                    ]
                }
            }
            event['attendees'] = [{"email": emails[person]}]
            # print(event)
            event = service.events().insert(calendarId=CALENDAR_ID,
                                            body=event, sendUpdates='all').execute()
            print(
                f"Created {shift_type} shift for {person} from {start_time} to {end_time} on {start_date.strftime('%Y-%m-%d')}")


def main():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'client_secret.json', SCOPES)
                creds = flow.run_local_server(port=0)
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                    range=RANGE_NAME).execute()
        values = result.get('values', [])

        if not values:
            print('No data found.')
            return

        data = []
        for row in values[1:]:
            # Print columns A and E, which correspond to indices 0 and 4.
            data.append(row)
        results = process(data)
        create_cal_events(results, creds)
    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()
