import json
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

SPREADSHEET_ID = '1SBPuMB1tSEX6CUsYmXuvRW1hKGDnx6KHL0LDE5amgvw'
READ_RANGE_NAME = 'A2:E'
WRITE_RANGE_NAME = 'E2:E'
VALUE_INPUT_OPTION = 'RAW'
STATUS_VALUE = 'yes'

def main():
    # Authentication
    credentials = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            credentials = pickle.load(token)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            credentials = flow.run_console()
        with open('token.pickle', 'wb') as token:
            pickle.dump(credentials, token)

    service = build('sheets', 'v4', credentials=credentials)

    # Read values
    sheet = service.spreadsheets()
    result = sheet.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=READ_RANGE_NAME,
    ).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
        return

    read_status_values = []
    for i, row in enumerate(values):
        if len(row) == 5 and row[4] == STATUS_VALUE:
            read_status_values.append([])
            continue

        print(f'{row[0]}, {row[1]}, {row[2]}, {row[3]}')
        read_status_values.append([STATUS_VALUE])

    # Write values
    result = service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=WRITE_RANGE_NAME,
        valueInputOption=VALUE_INPUT_OPTION,
        body=dict(values=read_status_values),
    ).execute()

if __name__ == '__main__':
    main()
