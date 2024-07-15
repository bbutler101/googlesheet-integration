import os
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Define the scope and credentials for Google Sheets API
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'path/to/your/service_account.json'

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Create the Sheets API client
service = build('sheets', 'v4', credentials=credentials)

# Define spreadsheet ID and range names
SPREADSHEET_ID = 'your_spreadsheet_id'
RANGE_NAME = 'Sheet1!A1:Z1000'
BACKUP_RANGE_NAME = 'BackupSheet!A1:Z1000'

def read_sheet():
    """Read data from Google Sheet."""
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    values = result.get('values', [])
    return pd.DataFrame(values[1:], columns=values[0])

def write_sheet(df):
    """Write data to Google Sheet."""
    values = [df.columns.values.tolist()] + df.values.tolist()
    body = {'values': values}
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME,
        valueInputOption='RAW', body=body).execute()

def backup_sheet(df):
    """Backup data to another sheet."""
    values = [df.columns.values.tolist()] + df.values.tolist()
    body = {'values': values}
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID, range=BACKUP_RANGE_NAME,
        valueInputOption='RAW', body=body).execute()

def delete_duplicates(df):
    """Delete duplicate rows based on specific columns."""
    return df.drop_duplicates(subset=['Column1', 'Column2'])

def sort_sheet(df):
    """Sort the sheet based on a specific column."""
    return df.sort_values(by=['Column3'], ascending=True)

def calculate_scores(df):
    """Calculate scores based on weights for specific columns."""
    df['Score'] = df['Column4'].astype(float) * 0.5 + df['Column5'].astype(float) * 0.3 + df['Column6'].astype(float) * 0.2
    return df

def main():
    # Step 1: Read data from Google Sheet
    df = read_sheet()

    # Step 2: Delete duplicates
    df = delete_duplicates(df)

    # Step 3: Sort the sheet
    df = sort_sheet(df)

    # Step 4: Calculate scores
    df = calculate_scores(df)

    # Step 5: Write data back to Google Sheet
    write_sheet(df)

    # Step 6: Backup the data
    backup_sheet(df)

if __name__ == '__main__':
    main()
