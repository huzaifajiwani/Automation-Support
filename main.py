import win32com.client
from win32com.client import constants as c
from datetime import datetime
import gspread
from google.oauth2 import service_account

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = service_account.Credentials.from_service_account_file("testing-project-405310-0116e9bc8128.json", scopes=scope)
gc = gspread.authorize(credentials)

spreadsheet = gc.open("test-1")

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace('MAPI')
inbox = namespace.GetDefaultFolder(c.olFolderInbox)
messages = inbox.Items

today = datetime.today().strftime('%Y-%m-%d')

messages = messages.Restrict("[UnRead] = false")

filtered_messages = [msg for msg in messages if msg.ReceivedTime.date() >= datetime.strptime(today, '%Y-%m-%d').date()]

try:
    worksheet = spreadsheet.sheet1
except gspread.exceptions.WorksheetNotFound:
    worksheet = spreadsheet.add_worksheet(title="Sheet1", rows="100", cols="20")

headers = ["Subject", "Sender Name", "Received Date", "CC", "Message To", "Priority"]

header_row = {header: index + 1 for index, header in enumerate(worksheet.row_values(1))}

default_values = {
    "id": "default_id_value",
    "reason": "default_reason_value",
}

for msg in filtered_messages:
    sender_name = msg.Sender.Name
    received_date = msg.ReceivedTime.replace(tzinfo=None).strftime('%Y-%m-%d %H:%M:%S')
    priority = msg.Importance

    if priority == c.olImportanceHigh:
        importance = "High"
    elif priority == c.olImportanceLow:
        importance = "Low"
    else:
        importance = "Medium"

    if "Technology" in msg.CC or "Technology" in msg.To:
        if "RE:" not in msg.Subject:
            row_values = [None] * len(headers)

            for header in headers:
                if header in header_row:
                    if header == "Subject":
                        row_values[header_row[header] - 1] = msg.Subject
                    elif header == "Sender Name":
                        row_values[header_row[header] - 1] = sender_name
                    elif header == "Received Date":
                        row_values[header_row[header] - 1] = received_date
                    elif header == "CC":
                        row_values[header_row[header] - 1] = msg.CC
                    elif header == "Message To":
                        row_values[header_row[header] - 1] = msg.To
                    elif header == "Priority":
                        row_values[header_row[header] - 1] = importance

            for col, value in default_values.items():
                if col in header_row:
                    row_values[header_row[col] - 1] = value

            existing_values = worksheet.col_values(header_row["Subject"])
            new_value = msg.Subject

            if new_value not in existing_values:
                worksheet.append_row(row_values)
