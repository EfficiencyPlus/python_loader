# Import required libraries
import openpyxl
import os
from datetime import datetime, date, time
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.oauth2 import service_account

# Define the path to the Excel file using file dialog box
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()
excel_file = filedialog.askopenfilename()

# Load the Excel file and select the first sheet
wb = openpyxl.load_workbook(excel_file)
ws = wb.worksheets[0]

# Create an empty list to store the values
values = []

# Loop through the rows and columns and append the values to the list
for row in ws.iter_rows():
    row_values = []
    for cell in row:
        if isinstance(cell.value, (datetime, date)):
            row_values.append(cell.value.strftime("%d/%m/%Y %H:%M:%S"))
        elif isinstance(cell.value, time):
            row_values.append(cell.value.strftime("%H:%M:%S"))
        else:
            row_values.append(cell.value)
    values.append(row_values)

# Define Google Sheets variables and credentials
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
KEY = 'key.json'
SPREADSHEET_ID ='1QeJpkDbajigV-RBy7Lkb70eM03nm-l0kwc-QpQT6ChQ'

creds = service_account.Credentials.from_service_account_file(KEY, scopes=SCOPES)
service = build('sheets', 'v4', credentials=creds)

# Define the range of cells to update
range_ = ws.calculate_dimension()

# Update the sheet with the values
try:
    result = service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_,
        valueInputOption='USER_ENTERED',
        body={
            'values': values
        }
    ).execute()
    print(f"{result['updatedCells']} cells updated.")
except Exception as e:
    print(f"Error al insertar los datos: {e}")
