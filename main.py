def load_File():
    from Send_Mail import send_mail
    # Read excel file
    import openpyxl
    from datetime import datetime, time
    import os

    # Define the path to the Excel files directory
    excel_dir = r"C:\Users\loren\Desktop\python_loader"

    # Set the log file path
    log_file_path = os.path.join(os.path.dirname(__file__), "Upload_Logs.txt")

    # Set the log file path
    log_file_path = os.path.join(os.path.dirname(__file__), "Upload_Logs.txt")

    # Create a list of dictionaries, where each dictionary represents a row in the Excel sheet
    rows = []
    print(f"Leyendo datos del archivo de Excel desde {excel_dir}")

    for filename in os.listdir(excel_dir):
        if filename.endswith('.xlsx'): # Check if the file is an Excel file
            excel_file = os.path.join(excel_dir, filename)
            
            # Load the Excel file and select the first sheet
            wb = openpyxl.load_workbook(excel_file)

            sheet = wb.active
            for row in sheet.iter_rows(values_only=True):
                row_dict = {}
                for idx, cell in enumerate(row):
                    if idx != 2:  # Exclude column C
                        if isinstance(cell, (int, float)):
                            cell = str(cell)
                        else:
                            cell = str(cell)
                    else:
                        if isinstance(cell, datetime):
                            cell = cell.strftime('%m/%d/%Y %H:%M:%S')
                        else:
                            # Handle the case if the cell value is not a datetime object
                            cell = str(cell).split()[0]
                    if isinstance(cell, time):
                        cell = cell.strftime('%H:%M:%S')
                    row_dict[sheet.cell(row=1, column=idx + 1).value] = cell
                rows.append(row_dict)


    # Write to G. sheets
    print("Autenticando en Google Sheets")
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
    from google.oauth2 import service_account

    # Define G. sheets variables & credentials
    current_directory = os.getcwd()
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    KEY = r'C:\Users\loren\Desktop\python_loader\key.json'
    SPREADSHEET_ID ='1iuqRmwP0j9xeynqYMsHkjpdcYXIQWKr8AfAI_Np8v40' #'1QeJpkDbajigV-RBy7Lkb70eM03nm-l0kwc-QpQT6ChQ'

    creds = None
    creds = service_account.Credentials.from_service_account_file(KEY, scopes=SCOPES)

    service = build('sheets', 'v4', credentials=creds)
    sheets_api = service.spreadsheets() # Rename the variable to avoid collision with sheet data

    # Convert the list of dictionaries to a list of lists
    data = []
    headers = list(rows[0].keys())
    for row in rows:
        data.append([row[header] for header in headers])

    # Retrieve the existing data from the sheet
    result = sheets_api.values().get(spreadsheetId=SPREADSHEET_ID, range='Cheques').execute()       
    existing_data = result.get('values', [])

    # Create a set of the existing unique IDs
    existing_unique_ids = set()
    for row in existing_data:
        existing_unique_ids.add(row[1])

    # Insert new data into the sheet
    print("Insertando datos en Google")
    new_data = []
    for row in data:
        if row[1] not in existing_unique_ids:
            new_data.append(row)

    if new_data:
        try:
            # Insert the new data into a new sheet in the Google Sheets document
            append_result = sheets_api.values().append(
                spreadsheetId=SPREADSHEET_ID,
                range='Cheques!A1',
                valueInputOption='USER_ENTERED',
                insertDataOption='INSERT_ROWS',
                body={'values': new_data}
            ).execute()
            rowNum = append_result['updates']['updatedRows']
            print(f"{rowNum} filas insertadas.")
        except Exception as e:
            print(f"Error al insertar los datos: {e}")
        else:
            print("Datos insertados correctamente!")
        finally:
            with open(log_file_path, "a") as log_file:
                log_file.write(f"Execution Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}: ")
                if 'rowNum' in locals():
                    log_file.write(f"{rowNum} filas insertadas.\n")
                else:
                    log_file.write("Error al insertar los datos.\n")
                log_file.write("\n")
    else:
        print("No hay datos nuevos para insertar")
        with open(log_file_path, "a") as log_file:
            log_file.write(f"Execution Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}: ")
            log_file.write("No hay datos nuevos para insertar\n")
            log_file.write("\n")


# Call the load_File function directly
load_File()
import time
print("Esta ventana se cerrará en 3 segundos")
time.sleep(3)