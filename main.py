def load_File():
    import openpyxl
    from datetime import datetime, time
    import os

    # Define variables
    log_file_path = os.path.join(os.path.dirname(__file__), "Upload_Logs.txt")
    excel_dir = os.path.dirname(os.path.abspath(__file__))
    SPREADSHEET_ID ='1iuqRmwP0j9xeynqYMsHkjpdcYXIQWKr8AfAI_Np8v40'
    tabname ="Cheques"

    rows = []

    print("""  
********************************************************************
  _____ _____ _____ ___ ____ ___ _____ _   _  ______   __        
 | ____|  ___|  ___|_ _/ ___|_ _| ____| \ | |/ ___\ \ / /    _   
 |  _| | |_  | |_   | | |    | ||  _| |  \| | |    \ V /   _| |_ 
 | |___|  _| |  _|  | | |___ | || |___| |\  | |___  | |   |_   _|
 |_____|_|   |_|   |___\____|___|_____|_| \_|\____| |_|     |_|  
 
********************************************************************
 """)
    print(f"Leyendo datos del archivo de Excel desde {excel_dir}")

    try:
        for filename in os.listdir(excel_dir):
            if filename.endswith('.xlsx'):
                excel_file = os.path.join(excel_dir, filename)
                wb = openpyxl.load_workbook(excel_file)
                sheet = wb.active
                for row in sheet.iter_rows(values_only=True):
                    row_dict = {}
                    for idx, cell in enumerate(row):
                        if idx != 2:
                            if isinstance(cell, (int, float)):
                                cell = str(cell)
                            else:
                                cell = str(cell)
                        else:
                            if isinstance(cell, datetime):
                                cell = cell.strftime('%Y-%m-%d %H:%M:%S')
                            else:
                                cell = str(cell).split()[0]
                        if isinstance(cell, time):
                            cell = cell.strftime('%H:%M:%S')
                        row_dict[sheet.cell(row=1, column=idx + 1).value] = cell
                    rows.append(row_dict)

        print("Autenticando en Google Sheets")
        from google.oauth2.credentials import Credentials
        from googleapiclient.discovery import build
        from google.oauth2 import service_account

        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        KEY = os.path.join(os.path.dirname(__file__), "key.json")

        creds = service_account.Credentials.from_service_account_file(KEY, scopes=SCOPES)

        service = build('sheets', 'v4', credentials=creds)
        sheets_api = service.spreadsheets()

        data = []
        headers = list(rows[0].keys())
        for row in rows:
            data.append([row[header] for header in headers])

        result = sheets_api.values().get(spreadsheetId=SPREADSHEET_ID, range=tabname).execute()
        existing_data = result.get('values', [])

        existing_unique_ids = set()
        for row in existing_data:
            existing_unique_ids.add(row[1])

        new_data = []
        for row in data:
            if row[1] not in existing_unique_ids:
                new_data.append(row)

        if new_data:
            try:
                append_result = sheets_api.values().append(
                    spreadsheetId=SPREADSHEET_ID,
                    range=f'{tabname}!A1',
                    valueInputOption='USER_ENTERED',
                    insertDataOption='INSERT_ROWS',
                    body={'values': new_data}
                ).execute()
                rowNum = append_result['updates']['updatedRows']
                print(f"{rowNum} filas insertadas")
            except Exception as e:
                print(f"Error al insertar los datos: {e}")
                with open(log_file_path, "a") as log_file:
                    log_file.write(f"Execution Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}: ")
                    log_file.write(f"Error al insertar los datos: {e}.\n")
                    log_file.write("\n")
            else:
                print("Datos insertados correctamente!")
                with open(log_file_path, "a") as log_file:
                    log_file.write(f"Execution Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}: ")
                    log_file.write(f"{rowNum} filas insertadas\n")
                    log_file.write("\n")
        else:
            print("No hay datos nuevos para insertar")
            with open(log_file_path, "a") as log_file:
                log_file.write(f"Execution Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}: ")
                log_file.write("No hay datos nuevos para insertar\n")
                log_file.write("\n")
    except Exception as e:
        print(f"Error al cargar el archivo de Excel: {e}")
        with open(log_file_path, "a") as log_file:
            log_file.write(f"Execution Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}: ")
            log_file.write(f"Error al cargar el archivo de Excel: {e}.\n")
            log_file.write("\n")

    import time
    print("Esta ventana se cerrará en 3 segundos")
    time.sleep(3)

# Call the load_File function directly
load_File()