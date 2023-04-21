import win32com.client
import datetime

def send_mail(rowNum):
    today = datetime.datetime.today().strftime('%Y-%m-%d')
    a = 10
    b = 20
    ol=win32com.client.Dispatch("outlook.application")
    olmailitem=0x0 #size of the new email
    newmail=ol.CreateItem(olmailitem)
    newmail.Subject = f'Batch run results for {today}'
    newmail.To='lorenzo.orenday@efficiency-plus.com'
    newmail.CC='lorenzo.orenday@efficiency-plus.com'
    newmail.Body = f'Hello,\n\nHere are the batch results for today\'s run;\n\n Rows added {rowNum}\n\nSincerely, Efficiency + Team'
    # attach='C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
    # newmail.Attachments.Add(attach)
    # To display the mail before sending it
    # newmail.Display() 
    try:
        newmail.Send()
        print("Mail sent successfully")
    except Exception as e:
        print(f"Mail failed to send due to: {e}")

