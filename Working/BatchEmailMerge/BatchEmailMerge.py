#! python3
# BatchEmailMerge.py - Create email draft by using word template according to the information in excel file (for example : Email, Name, attachment, and field)

import os.path
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
#outlook = win32.client.Dispatch("Outlook.Application").GetNamespace("MAPI")\


# TODO: Get information from excel for Email, Name and field
EmailList = os.getcwd() + '\EmailList.xlsx'
wb1 = excel.Workbooks.Open(EmailList)
excel.Visible = True
ws1 = wb1.Worksheets('EmailList')

# Email draft need to create if Batch name match
BatchToRun = ws1.Cells(1, 2).Value

AssistantList = {}
x = 4
while True:
    # If the Assistant field is empty, exit the loop
    if str(ws1.Cells(x, 2).Value) == 'None':
        break
    # If Batch field is empty will pass the loop
    if ws1.Cells(x, 1).Value == BatchToRun:
        Name = ws1.Cells(x, 2).Value
        # Create Assistant Folder if it does not exist
        if Name not in AssistantList:
            os.mkdir(Name)
            AssistantList.setdefault(Name, str(os.path.abspath(Name)))
        # Create 
        

    x += 1


# TODO: Create File for each assistant

# TODO: Use word template location to create draft email

# TODO: Convert the information to Word templatee

# TODO: Create email draft