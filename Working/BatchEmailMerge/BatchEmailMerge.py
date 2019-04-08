#! python3
# BatchEmailMerge.py - Create email draft by using word template according to the information in excel file (for example : Email, Name, attachment, and field)

import os.path, send2trash, mammoth
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
word = win32.DispatchEx("Word.Application")
#outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")\
word.Visible = False
excel.Visible = False

# TODO: Get information from excel for Email, Name and field
EmailList = os.getcwd() + '\EmailList.xlsx'
wb1 = excel.Workbooks.Open(EmailList)

ws1 = wb1.Worksheets('EmailList')

# List for column number for Field and Attachment
FieldList = []
AttachList = []
# Initital value for y, Field number and Attach number for loop function
y = 1
FieldNum = 1
AttachNum = 1
# To find the column number for Field and Attachment and add to the FieldList and AttachList
while True:
    # If the column header is empty, exit the loop
    if str(ws1.Cells(3, y).Value) == 'None':
        break
    # Find the column number for Field to replace
    if ws1.Cells(3, y).Value == 'Field' + str(FieldNum):
        FieldList.append(y)
        FieldNum += 1
    # Find the column number for Attachment
    if ws1.Cells(3, y).Value == 'Attach' + str(AttachNum):
        AttachList.append(y)
        AttachNum += 1
    y += 1

# Email draft need to create if Batch name match
BatchToRun = ws1.Cells(1, 2).Value

AssistantList = {}
# x is to find the End of the row number
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
        
        # Open word template to replace
        doc = word.Documents.Open(os.getcwd() + '\\EmailTemplate.docx', False, True)
        Num = 1
        for field in FieldList:
            word.Selection.Find.Execute(str('<Field' + str(Num) + '>'), False, False, False, False, False, True, 1, True, str(ws1.Cells(x, field).Value), 2)
            Num += 1
        # TODO: Amend the file name for Template file
        doc.SaveAs('D:\\Python\\Additional\\Email\\BatchEmailMerge\\Template' + Name + '.docx', FileFormat=12)
        doc.Close()
        
        # Convert the Word file into HTML text
        with open('D:\\Python\\Additional\\Email\\BatchEmailMerge\\Template' + Name + '.docx', 'rb') as docx_file:
            result = mammoth.convert_to_html(docx_file)
            html = result.value
        
        # TODO: Amend the file name for Template file
        #send2trash.send2trash('D:\\Python\\Additional\\Email\\BatchEmailMerge\\Template' + Name + '.docx')
        
        
        # TODO: Create email draft and use HTML as body
        
    x += 1
