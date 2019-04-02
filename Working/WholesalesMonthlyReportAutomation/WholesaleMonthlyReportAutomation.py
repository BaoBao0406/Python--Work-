#! pyhon3
# WholesaleMonthlyReportAutomation.py - find the latest file, copy the correct worksheet to the
# working file and click the Marco button

import os.path, PathNPassword
import win32com.client as win32
from win32com.client import constants
excel = win32.DispatchEx("Excel.Application")
from ctypes import windll


# Import the password and working file path from PathNPassword
Working_File_Path = PathNPassword.Working_File_Path
Working_File = PathNPassword.Working_File
NewWorking_File = PathNPassword.NewWorking_File
password = PathNPassword.password

# Create an empty Wholesale Date excel file
Wholesale_Data_Path = PathNPassword.Wholesale_Data_Path
Wholesale_Data = PathNPassword.Wholesale_Data

# Properties name for short and long form in Working File
PropDict = {'VMRH': 'Venetian', 'CMCC': 'Conrad', 'HICC': 'Holiday', 'PARIS':'Parisian'}

# Create Worksheet for each Properties
wb_WSData = excel.Workbooks.Add()
# Open excel file by using Manual Calculation
excel.Calculation = -4135
excel.Visible = True
for prop in PropDict.keys():
    wsP = wb_WSData.Worksheets.Add()
    wsP.Name = prop
# Delete empty worksheets
for sheet in wb_WSData.Worksheets:
    if "Sheet" in sheet.Name:
        sheet.Delete()

# Import the Revenue Path according to month
Revenue_Path = PathNPassword.Revenue_Path

# Function to open, copy the sheet to our RawData worksheet
def CopyNPaste(filename, prop):
    wb1 = excel.Workbooks.Open(filename, False, False, None)
    ws1 = wb1.Worksheets('All_MTD')
    RowRange = ws1.Range("B1:B400")
    for i, value in enumerate(RowRange, 1):
        if "grand total" in str(value).lower():
            LastCellNum = i
            break

    CopyRange = ws1.Range(ws1.Cells(1,1), ws1.Cells(LastCellNum, 26))
    CopyRange.Copy()
    ws2 = wb_WSData.Worksheets(prop)
    ws2.Range(ws2.Cells(1,1), ws2.Cells(LastCellNum, 26)).PasteSpecial(Paste = constants.xlPasteValues, Operation = constants.xlNone)
    # Clear Clipboard (for warning window appear while running)
    if windll.user32.OpenClipboard(None):
        windll.user32.EmptyClipboard()
        windll.user32.CloseClipboard()
    wb1.Close(True)

# Run the function if filename equal to the RawData worksheet name 
for filename in os.listdir(Revenue_Path):
    for propShort, propLong in PropDict.items():
        if propLong in filename:
            CopyNPaste(Revenue_Path + filename, propShort)
# Save the Wholesale RawData in our path
wb_WSData.SaveAs(Wholesale_Data_Path + Wholesale_Data)

# Open working file with password
wb2 = excel.Workbooks.Open(Working_File_Path + Working_File, False, False, None, password, password)

# Copy all worksheet in RawData to WorkingFile
for sheet in wb_WSData.Worksheets:
    sheet.Copy(Before=wb2.Worksheets('Summary'))

wb_WSData.Close(True)

# Save the Wholesale Monthly report
wb2.SaveAs(Working_File_Path + NewWorking_File)
# Click the macro button
excel.Run("RunAllSteps")

#excel.Application.Quit()
