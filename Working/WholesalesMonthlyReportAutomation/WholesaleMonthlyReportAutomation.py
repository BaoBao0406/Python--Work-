#! pyhon3
# WholesaleMonthlyReportAutomation.py - find the latest file, copy the correct worksheet to the
# working file and click the Marco button

import os.path, shutil, PathNPassword
from pathlib import Path
import win32com.client as win32
from win32com.client import constants
excel = win32.DispatchEx("Excel.Application")
from ctypes import windll


# Import the password and working file path from PathNPassword
Working_File_Path = PathNPassword.Working_File_Path
Working_File = PathNPassword.Working_File
password = PathNPassword.password

# Create an empty Wholesale Date excel file
Wholesale_Data_Path = PathNPassword.Wholesale_Data_Path
Wholesale_Data = PathNPassword.Wholesale_Data

# Properties in Working File
Properties = ['VMRH', 'CMCC', 'HICC', 'PARIS']

# Create Worksheet for each Properties
wb_WSData = excel.Workbooks.Add()
excel.Calculation = -4135
excel.Visible = True
for prop in Properties:
    wsP = wb_WSData.Worksheets.Add()
    wsP.Name = prop
# Delete empty worksheets
for sheet in wb_WSData.Worksheets:
    if "Sheet" in sheet.Name:
        sheet.Delete()

#wb_WSData.SaveAs(Wholesale_Data_Path + Wholesale_Data)

# Import the Revenue Path according to month
Revenue_Path = PathNPassword.Revenue_Path

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

for filename in os.listdir(Revenue_Path):
    if "Venetian" in filename:
        CopyNPaste(Revenue_Path + filename, 'VMRH')
    elif "Conrad" in filename:
        CopyNPaste(Revenue_Path + filename, 'CMCC')
    elif "Holiday" in filename:
        CopyNPaste(Revenue_Path + filename, 'HICC')
    elif "Parisian" in filename:
        CopyNPaste(Revenue_Path + filename, 'PARIS')
"""
#wb_WSData.Close(True)

# Open working file with password
#wb2 = excel.Workbooks.Open(Working_File, False, False, None, password, password)
#excel.Visible = True

# TODO: Save as a new file

# TODO: Click the macro button

#excel.Application.Quit()
"""
