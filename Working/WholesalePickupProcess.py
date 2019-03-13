#! python3
# WholesalePickupProcess - Move the file from revenue folder to specific folder, copy the data in excel
# Paste it in the report

import os.path, shutil, send2trash
from pathlib import Path
import win32com.client as win32
from win32com.client import constants
#excel = win32.gencache.EnsureDispatch('Excel.Application')
excel = win32.DispatchEx("Excel.Application")
from ctypes import windll
import PathNPassword


# Import Revenue file location from PathNPassword
VMRH = PathNPassword.VMRH
SCC = PathNPassword.SCC
PARIS = PathNPassword.PARIS
OriginalPath = PathNPassword.OriginalPath

# Import the path to store Revenue data from PathNPassword
Revenue_Data = PathNPassword.Revenue_Data

# Import the password and working file path from PathNPassword
Working_File_Path = PathNPassword.Working_File_Path
password = PathNPassword.password

# Revenue File path for data to be copied
LocationOfFile = []
    
# Get the last modified file path
def FileLocation(path1):
    last_modified_file = ''
    n = 0
    for filename in os.listdir(path1):
        filename = str(path1) + "\\" + str(filename)
        filename = Path(filename)
        time = filename.stat().st_mtime
        if time > n:
            n = time
            last_modified_file = filename
    return last_modified_file

# Loop four properties for file.
for path2 in OriginalPath:
    # Convert to Absolute path and added to List
    LocationOfFile.append(os.path.abspath(FileLocation(path2)))

# Remove last week's excel file in the specific folder
for file in os.listdir(Revenue_Data):
    send2trash.send2trash(Revenue_Data + file)

# Move four files to specific folder
for path3 in LocationOfFile:
    shutil.copy(path3, Revenue_Data)

# Find the latest working file and the Filename
Working_File = FileLocation(Working_File_Path)
Working_Filename = os.path.basename(Working_File)


# Dictionary to search for column in our working file 
Month = {'January': '4', 'February': '5', 'March': '6', 'April': '7', 'May': '8', 'June': '9',
         'July': '10', 'August': '11', 'September': '12', 'October': '13', 'November': '14', 'December': '15'}

# Find the correct column, then copy data in revenue data and paste it in our working file
def CopyNPaste(ws, Prop):
    RowRange = ws.Range("A1:A200")
    # Find the correct column in the revenue report to copy
    for i, value in enumerate(RowRange):
        if "Business pick" in str(value):
            LastCellNumber =  i
            break
    
    # Select the correct Worksheet in Working File
    wsWF = wb1.Worksheets(Prop)
    
    # Find the Month Column Number in Month dictionary
    CurrentMonth = (ws.Range("D2").Value).split(" ")[0]
    MonthColumn = Month[CurrentMonth]
    ColumnToPaste = [2, 3]
    for i in range(4):
        ColumnToPaste.append(int(MonthColumn) + i)
    
    # Find the correct column number for copy and find the correct column number to paste
    for Copy, Paste in zip(ColumnToCopy, ColumnToPaste):
        CopiedCell = ws.Range(ws.Cells(4, Copy), ws.Cells(LastCellNumber, Copy))
        CopiedCell.Copy()
        wsWF.Range(wsWF.Cells(3, Paste), wsWF.Cells(LastCellNumber-3, Paste)).PasteSpecial(Paste = constants.xlPasteValues, Operation = constants.xlNone)
        # Clear Clipboard (for warning window appear while running)
        if windll.user32.OpenClipboard(None):
            windll.user32.EmptyClipboard()
            windll.user32.CloseClipboard()

# Open our working file by using password
wb1 = excel.Workbooks.Open(Working_File, False, False, None, password, password)
excel.Visible = True
# Run Excel Macro in Working file (Unhide all Worksheet and Clear Data)
excel.Run("Module1.ClearData")
        
# Open Revenue Report for copying data to our working file
# Seperate into two conditions as HIMCC and CMCC are in the same file
for path4 in os.listdir(Revenue_Data):
    ColumnToCopy = [1, 2, 6, 9, 12, 15]
    
    if "SCC" in path4:
        wb2 = excel.Workbooks.Open(Revenue_Data + "\\" + path4)
        # Select the correct worksheet for HIMCC and CMCC
        # Conrad Worksheet
        wsCM = wb2.Worksheets('Report - Conrad')
        CopyNPaste(wsCM, 'CMCC Raw')
        # Holiday Inn Worksheet
        wsHI = wb2.Worksheets('Report - Holiday Inn')
        CopyNPaste(wsHI, 'HIMCC Raw')
        
        wb2.Close(True)
    
    elif "Parisian" in path4:
        wb2 = excel.Workbooks.Open(Revenue_Data + "\\" + path4)
        # Parisian Worksheet
        wsPA = wb2.Worksheets('Report')
        CopyNPaste(wsPA, 'PARIS Raw')
        
        wb2.Close(True)
    else:
        wb2 = excel.Workbooks.Open(Revenue_Data + "\\" + path4)
        # Venetian Worksheet
        wsVE = wb2.Worksheets('Report')
        CopyNPaste(wsVE, 'VMRH Raw')
        
        wb2.Close(True)

# Save As the excel file 
New_Working_Filename = PathNPassword.New_Working_Filename
wb1.SaveAs(Working_File_Path + "\\" + New_Working_Filename)

#excel.Application.Quit()
