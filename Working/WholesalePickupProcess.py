#! python3
# WholesalePickupProcess - Move the file from revenue folder to specific folder, copy the data in excel
# Paste it in the report

import os.path, shutil, send2trash
from pathlib import Path
import win32com.client as win32
from win32com.client import constants
excel = win32.gencache.EnsureDispatch('Excel.Application')

"""
# TODO: Save the Path and Data in another python file

# Path for all files
VMRH = 'D:\\Python\\Book 3'
SCC = 'D:\\Python\\Book 4'
PARIS = 'D:\\Python\\Book 5'
OriginalPath = [VMRH, SCC, PARIS]
"""
# Path to store the Revenue Data
Revenue_Data = 'D:\\Python\\Testing'
"""
# Woring file path and password
Working_File = 'D:\\Python\\Additional\\WholesalePickupProcess\\Wholesale Pick-up ComparisonTesting2.xlsx'
password = 'venetian2019'

# Revenue File path for data to be copied
LocationOfFile = []
    
# Get the last modified file path
def FileLocation(path1):
    last_modified_file = ''
    for filename in os.listdir(path1):
        filename = str(path1) + "\\" + str(filename)
        filename = Path(filename)
        n = 0
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
    print(file)
    send2trash.send2trash(Revenue_Data + "\\" + file)

# Move four files to specific folder
for path3 in LocationOfFile:
    shutil.copy(path3, Revenue_Data)


# Open our working file by using password
wb1 = excel.Workbooks.Open(Working_File, False, False, None, password, password)
excel.Visible = True
# TODO: Need to find filename
# Run Excel Macro in Working file
excel.Application.Run(".xlsm!Module1.ClearData")

"""

# Dictionary to search for column in our working file 
Month = {'January': '4', 'February': '5', 'March': '6', 'April': '7', 'May': '8', 'June': '9',
         'July': '10', 'August': '11', 'September': '12', 'October': '13', 'November': '14', 'December': '15'}


# Find the correct column in the revenue report to copy
def FindLastCell(ws):
    RowRange = ws.Range("A1:A200")
    # No need to start from 1 as it need to exclude the Business pick-up row
    for i, value in enumerate(RowRange):
        if "Business pick-up" in str(value):
            EndOfRow =  i
            break
    return EndOfRow
        
# Open Revenue Report for copying data to our working file
# Seperate into two conditions as HIMCC and CMCC are in the same file
for path4 in os.listdir(Revenue_Data):
    if "SCC" in path4:
        wb2 = excel.Workbooks.Open(Revenue_Data + "\\" + path4)
        # Select the correct worksheet for HIMCC and CMCC
        # Conrad Worksheet
        wsCM = wb2.Worksheets('Report - Conrad')
        LastCellNumber = FindLastCell(wsCM)
        
        
        # Holiday Inn Worksheet
        wsHI = wb2.Worksheets('Report - Holiday Inn')
        LastCellNumber = FindLastCell(wsHI)
        
        wb2.Close(True)
    else:
        # Venetian and Parisian Worksheets
        wb2 = excel.Workbooks.Open(Revenue_Data + "\\" + path4)
        ws2 = wb2.Worksheets('Report')
        LastCellNumber = FindLastCell(ws2)
        
        
        wb2.Close(True)

#excel.Application.Quit()
# TODO: Copy data in revenue data and paste it in our report
    # TODO: Find the correct column (month) in our report to paste (Using Dictionary)