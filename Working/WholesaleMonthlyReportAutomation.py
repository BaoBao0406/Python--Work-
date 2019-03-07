#! pyhon3
# WholesaleMonthlyReportAutomation.py - find the latest file, copy the correct worksheet to the
# working file and click the Marco button

import os.path, shutil, PathNPassword
from pathlib import Path
import win32com.client as win32
from win32com.client import constants
excel = win32.gencache.EnsureDispatch('Excel.Application')


# Import the password and working file path from PathNPassword
Working_File_Path = PathNPassword.Working_File_Path
Working_File = PathNPassword.Working_File
password = PathNPassword.password

# Create an empty Wholesale Date excel file
Wholesale_Data_Path = PathNPassword.Wholesale_Data_Path
Wholesale_Data = PathNPassword.Wholesale_Data
wb_WSData = excel.Workbooks.Add()
wb_WSData.SaveAs(Wholesale_Data_Path + Wholesale_Data)

# Import the Revenue Path according to month
Revenue_Path = PathNPassword.Revenue_Path

#for filename in os.listdir(Revenue_Path):
#    if filename.endswith('.xlsx') or filename.endswith('.xlsm'):
        

# TODO: Open each of the excel file for all properties
    # TODO: Turn off the Calculate Now
    # TODO: Copy and paste the value
    # TODO: Change the tab name according to the property
wb_WSData.Close(True)

# Open working file with password
wb1 = excel.Workbooks.Open(Working_File, False, False, None, password, password)
excel.Visible = True

# TODO: Save as a new file

# TODO: Click the macro button

excel.Application.Quit()