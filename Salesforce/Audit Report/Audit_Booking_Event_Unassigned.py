#! python3
# Audit_Booking_Event_Unassigned.py - run data from salesforce for Events with Function Space Unassigned
# and send email to booking owner

from simple_salesforce import Salesforce
import requests, datetime
from Others.stripJunkSimpleSalesforce import stripJunkSimpleSalesforce
from Others import password
import pandas as pd
import win32com.client as win32
outlook = win32.Dispatch("Outlook.Application")


# Excel SaveAs file path
year = datetime.datetime.today().year
FilePath = 'I:\\10-Sales\\Delphi\\Audit_Info\\Audits\\MPE\\' + str(year) + '\\Event\\Event Unassigned\\'

# Salesforce Login info
Username = password.Username
Password = password.Password
securitytoken = password.securitytoken

# Change proxies and ports
proxies = {"http": "http://10.96.250.10:80", "https":"https://10.96.250.10:80"}

# Login to Salesforce
sf = Salesforce(instance='na1.salesforce.com', session_id='', proxies=proxies, username = Username, password= Password, security_token = securitytoken)
session = requests.Session()

# Report Run date
now = datetime.datetime.now()
File_Date = str(now.year) + str('%02d'% now.month) + str('%02d'% now.day)

# Run SOQL to get data
BKdata1 = sf.query("SELECT Owner.Name, Owner.Email, nihrm__Property__c, nihrm__Booking__r.nihrm__Account__r.Name, nihrm__Booking__r.Name, nihrm__StartDate__c,  Name, nihrm__FunctionRoom__r.Name, \
                           nihrm__EventStatus__c, CreatedBy.Name \
                    FROM nihrm__BookingEvent__c \
                    WHERE (nihrm__FunctionRoom__r.Name IN ('Unassigned')) AND (nihrm__EventStatus__c IN ('Definite', 'Tentative')) AND (NOT nihrm__Booking__r.Name like '%Testing%')")
                    
# Convert the data to a readable format
BKdata2 = stripJunkSimpleSalesforce(BKdata1)
# Sorting the order for the columns
index = ['Owner.Name', 'nihrm__Property__c', 'nihrm__Booking__r.nihrm__Account__r.Name', 'nihrm__Booking__r.Name', 'nihrm__StartDate__c', 'Name', 'nihrm__FunctionRoom__r.Name',
         'nihrm__EventStatus__c', 'CreatedBy.Name', 'Owner.Email']
BKdata3 = pd.DataFrame((pd.DataFrame.from_dict(BKdata2)), columns = index)
# Add Owner Email to EmailList and then take out from the column
EmailList = list(set(BKdata3['Owner.Email'].tolist()))
del BKdata3['Owner.Email']
# Change column header
BKdata3.columns = ['Event Owner', 'Property', 'Account', 'Post As', 'Start Date', 'Event Name', 'Function Room', 'Event Status', 'CreatedBy']
# Transfer data into excel file
ExcelPath = FilePath + File_Date + '_audit_booking_event_unassigned' + '.xlsx'
writer = pd.ExcelWriter(ExcelPath, engine='xlsxwriter')
BKdata3.to_excel(writer, index=False, sheet_name='Report')

# Adjust the column width (Copied from Stackoverflow: TrigonaMinima)
worksheet = writer.sheets['Report']
for i, col in enumerate(BKdata3.columns):
    column_len = BKdata3[col].astype(str).str.len().max()
    column_len = max(column_len, len(col)) + 2
    worksheet.set_column(i, i, column_len)
writer.save()

# Send via Outlook
def SendEmail():
    mail = outlook.CreateItem(0)
    mail.To = ';'.join(EmailList)
    mail.CC = ';'.join(password.CCList)
    mail.Subject = 'Audit : Booking Event Unassigned'
    mail.Attachments.Add(ExcelPath)
    # Set email to High Importance
    mail.Importance = 2
    # Add Signature to Email first
    mail.GetInspector
    
    # Message Body + Image Add
    MessageBody = "<p>Dear All</p><p>&nbsp;</p><p>Please find the attachment of your booking(s)/event(s) and don&rsquo;t forget to update the function room space that is/are still under <strong>Unassigned</strong>.</p>"
    # Find and replace to add Message Body to HTML text
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body')) 
    mail.HTMLbody = mail.HTMLbody[:index + 1] + MessageBody + mail.HTMLbody[index + 1:]
    mail.SaveAs(Path='I:\\10-Sales\\Delphi\\Audit_Info\\Audits\\MPE\\Python Code\\Email.msg')
    #mail.send
    
# Send email to booking owner if EmailList is larger than 0
if len(EmailList) > 0:
    SendEmail()
