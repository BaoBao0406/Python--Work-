#! python3
# Audit_RSO_Bkg_without_Inquiry.py - run data from salesforce for RSO booking without Inquiry
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
FilePath = 'I:\\10-Sales\\Delphi\\Audit_Info\\Audits\\MPE\\' + str(year) + '\\Booking\\RSO Bkg without Inquiry\\'

# Salesforce Login info
Username = password.Username
Password = password.Password
securitytoken = password.securitytoken

# Change proxies and ports
proxies = {"http": "http://10.96.250.10:80", "https":"https://10.96.250.10:80"}

# Login to Salesforce
sf = Salesforce(instance='na1.salesforce.com', session_id='', proxies=proxies, username = Username, password= Password, security_token = securitytoken)
session = requests.Session()

# Date Range is last 14 days to now
now = datetime.datetime.now()
StartDate = now - datetime.timedelta(days=14)
Current_Date = str(now.year) + '-' + str('%02d'% now.month) + '-' + str('%02d'% now.day)
Start_Date = str(StartDate.year) + '-' + str('%02d'% StartDate.month) + '-' + str('%02d'% StartDate.day)
File_Date = str(now.year) + str('%02d'% now.month) + str('%02d'% now.day)

# Run SOQL to get data
BKdata1 = sf.query("SELECT Owner.Name, Owner.Email, nihrm__Property__c, nihrm__Account__r.Name, Name, nihrm__Inquiry__r.Name, RSO_Manager__r.Name, nihrm__BookingStatus__c, nihrm__BookedDate__c \
                    FROM nihrm__Booking__c \
                    WHERE (nihrm__BookedDate__c <= TODAY AND nihrm__BookedDate__c >= " + str(Start_Date) + ") AND (nihrm__BookingStatus__c IN ('Prospect', 'Tentative', 'Definite')) AND (nihrm__BookingTypeName__c NOT IN ('ALT Alternative', 'IN Internal')) \
                           AND (nihrm__Inquiry__r.Name = null) AND (RSO_Manager__r.Name != null)")
                    #WHERE (nihrm__BookedDate__c <= TODAY AND nihrm__BookedDate__c >= " + str(Start_Date) + ")")
                    
# Convert the data to a readable format
BKdata2 = stripJunkSimpleSalesforce(BKdata1)
# Sorting the order for the columns
index = ['Owner.Name', 'nihrm__Property__c', 'nihrm__Account__r.Name', 'Name', 'nihrm__Inquiry__r.Name', 'RSO_Manager__r.Name', 'nihrm__BookingStatus__c', 'Owner.Email']
BKdata3 = pd.DataFrame((pd.DataFrame.from_dict(BKdata2)), columns = index)
# Add Owner Email to EmailList and then take out from the column
EmailList = list(set(BKdata3['Owner.Email'].tolist()))
del BKdata3['Owner.Email']
# Change column header
BKdata3.columns = ['Owner', 'Property', 'Account', 'Post As', 'Inquiry Name', 'RSO Manager', 'Status']
# Transfer data into excel file
ExcelPath = FilePath + File_Date + '_audit_rso_bkg_without_inquiry' + '.xlsx'
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
    mail.Subject = 'Audit: RSO Bkg without Inquiry'
    mail.Attachments.Add(ExcelPath)
    # Add Signature to Email first
    mail.GetInspector
    
    # Message Body + Image Add
    MessageBody = "<p>Dear,</p><p>Please note that the below booking(s) that have RSO Manager involved are without the <u>Primary Inquiry attachment</u>.</p><p>Don&rsquo;t forget that all bookings come from RSO need to have Primary Inquiry (sent by RSO).</p>"
    # Find and replace to add Message Body to HTML text
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body')) 
    mail.HTMLbody = mail.HTMLbody[:index + 1] + MessageBody + mail.HTMLbody[index + 1:]
    mail.SaveAs(Path='I:\\10-Sales\\Delphi\\Audit_Info\\Audits\\MPE\\Python Code\\Email.msg')
    #mail.send
    
# Send email to booking owner if EmailList is larger than 0
if len(EmailList) > 0:
    SendEmail()
