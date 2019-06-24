#! python3
# Audit_Events_T_and_P_with_Past_Arrival_Date.py - run data from salesforce for Events contain Event with Past Arrival Date
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
FilePath = 'I:\\10-Sales\\Delphi\\Audit_Info\\Audits\\MPE\\' + str(year) + '\\Event\\Event Past Arrival Date P&T\\'

# Salesforce Login info
Username = password.Username
Password = password.Password
securitytoken = password.securitytoken

# Change proxies and ports
proxies = {"http": "http://10.96.250.10:80", "https":"https://10.96.250.10:80"}

# Login to Salesforce
sf = Salesforce(instance='na1.salesforce.com', session_id='', proxies=proxies, username = Username, password= Password, security_token = securitytoken)
session = requests.Session()

# Date Range is last 120 days to now
now = datetime.datetime.now()
StartDate = now - datetime.timedelta(days=120)
Current_Date = str(now.year) + '-' + str('%02d'% now.month) + '-' + str('%02d'% now.day)
Start_Date = str(StartDate.year) + '-' + str('%02d'% StartDate.month) + '-' + str('%02d'% StartDate.day)
File_Date = str(now.year) + str('%02d'% now.month) + str('%02d'% now.day)

# Run SOQL to get data
BKdata1 = sf.query("SELECT Owner.Name, Owner.Email, nihrm__Property__c, nihrm__Booking__r.Name, nihrm__Booking__r.nihrm__BookingStatus__c, nihrm__FunctionRoom__r.Name, Name, nihrm__StartDate__c, \
                           nihrm__EndDate__c, nihrm__EventStatus__c, nihrm__Booking__r.nihrm__ArrivalDate__c, CreatedBy.Name \
                    FROM nihrm__BookingEvent__c \
                    WHERE (nihrm__StartDate__c <= TODAY AND nihrm__StartDate__c >= " + str(Start_Date) + ") AND (nihrm__EventStatus__c IN ('Prospect', 'Tentative'))")
                    
# Convert the data to a readable format
BKdata2 = stripJunkSimpleSalesforce(BKdata1)
# Sorting the order for the columns
index = ['Owner.Name', 'nihrm__Property__c', 'nihrm__Booking__r.Name', 'nihrm__Booking__r.nihrm__BookingStatus__c', 'nihrm__FunctionRoom__r.Name', 'Name', 'nihrm__StartDate__c',
         'nihrm__EndDate__c', 'nihrm__EventStatus__c', 'nihrm__Booking__r.nihrm__ArrivalDate__c', 'CreatedBy.Name', 'Owner.Email']
BKdata3 = pd.DataFrame((pd.DataFrame.from_dict(BKdata2)), columns = index)
# Add Owner Email to EmailList and then take out from the column
EmailList = list(set(BKdata3['Owner.Email'].tolist()))
del BKdata3['Owner.Email']
# Change column header
BKdata3.columns = ['Event Owner', 'Property', 'Post As', 'Booking Status', 'Function Room', 'Event Name', 'Start Date', 'End Date', 'Event Status', 'Booking Arrival', 'CreatedBy']
# Transfer data into excel file
ExcelPath = FilePath + File_Date + '_Event Past Arrival Date P&T1' + '.xlsx'
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
    mail.Subject = 'Audit: Events (T and P) with Past Arrival Date'
    mail.Attachments.Add(ExcelPath)
    # Add image for Message Body
    attachment = mail.Attachments.Add("I:\\10-Sales\\Delphi\\Audit_Info\\Audits\\MPE\\Python Code\\Others\\AuditEventsPastArrival.jpg", 0x5, 0)
    imageCid = "AuditEventsPastArrival.jpg@123"
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imageCid)
    # Add Signature to Email first
    mail.GetInspector
    
    # Message Body + Image Add
    MessageBody = "<p>Dear&nbsp;all,</p> <p>We have found some events that were created inside Definite bookings, but the status of each event has not been updated.</p><p>Please find the attachment with the list of <strong>Prospect and Tentative events</strong> where the <strong><em><u>Event Arrival Date</u></em></strong> had passed already.</p> \
                  <p>Kindly don&rsquo;t forget to update the <u>event status</u> with either <strong><u>Event Cancelled</u></strong> (if the event did not take place) OR <strong><u>Definite</u></strong> (if the event did happen)</p>" + "<img src=\"cid:{0}\">".format(imageCid)
    # Find and replace to add Message Body to HTML text
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body')) 
    mail.HTMLbody = mail.HTMLbody[:index + 1] + MessageBody + mail.HTMLbody[index + 1:]
    mail.SaveAs(Path='I:\\10-Sales\\Personal Folder\\Admin & Assistant Team\\Patrick Leong\\Python Code\\Audit Report\\Email.msg')
    #mail.send
    
# Send email to booking owner if EmailList is larger than 0
if len(EmailList) > 0:
    SendEmail()
