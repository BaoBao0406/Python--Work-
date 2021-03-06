#! python3
# Audit_Booking_with_Past_Arrival_Date_P_and_T.py - run data from salesforce for Past Arrival Date with P and
# T status and send email to booking owner

from simple_salesforce import Salesforce
import requests, datetime
from Others.stripJunkSimpleSalesforce import stripJunkSimpleSalesforce
from Others import password
import pandas as pd
import win32com.client as win32
outlook = win32.Dispatch("Outlook.Application")


# Excel SaveAs file path
year = datetime.datetime.today().year
FilePath = 'I:\\10-Sales\\Delphi\\Audit_Info\\Audits\\MPE\\' + str(year) + '\\Booking\\Booking (T and P) with Past Arrival Date\\'

# Salesforce Login info
Username = password.Username
Password = password.Password
securitytoken = password.securitytoken

# Change proxies and ports
proxies = {"http": "http://10.96.250.10:80", "https":"https://10.96.250.10:80"}

# Login to Salesforce
sf = Salesforce(instance='na1.salesforce.com', session_id='', proxies=proxies, username = Username, password= Password, security_token = securitytoken)
session = requests.Session()

# Date Range is from last 90 days till now
now = datetime.datetime.now()
StartDate = now - datetime.timedelta(days=90)
Current_Date = str(now.year) + '-' + str('%02d'% now.month) + '-' + str('%02d'% now.day)
Start_Date = str(StartDate.year) + '-' + str('%02d'% StartDate.month) + '-' + str('%02d'% StartDate.day)
File_Date = str(now.year) + str('%02d'% now.month) + str('%02d'% now.day)

# Run SOQL to get data
BKdata1 = sf.query("SELECT Owner.Name, Owner.Email, nihrm__Account__r.Name, nihrm__Property__c, Name, nihrm__ArrivalDate__c, nihrm__AgreedGuestroomRevenueTotal__c, nihrm__AgreedRoomnightsTotal__c, nihrm__BookingStatus__c\
                  FROM nihrm__Booking__c \
                  WHERE (nihrm__ArrivalDate__c <= TODAY AND nihrm__ArrivalDate__c >= " + str(Start_Date) + ") AND (nihrm__BookingStatus__c IN ('Prospect', 'Tentative'))")
# Convert the data to a readable format
BKdata2 = stripJunkSimpleSalesforce(BKdata1)
# Sorting the order for the columns
index = ['nihrm__Property__c', 'Owner.Name', 'nihrm__Account__r.Name', 'Name', 'nihrm__AgreedRoomnightsTotal__c', 'nihrm__AgreedGuestroomRevenueTotal__c', 'nihrm__ArrivalDate__c', 'nihrm__BookingStatus__c', 'Owner.Email']
BKdata3 = pd.DataFrame((pd.DataFrame.from_dict(BKdata2)), columns = index)
# Add Owner Email to EmailList and then take out from the column
EmailList = list(set(BKdata3['Owner.Email'].tolist()))
del BKdata3['Owner.Email']
# Change column header
BKdata3.columns = ['Property', 'Booking Owner', 'Account', 'Post As', 'Agreed RNs', 'Agreed RN Revenue', 'Arrival', 'Status']

# Transfer data into excel file
ExcelPath = FilePath + File_Date + '_Audit Bookings (T and P) with Past Arrival Date' + '.xlsx'
writer = pd.ExcelWriter(ExcelPath, engine='xlsxwriter')
BKdata3.to_excel(writer, index=False, sheet_name='Past Arrival Date')

# Adjust the column width (Copied from Stackoverflow: TrigonaMinima)
worksheet = writer.sheets['Past Arrival Date']
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
    mail.Subject = 'Audit: Bookings (T and P) with Past Arrival Date'
    mail.Attachments.Add(ExcelPath)
    # Set email to High Importance
    mail.Importance = 2
    
    # Add Signature to Email first
    mail.GetInspector
    # Message Body
    MessageBody = "<html><body> <p>Dear all,</p><p>Please find the attachment with the list of <strong>Prospect and Tentative bookings</strong> where the <strong><em><u>Arrival Date</u></em>\
                    </strong> had past already.</p><p>Kindly either <u>Turn Down</u> the booking or change the Arrival Date if necessary.</p> <p>Thank you!</p> </body></html>"
    # Find and replace to add Message Body to HTML text
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body')) 
    mail.HTMLbody = mail.HTMLbody[:index + 1] + MessageBody + mail.HTMLbody[index + 1:]
    #mail.SaveAs(Path='I:\\10-Sales\\Personal Folder\\Admin & Assistant Team\\Patrick Leong\\Python Code\\Audit Report\\Email.msg')
    mail.send
    
# Send email to booking owner if EmailList is larger than 0
if len(EmailList) > 0:
    SendEmail()
