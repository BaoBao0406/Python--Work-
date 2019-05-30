#! python3
# Booking_Other_Income.py - run data from salesforce for Bookings contain Booking Other Income with Arrival Date within 60 days 
# P and T status and send email to booking owner

from simple_salesforce import Salesforce
import requests, password, datetime
import stripJunkSimpleSalesforce as stripForce
import pandas as pd
import win32com.client as win32
outlook = win32.Dispatch("Outlook.Application")


# Excel SaveAs file path
year = datetime.datetime.today().year
FilePath = 'I:\\10-Sales\\Delphi\\Audit_Info\\Audits\\MPE\\' + str(year) + '\\Booking Other Income\\'

# Salesforce Login info
Username = password.Username
Password = password.Password
securitytoken = password.securitytoken

# Change proxies and ports
proxies = {"http": "http://10.96.250.10:80", "https":"https://10.96.250.10:80"}

# Login to Salesforce
sf = Salesforce(instance='na1.salesforce.com', session_id='', proxies=proxies, username = Username, password= Password, security_token = securitytoken)
session = requests.Session()

# Date Range is from now til 60 days later
now = datetime.datetime.now()
StartDate = now + datetime.timedelta(days=60)
Current_Date = str(now.year) + '-' + str('%02d'% now.month) + '-' + str('%02d'% now.day)
Start_Date = str(StartDate.year) + '-' + str('%02d'% StartDate.month) + '-' + str('%02d'% StartDate.day)
File_Date = str(now.year) + str('%02d'% now.month) + str('%02d'% now.day)
"""
# Run SOQL to get data
BKdata1 = sf.query("SELECT nihrm__Booking__r.Owner.Name, nihrm__Booking__r.Owner.Email, nihrm__Booking__r.nihrm__Account__r.Name,	nihrm__Booking__r.nihrm__Agency__c, nihrm__Booking__r.Name, nihrm__Booking__r.nihrm__BookingStatus__c, nihrm__Booking__r.nihrm__ArrivalDate__c, \
                           nihrm__Booking__r.nihrm__DepartureDate__c, Name, nihrm__OtherIncome__c, nihrm__Description__c \
                    FROM nihrm__BookingOtherIncome__c \
                    WHERE (nihrm__Booking__r.nihrm__ArrivalDate__c <= TODAY AND nihrm__Booking__r.nihrm__ArrivalDate__c >= " + str(Start_Date) + ") AND (nihrm__Booking__r.nihrm__BookingStatus__c IN ('Prospect', 'Tentative'))")
# Convert the data to a readable format
BKdata2 = stripForce.stripJunkSimpleSalesforce(BKdata1)
# Sorting the order for the columns
index = ['nihrm__Property__c', 'Owner.Name', 'nihrm__Account__r.Name', 'Name', 'nihrm__AgreedRoomnightsTotal__c', 'nihrm__AgreedGuestroomRevenueTotal__c', 'nihrm__ArrivalDate__c', 'nihrm__BookingStatus__c', 'Owner.Email']
BKdata3 = pd.DataFrame((pd.DataFrame.from_dict(BKdata2)), columns = index)
# Add Owner Email to EmailList and then take out from the column
EmailList = BKdata3['Owner.Email'].tolist()
del BKdata3['Owner.Email']
# Change column header
BKdata3.columns = ['Property', 'Booking Owner', 'Account', 'Post As', 'Agreed RNs', 'Agreed RN Revenue', 'Arrival', 'Status']
"""
# Transfer data into excel file
ExcelPath = FilePath + File_Date + ' Booking Other income Report' + '.xlsx'
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
    mail.Subject = 'Booking Other Income Report'
    mail.Attachments.Add(ExcelPath)

    
    # Add Signature to Email first
    mail.GetInspector
    # Message Body
    MessageBody = "<p>Dear All</p><p>Attached is the <strong><em>Booking Other Income Report</em></strong> listing all bookings with a marked up room rate (i.e.: includes ferry ticket or Eiffel tower ticket) - for booking arrival within two months. Please don&rsquo;t forget to take action on the listed bookings.</p><p>Once assistants / coordinators send the approved form to the relevant department for issuing the tickets / vouchers, please input comment in the &ldquo;Description&rdquo; field under Booking Other income in Delphi.fdc. This must be done on the same date so that it can reflect in the &ldquo;Description&rdquo; column of the report that the task of the booking has been actioned.</p> \
                  <p>Guideline for inputting comment in Delphi.fdc:</p><p>Go to Booking Other income -&gt; Select &ldquo;Ferry &ndash; Cotai Class &ndash; Adult&rdquo; -&gt; Click &ldquo;Edit&rdquo; -&gt; Input comment in the field of Description</p>"
    # Find and replace to add Message Body to HTML text
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body')) 
    mail.HTMLbody = mail.HTMLbody[:index + 1] + MessageBody + mail.HTMLbody[index + 1:]
    mail.SaveAs(Path='I:\\10-Sales\\Personal Folder\\Admin & Assistant Team\\Patrick Leong\\Python Code\\Audit Report\\Email.msg')
    #mail.send
    
# Send email to booking owner if EmailList is larger than 0
EmailList = list(set(EmailList))
if len(EmailList) > 0:
    SendEmail()
