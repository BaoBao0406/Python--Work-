#! python3
# 21DaysReportAutomation.py - run data from salesforce, get data and export to excel. Send the excel
# with distribution list.

from simple_salesforce import Salesforce
import requests, password, datetime, os.path
import stripJunkSimpleSalesforce as stripForce
import pandas as pd
import win32com.client as win32
from time import sleep
excel = win32.gencache.EnsureDispatch('Excel.Application')
outlook = win32.Dispatch("Outlook.Application")

path = 'I:\\10-Sales\\01_Sales_Reports\\21 Days Report\\'

# Salesforce Login info
Username = password.Username
Password = password.Password
securitytoken = password.securitytoken

# Change proxies and ports
proxies = {"http": "http://10.96.250.10:80", "https":"https://10.96.250.10:80"}

# Login to Salesforce
sf = Salesforce(instance='na1.salesforce.com', session_id='', proxies=proxies, username = Username, password= Password, security_token = securitytoken)
session = requests.Session()

# Get Date for 21 days later
now = datetime.datetime.now()
EndDate = now + datetime.timedelta(days=21)
Current_Date = str(now.year) + '-' + str('%02d'% now.month) + '-' + str('%02d'% now.day)
End_Date = str(EndDate.year) + '-' + str('%02d'% EndDate.month) + '-' + str('%02d'% EndDate.day)

Status = ['Prospect', 'Tentative']
TabName = ['PROS', 'TENT']

# Adjust the column width (Credit To Stackoverflow: TrigonaMinima)
def AdjustColumnWidth(Sheet, Name):
    worksheet = writer.sheets[Sheet]
    for i, col in enumerate(Name.columns):
        column_len = Name[col].astype(str).str.len().max()
        column_len = max(column_len, len(col)) + 2
        worksheet.set_column(i, i, column_len)

# Run for both Prospect and Tentative status for Booking tab
for s, n in zip(Status, TabName):
    # Booking tab - Use SOQL languauges to export the Booking tab from Salesforce
    BKdata1 = sf.query("SELECT Owner.Name, Owner.Email, nihrm__Account__r.name, nihrm__Agency__r.name, nihrm__Property__c, Name, nihrm__ArrivalDate__c, nihrm__DepartureDate__c, nihrm__BookingTypeName__c, nihrm__ForecastRoomnightsTotal__c, nihrm__DecisionDate__c, nihrm__BookedDate__c \
                       FROM nihrm__Booking__c \
                       WHERE (nihrm__BookingStatus__c = '" + str(s) + "') AND (nihrm__BookingTypeName__c NOT IN ('Default', 'IN Internal', 'CN Concert')) AND (nihrm__ArrivalDate__c >= TODAY AND nihrm__ArrivalDate__c <= " + str(End_Date) + ") AND (nihrm__Property__c NOT IN ('Sands Macao Hotel'))")
    # Convert the data to a readable format
    BKdata2 = stripForce.stripJunkSimpleSalesforce(BKdata1)
    # Sorting the order for the columns
    index = ['Owner.Name', 'nihrm__Property__c', 'nihrm__Account__r.Name', 'nihrm__Agency__r.Name', 'Name', 'nihrm__ArrivalDate__c', 'nihrm__DepartureDate__c', \
              'nihrm__ForecastRoomnightsTotal__c', 'nihrm__DecisionDate__c', 'nihrm__BookedDate__c', 'nihrm__BookingTypeName__c', 'Owner.Email']
    BKdata3 = pd.DataFrame(pd.DataFrame.from_dict(BKdata2), columns = index)
    # Add to email distribution list
    if n == "PROS":
        ProsEmailList = BKdata3['Owner.Email'].tolist()
    elif n == "TENT":
        TentEmailList = BKdata3['Owner.Email'].tolist()
    del BKdata3['Owner.Email']
    # Change column header
    BKdata3.columns = ['Booking Owner', 'Property', 'Account', 'Agency', 'Post As', 'Arrival', 'Departure', 'Roomnights', 'Decision Due', 'Booked Date', 'Booking Type']
    # Transfer the data to excel file
    writer = pd.ExcelWriter(str(s) + '.xlsx', engine='xlsxwriter')
    BKdata3.to_excel(writer, index=False, sheet_name=n, startrow=2)
    # Adjust Column width
    AdjustColumnWidth(n, BKdata3)
    
    # Property (Location) Code to exclude in the report
    ExcludeProp = "('FSHM', 'SGMH', 'SANDS', 'TSRM')"
    # Room Block tab - Use SOQL languauges to export the Room Block tab from Salesforce
    RBdata1 = sf.query("SELECT Owner.Name, nihrm__Location__r.Name, nihrm__StartDate__c, Name, nihrm__Booking__r.nihrm__BookingTypeName__c, nihrm__RoomBlockStatus__c, nihrm__Booking__r.Name, nihrm__ForecastRoomnightsTotal__c \
                       FROM nihrm__BookingRoomBlock__c \
                       WHERE nihrm__Booking__r.nihrm__BookingTypeName__c NOT IN ('Default', 'IN Internal', 'CN Concert') AND (nihrm__RoomBlockStatus__c = '" + str(s) + "') AND (nihrm__StartDate__c >= TODAY AND nihrm__StartDate__c <= " + str(End_Date) + ") AND (nihrm__Location__r.Name NOT IN " + str(ExcludeProp) + ")")
    # Convert the data to a readable format
    RBdata2 = stripForce.stripJunkSimpleSalesforce(RBdata1)
    # Sorting the order for the columns
    index = ['nihrm__Location__r.Name', 'nihrm__Booking__r.Name', 'Name', 'Owner.Name', 'nihrm__Booking__r.nihrm__BookingTypeName__c', 'nihrm__RoomBlockStatus__c', 'nihrm__StartDate__c', 'nihrm__ForecastRoomnightsTotal__c']
    RBdata3 = pd.DataFrame(pd.DataFrame.from_dict(RBdata2), columns = index)
    RBdata3.columns = ['Property', 'Post As', 'Room Block Name', 'Booking Owner', 'Booking Type', 'Status', 'Start Date', 'Roomnights']
    # Transfer the data to excel file
    RBdata3.to_excel(writer, index=False, sheet_name="RN Block by Property", startrow=2)
    # Adjust Column width
    AdjustColumnWidth("RN Block by Property", RBdata3)
    writer.save()
    sleep(10)
"""
for s, n in zip(Status, TabName):    
    try:
        # Copy the Worksheet from (.xlsx) to Working File .(xlsm)
        wb1 = excel.Workbooks.Open(path + s + '.xlsx')
    except Exception as e:
        continue
    else:
        wb2 = excel.Workbooks.Open(path + '21Days' + s + '.xlsm')
        wsBK = wb1.Worksheets(n)
        wsBK.Copy(Before=wb2.Worksheets('Sheet1'))
        wsRB = wb1.Worksheets('RN Block by Property')
        wsRB.Copy(Before=wb2.Worksheets('Sheet1'))
        # Run the excel marco function to format the excel file
        #excel.Application.Run('21Days' + s + "!Module1.Step1")
        wb2.Close(SaveChanges=True)
        wb1.Close()
        excel.Quit()
"""
# Function to Create Prospect Email to send to DOS and SM
def EmailPROS(ToList, CCList, FollowUp):
    mail.To = ';'.join(ToList)
    mail.CC = ';'.join(CCList)
    mail.Subject = '21 days report Prospect'
    mail.Attachments.Add(path + 'Prospect.xlsx')
    # Add Signature to Email first
    mail.GetInspector
    # Message Body
    if FollowUp == False:
        MessageBody = "<p>Dear All</p><p><br /> Please refer to the attached 21 days PROSPECT business report. This report is run on a daily basis and sent to all DOS to follow up with their sales team.</p> \
                       <p>Note that there are two tabs <strong>&ldquo;PROS&rdquo; </strong>shows the main property of each group with total Room Nights and <strong>&ldquo;Room Block by Property&rdquo;</strong> in the report.</p> \
                       <p> Any question or problem, please feel free to contact the Systems Team.</p>"
        Filename = 'PEmail.msg'
    elif FollowUp == True:
        MessageBody = "<p>Dear All</p> <p><br /> Please refer to the attached 21 days PROSPECT business report.&nbsp;Please follow up on your booking(s) with expired decision due dates</p> \
                       <p>Note that there are two tabs <strong>&ldquo;PROS&rdquo; </strong>shows the main property of each group with total Room Nights and <strong>&ldquo;Room Block by Property&rdquo;</strong> in the report.</p> <p>&nbsp;</p>"
        # Set email to High Importance
        mail.Importance = 2
        Filename = 'PEmailFollowUp.msg'
    # Find and replace to add Message Body to HTML text
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body')) 
    mail.HTMLbody = mail.HTMLbody[:index + 1] + MessageBody + mail.HTMLbody[index + 1:]
    mail.SaveAs(Path='I:\\10-Sales\\01_Sales_Reports\\21 Days Report\\Python Code\\Testing\\' + Filename)
    mail.Close(0)

# Function to Create Tentative Email to send to DOS and SM
def EmailTENT(ToList, CCList, FollowUp):
    mail.To = ';'.join(ToList)
    mail.CC = ';'.join(CCList)
    mail.Subject = '21 days report Tentative'
    mail.Attachments.Add(path + 'Tentative.xlsx')
    # Add Signature to Email first
    mail.GetInspector
    # Message Body
    if FollowUp == False:
        MessageBody = "<p>Dear All<br /> <br /> Please refer to the attached 21 days short term business report. This report is run on a daily basis and sent to all DOS follow up and to other appropriate departments for reference. <br />\
                       <br /> The DOS is responsible for ensuring the sales managers confirm the groups by the decision due dates listed.<br /> <br /> Any contract received for groups on this list must be routed short term, to all the appropriate departments, by the sales assistant, ASAP. \
                       <br /> <br /> Any assistant who have problem with short term routing, please feel free to contact the Systems Team.</p> <p>&nbsp;</p> <p>Please note: These are all the Tentative groups that are arriving within the next 21 days.</p> \
                       <p>There are two tabs &ldquo;<strong>Tentative Booking</strong>&rdquo; shows the main property of each group with total Room Nights and <strong>&ldquo;Room Block by Property&rdquo;</strong> shows total Room Nights by property per each group in the report.</p>"
    elif FollowUp == True:
        MessageBody = "<p>Dear All</p> <p><br /> Please refer to the attached 21 days PROSPECT business report.&nbsp;Please follow up on your booking(s) with expired decision due dates</p> \
                       <p>Note that there are two tabs <strong>&ldquo;PROS&rdquo; </strong>shows the main property of each group with total Room Nights and <strong>&ldquo;Room Block by Property&rdquo;</strong> in the report.</p> <p>&nbsp;</p>"
    # Find and replace to add Message Body to HTML text
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body')) 
    mail.HTMLbody = mail.HTMLbody[:index + 1] + MessageBody + mail.HTMLbody[index + 1:]
    mail.SaveAs(Path='I:\\10-Sales\\01_Sales_Reports\\21 Days Report\\Python Code\\Testing\\TEmail.msg')
    mail.Close(0)

# Send two Emails to DOS and SM if Prospect report is larger than 0
if len(ProsEmailList) > 0:
    mail = outlook.CreateItem(0)
    EmailPROS(password.ToListforPROS, password.CCListforPROS, False)
    mail = outlook.CreateItem(0)
    EmailPROS(ProsEmailList, password.CCListforPROS, True)
# Send two Emails to DOS and SM if Tentative report is larger than 0
if len(TentEmailList) > 0:
    mail = outlook.CreateItem(0)
    EmailTENT(password.ToListforTENT, password.CCListforTENT, False)
    mail = outlook.CreateItem(0)
    EmailTENT(TentEmailList, password.CCListforPROS, True)
