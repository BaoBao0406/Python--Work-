#! python3
# 21DaysReportAutomation.py - run data from salesforce, get data and export to excel. Send the excel
# with distribution list.

from simple_salesforce import Salesforce
import requests, datetime, os.path, password
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
FileDate = str(now.year) + str('%02d'% now.month) + str('%02d'% now.day)

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
    # Booking tab
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
        ProsEmailList = list(set(BKdata3['Owner.Email'].tolist()))
    elif n == "TENT":
        TentEmailList = list(set(BKdata3['Owner.Email'].tolist()))
    del BKdata3['Owner.Email']
    # Change column header
    BKdata3.columns = ['Booking Owner', 'Property', 'Account', 'Agency', 'Post As', 'Arrival', 'Departure', 'Roomnights', 'Decision Due', 'Booked Date', 'Booking Type']

    # Room Block tab
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
    
    # Sumif function for Roomnights in Room Block tab into Booking tab
    x = 0
    RN = RBdata3.groupby(['Post As'])['Roomnights'].sum()
    while RN.count()-1 >= x:
        for i, j in BKdata3.iterrows():    
            if RN.index[x] == j['Post As']:
                BKdata3.at[i, 'Roomnights'] = RN.values[x]
        x += 1
    
    # Transfer the data to excel file
    writer = pd.ExcelWriter(path + FileDate + '_21 days report ' + str(s) + '.xlsx', engine='xlsxwriter')
    BKdata3.to_excel(writer, index=False, sheet_name=n)
    # Adjust Column width
    AdjustColumnWidth(n, BKdata3)          
    
    # Transfer the data to excel file
    RBdata3.to_excel(writer, index=False, sheet_name="RN Block by Property")
    # Adjust Column width
    AdjustColumnWidth("RN Block by Property", RBdata3)
    writer.save()
    sleep(10)

# Function to Create Prospect Email to send to DOS and SM
def EmailPROS(ToList, CCList, FollowUp):
    mail.To = ';'.join(ToList)
    mail.CC = ';'.join(CCList)
    mail.Subject = '21 days report Prospect'
    mail.Attachments.Add(path + FileDate + '_21 days report Prospect.xlsx')
    # Add Signature to Email first
    mail.GetInspector
    # Message Body
    if FollowUp == False:
        MessageBody = "<p>Dear All</p><p><br /> Please refer to the attached 21 days PROSPECT business report. This report is run on a daily basis and sent to all DOS to follow up with their sales team.</p> \
                       <p>Note that there are two tabs <strong>&ldquo;PROS&rdquo; </strong>shows the main property of each group with total Room Nights and <strong>&ldquo;Room Block by Property&rdquo;</strong> in the report.</p> \
                       <p> Any question or problem, please feel free to contact the Systems Team.</p>"
        #Filename = 'PEmail.msg'
    elif FollowUp == True:
        MessageBody = "<p>Dear All</p> <p><br /> Please refer to the attached 21 days PROSPECT business report.&nbsp;Please follow up on your booking(s) with expired decision due dates</p> \
                       <p>Note that there are two tabs <strong>&ldquo;PROS&rdquo; </strong>shows the main property of each group with total Room Nights and <strong>&ldquo;Room Block by Property&rdquo;</strong> in the report.</p> <p>&nbsp;</p>"
        # Set email to High Importance
        mail.Importance = 2
        #Filename = 'PEmailFollowUp.msg'
    # Find and replace to add Message Body to HTML text
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body')) 
    mail.HTMLbody = mail.HTMLbody[:index + 1] + MessageBody + mail.HTMLbody[index + 1:]
    #mail.SaveAs(Path='I:\\10-Sales\\01_Sales_Reports\\21 Days Report\\Python Code\\Testing\\' + Filename)
    mail.send


# Function to Create Tentative Email to send to DOS and SM
def EmailTENT(ToList, CCList, FollowUp):
    mail.To = ';'.join(ToList)
    mail.CC = ';'.join(CCList)
    mail.Subject = '21 days report Tentative'
    mail.Attachments.Add(path + FileDate + '_21 days report Tentative.xlsx')
    # Add Signature to Email first
    mail.GetInspector
    # Message Body
    if FollowUp == False:
        MessageBody = "<p>Dear All<br /> <br /> Please refer to the attached 21 days short term business report. This report is run on a daily basis and sent to all DOS follow up and to other appropriate departments for reference. <br />\
                       <br /> The DOS is responsible for ensuring the sales managers confirm the groups by the decision due dates listed.<br /> <br /> Any contract received for groups on this list must be routed short term, to all the appropriate departments, by the sales assistant, ASAP. \
                       <br /> <br /> Any assistant who have problem with short term routing, please feel free to contact the Systems Team.</p> <p>&nbsp;</p> <p>Please note: These are all the Tentative groups that are arriving within the next 21 days.</p> \
                       <p>There are two tabs &ldquo;<strong>Tentative Booking</strong>&rdquo; shows the main property of each group with total Room Nights and <strong>&ldquo;Room Block by Property&rdquo;</strong> shows total Room Nights by property per each group in the report.</p>"
        #Filename = 'TEmail.msg'
    elif FollowUp == True:
        MessageBody = "<p>Dear All</p> <p><br /> Please refer to the attached 21 days TENTATIVE business report.&nbsp;Please follow up on your booking(s) with expired decision due dates</p> \
                       <p>Note that there are two tabs <strong>&ldquo;TENT&rdquo; </strong>shows the main property of each group with total Room Nights and <strong>&ldquo;Room Block by Property&rdquo;</strong> in the report.</p> <p>&nbsp;</p>"
        #Filename = 'TEmailFollowUp.msg'
    # Find and replace to add Message Body to HTML text
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body')) 
    mail.HTMLbody = mail.HTMLbody[:index + 1] + MessageBody + mail.HTMLbody[index + 1:]
    #mail.SaveAs(Path='I:\\10-Sales\\01_Sales_Reports\\21 Days Report\\Python Code\\Testing\\' + Filename)
    mail.send


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
