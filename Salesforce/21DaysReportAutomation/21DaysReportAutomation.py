#! python3
# 21DaysReportAutomation.py - run data from salesforce, get data and export to excel. Send the excel
# with distribution list.

from simple_salesforce import Salesforce, SalesforceLogin
import requests, password, datetime
import stripJunkSimpleSalesforce as stripForce
import pandas as pd
import win32com.client as win32
from time import sleep
excel = win32.gencache.EnsureDispatch('Excel.Application')
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header

path = 'D:\\Python\\Additional\\Salesforce\\21DaysReport\\'

# Salesforce Login info
Username = password.username
Password = password.password
SecurityToken = password.securitytoken

# Email Login info
Email = password.Email
Epassword = password.Epassword
DistributionList_To = password.DistributionList_To
DistributionList_Cc = password.DistributionList_Cc
EmailBody = password.EmailBody

# Login to Salesforce
session_id, instance = SalesforceLogin(username= Username, password= Password, security_token= SecurityToken)
sf= Salesforce(instance=instance, session_id=session_id)
session = requests.Session()

# Get Date for 21 days later
now = datetime.datetime.now()
EndDate = now + datetime.timedelta(days=21)
Current_Date = str(now.year) + '-' + str('%02d'% now.month) + '-' + str('%02d'% now.day)
End_Date = str(EndDate.year) + '-' + str('%02d'% EndDate.month) + '-' + str('%02d'% EndDate.day)

Status = ['Prospect', 'Tentative']
TabName = ['PROS', 'TENT']

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
             'nihrm__BookingTypeName__c', 'nihrm__ForecastRoomnightsTotal__c', 'nihrm__DecisionDate__c', 'nihrm__BookedDate__c', 'Owner.Email']
    BKdata3 = pd.DataFrame.from_dict(BKdata2)
    BKdata4 = pd.DataFrame(BKdata3, columns = index)
    # Add to email distribution list
    if n == "PROS":
        ProsEmailList = BKdata4['Owner.Email'].tolist()
    elif n == "TENT":
        TentEmailList = BKdata4['Owner.Email'].tolist()
    
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
    RBdata3 = pd.DataFrame.from_dict(RBdata2)
    RBdata4 = pd.DataFrame(RBdata3, columns = index)
    
    # Transfer the data to excel file
    writer = pd.ExcelWriter(str(s) + '.xlsx', engine='xlsxwriter')
    BKdata4.to_excel(writer, index=False, sheet_name=n, startrow=2)
    RBdata4.to_excel(writer, index=False, sheet_name="RN Block by Property", startrow=2)

sleep(10)
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
    
# Send email according to the distribution list
smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
smtpObj.ehlo()
smtpObj.starttls()
smtpObj.login(Email, Epassword)

# TODO: Function for P and T status
def Email_Sent_PROSnTENT(Status, Email, DistributionList_To, DistributionList_Cc, EmailBody, FilePath):
    fromaddr = Email
    toaddr = DistributionList_To
    cc = DistributionList_Cc

    message = MIMEMultipart()
    message.attach(MIMEText(EmailBody, 'html', 'utf-8'))
    message['From'] = fromaddr
    message['To'] = ','.join(toaddr)
    message['Cc'] = ','.join(cc)

    Subject = '21 Days ' + str(Status) + ' Report'
    message['Subject'] = Header(Subject, 'utf-8')

    att1 = MIMEText(open(FilePath, 'rb').read(), 'base64', 'utf-8')
    att1["Content-Type"] = 'application/octet-steam'
    att1["Content-Disposition"] = 'attachment; filename="21Days' + str(Status) + '.xlsm"'
    message.attach(att1)

    try:
        smtpObj.sendmail(message['From'], [message['To'], message['Cc']], message.as_string())
        print('Email sent')
    except smtplib.SMTPException:
        print("Failed")

# PROS report email sent
# TODO: Check to see if the P and T file exist.
    
# TENT report email sent
# TODO: Check to see if the P and T file exist.        

# TODO: Send follow up email to manager who need to follow up on the booking
   # TODO: Use the email template as the body