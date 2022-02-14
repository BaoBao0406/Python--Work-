#! python3
# 21DaysReportAutomation.py - run data from sql, get data and export to excel. Send the excel
# with distribution list.

import datetime, os.path, mail_distribution_list, pyodbc
import pandas as pd
import win32com.client as win32
from time import sleep
excel = win32.gencache.EnsureDispatch('Excel.Application')
outlook = win32.Dispatch("Outlook.Application")

##################################################################################################

# filepath save
path = 'I:\\10-Sales\\+Operational Reports (5Y, Restricted)\\2021\\21 Days Report\\'

# Get Date for 21 days later
now = datetime.datetime.now()
EndDate = now + datetime.timedelta(days=21)
Current_Date = str(now.year) + '-' + str('%02d'% now.month) + '-' + str('%02d'% now.day)
End_Date = str(EndDate.year) + '-' + str('%02d'% EndDate.month) + '-' + str('%02d'% EndDate.day)
FileDate = str(now.year) + str('%02d'% now.month) + str('%02d'% now.day)

# connect sql server
conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=VOPPSCLDBN01\VOPPSCLDBI01;'
                      'Database=SalesForce;'
                      'Trusted_Connection=yes;')


# FDC User ID, Name, Email list
user = pd.read_sql('SELECT DISTINCT(Id), Name, Email \
                    FROM dbo.[User]', conn)
user_list = user.set_index('Id')['Name'].to_dict()
user_email = user.set_index('Id')['Email'].to_dict()

##################################################################################################

# Adjust the column width (Credit To Stackoverflow: TrigonaMinima)
def AdjustColumnWidth(writer, Sheet, Name):
    worksheet = writer.sheets[Sheet]
    for i, col in enumerate(Name.columns):
        column_len = Name[col].astype(str).str.len().max()
        column_len = max(column_len, len(col)) + 2
        worksheet.set_column(i, i, column_len)

# Sumif function for Roomnights in Room Block tab into Booking tab
def sumif_function(BKdata, RBdata):
    x = 0
    RN = RBdata.groupby(['Post As'])['Roomnights'].sum()
    while RN.count()-1 >= x:
        for i, j in BKdata.iterrows():    
            if RN.index[x] == j['Post As']:
                BKdata.at[i, 'Roomnights'] = RN.values[x]
        x += 1
    return BKdata

# Booking SQL query
def booking_sql_data(s):
    BKdata = pd.read_sql("SELECT BK.OwnerId, BK.nihrm__Property__c, ac.Name, ag.Name,  BK.Name, FORMAT(BK.nihrm__ArrivalDate__c, 'MM/dd/yyyy') AS ArrivalDate, FORMAT(BK.nihrm__DepartureDate__c, 'MM/dd/yyyy') AS DepartureDate, \
                                 BK.nihrm__ForecastRoomnightsTotal__c, FORMAT(BK.nihrm__DecisionDate__c, 'MM/dd/yyyy') AS LastStatusDate, FORMAT(BK.nihrm__BookedDate__c, 'MM/dd/yyyy') AS BookedDate, BK.nihrm__BookingTypeName__c \
                          FROM dbo.nihrm__Booking__c AS BK \
                          LEFT JOIN dbo.Account AS ac \
                              ON BK.nihrm__Account__c = ac.Id \
                          LEFT JOIN dbo.Account AS ag \
                              ON BK.nihrm__Agency__c = ag.Id \
                          WHERE (BK.nihrm__BookingStatus__c = '" + str(s) + "') AND (BK.nihrm__BookingTypeName__c NOT IN ('Default', 'IN Internal', 'CN Concert')) AND \
                                (BK.nihrm__ArrivalDate__c BETWEEN CONVERT(datetime, '" + Current_Date + "') AND CONVERT(datetime, '" + End_Date + "')) AND (BK.nihrm__Property__c NOT IN ('Sands Macao Hotel'))", conn)
    BKdata.columns = ['Booking Owner', 'Property', 'Account', 'Agency', 'Post As', 'Arrival', 'Departure', 'Roomnights', 'Decision Due', 'Booked Date', 'Booking Type']
    # create mail list
    mail_list = BKdata[['Booking Owner']]
    mail_list['Booking Owner'].replace(user_email, inplace=True)
    # replace id with Owner Name
    BKdata['Booking Owner'].replace(user_list, inplace=True)

    return BKdata, mail_list
    
# Room Block SQL query
def room_block_sql_data(s):
    RBdata = pd.read_sql("SELECT prop.Name, BK.Name, RoomB.Name, RoomB.OwnerId, BK.nihrm__BookingTypeName__c, RoomB.nihrm__RoomBlockStatus__c, FORMAT(RoomB.nihrm__StartDate__c, 'MM/dd/yyyy') AS StartDate, \
                                 RoomB.nihrm__ForecastRoomnightsTotal__c \
                          FROM dbo.nihrm__BookingRoomBlock__c AS RoomB \
                          LEFT JOIN dbo.nihrm__Booking__c AS BK \
                              ON RoomB.nihrm__Booking__c = BK.Id \
                          LEFT JOIN dbo.nihrm__Location__c AS prop \
                              ON RoomB.nihrm__Location__c = prop.Id \
                          WHERE (BK.nihrm__BookingTypeName__c NOT IN ('Default', 'IN Internal', 'CN Concert')) AND (RoomB.nihrm__RoomBlockStatus__c = '" + str(s) + "') AND \
                                (prop.Name NOT IN ('Sheraton Grand Macao', 'Four Seasons Hotel Macao Cotai Strip', 'Sands Macao Hotel', 'The St. Regis Macao')) AND \
                                (RoomB.nihrm__StartDate__c BETWEEN CONVERT(datetime, '" + Current_Date + "') AND CONVERT(datetime, '" + End_Date + "'))", conn)
    RBdata.columns = ['Property', 'Post As', 'Room Block Name', 'Booking Owner', 'Booking Type', 'Status', 'Start Date', 'Roomnights']
    RBdata['Booking Owner'].replace(user_list, inplace=True)
    
    return RBdata
    

# function to create P and T booking and room block data, then save as excel file
def P_and_T_booking():
    
    Status = ['Prospect', 'Tentative']
    TabName = ['PROS', 'TENT']
    
    for s, n in zip(Status, TabName):
        # run Booking data
        BKdata, mail_list = booking_sql_data(s)

        # run Room Block data
        RBdata = room_block_sql_data(s)
        
        # run sumif function
        BKdata = sumif_function(BKdata, RBdata)
        
        # Transfer the data to excel file
        writer = pd.ExcelWriter(path + FileDate + '_21 days report ' + str(s) + '.xlsx', engine='xlsxwriter')
        BKdata.to_excel(writer, index=False, sheet_name=n)
        # Adjust Column width
        AdjustColumnWidth(writer, n, BKdata)          
        
        # Transfer the data to excel file
        RBdata.to_excel(writer, index=False, sheet_name="RN Block by Property")
        # Adjust Column width
        AdjustColumnWidth(writer, "RN Block by Property", RBdata)
        writer.save()
        sleep(10)
        
        # create email distribution list for P and T
        if n == "PROS":
            ProsEmailList = list(set(mail_list['Booking Owner'].tolist()))
        elif n == "TENT":
            TentEmailList = list(set(mail_list['Booking Owner'].tolist()))
        
    return ProsEmailList, TentEmailList


# Function to Create Prospect Email to send to DOS and SM
def EmailPROS(ToList, CCList, SMList):
    mail.To = ';'.join(ToList)
    mail.CC = ';'.join(CCList) + ';' + ';'.join(SMList)
    mail.Subject = '21 days report Prospect'
    mail.Attachments.Add(path + FileDate + '_21 days report Prospect.xlsx')
    # Add Signature to Email first
    mail.GetInspector
    # Message Body
    MessageBody = "<p>Dear All</p> <p>Please refer to the attached 21 days PROSPECT business report. This report is run on a daily basis and sent to all DOS to follow up with their sales team.</p> \
                   <p>For Sales TeamManager - Please follow up on your booking(s) with expired decision due dates.</p><p>Note that there are two tabs <strong><strong>&ldquo;PROS&rdquo; </strong> \
                   </strong>shows the main property of each group with total Room Nights and <strong><strong>&ldquo;Room Block by Property&rdquo;</strong></strong> in the report.</p> \
                   <p>Any question or problem, please feel free to contact Systems Team.</p>"
    # Set email to High Importance
    #mail.Importance = 2
    # Find and replace to add Message Body to HTML text
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body')) 
    mail.HTMLbody = mail.HTMLbody[:index + 1] + MessageBody + mail.HTMLbody[index + 1:]
    #mail.SaveAs(Path='I:\\10-Sales\\01_Sales_Reports\\21 Days Report\\Python Code\\Testing\\Prospect.msg')
    mail.send

# Function to Create Tentative Email to send to DOS and SM
def EmailTENT(ToList, CCList, SMList):
    mail.To = ';'.join(ToList)
    mail.CC = ';'.join(CCList) + ';' + ';'.join(SMList)
    mail.Subject = '21 days report Tentative'
    mail.Attachments.Add(path + FileDate + '_21 days report Tentative.xlsx')
    # Add Signature to Email first
    mail.GetInspector
    # Message Body
    MessageBody = "<p>Dear All</p> <p>Please refer to the attached Tentative groups that are arriving within the next 21 days. This report is run on a daily basis and sent to all DOS follow up and to other appropriate departments for reference. <br /> \
                   <br /> Any contract received for groups on this list must be routed short term to all the appropriate departments by the sales assistant ASAP. <br /> <br /> Any assistant who has problem with short term routing, please feel free to contact Systems Team.</p> \
                   <p>For Sales Manager - Please follow up on your booking(s) with expired decision due dates.</p> <p>Note that there are two tabs <strong><strong>&ldquo;TENT&rdquo;</strong></strong> shows the main property of each group with total Room Nights and \
                   <strong><strong>&ldquo;Room Block by Property&rdquo; </strong></strong>shows total Room Nights by property per each group in the report.</p> "
        #Filename = 'TEmailFollowUp.msg'
    # Find and replace to add Message Body to HTML text
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body')) 
    mail.HTMLbody = mail.HTMLbody[:index + 1] + MessageBody + mail.HTMLbody[index + 1:]
    #mail.SaveAs(Path='I:\\10-Sales\\01_Sales_Reports\\21 Days Report\\Python Code\\Testing\\Tentative.msg')
    mail.send


# main function
if __name__ == '__main__':
    # run 
    ProsEmailList, TentEmailList = P_and_T_booking()

    # Send two Emails to DOS and SM if Prospect report is larger than 0
    if len(ProsEmailList) > 0:
        mail = outlook.CreateItem(0)
        EmailPROS(mail_distribution_list.ToListforPROS, mail_distribution_list.CCListforPROS, ProsEmailList)

    # Send two Emails to DOS and SM if Tentative report is larger than 0
    if len(TentEmailList) > 0:
        mail = outlook.CreateItem(0)
        EmailTENT(mail_distribution_list.ToListforTENT, mail_distribution_list.CCListforTENT, TentEmailList)
