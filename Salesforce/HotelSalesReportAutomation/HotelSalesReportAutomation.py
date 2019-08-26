#! python3
# HotelSalesReport.py - run data from salesforce, get data and export to excel. Send the excel


from simple_salesforce import Salesforce
import requests, datetime, os.path, password
import stripJunkSimpleSalesforce as stripForce
import pandas as pd
import win32com.client as win32
from time import sleep
excel = win32.gencache.EnsureDispatch('Excel.Application')

# Salesforce Login info
Username = password.Username
Password = password.Password
securitytoken = password.securitytoken

# Change proxies and ports
proxies = {"http": "http://10.96.250.10:80", "https":"https://10.96.250.10:80"}

# Date Range for Current Year
now = datetime.datetime.now()
EndDate = now + datetime.timedelta(days=21)
Start_Date = str(now.year) + '-' + str('01') + '-' + str('01')
End_Date = str(EndDate.year) + '-' + str('12') + '-' + str('31')
FileDate = str(now.year) + str('%02d'% now.month) + str('%02d'% now.day)

# Login to Salesforce
sf = Salesforce(instance='na1.salesforce.com', session_id='', proxies=proxies, username = Username, password= Password, security_token = securitytoken)
session = requests.Session()

# Run SOQL to get Hotel Sales Report data
BKdata1 = sf.query("SELECT nihrm__Property__c, End_User_Region__c, nihrm__BookingTypeName__c, RSO_Manager__r.Name, Owner.Name, nihrm__Account__r.Name, Name, nihrm__BookedDate__c, nihrm__ArrivalDate__c, nihrm__DepartureDate__c, \
                           nihrm__CurrentBlendedRoomnightsTotal__c, 	nihrm__BlendedGuestroomRevenueTotal__c, VCL_Blended_F_B_Revenue__c, 	nihrm__CurrentBlendedEventRevenue7__c, nihrm__CurrentBlendedEventRevenue3__c, nihrm__LastStatusDate__c, nihrm__BookingStatus__c, nihrm__StatusReasonName__c\
                    FROM nihrm__Booking__c \
                    WHERE (NOT nihrm__BookingTypeName__c IN ('ALT Alternative', 'Default', 'IN Internal', 'CN Concert')) AND (nihrm__BookingStatus__c IN ('Definite', 'Tentative', 'Prospect', 'Cancelled', 'TurnedDown')) AND (NOT nihrm__Property__c IN ('Sands Macao Hotel')) AND \
                    (NOT nihrm__StatusReasonName__c IN ('Alternative booking', 'Operator Error in Entry')) AND (nihrm__BookedDate__c >= " + str(Start_Date) + " AND nihrm__BookedDate__c <= " + str(End_Date) + ")")

# Convert the data to a readable format
BKdata2 = stripForce.stripJunkSimpleSalesforce(BKdata1)

BKdata3 = pd.DataFrame(pd.DataFrame.from_dict(BKdata2))
print(BKdata3.head())
