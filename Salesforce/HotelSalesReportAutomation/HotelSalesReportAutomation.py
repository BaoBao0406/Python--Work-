#! python3
# HotelSalesReport.py - run data from salesforce, get data and export to excel


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
Start_Date = str(now.year) + '-' + str('01') + '-' + str('01')
End_Date = str(now.year) + '-' + str('12') + '-' + str('31')
FileDate = str(now.year) + str('%02d'% now.month) + str('%02d'% now.day)

# Login to Salesforce
sf = Salesforce(instance='na1.salesforce.com', session_id='', proxies=proxies, username = Username, password= Password, security_token = securitytoken)
session = requests.Session()

# Run SOQL to get Hotel Sales Report data
BKdata1 = sf.query("SELECT nihrm__Property__c, End_User_Region__c, nihrm__BookingTypeName__c, VCL_Booking Team_N, Owner.Name, RSO_Manager__r.Name, nihrm__Account__r.Name, Name, nihrm__BookedDate__c, nihrm__ArrivalDate__c, nihrm__DepartureDate__c, \
                           nihrm__CurrentBlendedRoomnightsTotal__c, 	nihrm__BlendedGuestroomRevenueTotal__c, nihrm__CurrentBlendedADR__c, VCL_Blended_F_B_Revenue__c, 	nihrm__CurrentBlendedEventRevenue7__c, nihrm__CurrentBlendedEventRevenue3__c, nihrm__LastStatusDate__c, nihrm__BookingStatus__c\
                    FROM nihrm__Booking__c \
                    WHERE (NOT nihrm__BookingTypeName__c IN ('ALT Alternative', 'Default', 'IN Internal', 'CN Concert')) AND (nihrm__BookingStatus__c IN ('Definite', 'Tentative', 'Prospect', 'Cancelled', 'TurnedDown')) AND (NOT nihrm__Property__c IN ('Sands Macao Hotel')) AND \
                    (NOT nihrm__StatusReasonName__c IN ('Alternative booking', 'Operator Error in Entry')) AND (nihrm__BookedDate__c >= " + str(Start_Date) + " AND nihrm__BookedDate__c <= " + str(End_Date) + ")")

# Convert the data to a readable format
BKdata2 = stripForce.stripJunkSimpleSalesforce(BKdata1)
#print(BKdata2)
index = ['nihrm__Property__c', 'End_User_Region__c', 'nihrm__BookingTypeName__c', 'VCL_Booking Team_N', 'Owner.Name', 'RSO_Manager__r', 'nihrm__Account__r.Name', 'Name', 'nihrm__BookedDate__c', 'nihrm__ArrivalDate__c', 'nihrm__DepartureDate__c',
         'nihrm__CurrentBlendedRoomnightsTotal__c', 'nihrm__BlendedGuestroomRevenueTotal__c', 'nihrm__CurrentBlendedADR__c', 'VCL_Blended_F_B_Revenue__c', 'nihrm__CurrentBlendedEventRevenue7__c', 'nihrm__CurrentBlendedEventRevenue3__c', 'nihrm__LastStatusDate__c', 'nihrm__BookingStatus__c']
BKdata3 = pd.DataFrame(pd.DataFrame.from_dict(BKdata2), columns = index)
BKdata3.columns = ['Property', 'End User Region', 'BkgType', 'Group', 'BkdByID', 'BookedBy', 'RSO', 'Account Name', 'PostAs', 'CreateDate', 'ArrivalDate', 'BookedDate', 'DepartureDate', 'Room Night', ' Room Night Rev', 'AverageRate',
                   'F&B Revenue', 'Venue', 'Others', 'StatusChangeDate', 'Status']
#BKdata3[]
print(BKdata3)

