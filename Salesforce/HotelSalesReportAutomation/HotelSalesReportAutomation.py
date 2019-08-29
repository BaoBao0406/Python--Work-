#! python3
# HotelSalesReport.py - run data from salesforce, get data and export to excel

from simple_salesforce import Salesforce
import requests, datetime, os.path, password
import stripJunkSimpleSalesforce as stripForce
import pandas as pd
from time import sleep
import numpy as np
#import win32com.client as win32
#excel = win32.gencache.EnsureDispatch('Excel.Application')

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
BKdata1 = sf.query("SELECT nihrm__Property__c, End_User_Region__c, nihrm__BookingTypeName__c, VCL_Booking_Team_N__c, Owner.Name, RSO_Manager__r.Name, nihrm__Account__r.Name, Name, nihrm__BookedDate__c, nihrm__ArrivalDate__c, nihrm__DepartureDate__c, \
                           nihrm__CurrentBlendedRoomnightsTotal__c, 	nihrm__BlendedGuestroomRevenueTotal__c, nihrm__CurrentBlendedADR__c, VCL_Blended_F_B_Revenue__c, nihrm__CurrentBlendedEventRevenue7__c, nihrm__CurrentBlendedEventRevenue3__c, nihrm__LastStatusDate__c, nihrm__BookingStatus__c, nihrm__StatusReasonName__c \
                    FROM nihrm__Booking__c \
                    WHERE (NOT nihrm__BookingTypeName__c IN ('ALT Alternative', 'Default', 'IN Internal', 'CN Concert')) AND (nihrm__BookingStatus__c IN ('Definite', 'Tentative', 'Prospect', 'Cancelled', 'TurnedDown')) AND (NOT nihrm__Property__c IN ('Sands Macao Hotel')) AND \
                    (NOT nihrm__StatusReasonName__c IN ('Alternative booking', 'Operator Error in Entry')) AND (nihrm__BookedDate__c >= " + str(Start_Date) + " AND nihrm__BookedDate__c <= " + str(End_Date) + ")")

# Convert the data to a readable format
BKdata2 = stripForce.stripJunkSimpleSalesforce(BKdata1)

index = ['nihrm__Property__c', 'End_User_Region__c', 'nihrm__BookingTypeName__c', 'VCL_Booking_Team_N__c', 'Owner.Name', 'RSO_Manager__r.Name', 'nihrm__Account__r.Name', 'Name', 'nihrm__BookedDate__c', 'nihrm__ArrivalDate__c', 'nihrm__DepartureDate__c',
         'nihrm__CurrentBlendedRoomnightsTotal__c', 'nihrm__BlendedGuestroomRevenueTotal__c', 'nihrm__CurrentBlendedADR__c', 'VCL_Blended_F_B_Revenue__c', 'nihrm__CurrentBlendedEventRevenue7__c', 'nihrm__CurrentBlendedEventRevenue3__c', 'nihrm__LastStatusDate__c', 'nihrm__BookingStatus__c', 'nihrm__StatusReasonName__c']
BKdata3 = pd.DataFrame(pd.DataFrame.from_dict(BKdata2), columns = index)

# Date columns in List for process
DateFormat = ['nihrm__BookedDate__c', 'nihrm__ArrivalDate__c', 'nihrm__DepartureDate__c', 'nihrm__LastStatusDate__c']

# Convert column format to datetime
for date in DateFormat:
    BKdata3[str(date)] = pd.to_datetime(BKdata3[str(date)])

# Add 'Group' column - Show SM for Sales Manager or RSM for RSO
BKdata3['Group'] = BKdata3['RSO_Manager__r.Name'].apply(lambda x: 'SM' if pd.isnull(x) else 'RSM')
# Add 'Period' column - Display the Year and month
BKdata3['Period'] = pd.to_datetime(BKdata3['nihrm__BookedDate__c']).dt.strftime('%Y-%m')
# Add 'Total' column - Use (Room Night * AverageRate) to calculate the revenue
BKdata3['Total'] = BKdata3['nihrm__CurrentBlendedRoomnightsTotal__c'] * BKdata3['nihrm__CurrentBlendedADR__c']
# Add 'Lead Time' column - Add and calculate the difference between Arrival to Booked Date
BKdata3['Lead Time'] = BKdata3['nihrm__ArrivalDate__c'] - BKdata3['nihrm__BookedDate__c']
BKdata3['Lead Time'] = (BKdata3['Lead Time'] / np.timedelta64(1, 'D')) + 1
# Add 'Count_Status' column - Display '1' if Status in Definite, and '0' for other status
BKdata3['Count_Status'] = BKdata3['nihrm__BookingStatus__c'].apply(lambda x: 1 if x == 'Definite' else 0)
# Add 'Lead_Status' column - same value as 'Count_Status'
BKdata3['Lead_Status'] = BKdata3['Count_Status']
# Add 'Lead_all' column - 
BKdata3['Lead_all'] = BKdata3['nihrm__BookingStatus__c'].apply(lambda x: 0 if x == '' else 1)

#BKdata3['nihrm__BookedDate__c'] = pd.to_datetime(BKdata3['nihrm__BookedDate__c']).dt.strftime('%d/%m/%Y')

# Columne final order
#BKdata3.columns = ['Property', 'End User Region', 'BkgType', 'Group', 'BkdByID', 'BookedBy', 'RSO', 'Account Name', 'PostAs', 'CreateDate', 'ArrivalDate', 'BookedDate', 'DepartureDate', 'Room Night', ' Room Night Rev', 'AverageRate',
#                   'F&B Revenue', 'Venue', 'Others', 'StatusChangeDate', 'Status', 'Period', 'Total', 'Lead Time', 'Count_Status', 'Lead_Status', 'Lead_all']
