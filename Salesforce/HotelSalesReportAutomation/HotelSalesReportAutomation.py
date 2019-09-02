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
BKdata1 = sf.query("SELECT nihrm__Property__c, End_User_Region__c, nihrm__BookingTypeName__c, VCL_Booking_Team_M__c, Owner.Name, RSO_Manager__r.Name, nihrm__Account__r.Name, Name, nihrm__BookedDate__c, nihrm__ArrivalDate__c, nihrm__DepartureDate__c, \
                           nihrm__CurrentBlendedRoomnightsTotal__c, 	nihrm__BlendedGuestroomRevenueTotal__c, nihrm__CurrentBlendedADR__c, VCL_Blended_F_B_Revenue__c, nihrm__CurrentBlendedEventRevenue7__c, nihrm__CurrentBlendedEventRevenue3__c, nihrm__LastStatusDate__c, nihrm__BookingStatus__c, nihrm__StatusReasonName__c \
                    FROM nihrm__Booking__c \
                    WHERE (NOT nihrm__BookingTypeName__c IN ('ALT Alternative', 'Default', 'IN Internal', 'CN Concert')) AND (nihrm__BookingStatus__c IN ('Definite', 'Tentative', 'Prospect', 'Cancelled', 'TurnedDown')) AND (NOT nihrm__Property__c IN ('Sands Macao Hotel')) AND \
                    (NOT nihrm__StatusReasonName__c IN ('Alternative booking', 'Operator Error in Entry')) AND (nihrm__BookedDate__c >= " + str(Start_Date) + " AND nihrm__BookedDate__c <= " + str(End_Date) + ")")

# Convert the data to a readable format
BKdata2 = stripForce.stripJunkSimpleSalesforce(BKdata1)
BKdata3 = pd.DataFrame.from_dict(BKdata2)

# Rename column name
BKdata3.rename(columns={'nihrm__BookedDate__c': 'CreateDate', 'nihrm__ArrivalDate__c': 'ArrivalDate', 
                        'nihrm__DepartureDate__c': 'DepartureDate', 'nihrm__LastStatusDate__c': 'StatusChangeDate'}, inplace=True)
# Date columns in List for process
DateFormat = ['CreateDate', 'ArrivalDate', 'DepartureDate', 'StatusChangeDate']
# Convert column format to datetime
for date in DateFormat:
    BKdata3[str(date)] = pd.to_datetime(BKdata3[str(date)])

# Create columns to fit Excel report
# Add 'BookedDate' column - same date as CreatedDate
BKdata3['BookedDate'] = BKdata3['CreateDate']
# Add 'Group' column - Show SM for Sales Manager or RSM for RSO
BKdata3['Group'] = BKdata3['RSO_Manager__r.Name'].apply(lambda x: 'SM' if pd.isnull(x) else 'RSM')
# Add 'Period' column - Display the Year and month
BKdata3['Period'] = pd.to_datetime(BKdata3['CreateDate']).dt.strftime('%Y-%m')
# Add 'Total' column - Use (Room Night * AverageRate) to calculate the revenue
BKdata3['Total'] = BKdata3['nihrm__CurrentBlendedRoomnightsTotal__c'] * BKdata3['nihrm__CurrentBlendedADR__c']
# Add 'Lead Time' column - Add and calculate the difference between Arrival to Booked Date
BKdata3['Lead Time'] = BKdata3['ArrivalDate'] - BKdata3['CreateDate']
BKdata3['Lead Time'] = (BKdata3['Lead Time'] / np.timedelta64(1, 'D')) + 1
# Add 'Count_Status' column - Display '1' if Status in Definite, and '0' for other status
BKdata3['Count_Status'] = BKdata3['nihrm__BookingStatus__c'].apply(lambda x: 1 if x == 'Definite' else 0)
# Add 'Lead_Status' column - same value as 'Count_Status'
BKdata3['Lead_Status'] = BKdata3['Count_Status']
# Add 'Lead_all' column - 
BKdata3['Lead_all'] = BKdata3['nihrm__BookingStatus__c'].apply(lambda x: 0 if x == '' else 1)
# Replace NaT value with "None"
BKdata3['StatusChangeDate'] = BKdata3['StatusChangeDate'].apply(lambda x: None if x=="NaT" else x)

# Sort Column Order
index = ['nihrm__Property__c', 'End_User_Region__c', 'nihrm__BookingTypeName__c', 'Group', 'VCL_Booking_Team_M__c', 'Owner.Name', 'RSO_Manager__r.Name', 'nihrm__Account__r.Name', 'Name', 'CreateDate', 'ArrivalDate', 'BookedDate', 'DepartureDate', 'nihrm__CurrentBlendedRoomnightsTotal__c', 'nihrm__BlendedGuestroomRevenueTotal__c', 
         'nihrm__CurrentBlendedADR__c', 'VCL_Blended_F_B_Revenue__c', 'nihrm__CurrentBlendedEventRevenue7__c', 'nihrm__CurrentBlendedEventRevenue3__c', 'StatusChangeDate', 'nihrm__BookingStatus__c', 'nihrm__StatusReasonName__c', 'Period', 'Total', 'Lead Time', 'Count_Status', 'Lead_Status', 'Lead_all']
BKdata3 = pd.DataFrame(BKdata3, columns = index)
# Drop 'nihrm__StatusReasonName__c' column
BKdata3.drop('nihrm__StatusReasonName__c', axis=1, inplace=True)

# Re-arrange columns order
BKdata3.columns = ['Property', 'End User Region', 'BkgType', 'Group', 'BkdByID', 'BookedBy', 'RSO', 'Account Name', 'PostAs', 'CreateDate', 'ArrivalDate', 'BookedDate', 'DepartureDate', 'Room Night', ' Room Night Rev', 'AverageRate',
                   'F&B Revenue', 'Venue', 'Others', 'StatusChangeDate', 'Status', 'Period', 'Total', 'Lead Time', 'Count_Status', 'Lead_Status', 'Lead_all']

# Convert date column format to suit excel format
for date in DateFormat:
    BKdata3[str(date)] = BKdata3[str(date)].dt.strftime('%d/%m/%Y')

"""
# TODO: take out Sheraton RN and Revenue from the data
SheratonData1 =  = sf.query("SELECT nihrm__Location__r.Name, nihrm__Booking__r.Name, nihrm__ForecastRoomnightsTotal__c, nihrm__ForecastRevenueTotal__c, nihrm__BookedDate__c, Name, nihrm__Booking__r.nihrm__BookingTypeName__c, nihrm__RoomBlockStatus__c \
                             FROM nihrm__BookingRoomBlock__c \
                             WHERE nihrm__Booking__r.nihrm__BookingTypeName__c NOT IN ('Default', 'IN Internal', 'CN Concert', 'ALT Alternative') AND (nihrm__RoomBlockStatus__c = '" + str(s) + "') AND (nihrm__BookedDate__c >= " + str(Start_Date) + "  AND nihrm__BookedDate__c <= " + str(End_Date) + ") AND (nihrm__Location__r.Name IN ('FSHM', 'SGMH', 'TSRM'))")

# Convert the data to a readable format
SheratonData2 = stripForce.stripJunkSimpleSalesforce(SheratonData1)
SheratonData3 = pd.DataFrame.from_dict(SheratonData2)
# Rename Sheraton column name
SheratonData3.columns = ['Property', 'Post As', 'Room Night', ' Room Night Rev', 'CreateDate', 'RoomBlockName','BookingType', 'Status']

# Sumif function for Roomnights in Room Block tab into Booking tab
x = 0
RN = RBdata3.groupby(['Post As'])['Roomnights'].sum()
while RN.count()-1 >= x:
    for i, j in BKdata3.iterrows():    
        if RN.index[x] == j['Post As']:
            BKdata3.at[i, 'Roomnights'] = RN.values[x]
    x += 1
"""
# All lead Top3 - Create Groupby for Top3 data
Region1Top3 = BKdata3[['End User Region', 'Room Night', 'Lead_all']].groupby('End User Region').sum()
# Total_Demand - Top 3 Region for 'Total Demand'
Total_Demand = Region1Top3.sort_values('Room Night', ascending=False).head(3)
# Total_Leads - Top 3 Region for 'Total Leads'
Total_Leads = Region1Top3.sort_values('Lead_all', ascending=False).head(3)

# Def lead Top3 - Create Groupby for Top3 Definite data
Region2Top3 = BKdata3[BKdata3['Count_Status'] == 1]
Region2Top3 = Region2Top3[['End User Region', 'Room Night', 'Count_Status']].groupby('End User Region').sum()
# Def_RN - Top 3 Region for 'Definite RNs'
Def_RN = Region2Top3.sort_values(by='Room Night', ascending=False).head(3)
# Def_Leads - Top 3 Region for 'No of Definite'
Def_Leads = Region2Top3.sort_values(by='Count_Status', ascending=False).head(3)
