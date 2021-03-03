#! python3
# autorun_salesforce_data.py - run data from salesforce, get data and export to excel. 

from simple_salesforce import Salesforce
import requests, datetime, os.path, password
import stripJunkSimpleSalesforce as stripForce
import pandas as pd
from time import sleep

#path = 'I:\\10-Sales\\01_Sales_Reports\\21 Days Report\\'

# Salesforce Login info
Username = password.Username
Password = password.Password
securitytoken = password.securitytoken

# Change proxies and ports
proxies = {"https":"https://10.96.250.11:443"}

# Login to Salesforce
sf = Salesforce(instance='na1.salesforce.com', session_id='', proxies=proxies, username = Username, password= Password, security_token = securitytoken)
session = requests.Session()


# Date Range
now = datetime.datetime.now()
EndDate = now + datetime.timedelta(days=21)
Current_Date = str(now.year) + '-' + str('%02d'% now.month) + '-' + str('%02d'% now.day)
End_Date = str(EndDate.year) + '-' + str('%02d'% EndDate.month) + '-' + str('%02d'% EndDate.day)


# TODO: 
def convert_to_excel(data, filename):
    
    

# Account Information
def Account():
    filename = ''
    data = sf.query("SELECT Id, Owner.Name, Name, Type, nihrm__RegionName__c, Industry, BillingCountry, BillingState, BillingCity, Rating, nihrm__MarketSegmentName__c, LastModifiedDate, LastActivityDate, CreatedDate \
                     FROM Account")
    data = stripForce.stripJunkSimpleSalesforce(data)
    convert_to_excel(data, filename)

# Account and Activity query
def Account_Activity():
    data = sf.query("SELECT AccountId, Account.Name, Account.Type, Account.Owner.Name, Account.nihrm__RegionName__c, Account.Industry, Id, Owner.Name, Type, Subject, CreatedDate, ActivityDate \
                     FROM Task \
                     WHERE (ActivityDate >= TODAY AND ActivityDate <= " + str(End_Date) + ")")
    data = stripForce.stripJunkSimpleSalesforce(data)
    convert_to_excel(data, filename)

# TODO: rawdata booking (Outlet and Ancillary field number)
def Booking_information():
    data = sf.query("SELECT Owner.Name, nihrm__Property__c, nihrm__Account__r.name, nihrm__Account__r.BillingCity, nihrm__Account__r.BillingCountry, nihrm__Agency__r.name, Name, nihrm__ArrivalDate__c, nihrm__DepartureDate__c, Percentage_of_Attrition__c, 	nihrm__CommissionPercentage__c, Promotion__c, nihrm__AtDefiniteAgreedRoomnights__c, nihrm__CurrentBlendedRoomnightsTotal__c, \
                            nihrm__AtDefiniteAgreedGuestroomRevenue__c, nihrm__BlendedGuestroomRevenueTotal__c, nihrm__AtDefiniteBlendedEventRevenue1__c, nihrm__CurrentBlendedEventRevenue1__c, nihrm__AtDefiniteBlendedEventRevenue2__c, nihrm__CurrentBlendedEventRevenue2__c, VCL_Blended_F_B_Revenue__c, nihrm__AtDefiniteBlendedEventRevenue7__c, nihrm__CurrentBlendedEventRevenue7__c,\
                            nihrm__AtDefiniteBlendedEventRevenue4__c, nihrm__CurrentBlendedEventRevenue4__c, nihrm__AtDefiniteBlendedEventRevenue3__c, nihrm__CurrentBlendedEventRevenue3__c, nihrm__AtDefiniteBlendedEventRevenue6__c, nihrm__CurrentBlendedEventRevenue6__c, Sheraton_F_B_Revenue__c, Sheraton_Room_Rental_Revenue__c, nihrm__BookingStatus__c, nihrm__LastStatusDate__c, \
                            nihrm__BookedDate__c, End_User_Region__c, End_User_SIC__c, nihrm__BookingTypeName__c, nihrm__LostToCompetitorName__c, nihrm__StatusReasonName__c, Id \
                     FROM nihrm__Booking__c \
                     WHERE (nihrm__BookingStatus__c IN ('Prospect', 'Tentative', 'Definite', 'TurnedDown', 'Cancelled')) AND (nihrm__BookingTypeName__c NOT IN ('Default', 'IN Internal', 'CN Concert')) AND (nihrm__ArrivalDate__c >= TODAY AND nihrm__ArrivalDate__c <= " + str(End_Date) + ") AND (nihrm__Property__c NOT IN ('Sands Macao Hotel'))")
    data = stripForce.stripJunkSimpleSalesforce(data)
    convert_to_excel(data, filename)

# TODO: Booking and Activity
def Booking_Activity():
    

# TODO: Roomnight peak number and revenue (query not finished)
def Roomnight_peak():
    data = sf.query("SELECT nihrm__Booking__c, nihrm__Location__r.Name, nihrm__StartDate__c, nihrm__Booking__r.Name, nihrm__AgreedRoomnightsTotal__c, nihrm__PickupRoomnightsTotal__c \
                     FROM nihrm__BookingRoomBlock__c")
#    data = sf.query("SELECT Id (SELECT nihrm__RoomBlock__r.nihrm__PatternDate__c FROM nihrm__BookingRoomNight__r) \
#                 FROM nihrm__BookingRoomBlock__c")
    data = stripForce.stripJunkSimpleSalesforce(data)
    convert_to_excel(data, filename)
    
# Event max attendance
def Event_max_attendance():
    data = sf.query("SELECT nihrm__Booking__c, nihrm__AgreedAttendance__c \
                     FROM nihrm__BookingEvent__c")
    data = stripForce.stripJunkSimpleSalesforce(data)
    convert_to_excel(data, filename)


data = sf.query("SELECT Owner.Name, nihrm__Property__c, nihrm__Account__r.name, nihrm__Account__r.BillingCity, nihrm__Account__r.BillingCountry, nihrm__Agency__r.name, Name, nihrm__ArrivalDate__c, nihrm__DepartureDate__c, Percentage_of_Attrition__c, 	nihrm__CommissionPercentage__c, Promotion__c, nihrm__AtDefiniteAgreedRoomnights__c, nihrm__CurrentBlendedRoomnightsTotal__c, \
                        nihrm__AtDefiniteAgreedGuestroomRevenue__c, nihrm__BlendedGuestroomRevenueTotal__c, nihrm__AtDefiniteBlendedEventRevenue1__c, nihrm__CurrentBlendedEventRevenue1__c, nihrm__AtDefiniteBlendedEventRevenue2__c, nihrm__CurrentBlendedEventRevenue2__c, VCL_Blended_F_B_Revenue__c, nihrm__AtDefiniteBlendedEventRevenue7__c, nihrm__CurrentBlendedEventRevenue7__c,\
                        nihrm__AtDefiniteBlendedEventRevenue4__c, nihrm__CurrentBlendedEventRevenue4__c, nihrm__AtDefiniteBlendedEventRevenue3__c, nihrm__CurrentBlendedEventRevenue3__c, nihrm__AtDefiniteBlendedEventRevenue6__c, nihrm__CurrentBlendedEventRevenue6__c, Sheraton_F_B_Revenue__c, Sheraton_Room_Rental_Revenue__c, nihrm__BookingStatus__c, nihrm__LastStatusDate__c, \
                        nihrm__BookedDate__c, End_User_Region__c, End_User_SIC__c, nihrm__BookingTypeName__c, nihrm__LostToCompetitorName__c, nihrm__StatusReasonName__c, Id \
                 FROM nihrm__Booking__c \
                 WHERE (nihrm__BookingStatus__c IN ('Prospect', 'Tentative', 'Definite', 'TurnedDown', 'Cancelled')) AND (nihrm__BookingTypeName__c NOT IN ('Default', 'IN Internal', 'CN Concert')) AND (nihrm__ArrivalDate__c >= TODAY AND nihrm__ArrivalDate__c <= " + str(End_Date) + ") AND (nihrm__Property__c NOT IN ('Sands Macao Hotel'))")
data = stripForce.stripJunkSimpleSalesforce(data)
print(data)


if __name__ == '__main__':
    try:
        # run query
        Account()
        sleep(10)
        Account_Activity()
        sleep(10)
        Booking_information()
        sleep(10)
        Booking_Activity()
        sleep(10)
        Roomnight_peak()
        sleep(10)
        Event_max_attendance()
        
    except Exception as err:
        # display error message
        print('Error Reason: %s' % err)
