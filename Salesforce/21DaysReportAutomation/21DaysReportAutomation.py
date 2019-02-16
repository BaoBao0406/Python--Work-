#! python3
# 21DaysReportAutomation.py - run data from salesforce, get data and export to excel. Send the excel
# with distribution list.

from simple_salesforce import Salesforce, SalesforceLogin
import requests, password, datetime
import stripJunkSimpleSalesforce as stripForce
import pandas as pd

Username = password.username
Password = password.password
SecurityToken = password.securitytoken

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
for s, n in zip(Status, TabName) :
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
            
# TODO: Copy the Worksheet from (.xlsx) to Working File .(xlsm)
# TODO: Run the excel marco function to format the excel file

# TODO: Send email according to the distribution list
   # TODO: Use the email template as the body

# TODO: Send follow up email to manager who need to follow up on the booking
   # TODO: Use the email template as the body