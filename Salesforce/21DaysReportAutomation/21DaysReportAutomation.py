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
EndDate = datetime.datetime.now() + datetime.timedelta(days=20)
year = EndDate.year
month = '%02d'% EndDate.month
day = '%02d'% EndDate.day
End_Date = str(year) + '-' + str(month) + '-' + str(day)

Status = ['Prospect', 'Tentative']


# Run for both Prospect and Tentative status for Booking tab
for s in Status:
    # Use SOQL languauges to export the Booking tab from Salesforce
    data1 = sf.query("SELECT Owner.Name, Owner.Email, nihrm__Account__r.name, nihrm__Agency__r.name, nihrm__Property__c, Name, nihrm__ArrivalDate__c, nihrm__DepartureDate__c, nihrm__BookingTypeName__c, nihrm__ForecastRoomnightsTotal__c, nihrm__DecisionDate__c, nihrm__BookedDate__c FROM nihrm__Booking__c WHERE (nihrm__BookingStatus__c = '" + str(s) + "') AND (nihrm__BookingTypeName__c NOT IN ('Default', 'IN Internal', 'CN Concert')) AND (nihrm__ArrivalDate__c >= TODAY AND nihrm__ArrivalDate__c <= " + str(End_Date) + ") AND (nihrm__Property__c NOT IN ('Sands Macao Hotel'))")

    # Convert the data to a readable format
    data2 = stripForce.stripJunkSimpleSalesforce(data1)
    # Sorting the order for the columns
    index = ['Owner.Name', 'nihrm__Property__c', 'nihrm__Account__r.Name', 'nihrm__Agency__r.Name', 'Name', 'nihrm__ArrivalDate__c', 'nihrm__DepartureDate__c', 'nihrm__BookingTypeName__c', 'nihrm__ForecastRoomnightsTotal__c', 'nihrm__DecisionDate__c', 'nihrm__BookedDate__c', 'Owner.Email']
    data3 = pd.DataFrame.from_dict(data2)
    data4 = pd.DataFrame(data3, columns = index)
    # Transfer the data to excel file
    if s == 'Prospect':
        SheetName = 'PROS'
    else:
        SheetName = 'TENT'
    writer = pd.ExcelWriter(str(s) + '.xlsx', engine ='xlsxwriter')
    data4.to_excel(writer, index=False, sheet_name=SheetName)

    
# Run for both Prospect and Tentative status for Room Block tab
for s in Status:
    # Property (Location) Code to exclude in the report -- need to ask Amademus
    ExcludeProp = "('a0Y28000001RJ07', 'a0Y28000001RJ0A', 'a0Y28000001RJ0B', 'a0Y28000001RJ0C')"
    # Use SOQL languauges to export the Booking tab from Salesforce
    data1 = sf.query("SELECT Owner.Name, nihrm__Location__c, nihrm__StartDate__c, Name, nihrm__Booking__r.nihrm__BookingTypeName__c, nihrm__RoomBlockStatus__c, nihrm__Booking__r.Name, nihrm__ForecastRoomnightsTotal__c FROM nihrm__BookingRoomBlock__c WHERE nihrm__Booking__r.nihrm__BookingTypeName__c NOT IN ('Default', 'IN Internal', 'CN Concert') AND (nihrm__RoomBlockStatus__c = 'Prospect') AND (nihrm__StartDate__c >= TODAY AND nihrm__StartDate__c <= 2019-03-05) AND (nihrm__Location__c NOT IN " + str(ExcludeProp) + ")")
    
    # Convert the data to a readable format
    data2 = stripForce.stripJunkSimpleSalesforce(data1)
    # Sorting the order for the columns
    index = ['nihrm__Location__c', 'nihrm__Booking__r.Name', 'Name', 'Owner.Name', 'nihrm__Booking__r.nihrm__BookingTypeName__c', 'nihrm__RoomBlockStatus__c', 'nihrm__StartDate__c', 'nihrm__ForecastRoomnightsTotal__c']
    data3 = pd.DataFrame.from_dict(data2)
    data4 = pd.DataFrame(data3, columns = index)
    # Transfer the data to excel file
    writer = pd.ExcelWriter(str(s) + '.xlsx', engine ='xlsxwriter')
    data4.to_excel(writer, index=False, sheet_name="RN Block by Property")
    

# TODO: Run the excel marco function to format the excel file

# TODO: Send email according to the distribution list
   # TODO: Use the email template as the body

# TODO: Send follow up email to manager who need to follow up on the booking
   # TODO: Use the email template as the body