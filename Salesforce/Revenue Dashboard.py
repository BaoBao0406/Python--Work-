import pyodbc, datetime
import pandas as pd

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=VOPPSCLDBN01\VOPPSCLDBI01;'
                      'Database=SalesForce;'
                      'Trusted_Connection=yes;')

save_path = ''

# TODO: Date Range
now = datetime.datetime.now()
EndDate = now + datetime.timedelta(days=21)
Current_Date = str(now.year) + '-' + str('%02d'% now.month) + '-' + str('%02d'% now.day)
End_Date = str(EndDate.year) + '-' + str('%02d'% EndDate.month) + '-' + str('%02d'% EndDate.day)

# TODO: Convert data to excel format
def convert_to_excel(data, filename):
    
    return 


# TODO: "01Rawdata_Group Booking by arrival date" - Report
data = pd.read_sql('SELECT * \
                    FROM dbo.nihrm__Booking__c', conn)


# TODO: "02Account and Activities" - Report


# "03Agency and Booking ID" - Report
data = pd.read_sql('SELECT nihrm__Account__c, Booking_ID_Number__c \
                    FROM dbo.nihrm__Booking__c', conn)
data.columns = ['Account: Account ID', 'Booking ID#']
filename = '03Agency and Booking ID'
convert_to_excel(data, filename)


# "04Account and Booking ID" - Report
data = pd.read_sql('SELECT nihrm__Agency__c, Booking_ID_Number__c \
                    FROM dbo.nihrm__Booking__c', conn)
data.columns = ['Agency: Account ID', 'Booking ID#']
filename = '04Account and Booking ID'
convert_to_excel(data, filename)


# TODO: "05Booking and Activities not started" - Report


# "06Event max attendance" - Report
data = pd.read_sql('SELECT dbo.nihrm__Booking__c.Booking_ID_Number__c, nihrm__AgreedAttendance__c \
                    FROM dbo.nihrm__BookingEvent__c \
                    INNER JOIN dbo.nihrm__Booking__c \
                        ON dbo.nihrm__BookingEvent__c.nihrm__Booking__c = dbo.nihrm__Booking__c.Id', conn)
data.columns = ['Booking: Booking ID#', 'Agreed']
filename = '06Event max attendance'
convert_to_excel(data, filename)


# "07roomnight peak data and number" - Report
# TODO: Add date range WHERE clause
data = pd.read_sql("SELECT BK.Booking_ID_Number__c, GS.nihrm__Property__c, GS.Name, BK.Name, FORMAT(RoomN.nihrm__PatternDate__c, 'dd/MM/yyyy') AS PatternDate, \
                           RoomN.nihrm__AgreedRoomsTotal__c, RoomN.nihrm__PickupRoomsTotal__c \
                    FROM dbo.nihrm__BookingRoomNight__c AS RoomN \
                    INNER JOIN dbo.nihrm__Booking__c AS BK \
                        ON RoomN.nihrm__Booking__c = BK.Id \
                    INNER JOIN dbo.nihrm__BookingRoomBlock__c AS RoomB \
                        ON RoomN.nihrm__Booking__c = RoomB.nihrm__Booking__c \
                    INNER JOIN dbo.nihrm__GuestroomType__c AS GS \
                        ON RoomN.nihrm__GuestroomType__c = GS.Id", conn)
data.columns = ['Booking: Booking ID#', 'Property', 'Guestroom Type', 'Booking: Booking Post As', 'Pattern Date', 'Agreed Rooms Total', 'Pickup Rooms Total']
filename = '07roomnight peak data and number'
convert_to_excel(data, filename)


# "08Account Information" - Report
data = pd.read_sql("SELECT ac.Id, owner.Name, ac.Name, ac.Type, ac.nihrm__RegionName__c, ac.Industry, ac.BillingCountry, ac.BillingState, ac.BillingCity, \
                           ac.Rating, ac.nihrm__MarketSegment__c, FORMAT(ac.LastModifiedDate, 'dd/MM/yyyy') AS LastModifiedDate, \
                           FORMAT(ac.LastActivityDate, 'dd/MM/yyyy') AS LastActivityDate, FORMAT(ac.CreatedDate, 'dd/MM/yyyy') AS CreatedDate \
                    FROM dbo.Account AS ac \
                    INNER JOIN dbo.[User] AS owner \
                        ON ac.OwnerId = owner.Id", conn)
data.columns = ['Account ID', 'Account Owner', 'Account Name', 'Type', 'Region', 'Industry', 'Country', 'State/Province', 'City', 'Quality Rating', 
                'Market Segment', 'Last Modified Date', 'Last Activity', 'Created Date']
filename = '08Account Information'
convert_to_excel(data, filename)
