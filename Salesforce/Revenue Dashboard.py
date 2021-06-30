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


# Convert data to excel format
def convert_to_excel(data, filename):
    data.to_excel(save_path + filename + '.xlsx', sheet_name='Sheet1')


# "01Rawdata_Group Booking by arrival date" - Report
data = pd.read_sql("SELECT BK.OwnerId, BK.nihrm__Property__c, ac.Name, ac.BillingCity, ac.BillingCountry, ag.Name, BK.Name, FORMAT(BK.nihrm__ArrivalDate__c, 'dd/MM/yyyy') AS ArrivalDate, FORMAT(BK.nihrm__DepartureDate__c, 'dd/MM/yyyy') AS DepartureDate, \
                           BK.Percentage_of_Attrition__c, BK.nihrm__CommissionPercentage__c, BK.Promotion__c, BK.nihrm__AtDefiniteAgreedRoomnights__c, BK.nihrm__CurrentBlendedRoomnightsTotal__c, BK.nihrm__AtDefiniteAgreedGuestroomRevenue__c, BK.nihrm__BlendedGuestroomRevenueTotal__c, \
                           BK.nihrm__AtDefiniteBlendedEventRevenue1__c, BK.nihrm__CurrentBlendedEventRevenue1__c, BK.nihrm__AtDefiniteBlendedEventRevenue2__c, BK.nihrm__CurrentBlendedEventRevenue2__c, BK.nihrm__AtDefiniteBlendedEventRevenue9__c, BK.nihrm__CurrentBlendedEventRevenue9__c, \
                           BK.VCL_Blended_F_B_Revenue__c, BK.nihrm__AtDefiniteBlendedEventRevenue7__c, BK.nihrm__CurrentBlendedEventRevenue7__c, BK.nihrm__AtDefiniteBlendedEventRevenue4__c, BK.nihrm__CurrentBlendedEventRevenue4__c, \
                           BK.nihrm__AtDefiniteBlendedEventRevenue3__c, BK.nihrm__CurrentBlendedEventRevenue3__c, BK.nihrm__AtDefiniteBlendedEventRevenue8__c, BK.nihrm__CurrentBlendedEventRevenue8__c, BK.nihrm__AtDefiniteBlendedEventRevenue6__c, \
                           BK.nihrm__CurrentBlendedEventRevenue6__c, BK.Sheraton_F_B_Revenue__c, BK.Sheraton_Room_Rental_Revenue__c, BK.nihrm__BookingStatus__c, FORMAT(BK.nihrm__LastStatusDate__c, 'dd/MM/yyyy') AS LastStatusDate, \
                           FORMAT(BK.nihrm__BookedDate__c, 'dd/MM/yyyy') AS BookedDate, BK.End_User_Region__c, BK.End_User_SIC__c, BK.nihrm__BookingTypeName__c, BK.nihrm__LostToCompetitorName__c, BK.nihrm__StatusReasonName__c, BK.Booking_ID_Number__c \
                    FROM dbo.nihrm__Booking__c AS BK \
                    LEFT JOIN dbo.Account AS ac \
                        ON BK.nihrm__Account__c = ac.Id \
                    LEFT JOIN dbo.Account AS ag \
                        ON BK.nihrm__Agency__c = ag.Id", conn)
data.columns = ['Booking: Owner Name', 'Property', 'Account', 'Company City', 'Company Country', 'Agency', 'Booking: Booking Post As', 'Arrival', 'Departure', 'Percentage of Attrition', 'Commission %', 'Promotion', 'At Definite Agreed Roomnights',
                'Blended Roomnights', 'At Definite Agreed Guestroom Revenue', 'Blended Guestroom Revenue Total', 'At Definite Blended Food Revenue', 'Blended Food Revenue', 'At Definite Blended Beverage Revenue', 'Blended Beverage Revenue',
                'At Definite Blended Outlet Revenue', 'Blended Outlet Revenue', 'Blended F&B Revenue', 'At Definite Blended Rental Revenue', 'Blended Rental Revenue', 'At Definite Blended AV Revenue', 'Blended AV Revenue', 'At Definite Blended Resource Revenue',
                'Blended Resource Revenue', 'At Definite Blended Ancillary Revenue', 'Blended Ancillary Revenue', 'At Definite Blended Other Revenue', 'Blended Other Revenue', 'Sheraton F&B Revenue', 'Sheraton Room Rental Revenue	', 'Status',
                'Last Status Date', 'Booked', 'End User Region', 'End User SIC', 'Booking Type', 'Lost to Competitor', 'Lost Reason', 'Booking ID#']
filename = '01Rawdata_Group Booking by arrival date'
convert_to_excel(data, filename)


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
