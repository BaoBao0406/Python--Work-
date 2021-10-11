import pyodbc, datetime
import pandas as pd

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=VOPPSCLDBN01\VOPPSCLDBI01;'
                      'Database=SalesForce;'
                      'Trusted_Connection=yes;')

# temp file path
save_path = 'I:\\10-Sales\\Personal Folder\\Admin & Assistant Team\\Patrick Leong\\Python Code\\Revenue DataPipeline\\'

# TODO: Date Range
now = datetime.datetime.now()
EndDate = now + datetime.timedelta(days=21)
Current_Date = str(now.year) + '-' + str('%02d'% now.month) + '-' + str('%02d'% now.day)
End_Date = str(EndDate.year) + '-' + str('%02d'% EndDate.month) + '-' + str('%02d'% EndDate.day)


# Convert data to excel format
def convert_to_excel(data, filename):
    data.to_excel(save_path + filename + '.xlsx', sheet_name='Sheet1', index=False)


# FDC User ID and Name list
user = pd.read_sql('SELECT DISTINCT(Id), Name \
                    FROM dbo.[User]', conn)
user = user.set_index('Id')['Name'].to_dict()


# "01Rawdata_Group Booking by arrival date" - Report
data = pd.read_sql("SELECT BK.OwnerId, BK.nihrm__Property__c, ac.Name, ac.BillingCity, ac.BillingCountry, ag.Name, BK.Name, FORMAT(BK.nihrm__ArrivalDate__c, 'MM/dd/yyyy') AS ArrivalDate, FORMAT(BK.nihrm__DepartureDate__c, 'MM/dd/yyyy') AS DepartureDate, \
                           BK.Percentage_of_Attrition__c, BK.nihrm__CommissionPercentage__c, BK.Promotion__c, BK.nihrm__AtDefiniteAgreedRoomnights__c, BK.nihrm__CurrentBlendedRoomnightsTotal__c, BK.nihrm__AtDefiniteAgreedGuestroomRevenue__c, BK.nihrm__BlendedGuestroomRevenueTotal__c, \
                           BK.nihrm__AtDefiniteBlendedEventRevenue1__c, BK.nihrm__CurrentBlendedEventRevenue1__c, BK.nihrm__AtDefiniteBlendedEventRevenue2__c, BK.nihrm__CurrentBlendedEventRevenue2__c, BK.nihrm__AtDefiniteBlendedEventRevenue9__c, BK.nihrm__CurrentBlendedEventRevenue9__c, \
                           BK.VCL_Blended_F_B_Revenue__c, BK.nihrm__AtDefiniteBlendedEventRevenue7__c, BK.nihrm__CurrentBlendedEventRevenue7__c, BK.nihrm__AtDefiniteBlendedEventRevenue4__c, BK.nihrm__CurrentBlendedEventRevenue4__c, \
                           BK.nihrm__AtDefiniteBlendedEventRevenue3__c, BK.nihrm__CurrentBlendedEventRevenue3__c, BK.nihrm__AtDefiniteBlendedEventRevenue8__c, BK.nihrm__CurrentBlendedEventRevenue8__c, BK.nihrm__AtDefiniteBlendedEventRevenue6__c, \
                           BK.nihrm__CurrentBlendedEventRevenue6__c, BK.Sheraton_F_B_Revenue__c, BK.Sheraton_Room_Rental_Revenue__c, BK.nihrm__BookingStatus__c, FORMAT(BK.nihrm__LastStatusDate__c, 'MM/dd/yyyy') AS LastStatusDate, \
                           FORMAT(BK.nihrm__BookedDate__c, 'MM/dd/yyyy') AS BookedDate, BK.End_User_Region__c, BK.End_User_SIC__c, BK.nihrm__BookingTypeName__c, BK.nihrm__LostToCompetitorName__c, BK.nihrm__StatusReasonName__c, BK.Booking_ID_Number__c \
                    FROM dbo.nihrm__Booking__c AS BK \
                    LEFT JOIN dbo.Account AS ac \
                        ON BK.nihrm__Account__c = ac.Id \
                    LEFT JOIN dbo.Account AS ag \
                        ON BK.nihrm__Agency__c = ag.Id \
                    WHERE (BK.nihrm__BookingTypeName__c NOT IN ('CS Catering - Social', 'ALT Alternative', 'CC Catering - Corporate', 'CN Concert', 'IN Internal')) AND \
                        (BK.nihrm__Property__c NOT IN ('Sands Macao Hotel')) AND (BK.nihrm__BookingStatus__c IN ('Tentative', 'Definite', 'Prospect', 'Cancelled', 'TurnedDown'))", conn)
data.columns = ['Booking: Owner Name', 'Property', 'Account', 'Company City', 'Company Country', 'Agency', 'Booking: Booking Post As', 'Arrival', 'Departure', 'Percentage of Attrition', 'Commission %', 'Promotion', 'At Definite Agreed Roomnights',
                'Blended Roomnights', 'At Definite Agreed Guestroom Revenue', 'Blended Guestroom Revenue Total', 'At Definite Blended Food Revenue', 'Blended Food Revenue', 'At Definite Blended Beverage Revenue', 'Blended Beverage Revenue',
                'At Definite Blended Outlet Revenue', 'Blended Outlet Revenue', 'Blended F&B Revenue', 'At Definite Blended Rental Revenue', 'Blended Rental Revenue', 'At Definite Blended AV Revenue', 'Blended AV Revenue', 'At Definite Blended Resource Revenue',
                'Blended Resource Revenue', 'At Definite Blended Ancillary Revenue', 'Blended Ancillary Revenue', 'At Definite Blended Other Revenue', 'Blended Other Revenue', 'Sheraton F&B Revenue', 'Sheraton Room Rental Revenue	', 'Status',
                'Last Status Date', 'Booked', 'End User Region', 'End User SIC', 'Booking Type', 'Lost to Competitor', 'Lost Reason', 'Booking ID#']
data['Booking: Owner Name'].replace(user, inplace=True)
filename = '01Rawdata_Group Booking by arrival date'
convert_to_excel(data, filename)


# "02Account and Activities" - Report
column_name = ['Account ID', 'Account Name', 'Account Type', 'Account Owner', 'Region', 'Industry', 'Activity ID', 'Assigned', 'Type', 'Subject', 'Created Date', 'Start', 'Last Modified Date', 'Status']
ac_event = pd.read_sql("SELECT ac.Id, ac.Name, ac.Type, ac.OwnerId, ac.nihrm__RegionName__c, ac.Industry, ev.Id, ev.OwnerId, ev.Type, ev.Subject, FORMAT(ev.CreatedDate, 'MM/dd/yyyy') AS CreatedDate, FORMAT(ev.StartDateTime, 'MM/dd/yyyy') AS Start, \
                               FORMAT(ev.LastModifiedDate, 'MM/dd/yyyy') AS LastModifiedDate, ev.VCL_Status__c \
                        FROM dbo.Event AS ev \
                        INNER JOIN dbo.Account AS ac \
                            ON ev.WhatId = ac.Id \
                        WHERE ev.CreatedDate BETWEEN CONVERT(datetime, '2021-01-01') AND CONVERT(datetime, '2045-12-31')", conn)
ac_event.columns = column_name

ac_task = pd.read_sql("SELECT ac.Id, ac.Name, ac.Type, ac.OwnerId, ac.nihrm__RegionName__c, ac.Industry, tk.Id, tk.OwnerId, tk.Type, tk.Subject, FORMAT(tk.CreatedDate, 'MM/dd/yyyy') AS CreatedDate, FORMAT(tk.ActivityDate, 'MM/dd/yyyy') AS Start, \
                              FORMAT(tk.LastModifiedDate, 'MM/dd/yyyy') AS LastModifiedDate, tk.Status \
                      FROM dbo.Task AS tk \
                      INNER JOIN dbo.Account AS ac \
                          ON tk.WhatId = ac.Id \
                      WHERE tk.CreatedDate BETWEEN CONVERT(datetime, '2021-01-01') AND CONVERT(datetime, '2045-12-31')", conn)
ac_task.columns = column_name
# Concate event and task table
ac_activities = pd.concat([ac_event, ac_task])
ac_activities['Account Owner'].replace(user, inplace=True)
ac_activities['Assigned'].replace(user, inplace=True)
filename = '02Account and Activities'
convert_to_excel(ac_activities, filename)


# "03Agency and Booking ID" - Report
data = pd.read_sql('SELECT nihrm__Account__c, Booking_ID_Number__c \
                    FROM dbo.nihrm__Booking__c', conn)
data.columns = ['Account: Account ID', 'Booking ID#']
filename = '03Agency and Booking ID'
convert_to_excel(data, filename)


# "04Account and Booking ID" - Report
data = pd.read_sql('SELECT nihrm__Agency__c, Booking_ID_Number__c \
                    FROM dbo.nihrm__Booking__c \
                    WHERE nihrm__Agency__c IS NOT NULL', conn)
data.columns = ['Agency: Account ID', 'Booking ID#']
filename = '04Account and Booking ID'
convert_to_excel(data, filename)


# "05Booking and Activities not started" - Report
column_name = ['Booking: ID', 'Booking ID#', 'Booking: Booking Post As', 'Status', 'Booking: Owner Name', 'End User SIC', 'End User Region', 'Arrival', 'Booked', 'Activity ID', 'Assigned', 'Subject', 'Type', 'Created Date', 'Start', 'Last Modified Date', 'Status', 'Account', 'Agency']
bk_event = pd.read_sql("SELECT BK.Id, BK.Booking_ID_Number__c, BK.Name, BK.nihrm__BookingStatus__c, BK.OwnerId, BK.End_User_SIC__c, BK.End_User_Region__c, FORMAT(BK.nihrm__ArrivalDate__c, 'MM/dd/yyyy') AS ArrivalDate, FORMAT(BK.nihrm__BookedDate__c, 'MM/dd/yyyy') AS BookedDate, ev.Id, ev.OwnerId, ev.Subject, ev.Type, FORMAT(ev.CreatedDate, 'MM/dd/yyyy') AS CreatedDate, FORMAT(ev.StartDateTime, 'MM/dd/yyyy') AS Start, FORMAT(ev.LastModifiedDate, 'MM/dd/yyyy') AS LastModifiedDate, ev.VCL_Status__c, ac.Name, ag.Name \
                        FROM dbo.nihrm__Booking__c AS BK \
                        INNER JOIN dbo.Event AS ev \
                            ON BK.Id = ev.WhatId \
                        LEFT JOIN dbo.Account AS ac \
                             ON BK.nihrm__Account__c = ac.Id \
                         LEFT JOIN dbo.Account AS ag \
                             ON BK.nihrm__Agency__c = ag.Id \
                        WHERE BK.nihrm__ArrivalDate__c BETWEEN CONVERT(datetime, '2021-01-18') AND CONVERT(datetime, '2045-12-31')", conn)
bk_event.columns = column_name

bk_task = pd.read_sql("SELECT BK.Id, BK.Booking_ID_Number__c, BK.Name, BK.nihrm__BookingStatus__c, BK.OwnerId, BK.End_User_SIC__c, BK.End_User_Region__c, FORMAT(BK.nihrm__ArrivalDate__c, 'MM/dd/yyyy') AS ArrivalDate, FORMAT(BK.nihrm__BookedDate__c, 'MM/dd/yyyy') AS BookedDate, tk.Id, tk.OwnerId, tk.Subject, tk.Type, FORMAT(tk.CreatedDate, 'MM/dd/yyyy') AS CreatedDate, FORMAT(tk.ActivityDate, 'MM/dd/yyyy') AS Start, FORMAT(tk.LastModifiedDate, 'MM/dd/yyyy') AS LastModifiedDate, tk.Status, ac.Name, ag.Name \
                       FROM dbo.nihrm__Booking__c AS BK \
                       INNER JOIN dbo.Task AS tk \
                           ON BK.Id = tk.WhatId \
                       LEFT JOIN dbo.Account AS ac \
                           ON BK.nihrm__Account__c = ac.Id \
                       LEFT JOIN dbo.Account AS ag \
                           ON BK.nihrm__Agency__c = ag.Id \
                       WHERE BK.nihrm__ArrivalDate__c BETWEEN CONVERT(datetime, '2021-01-18') AND CONVERT(datetime, '2045-12-31')", conn)
bk_task.columns = column_name
# Concate event and task table
bk_activities = pd.concat([bk_event, bk_task])
bk_activities['Booking: Owner Name'].replace(user, inplace=True)
bk_activities['Assigned'].replace(user, inplace=True)
filename = '05Booking and Activities not started'
convert_to_excel(bk_activities, filename)


# "06Event max attendance" - Report
data = pd.read_sql('SELECT dbo.nihrm__Booking__c.Booking_ID_Number__c, nihrm__AgreedAttendance__c \
                    FROM dbo.nihrm__BookingEvent__c \
                    INNER JOIN dbo.nihrm__Booking__c \
                        ON dbo.nihrm__BookingEvent__c.nihrm__Booking__c = dbo.nihrm__Booking__c.Id \
                    WHERE nihrm__AgreedAttendance__c > 0', conn)
data.columns = ['Booking: Booking ID#', 'Agreed']
filename = '06Event max attendance'
convert_to_excel(data, filename)


# "07roomnight peak data and number" - Report
data = pd.read_sql("SELECT BK.Booking_ID_Number__c, GS.nihrm__Property__c, GS.Name, BK.Name, FORMAT(RoomN.nihrm__PatternDate__c, 'MM/dd/yyyy') AS PatternDate, \
                           RoomN.nihrm__AgreedRoomsTotal__c, RoomN.nihrm__PickupRoomsTotal__c \
                    FROM dbo.nihrm__BookingRoomNight__c AS RoomN \
                    INNER JOIN dbo.nihrm__Booking__c AS BK \
                        ON RoomN.nihrm__Booking__c = BK.Id \
                    INNER JOIN dbo.nihrm__GuestroomType__c AS GS \
                        ON RoomN.nihrm__GuestroomType__c = GS.Id \
                    WHERE RoomN.nihrm__PatternDate__c BETWEEN CONVERT(datetime, '2018-01-01') AND CONVERT(datetime, '2045-12-31')", conn)
data.columns = ['Booking: Booking ID#', 'Property', 'Guestroom Type', 'Booking: Booking Post As', 'Pattern Date', 'Agreed Rooms Total', 'Pickup Rooms Total']
filename = '07roomnight peak data and number'
convert_to_excel(data, filename)


# "08Account Information" - Report
data = pd.read_sql("SELECT ac.Id, owner.Name, ac.Name, ac.Type, ac.nihrm__RegionName__c, ac.Industry, ac.BillingCountry, ac.BillingState, ac.BillingCity, \
                           ac.Rating, ac.nihrm__MarketSegment__c, FORMAT(ac.LastModifiedDate, 'MM/dd/yyyy') AS LastModifiedDate, \
                           FORMAT(ac.LastActivityDate, 'MM/dd/yyyy') AS LastActivityDate, FORMAT(ac.CreatedDate, 'MM/dd/yyyy') AS CreatedDate \
                    FROM dbo.Account AS ac \
                    INNER JOIN dbo.[User] AS owner \
                        ON ac.OwnerId = owner.Id", conn)
data.columns = ['Account ID', 'Account Owner', 'Account Name', 'Type', 'Region', 'Industry', 'Country', 'State/Province', 'City', 'Quality Rating', 
                'Market Segment', 'Last Modified Date', 'Last Activity', 'Created Date']
filename = '08Account Information'
convert_to_excel(data, filename)
