Attribute VB_Name = "modSunriseSunset"

' =======================================================================================================================
' https://github.com/denbe52/Excel-VBA-Sunrise-Sunset
' Dennis Best - 2023-12-16 - dbExcelFunc@outlook.com
' =======================================================================================================================

' Calculation of local times of Sunrise, Solar Noon, Sunset, Elevation and Azimuth
' based on the calculation procedure by NOAA in the spreadsheet
' https://gml.noaa.gov/grad/solcalc/calcdetails.html
' https://gml.noaa.gov/grad/solcalc/NOAA_Solar_Calculations_year.xls
' https://www.timeanddate.com/astronomy/axial-tilt-obliquity.html

' Requires VBA Class module (clsSunriseSunset)

Option Explicit
Dim SS As New clsSunriseSunset  ' Define a global variable to avoid having to reload it on each call

' Return an array of sun tracking values at a given Latitude and Longitude for a given day and time
' Note that latitudes south of the equator are entered as negative
' And longitudes west of Greenwich are also entered as negative
' e.g. New York, USA : Latitude=  40.66, Longitude:  -73.94
' and  Sydney,  Aust : Latitude= -33.87, Longitude:  151.31
' If TZname (Time Zone Name) is missing it will default to the local time zone (on your computer)
' TZname must be entered as Standard time - it will be automatically converted to DST as necessary

' This function returns the following in one row:
' Column  1 - Azimuth   of the sun at the given DateTime
' Column  2 - Elevation of the sun at the given DateTime
' Column  3 - Time of Sunrise expressed relative to the specified time zone
' Column  4 - Azimuth of the sun at sunrise - note that the elevation of the sun at sunrise is -0.44 Deg
' Column  5 - Time of Solar Noon expressed relative to the specified time zone
' Column  6 - Elevation of the sun at Solar Noon - note that the azimuth of the sun at Solar noon is 180 Deg
' Column  7 - Time of Sunset expressed relative to the specified time zone
' Column  8 - Azimuth of the sun at sunset - note that the elevation of the sun at sunset is -0.44 Deg
' Column  9 - Hrs:Mins of Sunlight for the day
' Column 10 - Offset in hours from UTC for the specified Time Zone
' Column 11 - Time Zone Designation at the DateTime (e.g. Standard time or Summer Time)

Function SunriseSunset(Lat As Double, Lon As Double, DateTime As Date, Optional TZname As String = "") As Variant
    
    SS.Lat = Lat                        ' Latitudes South of the equator are negative
    SS.Lon = Lon                        ' Longitudes West of Greenwich are negative

    SS.DateTime = DateTime              ' Specify the DateTime for the calculations
    SS.TZname = TZname                  ' e.g. "Mountain Standard Time"
    
    SS.Calculate                        ' Calculate based on the preceding parameters

    Dim var(1 To 11) As Variant         ' Return an array of the results
    
    var(1) = SS.Azim_DateTime           ' Azimuth   of the sun at the given DateTime
    var(2) = SS.Elev_DateTime           ' Elevation of the sun at the given DateTime

    var(3) = SS.Time_Sunrise            ' Time of Sunrise
    var(4) = SS.Azim_Sunrise            ' Elevation at Sunrise = -0.44 Deg

    var(5) = SS.Time_SolarNoon          ' Time of SolarNoon
    var(6) = SS.Elev_SolarNoon          ' Azimuth at SolarNoon = 180.0 Deg

    var(7) = SS.Time_Sunset             ' Time of Sunset
    var(8) = SS.Azim_Sunset             ' Elevation at Sunset  = -0.44 Deg

    var(9) = SS.SunlightMinutes / 1440  ' convert minutes to fraction of a day. Format it in the spreadsheet as hh:mm:ss
    var(10) = SS.UTC_Offset             ' Offset from UTC in hours
    var(11) = SS.TZname_current         ' e.g. Standard Time or Summer Time

    SunriseSunset = var                 ' Return 11 values to the spreadsheet in one row
End Function

' This function returns the following in one row:
' Column  1 - Azimuth   of the sun at the given DateTime
' Column  2 - Elevation of the sun at the given DateTime
Function AzimuthElevation(Lat As Double, Lon As Double, DateTime As Date, Optional TZname As String = "") As Variant
    
    SS.Lat = Lat
    SS.Lon = Lon

    SS.DateTime = DateTime
    SS.TZname = TZname
    
    SS.Calculate

    Dim var(1 To 2) As Variant
    var(1) = SS.Azim_DateTime
    var(2) = SS.Elev_DateTime

    AzimuthElevation = var             ' Return Azimuth and Elevation at the given DateTime
End Function

