# Excel VBA Sunrise Sunset
Excel VBA Class for Sunrise, Sunset, Azimuth, and Elevation Calculations

> [!NOTE]
> 2023-12-17 - Programmed by Dennis Best - ExcelFunctions@natalko.com<br/>
> Calculation of local times of sunrise, solar noon, and sunset<br/>
> based on the calculation procedure by NOAA. See:<br/>
> https://gml.noaa.gov/grad/solcalc/calcdetails.html<br/>
> https://gml.noaa.gov/grad/solcalc/NOAA_Solar_Calculations_year.xls<br/><br/>
> Note that latitudes south of the equator are entered as negative<br/>
> And longitudes west of Greenwich are also entered as negative<br/>
> e.g. New York, USA : Latitude=  40.66, Longitude:  -73.94<br/>
> and  Sydney,  Aust : Latitude= -33.87, Longitude:  151.31<br/>

#### The following files are included:<br/>
1. clsSunriseSunset.cls - class module
2. modSunriseSunset.bas - module illustrating function calls
3. clsTimeZones.cls - class module<br/>
4. modTimeZones.bas - module illustrating function calls<br/>
5. Sunrise Sunset.xlsb - example spreadsheet<br/>
       


> [!NOTE]
> The SunriseSunset class requires the TimeZone class to determine the<br/> 
> offset hours from UTC for the given DateTime and Time Zone<br/>
> The Timezone class also makes corrections for DST<br/>
> For more information, see: https://github.com/denbe52/Excel-VBA-Time-Zone-Conversion<br/>

  
#### The available functions are:<br/>
1. SunriseSunset(Lat As Double, Lon As Double, DateTime As Date, Optional TZname As String = "")<br/>
2. AzimuthElevation(Lat As Double, Lon As Double, DateTime As Date, Optional TZname As String = "")<br/><br/> 

### Examples
```VBA
    =SunriseSunset(51, -115.1, now(), "Mountain Standard Time") 
```
> This function returns the following in one row:<br/>
> Column  1 - Azimuth   of the sun at the given DateTime<br/>
> Column  2 - Elevation of the sun at the given DateTime<br/>
> Column  3 - Time of Sunrise expressed relative to the specified time zone<br/>
> Column  4 - Azimuth of the sun at sunrise - note that the elevation of the sun at sunrise is -0.44 Deg<br/>
> Column  5 - Time of Solar Noon expressed relative to the specified time zone<br/>
> Column  6 - Elevation of the sun at Solar Noon - note that the azimuth of the sun at Solar noon is 180 Deg<br/>
> Column  7 - Time of Sunset expressed relative to the specified time zone<br/>
> Column  8 - Azimuth of the sun at sunset - note that the elevation of the sun at sunset is -0.44 Deg<br/>
> Column  9 - Hrs:Mins of Sunlight for the day<br/>
> Column 10 - Offset in hours from UTC for the specified Time Zone<br/>
> Column 11 - Time Zone Designation at the DateTime (e.g. Standard time or Summer Time)

<br/>

```VBA
    =AzimuthElevation(51, -115.1, now(), "Mountain Standard Time") 
```
> This function returns the following in one row:<br/>
> Column  1 - Azimuth   of the sun at the given DateTime<br/>
> Column  2 - Elevation of the sun at the given DateTime


<br/>

> [!IMPORTANT]
> If you install the modules into a new spreadsheet, you<br/> 
> must set a reference to the Outlook Library in the Visual Basic Editor.<br/>
>    In Excel, press Alt-F11 to open the VBA code editor.<br/>
>    Click on Tools, References and select "Microsoft Outlook 16.0 Object Library"

<br/>

> [!CAUTION]
> Bypass MalwareBytes Exploit Protection<br/>
> If you are using MalwareBytes and experience issues, you<br/>
> might need to make a modification to the settings in MalwareBytes.<br/>
> See: https://forums.malwarebytes.com/topic/78852-how-to-exclude-excel-addin-suddenly-showing-as-exploit-no-change-in-addin/
>
>```
>    Click Settings, Security, Advanced Settings (under Exploit Protection),
>    Advanced Exploit Protection Settings, Application behaviour protection tab
>    Remove check from both "Office VBA7 and VBE7 abuse protection" and Apply
>```

