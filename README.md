## Excel VBA Time Zone Conversion using the Outlook Library
 
> [!NOTE]
> 2023-12-16 - Programmed by Dennis Best - dbExcelFunc@outlook.com<br/>
> Adapted from: Julian Hess and Patrick Honorez<br/> 
> https://stackoverflow.com/a/45510712/78522

#### The following files are included:<br/>
1. clsTimeZones.cls - class module<br/>
2. modTimeZones.bas - module illustrating function calls<br/>
3. TimeZones.xlsb - example spreadsheet<br/>
       

  
### Excel VBA functions to calculate the time in other time zones
> These functions use the Microsoft Outlook 16.0 Object library to convert the date and time in a reference time zone to another destimation time zone, and include corrections for Daylight Savings Time.

> #### The available functions are:<br/>
1. ConvertDateTime(DateTime As Date, Reference_TZ As String, Destination_TZ As String)<br/>
2. isDST(DateTime As Date, Reference_TZ As String)<br/>
3. CurrentTimeZoneDesignation(DateTime as Date, ReferenceTimeZone as string)<br/>
4. Offset_Hrs(DateTime As Variant, TZ_from As String, TZ_to As String)<br/>
5. ListTimeZones()<br/>
6. ListAllTimeZoneData(Optional DateTime, Optional ReferenceTimeZone)<br/>
7. GetLocalTimeZone()<br/><br/> 

### Examples
```VBA
    =ConvertDateTime(now(), "Mountain Standard Time", "Eastern Standard Time") 
```
> returns "16 Dec 23 10:12:09"

<br/>

```VBA
    =isDST(now(), "Eastern Standard Time") 
```
> returns "DST" in the summer, and "-" in the winter

<br/>

```VBA
    =CurrentTimeZoneDesignation(=now(), "Mountain Standard Time") 
```
> returns "Mountain Summer Time" in the summer and "Mountain Standard Time" in the winter

<br/>

```VBA
    =Offset_Hrs(now(), "Mountain Standard Time", "Eastern Standard Time") 
```
> returns -2.0

<br/>

```VBA   
    =ListTimeZones()
```
> returns a column of the Standard Time Zones (about 141 Time Zones)

<br/>

```VBA
    =ListAllTimeZoneData(=now(), "Mountain Standard Time")
```
> returns a table of TZ.Name, TZ.ID, Bias, Standard Date, Daylight Date, 
> Daylight Designation, UTC_Offset, isDST, and DateTime in the Time Zone
> for all Time Zones. <br/>

> Note that you can optionally specify a DateTime and a TimeZone as the basis for the 
> calculation of the DateTime in all of the other time zones. <br/>

> The default DateTime and ReferenceTimeZone is the current DateTime and local time zone.

<br/>

```VBA
    =GetLocalTimeZone() 
```
> returns "Mountain Standard Time" (i.e. your computer's time zone)

<br/>

> [!IMPORTANT]
> <u>If you install the two modules into a new spreadsheet, you</u><br/> 
> must set a reference to the Outlook Library in the Visual Basic Editor.<br/>
>    In Excel, press Alt-F11 to open the VBA code editor.<br/>
>    Click on Tools, References and select "Microsoft Outlook 16.0 Object Library"

<br/>

> [!CAUTION]
> <u>Bypass MalwareBytes Exploit Protection</u><br/>
> If you are using MalwareBytes and experience issues, you<br/>
> might need to make a modification to the settings in MalwareBytes.<br/>
> See: https://forums.malwarebytes.com/topic/78852-how-to-exclude-excel-addin-suddenly-showing-as-exploit-no-change-in-addin/
>
>```
>    Click Settings, Security, Advanced Settings (under Exploit Protection),
>    Advanced Exploit Protection Settings, Application behaviour protection tab
>    Remove check from both "Office VBA7 and VBE7 abuse protection" and Apply
>```
