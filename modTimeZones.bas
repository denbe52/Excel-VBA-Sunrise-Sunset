Attribute VB_Name = "modTimeZones"
' https://github.com/denbe52/Excel-VBA-Time-Zone-Conversion
' Dennis Best - 2023-12-16 - ExcelFunctions@natalko.com

' =======================================================================================================================
' Functions to convert date and time between time zones using Outlook
' 2023-12-16 - Programmed by Dennis Best
' Adapted From: https://stackoverflow.com/questions/3120915/get-timezone-information-in-vba-excel
' Credits to Patrick Honorez and Julian Hess https://stackoverflow.com/a/45510712/78522

' =======================================================================================================================
' IMPORTANT - You must set a reference to the Outlook Library in the Visual Basic Editor
' Click on Tools, References and add "Microsoft Outlook 16.0 Object Library"
' =======================================================================================================================

' =======================================================================================================================
' Bypass MalwareBytes Exploit Protection - do this if you use MalwareBytes and have problems
' https://forums.malwarebytes.com/topic/278852-how-to-exclude-excel-addin-suddenly-showing-as-exploit-no-change-in-addin/
' Click Settings, Security, Advanced Settings (under Exploit Protection),
' Advanced Exploit Protection Settings, Application behaviour protection tab
' Remove check from both "Office VBA7 and VBE7 abuse protection" and Apply
' =======================================================================================================================

Option Explicit

Public TZ As New clsTimeZones ' Declare as global variable to speed up code execution - no need to reload for every call

' Converts the DateTime from Reference Time Zone to Destination Time Zone
Public Function ConvertDateTime(DateTime As Date, Reference_TZ As String, Destination_TZ As String)
    ConvertDateTime = TZ.ConvertDateTime(DateTime, Reference_TZ, Destination_TZ)
End Function

' Returns "DST" if the DateTime in the Reference Time Zone is DST - Returns "-" otherwise
Public Function isDST(DateTime As Date, Reference_TZ As String)
    isDST = TZ.isDST(DateTime, Reference_TZ)
End Function

' Returns the Current Time Zone Designation in the Reference Time Zone
' e.g. "Mountain Standard Time" or "Mountain Summer Time" (if DST)
Public Function CurrentTimeZoneDesignation(DateTime As Variant, ReferenceTimeZone As String)
    CurrentTimeZoneDesignation = TZ.CurrentTimeZoneDesignation(DateTime, ReferenceTimeZone)
End Function

' Returns the Offset Hours between two Time Zones (corrected for DST in either or both Time Zones)
Public Function Offset_Hrs(DateTime As Variant, ReferenceTimeZone As String, Destination_TZ As String)
    Offset_Hrs = TZ.Offset_Hrs(DateTime, ReferenceTimeZone, Destination_TZ)
End Function

' Returns a single column of the Standard Time Zone Designations (about 141 Time Zones)
' Use this data in a data validation list to simplify entry of the Time Zone names
Public Function ListTimeZones()
    ListTimeZones = TZ.ListTimeZones
End Function

' Returns a 10 column summary of all of the data for each time zone - e.g. Time Zone bias, start and end dates for DST, etc.
Public Function ListAllTimeZoneData(Optional DateTime As Date, Optional ReferenceTimeZone As String = "")
    ListAllTimeZoneData = TZ.ListAllTimeZoneData(DateTime, ReferenceTimeZone)
End Function

' Returns the name of the Local Time zone (i.e. the Time Zone that your computer is set to)
Public Function GetLocalTimeZone()
    GetLocalTimeZone = TZ.GetLocalTimeZone
End Function

