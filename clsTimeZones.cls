VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTimeZones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' =======================================================================================================================
' https://github.com/denbe52/Excel-VBA-Time-Zone-Conversion
' Dennis Best - 2023-12-16 - dbExcelFunc@outlook.com
' =======================================================================================================================

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

' Use a global variable, TZ, to keep a copy of the Outlook.TimeZones alive - no need to reload it every time.
Private TZ As Object

' Load Time Zone data from Outlook.TimeZones into global variable TZ
' Returns TRUE if successful
Private Function LoadTimeZoneData() As Boolean
    If TZ Is Nothing Then           ' Only load the TZ object once to speed up the calculations
        Set TZ = Outlook.TimeZones
        Debug.Print "Loaded Outlook.TimeZones at " & Format(Now, "yyyy-mm-dd hh:mm:ss")
'   Uncomment the next 9 lines if you have problems getting a TZ object
    Else
        On Error Resume Next
        If TZ.Count > 0 Then                                 ' This throws an error if TZ is invalid
'            Debug.Print "TimeZones Okay"
        Else                                                 ' If it is an error, then try one more time
            Debug.Print "Trying again"
            Set TZ = Nothing                                 ' Clear TZ first
            Set TZ = Outlook.TimeZones                       ' Retry loading
        End If
    End If
    LoadTimeZoneData = Not (TZ Is Nothing)                    ' TRUE if successful
End Function

' Return the Standard Designation for the Local Time Zone
Public Function GetLocalTimeZone()
    Dim TZlocal As Object
    If LoadTimeZoneData() Then
        Set TZlocal = TZ.CurrentTimeZone
        GetLocalTimeZone = TZlocal.ID
    Else
        GetLocalTimeZone = "Outlook Problem"
    End If
End Function

' The Outlook function, TZ.ConvertTime(), apparently rounds DateTime up or down to the nearest minute
' Function Add_seconds converts the seconds to a DateTime value
Private Function Add_Seconds(DateTime As Variant) As Double
    Add_Seconds = (Second(DateTime) - Int(Second(DateTime) / 30) * 60) / 86400
End Function

' Convert DateTime from Source time zone to Destination time zone
' This function uses Outlook, and properly handles Daylight Saving Times (past and future)
' TZ_from, TZ_to = e.g. "Mountain Standard Time", "AUS Eastern Standard Time", etc.
' These must always be expressed in "Standard" time.
' See "ListAllTimeZones" for valid list of time zones - TZ.ID - TZ_from and TZ_to
Public Function ConvertDateTime(DateTime As Variant, TZ_from As String, TZ_to As String) As Variant
    Dim TZsrce As Object, TZdest As Object
    If LoadTimeZoneData() Then
        Set TZsrce = TZ.Item(TZ_from)
        Set TZdest = TZ.Item(TZ_to)
        If TZsrce Is Nothing Then
            ConvertDateTime = "Check TZ_from"
        ElseIf TZdest Is Nothing Then
            ConvertDateTime = "Check TZ_to"
        ElseIf VarType(DateTime) <> vbDate Then
            ConvertDateTime = "Check DateTime"
        End If
        If VarType(ConvertDateTime) <> vbString Then
            ConvertDateTime = TZ.ConvertTime(DateTime, TZsrce, TZdest) + Add_Seconds(DateTime)
        End If
    Else
        ConvertDateTime = "Outlook Problem"
    End If
End Function

'' Return the offset in hours between two timezones (includes daylight savings time adjustments)
'' TZ_ID = "Mountain Standard Time", "AUS Eastern Standard Time", etc.
'' TZ_ID = Current Time Zone (if missing)
Public Function Offset_Hrs(DateTime As Variant, TZ_from As String, TZ_to As String) As Variant
    Dim ConvertedTime As Variant
    If TZ_from = "" Then TZ_from = GetLocalTimeZone()
    ConvertedTime = ConvertDateTime(DateTime, TZ_from, TZ_to)
    If VarType(ConvertedTime) <> vbDate Then
        Offset_Hrs = ConvertedTime
    Else
        Offset_Hrs = Application.Round((DateTime - ConvertedTime) * 24, 2)
    End If
End Function

'' Determine whether a given DateTime in a given TimeZone is using Daylight Savings Time
'' Returns "DST", "-" or "Check Data"
Public Function isDST(DateTime As Variant, ReferenceTimeZone As String) As Variant
    Dim UTC As Variant, UTC_Offset As Integer, Bias As Integer
    UTC = ConvertDateTime(DateTime, ReferenceTimeZone, "UTC")
    If VarType(UTC) <> vbDate Then
        isDST = "Check Data"
    Else
        UTC_Offset = Application.Round((DateTime - UTC) * 1440, 0)
        Bias = TZ.Item(ReferenceTimeZone).Bias
        isDST = IIf(UTC_Offset + Bias = 0, "-", "DST")
    End If
End Function

Public Function CurrentTimeZoneDesignation(DateTime As Variant, ReferenceTimeZone As String) As Variant
    Dim DST As Variant, TZone As Object
    If LoadTimeZoneData() Then
        Set TZone = TZ.Item(ReferenceTimeZone)
        If Not (TZone Is Nothing) Then
            DST = isDST(DateTime, ReferenceTimeZone)
            Select Case DST
                Case "DST"
                    CurrentTimeZoneDesignation = TZone.DaylightDesignation
                Case "-"
                    CurrentTimeZoneDesignation = TZone.ID
                Case Else
                    CurrentTimeZoneDesignation = "Check Data"
            End Select
        Else
            CurrentTimeZoneDesignation = "Check TZ"
        End If
    Else
        CurrentTimeZoneDesignation = "Outlook Problem"
    End If
End Function

' Return the Standard Designation TZ.IDs that can be used in List boxes
Public Function ListTimeZones()
    Dim numTZ As Integer, var() As Variant, I As Integer
    If LoadTimeZoneData() Then
        numTZ = TZ.Count
        ReDim var(0 To numTZ + 1)
        var(0) = "TimeZones"
        var(1) = ""
        For I = 1 To numTZ
            var(I + 1) = TZ(I).ID
        Next I
        ListTimeZones = Application.WorksheetFunction.Transpose(var)
    Else
        ListTimeZones = "Outlook Problem"
    End If
End Function

' Returns: Item, TZ.Name, TZ.ID, Bias, Standard Date, Daylight Date, Daylight Designation, UTC Offset, IsDST, DateTime
Public Function ListAllTimeZoneData(Optional DateTime As Date, Optional ReferenceTimeZone As String = "")
    Dim TZcurrent As Object, numTZ As Integer, var() As Variant, I As Integer, UTC As Date, TZone As Object
    If LoadTimeZoneData() Then
        numTZ = TZ.Count
        If DateTime = #12:00:00 AM# Then DateTime = Now  ' Default DateTime is now()
        If ReferenceTimeZone = "" Then                   ' Reference Time Zone is Local TimeZone
            Set TZcurrent = TZ.CurrentTimeZone
            ReferenceTimeZone = TZcurrent.ID
        Else
            Set TZone = TZ.Item(ReferenceTimeZone)
            If TZone Is Nothing Then
                ListAllTimeZoneData = "Check TZ"
                Exit Function
            End If
        End If
        UTC = TZ.ConvertTime(DateTime, TZ.Item(ReferenceTimeZone), TZ.Item("UTC")) + Add_Seconds(DateTime)

        ReDim var(0 To numTZ, 1 To 10)
        var(0, 1) = "TZ.Item"
        var(0, 2) = "TZ.Name"
        var(0, 3) = "TZ.ID (Standard Designation)"
        var(0, 4) = "Bias minutes"
        var(0, 5) = "Standard Date"
        var(0, 6) = "Daylight Date"
        var(0, 7) = "Daylight Designation"
        var(0, 8) = "UTC Offset hrs"
        var(0, 9) = "Is DST?"
        var(0, 10) = "DateTime"

        For I = 1 To numTZ
            var(I, 1) = I
            var(I, 2) = TZ(I).Name
            var(I, 3) = TZ(I).ID
            var(I, 4) = TZ(I).Bias
            If TZ(I).StandardDate = #1/1/1601# Then
                var(I, 5) = ""
                var(I, 6) = ""
            Else
                var(I, 5) = TZ(I).StandardDate
                var(I, 6) = TZ(I).DaylightDate
            End If
            var(I, 7) = TZ(I).DaylightDesignation
            var(I, 8) = TZ.ConvertTime(DateTime, TZ.Item(I), TZ.Item("UTC")) + Add_Seconds(DateTime)
            var(I, 8) = Application.Round((DateTime - var(I, 8)) * 24, 2)
            var(I, 9) = IIf(Application.Round((var(I, 4) + var(I, 8) * 60), 0) = 0, "", "DST")
            var(I, 10) = UTC + var(I, 8) / 24
        Next I
        ListAllTimeZoneData = var
    End If
End Function

