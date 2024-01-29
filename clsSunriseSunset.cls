VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSunriseSunset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' =======================================================================================================================
' https://github.com/denbe52/Excel-VBA-Sunrise-Sunset
' Dennis Best - 2023-12-16 - dbExcelFunc@outlook.com
' =======================================================================================================================

' Calculation of local times of sunrise, solar noon, and sunset
' based on the calculation procedure by NOAA in the spreadsheet
' https://gml.noaa.gov/grad/solcalc/calcdetails.html
' https://gml.noaa.gov/grad/solcalc/NOAA_Solar_Calculations_year.xls

' Note that latitudes south of the equator are entered as negative
' And longitudes west of Greenwich are also entered as negative
' e.g. New York, USA : Latitude=  40.66, Longitude:  -73.94
' and  Sydney,  Aust : Latitude= -33.87, Longitude:  151.31

Option Explicit

Private Const PI As Double = 3.14159265358979           ' Pi
Private Const PI_2 As Double = 1.5707963267949          ' Pi/2
Private Const DtoR As Double = 1.74532925199433E-02 ' Radians per degree
Private Const RtoD As Double = 57.2957795130823     ' Degrees per radian

Private TZ As New clsTimeZones      ' Used to calculate Offset from UTC

Private m_Lat As Double
Private m_Lon As Double
Private m_TZname As String          ' Optional - Defaults to Current Time Zone
Private m_TZname_current As String  ' Returns the current name of the timezone (e.g. Standard or Summer)
Private m_UTC_Offset As Double      ' Offset of Time zone from UTC in hours

Private m_DateTime As Date

Private m_Time_SolarNoon As Double
Private m_Time_Sunrise As Double
Private m_Time_Sunset As Double
Private m_SunlightMinutes As Double

Private m_Elev_DateTime As Double
Private m_Elev_Sunrise As Double
Private m_Elev_SolarNoon As Double
Private m_Elev_Sunset As Double

Private m_Azim_DateTime As Double
Private m_Azim_Sunrise As Double
Private m_Azim_SolarNoon As Double
Private m_Azim_Sunset As Double

' Location Information ===============================================================
Property Let Lat(ByVal Lat As Double)       ' Latitudes south of the equator are entered as negative
    m_Lat = Lat
End Property

Property Let Lon(ByVal Lon As Double)       ' Longitudes west of Greenwich are entered as negative
    m_Lon = Lon
End Property

Property Let TZname(ByVal TZname As String) ' e.g. "Mountain Standard Time"
    m_TZname = TZname
End Property

Property Get TZname_current() As String     ' Current name of the Time Zone (e.g. Standard or Summer)
    TZname_current = m_TZname_current
End Property

Property Get UTC_Offset() As Double         ' Offset in hours from UTC
    UTC_Offset = m_UTC_Offset
End Property

' Time ===============================================================================
Property Let DateTime(ByVal DateTime As Date)
    m_DateTime = DateTime
End Property

Property Get Time_Sunrise() As Double
    Time_Sunrise = m_Time_Sunrise
End Property

Property Get Time_SolarNoon() As Double
    Time_SolarNoon = m_Time_SolarNoon
End Property

Property Get Time_Sunset() As Double
    Time_Sunset = m_Time_Sunset
End Property

' Elevation ==========================================================================
Property Get Elev_DateTime() As Variant
    Elev_DateTime = m_Elev_DateTime
End Property
    
Property Get Elev_Sunrise() As Variant
    Elev_Sunrise = m_Elev_Sunrise
End Property
    
Property Get Elev_SolarNoon() As Variant
    Elev_SolarNoon = m_Elev_SolarNoon
End Property
    
Property Get Elev_Sunset() As Variant
    Elev_Sunset = m_Elev_Sunset
End Property
    
' Azimuth ============================================================================
Property Get Azim_DateTime() As Variant
    Azim_DateTime = m_Azim_DateTime
End Property
    
Property Get Azim_Sunrise() As Variant
    Azim_Sunrise = m_Azim_Sunrise
End Property
    
Property Get Azim_SolarNoon() As Variant
    Azim_SolarNoon = m_Azim_SolarNoon
End Property
    
Property Get Azim_Sunset() As Variant
    Azim_Sunset = m_Azim_Sunset
End Property
    
Property Get SunlightMinutes() As Double
    SunlightMinutes = m_SunlightMinutes
End Property

Public Function Calculate()
    Dim I As Integer
    Dim UTC_Offset As Double
    Dim TZname_current As String
    Dim JD As Double
    Dim JC As Double
    Dim GeoMeanLongSun As Double
    Dim GeoMeanAnomSun As Double
    Dim EccentEarthOrb As Double
    Dim SunEquOfCenter As Double
    Dim SunTrueLong As Double
    Dim SunTrueAnom As Double
    Dim SunRadVector As Double
    Dim SunAppLong As Double
    Dim MeanOblEcliptic As Double
    Dim ObliqCorr As Double
    Dim SunRightAsc As Double
    Dim SunDeclin As Double
    Dim varY As Double
    Dim EqnOfTime As Double
    Dim HAsunrise As Double
    Dim SolarNoon As Double
    Dim Sunrise As Double
    Dim Sunset As Double
    Dim SunlightMinutes As Double
    Dim TimeOfDay As Double
    Dim TrueSolarTime As Double
    Dim HourAngle As Double
    Dim SolarZenith As Double
    Dim SolarElevation As Double
    Dim AtmosphericRefraction As Double
    Dim SolarAzimuth As Double
    Dim Times(1 To 4) As Double     ' DateTime, Sunrise, Solar Noon, Sunset
    Dim Elevation(1 To 4) As Double
    Dim Azimuth(1 To 4) As Double
            
    UTC_Offset = TZ.Offset_Hrs(m_DateTime, m_TZname, "UTC")
    TZname_current = TZ.CurrentTimeZoneDesignation(m_DateTime, m_TZname)
   
    JD = 2415018.5 + m_DateTime - UTC_Offset / 24
    JC = (JD - 2451545) / 36525
    GeoMeanLongSun = Mod2((280.46646 + JC * (36000.76983 + JC * 0.0003032)), 360)
    GeoMeanAnomSun = 357.52911 + JC * (35999.05029 - 0.0001537 * JC)
    EccentEarthOrb = 0.016708634 - JC * (0.000042037 + 0.0000001267 * JC)
    SunEquOfCenter = Sin(GeoMeanAnomSun * DtoR) * (1.914602 - JC * (0.004817 + 0.000014 * JC)) + _
                     Sin(2 * GeoMeanAnomSun * DtoR) * (0.019993 - 0.000101 * JC) + _
                     Sin(3 * GeoMeanAnomSun * DtoR) * 0.000289
    SunTrueLong = SunEquOfCenter + GeoMeanLongSun
    SunTrueAnom = SunEquOfCenter + GeoMeanAnomSun
    SunRadVector = (1.000001018 * (1 - EccentEarthOrb ^ 2)) / (1 + EccentEarthOrb * Cos(SunTrueAnom * DtoR))
    SunAppLong = SunTrueLong - 0.00569 - 0.00478 * Sin((125.04 - 1934.136 * JC) * DtoR)
    
    MeanOblEcliptic = (46.815 + JC * (0.00059 - JC * 0.001813))
    MeanOblEcliptic = 23 + (26 + ((21.448 - JC * MeanOblEcliptic)) / 60) / 60
    
    ObliqCorr = MeanOblEcliptic + 0.00256 * Cos((125.04 - 1934.136 * JC) * DtoR)
    SunRightAsc = RtoD * (ArcTan2(Cos(SunAppLong * DtoR), Cos(ObliqCorr * DtoR) * Sin(SunAppLong * DtoR)))
    SunDeclin = RtoD * (ArcSin(Sin(DtoR * ObliqCorr) * Sin(DtoR * SunAppLong)))
    varY = Tan(DtoR * (ObliqCorr / 2)) ^ 2
    EqnOfTime = 4 * RtoD * (varY * Sin(2 * DtoR * GeoMeanLongSun) - 2 * EccentEarthOrb * _
                Sin(DtoR * GeoMeanAnomSun) + 4 * EccentEarthOrb * varY * Sin(DtoR * GeoMeanAnomSun) * _
                Cos(2 * DtoR * GeoMeanLongSun) - 0.5 * varY * varY * Sin(4 * DtoR * GeoMeanLongSun) _
                - 1.25 * EccentEarthOrb * EccentEarthOrb * Sin(2 * DtoR * GeoMeanAnomSun))
    HAsunrise = RtoD * (ArcCos(Cos(DtoR * (90.833)) / (Cos(DtoR * m_Lat) * _
                Cos(DtoR * SunDeclin)) - Tan(DtoR * m_Lat) * Tan(DtoR * SunDeclin)))
    SolarNoon = (720 - 4 * m_Lon - EqnOfTime + UTC_Offset * 60) / 1440
    Sunrise = (SolarNoon * 1440 - HAsunrise * 4) / 1440
    Sunset = (SolarNoon * 1440 + HAsunrise * 4) / 1440
    SunlightMinutes = 8 * HAsunrise
    TimeOfDay = m_DateTime - Int(m_DateTime)
    
    Times(1) = Mod2((TimeOfDay * 1440 + EqnOfTime + 4 * m_Lon - 60 * UTC_Offset), 1440)
    Times(2) = Mod2((Sunrise * 1440 + EqnOfTime + 4 * m_Lon - 60 * UTC_Offset), 1440)
    Times(3) = Mod2((SolarNoon * 1440 + EqnOfTime + 4 * m_Lon - 60 * UTC_Offset), 1440)
    Times(4) = Mod2((Sunset * 1440 + EqnOfTime + 4 * m_Lon - 60 * UTC_Offset), 1440)
    
    ' Calculate Elevation and Azimuth for DateTime, Sunrise, SolarNoon, and Sunset
    For I = 1 To 4
        TrueSolarTime = Times(I)
        HourAngle = IIf(TrueSolarTime / 4 < 0, TrueSolarTime / 4 + 180, TrueSolarTime / 4 - 180)
        SolarZenith = RtoD * (ArcCos(Sin(DtoR * m_Lat) * Sin(DtoR * SunDeclin) + _
                      Cos(DtoR * m_Lat) * Cos(DtoR * SunDeclin) * Cos(DtoR * HourAngle)))
        SolarElevation = 90 - SolarZenith
        
        Select Case SolarElevation
            Case Is > 85
                AtmosphericRefraction = 0
            Case Is > 5
                AtmosphericRefraction = 58.1 / Tan(DtoR * SolarElevation) _
                                      - 0.07 / Tan(DtoR * SolarElevation) ^ 3 _
                                  + 0.000086 / Tan(DtoR * SolarElevation) ^ 5
            Case Is > -0.575
                AtmosphericRefraction = 1735 - 518.2 * SolarElevation + 103.4 * SolarElevation ^ 2 _
                                      - 12.79 * SolarElevation ^ 3 + 0.711 * SolarElevation ^ 4
            Case Else
                AtmosphericRefraction = -20.772 / Tan(DtoR * SolarElevation)
        End Select
        AtmosphericRefraction = AtmosphericRefraction / 3600
        SolarElevation = SolarElevation + AtmosphericRefraction
        SolarAzimuth = RtoD * (ArcCos(((Sin(DtoR * m_Lat) * Cos(DtoR * SolarZenith)) - Sin(DtoR * SunDeclin)) / _
                      (Cos(DtoR * m_Lat) * Sin(DtoR * SolarZenith))))
        SolarAzimuth = IIf(HourAngle > 0, Mod2(SolarAzimuth + 180, 360), Mod2(540 - SolarAzimuth, 360))
        Elevation(I) = SolarElevation
        Azimuth(I) = SolarAzimuth
    Next
   
    m_Time_Sunrise = Sunrise
    m_Time_SolarNoon = SolarNoon
    m_Time_Sunset = Sunset
    
    m_Elev_DateTime = Elevation(1)
    m_Elev_Sunrise = Elevation(2)
    m_Elev_SolarNoon = Elevation(3)
    m_Elev_Sunset = Elevation(4)
    
    m_Azim_DateTime = Azimuth(1)
    m_Azim_Sunrise = Azimuth(2)
    m_Azim_SolarNoon = Azimuth(3)
    m_Azim_Sunset = Azimuth(4)
    
    m_UTC_Offset = UTC_Offset
    m_TZname_current = TZname_current
    m_SunlightMinutes = SunlightMinutes
    
End Function

' Calculate Mod to multiple decimal places - VBA Mod function rounds to an integer
Private Function Mod2(X, Y) As Double
    Mod2 = X / Y
    Mod2 = (Mod2 - Int(Mod2)) * Y
End Function

' Function to determine the arc cosine
Private Function ArcCos(X As Double) As Double
    If Round(X, 12) = 1# Then
        ArcCos = 0
    ElseIf Round(X, 12) = -1# Then
        ArcCos = PI
    Else
        ArcCos = Atn(-X / Sqr(1 - X ^ 2)) + 2 * Atn(1)
    End If
End Function

' Function to determine the arc sine
' https://stackoverflow.com/questions/46631569/how-to-calculate-define-arcsin-function-in-vba
' The VBA Sgn function returns an integer (+1, 0 or -1), representing the arithmetic sign of a supplied number.
Private Function ArcSin(X As Double) As Double
    If Abs(X) = 1 Then
        ArcSin = Sgn(X) * PI_2
    Else
        ArcSin = Atn(X / Sqr(1 - X ^ 2))
    End If
End Function

' https://stackoverflow.com/questions/5117265/does-vba-have-an-atan2-function
Private Function ArcTan2(X As Double, Y As Double) As Double
    Select Case X
        Case Is > 0
            ArcTan2 = Atn(Y / X)
        Case Is < 0
            ArcTan2 = Atn(Y / X) + PI * Sgn(Y)
            If Y = 0 Then ArcTan2 = ArcTan2 + PI
        Case Is = 0
            ArcTan2 = PI_2 * Sgn(Y)
    End Select
End Function



