Attribute VB_Name = "Math"
Option Explicit


Public Function GetJulianDate(Dt As MyDate, tm As MyTime) As Double
  'wenn Monat > 2 dann  Y = Jahr,   M = Monat
  '                  sonst Y = Jahr - 1, M = Monat + 12
  '
  '   D = Tag
  '
  '   H = Stunde / 24 + Minute / 1440 + Sekunde / 86400
  '
  '   wenn TT.MM.YYYY >= 15.10.1582
  '        dann gregorianischer Kalender: A = Int(Y/100), B = 2 - A + Int(A/4)
  '
  '   wenn TT.MM.YYYY <= 04.10.1582
  '        dann julianischer Kalender:                    B = 0
  '
  '   sonst Fehler: Das Datum zwischen dem 04.10.1582 und dem 15.10.1582 existiert nicht.
  '                 Auf den         04.10.1582 (julianischer Kalender) folgte
  '                 unmittelbar der 15.10.1582 (gregorianischer Kalender).
  '
  '   JD = Int(365,25*(Y+4716)) + Int(30,6001*(M+1)) + D + H + B - 1524,5
  '
  '   Das JD liefert bei 12:00:00 eine ganze Zahl. 4,5 / 24 = 0,1875
  '
  '   Beispiel: 25. März 2010, 16:30 UT (gregorianisch)    2.455.281,1875
  '  Dim A As Double
  Dim a As Double
  Dim B As Double
  Dim h As Double

  If Dt.MM <= 2 Then
    Dt.YY = Dt.YY - 1
    Dt.MM = Dt.MM + 12
  End If
  
  'Korrektur für den julianischen Kalender
  a = Int(Dt.YY / 100)
  B = 2 - a + Int(a / 4)

  h = tm.h / 24 + tm.M / 1440 + tm.s / 86400
  
  GetJulianDate = Int(365.25 * (Dt.YY + 4716)) + Int(30.6001 * (Dt.MM + 1)) + Dt.DD + h + B - 1524.5
End Function


Public Function GMST(LocalDate As MyDate, LocalTimeUT As MyTime) As MyTime
' https://de.wikipedia.org/wiki/Sternzeit
' https://de.wikibooks.org/wiki/Astronomische_Berechnungen_für_Amateure/_Zeit/_Zeitrechnungen

    Dim Time0hGMT As MyTime
    Dim LocalTimeSiderial As MyTime
    
    Dim t As Double
    Dim JD As Double
    Dim GMST_Zeit_s As Double
    Dim GMST_Zeit_h As Double
    Dim GMST_24 As Double
    Dim LokalTime As Double
    Dim GMSTindividualTime As Double

    '1.  Julianische Datum JD um 0h berechnen. Muß immer auf 0,5 enden
    Time0hGMT.h = 0
    Time0hGMT.M = 0
    Time0hGMT.s = 0
    JD = GetJulianDate(LocalDate, Time0hGMT)
    
    ' 2. Sternzeit in Greenwich berechnen
    ' Berechne die mittlere Sternzeit von Greenwich um 0 h UT zum gewünschten Datum.
    ' Addiere zum Ergebnis von 1) das Produkt t * 1.00273790935
    ' Der Faktor 1.002 737 909 35 berücksichtigt, dass die Sternzeit um so viel schneller abläuft als die Sonnenzeit.
    ' Das Resultat ist zum Schluss wieder auf [0; 24) zu normieren.
    ' Nächster Schritt außerhalb der Function:
    ' Soll die mittlere Sternzeit nicht für Greenwich, sondern für einen Ort der geografischen Länge L° ,
    ' addiere man zum Resultat L/15      (positiv gezählt nach Osten, negativ nach Westen)

    ' Anzahl der seit dem 1. Januar 2000, 12h UT1 (JD = 2451545.0 UT1) verstrichenen UT-Tage
    t = (JD - 2451545#) / 36525
  
    GMST_Zeit_s = 24110.54841 + 8640184.812866 * t + 0.093104 * t * t - 0.0000062 * t * t * t
    GMST_Zeit_h = GMST_Zeit_s / 3600
    GMST_24 = CutTime(GMST_Zeit_h)
  
    'Lokale Zeit auf siderische Zeit umgerechnet
    LocalTimeSiderial = TimeHMStoDez(LocalTimeUT)
    GMSTindividualTime = LocalTimeSiderial.TimeDec * 1.00273790935 + GMST_24
    GMSTindividualTime = CutTime(GMSTindividualTime)
    GMST = TimeDezToHMS(GMSTindividualTime)
    
End Function


Public Function CutTime(ByVal Hours As Double) As Double
  Dim HoursInt As Long
  HoursInt = Int(Hours / 24)
  CutTime = Hours - (HoursInt * 24)
End Function


Public Function CutAngle(ByVal Angle As Double) As Double
  Dim AngleInt As Long
  AngleInt = Int(Angle / 360)
  CutAngle = Angle - (AngleInt * 360)
End Function


Public Function TimeDezToHMS(TimeDezimal As Double) As MyTime
    Dim locTDec As Double
    Dim Negative As Boolean
    
    TimeDezToHMS.TimeDec = TimeDezimal
    
    If TimeDezimal < 0 Then
        locTDec = -TimeDezimal
        Negative = True
    Else
        locTDec = TimeDezimal
        Negative = False
    End If
    
    
    TimeDezToHMS.h = Int(locTDec)
    locTDec = locTDec - TimeDezToHMS.h
    
    TimeDezToHMS.s = locTDec * 3600
    
    TimeDezToHMS.M = Int(TimeDezToHMS.s / 60)
    TimeDezToHMS.s = TimeDezToHMS.s - (TimeDezToHMS.M * 60)

End Function


Public Function TimeHMStoDez(TimeIn As MyTime) As MyTime
    TimeHMStoDez.TimeDec = TimeIn.h + TimeIn.M / 60 + TimeIn.s / 3600
    TimeHMStoDez.h = TimeIn.h
    TimeHMStoDez.M = TimeIn.M
    TimeHMStoDez.s = TimeIn.s
    
End Function

Public Function GeoToDez(Coord As GeoCoord) As Double
    Dim s As String
    
    s = Mid(Coord.Sign, 1, 1)
    GeoToDez = Coord.Deg + Coord.Min / 60 + Coord.Sec / 3600
    If s = "o" Or s = "O" Or s = "E" Or s = "-" Or s = "s" Or s = "S" Or s = "e" Then
        GeoToDez = -GeoToDez
    ElseIf s = "w" Or s = "W" Or s = "n" Or s = "N" Or s = "+" Then
    
    End If
End Function

Public Function GradToTime(Deg As Double) As MyTime
    Dim h As Double
    
    h = Deg * 24 / 360
    GradToTime = TimeDezToHMS(h)
End Function




Public Function arcsin(x As Double) As Double
    arcsin = Atn(x / Sqr(-x * x + 1))
End Function






Public Sub RA_DEC_to_AZ_ALT(RA_Star As MyTime, DEC_Star As MyTime, Longitude As GeoCoord, Latitude As GeoCoord, LocalTimeUT As MyTime, LocalDate As MyDate, AZ As Double, ALT As Double, HourAngle As MyTime)
    ' matrix_method_rev_d.pdf Seite 15
    Dim SiderialTime As MyTime
    Dim LocalHourAngleHour As Double    'Local hour angle in hour (decimal)
    Dim LocalHourAngleDeg As Double     'Local hour angle in degree
    Dim LocalHourAngleRad As Double     'Local hour angle in radian
    Dim LongitudeDeg As Double
    
    LongitudeDeg = GeoToDez(Longitude)
       
    'Calculate siderial time at Greenwich
    SiderialTime = GMST(LocalDate, LocalTimeUT)
    
    ' Calculate local hour angle
    RA_Star = TimeHMStoDez(RA_Star)
    LocalHourAngleHour = SiderialTime.TimeDec - RA_Star.TimeDec
    LocalHourAngleDeg = LocalHourAngleHour * 15 - LongitudeDeg
    LocalHourAngleRad = LocalHourAngleDeg / (180 / Pi)
    
    HourAngle = GradToTime(CutAngle(LocalHourAngleDeg))
    
    ' Calculate star position in rectangular equatorial coordinate system
    Dim LMN_Equatorial As Vector         ' Rectangular equatorial coordinate system
    Dim LMN_EquaMatrix(10, 10) As Double
    
    Dim DeclinationDeg As Double
    Dim DeclinationRad As Double
      
    DEC_Star = TimeHMStoDez(DEC_Star)
    DeclinationDeg = DEC_Star.TimeDec
    DeclinationRad = DeclinationDeg / (180 / Pi)
    LMN_Equatorial = PolarKarthesisch(LocalHourAngleRad, DeclinationRad)
   
    LMN_EquaMatrix(0, 0) = LMN_Equatorial.x
    LMN_EquaMatrix(1, 0) = LMN_Equatorial.Y
    LMN_EquaMatrix(2, 0) = LMN_Equatorial.z
   
    'Calculate star position in rectangular horizontal coordinate system
    Dim LMN_Horizontal As Vector                ' Rectangular horizontal coordinate system
    Dim LMN_HorizMatrix(10, 10) As Double
    Dim TransformationMatrix(10, 10) As Double  ' Transformation-Matrix from equatorial Coordinates to horizontal Coordinates
    Dim LatitudeDeg As Double                  ' Observer’s latitude
    Dim LatitudeRad As Double
    Dim Phi As Double                           ' Observer’s latitude
    
    LatitudeDeg = GeoToDez(Latitude)
    LatitudeRad = LatitudeDeg / (180 / Pi)
   
    Phi = LatitudeRad
    TransformationMatrix(0, 0) = Cos(Phi - Pi / 2)
    TransformationMatrix(0, 1) = 0
    TransformationMatrix(0, 2) = Sin(Phi - Pi / 2)
    TransformationMatrix(1, 0) = 0
    TransformationMatrix(1, 1) = 1
    TransformationMatrix(1, 2) = 0
    TransformationMatrix(2, 0) = -Sin(Phi - Pi / 2)
    TransformationMatrix(2, 1) = 0
    TransformationMatrix(2, 2) = Cos(Phi - Pi / 2)

    MatrixProduct TransformationMatrix, 3, 3, LMN_EquaMatrix, 3, 1, LMN_HorizMatrix

    Dim Lh As Double
    Dim Mh As Double
    Dim Nh As Double
    
    Lh = LMN_HorizMatrix(0, 0)                  ' Rectangular horizontal  coordinate system
    Mh = LMN_HorizMatrix(1, 0)
    Nh = LMN_HorizMatrix(2, 0)
    
    
    'Calculate Star position in Altazimuth horizontal coordinate system
    Dim AzRad As Double         'Azimuth in radian
    Dim AzDeg As Double         'Azimuth in degree
    Dim AltRad As Double        'Altitude in radian
    Dim AltDeg As Double        'Altitude in degree
    Dim sin_h As Double
    
    AzRad = -Atn(Mh / Lh)
    AzDeg = AzRad / (Pi / 180)
    
    sin_h = Cos(Phi) * Cos(LocalHourAngleRad) * Cos(DeclinationRad) + Sin(Phi) * Sin(DeclinationRad)
    AltRad = arcsin(sin_h)
    AltDeg = AltRad / (Pi / 180)
    
    AZ = AzDeg
    ALT = AltDeg

End Sub





