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
  Dim H As Double

  If Dt.MM <= 2 Then
    Dt.YY = Dt.YY - 1
    Dt.MM = Dt.MM + 12
  End If
  
  'Korrektur für den julianischen Kalender
  a = Int(Dt.YY / 100)
  B = 2 - a + Int(a / 4)

  H = tm.H / 24 + tm.M / 1440 + tm.s / 86400
  
  GetJulianDate = Int(365.25 * (Dt.YY + 4716)) + Int(30.6001 * (Dt.MM + 1)) + Dt.DD + H + B - 1524.5
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
    Time0hGMT.H = 0
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
    
    
    TimeDezToHMS.H = Int(locTDec)
    locTDec = locTDec - TimeDezToHMS.H
    
    TimeDezToHMS.s = locTDec * 3600
    
    TimeDezToHMS.M = Int(TimeDezToHMS.s / 60)
    TimeDezToHMS.s = TimeDezToHMS.s - (TimeDezToHMS.M * 60)

End Function


Public Function TimeHMStoDez(TimeIn As MyTime) As MyTime
    TimeHMStoDez.TimeDec = TimeIn.H + TimeIn.M / 60 + TimeIn.s / 3600
    TimeHMStoDez.H = TimeIn.H
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
    Dim H As Double
    
    H = Deg * 24 / 360
    GradToTime = TimeDezToHMS(H)
End Function

Public Function RadToTime(Rad As Double) As MyTime
    Dim H As Double
    Dim Deg As Double
    
    Deg = CutAngle(RadToDeg(Rad))

    H = Deg * 24 / 360
    RadToTime = TimeDezToHMS(H)
End Function


Public Function TimeToRad(HMS As MyTime) As Double
    Dim tmp1 As MyTime
    Dim tmp2 As Double
    
    tmp1 = TimeHMStoDez(HMS)
    tmp2 = tmp1.TimeDec * 360 / 24
    TimeToRad = tmp2 / (180 / Pi)

End Function



Public Function StingsToDate(sTag As String, sMonat As String, sJahr As String, sStunden As String, sMinuten As String, sSekunden As String) As Date
    Dim iTag As Integer
    Dim iMonat As Integer
    Dim iJahr As Integer
    Dim iStunden As Integer
    Dim iMinuten As Integer
    Dim iSekunden As Integer
    
    iTag = Int(Zahl(sTag))
    iMonat = Int(Zahl(sMonat))
    iJahr = Int(Zahl(sJahr))
    iStunden = Int(Zahl(sStunden))
    iMinuten = Int(Zahl(sMinuten))
    iSekunden = Int(Zahl(sSekunden))
        
    If (iTag < 1) Or (iTag > 31) Then iTag = 1
    If (iMonat < 1) Or (iMonat > 12) Then iMonat = 1
    If (iStunden < 0) Or (iStunden > 23) Then iStunden = 0
    If (iMinuten < 0) Or (iMinuten > 59) Then iMinuten = 0
    If (iSekunden < 0) Or (iSekunden > 59) Then iSekunden = 0
  
    StingsToDate = iTag & "." & iMonat & "." & iJahr & " " & iStunden & ":" & iMinuten & ":" & iSekunden

End Function

' Calculates Hours [0h..24h] in radian [0..6,28]
' Example: 12h = 3,14
Public Function HourToRad(Hours As Double) As Double
    HourToRad = Hours * 2 * Pi / 24
End Function


Public Function DegToRad(Deg As Double) As Double
    DegToRad = Deg / (180 / Pi)
End Function


Public Function RadToDeg(Rad As Double) As Double
    RadToDeg = Rad * (180 / Pi)
End Function


Public Function arcsin(x As Double) As Double
    arcsin = Atn(x / Sqr(-x * x + 1))
End Function


Public Sub RA_DEC_to_AZ_ALT_radian(RA_Star_Rad As Double, DEC_Star_Rad As Double, Longitude As GeoCoord, Latitude As GeoCoord, LocalDateTime As Date, Az As Double, Alt As Double, LocalHourAngleRad As Double)
    ' matrix_method_rev_d.pdf Seite 15
    Dim SiderialTime As MyTime
    Dim SiderialTimeRad As Double
    Dim LongitudeDeg As Double
    Dim LongitudeRad As Double
    Dim LocalDate As MyDate
    Dim LocalTimeUT As MyTime

    LongitudeDeg = GeoToDez(Longitude)
    LongitudeRad = DegToRad(LongitudeDeg)
       
    'Calculate siderial time at Greenwich
    LocalDate.YY = Year(LocalDateTime)
    LocalDate.MM = Month(LocalDateTime)
    LocalDate.DD = Day(LocalDateTime)
    LocalTimeUT.H = Hour(LocalDateTime)
    LocalTimeUT.M = Minute(LocalDateTime)
    LocalTimeUT.s = Second(LocalDateTime)

    SiderialTime = GMST(LocalDate, LocalTimeUT)
    SiderialTimeRad = TimeToRad(SiderialTime)
    
    ' Calculate local hour angle
    LocalHourAngleRad = SiderialTimeRad - RA_Star_Rad
    LocalHourAngleRad = LocalHourAngleRad - LongitudeRad
    
    ' Calculate star position in rectangular equatorial coordinate system
    Dim LMN_Equatorial As Vector         ' Rectangular equatorial coordinate system
    Dim LMN_EquaMatrix(10, 10) As Double

    LMN_Equatorial = PolarKarthesisch(LocalHourAngleRad, DEC_Star_Rad)
   
    LMN_EquaMatrix(0, 0) = LMN_Equatorial.x
    LMN_EquaMatrix(1, 0) = LMN_Equatorial.Y
    LMN_EquaMatrix(2, 0) = LMN_Equatorial.z
   
    'Calculate star position in rectangular horizontal coordinate system
    Dim LMN_Horizontal As Vector                ' Rectangular horizontal coordinate system
    Dim LMN_HorizMatrix(10, 10) As Double
    Dim TransformationMatrix(10, 10) As Double  ' Transformation-Matrix from equatorial Coordinates to horizontal Coordinates
    Dim LatitudeDeg As Double                   ' Observer’s latitude
    Dim LatitudeRad As Double
   
    LatitudeDeg = GeoToDez(Latitude)
    LatitudeRad = DegToRad(LatitudeDeg)
   
    TransformationMatrix(0, 0) = Cos(LatitudeRad - Pi / 2)
    TransformationMatrix(0, 1) = 0
    TransformationMatrix(0, 2) = Sin(LatitudeRad - Pi / 2)
    TransformationMatrix(1, 0) = 0
    TransformationMatrix(1, 1) = 1
    TransformationMatrix(1, 2) = 0
    TransformationMatrix(2, 0) = -Sin(LatitudeRad - Pi / 2)
    TransformationMatrix(2, 1) = 0
    TransformationMatrix(2, 2) = Cos(LatitudeRad - Pi / 2)

    MatrixProduct TransformationMatrix, 3, 3, LMN_EquaMatrix, 3, 1, LMN_HorizMatrix

    Dim Lh As Double
    Dim Mh As Double
    Dim Nh As Double
    
    Lh = LMN_HorizMatrix(0, 0)                  ' Rectangular horizontal  coordinate system
    Mh = LMN_HorizMatrix(1, 0)
    Nh = LMN_HorizMatrix(2, 0)
    
    'Calculate Star position in Altazimuth horizontal coordinate system
    Dim sin_h As Double
'    When Lh >= 0, (-A) is in the 1st quadrant or the 4th quadrant.
'    When Lh < 0, (-A) is in the 2nd quadrant or the 3rd quadrant.
    Az = -Atn(Mh / Lh)
    If Lh < 0 Then
        Az = Az + Pi
    End If
    
    'geht möglicherweise einfacher: sin(h) = Nh
    sin_h = Cos(LatitudeRad) * Cos(LocalHourAngleRad) * Cos(DEC_Star_Rad) + Sin(LatitudeRad) * Sin(DEC_Star_Rad)
    Alt = arcsin(sin_h)

End Sub




Public Function VectorToAzAlt(V As Vector) As AzAlt
    VectorToAzAlt.Az = -Atn(V.Y / V.x)
'''    '    When V.x >= 0, (-A) is in the 1st quadrant or the 4th quadrant.
'''    '    When V.x < 0, (-A) is in the 2nd quadrant or the 3rd quadrant.
'''
'''    If V.x < 0 Then
'''        VectorToAzAlt.Az = VectorToAzAlt.Az + Pi
'''    End If


    VectorToAzAlt.Alt = arcsin(V.z)
End Function




Public Sub DerivateTeleskope(V As Vector)

End Sub


