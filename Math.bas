Attribute VB_Name = "Math"
Option Explicit


Public Function GetJulianDate(Dt As MyDate, Tm As MyTime) As Double
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

  H = Tm.H / 24 + Tm.M / 1440 + Tm.s / 86400
  
  GetJulianDate = Int(365.25 * (Dt.YY + 4716)) + Int(30.6001 * (Dt.MM + 1)) + Dt.DD + H + B - 1524.5
End Function



Public Function GetSiderialTime(LocalDate As MyDate, LocalTimeUT As MyTime, Laenge As Double) As MyTime
' https://de.wikipedia.org/wiki/Sternzeit
' https://de.wikibooks.org/wiki/Astronomische_Berechnungen_für_Amateure/_Zeit/_Zeitrechnungen
' Welchen Wert hatte die mittlere Sternzeit in Berlin (Länge = +13.5°) am 25. Dezember 2007 um 20 h UT (entspricht 21 MEZ in Berlin)?
' Ergebnis: 3h 09m 48,3s

        Dim Time0hGMT As MyTime
        Dim LocalTimeSiderial As MyTime
        
        Dim t As Double
        Dim JD As Double
        Dim GMST_Zeit_s As Double
        Dim GMST_Zeit_h As Double
        Dim GMST_24 As Double
        Dim LokalTime As Double
        Dim GMSTindividualTime As Double
        Dim GMSTindividualOrt As Double
    
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
   
    'Geographische Länge berücksichtigen
    GMSTindividualOrt = GMSTindividualTime + Laenge / 15
  
    GMSTindividualOrt = CutTime(GMSTindividualOrt)
    
    GetSiderialTime = TimeDezToHMS(GMSTindividualOrt)
  
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
    
    locTDec = TimeDezimal
    
    TimeDezToHMS.TimeDec = locTDec
    
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


Public Function arcsin(x As Double) As Double
    arcsin = Atn(x / Sqr(-x * x + 1))
End Function







