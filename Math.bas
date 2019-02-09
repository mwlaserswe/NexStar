Attribute VB_Name = "Math"
Public Function GetJulianDate(ByVal Y As Double, ByVal M As Double, ByVal D As Double, ByVal Hour As Double, ByVal Minute As Double, ByVal Second As Double) As Double
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
  Dim B As Double
  Dim H As Double

  If M <= 2 Then
    Y = Y - 1
    M = M + 12
  End If
  
  'Korrektur für den julianischen Kalender
  a = Int(Y / 100)
  B = 2 - a + Int(a / 4)

  H = Hour / 24 + Minute / 1440 + Second / 86400
  
  GetJulianDate = Int(365.25 * (Y + 4716)) + Int(30.6001 * (M + 1)) + D + H + B - 1524.5
End Function




Public Sub GetSiderialTime(Y As Double, M As Double, D As Double, Stunde As Double, Minuten As Double, Sekunde As Double, Zeitzone As Double, Laenge As Double, SiderialTime As Double)
' https://de.wikipedia.org/wiki/Sternzeit
' https://de.wikibooks.org/wiki/Astronomische_Berechnungen_für_Amateure/_Zeit/_Zeitrechnungen
' Welchen Wert hatte die mittlere Sternzeit in Berlin (Länge = +13.5°) am 25. Dezember 2007 um 20 h UT (entspricht 21 MEZ in Berlin)?
' Ergebnis: 3h 09m 48,3s

    Dim t As Double
    Dim JD As Double
    Dim GMST_Zeit_s As Double
    Dim GMST_Zeit_h As Double
    Dim GMST_24 As Double
    Dim LokalTime As Double
    Dim GMSTindividualTime As Double
    Dim GMSTindividualOrt As Double
    
    '1.  Julianische Datum JD um 0h berechnen. Muß immer auf 0,5 enden
     JD = GetJulianDate(Y, M, D, 0, 0, 0)
    
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
    LokalTime = (Stunde + Minuten / 60 + Sekunde / 3600) * 1.00273790935
 
    GMSTindividualTime = GMST_24 + LokalTime
  
    GMSTindividualTime = CutTime(GMSTindividualTime)
   
    'Geographische Länge berücksichtigen
    GMSTindividualOrt = GMSTindividualTime + Laenge / 15
  
    GMSTindividualOrt = CutTime(GMSTindividualOrt)
    SiderialTime = GMSTindividualOrt
  
End Sub



Public Sub ZeitDezToHMS(ByVal TimeDezimal As Double, Stunden As Double, Minuten As Double, Sekunden As Double)
  Stunden = Int(TimeDezimal)
  TimeDezimal = TimeDezimal - Stunden

  Sekunden = TimeDezimal * 3600

  Minuten = Int(Sekunden / 60)
  Sekunden = Int(Sekunden - (Minuten * 60))


End Sub
