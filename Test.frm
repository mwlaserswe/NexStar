VERSION 5.00
Begin VB.Form Test 
   Caption         =   "Test"
   ClientHeight    =   7320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Command2"
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   4080
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   4920
      TabIndex        =   14
      Top             =   1080
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Demo Stern"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton C_TestSiderialTime 
      Caption         =   "Siderial Time"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label L_AltStar 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label L_AzStar 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Azimuth"
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Altitude"
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label L_HourAngle 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Hour Angle"
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label L_SiderialTime 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Siderial Time"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.Label L_SiderialTimeHMS 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   5160
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Test siderial time
' https://de.wikibooks.org/wiki/Astronomische_Berechnungen_f%C3%BCr_Amateure/_Zeit/_Zeitrechnungen
' Welchen Wert hatte die mittlere Sternzeit?
' Berlin (Länge = +13.5°) am 25. Dezember 2007 um 20 h UT (entspricht 21 MEZ in Berlin)?
' Ergebnis: 3,1634161794371 = 3h 09m 48,3s
Private Sub C_TestSiderialTime_Click()

    Dim DemoDate As MyDate
    Dim DemoTime As MyTime
    Dim SiderialTime As MyTime
    Dim SiderialTimeGreenwich As MyTime
    Dim s As String

    DemoDate.YY = 2007
    DemoDate.MM = 12
    DemoDate.DD = 25
    DemoTime.H = 20
    DemoTime.M = 0
    DemoTime.s = 0

    SiderialTimeGreenwich = GMST(DemoDate, DemoTime)
    SiderialTime = TimeDezToHMS(SiderialTimeGreenwich.TimeDec + 13.5 / 15)
    L_SiderialTime = SiderialTime.TimeDec
    L_SiderialTimeHMS = SiderialTime.H & ":" & SiderialTime.M & ":" & Format(SiderialTime.s, "00.00")

End Sub

Private Sub Command1_Click()
    Dim ut As Date
    Dim AnyDateTime As Date

    Dim tst1 As Integer
    Dim tst2 As Integer
    Dim tst3 As Integer
    Dim tst4 As Integer
    Dim tst5 As Integer
    Dim tst6 As Integer

    ut = UtcTime(Now)

    AnyDateTime = "18.2.2019 1:0:0"
    ut = UtcTime(AnyDateTime)

    tst1 = Day(ut)
    tst2 = Month(ut)
    tst3 = Year(ut)
    tst4 = Hour(ut)
    tst5 = Minute(ut)
    tst6 = Second(ut)

    'Achtung: "2019.11.4 1:0:00" liefert nur "4.11.2019"

End Sub

Private Sub Command2_Click()
    TestJulianischesDatum.Show
End Sub

Private Sub Command3_Click()
    Dim a As String
    Dim B As String
    Dim erg As Long

    a = SetNexStarPosition(1234567)

    B = Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H11) & Chr$(&H24) & Chr$(&H80)
'    b = Chr$(&H0) & Chr$(&H3) & Chr$(&HE8)

    erg = GetNexStarPosition(a)

End Sub


Private Sub Command4_Click()
    Dim Jetzt As String
    Dim Lont As MyTime

    ' Datensatz Saturn Demo aus dem Sript
    Dim SaturnDateTime As Date
    SaturnDateTime = "13.11.1978 4:34:0"

    Dim RA_Saturn As MyTime
    RA_Saturn.H = 10
    RA_Saturn.M = 57              '57
    RA_Saturn.s = 35.681

    Dim DEC_Saturn As GeoCoord
    DEC_Saturn.Deg = 8
    DEC_Saturn.Min = 25
    DEC_Saturn.Sec = 58.1
    DEC_Saturn.Sign = "+"

    Lont = TimeDezToHMS(4.35808335) '  -4.358°              ' Observer’s longitude

    Dim Longitude As GeoCoord
    Longitude.Deg = Lont.H
    Longitude.Min = Lont.M
    Longitude.Sec = Lont.s
    Longitude.Sign = "E"

    Dim Latitude As GeoCoord    '  50°47'55''                 ' Observer’s latitude
    Latitude.Deg = 50
    Latitude.Min = 47
    Latitude.Sec = 55
    Latitude.Sign = "N"

    Dim Az As Double
    Dim Alt As Double
    Dim LocalHourAngleRad As Double
    Dim HourAngle As MyTime

    Dim RA_Saturn_Rad As Double
    RA_Saturn_Rad = TimeToRad(RA_Saturn)
    Dim DEC_Saturn_Rad As Double
    DEC_Saturn_Rad = DegToRad(GeoToDez(DEC_Saturn))

    RA_DEC_to_AZ_ALT_radian RA_Saturn_Rad, DEC_Saturn_Rad, Longitude, Latitude, SaturnDateTime, Az, Alt, LocalHourAngleRad

    If Mainform.O_OrientationNorth.Value Then Az = Az + Pi
    L_AzStar = CutAngle(RadToDeg(Az))
    L_AltStar = RadToDeg(Alt)

    HourAngle = RadToTime(LocalHourAngleRad)
    L_HourAngle = HourAngle.H & ":" & HourAngle.M & ":" & Format(HourAngle.s, "00.00")


 ' Capella Kassel
'''    Dim CapellaDateTime As Date
'''    CapellaDateTime = "2.2.2019 19:00:00"
'''
'''    Dim RA_Capella As MyTime
'''    RA_Capella.H = 5
'''    RA_Capella.M = 18
'''    RA_Capella.s = 6
'''
'''    Dim DEC_Capella As GeoCoord
'''    DEC_Capella.Deg = 46
'''    DEC_Capella.Min = 1
'''    DEC_Capella.Sec = 0
'''    DEC_Capella.Sign = "+"
'''
'''    Dim Longitude As GeoCoord                     ' Observer’s longitude
'''    Longitude.Deg = 9
'''    Longitude.Min = 18
'''    Longitude.Sec = 3
'''    Longitude.Sign = "E"
'''
'''    Dim Latitude As GeoCoord                     ' Observer’s latitude
'''    Latitude.Deg = 51
'''    Latitude.Min = 11
'''    Latitude.Sec = 27
'''    Latitude.Sign = "N"
'''
'''    Dim Az As Double
'''    Dim Alt As Double
'''    Dim LocalHourAngleRad As Double
'''    Dim HourAngle As MyTime
'''
'''    Dim RA_Capella_Rad As Double
'''    RA_Capella_Rad = TimeToRad(RA_Capella)
'''    Dim DEC_Capella_Rad As Double
'''    DEC_Capella_Rad = DegToRad(GeoToDez(DEC_Capella))
'''
'''    RA_DEC_to_AZ_ALT_radian RA_Capella_Rad, DEC_Capella_Rad, Longitude, Latitude, CapellaDateTime, Az, Alt, LocalHourAngleRad
'''
'''    If O_OrientationNorth.Value Then Az = Az + Pi
'''    L_AzStar = CutAngle(RadToDeg(Az))
'''    L_AltStar = RadToDeg(Alt)
'''
'''    HourAngle = RadToTime(LocalHourAngleRad)
'''    L_HourAngle = HourAngle.H & ":" & HourAngle.M & ":" & Format(HourAngle.s, "00.00")

 ' Deneb München
'''    Dim DenebDateTime As Date
'''    DenebDateTime = "2.2.2019 19:00:00"
'''
'''    Dim RA_Deneb As MyTime
'''    RA_Deneb.H = 20
'''    RA_Deneb.M = 42
'''    RA_Deneb.s = 4
'''
'''    Dim DEC_Deneb As GeoCoord
'''    DEC_Deneb.Deg = 45
'''    DEC_Deneb.Min = 21
'''    DEC_Deneb.Sec = 0
'''    DEC_Deneb.Sign = "+"
'''
'''    Dim Longitude As GeoCoord                     ' Observer’s longitude
'''    Longitude.Deg = 11
'''    Longitude.Min = 34
'''    Longitude.Sec = 0
'''    Longitude.Sign = "E"
'''
'''    Dim Latitude As GeoCoord                     ' Observer’s latitude
'''    Latitude.Deg = 48
'''    Latitude.Min = 8
'''    Latitude.Sec = 0
'''    Latitude.Sign = "N"
'''
'''    Dim Az As Double
'''    Dim Alt As Double
'''    Dim LocalHourAngleRad As Double
'''    Dim HourAngle As MyTime
'''
'''    Dim RA_Deneb_Rad As Double
'''    RA_Deneb_Rad = TimeToRad(RA_Deneb)
'''    Dim DEC_Deneb_Rad As Double
'''    DEC_Deneb_Rad = DegToRad(GeoToDez(DEC_Deneb))
'''
'''     RA_DEC_to_AZ_ALT_radian RA_Deneb_Rad, DEC_Deneb_Rad, Longitude, Latitude, DenebDateTime, Az, Alt, LocalHourAngleRad
'''
'''    If O_OrientationNorth.value Then Az = Az + Pi
'''    L_AzStar = CutAngle(RadToDeg(Az))
'''    L_AltStar = RadToDeg(Alt)
'''
'''    HourAngle = RadToTime(LocalHourAngleRad)
'''    L_HourAngle = HourAngle.H & ":" & HourAngle.M & ":" & Format(HourAngle.s, "00.00")



End Sub
Private Sub Command5_Click()
    ' matrix_method_rev_d.pdf Seite 37
    Dim tmp As Vector

    Dim tst As MyTime
    Dim InitTimerad As Double
    Dim ObservTime1Rad As Double
    Dim ObservTime2Rad As Double
    Dim RA1Rad As Double
    Dim RA2Rad As Double
    Dim DEC1Rad As Double
    Dim DEC2Rad As Double
    Dim TelHorizAngle1 As Double
    Dim TelHorizAngle2 As Double
    Dim TelElevAngle1 As Double
    Dim TelElevAngle2 As Double


    tst.H = 21
    tst.M = 0
    tst.s = 0
    InitTimerad = TimeToRad(tst)

    tst.H = 21
    tst.M = 27
    tst.s = 56
    ObservTime1Rad = TimeToRad(tst)

    tst.H = 0
    tst.M = 7
    tst.s = 54
    RA1Rad = TimeToRad(tst)
    DEC1Rad = DegToRad(29.038)
    TelHorizAngle1 = DegToRad(99.25)
    TelElevAngle1 = DegToRad(83.87)

    tst.H = 21
    tst.M = 37
    tst.s = 2
    ObservTime2Rad = TimeToRad(tst)

    tst.H = 2
    tst.M = 21
    tst.s = 45
    RA2Rad = TimeToRad(tst)
    DEC2Rad = DegToRad(89.222)
    TelHorizAngle2 = DegToRad(310.98)
    TelElevAngle2 = DegToRad(35.04)


    Dim lmn_Tel_1 As Vector     ' Telescope coordinates
    Dim lmn_Tel_2 As Vector
    Dim lmn_Tel_3 As Vector
    Dim LMN_Equ_1 As Vector
    Dim LMN_Equ_2 As Vector
    Dim LMN_Equ_3 As Vector
    Dim k As Double         ' Umrechnung Sonnenzeit in siderische Zeit 1.00273790935
    k = 1.00273790935

    'Equation (5.4-5)
    lmn_Tel_1.x = Cos(TelElevAngle1) * Cos(TelHorizAngle1)
    lmn_Tel_1.Y = Cos(TelElevAngle1) * Sin(TelHorizAngle1)
    lmn_Tel_1.z = Sin(TelElevAngle1)

    'Equation (5.4-6)
    LMN_Equ_1.x = Cos(DEC1Rad) * Cos(RA1Rad - k * (ObservTime1Rad - InitTimerad))
    LMN_Equ_1.Y = Cos(DEC1Rad) * Sin(RA1Rad - k * (ObservTime1Rad - InitTimerad))
    LMN_Equ_1.z = Sin(DEC1Rad)

    'Equation (5.4-7)
    lmn_Tel_2.x = Cos(TelElevAngle2) * Cos(TelHorizAngle2)
    lmn_Tel_2.Y = Cos(TelElevAngle2) * Sin(TelHorizAngle2)
    lmn_Tel_2.z = Sin(TelElevAngle2)

    'Equation (5.4-8)
    LMN_Equ_2.x = Cos(DEC2Rad) * Cos(RA2Rad - k * (ObservTime2Rad - InitTimerad))
    LMN_Equ_2.Y = Cos(DEC2Rad) * Sin(RA2Rad - k * (ObservTime2Rad - InitTimerad))
    LMN_Equ_2.z = Sin(DEC2Rad)

    Dim V1_cross_V2 As Vector
    Dim Len_V1_cross_V2 As Double

    'Equation (5.4-13)
    V1_cross_V2 = CrossProduct(lmn_Tel_1, lmn_Tel_2)
    Len_V1_cross_V2 = LenghtVector(V1_cross_V2)
    lmn_Tel_3 = ScalarProduct((1 / Len_V1_cross_V2), V1_cross_V2)

    'Equation (5.4-14)
    V1_cross_V2 = CrossProduct(LMN_Equ_1, LMN_Equ_2)
    Len_V1_cross_V2 = LenghtVector(V1_cross_V2)
    LMN_Equ_3 = ScalarProduct((1 / Len_V1_cross_V2), V1_cross_V2)


    'From equation(5.4 - 11)
    Dim LMN_Equ__Matrix(10, 10) As Double
    Dim LMN_Equ__MatrixInvers(10, 10) As Double
    Dim lmn_Tel__Matrix(10, 10) As Double
    Dim TransformationMatrix(10, 10) As Double

    LMN_Equ__Matrix(0, 0) = LMN_Equ_1.x
    LMN_Equ__Matrix(0, 1) = LMN_Equ_2.x
    LMN_Equ__Matrix(0, 2) = LMN_Equ_3.x
    LMN_Equ__Matrix(1, 0) = LMN_Equ_1.Y
    LMN_Equ__Matrix(1, 1) = LMN_Equ_2.Y
    LMN_Equ__Matrix(1, 2) = LMN_Equ_3.Y
    LMN_Equ__Matrix(2, 0) = LMN_Equ_1.z
    LMN_Equ__Matrix(2, 1) = LMN_Equ_2.z
    LMN_Equ__Matrix(2, 2) = LMN_Equ_3.z

    Calculate_Inverse 3, LMN_Equ__Matrix, LMN_Equ__MatrixInvers

    lmn_Tel__Matrix(0, 0) = lmn_Tel_1.x
    lmn_Tel__Matrix(0, 1) = lmn_Tel_2.x
    lmn_Tel__Matrix(0, 2) = lmn_Tel_3.x
    lmn_Tel__Matrix(1, 0) = lmn_Tel_1.Y
    lmn_Tel__Matrix(1, 1) = lmn_Tel_2.Y
    lmn_Tel__Matrix(1, 2) = lmn_Tel_3.Y
    lmn_Tel__Matrix(2, 0) = lmn_Tel_1.z
    lmn_Tel__Matrix(2, 1) = lmn_Tel_2.z
    lmn_Tel__Matrix(2, 2) = lmn_Tel_3.z

    '==================================================================================================
    'This is the TransformationMatrix which transforms a vector from eqatorial to telescope coordinates
    '==================================================================================================
    MatrixProduct lmn_Tel__Matrix, 3, 3, LMN_Equ__MatrixInvers, 3, 3, TransformationMatrix



    '=================================
    ' Example: Beta Cet: Deneb Kaitos
    '=================================

    Dim RA_BetaCet As MyTime
    Dim DEC_BetaCet As Double
    Dim AimTime As MyTime
    Dim RA_BetaCetRad As Double
    Dim DEC_BetaCetRad As Double
    Dim AimTimeRad As Double

    ' if you want to aim the telescope at Beta Cet (RA = 0h43m07s, DEC = -18.038°) at 21h52m12s
    RA_BetaCet.H = 0: RA_BetaCet.M = 43: RA_BetaCet.s = 7
    RA_BetaCetRad = TimeToRad(RA_BetaCet)
    DEC_BetaCetRad = DegToRad(-18.038)
    AimTime.H = 21: AimTime.M = 52: AimTime.s = 12
    AimTimeRad = TimeToRad(AimTime)

    'LMN_Equ_1: Vector points to Beta Cet in equatorial coordinats
    LMN_Equ_1.x = Cos(DEC_BetaCetRad) * Cos(RA_BetaCetRad - k * (AimTimeRad - InitTimerad))
    LMN_Equ_1.Y = Cos(DEC_BetaCetRad) * Sin(RA_BetaCetRad - k * (AimTimeRad - InitTimerad))
    LMN_Equ_1.z = Sin(DEC_BetaCetRad)



    LMN_Equ__Matrix(0, 0) = LMN_Equ_1.x
    LMN_Equ__Matrix(1, 0) = LMN_Equ_1.Y
    LMN_Equ__Matrix(2, 0) = LMN_Equ_1.z


    MatrixProduct TransformationMatrix, 3, 3, LMN_Equ__Matrix, 3, 1, lmn_Tel__Matrix

    'lmn_Tel__Matrix: Vector points to Beta Cet in equatorial coordinats

    lmn_Tel_1.x = lmn_Tel__Matrix(0, 0)
    lmn_Tel_1.Y = lmn_Tel__Matrix(1, 0)
    lmn_Tel_1.z = lmn_Tel__Matrix(2, 0)

    Dim AzAlt_BetaCet As AzAlt
    Dim Az_BetaCetRad As Double
    Dim Alt_BetaCetRad As Double
    Dim Az_BetaCet As Double
    Dim Alt_BetaCet As Double

    AzAlt_BetaCet = VectorToAzAlt(lmn_Tel_1)
    Az_BetaCetRad = AzAlt_BetaCet.Az
    Alt_BetaCetRad = AzAlt_BetaCet.Alt

    ' !!! hier muß möglicherweise noch 180° addiert werden !!!
    Az_BetaCet = RadToDeg(Az_BetaCetRad)

    Alt_BetaCet = RadToDeg(Alt_BetaCetRad)

End Sub

Private Sub Command7_Click()
    Dim i As Long
    
    Dim dummy As Double
    Label1 = "1"
        For i = 1 To 10000000
            dummy = dummy * Pi
        Next i
    Label1 = "2"
        For i = 1 To 10000000
            dummy = dummy * Pi
        Next i
    Label1 = "3"
End Sub
