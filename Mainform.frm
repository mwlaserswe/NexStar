VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Mainform 
   Caption         =   "Form1"
   ClientHeight    =   11235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   11235
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   195
      Left            =   600
      TabIndex        =   35
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   195
      Left            =   600
      TabIndex        =   34
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Timer Tim_Tracking 
      Interval        =   1000
      Left            =   7800
      Top             =   5520
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Demo Stern"
      Height          =   255
      Left            =   600
      TabIndex        =   30
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton C_TestSiderialTime 
      Caption         =   "Siderial Time"
      Height          =   255
      Left            =   600
      TabIndex        =   22
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Test Sternzeit"
      Height          =   495
      Left            =   600
      TabIndex        =   21
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   600
      TabIndex        =   20
      Top             =   3000
      Width           =   855
   End
   Begin VB.OptionButton O_OrientationSouth 
      Caption         =   "South"
      Height          =   195
      Left            =   480
      TabIndex        =   19
      Top             =   1440
      Width           =   1215
   End
   Begin VB.OptionButton O_OrientationNorth 
      Caption         =   "North"
      Height          =   195
      Left            =   480
      TabIndex        =   18
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Timer Tim_Simulation 
      Interval        =   100
      Left            =   7200
      Top             =   5280
   End
   Begin VB.Timer Tim_DisplayUpdate 
      Interval        =   250
      Left            =   6840
      Top             =   4680
   End
   Begin VB.CommandButton C_SetAzAlt 
      Caption         =   "Set Az Alt"
      Height          =   495
      Left            =   2880
      TabIndex        =   15
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox T_Alt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4320
      TabIndex        =   14
      Text            =   "0"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox T_Az 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4320
      TabIndex        =   13
      Text            =   "0"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.VScrollBar VS_ManualSkewingSpeed 
      Height          =   2295
      LargeChange     =   10
      Left            =   3960
      Max             =   0
      Min             =   100
      TabIndex        =   12
      Top             =   2520
      Value           =   100
      Width           =   255
   End
   Begin VB.CommandButton C_Le 
      Caption         =   "<"
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton C_Dn 
      Caption         =   "V"
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton C_Ri 
      Caption         =   ">"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton C_Up 
      Caption         =   "^"
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   3000
      Width           =   375
   End
   Begin VB.ListBox AlignmentStarList 
      Height          =   9615
      Left            =   5880
      TabIndex        =   7
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton C_SetEncoder 
      Caption         =   "Set Encoder"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton C_GetAlt 
      Caption         =   "Get Alt"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton C_GetAz 
      Caption         =   "Get Az"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin MSCommLib.MSComm NexStarComm 
      Left            =   7560
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   0   'False
      InputLen        =   1
      RThreshold      =   1
      BaudRate        =   4800
      InputMode       =   1
   End
   Begin VB.Label L_CurrentStar 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   36
      Top             =   10200
      Width           =   2775
   End
   Begin VB.Label L_UT 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   2400
      TabIndex        =   33
      Top             =   8040
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Hour Angle"
      Height          =   255
      Left            =   2280
      TabIndex        =   32
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label L_HourAngle 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   3120
      TabIndex        =   31
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Altitude"
      Height          =   255
      Left            =   2280
      TabIndex        =   29
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Azimuth"
      Height          =   255
      Left            =   2280
      TabIndex        =   28
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label L_AzStar 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   3120
      TabIndex        =   27
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label L_AltStar 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   3120
      TabIndex        =   26
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label L_SiderialTimeHMS 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   5520
      TabIndex        =   25
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Siderial Time"
      Height          =   255
      Left            =   2520
      TabIndex        =   24
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label L_SiderialTime 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   3720
      TabIndex        =   23
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label L_TelDegAlt 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   375
      Left            =   4800
      TabIndex        =   17
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label L_TelDegAz 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   375
      Left            =   4800
      TabIndex        =   16
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label L_Az 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label L_Alt 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
End
Attribute VB_Name = "Mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NexStarPortNr As Long
Dim NexStarBaudrate As Long
Dim Command As Long
Dim CommandCnt As Long
Dim InputBufferAz As String
Dim InputbufferAlt As String
Dim NexStarAz As String
Dim NexStarAlt As String
Dim TelIncrAz As Long
Dim TelIncrAlt As Long
Dim TelDegAz As Double
Dim TelDegAlt As Double

Dim ManualSkewingSpeed As Long

'Simulation
Dim SimIncrAz As Long
Dim SimIncrAlt As Long
Dim SimBntUp As Boolean
Dim SimBntDn As Boolean
Dim SimBntLe As Boolean
Dim SimBntRi As Boolean
Dim SimGotoAzAltActive As Boolean
Dim SimGotoAz As Long
Dim SimGotoAlt As Long






Private Sub AlignmentStarList_Click()
  Dim idx As Long
  
  idx = AlignmentStarList.ListIndex
  
  L_CurrentStar = AlignmentStarArray(idx).ProperName
  

End Sub

Private Sub C_GetAz_Click()
    If SimOffline Then
        TelIncrAz = SimIncrAz
                L_Az = TelIncrAz
                TelDegAz = TelIncrAz * 360 / EncoderResolution
                L_TelDegAz = Format(TelDegAz, "0.0000")
    Else
        NexStarComm.Output = Chr$(&H1)
        NexStarAz = ""
        Command = 1
    End If
End Sub

Private Sub C_GetAlt_Click()
    If SimOffline Then
        TelIncrAlt = SimIncrAlt
                 L_Alt = TelIncrAlt
                TelDegAlt = TelIncrAlt * 360 / EncoderResolution
                L_TelDegAlt = Format(TelDegAlt, "0.0000")
   Else
        NexStarComm.Output = Chr$(&H15)
        NexStarAlt = ""
        Command = 21
    End If
End Sub




Private Sub C_Up_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If SimOffline Then
        SimBntUp = True
    Else
        NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1A) & SetNexStarPosition(ManualSkewingSpeed)
    End If
End Sub

Private Sub C_Up_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If SimOffline Then
        SimBntUp = False
    Else
        NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1A) & SetNexStarPosition(0)
    End If
End Sub

Private Sub C_Dn_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If SimOffline Then
        SimBntDn = True
    Else
        NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1B) & SetNexStarPosition(ManualSkewingSpeed)
    End If
End Sub

Private Sub C_Dn_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If SimOffline Then
        SimBntDn = False
    Else
        NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1A) & SetNexStarPosition(0)
    End If
End Sub

Private Sub C_Le_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If SimOffline Then
        SimBntLe = True
    Else
        NexStarComm.Output = Chr$(&H7) & SetNexStarPosition(ManualSkewingSpeed) & Chr$(&H1A) & SetNexStarPosition(0)
    End If
End Sub

Private Sub C_Le_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If SimOffline Then
        SimBntLe = False
    Else
      NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1A) & SetNexStarPosition(0)
    End If
End Sub

Private Sub C_Ri_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If SimOffline Then
        SimBntRi = True
    Else
        NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(ManualSkewingSpeed) & Chr$(&H1A) & SetNexStarPosition(0)
    End If
End Sub

Private Sub C_Ri_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If SimOffline Then
        SimBntRi = False
    Else
        NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1A) & SetNexStarPosition(0)
    End If
End Sub









Private Sub C_SetAzAlt_Click()
    Dim SetAz As Long
    Dim SetAlt As Long
    
    SetAz = CLng(Zahl(T_Az))
    SetAlt = CLng(Zahl(T_Alt))

    SimGotoAzAltActive = True
    
    If SimOffline Then
        SimGotoAz = SetAz
        SimGotoAlt = SetAlt
    Else
        NexStarComm.Output = Chr$(&O2) & SetNexStarPosition(SetAz) & Chr$(&H16) & SetNexStarPosition(SetAlt)
    End If
    
End Sub

Private Sub C_SetEncoder_Click()
    If SimOffline Then
    Else
        NexStarComm.Output = Chr$(&HC) & SetNexStarPosition(EncoderResolution) & SetNexStarPosition(EncoderResolution)
    End If
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


Private Sub Command4_Click()
    Dim Jetzt As String
    Dim Lont As MyTime

' Datensatz Saturn Demo aus dem Sript
    Dim SaturnDemoDate As MyDate
    Dim SaturnDemoTime As MyTime
    SaturnDemoDate.YY = 1978    '13.11.1978
    SaturnDemoDate.MM = 11
    SaturnDemoDate.DD = 13
    SaturnDemoTime.H = 4        '4:34:00 UT   5:34:00 Ortszeit
    SaturnDemoTime.M = 34
    SaturnDemoTime.s = 0

    Dim RA_Saturn As MyTime
    RA_Saturn.H = 10
    RA_Saturn.M = 57              '57
    RA_Saturn.s = 35.681

    Dim DEC_Saturn As MyTime
    DEC_Saturn.H = 8
    DEC_Saturn.M = 25
    DEC_Saturn.s = 58.1

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
    Dim HourAngle As MyTime

    RA_DEC_to_AZ_ALT RA_Saturn, DEC_Saturn, Longitude, Latitude, SaturnDemoTime, SaturnDemoDate, Az, Alt, HourAngle

'    L_AzStar = AZ
    L_AzStar = CutAngle(Az)
    L_AltStar = Alt
    L_HourAngle = HourAngle.H & ":" & HourAngle.M & ":" & Format(HourAngle.s, "00.00")
    
    
    
    
''' ' Capella Kassel
'''    Dim CapellaDemoDate As MyDate
'''    Dim CapellaDemoTime As MyTime
'''    CapellaDemoDate.YY = 2019
'''    CapellaDemoDate.MM = 2
'''    CapellaDemoDate.DD = 2
'''    CapellaDemoTime.h = 19
'''    CapellaDemoTime.M = 0
'''    CapellaDemoTime.s = 0
'''
'''    Dim RA_Capella As MyTime
'''    RA_Capella.h = 5
'''    RA_Capella.M = 18
'''    RA_Capella.s = 6
'''
'''    Dim DEC_Capella As MyTime
'''    DEC_Capella.h = 46
'''    DEC_Capella.M = 1
'''    DEC_Capella.s = 0
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
'''    Dim AZ As Double
'''    Dim ALT As Double
'''    Dim HourAngle As MyTime
'''     RA_DEC_to_AZ_ALT RA_Capella, DEC_Capella, Longitude, Latitude, CapellaDemoTime, CapellaDemoDate, AZ, ALT, HourAngle
'''
''''    L_AzStar = AZ
'''    L_AzStar = CutAngle(AZ)
'''    L_AltStar = ALT
'''    L_HourAngle = HourAngle.h & ":" & HourAngle.M & ":" & Format(HourAngle.s, "00.00")
    
    
    
''' ' Deneb München
'''    Dim DenebDemoDate As MyDate
'''    Dim DenebDemoTime As MyTime
'''    DenebDemoDate.YY = 2019
'''    DenebDemoDate.MM = 2
'''    DenebDemoDate.DD = 2
'''    DenebDemoTime.h = 19
'''    DenebDemoTime.M = 0
'''    DenebDemoTime.s = 0
'''
'''    Dim RA_Deneb As MyTime
'''    RA_Deneb.h = 20
'''    RA_Deneb.M = 42
'''    RA_Deneb.s = 4
'''
'''    Dim DEC_Deneb As MyTime
'''    DEC_Deneb.h = 45
'''    DEC_Deneb.M = 21
'''    DEC_Deneb.s = 0
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
'''    Dim AZ As Double
'''    Dim ALT As Double
'''    Dim HourAngle As MyTime
'''
'''     RA_DEC_to_AZ_ALT RA_Deneb, DEC_Deneb, Longitude, Latitude, DenebDemoTime, DenebDemoDate, AZ, ALT, HourAngle
'''
''''    L_AzStar = AZ
'''    L_AzStar = CutAngle(AZ + 180)
'''    L_AltStar = ALT
'''    L_HourAngle = HourAngle.h & ":" & HourAngle.M & ":" & Format(HourAngle.s, "00.00")
   
    
    
    
    
    
    

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

Private Sub Command6_Click()
    Dim M(10, 10) As Double
    Dim i(10, 10) As Double
    
    
'    m(1, 1) = 1
'    m(1, 2) = 2
'    m(1, 3) = 4
'
'    m(2, 1) = 1
'    m(2, 2) = 4
'    m(2, 3) = 1
'
'    m(3, 1) = 4
'    m(3, 2) = 2
'    m(3, 3) = 2
    M(0, 0) = 1
    M(0, 1) = 2
    M(0, 2) = 4
    M(1, 0) = 1
    M(1, 1) = 4
    M(1, 2) = 1
    M(2, 0) = 4
    M(2, 1) = 2
    M(2, 2) = 2
    Calculate_Inverse 3, M, i
    
    Dim I11 As Double
    Dim I12 As Double
    Dim I13 As Double
    Dim I21 As Double
    Dim I22 As Double
    Dim I23 As Double
    Dim I31 As Double
    Dim I32 As Double
    Dim I33 As Double
    
'    I11 = i(1, 1)
'    I12 = i(1, 2)
'    I13 = i(1, 3)
'    I21 = i(2, 1)
'    I22 = i(2, 2)
'    I23 = i(2, 3)
'    I31 = i(3, 1)
'    I32 = i(3, 2)
'    I33 = i(3, 3)
 
I11 = i(0, 0)
I12 = i(0, 1)
I13 = i(0, 2)
I21 = i(1, 0)
I22 = i(1, 1)
I23 = i(1, 2)
I31 = i(2, 0)
I32 = i(2, 1)
I33 = i(2, 2)
   
End Sub

'''Private Sub Command5_Click()
'''    ' matrix_method_rev_d.pdf Seite 15
'''    Dim SaturnDemoDate As MyDate
'''    Dim SaturnDemoTime As MyTime
'''    Dim SiderialTime As MyTime
'''    Dim RA_Saturn As MyTime
'''    Dim tmp As Double
'''    Dim ttmp As MyTime
'''    Dim LocalHourAngleHour As Double    'Local hour angle in hour (decimal)
'''    Dim LocalHourAngleDeg As Double     'Local hour angle in degree
'''    Dim LocalHourAngleRad As Double     'Local hour angle in radian
'''
'''    Dim Longitude As MyTime                     ' Observer’s longitude
'''    Dim LongitudeDeg As Double
'''
'''
'''    Longitude.h = 0
'''    Longitude.M = 17
'''    Longitude.s = 25.94
'''    Longitude = TimeHMStoDez(Longitude)
'''    LongitudeDeg = Longitude.TimeDec * 15
'''
'''
'''
'''    SaturnDemoDate.YY = 1978
'''    SaturnDemoDate.MM = 11
'''    SaturnDemoDate.DD = 13
'''    SaturnDemoTime.h = 4
'''    SaturnDemoTime.M = 34
'''    SaturnDemoTime.s = 0
'''
'''    'Calculate siderial time
'''    SiderialTime = GetSiderialTime(SaturnDemoDate, SaturnDemoTime, LongitudeDeg)
'''
'''    RA_Saturn.h = 10
'''    RA_Saturn.M = 57
'''    RA_Saturn.s = 35.681
'''    RA_Saturn = TimeHMStoDez(RA_Saturn)
'''
'''    'Calculate local hour angle
'''    LocalHourAngleHour = SiderialTime.TimeDec - RA_Saturn.TimeDec
'''    LocalHourAngleDeg = LocalHourAngleHour * 15
'''    LocalHourAngleRad = LocalHourAngleDeg / (180 / Pi)
'''
'''    'Calculate Saturn position in rectangular equatorial coordinate system
'''    Dim LMN_Equatorial As Vector         ' Rectangular equatorial coordinate system
'''    Dim LMN_EquaMatrix(10, 10) As Double
'''
'''    Dim DEC_Saturn As MyTime
'''    Dim DeclinationDeg As Double
'''    Dim DeclinationRad As Double
'''
'''    DEC_Saturn.h = 8
'''    DEC_Saturn.M = 25
'''    DEC_Saturn.s = 58.1
'''
'''    DEC_Saturn = TimeHMStoDez(DEC_Saturn)
'''    DeclinationDeg = DEC_Saturn.TimeDec
'''    DeclinationRad = DeclinationDeg / (180 / Pi)
'''    LMN_Equatorial = PolarKarthesisch(LocalHourAngleRad, DeclinationRad)
'''
'''    LMN_EquaMatrix(0, 0) = LMN_Equatorial.x
'''    LMN_EquaMatrix(1, 0) = LMN_Equatorial.Y
'''    LMN_EquaMatrix(2, 0) = LMN_Equatorial.z
'''
'''    'Calculate Saturn position in rectangular horizontal coordinate system
'''    Dim LMN_Horizontal As Vector                ' Rectangular horizontal coordinate system
'''    Dim LMN_HorizMatrix(10, 10) As Double
'''    Dim TransformationMatrix(10, 10) As Double  ' Transformation-Matrix from equatorial Coordinates to horizontal Coordinates
'''    Dim Latitude As MyTime                     ' Observer’s latitude
'''    Dim LatitudeDeg As Double
'''    Dim LatitudeRad As Double
'''    Dim Phi As Double                           ' Observer’s latitude
'''
'''    Latitude.h = 50
'''    Latitude.M = 47
'''    Latitude.s = 55
'''    Latitude = TimeHMStoDez(Latitude)
'''    LatitudeDeg = Latitude.TimeDec
'''    LatitudeRad = LatitudeDeg / (180 / Pi)
'''
'''    Phi = LatitudeRad
'''    TransformationMatrix(0, 0) = Cos(Phi - Pi / 2)
'''    TransformationMatrix(0, 1) = 0
'''    TransformationMatrix(0, 2) = Sin(Phi - Pi / 2)
'''    TransformationMatrix(1, 0) = 0
'''    TransformationMatrix(1, 1) = 1
'''    TransformationMatrix(1, 2) = 0
'''    TransformationMatrix(2, 0) = -Sin(Phi - Pi / 2)
'''    TransformationMatrix(2, 1) = 0
'''    TransformationMatrix(2, 2) = Cos(Phi - Pi / 2)
'''
'''    MatrixProduct TransformationMatrix, 3, 3, LMN_EquaMatrix, 3, 1, LMN_HorizMatrix
'''
'''    Dim Lh As Double
'''    Dim Mh As Double
'''    Dim Nh As Double
'''
'''    Lh = LMN_HorizMatrix(0, 0)                  ' Rectangular horizontal  coordinate system
'''    Mh = LMN_HorizMatrix(1, 0)
'''    Nh = LMN_HorizMatrix(2, 0)
'''
'''
'''    'Calculate Saturn position in Altazimuth horizontal coordinate system
'''    Dim AzRad As Double         'Azimuth in radian
'''    Dim AzDeg As Double         'Azimuth in degree
'''    Dim AltRad As Double        'Altitude in radian
'''    Dim AltDeg As Double        'Altitude in degree
'''    Dim sin_h As Double
'''
'''    AzRad = -Atn(Mh / Lh)
'''    AzDeg = AzRad / (Pi / 180)
'''
'''    sin_h = Cos(Phi) * Cos(LocalHourAngleRad) * Cos(DeclinationRad) + Sin(Phi) * Sin(DeclinationRad)
'''    AltRad = arcsin(sin_h)
'''    AltDeg = AltRad / (Pi / 180)
'''
'''    L_AzSaturn = AzDeg
'''    L_AltSaturn = AltDeg
'''End Sub

Private Sub Form_Load()
    SimOffline = True
    
    O_OrientationNorth.Value = 1
    IniFileName = App.Path & "\NexStar.ini"
    InitNexStarComm
    
    Command = 0
    
    VS_ManualSkewingSpeed.Value = 10


    LoadAlignmetStarFile


End Sub


Private Sub InitNexStarComm()

  On Error GoTo v24error
  
  NexStarPortNr = Zahl(INIGetValue(IniFileName, "NexStar", "PortNr"))
  NexStarBaudrate = Zahl(INIGetValue(IniFileName, "NexStar", "Baudrate"))

  If SimOffline Then
    '
  ElseIf NexStarPortNr > 0 Then
    NexStarComm.CommPort = NexStarPortNr
    NexStarComm.Settings = NexStarBaudrate + ",n,8,1"
    NexStarComm.PortOpen = True
  Else
    NexStarComm.CommPort = 6
    NexStarComm.Settings = "4800,n,8,1"
    NexStarComm.PortOpen = True
  End If

  Exit Sub
  
v24error:
  MsgBox "NexStar RS232 Open error: " & Err.Description, , "Communication NexStar"
End Sub





Private Sub NexStarComm_OnComm()
  Dim pos As Long
  Dim vbuf As Variant
  Dim bbuf() As Byte
  Dim key As Integer
  Dim l As Long
  
  
  On Error GoTo msgError
 
  Select Case NexStarComm.CommEvent
  ' Behandeln jedes Ereignisses oder Fehlers durch
  ' Positionieren von Code unter jeder Case-Anweisung

  ' Fehler
    Case comBreak     ' Es wurde ein Anhaltesignal empfangen.
    Case comCDTO      ' CD-Zeitüberschreitung
    Case comCTSTO     ' CTS-Zeitüberschreitung
    Case comDSRTO     ' DSR-Zeitüberschreitung
    Case comFrame     ' Fehler im Übertragungsraster (Framing Error)
    Case comOverrun   ' Datenverlust
    Case comRxOver    ' Überlauf des Empfangspuffers
    Case comRxParity  ' Paritätsfehler
    Case comTxFull    ' Sendepuffer voll
    Case comDCB       ' Unerwarteter Fehler beim Abrufen des DCB]

  ' Ereignisse
    Case comEvCD  ' Pegeländerung auf DCD
    Case comEvCTS ' Pegeländerung auf CTS
    Case comEvDSR ' Pegeländerung auf DSR
    Case comEvRing  ' Pegeländerung auf RI(Ring Indicator)
    Case comEvReceive ' Anzahl empfangener Zeichen gleich RThreshold
    
            If Command = 1 Then
                Do
                    vbuf = NexStarComm.Input
                    bbuf = vbuf
                    NexStarAz = NexStarAz & Chr$(bbuf(0))
                     key = (bbuf(0))
                Loop While NexStarComm.InBufferCount > 0
                l = Len(NexStarAz)
                TelIncrAz = GetNexStarPosition(NexStarAz)
            ElseIf Command = 21 Then
                Do
                    vbuf = NexStarComm.Input
                    bbuf = vbuf
                    NexStarAlt = NexStarAlt & Chr$(bbuf(0))
                     key = (bbuf(0))
                Loop While NexStarComm.InBufferCount > 0
                l = Len(NexStarAlt)
                TelIncrAlt = GetNexStarPosition(NexStarAlt)
            End If
        
    Case comEvSend  ' Im Sendepuffer befinden sich SThreshold Zeichen
    Case comEvEOF ' Im Eingabestrom wurde ein EOF-Zeichen gefunden
  End Select
  If NexStarComm.CommEvent <> 2 Then    'empfangen'
'  Kommunikation_DMX_Scanner_OK = False
  Else
'   Kommunikation_DMX_Scanner_OK = True
  End If
  Exit Sub
msgError:
  MsgBox "Error: " + Err.Description + "in Function OnComm() in MainFrm."

End Sub



Private Sub Tim_DisplayUpdate_Timer()
    Static Toggle As Boolean
    
    If Toggle Then
        Toggle = False
        C_GetAz_Click
    Else
        Toggle = True
        C_GetAlt_Click
    End If
    
    L_Az = TelIncrAz
    L_Alt = TelIncrAlt
    
    If O_OrientationNorth.Value Then
        TelDegAz = TelIncrAz * 360 / EncoderResolution
    ElseIf O_OrientationSouth.Value Then
        TelDegAz = (TelIncrAz * 360 / EncoderResolution) + 180
    End If
    
    TelDegAlt = TelIncrAlt * 360 / EncoderResolution
    L_TelDegAz = Format(CutAngle(TelDegAz), "0.0000")
    L_TelDegAlt = Format(TelDegAlt, "0.0000")
    
    
End Sub

Private Sub Tim_Simulation_Timer()
    Dim SimScaling As Long
    Dim SimGotoStep As Long

    SimScaling = 10
    SimGotoStep = 1000

    If SimBntUp Then
        SimIncrAlt = SimIncrAlt + (ManualSkewingSpeed / SimScaling)
    End If
        
    If SimBntDn Then
        SimIncrAlt = SimIncrAlt - (ManualSkewingSpeed / SimScaling)
    End If
    
    If SimBntLe Then
        SimIncrAz = SimIncrAz - (ManualSkewingSpeed / SimScaling)
    End If
        
    If SimBntRi Then
        SimIncrAz = SimIncrAz + (ManualSkewingSpeed / SimScaling)
    End If
        
    If SimIncrAz > EncoderResolution Then
        SimIncrAz = 0
    ElseIf SimIncrAz < 0 Then
        SimIncrAz = EncoderResolution
    End If
        
    If SimIncrAlt > EncoderResolution Then
        SimIncrAlt = 0
    ElseIf SimIncrAlt < 0 Then
        SimIncrAlt = EncoderResolution
    End If
    
    
    If SimGotoAzAltActive Then
        If Abs(SimGotoAz - SimIncrAz) < SimGotoStep Then
            SimIncrAz = SimGotoAz
        ElseIf SimGotoAz > SimIncrAz Then
            SimIncrAz = SimIncrAz + SimGotoStep
        Else
            SimIncrAz = SimIncrAz - SimGotoStep
        End If
    
        If Abs(SimGotoAlt - SimIncrAlt) < SimGotoStep Then
            SimIncrAlt = SimGotoAlt
        ElseIf SimGotoAlt > SimIncrAlt Then
            SimIncrAlt = SimIncrAlt + SimGotoStep
        Else
            SimIncrAlt = SimIncrAlt - SimGotoStep
        End If
        
        
        If (SimIncrAz = SimGotoAz) And (SimIncrAlt = SimGotoAlt) Then
            SimGotoAzAltActive = False
        End If

    End If
    
    
    
    'Dim SimGotoAlt As Long

End Sub



Private Sub Tim_Tracking_Timer()
    Dim LocalUT As Date
    
    LocalUT = UtcTime(Now)

    L_UT = LocalUT

 ' Capella Kassel
    Dim CapellaDemoDate As MyDate
    Dim CapellaDemoTime As MyTime
    CapellaDemoDate.YY = Year(LocalUT)
    CapellaDemoDate.MM = Month(LocalUT)
    CapellaDemoDate.DD = Day(LocalUT)
    CapellaDemoTime.H = Hour(LocalUT)
    CapellaDemoTime.M = Minute(LocalUT)
    CapellaDemoTime.s = Second(LocalUT)

    Dim RA_Capella As MyTime
    RA_Capella.H = 5
    RA_Capella.M = 18
    RA_Capella.s = 6

    Dim DEC_Capella As MyTime
    DEC_Capella.H = 46
    DEC_Capella.M = 1
    DEC_Capella.s = 0

    Dim Longitude As GeoCoord                     ' Observer’s longitude
    Longitude.Deg = 9
    Longitude.Min = 18
    Longitude.Sec = 3
    Longitude.Sign = "E"

    Dim Latitude As GeoCoord                     ' Observer’s latitude
    Latitude.Deg = 51
    Latitude.Min = 11
    Latitude.Sec = 27
    Latitude.Sign = "N"

    Dim Az As Double
    Dim Alt As Double
    Dim HourAngle As MyTime
     RA_DEC_to_AZ_ALT RA_Capella, DEC_Capella, Longitude, Latitude, CapellaDemoTime, CapellaDemoDate, Az, Alt, HourAngle

'    L_AzStar = AZ
    If O_OrientationNorth.Value Then Az = Az + 180
    L_AzStar = CutAngle(Az)
    L_AltStar = Alt
    L_HourAngle = HourAngle.H & ":" & HourAngle.M & ":" & Format(HourAngle.s, "00.00")
    



End Sub

Private Sub VS_ManualSkewingSpeed_Change()
    Dim tmp As Long
    
    tmp = VS_ManualSkewingSpeed.Value
    ManualSkewingSpeed = 1000 * tmp
End Sub




Private Sub LoadAlignmetStarFile()
    Dim AlignmetStarFile As Integer
    Dim AlignmetStarFileName As String
    Dim i As Integer
    Dim Zeile As String
    Dim StarEntities() As String
    Dim idx As Long

    ReDim AlignmentStarArray(0 To 0)
    
    AlignmetStarFile = FreeFile
    On Error GoTo openErr:
    AlignmetStarFileName = App.Path & "\Alignment Stars.txt"
    Open AlignmetStarFileName For Input As AlignmetStarFile
    While Not EOF(AlignmetStarFile)
        Line Input #AlignmetStarFile, Zeile
        SepariereString Zeile, StarEntities, vbTab
        idx = UBound(AlignmentStarArray)
        AlignmentStarArray(idx).ProperName = StarEntities(0)
        AlignmentStarArray(idx).Constellation = StarEntities(1)
        AlignmentStarArray(idx).Bayer = StarEntities(2)
        AlignmentStarArray(idx).Flamsteed = StarEntities(3)
        AlignmentStarArray(idx).RA = Zahl(StarEntities(4))
        AlignmentStarArray(idx).DEC = Zahl(StarEntities(5))
        AlignmentStarArray(idx).Mag = Zahl(StarEntities(6))
        
        AlignmentStarList.AddItem AlignmentStarArray(idx).ProperName
        
        ReDim Preserve AlignmentStarArray(0 To UBound(AlignmentStarArray) + 1)
    Wend
    Close AlignmetStarFile
    
    
    Exit Sub
    
openErr:
    MsgBox Err.Description & vbCrLf & "Can't read Config File:" & AlignmetStarFileName, , " Error "
    Close AlignmetStarFile
End Sub

