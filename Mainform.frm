VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMm32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Mainform 
   Caption         =   "Form1"
   ClientHeight    =   11235
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   11235
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   63
      Text            =   "40"
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton C_SetBacklAlt 
      Caption         =   "Set Backl. Alt"
      Height          =   255
      Left            =   240
      TabIndex        =   62
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton C_SetBacklAz 
      Caption         =   "Set Backl. Az"
      Height          =   255
      Left            =   240
      TabIndex        =   61
      Top             =   480
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1095
      Left            =   480
      TabIndex        =   56
      Top             =   7200
      Width           =   1575
      Begin VB.OptionButton O_TimeSelectLocal 
         Caption         =   "Lokalzeit"
         Height          =   195
         Left            =   120
         TabIndex        =   58
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton O_TimeSelectSim 
         Caption         =   "Simulierte Zeit"
         Height          =   195
         Left            =   120
         TabIndex        =   57
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   975
      Left            =   480
      TabIndex        =   53
      Top             =   5400
      Width           =   1455
      Begin VB.OptionButton O_OrientationNorth 
         Caption         =   "North"
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton O_OrientationSouth 
         Caption         =   "South"
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.TextBox T_Long_Sign 
      Height          =   285
      Left            =   2640
      TabIndex        =   52
      Text            =   "Text1"
      Top             =   9480
      Width           =   255
   End
   Begin VB.CommandButton C_Set_ObserverLocation 
      Caption         =   "Command7"
      Height          =   255
      Left            =   2160
      TabIndex        =   51
      Top             =   10080
      Width           =   1215
   End
   Begin VB.TextBox T_Latt_Sign 
      Height          =   285
      Left            =   120
      TabIndex        =   50
      Text            =   "Text1"
      Top             =   9480
      Width           =   255
   End
   Begin VB.TextBox T_Latt_Grad 
      Height          =   285
      Left            =   480
      TabIndex        =   43
      Text            =   "Text1"
      Top             =   9480
      Width           =   495
   End
   Begin VB.TextBox T_Latt_Min 
      Height          =   285
      Left            =   1080
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   9480
      Width           =   495
   End
   Begin VB.TextBox T_Latt_Sek 
      Height          =   285
      Left            =   1680
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   9480
      Width           =   495
   End
   Begin VB.TextBox T_Long_Sek 
      Height          =   285
      Left            =   4200
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   9480
      Width           =   495
   End
   Begin VB.TextBox T_Long_Min 
      Height          =   285
      Left            =   3600
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   9480
      Width           =   495
   End
   Begin VB.TextBox T_Long_Grad 
      Height          =   285
      Left            =   3000
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   9480
      Width           =   495
   End
   Begin VB.TextBox T_Tag 
      Height          =   285
      Left            =   480
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   8760
      Width           =   495
   End
   Begin VB.TextBox T_Monat 
      Height          =   285
      Left            =   1080
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   8760
      Width           =   495
   End
   Begin VB.TextBox T_Jahr 
      Height          =   285
      Left            =   1680
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   8760
      Width           =   495
   End
   Begin VB.TextBox T_Sekunden 
      Height          =   285
      Left            =   4200
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   8760
      Width           =   495
   End
   Begin VB.TextBox T_Minuten 
      Height          =   285
      Left            =   3600
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   8760
      Width           =   495
   End
   Begin VB.TextBox T_Stunden 
      Height          =   285
      Left            =   3000
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   8760
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7320
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Tim_Tracking 
      Interval        =   1000
      Left            =   7800
      Top             =   5520
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Test Sternzeit"
      Height          =   495
      Left            =   0
      TabIndex        =   16
      Top             =   10680
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
      TabIndex        =   13
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox T_AltTel 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   4320
      TabIndex        =   12
      Text            =   "0"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox T_AzTel 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   4320
      TabIndex        =   11
      Text            =   "0"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.VScrollBar VS_ManualSkewingSpeed 
      Height          =   2295
      LargeChange     =   10
      Left            =   4560
      Max             =   0
      Min             =   100
      TabIndex        =   10
      Top             =   3240
      Value           =   100
      Width           =   255
   End
   Begin VB.CommandButton C_Le 
      Caption         =   "<"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton C_Dn 
      Caption         =   "V"
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton C_Ri 
      Caption         =   ">"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton C_Up 
      Caption         =   "^"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   3720
      Width           =   375
   End
   Begin VB.ListBox AlignmentStarList 
      Height          =   9615
      Left            =   5880
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton C_SetEncoder 
      Caption         =   "Set Encoder"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton C_GetAlt 
      Caption         =   "Get Alt"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton C_GetAz 
      Caption         =   "Get Az"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   1335
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
   Begin VB.Frame Frame3 
      Caption         =   "Time"
      Height          =   1455
      Left            =   240
      TabIndex        =   64
      Top             =   1920
      Width           =   4335
      Begin VB.Label L_UTime 
         Caption         =   "UT"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label L_SiderialTime 
         Caption         =   "Siderial Time"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   840
         Width           =   4095
      End
      Begin VB.Label L_LocalTime 
         Caption         =   "Local Time"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Label Label13 
      Caption         =   "Skewing Speed"
      Height          =   255
      Left            =   2880
      TabIndex        =   60
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label L_SkewingSpeed 
      Alignment       =   1  'Rechts
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "--"
      Height          =   255
      Left            =   2880
      TabIndex        =   59
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "Sek."
      Height          =   255
      Left            =   4200
      TabIndex        =   49
      Top             =   9240
      Width           =   375
   End
   Begin VB.Label Label15 
      Caption         =   "Min."
      Height          =   255
      Left            =   3600
      TabIndex        =   48
      Top             =   9240
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "Grad"
      Height          =   255
      Left            =   3000
      TabIndex        =   47
      Top             =   9240
      Width           =   375
   End
   Begin VB.Label Label19 
      Caption         =   "Sek."
      Height          =   255
      Left            =   1680
      TabIndex        =   46
      Top             =   9240
      Width           =   375
   End
   Begin VB.Label Label18 
      Caption         =   "Min."
      Height          =   255
      Left            =   1080
      TabIndex        =   45
      Top             =   9240
      Width           =   495
   End
   Begin VB.Label Label17 
      Caption         =   "Grad"
      Height          =   255
      Left            =   480
      TabIndex        =   44
      Top             =   9240
      Width           =   375
   End
   Begin VB.Label Label12 
      Caption         =   "Jahr"
      Height          =   255
      Left            =   1680
      TabIndex        =   37
      Top             =   8520
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "Monat"
      Height          =   255
      Left            =   1080
      TabIndex        =   36
      Top             =   8520
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "Tag"
      Height          =   255
      Left            =   480
      TabIndex        =   35
      Top             =   8520
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "H"
      Height          =   255
      Left            =   3120
      TabIndex        =   34
      Top             =   8520
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "M"
      Height          =   255
      Left            =   3720
      TabIndex        =   33
      Top             =   8520
      Width           =   255
   End
   Begin VB.Label Label11 
      Caption         =   "S"
      Height          =   255
      Left            =   4320
      TabIndex        =   32
      Top             =   8520
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "Ortszeit"
      Height          =   255
      Left            =   4920
      TabIndex        =   31
      Top             =   8760
      Width           =   615
   End
   Begin VB.Label L_CurrentStar 
      BorderStyle     =   1  'Fest Einfach
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
      TabIndex        =   24
      Top             =   10200
      Width           =   2775
   End
   Begin VB.Label L_UT 
      Alignment       =   1  'Rechts
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "--"
      Height          =   255
      Left            =   2400
      TabIndex        =   23
      Top             =   8040
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Hour Angle"
      Height          =   255
      Left            =   2280
      TabIndex        =   22
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label L_HourAngle 
      Alignment       =   1  'Rechts
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "--"
      Height          =   255
      Left            =   3120
      TabIndex        =   21
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Altitude"
      Height          =   255
      Left            =   2280
      TabIndex        =   20
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Azimuth"
      Height          =   255
      Left            =   2280
      TabIndex        =   19
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label L_AzStar 
      Alignment       =   1  'Rechts
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "--"
      Height          =   255
      Left            =   3120
      TabIndex        =   18
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label L_AltStar 
      Alignment       =   1  'Rechts
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "--"
      Height          =   255
      Left            =   3120
      TabIndex        =   17
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label L_TelDegAlt 
      Alignment       =   1  'Rechts
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "--"
      Height          =   375
      Left            =   4800
      TabIndex        =   15
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label L_TelDegAz 
      Alignment       =   1  'Rechts
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "--"
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label L_Az 
      Alignment       =   1  'Rechts
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "--"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label L_Alt 
      Alignment       =   1  'Rechts
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "--"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Menu M_Setup 
      Caption         =   "Setup"
   End
   Begin VB.Menu M_Test 
      Caption         =   "Test"
   End
   Begin VB.Menu M_Communication 
      Caption         =   "Communication"
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


Dim ObserverDateTimeUT As Date
Dim ObserverLatt As GeoCoord
Dim ObserverLong As GeoCoord
Dim ObserverRA As Double
Dim ObserverDEC As Double


Private Sub AlignmentStarList_Click()
    Dim idx As Long
    Dim Az As Double
    Dim Alt As Double
    Dim HourAngle As Double
    Dim HourAngleHMS As MyTime
    Dim SelectedStar As String

    ' Search star name in list
    idx = -1
    Do
        idx = idx + 1
    Loop Until (AlignmentStarArray(idx).ProperName = AlignmentStarList.Text) Or (idx >= UBound(AlignmentStarArray))
    L_CurrentStar = AlignmentStarArray(idx).ProperName
    
    ObserverRA = HourToRad(AlignmentStarArray(idx).RA)
    ObserverDEC = DegToRad(AlignmentStarArray(idx).DEC)
      
    RA_DEC_to_AZ_ALT_radian ObserverRA, ObserverDEC, ObserverLong, ObserverLatt, ObserverDateTimeUT, Az, Alt, HourAngle
  
    If O_OrientationNorth.Value Then Az = Az + Pi
    L_AzStar = CutAngle(RadToDeg(Az))
    L_AltStar = RadToDeg(Alt)

    HourAngleHMS = RadToTime(HourAngle)
    L_HourAngle = HourAngleHMS.H & ":" & HourAngleHMS.M & ":" & Format(HourAngleHMS.s, "00.00")

    T_AzTel = CutAngle(RadToDeg(Az))
    T_AltTel = RadToDeg(Alt)
    

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




Private Sub C_Set_ObserverLocation_Click()
    
    
    ObserverLatt.Min = Zahl(T_Latt_Min)
    ObserverLatt.Sec = Zahl(T_Latt_Sek)
    ObserverLatt.Sign = T_Latt_Sign
    ObserverLong.Deg = Zahl(T_Long_Grad)
    ObserverLong.Min = Zahl(T_Long_Min)
    ObserverLong.Sec = Zahl(T_Long_Sek)
    ObserverLong.Sign = T_Long_Sign
    
    INISetValue IniFileName, "Datum", "Tag", T_Tag
    INISetValue IniFileName, "Datum", "Monat", T_Monat
    INISetValue IniFileName, "Datum", "Jahr", T_Jahr
    INISetValue IniFileName, "Zeit", "Stunden", T_Stunden
    INISetValue IniFileName, "Zeit", "Minuten", T_Minuten
    INISetValue IniFileName, "Zeit", "Sekunden", T_Sekunden
  
    INISetValue IniFileName, "Ort", "LattGrad", T_Latt_Grad
    INISetValue IniFileName, "Ort", "LattMin", T_Latt_Min
    INISetValue IniFileName, "Ort", "LattSek", T_Latt_Sek
    INISetValue IniFileName, "Ort", "LattSign", T_Latt_Sign
    INISetValue IniFileName, "Ort", "LongGrad", T_Long_Grad
    INISetValue IniFileName, "Ort", "LongMin", T_Long_Min
    INISetValue IniFileName, "Ort", "LongSek", T_Long_Sek
    INISetValue IniFileName, "Ort", "LongSing", T_Long_Sign

End Sub

Private Sub C_SetBacklAlt_Click()
    Dim BacklashAlt As Long    '0..100

    BacklashAlt = Text1
    
    If SimOffline Then
    Else
        NexStarComm.Output = Chr$(&H1E) & SetNexStarPosition(BacklashAlt)
    End If
End Sub

Private Sub C_SetBacklAz_Click()
    Dim BacklashAz As Long    '0..100

    BacklashAz = Text1
    
    If SimOffline Then
    Else
        NexStarComm.Output = Chr$(&HA) & SetNexStarPosition(BacklashAz)
    End If
End Sub

Private Sub C_Up_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SimOffline Then
        SimBntUp = True
    Else
        NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1A) & SetNexStarPosition(ManualSkewingSpeed)
    End If
End Sub

Private Sub C_Up_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SimOffline Then
        SimBntUp = False
    Else
        NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1A) & SetNexStarPosition(0)
    End If
End Sub

Private Sub C_Dn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SimOffline Then
        SimBntDn = True
    Else
        NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1B) & SetNexStarPosition(ManualSkewingSpeed)
    End If
End Sub

Private Sub C_Dn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SimOffline Then
        SimBntDn = False
    Else
        NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1A) & SetNexStarPosition(0)
    End If
End Sub

Private Sub C_Le_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SimOffline Then
        SimBntLe = True
    Else
        NexStarComm.Output = Chr$(&H7) & SetNexStarPosition(ManualSkewingSpeed) & Chr$(&H1A) & SetNexStarPosition(0)
    End If
End Sub

Private Sub C_Le_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SimOffline Then
        SimBntLe = False
    Else
      NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1A) & SetNexStarPosition(0)
    End If
End Sub

Private Sub C_Ri_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SimOffline Then
        SimBntRi = True
    Else
        NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(ManualSkewingSpeed) & Chr$(&H1A) & SetNexStarPosition(0)
    End If
End Sub

Private Sub C_Ri_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SimOffline Then
        SimBntRi = False
    Else
        NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1A) & SetNexStarPosition(0)
    End If
End Sub


Private Sub C_SetAzAlt_Click()
    Dim SetAz As Long
    Dim SetAlt As Long
    
    SetAz = CLng(Zahl(T_AzTel) * EncoderResolution / 360)
    SetAlt = CLng(Zahl(T_AltTel) * EncoderResolution / 360)

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


Private Sub Form_Load()
    SimOffline = True
    
    O_OrientationNorth.Value = 1
    O_TimeSelectLocal.Value = 1
    IniFileName = App.Path & "\NexStar.ini"
    InitNexStarComm
    
    Command = 0
    
    VS_ManualSkewingSpeed.Value = 10


    LoadAlignmetStarFile

    T_Latt_Grad = INIGetValue(IniFileName, "Ort", "LattGrad")
    T_Latt_Min = INIGetValue(IniFileName, "Ort", "LattMin")
    T_Latt_Sek = INIGetValue(IniFileName, "Ort", "LattSek")
    T_Latt_Sign = INIGetValue(IniFileName, "Ort", "LattSign")
    T_Long_Grad = INIGetValue(IniFileName, "Ort", "LongGrad")
    T_Long_Min = INIGetValue(IniFileName, "Ort", "LongMin")
    T_Long_Sek = INIGetValue(IniFileName, "Ort", "LongSek")
    T_Long_Sign = INIGetValue(IniFileName, "Ort", "LongSing")
    
    ObserverLatt.Deg = Zahl(T_Latt_Grad)
    ObserverLatt.Min = Zahl(T_Latt_Min)
    ObserverLatt.Sec = Zahl(T_Latt_Sek)
    ObserverLatt.Sign = T_Latt_Sign
    ObserverLong.Deg = Zahl(T_Long_Grad)
    ObserverLong.Min = Zahl(T_Long_Min)
    ObserverLong.Sec = Zahl(T_Long_Sek)
    ObserverLong.Sign = T_Long_Sign
    
    T_Tag = INIGetValue(IniFileName, "Datum", "Tag")
    T_Monat = INIGetValue(IniFileName, "Datum", "Monat")
    T_Jahr = INIGetValue(IniFileName, "Datum", "Jahr")
    
    T_Stunden = INIGetValue(IniFileName, "Zeit", "Stunden")
    T_Minuten = INIGetValue(IniFileName, "Zeit", "Minuten")
    T_Sekunden = INIGetValue(IniFileName, "Zeit", "Sekunden")

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




Private Sub M_Communication_Click()
    Communication.Show
End Sub

Private Sub M_Test_Click()
    Test.Show
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
    Dim tTime As MyTime
    Dim tDate As MyDate
    Dim tTs As MyTime
  
    

    If O_TimeSelectLocal.Value = True Then
        ObserverDateTimeUT = UtcTime(Now)              ' Get current Time
        L_UT = ObserverDateTimeUT                      ' Convert to UT
        
        L_LocalTime = " Local time:   " & Now
        L_UTime = " UT:              " & ObserverDateTimeUT
        
        tTime.H = Hour(ObserverDateTimeUT)
        tTime.M = Minute(ObserverDateTimeUT)
        tTime.s = Second(ObserverDateTimeUT)
        tDate.YY = Year(ObserverDateTimeUT)
        tDate.MM = Month(ObserverDateTimeUT)
        tDate.DD = Day(ObserverDateTimeUT)
        tTs = GMST(tDate, tTime)
        L_SiderialTime = "Siderial time: " & tTs.H & ":" & tTs.M & ":" & CInt(tTs.s)
        
        
        
    Else
        ' Take simulation time
        ObserverDateTimeUT = StingsToDate(T_Tag, T_Monat, T_Jahr, T_Stunden, T_Minuten, T_Sekunden)
    End If

    Dim idx As Long
    Dim Az As Double
    Dim Alt As Double
    Dim HourAngle As Double
    Dim HourAngleHMS As MyTime
    
    idx = AlignmentStarList.ListIndex
    
    ' no star selected yet
    If idx < 0 Then
        Exit Sub
    End If
    
    ' Search star name in list
    idx = -1
    Do
        idx = idx + 1
    Loop Until (AlignmentStarArray(idx).ProperName = AlignmentStarList.Text) Or (idx >= UBound(AlignmentStarArray))
    L_CurrentStar = AlignmentStarArray(idx).ProperName
    
    ObserverRA = HourToRad(AlignmentStarArray(idx).RA)
    ObserverDEC = DegToRad(AlignmentStarArray(idx).DEC)
     
    RA_DEC_to_AZ_ALT_radian ObserverRA, ObserverDEC, ObserverLong, ObserverLatt, ObserverDateTimeUT, Az, Alt, HourAngle

    If O_OrientationNorth.Value Then Az = Az + Pi
    L_AzStar = CutAngle(RadToDeg(Az))
    L_AltStar = RadToDeg(Alt)

    HourAngleHMS = RadToTime(HourAngle)

    T_AzTel = CutAngle(RadToDeg(Az))
    T_AltTel = RadToDeg(Alt)
    
    If Alt < 0 Then
        L_CurrentStar.BackColor = RGB(255, 0, 0)
    ElseIf (Alt > 0) And (Alt < 0.3) Then
        L_CurrentStar.BackColor = RGB(255, 255, 0)
    Else
        L_CurrentStar.BackColor = RGB(0, 255, 0)
    End If
    
    
    L_HourAngle = HourAngleHMS.H & ":" & HourAngleHMS.M & ":" & Format(HourAngleHMS.s, "00.00")
    


End Sub

Private Sub VS_ManualSkewingSpeed_Change()
    Dim tmp As Long
    
    tmp = VS_ManualSkewingSpeed.Value
    ManualSkewingSpeed = 1000 * tmp
    L_SkewingSpeed = ManualSkewingSpeed
    
    'SkewingSpeed[Incr/sec] = ManualSkewingSpeed[Incr/sec] * 0,1 [Incr/sec]
    '
    
    
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
        AlignmentStarArray(idx).Index = Zahl(StarEntities(0))
        AlignmentStarArray(idx).ProperName = StarEntities(1)
        AlignmentStarArray(idx).Constellation = StarEntities(2)
        AlignmentStarArray(idx).Bayer = StarEntities(3)
        AlignmentStarArray(idx).Flamsteed = StarEntities(4)
        AlignmentStarArray(idx).RA = Zahl(StarEntities(5))
        AlignmentStarArray(idx).DEC = Zahl(StarEntities(6))
        AlignmentStarArray(idx).Mag = Zahl(StarEntities(7))
        
        AlignmentStarList.AddItem AlignmentStarArray(idx).ProperName
        
        ReDim Preserve AlignmentStarArray(0 To UBound(AlignmentStarArray) + 1)
    Wend
    Close AlignmetStarFile
    
    Dim tst
    tst = AlignmentStarList.List(2)
    
    
    Exit Sub
    
openErr:
    MsgBox Err.Description & vbCrLf & "Can't read Config File:" & AlignmetStarFileName, , " Error "
    Close AlignmetStarFile
End Sub








