VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Mainform 
   Caption         =   "Form1"
   ClientHeight    =   11235
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12690
   LinkTopic       =   "Form1"
   ScaleHeight     =   11235
   ScaleWidth      =   12690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton C_Tracking 
      Caption         =   "Tracking"
      Height          =   495
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   9240
      Width           =   1215
   End
   Begin VB.CommandButton C_SingleStarAlignment 
      Caption         =   "Single Star Alignment"
      Height          =   495
      Left            =   480
      TabIndex        =   99
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton C_GotoStarCalibrated 
      Caption         =   "GotoStar calibrated"
      Height          =   255
      Left            =   8640
      TabIndex        =   97
      Top             =   10800
      Width           =   2175
   End
   Begin VB.CommandButton C_CalibrateNow 
      Caption         =   "Calibrate now"
      Height          =   495
      Left            =   3000
      TabIndex        =   96
      Top             =   7560
      Width           =   2295
   End
   Begin VB.CommandButton C_SetCalibrationStar_2 
      Caption         =   "Set Calibration Star 2"
      Height          =   495
      Left            =   3000
      TabIndex        =   95
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton C_SetCalibrationStar_1 
      Caption         =   "Set Calibration Star 1"
      Height          =   495
      Left            =   3000
      TabIndex        =   94
      Top             =   6360
      Width           =   2295
   End
   Begin VB.CommandButton C_GotoStar 
      Caption         =   "GotoStar w/o calibration"
      Height          =   255
      Left            =   5880
      TabIndex        =   93
      Top             =   10800
      Width           =   2175
   End
   Begin VB.CommandButton C_SetNorth 
      Caption         =   "Set North"
      Height          =   495
      Left            =   3000
      TabIndex        =   92
      Top             =   5760
      Width           =   2295
   End
   Begin VB.Frame Frame4 
      Caption         =   "Altitude"
      Height          =   2415
      Left            =   8160
      TabIndex        =   81
      Top             =   6360
      Width           =   3855
      Begin VB.Label Label46 
         Caption         =   "Mot. Incr."
         Height          =   255
         Left            =   240
         TabIndex        =   91
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label L_AltMotorIncr 
         Caption         =   "--"
         Height          =   255
         Left            =   1200
         TabIndex        =   90
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label L_MatrixSystemAltIst 
         Caption         =   "--"
         Height          =   255
         Left            =   1200
         TabIndex        =   89
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label43 
         Caption         =   "Matr Sy Ist:"
         Height          =   255
         Left            =   240
         TabIndex        =   88
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label42 
         Caption         =   "AzAlt Sys:"
         Height          =   255
         Left            =   240
         TabIndex        =   87
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label41 
         Caption         =   "--"
         Height          =   255
         Left            =   1200
         TabIndex        =   86
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label L_MatrixSystemAltSoll 
         Caption         =   "--"
         Height          =   255
         Left            =   1200
         TabIndex        =   85
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label L_GlobalAltOffset 
         Caption         =   "--"
         Height          =   255
         Left            =   1200
         TabIndex        =   84
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label36 
         Caption         =   "Matr Soll"
         Height          =   255
         Left            =   240
         TabIndex        =   83
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Glob. Offs."
         Height          =   255
         Left            =   240
         TabIndex        =   82
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Azimut"
      Height          =   2415
      Left            =   8160
      TabIndex        =   70
      Top             =   3720
      Width           =   3855
      Begin VB.Label Label31 
         Caption         =   "Glob. Offs."
         Height          =   255
         Left            =   240
         TabIndex        =   80
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label32 
         Caption         =   "Matr Soll"
         Height          =   255
         Left            =   240
         TabIndex        =   79
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label L_GlobalAzOffset 
         Caption         =   "--"
         Height          =   255
         Left            =   1200
         TabIndex        =   78
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label L_MatrixSystemAzSoll 
         Caption         =   "--"
         Height          =   255
         Left            =   1200
         TabIndex        =   77
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label L_AzAltSystem 
         Caption         =   "--"
         Height          =   255
         Left            =   1200
         TabIndex        =   76
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label26 
         Caption         =   "AzAlt Sys:"
         Height          =   255
         Left            =   240
         TabIndex        =   75
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "Matr Sy Ist:"
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label L_MatrixSystemAzIst 
         Caption         =   "--"
         Height          =   255
         Left            =   1200
         TabIndex        =   73
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label L_AzMotorIncr 
         Caption         =   "--"
         Height          =   255
         Left            =   1200
         TabIndex        =   72
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Mot. Incr."
         Height          =   255
         Left            =   240
         TabIndex        =   71
         Top             =   1800
         Width           =   855
      End
   End
   Begin VB.Frame F_StarInfo 
      Caption         =   "--"
      Height          =   2415
      Left            =   8160
      TabIndex        =   55
      Top             =   360
      Width           =   3855
      Begin VB.CheckBox Ch_South 
         Caption         =   "South (VSky)"
         Height          =   255
         Left            =   2280
         TabIndex        =   98
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label25 
         Caption         =   "Horiz. xyz:"
         Height          =   255
         Left            =   240
         TabIndex        =   69
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label L_I_HorXYZ 
         Caption         =   "Alt:"
         Height          =   255
         Left            =   1200
         TabIndex        =   68
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label L_I_EquXYZ 
         Caption         =   "Alt:"
         Height          =   255
         Left            =   1200
         TabIndex        =   67
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label22 
         Caption         =   "Equ. xyz:"
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "Hour Angle:"
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label L_I_HourAngle 
         Caption         =   "Alt:"
         Height          =   255
         Left            =   1200
         TabIndex        =   64
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label L_I_Alt 
         Caption         =   "Alt:"
         Height          =   255
         Left            =   1200
         TabIndex        =   63
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label L_I_Az 
         Caption         =   "Az:"
         Height          =   255
         Left            =   1200
         TabIndex        =   62
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label L_I_DEC 
         Caption         =   "DEC:"
         Height          =   255
         Left            =   1200
         TabIndex        =   61
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label L_I_RA 
         Caption         =   "RA:"
         Height          =   255
         Left            =   1200
         TabIndex        =   60
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "Alt:"
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "Az:"
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "DEC:"
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "RA:"
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   240
         Width           =   375
      End
   End
   Begin MSCommLib.MSComm NexStarComm 
      Left            =   6120
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox T_Backlash 
      Height          =   285
      Left            =   480
      TabIndex        =   48
      Text            =   "40"
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton C_SetBacklAlt 
      Caption         =   "Set Backl. Alt"
      Height          =   255
      Left            =   240
      TabIndex        =   47
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton C_SetBacklAz 
      Caption         =   "Set Backl. Az"
      Height          =   255
      Left            =   240
      TabIndex        =   46
      Top             =   480
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   975
      Left            =   480
      TabIndex        =   41
      Top             =   5400
      Width           =   1455
      Begin VB.OptionButton O_OrientationNorth 
         Caption         =   "North"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton O_OrientationSouth 
         Caption         =   "South"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.TextBox T_Long_Sign 
      Height          =   285
      Left            =   2640
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   9480
      Width           =   255
   End
   Begin VB.CommandButton C_Set_ObserverLocation 
      Caption         =   "Command7"
      Height          =   255
      Left            =   2160
      TabIndex        =   39
      Top             =   10080
      Width           =   1215
   End
   Begin VB.TextBox T_Latt_Sign 
      Height          =   285
      Left            =   120
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   9480
      Width           =   255
   End
   Begin VB.TextBox T_Latt_Grad 
      Height          =   285
      Left            =   480
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   9480
      Width           =   495
   End
   Begin VB.TextBox T_Latt_Min 
      Height          =   285
      Left            =   1080
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   9480
      Width           =   495
   End
   Begin VB.TextBox T_Latt_Sek 
      Height          =   285
      Left            =   1680
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   9480
      Width           =   495
   End
   Begin VB.TextBox T_Long_Sek 
      Height          =   285
      Left            =   4200
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   9480
      Width           =   495
   End
   Begin VB.TextBox T_Long_Min 
      Height          =   285
      Left            =   3600
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   9480
      Width           =   495
   End
   Begin VB.TextBox T_Long_Grad 
      Height          =   285
      Left            =   3000
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   9480
      Width           =   495
   End
   Begin VB.TextBox T_Tag 
      Height          =   285
      Left            =   480
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   8760
      Width           =   495
   End
   Begin VB.TextBox T_Monat 
      Height          =   285
      Left            =   1080
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   8760
      Width           =   495
   End
   Begin VB.TextBox T_Jahr 
      Height          =   285
      Left            =   1680
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   8760
      Width           =   495
   End
   Begin VB.TextBox T_Sekunden 
      Height          =   285
      Left            =   4200
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   8760
      Width           =   495
   End
   Begin VB.TextBox T_Minuten 
      Height          =   285
      Left            =   3600
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   8760
      Width           =   495
   End
   Begin VB.TextBox T_Stunden 
      Height          =   285
      Left            =   3000
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   8760
      Width           =   495
   End
   Begin VB.Timer Tim_Tracking 
      Interval        =   1000
      Left            =   7320
      Top             =   3720
   End
   Begin VB.Timer Tim_Simulation 
      Interval        =   100
      Left            =   6720
      Top             =   3720
   End
   Begin VB.Timer Tim_DisplayUpdate 
      Interval        =   250
      Left            =   6120
      Top             =   3720
   End
   Begin VB.CommandButton C_SetAzAlt 
      Caption         =   "Set Az Alt"
      Height          =   495
      Left            =   2880
      TabIndex        =   11
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox T_AltTel 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4320
      TabIndex        =   10
      Text            =   "0"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox T_AzTel 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4320
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   3240
      Value           =   100
      Width           =   255
   End
   Begin VB.CommandButton C_Le 
      Caption         =   "<"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton C_Dn 
      Caption         =   "V"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton C_Ri 
      Caption         =   ">"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton C_Up 
      Caption         =   "^"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   3720
      Width           =   375
   End
   Begin VB.ListBox AlignmentStarList 
      Height          =   9615
      Left            =   5880
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton C_SetEncoder 
      Caption         =   "Set Encoder"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton C_GetAlt 
      Caption         =   "Get Alt"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
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
   Begin VB.Frame Frame3 
      Caption         =   "Time"
      Height          =   1695
      Left            =   240
      TabIndex        =   49
      Top             =   1920
      Width           =   3375
      Begin VB.OptionButton O_TimeSelectSim 
         Caption         =   "Simulierte Zeit"
         Height          =   195
         Left            =   1680
         TabIndex        =   54
         Top             =   1320
         Width           =   1335
      End
      Begin VB.OptionButton O_TimeSelectLocal 
         Caption         =   "Lokalzeit"
         Height          =   195
         Left            =   240
         TabIndex        =   53
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label L_UTime 
         Caption         =   "UT"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label L_SiderialTime 
         Caption         =   "Siderial Time"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label L_LocalTime 
         Caption         =   "Local Time"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Label Label27 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   255
      Left            =   8280
      TabIndex        =   102
      Top             =   9600
      Width           =   1215
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   255
      Left            =   8280
      TabIndex        =   100
      Top             =   9240
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Skewing Speed"
      Height          =   255
      Left            =   2880
      TabIndex        =   45
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label L_SkewingSpeed 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   2880
      TabIndex        =   44
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "Sek."
      Height          =   255
      Left            =   4200
      TabIndex        =   37
      Top             =   9240
      Width           =   375
   End
   Begin VB.Label Label15 
      Caption         =   "Min."
      Height          =   255
      Left            =   3600
      TabIndex        =   36
      Top             =   9240
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "Grad"
      Height          =   255
      Left            =   3000
      TabIndex        =   35
      Top             =   9240
      Width           =   375
   End
   Begin VB.Label Label19 
      Caption         =   "Sek."
      Height          =   255
      Left            =   1680
      TabIndex        =   34
      Top             =   9240
      Width           =   375
   End
   Begin VB.Label Label18 
      Caption         =   "Min."
      Height          =   255
      Left            =   1080
      TabIndex        =   33
      Top             =   9240
      Width           =   495
   End
   Begin VB.Label Label17 
      Caption         =   "Grad"
      Height          =   255
      Left            =   480
      TabIndex        =   32
      Top             =   9240
      Width           =   375
   End
   Begin VB.Label Label12 
      Caption         =   "Jahr"
      Height          =   255
      Left            =   1680
      TabIndex        =   25
      Top             =   8520
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "Monat"
      Height          =   255
      Left            =   1080
      TabIndex        =   24
      Top             =   8520
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "Tag"
      Height          =   255
      Left            =   480
      TabIndex        =   23
      Top             =   8520
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "H"
      Height          =   255
      Left            =   3120
      TabIndex        =   22
      Top             =   8520
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "M"
      Height          =   255
      Left            =   3720
      TabIndex        =   21
      Top             =   8520
      Width           =   255
   End
   Begin VB.Label Label11 
      Caption         =   "S"
      Height          =   255
      Left            =   4320
      TabIndex        =   20
      Top             =   8520
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "Ortszeit"
      Height          =   255
      Left            =   4920
      TabIndex        =   19
      Top             =   8760
      Width           =   615
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
      TabIndex        =   12
      Top             =   10200
      Width           =   2775
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
'Dim TelDegAz As Double
'Dim TelDegAlt As Double


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



Private Enum NxMode
    Unchanged = 0
    HMS = 1
    HourDec = 2
    DegDec = 3
End Enum


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
    
    ObserverRA = HourToRad(AlignmentStarArray(idx).Ra)
    ObserverDEC = DegToRad(AlignmentStarArray(idx).Dec)
      
    RA_DEC_to_AZ_ALT_radian ObserverRA, ObserverDEC, ObserverLong, ObserverLatt, ObserverDateTimeUT, Az, Alt, HourAngle
  
'''    If O_OrientationNorth.Value Then Az = Az + Pi
'''    L_AzStar = CutAngle(RadToDeg(Az))
'''    L_AltStar = RadToDeg(Alt)
'''
'''    HourAngleHMS = RadToTime(HourAngle)
'''    L_HourAngle = HourAngleHMS.H & ":" & HourAngleHMS.M & ":" & Format(HourAngleHMS.s, "00.00")
'''
'''    T_AzTel = CutAngle(RadToDeg(Az))
'''    T_AltTel = RadToDeg(Alt)
'''

End Sub

Private Sub C_CalibrateNow_Click()
Label6 = "--"
    CalibrateTelescope Cal_InitTime, _
                       Cal_RaStar_1, Cal_DecStar_1, Cal_TelHorizAngle_1, Cal_TelElevAngle_1, Cal_Time_1, _
                       Cal_RaStar_2, Cal_DecStar_2, Cal_TelHorizAngle_2, Cal_TelElevAngle_2, Cal_Time_2, _
                       TransformationMatrix

End Sub

Private Sub C_GetAz_Click()
       
    
    If SimOffline Then
        TelIncrAz = SimIncrAz
'                L_MotorIncrSystem = TelIncrAz
'                tmp = MotorIncrSystem_to_MatrixSystem(CDbl(TelIncrAz))
'                tmp = RadToDeg(tmp)
'                L_MatrixSystem = Format(tmp, "0.0000") & "°"
    Else
        NexStarComm.Output = Chr$(&H1)
        NexStarAz = ""
        Command = 1
    End If
    
    
    L_AzMotorIncr = TelIncrAz
    MatrixSystemAzIst = MotorIncrSystem_to_MatrixSystem(CDbl(TelIncrAz))
    L_MatrixSystemAzIst = Format(RadToDeg(MatrixSystemAzIst), "0.0000") & "°"
End Sub

Private Sub C_GetAlt_Click()
    Dim tmp As Double
    
    If SimOffline Then
        TelIncrAlt = SimIncrAlt
'                L_AltMotorIncr = TelIncrAlt
'                tmp = TelIncrAlt * 360 / EncoderResolution
'                L_AltMatrixSys = Format(tmp, "0.0000") & "°"
   Else
        NexStarComm.Output = Chr$(&H15)
        NexStarAlt = ""
        Command = 21
    End If
    
    L_AltMotorIncr = TelIncrAlt
   
    MatrixSystemAltIst = TelIncrAlt * (2 * Pi) / EncoderResolution
   
    L_MatrixSystemAltIst = Format(RadToDeg(MatrixSystemAltIst), "0.0000") & "°"

End Sub




Private Sub C_GotoStar_Click()
    'Set Az
    MatrixSystemAzSoll = AzAltSystem_to_MatrixSystem(ObserverAz)
    'Set Alt
    
    MatrixSystemAltSoll = ObserverAlt + GlobalAltOffset


    Dim MotorIncrAz As Long
    Dim MotorIncrAlt As Long
    
    MotorIncrAz = MatrixSystem_to_MotorIncrSystem(MatrixSystemAzSoll)
    MotorIncrAlt = MatrixSystemAltSoll * EncoderResolution / (2 * Pi)

    SimGotoAzAltActive = True
    
    If SimOffline Then
        SimGotoAz = MotorIncrAz
        SimGotoAlt = MotorIncrAlt
    Else
        NexStarComm.Output = Chr$(&O2) & SetNexStarPosition(MotorIncrAz) & Chr$(&H16) & SetNexStarPosition(MotorIncrAlt)
    End If


End Sub

Private Sub C_GotoStarCalibrated_Click()

    Dim AimTimeRad As Double
    Dim AzAlt_BetaCet As AzAlt
    AimTimeRad = TimeToRad(ObserverTimeUT)

    CalculateTelescopeCoordinates Cal_InitTime, _
                                  ObserverRA, ObserverDEC, AimTimeRad, TransformationMatrix, _
                                  AzAlt_BetaCet

 
    'Set Az     'Muß hier scheinbar invertiert werden  -->beobachten
'    MatrixSystemAzSoll = CutRad(-AzAlt_BetaCet.Az)      'plus
'    MatrixSystemAzSoll = CutRad(Pi - AzAlt_BetaCet.Az)  'minus

     MatrixSystemAzSoll = CutRad(AzAlt_BetaCet.Az)
     
     
    'Set Alt
    MatrixSystemAltSoll = AzAlt_BetaCet.Alt


    Dim MotorIncrAz As Long
    Dim MotorIncrAlt As Long
    
    MotorIncrAz = MatrixSystem_to_MotorIncrSystem(MatrixSystemAzSoll)
    MotorIncrAlt = MatrixSystemAltSoll * EncoderResolution / (2 * Pi)
    
    LastMotorIncrAz = MotorIncrAz
    LastMotorIncrAlt = MotorIncrAlt

    SimGotoAzAltActive = True
    
    If SimOffline Then
        SimGotoAz = MotorIncrAz
        SimGotoAlt = MotorIncrAlt
    Else
        NexStarComm.Output = Chr$(&O2) & SetNexStarPosition(MotorIncrAz) & Chr$(&H16) & SetNexStarPosition(MotorIncrAlt)
    End If









End Sub

Private Sub C_Set_ObserverLocation_Click()
    
    
    ObserverLatt.Min = Zahl(T_Latt_Min)
    ObserverLatt.Sec = Zahl(T_Latt_Sek)
    ObserverLatt.Sign = T_Latt_Sign
    ObserverLong.deg = Zahl(T_Long_Grad)
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

    BacklashAlt = T_Backlash
    
    If SimOffline Then
    Else
        NexStarComm.Output = Chr$(&H1E) & SetNexStarPosition(BacklashAlt)
    End If
End Sub

Private Sub C_SetBacklAz_Click()
    Dim BacklashAz As Long    '0..100

    BacklashAz = T_Backlash
    
    If SimOffline Then
    Else
        NexStarComm.Output = Chr$(&HA) & SetNexStarPosition(BacklashAz)
    End If
End Sub

Private Sub C_SetCalibrationStar_1_Click()
   
    Cal_RaStar_1 = ObserverRA
    Cal_DecStar_1 = ObserverDEC
    Cal_TelHorizAngle_1 = MatrixSystemAzIst
    Cal_TelElevAngle_1 = MatrixSystemAltIst

    'Set time reference star 1 for calibration
    Cal_Time_1 = TimeToRad(ObserverTimeUT)

End Sub

Private Sub C_SetCalibrationStar_2_Click()
    Cal_RaStar_2 = ObserverRA
    Cal_DecStar_2 = ObserverDEC
    Cal_TelHorizAngle_2 = MatrixSystemAzIst
    Cal_TelElevAngle_2 = MatrixSystemAltIst

    'Set time reference star 2 for calibration
    Cal_Time_2 = TimeToRad(ObserverTimeUT)





End Sub

Private Sub C_SetNorth_Click()
    Dim MatrixSystem As Double
    Dim tmp As Double
    Dim d1 As Double
    Dim d2 As Double
    
    'Azimut Offset
    MatrixSystem = MotorIncrSystem_to_MatrixSystem(CDbl(TelIncrAz))
    GlobalAzOffset = CutRad(MatrixSystem)
    
    'Altitude Offset
    tmp = TelIncrAlt * (2 * Pi) / EncoderResolution
    GlobalAltOffset = CutRad(tmp)
    
    'Set Initial for calibration
    Cal_InitTime = TimeToRad(ObserverTimeUT)
    
    
    d1 = RadToDeg(MatrixSystem)
    d2 = RadToDeg(tmp)
End Sub

Private Sub C_SingleStarAlignment_Click()
    Dim t1 As Double
    Dim t2 As Double
    Dim t3 As Double
    Dim t4 As Double
    
    
    t1 = MatrixSystemAzSoll
    t2 = MatrixSystemAzIst
    
    t3 = MatrixSystemAltIst - MatrixSystemAltSoll
    GlobalAltOffset = GlobalAltOffset + t3
    
    t4 = MatrixSystemAzIst - MatrixSystemAzSoll
     GlobalAzOffset = GlobalAzOffset + t4
   
End Sub

Private Sub C_Tracking_Click()
    If TrackingisON Then
        TrackingisON = False
    Else
        TrackingisON = True
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




Private Sub Form_Load()
    SimOffline = True
    CommTest = False
    
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
    
    ObserverLatt.deg = Zahl(T_Latt_Grad)
    ObserverLatt.Min = Zahl(T_Latt_Min)
    ObserverLatt.Sec = Zahl(T_Latt_Sek)
    ObserverLatt.Sign = T_Latt_Sign
    ObserverLong.deg = Zahl(T_Long_Grad)
    ObserverLong.Min = Zahl(T_Long_Min)
    ObserverLong.Sec = Zahl(T_Long_Sek)
    ObserverLong.Sign = T_Long_Sign
    
    T_Tag = INIGetValue(IniFileName, "Datum", "Tag")
    T_Monat = INIGetValue(IniFileName, "Datum", "Monat")
    T_Jahr = INIGetValue(IniFileName, "Datum", "Jahr")
    
    T_Stunden = INIGetValue(IniFileName, "Zeit", "Stunden")
    T_Minuten = INIGetValue(IniFileName, "Zeit", "Minuten")
    T_Sekunden = INIGetValue(IniFileName, "Zeit", "Sekunden")
        
    Cal_InitTime = Zahl(INIGetValue(IniFileName, "TransformationMatrix", "Cal_InitTime"))
    TransformationMatrix(0, 0) = Zahl(INIGetValue(IniFileName, "TransformationMatrix", "00"))
    TransformationMatrix(0, 1) = Zahl(INIGetValue(IniFileName, "TransformationMatrix", "01"))
    TransformationMatrix(0, 2) = Zahl(INIGetValue(IniFileName, "TransformationMatrix", "02"))
    TransformationMatrix(1, 0) = Zahl(INIGetValue(IniFileName, "TransformationMatrix", "10"))
    TransformationMatrix(1, 1) = Zahl(INIGetValue(IniFileName, "TransformationMatrix", "11"))
    TransformationMatrix(1, 2) = Zahl(INIGetValue(IniFileName, "TransformationMatrix", "12"))
    TransformationMatrix(2, 0) = Zahl(INIGetValue(IniFileName, "TransformationMatrix", "20"))
    TransformationMatrix(2, 1) = Zahl(INIGetValue(IniFileName, "TransformationMatrix", "21"))
    TransformationMatrix(2, 2) = Zahl(INIGetValue(IniFileName, "TransformationMatrix", "22"))
    
    
    
    
    

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
    NexStarComm.InputLen = 1
    NexStarComm.DTREnable = False
    NexStarComm.RThreshold = 1
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

' Goto AzAlt        0xO2 Az (3 Byte) 0x16 Alt (3 Bype)
' Move UP           0x06 0 (3 Byte) 0x1A & Speed  (3 Bype)
' Move DOWN         0x06 0 (3 Byte) 0x1B & Speed  (3 Bype)
' Move LEFT         0x07 Speed (3 Byte) 0x1A 0 (3 Bype)
' Move RIGHT        0x06 Speed (3 Byte) 0x1A 0 (3 Bype)
' Set EncRes        0x0C EncResAz (3 Byte) EncResAlt (3 Bype)
' Set Az Backlash   0x0A BacklashAz (3 Byte)
' Set Alt Backlash  0x1E BacklashAlt (3 Byte)
' Get Az Incr       0x01                            Antwort Az (3 Byte)
' Get Alt Incr      0x15                            Antwort Az (3 Byte)

' Skewing rate      [1/10 Motor Incr/sec]  i.e.  Skewing rate 10000: 10000 Incr in 10sec





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
            ElseIf TestCommMotorToHandheld Then
                NexStarChar1 = ""
                Do
                    vbuf = NexStarComm.Input
                    bbuf = vbuf
                    NexStarChar1 = NexStarChar1 & Chr$(bbuf(0))
                     key = (bbuf(0))
                Loop While NexStarComm.InBufferCount > 0
                l = Len(NexStarChar1)
                Communication.DisplayAzAltTracking NexStarChar1
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
    
    If Not CommTest Then
    
        If Toggle Then
            Toggle = False
            C_GetAz_Click
        Else
            Toggle = True
            C_GetAlt_Click
        End If

    End If
    
    L_GlobalAzOffset = Format(RadToDeg(GlobalAzOffset), "0.0000") & "°"
    L_GlobalAltOffset = Format(RadToDeg(GlobalAltOffset), "0.0000") & "°"
    L_MatrixSystemAzSoll = Format(RadToDeg(MatrixSystemAzSoll), "0.0000") & "°"
    L_MatrixSystemAltSoll = Format(RadToDeg(MatrixSystemAltSoll), "0.0000") & "°"
    
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
'    Dim tTime As MyTime
    Dim tDate As MyDate
    Dim tTs As MyTime
    Dim tTsRad As Double
    Dim LongitudeDeg As Double
    Dim LongitudeRad As Double
  
    

    If O_TimeSelectLocal.Value = True Then
        ObserverDateTimeUT = UtcTime(Now)              ' Get current Time UT
        L_LocalTime = " Local time:   " & Now
    Else
        ' Take simulation time
        ObserverDateTimeUT = StingsToDate(T_Tag, T_Monat, T_Jahr, T_Stunden, T_Minuten, T_Sekunden)
        L_LocalTime = " Local time:   " & "--"
    End If

    L_UTime = " UT:              " & ObserverDateTimeUT
    ObserverTimeUT.H = Hour(ObserverDateTimeUT)
    ObserverTimeUT.M = Minute(ObserverDateTimeUT)
    ObserverTimeUT.s = Second(ObserverDateTimeUT)
    tDate.YY = Year(ObserverDateTimeUT)
    tDate.MM = Month(ObserverDateTimeUT)
    tDate.DD = Day(ObserverDateTimeUT)
    
    LongitudeDeg = GeoToDez(ObserverLong)
    LongitudeRad = DegToRad(LongitudeDeg)
    
    'double check siderial time: https://tycho.usno.navy.mil/sidereal.html
    tTsRad = TimeToRad(GMST(tDate, ObserverTimeUT)) - LongitudeRad
    tTs = RadToTime(tTsRad)
    L_SiderialTime = "Siderial time: " & tTs.H & ":" & Format(tTs.M, "00") & ":" & Format(tTs.s, "00")




    Dim idx As Long
    Dim Az As Double
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
    F_StarInfo.Caption = AlignmentStarArray(idx).ProperName
    
    ObserverRA = HourToRad(AlignmentStarArray(idx).Ra)
    ObserverDEC = DegToRad(AlignmentStarArray(idx).Dec)
    DisplayCoordinate L_I_RA, ObserverRA, HMS
    DisplayCoordinate L_I_DEC, ObserverDEC, DegDec
     
    RA_DEC_to_AZ_ALT_radian ObserverRA, ObserverDEC, ObserverLong, ObserverLatt, ObserverDateTimeUT, ObserverAz, ObserverAlt, HourAngle

    Dim DisplObserverAz As Double
    If Ch_South.Value = 1 Then
        DisplObserverAz = CutRad(ObserverAz + Pi)
    Else
        DisplObserverAz = CutRad(ObserverAz)
    End If
    
    DisplayCoordinate L_I_Az, DisplObserverAz, DegDec
    DisplayCoordinate L_I_Alt, ObserverAlt, DegDec
    DisplayCoordinate L_I_HourAngle, HourAngle, HMS
    T_AzTel = RadToDeg(ObserverAz)
    T_AltTel = RadToDeg(ObserverAlt)

            'Just for testing: get matrix vectors
            Dim x As Double
            Dim Y As Double
            Dim z As Double
            Dim HorizAngle As Double
            Dim ElevAngle As Double
            
            HorizAngle = ObserverRA
            ElevAngle = ObserverDEC
            x = Cos(ElevAngle) * Cos(HorizAngle)
            Y = Cos(ElevAngle) * Sin(HorizAngle)
            z = Sin(ElevAngle)
            L_I_EquXYZ = Format(x, "0.0000") & " " & Format(Y, "0.0000") & " " & Format(z, "0.0000")
        
            HorizAngle = ObserverAz
            ElevAngle = ObserverAlt
            x = Cos(ElevAngle) * Cos(HorizAngle)
            Y = Cos(ElevAngle) * Sin(HorizAngle)
            z = Sin(ElevAngle)
            L_I_HorXYZ = Format(x, "0.0000") & " " & Format(Y, "0.0000") & " " & Format(z, "0.0000")


    If ObserverAlt < 0 Then
        L_CurrentStar.BackColor = RGB(255, 0, 0)
    ElseIf (ObserverAlt > 0) And (ObserverAlt < 0.3) Then
        L_CurrentStar.BackColor = RGB(255, 255, 0)
    Else
        L_CurrentStar.BackColor = RGB(0, 255, 0)
    End If
    
    
    
    
    
    Static TrackCount As Long
    Const TrackInterval = 10        'calculate new star positition ever ... sec
    Dim n As Long
    
    n = (TrackInterval * 1000) / Tim_Tracking.Interval

    If TrackingisON Then
            TrackCount = TrackCount + 1
            If TrackCount >= n Then
                TrackCount = 0
                
                C_Tracking.BackColor = RGB(0, 255, 0)
            
            
                Dim AimTimeRad As Double
                Dim AzAlt_BetaCet As AzAlt
                AimTimeRad = TimeToRad(ObserverTimeUT)
            
                CalculateTelescopeCoordinates Cal_InitTime, _
                                              ObserverRA, ObserverDEC, AimTimeRad, TransformationMatrix, _
                                              AzAlt_BetaCet
    
                'Set Az
                MatrixSystemAzSoll = CutRad(AzAlt_BetaCet.Az)
                 
                'Set Alt
                MatrixSystemAltSoll = AzAlt_BetaCet.Alt
            
                Dim MotorIncrAz As Long
                Dim MotorIncrAlt As Long
                Dim DiffMotorIncrAz As Long
                Dim DiffMotorIncrAlt As Long
                
                MotorIncrAz = MatrixSystem_to_MotorIncrSystem(MatrixSystemAzSoll)
                MotorIncrAlt = MatrixSystemAltSoll * EncoderResolution / (2 * Pi)
                DiffMotorIncrAz = MotorIncrAz - LastMotorIncrAz
                DiffMotorIncrAlt = MotorIncrAlt - LastMotorIncrAlt
                LastMotorIncrAz = MotorIncrAz
                LastMotorIncrAlt = MotorIncrAlt
            
                Label6 = DiffMotorIncrAz
                Label27 = DiffMotorIncrAlt
    
''                SimGotoAzAltActive = True
''
''                If SimOffline Then
''                    SimGotoAz = MotorIncrAz
''                    SimGotoAlt = MotorIncrAlt
''                Else
''                    NexStarComm.Output = Chr$(&O2) & SetNexStarPosition(MotorIncrAz) & Chr$(&H16) & SetNexStarPosition(MotorIncrAlt)
''                End If
                
                
            End If



    
    
    
    
    
    
    Else
            C_Tracking.BackColor = &H8000000F
    End If

    
    
    
    
    
    

End Sub

Private Sub VS_ManualSkewingSpeed_Change()
    Dim tmp As Long
    
    tmp = VS_ManualSkewingSpeed.Value
    ManualSkewingSpeed = 1000 * tmp
    L_SkewingSpeed = ManualSkewingSpeed
    
    'SkewingSpeed[Incr/sec] = ManualSkewingSpeed[Incr/sec] * 0,1 [Incr/sec]
    '
    
    
End Sub

'Coordinate [radian]
Private Sub DisplayCoordinate(l As Label, Coord As Double, Mode As NxMode)
    Dim TmpTime As MyTime
    
    If Mode = HMS Then
        TmpTime = RadToTime(Coord)
        l = TmpTime.H & ":" & TmpTime.M & ":" & Format(TmpTime.s, "00.00")
    ElseIf Mode = DegDec Then
        l = Format(RadToDeg(Coord), "0.0000") & "°"
    Else
        l = Coord
    End If
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
        AlignmentStarArray(idx).Ra = Zahl(StarEntities(5))
        AlignmentStarArray(idx).Dec = Zahl(StarEntities(6))
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








