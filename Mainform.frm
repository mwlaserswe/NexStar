VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form Mainform 
   Caption         =   "Form1"
   ClientHeight    =   11235
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14250
   LinkTopic       =   "Form1"
   ScaleHeight     =   11235
   ScaleWidth      =   14250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton C_Simulation 
      Caption         =   "Simulation"
      Height          =   495
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   113
      Top             =   1200
      Width           =   1215
   End
   Begin Project1.zzSlider zzSlider1 
      Height          =   255
      Left            =   8040
      TabIndex        =   112
      Top             =   9360
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SliderColor     =   12632319
      MaxValue        =   1000
      MinValue        =   -1000
      Value           =   300
      SmallChange     =   0
      LargeChange     =   0
   End
   Begin VB.Timer Tim_Startup 
      Interval        =   500
      Left            =   6480
      Top             =   5400
   End
   Begin VB.CommandButton C_GotoNorth 
      Caption         =   "Goto Noth"
      Height          =   495
      Left            =   4080
      TabIndex        =   111
      Top             =   360
      Width           =   1335
   End
   Begin VB.Timer Tim_TestStatus 
      Interval        =   500
      Left            =   6240
      Top             =   4440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   12360
      TabIndex        =   109
      Top             =   7560
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   5910
      Left            =   12240
      TabIndex        =   108
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton C_SetEncoder_Alt 
      Caption         =   "Set Enc. Alt"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   480
      Width           =   1215
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   315
      Left            =   7920
      TabIndex        =   105
      Top             =   8880
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   556
      _Version        =   393216
      Min             =   -1000
      Max             =   1000
      SelectRange     =   -1  'True
   End
   Begin VB.CommandButton C_Tracking 
      Caption         =   "Tracking"
      Height          =   495
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton C_SingleStarAlignment 
      Caption         =   "Single Star Alignment"
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton C_GotoStarCalibrated 
      Caption         =   "GotoStar calibrated"
      Height          =   255
      Left            =   8640
      TabIndex        =   91
      Top             =   10800
      Width           =   2175
   End
   Begin VB.CommandButton C_CalibrateNow 
      Caption         =   "Calibrate now"
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   7560
      Width           =   2295
   End
   Begin VB.CommandButton C_SetCalibrationStar_2 
      Caption         =   "Set Calibration Star 2"
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton C_SetCalibrationStar_1 
      Caption         =   "Set Calibration Star 1"
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   6360
      Width           =   2295
   End
   Begin VB.CommandButton C_GotoStar 
      Caption         =   "GotoStar w/o calibration"
      Height          =   255
      Left            =   5880
      TabIndex        =   87
      Top             =   10800
      Width           =   2175
   End
   Begin VB.CommandButton C_SetNorth 
      Caption         =   "Set North"
      Height          =   495
      Left            =   120
      TabIndex        =   86
      Top             =   5760
      Width           =   2295
   End
   Begin VB.Frame Frame4 
      Caption         =   "Altitude"
      Height          =   1695
      Left            =   8040
      TabIndex        =   75
      Top             =   4680
      Width           =   3855
      Begin VB.Label L_MotorSystemAltDiff 
         Caption         =   "--"
         Height          =   255
         Left            =   2520
         TabIndex        =   103
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label L_MatrixSystemAltDiff 
         Caption         =   "--"
         Height          =   255
         Left            =   2520
         TabIndex        =   101
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label46 
         Caption         =   "Mot. Incr."
         Height          =   255
         Left            =   240
         TabIndex        =   85
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label L_AltMotorIncr 
         Caption         =   "--"
         Height          =   255
         Left            =   1200
         TabIndex        =   84
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label L_MatrixSystemAltIst 
         Caption         =   "--"
         Height          =   255
         Left            =   1200
         TabIndex        =   83
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label43 
         Caption         =   "Matr Sy Ist:"
         Height          =   255
         Left            =   240
         TabIndex        =   82
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label42 
         Caption         =   "AzAlt Sys:"
         Height          =   255
         Left            =   240
         TabIndex        =   81
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label41 
         Caption         =   "--"
         Height          =   255
         Left            =   1200
         TabIndex        =   80
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label L_MatrixSystemAltSoll 
         Caption         =   "--"
         Height          =   255
         Left            =   1200
         TabIndex        =   79
         Top             =   600
         Width           =   975
      End
      Begin VB.Label L_GlobalAltOffset 
         Caption         =   "--"
         Height          =   255
         Left            =   1200
         TabIndex        =   78
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label36 
         Caption         =   "Matr Soll"
         Height          =   255
         Left            =   240
         TabIndex        =   77
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Glob. Offs."
         Height          =   255
         Left            =   240
         TabIndex        =   76
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Azimut"
      Height          =   1695
      Left            =   8040
      TabIndex        =   64
      Top             =   2880
      Width           =   3855
      Begin VB.Label L_MotorSystemAzDiff 
         Caption         =   "--"
         Height          =   255
         Left            =   2640
         TabIndex        =   102
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label L_MatrixSystemAzDiff 
         Caption         =   "--"
         Height          =   255
         Left            =   2520
         TabIndex        =   100
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label31 
         Caption         =   "Glob. Offs."
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label32 
         Caption         =   "Matr Soll"
         Height          =   255
         Left            =   240
         TabIndex        =   73
         Top             =   600
         Width           =   855
      End
      Begin VB.Label L_GlobalAzOffset 
         Caption         =   "--"
         Height          =   255
         Left            =   1200
         TabIndex        =   72
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label L_MatrixSystemAzSoll 
         Caption         =   "--"
         Height          =   255
         Left            =   1200
         TabIndex        =   71
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label L_AzAltSystem 
         Caption         =   "--"
         Height          =   255
         Left            =   1200
         TabIndex        =   70
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label26 
         Caption         =   "AzAlt Sys:"
         Height          =   255
         Left            =   240
         TabIndex        =   69
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "Matr Sy Ist:"
         Height          =   255
         Left            =   240
         TabIndex        =   68
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label L_MatrixSystemAzIst 
         Caption         =   "--"
         Height          =   255
         Left            =   1200
         TabIndex        =   67
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label L_AzMotorIncr 
         Caption         =   "--"
         Height          =   255
         Left            =   1200
         TabIndex        =   66
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Mot. Incr."
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Frame F_StarInfo 
      Caption         =   "--"
      Height          =   2415
      Left            =   8040
      TabIndex        =   49
      Top             =   360
      Width           =   3855
      Begin VB.CheckBox Ch_South 
         Caption         =   "South (VSky)"
         Height          =   255
         Left            =   2400
         TabIndex        =   92
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label L_CardinalOrientation 
         Caption         =   "--"
         Height          =   255
         Left            =   2160
         TabIndex        =   99
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label25 
         Caption         =   "Horiz. xyz:"
         Height          =   255
         Left            =   240
         TabIndex        =   63
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label L_I_HorXYZ 
         Caption         =   "Alt:"
         Height          =   255
         Left            =   1200
         TabIndex        =   62
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label L_I_EquXYZ 
         Caption         =   "Alt:"
         Height          =   255
         Left            =   1200
         TabIndex        =   61
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label22 
         Caption         =   "Equ. xyz:"
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "Hour Angle:"
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label L_I_HourAngle 
         Caption         =   "Alt:"
         Height          =   255
         Left            =   1200
         TabIndex        =   58
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label L_I_Alt 
         Caption         =   "Alt:"
         Height          =   255
         Left            =   1200
         TabIndex        =   57
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label L_I_Az 
         Caption         =   "Az:"
         Height          =   255
         Left            =   1200
         TabIndex        =   56
         Top             =   840
         Width           =   735
      End
      Begin VB.Label L_I_DEC 
         Caption         =   "DEC:"
         Height          =   255
         Left            =   1200
         TabIndex        =   55
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label L_I_RA 
         Caption         =   "RA:"
         Height          =   255
         Left            =   1200
         TabIndex        =   54
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "Alt:"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "Az:"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "DEC:"
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "RA:"
         Height          =   255
         Left            =   240
         TabIndex        =   50
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
      Left            =   360
      TabIndex        =   42
      Text            =   "40"
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton C_SetBacklAlt 
      Caption         =   "Set Backl. Alt"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton C_SetBacklAz 
      Caption         =   "Set Backl. Az"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox T_Long_Sign 
      Height          =   285
      Left            =   2640
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   9480
      Width           =   255
   End
   Begin VB.CommandButton C_Set_ObserverLocation 
      Caption         =   "Command7"
      Height          =   255
      Left            =   2160
      TabIndex        =   36
      Top             =   10080
      Width           =   1215
   End
   Begin VB.TextBox T_Latt_Sign 
      Height          =   285
      Left            =   120
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   9480
      Width           =   255
   End
   Begin VB.TextBox T_Latt_Grad 
      Height          =   285
      Left            =   480
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   9480
      Width           =   495
   End
   Begin VB.TextBox T_Latt_Min 
      Height          =   285
      Left            =   1080
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   9480
      Width           =   495
   End
   Begin VB.TextBox T_Latt_Sek 
      Height          =   285
      Left            =   1680
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   9480
      Width           =   495
   End
   Begin VB.TextBox T_Long_Sek 
      Height          =   285
      Left            =   4200
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   9480
      Width           =   495
   End
   Begin VB.TextBox T_Long_Min 
      Height          =   285
      Left            =   3600
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   9480
      Width           =   495
   End
   Begin VB.TextBox T_Long_Grad 
      Height          =   285
      Left            =   3000
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   9480
      Width           =   495
   End
   Begin VB.TextBox T_Tag 
      Height          =   285
      Left            =   480
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   8760
      Width           =   495
   End
   Begin VB.TextBox T_Monat 
      Height          =   285
      Left            =   1080
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   8760
      Width           =   495
   End
   Begin VB.TextBox T_Jahr 
      Height          =   285
      Left            =   1680
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   8760
      Width           =   495
   End
   Begin VB.TextBox T_Sekunden 
      Height          =   285
      Left            =   4200
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   8760
      Width           =   495
   End
   Begin VB.TextBox T_Minuten 
      Height          =   285
      Left            =   3600
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   8760
      Width           =   495
   End
   Begin VB.TextBox T_Stunden 
      Height          =   285
      Left            =   3000
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   8760
      Width           =   495
   End
   Begin VB.Timer Tim_Tracking 
      Interval        =   500
      Left            =   7320
      Top             =   3720
   End
   Begin VB.Timer Tim_Simulation 
      Interval        =   500
      Left            =   6720
      Top             =   3720
   End
   Begin VB.Timer Tim_DisplayUpdate 
      Interval        =   250
      Left            =   6120
      Top             =   3720
   End
   Begin VB.VScrollBar VS_ManualSlewingSpeed 
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
   Begin VB.CommandButton C_SetEncoder_Az 
      Caption         =   "Set Enc. Az"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
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
      TabIndex        =   43
      Top             =   1920
      Width           =   3375
      Begin VB.OptionButton O_TimeSelectSim 
         Caption         =   "Simulierte Zeit"
         Height          =   195
         Left            =   1680
         TabIndex        =   48
         Top             =   1320
         Width           =   1335
      End
      Begin VB.OptionButton O_TimeSelectLocal 
         Caption         =   "Lokalzeit"
         Height          =   195
         Left            =   240
         TabIndex        =   47
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label L_UTime 
         Caption         =   "UT"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label L_SiderialTime 
         Caption         =   "Siderial Time"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label L_LocalTime 
         Caption         =   "Local Time"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Label L_StatusMoving 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   375
      Left            =   9240
      TabIndex        =   110
      Top             =   9840
      Width           =   2415
   End
   Begin VB.Label L_ErrorCount 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ErrorCount:"
      Height          =   255
      Left            =   12360
      TabIndex        =   107
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   10560
      X2              =   10560
      Y1              =   8640
      Y2              =   9360
   End
   Begin VB.Label L_MotorSystemAzDiffSim 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   255
      Left            =   8040
      TabIndex        =   104
      Top             =   8280
      Width           =   2415
   End
   Begin VB.Label L_MotorSystemAzDiffReal 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   255
      Left            =   8040
      TabIndex        =   98
      Top             =   7920
      Width           =   2415
   End
   Begin VB.Label L_MatrixSystemAzDiffReal 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   255
      Left            =   8040
      TabIndex        =   97
      Top             =   7560
      Width           =   2415
   End
   Begin VB.Label Label27 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   255
      Left            =   8040
      TabIndex        =   96
      Top             =   7080
      Width           =   2415
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   255
      Left            =   8040
      TabIndex        =   94
      Top             =   6720
      Width           =   2415
   End
   Begin VB.Label Label13 
      Caption         =   "Slewing Speed"
      Height          =   255
      Left            =   2880
      TabIndex        =   39
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label L_SlewingSpeed 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   2880
      TabIndex        =   38
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "Sek."
      Height          =   255
      Left            =   4200
      TabIndex        =   34
      Top             =   9240
      Width           =   375
   End
   Begin VB.Label Label15 
      Caption         =   "Min."
      Height          =   255
      Left            =   3600
      TabIndex        =   33
      Top             =   9240
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "Grad"
      Height          =   255
      Left            =   3000
      TabIndex        =   32
      Top             =   9240
      Width           =   375
   End
   Begin VB.Label Label19 
      Caption         =   "Sek."
      Height          =   255
      Left            =   1680
      TabIndex        =   31
      Top             =   9240
      Width           =   375
   End
   Begin VB.Label Label18 
      Caption         =   "Min."
      Height          =   255
      Left            =   1080
      TabIndex        =   30
      Top             =   9240
      Width           =   495
   End
   Begin VB.Label Label17 
      Caption         =   "Grad"
      Height          =   255
      Left            =   480
      TabIndex        =   29
      Top             =   9240
      Width           =   375
   End
   Begin VB.Label Label12 
      Caption         =   "Jahr"
      Height          =   255
      Left            =   1680
      TabIndex        =   22
      Top             =   8520
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "Monat"
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      Top             =   8520
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "Tag"
      Height          =   255
      Left            =   480
      TabIndex        =   20
      Top             =   8520
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "H"
      Height          =   255
      Left            =   3120
      TabIndex        =   19
      Top             =   8520
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "M"
      Height          =   255
      Left            =   3720
      TabIndex        =   18
      Top             =   8520
      Width           =   255
   End
   Begin VB.Label Label11 
      Caption         =   "S"
      Height          =   255
      Left            =   4320
      TabIndex        =   17
      Top             =   8520
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "Ortszeit"
      Height          =   255
      Left            =   4920
      TabIndex        =   16
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
      TabIndex        =   9
      Top             =   10200
      Width           =   2775
   End
   Begin VB.Menu M_Setup 
      Caption         =   "Setup"
   End
   Begin VB.Menu M_Test 
      Caption         =   "Test"
      Begin VB.Menu M_Test1 
         Caption         =   "Test1"
      End
      Begin VB.Menu M_Test2 
         Caption         =   "Test2"
      End
      Begin VB.Menu M_Vis 
         Caption         =   "Visualisierung"
      End
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
Dim TelIncr As AzAlt

Dim ManualSlewingSpeed As Double

'Simulation
Dim SimIncr As AzAlt
                
Dim SimBntUp As Boolean
Dim SimBntDn As Boolean
Dim SimBntLe As Boolean
Dim SimBntRi As Boolean
Dim SimGotoAzAltActive As Boolean
'''Dim SimTrackingActive As Boolean
Dim SimGoto As AzAlt
Dim SimTrackingStep As AzAlt


Private Enum NxMode
    Unchanged = 0
    HMS = 1
    HourDec = 2
    DegDec = 3
End Enum


Private Sub AlignmentStarList_Click()
    Dim idx As Long
    Dim HourAngle As Double
    Dim HourAngleHMS As MyTime
    Dim SelectedStar As String

    ' Search star name in list
    idx = -1
    Do
        idx = idx + 1
    Loop Until (AlignmentStarArray(idx).ProperName = AlignmentStarList.Text) Or (idx >= UBound(AlignmentStarArray))
    L_CurrentStar = AlignmentStarArray(idx).ProperName
    
    ObserverRaDec.Ra = HourToRad(AlignmentStarArray(idx).Ra)
    ObserverRaDec.Dec = DegToRad(AlignmentStarArray(idx).Dec)
         
         
    Dim StarPos As AzAlt
    StarPos = RA_DEC_to_AZ_ALT_new(ObserverRaDec, GlbOberverPos, GlbSiderialTime)
'    StarPos = AzAlt_to_MatrixSystem(StarPos)
    DispAlignmentStar StarPos
  
  
End Sub



Private Sub C_GetAz_Click()
       
    
    If SimOffline Then
        TelIncr = SimIncr
    Else
        NexStarComm.Output = Chr$(&H1)
        NexStarAz = ""
        Command = 1
    End If
    
End Sub


Private Sub C_GetAlt_Click()
    Dim tmp As Double
    
    If SimOffline Then
        TelIncr = SimIncr
    Else
        NexStarComm.Output = Chr$(&H15)
        NexStarAlt = ""
        Command = 21
    End If
    
End Sub


Private Sub C_GotoNorth_Click()
    TestStatus = True
    StatusMoving = 0
    
    Dim MotorIncr As AzAlt
    
    MotorIncr = Matrix_To_MotorIncrSystem(GlobalOffset)
    
    
    If SimOffline Then
        SimGotoAzAltActive = True
        SimGoto = MotorIncr
    Else
        NexStarComm.Output = Chr$(&O2) & SetNexStarPosition(CLng(MotorIncr.Az)) & Chr$(&H16) & SetNexStarPosition(CLng(MotorIncr.Alt))
    End If

End Sub

Private Sub C_GotoStar_Click()


    TestStatus = True
    StatusMoving = 0


    Dim tmp As AzAlt
    tmp.Az = ObserverAzAlt.Az
    tmp.Alt = ObserverAzAlt.Alt
    MatrixSystemSoll = AzAlt_to_MatrixSystem(tmp)

    Dim MotorIncr As AzAlt
    

    MotorIncr = Matrix_To_MotorIncrSystem(MatrixSystemSoll)
    
    If SimOffline Then
        SimGotoAzAltActive = True
        SimGoto = MotorIncr
    Else
        NexStarComm.Output = Chr$(&O2) & SetNexStarPosition(CLng(MotorIncr.Az)) & Chr$(&H16) & SetNexStarPosition(CLng(MotorIncr.Alt))
    End If


End Sub

Private Sub C_GotoStarCalibrated_Click()

    TestStatus = True
    StatusMoving = 0
    

    Dim AimTimeRad As Double
    Dim AzAlt_BetaCet As AzAlt
    AimTimeRad = TimeToRad(ObserverTimeUT)

    TrackingisON = False

    CalculateTelescopeCoordinates Cal_InitTime, _
                                  ObserverRaDec.Ra, ObserverRaDec.Dec, AimTimeRad, TransformationMatrix, _
                                  AzAlt_BetaCet

 
    'Set Az
    MatrixSystemSoll.Az = CutRad(AzAlt_BetaCet.Az)
    'Set Alt
    MatrixSystemSoll.Alt = AzAlt_BetaCet.Alt
    
    MatrixLastCalc = MatrixSystemSoll


    Dim MotorIncr As AzAlt
    MotorIncr = Matrix_To_MotorIncrSystem(MatrixSystemSoll)
    MotorLastCalc = MotorIncr



    
    If SimOffline Then
        SimGotoAzAltActive = True
        SimGoto = MotorIncr
    Else
        NexStarComm.Output = Chr$(&O2) & SetNexStarPosition(CLng(MotorIncr.Az)) & Chr$(&H16) & SetNexStarPosition(CLng(MotorIncr.Alt))
    End If


End Sub

Private Sub C_Set_ObserverLocation_Click()
    
    
    ObserverLatt.deg = Zahl(T_Latt_Grad)
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
   
    Cal_RaStar_1 = ObserverRaDec.Ra
    Cal_DecStar_1 = ObserverRaDec.Dec
    Cal_TelHorizAngle_1 = MatrixSystemIst.Az
    Cal_TelElevAngle_1 = MatrixSystemIst.Alt

    'Set time reference star 1 for calibration
    Cal_Time_1 = TimeToRad(ObserverTimeUT)

    C_SetCalibrationStar_1.BackColor = vbGreen
End Sub

Private Sub C_SetCalibrationStar_2_Click()
    Cal_RaStar_2 = ObserverRaDec.Ra
    Cal_DecStar_2 = ObserverRaDec.Dec
    Cal_TelHorizAngle_2 = MatrixSystemIst.Az
    Cal_TelElevAngle_2 = MatrixSystemIst.Alt

    'Set time reference star 2 for calibration
    Cal_Time_2 = TimeToRad(ObserverTimeUT)

    C_SetCalibrationStar_2.BackColor = vbGreen
End Sub

Private Sub C_CalibrateNow_Click()
Label6 = "--"
    CalibrateTelescope Cal_InitTime, _
                       Cal_RaStar_1, Cal_DecStar_1, Cal_TelHorizAngle_1, Cal_TelElevAngle_1, Cal_Time_1, _
                       Cal_RaStar_2, Cal_DecStar_2, Cal_TelHorizAngle_2, Cal_TelElevAngle_2, Cal_Time_2, _
                       TransformationMatrix

    C_CalibrateNow.BackColor = vbGreen
End Sub

Private Sub C_SetEncoder_Alt_Click()
    If SimOffline Then
    Else
        NexStarComm.Output = Chr$(&H1F) & SetNexStarPosition(EncoderResolution)
    End If
End Sub

Private Sub C_SetEncoder_Az_Click()
    If SimOffline Then
    Else
    
        WriteComm " Encoder Rsolution Az:   0x0C " & EncoderResolution, Send
        NexStarComm.Output = Chr$(&HC) & SetNexStarPosition(EncoderResolution)
    End If
End Sub

Private Sub C_Simulation_Click()
    SimOffline = Not SimOffline
End Sub

Private Sub C_SingleStarAlignment_Click()
''''    Dim MatrixSystem As AzAlt
''''    Dim tmp As Double
''''    Dim d1 As Double
''''    Dim d2 As Double
''''
''''    MatrixSystem = MotorIncr_To_MatrixSystem(TelIncr)
''''    GlobalOffset.Az = CutRad(MatrixSystem.Az)
''''    GlobalOffset.Alt = CutRad(MatrixSystem.Alt)
    
    
    
'    GlobalOffset.Az = GlobalOffset.Az + (MatrixSystemIst.Az - MatrixSystemSoll.Az)
'    GlobalOffset.Alt = GlobalOffset.Alt + (MatrixSystemIst.Alt - MatrixSystemSoll.Alt)

    GlobalOffset = AddAzAlt(GlobalOffset, SubAzAlt(MatrixSystemIst, MatrixSystemSoll))

    ' Set status 1 point calibration done
    GlbCalibStatus = 1
  
    'Set Initial for calibration
    Cal_InitTime = TimeToRad(ObserverTimeUT)
    
    C_SingleStarAlignment.BackColor = vbGreen
    
End Sub


'''Private Sub C_SingleStarAlignment_Click()
'''    GlobalOffset.Az = GlobalOffset.Az + (MatrixSystemIst.Az - MatrixSystemSoll.Az)
'''    GlobalOffset.Alt = GlobalOffset.Alt + (MatrixSystemIst.Alt - MatrixSystemSoll.Alt)
'''End Sub


Private Sub C_Tracking_Click()
    If TrackingisON Then
        TrackingisON = False
        
        If SimOffline Then
        Else
            NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1A) & SetNexStarPosition(0)
        End If
    Else
        TrackingisON = True
    End If
End Sub


Private Sub C_Up_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SimOffline Then
        SimBntUp = True
    Else
        NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1A) & SetNexStarPosition(CDbl(ManualSlewingSpeed))
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
        NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1B) & SetNexStarPosition(CDbl(ManualSlewingSpeed))
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
        NexStarComm.Output = Chr$(&H7) & SetNexStarPosition(CDbl(ManualSlewingSpeed)) & Chr$(&H1A) & SetNexStarPosition(0)
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
        NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(CDbl(ManualSlewingSpeed)) & Chr$(&H1A) & SetNexStarPosition(0)
    End If
End Sub

Private Sub C_Ri_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SimOffline Then
        SimBntRi = False
    Else
        NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1A) & SetNexStarPosition(0)
    End If
End Sub


Private Sub Command1_Click()
    Command = 13
    NexStarComm.Output = Chr$(&HD)
End Sub


Private Sub Form_Load()
    SimOffline = True
    CommTest = False
    
    O_TimeSelectLocal.Value = 1
    IniFileName = App.Path & "\NexStar.ini"
    InitNexStarComm
    
    Command = 0
    
    VS_ManualSlewingSpeed.Value = 10


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
    

    
    Dim nw As String
    Dim CommFile As Integer
    nw = Now
     
'    CommFileName = App.Path & "\Commu_" & Year(nw) & Month(nw) & Day(nw) & Hour(nw) & Minute(nw) & Second(nw) & ".txt"
    CommFileName = App.Path & "\Commu_.txt"
    CommFile = FreeFile
    On Error GoTo openErr:
    Open CommFileName For Output As CommFile
    Close CommFile
    
    Vis.Show
    
    
    Exit Sub
  
openErr:
    MsgBox CommFileName & vbCr & Err.Description, , "NexStar"
    
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



Private Sub M_Test1_Click()
    Test.Show
End Sub

Private Sub M_Test2_Click()
    Test2.Show
End Sub

Private Sub M_Vis_Click()
    Vis.Show
End Sub

' Goto AzAlt        0xO2 Az (3 Byte) 0x16 Alt (3 Bype)
' Move UP           0x06 0 (3 Byte) 0x1A & Speed  (3 Bype)
' Move DOWN         0x06 0 (3 Byte) 0x1B & Speed  (3 Bype)
' Move LEFT         0x07 Speed (3 Byte) 0x1A 0 (3 Bype)
' Move RIGHT        0x06 Speed (3 Byte) 0x1A 0 (3 Bype)
' Set Az EncRes     0x0C EncResAz (3 Byte)
' Set Alt EncRes    0x1F EncResAlt (3 Bype)
' Set Az Backlash   0x0A BacklashAz (3 Byte)
' Set Alt Backlash  0x1E BacklashAlt (3 Byte)
' Get Az Incr       0x01                            Antwort Az (3 Byte)
' Get Alt Incr      0x15                            Antwort Az (3 Byte)
' ????              0x18 ??? (3 Byte)

' Get Status        0x0D    Antwort Status (1 Byte) 0x00: Busy  0xFF: Idle

' Slewing rate      [1/10 Motor Incr/sec]  i.e.  Slewing rate 10000: 10000 Incr in 10sec

' Slewing rate 9: "CE BB 02" = "CE BB 02" = 13548290 = 1354829,0 [Incr pro sec]



' Tracking a star:
' 01 15             Answer: az az az al al al        Get Az Alt Incr
' 18 0A F0 46  ????
' 01 15             Answer: az az az al al al        Get Az Alt Incr
' 06 00 00 5E 1B 00 00 05                            Slew Az=94 Alt=5









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
                
                If l = 3 Then
                    TelIncr.Az = GetNexStarPosition(NexStarAz)
                    WriteComm "Az: " & TelIncr.Az, Receive
                Else
                    ErrorCount = ErrorCount + 1
                End If

            ElseIf Command = 21 Then
                Do
                    vbuf = NexStarComm.Input
                    bbuf = vbuf
                    NexStarAlt = NexStarAlt & Chr$(bbuf(0))
                     key = (bbuf(0))
                Loop While NexStarComm.InBufferCount > 0
                l = Len(NexStarAlt)
                
                If l = 3 Then
                    TelIncr.Alt = GetNexStarPosition(NexStarAlt)
                    WriteComm "Alt: " & TelIncr.Alt, Receive
                Else
                    ErrorCount = ErrorCount + 1
                End If
            ElseIf Command = 13 Then
                 NexStarChar1 = ""
                Do
                    vbuf = NexStarComm.Input
                    bbuf = vbuf
                    NexStarChar1 = NexStarChar1 & Chr$(bbuf(0))
                     key = (bbuf(0))
                    List1.AddItem key
                    
                    If key = 0 Then
                        StatusMoving = 1      'Busy
                    ElseIf key = 255 Then
                        StatusMoving = 2      'Idle
                    End If
                    
                Loop While NexStarComm.InBufferCount > 0
                l = Len(NexStarChar1)
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
    
    If SimOffline Then
        C_Simulation.BackColor = RGB(255, 255, 0)
    Else
        C_Simulation.BackColor = RGB(255, 255, 255)
    End If
    
    If Not CommTest Then
        
        Static GetAzAltCount As Long
        Const GetAzAltInterval = 1        'calculate new star positition ever ... sec
        Dim N As Long
        
        N = (GetAzAltInterval * 1000) / Tim_DisplayUpdate.Interval
        
        GetAzAltCount = GetAzAltCount + 1
            
         ' this code only every "GetAzAltInterval" sec
        If GetAzAltCount >= N Then
            GetAzAltCount = 0
            If Toggle Then
                Toggle = False
                C_GetAz_Click
            Else
                Toggle = True
                C_GetAlt_Click
            End If
        End If

    End If
    
    
    MatrixSystemIst = MotorIncr_To_MatrixSystem(TelIncr)
    
    L_GlobalAzOffset = Format(RadToDeg(GlobalOffset.Az), "0.0000") & "°"
    L_MatrixSystemAzSoll = Format(RadToDeg(MatrixSystemSoll.Az), "0.0000") & "°"
    L_MatrixSystemAzDiff = Format(RadToDeg(MatrixDiffCalc.Az), "0.0000") & "°/s"
    L_MatrixSystemAzIst = Format(RadToDeg(MatrixSystemIst.Az), "0.0000") & "°"
    L_AzMotorIncr = Format(TelIncr.Az, "0.0")
    L_MotorSystemAzDiff = Format(MotorDiffCalc.Az, "0.0")
   
   
    L_GlobalAltOffset = Format(RadToDeg(GlobalOffset.Alt), "0.0000") & "°"
    L_MatrixSystemAltSoll = Format(RadToDeg(MatrixSystemSoll.Alt), "0.0000") & "°"
    L_MatrixSystemAltDiff = Format(RadToDeg(MatrixDiffCalc.Alt), "0.0000") & "°/s"
    L_MatrixSystemAltIst = Format(RadToDeg(MatrixSystemIst.Alt), "0.0000") & "°"
    L_AltMotorIncr = Format(TelIncr.Alt, "0.0")
    L_MotorSystemAltDiff = Format(MotorDiffCalc.Alt, "0.0")
   
   
    L_MatrixSystemAzDiffReal = "Real track rate: " & Format(RadToDeg(MatrixSystemDiffPerSec.Az), "0.0000") & "°/s"
    L_MotorSystemAzDiffReal = "Real (RS232):      " & Format(TrackingSpeed.Az / 10, "0.000") & " Incr/s"
    L_MotorSystemAzDiffSim = "Real (Simulation): " & Format(SimTrackingStep.Az, "0.000") & " Incr/s"
   
    If SimOffline Then
        Slider1.Value = SimIncr.Az - Matrix_To_MotorIncrSystem(MatrixSystemSoll).Az
        zzSlider1.Value = SimIncr.Az - Matrix_To_MotorIncrSystem(MatrixSystemSoll).Az
    Else
        Slider1.Value = TelIncr.Az - Matrix_To_MotorIncrSystem(MatrixSystemSoll).Az
        zzSlider1.Value = TelIncr.Az - Matrix_To_MotorIncrSystem(MatrixSystemSoll).Az
    End If
    
    L_ErrorCount = "Error cout: " & ErrorCount
    
    
    
    Select Case StatusMoving
        Case 0
            L_StatusMoving = "--"
            L_StatusMoving.BackColor = vbWhite
        Case 1
            L_StatusMoving = "Busy"
            L_StatusMoving.BackColor = vbYellow
        Case 2
            L_StatusMoving = "Idle"
            L_StatusMoving.BackColor = vbGreen
    End Select
        
        
    
End Sub

Private Sub Tim_Simulation_Timer()
    Dim SimScaling As Double
    Dim SimGotoStep As Double

    SimScaling = 100
    SimGotoStep = 10000

    If SimBntUp Then
        SimIncr.Alt = SimIncr.Alt + (ManualSlewingSpeed / SimScaling)
    End If
        
    If SimBntDn Then
        SimIncr.Alt = SimIncr.Alt - (ManualSlewingSpeed / SimScaling)
    End If
    
    If SimBntLe Then
        SimIncr.Az = SimIncr.Az - (ManualSlewingSpeed / SimScaling)
    End If
        
    If SimBntRi Then
        SimIncr.Az = SimIncr.Az + (ManualSlewingSpeed / SimScaling)
    End If
    
    
    ' movement active
    If SimGotoAzAltActive Then
        If CheckDeltaIncr(SimGoto.Az, SimIncr.Az, SimGotoStep * 2) Then
            SimIncr.Az = SimGoto.Az
        Else
            SimIncr.Az = SimIncr.Az + GetShortestWay(SimGoto.Az, SimIncr.Az) * SimGotoStep
        End If
    
        If CheckDeltaIncr(SimGoto.Alt, SimIncr.Alt, SimGotoStep * 2) Then
            SimIncr.Alt = SimGoto.Alt
        ElseIf SimGoto.Alt > SimIncr.Alt Then
            SimIncr.Alt = SimIncr.Alt + SimGotoStep
        Else
            SimIncr.Alt = SimIncr.Alt - SimGotoStep
        End If
        
        ' movement finished
        If (SimIncr.Az = SimGoto.Az) And (SimIncr.Alt = SimGoto.Alt) Then
            SimGotoAzAltActive = False
        End If

    End If
    
    
    ' tracking active
    If TrackingisON Then
        SimIncr = AddAzAlt(SimIncr, SimTrackingStep)
'''        SimTrackingActive = False
       
    End If
    
    
    SimIncr.Az = CutIncr(SimIncr.Az)
    SimIncr.Alt = CutIncr(SimIncr.Alt)
    
    Dim DispSimInc As AzAlt
    ' convert SimIncr to radian
    DispSimInc.Az = (SimIncr.Az / EncoderResolution) * 2 * Pi
    DispSimInc.Alt = (SimIncr.Alt / EncoderResolution) * 2 * Pi
    
    DispTelescopePos DispSimInc
    
End Sub



Private Sub Tim_Startup_Timer()
    Static StartupStep As Long
    
    Select Case StartupStep
    Case 0
        'check communication
        If SimOffline Then
            StartupStep = 10
        Else
            Command = 13
            NexStarComm.Output = Chr$(&HD)
            StatusMoving = 0
            StartupStep = 10
        End If
        
    
    Case 10
        If SimOffline Then
            StartupStep = 20
        Else
            If StatusMoving = 1 Or StatusMoving = 2 Then
                StartupStep = 20    'Communication OK
            Else
                StartupStep = 0     'Communication not OK. Retry
            End If
        End If
        
    Case 20
        C_SetEncoder_Az_Click
        C_SetEncoder_Az.BackColor = vbGreen
        StartupStep = 30
        
    Case 30
        C_SetEncoder_Alt_Click
        C_SetEncoder_Alt.BackColor = vbGreen
        StartupStep = 40
        
    Case 40
        C_SetBacklAz_Click
        C_SetBacklAz.BackColor = vbGreen
        StartupStep = 50

    Case 50
        C_SetBacklAlt_Click
        C_SetBacklAlt.BackColor = vbGreen
        StartupStep = 60
        
    Case 60
        If SimOffline Then
            StartupStep = 70
        Else
        End If


    End Select
    

End Sub

Private Sub Tim_TestStatus_Timer()
    If TestStatus Then
        Command = 13
        If SimOffline Then
        Else
            NexStarComm.Output = Chr$(&HD)
        End If
    End If
    
    
    If StatusMoving = 2 Then
        TestStatus = False
    End If

    
    
End Sub

Private Sub Tim_Tracking_Timer()
'    Dim tTime As MyTime
    Dim tDate As MyDate
    Dim tTs As MyTime
    Dim tTsRad As Double
    Dim LongitudeDeg As Double
    Dim LongitudeRad As Double
    Dim LatitudeDeg As Double
    Dim LatitudeRad As Double
  
    

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
    
    LatitudeDeg = GeoToDez(ObserverLatt)
    LatitudeRad = DegToRad(LatitudeDeg)
    LongitudeDeg = GeoToDez(ObserverLong)
    LongitudeRad = DegToRad(LongitudeDeg)

    GlbOberverPos.Latitude = LatitudeRad
    GlbOberverPos.Longitude = LongitudeRad
    
    'double check siderial time: https://tycho.usno.navy.mil/sidereal.html
    tTsRad = TimeToRad(GMST(tDate, ObserverTimeUT)) - LongitudeRad
    GlbSiderialTime = TimeToRad(GMST(tDate, ObserverTimeUT))
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
    
'    ObserverRA = HourToRad(AlignmentStarArray(idx).Ra)
'    ObserverDEC = DegToRad(AlignmentStarArray(idx).Dec)
    ObserverRaDec.Ra = HourToRad(AlignmentStarArray(idx).Ra)
    ObserverRaDec.Dec = DegToRad(AlignmentStarArray(idx).Dec)
    DisplayCoordinate L_I_RA, ObserverRaDec.Ra, HMS
    DisplayCoordinate L_I_DEC, ObserverRaDec.Dec, DegDec
     
'    RA_DEC_to_AZ_ALT_radian ObserverRA, ObserverDEC, ObserverLong, ObserverLatt, ObserverDateTimeUT, ObserverAzAlt.Az, ObserverAzAlt.Alt, HourAngle
    ObserverAzAlt = RA_DEC_to_AZ_ALT_new(ObserverRaDec, GlbOberverPos, GlbSiderialTime)
    ObserverAzAlt.Az = ObserverAzAlt.Az
    ObserverAzAlt.Alt = ObserverAzAlt.Alt

    Dim DisplObserverAz As Double
    If Ch_South.Value = 1 Then
        DisplObserverAz = CutRad(ObserverAzAlt.Az + Pi)
    Else
        DisplObserverAz = CutRad(ObserverAzAlt.Az)
    End If
    
    DisplayCoordinate L_I_Az, DisplObserverAz, DegDec
    
    Dim tmp As AzAlt
    tmp.Az = ObserverAzAlt.Az
    tmp.Alt = ObserverAzAlt.Alt
    L_CardinalOrientation = GetCardinalDrection(AzAlt_to_MatrixSystem(tmp).Az)
    DisplayCoordinate L_I_Alt, ObserverAzAlt.Alt, DegDec
    DisplayCoordinate L_I_HourAngle, HourAngle, HMS

            'Just for testing: get matrix vectors
            Dim X As Double
            Dim Y As Double
            Dim z As Double
            Dim HorizAngle As Double
            Dim ElevAngle As Double
            
            HorizAngle = ObserverRaDec.Ra
            ElevAngle = ObserverRaDec.Dec
            X = Cos(ElevAngle) * Cos(HorizAngle)
            Y = Cos(ElevAngle) * Sin(HorizAngle)
            z = Sin(ElevAngle)
            L_I_EquXYZ = Format(X, "0.0000") & " " & Format(Y, "0.0000") & " " & Format(z, "0.0000")
        
            HorizAngle = ObserverAzAlt.Az
            ElevAngle = ObserverAzAlt.Alt
            X = Cos(ElevAngle) * Cos(HorizAngle)
            Y = Cos(ElevAngle) * Sin(HorizAngle)
            z = Sin(ElevAngle)
            L_I_HorXYZ = Format(X, "0.0000") & " " & Format(Y, "0.0000") & " " & Format(z, "0.0000")


    If ObserverAzAlt.Alt < 0 Then
        L_CurrentStar.BackColor = RGB(255, 0, 0)
    ElseIf (ObserverAzAlt.Alt > 0) And (ObserverAzAlt.Alt < 0.3) Then
        L_CurrentStar.BackColor = RGB(255, 255, 0)
    Else
        L_CurrentStar.BackColor = RGB(0, 255, 0)
    End If
    
    
    
    
    
    Static TrackCount As Long
    Const TrackInterval = 30        'calculate new star positition ever ... sec
    Dim N As Long
    
    N = (TrackInterval * 1000) / Tim_Tracking.Interval

    ' start with tracking immediately
    If Not TrackingisON Then
        TrackCount = N
    End If

    If TrackingisON Then
            TrackCount = TrackCount + 1
            
            ' this code only every "TrackInterval" sec
            If TrackCount >= N Then
                        '==== Life counter ====
                        Static LifeCounter As Long
                        Dim i As Long
                        Dim s As String
                        If LifeCounter >= 10 Then
                            LifeCounter = 0
                            s = ""
                        Else
                            LifeCounter = LifeCounter + 1
                            For i = 0 To LifeCounter
                                s = s & "."
                            Next i
                        End If
          
                TrackCount = 0
                
                C_Tracking.BackColor = RGB(0, 255, 0)
            
                If GlbCalibStatus = 0 Then
                ElseIf GlbCalibStatus = 1 Then
                
                                
                            Dim Tmp1 As AzAlt
                            Tmp1.Az = CutRad(ObserverAzAlt.Az)
                            Tmp1.Alt = ObserverAzAlt.Alt
                            MatrixSystemSoll = AzAlt_to_MatrixSystem(Tmp1)
                            
'                             'Set Az
'                            MatrixSystemSoll.Az = CutRad(ObserverAzAlt.Az)
'                            'Set Alt
'                            MatrixSystemSoll.Alt = ObserverAzAlt.Alt
  
                ElseIf GlbCalibStatus = 2 Then
                            Dim AimTimeRad As Double
                            Dim AzAlt_BetaCet As AzAlt
                            Dim TimeDiff As MyTime
                            
                            
                            TimeDiff.s = TrackInterval
                            
                            AimTimeRad = TimeToRad(ObserverTimeUT) + TimeToRad(TimeDiff)
                            JetztTime = TimeToRad(ObserverTimeUT)
                        
                            CalculateTelescopeCoordinates Cal_InitTime, _
                                                          ObserverRaDec.Ra, ObserverRaDec.Dec, AimTimeRad, TransformationMatrix, _
                                                          AzAlt_BetaCet
                            'Set Az
                            MatrixSystemSoll.Az = CutRad(AzAlt_BetaCet.Az)
                            'Set Alt
                            MatrixSystemSoll.Alt = AzAlt_BetaCet.Alt
                End If
    
            

                Dim MatrixSystemDiff As AzAlt
                'real values
                
                'das funktioniert nicht wenn mal gegen den Sternenenlauf positioniert werden soll...
                 '... oder beim Durchgang durch 0°
                MatrixSystemDiff = SubAzAlt(MatrixSystemSoll, MatrixSystemIst)
                 'Ersatz...
                MatrixDiffCalc.Az = GetShortestRad(MatrixSystemSoll.Az, MatrixSystemIst.Az)
                MatrixDiffCalc.Alt = GetShortestRad(MatrixSystemSoll.Alt, MatrixSystemIst.Alt)
               
               
               
                MatrixSystemDiffPerSec.Az = MatrixSystemDiff.Az / TrackInterval
                MatrixSystemDiffPerSec.Alt = MatrixSystemDiff.Alt / TrackInterval
                
                
                
                'Hier geht das Vorzeichen verloren
'                DiffMotorIncr = Matrix_To_MotorIncrSystem(MatrixSystemDiff)
                'Ersatz
                DiffMotorIncr.Az = -MatrixSystemDiff.Az * EncoderResolution / (2 * Pi)
                DiffMotorIncr.Alt = MatrixSystemDiff.Alt * EncoderResolution / (2 * Pi)
                
                'check motor movement with calculated values
                
                'das funktioniert nicht wenn mal gegen den Sternenenlauf positioniert werden soll...
                '... oder beim Durchgang durch 0°
                'MatrixDiffCalc = SubAzAlt(MatrixSystemSoll, MatrixLastCalc)
                
                
                MatrixDiffCalc.Az = MatrixDiffCalc.Az / TrackInterval
                MatrixDiffCalc.Alt = MatrixDiffCalc.Alt / TrackInterval
                MatrixLastCalc = MatrixSystemSoll
                
                MotorDiffCalc = SubAzAlt(Matrix_To_MotorIncrSystem(MatrixSystemSoll), MotorLastCalc)
                MotorDiffCalc.Az = MotorDiffCalc.Az / TrackInterval
                MotorDiffCalc.Alt = MotorDiffCalc.Alt / TrackInterval
                MotorLastCalc = Matrix_To_MotorIncrSystem(MatrixSystemSoll)
  
                
                Label6 = Format(RadToDeg(MatrixSystemDiff.Az), "0.0000") & "° = " & Format(DiffMotorIncr.Az, "0.0") & " Incr pro " & TrackInterval & " sec"
                Label27 = Format(RadToDeg(MatrixSystemDiff.Alt), "0.0000") & "° = " & Format(DiffMotorIncr.Alt, "0.0") & " Incr pro " & TrackInterval & " sec"
  
               
                TrackingSpeed.Az = (DiffMotorIncr.Az * 10) / TrackInterval
                TrackingSpeed.Alt = (DiffMotorIncr.Alt * 10) / TrackInterval
                  
             
                Dim LogString As String
                Dim AzString As String
                Dim AltString As String
                
                'If DiffMotorIncr.Az < 0 Then
                If MatrixSystemDiff.Az > 0 Then
                    'CCW
                    AzString = Chr$(&H7) & SetNexStarPosition(-CLng(TrackingSpeed.Az))
                Else
                    'CW (normal direction)
                    AzString = Chr$(&H6) & SetNexStarPosition(CLng(TrackingSpeed.Az))
                End If
                
                If DiffMotorIncr.Alt < 0 Then
                    'descending
                    AltString = Chr$(&H1B) & SetNexStarPosition(-CLng(TrackingSpeed.Alt))
                Else
                    'ascending
                    AltString = Chr$(&H1A) & SetNexStarPosition(CLng(TrackingSpeed.Alt))
                End If
                    
                    
                If Not SimOffline Then
                    LogString = "Tracking: " & TrackingSpeed.Az & " " & TrackingSpeed.Alt
                    WriteComm LogString, Send
                    NexStarComm.Output = AzString & AltString
                    
                End If
            End If
                
                
                
            ' every TimTracking.Interval
            If SimOffline Then
'''                SimTrackingActive = True

                SimTrackingStep.Az = (DiffMotorIncr.Az / TrackInterval) * (Tim_Tracking.Interval / 1000#)
                SimTrackingStep.Alt = (DiffMotorIncr.Alt / TrackInterval) * (Tim_Tracking.Interval / 1000#)
            End If

    
    Else
            'Tracking is OFF
            C_Tracking.BackColor = &H8000000F
    End If



End Sub

Private Sub VS_ManualSlewingSpeed_Change()
    Dim tmp As Double
    
    tmp = VS_ManualSlewingSpeed.Value
    ManualSlewingSpeed = 1000 * tmp
    L_SlewingSpeed = ManualSlewingSpeed
    
    'TrackingSpeed[Incr/sec] = ManualSlewingSpeed[Incr/sec] * 0,1 [Incr/sec]

    
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








