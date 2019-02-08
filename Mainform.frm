VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Mainform 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
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
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   6720
      TabIndex        =   7
      Top             =   360
      Width           =   1215
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




Private Sub C_Dn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1 = "down"
    NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1B) & SetNexStarPosition(ManualSkewingSpeed)
End Sub

Private Sub C_Dn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1 = "down"
    NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1A) & SetNexStarPosition(0)
End Sub

Private Sub C_GetAz_Click()
    NexStarComm.Output = Chr$(&H1)
    NexStarAz = ""
    List1.Clear
    Command = 1
End Sub

Private Sub C_GetAlt_Click()
    NexStarComm.Output = Chr$(&H15)
    NexStarAlt = ""
    List1.Clear
    Command = 21
End Sub



Private Sub C_Le_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    NexStarComm.Output = Chr$(&H7) & SetNexStarPosition(ManualSkewingSpeed) & Chr$(&H1A) & SetNexStarPosition(0)
End Sub

Private Sub C_Le_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1A) & SetNexStarPosition(0)
End Sub


Private Sub C_Ri_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(ManualSkewingSpeed) & Chr$(&H1A) & SetNexStarPosition(0)
End Sub

Private Sub C_Ri_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1A) & SetNexStarPosition(0)
End Sub

Private Sub C_SetAzAlt_Click()
    Dim SetAz As Long
    Dim SetAlt As Long
    
    SetAz = CLng(Zahl(T_Az))
    SetAlt = CLng(Zahl(T_Alt))

    NexStarComm.Output = Chr$(&O2) & SetNexStarPosition(SetAz) & Chr$(&H16) & SetNexStarPosition(SetAlt)
End Sub

Private Sub C_SetEncoder_Click()
    NexStarComm.Output = Chr$(&HC) & SetNexStarPosition(726559) & SetNexStarPosition(726559)
End Sub



Private Sub C_Up_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1 = "down"
    NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1A) & SetNexStarPosition(ManualSkewingSpeed)
End Sub

Private Sub C_Up_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1 = "up"
    NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1A) & SetNexStarPosition(0)
End Sub





Private Sub Command3_Click()
    Dim a As String
    Dim b As String
    Dim erg As Long
    
    a = SetNexStarPosition(1234567)
    
    b = Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H11) & Chr$(&H24) & Chr$(&H80)
'    b = Chr$(&H0) & Chr$(&H3) & Chr$(&HE8)
    
    erg = GetNexStarPosition(a)
    
End Sub



Private Sub Form_Load()

    IniFileName = App.Path & "\NexStar.ini"
    InitNexStarComm
    ClearNexStarComm
    Command = 0
    
    VS_ManualSkewingSpeed.Value = 10
End Sub


Private Sub InitNexStarComm()

  On Error GoTo v24error
  
  NexStarPortNr = Zahl(INIGetValue(IniFileName, "NexStar", "PortNr"))
  NexStarBaudrate = Zahl(INIGetValue(IniFileName, "NexStar", "Baudrate"))

  If NexStarPortNr > 0 Then
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

Private Sub ClearNexStarComm()
    Dim tmp As String
    
    While NexStarComm.InBufferCount > 0
        tmp = NexStarComm.Input
    Wend
    
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
    Case comCDTO      ' CD-Zeit�berschreitung
    Case comCTSTO     ' CTS-Zeit�berschreitung
    Case comDSRTO     ' DSR-Zeit�berschreitung
    Case comFrame     ' Fehler im �bertragungsraster (Framing Error)
    Case comOverrun   ' Datenverlust
    Case comRxOver    ' �berlauf des Empfangspuffers
    Case comRxParity  ' Parit�tsfehler
    Case comTxFull    ' Sendepuffer voll
    Case comDCB       ' Unerwarteter Fehler beim Abrufen des DCB]

  ' Ereignisse
    Case comEvCD  ' Pegel�nderung auf DCD
    Case comEvCTS ' Pegel�nderung auf CTS
    Case comEvDSR ' Pegel�nderung auf DSR
    Case comEvRing  ' Pegel�nderung auf RI(Ring Indicator)
    Case comEvReceive ' Anzahl empfangener Zeichen gleich RThreshold
    
            If Command = 1 Then
                Do
                    vbuf = NexStarComm.Input
                    bbuf = vbuf
                    NexStarAz = NexStarAz & Chr$(bbuf(0))
                     key = (bbuf(0))
                   List1.AddItem (key)
                Loop While NexStarComm.InBufferCount > 0
                l = Len(NexStarAz)
                TelIncrAz = GetNexStarPosition(NexStarAz)
                L_Az = TelIncrAz
                TelDegAz = TelIncrAz * 360 / 726559
                L_TelDegAz = Format(TelDegAz, "0.0000")
            ElseIf Command = 21 Then
                Do
                    vbuf = NexStarComm.Input
                    bbuf = vbuf
                    NexStarAlt = NexStarAlt & Chr$(bbuf(0))
                     key = (bbuf(0))
                   List1.AddItem (key)
                Loop While NexStarComm.InBufferCount > 0
                l = Len(NexStarAlt)
                TelIncrAlt = GetNexStarPosition(NexStarAlt)
                L_Alt = TelIncrAlt
                TelDegAlt = TelIncrAlt * 360 / 726559
                L_TelDegAlt = Format(TelDegAlt, "0.0000")
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
End Sub

Private Sub VS_ManualSkewingSpeed_Change()
    Dim tmp As Long
    
    tmp = VS_ManualSkewingSpeed.Value
    ManualSkewingSpeed = 1000 * tmp
End Sub
