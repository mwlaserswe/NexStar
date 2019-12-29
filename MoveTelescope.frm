VERSION 5.00
Begin VB.Form MoveTelescope 
   Caption         =   "Move Telescope"
   ClientHeight    =   6195
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer MouseTimer 
      Interval        =   200
      Left            =   5280
      Top             =   1080
   End
   Begin VB.Frame Frame_Aktuelle_Galvo_Position 
      Caption         =   "Aktuelle Galvo Position"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   3015
      Begin VB.Label Label_Achse1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Achse 1"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label_Achse_1_Value 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label_Achse2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Achse 2"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label_Achse2_Value 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame Frame_Mouse_Move 
      Caption         =   "Mouse Move"
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      Begin VB.PictureBox Picture_Mouse_Move 
         Height          =   3615
         Left            =   240
         ScaleHeight     =   3555
         ScaleWidth      =   3675
         TabIndex        =   1
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Label L_SpeedY 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   4920
      TabIndex        =   11
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label L_SpeedX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "[ESC] Move Galvo Stop"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   4560
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "[F12] Punkt übernehmen"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "[F1] Grob [F2] Mittel [F3] Fein"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   5160
      Width           =   2895
   End
   Begin VB.Menu M_PopUp 
      Caption         =   "PopUp"
      Begin VB.Menu M_Move_Galvo_Start 
         Caption         =   "Move Galvo Start"
      End
      Begin VB.Menu M_Move_Galvo_Stop 
         Caption         =   "[ESC] Move Galvo Stop"
      End
      Begin VB.Menu M_Punkt_UEB 
         Caption         =   "[F12] Punkt übernehmen"
      End
      Begin VB.Menu M_Trennlinie 
         Caption         =   "-"
      End
      Begin VB.Menu M_Grob 
         Caption         =   "[F1] Grob"
      End
      Begin VB.Menu M_Mittel 
         Caption         =   "[F2] Mittel"
      End
      Begin VB.Menu M_Fein 
         Caption         =   "[F3] Fein"
      End
   End
End
Attribute VB_Name = "MoveTelescope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''Dim Winsock1SendComplete As Boolean
'''Dim WinsockInputbuffer As New Collection

Public MoveGalvoAKT As Boolean

Dim PixelFaktor As Double
Dim GalvoLastPositionX As Double
Dim GalvoLastPositionY As Double

Dim PixelFaktorGrob As Double
Dim PixelFaktorMittel As Double
Dim PixelFaktorFein As Double

'''Dim GalvopointX() As Double
'''Dim GalvopointY() As Double
'''Dim GalvoPointcounter As Long
'''Dim PointsInArray As Boolean

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Dim LastMousePositionX As Double
Dim LastMousepositionY As Double
Dim P1 As POINTAPI
Dim P2 As POINTAPI

Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

'''Private Sub Befehls_Timer_Timer()
'''
'''Dim befehlsnummer As Integer
'''Dim Argumente() As String
'''Dim i As Long, f As Long
'''Dim MoveX As Double
'''Dim MoveY As Double
'''Dim PointReturn As String
'''
'''If WinsockInputbuffer.Count <> 0 Then
'''
'''    befehlsnummer = readnr(WinsockInputbuffer(1))
'''
'''    If befehlsnummer = -1 Then
'''      WinsockInputbuffer.Remove (1)
'''      Exit Sub
'''    End If
'''
'''  Select Case befehlsnummer
'''
'''    Case 10:
'''           ReDim GalvopointX(0 To 0)
'''           ReDim GalvopointY(0 To 0)
'''           GalvoPointcounter = 0
'''           List_Galvopoint.Clear
'''           MoveGalvoAKT = False
'''           MoveGalvo 0, 0
'''           PointsInArray = False
'''           WinsockSendData "10 MG Ready"
'''
'''    Case 20:
'''            SepariereString WinsockInputbuffer(1), Argumente, " "
'''            For i = LBound(Argumente) To UBound(Argumente)
'''              Select Case Mid(Argumente(i), 1, 1)
'''                Case "X", "x": MoveX = CDbl(Zahl(Mid(Argumente(i), 2, Len(Argumente(i)))))
'''                Case "Y", "y": MoveY = CDbl(Zahl(Mid(Argumente(i), 2, Len(Argumente(i)))))
'''              End Select
'''            Next i
'''            If IsNumeric(MoveX) And IsNumeric(MoveY) Then
'''              MoveGalvo MoveX, MoveY
'''              WinsockSendData "21 Move Ready"
'''            Else
'''              WinsockSendData "29 Abbruch"
'''            End If
'''    Case 30:
'''            If PointsInArray Then
'''            PointReturn = "31 Point Return "
'''            For f = LBound(GalvopointX) To UBound(GalvopointX)
'''              PointReturn = PointReturn & " X" & CStr(GalvopointX(f)) & " Y" & CStr(GalvopointY(f))
'''            Next f
'''            WinsockSendData PointReturn
'''            Else
'''            WinsockSendData "39 No Points"
'''            End If
'''  End Select
'''
'''  WinsockInputbuffer.Remove (1)
'''
'''End If
'''
'''End Sub

'''Private Sub Button_Achse12_Set_0_Click()
'''  MoveGalvo 0, 0
  
  
''
''
'' DoCommand "LaserOff"
''
'' DoCommand "Start"
''
''
 
 
 
 
''' End Sub

'''Private Sub Button_Clear_All_Click()
'''  ReDim GalvopointX(0 To 0)
'''  ReDim GalvopointY(0 To 0)
'''  GalvoPointcounter = 0
'''  List_Galvopoint.Clear
'''  PointsInArray = False
'''End Sub

'''Private Sub Button_Start_Click()
'''  Dim i As Long
'''  Dim Q As Long
'''
'''
'''
'''
'''End Sub

'''Private Sub button_test_Click()
'''Dim i As Integer
'''Dim outer As Integer
'''
'''Dim X As Double
'''Dim Y As Double
'''
''''DoCommand "Move", Str(X) & "," & Str(Y)
'''
'''For outer = 1 To 1
'''  For i = 1 To 360 Step 2
'''  'DoCommand "DoLine", Str(Cos(i) * 25) & "," & Str(Sin(i) * 25) & "," & Str(Cos(i + 1) * 25) & "," & Str(Sin(i + 1) * 25) & ",100"
'''
'''  Select Case i
'''  Case 0 To 90:
'''    X = Sin(i)
'''    Y = Cos(i)
'''  Case 91 To 180:
'''      X = Sin(i)
'''    Y = Cos(i)
'''  Case 181 To 270:
'''      X = Sin(i)
'''    Y = Cos(i)
'''  Case 271 To 360:
'''      X = Sin(i)
'''    Y = Cos(i)
'''  End Select
'''
'''
'''  Next i
'''Next outer
'''
'''End Sub


Private Sub Form_Load()
  
  InitMoveGalvo
  M_Move_Galvo_Start_Click
'''  GalvoPointcounter = 0

End Sub


'''Private Sub KeyTimer_Timer()
'''Dim tmpMouseP As POINTAPI
'''If MoveGalvoAKT = True Then
'''  If CompKey(vbKeyUp) Then
'''  GetCursorPos tmpMouseP
'''  SetCursorPos tmpMouseP.X, tmpMouseP.Y - 1
'''  Exit Sub
'''  End If
'''
'''  If CompKey(vbKeyDown) Then
'''  GetCursorPos tmpMouseP
'''  SetCursorPos tmpMouseP.X, tmpMouseP.Y + 1
'''  Exit Sub
'''  End If
'''
'''  If CompKey(vbKeyLeft) Then
'''  GetCursorPos tmpMouseP
'''  SetCursorPos tmpMouseP.X - 1, tmpMouseP.Y
'''  Exit Sub
'''  End If
'''
'''  If CompKey(vbKeyRight) Then
'''  GetCursorPos tmpMouseP
'''  SetCursorPos tmpMouseP.X + 1, tmpMouseP.Y
'''  Exit Sub
'''  End If
'''
'''  If CompKey(vbKeyF12) Then
'''  M_Punkt_UEB_Click
'''  Exit Sub
'''  End If
'''End If
'''
'''If CompKey(vbKeyF1) Then
'''  PixelFaktor = PixelFaktorGrob
'''  Exit Sub
'''End If
'''
'''If CompKey(vbKeyF2) Then
'''  PixelFaktor = PixelFaktorMittel
'''  Exit Sub
'''End If
'''
'''If CompKey(vbKeyF3) Then
'''  PixelFaktor = PixelFaktorFein
'''  Exit Sub
'''End If
'''
'''If CompKey(vbKeyEscape) Then
'''  M_Move_Galvo_Stop_Click
'''  Exit Sub
'''End If
'''
'''End Sub

Private Sub M_Fein_Click()
  PixelFaktor = PixelFaktorFein
  MouseTimer.Enabled = True
End Sub

Private Sub M_Grob_Click()
  PixelFaktor = PixelFaktorGrob
  MouseTimer.Enabled = True
End Sub

Private Sub M_Mittel_Click()
  PixelFaktor = PixelFaktorMittel
  MouseTimer.Enabled = True
End Sub

Private Sub M_Move_Galvo_Start_Click()
  MoveGalvoAKT = True
  P1.X = ScaleX(Picture_Mouse_Move.Left, ScaleMode, vbPixels)
  P1.Y = ScaleY(Picture_Mouse_Move.Top, ScaleMode, vbPixels)
  ClientToScreen hWnd, P1
  P2.X = ScaleX(Picture_Mouse_Move.Left + Picture_Mouse_Move.Width, ScaleMode, vbPixels)
  P2.Y = ScaleY(Picture_Mouse_Move.Top + Picture_Mouse_Move.Height, ScaleMode, vbPixels)
  ClientToScreen hWnd, P2
  MouseTimer.Enabled = True
End Sub

Private Sub M_Move_Galvo_Stop_Click()
    MoveGalvoAKT = False
    
    L_SpeedX = 0
    L_SpeedY = 0

    Mainform.NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(0) & Chr$(&H1A) & SetNexStarPosition(0)

End Sub

'''Private Sub M_Punkt_UEB_Click()
'''
'''  PointsInArray = True
'''
'''  ReDim Preserve GalvopointX(0 To GalvoPointcounter)
'''  ReDim Preserve GalvopointY(0 To GalvoPointcounter)
'''
'''  GalvopointX(GalvoPointcounter) = GalvoLastPositionX
'''  GalvopointY(GalvoPointcounter) = GalvoLastPositionY
'''  List_Galvopoint.AddItem CStr(GalvopointX(GalvoPointcounter)) & "      " & CStr(GalvopointY(GalvoPointcounter))
'''  GalvoPointcounter = GalvoPointcounter + 1
'''  MouseTimer.Enabled = True
'''End Sub

Private Sub MouseTimer_Timer()

    Dim GlobalMousePosition As POINTAPI
    Dim XDiffPix As Double
    Dim YDiffPix As Double
    Dim XDiffRel As Double
    Dim YDiffRel As Double
    Dim CommString As String
    Dim CmdX As String
    Dim CmdY As String
    Dim CommentX As String
    Dim CommentY As String

    
    If MoveGalvoAKT = True Then
    
        GetCursorPos GlobalMousePosition
        
        XDiffPix = CDbl(GlobalMousePosition.X) - LastMousePositionX
        YDiffPix = CDbl(GlobalMousePosition.Y) - LastMousepositionY
        
        
        XDiffRel = XDiffPix * PixelFaktor
        YDiffRel = -YDiffPix * PixelFaktor
        
        
        If Abs(XDiffRel) > (Abs(YDiffRel)) Then
            YDiffRel = 0
        ElseIf Abs(YDiffRel) > (Abs(XDiffRel)) Then
            XDiffRel = 0
        End If
        
        
        L_SpeedX = XDiffRel
        L_SpeedY = YDiffRel
        
        
        
        
        
        
                
        If Command <> 0 Then
            If XDiffRel >= 0 Then
    '''            Mainform.NexStarComm.Output = Chr$(&H6) & SetNexStarPosition(CDbl(XDiffRel))
                CmdX = Chr$(&H6)
                CommentX = "Move right (0x06) " & CDbl(XDiffRel)
            Else
    '''             Mainform.NexStarComm.Output = Chr$(&H7) & SetNexStarPosition(CDbl(-XDiffRel))
                CmdX = Chr$(&H7)
                XDiffRel = -XDiffRel
                CommentX = "Move left (0x07) " & CDbl(XDiffRel)
            End If
    
            If YDiffRel >= 0 Then
    '''            Mainform.NexStarComm.Output = Chr$(&H1A) & SetNexStarPosition(CDbl(YDiffRel))
                CmdY = Chr$(&H1A)
                CommentY = ", up (0x1A) " & CDbl(YDiffRel)
            Else
    '''             Mainform.NexStarComm.Output = Chr$(&H1B) & SetNexStarPosition(CDbl(-YDiffRel))
                CmdY = Chr$(&H1B)
                YDiffRel = -YDiffRel
                CommentY = ", down (0x1B) " & CDbl(YDiffRel)
            End If
           
            CommString = CmdX & SetNexStarPosition(CDbl(XDiffRel)) & CmdY & SetNexStarPosition(CDbl(YDiffRel))
            NexStarCommunication CommString, CommentX & CommentY, Send
        End If
        
        
        
        
        LastMousePositionX = CDbl(GlobalMousePosition.X)
        LastMousepositionY = CDbl(GlobalMousePosition.Y)
        
        If GlobalMousePosition.X < P1.X Or GlobalMousePosition.X > P2.X Or _
           GlobalMousePosition.Y < P1.Y Or GlobalMousePosition.Y > P2.Y Then
           
           SetCursorPos (P1.X + ((P2.X - P1.X) / 2)), (P1.Y + ((P2.Y - P1.Y) / 2))
           
           LastMousePositionX = CDbl((P1.X + ((P2.X - P1.X) / 2)))
           LastMousepositionY = CDbl((P1.Y + ((P2.Y - P1.Y) / 2)))
           
        End If
        
    
    End If
    
    If CompKey(vbKeyF1) Then
      PixelFaktor = PixelFaktorGrob
      Exit Sub
    End If
      
    If CompKey(vbKeyF2) Then
      PixelFaktor = PixelFaktorMittel
      Exit Sub
    End If
    
    If CompKey(vbKeyF3) Then
      PixelFaktor = PixelFaktorFein
      Exit Sub
    End If

    If CompKey(vbKeyEscape) Then
                 
      M_Move_Galvo_Stop_Click
      Exit Sub
    End If

End Sub

Private Sub Picture_Mouse_Move_Click()
  If MoveGalvoAKT Then MouseTimer.Enabled = True
End Sub

Private Sub Picture_Mouse_Move_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 Then
  MouseTimer.Enabled = False
  PopupMenu M_PopUp
 End If
End Sub

Private Sub InitMoveGalvo()

  PixelFaktorGrob = 200
  PixelFaktorMittel = 100
  PixelFaktorFein = 1

  PixelFaktor = PixelFaktorMittel

  MoveGalvoAKT = False

'''  MoveGalvo 0, 0

End Sub

'''Private Sub MoveGalvo(X As Double, Y As Double)
'''
'''
'''
'''  GalvoLastPositionX = X
'''  GalvoLastPositionY = Y
'''
'''  Label_Achse_1_Value = Format(Str(X) & " mm  ", "0.000")
'''  Label_Achse2_Value = Format(Str(Y) & " mm  ", "0.000")
'''
'''End Sub


Private Function CompKey(KCode&) As Boolean
  Dim result%
    result = GetAsyncKeyState(KCode)
    If result = -32767 Then
      CompKey = True
    Else
      CompKey = False
    End If
End Function

'''Private Sub Winsock1_Close()
'''  Winsock1.Close
'''  Winsock1.Listen
'''End Sub
'''
'''Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
'''  If Winsock1.State <> sckClosed Then Winsock1.Close
'''  Winsock1.Accept requestID
'''End Sub
'''
'''Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'''  Dim DataIn As String
'''  Dim i As Long, f As Long
'''
'''  Winsock1.GetData DataIn
'''
'''    If Right(DataIn, 1) = vbCr Then
'''      For i = 1 To bytesTotal
'''        If Mid(DataIn, i, 1) = vbCr Then
'''          WinsockInputbuffer.Add (Mid(DataIn, f + 1, i - f - 1))
'''          f = i
'''        End If
'''      Next i
'''    End If
'''  DataIn = ""
'''
'''End Sub
'''
'''Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'''  Winsock1.Close
'''  CancelDisplay = True
'''End Sub
'''
'''Private Sub Winsock1_SendComplete()
'''  Winsock1SendComplete = True
'''End Sub
'''
'''Private Sub WinsockSendData(Data As String)
'''  Winsock1SendComplete = False
'''
'''  On Error GoTo WinsockError
'''
'''  If Winsock1.State = sckClosed Then Exit Sub
'''
'''  Winsock1.SendData (Data & vbCr)
'''
'''  Exit Sub
'''
'''WinsockError:
'''Winsock1SendComplete = True
'''Winsock1.Close
'''
'''End Sub
'''
'''Private Function readnr(Line As String) As Integer
'''' Liest die Befehlsnummer aus der Zeile
'''' Wenn die ersten beiden Zeichen keine Ziffern sind, wird -1 zurückgeliefert
'''  If Trim(Line) Like "##*" Then
'''    readnr = CInt(Left(Trim(Line), 2))
'''  Else
'''    readnr = -1
'''  End If
'''End Function
'''
'''Private Sub SepariereString(Zeile As String, WortArray() As String, Delimiter As String)
'''  Dim Pos1 As Long
'''  Dim Pos2 As Long
'''  Dim AnzahlWorte As Long               'Anzahl der Worte der Zeile
'''
'''  ReDim WortArray(0 To 0)                         'WortArray löschen
'''  AnzahlWorte = 0
'''  Pos2 = 0
'''
'''  If Delimiter = " " Then Zeile = Trim(Zeile)
'''
'''  Zeile = Trim(Zeile)
'''  Do
'''    Pos1 = Pos2
'''
'''    'Trennzeichen [CR]: [LF] werden überlesen
'''    If Delimiter = vbCr Then
'''      If Mid(Zeile, Pos1 + 1, 1) = vbLf Then
'''        Pos1 = Pos1 + 1                             'LF übergehen
'''      End If
'''    End If
'''
'''     'Trennzeichen [Space]: [Space] werden überlesen
'''    If Delimiter = " " Then
'''      Do While Mid(Zeile, Pos1 + 1, 1) = " "
'''        Pos1 = Pos1 + 1                             'Space übergehen
'''      Loop
'''    End If
'''
'''    Pos2 = InStr(Pos1 + 1, Zeile, Delimiter)      'nach Trennzeichen suchen
'''    If Pos2 > 0 Then                              'noch ein Trennzeichen in der Zeile
'''      WortArray(AnzahlWorte) = Mid(Zeile, Pos1 + 1, Pos2 - Pos1 - 1)
'''      ReDim Preserve WortArray(0 To UBound(WortArray) + 1)
'''      AnzahlWorte = AnzahlWorte + 1
'''    Else                                          'kein Trennzeichen mehr vorhanden
'''      WortArray(AnzahlWorte) = Mid(Zeile, Pos1 + 1)
'''    End If
'''  Loop While Pos2 > 0
'''End Sub
'''
'''Private Function Zahl(Txt As String) As Double
'''' Wandelt die Zahl in einem String in eine Zahl um
'''' dabei werden "," in "." umgewandelt und alle Zeichen
'''' die nicht passen in Leerzeichen gewandelt
'''' 22.07.2002 Exponent möglich
''''            wenn keine Ziffern vorhanden sind, wird Err.number = 1 gesetzt
'''  Dim i As Integer
'''  Dim s As String
'''  Dim noVorz As Boolean, noKomma As Boolean, noExpo As Boolean, haveDigits As Boolean
'''  s = ""
'''  For i = 1 To Len(Txt)
'''    Select Case Mid(Txt, i, 1)
'''      Case "+", "-"
'''        If Not noVorz Then
'''          s = s + Mid(Txt, i, 1)
'''          noVorz = True
'''        Else
'''          Exit For
'''        End If
'''      Case ",", "."
'''        If Not noKomma Then
'''          s = s + "."
'''          noKomma = True
'''          noVorz = True
'''        Else
'''          Exit For
'''        End If
'''      Case "0" To "9"
'''        s = s + Mid(Txt, i, 1)
'''        noVorz = True
'''        haveDigits = True
'''      Case "&"
'''        s = s + Mid(Txt, i, 2)
'''        noVorz = True
'''      Case "E", "e"
'''        If Not noExpo Then
'''           s = s + Mid(Txt, i, 1)
'''          noVorz = False
'''          noKomma = True
'''          noExpo = True
'''        Else
'''          Exit For
'''        End If
'''      Case " "
'''      Case Else
'''        If noVorz Then Exit For
'''    End Select
'''  Next i
'''  If Not haveDigits Then
'''    Err.Number = 1
'''    Err.Description = "Zahl set to 0. No Digits in String"
'''  End If
'''  Zahl = Val(s)
'''End Function
'''
'''
'''
