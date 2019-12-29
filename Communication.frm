VERSION 5.00
Begin VB.Form Communication 
   Caption         =   "Communication"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton C_Stop 
      Caption         =   "Stop"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ListBox List_log 
      Height          =   5325
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   4215
   End
   Begin VB.Timer DisplayUpdateTimer 
      Interval        =   100
      Left            =   720
      Top             =   2760
   End
   Begin VB.CommandButton C_TestCommMotorToHandheld 
      Caption         =   "Motor -> Handheld"
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton C_TestCommHandheldToMotor 
      Caption         =   "Handheld -> Motor"
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "Communication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim LineString As String
Dim iAzLastValue As Long
Dim iAltLastValue As Long
Public StopFlag As Boolean




Private Sub C_Stop_Click()
    StopFlag = Not StopFlag
End Sub

Private Sub C_TestCommHandheldToMotor_Click()
    If TestCommHandheldToMotor Then
        TestCommHandheldToMotor = False
    Else
        TestCommHandheldToMotor = True
        TestCommMotorToHandheld = False
    End If
End Sub

Private Sub C_TestCommMotorToHandheld_Click()
    If TestCommMotorToHandheld Then
        TestCommMotorToHandheld = False
    Else
        TestCommMotorToHandheld = True
        TestCommHandheldToMotor = False
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DisplayUpdateTimer_Timer()
    If TestCommHandheldToMotor Then
       C_TestCommHandheldToMotor.BackColor = RGB(0, 255, 0)
    Else
       C_TestCommHandheldToMotor.BackColor = &H8000000F
    End If
    
    If TestCommMotorToHandheld Then
        C_TestCommMotorToHandheld.BackColor = RGB(0, 255, 0)
    Else
        C_TestCommMotorToHandheld.BackColor = &H8000000F
    End If
End Sub


Public Sub DisplayAzAltTracking(s As String)
    Dim i As Long
    Dim sLine As String
    Dim s1 As String
    Dim s2 As String
    Dim sAz As String
    Dim sAlt As String
    Dim iAz As Long
    Dim iAlt As Long
    Dim iAzDiff As Long
    Dim iAltDiff As Long
    
    LineString = LineString + s
    
    If Len(LineString) >= 6 Then
    
        sLine = "--> "
        
        For i = 1 To Len(LineString)
            s1 = Mid(LineString, i, 1)
            s2 = Hex(Asc(s1))
           sLine = sLine & s2 & " "
           
           
           
           
           
        Next i
        
'''        iAz = GetNexStarPosition(Mid(LineString, 1, 3))
'''        iAlt = GetNexStarPosition(Mid(LineString, 4, 3))
        
        iAzDiff = iAz - iAzLastValue
        iAltDiff = iAlt - iAltLastValue
        
        
        sLine = sLine & "  Az:" & iAz & " Alt:" & iAlt & " AzDiff:" & iAzDiff & " AltDiff:" & iAltDiff
        
        iAzLastValue = iAz
        iAltLastValue = iAlt
        
        LineString = ""
        
    End If
    
    List_log.AddItem sLine

End Sub


