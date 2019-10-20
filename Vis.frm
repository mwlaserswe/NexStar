VERSION 5.00
Begin VB.Form Vis 
   Caption         =   "Form1"
   ClientHeight    =   8760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   11835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   10440
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.PictureBox Pic 
      Height          =   7335
      Left            =   240
      ScaleHeight     =   7275
      ScaleWidth      =   9915
      TabIndex        =   0
      Top             =   240
      Width           =   9975
   End
End
Attribute VB_Name = "Vis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Pic.Line (100, 100)-(2000, 1000), vbRed
'    Pic.TextHeight = 50
'    Pic.TextWidth = 50
Pic.PSet (500, 1600)
    Pic.Print "Hallo"
End Sub

Private Sub Form_Activate()
   DispInit
    DispCoordinateSystem
End Sub

 
