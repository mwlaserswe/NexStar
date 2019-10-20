VERSION 5.00
Begin VB.Form Test2 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Test2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Command1_Click()
    Dim SaturnRaDec As RaDec
    

    Dim SaturnAzAlt As AzAlt
    
    SaturnAzAlt.Az = DegToRad(-51.6992)           ' A Azimut
    SaturnAzAlt.Alt = DegToRad(36.5405)           ' h Höhe
   
    SaturnRaDec = AZ_ALT_to_RA_DEC(SaturnAzAlt, GlbOberverPos, GlbSiderialTime)
    
    Dim dmy1 As Double
    Dim dmy2 As Double
    dmy1 = SaturnRaDec.Ra
    dmy2 = SaturnRaDec.Dec
    
End Sub
