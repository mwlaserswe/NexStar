VERSION 5.00
Begin VB.Form TestJulianischesDatum 
   Caption         =   "Test Julianisches Datum"
   ClientHeight    =   6390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox T_Zeitzone 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox T_Laenge 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox T_Stunden 
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox T_Minuten 
      Height          =   285
      Left            =   3480
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox T_Sekunden 
      Height          =   285
      Left            =   4080
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox T_Jahr 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox T_Monat 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox T_Tag 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Starte Berechnung"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   5400
      Width           =   3735
   End
   Begin VB.Label Label23 
      Caption         =   "Ost ist positiv"
      Height          =   255
      Left            =   3000
      TabIndex        =   39
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label19 
      Caption         =   "Ost ist positiv"
      Height          =   255
      Left            =   3000
      TabIndex        =   38
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Zeit in UT"
      Height          =   255
      Left            =   4800
      TabIndex        =   37
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Zeit u. Ort (H:M:S)"
      Height          =   255
      Left            =   360
      TabIndex        =   36
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label L_Zeit_Ort_2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Left            =   1920
      TabIndex        =   35
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label22 
      Caption         =   "Zeitzone"
      Height          =   255
      Left            =   840
      TabIndex        =   34
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "Grad"
      Height          =   255
      Left            =   3960
      TabIndex        =   33
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label20 
      Caption         =   "Zeit u. Ort (dezimal)"
      Height          =   255
      Left            =   360
      TabIndex        =   32
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label L_Zeit_Ort 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Left            =   1920
      TabIndex        =   31
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Left            =   5400
      TabIndex        =   30
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label16 
      Caption         =   "Geo. Länge"
      Height          =   255
      Left            =   840
      TabIndex        =   29
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Left            =   5400
      TabIndex        =   28
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label L_Zeit2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Left            =   1920
      TabIndex        =   27
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label15 
      Caption         =   "Zeit"
      Height          =   255
      Left            =   360
      TabIndex        =   26
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label14 
      Caption         =   "Grad"
      Height          =   255
      Left            =   3960
      TabIndex        =   25
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label L_GMST_Grad 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Left            =   5400
      TabIndex        =   24
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label L_GMST_Zeit 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Left            =   1920
      TabIndex        =   23
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label13 
      Caption         =   "GMST Zeit"
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "GMST_Grad"
      Height          =   255
      Left            =   3960
      TabIndex        =   21
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "S"
      Height          =   255
      Left            =   4200
      TabIndex        =   20
      Top             =   480
      Width           =   135
   End
   Begin VB.Label Label10 
      Caption         =   "M"
      Height          =   255
      Left            =   3600
      TabIndex        =   19
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "H"
      Height          =   255
      Left            =   3000
      TabIndex        =   18
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "Tag"
      Height          =   255
      Left            =   840
      TabIndex        =   17
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "Monat"
      Height          =   255
      Left            =   1440
      TabIndex        =   16
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Jahr"
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "JD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Grad"
      Height          =   255
      Left            =   3960
      TabIndex        =   13
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Zeit GMST 0h"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label L_Zeit 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label L_Grad 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label L_JD 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   2400
      Width           =   1695
   End
End
Attribute VB_Name = "TestJulianischesDatum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command2_Click()
    Dim JD As Double
    Dim Std As Double
    Dim Min As Double
    Dim Sek As Double
    
    Dim D As Double
    Dim M As Double
    Dim Y As Double
    
    
    Dim Stunde As Double
    Dim Minuten As Double
    Dim Sekunde As Double
    Dim Zeitzone As Double
    Dim Laenge As Double
    
    Dim t As Double
    Dim GMST_Zeit_s As Double
    Dim GMST_Zeit_h As Double
    Dim GMST_24 As Double
    Dim LokalTime As Double
    Dim GMSTindividualTime As Double
    Dim GMSTindividualOrt As Double
    
    
  
  
  '1.  Julianische Datum JD berechnen
  
  
    D = Zahl(T_Tag)
    M = Zahl(T_Monat)
    Y = Zahl(T_Jahr)
    INISetValue IniFileName, "Datum", "Tag", T_Tag
    INISetValue IniFileName, "Datum", "Monat", T_Monat
    INISetValue IniFileName, "Datum", "Jahr", T_Jahr
  
    Stunde = Zahl(T_Stunden)
    Minuten = Zahl(T_Minuten)
    Sekunde = Zahl(T_Sekunden)
    Zeitzone = Zahl(T_Zeitzone)

    INISetValue IniFileName, "Zeit", "Stunden", T_Stunden
    INISetValue IniFileName, "Zeit", "Minuten", T_Minuten
    INISetValue IniFileName, "Zeit", "Sekunden", T_Sekunden
    INISetValue IniFileName, "Zeit", "Zeitzone", T_Zeitzone
  
  
  Laenge = Zahl(T_Laenge)
  INISetValue IniFileName, "Koordinaten", "Länge", T_Laenge
  
  
    

  JD = GetJulianDate(Y, M, D, 0, 0, 0)
  
  L_JD = JD
  
  
' 2. Sternzeit in Greenwich berechnen
' Berechne die mittlere Sternzeit von Greenwich um 0 h UT zum gewünschten Datum.
' Addiere zum Ergebnis von 1) das Produkt t * 1.00273790935
' Der Faktor 1.002 737 909 35 berücksichtigt, dass die Sternzeit um so viel schneller abläuft als die Sonnenzeit.
' Das Resultat ist zum Schluss wieder auf [0; 24) zu normieren.
' Soll die mittlere Sternzeit nicht für Greenwich, sondern für einen Ort der geografischen Länge L° ,
' addiere man zum Resultat L/15      (positiv gezählt nach Osten, negativ nach Westen)

  t = (JD - 2451545#) / 36525
  
  
  GMST_Zeit_s = 24110.54841 + 8640184.812866 * t + 0.093104 * t * t - 0.0000062 * t * t * t
  GMST_Zeit_h = GMST_Zeit_s / 3600
            L_GMST_Zeit = GMST_Zeit_h
  GMST_24 = CutTime(GMST_Zeit_h)
  L_Zeit = GMST_24
  
  'Lokale Zeit auf siderische Zeit umgerechnet
  LokalTime = (Stunde + Minuten / 60 + Sekunde / 3600) * 1.00273790935
 
  GMSTindividualTime = GMST_24 + LokalTime
  
  GMSTindividualTime = CutTime(GMSTindividualTime)
  
  L_Zeit2 = GMSTindividualTime
  
  'Geographische Länge berücksichtigen
  GMSTindividualOrt = GMSTindividualTime + Laenge / 15
  
  GMSTindividualOrt = CutTime(GMSTindividualOrt)
  
  L_Zeit_Ort = GMSTindividualOrt
  
  ZeitDezToHMS GMSTindividualOrt, Std, Min, Sek
  
  L_Zeit_Ort_2 = (Std & ":" & Min & ":" & Sek)
 
 
 
'
'  GMST_Grad = 100.46061837 + 36000.770053608 * t + 0.000387933 * t * t - ((t * t * t) / 38710000)
'  L_GMST_Grad = GMST_Grad
'  Grad_Integer = Int(GMST_Grad / 360)
'  Grad_360 = GMST_Grad - (Grad_Integer * 360)
'  L_Grad = Grad_360
    
  
End Sub

Private Sub Form_Load()


    T_Tag = INIGetValue(IniFileName, "Datum", "Tag")
    T_Monat = INIGetValue(IniFileName, "Datum", "Monat")
    T_Jahr = INIGetValue(IniFileName, "Datum", "Jahr")
    
    T_Stunden = INIGetValue(IniFileName, "Zeit", "Stunden")
    T_Minuten = INIGetValue(IniFileName, "Zeit", "Minuten")
    T_Sekunden = INIGetValue(IniFileName, "Zeit", "Sekunden")
    T_Zeitzone = INIGetValue(IniFileName, "Zeit", "Zeitzone")
    
    T_Laenge = INIGetValue(IniFileName, "Koordinaten", "Länge")


End Sub
