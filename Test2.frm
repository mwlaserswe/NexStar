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

Dim LMN_EquaMatrix(10, 10) As Double
Dim LMN_HorizMatrix(10, 10) As Double
Dim AusgangsMatrix(10, 10) As Double

Dim dmy0 As Double
Dim dmy1 As Double
Dim dmy2 As Double


Private Sub Command1_Click()
    ' matrix_method_rev_d.pdf  Seite 15,16
    ' Rückwärts rechnen. Siehe "Matrix Method.doc"
    ' Input: Object Alt and Az, ObserverLatt,
    ' Output: Object RA and DEC
    
    Dim SaturnAzAlt As AzAlt
    Dim LMN_Horizontal As Vector
    Dim LMN_HorizMatrix(10, 10) As Double
    Dim ObsLocMatrix(10, 10) As Double  ' Observer location Matrix
    
    SaturnAzAlt.Az = DegToRad(-51.6992)           ' A Azimut
    SaturnAzAlt.Alt = DegToRad(36.5405)           ' h Höhe
    
    LMN_Horizontal.X = Cos(SaturnAzAlt.Alt) * Cos(-SaturnAzAlt.Az)
    LMN_Horizontal.Y = Cos(SaturnAzAlt.Alt) * Sin(-SaturnAzAlt.Az)
    LMN_Horizontal.z = Sin(SaturnAzAlt.Alt)
    
    
    Dim LatitudeDeg As Double                   ' Observer’s latitude
    Dim LatitudeRad As Double
   
    LatitudeDeg = GeoToDez(ObserverLatt)
    LatitudeRad = DegToRad(LatitudeDeg)
   
    ObsLocMatrix(0, 0) = Cos(LatitudeRad - Pi / 2):     ObsLocMatrix(0, 1) = 0:     ObsLocMatrix(0, 2) = Sin(LatitudeRad - Pi / 2)
    ObsLocMatrix(1, 0) = 0:                             ObsLocMatrix(1, 1) = 1:     ObsLocMatrix(1, 2) = 0
    ObsLocMatrix(2, 0) = -Sin(LatitudeRad - Pi / 2):    ObsLocMatrix(2, 1) = 0:     ObsLocMatrix(2, 2) = Cos(LatitudeRad - Pi / 2)
   
    Dim InverseMatrix(10, 10) As Double             'Zeile, Spalte
    Calculate_Inverse 3, ObsLocMatrix, InverseMatrix
    
    LMN_HorizMatrix(0, 0) = LMN_Horizontal.X
    LMN_HorizMatrix(1, 0) = LMN_Horizontal.Y
    LMN_HorizMatrix(2, 0) = LMN_Horizontal.z
    
    Dim LMN_EquaMatrix(10, 10) As Double
    MatrixProduct InverseMatrix, 3, 3, LMN_HorizMatrix, 3, 1, LMN_EquaMatrix
    
    
    dmy0 = LMN_EquaMatrix(0, 0)
    dmy1 = LMN_EquaMatrix(1, 0)
    dmy2 = LMN_EquaMatrix(2, 0)
 
    
    Dim Object As RaDec
    Object.Dec = arcsin(LMN_EquaMatrix(2, 0))
    
    Dim HourAngle As Double
    HourAngle = arcsin(-LMN_EquaMatrix(1, 0) / Cos(Object.Dec))
    
    Dim LongitudeDeg As Double                  ' Observer’s longitude
    Dim LongitudeRad As Double
    LongitudeDeg = GeoToDez(ObserverLong)
    LongitudeRad = DegToRad(LongitudeDeg)
  
    Dim LocalDateTime As Date
    Dim LocalDate As MyDate
    Dim LocalTimeUT As MyTime
    LocalDateTime = ObserverDateTimeUT
    'Calculate siderial time at Greenwich
    LocalDate.YY = Year(LocalDateTime)
    LocalDate.MM = Month(LocalDateTime)
    LocalDate.DD = Day(LocalDateTime)
    LocalTimeUT.H = Hour(LocalDateTime)
    LocalTimeUT.M = Minute(LocalDateTime)
    LocalTimeUT.s = Second(LocalDateTime)

    Dim SiderialTime As MyTime
    Dim SiderialTimeRad As Double
    SiderialTime = GMST(LocalDate, LocalTimeUT)
    SiderialTimeRad = TimeToRad(SiderialTime)
    
     Object.Ra = SiderialTimeRad - LongitudeRad - HourAngle
  
End Sub

