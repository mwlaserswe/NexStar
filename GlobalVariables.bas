Attribute VB_Name = "GlobalVariables"
'ToDo

Option Explicit

Public Const Pi = 3.14159265359
Public Const EncoderResolution = 726559
Public Const SidConst = 1.00273790935

Public Enum ProtokollMode
    Send = 0
    Receive = 1
End Enum

Public Type MyDate
    YY As Double
    MM As Double
    DD As Double
End Type

Public Type MyTime
    TimeDec As Double
    H As Double
    M As Double
    s As Double
End Type

Public Type GeoDegMinSec
    deg As Double
    Min As Double
    Sec As Double
    Sign As String
End Type

Public Type GeoCoordinates
    Latitude As Double
    Longitude As Double
End Type

Public Type Vector
  X As Double
  Y As Double
  z As Double
End Type

Public Type RaDec
    Ra As Double        ' Rectascension as randian
    Dec As Double       ' Declination as radian
End Type

Public Type AzAlt
    Az As Double        ' Azimut as randian
    Alt As Double       ' Altitude as radian
End Type

Public Type StarDescription
    Index As Long
    ProperName As String
    Bayer As String
    Constellation As String
    Flamsteed As String
    Ra As Double
    Dec As Double
    Mag As Double
    StarDsc1 As String
    StarDsc2 As String
    StarDsc3 As String
    StarDsc4 As String
    StarDsc5 As String
End Type





Public SimOffline As Boolean
Public ErrorCount As Long
Public CommTest As Boolean
Public IniFileName As String
Public CommFileName As String
Public AlignmentStarArray() As StarDescription

Public TestStatus As Boolean
Public StatusMoving As Integer


Public ObserverDateTimeUT As Date
Public ObserverTimeUT As MyTime
Public ObserverLatt As GeoDegMinSec
Public ObserverLong As GeoDegMinSec
Public ObserverRaDec As RaDec
Public ObserverAzAlt As AzAlt
Public GlobalOffset As AzAlt

Public GlbSiderialTime As Double
Public GlbOberverPos As GeoCoordinates

' Main Horizontal System für die Matrixmetode in [radian]
' mathematischer Sinn gegen den Uhtzeigersinn (CCW)
Public MatrixSystemSoll As AzAlt
Public MatrixSystemIst As AzAlt

Public MotorIncrSystem As Double       ' Horizontalsystem in [Increments] 0..726559 [CW]
Public AzAltSystem As Double           ' Horzontsystem in [radian] aus RA DEC berechnet


'==== Calibration ====
Public GlbCalibStatus As Integer        ' 0: not calibrated
                                        ' 1: 1 point calibration
                                        ' 2: 2 point calibration

'==== Init Time ====
Public Cal_InitTime As Double
Public TransformationMatrix(10, 10) As Double

'==== Reference Star 1 ====
Public Cal_RaStar_1 As Double
Public Cal_DecStar_1 As Double
Public Cal_TelHorizAngle_1 As Double
Public Cal_TelElevAngle_1 As Double
Public Cal_Time_1 As Double

'==== Reference Star 2 ====
Public Cal_RaStar_2 As Double
Public Cal_DecStar_2 As Double
Public Cal_TelHorizAngle_2 As Double
Public Cal_TelElevAngle_2 As Double
Public Cal_Time_2 As Double

Public TrackingisON As Boolean
Public DiffMotorIncr As AzAlt
Public MatrixLastCalc As AzAlt
Public MatrixDiffCalc As AzAlt
Public MotorLastCalc As AzAlt
Public MotorDiffCalc As AzAlt

Public MatrixSystemDiffPerSec As AzAlt
Public TrackingSpeed As AzAlt

'=== Test Only ===
Public LastVal As AzAlt
Public JetztTime As Double
                



'=== Communication Test ===
Public TestCommHandheldToMotor As Boolean
Public TestCommMotorToHandheld As Boolean
Public NexStarChar1 As String
Public NexStarChar2 As String
Public NexStarChar3 As String

'=== Visualisieung ===
Public GlbScale As Double
Public GlbCx As Double
Public GlbCY As Double


