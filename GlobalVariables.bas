Attribute VB_Name = "GlobalVariables"
Option Explicit

Public Const Pi = 3.14159265359
Public Const EncoderResolution = 726559

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

Public Type GeoCoord
    Deg As Double
    Min As Double
    Sec As Double
    Sign As String
End Type

Public Type Vector
  x As Double
  Y As Double
  z As Double
End Type

Public SimOffline As Boolean
Public IniFileName As String


