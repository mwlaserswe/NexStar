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


Public m11 As Double
Public m12 As Double
Public m13 As Double
Public m21 As Double
Public m22 As Double
Public m23 As Double
Public m31 As Double
Public m32 As Double
Public m33 As Double
Public m14 As Double
Public m15 As Double
Public m16 As Double
Public m24 As Double
Public m25 As Double
Public m26 As Double
Public m34 As Double
Public m35 As Double
Public m36 As Double
