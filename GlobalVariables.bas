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

Public Type AzAlt
    Az As Double        ' Azimut as randian
    Alt As Double       ' Altitude as radian
End Type

Public Type StarDescription
    ProperName As String
    Bayer As String
    Constellation As String
    Flamsteed As String
    RA As Double
    DEC As Double
    Mag As Double
    StarDsc1 As String
    StarDsc2 As String
    StarDsc3 As String
    StarDsc4 As String
    StarDsc5 As String
End Type





Public SimOffline As Boolean
Public IniFileName As String
Public AlignmentStarArray() As StarDescription


