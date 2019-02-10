Attribute VB_Name = "GlobalVariables"
Option Explicit

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
    S As Double
End Type

Public SimOffline As Boolean
Public IniFileName As String


