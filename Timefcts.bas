Attribute VB_Name = "Timefcts"
Option Explicit

' Use the PtrSafe attribute for x64 installations
Private Declare Function FileTimeToLocalFileTime Lib "Kernel32" (lpFileTime As FILETIME, ByRef lpLocalFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "Kernel32" (lpLocalFileTime As FILETIME, ByRef lpFileTime As FILETIME) As Long
Private Declare Function SystemTimeToFileTime Lib "Kernel32" (lpSystemTime As SYSTEMTIME, ByRef lpFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "Kernel32" (lpFileTime As FILETIME, ByRef lpSystemTime As SYSTEMTIME) As Long

Public Type FILETIME
  LowDateTime As Long
  HighDateTime As Long
End Type

Public Type SYSTEMTIME
  Year As Integer
  Month As Integer
  DayOfWeek As Integer
  Day As Integer
  Hour As Integer
  Minute As Integer
  Second As Integer
  Milliseconds As Integer
End Type


'===============================================================================
' Convert local time to UTC
'===============================================================================
Public Function UtcTime(LOCALTIME As Date) As Date
  Dim oLocalFileTime As FILETIME
  Dim oUtcFileTime As FILETIME
  Dim oSystemTime As SYSTEMTIME

  ' Convert to a SYSTEMTIME
  oSystemTime = DateToSystemTime(LOCALTIME)

  ' 1. Convert to a FILETIME
  ' 2. Convert to UTC time
  ' 3. Convert to a SYSTEMTIME
  Call SystemTimeToFileTime(oSystemTime, oLocalFileTime)
  Call LocalFileTimeToFileTime(oLocalFileTime, oUtcFileTime)
  Call FileTimeToSystemTime(oUtcFileTime, oSystemTime)

  ' Convert to a Date
  UtcTime = SystemTimeToDate(oSystemTime)
End Function



'===============================================================================
' Convert UTC to local time
'===============================================================================
Public Function LOCALTIME(UtcTime As Date) As Date
  Dim oLocalFileTime As FILETIME
  Dim oUtcFileTime As FILETIME
  Dim oSystemTime As SYSTEMTIME

  ' Convert to a SYSTEMTIME.
  oSystemTime = DateToSystemTime(UtcTime)

  ' 1. Convert to a FILETIME
  ' 2. Convert to local time
  ' 3. Convert to a SYSTEMTIME
  Call SystemTimeToFileTime(oSystemTime, oUtcFileTime)
  Call FileTimeToLocalFileTime(oUtcFileTime, oLocalFileTime)
  Call FileTimeToSystemTime(oLocalFileTime, oSystemTime)

  ' Convert to a Date
  LOCALTIME = SystemTimeToDate(oSystemTime)
End Function



'===============================================================================
' Convert a Date to a SYSTEMTIME
'===============================================================================
Private Function DateToSystemTime(Value As Date) As SYSTEMTIME
  With DateToSystemTime
    .Year = Year(Value)
    .Month = Month(Value)
    .Day = Day(Value)
    .Hour = Hour(Value)
    .Minute = Minute(Value)
    .Second = Second(Value)
  End With
End Function



'===============================================================================
' Convert a SYSTEMTIME to a Date
'===============================================================================
Private Function SystemTimeToDate(Value As SYSTEMTIME) As Date
  With Value
    SystemTimeToDate = _
      DateSerial(.Year, .Month, .Day) + _
      TimeSerial(.Hour, .Minute, .Second)
  End With
End Function

