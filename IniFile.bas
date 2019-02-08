Attribute VB_Name = "IniOld"
Option Explicit

Private Declare Function WritePrivateProfileString Lib _
        "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal _
        lpKeyName As Any, ByVal lpString As Any, ByVal _
        lpFileName As String) As Long
        
Private Declare Function GetPrivateProfileString Lib _
        "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal _
        lpKeyName As Any, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize _
        As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileSection Lib _
        "kernel32" Alias "WritePrivateProfileSectionA" _
        (ByVal lpAppName As String, ByVal lpString As _
        String, ByVal lpFileName As String) As Long
        
Private Declare Function GetPrivateProfileSection Lib _
        "kernel32" Alias "GetPrivateProfileSectionA" _
        (ByVal lpAppName As String, ByVal lpReturnedString _
        As String, ByVal nSize As Long, ByVal lpFileName _
        As String) As Long
        
Private Declare Function GetPrivateProfileSectionNames Lib _
        "kernel32" Alias "GetPrivateProfileSectionNamesA" _
        (ByVal lpszReturnBuffer As String, ByVal nSize As _
        Long, ByVal lpFileName As String) As Long




Public Sub INISetValue(ByVal Path$, ByVal Sect$, ByVal key$, ByVal value$)

Dim result&
Dim antwort As Integer
    
    result = WritePrivateProfileString(Sect, key, value, Path)
'    If Result = 0 Then
'      antwort = MsgBox("Filename: " & Path & vbCr & _
'                "Section:  " & Sect & vbCr & _
'                "Key:        " & Key & vbCr & _
'                "Value:     " & Value, _
'                vbCritical Or vbOKCancel, "Error SetValue of INI-File")
'      If antwort = vbCancel Then End
'    End If
End Sub

Public Function INIGetValue(ByVal Path$, ByVal Sect$, ByVal key$) As String
  
Dim result&, Buffer$
Dim antwort As Integer
    
    Buffer = space$(32767)
    result = GetPrivateProfileString(Sect, key, vbNullString, Buffer, Len(Buffer), Path)
    INIGetValue = left$(Buffer, result)
'    If Result = 0 Then
'      antwort = MsgBox("Filename: " & Path & vbCr & _
'                "Section:  " & Sect & vbCr & _
'                "Key:        " & Key, _
'                vbCritical Or vbOKCancel, "Error GetValue of INI-File")
'      If antwort = vbCancel Then End
'    End If
End Function

Public Sub INIGetArray(ByVal Path$, ByVal Sect$, xArray() As String)
  Dim result&, Buffer$
  Dim l%, p%, z%
    'String lesen
    Buffer = space(32767)
    result = GetPrivateProfileSection(Sect, Buffer, Len(Buffer), Path)
    
    Buffer = left$(Buffer, result)
    
    If Buffer <> "" Then
      'String mit Trennzeichen Chr$(0) in ein Feld umwandeln
      l = 1
      ReDim xArray(0)
      Do While l < result
        p = InStr(l, Buffer, Chr$(0))
        If p = 0 Then Exit Do
        
        xArray(z) = Mid$(Buffer, l, p - l)
        z = z + 1
        ReDim Preserve xArray(0 To z)
        l = p + 1
      Loop
    End If
End Sub


