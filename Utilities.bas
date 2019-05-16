Attribute VB_Name = "Utilities"
Option Explicit

  Private Declare Function FindFirstFile Lib "Kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
  Private Declare Function FindNextFile Lib "Kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
  Private Declare Function FindClose Lib "Kernel32" (ByVal hFindFile As Long) As Long
  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
  

    
  Const LB_SETHORIZONTAL = &H194
  
  
  Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
  End Type
  
  Const MAX_PATH = 259
  
  Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
  End Type
  
  Const FILE_ATTRIBUTE_ARCHIVE = &H20
  Const FILE_ATTRIBUTE_COMPRESSED = &H800
  Const FILE_ATTRIBUTE_DIRECTORY = &H10
  Const FILE_ATTRIBUTE_HIDDEN = &H2
  Const FILE_ATTRIBUTE_NORMAL = &H80
  Const FILE_ATTRIBUTE_READONLY = &H1
  Const FILE_ATTRIBUTE_SYSTEM = &H4
  Const FILE_ATTRIBUTE_TEMPORARY = &H100

  Public Const maxlist = 100


Declare Function GetShortPathName Lib "Kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public bitMaske(0 To 7) As Byte

'=============================================
'= Definitionen für Function T (Michael)
'=============================================
Public SpracheFileName As String   'Variable, die den Dateinamen und den Pfad, der Ausgabetextsammlng angibt
Public SprachFileSprache As String 'Variable, mit der man die Sprache der Ausgabe einstellt (z.B. DE für Deutsch)
Public ErrorCollection As New Collection
Public SprachFileNummernAnzeigen As Boolean
Public sprachcollection As New Collection 'Hier werden die Einträge der Ausgabetext-Sammlung gespeichert.
                                          'Die Ausgabetexte sind mit einem eineindeutigen Schlüssel abrufbar.
                                          'Dieser Schlüssel setzt sich aus der Textnummer und der einem Sprachkürzel (z.B. "DE") zsammen.
                                          'Beispiele für einen solchen Schlüssel sind "815EN" oder "753RO"

Private LanguageCollection As New Collection 'Speicher für alle vorkommenden Sprachen z.B. EN für Englisch , DE für Deutsch...
Private country_abbrev As String
Private complete_entry  As String
Private Complete_TextNr_Collection  As New Collection

Private flagMMI As Boolean

'=============================================
'= Definitionen für MWMotionDriver
'=============================================
Private Declare Function SendCommand Lib "C:\MWMotionDriver\MWMotionDriver.dll" _
         (ByVal Command As String, _
          ByVal Parameters As String, _
          ByVal reply As String, _
          ByVal maxLength As Long) As Long








Public Function Zahl(Txt As String) As Double
' Wandelt die Zahl in einem String in eine Zahl um
' dabei werden "," in "." umgewandelt und alle Zeichen
' die nicht passen in Leerzeichen gewandelt
' 22.07.2002 Exponent möglich
'            wenn keine Ziffern vorhanden sind, wird Err.number = 1 gesetzt
  Dim i As Integer
  Dim s As String
  Dim noVorz As Boolean, noKomma As Boolean, noExpo As Boolean, haveDigits As Boolean
  s = ""
  For i = 1 To Len(Txt)
    Select Case Mid(Txt, i, 1)
      Case "+", "-"
        If Not noVorz Then
          s = s + Mid(Txt, i, 1)
          noVorz = True
        Else
          Exit For
        End If
      Case ",", "."
        If Not noKomma Then
          s = s + "."
          noKomma = True
          noVorz = True
        Else
          Exit For
        End If
      Case "0" To "9"
        s = s + Mid(Txt, i, 1)
        noVorz = True
        haveDigits = True
      Case "&"
        s = s + Mid(Txt, i, 2)
        noVorz = True
      Case "E", "e"
        If Not noExpo Then
           s = s + Mid(Txt, i, 1)
          noVorz = False
          noKomma = True
          noExpo = True
        Else
          Exit For
        End If
      Case " "
      Case Else
        If noVorz Then Exit For
    End Select
  Next i
  If Not haveDigits Then
    Err.Number = 1
    Err.Description = "Zahl set to 0. No Digits in String"
  End If
  Zahl = Val(s)
End Function

Public Function TestZeile(varName As String, Zeile As String, wert As String) As Boolean
' Überprüft ob die Zeile den Parameter varName enthält
' dabei muß varName am Anfang der Zeile stehen
' Groß/kleinschreibung bleibt unberücksichtigt
' Ist kein Wert vorhanden, liefert die Funktion FALSE
' Der Wert darf auch Leerzeichen enthalten
' Input: varName : Name des Parameters
'        zeile : enthält Parameternamen am anfang und wert
' Output : wert : string ohne Parameternamen
' Return : True : zeile enthält varName
'
' 18.02.02 varName auf Länge überprüft, wert ohne führende Leerzeichen

  Dim SpacePos As Integer

  TestZeile = False
  If InStr(1, Zeile, varName, vbTextCompare) <> 1 Then Exit Function
  SpacePos = InStr(1, Zeile, " ")
  If SpacePos > 0 Then
    If Len(varName) <> SpacePos - 1 Then Exit Function
  Else
    Exit Function
  End If
  TestZeile = True
  wert = Trim(Mid(Zeile, Len(varName) + 2))

End Function

' Dateiname ohne Pfad extrahieren
Function DateiName$(vollständig$)
  Dim pos&
  pos = InStrRev(vollständig, "\")
  If pos <> 0 Then
    DateiName = Mid(vollständig, pos + 1)
  Else
    DateiName = vollständig
  End If
End Function

' Pfad ohne Dateinamen extrahieren
Function Pfad$(vollständig$)
  Dim pos&
  pos = InStrRev(vollständig, "\")
  If pos <> 0 Then
    Pfad = Left(vollständig, pos)
  End If
End Function

' Laufwerksname extrahieren
Function Laufwerk$(vollständig$)
  Dim pos&
  On Error Resume Next
  If Left(vollständig, 2) = "\\" Then
    pos = InStr(3, vollständig, "\")
    Laufwerk = Left(vollständig, pos - 1)
  Else
    pos = InStr(vollständig, ":")
    Laufwerk = Left(vollständig, pos)
  End If
End Function

' Dateikennung extrahieren
Function Kennung$(DateiName$)
  Dim pos&
  pos = InStrRev(DateiName, ".")
  If pos <> 0 Then
    Kennung = Mid(DateiName, pos + 1)
  End If
End Function

' Dateikennung entfernen
Function NameOhneKennung$(DateiName$)
  Dim pos&
  pos = InStrRev(DateiName, ".")
  If pos <> 0 Then
    NameOhneKennung = Left(DateiName, pos - 1)
  Else
    NameOhneKennung = DateiName
  End If
End Function


Public Function InStrRev(String1 As String, String2 As String) As Long
  Dim Pos1 As Long
  Dim Pos2 As Long
  
  If (String1 = "") Or (String2 = "") Or (Len(String2) > Len(String1)) Then
      InStrRev = 0
  Else
      Pos2 = 0
      Do
        Pos1 = Pos2
        Pos2 = InStr(Pos1 + 1, String1, String2)
      Loop While Pos2 > 0
      InStrRev = Pos1
  End If
End Function

Public Function KommaToPunkt(Txt As String) As String
Dim i As Integer
Dim s As String

  For i = 1 To Len(Txt)
    Select Case Mid(Txt, i, 1)
      Case ","
          s = s + "."
      Case Else
        s = s + Mid(Txt, i, 1)
    End Select
  Next i
  KommaToPunkt = s

End Function

Public Function VbGetShortPathName(LongPathName As String)
Dim sBuffer As String, lLen As Long

    sBuffer = Space$(512)
    lLen = GetShortPathName(LongPathName, sBuffer, Len(sBuffer))
    VbGetShortPathName = Left$(sBuffer, lLen)

End Function

Public Function DelCrtlChar(Srt1) As String
    Dim i As Integer
    
    DelCrtlChar = ""
    For i = 1 To Len(Srt1)
      If Mid(Srt1, i, 1) > Chr(31) And Mid(Srt1, i, 1) < Chr(128) Then
        DelCrtlChar = DelCrtlChar + Mid(Srt1, i, 1)
      End If
    Next i
    
End Function

Public Sub InitVariablen()
'  Static isInit As Boolean
'  Dim i As Integer, n As Long
'
'  If isInit Then Exit Sub
'   isInit = True
'
'    For i = 0 To 7
'      bitMaske(i) = 2 ^ i
'    Next i
'
'    n = 1
'    For i = 0 To 15
'      bitMaske_Int(i) = n
'      bitMaskeInv_Int(i) = Not n
'      n = n * 2
'    Next i

End Sub

Public Sub BitSet(daten As Byte, BitNr As Integer)
  daten = daten Or bitMaske(BitNr)
End Sub
Public Sub BitReset(daten As Byte, BitNr As Integer)
  daten = daten And (Not bitMaske(BitNr))
End Sub


Public Sub BitSet_Int(daten As Integer, BitNr As Integer)
'  daten = daten Or bitMaske_Int(BitNr)
End Sub
Public Sub BitReset_Int(daten As Integer, BitNr As Integer)
'  daten = daten And bitMaskeInv_Int(BitNr)
End Sub


Public Function BitTest(daten As Byte, BitNr As Integer) As Boolean
'  BitTest = False
'  If (daten And bitMaske(BitNr)) <> 0 Then
'    BitTest = True
'  End If
End Function


Public Sub PrintControlCaption(WertVar As Variant, Steuerelement As Variant)
  Dim WertOld As Variant
  Dim wert As String
  
  '1. sicherstellen, daß alle Werte als String vorliegen
  If Not (VarType(WertVar) = 8) Then
    wert = CStr(WertVar)
  Else
    wert = WertVar
  End If
  
  If (TypeOf Steuerelement Is Label) _
     Or (TypeOf Steuerelement Is Form) _
     Or (TypeOf Steuerelement Is Frame) _
     Or (TypeOf Steuerelement Is CommandButton) _
  Then
    WertOld = Steuerelement.Caption
    If (WertOld <> wert) Then         'hat sich die Wert geändert?
      Steuerelement.Caption = wert
      
    End If
  End If
  If TypeOf Steuerelement Is TextBox Then
    WertOld = Steuerelement.Text
    If (WertOld <> wert) Then         'hat sich die Wert geändert?
      Steuerelement.Text = wert
    End If
  End If
End Sub
Public Sub PrintControlFehler(WertVar As Variant, Steuerelement As Variant)
  Dim WertOld As Variant
  Dim wert As String
  
  '1. sicherstellen, daß alle Werte als String vorliegen
  If Not (VarType(WertVar) = 8) Then
    wert = CStr(WertVar)
  Else
    wert = WertVar
  End If
  
  If (TypeOf Steuerelement Is Label) _
     Or (TypeOf Steuerelement Is Form) _
     Or (TypeOf Steuerelement Is Frame) _
     Or TypeOf Steuerelement Is TextBox _
     Or (TypeOf Steuerelement Is CommandButton) _
  Then

WertOld = Steuerelement
    If (WertOld <> wert) Then         'hat sich die Wert geändert?
'      ProtokollMessage wert

    End If
     
  End If
  
End Sub

Public Sub PrintControlColor(wert As Long, Steuerelement As Variant)
  Dim WertOld As Long
  
  WertOld = Steuerelement.BackColor
  If (WertOld <> wert) Then         'hat sich die Wert geändert?
    Steuerelement.BackColor = wert
  End If
End Sub

Public Sub PrintControlForeColor(wert As Long, Steuerelement As Variant)
  Dim WertOld As Long
  
  WertOld = Steuerelement.ForeColor
  If (WertOld <> wert) Then         'hat sich die Wert geändert?
    Steuerelement.ForeColor = wert
  End If
End Sub


'**************************************************************************************************
'                        Ab hier alle Funktionen für Recent Files
'**************************************************************************************************

'// Updates the file menu:
'// either adds a new item at the top for new file
'// or moves the existing item to the top
Public Sub UpdateFileMenu(strFileName As String, MenüForm As Form)
    '// Check if the open filename is already in the File menu control array.
    If OnRecentFilesList(strFileName, MenüForm) Then
        '// move the existing item to the top
        MoveRecentFiles strFileName, MenüForm
    Else
        '// add a new item at the top for new file
        WriteRecentFiles strFileName, MenüForm
    End If
    '// Update the list of the most recently opened files in the File menu control array.
'    GetRecentFiles
End Sub


'// adds a new file to the top of the list
Private Sub WriteRecentFiles(strFileName As String, MenüForm As Form)
    Dim i       As Integer
    Dim strFile As String
    Dim strKey  As String
    '// Move all items down one
    '// start from one up from bottom, so that bottom
    '// item is overwritten
    
    'Trennlinie sichtbar machen
    MenüForm.mnuRecentFileSep.Visible = True
    
    For i = MenüForm.mnurecentfile().Count - 1 To 1 Step -1
        strFile = MenüForm.mnurecentfile(i).Tag
        If strFile <> "" Then
            MenüForm.mnurecentfile(i + 1).Tag = strFile
            MenüForm.mnurecentfile(i + 1).Caption = i + 1 & " " & FilenameKürzen(strFile)
            MenüForm.mnurecentfile(i + 1).Visible = True
            strKey = "RecentFile" & (i + 1)
            ' Eintrag in Windows-Registry
            SaveSetting App.Title, "RecentFiles", strKey, strFile
        End If
    Next i
    
    'Neuen Eintrag an die erste Stelle schreiben
    MenüForm.mnurecentfile(1).Tag = strFileName
    MenüForm.mnurecentfile(1).Caption = "1 " & FilenameKürzen(strFileName)
    MenüForm.mnurecentfile(1).Visible = True
    
    ' Eintrag in Windows-Registry
    SaveSetting App.Title, "RecentFiles", "RecentFile1", strFileName
' SaveSetting App.Title, "RecentFiles", &quot;RecentFile1&quot;, strFileName
End Sub


'// This sub moves the specified file to the top of the list
'// from wherever it is
Private Sub MoveRecentFiles(strFileName As String, MenüForm As Form)
    Dim intLocation  As Integer
    Dim i            As Integer
    Dim strFile      As String
    Dim strKey       As String
    '// Get location of specified file
    For intLocation = 1 To MenüForm.mnurecentfile.Count
        strFile = MenüForm.mnurecentfile(intLocation).Tag
        strKey = "RecentFile" & intLocation
        strFile = GetSetting(App.Title, "RecentFiles", strKey)
        If strFile = strFileName Then
            '// found item
            Exit For
        End If
    Next
    '// Move all items down upto location of strFileName
    '// start from item before the loc of strFileName
    For i = intLocation - 1 To 1 Step -1
        strFile = MenüForm.mnurecentfile(i).Tag
    
        If strFile <> "" Then
            'um eins nach unten verschieben
            MenüForm.mnurecentfile(i + 1).Tag = strFile
            MenüForm.mnurecentfile(i + 1).Caption = i + 1 & " " & FilenameKürzen(strFile)
            MenüForm.mnurecentfile(i + 1).Visible = True
            ''// save new location
            strKey = "RecentFile" & (i + 1)
            SaveSetting App.Title, "RecentFiles", strKey, strFile
        End If
        
    Next i
    '// save deleted item at top
    MenüForm.mnurecentfile(1).Tag = strFileName
    MenüForm.mnurecentfile(1).Caption = "1 " & FilenameKürzen(strFileName)
    MenüForm.mnurecentfile(1).Visible = True
    strKey = "RecentFile1"
    SaveSetting App.Title, "RecentFiles", strKey, strFileName
End Sub


'// Checks to see if item is on the File list already
Private Function OnRecentFilesList(Filename, MenüForm As Form) As Boolean
    Dim i As Integer '// Counter variable.
    For i = 1 To MenüForm.mnurecentfile.Count
        If MenüForm.mnurecentfile(i).Tag = Filename Then
            OnRecentFilesList = True
            Exit Function
        End If
    Next i
End Function


'Kürzt einen Dateinamen, damit er ins Menü passt.
'Denkbar wäre:
' 1. - Nur den Dateinamen ohne Pfad anzeigen
' - Eine elegantere Lösung, in der man ansatzweise den Pfad erkennen kann
Public Function FilenameKürzen(Filename As String) As String
  Dim EchterDateiname As String
  Dim TempPfad As String
  Dim LetzterPfadTeil As String
  
' 1. Variante: Nur den Dateinamen ohne Pfad anzeigen
'  FilenameKürzen = DateiName$(Filename)

' 2.Variante: der letzte Pfad-teil wird ebenfalls angezeigt: ...\pfad\datei.ext
  EchterDateiname = DateiName$(Filename)
  TempPfad = Pfad$(Filename)
  'letzen Backslash eliminieren, so daß es wie ein Dateiname aussieht
  TempPfad = Mid(TempPfad, 1, Len(TempPfad) - 1)
  'Nun kann man die DateiName$-Funktion mißbrauchen
  LetzterPfadTeil = DateiName$(TempPfad)
  FilenameKürzen = "...\" & LetzterPfadTeil & "\" & EchterDateiname
  
End Function

'// gets all the recent files stored in the registry
Public Sub GetRecentFiles(MenüForm As Form)
    Dim i        As Integer
    Dim varFiles As Variant ' Variable to store the returned array.
    Dim strTitle As String
    Dim strPath  As String
    Dim strFile  As String
    Dim Item     As Menu
    '// App.Title and RECENT_FILE_KEY are constants defined in this module.
    '// hide all the items
    On Error Resume Next
    For Each Item In MenüForm.mnurecentfile
        Item.Visible = False
    Next
    
    If GetSetting(App.Title, "RecentFiles", "RecentFile1") = Empty Then
        '// no files on list, hide seperator
        MenüForm.mnuRecentFileSep.Visible = False
        Exit Sub
    Else
        '// load items
        For i = 1 To MenüForm.mnurecentfile.Count
        
            'get Filename
            strFile = GetSetting(App.Title, "RecentFiles", "RecentFile" & i)
            If strFile = "" Then
                '// no more items, exit for
                Exit For
            End If
            MenüForm.mnurecentfile(i).Tag = strFile
            MenüForm.mnurecentfile(i).Caption = i & " " & FilenameKürzen(strFile)
            MenüForm.mnurecentfile(i).Visible = True
        Next
        '// show seperator
        MenüForm.mnuRecentFileSep.Visible = True
    End If
End Sub
'**************************************************************************************************
'                        Recent Files   E N D E
'**************************************************************************************************






Public Sub PrintCommImage(PortNr As String, CommStatus As String, Pic As PictureBox)
  Dim TagOld As String
  Dim CommStatusOld As String
  
  CommStatusOld = Pic.Tag
  If CommStatusOld <> CommStatus Then
    Pic.Tag = CommStatus
    If CommStatus = True Then
      Pic.Cls
      Pic.BackColor = vbGreen
      Pic.Print PortNr
    Else
      Pic.Cls
      Pic.BackColor = vbRed
      Pic.Print PortNr
    End If
  End If
  
End Sub


Public Sub SepariereString(Zeile As String, WortArray() As String, Delimiter As String)
  Dim Pos1 As Long
  Dim Pos2 As Long
  Dim AnzahlWorte As Long               'Anzahl der Worte der Zeile
  
  ReDim WortArray(0 To 0)                         'WortArray löschen
  AnzahlWorte = 0
  Pos2 = 0
  
  If Delimiter = " " Then Zeile = Trim(Zeile)
  
  Zeile = Trim(Zeile)
  Do
    Pos1 = Pos2
    
    'Trennzeichen [CR]: [LF] werden überlesen
    If Delimiter = vbCr Then
      If Mid(Zeile, Pos1 + 1, 1) = vbLf Then
        Pos1 = Pos1 + 1                             'LF übergehen
      End If
    End If
    
     'Trennzeichen [Space]: [Space] werden überlesen
    If Delimiter = " " Then
      Do While Mid(Zeile, Pos1 + 1, 1) = " "
        Pos1 = Pos1 + 1                             'Space übergehen
      Loop
    End If
   
    Pos2 = InStr(Pos1 + 1, Zeile, Delimiter)      'nach Trennzeichen suchen
    If Pos2 > 0 Then                              'noch ein Trennzeichen in der Zeile
      WortArray(AnzahlWorte) = Mid(Zeile, Pos1 + 1, Pos2 - Pos1 - 1)
      ReDim Preserve WortArray(0 To UBound(WortArray) + 1)
      AnzahlWorte = AnzahlWorte + 1
    Else                                          'kein Trennzeichen mehr vorhanden
      WortArray(AnzahlWorte) = Mid(Zeile, Pos1 + 1)
    End If
  Loop While Pos2 > 0
End Sub

Public Sub Add_List(CurrentForm As Form, Listbox As Control, Text As String)

    Dim x&, Max&, Akt&, result&, cForm As Form

    Text = Text & "   "

    Listbox.AddItem Text
    If Listbox.ListCount > maxlist Then Listbox.RemoveItem 0
    Listbox.ListIndex = Listbox.ListCount - 1 'Letzten Eintrag hinterlegen
    
      
    Set cForm = CurrentForm
    Set cForm.Font = Listbox.Font
    
    For x = 0 To Listbox.ListCount - 1
      Akt = cForm.TextWidth(Listbox.List(x))
      If Akt > Max Then Max = Akt
    Next
    
    Max = Max \ Screen.TwipsPerPixelX
    result = SendMessage(Listbox.hwnd, LB_SETHORIZONTAL, _
                         Max, ByVal 0)
    
    Set cForm = Nothing
    
End Sub

Public Sub GetAllFiles(ByVal Pfad As String, ByVal Patt$, ByRef Field() As String)
  Dim Datei$, hFile&, FD As WIN32_FIND_DATA
  
  If Right(Pfad, 1) <> "\" Then Pfad = Pfad & "\"
  hFile = FindFirstFile(Pfad & Patt, FD)
  If hFile = 0 Then Exit Sub
  
  Do
    Datei = Left(FD.cFileName, InStr(FD.cFileName, Chr(0)) - 1)
    If (FD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) _
        = FILE_ATTRIBUTE_DIRECTORY Then
      If (Datei <> ".") And (Datei <> "..") Then
        Field(UBound(Field)) = "°" & Pfad & Datei
        ReDim Preserve Field(0 To UBound(Field) + 1)
        GetAllFiles Pfad & Datei, Patt, Field
      End If
    Else
      Field(UBound(Field)) = Pfad & Datei
      ReDim Preserve Field(0 To UBound(Field) + 1)
    End If
  Loop While FindNextFile(hFile, FD)
  
  FindClose hFile
End Sub



Public Function t_variable_texte(Index As Integer, DefaultText As String) As String
  Static Init         As Boolean
  Dim first_of_more   As Boolean
  Static language_supported As Boolean
  Static firsterror   As Boolean
  Static fileheaderexists As Boolean
  Dim textnumber      As Long
  Static firstEntryInErrorFile As Boolean
  
  Dim KommaPos        As Integer
  Dim first_entry     As Integer
  
  Dim Item2           As Variant
  Dim entry           As Variant
  Dim entry2          As Variant
  
  Dim languagefile    As String
  Dim single_line     As String
  Dim nextline        As String
  Dim language_entry  As String
  Dim old_entry       As String   'Hilfvariable (vorheriger Eintrag)
  Dim returnval       As String   'Rückgabewert von t
  Dim errorfile       As String
  Dim errorfilename   As String   'Dateiname des Error-Log
  Dim avail_language  As String   'Namen der verfügbaren Sprachen

  ''''''Variablendeklaration zu Ende
  
  If Not Init Then 'Beim ersten Aufruf der t- Funktion werden alle Einträge der Sammlung mit entsprechendem Schlüssel in eine Collecion (sprachcollection) geladen

      Init = True
      languagefile = FreeFile                 'Öffnen der Datei
      On Error GoTo openErr
      Open SpracheFileName For Input As languagefile
      Do While Not (Left(single_line, 1) = "@")
        Line Input #languagefile, single_line
        If (Left(single_line, 1) = "@") Then GoTo AtGefunden
      Loop
      Do While Not EOF(languagefile)
        Line Input #languagefile, single_line 'Zeilen der Datei einzeln einlesen
AtGefunden:
      If "@" = Left(single_line, 1) Then 'Falls erstes Zeichen in der Zeile eine "@" ist, so liegt ein Schlüssel und somit ein Eintrag vor
              first_of_more = True
              KommaPos = InStr(1, single_line, ",", vbTextCompare)  'Kommaposition bestimmen (um zu wissen, wo der Schlüssel aufhört und der Eintrag anfängt
              country_abbrev = Mid(single_line, KommaPos - 2, 2)    '"Länderkürzel" auslesen (z.B. "DE" oder "FR")
              complete_entry = Mid(single_line, 2, KommaPos - 2)    'Gesamten Schlüssel aus der Zeile auslesen
              language_entry = Mid(single_line, KommaPos + 1)       'Ausgabetext-Eintrag (z.B. "Roboter in Endposition") auslesen und speichern
            
          If Not (country_abbrev Like "[0-9]") Then ' MS 31012006 V1.0.2
            On Error Resume Next
              sAdd_Key 'Schlüssel zu der Schlüsselsammlung Complete_TextNr_Collection   hinzufügen

            On Error Resume Next
              sAdd_Language 'Sprachkürzel zu der Sprachensammlung LanguageCollection   hinzufügen

            On Error Resume Next
              sprachcollection.Add language_entry, complete_entry 'Eintrag des Ausgabetextes mit Schlüssel
          Else
            GoTo NoValidLine
          End If
      Else 'Erstes Zeichen kein "@" also ein mehrzeiliger Eintrag
                If first_of_more Then 'die erste Zeile eines mehrzeiligen Eintrags mit Carriage Return
                  old_entry = sprachcollection(complete_entry) + Chr(13)
                Else: old_entry = sprachcollection(complete_entry)
                End If
                
                language_entry = old_entry + single_line + Chr(13)
                sprachcollection.Remove complete_entry ' Damit kein Doppeleintrag vorliegt: Eintrag löschen und dann..
                sprachcollection.Add language_entry, complete_entry 'Sprachkürzel zu der Sprachensammlung LanguageCollection   hinzufügen
                first_of_more = False 'Falls es mehr als eine Zeile gibt
'            End If
      End If 'Abfrage von "@" am Zeilenanfang
NoValidLine:
    Loop
    '(Hier sind bereits alle Einträge in der sprachcollection vorgenommen)
    





    first_entry = CLng(Mid(Complete_TextNr_Collection(1), 1, Len(Complete_TextNr_Collection(1)) - 2))    'Die Nummer des ersten Eintrags suchen

      'Falls die behandelte Sprache keinen ersten Eintrag in der Sprachdatei hat
      'lösche alle Nummer mit dieser Sprache aus der Liste
    firstEntryInErrorFile = True
    For Each Item2 In LanguageCollection
      If Not ReadCollectionItem(first_entry & Item2, language_entry) Then
              errorfilename = App.Path & "\SpracheError.log"
              errorfile = FreeFile
              ' MS 20012006 V1.0.1
              If firstEntryInErrorFile Then
                Open errorfilename For Output As errorfile
                firstEntryInErrorFile = False
              Else
                Open errorfilename For Append As errorfile
              End If
                  Print #errorfile, "Warning: Not supported Language:" & Item2
              Close errorfile
        For Each entry In Complete_TextNr_Collection
            If Mid(entry, Len(entry) - 1, 2) = Item2 Then
              Complete_TextNr_Collection.Remove (entry)
              sprachcollection.Remove (entry) 'Dieser Eintrag löscht alle nicht gepflegten Sprachen aus der
              'sprachcollection. Der Eintrag kann bei sehr großen Collections hilfreich sein um
              'Speicher zu sparen, allerdings ist der Aufwand beim ersten Laden größer.
            End If '
        Next entry
        LanguageCollection.Remove (Item2)
      End If ' Not ReadCollectionItem (Falls Eintrag nicht existiert)
    Next Item2



    'Alle erkannten Sprachen mit allen erkannten Textnummern durchlaufen lassen
    'Textnummern zu denen es keinen Eintrag gibt in Log File speichern
    For Each Item2 In LanguageCollection
      For Each entry In Complete_TextNr_Collection
          textnumber = CLng(Mid(entry, 1, Len(entry) - 2))
          If Not ReadCollectionItem(textnumber & Item2, language_entry) Then
              For Each entry2 In LanguageCollection
                If ReadCollectionItem(textnumber & entry2, language_entry) Then
                  ErrorCollection.Add "Error: @" & textnumber & " is available for " & entry2 & " but not for " & Item2
                End If
              Next entry2
          End If ' Not ReadCollectionItem (Falls Eintrag nicht existiert)
      Next entry
    Next Item2
      
      
    'Ist die angeforderte Sprachdatei (aus der INI Datei) überhaupt verfügbar?
    For Each Item2 In LanguageCollection
        If SprachFileSprache = Item2 Then
          language_supported = True
        Else: avail_language = Item2 & "," & avail_language 'Verfügbare Sprachen
        End If
    Next Item2
    If Not language_supported Then ErrorCollection.Add "Language " & SprachFileSprache & "  not supported. Please adjust the .INI file!!"
    
    'Hier alle Fehler gesammelt in das Logfile schreiben
    If ErrorCollection.Count > 0 Then
      On Error GoTo openErr
      errorfilename = App.Path & "\SpracheError.log"
      errorfile = FreeFile
        Open errorfilename For Output As errorfile
          For Each Item2 In ErrorCollection
            Print #errorfile, Item2
          Next Item2
        Close errorfile
      
      'MsgBox ("There were errors. Please refer to file " & Chr(13) & App.Path & "\SpracheError.log" & Chr(13) & "for details")
'      logger.errorLine "Utilities::t() - Language File Error(s) - refer to " & Chr(13) & App.Path & "\SpracheError.log" & " for details"
      
      
      firsterror = False
    Else: firsterror = True
    End If



  End If 'Not init Then    Erster Aufruf von t (Der Quellcode bis hierher wird nur ein einziges Mal durchlaufen)
  
  
  
  
  
  If language_supported Then
    If ReadCollectionItem(CStr(Index) & SprachFileSprache, returnval) Then
      If SprachFileNummernAnzeigen Then
        t_variable_texte = CStr(Index) + "," + returnval
      Else
         t_variable_texte = returnval
      End If
    Else: t_variable_texte = DefaultText
        ErrorCollection.Add CStr(Index) & SprachFileSprache & " is not available ", CStr(Index) & SprachFileSprache
        sprachcollection.Add DefaultText, CStr(Index) & SprachFileSprache
        
        If firsterror Then
          MsgBox ("There were errors. Please refer to file " & Chr(13) & App.Path & "\SpracheError.log" & Chr(13) & "for details")
          firsterror = False
        End If
        
        On Error GoTo openErr
          errorfilename = App.Path & "\SpracheError.log"
          errorfile = FreeFile
          Open errorfilename For Append As errorfile
              Print #errorfile, "@" & CStr(Index) & SprachFileSprache & vbTab & " was requested, but is not available!"
          Close errorfile
    End If
  Else:
      t_variable_texte = DefaultText ' " Language " & SprachFileSprache & " not supported. Please adjust the .INI file!! Available languages are " & avail_language
  End If
  Exit Function
    

openErr:
        MsgBox (App.Path & SpracheFileName & " could not be opened") 'dateinamen !!!!
  Exit Function
NoLanguageAbbrev:
        MsgBox ("Entrys in file " + App.Path & SpracheFileName + " do not have country abbreviations (for example @500 instead of @500EN). ")
  Exit Function
End Function

Private Function ReadCollectionItem(varKey, var_text) As Boolean

  On Error GoTo not_found
    var_text = CStr(sprachcollection.Item(varKey))
    ReadCollectionItem = True
  Exit Function

not_found:
  ReadCollectionItem = False
End Function

Sub sAdd_Language()
  On Error Resume Next
    LanguageCollection.Add country_abbrev, country_abbrev

End Sub

Sub sAdd_Key()
  On Error GoTo Doppelter_Eintrag
    Complete_TextNr_Collection.Add complete_entry, complete_entry
  Exit Sub

Doppelter_Eintrag:
  ErrorCollection.Add "Error: @" & complete_entry & " already exists for this language"
End Sub



' PM 12082004 V2.0.73 (
Public Function MWFileExist(Filename As String) As Boolean
' Return: TRUE=File vorhanden und File ist größer 0
  Dim fNr As Integer
  
  MWFileExist = False
  If Len(Dir(Filename)) < 1 Then Exit Function
  On Error Resume Next
  If FileLen(Filename) < 1 Then Exit Function
  fNr = FreeFile
  On Error GoTo openErr:
  Open Filename For Input As fNr
  MWFileExist = True
openErr:
  Close fNr
End Function
' PM 12082004 V2.0.73 )

Public Sub SortString(DatenFeld() As String)
  Dim P1 As Long
  Dim P2 As Long
  Dim Temp As Variant

  If IsArray(DatenFeld) Then
    For P1 = 0 To UBound(DatenFeld)
      For P2 = P1 To UBound(DatenFeld)
        If UCase(DatenFeld(P2)) < UCase(DatenFeld(P1)) Then
          Temp = DatenFeld(P1)
          DatenFeld(P1) = DatenFeld(P2)
          DatenFeld(P2) = Temp
        End If
      Next P2
    Next P1
  End If
End Sub

' Zerlegt eine 32-Bit Zahl in 4 Byte(0..3)
Public Sub LongToByte(ByVal l As Long, B() As Byte)
Dim i As Integer

    For i = 0 To 3
        B(i) = l And 255&
        l = Int(l / 256&)
'         L = L \ 256&           ' funktioniert nicht !!
   Next i

End Sub

' Vereint 4 Bytes(0..3) zu einer 32-Bit Zahl mit Vorzeichen
Public Function ByteToLong(B() As Byte) As Long
Dim i As Integer
Dim negativ As Boolean
Dim ByteLocal(3) As Byte

    ByteLocal(0) = B(0):    ByteLocal(1) = B(1)
    ByteLocal(2) = B(2):    ByteLocal(3) = B(3)

    negativ = False
    If (ByteLocal(3) And &H80) Then
        For i = 0 To 3
            ByteLocal(i) = Not (ByteLocal(i))
            negativ = True
        Next i
    End If
    
    For i = 3 To 0 Step -1
        ByteToLong = (256 * ByteToLong) + ByteLocal(i)
    Next i
    
    If negativ Then
        ByteToLong = Not (ByteToLong)
    End If
    
End Function


Public Sub WriteComm(Txt As String, Mode As ProtokollMode)
    Dim CommFile As Integer
    Dim i As Integer
    Dim Prexit As String
    
    CommFile = FreeFile                'Nächste freie DateiNr.
    On Error GoTo OpenError
    Open CommFileName For Append As CommFile
    
    Select Case Mode
      Case Send
        Print #CommFile, "--> Send:   " & Txt
       
      Case Receive
        Print #CommFile, "--> Recive:   " & Txt
    
    End Select
    
    Close CommFile
    
    
    Exit Sub

OpenError:
  MsgBox CommFileName, , "Write error"

End Sub
