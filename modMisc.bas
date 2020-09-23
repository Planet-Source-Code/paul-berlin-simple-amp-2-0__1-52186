Attribute VB_Name = "modMisc"
'This module contains misc public functions and subs

Option Explicit

'APIs
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hDCSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Enum enumKeyPressAllowTypes
  NumbersOnly = 2 ^ 0
  Uppercase = 2 ^ 1
  NoSpaces = 2 ^ 2
  NoSingleQuotes = 2 ^ 3
  NoDoubleQuotes = 2 ^ 4
  AllowDecimal = 2 ^ 5
  AllowNegative = 2 ^ 6
  DatesOnly = 2 ^ 7
  TimesOnly = 2 ^ 8
  LettersOnly = 2 ^ 9
  AllowSpaces = 2 ^ 10
  AllowStars = 2 ^ 11
  AllowPounds = 2 ^ 12
End Enum

Public Enum enumFileNameParts
  efpFileName = 2 ^ 0
  efpFileExt = 2 ^ 1
  efpFilePath = 2 ^ 2
  efpFileNameAndExt = efpFileName + efpFileExt
  efpFileNameAndPath = efpFilePath + efpFileName
End Enum

'SetWindowPos constants
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
'Used for setting window style
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2

Public Sub MakeTransparent(ByVal hwnd As Long, ByVal Rate As Byte)
  'Makes window transparent 0-254
  'Requires Win 2k/XP
  On Error Resume Next
  
  If Rate = 255 Then 'turn off layered style
    SetWindowLong hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Xor WS_EX_LAYERED
  Else
    SetWindowLong hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes hwnd, 0, Rate, LWA_ALPHA
  End If
  
End Sub

Public Function ConvertTime(ByVal Sec As Long, Optional ByVal bQuestion As Boolean) As String
  'Converts seconds to the format 00:00:00/00:00 as string
  On Error Resume Next
  Dim Minutes As Long, strMinutes As String
  Dim Seconds As Long, strSeconds As String, Hours As Long
  
  If Sec < 0 And bQuestion Then ConvertTime = "??:??": Exit Function
  
  Seconds = Sec
  
  Minutes = Seconds \ 60
  Seconds = Seconds - (Minutes * 60)
  Hours = Minutes \ 60
  Minutes = Minutes - (Hours * 60)
  
  If Seconds < 10 Then strSeconds = "0" & Seconds Else strSeconds = Seconds
  If Minutes < 10 Then strMinutes = "0" & Minutes Else strMinutes = Minutes
    
  If Hours > 0 Then
    ConvertTime = Hours & ":" & strMinutes & ":" & strSeconds
  Else
    ConvertTime = strMinutes & ":" & strSeconds
  End If

End Function

Public Function FileExists(ByVal FileName As String) As Boolean
  'Cheks if file exists
  On Error Resume Next
  FileExists = Not (Dir(FileName) = "")
End Function

Public Function SavePlaylist(ByVal fname As String) As Boolean
  'Saves playlist to file FName, returns true if successful
  Dim X As Long, y As Long
  Dim File As New clsDatafile
  Dim tDir As String, tmpName As String
  
  On Error GoTo ErrHandler

  If FileExists(fname) Then Kill fname
  SavePlaylist = True
  
  File.FileName = fname
  
  'figure out the Common dir for all entries
  '(to make file size smaller)
  For X = 1 To UBound(Playlist)
    If Len(tDir) = 0 Then
      tDir = Library(Playlist(X).Reference).sFilename
    Else
      For y = 1 To Len(tDir)
        If Left(tDir, y) = Left(Library(Playlist(X).Reference).sFilename, y) Then
          tmpName = Left(tDir, y)
        Else
          tDir = tmpName
        End If
      Next y
    End If
  Next X
  
  'Write header
  File.WriteStrFixed "SAMPEXT"
  'Write total number of items in list
  File.WriteNumber UBound(Playlist)
  'Write common dir
  File.WriteStr tDir
  For X = 1 To UBound(Playlist)
    If Not Playlist(X).Removed Then
      'Write filename & reduce its size with common dir if possible
      If Left(Library(Playlist(X).Reference).sFilename, Len(tDir)) = tDir <> 0 Then
        File.WriteStr CStr(Right(Library(Playlist(X).Reference).sFilename, Len(Library(Playlist(X).Reference).sFilename) - Len(tDir)))
      Else
        File.WriteStr Library(Playlist(X).Reference).sFilename
      End If
      'write index
      File.WriteNumber Playlist(X).Index
      'write Artist & Title
      File.WriteStr Library(Playlist(X).Reference).sArtistTitle
      'write length
      File.WriteNumber Library(Playlist(X).Reference).lLength
    End If
  Next X
  
  Exit Function
ErrHandler:
  SavePlaylist = False
End Function

Public Function LoadPlaylist(ByVal fname As String) As Boolean
  'Loads playlist FName, returns true if successful
  Dim X As Long, tDir As String
  Dim File As New clsDatafile
  Dim nums As Long
  Dim tStr As String, b As Long
  
  On Error GoTo ErrHandler
  frmPlaylist.MousePointer = vbHourglass
  LoadPlaylist = True
  
  File.FileName = fname
  
  'Reads header & goes to error if it is not 'SAMPEXT'
  If File.ReadStrFixed(7) <> "SAMPEXT" Then GoTo ErrHandler
  nums = File.ReadNumber
  'Reads common dir
  tDir = File.ReadStr
  
  'Starts getting info for each item
  For X = 1 To nums
    
    'Start with filename
    tStr = File.ReadStr
    If Mid(tStr, 2, 1) <> ":" And Left(tStr, 2) <> "//" And Left(tStr, 2) <> "\\" Then
      tStr = tDir & tStr
    End If
      
      b = LibraryCheck(tStr)
      If b > 0 Then
        CreateLibrary tStr
        File.SkipField 3
      Else
        If Left(tStr, 4) = "//CD" Then 'If it is cd audio tracks
          ReDim Preserve Library(UBound(Library) + 1)
          ReDim Preserve Playlist(UBound(Playlist) + 1)
          Library(UBound(Library)).sFilename = tStr
          Library(UBound(Library)).eType = TYPE_CDA
          Playlist(UBound(Playlist)).Index = File.ReadNumber
          Playlist(UBound(Playlist)).Reference = UBound(Library)
          Library(UBound(Library)).sArtistTitle = File.ReadStr
          Library(UBound(Library)).lLength = File.ReadNumber
        Else 'other files
          ReDim Preserve LibraryIndex(UBound(LibraryIndex) + 1)
          ReDim Preserve Library(UBound(Library) + 1)
          ReDim Preserve Playlist(UBound(Playlist) + 1)
          LibraryIndex(UBound(LibraryIndex)).sFilename = tStr
          LibraryIndex(UBound(LibraryIndex)).lPointer = 0
          LibraryIndex(UBound(LibraryIndex)).lLen = Len(tStr)
          LibraryIndex(UBound(LibraryIndex)).lReference = UBound(Library)
          Library(UBound(Library)).sFilename = tStr
          'read index
          Playlist(UBound(Playlist)).Index = File.ReadNumber
          Playlist(UBound(Playlist)).Reference = UBound(Library)
          'Read artist & title
          Library(UBound(Library)).sArtistTitle = File.ReadStr
          'read length
          Library(UBound(Library)).lLength = File.ReadNumber
        End If
      End If
      
    frmPlaylist.lblTotalNum = X & " files."
    DoEvents
  Next X

  frmPlaylist.MousePointer = vbDefault
  UpdateList
  
  Exit Function
ErrHandler:
  LoadPlaylist = False
End Function

Public Function GetMp3VBR(ByVal sFilename As String) As Boolean
  'I grabbed this code from an mp3 header reading program and
  'adjusted it to only get if an mp3 is VBR.
  'It seems odd to me that VBRs contains "Xing" (As in the Xing En/De-coder)
  'in the header, there is no mention of this in the mp3 header specs on
  'www.mp3-tech.org (or .com was it?), but every VBR mp3 i have contains
  '"Xing" so it should work. /Paul
  'If anybody knows how to get this in any other, simpler way, please tell me!
  On Error GoTo ErrHand
  
  Dim XingH As String * 4
  Dim FIO As Long
  Dim i As Long
  Dim X As Byte
  Dim HeadStart As Long
  Dim MaxSize As Long
  Dim Max As Long
         
  FIO = FreeFile
    
  'read the header
  Open sFilename For Binary Access Read As FIO
  If LOF(FIO) < 256 Then GoTo ErrHand
  If LOF(FIO) < Max Then
    Max = LOF(FIO)
  Else
    Max = 10240
  End If
  
  For i = 1 To Max
    Get #FIO, i, X
    If X = 255 Then
      Get #FIO, i + 1, X
      If X > 249 And X < 252 Then
        HeadStart = i
        Exit For
      End If
    End If
  Next i
        
  'no header start position was found
  If HeadStart = 0 Then GoTo ErrHand
    
  'start check for XingHeader
  Get #FIO, HeadStart + 36, XingH
  If XingH = "Xing" Then GetMp3VBR = True

  Close FIO

  Exit Function
ErrHand:
  GetMp3VBR = False
  Close FIO
End Function

Public Function LoadPls(ByVal sFilename As String) As Boolean
  'Loads an playlist in format pls
  'pls is built as an ini-file
  On Error GoTo ErrHandler
  Dim X As Long, tStr As String
  Dim cINI As New clsINI
  LoadPls = True
  
  cINI.sFilename = sFilename
  cINI.sSection = "Playlist"
  
  If cINI.ReadNumber("Version", 2) <> 2 Then GoTo ErrHandler
  
  For X = 1 To cINI.ReadNumber("NumberOfEntries")
    tStr = cINI.ReadString("File" & X)
    If Not FileExists(tStr) Then 'check if file can be found
      If FileExists(modMisc.sFilename(sFilename, efpFilePath) & tStr) Then 'check if saved in reference
        tStr = modMisc.sFilename(sFilename, efpFilePath) & tStr
      End If
    End If
    If LibraryCheck(tStr) > 0 Then
      CreateLibrary tStr
    Else
      ReDim Preserve LibraryIndex(UBound(LibraryIndex) + 1)
      ReDim Preserve Library(UBound(Library) + 1)
      ReDim Preserve Playlist(UBound(Playlist) + 1)
      LibraryIndex(UBound(LibraryIndex)).sFilename = tStr
      LibraryIndex(UBound(LibraryIndex)).lPointer = 0
      LibraryIndex(UBound(LibraryIndex)).lLen = Len(tStr)
      LibraryIndex(UBound(LibraryIndex)).lReference = UBound(Library)
      Library(UBound(Library)).sFilename = tStr
      'read index
      Playlist(UBound(Playlist)).Index = UBound(Playlist)
      Playlist(UBound(Playlist)).Reference = UBound(Library)
      'Read artist & title
      Library(UBound(Library)).sArtistTitle = cINI.ReadString("Title" & X)
      Library(UBound(Library)).lLength = cINI.ReadNumber("Length" & X)
    End If
    DoEvents
  Next X
  
  UpdateList
    
  Exit Function
ErrHandler:
  LoadPls = False
End Function

Public Function LoadM3u(ByVal sFilename As String) As Boolean
  'This loads an m3u, both normal and extended versions are supported
  On Error GoTo ErrHandler
  Dim FF As Long, rStr As String, rStr2 As String
  Dim ll As Long, sAt As String
  
  LoadM3u = True
  
  FF = FreeFile
  Open sFilename For Input As FF
  Line Input #FF, rStr
  
  If InStr(1, rStr, Chr(10)) > 0 Then
    Close FF
    CleanFile sFilename
    Exit Function
  End If
  
  If rStr = "#EXTM3U" Then
    'Extended m3u
    
    Do Until EOF(FF)
      
        'First, get song data (length & title)
        Line Input #FF, rStr
        ll = Val(Mid(rStr, 9, InStr(1, rStr, ",") - 9))
        sAt = Right(rStr, Len(rStr) - InStr(1, rStr, ","))
      
        'Then, get filename
        Line Input #FF, rStr
        If Left(rStr, 1) = "\" Then
          If FileExists(Left(sFilename, InStrRev(sFilename, "\") - 1) & rStr) Then
            rStr2 = Left(sFilename, InStrRev(sFilename, "\") - 1) & rStr
          ElseIf FileExists(Left(sFilename, 2) & rStr) Then
            rStr2 = Left(sFilename, 2) & rStr
          Else
            rStr2 = rStr
          End If
        Else
          If FileExists(Left(sFilename, InStrRev(sFilename, "\")) & rStr) Then
            rStr2 = Left(sFilename, InStrRev(sFilename, "\")) & rStr
          ElseIf FileExists(Left(sFilename, 3) & rStr) Then
            rStr2 = Left(sFilename, 3) & rStr
          Else
            rStr2 = rStr
          End If
        End If
        DoEvents
        
          If LibraryCheck(rStr2) > 0 Then
            CreateLibrary rStr2
          Else
            ReDim Preserve LibraryIndex(UBound(LibraryIndex) + 1)
            ReDim Preserve Library(UBound(Library) + 1)
            ReDim Preserve Playlist(UBound(Playlist) + 1)
            LibraryIndex(UBound(LibraryIndex)).sFilename = rStr2
            LibraryIndex(UBound(LibraryIndex)).lLen = Len(rStr2)
            LibraryIndex(UBound(LibraryIndex)).lPointer = 0
            LibraryIndex(UBound(LibraryIndex)).lReference = UBound(Library)
            Library(UBound(Library)).sFilename = rStr2
            'read index
            Playlist(UBound(Playlist)).Index = UBound(Playlist)
            Playlist(UBound(Playlist)).Reference = UBound(Library)
            'Read artist & title
            Library(UBound(Library)).sArtistTitle = sAt
            'read length
            Library(UBound(Library)).lLength = ll
          End If
        
      
    Loop
    
  Else
    'Normal version
    
    Close FF
    Open sFilename For Input As FF
    
    Do Until EOF(FF)
      
        'Get filename
        Line Input #FF, rStr
        If InStr(1, rStr, ":") <> 0 Then
          rStr2 = rStr
        Else
          If FileExists(Left(sFilename, InStrRev(sFilename, "\")) & rStr) Then
            rStr2 = Left(sFilename, InStrRev(sFilename, "\")) & rStr
          Else
            rStr2 = rStr
          End If
        End If
        If LibraryCheck(rStr2) > 0 Then
          CreateLibrary rStr2
        Else
          ReDim Preserve LibraryIndex(UBound(LibraryIndex) + 1)
          ReDim Preserve Library(UBound(Library) + 1)
          ReDim Preserve Playlist(UBound(Playlist) + 1)
          LibraryIndex(UBound(LibraryIndex)).sFilename = rStr2
          LibraryIndex(UBound(LibraryIndex)).lLen = Len(rStr2)
          LibraryIndex(UBound(LibraryIndex)).lReference = UBound(Library)
          Library(UBound(Library)).sFilename = rStr2
          'read index
          Playlist(UBound(Playlist)).Index = UBound(Playlist)
          Playlist(UBound(Playlist)).Reference = UBound(Library)
          'Read artist & title
          Library(UBound(Library)).sArtistTitle = modMisc.sFilename(rStr2, efpFileNameAndExt)
        End If
        
    Loop
    
  End If
  
  UpdateList
  
  Close FF
  Exit Function
ErrHandler:
  LoadM3u = False
  Close FF
End Function

Public Sub CleanFile(ByVal sFilename As String)
  'This is an helper file for sub LoadM3u. It removes all chr(10) linefeeds and replaces them
  'with linebreaks (i think, chr(13)), and saves over the old file. Some normal version m3u's
  'saves with chr(10) and Line Input does not recognize them as linebreaks, and therefore reads
  'several lines in the first line input.
  On Error GoTo ErrHandler
  Dim FF As Long, rStr As String, i As Long
  Dim Spit() As String
  
  FF = FreeFile
  Open sFilename For Input As FF
  Line Input #FF, rStr
  Close FF
  
  Spit() = Split(rStr, Chr(10))
  
  Open sFilename For Output As FF
  For i = LBound(Spit) To UBound(Spit)
    If Len(Spit(i)) > 0 Then
      Print #FF, Spit(i)
    End If
  Next
  
  Close FF
  LoadM3u sFilename
  Exit Sub
  
ErrHandler:
  MsgBox "Could not read """ & sFilename & """.", vbExclamation
  Close FF
End Sub

Public Sub AlwaysOnTop(ByRef FormName As Form, ByVal OnTop As Boolean)
  'This sub sets FormName to always stay ontop without moving or resizing
  On Error Resume Next

  SetWindowPos FormName.hwnd, IIf(OnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Function GetPlaylistLength(ByVal sFilename As String)
  'This function get number of items in playlist file, for use in file browser window
  'works on .playlist, .m3u, .pls
  On Error Resume Next
  Select Case LCase(modMisc.sFilename(sFilename, efpFileExt))
  
    Case "playlist"
      Dim File As New clsDatafile
      File.FileName = sFilename
      If File.ReadStrFixed(7) = "SAMPEXT" Then
        'Reads total num items
        GetPlaylistLength = File.ReadNumber
      End If
      
    Case "m3u"
      Dim FF As Long, rStr As String
      FF = FreeFile
      Open sFilename For Input As FF
      Line Input #FF, rStr
  
      If InStr(1, rStr, Chr(10)) > 0 Then
        Close FF
        Exit Function
      End If
  
      If rStr = "#EXTM3U" Then
        Do
          Line Input #FF, rStr
          Line Input #FF, rStr
          GetPlaylistLength = GetPlaylistLength + 1
        Loop Until EOF(FF)
      Else
        GetPlaylistLength = 1
        Do
          Line Input #FF, rStr
          GetPlaylistLength = GetPlaylistLength + 1
        Loop Until EOF(FF)
      End If
      
      Close FF
  
    Case "pls"
      Dim cINI As clsINI
      cINI.sFilename = sFilename
      cINI.sSection = "Playlist"
      If cINI.ReadNumber("Version", 2) = 2 Then
        GetPlaylistLength = cINI.ReadNumber("NumberOfEntries")
      End If
      
  End Select
  
End Function

Public Function StrFormat(ByVal str As String) As String
  'This function formats a string to fix upper/lower case,
  '_, %20 to space and other stuff
  On Error Resume Next
  str = sNT(str)

  Do While InStr(1, str, "_") > 0
    str = Replace(str, "_", " ")
  Loop
  Do While InStr(1, str, "%20") > 0
    str = Replace(str, "%20", " ")
  Loop
  
  str = StrConv(str, vbProperCase)
  
  Dim X As Long
  X = 1
  Do
    Select Case Mid(str, X, 1)
      Case "[", "(", "{", ".", "/", "\", "-", Chr(34)
        str = Left(str, X) & UCase(Mid(str, X + 1, 1)) & Right(str, Len(str) - X - 1)
    End Select
    X = X + 1
  Loop Until X > Len(str) - 1

  StrFormat = str

End Function

Public Function InvertColor(ByVal lColor As Long) As Long
  'this inverts an RGB color
  InvertColor = RGB(255 - (lColor Mod &H100), 255 - ((lColor \ &H100) Mod &H100), 255 - ((lColor \ &H10000) Mod &H100))
End Function

Public Function ctlKeyPress(ByVal KeyAscii As KeyCodeConstants, ByVal TypeToAllow As enumKeyPressAllowTypes) As Integer
  'Written by the Frog Prince
    
  Dim ltrKeyAscii As Integer
  ltrKeyAscii = Asc(UCase(Chr(KeyAscii)))
    
  ' By default pass the keystroke through and then optionally kill it
  ctlKeyPress = KeyAscii
    
  ' Default Keystrokes to allow (enter, backspace, delete, escape)
  If _
    KeyAscii = vbKeyReturn Or _
    KeyAscii = vbKeyEscape Or _
    KeyAscii = vbKeyBack Or _
    KeyAscii = vbKeyDelete Then
      
    Exit Function
  End If
    
  ' NumbersOnly
  If (TypeToAllow And NumbersOnly) Then
    Select Case True
      Case KeyAscii >= vbKey0 And KeyAscii <= vbKey9
      Case (KeyAscii = vbKeySubtract Or KeyAscii = Asc("-")) And (TypeToAllow And AllowNegative)
      Case KeyAscii = Asc("#") And (TypeToAllow And AllowPounds)
      Case KeyAscii = Asc("*") And (TypeToAllow And AllowStars)
      Case KeyAscii = vbKeyDecimal And (TypeToAllow And AllowDecimal)
      Case KeyAscii = vbKeySpace And (TypeToAllow And AllowSpaces)
      Case Else
        KeyAscii = 0
    End Select
  End If
    
  ' DatesOnly
  If (TypeToAllow And DatesOnly) Then
    Select Case True
      Case KeyAscii >= vbKey0 And KeyAscii <= vbKey9
      Case KeyAscii = vbKeyDivide Or KeyAscii = Asc("/")
      Case Else
        KeyAscii = 0
    End Select
  End If
    
  ' TimesOnly
  If (TypeToAllow And TimesOnly) Then
    Select Case True
      Case KeyAscii >= vbKey0 And KeyAscii <= vbKey9
      Case KeyAscii = Asc(":") Or KeyAscii = Asc(";")
        ctlKeyPress = Asc(":")
      Case ltrKeyAscii = vbKeyA Or ltrKeyAscii = vbKeyP Or ltrKeyAscii = vbKeyM
      Case Else
        KeyAscii = 0
    End Select
  End If
            
  ' LettersOnly
  If (TypeToAllow And LettersOnly) Then
    Select Case True
      Case ltrKeyAscii >= vbKeyA And ltrKeyAscii <= vbKeyZ
      Case Else
        KeyAscii = 0
    End Select
  End If
            
  ' UpperCase
  If (TypeToAllow And Uppercase) Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
  End If
    
  ' No Spaces
  If (TypeToAllow And NoSpaces) And KeyAscii = vbKeySpace Then
    KeyAscii = 0
  End If
    
  ' No Double Quotes
  If (TypeToAllow And NoDoubleQuotes) And KeyAscii = Asc("""") Then
    KeyAscii = Asc("'")
  End If
    
  ' No Single Quotes
  If (TypeToAllow And NoSingleQuotes) And KeyAscii = Asc("'") Then
    KeyAscii = 0
  End If
    
  ctlKeyPress = KeyAscii
    
End Function

Public Function sNT(ByVal sStr As String) As String
  Dim iNL As Integer
  
  iNL = InStr(sStr, Chr(0))
  If iNL > 0 Then
    sNT = Left(sStr, iNL - 1)
  Else
    sNT = sStr
  End If
End Function

Public Function sFilename(ByVal sFile As String, ByVal ePortions As enumFileNameParts) As String
  'Written by the Frog Prince
  Dim lFirstPeriod As Long, lFirstBackSlash As Long
  Dim sName As String, sExt As String
  Dim sRet As String
  
  lFirstPeriod = InStrRev(sFile, ".")
  lFirstBackSlash = InStrRev(sFile, "\")
  
  If ePortions And efpFilePath Then
    If lFirstBackSlash > 0 Then
      sRet = Left(sFile, lFirstBackSlash)
    End If
  End If
  If ePortions And efpFileName Then
    If lFirstPeriod > 0 And lFirstPeriod > lFirstBackSlash Then
      sName = Mid(sFile, lFirstBackSlash + 1, lFirstPeriod - lFirstBackSlash - 1)
    Else
      sName = Mid(sFile, lFirstBackSlash + 1)
    End If
    sRet = sRet & sName
  End If
  If ePortions And efpFileExt Then
    If lFirstPeriod > 0 And lFirstPeriod > lFirstBackSlash Then
      sExt = Mid(sFile, lFirstPeriod + 1)
    End If
    If sRet <> "" Then
      sRet = sRet & "." & sExt
    Else
      sRet = sRet & sExt
    End If
  End If
  
  sFilename = sRet
    
End Function

Public Function sAppend(ByVal s2AppendTo As String, ByVal sChars2Append As String) As String
  If Right(s2AppendTo, Len(sChars2Append)) <> sChars2Append Then
    sAppend = s2AppendTo & sChars2Append
  Else
    sAppend = s2AppendTo
  End If
End Function

Public Function ctlSetFocus(ByRef ObjToSetFocusTo As Object) As Boolean
  On Error Resume Next
  ObjToSetFocusTo.SetFocus
  ctlSetFocus = Err.Number = 0
  On Error GoTo 0
End Function
