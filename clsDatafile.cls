VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatafile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'-------------------------------------------------------------
'clsDatafile v1.6 - 26 May 2003
'This class contains routines for reading and writing to files
'Written by Paul Berlin 2002-2003
'-------------------------------------------------------------
'
'clsDatafile can write and read several types of data to/from a file.
'It CANNOT write single or double values.
'
'This is what it writes:
'* LONG NUMBER (5 bytes)
'  1. byte value 100 (used for identification when reading)
'  2. Long number compressed to 4 bytes
'* NEGATIVE LONG NUMBER (5 bytes)
'  1. byte value 101 (used for identification when reading)
'  2. Negative long number converted to positive compressed to 4 bytes
'* BYTE NUMBER (<=255) (2 bytes)
'  1. byte value 102 (used for identification when reading)
'  2. the byte
'* FIXED LENGTH STRING
'  This writes an string of known fixed length (for example an file header)
'  and nothing else. If used in the middle of an file you cannot use SkipField
'  to skip this field! Best used for headers.
'  1. the string
'* SHORT STRING (<=255 chars) (3-257 bytes)
'  1. byte value 103 (used for identification when reading)
'  2. length of string as an byte
'  3. the string itself
'* LONG STRING (>255 chars) (258+ bytes)
'  1. byte value 104 (used for identification when reading)
'  2. length of string as an compressed long into 4 bytes
'  3. the string itself
'* DATES & TIME (10 bytes)
'  It's not very compact, but I don't know how to edit single bits so I could
'  for example put Month, Day and Hour in one byte (the max for those is only 67).
'  1. byte value 105 (used for identification when reading)
'  2. Year as an 4 byte long
'  3. Month as 1 byte
'  4. Day as 1 byte
'  5. Hour as 1 byte
'  6. Minute as 1 byte
'  7. Second as 1 byte
'* EMBEDDED FILE (6+ bytes)
'  1. byte value 106 (used for identification when reading)
'  2. length of file as 4 byte long
'  3. the file itself
'* SKIPPED VALUE
'  Whenever an written value is 0 or written string has an length of 0,
'  to save space:
'  1. byte value 200 (used for identification when reading)

'FUNCTIONS & SUBROUTINES:
'[PROPERTIES]
'* Filename [String]
'  Assigns an filename to read/write from.
'  This is the first and only thing you have to set before starting to read/write.
'* Position [Long]
'  This is the current byte position in the file you are at.
'  You do not have to touch this property in any simple reading/writing tasks.
'  Perhaps most useful to get the final size of the file when finished writing.
'* XorCode [Boolean]
'  Turn on/off simple Xor coding of strings. Remember to turn on this when
'  reading from Xor coded files.
'[WRITING]
'* WriteStr [String]
'  This simply writes an string to the current position. The string can be any size
'  and the proper writing method (LONG STRING/SHORT STRING/SKIPPED) will be used
'  depending on the length.
'* WriteNumber [Number]
'  This simply writes the supplied number with the proper writing method
'  (POS LONG/NEG LONG/BYTE/SKIPPED) depending on the length of the number at
'  the current file position.
'* WriteStrFixed [String]
'  Writes an fixed length string at current position. This can only be read with
'  ReadStrFixed, and cannot be skipped with SkipField. This should generally
'  only be used for writing file headers.
'* WriteDate [Date]
'  Writes an date at current file position.
'* WriteFile [File]
'  Writes an entire file into the datafile at current position.
'* There are other writing subroutines, but only the above are needed.
'[READING]
'* ReadStr
'  This reads an string from the current position and returns it.
'  If the value at the current file position is not an string, an error will be
'  raised.
'* ReadNumber
'  This reads an number from the current position and returns it.
'  If the value at the current file position is not an number, an error will be
'  raised.
'* ReadStrFixed [Length]
'  Reads an fixed length string at current position.
'* ReadDate
'  This reads an date from the current position and returns it.
'  If the value at the current file position is not an date, an error will be
'  raised.
'* ReadFile [File]
'  Reads an file embedded into the datafile and saves it to [File].
'  If the value at the current file position is not an file, an error will be
'  raised.
'* There are other writing subroutines, but only the above are needed.
'[OTHER]
'* SkipFiled [NumberSkips]
'  This skips any number of fields, as long as they are recognised.

'CLOSING THE FILE:
'The datafile will be closed whenever you change the filename property or
'you unload the class.

'SIMPLE EXAMPLE:
'--------------
'
'Dim File as new clsDatafile
'
'File.Filename = App.Path & "\file.tmp"
'File.WriteStrFixed "DATA"      'Fixed str header
'File.WriteNumber 100           'Byte number
'File.WriteNumber 100000        'Long number
'File.WriteNumber -100000       'Neg. Long number
'File.WriteNumber 0             'Skipped number
'File.WriteStr "Hello!"         'Byte length string
'File.WriteStr String(500, "X") 'Long length string
'File.WriteStr ""               'Skipped str
'File.WriteDate Now             'Date
'File.WriteFile sFile           'File (must be set to an valid file)
'Set File = Nothing             'Close
'Set File = New clsDatafile
'File.Filename = App.Path & "\file.tmp"
'Debug.Print File.ReadStrFixed 4        'Read fixed str header
'Debug.Print File.ReadNumber            'Byte number
'Debug.Print File.ReadNumber            'Long number
'Debug.Print File.ReadNumber            'Neg. Long number
'Debug.Print File.ReadNumber            'Skipped number
'Debug.Print File.ReadStr               'Byte length string
'Debug.Print File.ReadStr               'Long length string
'Debug.Print File.ReadStr               'Skipped str
'Debug.Print File.ReadDate              'Date
'Debug.Print File.ReadFile sFile        'File (must be set to an valid file)
'

Option Explicit

Const MaxBuffer As Long = 1048576 'The MAX number of bytes WriteFile can buffer

Private DoXor         As Boolean
Private strFilename   As String 'Filename
Private lngPosition   As Long   'File Position
Private Free          As Integer 'File number
Private IsOpen        As Boolean 'True if file is open, false if not

Public Function ReadStr() As String
  'This function reads an string from the current file
  'position and returns it. It reads both short(max 255
  'chars) and long strings.
  '---------------------------------------------------
  Dim vSize As Byte   'The size of the short string
  Dim vByte(3) As Byte 'The size of the long string
  
  If Not IsOpen Then OpenFile
  
  Get Free, lngPosition, vSize: lngPosition = lngPosition + 1
  If vSize = 103 Then 'short string (=<255 char)
  
    Get Free, lngPosition, vSize: lngPosition = lngPosition + 1 'Get size
    ReadStr = Space(vSize) 'Setup the size of the string to read into
    Get Free, lngPosition, ReadStr: lngPosition = lngPosition + vSize

  ElseIf vSize = 104 Then 'long string (>255 char)
  
    Get Free, lngPosition, vByte: lngPosition = lngPosition + 4 'get size
    ReadStr = Space(Byte2Long(vByte)) 'Setup the size of the string to read into
    Get Free, lngPosition, ReadStr: lngPosition = lngPosition + Byte2Long(vByte)
    
  ElseIf vSize = 200 Then 'The string is 0 length, skip
    ReadStr = ""
  
  Else 'This ain't no string
    Debug.Print "unknown string"
    Err.Raise 566
  End If
  
  If DoXor Then ReadStr = XorString(ReadStr)
  
End Function

Public Function ReadStrFixed(ByVal vSize As Long) As String
  'This function reads an fixed size string from the
  'current file position and returns it. vSize is the
  'size of the string to read.
  '---------------------------------------------------
  
  If Not IsOpen Then OpenFile
  
  ReadStrFixed = Space(vSize) 'Setup the size of the string to read into
  Get Free, lngPosition, ReadStrFixed: lngPosition = lngPosition + vSize
  
End Function

Public Sub WriteNumber(ByVal wNumber As Long)
  'This sub will examine the length of wNumber and use the
  'appropriate method to write it.
  '------------------------------------------------------
  
  If wNumber > 255 Or wNumber < 0 Then
    WriteLong wNumber
  Else
    WriteByte wNumber
  End If
  
End Sub

Public Sub WriteStr(ByVal wString As String)
  'This sub will examine the length of wString and use the
  'appropriate method to write it.
  '------------------------------------------------------

  If Len(wString) > 255 Then
    WriteLongStr wString
  Else
    WriteShortStr wString
  End If

End Sub

Public Sub WriteShortStr(ByVal wString As String)
  'This sub writes an string to the current file position
  '------------------------------------------------------
  
  If Not IsOpen Then OpenFile
  
  If Len(wString) = 0 Then
    Put Free, lngPosition, 200: lngPosition = lngPosition + 1
    Exit Sub
  End If
  
  'Crop string if it is longer than 255 characters
  If Len(wString) > 255 Then wString = Left(wString, 255)
  
  If DoXor Then wString = XorString(wString)
  
  Put Free, lngPosition, 103: lngPosition = lngPosition + 1
  Put Free, lngPosition, CByte(Len(wString)): lngPosition = lngPosition + 1
  Put Free, lngPosition, wString: lngPosition = lngPosition + Len(wString)

End Sub

Public Sub WriteLongStr(ByVal wString As String)
  'This sub writes an string to the current file position
  'The string can be over 255 characters long.
  '------------------------------------------------------
  
  If Not IsOpen Then OpenFile
  
  If Len(wString) = 0 Then
    Put Free, lngPosition, 200: lngPosition = lngPosition + 1
    Exit Sub
  End If
  
  If DoXor Then wString = XorString(wString)
  
  Put Free, lngPosition, 104: lngPosition = lngPosition + 1
  Put Free, lngPosition, Long2ByteStr(Len(wString)): lngPosition = lngPosition + 4
  Put Free, lngPosition, wString: lngPosition = lngPosition + Len(wString)

End Sub

Public Sub WriteStrFixed(ByVal wString As String)
  'This sub writes an fixed size string to the current
  'file position. The length can be as long as you wish.
  'No Identification number will be written when writing
  'this string, so you should only use this to write file
  'headers.
  '------------------------------------------------------
  
  If Not IsOpen Then OpenFile
  
  Put Free, lngPosition, wString: lngPosition = lngPosition + Len(wString)

End Sub

Public Sub WriteByte(ByVal wByte As Byte)
  'This sub writes an byte to the current file position
  '----------------------------------------------------
  
  If wByte = 0 Then
    Put Free, lngPosition, 200: lngPosition = lngPosition + 1
    Exit Sub
  End If
  
  If Not IsOpen Then OpenFile
  
  Put Free, lngPosition, 102: lngPosition = lngPosition + 1
  Put Free, lngPosition, wByte: lngPosition = lngPosition + 1

End Sub

Public Sub WriteLong(ByVal wLong As Long)
  'This sub writes an long to the current file position
  '----------------------------------------------------
  
  If Not IsOpen Then OpenFile
  
  If wLong = 0 Then
    Put Free, lngPosition, 200: lngPosition = lngPosition + 1
    Exit Sub
  End If
  
  If wLong < 0 Then
    Put Free, lngPosition, CByte(101)
    lngPosition = lngPosition + 1
    Put Free, lngPosition, Long2ByteStr(wLong * -1) 'convert to a positive long
    lngPosition = lngPosition + 4
  Else
    Put Free, lngPosition, CByte(100)
    lngPosition = lngPosition + 1
    Put Free, lngPosition, Long2ByteStr(wLong)
    lngPosition = lngPosition + 4
  End If
  
End Sub

Public Property Get Position() As Long
  Position = lngPosition
End Property

Public Property Let Position(ByVal lPos As Long)
  lngPosition = lPos
  If lngPosition < 1 Then lngPosition = 1 'The position cannot be lower than 1
End Property

Public Property Let FileName(ByVal sFilename As String)
  If IsOpen Then
    Close Free
    lngPosition = 1
    IsOpen = False
  End If
  strFilename = sFilename
End Property

Public Property Get FileName() As String
  FileName = strFilename
End Property

Private Function Long2ByteStr(ByVal Val As Long) As String
  'This sub converts an Long to 4 bytes, returned as an string
  Dim x As Integer

  ReDim byteArray(3) As Byte
  For x = 0 To 3
    byteArray(x) = Val \ (2 ^ ((3 - x) * 7)) And 127
  Next x
  
  Long2ByteStr = Chr(byteArray(0)) & Chr(byteArray(1)) & Chr(byteArray(2)) & Chr(byteArray(3))

End Function

Private Function Byte2Long(ByRef byteArray() As Byte) As Long
  'This function converts 4 bytes in byteArray() to an long and retuns it.
  Dim x As Integer

  Byte2Long = 0
  For x = 0 To 3
    Byte2Long = Byte2Long + (byteArray(x) * (2 ^ ((3 - x) * 7)))
  Next x

End Function

Private Sub Class_Initialize()
  'Setup file position
  lngPosition = 1
End Sub

Private Sub OpenFile()
  'This initializes the file
  Free = FreeFile
  Open strFilename For Binary As Free
  IsOpen = True
End Sub

Private Sub Class_Terminate()
  'This closes the file
  Close Free
End Sub

Public Sub WriteFile(ByVal sFilename As String)
  'This sub writes an other file into the file
  '-------------------------------------------
  
  Dim sBuffer As String
  Dim lCurLoc As Long
  Dim lFile As Long
  Dim lLen As Long
   
  If Not IsOpen Then OpenFile
  
  lLen = FileLen(sFilename)
  
  Put Free, lngPosition, 106: lngPosition = lngPosition + 1
  Put Free, lngPosition, Long2ByteStr(lLen)
  lngPosition = lngPosition + 4
  
  lFile = FreeFile
  lCurLoc = 1
  Open sFilename For Binary As lFile
  Do
    If lLen - lCurLoc < MaxBuffer Then 'figure out the size of the buffer
      sBuffer = String(lLen - lCurLoc + 1, " ")
    Else
      sBuffer = String(MaxBuffer, " ")
    End If
    Get lFile, lCurLoc, sBuffer 'get from file
    lCurLoc = lCurLoc + Len(sBuffer) 'adjust file position
    Put Free, lngPosition, sBuffer 'put in file
    lngPosition = lngPosition + Len(sBuffer) 'adjust file position
  Loop Until lCurLoc >= lLen

End Sub

Public Sub ReadFile(ByVal sFilename As String)
  'This sub reads and extracts an file and saves it to sFilename
  Dim sBuffer As String
  Dim lCurLoc As Long
  Dim lFile As Long
  Dim lLen As Long
  Dim vByte(3) As Byte
  
  If Not IsOpen Then OpenFile
  
  Get Free, lngPosition, vByte(0): lngPosition = lngPosition + 1 'The ID
  If vByte(0) <> 106 Then
    Err.Raise 566
    Exit Sub
  End If
  Get Free, lngPosition, vByte: lngPosition = lngPosition + 4
  lLen = lngPosition + Byte2Long(vByte())
  lFile = FreeFile
  lCurLoc = 1
  Open sFilename For Binary As lFile
  Do
    If lLen - lngPosition < MaxBuffer Then
      sBuffer = String(lLen - lngPosition, " ")
    Else
      sBuffer = String(MaxBuffer, " ")
    End If
    Get Free, lngPosition, sBuffer
    lngPosition = lngPosition + Len(sBuffer)
    Put lFile, lCurLoc, sBuffer
    lCurLoc = lCurLoc + Len(sBuffer)
  Loop Until lngPosition >= lLen
  
  Close lFile

  
End Sub

Public Function ReadNumber() As Long
  'This function reads an number from the current file
  'position and returns it. It reads both bytes and longs.
  '---------------------------------------------------
  Dim vType As Byte   'The type of number to read
  Dim vByte(3) As Byte '4 bytes to read longs into
  
  If Not IsOpen Then OpenFile
  
  Get Free, lngPosition, vType: lngPosition = lngPosition + 1
  If vType = 102 Then 'byte
  
    Get Free, lngPosition, vType: lngPosition = lngPosition + 1 'Get size
    ReadNumber = vType
    
  ElseIf vType = 101 Then 'negative long
  
    Get Free, lngPosition, vByte: lngPosition = lngPosition + 4
    ReadNumber = Byte2Long(vByte()) * -1

  ElseIf vType = 100 Then 'positive long
  
    Get Free, lngPosition, vByte: lngPosition = lngPosition + 4
    ReadNumber = Byte2Long(vByte())
    
  ElseIf vType = 200 Then 'The number is 0
    
    ReadNumber = 0
    
  Else 'This aint no number
    Debug.Print "unknown number"
    Err.Raise 566
  End If
  
End Function

Public Sub WriteDate(ByVal wDate As Date)
  'This sub writes an date to the current file position.
  '-----------------------------------------------------
  'What is writes:
  '1. One byte with value 105 (identifying it as an date)
  '2. Year as long
  '3. Month as byte
  '4. Day as byte
  '5. Hour as byte
  '6. Minute as byte
  '7. Second as byte
  '(total 10 bytes)
  'I know it's not very compact, you do it better and send it to me!
  
  If Not IsOpen Then OpenFile
  
  Put Free, lngPosition, 105: lngPosition = lngPosition + 1
  Put Free, lngPosition, Long2ByteStr(Year(wDate)): lngPosition = lngPosition + 4
  Put Free, lngPosition, CByte(Month(wDate)): lngPosition = lngPosition + 1
  Put Free, lngPosition, CByte(Day(wDate)): lngPosition = lngPosition + 1
  Put Free, lngPosition, CByte(Hour(wDate)): lngPosition = lngPosition + 1
  Put Free, lngPosition, CByte(Minute(wDate)): lngPosition = lngPosition + 1
  Put Free, lngPosition, CByte(Second(wDate)): lngPosition = lngPosition + 1
  
End Sub

Public Function ReadDate() As Date
  'This function reads an date at the current file position
  'and returns it.
  '-----------------------------------------------------
  Dim vByte(3) As Byte 'Holds 4 bytes for year
  Dim vByte2(4) As Byte 'Holds the rest of the values
  
  If Not IsOpen Then OpenFile
  
  Get Free, lngPosition, vByte2(0): lngPosition = lngPosition + 1
  If vByte2(0) <> 105 Then
    Err.Raise 566
    Exit Function
  End If
  Get Free, lngPosition, vByte: lngPosition = lngPosition + 4
  Get Free, lngPosition, vByte2(0): lngPosition = lngPosition + 1
  Get Free, lngPosition, vByte2(1): lngPosition = lngPosition + 1
  Get Free, lngPosition, vByte2(2): lngPosition = lngPosition + 1
  Get Free, lngPosition, vByte2(3): lngPosition = lngPosition + 1
  Get Free, lngPosition, vByte2(4): lngPosition = lngPosition + 1
  ReadDate = Byte2Long(vByte()) & "-" & vByte2(0) & "-" & vByte2(1) & " " & vByte2(2) & ":" & vByte2(3) & ":" & vByte2(4)
  
End Function

Public Sub SkipField(Optional ByVal lSkips As Integer = 1)
  'This sub skips lSkips values, and moves the current file position
  'after them without reading them.
  '-----------------------------------------------------------------
  Dim x As Integer
  Dim vByte As Byte
  Dim vByte2 As Byte
  Dim vBytes(3) As Byte
  
  For x = 1 To lSkips
    Get Free, lngPosition, vByte: lngPosition = lngPosition + 1 'Check which value this is
    Select Case vByte
      Case 100, 101 'Longs
        lngPosition = lngPosition + 4
      Case 102 'Byte
        lngPosition = lngPosition + 1
      Case 103 'Short string
        Get Free, lngPosition, vByte2
        lngPosition = lngPosition + vByte2 + 1
      Case 104, 106  'Long string & file
        Get Free, lngPosition, vBytes
        lngPosition = lngPosition + Byte2Long(vBytes()) + 4
      Case 105 'Date
        lngPosition = lngPosition + 9
      Case 200
        'Do nothing
      Case Else 'Could not identify an header, position might be off...
        Err.Raise 566
    End Select
  Next
  
End Sub

Private Function XorString(ByVal Text As String) As String
  Dim A As Integer
  For A = 1 To Len(Text) 'code all
   XorString = XorString & Chr(Asc(Mid(Text, A, 1)) Xor 1) 'Code text
  Next A
End Function

Public Property Get XorCode() As Boolean
  XorCode = DoXor
End Property

Public Property Let XorCode(ByVal bXor As Boolean)
  DoXor = bXor
End Property
