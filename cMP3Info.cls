VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMP3Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' AmBra (ambra@blic.net) 2001-06-25

' This Class module is built according to informations
' found on http://www.mp3-tech.org/. These informations
' can also be found on http://www.id3.org, but it seems
' that they slightly differ from those found on
' http://www.mp3-tech.org/. This class module reads first
' frame header of an MP3 file, and according to it, it
' collects informations about the file.

Option Explicit

Private m_sFileName As String
Private m_bValidHeader As Boolean
Private m_sID As String
Private m_sLayer As String
Private m_bProtectionBitSet As Boolean
Private m_sProtectionChecksum As String
Private m_iBitrate As Integer
Private m_lFrequency As Long
Private m_bPadded As Boolean
Private m_bPrivateBitSet As Boolean
Private m_sMode As String
Private m_iModeExt As Integer
Private m_bCopyrighted As Boolean
Private m_bOriginal As Boolean
Private m_sEmphasis As String

' DEFINING PROPERTIES

Public Property Let FileName(sInput As String)
    m_sFileName = sInput
End Property

Public Property Get FileName() As String
    FileName = m_sFileName
End Property

Public Property Get ValidHeader() As Boolean
    ValidHeader = m_bValidHeader
End Property

Public Property Get ID() As String
    ID = m_sID
End Property

Public Property Get Layer() As String
    Layer = m_sLayer
End Property

Public Property Get ProtectionBitSet() As Boolean
    ProtectionBitSet = m_bProtectionBitSet
End Property

Public Property Get ProtectionChecksum() As String
    ProtectionChecksum = m_sProtectionChecksum
End Property

Public Property Get Bitrate() As Integer
    Bitrate = m_iBitrate
End Property

Public Property Get Frequency() As Long
    Frequency = m_lFrequency
End Property

Public Property Get Padded() As Boolean
    Padded = m_bPadded
End Property

Public Property Get PrivateBitSet() As Boolean
    PrivateBitSet = m_bPrivateBitSet
End Property

Public Property Get Mode() As String
    Mode = m_sMode
End Property

Public Property Get ModeExt() As Integer
    ModeExt = m_iModeExt
End Property

Public Property Get Copyrighted() As Boolean
    Copyrighted = m_bCopyrighted
End Property

Public Property Get Original() As Boolean
    Original = m_bOriginal
End Property

Public Property Get Emphasis() As String
    Emphasis = m_sEmphasis
End Property

' READING HEADER

Public Sub ReadMP3Header()
On Error GoTo fault

Dim bTMP1 As Byte, bTMP2 As Byte
Dim filenum As Integer
filenum = FreeFile
Dim I As Long
Dim StartByte As Long
Open m_sFileName For Binary Access Read As #filenum
    
' FINDING FIRST FRAME HEADER

    For I = 1 To LOF(filenum) - 1
        Get #filenum, I, bTMP1
        If bTMP1 = &HFF Then
            Get #filenum, I + 1, bTMP2
            If bTMP2 And &HE0 = &HE0 Then
                m_bValidHeader = True
                StartByte = I + 1
                Exit For
            End If
        End If
    Next
    
    If Not m_bValidHeader Then Exit Sub
    
    Dim b1 As Byte, b2 As Byte, b3 As Byte, b4 As Byte, b5 As Byte
    
' GETTING BYTES THAT CONTAIN HEADER INFORMATIONS

    Get #filenum, StartByte, b1
    Get #filenum, StartByte + 1, b2
    Get #filenum, StartByte + 2, b3
    Get #filenum, StartByte + 3, b4
    Get #filenum, StartByte + 4, b5
    
Close #filenum

' READING MPEG MODE
Select Case CInt(b1 And &H18) / 8
    Case 0
        m_sID = "Mpeg 2.5"
    Case 1
        m_sID = "Not defined"
    Case 2
        m_sID = "Mpeg 2"
    Case 3
        m_sID = "Mpeg 1"
End Select
        
    
'READING LAYER INFO
Select Case (b1 And &H6)
    Case &H0
        m_sLayer = "Not defined"
    Case &H2
        m_sLayer = "Layer III"
    Case &H4
        m_sLayer = "Layer II"
    Case &H6
        m_sLayer = "Layer I"
End Select

' READING PROTECTION BIT, AND PROTECTION CHECKSUM IF THE BIT IS NOT SET
If (b1 And &H1) = &H1 Then
    m_bProtectionBitSet = True
Else
    m_bProtectionBitSet = False
    m_sProtectionChecksum = Hex(b4) & " " & Hex(b5)
End If

' READING BITRATE INFO
Dim arg1 As Integer, arg2 As Integer, arg3 As Integer
arg1 = CInt(b1 And &H8) / 8
arg2 = CInt(b1 And &H6) / 2
arg3 = CInt(b2 And &HF0) / 16
m_iBitrate = fnGetBitrate(arg1, arg2, arg3)

' READING FREQUENCY (SAMPLERATE)
Dim k As Long
Select Case CInt(b1 And &H18) / 8
    Case 0
        k = 1
    Case 1
        k = 0
    Case 2
        k = 2
    Case 3
        k = 4
End Select
Select Case CInt(b2 And &HC) / 4
    Case 0
        m_lFrequency = k * 11025
    Case 1
        m_lFrequency = k * 12000
    Case 2
        m_lFrequency = k * 8000
End Select

' READING PADDING BIT
If (b2 And &H2) = &H2 Then m_bPadded = True

' READING PRIVATE BIT
If (b2 And &H1) = &H1 Then m_bPrivateBitSet = True

' READING CHANNEL MODE
Select Case CInt(b3 And &HC0) / CInt(&H40)
    Case 0
        m_sMode = "Stereo"
    Case 1
        m_sMode = "Joint stereo"
    Case 2
        m_sMode = "Dual channel"
    Case 3
        m_sMode = "Mono"
End Select

' READING MODE EXTENSION (I DON'T KNOW WHAT IT IS)
m_iModeExt = CInt(b3 And &H30) / CInt(&H10)

' READING COPYRIGHT BIT
If (b3 And &H8) = &H8 Then m_bCopyrighted = True

' READING ORIGINAL HOME BIT
If (b3 And &H4) = &H4 Then m_bOriginal = True

' READING EMPHASIS INFO
Select Case b3 And &H3
    Case 0
        m_sEmphasis = "None"
    Case 2
        m_sEmphasis = "Not defined"
    Case 1
        m_sEmphasis = "50/15 ms"
    Case 3
        m_sEmphasis = "CCITT j. 17"
End Select
    
GoTo finish

fault:
m_bValidHeader = False
Close filenum

finish:
End Sub


' FUNCTION FOR GETTING BITRATE INFO
Private Function fnGetBitrate(arg1 As Integer, arg2 As Integer, arg3 As Integer) As Integer
Dim I As Integer
Dim A(1, 3, 15) As Integer
For I = 1 To 14
    A(1, 3, I) = I * 32
    If I < 5 Then
        A(1, 2, I) = 8 * (I + 4)
        A(1, 1, I) = 8 * (I + 3)
    Else
        A(1, 2, I) = A(1, 2, I - 4) * 2
        A(1, 1, I) = A(1, 1, I - 4) * 2
    End If
    If I < 9 Then
        A(0, 1, I) = I * 8
    Else
        A(0, 1, I) = (I - 4) * 16
    End If
    A(0, 2, I) = A(0, 1, I)
Next
A(1, 2, 1) = 32
A(0, 3, 1) = 32
A(0, 3, 2) = 48
A(0, 3, 3) = 56
A(0, 3, 4) = 64
A(0, 3, 5) = 80
A(0, 3, 6) = 96
A(0, 3, 7) = 112
A(0, 3, 8) = 128
A(0, 3, 9) = 144
A(0, 3, 10) = 160
A(0, 3, 11) = 176
A(0, 3, 12) = 192
A(0, 3, 13) = 224
A(0, 3, 14) = 256

fnGetBitrate = A(arg1, arg2, arg3)
If arg3 = 15 Then fnGetBitrate = 1
If arg3 = 0 Then fnGetBitrate = 0
End Function



