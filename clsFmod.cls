VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFMOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ClsFMOD - An easy to use class interface for FMOD
'Version 1.2 - 16:29 2003-10-02
'By Paul Berlin for Simple Amp

Option Explicit

Private Const SHOW_ERROR_MSGBOX As Boolean = True

Dim lStreamPtr  As Long   'Pointer to loaded stream
Dim lChannel    As Long   'Pointer to used channel
Dim lModulePtr  As Long   'Pointer to loaded module

Private Sub Class_Terminate()
  On Error Resume Next
  
  ManStopped = True
  
  FSOUND_StopSound FSOUND_ALL
  If lStreamPtr <> 0 Then FSOUND_Stream_Close lStreamPtr
  If lModulePtr <> 0 Then
    FMUSIC_StopSong lModulePtr
    FMUSIC_FreeSong lModulePtr
  End If
  FSOUND_Close
  
End Sub

'------------------------
'INITIALIZATION FUNCTIONS
'------------------------

Public Function Init(ByVal lMixrate As Long, ByVal lMaxChannels As Long) As Boolean
  'This inits FMOD with selected settings
  On Error GoTo errh
  Init = True
  
  If FSOUND_Init(lMixrate, lMaxChannels, FSOUND_INIT_GLOBALFOCUS Or FSOUND_INIT_ACCURATEVULEVELS Or FSOUND_INIT_ENABLEOUTPUTFX) = False Then GoTo errh
  
  Exit Function
errh:
  If SHOW_ERROR_MSGBOX Then MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "FMOD Error"
  Init = False
End Function

Public Function ReturnDrivers(ByRef sDrvArray() As String) As Boolean
  'This returns the names of availabe drives in the array
  'Output must be set before using this or there will be 0 drivers
  On Error GoTo errh
  
  Dim X As Integer
  
  If FSOUND_GetNumDrivers > 0 Then
  
    ReDim sDrvArray(FSOUND_GetNumDrivers - 1)
    For X = 0 To FSOUND_GetNumDrivers - 1
      sDrvArray(X) = GetStringFromPointer(FSOUND_GetDriverName(X))
    Next X
    
    ReturnDrivers = True
  Else
    ReturnDrivers = False
  End If
  
  Exit Function
errh:
  If SHOW_ERROR_MSGBOX Then MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "FMOD Error"
  ReturnDrivers = False
End Function

Public Function InitOutput(ByVal eOutType As FSOUND_OUTPUTTYPES) As Boolean
  'Sets output type
  'MUST BE CALLED BEFORE Init, If you do not call this before Init,
  'the best output type will be autodetected.
  'use ReturnDrivers to see which drives support the selected output type
  On Error GoTo errh
  InitOutput = True
  
  If FSOUND_SetOutput(eOutType) = False Then GoTo errh
  
  Exit Function
errh:
  If SHOW_ERROR_MSGBOX Then MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "FMOD Error"
  InitOutput = False
End Function

Public Function InitMixer(ByVal eMixerType As FSOUND_MIXERTYPES) As Boolean
  'Sets mixer to use
  'MUST BE CALLED BEFORE Init, If you do not call this before Init,
  'the best mixer type will be autodetected.
  On Error GoTo errh
  InitMixer = True
  
  If FSOUND_SetMixer(eMixerType) = False Then GoTo errh
  
  Exit Function
errh:
  If SHOW_ERROR_MSGBOX Then MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "FMOD Error"
  InitMixer = False
End Function

Public Function InitBuffer(ByVal lBufferLenMs As Long) As Boolean
  'Sets buffer size in milliseconds.
  'Buffer size < 100 might cause problem on certain settings.
  'MUST BE CALLED BEFORE Init, If you do not call this before init,
  'an optimal value will be autodetected.
  InitBuffer = True
  
  If FSOUND_SetBufferSize(lBufferLenMs) = False Then GoTo errh

  Exit Function
errh:
  If SHOW_ERROR_MSGBOX Then MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "FMOD Error"
  InitBuffer = False
End Function

Public Function InitDriver(ByVal iDriveNum As Integer) As Boolean
  'Must be called before init, If you do not call this before init,
  'an the default driver will be selected.
  'Drive num 0 will set to default driver.
  'You can get available drivers and their numbers with ReturnDrivers.
  On Error GoTo errh
  InitDriver = True
  
  If FSOUND_SetDriver(iDriveNum) = False Then GoTo errh
  
  Exit Function
errh:
  If SHOW_ERROR_MSGBOX Then MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "FMOD Error"
  InitDriver = False
End Function

'----------------
'STREAM FUNCTIONS
'----------------
'Plays mp3, mp2, wma, ogg, wav, asf, midi

Public Function StreamPlay(ByVal sFilename As String, Optional ByVal bVBR As Boolean = False, Optional bVolume As Byte = 255) As Boolean
  'Will play the sFilename stream.
  'bVBR is for variable bitrate mp3s.
  On Error GoTo errh
  StreamPlay = True

  FSOUND_StopSound FSOUND_ALL
  If lStreamPtr <> 0 Then FSOUND_Stream_Close lStreamPtr
  
  If bVBR Then
    lStreamPtr = FSOUND_Stream_Open(sFilename, FSOUND_NORMAL Or FSOUND_MPEGACCURATE, 0, 0)
  Else
    lStreamPtr = FSOUND_Stream_Open(sFilename, FSOUND_NORMAL, 0, 0)
  End If
  
  If lStreamPtr = 0 Then
    MsgBox "Could not recognize & play """ & sFilename & """.", vbExclamation, "FMOD Error"
    StreamPlay = False
    Exit Function
  End If
  
  lChannel = FSOUND_Stream_Play(FSOUND_FREE, lStreamPtr)
  FSOUND_SetVolume lChannel, bVolume
  
  'lChannel = FSOUND_PlaySoundEx(FSOUND_FREE, lStreamPtr, 0, True)
  'FSOUND_SetPaused lChannel, False
  'Call FSOUND_Stream_SetEndCallback(lStreamPtr, AddressOf StreamEndCallback, 0)
  
  Exit Function
errh:
  If SHOW_ERROR_MSGBOX Then MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "FMOD Error"
  StreamPlay = False
End Function

Public Function StreamPausePlay() As Boolean
  'this will pause or play the current stream
  On Error GoTo errh
  
  If lStreamPtr <> 0 Then
    StreamPausePlay = True
    
    ManStopped = Not CBool(FSOUND_GetPaused(lChannel))
    FSOUND_SetPaused lChannel, ManStopped
    
  End If
  
  Exit Function
errh:
  If SHOW_ERROR_MSGBOX Then MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "FMOD Error"
  StreamPausePlay = False
End Function

Public Function StreamStop() As Boolean
  'this will stop and restart stream
  On Error GoTo errh
  StreamStop = True

  If lStreamPtr <> 0 Then
    StreamStop = True
    
    If lChannel <> 0 Then
      ManStopped = True
      FSOUND_StopSound lChannel
      FSOUND_Stream_SetTime lStreamPtr, 0
      lChannel = 0
    Else
      lChannel = FSOUND_Stream_Play(FSOUND_FREE, lStreamPtr)
    End If
    
  End If

  Exit Function
errh:
  If SHOW_ERROR_MSGBOX Then MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "FMOD Error"
  StreamStop = False
End Function

Public Function StreamUnload() As Boolean
  'this will stop and unload stream
  On Error GoTo errh
  StreamUnload = True

  ManStopped = True
  FSOUND_StopSound FSOUND_ALL
  If lStreamPtr <> 0 Then FSOUND_Stream_Close lStreamPtr
  lStreamPtr = 0
  lChannel = 0

  Exit Function
errh:
  If SHOW_ERROR_MSGBOX Then MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "FMOD Error"
  StreamUnload = False
End Function

Public Property Get StreamFrequency() As Long
  'NOTE: FSOUND_GetFrequency(Channel) does not seem to work with wma, asf
  StreamFrequency = FSOUND_GetFrequency(lChannel)
End Property

Public Property Let StreamFrequency(ByVal lFreq As Long)
  FSOUND_SetFrequency lChannel, lFreq
End Property

Public Property Get StreamVolume() As Byte
  StreamVolume = FSOUND_GetVolume(lChannel)
End Property

Public Property Let StreamVolume(ByVal bVolume As Byte)
  FSOUND_SetVolume lChannel, bVolume
End Property

Public Property Get StreamIsLoaded() As Boolean
  StreamIsLoaded = (lStreamPtr <> 0)
End Property

Public Property Get StreamIsPlaying() As Boolean
  If lChannel <> 0 Then
    StreamIsPlaying = (FSOUND_GetPaused(lChannel) = 0)
  End If
End Property

Public Property Get StreamSongLen() As Long
  If lStreamPtr <> 0 Then
    StreamSongLen = FSOUND_Stream_GetLengthMs(lStreamPtr)
  End If
End Property

Public Property Get StreamSongPos() As Long
  If lStreamPtr <> 0 And lChannel <> 0 Then
    StreamSongPos = FSOUND_Stream_GetTime(lStreamPtr)
  End If
End Property

Public Property Let StreamSongPos(ByVal lPos As Long)
  If lStreamPtr <> 0 And lChannel <> 0 Then
    FSOUND_Stream_SetTime lStreamPtr, lPos
  End If
End Property

Public Property Get StreamKbps() As Long
  'This calculates Kbps by dividing length of stream in bits with length of stream in seconds.
  On Error Resume Next
  StreamKbps = (FSOUND_Stream_GetLength(lStreamPtr) \ (FSOUND_Stream_GetLengthMs(lStreamPtr) \ 1000)) * 8 \ 1000
End Property

Public Property Get StreamPanning() As Byte
  StreamPanning = FSOUND_GetPan(lChannel)
End Property

Public Property Let StreamPanning(ByVal bPanning As Byte)
  If bPanning = 128 Then
    FSOUND_SetPan FSOUND_ALL, FSOUND_STEREOPAN
  Else
    FSOUND_SetPan FSOUND_ALL, bPanning
  End If
End Property

Public Property Get StreamSurround() As Boolean
  StreamSurround = FSOUND_GetSurround(lChannel)
End Property

Public Property Let StreamSurround(ByVal bSurround As Boolean)
  Call FSOUND_SetSurround(FSOUND_ALL, bSurround)
End Property

'the below stream functions are used to get info from the stream
'when you do not want to play it
'use GetStreamOpenFile to open a file
'use GetStreamCloseFile to close it after you are done
'not closing after you are done will produce errors or hangs

Public Function GetStreamOpenFile(ByVal sFilename As String, Optional ByVal bVBR As Boolean = False) As Long
  'opens file and returns handle
  'make sure you close handle with GetStreamCloseFile later
  
  If bVBR Then
    GetStreamOpenFile = FSOUND_Stream_Open(sFilename, FSOUND_NORMAL Or FSOUND_MPEGACCURATE, 0, 0)
  Else
    GetStreamOpenFile = FSOUND_Stream_Open(sFilename, FSOUND_NORMAL, 0, 0)
  End If
  
End Function

Public Sub GetStreamCloseFile(ByVal lHandle As Long)
  'closes file opened with GetStreamOpenFile
  
  Call FSOUND_Stream_Close(lHandle)
  
End Sub

Public Function GetStreamLength(ByVal lHandle As Long) As Long
  'gets length of opened stream
  
  GetStreamLength = FSOUND_Stream_GetLengthMs(lHandle) / 1000
  
End Function

Public Function GetStreamTag(ByVal lHandle As Long, ByVal eTag As FSOUND_TAGFIELD_TYPE, ByRef sTagName As String) As String
  'gets a tag from opened stream handle
  Dim lFieldvalue As Long
  Dim lFieldlength As Long
  
  If FSOUND_Stream_FindTagField(lHandle, eTag, sTagName, lFieldvalue, lFieldlength) Then
    If eTag = FSOUND_TAGFIELD_ID3V2 Then
      If sTagName = "COMM" Or Left(sTagName, 1) = "W" Then
        'comment tag may contain other information than just the comment,
        'seperated by chr(0), but the comment is always last.
        'We extract the comment from the other information below
        GetStreamTag = GetStringFromPointerEx(lFieldvalue, lFieldlength)
        GetStreamTag = Left(GetStreamTag, Len(GetStreamTag) - 1)
        GetStreamTag = Mid(GetStreamTag, InStrRev(GetStreamTag, Chr(0)) + 1)
      Else
        GetStreamTag = Trim(Replace(GetStringFromPointerEx(lFieldvalue, lFieldlength), Chr(0), " "))
      End If
    Else
      GetStreamTag = Trim(GetStringFromPointer(lFieldvalue))
    End If
  End If
  
End Function

Public Function GetStreamTagNumTags(ByVal lHandle As Long) As Long
  Call FSOUND_Stream_GetNumTagFields(lHandle, GetStreamTagNumTags)
End Function

Public Sub GetStreamTagByNum(ByVal lHandle As Long, ByVal lNum As Long, ByRef eTagType As FSOUND_TAGFIELD_TYPE, ByRef sTagName As String, ByRef sTagValue As String)
  Dim lVal As Long, lLen As Long, lName As Long
  Call FSOUND_Stream_GetTagField2(lHandle, lNum, eTagType, lName, lVal, lLen)
  
  sTagName = sNT(GetStringFromPointer(lName))
  
  If eTagType = FSOUND_TAGFIELD_ID3V2 Then
    If sTagName = "COMM" Or Left(sTagName, 1) = "W" Then
      'comment tag may contain other information than just the comment,
      'seperated by chr(0), but the comment is always last.
      'We extract the comment from the other information below
      sTagValue = GetStringFromPointerEx(lVal, lLen)
      sTagValue = Left(sTagValue, Len(sTagValue) - 1)
      sTagValue = Mid(sTagValue, InStrRev(sTagValue, Chr(0)) + 1)
    Else
      sTagValue = Trim(Replace(GetStringFromPointerEx(lVal, lLen), Chr(0), " "))
    End If
  Else
    sTagValue = Trim(Replace(GetStringFromPointerEx(lVal, lLen), Chr(0), " "))
  End If
  
End Sub

'---------------
'MUSIC FUNCTIONS
'---------------
'Plays Mod, S3m, Xm, It, Midi, Rmi, Sgt

Public Function MusicPlay(ByVal sFilename As String, Optional ByVal bLoopit As Boolean = False, Optional bVolume As Byte = 255) As Boolean
  'Will play the filename music
  On Error GoTo errh
  MusicPlay = True

  If lModulePtr <> 0 Then
    FMUSIC_StopSong lModulePtr
    FMUSIC_FreeSong lModulePtr
  End If
  lModulePtr = FMUSIC_LoadSong(sFilename)
  
  If lModulePtr = 0 Then
    MsgBox "Could not recognize & play """ & sFilename & """.", vbExclamation, "FMOD Error"
    MusicPlay = False
    Exit Function
  End If
  
  Call FMUSIC_SetLooping(lModulePtr, bLoopit)
  Call FMUSIC_PlaySong(lModulePtr)
  FMUSIC_SetMasterVolume lModulePtr, bVolume

  Exit Function
errh:
  If SHOW_ERROR_MSGBOX Then MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "FMOD Error"
  MusicPlay = False
End Function

Public Function MusicStop() As Boolean
  'this will stop and resume music
  On Error GoTo errh

  If lModulePtr <> 0 Then
    MusicStop = True
    If FMUSIC_IsPlaying(lModulePtr) Then
      FMUSIC_StopSong lModulePtr
    Else
      FMUSIC_PlaySong lModulePtr
    End If
  End If
    
  Exit Function
errh:
  If SHOW_ERROR_MSGBOX Then MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "FMOD Error"
  MusicStop = False
End Function

Public Function MusicUnload() As Boolean
  'this will stop and unload music
  On Error GoTo errh

  If lModulePtr <> 0 Then
    MusicUnload = True
    FMUSIC_StopSong lModulePtr
    FMUSIC_FreeSong lModulePtr
    lModulePtr = 0
  End If
    

  Exit Function
errh:
  If SHOW_ERROR_MSGBOX Then MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "FMOD Error"
  MusicUnload = False
End Function

Public Function MusicPausePlay() As Boolean
  'this will stop and unload music
  On Error GoTo errh

  If lModulePtr <> 0 Then
    MusicPausePlay = True

    FMUSIC_SetPaused lModulePtr, Not FMUSIC_GetPaused(lModulePtr)
  End If
    

  Exit Function
errh:
  If SHOW_ERROR_MSGBOX Then MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "FMOD Error"
  MusicPausePlay = False
End Function

Public Property Get MusicVolume() As Byte
  MusicVolume = FMUSIC_GetMasterVolume(lModulePtr)
End Property

Public Property Let MusicVolume(ByVal bVolume As Byte)
  FMUSIC_SetMasterVolume lModulePtr, bVolume
End Property

Public Property Get MusicName() As String
  MusicName = GetStringFromPointer(FMUSIC_GetName(lModulePtr))
End Property

Public Property Get MusicNumChannels() As Long
  MusicNumChannels = FMUSIC_GetNumChannels(lModulePtr)
End Property

Public Property Get MusicNumIntruments() As Long
  MusicNumIntruments = FMUSIC_GetNumInstruments(lModulePtr)
End Property

Public Property Get MusicNumOrders() As Long
  MusicNumOrders = FMUSIC_GetNumOrders(lModulePtr)
End Property

Public Property Get MusicNumPatterns() As Long
  MusicNumPatterns = FMUSIC_GetNumPatterns(lModulePtr)
End Property

Public Property Get MusicNumSamples() As Long
  MusicNumSamples = FMUSIC_GetNumSamples(lModulePtr)
End Property

Public Property Get MusicBPM() As Long
  MusicBPM = FMUSIC_GetBPM(lModulePtr)
End Property

Public Property Get MusicIsLoaded() As Boolean
  If lModulePtr <> 0 Then MusicIsLoaded = True
End Property

Public Property Get MusicOrder() As Long
  MusicOrder = FMUSIC_GetOrder(lModulePtr)
End Property

Public Property Let MusicOrder(ByVal lNewOrder As Long)
  FMUSIC_SetOrder lModulePtr, lNewOrder
End Property

Public Property Get MusicPattern() As Long
  MusicPattern = FMUSIC_GetPattern(lModulePtr)
End Property

Public Property Get MusicRow() As Long
  MusicRow = FMUSIC_GetRow(lModulePtr)
End Property

Public Property Get MusicSpeed() As Long
  MusicSpeed = FMUSIC_GetSpeed(lModulePtr)
End Property

Public Property Get MusicTime() As Long
  MusicTime = FMUSIC_GetTime(lModulePtr)
End Property

Public Property Get MusicIsPlaying() As Boolean
  If FMUSIC_GetPaused(lModulePtr) = False Then
    MusicIsPlaying = FMUSIC_IsPlaying(lModulePtr)
  End If
End Property

Public Property Get MusicIsFinished() As Boolean
  MusicIsFinished = FMUSIC_IsFinished(lModulePtr)
End Property

Public Property Get MusicNumRows() As Long
  MusicNumRows = FMUSIC_GetPatternLength(lModulePtr, FMUSIC_GetOrder(lModulePtr))
End Property

Public Property Let MusicPanSep(ByVal sPan As Single)
  FMUSIC_SetPanSeperation lModulePtr, sPan
End Property

Public Property Get MusicSongLen() As Long
  'This will calculate the songs length in ms
  'This is mostly accurate, but can be highly inaccurate
  On Error Resume Next
  Dim Kbps As Double, st As Double, i As Long, t As Long
  
  Kbps = (FMUSIC_GetBPM(lModulePtr) * 2) / 5
  st = ((1 / Kbps) * 1000) * FMUSIC_GetSpeed(lModulePtr)

  For i = 0 To (FMUSIC_GetNumOrders(lModulePtr) - 1)
    t = t + FMUSIC_GetPatternLength(lModulePtr, i)
  Next i

  MusicSongLen = t * st
End Property

Public Property Get MusicSongPos() As Long
  'This will calculate the current song time in ms
  'This is mostly accurate, but can be highly inaccurate
  On Error Resume Next
  Dim Kbps As Double, st As Double, i As Long, t As Long
  
  Kbps = (FMUSIC_GetBPM(lModulePtr) * 2) / 5
  st = ((1 / Kbps) * 1000) * FMUSIC_GetSpeed(lModulePtr)

  For i = 0 To (FMUSIC_GetOrder(lModulePtr))
    t = t + FMUSIC_GetPatternLength(lModulePtr, i)
  Next i
  t = t - (FMUSIC_GetPatternLength(lModulePtr, FMUSIC_GetOrder(lModulePtr)) - FMUSIC_GetRow(lModulePtr))
  
  MusicSongPos = t * st
End Property

Public Function MusicGetRows(ByVal lOrder As Long) As Long
  MusicGetRows = FMUSIC_GetPatternLength(lModulePtr, lOrder)
End Function

Public Function GetMusicOK(ByVal sFilename As String) As Boolean
  'This opens an music file and checks if fmod recognises it
  On Error GoTo errh
  Dim tmpMusic As Long
  
  tmpMusic = FMUSIC_LoadSong(sFilename)
  GetMusicOK = (FMUSIC_GetType(tmpMusic) <> 0)
  FMUSIC_FreeSong tmpMusic
  
errh:
End Function

Public Sub GetMusicData(ByVal sFilename As String, ByRef lLength As Long, ByRef sName As String)
  'This opens an music file, gets the length and name, and closes it
  'This will calculate the songs length in ms
  'This is mostly accurate, but can be highly inaccurate
  On Error Resume Next
  Dim Kbps As Double, st As Double, i As Long, t As Long
  Dim tmpMusic As Long
  
  tmpMusic = FMUSIC_LoadSong(sFilename)  'load song
  If FMUSIC_GetType(tmpMusic) <> FMUSIC_TYPE_NONE Then
  
    Call FMUSIC_SetMasterVolume(tmpMusic, 0)  'set volume to 0
    Call FMUSIC_PlaySong(tmpMusic)  'play song, must be played to get BPM
    'set order to in the middle, because this should increase the accuracy of the time
    Call FMUSIC_SetOrder(tmpMusic, FMUSIC_GetNumOrders(tmpMusic) / 2)
    sName = Trim(GetStringFromPointer(FMUSIC_GetName(tmpMusic))) 'get name
    
    'calc time
    Kbps = (FMUSIC_GetBPM(tmpMusic) * 2) / 5
    st = ((1 / Kbps) * 1000) * FMUSIC_GetSpeed(tmpMusic)
    For i = 0 To (FMUSIC_GetNumOrders(tmpMusic) - 1)
      t = t + FMUSIC_GetPatternLength(tmpMusic, i)
    Next i
    lLength = (t * st) \ 1000
    If lLength > 3600 Then lLength = 0
    If lLength < 0 Then lLength = 0
    
    'stop song
    FMUSIC_StopSong tmpMusic
    
  End If
  
  'unload song
  FMUSIC_FreeSong tmpMusic
  
End Sub

'------------------
'CD AUDIO FUNCTIONS
'------------------

Public Sub CDEject(Optional ByVal bDrive As Byte = 0)
  FSOUND_CD_Eject bDrive
End Sub

Public Function CDGetNumTracks(Optional ByVal bDrive As Byte = 0) As Long
  CDGetNumTracks = FSOUND_CD_GetNumTracks(bDrive)
End Function

Public Function CDPlay(ByVal lTrack As Long, Optional ByVal bDrive As Byte = 0) As Boolean
  CDPlay = True
  If FSOUND_CD_SetPlayMode(bDrive, FSOUND_CD_PLAYONCE) = 0 Then CDPlay = False
  If FSOUND_CD_Play(bDrive, lTrack) = 0 Then CDPlay = False
End Function

Public Sub CDPausePlay(Optional ByVal bDrive As Byte = 0)
  FSOUND_CD_SetPaused bDrive, Not CBool(FSOUND_CD_GetPaused(bDrive))
End Sub

Public Sub CDStop(Optional ByVal bDrive As Byte = 0)
  FSOUND_CD_Stop bDrive
End Sub

Public Function CDGetTrackLength(ByVal lTrack As Long, Optional ByVal bDrive As Byte = 0) As Long
  CDGetTrackLength = FSOUND_CD_GetTrackLength(bDrive, lTrack)
End Function

Public Function CDGetTrackTime(Optional ByVal bDrive As Byte = 0) As Long
  CDGetTrackTime = FSOUND_CD_GetTrackTime(bDrive)
End Function

Public Sub CDSetTrackTime(ByVal lMs As Long, Optional ByVal bDrive As Byte = 0)
  FSOUND_CD_SetTrackTime bDrive, lMs
End Sub

Public Sub CDSetVolume(ByVal bVol As Byte, Optional ByVal bDrive As Byte = 0)
  FSOUND_CD_SetVolume bDrive, bVol
End Sub

'---------------------
'OTHER MISC. FUNCTIONS
'---------------------

Public Property Get CPUUsage() As Single
  CPUUsage = FSOUND_GetCPUUsage
End Property

Public Property Get FMODVersion() As Single
  FMODVersion = FSOUND_GetVersion
End Property

Public Property Get ChannelsPlaying() As Long
  ChannelsPlaying = FSOUND_GetChannelsPlaying
End Property

Public Property Let SpeakerSetup(ByVal Spk As FSOUND_SPEAKERMODES)
  FSOUND_SetSpeakerMode Spk
End Property

Public Sub GetVU(ByRef nLeft As Single, ByRef nRight As Single)
  Call FSOUND_GetCurrentLevels(lChannel, nLeft, nRight)
End Sub

Public Sub GetSpecificVU(ByVal lSpecChannel, ByRef nLeft As Single, ByRef nRight As Single)
  Call FSOUND_GetCurrentLevels(lSpecChannel, nLeft, nRight)
End Sub

Public Function TestFile(ByVal sFilename As String) As Byte
  'This function tests the file if fmod can open it.
  'Returns: 0 = No
  '1 = Stream
  '2 = Module
  On Error Resume Next
  Dim tmpPtr As Long
  
  tmpPtr = FSOUND_Stream_Open(sFilename, FSOUND_NORMAL, 0, 0)
  If tmpPtr <> 0 Then
    FSOUND_Stream_Close tmpPtr
    TestFile = 1
    Exit Function
  End If
  
  tmpPtr = FMUSIC_LoadSong(sFilename)
  If tmpPtr <> 0 Then
    FMUSIC_FreeSong tmpPtr
    TestFile = 2
    Exit Function
  End If
  
End Function
