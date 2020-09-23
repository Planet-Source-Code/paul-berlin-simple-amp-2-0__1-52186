Attribute VB_Name = "modMain"
'This modules contains program startup sub Main as well as other
'program specific subs & functions and APIs

Option Explicit

'Declaration of API Functions
Public Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'public APIs for drawing visualizations
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

'Used to get screen area
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48

Private Sub Main()
  'startup sub
  On Error Resume Next
  Dim sw As String
  Dim Comnd As String
  Dim Ignore As Boolean
  Const sB As String = "Simple Amp 2.0 RC 010"
  
  'format switch (only one is allowed + filename)
  Comnd = Trim(Command)
  If InStr(1, Comnd, " ") <> 0 Then
    sw = Left(Comnd, InStr(1, Comnd, " ") - 1)
  Else
    sw = Comnd
  End If
  
  'handle switch
  Select Case sw
    Case "-playpause", "-pauseplay"
      SendString sB, "0"
    Case "-prev", "-previous"
      SendString sB, "1"
    Case "-next"
      SendString sB, "2"
    Case "-toggleshuffle", "-shuffle"
      SendString sB, "3"
    Case "-togglerepeat", "-repeat"
      SendString sB, "4"
    Case "-showhide", "-hideshow"
      SendString sB, "5"
    Case "-restart", "-stopplay"
      SendString sB, "7"
    Case "-forward", "-fwd"
      SendString sB, "8"
    Case "-back", "-bck"
      SendString sB, "9"
    Case "-togglesurround", "-surround"
      SendString sB, "a"
    Case "-showopen"
      SendString sB, "b"
    Case "-ignore"
      Ignore = True
    Case "-open"
      FileOpen = Right(Comnd, Len(Comnd) - InStr(1, Comnd, " "))
  End Select

  'if the switch isn't -ignore, end simple amp if there is one copy of the program
  'running already.
  If Not Ignore Then
    If FindWindow(vbNullString, "Simple Amp 2.0 RC 010") <> 0 Then
      If Len(FileOpen) > 0 Then SendString sB, "6" & FileOpen
      End
    End If
  End If
  
  cLog.bLogDebug = LOG_DEBUG
  cLog.iLogDebugLevel = LOG_LEVEL
  
  cLog.Log "STARTING UP SIMPLE AMP", 5
  cLog.Log "----------------------", 5
  
  LoadSettings  'Load settings from ini-file
  
  Genre = Split(sGenreMatrix, "|") 'setup ID3v1 genre array
  
  If Settings.AssOnStart Then AssociateTypes
  
  ReDim Images(0)
  ReDim Library(0)
  ReDim LibraryIndex(0)
  ReDim Playlist(0)
  
  If Settings.UseLibrary Then LoadIndex
  fLibrary.FileName = App.Path & "\database.db"
  fStat.FileName = App.Path & "\stat.db"
  
  Load frmMain
  Load frmPlaylist
  
  frmMain.Move Settings.WinLeft, Settings.WinTop
  
  frmMain.Init
End Sub

Public Sub LoadSettings()
  'This sub loads program setting from ini-file
  On Error Resume Next
  Dim cINI As New clsINI, X As Long
  
  'show about first time started
  If Not FileExists(App.Path & "\settings.ini") Then frmAbout.Show
  
  cINI.sFilename = App.Path & "\settings.ini"
  
  cLog.Log "LOADING SETTINGS (settings.ini)...", 2, False
  cLog.StartTimer
  
  With Settings
    cINI.sSection = "Simple Amp"
    For X = 0 To 4
      .LibSearch(X) = cINI.ReadNumber("LibSearch" & X, 1)
    Next
    lCurVisPreset = cINI.ReadString("CurVisPreset", 1)
    .BrowseMode = cINI.ReadNumber("BrowseMode", 0)
    .DynamicColumns = cINI.ReadString("DynamicColumns", False)
    .LibHideNonMusic = cINI.ReadString("HideNonMusic", True)
    .AddAutosize = cINI.ReadString("AddAutosize", False)
    .LibAutosize = cINI.ReadString("Autosize", True)
    .LibGetName = cINI.ReadString("GetName", False)
    .LibFilter = cINI.ReadNumber("PrimaryFilter", 0)
    .Advance = cINI.ReadString("Advance", True)
    .AddView = cINI.ReadString("AddView", 0)
    CurrentSkin = cINI.ReadString("Skin", DEFAULTSKIN)
    .Repeat = cINI.ReadNumber("Repeat", 0)
    .Shuffle = cINI.ReadNumber("Shuffle", 0)
    .CurrentVolume = cINI.ReadNumber("Volume", 255)
    Spectrum = cINI.ReadNumber("Spectrum", 1)
    .WinLeft = cINI.ReadNumber("WindowX", 500)
    .WinTop = cINI.ReadNumber("WindowY", 400)
    frmPlaylist.Left = cINI.ReadNumber("PlaylistX", 500)
    frmPlaylist.Top = cINI.ReadNumber("PlaylistY", 2380)
    .AlwaysTray = cINI.ReadString("AlwaysTray", True)
    .OnTop = cINI.ReadString("OnTop", False)
    .TrayIcon = cINI.ReadNumber("TrayIcon", 1)
    .StartInTray = cINI.ReadString("StartInTray", False)
    .Snap = cINI.ReadString("Snap", True)
    .AddDir = cINI.ReadString("AddDir", App.Path)
    Playing = cINI.ReadNumber("Playing", 1)
    .PlaylistOn = cINI.ReadString("PlaylistOn", False)
    .PlaylistSmall = cINI.ReadString("PlaylistSmall", True)
    .Fade = cINI.ReadString("Fade", True)
    .AddMax = cINI.ReadString("AddMax", False)
    .AddWidth = cINI.ReadNumber("AddWidth")
    .AddHeight = cINI.ReadNumber("AddHeight")
    .AddFolderWidth = cINI.ReadNumber("AddFolderWidth")
    .LibMax = cINI.ReadString("LibMax", False)
    .LibWidth = cINI.ReadNumber("LibWidth")
    .LibHeight = cINI.ReadNumber("LibHeight")
    .UseLibrary = cINI.ReadString("UseLibrary", True)
    .ButtonDefault = cINI.ReadString("ButtonDefault", False)
    Docked = cINI.ReadString("Docked", False)
    DockedLeft = cINI.ReadNumber("DockedLeft", 0)
    DockedTop = cINI.ReadNumber("DockedTop", 0)
    SnapWidth = cINI.ReadNumber("SnapWidth", 15)
    .NoPresetFade = cINI.ReadString("NoPresetFade", False)
    
    'scope settings
    With ScopeSettings
      cINI.sSection = "VisScope"
      .bBrushSizeL = cINI.ReadNumber("BrushSizeL", 2)
      .bBrushSizeR = cINI.ReadNumber("BrushSizeR", 1)
      .bDetail = cINI.ReadNumber("Detail", 0)
      .bFade = cINI.ReadNumber("Fade", 0)
      .bFall = cINI.ReadNumber("Fall", 1)
      .bPeakCount = cINI.ReadNumber("PeakCount", 2)
      .bPeakDetail = cINI.ReadNumber("PeakDetail", 1)
      .bPeaks = cINI.ReadNumber("Peaks", 0)
      .bSkip = cINI.ReadNumber("Skip", 2)
      .bType = cINI.ReadNumber("Type", 1)
      .lColorL = cINI.ReadNumber("ColorL", RGB(106, 121, 90))
      .lColorPeakL = cINI.ReadNumber("ColorPeakL", RGB(106, 121, 90))
      .lColorPeakR = cINI.ReadNumber("ColorPeakR", RGB(76, 91, 60))
      .lColorR = cINI.ReadNumber("ColorR", RGB(76, 91, 60))
      .lPeakDec = cINI.ReadNumber("PeakDec", 75)
      .lPeakPause = cINI.ReadNumber("PeakPause", 250)
    End With
    
    'spectrum settings
    With SpectrumSettings
      cINI.sSection = "VisSpectrum"
      .bCorrection = cINI.ReadNumber("Correction", 0)
      .bBarSize = cINI.ReadNumber("BarSize", 2)
      .bBrushSize = cINI.ReadNumber("BrushSize", 1)
      .bDrawStyle = cINI.ReadNumber("DrawStyle", 1)
      .bFade = cINI.ReadNumber("Fade", 0)
      .bFall = cINI.ReadNumber("Fall", 1)
      .bPeakFall = cINI.ReadNumber("PeakFall", 1)
      .bPeaks = cINI.ReadNumber("Peaks", 1)
      .bType = cINI.ReadNumber("Type", 1)
      .iView = cINI.ReadNumber("View", 128)
      .lColorDn = cINI.ReadNumber("ColorDn", RGB(69, 80, 56))
      .lColorLine = cINI.ReadNumber("ColorLine", RGB(69, 80, 56))
      .lColorUp = cINI.ReadNumber("ColorUp", RGB(113, 131, 92))
      .lPause = cINI.ReadNumber("Pause", 0)
      .lPeakColor = cINI.ReadNumber("PeakColor", RGB(55, 64, 45))
      .lPeakDec = cINI.ReadNumber("PeakDec", 7)
      .lPeakPause = cINI.ReadNumber("PeakPause", 250)
      .nDec = cINI.ReadNumber("Dec", 30)
      .nZoom = cINI.ReadNumber("Zoom", 1.7)
    End With
    
    'volume settings
    With VolumeSettings
      cINI.sSection = "VisVolume"
      .bDrawStyle = cINI.ReadNumber("DrawStyle", 0)
      .bFade = cINI.ReadNumber("Fade", 0)
      .bFall = cINI.ReadNumber("Fall", 1)
      .bType = cINI.ReadNumber("Type", 0)
      .lColorDn = cINI.ReadNumber("ColorDn", RGB(69, 80, 56))
      .lColorUp = cINI.ReadNumber("ColorUp", RGB(113, 131, 92))
      .lPause = cINI.ReadNumber("Pause", 0)
      .nDec = cINI.ReadNumber("Dec", 30)
      .sFile = cINI.ReadString("File")
    End With
    
    'beat settings
    With BeatSettings
      cINI.sSection = "VisBeat"
      .bFade = cINI.ReadNumber("Fade", 5)
      .bType = cINI.ReadNumber("Type", 0)
      .iDetectHigh = cINI.ReadNumber("DetectHigh", 25)
      .iDetectLow = cINI.ReadNumber("DetectLow", 0)
      .nMin = cINI.ReadNumber("Min", 0.75)
      .nMulti = cINI.ReadNumber("Multi", 3)
      .nRotMin = cINI.ReadNumber("RotMin", 0.2)
      .nRotMove = cINI.ReadNumber("RotMove", 0.1)
      .nRotSpeed = cINI.ReadNumber("RotSpeed", 0.15)
      .sFile = cINI.ReadString("File")
    End With
       
    'Equalizer
    cINI.sSection = "Equalizer"
    .EQon = cINI.ReadString("EQon", False)
    For X = 0 To 9
      EQValue(X) = cINI.ReadNumber("EQ" & X)
    Next
    
    'Visualization options
    cINI.sSection = "Visualization"
    VisUpdateInt = cINI.ReadNumber("UpdateInterval", 22)
       
    'Browser filters
    cINI.sSection = "Filter"
    For X = 0 To UBound(.Filter)
      .Filter(X) = cINI.ReadString("Filter" & X, True)
    Next X
    
    'Association
    cINI.sSection = "Associate"
    For X = 0 To UBound(.AssType)
      .AssType(X) = cINI.ReadString("Type" & X, False)
    Next X
    .AssOnStart = cINI.ReadString("OnStart", False)
    .AssAction = cINI.ReadNumber("Action", 0)
    
    'sound device settings
    cINI.sSection = "Device"
    Settings.DXFXon = cINI.ReadString("DXFXon", False)
    devType = cINI.ReadNumber("Type")
    devDevice = cINI.ReadNumber("Device")
    devMixer = cINI.ReadNumber("Mixer")
    devFreq = cINI.ReadNumber("Frequency", 44100)
    devChannels = cINI.ReadNumber("Channels", 64)
    devBuffer = cINI.ReadNumber("Buffer", 100)
    devPanning = cINI.ReadNumber("Panning", 128)
    devSurround = cINI.ReadString("Surround", False)
    devSpeaker = cINI.ReadNumber("Speaker", 0)
    devPanSep = cINI.ReadNumber("PanSep", 1)
    devDSP = cINI.ReadString("DSP", True)
    
  End With
   
  cLog.Log "DONE. (" & cLog.GetTimer & " ms)", 2
    
End Sub

Public Sub SaveSettings()
  'This sub saves program settings to ini-file
  On Error GoTo ErrHandler
  
  Dim cINI As New clsINI, X As Long
  cINI.sFilename = App.Path & "\settings.ini"
  
  cLog.Log "SAVING SETTINGS (settings.ini)...", 2, False
  cLog.StartTimer
  
  With Settings
    'Program settings
    cINI.sSection = "Simple Amp"
    For X = 0 To 4
      cINI.WriteKey "LibSearch" & X, .LibSearch(X)
    Next
    cINI.WriteKey "BrowseMode", .BrowseMode
    cINI.WriteKey "CurVisPreset", lCurVisPreset
    cINI.WriteKey "DynamicColumns", .DynamicColumns
    cINI.WriteKey "HideNonMusic", .LibHideNonMusic
    cINI.WriteKey "GetName", .LibGetName
    cINI.WriteKey "AutoSize", .LibAutosize
    cINI.WriteKey "AddAutoSize", .AddAutosize
    cINI.WriteKey "PrimaryFilter", .LibFilter
    cINI.WriteKey "AddView", .AddView
    cINI.WriteKey "Skin", CurrentSkin
    cINI.WriteKey "Repeat", .Repeat
    cINI.WriteKey "Shuffle", .Shuffle
    cINI.WriteKey "Advance", .Advance
    cINI.WriteKey "Volume", .CurrentVolume
    cINI.WriteKey "WindowX", frmMain.Left
    cINI.WriteKey "WindowY", frmMain.Top
    cINI.WriteKey "PlaylistX", frmPlaylist.Left
    cINI.WriteKey "PlaylistY", frmPlaylist.Top
    cINI.WriteKey "Spectrum", Spectrum
    cINI.WriteKey "AlwaysTray", .AlwaysTray
    cINI.WriteKey "TrayIcon", .TrayIcon
    cINI.WriteKey "OnTop", .OnTop
    cINI.WriteKey "StartInTray", .StartInTray
    cINI.WriteKey "Snap", .Snap
    cINI.WriteKey "AddDir", .AddDir
    cINI.WriteKey "PlaylistOn", .PlaylistOn
    cINI.WriteKey "PlaylistSmall", .PlaylistSmall
    cINI.WriteKey "Playing", Playing
    cINI.WriteKey "Fade", .Fade
    cINI.WriteKey "AddMax", .AddMax
    cINI.WriteKey "AddWidth", .AddWidth
    cINI.WriteKey "AddHeight", .AddHeight
    cINI.WriteKey "AddFolderWidth", .AddFolderWidth
    cINI.WriteKey "LibMax", .LibMax
    cINI.WriteKey "LibWidth", .LibWidth
    cINI.WriteKey "LibHeight", .LibHeight
    cINI.WriteKey "UseLibrary", .UseLibrary
    cINI.WriteKey "ButtonDefault", .ButtonDefault
    cINI.WriteKey "Docked", Docked
    cINI.WriteKey "DockedLeft", DockedLeft
    cINI.WriteKey "DockedTop", DockedTop
    cINI.WriteKey "SnapWidth", SnapWidth
    cINI.WriteKey "NoPresetFade", .NoPresetFade
    
    'scope settings
    With ScopeSettings
      cINI.sSection = "VisScope"
      cINI.WriteKey "BrushSizeL", .bBrushSizeL
      cINI.WriteKey "BrushSizeR", .bBrushSizeR
      cINI.WriteKey "Detail", .bDetail
      cINI.WriteKey "Fade", .bFade
      cINI.WriteKey "Fall", .bFall
      cINI.WriteKey "PeakCount", .bPeakCount
      cINI.WriteKey "PeakDetail", .bPeakDetail
      cINI.WriteKey "Peaks", .bPeaks
      cINI.WriteKey "Skip", .bSkip
      cINI.WriteKey "Type", .bType
      cINI.WriteKey "ColorL", .lColorL
      cINI.WriteKey "ColorPeakL", .lColorPeakL
      cINI.WriteKey "ColorPeakR", .lColorPeakR
      cINI.WriteKey "ColorR", .lColorR
      cINI.WriteKey "PeakDec", .lPeakDec
      cINI.WriteKey "PeakPause", .lPeakPause
    End With
    
    'spectrum settings
    With SpectrumSettings
      cINI.sSection = "VisSpectrum"
      cINI.WriteKey "Correction", .bCorrection
      cINI.WriteKey "BarSize", .bBarSize
      cINI.WriteKey "BrushSize", .bBrushSize
      cINI.WriteKey "DrawStyle", .bDrawStyle
      cINI.WriteKey "Fade", .bFade
      cINI.WriteKey "Fall", .bFall
      cINI.WriteKey "PeakFall", .bPeakFall
      cINI.WriteKey "Peaks", .bPeaks
      cINI.WriteKey "Type", .bType
      cINI.WriteKey "View", .iView
      cINI.WriteKey "ColorDn", .lColorDn
      cINI.WriteKey "ColorLine", .lColorLine
      cINI.WriteKey "ColorUp", .lColorUp
      cINI.WriteKey "Pause", .lPause
      cINI.WriteKey "PeakColor", .lPeakColor
      cINI.WriteKey "PeakDec", .lPeakDec
      cINI.WriteKey "PeakPause", .lPeakPause
      cINI.WriteKey "Dec", .nDec
      cINI.WriteKey "Zoom", .nZoom
    End With
    
    'volume settings
    With VolumeSettings
      cINI.sSection = "VisVolume"
      cINI.WriteKey "DrawStyle", .bDrawStyle
      cINI.WriteKey "Fade", .bFade
      cINI.WriteKey "Fall", .bFall
      cINI.WriteKey "Type", .bType
      cINI.WriteKey "ColorDn", .lColorDn
      cINI.WriteKey "ColorUp", .lColorUp
      cINI.WriteKey "Pause", .lPause
      cINI.WriteKey "Dec", .nDec
      cINI.WriteKey "File", .sFile
    End With
    
    'beat settings
    With BeatSettings
      cINI.sSection = "VisBeat"
      cINI.WriteKey "Fade", .bFade
      cINI.WriteKey "Type", .bType
      cINI.WriteKey "DetectHigh", .iDetectHigh
      cINI.WriteKey "DetectLow", .iDetectLow
      cINI.WriteKey "Min", .nMin
      cINI.WriteKey "Multi", .nMulti
      cINI.WriteKey "RotMin", .nRotMin
      cINI.WriteKey "RotMove", .nRotMove
      cINI.WriteKey "RotSpeed", .nRotSpeed
      cINI.WriteKey "File", .sFile
    End With
    
    'Equalizer
    cINI.sSection = "Equalizer"
    cINI.WriteKey "EQon", .EQon
    For X = 0 To 9
      cINI.WriteKey "EQ" & X, EQValue(X)
    Next
    
    'Visualization options
    cINI.sSection = "Visualization"
    cINI.WriteKey "UpdateInterval", VisUpdateInt
    
    'File browser filter
    cINI.sSection = "Filter"
    For X = 0 To UBound(.Filter)
      cINI.WriteKey "Filter" & X, .Filter(X)
    Next X
    
    'file type association
    cINI.sSection = "Associate"
    For X = 0 To UBound(.AssType)
      cINI.WriteKey "Type" & X, .AssType(X)
    Next X
    cINI.WriteKey "OnStart", .AssOnStart
    cINI.WriteKey "Action", .AssAction
    
    'Sound device settings
    cINI.sSection = "Device"
    cINI.WriteKey "DXFXon", .DXFXon
    cINI.WriteKey "Type", devType
    cINI.WriteKey "Device", devDevice
    cINI.WriteKey "Mixer", devMixer
    cINI.WriteKey "Frequency", devFreq
    cINI.WriteKey "Channels", devChannels
    cINI.WriteKey "Buffer", devBuffer
    cINI.WriteKey "Panning", devPanning
    cINI.WriteKey "Surround", devSurround
    cINI.WriteKey "Speaker", devSpeaker
    cINI.WriteKey "PanSep", devPanSep
    cINI.WriteKey "DSP", devDSP
 
  End With
  
  cLog.Log "DONE. (" & cLog.GetTimer & " ms)", 2
  
  Exit Sub
ErrHandler:
  If cLog.ErrorMsg(Err, "modMain, SaveSettings()") = vbYes Then Resume Next Else frmMain.UnloadAll
End Sub

Public Sub SimpleAddDir(ByVal sFldr As String)
  'This adds all song in dir & subdirs
  Dim File As New clsFind
  Dim X As Long
  
  On Error GoTo ErrHandler
  
  sFldr = sAppend(sFldr, "\")
  File.Find sFldr, "*.mp2;*.mp3;*.ogg;*.wma;*.asf;*.wav;*.mod;*.it;*.s3m;*.xm;*.midi;*.mid;*.rmi;*.sgm", True
  
  frmMain.MousePointer = vbHourglass
  If Settings.PlaylistOn Then frmPlaylist.MousePointer = vbHourglass
  
  If File.Count > 0 Then
    For X = 1 To File.Count
      CreateLibrary File(X).sFilename
    Next X
  End If
 
  frmMain.MousePointer = vbDefault
  If Settings.PlaylistOn Then frmPlaylist.MousePointer = vbDefault
 
  Exit Sub
ErrHandler:
  If cLog.ErrorMsg(Err, "modMain, SimpleAddDir") = vbYes Then Resume Next Else frmMain.UnloadAll
End Sub

Public Sub SimpleAddFile(ByVal sFile As String)
  'This is used when an associated file is started
  'it will load the file
  On Error GoTo ErrHandler
  Dim X As Long
  
  If Not FileExists(sFile) Then Exit Sub
  
  If Settings.AssAction = 0 Then frmMenus.menDeleteAll_Click 'clear playlist
  
  X = UBound(Playlist)
  
  Select Case LCase(sFilename(sFile, efpFileExt))
    Case "playlist", "m3u", "pls"
      HandlePlaylist sFile
    Case Else
      CreateLibrary sFile
  End Select
  
  If Settings.AssAction = 0 Then
    UpdateList
    If UBound(Playlist) > 0 Then frmMain.PlayNext
  Else
    UpdateList
    If Settings.AssAction = 1 Then
      If Not Sound.StreamIsPlaying And Not Sound.MusicIsPlaying Then
        Playing = X + 1
        ShuffleNum = ShuffleNum + 1
        frmMain.Play
      End If
    ElseIf Settings.AssAction = 2 Then
      Playing = X + 1
      ShuffleNum = ShuffleNum + 1
      frmMain.Play
    End If
  End If
  

  Exit Sub
ErrHandler:
  If cLog.ErrorMsg(Err, "modMain, SimpleAddFile") = vbYes Then Resume Next Else frmMain.UnloadAll
End Sub

Public Sub AssociateTypes()
  'This associates supported filetypes with simple amp
  On Error GoTo ErrHandler
  
  cLog.Log "ASSOCIATING FILE TYPES...", 2, False
  
  Call CreateFileType("SimpleAmp.Stream", "Simple Amp Streamed Media File", App.Path & "\Simple Amp.exe,1")
  Call CreateFileTypeAction("SimpleAmp.Stream", "Open", Chr(34) & App.Path & "\Simple Amp.exe" & Chr(34) & " -open %1")
  Call CreateFileType("SimpleAmp.Module", "Simple Amp Module Media File", App.Path & "\Simple Amp.exe,2")
  Call CreateFileTypeAction("SimpleAmp.Module", "Open", Chr(34) & App.Path & "\Simple Amp.exe" & Chr(34) & " -open %1")
  Call CreateFileType("SimpleAmp.Sequence", "Simple Amp Sequenced Media File", App.Path & "\Simple Amp.exe,3")
  Call CreateFileTypeAction("SimpleAmp.Sequence", "Open", Chr(34) & App.Path & "\Simple Amp.exe" & Chr(34) & " -open %1")
  Call CreateFileType("SimpleAmp.Plist", "Simple Amp Playlist", App.Path & "\Simple Amp.exe,4")
  Call CreateFileTypeAction("SimpleAmp.Plist", "Open", Chr(34) & App.Path & "\Simple Amp.exe" & Chr(34) & " -open %1")
  
  If Settings.AssType(0) Then
    If Not CheckAssociation(".mp3", "SimpleAmp.Stream") Then CreateAssociation ".mp3", "SimpleAmp_Old", "SimpleAmp.Stream"
    If Not CheckAssociation(".mp2", "SimpleAmp.Stream") Then CreateAssociation ".mp2", "SimpleAmp_Old", "SimpleAmp.Stream"
  Else
    RestoreAssociation ".mp3", "SimpleAmp_Old"
    RestoreAssociation ".mp2", "SimpleAmp_Old"
  End If
  If Settings.AssType(1) Then
    If Not CheckAssociation(".ogg", "SimpleAmp.Stream") Then CreateAssociation ".ogg", "SimpleAmp_Old", "SimpleAmp.Stream"
  Else
    RestoreAssociation ".ogg", "SimpleAmp_Old"
  End If
  If Settings.AssType(2) Then
    If Not CheckAssociation(".wma", "SimpleAmp.Stream") Then CreateAssociation ".wma", "SimpleAmp_Old", "SimpleAmp.Stream"
  Else
    RestoreAssociation ".wma", "SimpleAmp_Old"
  End If
  If Settings.AssType(3) Then
    If Not CheckAssociation(".asf", "SimpleAmp.Stream") Then CreateAssociation ".asf", "SimpleAmp_Old", "SimpleAmp.Stream"
  Else
    RestoreAssociation ".asf", "SimpleAmp_Old"
  End If
  If Settings.AssType(4) Then
    If Not CheckAssociation(".wav", "SimpleAmp.Stream") Then CreateAssociation ".wav", "SimpleAmp_Old", "SimpleAmp.Stream"
  Else
    RestoreAssociation ".wav", "SimpleAmp_Old"
  End If
  If Settings.AssType(5) Then
    If Not CheckAssociation(".mod", "SimpleAmp.Module") Then CreateAssociation ".mod", "SimpleAmp_Old", "SimpleAmp.Module"
  Else
    RestoreAssociation ".mod", "SimpleAmp_Old"
  End If
  If Settings.AssType(6) Then
    If Not CheckAssociation(".s3m", "SimpleAmp.Module") Then CreateAssociation ".s3m", "SimpleAmp_Old", "SimpleAmp.Module"
  Else
    RestoreAssociation ".s3m", "SimpleAmp_Old"
  End If
  If Settings.AssType(7) Then
    If Not CheckAssociation(".xm", "SimpleAmp.Module") Then CreateAssociation ".xm", "SimpleAmp_Old", "SimpleAmp.Module"
  Else
    RestoreAssociation ".xm", "SimpleAmp_Old"
  End If
  If Settings.AssType(8) Then
    If Not CheckAssociation(".it", "SimpleAmp.Module") Then CreateAssociation ".it", "SimpleAmp_Old", "SimpleAmp.Module"
  Else
    RestoreAssociation ".it", "SimpleAmp_Old"
  End If
  If Settings.AssType(9) Then
    If Not CheckAssociation(".mid", "SimpleAmp.Sequence") Then CreateAssociation ".mid", "SimpleAmp_Old", "SimpleAmp.Sequence"
    If Not CheckAssociation(".midi", "SimpleAmp.Sequence") Then CreateAssociation ".midi", "SimpleAmp_Old", "SimpleAmp.Sequence"
    If Not CheckAssociation(".rmi", "SimpleAmp.Sequence") Then CreateAssociation ".rmi", "SimpleAmp_Old", "SimpleAmp.Sequence"
  Else
    RestoreAssociation ".mid", "SimpleAmp_Old"
    RestoreAssociation ".rmi", "SimpleAmp_Old"
    RestoreAssociation ".midi", "SimpleAmp_Old"
  End If
  If Settings.AssType(10) Then
    If Not CheckAssociation(".sgm", "SimpleAmp.Sequence") Then CreateAssociation ".sgm", "SimpleAmp_Old", "SimpleAmp.Sequence"
  Else
    RestoreAssociation ".sgm", "SimpleAmp_Old"
  End If
  If Settings.AssType(11) Then
    If Not CheckAssociation(".playlist", "SimpleAmp.Plist") Then CreateAssociation ".playlist", "SimpleAmp_Old", "SimpleAmp.Plist"
  Else
    RestoreAssociation ".playlist", "SimpleAmp_Old"
  End If
  If Settings.AssType(12) Then
    If Not CheckAssociation(".m3u", "SimpleAmp.Plist") Then CreateAssociation ".m3u", "SimpleAmp_Old", "SimpleAmp.Plist"
    If Not CheckAssociation(".pls", "SimpleAmp.Plist") Then CreateAssociation ".pls", "SimpleAmp_Old", "SimpleAmp.Plist"
  Else
    RestoreAssociation ".m3u", "SimpleAmp_Old"
    RestoreAssociation ".pls", "SimpleAmp_Old"
  End If
  
  cLog.Log "DONE.", 2
  
  Exit Sub
ErrHandler:
  If cLog.ErrorMsg(Err, "modMain, AssociateTypes") = vbYes Then Resume Next Else frmMain.UnloadAll
End Sub

Public Sub UpdateList()
  'This will update the playlist and all its associated info
  On Error GoTo ErrHandler
  Dim TotLength As Long, X As Long
  
  'Calculate total playlist length
  'would probably be better with a global var
  'that gets changed for each added file instead.
  For X = 1 To UBound(Playlist)
    TotLength = TotLength + Library(Playlist(X).Reference).lLength
  Next
  
  'Updates form
  frmMain.UpdatePlaylistColor
  With frmPlaylist
    .List.Refresh
    .lblTotalTime = ConvertTime(TotLength)
    .lblTotalNum = UBound(Playlist) & " files."
    If UBound(Playlist) = 0 Then .Scroll.Value = 0
    .Scroll.Max = .List.Max
  End With
    
  Exit Sub
ErrHandler:
  If cLog.ErrorMsg(Err, "modMain, UpdateList") = vbYes Then Resume Next Else frmMain.UnloadAll
End Sub

Public Sub UpdateEQ()
  'Updates EQ values
  On Error Resume Next
  
  If Settings.EQon And Settings.DXFXon Then
    Call FSOUND_FX_SetParamEQ(EQHandle(0), 80, 18, EQValue(0))
    Call FSOUND_FX_SetParamEQ(EQHandle(1), 170, 18, EQValue(1))
    Call FSOUND_FX_SetParamEQ(EQHandle(2), 310, 18, EQValue(2))
    Call FSOUND_FX_SetParamEQ(EQHandle(3), 600, 18, EQValue(3))
    Call FSOUND_FX_SetParamEQ(EQHandle(4), 1000, 18, EQValue(4))
    Call FSOUND_FX_SetParamEQ(EQHandle(5), 3000, 18, EQValue(5))
    Call FSOUND_FX_SetParamEQ(EQHandle(6), 6000, 18, EQValue(6))
    Call FSOUND_FX_SetParamEQ(EQHandle(7), 12000, 18, EQValue(7))
    Call FSOUND_FX_SetParamEQ(EQHandle(8), 14000, 18, EQValue(8))
    Call FSOUND_FX_SetParamEQ(EQHandle(9), 16000, 18, EQValue(9))
  End If
  
End Sub

Public Sub UpdateFX()
  'This sub pauses the sound, closes every FX and then reactivates every FX, then unpauses.
  'Doing this any other way seems to work bad... :(
  On Error GoTo errh
  Dim X As Integer, lc As Integer
  
  If Not Settings.DXFXon Then Exit Sub
  
  With frmStudio
  
    'Pause
    Call FSOUND_SetPaused(FSOUND_SYSTEMCHANNEL, True)
    
    'Close everything
    For X = 0 To 9
      If EQHandle(X) <> 0 Then
        Call FSOUND_FX_Disable(EQHandle(X))
        EQHandle(X) = 0
      End If
    Next X
    For X = 0 To 7
      If DXFXHandle(X) <> 0 Then
        Call FSOUND_FX_Disable(DXFXHandle(X))
        DXFXHandle(X) = 0
      End If
    Next X
  
    'enable everything that is on
    If Settings.EQon Then
      lc = lc + 10
      For X = 0 To 9
        EQHandle(X) = FSOUND_FX_Enable(FSOUND_SYSTEMCHANNEL, FSOUND_FX_PARAMEQ)
      Next X
    End If
    
    If .chkEffectOn(0).Value = vbChecked Then
      lc = lc + 1
      DXFXHandle(0) = FSOUND_FX_Enable(FSOUND_SYSTEMCHANNEL, FSOUND_FX_CHORUS)
    End If
    If .chkEffectOn(1).Value = vbChecked Then
      lc = lc + 1
      DXFXHandle(1) = FSOUND_FX_Enable(FSOUND_SYSTEMCHANNEL, FSOUND_FX_COMPRESSOR)
    End If
    If .chkEffectOn(2).Value = vbChecked Then
      lc = lc + 1
      DXFXHandle(2) = FSOUND_FX_Enable(FSOUND_SYSTEMCHANNEL, FSOUND_FX_DISTORTION)
    End If
    If .chkEffectOn(3).Value = vbChecked Then
      lc = lc + 1
      DXFXHandle(3) = FSOUND_FX_Enable(FSOUND_SYSTEMCHANNEL, FSOUND_FX_ECHO)
    End If
    If .chkEffectOn(4).Value = vbChecked Then
      lc = lc + 1
      DXFXHandle(4) = FSOUND_FX_Enable(FSOUND_SYSTEMCHANNEL, FSOUND_FX_FLANGER)
    End If
    If .chkEffectOn(5).Value = vbChecked Then
      lc = lc + 1
      DXFXHandle(5) = FSOUND_FX_Enable(FSOUND_SYSTEMCHANNEL, FSOUND_FX_GARGLE)
    End If
    If .chkEffectOn(6).Value = vbChecked Then
      lc = lc + 1
      DXFXHandle(6) = FSOUND_FX_Enable(FSOUND_SYSTEMCHANNEL, FSOUND_FX_I3DL2REVERB)
    End If
    If .chkEffectOn(7).Value = vbChecked Then
      lc = lc + 1
      DXFXHandle(7) = FSOUND_FX_Enable(FSOUND_SYSTEMCHANNEL, FSOUND_FX_WAVES_REVERB)
    End If
    
    'Update values of everything
    .UpdateCompressor
    .UpdateDistortion
    .UpdateChorus
    .UpdateEcho
    .UpdateFlanger
    .UpdateGargle
    .UpdateI3DL2
    .UpdateWave
    UpdateEQ
  
  End With
  
  'Unpause
  Call FSOUND_SetPaused(FSOUND_SYSTEMCHANNEL, False)
  
  If lc <= 16 Then Exit Sub
errh:
  If lc > 16 Then
    MsgBox "DirectX can only support 16 effects at the same time. You have selected more than that, so some effects will not be heard. The equalizer counts as 10 effects, so it's easiest to just turn it off.", vbExclamation, "Too many effects"
  Else
    If cLog.ErrorMsg(Err, "modMain, UpdateFx()") = vbYes Then Resume Next Else frmMain.UnloadAll
  End If
End Sub

Function ScopeCallback(ByVal originalbuffer As Long, ByVal newbuffer As Long, ByVal length As Long, ByVal param As Long) As Long
  'This callback takes note of the pointer to the sound buffer
  RealtimeBuffer = newbuffer
  
  ScopeCallback = newbuffer
End Function

Public Sub SaveLibrary()
  'This sub saves the media library
  On Error GoTo ErrHandler
  Dim File As New clsDatafile
  Dim X As Long, v As enumFileType
  
  If UBound(LibraryIndex) > 0 Then
    
    cLog.Log "SAVING MEDIA LIBRARY DATABASE...", 5, False
    cLog.StartTimer
    
    If FileExists(App.Path & "\database.db_") Then Kill App.Path & "\database.db_"
    File.FileName = App.Path & "\database.db_"
    File.WriteStrFixed "SADB"
    For X = 1 To UBound(LibraryIndex)
      If LibraryIndex(X).lReference > 0 Then
        'the file info has been loaded from library/updated
        'so write from memory to new database
        With Library(LibraryIndex(X).lReference)
          LibraryIndex(X).lPointer = File.Position
          File.WriteNumber .eType
          File.WriteDate .dLastUpdateDate
          File.WriteNumber .lLength
          If .eType = TYPE_MP2_MP3 Or .eType = TYPE_OGG Then
            'Write following values only if mp3,mp3,ogg
            File.WriteNumber Abs(.bIsVBR)
            File.WriteStr .sAlbum
            File.WriteStr .sArtist
            File.WriteStr .sComments
            File.WriteStr .sGenre
            File.WriteStr .sTitle
            File.WriteStr .sTrack
            File.WriteStr .sYear
          ElseIf .eType = TYPE_IT Or .eType = TYPE_MOD Or .eType = TYPE_S3M Or .eType = TYPE_XM Then
            'only write title on modules
            File.WriteStr .sTitle
          End If
        End With
      Else
        'it has not been loaded/updated...
        If LibraryIndex(X).lPointer > 0 Then
          '...but it exists in file, so read from old database
          'and write into new database
          fLibrary.Position = LibraryIndex(X).lPointer
          v = fLibrary.ReadNumber
          LibraryIndex(X).lPointer = File.Position
          File.WriteNumber v
          File.WriteDate fLibrary.ReadDate
          File.WriteNumber fLibrary.ReadNumber
          If v = TYPE_MP2_MP3 Or v = TYPE_OGG Then
            'Write following values only if mp3,mp3,ogg
            File.WriteNumber fLibrary.ReadNumber
            File.WriteStr fLibrary.ReadStr
            File.WriteStr fLibrary.ReadStr
            File.WriteStr fLibrary.ReadStr
            File.WriteStr fLibrary.ReadStr
            File.WriteStr fLibrary.ReadStr
            File.WriteStr fLibrary.ReadStr
            File.WriteStr fLibrary.ReadStr
          ElseIf v = TYPE_IT Or v = TYPE_MOD Or v = TYPE_S3M Or v = TYPE_XM Then
            'only write title on modules
            File.WriteStr fLibrary.ReadStr
          End If
        Else
          '...and it does not exist in file, so write
          'empty data into database
          With Library(LibraryIndex(X).lReference)
            LibraryIndex(X).lPointer = File.Position
            File.WriteNumber .eType
            File.WriteDate .dLastUpdateDate
            File.WriteNumber .lLength
            If .eType = TYPE_MP2_MP3 Or .eType = TYPE_OGG Then
              'Write following values only if mp3,mp3,ogg
              File.WriteNumber Abs(.bIsVBR)
              File.WriteStr .sAlbum
              File.WriteStr .sArtist
              File.WriteStr .sComments
              File.WriteStr .sGenre
              File.WriteStr .sTitle
              File.WriteStr .sTrack
              File.WriteStr .sYear
            ElseIf .eType = TYPE_IT Or .eType = TYPE_MOD Or .eType = TYPE_S3M Or .eType = TYPE_XM Then
              'only write title on modules
              File.WriteStr .sTitle
            End If
          End With
        End If
      End If
    Next

    cLog.Log "DONE. (" & File.Position - 1 & " bytes, " & cLog.GetTimer & " ms)", 5
    cLog.Log "SAVING MEDIA LIBRARY INDEX...", 5, False
    cLog.StartTimer
    
    Dim tStr As String
    If FileExists(App.Path & "\index.db_") Then Kill App.Path & "\index.db_"
    File.FileName = App.Path & "\index.db_"
    File.WriteStrFixed "SADBINDEX"
    File.WriteNumber UBound(LibraryIndex)
    For X = 1 To UBound(LibraryIndex)
      With LibraryIndex(X)
        If sFilename(.sFilename, efpFilePath) = tStr Then
          File.WriteStr "\" & sFilename(.sFilename, efpFileNameAndExt)
        Else
          File.WriteStr .sFilename
          tStr = sFilename(.sFilename, efpFilePath)
        End If
        File.WriteNumber .lPointer
      End With
    Next

    cLog.Log "DONE. (" & File.Position - 1 & " bytes, " & cLog.GetTimer & " ms)", 5
    
    Set File = Nothing
    Set fLibrary = Nothing 'close original file
    
    If FileExists(App.Path & "\index.db") Then Kill App.Path & "\index.db"
    If FileExists(App.Path & "\database.db") Then Kill App.Path & "\database.db"
    Name App.Path & "\index.db_" As App.Path & "\index.db"
    Name App.Path & "\database.db_" As App.Path & "\database.db"

  End If
   
  Exit Sub
ErrHandler:
  MsgBox "There was an error while saving the media library (database.db, index.db). The files might be write protected.", vbCritical, "Write Error"
  If FileExists(App.Path & "\index.db_") Then Kill App.Path & "\index.db_"
  If FileExists(App.Path & "\database.db_") Then Kill App.Path & "\database.db_"
End Sub

Public Function LibraryCheck(ByVal sFilename As String) As Long
  'This function checks the media library after the filename
  'and returns its index, or 0 is there is none.
  On Error Resume Next
  Dim X As Long, y As Long
  
  sFilename = LCase(sFilename)
  y = Len(sFilename)
  
  'Search thought library
  For X = 1 To UBound(LibraryIndex)
    If LibraryIndex(X).lLen = y Then 'used to speed up
      If LCase(LibraryIndex(X).sFilename) = sFilename Then
        LibraryCheck = X
        Exit For
      End If
    End If
  Next

End Function

Public Function CreateLibrary(ByVal sFilename As String, Optional AddPlaylist As Boolean = True) As Long
  'This function creates/gets sFilename in the library and adds it
  'to the playlist array.
  On Error Resume Next
  Dim Item As Long

  'check if file exists in index library
  Item = LibraryCheck(sFilename)
  'if it doesn't...
  If Item = 0 Then
 
    'check if module is valid
    Select Case LCase(modMisc.sFilename(sFilename, efpFileExt))
      Case "it", "mod", "xm", "s3m"
        If Not Sound.GetMusicOK(sFilename) Then
          'module was not OK, so return 0
          Exit Function
        End If
    End Select
    
    'Everything is ok, create new item in index library
    ReDim Preserve LibraryIndex(UBound(LibraryIndex) + 1)
    With LibraryIndex(UBound(LibraryIndex))
      .lPointer = 0
      .sFilename = sFilename
      .lLen = Len(.sFilename)
      .lReference = UBound(Library) + 1 'the item we are going to create
    End With
    
    'create new item in media library
    ReDim Preserve Library(UBound(Library) + 1)
    Item = UBound(Library)
    Library(Item).sFilename = sFilename
    UpdateLibrary Item
  Else
    'This file exists in index library, so...
    If LibraryIndex(Item).lReference = 0 Then '...if it has not yet been loaded
      'create the new item in the media library
      ReDim Preserve Library(UBound(Library) + 1)
      LibraryIndex(Item).lReference = UBound(Library)
      'but the item exists in the library (in file), so load it
      LoadItem UBound(Library), Item
      Item = UBound(Library) 'and set up the reference
    Else '...if it has been loaded
      Item = LibraryIndex(Item).lReference 'just set up the reference
    End If
  End If
  
  If AddPlaylist Then
    ReDim Preserve Playlist(UBound(Playlist) + 1) As PlaylistData
    Playlist(UBound(Playlist)).Reference = Item
    Playlist(UBound(Playlist)).index = UBound(Playlist)
  End If
  
  CreateLibrary = Item
  
End Function

Public Sub UpdateLibrary(ByVal Item As Long)
  'This sub updates the info of item if it is outdated.
  'Doing this check every time an file is loaded into the playlist is
  'too slow, so instead the check is made whenever the item is played.
  On Error Resume Next
  Dim File As New clsFile
  Dim hFile As Long
  
  If Not FileExists(Library(Item).sFilename) Then Exit Sub

  If LibraryCheck(Library(Item).sFilename) = 0 Then
    CreateLibrary Library(Item).sFilename, False
    Exit Sub
  End If
  
  'Now, update file info if file has been changed since last time.
  'Newly created items will always be updated.
  With Library(Item)
    File.sFilename = .sFilename
      
    'If the latest noted modified date is not the same, update file
    If .dLastUpdateDate <> File.dLastWriteTime Then
        
        .bIsVBR = False
        .lLength = 0
        .sAlbum = ""
        .sArtist = ""
        .sArtistTitle = ""
        .sComments = ""
        .sGenre = ""
        .sTitle = ""
        .sTrack = ""
        .sYear = ""
       
        LibraryChanged = True
        .dLastUpdateDate = File.dLastWriteTime
        
        'Check which file type this is
        Select Case LCase(File.sExtension)
          Case "mp2", "mp3"
           
            .eType = TYPE_MP2_MP3

            hFile = Sound.GetStreamOpenFile(.sFilename, .bIsVBR)
            'Read ID3v2 tags
            .sTitle = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "TIT2")
            .sArtist = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "TPE1")
            .sAlbum = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "TALB")
            .sGenre = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "TCON")
            .sYear = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "TYER")
            .sTrack = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "TRCK")
            .sComments = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "COMM")
            'Read ID3v1 tags if there where no ID3v2 tag
            If Len(.sTitle) = 0 Then
              .sTitle = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V1, "TITLE")
            End If
            If Len(.sArtist) = 0 Then
              .sArtist = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V1, "ARTIST")
            End If
            If Len(.sAlbum) = 0 Then
              .sAlbum = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V1, "ALBUM")
            End If
            If Len(.sGenre) = 0 Then
              .sGenre = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V1, "GENRE")
              If Len(.sGenre) > 0 Then
                If Val(.sGenre) <= UBound(Genre) Then
                  .sGenre = Genre(Val(.sGenre))
                Else
                  .sGenre = ""
                End If
              End If
            Else 'format Id3v2 genre (remove number)
              If Left(.sGenre, 1) = "(" Then
                .sGenre = Right(.sGenre, Len(.sGenre) - InStr(1, .sGenre, ")"))
              End If
            End If
            If Len(.sYear) = 0 Then
              .sYear = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V1, "YEAR")
            End If
            If Len(.sTrack) = 0 Then
              .sTrack = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V1, "TRACK")
            End If
            If Len(.sComments) = 0 Then
              .sComments = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V1, "COMMENT")
            End If
            .lLength = Sound.GetStreamLength(hFile) 'get length
            Sound.GetStreamCloseFile hFile 'close file
                       
            .bIsVBR = GetMp3VBR(.sFilename)
            If Len(.sTitle) > 0 And Len(.sArtist) > 0 Then
              .sArtistTitle = StrFormat(.sArtist & " - " & .sTitle)
            Else
              .sArtistTitle = StrFormat(sFilename(.sFilename, efpFileNameAndExt))
            End If
            
          Case "ogg"
            
            .eType = TYPE_OGG
            
            'open file
            hFile = Sound.GetStreamOpenFile(.sFilename, False)
            'get vorbis tags
            .sTitle = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_VORBISCOMMENT, "TITLE")
            .sArtist = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_VORBISCOMMENT, "ARTIST")
            .sAlbum = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_VORBISCOMMENT, "ALBUM")
            .sYear = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_VORBISCOMMENT, "DATE")
            .sGenre = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_VORBISCOMMENT, "GENRE")
            .sTrack = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_VORBISCOMMENT, "TRACKNUMBER")
            .sComments = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_VORBISCOMMENT, "COMMENT")
            .lLength = Sound.GetStreamLength(hFile) 'Get length
            Sound.GetStreamCloseFile hFile 'close file
            
            If Len(.sTitle) > 0 And Len(.sArtist) > 0 Then
              .sArtistTitle = StrFormat(.sArtist & " - " & .sTitle)
            Else
              .sArtistTitle = StrFormat(sFilename(.sFilename, efpFileNameAndExt))
            End If
            
          Case "wma"
            .eType = TYPE_WMA
            
            'open file
            hFile = Sound.GetStreamOpenFile(.sFilename, False)
            'get ASF tags
            .sTitle = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ASF, "TITLE")
            .sArtist = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ASF, "AUTHOR")
            .lLength = Sound.GetStreamLength(hFile) 'Get length
            Sound.GetStreamCloseFile hFile 'close file
            
            If Len(.sTitle) > 0 And Len(.sArtist) > 0 Then
              .sArtistTitle = StrFormat(.sArtist & " - " & .sTitle)
            Else
              .sArtistTitle = StrFormat(sFilename(.sFilename, efpFileNameAndExt))
            End If

          Case "asf"
            .eType = TYPE_ASF
            
            'open file
            hFile = Sound.GetStreamOpenFile(.sFilename, False)
            'get ASF tags
            .sTitle = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ASF, "TITLE")
            .sArtist = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ASF, "AUTHOR")
            .lLength = Sound.GetStreamLength(hFile) 'Get length
            Sound.GetStreamCloseFile hFile 'close file
            
            If Len(.sTitle) > 0 And Len(.sArtist) > 0 Then
              .sArtistTitle = StrFormat(.sArtist & " - " & .sTitle)
            Else
              .sArtistTitle = StrFormat(sFilename(.sFilename, efpFileNameAndExt))
            End If
            
          Case "wav"
            .eType = TYPE_WAV
            
            hFile = Sound.GetStreamOpenFile(.sFilename, False) 'open file
            .lLength = Sound.GetStreamLength(hFile) 'Get length
            Sound.GetStreamCloseFile hFile 'close file

            .sArtistTitle = StrFormat(sFilename(.sFilename, efpFileNameAndExt))
            
          Case "mod"
            .eType = TYPE_MOD
            Sound.GetMusicData .sFilename, .lLength, .sTitle
            .sArtistTitle = StrFormat(sFilename(.sFilename, efpFileNameAndExt))
            
          Case "it"
            .eType = TYPE_IT
            Sound.GetMusicData .sFilename, .lLength, .sTitle
            .sArtistTitle = StrFormat(sFilename(.sFilename, efpFileNameAndExt))
            
          Case "xm"
            .eType = TYPE_XM
            Sound.GetMusicData .sFilename, .lLength, .sTitle
            .sArtistTitle = StrFormat(sFilename(.sFilename, efpFileNameAndExt))
            
          Case "s3m"
            .eType = TYPE_S3M
            Sound.GetMusicData .sFilename, .lLength, .sTitle
            .sArtistTitle = StrFormat(sFilename(.sFilename, efpFileNameAndExt))
            
          Case "mid", "rmi", "midi"
            .eType = TYPE_MID_RMI
            
            hFile = Sound.GetStreamOpenFile(.sFilename, False)
            .lLength = Sound.GetStreamLength(hFile) 'Get length
            Sound.GetStreamCloseFile hFile 'close file

            .sArtistTitle = StrFormat(sFilename(.sFilename, efpFileNameAndExt))
            
          Case "sgm"
            .eType = TYPE_SGM
            
            hFile = Sound.GetStreamOpenFile(.sFilename, False)
            .lLength = Sound.GetStreamLength(hFile) 'Get length
            Sound.GetStreamCloseFile hFile 'close file
            
            .sArtistTitle = StrFormat(sFilename(.sFilename, efpFileNameAndExt))
            
          Case Else
            Select Case Sound.TestFile(.sFilename)
              Case 1
                'this unknown stream can be opened, we assume it is an mp3
                'just in case we check for all sorts of tags
                
                .eType = TYPE_MP2_MP3
                
                hFile = Sound.GetStreamOpenFile(.sFilename, .bIsVBR)
                'Read ID3v2 tags
                .sTitle = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "TIT2")
                .sArtist = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "TPE1")
                .sAlbum = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "TALB")
                .sGenre = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "TCON")
                .sYear = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "TYER")
                .sTrack = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "TRCK")
                .sComments = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "COMM")
                'Read ID3v1 tags if there where no ID3v2 tag
                If Len(.sTitle) = 0 Then
                  .sTitle = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V1, "TITLE")
                End If
                If Len(.sArtist) = 0 Then
                  .sArtist = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V1, "ARTIST")
                End If
                If Len(.sAlbum) = 0 Then
                  .sAlbum = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V1, "ALBUM")
                End If
                If Len(.sGenre) = 0 Then
                  .sGenre = Genre(Val(Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V1, "GENRE")))
                Else 'format Id3v2 genre (remove number)
                  If Left(.sGenre, 1) = "(" Then
                    .sGenre = Right(.sGenre, Len(.sGenre) - InStr(1, .sGenre, ")"))
                  End If
                End If
                If Len(.sYear) = 0 Then
                  .sYear = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V1, "YEAR")
                End If
                If Len(.sTrack) = 0 Then
                  .sTrack = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V1, "TRACK")
                End If
                If Len(.sComments) = 0 Then
                  .sComments = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V1, "COMMENT")
                End If
                'Read vorbis tags
                If Len(.sTitle) = 0 Then
                  .sTitle = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_VORBISCOMMENT, "TITLE")
                End If
                If Len(.sArtist) = 0 Then
                  .sArtist = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_VORBISCOMMENT, "ARTIST")
                End If
                If Len(.sAlbum) = 0 Then
                  .sAlbum = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_VORBISCOMMENT, "ALBUM")
                End If
                If Len(.sGenre) = 0 Then
                  .sGenre = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_VORBISCOMMENT, "GENRE")
                End If
                If Len(.sYear) = 0 Then
                  .sYear = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_VORBISCOMMENT, "DATE")
                End If
                If Len(.sTrack) = 0 Then
                  .sTrack = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_VORBISCOMMENT, "TRACKNUMBER")
                End If
                If Len(.sComments) = 0 Then
                  .sComments = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_VORBISCOMMENT, "COMMENT")
                End If
                'Read ASF tags
                If Len(.sTitle) = 0 Then
                  .sTitle = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ASF, "TITLE")
                End If
                If Len(.sArtist) = 0 Then
                  .sArtist = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ASF, "AUTHOR")
                End If
                .lLength = Sound.GetStreamLength(hFile) 'get length
                Sound.GetStreamCloseFile hFile 'close file
                
                .bIsVBR = GetMp3VBR(.sFilename)
                If Len(.sTitle) > 0 And Len(.sArtist) > 0 Then
                  .sArtistTitle = StrFormat(.sArtist & " - " & .sTitle)
                Else
                  .sArtistTitle = StrFormat(sFilename(.sFilename, efpFileNameAndExt))
                End If
                
              Case 2
                .eType = TYPE_MOD
                Sound.GetMusicData .sFilename, .lLength, .sTitle
                .sArtistTitle = StrFormat(sFilename(.sFilename, efpFileNameAndExt))
              Case Else
                MsgBox "The file """ & .sFilename & """ is not valid! This error may result in further errors. ", vbCritical
            End Select
        End Select
      
    End If
  End With
  
End Sub

Public Sub HandlePlaylist(ByVal sFilename As String)
  'If these items are playlists, load them via other subroutines
  On Error Resume Next
  Select Case LCase(modMisc.sFilename(sFilename, efpFileExt))
    Case "playlist"
      If Not LoadPlaylist(sFilename) Then
        MsgBox "Loading """ & sFilename & """ was unsuccessful. The file might be corrupt, or worse!", vbCritical, "Load Error"
        Exit Sub
      End If
      
    Case "pls"
      If Not LoadPls(sFilename) Then
        MsgBox "Loading """ & sFilename & """ was unsuccessful. The file might be corrupt, or worse!", vbCritical, "Load Error"
        Exit Sub
      End If
      
    Case "m3u"
      If Not LoadM3u(sFilename) Then
        MsgBox "Loading """ & sFilename & """ was unsuccessful. The file might be corrupt, or worse!", vbCritical, "Load Error"
        Exit Sub
      End If
  End Select
End Sub

Public Sub SaveStat()
  'Saves statistics to stat.db
  On Error GoTo ErrHandler
  Dim File As New clsDatafile
  Dim X As Long
  
  If UBound(LibraryIndex) > 0 Then
    
    cLog.Log "SAVING MEDIA LIBRARY STATISTICS...", 5, False
    cLog.StartTimer
        
    Set fStat = Nothing
    
    File.FileName = App.Path & "\stat.db"
    File.WriteStrFixed "SADBSTAT"
    For X = 1 To UBound(LibraryIndex)
      If LibraryIndex(X).lReference > 0 Then
        File.Position = ((X - 1) * 20) + 8 'index * size of each index + header size
        With Library(LibraryIndex(X).lReference)
          File.WriteDate .dLastPlayDate
          File.WriteLong .lTimesPlayed
          File.WriteLong .lTimesSkipped
        End With
      End If
    Next
  
    cLog.Log "DONE. (" & File.Position - 1 & " bytes, " & cLog.GetTimer & " ms)", 5
    Set File = Nothing
    
  End If

  Exit Sub
ErrHandler:
  MsgBox "There was an error while saving the statistics database (stat.db). The file might be write protected.", vbCritical, "Write Error"
End Sub

Public Sub LoadIndex()
  'This sub load the media library index
  On Error GoTo ErrHandler
  Dim File As New clsDatafile
  Dim X As Long, tStr As String
  
  If FileExists(App.Path & "\index.db") Then
    cLog.Log "LOADING MEDIA LIBRARY INDEX...", 5, False
    cLog.StartTimer
  
    File.FileName = App.Path & "\index.db"
    If File.ReadStrFixed(9) <> "SADBINDEX" Then GoTo ErrHandler
    ReDim LibraryIndex(File.ReadNumber)
    For X = 1 To UBound(LibraryIndex)
      With LibraryIndex(X)
        .sFilename = File.ReadStr
        If Left(.sFilename, 1) = "\" And Left(.sFilename, 2) <> "\\" Then
          .sFilename = tStr & Right(.sFilename, Len(.sFilename) - 1)
        Else
          tStr = sFilename(.sFilename, efpFilePath)
        End If
        .lLen = Len(.sFilename)
        .lPointer = File.ReadNumber
        .lReference = 0
      End With
    Next
    
    cLog.Log "DONE. (" & File.Position - 1 & " bytes, " & cLog.GetTimer & " ms)", 5
  End If

  Exit Sub
ErrHandler:
  MsgBox "There was an error while loading the media library index (index.db). Simple Amp can continue but all of the items contained in the media library will be lost.", vbCritical, "Read Error"
End Sub

Public Sub LoadItem(ByVal libindex As Long, ByVal index As Long)
  'This loads media library information into index from file position pos
  On Error Resume Next
  
  LoadedMedia = LoadedMedia + 1
  
  fLibrary.Position = LibraryIndex(index).lPointer
  With Library(libindex)
    .sFilename = LibraryIndex(index).sFilename
    .eType = fLibrary.ReadNumber
    .dLastUpdateDate = fLibrary.ReadDate
    .lLength = fLibrary.ReadNumber
    If .eType = TYPE_MP2_MP3 Or .eType = TYPE_OGG Then
      'read following values only if mp3,mp3,ogg
      .bIsVBR = CBool(fLibrary.ReadNumber)
      .sAlbum = fLibrary.ReadStr
      .sArtist = fLibrary.ReadStr
      .sComments = fLibrary.ReadStr
      .sGenre = fLibrary.ReadStr
      .sTitle = fLibrary.ReadStr
      .sTrack = fLibrary.ReadStr
      .sYear = fLibrary.ReadStr
    ElseIf .eType = TYPE_IT Or .eType = TYPE_MOD Or .eType = TYPE_S3M Or .eType = TYPE_XM Then
      'only read title on modules
      .sTitle = fLibrary.ReadStr
    End If
    If Len(.sArtist) > 0 And Len(.sTitle) > 0 Then
      .sArtistTitle = StrFormat(.sArtist & " - " & .sTitle)
    Else
      .sArtistTitle = StrFormat(sFilename(.sFilename, efpFileNameAndExt))
    End If
    'This reads items stat settings from stat.db by calculating its location from index
    If FileLen(App.Path & "\stat.db") >= ((index - 1) * 20) + 8 Then
      fStat.Position = ((index - 1) * 20) + 8 'index * size of each index + header size
      .dLastPlayDate = fStat.ReadDate
      .lTimesPlayed = fStat.ReadNumber
      .lTimesSkipped = fStat.ReadNumber
    End If
  End With
  
End Sub

'Function StreamEndCallback(ByVal stream As Long, ByVal buff As Long, ByVal Length As Long, ByVal param As Long) As Long
  'This simple callback monitors the end of an stream.
  'This does not work on ASF/WMA...
'  If Sound.StreamIsLoaded And Settings.Advance And Not ManStopped Then
'    If Settings.Repeat = 2 Then
'      frmMain.Play
'    Else
'      frmMain.PlayStop
'      frmMain.PlayNext
'    End If
'  End If
'End Function

Public Sub AddCDAudio()
  On Error Resume Next
  Dim X As Long, y As Long
  X = Sound.CDGetNumTracks
  If X > 0 Then
    For y = 1 To X
      ReDim Preserve Library(UBound(Library) + 1)
      With Library(UBound(Library))
        .sFilename = "//CD Audio Track " & y
        .eType = TYPE_CDA
        .lLength = Sound.CDGetTrackLength(y) \ 1000
        .sArtistTitle = "CD Track " & y
      End With
      ReDim Preserve Playlist(UBound(Playlist) + 1)
      With Playlist(UBound(Playlist))
        .Reference = UBound(Library)
        .index = UBound(Playlist)
      End With
    Next
    UpdateList
  Else
    MsgBox "No tracks found in the default CD drive.", vbInformation
  End If
End Sub

Public Function GetColor(ByVal Color As enumSkinColors) As Long
  GetColor = SkinColor(Color)
End Function

Public Function GetImage(ByVal image As enumSkinImgComponent) As Picture
  Set GetImage = Images(SkinImage(image))
End Function

Public Sub SavePreset(ByVal sName As String, ByVal lNum As Long)
  'This saves an visualization preset
  On Error GoTo ErrHandler
  Dim cFile As New clsDatafile
  
  sName = sAppend(sName, ".sap")
  If FileExists(sName) Then Kill sName
  
  cFile.FileName = sName
  cFile.WriteStrFixed "SAPRESET"
  cFile.WriteNumber PRESETVERSION
  cFile.WriteNumber lNum
  Select Case lNum
    Case 1
      With ScopeSettings
        cFile.WriteNumber .bBrushSizeL
        cFile.WriteNumber .bBrushSizeR
        cFile.WriteNumber .bDetail
        cFile.WriteNumber .bFade
        cFile.WriteNumber .bFall
        cFile.WriteNumber .bPeakCount
        cFile.WriteNumber .bPeakDetail
        cFile.WriteNumber .bPeaks
        cFile.WriteNumber .bSkip
        cFile.WriteNumber .bType
        cFile.WriteNumber .lColorL
        cFile.WriteNumber .lColorPeakL
        cFile.WriteNumber .lColorPeakR
        cFile.WriteNumber .lColorR
        cFile.WriteNumber .lPeakDec
        cFile.WriteNumber .lPeakPause
      End With
    Case 2
      With SpectrumSettings
        cFile.WriteNumber .bBarSize
        cFile.WriteNumber .bBrushSize
        cFile.WriteNumber .bDrawStyle
        cFile.WriteNumber .bFade
        cFile.WriteNumber .bFall
        cFile.WriteNumber .bPeakFall
        cFile.WriteNumber .bPeaks
        cFile.WriteNumber .bType
        cFile.WriteNumber .iView
        cFile.WriteNumber .lColorDn
        cFile.WriteNumber .lColorLine
        cFile.WriteNumber .lColorUp
        cFile.WriteNumber .lPause
        cFile.WriteNumber .lPeakColor
        cFile.WriteNumber .lPeakDec
        cFile.WriteNumber .lPeakPause
        cFile.WriteNumber .nDec
        cFile.WriteNumber .nZoom * 100
        cFile.WriteNumber .bCorrection
      End With
    Case 3
      With VolumeSettings
        cFile.WriteNumber .bDrawStyle
        cFile.WriteNumber .bFade
        cFile.WriteNumber .bFall
        cFile.WriteNumber .bType
        cFile.WriteNumber .lColorDn
        cFile.WriteNumber .lColorUp
        cFile.WriteNumber .lPause
        cFile.WriteNumber .nDec
        cFile.WriteStr .sFile
      End With
    Case 4
      With BeatSettings
        cFile.WriteNumber .bFade
        cFile.WriteNumber .bType
        cFile.WriteNumber .iDetectHigh
        cFile.WriteNumber .iDetectLow
        cFile.WriteNumber .nMin * 1000
        cFile.WriteNumber .nMulti * 1000
        cFile.WriteNumber .nRotMin * 1000
        cFile.WriteNumber .nRotMove * 1000
        cFile.WriteNumber .nRotSpeed * 1000
        cFile.WriteStr .sFile
      End With
  End Select
  
  cLog.Log "SAVED VISUALIZATION PRESET (" & sName & ")", 2
  
  Exit Sub
ErrHandler:
  MsgBox "Could not properly save to """ & sName & """. This must be because of your complete lack of talent.", vbExclamation, "Saving Problems"
End Sub

Public Sub LoadPreset(ByVal sName As String, Optional bNoVisual As Boolean = False)
  'This loads an visualization preset into current settings
  On Error GoTo ErrHandler
  Dim cFile As New clsDatafile
  
  lCurVisPreset = 0
  
  cFile.FileName = sName
  If cFile.ReadStrFixed(8) <> "SAPRESET" Then Err.Raise 3223
  If cFile.ReadNumber <> PRESETVERSION Then Err.Raise 3223
  Spectrum = cFile.ReadNumber
  Select Case Spectrum
    Case 1
      With ScopeSettings
        .bBrushSizeL = cFile.ReadNumber
        .bBrushSizeR = cFile.ReadNumber
        .bDetail = cFile.ReadNumber
        .bFade = cFile.ReadNumber
        .bFall = cFile.ReadNumber
        .bPeakCount = cFile.ReadNumber
        .bPeakDetail = cFile.ReadNumber
        .bPeaks = cFile.ReadNumber
        .bSkip = cFile.ReadNumber
        .bType = cFile.ReadNumber
        If Not bNoVisual Then
          .lColorL = cFile.ReadNumber
          .lColorPeakL = cFile.ReadNumber
          .lColorPeakR = cFile.ReadNumber
          .lColorR = cFile.ReadNumber
        Else
          cFile.SkipField 4
        End If
        .lPeakDec = cFile.ReadNumber
        .lPeakPause = cFile.ReadNumber
        If Settings.NoPresetFade Then .bFade = 0
      End With
      
      DeleteObject hPenRight
      hPenRight = CreatePen(0, ScopeSettings.bBrushSizeR, ScopeSettings.lColorR)
      hBrushSolidRight = CreateSolidBrush(ScopeSettings.lColorR)
      DeleteObject hPenLeft
      hPenLeft = CreatePen(0, ScopeSettings.bBrushSizeL, ScopeSettings.lColorL)
      hBrushSolidLeft = CreateSolidBrush(ScopeSettings.lColorL)
      
      DeleteObject hPenPeakRight
      hPenPeakRight = CreatePen(0, 1, ScopeSettings.lColorPeakR)
      DeleteObject hPenPeakLeft
      hPenPeakLeft = CreatePen(0, 1, ScopeSettings.lColorPeakL)

    Case 2
      With SpectrumSettings
        .bBarSize = cFile.ReadNumber
        .bBrushSize = cFile.ReadNumber
        .bDrawStyle = cFile.ReadNumber
        .bFade = cFile.ReadNumber
        .bFall = cFile.ReadNumber
        .bPeakFall = cFile.ReadNumber
        .bPeaks = cFile.ReadNumber
        .bType = cFile.ReadNumber
        .iView = cFile.ReadNumber
        If Not bNoVisual Then
          .lColorDn = cFile.ReadNumber
          .lColorLine = cFile.ReadNumber
          .lColorUp = cFile.ReadNumber
        Else
          cFile.SkipField 3
        End If
        .lPause = cFile.ReadNumber
        If Not bNoVisual Then
          .lPeakColor = cFile.ReadNumber
        Else
          cFile.SkipField
        End If
        .lPeakDec = cFile.ReadNumber
        .lPeakPause = cFile.ReadNumber
        .nDec = cFile.ReadNumber
        .nZoom = cFile.ReadNumber / 100
        .bCorrection = cFile.ReadNumber
        If Settings.NoPresetFade Then .bFade = 0
      End With
      
      frmMain.DoGrad SpectrumSettings.lColorUp, SpectrumSettings.lColorDn
      cGradBar.Create 1, frmMain.picVis.ScaleHeight, frmMain.hdc
      cGradBar.BitBltFrom frmMain.p.hdc
      frmMain.DoGrad SpectrumSettings.lColorUp, SpectrumSettings.lColorDn
      cGradBar.Create 1, frmMain.picVis.ScaleHeight, frmMain.hdc
      cGradBar.BitBltFrom frmMain.p.hdc
      
      DeleteObject hPenLineSpec
      hPenLineSpec = CreatePen(0, SpectrumSettings.bBrushSize, SpectrumSettings.lColorLine)
      DeleteObject hPenPeakSpec
      hPenPeakSpec = CreatePen(0, 1, SpectrumSettings.lPeakColor)
      
    Case 3
      With VolumeSettings
        .bDrawStyle = cFile.ReadNumber
        .bFade = cFile.ReadNumber
        .bFall = cFile.ReadNumber
        .bType = cFile.ReadNumber
        If Not bNoVisual Then
          .lColorDn = cFile.ReadNumber
          .lColorUp = cFile.ReadNumber
        Else
          cFile.SkipField 2
        End If
        .lPause = cFile.ReadNumber
        .nDec = cFile.ReadNumber
        .sFile = cFile.ReadStr
        .lImage = 0
        If Settings.NoPresetFade Then .bFade = 0
      End With
      
      frmMain.DoGrad VolumeSettings.lColorUp, VolumeSettings.lColorDn
      cGradVol.Create 2, frmMain.picVis.ScaleHeight, frmMain.hdc
      cGradVol.BitBltFrom frmMain.p.hdc
      
    Case 4
      With BeatSettings
        .bFade = cFile.ReadNumber
        .bType = cFile.ReadNumber
        .iDetectHigh = cFile.ReadNumber
        .iDetectLow = cFile.ReadNumber
        .nMin = cFile.ReadNumber / 1000
        .nMulti = cFile.ReadNumber / 1000
        .nRotMin = cFile.ReadNumber / 1000
        .nRotMove = cFile.ReadNumber / 1000
        .nRotSpeed = cFile.ReadNumber / 1000
        .sFile = cFile.ReadStr
        .lImage = 0
        If Settings.NoPresetFade Then .bFade = 0
      End With
    Case Else
      Err.Raise 3223 'whatever, just error
  End Select

  cLog.Log "LOADED VISUALIZATION PRESET (" & sName & ")", 2

  frmMain.UpdateSpectrum

  Exit Sub
ErrHandler:
  MsgBox "This preset could not be loaded, because of an fatal error inside it's main core binary structure.", vbExclamation, "Loading Error"
End Sub

Public Sub SaveEQ(ByVal sName As String)
  'saves an equalizer preset
  On Error GoTo errh
  Dim cFile As New clsDatafile
  Dim X As Long
  
  sName = sAppend(sName, ".seq")
  If FileExists(sName) Then Kill sName
  
  With cFile
    .FileName = sName
    .WriteStrFixed "SAEQ"
    .WriteNumber EQVERSION
    For X = 0 To 9
      .WriteNumber EQValue(X)
    Next
  End With
  
  cLog.Log "SAVED EQUALIZER PRESET (" & sName & ")", 2
  
  Exit Sub
errh:
  MsgBox "Could not save the equalizer preset.", vbExclamation, "Writing Error"
End Sub

Public Sub LoadEQ(ByVal sName As String)
  'loads an equalizer preset
  On Error GoTo errh
  Dim cFile As New clsDatafile
  Dim X As Long
  
  With cFile
    .FileName = sName
    If .ReadStrFixed(4) = "SAEQ" And .ReadNumber = EQVERSION Then
      For X = 0 To 9
        EQValue(X) = .ReadNumber
      Next
    Else
      GoTo errh
    End If
  End With
  
  cLog.Log "LOADED EQUALIZER PRESET (" & sName & ")", 2
  
  Exit Sub
errh:
  MsgBox "Could not load the equalizer preset.", vbExclamation, "Reading Error"
End Sub

Public Sub ShowPopup(ByRef frm As Form, ByRef menu As menu, ByRef obj As Object, ByVal lWidth As Long)
  'this sub will align an popupmenu left or right depending on if it fits on screen
  'menu is the menu to display as popup
  'obj is the control that the menu is aligned to
  'lwidth is the width of the menu (use TextWidth on longest caption in menu + ~45 pixels)
  On Error Resume Next
  RefreshScreenRECT
  Dim lLeft As Long
  
  lLeft = frm.Left
  
  If frm.ScaleMode = vbTwips Then
    Scr.Right = Scr.Right * Screen.TwipsPerPixelX
  Else
    lLeft = lLeft / Screen.TwipsPerPixelX
  End If
  
  'Debug.Print lLeft, obj.Left, lWidth
  
  If lLeft + obj.Left + lWidth > Scr.Right Then
    frm.PopupMenu menu, vbPopupMenuRightAlign, obj.Left + obj.Width, obj.Top + obj.Height
  Else
    frm.PopupMenu menu, vbPopupMenuLeftAlign, obj.Left, obj.Top + obj.Height
  End If
  
End Sub

Public Sub RefreshScreenRECT()
  SystemParametersInfo SPI_GETWORKAREA, 0&, Scr, 0&
End Sub

Public Sub LoadSkinPreset(ByVal lNumPreset As Long)
  On Error Resume Next

  If lNumPreset > UBound(SkinPresets) Then lNumPreset = 1

  Spectrum = SkinPresets(lNumPreset).bType
  lCurVisPreset = lNumPreset
  Select Case SkinPresets(lNumPreset).bType
    Case 1
      ScopeSettings = SkinPresets(lNumPreset).t1
            
      DeleteObject hPenRight
      hPenRight = CreatePen(0, ScopeSettings.bBrushSizeR, ScopeSettings.lColorR)
      hBrushSolidRight = CreateSolidBrush(ScopeSettings.lColorR)
      DeleteObject hPenLeft
      hPenLeft = CreatePen(0, ScopeSettings.bBrushSizeL, ScopeSettings.lColorL)
      hBrushSolidLeft = CreateSolidBrush(ScopeSettings.lColorL)
      
      DeleteObject hPenPeakRight
      hPenPeakRight = CreatePen(0, 1, ScopeSettings.lColorPeakR)
      DeleteObject hPenPeakLeft
      hPenPeakLeft = CreatePen(0, 1, ScopeSettings.lColorPeakL)
    Case 2
      SpectrumSettings = SkinPresets(lNumPreset).t2
      
      frmMain.DoGrad SpectrumSettings.lColorUp, SpectrumSettings.lColorDn
      cGradBar.Create 1, frmMain.picVis.ScaleHeight, frmMain.hdc
      cGradBar.BitBltFrom frmMain.p.hdc
      frmMain.DoGrad SpectrumSettings.lColorUp, SpectrumSettings.lColorDn
      cGradBar.Create 1, frmMain.picVis.ScaleHeight, frmMain.hdc
      cGradBar.BitBltFrom frmMain.p.hdc
      
      DeleteObject hPenLineSpec
      hPenLineSpec = CreatePen(0, SpectrumSettings.bBrushSize, SpectrumSettings.lColorLine)
      DeleteObject hPenPeakSpec
      hPenPeakSpec = CreatePen(0, 1, SpectrumSettings.lPeakColor)
    Case 3
      VolumeSettings = SkinPresets(lNumPreset).t3
      
      frmMain.DoGrad VolumeSettings.lColorUp, VolumeSettings.lColorDn
      cGradVol.Create 2, frmMain.picVis.ScaleHeight, frmMain.hdc
      cGradVol.BitBltFrom frmMain.p.hdc
      
    Case 4
      BeatSettings = SkinPresets(lNumPreset).t4
  End Select
  frmMain.UpdateSpectrum
  
End Sub

