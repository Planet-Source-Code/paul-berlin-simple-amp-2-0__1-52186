Attribute VB_Name = "modVariables"
Option Explicit

Public Type POINTAPI
  X As Long
  y As Long
End Type

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

'File types
'----------
'Used to identify which file type the file is, so the right action can be
'performed on it when for example it is to be played.
Public Enum enumFileType
  TYPE_UNKNOWN = 0
  TYPE_MP2_MP3 = 1
  TYPE_WAV = 2
  TYPE_OGG = 3
  TYPE_WMA = 4
  TYPE_ASF = 5
  TYPE_MOD = 6
  TYPE_S3M = 7
  TYPE_XM = 8
  TYPE_IT = 9
  TYPE_MID_RMI = 10
  TYPE_SGM = 11
  TYPE_PLAYLIST = 12
  TYPE_CDA = 13
End Enum

'All images used in the current skin
'-----------------------------------
'This is used when accessing each image of a skin from memory
'for easy programming.
Public Enum enumSkinImgComponent
  Preview = 0 'Not loaded!
  SI_MAIN_BG = 1
  SI_PLIST_BG = 2
  SI_PLIST_COLUMNS = 3
  SI_MAIN_STEREO = 4
  SI_MAIN_MONO = 5
  SI_PLIST2_BG = 6
  SI_PLIST2_COLUMNS = 7
  SCR_MAIN_POS_AFTER = 8
  SCR_MAIN_POS_BEFORE = 9
  SCR_MAIN_POS_BAR = 10
  SCR_MAIN_VOL_AFTER = 11
  SCR_MAIN_VOL_BEFORE = 12
  SCR_MAIN_VOL_BAR = 13
  SCR_PLIST_SCRL_AFTER = 14
  SCR_PLIST_SCRL_BEFORE = 15
  SCR_PLIST_SCRL_BAR = 16
  SCR_PLIST2_SCRL_AFTER = 17
  SCR_PLIST2_SCRL_BEFORE = 18
  SCR_PLIST2_SCRL_BAR = 19
  BTN_PLIST_ADD_DOWN = 20
  BTN_PLIST_ADD_UP = 21
  BTN_PLIST_ADD_UPM = 22
  BTN_PLIST_REM_DOWN = 23
  BTN_PLIST_REM_UP = 24
  BTN_PLIST_REM_UPM = 25
  BTN_PLIST_SEL_DOWN = 26
  BTN_PLIST_SEL_UP = 27
  BTN_PLIST_SEL_UPM = 28
  BTN_PLIST_LST_DOWN = 29
  BTN_PLIST_LST_UP = 30
  BTN_PLIST_LST_UPM = 31
  BTN_MAIN_PREV_DOWN = 32
  BTN_MAIN_PREV_UP = 33
  BTN_MAIN_PREV_UPM = 34
  BTN_MAIN_NEXT_DOWN = 35
  BTN_MAIN_NEXT_UP = 36
  BTN_MAIN_NEXT_UPM = 37
  BTN_MAIN_PLAY_DOWN = 38
  BTN_MAIN_PLAY_UP = 39
  BTN_MAIN_PLAY_UPM = 40
  BTN_MAIN_PAUSE_DOWN = 41
  BTN_MAIN_PAUSE_UP = 42
  BTN_MAIN_PAUSE_UPM = 43
  BTN_MAIN_STOP_DOWN = 44
  BTN_MAIN_STOP_UP = 45
  BTN_MAIN_STOP_UPM = 46
  BTN_MAIN_CLOSE_DOWN = 47
  BTN_MAIN_CLOSE_UP = 48
  BTN_MAIN_CLOSE_UPM = 49
  BTN_MAIN_MIN_DOWN = 50
  BTN_MAIN_MIN_UP = 51
  BTN_MAIN_MIN_UPM = 52
  BTN_PLIST_CLOSE_DOWN = 53
  BTN_PLIST_CLOSE_UP = 54
  BTN_PLIST_CLOSE_UPM = 55
  BTN_PLIST_SIZE_DOWN = 56
  BTN_PLIST_SIZE_UP = 57
  BTN_PLIST_SIZE_UPM = 58
  BTN_PLIST2_ADD_DOWN = 59
  BTN_PLIST2_ADD_UP = 60
  BTN_PLIST2_ADD_UPM = 61
  BTN_PLIST2_REM_DOWN = 62
  BTN_PLIST2_REM_UP = 63
  BTN_PLIST2_REM_UPM = 64
  BTN_PLIST2_SEL_DOWN = 65
  BTN_PLIST2_SEL_UP = 66
  BTN_PLIST2_SEL_UPM = 67
  BTN_PLIST2_LST_DOWN = 68
  BTN_PLIST2_LST_UP = 69
  BTN_PLIST2_LST_UPM = 70
  BTN_PLIST2_CLOSE_DOWN = 71
  BTN_PLIST2_CLOSE_UP = 72
  BTN_PLIST2_CLOSE_UPM = 73
  BTN_PLIST2_SIZE_DOWN = 74
  BTN_PLIST2_SIZE_UP = 75
  BTN_PLIST2_SIZE_UPM = 76
  XBTN_MAIN_PLIST_DOWN = 77
  XBTN_MAIN_PLIST_OFFUP = 78
  XBTN_MAIN_PLIST_OFFUPM = 79
  XBTN_MAIN_PLIST_ONUP = 80
  XBTN_MAIN_PLIST_ONUPM = 81
  XBTN_MAIN_RPT_DOWN = 82
  XBTN_MAIN_RPT_OFFUP = 83
  XBTN_MAIN_RPT_OFFUPM = 84
  XBTN_MAIN_RPT_ONUP = 85
  XBTN_MAIN_RPT_ONUPM = 86
  XBTN_MAIN_SHFL_DOWN = 87
  XBTN_MAIN_SHFL_OFFUP = 88
  XBTN_MAIN_SHFL_OFFUPM = 89
  XBTN_MAIN_SHFL_ONUP = 90
  XBTN_MAIN_SHFL_ONUPM = 91
  SPEC_BAR = 92
  SPEC_BG1 = 93
  SPEC_BG2 = 94
  SPEC_BG3 = 95
  SPEC_BG4 = 96
  SPEC_BGOFF = 97
  SCR_MAIN_POS_BARDRAG = 98
  SCR_MAIN_VOL_BARDRAG = 99
  SCR_PLIST_SCRL_BARDRAG = 100
  SCR_PLIST2_SCRL_BARDRAG = 101
  PLIST_LST_BG = 102
  PLIST2_LST_BG = 103
  SPEC_BG5 = 104
  SPEC_BEAT = 105
End Enum

'All colors used in the current skin
'-----------------------------------
'used when accessing color values for each component in the skin
Public Enum enumSkinColors
  TXT_MAIN_ARTISTTITLE_EN = 1
  TXT_MAIN_ARTISTTITLE_DIS = 0
  TXT_MAIN_ALBUM_EN = 3
  TXT_MAIN_ALBUM_DIS = 2
  TXT_MAIN_GENRE_EN = 5
  TXT_MAIN_GENRE_DIS = 4
  TXT_MAIN_YEAR_EN = 7
  TXT_MAIN_YEAR_DIS = 6
  TXT_MAIN_COM_EN = 9
  TXT_MAIN_COM_DIS = 8
  TXT_MAIN_INF_EN = 11
  TXT_MAIN_INF_DIS = 10
  TXT_MAIN_TIME_EN = 13
  TXT_MAIN_TIME_DIS = 12
  TXT_MAIN_TOTTIME_EN = 15
  TXT_MAIN_TOTTIME_DIS = 14
  TXT_PLIST_NUM_EN = 17
  TXT_PLIST_NUM_DIS = 16
  TXT_PLIST_TIME_EN = 19
  TXT_PLIST_TIME_DIS = 18
  TXT_PLIST2_NUM_EN = 21
  TXT_PLIST2_NUM_DIS = 20
  TXT_PLIST2_TIME_EN = 23
  TXT_PLIST2_TIME_DIS = 22
  LST_PLIST_BG = 24
  LST_PLIST_FONT = 25
  LST_PLIST2_BG = 26
  LST_PLIST2_FONT = 27
'  SPEC_COL1 = 28
'  SPEC_COL2 = 29
'  SPEC_COL3 = 30
'  SPEC_COL4 = 31
  LST_PLIST_SELECT = 28
  LST_PLIST2_SELECT = 29
'  SPEC_COL1B = 34
End Enum

'For storing X, Y position of an skin component as well as Width & Height.
Type XYHW
  X As Long
  y As Long
  H As Long
  W As Long
End Type

'The stored data for skin component, text label
Type TXT
  Font As String    'Font name
  Size As Long      'Font size
  Bold As Boolean   'Is Bold?
  Italic As Boolean 'IS Italic?
  Align As Byte     'Alignment
  Pos As XYHW       'Position & Size
End Type

'The stored data for skin component, listbox
Type tLIST
  Font As String 'Font Name
  FontSize As Integer 'Font size
  Pos As XYHW    'Position & Size
  ColAT As Long  'Width of Artist&Title Column
  ColA As Long   'Width of Album Column
  ColG As Long   'Width of Genre Column
  ColT As Long   'Width of Time Column
End Type

'The stored data for skin component, playlist windows
Type PlaylistWindow
  Self As XYHW      'The window itselfs size (position is not used)
  Trans As Boolean  'Has this window transparent areas?
  iColumn As XYHW   'The position & size of Column image
  bAdd As XYHW      'The position & size of Add button
  bRemove As XYHW   'The position & size of Remove Button
  bSelect As XYHW   'The position & size of Select button
  bList As XYHW     'The position & size of List button
  bClose As XYHW    'The position & size of Close button
  bSize As XYHW     'The position & size of Size button
  sScroll As XYHW   'The position & size of Scrollbar
  num As TXT        'The text data for total number label
  Time As TXT       'The text data for total time label
  List As tLIST     'The data for the list
  HasBg As Boolean  'Does the list have an background image?
End Type

'This is the Media Library Type.
'-------------------------------
'The Library Array using this type is an database of information of all the
'music files you have opened in simple amp, if you have Media library enabled,
'ever. If you have Media library off, this contains all files since program start.
'Using the media library, one file does not have to be loaded more than one.
Public Type tMediaLibrary
  '-File data------------ (Saved in Database.db)
  sFilename       As String 'Filename of indexed media file
  dLastUpdateDate As Date   'Last time file was changed (so we can se when we need to update)
  eType           As enumFileType 'Type of file this was identified as
  '-Music data----------- (Saved in Database.db)
  bIsVBR          As Byte 'Is the mp3 an VBR? (MP3/MP2)
  lLength         As Long   'Length of file in seconds
  sTitle          As String 'Title of file (ID3/Vorbis Tag/Module Title)
  sArtist         As String 'Artist of file (ID3/Vorbis Tag)
  sAlbum          As String 'Album of file (ID3/Vorbis Tag)
  sGenre          As String 'Genre of file (ID3/Vorbis Tag)
  sYear           As String 'Year of file (ID3/Vorbis Tag)
  sComments       As String 'Comments of file (ID3/Vorbis Tag)
  sTrack          As String 'Track number (ID3/Vorbis Tag)
  '-Statistics----------- (Saved in Stat.db)
  dLastPlayDate   As Date   'Last time played in Simple Amp
  lTimesPlayed    As Long   'Number of times played in Simple Amp
  lTimesSkipped   As Long   'Number of times this song was skipped in Simple Amp
  '-Not Saved-----------
  sArtistTitle    As String 'Temporary. 'Artist - Title'/filename/module name
End Type

'Media Index
'-----------
'This is used to keep track of which media libraries have been loaded into memory
Public Type tMediaIndex
  sFilename     As String 'Filename
  lLen          As Long   'Length of sFilename, used to speed up LibraryCheck
  lPointer      As Long   'Location at which the files data is saved in database.db in bytes
  lReference    As Long   'Reference to Library Array, 0 if file info has not been loaded
End Type

Public Type PlaylistData
  Reference   As Long    'This is an reference to the
                         'item in the media library to take data from
  lShuffleIndex As Long  'Shuffle index (for ordered mode)
  '-Used by CtrlLister--
  Removed     As Byte    'true when item is removed (to speed up)
  Index       As Long    'order of item in playlist (so it can be moved)
  IsBold      As Byte    'Is text bold?
  Selected    As Byte    'Is selected?
End Type

'The folowing types are used for visualizations
'----------------------------------------------

'Settings for the Oscilliscope
Type tVisScope
  bType As Byte         'Type of scope, 0 = Dot, 1 = Line, 2 = Solid, 3 = History
  bDetail As Byte       'Detail of scope, channels, 0 = Stereo, 1 = Mono
  bSkip As Byte         'Can only be even numbers, 2,4,6,8,10 etc. 2=default
  bBrushSizeL As Byte   'Left Channel Thickness in pixels
  bBrushSizeR As Byte   'Right Channel Thickness in pixels
  lColorL As Long
  lColorR As Long
  'PEAK
  bPeaks As Byte        '0 = none, 1 = dot, 2 = line
  bPeakDetail As Byte   'Detail of peaks, channels, 0 = Stereo, 1 = Mono
  bPeakCount As Byte    'Detail of scope, upper/lower. 0 = upper, 1 = lower, 2 = both
  lPeakPause As Long    'Pause length at top in ms
  lPeakDec As Long      '0 = off, > 0 = how much peak to add to the decrease value each time (creating an faster & faster effect...)
  bFall As Byte         '0 = off, 1 = lower peak faster & faster
  lColorPeakL As Long
  lColorPeakR As Long
  'MISC
  bFade As Byte         'Fade speed, 0 = no fade
End Type

'Settings for the Spectrum Analyzer
Type tVisSpectrum
  bType As Byte          'Type of spectrum, 0 = Normal, 1 = Bars, 2 = Lines
  nZoom As Single        'Zoom value, default 2
  iView As Integer       'Width of view, 511 = full
  bBarSize As Byte       'Width of bars in bType = 1, 4 (Thin) is default
  lColorUp As Long       'Upper gradient color
  lColorDn As Long       'Lower gradient color
  lColorLine As Long     'Line type color
  bDrawStyle As Byte     'Style to draw bars, 0 = normal (crop), 1 = fire style (scale)
  lPause As Long         'Pause length at top in ms
  nDec As Single         '0 = off, > 0 = how to lower bar each interval
  bFall As Byte          '0 = off, 1 = lower peak faster & faster
  bBrushSize As Byte     'Thickness of line in type=2
  bCorrection As Byte    'Correction on bType = 1, def. 0
  'PEAK
  bPeaks As Byte         '0 = none, 1 = on
  lPeakPause As Long     'Pause length at top in ms
  lPeakDec As Long       '0 = off, > 0 = how much to lower the peak each interval
  bPeakFall As Byte      '0 = off, 1 = lower peak faster & faster
  lPeakColor As Long     'Color of peak lines
  'MISC
  bFade As Byte          'Fade speed, 0 = no fade
End Type

'Settings for the Volume meter
Type tVisVolume
  bType As Byte          'Type of volume meter, 0 = 2 bars, 1 = history
  sFile As String        'Image file if any
  lImage As Long         'pointer to skin image to use, if sfile is ""
  lPause As Long         'Pause length at top in ms
  nDec As Single         '0 = off, > 0 = how to lower bar each interval
  bFall As Byte          '0 = off, 1 = lower peak faster & faster
  lColorUp As Long       'Upper gradient color
  lColorDn As Long       'Lower gradient color
  bDrawStyle As Byte     'Style to draw bars, 0 = normal (crop), 1 = fire style (scale)
  'MISC
  bFade As Byte          'Fade speed, 0 = no fade
End Type

'Settings for the Beat Detector
Type tVisBeat
  bType As Byte          '0 = Swinger, 1 = Rotator
  sFile As String        'Image file if any (must be square!)
  lImage As Long         'pointer to skin image to use, if sfile is ""
  iDetectLow As Integer  'Value to begin detection on, 0-511, < iDetectHigh, def. 0
  iDetectHigh As Integer 'Value to end detection on, 0-511, > iDetectLow, def. 25
  nMulti As Single       'Multiply value, def. 2.5
  nMin As Single         'Minimum Zoom value, def. 0.75
  nRotMin As Single      'Minimum zoom value increase to begin rotation at, def 0.25
  nRotMove As Single     'Move speed of roation, in radians per frame update, def 0.05
  nRotSpeed As Single    'Rotation speed in .bType = 1, def. 0.05
  bFade As Byte
End Type

'type used for peak data
Type tPeaks
  Value As Single 'current value
  Time As Long  'time elapsed since last value decrease
  Pause As Long 'time elapsed since paused
  Dec As Single 'Decrease after each time pass
End Type

Type tSkinPresets
  bType As Byte '1-4, t1-t4
  sName As String 'name
  t1 As tVisScope 'only one of t1-t4 used... a waste of memory but it's
  t2 As tVisSpectrum 'easier to program... anyways, skins shouldn't have
  t3 As tVisVolume 'more than 10-15 presets anyway.
  t4 As tVisBeat
End Type

'The following variables are all saved in settings.ini
'-----------------------------------------------------
'Settings type
Type tSettings
  DynamicColumns  As Boolean 'if true, use dynamic size columns in playlist, def. false
  NoPresetFade    As Boolean 'if true, fade is always of when loading presets
  LibHideNonMusic As Boolean 'If true, in the media library window, modules, midi, wave etc. are hidden.
  LibGetName      As Boolean 'get artist & title from filename?
  LibAutosize     As Boolean 'Do autosize in media library?
  LibSearch(4)    As Byte    'Artist/Title/Album/Comments/Filename, on or off
  LibFilter       As Byte    '0/1/2/3 = Artist/Album/Genre/Year
  ButtonDefault   As Boolean 'Do default action when leftclicking playlist buttons?
  UseLibrary      As Boolean 'Use media library?
  DXFXon          As Boolean 'Use DirectX Effects?
  EQon            As Boolean 'Is Equalizer on?
  AddMax          As Boolean 'Is Browse window maximized
  AddWidth        As Long    'Browse window width
  AddHeight       As Long    'Browse window height
  AddFolderWidth  As Long    'Browse window folder view width
  AddAutosize     As Boolean 'Do autosize on columns in add window?
  LibMax          As Boolean 'Is Library window maximized
  LibWidth        As Long    'Library window width
  LibHeight       As Long    'Library window height
  WinTop          As Integer 'Top pos of main window
  WinLeft         As Integer 'Left pos of main window
  PlaylistOn      As Boolean 'Is playlist on?
  PlaylistSmall   As Boolean 'Is playlist small?
  Fade            As Boolean 'Is fade on?
  Advance         As Boolean 'Is auto advance on?
  Repeat          As Byte    'Repeat mode 0=off/1=list/2=song
  Shuffle         As Byte    'Shuffle mode 0=off/1=ordered/2=random
  CurrentVolume   As Byte    'Volume
  TrayIcon        As Byte    'Tray icon to use
  AlwaysTray      As Boolean 'Always show tray icon?
  OnTop           As Boolean 'Always ontop?
  StartInTray     As Boolean 'Start in tray?
  Snap            As Boolean 'Snap to screen edges?
  AddView         As Byte    'Is the list in the browse window extended?
  AddDir          As String  'Last dir of add window
  Filter(11)      As Boolean 'browse filter
  AssType(12)     As Boolean 'associate types
  AssOnStart      As Boolean 'associate type on start?
  AssAction       As Byte    'associate action 0 = open & play, 1 = add to playlist, 2 = add to playlist & play
  BrowseMode      As Byte    'File Browser browse mode, 0 = full always, 1 = full on local drives, 2 = never full
End Type

'General Program varibales
'-------------------------
Public LoadedMedia          As Long      'Number of media library items loaded
Public CurrentSkin          As String    'Currently used skin
Public Playing              As Long      'Currently playing song in the list
Public PlayingLib           As Long      'Currently playing song in the list
Public Scr                  As RECT      'holds screen size
Public CurMono              As Boolean   'True if currently playing music is mono
Public ManStopped           As Boolean   'Used for stream end callback when manual stopped
Public FileOpen             As String    'This holds filename if simple amp was started with -open [filename], until the program is started ok
Public StreamStart          As Boolean   'Used to control music flow
Public LibraryChanged       As Boolean   'Has the media library changed since load?
Public ShuffleNum           As Long      'The current number of shuffles done and which item to play from its lShuffleIndex
Public Docked               As Boolean   'If true, playlist window is docked with main window
Public DockedLeft           As Long      'Difference between main window & playlist window in left pos
Public DockedTop            As Long      'Difference between main window & playlist window in top pos
Public SnapWidth            As Integer   'Width of snap area

Public Genre()              As String    'this array holds ID3v1 genres

'DSP & Visualization variables, some are saved to settings.ini
'-------------------------------------------------------------
'Equalizer & DSP
Public EQValue(9)           As Long    'The current eq values
Public EQHandle(9)          As Long    'The handles to the EQ settings
Public DXFXHandle(7)        As Long    'Handles to DirectX FX settings
Public DSP_OK               As Boolean 'Is DSP started & OK?
Public DSP_Handle           As Long    'Handle to dsp unit
Public RealtimeBuffer       As Long    'Pointer to DSP buffer
Public lCurVisPreset        As Integer 'the current visible skin preset used, used when left-clicking on vis area
Public Spectrum             As Integer '0 = off, 1 = scope, 2 = spectrum, 3 = volume
Public VisUpdateInt         As Integer 'Visualization update interval
'OSCILLISCOPE
Public ScopeSettings        As tVisScope 'Scope settings
Public ScopeBufferINT()     As Integer 'Array with DSP buffer, MMX format
Public ScopeBufferFPU()     As Single  'Array with DSP buffer, CPU/FPU format
Public ScopeUPeaks()        As tPeaks  'Upper peaks values for scope
Public ScopeLPeaks()        As tPeaks  'Lower peaks values for scope
Public ScopeHistoryL()      As Single  'Holds data for history scope, Left Channel
Public ScopeHistoryR()      As Single  'Holds data for history scope, Right Channel
Public hPenLeft             As Long    'Handle to left oscilliscope pen
Public hPenRight            As Long    'Handle to right oscilliscope pen
Public hPenPeakLeft         As Long    'Handle to left oscilliscope peak pen
Public hPenPeakRight        As Long    'Handle to right oscilliscope peak pen
Public hBrushSolidLeft      As Long    'Handle to left solid oscilliscope brush
Public hBrushSolidRight     As Long    'Handle to right solid oscilliscope brush
'SPECTRUM ANALYZER
Public SpectrumSettings     As tVisSpectrum 'Spectrum settings
Public SpectrumPeaks()      As tPeaks  'Holds the peaks info
Public SpectrumBars()       As tPeaks  'Holds the bars info
Public hPenPeakSpec         As Long    'Handle to peak pen
Public hPenLineSpec         As Long    'Handle to type=2 line pen
'VOLUME BARS
Public VolumeSettings       As tVisVolume 'Volume meter settings
Public VolumeBars(1)        As tPeaks  'Holds the values of the two volume bars
Public VolumeHistory()      As Single  'Holds the history values of the volume meter
'BEAT DETECTOR
Public BeatSettings         As tVisBeat 'Beat detector settings
'SKIN PRESETS
Public SkinPresets()        As tSkinPresets

'Engine settings saved in settings.ini
'-------------------------------------
'Variables for device settings
Public devType              As Byte     'Type of output system used, 0 = autodetect
Public devDevice            As Byte     'Number of device used, 0 = default
Public devMixer             As Byte     'Number of mixer used, 0 = autodetect
Public devFreq              As Long     'Sampling frequency
Public devChannels          As Integer  'max channels
Public devBuffer            As Integer  'buffer size
Public devPanning           As Byte     'panning of sound 0=left, 255 = right
Public devSurround          As Boolean  'Surround mode
Public devSpeaker           As Byte     'Speaker setup
Public devPanSep            As Single   'Pan seperation 0.0-1.0
Public devDSP               As Boolean  'This controls if DSP is on or off

'Program constants
'-----------------
Public Const DEFAULTSKIN As String = "Console"  'Default skin name
Public Const COMPILEVERSION As Byte = 107       'The version of skins this version of simple amp can open
Public Const PRESETVERSION As Byte = 101
Public Const EQVERSION As Byte = 100
Public Const LOG_DEBUG As Boolean = True 'logging constants not useful when compiled
Public Const LOG_LEVEL As Integer = 1

Public Const sGenreMatrix = "Blues|Classic Rock|Country|Dance|Disco|Funk|Grunge|" + _
    "Hip-Hop|Jazz|Metal|New Age|Oldies|Other|Pop|R&B|Rap|Reggae|Rock|Techno|" + _
    "Industrial|Alternative|Ska|Death Metal|Pranks|Soundtrack|Euro-Techno|" + _
    "Ambient|Trip Hop|Vocal|Jazz+Funk|Fusion|Trance|Classical|Instrumental|Acid|" + _
    "House|Game|Sound Clip|Gospel|Noise|Alt. Rock|Bass|Soul|Punk|Space|Meditative|" + _
    "Instrumental Pop|Instrumental Rock|Ethnic|Gothic|Darkwave|Techno-Industrial|Electronic|" + _
    "Pop-Folk|Eurodance|Dream|Southern Rock|Comedy|Cult|Gangsta Rap|Top 40|Christian Rap|" + _
    "Pop/Punk|Jungle|Native American|Cabaret|New Wave|Phychedelic|Rave|Showtunes|Trailer|" + _
    "Lo-Fi|Tribal|Acid Punk|Acid Jazz|Polka|Retro|Musical|Rock & Roll|Hard Rock|Folk|" + _
    "Folk/Rock|National Folk|Swing|Fast-Fusion|Bebob|Latin|Revival|Celtic|Blue Grass|" + _
    "Avantegarde|Gothic Rock|Progressive Rock|Psychedelic Rock|Symphonic Rock|Slow Rock|" + _
    "Big Band|Chorus|Easy Listening|Acoustic|Humour|Speech|Chanson|Opera|Chamber Music|" + _
    "Sonata|Symphony|Booty Bass|Primus|Porn Groove|Satire|Slow Jam|Club|Tango|Samba|Folklore|" + _
    "Ballad|Power Ballad|Rhythmic Soul|Freestyle|Duet|Punk Rock|Drum Solo|A Capella|Euro-House|" + _
    "Dance Hall|Goa|Drum & Bass|Club-House|Hardcore|Terror|Indie|Brit Pop|Negerpunk|Polsk Punk|" + _
    "Beat|Christian Gangsta Rap|Heavy Metal|Black Metal|Crossover|Comteporary Christian|" + _
    "Christian Rock|Merengue|Salsa|Trash Metal|Anime|JPop|Synth Pop"

'The sound device controller
'---------------------------
Public Sound                As New clsFMOD 'the main sound control class
                                           'everything except DSP & callbacks

'Holds availabe devices for each output type
Public Direct_Dev()         As String
Public WinOut_Dev()         As String

'Image buffers
Public cBack                As New clsImg 'Visualization backbuffer
Public cBackOrig            As New clsImg 'Background image, 'cleared' state of visualization
'Image buffer holding volume meter bar image
Public cBar                 As New clsImg
'Image buffer holding image used in beat detection visualization
Public cBeat                As New clsImg
Public cBeatTemp            As New clsImg
'Image buffer holding gradient version of spectrum analyzer bars
Public cGradBar             As New clsImg
Public cGradVol             As New clsImg

'Region data (used when setting transparent areas of windows)
'Apparently you must rebuild the region data each time you use it,
'so we have to keep each rectangle to be able to recreate the
'the data later.
Public R1()                 As RECT 'This is an collection of rectangles
Public R2()                 As RECT 'that form the pieces of the window
Public R3()                 As RECT 'that should be transparent.

'Skin data
'---------
Public Images()             As Picture  'Holds all images
Public SkinImage(105)       As Long     'Pointers to correct images
Public SkinColor(29)        As Long     'Holds all color values
Public MainTrans            As Boolean  'Is main window transparent?
Public PlaylistWin(1)       As PlaylistWindow 'holds position, fonts etc of all items in each size
                                 'of playlist window

'The file controls for database.db & stat.db
Public fLibrary             As New clsDatafile
Public fStat                As New clsDatafile

'Settings
Public Settings             As tSettings

'The playlist
Public Playlist()           As PlaylistData
'The Media Library
Public Library()            As tMediaLibrary
Public LibraryIndex()       As tMediaIndex

Public cLog                 As New clsLog 'used to log events & display error messages
