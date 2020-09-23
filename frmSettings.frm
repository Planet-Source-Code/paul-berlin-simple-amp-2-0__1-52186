VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Program Settings"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm 
      Caption         =   "General"
      Height          =   3975
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   4695
      Begin VB.CheckBox chkGeneral 
         Caption         =   "&Dynamic size playlist columns."
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   2535
      End
      Begin VB.CheckBox chkGeneral 
         Caption         =   "&Fade Windows"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CheckBox chkGeneral 
         Caption         =   "&Start minimized to tray."
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox chkGeneral 
         Caption         =   "&Always on top."
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox chkGeneral 
         Caption         =   "Always show &tray icon."
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
      Begin VB.CheckBox chkGeneral 
         Caption         =   "Snap &windows to screen edges."
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1330
         Width           =   2655
      End
      Begin VB.CheckBox chkGeneral 
         Caption         =   "&Playlist buttons do default action when clicked."
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   3735
      End
      Begin MSComctlLib.Slider sldIcon 
         Height          =   375
         Left            =   1200
         TabIndex        =   12
         Top             =   3120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   7
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tray &Icon:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   2640
         Width           =   720
      End
      Begin VB.Image imgIcon 
         Height          =   375
         Left            =   480
         Top             =   3120
         Width           =   375
      End
   End
   Begin VB.Frame frm 
      Caption         =   "Media Library"
      Height          =   3975
      Index           =   3
      Left            =   2520
      TabIndex        =   18
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton cmdClean 
         Caption         =   "Library &Maintenace"
         Height          =   375
         Left            =   2760
         TabIndex        =   23
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   "Media Library Browser Settings"
         Height          =   2295
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   4455
         Begin VB.CheckBox chkLibrary 
            Caption         =   "&When there are no tags, get artist && title from filename."
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   4215
         End
         Begin VB.CheckBox chkLibrary 
            Caption         =   "&Autosize the list columns to fit their content."
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   21
            Top             =   1320
            Width           =   3375
         End
         Begin VB.CheckBox chkLibrary 
            Caption         =   "Hi&de Modules, Midis and Wave-files in the Media Library Browser."
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   22
            Top             =   1680
            Width           =   3855
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmSettings.frx":0000
            Height          =   675
            Left            =   360
            TabIndex        =   25
            Top             =   525
            Width           =   4035
         End
      End
      Begin VB.CheckBox chkLibrary 
         Caption         =   "&Use Media Library"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame frm 
      Caption         =   "Visualizations"
      Height          =   3975
      Index           =   1
      Left            =   2520
      TabIndex        =   13
      Top             =   120
      Width           =   4695
      Begin VB.CheckBox chkVis 
         Caption         =   "&Disable fade on all presets."
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   2295
      End
      Begin MSComctlLib.Slider sldDelay 
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   720
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   661
         _Version        =   393216
         Min             =   10
         Max             =   100
         SelStart        =   10
         TickFrequency   =   10
         Value           =   10
      End
      Begin VB.Label Label17 
         Caption         =   $"frmSettings.frx":00A3
         Height          =   855
         Left            =   360
         TabIndex        =   56
         Top             =   2880
         Width           =   3975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "every 20 milliseconds"
         Height          =   195
         Left            =   2835
         TabIndex        =   17
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Visualization update speed:"
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame frm 
      Caption         =   "Sound Device"
      Height          =   3975
      Index           =   4
      Left            =   2520
      TabIndex        =   26
      Top             =   120
      Width           =   4695
      Begin VB.ComboBox cmbEngine 
         Height          =   315
         ItemData        =   "frmSettings.frx":015B
         Left            =   240
         List            =   "frmSettings.frx":0168
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   600
         Width           =   4215
      End
      Begin VB.ComboBox cmbDevice 
         Height          =   315
         ItemData        =   "frmSettings.frx":01A3
         Left            =   240
         List            =   "frmSettings.frx":01A5
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1200
         Width           =   4215
      End
      Begin VB.ComboBox cmbMixer 
         Height          =   315
         ItemData        =   "frmSettings.frx":01A7
         Left            =   240
         List            =   "frmSettings.frx":01B7
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1800
         Width           =   2295
      End
      Begin VB.ComboBox cmbChannels 
         Height          =   315
         ItemData        =   "frmSettings.frx":021E
         Left            =   2760
         List            =   "frmSettings.frx":023A
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   2400
         Width           =   1215
      End
      Begin VB.ComboBox cmbBuffer 
         Height          =   315
         ItemData        =   "frmSettings.frx":0260
         Left            =   240
         List            =   "frmSettings.frx":0276
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   2400
         Width           =   1575
      End
      Begin VB.ComboBox cmbFreq 
         Height          =   315
         ItemData        =   "frmSettings.frx":0298
         Left            =   2760
         List            =   "frmSettings.frx":02A8
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CheckBox chkDSP 
         Caption         =   "Enable D&SP"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CheckBox chkDXFX 
         Caption         =   "Enable Direct&X 8 Effects"
         Height          =   255
         Left            =   2160
         TabIndex        =   50
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "&Max Channels:"
         Height          =   195
         Left            =   2760
         TabIndex        =   44
         Top             =   2160
         Width           =   1050
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "&Buffer size (ms):"
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   2160
         Width           =   1110
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Sampling &Frequency:"
         Height          =   195
         Left            =   2760
         TabIndex        =   36
         Top             =   1560
         Width           =   1485
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Output &Engine:"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Mixer:"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Output &Device:"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   960
         Width           =   1080
      End
   End
   Begin VB.Frame frm 
      Caption         =   "Speakers"
      Height          =   3975
      Index           =   5
      Left            =   2520
      TabIndex        =   31
      Top             =   120
      Width           =   4695
      Begin VB.CheckBox chkSurround 
         Caption         =   "S&urround"
         Height          =   255
         Left            =   3240
         TabIndex        =   41
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox cmbSpeaker 
         Height          =   315
         ItemData        =   "frmSettings.frx":02C8
         Left            =   240
         List            =   "frmSettings.frx":02DE
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   600
         Width           =   4215
      End
      Begin MSComctlLib.Slider sldPanning 
         Height          =   495
         Left            =   360
         TabIndex        =   39
         Top             =   1440
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   51
         Max             =   255
         TickFrequency   =   128
      End
      Begin MSComctlLib.Slider sldPanSeperation 
         Height          =   495
         Left            =   360
         TabIndex        =   45
         Top             =   2280
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   3
         TickFrequency   =   10
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Speaker &Panning:"
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Pan S&eperation:"
         Height          =   195
         Left            =   240
         TabIndex        =   43
         Top             =   2040
         Width           =   1140
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "L"
         Height          =   195
         Left            =   240
         TabIndex        =   49
         Top             =   1560
         Width           =   90
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "R"
         Height          =   195
         Left            =   2640
         TabIndex        =   47
         Top             =   1560
         Width           =   120
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "&Speaker Setup:"
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   1110
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   6165
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   706
      LabelEdit       =   1
      Style           =   4
      Appearance      =   1
   End
   Begin VB.Frame frm 
      Caption         =   "File Types"
      Height          =   3975
      Index           =   6
      Left            =   2520
      TabIndex        =   51
      Top             =   120
      Width           =   4695
      Begin VB.ComboBox cmbAction 
         Height          =   315
         ItemData        =   "frmSettings.frx":0384
         Left            =   1560
         List            =   "frmSettings.frx":0391
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   3480
         Width           =   2295
      End
      Begin VB.CheckBox chkReg 
         Caption         =   "&Register file types on Simple Amp start."
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   3120
         Width           =   3135
      End
      Begin MSComctlLib.ListView lvwType 
         Height          =   2775
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "types"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Default &Action:"
         Height          =   195
         Left            =   240
         TabIndex        =   54
         Top             =   3530
         Width           =   1050
      End
   End
   Begin VB.Frame frm 
      Caption         =   "File Browser"
      Height          =   3975
      Index           =   2
      Left            =   2520
      TabIndex        =   57
      Top             =   120
      Width           =   4695
      Begin VB.Frame Frame2 
         Caption         =   "Read Extended File Information"
         Height          =   3375
         Left            =   240
         TabIndex        =   58
         Top             =   360
         Width           =   4215
         Begin VB.OptionButton optFile 
            Caption         =   "&Never."
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   61
            Top             =   1080
            Width           =   975
         End
         Begin VB.OptionButton optFile 
            Caption         =   "&Only when browsing local and fixed drives."
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   60
            Top             =   720
            Width           =   3375
         End
         Begin VB.OptionButton optFile 
            Caption         =   "&Always."
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   59
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "If you have this file in your media library, however, the extended info is read from there, always."
            Height          =   495
            Left            =   240
            TabIndex        =   63
            Top             =   2880
            Width           =   3495
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmSettings.frx":03CF
            Height          =   1455
            Left            =   240
            TabIndex        =   62
            Top             =   1440
            Width           =   3735
         End
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkSurround_Click()
  On Error Resume Next
  devSurround = CBool(chkSurround.Value)
  Sound.StreamSurround = devSurround
End Sub

Private Sub cmbEngine_Click()
  ShowDrivers
End Sub

Private Sub cmbMixer_Click()
  On Error Resume Next
  If cmbMixer.ListIndex > 1 Then
    chkSurround.Enabled = False
  Else
    chkSurround.Enabled = True
  End If
End Sub

Private Sub cmbSpeaker_Click()
  On Error Resume Next
  devSpeaker = cmbSpeaker.ListIndex
  If devSpeaker = 1 Then
    Sound.SpeakerSetup = FSOUND_SPEAKERMODE_HEADPHONE
  ElseIf devSpeaker = 2 Then
    Sound.SpeakerSetup = FSOUND_SPEAKERMODE_STEREO
  ElseIf devSpeaker = 3 Then
    Sound.SpeakerSetup = FSOUND_SPEAKERMODE_QUAD
  ElseIf devSpeaker = 4 Then
    Sound.SpeakerSetup = FSOUND_SPEAKERMODE_SURROUND
  ElseIf devSpeaker = 5 Then
    Sound.SpeakerSetup = FSOUND_SPEAKERMODE_DOLBYDIGITAL
  End If
End Sub

Private Sub cmdClean_Click()
  frmLibClean.Show , Me
End Sub

Private Sub cmdClose_Click()
  On Error GoTo errh
  Dim X As Long, bRestart As Boolean
  
  With Settings
  
    .AlwaysTray = CBool(chkGeneral(0))
    .StartInTray = CBool(chkGeneral(1))
    .OnTop = CBool(chkGeneral(2))
    .Snap = CBool(chkGeneral(3))
    .ButtonDefault = CBool(chkGeneral(4))
    .Fade = CBool(chkGeneral(5))
    .DynamicColumns = CBool(chkGeneral(6))
    
    .NoPresetFade = CBool(chkVis)
    
    For X = 0 To 2
      If optFile(X).Value Then
        .BrowseMode = X
        Exit For
      End If
    Next X
    
    .UseLibrary = CBool(chkLibrary(0))
    .LibGetName = CBool(chkLibrary(1))
    .LibAutosize = CBool(chkLibrary(2))
    .LibHideNonMusic = CBool(chkLibrary(3))
    
    .DXFXon = CBool(chkDXFX)
    
    If devDSP <> CBool(chkDSP) Then bRestart = True
    If devType <> cmbEngine.ListIndex Then bRestart = True
    If devDevice <> cmbDevice.ListIndex Then bRestart = True
    If devMixer <> cmbMixer.ListIndex Then bRestart = True
    If devFreq <> Val(cmbFreq.List(cmbFreq.ListIndex)) Then bRestart = True
    If devChannels <> Val(cmbChannels.List(cmbChannels.ListIndex)) Then bRestart = True
    If devBuffer <> Val(cmbBuffer.List(cmbBuffer.ListIndex)) Then bRestart = True
    
    devDSP = CBool(chkDSP)
    devType = cmbEngine.ListIndex
    devDevice = cmbDevice.ListIndex
    devMixer = cmbMixer.ListIndex
    devFreq = Val(cmbFreq.List(cmbFreq.ListIndex))
    devChannels = Val(cmbChannels.List(cmbChannels.ListIndex))
    devBuffer = Val(cmbBuffer.List(cmbBuffer.ListIndex))
    
    .AssOnStart = CBool(chkReg)
    .AssAction = cmbAction.ListIndex
    
    Dim bAss As Boolean
    For X = 0 To UBound(Settings.AssType)
      Settings.AssType(X) = lvwType.ListItems(X + 1).Checked
      If Not bAss Then bAss = lvwType.ListItems(X + 1).Checked
    Next
    
    'associate types now if they are not to be associated at start and
    'there are types selected.
    If Not .AssOnStart And bAss Then modMain.AssociateTypes
    
    SaveSettings
    
    If bRestart Then
      If MsgBox("You have made changes that require the sound engine to restart. If you do not restart the engine now, the changes you have made will not take effect until the next time you start Simple Amp." & vbCrLf & vbCrLf & "Do you want to restart the sound engine now?", vbQuestion Or vbYesNo, "Engine restart needed") = vbYes Then
        'restart fmod!
        'close
        Dim b As Boolean
        With Sound
          If Sound.StreamIsPlaying Or Sound.MusicIsPlaying Then
            frmMain.PlayStop
            b = True
          End If
          .StreamUnload
          .MusicUnload
          .CDStop
          frmMain.tmrScope.Enabled = False
          If devDSP Then
            FSOUND_DSP_SetActive FSOUND_DSP_GetFFTUnit, False
            FSOUND_DSP_SetActive DSP_Handle, False
            FSOUND_DSP_Free DSP_Handle
          End If
        End With
        Set Sound = Nothing
        'start
        Set Sound = New clsFMOD
        With Sound
          Select Case devType
            Case 1
              .InitOutput FSOUND_OUTPUT_DSOUND
            Case 2
              .InitOutput FSOUND_OUTPUT_WINMM
            Case Else
              .InitOutput -1
          End Select
          .InitDriver devDevice
          Select Case devMixer
            Case 1
              .InitMixer FSOUND_MIXER_QUALITY_FPU
            Case 2
              .InitMixer FSOUND_MIXER_QUALITY_MMXP5
            Case 3
              .InitMixer FSOUND_MIXER_QUALITY_MMXP6
            Case Else
              .InitMixer FSOUND_MIXER_QUALITY_AUTODETECT
          End Select
          .InitBuffer devBuffer
          .Init devFreq, devChannels
          If devDSP Then
            ReDim Preserve ScopeBufferFPU(FSOUND_DSP_GetBufferLength)
            ReDim Preserve ScopeBufferINT(FSOUND_DSP_GetBufferLength)
            ReDim Preserve ScopeUPeaks(UBound(ScopeBufferFPU))
            ReDim Preserve ScopeLPeaks(UBound(ScopeBufferFPU))
            DSP_Handle = FSOUND_DSP_Create(AddressOf ScopeCallback, FSOUND_DSP_DEFAULTPRIORITY_USER + 3, 0)
            FSOUND_DSP_SetActive FSOUND_DSP_GetFFTUnit, True
            FSOUND_DSP_SetActive DSP_Handle, True
            Call FSOUND_DSP_SetActive(FSOUND_DSP_GetFFTUnit, True) 'Set DSP FFT Unit to active
            If DSP_Handle <> 0 Then DSP_OK = True
          End If
          'setup speaker, devSpeaker = 0 : no change
          If devSpeaker = 1 Then
            .SpeakerSetup = FSOUND_SPEAKERMODE_HEADPHONE
          ElseIf devSpeaker = 2 Then
            .SpeakerSetup = FSOUND_SPEAKERMODE_STEREO
          ElseIf devSpeaker = 3 Then
            .SpeakerSetup = FSOUND_SPEAKERMODE_QUAD
          ElseIf devSpeaker = 4 Then
            .SpeakerSetup = FSOUND_SPEAKERMODE_SURROUND
          ElseIf devSpeaker = 5 Then
            .SpeakerSetup = FSOUND_SPEAKERMODE_DOLBYDIGITAL
          End If
        End With
        frmMain.tmrScope.Enabled = True
        If b Then frmMain.Play
      End If
    End If
    
  End With
  
  frmPlaylist.List.Refresh
  Unload Me
  Exit Sub
errh:
  If cLog.ErrorMsg(Err, "frmSettings, cmdClose_Click") = vbYes Then Resume Next Else frmMain.UnloadAll
End Sub

Private Sub cmdHelp_Click()
  frmHelp.Show , Me
  frmHelp.web.Navigate App.Path & "\docs\config.html"
End Sub

Private Sub Form_Activate()
  AlwaysOnTop Me, Settings.OnTop
End Sub

Private Sub Form_Load()
  On Error Resume Next
  Dim NewNode As Node
  Set NewNode = tvw.Nodes.Add(, , "m0", "Options")
  NewNode.Expanded = True
  Set NewNode = tvw.Nodes.Add("m0", tvwChild, "d0", "General")
  NewNode.Selected = True
  Set NewNode = tvw.Nodes.Add("m0", tvwChild, "d1", "Visualizations")
  Set NewNode = tvw.Nodes.Add("m0", tvwChild, "d2", "File Browser")
  Set NewNode = tvw.Nodes.Add("m0", tvwChild, "d3", "Media Library")
  Set NewNode = tvw.Nodes.Add(, , "m4", "System")
  NewNode.Expanded = True
  Set NewNode = tvw.Nodes.Add("m4", tvwChild, "d4", "Sound Device")
  Set NewNode = tvw.Nodes.Add("m4", tvwChild, "d5", "Speakers")
  Set NewNode = tvw.Nodes.Add("m4", tvwChild, "d6", "File Types")
  
  UpdateCtrls
  
End Sub

Private Sub lvwType_KeyUp(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  Dim X As Long
  If Shift = vbCtrlMask And KeyCode = vbKeyA Then
    For X = 1 To lvwType.ListItems.Count
      lvwType.ListItems(X).Checked = True
    Next
  ElseIf Shift = vbCtrlMask And KeyCode = vbKeyN Then
    For X = 1 To lvwType.ListItems.Count
      lvwType.ListItems(X).Checked = False
    Next
  End If
End Sub

Private Sub sldDelay_Scroll()
  On Error Resume Next
  VisUpdateInt = sldDelay.Value
  frmMain.tmrScope.Interval = VisUpdateInt
  Label3 = "every " & sldDelay.Value & " milliseconds"
End Sub

Private Sub sldIcon_Scroll()
  On Error Resume Next
  Settings.TrayIcon = sldIcon.Value
  imgIcon.Picture = LoadResPicture(Settings.TrayIcon + 104, vbResIcon)
  frmMain.SetupSystrayIcon Settings.TrayIcon
End Sub

Private Sub sldPanning_Scroll()
  On Error Resume Next
  If sldPanning.Value < 138 And sldPanning.Value > 118 Then sldPanning.Value = 128
  devPanning = sldPanning.Value
  Sound.StreamPanning = devPanning
End Sub

Private Sub sldPanSeperation_Scroll()
  On Error Resume Next
  devPanSep = sldPanSeperation.Value / 10
  Sound.MusicPanSep = devPanSep
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
  On Error Resume Next
  Dim X As Long
  For X = 0 To 6
    frm(X).Visible = (X = Val(Right(Node.Key, 1)))
  Next X
End Sub

Public Sub UpdateCtrls()
  On Error Resume Next
  Dim X As Long
  
  With Settings
    chkGeneral(0).Value = Abs(.AlwaysTray)
    chkGeneral(1).Value = Abs(.StartInTray)
    chkGeneral(2).Value = Abs(.OnTop)
    chkGeneral(3).Value = Abs(.Snap)
    chkGeneral(4).Value = Abs(.ButtonDefault)
    chkGeneral(5).Value = Abs(.Fade)
    chkGeneral(6).Value = Abs(.DynamicColumns)
  
    sldIcon.Value = .TrayIcon
    imgIcon.Picture = LoadResPicture(.TrayIcon + 104, vbResIcon)
    
    sldDelay.Value = VisUpdateInt
    Label3 = "every " & sldDelay.Value & " milliseconds"
    chkVis.Value = Abs(.NoPresetFade)
    
    optFile(.BrowseMode).Value = True
    
    chkLibrary(0).Value = Abs(.UseLibrary)
    chkLibrary(1).Value = Abs(.LibGetName)
    chkLibrary(2).Value = Abs(.LibAutosize)
    chkLibrary(3).Value = Abs(.LibHideNonMusic)
    
    cmbEngine.ListIndex = devType
    ShowDrivers
    cmbMixer.ListIndex = devMixer
    
    For X = 0 To cmbFreq.ListCount - 1
      If Val(cmbFreq.List(X)) = devFreq Then
        cmbFreq.ListIndex = X
        Exit For
      End If
    Next X
    If cmbFreq.ListIndex = -1 Then
      cmbFreq.ListIndex = 2
    End If
    
    For X = 0 To cmbChannels.ListCount - 1
      If Val(cmbChannels.List(X)) = devChannels Then
        cmbChannels.ListIndex = X
        Exit For
      End If
    Next X
    If cmbChannels.ListIndex = -1 Then
      cmbChannels.ListIndex = 3
    End If
    
    For X = 0 To cmbBuffer.ListCount - 1
      If Val(cmbBuffer.List(X)) = devBuffer Then
        cmbBuffer.ListIndex = X
        Exit For
      End If
    Next X
    If cmbBuffer.ListIndex = -1 Then
      cmbBuffer.ListIndex = 5
    End If
    
    chkDXFX.Value = Abs(.DXFXon)
    chkDSP.Value = Abs(devDSP)
    sldPanning.Value = devPanning
    sldPanning.Value = devPanning
    chkSurround.Value = Abs(devSurround)
    sldPanSeperation.Value = devPanSep * 10
    cmbSpeaker.ListIndex = devSpeaker
    
    With lvwType.ListItems
      .Add , , "Mpeg Audio Files (mp3, mp2)"
      .Add , , "Ogg Vorbis Audio Files (ogg)"
      .Add , , "Microsoft Windows Media Audio Files (wma)"
      .Add , , "Microsoft Advanced Sound Format Audio Files (asf)"
      .Add , , "Microsoft Waveform Audio Files (wav)"
      .Add , , "Protracker/Fasttracker Module Files (mod)"
      .Add , , "Screamtracker 3 Module Files (s3m)"
      .Add , , "Fasttracker 2 Module Files (xm)"
      .Add , , "Impulse Tracker Module Files (it)"
      .Add , , "MIDI Files (mid, midi, rmi)"
      .Add , , "DirectMusic Segment Files (sgm)"
      .Add , , "Simple Amp Playlist Files (playlist)"
      .Add , , "Other Supported Playlist Files (m3u, pls)"
    End With
    
    For X = 0 To UBound(.AssType)
      lvwType.ListItems(X + 1).Checked = .AssType(X)
    Next
    
    chkReg.Value = Abs(.AssOnStart)
    cmbAction.ListIndex = .AssAction
    
    If cmbMixer.ListIndex > 1 Then chkSurround.Enabled = False
    
  End With
  
End Sub

Private Sub ShowDrivers()
  Dim X As Long
  On Error Resume Next
  
  cmbDevice.Clear
  If cmbEngine.ListIndex = 0 Then
    cmbDevice.AddItem "<Default>"
  ElseIf cmbEngine.ListIndex = 1 Then
    For X = 0 To UBound(Direct_Dev)
      cmbDevice.AddItem Direct_Dev(X)
    Next X
  ElseIf cmbEngine.ListIndex = 2 Then
    For X = 0 To UBound(WinOut_Dev)
      cmbDevice.AddItem WinOut_Dev(X)
    Next X
  End If
  
  If cmbDevice.ListCount = 0 Then
    cmbDevice.AddItem "<No Devices Detected>"
  End If
  
  If devDevice > cmbDevice.ListCount - 1 Then
    cmbDevice.ListIndex = 0
  Else
    cmbDevice.ListIndex = devDevice
  End If

End Sub
