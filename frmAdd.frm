VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{19B7F2A2-1610-11D3-BF30-1AF820524153}#1.2#0"; "ccrpftv6.ocx"
Begin VB.Form frmAdd 
   Caption         =   "File Browser"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   150
   ClientWidth     =   10410
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   10410
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   7800
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdd.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdd.frx":0C1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdd.frx":0F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdd.frx":12C6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrShutdown 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6840
      Top             =   720
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Height          =   330
      Index           =   0
      Left            =   3720
      TabIndex        =   4
      Top             =   60
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ButtonWidth     =   2381
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Selected"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar prbMain 
      Height          =   255
      Left            =   7800
      TabIndex        =   3
      Top             =   6795
      Width           =   1695
      Visible         =   0   'False
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6675
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   661
      Style           =   1
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   1680
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdd.frx":161A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdd.frx":196E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdd.frx":1CC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdd.frx":2016
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdd.frx":236A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdd.frx":26BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdd.frx":2A12
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   6135
      Left            =   3600
      TabIndex        =   1
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   10821
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imlIcons"
      SmallIcons      =   "imlIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Length"
         Object.Width           =   2540
      EndProperty
   End
   Begin CCRPFolderTV6.FolderTreeview ftwMain 
      Height          =   6570
      Left            =   50
      TabIndex        =   0
      Top             =   50
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   11589
      AutoUpdate      =   0   'False
      IntegralHeight  =   0   'False
      VirtualFolders  =   0   'False
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Height          =   330
      Index           =   1
      Left            =   4440
      TabIndex        =   5
      Top             =   60
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   582
      ButtonWidth     =   1588
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add All"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Height          =   330
      Index           =   2
      Left            =   5400
      TabIndex        =   6
      Top             =   60
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ButtonWidth     =   2090
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Folder"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Height          =   330
      Index           =   4
      Left            =   6600
      TabIndex        =   7
      Top             =   60
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      ButtonWidth     =   1349
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Filter"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Height          =   330
      Index           =   5
      Left            =   7440
      TabIndex        =   8
      Top             =   60
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      ButtonWidth     =   1296
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "View"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Height          =   330
      Index           =   3
      Left            =   8280
      TabIndex        =   9
      Top             =   60
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ButtonWidth     =   2328
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear Playlist"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.Image imgDrag 
      Height          =   6135
      Left            =   3480
      MousePointer    =   9  'Size W E
      Top             =   480
      Width           =   135
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'---------
'Shell file remove/copy/move etc...
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As Byte) As Long

Private Type SHFILEOPSTRUCT
  hwnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAnyOperationsAborted As Long
  hNameMappings As Long
  lpszProgressTitle As String
End Type

Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_FILESONLY = &H80
'---------

Private Const DRIVE_CDROM = 5
Private Const DRIVE_FIXED = 3
Private Const DRIVE_RAMDISK = 6
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_NOTFOUND = 1
Private Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Type tListData
  'mp3 info
  dDate        As Date
  lSize        As Long
  sFilename    As String 'Filename
  sLength      As String 'formatted
  sArtistTitle As String 'First field
  sAlbum       As String 'Second field
  sGenre       As String 'Third field
  eType        As enumFileType
  lItem        As Long
End Type

Private Mousepos As POINTAPI 'used for doubleclick in lvwMain

Public bWorking As Boolean 'if true, files are being read/displayed
Private bDrag As Boolean
Private OldX As Long
Private AddList() As tListData

Private Sub Form_Activate()
  On Error Resume Next
  
  tmrShutdown.Enabled = False
  AlwaysOnTop Me, Settings.OnTop
End Sub

Private Sub Form_Load()
  On Error Resume Next
  Dim X As Long
  
  ReDim AddList(0) 'reset list data
  'setup window from saved settings
  If Settings.AddWidth > 0 Then Me.Width = Settings.AddWidth
  If Settings.AddHeight > 0 Then Me.Height = Settings.AddHeight
  If Settings.AddFolderWidth > 0 Then imgDrag.Left = Settings.AddFolderWidth
  If Settings.AddMax Then Me.WindowState = 2
  frmMenus.menView(Settings.AddView).Checked = True
  frmMenus.menAutosizeColumns.Checked = Settings.AddAutosize
  
  Me.Show 'show window
  DoEvents 'wait a bit
  
  AlwaysOnTop Me, Settings.OnTop
  
  CheckExtended
  'setup more stuff from saved settings
  If Len(Settings.AddDir) > 0 Then ftwMain.SelectedFolder = Settings.AddDir
  For X = 0 To UBound(Settings.Filter)
    frmMenus.menFilt(X).Checked = Settings.Filter(X)
  Next X
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  
  'Resize controls
  ftwMain.Move ftwMain.Left, ftwMain.Top, imgDrag.Left - 30, Me.Height - stbMain.Height - 540
  lvwMain.Move imgDrag.Left + 100, lvwMain.Top, Me.Width - ftwMain.Width - 300, Me.Height - stbMain.Height - 640 - tbrMain(0).Height
  tbrMain(0).Left = lvwMain.Left + 30
  tbrMain(1).Left = tbrMain(0).Left + tbrMain(0).Width
  tbrMain(2).Left = tbrMain(1).Left + tbrMain(1).Width + 30
  tbrMain(3).Left = tbrMain(2).Left + tbrMain(2).Width + 30
  tbrMain(4).Left = tbrMain(3).Left + tbrMain(3).Width + 30
  tbrMain(5).Left = tbrMain(4).Left + tbrMain(4).Width + 30
  imgDrag.Height = lvwMain.Height
  
  prbMain.Move Me.Width - prbMain.Width - 450, stbMain.Top + 90
  'resize list columns
  If Not Settings.AddAutosize Then
    Select Case Settings.AddView
      Case 1
        lvwMain.ColumnHeaders(1).Width = lvwMain.Width - 1200
        lvwMain.ColumnHeaders(2).Width = 800
      Case 2
        lvwMain.ColumnHeaders(1).Width = (lvwMain.Width - 2400) * 0.68
        lvwMain.ColumnHeaders(2).Width = (lvwMain.Width - 2400) * 0.3
        lvwMain.ColumnHeaders(3).Width = 1200
        lvwMain.ColumnHeaders(4).Width = 900
      Case Else
        lvwMain.ColumnHeaders(1).Width = lvwMain.Width - 3200
        lvwMain.ColumnHeaders(2).Width = 1800
        lvwMain.ColumnHeaders(3).Width = 1000
    End Select
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  
  If bWorking Then
    Me.Hide
    tmrShutdown.Enabled = True
    Cancel = 1
  End If
 
  'Save settings
  Settings.AddDir = ftwMain.SelectedFolder
  If Me.WindowState = 2 Then
    Settings.AddMax = True
  Else
    Settings.AddMax = False
    Settings.AddWidth = Me.Width
    Settings.AddHeight = Me.Height
  End If
  Settings.AddFolderWidth = ftwMain.Width
End Sub

Private Sub ftwMain_FolderClick(Folder As CCRPFolderTV6.Folder, Location As CCRPFolderTV6.ftvHitTestConstants)
  On Error Resume Next
  If Not bWorking Then
    UpdateFolder Folder.FullPath
  Else
    Folder.Selected = False
  End If
End Sub

Public Sub UpdateFolder(ByVal sFldr As String)
  'Finds files in current folder
  On Error GoTo ErrHandler
  Dim File As New clsFind
  Dim sFilter As String
  Dim lTime As Long, X As Long, tItem As Long
  Dim bLocalDrive As Boolean
  
  If Len(sFldr) = 0 Then Exit Sub
  cLog.StartTimer
  cLog.Log "GETTING FILES FROM FOLDER...", 3, False
  bWorking = True
  sFldr = sAppend(sFldr, "\")
  Me.MousePointer = vbArrowHourglass
  stbMain.SimpleText = "Finding files..."
  DoEvents
  
  'check driver type
  On Error Resume Next
  
  Select Case GetDriveType(Left(sFldr, 3))
    Case DRIVE_FIXED, DRIVE_RAMDISK
      bLocalDrive = True
    Case Else 'drive not found or other error
      bLocalDrive = False
  End Select
  
  On Error GoTo ErrHandler
  
  'Figure out the filter
  With Settings
    If .Filter(0) Then sFilter = sFilter + "*.wav;"
    If .Filter(1) Then sFilter = sFilter + "*.mp2;*.mp3;"
    If .Filter(2) Then sFilter = sFilter + "*.ogg;"
    If .Filter(3) Then sFilter = sFilter + "*.wma;"
    If .Filter(4) Then sFilter = sFilter + "*.asf;"
    If .Filter(5) Then sFilter = sFilter + "*.mod;"
    If .Filter(6) Then sFilter = sFilter + "*.s3m;"
    If .Filter(7) Then sFilter = sFilter + "*.xm;"
    If .Filter(8) Then sFilter = sFilter + "*.it;"
    If .Filter(9) Then sFilter = sFilter + "*.mid;*.rmi;*.midi;"
    If .Filter(10) Then sFilter = sFilter + "*.sgm;"
    If .Filter(11) Then sFilter = sFilter + "*.playlist;*.pls;*.m3u;"
    If Len(sFilter) = 0 Then sFilter = "*.mp3;*.mp2;"
    sFilter = Left(sFilter, Len(sFilter) - 1) 'Remove ending ';'
  End With

  'Ok, find files
  File.Find sFldr, sFilter
  ReDim AddList(0) 'reset file list
  
  'If files were found
  If File.Count > 0 Then
    
    prbMain.Value = 0
    prbMain.Visible = True
    prbMain.Max = File.Count
    For X = 1 To File.Count
      prbMain.Value = X
      stbMain.SimpleText = "Reading file " & X & "/" & File.Count & "... (" & File(X).sNameAndExtension & ")"
      DoEvents
      
      'If it is an playlist, just add the file to list.
      'if it is an other media file, add to library and grab its data.
      Select Case File(X).sExtension
        Case "playlist", "m3u", "pls" 'file is playlist
        
          ReDim Preserve AddList(UBound(AddList) + 1)
          With AddList(UBound(AddList))
            .sFilename = File(X).sFilename
            .eType = TYPE_PLAYLIST
            .sArtistTitle = File(X).sNameAndExtension
            If Settings.BrowseMode = 0 Or (Settings.BrowseMode = 1 And bLocalDrive = True) Then
              .sLength = GetPlaylistLength(.sFilename) & " items"
            Else
              .sLength = "? items"
            End If
          End With
          
        Case Else 'file is media

          tItem = LibraryCheck(File(X).sFilename) 'get library index
          If tItem <> 0 Then 'this file exists in library already!
          
            'This file exists in index library, so...
            If LibraryIndex(tItem).lReference = 0 Then '...if it has not yet been loaded
              'create the new item in the media library
              ReDim Preserve Library(UBound(Library) + 1)
              LibraryIndex(tItem).lReference = UBound(Library)
              'but the item exists in the library (in file), so load it
              LoadItem UBound(Library), tItem
              tItem = UBound(Library) 'and set up the reference
            Else '...if it has been loaded
              tItem = LibraryIndex(tItem).lReference 'just set up the reference
            End If
            
            'add to browser list
            ReDim Preserve AddList(UBound(AddList) + 1)
            With AddList(UBound(AddList))
              .dDate = File(X).dLastWriteTime
              .lSize = File(X).lSize / 1024
              .lItem = tItem
              .sFilename = File(X).sFilename
              .sArtistTitle = Library(tItem).sArtistTitle
              .eType = Library(tItem).eType
              .sGenre = Library(tItem).sGenre
              .sLength = ConvertTime(Library(tItem).lLength, True)
              lTime = lTime + Library(tItem).lLength
              If .eType = TYPE_MOD Or .eType = TYPE_IT Or .eType = TYPE_S3M Or .eType = TYPE_XM Then
                .sAlbum = Library(tItem).sTitle
              Else
                .sAlbum = Library(tItem).sAlbum
              End If
            End With
            
          Else 'this file does not exist in library...
          
            If Settings.BrowseMode = 0 Or (Settings.BrowseMode = 1 And bLocalDrive = True) Then
              '...but settings allow us to read it's info so we do that
              tItem = CreateLibrary(File(X).sFilename, False)
              
              If tItem <> 0 Then
                'add to browser list
                ReDim Preserve AddList(UBound(AddList) + 1)
                With AddList(UBound(AddList))
                  .dDate = File(X).dLastWriteTime
                  .lSize = File(X).lSize / 1024
                  .lItem = tItem
                  .sFilename = File(X).sFilename
                  .sArtistTitle = Library(tItem).sArtistTitle
                  .eType = Library(tItem).eType
                  .sGenre = Library(tItem).sGenre
                  .sLength = ConvertTime(Library(tItem).lLength, True)
                  lTime = lTime + Library(tItem).lLength
                  If .eType = TYPE_MOD Or .eType = TYPE_IT Or .eType = TYPE_S3M Or .eType = TYPE_XM Then
                    .sAlbum = Library(tItem).sTitle
                  Else
                    .sAlbum = Library(tItem).sAlbum
                  End If
                End With
              Else
                cLog.Log "FILE ERROR (NOT LISTING): " & File(X).sNameAndExtension, 5, True
              End If
              
            Else
              '...and settings do not allow us to read it's info so we do what
              'we can with the little info we have.
              
              ReDim Preserve AddList(UBound(AddList) + 1)
              With AddList(UBound(AddList))
                .dDate = File(X).dLastWriteTime
                .lSize = File(X).lSize / 1024
                .sFilename = File(X).sFilename
                .sArtistTitle = File(X).sNameAndExtension
                .sAlbum = " " 'you must set to one space or listview will
                .sGenre = " " 'act strange when adding these vars to columns
                .sLength = " "
                Select Case LCase(File(X).sExtension)
                  Case "mp3", "mp2", "wma", "asf", "wav", "ogg"
                    .eType = TYPE_MP2_MP3
                  Case "mod", "xm", "it", "s3m"
                    .eType = TYPE_MOD
                  Case Else
                    .eType = TYPE_MID_RMI
                End Select
              End With
              
            End If
            
          End If
        
      End Select
      
      'setting bworking to false will abort any operation
      If Not bWorking Then Exit Sub 'abort
      
    Next X
      
  End If
  
  cLog.Log "OK. (" & File.Count & " files found, " & cLog.GetTimer & " ms)", 3
  
  UpdateList
  
  prbMain.Visible = False
  If UBound(AddList) > 0 Then
    stbMain.SimpleText = UBound(AddList) & " item(s) found, total " & ConvertTime(lTime, False) & "."
  Else
    stbMain.SimpleText = "0 item(s) found."
  End If
  
  'ftwMain.SelectedFolder = sFldr
  
  Me.MousePointer = vbDefault
  bWorking = False
  
  Exit Sub
ErrHandler:
  If cLog.ErrorMsg(Err, "frmAdd, UpdateFolder(" & sFldr & ")") = vbNo Then frmMain.UnloadAll Else Resume Next
End Sub

Private Sub imgDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  On Error Resume Next
  bDrag = True
  OldX = imgDrag.Left - X
End Sub

Private Sub imgDrag_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  'Changing size of tree view window
  On Error Resume Next
  If bDrag And OldX + X > 2000 And OldX + X < 6000 Then
    If OldX + X <> imgDrag.Left Then
      ftwMain.Width = OldX + X - 30
      lvwMain.Move OldX + X + 100, lvwMain.Top, Me.Width - ftwMain.Width - 300
      tbrMain(0).Left = lvwMain.Left + 30
      tbrMain(1).Left = tbrMain(0).Left + tbrMain(0).Width
      tbrMain(2).Left = tbrMain(1).Left + tbrMain(1).Width + 30
      tbrMain(3).Left = tbrMain(2).Left + tbrMain(2).Width + 30
      tbrMain(4).Left = tbrMain(3).Left + tbrMain(3).Width + 30
      tbrMain(5).Left = tbrMain(4).Left + tbrMain(4).Width + 30
    End If
  End If
End Sub

Private Sub imgDrag_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  On Error Resume Next
  bDrag = False
  imgDrag.Left = lvwMain.Left - 100
  Form_Resize
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  'the list is not sorted first, in case you want to add the files
  'in their original order, but once you click a column the list gets
  'sorted.
  On Error Resume Next
  If Not lvwMain.Sorted Then
    lvwMain.Sorted = True
  Else
    If lvwMain.SortKey = ColumnHeader.index - 1 Then
      If lvwMain.SortOrder = lvwAscending Then
        lvwMain.SortOrder = lvwDescending
      Else
        lvwMain.SortOrder = lvwAscending
      End If
    End If
  End If
  lvwMain.SortKey = ColumnHeader.index - 1
End Sub

Private Sub lvwMain_DblClick()
  On Error GoTo errh
  Dim hit As ListItem
  Set hit = lvwMain.HitTest(Mousepos.X, Mousepos.y)
  
  'play file doubleclicked on, without adding to playlist
  hit.Selected = True
  PlaySelected
  
errh:
End Sub

Private Sub lvwMain_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  On Error Resume Next
  Mousepos.X = X
  Mousepos.y = y
End Sub

Private Sub lvwMain_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  On Error Resume Next
  frmMenus.menAutosizeColumns.Checked = Settings.AddAutosize
  If Button = 2 Then PopupMenu frmMenus.menAddRClick, , , , frmMenus.menPlayfoca
End Sub

Private Sub tbrMain_ButtonClick(index As Integer, ByVal Button As MSComctlLib.Button)
  On Error Resume Next
  Select Case index
  Case 0
    AddSelected
  Case 1
    AddAll
  Case 2
    AddDirs ftwMain.SelectedFolder
  Case 3
    frmMenus.menDeleteAll_Click
  Case 4  'filter
    ShowPopup frmAdd, frmMenus.menFilter, tbrMain(4), TextWidth(frmMenus.menFilt(11).Caption) + (45 * Screen.TwipsPerPixelX)
    frmMenus.menFilter.Visible = True
  Case 5  'ext list
    ShowPopup frmAdd, frmMenus.menAddViews, tbrMain(5), TextWidth(frmMenus.menAutosizeColumns.Caption) + (45 * Screen.TwipsPerPixelX)
    frmMenus.menAddViews.Visible = True
  End Select
End Sub

Public Sub CheckExtended()
  On Error Resume Next
  
  lvwMain.ListItems.Clear
  lvwMain.ColumnHeaders.Clear
  
  With lvwMain.ColumnHeaders
    Select Case Settings.AddView
      Case 1
        .Add , , "Filename"
        .Add , , "Length", , lvwColumnRight
      Case 2
        .Add , , "Title & Artist"
        .Add , , "Album"
        .Add , , "Genre"
        .Add , , "Length", , lvwColumnRight
      Case Else
        .Add , , "Filename"
        .Add , , "Latest Change"
        .Add , , "Size", , lvwColumnRight
    End Select
  End With
  
  Form_Resize
  UpdateList
End Sub

Private Sub UpdateList()
  On Error GoTo ErrHandler
  Dim List As ListItem, X As Long
  Dim tStr(3) As String 'Keeps track of longest string for column autosize
  
  bWorking = True
  
  'LockWindowUpdate lvwMain.hwnd
  lvwMain.ListItems.Clear
  
  For X = 1 To UBound(AddList) 'add all found files to list
    Set List = lvwMain.ListItems.Add
    With List
      
      Select Case Settings.AddView
        Case 2 'tag view
          .Text = AddList(X).sArtistTitle
          .SubItems(1) = AddList(X).sAlbum
          .SubItems(2) = AddList(X).sGenre
          .SubItems(3) = AddList(X).sLength
          If Settings.AddAutosize Then 'refresh colums sizes
            If Len(.Text) > Len(tStr(0)) Then tStr(0) = .Text
            If Len(.SubItems(1)) > Len(tStr(1)) Then tStr(1) = .SubItems(1)
            If Len(.SubItems(2)) > Len(tStr(2)) Then tStr(2) = .SubItems(2)
            If Len(.SubItems(3)) > Len(tStr(3)) Then tStr(3) = .SubItems(3)
          End If
        Case 1 'mix view
          .Text = sFilename(AddList(X).sFilename, efpFileNameAndExt)
          .SubItems(1) = AddList(X).sLength
          If Settings.AddAutosize Then 'refresh column sizes
            If Len(.Text) > Len(tStr(0)) Then tStr(0) = .Text
            If Len(.SubItems(1)) > Len(tStr(1)) Then tStr(1) = .SubItems(1)
          End If
        Case Else 'file view
          .Text = sFilename(AddList(X).sFilename, efpFileNameAndExt)
          .SubItems(1) = AddList(X).dDate
          .SubItems(2) = FormatNumber(AddList(X).lSize, 0) & " kB"
          If Settings.AddAutosize Then 'refresh colums sizes
            If Len(.Text) > Len(tStr(0)) Then tStr(0) = .Text
            If Len(.SubItems(1)) > Len(tStr(1)) Then tStr(1) = .SubItems(1)
            If Len(.SubItems(2)) > Len(tStr(2)) Then tStr(2) = .SubItems(2)
          End If
      End Select
      
      If AddList(X).eType = TYPE_ASF Or AddList(X).eType = TYPE_MP2_MP3 Or AddList(X).eType = TYPE_OGG Or AddList(X).eType = TYPE_WAV Or AddList(X).eType = TYPE_WMA Then
        .SmallIcon = 1
      ElseIf AddList(X).eType = TYPE_IT Or AddList(X).eType = TYPE_MOD Or AddList(X).eType = TYPE_S3M Or AddList(X).eType = TYPE_XM Then
        .SmallIcon = 2
      ElseIf AddList(X).eType = TYPE_MID_RMI Or AddList(X).eType = TYPE_SGM Then
        .SmallIcon = 3
      Else
        .SmallIcon = 4
      End If
      .Tag = X
    
    End With
    
    'setting bworking to false will abort any operation
    If Not bWorking Then Exit Sub 'abort
    
  Next X
  
  If Settings.AddAutosize Then 'autosize columns
    lvwMain.ColumnHeaders(1).Width = Me.TextWidth(tStr(0)) + 480
    lvwMain.ColumnHeaders(2).Width = Me.TextWidth(tStr(1)) + 240
    If Settings.AddView <> 1 Then lvwMain.ColumnHeaders(3).Width = Me.TextWidth(tStr(2)) + 240
    If Settings.AddView = 2 Then lvwMain.ColumnHeaders(4).Width = Me.TextWidth(tStr(3)) + 240
  Else
    Form_Resize
  End If
  
  'LockWindowUpdate 0
  
  bWorking = False
  
  Exit Sub
ErrHandler:
  If cLog.ErrorMsg(Err, "frmAdd, UpdateList()") = vbNo Then frmMain.UnloadAll Else Resume Next
End Sub

Private Sub AddSelected()
  'adds selected files to playlist
  On Error GoTo ErrHandler
  Dim X As Long, AddNew As Boolean, lc As Long

  bWorking = True

  'If playlist is empty, set addnew to true, which starts the first added song later
  If UBound(Playlist) = 0 Then AddNew = True
  
  Me.MousePointer = vbArrowHourglass

  prbMain.Value = 0
  prbMain.Visible = True
  prbMain.Max = UBound(AddList)
  For X = 1 To lvwMain.ListItems.Count
    If lvwMain.ListItems(X).Selected Then  'make sure it is checked
      prbMain.Value = X
      stbMain.SimpleText = "Adding file... (" & sFilename(AddList(X).sFilename, efpFileNameAndExt) & ")"
      DoEvents
      
      lc = lc + 1
      
      'Loads item into playlist array
      If AddList(lvwMain.ListItems(X).Tag).eType = TYPE_PLAYLIST Then
        HandlePlaylist AddList(lvwMain.ListItems(X).Tag).sFilename
      Else
        CreateLibrary AddList(lvwMain.ListItems(X).Tag).sFilename
      End If

    End If
    
    'setting bworking to false will abort any operation
    If Not bWorking Then Exit Sub 'abort
    
  Next X
  
  prbMain.Visible = False
  stbMain.SimpleText = lc & " item(s) added to playlist."
  
  modMain.UpdateList
  
  If AddNew And Not tmrShutdown.Enabled Then frmMain.PlayNext
  
  Me.MousePointer = vbDefault
  bWorking = False
  
  Exit Sub
ErrHandler:
  If cLog.ErrorMsg(Err, "frmAdd, AddSelected()") = vbNo Then frmMain.UnloadAll Else Resume Next
End Sub

Private Sub AddAll()
  'adds all files to playlist
  On Error GoTo ErrHandler
  Dim X As Long, AddNew As Boolean
  
  If UBound(AddList) > 0 Then 'Make sure thare are items
        
    bWorking = True
    
    If UBound(Playlist) = 0 Then AddNew = True
    Me.MousePointer = vbArrowHourglass
    
    prbMain.Value = 0
    prbMain.Visible = True
    prbMain.Max = UBound(AddList)
    For X = 1 To UBound(AddList)
      prbMain.Value = X
      stbMain.SimpleText = "Adding file " & X & "/" & UBound(AddList) & "... (" & sFilename(AddList(X).sFilename, efpFileNameAndExt) & ")"
      DoEvents
      
      'Loads item into playlist array
      If AddList(X).eType = TYPE_PLAYLIST Then
        HandlePlaylist AddList(X).sFilename
      Else
        CreateLibrary AddList(X).sFilename
      End If
      
      'setting bworking to false will abort any operation
      If Not bWorking Then Exit Sub 'abort
      
    Next X
    
    prbMain.Visible = False
    stbMain.SimpleText = UBound(AddList) & " item(s) added to playlist."
    
    modMain.UpdateList

    If AddNew And Not tmrShutdown.Enabled Then frmMain.PlayNext
    
    Me.MousePointer = vbDefault
    bWorking = False
    
  End If
  
  Exit Sub
ErrHandler:
  If cLog.ErrorMsg(Err, "frmAdd, AddAll()") = vbNo Then frmMain.UnloadAll Else Resume Next
End Sub

Private Sub AddDirs(ByVal sFldr As String)
  'adds files in current dir and all it's subdirs
  On Error GoTo ErrHandler
  Dim File As New clsFind
  Dim AddNew As Boolean
  Dim sFilter As String, X As Long
  
  bWorking = True
  Me.MousePointer = vbHourglass
  stbMain.SimpleText = "Finding files..."
  sFldr = sAppend(sFldr, "\")
  DoEvents
  
  'Figure out the filter
  'playlists are not read in this mode
  With Settings
    If .Filter(0) Then sFilter = sFilter + "*.wav;"
    If .Filter(1) Then sFilter = sFilter + "*.mp2;*.mp3;"
    If .Filter(2) Then sFilter = sFilter + "*.ogg;"
    If .Filter(3) Then sFilter = sFilter + "*.wma;"
    If .Filter(4) Then sFilter = sFilter + "*.asf;"
    If .Filter(5) Then sFilter = sFilter + "*.mod;"
    If .Filter(6) Then sFilter = sFilter + "*.s3m;"
    If .Filter(7) Then sFilter = sFilter + "*.xm;"
    If .Filter(8) Then sFilter = sFilter + "*.it;"
    If .Filter(9) Then sFilter = sFilter + "*.mid;*.rmi;*.midi;"
    If .Filter(10) Then sFilter = sFilter + "*.sgm;"
    If Len(sFilter) = 0 Then sFilter = "*.mp3;*.mp2;"
    sFilter = Left(sFilter, Len(sFilter) - 1) 'Remove ending ';'
  End With

  'Do the search
  File.Find sFldr, sFilter, True
  
  If File.Count > 0 Then
    If UBound(Playlist) = 0 Then AddNew = True
      
    prbMain.Value = 0
    prbMain.Visible = True
    prbMain.Max = File.Count
    For X = 1 To File.Count
      prbMain.Value = X
      stbMain.SimpleText = "Adding file " & X & "/" & File.Count & "... (" & File(X).sNameAndExtension & ")"
      DoEvents

      CreateLibrary File(X).sFilename 'create in library & add to playlist
      
      'setting bworking to false will abort any operation
      If Not bWorking Then Exit Sub 'abort
      
    Next X
       
    modMain.UpdateList
    If AddNew And Not tmrShutdown.Enabled Then frmMain.PlayNext
    
  End If
 
  prbMain.Visible = False
  stbMain.SimpleText = File.Count & " item(s) added to playlist."
  Me.MousePointer = vbDefault
  bWorking = False
  
  Exit Sub
ErrHandler:
  If cLog.ErrorMsg(Err, "frmAdd, AddDirs()") = vbNo Then frmMain.UnloadAll Else Resume Next
End Sub

Public Sub RemoveFile()
  'remove file with window shell dialog box... nifty, eh?
  On Error GoTo errh
  Dim fos As SHFILEOPSTRUCT  ' structure to pass to the function
  Dim sa(1 To 32) As Byte    ' byte array to make structure properly sized
  Dim X As Long
  Dim sFrom As String
  
  'get all selected files
  For X = 1 To lvwMain.ListItems.Count
    If lvwMain.ListItems(X).Selected Then
      sFrom = sFrom & AddList(lvwMain.ListItems(X).Tag).sFilename & vbNullChar
    End If
  Next X
  
  With fos
    .hwnd = Me.hwnd
    .wFunc = FO_DELETE 'Delete the specified files.
    ' The list of files to delete.
    .pFrom = sFrom & vbNullChar
    .pTo = vbNullChar & vbNullChar
    .fFlags = FOF_ALLOWUNDO Or FOF_FILESONLY
    .fAnyOperationsAborted = 0
    .hNameMappings = 0
    .lpszProgressTitle = vbNullChar
  End With
  
  ' Transfer the contents of the structure object into the byte
  ' array in order to compensate for a byte alignment problem.
  CopyMemory sa(1), fos, LenB(fos)
  CopyMemory sa(19), sa(21), 12
  
  X = SHFileOperation(sa(1)) 'do it
  
  ' Transfer the contents of the byte array to structure object
  ' so we can get info of operation aborted.
  ''CopyMemory sa(21), sa(19), 12
  ''CopyMemory fos, sa(1), Len(fos)
  
  ''If fos.fAnyOperationsAborted = 0 Then UpdateFolder ftwMain.SelectedFolder
  
  
  UpdateFolder ftwMain.SelectedFolder
  
  Exit Sub
errh:
  If cLog.ErrorMsg(Err, "frmAdd, RemoveFile()") = vbNo Then frmMain.UnloadAll Else Resume Next
End Sub

Public Sub SelectArtist()
  On Error Resume Next
  Dim tStr As String, X As Long
  
  tStr = Trim(LCase(Left(AddList(lvwMain.SelectedItem.Tag).sArtistTitle, InStr(1, AddList(lvwMain.SelectedItem.Tag).sArtistTitle, "-"))))
  
  If tStr <> "" Then
    For X = 1 To UBound(AddList)
      If Trim(LCase(Left(AddList(X).sArtistTitle, InStr(1, AddList(X).sArtistTitle, "-")))) = tStr Then
        lvwMain.ListItems(X).Selected = True
      End If
    Next
  End If
  
End Sub

Public Sub SelectAlbum()
  On Error Resume Next
  Dim tStr As String, X As Long
  
  tStr = AddList(lvwMain.SelectedItem.Tag).sAlbum
  
  If tStr <> "" Then
    For X = 1 To UBound(AddList)
      If AddList(X).sAlbum = tStr Then lvwMain.ListItems(X).Selected = True
    Next
  Else
    stbMain.SimpleText = "Selected item has no tags or an empty album field."
  End If
  
End Sub

Public Sub SelectGenre()
  On Error Resume Next
  Dim tStr As String, X As Long
  
  tStr = AddList(lvwMain.SelectedItem.Tag).sGenre
  
  If tStr <> "" Then
    For X = 1 To UBound(AddList)
      If AddList(X).sGenre = tStr Then lvwMain.ListItems(X).Selected = True
    Next
  Else
    stbMain.SimpleText = "Selected item has no tags or an empty genre field."
  End If
  
End Sub

Public Sub SelectType()
  On Error Resume Next
  Dim tStr As Integer, X As Long
  
  tStr = AddList(lvwMain.SelectedItem.Tag).eType
  
  If tStr <> 0 Then
    For X = 1 To UBound(AddList)
      If AddList(X).eType = tStr Then lvwMain.ListItems(X).Selected = True
    Next
  End If
  
End Sub

Private Sub tmrShutdown_Timer()
  'this timer gets started whenever the window is working and you try to close it.
  'The window will then be hidden and this timer will activated to check
  'if it is safe to unload the window.
  'so now you don't have to wait for the window to finish adding files etc.
  On Error Resume Next
  If Not bWorking Then
    Unload Me 'done with work, close self
  Else
    modMain.UpdateList 'refresh playlist
  End If
End Sub

Public Sub ShowInfo()
  On Error Resume Next
  If AddList(lvwMain.SelectedItem.Tag).lItem = 0 Then
    AddList(lvwMain.SelectedItem.Tag).lItem = CreateLibrary(AddList(lvwMain.SelectedItem.Tag).sFilename, False)
  End If
  Select Case AddList(lvwMain.SelectedItem.Tag).eType
    Case TYPE_CDA, TYPE_UNKNOWN, TYPE_PLAYLIST
      MsgBox "This file type is not supported by the file viewer.", vbExclamation, "Type not supported"
    Case Else
      frmView.View AddList(lvwMain.SelectedItem.Tag).lItem
  End Select
End Sub

Public Sub PlaySelected()
  On Error Resume Next
  If AddList(lvwMain.SelectedItem.Tag).lItem = 0 Then
    AddList(lvwMain.SelectedItem.Tag).lItem = CreateLibrary(AddList(lvwMain.SelectedItem.Tag).sFilename, False)
  End If
  Playing = 0
  PlayingLib = AddList(lvwMain.SelectedItem.Tag).lItem
  frmMain.Play
End Sub
