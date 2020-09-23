VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLibrary 
   Caption         =   "Media Library Browser"
   ClientHeight    =   7275
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10050
   ClipControls    =   0   'False
   Icon            =   "frmLibrary.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tbrFilter 
      Height          =   330
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   582
      ButtonWidth     =   2223
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Filter Artist:"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwFilter 
      Height          =   2655
      Left            =   60
      TabIndex        =   6
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   4683
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrSettings 
      Height          =   330
      Left            =   8880
      TabIndex        =   5
      ToolTipText     =   "Settings"
      Top             =   60
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
      ButtonWidth     =   1746
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Settings"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbSearch 
      Height          =   315
      Left            =   3600
      TabIndex        =   1
      Top             =   60
      Width           =   5175
   End
   Begin MSComctlLib.ProgressBar prbMain 
      Height          =   255
      Left            =   8160
      TabIndex        =   3
      Top             =   6990
      Width           =   1575
      _ExtentX        =   2778
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
      Top             =   6900
      Width           =   10050
      _ExtentX        =   17727
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
   Begin MSComctlLib.ImageList imlMain 
      Left            =   5280
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":0C1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibrary.frx":0F72
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   5895
      Left            =   2640
      TabIndex        =   0
      Top             =   480
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   10398
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Artist & Title"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Album"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Genre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Last Playing Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrSearch 
      Height          =   330
      Left            =   2520
      TabIndex        =   4
      ToolTipText     =   "Search Filter"
      Top             =   60
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
      ButtonWidth     =   1693
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search:"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu menf 
      Caption         =   "f1"
      Visible         =   0   'False
      Begin VB.Menu menfo 
         Caption         =   "&Artist"
         Index           =   0
      End
      Begin VB.Menu menfo 
         Caption         =   "&Album"
         Index           =   1
      End
      Begin VB.Menu menfo 
         Caption         =   "&Genre"
         Index           =   2
      End
      Begin VB.Menu menfo 
         Caption         =   "&Year"
         Index           =   3
      End
   End
   Begin VB.Menu mens 
      Caption         =   "s"
      Visible         =   0   'False
      Begin VB.Menu menso 
         Caption         =   "&Artist"
         Index           =   0
      End
      Begin VB.Menu menso 
         Caption         =   "&Title"
         Index           =   1
      End
      Begin VB.Menu menso 
         Caption         =   "&Album"
         Index           =   2
      End
      Begin VB.Menu menso 
         Caption         =   "&Comments"
         Index           =   3
      End
      Begin VB.Menu menso 
         Caption         =   "&Filename"
         Index           =   4
      End
   End
   Begin VB.Menu menMenu 
      Caption         =   "m"
      Visible         =   0   'False
      Begin VB.Menu menPlay 
         Caption         =   "&Play Focused Item"
      End
      Begin VB.Menu menInfo 
         Caption         =   "&File Info..."
      End
      Begin VB.Menu menline 
         Caption         =   "-"
      End
      Begin VB.Menu menAddSel 
         Caption         =   "Add &Selected Items"
      End
      Begin VB.Menu menAddAll 
         Caption         =   "Add &All to Playlist"
      End
      Begin VB.Menu menClear 
         Caption         =   "&Clear Playlist"
      End
      Begin VB.Menu ml1 
         Caption         =   "-"
      End
      Begin VB.Menu menRefList 
         Caption         =   "&Refresh List"
      End
      Begin VB.Menu menReload 
         Caption         =   "R&eload Library"
      End
   End
End
Attribute VB_Name = "frmLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public LibClose As Boolean

'these arrays contains the items in the sorting list
Dim Artist() As String
Dim Album() As String
Dim Genre() As String
Dim Year() As String

Private Type ExData
  Artist As String
  Title As String
  bShow As Boolean 'True if format is mp3,mp2,ogg,wma,asf
End Type

Private Mousepos As POINTAPI 'for doubleclick in lvwMain

Dim ExData() As ExData 'holds extra data used to speed up
Dim NumEx As Long

Option Explicit

Private Sub cmbSearch_KeyUp(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  If KeyCode = vbKeyReturn Then
    Dim X As Long, A As Boolean
    'Add current search text to list in combobox, but only if it doesn't exist already
    For X = 0 To cmbSearch.ListCount - 1
      If cmbSearch.Text = cmbSearch.List(X) Then
        A = True
        Exit For
      End If
    Next
    If Not A Then cmbSearch.AddItem cmbSearch.Text
    'do search
    RefreshList
  End If
End Sub

Private Sub Form_Activate()
  On Error Resume Next
  
  AlwaysOnTop Me, Settings.OnTop
  LoadLib
End Sub

Private Sub Form_Load()
  On Error Resume Next
  
  'reset all lists
  ReDim Artist(0)
  ReDim Album(0)
  ReDim Genre(0)
  ReDim Year(0)
  ReDim ExData(0)
  
  'restore saved window settings
  If Settings.LibWidth > 0 Then Me.Width = Settings.LibWidth
  If Settings.LibHeight > 0 Then Me.Height = Settings.LibHeight
  If Settings.LibMax Then Me.WindowState = vbMaximized
  
  Me.Show
  DoEvents
  
  AlwaysOnTop Me, Settings.OnTop
  
  UpdateFilters
  DoEvents
  
  Examine
  
  DoEvents
  
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lvwMain.Move lvwFilter.Width + lvwFilter.Left + 60, _
  cmbSearch.Top + cmbSearch.Height + 60, Me.Width - lvwFilter.Width - 280, _
  Me.Height - lvwMain.Top - 540 - stbMain.Height
  
  lvwFilter.Move lvwFilter.Left, tbrFilter.Top + tbrFilter.Height + 60, _
  lvwFilter.Width, lvwMain.Height
  
  cmbSearch.Width = Me.Width - cmbSearch.Left - tbrSettings.Width - 200
  tbrSettings.Left = Me.Width - tbrSettings.Width - 150
  
  prbMain.Move Me.Width - prbMain.Width - 450, stbMain.Top + 90
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  
  'save window settings
  If Me.WindowState = vbMaximized Then
    Settings.LibMax = True
  Else
    Settings.LibMax = False
    Settings.LibWidth = Me.Width
    Settings.LibHeight = Me.Height
  End If
  
  'Because it takes so long to examine the media library each time the window
  'loads, instead of unloading it we will just hide it, so it doesn't have to
  'reload next time we start the media library. However, if we want to end the
  'program we should set LibClose to true.
  If Not LibClose Then
    Cancel = 1
    Me.Hide
  End If
End Sub

Private Sub lvwFilter_DblClick()
  'cmbSearch.Text = ""
  RefreshList
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  On Error Resume Next
  If lvwMain.SortKey = ColumnHeader.Index - 1 Then
    If lvwMain.SortOrder = lvwAscending Then
      lvwMain.SortOrder = lvwDescending
    Else
      lvwMain.SortOrder = lvwAscending
    End If
  End If
  lvwMain.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub lvwMain_DblClick()
  On Error GoTo errh
  Dim hit As ListItem
  Set hit = lvwMain.HitTest(Mousepos.X, Mousepos.y)
  
  PlayingLib = lvwMain.SelectedItem.Tag
  Playing = 0
  frmMain.Play
  
errh:
End Sub

Private Sub lvwMain_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  Mousepos.X = X
  Mousepos.y = y
End Sub

Private Sub lvwMain_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  If Button = 2 Then PopupMenu menMenu, , , , menPlay
End Sub

Private Sub menAddAll_Click()
  'add all items in list to playlist
  On Error Resume Next
  Dim X As Long
  For X = 1 To lvwMain.ListItems.Count
    ReDim Preserve Playlist(UBound(Playlist) + 1)
    Playlist(UBound(Playlist)).Reference = lvwMain.ListItems(X).Tag
    Playlist(UBound(Playlist)).Index = UBound(Playlist)
  Next
  UpdateList
End Sub

Private Sub menAddSel_Click()
  'add selected items in list to playlist
  On Error Resume Next
  Dim X As Long
  For X = 1 To lvwMain.ListItems.Count
    If lvwMain.ListItems(X).Selected Then
      ReDim Preserve Playlist(UBound(Playlist) + 1)
      Playlist(UBound(Playlist)).Reference = lvwMain.ListItems(X).Tag
      Playlist(UBound(Playlist)).Index = UBound(Playlist)
    End If
  Next
  UpdateList
End Sub

Private Sub menClear_Click()
  'clear playlist
  frmMenus.menDeleteAll_Click
End Sub

Private Sub menfo_Click(Index As Integer)
  Settings.LibFilter = Index
  UpdateFilters
  RefreshFilter
End Sub

Private Sub menInfo_Click()
  frmView.View lvwMain.SelectedItem.Tag
End Sub

Private Sub menPlay_Click()
  'play without adding to playlist
  Playing = 0
  PlayingLib = lvwMain.SelectedItem.Tag
  frmMain.Play
End Sub

Private Sub menRefList_Click()
  RefreshList
End Sub

Private Sub menReload_Click()
  LoadLib
  Examine
End Sub

Private Sub menso_Click(Index As Integer)
  Settings.LibSearch(Index) = Not CBool(Settings.LibSearch(Index))
End Sub

Public Sub UpdateFilters()
  Select Case Settings.LibFilter
    Case 0
      tbrFilter.Buttons(1).Caption = "Filter Artist:"
    Case 1
      tbrFilter.Buttons(1).Caption = "Filter Album:"
    Case 2
      tbrFilter.Buttons(1).Caption = "Filter Genre:"
    Case 3
      tbrFilter.Buttons(1).Caption = "Filter Year:"
  End Select
End Sub

Public Sub RefreshFilter()
  'updates the filter list with the selected filter's items
  On Error Resume Next
  Dim ListAdd As ListItem, X As Long, y As String

  'LockWindowUpdate lvwFilter.hWnd
  
  lvwFilter.ListItems.Clear
  lvwFilter.ListItems.Add , , " <All>"
  lvwFilter.ListItems.Add , , " <None>"
  
  Select Case Settings.LibFilter
    Case 0
      For X = 1 To UBound(Artist)
        Set ListAdd = lvwFilter.ListItems.Add
        ListAdd.Text = StrConv(Artist(X), vbProperCase)
        If Len(Artist(X)) > Len(y) Then y = Artist(X)
      Next
    Case 1
      For X = 1 To UBound(Album)
        Set ListAdd = lvwFilter.ListItems.Add
        ListAdd.Text = StrConv(Album(X), vbProperCase)
        If Len(Album(X)) > Len(y) Then y = Album(X)
      Next
    Case 2
      For X = 1 To UBound(Genre)
        Set ListAdd = lvwFilter.ListItems.Add
        ListAdd.Text = StrConv(Genre(X), vbProperCase)
        If Len(Genre(X)) > Len(y) Then y = Genre(X)
      Next
    Case 3
      For X = 1 To UBound(Year)
        Set ListAdd = lvwFilter.ListItems.Add
        ListAdd.Text = Year(X)
        If Len(Year(X)) > Len(y) Then y = Year(X)
      Next
  End Select
    
  lvwFilter.ColumnHeaders(1).Width = Me.TextWidth(y) + 512
  
  'LockWindowUpdate 0

End Sub

Public Sub RefreshList()
  'This sub refreshes the main list and applies the selected filters
  On Error Resume Next
  
  If NumEx <> UBound(LibraryIndex) Then Examine
  
  Dim ListAdd As ListItem
  Dim X As Long, y As Long, A As Boolean, b As Boolean
  Dim tStr(4) As String, sF As String
  sF = lvwFilter.SelectedItem.Text
  
  cLog.Log "APPLYING FILTERS AND REFRESHING LIST...", 3, False
  cLog.StartTimer
  
  stbMain.SimpleText = "Searching Library..."
  prbMain.Value = 0
  prbMain.Visible = True
  prbMain.Max = UBound(LibraryIndex)
  Me.MousePointer = vbArrowHourglass
  DoEvents
  
  'LockWindowUpdate lvwMain.hWnd
  
  lvwMain.ListItems.Clear
  For X = 1 To UBound(LibraryIndex)
    prbMain.Value = X
    
    If ExData(X).bShow Then
      A = False
      b = False
      
      With Library(LibraryIndex(X).lReference)
      
        'First filter
        Select Case Settings.LibFilter
          Case 0 'artist
            If sF = " <All>" Then
              A = True
            ElseIf sF = " <None>" Then
              If ExData(X).Artist = "" Then A = True
            ElseIf LCase(sF) = ExData(X).Artist Then
              A = True
            End If
            
          Case 1 'album
            If sF = " <All>" Then
              A = True
            ElseIf sF = " <None>" Then
              If .sAlbum = "" Then A = True
            ElseIf LCase(sF) = LCase(.sAlbum) Then
              A = True
            End If
            
          Case 2 'genre
            If sF = " <All>" Then
              A = True
            ElseIf sF = " <None>" Then
              If .sGenre = "" Then A = True
            ElseIf LCase(sF) = LCase(.sGenre) Then
              A = True
            End If
            
          Case 3 'year
            If sF = " <All>" Then
              A = True
            ElseIf sF = " <None>" Then
              If .sYear = "" Then A = True
            ElseIf LCase(sF) = LCase(.sYear) Then
              A = True
            End If
            
        End Select
        
        'search filter
        If Len(Trim(cmbSearch.Text)) > 0 Then
          If Settings.LibSearch(0) And InStr(1, ExData(X).Artist, cmbSearch.Text, vbTextCompare) > 0 Then b = True
          If Settings.LibSearch(1) And InStr(1, ExData(X).Title, cmbSearch.Text, vbTextCompare) > 0 Then b = True
          If Settings.LibSearch(2) And InStr(1, .sAlbum, cmbSearch.Text, vbTextCompare) > 0 Then b = True
          If Settings.LibSearch(3) And InStr(1, .sComments, cmbSearch.Text, vbTextCompare) > 0 Then b = True
          If Settings.LibSearch(4) And InStr(1, sFilename(.sFilename, efpFileName), cmbSearch.Text, vbTextCompare) > 0 Then b = True
        Else
          b = True
        End If
        
        If (A And b) Then 'add to list
          Set ListAdd = lvwMain.ListItems.Add
          ListAdd.Text = .sArtistTitle
          ListAdd.SubItems(1) = .sAlbum
          ListAdd.SubItems(2) = .sGenre
          ListAdd.SubItems(3) = ConvertTime(.lLength)
          ListAdd.SubItems(4) = .dLastPlayDate
          ListAdd.Tag = LibraryIndex(X).lReference
          
          If Settings.LibAutosize Then 'get size of columns
            If Len(ListAdd.Text) > Len(tStr(0)) Then tStr(0) = ListAdd.Text
            If Len(ListAdd.SubItems(1)) > Len(tStr(1)) Then tStr(1) = ListAdd.SubItems(1)
            If Len(ListAdd.SubItems(2)) > Len(tStr(2)) Then tStr(2) = ListAdd.SubItems(2)
            If Len(ListAdd.SubItems(3)) > Len(tStr(3)) Then tStr(3) = ListAdd.SubItems(3)
            If Len(ListAdd.SubItems(4)) > Len(tStr(4)) Then tStr(4) = ListAdd.SubItems(4)
          End If
        End If
      End With
    End If
       
  Next
  
  'resize columns
  If Settings.LibAutosize And lvwMain.ListItems.Count > 0 Then
    lvwMain.ColumnHeaders(1).Width = Me.TextWidth(tStr(0)) + 240
    lvwMain.ColumnHeaders(2).Width = Me.TextWidth(tStr(1)) + 240
    lvwMain.ColumnHeaders(3).Width = Me.TextWidth(tStr(2)) + 240
    lvwMain.ColumnHeaders(4).Width = Me.TextWidth(tStr(3)) + 240
    lvwMain.ColumnHeaders(5).Width = Me.TextWidth(tStr(4)) + 240
  End If
  
  'LockWindowUpdate 0
  
  Me.MousePointer = vbDefault
  prbMain.Visible = False
  stbMain.SimpleText = lvwMain.ListItems.Count & " item(s) found with "
  Select Case Settings.LibFilter
    Case 0: stbMain.SimpleText = stbMain.SimpleText & "Artist: "
    Case 1: stbMain.SimpleText = stbMain.SimpleText & "Album: "
    Case 2: stbMain.SimpleText = stbMain.SimpleText & "Genre: "
    Case 3: stbMain.SimpleText = stbMain.SimpleText & "Year: "
  End Select
  stbMain.SimpleText = stbMain.SimpleText & "'" & sF & "' filter selected"
  If Len(Trim(cmbSearch.Text)) > 0 Then
    stbMain.SimpleText = stbMain.SimpleText & " and search string '" & Trim(cmbSearch.Text) & "'"
  End If
  stbMain.SimpleText = stbMain.SimpleText & "."
  
  cLog.Log "DONE. (" & lvwMain.ListItems.Count & " MATCHES, " & cLog.GetTimer & " ms)", 3
    
End Sub

Private Sub tbrFilter_ButtonClick(ByVal Button As MSComctlLib.Button)
  On Error Resume Next
  ShowPopup frmLibrary, menf, tbrFilter, TextWidth(menfo(1).Caption) + (45 * Screen.TwipsPerPixelX)
End Sub

Private Sub tbrSearch_ButtonClick(ByVal Button As MSComctlLib.Button)
  On Error Resume Next
  Dim X As Integer
  For X = 0 To 4
    menso(X).Checked = CBool(Settings.LibSearch(X))
  Next
  ShowPopup frmLibrary, mens, tbrSearch, TextWidth(menso(3).Caption) + (45 * Screen.TwipsPerPixelX)
End Sub

Private Sub tbrSettings_ButtonClick(ByVal Button As MSComctlLib.Button)
  frmLibSettings.Show , Me
End Sub

Public Sub Examine()
  'Get all availabe artists, albums, genres & years
  On Error Resume Next
  
  cLog.Log "PROCESSING MEDIA LIBRARY...", 3, False
  cLog.StartTimer
  
  stbMain.SimpleText = "Processing Media Library..."
  Me.MousePointer = vbArrowHourglass
  prbMain.Value = 0
  prbMain.Max = UBound(LibraryIndex)
  prbMain.Visible = True
  DoEvents
  
  ReDim Artist(0)
  ReDim Album(0)
  ReDim Genre(0)
  ReDim Year(0)
  ReDim ExData(UBound(LibraryIndex))
  NumEx = UBound(LibraryIndex)
  
  Dim y As Long, X As Long, f As Boolean, tStr As String
  'Then examine all items to be able to get the artist/album/genre/Year data
  For X = 1 To UBound(LibraryIndex)
    prbMain.Value = X
    DoEvents
    
    With Library(LibraryIndex(X).lReference)
    
      ExData(X).bShow = True
      If Settings.LibHideNonMusic Then
        If (.eType = TYPE_IT Or .eType = TYPE_MID_RMI Or .eType = TYPE_MOD Or .eType = TYPE_S3M Or .eType = TYPE_SGM Or .eType = TYPE_WAV Or .eType = TYPE_XM) Then
          ExData(X).bShow = False
        End If
      End If
      
      If ExData(X).bShow Then
      
        'get artist & title from tags or from filename
        ExData(X).Artist = LCase(.sArtist)
        ExData(X).Title = LCase(.sTitle)
        If ExData(X).Artist = "" And Settings.LibGetName And InStr(1, sFilename(.sFilename, efpFileName), "-") > 0 Then
          ExData(X).Artist = sFilename(.sFilename, efpFileName)
          ExData(X).Artist = LCase(Trim(Left(ExData(X).Artist, InStr(1, ExData(X).Artist, "-") - 1)))
        End If
        If ExData(X).Title = "" And Settings.LibGetName And InStr(1, sFilename(.sFilename, efpFileName), "-") > 0 Then
          ExData(X).Title = sFilename(.sFilename, efpFileName)
          ExData(X).Title = LCase(Trim(Right(ExData(X).Title, InStrRev(ExData(X).Title, "-"))))
        End If

        'make sure there are no duplicates
        'Artist first
        f = False
        tStr = ExData(X).Artist
        If tStr <> "" Then
          For y = 1 To UBound(Artist)
            If tStr = Artist(y) Then
              f = True
              Exit For
            End If
          Next
          If Not f Then  'If artist was not found in list
            ReDim Preserve Artist(UBound(Artist) + 1)
            Artist(UBound(Artist)) = tStr
          End If
        End If
        
        'then album
        f = False
        tStr = LCase(.sAlbum)  'buffered lcase to speed up
        If tStr <> "" Then
          For y = 1 To UBound(Album)
            If tStr = Album(y) Then
              f = True
              Exit For
            End If
          Next
          If Not f Then  'If album was not found in list
            ReDim Preserve Album(UBound(Album) + 1)
            Album(UBound(Album)) = tStr
          End If
        End If
        
        'then genre
        f = False
        tStr = LCase(.sGenre)  'buffered lcase to speed up
        If tStr <> "" Then
          For y = 1 To UBound(Genre)
            If tStr = Genre(y) Then
              f = True
              Exit For
            End If
          Next
          If Not f Then  'If Genre was not found in list
            ReDim Preserve Genre(UBound(Genre) + 1)
            Genre(UBound(Genre)) = tStr
          End If
        End If
        
        'and last year
        f = False
        tStr = LCase(.sYear)  'buffered lcase to speed up
        If tStr <> "" Then
          For y = 1 To UBound(Year)
            If tStr = Year(y) Then
              f = True
              Exit For
            End If
          Next
          If Not f Then  'If Genre was not found in list
            ReDim Preserve Year(UBound(Year) + 1)
            Year(UBound(Year)) = tStr
          End If
        End If
      
      End If
    End With
    
  Next X
  
  stbMain.SimpleText = ""
  prbMain.Visible = False
  Me.MousePointer = vbDefault
  
  cLog.Log "DONE. (" & cLog.GetTimer & " ms)", 3
  
  RefreshFilter
  RefreshList
  
End Sub

Public Sub LoadLib()
  'Makes sure the entire media library is loaded
  On Error Resume Next
  
  If LoadedMedia < UBound(LibraryIndex) Then 'dont load when everything is loaded already
    
    cLog.Log "LOADING ENTIRE MEDIA LIBRARY...(ALREADY LOADED: " & LoadedMedia & "/" & UBound(LibraryIndex) & ")...", 3, False
    cLog.StartTimer
  
    Dim X As Long
  
    Me.MousePointer = vbHourglass
    stbMain.SimpleText = "Loading Media Library..."
    prbMain.Value = 0
    prbMain.Max = UBound(LibraryIndex)
    prbMain.Visible = True
    
    For X = 1 To UBound(LibraryIndex)
      prbMain.Value = X
      DoEvents
      If LibraryIndex(X).lReference = 0 Then
        'there is no reference to Library array which means it has
        'not been loaded yet, so load it!
        ReDim Preserve Library(UBound(Library) + 1)
        LibraryIndex(X).lReference = UBound(Library)
        LoadItem UBound(Library), X
      End If
    Next
    
    LoadedMedia = UBound(LibraryIndex)
    
    stbMain.SimpleText = ""
    Me.MousePointer = vbDefault
    prbMain.Visible = False
    
    cLog.Log "DONE. (" & cLog.GetTimer & " ms)", 3
    
  End If
  
End Sub
