VERSION 5.00
Begin VB.Form frmMenus 
   Caption         =   "Menus"
   ClientHeight    =   1155
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4095
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1155
   ScaleWidth      =   4095
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "menus cannot be placed on frmMain and frmPlaylist when they have no borders..."
      Height          =   585
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      WordWrap        =   -1  'True
   End
   Begin VB.Menu menAddRClick 
      Caption         =   "AddRClick"
      Begin VB.Menu menPlayfoca 
         Caption         =   "&Play Focused Item"
      End
      Begin VB.Menu menFinfo 
         Caption         =   "&File Info..."
      End
      Begin VB.Menu menaddline 
         Caption         =   "-"
      End
      Begin VB.Menu menSelAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu menSelNone 
         Caption         =   "Select &None"
      End
      Begin VB.Menu menSelArtist 
         Caption         =   "Select all from this A&rtist"
      End
      Begin VB.Menu menSelAlbum 
         Caption         =   "Select all from this A&lbum"
      End
      Begin VB.Menu menSelGenre 
         Caption         =   "Select all of this &Genre"
      End
      Begin VB.Menu menSelType 
         Caption         =   "Select all of this &Type"
      End
      Begin VB.Menu menaddline2 
         Caption         =   "-"
      End
      Begin VB.Menu menClrPlist 
         Caption         =   "&Clear Playlist"
      End
      Begin VB.Menu menFilter 
         Caption         =   "F&ilter"
         Begin VB.Menu menAllSupp 
            Caption         =   "All Supported Types"
         End
         Begin VB.Menu menTypeAllNoPlay 
            Caption         =   "All Except Playlists"
         End
         Begin VB.Menu menTypeStream 
            Caption         =   "Toggle Streaming Types"
         End
         Begin VB.Menu menTypeSeq 
            Caption         =   "Toggle Sequenced Types"
         End
         Begin VB.Menu menlline 
            Caption         =   "-"
         End
         Begin VB.Menu menFilt 
            Caption         =   "&Waveform Audio Files (*.wav)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu menFilt 
            Caption         =   "&MPEG Audio Files (*.mp2;*.mp3)"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu menFilt 
            Caption         =   "&Ogg Files (*.ogg)"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu menFilt 
            Caption         =   "Windows Media &Audio Files (*.wma)"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu menFilt 
            Caption         =   "Advanced &Sound Format Files (*.asf)"
            Checked         =   -1  'True
            Index           =   4
         End
         Begin VB.Menu menFilt 
            Caption         =   "Protracker/Fasttracker &Modules (*.mod)"
            Checked         =   -1  'True
            Index           =   5
         End
         Begin VB.Menu menFilt 
            Caption         =   "Screamtracker &3 Modules (*.s3m)"
            Checked         =   -1  'True
            Index           =   6
         End
         Begin VB.Menu menFilt 
            Caption         =   "Fasttracker 2 Modules (*.&xm)"
            Checked         =   -1  'True
            Index           =   7
         End
         Begin VB.Menu menFilt 
            Caption         =   "&Impulse Tracker Modules (*.it)"
            Checked         =   -1  'True
            Index           =   8
         End
         Begin VB.Menu menFilt 
            Caption         =   "MI&DI Files (*.mid;*.rmi)"
            Checked         =   -1  'True
            Index           =   9
         End
         Begin VB.Menu menFilt 
            Caption         =   "Di&rectMusic Segment Files (*.sgm)"
            Checked         =   -1  'True
            Index           =   10
         End
         Begin VB.Menu menFilt 
            Caption         =   "Supported &Playlist Files (*.playlist;*.pls;*.m3u)"
            Checked         =   -1  'True
            Index           =   11
         End
      End
      Begin VB.Menu menAddViews 
         Caption         =   "&View"
         Begin VB.Menu menView 
            Caption         =   "&File View"
            Index           =   0
         End
         Begin VB.Menu menView 
            Caption         =   "&Mixed View"
            Index           =   1
         End
         Begin VB.Menu menView 
            Caption         =   "&Tag View"
            Index           =   2
         End
         Begin VB.Menu ippo 
            Caption         =   "-"
         End
         Begin VB.Menu menAutosizeColumns 
            Caption         =   "&Autosize Columns"
         End
      End
      Begin VB.Menu addline3 
         Caption         =   "-"
      End
      Begin VB.Menu menRemoveFile 
         Caption         =   "&Remove File(s)..."
      End
   End
   Begin VB.Menu menRClick 
      Caption         =   "RClickMenu"
      Begin VB.Menu menPlay 
         Caption         =   "&Play Item"
      End
      Begin VB.Menu menFileInfo 
         Caption         =   "&File Info..."
      End
      Begin VB.Menu linje1 
         Caption         =   "-"
      End
      Begin VB.Menu menAdd 
         Caption         =   "&Add"
         Begin VB.Menu menAddWin 
            Caption         =   "&File Browser..."
         End
         Begin VB.Menu menLibrary 
            Caption         =   "&Media Library Browser..."
         End
         Begin VB.Menu menAudioCD 
            Caption         =   "&Audio CD (default drive)"
         End
      End
      Begin VB.Menu menDel 
         Caption         =   "&Delete"
         Begin VB.Menu menDeleteFile 
            Caption         =   "&Item(s)"
         End
         Begin VB.Menu menDeleteAll 
            Caption         =   "&All"
         End
         Begin VB.Menu menCrop 
            Caption         =   "&Crop"
         End
      End
      Begin VB.Menu menSelect 
         Caption         =   "&Select"
         Begin VB.Menu menSelectAll 
            Caption         =   "&All"
         End
         Begin VB.Menu menSelectNone 
            Caption         =   "&None"
         End
         Begin VB.Menu menSelectInvert 
            Caption         =   "&Invert"
         End
         Begin VB.Menu menSelPlaying 
            Caption         =   "&Playing Item"
         End
      End
      Begin VB.Menu menList 
         Caption         =   "&List"
         Begin VB.Menu menListLoad 
            Caption         =   "&Load Playlist..."
         End
         Begin VB.Menu menListSave 
            Caption         =   "&Save Playlist..."
         End
         Begin VB.Menu menPlistSearch 
            Caption         =   "&Find..."
         End
         Begin VB.Menu menSort 
            Caption         =   "Sort List"
            Begin VB.Menu menSortArtistTitle 
               Caption         =   "By Artist && &Title"
            End
            Begin VB.Menu menSortAlbum 
               Caption         =   "By &Album"
            End
            Begin VB.Menu menSortGenre 
               Caption         =   "By &Genre"
            End
            Begin VB.Menu menSortTime 
               Caption         =   "By Ti&me"
            End
            Begin VB.Menu menSortFilename 
               Caption         =   "By &Filename"
            End
            Begin VB.Menu menSortType 
               Caption         =   "By Filety&pe"
            End
            Begin VB.Menu menSortDate 
               Caption         =   "By Play&date"
            End
            Begin VB.Menu menTimesPlayed 
               Caption         =   "By &Times Started"
            End
            Begin VB.Menu menTimesSkipped 
               Caption         =   "By Times &Skipped"
            End
            Begin VB.Menu menReset 
               Caption         =   "By &Original Order"
            End
            Begin VB.Menu op 
               Caption         =   "-"
            End
            Begin VB.Menu menReverse 
               Caption         =   "&Reverse List"
            End
         End
      End
   End
   Begin VB.Menu menMain 
      Caption         =   "MainMenu"
      Begin VB.Menu menAbout 
         Caption         =   "About Simple Amp..."
      End
      Begin VB.Menu menHelp 
         Caption         =   "&Help..."
      End
      Begin VB.Menu menLine 
         Caption         =   "-"
      End
      Begin VB.Menu menPlayback 
         Caption         =   "&Playback"
         Begin VB.Menu menPrev 
            Caption         =   "&Previous"
         End
         Begin VB.Menu menPlayPause 
            Caption         =   "P&lay/Pause"
         End
         Begin VB.Menu menStop 
            Caption         =   "&Stop"
         End
         Begin VB.Menu menNext 
            Caption         =   "&Next"
         End
         Begin VB.Menu menlinjegram 
            Caption         =   "-"
         End
         Begin VB.Menu menFwd 
            Caption         =   "&Forward 5 Sec."
         End
         Begin VB.Menu menBck 
            Caption         =   "&Back 5 Sec."
         End
         Begin VB.Menu menenlinje 
            Caption         =   "-"
         End
         Begin VB.Menu menVolUp 
            Caption         =   "&Raise Volume"
         End
         Begin VB.Menu menVolDown 
            Caption         =   "&Lower Volume"
         End
      End
      Begin VB.Menu menAdvance 
         Caption         =   "Auto &Advance"
      End
      Begin VB.Menu menRepeat 
         Caption         =   "&Repeat"
         Begin VB.Menu menRepeatVal 
            Caption         =   "&Off"
            Index           =   0
         End
         Begin VB.Menu menRepeatVal 
            Caption         =   "&List"
            Index           =   1
         End
         Begin VB.Menu menRepeatVal 
            Caption         =   "&Song"
            Index           =   2
         End
      End
      Begin VB.Menu menShuffle 
         Caption         =   "S&huffle"
         Begin VB.Menu menShuffleVal 
            Caption         =   "&Off"
            Index           =   0
         End
         Begin VB.Menu menShuffleVal 
            Caption         =   "Ord&ered"
            Index           =   1
         End
         Begin VB.Menu menShuffleVal 
            Caption         =   "&Random"
            Index           =   2
         End
         Begin VB.Menu menk 
            Caption         =   "-"
         End
         Begin VB.Menu menShufReset 
            Caption         =   "Reset Shuffle Order"
         End
      End
      Begin VB.Menu menPlaylist 
         Caption         =   "&Playlist"
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu menMBrowse 
         Caption         =   "&File Browser..."
      End
      Begin VB.Menu menMMediaLib 
         Caption         =   "&Media Library Browser..."
      End
      Begin VB.Menu bubba 
         Caption         =   "-"
      End
      Begin VB.Menu menVisLoad 
         Caption         =   "Load Visualization..."
      End
      Begin VB.Menu menVisEdit 
         Caption         =   "Edit Visualization..."
      End
      Begin VB.Menu hubba 
         Caption         =   "-"
      End
      Begin VB.Menu menSoundStudio 
         Caption         =   "S&ound Studio..."
      End
      Begin VB.Menu menSkin 
         Caption         =   "&Change Skin..."
      End
      Begin VB.Menu menSettings 
         Caption         =   "&Settings..."
      End
      Begin VB.Menu menline2 
         Caption         =   "-"
      End
      Begin VB.Menu menExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu menVisMenu 
      Caption         =   "VisMenu"
      Begin VB.Menu menTurnOff 
         Caption         =   "&Turn off"
      End
      Begin VB.Menu menNextSkinPreset 
         Caption         =   "&Next Skin Preset"
      End
      Begin VB.Menu menLinjal 
         Caption         =   "-"
      End
      Begin VB.Menu menPresetLoad 
         Caption         =   "&Load Preset..."
      End
      Begin VB.Menu menPresetEdit 
         Caption         =   "&Edit Preset..."
      End
   End
End
Attribute VB_Name = "frmMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  'setup menu captions to reflect keyboard commands.
  'vb menu editor wont let you type tabs...
  'but you have to make sure the commands actually work yourself.
  'most ctrl+ commands can be done automatically by vb, but
  'that doesn't work in this case, when i am using these menus
  'on other forms!
  
  'playlist rightclick
  menPlay.Caption = menPlay.Caption & vbTab & "Enter"
  menAddWin.Caption = menAddWin.Caption & vbTab & "F2"
  menLibrary.Caption = menLibrary.Caption & vbTab & "F3"
  menDeleteFile.Caption = menDeleteFile.Caption & vbTab & "Delete"
  menSelectAll.Caption = menSelectAll.Caption & vbTab & "Ctrl+A"
  menSelectNone.Caption = menSelectNone.Caption & vbTab & "Ctrl+N"
  menSelectInvert.Caption = menSelectInvert.Caption & vbTab & "Ctrl+I"
  menSelPlaying.Caption = menSelPlaying.Caption & vbTab & "H"
  menPlistSearch.Caption = menPlistSearch.Caption & vbTab & "Ctrl+F"
  menSortArtistTitle.Caption = menSortArtistTitle.Caption & vbTab & "1"
  menSortAlbum.Caption = menSortAlbum.Caption & vbTab & "2"
  menSortGenre.Caption = menSortGenre.Caption & vbTab & "3"
  menSortTime.Caption = menSortTime.Caption & vbTab & "4"
  menSortFilename.Caption = menSortFilename.Caption & vbTab & "5"
  menSortType.Caption = menSortType.Caption & vbTab & "6"
  menSortDate.Caption = menSortDate.Caption & vbTab & "7"
  menTimesPlayed.Caption = menTimesPlayed.Caption & vbTab & "8"
  menTimesSkipped.Caption = menTimesSkipped.Caption & vbTab & "9"
  menReset.Caption = menReset.Caption & vbTab & "0"
  
  'main menu
  menHelp.Caption = menHelp.Caption & vbTab & "F1"
  menAdvance.Caption = menAdvance.Caption & vbTab & "A"
  menRepeat.Caption = menRepeat.Caption & vbTab & "R"
  menShuffle.Caption = menShuffle.Caption & vbTab & "S"
  menPlaylist.Caption = menPlaylist.Caption & vbTab & "P"
  menPrev.Caption = menPrev.Caption & vbTab & "Z"
  menPlayPause.Caption = menPlayPause.Caption & vbTab & "X"
  menStop.Caption = menStop.Caption & vbTab & "C"
  menNext.Caption = menNext.Caption & vbTab & "V"
  menFwd.Caption = menFwd.Caption & vbTab & "Right"
  menBck.Caption = menBck.Caption & vbTab & "Left"
  menVolUp.Caption = menVolUp.Caption & vbTab & "Up"
  menVolDown.Caption = menVolDown.Caption & vbTab & "Down"
  menMBrowse.Caption = menMBrowse.Caption & vbTab & "F2"
  menMMediaLib.Caption = menMMediaLib.Caption & vbTab & "F3"
  
  'Set menus to have radio checkbuttons
  Dim hMenu As Long, hSubMenu As Long
  hMenu = modMenu.GetMenuHandle(Me.hwnd) 'get window menu handle
  hMenu = modMenu.GetSubMenuHandle(hMenu, 2) 'get MainMenu handle
  hSubMenu = modMenu.GetSubMenuHandle(hMenu, 5) 'Repeat menu handle
  modMenu.SetMenuRadio hSubMenu, 0
  modMenu.SetMenuRadio hSubMenu, 1
  modMenu.SetMenuRadio hSubMenu, 2
  hSubMenu = modMenu.GetSubMenuHandle(hMenu, 6) 'shuffle menu handle
  modMenu.SetMenuRadio hSubMenu, 0
  modMenu.SetMenuRadio hSubMenu, 1
  modMenu.SetMenuRadio hSubMenu, 2
  
  hMenu = modMenu.GetMenuHandle(Me.hwnd)
  hMenu = modMenu.GetSubMenuHandle(hMenu, 0)
  hSubMenu = modMenu.GetSubMenuHandle(hMenu, 12)
  modMenu.SetMenuRadio hSubMenu, 0
  modMenu.SetMenuRadio hSubMenu, 1
  modMenu.SetMenuRadio hSubMenu, 2
  
End Sub

Private Sub menAbout_Click()
  'Shows about window
  If Not frmMain.Visible Then frmMain.cTray_LButtonDblClk 'restore from tray
  frmAbout.Show , frmMain
End Sub

Public Sub menAddWin_Click()
  frmAdd.Show
End Sub

Public Sub menAdvance_Click()
  Settings.Advance = Not Settings.Advance
  menAdvance.Checked = Settings.Advance
End Sub

Private Sub menAudioCD_Click()
  AddCDAudio
End Sub

Public Sub menBck_Click()
  frmMain.Position.Value = frmMain.Position.Value - 5000
  frmMain.Position_Change
End Sub

Private Sub menCrop_Click()
  'Removes files NOT selected from playlist
  On Error Resume Next
  Dim X As Long
  Dim Count As Long
  
  'loop through all items in the list last to first
  For X = UBound(Playlist) To 1 Step -1
    'If the current item ISN'T selected
    If Not Playlist(X).Selected Then
      If Playing = X Then Playing = 0
      'Set to be removed
      Playlist(X).Removed = True
      'Add one to cound to keep track of how many was removed
      Count = Count + 1
    End If
  Next X

  'Cleans up the playlist from unused entries
  CleanUpPlaylist Count

  'refresh playlist
  UpdateList
  
End Sub

Public Sub menDeleteAll_Click()
  'Removes all files
  
  On Error Resume Next
  
  'Clears array & list
  ReDim Playlist(0) As PlaylistData
  ShuffleNum = 0
  Playing = 0
  UpdateList
  
End Sub

Public Sub menDeleteFile_Click()
  'Removes Selected files
  On Error Resume Next
  Dim X As Long
  Dim Count As Long
   
  'loop through all files in list
  For X = UBound(Playlist) To 1 Step -1  'reverse
    If Playlist(X).Selected Then  'if selected
      If Playing = X Then Playing = 0
      'Update array
      Playlist(X).Removed = True
      'Remove from list
      Count = Count + 1
    End If
  Next X

  'Cleans up the playlist from unused entries
  CleanUpPlaylist Count

  'updates form
  UpdateList
  
End Sub

Private Sub menExit_Click()
  Unload frmMain
End Sub

Private Sub menFileInfo_Click()
  frmView.View Playlist(frmPlaylist.List.ItemSelected).Reference
End Sub

Public Sub menFwd_Click()
  On Error Resume Next
  frmMain.Position.Value = frmMain.Position.Value + 5000
  frmMain.Position_Change
End Sub

Public Sub menHelp_Click()
  If Not frmMain.Visible Then frmMain.cTray_LButtonDblClk 'restore from tray
  frmHelp.Show , frmMain
  frmHelp.web.Navigate App.Path & "\docs\index.html"
End Sub

Private Sub menLibrary_Click()
  frmLibrary.Show
End Sub

Private Sub menListSave_Click()
  'Saves playlist
 
  On Error GoTo ErrHandler
  Dim cdg As New clsCommonDialog
 
  If UBound(Playlist) > 0 Then
    'Shows save dialog box
    With cdg
      Set .Parent = frmPlaylist
      .CancelError = True
      .Filter = "Simple Amp Playlist (*.playlist)|*.playlist"
      .DialogTitle = "Save Playlist"
      .Flags = cdlOFNCreatePrompt Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
      .ShowSave
      
      .FileName = sAppend(LCase(.FileName), ".playlist")
      If Not SavePlaylist(.FileName) Then
        MsgBox "Saving to " & Chr(34) & .FileName & Chr(34) & " was unsuccessful. Maybe it's my fault, or yours... It doesn't matter as long as we all just forget about this!", vbCritical, "Saving Error"
      End If
    
    End With
  Else
    MsgBox "The Playlist is empty. Fill it up with some music first!", vbInformation
  End If

  Exit Sub
ErrHandler:
  If Err.Number = 32755 Then Exit Sub 'dialog cancel
  If cLog.ErrorMsg(Err, "frmMenus, menListSave") = vbYes Then Resume Next Else frmMain.UnloadAll
End Sub

Private Sub menMBrowse_Click()
  frmAdd.Show
End Sub

Private Sub menMMediaLib_Click()
  frmLibrary.Show
End Sub

Private Sub menNext_Click()
  frmMain.PlayNext
End Sub

Private Sub menNextSkinPreset_Click()
  lCurVisPreset = lCurVisPreset + 1
  If lCurVisPreset > UBound(SkinPresets) Then lCurVisPreset = 0
  LoadSkinPreset lCurVisPreset
End Sub

Private Sub menPlay_Click()
  On Error Resume Next
  
  Playing = frmPlaylist.List.ItemSelected
  frmMain.Play
  
End Sub

Private Sub menPlaylist_Click()
  On Error Resume Next
  frmMain.btnPlaylist.Value = Not frmMain.btnPlaylist.Value
  If frmMain.Visible Then
    frmMain.HideShowPlaylist
  Else
    Settings.PlaylistOn = frmMain.btnPlaylist.Value
    menPlaylist.Checked = Settings.PlaylistOn
  End If
End Sub

Private Sub menPlayPause_Click()
  frmMain.PlayPause
End Sub

Public Sub menPlistSearch_Click()
  frmSearch.Show , frmPlaylist
End Sub

Private Sub menPresetEdit_Click()
  frmPresetEdit.Show , frmMain
End Sub

Private Sub menPresetLoad_Click()
  frmPresetLoad.Show , frmMain
End Sub

Private Sub menPrev_Click()
  frmMain.PlayPrev
End Sub

Private Sub menRepeatVal_Click(index As Integer)
  Settings.Repeat = index
  frmMain.RepeatOnOff
End Sub

Private Sub menReset_Click()
  frmPlaylist.List.ColumnSort Sort_OriginalOrder
End Sub

Private Sub menReverse_Click()
  On Error Resume Next
  frmPlaylist.List.ReverseList
  frmPlaylist.List.Refresh
End Sub

Public Sub menSelPlaying_Click()
  If Playing > 0 Then
    frmPlaylist.List.SelectNone
    Playlist(Playing).Selected = True
    frmPlaylist.List.MakeVisible Playlist(Playing).index
    frmPlaylist.List.lKeySelected = Playing
    frmPlaylist.List.Refresh
  End If
End Sub

Private Sub menShuffleVal_Click(index As Integer)
  Settings.Shuffle = index
  frmMain.ShuffleOnOff
End Sub

Private Sub menShufReset_Click()
  Dim X As Long
  For X = 1 To UBound(Playlist)
    Playlist(X).lShuffleIndex = 0
  Next
  ShuffleNum = 0
End Sub

Private Sub menSoundStudio_Click()
  If Not frmMain.Visible Then frmMain.cTray_LButtonDblClk
  frmStudio.Show , frmMain
End Sub

Private Sub menTimesPlayed_Click()
  frmPlaylist.List.ColumnSort Sort_PlayTimes
End Sub

Private Sub menTimesSkipped_Click()
  frmPlaylist.List.ColumnSort Sort_SkipTimes
End Sub

Private Sub menTurnOff_Click()
  lCurVisPreset = 0
  Spectrum = 0
  frmMain.UpdateSpectrum
End Sub

Private Sub menVisEdit_Click()
  If Not frmMain.Visible Then frmMain.cTray_LButtonDblClk 'restore from tray
  frmPresetEdit.Show , frmMain
End Sub

Private Sub menVisLoad_Click()
  If Not frmMain.Visible Then frmMain.cTray_LButtonDblClk 'restore from tray
  frmPresetLoad.Show , frmMain
End Sub

Private Sub menVolDown_Click()
  frmMain.volume.Value = frmMain.volume.Value - 10
  frmMain.Volume_Change
End Sub

Public Sub menSelectAll_Click()
  'Selects all items in list
  On Error Resume Next
  Dim X As Long
  
  For X = 1 To UBound(Playlist)
    Playlist(X).Selected = True
  Next X
  frmPlaylist.List.Refresh
  
End Sub

Public Sub menSelectInvert_Click()
  'Inverts selection in list
  On Error Resume Next
  Dim X As Long
  
  For X = 1 To UBound(Playlist)
    Playlist(X).Selected = Not Playlist(X).Selected
  Next X
  frmPlaylist.List.Refresh
  
End Sub

Public Sub menSelectNone_Click()
  'Deselect all items in list
  On Error Resume Next
  Dim X As Long
  
  For X = 1 To UBound(Playlist)
    Playlist(X).Selected = False
  Next X
  frmPlaylist.List.Refresh

End Sub

Private Sub menSettings_Click()
  If Not frmMain.Visible Then frmMain.cTray_LButtonDblClk
  frmSettings.Show , frmMain
End Sub

Private Sub menSkin_Click()
  If Not frmMain.Visible Then frmMain.cTray_LButtonDblClk
  frmSkin.Show , frmMain
End Sub

Private Sub menSortAlbum_Click()
  frmPlaylist.List.ColumnSort Sort_Album
End Sub

Private Sub menSortArtistTitle_Click()
  frmPlaylist.List.ColumnSort Sort_AristTitle
End Sub

Private Sub menSortDate_Click()
  frmPlaylist.List.ColumnSort Sort_PlayDate
End Sub

Private Sub menSortFilename_Click()
  frmPlaylist.List.ColumnSort Sort_FileName
End Sub

Private Sub menSortGenre_Click()
  frmPlaylist.List.ColumnSort Sort_Genre
End Sub

Private Sub menSortTime_Click()
  frmPlaylist.List.ColumnSort Sort_Time
End Sub

Private Sub menSortType_Click()
  frmPlaylist.List.ColumnSort Sort_FileType
End Sub

Private Sub menStop_Click()
  frmMain.PlayStop
End Sub

Public Sub CleanUpPlaylist(Optional ByVal num2 As Long = 0)
  'This sub removes unused entries in the array Playlist()
  On Error GoTo ErrHandler
  
  Dim X As Long, Count As Long, x2 As Long, z As Long, f As Long

  'First, move existing items over thoses that doesn't, so in the end
  'we get an array with all the still existing items in the beginning
  'and removed items at the end.
  X = 1
  Do While X <= UBound(Playlist)
    If Playlist(X).Removed Then
      
      x2 = X + 1
      Do While x2 <= UBound(Playlist)
        If Playlist(x2).Removed = False Then
          Exit Do
        End If
        x2 = x2 + 1
      Loop
      
      If x2 > UBound(Playlist) Then
        Exit Do
      Else
        Playlist(X) = Playlist(x2)
        Playlist(x2).Removed = True
        Playlist(X).Removed = False
        'Playlist(X).IsBold = Playlist(x2).IsBold
        'Playlist(X).Reference = Playlist(x2).Reference
        'Playlist(X).Selected = Playlist(x2).Selected
        'Playlist(X).Index = Playlist(x2).Index
        'Playlist(X).lShuffleIndex = Playlist(x2).lShuffleIndex
        Count = Count + 1
      End If
    Else
      Count = Count + 1
    End If
    X = X + 1
  Loop
  ReDim Preserve Playlist(Count)
  
  'But now we have to setup the moved items new indexes.
  Dim minindex As Long
  Dim maxindex As Long
  
  'first get max index & min index
  For X = 1 To UBound(Playlist)
    If Playlist(X).index > maxindex Then
      maxindex = Playlist(X).index
    End If
  Next
  minindex = maxindex
  For X = 1 To UBound(Playlist)
    If Playlist(X).index < minindex Then
      minindex = Playlist(X).index
    End If
  Next
  
  'then go though each of the indexes and change them so they are after
  'eachother.
  x2 = 1
  For X = minindex To maxindex
  
    For f = 1 To UBound(Playlist)
      If Playlist(f).index = X Then
        z = f
        Exit For
      End If
    Next
    
    If z > 0 Then
      Playlist(z).index = x2
      x2 = x2 + 1
    End If
    z = 0
    
  Next
  
  'Now do the same as above, but with lShuffleIndex
  'first get max lShuffleIndex & min lShuffleIndex
  For X = 1 To UBound(Playlist)
    If Playlist(X).lShuffleIndex > maxindex Then
      maxindex = Playlist(X).lShuffleIndex
    End If
  Next
  minindex = maxindex
  For X = 1 To UBound(Playlist)
    If Playlist(X).lShuffleIndex < minindex Then
      minindex = Playlist(X).lShuffleIndex
    End If
  Next
  
  'then go though each of the lShuffleIndexes and change them so they are after
  'eachother.
  x2 = 1
  For X = minindex To maxindex
    For f = 1 To UBound(Playlist)
      If Playlist(f).lShuffleIndex = X Then
        z = f
        Exit For
      End If
    Next
    
    If z > 0 Then
      If ShuffleNum = Playlist(z).lShuffleIndex Then ShuffleNum = x2
      Playlist(z).lShuffleIndex = x2
      x2 = x2 + 1
    End If
    z = 0
  Next
  
  If UBound(Playlist) = 0 Then ShuffleNum = 0

  Exit Sub
ErrHandler:
  If cLog.ErrorMsg(Err, "frmMenus, CleanUpPlaylist") = vbYes Then Resume Next Else frmMain.UnloadAll
End Sub

Public Sub menListLoad_Click()
  'Loads playlist
  
  On Error GoTo ErrHandler
  Dim cdg As New clsCommonDialog
  
  'Shows open dialog box
  With cdg
    Set .Parent = frmPlaylist
    .CancelError = True
    .DialogTitle = "Load Playlist"
    .Filter = "All Supported Playlist Types|*.playlist;*.m3u;*.pls|Simple Amp Playlist (*.playlist)|*.playlist|M3U Playlist (*.m3u)|*.m3u|PLS Playlist (*.pls)|*.pls"
    .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNPathMustExist
    .ShowOpen
    
    Select Case LCase(Right(.FileName, 3))
      Case "ist" '.playlist
        If Not LoadPlaylist(.FileName) Then
          MsgBox "Loading " & Chr(34) & .FileName & Chr(34) & " was unsuccessful. The file might be corrupt, or worse!", vbCritical, "Load Error"
          Exit Sub
        End If
      Case "m3u"
        If Not LoadM3u(.FileName) Then
          MsgBox "Loading " & Chr(34) & .FileName & Chr(34) & " was unsuccessful. The file might be corrupt, or worse!", vbCritical, "Load Error"
          Exit Sub
        End If
      Case "pls"
        If Not LoadPls(.FileName) Then
          MsgBox "Loading " & Chr(34) & .FileName & Chr(34) & " was unsuccessful. The file might be corrupt, or worse!", vbCritical, "Load Error"
          Exit Sub
        End If
    End Select
    
  End With
  
  Exit Sub
ErrHandler:
  If Err.Number = 32755 Then Exit Sub 'Open dialog cancel
  If cLog.ErrorMsg(Err, "frmMenus, menListLoad") = vbYes Then Resume Next Else frmMain.UnloadAll
End Sub

Private Sub menVolUp_Click()
  frmMain.volume.Value = frmMain.volume.Value + 10
  frmMain.Volume_Change
End Sub

Private Sub menAllSupp_Click()
  On Error Resume Next
  Dim X As Byte
  
  For X = 0 To 11
    Settings.Filter(X) = True
    menFilt(X).Checked = True
  Next X
  frmAdd.UpdateFolder frmAdd.ftwMain.SelectedFolder
End Sub

Private Sub menAutosizeColumns_Click()
  On Error Resume Next
  Settings.AddAutosize = Not Settings.AddAutosize
  menAutosizeColumns.Checked = Settings.AddAutosize
  frmAdd.CheckExtended
End Sub

Private Sub menFinfo_Click()
  frmAdd.ShowInfo
End Sub

Private Sub menFilt_Click(index As Integer)
  'toggle checkmark on selected filter
  On Error Resume Next
  Settings.Filter(index) = Not Settings.Filter(index)
  menFilt(index).Checked = Settings.Filter(index)
  frmAdd.UpdateFolder frmAdd.ftwMain.SelectedFolder
End Sub

Private Sub menClrPlist_Click()
  menDeleteAll_Click
End Sub

Private Sub menView_Click(index As Integer)
  Settings.AddView = index
  menView(0).Checked = (index = 0)
  menView(1).Checked = (index = 1)
  menView(2).Checked = (index = 2)
  frmAdd.CheckExtended
End Sub

Private Sub menRemoveFile_Click()
  frmAdd.RemoveFile
End Sub

Private Sub menPlayfoca_Click()
  On Error Resume Next
  frmAdd.PlaySelected
End Sub

Private Sub menSelArtist_Click()
  On Error Resume Next
  frmAdd.SelectArtist
End Sub

Private Sub menSelGenre_Click()
  On Error Resume Next
  frmAdd.SelectGenre
End Sub

Private Sub menSelAlbum_Click()
  On Error Resume Next
  frmAdd.SelectAlbum
End Sub

Private Sub menTypeSeq_Click()
  On Error Resume Next
  Dim X As Byte
  
  For X = 5 To 10
    Settings.Filter(X) = Not Settings.Filter(X)
    menFilt(X).Checked = Not menFilt(X).Checked
  Next X
  frmAdd.UpdateFolder frmAdd.ftwMain.SelectedFolder
End Sub

Private Sub menTypeStream_Click()
  On Error Resume Next
  Dim X As Byte
  
  For X = 0 To 4
    Settings.Filter(X) = Not Settings.Filter(X)
    menFilt(X).Checked = Not menFilt(X).Checked
  Next X
  frmAdd.UpdateFolder frmAdd.ftwMain.SelectedFolder
End Sub

Private Sub menTypeAllNoPlay_Click()
  On Error Resume Next
  Dim X As Byte
  
  For X = 0 To 10
    Settings.Filter(X) = True
    menFilt(X).Checked = True
  Next X
  Settings.Filter(11) = False
  menFilt(11).Checked = False
  frmAdd.UpdateFolder frmAdd.ftwMain.SelectedFolder
End Sub

Private Sub menSelType_Click()
  On Error Resume Next
  frmAdd.SelectType
End Sub

Public Sub menSelNone_Click()
  On Error Resume Next
  Dim X As Long
  
  For X = 1 To frmAdd.lvwMain.ListItems.Count
    frmAdd.lvwMain.ListItems(X).Selected = False
  Next X
End Sub

Private Sub menSelAll_Click()
  On Error Resume Next
  Dim X As Long

  For X = 1 To frmAdd.lvwMain.ListItems.Count
    frmAdd.lvwMain.ListItems(X).Selected = True
  Next X
End Sub
