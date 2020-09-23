VERSION 5.00
Begin VB.Form frmPlaylist 
   BorderStyle     =   0  'None
   Caption         =   "Playlist"
   ClientHeight    =   2070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6435
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmPlaylist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   138
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   429
   ShowInTaskbar   =   0   'False
   Begin SimpleAmp.CtrlLister List 
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      SelColor        =   16777215
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SimpleAmp.CtrlButton btnSize 
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   240
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin SimpleAmp.CtrlScroller Scroll 
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Vertical        =   -1  'True
      Max             =   0
   End
   Begin SimpleAmp.CtrlButton btnClose 
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   240
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin SimpleAmp.CtrlButton btnAdd 
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Top             =   240
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin SimpleAmp.CtrlButton btnSelect 
      Height          =   495
      Left            =   5400
      TabIndex        =   6
      Top             =   240
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin SimpleAmp.CtrlButton btnRem 
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin SimpleAmp.CtrlButton btnList 
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.Image imgColumns 
      Height          =   255
      Left            =   120
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Top             =   360
      Width           =   300
   End
   Begin VB.Label lblTotalTime 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblTotalNum 
      BackStyle       =   0  'Transparent
      Caption         =   "0 files."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
   Begin VB.Image imgDropdown 
      Height          =   255
      Left            =   120
      OLEDropMode     =   1  'Manual
      Top             =   0
      Width           =   255
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmPlaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MovePlaylist As Boolean 'True if in move mode
Private MovePlaylistOldX As Long  'Old position X
Private MovePlaylistOldY As Long  'Old position Y

Private Sub btnAdd_KeyDown(KeyCode As Integer, Shift As Integer)
  List_KeyDown KeyCode, Shift
End Sub

Private Sub btnAdd_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub btnAdd_Pressed(ByVal Button As Integer)
  On Error Resume Next
  If Button = 1 And Settings.ButtonDefault Then
    frmMenus.menAddWin_Click
  Else
    ShowPopup frmPlaylist, frmMenus.menAdd, btnAdd, TextWidth(frmMenus.menLibrary.Caption) + 45
  End If
End Sub

Private Sub btnClose_KeyDown(KeyCode As Integer, Shift As Integer)
  List_KeyDown KeyCode, Shift
End Sub

Private Sub btnClose_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub btnClose_Pressed(ByVal Button As Integer)
  If Button = 1 Then
    frmMain.btnPlaylist.Value = False
    frmMain.HideShowPlaylist
  End If
End Sub

Private Sub btnList_KeyDown(KeyCode As Integer, Shift As Integer)
  List_KeyDown KeyCode, Shift
End Sub

Private Sub btnList_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub btnList_Pressed(ByVal Button As Integer)
  ShowPopup frmPlaylist, frmMenus.menList, btnList, TextWidth(frmMenus.menListLoad.Caption) + 45 + 50
End Sub

Private Sub btnRem_KeyDown(KeyCode As Integer, Shift As Integer)
  List_KeyDown KeyCode, Shift
End Sub

Private Sub btnRem_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub btnRem_Pressed(ByVal Button As Integer)
  On Error Resume Next
  If Button = 1 And Settings.ButtonDefault Then
    frmMenus.menDeleteFile_Click
  Else
    ShowPopup frmPlaylist, frmMenus.menDel, btnRem, TextWidth(frmMenus.menDeleteFile.Caption) + 45
  End If
End Sub

Private Sub btnSelect_KeyDown(KeyCode As Integer, Shift As Integer)
  List_KeyDown KeyCode, Shift
End Sub

Private Sub btnSelect_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub btnSelect_Pressed(ByVal Button As Integer)
  On Error Resume Next
  If Button = 1 And Settings.ButtonDefault Then
    frmMenus.menSelectAll_Click
  Else
    ShowPopup frmPlaylist, frmMenus.menSelect, btnSelect, TextWidth(frmMenus.menSelPlaying.Caption) + 45
  End If
End Sub

Private Sub btnSize_KeyDown(KeyCode As Integer, Shift As Integer)
  List_KeyDown KeyCode, Shift
End Sub

Private Sub btnSize_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub btnSize_Pressed(ByVal Button As Integer)
  If Button = 1 Then ToggleSize
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  List_KeyDown KeyCode, Shift
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseDown Button, Shift, X * Screen.TwipsPerPixelX, y * Screen.TwipsPerPixelY
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseMove Button, Shift, X * Screen.TwipsPerPixelX, y * Screen.TwipsPerPixelY
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseUp Button, Shift, X * Screen.TwipsPerPixelX, y * Screen.TwipsPerPixelY
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub imgColumns_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseDown Button, Shift, X, y
End Sub

Private Sub imgColumns_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseMove Button, Shift, X, y
End Sub

Private Sub imgColumns_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseUp Button, Shift, X, y
End Sub

Private Sub imgColumns_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub imgDropdown_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

  On Error Resume Next

  If Button = 2 Then
    PopupMenu frmMenus.menMain
  Else

    'Get windows work area size to type variable Scr
    RefreshScreenRECT
    MovePlaylist = True
    MovePlaylistOldX = X
    MovePlaylistOldY = y
  
  End If
  ctlSetFocus List
  
End Sub

Private Sub imgDropdown_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  'gah this sub is messy! i don't have the energy to go through it, so you do it
  On Error GoTo ErrHandler
  
  If MovePlaylist Then
  
    'If snap is on, dock window to screen edges & main window edges.
    If Settings.Snap Then
      If (Left + (X - MovePlaylistOldX)) \ Screen.TwipsPerPixelX >= Scr.Right Or (Top + (y - MovePlaylistOldY)) \ Screen.TwipsPerPixelY >= Scr.Bottom Then Exit Sub
      With Scr
        'Calculate Window position in pixels
        Dim DragLeft As Long, DragLeftWidth As Long, DragTop As Long, DragTopHeight As Long
        DragLeft = (Left + (X - MovePlaylistOldX)) \ Screen.TwipsPerPixelX
        DragLeftWidth = (Left + (X - MovePlaylistOldX) + Width) \ Screen.TwipsPerPixelX
        DragTopHeight = (Top + (y - MovePlaylistOldY) + Height) \ Screen.TwipsPerPixelY
        DragTop = (Top + (y - MovePlaylistOldY)) \ Screen.TwipsPerPixelY
        Dim NewLeft As Long, NewTop As Long
        
        'Snap to right or left edge of screen
        If DragLeftWidth > .Right - SnapWidth And DragLeftWidth < .Right + SnapWidth Then
          NewLeft = (.Right * Screen.TwipsPerPixelX) - Width
          Docked = False
        ElseIf DragLeft < .Left + SnapWidth And DragLeft > .Left - SnapWidth Then
          NewLeft = (.Left * Screen.TwipsPerPixelX)
          Docked = False
        'Snap to main window left
        ElseIf DragLeftWidth > (frmMain.Left \ Screen.TwipsPerPixelX) - SnapWidth And DragLeftWidth < (frmMain.Left \ Screen.TwipsPerPixelY) + SnapWidth And DragTop >= (frmMain.Top - Me.Height) \ Screen.TwipsPerPixelY And DragTop <= (frmMain.Top + frmMain.Height) \ Screen.TwipsPerPixelY Then
          NewLeft = frmMain.Left - Width
          If DragTopHeight > ((frmMain.Top + frmMain.Height) \ Screen.TwipsPerPixelY) - SnapWidth And DragTopHeight < ((frmMain.Top + frmMain.Height) \ Screen.TwipsPerPixelY) + SnapWidth Then
            NewTop = frmMain.Top + frmMain.Height - Me.Height
          ElseIf DragTop > (frmMain.Top \ Screen.TwipsPerPixelY) - SnapWidth And DragTop < (frmMain.Top \ Screen.TwipsPerPixelY) + SnapWidth Then
            NewTop = frmMain.Top
          End If
          Docked = True
        'Snap to main window right
        ElseIf DragLeft > ((frmMain.Left + frmMain.Width) \ Screen.TwipsPerPixelX) - SnapWidth And DragLeft < ((frmMain.Left + frmMain.Width) \ Screen.TwipsPerPixelX) + SnapWidth And DragTop >= (frmMain.Top - Me.Height) \ Screen.TwipsPerPixelY And DragTop <= (frmMain.Top + frmMain.Height) \ Screen.TwipsPerPixelY Then
          NewLeft = frmMain.Left + frmMain.Width
          If DragTopHeight > ((frmMain.Top + frmMain.Height) \ Screen.TwipsPerPixelY) - SnapWidth And DragTopHeight < ((frmMain.Top + frmMain.Height) \ Screen.TwipsPerPixelY) + SnapWidth Then
            NewTop = frmMain.Top + frmMain.Height - Me.Height
          ElseIf DragTop > (frmMain.Top \ Screen.TwipsPerPixelY) - SnapWidth And DragTop < (frmMain.Top \ Screen.TwipsPerPixelY) + SnapWidth Then
            NewTop = frmMain.Top
          End If
          Docked = True
        Else
          NewLeft = Left + (X - MovePlaylistOldX)
          Docked = False
        End If
        
        'Snap to lower or upper edge of screen
        If NewTop = 0 Then
          If DragTopHeight > .Bottom - SnapWidth And DragTopHeight < .Bottom + SnapWidth Then
            NewTop = (.Bottom * Screen.TwipsPerPixelY) - Height
            Docked = False
          ElseIf DragTop < .Top + SnapWidth And DragTop > .Top - SnapWidth Then
            NewTop = (.Top * Screen.TwipsPerPixelY)
            Docked = False
          'Snap to main window top
          ElseIf DragTopHeight > (frmMain.Top \ Screen.TwipsPerPixelY) - SnapWidth And DragTopHeight < (frmMain.Top \ Screen.TwipsPerPixelY) + SnapWidth And DragLeft >= (frmMain.Left - Me.Width) \ Screen.TwipsPerPixelX And DragLeft <= (frmMain.Left + frmMain.Width) \ Screen.TwipsPerPixelX Then
            NewTop = frmMain.Top - Height
            If DragLeftWidth > ((frmMain.Left + frmMain.Width) \ Screen.TwipsPerPixelX) - SnapWidth And DragLeftWidth < ((frmMain.Left + frmMain.Width) \ Screen.TwipsPerPixelX) + SnapWidth Then
              NewLeft = frmMain.Left + frmMain.Width - Me.Width
            ElseIf DragLeft > (frmMain.Left \ Screen.TwipsPerPixelX) - SnapWidth And DragLeft < (frmMain.Left \ Screen.TwipsPerPixelX) + SnapWidth Then
              NewLeft = frmMain.Left
            End If
            Docked = True
          'Snap to main window bottom
          ElseIf DragTop > ((frmMain.Top + frmMain.Height) \ Screen.TwipsPerPixelY) - SnapWidth And DragTop < ((frmMain.Top + frmMain.Height) \ Screen.TwipsPerPixelY) + SnapWidth And DragLeft >= (frmMain.Left - Me.Width) \ Screen.TwipsPerPixelX And DragLeft <= (frmMain.Left + frmMain.Width) \ Screen.TwipsPerPixelX Then
            NewTop = frmMain.Top + frmMain.Height
            If DragLeftWidth > ((frmMain.Left + frmMain.Width) \ Screen.TwipsPerPixelX) - SnapWidth And DragLeftWidth < ((frmMain.Left + frmMain.Width) \ Screen.TwipsPerPixelX) + SnapWidth Then
              NewLeft = frmMain.Left + frmMain.Width - Me.Width
            ElseIf DragLeft > (frmMain.Left \ Screen.TwipsPerPixelX) - SnapWidth And DragLeft < (frmMain.Left \ Screen.TwipsPerPixelX) + SnapWidth Then
              NewLeft = frmMain.Left
            End If
            Docked = True
          Else
            NewTop = Top + (y - MovePlaylistOldY)
          End If
        End If
        
        Left = NewLeft
        Top = NewTop
        If Docked Then
          DockedLeft = Me.Left - frmMain.Left
          DockedTop = Me.Top - frmMain.Top
        Else
          DockedLeft = 0
          DockedTop = 0
        End If
        
      End With
    Else
      Left = Left + (X - MovePlaylistOldX)
      Top = Top + (y - MovePlaylistOldY)
    End If
  End If
  
 
  Exit Sub
ErrHandler:
  If cLog.ErrorMsg(Err, "frmMenus, imgDropDown_MouseMove") = vbYes Then Resume Next Else frmMain.UnloadAll
End Sub

Private Sub imgDropdown_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  On Error Resume Next
  MovePlaylist = False
End Sub

Public Sub imgDropdown_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  'Drag & drop
  'This will add all files dropped on the playlist window to the playlist
  On Error GoTo ErrHandler
  
  Dim i As Long
  Dim AddNew As Boolean
  Dim File As New clsFile
  
  If UBound(Playlist) = 0 Then AddNew = True

  If Data.GetFormat(vbCFFiles) Then 'true if data is list of files
    If Shift = vbShiftMask Then 'clear list and add if shift, frmmain does this by default
      frmMain.imgDropdown_OLEDragDrop Data, Effect, Button, 0, X, y
      Exit Sub
    End If
    
    For i = 1 To Data.Files.Count 'go through each file in dropped list
      File.sFilename = Data.Files(i)
      If File.eFileAttributes And eDIRECTORY Then 'add dir and subdirs
        SimpleAddDir Data.Files(i)
      Else
        Select Case LCase(File.sExtension)
          Case "playlist", "m3u", "pls"
            HandlePlaylist Data.Files(i)
          Case "mp3", "mp2", "asf", "wma", "ogg", "mid", "midi", "rmi", "mod", "xm", "it", "s3m", "wav", "sgm"
            If CreateLibrary(Data.Files(i)) = 0 Then
              MsgBox """" & Data.Files(i) & """ could not be opened.", vbExclamation, "Error in file"
            End If
          Case Else 'not recognised extention
            If Sound.TestFile(Data.Files(i)) = 0 Then 'test if file can be opened
              MsgBox "The file " & Chr(34) & Data.Files(i) & Chr(34) & " is not supported.", vbExclamation
            Else
              CreateLibrary Data.Files(i)
            End If
        End Select
      End If
    Next i
    
    UpdateList

    If UBound(Playlist) > 0 And AddNew Then frmMain.PlayNext
  Else
    MsgBox "Whatever you dragged to Simple Amp, it isn't supported." & vbCrLf & "What you can drag here:" & vbCrLf & vbCrLf & "- One or more supported filetypes." & vbCrLf & "- One ore more folders (all files in them and their subfolders will be added)." & vbCrLf & "Dragging files to the player window will clear the playlist before adding them. Dragging files to the playlist window will add them to the list. Hold shift while dropping to do the opposite, i.e. in the player window to not clear the list.", vbInformation, "Simple Amp"
  End If
    
 
  Exit Sub
ErrHandler:
  If cLog.ErrorMsg(Err, "frmMenus, imgDropdown_OLEDragDrop") = vbYes Then Resume Next Else frmMain.UnloadAll
End Sub

Private Sub lblTotalNum_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseDown Button, Shift, X, y
End Sub

Private Sub lblTotalNum_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseMove Button, Shift, X, y
End Sub

Private Sub lblTotalNum_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseUp Button, Shift, X, y
End Sub

Private Sub lblTotalNum_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub lblTotalTime_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseDown Button, Shift, X, y
End Sub

Private Sub lblTotalTime_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseMove Button, Shift, X, y
End Sub

Private Sub lblTotalTime_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseUp Button, Shift, X, y
End Sub

Private Sub lblTotalTime_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub List_ItemClick(Item As Long, num As Long)
  If frmView.Visible Then frmView.View Playlist(num).Reference
End Sub

Private Sub List_ItemDblClick(Item As Long, num As Long)
  On Error Resume Next
  
  If num > 0 And num <= UBound(Playlist) Then
    Playing = num
    If Playlist(Playing).lShuffleIndex = 0 And Settings.Shuffle > 0 Then ShuffleNum = ShuffleNum + 1
    StreamStart = True
    frmMain.Play
  End If
  
End Sub

Private Sub List_KeyDown(KeyCode As Integer, Shift As Integer)
  'This sub catches the keypresses of the playlist window
  On Error Resume Next

  Select Case KeyCode
    Case vbKeyDelete
      frmMenus.menDeleteFile_Click
    Case vbKeyReturn
      Playing = List.ItemSelected
      If Playlist(Playing).lShuffleIndex = 0 And Settings.Shuffle > 0 Then ShuffleNum = ShuffleNum + 1
      frmMain.Play
    Case vbKeyF1
      frmMenus.menHelp_Click
    Case vbKeyF2
      frmAdd.Show
    Case vbKeyF3
      frmLibrary.Show
    Case vbKeyA
      If Shift = vbCtrlMask Then
        frmMenus.menSelectAll_Click
      Else
        frmMenus.menAdvance_Click
      End If
    Case vbKeyI
      If Shift = vbCtrlMask Then frmMenus.menSelectInvert_Click
    Case vbKeyN
      If Shift = vbCtrlMask Then frmMenus.menSelectNone_Click
    Case vbKeyH
      frmMenus.menSelPlaying_Click
    Case vbKeyZ
      frmMain.PlayPrev
    Case vbKeyX
      frmMain.PlayPause
    Case vbKeyC
      frmMain.PlayStop
    Case vbKeyV
      frmMain.PlayNext
    Case vbKeyP
      frmMain.btnPlaylist.Value = Not frmMain.btnPlaylist.Value
      frmMain.HideShowPlaylist
    Case vbKeyR
      frmMain.btnRepeat.Value = Not frmMain.btnRepeat.Value
      frmMain.RepeatOnOff
    Case vbKeyS
      frmMain.btnShuffle.Value = Not frmMain.btnShuffle.Value
      frmMain.ShuffleOnOff
    Case vbKeyLeft
      If Sound.StreamIsLoaded Then
        frmMain.Position.Value = frmMain.Position.Value - 5000
      ElseIf Sound.MusicIsLoaded Then
        frmMain.Position.Value = frmMain.Position.Value - Sound.MusicGetRows(Sound.MusicOrder)
      End If
      frmMain.Position_Change
    Case vbKeyRight
      If Sound.StreamIsLoaded Then
        frmMain.Position.Value = frmMain.Position.Value + 5000
      ElseIf Sound.MusicIsLoaded Then
        frmMain.Position.Value = frmMain.Position.Value + Sound.MusicGetRows(Sound.MusicOrder)
      End If
      frmMain.Position_Change
    Case vbKeyF
      If Shift = vbCtrlMask Then frmMenus.menPlistSearch_Click
    Case vbKey1
      List.ColumnSort Sort_AristTitle
    Case vbKey2
      List.ColumnSort Sort_Album
    Case vbKey3
      List.ColumnSort Sort_Genre
    Case vbKey4
      List.ColumnSort Sort_Time
    Case vbKey5
      List.ColumnSort Sort_FileName
    Case vbKey6
      List.ColumnSort Sort_FileType
    Case vbKey7
      List.ColumnSort Sort_PlayDate
    Case vbKey8
      List.ColumnSort Sort_PlayTimes
    Case vbKey9
      List.ColumnSort Sort_SkipTimes
    Case vbKey0
      List.ColumnSort Sort_OriginalOrder
  End Select
  
End Sub

Private Sub List_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  If Button = 2 Then PopupMenu frmMenus.menRClick
End Sub

Private Sub List_Scroll()
  On Error Resume Next
  Scroll.Value = List.Value
  Scroll.Max = List.Max
End Sub

Private Sub Scroll_Change()
  On Error Resume Next
  List.Value = Scroll.Value
End Sub

Public Sub ToggleSize()
  On Error Resume Next
  Settings.PlaylistSmall = Not Settings.PlaylistSmall
  frmMain.UpdatePlaylist
End Sub

Private Sub Scroll_KeyDown(KeyCode As Integer, Shift As Integer)
  List_KeyDown KeyCode, Shift
End Sub

Private Sub Scroll_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub
