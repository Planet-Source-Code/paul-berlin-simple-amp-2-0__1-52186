VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Simple Amp 1.2"
   ClientHeight    =   3555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   Icon            =   "frmMain.frx":0000
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   237
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   250
   Begin VB.PictureBox p 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   20
      Top             =   120
      Width           =   255
      Visible         =   0   'False
   End
   Begin VB.Timer tmrCDAudio 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3120
      Top             =   1320
   End
   Begin VB.Timer tmrStream 
      Interval        =   1000
      Left            =   2640
      Top             =   840
   End
   Begin VB.Timer tmrScope 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   2640
      Top             =   1320
   End
   Begin SimpleAmp.CtrlScroller Volume 
      Height          =   495
      Left            =   1560
      TabIndex        =   19
      Top             =   1200
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Max             =   255
   End
   Begin SimpleAmp.CtrlScroller Position 
      Height          =   495
      Left            =   840
      TabIndex        =   18
      Top             =   1200
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Min             =   1
      Max             =   1
      Value           =   1
   End
   Begin SimpleAmp.CtrlButton btnMinimize 
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.PictureBox Keys 
      Height          =   255
      Left            =   -750
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   -750
      Width           =   255
   End
   Begin VB.Timer tmrMain 
      Interval        =   100
      Left            =   3120
      Top             =   840
   End
   Begin SimpleAmp.CtrlButton btnClose 
      Height          =   495
      Left            =   840
      TabIndex        =   10
      Top             =   1920
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin SimpleAmp.CtrlButton btnRepeat 
      Height          =   495
      Left            =   1560
      TabIndex        =   11
      Top             =   1920
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Extended        =   -1  'True
   End
   Begin SimpleAmp.CtrlButton btnShuffle 
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   1920
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Extended        =   -1  'True
   End
   Begin SimpleAmp.CtrlButton btnPrev 
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin SimpleAmp.CtrlButton btnNext 
      Height          =   495
      Left            =   840
      TabIndex        =   14
      Top             =   2640
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin SimpleAmp.CtrlButton btnPlayPause 
      Height          =   495
      Left            =   1560
      TabIndex        =   15
      Top             =   2640
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin SimpleAmp.CtrlButton btnStop 
      Height          =   495
      Left            =   2280
      TabIndex        =   16
      Top             =   2640
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin SimpleAmp.CtrlButton btnPlaylist 
      Height          =   495
      Left            =   3000
      TabIndex        =   17
      Top             =   1920
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Extended        =   -1  'True
   End
   Begin VB.PictureBox picVis 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   120
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Width           =   300
   End
   Begin SimpleAmp.ctrlLabel lblArtistTitle 
      Height          =   255
      Left            =   1320
      TabIndex        =   21
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblComments 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   6
      Top             =   840
      UseMnemonic     =   0   'False
      Width           =   855
   End
   Begin VB.Label lblGenre 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1080
      OLEDropMode     =   1  'Manual
      TabIndex        =   8
      Top             =   840
      UseMnemonic     =   0   'False
      Width           =   495
   End
   Begin VB.Label lblYear 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1080
      OLEDropMode     =   1  'Manual
      TabIndex        =   5
      Top             =   405
      UseMnemonic     =   0   'False
      Width           =   375
   End
   Begin VB.Label lblAlbum 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   495
   End
   Begin VB.Label lblTotalTime 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2640
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   480
      UseMnemonic     =   0   'False
      Width           =   675
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1920
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   480
      UseMnemonic     =   0   'False
      Width           =   675
   End
   Begin VB.Image imgDropdown 
      Height          =   255
      Left            =   120
      OLEDropMode     =   1  'Manual
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgStereoMono 
      Height          =   255
      Left            =   480
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1560
      OLEDropMode     =   1  'Manual
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
'                      Simple Amp 2.0
'                  By Paul Berlin 2002-2003
'                  berlin_paul@hotmail.com
'                    http://pab.mydns.org
'----------------------------------------------------------
'
'For more info look in docs\index.html
'
'Feel free to change/add/improve this code, but if you do
'it would be nice if you send me the updated code =).
'
'You are free to use this code in your own programs,
'as long as you give credit where credit is due.


Option Explicit

'Window region API
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Const RGN_DIFF = 4

'APIs for drawing the visualization
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function TransparentBlt Lib "msimg32" (ByVal hDCDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hDCSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long

Private WithEvents cTray As clsSysTray
Attribute cTray.VB_VarHelpID = -1
Private MoveMain As Boolean 'True if in move mode
Private MoveMainOldX As Long 'Old Move X
Private MoveMainOldY As Long 'Old Move Y
Private CDTime As Long

Private Sub btnClose_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_KeyDown KeyCode, Shift
End Sub

Private Sub btnClose_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub btnClose_Pressed(ByVal Button As Integer)
  If Button = 1 Then UnloadAll
End Sub

Private Sub btnMinimize_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_KeyDown KeyCode, Shift
End Sub

Private Sub btnMinimize_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub btnMinimize_Pressed(ByVal Button As Integer)
  If Button = 1 Then Minimize
End Sub

Private Sub btnNext_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_KeyDown KeyCode, Shift
End Sub

Private Sub btnNext_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub btnNext_Pressed(ByVal Button As Integer)
  If Button = 1 Then PlayNext
End Sub

Private Sub btnPlaylist_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_KeyDown KeyCode, Shift
End Sub

Private Sub btnPlaylist_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub btnPlaylist_Pressed(ByVal Button As Integer)
  HideShowPlaylist
End Sub

Private Sub btnPlayPause_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_KeyDown KeyCode, Shift
End Sub

Private Sub btnPlayPause_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub btnPlayPause_Pressed(ByVal Button As Integer)
  If Button = 1 Then PlayPause
End Sub

Private Sub btnPrev_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_KeyDown KeyCode, Shift
End Sub

Private Sub btnPrev_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub btnPrev_Pressed(ByVal Button As Integer)
  If Button = 1 Then PlayPrev
End Sub

Private Sub btnRepeat_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_KeyDown KeyCode, Shift
End Sub

Private Sub btnRepeat_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub btnRepeat_Pressed(ByVal Button As Integer)
  On Error Resume Next
  If Button = 2 Then
    btnRepeat.Value = Not btnRepeat.Value 'reset change & show popup instead
    ShowPopup frmMain, frmMenus.menRepeat, btnRepeat, TextWidth(frmMenus.menRepeatVal(2).Caption) + 45
  Else
    If btnRepeat.Value Then
      Settings.Repeat = 1
    Else
      Settings.Repeat = 0
    End If
    RepeatOnOff
  End If
End Sub

Private Sub btnShuffle_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_KeyDown KeyCode, Shift
End Sub

Private Sub btnShuffle_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub btnShuffle_Pressed(ByVal Button As Integer)
  On Error Resume Next
  If Button = 2 Then
    btnShuffle.Value = Not btnShuffle.Value 'reset change & show popup instead
    ShowPopup frmMain, frmMenus.menShuffle, btnShuffle, TextWidth(frmMenus.menShuffleVal(2).Caption) + 45
  Else
    If btnShuffle.Value Then
      Settings.Shuffle = 1
    Else
      Settings.Shuffle = 0
    End If
    ShuffleOnOff
  End If
End Sub

Private Sub btnStop_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_KeyDown KeyCode, Shift
End Sub

Private Sub btnStop_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub btnStop_Pressed(ByVal Button As Integer)
  If Button = 1 Then PlayStop
End Sub

Public Sub cTray_LButtonDblClk()
  On Error Resume Next
  'restore windows
  If Not Settings.AlwaysTray Then cTray.RemoveFromSysTray
  Me.Show
  AlwaysOnTop Me, Settings.OnTop
  If Settings.PlaylistOn Then frmPlaylist.Show
  AlwaysOnTop frmPlaylist, Settings.OnTop
End Sub

Private Sub cTray_RButtonUp()
  On Error Resume Next
  'when you rightclick the on the systray icon, show popup menu
  PopupMenu frmMenus.menMain
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_KeyDown KeyCode, Shift
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

Private Sub imgDropdown_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  On Error Resume Next
  
  If Button = 2 Then
    PopupMenu frmMenus.menMain
  Else
    'Get windows work area size to type variable Scr
    RefreshScreenRECT
    MoveMain = True
    MoveMainOldX = X
    MoveMainOldY = y
  End If
  ctlSetFocus Keys
  
End Sub

Private Sub imgDropdown_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  On Error Resume Next
   
  If MoveMain Then
    If Settings.Snap Then  'If Snap window option is on, snap it to screen edges
      If (Left + (X - MoveMainOldX)) / Screen.TwipsPerPixelX >= Scr.Right Or (Top + (y - MoveMainOldY)) / Screen.TwipsPerPixelY >= Scr.Bottom Then Exit Sub
      With Scr
      
      '########### Commented out:
      '########### Unfinished code for snapping to screen edges with
      '########### playlistwindow included when docked.
      '########### could not get it right, so it's out
      '########### possibly best to just remove everything and rewrite from scratch
      
'        Dim RightOffset As Long, LeftOffset As Long
'        Dim BottomOffset As Long, TopOffset As Long
        Dim DragLeft As Long, DragLeftWidth As Long, DragTop As Long, DragTopHeight As Long
        DragLeft = (Left + (X - MoveMainOldX)) / Screen.TwipsPerPixelX
        DragLeftWidth = (Left + (X - MoveMainOldX) + Width) / Screen.TwipsPerPixelX
        DragTopHeight = (Top + (y - MoveMainOldY) + Height) / Screen.TwipsPerPixelY
        DragTop = (Top + (y - MoveMainOldY)) / Screen.TwipsPerPixelY
        
'
'        'Recalculate bound with playlist window included
'        If Docked And frmPlaylist.Visible Then
'          If DragLeft + ((frmPlaylist.Width + DockedLeft) / Screen.TwipsPerPixelX) > DragLeftWidth Then
'            DragLeftWidth = DragLeft + ((frmPlaylist.Width + DockedLeft) / Screen.TwipsPerPixelX)
'          End If
'          If DragLeft + (DockedLeft / Screen.TwipsPerPixelX) < DragLeft Then
'            DragLeft = DragLeft + (DockedLeft / Screen.TwipsPerPixelX)
'          End If
'          If DragTop + ((frmPlaylist.Height + DockedTop) / Screen.TwipsPerPixelY) > DragTopHeight Then
'            DragTopHeight = DragTop + ((frmPlaylist.Height + DockedTop) / Screen.TwipsPerPixelY)
'          End If
'          If DragTop + (DockedTop / Screen.TwipsPerPixelY) < DragTop Then
'            DragTop = DragTop + (DockedTop / Screen.TwipsPerPixelY)
'          End If
'
'          If DockedLeft > 0 Then
'            RightOffset = DragLeftWidth - DragLeft
'          Else
'            RightOffset = Me.ScaleWidth
'          End If
'          If DockedLeft < 0 Then
'            LeftOffset = (DockedLeft * -1) / Screen.TwipsPerPixelX
'          End If
'          If DockedTop > 0 Then
'            BottomOffset = DragTopHeight - DragTop
'          Else
'            BottomOffset = Me.ScaleHeight
'          End If
'          If DockedTop < 0 Then
'            TopOffset = (DockedTop * -1) / Screen.TwipsPerPixelX
'          End If
'        Else
'          RightOffset = Me.ScaleWidth
'          BottomOffset = Me.ScaleHeight
'        End If
'
'
'
'
'        'Snap to right or left edge of screen
'        If DragLeftWidth > .Right - SnapWidth And DragLeftWidth < .Right + SnapWidth Then
'          Left = (.Right - RightOffset) * Screen.TwipsPerPixelX
'        ElseIf DragLeft < .Left + SnapWidth And DragLeft > .Left - SnapWidth Then
'          Left = (.Left + LeftOffset) * Screen.TwipsPerPixelX
'        Else
'          Left = Left + (x - MoveMainOldX)
'        End If
'        'Snap to lower or upper edge of screen
'        If DragTopHeight > .Bottom - SnapWidth And DragTopHeight < .Bottom + SnapWidth Then
'          Top = (.Bottom - BottomOffset) * Screen.TwipsPerPixelY
'        ElseIf DragTop < .Top + SnapWidth And DragTop > .Top - SnapWidth Then
'          Top = (.Top + TopOffset) * Screen.TwipsPerPixelY
'        Else
'          Top = Top + (y - MoveMainOldY)
'        End If
        'Snap to right or left edge of screen
        If DragLeftWidth > .Right - SnapWidth And DragLeftWidth < .Right + SnapWidth Then
          Left = (.Right * Screen.TwipsPerPixelX) - Width
        ElseIf DragLeft < .Left + SnapWidth And DragLeft > .Left - SnapWidth Then
          Left = .Left * Screen.TwipsPerPixelX
        Else
          Left = Left + (X - MoveMainOldX)
        End If
        'Snap to lower or upper edge of screen
        If DragTopHeight > .Bottom - SnapWidth And DragTopHeight < .Bottom + SnapWidth Then
          Top = (.Bottom * Screen.TwipsPerPixelY) - Height
        ElseIf DragTop < .Top + SnapWidth And DragTop > .Top - SnapWidth Then
          Top = .Top * Screen.TwipsPerPixelY
        Else
          Top = Top + (y - MoveMainOldY)
        End If
      End With
    Else
      Left = Left + (X - MoveMainOldX)
      Top = Top + (y - MoveMainOldY)
    End If
    
    If Docked Then
      frmPlaylist.Left = Me.Left + DockedLeft
      frmPlaylist.Top = Me.Top + DockedTop
    End If
  End If
  
End Sub

Private Sub imgDropdown_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  On Error Resume Next

  If MoveMain Then MoveMain = False
  
End Sub

Public Sub imgDropdown_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  'Drag & drop
  'This will add all files dropped on the player to the playlist
  'The playlist will be cleared first, though
  On Error Resume Next
  
  Dim i As Long
  Dim File As New clsFile
  
  If Data.GetFormat(vbCFFiles) Then 'true if data is list of files
  
    If Shift = vbShiftMask Then 'if shift is down don't clear list
      frmPlaylist.imgDropdown_OLEDragDrop Data, Effect, Button, 0, X, y
      Exit Sub
    End If
  
    frmMenus.menDeleteAll_Click 'clear playlist
    
    For i = 1 To Data.Files.Count 'go through each file in dropped list
      File.sFilename = Data.Files(i)
      If File.eFileAttributes And eDIRECTORY Then
        SimpleAddDir Data.Files(i)
      Else
        Select Case LCase(File.sExtension)
          Case "playlist", "m3u", "pls"
            HandlePlaylist Data.Files(i)
          Case "mp3", "mp2", "asf", "wma", "ogg", "mid", "midi", "rmi", "mod", "xm", "it", "s3m", "wav", "sgm"
            If CreateLibrary(Data.Files(i)) = 0 Then
              MsgBox """" & Data.Files(i) & """ could not be opened.", vbExclamation, "Error in file"
            End If
          Case Else
            If Sound.TestFile(Data.Files(i)) = 0 Then
              MsgBox "The file " & Chr(34) & Data.Files(i) & Chr(34) & " is not supported.", vbExclamation
            Else
              CreateLibrary Data.Files(i)
            End If
        End Select
      End If
    Next i
    
    UpdateList
    If UBound(Playlist) > 0 Then PlayNext
  Else
    MsgBox "Whatever you dragged to Simple Amp, it isn't supported." & vbCrLf & "What you can drag here:" & vbCrLf & vbCrLf & "- One or more supported filetypes." & vbCrLf & "- One ore more folders (all files in them and their subfolders will be added)." & vbCrLf & "Dragging files to the player window will clear the playlist before adding them. Dragging files to the playlist window will add them to the list. Hold shift while dropping to do the opposite, i.e. in the player window to not clear the list.", vbInformation, "Simple Amp"
  End If
    
End Sub

Private Sub imgStereoMono_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseDown Button, Shift, X, y
End Sub

Private Sub imgStereoMono_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseMove Button, Shift, X, y
End Sub

Private Sub imgStereoMono_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseUp Button, Shift, X, y
End Sub

Private Sub imgStereoMono_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub Keys_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  
  Select Case KeyCode
    Case vbKeyF1
      frmMenus.menHelp_Click
    Case vbKeyF2
      frmAdd.Show
    Case vbKeyF3
      frmLibrary.Show
    Case vbKeyA
      frmMenus.menAdvance_Click
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
      btnPlaylist.Value = Not btnPlaylist.Value
      HideShowPlaylist
    Case vbKeyR
      btnRepeat.Value = Not btnRepeat.Value
      RepeatOnOff
    Case vbKeyS
      btnShuffle.Value = Not btnShuffle.Value
      ShuffleOnOff
    Case vbKeyLeft
      If Sound.StreamIsLoaded Then
        Position.Value = Position.Value - 5000
      ElseIf Sound.MusicIsLoaded Then
        Position.Value = Position.Value - Sound.MusicGetRows(Sound.MusicOrder)
      End If
      Position_Change
    Case vbKeyRight
      If Sound.StreamIsLoaded Then
        Position.Value = Position.Value + 5000
      ElseIf Sound.MusicIsLoaded Then
        Position.Value = Position.Value + Sound.MusicGetRows(Sound.MusicOrder)
      End If
      Position_Change
    Case vbKeyF
      If Shift = vbCtrlMask Then frmMenus.menPlistSearch_Click
    Case vbKey1
      frmPlaylist.List.ColumnSort Sort_AristTitle
    Case vbKey2
      frmPlaylist.List.ColumnSort Sort_Album
    Case vbKey3
      frmPlaylist.List.ColumnSort Sort_Genre
    Case vbKey4
      frmPlaylist.List.ColumnSort Sort_Time
    Case vbKey5
      frmPlaylist.List.ColumnSort Sort_FileName
    Case vbKey6
      frmPlaylist.List.ColumnSort Sort_FileType
    Case vbKey7
      frmPlaylist.List.ColumnSort Sort_PlayDate
    Case vbKey8
      frmPlaylist.List.ColumnSort Sort_PlayTimes
    Case vbKey9
      frmPlaylist.List.ColumnSort Sort_SkipTimes
    Case vbKey0
      frmPlaylist.List.ColumnSort Sort_OriginalOrder
    Case vbKeyUp
      Volume.Value = Volume.Value + 10
      Volume_Change
    Case vbKeyDown
      Volume.Value = Volume.Value - 10
      Volume_Change
  End Select
  
End Sub

Private Sub lblAlbum_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseDown Button, Shift, X, y
End Sub

Private Sub lblAlbum_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseMove Button, Shift, X, y
End Sub

Private Sub lblAlbum_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseUp Button, Shift, X, y
End Sub

Private Sub lblAlbum_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub lblArtistTitle_DblClick()
  If PlayingLib > 0 Then frmView.View PlayingLib
End Sub

Private Sub lblArtistTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseDown Button, Shift, X * Screen.TwipsPerPixelX, y * Screen.TwipsPerPixelY
End Sub

Private Sub lblArtistTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseMove Button, Shift, X * Screen.TwipsPerPixelX, y * Screen.TwipsPerPixelY
End Sub

Private Sub lblArtistTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseUp Button, Shift, X * Screen.TwipsPerPixelX, y * Screen.TwipsPerPixelY
End Sub

Private Sub lblArtistTitle_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub lblComments_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseDown Button, Shift, X, y
End Sub

Private Sub lblComments_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseMove Button, Shift, X, y
End Sub

Private Sub lblComments_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseUp Button, Shift, X, y
End Sub

Private Sub lblComments_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub lblGenre_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseDown Button, Shift, X, y
End Sub

Private Sub lblGenre_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseMove Button, Shift, X, y
End Sub

Private Sub lblGenre_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseUp Button, Shift, X, y
End Sub

Private Sub lblGenre_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub lblInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseDown Button, Shift, X, y
End Sub

Private Sub lblInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseMove Button, Shift, X, y
End Sub

Private Sub lblInfo_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseUp Button, Shift, X, y
End Sub

Private Sub lblInfo_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub lblTime_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseDown Button, Shift, X, y
End Sub

Private Sub lblTime_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseMove Button, Shift, X, y
End Sub

Private Sub lblTime_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseUp Button, Shift, X, y
End Sub

Private Sub lblTime_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
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

Private Sub lblYear_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseDown Button, Shift, X, y
End Sub

Private Sub lblYear_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseMove Button, Shift, X, y
End Sub

Private Sub lblYear_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_MouseUp Button, Shift, X, y
End Sub

Private Sub lblYear_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub p_GotFocus()
  ctlSetFocus Keys
End Sub

Private Sub p_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_KeyDown KeyCode, Shift
End Sub

Private Sub picVis_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_KeyDown KeyCode, Shift
End Sub

Private Sub picVis_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  On Error Resume Next
  
  If devDSP Then
    If Button = 2 Then
      PopupMenu frmMenus.menVisMenu
    Else
      If Shift = vbShiftMask Then
        lCurVisPreset = lCurVisPreset - 1
        If lCurVisPreset < 0 Then lCurVisPreset = UBound(SkinPresets)
      Else
        lCurVisPreset = lCurVisPreset + 1
        If lCurVisPreset > UBound(SkinPresets) Then lCurVisPreset = 0
      End If
      LoadSkinPreset lCurVisPreset
    End If
  End If
End Sub

Private Sub picVis_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Public Sub Position_Change()
  On Error Resume Next
  With Sound
    If .StreamIsLoaded Then
      '.StreamVolume = 0
      .StreamSongPos = Position.Value
      '.StreamVolume = Settings.CurrentVolume
    ElseIf .MusicIsLoaded Then
      Dim X As Long
      X = CInt(.MusicNumOrders * (Position.Value / Position.Max))
      If X <> .MusicOrder Then
        .MusicOrder = X
      End If
    End If
    If Library(PlayingLib).eType = TYPE_CDA Then
      Sound.CDSetTrackTime Position.Value * 1000
    End If
  End With
End Sub

Private Sub Position_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_KeyDown KeyCode, Shift
End Sub

Private Sub Position_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub tmrCDAudio_Timer()
  'tries to determine when an CDA track has completed.
  'fmod can't track this by itself.
  'this works bad.
  On Error Resume Next
  If Library(PlayingLib).eType = TYPE_CDA Then
    If FSOUND_CD_GetPaused(0) = False And Not ManStopped Then
      Dim X As Long
      X = CDTime
      CDTime = Sound.CDGetTrackTime
      If X = CDTime And Settings.Advance Then  'next song
        If Settings.Repeat = 2 Then
          Play
        Else
          PlayNext
        End If
      End If
    End If
  End If
End Sub

Private Sub tmrMain_Timer()
  'this timer refreshes lables & other misc controls.
  'it also checks for end of music.
  On Error Resume Next
  Dim X As Long, y As Long

  With Sound
  
    If .StreamIsPlaying And WindowState = vbNormal And Visible Then
      Position.Max = .StreamSongLen
      Position.Value = .StreamSongPos
      lblTime = ConvertTime(.StreamSongPos \ 1000)
      lblTotalTime = ConvertTime(.StreamSongLen \ 1000)
    ElseIf .MusicIsPlaying And WindowState = vbNormal And Visible Then
      'calculate slider position
      For X = 0 To .MusicNumOrders - 1
        y = y + .MusicGetRows(X)
      Next
      Position.Max = y
      y = 0
      For X = 1 To .MusicOrder
        y = y + .MusicGetRows(X)
      Next
      y = y + .MusicRow
      Position.Value = y
      
      lblTime = ConvertTime(.MusicSongPos \ 1000)
      lblTotalTime = ConvertTime(.MusicSongLen \ 1000)
      
      'Update labels
      lblYear = .MusicPattern & "\" & .MusicNumPatterns
      lblGenre = "BPM: " & .MusicBPM
      lblComments = .MusicNumIntruments & " Instruments, " & .MusicNumSamples & " Samples, " & .ChannelsPlaying & "/" & .MusicNumChannels & " Channels, Speed: " & .MusicSpeed
      lblInfo = .MusicOrder & "\" & .MusicNumOrders & ", " & .MusicRow & "\" & .MusicNumRows
    End If
    
    If Library(PlayingLib).eType = TYPE_CDA Then
      Position.Max = Library(PlayingLib).lLength
      Position.Value = Sound.CDGetTrackTime \ 1000
      lblTime = ConvertTime(Position.Value)
      lblTotalTime = ConvertTime(Library(PlayingLib).lLength)
    End If
    
    'this tries to track when a stream has ended
    'works pretty good for such an half-ass way when you could use callbacks
    'callback does not work atm, it just crashes!
    If Not StreamStart Then
      If .StreamIsLoaded And .StreamSongPos = .StreamSongLen And Settings.Advance Then
        If Settings.Repeat = 2 Then
          Play
        Else
          PlayStop
          PlayNext
        End If
      ElseIf .StreamIsLoaded And .StreamSongPos = 0 And Not ManStopped And Settings.Advance Then
        If Settings.Repeat = 2 Then
          Play
        Else
          PlayStop
          PlayNext
        End If
      End If
    End If
    If .MusicIsFinished And Settings.Advance Then
      If Settings.Repeat = 2 Then
        Play
      Else
        PlayNext
      End If
    End If
  
  End With
  
End Sub

Private Sub tmrScope_Timer()
  On Error Resume Next
  'This timer controls the visualizations.

  If Me.WindowState = vbNormal And Me.Visible And devDSP Then
  
    Select Case Spectrum
      Case 1
        If DSP_OK Then visScope
      Case 2
        If DSP_OK Then visSpectrum
      Case 3
        visVolume
      Case 4
        If DSP_OK Then visBeat
    End Select
    
  Else
    Set picVis.Picture = GetImage(SPEC_BGOFF)
  End If
End Sub

Private Sub tmrStream_Timer()
  'stupid timer that helps to detect when an stream has ended
  'tmrMain_Timer does the rest of the checking
  'would use callbacks if it didn't crash everything...
  On Error Resume Next
  StreamStart = False
  tmrStream.Enabled = False
End Sub

Public Sub Volume_Change()
  'volume change
  On Error Resume Next
  
  'simple workaround to an bug in CtrlScroller that would let Volume.Value
  'get higher than Volume.Max, which Sound.StreamVolume could not handle,
  'thus setting the volume to 0 instead.
  If Volume.Value > 255 Then Volume.Value = 255
  
  Settings.CurrentVolume = Volume.Value
  Sound.StreamVolume = Volume.Value
  Sound.MusicVolume = Volume.Value
  If Library(PlayingLib).eType = TYPE_CDA Then
    Sound.CDSetVolume Volume.Value
  End If
End Sub

Public Sub Play()
  'Plays the current song
  On Error GoTo ErrHandler
  Dim X As Long, tL As String

  'Playing is the playing item in playlist array
  'PlayingLib is the playing item in Library
  If (Playing < 1 Or Playing > UBound(Playlist)) And PlayingLib = 0 Then Exit Sub
  If Playing > 0 And Playing <= UBound(Playlist) Then
    PlayingLib = Playlist(Playing).Reference
  End If
  
  'file not found
  If Not FileExists(Library(PlayingLib).sFilename) And Not Library(PlayingLib).eType = TYPE_CDA Then
    MsgBox """" & Library(PlayingLib).sFilename & """ was not found.", vbExclamation, "File not found"
    Exit Sub
  End If
  
  'reset shit
  tmrMain.Enabled = False
  tmrStream.Enabled = False
  StreamStart = True
  
  If CDTime <> -1 Then Sound.CDStop
  If Sound.MusicIsPlaying Then Sound.MusicUnload
  If Sound.StreamIsPlaying Then Sound.StreamUnload
  
  'set up shuffle indexes
  If Playlist(Playing).lShuffleIndex = 0 Then
    Playlist(Playing).lShuffleIndex = ShuffleNum
  End If
  
  lblArtistTitle.Caption = "Artist & Title N/A"
  lblAlbum = "Album N/A"
  lblYear = "N/A"
  lblGenre = "N/A"
  lblComments = "Comments N/A"
  lblInfo = ""
  tmrCDAudio.Enabled = False
  CDTime = -1
  
  'Refresh this items data in the library
  If Not Library(PlayingLib).eType = TYPE_CDA Then UpdateLibrary PlayingLib
  
  'check type
  With Library(PlayingLib)
    'MODULES
    If .eType = TYPE_IT Or .eType = TYPE_MOD Or .eType = TYPE_S3M Or .eType = TYPE_XM Then
    
      'play
      If Not Sound.MusicPlay(.sFilename, False, Settings.CurrentVolume) Then Exit Sub
      
      'setup labels & misc.
      lblArtistTitle.Caption = .sArtistTitle
      lblAlbum = "Title: " & .sTitle
      If Len(.sTitle) > 0 Then
        cTray.ToolTip = .sTitle
      Else
        cTray.ToolTip = lblArtistTitle.Caption
      End If
      
      CurMono = False

    'CD audio (works bad, but mp3's the thing now you know, forget your old discs!)
    ElseIf .eType = TYPE_CDA Then
      
      Dim trk As Long
      trk = Val(Right(.sFilename, Len(.sFilename) - 17))
      
      If Not Sound.CDPlay(trk) Then
        MsgBox "Could not play track " & trk & " on an Audio CD in the default drive.", vbExclamation
        Exit Sub
      End If
      Sound.CDSetVolume Settings.CurrentVolume
      
      .lLength = Sound.CDGetTrackLength(trk) \ 1000
      lblArtistTitle.Caption = .sArtistTitle
      cTray.ToolTip = .sArtistTitle
      lblInfo = "44100 hz"
      CurMono = True
      tmrCDAudio.Enabled = True
    
    'MIDIs
    ElseIf .eType = TYPE_MID_RMI Or .eType = TYPE_SGM Then
      'play
      If Not Sound.StreamPlay(.sFilename, , Settings.CurrentVolume) Then Exit Sub
      
      'setup labels & misc.
      lblArtistTitle.Caption = .sArtistTitle
      cTray.ToolTip = lblArtistTitle.Caption
      
      CurMono = False
      
    'MP2, MP3
    ElseIf .eType = TYPE_MP2_MP3 Then
      Dim MP3 As New cMP3Info
   
      'Find out if this is stereo or mono
      MP3.FileName = .sFilename
      MP3.ReadMP3Header
      CurMono = (MP3.Mode = "Mono")
      
      'play
      If Not Sound.StreamPlay(.sFilename, .bIsVBR, Settings.CurrentVolume) Then Exit Sub
      
      'Update labels & misc.
      lblArtistTitle.Caption = .sArtistTitle
      If Len(.sAlbum) > 0 Then lblAlbum = .sAlbum
      If Len(.sComments) > 0 Then lblComments = .sComments
      If Len(.sGenre) > 0 Then lblGenre = .sGenre
      If Len(.sYear) > 0 Then lblYear = .sYear
      
      cTray.ToolTip = lblArtistTitle.Caption
      lblInfo = Sound.StreamKbps & "kbps, " & Sound.StreamFrequency \ 1000 & "khz"

    'OGG
    ElseIf .eType = TYPE_OGG Then
      
      'play
      If Not Sound.StreamPlay(.sFilename, , Settings.CurrentVolume) Then Exit Sub
      
      'Update labels & misc.
      lblArtistTitle.Caption = .sArtistTitle
      If Len(.sAlbum) > 0 Then lblAlbum = .sAlbum
      If Len(.sComments) > 0 Then lblComments = .sComments
      If Len(.sGenre) > 0 Then lblGenre = .sGenre
      If Len(.sYear) > 0 Then lblYear = .sYear
      
      cTray.ToolTip = lblArtistTitle.Caption
      lblInfo = Sound.StreamKbps & "kbps, " & Sound.StreamFrequency \ 1000 & "khz"
      CurMono = False
    
    'ASF, WMA
    ElseIf .eType = TYPE_ASF Or .eType = TYPE_WMA Then
      'play
      If Not Sound.StreamPlay(.sFilename, , Settings.CurrentVolume) Then Exit Sub
      
      'Update labels & misc.
      lblArtistTitle.Caption = .sArtistTitle
      CurMono = False
      cTray.ToolTip = lblArtistTitle.Caption
      lblInfo = Sound.StreamKbps & "kbps, " & Sound.StreamFrequency \ 1000 & "khz"
    
    'WAV
    ElseIf .eType = TYPE_WAV Then
    
      'play
      If Not Sound.StreamPlay(.sFilename, , Settings.CurrentVolume) Then Exit Sub
      
      'Update labels & misc.
      lblArtistTitle.Caption = .sArtistTitle
      CurMono = False
      cTray.ToolTip = lblArtistTitle.Caption
      lblInfo = Sound.StreamKbps & "kbps, " & Sound.StreamFrequency \ 1000 & "khz"
    
    'Type is unknown, skip
    Else
      PlayNext
      Exit Sub
    End If
  
    .dLastPlayDate = Now
    .lTimesPlayed = .lTimesPlayed + 1
    Sound.StreamSurround = devSurround
    Sound.StreamPanning = devPanning
    Sound.MusicPanSep = devPanSep
    ManStopped = False

  End With

  'Init Equalizer if it is on
  If Settings.EQon And EQHandle(0) = 0 Then UpdateFX
  
  'reset the pitch
  If frmStudio.Visible Then
    frmStudio.ResetPitch
  End If
  
  'update forms & controls
  Position.Enabled = True
  For X = 1 To UBound(Playlist)
    Playlist(X).IsBold = False
  Next X
  Playlist(Playing).IsBold = True
  frmPlaylist.List.MakeVisible Playlist(Playing).index
  frmPlaylist.List.Refresh
  
  'Show pause button instead of play
  UpdatePlayPause True
  
  'Update Stereo/Mono indicator
  If CurMono Then
    imgStereoMono.Picture = GetImage(SI_MAIN_MONO)
  Else
    imgStereoMono.Picture = GetImage(SI_MAIN_STEREO)
  End If
  
  'Update main text colors
  UpdateMainColor
  
  tmrMain.Enabled = True
  tmrStream.Enabled = True
   
  Exit Sub
ErrHandler:
  If cLog.ErrorMsg(Err, "frmMain, Play()") = vbYes Then Resume Next Else UnloadAll
End Sub

Public Sub PlayNext()
  'This plays the next song in the list
  On Error Resume Next
  
  StreamStart = False
  
  If Sound.StreamIsPlaying Or Sound.MusicIsPlaying Then
    Library(PlayingLib).lTimesSkipped = Library(PlayingLib).lTimesSkipped + 1
  End If
  
  Dim X As Long, y As Long
  X = Playing
  
  If Settings.Shuffle Then
    If Settings.Shuffle = 1 Then 'ordered shuffle
      If ShuffleNum = UBound(Playlist) And Settings.Repeat <> 1 Then Exit Sub
      ShuffleNum = ShuffleNum + 1
      If ShuffleNum > UBound(Playlist) Then ShuffleNum = 1
      Playing = GetShuffle
    Else 'random shuffle (pretty useless actually)
      If UBound(Playlist) > 0 Then
        Do
          Playing = CInt(Rnd * UBound(Playlist)) + 1
        Loop Until Playing <> X
      End If
    End If
  Else
    Playing = Playlist(Playing).index + 1
    'handle repeat list
    If Playing > UBound(Playlist) Then
      If Settings.Repeat = 1 Then
        Playing = 1
      Else
        Playing = UBound(Playlist)
      End If
    End If
    Playing = frmPlaylist.List.GetNumFromIndex(Playing)
  End If
  
  If Playing <> X Then
    Play
  ElseIf Playing = X And Settings.Repeat = 1 Then
    Play
  End If
  
End Sub

Public Sub PlayPrev()
  'Plays previous item in playlist
  On Error Resume Next
  
  If Sound.StreamIsPlaying Or Sound.MusicIsPlaying Then
    Library(PlayingLib).lTimesSkipped = Library(PlayingLib).lTimesSkipped + 1
  End If
  
  Dim X As Long
  X = Playing
  
  If Settings.Shuffle Then
    If Settings.Shuffle = 1 Then 'ordered shuffle
      If ShuffleNum = 1 And Settings.Repeat <> 1 Then Exit Sub
      ShuffleNum = ShuffleNum - 1
      If ShuffleNum < 1 Then ShuffleNum = UBound(Playlist)
      Playing = GetShuffle
    Else 'random shuffle (pretty useless actually)
      If UBound(Playlist) > 0 Then
        Do
          Playing = CInt(Rnd * UBound(Playlist)) + 1
        Loop Until Playing <> X
      End If
    End If
  Else
    Playing = Playlist(Playing).index - 1
    'handle repeat list
    If Playing < 1 Then
      If Settings.Repeat = 1 Then
        Playing = UBound(Playlist)
      Else
        Playing = 1
      End If
    End If
    Playing = frmPlaylist.List.GetNumFromIndex(Playing)
  End If
  
  If Playing <> X Then
    Play
  ElseIf Playing = X And Settings.Repeat = 1 Then
    Play
  End If
  
End Sub

Public Sub SetupSystray(ByVal Mode As Boolean)
  On Error Resume Next
  'This sub just changes the ctray mode, used from another form,
  'as cTray can only be used in this one (private class with events)
  If Mode Then
    cTray.IconInSysTray
  Else
    cTray.RemoveFromSysTray
  End If
End Sub

Public Sub SetupSystrayIcon(ByVal Icon As Byte)
  On Error Resume Next
  'This sub just changes the ctray icon, used from another form,
  'as cTray can only be used in this one (private class with events)
  cTray.Icon = LoadResPicture(Icon + 104, vbResIcon)
End Sub

Public Sub UnloadAll()
  'This sub saves settings and unloads all data
  On Error Resume Next
  Dim X As Long
  
  'SAVE STUFF
  
  SaveSettings
  
  PlayStop
  Me.Hide 'hide simple amp as fast as possible so it won't bother the user
          'allthough saving library etc. will continue in the background
  If frmPlaylist.Visible Then frmPlaylist.Hide
  DoEvents
  
  'save library
  If Settings.UseLibrary Then
    If LibraryChanged Then SaveLibrary 'only save if it has changed!
    SaveStat
  End If
  
  'save current playlist
  If UBound(Playlist) > 0 Then
    Call SavePlaylist(App.Path & "\current.playlist")
  Else
    If FileExists(App.Path & "\current.playlist") Then Kill App.Path & "\current.playlist"
  End If
  
  'UNLOAD STUFF
  
  cLog.Log "UNLOADING SIMPLE AMP...", 5, False
  
  'make sure systray is unloaded
  cTray.RemoveFromSysTray
  Set cTray = Nothing
  'make sure fmod sound system is unloaded
  Sound.StreamUnload
  Sound.MusicUnload
  If CDTime <> 0 Then Sound.CDStop
  If devDSP Then
    FSOUND_DSP_SetActive FSOUND_DSP_GetFFTUnit, False
    FSOUND_DSP_SetActive DSP_Handle, False
    FSOUND_DSP_Free DSP_Handle
  End If
  Set Sound = Nothing
  'unload skin, don't think this is really neccesary
  For X = 1 To UBound(Images)
    Set Images(X) = Nothing
  Next X
  'delete pens & brushes
  DeleteObject hPenLeft
  DeleteObject hPenRight
  DeleteObject hPenPeakLeft
  DeleteObject hPenPeakRight
  DeleteObject hBrushSolidLeft
  DeleteObject hBrushSolidRight
  DeleteObject hPenPeakSpec
  DeleteObject hPenLineSpec
  'unload forms
  Unload frmAbout
  Unload frmDDE
  Unload frmView
  Unload frmMenus
  Unload frmPlaylist
  Unload frmSettings
  Unload frmSkin
  Unload frmStudio
  frmAdd.bWorking = False 'make sure it exits even if it is working
  Unload frmAdd
  frmLibrary.LibClose = True
  Unload frmLibrary
  Unload frmLibClean
  Unload frmLibSettings
  Unload frmPresetEdit
  Unload frmPresetLoad
  
  cLog.Log "BYE!", 5
  
  Unload Me

End Sub

Public Sub PlayStop()
  'This sub stops the music
  On Error Resume Next
  
  lblTime = "00:00"
  Position.Enabled = False
  Position.Value = 1
  UpdatePlayPause False
  
  
  If Sound.StreamIsLoaded Then
    Sound.StreamUnload
  End If
  If Sound.MusicIsLoaded Then
    Sound.MusicUnload
  End If
  If Library(PlayingLib).eType = TYPE_CDA Then
    Sound.CDStop
  End If
  
  UpdateMainColor
  PlayingLib = 0
  
End Sub

Public Sub Minimize()
  'This sub minimizes the main window and the playlist window
  On Error Resume Next
  
  Me.Hide
  If Settings.PlaylistOn Then frmPlaylist.Hide
  cTray.IconInSysTray
  
End Sub

Public Sub PlayPause()
  'This sub pauses/plays the music
  On Error Resume Next
  
  If Library(PlayingLib).eType = TYPE_CDA Then
    Sound.CDPausePlay
    UpdatePlayPause Not CBool(FSOUND_CD_GetPaused(0))
    Exit Sub
  End If
  
  If Not Sound.StreamIsLoaded And Not Sound.MusicIsLoaded Then Play: Exit Sub
  Sound.StreamPausePlay
  Sound.MusicPausePlay
  
  If Sound.StreamIsPlaying Or Sound.MusicIsPlaying Then
    UpdatePlayPause True
  Else
    UpdatePlayPause False
  End If
      
End Sub

Public Sub HideShowPlaylist()
  'This sub Shows/hides the playlist window
  On Error Resume Next
  
  Settings.PlaylistOn = btnPlaylist.Value
  frmMenus.menPlaylist.Checked = Settings.PlaylistOn
  
  If Settings.PlaylistOn Then
    frmPlaylist.Show
    AlwaysOnTop frmPlaylist, Settings.OnTop
  Else
    frmPlaylist.Hide
  End If
  
End Sub

Public Sub RepeatOnOff()
  'This sub turns Repeat On/off
  On Error Resume Next
  Dim X As Long
  
  If Settings.Repeat > 0 Then
    btnRepeat.Value = True
  Else
    btnRepeat.Value = False
  End If
  For X = 0 To 2
    frmMenus.menRepeatVal(X).Checked = False
  Next
  frmMenus.menRepeatVal(Settings.Repeat).Checked = True
  
End Sub

Public Sub ShuffleOnOff()
  'This sub turns shuffle On/off
  On Error Resume Next
  Dim X As Long

  If Settings.Shuffle > 0 Then
    btnShuffle.Value = True
  Else
    btnShuffle.Value = False
  End If
  For X = 0 To 2
    frmMenus.menShuffleVal(X).Checked = False
  Next
  frmMenus.menShuffleVal(Settings.Shuffle).Checked = True
  
End Sub

Public Sub UpdateSpectrum()
  On Error Resume Next
  
  If Spectrum = 0 Then
    'cBack.ClonePicture GetImage(SPEC_BGOFF)
    'cBackOrig.ClonePicture GetImage(SPEC_BGOFF)
    Set picVis.Picture = GetImage(SPEC_BGOFF)
  ElseIf Spectrum = 1 Then
    cBack.ClonePicture GetImage(SPEC_BG1)
    cBackOrig.ClonePicture GetImage(SPEC_BG1)
  ElseIf Spectrum = 2 Then
    cBack.ClonePicture GetImage(SPEC_BG2)
    cBackOrig.ClonePicture GetImage(SPEC_BG2)
  ElseIf Spectrum = 3 Then
    If VolumeSettings.bType = 0 Then
      cBack.ClonePicture GetImage(SPEC_BG3)
      cBackOrig.ClonePicture GetImage(SPEC_BG3)
    ElseIf VolumeSettings.bType = 1 Then
      cBack.ClonePicture GetImage(SPEC_BG4)
      cBackOrig.ClonePicture GetImage(SPEC_BG4)
    End If
  ElseIf Spectrum = 4 Then
    cBack.ClonePicture GetImage(SPEC_BG5)
    cBackOrig.ClonePicture GetImage(SPEC_BG5)
    
    'Init the beat image (with the possibility of using external bmp)
    cBeat.Destroy
    cBeatTemp.Destroy
    If Len(BeatSettings.sFile) > 0 And FileExists(App.Path & "\vis\data\" & BeatSettings.sFile) Then
      cBeat.LoadImg App.Path & "\vis\data\" & BeatSettings.sFile
      cBeatTemp.LoadImg App.Path & "\vis\data\" & BeatSettings.sFile
    Else
      If BeatSettings.lImage > 0 Then
        'temp solution... cBeat.ClonePicture Images(BeatSettings.lImage)
        'hangs for some reason...
        SavePicture Images(BeatSettings.lImage), App.Path & "\t.bmp"
        cBeat.LoadImg App.Path & "\t.bmp"
        cBeatTemp.LoadImg App.Path & "\t.bmp"
        Kill App.Path & "\t.bmp"
      Else
        cBeat.ClonePicture GetImage(SPEC_BEAT)
        cBeatTemp.ClonePicture GetImage(SPEC_BEAT)
      End If
    End If
    cBeat.InitSA
    cBeatTemp.InitSA
  End If
  
  cBack.InitFade cBackOrig

End Sub

Public Sub Init()
  On Error GoTo ErrHandler
  
  Dim X As Long
  Set cTray = New clsSysTray
  
  cLog.Log "INITIALIZING SIMPLE AMP.", 3
  
  Caption = "Simple Amp " & App.Major & "." & App.Minor
  Load frmDDE 'hooks up simple amp to recieve messages from other simple amps
  
  'StreamStart = True
  
  ReDim SpectrumPeaks(0)
  ReDim SkinPresets(0)
  
  'Init random number generator
  Randomize Timer
  
  'Setup systray
  Set cTray.SourceWindow = Me
  cTray.Icon = LoadResPicture(Settings.TrayIcon + 104, vbResIcon)  'Set icon
  cTray.ToolTip = "Simple Amp " & App.Major & "." & App.Minor 'Set tip text
  cTray.DefaultDblClk = False
  If Settings.AlwaysTray Then cTray.IconInSysTray
  
  'Loads skin
  LoadSkin CurrentSkin
  
  If Settings.PlaylistOn And Not Settings.StartInTray Then
    If Settings.Fade Then MakeTransparent frmPlaylist.hwnd, 0
    frmPlaylist.Show
    AlwaysOnTop frmPlaylist, Settings.OnTop
  End If
  If Settings.StartInTray Then
    Minimize
  Else
    'Sets Ontop
    If Settings.Fade Then MakeTransparent Me.hwnd, 0
    Show
  End If
  
  If Sound.FMODVersion <> FMOD_VERSION Then
    cLog.Log "WARNING: WRONG FMOD VERSION... FOUND: " & Sound.FMODVersion, 10
    cLog.Log "SIMPLE AMP WAS WRITTEN FOR " & FMOD_VERSION & ", PROBLEMS MAY OCCUR!", 10
  End If
  
  With Sound
    
    cLog.Log "STARTING FMOD...", 5, False
    
    'Get devices for each output type
    .InitOutput FSOUND_OUTPUT_DSOUND
    .ReturnDrivers Direct_Dev
    .InitOutput FSOUND_OUTPUT_WINMM
    .ReturnDrivers WinOut_Dev
  
    'Init sound
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
    .Init devFreq, devChannels + 1
    If devDSP Then
      ReDim ScopeBufferFPU(FSOUND_DSP_GetBufferLength)
      ReDim ScopeBufferINT(FSOUND_DSP_GetBufferLength)
      ReDim ScopeUPeaks(UBound(ScopeBufferFPU))
      ReDim ScopeLPeaks(UBound(ScopeBufferFPU))
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
    DoEvents
    
    cLog.Log "DONE.", 5
    cLog.Log "RUNNING WITH " & devBuffer & " MS BUFFER.", 1
    cLog.Log "DIGITAL SOUND PROCESSING IS " & IIf(DSP_OK, "ON.", "OFF."), 1
    cLog.Log "DIRECT X 8 EFFECTS IS " & IIf(Settings.DXFXon, "ON.", "OFF."), 1
  End With
  
  'set volume
  Volume.Value = Settings.CurrentVolume
  ShuffleNum = 1
  
  tmrScope.Interval = VisUpdateInt
  frmMenus.menPlaylist.Checked = Settings.PlaylistOn
  frmMenus.menAdvance.Checked = Settings.Advance
  ShuffleOnOff
  RepeatOnOff
  
  tmrScope.Enabled = True

  If Me.Visible Then
    If Settings.Fade Then MakeTransparent Me.hwnd, 0
    Me.Show
    AlwaysOnTop Me, Settings.OnTop
  End If
  
  'fades in window (could take a while on slower computers...)
  If Me.Visible And Settings.Fade Then
    For X = 0 To 250 Step 10
      MakeTransparent Me.hwnd, X
      If frmPlaylist.Visible Then MakeTransparent frmPlaylist.hwnd, X
      DoEvents
    Next X
    MakeTransparent Me.hwnd, 255
    If frmPlaylist.Visible Then MakeTransparent frmPlaylist.hwnd, 255
  End If

  'This loads the autosave playlist if there is one
  If FileExists(App.Path & "\current.playlist") And Len(FileOpen) = 0 Then
    cLog.Log "RESTORING PLAYLIST", 3
    If LoadPlaylist(App.Path & "\current.playlist") Then
      UpdateList
    End If
  End If

  'open files added from command line
  If Len(FileOpen) > 0 Then
    SimpleAddFile FileOpen
  Else
    If UBound(Playlist) > 0 Then Play
  End If
  
  cLog.Log "SIMPLE AMP STARTED... PRAISE THE LORD! =D", 5
  
  Exit Sub
ErrHandler:
  If cLog.ErrorMsg(Err, "frmMain, Init()") = vbYes Then Resume Next Else UnloadAll
End Sub

Public Sub LoadSkin(ByVal Skin As String, Optional bDefPreset As Boolean = False)
  'Loads skin 'Skin' by reading from app.path & "\skins\" & skin & ".sas"
  On Error Resume Next
  Dim File As New clsDatafile
  Dim tmpFile As String
  Dim cFnt As New StdFont
  
  lblArtistTitle.AutoScroll = False
  
  cLog.Log "LOADING SKIN (" & Skin & "...", 3, False
  cLog.StartTimer
  
  Dim X As Long, y As Long
  For X = 0 To UBound(Images)
    Set Images(X) = Nothing
  Next
  For X = 1 To UBound(SkinImage)
    SkinImage(X) = 0
  Next
  
  On Error GoTo SkinError
  
  With File
  
    tmpFile = App.Path & "\img.tmp"
    
    LockWindowUpdate Me.hwnd
    
    .FileName = App.Path & "\skins\" & Skin & ".sas"
    If .ReadStrFixed(3) <> "SAS" Then GoTo SkinError
    If .ReadNumber <> COMPILEVERSION Then Err.Raise 30335
    .SkipField 5 'skip 5 fields (all the skin info & preview image)
    
    'Read all images
    ReDim Images(.ReadNumber)
    For X = 1 To UBound(Images)
      .ReadFile tmpFile
      Set Images(X) = LoadPicture(tmpFile)
    Next
    
    'read main player background
    Me.Height = .ReadNumber * Screen.TwipsPerPixelY
    Me.Width = .ReadNumber * Screen.TwipsPerPixelX
    Call .ReadNumber 'x
    Call .ReadNumber 'y
    MainTrans = CBool(.ReadNumber)
    SkinImage(1) = .ReadNumber
    'read playlist background
    PlaylistWin(0).Self.H = .ReadNumber * Screen.TwipsPerPixelY
    PlaylistWin(0).Self.W = .ReadNumber * Screen.TwipsPerPixelX
    Call .ReadNumber 'x
    Call .ReadNumber 'y
    PlaylistWin(0).Trans = CBool(.ReadNumber)
    SkinImage(2) = .ReadNumber
    'read playlist columns
    PlaylistWin(0).iColumn.H = .ReadNumber
    PlaylistWin(0).iColumn.W = .ReadNumber
    PlaylistWin(0).iColumn.X = .ReadNumber
    PlaylistWin(0).iColumn.y = .ReadNumber
    Call .ReadNumber 'trans
    SkinImage(3) = .ReadNumber
    'read stereo/mono
    imgStereoMono.Height = .ReadNumber
    imgStereoMono.Width = .ReadNumber
    imgStereoMono.Left = .ReadNumber
    imgStereoMono.Top = .ReadNumber
    Call .ReadNumber 'trans
    SkinImage(4) = .ReadNumber
    Call .ReadNumber
    Call .ReadNumber
    Call .ReadNumber
    Call .ReadNumber
    Call .ReadNumber 'trans
    SkinImage(5) = .ReadNumber
    'read playlist2 background
    PlaylistWin(1).Self.H = .ReadNumber * Screen.TwipsPerPixelY
    PlaylistWin(1).Self.W = .ReadNumber * Screen.TwipsPerPixelX
    Call .ReadNumber 'x
    Call .ReadNumber 'y
    PlaylistWin(1).Trans = CBool(.ReadNumber)
    SkinImage(6) = .ReadNumber
    'read playlist2 columns
    PlaylistWin(1).iColumn.H = .ReadNumber
    PlaylistWin(1).iColumn.W = .ReadNumber
    PlaylistWin(1).iColumn.X = .ReadNumber
    PlaylistWin(1).iColumn.y = .ReadNumber
    Call .ReadNumber 'trans
    SkinImage(7) = .ReadNumber
    'read position scroller
    Position.Height = .ReadNumber
    Position.Width = .ReadNumber
    Position.Left = .ReadNumber
    Position.Top = .ReadNumber
    SkinImage(8) = .ReadNumber
    SkinImage(9) = .ReadNumber
    SkinImage(10) = .ReadNumber
    SkinImage(98) = .ReadNumber
    'read volume scroller
    Volume.Height = .ReadNumber
    Volume.Width = .ReadNumber
    Volume.Left = .ReadNumber
    Volume.Top = .ReadNumber
    SkinImage(11) = .ReadNumber
    SkinImage(12) = .ReadNumber
    SkinImage(13) = .ReadNumber
    SkinImage(99) = .ReadNumber
    'read playlist window scroller
    PlaylistWin(0).sScroll.H = .ReadNumber
    PlaylistWin(0).sScroll.W = .ReadNumber
    PlaylistWin(0).sScroll.X = .ReadNumber
    PlaylistWin(0).sScroll.y = .ReadNumber
    SkinImage(14) = .ReadNumber
    SkinImage(15) = .ReadNumber
    SkinImage(16) = .ReadNumber
    SkinImage(100) = .ReadNumber
    'read playlist window 2 scroller
    PlaylistWin(1).sScroll.H = .ReadNumber
    PlaylistWin(1).sScroll.W = .ReadNumber
    PlaylistWin(1).sScroll.X = .ReadNumber
    PlaylistWin(1).sScroll.y = .ReadNumber
    SkinImage(17) = .ReadNumber
    SkinImage(18) = .ReadNumber
    SkinImage(19) = .ReadNumber
    SkinImage(101) = .ReadNumber
    'read playlist add button
    PlaylistWin(0).bAdd.H = .ReadNumber
    PlaylistWin(0).bAdd.W = .ReadNumber
    PlaylistWin(0).bAdd.X = .ReadNumber
    PlaylistWin(0).bAdd.y = .ReadNumber
    SkinImage(20) = .ReadNumber
    SkinImage(21) = .ReadNumber
    SkinImage(22) = .ReadNumber
    'read playlist remove button
    PlaylistWin(0).bRemove.H = .ReadNumber
    PlaylistWin(0).bRemove.W = .ReadNumber
    PlaylistWin(0).bRemove.X = .ReadNumber
    PlaylistWin(0).bRemove.y = .ReadNumber
    SkinImage(23) = .ReadNumber
    SkinImage(24) = .ReadNumber
    SkinImage(25) = .ReadNumber
    'read playlist select button
    PlaylistWin(0).bSelect.H = .ReadNumber
    PlaylistWin(0).bSelect.W = .ReadNumber
    PlaylistWin(0).bSelect.X = .ReadNumber
    PlaylistWin(0).bSelect.y = .ReadNumber
    SkinImage(26) = .ReadNumber
    SkinImage(27) = .ReadNumber
    SkinImage(28) = .ReadNumber
    'read playlist select button
    PlaylistWin(0).bList.H = .ReadNumber
    PlaylistWin(0).bList.W = .ReadNumber
    PlaylistWin(0).bList.X = .ReadNumber
    PlaylistWin(0).bList.y = .ReadNumber
    SkinImage(29) = .ReadNumber
    SkinImage(30) = .ReadNumber
    SkinImage(31) = .ReadNumber
    'read main previous
    btnPrev.Height = .ReadNumber
    btnPrev.Width = .ReadNumber
    btnPrev.Left = .ReadNumber
    btnPrev.Top = .ReadNumber
    SkinImage(32) = .ReadNumber
    SkinImage(33) = .ReadNumber
    SkinImage(34) = .ReadNumber
    'read main next
    btnNext.Height = .ReadNumber
    btnNext.Width = .ReadNumber
    btnNext.Left = .ReadNumber
    btnNext.Top = .ReadNumber
    SkinImage(35) = .ReadNumber
    SkinImage(36) = .ReadNumber
    SkinImage(37) = .ReadNumber
    'read main play/pause
    btnPlayPause.Height = .ReadNumber
    btnPlayPause.Width = .ReadNumber
    btnPlayPause.Left = .ReadNumber
    btnPlayPause.Top = .ReadNumber
    SkinImage(38) = .ReadNumber
    SkinImage(39) = .ReadNumber
    SkinImage(40) = .ReadNumber
    Call .ReadNumber
    Call .ReadNumber
    Call .ReadNumber
    Call .ReadNumber
    SkinImage(41) = .ReadNumber
    SkinImage(42) = .ReadNumber
    SkinImage(43) = .ReadNumber
    'read main stop
    btnStop.Height = .ReadNumber
    btnStop.Width = .ReadNumber
    btnStop.Left = .ReadNumber
    btnStop.Top = .ReadNumber
    SkinImage(44) = .ReadNumber
    SkinImage(45) = .ReadNumber
    SkinImage(46) = .ReadNumber
    'read main close
    btnClose.Height = .ReadNumber
    btnClose.Width = .ReadNumber
    btnClose.Left = .ReadNumber
    btnClose.Top = .ReadNumber
    SkinImage(47) = .ReadNumber
    SkinImage(48) = .ReadNumber
    SkinImage(49) = .ReadNumber
    'read main close
    btnMinimize.Height = .ReadNumber
    btnMinimize.Width = .ReadNumber
    btnMinimize.Left = .ReadNumber
    btnMinimize.Top = .ReadNumber
    SkinImage(50) = .ReadNumber
    SkinImage(51) = .ReadNumber
    SkinImage(52) = .ReadNumber
    'read playlist close
    PlaylistWin(0).bClose.H = .ReadNumber
    PlaylistWin(0).bClose.W = .ReadNumber
    PlaylistWin(0).bClose.X = .ReadNumber
    PlaylistWin(0).bClose.y = .ReadNumber
    SkinImage(53) = .ReadNumber
    SkinImage(54) = .ReadNumber
    SkinImage(55) = .ReadNumber
    'read playlist size
    PlaylistWin(0).bSize.H = .ReadNumber
    PlaylistWin(0).bSize.W = .ReadNumber
    PlaylistWin(0).bSize.X = .ReadNumber
    PlaylistWin(0).bSize.y = .ReadNumber
    SkinImage(56) = .ReadNumber
    SkinImage(57) = .ReadNumber
    SkinImage(58) = .ReadNumber
    'read playlist2 add
    PlaylistWin(1).bAdd.H = .ReadNumber
    PlaylistWin(1).bAdd.W = .ReadNumber
    PlaylistWin(1).bAdd.X = .ReadNumber
    PlaylistWin(1).bAdd.y = .ReadNumber
    SkinImage(59) = .ReadNumber
    SkinImage(60) = .ReadNumber
    SkinImage(61) = .ReadNumber
    'read playlist2 remove
    PlaylistWin(1).bRemove.H = .ReadNumber
    PlaylistWin(1).bRemove.W = .ReadNumber
    PlaylistWin(1).bRemove.X = .ReadNumber
    PlaylistWin(1).bRemove.y = .ReadNumber
    SkinImage(62) = .ReadNumber
    SkinImage(63) = .ReadNumber
    SkinImage(64) = .ReadNumber
    'read playlist2 Select
    PlaylistWin(1).bSelect.H = .ReadNumber
    PlaylistWin(1).bSelect.W = .ReadNumber
    PlaylistWin(1).bSelect.X = .ReadNumber
    PlaylistWin(1).bSelect.y = .ReadNumber
    SkinImage(65) = .ReadNumber
    SkinImage(66) = .ReadNumber
    SkinImage(67) = .ReadNumber
    'read playlist2 list
    PlaylistWin(1).bList.H = .ReadNumber
    PlaylistWin(1).bList.W = .ReadNumber
    PlaylistWin(1).bList.X = .ReadNumber
    PlaylistWin(1).bList.y = .ReadNumber
    SkinImage(68) = .ReadNumber
    SkinImage(69) = .ReadNumber
    SkinImage(70) = .ReadNumber
    'read playlist2 close
    PlaylistWin(1).bClose.H = .ReadNumber
    PlaylistWin(1).bClose.W = .ReadNumber
    PlaylistWin(1).bClose.X = .ReadNumber
    PlaylistWin(1).bClose.y = .ReadNumber
    SkinImage(71) = .ReadNumber
    SkinImage(72) = .ReadNumber
    SkinImage(73) = .ReadNumber
    'read playlist2 size
    PlaylistWin(1).bSize.H = .ReadNumber
    PlaylistWin(1).bSize.W = .ReadNumber
    PlaylistWin(1).bSize.X = .ReadNumber
    PlaylistWin(1).bSize.y = .ReadNumber
    SkinImage(74) = .ReadNumber
    SkinImage(75) = .ReadNumber
    SkinImage(76) = .ReadNumber
    'read main playlist btn
    btnPlaylist.Height = .ReadNumber
    btnPlaylist.Width = .ReadNumber
    btnPlaylist.Left = .ReadNumber
    btnPlaylist.Top = .ReadNumber
    SkinImage(77) = .ReadNumber
    SkinImage(78) = .ReadNumber
    SkinImage(79) = .ReadNumber
    SkinImage(80) = .ReadNumber
    SkinImage(81) = .ReadNumber
    'read main repeat btn
    btnRepeat.Height = .ReadNumber
    btnRepeat.Width = .ReadNumber
    btnRepeat.Left = .ReadNumber
    btnRepeat.Top = .ReadNumber
    SkinImage(82) = .ReadNumber
    SkinImage(83) = .ReadNumber
    SkinImage(84) = .ReadNumber
    SkinImage(85) = .ReadNumber
    SkinImage(86) = .ReadNumber
    'read main shuffle btn
    btnShuffle.Height = .ReadNumber
    btnShuffle.Width = .ReadNumber
    btnShuffle.Left = .ReadNumber
    btnShuffle.Top = .ReadNumber
    SkinImage(87) = .ReadNumber
    SkinImage(88) = .ReadNumber
    SkinImage(89) = .ReadNumber
    SkinImage(90) = .ReadNumber
    SkinImage(91) = .ReadNumber
    'text artist title
    lblArtistTitle.Height = .ReadNumber
    lblArtistTitle.Width = .ReadNumber
    lblArtistTitle.Left = .ReadNumber
    lblArtistTitle.Top = .ReadNumber
    lblArtistTitle.Alignment = .ReadNumber
    SkinColor(0) = .ReadNumber
    SkinColor(1) = .ReadNumber
    cFnt.Bold = CBool(.ReadNumber)
    cFnt.Italic = CBool(.ReadNumber)
    cFnt.Size = .ReadNumber
    cFnt.name = .ReadStr
    'lblArtistTitle.FontBold = CBool(.ReadNumber)
    'lblArtistTitle.FontItalic = CBool(.ReadNumber)
    'x = .ReadNumber
    'lblArtistTitle.Font = .ReadStr
    'lblArtistTitle.FontSize = x
    Set lblArtistTitle.Font = cFnt
    'lblArtistTitle.ResetBg
    'text album
    lblAlbum.Height = .ReadNumber
    lblAlbum.Width = .ReadNumber
    lblAlbum.Left = .ReadNumber
    lblAlbum.Top = .ReadNumber
    lblAlbum.Alignment = .ReadNumber
    SkinColor(2) = .ReadNumber
    SkinColor(3) = .ReadNumber
    lblAlbum.FontBold = CBool(.ReadNumber)
    lblAlbum.FontItalic = CBool(.ReadNumber)
    X = .ReadNumber
    lblAlbum.Font = .ReadStr
    lblAlbum.FontSize = X
    'text genre
    lblGenre.Height = .ReadNumber
    lblGenre.Width = .ReadNumber
    lblGenre.Left = .ReadNumber
    lblGenre.Top = .ReadNumber
    lblGenre.Alignment = .ReadNumber
    SkinColor(4) = .ReadNumber
    SkinColor(5) = .ReadNumber
    lblGenre.FontBold = CBool(.ReadNumber)
    lblGenre.FontItalic = CBool(.ReadNumber)
    X = .ReadNumber
    lblGenre.Font = .ReadStr
    lblGenre.FontSize = X
    'text year
    lblYear.Height = .ReadNumber
    lblYear.Width = .ReadNumber
    lblYear.Left = .ReadNumber
    lblYear.Top = .ReadNumber
    lblYear.Alignment = .ReadNumber
    SkinColor(6) = .ReadNumber
    SkinColor(7) = .ReadNumber
    lblYear.FontBold = CBool(.ReadNumber)
    lblYear.FontItalic = CBool(.ReadNumber)
    X = .ReadNumber
    lblYear.Font = .ReadStr
    lblYear.FontSize = X
    'text comments
    lblComments.Height = .ReadNumber
    lblComments.Width = .ReadNumber
    lblComments.Left = .ReadNumber
    lblComments.Top = .ReadNumber
    lblComments.Alignment = .ReadNumber
    SkinColor(8) = .ReadNumber
    SkinColor(9) = .ReadNumber
    lblComments.FontBold = CBool(.ReadNumber)
    lblComments.FontItalic = CBool(.ReadNumber)
    X = .ReadNumber
    lblComments.Font = .ReadStr
    lblComments.FontSize = X
    'text info
    lblInfo.Height = .ReadNumber
    lblInfo.Width = .ReadNumber
    lblInfo.Left = .ReadNumber
    lblInfo.Top = .ReadNumber
    lblInfo.Alignment = .ReadNumber
    SkinColor(10) = .ReadNumber
    SkinColor(11) = .ReadNumber
    lblInfo.FontBold = CBool(.ReadNumber)
    lblInfo.FontItalic = CBool(.ReadNumber)
    X = .ReadNumber
    lblInfo.Font = .ReadStr
    lblInfo.FontSize = X
    'text time
    lblTime.Height = .ReadNumber
    lblTime.Width = .ReadNumber
    lblTime.Left = .ReadNumber
    lblTime.Top = .ReadNumber
    lblTime.Alignment = .ReadNumber
    SkinColor(12) = .ReadNumber
    SkinColor(13) = .ReadNumber
    lblTime.FontBold = CBool(.ReadNumber)
    lblTime.FontItalic = CBool(.ReadNumber)
    X = .ReadNumber
    lblTime.Font = .ReadStr
    lblTime.FontSize = X
    'text total time
    lblTotalTime.Height = .ReadNumber
    lblTotalTime.Width = .ReadNumber
    lblTotalTime.Left = .ReadNumber
    lblTotalTime.Top = .ReadNumber
    lblTotalTime.Alignment = .ReadNumber
    SkinColor(14) = .ReadNumber
    SkinColor(15) = .ReadNumber
    lblTotalTime.FontBold = CBool(.ReadNumber)
    lblTotalTime.FontItalic = CBool(.ReadNumber)
    X = .ReadNumber
    lblTotalTime.Font = .ReadStr
    lblTotalTime.FontSize = X
    'text plist num
    PlaylistWin(0).num.Pos.H = .ReadNumber
    PlaylistWin(0).num.Pos.W = .ReadNumber
    PlaylistWin(0).num.Pos.X = .ReadNumber
    PlaylistWin(0).num.Pos.y = .ReadNumber
    PlaylistWin(0).num.Align = .ReadNumber
    SkinColor(16) = .ReadNumber
    SkinColor(17) = .ReadNumber
    PlaylistWin(0).num.Bold = CBool(.ReadNumber)
    PlaylistWin(0).num.Italic = CBool(.ReadNumber)
    PlaylistWin(0).num.Size = .ReadNumber
    PlaylistWin(0).num.Font = .ReadStr
    'text plist time
    PlaylistWin(0).Time.Pos.H = .ReadNumber
    PlaylistWin(0).Time.Pos.W = .ReadNumber
    PlaylistWin(0).Time.Pos.X = .ReadNumber
    PlaylistWin(0).Time.Pos.y = .ReadNumber
    PlaylistWin(0).Time.Align = .ReadNumber
    SkinColor(18) = .ReadNumber
    SkinColor(19) = .ReadNumber
    PlaylistWin(0).Time.Bold = CBool(.ReadNumber)
    PlaylistWin(0).Time.Italic = CBool(.ReadNumber)
    PlaylistWin(0).Time.Size = .ReadNumber
    PlaylistWin(0).Time.Font = .ReadStr
    'text plist num
    PlaylistWin(1).num.Pos.H = .ReadNumber
    PlaylistWin(1).num.Pos.W = .ReadNumber
    PlaylistWin(1).num.Pos.X = .ReadNumber
    PlaylistWin(1).num.Pos.y = .ReadNumber
    PlaylistWin(1).num.Align = .ReadNumber
    SkinColor(20) = .ReadNumber
    SkinColor(21) = .ReadNumber
    PlaylistWin(1).num.Bold = CBool(.ReadNumber)
    PlaylistWin(1).num.Italic = CBool(.ReadNumber)
    PlaylistWin(1).num.Size = .ReadNumber
    PlaylistWin(1).num.Font = .ReadStr
    'text plist time
    PlaylistWin(1).Time.Pos.H = .ReadNumber
    PlaylistWin(1).Time.Pos.W = .ReadNumber
    PlaylistWin(1).Time.Pos.X = .ReadNumber
    PlaylistWin(1).Time.Pos.y = .ReadNumber
    PlaylistWin(1).Time.Align = .ReadNumber
    SkinColor(22) = .ReadNumber
    SkinColor(23) = .ReadNumber
    PlaylistWin(1).Time.Bold = CBool(.ReadNumber)
    PlaylistWin(1).Time.Italic = CBool(.ReadNumber)
    PlaylistWin(1).Time.Size = .ReadNumber
    PlaylistWin(1).Time.Font = .ReadStr
    'list playlist
    PlaylistWin(0).List.Pos.H = .ReadNumber
    PlaylistWin(0).List.Pos.W = .ReadNumber
    PlaylistWin(0).List.Pos.X = .ReadNumber
    PlaylistWin(0).List.Pos.y = .ReadNumber
    PlaylistWin(0).List.ColA = .ReadNumber
    PlaylistWin(0).List.ColAT = .ReadNumber
    PlaylistWin(0).List.ColG = .ReadNumber
    PlaylistWin(0).List.ColT = .ReadNumber
    SkinColor(24) = .ReadNumber
    SkinColor(25) = .ReadNumber
    SkinColor(28) = .ReadNumber
    PlaylistWin(0).List.FontSize = .ReadNumber
    PlaylistWin(0).List.Font = .ReadStr
    SkinImage(102) = .ReadNumber
    If SkinImage(102) = 0 Then
      PlaylistWin(0).HasBg = False
    Else
      PlaylistWin(0).HasBg = True
    End If
    'list playlist2
    PlaylistWin(1).List.Pos.H = .ReadNumber
    PlaylistWin(1).List.Pos.W = .ReadNumber
    PlaylistWin(1).List.Pos.X = .ReadNumber
    PlaylistWin(1).List.Pos.y = .ReadNumber
    PlaylistWin(1).List.ColA = .ReadNumber
    PlaylistWin(1).List.ColAT = .ReadNumber
    PlaylistWin(1).List.ColG = .ReadNumber
    PlaylistWin(1).List.ColT = .ReadNumber
    SkinColor(26) = .ReadNumber
    SkinColor(27) = .ReadNumber
    SkinColor(29) = .ReadNumber
    PlaylistWin(1).List.FontSize = .ReadNumber
    PlaylistWin(1).List.Font = .ReadStr
    SkinImage(103) = .ReadNumber
    If SkinImage(103) = 0 Then
      PlaylistWin(1).HasBg = False
    Else
      PlaylistWin(1).HasBg = True
    End If
    'spectrum
    picVis.Height = .ReadNumber
    picVis.Width = .ReadNumber
    picVis.Left = .ReadNumber
    picVis.Top = .ReadNumber
    SkinImage(92) = .ReadNumber
    SavePicture GetImage(SPEC_BAR), tmpFile
    cBar.LoadImg tmpFile
    SkinImage(105) = .ReadNumber
    SavePicture GetImage(SPEC_BEAT), tmpFile
    cBeat.LoadImg tmpFile
    cBeatTemp.LoadImg tmpFile
    SkinImage(93) = .ReadNumber
    SkinImage(94) = .ReadNumber
    SkinImage(95) = .ReadNumber
    SkinImage(96) = .ReadNumber
    SkinImage(104) = .ReadNumber
    SkinImage(97) = .ReadNumber
    y = .ReadNumber
    ReDim SkinPresets(.ReadNumber)
    For X = 1 To UBound(SkinPresets) 'loads all presets + activates the default one
      SkinPresets(X).sName = .ReadStr
      SkinPresets(X).bType = .ReadNumber
      Select Case SkinPresets(X).bType
        Case 1
          With SkinPresets(X).t1
            .bBrushSizeL = File.ReadNumber
            .bBrushSizeR = File.ReadNumber
            .bDetail = File.ReadNumber
            .bFade = File.ReadNumber
            .bFall = File.ReadNumber
            .bPeakCount = File.ReadNumber
            .bPeakDetail = File.ReadNumber
            .bPeaks = File.ReadNumber
            .bSkip = File.ReadNumber
            .bType = File.ReadNumber
            .lColorL = File.ReadNumber
            .lColorPeakL = File.ReadNumber
            .lColorPeakR = File.ReadNumber
            .lColorR = File.ReadNumber
            .lPeakDec = File.ReadNumber
            .lPeakPause = File.ReadNumber
          End With
          If y = X And bDefPreset Then
            ScopeSettings = SkinPresets(X).t1
          
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
            Spectrum = 1
            lCurVisPreset = X
          End If
        Case 2
          With SkinPresets(X).t2
            .bBarSize = File.ReadNumber
            .bBrushSize = File.ReadNumber
            .bDrawStyle = File.ReadNumber
            .bFade = File.ReadNumber
            .bFall = File.ReadNumber
            .bPeakFall = File.ReadNumber
            .bPeaks = File.ReadNumber
            .bType = File.ReadNumber
            .iView = File.ReadNumber
            .lColorDn = File.ReadNumber
            .lColorLine = File.ReadNumber
            .lColorUp = File.ReadNumber
            .lPause = File.ReadNumber
            .lPeakColor = File.ReadNumber
            .lPeakDec = File.ReadNumber
            .lPeakPause = File.ReadNumber
            .nDec = File.ReadNumber
            .nZoom = File.ReadNumber / 100
            .bCorrection = File.ReadNumber
          End With
          If X = y And bDefPreset Then
            SpectrumSettings = SkinPresets(X).t2
            
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
            Spectrum = 2
            lCurVisPreset = X
          End If
        Case 3
          With SkinPresets(X).t3
            .bDrawStyle = File.ReadNumber
            .bFade = File.ReadNumber
            .bFall = File.ReadNumber
            .bType = File.ReadNumber
            .lColorDn = File.ReadNumber
            .lColorUp = File.ReadNumber
            .lPause = File.ReadNumber
            .nDec = File.ReadNumber
            .sFile = ""
            If File.ReadNumber = 200 Then
              .lImage = File.ReadNumber
            End If
          End With
          If X = y And bDefPreset Then
            VolumeSettings = SkinPresets(X).t3
            
            frmMain.DoGrad VolumeSettings.lColorUp, VolumeSettings.lColorDn
            cGradVol.Create 2, frmMain.picVis.ScaleHeight, frmMain.hdc
            cGradVol.BitBltFrom frmMain.p.hdc
            Spectrum = 3
            lCurVisPreset = X
          End If
        Case 4
          With SkinPresets(X).t4
            .bFade = File.ReadNumber
            .bType = File.ReadNumber
            .iDetectHigh = File.ReadNumber
            .iDetectLow = File.ReadNumber
            .nMin = File.ReadNumber / 1000
            .nMulti = File.ReadNumber / 1000
            .nRotMin = File.ReadNumber / 1000
            .nRotMove = File.ReadNumber / 1000
            .nRotSpeed = File.ReadNumber / 1000
            .sFile = ""
            If File.ReadNumber = 200 Then
              .lImage = File.ReadNumber
            End If
          End With
          If X = y And bDefPreset Then
            BeatSettings = SkinPresets(X).t4
            Spectrum = 4
            lCurVisPreset = X
          End If
      End Select
    Next
    
    Set picVis.Picture = GetImage(SPEC_BGOFF)
    
    cBeat.InitSA
    cBeatTemp.InitSA
    
    ReDim VolumeHistory((picVis.Width \ 2) - 2)
    ReDim ScopeHistoryL(UBound(VolumeHistory))
    ReDim ScopeHistoryR(UBound(VolumeHistory))
    'create pens, brushes & gradient images for visualization
    hPenLeft = CreatePen(0, ScopeSettings.bBrushSizeL, ScopeSettings.lColorL)
    hPenRight = CreatePen(0, ScopeSettings.bBrushSizeR, ScopeSettings.lColorR)
    hPenPeakLeft = CreatePen(0, 1, ScopeSettings.lColorPeakL)
    hPenPeakRight = CreatePen(0, 1, ScopeSettings.lColorPeakR)
    hPenPeakSpec = CreatePen(0, 1, SpectrumSettings.lPeakColor)
    hPenLineSpec = CreatePen(0, SpectrumSettings.bBrushSize, SpectrumSettings.lColorLine)
    hBrushSolidLeft = CreateSolidBrush(ScopeSettings.lColorL)
    hBrushSolidRight = CreateSolidBrush(ScopeSettings.lColorR)
    p.Height = picVis.Height
    DoGrad SpectrumSettings.lColorUp, SpectrumSettings.lColorDn
    cGradBar.Create 1, picVis.ScaleHeight, Me.hdc
    cGradBar.BitBltFrom p.hdc
    DoGrad VolumeSettings.lColorUp, VolumeSettings.lColorDn
    cGradVol.Create 2, picVis.ScaleHeight, Me.hdc
    cGradVol.BitBltFrom p.hdc
    
    'Read window regions
    'Main window first
    ReDim R1(.ReadNumber)
    If UBound(R1) > 0 Then
      For X = 1 To UBound(R1)
        R1(X).Left = .ReadNumber
        R1(X).Top = .ReadNumber
        R1(X).Right = .ReadNumber
        R1(X).Bottom = .ReadNumber
      Next
    End If
    
    'playlist big window
    ReDim R2(.ReadNumber)
    If UBound(R2) > 0 Then
      For X = 1 To UBound(R2)
        R2(X).Left = .ReadNumber
        R2(X).Top = .ReadNumber
        R2(X).Right = .ReadNumber
        R2(X).Bottom = .ReadNumber
      Next
    End If

    'playlist small window
    ReDim R3(.ReadNumber)
    If UBound(R3) > 0 Then
      For X = 1 To UBound(R3)
        R3(X).Left = .ReadNumber
        R3(X).Top = .ReadNumber
        R3(X).Right = .ReadNumber
        R3(X).Bottom = .ReadNumber
      Next
    End If

    If FileExists(tmpFile) Then Kill tmpFile
    
    'EVERYTHING LOADED & DONE!
    '... but now we have to update & draw images at the windows...
    LockWindowUpdate 0
    
    cLog.Log "DONE. (" & cLog.GetTimer & " ms)", 3
    
    DrawEverything
  
  End With
  
  Exit Sub
SkinError:
  If Skin <> DEFAULTSKIN Then 'error on non-default skin
    If Err.Number = 30335 Then
      MsgBox "The skin you tried to load was made for an older version of Simple Amp. The default skin will be loaded.", vbCritical, "Skin Error"
    Else
      MsgBox "There was an error while loading the selected skin. The default skin will be loaded.", vbCritical, "Skin Error"
    End If
    CurrentSkin = DEFAULTSKIN
    LoadSkin CurrentSkin
  Else
    If Err.Number = 30335 Then 'error on default skin
      MsgBox "The default skin was made for an older version of Simple Amp. Please download the newest version of Simple Amp. If you have other skins, try changing to an other skin via the ini-file. ", vbCritical, "Skin Error"
    Else
      MsgBox "There was an error while loading the default skin. Simple Amp cannot be started. If you have other skins, try changing to an other skin via the ini-file.", vbCritical, "Skin Error"
    End If
    UnloadAll
  End If
End Sub

Public Sub DrawEverything()
  On Error Resume Next
  'this will draw everything (after load of skin)
  'first, begin with main window

  LockWindowUpdate Me.hwnd
  
  Me.Picture = GetImage(SI_MAIN_BG)
  If MainTrans Then
    Dim z As Long, X As Long, y As Long
    p.Picture = GetImage(SI_MAIN_BG)
    z = CreateRectRgn(0, 0, p.ScaleWidth, p.ScaleHeight)
    For X = 1 To UBound(R1) 'creates region by combining rectangles in memory
      y = CreateRectRgn(R1(X).Left, R1(X).Top, R1(X).Right, R1(X).Bottom)
      CombineRgn z, z, y, RGN_DIFF
    Next
    p.Picture = Nothing
    SetWindowRgn Me.hwnd, z, True
  Else
    Call SetWindowRgn(Me.hwnd, 0, True)
  End If
  If CurMono Then
    imgStereoMono.Picture = GetImage(SI_MAIN_MONO)
  Else
    imgStereoMono.Picture = GetImage(SI_MAIN_STEREO)
  End If
  Set Position.Bar = GetImage(SCR_MAIN_POS_BAR)
  Set Position.BarOver = GetImage(SCR_MAIN_POS_BARDRAG)
  Set Position.ScrollBefore = GetImage(SCR_MAIN_POS_BEFORE)
  Set Position.ScrollAfter = GetImage(SCR_MAIN_POS_AFTER)
  Set Volume.Bar = GetImage(SCR_MAIN_VOL_BAR)
  Set Volume.BarOver = GetImage(SCR_MAIN_VOL_BARDRAG)
  Set Volume.ScrollBefore = GetImage(SCR_MAIN_VOL_BEFORE)
  Set Volume.ScrollAfter = GetImage(SCR_MAIN_VOL_AFTER)
  Set btnPrev.gfxUp = GetImage(BTN_MAIN_PREV_UP)
  Set btnPrev.gfxDown = GetImage(BTN_MAIN_PREV_DOWN)
  Set btnPrev.gfxUpOver = GetImage(BTN_MAIN_PREV_UPM)
  Set btnNext.gfxUp = GetImage(BTN_MAIN_NEXT_UP)
  Set btnNext.gfxDown = GetImage(BTN_MAIN_NEXT_DOWN)
  Set btnNext.gfxUpOver = GetImage(BTN_MAIN_NEXT_UPM)
  Set btnStop.gfxUp = GetImage(BTN_MAIN_STOP_UP)
  Set btnStop.gfxDown = GetImage(BTN_MAIN_STOP_DOWN)
  Set btnStop.gfxUpOver = GetImage(BTN_MAIN_STOP_UPM)
  Set btnClose.gfxUp = GetImage(BTN_MAIN_CLOSE_UP)
  Set btnClose.gfxDown = GetImage(BTN_MAIN_CLOSE_DOWN)
  Set btnClose.gfxUpOver = GetImage(BTN_MAIN_CLOSE_UPM)
  Set btnMinimize.gfxUp = GetImage(BTN_MAIN_MIN_UP)
  Set btnMinimize.gfxDown = GetImage(BTN_MAIN_MIN_DOWN)
  Set btnMinimize.gfxUpOver = GetImage(BTN_MAIN_MIN_UPM)
  If Sound.StreamIsPlaying Or Sound.MusicIsPlaying Then
    Set btnPlayPause.gfxUp = GetImage(BTN_MAIN_PAUSE_UP)
    Set btnPlayPause.gfxDown = GetImage(BTN_MAIN_PAUSE_DOWN)
    Set btnPlayPause.gfxUpOver = GetImage(BTN_MAIN_PAUSE_UPM)
  Else
    Set btnPlayPause.gfxUp = GetImage(BTN_MAIN_PLAY_UP)
    Set btnPlayPause.gfxDown = GetImage(BTN_MAIN_PLAY_DOWN)
    Set btnPlayPause.gfxUpOver = GetImage(BTN_MAIN_PLAY_UPM)
  End If
  Set btnRepeat.gfxDown = GetImage(XBTN_MAIN_RPT_DOWN)
  Set btnRepeat.gfxUp = GetImage(XBTN_MAIN_RPT_ONUP)
  Set btnRepeat.gfxUpOver = GetImage(XBTN_MAIN_RPT_ONUPM)
  Set btnRepeat.gfxOffUp = GetImage(XBTN_MAIN_RPT_OFFUP)
  Set btnRepeat.gfxOffUpOver = GetImage(XBTN_MAIN_RPT_OFFUPM)
  RepeatOnOff
  Set btnShuffle.gfxDown = GetImage(XBTN_MAIN_SHFL_DOWN)
  Set btnShuffle.gfxUp = GetImage(XBTN_MAIN_SHFL_ONUP)
  Set btnShuffle.gfxUpOver = GetImage(XBTN_MAIN_SHFL_ONUPM)
  Set btnShuffle.gfxOffUp = GetImage(XBTN_MAIN_SHFL_OFFUP)
  Set btnShuffle.gfxOffUpOver = GetImage(XBTN_MAIN_SHFL_OFFUPM)
  ShuffleOnOff
  Set btnPlaylist.gfxDown = GetImage(XBTN_MAIN_PLIST_DOWN)
  Set btnPlaylist.gfxUp = GetImage(XBTN_MAIN_PLIST_ONUP)
  Set btnPlaylist.gfxUpOver = GetImage(XBTN_MAIN_PLIST_ONUPM)
  Set btnPlaylist.gfxOffUp = GetImage(XBTN_MAIN_PLIST_OFFUP)
  Set btnPlaylist.gfxOffUpOver = GetImage(XBTN_MAIN_PLIST_OFFUPM)
  btnPlaylist.Value = Settings.PlaylistOn
  
  UpdateSpectrum 'draw scope
  UpdateMainColor 'update text colors
  LockWindowUpdate 0
  UpdatePlaylist
  
  lblArtistTitle.AutoScroll = True
  
End Sub

Public Sub UpdateMainColor()
  On Error Resume Next
  'this sub updates all of the main windows text colors
  If Sound.StreamIsLoaded Or Sound.MusicIsLoaded Then
    lblArtistTitle.ForeColor = GetColor(TXT_MAIN_ARTISTTITLE_EN)
    lblAlbum.ForeColor = GetColor(TXT_MAIN_ALBUM_EN)
    lblGenre.ForeColor = GetColor(TXT_MAIN_GENRE_EN)
    lblYear.ForeColor = GetColor(TXT_MAIN_YEAR_EN)
    lblComments.ForeColor = GetColor(TXT_MAIN_COM_EN)
    lblInfo.ForeColor = GetColor(TXT_MAIN_INF_EN)
    lblTime.ForeColor = GetColor(TXT_MAIN_TIME_EN)
    lblTotalTime.ForeColor = GetColor(TXT_MAIN_TOTTIME_EN)
  Else
    lblArtistTitle.ForeColor = GetColor(TXT_MAIN_ARTISTTITLE_DIS)
    lblAlbum.ForeColor = GetColor(TXT_MAIN_ALBUM_DIS)
    lblGenre.ForeColor = GetColor(TXT_MAIN_GENRE_DIS)
    lblYear.ForeColor = GetColor(TXT_MAIN_YEAR_DIS)
    lblComments.ForeColor = GetColor(TXT_MAIN_COM_DIS)
    lblInfo.ForeColor = GetColor(TXT_MAIN_INF_DIS)
    lblTime.ForeColor = GetColor(TXT_MAIN_TIME_DIS)
    lblTotalTime.ForeColor = GetColor(TXT_MAIN_TOTTIME_DIS)
  End If
End Sub

Public Sub UpdatePlaylist()
  On Error Resume Next
  Dim X As Long
  'this sub totally redraws & updates the playlist window
  X = Abs(Settings.PlaylistSmall)
  
  With frmPlaylist
  
    'LockWindowUpdate .hwnd
    
    'first, update the size & location of every control
    .imgColumns.Width = PlaylistWin(X).iColumn.W
    .imgColumns.Height = PlaylistWin(X).iColumn.H
    .imgColumns.Top = PlaylistWin(X).iColumn.y
    .imgColumns.Left = PlaylistWin(X).iColumn.X
    .Scroll.Width = PlaylistWin(X).sScroll.W
    .Scroll.Height = PlaylistWin(X).sScroll.H
    .Scroll.Top = PlaylistWin(X).sScroll.y
    .Scroll.Left = PlaylistWin(X).sScroll.X
    .btnAdd.Width = PlaylistWin(X).bAdd.W
    .btnAdd.Height = PlaylistWin(X).bAdd.H
    .btnAdd.Top = PlaylistWin(X).bAdd.y
    .btnAdd.Left = PlaylistWin(X).bAdd.X
    .btnRem.Width = PlaylistWin(X).bRemove.W
    .btnRem.Height = PlaylistWin(X).bRemove.H
    .btnRem.Top = PlaylistWin(X).bRemove.y
    .btnRem.Left = PlaylistWin(X).bRemove.X
    .btnSelect.Width = PlaylistWin(X).bSelect.W
    .btnSelect.Height = PlaylistWin(X).bSelect.H
    .btnSelect.Top = PlaylistWin(X).bSelect.y
    .btnSelect.Left = PlaylistWin(X).bSelect.X
    .btnList.Width = PlaylistWin(X).bList.W
    .btnList.Height = PlaylistWin(X).bList.H
    .btnList.Top = PlaylistWin(X).bList.y
    .btnList.Left = PlaylistWin(X).bList.X
    .btnClose.Width = PlaylistWin(X).bClose.W
    .btnClose.Height = PlaylistWin(X).bClose.H
    .btnClose.Top = PlaylistWin(X).bClose.y
    .btnClose.Left = PlaylistWin(X).bClose.X
    .btnSize.Width = PlaylistWin(X).bSize.W
    .btnSize.Height = PlaylistWin(X).bSize.H
    .btnSize.Top = PlaylistWin(X).bSize.y
    .btnSize.Left = PlaylistWin(X).bSize.X
    .lblTotalNum.Alignment = PlaylistWin(X).num.Align
    .lblTotalNum.FontBold = PlaylistWin(X).num.Bold
    .lblTotalNum.Font = PlaylistWin(X).num.Font
    .lblTotalNum.FontItalic = PlaylistWin(X).num.Italic
    .lblTotalNum.FontSize = PlaylistWin(X).num.Size
    .lblTotalNum.Height = PlaylistWin(X).num.Pos.H
    .lblTotalNum.Width = PlaylistWin(X).num.Pos.W
    .lblTotalNum.Left = PlaylistWin(X).num.Pos.X
    .lblTotalNum.Top = PlaylistWin(X).num.Pos.y
    .lblTotalTime.Alignment = PlaylistWin(X).Time.Align
    .lblTotalTime.FontBold = PlaylistWin(X).Time.Bold
    .lblTotalTime.Font = PlaylistWin(X).Time.Font
    .lblTotalTime.FontItalic = PlaylistWin(X).Time.Italic
    .lblTotalTime.FontSize = PlaylistWin(X).Time.Size
    .lblTotalTime.Height = PlaylistWin(X).Time.Pos.H
    .lblTotalTime.Width = PlaylistWin(X).Time.Pos.W
    .lblTotalTime.Left = PlaylistWin(X).Time.Pos.X
    .lblTotalTime.Top = PlaylistWin(X).Time.Pos.y
    .List.Height = PlaylistWin(X).List.Pos.H
    .List.Width = PlaylistWin(X).List.Pos.W
    .List.Left = PlaylistWin(X).List.Pos.X
    .List.Top = PlaylistWin(X).List.Pos.y
    .List.Font = PlaylistWin(X).List.Font
    .List.FontSize = PlaylistWin(X).List.FontSize
    .List.SetColumnWidth 0, PlaylistWin(X).List.ColAT
    .List.SetColumnWidth 1, PlaylistWin(X).List.ColA
    .List.SetColumnWidth 2, PlaylistWin(X).List.ColG
    .List.SetColumnWidth 3, PlaylistWin(X).List.ColT
    
    Dim z As Long, x2 As Long, y As Long
    
    'now, update images! fun fun fun!
    If Settings.PlaylistSmall Then
      .Picture = GetImage(SI_PLIST2_BG)
      If PlaylistWin(1).Trans Then
        p.Picture = GetImage(SI_PLIST2_BG)
        z = CreateRectRgn(0, 0, p.ScaleWidth, p.ScaleHeight)
        For x2 = 1 To UBound(R3)
          y = CreateRectRgn(R3(x2).Left, R3(x2).Top, R3(x2).Right, R3(x2).Bottom)
          CombineRgn z, z, y, RGN_DIFF
        Next
        p.Picture = Nothing
        SetWindowRgn .hwnd, z, True
      Else
        Call SetWindowRgn(.hwnd, 0, True)
      End If
      .imgColumns.Picture = GetImage(SI_PLIST2_COLUMNS)
      Set .Scroll.ScrollBefore = GetImage(SCR_PLIST2_SCRL_BEFORE)
      Set .Scroll.ScrollAfter = GetImage(SCR_PLIST2_SCRL_AFTER)
      Set .Scroll.Bar = GetImage(SCR_PLIST2_SCRL_BAR)
      Set .Scroll.BarOver = GetImage(SCR_PLIST2_SCRL_BARDRAG)
      Set .btnAdd.gfxUp = GetImage(BTN_PLIST2_ADD_UP)
      Set .btnAdd.gfxDown = GetImage(BTN_PLIST2_ADD_DOWN)
      Set .btnAdd.gfxUpOver = GetImage(BTN_PLIST2_ADD_UPM)
      Set .btnRem.gfxUp = GetImage(BTN_PLIST2_REM_UP)
      Set .btnRem.gfxDown = GetImage(BTN_PLIST2_REM_DOWN)
      Set .btnRem.gfxUpOver = GetImage(BTN_PLIST2_REM_UPM)
      Set .btnSelect.gfxUp = GetImage(BTN_PLIST2_SEL_UP)
      Set .btnSelect.gfxDown = GetImage(BTN_PLIST2_SEL_DOWN)
      Set .btnSelect.gfxUpOver = GetImage(BTN_PLIST2_SEL_UPM)
      Set .btnList.gfxUp = GetImage(BTN_PLIST2_LST_UP)
      Set .btnList.gfxDown = GetImage(BTN_PLIST2_LST_DOWN)
      Set .btnList.gfxUpOver = GetImage(BTN_PLIST2_LST_UPM)
      Set .btnClose.gfxUp = GetImage(BTN_PLIST2_CLOSE_UP)
      Set .btnClose.gfxDown = GetImage(BTN_PLIST2_CLOSE_DOWN)
      Set .btnClose.gfxUpOver = GetImage(BTN_PLIST2_CLOSE_UPM)
      Set .btnSize.gfxUp = GetImage(BTN_PLIST2_SIZE_UP)
      Set .btnSize.gfxDown = GetImage(BTN_PLIST2_SIZE_DOWN)
      Set .btnSize.gfxUpOver = GetImage(BTN_PLIST2_SIZE_UPM)
      .List.ForeColor = GetColor(LST_PLIST2_FONT)
      .List.BackColor = GetColor(LST_PLIST2_BG)
      .List.SelectColor = GetColor(LST_PLIST2_SELECT)
      If PlaylistWin(1).HasBg Then
        Set .List.BgPicture = GetImage(PLIST2_LST_BG)
      Else
        Set .List.BgPicture = Nothing
      End If
    Else
      .Picture = GetImage(SI_PLIST_BG)
      If PlaylistWin(0).Trans Then
        p.Picture = GetImage(SI_PLIST_BG)
        z = CreateRectRgn(0, 0, p.ScaleWidth, p.ScaleHeight)
        For x2 = 1 To UBound(R2)
          y = CreateRectRgn(R2(x2).Left, R2(x2).Top, R2(x2).Right, R2(x2).Bottom)
          CombineRgn z, z, y, RGN_DIFF
        Next
        p.Picture = Nothing
        SetWindowRgn .hwnd, z, True
      Else
        Call SetWindowRgn(.hwnd, 0, True)
      End If
      .imgColumns.Picture = GetImage(SI_PLIST_COLUMNS)
      Set .Scroll.ScrollBefore = GetImage(SCR_PLIST_SCRL_BEFORE)
      Set .Scroll.ScrollAfter = GetImage(SCR_PLIST_SCRL_AFTER)
      Set .Scroll.Bar = GetImage(SCR_PLIST_SCRL_BAR)
      Set .Scroll.BarOver = GetImage(SCR_PLIST_SCRL_BARDRAG)
      Set .btnAdd.gfxUp = GetImage(BTN_PLIST_ADD_UP)
      Set .btnAdd.gfxDown = GetImage(BTN_PLIST_ADD_DOWN)
      Set .btnAdd.gfxUpOver = GetImage(BTN_PLIST_ADD_UPM)
      Set .btnRem.gfxUp = GetImage(BTN_PLIST_REM_UP)
      Set .btnRem.gfxDown = GetImage(BTN_PLIST_REM_DOWN)
      Set .btnRem.gfxUpOver = GetImage(BTN_PLIST_REM_UPM)
      Set .btnSelect.gfxUp = GetImage(BTN_PLIST_SEL_UP)
      Set .btnSelect.gfxDown = GetImage(BTN_PLIST_SEL_DOWN)
      Set .btnSelect.gfxUpOver = GetImage(BTN_PLIST_SEL_UPM)
      Set .btnList.gfxUp = GetImage(BTN_PLIST_LST_UP)
      Set .btnList.gfxDown = GetImage(BTN_PLIST_LST_DOWN)
      Set .btnList.gfxUpOver = GetImage(BTN_PLIST_LST_UPM)
      Set .btnClose.gfxUp = GetImage(BTN_PLIST_CLOSE_UP)
      Set .btnClose.gfxDown = GetImage(BTN_PLIST_CLOSE_DOWN)
      Set .btnClose.gfxUpOver = GetImage(BTN_PLIST_CLOSE_UPM)
      Set .btnSize.gfxUp = GetImage(BTN_PLIST_SIZE_UP)
      Set .btnSize.gfxDown = GetImage(BTN_PLIST_SIZE_DOWN)
      Set .btnSize.gfxUpOver = GetImage(BTN_PLIST_SIZE_UPM)
      .List.ForeColor = GetColor(LST_PLIST_FONT)
      .List.BackColor = GetColor(LST_PLIST_BG)
      .List.SelectColor = GetColor(LST_PLIST_SELECT)
      If PlaylistWin(0).HasBg Then
        Set .List.BgPicture = GetImage(PLIST_LST_BG)
      Else
        Set .List.BgPicture = Nothing
      End If
    End If
  
    UpdatePlaylistColor
    'LockWindowUpdate 0
    
    .List.Refresh
    
    .Width = PlaylistWin(X).Self.W
    .Height = PlaylistWin(X).Self.H
  
  End With
  
End Sub

Public Sub UpdatePlaylistColor()
  On Error Resume Next
  'this sub updates all of the main windows text colors
  If UBound(Playlist) > 0 Then
    frmPlaylist.lblTotalNum.ForeColor = GetColor(TXT_PLIST_NUM_EN)
    frmPlaylist.lblTotalTime.ForeColor = GetColor(TXT_PLIST_TIME_EN)
  Else
    frmPlaylist.lblTotalNum.ForeColor = GetColor(TXT_PLIST_NUM_DIS)
    frmPlaylist.lblTotalTime.ForeColor = GetColor(TXT_PLIST_TIME_DIS)
  End If
End Sub

Public Function GetShuffle() As Long
  'This function returns the item index of the next song to be played
  'according to ShuffleNum.
  On Error Resume Next
  Dim X As Long
  
  For X = 1 To UBound(Playlist)
    If Playlist(X).lShuffleIndex = ShuffleNum Then
      GetShuffle = X
      Exit For
    End If
  Next
  
  'If there was none, get the ones with 0 as lShuffleIndex and choose
  'randomly between them
  If GetShuffle = 0 Then
    Dim Shuf As New Collection
        
    For X = 1 To UBound(Playlist)
      If Playlist(X).lShuffleIndex = 0 Then
        Shuf.Add X
      End If
    Next

    If Shuf.Count > 0 Then
      If Shuf.Count > 1 Then
        X = CInt(Rnd * (Shuf.Count - 1)) + 1
      Else
        X = 1
      End If
      'Playlist(Shuf(x)).lShuffleIndex = ShuffleNum
      'the commented out is set in the Play() sub instead
      GetShuffle = Shuf(X)
    End If
    
  End If

End Function

Public Sub UpdatePlayPause(ByVal bPause As Boolean)
  On Error Resume Next
  With btnPlayPause
    If bPause Then
      Set .gfxUp = GetImage(BTN_MAIN_PAUSE_UP)
      Set .gfxDown = GetImage(BTN_MAIN_PAUSE_DOWN)
      Set .gfxUpOver = GetImage(BTN_MAIN_PAUSE_UPM)
    Else
      Set .gfxUp = GetImage(BTN_MAIN_PLAY_UP)
      Set .gfxDown = GetImage(BTN_MAIN_PLAY_DOWN)
      Set .gfxUpOver = GetImage(BTN_MAIN_PLAY_UPM)
    End If
  End With
End Sub

Private Sub Volume_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_KeyDown KeyCode, Shift
End Sub

Private Sub Volume_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Public Sub visScope()
  'Draw oscilliscope
  On Error Resume Next

  If Not Sound.StreamIsPlaying And Not Sound.MusicIsPlaying Then
    'If no sound is playing, just fade out/clear the screen
    
    If ScopeSettings.bFade > 0 Then
      cBack.Fade ScopeSettings.bFade
    Else
      cBack.BitBltFrom cBackOrig.hdc
    End If
    
  Else
    'Sound IS playing, so draw the oscilliscope
    
    Dim point As POINTAPI 'this point is just for MoveToEx, but never used
    Dim mRect As RECT
    Dim X As Integer, H As Integer, W As Integer
    
    'clear backbuffer before drawing
    If ScopeSettings.bFade > 0 Then
      cBack.Fade ScopeSettings.bFade
    Else
      cBack.BitBltFrom cBackOrig.hdc
    End If
    
    If FSOUND_GetMixer = FSOUND_MIXER_QUALITY_FPU Then
      'If mixer is CPU/FPU, we have to work with Single values
      
      Call GetScopeFPU(ScopeBufferFPU) 'Get sound data
      'data is placed in the buffer switching between left & right channel
      '(value 1 left, value 1 right, value 2 left, value 2 right ...)
      
      'Update History (for history mode vis)
      For X = 0 To UBound(ScopeHistoryL) - 1
        ScopeHistoryL(X) = ScopeHistoryL(X + 1)
        ScopeHistoryR(X) = ScopeHistoryR(X + 1)
      Next
      'Actually I don't really know which comes first, Left or Right but who cares, right
      ScopeHistoryL(UBound(ScopeHistoryL)) = ScopeBufferFPU(0)
      ScopeHistoryR(UBound(ScopeHistoryR)) = ScopeBufferFPU(1)
      
      'Handle peaks
      If ScopeSettings.bPeaks > 0 Then
        For X = 0 To UBound(ScopeBufferFPU)
          'upper peaks first!
          With ScopeUPeaks(X)
            If ScopeSettings.bFall = 1 Then
              If ScopeSettings.lPeakDec = 0 Then ScopeSettings.lPeakDec = 20
            Else
              If ScopeSettings.lPeakDec = 0 Then ScopeSettings.lPeakDec = 75
              .Dec = ScopeSettings.lPeakDec * 10
            End If
            If .Value - .Dec > ScopeBufferFPU(X) Then
              .Value = ScopeBufferFPU(X)
              .Time = timeGetTime
              .Pause = .Time
              .Dec = 0
            ElseIf timeGetTime - .Pause > ScopeSettings.lPeakPause Then
              If timeGetTime - .Time > 10 Then
                .Time = timeGetTime
                If ScopeSettings.bFall = 1 Then
                  .Dec = .Dec + ScopeSettings.lPeakDec
                End If
                .Value = .Value + .Dec
                If .Value > 0 Then
                  .Value = 0
                  .Dec = 0
                End If
              End If
            End If
          End With
          'lower peaks
          With ScopeLPeaks(X)
            If ScopeSettings.bFall = 1 Then
              If ScopeSettings.lPeakDec = 0 Then ScopeSettings.lPeakDec = 20
            Else
              If ScopeSettings.lPeakDec = 0 Then ScopeSettings.lPeakDec = 75
              .Dec = ScopeSettings.lPeakDec * 10
            End If
            If .Value + .Dec < ScopeBufferFPU(X) Then
              .Value = ScopeBufferFPU(X)
              .Time = timeGetTime
              .Pause = .Time
              .Dec = 0
            ElseIf timeGetTime - .Pause > ScopeSettings.lPeakPause Then
              If timeGetTime - .Time > 10 Then
                .Time = timeGetTime
                If ScopeSettings.bFall = 1 Then
                  .Dec = .Dec + ScopeSettings.lPeakDec
                End If
                .Value = .Value - .Dec
                If .Value < 0 Then
                  .Value = 0
                  .Dec = 0
                End If
              End If
            End If
          End With
        Next
      End If

      'Draw the scope
      If ScopeSettings.bType = 0 Then 'Dot Scope

        If ScopeSettings.bDetail = 0 Then 'left channel first
          For X = 0 To UBound(ScopeBufferFPU) Step ScopeSettings.bSkip
            H = ((ScopeBufferFPU(X) + 32768) / 65535 * cBack.lHeight)
            W = (cBack.lWidth * X * 2) / UBound(ScopeBufferFPU)
            SetPixelV cBack.hdc, W, H, ScopeSettings.lColorL
          Next
        End If

        For X = 1 To UBound(ScopeBufferFPU) Step ScopeSettings.bSkip 'right channel
          H = ((ScopeBufferFPU(X) + 32768) / 65535 * cBack.lHeight)
          W = (cBack.lWidth * X * 2) / UBound(ScopeBufferFPU)
          SetPixelV cBack.hdc, W, H, ScopeSettings.lColorR
        Next
               
      ElseIf ScopeSettings.bType = 1 Then 'Line Scope

        If ScopeSettings.bDetail = 0 Then 'left channel
          SelectObject cBack.hdc, hPenLeft 'select the correct pen for drawing
          MoveToEx cBack.hdc, 0, ((ScopeBufferFPU(0) + 32768) / 65535 * cBack.lHeight), point
          For X = 0 To UBound(ScopeBufferFPU) Step ScopeSettings.bSkip
            H = ((ScopeBufferFPU(X) + 32768) / 65535 * cBack.lHeight)
            W = (cBack.lWidth * X * 2) / UBound(ScopeBufferFPU)
            LineTo cBack.hdc, W, H
          Next
        End If

        'right channel
        SelectObject cBack.hdc, hPenRight
        MoveToEx cBack.hdc, 0, ((ScopeBufferFPU(1) + 32768) / 65535 * cBack.lHeight), point
        For X = 1 To UBound(ScopeBufferFPU) Step ScopeSettings.bSkip
          H = ((ScopeBufferFPU(X) + 32768) / 65535 * cBack.lHeight)
          W = (cBack.lWidth * X * 2) / UBound(ScopeBufferFPU)
          LineTo cBack.hdc, W, H
        Next

      ElseIf ScopeSettings.bType = 2 Then 'Solid Scope
        
        If ScopeSettings.bDetail = 0 Then 'left
          For X = 0 To UBound(ScopeBufferFPU) Step ScopeSettings.bSkip
            mRect.Bottom = cBack.lHeight / 2
            mRect.Top = ((ScopeBufferFPU(X) + 32768) / 65535 * cBack.lHeight)
            mRect.Left = (cBack.lWidth * X * 2) / UBound(ScopeBufferFPU)
            mRect.Right = (cBack.lWidth * (X + 2) * 2) / UBound(ScopeBufferFPU)
            FillRect cBack.hdc, mRect, hBrushSolidLeft
          Next
        End If
        
        'right channel
        'MoveToEx cBack.hdc, 0, (cBack.lHeight / 2), point
        'LineTo cBack.hdc, cBack.lWidth, (cBack.lHeight / 2)
        For X = 1 To UBound(ScopeBufferFPU) Step ScopeSettings.bSkip
          mRect.Bottom = cBack.lHeight / 2
          mRect.Top = ((ScopeBufferFPU(X) + 32768) / 65535 * cBack.lHeight)
          mRect.Left = (cBack.lWidth * X * 2) / UBound(ScopeBufferFPU)
          mRect.Right = (cBack.lWidth * (X + 2) * 2) / UBound(ScopeBufferFPU)
          FillRect cBack.hdc, mRect, hBrushSolidRight
        Next

      Else 'History Scope

        If ScopeSettings.bDetail = 0 Then 'left
          SelectObject cBack.hdc, hPenLeft
          MoveToEx cBack.hdc, 0, ((ScopeHistoryL(0) + 32768) / 65535 * cBack.lHeight), point
          For X = 0 To UBound(ScopeHistoryL)
            LineTo cBack.hdc, X * 2, ((ScopeHistoryL(X) + 32768) / 65535 * cBack.lHeight)
          Next
        End If

        'right
        SelectObject cBack.hdc, hPenRight
        MoveToEx cBack.hdc, 0, ((ScopeHistoryR(0) + 32768) / 65535 * cBack.lHeight), point
        For X = 0 To UBound(ScopeHistoryR)
          LineTo cBack.hdc, X * 2, ((ScopeHistoryR(X) + 32768) / 65535 * cBack.lHeight)
        Next

      End If
  
      'Draw Peaks!
      If ScopeSettings.bType < 3 Then
        If ScopeSettings.bPeaks = 1 Then  'dots
  
          'stereo channel
          If ScopeSettings.bDetail = 0 And ScopeSettings.bPeakDetail = 0 Then
            For X = 0 To UBound(ScopeUPeaks) Step 2
              W = (cBack.lWidth * X * 2) / UBound(ScopeUPeaks)
              'Upper or both
              If ScopeSettings.bPeakCount = 0 Or ScopeSettings.bPeakCount = 2 Then
                H = ((ScopeUPeaks(X).Value + 32768) / 65535 * cBack.lHeight)
                SetPixelV cBack.hdc, W, H, ScopeSettings.lColorPeakL
              End If
              'lower or both
              If ScopeSettings.bPeakCount > 0 Then
                H = ((ScopeLPeaks(X).Value + 32768) / 65535 * cBack.lHeight)
                SetPixelV cBack.hdc, W, H, ScopeSettings.lColorPeakL
              End If
            Next
          End If
  
          'mono channel
          For X = 1 To UBound(ScopeUPeaks) Step 2
            W = (cBack.lWidth * X * 2) / UBound(ScopeUPeaks)
            'Upper or both
            If ScopeSettings.bPeakCount = 0 Or ScopeSettings.bPeakCount = 2 Then
              H = ((ScopeUPeaks(X).Value + 32768) / 65535 * cBack.lHeight)
              SetPixelV cBack.hdc, W, H, ScopeSettings.lColorPeakR
            End If
            'lower or both
            If ScopeSettings.bPeakCount > 0 Then
              H = ((ScopeLPeaks(X).Value + 32768) / 65535 * cBack.lHeight)
              SetPixelV cBack.hdc, W, H, ScopeSettings.lColorPeakR
            End If
          Next
  
        ElseIf ScopeSettings.bPeaks = 2 Then  'lines
  
          'Stereo channel
          If ScopeSettings.bDetail = 0 And ScopeSettings.bPeakDetail = 0 Then
            SelectObject cBack.hdc, hPenPeakLeft
            'upper or both
            If ScopeSettings.bPeakCount = 0 Or ScopeSettings.bPeakCount = 2 Then
              MoveToEx cBack.hdc, 0, ((ScopeUPeaks(0).Value + 32768) / 65535 * cBack.lHeight), point
              For X = 0 To UBound(ScopeUPeaks) Step 2
                W = (cBack.lWidth * X * 2) / UBound(ScopeUPeaks)
                H = ((ScopeUPeaks(X).Value + 32768) / 65535 * cBack.lHeight)
                LineTo cBack.hdc, W, H
              Next
            End If
            'lower or both
            If ScopeSettings.bPeakCount > 0 Then
              MoveToEx cBack.hdc, 0, ((ScopeLPeaks(0).Value + 32768) / 65535 * cBack.lHeight), point
              For X = 0 To UBound(ScopeLPeaks) Step 2
                W = (cBack.lWidth * X * 2) / UBound(ScopeLPeaks)
                H = ((ScopeLPeaks(X).Value + 32768) / 65535 * cBack.lHeight)
                LineTo cBack.hdc, W, H
              Next
            End If
          End If
  
          'mono channel
          SelectObject cBack.hdc, hPenPeakRight
          'upper or both
          If ScopeSettings.bPeakCount = 0 Or ScopeSettings.bPeakCount = 2 Then
            MoveToEx cBack.hdc, 0, ((ScopeUPeaks(1).Value + 32768) / 65535 * cBack.lHeight), point
            For X = 1 To UBound(ScopeUPeaks) Step 2
              W = (cBack.lWidth * X * 2) / UBound(ScopeUPeaks)
              H = ((ScopeUPeaks(X).Value + 32768) / 65535 * cBack.lHeight)
              LineTo cBack.hdc, W, H
            Next
          End If
          'lower or both
          If ScopeSettings.bPeakCount > 0 Then
            MoveToEx cBack.hdc, 0, ((ScopeLPeaks(1).Value + 32768) / 65535 * cBack.lHeight), point
            For X = 1 To UBound(ScopeLPeaks) Step 2
              W = (cBack.lWidth * X * 2) / UBound(ScopeLPeaks)
              H = ((ScopeLPeaks(X).Value + 32768) / 65535 * cBack.lHeight)
              LineTo cBack.hdc, W, H
            Next
          End If

        End If
      End If

    Else 'mixer is MMX
      'The following is EXACTLY the same as the above, but using
      'ScopeBufferINT (integer) instead of ScopeBufferFPU (single)
      'Sure, I could convert the single values to intgers to save
      'alot of code, but... ah what the hell.
      
      Call GetScopeINT(ScopeBufferINT) 'Get sound data
      'we get data in format INT for MMX mixers
      
      'Update History (for the history vis)
      For X = 0 To UBound(ScopeHistoryL) - 1
        ScopeHistoryL(X) = ScopeHistoryL(X + 1)
        ScopeHistoryR(X) = ScopeHistoryR(X + 1)
      Next
      'Actually I don't really know which comes first, Left or Right but who cares, right
      ScopeHistoryL(UBound(ScopeHistoryL)) = ScopeBufferINT(0)
      ScopeHistoryR(UBound(ScopeHistoryR)) = ScopeBufferINT(1)
      
      'Handle peaks
      If ScopeSettings.bPeaks > 0 Then
        For X = 0 To UBound(ScopeBufferINT)
          'upper peaks first!
          With ScopeUPeaks(X)
            If ScopeSettings.bFall = 1 Then
              If ScopeSettings.lPeakDec = 0 Then ScopeSettings.lPeakDec = 20
            Else
              If ScopeSettings.lPeakDec = 0 Then ScopeSettings.lPeakDec = 75
              .Dec = ScopeSettings.lPeakDec * 10
            End If
            If .Value - .Dec > ScopeBufferINT(X) Then
              .Value = ScopeBufferINT(X)
              .Time = timeGetTime
              .Pause = .Time
              .Dec = 0
            ElseIf timeGetTime - .Pause > ScopeSettings.lPeakPause Then
              If timeGetTime - .Time > 10 Then
                .Time = timeGetTime
                If ScopeSettings.bFall = 1 Then
                  .Dec = .Dec + ScopeSettings.lPeakDec
                End If
                .Value = .Value + .Dec
                If .Value > 0 Then
                  .Value = 0
                  .Dec = 0
                End If
              End If
            End If
          End With
          
          With ScopeLPeaks(X)
            If ScopeSettings.bFall = 1 Then
              If ScopeSettings.lPeakDec = 0 Then ScopeSettings.lPeakDec = 20
            Else
              If ScopeSettings.lPeakDec = 0 Then ScopeSettings.lPeakDec = 75
              .Dec = ScopeSettings.lPeakDec * 10
            End If
            If .Value + .Dec < ScopeBufferINT(X) Then
              .Value = ScopeBufferINT(X)
              .Time = timeGetTime
              .Pause = .Time
              .Dec = 0
            ElseIf timeGetTime - .Pause > ScopeSettings.lPeakPause Then
              If timeGetTime - .Time > 10 Then
                .Time = timeGetTime
                If ScopeSettings.bFall = 1 Then
                  .Dec = .Dec + ScopeSettings.lPeakDec
                End If
                .Value = .Value - .Dec
                If .Value < 0 Then
                  .Value = 0
                  .Dec = 0
                End If
              End If
            End If
          End With
        Next
      End If

      'Draw the scope
      If ScopeSettings.bType = 0 Then 'Dot Scope
        If ScopeSettings.bDetail = 0 Then
          For X = 0 To UBound(ScopeBufferINT) Step ScopeSettings.bSkip
            H = ((ScopeBufferINT(X) + 32768) / 65535 * cBack.lHeight)
            W = (cBack.lWidth * X * 2) / UBound(ScopeBufferINT)
            SetPixelV cBack.hdc, W, H, ScopeSettings.lColorL
          Next
        End If

        For X = 1 To UBound(ScopeBufferINT) Step ScopeSettings.bSkip
          H = ((ScopeBufferINT(X) + 32768) / 65535 * cBack.lHeight)
          W = (cBack.lWidth * X * 2) / UBound(ScopeBufferINT)
          SetPixelV cBack.hdc, W, H, ScopeSettings.lColorR
        Next
               
      ElseIf ScopeSettings.bType = 1 Then 'Line Scope

        If ScopeSettings.bDetail = 0 Then
          SelectObject cBack.hdc, hPenLeft
          MoveToEx cBack.hdc, 0, ((ScopeBufferINT(0) + 32768) / 65535 * cBack.lHeight), point
          For X = 0 To UBound(ScopeBufferINT) Step ScopeSettings.bSkip
            H = ((ScopeBufferINT(X) + 32768) / 65535 * cBack.lHeight)
            W = (cBack.lWidth * X * 2) / UBound(ScopeBufferINT)
            LineTo cBack.hdc, W, H
          Next
        End If

        SelectObject cBack.hdc, hPenRight
        MoveToEx cBack.hdc, 0, ((ScopeBufferINT(1) + 32768) / 65535 * cBack.lHeight), point
        For X = 1 To UBound(ScopeBufferINT) Step ScopeSettings.bSkip
          H = ((ScopeBufferINT(X) + 32768) / 65535 * cBack.lHeight)
          W = (cBack.lWidth * X * 2) / UBound(ScopeBufferINT)
          LineTo cBack.hdc, W, H
        Next

      ElseIf ScopeSettings.bType = 2 Then 'Solid Scope
        
        If ScopeSettings.bDetail = 0 Then
          For X = 0 To UBound(ScopeBufferINT) Step ScopeSettings.bSkip
            mRect.Bottom = cBack.lHeight / 2
            mRect.Top = ((ScopeBufferINT(X) + 32768) / 65535 * cBack.lHeight)
            mRect.Left = (cBack.lWidth * X * 2) / UBound(ScopeBufferINT)
            mRect.Right = (cBack.lWidth * (X + 2) * 2) / UBound(ScopeBufferINT)
            FillRect cBack.hdc, mRect, hBrushSolidLeft
          Next
        End If
        
        'MoveToEx cBack.hdc, 0, (cBack.lHeight / 2), point
        'LineTo cBack.hdc, cBack.lWidth, (cBack.lHeight / 2)
        For X = 1 To UBound(ScopeBufferINT) Step ScopeSettings.bSkip
          mRect.Bottom = cBack.lHeight / 2
          mRect.Top = ((ScopeBufferINT(X) + 32768) / 65535 * cBack.lHeight)
          mRect.Left = (cBack.lWidth * X * 2) / UBound(ScopeBufferINT)
          mRect.Right = (cBack.lWidth * (X + 2) * 2) / UBound(ScopeBufferINT)
          FillRect cBack.hdc, mRect, hBrushSolidRight
        Next

      Else 'History Scope

        If ScopeSettings.bDetail = 0 Then
          SelectObject cBack.hdc, hPenLeft
          MoveToEx cBack.hdc, 0, ((ScopeHistoryL(0) + 32768) / 65535 * cBack.lHeight), point
          For X = 0 To UBound(ScopeHistoryL)
            LineTo cBack.hdc, X * 2, ((ScopeHistoryL(X) + 32768) / 65535 * cBack.lHeight)
          Next
        End If

        SelectObject cBack.hdc, hPenRight
        MoveToEx cBack.hdc, 0, ((ScopeHistoryR(0) + 32768) / 65535 * cBack.lHeight), point
        For X = 0 To UBound(ScopeHistoryR)
          LineTo cBack.hdc, X * 2, ((ScopeHistoryR(X) + 32768) / 65535 * cBack.lHeight)
        Next

      End If

      'Draw Peaks!
      If ScopeSettings.bType < 3 Then
        If ScopeSettings.bPeaks = 1 Then  'dots
  
          'stereo channel
          If ScopeSettings.bDetail = 0 And ScopeSettings.bPeakDetail = 0 Then
            For X = 0 To UBound(ScopeUPeaks) Step 2
              W = (cBack.lWidth * X * 2) / UBound(ScopeUPeaks)
              'Upper or both
              If ScopeSettings.bPeakCount = 0 Or ScopeSettings.bPeakCount = 2 Then
                H = ((ScopeUPeaks(X).Value + 32768) / 65535 * cBack.lHeight)
                SetPixelV cBack.hdc, W, H, ScopeSettings.lColorPeakL
              End If
              'lower or both
              If ScopeSettings.bPeakCount > 0 Then
                H = ((ScopeLPeaks(X).Value + 32768) / 65535 * cBack.lHeight)
                SetPixelV cBack.hdc, W, H, ScopeSettings.lColorPeakL
              End If
            Next
          End If
  
          'mono channel
          For X = 1 To UBound(ScopeUPeaks) Step 2
            W = (cBack.lWidth * X * 2) / UBound(ScopeUPeaks)
            'Upper or both
            If ScopeSettings.bPeakCount = 0 Or ScopeSettings.bPeakCount = 2 Then
              H = ((ScopeUPeaks(X).Value + 32768) / 65535 * cBack.lHeight)
              SetPixelV cBack.hdc, W, H, ScopeSettings.lColorPeakR
            End If
            'lower or both
            If ScopeSettings.bPeakCount > 0 Then
              H = ((ScopeLPeaks(X).Value + 32768) / 65535 * cBack.lHeight)
              SetPixelV cBack.hdc, W, H, ScopeSettings.lColorPeakR
            End If
          Next
  
        ElseIf ScopeSettings.bPeaks = 2 Then  'lines
  
          'Stereo channel
          If ScopeSettings.bDetail = 0 And ScopeSettings.bPeakDetail = 0 Then
            SelectObject cBack.hdc, hPenPeakLeft
            'upper or both
            If ScopeSettings.bPeakCount = 0 Or ScopeSettings.bPeakCount = 2 Then
              MoveToEx cBack.hdc, 0, ((ScopeUPeaks(0).Value + 32768) / 65535 * cBack.lHeight), point
              For X = 0 To UBound(ScopeUPeaks) Step 2
                W = (cBack.lWidth * X * 2) / UBound(ScopeUPeaks)
                H = ((ScopeUPeaks(X).Value + 32768) / 65535 * cBack.lHeight)
                LineTo cBack.hdc, W, H
              Next
            End If
            'lower or both
            If ScopeSettings.bPeakCount > 0 Then
              MoveToEx cBack.hdc, 0, ((ScopeLPeaks(0).Value + 32768) / 65535 * cBack.lHeight), point
              For X = 0 To UBound(ScopeLPeaks) Step 2
                W = (cBack.lWidth * X * 2) / UBound(ScopeLPeaks)
                H = ((ScopeLPeaks(X).Value + 32768) / 65535 * cBack.lHeight)
                LineTo cBack.hdc, W, H
              Next
            End If
          End If
  
          'mono channel
          SelectObject cBack.hdc, hPenPeakRight
          'upper or both
          If ScopeSettings.bPeakCount = 0 Or ScopeSettings.bPeakCount = 2 Then
            MoveToEx cBack.hdc, 0, ((ScopeUPeaks(1).Value + 32768) / 65535 * cBack.lHeight), point
            For X = 1 To UBound(ScopeUPeaks) Step 2
              W = (cBack.lWidth * X * 2) / UBound(ScopeUPeaks)
              H = ((ScopeUPeaks(X).Value + 32768) / 65535 * cBack.lHeight)
              LineTo cBack.hdc, W, H
            Next
          End If
          'lower or both
          If ScopeSettings.bPeakCount > 0 Then
            MoveToEx cBack.hdc, 0, ((ScopeLPeaks(1).Value + 32768) / 65535 * cBack.lHeight), point
            For X = 1 To UBound(ScopeLPeaks) Step 2
              W = (cBack.lWidth * X * 2) / UBound(ScopeLPeaks)
              H = ((ScopeLPeaks(X).Value + 32768) / 65535 * cBack.lHeight)
              LineTo cBack.hdc, W, H
            Next
          End If
  
        End If
      End If
    End If
  End If
  
  'present the backbuffer
  cBack.PaintTo picVis.hdc, 0, 0, vbSrcCopy
  
End Sub

Public Sub visSpectrum()
  'Draws an Spectrum Analyzer
  On Error Resume Next
  
  If Not Sound.StreamIsPlaying And Not Sound.MusicIsPlaying Then
    'If no sound is playing, just fade out/clear the screen
    
    If SpectrumSettings.bFade > 0 Then
      cBack.Fade SpectrumSettings.bFade
    Else
      cBack.BitBltFrom cBackOrig.hdc
    End If
    
  Else
    'If sound IS playing, draw the visualization!
    
    Dim Spec() As Single, point As POINTAPI 'point for MoveToEx
    Dim X As Single, y As Long, t As Long
    Dim h2 As Single, w2 As Single, nH As Single, b As Integer
  
    'clear backbuffer
    If SpectrumSettings.bFade > 0 Then
      cBack.Fade SpectrumSettings.bFade
    Else
      cBack.BitBltFrom cBackOrig.hdc
    End If
  
    'redimension the array if needed
    If UBound(Spec) <> SpectrumSettings.iView Then
      ReDim Spec(SpectrumSettings.iView) As Single
    End If
    GetSpectrum Spec() 'Get spectrum data

    If SpectrumSettings.bType = 0 Then 'Normal style
      'This just plots all the values in Spec() to the screen, left to right.
      w2 = cBack.lWidth \ SpectrumSettings.iView + 1
      If w2 < 1 Then w2 = 1
      h2 = cBack.lHeight * SpectrumSettings.nZoom
      For t = 0 To SpectrumSettings.iView
        X = X + w2
        y = (Spec(t) + ((Spec(t) * (t / 511)) * SpectrumSettings.bCorrection)) * h2
        If y > cBack.lHeight Then y = cBack.lHeight
        If SpectrumSettings.bDrawStyle = 0 Then
          TransparentBlt cBack.hdc, X - w2, cBack.lHeight - y, w2, y, _
          cGradBar.hdc, 0, cBack.lHeight - y, 1, y, 0
        ElseIf SpectrumSettings.bDrawStyle = 1 Then
          TransparentBlt cBack.hdc, X - w2, cBack.lHeight - y, w2, y, _
          cGradBar.hdc, 0, 0, 1, cGradBar.lHeight, 0
        Else
          TransparentBlt cBack.hdc, X - w2, cBack.lHeight - y, w2, y, _
          cGradBar.hdc, 0, cBack.lHeight - y, 1, 1, 0
        End If
        
      Next t
      
    ElseIf SpectrumSettings.bType = 1 Then 'Bar style
      'This plots all the values in Spec() to the screen in bars.
      'If you have 20 bars and there are 512 values, each bar will
      'show 512 / 20 values (the values will be added and max value
      'will be 512 / 20). It will also draw peaks if selected.
  
      'Make sure array to hold bars is correctly dimensioned
      y = (cBack.lWidth - 3) / (SpectrumSettings.bBarSize + 1)
      'Debug.Print y, y * (SpectrumSettings.bBarSize + 1), cBack.lWidth
      If UBound(SpectrumBars) <> y Then ReDim Preserve SpectrumBars(y)
      If UBound(SpectrumPeaks) <> y Then ReDim Preserve SpectrumPeaks(y)
  
      'Debug.Print UBound(Spec) / y
      y = CInt(UBound(Spec) / y)
      
      If y < 1 Then y = 1
      h2 = cBack.lHeight * SpectrumSettings.nZoom
      X = 2
      
      SelectObject cBack.hdc, hPenPeakSpec 'select the correct pen
      
      'Begin processing & drawing each bar
      For t = 0 To UBound(Spec) Step y
        nH = 0
        
        'This adds all values within one bar
        For b = t To t + (y - 1)
          nH = nH + Spec(b) + ((Spec(b) * (b / 511)) * SpectrumSettings.bCorrection)
        Next b
        nH = nH / y

        X = X + (SpectrumSettings.bBarSize + 1)
        b = Int(t \ y)
        
        'calculate the new position of the bar
        With SpectrumBars(b)
          If SpectrumSettings.bFall = 1 Then
            If SpectrumSettings.nDec = 0 Then SpectrumSettings.nDec = 10
          Else
            If SpectrumSettings.nDec = 0 Then SpectrumSettings.nDec = 50
            .Dec = SpectrumSettings.nDec / 1000
          End If
          If nH > .Value - .Dec Then
            .Value = nH
            .Time = timeGetTime
            .Pause = .Time
            .Dec = 0
          ElseIf timeGetTime - .Pause > SpectrumSettings.lPause Then
            If timeGetTime - .Time > 10 Then
              .Time = timeGetTime
              If SpectrumSettings.bFall = 1 Then
                .Dec = .Dec + y * (SpectrumSettings.nDec / 10000)
              End If
              .Value = .Value - .Dec
              If .Value < 0 Then
                .Value = 0
                .Dec = 0
              End If
            End If
          End If
        End With
        
        'draw bar
        w2 = (SpectrumBars(b).Value * h2)
        If w2 > cBack.lHeight Then w2 = cBack.lHeight
        If SpectrumSettings.bDrawStyle = 0 Then
          TransparentBlt cBack.hdc, X - SpectrumSettings.bBarSize, cBack.lHeight - w2, SpectrumSettings.bBarSize, w2, _
          cGradBar.hdc, 0, cBack.lHeight - w2, 1, w2, 0
        ElseIf SpectrumSettings.bDrawStyle = 1 Then
          TransparentBlt cBack.hdc, X - SpectrumSettings.bBarSize, cBack.lHeight - w2, SpectrumSettings.bBarSize, w2, _
          cGradBar.hdc, 0, 0, 1, cGradBar.lHeight, 0
        Else
          TransparentBlt cBack.hdc, X - SpectrumSettings.bBarSize, cBack.lHeight - w2, SpectrumSettings.bBarSize, w2, _
          cGradBar.hdc, 0, cBack.lHeight - w2, 1, 1, 0
        End If
  
        'process peak for bar
        If SpectrumSettings.bPeaks Then
          With SpectrumPeaks(b)
            If SpectrumSettings.bPeakFall = 1 Then
              If SpectrumSettings.lPeakDec = 0 Then SpectrumSettings.lPeakDec = 2.5
            Else
              If SpectrumSettings.lPeakDec = 0 Then SpectrumSettings.lPeakDec = 15
              .Dec = SpectrumSettings.lPeakDec / 1000
            End If
            If SpectrumBars(b).Value > .Value - .Dec Then
              .Value = SpectrumBars(b).Value
              .Time = timeGetTime
              .Pause = .Time
              .Dec = 0
            ElseIf timeGetTime - .Pause > SpectrumSettings.lPeakPause Then
              If timeGetTime - .Time > 10 Then
                .Time = timeGetTime
                If SpectrumSettings.bPeakFall = 1 Then
                  .Dec = .Dec + y * (SpectrumSettings.lPeakDec / 10000)
                End If
                .Value = .Value - .Dec
                If .Value < 0 Then
                  .Value = 0
                  .Dec = 0
                End If
              End If
            End If
          End With
          
          'Draw peak line
          w2 = cBack.lHeight - (SpectrumPeaks(b).Value * h2)
          If w2 < 0 Then w2 = 0
          MoveToEx cBack.hdc, X - SpectrumSettings.bBarSize, w2, point
          LineTo cBack.hdc, X, w2
        End If
    
      Next t
      
    ElseIf SpectrumSettings.bType = 2 Then 'line style

      SelectObject cBack.hdc, hPenLineSpec
      
      w2 = cBack.lWidth \ SpectrumSettings.iView + 1
      If w2 < 1 Then w2 = 1
      h2 = cBack.lHeight * SpectrumSettings.nZoom
      X = 1
      y = cBack.lHeight - ((Spec(t) + ((Spec(t) * (t / 511)) * SpectrumSettings.bCorrection)) * h2)
      If y > cBack.lHeight Then y = cBack.lHeight
      MoveToEx cBack.hdc, X, y, point
      For t = 1 To SpectrumSettings.iView
        X = X + w2
        y = cBack.lHeight - ((Spec(t) + ((Spec(t) * (t / 511)) * SpectrumSettings.bCorrection)) * h2)
        If y > cBack.lHeight Then y = cBack.lHeight
        LineTo cBack.hdc, X, y
      Next t
 
    End If
    
  End If
  
  'show backbuffer to picturebox
  cBack.PaintTo picVis.hdc, 0, 0, vbSrcCopy

End Sub

Public Sub visVolume()
  On Error Resume Next

    
  If Not Sound.StreamIsPlaying And Not Sound.MusicIsPlaying Then
    'If no sound is playing, just fade out/clear the screen

    If VolumeSettings.bFade > 0 Then
      cBack.Fade VolumeSettings.bFade
    Else
      cBack.BitBltFrom cBackOrig.hdc
    End If
  
  Else
    'If sound IS playing, draw the visualization!

    Dim VU(1) As Single, point As POINTAPI
    Dim W As Integer, H As Integer, X As Integer, vb(1) As Single
    
    'clear backbuffer
    If VolumeSettings.bFade > 0 Then
      cBack.Fade VolumeSettings.bFade
    Else
      cBack.BitBltFrom cBackOrig.hdc
    End If
    
    If Sound.MusicIsLoaded Then
      'When modules are playing, get the channel playing the highest volume
      For H = 1 To FSOUND_GetMaxChannels
        Sound.GetSpecificVU H, vb(0), vb(1)
        If vb(0) > VU(0) Then VU(0) = vb(0)
        If vb(1) > VU(1) Then VU(1) = vb(1)
      Next
      'mod, s3m & xm only goes up to ~0.34 instead of 1.0, so multiply them
      Select Case Library(PlayingLib).eType
        Case TYPE_MOD, TYPE_S3M, TYPE_XM
          VU(0) = VU(0) * 2.87
          VU(1) = VU(1) * 2.87
      End Select
    ElseIf Sound.StreamIsLoaded Then
      Sound.GetVU VU(0), VU(1) 'Get VU values
    End If
    
    For W = 0 To 1
      With VolumeBars(W)
        If VolumeSettings.bFall = 1 Then
          If VolumeSettings.nDec = 0 Then VolumeSettings.nDec = 1
        Else
          If VolumeSettings.nDec = 0 Then VolumeSettings.nDec = 50
          .Dec = VolumeSettings.nDec / 1000
        End If
        If VU(W) > .Value - .Dec Then
          .Value = VU(W)
          .Time = timeGetTime
          .Pause = .Time
          .Dec = 0
        ElseIf timeGetTime - .Pause > VolumeSettings.lPause Then
          If timeGetTime - .Time > 10 Then
            .Time = timeGetTime
            If VolumeSettings.bFall = 1 Then
              .Dec = .Dec + (VolumeSettings.nDec / 1000)
            End If
            .Value = .Value - .Dec
            If .Value < 0 Then
              .Value = 0
              .Dec = 0
            End If
          End If
        End If
      End With
    Next W
    
    'setup history (for history volume meters)
    For W = 0 To UBound(VolumeHistory) - 1 'move history
      VolumeHistory(W) = VolumeHistory(W + 1)
    Next W
    VolumeHistory(UBound(VolumeHistory)) = VU(0) + VU(1) 'add new value
    
    If VolumeSettings.bType = 0 Then 'Two graphic bars
      H = (cBack.lHeight / 4) - (cBar.lHeight / 2)
      W = cBar.lWidth * VolumeBars(0).Value
      cBar.PaintToTransCrop cBack.hdc, 2, H, W
      H = ((cBack.lHeight / 4) * 3) - (cBar.lHeight / 2)
      W = cBar.lWidth * VolumeBars(1).Value
      cBar.PaintToTransCrop cBack.hdc, 2, H, W
    
    ElseIf VolumeSettings.bType = 1 Then
      
      For W = 0 To UBound(VolumeHistory)
        X = X + 2
        H = ((cBack.lHeight / 2) * VolumeHistory(W))

        If VolumeSettings.bDrawStyle = 0 Then
          TransparentBlt cBack.hdc, X - 1, cBack.lHeight - H, 2, H, _
          cGradVol.hdc, 0, cBack.lHeight - H, 2, H, 0
        ElseIf VolumeSettings.bDrawStyle = 1 Then
          TransparentBlt cBack.hdc, X - 1, cBack.lHeight - H, 2, H, _
          cGradVol.hdc, 0, 0, 2, cGradVol.lHeight, 0
        Else
          TransparentBlt cBack.hdc, X - 1, cBack.lHeight - H, 2, H, _
          cGradVol.hdc, 0, cBack.lHeight - H, 2, 1, 0
        End If
        
      Next
  
    End If
  
  End If
  
  'show backbuffer to picturebox
  cBack.PaintTo picVis.hdc, 0, 0, vbSrcCopy
  
End Sub

Public Sub visBeat()
  'Draws an image pulsating after beat
  On Error Resume Next
  
  If Not Sound.StreamIsPlaying And Not Sound.MusicIsPlaying Then
    'If no sound is playing, just fade out/clear the screen
    
    If BeatSettings.bFade > 0 Then
      cBack.Fade BeatSettings.bFade
    Else
      cBack.BitBltFrom cBackOrig.hdc
    End If
  
  Else
    'If sound IS playing, draw the visualization!
  
    Dim Spec() As Single, X As Long, bV2 As Single
    Static bV As Single, bH As Single, bH2 As Single, bD As Byte 'static vars
    
    'dimension the array
    ReDim Spec(275) As Single
    GetSpectrum Spec() 'Get spectrum data
    
    'Calculate beat
    For X = BeatSettings.iDetectLow To BeatSettings.iDetectHigh
      bV2 = bV2 + Spec(X)
    Next
    
    'Figure out zoom value
    'zoom = (all read values / # read values) * multiplyer + minimum zoom
    bV2 = (bV2 / (BeatSettings.iDetectHigh - BeatSettings.iDetectLow)) * BeatSettings.nMulti + BeatSettings.nMin
    
    If BeatSettings.bType = 0 Then 'swing
    
      'Figure out rotation
      'bD = direction to rotate, switches each time
      If bV = 0 Then bV = bV2
      If bV2 - bV > BeatSettings.nRotMin Then  'if increase in zoom was high enough, rotate
        'determination of final destinaion pos
        If bD = 0 Then
          bH = (bV2 - bV) '/ 2
          bD = 1
        Else
          bD = 0
          bH = (bV - bV2) '/ 2
        End If
      Else
        'movement of destination pos
        If bH < -BeatSettings.nRotMove Then
          bH = bH + BeatSettings.nRotMove
        ElseIf bH > BeatSettings.nRotMove Then
          bH = bH - BeatSettings.nRotMove
        Else
          bH = 0
        End If
      End If
      'movement of current pos
      If bH2 > bH + BeatSettings.nRotMove Then
        bH2 = bH2 - BeatSettings.nRotMove
      ElseIf bH2 < bH - BeatSettings.nRotMove Then
        bH2 = bH2 + BeatSettings.nRotMove
      Else
        bH2 = bH
      End If
        
      bV = bV2
    
    ElseIf BeatSettings.bType = 1 Then 'rotation
      'bH2 = rotation acc.
      
      If bV2 - bV > BeatSettings.nRotMin Then  'if increase in zoom was high enough, rotate
        bH = (bV2 - bV) / 4
      ElseIf bH > 0 Then
        bH = bH - 0.01
        If bH < 0 Then bH = 0
      End If
      
      bH2 = bH2 + BeatSettings.nRotSpeed + bH
      If bH2 > 6.28 Then bH2 = bH2 - 6.28 'wrap around radians (2*PI = full circle)
      
      bV = bV2
    End If
    
    'clear backbuffer
    If BeatSettings.bFade > 0 Then
      cBack.Fade BeatSettings.bFade
    Else
      cBack.BitBltFrom cBackOrig.hdc
    End If
    
    'do the actual rotation & zooming
    cBeat.RotateZoomTo cBeatTemp, cBack, 1.58 + bH2, bV
    
  End If
    
  'show backbuffer to picturebox
  cBack.PaintTo picVis.hdc, 0, 0, vbSrcCopy
  
End Sub

Public Sub DoGrad(ByVal lUpper As Long, ByVal lLower As Long)
  'Draws gradient on picturebox p
  'same height as visualization drawing area
  On Error Resume Next
  Dim RGB1(2) As Single
  Dim RGB2(2) As Single

  p.Cls
  
  'Convert to rgb
  RGB1(0) = lUpper Mod &H100
  RGB1(1) = (lUpper \ &H100) Mod &H100
  RGB1(2) = (lUpper \ &H10000) Mod &H100
  RGB2(0) = lLower Mod &H100
  RGB2(1) = (lLower \ &H100) Mod &H100
  RGB2(2) = (lLower \ &H10000) Mod &H100
  Dim y As Long, nC(2) As Single, cc(2) As Single
  cc(0) = RGB1(0): cc(1) = RGB1(1): cc(2) = RGB1(2)
  nC(0) = (RGB2(0) - RGB1(0)) / picVis.ScaleHeight
  nC(1) = (RGB2(1) - RGB1(1)) / picVis.ScaleHeight
  nC(2) = (RGB2(2) - RGB1(2)) / picVis.ScaleHeight
  For y = 0 To picVis.ScaleHeight
    p.Line (0, y)-(10, y), RGB(cc(0), cc(1), cc(2))
    cc(0) = cc(0) + nC(0)
    cc(1) = cc(1) + nC(1)
    cc(2) = cc(2) + nC(2)
  Next
  
End Sub
