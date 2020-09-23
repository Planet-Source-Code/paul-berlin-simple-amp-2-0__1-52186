VERSION 5.00
Begin VB.UserControl ctrlLabel 
   BackColor       =   &H00C0C0FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2355
   ClipBehavior    =   0  'None
   ScaleHeight     =   54
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   157
   Windowless      =   -1  'True
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   360
   End
End
Attribute VB_Name = "ctrlLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ctrlLabel 1.0
'By Paul Berlin 2003
'This was written with simple amp in mind, and may need some modification
'if used in other projects.

Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
'Private Declare Function timeGetTime Lib "winmm.dll" () As Long

'Private Type RECT
'  Left As Long
'  Top As Long
'  Right As Long
'  Bottom As Long
'End Type

Private Const DT_VCENTER = &H4
Private Const DT_NOPREFIX = &H800

Private lAlignment As Long 'is ignored when scrolling, scroll is always left
Private sCaption As String

Private Const PAUSE_LENGTH As Long = 3000 'ms to pause at end of each scroll

Private lPauseStart As Long 'time of start of pause, if over PAUSE_LENGTH has passed it will scroll
Private lCaptionLenPx As Long 'caption length in pixels
Private lCurScrollPx As Long 'Current horizontal text offset
Private lScrollLeft As Boolean 'true if scroll left, false if right

Private bGotBack As Boolean
Private cBack As New clsImg 'original back image

Private bMouseDown As Boolean
Private lOldMouseX As Long

Event Click()
Event DblClick()
Event MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)

'#### PROPERTIES ####
'FONT
Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
  Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set UserControl.Font = New_Font
  PropertyChanged "Font"
  ResetVars
  UserControl_Paint
End Property

'FORECOLOR/TEXTCOLOR
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
  ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal lNew_Color As OLE_COLOR)
  UserControl.ForeColor = lNew_Color
  PropertyChanged "ForeColor"
  ResetVars
  UserControl_Paint
End Property

'CAPTION
Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
  Caption = sCaption
End Property

Public Property Let Caption(ByVal sNew_Cap As String)
  sCaption = sNew_Cap
  PropertyChanged "Caption"
  ResetVars
  'UserControl_Paint
End Property

'AUTOSCROLL
Public Property Get AutoScroll() As Boolean
  AutoScroll = tmrScroll.Enabled
End Property

Public Property Let AutoScroll(ByVal bNew As Boolean)
  tmrScroll.Enabled = bNew
  bGotBack = 0
End Property

'ALIGNMENT
Public Property Get Alignment() As Long
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Misc"
  Alignment = lAlignment
End Property

Public Property Let Alignment(ByVal lNewAlign As Long)
  If lNewAlign = 3 Then
    lNewAlign = 2
  ElseIf lNewAlign = 2 Then
    lNewAlign = 3
  End If
  lAlignment = lNewAlign
  PropertyChanged "Alignment"
  ResetVars
  UserControl_Paint
End Property

Private Sub tmrScroll_Timer()
  If timeGetTime - lPauseStart > PAUSE_LENGTH Then
    If lCaptionLenPx > UserControl.ScaleWidth Then
      If lScrollLeft Then
        lCurScrollPx = lCurScrollPx - 1
        If lCurScrollPx <= UserControl.ScaleWidth - lCaptionLenPx - 5 Then
          lScrollLeft = False
          lPauseStart = timeGetTime
        End If
      Else
        lCurScrollPx = lCurScrollPx + 1
        If lCurScrollPx >= 5 Then
          lScrollLeft = True
          lPauseStart = timeGetTime
        End If
      End If
      
    End If
  End If
  UserControl_Paint
End Sub

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub UserControl_HitTest(X As Single, y As Single, HitResult As Integer)
  HitResult = 3
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  If lCaptionLenPx > UserControl.ScaleWidth Then
    bMouseDown = True
    lOldMouseX = X - lCurScrollPx
    lPauseStart = timeGetTime
  Else
    RaiseEvent MouseDown(Button, Shift, X, y)
  End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  If lCaptionLenPx > UserControl.ScaleWidth Then
    If bMouseDown Then
      lCurScrollPx = X - lOldMouseX
      If lCurScrollPx <= UserControl.ScaleWidth - lCaptionLenPx - 5 Then
        lCurScrollPx = UserControl.ScaleWidth - lCaptionLenPx - 5
      ElseIf lCurScrollPx >= 5 Then
        lCurScrollPx = 5
      End If
      UserControl_Paint
      lPauseStart = timeGetTime
    End If
  Else
    RaiseEvent MouseMove(Button, Shift, X, y)
  End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  If lCaptionLenPx > UserControl.ScaleWidth Then
    bMouseDown = False
  Else
    RaiseEvent MouseUp(Button, Shift, X, y)
  End If
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, y)
End Sub

'#### USERCONTROL FUNCTIONS ####
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  'Load properties
  UserControl.Font = PropBag.ReadProperty("Font", "MS Sans Serif")
  UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
  sCaption = PropBag.ReadProperty("Caption", "")
  lAlignment = PropBag.ReadProperty("Alignment", 0)
  
  'setup other
  ResetVars
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  'save properties
  PropBag.WriteProperty "Font", UserControl.Font, "MS Sans Serif"
  PropBag.WriteProperty "ForeColor", UserControl.ForeColor, &H80000012
  PropBag.WriteProperty "Caption", sCaption, ""
  PropBag.WriteProperty "Alignment", lAlignment, 0
End Sub

'Drawing of control
Private Sub UserControl_Paint()
  On Error GoTo errh
  Dim r As RECT
  
  If Not bGotBack Then
    'grab background image
    Set frmMain.p.Picture = frmMain.Picture
    cBack.Create UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hdc
    BitBlt cBack.hdc, 0, 0, UserControl.Extender.Width, UserControl.Extender.Height, frmMain.p.hdc, UserControl.Extender.Left, UserControl.Extender.Top, vbSrcCopy
    Set frmMain.p.Picture = Nothing
    bGotBack = True
  End If
  
  With UserControl
    
    Set frmMain.p.Font = .Font
    frmMain.p.ForeColor = .ForeColor
    frmMain.p.Width = .ScaleWidth
    cBack.PaintTo frmMain.p.hdc, 0, 0, vbSrcCopy
    
    If lCaptionLenPx > .ScaleWidth Then  'scroll needed
      r.Left = lCurScrollPx
      r.Top = 0
      r.Right = lCaptionLenPx
      r.Bottom = .ScaleHeight

      DrawText frmMain.p.hdc, sCaption, Len(sCaption), r, DT_NOPREFIX Or DT_VCENTER

    Else
      r.Left = 0
      r.Top = 0
      r.Right = .ScaleWidth
      r.Bottom = .ScaleHeight

      DrawText frmMain.p.hdc, sCaption, Len(sCaption), r, lAlignment Or DT_NOPREFIX Or DT_VCENTER

    End If
    
    BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, frmMain.p.hdc, 0, 0, vbSrcCopy

  End With
  
errh:
End Sub

Private Sub ResetVars()
  lCaptionLenPx = UserControl.TextWidth(sCaption)
  lCurScrollPx = 5
  lScrollLeft = True
  lPauseStart = timeGetTime
End Sub

