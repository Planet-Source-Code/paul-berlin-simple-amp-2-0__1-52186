VERSION 5.00
Begin VB.UserControl CtrlButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5475
   ClipControls    =   0   'False
   MaskColor       =   &H00FF00FF&
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   246
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   365
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4080
      Top             =   2760
   End
   Begin VB.PictureBox picOffUpOver 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   495
      Left            =   3120
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   4
      Top             =   2760
      Width           =   495
      Visible         =   0   'False
   End
   Begin VB.PictureBox picOffUp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   495
      Left            =   2400
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   3
      Top             =   2760
      Width           =   495
      Visible         =   0   'False
   End
   Begin VB.PictureBox picDown 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   495
      Left            =   1680
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   2
      Top             =   2760
      Width           =   495
      Visible         =   0   'False
   End
   Begin VB.PictureBox picUpOver 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   495
      Left            =   960
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   1
      Top             =   2760
      Width           =   495
      Visible         =   0   'False
   End
   Begin VB.PictureBox picUp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   495
      Left            =   240
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   0
      Top             =   2760
      Width           =   495
      Visible         =   0   'False
   End
End
Attribute VB_Name = "CtrlButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CtrlButton 1.1
'Created by Paul Berlin 2002
'Simple Button & On/Off Button

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'Private Type POINTAPI
'  X As Long
'  y As Long
'End Type

' Declarations
Dim bExt As Boolean
Dim bOn As Boolean

Dim bPressed As Boolean 'if true, button is down
Dim bUsedto As Boolean 'if true, button was pressed and mouse is outside button, mousebutton down
Dim bOver As Boolean 'mouse is over

' Events
Event Pressed(ByVal Button As Integer)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)

Private Sub Draw()
  On Error Resume Next
  UserControl.Cls
  
  If bPressed Then
    BitBlt UserControl.hdc, 0, 0, picDown.Width, picDown.Height, picDown.hdc, 0, 0, vbSrcCopy
  Else
    If bExt Then
      If bOn Then
        BitBlt UserControl.hdc, 0, 0, picUp.Width, picUp.Height, picUp.hdc, 0, 0, vbSrcCopy
      Else
        BitBlt UserControl.hdc, 0, 0, picOffUp.Width, picOffUp.Height, picOffUp.hdc, 0, 0, vbSrcCopy
      End If
    Else
      BitBlt UserControl.hdc, 0, 0, picUp.Width, picUp.Height, picUp.hdc, 0, 0, vbSrcCopy
    End If
  End If
  If bOver Then
    If bExt Then
      If bOn Then
        BitBlt UserControl.hdc, 0, 0, picUpOver.Width, picUpOver.Height, picUpOver.hdc, 0, 0, vbSrcCopy
      Else
        BitBlt UserControl.hdc, 0, 0, picOffUpOver.Width, picOffUpOver.Height, picOffUpOver.hdc, 0, 0, vbSrcCopy
      End If
    Else
      BitBlt UserControl.hdc, 0, 0, picUpOver.Width, picUpOver.Height, picUpOver.hdc, 0, 0, vbSrcCopy
    End If
  End If
  
  UserControl.Refresh
End Sub

Private Sub Timer1_Timer()
  If Not isMouseOver Then
    Timer1.Enabled = False
    bOver = False
    Draw
  End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

  bPressed = True
  Draw

  UserControl_MouseMove Button, Shift, X, y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  On Error Resume Next
  Timer1.Enabled = True
  
  If bPressed Then
    bOver = False
    If X < 0 Or X > UserControl.ScaleWidth Or y < 0 Or y > UserControl.ScaleHeight Then
      bPressed = False
      bUsedto = True
    End If
  Else
    bOver = isMouseOver
  End If
  If bUsedto Then
    If X > 0 And X < UserControl.ScaleWidth And y > 0 And y < UserControl.ScaleHeight Then
      bPressed = True
      bUsedto = False
    End If
  End If
  
  Draw
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  On Error Resume Next
  bPressed = False
  bUsedto = False
  
  If X > 0 And X < UserControl.ScaleWidth And y > 0 And y < UserControl.ScaleHeight Then
    If bExt Then bOn = Not bOn
    PropertyChanged "Value"
    RaiseEvent Pressed(Button)
  End If
  
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Set picUp.Picture = PropBag.ReadProperty("Up", Nothing)
  Set picUpOver.Picture = PropBag.ReadProperty("UpOver", Nothing)
  Set picDown.Picture = PropBag.ReadProperty("Down", Nothing)
  Set picOffUp.Picture = PropBag.ReadProperty("OffUp", Nothing)
  Set picOffUpOver.Picture = PropBag.ReadProperty("OffUpOver", Nothing)
  
  bExt = PropBag.ReadProperty("Extended", False)
  bOn = PropBag.ReadProperty("Value", False)
  UserControl_Resize
  
End Sub

Private Sub UserControl_Resize()
  On Error Resume Next
  On Error Resume Next
  UserControl.Width = picUp.Width * Screen.TwipsPerPixelX
  UserControl.Height = picUp.Height * Screen.TwipsPerPixelY

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Up", picUp.Picture, Nothing)
  Call PropBag.WriteProperty("UpOver", picUpOver.Picture, Nothing)
  Call PropBag.WriteProperty("Down", picDown.Picture, Nothing)
  Call PropBag.WriteProperty("OffUp", picOffUp.Picture, Nothing)
  Call PropBag.WriteProperty("OffUpOver", picOffUpOver.Picture, Nothing)
  Call PropBag.WriteProperty("Extended", bExt, False)
  Call PropBag.WriteProperty("Value", bOn, False)
End Sub

Public Property Get gfxUp() As Picture
  Set gfxUp = picUp.Picture
End Property

Public Property Set gfxUp(ByVal New_Img As Picture)
  Set picUp.Picture = New_Img
  
  UserControl_Resize
  Draw
  PropertyChanged "Up"
End Property

Public Property Get gfxUpOver() As Picture
  Set gfxUpOver = picUpOver.Picture
End Property

Public Property Set gfxUpOver(ByVal New_Img As Picture)
  Set picUpOver.Picture = New_Img
  
  Draw
  PropertyChanged "UpOver"
End Property

Public Property Get gfxDown() As Picture
  Set gfxDown = picDown.Picture
End Property

Public Property Set gfxDown(ByVal New_Img As Picture)
  Set picDown.Picture = New_Img
  
  Draw
  PropertyChanged "Down"
End Property

Public Property Get gfxOffUp() As Picture
  Set gfxOffUp = picOffUp.Picture
End Property

Public Property Set gfxOffUp(ByVal New_Img As Picture)
  Set picOffUp.Picture = New_Img
  
  Draw
  PropertyChanged "OffUp"
End Property

Public Property Get gfxOffUpOver() As Picture
  Set gfxOffUpOver = picOffUpOver.Picture
End Property

Public Property Set gfxOffUpOver(ByVal New_Img As Picture)
  Set picOffUpOver.Picture = New_Img
  
  Draw
  PropertyChanged "OffUpOver"
End Property

Public Property Get Extended() As Boolean
  Extended = bExt
End Property

Public Property Let Extended(ByVal bNew As Boolean)
  bExt = bNew
  PropertyChanged "Extended"
End Property

Public Property Get Value() As Boolean
  Value = bOn
End Property

Public Property Let Value(ByVal bNew As Boolean)
  bOn = bNew
  Draw
  PropertyChanged "Value"
End Property

Public Sub Reset()
  bPressed = False
  bUsedto = False
  bOver = False
  Draw
End Sub

Private Function isMouseOver() As Boolean
  Dim pt As POINTAPI

  GetCursorPos pt
  isMouseOver = (WindowFromPoint(pt.X, pt.y) = hwnd)
End Function
