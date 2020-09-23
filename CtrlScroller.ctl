VERSION 5.00
Begin VB.UserControl CtrlScroller 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   ClientHeight    =   3690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4275
   ClipControls    =   0   'False
   MaskColor       =   &H00FF00FF&
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   246
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   285
   Begin VB.PictureBox picScrBefore 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
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
   Begin VB.PictureBox picScrAfter 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
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
   Begin VB.PictureBox picBarOver 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
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
   Begin VB.PictureBox picBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
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
Attribute VB_Name = "CtrlScroller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CtrlScroller 1.1
'Created by Paul Berlin 2002
'Vertical & Horizontal Scroller

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32" (ByVal hDCDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hDCSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long

' Declarations
Dim iMin As Long
Dim iMax As Long
Dim iValue As Long
Dim ix As Long
Dim iy As Long

Dim IsOn As Boolean
Dim bVertical As Boolean
Dim bDrag As Boolean
Dim bSnap As Boolean

' Events
Event Change()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)

Private Sub DrawBar()
  On Error Resume Next
  UserControl_Resize
  UserControl.Cls
    
  BitBlt UserControl.hdc, 0, 0, picScrBefore.Width, picScrBefore.Height, picScrBefore.hdc, 0, 0, vbSrcCopy
    
  If IsOn Then
    If bVertical Then
      BitBlt UserControl.hdc, 0, iy + (picBar.Height / 2), picScrAfter.Width, picScrAfter.Height - (iy + (picBar.Height / 2)), picScrAfter.hdc, 0, iy + (picBar.Height / 2), vbSrcCopy
    Else
      BitBlt UserControl.hdc, ix + (picBar.Width / 2), 0, picScrAfter.Width - (ix + (picBar.Width / 2)), picScrAfter.Height, picScrAfter.hdc, ix + (picBar.Width / 2), 0, vbSrcCopy
    End If
    If bDrag Then
      TransparentBlt UserControl.hdc, ix, iy, picBarOver.Width, picBarOver.Height, picBarOver.hdc, 0, 0, picBarOver.Width, picBarOver.Height, &HFF00FF
    Else
      TransparentBlt UserControl.hdc, ix, iy, picBar.Width, picBar.Height, picBar.hdc, 0, 0, picBar.Width, picBar.Height, &HFF00FF
    End If
  End If

  UserControl.Refresh
End Sub

Public Property Get Max() As Long
  Max = iMax
End Property

Public Property Let Max(New_Max As Long)
  If iValue > New_Max Then
    iValue = New_Max
  End If
    
  iMax = New_Max
  DrawBar
    
  PropertyChanged "Max"
End Property

Public Property Get Min() As Long
  Min = iMin
End Property

Public Property Let Min(New_Min As Long)
  If New_Min > iValue Then
    New_Min = iValue
  End If
    
  iMin = New_Min
  DrawBar
    
  PropertyChanged "Min"
End Property

Public Property Get Value() As Long
  Value = iValue
End Property

Public Property Let Value(New_Value As Long)
  On Error Resume Next
  If New_Value < iMin Then
    New_Value = iMin
  ElseIf New_Value > iMax Then
    New_Value = iMax
  End If
  
  iValue = New_Value
  If bVertical Then
    iy = (iValue - iMin) / (iMax - iMin) * (UserControl.ScaleHeight - picBar.Height)
  Else
    ix = (iValue - iMin) / (iMax - iMin) * (UserControl.ScaleWidth - picBar.Width)
  End If
  DrawBar
    
  PropertyChanged "Value"
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  bDrag = True
  UserControl_MouseMove Button, Shift, X, y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  On Error Resume Next
  
  If bDrag And IsOn Then
    If bVertical Then
      iy = y

      If iy > picScrBefore.Height - (picBar.Height / 2) Then iy = picScrBefore.Height - (picBar.Height / 2)
      If iy < picBar.Height / 2 Then iy = picBar.Height / 2

      iy = iy - picBar.Height / 2
    
      iValue = iy / (picScrBefore.Height - picBar.Height) * (iMax - iMin) + iMin
      If bSnap Then iy = (iValue - iMin) / (iMax - iMin) * (UserControl.ScaleHeight - picBar.Height)
    Else
      ix = X

      If ix > picScrBefore.Width - (picBar.Width / 2) Then ix = picScrBefore.Width - (picBar.Width / 2)
      If ix < picBar.Width / 2 Then ix = picBar.Width / 2

      ix = ix - picBar.Width / 2
    
      iValue = ix / (picScrBefore.Width - picBar.Width) * (iMax - iMin) + iMin
      If bSnap Then ix = (iValue - iMin) / (iMax - iMin) * (UserControl.ScaleWidth - picBar.Width)
    End If
    DrawBar
    RaiseEvent Change
  End If
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  bDrag = False
  DrawBar
End Sub

Private Sub UserControl_Initialize()
  If iMax = 0 Then iMax = 100
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Set picScrBefore.Picture = PropBag.ReadProperty("ScrBefore", Nothing)
  Set picScrAfter.Picture = PropBag.ReadProperty("ScrollAfter", Nothing)
  Set picBar.Picture = PropBag.ReadProperty("Bar", Nothing)
  Set picBarOver.Picture = PropBag.ReadProperty("BarOver", Nothing)
  bVertical = PropBag.ReadProperty("Vertical", False)
  bSnap = PropBag.ReadProperty("Snap", False)
  IsOn = PropBag.ReadProperty("Enabled", True)
  iMin = PropBag.ReadProperty("Min", 0)
  iMax = PropBag.ReadProperty("Max", 100)
  iValue = PropBag.ReadProperty("Value", 0)
  
  DrawBar
End Sub

Private Sub UserControl_Resize()
  On Error Resume Next
  UserControl.Width = picScrBefore.Width * Screen.TwipsPerPixelX
  UserControl.Height = picScrBefore.Height * Screen.TwipsPerPixelY
  If bVertical Then
    iy = (iValue - iMin) / (iMax - iMin) * (UserControl.ScaleHeight - picBar.Height)
  Else
    ix = (iValue - iMin) / (iMax - iMin) * (UserControl.ScaleWidth - picBar.Width)
  End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("ScrBefore", picScrBefore.Picture, Nothing)
  Call PropBag.WriteProperty("ScrollAfter", picScrAfter.Picture, Nothing)
  Call PropBag.WriteProperty("Bar", picBar.Picture, Nothing)
  Call PropBag.WriteProperty("BarOver", picBarOver.Picture, Nothing)
  Call PropBag.WriteProperty("Enabled", IsOn, True)
  Call PropBag.WriteProperty("Vertical", bVertical, False)
  Call PropBag.WriteProperty("Snap", bSnap, False)
  Call PropBag.WriteProperty("Min", iMin, 0)
  Call PropBag.WriteProperty("Max", iMax, 100)
  Call PropBag.WriteProperty("Value", iValue, 0)
End Sub

Public Property Get Bar() As Picture
  Set Bar = picBar.Picture
End Property

Public Property Set Bar(ByVal New_Img As Picture)
  Set picBar.Picture = New_Img
  
  Call DrawBar
  PropertyChanged "Bar"
End Property

Public Property Get BarOver() As Picture
  Set BarOver = picBarOver.Picture
End Property

Public Property Set BarOver(ByVal New_Img As Picture)
  Set picBarOver.Picture = New_Img
    
  DrawBar
  PropertyChanged "BarOver"
End Property

Public Property Get ScrollBefore() As Picture
  Set ScrollBefore = picScrBefore.Picture
End Property

Public Property Set ScrollBefore(ByVal New_Img As Picture)
  Set picScrBefore.Picture = New_Img
    
  DrawBar
  PropertyChanged "ScrBefore"
End Property

Public Property Get ScrollAfter() As Picture
  Set ScrollAfter = picScrAfter.Picture
End Property

Public Property Set ScrollAfter(ByVal New_Img As Picture)
  Set picScrAfter.Picture = New_Img
    
  DrawBar
  PropertyChanged "ScrollAfter"
End Property

Public Property Get Enabled() As Boolean
  Enabled = IsOn
End Property

Public Property Let Enabled(ByVal bNew As Boolean)
  IsOn = bNew
  PropertyChanged "IsOn"
End Property

Public Property Get Vertical() As Boolean
  Vertical = bVertical
End Property

Public Property Let Vertical(ByVal bNew As Boolean)
  bVertical = bNew
  PropertyChanged "Vertical"
End Property

Public Property Get Snap() As Boolean
  Snap = bSnap
End Property

Public Property Let Snap(ByVal vNewValue As Boolean)
  bSnap = vNewValue
  PropertyChanged "Snap"
End Property
