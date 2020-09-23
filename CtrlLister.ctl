VERSION 5.00
Begin VB.UserControl CtrlLister 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4275
   ClipControls    =   0   'False
   MaskColor       =   &H00FF00FF&
   OLEDropMode     =   1  'Manual
   PropertyPages   =   "CtrlLister.ctx":0000
   ScaleHeight     =   246
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   285
   Begin VB.PictureBox picBg 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   495
      Left            =   120
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   0
      Top             =   3000
      Width           =   495
      Visible         =   0   'False
   End
End
Attribute VB_Name = "CtrlLister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CtrlLister 1.5 - 04/12 2002
'Created by Paul Berlin
'An Listbox with four columns, sorting abilities and more
'It was made with simple amp in mind, and is not very customizable
'for other programs
Option Compare Text
Option Explicit

'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_RIGHT = &H2
Private Const DT_LEFT = &H0
Private Const DT_NOPREFIX = &H800

Private Const QTHRESH As Long = 7          'Threshhold for switching from QuickSort to Insertion Sort
Private Const MinLong As Long = &H80000000

Public Enum enumSortField
  Sort_AristTitle = 0
  Sort_Album = 1
  Sort_Genre = 2
  Sort_Time = 3
  Sort_FileName = 4
  Sort_PlayDate = 5
  Sort_FileType = 6
  Sort_OriginalOrder = 7
  Sort_PlayTimes = 8
  Sort_SkipTimes = 9
End Enum

Private ColumnWidth(3) As Long

Private lSelColor As Long 'color of selected box

Private lItemHeight As Long 'height of an item
Private lItemsPerPage As Long 'items that can fit per page
Private lPages As Long 'number of pages
Private lCurPage As Long 'the current page
Private lItemCut As Long 'Number of pixels that is not visible of lowest visible item

Public lKeySelected As Long 'The item where focus is
Private lShiftIndex As Long 'The item where shift is pressed

'Used when dragging files
Private lDragBeginIndex As Long
Private lDragIndex As Long

'Button & mouse control
Private bDown As Boolean 'Is m button down?
Private bDrag As Boolean 'Is drag in progress?
Private bShift As Boolean 'Is Shift down?
Private bCtrl As Boolean 'Is Ctrl down?
Private sY As Single     'Holds Y coordinates for doubleklick event

'Events
Event ItemDblClick(Item As Long, num As Long)
Event ItemClick(Item As Long, num As Long)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Event Scroll()

'### FONT
Public Property Get Font() As Font
  Set Font = UserControl.Font
End Property

Public Property Set Font(ByRef newFont As Font)
  Set UserControl.Font = newFont
  UserControl.FontBold = False
  CalcSize
  PropertyChanged "FONT"
  DrawList
End Property

Public Property Get FontSize() As Integer
  FontSize = UserControl.FontSize
End Property

Public Property Let FontSize(ByVal newSize As Integer)
  UserControl.FontSize = newSize
  CalcSize
  PropertyChanged "FONT"
  DrawList
End Property

'#### COLORS
Public Property Get SelectColor() As OLE_COLOR
  SelectColor = lSelColor
End Property

Public Property Let SelectColor(ByVal theCol As OLE_COLOR)
  lSelColor = theCol
  PropertyChanged "SelColor"
  DrawList
End Property

Public Property Get BackColor() As OLE_COLOR
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal theCol As OLE_COLOR)
  UserControl.BackColor = theCol
  PropertyChanged "BgColor"
  DrawList
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_Col As OLE_COLOR)
  UserControl.ForeColor = New_Col
  PropertyChanged "TextColor"
End Property

'###### Pictures
Public Property Get BgPicture() As Picture
  Set BgPicture = picBg.Picture
End Property

Public Property Set BgPicture(ByVal New_Img As Picture)
  Set picBg.Picture = New_Img

  DrawList
  PropertyChanged "BgPicture"
End Property
'###### END OF PROPERTIES

Private Sub DrawList()
  On Error Resume Next
  Dim X As Single, y As Long, c As Long, W As Single
  Dim rec As RECT
  
  UserControl.Cls
  'Draw background
  If picBg.Picture <> 0 Then
    For X = 0 To UserControl.ScaleWidth Step picBg.ScaleWidth
      For y = 0 To UserControl.ScaleHeight Step picBg.ScaleHeight
        BitBlt UserControl.hdc, X, y, picBg.ScaleWidth, picBg.ScaleHeight, picBg.hdc, 0, 0, vbSrcCopy
      Next
    Next
  End If
  'Draw all text
  
  'Dim lstr As String
  
  y = (lCurPage / lItemHeight)
  W = Not ((lCurPage / lItemHeight) + 1 - y) * lItemHeight
  For X = W To UserControl.ScaleHeight Step lItemHeight
    If y > UBound(Playlist) Then Exit For
    c = GetNum(y)
    
    If Playlist(c).Selected Then
      UserControl.Line (3, X)-(UserControl.ScaleWidth - 4, X + lItemHeight), lSelColor, BF
    End If
    If Playlist(c).IsBold Then UserControl.FontBold = True
    
    'lstr = Library(Playlist(c).Reference).sArtistTitle
    'Library(Playlist(c).Reference).sArtistTitle = Library(Playlist(c).Reference).sArtistTitle & " (" & Playlist(c).Index & ", " & Playlist(c).lShuffleIndex & ")"
    
    If Settings.DynamicColumns And Len(Library(Playlist(c).Reference).sAlbum) = 0 And Len(Library(Playlist(c).Reference).sGenre) = 0 Then
      DrawText UserControl.hdc, Library(Playlist(c).Reference).sArtistTitle, Len(Library(Playlist(c).Reference).sArtistTitle), rectMake(3, X, ColumnWidth(0) + ColumnWidth(1) + ColumnWidth(2), lItemHeight * y), DT_LEFT Or DT_NOPREFIX
    ElseIf Settings.DynamicColumns And Len(Library(Playlist(c).Reference).sAlbum) = 0 Then
      DrawText UserControl.hdc, Library(Playlist(c).Reference).sArtistTitle, Len(Library(Playlist(c).Reference).sArtistTitle), rectMake(3, X, ColumnWidth(0) + ColumnWidth(1), lItemHeight * y), DT_LEFT Or DT_NOPREFIX
    Else
      DrawText UserControl.hdc, Library(Playlist(c).Reference).sArtistTitle, Len(Library(Playlist(c).Reference).sArtistTitle), rectMake(3, X, ColumnWidth(0), lItemHeight * y), DT_LEFT Or DT_NOPREFIX
    End If
    If Settings.DynamicColumns And Len(Library(Playlist(c).Reference).sAlbum) > 0 And Len(Library(Playlist(c).Reference).sGenre) = 0 Then
      DrawText UserControl.hdc, Library(Playlist(c).Reference).sAlbum, Len(Library(Playlist(c).Reference).sAlbum), rectMake(ColumnWidth(0) + 5, X, ColumnWidth(0) + ColumnWidth(1) + ColumnWidth(2), lItemHeight * y), DT_LEFT Or DT_NOPREFIX
    Else
      DrawText UserControl.hdc, Library(Playlist(c).Reference).sAlbum, Len(Library(Playlist(c).Reference).sAlbum), rectMake(ColumnWidth(0) + 5, X, ColumnWidth(0) + ColumnWidth(1), lItemHeight * y), DT_LEFT Or DT_NOPREFIX
    End If
    If Len(Library(Playlist(c).Reference).sGenre) > 0 Then
      DrawText UserControl.hdc, Library(Playlist(c).Reference).sGenre, Len(Library(Playlist(c).Reference).sGenre), rectMake(ColumnWidth(0) + ColumnWidth(1) + 5, X, ColumnWidth(0) + ColumnWidth(1) + ColumnWidth(2), lItemHeight * y), DT_LEFT Or DT_NOPREFIX
    End If
    DrawText UserControl.hdc, ConvertTime(Library(Playlist(c).Reference).lLength), Len(ConvertTime(Library(Playlist(c).Reference).lLength)), rectMake(ColumnWidth(0) + ColumnWidth(1) + ColumnWidth(2), X, ColumnWidth(0) + ColumnWidth(1) + ColumnWidth(2) + ColumnWidth(3) - 3, lItemHeight * y), DT_RIGHT Or DT_NOPREFIX
    If Playlist(c).IsBold Then UserControl.FontBold = False
    
    'Library(Playlist(c).Reference).sArtistTitle = lstr

    y = y + 1
  Next X

  UserControl.Refresh
End Sub

Private Sub UserControl_DblClick()
  On Error Resume Next
  Dim d As Long
  d = Int((lCurPage / lItemHeight) + (sY / lItemHeight) + 1)
  RaiseEvent ItemDblClick(d, GetNum(d))
End Sub

Private Sub UserControl_EnterFocus()
  bDown = False
End Sub

Private Sub UserControl_Initialize()
  CalcSize
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  Dim b As Long
  'If bShift Or bCtrl Then Exit Sub
  If Shift = vbShiftMask And Not bShift Then lShiftIndex = lKeySelected
  
  Select Case KeyCode
    Case vbKeyUp 'up key
      lKeySelected = lKeySelected - 1
      If lKeySelected < 1 Then lKeySelected = 1
      SelectNone
      MakeVisible lKeySelected
      Playlist(GetNum(lKeySelected)).Selected = True
      
    Case vbKeyDown 'down key
      lKeySelected = lKeySelected + 1
      If lKeySelected > UBound(Playlist) Then lKeySelected = UBound(Playlist)
      SelectNone
      MakeVisible lKeySelected
      Playlist(GetNum(lKeySelected)).Selected = True
      
    Case vbKeyPageDown 'page down
      lKeySelected = lKeySelected + lItemsPerPage
      If lKeySelected > UBound(Playlist) Then lKeySelected = UBound(Playlist)
      SelectNone
      MakeVisible lKeySelected
      Playlist(GetNum(lKeySelected)).Selected = True
      
    Case vbKeyPageUp 'page up button
      SelectNone
      lKeySelected = lKeySelected - lItemsPerPage
      If lKeySelected < 1 Then lKeySelected = 1
      MakeVisible lKeySelected
      Playlist(GetNum(lKeySelected)).Selected = True
      
    Case vbKeyHome 'home button
      SelectNone
      lKeySelected = 1
      MakeVisible lKeySelected
      Playlist(GetNum(lKeySelected)).Selected = True
     
    Case vbKeyEnd 'End button
      SelectNone
      lKeySelected = UBound(Playlist)
      MakeVisible lKeySelected
      Playlist(GetNum(lKeySelected)).Selected = True
      
    Case Else
      If Shift = vbShiftMask Then
        bShift = True
      ElseIf Shift = vbCtrlMask Then
        bCtrl = True
      End If
      
  End Select
  
  If Shift = vbShiftMask Then
    If lShiftIndex > lKeySelected Then
      For b = lKeySelected To lShiftIndex
        Playlist(GetNum(b)).Selected = True
      Next
    Else
      For b = lShiftIndex To lKeySelected
        Playlist(GetNum(b)).Selected = True
      Next
    End If
  End If
  
  RaiseEvent ItemClick(lKeySelected, GetNum(lKeySelected))
  RaiseEvent KeyDown(KeyCode, Shift)
  DrawList
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  If Shift <> vbShiftMask Then bShift = False
  If Shift <> vbCtrlMask Then bCtrl = False

  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  On Error Resume Next
  bDown = True
  lDragBeginIndex = Fix((lCurPage / lItemHeight) + (y / lItemHeight) + 1)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  On Error Resume Next
  Dim b As Long
  
  sY = y
  If bDown Then
    bDrag = True
    lDragIndex = Fix((lCurPage / lItemHeight) + (y / lItemHeight) + 1)
    If lDragIndex <> lDragBeginIndex Then 'if y coordinates have moved over one lItemHeight
      
      If lDragIndex - lDragBeginIndex = 1 Then 'move down one step
        For b = UBound(Playlist) To 1 Step -1
          If Playlist(b).Selected Then
            If Playlist(b).Index + 1 > UBound(Playlist) Then Exit For
            Playlist(GetNum(Playlist(b).Index + 1)).Index = Playlist(GetNum(Playlist(b).Index + 1)).Index - 1
            Playlist(b).Index = Playlist(b).Index + 1
          End If
        Next b
        For b = UBound(Playlist) To 1 Step -1 'Move view to lowest item
          If Playlist(b).Selected Then
            MakeVisible Playlist(b).Index
            Exit For
          End If
        Next b
        
      ElseIf lDragIndex - lDragBeginIndex = -1 Then 'move up one step
        For b = 1 To UBound(Playlist)
          If Playlist(b).Selected Then
            If Playlist(b).Index - 1 < 1 Then Exit For
            Playlist(GetNum(Playlist(b).Index - 1)).Index = Playlist(GetNum(Playlist(b).Index - 1)).Index + 1
            Playlist(b).Index = Playlist(b).Index - 1
          End If
        Next b
        For b = 1 To UBound(Playlist) 'Move view to highest item
          If Playlist(b).Selected Then
            MakeVisible Playlist(b).Index
            Exit For
          End If
        Next b
      End If
      
      lDragBeginIndex = lDragIndex
    End If
    DrawList
  End If
  
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  On Error Resume Next
  Dim b As Long, c As Long

  If bShift And Not bDrag And lKeySelected > 0 Then 'select with shift
    b = Int((lCurPage / lItemHeight) + (y / lItemHeight) + 1)
    SelectNone
    If b < lKeySelected Then
      For c = b To lKeySelected
        Playlist(GetNum(c)).Selected = True
      Next
    ElseIf b > lKeySelected Then
      For c = lKeySelected To b
        Playlist(GetNum(c)).Selected = True
      Next
    Else
      Playlist(GetNum(lKeySelected)).Selected = True
    End If
  ElseIf Not bDrag Then 'select normal or with ctrl
    If Not bCtrl Then SelectNone
    lKeySelected = Int((lCurPage / lItemHeight) + (y / lItemHeight) + 1)
    If bCtrl Then
      Playlist(GetNum(lKeySelected)).Selected = Not Playlist(GetNum(lKeySelected)).Selected
    Else
      Playlist(GetNum(lKeySelected)).Selected = True
    End If
  End If
  
  bDown = False
  bDrag = False

  If lKeySelected > UBound(Playlist) Or lKeySelected < 1 Then
    lKeySelected = 0
  Else
    RaiseEvent ItemClick(lKeySelected, GetNum(lKeySelected))
  End If
  RaiseEvent MouseUp(Button, Shift, X, y)
  DrawList
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  frmPlaylist.imgDropdown_OLEDragDrop Data, Effect, Button, Shift, X, y
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Set picBg.Picture = PropBag.ReadProperty("BgPicture", Nothing)
  UserControl.BackColor = PropBag.ReadProperty("BgColor", vbWhite)
  UserControl.ForeColor = PropBag.ReadProperty("TextColor", vbBlack)
  lSelColor = PropBag.ReadProperty("SelColor", vbBlue)
  UserControl.Font = PropBag.ReadProperty("FONT", "Ms Sans Serif")

  CalcSize
  DrawList
End Sub

Private Sub UserControl_Resize()
  CalcSize
  DrawList
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("BgPicture", picBg.Picture, Nothing)
  Call PropBag.WriteProperty("BgColor", UserControl.BackColor, vbWhite)
  Call PropBag.WriteProperty("TextColor", UserControl.ForeColor, vbBlack)
  Call PropBag.WriteProperty("SelColor", lSelColor, vbBlue)
  Call PropBag.WriteProperty("FONT", UserControl.Font, "Ms Sans Serif")

End Sub

Public Sub Refresh()
  CalcSize
  DrawList
End Sub

Public Sub CalcSize()
  On Error Resume Next
  lItemHeight = UserControl.TextHeight("I")
  lItemsPerPage = Fix(UserControl.ScaleHeight / lItemHeight)
  lItemCut = (lItemHeight - (((UserControl.ScaleHeight / lItemHeight) - lItemsPerPage) * lItemHeight))
  lPages = ((UBound(Playlist) - lItemsPerPage) * lItemHeight) - lItemHeight + lItemCut
  If lPages < 0 Then lPages = 0
  If lCurPage > lPages Then lCurPage = lPages
  If lCurPage < 0 Then lCurPage = 0
  If UBound(Playlist) > 0 Then
    If lKeySelected > UBound(Playlist) Then
      lKeySelected = UBound(Playlist)
      If lKeySelected > 0 Then
        MakeVisible lKeySelected - lItemsPerPage + 1
        RaiseEvent Scroll
      End If
    ElseIf lKeySelected < 1 Then
      lKeySelected = 1
    End If
  End If
End Sub

Public Sub SetColumnWidth(ByVal Column_Num As Long, ByVal Value As Long)
  If Column_Num < 0 Or Column_Num > 3 Then Exit Sub
  ColumnWidth(Column_Num) = Value
End Sub

Private Function rectMake(ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long) As RECT
  Dim tRet As RECT
  tRet.Bottom = lBottom
  tRet.Top = lTop
  tRet.Left = lLeft
  tRet.Right = lRight
  rectMake = tRet
End Function

Public Property Get Max() As Long
  CalcSize
  Max = lPages
End Property

Public Property Get Value() As Long
  Value = lCurPage
End Property

Public Property Let Value(ByVal vNewValue As Long)
  If vNewValue < 0 Then vNewValue = 0
  If vNewValue > lPages Then vNewValue = lPages
  lCurPage = vNewValue
  DrawList
End Property

Public Sub SelectNone()
  On Error Resume Next
  Dim X As Long
  For X = 1 To UBound(Playlist)
    Playlist(X).Selected = False
  Next
End Sub

Public Sub SelectAll()
  On Error Resume Next
  Dim X As Long
  For X = 1 To UBound(Playlist)
    Playlist(X).Selected = True
  Next
  DrawList
End Sub

Public Sub MakeVisible(ByVal Index As Long)
  'scrolls to the index item in the list
  On Error Resume Next
  Dim l As Single
  
  l = (lCurPage / lItemHeight) + 1

  If Index < l Or Index > l + lItemsPerPage Then
    If Index <= l Then
      lCurPage = (Index - 1) * lItemHeight
    Else
      lCurPage = ((Index - 1 - lItemsPerPage) * lItemHeight) + lItemCut
    End If
    RaiseEvent Scroll
  End If
End Sub

Private Function GetNum(ByVal Index As Long)
  'Gets number in array that has the specified index
  On Error Resume Next
  Dim X As Long
  
  For X = 1 To UBound(Playlist)
    If Playlist(X).Index = Index Then
      GetNum = X
      Exit Function
    End If
  Next
End Function

Public Sub ColumnSort(ByVal Column As enumSortField, Optional ByVal bReverse As Boolean)
  'This will sort the list after the specified column, and reverse it if specified
  On Error Resume Next
  Dim dData() As Date
  Dim sData() As String
  Dim lData() As Long
  Dim Index() As Long
  Dim X As Long
  Dim t As Long
  t = timeGetTime
  
  ReDim Index(1 To UBound(Playlist))

  If Column = Sort_AristTitle Then 'Sort Artist & Title
    
    ReDim sData(1 To UBound(Playlist))
    
    For X = 1 To UBound(Playlist)
      sData(X) = Library(Playlist(X).Reference).sArtistTitle
      Index(X) = Playlist(X).Index
    Next
    
    SortStringIndexArray sData(), Index()
    
  ElseIf Column = Sort_Album Then 'Sort Album
  
    ReDim sData(1 To UBound(Playlist))
    
    For X = 1 To UBound(Playlist)
      sData(X) = Library(Playlist(X).Reference).sAlbum
      Index(X) = Playlist(X).Index
    Next
    
    SortStringIndexArray sData(), Index()
  
  ElseIf Column = Sort_Genre Then 'Sort Genre
  
    ReDim sData(1 To UBound(Playlist))
    
    For X = 1 To UBound(Playlist)
      sData(X) = Library(Playlist(X).Reference).sGenre
      Index(X) = Playlist(X).Index
    Next
    
    SortStringIndexArray sData(), Index()
  
  ElseIf Column = Sort_Time Then 'Sort Time
  
    ReDim lData(1 To UBound(Playlist))
    
    For X = 1 To UBound(Playlist)
      lData(X) = Library(Playlist(X).Reference).lLength
      Index(X) = Playlist(X).Index
    Next
    
    SortLongIndexArray lData(), Index()
  
  ElseIf Column = Sort_FileName Then 'Sort Filename
      
    ReDim sData(1 To UBound(Playlist))
    
    For X = 1 To UBound(Playlist)
      sData(X) = Library(Playlist(X).Reference).sFilename
      Index(X) = Playlist(X).Index
    Next
    
    SortStringIndexArray sData(), Index()
    
  ElseIf Column = Sort_FileType Then 'Sort File type
  
    ReDim lData(1 To UBound(Playlist))
    
    For X = 1 To UBound(Playlist)
      lData(X) = Library(Playlist(X).Reference).eType
      Index(X) = Playlist(X).Index
    Next
    
    SortLongIndexArray lData(), Index()
  
  ElseIf Column = Sort_PlayDate Then 'Sort last play date
    
    ReDim dData(1 To UBound(Playlist))
    
    For X = 1 To UBound(Playlist)
      dData(X) = Library(Playlist(X).Reference).dLastPlayDate
      Index(X) = Playlist(X).Index
    Next
    
    SortDateIndexArray dData(), Index()
    bReverse = Not bReverse

  ElseIf Column = Sort_PlayTimes Then 'Sort after # play times
    
    ReDim lData(1 To UBound(Playlist))
    
    For X = 1 To UBound(Playlist)
      lData(X) = Library(Playlist(X).Reference).lTimesPlayed
      Index(X) = Playlist(X).Index
    Next
    
    SortLongIndexArray lData(), Index()
  
  ElseIf Column = Sort_SkipTimes Then 'Sort after # skip times
    
    ReDim lData(1 To UBound(Playlist))
    
    For X = 1 To UBound(Playlist)
      lData(X) = Library(Playlist(X).Reference).lTimesSkipped
      Index(X) = Playlist(X).Index
    Next
    
    SortLongIndexArray lData(), Index()
  
  ElseIf Column = Sort_OriginalOrder Then 'Sort after original order
  
    For X = 1 To UBound(Playlist)
      Index(X) = X
    Next
  
  End If
  
  For X = 1 To UBound(Playlist)
    Playlist(Index(X)).Index = X
  Next
  If bReverse Then ReverseList
   
  DrawList
  
  Debug.Print timeGetTime - t & " ms"
End Sub

Public Sub ReverseList()
  On Error Resume Next
  Dim X As Long
  
  For X = 1 To UBound(Playlist)
    Playlist(X).Index = UBound(Playlist) - Playlist(X).Index + 1
  Next
End Sub

Public Property Get ItemSelected() As Long
  On Error Resume Next
  If lKeySelected > 0 And lKeySelected <= UBound(Playlist) Then
    ItemSelected = GetNum(lKeySelected)
  Else
    ItemSelected = 1
  End If
End Property

Public Function GetNumFromIndex(ByVal Index As Long)
  GetNumFromIndex = GetNum(Index)
End Function

Private Sub SwapLong(A As Long, b As Long)
    Static c As Long
    c = A
    A = b
    b = c
End Sub

Public Sub SortStringIndexArray(TheArray() As String, TheIndex() As Long, Optional LowerBound As Long = MinLong, Optional UpperBound As Long = MinLong)
    Dim f          As Long
    Dim g          As Long
    Dim H          As Long
    Dim i          As Long
    Dim j          As Long

    Dim s(1 To 64) As Long
    Dim t          As Long

    Dim swp        As String
    Dim indxt      As Long

    If LowerBound = MinLong Then f = LBound(TheIndex) Else f = LowerBound
    If UpperBound = MinLong Then g = UBound(TheIndex) Else g = UpperBound

    t = 0
    Do
        If g - f < QTHRESH Then
            For j = f + 1 To g
                indxt = TheIndex(j)
                CopyMemory ByVal VarPtr(swp), ByVal VarPtr(TheArray(indxt)), 4 ' swp = TheArray(indxt)
                For i = j - 1 To f Step -1
                    If TheArray(TheIndex(i)) <= swp Then Exit For
                    TheIndex(i + 1) = TheIndex(i)
                Next i
                TheIndex(i + 1) = indxt
            Next j
            If t = 0 Then Exit Do
            g = s(t)
            f = s(t - 1)
            t = t - 2
        Else
            H = (f + g) \ 2
            SwapLong TheIndex(H), TheIndex(f + 1)

            If TheArray(TheIndex(f)) > TheArray(TheIndex(g)) Then SwapLong TheIndex(f), TheIndex(g)
            If TheArray(TheIndex(f + 1)) > TheArray(TheIndex(g)) Then SwapLong TheIndex(f + 1), TheIndex(g)
            If TheArray(TheIndex(f)) > TheArray(TheIndex(f + 1)) Then SwapLong TheIndex(f), TheIndex(f + 1)

            i = f + 1
            j = g
            indxt = TheIndex(f + 1)
            CopyMemory ByVal VarPtr(swp), ByVal VarPtr(TheArray(indxt)), 4 ' swp = TheArray(indxt)
            Do
                Do
                  i = i + 1
                Loop While TheArray(TheIndex(i)) < swp
                Do
                    j = j - 1
                Loop While TheArray(TheIndex(j)) > swp
                If j < i Then Exit Do
                SwapLong TheIndex(i), TheIndex(j)
            Loop

            TheIndex(f + 1) = TheIndex(j)
            TheIndex(j) = indxt

            t = t + 2
            If g - i + 1 >= j - f Then
                s(t) = g
                s(t - 1) = i
                g = j - 1
            Else
                s(t) = j - 1
                s(t - 1) = f
                f = i
            End If
        End If
    Loop

    CopyMemory ByVal VarPtr(swp), 0&, 4  'Clear the string pointer.  This is necessary.
                                         'especially if this code is run under Win NT 4.0
End Sub

Public Sub SortLongIndexArray(TheArray() As Long, TheIndex() As Long, Optional LowerBound As Long = MinLong, Optional UpperBound As Long = MinLong)
    Dim f          As Long
    Dim g          As Long
    Dim H          As Long
    Dim i          As Long
    Dim j          As Long

    Dim s(1 To 64) As Long
    Dim t          As Long

    Dim swp        As Long
    Dim indxt      As Long

    If LowerBound = MinLong Then f = LBound(TheIndex) Else f = LowerBound
    If UpperBound = MinLong Then g = UBound(TheIndex) Else g = UpperBound

    t = 0
    Do
        If g - f < QTHRESH Then
            For j = f + 1 To g
                indxt = TheIndex(j)
                swp = TheArray(indxt)
                For i = j - 1 To f Step -1
                    If TheArray(TheIndex(i)) <= swp Then Exit For
                    TheIndex(i + 1) = TheIndex(i)
                Next i
                TheIndex(i + 1) = indxt
            Next j
            If t = 0 Then Exit Do
            g = s(t)
            f = s(t - 1)
            t = t - 2
        Else
            H = (f + g) \ 2
            SwapLong TheIndex(H), TheIndex(f + 1)

            If TheArray(TheIndex(f)) > TheArray(TheIndex(g)) Then SwapLong TheIndex(f), TheIndex(g)
            If TheArray(TheIndex(f + 1)) > TheArray(TheIndex(g)) Then SwapLong TheIndex(f + 1), TheIndex(g)
            If TheArray(TheIndex(f)) > TheArray(TheIndex(f + 1)) Then SwapLong TheIndex(f), TheIndex(f + 1)

            i = f + 1
            j = g
            indxt = TheIndex(f + 1)
            swp = TheArray(indxt)
            Do
                Do
                  i = i + 1
                Loop While TheArray(TheIndex(i)) < swp
                Do
                    j = j - 1
                Loop While TheArray(TheIndex(j)) > swp
                If j < i Then Exit Do
                SwapLong TheIndex(i), TheIndex(j)
            Loop

            TheIndex(f + 1) = TheIndex(j)
            TheIndex(j) = indxt

            t = t + 2
            If g - i + 1 >= j - f Then
                s(t) = g
                s(t - 1) = i
                g = j - 1
            Else
                s(t) = j - 1
                s(t - 1) = f
                f = i
            End If
        End If
    Loop

End Sub

Public Sub SortDateIndexArray(TheArray() As Date, TheIndex() As Long, Optional LowerBound As Long = MinLong, Optional UpperBound As Long = MinLong)
    Dim f          As Long
    Dim g          As Long
    Dim H          As Long
    Dim i          As Long
    Dim j          As Long

    Dim s(1 To 64) As Long
    Dim t          As Long

    Dim swp        As Date
    Dim indxt      As Long

    If LowerBound = MinLong Then f = LBound(TheIndex) Else f = LowerBound
    If UpperBound = MinLong Then g = UBound(TheIndex) Else g = UpperBound

    t = 0
    Do
        If g - f < QTHRESH Then
            For j = f + 1 To g
                indxt = TheIndex(j)
                swp = TheArray(indxt)
                For i = j - 1 To f Step -1
                    If TheArray(TheIndex(i)) <= swp Then Exit For
                    TheIndex(i + 1) = TheIndex(i)
                Next i
                TheIndex(i + 1) = indxt
            Next j
            If t = 0 Then Exit Do
            g = s(t)
            f = s(t - 1)
            t = t - 2
        Else
            H = (f + g) \ 2
            SwapLong TheIndex(H), TheIndex(f + 1)

            If TheArray(TheIndex(f)) > TheArray(TheIndex(g)) Then SwapLong TheIndex(f), TheIndex(g)
            If TheArray(TheIndex(f + 1)) > TheArray(TheIndex(g)) Then SwapLong TheIndex(f + 1), TheIndex(g)
            If TheArray(TheIndex(f)) > TheArray(TheIndex(f + 1)) Then SwapLong TheIndex(f), TheIndex(f + 1)

            i = f + 1
            j = g
            indxt = TheIndex(f + 1)
            swp = TheArray(indxt)
            Do
                Do
                  i = i + 1
                Loop While TheArray(TheIndex(i)) < swp
                Do
                    j = j - 1
                Loop While TheArray(TheIndex(j)) > swp
                If j < i Then Exit Do
                SwapLong TheIndex(i), TheIndex(j)
            Loop

            TheIndex(f + 1) = TheIndex(j)
            TheIndex(j) = indxt

            t = t + 2
            If g - i + 1 >= j - f Then
                s(t) = g
                s(t - 1) = i
                g = j - 1
            Else
                s(t) = j - 1
                s(t - 1) = f
                f = i
            End If
        End If
    Loop

End Sub

Private Function Sine(ByVal Degrees_Arg As Single) As Single
  Sine = Sin(Degrees_Arg * Atn(1) / 45)
End Function
