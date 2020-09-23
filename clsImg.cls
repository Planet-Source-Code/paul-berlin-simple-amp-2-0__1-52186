VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Created by Paul Berlin 2003 - Latest update 22/07
'Used when drawing certain effects of the visualizations, like
'fade & zoom/rotation.

Option Explicit

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32" (ByVal hDCDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hDCSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long

'Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public hdc As Long
Public lWidth As Long
Public lHeight As Long

Private hBM As Long
Private hOldBM As Long

Private bSA As Boolean
Private ba() As Byte, ba2() As Byte
Private bmp As BITMAP
Private sa As SAFEARRAY2D, na As SAFEARRAY2D, na2 As SAFEARRAY2D

Public Sub Create(ByVal lW As Long, ByVal lH As Long, ByVal hSourceDC As Long)
  'hSourceDC can be just the dc of any form, like Me.hdc
  'At the moment safearray does not work with Create, this is no
  'issue for Simple Amp so I don't care =) try to fix it yoursef if you want it
  Destroy
  
  lWidth = lW
  lHeight = lH
  
  hdc = CreateCompatibleDC(0)                               'Create DC
  hBM = CreateCompatibleBitmap(hSourceDC, lWidth, lHeight)  'Create bitmap width same dimensions as image
  hOldBM = SelectObject(hdc, hBM)                           'Select bitmap into DC
  
  'GetObjectAPI hBM, Len(bmp), bmp 'get bitmap info for SAFEARRAY
  
End Sub

Public Sub ClonePicture(ByRef Pic As StdPicture)
  'Works just as well with IPictureDisp, just change the var type
  Dim hPDC As Long, hOldObj As Long
  
  Destroy
  
  hPDC = CreateCompatibleDC(0)      'Create DC for image
  hOldObj = SelectObject(hPDC, Pic) 'Select image into DC
  
  GetObjectAPI Pic, Len(bmp), bmp   'Get bitmap info from image
  lWidth = bmp.bmWidth              'Get width
  lHeight = bmp.bmHeight            'Get height
  
  hdc = CreateCompatibleDC(0)                         'Create DC
  hBM = CreateCompatibleBitmap(hPDC, lWidth, lHeight) 'Create bitmap width same dimensions as image
  hOldBM = SelectObject(hdc, hBM)                     'Select bitmap into DC
  
  BitBlt hdc, 0, 0, lWidth, lHeight, hPDC, 0, 0, vbSrcCopy 'BitBlt image into DC
  
  GetObjectAPI hBM, Len(bmp), bmp 'get bitmap info for SAFEARRAY
  
  SelectObject hPDC, hOldObj  'Deselect image
  DeleteDC hPDC               'Delete DC containing image
  Set Pic = Nothing           'Unload image

End Sub

Public Sub LoadImg(ByVal sFilename As String)
  Dim Pic As StdPicture, hPDC As Long, hOldObj As Long
  
  Destroy
  
  Set Pic = LoadPicture(sFilename)  'Load image
  hPDC = CreateCompatibleDC(0)      'Create DC for image
  hOldObj = SelectObject(hPDC, Pic) 'Select image into DC
  
  GetObjectAPI Pic, Len(bmp), bmp   'Get bitmap info from image
  lWidth = bmp.bmWidth              'Get width
  lHeight = bmp.bmHeight            'Get height
  
  hdc = CreateCompatibleDC(0)                         'Create DC
  hBM = CreateCompatibleBitmap(hPDC, lWidth, lHeight) 'Create bitmap width same dimensions as image
  hOldBM = SelectObject(hdc, hBM)                     'Select bitmap into DC
  
  BitBlt hdc, 0, 0, lWidth, lHeight, hPDC, 0, 0, vbSrcCopy 'BitBlt image into DC
  
  GetObjectAPI hBM, Len(bmp), bmp 'get bitmap info for SAFEARRAY
   
  SelectObject hPDC, hOldObj  'Deselect image
  DeleteDC hPDC               'Delete DC containing image
  Set Pic = Nothing           'Unload image

End Sub

Private Sub Class_Terminate()
  Destroy
End Sub

Public Sub PaintTo(ByVal toDC As Long, ByVal atX As Long, ByVal atY As Long, ByVal dwRop As Long)
  BitBlt toDC, atX, atY, lWidth, lHeight, hdc, 0, 0, dwRop
End Sub

Public Sub PaintToTransCrop(ByVal toDC As Long, ByVal atX As Long, ByVal atY As Long, ByVal atWidth As Long)
  If atWidth > lWidth Then atWidth = lWidth
  TransparentBlt toDC, atX, atY, atWidth, lHeight, hdc, 0, 0, atWidth, lHeight, vbMagenta
End Sub

Public Sub PaintToTransZoom(ByVal toDC As Long, ByVal centerX As Long, ByVal centerY As Long, ByVal Ratio As Single)
  TransparentBlt toDC, centerX - (lWidth * Ratio) / 2, centerY - (lHeight * Ratio) / 2, lWidth * Ratio, lHeight * Ratio, hdc, 0, 0, lWidth, lHeight, vbMagenta
End Sub

Public Sub Destroy()
  If hdc <> 0 Then
    If bSA Then CopyMemory ByVal VarPtrArray(ba), 0&, 4
    bSA = False
    
    SelectObject hdc, hOldBM  'Select old bitmap
    DeleteObject hBM          'Delete bitmap
    DeleteDC hdc              'Delete DC
  End If
End Sub

Public Sub InitSA()
  bSA = True
  With sa
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = bmp.bmHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = bmp.bmWidthBytes
    .pvData = bmp.bmBits
  End With
  CopyMemory ByVal VarPtrArray(ba), VarPtr(sa), 4
  
End Sub

Sub RotateZoomTo(ByRef rotImg As clsImg, ByRef destImg As clsImg, ByVal angle As Double, ByVal zoom As Double)
  'UPDATE: 19-08-03
  'changed asin & acos from double to single = +1 fps!, no need for double precision
  'changed all divisions from / to integer division \ = +2-3 fps
  'Note: these fps changes are when fps where ~280 normally (with fade speed 5 on).
  '      but not that high fps's inside simple amp, these where measured
  '      in an standalone project running just an rotating image, zooming in & out.
  
  Dim asin As Single, acos As Single
  Dim mx As Integer, my As Integer 'center pos of source image
  Dim cx As Integer, cy As Integer 'for...to vars
  Dim sx As Integer, sY As Integer 'calcuated source pos
  Dim c1 As Byte, c2 As Byte, c3 As Byte 'read color values
  Dim d As Integer ', d2 As Integer 'dest image pos upper left
  Dim X As Integer, y As Integer 'dest image center
  Dim hmin As Integer, hmax As Integer
  Dim wmin As Integer, wmax As Integer
  Dim ax As Integer
  
  Dim wbytes As Long, b() As Byte
  'rotImg.SetColor 255, 0, 255
  rotImg.Map b()

  X = rotImg.lWidth \ 2
  y = rotImg.lHeight \ 2
  wbytes = rotImg.lWidth * 3

  hmax = rotImg.lHeight - 1
  wmax = wbytes

  mx = bmp.bmWidth \ 2
  my = bmp.bmHeight \ 2
    
  'Precalculate sin & cos
  asin = Sin(angle)
  acos = Cos(angle)
  
  ax = wmin
  For cx = wmin * 3 To wmax - 2 Step 3
    ax = ax + 1
    For cy = hmin To hmax
      sx = mx + (ax - X) * asin + (cy - y) * acos
      sY = my + (cy - y) * asin - (ax - X) * acos
      If (sx < bmp.bmWidth And sx >= 0) Then
        If (sY < bmp.bmHeight And sY >= 0) Then
          sx = sx * 3
          c1 = ba(sx, sY)
          c2 = ba(sx + 1, sY)
          c3 = ba(sx + 2, sY)
          'If (c1 <> 255 And c2 <> 0 And c3 <> 255) Then
            b(cx, cy) = c1
            b(cx + 1, cy) = c2
            b(cx + 2, cy) = c3
          'End If
        End If
      End If
    Next
  Next
  
  rotImg.UnMap b()
  rotImg.PaintToTransZoom destImg.hdc, destImg.lWidth \ 2, destImg.lHeight \ 2, zoom
    
End Sub

Public Sub Map(ByRef b() As Byte)
  With sa
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = bmp.bmHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = bmp.bmWidthBytes
    .pvData = bmp.bmBits
  End With
  CopyMemory ByVal VarPtrArray(b), VarPtr(sa), 4
End Sub

Public Sub UnMap(ByRef b() As Byte)
  CopyMemory ByVal VarPtrArray(b), 0&, 4
End Sub

Public Sub SetColor(ByVal bR As Byte, ByVal bG As Byte, ByVal bB As Byte)
  Dim cx As Long, cy As Long
  For cy = 0 To lHeight - 1
    For cx = 0 To bmp.bmWidthBytes - 2 Step 3
      ba(cx, cy) = bR
      ba(cx + 1, cy) = bG
      ba(cx + 2, cy) = bB
    Next
  Next
End Sub

Public Sub Fade(ByVal speed As Integer)
  Dim X As Long
  
  For X = 0 To UBound(ba)
    If ba(X) <> ba2(X) Then
      If ba(X) > ba2(X) + speed Then
        ba(X) = ba(X) - speed
      ElseIf ba(X) < ba2(X) - speed Then
        ba(X) = ba(X) + speed
      Else
        ba(X) = ba2(X)
      End If
    End If
  Next

End Sub

Public Function GetBits() As Long
  GetBits = bmp.bmBits
End Function

Public Function GetWidthBytes() As Long
  GetWidthBytes = bmp.bmWidthBytes
End Function

Public Sub InitFade(ByRef origImg As clsImg)
  With na2
    .cbElements = 1
    .cDims = 1
    .Bounds(0).cElements = origImg.lHeight * origImg.GetWidthBytes
    .Bounds(0).lLbound = 0
    .pvData = origImg.GetBits
  End With
  CopyMemory ByVal VarPtrArray(ba2), VarPtr(na2), 4
  
  With na
    .cbElements = 1
    .cDims = 1
    .Bounds(0).cElements = bmp.bmHeight * bmp.bmWidthBytes
    .Bounds(0).lLbound = 0
    .pvData = bmp.bmBits
  End With
  CopyMemory ByVal VarPtrArray(ba), VarPtr(na), 4
End Sub

Public Sub BitBltFrom(ByVal fromDC As Long)
  BitBlt hdc, 0, 0, lWidth, lHeight, fromDC, 0, 0, vbSrcCopy
End Sub
