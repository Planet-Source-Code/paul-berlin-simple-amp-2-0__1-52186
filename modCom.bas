Attribute VB_Name = "modCom"
'Written by Paul Berlin 2002
'Used to communicate between two programs
'in this case, between the already running instance of Simple Amp and
'the newly started one.

Option Explicit

Private Type COPYDATASTRUCT
  dwData As Long
  cbData As Long
  lpData As Long
End Type

Private Const GWL_WNDPROC = (-4)
Private Const WM_COPYDATA = &H4A
Private lpPrevWndProc As Long

'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lngParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Sub SendString(ByVal sWinCaption As String, ByVal sData As String)
  Dim cdCopyData As COPYDATASTRUCT
  Dim ThWnd As Long
  Dim byteBuffer(1 To 255) As Byte
    
  'Get the hWnd of the target application
  ThWnd = FindWindow(vbNullString, sWinCaption)
  
  'Copy the string into a byte array, converting it to ASCII
  CopyMemory byteBuffer(1), ByVal sData, Len(sData)
  cdCopyData.dwData = 3
  cdCopyData.cbData = Len(sData) + 1
  cdCopyData.lpData = VarPtr(byteBuffer(1))
  Call SendMessage(ThWnd, WM_COPYDATA, frmMain.hwnd, cdCopyData)

End Sub

Public Sub Hook(ByVal hwnd As Long)
  'Grabs an window for viewing its messages
  lpPrevWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook(ByVal hwnd As Long)
  'Ungrab
  Call SetWindowLong(hwnd, GWL_WNDPROC, lpPrevWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lngParam As Long) As Long
  'This function intercepts the window messages
  If uMsg = WM_COPYDATA Then
    Call InterProcessComms(lngParam)
  End If
  WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lngParam)
End Function

Sub InterProcessComms(ByVal lngParam As Long)
  Dim cdCopyData As COPYDATASTRUCT
  Dim byteBuffer(1 To 255) As Byte
  Dim strTemp As String
          
  Call CopyMemory(cdCopyData, ByVal lngParam, Len(cdCopyData))
  
  Select Case cdCopyData.dwData
    Case 3
      Call CopyMemory(byteBuffer(1), ByVal cdCopyData.lpData, cdCopyData.cbData)
      strTemp = StrConv(byteBuffer, vbUnicode)
      strTemp = Left$(strTemp, InStr(1, strTemp, Chr(0)) - 1) 'remove null chars
      frmDDE.CheckCommand strTemp
  End Select
  
End Sub
