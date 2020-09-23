VERSION 5.00
Begin VB.Form frmDDE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Amp 2.0 RC 010"
   ClientHeight    =   420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleWidth      =   2970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmDDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This is the form used for communicating with other instances of this program
'The window caption is used with FindWindow
'The name frmDDE is left from the times Simple Amp used DDE...

Private Sub Form_Load()
  Hook Me.hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unhook Me.hwnd
End Sub

Public Sub CheckCommand(ByVal CmdStr As String)
  On Error Resume Next
  
  Select Case Left(CmdStr, 1)
    Case "0"
      frmMain.PlayPause
    Case "1"
      frmMain.PlayPrev
    Case "2"
      frmMain.PlayNext
    Case "3"
      If Settings.Shuffle > 0 Then
        Settings.Shuffle = 0
      Else
        Settings.Shuffle = 1
      End If
      frmMain.ShuffleOnOff
    Case "4"
      If Settings.Repeat > 0 Then
        Settings.Repeat = 0
      Else
        Settings.Repeat = 1
      End If
      frmMain.RepeatOnOff
    Case "5"
      If frmMain.Visible Then
        frmMain.Minimize
      Else
        frmMain.cTray_LButtonDblClk
      End If
    Case "6"
      SimpleAddFile Right(CmdStr, Len(CmdStr) - 1)
    Case "7"
      frmMain.PlayStop
      frmMain.Play
    Case "8"
      frmMenus.menFwd_Click
    Case "9"
      frmMenus.menBck_Click
    Case "a"
      devSurround = Not devSurround
      Sound.StreamSurround = devSurround
    Case "b"
      If Not frmMain.Visible Then
        frmMain.cTray_LButtonDblClk
      End If
      frmAdd.Show
  End Select
End Sub
