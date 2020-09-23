VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   Caption         =   "About Simple Amp"
   ClientHeight    =   5625
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4140
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmAbout.frx":08CA
   ScaleHeight     =   5625
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCPU 
      Interval        =   50
      Left            =   120
      Top             =   5040
   End
   Begin VB.Label lblCPU 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   3150
      TabIndex        =   2
      Top             =   5250
      Width           =   855
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   2655
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      UseMnemonic     =   0   'False
      Width           =   3735
   End
   Begin VB.Label lblBuild 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Build"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   555
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   1875
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
  On Error Resume Next

  AlwaysOnTop Me, Settings.OnTop

End Sub

Private Sub Form_Click()
  On Error Resume Next
  Unload Me
End Sub

Private Sub Form_Load()
  On Error Resume Next
  
  lblBuild = "v" & App.Major & "." & App.Minor & " Build " & App.Revision & " Compiled " & FileDateTime(App.Path & "\" & App.EXEName & ".exe")
  
  lblCredits = "Copyright (c) Paul Berlin 2002-2003" & vbCrLf & _
               "Freeware!" & vbCrLf & vbCrLf & _
               "Using FMOD Sound Engine v" & Sound.FMODVersion & vbCrLf & _
               "http://www.fmod.org" & vbCrLf & vbCrLf & _
               "Using ComDlg32 by Roal Zanazzi, ID3v23x.DLL (with modification) by R. Glenn Scott, cMP3Info by AmBra and clsSysTray by Martin Richardson." & vbCrLf & vbCrLf & _
               "Beta testing by Rudi Nilsson" & vbCrLf & "(semi-passive testing at least)"
               
End Sub

Private Sub lblBuild_Click()
  Form_Click
End Sub

Private Sub lblCPU_Click()
  Form_Click
End Sub

Private Sub lblCredits_Click()
  Form_Click
End Sub

Private Sub tmrCPU_Timer()
  On Error Resume Next
  lblCPU = Format(Sound.CPUUsage, "##0.#0") & "%"
End Sub
