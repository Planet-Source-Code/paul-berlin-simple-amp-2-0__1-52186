VERSION 5.00
Begin VB.Form frmLibSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Media Library Settings"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmLibSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkLibOpt 
      Caption         =   "&Hide Modules, Midis and Wave-files in the Media Library."
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   4335
   End
   Begin VB.CommandButton cmdClean 
      Caption         =   "Library Maintenace"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CheckBox chkLibOpt 
      Caption         =   "&Autosize the list columns to fit their content."
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   3375
   End
   Begin VB.CheckBox chkLibOpt 
      Caption         =   "&When there are no tags, get artist && title from filename."
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLibSettings.frx":058A
      Height          =   585
      Left            =   360
      TabIndex        =   1
      Top             =   405
      Width           =   4215
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmLibSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim b As Byte, c As Byte, d As Byte

Private Sub cmdClean_Click()
  On Error Resume Next
  frmLibClean.Show , frmLibrary
  c = True
End Sub

Private Sub Command1_Click()
  On Error Resume Next
  b = Abs(Settings.LibGetName)
  d = Abs(Settings.LibHideNonMusic)
  Settings.LibGetName = CBool(chkLibOpt(0).Value)
  Settings.LibAutosize = CBool(chkLibOpt(1).Value)
  Settings.LibHideNonMusic = CBool(chkLibOpt(2).Value)
  If (b <> Abs(Settings.LibGetName) Or d <> Abs(Settings.LibHideNonMusic) Or c) Then
    frmLibrary.LoadLib
    frmLibrary.Examine
  End If
  Unload Me
End Sub

Private Sub Form_Activate()
  On Error Resume Next
  
  AlwaysOnTop Me, Settings.OnTop
End Sub

Private Sub Form_Load()
  On Error Resume Next
  chkLibOpt(0).Value = Abs(Settings.LibGetName)
  chkLibOpt(1).Value = Abs(Settings.LibAutosize)
  chkLibOpt(2).Value = Abs(Settings.LibHideNonMusic)
End Sub
