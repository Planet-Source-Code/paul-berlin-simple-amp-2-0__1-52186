VERSION 5.00
Begin VB.Form frmColor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Color Selector"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Color Selector (click)"
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   2775
      Begin VB.Label lblColor 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Enter from Hex"
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   2775
      Begin VB.TextBox txtHTML 
         Height          =   315
         Left            =   360
         MaxLength       =   6
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "#"
         Height          =   195
         Left            =   195
         TabIndex        =   10
         Top             =   360
         Width           =   105
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(Ex: FFCC00)"
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   360
         Width           =   930
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Enter from RGB"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   2775
      Begin VB.TextBox txtRGB 
         Height          =   315
         Index           =   0
         Left            =   360
         MaxLength       =   3
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtRGB 
         Height          =   315
         Index           =   1
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtRGB 
         Height          =   315
         Index           =   2
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   390
         Width           =   165
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Left            =   960
         TabIndex        =   6
         Top             =   390
         Width           =   165
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Left            =   1800
         TabIndex        =   5
         Top             =   390
         Width           =   150
      End
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
  On Error Resume Next
  AlwaysOnTop Me, Settings.OnTop
  lblColor.BackColor = Val(Me.Tag)
  lblColor.ForeColor = RGB(255 - (lblColor.BackColor Mod &H100), 255 - ((lblColor.BackColor \ &H100) Mod &H100), 255 - ((lblColor.BackColor \ &H10000) Mod &H100))
  lblColor.Caption = lblColor.BackColor
  txtRGB(0) = lblColor.BackColor Mod &H100
  txtRGB(1) = (lblColor.BackColor \ &H100) Mod &H100
  txtRGB(2) = (lblColor.BackColor \ &H10000) Mod &H100
  txtHTML = String(2 - Len(Hex(txtRGB(0))), "0") & Hex(txtRGB(0))
  txtHTML = txtHTML & String(2 - Len(Hex(txtRGB(1))), "0") & Hex(txtRGB(1))
  txtHTML = txtHTML & String(2 - Len(Hex(txtRGB(2))), "0") & Hex(txtRGB(2))
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  'the color is in Me.Tag so we cant unload until the parent
  'window has read the tag. To close the window the parent window
  'then sets the tag to 'Y' and unloads window.
  'it would probably simpler to just use a global var... never mind. =)
  If Me.Tag <> "Y" Then
    Me.Tag = lblColor.BackColor
    Cancel = 1
    Me.Hide
  End If
End Sub

Private Sub lblColor_Click()
  On Error GoTo errh
  Dim cdg As New clsCommonDialog
  With cdg
    Set .Parent = Me
    .CancelError = True
    .Color = lblColor.BackColor
    .Flags = cdlCCRGBInit Or cdlCCFullOpen
    .ShowColor
    lblColor.BackColor = .Color
    lblColor.ForeColor = RGB(255 - (lblColor.BackColor Mod &H100), 255 - ((lblColor.BackColor \ &H100) Mod &H100), 255 - ((lblColor.BackColor \ &H10000) Mod &H100))
    lblColor.Caption = lblColor.BackColor
    txtRGB(0) = lblColor.BackColor Mod &H100
    txtRGB(1) = (lblColor.BackColor \ &H100) Mod &H100
    txtRGB(2) = (lblColor.BackColor \ &H10000) Mod &H100
    txtHTML = String(2 - Len(Hex(txtRGB(0))), "0") & Hex(txtRGB(0))
    txtHTML = txtHTML & String(2 - Len(Hex(txtRGB(1))), "0") & Hex(txtRGB(1))
    txtHTML = txtHTML & String(2 - Len(Hex(txtRGB(2))), "0") & Hex(txtRGB(2))
  End With
errh:
End Sub

Private Sub txtHTML_KeyPress(KeyAscii As Integer)
  On Error Resume Next
  KeyAscii = ctlKeyPress(KeyAscii, Uppercase Or NoSpaces Or NoSingleQuotes Or NoDoubleQuotes)
End Sub

Private Sub txtHTML_KeyUp(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  Dim sTmp As String
  sTmp = txtHTML
  If Len(sTmp) < 6 Then
    sTmp = sTmp & String(6 - Len(sTmp), "0")
  End If
  lblColor.BackColor = "&H00" & Right(txtHTML, 2) & Mid(txtHTML, 3, 2) & Left(txtHTML, 2)
  lblColor.ForeColor = RGB(255 - (lblColor.BackColor Mod &H100), 255 - ((lblColor.BackColor \ &H100) Mod &H100), 255 - ((lblColor.BackColor \ &H10000) Mod &H100))
  lblColor.Caption = lblColor.BackColor
  txtRGB(0) = lblColor.BackColor Mod &H100
  txtRGB(1) = (lblColor.BackColor \ &H100) Mod &H100
  txtRGB(2) = (lblColor.BackColor \ &H10000) Mod &H100
End Sub

Private Sub txtRGB_KeyPress(Index As Integer, KeyAscii As Integer)
  On Error Resume Next
  KeyAscii = ctlKeyPress(KeyAscii, NumbersOnly)
End Sub

Private Sub txtRGB_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  lblColor.BackColor = RGB(Val(txtRGB(0)), Val(txtRGB(1)), Val(txtRGB(2)))
  lblColor.ForeColor = RGB(255 - (lblColor.BackColor Mod &H100), 255 - ((lblColor.BackColor \ &H100) Mod &H100), 255 - ((lblColor.BackColor \ &H10000) Mod &H100))
  lblColor.Caption = lblColor.BackColor
  txtHTML = String(2 - Len(Hex(txtRGB(0))), "0") & Hex(txtRGB(0))
  txtHTML = txtHTML & String(2 - Len(Hex(txtRGB(1))), "0") & Hex(txtRGB(1))
  txtHTML = txtHTML & String(2 - Len(Hex(txtRGB(2))), "0") & Hex(txtRGB(2))
End Sub
