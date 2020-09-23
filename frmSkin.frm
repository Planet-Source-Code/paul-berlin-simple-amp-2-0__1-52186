VERSION 5.00
Begin VB.Form frmSkin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Skin"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7470
   ControlBox      =   0   'False
   Icon            =   "frmSkin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Frame frmInfo 
      Caption         =   "Skin Info"
      Height          =   6495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5055
      Begin VB.TextBox txtNotes 
         Height          =   1335
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   4920
         Width           =   4815
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Notes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   4680
         Width           =   570
      End
      Begin VB.Label lblVersion 
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   4440
         Width           =   3855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Version:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   4440
         Width           =   705
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4695
      End
      Begin VB.Image imgPreview 
         BorderStyle     =   1  'Fixed Single
         Height          =   3600
         Left            =   120
         Stretch         =   -1  'True
         Top             =   480
         Width           =   4740
      End
      Begin VB.Label lblAuthor 
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   4200
         Width           =   3975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Author:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   4200
         Width           =   630
      End
   End
   Begin VB.ListBox lstSkins 
      Height          =   3765
      Left            =   5280
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Available skins:"
      Height          =   195
      Left            =   5280
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Skin
  name As String
  Author As String
  Version As String
  Notes As String
  Preview As Picture
  fname As String
End Type

Private Skins() As Skin 'info for each skin found

Private Sub cmdLoad_Click()
  'Loads the selected skin
  On Error Resume Next
  CurrentSkin = Skins(lstSkins.ListIndex + 1).fname
  'Tur off docked status
  Docked = False
  DockedLeft = 0
  DockedTop = 0
  frmMain.LoadSkin CurrentSkin, True 'load new skin
End Sub

Private Sub cmdOk_Click()
  'Closes window
  Unload Me
End Sub

Private Sub Form_Activate()
  AlwaysOnTop Me, Settings.OnTop
End Sub

Private Sub Form_Load()
  'Setup form with each skin availabe
  GetInfo
End Sub

Private Sub lstSkins_Click()
  On Error Resume Next
  imgPreview.Picture = Skins(lstSkins.ListIndex + 1).Preview
  lblName = Skins(lstSkins.ListIndex + 1).name
  lblAuthor = Skins(lstSkins.ListIndex + 1).Author
  lblVersion = Skins(lstSkins.ListIndex + 1).Version
  txtNotes = Skins(lstSkins.ListIndex + 1).Notes
  
End Sub

Public Sub GetInfo()
  'finds all skins & reads info/gets preview image from them
  'adds skins to list
  On Error Resume Next
  Dim cFiles As New clsFind  'holds found skins
  Dim Read As New clsDatafile
  Dim X As Long, y As Long
  
  ReDim Skins(0)
  
  cFiles.Find App.Path & "\skins\", "*.sas", False 'find all skins
  For X = 1 To cFiles.Count 'process all found skins
    Read.FileName = cFiles(X).sFilename
    If Read.ReadStrFixed(3) = "SAS" Then 'correct header
      ReDim Preserve Skins(UBound(Skins) + 1) As Skin
    Else
      GoTo nxt
    End If
    y = Read.ReadNumber 'read skin file version
    With Skins(UBound(Skins)) 'read other info
      .fname = cFiles(X).sName
      .name = Read.ReadStr
      If y <> COMPILEVERSION Then .name = .name & " (wrong version)"
      .Author = Read.ReadStr
      .Notes = Read.ReadStr
      .Version = Read.ReadStr
      Read.ReadFile App.Path & "\skin.tmp" 'get preview image
      Set .Preview = LoadPicture(App.Path & "\skin.tmp")
      Kill App.Path & "\skin.tmp"
    End With
nxt:
  Next X
  
  For X = 1 To UBound(Skins) 'add skins to list & select current skin
    lstSkins.AddItem Skins(X).name
    If LCase(sFilename(Skins(X).fname, efpFileName)) = LCase(CurrentSkin) Then lstSkins.ListIndex = X - 1
  Next X
  
End Sub
