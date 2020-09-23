VERSION 5.00
Begin VB.Form frmPresetLoad 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Load Preset"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3150
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Presets"
      Height          =   4095
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   2895
      Begin VB.ComboBox cmbType 
         Height          =   315
         ItemData        =   "frmPresetLoad.frx":0000
         Left            =   120
         List            =   "frmPresetLoad.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   2655
      End
      Begin VB.ListBox lst 
         Height          =   3375
         ItemData        =   "frmPresetLoad.frx":0062
         Left            =   120
         List            =   "frmPresetLoad.frx":0069
         TabIndex        =   4
         Top             =   600
         Width           =   2655
      End
   End
   Begin VB.CheckBox chkLoad 
      Caption         =   "&Don't load colors and images."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Will prevent load of preset specific colors and images, making it fit better for the current skin"
      Top             =   4560
      Width           =   2415
   End
   Begin VB.OptionButton optType 
      Caption         =   "&Custom Presets"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      ToolTipText     =   "Load your own presets"
      Top             =   120
      Width           =   1455
   End
   Begin VB.OptionButton optType 
      Caption         =   "&Skin Presets"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Shows the built-in skin presets"
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
End
Attribute VB_Name = "frmPresetLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tPreset
  sFile As String
  sName As String
  bType As Byte
End Type

Private Presets() As tPreset

Private Sub cmbType_Click()
  On Error Resume Next
  Dim X As Long
  lst.Clear
  
  If optType(0) Then
    ShowSkinPlugins
  Else
    For X = 1 To UBound(Presets)
      If cmbType.ListIndex = 0 Or Presets(X).bType = cmbType.ListIndex Then
        lst.AddItem Presets(X).sName
      End If
    Next X
  End If
  
End Sub

Private Sub Form_Activate()
  AlwaysOnTop Me, Settings.OnTop
End Sub

Private Sub Form_Load()
  'FindPresets
  cmbType.ListIndex = 0
  chkLoad.Enabled = optType(1).Value
  ShowSkinPlugins
End Sub

Public Sub FindPresets()
  On Error Resume Next
  Dim cFind As New clsFind, X As Long
  Dim cFile As New clsDatafile
  
  ReDim Presets(0)
  
  cFind.Find App.Path & "\vis", "*.sap"
  For X = 1 To cFind.Count
    cFile.FileName = cFind(X).sFilename
    If cFile.ReadStrFixed(8) = "SAPRESET" And cFile.ReadNumber = PRESETVERSION Then
      ReDim Preserve Presets(UBound(Presets) + 1)
      With Presets(UBound(Presets))
        .sFile = cFind(X).sFilename
        .sName = cFind(X).sName
        .bType = cFile.ReadNumber
      End With
    End If
  Next X
  
  cmbType.ListIndex = 0
  cmbType_Click
  
End Sub

Private Sub lst_DblClick()
  On Error GoTo errh
  Dim X As Long
  
  If optType(0) Then
    For X = 1 To UBound(SkinPresets)
      If SkinPresets(X).sName = lst.List(lst.ListIndex) Then
        LoadSkinPreset X
      End If
    Next X
  Else
    For X = 1 To UBound(Presets)
      If Presets(X).sName = lst.List(lst.ListIndex) Then
        LoadPreset Presets(X).sFile, CBool(Abs(chkLoad.Value))
      End If
    Next X
  End If
  
  If frmPresetEdit.Visible Then
    frmPresetEdit.cmbType.ListIndex = Spectrum - 1
    frmPresetEdit.cmbType_Click
  End If
  
  Exit Sub
errh:
  If cLog.ErrorMsg(Err, "frmPresetLoad, lst_DblClick") = vbYes Then Resume Next Else frmMain.UnloadAll
End Sub

Private Sub optType_Click(Index As Integer)
  On Error Resume Next
  If Index = 0 Then
    ShowSkinPlugins
  Else
    FindPresets
  End If
  chkLoad.Enabled = optType(1).Value
End Sub

Public Sub ShowSkinPlugins()
  On Error Resume Next
  Dim X As Long
  
  lst.Clear
  
  For X = 1 To UBound(SkinPresets)
    If SkinPresets(X).bType = cmbType.ListIndex Or cmbType.ListIndex = 0 Then
      lst.AddItem SkinPresets(X).sName
    End If
  Next
  
End Sub
