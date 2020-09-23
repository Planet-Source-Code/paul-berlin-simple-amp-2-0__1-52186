VERSION 5.00
Begin VB.Form frmSearch 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAll 
      Caption         =   "Select &All"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Find &Next"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox cmbSearch 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Search for:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   780
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SearchIndex As Long

Private Sub cmdAll_Click()
  On Error GoTo errh
  Dim A As Boolean, X As Long, sTxt As String
  
  sTxt = Trim(cmbSearch.Text)
  
  If Len(sTxt) > 0 Then
    
    For X = 0 To cmbSearch.ListCount - 1
      If cmbSearch.Text = cmbSearch.List(X) Then A = True
    Next
    If Not A Then cmbSearch.AddItem cmbSearch.Text
    
    For X = 1 To UBound(Playlist)
      A = False
      If InStr(1, Library(Playlist(X).Reference).sArtistTitle, sTxt, vbTextCompare) Then A = True
      If InStr(1, Library(Playlist(X).Reference).sAlbum, sTxt, vbTextCompare) Then A = True
      Playlist(X).Selected = A
    Next
    
    frmPlaylist.List.Refresh
    
  End If
  
  Exit Sub
errh:
  If cLog.ErrorMsg(Err, "frmSearch, cmdAll_Click") = vbYes Then Resume Next Else frmMain.UnloadAll
End Sub

Private Sub cmdNext_Click()
  On Error GoTo errh
  Dim A As Boolean, X As Long, sTxt As String
  
  sTxt = Trim(cmbSearch.Text)
  
  If Len(sTxt) > 0 Then

    For X = 0 To cmbSearch.ListCount - 1
      If cmbSearch.Text = cmbSearch.List(X) Then A = True
    Next
    If Not A Then cmbSearch.AddItem cmbSearch.Text
    
    If SearchIndex = 0 Then SearchIndex = 1
    
    frmMenus.menSelectNone_Click
    For X = SearchIndex To UBound(Playlist)
      A = False
      If InStr(1, Library(Playlist(X).Reference).sArtistTitle, sTxt, vbTextCompare) Then A = True
      If InStr(1, Library(Playlist(X).Reference).sAlbum, sTxt, vbTextCompare) Then A = True
      Playlist(X).Selected = A
      If A Then
        frmPlaylist.List.MakeVisible Playlist(X).Index
        SearchIndex = X + 1
        Exit For
      End If
    Next
    
    If Not A Then
      MsgBox Chr(34) & cmbSearch.Text & Chr(34) & " could not be found.", vbInformation, "Not found"
      SearchIndex = 0
    End If

    frmPlaylist.List.Refresh
    
  End If
  
  Exit Sub
errh:
  If cLog.ErrorMsg(Err, "frmSearch, cmdNext_Click") = vbYes Then Resume Next Else frmMain.UnloadAll
End Sub

Private Sub Form_Activate()
  On Error Resume Next
  AlwaysOnTop Me, Settings.OnTop
End Sub

Private Sub Form_Load()
  ctlSetFocus Me
End Sub
