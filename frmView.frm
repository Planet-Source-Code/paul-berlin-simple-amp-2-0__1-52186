VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmView 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "File Viewer"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9990
   Icon            =   "frmView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Caption         =   "ID3v2"
      Height          =   5175
      Index           =   1
      Left            =   240
      TabIndex        =   43
      Top             =   840
      Width           =   4575
      Visible         =   0   'False
      Begin VB.CheckBox chkv2 
         Caption         =   "&ID3v2 tag"
         Height          =   195
         Left            =   1080
         TabIndex        =   1
         Top             =   135
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox txtV2 
         Height          =   315
         Index           =   0
         Left            =   3960
         TabIndex        =   3
         Top             =   30
         Width           =   495
      End
      Begin VB.TextBox txtV2 
         Height          =   795
         Index           =   5
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   2130
         Width           =   3375
      End
      Begin VB.ComboBox cmbV2 
         Height          =   315
         Left            =   2400
         TabIndex        =   13
         Top             =   1695
         Width           =   2055
      End
      Begin VB.TextBox txtV2 
         Height          =   315
         Index           =   4
         Left            =   1080
         TabIndex        =   11
         Top             =   1695
         Width           =   615
      End
      Begin VB.TextBox txtV2 
         Height          =   315
         Index           =   3
         Left            =   1080
         TabIndex        =   9
         Top             =   1275
         Width           =   3375
      End
      Begin VB.TextBox txtV2 
         Height          =   315
         Index           =   2
         Left            =   1080
         TabIndex        =   7
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txtV2 
         Height          =   315
         Index           =   1
         Left            =   1080
         TabIndex        =   5
         Top             =   450
         Width           =   3375
      End
      Begin VB.TextBox txtV2 
         Height          =   315
         Index           =   6
         Left            =   1080
         TabIndex        =   17
         Top             =   3030
         Width           =   3375
      End
      Begin VB.TextBox txtV2 
         Height          =   315
         Index           =   7
         Left            =   1080
         TabIndex        =   19
         Top             =   3450
         Width           =   3375
      End
      Begin VB.TextBox txtV2 
         Height          =   315
         Index           =   8
         Left            =   1080
         TabIndex        =   21
         Top             =   3855
         Width           =   3375
      End
      Begin VB.TextBox txtV2 
         Height          =   315
         Index           =   9
         Left            =   1080
         TabIndex        =   23
         Top             =   4275
         Width           =   3375
      End
      Begin VB.TextBox txtV2 
         Height          =   315
         Index           =   10
         Left            =   1080
         TabIndex        =   25
         Top             =   4695
         Width           =   3375
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Trac&k #:"
         Height          =   195
         Left            =   3240
         TabIndex        =   2
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "C&omments:"
         Height          =   195
         Left            =   195
         TabIndex        =   14
         Top             =   2160
         Width           =   780
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Genre:"
         Height          =   195
         Left            =   1815
         TabIndex        =   12
         Top             =   1770
         Width           =   480
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Year:"
         Height          =   195
         Left            =   600
         TabIndex        =   10
         Top             =   1770
         Width           =   375
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Al&bum:"
         Height          =   195
         Left            =   495
         TabIndex        =   8
         Top             =   1365
         Width           =   480
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "A&rtist:"
         Height          =   195
         Left            =   585
         TabIndex        =   6
         Top             =   945
         Width           =   390
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Title:"
         Height          =   195
         Left            =   630
         TabIndex        =   4
         Top             =   525
         Width           =   345
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Com&poser:"
         Height          =   195
         Left            =   225
         TabIndex        =   16
         Top             =   3105
         Width           =   750
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ori&g. Artist:"
         Height          =   195
         Left            =   210
         TabIndex        =   18
         Top             =   3525
         Width           =   765
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cop&yright:"
         Height          =   195
         Left            =   270
         TabIndex        =   20
         Top             =   3930
         Width           =   705
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&URL:"
         Height          =   195
         Left            =   600
         TabIndex        =   22
         Top             =   4350
         Width           =   375
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "En&coded by:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   4770
         Width           =   900
      End
   End
   Begin MSComctlLib.Toolbar tbrSave 
      Height          =   330
      Left            =   3360
      TabIndex        =   53
      Top             =   6200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ButtonWidth     =   2540
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "iml"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save Changes"
            Key             =   "save"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrRevert 
      Height          =   330
      Left            =   2280
      TabIndex        =   52
      Top             =   6200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      ButtonWidth     =   1773
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "iml"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Revert"
            Key             =   "revert"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iml 
      Left            =   6480
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":08DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":0C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":0F86
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Caption         =   "ID3v1"
      Height          =   5175
      Index           =   2
      Left            =   5160
      TabIndex        =   45
      Top             =   600
      Width           =   4575
      Begin VB.Frame Frame2 
         Caption         =   "Other Tags (empty tags ignored)"
         Height          =   4935
         Left            =   120
         TabIndex        =   56
         Top             =   120
         Width           =   4335
         Begin VB.TextBox txtTags 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   4560
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   57
            Top             =   240
            Width           =   4095
         End
      End
   End
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Caption         =   "ID3v1"
      Height          =   5175
      Index           =   3
      Left            =   5160
      TabIndex        =   46
      Top             =   720
      Width           =   4575
      Begin VB.Frame Frame1 
         Caption         =   "Audio Information"
         Height          =   4935
         Index           =   0
         Left            =   120
         TabIndex        =   47
         Top             =   120
         Width           =   4335
         Begin VB.TextBox txtAudio 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   2520
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   26
            Top             =   2280
            Width           =   4095
         End
         Begin VB.Image imgCover 
            Height          =   1935
            Index           =   1
            Left            =   2160
            MousePointer    =   5  'Size
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1935
         End
         Begin VB.Image imgCover 
            Height          =   1935
            Index           =   0
            Left            =   120
            MousePointer    =   5  'Size
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1935
         End
      End
   End
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Caption         =   "ID3v1"
      Height          =   5175
      Index           =   0
      Left            =   240
      TabIndex        =   42
      Top             =   840
      Width           =   4575
      Begin MSComctlLib.Toolbar tbrCopyfrom 
         Height          =   330
         Left            =   1080
         TabIndex        =   54
         Top             =   2640
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         ButtonWidth     =   2884
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "iml"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Copy from ID3v2"
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtV1 
         Height          =   315
         Index           =   1
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   31
         Top             =   450
         Width           =   3375
      End
      Begin VB.TextBox txtV1 
         Height          =   315
         Index           =   2
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   33
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txtV1 
         Height          =   315
         Index           =   3
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   35
         Top             =   1275
         Width           =   3375
      End
      Begin VB.TextBox txtV1 
         Height          =   315
         Index           =   4
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   37
         Top             =   1695
         Width           =   615
      End
      Begin VB.ComboBox cmbV1 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   1695
         Width           =   2055
      End
      Begin VB.TextBox txtV1 
         Height          =   315
         Index           =   5
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   41
         Top             =   2130
         Width           =   3375
      End
      Begin VB.TextBox txtV1 
         Height          =   315
         Index           =   0
         Left            =   3960
         MaxLength       =   3
         TabIndex        =   29
         Top             =   30
         Width           =   495
      End
      Begin VB.CheckBox chkv1 
         Caption         =   "&ID3v1 tag"
         Height          =   195
         Left            =   1080
         TabIndex        =   27
         Top             =   135
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin MSComctlLib.Toolbar tbrCopyTo 
         Height          =   330
         Left            =   2880
         TabIndex        =   55
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         ButtonWidth     =   2566
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "iml"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Copy to ID3v2"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note: You can select other files in the playlist without closing this window."
         Height          =   390
         Left            =   120
         TabIndex        =   51
         Top             =   4680
         Width           =   4425
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Title:"
         Height          =   195
         Left            =   630
         TabIndex        =   30
         Top             =   525
         Width           =   345
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "A&rtist:"
         Height          =   195
         Left            =   585
         TabIndex        =   32
         Top             =   945
         Width           =   390
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Al&bum:"
         Height          =   195
         Left            =   495
         TabIndex        =   34
         Top             =   1365
         Width           =   480
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Year:"
         Height          =   195
         Left            =   600
         TabIndex        =   36
         Top             =   1770
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Genre:"
         Height          =   195
         Left            =   1815
         TabIndex        =   38
         Top             =   1770
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "C&omments:"
         Height          =   195
         Left            =   195
         TabIndex        =   40
         Top             =   2160
         Width           =   780
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Trac&k #:"
         Height          =   195
         Left            =   3240
         TabIndex        =   28
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Caption         =   "ID3v1"
      Height          =   5175
      Index           =   4
      Left            =   240
      TabIndex        =   48
      Top             =   840
      Width           =   4575
      Begin VB.Frame Frame1 
         Caption         =   "Misc Information"
         Height          =   4935
         Index           =   1
         Left            =   120
         TabIndex        =   49
         Top             =   120
         Width           =   4335
         Begin VB.TextBox txtMisc 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   4575
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   50
            Top             =   240
            Width           =   4095
         End
      End
   End
   Begin VB.TextBox txtLocation 
      ForeColor       =   &H80000011&
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   120
      Width           =   4785
   End
   Begin MSComctlLib.TabStrip tabTags 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   9975
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ID3v&1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ID3v&2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Other Tags"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Audio"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Misc"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sFile As String
Private index As Long
Private File As New clsFile
Private Covers As New clsFind
Private FrontBack(1) As Long
Private OldFrontBack(1) As String

Public Sub View(ByVal libindex As String)
  On Error Resume Next
  Dim eTagType As FSOUND_TAGFIELD_TYPE
  Dim sTagName As String, sTagValue As String
  Dim X As Long, lMod As Long, hFile As Long
  Dim ChangeTab As Boolean
  
  ChangeTab = Not Me.Visible
  If ChangeTab Then
    OldFrontBack(0) = ""
    OldFrontBack(1) = ""
  End If
  
  sFile = Library(libindex).sFilename
  index = libindex
  txtLocation = sFile
  txtLocation.SelStart = Len(txtLocation)
  
  'Update data
  If FileExists(sFile) Then
    
    'Clear & setup controls
    txtAudio = ""
    txtMisc = ""
    txtTags = ""
    chkv1.Value = vbUnchecked
    chkv2.Value = vbUnchecked
    If cmbV1.ListCount = 0 Then 'only first time
      For X = 0 To UBound(Genre) 'Fill comboboxes with genres
        cmbV1.AddItem Genre(X)
        cmbV2.AddItem Genre(X)
      Next
    End If
    
    File.sFilename = sFile
    
    'Refresh Misc tab
    txtMisc = txtMisc & "Size: " & FormatNumber(File.lSize, 0) & " bytes" & vbCrLf & vbCrLf
    txtMisc = txtMisc & "Created: " & File.dCreationTime & vbCrLf
    txtMisc = txtMisc & "Last Modified: " & File.dLastWriteTime & vbCrLf
    txtMisc = txtMisc & "Last Accessed: " & File.dLastAccessTime & vbCrLf & vbCrLf
    If File.eFileAttributes And eREADONLY Then
      txtMisc = txtMisc & "Read only: Yes" & vbCrLf
    Else
      txtMisc = txtMisc & "Read only: No" & vbCrLf
    End If
    If File.eFileAttributes And eHIDDEN Then
      txtMisc = txtMisc & "Hidden: Yes" & vbCrLf & vbCrLf
    Else
      txtMisc = txtMisc & "Hidden: No" & vbCrLf & vbCrLf
    End If
    txtMisc = txtMisc & "Last Playing Date: " & Library(libindex).dLastPlayDate & vbCrLf
    txtMisc = txtMisc & "Times Started: " & Library(libindex).lTimesPlayed & vbCrLf
    txtMisc = txtMisc & "Times Skipped: " & Library(libindex).lTimesSkipped & vbCrLf
    txtMisc = txtMisc & "Times Played: " & Library(libindex).lTimesPlayed - Library(libindex).lTimesSkipped & vbCrLf
    
    If Library(libindex).eType <= 5 Then 'Stream type
    
      hFile = Sound.GetStreamOpenFile(sFile, Library(libindex).bIsVBR)
      
      'Read all tags
      lMod = -1 'temp use, to see change of tag type
      For X = 0 To Sound.GetStreamTagNumTags(hFile) - 1
        Call Sound.GetStreamTagByNum(hFile, X, eTagType, sTagName, sTagValue)
        If Len(sTagValue) > 0 Then
          If lMod <> eTagType Then
            Select Case eTagType
              Case 0
                txtTags = txtTags & vbCrLf & "OGG VORBIS TAG:" & vbCrLf & vbCrLf
              Case 1
                txtTags = txtTags & vbCrLf & "ID3V1 TAG:" & vbCrLf & vbCrLf
                chkv1.Value = vbChecked
              Case 2
                txtTags = txtTags & vbCrLf & "ID3V2 TAG:" & vbCrLf & vbCrLf
                chkv2.Value = vbChecked
              Case 3
                txtTags = txtTags & vbCrLf & "SHOUTCAST TAG:" & vbCrLf & vbCrLf
              Case 4
                txtTags = txtTags & vbCrLf & "ICECAST TAG:" & vbCrLf & vbCrLf
              Case 5
                txtTags = txtTags & vbCrLf & "ASF TAG:" & vbCrLf & vbCrLf
            End Select
            lMod = eTagType
          End If
          txtTags = txtTags & sTagName & ": " & vbTab & sTagValue & vbCrLf
        End If
      Next X
      
      lMod = 0
      
    End If
    
    'Add file type specific info
    If Library(libindex).eType = TYPE_MP2_MP3 Then
      'MP3, MP2

      'Setup ID3v1 textboxes
      If chkv1.Value Then
        txtV1(0) = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V1, "TRACK")
        txtV1(1) = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V1, "TITLE")
        txtV1(2) = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V1, "ARTIST")
        txtV1(3) = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V1, "ALBUM")
        txtV1(4) = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V1, "YEAR")
        txtV1(5) = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V1, "COMMENT")
        If Len(Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V1, "GENRE")) > 0 Then
          lMod = Val(Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V1, "GENRE"))
          If lMod < cmbV1.ListCount - 1 Then
            cmbV1.ListIndex = lMod
          End If
          lMod = 0
        End If
      Else
        cmbV1.ListIndex = -1
        For X = 0 To 5
          txtV1(X).Text = ""
        Next
      End If
      
      'Setup ID3v2 textboxes
      If chkv2.Value Then
        txtV2(0) = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "TRCK")
        txtV2(1) = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "TIT2")
        txtV2(2) = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "TPE1")
        txtV2(3) = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "TALB")
        txtV2(4) = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "TYER")
        txtV2(5) = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "COMM")
        txtV2(6) = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "TCOM")
        txtV2(7) = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "TOPE")
        txtV2(8) = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "TCOP")
        txtV2(9) = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "WXXX")
        txtV2(10) = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "TENC")
        cmbV2.Text = Sound.GetStreamTag(hFile, FSOUND_TAGFIELD_ID3V2, "TCON")
        If Left(cmbV2.Text, 1) = "(" Then
          cmbV2.Text = Right(cmbV2.Text, Len(cmbV2.Text) - InStr(1, cmbV2.Text, ")"))
        End If
      Else
        cmbV2.Text = ""
        For X = 0 To 10
          txtV2(X).Text = ""
        Next
      End If
      
      'Update MPEG info
      Dim cMpeg As New cMP3Info
      cMpeg.FileName = sFile
      cMpeg.ReadMP3Header
      txtAudio = txtAudio & "Type: " & cMpeg.ID & " " & cMpeg.Layer & vbCrLf & vbCrLf
      If chkv1 And chkv2 Then
        txtAudio = txtAudio & "Tags: ID3v1, ID3v2" & vbCrLf
      ElseIf chkv1 Then
        txtAudio = txtAudio & "Tags: ID3v1" & vbCrLf
      ElseIf chkv2 Then
        txtAudio = txtAudio & "Tags: ID3v2" & vbCrLf
      Else
        txtAudio = txtAudio & "Tags: None" & vbCrLf
      End If
      txtAudio = txtAudio & "Playing Length: " & ConvertTime(Library(libindex).lLength) & vbCrLf & vbCrLf
      
      txtAudio = txtAudio & "Variable Bitrate: " & CBool(Library(libindex).bIsVBR) & vbCrLf
      txtAudio = txtAudio & "Channel Mode: " & cMpeg.Mode & vbCrLf
      txtAudio = txtAudio & "Bitrate: " & cMpeg.Bitrate & " kbps" & vbCrLf
      txtAudio = txtAudio & "Frequency: " & cMpeg.Frequency & " hz" & vbCrLf
      txtAudio = txtAudio & "Copyright: " & cMpeg.Copyrighted & vbCrLf
      txtAudio = txtAudio & "Emphasis: " & cMpeg.Emphasis & vbCrLf
      txtAudio = txtAudio & "Original: " & cMpeg.Original & vbCrLf
      Set cMpeg = Nothing
      
    Else 'If not mp3, clear controls
      cmbV1.ListIndex = -1
      cmbV2.Text = ""
      For X = 0 To 5
        txtV1(X).Text = ""
      Next
      For X = 0 To 10
        txtV2(X).Text = ""
      Next
    End If
    
    If Library(libindex).eType = TYPE_OGG Then
      txtAudio = txtAudio & "Type: Ogg Vorbis File" & vbCrLf & vbCrLf
      txtAudio = txtAudio & "Playing Length: " & ConvertTime(Library(libindex).lLength) & vbCrLf & vbCrLf
      txtAudio = txtAudio & "Bitrate: " & CStr(Round((FileLen(sFile) / Library(libindex).lLength) * 8 / 1000, 2)) & " Kbps" & vbCrLf
    
    ElseIf Library(libindex).eType = TYPE_ASF Then
      txtAudio = txtAudio & "Type: Advanced Sound Format" & vbCrLf & vbCrLf
      txtAudio = txtAudio & "Playing Length: " & ConvertTime(Library(libindex).lLength) & vbCrLf & vbCrLf
      txtAudio = txtAudio & "Bitrate: " & CStr(Round((FileLen(sFile) / Library(libindex).lLength) * 8 / 1000, 2)) & " Kbps" & vbCrLf
      
    ElseIf Library(libindex).eType = TYPE_IT Then
      txtAudio = txtAudio & "Type: Impulse Tracker Module" & vbCrLf & vbCrLf
      txtAudio = txtAudio & "Playing Length: " & ConvertTime(Library(libindex).lLength) & vbCrLf & vbCrLf
      If lMod <> 0 Then FMUSIC_FreeSong lMod
      lMod = FMUSIC_LoadSong(sFile)
      txtAudio = txtAudio & "Title: " & GetStringFromPointer(FMUSIC_GetName(lMod)) & vbCrLf & vbCrLf
      txtAudio = txtAudio & "Channels: " & FMUSIC_GetNumChannels(lMod) & vbCrLf
      txtAudio = txtAudio & "Instruments: " & FMUSIC_GetNumInstruments(lMod) & vbCrLf
      txtAudio = txtAudio & "Samples: " & FMUSIC_GetNumSamples(lMod) & vbCrLf
      txtAudio = txtAudio & "Orders: " & FMUSIC_GetNumOrders(lMod) & vbCrLf
      txtAudio = txtAudio & "Patterns: " & FMUSIC_GetNumPatterns(lMod) & vbCrLf
      FMUSIC_FreeSong lMod
      lMod = 0
      
    ElseIf Library(libindex).eType = TYPE_MID_RMI Then
      txtAudio = txtAudio & "Type: MIDI File" & vbCrLf & vbCrLf
      txtAudio = txtAudio & "Playing Length: " & ConvertTime(Library(libindex).lLength) & vbCrLf & vbCrLf
      
    ElseIf Library(libindex).eType = TYPE_MOD Then
      txtAudio = txtAudio & "Type: Protracker/Fasttracker Module" & vbCrLf & vbCrLf
      txtAudio = txtAudio & "Playing Length: " & ConvertTime(Library(libindex).lLength) & vbCrLf & vbCrLf
      If lMod <> 0 Then FMUSIC_FreeSong lMod
      lMod = FMUSIC_LoadSong(sFile)
      txtAudio = txtAudio & "Title: " & GetStringFromPointer(FMUSIC_GetName(lMod)) & vbCrLf & vbCrLf
      txtAudio = txtAudio & "Channels: " & FMUSIC_GetNumChannels(lMod) & vbCrLf
      txtAudio = txtAudio & "Instruments: " & FMUSIC_GetNumInstruments(lMod) & vbCrLf
      txtAudio = txtAudio & "Samples: " & FMUSIC_GetNumSamples(lMod) & vbCrLf
      txtAudio = txtAudio & "Orders: " & FMUSIC_GetNumOrders(lMod) & vbCrLf
      txtAudio = txtAudio & "Patterns: " & FMUSIC_GetNumPatterns(lMod) & vbCrLf
      FMUSIC_FreeSong lMod
      lMod = 0
        
    ElseIf Library(libindex).eType = TYPE_S3M Then
      txtAudio = txtAudio & "Type: Screamtracker 3 Module" & vbCrLf & vbCrLf
      If lMod <> 0 Then FMUSIC_FreeSong lMod
      lMod = FMUSIC_LoadSong(sFile)
      txtAudio = txtAudio & "Title: " & GetStringFromPointer(FMUSIC_GetName(lMod)) & vbCrLf & vbCrLf
      txtAudio = txtAudio & "Playing Length: " & ConvertTime(Library(libindex).lLength) & vbCrLf & vbCrLf
      txtAudio = txtAudio & "Channels: " & FMUSIC_GetNumChannels(lMod) & vbCrLf
      txtAudio = txtAudio & "Instruments: " & FMUSIC_GetNumInstruments(lMod) & vbCrLf
      txtAudio = txtAudio & "Samples: " & FMUSIC_GetNumSamples(lMod) & vbCrLf
      txtAudio = txtAudio & "Orders: " & FMUSIC_GetNumOrders(lMod) & vbCrLf
      txtAudio = txtAudio & "Patterns: " & FMUSIC_GetNumPatterns(lMod) & vbCrLf
      FMUSIC_FreeSong lMod
      lMod = 0
      
    ElseIf Library(libindex).eType = TYPE_SGM Then
      txtAudio = txtAudio & "Type: DirectMusic Segment File" & vbCrLf & vbCrLf
      txtAudio = txtAudio & "Playing Length: " & ConvertTime(Library(libindex).lLength) & vbCrLf & vbCrLf
      
    ElseIf Library(libindex).eType = TYPE_WAV Then
      txtAudio = txtAudio & "Type: Waveform Audio File" & vbCrLf & vbCrLf
      txtAudio = txtAudio & "Playing Length: " & ConvertTime(Library(libindex).lLength) & vbCrLf & vbCrLf
      txtAudio = txtAudio & "Bitrate: " & CStr(Round((FileLen(sFile) / Library(libindex).lLength) * 8 / 1000, 2)) & " Kbps" & vbCrLf
      
    ElseIf Library(libindex).eType = TYPE_WMA Then
      txtAudio = txtAudio & "Type: Windows Media Audio File" & vbCrLf & vbCrLf
      txtAudio = txtAudio & "Playing Length: " & ConvertTime(Library(libindex).lLength) & vbCrLf & vbCrLf
      txtAudio = txtAudio & "Bitrate: " & CStr(Round((FileLen(sFile) / Library(libindex).lLength) * 8 / 1000, 2)) & " Kbps" & vbCrLf
      
    ElseIf Library(libindex).eType = TYPE_XM Then
      txtAudio = txtAudio & "Type: Fasttracker 2 Module" & vbCrLf & vbCrLf
      txtAudio = txtAudio & "Playing Length: " & ConvertTime(Library(libindex).lLength) & vbCrLf & vbCrLf
      If lMod <> 0 Then FMUSIC_FreeSong lMod
      lMod = FMUSIC_LoadSong(sFile)
      txtAudio = txtAudio & "Title: " & GetStringFromPointer(FMUSIC_GetName(lMod)) & vbCrLf & vbCrLf
      txtAudio = txtAudio & "Channels: " & FMUSIC_GetNumChannels(lMod) & vbCrLf
      txtAudio = txtAudio & "Instruments: " & FMUSIC_GetNumInstruments(lMod) & vbCrLf
      txtAudio = txtAudio & "Samples: " & FMUSIC_GetNumSamples(lMod) & vbCrLf
      txtAudio = txtAudio & "Orders: " & FMUSIC_GetNumOrders(lMod) & vbCrLf
      txtAudio = txtAudio & "Patterns: " & FMUSIC_GetNumPatterns(lMod) & vbCrLf
      FMUSIC_FreeSong lMod
      lMod = 0
      
    End If
    
    If hFile <> 0 Then
      Sound.GetStreamCloseFile hFile 'close file
      hFile = 0
    End If
    Set File = Nothing
    
  Else
    MsgBox "The file """" & sfile & """" cannot be viewed because it does not exist!", vbExclamation, "File Error"
  End If
  
  'Get covers
  Erase FrontBack
  If Library(libindex).eType = TYPE_MP2_MP3 Or Library(libindex).eType = TYPE_OGG Or Library(libindex).eType = TYPE_WMA Then
    Covers.Clear
    Covers.Find sFilename(sFile, efpFilePath), "*.jpg;*.jpeg;*.jpe;*.bmp;*.gif"
    
    If Covers.Count > 0 Then
      
      Dim Front() As Long
      Dim Back() As Long
      Dim sArtist As String
      Dim sAlbum As String
      Dim tStr As String
      
      'Establish artist & album
      If Len(txtV1(2)) > 0 Then
        sArtist = LCase(txtV1(2))
      End If
      If Len(txtV2(2)) > 0 Then
        sArtist = LCase(txtV2(2))
      End If
      If Len(sArtist) = 0 And InStr(1, sFilename(sFile, efpFileName), "-") > 0 Then
        tStr = sFilename(sFile, efpFileName)
        sArtist = LCase(Trim(Left(tStr, InStr(1, tStr, "-") - 1)))
      End If
      If Len(txtV1(3)) > 0 Then
        sAlbum = LCase(txtV1(3))
      End If
      If Len(txtV2(3)) > 0 Then
        sAlbum = LCase(txtV2(3))
      End If
      If Len(sAlbum) = 0 Then
        tStr = Left(sFilename(sFile, efpFilePath), Len(sFilename(sFile, efpFilePath)) - 1)
        If InStrRev(tStr, "-") > InStrRev(tStr, "\") Then
          sAlbum = LCase(Trim(Right$(tStr, Len(tStr) - InStrRev(tStr, "-"))))
        Else
          sAlbum = LCase(Trim(Right$(tStr, Len(tStr) - InStrRev(tStr, "\"))))
        End If
      End If
      
      'Get data from each found cover
      ReDim Front(Covers.Count)
      ReDim Back(Covers.Count)
      For X = 1 To Covers.Count
        tStr = LCase(Covers(X).sFilename)
        If InStr(1, tStr, sArtist) Then
          Front(X) = Front(X) + 1
          Back(X) = Back(X) + 1
        End If
        If InStr(1, tStr, sAlbum) Then
          Front(X) = Front(X) + 1
          Back(X) = Back(X) + 1
        End If
        If InStr(1, tStr, "front") Then
          Front(X) = Front(X) + 1
        End If
        If InStr(1, tStr, "back") Then
          Back(X) = Back(X) + 1
        End If
      Next
      
      Dim y As Long
      Dim d As Long
      'See which found cover filename was the best matches
      For X = 1 To Covers.Count
        If Front(X) >= y Then
          y = Front(X)
          d = X
        End If
      Next
      FrontBack(0) = d
      y = 0: d = 0
      For X = 1 To Covers.Count
        If Back(X) >= y Then
          y = Back(X)
          d = X
        End If
      Next
      FrontBack(1) = d
      
      If FrontBack(0) = FrontBack(1) Then
        FrontBack(1) = 0
      End If
      
      If FrontBack(1) > 0 Then
        If Covers(FrontBack(0)).sFilename <> OldFrontBack(0) Then
          imgCover(0).Picture = LoadPicture(Covers(FrontBack(0)).sFilename)
          imgCover(0).ToolTipText = sFilename(Covers(FrontBack(0)).sFilename, efpFileNameAndExt)
        End If
        If Covers(FrontBack(1)).sFilename <> OldFrontBack(1) Then
          imgCover(1).Picture = LoadPicture(Covers(FrontBack(1)).sFilename)
          imgCover(1).ToolTipText = sFilename(Covers(FrontBack(1)).sFilename, efpFileNameAndExt)
        End If
      Else
        If Covers(FrontBack(0)).sFilename <> OldFrontBack(0) Then
          imgCover(0).Picture = LoadPicture(Covers(FrontBack(0)).sFilename)
          imgCover(0).ToolTipText = sFilename(Covers(FrontBack(0)).sFilename, efpFileNameAndExt)
        End If
        imgCover(1).Picture = Nothing
        imgCover(1).Visible = False
      End If

    End If
  End If
  
  If FrontBack(0) + FrontBack(1) = 0 Then
    imgCover(1).Picture = Nothing
    imgCover(0).Picture = Nothing
    imgCover(0).Visible = False
    imgCover(1).Visible = False
    txtAudio.Top = 240
    txtAudio.Height = 4560
    OldFrontBack(0) = ""
    OldFrontBack(1) = ""
  Else
    imgCover(0).Visible = True
    'imgCover(1).Visible = True
    txtAudio.Top = 2280
    txtAudio.Height = 2520
    OldFrontBack(0) = Covers(FrontBack(0)).sFilename
    OldFrontBack(1) = Covers(FrontBack(1)).sFilename
  End If
  
  'update visual aspects of controls
  If Me.Visible = False Then
    For X = 0 To frm.Count - 1
      frm(X).Left = frm(1).Left
      frm(X).Top = frm(1).Top
    Next
    Me.Width = 5130
    tabTags_Click
    Me.Show
    DoEvents
    
    AlwaysOnTop Me, Settings.OnTop
  End If
  
  If ChangeTab Then
    If chkv2.Value Then
      tabTags.tabs(2).Selected = True
    ElseIf Not chkv1.Value Then
      tabTags.tabs(3).Selected = True
    End If
  End If
  
End Sub

Private Sub imgCover_DblClick(index As Integer)
  On Error GoTo errh
  ShellExecute Me.hwnd, "open", Covers(FrontBack(index)).sFilename, vbNullString, vbNullString, 1
  
  Exit Sub
errh:
  MsgBox "Could not open this file with an associated program. Mayhaps ye have not associated it?", vbCritical
End Sub

Private Sub tabTags_Click()
  On Error Resume Next
  Dim X As Integer
  For X = 0 To frm.Count - 1
    frm(X).Visible = False
  Next
  frm(tabTags.SelectedItem.index - 1).Visible = True
End Sub

Private Sub tbrCopyfrom_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim X As Long
  chkv1.Value = vbChecked
  For X = 0 To 5
    txtV1(X) = txtV2(X)
  Next
End Sub

Private Sub tbrCopyTo_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim X As Long
  chkv2.Value = vbChecked
  For X = 0 To 5
    txtV2(X) = txtV1(X)
  Next
  cmbV2.Text = cmbV1.List(cmbV1.ListIndex)
End Sub

Private Sub tbrRevert_ButtonClick(ByVal Button As MSComctlLib.Button)
  View index
End Sub

Private Sub tbrSave_ButtonClick(ByVal Button As MSComctlLib.Button)
  On Error Resume Next
  Dim v1 As New clsID3v1
  Dim v2 As New clsID3v23
  v1.FileName = sFile
  v2.FileName = sFile
  'ID3v1
  If chkv1.Value = 0 And v1.HasTag Then
    If index = PlayingLib Then
      MsgBox "ID3 Tags cannot be removed when the file is being played. Stop the song and try again.", vbExclamation
      Exit Sub
    Else
      v1.RemoveTag
    End If
  Else
    If Val(txtV1(0)) > 0 Then
      v1.tagTrack = Val(txtV1(0))
    End If
    v1.tagTitle = txtV1(1)
    v1.tagArtist = txtV1(2)
    v1.tagAlbum = txtV1(3)
    v1.tagYear = txtV1(4)
    v1.tagComments = txtV1(5)
    If cmbV1.ListIndex < 0 Then
      v1.tagGenre = 255
    Else
      v1.tagGenre = cmbV1.ListIndex
    End If
    v1.SaveTag
  End If
  'ID3v2
  If chkv2.Value = 0 And v2.HasTag Then
    If index = PlayingLib Then
      MsgBox "ID3 Tags cannot be removed when the file is being played. Stop the song and try again.", vbExclamation
      Exit Sub
    Else
      v2.RemoveTag
    End If
  Else
    v2.tagTrack = txtV2(0)
    v2.tagTitle = txtV2(1)
    v2.tagArtist = txtV2(2)
    v2.tagAlbum = txtV2(3)
    v2.tagYear = txtV2(4)
    v2.tagComments = txtV2(5)
    v2.tagComposer = txtV2(6)
    v2.tagOrigArtist = txtV2(7)
    v2.tagCopyright = txtV2(8)
    v2.tagURL = txtV2(9)
    v2.tagEncodedBy = txtV2(10)
    v2.tagGenre = cmbV2.Text
    v2.SaveTag
  End If
  modMain.UpdateLibrary index
  frmPlaylist.List.Refresh
End Sub
