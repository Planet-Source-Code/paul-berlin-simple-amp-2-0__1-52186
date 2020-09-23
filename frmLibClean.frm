VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLibClean 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Media Library Maintenance"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   ControlBox      =   0   'False
   Icon            =   "frmLibClean.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   1575
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   4455
      Begin VB.CheckBox chkOpt 
         Caption         =   "Remove all files that are shorter than 20 seconds."
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   3855
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Remove all midi files from the library."
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   3015
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Remove all wave files from the library."
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   4215
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Do not remove files on removable drives (such as CD-Drives, etc.) from the library."
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   1  'Checked
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   4455
      Begin MSComctlLib.ProgressBar prbMain 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Min             =   1e-4
         Max             =   200
         Scrolling       =   1
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Not Started."
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   $"frmLibClean.frx":058A
      Height          =   585
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4455
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmLibClean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DRIVE_CDROM = 5
Private Const DRIVE_FIXED = 3
Private Const DRIVE_RAMDISK = 6
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_NOTFOUND = 1
Private Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdStart_Click()
  On Error GoTo ErrH
  Dim x As Long, Chk() As Boolean, Count As Long
  
  For x = 0 To 3
    chkOpt(x).Enabled = False
  Next x
  cmdStart.Enabled = False
  cmdCancel.Enabled = False
  
  ReDim Chk(UBound(LibraryIndex)) 'array used to keep track of missing files, True removes file
  
  'First step, find missing files
  lblInfo = "Checking if files exist..."
  For x = 1 To UBound(LibraryIndex)
    prbMain.Value = (x / UBound(LibraryIndex)) * 100
    DoEvents
    
    Chk(x) = False 'default to keep file in database
    
    If chkOpt(1).Value = vbChecked Or chkOpt(2).Value = vbChecked Then
      'remove wave-files and midi-files
    
      If LibraryIndex(x).lReference = 0 Then 'load file info from database
        ReDim Preserve Library(UBound(Library) + 1)
        LibraryIndex(x).lReference = UBound(Library)
        LoadItem UBound(Library), x
      End If
      
      If chkOpt(1).Value = vbChecked Then
        If Library(LibraryIndex(x).lReference).eType = TYPE_WAV Then Chk(x) = True
      End If
      If chkOpt(2).Value = vbChecked Then
        If Library(LibraryIndex(x).lReference).eType = TYPE_MID_RMI Then Chk(x) = True
      End If
    End If
    
    If chkOpt(3).Value = vbChecked And Not Chk(x) Then
      'remove short files (<= 20 sec)
      
      If LibraryIndex(x).lReference = 0 Then 'load file info from database
        ReDim Preserve Library(UBound(Library) + 1)
        LibraryIndex(x).lReference = UBound(Library)
        LoadItem UBound(Library), x
      End If
      
      If Library(LibraryIndex(x).lReference).lLength <= 20 Then Chk(x) = True
      
    End If
    
    If Not Chk(x) Then
      Select Case GetDriveType(Left(LibraryIndex(x).sFilename, 3))
        Case DRIVE_CDROM, DRIVE_REMOVABLE
          If chkOpt(0).Value = vbUnchecked Then Chk(x) = True 'if on removable drive, and not checkbox unchecked, remove
        Case DRIVE_FIXED, DRIVE_RAMDISK, DRIVE_REMOTE
          If Not FileExists(LibraryIndex(x).sFilename) Then Chk(x) = True 'if on fixed disk, and does not exist, remove
        Case Else 'drive not found or other error
          Chk(x) = True 'remove
      End Select
    End If
    
    If Chk(x) Then Count = Count + 1
  Next
  
  'Second step, rearrange LibraryIndex()
  'It moves all existing files to beginning of array, and then
  'redims so the old ones are removed.
  lblInfo = "Updating Media Library..."
  If Count > 0 Then
    x = 1
    Dim x2 As Long, d As Boolean
    Do
      prbMain.Value = (x / UBound(LibraryIndex)) * 100 + 100
      DoEvents
      If Chk(x) Then
        x2 = x + 1
        d = False
        Do
          If Chk(x2) Then
            x2 = x2 + 1
          Else
            Chk(x2) = True
            d = True
            LibraryIndex(x) = LibraryIndex(x2)
          End If
        Loop Until d Or x2 >= UBound(LibraryIndex)
        If x2 >= UBound(LibraryIndex) Then Exit Do
      End If
      x = x + 1
    Loop Until x >= UBound(LibraryIndex) - Count
    ReDim Preserve LibraryIndex(UBound(LibraryIndex) - Count)
    LibraryChanged = True
  End If
  prbMain.Value = 200
    
  lblInfo = "Finished. " & Count & " files where removed from the library."
  
  cmdCancel.Caption = "&Close"
  cmdCancel.Enabled = True
  
  Exit Sub
ErrH:
  If cLog.ErrorMsg(Err, "frmLibClean, cmdStart_Click()") = vbNo Then frmMain.UnloadAll Else Resume Next
End Sub

Private Sub Form_Activate()
  On Error Resume Next

  AlwaysOnTop Me, Settings.OnTop
End Sub

Private Sub Form_Load()
  On Error Resume Next
  lblInfo = "There are " & UBound(LibraryIndex) & " items in your Media Library."
End Sub
