VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"mCol"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--------------------------------------------
'API Filesearch Class v1.0 (20:54 2003-07-12)
'by Paul Berlin (berlin_paul@hotmail.com)
'--------------------------------------------
'Other required files: clsFile.cls, modFind.bas
'You can spcify more than one filename when
'searching by seperating them with semicolon(;).
'
'Example:
'--------
'Dim cFind As New clsFind
'Dim x As Long
'
'For x = 1 to cFind.Find(App.Path, "*.cls;*.bas;*.frm", False)
'  Debug.Print x; cFind(x).sFilename, FormatNumber(cFind(x).lSize / 1024, 1) & " kb"
'Next x
Option Explicit

Private mCol As Collection

Public Function Clear()
  Set mCol = Nothing
  Set mCol = New Collection
End Function

Public Property Get Item(vntIndexKey As Variant) As clsFile
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
  Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
  mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As clsFile
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
  Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
  Set mCol = Nothing
End Sub

Public Function Find(ByVal sFindPath As String, ByVal sFindSpec As String, Optional ByVal bRecursive As Boolean = False) As Long
  Dim hFind As Long, lNext As Long, x As Long
  Dim tInfo As WIN32_FIND_DATA
  Dim oFile As clsFile
  Dim aFiles() As String
    
  'Make sure path ends with an '\'
  If Right(sFindPath, 1) <> "\" Then sFindPath = sFindPath & "\"
  
  'Scan for subdirs first
  If bRecursive Then
    hFind = FindFirstFile(sFindPath & "*.*", tInfo)
    lNext = 1
    Do While hFind > 0 And lNext > 0
      'If we found another dir, scan it to
      If (tInfo.dwFileAttributes And eDIRECTORY) = eDIRECTORY Then
        Dim sNTName As String
        sNTName = sNT(tInfo.cFileName) 'null trim filename
        If sNTName <> "." And sNTName <> ".." Then
          Find sFindPath & sNTName, sFindSpec, bRecursive
        End If
      End If
      lNext = FindNextFile(hFind, tInfo)
    Loop
    FindClose hFind
  End If
  
  'Get file names
  aFiles = Split(sFindSpec, ";")
  
  'Scan for files
  For x = 0 To UBound(aFiles)
    sFindSpec = sFindPath & aFiles(x)
    
    hFind = FindFirstFile(sFindSpec, tInfo)
    lNext = 1
    Do While hFind > 0 And lNext > 0
      'Add file if it isn't an directory
      If Not (tInfo.dwFileAttributes And eDIRECTORY) = eDIRECTORY Then
        Set oFile = New clsFile
        oFile.sFilename = sFindPath & sNT(tInfo.cFileName)
        mCol.Add oFile
      End If
      lNext = FindNextFile(hFind, tInfo)
    Loop
    FindClose hFind
  Next x
  
  Find = mCol.Count
  
End Function
