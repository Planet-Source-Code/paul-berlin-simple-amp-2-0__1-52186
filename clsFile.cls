VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--------------------------------------------
'API Filesearch Class v1.01 (15:06 2003-08-25)
'by Paul Berlin (berlin_paul@hotmail.com)
'--------------------------------------------
'Other required files: clsFind.cls, modFind.bas
Option Explicit

Private tFileInfo As BY_HANDLE_FILE_INFORMATION
Private mvarFilename As String

Public Property Let eFileAttributes(ByVal vData As enumFileAttributes)
  SetFileAttributes mvarFilename, vData
  tFileInfo.dwFileAttributes = vData
End Property

Public Property Get eFileAttributes() As enumFileAttributes
  eFileAttributes = tFileInfo.dwFileAttributes
End Property

Public Property Let sFilename(ByVal vData As String)
  mvarFilename = vData
  RefreshInfo
End Property

Public Property Get sFilename() As String
Attribute sFilename.VB_UserMemId = 0
  sFilename = mvarFilename
End Property

Public Function RefreshInfo() As Boolean
  Dim hFile As Long
  
  hFile = CreateFile(mvarFilename, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal CLng(0), OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0)
  If GetFileInformationByHandle(hFile, tFileInfo) = 0 Then
    tFileInfo.dwFileAttributes = GetFileAttributes(mvarFilename)
  End If
  CloseHandle hFile
End Function

Public Property Get dCreationTime() As Date
  dCreationTime = ConvertDate(tFileInfo.ftCreationTime)
End Property

Public Property Get dLastAccessTime() As Date
  dLastAccessTime = ConvertDate(tFileInfo.ftLastAccessTime)
End Property

Public Property Get dLastWriteTime() As Date
  dLastWriteTime = ConvertDate(tFileInfo.ftLastWriteTime)
End Property

Private Function ConvertDate(ftDate As FILETIME) As Date
  Dim tST As SYSTEMTIME
  
  Call FileTimeToLocalFileTime(ftDate, ftDate)
  Call FileTimeToSystemTime(ftDate, tST)
  ConvertDate = CDate(Format(tST.wMonth, "00") & "/" & Format(tST.wDay, "00") & "/" & Format(tST.wYear, "0000") & " " & Format(tST.wHour, "00") & ":" & Format(tST.wMinute, "00") & ":" & Format(tST.wSecond, "00"))
End Function

Public Property Get lSize() As Long
  lSize = tFileInfo.nFileSizeHigh * 2 ^ 32 + tFileInfo.nFileSizeLow
End Property

Public Property Get sExtension() As String
  sExtension = modMisc.sFilename(mvarFilename, efpFileExt)
End Property

Public Property Get sName() As String
  sName = modMisc.sFilename(mvarFilename, efpFileName)
End Property

Public Property Get sNameAndExtension() As String
  sNameAndExtension = modMisc.sFilename(mvarFilename, efpFileNameAndExt)
End Property
