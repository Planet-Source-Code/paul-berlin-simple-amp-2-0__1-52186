VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This simple class reads/writes to ini-files
'You have to set up sFilename and sSection before reading or writing.
'By Paul Berlin 2003
Option Explicit

Private Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private slFilename As String
Private slSection As String

Public Property Get sFilename() As String
  sFilename = slFilename
End Property

Public Property Let sFilename(ByVal sFile As String)
  slFilename = sFile
End Property

Public Property Get sSection() As String
  sSection = slSection
End Property

Public Property Let sSection(ByVal sSect As String)
  slSection = sSect
End Property

Public Sub WriteKey(ByVal sKey As String, ByVal vValue As Variant)
  If Len(slFilename) > 0 And Len(slSection) > 0 Then
    WritePrivateProfileString slSection, sKey, CStr(vValue), slFilename
  End If
End Sub

Public Function ReadNumber(ByVal sKey As String, Optional ByVal nDefault As Single) As Single
  If Len(slFilename) > 0 And Len(slSection) > 0 Then
    Dim sTemp As String
    sTemp = ReadString(sKey, CStr(nDefault))
    
    'When using some language settings in windows a Single number like 75.34
    'will be written 75,34. This will not work when reading the number later,
    'which will be returned as just 75. Therefore I am reading all numbers
    'as strings and replacing all commas with periods, so it will work.
    sTemp = Replace(sTemp, ",", ".")
    
    ReadNumber = Val(sTemp)
    
  End If
End Function

Public Function ReadString(ByVal sKey As String, Optional ByVal sDefault As String) As String
  If Len(slFilename) > 0 And Len(slSection) > 0 Then
    Dim sBuffer As String * 255
    
    GetPrivateProfileString slSection, sKey, CStr(sDefault), sBuffer, Len(sBuffer), slFilename
    
    If InStr(sBuffer, Chr(0)) > 0 Then
      ReadString = Trim(Left(sBuffer, InStr(sBuffer, Chr(0)) - 1))
    Else
      ReadString = Trim(sBuffer)
    End If
    
  End If
End Function
