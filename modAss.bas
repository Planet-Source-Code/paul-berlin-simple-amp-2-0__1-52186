Attribute VB_Name = "modAss"
'Association Routines 1.1
'------------------------
'By Paul Berlin 2002
'I snatched the registry stuff from my friend Davey Taylor, cause
'It was the most compact I could find. I fixed some bugs with it though.

Option Explicit

'// Registry access API
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Private Const Reg_HKCR As Long = &H80000000
Private Const Reg_HKCU As Long = &H80000001
Private Const Reg_HKLM As Long = &H80000002
Private Const Reg_HKU  As Long = &H80000003
Private Const Reg_HKPD As Long = &H80000004
Private Const Reg_HKCC As Long = &H80000005
Private Const Reg_HKDD As Long = &H80000006

Private Const RB  As Long = 1
Private Const RDW As Long = 4

'REGISTRY ROUTINES
'-----------------
Function CreateKey(ByVal hKey As Long, ByVal subKey As String) As Long
 On Error Resume Next
 Dim hReg As Long
 If RegCreateKey(hKey, subKey, hReg) = 0 Then
  CreateKey = hReg
 Else
  Err.Raise 335
 End If
End Function

Function OpenKey(ByVal hKey As Long, ByVal subKey As String) As Long
 On Error Resume Next
 Dim hReg As Long
 If RegOpenKey(hKey, subKey, hReg) = 0 Then
  OpenKey = hReg
 Else
  Err.Raise 335
 End If
End Function

Function GetDWord(ByVal hReg As Long, ByVal rValue As String) As Long
 On Error Resume Next
 Dim dwData As Long
 If RegQueryValueEx(hReg, rValue, 0, RDW, dwData, RDW) = 0 Then
  GetDWord = dwData
 Else
   Err.Raise 335
 End If
End Function

Sub SetString(ByVal hKey As Long, ByVal rValue As String, ByVal strData As String)
 On Error Resume Next
 Dim strLength As Long
 strLength = Len(strData)
 If RegSetValueEx(hKey, rValue, 0, RB, ByVal strData, strLength) <> 0 Then
  Err.Raise 335
 End If
End Sub

Sub SetDWord(ByVal hReg As Long, ByVal rValue As String, ByVal dwData As Long)
 On Error Resume Next
 If RegSetValueEx(hReg, rValue, 0, RDW, ByVal dwData, RDW) <> 0 Then
  Err.Raise 335
 End If
End Sub

Sub Delete(ByVal hReg As Long, ByVal rValue As String)
 On Error Resume Next
 If RegDeleteValue(hReg, rValue) <> 0 Then
  Err.Raise 335
 End If
End Sub

Function GetString(ByVal hReg As Long, ByVal rValue As String) As String
 On Error Resume Next
 Dim strData As String
 Dim strLength As Long

 strLength = 1024
 strData = Space$(strLength)

 If RegQueryValueEx(hReg, rValue, 0, RB, ByVal strData, strLength) = 0 Then
  If strLength > 0 Then GetString = Left$(strData, strLength - 1)
  'Theres lots of shit at the end of the string sometimes
  If InStr(1, strData, Chr(32)) > 0 Then GetString = Left$(strData, InStr(1, strData, Chr(32)) - 1)
  If InStr(1, strData, Chr(0)) > 0 Then GetString = Left$(strData, InStr(1, strData, Chr(0)) - 1)
 Else
  Err.Raise 335
 End If
End Function

Sub CloseKey(ByVal hReg As Long)
 On Error Resume Next
 If RegCloseKey(hReg) <> 0 Then
  Err.Raise 335
 End If
End Sub

'ASSOCIATION SUBROUTINES
'-----------------------

'This creates an file type, for use with CreateAssociation
'------------------------------------------
'sName = Name of file type, ex. "Winamp.mp3"
'sDescr = Decription of type, ex. "Mpeg Layer 3"
'sDefaultIcon = Icon of type, ex. "C:\prg\winamp\winamp.exe,0" (the number being the number of the icon embedded in the exe)
Public Sub CreateFileType(ByVal sName As String, ByVal sDescr As String, ByVal sDefaultIcon As String)
  Dim lKey As Long
  Dim lKey2 As Long
  On Error Resume Next
  
  lKey = CreateKey(Reg_HKCR, sName)
  Call SetString(lKey, "", sDescr)
  lKey2 = CreateKey(lKey, "DefaultIcon")
  Call SetString(lKey2, "", sDefaultIcon)
  Call CloseKey(lKey)
  Call CloseKey(lKey2)
  
End Sub

'This creates an action for the filetype
'------------------------------------------
'sTypeName = Name of the filetype to add this to, ex. "Winamp.mp3"
'sName = Name of action when you rightclick on file, ex. "Open"
'sCmd = Command for this action, ex. "C:\prg\winamp\winamp.exe -open %1"
Public Sub CreateFileTypeAction(ByVal sTypeName As String, ByVal sName As String, ByVal sCmd As String)
  Dim lKey As Long
  Dim lKey2 As Long
  On Error Resume Next
  
  lKey = OpenKey(Reg_HKCR, sTypeName)
  lKey2 = CreateKey(lKey, "shell\" & sName & "\command")
  Call SetString(lKey2, "", sCmd)
  Call CloseKey(lKey)
  Call CloseKey(lKey2)
  
End Sub

'This creates an association with an extention to an earlier created filetype
'------------------------------------------
'sExtention = Extention of files to associate with, ex. ".mp3"
'sBackupName = The name of the backup entry, ex. "Winamp_bak"
'sFileType = the Filetype created with CreateFileType to use, ex. "Winamp.mp3"
Public Sub CreateAssociation(ByVal sExtention As String, ByVal sBackupName As String, ByVal sFileType As String)
  Dim lKey As Long
  On Error Resume Next
  
  lKey = CreateKey(Reg_HKCR, sExtention)
  If GetString(lKey, "") <> sFileType And Len(GetString(lKey, "")) > 0 Then 'If the current value is not sFileType
    Call SetString(lKey, sBackupName, GetString(lKey, "")) 'create backup
  End If
  Call SetString(lKey, "", sFileType)
  Call CloseKey(lKey)
  
End Sub

'This restores the association to the one in the backup
'------------------------------------------
'sExtention = Extention of files to associate with, ex. ".mp3"
'sBackupName = The name of the backup entry, ex. "Winamp_bak"
Public Sub RestoreAssociation(ByVal sExtention As String, ByVal sBackupName As String)
  Dim lKey As Long
  On Error Resume Next

  lKey = OpenKey(Reg_HKCR, sExtention)
  Call SetString(lKey, "", GetString(lKey, sBackupName)) 'restore backup
  Call Delete(lKey, sBackupName)
  Call CloseKey(lKey)
  
End Sub

'This checks if association of extention is sfilename
'------------------------------------------
'sExtention = Extention of files to associate with, ex. ".mp3"
'sFileType = the Filetype created with CreateFileType to use, ex. "Winamp.mp3"
Public Function CheckAssociation(ByVal sExtention As String, ByVal sFileType As String)
  Dim lKey As Long
  On Error Resume Next
  
  lKey = CreateKey(Reg_HKCR, sExtention)
  If GetString(lKey, "") = sFileType Then CheckAssociation = True
  Call CloseKey(lKey)
  
End Function
