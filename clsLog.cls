VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'By Paul Berlin 2003, update 03:07 2003-09-01
'Small class with some basic logging and error helper functions
'You can also easily measure time for benchmarking your app

'NOTE: All logging start disabled so be sure to enable if wanted

Option Explicit

'Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Public bLogFile As Boolean 'Do log to file?
Public bLogDebug As Boolean 'Do log to immediate window

Private iFileNum As Long
Private lTimer As Long

Public bAutoErrorLog As Boolean 'Log Errors when using ErrorMsg sub automatically
Public bAppendFile As Boolean 'Append to logfile?
Public sLogFile As String 'log filename
Public iLogFileLevel As Integer '1-10 minimum 'importance-level' of messages to display
Public iLogDebugLevel As Integer '1-10

Public Sub Log(ByVal sText As String, Optional ByVal lImportance As Integer = 10, Optional bNewLine As Boolean = True)
  
  If bLogDebug And lImportance >= iLogDebugLevel Then
    If bNewLine Then
      Debug.Print sText
    Else
      Debug.Print sText;
    End If
  End If
  
  If iFileNum = 0 And bLogFile Then
    iFileNum = FreeFile
    If bAppendFile And Not (Dir(sLogFile) = "") Then
      Open sLogFile For Append As iFileNum
    Else
      Open sLogFile For Output As iFileNum
      Print #iFileNum, App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision & " logfile created " & Now
      Print #iFileNum, String(Len(App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision & " logfile created " & Now), "=")
    End If
  End If
  
  If bLogFile And iFileNum <> 0 And lImportance >= iLogFileLevel Then
    If bNewLine Then
      Print #iFileNum, sText
    Else
      Print #iFileNum, sText;
    End If
  End If
  
End Sub

Private Sub Class_Initialize()
  sLogFile = App.Path & "\" & App.EXEName & ".log"
  iLogFileLevel = 5
  iLogDebugLevel = 1
End Sub

Private Sub Class_Terminate()
  If iFileNum <> 0 Then Close iFileNum
End Sub

Public Sub StartTimer()
  lTimer = timeGetTime
End Sub

Public Function ErrorMsg(ByRef oErr As ErrObject, Optional ByVal sSource As String) As VbMsgBoxResult
  Dim sMsg As String
  
  If Len(sSource) = 0 Then sSource = Err.Source
  
  sMsg = "The following application error has occurred:" & vbCrLf & vbCrLf
  sMsg = sMsg & "Error Number:" & vbTab & Err.Number & vbCrLf
  sMsg = sMsg & "Description:" & vbTab & Err.Description & vbCrLf & vbCrLf
  sMsg = sMsg & "Last DLL Error:" & vbTab & Err.LastDllError & vbCrLf
  sMsg = sMsg & "Source:" & vbTab & vbTab & sSource & vbCrLf & vbCrLf
  sMsg = sMsg & "Do you want to resume program execution?"
  
  If bAutoErrorLog Then
    Log "APPLICATION ERROR # " & Err.Number & ", at " & Now, 10
    Log "ERROR SOURCE: " & sSource, 10
    Log "ERROR DESCRIPTION: " & Err.Description, 10
  End If
  
  ErrorMsg = MsgBox(sMsg, vbYesNo Or vbDefaultButton1 Or vbCritical, "Application Error")
  
  If bAutoErrorLog Then
    Log "ERROR RESPONSE: " & IIf(ErrorMsg = vbYes, "RESUME EXECUTION.", "TERMINATE EXECUTION."), 10
  End If
  
End Function

Public Function GetTimer() As Long
  GetTimer = timeGetTime - lTimer
End Function
