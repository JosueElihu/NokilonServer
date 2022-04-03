VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function GetLongPathName Lib "kernel32.dll" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long


Private m_FNum  As Long
Private m_File  As String

Private Sub Class_Initialize()
    m_File = TempFile
    'If PathFileExists(m_File) Then CleanFile
    m_FNum = FreeFile
    Open m_File For Append Access Write As #m_FNum
    
End Sub
Private Sub Class_Terminate()
    If m_FNum Then Close #m_FNum: m_FNum = 0
    'CleanFile
End Sub

Property Get FileName() As String: FileName = m_File: End Property
Public Sub PushLog(sData As String)
    If m_FNum = 0 Then Exit Sub
    'Write #m_FNum, sData
    Print #m_FNum, sData
     
End Sub




Private Function TempFile() As String
On Error Resume Next

    If Environ$("temp") <> vbNullString Then
       Dim tmp As String
       tmp = String(255, 0)
       Call GetLongPathName(Environ$("Temp"), tmp, 255)
       TempFile = StripNull(tmp)
    Else
        TempFile = App.Path
    End If
     
    If Right$(TempFile, 1) <> "\" Then TempFile = TempFile & "\"
    TempFile = TempFile & App.EXEName & "-" & App.ThreadID & "-" & App.hInstance & ".log"
    
End Function

Private Function StripNull(sName As String) As String
    StripNull = Left$(sName, InStr(sName, Chr$(0)) - 1)
End Function

Private Function CleanFile()
On Error GoTo e
    Kill m_File
e:
End Function


