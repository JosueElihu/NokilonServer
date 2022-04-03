VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_PATH                  As Long = 260
Private Const INVALID_HANDLE_VALUE      As Long = -1
Private Const FILE_ATTRIBUTE_DIRECTORY  As Long = &H10

Private Type FILETIME
  dwLowDateTime       As Long
  dwHighDateTime      As Long
End Type

Private Type WIN32_FIND_DATA
  dwFileAttributes    As Long
  ftCreationTime      As FILETIME
  ftLastAccessTime    As FILETIME
  ftLastWriteTime     As FILETIME
  nFileSizeHigh       As Long
  nFileSizeLow        As Long
  dwReserved0         As Long
  dwReserved1         As Long
  cFileName           As String * MAX_PATH
  cAlternate          As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Private Type LARGE_INTEGER
  LowPart     As Long
  HighPart    As Long
End Type


Private m_Path      As String
Private m_bDoEvt    As Boolean

Private m_File()    As String
Private m_Folder()  As String

Property Get Path() As String: Path = m_Path: End Property



Private Sub Class_Initialize()
'
End Sub
Private Sub Class_Terminate()
    Clear
End Sub

Public Function List(ByVal oPath As String, Optional CallDoEvents As Boolean) As Boolean
On Error Resume Next
Dim lResult         As Long
Dim lHandle     As Long
Dim tWFD        As WIN32_FIND_DATA
Dim sName       As String
Dim l           As Long

    
    If Right$(oPath, 1) <> "\" Then oPath = oPath & "\"
    m_Path = oPath
    Clear
    
    lHandle = FindFirstFile(oPath & "*", tWFD)
    If lHandle = -1 Then Exit Function
    
    Do
        
        sName = StripNulls(tWFD.cFileName)
        
        If sName = "." Or sName = ".." Then GoTo next_
        If ((tWFD.dwFileAttributes And vbSystem) = vbSystem) And ((tWFD.dwFileAttributes And vbHidden) = vbHidden) Then
            GoTo next_
        End If
        
        If (tWFD.dwFileAttributes And vbDirectory) <> vbDirectory Then
            
            l = FileCount
            ReDim Preserve m_File(l)
            m_File(l) = sName
            
        Else
        
            l = FolderCount
            ReDim Preserve m_Folder(l)
            m_Folder(l) = sName
            
        End If
        
next_:
        lResult = FindNextFile(lHandle, tWFD)
        If CallDoEvents Then DoEvents
        
    Loop While lResult
    
    Call FindClose(lHandle)
    
End Function

Public Sub Clear()
    Erase m_File
    Erase m_Folder
End Sub

Property Get FileCount() As Long
On Error GoTo e
    FileCount = UBound(m_File) + 1
e:
End Property
Property Get FolderCount() As Long
On Error GoTo e
    FolderCount = UBound(m_Folder) + 1
e:
End Property

Property Get FileName(ByVal Index As Long) As String
On Error GoTo e
    FileName = m_File(Index)
e:
End Property
Property Get Folder(ByVal Index As Long) As String
On Error GoTo e
    Folder = m_Folder(Index)
e:
End Property


Property Get FilePath(ByVal Index As Long) As String
On Error GoTo e
    FilePath = m_Path & m_File(Index)
e:
End Property
Property Get FolderPath(ByVal Index As Long) As String
On Error GoTo e
    FolderPath = m_Path & m_Folder(Index)
e:
End Property

Property Get Attr(ByVal Index As Long, Optional IsFolder As Boolean = True) As VbFileAttribute
On Error GoTo e
    
    If IsFolder Then
        Attr = GetAttr(m_Path & m_Folder(Index))
    Else
        Attr = GetAttr(m_Path & m_File(Index))
    End If
e:
End Property
'Property Let Attr(ByVal Index As Long, ByVal IsFolder As Boolean, ByVal Value As VbFileAttribute)
'On Error GoTo Err
'
'    If IsFolder Then
'        SetAttr m_Path & m_Folder(Index), Value
'    Else
'        SetAttr m_Path & m_File(Index), Value
'    End If
'
'Err:
'End Property
Property Get FileDate(ByVal Index As Long) As String
On Error GoTo e
   FileDate = FileDateTime(m_Path & m_File(Index))
e:
End Property
Property Get FolderDate(ByVal Index As Long) As String
On Error GoTo e
   FolderDate = FileDateTime(m_Path & m_Folder(Index))
e:
End Property



Private Function StripNulls(sData As String) As String
    StripNulls = Left$(sData, lstrlen(sData))
End Function

