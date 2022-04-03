Attribute VB_Name = "mFileIO"
'--------------------------------------------------------------------------------
'    Component  : mFileIO
'    Autor      : J. Elihu
'    Description: Can handle file sizes greater than 2gb.
'    Modified   : 13/09/2021
'--------------------------------------------------------------------------------
Option Explicit


Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Any) As Long
    
Private Declare Function PathFileExistsA Lib "shlwapi" (ByVal pszPath As String) As Long

Private Const FILE_ATTRIBUTE_NORMAL   As Long = &H80
Private Const FILE_BEGIN = 0, FILE_CURRENT = 1, FILE_END = 2

Private Const FILE_SHARE_READ         As Long = &H1
Private Const FILE_SHARE_WRITE        As Long = &H2
Private Const FILE_FLAG_WRITE_THROUGH As Long = &H80000000
Private Const GENERIC_READ            As Long = &H80000000
Private Const GENERIC_WRITE           As Long = &H40000000
Private Const OPEN_EXISTING           As Long = 3
Private Const OPEN_ALWAYS             As Long = 4
Private Const INVALID_HANDLE_VALUE    As Long = -1
Private Const MAX_LONG                As Long = 2147483647

Private Type LARGE_INTEGER
  LowPart   As Long
  HighPart  As Long
End Type



Public Function Open_(ByRef sFileName As String, Optional ByVal ReadOnly As Boolean, Optional lErr As Long) As Long

    If ReadOnly Then
        If PathFileExistsA(sFileName) = 0 Then Exit Function
        Open_ = CreateFile(sFileName, GENERIC_READ, FILE_SHARE_READ, 0&, OPEN_EXISTING, 0, 0)
    Else
        Open_ = CreateFile(sFileName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL Or FILE_FLAG_WRITE_THROUGH, 0)
       'Open_ = CreateFile(sFileName, GENERIC_WRITE Or GENERIC_READ, 0, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    End If
    
    If Open_ = INVALID_HANDLE_VALUE Then
        '123 El nombre de archivo, el nombre de directorio o la sintaxis de la etiqueta del volumen no son correctos.
        lErr = Err.LastDllError
        Debug.Print "mFileIO::Open: " & DecodeAPIErrors(lErr)
        PushLog "FileIO::Open: " & DecodeAPIErrors(lErr)
        CloseHandle Open_
        Open_ = 0
    End If
    'SetFilePointer m_hFile, 0, 0, FILE_BEGIN
    
End Function
Public Function FileClose_(lFile As Long)
    If lFile <> 0 Then
        CloseHandle lFile
        lFile = 0
    End If
End Function


Public Function Read_(ByRef lFile As Long, ByRef Out() As Byte, Optional Pos As Variant) As Long
On Error GoTo e
    
    If Not IsMissing(Pos) Then Seek_ lFile, Pos
    If ReadFile(lFile, Out(LBound(Out)), UBound(Out) - LBound(Out) + 1, Read_, 0&) Then
        'Debug.Print Read_
    End If
e:
    On Error GoTo 0   ' Nullify this error trap
End Function
Public Function Write_(ByRef lFile As Long, ByRef Out() As Byte, Optional Pos As Variant, Optional ByVal mbFlush As Boolean) As Long
On Error GoTo e

    If Not IsMissing(Pos) Then Seek_ lFile, Pos
    If WriteFile(lFile, Out(LBound(Out)), UBound(Out) - LBound(Out) + 1, Write_, 0&) Then
        If mbFlush Then Flush_ lFile
        'Debug.Print WriteBytes
    End If
e:
    On Error GoTo 0   ' Nullify this error trap
End Function

Public Function Flush_(ByRef lFile As Long) As Boolean
   Flush_ = FlushFileBuffers(lFile)
End Function
Public Function Seek_(ByRef lFile As Long, ByVal Value As Currency) As Boolean
    'If Not m_hFile <> 0 Then Exit Property
    
    Dim lInt As LARGE_INTEGER
    'Calculate current position within file
    Size2Long Value, lInt.LowPart, lInt.HighPart
    SetFilePointer lFile, lInt.LowPart, lInt.HighPart, FILE_BEGIN
    
End Function


Private Sub Size2Long(ByVal curFileSize As Currency, ByRef lngLowOrder As Long, ByRef lngHighOrder As Long)
Dim curCutoff As Currency

    curCutoff = CCur(MAX_LONG)
    curCutoff = curCutoff + MAX_LONG
    curCutoff = curCutoff + 1       ' now we hold the value of 4294967295 and not -1

    lngHighOrder = 0: lngLowOrder = 0

    Do Until curFileSize < curCutoff
        lngHighOrder = lngHighOrder + 1
        curFileSize = curFileSize - curCutoff
    Loop

    If curFileSize > MAX_LONG Then
        lngLowOrder = CLng(-(curCutoff - (curFileSize - 1)))   ' Larger than 2gb
    Else
        lngLowOrder = CLng(curFileSize)   ' Less than 2gb
    End If

End Sub
Private Function ULongToCurrency(ByVal Value As Long) As Currency
    If Value < 0 Then
        ULongToCurrency = CCur(Value And &H7FFFFFFF) + 2147483648#
    Else
        ULongToCurrency = CCur(Value)
    End If
End Function
Public Function DecodeAPIErrors(ByVal ErrorCode As Long) As String
    Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
    Dim tmp As String, ln As Long

    tmp = Space$(256)
    ln = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, ErrorCode, 0&, tmp, 256&, 0&)
    If ln > 0 Then DecodeAPIErrors = Left(tmp, ln - 2) Else DecodeAPIErrors = "Unknown Error."
    
End Function '&
