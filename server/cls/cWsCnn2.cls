VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWsCnn2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'FTP data Socket Manager

Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Enum FTPTransferMode
    emNone
    emDownload
    emUpload
End Enum
Public Enum FTPResult
    emSuccess
    emError
    emCancel
End Enum

Private WithEvents mTimer  As cTimer
Attribute mTimer.VB_VarHelpID = -1

Public LISTENER_SOCKET  As Long             '/* Port Socket Handle      */
Public CONTROL_SOCKET   As Long             '/* Control Socket Handle   */
Public SOCKET_TYPE      As FtpSocketType    '/* Plug Type               */
Public SOCKET_ID        As Long
Public SOCKET_IP        As String
Public Mode             As FTPTransferMode

Private LEN_STREAM  As Long
Private m_Stream()  As Byte
Private m_Size      As Currency
Private m_Pos       As Currency
Private m_lFile     As Long

Private m_bCancel   As Boolean
Private m_bResume   As Boolean
Private m_StartTime As Long
Private m_TimeStamp As String


Private Sub Class_Initialize()
    SOCKET_TYPE = FTP_DATA_SOCKET
    Set mTimer = New cTimer
End Sub
Private Sub Class_Terminate()
On Error GoTo e

    Set mTimer = Nothing
    
    '/* To control Socket */
    If CONTROL_SOCKET <> 0 Then mcSockets(CStr(CONTROL_SOCKET)).SOCKET_DATA = 0: CONTROL_SOCKET = 0
  
    '/* Disconnect */
    If SOCKET_ID <> 0 Then mWinSock.WsDisconnect SOCKET_ID: SOCKET_ID = 0

    '/* Close opened port */
    If LISTENER_SOCKET <> 0 Then Call mWinSock.WsClose(LISTENER_SOCKET)
  
    '/* Close file */
    If m_lFile Then mFileIO.FileClose_ m_lFile: m_bCancel = False
e:
    Err.Clear
    Resume Next
End Sub

Property Get Speed() As String
On Error GoTo e
Dim KBDesc      As Currency
Dim TimeElapsed As Long

    If Mode = emUpload Then
        KBDesc = m_Size \ 1024
        'TimeElapsed = (GetTickCount - m_StartTime) \ 1000
        TimeElapsed = (GetTickCount - m_StartTime) '\ 1000
        Speed = KBDesc \ TimeElapsed
    Else
        'TimeElapsed = (GetTickCount - m_StartTime) \ 1000
        TimeElapsed = (GetTickCount - m_StartTime) '\ 1000
        Speed = m_Pos \ TimeElapsed
    End If
e:
End Property
Property Get CurrentBytes() As Currency
    Select Case Mode
        Case emDownload: CurrentBytes = m_Pos
        Case emUpload: CurrentBytes = m_Size
    End Select
End Property
Property Get TotalBytes() As Currency: TotalBytes = m_Size: End Property
Property Get Percent() As Long
On Error GoTo e:
    If Mode = emDownload Then Percent = CInt((m_Pos / m_Size) * 100)
e:
End Property
Property Get TimeStamp() As String: TimeStamp = m_TimeStamp: End Property

Private Sub mTimer_Timer(ByVal ThisTime As Long)
    mTimer.DestroyTimer
    Call SendParts
End Sub

Public Sub WSconnected(ID As Long, IP As String)
    SOCKET_ID = ID: SOCKET_IP = IP
    mcSockets(CStr(CONTROL_SOCKET)).SOCKET_DATA = ID
    mcSockets(CStr(CONTROL_SOCKET)).ResetTimeOut
End Sub
Public Sub WSarrival(ID As Long, data As String, ByVal lBytes As Long)
On Error GoTo e

    mcSockets(CStr(CONTROL_SOCKET)).ResetTimeOut
    Me.AddBytes lBytes, eBytesReceived
    
    If m_lFile = 0 Then Exit Sub
    mFileIO.Write_ m_lFile, StrConv(data, vbFromUnicode)
     
    m_Size = m_Size + lBytes

    Exit Sub
e:
    FrmMain.FTPtransferEnd SOCKET_ID, emUpload, m_Size, emError, 4, CONTROL_SOCKET
    mcSockets(CStr(CONTROL_SOCKET)).Send "550 " & Err.Description
    RemoveSH CONTROL_SOCKET
End Sub
Public Sub WSwritable(ID As Long)
     If Mode = emDownload And m_bResume Then SendParts
End Sub
Public Sub WSdisconnected(ID As Long)

    If Mode = emUpload Then
        FrmMain.FTPtransferEnd SOCKET_ID, emUpload, m_Size, IIf(m_bCancel, emCancel, emSuccess), 4, CONTROL_SOCKET
        mcSockets(CStr(CONTROL_SOCKET)).Send "226 File successfully upload."
    ElseIf Mode = emDownload Then
        If m_Pos < m_Size Then
            mcSockets(CStr(CONTROL_SOCKET)).Send "550 Send file canceled, connection closed by client."
            FrmMain.FTPtransferEnd ID, emDownload, m_Pos, emCancel, Percent, CONTROL_SOCKET '[MAIN EVENT CALLBACK]
        End If
    End If
    
    SOCKET_ID = 0
    If CONTROL_SOCKET Then
        mcSockets(CStr(CONTROL_SOCKET)).SOCKET_DATA = 0
        CONTROL_SOCKET = 0
    End If
    RemoveSH ID
    
End Sub


'TODO: File List Functions
'======================================================================================================================
Public Function SendList(ByVal Args As String, ByVal Path As String, cColl As Collection, Cmd As String)
On Error GoTo e
Dim tmp     As String
Dim i       As Long

    '-ARGS
    If LenB(Args) <> 0 Then
        If PathDirectory(Args) Then tmp = mvList(Args, Cmd, i) Else tmp = mvFileI(Args, Cmd)
        GoTo m
    End If
    
    '-ROOT
    If Path = vbNullString Then tmp = mvRootList(cColl, Cmd, i): GoTo m

    If Not PathExist(Path) Then SendList = 550: Exit Function
    tmp = mvList(Path, Cmd, i)
m:
    If LenB(tmp) Then
        If mWinSock.WsSend(SOCKET_ID, tmp) Then SendList = 226
    Else
        SendList = 226
    End If
e:
     If Err.Number Then Debug.Print "WsData::SendList: " & Err.Description
End Function



'TODO: File Transfer Functions
'======================================================================================================================
Public Function PrepareFile(ByRef FileName As String, Optional lMode As FTPTransferMode, Optional Reset As Boolean, Optional Pos As Currency) As Boolean
On Error GoTo e
Dim lErr    As Long

    If lMode = emUpload Then m_Size = 0: GoTo m
 
    LEN_STREAM = 1024& * SERVER.STREAM_SIZE&
    ReDim m_Stream(1 To LEN_STREAM)
    m_Size = FileLen2(FileName)
    
    If Reset Then
        m_Pos = Pos
        If m_Pos < 0 Then m_Pos = 0
        If m_Pos > m_Size Then m_Pos = m_Size
    Else
        m_Pos = 0
    End If
    Pos = 0
m:
    Mode = lMode
    m_bCancel = False
    
Open_:

    m_lFile = mFileIO.Open_(FileName, lMode = emDownload, lErr)
    If m_lFile = 0 And lErr = 123 Then FileName = GetSafeFileName(FileName): m_lFile = mFileIO.Open_(FileName, , lErr)
    If m_lFile = 0 Then Exit Function
    
    m_StartTime = GetTickCount
    m_TimeStamp = Format(Now, "hh:mm:ss - dd/mm/yy")
    PrepareFile = True
    
    Exit Function
e:
    PushLog "WsData::PrepareFile: " & Err.Description, enmPSError, mclERRORS
End Function

Public Sub SendParts()
On Error GoTo e
Dim lErr    As WS_ERROR
Dim lBytes  As Long

    If m_lFile = 0 Then GoTo e
    If m_bCancel Then GoTo e

    'If (m_Pos + LEN_STREAM) > m_Size Then ReDim m_Stream(m_Size - m_Pos)
    If (m_Pos + LEN_STREAM) > m_Size Then ReDim m_Stream(m_Size - m_Pos - 1)
    
    mFileIO.Read_ m_lFile, m_Stream, m_Pos
     
    '/To main thread
    mcSockets(CStr(CONTROL_SOCKET)).ResetTimeOut
    
    m_bResume = False
    If mWinSock.WsSend2(SOCKET_ID, m_Stream, lBytes, lErr) = False Then
       If lErr <> WSAEWOULDBLOCK Then GoTo e
       m_Pos = m_Pos + lBytes
       m_bResume = True
       Exit Sub
    End If
    
    m_Pos = m_Pos + LEN_STREAM
    If Not (m_Pos > m_Size - 1) Then mTimer.CreateTimer 1: Exit Sub

    mcSockets(CStr(CONTROL_SOCKET)).Send "226 Successfully transferred File"
    FrmMain.FTPtransferEnd SOCKET_ID, emDownload, m_Size, emSuccess, Percent, CONTROL_SOCKET '[MAIN EVENT CALLBACK]
    RemoveSH SOCKET_ID
    Exit Sub
e:
    FrmMain.FTPtransferEnd SOCKET_ID, emDownload, m_Pos, IIf(m_bCancel, emCancel, emError), Percent, CONTROL_SOCKET '[MAIN EVENT CALLBACK]
    If ExistSH(CONTROL_SOCKET) Then
        Dim tmp As String
        If Err.Number Then tmp = "502 Failed to send file: " & Err.Description & ".": PushLog "WsData::SendFile: " & Err.Description, enmPSError, mclERRORS
        If m_bCancel Then tmp = "550 Send file canceled."
        If tmp = vbNullString Then tmp = "550 Failed to send file."
        mcSockets(CStr(CONTROL_SOCKET)).Send tmp
    End If
    RemoveSH CONTROL_SOCKET
End Sub

Public Sub Cancel()
    If Mode = 0 Then Exit Sub
    m_bCancel = True
    Select Case Mode
        Case emDownload
            SendParts 'WARNING: Raise cancel now
        Case emUpload
            FrmMain.FTPtransferEnd SOCKET_ID, emUpload, m_Size, emCancel, 4, CONTROL_SOCKET
            RemoveSH SOCKET_ID
            mcSockets(CStr(CONTROL_SOCKET)).Send "550 file upload canceled"
    End Select
End Sub

Public Function CreateWS(Optional IPv6 As Boolean) As Long
    LISTENER_SOCKET = mWinSock.WsListen(GetFreePortNum(2000), IPv6)
    If LISTENER_SOCKET <> 0 Then CreateWS = LISTENER_SOCKET
End Function

Public Function Dump(obj As Object)
On Error GoTo e

    obj.SOCKET_ID = SOCKET_ID
    SOCKET_ID = 0
    If obj.SOCKET_TYPE = FTP_DATA_SOCKET Then
        obj.CONTROL_SOCKET = CONTROL_SOCKET
        obj.LISTENER_SOCKET = LISTENER_SOCKET
        CONTROL_SOCKET = 0
        LISTENER_SOCKET = 0
    End If
e:
End Function

Public Sub AddBytes(Bytes As Long, Optional lpType As enmByteType = 1)
On Error GoTo e
    If CONTROL_SOCKET <> 0 Then mcSockets(CStr(CONTROL_SOCKET)).AddBytes Bytes, lpType     '/INCREASE BYTES
e:
End Sub


'TODO: Private Sub's
'======================================================================================================================
Private Function mvList(sPath As String, Cmd As String, i As Long) As String
    With FileList(sPath)
        For i = 0 To .FolderCount - 1
             Select Case Cmd
                Case "LIST"
                   'mvList = mvList & "drwxrwxr-x 1 ftp ftp " & mvSpc(0) & " " & FTPDate(.FolderDate(i)) & " " & ToUTF8(.Folder(i)) & vbCrLf
                   'mvList = mvList & "drwx------ 1 ftp ftp " & mvSpc(0) & " " & FTPDate(.FolderDate(i)) & " " & ToUTF8(.Folder(i)) & vbCrLf
                    mvList = mvList & "drwxrwxrwx 1 ftp ftp " & mvSpc(0) & " " & FTPDate(.FolderDate(i)) & " " & ToUTF8(.Folder(i)) & vbCrLf
                Case "NLST"
                   mvList = mvList & ToUTF8(.Folder(i)) & vbCrLf
                Case "MLSD"
                   mvList = mvList & "type=dir;modify=" & Format(.FolderDate(i), "YYYYMMDDhhmmss") & ";perms=cplemdf; " & ToUTF8(.Folder(i)) & vbCrLf
             End Select
        Next
        
        For i = 0 To .FileCount - 1
            Select Case Cmd
                Case "LIST"
                   'mvList = mvList & "-rwx------ 1 ftp ftp " & mvSpc(FileLen2(.FilePath(i))) & " " & FTPDate(.FileDate(i)) & " " & ToUTF8(.FileName(i)) & vbCrLf
                    mvList = mvList & "-rw-rw-rw- 1 ftp ftp " & mvSpc(FileLen2(.FilePath(i))) & " " & FTPDate(.FileDate(i)) & " " & ToUTF8(.FileName(i)) & vbCrLf
                Case "NLST"
                    mvList = mvList & ToUTF8(.FileName(i)) & vbCrLf
                Case "MLSD"
                    mvList = mvList & "type=file;size=" & FileLen2(.FilePath(i)) & ";modify=" & Format(.FileDate(i), "YYYYMMDDhhmmss") & ";perms=awrdf; " & ToUTF8(.FileName(i)) & vbCrLf
            End Select
        Next
    End With
End Function

Private Function mvRootList(cColl As Collection, Cmd As String, i As Long) As String

    '/* {name, path, access} */
    For i = 1 To cColl.Count
        Select Case Cmd
            Case "LIST"
                'mvRootList = mvRootList & "drwx------ 1 ftp ftp " & mvSpc(0) & " " & FTPDate(Now) & " " & ToUTF8(cColl(i)(0)) & vbCrLf
                mvRootList = mvRootList & "drwxrwxrwx 1 ftp ftp " & mvSpc(0) & " " & FTPDate(Now) & " " & ToUTF8(cColl(i)(0)) & vbCrLf
            Case "NLST"
                mvRootList = mvRootList & ToUTF8(cColl(i)(0)) & vbCrLf
            Case "MLSD"
                mvRootList = mvRootList & "type=dir;modify=" & Format(Now, "YYYYMMDDhhmmss") & ";perms=cplemdf; " & ToUTF8(cColl(i)(0)) & vbCrLf
        End Select
    Next
        
    
    '/NT FILE
    Select Case Cmd
        Case "LIST"
           'mvRootList = mvRootList & "-rw-rw-rw- 1 ftp ftp " & mvSpc(0) & " " & FTPDate(Now) & " " & ToUTF8("NO SUBIR ARCHIVOS AQUI") & vbCrLf
            mvRootList = mvRootList & "-r--r--r-- 1 ftp ftp " & mvSpc(0) & " " & FTPDate(Now) & " " & ToUTF8("NO SUBIR ARCHIVOS AQUI") & vbCrLf
        Case "NLST"
            mvRootList = mvRootList & ToUTF8("NO SUBIR ARCHIVOS AQUI") & vbCrLf
        Case "MLSD"
            mvRootList = mvRootList & "type=file;size=0;modify=" & Format(Now, "YYYYMMDDhhmmss") & ";perms=awrdf; " & ToUTF8("NO SUBIR ARCHIVOS AQUI") & vbCrLf
    End Select
    
End Function

Private Function mvFileI(sPath As String, Cmd As String) As String
    Select Case Cmd
        Case "LIST"
            mvFileI = "-rw-rw-rw- 1 ftp ftp " & mvSpc(FileLen2(sPath)) & " " & FTPDate(FileDateTime2(sPath)) & " " & ToUTF8(GetFileName(sPath)) & vbCrLf
        Case "NLST"
            mvFileI = "/" & ToUTF8(GetFileName(sPath)) & vbCrLf
        Case "MLSD"
            mvFileI = "type=file;size=" & FileLen2(sPath) & ";modify=" & Format(FileDateTime2(sPath), "YYYYMMDDhhmmss") & ";perms=awrdf; " & ToUTF8(GetFileName(sPath)) & vbCrLf
    End Select
End Function

Private Function mvSpc(data As String) As String
    mvSpc = String$(15 - Len(data), " ") & data
End Function

