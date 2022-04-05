VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWsCnn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'   FTP Control Socket Helper
'-------------------------------------------------------------------------------------------------
Option Explicit

Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public SOCKET_ID     As Long
Public SOCKET_IP     As String
Public SOCKET_DATA   As Long            '/* Data Socket Handle   */
Public SOCKET_TYPE   As FtpSocketType   '/* Plug type            */
Public SOCKET_FAMILY As AddressFamilies

Private mvUser      As mtUSER
Private m_Mounts    As Collection

Private m_Path      As String       ' Local Path
Private m_url       As String       ' Virtual Path
Private m_RnmFrom   As String
Private m_Timeout   As Long
Private m_Index     As Long

'/ TRANSFER FLAGS
Private mtReset     As Boolean
Private mtPos       As Currency
Private OnlyEPSV    As Boolean

'/TRANSFER Bytes
Private m_lBytes1   As Currency         ' Bytes Received
Private m_lBytes2   As Currency         ' Bytes Send


Private Sub Class_Initialize()
    mvUser.m_Name = "Not logged on"
    Set m_Mounts = New Collection
End Sub
Private Sub Class_Terminate()

    If SOCKET_DATA <> 0 Then RemoveSH SOCKET_DATA
    If SOCKET_ID <> 0 Then
        mWinSock.WsDisconnect SOCKET_ID
        If mvUser.m_logged Then FrmMain.FTPuserlogout SOCKET_ID
    End If
    Set m_Mounts = Nothing
    
End Sub

Property Get UserName() As String: UserName = mvUser.m_Name: End Property
Property Get UserID() As String: UserID = mvUser.m_id: End Property
Property Get Bytes(ByVal lBType As enmByteType) As Currency
    Select Case lBType
        Case 1: Bytes = m_lBytes1
        Case 2: Bytes = m_lBytes2
    End Select
End Property

Public Sub WSconnected(ID As Long, IP As String)
    SOCKET_ID = ID
    SOCKET_IP = IP
    SOCKET_FAMILY = mWinSock.SocketFamily(ID)
    Send "220 " & SERVER_NAME
    ResetTimeOut
End Sub

Public Sub WSarrival(ID As Long, data As String, ByVal lBytes As Long)
Dim Cmd  As String
Dim Args As String

Dim argv()  As String
Dim i       As Long

    ResetTimeOut
    m_lBytes1 = m_lBytes1 + lBytes   'Increase Bytes
    
    argv = Split(data, vbCrLf)
    If argc(argv) > 0 Then
        For i = 0 To argc(argv) - 1
            Call mvParseArgs(argv(i), Cmd, Args)
            Call mvProcessCommand(Cmd, Args, ID)
        Next
    Else
        If argbb(argv) Then
            Call mvParseArgs(argv(0), Cmd, Args)
            Call mvProcessCommand(Cmd, Args, ID)
        Else
            Call Send("501 Syntax error")
        End If
    End If
End Sub
Public Sub WSwritable(ID As Long)
    '
End Sub
Public Sub WSdisconnected(ID As Long)
    SOCKET_ID = 0
    If mvUser.m_logged Then FrmMain.FTPuserlogout ID
    If SOCKET_DATA <> 0 Then RemoveSH SOCKET_DATA
    RemoveSH ID
End Sub

Public Sub CloseSocket(Optional Notify As Boolean, Optional ByVal MSGIndex As Long)
    If SOCKET_DATA <> 0 Then mcSockets(CStr(SOCKET_DATA)).Cancel
    If Notify Then
        Select Case MSGIndex
            Case 0: Send "421 Kicked by Administrator"
            Case 1: Send "421 Server is going offline"
            Case 2: Send "421 Connection timed out."
            Case 3: Send "421 No-transfer-time exceeded. Closing control connection."
            Case 4: Send "421 Login time exceeded. Closing control connection."
        End Select
    End If
    RemoveSH SOCKET_ID
End Sub

Public Function Send(data As String) As Boolean
    If Not mWinSock.WsSend(SOCKET_ID, data & vbCrLf) Then CloseSocket: Exit Function
    FrmMain.FTPcommandSent SOCKET_ID, data, mvUser.m_Name
    ResetTimeOut
End Function

Public Function ResetTimeOut(): m_Timeout = GetTickCount: End Function
Public Sub AddBytes(Bytes As Long, Optional lpType As enmByteType = 1)
    'TODO: Call To Increase Bytes
    Select Case lpType
        Case 1: m_lBytes1 = m_lBytes1 + Bytes   'Received
        Case 2: m_lBytes2 = m_lBytes2 + Bytes   'Sent
    End Select
End Sub
Public Sub UpdateMounts()
    LoadMountPoints mvUser.m_id, m_Mounts
End Sub




'TODO: Private subs
'---------------------------------------------------------------------------------------------------------------------

Private Function mvRootPath() As Boolean
    m_url = "/"
    m_Path = vbNullString
End Function
Private Function mvParseLocalPath(urlp As String) As String
    mvParseLocalPath = IIf(Left(urlp, 1) <> "/", m_url & "/" & urlp, urlp)
    mvParseLocalPath = ParseMountPath(mvParseLocalPath, m_Mounts, m_Index)
End Function
Private Function mvWaitDataSocket() As Boolean
    ResetTimeOut
    Do While SOCKET_DATA = 0
        DoEvents
        If m_Timeout + 80000 < GetTickCount Then Exit Do
    Loop
    mvWaitDataSocket = SOCKET_DATA <> 0
    If mvWaitDataSocket = False Then Send "425 No data connection": PushLog "WARNING: No data connection   WS:" & SOCKET_ID, enmPSNone, mclERRORS
    
End Function
Private Function mvParseArgs(data As String, Cmd As String, Args As String) As Boolean
Dim i    As Long
    i = InStr(data, " ")
    If i Then
        Cmd = UCase$(Left(data, i - 1))
        Args = ToUnicode(Mid(data, i + 1, Len(data) - i))
    Else
        Cmd = UCase$(data)
        Args = vbNullString
    End If
End Function
Private Sub mvProcessCommand(Cmd As String, Args As String, Optional WS As Long)
Dim tmp  As String
Dim tmp2 As String
Dim obj  As cWsCnn2
Dim i    As Long

    FrmMain.FTPcommandReceived WS, Cmd & " " & Args, mvUser.m_Name
    If CmdNeesdAuth(Cmd) And mvUser.m_logged = False Then Send "530 Please log in with USER and PASS first.": Exit Sub
    
    Select Case Cmd
        Case "USER"

            Select Case Login(Args, mvUser)
                Case LOGIN_INVALID: Send "530 user incorrect!": CloseSocket
                Case LOGIN_LOCKED: Send "530 Not logged in, user account has been disabled": CloseSocket
                Case LOGIN_PASSWORD_REQUIRED: Send "331 Password required for " & Args
                Case LOGIN_SUCCESS
                    
                    LoadMountPoints mvUser.m_id, m_Mounts
                    GetHomeMount m_url, m_Path
                    Send "230 Logged on"
                    FrmMain.FTPuserlogin WS, SOCKET_IP, mvUser.m_Name, m_lBytes1, m_lBytes2
            End Select
            
        Case "PASS"
        
            If mvUser.m_Name = vbNullString Then Send "530 Please log in with USER and PASS first.": Exit Sub
            If mvUser.m_logged Then Send "503 Bad sequence of commands.": Exit Sub
            If Login(Args, mvUser) <> LOGIN_SUCCESS Then Send "530 password incorrect!": Exit Sub
            
            LoadMountPoints mvUser.m_id, m_Mounts
            GetHomeMount m_url, m_Path
            Send "230 Logged on"
            FrmMain.FTPuserlogin WS, SOCKET_IP, mvUser.m_Name, m_lBytes1, m_lBytes2
          
        Case "SYST": Send "215 UNIX emulated by " & SERVER_NAME
        Case "NOOP", "NOP": Send "200 OK"
        Case "OPTS"
            Select Case UCase$(Args)
                Case "UTF-8 ON", "UTF8 ON": Send "202 UTF8 mode is always enabled. No need to send this command."
                Case "UTF-8 OFF", "UTF8 OFF": Send "504 UTF8 mode cannot be disabled."
                Case Else: Send "501 Option not understood"
            End Select
            
        Case "TYPE"
        
            Select Case Args
                Case "A", "I": Send "200 Type set to " & Args     ' A = ASCII, I = Binario, L 8 = ??
                Case Else: Send "501 Unsupported type. Supported types are I and A"
            End Select
            
         Case "PWD" '/* Print Working Directory */
        
            tmp = ParseUrl(m_url)
            If Not tmp <> "" Then Send "550 Directory could not be accessed": Exit Sub
            Send "257 " & Quot(ToUTF8(tmp))
            
        Case "CWD" '/* Change Working Directory */
        
            If Args = "/" Then mvRootPath: Send "250 CWD successful": Exit Sub
            
            tmp = m_Path: tmp2 = m_url
            If Left(Args, 1) = "/" Then
                m_url = Left(Args, Len(Args) - IIf(Right$(Args, 1) <> "/", 0, 1)) ' \Unidad\ -> \Unidad
            Else
                m_url = IIf(m_url <> "/", m_url & "/" & Args, "/" & Args)  ' & "/"
            End If
            
            m_Path = ParseMountPath(m_url, m_Mounts, m_Index)
            If PathExist(m_Path) Then
                Send "250 CWD successful."
            Else
                Send "550 CWD failed."
                m_Path = tmp: m_url = tmp2
            End If
            
        Case "CDUP" '/* Go up one directory */
        
            If m_url = "/" Then Send "200 CDUP successful.": Exit Sub
            
            tmp = ParentVirtualPath(m_url)
            If Len(tmp) = 0 Then tmp = "/"
            If tmp = "/" Then mvRootPath: Send "200 CDUP successful.": Exit Sub
            
            tmp2 = ParseMountPath(tmp, m_Mounts, m_Index)
            
            If Not PathExist(tmp2) Then Send "550 CDUP failed": Exit Sub
            m_url = tmp: m_Path = tmp2
            Send "200 CDUP successful."
            
        Case "PASV" '/* PASIVE MODE */
        
            If OnlyEPSV Then Send "500 Illegal EPRT command, EPSV ALL in effect.": Exit Sub
            If SOCKET_FAMILY = AF_INET6 Then Send "You are connected using IPv6. PASV is only for IPv4. You have to use the EPSV command instead.": Exit Sub
            
            '/* CREATE WS */
            Set obj = New cWsCnn2
            If obj.CreateWS() Then
                obj.CONTROL_SOCKET = SOCKET_ID
                mcSockets.Add obj, "#" & CStr(obj.LISTENER_SOCKET)
                Send "227 Entering Passive Mode " & MakePasvReply(mWinSock.SocketLocalPort(obj.LISTENER_SOCKET))
            Else
                Set obj = Nothing
                'Send "421 Could not create socket, unable to query socket for used port."
                Send "425 Cannot prepare for data connection."
            End If

        Case "EPSV" '/* EXTENED PASIVE MODE */
                
            If UCase$(Args) = "ALL" Then OnlyEPSV = True: Send "200 EPSV ALL command successful.": Exit Sub
            If Args <> vbNullString Then
                Select Case Args
                    Case "1" 'IPv4
                    Case "2" 'IPv6
                    Case Else: Send "522 Network protocol not supported, use (1,2)": Exit Sub
                End Select
            End If
            
            Set obj = New cWsCnn2
            If obj.CreateWS(Args <> "1") Then
                obj.CONTROL_SOCKET = SOCKET_ID
                mcSockets.Add obj, "#" & CStr(obj.LISTENER_SOCKET)
                Send "229 Entering Extended Passive Mode (|||" & mWinSock.SocketLocalPort(obj.LISTENER_SOCKET) & "|)"
            Else
                Send "425 Extended EPSV Failure, Could not create socket"
                Set obj = Nothing
            End If
                
        Case "PORT" '/* ACTIVE MODE */
            
            If OnlyEPSV Then Send "500 Illegal PORT command, EPSV ALL in effect.": Exit Sub
            If SOCKET_FAMILY = AF_INET6 Then Send "500 Illegal PORT command.": Exit Sub
            
            Dim sAddr() As String
            Dim lPort As Long

            sAddr = Split(Args, ",")
            If CheckBound(sAddr, 5) Then Send "501 Syntax error": Exit Sub
            lPort = (sAddr(4) * 256) + sAddr(5)
            tmp = sAddr(0) & "." & sAddr(1) & "." & sAddr(2) & "." & sAddr(3)

            If SOCKET_DATA <> 0 Then RemoveSH SOCKET_DATA
            SOCKET_DATA = mWinSock.WsConnect(tmp, lPort)
            
            If SOCKET_DATA Then
                '/* Socket Helper */
                Call CreateDataSH(SOCKET_DATA, WS, SOCKET_IP)
                Send "200 Port command successful"
            Else
                Send "421 Could not create socket"
            End If
            
        Case "EPRT" '/* EXTENDED ACTIVE MODE */
            
            If OnlyEPSV Then Send "500 Illegal EPRT command, EPSV ALL in effect.": Exit Sub
            If SOCKET_FAMILY = AF_INET Then Send "500 Illegal PORT command.": Exit Sub
            
            '|2|::1|54635|
            sAddr = Split(Args, "|")
            If Not CheckBound(sAddr, 4) Then Send "500 Illegal EPRT command.": Exit Sub
            
            If SOCKET_DATA <> 0 Then RemoveSH SOCKET_DATA
            Select Case sAddr(1)
                Case "1": SOCKET_DATA = mWinSock.WsConnect(sAddr(2), sAddr(3))
                Case "2": SOCKET_DATA = mWinSock.WsConnect(sAddr(2), sAddr(3), True)
                Case Else: Send "522 Network protocol not supported, use (1,2)": Exit Sub
            End Select
            
            If SOCKET_DATA Then
                '/* Socket Helper */
                Call CreateDataSH(SOCKET_DATA, WS, SOCKET_IP)
                Send "200 EPRT command successful"
            Else
                Send "522 Extended Port Failure, Could not create socket"
            End If
            
        Case "LIST", "MLSD", "NLST"
            
            If LenB(Args) Then
                Select Case LCase(Args)
                    Case "-a .", "-a", "-l": Args = vbNullString
                    Case Else
                    
                        If Left$(Args, 3) = "-a " Then Args = Right$(Args, Len(Args) - 3)
                        If Left$(Args, 2) = "-l" Then Args = Right$(Args, Len(Args) - 2)
                        
                        Args = Replace$(Args, "\", "")
                        Args = mvParseLocalPath(Args)
                        If Not PathExist(Args) Then
                            Send "450 Invalid file name or path"
                            If SOCKET_DATA Then RemoveSH SOCKET_DATA
                            Exit Sub
                        End If
                End Select
            End If
            
            Send "150 Opening data channel for directory listing"
            If Not mvWaitDataSocket Then CloseSocket: Exit Sub
            
            Select Case mcSockets(CStr(SOCKET_DATA)).SendList(Args, m_Path, m_Mounts, Cmd)
                Case 550: Send "550 File not found"
                Case 226: Send "226 Successfully transferred list"
                Case Else: Send "450 Couldn't open the file or directory"
            End Select
            RemoveSH SOCKET_DATA
            
        Case "MLST"
            
            If Len(Args) Then tmp = mvParseLocalPath(Args) Else tmp = m_Path: Args = m_url
            If Not PathExist(tmp) Then Send "550 Couldn't open the file or directory": Exit Sub
            
            If PathDirectory(tmp) Then
                tmp = "250-Listing " & Args & ENDL & _
                      " type=dir;modify=" & Format(FileDateTime2(tmp), "YYYYMMDDhhmmss") & ";perms=awrdf; " & ToUTF8(Args) & ENDL & _
                      " 250 End"
            Else
                tmp = "250-Listing " & Args & ENDL & _
                      " type=file;size=" & FileLen2(tmp) & "modify=" & Format(FileDateTime2(tmp), "YYYYMMDDhhmmss") & ";perms=awrdf; " & ToUTF8(Args) & ENDL & _
                      " 250 End"
            End If
            Send tmp
            
            '250-Listing
            ' type=dir;modify=20220108134311.628;perms=cpledf; /
            '250 End
            
            'MLST /programs
            '250-Listing /programs
            ' type=dir;modify=20220103023810.218;perms=cpledf; /programs
            '250 End
            
            '250-Listing coffe.png
            ' type=file;size=6392;modify=20211004150352.610;perms=awrdf; /coffe.png
            '250 End
            
            '250-Listing /coffe.png
            ' type=file;size=6392;modify=20211004150352.610;perms=awrdf; /coffe.png
            '250 End

        
        Case "SIZE"
            tmp = mvParseLocalPath(Args)
            If PathExist(tmp) Then Send "213 " & FileLen2(tmp) Else Send "550 File not found"
            
        Case "MDTM"
            tmp = mvParseLocalPath(Args)
            If PathExist(tmp) Then Send "213 " & CustomFileDateTime(tmp) Else Send "550 File not found"
            
        Case "RNFR"

            m_RnmFrom = mvParseLocalPath(Args)
            If Not Permission(READ_WRITE) Then Send "550 Permission denied": m_RnmFrom = vbNullString: Exit Sub
            If Not PathExist(m_RnmFrom) Then Send "550 file/directory not found": Exit Sub
            Send "350 Ready for destination name."
            
        Case "RNTO"
             If Not m_RnmFrom <> vbNullString Then Send "503 Bad sequence of commands!": Exit Sub
             tmp = mvParseLocalPath(Args)
             If Not Permission(READ_WRITE) Then Send "550 Permission denied": Exit Sub
             If PathExist(tmp) Then Send "553 file exists": Exit Sub
             If Not RenamePath(m_RnmFrom, tmp) Then Send "450 Internal error renaming the file": Exit Sub
             Send "250 file renamed successfully"
             m_RnmFrom = vbNullString
             
        Case "MKD"
            tmp = mvParseLocalPath(Args)
            If Not Permission(READ_WRITE) Then Send "550 Permission denied": Exit Sub
            
            If PathExist(tmp) Then Send "550 Directory already exists.": Exit Sub
            If CreateDirectory(tmp) Then Send "257 created successfully ": Exit Sub
            Send "450 Internal error creating the directory."
        
        Case "DELE", "RMD" '/* DeleteFile || DeleteDIR */
        
            tmp = mvParseLocalPath(Args)
            If Not (Permission(READ_WRITE)) Then Send "550 Permission denied": Exit Sub
            If Not PathExist(tmp) Then Send "550 " & IIf(Args = "DELE", "File", "Directory") & " not found": Exit Sub
            If Not ToRecycleBin(tmp, True, False) Then Send IIf(Args = "DELE", "500 Failed to delete the file", "450 Internal error deleting the directory"): Exit Sub
            Send "250 " & IIf(Args = "DELE", "File", "Directory") & " deleted successfully"
            
        Case "MODE"
        
            'If UCase$(Left(Args, 1)) = "S" Then m_Mode = emDownload: Send "200 MODE set to S.": Exit Sub
            If UCase$(Left(Args, 1)) = "S" Then Send "200 MODE set to S.": Exit Sub
            Send "502 Unimplemented MODE type"

        Case "RETR" '/* DOWNLOAD */
        
            tmp = mvParseLocalPath(Args)
            If Not Permission(READ_ONLY) Then Send "550 Permission denied": Exit Sub
            If Not PathExist(tmp) Then Send "550 File not found": Exit Sub

            Set obj = New cWsCnn2
            If obj.PrepareFile(tmp, emDownload, mtReset, mtPos) Then
                obj.CONTROL_SOCKET = SOCKET_ID
                Send "150 Opening data channel for file download from server"
                If Not mvWaitDataSocket Then CloseSocket True: Exit Sub
                Call DelegateSH(SOCKET_DATA, obj)
                FrmMain.FTPtransferStart SOCKET_DATA, tmp, emDownload, obj.TotalBytes, obj.TimeStamp, SOCKET_ID, mvUser.m_Name '[MAIN EVENT CALLBACK]
                mcSockets(CStr(SOCKET_DATA)).SendParts
            Else
                Set obj = Nothing
                Send "550 Could not open file for reading."
                If SOCKET_DATA <> 0 Then RemoveSH SOCKET_DATA
            End If
            
            
        Case "STOR", "APPE" '[UPLOAD]
        
            tmp = mvParseLocalPath(Args)
            If Not Permission(READ_WRITE) Then Send "550 Permission denied": Exit Sub
            'If args = "STOR" Then ClearFile tmp
            
            Set obj = New cWsCnn2
            If obj.PrepareFile(tmp, emUpload) Then
                obj.CONTROL_SOCKET = SOCKET_ID
                Send "150 Opening data channel for file upload to server"
                If Not mvWaitDataSocket Then CloseSocket True: Exit Sub
                Call DelegateSH(SOCKET_DATA, obj)
                FrmMain.FTPtransferStart SOCKET_DATA, tmp, emUpload, 0, obj.TimeStamp, SOCKET_ID, mvUser.m_Name
            Else
                Set obj = Nothing
                Send "550 Filename invalid."
                If SOCKET_DATA <> 0 Then RemoveSH SOCKET_DATA
            End If
            
        Case "REST" '/* RESTART DOWNLOAD */
        
            If Not IsNumeric(Args) Then Send "501 Bad parameter. Numeric value required": Exit Sub
            mtPos = Args: mtReset = True
            Send "350 Rest supported. Restarting"
            
        Case "MFMT"
            
            i = InStr(Args, " ")
            If i = 0 Then Send "501 Syntax error": Exit Sub
            tmp = Left(Args, i - 1)
            If Not IsNumeric(tmp) Or Len(tmp) < 14 Then Send "501 Not a valid date": Exit Sub
            
            Args = Right$(Args, Len(Args) - i)
            tmp2 = mvParseLocalPath(Args)
            If Not Permission(READ_WRITE) Then Send "550 Permission denied": Exit Sub
            tmp2 = GetSafeFileName(tmp2)
            
            If Not PathExist(tmp2) Then Send "550 File not found": Exit Sub
            If SetLastWriteFileDateTime(tmp2, tmp) Then
                Send "213 modify=" & tmp & Args
            Else
                Send "550 Failed to set file modification time"
            End If
        
        Case "FEAT"
        
            '211-Features:
            ' MDTM
            ' REST STREAM
            ' SIZE
            ' MODE Z
            ' MLST type*;size*;modify*;
            ' MLSD
            ' UTF8
            ' MFMT
            ' EPSV
            ' EPRT
            '211 End
            
            tmp = "211-Features:" & ENDL & _
                  " MDTM" & ENDL & _
                  " REST STREAM" & ENDL & _
                  " SIZE" & ENDL & _
                  " MLST type*;size*;modify*;" & ENDL & _
                  " MLSD" & ENDL & _
                  " UTF8" & ENDL & _
                  " MFMT" & ENDL & _
                  " EPSV" & ENDL & _
                  " EPRT" & ENDL & _
                  "211 End" & ENDL
                  
            Send tmp
            
        Case "SITE"
    
'            If Args = "help" Then
'                Send "200-CWD,NOOP,PWD,SYST,TYPE,PASV,MDTM,RNFR,RNTO"
'                Send "200 SIZE,RETR,STOR,MKD,DELE,PORT,LIST,QUIT,PASS,USER,SITE,OPTS"
'            Else
'                Send "504 Command not implemented for that parameter"
'            End If

             Send "500 Wrong command."
            
        Case "QUIT"
            Send "221 Goodbye"
            'Send "421 Server is going offline"
            'Send "421 Connection timed out."
            'Send "421 No-transfer-time exceeded. Closing control connection."
            'Send "421 Login time exceeded. Closing control connection."
            'Send "421 Kicked by Administrator"
            CloseSocket
        Case Else
            If Cmd <> "AUTH" Then Debug.Print Cmd & "->" & Args
            Send "502 Command not implemented."
            'CloseSocket
    End Select
    
End Sub

Private Property Get Permission(FLAG As FTPAccess) As Boolean
On Error GoTo e

    If Not mvUser.m_logged Then Exit Property
    '/* {name, path, access} */
    Select Case FLAG
        Case DISABLED_:
        Case READ_ONLY: If (m_Mounts(m_Index)(2) = READ_ONLY) Or (m_Mounts(m_Index)(2) = READ_WRITE) Then Permission = True
        Case READ_WRITE: If (m_Mounts(m_Index)(2) = READ_WRITE) Then Permission = True
    End Select
e:
End Property

Private Property Get argc(argv() As String) As Long
On Error GoTo err_
    argc = UBound(argv)
    Exit Property
err_:
    argc = -1
End Property
Private Property Get argbb(argv() As String) As Boolean
On Local Error GoTo err_
    argbb = UBound(argv) >= 0
err_:
End Property
