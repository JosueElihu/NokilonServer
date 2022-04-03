Attribute VB_Name = "mWinSock"
'--------------------------------------------------------------------------------
'    Component  : mWinSock 1.4.6
'    Autor      : J. Elihu
'    Original   : Leandro Ascierto
'    Modified   : 18/01/2022
'--------------------------------------------------------------------------------

Option Explicit

'/ ERROR CODES
Private Const WSABASEERR         As Long = 10000
Public Enum WS_ERROR
  WSAEINTR = (WSABASEERR + 4)
  WSAEACCES = (WSABASEERR + 13)
  WSAEFAULT = (WSABASEERR + 14)
  WSAEINVAL = (WSABASEERR + 22)
  WSAEMFILE = (WSABASEERR + 24)
  WSAEWOULDBLOCK = (WSABASEERR + 35)
  WSAEINPROGRESS = (WSABASEERR + 36)
  WSAEALREADY = (WSABASEERR + 37)
  WSAENOTSOCK = (WSABASEERR + 38)
  WSAEDESTADDRREQ = (WSABASEERR + 39)
  WSAEMSGSIZE = (WSABASEERR + 40)
  WSAEPROTOTYPE = (WSABASEERR + 41)
  WSAENOPROTOOPT = (WSABASEERR + 42)
  WSAEPROTONOSUPPORT = (WSABASEERR + 43)
  WSAESOCKTNOSUPPORT = (WSABASEERR + 44)
  WSAEOPNOTSUPP = (WSABASEERR + 45)
  WSAEPFNOSUPPORT = (WSABASEERR + 46)
  WSAEAFNOSUPPORT = (WSABASEERR + 47)
  WSAEADDRINUSE = (WSABASEERR + 48)
  WSAEADDRNOTAVAIL = (WSABASEERR + 49)
  WSAENETDOWN = (WSABASEERR + 50)
  WSAENETUNREACH = (WSABASEERR + 51)
  WSAENETRESET = (WSABASEERR + 52)
  WSAECONNABORTED = (WSABASEERR + 53)
  WSAECONNRESET = (WSABASEERR + 54)
  WSAENOBUFS = (WSABASEERR + 55)
  WSAEISCONN = (WSABASEERR + 56)
  WSAENOTCONN = (WSABASEERR + 57)
  WSAESHUTDOWN = (WSABASEERR + 58)
  WSAETIMEDOUT = (WSABASEERR + 60)
  WSAEHOSTUNREACH = (WSABASEERR + 65)
  WSAECONNREFUSED = (WSABASEERR + 61)
  WSAEPROCLIM = (WSABASEERR + 67)
  WSASYSNOTREADY = (WSABASEERR + 91)
  WSAVERNOTSUPPORTED = (WSABASEERR + 92)
  WSANOTINITIALISED = (WSABASEERR + 93)
  WSAHOST_NOT_FOUND = (WSABASEERR + 1001)
  WSATRY_AGAIN = (WSABASEERR + 1002)
  WSANO_RECOVERY = (WSABASEERR + 1003)
  WSANO_DATA = (WSABASEERR + 1004)
End Enum

'- WinSock
Private Declare Function WSAStartup Lib "ws2_32" (ByVal WsVersionRequired As Long, lpWSAData As Any) As Long
Private Declare Function WSACleanup Lib "ws2_32" () As Long
Private Declare Function WSAIsBlocking Lib "ws2_32" () As Long
Private Declare Function WSACancelBlockingCall Lib "ws2_32" () As Long
Private Declare Function WSAAsyncSelect Lib "ws2_32" (ByVal s As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Declare Function WSAGetLastError Lib "ws2_32" () As Long

Private Declare Function Ws_htonl Lib "ws2_32" Alias "htonl" (ByVal hostlong As Long) As Long
Private Declare Function Ws_htons Lib "ws2_32" Alias "htons" (ByVal hostshort As Long) As Integer
Private Declare Function Ws_ntohs Lib "ws2_32" Alias "ntohs" (ByVal netshort As Long) As Integer
Private Declare Function WS_Socket Lib "ws2_32" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Private Declare Function Ws_connect Lib "ws2_32" Alias "connect" (ByVal s As Long, ByRef name As Any, ByVal namelen As Long) As Long
Private Declare Function Ws_inet_ntoa Lib "ws2_32" Alias "inet_ntoa" (ByVal inn As Long) As Long
Private Declare Function Ws_inet_ntop Lib "ws2_32" Alias "inet_ntop" (ByVal af As Long, ByRef ppAddr As Any, ByRef pStringBuf As Any, ByVal StringBufSize As Long) As Long
Private Declare Function Ws_bind Lib "ws2_32" Alias "bind" (ByVal s As Long, addr As Any, ByVal namelen As Long) As Long
Private Declare Function Ws_listen Lib "ws2_32" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function Ws_accept Lib "ws2_32" Alias "accept" (ByVal s As Long, addr As sockaddr, addrlen As Long) As Long
Private Declare Function Ws_closesocket Lib "ws2_32" Alias "closesocket" (ByVal s As Long) As Long
Private Declare Function Ws_shutdown Lib "ws2_32" Alias "shutdown" (ByVal s As Long, ByVal How As Long) As Long
Private Declare Function Ws_send Lib "ws2_32" Alias "send" (ByVal s As Long, buf As Any, ByVal BufLen As Long, ByVal Flags As Long) As Long
Private Declare Function Ws_sendto Lib "ws2_32" Alias "sendto" (ByVal s As Long, buf As Any, ByVal BufLen As Long, ByVal Flags As Long, addr As sockaddr, ByVal tolen As Long) As Long
Private Declare Function Ws_ioctlsocket Lib "ws2_32" Alias "ioctlsocket" (ByVal s As Long, ByVal Cmd As Long, argp As Long) As Long
Private Declare Function Ws_recv Lib "ws2_32" Alias "recv" (ByVal s As Long, buf As Any, ByVal BufLen As Long, ByVal Flags As Long) As Long
Private Declare Function Ws_recvfrom Lib "ws2_32" Alias "recvfrom" (ByVal s As Long, buf As Any, ByVal BufLen As Long, ByVal Flags As Long, addr As sockaddr, ByRef fromlen As Long) As Long
Private Declare Function Ws_inet_addr Lib "ws2_32" Alias "inet_addr" (ByVal cp As String) As Long
Private Declare Function Ws_getsockopt Lib "ws2_32" Alias "getsockopt" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Private Declare Function Ws_setsockopt Lib "ws2_32" Alias "setsockopt" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function Ws_getsockname Lib "ws2_32" Alias "getsockname" (ByVal s As Long, name As sockaddr, ByRef namelen As Long) As Long
Private Declare Function Ws_getpeername Lib "ws2_32" Alias "getpeername" (ByVal s As Long, name As sockaddr, namelen As Long) As Long
Private Declare Function Ws_gethostbyname Lib "ws2_32" Alias "gethostbyname" (ByVal host_name As String) As Long
Private Declare Function Ws_gethostname Lib "ws2_32" Alias "gethostname" (ByVal host_name As String, ByVal namelen As Long) As Long
Private Declare Function Ws_gethostbyaddr Lib "ws2_32" Alias "gethostbyaddr" (haddr As Any, ByVal hnlen As Long, ByVal addrtype As Long) As Long
Private Declare Function Ws_inet_pton Lib "ws2_32" Alias "inet_pton" (ByVal Family As Long, ByVal pszAddrString As String, pAddrBuf As Any) As Long
Private Declare Function Ws_gethostbyname2 Lib "ws2_32" Alias "gethostbyname2" (ByVal host_name As String, ByVal af As Long) As Long
Private Declare Function Ws_freeaddrinfo Lib "ws2_32" Alias "freeaddrinfo" (ByVal lRes As Long) As Long
Private Declare Function Ws_getaddrinfo Lib "ws2_32" Alias "getaddrinfo" (ByVal NodeName As String, ByVal ServiceName As String, ByRef lpHints As addrinfo, lpResult As Long) As Long

'- User32
Private Declare Function CreateWindowExA Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'- Kernel32
Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal cb As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Any) As Long
Private Declare Function lstrcpyA Lib "kernel32" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Private Const WINSOCK_MESSAGE   As Long = 1025
Private Const INADDR_NONE       As Long = &HFFFF
Private Const INADDR_ANY        As Long = &H0
Private Const INVALID_SOCKET    As Long = -1
Private Const SOCKET_ERROR      As Long = -1
Private Const SOCK_STREAM       As Long = 1
Private Const AF_UNSPEC         As Long = 0
'Private Const AF_INET           As Long = 2
'Private Const AF_INET6          As Long = 23

Private Const FD_READ           As Long = &H1&
Private Const FD_WRITE          As Long = &H2&
Private Const FD_OOB            As Long = &H4&
Private Const FD_ACCEPT         As Long = &H8&
Private Const FD_CONNECT        As Long = &H10&
Private Const FD_CLOSE          As Long = &H20&
Private Const FD_ALL            As Long = &H3F&

Private Const GWL_WNDPROC       As Long = (-4)

Private Const SOL_SOCKET        As Long = 65535
Private Const SO_SNDBUF         As Long = &H1001&
Private Const SO_RCVBUF         As Long = &H1002&
Private Const FIONREAD          As Long = &H4004667F
Private Const FIONBIO           As Long = &H8004667E
Private Const IPPROTO_TCP       As Long = 6

' -=  IPv4 and IPv6  =-
Private Const INET_ADDRSTRLEN   As Long = 16
Private Const INET6_ADDRSTRLEN  As Long = 46
Private Const AI_PASSIVE        As Long = 1

Public Enum AddressFamilies
    AF_INET = 2         'IPv4
    AF_INET6 = 23       'IPv6
End Enum

Private Type sockaddr
    sa_family           As Integer  '  2 bytes
    sa_data(25)         As Byte     ' 26 bytes
End Type                            ' ========
                                    ' 28 bytes
Private Type sockaddr_in
    sin_family      As Integer      '  2 bytes
    sin_port        As Integer      '  2 bytes
    sin_addr        As Long         '  4 bytes
    sin_zero        As String * 8   ' 16 bytes
End Type                            ' ========
                                    ' 24 bytes
Private Type sockaddr_in6
    sin6_family         As Integer  '  2 bytes
    sin6_port           As Integer  '  2 bytes
    sin6_flowinfo       As Long     '  4 bytes
    sin6_addr(15)       As Byte     ' 16 bytes
    sin6_scope_id       As Long     '  4 bytes
End Type                            ' ========
                                    ' 28 bytes
Private Type addrinfo
    Flags     As Long
    Family    As Long
    socktype  As Long
    protocol  As Long
    addrlen   As Long
    canonname As Long     ' String     Ptr
    addr      As Long     ' Sockaddr   Ptr
    next      As Long     ' Addrinfo   Ptr
End Type

'-=  Adtional  =-
Public Type mtWSPortItem
    Num      As Long
    Handle   As Long
    Family   As AddressFamilies
End Type

Private m_CallBack          As Object
Private m_hwnd              As Long
Private m_bStartup          As Boolean
Private m_PrevProc          As Long
Private m_lWSErr            As Long

Private mmSockets           As Collection   '/* mmSockets: {SOCKET, SERVER_SOCKET, SEND_BUF}    */
Private mmPorts             As Collection   '/* mmPorts  : {SERVER_SOCKET, PORT_NUM, FAMILY}    */


Public Function WsStart(Callback As Object) As Boolean
Dim baWSAData() As Byte
    
    If m_bStartup Then WsStart = True: Exit Function
    
    ReDim baWSAData(0 To 1000) As Byte
    If WSAStartup(&H101, baWSAData(0)) <> 0 Then WsError "Start", Err.LastDllError: Exit Function
    
    m_bStartup = True
    m_hwnd = CreateWindowExA(0&, "STATIC", "SOCKET_WINDOW", 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)
    m_PrevProc = SetWindowLongA(m_hwnd, GWL_WNDPROC, AddressOf WindowProc)
    
    Set m_CallBack = Callback
    Set mmSockets = New Collection
    Set mmPorts = New Collection
    WsStart = m_bStartup

End Function

Public Sub WsTerminate()
On Error Resume Next
Dim i As Long

    For i = 1 To mmSockets.Count
        WsDisconnect mmSockets.Item(1)(0)
    Next
    
    For i = 1 To mmPorts.Count
        Ws_closesocket mmPorts.Item(1)(0)
        mmPorts.Remove (1)
    Next
    
    If WSAIsBlocking Then WSACancelBlockingCall
    Call WSACleanup
    
    m_bStartup = False
    SetWindowLongA m_hwnd, GWL_WNDPROC, m_PrevProc
    DestroyWindow m_hwnd
    
    Set mmSockets = Nothing
    Set mmPorts = Nothing
    Set m_CallBack = Nothing
    
End Sub

Public Function WsConnect(ByVal Address As String, ByVal RemotePort As Long, Optional ByVal IPv6 As Boolean) As Long
Dim hSocket As Long

    hSocket = WS_Socket(IIf(IPv6, AF_INET6, AF_INET), SOCK_STREAM, IPPROTO_TCP)
    If hSocket = INVALID_SOCKET Then hSocket = 0: GoTo e_
    
    ' • TODO:  IPv6
    '----------------------------------------------------------------------
    '   Dim Sin6 As sockaddr_in6
    '   Sin6.sin6_family = AF_INET6
    '   Sin6.sin6_port = Ws_htons(RemotePort)
    '   If Ws_inet_pton(AF_INET6, Address, Sin6.sin6_addr(0)) = 0 Then GoTo e_
    '   If Sin6.sin6_port = INVALID_SOCKET Then GoTo e_
    '----------------------------------------------------------------------
    
    ' • TODO:  IPv4
    '----------------------------------------------------------------------
    '   Dim Sin4 As sockaddr_in
    '   Sin4.sin_family = AF_INET
    '   Sin4.sin_addr = mv_inet_addr(Address)
    '   Sin4.sin_port = Ws_htons(RemotePort)
    '   If Sin4.sin_port = INVALID_SOCKET Then GoTo e_
    '   If Sin4.sin_addr = INADDR_NONE Then GoTo e_
    '-------------------------------------------------------------------------
    
    Dim uAddr  As sockaddr
    
    '/* IPv4 - IPv6 */
    If mv_addr_info(Address, RemotePort, IIf(IPv6, AF_INET6, AF_INET), uAddr) = False Then GoTo e_
    
    '/* Connect to server */
    If Ws_connect(hSocket, uAddr, LenB(uAddr)) = SOCKET_ERROR Then GoTo e_
    If Not mvAsyncSelect(hSocket) Then GoTo e_
    
    '/* mmSockets: {SOCKET, SERVER_SOCKET, SEND_BUF}    */
    mvAddCollection mmSockets, hSocket, hSocket, 0, mvSendBuf(hSocket)
    WsConnect = hSocket
    Exit Function
e_:
    WsError "Connect", WSAGetLastError
    If hSocket Then Ws_closesocket hSocket
End Function


Public Function WsDisconnect(ByVal WS As Long) As Boolean
On Error Resume Next
    WsDisconnect = Ws_closesocket(WS) = 0
    If WsDisconnect Then mmSockets.Remove CStr(WS) Else WsError "WsDiconnect", WSAGetLastError
End Function

'TODO: Send && Wait complete
'----------------------------------------------------------------------------------------------------------------------
Public Function WsSend(WS As Long, data As Variant) As Boolean
Dim Out()   As Byte
Dim tmp     As String
Dim lBytes  As Long
Dim lResult As Long
Dim lErr    As Long
Dim ln      As Long
Dim lBuff   As Long

    Select Case VarType(data)
        Case 8209:  tmp = data                           '- byte array
        Case 8:     tmp = StrConv(data, vbFromUnicode)   '- String
        Case Else:  tmp = StrConv(CStr(data), vbFromUnicode)
    End Select
    If Len(tmp) = 0 Then Exit Function Else Out = tmp
    

'    Do Until lBytes > UBound(out)
'        lResult = ws_send(WS, out(lBytes), (UBound(out) + 1) - lBytes, 0&)
'        If lResult = SOCKET_ERROR Then
'            lErr = Err.LastDllError
'            If lErr = WSAEWOULDBLOCK Then mvWait (1) Else WsError "WsSend[" & WS & "]", lErr: GoTo e
'        Else
'            lBytes = lBytes + lResult
'        End If
'    Loop

    lBuff = mmSockets(CStr(WS))(2)  '/* SO_SNDBUF */
    Do Until lBytes > UBound(Out)

        ln = (UBound(Out) + 1) - lBytes
        If ln > lBuff Then ln = lBuff

        lResult = Ws_send(WS, Out(lBytes), ln, 0&)
        If lResult = SOCKET_ERROR Then
            lErr = Err.LastDllError
            If lErr = WSAEWOULDBLOCK Then mvWait (1) Else WsError "WsSend[" & WS & "]", lErr: GoTo e
        Else
            lBytes = lBytes + lResult
        End If
    Loop
    
e:
    WsSend = (lResult <> SOCKET_ERROR)
    If Not WsSend Then WsError "WsSend", WSAGetLastError
    Call m_CallBack.WSOnSend(WS, WsSend, lBytes, lErr)
End Function


'TODO: Send && Wait FD_WRITE MSG
'----------------------------------------------------------------------------------------------------------------------
Public Function WsSend2(WS As Long, data As Variant, Optional ByRef lSendBytes As Long, Optional ByRef lReturnErr As Long) As Boolean
Dim Out()   As Byte
Dim tmp     As String
Dim lBytes  As Long
Dim lResult As Long
Dim lErr    As Long
Dim ln      As Long
Dim lBuff   As Long

    Select Case VarType(data)
        Case 8209:  tmp = data                           '- byte array
        Case 8:     tmp = StrConv(data, vbFromUnicode)   '- String
        Case Else:  tmp = StrConv(CStr(data), vbFromUnicode)
    End Select
    If Len(tmp) = 0 Then Exit Function Else Out = tmp

'    Do Until lResult = SOCKET_ERROR Or lBytes > UBound(out)
'        lResult = ws_send(WS, out(lBytes), (UBound(out) + 1) - lBytes, 0&)
'        If lResult = SOCKET_ERROR Then
'            lErr = Err.LastDllError
'            If lErr <> WSAEWOULDBLOCK Then WsError "WsSend[" & WS & "]", lErr
'        Else
'            lBytes = lBytes + lResult
'        End If
'    Loop
    
    lBuff = mmSockets(CStr(WS))(2)  '/* SO_SNDBUF */
    Do Until lResult = SOCKET_ERROR Or lBytes > UBound(Out)

        ln = (UBound(Out) + 1) - lBytes
        If ln > lBuff Then ln = lBuff

        lResult = Ws_send(WS, Out(lBytes), ln, 0&)
        If lResult = SOCKET_ERROR Then
            lErr = Err.LastDllError
            If lErr <> WSAEWOULDBLOCK Then WsError "WsSend[" & WS & "]", lErr
        Else
            lBytes = lBytes + lResult
        End If
    Loop
    
    lSendBytes = lBytes
    lReturnErr = lErr
    
    WsSend2 = (lResult <> SOCKET_ERROR)
    If Not WsSend2 Then WsError "WsSend2", WSAGetLastError
    Call m_CallBack.WSOnSend(WS, WsSend2, lBytes, lErr)
End Function

Public Function WsListen(ByVal LocalPort As Long, Optional ByVal IPv6 As Boolean) As Long
Dim hSocket  As Long

    hSocket = WS_Socket(IIf(IPv6, AF_INET6, AF_INET), SOCK_STREAM, IPPROTO_TCP)
    If hSocket = SOCKET_ERROR Then hSocket = 0: GoTo e_
    
    
    ' IPv4
    '----------------------------------------------------
    '   Dim Sin4  As sockaddr_in
    '   Sin4.sin_family = AF_INET
    '   Sin4.sin_port = Ws_htons(LocalPort)
    '   Sin4.sin_addr = Ws_htonl(INADDR_ANY)
    '   If Sin4.sin_port = INVALID_SOCKET Then GoTo e_
    '   If Sin4.sin_addr = INADDR_NONE Then GoTo e_
    '----------------------------------------------------
    
    Dim uAddr   As sockaddr
    uAddr.sa_family = IIf(IPv6, AF_INET6, AF_INET)
    uAddr.sa_data(0) = mvPeekB(VarPtr(LocalPort) + 1)
    uAddr.sa_data(1) = mvPeekB(VarPtr(LocalPort))
    
    '/* Bind Socket */
    If Ws_bind(hSocket, uAddr, LenB(uAddr)) = SOCKET_ERROR Then GoTo e_
    If Not mvAsyncSelect(hSocket) Then GoTo e_
    If Ws_listen(hSocket, 1) = SOCKET_ERROR Then GoTo e_
    
    '/* mmPorts  : {SERVER_SOCKET, PORT_NUM, FAMILY}  */
    mvAddCollection mmPorts, hSocket, hSocket, LocalPort, IIf(IPv6, AF_INET6, AF_INET)
    WsListen = hSocket
    Exit Function
e_:
    WsError "WsListen", WSAGetLastError
    If hSocket Then Ws_closesocket hSocket
End Function


Public Function WsClose(ByVal NumOrWSP As Long, Optional CloseConnections As Boolean = True) As Boolean
On Error GoTo ee_
Dim i  As Long
Dim j  As Long
Dim nn As Long

    If NumOrWSP = 0 Then Exit Function
    
    '/* Check NumOrWPS is Socket Handle  */
    If mvWspIsHandle(NumOrWSP) Then GoTo SSHH_
    'If SocketLocalPort(NumOrWSP) = SOCKET_ERROR Then GoTo SSHH_
    
    '/* Close Port by Num */
    For i = mmPorts.Count To 1 Step -1
        If NumOrWSP = mmPorts(i)(1) Then
            nn = mmPorts(i)(0)
            WsClose = (Ws_closesocket(nn) = 0)
            If Not WsClose Then GoTo ee_
            If WsClose Then
                If CloseConnections Then
                    For j = mmSockets.Count To 1 Step -1
                        If mmSockets(j)(1) = nn Then WsDisconnect mmSockets(j)(0)
                    Next
                End If
                mmPorts.Remove i
            End If
        End If
    Next
    GoTo ee_
    
SSHH_:  '/* Close Port by Handle */
    WsClose = (Ws_closesocket(NumOrWSP) = 0)
    If Not WsClose Then GoTo ee_
    If CloseConnections Then
        For i = mmSockets.Count To 1 Step -1
            If mmSockets(i)(1) = NumOrWSP Then WsDisconnect mmSockets(i)(0)
        Next
    End If
    mmPorts.Remove CStr(NumOrWSP)
    Exit Function
ee_:
    WsError "WsClose", WSAGetLastError
End Function


' PROPERTIIES
'======================================================================================================================
Property Get LocalHostName() As String
    LocalHostName = String$(256, 0)
    If Ws_gethostname(LocalHostName, 256) <> INVALID_SOCKET Then LocalHostName = Left(LocalHostName, InStr(LocalHostName, Chr(0)) - 1) Else LocalHostName = vbNullString
End Property

Property Get LocalHostIP(Optional IPv6 As Boolean) As String
Dim uAddr As sockaddr
Dim Sin4  As sockaddr_in
Dim Sin6  As sockaddr_in6
Dim lPtr  As Long
Dim Out() As Byte
    
    If mv_addr_info("", 0, IIf(IPv6, AF_INET6, AF_INET), uAddr) Then
        If uAddr.sa_family = AF_INET6 Then
            ReDim Out(INET6_ADDRSTRLEN - 1)
            MemCopy Sin6, uAddr, LenB(Sin6)
            lPtr = Ws_inet_ntop(AF_INET6, Sin6.sin6_addr(0), 0&, INET6_ADDRSTRLEN)
        Else
            MemCopy Sin4, uAddr, LenB(Sin4)
            lPtr = Ws_inet_ntoa(Sin4.sin_addr)
        End If
    End If
    If lPtr Then LocalHostIP = StrFromPtr(lPtr) Else LocalHostIP = IIf(IPv6, "::", "0.0.0.0")
    
End Property

Property Get SocketName(ByVal WS As Long) As String
Dim uAddr   As sockaddr
Dim lPtr    As Long

    If Ws_getpeername(WS, uAddr, LenB(uAddr)) = 0 Then
        Dim Sin4 As sockaddr_in
        Dim Sin6 As sockaddr_in6
        
        If uAddr.sa_family = AF_INET6 Then
            MemCopy Sin6, uAddr, LenB(Sin6)
            lPtr = Ws_gethostbyaddr(Sin6.sin6_addr(0), 16, AF_INET6)
        Else
            MemCopy Sin4, uAddr, LenB(Sin4)
            lPtr = Ws_gethostbyaddr(Sin4.sin_addr, 4, AF_INET)
        End If
        
        If lPtr = 0 Then Exit Property
        MemCopy lPtr, ByVal lPtr, 4
        SocketName = StrFromPtr(lPtr)
    End If
End Property

Property Get SocketIP(ByVal WS As Long) As String
Dim uAddr   As sockaddr
    If Ws_getpeername(WS, uAddr, LenB(uAddr)) = 0 Then
        SocketIP = mv_addr_ip(uAddr)
    Else
        SocketIP = "0.0.0.0"
    End If
End Property
Property Get SocketFamily(ByVal WS As Long) As AddressFamilies
Dim uAddr As sockaddr
    If Ws_getsockname(WS, uAddr, LenB(uAddr)) <> SOCKET_ERROR Then
        SocketFamily = uAddr.sa_family
    Else
        SocketFamily = 0
    End If
End Property

Property Get SocketLocalPort(ByVal WS As Long) As Long
Dim uAddr   As sockaddr

    '- GET: PortNum From WS (Local)
    If Ws_getsockname(WS, uAddr, LenB(uAddr)) <> SOCKET_ERROR Then
        Dim Sin4 As sockaddr_in
        Dim Sin6 As sockaddr_in6
        If uAddr.sa_family = AF_INET6 Then
            MemCopy Sin6, uAddr, LenB(Sin6)
            SocketLocalPort = mvUnsigned(Ws_ntohs(Sin6.sin6_port))
        Else
            MemCopy Sin4, uAddr, LenB(Sin4)
            SocketLocalPort = mvUnsigned(Ws_ntohs(Sin4.sin_port))
        End If
    Else
        SocketLocalPort = SOCKET_ERROR
    End If

End Property
Property Get SocketRemotePort(ByVal WS As Long) As Long
Dim uAddr   As sockaddr
    
    '- GET: PortNum From WS (Remote)
    If Ws_getpeername(WS, uAddr, LenB(uAddr)) = 0 Then
        Dim Sin4 As sockaddr_in
        Dim Sin6 As sockaddr_in6
        If uAddr.sa_family = AF_INET6 Then
            MemCopy Sin6, uAddr, LenB(Sin6)
            SocketRemotePort = mvUnsigned(Ws_ntohs(Sin6.sin6_port))
        Else
            MemCopy Sin4, uAddr, LenB(Sin4)
            SocketRemotePort = mvUnsigned(Ws_ntohs(Sin4.sin_port))
        End If
    Else
        SocketRemotePort = SOCKET_ERROR
    End If
End Property


'- Additional
'=====================================================================================================================
Property Get OpenPorts() As Long: OpenPorts = mmPorts.Count: End Property
Property Get OpenPort(ByVal Index As Long) As mtWSPortItem
On Error GoTo e
    OpenPort.Num = mmPorts(Index)(1)
    OpenPort.Handle = mmPorts(Index)(0)
    OpenPort.Family = mmPorts(Index)(2)
e:
End Property
Property Get Sockets() As Collection
On Error GoTo e
Dim i As Long
    '/* Return Connection Sockets                       */
    '/* mmSockets: {SOCKET, SERVER_SOCKET, SEND_BUF}    */
    Set Sockets = New Collection
    For i = 1 To mmSockets.Count
        Sockets.Add mmSockets(i)(0)
    Next
e:
End Property
Property Get WSLastError() As WS_ERROR: WSLastError = m_lWSErr: End Property


' TODO:  PRIVATE FUNCTIONS
'=====================================================================================================================

Private Function mvAsyncSelect(hSocket As Long) As Boolean
    mvAsyncSelect = WSAAsyncSelect(hSocket, m_hwnd, ByVal WINSOCK_MESSAGE, ByVal FD_READ Or FD_WRITE Or FD_ACCEPT Or FD_CONNECT Or FD_CLOSE) <> SOCKET_ERROR
End Function
Private Function mvWspIsHandle(WSS As Long) As Boolean
On Error GoTo e
    mvWspIsHandle = mmPorts(CStr(WSS))(0) <> 0
e:
End Function
Private Function mv_inet_addr(Address As String) As Long
Dim lPtr As Long
    mv_inet_addr = Ws_inet_addr(Address)
    If mv_inet_addr = INADDR_NONE Then
        lPtr = Ws_gethostbyname(Address)
        If lPtr = 0 Then Exit Function
        Call MemCopy(lPtr, ByVal mvUnsignedAdd(lPtr, 12), 4)
        Call MemCopy(lPtr, ByVal lPtr, 4)
        Call MemCopy(mv_inet_addr, ByVal lPtr, 4)
    End If
End Function

Private Function mv_addr_info(Address As String, lPort As Long, Family As Long, uAddr As sockaddr) As Boolean
Dim addr_i  As addrinfo
Dim lPtr    As Long

    addr_i.Family = Family
    If LenB(Address) = 0 Then addr_i.Flags = AI_PASSIVE
    If Ws_getaddrinfo(Address, lPort, addr_i, lPtr) <> 0 Then GoTo e_
    addr_i.next = lPtr                                                  ' Point to first structure in linked list
    MemCopy addr_i, ByVal addr_i.next, LenB(addr_i)                     ' Copy next address info to Hints
    MemCopy uAddr, ByVal addr_i.addr, LenB(uAddr)                       ' Save sockaddr portion
    mv_addr_info = True
e_:
    If lPtr Then Ws_freeaddrinfo lPtr
End Function

Private Function mv_addr_ip(uAddr As sockaddr) As String
Dim Sin4 As sockaddr_in
Dim Sin6 As sockaddr_in6
Dim lPtr    As Long
Dim Out()  As Byte
    
    If uAddr.sa_family = AF_INET6 Then
        ReDim Out(INET6_ADDRSTRLEN - 1)
        MemCopy Sin6, uAddr, LenB(Sin6)
        lPtr = Ws_inet_ntop(AF_INET6, Sin6.sin6_addr(0), Out(0), INET6_ADDRSTRLEN)
    Else
        MemCopy Sin4, uAddr, LenB(Sin4)
        lPtr = Ws_inet_ntoa(Sin4.sin_addr)
    End If
    If lPtr Then mv_addr_ip = StrFromPtr(lPtr)
    
End Function
Private Function mvSendBuf(WS As Long) As Long
    If Ws_getsockopt(WS, SOL_SOCKET, SO_SNDBUF, mvSendBuf, LenB(mvSendBuf)) = SOCKET_ERROR Then mvSendBuf = 1024
End Function
Private Function StrFromPtr(ByVal lPtr As Long) As String
    StrFromPtr = String$(lstrlenA(ByVal lPtr), 0)
    lstrcpyA ByVal StrFromPtr, ByVal lPtr
End Function

Private Sub mvAddCollection(mColl As Collection, ByVal Key As String, ParamArray elements() As Variant)
On Error GoTo e

    '/* mmSockets: {SOCKET, SERVER_SOCKET, SEND_BUF}    */
    '/* mmPorts  : {SERVER_SOCKET, PORT_NUM, FAMILY}    */
    mColl.Add elements, Key
e:
End Sub
Private Function mvUnsigned(Value As Integer) As Long
    If Value < 0 Then mvUnsigned = Value + 65536 Else mvUnsigned = Value
End Function
Private Function mvUnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    mvUnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function
Private Function mvPeekB(ByVal lpdwData As Long) As Byte
    MemCopy mvPeekB, ByVal lpdwData, 1
End Function

Private Function mv_Remote_Info(hSocket As Long, ByRef lRemotePort As Long, ByRef sIP As String, ByRef sHost As String) As Boolean
Dim uAddr As sockaddr
Dim lPtr  As Long
Dim lPtr2 As Long

    If Ws_getpeername(hSocket, uAddr, LenB(uAddr)) = 0 Then
    
        Dim Sin4 As sockaddr_in
        Dim Sin6 As sockaddr_in6
        Dim Out() As Byte
        
        If uAddr.sa_family = AF_INET6 Then
            MemCopy Sin6, uAddr, LenB(Sin6)
            lRemotePort = mvUnsigned(Ws_ntohs(Sin6.sin6_port))
            ReDim Out(0 To INET6_ADDRSTRLEN - 1)
            
            lPtr = Ws_inet_ntop(AF_INET6, Sin6.sin6_addr(0), Out(0), INET6_ADDRSTRLEN)
            lPtr2 = Ws_gethostbyaddr(Sin6.sin6_addr(0), 16, AF_INET6)
            
        Else
            MemCopy Sin4, uAddr, LenB(Sin4)
            lRemotePort = mvUnsigned(Ws_ntohs(Sin4.sin_port))
            ReDim Out(0 To INET_ADDRSTRLEN - 1)
            
            lPtr = Ws_inet_ntoa(Sin4.sin_addr)
            lPtr2 = Ws_gethostbyaddr(Sin4.sin_addr, 4, AF_INET)
            
        End If
        
        If lPtr Then sIP = StrFromPtr(lPtr)
        
        If lPtr2 Then
            MemCopy lPtr2, ByVal lPtr2, 4
            sHost = StrFromPtr(lPtr2)
        Else
            sHost = vbNullString
        End If
        
        mv_Remote_Info = True
    Else
        lRemotePort = 0
        sIP = vbNullString
        sHost = vbNullString
    End If
End Function


Private Function WsError(Source As String, ByVal Num As Long)
Dim tmp As String
    Select Case Num
        Case WSAEACCES:         tmp = "Permission denied."
        Case WSAEADDRINUSE:     tmp = "Address already in use."
        Case WSAEADDRNOTAVAIL:  tmp = "Cannot assign requested address."
        Case WSAEAFNOSUPPORT:   tmp = "Address family not supported by protocol family."
        Case WSAEALREADY:       tmp = "Operation already in progress."
        Case WSAECONNABORTED:   tmp = "Software caused connection abort."
        Case WSAECONNREFUSED:   tmp = "Connection refused."
        Case WSAECONNRESET:     tmp = "Connection reset by peer."
        Case WSAEDESTADDRREQ:   tmp = "Destination address required."
        Case WSAEFAULT:         tmp = "Bad address."
        Case WSAEHOSTUNREACH:   tmp = "No route to host."
        Case WSAEINPROGRESS:    tmp = "Operation now in progress."
        Case WSAEINTR:          tmp = "Interrupted function call."
        Case WSAEINVAL:         tmp = "Invalid argument."
        Case WSAEISCONN:        tmp = "Socket is already connected."
        Case WSAEMFILE:         tmp = "Too many open files."
        Case WSAEMSGSIZE:       tmp = "Message too long."
        Case WSAENETDOWN:       tmp = "Network is down."
        Case WSAENETRESET:      tmp = "Network dropped connection on reset."
        Case WSAENETUNREACH:    tmp = "Network is unreachable."
        Case WSAENOBUFS:        tmp = "No buffer space available."
        Case WSAENOPROTOOPT:    tmp = "Bad protocol option."
        Case WSAENOTCONN:       tmp = "Socket is not connected."
        Case WSAENOTSOCK:       tmp = "Socket operation on nonsocket."
        Case WSAEOPNOTSUPP:     tmp = "Operation not supported."
        Case WSAEPFNOSUPPORT:   tmp = "Protocol family not supported."
        Case WSAEPROCLIM:       tmp = "Too many processes."
        Case WSAEPROTOTYPE:     tmp = "Protocol wrong type for socket."
        Case WSAESHUTDOWN:      tmp = "Cannot send after socket shutdown."
        Case WSAETIMEDOUT:      tmp = "Connection timed out."
        Case WSAEWOULDBLOCK:    tmp = "Resource temporarily unavailable."
        Case WSAHOST_NOT_FOUND: tmp = "Host not found."
        Case WSANOTINITIALISED: tmp = "Successful WSAStartup not yet performed."
        Case WSANO_DATA:        tmp = "Valid name, no data record of requested type."
        Case WSANO_RECOVERY:    tmp = "This is a nonrecoverable error."
        Case WSASYSNOTREADY:    tmp = "Network subsystem is unavailable."
        Case WSATRY_AGAIN:      tmp = "Nonauthoritative host not found."
        Case WSAEPROTONOSUPPORT: tmp = "Protocol not supported."
        Case WSAESOCKTNOSUPPORT: tmp = "Socket type not supported."
        Case WSAVERNOTSUPPORTED: tmp = "Winsock.dll version out of range."
        Case Else:               tmp = "Unknown error."
    End Select
    m_lWSErr = Num
    'Debug.Print "mWinSock::" & Source & ": " & tmp
    Call m_CallBack.WSOnError(Num, tmp, Source)
End Function
Private Function mvWait(Optional Seconds As Single = 1) As Single
    mvWait = Timer
    Do While Timer - mvWait < Seconds: DoEvents: Loop
End Function
Private Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next

    If uMsg = WINSOCK_MESSAGE Then
    
        Dim msIP    As String
        
        Select Case lParam
            Case FD_ACCEPT
            
                Dim hSocket As Long
                Dim uAddr   As sockaddr
                
                hSocket = Ws_accept(wParam, uAddr, LenB(uAddr))
                If hSocket <> SOCKET_ERROR Then
                
                    msIP = mv_addr_ip(uAddr)
                    '/* mmSockets: {SOCKET, SERVER_SOCKET, SEND_BUF}    */
                    mvAddCollection mmSockets, hSocket, hSocket, wParam, mvSendBuf(hSocket)
                    Call m_CallBack.WSOnConnect(hSocket, msIP, wParam, uAddr.sa_family = AF_INET6)
                    
                Else
                    WsError "Accept", Err.LastDllError
                End If
                
            Case FD_CONNECT:
                'Debug.Print "FD_CONNECT ", wParam, lParam
            Case FD_WRITE:
            
                '------------------------------------------------------------------------
                'Socket in a writable state, buffer for outgoing data of the transport
                'service is empty and ready to receive data to send through the network.
                '(WSAEWOULDBLOCK)
                '------------------------------------------------------------------------
                 Call m_CallBack.WSOnWritable(wParam)
                 
            Case FD_READ
        
                Dim Out()  As Byte
                Dim lBytes As Long
                Dim lPos   As Long
                Dim lSize  As Long
                
                If Ws_ioctlsocket(wParam, FIONREAD, lBytes) = SOCKET_ERROR Then lBytes = SOCKET_ERROR
                If lBytes <= 0 Then Exit Function
                
                ReDim Out(0 To lBytes - 1) As Byte
                Do
                    lPos = lPos + lSize
                    lSize = Ws_recv(wParam, Out(lPos), lBytes - lPos, 0&)
                    If lSize = SOCKET_ERROR Then
                       WsError "Read", Err.LastDllError: Exit Function
                    ElseIf lSize = 0 Then
                        If lPos = 0 Then
                            Out = vbNullString
                        Else
                            ReDim Preserve Out(0 To lPos - 1)
                        End If
                        Exit Do
                    End If
                Loop While lPos + lSize < lBytes
                Call m_CallBack.WSOnRead(wParam, StrConv(Out(), vbUnicode), lBytes)
               
            Case Else 'FD_CLOSE
                WsDisconnect wParam
                Call m_CallBack.WSOnDisconnect(wParam)
        End Select
        
    Else
        WindowProc = CallWindowProcA(m_PrevProc, hwnd, uMsg, wParam, lParam)
    End If
    If Err.Number Then Debug.Print "WinSok::WindowProc: " & Err.Description
    
End Function


'WINSOCK CALLBACKS: Include these events in yor callback object
'====================================================================================================================
'Public Sub WSOnConnect(WS As Long, IP As String, SERVER_SOCKET As Long, ByVal IPv6 As Boolean)
'''
'End Sub
'Public Sub WSOnWritable(WS As Long)
'''
'End Sub
'Public Sub WSOnRead(WS As Long, Data As String, ByVal lBytes As Long)
'''
'End Sub
'Public Sub WSOnSend(WS As Long, ByVal Result As Boolean, ByVal lBytes As Long, ByVal lErr As WS_ERROR)
'''
'End Sub
'Public Sub WSOnDisconnect(WS As Long)
'''
'End Sub
'Public Sub WSOnError(ByVal Num As Long, ByVal Description As String, Source As String)
'''
'End Sub
'====================================================================================================================
