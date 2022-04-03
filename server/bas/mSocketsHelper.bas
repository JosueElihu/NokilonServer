Attribute VB_Name = "mSocketsHelper"
Option Explicit


Private Declare Function GetTcpTable Lib "iphlpapi" (ByRef pTcpTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Private Declare Function GetExtendedTcpTable Lib "iphlpapi" (ByRef pTcpTable As Any, ByRef dwOutBufLen As Long, ByVal bSort As Boolean, ByVal ipVersion As Integer, ByVal tblClass As TCP_TABLE_CLASS, ByVal Reserved As Long) As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal cb As Long)

Private Declare Function Ws_ntohs Lib "ws2_32" Alias "ntohs" (ByVal netshort As Long) As Integer
Private Declare Function GetModuleFileNameExA Lib "psapi" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Private Enum TCP_TABLE_CLASS
  TCP_TABLE_BASIC_LISTENER
  TCP_TABLE_BASIC_CONNECTIONS
  TCP_TABLE_BASIC_ALL
  TCP_TABLE_OWNER_PID_LISTENER
  TCP_TABLE_OWNER_PID_CONNECTIONS
  TCP_TABLE_OWNER_PID_ALL
  TCP_TABLE_OWNER_MODULE_LISTENER
  TCP_TABLE_OWNER_MODULE_CONNECTIONS
  TCP_TABLE_OWNER_MODULE_ALL
End Enum

Private Type MIB_TCPROW
  dwState       As Long
  dwLocalAddr   As Long
  dwLocalPort   As String * 2
  dwRemoteAddr  As Long
  dwRemotePort  As String * 2
End Type
Private Type MIB_TCPROW_OWNER_PID
  dwState      As Long
  dwLocalAddr  As Long
  dwLocalPort  As String * 2
  dwRemoteAddr As Long
  dwRemotePort As String * 2
  dwOwningPid  As Long
End Type
Private Type MIB_TCP6ROW_OWNER_PID
  ucLocalAddr(15)   As Byte         ' 16 bytes
  dwLocalScopeId    As Long
  dwLocalPort       As Long
  ucRemoteAddr(15)  As Byte         ' 16 bytes
  dwRemoteScopeId   As Long
  dwRemotePort      As Long
  dwState           As Long
  dwOwningPid       As Long
End Type

Public Enum FtpSocketType
  FTP_CONTROL_SOCKET = 0
  FTP_DATA_SOCKET = 2
End Enum

Public mcSockets     As Collection       '/* Sockect Connections Helpers */

Public Function CreateSH(ID As Long, IP As String, WSS As Long) As Boolean

    '/* Initialize Socket Helpers */
    If mcSockets Is Nothing Then Set mcSockets = New Collection
    
    '/* Check connection is in Server Sockets */
    '/* If true, create FTP Control Object  */
    mcSockets.Add IIf(WSS = SERVER.SOCKET4 Or WSS = SERVER.SOCKET6, New cWsCnn, New cWsCnn2), CStr(ID)
    
    If mcSockets(CStr(ID)).SOCKET_TYPE = FTP_DATA_SOCKET Then
        If ExistSH("#" & WSS) Then
        
            '/* Change opened port handle to data connection */
            mcSockets("#" & WSS).Dump mcSockets(CStr(ID))
            mcSockets.Remove "#" & WSS
            
        Else
            '/* REMOVE IF THE MAIN THREAD DOES NOT EXIST */
            PushLog "WARNING: Main Socket does not exist: " & ID, enmPSNone, mclERRORS
            RemoveSH ID
            Exit Function
        End If
    End If
    CreateSH = True
    
End Function

Public Function CreateDataSH(WS As Long, WS2 As Long, IP As String) As Boolean
    mcSockets.Add New cWsCnn2, CStr(WS)
    mcSockets(CStr(WS)).CONTROL_SOCKET = WS2
    mcSockets(CStr(WS)).WSconnected WS, IP
    CreateDataSH = True
End Function

Public Sub RemoveSH(ByVal WS As Long, Optional ByVal Disconnect As Boolean = True)
On Error GoTo e
    If Not Disconnect Then mcSockets(CStr(WS)).SOCKET_ID = 0
    mcSockets.Remove CStr(WS)
e: If Err.Number Then Err.Clear
End Sub

Public Function ExistSH(ByVal WS As String) As Boolean
On Error GoTo e
    ExistSH = ObjPtr(mcSockets(WS)) <> 0
e: If Err.Number Then Err.Clear
End Function

Public Function DelegateSH(ByVal WS As Long, Obj As Object) As Boolean
On Error GoTo e
Dim WsM As Long
Dim WSP As Long

    If Obj.SOCKET_TYPE = FTP_DATA_SOCKET Then
        mcSockets(CStr(WS)).Dump Obj
    Else
        mcSockets(CStr(WS)).SOCKET_ID = 0
    End If
    
    mcSockets.Remove CStr(WS)
    Obj.SOCKET_ID = WS
    mcSockets.Add Obj, CStr(WS)
e:
End Function
Public Function GetSocketControl(ByVal SOCKET_DATA As Long) As Long
On Error GoTo e
    GetSocketControl = mcSockets(CStr(SOCKET_DATA)).CONTROL_SOCKET
e:
End Function
Public Function GetFreePortNum(ByVal LastPort As Long) As Long
    Do While Not mvIsFreePort(LastPort)
        LastPort = LastPort + 1
    Loop
    GetFreePortNum = LastPort
End Function

Public Function TCPListenerModule(ByVal NumPort As Long, Optional ByVal IPv6 As Boolean) As String
Dim lSize   As Long
Dim Out()   As Byte
Dim lRows   As Long
Dim i       As Long
Dim lPID    As Long

    Call GetExtendedTcpTable(0&, lSize, 1, IIf(IPv6, AF_INET6, AF_INET), TCP_TABLE_OWNER_PID_LISTENER, 0&)
    If lSize = 0 Then Exit Function
    
    ReDim Out(lSize - 1)
    If GetExtendedTcpTable(Out(0), lSize, 1, IIf(IPv6, AF_INET6, AF_INET), TCP_TABLE_OWNER_PID_LISTENER, 0&) <> 0 Then Exit Function
    MemCopy lRows, Out(0), 4    '/* Copy dwNumEntries in to lRows */
    
    If Not IPv6 Then
        Dim mRow4 As MIB_TCPROW_OWNER_PID
        For i = 0 To lRows - 1
            MemCopy mRow4, Out(4 + (i * LenB(mRow4))), LenB(mRow4)
            If mvGetAscPort(mRow4.dwLocalPort) = NumPort Then lPID = mRow4.dwOwningPid: Exit For
        Next
    Else
        Dim mRow6 As MIB_TCP6ROW_OWNER_PID
        
        For i = 0 To lRows - 1
           
            MemCopy mRow6, Out(4 + (i * LenB(mRow6))), LenB(mRow6)
            'PushLog "i  " & mvUnsigned(Ws_ntohs(mRow6.dwLocalPort)) & "   PID " & mRow6.dwOwningPid
            If mvUnsigned(Ws_ntohs(mRow6.dwLocalPort)) Then lPID = mRow6.dwOwningPid: Exit For
            'If mvGetAscPort(mRow6.dwLocalPort) = NumPort Then lPID = mRow6.dwOwningPid: Exit For
        Next
    End If
    
    If lPID Then TCPListenerModule = mvProcessModule(lPID)
    If lPID And Len(TCPListenerModule) = 0 Then TCPListenerModule = mvProcessModule(lPID, True)
    If Len(TCPListenerModule) = 0 Then TCPListenerModule = "Unable to open process"
    
    
End Function



'TODO: Private
'---------------------------------------------------------------------------------------------------------------------

Private Function mvIsFreePort(NumPort As Long) As Boolean
Dim tRow  As MIB_TCPROW
Dim Out() As Byte
Dim lRows As Long
Dim lSize  As Long
Dim i     As Long

    Call GetTcpTable(ByVal 0&, lSize, True)
    If lSize = 0 Then PushLog "GetTcpTable error in get stucture size": Exit Function
    
    ReDim Out(lSize - 1)
    If GetTcpTable(Out(0), lSize, 1) <> 0 Then PushLog "GetTcpTable error in get stucture data": Exit Function
    
    '/* First 4 bytes indicate the number of entries */
    '/* Copy dwNumEntries in to lRows                */
    MemCopy lRows, Out(0), 4
    
    For i = 0 To lRows - 1
        '/* First 4 bytes indicate the number of entries */
        '/* Get data and cast into a TcpRow stucture    */
        MemCopy tRow, Out(4 + (i * LenB(tRow))), LenB(tRow)
        If NumPort = mvGetAscPort(tRow.dwLocalPort) Then Exit Function
        
        '/* It can be used 'ntohs' api for get port number  */
        'Debug.Print Unsigned(ntohs(tRow.dwLocalPort))
    Next
    mvIsFreePort = True
    
End Function
Private Function mvGetAscPort(NumPort As String) As Long
    mvGetAscPort = Asc(Mid(NumPort, 1, 1))
    mvGetAscPort = mvGetAscPort * 256
    mvGetAscPort = mvGetAscPort + Asc(Mid(NumPort, 2, 1))
End Function
Private Function mvUnsigned(Value As Integer) As Long
    If Value < 0 Then mvUnsigned = Value + 65536 Else mvUnsigned = Value
End Function

Private Function mvProcessModule(ProcessID As Long, Optional ForceWMI As Boolean) As String
On Error GoTo e
Dim hProcess As Long
Dim sPath    As String * 256
Dim lRet     As Long
    
    If ForceWMI Then GoTo wmi_
    hProcess = OpenProcess(&H1F0FFF, False, ProcessID) ' PROCESS_ALL_ACCESS = &H1F0FFF
    If hProcess = 0 Then GoTo wmi_
   
    lRet = GetModuleFileNameExA(hProcess, 0, sPath, 256)
    If lRet Then
        mvProcessModule = Left$(sPath, lRet)
        mvProcessModule = Replace$(mvProcessModule, "\SystemRoot", Environ("WINDIR"))
        mvProcessModule = Replace$(mvProcessModule, "\??\", vbNullString)
        mvProcessModule = Right(mvProcessModule, Len(mvProcessModule) - InStrRev(mvProcessModule, "\"))
    End If
    CloseHandle (hProcess)
    Exit Function
    
wmi_:
 
    Dim Results As Object
    Dim Info    As Object
      
    Set Results = GetObject("Winmgmts:").ExecQuery("SELECT * FROM Win32_Process WHERE ProcessID=" & ProcessID)
    For Each Info In Results
        'Info.Caption
        'Info.CommandLine
        'Info.ExecutablePath
        'Info.Handle
        'Info.ProcessID
        mvProcessModule = Info.Caption
    Next
e:
    If Err.Number Then Err.Clear
End Function
