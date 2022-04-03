VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00FAF7F5&
   BorderStyle     =   0  'None
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   390
   ScaleWidth      =   450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin NokilonServer.JImageListEx iml 
      Left            =   0
      Top             =   0
      _ExtentX        =   794
      _ExtentY        =   688
      Count           =   4
      Data_0          =   "FrmMain.frx":22662
      Data_1          =   "FrmMain.frx":23EC3
      Data_2          =   "FrmMain.frx":24C40
      Data_3          =   "FrmMain.frx":25966
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Type NOTIFYICONDATA
  cbSize        As Long
  hWnd          As Long
  uID           As Long
  uFlags        As Long
  uCallbackMessage As Long
  hIcon         As Long
  szTip         As String * 128
  dwState       As Long
  dwStateMask   As Long
  szInfo        As String * 256
  uTimeout      As Long
  szInfoTitle   As String * 64
  dwInfoFlags   As Long
End Type

Private WithEvents c_DDE    As cDDE
Attribute c_DDE.VB_VarHelpID = -1
Private WithEvents c_Timer  As cTimer
Attribute c_Timer.VB_VarHelpID = -1

Private c_mnu       As cMenuApi
Private hWndGUI     As Long
Private mtSysTray   As NOTIFYICONDATA
Private msState     As SERVER_STATE

Private Sub Form_Load()
Dim mObj    As cConfig
    
    If App.PrevInstance Then End
    If Not mWinSock.WsStart(Me) Then
        MsgBox "Winsock Error" & vbNewLine & String$(79, "•") & vbNewLine & "Source" & vbTab & ": WSAStartup" & _
                vbNewLine & "Lib" & vbTab & ": ws2_32.dll" & _
                vbNewLine & String$(79, "-") & vbNewLine & "Error al iniciar el modulo de red", vbCritical, SERVER_NAME
        End
    End If
    
    Me.Caption = "Nokilon " & App.Major & "." & App.Minor & "." & App.Revision
    
    Set c_DDE = New cDDE
    Set c_mnu = New cMenuApi
    Set c_Timer = New cTimer
    Set dbcnn = SQLite.Connection(App.Path & "\noz-db3")
    c_DDE.InitDDE ("Nokilon-SERVER-DDE")
    
    c_mnu.AddItem 100, "Administrar", , , iml.hBitmap(0, 16, 16)
    With c_mnu.AddSubMenu("SERVER", "Servidor", iml.hBitmap(1, 16, 16))
        .AddItem 101, "Encender": .AddItem 102, "Apagar"
        .ItemRadioCheck(0) = True: .ItemRadioCheck(1) = True
    End With
    c_mnu.AddItem 0, "", True
    c_mnu.AddItem 103, "Acerca de...", , , iml.hBitmap(2, 16, 16, &HB78255)
    c_mnu.AddItem 0, "", True
    c_mnu.AddItem 104, "Salir", , , iml.hBitmap(3, 16, 16, , 80)
    c_mnu.ItemDefault(0) = True
    
    Set mObj = New cConfig
    mObj.ReadAll dbcnn
    SERVER.PORT = mObj.GetValue2("Port", 21)
    SERVER.STREAM_SIZE = mObj.GetValue2("Stream-size", 16)
    SERVER.SOCKET_TIME_OUT = mObj.GetValue2("Time-out", 0)
    
    CONFIG.Rmvct = mObj.GetValue("Rmvct", False)
    CONFIG.Rmvit = mObj.GetValue("Rmvit", False)
    CONFIG.Console = mObj.GetValue2("Events", 7)
    CONFIG.Sockets = mObj.GetValue2("Sockets", 3)
    
    If SERVER.PORT = 0 Then SERVER.PORT = 21
    If SERVER.SOCKET_TIME_OUT < 0 Then SERVER.SOCKET_TIME_OUT = 0
    If Not SERVER.STREAM_SIZE > 0 Then SERVER.STREAM_SIZE = 16
    
    Call dbcnn.Execute("CREATE TABLE temp.transfers (ws,tstamp,mode,user,fpath,state,cbytes,tbytes,tpval);")
    
    If App.LogMode <> 0 Then Me.Visible = False
    Call mvSetTray
    Call CheckDataBase
    Call StartServer
    Call ConnectToGUI
    iml.Clear

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Msg As Long
Const WM_LBUTTONDBLCLK = &H203
Const WM_LBUTTONDOWN = &H201
Const WM_RBUTTONUP = &H205

    Msg = x / Screen.TwipsPerPixelX
    If Msg = WM_RBUTTONUP Then
        Call SetForegroundWindow(Me.hWnd)
        Select Case c_mnu.PopupMenu(, , Me.hWnd, TPM_CENTERALIGN)
            Case 100:
                    If hWndGUI Then
                        c_DDE.SendData "000", hWndGUI
                    Else
                        If PathExist(App.Path & "\nokilon-server-gui.exe") Then Shell2 App.Path & "\nokilon-server-gui.exe"
                    End If
            Case 101: StartServer
            Case 102: StopServer
            Case 103: FrmAbout.Show , Me
            Case 104: Unload Me
        End Select
    ElseIf Msg = WM_LBUTTONDBLCLK Then
        If hWndGUI Then
            c_DDE.SendData "000", hWndGUI
        Else
            If PathExist(App.Path & "\nokilon-server-gui.exe") Then Shell2 App.Path & "\nokilon-server-gui.exe"
        End If
    End If
    
End Sub
Private Sub Form_Resize()
'
End Sub


'TODO: WINSOCK CALLBACKS
'======================================================================================================================
Public Sub WSOnConnect(WS As Long, IP As String, SERVER_SOCKET As Long, ByVal IPv6 As Boolean)
    If CreateSH(WS, IP, SERVER_SOCKET) Then
        mcSockets(CStr(WS)).WSconnected WS, IP
    End If
End Sub
Public Sub WSOnWritable(WS As Long)
On Error GoTo e
    mcSockets(CStr(WS)).WSwritable (WS)
e:
End Sub
Public Sub WSOnRead(WS As Long, data As String, ByVal lBytes As Long)
'On Error Resume Next
    mcSockets(CStr(WS)).WSarrival WS, data, lBytes
    SERVER.BYTES_RECEIVED = SERVER.BYTES_RECEIVED + lBytes
End Sub
Public Sub WSOnSend(WS As Long, ByVal Result As Boolean, ByVal lBytes As Long, ByVal lErr As WS_ERROR)
On Error GoTo e

    SERVER.BYTES_SENT = SERVER.BYTES_SENT + lBytes
    mcSockets(CStr(WS)).AddBytes lBytes, eBytesSent
e:
End Sub
Public Sub WSOnDisconnect(WS As Long)
On Error GoTo e
     mcSockets(CStr(WS)).WSdisconnected WS
e:
End Sub
Public Sub WSOnError(ByVal Num As WS_ERROR, ByVal Description As String, Source As String)
'''
End Sub


'TODO: FTP CALLBACKS
'======================================================================================================================
Public Sub FTPcommandReceived(WS As Long, ByVal ftpcmd As String, User As String)
    PushLog "[" & User & "] " & ftpcmd, enmPSInfo, mclFTP
End Sub
Public Sub FTPcommandSent(WS As Long, ByVal ftpcmd As String, User As String)
Dim lStyle As PrintStyle
    Select Case Val(Left(ftpcmd, InStr(ftpcmd, " ")))
        Case 100: lStyle = enmPSInfo
        Case Is >= 400: lStyle = enmPSError
        Case Else: lStyle = enmPSSuccess
    End Select
    PushLog "[" & User & "] " & ftpcmd, lStyle, mclFTP
End Sub
Public Sub FTPuserlogin(WS As Long, IP As String, UserName As String, BytesReceived As Currency, BytesSend As Currency)
    If hWndGUI Then c_DDE.SendData "101" & AddNulls(WS, IP, UserName, BytesReceived, BytesSend), hWndGUI
End Sub
Public Sub FTPuserlogout(WS As Long)
    If hWndGUI Then c_DDE.SendData "102" & WS, hWndGUI
End Sub

Public Sub FTPtransferStart(WS As Long, FileName As String, FTPMode As FTPTransferMode, lTotalBytes As Currency, tStamp As String, WsM As Long, UserName As String)
On Error GoTo e

    '(ws,tstamp,mode,user,fpath,state,cbytes,tbytes,tpval);
    db_exec "INSERT INTO temp.transfers VALUES(?,?,?,?,?,?,?,?,?);", Array(WS, tStamp, FTPMode, UserName, FileName, Null, 0, CStr(lTotalBytes), 0)
    
    If hWndGUI = 0 Then Exit Sub
    c_DDE.SendData "201" & AddNulls(WS, FileName, FTPMode, lTotalBytes, tStamp, WsM, UserName), hWndGUI
e:
End Sub
Public Sub FTPtransferEnd(WS As Long, FTPMode As FTPTransferMode, lBytes As Currency, Result As FTPResult, lPercent As Long, WsM As Long)
Dim tmp As String
    
    Select Case Result
        Case emSuccess: tmp = IIf(FTPMode = emDownload, "Completado - 100%", "Completado")
        Case Else
            tmp = IIf(Result = emError, "Error", "Cancelado")
            If FTPMode = emDownload Then tmp = tmp & " - " & lPercent & "%"
    End Select
        
    '/* ws,tstamp,mode,user,fpath,state,cbytes,tbytes,tpval */
    
    If (CONFIG.Rmvit And Result <> emSuccess) Or CONFIG.Rmvct Then
        dbcnn.Execute "DELETE FROM temp.transfers WHERE ws=" & WS & ";"
    Else
        Select Case FTPMode
            Case emDownload
                db_exec "UPDATE temp.transfers SET ws=?, state=?, cbytes=?, tpval=? WHERE ws=?", Array(0, tmp, CStr(lBytes), lPercent, WS)
            Case emUpload
                db_exec "UPDATE temp.transfers SET ws=?, state=?, cbytes=?,tbytes=?, tpval=? WHERE ws=?", Array(0, tmp, CStr(lBytes), CStr(lBytes), lPercent, WS)
        End Select
    End If
    
    If hWndGUI = 0 Then Exit Sub
    c_DDE.SendData "202" & AddNulls(WS, FTPMode, lBytes, Result, lPercent, WsM), hWndGUI
    
End Sub


'TODO: DDE EVENTS
'======================================================================================================================

Private Sub c_DDE_Request(ByVal Plug As Long, ByVal Key As String, Cancel As Long)
    If Key = "Nokilon-GUI" Then hWndGUI = Plug
End Sub
Private Sub c_DDE_Arrival(ByVal data As String, ByVal Plug As Long, ByVal Key As String, REPLY As String)
On Error Resume Next
Dim sCmd  As String
Dim mvObj As cWsCnn
Dim i     As Long

    sCmd = Left$(data, 3)
    data = Right$(data, Len(data) - 3)
    
    Select Case sCmd
        Case "000" 'Prev Instance
        Case "001" 'START SOCKET
        
            Call StartServer
            REPLY = Abs((SERVER.SOCKET4 <> 0) Or (SERVER.SOCKET6 <> 0))
        
        Case "002" 'STOP SERVICE
        
            Call StopServer
            Unload Me
            
        Case "004" 'SOCKET STATUS
        
            REPLY = Abs((SERVER.SOCKET4 <> 0) Or (SERVER.SOCKET6 <> 0))
            
        Case "005" 'SERVER INFO

            '-=  NAME, SOCKET, IP, PORT, BYTES_SENT, BYTES_RECEIVED
            REPLY = AddNulls(SERVER_NAME, Abs(SERVER.SOCKET4 Or SERVER.SOCKET6), mWinSock.LocalHostIP, SERVER.PORT, SERVER.BYTES_SENT, SERVER.BYTES_RECEIVED)
            
        Case "006" 'SERVER INFO2

            '-= BYTES_SENT, BYTES_RECEIVED, STATE
            REPLY = AddNulls(SERVER.BYTES_SENT, SERVER.BYTES_RECEIVED, Abs((SERVER.SOCKET4 <> 0) Or (SERVER.SOCKET6 <> 0)))
            
        Case "007" 'SERVER INFO3
        
            '-= IP, PORT
            REPLY = AddNulls(mWinSock.LocalHostIP, SERVER.PORT)
            
        Case "010" 'SOCKET BYTES
            REPLY = AddNulls(mcSockets(data).Bytes(eBytesSent), mcSockets(data).Bytes(eBytesReceived))
            
        Case "011" 'SOCKET EXITS
            If ExistSH(data) Then REPLY = True
        Case "012" 'DISCONNECT USER
            mcSockets(data).CloseSocket True
        Case "013" 'CANCEL TRANSFER
            mcSockets(data).Cancel
        Case "014" 'CLEAR TRANSFERS
        
            'dbcnn.Execute "DELETE FROM temp.transfers WHERE ws=0;"
            'dbcnn.Execute "DELETE FROM temp.transfers WHERE ws=0 AND state NOT LIKE 'Completado%';"
            'dbcnn.Execute "DELETE FROM temp.transfers WHERE ws=0 AND state LIKE 'Completado%';"
             dbcnn.Execute "DELETE FROM temp.transfers WHERE ws=0;"
             
        Case "020" 'LOG SERVER STATS
            
            If SERVER.SOCKET4 And SERVER.SOCKET6 Then
                PushLog "[SERVER] FTP Service is already running on port " & SERVER.PORT, , mclSERVER
            ElseIf SERVER.SOCKET4 = 0 And SERVER.SOCKET6 = 0 Then
                PushLog "[SERVER] FTP service is not listening", , mclSERVER
            Else
                
                If SERVER.SOCKET4 Then _
                    PushLog "[SERVER] FTP Service is already listening on port " & SERVER.PORT & " for IPv4", , mclSERVER
                
                If SERVER.SOCKET6 Then _
                    PushLog "[SERVER] FTP Service is already listening on port " & SERVER.PORT & " for IPv6", , mclSERVER
        
                If SERVER.SOCKET4 = 0 And (CONFIG.Sockets And SS_IPv4) Then _
                    PushLog "[SERVER] Port " & SERVER.PORT & " in IPv4 in use by " & Chr(34) & TCPListenerModule(SERVER.PORT) & Chr(34) & " FTP service is not listening in IPv4", enmPSError, mclSERVER
                
                If SERVER.SOCKET6 = 0 And (CONFIG.Sockets And SS_IPv6) Then _
                    PushLog "[SERVER] Port " & SERVER.PORT & " in IPv6 in use by " & Chr(34) & TCPListenerModule(SERVER.PORT, True) & Chr(34) & " FTP service is not listening in IPv6", enmPSError, mclSERVER
                    
            End If
        
        'TODO: SETUPS SERVER
        '=============================================================================================================
        Case "050" 'Startup Socket
        Case "051" 'Server port
        
            If SERVER.PORT = CInt(Val(data)) Then Exit Sub
            For Each mvObj In mcSockets
                 mvObj.CloseSocket True, 1
            Next
            If Not StopServer Then Exit Sub
            SERVER.PORT = CInt(Val(data))
            If SERVER.PORT = 0 Then SERVER.PORT = 21
            Call StartServer

        Case "052" 'Stream Size
            SERVER.STREAM_SIZE = CInt(Val(data))
            If Not SERVER.STREAM_SIZE > 0 Then SERVER.STREAM_SIZE = 8
        Case "053" 'Time out
            SERVER.SOCKET_TIME_OUT = CInt(Val(data))
            If SERVER.SOCKET_TIME_OUT < 0 Then SERVER.SOCKET_TIME_OUT = 0
        Case "054" 'Rmvct
            CONFIG.Rmvct = CBool(data)
        Case "055" 'Rmvit
            CONFIG.Rmvit = CBool(data)
        Case "056" 'Events logs
            CONFIG.Console = Abs(Val(data))
            
        Case "057" 'IPv4 && IPv6
            
            Dim lFlag As SOCKET_FLAGS

            If CONFIG.Sockets = CLng(Val(data)) Then Debug.Print "Salir": Exit Sub
            lFlag = CLng(Val(data))
        
            If (lFlag And SS_IPv4) Then
                If SERVER.SOCKET4 = 0 Then Call StartServer(True, False)
            Else
                If SERVER.SOCKET4 <> 0 Then
                    For Each mvObj In mcSockets
                         If mvObj.SOCKET_FAMILY = AF_INET Then mvObj.CloseSocket True, 1
                    Next
                    Call StopServer(False, True, False)
                End If
            End If
            
            
            If (lFlag And SS_IPv6) Then
                If SERVER.SOCKET6 = 0 Then Call StartServer(False, True)
            Else
                If SERVER.SOCKET6 <> 0 Then
                    For Each mvObj In mcSockets
                         If mvObj.SOCKET_FAMILY = AF_INET6 Then mvObj.CloseSocket True, 1
                    Next
                    Call StopServer(False, False, True)
                End If
            End If
            CONFIG.Sockets = lFlag
            
        
        'TODO: FTP USERS
        '=============================================================================================================
        
        Case "100" 'LOAD USERS
        
            '{WS, IP ,UserName, BytesSent, BytesReceived}
            For Each mvObj In mcSockets
                c_DDE.SendData "101" & AddNulls(mvObj.SOCKET_ID, mvObj.SOCKET_IP, mvObj.UserName, mvObj.Bytes(eBytesSent), mvObj.Bytes(eBytesReceived)), Plug
            Next
        
        Case "101" '/* User login  (server->gui) */
        Case "102" '/* User logout (server->gui) */
        
        Case "105" '/* User deleted || locked || renamed */
            For Each mvObj In mcSockets
                If mvObj.UserID = data Then mvObj.CloseSocket True
            Next
        Case "106" '/*  Update user mounts   */
            For Each mvObj In mcSockets
                If mvObj.UserID = data Then mvObj.UpdateMounts
            Next
        
        'TODO: FILE TRANSFERS
        '=============================================================================================================
        
        Case "200" 'LOAD FILES
        
            Dim WS1     As Long
            Dim WS2     As Long
            
            With dbcnn.Query("SELECT * FROM temp.transfers;")
                Do While .Step = SQLITE_ROW
                    WS1 = Val(.Value(0))
                    If WS1 <> 0 And Not ExistSH(WS1) Then WS1 = 0
                    WS2 = IIf(WS1, GetSocketControl(WS1), 0)
                    
                    '--------------------------------------------------------------------------
                    ' ws, tstamp, mode, user, fpath, state, cbytes, tbytes, tpval | ws2, speed
                    ' 0     1      2    3       4      5      6        7      8   |  9    10
                    '--------------------------------------------------------------------------
                    
                    If WS1 Then
                       
                        c_DDE.SendData "220" & AddNulls(WS1, .Value(1), .Value(2), .Value(3), .Value(4), _
                            IIf(.Value(2) = 1, "Descargando - " & mcSockets(CStr(WS1)).Percent & "%", "Recibiendo"), _
                            mcSockets(CStr(WS1)).CurrentBytes, _
                            IIf(.Value(2) = 1, mcSockets(CStr(WS1)).TotalBytes, "-"), _
                            mcSockets(CStr(WS1)).Percent, WS2, mcSockets(CStr(WS1)).Speed), hWndGUI
                            
                    Else
                        c_DDE.SendData "220" & AddNulls(WS1, .Value(1), .Value(2), .Value(3), .Value(4), .Value(5), .Value(6), .Value(7), .Value(8), WS2, "-"), hWndGUI
                    End If
                Loop
            End With
            
        Case "201" '/* Begin file transfer (server->gui) */
        Case "202" '/* End file transfer   (server->gui) */
        Case "205" 'TRANSFER INFO
        
            '---------------------------------------------
            ' Mode, Percent, CurrentBytes, Speed
            '---------------------------------------------
            If ExistSH(data) Then
                REPLY = AddNulls(mcSockets(data).Mode, mcSockets(data).Percent, mcSockets(data).CurrentBytes, mcSockets(data).Speed)
            End If
            
        Case "220" '/* file info (server->gui) */
            
            
        'TODO: SERVER_PORT SOCKETS
        '=============================================================================================================
        
        Case "300" '
        Case "301" 'ADD SERVER_PORT
        Case "302" 'CHANGE SERVER_PORT
        Case "303" 'REMOVE SERVER_PORT
        Case Else
    End Select
    
End Sub
Private Sub c_DDE_Disconnected(ByVal Plug As Long, ByVal Key As String)
    If Plug = hWndGUI Then hWndGUI = 0
End Sub
'======================================================================================================================

Private Sub c_Timer_Timer(ByVal ThisTime As Long)

    If msState = SS_WAIT Then
        Call StartServer
    Else
        c_Timer.DestroyTimer
    End If
    Exit Sub
    
    If (CONFIG.Sockets And SS_IPv4) And SERVER.SOCKET4 = 0 Then
        StartServer True, False
    End If
    
    If (CONFIG.Sockets And SS_IPv6) And SERVER.SOCKET6 = 0 Then
        StartServer False, True
    End If
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 And App.LogMode <> 0 Then
        Me.Visible = False
        Cancel = 1
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    
    Set mcSockets = Nothing
    
    Call mWinSock.WsTerminate
    Call dbcnn.Execute("DROP TABLE IF EXISTS temp.transfers;")
    
    Set c_Timer = Nothing
    Set c_DDE = Nothing
    Set c_mnu = Nothing
    Set dbcnn = Nothing
    
    Call Shell_NotifyIcon(&H2, mtSysTray)
    
End Sub


'TODO: PUBLIC SUBS
'======================================================================================================================
Public Function StartServer(Optional IPv4 As Boolean = True, Optional IPv6 As Boolean = True) As Boolean
Dim mbErr4  As Boolean
Dim mbErr6  As Boolean

    If CONFIG.Sockets And SS_IPv4 Then
        SERVER.SOCKET4 = mWinSock.WsListen(SERVER.PORT)
        mbErr4 = mWinSock.WSLastError = WSAEADDRINUSE
    End If
    
    If CONFIG.Sockets And SS_IPv6 Then
        SERVER.SOCKET6 = mWinSock.WsListen(SERVER.PORT, True)
        mbErr6 = mWinSock.WSLastError = WSAEADDRINUSE
    End If
    

    If (SERVER.SOCKET4 And SERVER.SOCKET6) And (IPv4 And IPv6) Then
        PushLog "[SERVER] FTP Service is already running on port " & SERVER.PORT, , mclSERVER
    Else
    
        If mbErr4 And mbErr6 Then
            PushLog "[SERVER] Port " & SERVER.PORT & " in use by " & Chr(34) & TCPListenerModule(SERVER.PORT) & Chr(34) & "!", enmPSError, mclSERVER
            PushLog "[SERVER] FTP Service WILL NOT start without the configured ports free", enmPSError, mclSERVER
            GoTo Skip_
        End If
        
        If SERVER.SOCKET4 And IPv4 And Not mbErr4 Then _
            PushLog "[SERVER] FTP Service is already running on port " & SERVER.PORT & " for IPv4", , mclSERVER
       
        If SERVER.SOCKET6 And IPv6 And Not mbErr6 Then _
            PushLog "[SERVER] FTP Service is already running on port " & SERVER.PORT & " for IPv6", , mclSERVER

        
        If IPv4 And mbErr4 Then _
            PushLog "[SERVER] Port " & SERVER.PORT & " in IPv4 in use by " & Chr(34) & TCPListenerModule(SERVER.PORT) & Chr(34) & " FTP service is not listening in IPv4", enmPSError, mclSERVER
       
        If IPv6 And mbErr6 Then _
            PushLog "[SERVER] Port " & SERVER.PORT & " in IPv6 in use by " & Chr(34) & TCPListenerModule(SERVER.PORT, True) & Chr(34) & " FTP service is not listening in IPv6", enmPSError, mclSERVER
       
    End If
    
Skip_:

    c_mnu.SubMenu("SERVER").ItemCheck(0) = (SERVER.SOCKET4 <> 0) Or (SERVER.SOCKET6 <> 0)
    c_mnu.SubMenu("SERVER").ItemCheck(1) = (SERVER.SOCKET4 = 0) And (SERVER.SOCKET6 = 0)
    If hWndGUI Then c_DDE.SendData "020" & ((SERVER.SOCKET4 <> 0) Or (SERVER.SOCKET6 <> 0)), hWndGUI 'Notify CHANGED STATUS
    StartServer = (SERVER.SOCKET4 <> 0) Or (SERVER.SOCKET6 <> 0)
    
    If Not StartServer Then
        c_Timer.CreateTimer 5000
        msState = SS_WAIT
    Else
        msState = SS_LISTEN
    End If
    
    
End Function

Public Function StopServer(Optional DestroySH As Boolean = True, Optional IPv4 As Boolean = True, Optional IPv6 As Boolean = True) As Boolean
Dim mbErr4  As Boolean
Dim mbErr6  As Boolean

    If DestroySH Then Set mcSockets = Nothing
    
    If SERVER.SOCKET4 And IPv4 Then
        If mWinSock.WsClose(SERVER.SOCKET4) Then SERVER.SOCKET4 = 0 Else mbErr4 = True
    End If
    
    If SERVER.SOCKET6 And IPv6 Then
        If mWinSock.WsClose(SERVER.SOCKET6) Then SERVER.SOCKET6 = 0 Else mbErr6 = True
    End If
    
    If (Not mbErr4 And Not mbErr6) And (IPv4 And IPv6) Then
        PushLog "[SERVER] FTP Service stopped ", , mclSERVER
    ElseIf Not mbErr4 And IPv4 Then
        PushLog "[SERVER] FTP Service stopped listening in IPv4", , mclSERVER
    ElseIf Not mbErr6 And IPv6 Then
        PushLog "[SERVER] FTP Service stopped listening in IPv6", , mclSERVER
    Else
        
    End If
    
    c_mnu.SubMenu("SERVER").ItemCheck(0) = (SERVER.SOCKET4 <> 0) Or (SERVER.SOCKET6 <> 0)
    c_mnu.SubMenu("SERVER").ItemCheck(1) = (SERVER.SOCKET4 = 0) And (SERVER.SOCKET6 = 0)
    If hWndGUI Then c_DDE.SendData "020" & ((SERVER.SOCKET4 <> 0) Or (SERVER.SOCKET6 <> 0)), hWndGUI 'Notify CHANGED STATUS
    StopServer = ((SERVER.SOCKET4 = 0) And (SERVER.SOCKET6 = 0))
    
End Function

Property Get DDE() As cDDE: Set DDE = c_DDE: End Property
Property Get GUI() As Long: GUI = hWndGUI: End Property



'TODO: PRIVATE SUBS
'======================================================================================================================
Private Function AddNulls(ParamArray elements() As Variant) As String
On Error GoTo e
Dim i As Long
    
    If UBound(elements) = 0 Then AddNulls = elements(0): Exit Function
    For i = 0 To UBound(elements) - 1
        AddNulls = AddNulls & elements(i) & vbNullChar
    Next
    If UBound(elements) > 0 Then AddNulls = AddNulls & elements(UBound(elements))
e:
End Function
Private Sub ConnectToGUI()
    With New cDDE
        If .Connect("Nokilon-GUI2", "SVR-START") Then .SendData "001"
    End With
End Sub
Private Sub mvSetTray()

    ' {NIF_MESSAGE = &H1, NIF_ICON = &H2, NIF_TIP = &H4, NIF_STATE = &H8, NIF_INFO = &H10}
    
    With mtSysTray
        .cbSize = Len(mtSysTray)
        .hWnd = Me.hWnd
        .uID = 1&
        .uFlags = &H2 Or &H10 Or &H1 Or &H4
        .uCallbackMessage = WM_LBUTTONDOWN
        .hIcon = IconFromStream(iml.Stream(0, 16, 16))
        .szTip = IIf(App.LogMode, SERVER_NAME & Chr$(0), "IDE " & Chr$(0))
    End With
    
    Call Shell_NotifyIcon(&H0, mtSysTray)
End Sub
