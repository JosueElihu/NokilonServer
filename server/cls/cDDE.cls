VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Component    : cDDE v1.5
'   Autor        : J. Elihu
'   Requirements : cSubclass
'   Credits      : http://leandroascierto.com/blog/clase-dde/
'--------------------------------------------------------------------------------

Option Explicit

Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function RegisterWindowMessage Lib "user32.dll" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Const WM_COPYDATA& = &H4A
Private Const WM_APP& = &H8000&

'------------------------------------------------------------------------------------------

'?EVENTS
Event Request(ByVal hWin As Long, ByVal Key As String, ByRef Cancel As Long)
Event Arrival(ByVal data As String, ByVal hWin As Long, ByVal Key As String, ByRef REPLY As String)
Event Disconnected(ByVal hWin As Long, ByVal Key As String)

Private Const SEND_DATA     As Long = 1
Private Const REPLY_DATA    As Long = 2
Private Const READED_REPLY  As Long = 3
Private Const READ_CONTENT  As Long = 4

'CUSTOM MESSAGES
Private WM_CONNECTION_REQUEST   As Long
Private WM_CONNECTION_CLOSE     As Long

Private Type COPYDATASTRUCT
  lData   As Long
  lSize   As Long
  lPtr    As Long
End Type

Public Enum DDEMODE
  DDENone
  DDEServer
  DDEClient
End Enum

Private c_SubClass As cSubClass

Private mPlugs      As Collection   '[CLIENT CONNECTIONS]
Private m_Main      As Long         '[MAIN WINDOW]      -> Used in Client mode
Private m_hWnd      As Long         '[WINDOW CALLBACK]  -> Used in Server mode
Private m_Mode      As DDEMODE
Private m_Key       As String
Private m_StrData   As String

Property Get Main() As Long: Main = m_Main: End Property
Property Get Client(ByVal Index As Long) As Long
On Error GoTo e
    Client = mPlugs(Index)(0)
e:
End Property
Property Get ClientCount() As Long: ClientCount = mPlugs.Count: End Property
Private Function ClientKey(ByVal hWin As Long) As String
On Error GoTo e
    ClientKey = mPlugs(CStr(hWin))(1)
e:
End Function

Private Sub Class_Initialize()
    Set c_SubClass = New cSubClass
    Set mPlugs = New Collection
    
    WM_CONNECTION_REQUEST = RegisterWindowMessage("WM_CONNECTION_REQUEST")
    WM_CONNECTION_CLOSE = RegisterWindowMessage("WM_CONNECTION_CLOSE")

End Sub
Private Sub Class_Terminate()
    Call StopDDE
    Set mPlugs = Nothing
    Set c_SubClass = Nothing
End Sub

Public Function InitDDE(ByVal Key As String) As Boolean
    Call StopDDE
    m_hWnd = CreateWindowEx(0, "Static", Key, 0, 0, 0, 0, 0, 0, 0, App.hInstance, 0&)
    If m_hWnd <> 0 Then
        Call SetSubclass(m_hWnd)
        InitDDE = True
        m_Key = Key
        m_Mode = DDEServer
    End If
End Function

Public Sub StopDDE()
Dim i As Long

    If (m_Mode = 0) Or (m_hWnd = 0) Then Exit Sub
    c_SubClass.UnSubclass m_hWnd
    
    Select Case m_Mode
        Case 1 'SERVER
            For i = mPlugs.Count To 1 Step -1
                SendMessage mPlugs(i)(0), WM_CONNECTION_CLOSE, m_hWnd, ByVal 0&
                mPlugs.Remove i
            Next
        Case 2 'CLIENT
            SendMessage m_Main, WM_CONNECTION_CLOSE, m_hWnd, ByVal 0&
            m_Main = 0
    End Select
    
    DestroyWindow m_hWnd
    m_hWnd = 0: m_Mode = DDENone
    m_Key = vbNullString
    
End Sub

Public Function Connect(ByVal ServerKey As String, ByVal ClientKey As String) As Boolean

    Call StopDDE
    m_Main = FindWindow("Static", ServerKey)
    If m_Main = 0 Then Exit Function
    
    m_hWnd = CreateWindowEx(0, "Static", ClientKey, 0, 0, 0, 0, 0, 0, 0, App.hInstance, 0&)
    If m_hWnd = 0 Then Exit Function
  
    If SendMessage(m_Main, WM_CONNECTION_REQUEST, m_hWnd, ByVal 0&) = 0 Then
        Call SetSubclass(m_hWnd)
        m_Mode = DDEClient
        m_Key = ServerKey
        Connect = True
    Else
        DestroyWindow m_hWnd
        m_hWnd = 0
    End If
    
End Function

Public Sub Disconnect()
    If m_Mode = DDEClient Then Call StopDDE
End Sub

Public Function SendData(ByVal data As String, Optional ByVal Client As String) As String
Dim CDS     As COPYDATASTRUCT
Dim lWin    As Long
Dim lRes    As String

    Select Case m_Mode
        Case 1 'SERVER - TO - CLIENT
            If Client = vbNullString Then Client = Me.Client(Me.ClientCount)
            If Not IsNumeric(Client) Then lWin = FindWindow("Static", Client) Else lWin = Client
        Case 2 'CLIENT - TO - SERVER
            lWin = m_Main
        Case Else: Exit Function
    End Select
    
    If IsWindow(lWin) = 0 Then Exit Function

    lRes = Space$(256)
    With CDS
        .lData = SEND_DATA
        .lSize = LenB(data)
        .lPtr = StrPtr(data)
    End With
    
    If SendMessage(lWin, WM_COPYDATA, m_hWnd, CDS) = READ_CONTENT Then
        SendData = m_StrData
        m_StrData = vbNullString
    End If

End Function

Public Function FindDDE(ByVal sKey As String) As Boolean
    FindDDE = FindWindow("Static", sKey)
End Function

Private Function SetSubclass(hWin As Long) As Boolean
    With c_SubClass
        .Subclass hWin, , , Me
        .AddMsg hWin, WM_COPYDATA, MSG_AFTER
        .AddMsg hWin, WM_SETTEXT, MSG_BEFORE
        .AddMsg hWin, WM_CONNECTION_REQUEST, MSG_AFTER
        .AddMsg hWin, WM_CONNECTION_CLOSE, MSG_AFTER
        .AddMsg hWin, WM_CLOSE, MSG_AFTER
    End With
End Function
Private Sub AddCollection(ByVal Key As String, ParamArray elements() As Variant)
On Error GoTo e
    mPlugs.Add elements, Key     '[hWin, WinName]
e:
End Sub

Private Function hWinText(hWin As Long) As String
Dim n   As Long
    hWinText = String$(256, Chr(0))
    n = GetWindowText(hWin, hWinText, 256)
    hWinText = Left(hWinText, n)
End Function


'- ordinal #1
Private Sub WndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lParamUser As Long)
Dim tmp As String

    Select Case uMsg
    
        Case WM_CONNECTION_REQUEST

            If IsWindow(wParam) = 0 Then Exit Sub
            tmp = hWinText(wParam)
            RaiseEvent Request(wParam, tmp, lReturn)
            If lReturn = 0 Then AddCollection wParam, wParam, tmp
            
        Case WM_COPYDATA
        
            Dim CDS     As COPYDATASTRUCT
            Dim sReply  As String

            Call CopyMemory(CDS, ByVal lParam, Len(CDS))
            
            tmp = String(CDS.lSize, 0)
            Call CopyMemory(ByVal tmp, ByVal CDS.lPtr, CDS.lSize)
            tmp = StrConv(tmp, vbFromUnicode)
            
            Select Case CDS.lData
                Case SEND_DATA
                
                    RaiseEvent Arrival(tmp, wParam, hWinText(wParam), sReply)
                    If Len(sReply) Then
                        CDS.lData = REPLY_DATA
                        CDS.lSize = LenB(sReply)
                        CDS.lPtr = StrPtr(sReply)
                        If SendMessage(wParam, WM_COPYDATA, m_hWnd, CDS) = READED_REPLY Then lReturn = READ_CONTENT
                    End If
                    
                Case REPLY_DATA
                    m_StrData = tmp
                    lReturn = READED_REPLY
                    
            End Select
            
            
        Case WM_CONNECTION_CLOSE
            
            Dim i   As Long
            
            Select Case m_Mode
                Case 1 'SERVER
                    For i = mPlugs.Count To 1 Step -1
                        If mPlugs(i)(0) = wParam Then
                            tmp = mPlugs(i)(1)
                            mPlugs.Remove i
                        End If
                    Next
                Case 2 'CLIENT
                    If wParam = m_hWnd Then
                        m_Mode = 0: m_Main = 0
                        tmp = m_Key
                        If m_hWnd <> 0 Then
                            c_SubClass.UnSubclass m_hWnd
                            DestroyWindow m_hWnd
                            m_hWnd = 0
                        End If
                    Else
                        tmp = hWinText(wParam)
                    End If
            End Select
            
            RaiseEvent Disconnected(wParam, tmp)
            
        Case WM_SETTEXT
            lReturn = -1
            bHandled = True
        Case Else
            
    End Select

End Sub


