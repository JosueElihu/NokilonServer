Attribute VB_Name = "mFtpHelper"
Option Explicit

Public Enum enmLoginResult
    LOGIN_SUCCESS
    LOGIN_INVALID
    LOGIN_PASSWORD_REQUIRED
    LOGIN_LOCKED
End Enum

'Public Enum enmUserPermissions
'    UP_READ = 0
'    UP_DOWNLOAD = 1
'    UP_UPLOAD = 2
'    UP_RENAME = 4
'    UP_REMOVE = 8
'    UP_MKDIR = 16
'End Enum

Public Enum FTPAccess
    DISABLED_ = 0
    READ_ONLY = 1
    READ_WRITE = 2
End Enum

Public Type mtUSER
  m_id      As Long
  m_Name    As String
  m_Pwd     As String
  m_logged  As Boolean
End Type

Public Enum PrintStyle
  enmPSNone
  enmPSSuccess
  enmPSInfo
  enmPSError
  enmPSWarning
  enmPSServer
End Enum
Public Enum enmByteType
  eBytesReceived = 1
  eBytesSent = 2
End Enum
Public Enum CONSOLE_LOGS
  mclNONE = 0
  mclSERVER = 1
  mclFTP = 2
  mclERRORS = 4
End Enum



Public Function Login(Args As String, miUser As mtUSER) As enmLoginResult

    If miUser.m_logged Then Exit Function

    '/* Check Password */
    If miUser.m_Name <> vbNullString And miUser.m_Pwd <> vbNullString Then
        If Sha1(Args) <> miUser.m_Pwd Then Login = LOGIN_INVALID: GoTo e
        miUser.m_Pwd = vbNullString
        Login = LOGIN_SUCCESS: GoTo e
    End If

    With dbcnn.Query("SELECT * FROM users WHERE user='" & Args & "';")
        If .Step <> SQLITE_ROW Then Login = LOGIN_INVALID: GoTo e
        miUser.m_id = .Value(0)
        miUser.m_Name = .Value(1)
        miUser.m_Pwd = .Value(2)
        If CBool(Val(.Value(3))) = False Then Login = LOGIN_LOCKED: GoTo e
    End With

    If miUser.m_Pwd <> vbNullString Then Login = LOGIN_PASSWORD_REQUIRED: GoTo e
    Login = LOGIN_SUCCESS
e:
    If Login = LOGIN_SUCCESS Then miUser.m_logged = True
End Function

Public Function ParseUrl(ByVal URL As String) As String
On Error GoTo e
    If URL = "/" Then ParseUrl = "/": Exit Function
    If Left(URL, 1) <> "/" Then URL = "/" & URL
    If Right(URL, 1) = "/" Then URL = Left(URL, Len(URL) - 1)
    ParseUrl = URL
e:
End Function

Public Function ParentVirtualPath(ByVal virtual_path As String) As String
On Error GoTo e
    If Right(virtual_path, 1) = "/" Then virtual_path = Left(virtual_path, Len(virtual_path) - 1)
    ParentVirtualPath = Left(virtual_path, InStrRev(virtual_path, "/") - 1)
e:
End Function

Public Function FTPDate(Fecha As String) As String
    FTPDate = Format(Fecha, "mmm dd hh:mm")
    FTPDate = Replace(FTPDate, "ene", "Jan", compare:=vbTextCompare)
    FTPDate = Replace(FTPDate, "abr", "Apr", compare:=vbTextCompare)
    FTPDate = Replace(FTPDate, "ago", "Aug", compare:=vbTextCompare)
    FTPDate = Replace(FTPDate, "dic", "Dec", compare:=vbTextCompare)
    FTPDate = Replace(FTPDate, "set", "Sep", compare:=vbTextCompare)
End Function
Public Function MakePasvReply(ByVal lNumPort As Long, Optional Extended As Boolean) As String
    If Not Extended Then
        Dim a, b As Byte
        a = lNumPort Mod 256
        b = (lNumPort - a) / 256
        MakePasvReply = "(" & Replace(mWinSock.LocalHostIP, ".", ",") & "," & b & "," & a & ")"
    End If
End Function
Public Function CmdNeesdAuth(Cmd As String) As Boolean
    CmdNeesdAuth = InStr(no_needs_auth, Cmd) = 0
End Function
Public Function ParseMountPath(ByVal vPath As String, cColl As Collection, Optional SaveIndex As Long) As String
On Error GoTo e
Dim mMask   As String
Dim tmp     As String
Dim i       As Long

    If InStr(vPath, "/") Then vPath = Replace(vPath, "/", "\")                          ' /Unidad/    ->    \Unidad\
    If Left(vPath, 1) = "\" Then vPath = Right(vPath, Len(vPath) - 1)                   ' \Unidad\    ->    Unidad\
    'If InStr(vPath, "\") = 0 And Right$(vPath, 1) <> "\" Then vPath = vPath & "\"      ' Unidad      ->    Unidad\
    
    If InStr(vPath, "\") Then mMask = Left(vPath, InStr(vPath, "\") - 1) Else mMask = vPath
    
    '/* cColl = {name, path, access} */
    For i = 1 To cColl.Count
        If Trim(cColl(i)(0)) = mMask Then
            tmp = cColl(i)(1)
            If Right$(tmp, 1) = "\" Then tmp = Left$(tmp, Len(tmp) - 1)
            ParseMountPath = tmp & Right(vPath, Len(vPath) - Len(mMask))
            SaveIndex = i
        End If
    Next
e:
End Function

Public Sub LoadMountPoints(ByVal muid As Long, cColl As Collection)
Dim i As Long
Dim lFlag As FTPAccess

    
    Set cColl = New Collection
    With dbcnn.Query("SELECT * FROM mounts WHERE id_user='" & muid & "';")
        Do While .Step = SQLITE_ROW
        
            If PathExist(.Value(3)) = 0 Then GoTo next_
            Select Case UCase$(.Value(4)) 'ACCESS
                Case "READ ONLY": lFlag = READ_ONLY
                Case "READ + WRITE": lFlag = READ_WRITE
                Case Else:  lFlag = DISABLED_
            End Select
            If lFlag = DISABLED_ Then GoTo next_
            
            '/* {name, path, access} */
            mvAddCollection cColl, .Value(2), .Value(3), lFlag
            
next_:
        Loop
    End With
End Sub
Public Function GetHomeMount(VirtualPath As String, LocalPath As String) As Boolean
'On Error Resume Next

'    If Mounts = 0 Then
'        VirtualPath = "/"
'        LocalPath = ""
'        Exit Function
'    End If
    
    'If Mounts = 1 Then
        'VirtualPath = "/" & m_mount(0).sName '& "/"
        'LocalPath = ParsePath(VirtualPath)
    'End If
    VirtualPath = "/"
    LocalPath = ""
    
End Function


Private Sub mvAddCollection(mColl As Collection, ParamArray elements() As Variant)
On Error GoTo e
    mColl.Add elements
e:
End Sub
