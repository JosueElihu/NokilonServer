Attribute VB_Name = "Global"
Option Explicit

' /* UTF8 - UTF16 */
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

'/* DPI */
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal Hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal Hdc As Long) As Long

'/* ICONS */
Private Declare Function CreateIconFromResourceEx Lib "user32" (ByRef presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

'/* INI */
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Enum SERVER_STATE
    SS_WAIT
    SS_LISTEN
End Enum

Public Enum SOCKET_FLAGS
    SS_IPv4 = 1     '/* Server listen in IPv4   */
    SS_IPv6 = 2     '/* Server listen in IPv6   */
End Enum

Public Type mtConfig
  Rmvct             As Boolean
  Rmvit             As Boolean
  Console           As CONSOLE_LOGS
  Sockets           As SOCKET_FLAGS
End Type

Public Type mtSERVER
  SOCKET4           As Long
  SOCKET6           As Long
  PORT              As Long
  STREAM_SIZE       As Long
  SOCKET_TIME_OUT   As Long
  BYTES_RECEIVED    As Currency
  BYTES_SENT        As Currency
End Type

Public Const no_needs_auth = "AUTH,FEAT,HELP,NOOP,OPTS,USER,PASS,QUIT,SYST"
Public Const SERVER_NAME = "Nokilon v1.5"
Public Const ENDL = vbCrLf

Public SERVER           As mtSERVER
Public CONFIG           As mtConfig
Public dbcnn            As SQLiteConnection

Public Function Sha1(data As String) As String
    With New cSha1: Sha1 = .Hash(data): End With
End Function
Public Function FileSha1(sFileName As String) As String
    With New cSha1
        FileSha1 = .HashFile(sFileName)
    End With
End Function

Public Function db_exec(Sql As String, data As Variant) As ssSQliteResult
On Error GoTo e
    With dbcnn.Command(Sql)
        .Bindings = data
        .Step
    End With
e:
End Function

Public Sub PushLog(ByVal Text As String, Optional PStyle As PrintStyle, Optional lConsole As CONSOLE_LOGS)
    If FrmMain.GUI = 0 Then Exit Sub
    If (lConsole) And Not (lConsole And CONFIG.Console) Then Exit Sub
    FrmMain.DDE.SendData (800 + PStyle) & Text ', FrmMain.GUI
End Sub
Public Sub CheckDataBase()
    If Val(dbcnn.Pragma("user_version")) <> 1 Then
        If dbcnn.Errcode = SQLITE_NOTADB Then dbcnn.ShowError: Exit Sub
        dbcnn.Execute "CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY AUTOINCREMENT,user TEXT UNIQUE,pwd TEXT,enabled INTEGER DEFAULT (1),max_connections INTEGER);"
        dbcnn.Execute "CREATE TABLE IF NOT EXISTS mounts (id INTEGER PRIMARY KEY AUTOINCREMENT,id_user INTEGER REFERENCES users (id) ON DELETE CASCADE ON UPDATE CASCADE,name TEXT,path TEXT,access INTEGER DEFAULT (0),def INTEGER DEFAULT (0),time_stamp Text);"
        dbcnn.Pragma("user_version") = 1
    End If
End Sub

Public Function Quot(data As String) As String: Quot = Chr(34) & data & Chr(34): End Function

Public Function CheckBound(sElmnt() As String, Optional lMin As Long = 0) As Boolean
On Error GoTo e:
    If UBound(sElmnt) = -1 Then Exit Function
    If UBound(sElmnt) < lMin Then Exit Function
    CheckBound = True
e:
End Function

Property Get FileList(Path As String) As cFileList
    Set FileList = New cFileList
    FileList.List Path
End Property

Public Function WindowsDPI() As Double
Dim Hdc  As Long
Dim lPx  As Double

    Hdc = GetDC(0)
    lPx = CDbl(GetDeviceCaps(Hdc, 88))
    ReleaseDC 0, Hdc
    If (lPx = 0) Then WindowsDPI = 1# Else WindowsDPI = lPx / 96#
    
End Function

Public Sub PutIcon32Bit(ByVal hWnd As Long, ResIcon As Variant)
Dim hIcon As Long
Dim DPI As Double
    
    DPI = DPI = WindowsDPI
    hIcon = LoadImage(App.hInstance, ResIcon, 1, 32 * DPI, 32 * DPI, &H8000& Or &H1000)
    If hIcon Then DestroyIcon SendMessage(hWnd, &H80, 1, ByVal hIcon)
    hIcon = LoadImage(App.hInstance, ResIcon, 1, 16 * DPI, 16 * DPI, &H8000& Or &H1000)
    If hIcon Then DestroyIcon SendMessage(hWnd, &H80, 0, ByVal hIcon)
    
End Sub

Public Function IconFromStream(bvData() As Byte, Optional ByVal W As Long, Optional ByVal H As Long) As Long
On Error GoTo e
    IconFromStream = CreateIconFromResourceEx(bvData(LBound(bvData)), UBound(bvData) + 1&, 1&, &H30000, W&, H&, 0&)
e:
End Function

Public Function ReadIniData(ByVal Section As String, ByVal KeyName As String, Optional Default As String = vbNullString, Optional ByVal IniFile As String) As String
Dim tmp As String
Dim ln  As Long

    tmp = String(256, Chr(0))
    If IniFile = vbNullString Then IniFile = App.Path & "\plugins\data"
    ln = GetPrivateProfileString(Section, KeyName, Default, tmp, 256, IniFile)
    ReadIniData = Left$(tmp, ln)
End Function
Public Sub WriteIniData(ByVal Section As String, ByVal KeyName As String, ByVal Value As String, Optional ByVal IniFile As String)
    If IniFile = vbNullString Then IniFile = App.Path & "\plugins\data"
    Call WritePrivateProfileString(Section, KeyName, Value, IniFile)
End Sub


Public Function ToUnicode(ByVal data As String) As String
On Error GoTo e
    If LenB(data) = 0 Then Exit Function
    
    Dim Out() As Byte
    Dim ln    As Long
    Dim ln2   As String
    Dim tmp   As String
    
    Out = StrConv(data, vbFromUnicode)
    ln = (UBound(Out) + 1) * 2
    tmp = String$(ln, vbNullChar)
    ln2 = MultiByteToWideChar(65001, 0, Out(0), UBound(Out) + 1, StrPtr(tmp), ln)
    If ln2 Then ToUnicode = Left$(tmp, ln2)
e:
End Function
Public Function ToUTF8(ByVal data As String) As String
Dim Out() As Byte
Dim ln As Long
Dim lRet  As Long

    If LenB(data) = 0 Then Exit Function
    ln = Len(data) * 4
    ReDim Out(ln)
    lRet = WideCharToMultiByte(65001, 0, StrPtr(data), Len(data), Out(0), ln + 1, vbNullString, 0)
    If lRet = 0 Then Exit Function
    ReDim Preserve Out(lRet - 1)
    ToUTF8 = StrConv(Out, vbUnicode)
End Function


