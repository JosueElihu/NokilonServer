Attribute VB_Name = "mGlobal"
Option Explicit

Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Enum PrintStyle
  enmPSNone
  enmPSSuccess
  enmPSInfo
  enmPSError
  enmPSWarning
  enmPSServer
End Enum
Public Enum FTPAccess
  DISABLED_ = 0
  READ_ONLY = 1
  READ_WRITE = 2
End Enum
Public Enum CONSOLE_LOGS
  mclNONE = 0
  mclSERVER = 1
  mclFTP = 2
  mclERRORS = 4
End Enum

Public dbcnn        As SQLiteConnection
Public LOGS_DATA    As CONSOLE_LOGS
Public ShowProps_   As Boolean

Public Sub WindowOnTop(hWnd As Long, Value As Boolean)
    SetWindowPos hWnd, IIf(Value, -1, -2), 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
End Sub

Public Sub PushLog(ByVal Text As String, Optional PStyle As PrintStyle)
Dim lPos    As Long
Dim lEnd    As Long
Dim lColor  As Long
    
    Select Case PStyle
        Case enmPSNone: lColor = &H333333
        Case enmPSSuccess: lColor = &H158719
        Case enmPSInfo: lColor = &HCC8A35
        Case enmPSError: lColor = &HE2BB1
        Case enmPSWarning: lColor = &H9CDE&    '&H59BCF2
        Case enmPSServer: lColor = &HAE1E7D
        '&H00C38E31&
    End Select

    With FrmMain.RTE
        .HideSelection = True
        If .TextLen > 0 Then .AddText vbNewLine
        .AddText Format(Time, "hh:mm:ss") & " - ", &H6E6E6E
        lPos = InStr(Text, "[")
        lEnd = InStr(Text, "]")
        If lPos = 1 And lEnd <> 0 Then .AddText Mid$(Text, lPos, lEnd), &H404080: Text = Right$(Text, Len(Text) - lEnd)

        'If FontBold Then .SelBold = True
        .AddText Text, lColor
        'If FontBold Then .SelBold = False
        .HideSelection = False
        .SetRange -1, -1
    End With
    
End Sub
Public Sub CheckDataBase()
    If Val(dbcnn.Pragma("user_version")) <> 1 Then
        If dbcnn.errCode = SQLITE_NOTADB Then dbcnn.ShowError: Exit Sub
        dbcnn.Execute "CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY AUTOINCREMENT,user TEXT UNIQUE,pwd TEXT,enabled INTEGER DEFAULT (1),max_connections INTEGER);"
        dbcnn.Execute "CREATE TABLE IF NOT EXISTS mounts (id INTEGER PRIMARY KEY AUTOINCREMENT,id_user INTEGER REFERENCES users (id) ON DELETE CASCADE ON UPDATE CASCADE,name TEXT,path TEXT,access INTEGER DEFAULT (0),def INTEGER DEFAULT (0),time_stamp Text);"
        dbcnn.Pragma("user_version") = 1
    End If
End Sub
Public Function db_exec(SQL As String, data As Variant) As ssSQliteResult
On Error GoTo e_
    With dbcnn.Command(SQL)
        .Bindings = data
        .Step
    End With
e_:
End Function
Public Function SaveSettingDb(Key As String, ByVal Value As String)
    With New cConfig
        .WriteValue Key, Value, dbcnn
    End With
End Function

Public Function CheckArgs(sElmnt() As String, Optional lNums As Long = 0) As Boolean
On Error GoTo e:
    If UBound(sElmnt) = -1 Then Exit Function
    If UBound(sElmnt) < lNums Then Exit Function
    CheckArgs = True
e:
End Function
Public Function FmtSpeed(Speed As String) As String
    If IsNumeric(Speed) Then FmtSpeed = FmtSize(Speed) & "/Seg" Else FmtSpeed = Speed
End Function
Public Function ParseLong(ByVal Value As String) As Long
On Error GoTo e
    ParseLong = Abs(Val(Value))
e:
End Function
Public Function Msgbox2(Text As String, Title As String, mbButtons As VbMsgBoxStyle) As VbMsgBoxResult
Dim n As Long
    n = Len(Title)
    n = 50
   ' If n < 50 Then n = 50
    Msgbox2 = MsgBox(Title & vbNewLine & String$(n, "-") & vbNewLine & Text, mbButtons, "USUARIOS")
End Function
Public Function Sha1(data As String) As String
    With New cSha1: Sha1 = .Hash(data): End With
End Function
Public Function AddNulls(ParamArray elements() As Variant) As String
On Error GoTo e
Dim i As Long
    
    If UBound(elements) = 0 Then AddNulls = elements(0): Exit Function
    For i = 0 To UBound(elements) - 1
        AddNulls = AddNulls & elements(i) & vbNullChar
    Next
    If UBound(elements) > 0 Then AddNulls = AddNulls & elements(UBound(elements))
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


Public Function SaveSetting(ByVal Section As String, ByVal KeyName As String, ByVal Value As String)
    VBA.SaveSetting "Nokilon-Server", Section, KeyName, Value
End Function
Public Function GetSetting(ByVal Section As String, ByVal KeyName As String, Optional Default As String = vbNullString)
    GetSetting = VBA.GetSetting("Nokilon-Server", Section, KeyName, Default)
End Function
Public Sub DeleteAllSettings()
On Error GoTo e
    VBA.DeleteSetting "Nokilon-Server"
e:
End Sub


Public Function FmtSize(ByVal Bytes As Currency) As String
    If Bytes >= 1024@ Then
        If Bytes >= 1073741824 Then
            FmtSize = Format((Bytes / 1073741824), "##,###,##0.00") & " GB"
        Else
            If Bytes >= 1048576 Then
                FmtSize = Format((Bytes / 1048576), "##,###,##0.00") & " MB"
            Else
                FmtSize = Format((Bytes \ 1024), "##,###,##0") & " KB"
            End If
        End If
    Else
        FmtSize = CStr(Format(Bytes, "##,###,##0") & " Bytes")
    End If
    FmtSize = CStr(FmtSize)
End Function

Public Function BytesMS(ByVal Filesize As Variant) As String
    Dim Size As Variant
    Select Case VarType(Filesize)
    Case vbByte, vbInteger, vbLong, vbCurrency, vbSingle, vbDouble, vbDecimal
        Filesize = CDec(Filesize)
        If Filesize >= 1024 Then
            For Each Size In Array(" kB", " MB", " GB", " TB", " PB", " EB", " ZB", " YB")
                Select Case Filesize
                Case Is < 10240
                    BytesMS = Format$(Filesize / 1024, "0.00") & Size
                    Exit For
                Case Is < 102400
                    BytesMS = Format$(Filesize / 1024, "0.0") & Size
                    Exit For
                Case Is < 1024000
                    BytesMS = Format$(Filesize / 1024, "0") & Size
                    Exit For
                Case Else
                    Filesize = Filesize / 1024
                End Select
            Next
        Else
            BytesMS = Filesize & " bytes"
        End If
    End Select
End Function
