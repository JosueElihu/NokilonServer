VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'   Component    : cConfig v1.0
'   Autor        : J. Elihu
'   Requirements : SQLiteConnection + SQliteQuery
'--------------------------------------------------------------------------------

Option Explicit


Private m_Table     As String
Private m_Items     As Collection   '{key,value}

Private Sub Class_Initialize()
    m_Table = "config"
    Set m_Items = New Collection
End Sub
Private Sub Class_Terminate()
    Set m_Items = Nothing
End Sub
Property Get TableName() As String: TableName = m_Table: End Property
Property Let TableName(ByVal Value As String): m_Table = Value: End Property

Public Function ReadAll(dbc As SQLiteConnection) As Boolean
On Error GoTo e_
    With dbc.Query("SELECT * FROM " & m_Table & ";")
        Do While .Step = SQLITE_ROW
            mvAddItem .Value(0), .Value(0), .Value(1)
        Loop
    End With
e_:
End Function

Public Function SetValue(ByVal Key As String, ByVal Value As String)
    If KeyExist(Key) Then m_Items.Remove Key
    mvAddItem Key, Key, Value   '{key,value}
End Function
Public Function GetValue(ByVal Key As String, Optional Default As String = vbNullString) As String
On Error GoTo e_
    GetValue = m_Items(Key)(1) '{key,value}
    Exit Function
e_:
    GetValue = Default
End Function

'/* Long data */
Public Function GetValue2(ByVal Key As String, Optional Default As Long = 0) As Long
On Error GoTo e_
    GetValue2 = Abs(Val(m_Items(Key)(1))) '{key,value}
    Exit Function
e_:
    GetValue2 = Default
End Function

'/* Boolean data */
Public Function GetValue3(ByVal Key As String, Optional Default As Boolean) As Boolean
On Error GoTo e_
    GetValue3 = CBool(m_Items(Key)(1)) '{key,value}
    Exit Function
e_:
    GetValue3 = Default
End Function


Public Function WriteValue(ByVal Key As String, ByVal Value As String, dbc As SQLiteConnection)
    dbc.Execute "CREATE TABLE IF NOT EXISTS " & m_Table & " (key TEXT PRIMARY KEY,value);"
    dbc.Execute "INSERT INTO " & TableName & "(key,value) VALUES('" & Key & "','" & Value & "') " & _
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value;"
End Function
Public Function ReadValue(ByVal Key As String, dbc As SQLiteConnection, Optional Default As String = vbNullString) As String
On Error GoTo e_
    With dbc.Query("SELECT * FROM " & m_Table & "WHERE key='" & Key & "';")
        If .Step <> SQLITE_ROW Then GoTo e_
        ReadValue = .Value(1)
    End With
    Exit Function
e_:
    ReadValue = Default
End Function

'Public Sub RemoveValue(ByVal Key As String)
'On Error GoTo e
    'm_Items.Remove Key
'e_:
'End Sub



Public Function SaveAll(dbc As SQLiteConnection)
Dim i   As Long
Dim tmp As String

    dbc.Execute "CREATE TABLE IF NOT EXISTS " & m_Table & " (key TEXT PRIMARY KEY,value);"
    
'    For i = 1 To m_Items.Count
'        tmp = "INSERT INTO " & m_Table & "(key,value) VALUES('" & m_Items(i)(0) & "','" & m_Items(i)(1) & "') " & _
'              "ON CONFLICT(key) DO UPDATE SET value=excluded.value;"
'        dbc.Execute tmp
'    Next

    With dbc.Command("INSERT INTO " & m_Table & " (key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value;")
        For i = 1 To m_Items.Count
            .Bind 1, m_Items(i)(0)
            .Bind 2, m_Items(i)(1)
            .Step
        Next
    End With
    

End Function





Private Function KeyExist(ByVal Key As String) As Boolean
On Error GoTo e_
    KeyExist = m_Items(Key)(0) <> ""
e_:
End Function
Private Sub mvAddItem(ByVal Key As String, ParamArray elements() As Variant)
On Error GoTo e_
    m_Items.Add elements, Key
e_:
End Sub
