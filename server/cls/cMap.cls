VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private m_Items     As Collection
Private m_Flag      As Boolean

Private Sub Class_Initialize()
    Set m_Items = New Collection
End Sub
Private Sub Class_Terminate()
    Set m_Items = Nothing
End Sub


Property Get UseObjects() As Boolean: UseObjects = m_Flag: End Property
Property Let UseObjects(ByVal Value As Boolean): m_Flag = Value: End Property

Public Function KeyExist(ByVal Key As String) As Boolean
On Error GoTo e
    KeyExist = CStr(m_Items(Key)) <> ""
e:
End Function

Public Function PutItem(ByVal Key As String, Value As Variant) As Boolean
On Error GoTo e
    If KeyExist(Key) Then RemoveItem Key
    m_Items.Add Value, Key
e:
End Function

Property Get Item(ByVal Key As String) As Variant
On Error GoTo e
    Item = m_Items(Key)
    If m_Flag Then Set Item = m_Items(Key)
e:
End Property

Public Sub RemoveItem(ByVal Key As String)
On Error GoTo e
    m_Items.Remove Key
e:
End Sub
