VERSION 5.00
Begin VB.Form FrmSetup 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   368
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   346
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   360
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   297
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2520
      Width           =   4455
      Begin VB.TextBox TxtPort 
         Height          =   285
         Left            =   3000
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CheckBox ChkIP 
         Caption         =   "IPv4"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   14
         Top             =   1560
         Width           =   855
      End
      Begin VB.CheckBox ChkIP 
         Caption         =   "IPv6"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   13
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox TxtTimeOut 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TxtStreamSize 
         Height          =   285
         Left            =   3000
         TabIndex        =   6
         Text            =   "32 KiB"
         Top             =   240
         Width           =   1215
      End
      Begin NokilonGui.JImage JImage2 
         Height          =   855
         Left            =   360
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         imgStream       =   "FrmSetup.frx":000C
         eScale          =   1
         Color           =   -1
         Alpha           =   80
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Puerto"
         Height          =   195
         Left            =   2160
         TabIndex        =   16
         Top             =   1200
         Width           =   480
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000016&
         X1              =   0
         X2              =   280
         Y1              =   72
         Y2              =   72
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo maximo de conexion inactiva"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   2610
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Socket stream buffer"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Servidor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   11
      Top             =   2280
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   4695
      Begin VB.CheckBox Chk 
         Caption         =   "Mostrar interfaz en el escritorio"
         BeginProperty DataFormat 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "True"
            FalseValue      =   "False"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   7
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   3015
      End
      Begin VB.CheckBox Chk 
         Caption         =   "Iniciar servidor con windows"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.ComboBox cbIPv 
      Height          =   315
      ItemData        =   "FrmSetup.frx":9B7B
      Left            =   5760
      List            =   "FrmSetup.frx":9B85
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin NokilonGui.JGrid lv2 
      Height          =   615
      Left            =   5760
      TabIndex        =   2
      Top             =   1080
      Width           =   690
      _extentx        =   8202
      _extenty        =   4260
      headerh         =   28
      linecolor       =   15790320
      gridstyle       =   3
      striped         =   -1
      stripedcolor    =   16645629
      selcolor        =   -2147483635
      itemh           =   0
      bordercolor     =   9471874
      header          =   -1
      fullrow         =   -1
      focusrect       =   0
      forecolor       =   0
      editable        =   -1
      editborder      =   14265726
      editback        =   16777215
      editsize        =   1
      drawempty       =   -1
      headercustomdraw=   0
      alphablend      =   -1
      border          =   1
      backcolor       =   16777215
      font            =   "FrmSetup.frx":9B95
      font2           =   "FrmSetup.frx":9BBD
   End
   Begin NokilonGui.JButton BtnMain 
      Height          =   495
      Index           =   0
      Left            =   3360
      TabIndex        =   0
      Top             =   4920
      Width           =   1575
      _extentx        =   2778
      _extenty        =   873
      skinres         =   ""
      text            =   "Guardar"
      textalign       =   2
      bitmapalign     =   1
      margins         =   "8,8"
      buttontype      =   0
      value           =   0
      edown           =   -1
      ambientimage    =   0
      bitmapresize    =   "16x16"
      bitmapcolor     =   -1
      bitmapspace     =   3
      nobkgnd         =   0
      enabled         =   -1
      backcolor       =   -2147483633
      font            =   "FrmSetup.frx":9BEF
      fore0           =   -1
      fore1           =   -1
      fore2           =   -1
      fore3           =   -1
      fore4           =   -1
      image           =   "FrmSetup.frx":9C17
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000016&
      X1              =   16
      X2              =   328
      Y1              =   48
      Y2              =   48
   End
   Begin NokilonGui.JImage JImage1 
      Height          =   480
      Left            =   240
      Top             =   120
      Width           =   480
      _extentx        =   741
      _extenty        =   741
      imgstream       =   "FrmSetup.frx":A554
      escale          =   1
      color           =   -1
      alpha           =   100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Configuración del servidor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   270
      Width           =   2205
   End
End
Attribute VB_Name = "FrmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function SendMessageAsLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

Private Const ECM_FIRST As Long = &H1500
Private Const EM_SETCUEBANNER As Long = (ECM_FIRST + 1)
Private Const EM_SETMARGINS As Long = &HD3
Private Const EC_RIGHTMARGIN As Long = &H2

Private Enum SOCKET_FLAGS
  SS_IPv4 = 1     '/* Server listen in IPv4   */
  SS_IPv6 = 2     '/* Server listen in IPv6   */
End Enum

Private Type tSetup
  lPort       As Long
  StreamSize  As Long
  TimeOut     As Long
  Slipv4      As Boolean
  Slipv6      As Boolean
  Sockets    As SOCKET_FLAGS
End Type

Private mSetup      As tSetup
Private c_mnu1      As cMenuApi

Private Sub Form_Load()
Dim mConfig As cConfig
Dim i       As Long

    Set mConfig = New cConfig
    mConfig.ReadAll dbcnn
    
    With mSetup
        .lPort = mConfig.GetValue2("Port", 21)
        .StreamSize = mConfig.GetValue2("Stream-size", 16)
        .TimeOut = mConfig.GetValue2("Time-out", 0)
        .Slipv4 = mConfig.GetValue3("Slipv4", True)
        .Slipv6 = mConfig.GetValue3("Slipv6", False)
        .Sockets = mConfig.GetValue2("Sockets", 3)
        
        If .lPort = 0 Then .lPort = 21
        If .StreamSize = 0 Then .StreamSize = 8
    End With
    
    
    '/* Server Ports */
    '-------------------------------------------------------------------------------------
    'With lv2
    '    .CreateImageListEx iml.Stream(2), 16 * mdpi_
    '    .AddColumn "Direccion", 110, , True, True
    '    .AddColumn "IP", 80, , True
    '    .AddColumn "Puerto", 70, vbRightJustify, True
    '    .AlignmentItemIcons(2) = vbRightJustify
    '    .itemHeight = 22
    '    .SetControlToGrid cbIPv
    '    With dbcnn.Query("SELECT * FROM listeners;")
    '        Do While .Step = SQLITE_ROW
    '            i = lv2.AddItem(IIf(.Value(1) = 6, "::", "0.0.0.0"), 0, .Value(0), False)
    '            lv2.ItemText(i, 1) = IIf(.Value(1) = 6, "IPv6", "IPv4")
    '            lv2.SetItem i, 2, .Value(2), 1
    '        Loop
    '    End With
    'End With
    '
    'Set c_mnu1 = New cMenuApi
    'c_mnu1.AddItem 100, "Añadir"
    'c_mnu1.AddItem 0, , True
    'c_mnu1.AddItem 101, "Eliminar"
    '-------------------------------------------------------------------------------------
    
    If PathExist(SystemFolder(uStartUp) & "\nokilon-server.lnk") Then Chk(0).Value = 1
    If PathExist(SystemFolder(uDesktop) & "\Nokilon Server Interface.lnk") Then Chk(1).Value = 1
    
    TxtStreamSize = mSetup.StreamSize & " KiB"
    TxtTimeOut = mSetup.TimeOut & " Min"
    TxtPort = mSetup.lPort
    
    'If mSetup.Slipv4 Then ChkIP(0).Value = 1
    'If mSetup.Slipv6 Then ChkIP(1).Value = 1
    
    If mSetup.Sockets And SS_IPv4 Then ChkIP(0).Value = 1
    If mSetup.Sockets And SS_IPv6 Then ChkIP(1).Value = 1
    
    
    'Dim tmp As String
    'tmp = StrConv("Desactivado", vbUnicode)
    'SendMessage TxtTimeOut.hwnd, EM_SETCUEBANNER, 0&, ByVal tmp
    
    'Da un margen derecho sobre el TextBox para que el icono de buscar no tape lo que se escrive
    'RightMargin = 20 * &H10000
    'SendMessage Text1.hwnd, EM_SETMARGINS, EC_RIGHTMARGIN, ByVal RightMargin
    
End Sub
Private Sub btnMain_Click(Index As Integer)
Dim tmp As String

    Select Case Index
    
        Case 0
        
            If ChkIP(0).Value = 0 And ChkIP(1).Value = 0 Then
                'MsgBox "Debe activar una o dos de las versiones IP para el servidor", vbInformation: Exit Sub
            End If
        
            '[?] StarUp
            tmp = SystemFolder(uStartUp)
            If Chk(0).Value = 1 Then
                If Not PathExist(tmp & "\nokilon-server.lnk") Then CreateShortCut App.Path & "\nokilon-server.exe", tmp, "nokilon-server", Args:="-startup"
            Else
                If PathExist(tmp & "\nokilon-server.lnk") Then Kill tmp & "\nokilon-server.lnk"
            End If
            
            '[?] Desktop
            tmp = SystemFolder(uDesktop)
            If Chk(1).Value = 1 Then
                If Not PathExist(tmp & "\Nokilon Server Interface.lnk") Then CreateShortCut App.Path & "\nokilon-server-gui.exe", tmp, "Nokilon Server Interface", App.Path
            Else
                If PathExist(tmp & "\Nokilon Server Interface.lnk") Then Kill tmp & "\Nokilon Server Interface.lnk"
            End If
            
            '[?] Socket Stream Size
            If CInt(Val(TxtStreamSize)) <> mSetup.StreamSize Then
                mSetup.StreamSize = CInt(Val(TxtStreamSize))
                SaveSettingDb "Stream-size", mSetup.StreamSize
                If FrmMain.DDE.Main Then FrmMain.DDE.SendData "052" & mSetup.StreamSize
            End If

            '[?] TimeOut
            If CInt(Val(TxtTimeOut)) <> mSetup.TimeOut Then
                mSetup.TimeOut = CInt(Val(TxtTimeOut))
                If mSetup.TimeOut < 0 Then mSetup.TimeOut = 0
                SaveSettingDb "Time-out", mSetup.TimeOut
                If FrmMain.DDE.Main Then FrmMain.DDE.SendData "053" & mSetup.TimeOut
            End If

            '[?] Server Port
            If CInt(Val(TxtPort)) <> mSetup.lPort Then
                mSetup.lPort = CInt(Val(TxtPort))
                SaveSettingDb "Port", mSetup.lPort
                If FrmMain.DDE.Main Then FrmMain.DDE.SendData "051" & mSetup.lPort
            End If
            
            Dim SS_FLAG As SOCKET_FLAGS
            If ChkIP(0).Value = 1 Then SS_FLAG = SS_IPv4
            If ChkIP(1).Value = 1 Then SS_FLAG = SS_FLAG Or SS_IPv6
            
            If SS_FLAG <> mSetup.Sockets Then
                mSetup.Sockets = SS_FLAG
                SaveSettingDb "Sockets", mSetup.Sockets
                If FrmMain.DDE.Main Then FrmMain.DDE.SendData "057" & mSetup.Sockets
            End If
            Unload Me
            
    End Select

    
End Sub
Private Sub txtPort_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 13 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub
Private Sub txtPort_Validate(Cancel As Boolean)
    If Val(TxtPort) < 1 Then TxtPort = mSetup.lPort
End Sub
Private Sub TxtStreamSize_Validate(Cancel As Boolean)
Dim ln As Long

    ln = Abs(Val(TxtStreamSize))
    If ln < 1 Then ln = 1
    TxtStreamSize = ln & " KiB"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set c_mnu1 = Nothing
End Sub



Private Function mvQueryServerStatus() As Boolean
    If Val(FrmMain.DDE.SendData("004")) Then mvQueryServerStatus = True
End Function



'TODO: Manage Server Ports
'----------------------------------------------------------------------------------------------------------------------
'Private Sub lv2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button <> 2 Then Exit Sub
'    c_mnu1.ItemDisabled(2) = (lv2.SelectedItem = -1)
'    Select Case c_mnu1.PopupMenu
'        Case 100
'            Dim i As Long
'            i = lv2.AddItem("0.0.0.0", 0, 0, True)
'            lv2.ItemText(i, 1) = "IPv4"
'            lv2.SetItem i, 2, CLng(Timer), 1
'        Case 101:
'
'            Dim lIn As Long
'            If lv2.SelectedItem = -1 Then Exit Sub
'            lIn = lv2.SelectedItemData
'            If lIn Then
'                If dbcnn.Execute("DELETE FROM listeners WHERE id=" & lIn & ";") <> SQLITE_OK Then dbcnn.ShowError: Exit Sub
'                FrmMain.DDE.SendData "303" & lIn  'remove listener
'            End If
'            lv2.RemoveItem lv2.SelectedItem
'    End Select
'End Sub
'Private Sub lv2_EditStart(ByVal Item As Long, ByVal Column As Long, X As Long, Y As Long, W As Long, H As Long, Text As String, ObjEdit As Control, Cancel As Boolean, MoveObj As Boolean)
'    Select Case Column
'        Case 1: Set ObjEdit = cbIPv
'    End Select
'End Sub
'Private Sub lv2_EditShow(ByVal Item As Long, ByVal Column As Long, ObjEdit As Control, Visible As Boolean)
'    If Column = 1 Then SendMessageAsLong cbIPv.hwnd, &H14F, 1&, 0&
'End Sub
'Private Sub lv2_EditEnd(ByVal Item As Long, ByVal Column As Long, NewText As String, ObjEdit As Control, Cancel As Boolean)
'    Select Case Column
'        Case 1
'            If lv2.ItemText(Item, 1) = NewText Then Exit Sub
'            lv2.ItemText(Item, 0) = IIf(NewText = "IPv6", "::", "0.0.0.0")
'        Case 2:
'
'            If lv2.ItemText(Item, 2) = NewText Then Exit Sub
'            If Not IsNumeric(NewText) Then GoTo e
'            If val(NewText) < 1 Then GoTo e
'            NewText = CInt(val(NewText))
'
'    End Select
'    lv2.ItemTag(Item) = True
'    Exit Sub
'e:
'    Cancel = True
'End Sub
'Private Sub SaveNotifyServerPorts()
'Dim i As Long
'Dim tmp As String
'Dim lIn As Long
'
'    For i = 0 To lv2.ItemCount - 1
'        If lv2.ItemTag(i) = True Then
'
'            lIn = IIf(lv2.ItemText(i, 1) = "IPv6", 6, 4)
'            '/* {id,version,port) */
'            If lv2.ItemData(i) = 0 Then
'                db_exec "INSERT INTO listeners VALUES (?,?,?,?);", Array(Null, lIn, lv2.ItemText(i, 2), 1)
'                tmp = AddNulls(dbcnn.LastInsertID, lIn, lv2.ItemText(i, 2))
'                FrmMain.DDE.SendData "301" & tmp    'add listener
'            Else
'                tmp = AddNulls(lv2.ItemData(i), lIn, lv2.ItemText(i, 2))
'                FrmMain.DDE.SendData "302" & tmp    'edit listener
'                db_exec "UPDATE listeners SET version=?,port=? WHERE id=?;", Array(lIn, lv2.ItemText(i, 2), lv2.ItemData(i))
'            End If
'
'        End If
'    Next
'End Sub
