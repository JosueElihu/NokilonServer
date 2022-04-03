VERSION 5.00
Begin VB.Form FrmUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   540
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   667
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicTitle 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   120
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   649
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   840
      Width           =   9735
      Begin NokilonGui.JButton BtnMain 
         Height          =   345
         Index           =   0
         Left            =   2640
         TabIndex        =   12
         ToolTipText     =   "Añadir usuario"
         Top             =   45
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         SkinRes         =   ""
         Text            =   "Añadir"
         TextAlign       =   0
         BitmapAlign     =   1
         Margins         =   "8,8"
         ButtonType      =   0
         Value           =   0   'False
         EDown           =   -1  'True
         AmbientImage    =   0   'False
         BitmapResize    =   "16x16"
         BitmapColor     =   -1
         BitmapSpace     =   3
         NoBkgnd         =   -1  'True
         Fore0           =   -1
         Fore1           =   -1
         Fore2           =   -1
         Fore3           =   -1
         Fore4           =   -1
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Image           =   "FrmUsers.frx":000C
      End
      Begin NokilonGui.JButton BtnMain 
         Height          =   345
         Index           =   1
         Left            =   3750
         TabIndex        =   13
         ToolTipText     =   "Eliminar usuario"
         Top             =   45
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         SkinRes         =   ""
         Text            =   "Eliminar"
         TextAlign       =   0
         BitmapAlign     =   1
         Margins         =   "8,8"
         ButtonType      =   0
         Value           =   0   'False
         EDown           =   -1  'True
         AmbientImage    =   0   'False
         BitmapResize    =   "16x16"
         BitmapColor     =   -1
         BitmapSpace     =   3
         NoBkgnd         =   -1  'True
         Fore0           =   -1
         Fore1           =   -1
         Fore2           =   -1
         Fore3           =   -1
         Fore4           =   -1
         Enabled         =   0   'False
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Image           =   "FrmUsers.frx":0CD0
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuarios"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   765
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   170
         X2              =   170
         Y1              =   4
         Y2              =   24
      End
   End
   Begin VB.PictureBox PicTitle 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   1
      Left            =   120
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   649
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4320
      Width           =   9735
      Begin NokilonGui.JButton BtnMain 
         Height          =   345
         Index           =   2
         Left            =   2640
         TabIndex        =   8
         ToolTipText     =   "Añadir directorio"
         Top             =   45
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         SkinRes         =   ""
         Text            =   "Añadir"
         TextAlign       =   0
         BitmapAlign     =   1
         Margins         =   "8,8"
         ButtonType      =   0
         Value           =   0   'False
         EDown           =   -1  'True
         AmbientImage    =   0   'False
         BitmapResize    =   "16x16"
         BitmapColor     =   -1
         BitmapSpace     =   3
         NoBkgnd         =   -1  'True
         Fore0           =   -1
         Fore1           =   -1
         Fore2           =   -1
         Fore3           =   -1
         Fore4           =   -1
         Enabled         =   0   'False
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Image           =   "FrmUsers.frx":1769
      End
      Begin NokilonGui.JButton BtnMain 
         Height          =   345
         Index           =   3
         Left            =   3750
         TabIndex        =   10
         ToolTipText     =   "Eliminar directorio"
         Top             =   45
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         SkinRes         =   ""
         Text            =   "Eliminar"
         TextAlign       =   0
         BitmapAlign     =   1
         Margins         =   "8,8"
         ButtonType      =   0
         Value           =   0   'False
         EDown           =   -1  'True
         AmbientImage    =   0   'False
         BitmapResize    =   "16x16"
         BitmapColor     =   -1
         BitmapSpace     =   3
         NoBkgnd         =   -1  'True
         Fore0           =   -1
         Fore1           =   -1
         Fore2           =   -1
         Fore3           =   -1
         Fore4           =   -1
         Enabled         =   0   'False
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Image           =   "FrmUsers.frx":242D
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   170
         X2              =   170
         Y1              =   4
         Y2              =   24
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Directorios"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   667
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   735
      Width           =   10005
   End
   Begin NokilonGui.JGrid lv 
      Height          =   2880
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   5080
      HeaderH         =   28
      LineColor       =   15790320
      GridStyle       =   3
      Striped         =   -1  'True
      StripedColor    =   16645629
      SelColor        =   -2147483635
      ItemH           =   0
      BorderColor     =   9471874
      Header          =   -1  'True
      FullRow         =   -1  'True
      FocusRect       =   0   'False
      ForeColor       =   0
      Editable        =   -1  'True
      EditBorder      =   14265726
      EditBack        =   16777215
      EditSize        =   1
      DrawEmpty       =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Border          =   1
      BackColor       =   16777215
      BeginProperty Font2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderCustomDraw=   0   'False
      AlphaBlend      =   -1  'True
   End
   Begin VB.ComboBox cbPermissions 
      Height          =   315
      ItemData        =   "FrmUsers.frx":2EC6
      Left            =   6840
      List            =   "FrmUsers.frx":2ED3
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   7320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin NokilonGui.JImageListEx iml 
      Left            =   9240
      Top             =   120
      _ExtentX        =   794
      _ExtentY        =   688
      Count           =   8
      Data_0          =   "FrmUsers.frx":2EFA
      Data_1          =   "FrmUsers.frx":6BA3
      Data_2          =   "FrmUsers.frx":B154
      Data_3          =   "FrmUsers.frx":B79C
      Data_4          =   "FrmUsers.frx":BCA1
      Data_5          =   "FrmUsers.frx":ECF9
      Data_6          =   "FrmUsers.frx":FA06
      Data_7          =   "FrmUsers.frx":1049F
   End
   Begin NokilonGui.JGrid lv2 
      Height          =   3120
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   5503
      HeaderH         =   28
      LineColor       =   15790320
      GridStyle       =   3
      Striped         =   -1  'True
      StripedColor    =   16645629
      SelColor        =   -2147483635
      ItemH           =   0
      BorderColor     =   9471874
      Header          =   -1  'True
      FullRow         =   -1  'True
      FocusRect       =   0   'False
      ForeColor       =   0
      Editable        =   -1  'True
      EditBorder      =   14265726
      EditBack        =   16777215
      EditSize        =   1
      DrawEmpty       =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Border          =   1
      BackColor       =   16777215
      BeginProperty Font2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderCustomDraw=   0   'False
      AlphaBlend      =   -1  'True
   End
   Begin VB.PictureBox PicTop 
      Align           =   1  'Align Top
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   667
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   10005
      Begin NokilonGui.JImage img0 
         Height          =   240
         Left            =   6960
         Top             =   240
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   423
         imgStream       =   "FrmUsers.frx":10C16
         eScale          =   1
         Color           =   -1
         Alpha           =   100
      End
      Begin NokilonGui.JImage img1 
         Height          =   240
         Left            =   6600
         Top             =   240
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   423
         imgStream       =   "FrmUsers.frx":10C2E
         eScale          =   1
         Color           =   -1
         Alpha           =   100
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Configuracion de usuarios, directorios y permisos"
         ForeColor       =   &H80000015&
         Height          =   195
         Left            =   720
         TabIndex        =   5
         Top             =   390
         Width           =   3525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Administración de Usuarios"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   4
         Top             =   150
         Width           =   2625
      End
      Begin NokilonGui.JImage JImage1 
         Height          =   510
         Left            =   150
         Top             =   120
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   900
         imgStream       =   "FrmUsers.frx":10C46
         eScale          =   1
         Color           =   -1
         Alpha           =   100
      End
   End
   Begin NokilonGui.JImage img3 
      Height          =   300
      Left            =   120
      Top             =   3600
      Visible         =   0   'False
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      imgStream       =   "FrmUsers.frx":1798C
      eScale          =   1
      Color           =   -1
      Alpha           =   100
   End
End
Attribute VB_Name = "FrmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal Hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32" (ByRef lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal Hdc As Long, ByVal crColor As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Long) As Long
Private Declare Function SendMessageAsLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function PathIsRoot Lib "shlwapi" Alias "PathIsRootA" (ByVal pszPath As String) As Long

Private Const CB_SETITEMHEIGHT As Long = &H153
Private Const CB_SHOWDROPDOWN  As Long = &H14F

Private c_mnu1  As cMenuApi
Private c_mnu2  As cMenuApi
Private mvObj   As cGDIPDraw

Private Sub Form_Load()
    
    SetIconStream Me.hWnd, iml.Stream(4, 16, 16)
    With lv
        .AddColumn "Usuario", 167
        .AddColumn "Contraseña", 280
        .AddColumn "Activo", 80, vbCenter, True, True
        .CreateImageListEx iml.Stream(0), 16
        .AlignmentItemIcons(3) = vbRightJustify
        .ItemHeight = 22
    End With
    
    With lv2
        .AddColumn "Nombre", 167 ', , True
        .AddColumn "Permisos", 90, , True
        .AddColumn "Rutal local", 330, , , True
        .CreateImageListEx iml.Stream(1), 16
        .ItemHeight = 22
        .SetControlToGrid cbPermissions
    End With
    
    Set mvObj = New cGDIPDraw
    
    Set c_mnu1 = New cMenuApi
    c_mnu1.AddItem 100, "Nuevo usuario", , , iml.hBitmap(5, 16, 16)
    c_mnu1.AddItem 0, "", True
    c_mnu1.AddItem 101, "Eliminar usuario", , , iml.hBitmap(6, 16, 16)
    c_mnu1.ItemDefault(0) = True
    
    Set c_mnu2 = New cMenuApi
    c_mnu2.AddItem 100, "Añadir carpeta", , , iml.hBitmap(5, 16, 16)
    c_mnu2.AddItem 0, "", True
    c_mnu2.AddItem 101, "Eliminar carpeta", , , iml.hBitmap(6, 16, 16)
    c_mnu2.ItemDefault(0) = True

    img0.LoadPictureFromStream iml.Stream(2, 16, 16, &H404040, 90)
    img1.LoadPictureFromStream iml.Stream(3, 16, 16, , 90)
    iml.Clear
    mvLoadUsers
    
    'FillGradient PicTitle(1).Hdc, 0, 0, PicTitle(1).ScaleWidth, PicTitle(1).ScaleHeight, PicTitle(1).BackColor, &H80000016, True
    DrawLine PicTitle(0).Hdc, &H80000005, 0, 1 * mdpi_, PicTitle(0).ScaleWidth, 1 * mdpi_
    DrawRectBorder PicTitle(0).Hdc, &H80000010, 0, 0, PicTitle(0).ScaleWidth, PicTitle(0).ScaleHeight
    
    'FillGradient PicTitle(1).Hdc, 0, 0, PicTitle(1).ScaleWidth, PicTitle(1).ScaleHeight, PicTitle(1).BackColor, &H80000016, True
    DrawLine PicTitle(1).Hdc, &H80000005, 0, 1 * mdpi_, PicTitle(1).ScaleWidth, 1 * mdpi_
    DrawRectBorder PicTitle(1).Hdc, &H80000010, 0, 0, PicTitle(1).ScaleWidth, PicTitle(1).ScaleHeight
    
End Sub
Private Sub btnMain_Click(Index As Integer)
Dim i As Long

    Select Case Index
        Case 0
        
            Dim tmp As String
            tmp = GetSafeName(lv, "user")
            If dbcnn.Execute("INSERT INTO users (user,pwd) VALUES ('" & tmp & "','');") = SQLITE_OK Then
                i = dbcnn.LastInsertID
                i = lv.AddItem(tmp, 0, i)
                lv.SetItem i, 1, "", 2
                lv.ItemText(i, 2) = 1
                lv.SelectedItem = i
                lv.EditStart i, 0
            Else
                dbcnn.ShowError
            End If
        Case 1
        
            If lv.SelectedItem = -1 Then Exit Sub
            i = lv.SelectedItemData
            If i = 0 Then Exit Sub
            If myCustomMsg("¿Desea eliminar el usuario?" & vbNewLine & "Usuario:  '" & lv.SelectedItemText & "'", "ELIMINAR USUARIO") = vbNo Then Exit Sub
            If dbcnn.Execute("DELETE FROM users WHERE id=" & i & ";") <> SQLITE_OK Then dbcnn.ShowError: Exit Sub
            lv.RemoveItem lv.SelectedItem
            FrmMain.DDE.SendData "105" & i  'Update connections
            lv2.Clear
            
        Case 2:
            
            If lv.SelectedItem = -1 Then Msgbox2 "Ningun usuario seleccionado", "Añadir carpeta", vbCritical: Exit Sub
            i = lv.SelectedItemData
            If i = 0 Then Exit Sub
            
            tmp = mvSelectFolder
            If Len(tmp) = 0 Then Exit Sub
            
            Dim sName As String
            
            If PathIsRoot(tmp) Then sName = "UNIDAD-[" & Left$(tmp, 1) & "]" Else sName = GetFileTitle(tmp)
            sName = GetSafeName(lv2, sName)
            If db_exec("INSERT INTO mounts (id_user,name,path,access,time_stamp) VALUES (?,?,?,?,?);", Array(i, sName, tmp, "Read + Write", Now)) <> SQLITE_OK Then Exit Sub
            
            i = lv2.AddItem(sName, 0, dbcnn.LastInsertID, 0)
            lv2.SetItem i, 2, tmp, 2
            lv2.ItemText(i, 1) = "Read + Write"
            FrmMain.DDE.SendData "106" & lv.SelectedItemData  'Update Mounts
            
            
        Case 3
            If lv2.SelectedItem = -1 Then Exit Sub
            i = lv2.SelectedItemData
            If i = 0 Then Exit Sub
            If myCustomMsg("¿Desea eliminar el directorio?" & vbNewLine & "Nombre:  '" & lv2.SelectedItemText & "'", "ELIMINAR DIRECTORIO") = vbNo Then Exit Sub
            If dbcnn.Execute("DELETE FROM mounts WHERE id=" & i & ";") <> SQLITE_OK Then dbcnn.ShowError: Exit Sub
            lv2.RemoveItem lv2.SelectedItem
            FrmMain.DDE.SendData "106" & lv.SelectedItemData  'Update Mounts
        
        Case 4

            
    End Select
End Sub
Private Sub BtnMain_BeforePaint(Index As Integer, Hdc As Long, hGraphic As Long, ByVal Width As Long, ByVal Height As Long, Evt As JButtonState, Cancel As Boolean)

    mvObj.Graphic = hGraphic
    Select Case Evt
        Case lHotBtn
            mvObj.DrawRectangle 0, 0, Width, Height, &H8000000D, 15, 1, &H8000000D, 20
        Case lDownBtn
            mvObj.DrawRectangle 0, 0, Width, Height, &H8000000D, 25, 1, &H8000000D, 30
        Case Else
            'FillGradient Hdc, 0, 0, Width, Height, PicTitle(1).BackColor, &H80000016, True
    End Select
End Sub

Private Sub lv_EditStart(ByVal Item As Long, ByVal Column As Long, x As Long, y As Long, W As Long, H As Long, Text As String, ObjEdit As Control, Cancel As Boolean, MoveObj As Boolean)
    Select Case Column
        Case 0: lv.Edit.PasswordChar = vbNullString
        Case 1: lv.Edit.PasswordChar = "•"
    End Select
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    c_mnu1.ItemDisabled(2) = lv.SelectedItemData = 0
    Select Case c_mnu1.PopupMenu
        Case 100: btnMain_Click 0: lv.EditStart lv.ItemCount - 1, 0
        Case 101: btnMain_Click 1
    End Select
End Sub
Private Sub lv_ItemDblClick(ByVal Item As Long, ByVal Column As Long)
Dim vmp As String
Dim tmp As String
    
    If Column <> 2 Then Exit Sub
    vmp = IIf(lv.ItemText(Item, Column) = "0", "1", "0")
    'If Msgbox2("Usuario: " & lv.SelectedItemText, IIf(vmp = "0", "BLOQUEAR", "DESBLOQUEAR") & " USUARIO", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    tmp = "UPDATE users SET enabled ='" & vmp & "' WHERE id='" & lv.ItemData(Item) & "';"
    If dbcnn.Execute(tmp) <> SQLITE_OK Then dbcnn.ShowError: Exit Sub
    lv.ItemText(Item, Column) = vmp
    If vmp = 0 Then FrmMain.DDE.SendData "105" & lv.ItemData(Item) 'Close if conected
    
End Sub
Private Sub lv_SelectionChanged(ByVal Item As Long, ByVal Column As Long)
    mvLoadFolders lv.ItemData(Item)
    
    btnMain(1).Enabled = (Item <> -1 And lv.ItemCount > 0)
    btnMain(2).Enabled = (Item <> -1 And lv.ItemCount > 0)
    'BtnMain(4).Enabled = (lv2.SelectedItem <> -1 And lv2.ItemCount > 0)
End Sub
Private Sub lv_EditEnd(ByVal Item As Long, ByVal Column As Long, NewText As String, ObjEdit As Control, Cancel As Boolean)
Dim tmp As String

    If lv.ItemData(Item) = 0 Then Exit Sub
    Select Case Column
        Case 0
        
            'If IsUnike(lv, NewText, Item) = False Then MsgBox "El usuario '" & NewText & "' ya existe", vbCritical, "": Cancel = True: Exit Sub
            If lv.ItemText(Item, Column) = NewText Then Exit Sub
            If IsUnike(lv, NewText, Item) = False Then Msgbox2 "El usuario '" & NewText & "' ya existe", "Usuarios", vbCritical: Cancel = True: Exit Sub
            tmp = "UPDATE users SET user='" & NewText & "'" & " WHERE id='" & lv.ItemData(Item) & "';"
            If dbcnn.Execute(tmp) <> SQLITE_OK Then Cancel = True: dbcnn.ShowError: Exit Sub
            FrmMain.DDE.SendData "105" & lv.ItemData(Item)  'Update connections
            
        Case 1:
            
            If lv.ItemText(Item, Column) = NewText Then Cancel = True: Exit Sub
            tmp = Sha1(NewText)
            If lv.ItemText(Item, Column) = tmp Then Cancel = True: Exit Sub
            
            NewText = tmp
            tmp = "UPDATE users SET pwd='" & tmp & "'" & " WHERE id='" & lv.ItemData(Item) & "';"
            If dbcnn.Execute(tmp) <> SQLITE_OK Then Cancel = True: dbcnn.ShowError: Exit Sub
            FrmMain.DDE.SendData "105" & lv.ItemData(Item)  'Update connections
            
    End Select

    
End Sub
Private Sub lv_ItemDraw(ByVal Item As Long, ByVal Column As Long, Hdc As Long, hGraphic As Long, ByVal x As Long, ByVal y As Long, ByVal W As Long, ByVal H As Long, CancelDraw As Boolean)
    'If Column <> 2 Then Exit Sub
    
    Select Case Column
    
        Case 0
        Case 1
            
            If Len(lv.ItemText(Item, Column)) > 0 Then Exit Sub
            Dim Rct As RECT
            SetRect Rct, x + (23 * mdpi_), y, x + W - (23 * mdpi_), y + H
            SetTextColor Hdc, &HC0C0C0
            DrawText Hdc, StrPtr("Sin contraseña"), -1, Rct, &H20 Or &H4 Or &H40000
            SetTextColor Hdc, lv.ForeColor
            
        Case 2
        
            Dim Px As Long
            Px = 16 * lv.DpiScale
            
            If Val(lv.ItemText(Item, 2)) > 0 Then
                img1.Render 0, hGraphic, x + ((W - Px) \ 2), y + ((H - Px) \ 2)
            Else
                img0.Render 0, hGraphic, x + ((W - Px) \ 2), y + ((H - Px) \ 2)
            End If
            CancelDraw = True
            
    End Select
    

End Sub
Private Sub lv_ItemDrawData(ByVal Item As Long, ByVal Column As Long, ForeColor As Long, BackColor As Long, BorderColor As Long, ItemIdent As Long)
    ''
End Sub



Private Sub lv2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If lv.SelectedItem = -1 Then Exit Sub
    
    c_mnu2.ItemDisabled(2) = lv2.SelectedItemData = 0
    Select Case c_mnu2.PopupMenu
        Case 100: btnMain_Click 2
        Case 101: btnMain_Click 3
    End Select
End Sub

Private Sub lv2_SelectionChanged(ByVal Item As Long, ByVal Column As Long)
    btnMain(3).Enabled = (Item <> -1) And lv2.ItemCount > 0
End Sub
Private Sub lv2_EditStart(ByVal Item As Long, ByVal Column As Long, x As Long, y As Long, W As Long, H As Long, Text As String, ObjEdit As Control, Cancel As Boolean, MoveObj As Boolean)
    Select Case Column
        Case 1: Set ObjEdit = cbPermissions
    
    End Select
End Sub
Private Sub lv2_EditShow(ByVal Item As Long, ByVal Column As Long, ObjEdit As Control, Visible As Boolean)
    If Column = 1 Then SendMessageAsLong cbPermissions.hWnd, CB_SHOWDROPDOWN, 1&, 0&
End Sub
Private Sub lv2_EditEnd(ByVal Item As Long, ByVal Column As Long, NewText As String, ObjEdit As Control, Cancel As Boolean)
    Select Case Column
        Case 0:
                If lv2.ItemText(Item, Column) = NewText Then Exit Sub
                If IsUnike(lv2, NewText, Item) = False Then
                    MsgBox "El directorio '" & NewText & "' ya existe" & IIf(lv.SelectedItem <> -1, " para el usuario '" & lv.SelectedItemText & "'", ""), vbCritical, "Anadir directorio"
                    Cancel = True
                    Exit Sub
                End If
                
                Dim lID As Long
                Dim tmp As String
                
                lID = lv2.ItemData(Item)
                If lID = 0 Then Cancel = True: Exit Sub
                
                tmp = "UPDATE mounts SET name='" & NewText & "' WHERE id='" & lID & "';"
                If dbcnn.Execute(tmp) <> SQLITE_OK Then Cancel = True: dbcnn.ShowError: Exit Sub
                If lv.SelectedItemData Then FrmMain.DDE.SendData "106" & lv.SelectedItemData  'Update Mounts
                
        Case 1
        
            lID = lv2.SelectedItemData
            If lID = 0 Then Cancel = True: Exit Sub
            tmp = "UPDATE mounts SET access='" & NewText & "' WHERE id='" & lID & "';"
            If dbcnn.Execute(tmp) <> SQLITE_OK Then Cancel = True: dbcnn.ShowError
            If lv.SelectedItemData Then FrmMain.DDE.SendData "106" & lv.SelectedItemData  'Update Mounts
            
    End Select
    
End Sub
Private Sub lv2_ItemDrawData(ByVal Item As Long, ByVal Column As Long, ForeColor As Long, BackColor As Long, BorderColor As Long, ItemIdent As Long)
    If Column > 0 Then ForeColor = &H414141
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set c_mnu1 = Nothing
    Set c_mnu2 = Nothing
    Set mvObj = Nothing
End Sub

Private Sub mvLoadUsers()
Dim j As Long

    lv.NoDraw = True
    lv.Clear
    With dbcnn.Query("SELECT * FROM users;")
        Do While .Step = SQLITE_ROW
            j = lv.AddItem(.Value(1), 0, .Value(0), .Value(2))
            'lv.SetItem j, 1, String(Len(.Value(2)), PASSWORD_CHAR), 1
            lv.SetItem j, 1, .Value(2), 2
            lv.ItemText(j, 2) = Val(.Value(3))
        Loop
    End With
    lv.NoDraw = False
    
End Sub

Private Sub mvLoadFolders(Id As Long)
Dim j As Long

    lv2.Clear
    If Not Id > 0 Then Exit Sub
    lv2.NoDraw = True
    With dbcnn.Query("SELECT * FROM mounts WHERE id_user=" & Id & ";")
        Do While .Step = SQLITE_ROW
            j = lv2.AddItem(.Value(2), 0, .Value(0), Val(.Value(4)))
            lv2.SetItem j, 2, .Value(3), 2
            lv2.ItemText(j, 1) = .Value(4)
            'Format(500000), "@@@@-@@@@")
            'lv2.ItemText(j, 1) = Format(Hex(56), "000")
        Loop
    End With
    lv2.NoDraw = False
End Sub

Private Function IsUnike(moLv As JGrid, Value As String, Optional lIndex As Long = -1) As Boolean
Dim i As Long
    i = moLv.ItemFind(Value)
    If (lIndex <> -1) And (i = lIndex) Then i = -1
    IsUnike = (i = -1)
End Function
Private Function GetSafeName(moLv As JGrid, sName As String) As String
Dim i As Long
    If moLv.ItemFind(sName) = -1 Then GetSafeName = sName: Exit Function
    i = 1
    Do While moLv.ItemFind(sName & "-" & i) <> -1
        i = i + 1
    Loop
    GetSafeName = sName & "-" & i
End Function
Private Function myCustomMsg(pzp_msg As String, Optional pz_title As String = "USUARIOS") As VbMsgBoxResult
Dim tmp As String
    tmp = pz_title & vbNewLine & String$(40, "-") & vbNewLine & pzp_msg
    myCustomMsg = MsgBox(tmp, vbQuestion + vbYesNo)
End Function

Private Function mvSelectFolder() As String
    With New CmnDialogEx
        .FlagsDialog = DLG__BaseOpenDialogFlags Or DLGex_PickFolders Or DLG_Explorer
        If .Version = dvXP_Win2K Then
            If .ShowBrowseForFolder(Me.hWnd) = False Then Exit Function
        Else
            If .ShowOpen(Me.hWnd) = False Then Exit Function
        End If
        If Len(.FileName) Then mvSelectFolder = .FileName
    End With
End Function



