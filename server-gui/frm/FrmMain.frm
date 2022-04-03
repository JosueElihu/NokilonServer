VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "NokilonServer - Interface"
   ClientHeight    =   8730
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13830
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
   ScaleHeight     =   582
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   922
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicTitle 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   120
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   649
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4440
      Width           =   9735
      Begin NokilonGui.JButton BtnMain 
         Height          =   345
         Index           =   0
         Left            =   2640
         TabIndex        =   4
         ToolTipText     =   "Cancelar Transferencia"
         Top             =   45
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         SkinRes         =   ""
         Text            =   "Cancelar"
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
         Image           =   "FrmMain.frx":5B4B2
      End
      Begin NokilonGui.JButton BtnMain 
         Height          =   345
         Index           =   1
         Left            =   3750
         TabIndex        =   5
         ToolTipText     =   "Actualizar lista"
         Top             =   45
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         SkinRes         =   ""
         Text            =   "Actualizar"
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
         Image           =   "FrmMain.frx":5D985
      End
      Begin NokilonGui.JButton BtnMain 
         Height          =   345
         Index           =   2
         Left            =   5160
         TabIndex        =   9
         ToolTipText     =   "Limpiar lista"
         Top             =   45
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         SkinRes         =   ""
         Text            =   "Limpiar"
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
         Image           =   "FrmMain.frx":5F066
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   336
         X2              =   336
         Y1              =   4
         Y2              =   24
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
         Caption         =   "Transferencias"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   765
      End
   End
   Begin NokilonGui.ucStatusbar Sb 
      Height          =   375
      Left            =   4680
      Top             =   7200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
   End
   Begin NokilonGui.JGrid lv2 
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   3201
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
      Editable        =   0   'False
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
   Begin VB.PictureBox PicSpliter 
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   60
      Left            =   120
      MousePointer    =   7  'Size N S
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   641
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4200
      Width           =   9615
   End
   Begin NokilonGui.JImageListEx Iml 
      Left            =   3120
      Top             =   120
      _ExtentX        =   794
      _ExtentY        =   688
      Count           =   11
      Data_0          =   "FrmMain.frx":60C78
      Data_1          =   "FrmMain.frx":64CE8
      Data_2          =   "FrmMain.frx":6A0EC
      Data_3          =   "FrmMain.frx":70BF4
      Data_4          =   "FrmMain.frx":72F92
      Data_5          =   "FrmMain.frx":746AC
      Data_6          =   "FrmMain.frx":760A2
      Data_7          =   "FrmMain.frx":78F2D
      Data_8          =   "FrmMain.frx":7BC50
      Data_9          =   "FrmMain.frx":7D472
      Data_10         =   "FrmMain.frx":7E943
   End
   Begin NokilonGui.TabStrip Ts 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5741
      Begin NokilonGui.JGrid lv 
         Height          =   1335
         Left            =   120
         TabIndex        =   7
         Top             =   420
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   2355
         HeaderH         =   28
         LineColor       =   15790320
         GridStyle       =   3
         Striped         =   -1  'True
         StripedColor    =   16645629
         SelColor        =   -2147483635
         ItemH           =   0
         BorderColor     =   -2147483632
         Header          =   -1  'True
         FullRow         =   -1  'True
         FocusRect       =   0   'False
         ForeColor       =   0
         Editable        =   0   'False
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
      Begin VB.PictureBox PicLog 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   73
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   420
         Width           =   1095
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "Servidor"
      Index           =   0
      Begin VB.Menu mnuServer 
         Caption         =   "Iniciar servicio"
         Index           =   0
      End
      Begin VB.Menu mnuServer 
         Caption         =   "Detener servicio"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuServer 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuServer 
         Caption         =   "Propiedades"
         Enabled         =   0   'False
         Index           =   3
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "Configuración"
      Index           =   1
      Begin VB.Menu mnuSetup 
         Caption         =   "Administrar usuarios"
         Index           =   0
      End
      Begin VB.Menu mnuSetup 
         Caption         =   "Configurar servidor"
         Index           =   1
      End
      Begin VB.Menu mnuSetup 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuSetup 
         Caption         =   "Mostrar eventos"
         Index           =   3
         Begin VB.Menu mnuLog 
            Caption         =   "Servidor"
            Index           =   0
         End
         Begin VB.Menu mnuLog 
            Caption         =   "Comandos ftp"
            Index           =   1
         End
         Begin VB.Menu mnuLog 
            Caption         =   "Errores"
            Index           =   2
         End
      End
      Begin VB.Menu mnuSetup 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuSetup 
         Caption         =   "Limpiar registro"
         Index           =   5
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "Ayuda"
      Index           =   2
      Begin VB.Menu mnuHelp 
         Caption         =   "Acerca de.."
         Index           =   0
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents c_DDE    As cDDE
Attribute c_DDE.VB_VarHelpID = -1
Private WithEvents c_RTE    As cRichEdit
Attribute c_RTE.VB_VarHelpID = -1
Private WithEvents c_Timer  As cTimer
Attribute c_Timer.VB_VarHelpID = -1

Private lPH1        As Single
Private c_mnu1      As cMenuApi
Private c_mnu2      As cMenuApi
Private c_Draw      As cGDIPDraw
Private c_SysIml    As cSystemIml


Private Sub Form_Load()
Dim mConfig As cConfig

    If App.PrevInstance Then
        mvGuiPrevInstance
        Unload Me
        Exit Sub
    End If

    Set c_DDE = New cDDE
    Set c_RTE = New cRichEdit
    Set c_Timer = New cTimer
    Set c_Draw = New cGDIPDraw
    Set mConfig = New cConfig
    Set dbcnn = SQLite.Connection(App.Path & "\noz-db3")
    
    mdpi_ = WindowsDPI
    lPH1 = (PicSpliter.Top + (2 * mdpi_)) * 100 / Me.ScaleHeight
    mConfig.ReadAll dbcnn
    PutIcon32Bit Me.hWnd, "ALPHA"
    
    With c_RTE
        .Create PicLog.hWnd, , Both, True, False
        .LeftMargin = 5 * mdpi_
    End With
    
    With Ts
        '.CreateImageList 18 * mdpi_, 18 * mdpi_ ', Iml.hBitmap(5, 18  * 2, 18 )
        .AddTab 0, "Conexiones", , 0
        .AddTab 1, "Eventos", , 1
    End With
    
    'Grids
    With lv
        .CreateImageListEx iml.Stream(0), 16
        .AddColumn "WS", 60, vbCenter, True
        .AddColumn "Usuario", 150
        .AddColumn "IP", 150
        .AddColumn "Descargados", 110, vbRightJustify, True
        .AddColumn "Recibidos", 110, vbRightJustify, True

        .AlignmentItemIcons(3) = vbRightJustify
        .AlignmentItemIcons(4) = vbRightJustify
        
        .ItemHeight = 22
        '.Left = 4 * mdpi_
    End With
    With lv2
        .CreateImageListEx iml.Stream(1), 16
        .AddColumn "Usuario", 100
        .AddColumn "Archivo", 230, MinWidth:=80
        .AddColumn "Estado", 150, , , , 80  ', vbCenter ', True
        .AddColumn "Bytes", 80, vbRightJustify  ', True
        .AddColumn "Total", 70, vbRightJustify  ', True
        .AddColumn "Velocidad", 100, vbCenter, True
        .AddColumn "Ruta", 210
        .AddColumn "Hora", 160, vbRightJustify
        '.AddColumn "ID", 50, vbCenter, True
        .AlignmentItemIcons(1) = vbRightJustify
        .ItemHeight = 22
    End With
    
    With Sb
        .Initialize True, True
        .InitializeIconList 16 * mdpi_, 16 * mdpi_, iml.hBitmap(2, 16 * 6, 16)
        .AddPanel , , , sbSpring, "No conectado", , 0
        .AddPanel , 160 * mdpi_, 160 * mdpi_, sbContents, "" ', , 1  ' SERVER STATUS
        .AddPanel , 160 * mdpi_, 160 * mdpi_, sbContents, ""       ' SERVER IP
        .AddPanel , 160 * mdpi_, 160 * mdpi_, sbContents, ""       ' SERVER PORT
    End With
    
    Set c_mnu1 = New cMenuApi
    c_mnu1.AddItem 100, "Desconectar", , , iml.hBitmap(3, 16, 16)
    c_mnu1.AddItem 0, "-", True
    c_mnu1.AddItem 101, "Actualizar", , , iml.hBitmap(4, 16, 16)
    c_mnu1.ItemDefault(0) = True
    
    Set c_mnu2 = New cMenuApi
    c_mnu2.AddItem 100, "Cancelar", , , iml.hBitmap(3, 16, 16)
    c_mnu2.AddItem 0, , True
    c_mnu2.AddItem 101, "Actualizar", , , iml.hBitmap(4, 16, 16)
    c_mnu2.AddItem 102, "Limpiar" ', , , Iml.hBitmap(5, 16 , 16 )
    c_mnu2.AddItem 0, , True
    
    With c_mnu2.AddSubMenu("Submnu", "Limpieza automática")
        .AddItem 103, "Completos"
        .AddItem 104, "Incompletos"
        .ItemCheck(0) = mConfig.GetValue("Rmvct", False)
        .ItemCheck(1) = mConfig.GetValue("Rmvit", False)
    End With
    c_mnu2.ItemDefault(0) = True
    

    ' Server
    '------------------------------------------------
    PutIconToVBMenu Me.hWnd, iml.hBitmap(5, 16, 16), 0, 0
    PutIconToVBMenu Me.hWnd, iml.hBitmap(3, 16, 16), 1, 0
    PutIconToVBMenu Me.hWnd, iml.hBitmap(6, 16, 16), 3, 0
    
    ' Config
    '------------------------------------------------
    PutIconToVBMenu Me.hWnd, iml.hBitmap(7, 16, 16), 0, 1
    PutIconToVBMenu Me.hWnd, iml.hBitmap(8, 16, 16), 1, 1
    PutIconToVBMenu Me.hWnd, iml.hBitmap(9, 16, 16), 3, 1
    
    ' Help
    '------------------------------------------------
    PutIconToVBMenu Me.hWnd, iml.hBitmap(10, 16, 16, &HC89568), 0, 2
    iml.Clear
    
    LOGS_DATA = Abs(Val(mConfig.GetValue("Events", 7)))
    ShowProps_ = GetSetting("Gui", "Server-info", True)
    
    If LOGS_DATA And mclSERVER Then mnuLog(0).Checked = True
    If LOGS_DATA And mclFTP Then mnuLog(1).Checked = True
    If LOGS_DATA And mclERRORS Then mnuLog(2).Checked = True
    Call mvDDEConnect
    If c_DDE.Main = 0 Then CheckDataBase
    mvFormPos True
   
    Set c_SysIml = New cSystemIml
    c_SysIml.CreateSystemImageList emSysIml_sys_small
    
End Sub
Private Sub Form_Resize()
On Error Resume Next
Dim i As Long

    Ts.Width = Me.ScaleWidth - (Ts.Left * 2)
    PicLog.Width = Ts.Width * Screen.TwipsPerPixelX - (PicLog.Left * 2) 'Me.ScaleWidth - (PicLog.Left * 2)
    lv.Width = Ts.Width * Screen.TwipsPerPixelX - (lv.Left * 2) 'Me.ScaleWidth - (lv.Left * 2)
    lv2.Width = Me.ScaleWidth - (lv2.Left * 2)
    
    PicTitle.Width = Me.ScaleWidth - (PicTitle.Left * 2)
    PicSpliter.Width = Me.ScaleWidth - PicSpliter.Left * 2
    'PicSpliter.Top = Me.ScaleHeight * lPH1 / 100
    
    With PicSpliter
        PicSpliter.Top = Me.ScaleHeight * lPH1 / 100
        If .Top + 150 > Me.ScaleHeight Then .Top = Me.ScaleHeight - 150
    End With
    Call mvAdjustControls
    
End Sub

Private Sub PicTitle_Paint()
    PicTitle.Cls
    'FillGradient PicTitle(1).Hdc, 0, 0, PicTitle(1).ScaleWidth, PicTitle(1).ScaleHeight, PicTitle(1).BackColor, &H80000016, True
    DrawLine PicTitle.Hdc, &H80000005, 0, 1 * mdpi_, PicTitle.ScaleWidth, 1 * mdpi_
    DrawRectBorder PicTitle.Hdc, &H80000010, 0, 0, PicTitle.ScaleWidth, PicTitle.ScaleHeight
End Sub
Private Sub PicTitle_Resize()
    PicTitle_Paint
End Sub
Private Sub PicSpliter_Resize()
Dim i As Long

    With PicSpliter

        PicSpliter.Top = Me.ScaleHeight * lPH1 / 100
        If .Top + 150 > Me.ScaleHeight Then .Top = Me.ScaleHeight - 150

        .AutoRedraw = True
        .Cls
        .ForeColor = vb3DHighlight
        For i = PicSpliter.ScaleWidth / 2 - (40 * mdpi_) To PicSpliter.ScaleWidth / 2 + (40 / mdpi_) Step (7 * mdpi_)
            PicSpliter.PSet (i, 2)
        Next
        .ForeColor = &H80000010  'vbInactiveBorder
        For i = PicSpliter.ScaleWidth / 2 - (40 * mdpi_) To PicSpliter.ScaleWidth / 2 + (40 / mdpi_) Step (7 * mdpi_)
            PicSpliter.PSet (i - 1, 1)
        Next
        .Refresh
        .AutoRedraw = False
    End With
End Sub

Private Sub PicSpliter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Static LastY As Single
Dim lPos As Long
    
    If Button = 1 Then
        lPos = PicSpliter.Top + (y - LastY) '* Screen.TwipsPerPixelY
        If lPos > 150 And lPos < Me.ScaleHeight - 150 Then
            PicSpliter.Top = lPos
            lPH1 = (PicSpliter.Top) * 100 / Me.ScaleHeight
            Call mvAdjustControls
        End If
    Else
        LastY = y
    End If
End Sub

Private Sub Ts_TabClick(ByVal lTab As Long)
    PicLog.Visible = lTab = 1
    lv.Visible = lTab = 0
End Sub
Private Sub mnuServer_Click(Index As Integer)
On Error GoTo e

    Select Case Index
        Case 0 '/* Start Service */
        
            If c_DDE.Main Then Exit Sub
            
            If c_DDE.Connect("Nokilon-SERVER-DDE", "Nokilon-GUI") Then
                'If LOGS_DATA And mclSERVER Then PushLog "Service started successfully."
                mvOnServerConnection True
                If ShowProps_ Then FrmServer.Show , Me
            Else
            
                c_DDE.InitDDE "Nokilon-GUI2"
                If PathExist(App.Path & "\nokilon-server.exe") Then
                    Shell2 App.Path & "\nokilon-server.exe"
                    Dim mTick As Single
                    mTick = Timer
                    Do While Timer - mTick < 5
                        DoEvents
                        If c_DDE.Main <> 0 Then
                            'If LOGS_DATA And mclSERVER Then PushLog "Service started successfully."
                            Exit Sub
                        End If
                    Loop
                End If
                If LOGS_DATA And mclSERVER Then PushLog "Couldn't start service.", enmPSError
            End If
            
        Case 1 '/* End Service */
        
            If c_DDE.Main Then
                c_DDE.SendData "002"
            Else
                Shell "cmd.exe /k TASKKILL /IM nokilon-server.exe", vbHide
            End If
            
        Case 3: '/* Porperties */
            FrmServer.Show , Me
    End Select
    
e:
End Sub
Private Sub mnuSetup_Click(Index As Integer)
    Select Case Index
        Case 0: FrmUsers.Show , Me
        Case 1: FrmSetup.Show 1, Me
        Case 5:
                c_RTE.Clear
                c_RTE.SetFocus
    End Select
End Sub
Private Sub mnuLog_Click(Index As Integer)
Dim lConsole As CONSOLE_LOGS

    mnuLog(Index).Checked = Not mnuLog(Index).Checked
    If mnuLog(0).Checked Then lConsole = lConsole Or mclSERVER
    If mnuLog(1).Checked Then lConsole = lConsole Or mclFTP
    If mnuLog(2).Checked Then lConsole = lConsole Or mclERRORS
    LOGS_DATA = lConsole
    
    SaveSettingDb "Events", LOGS_DATA
    If c_DDE.Main Then c_DDE.SendData "056", LOGS_DATA
    
End Sub
Private Sub mnuHelp_Click(Index As Integer)
    'Debug.Print c_DDE.FindDDE("Nokilon-SERVER-DDE"), c_DDE.Server
    FrmAbout.Show 1, Me
End Sub

Private Sub BtnMain_BeforePaint(Index As Integer, Hdc As Long, hGraphic As Long, ByVal Width As Long, ByVal Height As Long, Evt As JButtonState, Cancel As Boolean)
    c_Draw.Graphic = hGraphic
    Select Case Evt
        Case lHotBtn
            c_Draw.DrawRectangle 0, 0, Width, Height, &H8000000D, 15, 1, &H8000000D, 20
        Case lDownBtn
            c_Draw.DrawRectangle 0, 0, Width, Height, &H8000000D, 25, 1, &H8000000D, 30
        Case Else
            'FillGradient Hdc, 0, 0, Width, Height, PicTitle(1).BackColor, &H80000016, True
    End Select
End Sub
Private Sub btnMain_Click(Index As Integer)
    Select Case Index
        Case 0 'Cancel Transfer
            If lv2.SelectedItemData = 0 Then Exit Sub
            c_DDE.SendData ("013" & lv2.SelectedItemData)
            
        Case 1 'Reload transfers
        
            lv2.Clear
            lv2.NoDraw = True
            c_DDE.SendData "200" 'LOAD TRANSFERS
            lv2.NoDraw = False
            
        Case 2 'Clear transfer
        
            If lv2.ItemCount = 0 Then Exit Sub
            c_Timer.DestroyTimer
            lv2.NoDraw = True
            lv2.Clear
            c_DDE.SendData "014"
            c_DDE.SendData "200" 'LOAD TRANSFERS
            lv2.NoDraw = False
            c_Timer.CreateTimer 1000
             
    End Select
End Sub


Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    If Button <> 2 Then Exit Sub
    If c_DDE.Main = 0 Then Exit Sub
    If c_DDE.SendData("011" & lv.SelectedItemData) = vbNullString And lv.SelectedItemData <> 0 Then lv.ItemData(lv.SelectedItem) = 0
    
    c_mnu1.ItemDisabled(0) = (lv.SelectedItemData = 0)
    c_mnu1.ItemDisabled(1) = (lv.SelectedItemData = 0)
    
    Select Case c_mnu1.PopupMenu
        Case 100: c_DDE.SendData ("012" & lv.SelectedItemData)
        Case 101:
                lv.Clear
                c_DDE.SendData "100" 'Load users
    End Select
    
End Sub
Private Sub lv_ItemDrawData(ByVal Item As Long, ByVal Column As Long, ForeColor As Long, BackColor As Long, BorderColor As Long, ItemIdent As Long)
    If Column = 0 Then ForeColor = &H808080   '&H59BCF2
End Sub


Private Sub lv2_ItemDblClick(ByVal Item As Long, ByVal Column As Long)
    '
End Sub
Private Sub lv2_SelectionChanged(ByVal Item As Long, ByVal Column As Long)
    btnMain(0).Enabled = Item <> -1 And lv2.ItemCount > 0 And lv2.SelectedItemData <> 0
End Sub
Private Sub lv2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'On Error Resume Next
Dim i As Long

    If Button <> 2 Then Exit Sub
    If c_DDE.Main = 0 Then Exit Sub
    If c_DDE.SendData("011" & lv2.SelectedItemData) = vbNullString And lv2.SelectedItemData <> 0 Then lv2.ItemData(lv2.SelectedItem) = 0
    
    c_mnu2.ItemDisabled(0) = (lv2.SelectedItemData = 0)
    'c_mnu2.ItemDisabled(1) = (lv2.ItemCount = 0)
    
    Select Case c_mnu2.PopupMenu
        Case 100: 'CANCEL
            If lv2.SelectedItemData = 0 Then Exit Sub
            c_DDE.SendData ("013" & lv2.SelectedItemData)

        Case 101: 'RELOAD
            
            lv2.NoDraw = True
            lv2.Clear
            c_DDE.SendData "200" 'LOAD TRANSFERS
            lv2.NoDraw = False
            
        Case 102: 'CLEAR
        
            If lv2.ItemCount = 0 Then Exit Sub
            c_Timer.DestroyTimer
            lv2.NoDraw = True
            lv2.Clear
            c_DDE.SendData "014"
            c_DDE.SendData "200" 'LOAD TRANSFERS
            lv2.NoDraw = False
            c_Timer.CreateTimer 1000
            
        Case 103:
            c_mnu2.SubMenu("Submnu").ItemCheck(0) = Not c_mnu2.SubMenu("Submnu").ItemCheck(0)
            c_DDE.SendData "054" & c_mnu2.SubMenu("Submnu").ItemCheck(0)
            SaveSettingDb "Rmvct", c_mnu2.SubMenu("Submnu").ItemCheck(0)
        Case 104
            c_mnu2.SubMenu("Submnu").ItemCheck(1) = Not c_mnu2.SubMenu("Submnu").ItemCheck(1)
            c_DDE.SendData "055" & c_mnu2.SubMenu("Submnu").ItemCheck(0)
            SaveSettingDb "Rmvit", c_mnu2.SubMenu("Submnu").ItemCheck(1)
            
    End Select
    
End Sub
Private Sub lv2_ItemDraw(ByVal Item As Long, ByVal Column As Long, Hdc As Long, hGraphic As Long, ByVal x As Long, ByVal y As Long, ByVal W As Long, ByVal H As Long, CancelDraw As Boolean)
On Error Resume Next

    If Column = 1 Then
        c_SysIml.DrawFileIcon Hdc, x + (4 * mdpi_), y + ((H - c_SysIml.Height) \ 2), lv2.ItemText(Item, 6) & lv2.ItemText(Item, 1)
    End If
    If Column <> 2 Then Exit Sub

    Dim lColor  As Long
    Dim lp      As Long

    lp = Val(lv2.ItemTag(Item, 2))
    If lp > 100 Then lp = 100
    If lp < 0 Then lp = 0
    
    lColor = &H19D73E
    If Left(lv2.ItemText(Item, 2), 5) = "Error" Then lColor = &HE35F5
    If Left(lv2.ItemText(Item, 2), 9) = "Cancelado" Then lColor = &H3B9BF3
    
     If lv2.ItemIcon(Item, 1) = 2 Then 'And lv2.ItemText(Item, 2) = "Recibiendo" Then
        DrawWaiting Hdc, x + (4 * mdpi_), y + ((H - (7 * mdpi_)) \ 2), 14 * mdpi_, 7 * mdpi_, lp, lColor
     Else
        DrawCircularProgress Hdc, x + (4 * mdpi_), y + ((H - (14 * mdpi_)) \ 2), 14 * mdpi_, 14 * mdpi_, lp, lColor, hGraphic
     End If
     
End Sub
Private Sub lv2_ItemDrawData(ByVal Item As Long, ByVal Column As Long, ForeColor As Long, BackColor As Long, BorderColor As Long, ItemIdent As Long)
    If Column = 1 Then ItemIdent = 22
    If Column = 2 Then ItemIdent = 19
    If Column = 7 Then ForeColor = &H717175
    
End Sub
Private Sub lv2_ItemDrawMeasureText(ByVal Item As Long, ByVal Column As Long, x As Long, y As Long, x2 As Long, y2 As Long)
    If Column = 1 Then x2 = x2 - (6 * mdpi_):
End Sub



Private Sub c_DDE_Request(ByVal Plug As Long, ByVal Key As String, Cancel As Long)
'
End Sub
Private Sub c_DDE_Arrival(ByVal data As String, ByVal Plug As Long, ByVal Key As String, REPLY As String)
Dim sCmd        As String
Dim sElmnt()    As String
Dim mc          As Long
Dim i           As Long

    'Debug.Print "DDE ARRIVAL: " & data
    sCmd = Left$(data, 3)
    data = Right$(data, Len(data) - 3)
    
    Select Case sCmd
    
        Case "000" 'PREV INSTANCE
            If Me.WindowState = 1 Then Me.WindowState = 0
            WindowOnTop Me.hWnd, True
            WindowOnTop Me.hWnd, False
            Me.SetFocus
            
        Case "001" 'SERVER RUN
        
            mvDDEConnect
            If Me.WindowState = 1 Then Me.WindowState = 0
            WindowOnTop Me.hWnd, True
            WindowOnTop Me.hWnd, False
            Me.SetFocus
 
        Case "020": 'SERVER CHANGED STATUS
            mvOnServerStatus CBool(data)
            
        Case "101"  'USER Login
        
            '{WS, IP ,UserName, BytesSent, BytesReceived}
            sElmnt = Split(data, vbNullChar)
            i = lv.AddItem(sElmnt(0), , sElmnt(0))
            lv.SetRow i, sElmnt(0), sElmnt(2), sElmnt(1), FmtSize(sElmnt(3)), FmtSize(sElmnt(4))
            lv.SetRowIcons i, -1, 0, 1, 2, 3
            
        Case "102"  'USER Logout
        
            For i = lv.ItemCount - 1 To 0 Step -1
                If lv.ItemText(i) = data Then
                    lv.RemoveItem i
                End If
            Next
                
            For i = lv2.ItemCount - 1 To 0 Step -1
                If lv2.ItemTag(i) = data Then
                    ' - Clear??
                    If (c_mnu2.SubMenu("Submnu").ItemCheck(1)) Or (c_mnu2.SubMenu("Submnu").ItemCheck(0)) Then
                        lv2.RemoveItem i
                    Else
                        lv2.ItemData(i) = 0
                        lv2.ItemTag(i) = vbNullString
                        If Left(lv2.ItemText(i, 2), 10) <> "Completado" And Left(lv2.ItemText(i, 2), 9) <> "Cancelado" Then lv2.ItemText(i, 2) = "Error"
                    End If
                    
                    'Disabled cancel button
                    If lv2.SelectedItem = i Then lv2_SelectionChanged i, 0
                End If
            Next
        
        Case "201" 'START FILE
            Dim lWS1 As Long
            
            sElmnt = Split(data, vbNullChar)
            '---------------------------------------------------------------------------
            ' 0 WS, 1 FileName, 2 FTPMode, 3 lTotalBytes, 4 tStamp, 5 WsM, 6 UserName
            '---------------------------------------------------------------------------
            
            With lv2
            
                i = .AddItem(sElmnt(6), 0, sElmnt(0), sElmnt(5))
        
                .SetItem i, 1, GetFileTitle(sElmnt(1)), IIf(sElmnt(2) = 1, 1, 2)            'File
                .ItemText(i, 2) = IIf(sElmnt(2) = 1, "Descargando - 0%", "Recibiendo")      'State
                .ItemText(i, 3) = "0 Bytes"                                                 'Bytes
                .ItemText(i, 4) = IIf(sElmnt(2) = 1, FmtSize(sElmnt(3)), "-")            'Total Bytes
                .ItemText(i, 5) = "-"                                                       'Speed
                .SetItem i, 6, GetFilePath(sElmnt(1)), 3                                    'Ruta
                .ItemText(i, 7) = sElmnt(4)                                                 'Time
                '.ItemText(i, 8) = sElmnt(0)                                                'ID
                
                .ItemTag(i, 2) = 0                                                          '[Save Percent]
            End With
            
        Case "202" 'END FILE
            
            sElmnt = Split(data, vbNullChar)
            '-------------------------------------------------------------------------------
            ' WS, FTPMode, lBytes, Result, lPercent | WSM
            '-------------------------------------------------------------------------------
            ' FTPMode = [emNone, emDownload, emUpload]
            ' Result  = [emSuccess, emError, emCancel]
            
            i = lv2.ItemFindData(Val(sElmnt(0)))
            If i = -1 Then Exit Sub
            
            'Disabled cancel button
            If i = lv2.SelectedItem Then btnMain(0).Enabled = False
            
            ' - Clear?
            If (c_mnu2.SubMenu("Submnu").ItemCheck(1) And (sElmnt(3) <> 0)) Or (c_mnu2.SubMenu("Submnu").ItemCheck(0)) Then
                lv2.RemoveItem i
                mvUpdateUserList sElmnt(5)
                Exit Sub
            End If
            
            With lv2
                .NoDraw = True
                .ItemData(i) = 0
                .ItemText(i, 5) = "-"                    'Speed
                .ItemText(i, 3) = FmtSize(sElmnt(2))     'Bytes
                
                If sElmnt(1) = 2 Then .ItemText(i, 4) = FmtSize(sElmnt(2)) 'IF UPLOAD
                
                Select Case sElmnt(3) 'Result
                    Case 0 'emSuccess
                        .ItemText(i, 2) = IIf(sElmnt(1) = 1, "Completado - 100%", "Completado")
                        .ItemTag(i, 2) = IIf(sElmnt(1) = 1, 100, 4) ' [TRANSFER PERCENT TAG FULL]
                    Case Else
                        .ItemText(i, 2) = IIf(sElmnt(3) = 1, "Error", "Cancelado")
                        
                        If sElmnt(1) = 1 Then ' FTPMode = emDownload
                            .ItemTag(i, 2) = sElmnt(4)
                            .ItemText(i, 2) = .ItemText(i, 2) & " - " & sElmnt(4) & "%"
                        Else
                            .ItemTag(i, 2) = 4
                        End If
                        
                End Select
                
                .NoDraw = False
            End With
            mvUpdateUserList sElmnt(5)

            
        Case "220" 'FILE  (Reply to 200)
        
             '--------------------------------------------------------------------------
             ' ws, tstamp, mode, user, fpath, state, cbytes, tbytes, tpval | ws2, speed
             ' 0     1      2    3       4      5      6        7      8   |  9    10
             '--------------------------------------------------------------------------
             
            sElmnt = Split(data, vbNullChar)
            i = lv2.AddItem(sElmnt(3), 0, sElmnt(0), sElmnt(9))
            lv2.SetItem i, 1, GetFileTitle(sElmnt(4)), IIf(sElmnt(2) = 1, 1, 2)       'File
            lv2.SetItem i, 6, GetFilePath(sElmnt(4)), 3
            lv2.ItemText(i, 7) = sElmnt(1)
             
            lv2.ItemText(i, 2) = sElmnt(5)                  'State
            lv2.ItemText(i, 3) = FmtSize(sElmnt(6))      'Bytes
            lv2.ItemText(i, 4) = FmtSize(sElmnt(7))      'Total Bytes
            lv2.ItemText(i, 5) = FmtSpeed(sElmnt(10))       'Speed
            lv2.ItemTag(i, 2) = sElmnt(8)                   'Save Percent

             
        Case Is >= 800: PushLog data, Val(sCmd) - 800

    End Select
    
End Sub
Private Sub c_DDE_Disconnected(ByVal Plug As Long, ByVal Key As String)
    If Key <> "Nokilon-SERVER-DDE" Then Exit Sub
    If LOGS_DATA And mclSERVER Then PushLog "Disconnected from Server"
    mvOnServerConnection False
    c_DDE.InitDDE "Nokilon-GUI2"
End Sub

Private Sub c_Timer_Timer(ByVal ThisTime As Long)
    
    If c_DDE.FindDDE("Nokilon-SERVER-DDE") = 0 Then
        If LOGS_DATA And mclSERVER Then PushLog "Disconnected from Server", enmPSError
        mvOnServerConnection False
        c_DDE.InitDDE "Nokilon-GUI2"
        Exit Sub
    End If

    If lv.ItemCount = 0 Then Exit Sub
    Dim i   As Long
    Dim WS  As Long
    
   
    With lv
        .NoDraw = True
        For i = 0 To .ItemCount - 1
            mvUpdateUserList .ItemData(i), i
        Next
        .NoDraw = False
    End With

    If lv2.ItemCount = 0 Then Exit Sub
    
    Dim sElmnt()    As String
    
    With lv2
        .NoDraw = True
        For i = 0 To .ItemCount - 1
            If .ItemData(i) > 0 Then
            
                'TRANSFER INFO
                sElmnt = Split(c_DDE.SendData("205" & .ItemData(i)), vbNullChar)
                '---------------------------------------------
                ' Mode, Percent, CurrentBytes, Speed
                '---------------------------------------------
            
                If UBound(sElmnt) > -1 Then
                    Select Case sElmnt(0)
                        Case 1 ' emDownload:
                            .ItemTag(i, 2) = sElmnt(1)
                            .ItemText(i, 2) = "Descargando - " & .ItemTag(i, 2) & "%"
                        Case 2 'emUpload:
                            If Val(.ItemTag(i, 2)) > 2 Then .ItemTag(i, 2) = 0 Else .ItemTag(i, 2) = Val(.ItemTag(i, 2)) + 1
                    End Select
                    .ItemText(i, 3) = FmtSize(sElmnt(2))     'Bytes
                    .ItemText(i, 5) = FmtSpeed(sElmnt(3))       'Speed
                    
                Else
                    .ItemData(i) = 0
                    
                    'Disabled cancel button
                    If .SelectedItem = i Then lv2_SelectionChanged i, 0
   
                End If
            End If
        Next
        .NoDraw = False
    End With
    
End Sub

Property Get RTE() As cRichEdit: Set RTE = c_RTE: End Property
Property Get DDE() As cDDE: Set DDE = c_DDE: End Property
Private Sub Form_Unload(Cancel As Integer)
    
    mvFormPos False

    Set c_Timer = Nothing
    Set c_mnu1 = Nothing
    Set c_mnu2 = Nothing
    Set c_RTE = Nothing
    Set c_DDE = Nothing
    Set c_Draw = Nothing
End Sub


Private Sub mvDDEConnect()
On Error GoTo e

    If c_DDE.Main <> 0 Then Exit Sub
    If c_DDE.Connect("Nokilon-SERVER-DDE", "Nokilon-GUI") Then
        mvOnServerConnection c_DDE.Main <> 0
        If ShowProps_ Then FrmServer.Show , Me
    Else
        c_DDE.InitDDE "Nokilon-GUI2"
        If LOGS_DATA And mclSERVER Then PushLog "Couldn't connect to Server.", enmPSError
    End If
e:
End Sub


Private Sub mvUpdateUserList(ByVal WS As Long, Optional ByVal i As Long = -1)
On Error GoTo e
    If i = -1 Then i = lv.ItemFindData(WS)
    If i = -1 Then Exit Sub
    
    Dim sBytes() As String
    sBytes = Split(c_DDE.SendData("010" & WS), vbNullChar)
    '{BYTES_SENT, BYTES_RECEIVED}
    If UBound(sBytes) > -1 Then
        lv.ItemText(i, 3) = FmtSize(sBytes(0))
        lv.ItemText(i, 4) = FmtSize(sBytes(1))
    End If
e:
End Sub


Private Sub mvOnServerConnection(ByVal Connection As Boolean)

    If lv.ItemCount Then lv.Clear
    If lv2.ItemCount Then lv2.Clear
    
    If Connection Then
    
        c_DDE.SendData "100" 'Load users
        c_DDE.SendData "200" 'Load transfers
        lv2.Redraw
        
        mvOnServerStatus CBool(c_DDE.SendData("004"))
        Set c_Timer = New cTimer
        c_Timer.CreateTimer 1000
        
        c_DDE.SendData "020" 'Log state
    Else
        Set c_Timer = Nothing
        Sb.PanelText(2) = "": Sb.PanelIconIndex(2) = -1     'Status
        Sb.PanelText(3) = "": Sb.PanelIconIndex(3) = -1     'IP
        Sb.PanelText(4) = "": Sb.PanelIconIndex(4) = -1     'Port
    End If
    
    Sb.PanelText(1) = IIf(Connection, "Conectado", "No conectado")
    Sb.Refresh
    
    mnuServer(0).Enabled = Not Connection
    mnuServer(1).Enabled = Connection
    mnuServer(3).Enabled = Connection

    If Not Connection Then btnMain(0).Enabled = False
    btnMain(1).Enabled = Connection
    btnMain(2).Enabled = Connection
    
End Sub
Private Sub mvOnServerStatus(ByVal mbState As Boolean)
    
    Dim sElmnt() As String
    sElmnt = Split(c_DDE.SendData("007"), vbNullChar)
    
    If CheckArgs(sElmnt, 1) Then
        Sb.PanelText(3) = "IP : " & sElmnt(0): Sb.PanelIconIndex(3) = 3         'IP
        Sb.PanelText(4) = "Puerto : " & sElmnt(1): Sb.PanelIconIndex(4) = 4     'Puerto
    End If

    If mbState Then
        Sb.PanelText(2) = "Servicio activo": Sb.PanelIconIndex(2) = 2     'Status
    Else
        Sb.PanelText(2) = "Servicio inactivo": Sb.PanelIconIndex(2) = 1    'Status
    End If
    
    
End Sub

Private Sub mvGuiPrevInstance()
     With New cDDE
        If .Connect("Nokilon-GUI", "Nokilon-GUI" & Me.hWnd) Then Call .SendData("000"): Exit Sub
        If .Connect("Nokilon-GUI2", "Nokilon-GUI2" & Me.hWnd) Then Call .SendData("000")
    End With
End Sub
Private Sub mvAdjustControls()
On Error Resume Next
    With PicSpliter
    
        Ts.Height = .Top - Ts.Top - (4 * mdpi_)
        Call Ts.UpdateSize
        DoEvents
        
        'DoEvents
        PicTitle.Top = .Top + .Height + (4 * mdpi_)
        lv2.Top = .Top + .Height + (6 * mdpi_) + PicTitle.ScaleHeight
        lv2.Height = Me.ScaleHeight - lv2.Top - (28 * mdpi_)  'Sb.Height
        
        PicLog.Height = (Ts.Height * Screen.TwipsPerPixelY) - PicLog.Top - (8 * Screen.TwipsPerPixelY) '.Top - PicLog.Top - (12 * mdpi_)
        lv.Height = (Ts.Height * Screen.TwipsPerPixelY) - lv.Top - (8 * Screen.TwipsPerPixelY)
        
    End With
    c_RTE.Move 1 * mdpi_, 1 * mdpi_, PicLog.ScaleWidth - (2 * mdpi_), PicLog.ScaleHeight - (2 * mdpi_)
End Sub

Private Sub mvFormPos(Optional ByVal Startup As Boolean)
Dim sTmp As String

    If Startup Then
    
        Dim lState As Long
        
        sTmp = GetSetting("Gui", "FormPos")
        lState = Val(GetSetting("Gui", "FormState"))
        
        If Len(sTmp) Then
            Dim sTmps() As String
            sTmps = Split(sTmp, ",")
            If Not CheckArgs(sTmps, 4) Then Exit Sub
            If lState <> 2 Then
                
                If sTmps(0) < 0 Then sTmps(0) = 0
                If sTmps(1) < 0 Then sTmps(1) = 0
                If sTmps(2) < 300 * mdpi_ Then sTmps(2) = 300 * mdpi_
                If sTmps(3) < 300 * mdpi_ Then sTmps(3) = 300 * mdpi_
                
                Me.Left = sTmps(0): Me.Top = sTmps(1)
                Me.Width = sTmps(2): Me.Height = sTmps(3)
                Me.Move sTmps(0), sTmps(1), sTmps(2), sTmps(3)
            End If
            
            If Val(sTmps(4)) < 8 Then sTmps(4) = 8
            If Val(sTmps(4)) Then lPH1 = Val(sTmps(4))
            
        End If
        
        Ts.SelectedItem = Val(GetSetting("Gui", "Tab"))
        If lState = 2 Then Me.WindowState = 2
        
    Else
        sTmp = Me.Left & "," & Me.Top & "," & Me.Width & "," & Me.Height & "," & lPH1
        If Me.WindowState = 0 Then SaveSetting "Gui", "FormPos", sTmp
        SaveSetting "Gui", "Tab", Ts.SelectedItem
        SaveSetting "Gui", "FormState", Me.WindowState
    End If
    
End Sub
