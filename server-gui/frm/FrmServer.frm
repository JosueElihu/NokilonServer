VERSION 5.00
Begin VB.Form FrmServer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Propiedades"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4755
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   225
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   317
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Chk 
      Caption         =   "No mostrar esta ventana al conectarse al servidor"
      Height          =   270
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   2280
      Width           =   3855
   End
   Begin NokilonGui.JButton JButton1 
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   2760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      SkinRes         =   ""
      Text            =   "Aceptar"
      TextAlign       =   2
      BitmapAlign     =   1
      Margins         =   "8,8"
      ButtonType      =   0
      Value           =   0   'False
      EDown           =   -1  'True
      AmbientImage    =   0   'False
      BitmapResize    =   "14x14"
      BitmapColor     =   -1
      BitmapSpace     =   5
      NoBkgnd         =   0   'False
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
      Image           =   "FrmServer.frx":000C
   End
   Begin VB.Timer Timer1 
      Left            =   3720
      Top             =   120
   End
   Begin VB.Label lblSpeedU 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ": 0 Bytes/seg"
      Height          =   195
      Left            =   2370
      TabIndex        =   12
      Top             =   1800
      Width           =   960
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Velocidad de subida"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   11
      Top             =   1800
      Width           =   1560
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   24
      X2              =   296
      Y1              =   176
      Y2              =   176
   End
   Begin VB.Label lblSpeedD 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ": 0 Bytes/seg"
      Height          =   195
      Left            =   2370
      TabIndex        =   10
      Top             =   1560
      Width           =   960
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Velocidad de descarga"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   9
      Top             =   1560
      Width           =   1740
   End
   Begin NokilonGui.JImage JImage1 
      Height          =   480
      Left            =   360
      Top             =   120
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      imgStream       =   "FrmServer.frx":0949
      eScale          =   1
      Color           =   -1
      Alpha           =   100
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   24
      X2              =   296
      Y1              =   144
      Y2              =   144
   End
   Begin VB.Label lblServerName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "My FTP Server"
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
      Left            =   960
      TabIndex        =   8
      Top             =   240
      Width           =   1380
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   24
      X2              =   296
      Y1              =   48
      Y2              =   48
   End
   Begin VB.Label lblPort 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ": 4321"
      Height          =   195
      Left            =   2370
      TabIndex        =   7
      Top             =   1320
      Width           =   465
   End
   Begin VB.Label lblIP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":127.0.0.1"
      Height          =   195
      Left            =   2370
      TabIndex        =   6
      Top             =   1080
      Width           =   675
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ": No activo"
      Height          =   195
      Left            =   2370
      TabIndex        =   5
      Top             =   840
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Puerto"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Servidor"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   645
   End
End
Attribute VB_Name = "FrmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mtSpeed
  lBytes1   As Currency
  lBytes2   As Currency
  lSpeed1   As Currency
  lSpeed2   As Currency
End Type

Private sElmnt()    As String
Private mSpeed      As mtSpeed
Private mbLoad      As Boolean

Private Sub Form_Load()
    
    sElmnt = Split(FrmMain.DDE.SendData("005"), vbNullChar)
    
    If Not CheckArgs(sElmnt, 5) Then PushLog "Could not get server information", enmPSError: Unload Me: Exit Sub
    '------------------------------------------------------------------
    ' NAME, SERVER.SOCKET, IP, PORT, BYTES_SEND, BYTES_RECEIVED
    '-------------------------------------------------------------------

    lblServerName = sElmnt(0)
    lblStatus = ": " & IIf(Val(sElmnt(1)), "Encendido", "Apagado")
    lblIP = ": " & sElmnt(2)
    lblPort = ": " & sElmnt(3)
    
    'mSpeed.lBytes1 = Val(sElmnt(4)) 'BYTES_SENT
    'mSpeed.lBytes2 = Val(sElmnt(5)) 'BYTES_RECEIVED
    
    mSpeed.lSpeed1 = Val(sElmnt(4)) 'BYTES_SENT
    mSpeed.lSpeed2 = Val(sElmnt(5)) 'BYTES_RECEIVED
    
    mbLoad = True
    If Not ShowProps_ Then Chk(0).Value = 1
    mbLoad = False
    Timer1.Interval = 1000
    
End Sub
Private Sub JButton1_Click()
    Unload Me
End Sub


Private Sub Timer1_Timer()

    If FrmMain.DDE.Main = 0 Then Unload Me
    
    
    '-------------------------------------
    ' BYTES_SENT, BYTES_RECEIVED, STATE
    '-------------------------------------
            
    sElmnt = Split(FrmMain.DDE.SendData("006"), vbNullChar)
    If UBound(sElmnt) < 0 Then Exit Sub
    
    'lblBytesSent = ": " & FmtSize(sElmnt(0))
    'lblBytesReceived = ": " & FmtSize(sElmnt(1))

    lblSpeedD = ": " & FmtSize(sElmnt(0) - mSpeed.lSpeed1) & "/seg"
    lblSpeedU = ": " & FmtSize(sElmnt(1) - mSpeed.lSpeed2) & "/seg"
    
    mSpeed.lSpeed1 = sElmnt(0)
    mSpeed.lSpeed2 = sElmnt(1)
    
    lblStatus = ": " & IIf(Val(sElmnt(2)), "Encendido", "Apagado")
    
End Sub
Private Sub Chk_Click(Index As Integer)
    If mbLoad Then Exit Sub
    ShowProps_ = Not (Chk(0).Value = 1)
    SaveSetting "Gui", "Server-info", ShowProps_
End Sub
Private Sub Form_Unload(Cancel As Integer)
'
End Sub

