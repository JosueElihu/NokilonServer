VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de..."
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   154
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin NokilonServer.JButton btnMain 
      Height          =   510
      Left            =   3240
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   900
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
      BitmapSpace     =   3
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
      Image           =   "FrmAbout.frx":000C
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00BFBFBF&
      X1              =   24
      X2              =   320
      Y1              =   104
      Y2              =   104
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Powered By FloresSystems - 2022"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   2940
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desarollado por J. Elihu - elihulgts.10@gmail.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   4230
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00BFBFBF&
      X1              =   24
      X2              =   320
      Y1              =   56
      Y2              =   56
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nokilon Server"
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
      TabIndex        =   1
      Top             =   240
      Width           =   1395
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.5"
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   810
   End
   Begin NokilonServer.JImage JImage1 
      Height          =   480
      Left            =   360
      Top             =   240
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      imgStream       =   "FrmAbout.frx":0814
      eScale          =   1
      Color           =   -1
      Alpha           =   100
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit






Private Sub Form_Load()
    'PutIcon32Bit Me.hWnd, "ALPHA"
    Label2 = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
Private Sub Form_Deactivate()
    Unload Me
End Sub
Private Sub btnMain_Click()
    Unload Me
End Sub
