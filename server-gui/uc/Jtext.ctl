VERSION 5.00
Begin VB.UserControl Jtext 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   PropertyPages   =   "Jtext.ctx":0000
   ScaleHeight     =   35
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "Jtext.ctx":000F
   Begin VB.TextBox Edit 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Jtext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left    As Long
    Top     As Long
    Width   As Long
    Height  As Long
End Type

Private Type BitmapData
    Width   As Long
    Height  As Long
    stride  As Long
    PixelFormat As Long
    Scan0Ptr    As Long
    ReservedPtr As Long
End Type

Private Type GUID
    Data1       As Long
    Data2       As Integer
    Data3       As Integer
    Data4(0 To 7) As Byte
End Type

'/DPI
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal Hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal Hdc As Long) As Long

'/GDI+
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As Long, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal Hdc As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal Graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal ARGB As Long, ByRef brush As Long) As Long
Private Declare Function GdipClosePathFigures Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipSetPenMode Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mPenMode As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus.dll" (ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStride As Long, ByVal mPixelFormat As Long, ByVal mScan0 As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipBitmapUnlockBits Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mLockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapLockBits Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mRect As RECT, ByVal mFlags As Long, ByVal mPixelFormat As Long, ByRef mLockedBitmapData As BitmapData) As Long
Private Declare Function GdipGetImageHeight Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mHeight As Long) As Long
Private Declare Function GdipGetImageWidth Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mWidth As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal Image As Long, ByRef Width As Single, ByRef Height As Single) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal Graphics As Long, ByVal InterpolationMode As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal Graphics As Long, ByVal PixelOffsetMode As Long) As Long

'/


Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function SetRect Lib "user32" (lpRect As Any, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As Any) As Long
Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)


Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal sFileName As String, clsidEncoder As Any, encoderParams As Any) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Any, ByRef Image As Long) As Long


Private Const PixelFormat32bppPARGB As Long = &HE200B
Private Const PASSWORD_CHAR As String = "•"

Public Enum JTextImageAlignment
    ImageInLeft
    ImageInRight
End Enum


Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)

Private m_bvData()  As Byte

Private m_token         As Long
Private m_BmpS          As Long
Private m_Bmp           As Long
Private m_BmpRct        As RECT
Private m_BmpSrcW       As Single
Private m_BmpSrcH       As Single

Private m_BackColor     As OLE_COLOR
Private m_fBack         As OLE_COLOR
Private m_cBorder       As OLE_COLOR
Private m_fBorder       As OLE_COLOR
Private m_ShadowColor   As OLE_COLOR
Private m_AutoSel       As Boolean
Private m_Pwd           As Boolean
Private m_Round         As Long
Private m_Shadow        As Long
Private m_BmpAlignment  As JTextImageAlignment
Private m_BmpSize       As String

Private bFocus          As Boolean
Private dpiScale        As Double



Private Sub UserControl_Initialize()
    ppGdipStart True
    dpiScale = GetWindowsDPI
End Sub
Private Sub UserControl_Terminate()
    If m_BmpS Then GdipDisposeImage m_BmpS
    If m_Bmp Then GdipDisposeImage m_Bmp
    ppGdipStart False
End Sub

Private Sub UserControl_InitProperties()
    m_BackColor = vbWhite
    m_fBack = vbWhite
    m_cBorder = &HB2ACA5
    m_fBorder = &HE8A859   'RGB(148, 199, 240)
    m_ShadowColor = &HE8A859
    Edit.Text = Extender.Name
    m_Round = 3
    m_AutoSel = True
    m_Shadow = 3
    m_BmpSize = "0x0"
End Sub
Private Sub UserControl_AmbientChanged(PropertyName As String)
    Call ppCopyAmbient
    Call ppDraw
End Sub
Private Sub UserControl_EnterFocus()
    bFocus = True
    If m_AutoSel Then
        Edit.SelStart = 0
        Edit.SelLength = Len(Edit.Text)
    End If
    ppDraw
End Sub
Private Sub UserControl_ExitFocus()
    bFocus = False
    ppDraw
End Sub
Private Sub UserControl_Resize()
On Error Resume Next
Dim Th      As Long
Dim Px      As Long
Dim bS      As Long

    With UserControl
        bS = (2 * dpiScale) + (m_Shadow * dpiScale) + (m_Round * dpiScale)
        If m_Round = 0 Then bS = (bS + (2 * dpiScale))
        
        If m_Bmp Then
            Px = m_BmpRct.Width + (2 * dpiScale)
            m_BmpRct.Left = bS
            m_BmpRct.Top = (.ScaleHeight - m_BmpRct.Height) \ 2
            If m_BmpAlignment = 1 Then m_BmpRct.Left = UserControl.ScaleWidth - (bS + m_BmpRct.Width)
        End If
        
        Th = .TextHeight("Ájq\")
        Select Case m_BmpAlignment
            Case 0: Edit.Move bS + Px, (.ScaleHeight - Th) \ 2, .ScaleWidth - (bS * 2) - Px, Th
            Case 1: Edit.Move bS, (.ScaleHeight - Th) \ 2, .ScaleWidth - (bS * 2) - Px, Th
        End Select
        
    End With
    
    ppCopyAmbient
    If Ambient.UserMode Then ppCreateShadow
    ppDraw
    
End Sub
Private Sub UserControl_Show()
    Call ppDraw
End Sub

Private Sub Edit_Change()
'
End Sub
Private Sub Edit_Click()
'
End Sub
Private Sub Edit_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub Edit_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub Edit_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub


Property Get DPI() As Double: DPI = dpiScale: End Property
Property Get ForeColor() As OLE_COLOR: ForeColor = Edit.ForeColor: End Property
Property Let ForeColor(ByVal Value As OLE_COLOR)
    Edit.ForeColor = Value
    PropertyChanged "Fore"
End Property
Property Get BackColor() As OLE_COLOR: BackColor = m_BackColor: End Property
Property Let BackColor(ByVal Value As OLE_COLOR)
    m_BackColor = Value
    ppDraw
    PropertyChanged "Back"
End Property
Property Get BackColorFocus() As OLE_COLOR: BackColorFocus = m_fBack: End Property
Property Let BackColorFocus(ByVal Value As OLE_COLOR)
    m_fBack = Value
    PropertyChanged "BackFocus"
End Property
Property Get BorderColor() As OLE_COLOR: BorderColor = m_cBorder: End Property
Property Let BorderColor(ByVal Value As OLE_COLOR)
    m_cBorder = Value
    ppDraw
    PropertyChanged "BorderColor"
End Property
Property Get BorderColorFocus() As OLE_COLOR: BorderColorFocus = m_fBorder: End Property
Property Let BorderColorFocus(ByVal Value As OLE_COLOR)
    m_fBorder = Value
    PropertyChanged "BorderColorF"
End Property
Property Get Text() As String: Text = Edit.Text: End Property
Attribute Text.VB_Description = "Asigna o recupera el texto"
Attribute Text.VB_UserMemId = 0
Public Property Let Text(ByVal Value As String)
    Edit = Value
    ppDraw
    PropertyChanged "Text"
End Property
Property Get SelFocus() As Boolean: SelFocus = m_AutoSel: End Property
Property Let SelFocus(ByVal Value As Boolean)
    m_AutoSel = Value
    PropertyChanged "AutoSel"
End Property
Property Get Readonly() As Boolean: Readonly = Edit.Locked: End Property
Property Let Readonly(ByVal Value As Boolean)
    Edit.Locked = Value
    PropertyChanged "ReadOnly"
End Property
Property Get CornerRound() As Long: CornerRound = m_Round: End Property
Property Let CornerRound(ByVal Value As Long)
    m_Round = Value
    UserControl_Resize
    PropertyChanged "BorderRoud"
End Property
Property Get Font() As StdFont: Set Font = Edit.Font: End Property
Property Set Font(ByVal Value As StdFont)
    Set Edit.Font() = Value
    Set UserControl.Font() = Value
    PropertyChanged "Font"
End Property
Property Get Alignment() As AlignmentConstants: Alignment = Edit.Alignment: End Property
Property Let Alignment(ByVal Value As AlignmentConstants)
    Edit.Alignment = Value
    PropertyChanged "Alignment"
End Property

Property Get ShadowSize() As Long: ShadowSize = m_Shadow: End Property
Property Let ShadowSize(ByVal Value As Long)
    If Value < 0 Then Value = 0
    m_Shadow = Value
    If m_BmpS Then
        GdipDisposeImage m_BmpS
        m_BmpS = 0
    End If
    UserControl_Resize
End Property
Property Get ShadowColor() As OLE_COLOR: ShadowColor = m_ShadowColor: End Property
Property Let ShadowColor(ByVal Value As OLE_COLOR)
    m_ShadowColor = Value
    Call ppCreateShadow
    If bFocus Then ppDraw
    PropertyChanged "ShadowColor"
End Property
Property Get Password() As Boolean: Password = m_Pwd: End Property
Property Let Password(ByVal Value As Boolean)
    m_Pwd = Value
    Edit.PasswordChar = IIf(m_Pwd, PASSWORD_CHAR, "")
    PropertyChanged "Pwd"
End Property


Property Get ImageSize() As String: ImageSize = m_BmpSrcW & "x" & m_BmpSrcH: End Property
Property Let ImageSize(ByVal Value As String): End Property
Property Get ImageResize() As String: ImageResize = m_BmpSize: End Property
Property Let ImageResize(ByVal Value As String)
On Error Resume Next
Dim lSep As String
Dim lW   As Long
Dim lH   As Long

    If InStr(Value, "*") Then lSep = "*"
    If InStr(LCase(Value), "x") Then lSep = "x"
    lW = Val(Split(Value, lSep)(0))
    lH = Val(Split(Value, lSep)(1))
    
    m_BmpSize = lW & "x" & lH
    m_BmpRct.Width = IIf(lW > 0, lW * dpiScale, m_BmpSrcW)
    m_BmpRct.Height = IIf(lW > 0, lW * dpiScale, m_BmpSrcH)
    
    Call UserControl_Resize
    PropertyChanged "ImageResize"
End Property
Property Get ImageAlignment() As JTextImageAlignment: ImageAlignment = m_BmpAlignment: End Property
Property Let ImageAlignment(ByVal Value As JTextImageAlignment)
    m_BmpAlignment = Value
    Call UserControl_Resize
    PropertyChanged "ImageAlignment"
End Property

Friend Function ppgGetStream() As Byte(): ppgGetStream = m_bvData: End Function
Friend Function ppgSetStream(bvData() As Byte)
    m_bvData = bvData
    Call SetPictureStream(bvData)
    PropertyChanged "ImageStream"
End Function

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_BackColor = .ReadProperty("Back", vbWhite)
        m_fBack = .ReadProperty("BackFocus", vbWhite)
        m_cBorder = .ReadProperty("BorderColor", &HB2ACA5)
        m_fBorder = .ReadProperty("BorderColorF", &HE8A859)
        m_AutoSel = .ReadProperty("AutoSel", True)
        m_Round = .ReadProperty("BorderRound", 3)
        m_Shadow = .ReadProperty("ShadowSize", 0)
        m_ShadowColor = .ReadProperty("ShadowColor", &HE8A859)
        m_Pwd = .ReadProperty("Pwd", False)
        m_BmpSize = .ReadProperty("ImageResize", "0x0")
        m_BmpAlignment = .ReadProperty("ImageAlignment", 0)
        
        Set Edit.Font() = .ReadProperty("Font", Edit.Font)
        Set UserControl.Font() = .ReadProperty("Font", UserControl.Font)

        Edit.Alignment = .ReadProperty("Alignment", 0)
        Edit.Text = .ReadProperty("Text", "")
        Edit.Locked = .ReadProperty("ReadOnly", False)
        Edit.ForeColor = .ReadProperty("Fore", 0)
        
        If m_Pwd Then Edit.PasswordChar = PASSWORD_CHAR
        
        m_bvData() = .ReadProperty("ImageStream", "")
        Call SetPictureStream(m_bvData())
        If Ambient.UserMode Then Erase m_bvData
        
    End With
    Call UserControl_Resize
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Back", m_BackColor
        .WriteProperty "BackFocus", m_fBack
        .WriteProperty "BorderColor", m_cBorder
        .WriteProperty "BorderColorF", m_fBorder
        .WriteProperty "AutoSel", m_AutoSel
        .WriteProperty "BorderRound", m_Round
        .WriteProperty "ShadowSize", m_Shadow
        .WriteProperty "Font", Edit.Font
        .WriteProperty "Text", Edit.Text
        .WriteProperty "ReadOnly", Edit.Locked
        .WriteProperty "Fore", Edit.ForeColor
        
        .WriteProperty "Alignment", Edit.Alignment
        .WriteProperty "ShadowColor", m_ShadowColor
        .WriteProperty "Pwd", m_Pwd
        .WriteProperty "Font", Edit.Font
        .WriteProperty "ImageResize", m_BmpSize
        .WriteProperty "ImageAlignment", m_BmpAlignment
        .WriteProperty "ImageStream", m_bvData
    End With
End Sub


Public Function SetPictureStream(bvData() As Byte)
On Error GoTo e
Dim hGrph   As Long
Dim hBmp    As Long
Dim lW      As Long
Dim lH      As Long

    If m_Bmp Then
        GdipDisposeImage m_Bmp
        m_Bmp = 0
        m_BmpSrcW = 0
        m_BmpSrcH = 0
    End If
    
    ppCreateBitmapFromStream bvData, hBmp
    If hBmp = 0 Then GoTo e
    GdipGetImageDimension hBmp, m_BmpSrcW, m_BmpSrcH
    m_Bmp = hBmp
    
    lW = Split(m_BmpSize, "x")(0)
    lH = Split(m_BmpSize, "x")(1)
    m_BmpRct.Width = IIf(lW > 0, lW * dpiScale, m_BmpSrcW)
    m_BmpRct.Height = IIf(lW > 0, lW * dpiScale, m_BmpSrcH)
    
e:
    Call UserControl_Resize
End Function

Private Sub ppDraw()
Dim hGrph   As Long
Dim hPath   As Long
Dim hPen    As Long
Dim hBrush  As Long

Dim lW      As Long
Dim lH      As Long
Dim lx      As Long
Dim ly      As Long


    With UserControl
        
        lx = m_Shadow * dpiScale
        ly = m_Shadow * dpiScale
        
        lW = (.ScaleWidth) - (lx * 2)
        lH = (.ScaleHeight) - (ly * 2)
    
        .Cls
        If bFocus Then
            Edit.BackColor = m_fBack
        Else
            Edit.BackColor = m_BackColor
        End If
        
        If GdipCreateFromHDC(.Hdc, hGrph) <> 0 Then Exit Sub
        GdipSetSmoothingMode hGrph, 4 '-> SmoothingModeAntiAlias
        
        If m_BmpS And bFocus Then
            GdipDrawImageRectI hGrph, m_BmpS, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        End If
        hPath = ppRound(lx, ly, lW, lH)
        
        '\Back
        GdipCreateSolidFill ConvertColor(IIf(bFocus, m_fBack, m_BackColor), 100), hBrush
        GdipFillPath hGrph, hBrush, hPath
        GdipDeleteBrush hBrush

        '\Border
        GdipCreatePen1 ConvertColor(IIf(bFocus, m_fBorder, m_cBorder), 100), (1 * dpiScale), &H2&, hPen
        'GdipSetPenMode hPen, &H1
        GdipDrawPath hGrph, hPen, hPath
        GdipDeletePen hPen
        
        If m_Bmp Then
            Call GdipSetInterpolationMode(hGrph, 7&)  'HIGH_QUALYTY_BICUBIC
            Call GdipSetPixelOffsetMode(hGrph, 4&)
            GdipDrawImageRectI hGrph, m_Bmp, m_BmpRct.Left, m_BmpRct.Top, m_BmpRct.Width, m_BmpRct.Height
        End If
        
        
        Call GdipDeletePath(hPath)
        Call GdipDeleteGraphics(hGrph)
    End With
    
End Sub

Private Function ppRound(x As Long, Y As Long, ByVal W As Long, ByVal H As Long) As Long
Dim ePath   As Long

Dim BCLT    As Integer
Dim BCRT    As Integer
Dim BCBR    As Integer
Dim BCBL    As Integer

    W = W - 1 'Antialias pixel
    H = H - 1 'Antialias pixel
    
    BCLT = GetSafeRound(m_Round * dpiScale, W, H)
    BCRT = GetSafeRound(m_Round * dpiScale, W, H)
    BCBR = GetSafeRound(m_Round * dpiScale, W, H)
    BCBL = GetSafeRound(m_Round * dpiScale, W, H)
    
    Call GdipCreatePath(&H0, ePath)
    If BCLT Then GdipAddPathArcI ePath, x, Y, BCLT * 2, BCLT * 2, 180, 90
    If BCLT = 0 Then GdipAddPathLineI ePath, x, Y, x + W - BCRT, Y
        
    If BCRT Then GdipAddPathArcI ePath, x + W - BCRT * 2, Y, BCRT * 2, BCRT * 2, 270, 90
    If BCRT = 0 Then GdipAddPathLineI ePath, x + W, Y, x + W, Y + H - BCBR
        
    If BCBR Then GdipAddPathArcI ePath, x + W - BCBR * 2, Y + H - BCBR * 2, BCBR * 2, BCBR * 2, 0, 90
    If BCBR = 0 Then GdipAddPathLineI ePath, x + W, Y + H, x + BCBL, Y + H
    
    If BCBL Then GdipAddPathArcI ePath, x, Y + H - BCBL * 2, BCBL * 2, BCBL * 2, 90, 90
    If BCBL = 0 Then GdipAddPathLineI ePath, x, Y + H, x, Y + BCLT
    
    GdipClosePathFigures ePath
    ppRound = ePath
    
End Function

Private Sub ppCreateShadow()
Dim hGrph   As Long
Dim hBmp    As Long
Dim hPath   As Long
Dim hPen    As Long
Dim hBrush  As Long

Dim lW      As Long
Dim lH      As Long
Dim lpz     As Long

    If m_BmpS Then GdipDisposeImage m_BmpS: m_BmpS = 0
    If m_Shadow = 0 Then Exit Sub
    
    lpz = m_Shadow * dpiScale
    
    lW = UserControl.ScaleWidth - (lpz * 2)
    lH = UserControl.ScaleHeight - (lpz * 2)
    
    
    GdipCreateBitmapFromScan0 lW, lH, 0&, &HE200B, ByVal 0&, hBmp
    GdipGetImageGraphicsContext hBmp, hGrph
    GdipSetSmoothingMode hGrph, 4& '->SmoothingModeAntiAlias
    
    hPath = ppRound(0, 0, lW, lH)
    GdipCreateSolidFill ConvertColor(m_fBorder, 100), hBrush
    GdipFillPath hGrph, hBrush, hPath
    GdipDeleteBrush hBrush
    
    m_BmpS = ppBlur(hBmp, m_ShadowColor, m_Shadow)

    GdipDeletePath hPath
    GdipDeleteGraphics hGrph
    GdipDisposeImage hBmp
    
End Sub

Private Function ppBlur(hImage As Long, Color As Long, blurDepth As Long, Optional ByVal Width As Long, Optional ByVal Height As Long) As Long
On Error Resume Next
Dim REC As RECT
Dim x As Long, Y As Long
Dim hImgShadow As Long
Dim bmpData1 As BitmapData
Dim bmpData2 As BitmapData
Dim t2xBlur As Long
Dim R As Long, G As Long, B As Long
Dim Alpha As Byte
Dim lSrcAlpha As Long, lDestAlpha As Long
Dim dBytes() As Byte
Dim srcBytes() As Byte
Dim vTally() As Long
Dim tAlpha As Long, tColumn As Long, tAvg As Long
Dim initY As Long, initYstop As Long, initYstart As Long
Dim initX As Long, initXstop As Long
    
    If hImage = 0& Then Exit Function
 
    If Width = 0& Then Call GdipGetImageWidth(hImage, Width)
    If Height = 0& Then Call GdipGetImageHeight(hImage, Height)
 

    t2xBlur = blurDepth * 2
 
    R = Color And &HFF
    G = (Color \ &H100&) And &HFF
    B = (Color \ &H10000) And &HFF
 
    SetRect REC, 0, 0, Width, Height
 
    ReDim srcBytes(REC.Width * 4 - 1&, REC.Height - 1&)
    With bmpData1
        .Scan0Ptr = VarPtr(srcBytes(0&, 0&))
        .stride = 4& * REC.Width
    End With
   
    Call GdipBitmapLockBits(hImage, REC, &H4 Or &H1, PixelFormat32bppPARGB, bmpData1)
 
 
    SetRect REC, 0, 0, Width + t2xBlur, Height + t2xBlur
    Call GdipCreateBitmapFromScan0(REC.Width, REC.Height, 0&, PixelFormat32bppPARGB, ByVal 0&, hImgShadow)

    ReDim dBytes(REC.Width * 4 - 1&, REC.Height - 1&)
    
    With bmpData2
        .Scan0Ptr = VarPtr(dBytes(0&, 0&))
        .stride = 4& * REC.Width
    End With
    
    Call GdipBitmapLockBits(hImgShadow, REC, &H4 Or &H1 Or &H2, PixelFormat32bppPARGB, bmpData2)
 
    R = Color And &HFF
    G = (Color \ &H100&) And &HFF
    B = (Color \ &H10000) And &HFF
    
    tAvg = (t2xBlur + 1) * (t2xBlur + 1)    ' how many pixels are being blurred
    
    ReDim vTally(0 To t2xBlur)              ' number of blur columns per pixel
    
    For Y = 0 To Height + t2xBlur - 1     ' loop thru shadow dib
    
        FillMemory vTally(0), (t2xBlur + 1) * 4, 0  ' reset column totals
        
        If Y < t2xBlur Then         ' y does not exist in source
            initYstart = 0          ' use 1st row
        Else
            initYstart = Y - t2xBlur ' start n blur rows above y
        End If
        ' how may source rows can we use for blurring?
        If Y < Height Then initYstop = Y Else initYstop = Height - 1
        
        tAlpha = 0  ' reset alpha sum
        tColumn = 0    ' reset column counter
        
        ' the first n columns will all be zero
        ' only the far right blur column has values; tally them
        For initY = initYstart To initYstop
            tAlpha = tAlpha + srcBytes(3, initY)
        Next
        ' assign the right column value
        vTally(t2xBlur) = tAlpha
        
        For x = 3 To (Width - 2) * 4 - 1 Step 4
            ' loop thru each source pixel's alpha
            
            ' set shadow alpha using blur average
            dBytes(x, Y) = tAlpha \ tAvg
            ' and set shadow color
            Select Case dBytes(x, Y)
            Case 255
                dBytes(x - 1, Y) = R
                dBytes(x - 2, Y) = G
                dBytes(x - 3, Y) = B
            Case 0
            Case Else
                dBytes(x - 1, Y) = R * dBytes(x, Y) \ 255
                dBytes(x - 2, Y) = G * dBytes(x, Y) \ 255
                dBytes(x - 3, Y) = B * dBytes(x, Y) \ 255
            End Select
            ' remove the furthest left column's alpha sum
            tAlpha = tAlpha - vTally(tColumn)
            ' count the next column of alphas
            vTally(tColumn) = 0&
            For initY = initYstart To initYstop
                vTally(tColumn) = vTally(tColumn) + srcBytes(x + 4, initY)
            Next
            ' add the new column's sum to the overall sum
            tAlpha = tAlpha + vTally(tColumn)
            ' set the next column to be recalculated
            tColumn = (tColumn + 1) Mod (t2xBlur + 1)
        Next
        
        ' now to finish blurring from right edge of source
        For x = x To (Width + t2xBlur - 1) * 4 - 1 Step 4
            dBytes(x, Y) = tAlpha \ tAvg
            Select Case dBytes(x, Y)
            Case 255
                dBytes(x - 1, Y) = R
                dBytes(x - 2, Y) = G
                dBytes(x - 3, Y) = B
            Case 0
            Case Else
                dBytes(x - 1, Y) = R * dBytes(x, Y) \ 255
                dBytes(x - 2, Y) = G * dBytes(x, Y) \ 255
                dBytes(x - 3, Y) = B * dBytes(x, Y) \ 255
            End Select
            ' remove this column's alpha sum
            tAlpha = tAlpha - vTally(tColumn)
            ' set next column to be removed
            tColumn = (tColumn + 1) Mod (t2xBlur + 1)
        Next
    Next
 
    Call GdipBitmapUnlockBits(hImage, bmpData1)
    Call GdipBitmapUnlockBits(hImgShadow, bmpData2)
    
    ppBlur = hImgShadow
End Function

Private Function ppCreateBitmapFromStream(bvData() As Byte, lBitmap As Long) As Long
On Error GoTo e
Dim IStream     As IUnknown
    Call CreateStreamOnHGlobal(bvData(0), 0&, IStream)
    If Not IStream Is Nothing Then Call GdipLoadImageFromStream(IStream, lBitmap)
e:
    Set IStream = Nothing
End Function


Private Function ppSaveBmp(ByVal FileName As String, Bmp As Long) As Boolean
Dim eGuid   As GUID
    If Bmp = 0& Then Exit Function
    CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), eGuid
    ppSaveBmp = GdipSaveImageToFile(Bmp, StrConv(FileName, vbUnicode), eGuid, ByVal 0&) = 0&
End Function

'?GDIP
Private Sub ppGdipStart(ByVal StartUp As Boolean)
    If StartUp Then
        If m_token = 0& Then
            Dim gdipSI(3) As Long
            gdipSI(0) = 1&
            Call GdiplusStartup(m_token, gdipSI(0), ByVal 0)
        End If
    Else
        If m_token <> 0 Then Call GdiplusShutdown(m_token): m_token = 0
    End If
End Sub

Public Sub ppCopyAmbient()
On Error GoTo e
Dim oPic As StdPicture
    With UserControl
        Set .Picture = Nothing
        Set oPic = Extender.Container.Image
        .BackColor = Extender.Container.BackColor
        UserControl.PaintPicture oPic, 0, 0, , , Extender.Left, Extender.Top ', Extender.Width * 15, Extender.Height * 15
        Set .Picture = .Image
    End With
    Exit Sub
e:
End Sub

Public Function GetWindowsDPI() As Double
Dim Hdc As Long, lPx  As Double
Const LOGPIXELSX    As Long = 88
    Hdc = GetDC(0)
    lPx = CDbl(GetDeviceCaps(Hdc, LOGPIXELSX))
    ReleaseDC 0, Hdc
    
    If (lPx = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = lPx / 96#
    End If
End Function
Private Function ConvertColor(ByVal Color As Long, ByVal Opacity As Long) As Long
Dim BGRA(0 To 3) As Byte
    OleTranslateColor Color, 0, VarPtr(Color)
  
    BGRA(3) = CByte((Abs(Opacity) / 100) * 255)
    BGRA(0) = ((Color \ &H10000) And &HFF)
    BGRA(1) = ((Color \ &H100) And &HFF)
    BGRA(2) = (Color And &HFF)
    CopyMemory ConvertColor, BGRA(0), 4&
End Function
Private Function GetSafeRound(Angle As Integer, Width As Long, Height As Long) As Integer
    Dim lRet As Integer
    lRet = Angle
    If lRet * 2 > Height Then lRet = Height \ 2
    If lRet * 2 > Width Then lRet = Width \ 2
    GetSafeRound = lRet
End Function
Private Sub SafeRange(Value, Min, Max)
    If Value < Min Then Value = Min
    If Value > Max Then Value = Max
End Sub
