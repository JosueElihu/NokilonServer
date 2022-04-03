VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGDIPDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECTL
  Left    As Long
  Top     As Long
  Width   As Long
  Height  As Long
End Type

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal Hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal Hdc As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As Any, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As Long, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal Hdc As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal Graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal Graphics As Long, ByVal mode As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal Graphics As Long, ByVal PixelOffsetMode As Long) As Long

Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipSetPenMode Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mPenMode As Long) As Long
Private Declare Function GdipDrawLineI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long

Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal ARGB As Long, ByRef brush As Long) As Long
Private Declare Function GdipFillEllipseI Lib "gdiplus" (ByVal Graphics As Long, ByVal brush As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipFillRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipClosePathFigures Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "GdiPlus.dll" (ByRef mRect As RECTL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As WrapMode, ByRef mLineGradient As Long) As Long

Private Declare Function GdipSetClipPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPath As Long, ByVal mCombineMode As Long) As Long
Private Declare Function GdipAddPathEllipseI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long

Private Enum WrapMode
    WrapModeTile = &H0
    WrapModeTileFlipX = &H1
    WrapModeTileFlipy = &H2
    WrapModeTileFlipXY = &H3
    WrapModeClamp = &H4
End Enum

Private Const UnitPixel     As Long = &H2&
Private Const LOGPIXELSY    As Long = 90

Private m_Token     As Long
Private m_Scale     As Single
Private m_Graphic   As Long
Private m_bFlag     As Boolean

Private m_Round(3)  As Single   ' Border Rounds

Private Sub Class_Initialize()
    ManageGdip True
    m_Scale = GetWindowsDPI
End Sub
Private Sub Class_Terminate()
    Graphic = 0
    ManageGdip False
End Sub

Property Let Hdc(Value As Long)
    If m_Graphic <> 0 And m_bFlag = False Then
        GdipDeleteGraphics m_Graphic
        m_Graphic = 0
    End If
    If Value <> 0 Then
        If GdipCreateFromHDC(Value, m_Graphic) = 0 Then
            m_bFlag = False
        Else
            Debug.Print "GDIP Error: GdipCreateFromHDC"
        End If
    End If
End Property

Property Let Graphic(Value As Long)
    If m_Graphic <> 0 And m_bFlag = False Then
        GdipDeleteGraphics m_Graphic
        m_Graphic = 0
    End If
    If Value <> 0 Then
        m_Graphic = Value
        m_bFlag = True
    Else
        m_Graphic = 0
        m_bFlag = False
    End If
End Property


Public Function DrawLine(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, Optional ByVal Color As OLE_COLOR = vbBlack, Optional ByVal Opacity As Integer = 100, Optional ByVal PenWidth As Integer = 1) As Boolean
Dim hPen As Long
    If m_Graphic = 0 Then Exit Function
    GdipSetSmoothingMode m_Graphic, 4 '-> SmoothingModeAntiAlias
    GdipCreatePen1 ConvertColor(Color, Opacity), PenWidth * m_Scale, UnitPixel, hPen
    DrawLine = GdipDrawLineI(m_Graphic, hPen, x1, y1, x2, y2) = 0
    GdipDeletePen hPen
End Function


Public Sub SetBorderRound(Optional Ltop As Long, Optional Rtop As Long, Optional Rbottom As Long, Optional Lbottom As Long)
    m_Round(0) = Ltop
    m_Round(1) = Rtop
    m_Round(2) = Rbottom
    m_Round(3) = Lbottom
End Sub

Public Sub DrawRectangle(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, BackColor As Long, Optional ByVal BackColorAlpha As Integer = 100, Optional BorderSize As Long, Optional BorderColor As Long, Optional BorderAlpha As Integer = 100)
Dim lPath   As Long
Dim lBrush  As Long
Dim lPen    As Long

Dim BCLT    As Integer
Dim BCRT    As Integer
Dim BCBR    As Integer
Dim BCBL    As Integer

    If m_Graphic = 0 Then Exit Sub
    Height = Height - 1 'Antialias pixel
    Width = Width - 1   'Antialias pixel
    
    GdipSetSmoothingMode m_Graphic, 4 '-> SmoothingModeAntiAlias
    Call GdipSetInterpolationMode(m_Graphic, 7&)    '> InterpolationModeHighQualityBicubic
    'GdipTranslateWorldTransform m_Graphic, PosX, PosY, &H1
    'Call GdipSetPixelOffsetMode(m_Graphic, 4&)
    
    Call GdipCreatePath(&H0, lPath)
    
    BCLT = GetSafeRound(m_Round(0) * m_Scale, Width, Height)
    BCRT = GetSafeRound(m_Round(1) * m_Scale, Width, Height)
    BCBR = GetSafeRound(m_Round(2) * m_Scale, Width, Height)
    BCBL = GetSafeRound(m_Round(3) * m_Scale, Width, Height)
    
    GdipAddPathArcI lPath, x, y, BCLT * 2, BCLT * 2, 180, 90
    If BCLT = 0 Then GdipAddPathLineI lPath, x, y, x + Width, y
    
    GdipAddPathArcI lPath, x + Width - BCRT * 2, y, BCRT * 2, BCRT * 2, 270, 90
    If BCRT = 0 Then GdipAddPathLineI lPath, x + Width, y, x + Width, y + Height - BCBR

    GdipAddPathArcI lPath, x + Width - BCBR * 2, y + Height - BCBR * 2, BCBR * 2, BCBR * 2, 0, 90
    If BCBR = 0 Then GdipAddPathLineI lPath, x + Width, y + Height, x + BCBL, y + Height
    
    GdipAddPathArcI lPath, x, y + Height - BCBL * 2, BCBL * 2, BCBL * 2, 90, 90
    If BCBL = 0 Then GdipAddPathLineI lPath, x, y + Height, x, y + BCLT
    
    GdipCreateSolidFill ConvertColor(BackColor, BackColorAlpha), lBrush
    
    GdipFillPath m_Graphic, lBrush, lPath
    GdipClosePathFigures lPath
    
    If BorderSize > 0 Then
        GdipCreatePen1 ConvertColor(BorderColor, BorderAlpha), (BorderSize * m_Scale), UnitPixel, lPen
        'GdipSetPenMode lPen, &H1
        GdipDrawPath m_Graphic, lPen, lPath
        GdipDeletePen lPen
    End If
    
    GdipDeleteBrush lBrush
    GdipDeletePath lPath
    
End Sub

Public Sub DrawRectangleGradient(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, Color1 As Long, Color2 As Long, Optional ByVal AlphaColor1 As Integer = 100, Optional ByVal AlphaColor2 As Integer = 100, Optional BorderSize As Long, Optional BorderColor As Long, Optional BorderAlpha As Integer = 100, Optional GradientAngle As Long)
Dim lPath   As Long
Dim lBrush  As Long
Dim lPen    As Long
Dim RECTL   As RECTL
Dim BCLT    As Integer
Dim BCRT    As Integer
Dim BCBR    As Integer
Dim BCBL    As Integer

    If m_Graphic = 0 Then Exit Sub
    Height = Height - 1 'Antialias pixel
    Width = Width - 1   'Antialias pixel
    
    GdipSetSmoothingMode m_Graphic, 4 '-> SmoothingModeAntiAlias
    Call GdipSetInterpolationMode(m_Graphic, 7&)    '> InterpolationModeHighQualityBicubic
    'GdipTranslateWorldTransform m_Graphic, PosX, PosY, &H1
    'Call GdipSetPixelOffsetMode(m_Graphic, 4&)
    
    Call GdipCreatePath(&H0, lPath)
    
    BCLT = GetSafeRound(m_Round(0) * m_Scale, Width, Height)
    BCRT = GetSafeRound(m_Round(1) * m_Scale, Width, Height)
    BCBR = GetSafeRound(m_Round(2) * m_Scale, Width, Height)
    BCBL = GetSafeRound(m_Round(3) * m_Scale, Width, Height)
    
    GdipAddPathArcI lPath, x, y, BCLT * 2, BCLT * 2, 180, 90
    If BCLT = 0 Then GdipAddPathLineI lPath, x, y, x + Width, y
    
    GdipAddPathArcI lPath, x + Width - BCRT * 2, y, BCRT * 2, BCRT * 2, 270, 90
    If BCRT = 0 Then GdipAddPathLineI lPath, x + Width, y, x + Width, y + Height - BCBR

    GdipAddPathArcI lPath, x + Width - BCBR * 2, y + Height - BCBR * 2, BCBR * 2, BCBR * 2, 0, 90
    If BCBR = 0 Then GdipAddPathLineI lPath, x + Width, y + Height, x + BCBL, y + Height
    
    GdipAddPathArcI lPath, x, y + Height - BCBL * 2, BCBL * 2, BCBL * 2, 90, 90
    If BCBL = 0 Then GdipAddPathLineI lPath, x, y + Height, x, y + BCLT
    
    SetRect RECTL, x, y, Width, Height
    GdipCreateLineBrushFromRectWithAngleI RECTL, ConvertColor(Color1, AlphaColor1), ConvertColor(Color2, AlphaColor2), GradientAngle + 90, 0, WrapModeTileFlipXY, lBrush
    
    GdipFillPath m_Graphic, lBrush, lPath
    GdipClosePathFigures lPath
    
    If BorderSize > 0 Then
        GdipCreatePen1 ConvertColor(BorderColor, BorderAlpha), (BorderSize * m_Scale), UnitPixel, lPen
        'GdipSetPenMode lPen, &H1
        GdipDrawPath m_Graphic, lPen, lPath
        GdipDeletePen lPen
    End If
    
    GdipDeleteBrush lBrush
    GdipDeletePath lPath
    
End Sub

Public Sub DrawCircle(ByVal x As Long, ByVal y As Long, ByVal W As Long, ByVal H As Long, Color As Long, Optional Alpha As Integer = 100)
Dim hPath As Long
Dim hBrush As Long

    If m_Graphic = 0 Then Exit Sub
    W = W - 1   'Antialias pixel
    H = H - 1   'Antialias pixel
    
    Call GdipSetSmoothingMode(m_Graphic, &H4) '->SmoothingModeAntiAlias
    Call GdipCreatePath(&H0, hPath)
    
    GdipAddPathEllipseI hPath, x - (W / 2), y - H / 2, W * 2, H * 2         'Dibuja un Circulo en el path
    GdipCreateSolidFill ConvertColor(Color, Alpha), hBrush
    Call GdipFillEllipseI(m_Graphic, hBrush, x, y, W, H)

    Call GdipDeleteBrush(hBrush)                                                                 'Descarga la brocha
    Call GdipDeletePath(hPath)

End Sub



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

Private Function ConvertColor(ByVal Color As Long, ByVal Opacity As Long) As Long
Dim BGRA(0 To 3) As Byte
    OleTranslateColor Color, 0, VarPtr(Color)
  
    BGRA(3) = CByte((Abs(Opacity) / 100) * 255)
    BGRA(0) = ((Color \ &H10000) And &HFF)
    BGRA(1) = ((Color \ &H100) And &HFF)
    BGRA(2) = (Color And &HFF)
    CopyMemory ConvertColor, BGRA(0), 4&
End Function

Private Function GetWindowsDPI() As Double
Dim lHdc As Long, lPx  As Double
Const LOGPIXELSX    As Long = 88

    lHdc = GetDC(0)
    lPx = CDbl(GetDeviceCaps(lHdc, LOGPIXELSX))
    ReleaseDC 0, lHdc

    If (lPx = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = lPx / 96#
    End If
    
End Function


'?GDIP
Private Function ManageGdip(ByVal StartUp As Boolean) As Boolean
    If StartUp Then
        If m_Token <> 0& Then Exit Function
        Dim gdipSI(3) As Long
        gdipSI(0) = 1&
        ManageGdip = GdiplusStartup(m_Token, gdipSI(0), ByVal 0) = 0
    Else
        If m_Token = 0 Then Exit Function
        Call GdiplusShutdown(m_Token)
        m_Token = 0
    End If
End Function


