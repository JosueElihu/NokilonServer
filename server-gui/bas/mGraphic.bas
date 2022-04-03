Attribute VB_Name = "mGraphic"
Option Explicit

Private Type POINTAPI
  x       As Long
  y       As Long
End Type
Private Type RECT
  Left    As Long
  Top     As Long
  Right   As Long
  Bottom  As Long
End Type
Private Type TRIVERTEX
  x       As Long
  y       As Long
  Red     As Integer
  Green   As Integer
  Blue    As Integer
  Alpha   As Integer
End Type

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal Hdc As Long, ByVal hObject As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal Hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal Hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal Hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal Hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal Hdc As Long, ByVal nIndex As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal Hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal Hdc As Long) As Long
Private Declare Function CreateIconFromResourceEx Lib "user32" (ByRef presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function SetRect Lib "user32" (ByRef lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long

Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal mPen As Long) As Long
Private Declare Function GdipDrawArc Lib "gdiplus" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal Hdc As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal Graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal Graphics As Long, ByVal InterpolationMode As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal Graphics As Long, ByVal PixelOffsetMode As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Function GradientFill Lib "msimg32" (ByVal Hdc As Long, ByRef Vertex As TRIVERTEX, ByVal nVertex As Long, ByRef Mesh As POINTAPI, ByVal nMesh As Long, ByVal mode As Long) As Long

Private Declare Function OleTranslateColor2 Lib "olepro32" Alias "OleTranslateColor" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long

Public mdpi_      As Single


Public Sub PutIcon32Bit(ByVal hWnd As Long, ResIcon As Variant)
Dim hIcon As Long

    hIcon = LoadImage(App.hInstance, ResIcon, 1, 32 * mdpi_, 32 * mdpi_, &H8000& Or &H1000)
    If hIcon Then DestroyIcon SendMessage(hWnd, &H80, 1, ByVal hIcon)
    hIcon = LoadImage(App.hInstance, ResIcon, 1, 16 * mdpi_, 16 * mdpi_, &H8000& Or &H1000)
    If hIcon Then DestroyIcon SendMessage(hWnd, &H80, 0, ByVal hIcon)
    
    ' Const ICON_BIG As Long = 1
    ' Const ICON_SMALL As Long = 0
    ' Const WM_SETICON As Long = &H80

'    hIcon = LoadIcon(App.hInstance, ResIcon)
'    Call SendMessage(hWnd, &H80, 1, ByVal hIcon)
'    Call SendMessage(hWnd, &H80, 0, ByVal hIcon)
'    Call DestroyIcon(hIcon)
End Sub
Public Sub SetIconStream(hWnd As Long, bvData() As Byte)
On Error GoTo e_
Dim hIcon As Long

    
    'hIcon = CreateIconFromResourceEx(bvData(0), UBound(bvData) + 1&, 1&, &H30000, 0&, 0&, 0&)
    'If hIcon Then DestroyIcon SendMessage(hwnd, &H80, 1, ByVal hIcon) 'BIG
    
    hIcon = CreateIconFromResourceEx(bvData(LBound(bvData)), UBound(bvData) + 1&, 1&, &H30000, 0&, 0&, 0&)
    If hIcon Then DestroyIcon SendMessage(hWnd, &H80, 0, ByVal hIcon) 'SMALL
e_:
End Sub
Public Function IconFromStream(bvData() As Byte, Optional ByVal W As Long, Optional ByVal H As Long) As Long
On Error GoTo e
    IconFromStream = CreateIconFromResourceEx(bvData(LBound(bvData)), UBound(bvData) + 1&, 1&, &H30000, W&, H&, 0&)
e:
End Function

Public Sub FillBack(dvc As Long, Color As Long, x As Long, y As Long, W As Long, H As Long)
Dim hBrush  As Long
Dim Rct     As RECT

    SetRect Rct, x, y, x + W, y + H
    hBrush = CreateSolidBrush(Color)
    Call FillRect(dvc, Rct, hBrush)
    Call DeleteObject(hBrush)
End Sub

Public Sub DrawRectBorder(dvc As Long, Color As Long, x As Long, y As Long, W As Long, H As Long, Optional PenSize As Long = 1)
Dim hPen As Long
Dim Px1  As Long
    
    Px1 = PenSize \ 2
    OleTranslateColor2 Color, 0, Color
    hPen = CreatePen(0, PenSize, Color)
    Call SelectObject(dvc, hPen)
    RoundRect dvc, x + Px1, y + Px1, x + (W - Px1), y + (H - Px1), 0, 0
    DeleteObject hPen
End Sub
Public Sub DrawLine(dvc As Long, Color As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, Optional PenSize As Long = 1)
Dim PT      As POINTAPI
Dim hPen    As Long
Dim hPenOld As Long

    OleTranslateColor2 Color, 0, Color
    hPen = CreatePen(0, 1 * mdpi_, Color)
    hPenOld = SelectObject(dvc, hPen)
    Call MoveToEx(dvc, x1, y1, PT)
    Call LineTo(dvc, x2, y2)
    Call SelectObject(dvc, hPenOld)
    Call DeleteObject(hPen)
End Sub

Public Sub DrawHeader(dvc As Long, x As Long, Width As Long, Height As Long, oPic As PictureBox)
Dim lPx As Long

    BitBlt dvc, x, 0, 3 * mdpi_, Height, oPic.Hdc, 0, 0, vbSrcCopy '[LEFT]
    StretchBlt dvc, x + (3 * mdpi_), 0, Width - (6 * mdpi_), Height, oPic.Hdc, 3 * mdpi_, 0, 3, oPic.ScaleHeight, vbSrcCopy
    BitBlt dvc, x + (Width - (3 * mdpi_)), 0, 3 * mdpi_, Height, oPic.Hdc, oPic.ScaleWidth - (3 * mdpi_), 0, vbSrcCopy '[RIGHT]
    
End Sub




Public Sub DrawCircularProgress(Hdc As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Percent As Single, ForeColor As Long, ByVal hGraphics As Long, Optional Marquee As Boolean)
Dim hPen    As Long
Dim mbFlag  As Boolean

    If hGraphics = 0 Then
        GdipCreateFromHDC Hdc, hGraphics
        mbFlag = True
    End If
    
    GdipSetSmoothingMode hGraphics, 4&
    Call GdipSetInterpolationMode(hGraphics, 7&)  '-> InterpolationModeHighQualityBicubic
    Call GdipSetPixelOffsetMode(hGraphics, 4&)
    
    GdipCreatePen1 ConvertColor(&HD9D9D9, 100), 2 * mdpi_, &H2&, hPen
    GdipDrawArc hGraphics, hPen, x, y, Width, Height, 0, 360
    GdipDeletePen hPen
    GdipCreatePen1 ConvertColor(ForeColor, 100), 2 * mdpi_, &H2&, hPen
    Percent = Percent / 100
    If Not Marquee Then
        GdipDrawArc hGraphics, hPen, x, y, Width, Height, -90, 360 * Percent
    Else
        GdipDrawArc hGraphics, hPen, x, y, Width, Height, -90 + (360 * Percent), 72   '36
    End If
    GdipDeletePen hPen
    
    If mbFlag Then
        GdipDeleteGraphics hGraphics
    End If
End Sub

Public Sub DrawWaiting(Hdc As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal lSteep As Long, ForeColor As Long)
Dim i      As Long
Dim hBrush As Long
Dim Rct    As RECT
Dim Xx     As Long

     Xx = (Width / 3)
     For i = 0 To 2
        
        If i = lSteep Or lSteep = 4 Then
            hBrush = CreateSolidBrush(ForeColor)
        Else
            hBrush = CreateSolidBrush(&HD9D9D9)
        End If

        SetRect Rct, x, y, x + (Xx - (1 * mdpi_)), y + Height
        FillRect Hdc, Rct, hBrush
        DeleteObject hBrush
        
        x = x + Xx
     Next
End Sub

Public Sub FillGradient(lHdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal Color1 As OLE_COLOR, ByVal Color2 As OLE_COLOR, Optional ByVal FillVertical As Boolean)
Dim TRVRT(1)   As TRIVERTEX
Dim PT         As POINTAPI

    OleTranslateColor2 Color1, 0, Color1
    OleTranslateColor2 Color2, 0, Color2
    With TRVRT(0)
        .x = x1
        .y = y1
        .Red = LongToSignedShort(Color1 And &HFF& * 256)
        .Green = LongToSignedShort(((Color1 And &HFF00&) / &H100&) * 256)
        .Blue = LongToSignedShort(((Color1 And &HFF0000) / &H10000) * 256)
        .Alpha = 0
    End With
    
    With TRVRT(1)
        .x = x2
        .y = y2
        .Red = LongToSignedShort((Color2 And &HFF&) * 256)
        .Green = LongToSignedShort(((Color2 And &HFF00&) / &H100&) * 256)
        .Blue = LongToSignedShort(((Color2 And &HFF0000) / &H10000) * 256)
        .Alpha = 0
    End With
    PT.y = 1
    GradientFill lHdc, TRVRT(0), 2, PT, 1, Abs(FillVertical)
End Sub

Public Function WindowsDPI() As Double
Dim Hdc  As Long
Dim lPx  As Double
Const LOGPIXELSX As Long = 88

    Hdc = GetDC(0)
    lPx = CDbl(GetDeviceCaps(Hdc, LOGPIXELSX))
    ReleaseDC 0, Hdc
    If (lPx = 0) Then WindowsDPI = 1# Else WindowsDPI = lPx / 96#
    
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
Private Function LongToSignedShort(dwUnsigned As Long) As Integer
   If dwUnsigned < 32768 Then
      LongToSignedShort = CInt(dwUnsigned)
   Else
      LongToSignedShort = CInt(dwUnsigned - &H10000)
   End If
End Function
