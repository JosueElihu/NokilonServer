VERSION 5.00
Begin VB.UserControl JImage 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2115
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   DrawStyle       =   2  'Dot
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   PropertyPages   =   "JImage.ctx":0000
   ScaleHeight     =   150
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   141
   ToolboxBitmap   =   "JImage.ctx":0010
   Windowless      =   -1  'True
End
Attribute VB_Name = "JImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As Long, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As Long, ByRef Image As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Any, ByRef Image As Long) As Long

Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef imageattr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imageattr As Long, ByVal ColorAdjust As Long, ByVal EnableFlag As Boolean, ByRef MatrixColor As COLORMATRIX, ByRef MatrixGray As COLORMATRIX, ByVal Flags As Long) As Long

Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal Image As Long, ByRef Width As Single, ByRef Height As Single) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imageattr As Long) As Long

Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstWidth As Long, ByVal DstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal CallBack As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal Hdc As Long, ByRef Graphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal Graphics As Long, ByVal InterpolationMode As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal Graphics As Long, ByVal PixelOffsetMode As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal BITMAP As Long, ByRef hbmReturn As Long, ByVal Background As Long) As Long
Private Declare Function GdipCreateHICONFromBitmap Lib "gdiplus" (ByVal BITMAP As Long, hbmReturn As Long) As Long

Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long

Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)

Private Const QmInvalid                    As Long = -1&
Private Const QmDefault                    As Long = 0&
Private Const QmLow                        As Long = 1&
Private Const QmHigh                       As Long = 2&

Private Const InterpolationModeNearestNeighbor      As Long = QmHigh + 3
Private Const PixelOffsetModeHalf                   As Long = QmHigh + 2

Private Type COLORMATRIX
    m(0 To 4, 0 To 4)           As Single
End Type

Enum enmJImageScale
    em_None
    em_Stretch
    em_ScaleDown
    em_Scale
    em_ScaleUp
End Enum

Private m_token     As Long
Private m_Bitmap    As Long

Private m_BitmapW     As Single
Private m_BitmapH    As Single

Private m_bvData()  As Byte
Private m_Scale     As enmJImageScale
Private m_Color     As Long
Private m_Alpha     As Long

Private m_bResize   As Boolean

Private Sub UserControl_Initialize()
    ManageGDIP True
End Sub
Private Sub UserControl_InitProperties()
    m_Color = -1
    m_Alpha = 100
    m_Scale = 1
End Sub

Private Sub UserControl_Resize()
    If m_bResize Then Exit Sub
    CheckSize
End Sub

Private Sub UserControl_Terminate()
    ManageGDIP False
End Sub
Private Sub UserControl_Paint()
On Error Resume Next
Dim lW As Long, lH As Long, lT As Long, lL As Long
Dim mColor      As COLORMATRIX
Dim mGray       As COLORMATRIX
Dim hGraphics   As Long
Dim hAttr       As Long

    
    With mColor
        If m_Color <> -1 Then
            Dim R As Byte, G As Byte, B As Byte
            B = ((m_Color \ &H10000) And &HFF)
            G = ((m_Color \ &H100) And &HFF)
            R = (m_Color And &HFF)
            .m(0, 0) = R / 255
            .m(1, 0) = G / 255
            .m(2, 0) = B / 255
            .m(0, 4) = R / 255
            .m(1, 4) = G / 255
            .m(2, 4) = B / 255
        Else
            .m(0, 0) = 1
            .m(1, 1) = 1
            .m(2, 2) = 1
            .m(4, 4) = 1
        End If
        .m(3, 3) = m_Alpha / 100
    End With
    
    With UserControl
        If m_Bitmap <> 0 Then
            ScalePicture m_Scale, m_BitmapW, m_BitmapH, .ScaleWidth, .ScaleHeight, lW, lH, lL, lT
            If GdipCreateFromHDC(.Hdc, hGraphics) = 0 Then

                If lW <> m_BitmapW Or lH <> m_BitmapH Then
                    Call GdipSetInterpolationMode(hGraphics, 7&)  '-> InterpolationModeHighQualityBicubic
                    Call GdipSetPixelOffsetMode(hGraphics, 4&)
                Else
                    Call GdipSetPixelOffsetMode(hGraphics, PixelOffsetModeHalf)
                End If
                
                If GdipCreateImageAttributes(hAttr) = 0 Then
                    Call GdipSetImageAttributesColorMatrix(hAttr, 0&, True, mColor, mGray, 0&)
                    Call GdipDrawImageRectRectI(hGraphics, m_Bitmap, lL, lT, lW, lH, 0, 0, m_BitmapW, m_BitmapH, &H2, hAttr)
                    GdipDisposeImageAttributes hAttr
                End If
                GdipDeleteGraphics hGraphics
            End If
            
        Else
            If Not Ambient.UserMode Then UserControl.Line (0, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), vbHighlight, B
        End If
    End With
    
    
End Sub
Private Sub UserControl_HitTest(x As Single, y As Single, HitResult As Integer)
    HitResult = vbHitResultHit
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_bvData() = .ReadProperty("imgStream", "")
        m_Scale = .ReadProperty("eScale", 0)
        m_Color = .ReadProperty("Color", -1)
        m_Alpha = .ReadProperty("Alpha", 100)
    End With
    If m_Alpha > 100 Then m_Alpha = 100
    If m_Alpha < 0 Then m_Alpha = 0
    
    Call LoadPictureFromStream(m_bvData())
    If Ambient.UserMode Then Erase m_bvData
    
    'UserControl_Resize
    
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "imgStream", m_bvData
        .WriteProperty "eScale", m_Scale
        .WriteProperty "Color", m_Color
        .WriteProperty "Alpha", m_Alpha
    End With
End Sub

Friend Function ppgGetStream() As Byte(): ppgGetStream = m_bvData: End Function
Friend Function ppgSetStream(vData() As Byte)
    m_bvData = vData
    Call LoadPictureFromStream(m_bvData)
    CheckSize
    PropertyChanged "imgStream"
End Function



Property Get ScaleMode() As enmJImageScale: ScaleMode = m_Scale: End Property
Property Let ScaleMode(ByVal Value As enmJImageScale)
    m_Scale = Value
    Call PropertyChanged("eScale")
    Call Me.Refresh
End Property

Property Get hIcon() As Long
    If m_Bitmap = 0 Then Exit Property
    Call GdipCreateHICONFromBitmap(m_Bitmap, hIcon)
End Property
Property Get hBitmap() As Long
Attribute hBitmap.VB_UserMemId = 0
    If m_Bitmap = 0 Then Exit Property
    GdipCreateHBITMAPFromBitmap m_Bitmap, hBitmap, 0 '&HE200B
End Property
Property Get Color() As OLE_COLOR: Color = m_Color: End Property
Property Let Color(ByVal Value As OLE_COLOR)
    m_Color = Value
    Call Me.Refresh
    Call PropertyChanged("Color")
End Property
Property Get Alpha() As Long: Alpha = m_Alpha: End Property
Property Let Alpha(ByVal Value As Long)
    m_Alpha = Value
    If m_Alpha > 100 Then m_Alpha = 100
    If m_Alpha < 0 Then m_Alpha = 0
    Call Me.Refresh
    Call PropertyChanged("Alpha")
End Property

Public Function GetStream() As Byte()
     GetStream = m_bvData
End Function


Public Sub Render(lHdc As Long, ByVal hGraphic As Long, ByVal x As Long, ByVal y As Long, Optional ByVal H As Long, Optional ByVal W As Long, Optional ByVal SrcX As Long, Optional ByVal SrcY As Long, Optional ByVal SrcW As Long, Optional ByVal SrcH As Long)
Dim mbFlag As Boolean
    If m_Bitmap = 0 Then Exit Sub
    If W = 0 Then W = m_BitmapW
    If H = 0 Then H = m_BitmapH
    If SrcW = 0 Then SrcW = m_BitmapW
    If SrcH = 0 Then SrcH = m_BitmapH
    If hGraphic = 0 Then GdipCreateFromHDC lHdc, hGraphic: mbFlag = True
    GdipSetInterpolationMode hGraphic, 7&
    GdipSetPixelOffsetMode hGraphic, 4&
    GdipDrawImageRectRectI hGraphic, m_Bitmap, x, y, W, H, SrcX, SrcY, SrcW, SrcH, &H2, 0&, 0&, 0&
    If mbFlag Then GdipDeleteGraphics hGraphic
End Sub


'?GDIP
Private Sub ManageGDIP(ByVal Startup As Boolean)
    If Startup Then
        If m_token <> 0& Then Exit Sub
        Dim gdipSI(3) As Long
        gdipSI(0) = 1&
        Call GdiplusStartup(m_token, gdipSI(0), ByVal 0)
    Else
        If m_token = 0 Then Exit Sub
        Call GdiplusShutdown(m_token)
        m_token = 0
    End If
End Sub
Public Sub CleanUp()
    If m_Bitmap Then
        Call GdipDisposeImage(m_Bitmap)
        m_Bitmap = 0
        m_BitmapW = 0
        m_BitmapH = 0
    End If
End Sub
Public Function LoadPictureFromFile(ByVal FileName As String) As Boolean
    Call CleanUp
    If GdipLoadImageFromFile(StrPtr(FileName), m_Bitmap) Then
        GdipGetImageDimension m_Bitmap, m_BitmapW, m_BitmapH
        LoadPictureFromFile = True
    End If
End Function
Public Function LoadPictureFromStream(bvData() As Byte) As Boolean
On Error GoTo e
Dim IStream   As IUnknown

    CleanUp
    Call CreateStreamOnHGlobal(bvData(0), 0&, IStream)
    If Not IStream Is Nothing Then
        If GdipLoadImageFromStream(IStream, m_Bitmap) = 0 Then
            GdipGetImageDimension m_Bitmap, m_BitmapW, m_BitmapH
            Call CheckSize
            LoadPictureFromStream = True
            Me.Refresh
        End If
    End If
    Set IStream = Nothing
e:
End Function

Public Sub Refresh()
    CheckSize
    UserControl.Refresh
End Sub

Private Sub CheckSize()
    m_bResize = True
    If m_Scale = em_None And m_Bitmap Then
        UserControl.Height = m_BitmapH * 15
        UserControl.Width = m_BitmapW * 15
    End If
    m_bResize = False
End Sub



Private Function ScalePicture( _
       ByVal eScaleMode As enmJImageScale, _
       ByVal lSrcWidth As Long, _
       ByVal lSrcHeight As Long, _
       ByVal lDstWidth As Long, _
       ByVal lDstHeight As Long, _
       ByRef lNewWidth As Long, _
       ByRef lNewHeight As Long, _
       ByRef lNewLeft As Long, _
       ByRef lNewTop As Long)

    Dim dHRatio As Double
    Dim dVRatio As Double
    Dim dRatio  As Double
    
    dHRatio = lSrcWidth / lDstWidth
    dVRatio = lSrcHeight / lDstHeight
     
    Select Case eScaleMode
        Case em_None
            lNewWidth = lSrcWidth
            lNewHeight = lSrcHeight
        Case em_Stretch
            lNewWidth = lDstWidth
            lNewHeight = lDstHeight
        Case em_ScaleDown
            If dHRatio > 1 Or dVRatio > 1 Then
                If dHRatio > dVRatio Then
                    dRatio = dHRatio
                Else
                    dRatio = dVRatio
                End If
            Else
                lNewWidth = lSrcWidth
                lNewHeight = lSrcHeight
            End If
        Case em_Scale
            If dHRatio > dVRatio Then
                dRatio = dHRatio
            Else
                dRatio = dVRatio
            End If
        Case em_ScaleUp
            If dHRatio < dVRatio Then
                dRatio = dHRatio
            Else
                dRatio = dVRatio
            End If
    End Select
    
    If Not dRatio = 0 Then
        lNewWidth = lSrcWidth / dRatio
        lNewHeight = lSrcHeight / dRatio
    End If
    
    lNewLeft = (lDstWidth - lNewWidth) / 2
    lNewTop = (lDstHeight - lNewHeight) / 2
End Function

