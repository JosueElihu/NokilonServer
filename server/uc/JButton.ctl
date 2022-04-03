VERSION 5.00
Begin VB.UserControl JButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3465
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   PropertyPages   =   "JButton.ctx":0000
   ScaleHeight     =   54
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   231
   ToolboxBitmap   =   "JButton.ctx":0011
End
Attribute VB_Name = "JButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : JButton 2.5.1
'    Autor      : J. Elihu
'--------------------------------------------------------------------------------
'    Description: VB6 Standar button replacement
'    Req        : cSubClass
'--------------------------------------------------------------------------------

Option Explicit

'?Mouse
Private Declare Function GetCapture Lib "user32.dll" () As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SetCapture Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32.dll" (ByRef lpEventTrack As Long) As Long ' Win98 or later

'/DPI
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal Hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal Hdc As Long) As Long

'/Theme
Private Declare Function DrawFocusRect Lib "user32" (ByVal Hdc As Long, lpRect As RECT) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackGround Lib "uxtheme.dll" Alias "DrawThemeBackground" (ByVal hTheme As Long, ByVal lHdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As Any) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal Hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long

'/Text
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal Hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal Hdc As Long, ByVal crColor As Long) As Long

'/Line
Private Declare Function MoveToEx Lib "gdi32" (ByVal Hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal Hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal Hdc As Long, lpPoint As Any, ByVal nCount As Long) As Long

'/Rect
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CopyRect Lib "user32.dll" (ByRef lpDestRect As RECT, ByRef lpSourceRect As RECT) As Long

'?GDI
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As Long, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Any, ByRef Image As Long) As Long
Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal Image As Long, ByRef Width As Single, ByRef Height As Single) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal Hdc As Long, ByRef Graphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal Graphics As Long, ByVal InterpolationMode As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal Graphics As Long, ByVal PixelOffsetMode As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstWidth As Long, ByVal DstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal Graphics As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef imageattr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imageattr As Long, ByVal ColorAdjust As Long, ByVal EnableFlag As Boolean, ByRef MatrixColor As COLORMATRIX, MatrixGray As Any, ByVal Flags As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imageattr As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal Format As Long, ByRef Scan0 As Any, ByRef BITMAP As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipDrawImageRect Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long


Private Type COLORMATRIX
    m(0 To 4, 0 To 4) As Single
End Type

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

Enum JButtonType
    NormalBtn
    CheckBtn
    OptionBtn
End Enum
Enum JButtonImageAlignment
    ImageInTopBtn
    ImageInLeftBtn
    ImageInRightBtn
    ImageInBehindBtn
End Enum
Enum JButtonState
    lNormalBtn
    lHotBtn
    lDownBtn
    lDisabledBtn
    lFocusedBtn
End Enum
Enum JButtonArrowStyle
    ToDownArrow
    ToUpArrow
    ToLeftArrow
    ToRightArrow
End Enum

Event Click()
Event DblClick()
Event MouseLeave()
Event MouseEnter()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event ValueChanged(ByRef Value As Boolean, ByVal EventCode As Integer)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)

Event BeforePaint(Hdc As Long, hGraphic As Long, ByVal Width As Long, ByVal Height As Long, ByRef Evt As JButtonState, Cancel As Boolean)
Event AffterPaint(Hdc As Long, hGraphic As Long, ByVal Width As Long, ByVal Height As Long, ByVal Evt As JButtonState)

Event OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Event OLEDragOver(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Event OLECompleteDrag(Effect As Long)

Private cSubClass   As cSubClass

Private m_Token   As Long
Private m_Skin     As Long
Private m_SkinW    As Single
Private m_SkinH    As Single

Private m_Bitmap   As Long
Private m_BmpSrcW  As Single
Private m_BmpSrcH  As Single
Private m_bvData() As Byte

Private m_Text     As String
Private m_SkinRes  As String
Private m_Margins  As String
Private m_bValue   As Boolean
Private m_algnText As AlignmentConstants
Private m_algnBmp  As JButtonImageAlignment
Private m_BtnType  As JButtonType
Private m_beDown   As Boolean
Private m_bNoBkgnd As Boolean
Private m_lColor(4) As Long
Private m_aImage    As Boolean
Private m_BmpSize   As String
Private m_BmpColor  As Long
Private m_bmpSpc    As Long

Private m_bTrack    As Boolean
Private m_bSpaceDwn As Boolean
Private m_lKey      As Integer
Private m_bFocus    As Boolean
Private m_lState    As JButtonState
Private m_TRct      As RECT
Private m_IRct      As RECT
Private m_msTrack(3) As Long

Private m_lButton    As Integer
Private m_lBtnDown   As Boolean
Private m_lMouseDown As Boolean
Private dpi_         As Single


'/ User Control Rutines      ---------------
Private Sub UserControl_Initialize()
    ManageGdip True
    dpi_ = GetWindowsDPI
End Sub
Private Sub UserControl_Terminate()
    ppUnloadBitmap m_Skin
    ppUnloadBitmap m_Bitmap
    ManageGdip False
    
    Set cSubClass = Nothing
End Sub
Private Sub UserControl_InitProperties()
Dim i As Integer

    m_Text = Extender.Name
    m_Margins = "8,8"
    m_beDown = True
    m_algnText = vbCenter
    m_algnBmp = ImageInLeftBtn
    
    For i = 0 To 4
        m_lColor(i) = -1
    Next
    m_BmpSize = "0x0"
    m_BmpColor = -1
    m_bmpSpc = 3
    
End Sub
Private Sub UserControl_Click()
    If m_lButton <> 1 Or m_bSpaceDwn = True Then Exit Sub
    '/ButtonCheck || ButtonOption
    Select Case m_BtnType
        Case 1: Value = Not Value
        Case 2: If Not Value Then Value = True
    End Select
    RaiseEvent Click
End Sub
Private Sub UserControl_DblClick()
    If m_lButton = 1 Then
        pvDraw 2, True
        RaiseEvent DblClick
    End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    RaiseEvent KeyDown(KeyCode, Shift)
    m_lKey = KeyCode
    Select Case KeyCode
        Case 32 '\Spacio
                If (Not Shift = 4) Then ' vbAltMask
                    m_lBtnDown = True
                    m_lButton = 1
                    m_bSpaceDwn = True
                    If m_lState <> 2 Then pvDraw 2
                    If (Not GetCapture = hWnd) Then SetCapture hWnd ' Restrict user from selecting other
                End If
        Case 39, 40: SendKeys "{Tab}" 'Right & Down arrows
        Case 37, 38: SendKeys "+{Tab}" 'Left & Up arrows
    End Select
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
    If (KeyCode = 32) And (m_lKey = 32) Then
        
        m_lBtnDown = m_lMouseDown
        m_bSpaceDwn = False
        
        If pMouseOnButton Then pvDraw IIf(m_lButton = 1, 2, 1) Else pvDraw 0
        'If (m_lButton = 1) Then
            'If (Not GetCapture = hWnd) Then SetCapture hWnd
        'Else
            If (GetCapture = hWnd) Then ReleaseCapture
        'End If
        If Not m_lBtnDown Then UserControl_Click

    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
    m_lButton = Button
    If Button = 1 Then
        m_lMouseDown = True
        m_lBtnDown = True
        pvDraw 2
    End If
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not m_bTrack Then
        TrackMouseEvent m_msTrack(0)
        RaiseEvent MouseEnter
        m_bTrack = True
    End If
    RaiseEvent MouseMove(Button, Shift, x, y)
    If pMouseOnButton Then
        pvDraw IIf(Button = 1 Or m_bSpaceDwn, 2, 1)
    Else
        m_lButton = 0
        pvDraw 0
    End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
    If Button = 1 Then
        m_lMouseDown = False
        m_lBtnDown = m_bSpaceDwn
        pvDraw 1
    End If
End Sub
Private Sub UserControl_Resize()
    pvUpdateRects
    If m_aImage Then pvSetTrans
    pvDraw m_lState, True
End Sub
Private Sub UserControl_Show()
    If m_aImage Then pvSetTrans
    pvDraw m_lState, True
End Sub


'/ Control Properties --------------------------------
Property Get hWnd() As Long: hWnd = UserControl.hWnd: End Property
Property Get BackColor() As OLE_COLOR: BackColor = UserControl.BackColor: End Property
Property Let BackColor(ByVal Value As OLE_COLOR)
    UserControl.BackColor = Value
    pvDraw m_lState, True
    PropertyChanged "BackColor"
End Property
Property Get Text() As String: Text = m_Text: End Property
Property Let Text(ByVal Text As String)
    m_Text = Text
    PropertyChanged ("Text")
    Call pvUpdateRects: pvDraw m_lState, True
End Property
Public Property Get SkinRes() As String: SkinRes = m_SkinRes: End Property
Public Property Let SkinRes(ByVal StrValue As String)
    m_SkinRes = StrValue
    Call ppCreateSkin
    pvDraw 0, True
    PropertyChanged "SkinRes"
End Property
Property Get Enabled() As Boolean: Enabled = UserControl.Enabled: End Property
Property Let Enabled(ByVal Value As Boolean)
    UserControl.Enabled() = Value
    pvDraw 0, True
    PropertyChanged ("Enabled")
End Property
Public Property Get Font() As Font: Set Font = UserControl.Font: End Property
Public Property Set Font(ByRef oFont As Font)
    Set UserControl.Font = oFont
    PropertyChanged ("Font")
    Call pvUpdateRects: pvDraw m_lState, True
End Property
Public Property Get TextAlignment() As AlignmentConstants: TextAlignment = m_algnText: End Property
Public Property Let TextAlignment(ByVal Value As AlignmentConstants)
    If (Value = 0 Or Value = 1) And (m_algnBmp = 0 Or m_algnBmp = 3) Then m_algnBmp = 1
    '0 Left '1 Right '2 Center
    m_algnText = Value
    Call PropertyChanged("TextAlign")
    Call pvUpdateRects
    Call pvDraw(m_lState, True)
End Property
Public Property Get ImageAlignment() As JButtonImageAlignment: ImageAlignment = m_algnBmp: End Property
Public Property Let ImageAlignment(ByVal Value As JButtonImageAlignment)
    If (Value = ImageInBehindBtn Or Value = ImageInTopBtn) And Not (m_algnText = 2) Then Exit Property
    m_algnBmp = Value
    Call pvUpdateRects
    Call pvDraw(m_lState, True)
    PropertyChanged ("BitmapAlign")
End Property
Public Property Get ButtonMargins() As String: ButtonMargins = m_Margins: End Property
Public Property Let ButtonMargins(ByVal Value As String)
On Error Resume Next
Dim Lm() As String
Dim i As Integer
    Lm = Split(Value, ",", 2)
    If UBound(Lm) < 1 Then ReDim Preserve Lm(1)
    For i = 0 To UBound(Lm)
        Lm(i) = Val(Lm(i))
    Next
    m_Margins = Join(Lm, ",")
    Call pvUpdateRects: pvDraw m_lState, True
    PropertyChanged "Margins"
End Property
Public Property Get ButtonType() As JButtonType: ButtonType = m_BtnType: End Property
Public Property Let ButtonType(ByVal Value As JButtonType)
    m_BtnType = Value
    Call PropertyChanged("ButtonType")
    If m_BtnType = 0 And m_bValue Then Value = False
End Property
Public Property Get Value() As Boolean: Value = m_bValue: End Property
Public Property Let Value(ByVal Value As Boolean)
On Error Resume Next

    If m_BtnType = 0 And Value Then Exit Property
    If Value = m_bValue Then Exit Property
    
    RaiseEvent ValueChanged(Value, 0)
    m_bValue = Value
    If m_BtnType = 2 And Value Then pUpdateOptionButtons
    
    Call PropertyChanged("Value")
    pvDraw 0, True
End Property
Property Get DownEffect() As Boolean: DownEffect = m_beDown: End Property
Property Let DownEffect(ByVal oDownEffect As Boolean)
     m_beDown = oDownEffect
     PropertyChanged "EDown"
End Property
'/ ForeColor
Property Get ForeNormal() As OLE_COLOR: ForeNormal = m_lColor(0): End Property
Property Let ForeNormal(ByVal Value As OLE_COLOR)
    m_lColor(0) = Value
    pvDraw m_lState, True
    PropertyChanged "Fore0"
End Property
Property Get ForeHot() As OLE_COLOR: ForeHot = m_lColor(1): End Property
Property Let ForeHot(ByVal Value As OLE_COLOR)
    m_lColor(1) = Value
    pvDraw m_lState, True
    PropertyChanged "Fore1"
End Property
Property Get ForePressed() As OLE_COLOR: ForePressed = m_lColor(2): End Property
Property Let ForePressed(ByVal Value As OLE_COLOR)
    m_lColor(2) = Value
    pvDraw m_lState, True
    PropertyChanged "Fore2"
End Property
Property Get ForeDisabled() As OLE_COLOR: ForeDisabled = m_lColor(3): End Property
Property Let ForeDisabled(ByVal Value As OLE_COLOR)
    m_lColor(3) = Value
    pvDraw m_lState, True
    PropertyChanged "Fore3"
End Property
Property Get ForeFocused() As OLE_COLOR: ForeFocused = m_lColor(4): End Property
Property Let ForeFocused(ByVal Value As OLE_COLOR)
    m_lColor(4) = Value
    pvDraw m_lState, True
    PropertyChanged "Fore4"
End Property
Property Get AmbientImage() As Boolean: AmbientImage = m_aImage: End Property
Property Let AmbientImage(Value As Boolean)
    m_aImage = Value
    pvSetTrans
    pvDraw m_lState, True
    PropertyChanged "AmbientImage"
End Property
Property Get Background() As Boolean: Background = Not m_bNoBkgnd: End Property
Property Let Background(ByVal Value As Boolean)
    m_bNoBkgnd = Not Value
    pvDraw m_lState, True
    PropertyChanged "NoBkgnd"
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
    Call pvUpdateRects: pvDraw m_lState, True
    PropertyChanged "BitmapResize"
End Property
Property Get ImageColorize() As OLE_COLOR: ImageColorize = m_BmpColor: End Property
Property Let ImageColorize(ByVal Value As OLE_COLOR)
    m_BmpColor = Value
    pvDraw m_lState, True
    PropertyChanged "BitmapColor"
End Property
Property Get ImageSpace() As Long: ImageSpace = m_bmpSpc: End Property
Property Let ImageSpace(ByVal Value As Long)
    m_bmpSpc = Value
    PropertyChanged "BitmapSpace"
    Call pvUpdateRects: pvDraw m_lState, True
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
Dim i As Integer

    With PropBag
        m_SkinRes = .ReadProperty("SkinRes", "")
        m_Text = .ReadProperty("Text", "")
        m_algnText = .ReadProperty("TextAlign", 0)
        m_algnBmp = .ReadProperty("BitmapAlign", 0)
        m_Margins = .ReadProperty("Margins", "8,8")
        m_BtnType = .ReadProperty("ButtonType", 0)
        m_bValue = .ReadProperty("Value", False)
        m_beDown = .ReadProperty("EDown", True)
        m_aImage = .ReadProperty("AmbientImage", False)
        m_BmpSize = .ReadProperty("BitmapResize", "0x0")
        m_BmpColor = .ReadProperty("BitmapColor", -1)
        m_bmpSpc = .ReadProperty("BitmapSpace", 3)
        m_bNoBkgnd = .ReadProperty("NoBkgnd", False)

        UserControl.Enabled() = .ReadProperty("Enabled", True)
        UserControl.BackColor() = .ReadProperty("BackColor", vbButtonFace)
        Set UserControl.Font = .ReadProperty("Font", UserControl.Font)
        
        For i = 0 To 4
            m_lColor(i) = .ReadProperty("Fore" & i, -1)
        Next
        
        If Not Ambient.UserMode Then
            m_bvData() = .ReadProperty("Image", "")
            ppCreateBitmapFromStream m_bvData, m_Bitmap
        Else
           ppCreateBitmapFromStream .ReadProperty("Image", ""), m_Bitmap
        End If
        If m_Bitmap Then GdipGetImageDimension m_Bitmap, m_BmpSrcW, m_BmpSrcH
    End With
    
    If Ambient.UserMode Then
        
        m_msTrack(0) = 16&
        m_msTrack(1) = &H2
        m_msTrack(2) = Me.hWnd
         
        Set cSubClass = New cSubClass
        With cSubClass
            .Subclass hWnd, , , Me
            .AddMsg hWnd, WM_MOUSELEAVE, MSG_AFTER
            .AddMsg hWnd, WM_SETFOCUS, MSG_AFTER
            .AddMsg hWnd, WM_KILLFOCUS, MSG_AFTER
        End With
        'SetSubclassing UserControl.hwnd
    End If
    ppCreateSkin
    pvUpdateRects
    pvDraw lNormalBtn, True
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim i As Integer
    With PropBag
        .WriteProperty "SkinRes", m_SkinRes
        .WriteProperty "Text", m_Text, ""
        .WriteProperty "TextAlign", m_algnText
        .WriteProperty "BitmapAlign", m_algnBmp
        .WriteProperty "Margins", m_Margins
        .WriteProperty "ButtonType", m_BtnType
        .WriteProperty "Value", m_bValue
        .WriteProperty "EDown", m_beDown
        .WriteProperty "AmbientImage", m_aImage
        .WriteProperty "BitmapResize", m_BmpSize
        .WriteProperty "BitmapColor", m_BmpColor
        .WriteProperty "BitmapSpace", m_bmpSpc
        .WriteProperty "NoBkgnd", m_bNoBkgnd
        
        For i = 0 To 4
            .WriteProperty "Fore" & i, m_lColor(i)
        Next
        
        .WriteProperty "Enabled", UserControl.Enabled
        .WriteProperty "BackColor", UserControl.BackColor
        .WriteProperty "Font", UserControl.Font
        
        .WriteProperty "Image", m_bvData
        
    End With
End Sub

'?Public Subs   -------------------------

Public Sub RenderText(Text As String, Optional ByVal Fore As Long = -1, Optional Font As StdFont, Optional ByVal x1 As Long, Optional ByVal y1 As Long, _
                        Optional ByVal x2 As Long, Optional ByVal y2 As Long, Optional ByVal lflags As Long = -1)
Dim TRct As RECT
Dim lColor As Long
Dim oldFont As StdFont


    If Trim(Text) = "" Then Exit Sub
    With UserControl
        lColor = .ForeColor
        If Fore <> -1 Then .ForeColor = Fore
        Set oldFont = .Font
        If Not Font Is Nothing Then Set .Font = Font
        If x2 = 0 Then x2 = .ScaleWidth
        If y2 = 0 Then y2 = .ScaleHeight
        'Left
        If lflags = -1 Then lflags = &H0& Or &H40000
        SetRect TRct, x1, y1, x2, y2
        DrawText .Hdc, Text, Len(Text), TRct, lflags
        If Fore <> -1 Then .ForeColor = lColor
        If Not Font Is Nothing Then Set .Font = oldFont
    End With
End Sub
Public Sub RenderLine(Optional ByVal Fore As Long = -1, Optional ByVal StyleLine As DrawStyleConstants, Optional ByVal lWidth As Long = 1, _
                Optional x1 As Long, Optional y1 As Long, Optional ByVal x2 As Long = -1, Optional ByVal y2 As Long = -1)

Dim PT As POINTAPI
Dim oldFore As Long
Dim oldWidth As Integer
Dim oldStyle As Integer
    With UserControl
        
        oldFore = .ForeColor
        oldWidth = .DrawWidth
        oldStyle = .DrawStyle
        If Fore <> -1 Then .ForeColor = Fore
        .DrawStyle = StyleLine
        .DrawWidth = lWidth * dpi_
        
        If x2 = -1 Then x2 = .ScaleWidth - (pvGetvalue(0) * dpi_)
        If y2 = -1 Then y2 = y1
        
        MoveToEx .Hdc, x1, y1, PT
        LineTo .Hdc, x2, y2
        
        If .ForeColor <> oldFore Then .ForeColor = oldFore
        If .DrawWidth <> oldWidth Then .DrawWidth = oldWidth
        If .DrawStyle <> oldStyle Then .DrawStyle = oldStyle
    End With
End Sub
Public Sub RenderSquare(Optional ByVal Fore As Long = -1, Optional ByVal StyleLine As DrawStyleConstants, Optional ByVal lWidth As Long = 1, Optional x1 As Long, Optional y1 As Long, _
                        Optional ByVal x2 As Long, Optional ByVal y2 As Long, Optional bFill As Boolean)
Dim oldFore As Long
Dim oldWidth As Integer
Dim oldStyle As Integer

    With UserControl
        
        oldWidth = .DrawWidth
        oldStyle = .DrawStyle
        If Fore <> -1 Then oldFore = Fore Else oldFore = .ForeColor
        .DrawStyle = StyleLine
        .DrawWidth = lWidth
        
        If x2 = 0 Then x2 = .ScaleWidth
        If y2 = 0 Then y2 = y1
        
        If bFill Then
            Line (x1, y1)-(x2, y2), oldFore, BF
        Else
            Line (x1, y1)-(x2, y2), oldFore, B
        End If
    
        If .DrawWidth <> oldWidth Then .DrawWidth = oldWidth
        If .DrawStyle <> oldStyle Then .DrawStyle = oldStyle
    End With
End Sub
Public Sub RenderArrow(Optional ByVal Fore As Long = -1, Optional ByVal x1 As Long, Optional ByVal y1 As Long, Optional ByVal aSize As Long = 3, Optional bFill As Boolean = True, Optional ArrowDir As JButtonArrowStyle)
Dim PT(2)       As POINTAPI
Dim oldFore     As Long
Dim oldFill     As Integer
   
    oldFore = UserControl.ForeColor
    oldFill = UserControl.FillStyle
    If Fore <> -1 Then UserControl.ForeColor = Fore
    UserControl.FillStyle = IIf(bFill, 0, 1)
    UserControl.FillColor = UserControl.ForeColor
    
    If aSize Mod 2 = 0 Then aSize = aSize + 1
    aSize = aSize - 1
    
    Select Case ArrowDir
        Case 0 'ToDown
            PT(0).x = x1:                   PT(0).y = y1
            PT(1).x = x1 + aSize:           PT(1).y = y1
            PT(2).x = x1 + (aSize \ 2):     PT(2).y = y1 + (aSize \ 2)
        Case 1 'ToUp
            PT(0).x = x1 + (aSize \ 2):     PT(0).y = y1
            PT(1).x = x1:                   PT(1).y = y1 + (aSize \ 2)
            PT(2).x = x1 + aSize:           PT(2).y = y1 + (aSize \ 2)
        Case 2 'ToLeft
            PT(0).x = x1:                   PT(0).y = y1
            PT(1).x = x1 + (aSize \ 2):     PT(1).y = y1 + (aSize \ 2)
            PT(2).x = x1:                   PT(2).y = y1 + aSize
        Case 3 ' ToRight
            PT(0).x = x1:                   PT(0).y = y1 + (aSize \ 2)
            PT(1).x = x1 + (aSize \ 2):     PT(1).y = y1
            PT(2).x = x1 + (aSize \ 2):     PT(2).y = y1 + aSize
    End Select
    Polygon Hdc, PT(0), 3
    If UserControl.ForeColor <> oldFore Then UserControl.ForeColor = oldFore
    If UserControl.FillStyle <> oldFill Then UserControl.FillStyle = oldFill
End Sub
Public Sub DrawSkin(DC As Long, ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal btnState As JButtonState)
Dim lPx As Integer
    lPx = btnState * (m_SkinH / 5)
    RenderStretchPlus DC, 0, x, y, Width, Height, m_Skin, 0, lPx, m_SkinW, m_SkinH / 5, 5
End Sub
Public Sub Redraw()
    pvDraw m_lState, True
End Sub

Public Sub SetPictureStream(bvData() As Byte)
On Error GoTo e
    ppUnloadBitmap m_Bitmap
    ppCreateBitmapFromStream bvData, m_Bitmap
    GdipGetImageDimension m_Bitmap, m_BmpSrcW, m_BmpSrcH
    pvUpdateRects
    pvDraw m_lState, True
e:
End Sub

Public Sub SetSkinStream(bvData() As Byte)
On Error GoTo e
Dim IStream     As IUnknown
    ppUnloadBitmap m_Skin
    Call CreateStreamOnHGlobal(bvData(0), 0&, IStream)
    If Not IStream Is Nothing Then
        If GdipLoadImageFromStream(IStream, m_Skin) = 0 Then
            Call GdipGetImageDimension(m_Skin, m_SkinW, m_SkinH)
            pvDraw m_lState, True
        End If
    End If
e:
    Set IStream = Nothing
End Sub


'?Button Sub -------------------

Private Sub pvDraw(ByVal Evt As JButtonState, Optional Force As Boolean)

    If m_bSpaceDwn Then Evt = 2
    If Evt = 0 And m_bFocus Then Evt = 4
    If Evt = m_lState And Not Force Then Exit Sub
    If Not UserControl.Enabled Then Evt = 3
         
    If m_bValue Then
        Select Case Evt
         Case 0, 1, 2, 4: Evt = 2
        End Select
    End If
    
    UserControl.Cls
    Dim hGraphic  As Long
    Dim bEvt      As Boolean
    
    Call GdipCreateFromHDC(UserControl.Hdc, hGraphic)
    RaiseEvent BeforePaint(UserControl.Hdc, hGraphic, UserControl.ScaleWidth, UserControl.ScaleHeight, Evt, bEvt)
    
    If Not bEvt And Not m_bNoBkgnd Then
        If m_Skin Then Call pvDrawSkinButton(Evt, hGraphic) Else pvDrawThemeButton Evt
    End If
    
    If m_Bitmap = 0 Then GoTo e

    Dim mtGray  As COLORMATRIX
    Dim mtColor As COLORMATRIX
    Dim hAtrr       As Long

    With mtColor
        .m(3, 3) = 1 ' [ALPHA]
        If m_BmpColor <> -1 And UserControl.Enabled = True Then
            Dim R As Byte, G As Byte, b As Byte
            b = ((m_BmpColor \ &H10000) And &HFF)
            G = ((m_BmpColor \ &H100) And &HFF)
            R = (m_BmpColor And &HFF)
            .m(0, 0) = R / 255
            .m(1, 0) = G / 255
            .m(2, 0) = b / 255
            .m(0, 4) = R / 255
            .m(1, 4) = G / 255
            .m(2, 4) = b / 255
        ElseIf UserControl.Enabled = False Then
            .m(0, 0) = 0.299
            .m(1, 0) = .m(0, 0)
            .m(2, 0) = .m(0, 0)
            .m(0, 1) = 0.587
            .m(1, 1) = .m(0, 1)
            .m(2, 1) = .m(0, 1)
            .m(0, 2) = 0.114
            .m(1, 2) = .m(0, 2)
            .m(2, 2) = .m(0, 2)
            .m(3, 3) = 0.5
            .m(4, 4) = 1
        Else
            .m(0, 0) = 1
            .m(1, 1) = 1
            .m(2, 2) = 1
            .m(4, 4) = 1
        End If
        End With
    
    If (m_IRct.Right <> m_BmpSrcW) Or (m_IRct.Bottom <> m_BmpSrcH) Then
        Call GdipSetInterpolationMode(hGraphic, 7&)  '-> InterpolationModeHighQualityBicubic
        Call GdipSetPixelOffsetMode(hGraphic, 4&)
    End If
        
    If GdipCreateImageAttributes(hAtrr) = 0 Then
        GdipSetImageAttributesColorMatrix hAtrr, 0, True, mtColor, mtGray, 0
        GdipDrawImageRectRectI hGraphic, m_Bitmap, m_IRct.Left, m_IRct.Top + IIf(m_beDown And Evt = 2, 1 * dpi_, 0), m_IRct.Right, m_IRct.Bottom, 0, 0, m_BmpSrcW, m_BmpSrcH, &H2, hAtrr, 0&, 0&
        Call GdipDisposeImageAttributes(hAtrr)
    End If
        
e:
   If Len(Trim(m_Text)) = 0 Then GoTo q
    
    Dim lFlag As Long
    Select Case m_algnText
        Case 0: lFlag = &H0& Or &H4& 'Or &H20&
        Case 1: lFlag = &H2& Or &H4& 'Or &H20&
        Case 2: lFlag = &H1& Or &H4& 'Or &H20&
    End Select
    
    UserControl.ForeColor = pvGetColor(Evt)
    If m_beDown And Evt = 2 Then OffsetRect m_TRct, 0, 1 * dpi_
    DrawText UserControl.Hdc, m_Text, -1, m_TRct, lFlag Or &H10 'Or &H100
    If m_beDown And Evt = 2 Then OffsetRect m_TRct, 0, -1 * dpi_

q:
    RaiseEvent AffterPaint(UserControl.Hdc, hGraphic, UserControl.ScaleWidth, UserControl.ScaleHeight, Evt)
    If hGraphic Then GdipDeleteGraphics (hGraphic): hGraphic = 0
    m_lState = Evt
    
End Sub
Private Sub pvDrawSkinButton(ByVal Evt As Integer, hGraphic As Long)
Dim lPx As Integer
    lPx = Evt * (m_SkinH / 5)
    With UserControl
        RenderStretchPlus .Hdc, hGraphic, 0, 0, .ScaleWidth, .ScaleHeight, m_Skin, 0, lPx, m_SkinW, m_SkinH / 5, 5 'm_iRct.nWidth / 4
    End With
End Sub

Private Sub pvDrawThemeButton(ByVal Evt As Integer)
Dim uRct As RECT
Dim hTheme As Long

    'If Evt = 0 Then Exit Sub
    With UserControl
        SetRect uRct, 0, 0, .ScaleWidth, .ScaleHeight
        hTheme = OpenThemeData(.hWnd, StrPtr("Button"))
        
        If hTheme Then
            If Not m_bValue Then
                Call DrawThemeBackGround(hTheme, .Hdc, 1, Evt + 1, uRct, ByVal 0&) '/Normal
            Else
                Select Case Evt
                    Case 0, 1, 2, 4
                        Call DrawThemeBackGround(hTheme, .Hdc, 1, 3, uRct, ByVal 0&) '/Pressed
                    Case 3
                        Call DrawThemeBackGround(hTheme, .Hdc, 1, 4, uRct, ByVal 0&) '/Disabled
                    Case Else
                        'Debug.Print lSE
                End Select
            End If
            CloseThemeData hTheme
        Else
            If m_bValue Then
                DrawFrameControl .Hdc, uRct, &H4, &H10 Or &H200
            Else
                Select Case Evt
                    Case 0: DrawFrameControl .Hdc, uRct, &H4, &H10
                    Case 1: DrawFrameControl .Hdc, uRct, &H4, &H10 Or &H1000
                    Case 2: DrawFrameControl .Hdc, uRct, &H4, &H10 Or &H200
                    Case 4: DrawFrameControl .Hdc, uRct, &H4, &H10
                    Case Else
                End Select
            End If
        End If
    End With
End Sub
Private Sub ppUnloadBitmap(lBmp As Long)
    If lBmp <> 0 Then Call GdipDisposeImage(lBmp): lBmp = 0
End Sub
Private Sub ppCreateSkin()
On Error GoTo e
Dim sData()     As String
Dim bvData()    As Byte
Dim IStream     As IUnknown
    
    ppUnloadBitmap m_Skin
    If Trim(m_SkinRes) = "" Then Exit Sub
    
    sData = Split(m_SkinRes, "|")
    bvData = LoadResData(Val(sData(0)), sData(1))
    'ppCreateBitmapFromStream bvData, m_Skin
    Call CreateStreamOnHGlobal(bvData(0), 0&, IStream)
    
    If Not IStream Is Nothing Then
        If GdipLoadImageFromStream(IStream, m_Skin) = 0 Then
            Call GdipGetImageDimension(m_Skin, m_SkinW, m_SkinH)
        End If
    End If
e:
    Set IStream = Nothing
End Sub
Private Function ppCreateBitmapFromStream(bvData() As Byte, lBitmap As Long) As Long
On Error GoTo e
Dim IStream As IUnknown

    Call CreateStreamOnHGlobal(bvData(0), 0&, IStream)
    If Not IStream Is Nothing Then
        Call GdipLoadImageFromStream(IStream, lBitmap)
    End If
e:
    Set IStream = Nothing
End Function


'? PropertyPage     --------------------------------------------
Friend Function ppgGetImageStream() As Byte()
    ppgGetImageStream = m_bvData
End Function
Friend Function ppgSetImageStream(lData() As Byte)
    m_bvData() = lData
    ppUnloadBitmap m_Bitmap
    ppCreateBitmapFromStream lData, m_Bitmap
    GdipGetImageDimension m_Bitmap, m_BmpSrcW, m_BmpSrcH
    pvUpdateRects
    pvDraw m_lState, True
    PropertyChanged "Image"
End Function


'?GDI
Private Sub ManageGdip(ByVal StartUp As Boolean)
    If StartUp Then
        If m_Token = 0& Then
            Dim gdipSI(3) As Long
            gdipSI(0) = 1&
            Call GdiplusStartup(m_Token, gdipSI(0), ByVal 0)
        End If
    Else
        If m_Token <> 0 Then
            Call GdiplusShutdown(m_Token)
            m_Token = 0
        End If
    End If
End Sub

Private Sub RenderStretchPlus(ByVal DestHdc As Long, ByVal hGraphics As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal DestW As Long, _
                            ByVal DestH As Long, ByVal hImage As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Size As Long)
Dim Sx2         As Long
Dim mbFlag      As Boolean
    
    Sx2 = Size * 2
    If hGraphics = 0 Then
        mbFlag = GdipCreateFromHDC(DestHdc, hGraphics) = 0
        If Not mbFlag Then Exit Sub
    End If
    
    Call GdipSetInterpolationMode(hGraphics, 5&)
    Call GdipSetPixelOffsetMode(hGraphics, 4&)

    GdipDrawImageRectRectI hGraphics, hImage, DestX, DestY, Size, Size, x, y, Size, Size, &H2, 0&, 0&, 0& 'TOP_LEFT
    GdipDrawImageRectRectI hGraphics, hImage, DestX + Size, DestY, DestW - Sx2, Size, x + Size, y, Width - Sx2, Size, &H2, 0&, 0&, 0& 'TOP_CENTER
    GdipDrawImageRectRectI hGraphics, hImage, DestX + DestW - Size, DestY, Size, Size, x + Width - Size, y, Size, Size, &H2, 0&, 0&, 0& 'TOP_RIGHT
    GdipDrawImageRectRectI hGraphics, hImage, DestX, DestY + Size, Size, DestH - Sx2, x, y + Size, Size, Height - Sx2, &H2, 0&, 0&, 0& 'MID_LEFT
    GdipDrawImageRectRectI hGraphics, hImage, DestX + Size, DestY + Size, DestW - Sx2, DestH - Sx2, x + Size, y + Size, Width - Sx2, Height - Sx2, &H2, 0&, 0&, 0& 'MID_CENTER
    GdipDrawImageRectRectI hGraphics, hImage, DestX + DestW - Size, DestY + Size, Size, DestH - Sx2, x + Width - Size, y + Size, Size, Height - Sx2, &H2, 0&, 0&, 0& 'MID_RIGHT
    GdipDrawImageRectRectI hGraphics, hImage, DestX, DestY + DestH - Size, Size, Size, x, y + Height - Size, Size, Size, &H2, 0&, 0&, 0& 'BOTTOM_LEFT
    GdipDrawImageRectRectI hGraphics, hImage, DestX + Size, DestY + DestH - Size, DestW - Sx2, Size, x + Size, y + Height - Size, Width - Sx2, Size, &H2, 0&, 0&, 0& 'BOTTOM_CENTER
    GdipDrawImageRectRectI hGraphics, hImage, DestX + DestW - Size, DestY + DestH - Size, Size, Size, x + Width - Size, y + Height - Size, Size, Size, &H2, 0&, 0&, 0& 'BOTTOM_RIGHT

    If mbFlag Then Call GdipDeleteGraphics(hGraphics)

    
End Sub

'? Button Size
Private Sub pvUpdateRects()
Dim TRct As RECT
Dim IRct As RECT
Dim Lm   As Integer
Dim Rm   As Integer
Dim PxSpc  As Long

Dim lBmpW  As Single
Dim lBmpH  As Single

    With UserControl
        
        .Cls
        Lm = pvGetvalue(0) * dpi_
        Rm = pvGetvalue(1) * dpi_

        If m_Bitmap Then
        
            PxSpc = m_bmpSpc * dpi_
    
            IRct.Right = pvGetvalue(0, 1) * dpi_
            IRct.Bottom = pvGetvalue(1, 1) * dpi_
            If IRct.Right <> 0 Then lBmpW = IRct.Right Else lBmpW = m_BmpSrcW
            If IRct.Bottom <> 0 Then lBmpH = IRct.Bottom Else lBmpH = m_BmpSrcH
            SetRect IRct, 0, 0, lBmpW, lBmpH
            
            
            '0-Top, 1-Left, 2-Right, 3-Behind
            Select Case m_algnBmp
                Case 0: SetRect TRct, Lm, lBmpH, .ScaleWidth - Rm, .ScaleHeight         'TOP
                Case 1: SetRect TRct, Lm, 0, .ScaleWidth - Rm - lBmpW, .ScaleHeight     'LEFT
                Case 2: SetRect TRct, Lm, 0, .ScaleWidth - Rm - lBmpW, .ScaleHeight     'RIGHT
                Case 3: SetRect TRct, Lm, 0, .ScaleWidth - Rm, .ScaleHeight             'BEHIND
            End Select
            
           
        Else
            SetRect TRct, Lm, 0, .ScaleWidth - Rm, .ScaleHeight
            SetRect IRct, 0, 0, 0, 0
        End If
        
        If Len(Trim(m_Text)) Then
            DrawText Hdc, m_Text, Len(m_Text), TRct, &H10 Or &H400 Or &H100
            OffsetRect TRct, -TRct.Left, -TRct.Top
        Else
            PxSpc = 0
            SetRect TRct, 0, 0, 0, 0
        End If
        
        
        '?Final
        Select Case m_algnText
        
            Case 2: 'Center Text
            
            Select Case m_algnBmp
                Case 0 ' Top
                    OffsetRect IRct, (.ScaleWidth - IRct.Right) \ 2, (.ScaleHeight - (IRct.Bottom + TRct.Bottom)) \ 2
                    OffsetRect TRct, ((.ScaleWidth - TRct.Right) \ 2), IRct.Bottom + PxSpc
                Case 1 ' Left
                    OffsetRect IRct, ((.ScaleWidth - (IRct.Right + TRct.Right)) \ 2), (.ScaleHeight - IRct.Bottom) \ 2
                    OffsetRect TRct, IRct.Right + PxSpc, (.ScaleHeight - TRct.Bottom) \ 2
                Case 2 ' Right
                    OffsetRect TRct, ((.ScaleWidth - (IRct.Right + TRct.Right)) \ 2), (.ScaleHeight - TRct.Bottom) \ 2
                    OffsetRect IRct, TRct.Right + PxSpc, (.ScaleHeight - IRct.Bottom) \ 2
                Case 3
                    OffsetRect TRct, (.ScaleWidth - TRct.Right) \ 2, (.ScaleHeight - TRct.Bottom) \ 2
            End Select
            
            Case 0 'Left Text
            
            Select Case m_algnBmp
                Case 1 ' Left
                    OffsetRect IRct, Lm, (.ScaleHeight - IRct.Bottom) \ 2
                    OffsetRect TRct, IRct.Right + PxSpc, (.ScaleHeight - TRct.Bottom) \ 2
                Case 2 ' Right
                    OffsetRect TRct, Lm, (.ScaleHeight - TRct.Bottom) \ 2
                    OffsetRect IRct, TRct.Right + PxSpc, (.ScaleHeight - IRct.Bottom) \ 2
                Case 3
                    OffsetRect TRct, Lm, (.ScaleHeight - TRct.Bottom) \ 2
            End Select
            
            Case 1 'Right Text
            
            Select Case m_algnBmp
                Case 1 ' Left
                    OffsetRect TRct, .ScaleWidth - Rm - TRct.Right, (.ScaleHeight - TRct.Bottom) \ 2
                    OffsetRect IRct, TRct.Left - IRct.Right - PxSpc, (.ScaleHeight - IRct.Bottom) \ 2
                Case 2 ' Right
                    OffsetRect IRct, .ScaleWidth - Rm - IRct.Right, (.ScaleHeight - IRct.Bottom) \ 2
                    OffsetRect TRct, IRct.Left - TRct.Right - PxSpc, (.ScaleHeight - TRct.Bottom) \ 2
                Case 3
                    OffsetRect TRct, .ScaleWidth - Rm - TRct.Right, (.ScaleHeight - TRct.Bottom) \ 2
            End Select
        End Select
        
        If m_algnBmp = 3 Then
            OffsetRect IRct, (.ScaleWidth - lBmpW) \ 2, (.ScaleHeight - lBmpH) \ 2
        End If
        
    End With
    CopyRect m_TRct, TRct
    CopyRect m_IRct, IRct
    SetRect m_IRct, m_IRct.Left, m_IRct.Top, lBmpW, lBmpH

End Sub
Private Function pvGetvalue(Index As Integer, Optional lType As Integer) As Long
On Error GoTo e
    If lType = 0 Then pvGetvalue = Val(Split(m_Margins, ",")(Index))
    If lType = 1 Then pvGetvalue = Val(Split(m_BmpSize, "x")(Index))
e:
End Function
Private Function pvGetColor(ByVal eState As Integer) As Long
    If m_lColor(eState) <> -1 Then
        pvGetColor = m_lColor(eState)
    Else
        Select Case eState
            Case 0, 1, 2, 4
                pvGetColor = vbButtonText
            Case 3
                pvGetColor = vbGrayText
        End Select
    End If
End Function
Private Sub pUpdateOptionButtons()
Dim lhWnd           As Long
Dim TmpValue        As Boolean
Dim Ctrl            As Control


    Dim Frm As Object
    Set Frm = Extender.Parent
    lhWnd = Extender.Container.hWnd
    For Each Ctrl In Frm.Controls
        DoEvents
        With Ctrl
           If TypeOf Ctrl Is JButton Then
              If .ButtonType = 2 Then
                 If (.Container.hWnd = lhWnd) And (Ctrl.hWnd <> UserControl.hWnd) Then
                    If .Value Then .Value = False
                 End If
              End If
           End If
        End With
    Next
End Sub
Private Sub pvSetTrans()
On Error GoTo e
Dim oPic As StdPicture
    With UserControl
        Set .Picture = Nothing
        If m_aImage Then
            Set oPic = Extender.Container.Image
            UserControl.PaintPicture oPic, 0, 0, , , Extender.Left, Extender.Top ', Extender.Width * 15, Extender.Height * 15
            Set .Picture = .Image
        End If
    End With
    Exit Sub
e:
    Set UserControl.Picture = Nothing
End Sub

'?Mouse
Private Function pMouseOnButton() As Boolean
Dim PT As POINTAPI
    GetCursorPos PT
    pMouseOnButton = (WindowFromPoint(PT.x, PT.y) = UserControl.hWnd)
End Function


Private Function GetWindowsDPI() As Double
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


'- ordinal #1
Private Sub WndProc(ByVal bBefore As Boolean, _
       ByRef bHandled As Boolean, _
       ByRef lReturn As Long, _
       ByVal hWnd As Long, _
       ByVal uMsg As Long, _
       ByVal wParam As Long, _
       ByVal lParam As Long, _
       ByRef lParamUser As Long)

    Select Case uMsg
        'Case WM_DESTROY
        Case WM_MOUSELEAVE
                m_bTrack = False
                m_lButton = 0
                pvDraw 0
                RaiseEvent MouseLeave
        Case WM_KILLFOCUS, WM_SETFOCUS
                m_bFocus = (uMsg = WM_SETFOCUS)
                If uMsg = WM_KILLFOCUS Then
                    m_lButton = 0
                    m_lMouseDown = True
                    m_lBtnDown = True
                    m_bSpaceDwn = False
                End If
                If Not pMouseOnButton Then pvDraw 0, True
                
        'Case WM_NCPAINT
    End Select
End Sub


