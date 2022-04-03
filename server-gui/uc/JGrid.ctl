VERSION 5.00
Begin VB.UserControl JGrid 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5340
   ControlContainer=   -1  'True
   FillColor       =   &H000040C0&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   ScaleHeight     =   179
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   356
   ToolboxBitmap   =   "JGrid.ctx":0000
End
Attribute VB_Name = "JGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================================================================================
'    Component  : JGrid 2.8.5
'    Autor      : J. Elihu
'    Description: Quick Grid Control
'    Req        : cScrollBar && cSubClass
'=====================================================================================================================

Option Explicit

Private Type POINTAPI
  x      As Long
  Y      As Long
End Type
Private Type Rect
  Left   As Long
  Top    As Long
  Right  As Long
  Bottom As Long
End Type

'/Header Style
Private Const HDS_HORZ = &H0
Private Const HDS_BUTTONS = &H2
Private Const HDS_HIDDEN = &H8
Private Const HDS_HOTTRACK = &H4
Private Const HDS_DRAGDROP = &H40
Private Const HDS_FULLDRAG = &H80

'/Header Item
Private Const HDI_WIDTH = &H1
Private Const HDI_HEIGHT = HDI_WIDTH
Private Const HDI_TEXT = &H2
Private Const HDI_FORMAT = &H4
Private Const HDI_LPARAM = &H8
Private Const HDI_BITMAP = &H10
Private Const HDI_IMAGE = &H20
Private Const HDI_DI_SETITEM = &H40
Private Const HDI_ORDER = &H80
Private Const HDI_FILTER = &H100

'/Header Messages
Private Const HDM_FIRST = &H1200
Private Const HDM_GETITEMCOUNT = (HDM_FIRST + 0)
Private Const HDM_INSERTITEM = (HDM_FIRST + 1)
Private Const HDM_HITTEST = (HDM_FIRST + 6)
Private Const HDM_GETITEMDROPDOWNRECT = (HDM_FIRST + 25)
Private Const HDM_DELETEITEM = (HDM_FIRST + 2)
Private Const HDM_GETITEM = (HDM_FIRST + 3)
Private Const HDM_SETITEM = (HDM_FIRST + 4)
Private Const HDM_LAYOUT = (HDM_FIRST + 5)
Private Const HDM_ORDERTOINDEX = (HDM_FIRST + 15)
Private Const HDM_GETITEMRECT = (HDM_FIRST + 7)
Private Const HDM_SETIMAGELIST = (HDM_FIRST + 8)
Private Const HDM_GETIMAGELIST = (HDM_FIRST + 9)

'/Header Messages
Private Const H_MAX                 As Long = &HFFFF + 1
Private Const HDN_FIRST             As Long = H_MAX - 300&                  '// header
Private Const HDN_LAST              As Long = H_MAX - 399&
Private Const HDN_ITEMCHANGING      As Long = (HDN_FIRST - 20) 'Unicode
Private Const HDN_ITEMCLICK         As Long = (HDN_FIRST - 22) 'Unicode
Private Const HDN_ITEMDBLCLICK      As Long = (HDN_FIRST - 23) 'Unicode
Private Const HDN_DIVIDERDBLCLICK   As Long = (HDN_FIRST - 25) 'unicode
Private Const HDN_BEGINTRACK        As Long = (HDN_FIRST - 26) 'Unicode
Private Const HDN_ENDTRACK          As Long = (HDN_FIRST - 27) 'Unicode
Private Const HDN_TRACK             As Long = (HDN_FIRST - 28) 'Unicode
Private Const HDN_DROPDOWN          As Long = (HDN_FIRST - 18)
Private Const HDN_FILTERBTNCLICK    As Long = (HDN_FIRST - 13)
Private Const HDN_FILTERCHANGE      As Long = (HDN_FIRST - 12)
Private Const HDN_ITEMCHECK         As Long = (HDN_FIRST - 16) 'The name is invented, not found his real name
Private Const HDN_BEGINDRAG         As Long = (HDN_FIRST - 10)
Private Const HDN_ENDDRAG           As Long = (HDN_FIRST - 11)

'/Header Flags
Private Const HDF_OWNERDRAW         As Long = &H8000
Private Const HDF_STRING            As Long = &H4000
Private Const HDF_BITMAP            As Long = &H2000
Private Const HDF_IMAGE             As Long = &H800

Private Type HDITEM
  mask        As Long
  cxy         As Long
  pszText     As String
  hbm         As Long
  cchTextMax  As Long
  fmt         As Long
  lParam      As Long
  iImage      As Long
  iOrder      As Long
  type        As Long
  pvFilter    As Long
End Type
Private Type HDHITTESTINFO
  PT          As POINTAPI
  Flags       As Long
  iItem       As Long
End Type
Private Type NMHDR
  hwndFrom    As Long
  idfrom      As Long
  code        As Long
End Type
Private Type NMHEADER
  Hdr         As NMHDR
  iItem       As Long
  iButton     As Long
  lPtrHDItem  As Long '    HDITEM  FAR* pItem
End Type

Private Type PAINTSTRUCT
  Hdc         As Long
  fErase      As Long
  RctPaint    As Rect
  fRestore    As Long
  fIncUpdate  As Long
  rgbReserved(1 To 32) As Byte
End Type

'/HEADER CUSTOM DRAW
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function GetCurrentObject Lib "gdi32" (ByVal Hdc As Long, ByVal uObjectType As Long) As Long

'/HEADER WINDOW
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetFocusEx Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32.dll" (ByRef lpEventTrack As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'/MOUSE
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long

'/Border
Private Declare Function GetWindowRect& Lib "user32" (ByVal hWnd As Long, lpRect As Rect)
Private Declare Function ExcludeClipRect Lib "gdi32.dll" (ByVal Hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal Hdc As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal Hdc As Long, lpRect As Rect) As Long

'/WINDOW MESSAGES
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'/ImageList
Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal Flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
Private Declare Function ImageList_AddMasked Lib "comctl32.dll" (ByVal himl As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_Add Lib "comctl32.dll" (ByVal himl As Long, ByVal hbmImage As Long, ByVal hbmMask As Long) As Long
Private Declare Function ImageList_GetIconSize Lib "comctl32.dll" (ByVal hImageList As Long, cx As Long, cy As Long) As Long
Private Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal Flags As Long) As Long
Private Declare Function ImageList_DrawEx Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long

'/Draw
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal Hdc As Long, ByVal x As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Function OleTranslateColor2 Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal Hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal Hdc As Long, ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal Hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As Rect, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As Rect, lpSourceRect As Rect) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As Rect, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal Hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal Hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal Hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal Hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal Hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal Hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal Hdc As Long, ByVal nIndex As Long) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function Polygon Lib "gdi32" (ByVal Hdc As Long, lpPoint As Any, ByVal nCount As Long) As Long

'/GDI+
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As Long, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal Image As Long, ByRef Width As Single, ByRef Height As Single) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Any, ByRef Image As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstWidth As Long, ByVal DstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal Hdc As Long, ByRef Graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal Graphics As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal Format As Long, ByRef Scan0 As Any, ByRef BITMAP As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal Graphics As Long, ByVal InterpolationMode As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal Graphics As Long, ByVal PixelOffsetMode As Long) As Long

Private Declare Sub CreateStreamOnHGlobal Lib "ole32" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)


'/JGrid Events
Event MouseEnter()
Event MouseExit()
Event ItemClick(ByVal Item As Long, ByVal Column As Long)
Event ItemDblClick(ByVal Item As Long, ByVal Column As Long)
Event ItemMouseUp(ByVal Item As Long, ByVal Column As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
Event ItemMouseDown(ByVal Item As Long, ByVal Column As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
Event SelectionChanged(ByVal Item As Long, ByVal Column As Long)

Event ItemDrawData(ByVal Item As Long, ByVal Column As Long, ByRef ForeColor As Long, ByRef BackColor As Long, ByRef BorderColor As Long, ByRef ItemIdent As Long)
Event ItemDraw(ByVal Item As Long, ByVal Column As Long, ByRef Hdc As Long, ByRef hGraphic As Long, ByVal x As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, ByRef CancelDraw As Boolean)
Event ItemDrawMeasureText(ByVal Item As Long, ByVal Column As Long, ByRef x As Long, ByRef Y As Long, ByRef x2 As Long, ByRef y2 As Long)

Event EditStart(ByVal Item As Long, ByVal Column As Long, ByRef x As Long, ByRef Y As Long, ByRef W As Long, ByRef H As Long, ByRef Text As String, ByRef ObjEdit As Control, ByRef Cancel As Boolean, ByRef MoveObj As Boolean)
Event EditEnd(ByVal Item As Long, ByVal Column As Long, ByRef NewText As String, ByRef ObjEdit As Control, ByRef Cancel As Boolean)
Event EditShow(ByVal Item As Long, ByVal Column As Long, ByRef ObjEdit As Control, ByRef Visible As Boolean)

'/Header Events
Event ColumnClick(ByVal Column As Long)
Event ColumnRightClick(ByVal Column As Long)
Event ColumnDblClick(ByVal Column As Long)
Event ColumnSizeChangeStart(ByVal Column As Long, ByVal Width As Long, Cancel As Boolean)
Event ColumnSizeChanging(ByVal Column As Long, ByVal Width As Long, Cancel As Boolean)
Event ColumnSizeChanged(ByVal Column As Long, ByVal Width As Long)
Event ColumnDividerDblClick(ByVal Column As Long)

Event HeaderBkgndDraw(Hdc As Long, W As Long, H As Long)
Event HeaderColumnDraw(ByVal Column As Long, Hdc As Long, x As Long, W As Long, H As Long, lBState As JGridMouseState)
Event HeaderColumnTextDraw(ByVal Column As Long, Hdc As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, Color As Long, lBState As JGridMouseState, Cancel As Boolean)

'/Scroll
Event Scroll(eBar As EFSScrollBarConstants)
Event ScrollChange(eBar As EFSScrollBarConstants)
Event ScrollClick(eBar As EFSScrollBarConstants, eButton As MouseButtonConstants)

'/Standart
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)


Public Enum JGridCoincidence
  [CC_WHOLE_WORD] = 0
  [CC_PARTIAL] = 1
End Enum
Public Enum JGridSortOrder
  [ASC_ORDER] = 0
  [DESC_ORDER] = 1
End Enum
Public Enum JGridMouseState
  HBS_NONE
  HBS_HOT
  HBS_DOWN
End Enum

Private Type tHeader
  Text    As String
  Image   As Long
  Width   As Long
  Id      As Long
  Fixed   As Boolean
  Aling   As Integer
  IAlign  As Integer
  MinW    As Long
  NoEdit  As Boolean
End Type

Private Type tSubItem
  Text    As String
  Icon    As Long
  Tag     As String
End Type

Private Type tItem
  Item()  As tSubItem
  data    As Long
End Type

Private Type tEventDrawing
  Fore    As Long
  Back    As Long
  Border  As Long
  Ident   As Long
  Cancel  As Boolean
End Type

Private m_ItemH         As Long
Private m_PropItemH     As Long
Private m_HeaderH       As Long
Private m_GridLineColor As Long
Private m_GridStyle     As Integer
Private m_Striped       As Boolean
Private m_Header        As Boolean
Private m_FullRow       As Boolean
Private m_Editable      As Boolean
Private m_DrawEmpty     As Boolean
Private m_CustomDraw    As Boolean
Private m_FocusRect     As Boolean
Private m_AlphaBlend    As Boolean
Private m_eBorderSize   As Integer

Private m_IFont         As IFont
Private m_StripedColor  As OLE_COLOR
Private m_ForeColor     As OLE_COLOR
Private m_SelColor      As OLE_COLOR
Private m_ForeSel       As OLE_COLOR
Private m_BorderColor   As OLE_COLOR
Private m_eBorderColor  As OLE_COLOR
Private m_eBackColor    As OLE_COLOR

Private WithEvents c_SubClass   As cSubClass
Attribute c_SubClass.VB_VarHelpID = -1
Private WithEvents c_Scroll     As cScrollBars
Attribute c_Scroll.VB_VarHelpID = -1
Private WithEvents c_Edit       As TextBox
Attribute c_Edit.VB_VarHelpID = -1

Private gdip_       As Long
Private dpi_        As Single
Private Track_(3)   As Long

Private m_hWnd      As Long
Private m_iml       As Long
Private m_bmdhFlag  As Boolean
Private m_Bmp       As Long

Private m_RowH      As Long
Private m_GridW     As Long
Private m_ImgW      As Long
Private m_ImgH      As Long
Private m_SelCol    As Long
Private m_SelRow    As Long
Private t_Col       As Long
Private t_Row       As Long
Private e_Row       As Long
Private e_Col       As Long

Private m_Cols()    As tHeader
Private m_Items()   As tItem

Private e_Ctrl      As Control
Private e_hWnd      As Long

Private mbTrack     As Boolean
Private mbImlFlag   As Boolean
Private mbEditFlag  As Boolean
Private mbNoDraw    As Boolean
Private mbResize    As Boolean
Private mbSelFirst  As Boolean
Private mlHdrBtn    As Long
Private mlSortCol   As Long


Private Sub UserControl_Initialize()
Dim gdipSI(3) As Long

    gdipSI(0) = 1&
    Call GdiplusStartup(gdip_, gdipSI(0), ByVal 0)
            
    Set c_SubClass = New cSubClass
    Set c_Scroll = New cScrollBars
    
    t_Col = -1: t_Row = -1
    m_SelCol = -1: m_SelRow = -1
    
    dpi_ = mvGetWindowsDPI
    mlHdrBtn = -1
    mlSortCol = -1
    
End Sub

Private Sub UserControl_InitProperties()
    m_HeaderH = 24
    m_GridLineColor = &HF0F0F0
    m_GridStyle = 3
    m_Striped = True
    m_StripedColor = &HFDFDFD
    m_SelColor = &HFF6600 'vbHighlight  '&HDDAC84
    m_BorderColor = &H908782  '&HB2ACA5
    m_eBorderColor = &HD9AD7E
    m_eBackColor = vbWhite
    m_eBorderSize = 1
    
    m_FullRow = True
    m_Header = True
    m_DrawEmpty = True
    m_AlphaBlend = True

    Set m_IFont = New StdFont
    m_IFont.Name = "Tahoma"
    
End Sub
Private Sub UserControl_Click()
    If mbEditFlag Then EditEnd
    If t_Row <> -1 And t_Col <> -1 Then
        RaiseEvent ItemClick(t_Row, t_Col)
        If mbSelFirst And m_Editable Then EditStart t_Row, t_Col
    End If
    RaiseEvent Click
End Sub
Private Sub UserControl_DblClick()
    If t_Row <> -1 And t_Col <> -1 Then
        RaiseEvent ItemDblClick(t_Row, t_Col)
        If m_Editable Then EditStart t_Row, t_Col
    End If
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Dim iRow As Long
Dim iCol As Long

    RaiseEvent KeyDown(KeyCode, Shift)
    If ItemCount = 0 Then Exit Sub
    If mbEditFlag Then EditEnd
    
    If m_Editable Then
    
        Select Case KeyCode
            Case vbKeyDown, vbKeyUp, vbKeyRight, vbKeyTab, vbKeyLeft, vbKeyEnd, vbKeyHome, vbKeyPageDown, vbKeyPageUp: GoTo Evt
            Case Else
                'If EditStart(m_SelRow, m_SelCol) Then
                    'Select Case KeyCode
                        'Case 8: e_Ctrl.Text = ""
                        'Case Else
                            'e_Ctrl.Text = Chr(KeyCode)
                            'e_Ctrl.SelStart = Len(e_Ctrl.Text) ': e_Ctrl.SelLength = Len(e_Ctrl.Text)
                    'End Select
                    'Exit Sub
                'End If
        End Select
        
    End If
    
Evt:
    Select Case KeyCode
        Case vbKeyDown
            If m_SelRow < ItemCount - 1 Then ChangeSelection m_SelRow + 1, m_SelCol
        Case vbKeyUp
            If m_SelRow > 0 Then ChangeSelection m_SelRow - 1, m_SelCol
            
        Case vbKeyRight, vbKeyTab
            If m_SelCol < ColumnCount - 1 And Not m_FullRow Then
                ChangeSelection m_SelRow, m_SelCol + 1
            ElseIf m_SelCol = ColumnCount - 1 And Not m_FullRow Then
                If m_SelRow < ItemCount - 1 Then ChangeSelection m_SelRow + 1, 0
            End If
        Case vbKeyLeft
            If m_SelCol > 0 And Not m_FullRow Then
                ChangeSelection m_SelRow, m_SelCol - 1
            ElseIf m_SelCol = 0 And Not m_FullRow Then
                If Not m_SelRow = 0 Then ChangeSelection m_SelRow - 1, ColumnCount - 1
            End If
        Case vbKeyEnd, vbKeyHome
            If KeyCode = vbKeyEnd Then ChangeSelection ItemCount - 1, m_SelCol
            If KeyCode = vbKeyHome Then ChangeSelection 0, m_SelCol
        Case vbKeyPageDown, vbKeyPageUp
            If KeyCode = vbKeyPageDown Then c_Scroll.Value(efsVertical) = c_Scroll.Value(efsVertical) + c_Scroll.LargeChange(efsVertical)
            If KeyCode = vbKeyPageUp Then c_Scroll.Value(efsVertical) = c_Scroll.Value(efsVertical) - c_Scroll.LargeChange(efsVertical)
        Case Else
            
            On Error Resume Next
            Dim j           As Long
            Dim lStart      As Long
            Dim pChar       As String
            Dim iChar       As String
            Dim bFound      As Boolean
            Dim lCol        As Long
        
            lStart = m_SelRow + 1
            lCol = IIf(m_FullRow, 0, m_SelCol)
            If lStart > ItemCount - 1 Then lStart = 0
            pChar = Chr(KeyCode)
            If pChar = "" Then Exit Sub
            
            For j = lStart To ItemCount - 1
                iChar = UCase(Left(m_Items(j).Item(lCol).Text, 1))
                If iChar <> "" And pChar = iChar Then
                    ChangeSelection j, lCol
                    bFound = True
                    Exit For
                End If
            Next
            If Not bFound And lStart > 0 Then
                For j = 0 To lStart '- 1
                    iChar = UCase(Left(m_Items(j).Item(lCol).Text, 1))
                    If iChar <> "" And pChar = iChar Then
                        ChangeSelection j, lCol
                        Exit For
                    End If
                Next
            End If
            
    End Select
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    If m_Editable And KeyAscii = 13 And m_SelRow <> -1 And m_SelCol <> -1 And IsVisibleItem(m_SelRow, m_SelCol) Then EditStart m_SelRow, m_SelCol
End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    RaiseEvent MouseDown(Button, Shift, x, Y)
    If t_Row = -1 And t_Col = -1 Then
        t_Row = GetRowFromY(Y)
        t_Col = GetColFromX(x)
        If m_FullRow And t_Row <> -1 And t_Col = -1 Then t_Row = -1
    End If
    
    If t_Row <> -1 And t_Col <> -1 Then
        ChangeSelection t_Row, t_Col
        RaiseEvent ItemMouseDown(t_Row, t_Col, Button, Shift, x, Y)
    Else
        ChangeSelection -1, -1
    End If
    
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim iCol    As Long
Dim iRow    As Long

    If Not mbTrack Then
        TrackMouseEvent Track_(0)
        RaiseEvent MouseEnter
        mbTrack = True
    End If
    iRow = GetRowFromY(Y)
    iCol = GetColFromX(x)
    If m_FullRow And iRow <> -1 And iCol = -1 Then iRow = -1
    If iRow <> t_Row Or iCol <> t_Col Then
        t_Col = iCol: t_Row = iRow
        DrawGrid
    End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

    If Button = 1 And mlHdrBtn <> -1 Then
        mlHdrBtn = -1
        m_bmdhFlag = False
        If m_CustomDraw Then RedrawHeader
    End If
    
    If (t_Row = m_SelRow And t_Col = m_SelCol) And (m_SelRow <> -1 And m_SelCol <> -1) Then
        RaiseEvent ItemMouseUp(m_SelRow, m_SelCol, Button, Shift, x, Y)
    End If
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Private Sub UserControl_Resize()
On Local Error Resume Next
    If mbEditFlag Then EditEnd
    If mbResize = True Then Exit Sub
    mbResize = True
    Call Update
    mbResize = False
    RedrawHeader
End Sub
Private Sub UserControl_Show()
    DrawGrid True
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "HeaderH", m_HeaderH
        .WriteProperty "LineColor", m_GridLineColor
        .WriteProperty "GridStyle", m_GridStyle
        .WriteProperty "Striped", m_Striped
        .WriteProperty "StripedColor", m_StripedColor
        .WriteProperty "SelColor", m_SelColor
        .WriteProperty "ItemH", m_PropItemH
        .WriteProperty "BorderColor", m_BorderColor
        .WriteProperty "Header", m_Header
        .WriteProperty "FullRow", m_FullRow
        .WriteProperty "FocusRect", m_FocusRect
        .WriteProperty "ForeColor", m_ForeColor
        .WriteProperty "Editable", m_Editable
        .WriteProperty "EditBorder", m_eBorderColor
        .WriteProperty "EditBack", m_eBackColor
        .WriteProperty "EditSize", m_eBorderSize
        .WriteProperty "DrawEmpty", m_DrawEmpty
        .WriteProperty "Font", UserControl.Font
        .WriteProperty "Border", UserControl.BorderStyle
        .WriteProperty "BackColor", UserControl.BackColor
        .WriteProperty "Font2", m_IFont
        .WriteProperty "HeaderCustomDraw", m_CustomDraw
        .WriteProperty "AlphaBlend", m_AlphaBlend
    End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim H As Double

    With PropBag
        m_HeaderH = .ReadProperty("HeaderH", 24)
        m_GridLineColor = .ReadProperty("LineColor", &HF0F0F0)
        m_GridStyle = .ReadProperty("GridStyle", 3)
        m_Striped = .ReadProperty("Striped", True)
        m_StripedColor = .ReadProperty("StripedColor", &HFDFDFD)
        m_SelColor = .ReadProperty("SelColor", &HFF6600) '&HDDAC84)
        m_PropItemH = .ReadProperty("ItemH", 0)
        m_BorderColor = .ReadProperty("BorderColor", &HB2ACA5)
        m_Header = .ReadProperty("Header", True)
        m_FullRow = .ReadProperty("FullRow", True)
        m_FocusRect = .ReadProperty("FocusRect", False)
        m_ForeColor = .ReadProperty("ForeColor", 0)
        m_Editable = .ReadProperty("Editable", False)
        m_eBorderColor = .ReadProperty("EditBorder", &HD9AD7E)
        m_eBackColor = .ReadProperty("EditBack", vbWhite)
        m_eBorderSize = .ReadProperty("EditSize", 1)
        m_DrawEmpty = .ReadProperty("DrawEmpty", False)
        m_CustomDraw = .ReadProperty("HeaderCustomDraw", False)
        m_AlphaBlend = .ReadProperty("AlphaBlend", True)

        UserControl.BorderStyle = Abs(.ReadProperty("Border", True))
        UserControl.BackColor = .ReadProperty("BackColor", UserControl.BackColor)
        Set UserControl.Font() = .ReadProperty("Font", UserControl.Font)
        Set m_IFont = .ReadProperty("Font2", UserControl.Font)
        
        m_ItemH = m_PropItemH * dpi_
    End With
    

    If Ambient.UserMode Then
    
        Set c_Edit = UserControl.Controls.Add("VB.TextBox", "c_Edit")
        c_Edit.BorderStyle = 0
        c_Edit.Appearance = 0
        c_Edit.BackColor = m_eBackColor

        With c_Scroll
            .Create hWnd
            .SmallChange(efsHorizontal) = 20 '48
            .SmallChange(efsVertical) = 16
        End With
        
        With c_SubClass
            If .Subclass(hWnd, , , Me) Then
                .AddMsg hWnd, WM_NOTIFY, MSG_AFTER
                .AddMsg hWnd, WM_MOUSELEAVE, MSG_AFTER
                .AddMsg hWnd, WM_NCPAINT, MSG_AFTER
            End If
            
        End With
        
        Track_(0) = 16&
        Track_(1) = &H2
        Track_(2) = hWnd
        
        mvCreateHeader
        UpdateSizes
    End If
    
End Sub
Private Sub UserControl_Terminate()

    mvDestroyHeader
    If m_iml And mbImlFlag Then ImageList_Destroy m_iml: m_iml = 0
    If m_Bmp Then GdipDisposeImage m_Bmp

    Set c_SubClass = Nothing
    Set c_Scroll = Nothing
    
    If gdip_ <> 0 Then Call GdiplusShutdown(gdip_)
  
End Sub

Private Sub c_Scroll_Change(eBar As EFSScrollBarConstants)
    If mbEditFlag Then EditEnd
    DrawGrid True, True
    If eBar = 0 Then MoveHeader -ScrollValue(eBar)
    RaiseEvent ScrollChange(eBar)
End Sub
Private Sub c_Scroll_MouseWheel(eBar As EFSScrollBarConstants, lAmount As Long)
    DrawGrid True, True
End Sub
Private Sub c_Scroll_Scroll(eBar As EFSScrollBarConstants)
    If mbEditFlag Then EditEnd
    DrawGrid True, True
    If eBar = 0 Then MoveHeader -ScrollValue(eBar)
    RaiseEvent Scroll(eBar)
End Sub
Private Sub c_Scroll_ScrollClick(eBar As EFSScrollBarConstants, eButton As MouseButtonConstants)
    If mbEditFlag Then EditEnd
    SetFocusEx Me.hWnd
    RaiseEvent ScrollClick(eBar, eButton)
End Sub

Property Get hWnd() As Long: hWnd = UserControl.hWnd: End Property
Property Get hWndHeader() As Long: hWndHeader = m_hWnd: End Property
Property Get DpiScale() As Single: DpiScale = dpi_: End Property
Property Get Font() As StdFont: Set Font = UserControl.Font: End Property
Property Set Font(ByVal Value As StdFont)
    Set UserControl.Font() = Value
    UpdateSizes
    Update
    PropertyChanged "Font"
End Property
Property Get LastSortColumn() As Long: LastSortColumn = mlSortCol: End Property
Property Get HeaderVisible() As Boolean: HeaderVisible = m_Header: End Property
Property Let HeaderVisible(Value As Boolean)
    m_Header = Value
    If m_hWnd <> 0 Then ShowWindow m_hWnd, Abs(m_Header)
    UpdateScrollV
    DrawGrid
    PropertyChanged "Header"
End Property
Property Get HeaderHeight() As Long: HeaderHeight = m_HeaderH: End Property
Property Let HeaderHeight(ByVal Value As Long)
    m_HeaderH = Value
    MoveHeader lHeight:=m_HeaderH
    UpdateScrollV
    RedrawHeader
    DrawGrid
    PropertyChanged "HeaderH"
End Property
Property Get HeaderCustomDraw() As Boolean: HeaderCustomDraw = m_CustomDraw: End Property
Property Let HeaderCustomDraw(ByVal Value As Boolean)
    m_CustomDraw = Value
    RedrawHeader
    PropertyChanged "HeaderCustomDraw"
End Property
Property Get FontHeader() As IFont: Set FontHeader = m_IFont: End Property
Property Set FontHeader(ByVal Value As IFont)
    Set m_IFont = Value
    If m_hWnd <> 0 Then SendMessage m_hWnd, &H30, m_IFont.hFont, 0&
    PropertyChanged "Font2"
End Property

Property Get ItemHeight() As Long: ItemHeight = m_PropItemH: End Property
Property Let ItemHeight(ByVal Value As Long)
    m_PropItemH = Value
    m_ItemH = Value * dpi_
    UpdateSizes
    UserControl_Resize
    PropertyChanged "ItemH"
End Property
Property Get GridLineColor() As OLE_COLOR: GridLineColor = m_GridLineColor: End Property
Property Let GridLineColor(ByVal Value As OLE_COLOR)
    m_GridLineColor = Value
    PropertyChanged "LineColor"
    DrawGrid
End Property
Property Get GridLineStyle() As ScrollBarConstants: GridLineStyle = m_GridStyle: End Property
Property Let GridLineStyle(ByVal Value As ScrollBarConstants)
    m_GridStyle = Value
    UpdateSizes
    Update
    PropertyChanged "GridStyle"
End Property
Property Get StripedGrid() As Boolean: StripedGrid = m_Striped: End Property
Property Let StripedGrid(ByVal Value As Boolean)
    m_Striped = Value
    DrawGrid
    PropertyChanged "Striped"
End Property
Property Get BackColor() As OLE_COLOR: BackColor = UserControl.BackColor: End Property
Property Let BackColor(ByVal Value As OLE_COLOR)
    UserControl.BackColor = Value
    DrawGrid
    PropertyChanged "BackColor"
End Property
Property Get StripBackColor() As OLE_COLOR: StripBackColor = m_StripedColor: End Property
Property Let StripBackColor(ByVal Value As OLE_COLOR)
    m_StripedColor = Value
    PropertyChanged "StripedColor"
    DrawGrid
End Property
Property Get SelectionColor() As OLE_COLOR: SelectionColor = m_SelColor: End Property
Property Let SelectionColor(ByVal Value As OLE_COLOR)
    m_SelColor = Value
    DrawGrid
    PropertyChanged "SelColor"
End Property
Property Get SelectionAlphaBlend() As Boolean: SelectionAlphaBlend = m_AlphaBlend: End Property
Property Let SelectionAlphaBlend(ByVal Value As Boolean)
    m_AlphaBlend = Value
    PropertyChanged "AlphaBlend"
End Property

Property Get BorderColor() As OLE_COLOR: BorderColor = m_BorderColor: End Property
Property Let BorderColor(ByVal Value As OLE_COLOR)
    m_BorderColor = Value
    PropertyChanged "BorderColor"
End Property
Property Get FullRowSelection() As Boolean: FullRowSelection = m_FullRow: End Property
Property Let FullRowSelection(Value As Boolean)
    m_FullRow = Value
    PropertyChanged "FullRow"
    DrawGrid
End Property
Property Get FullRowFocusRect() As Boolean: FullRowFocusRect = m_FocusRect: End Property
Property Let FullRowFocusRect(ByVal Value As Boolean)
    m_FocusRect = Value
    DrawGrid
    PropertyChanged "FocusRect"
End Property

Property Get ForeColor() As OLE_COLOR: ForeColor = m_ForeColor: End Property
Property Let ForeColor(ByVal Value As OLE_COLOR)
    m_ForeColor = Value
    DrawGrid
    PropertyChanged "ForeColor"
End Property
Property Get Editable() As Boolean: Editable = m_Editable: End Property
Property Let Editable(ByVal Value As Boolean)
    m_Editable = Value
    PropertyChanged "Editable"
End Property
Property Get EditBorderColor() As OLE_COLOR: EditBorderColor = m_eBorderColor: End Property
Property Let EditBorderColor(ByVal Value As OLE_COLOR)
    m_eBorderColor = Value
    PropertyChanged "EditBorder"
End Property
Property Get EditBackColor() As OLE_COLOR: EditBackColor = m_eBackColor: End Property
Property Let EditBackColor(ByVal Value As OLE_COLOR)
    m_eBackColor = Value
    If Not c_Edit Is Nothing Then c_Edit.BackColor = m_eBackColor
    PropertyChanged "EditBack"
End Property
Property Get EditBorderSize() As Integer: EditBorderSize = m_eBorderSize: End Property
Property Let EditBorderSize(ByVal Value As Integer)
    m_eBorderSize = Value
    PropertyChanged "EditBorderSize"
End Property

Property Get DrawEmptyGrid() As Boolean: DrawEmptyGrid = m_DrawEmpty: End Property
Property Let DrawEmptyGrid(Value As Boolean)
    m_DrawEmpty = Value
    DrawGrid
    PropertyChanged "DrawEmpty"
End Property
Property Get Border() As Boolean: Border = UserControl.BorderStyle: End Property
Property Let Border(ByVal Value As Boolean)
    UserControl.BorderStyle() = Abs(Value)
    PropertyChanged "Border"
End Property
Property Get Edit() As TextBox: Set Edit = c_Edit: End Property



'[COLUMN FUNCTIONS]
Property Get ColumnWidth(ByVal Index As Long) As Long
On Error Resume Next
    ColumnWidth = m_Cols(Index).Width
End Property
Property Let ColumnWidth(ByVal Index As Long, ByVal Value As Long)
On Error GoTo e
Dim tHI As HDITEM
Dim i As Long
    If Not (m_Cols(Index).Width = Value) Then
        tHI.mask = HDI_WIDTH
        tHI.cxy = Value
        If (pSetHeaderItemInfo(Index, tHI)) Then
            m_GridW = (m_GridW - m_Cols(Index).Width) + Value
            m_Cols(Index).Width = Value
            UpdateScrollH
            DrawGrid
            RaiseEvent ColumnSizeChanged(Index, Value)
        End If
    End If
e:
End Property
Property Get ColumnMinWidth(ByVal Index As Long) As Long
On Error Resume Next
    ColumnMinWidth = m_Cols(Index).MinW
End Property
Property Let ColumnMinWidth(ByVal Index As Long, ByVal Value As Long)
On Error GoTo e
Dim tHI As HDITEM
Dim i   As Long

    Value = Value * dpi_
    m_Cols(Index).MinW = Value
    If Value <> 0 Then
        If ColumnWidth(Index) < Value Then ColumnWidth(Index) = Value
    End If
e:
End Property


'ITEMS FUNCTIONS
'=====================================================================================================================
Property Get ItemText(ByVal Item As Long, Optional ByVal Column As Long) As String
On Local Error Resume Next
    ItemText = m_Items(Item).Item(Column).Text
End Property
Property Let ItemText(ByVal Item As Long, Optional ByVal Column As Long, Value As String)
On Local Error Resume Next
    If m_Items(Item).Item(Column).Text = Value Then Exit Property
    m_Items(Item).Item(Column).Text = Value
    If mbNoDraw Then Exit Property
    If IsVisibleItem(Item, Column) Then DrawGrid
End Property
Property Get ItemIcon(ByVal Item As Long, Optional ByVal Column As Long) As Long
On Local Error GoTo e
    ItemIcon = m_Items(Item).Item(Column).Icon - 1  '/* Decrease Icon Index */
    Exit Property
e:
    ItemIcon = -1
End Property
Property Let ItemIcon(ByVal Item As Long, Optional ByVal Column As Long, Value As Long)
On Error GoTo e
    'If m_Items(Item).Item(Column).Icon - 1 = Value Then Exit Property '/* Decrease Icon Index */
    m_Items(Item).Item(Column).Icon = Value + 1     '/* Increase Icon Index */
    If mbNoDraw Then Exit Property
    If IsVisibleItem(Item, Column) Then DrawGrid
e:
End Property
Property Get ItemTag(ByVal Item As Long, Optional ByVal Column As Long) As String
On Local Error Resume Next
    ItemTag = m_Items(Item).Item(Column).Tag
End Property
Property Let ItemTag(ByVal Item As Long, Optional ByVal Column As Long, Value As String)
On Local Error Resume Next
     m_Items(Item).Item(Column).Tag = Value
End Property
Property Get ItemData(ByVal Item As Long) As Long
On Local Error Resume Next
    ItemData = m_Items(Item).data
End Property
Property Let ItemData(ByVal Item As Long, Value As Long)
On Local Error Resume Next
     m_Items(Item).data = Value
End Property

Property Get ColumnCount() As Long
On Local Error Resume Next
    ColumnCount = UBound(m_Cols) + 1
End Property
Property Get ItemCount() As Long
On Error GoTo e
    ItemCount = UBound(m_Items) + 1
e:
End Property
Property Let ItemCount(ByVal Value As Long)
Dim i As Long

    If Value < 1 Then Me.Clear: Exit Property
    ReDim m_Items(Value)
    For i = 0 To Value - 1
      ReDim m_Items(i).Item(ColumnCount - 1)
    Next
End Property

Property Get SelectedItem() As Long: SelectedItem = m_SelRow: End Property
Property Let SelectedItem(ByVal Value As Long)
    If Value < 0 Then Value = -1
    If Value > ItemCount - 1 Then Value = -1
    If m_SelRow <> Value Then
        ChangeSelection Value, m_SelCol
    End If
End Property
Property Get SelectedColumn() As Long: SelectedColumn = m_SelCol: End Property
Property Let SelectedColumn(ByVal Value As Long)
    If Value < 0 Then Value = -1
    If Value > ColumnCount - 1 Then Value = -1
    If m_SelCol <> Value Then
        ChangeSelection m_SelRow, Value
    End If
End Property
Property Get AlignmentItemIcons(Column As Long) As AlignmentConstants
On Error Resume Next
    AlignmentItemIcons = m_Cols(Column).IAlign
End Property
Property Let AlignmentItemIcons(Column As Long, Value As AlignmentConstants)
On Error Resume Next
    m_Cols(Column).IAlign = Value
    If m_iml Then DrawGrid
End Property

Property Get ImagelistHeigth() As Long
    ImagelistHeigth = m_ImgH
End Property
Property Get ImagelistWidth() As Long
    ImagelistWidth = m_ImgW
End Property
Property Let NoDraw(Value As Boolean)
    mbNoDraw = Value
    If Not Value Then Update
End Property




'Private Subs && Properties
'======================================================================================================================
Private Property Get lHeaderH() As Long
    lHeaderH = IIf(m_Header, m_HeaderH * dpi_, 0)
End Property
Private Property Get lGridH() As Long
    lGridH = UserControl.ScaleHeight - lHeaderH
End Property

Public Sub AddColumn(ByVal Text As String, Optional ByVal Width As Long = 100, Optional ByVal Alignment As AlignmentConstants, Optional Fixed As Boolean, Optional ByVal LockEdit As Boolean, Optional ByVal MinWidth As Long)
Dim tHI    As HDITEM
Dim mc     As Long
Dim i      As Long


    Width = Width * dpi_
    mc = ColumnCount
    
    ReDim Preserve m_Cols(mc)
    With m_Cols(mc)
        .Text = Text
        .Width = Width
        .Aling = Alignment
        .Id = mc
        .Fixed = Fixed
        .MinW = MinWidth * dpi_
        .NoEdit = LockEdit
    End With
    
    'i = SendMessage(m_hWnd, HDM_GETITEMCOUNT, 0, ByVal 0)
    tHI.cxy = Width
    tHI.mask = HDI_TEXT Or HDI_WIDTH Or HDI_FORMAT Or HDI_LPARAM
    tHI.fmt = Alignment Or HDF_STRING
    tHI.lParam = mc
    tHI.pszText = Text
    
    Call SendMessage(m_hWnd, HDM_INSERTITEM, mc, tHI)
    m_GridW = m_GridW + (Width)
    If ItemCount And mc < ColumnCount Then
        For i = 0 To ItemCount - 1
            ReDim Preserve m_Items(i).Item(mc)
        Next
    End If
    UpdateScrollH
End Sub

Public Function AddItem(ByVal Text As String, Optional ByVal IconIndex As Long = -1, Optional ByVal ItemData As Long, Optional ByVal ItemTag As String = "") As Long
On Local Error Resume Next
Dim mc   As Long
Dim i   As Long

    mc = ItemCount
    ReDim Preserve m_Items(mc)
    
    With m_Items(mc)
        ReDim .Item(ColumnCount - 1)
        .Item(0).Text = Text
        .Item(0).Icon = IconIndex + 1   '/* Increase Icon Index */
        .Item(0).Tag = ItemTag
        .data = ItemData
    End With
    
    AddItem = mc
    If mbNoDraw Then Exit Function
    UpdateScrollV
    If IsVisibleRow(mc) Then DrawGrid
    
End Function
Public Sub RemoveItem(ByVal Index As Long)
On Local Error Resume Next
Dim j As Integer

    If ItemCount = 0 Or Index > ItemCount - 1 Or ItemCount < 0 Or Index < 0 Then Exit Sub
    
    If ItemCount > 1 Then
         For j = Index To UBound(m_Items) - 1
            LSet m_Items(j) = m_Items(j + 1)
         Next
        ReDim Preserve m_Items(UBound(m_Items) - 1)
    Else
        Erase m_Items
    End If
    
    UpdateScrollV
    If m_SelRow <> -1 Then
        If m_SelRow = Index Then m_SelRow = -1
        If m_SelRow > Index Then m_SelRow = m_SelRow - 1
    End If
    DrawGrid
    
End Sub
Public Sub Clear(Optional ByVal Columns As Boolean)
    
    Erase m_Items
    If Not Columns Then
        UpdateScrollV
        Call ChangeSelection(-1, -1)
        DrawGrid
    Else
        Erase m_Cols
        m_GridW = 0
        Do While SendMessage(m_hWnd, HDM_GETITEMCOUNT, 0, ByVal 0) <> 0
          Call SendMessageByLong(m_hWnd, HDM_DELETEITEM, 0, 0)
        Loop
        Update
    End If
End Sub

Public Sub SortItems(ByVal Column As Long, ByVal Order As JGridSortOrder)
Dim tHI    As HDITEM
Const HDF_SORTUP = &H400
Const HDF_SORTDOWN = &H200

    If ItemCount = 0 Then Exit Sub
    If Column < 0 Or Column > ColumnCount - 1 Then Exit Sub
    
    tHI.mask = HDI_FORMAT
    If mlSortCol <> -1 Then
        If pGetHeaderItemInfo(mlSortCol, tHI) Then
            tHI.fmt = tHI.fmt And Not (HDF_SORTUP Or HDF_SORTDOWN)
            pSetHeaderItemInfo mlSortCol, tHI
        End If
    End If
    
    If pGetHeaderItemInfo(Column, tHI) Then
        Select Case Order
            Case ASC_ORDER: tHI.fmt = (tHI.fmt Or HDF_SORTDOWN) And Not HDF_SORTUP
            Case DESC_ORDER: tHI.fmt = (tHI.fmt Or HDF_SORTUP) And Not HDF_SORTDOWN
            Case Else: tHI.fmt = tHI.fmt And Not (HDF_SORTUP Or HDF_SORTDOWN)
        End Select
        If pSetHeaderItemInfo(Column, tHI) Then mlSortCol = Column
    End If
    
    'SORT NOW
    Call QuickSort(0, UBound(m_Items), Column, Order)
    DrawGrid True

End Sub

Public Sub CreateImageList(Optional Width As Integer = 16, Optional Height As Integer = 16, Optional hBitmap As Long, Optional MaskColor As Long = &HFFFFFFFF)
    If m_iml And mbImlFlag Then ImageList_Destroy m_iml
    m_iml = ImageList_Create(Width * dpi_, Height * dpi_, &H20, 1, 1)
    mbImlFlag = m_iml <> 0
    If m_iml And hBitmap Then
        If (MaskColor <> &HFFFFFFFF) Then
            ImageList_AddMasked m_iml, hBitmap, MaskColor
        Else
            ImageList_Add m_iml, hBitmap, 0
        End If
    End If
    If m_iml Then
        m_ImgW = Width
        m_ImgH = Height
    End If
    
    UpdateSizes
End Sub

Public Sub CreateImageListEx(bvData() As Byte, Optional ByVal IconSize As Long)
On Error GoTo e
Dim IStream     As IUnknown
Dim lW          As Single
Dim lH          As Single
Dim lWX         As Long
Dim Grph        As Long
Dim Bmp         As Long
Dim Nums        As Long
    
    IconSize = IconSize * dpi_
    If m_Bmp Then GdipDisposeImage m_Bmp: m_Bmp = 0
    Call CreateStreamOnHGlobal(bvData(0), 0&, IStream)
    If IStream Is Nothing Then GoTo e
    
    If GdipLoadImageFromStream(IStream, Bmp) = 0 Then
        
        GdipGetImageDimension Bmp, lW, lH
        Nums = lW \ lH
        lWX = IconSize * Nums
        If ((lWX <> lW) Or (IconSize <> lH)) And (lWX <> 0 And IconSize <> 0) Then
            If GdipCreateBitmapFromScan0(lWX, Height, 0&, &HE200B, ByVal 0&, m_Bmp) = 0 Then
                If GdipGetImageGraphicsContext(m_Bmp, Grph) = 0 Then
                    Call GdipSetInterpolationMode(Grph, 7&)
                    Call GdipSetPixelOffsetMode(Grph, 4&)
                    Call GdipDrawImageRectRectI(Grph, Bmp, 0, 0, lWX, IconSize, 0, 0, lW, lH, &H2)
                    m_ImgW = IconSize: m_ImgH = IconSize
                    GdipDeleteGraphics Grph
                End If
            End If
            Call GdipDisposeImage(Bmp)
        Else
            m_ImgW = lH: m_ImgH = lH
            m_Bmp = Bmp
        End If
        UpdateSizes
    End If
    Exit Sub
e:
   If m_Bmp Then GdipDisposeImage m_Bmp: m_Bmp = 0
   m_ImgW = 0: m_ImgH = 0
   UpdateSizes
End Sub

Public Sub Update()
    UpdateScrollV
    UpdateScrollH
    DrawGrid
End Sub
Public Function EditStart(ByVal Item As Long, Subitem As Long) As Boolean
On Local Error Resume Next
Dim Rct   As Rect
Dim Evt0  As Boolean
Dim Evt1  As Boolean
Dim tmp   As String
Dim Px    As Long

    If Not m_Editable Then Exit Function
    
    If Item = -1 Or Subitem = -1 Then Exit Function
    If Item > ItemCount - 1 Or Subitem > ColumnCount - 1 Then Exit Function
    If m_Cols(Subitem).NoEdit Then Exit Function
    
    If mbEditFlag Then EditEnd
    
    e_Row = Item: e_Col = Subitem
    If Not IsCompleteVisibleItem(e_Row, e_Col) Then SetVisibleItem e_Row, e_Col
    
    SendMessage m_hWnd, HDM_GETITEMRECT, e_Col, Rct
    Rct.Left = Rct.Left - ScrollValue(0)
    Rct.Top = ((e_Row * m_RowH) + lHeaderH) - ScrollValue(1)
    Rct.Right = m_Cols(e_Col).Width - IIf(m_GridStyle = 2 Or m_GridStyle = 3, 1 * dpi_, 0)
    Rct.Bottom = m_ItemH
    
    If Rct.Right < 0 Then Exit Function
    
    Evt1 = True
    tmp = m_Items(e_Row).Item(e_Col).Text
    RaiseEvent EditStart(e_Row, e_Col, Rct.Left, Rct.Top, Rct.Right, Rct.Bottom, tmp, e_Ctrl, Evt0, Evt1)
    If Evt0 Then Exit Function

    If (e_Ctrl Is Nothing) Then
        Px = IIf(m_Items(e_Row).Item(e_Col).Icon > 0, m_ImgW + (6 * dpi_), 2 * dpi_)
        Select Case m_Cols(e_Col).IAlign
            Case 0: SetRect Rct, Rct.Left + Px, Rct.Top + ((m_ItemH - c_Edit.Height) \ 2), Rct.Right - Px - (2 * dpi_), c_Edit.Height
            Case 1: SetRect Rct, Rct.Left + (2 * dpi_), Rct.Top + ((m_ItemH - c_Edit.Height) \ 2), Rct.Right - Px - 3, c_Edit.Height
            Case Else
        End Select
        If Rct.Right < 0 Then Exit Function
        
        Evt1 = True
        Set e_Ctrl = c_Edit
    End If
    
    e_hWnd = e_Ctrl.hWnd
    If Evt1 Then MoveWindow e_hWnd, Rct.Left, Rct.Top, Rct.Right, Rct.Bottom, 1
    
    e_Ctrl.Text = tmp
    e_Ctrl.Alignment = m_Cols(e_Col).Aling
    e_Ctrl.SelStart = 0
    e_Ctrl.SelLength = Len(tmp)
    e_Ctrl.ZOrder 0
    
    mbEditFlag = True
    Call DrawGrid
    If e_hWnd Then
         With c_SubClass
            If .Subclass(e_hWnd, , , Me) Then
                .AddMsg e_hWnd, WM_KILLFOCUS, MSG_AFTER
                .AddMsg e_hWnd, WM_CHAR, MSG_BEFORE
                .AddMsg e_hWnd, WM_KEYDOWN, MSG_BEFORE
            End If
        End With
    End If
    Evt0 = True
    RaiseEvent EditShow(e_Row, e_Col, e_Ctrl, Evt0)
    If Evt0 Then e_Ctrl.Visible = True
    e_Ctrl.SetFocus
    EditStart = True
End Function
Public Sub EditEnd()
On Local Error Resume Next
Dim tmp As String
Dim Evt As Boolean

    If Not mbEditFlag Then Exit Sub
    If e_Ctrl Is Nothing Then Exit Sub
    If e_hWnd Then c_SubClass.UnSubclass e_hWnd
    
    e_hWnd = 0
    tmp = e_Ctrl.Text
    e_Ctrl.Visible = False
    SetFocusEx Me.hWnd
       
    RaiseEvent EditEnd(e_Row, e_Col, tmp, e_Ctrl, Evt)
    If Evt Then GoTo e
    If m_Items(e_Row).Item(e_Col).Text <> tmp Then ItemText(e_Row, e_Col) = tmp
e:
    Set e_Ctrl = Nothing
    mbEditFlag = False
    DrawGrid
End Sub

Public Sub Redraw()
    DrawGrid True
End Sub

Public Function ItemFind(ByVal Text As String, Optional ByVal Coincidence As JGridCoincidence = [CC_WHOLE_WORD], Optional ByVal IgnoreCase As Boolean = True, Optional ByVal Column As Long, Optional ByVal Start As Long = 0) As Long
On Local Error GoTo e
Dim i As Long
Dim b As Boolean
    
    ItemFind = -1
    If IgnoreCase Then Text = LCase(Text)
    For i = Start To ItemCount - 1
        If Coincidence = CC_WHOLE_WORD Then
            b = IIf(IgnoreCase, LCase(m_Items(i).Item(Column).Text) = Text, m_Items(i).Item(Column).Text = Text)
        Else
            b = InStr(1, IIf(IgnoreCase, LCase(m_Items(i).Item(Column).Text), m_Items(i).Item(Column).Text), Text) <> 0
        End If
        If b Then ItemFind = i: Exit Function
    Next
e:
End Function

Public Function ItemFindTag(ByVal Tag As String, Optional ByVal Column As Long) As Long
On Local Error GoTo e
Dim i As Long

    ItemFindTag = -1
    For i = 0 To ItemCount - 1
        If m_Items(i).Item(Column).Tag = Tag Then ItemFindTag = i: Exit Function
    Next
e:
End Function

Public Function ItemFindData(data As Long) As Long
On Local Error GoTo e
Dim i As Long
    ItemFindData = -1
    For i = 0 To ItemCount - 1
        If m_Items(i).data = data Then ItemFindData = i: Exit Function
    Next
e:
End Function

Public Function SelectedItemData() As Long
On Local Error GoTo e
    If m_SelRow = -1 Then Exit Function
    SelectedItemData = m_Items(m_SelRow).data
    Exit Function
e:
    SelectedItemData = -1
End Function
Public Function SelectedItemText(Optional ByVal Column As Long) As String
On Error GoTo e
    SelectedItemText = m_Items(m_SelRow).Item(Column).Text
e:
End Function
Public Function SetControlToGrid(Ctrl As Control) As Boolean
On Local Error Resume Next
Dim lWin As Long
    
    lWin = Ctrl.hWnd
    If lWin Then SetControlToGrid = SetParent(lWin, hWnd) <> 0

End Function
Public Function MoveControlToField(ByVal Item As Long, ByVal Column As Long, Ctrl As Control)
Dim lWin As Long
Dim IRct As Rect
    lWin = Ctrl.hWnd
    If lWin Then
        'SetControlToGrid = SetParent(lWin, hwnd) <> 0
        SendMessage m_hWnd, HDM_GETITEMRECT, Column, IRct
        IRct.Left = IRct.Left - ScrollValue(efsHorizontal)
        IRct.Top = ((Item * m_RowH) + lHeaderH) - ScrollValue(efsVertical)
        IRct.Right = m_Cols(Column).Width - IIf(m_GridStyle = 2 Or m_GridStyle = 3, 1 * dpi_, 0)
        IRct.Bottom = m_ItemH
        
        MoveWindow lWin, IRct.Left, IRct.Top, IRct.Right, IRct.Bottom, 1
    End If
    
End Function
Public Sub SetItem(ByVal Item As Long, ByVal Column As Long, ByVal Text As String, Optional ByVal Icon As Long = -1, Optional Tag As String = vbNullString)
On Error GoTo e
     m_Items(Item).Item(Column).Text = Text
     m_Items(Item).Item(Column).Icon = Icon + 1 '/* Increase Icon Index */
     m_Items(Item).Item(Column).Tag = Tag
     If mbNoDraw Then Exit Sub
     If IsVisibleItem(Item, Column) Then DrawGrid
e:
End Sub
Public Sub SetRow(ByVal Item As Long, ParamArray Fields() As Variant)
On Error GoTo e
Dim i As Long
    For i = 0 To UBound(Fields)
        If i > ColumnCount - 1 Then Exit For
        m_Items(Item).Item(i).Text = CStr(Fields(i))
    Next
e:
    If IsVisibleRow(Item) Then Call DrawGrid
End Sub
Public Sub SetRowIcons(ByVal Item As Long, ParamArray Icons() As Variant)
On Error GoTo e
Dim i As Long
    For i = 0 To UBound(Icons)
        If i > ColumnCount - 1 Then Exit For
        m_Items(Item).Item(i).Icon = Val(Icons(i)) + 1  '/* Increase Icon Index */
    Next
e:
    If IsVisibleRow(Item) Then Call DrawGrid
End Sub



'TODO: Private
'----------------------------------------------------------------------------------------------------------------------
Private Sub DrawGrid(Optional ByVal bForce As Boolean, Optional HighlightField As Boolean)
On Local Error Resume Next
Dim lCol    As Long
Dim lRow    As Long
Dim ly      As Long
Dim lx      As Long
Dim lSx     As Long 'Start X
Dim lSCol   As Long 'Start Col
Dim lColW   As Long
Dim dvc     As Long
Dim IRct    As Rect
Dim TRct    As Rect
Dim lPx     As Long
Dim lPx2    As Long
Dim tEvt    As tEventDrawing
Dim PT      As POINTAPI

Dim hGrph As Long

    If Extender.Visible = False And Not bForce Then Exit Sub
    If mbNoDraw Then Exit Sub
    
    UserControl.AutoRedraw = True
    UserControl.Cls

    lCol = 0
    lRow = 0

    lx = -ScrollValue(0)
    ly = -ScrollValue(1)
    dvc = UserControl.Hdc

    ly = ly + lHeaderH
    lSx = lx
    lSCol = -1

    
    If GdipCreateFromHDC(dvc, hGrph) <> 0 Then hGrph = 0
    If HighlightField Then UpdateScrollHitTest
    Do While lRow <= ItemCount - 1 And ly < UserControl.ScaleHeight
        
        If ly + m_RowH > 0 Then '?Visible
            
            
            SetRect IRct, 0, ly, UserControl.ScaleWidth, ly + m_ItemH
            If m_Striped And lRow Mod 2 Then DrawBack dvc, m_StripedColor, IRct
            
            'FullSeleccion
            If m_FullRow Then
            
                lPx = -ScrollValue(efsHorizontal)
                lPx2 = m_GridW - dpi_
                
                '\ Calcular no pintar excesivamente
                If lPx < -(4 * dpi_) Then lPx = -(4 * dpi_): lPx2 = lPx2 + (4 * dpi_) - ScrollValue(efsHorizontal)
                If lPx2 > UserControl.ScaleWidth + (8 * dpi_) Then lPx2 = UserControl.ScaleWidth + (8 * dpi_)
                
                If lRow = m_SelRow And lRow = t_Row Then
                    DrawSelection dvc, lPx, ly, lPx2, m_ItemH, 150 'SEL OVER
                ElseIf lRow = m_SelRow Then
                    DrawSelection dvc, lPx, ly, lPx2, m_ItemH, 100  'SEL
                ElseIf lRow = t_Row Then
                    DrawSelection dvc, lPx, ly, lPx2, m_ItemH, 50   'OVER
                End If
                
            End If
            
            lPx2 = dpi_ \ 2
            
            '?GridLines 0N,1H,2V,3B -> Horizontal
            If m_GridStyle = 1 Or m_GridStyle = 3 Then DrawLine dvc, lx, ly + m_ItemH + lPx2, UserControl.ScaleWidth, ly + m_ItemH + lPx2, m_GridLineColor
            
             Do While lCol < ColumnCount And lx < UserControl.ScaleWidth
             
                lColW = m_Cols(lCol).Width
                If m_GridStyle = 2 Or m_GridStyle = 3 Then lColW = lColW - dpi_
                
                
                If Not (lx + lColW > 0) Then GoTo Next_
     
                
                    If (lSCol = -1) Then
                        lSCol = lCol
                        lSx = lx
                    End If
                    
                    SetRect IRct, lx, ly, lx + lColW, ly + m_ItemH
                   
                    'Selection
                    If Not m_FullRow Then
                        If (lRow = m_SelRow And lCol = m_SelCol) And (lRow = t_Row And lCol = t_Col) Then
                            DrawSelection dvc, lx, ly, lColW, m_ItemH, 150
                        ElseIf (lRow = m_SelRow And lCol = m_SelCol) Then
                            DrawSelection dvc, lx, ly, lColW, m_ItemH, 100
                        ElseIf lRow = t_Row And lCol = t_Col Then
                            DrawSelection dvc, lx, ly, lColW, m_ItemH, 50
                        End If
                    Else
                        If (lRow = m_SelRow And lCol = m_SelCol) And m_FocusRect Then DrawFocusRect dvc, IRct
                        'If (lRow = m_SelRow And lCol = m_SelCol) Then DrawBorder dvc, &H1E2F2B, lx, ly, lColW, m_ItemH
                    End If
                    
                    '?GridLines 0N,1H,2V,3B - > Vertical
                    If m_GridStyle = 2 Or m_GridStyle = 3 Then DrawLine dvc, lx + lColW + lPx2, ly, lx + lColW + lPx2, ly + m_ItemH, m_GridLineColor
                       
                    tEvt = EventDrawingField(lRow, lCol)
                    
                    If tEvt.Back <> -1 Then DrawBack dvc, tEvt.Back, IRct
                    If tEvt.Border <> -1 Then DrawBorder dvc, tEvt.Border, lx, ly, lColW, m_ItemH
                        
                    RaiseEvent ItemDraw(lRow, lCol, dvc, hGrph, IRct.Left, IRct.Top, IRct.Right - IRct.Left, IRct.Bottom - IRct.Top, tEvt.Cancel)
                    If tEvt.Cancel Then GoTo Next_
                    
                    If mbEditFlag Then
                        If e_Row = lRow And e_Col = lCol Then
                            DrawBack dvc, m_eBackColor, IRct
                            DrawBorder dvc, m_eBorderColor, lx, ly, lColW, m_ItemH, m_eBorderSize
                        End If
                    End If
                    
                    '/* Draw Icons */
                    If (m_iml Or m_Bmp) And (m_Items(lRow).Item(lCol).Icon > 0) Then
                        lPx = IIf(m_ImgW + (6 * dpi_) > lColW, lColW - (6 * dpi_), m_ImgW)
                        Select Case m_Cols(lCol).IAlign
                            Case 0 'ICON LEFT
                                    SetRect TRct, (4 * dpi_), ((m_ItemH - m_ImgH) \ 2), lPx, m_ImgH
                            Case 1 'ICON RIGHT
                                    SetRect TRct, lColW - lPx - (4 * dpi_), ((m_ItemH - m_ImgH) \ 2), lPx, m_ImgH
                            Case 2 'ICON CENTER
                                    SetRect TRct, (lColW - m_ImgW) \ 2, ((m_ItemH - m_ImgH) \ 2), lPx, m_ImgH
                                    If lColW - (4 * dpi_) < m_ImgW Then TRct.Right = 0
                        End Select
                        
                        If TRct.Right > 0 Then
                             '/* Decrease Icon Index */
                            If m_iml Then ImageList_DrawEx m_iml, m_Items(lRow).Item(lCol).Icon - 1, dvc, lx + TRct.Left, ly + TRct.Top, TRct.Right, 0, &HFFFFFFFF, &HFF000000, 0
                            If m_Bmp Then GdipDrawImageRectRectI hGrph, m_Bmp, lx + TRct.Left, ly + TRct.Top, TRct.Right, m_ImgH, (m_Items(lRow).Item(lCol).Icon - 1) * m_ImgW, 0, TRct.Right, m_ImgH, &H2, 0&, 0&, 0&
                        End If
                        
                        lPx = m_ImgW + (3 * dpi_)
                    Else
                        lPx = 0
                    End If
                    
                    If LenB(Trim(m_Items(lRow).Item(lCol).Text)) = 0 Then GoTo Next_
                    SetRect TRct, lx + (4 * dpi_) + lPx + tEvt.Ident, ly, lx + lColW - (2 * dpi_), ly + m_ItemH
                    If lPx Then
                        Select Case m_Cols(lCol).IAlign
                            Case 0 'LEFT
                            Case 1 'RIGHT
                                OffsetRect TRct, -lPx - IIf(lPx > 0, 2 * dpi_, 0), 0
                            Case 2 'CENTER
                                SetRect TRct, lx + (4 * dpi_), ly, lx + lColW - (2 * dpi_), ly + m_ItemH
                        End Select
                    End If
                    If TRct.Right < TRct.Left Then GoTo Next_
                    RaiseEvent ItemDrawMeasureText(lRow, lCol, TRct.Left, TRct.Top, TRct.Right, TRct.Bottom)
                    UserControl.ForeColor = IIf(tEvt.Fore <> -1, tEvt.Fore, m_ForeColor)
                    DrawText dvc, m_Items(lRow).Item(lCol).Text, -1, TRct, GetTextFlag(lCol)
Next_:
                lx = lx + m_Cols(lCol).Width
                lCol = lCol + 1
             Loop
             
            '?Reset to Scroll Position
            lCol = lSCol
            lx = lSx
        End If
        
        ly = ly + m_RowH
        lRow = lRow + 1
    Loop
    
    'Completar Rows
    If ly < UserControl.ScaleHeight And Not m_RowH = 0 And m_DrawEmpty And Ambient.UserMode Then
    
        lPx = ly
        lPx2 = dpi_ \ 2
        
        If m_GridStyle = 1 Or m_GridStyle = 3 Then
            Do While ly < UserControl.ScaleHeight
            
                '?StripedGrid
                If lRow Mod 2 And m_Striped Then
                    SetRect IRct, 0, ly, UserControl.ScaleWidth, ly + m_ItemH
                    DrawBack dvc, m_StripedColor, IRct
                End If
                
                '?GridLines 0N,1H,2V,3B -> Horizontal
                If m_GridStyle = 1 Or m_GridStyle = 3 Then DrawLine dvc, 0, ly + m_ItemH + lPx2, UserControl.ScaleWidth, ly + m_ItemH + lPx2, m_GridLineColor
    
                ly = ly + m_RowH
                lRow = lRow + 1
            Loop
        End If
        
        '?GridLines 0N,1H,2V,3B -> Vertical
        If m_GridStyle = 2 Or m_GridStyle = 3 Then
            
            lCol = 0
            lx = -ScrollValue(efsHorizontal)
          
            If (lSCol <> -1) Then
                lCol = lSCol
                lx = lSx
            End If
          
            Do While lCol < ColumnCount And lx < UserControl.ScaleWidth
                
                lColW = m_Cols(lCol).Width - dpi_
                '?Visible Left
                If lx + lColW > 0 Then DrawLine dvc, lx + lColW + (dpi_ \ 2), lPx, lx + lColW + (dpi_ \ 2), lPx + UserControl.ScaleHeight, m_GridLineColor
                        
                lx = lx + m_Cols(lCol).Width
                lCol = lCol + 1
            Loop
          
        End If
    End If

    If hGrph Then Call GdipDeleteGraphics(hGrph)
    
    UserControl.AutoRedraw = False

    
End Sub
Private Function EventDrawingField(lRow As Long, lCol As Long) As tEventDrawing

    With EventDrawingField
        .Back = -1
        .Border = -1
        .Fore = -1
        .Cancel = False
        .Ident = 0
         RaiseEvent ItemDrawData(lRow, lCol, .Fore, .Back, .Border, .Ident)
         .Ident = .Ident * dpi_
    End With
   
End Function
Private Function GetTextFlag(Col As Long) As Long
    'VerticalCenter-SingleLine-WordElipsis
    GetTextFlag = &H4 Or &H20 Or &H40000
    Select Case m_Cols(Col).Aling
        Case 1: GetTextFlag = GetTextFlag Or &H2
        Case 2: GetTextFlag = GetTextFlag Or &H1
    End Select
    
End Function

Private Function SysColor(oColor As Long) As Long: OleTranslateColor2 oColor, 0, SysColor: End Function

Private Function mvCreateHeader() As Boolean
Dim wStyle      As Long

    wStyle = &H40000000 Or &H10000000 Or HDS_HORZ Or HDS_BUTTONS Or HDS_HOTTRACK ' [WS_CHILD|WS_VISIBLE]
    'wStyle = wStyle Or HDS_HOTTRACK
    'wStyle = wStyle Or HDS_DRAGDROP
    'wStyle = wStyle Or HDS_BUTTONS
    
    m_hWnd = CreateWindowEx(0, "SysHeader32", "", wStyle, 0, 0, UserControl.ScaleWidth, m_HeaderH * dpi_, hWnd, 0, App.hInstance, 0)
    
    If m_hWnd Then
    
         Const CCM_FIRST              As Long = &H2000
         Const CCM_SETUNICODEFORMAT   As Long = (CCM_FIRST + 5)
         Const CCM_GETUNICODEFORMAT   As Long = (CCM_FIRST + 6)

        'SendMessage m_hWnd, CCM_SETUNICODEFORMAT, 1&, ByVal 0&
        
        ShowWindow m_hWnd, Abs(m_Header)
        SendMessage m_hWnd, &H30, m_IFont.hFont, 0&
        SendMessage m_hWnd, &H2000 + 5, 1&, ByVal 0&
        
        SetTextColor GetDC(m_hWnd), vbRed
        With c_SubClass
            If .Subclass(m_hWnd, , , Me) Then
                .AddMsg m_hWnd, WM_PAINT, MSG_BEFORE
                .AddMsg m_hWnd, WM_LBUTTONUP, MSG_BEFORE
                .AddMsg m_hWnd, WM_LBUTTONDOWN, MSG_BEFORE
                .AddMsg m_hWnd, WM_SIZE, MSG_BEFORE
                .AddMsg m_hWnd, WM_ERASEBKGND, MSG_AFTER
            End If
        End With
        
    End If
End Function
Private Sub mvDestroyHeader()
    If m_hWnd Then
        c_SubClass.UnSubclass m_hWnd
        ShowWindow m_hWnd, 0
        DestroyWindow m_hWnd
        m_hWnd = 0
    End If
End Sub
Private Sub MoveHeader(Optional ByVal lLeft As Long = -1, Optional ByVal lWidth As Long = -1, Optional ByVal lHeight = -1)
    If lLeft = -1 Then lLeft = -ScrollValue(0)
    If lWidth = -1 Then lWidth = m_GridW + (5 * dpi_)
    If lHeight = -1 Then lHeight = m_HeaderH * dpi_
    Call MoveWindow(m_hWnd, lLeft, 0, lWidth, lHeight, 1)
End Sub

Private Sub RedrawHeader()
    RedrawWindow m_hWnd, ByVal 0&, ByVal 0&, &H1
End Sub

Private Function ScrollValue(eBar As EFSScrollBarConstants) As Long
    ScrollValue = IIf(c_Scroll.Visible(eBar), c_Scroll.Value(eBar), 0)
End Function

Private Function GetRowFromY(ByVal Y As Long) As Long
    Y = Y + ScrollValue(efsVertical) - lHeaderH
    GetRowFromY = Y \ m_RowH
    If GetRowFromY >= ItemCount Then GetRowFromY = -1
End Function

Private Function GetColFromX(ByVal x As Long) As Long
Dim iCol    As Long
Dim tHDI    As HDHITTESTINFO

    x = x + ScrollValue(0) + 8
    tHDI.PT.x = x
    Call SendMessage(m_hWnd, HDM_HITTEST, 0, tHDI)
    GetColFromX = tHDI.iItem
    
End Function


Private Sub UpdateSizes()
Dim Th As Integer
Dim Px As Long
    
    Px = 4 * dpi_
    Th = UserControl.TextHeight("ÀJq")
    If Not c_Edit Is Nothing Then c_Edit.Height = Th
    
    If Th + Px > m_ItemH Then m_ItemH = Th + Px
    If m_ImgH + Px > m_ItemH Then m_ItemH = m_ImgH + Px
    m_RowH = m_ItemH
    
    If m_GridStyle = 1 Or m_GridStyle = 3 Then m_RowH = m_RowH + (1 * dpi_)
  
End Sub

Private Sub UpdateScrollV()
On Local Error Resume Next
Dim lHeight     As Long
Dim lProportion As Long
Dim ly          As Long
Dim bFlag       As Boolean

    bFlag = c_Scroll.Visible(1)
    ly = lGridH
    lHeight = ((ItemCount * m_RowH) + 5) - ly
    
    If (lHeight > 0) Then
      lProportion = lHeight \ (ly + 1)
      c_Scroll.LargeChange(1) = lHeight \ lProportion
      c_Scroll.Max(1) = lHeight '+ 1
      c_Scroll.Visible(1) = True
    Else
      c_Scroll.Visible(1) = False
    End If
    If bFlag <> c_Scroll.Visible(1) Then UpdateScrollH
End Sub

Private Sub UpdateScrollH()
On Local Error Resume Next
Dim lWidth      As Long
Dim lProportion As Long
Dim bFlag       As Boolean
    
    bFlag = c_Scroll.Visible(0)
    lWidth = m_GridW - (UserControl.ScaleWidth - (5 * dpi_))
    MoveHeader 0, UserControl.ScaleWidth + (5 * dpi_)
    
    If (lWidth > 0) Then
        lProportion = lWidth \ (UserControl.ScaleWidth) + 1
        c_Scroll.LargeChange(0) = lWidth \ lProportion
        If c_Scroll.LargeChange(0) < (20 * dpi_) Then c_Scroll.LargeChange(0) = (20 * dpi_)
        c_Scroll.Max(0) = lWidth
        c_Scroll.Visible(0) = True
        MoveHeader -ScrollValue(0), m_GridW + (12 * dpi_)
    Else
        c_Scroll.Visible(0) = False:
        MoveHeader 0, UserControl.ScaleWidth + (5 * dpi_)
    End If
    If bFlag <> c_Scroll.Visible(0) Then UpdateScrollV
    
End Sub

Private Function IsVisibleRow(ByVal eRow As Long) As Boolean
On Error Resume Next
Dim Y As Long
    If c_Scroll.Visible(1) = False Then IsVisibleRow = True: Exit Function
    Y = (eRow * m_RowH) - ScrollValue(1)
    IsVisibleRow = (Y + m_ItemH > 0) And Y <= lGridH
End Function

Private Function IsVisibleItem(eRow As Long, ByVal eCol As Long) As Boolean
On Error Resume Next
Dim Y       As Long
Dim x       As Long
Dim bRow    As Boolean
Dim bCol    As Boolean
Dim Rct     As Rect

    SendMessage m_hWnd, HDM_GETITEMRECT, eCol, Rct
    Y = (eRow * m_RowH) - ScrollValue(1)
    x = Rct.Left - (ScrollValue(0))
    
    bRow = (Y + m_ItemH > 0) And Y <= lGridH
    bCol = (x + m_Cols(eCol).Width) >= 0 And x <= UserControl.ScaleWidth
    IsVisibleItem = bRow And bCol

End Function

Private Function IsCompleteVisibleItem(eRow As Long, eCol As Long) As Boolean
On Local Error Resume Next
Dim Y       As Long
Dim x       As Long
Dim bRow    As Boolean
Dim bCol    As Boolean
Dim Rct     As Rect

    
    SendMessage m_hWnd, HDM_GETITEMRECT, eCol, Rct
    Y = (eRow * m_RowH) - ScrollValue(1)
    x = Rct.Left - (ScrollValue(0))
    
    bRow = (Y >= 0) And (Y + m_ItemH <= lGridH)
    bCol = (x >= 0) And (x + m_Cols(eCol).Width <= UserControl.ScaleWidth) Or m_FullRow
    IsCompleteVisibleItem = bRow And bCol
    
End Function

Private Sub ChangeSelection(eRow As Long, eCol As Long)

    If eRow = m_SelRow And eCol = m_SelCol Then
        If Not IsCompleteVisibleItem(eRow, eCol) Then SetVisibleItem eRow, eCol
        Exit Sub
    End If
    
    m_SelRow = eRow
    m_SelCol = eCol
    
    If m_SelRow = -1 Or m_SelCol = -1 Then DrawGrid: GoTo Evt
    If Not IsCompleteVisibleItem(eRow, eCol) Then
        SetVisibleItem eRow, eCol
    Else
        DrawGrid
    End If
    
Evt:
    RaiseEvent SelectionChanged(eRow, eCol)
End Sub

Private Sub SetVisibleItem(eRow As Long, eCol As Long)
On Error GoTo e
Dim lx      As Long
Dim ly      As Long
Dim Rct     As Rect
    
    
    If eRow = -1 Or eCol = -1 Then Exit Sub
    
    SendMessage m_hWnd, HDM_GETITEMRECT, eCol, Rct
    ly = eRow * m_RowH
    lx = Rct.Left
    
    '?Vertical
    If (ly + m_RowH) - lGridH > ScrollValue(1) Then
        c_Scroll.Value(1) = ((ly + m_RowH) + 2) - lGridH
    ElseIf ly < ScrollValue(1) Then
        c_Scroll.Value(1) = ly
    End If
    
    '?Horizantal
    If lx + m_Cols(eCol).Width > UserControl.ScaleWidth + ScrollValue(0) Then
        c_Scroll.Value(0) = ((lx + m_Cols(eCol).Width) + 20) - UserControl.ScaleWidth '-> Right
    ElseIf (lx - ScrollValue(0)) < 0 Then
        c_Scroll.Value(efsHorizontal) = lx - 5
    End If
    
e:
    DrawGrid
End Sub

Private Function IsVisibleColumnForDraw(lCol As Long, lx As Long) As Boolean
      IsVisibleColumnForDraw = (lx - (ScrollValue(0)) + m_Cols(lCol).Width) >= 0 And (lx - (ScrollValue(0)) <= UserControl.ScaleWidth)
End Function

Private Sub UpdateScrollHitTest()
Dim PT As POINTAPI
    GetCursorPos PT
    If WindowFromPoint(PT.x, PT.Y) = hWnd Then
        ScreenToClient hWnd, PT
        t_Row = GetRowFromY(PT.Y)
        t_Col = GetColFromX(PT.x)
        If t_Row <> -1 And t_Col = -1 And m_FullRow Then t_Row = -1
    Else
        t_Row = -1: t_Col = -1
    End If
End Sub

Private Sub SetHeaderWidth(eCol As Long, ByVal lWidth As Long)
Dim tHI As HDITEM
    
    tHI.mask = HDI_WIDTH
    Call pGetHeaderItemInfo(eCol, tHI)
    'If tHI.cxy <> lWidth Then
    tHI.cxy = lWidth
    If (pSetHeaderItemInfo(eCol, tHI)) Then
        'RaiseEvent ColumnSizeChanged(Index, Value)
    End If
    'End If
End Sub

Private Function pSetHeaderItemInfo(ByVal lCol As Long, tHI As HDITEM) As Boolean
      If Not (SendMessage(m_hWnd, HDM_SETITEM, lCol, tHI) = 0) Then
         pSetHeaderItemInfo = True
      End If
End Function
Private Function pGetHeaderItemInfo(ByVal lCol As Long, tHI As HDITEM) As Boolean
      If Not (SendMessage(m_hWnd, HDM_GETITEM, lCol, tHI) = 0) Then
         pGetHeaderItemInfo = True
      End If
End Function

Private Sub DrawSelection(dvc As Long, x As Long, Y As Long, W As Long, H As Long, Intenc As Long)
Dim hPen        As Long
Dim hBrush      As Long
Dim OldBrush    As Long
Dim OldPen      As Long
Dim lColor      As Long

    '150 SEL - OVER
    '100 SEL
    '50  OVER
    
    If m_AlphaBlend = False Then GoTo e
    
    lColor = BlendColor(m_SelColor, UserControl.BackColor, Intenc)
    hPen = CreatePen(0, 1 * dpi_, lColor)
    hBrush = CreateSolidBrush(BlendColor(lColor, UserControl.BackColor, 115))

    OldBrush = SelectObject(dvc, hBrush)
    OldPen = SelectObject(dvc, hPen)
    
    RoundRect dvc, x, Y, x + W, Y + H, 0, 0
    Call SelectObject(dvc, OldPen)
    Call SelectObject(dvc, OldBrush)
    
    DeleteObject hPen
    DeleteObject hBrush
    
    Exit Sub
e:
Dim Rct As Rect
    
    SetRect Rct, x, Y, x + W, Y + H
    hBrush = CreateSolidBrush(BlendColor(m_SelColor, UserControl.BackColor, Intenc))
    Call FillRect(dvc, Rct, hBrush)
    Call DeleteObject(hBrush)
    'DrawFocusRect dvc, Rct
End Sub

Private Function BlendColor(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
Dim clrFore(3)      As Byte
Dim clrBack(3)      As Byte

    OleTranslateColor clrFirst, 0, VarPtr(clrFore(0))
    OleTranslateColor clrSecond, 0, VarPtr(clrBack(0))
    
    clrFore(0) = (clrFore(0) * lAlpha + clrBack(0) * (255 - lAlpha)) / 255
    clrFore(1) = (clrFore(1) * lAlpha + clrBack(1) * (255 - lAlpha)) / 255
    clrFore(2) = (clrFore(2) * lAlpha + clrBack(2) * (255 - lAlpha)) / 255
    CopyMemory BlendColor, clrFore(0), 4
End Function
Private Sub DrawBorder(dvc As Long, Color As Long, x As Long, Y As Long, W As Long, H As Long, Optional ByVal PenSize As Long = 1)
Dim hPen As Long
Dim Px1  As Long
    
    Px1 = dpi_ \ 2
    hPen = CreatePen(0, PenSize * dpi_, SysColor(Color))
    Call SelectObject(dvc, hPen)
    RoundRect dvc, x + Px1, Y + Px1, x + (W - Px1), Y + (H - Px1), 0, 0
    DeleteObject hPen
    
End Sub
Private Sub DrawBack(lpDC As Long, Color As Long, Rct As Rect)
Dim hBrush  As Long

    hBrush = CreateSolidBrush(SysColor(Color))
    Call FillRect(lpDC, Rct, hBrush)
    Call DeleteObject(hBrush)
    
End Sub
Private Sub DrawLine(lpDC As Long, x As Long, Y As Long, x2 As Long, y2 As Long, Color As Long)
Dim PT      As POINTAPI
Dim hPen    As Long
Dim hPenOld As Long

    hPen = CreatePen(0, 1 * dpi_, SysColor(Color))
    hPenOld = SelectObject(lpDC, hPen)
    Call MoveToEx(lpDC, x, Y, PT)
    Call LineTo(lpDC, x2, y2)
    Call SelectObject(lpDC, hPenOld)
    Call DeleteObject(hPen)
End Sub


' TODO: Quick SortItems: ELIHU
Private Sub QuickSort(inLow As Long, inHi As Long, Col As Long, Ord As JGridSortOrder)
Dim pivot   As Variant
Dim tmpSwap As tItem
Dim tmpLow  As Long
Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

    pivot = GetVar((inLow + inHi) \ 2, Col)
    While (tmpLow <= tmpHi)
        If Ord = ASC_ORDER Then
            While (GetVar(tmpLow, Col) < pivot And tmpLow < inHi)
                tmpLow = tmpLow + 1
            Wend
            While (pivot < GetVar(tmpHi, Col) And tmpHi > inLow)
                tmpHi = tmpHi - 1
            Wend
        Else
            While (GetVar(tmpLow, Col) > pivot And tmpLow < inHi)
                tmpLow = tmpLow + 1
            Wend
            While (pivot > GetVar(tmpHi, Col) And tmpHi > inLow)
                tmpHi = tmpHi - 1
            Wend
        End If
        If (tmpLow <= tmpHi) Then
            tmpSwap = m_Items(tmpLow)
            m_Items(tmpLow) = m_Items(tmpHi)
            m_Items(tmpHi) = tmpSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Wend

  If (inLow < tmpHi) Then QuickSort inLow, tmpHi, Col, Ord
  If (tmpLow < inHi) Then QuickSort tmpLow, inHi, Col, Ord
End Sub

Private Function GetVar(Item As Long, Col As Long) As Variant
    If IsNumeric(m_Items(Item).Item(Col).Text) Then
        GetVar = Val(m_Items(Item).Item(Col).Text)
    Else
        If IsDate(m_Items(Item).Item(Col).Text) Then GetVar = CDate(m_Items(Item).Item(Col).Text) Else GetVar = m_Items(Item).Item(Col).Text
    End If
End Function

Private Sub CustomDrawHeader()
Dim PS      As PAINTSTRUCT
Dim HDHII   As HDHITTESTINFO
Dim IRct    As Rect
Dim TRct    As Rect
Dim lIndex  As Long
Dim TColor  As Long
Dim lBState As JGridMouseState
Dim Evt0    As Boolean
Dim i       As Long


    Call BeginPaint(m_hWnd, PS)
    SetBkMode PS.Hdc, 1
    'SelectObject PS.Hdc, GetCurrentObject(UserControl.Hdc, 6&)
    SelectObject PS.Hdc, m_IFont.hFont
    RaiseEvent HeaderBkgndDraw(PS.Hdc, UserControl.ScaleWidth + (20 * dpi_), lHeaderH)
    
    Call GetCursorPos(HDHII.PT)
    Evt0 = (WindowFromPoint(HDHII.PT.x, HDHII.PT.Y) = m_hWnd)
    Call ScreenToClient(m_hWnd, HDHII.PT)
    Call SendMessage(m_hWnd, HDM_HITTEST, 0, HDHII)
    
    lIndex = HDHII.iItem
    If mlHdrBtn = -1 And m_bmdhFlag Then mlHdrBtn = lIndex
    If Not Evt0 Then lIndex = -1
    
    For i = 0 To Me.ColumnCount - 1
    
        SendMessage m_hWnd, HDM_GETITEMRECT, i, IRct
        SetRect TRct, IRct.Left + (4 * dpi_), 0, IRct.Right - (8 * dpi_), IRct.Bottom
        
        '/Visible?
        'If Not IsVisibleColumnForDraw(i, IRct.Left) Then GoTo lNext

        lBState = 0
        If lIndex = i And m_bmdhFlag And (mlHdrBtn = i) Then lBState = HBS_DOWN: OffsetRect TRct, 1 * dpi_, 1 * dpi_
        If lIndex = i And Not m_bmdhFlag Then lBState = HBS_HOT
        RaiseEvent HeaderColumnDraw(i, PS.Hdc, IRct.Left, IRct.Right - IRct.Left, IRct.Bottom, lBState)
        
        If TRct.Right < TRct.Left Then TRct.Right = TRct.Left
        TColor = vbWindowText
        Evt0 = False
        RaiseEvent HeaderColumnTextDraw(i, PS.Hdc, TRct.Left, TRct.Top, TRct.Right, TRct.Bottom, TColor, lBState, Evt0)
        If Not Evt0 Then
            SetTextColor PS.Hdc, TColor
            DrawText PS.Hdc, m_Cols(i).Text, -1, TRct, GetTextFlag(i)
            If mlSortCol = i And (TRct.Right - TRct.Left) > 3 Then
                RenderArrow PS.Hdc, IRct.Right - (9 * dpi_), (IRct.Bottom - 6) \ 2, 5, ppSortBitmap(i), TColor
            End If
        End If
lNext:
    Next
    Call EndPaint(m_hWnd, PS)
End Sub
Private Sub RenderArrow(dvc As Long, Optional ByVal x1 As Long, Optional ByVal y1 As Long, Optional ByVal aSize As Long = 3, Optional ArrowDir As Long, Optional Color As Long)
Dim PT(2)       As POINTAPI
Dim hPen        As Long
Dim hBrush      As Long
Dim OldBrush    As Long
Dim OldPen      As Long


    aSize = aSize * dpi_
    If aSize Mod 2 = 0 Then aSize = aSize + 1
    aSize = aSize - 1
    
    hPen = CreatePen(6, 1 * dpi_, Color)
    hBrush = CreateSolidBrush(Color)

    OldBrush = SelectObject(dvc, hBrush)
    OldPen = SelectObject(dvc, hPen)
    
    Select Case ArrowDir
        Case 0 'ToDown
            PT(0).x = x1:                   PT(0).Y = y1
            PT(1).x = x1 + aSize:           PT(1).Y = y1
            PT(2).x = x1 + (aSize \ 2):     PT(2).Y = y1 + (aSize \ 2)
        Case 1 'ToUp
            PT(0).x = x1 + (aSize \ 2):     PT(0).Y = y1
            PT(1).x = x1:                   PT(1).Y = y1 + (aSize \ 2)
            PT(2).x = x1 + aSize:           PT(2).Y = y1 + (aSize \ 2)
    End Select
    Polygon dvc, PT(0), 3
    
    Call SelectObject(dvc, OldPen)
    Call SelectObject(dvc, OldBrush)
    
    DeleteObject hPen
    DeleteObject hBrush
End Sub
Private Function ppSortBitmap(ByVal Col As Long) As JGridSortOrder
Dim tHI As HDITEM
    ppSortBitmap = -1
    tHI.mask = HDI_FORMAT
   If (pGetHeaderItemInfo(Col, tHI)) Then
      If (tHI.fmt And &H200) = &H200 Then ppSortBitmap = ASC_ORDER
      If (tHI.fmt And &H400) = &H400 Then ppSortBitmap = DESC_ORDER
   End If
End Function

Private Function mvGetWindowsDPI() As Double
Dim Hdc  As Long
Dim lPx  As Double

    Hdc = GetDC(0)
    lPx = CDbl(GetDeviceCaps(Hdc, 88))
    ReleaseDC 0, Hdc
    If (lPx = 0) Then mvGetWindowsDPI = 1# Else mvGetWindowsDPI = lPx / 96#
    
End Function


'- ordinal #1
Private Sub WndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lParamUser As Long)
On Error Resume Next
Dim Evt As Boolean

    Select Case hWnd
        Case UserControl.hWnd
            Select Case uMsg
                Case WM_NOTIFY
                        
                        Dim tNMH    As NMHDR
                        Dim tHDN    As NMHEADER
                        Dim lHDI()  As Long

                        'CopyMemory tNMH, ByVal lParam, LenB(tNMH)
                        CopyMemory tHDN, ByVal lParam, Len(tHDN)
                        
                        ReDim lHDI(1)
                        Select Case tHDN.Hdr.code
                            Case HDN_BEGINTRACK
                                If mbEditFlag Then EditEnd
                                
                                CopyMemory lHDI(0), ByVal tHDN.lPtrHDItem, 8
                                If m_Cols(tHDN.iItem).Fixed Then lReturn = 1: bHandled = True: Exit Sub
                                m_bmdhFlag = False
                                RaiseEvent ColumnSizeChangeStart(tHDN.iItem, lHDI(1), Evt)
                                If Evt Then lReturn = 1: bHandled = True
                                
                            Case HDN_TRACK
                                CopyMemory lHDI(0), ByVal tHDN.lPtrHDItem, 8
                                
                                RaiseEvent ColumnSizeChanging(tHDN.iItem, lHDI(1), Evt)
                                
                                If Evt Then
                                    lReturn = 1: bHandled = True
                                    lParam = lHDI(1)
                                    SendMessage m_hWnd, WM_PAINT, 0&, 0&
                                End If
                                
                                'm_GridW = (m_GridW - m_Cols(tHDN.iItem).Width) + lHDI(1)
                                'm_Cols(tHDN.iItem).Width = lHDI(1)
                                'UpdateScrollH
                                'Call DrawGrid(True)
                                'SetHeaderWidth tHDN.iItem, lHDI(1)
                                'DoEvents
                                
                            Case HDN_ENDTRACK
                            
                                CopyMemory lHDI(0), ByVal tHDN.lPtrHDItem, 8
                            
                                m_GridW = (m_GridW - m_Cols(tHDN.iItem).Width) + lHDI(1)
                                m_Cols(tHDN.iItem).Width = lHDI(1)
                                UpdateScrollH
                                DrawGrid
                                RaiseEvent ColumnSizeChanged(tHDN.iItem, lHDI(1))
                                If m_Cols(tHDN.iItem).MinW Then
                                    If lHDI(1) < m_Cols(tHDN.iItem).MinW Then
                                        Me.ColumnWidth(tHDN.iItem) = Me.ColumnMinWidth(tHDN.iItem)
                                        lParam = Me.ColumnMinWidth(tHDN.iItem)
                                        lReturn = 1
                                    End If
                                End If
                                
                            Case HDN_DIVIDERDBLCLICK
                                If mbEditFlag Then EditEnd
                                RaiseEvent ColumnDividerDblClick(tHDN.iItem)
                            Case HDN_ITEMCLICK
                                If mbEditFlag Then EditEnd
                                RaiseEvent ColumnClick(tHDN.iItem)
                            Case HDN_ITEMDBLCLICK
                                RaiseEvent ColumnDblClick(tHDN.iItem)
                            Case HDN_BEGINDRAG
                                If mbEditFlag Then EditEnd
                                'RaiseEvent ColumnDragStart(tHDN.iItem, Evt)
                                'If Evt Then lReturn = 1: bHandled = True
                            Case HDN_ENDDRAG
                                ReDim lHDI(8)
                                CopyMemory lHDI(0), ByVal tHDN.lPtrHDItem, 36
                                'Debug.Print "Drag "; tHDN.iItem; vbTab; lHDI(8)
                                'If (lHDI(8) > -1) Then
                                    'PostMessage m_hWnd, UM_ENDDRAG, tHDN.iItem, lHDI(8)
                                'End If
                        End Select
                        
                Case WM_MOUSELEAVE
                
                    If t_Row <> -1 Or t_Col <> -1 Then
                        t_Col = -1: t_Row = -1
                        DrawGrid
                    End If
                    mbTrack = False
                    RaiseEvent MouseExit

                Case WM_NCPAINT
                    
                    If UserControl.BorderStyle = 0 Then Exit Sub
                    Dim Rct As Rect
                    Dim DC As Long
                    Dim BZ As Long
                    Dim Px1  As Long
    
                    Px1 = dpi_ \ 2
                    DC = GetWindowDC(hWnd)
                    
                    GetWindowRect hWnd, Rct
                    Rct.Right = (Rct.Right - Rct.Left) - Px1
                    Rct.Bottom = (Rct.Bottom - Rct.Top) - Px1
                    Rct.Left = Px1
                    Rct.Top = Px1
                    
                    BZ = GetSystemMetrics(6)
                    ExcludeClipRect DC, BZ + (1 * dpi_), BZ + (1 * dpi_), Rct.Right - (BZ + (1 * dpi_)), Rct.Bottom - (BZ + (1 * dpi_))

                    Dim hPen        As Long
                    Dim OldPen      As Long
                    Dim hBrush      As Long
                    Dim OldBrush    As Long
            
                    hPen = CreatePen(0, 1 * dpi_, SysColor(m_BorderColor))
                    hBrush = CreateSolidBrush(SysColor(UserControl.BackColor))
                    
                    OldPen = SelectObject(DC, hPen)
                    OldBrush = SelectObject(DC, hBrush)
                    
                    Rectangle DC, Rct.Left, Rct.Top, Rct.Right, Rct.Bottom
                    Call SelectObject(DC, OldPen)
                    Call SelectObject(DC, OldBrush)
                    
                    DeleteObject hPen
                    DeleteObject hBrush
                    
                    ReleaseDC hWnd, DC
                    
            End Select
            
        Case m_hWnd
            'm_Header.WndProc bBefore, bHandled, lReturn, hWnd, uMsg, wParam, lParam, lParamUser

            Select Case uMsg
                Case WM_PAINT
                
                    If m_CustomDraw Then Call CustomDrawHeader
                    
                Case WM_ERASEBKGND
                    Dim WRct As Rect
                    
                    GetClientRect m_hWnd, WRct
                    If m_CustomDraw Then RaiseEvent HeaderBkgndDraw(wParam, WRct.Right, WRct.Bottom)
                    lReturn = 1: bHandled = True
                    
                Case WM_SIZE
                
                     If Not mbResize Then
                        GetWindowRect m_hWnd, Rct
                        Rct.Right = Rct.Right - Rct.Left
                        If Rct.Right < UserControl.ScaleWidth Then MoveHeader 0, UserControl.ScaleWidth + (5 * dpi_)
                     End If
                     
                Case WM_LBUTTONDOWN, WM_LBUTTONUP
                    
                    m_bmdhFlag = uMsg = WM_LBUTTONDOWN
                    If uMsg = WM_LBUTTONUP And mlHdrBtn <> -1 Then
                        mlHdrBtn = -1
                        If m_CustomDraw Then RedrawHeader
                    End If
  
            End Select
            
        Case e_hWnd
            Select Case uMsg
                Case WM_KILLFOCUS
                    If mbEditFlag Then EditEnd
                Case WM_CHAR
                    If wParam = 27 Then
                    ElseIf wParam = 13 Then
                        If mbEditFlag Then EditEnd
                        lReturn = 0
                        bHandled = True
                    ElseIf wParam = 9 Then
                    End If
                Case WM_NCPAINT
                Case WM_KEYDOWN
                    'Debug.Print wParam
            End Select
    Case Else
            Debug.Print "SASS"
    End Select
 
End Sub
