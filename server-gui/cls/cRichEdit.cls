VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRichEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module      : cRichEdit
' DateTime    : 02/05/2021
' Author      : J. Elihu
'---------------------------------------------------------------------------------------
Option Explicit

Private Const CFM_BACKCOLOR     As Long = &H4000000
Private Const CFM_COLOR         As Long = &H40000000
Private Const CFM_BOLD          As Long = &H1
Private Const CFM_ITALIC        As Long = &H2
Private Const CFM_UNDERLINE     As Long = &H4
Private Const CFM_STRIKEOUT     As Long = &H8
Private Const CFM_PROTECTED     As Long = &H10
Private Const CFM_HIDDEN        As Long = &H100
Private Const CFM_LINK          As Long = &H20
Private Const CFM_SIZE          As Long = &H80000000
Private Const CFM_FACE          As Long = &H20000000
Private Const CFM_OFFSET        As Long = &H10000000
Private Const CFM_CHARSET       As Long = &H8000000

Private Const CFE_AUTOBACKCOLOR As Long = CFM_BACKCOLOR
Private Const CFE_AUTOCOLOR     As Long = CFM_COLOR
Private Const CFE_BOLD          As Long = &H1
Private Const CFE_ITALIC        As Long = &H2
Private Const CFE_UNDERLINE     As Long = &H4
Private Const CFE_STRIKEOUT     As Long = &H8
Private Const CFE_PROTECTED     As Long = &H10
Private Const CFE_HIDDEN        As Long = &H100
Private Const CFE_LINK          As Long = &H20
Private Const CFE_SUBSCRIPT     As Long = &H10000
Private Const CFE_SUPERSCRIPT   As Long = &H20000

Private Const LF_FACESIZE       As Long = 32&

Private Const GWL_STYLE         As Long = (-16)
Private Const GWL_EXSTYLE       As Long = (-20)

Private Const SCF_DEFAULT       As Long = &H0
Private Const SCF_SELECTION     As Long = &H1

Enum eScrollBars
  None
  Vertical
  Horizontal
  Both
End Enum

Enum eBorderStyle
  [None ]
  [Fixed Single]
End Enum

Enum RTFFindOptions
  RTFFindWholeword = &H2
  RTFFindMatchcase = &H4
  RTFFindNoHighlight = &H8
End Enum

Private Type POINTAPI
  x     As Long
  y     As Long
End Type
Private Type SETTEXTEX
  Flags As Long
  CP    As Long
End Type
Private Type CHARRANGE
  Min   As Long
  Max   As Long
End Type

Private Type CHARFORMAT2W
  cbSize          As Long
  dwMask          As Long
  dwEffects       As Long
  YHeight         As Long
  YOffset         As Long
  TextColor       As Long
  Charset         As Byte
  PitchAndFamily  As Byte
  FaceName(0 To ((LF_FACESIZE * 2) - 1)) As Byte
  Weight          As Integer
  Spacing         As Integer
  BackColor       As Long
  LCID            As Long
  dwReserved      As Long
  Style           As Integer
  Kerning         As Integer
  UnderlineType   As Byte
  Animation       As Byte
  RevAuthor       As Byte
  UnderlineColor  As Byte
End Type

Private Type REFINDTEXTEX
  CHARRANGE      As CHARRANGE
  lpstrText      As Long
  CharRangeText  As CHARRANGE
End Type

Private Const GTL_USECRLF = 1
Private Const GTL_PRECISE = 2
Private Const GTL_NUMCHARS = 8
Private Const CP_UNICODE = 1200
Private Const GT_USECRLF = 1

Private Type GETTEXTEX
  cb              As Long
  Flags           As Long
  CodePage        As Integer
  lpDefaultChar   As String
  lpUsedDefChar   As Boolean
End Type
Private Type GETTEXTLENGTHEX
  Flags           As Long       ' /* flags (see GTL_XXX defines)              */
  CodePage        As Long       ' /* code page for translation (CP_ACP for default,                */
End Type

Private Const EM_REPLACESEL     As Long = &HC2
Private Const EM_SETSEL         As Long = &HB1

Private Const ES_READONLY       As Long = &H800&
Private Const ES_MULTILINE      As Long = &H4&
Private Const ES_NOHIDESEL      As Long = &H100
Private Const ES_PASSWORD       As Long = &H20
Private Const ES_AUTOHSCROLL    As Long = &H80
Private Const ES_AUTOVSCROLL    As Long = &H40

Private Const WS_VISIBLE        As Long = &H10000000
Private Const WS_CHILD          As Long = &H40000000
Private Const WS_VSCROLL        As Long = &H200000
Private Const WS_BORDER         As Long = &H800000
Private Const WS_HSCROLL        As Long = &H100000

Private Const WM_SETFONT        As Long = &H30
Private Const WM_SETFOCUS       As Long = &H7
Private Const WM_KILLFOCUS      As Long = &H8
Private Const WM_ENABLE         As Long = &HA
Private Const WM_THEMECHANGED   As Long = &H31A
Private Const WM_STYLECHANGED   As Long = &H7D
Private Const WM_COMMAND        As Long = &H111
Private Const WM_KEYDOWN        As Long = &H100
Private Const WM_KEYUP          As Long = &H101
Private Const WM_CHAR           As Long = &H102
Private Const WM_UNICHAR        As Long = &H109, UNICODE_NOCHAR As Long = &HFFFF&
Private Const WM_INPUTLANGCHANGE As Long = &H51
Private Const WM_IME_SETCONTEXT As Long = &H281
Private Const WM_IME_CHAR       As Long = &H286
Private Const WM_LBUTTONDOWN    As Long = &H201
Private Const WM_LBUTTONUP      As Long = &H202
Private Const WM_MBUTTONDOWN    As Long = &H207
Private Const WM_MBUTTONUP      As Long = &H208
Private Const WM_RBUTTONDOWN    As Long = &H204
Private Const WM_RBUTTONUP      As Long = &H205
Private Const WM_LBUTTONDBLCLK  As Long = &H203
Private Const WM_MBUTTONDBLCLK  As Long = &H209
Private Const WM_RBUTTONDBLCLK  As Long = &H206
Private Const WM_MOUSEMOVE      As Long = &H200
Private Const WM_MOUSELEAVE     As Long = &H2A3
Private Const WM_HSCROLL        As Long = &H114
Private Const WM_VSCROLL        As Long = &H115
Private Const WM_CONTEXTMENU    As Long = &H7B
Private Const WM_NOTIFY         As Long = &H4E
Private Const WM_SETCURSOR      As Long = &H20, HTCLIENT As Long = 1
Private Const WM_GETTEXTLENGTH  As Long = &HE
Private Const WM_GETTEXT        As Long = &HD
Private Const WM_SETTEXT        As Long = &HC
Private Const WM_DROPFILES      As Long = &H233
Private Const WM_COPY           As Long = &H301
Private Const WM_CUT            As Long = &H300
Private Const WM_PASTE          As Long = &H302
Private Const WM_CLEAR          As Long = &H303

Private Const EM_GETSEL         As Long = &HB0&
Private Const WM_USER           As Long = &H400
Private Const EM_SCROLLCARET    As Long = WM_USER + 49
Private Const EM_SETCHARFORMAT  As Long = WM_USER + 68
Private Const EM_GETCHARFORMAT  As Long = (WM_USER + 58)
Private Const EM_DISPLAYBAND    As Long = (WM_USER + 51)
Private Const EM_EXGETSEL       As Long = (WM_USER + 52)
Private Const EM_EXSETSEL       As Long = (WM_USER + 55)
Private Const EM_HIDESELECTION  As Long = (WM_USER + 63)
Private Const EM_SETREADONLY    As Long = &HCF
Private Const EM_GETLINECOUNT   As Long = &HBA
Private Const EM_LINESCROLL     As Long = &HB6
Private Const EM_LINEINDEX      As Long = &HBB
Private Const EM_GETMARGINS     As Long = &HD4
Private Const EM_SETMARGINS     As Long = &HD3

Private Const EM_FINDTEXTEXA    As Long = (WM_USER + 79)
Private Const EM_FINDTEXTEXW    As Long = (WM_USER + 124)
Private Const EM_FINDTEXTEX     As Long = EM_FINDTEXTEXW
Private Const EM_GETTEXTEX      As Long = WM_USER + 94
Private Const EM_GETTEXTLENGTHEX As Long = WM_USER + 95

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageWLng Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SetFocus2 Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Long) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Declare Function GetCurrentObject Lib "gdi32" (ByVal Hdc As Long, ByVal uObjectType As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal Color As Long, ByVal hPal As Long, ByRef RGBResult As Long) As Long

Event ContextMenu(ByVal x As Single, ByVal y As Single, Cancel As Boolean)

Private c_SubClass         As cSubClass
Private m_hMod             As Long
Private m_hWnd             As Long
Private m_ClassName        As String



Private Sub Class_Initialize()

    m_hMod = LoadLibrary("Msftedit.dll")
    If m_hMod = 0 Then m_hMod = LoadLibrary(App.Path & "\Msftedit.dll")
    If m_hMod <> 0 Then
        m_ClassName = "RichEdit50W"
    Else
        m_hMod = LoadLibrary("Riched20.dll")
        m_ClassName = "RichEdit20W"
    End If
    
    If m_hMod = 0 Then m_hMod = LoadLibrary(App.Path & "\plugins\msftedit.dll")
    If m_hMod = 0 Then
        'MsgBox "Msftedit.dll or Riched20.dll library not found.", vbCritical
        Err.Raise 550, "Plugins", "Msftedit.dll or Riched20.dll library not found."
    End If
    
    'MSVCRT.DLL
    Set c_SubClass = New cSubClass
    
End Sub

Private Sub Class_Terminate()
    Set c_SubClass = Nothing
    If Not m_hWnd = 0 Then Call DestroyWindow(m_hWnd)
    If Not m_hMod = 0 Then Call FreeLibrary(m_hMod)
End Sub

Public Function Create(ByVal hWndParent As Long, Optional ByVal BorderStyle As eBorderStyle, Optional ByVal ScrollBars As eScrollBars, Optional mbReadOnly As Boolean, Optional mbHideSelection As Boolean)
Dim lStyle  As Long
Dim ExStyle As Long
Dim lRct(3) As Long

    lStyle = WS_CHILD Or WS_VISIBLE Or WS_CHILD Or ES_MULTILINE
    If mbReadOnly Then lStyle = lStyle Or ES_READONLY
    If Not mbHideSelection Then lStyle = lStyle Or ES_NOHIDESEL
  
    If BorderStyle = [Fixed Single] Then lStyle = lStyle Or WS_BORDER
    Select Case ScrollBars
        Case Vertical:      lStyle = lStyle Or WS_VSCROLL
        Case Horizontal:    lStyle = lStyle Or WS_HSCROLL
        Case Both:          lStyle = lStyle Or WS_HSCROLL Or WS_VSCROLL Or ES_AUTOVSCROLL Or ES_AUTOHSCROLL
    End Select

    Call GetWindowRect(hWndParent, lRct(0))
    m_hWnd = CreateWindowEx(0, m_ClassName, vbNullString, lStyle, 0, 0, lRct(1), lRct(2), hWndParent, 0, App.hInstance, ByVal 0&)
    SendMessage m_hWnd, WM_SETFONT, GetCurrentObject(GetDC(hWndParent), 6&), 1

    If m_hWnd Then
        Call c_SubClass.Subclass(hWndParent, , , Me)
        c_SubClass.AddMsg hWndParent, ALL_MESSAGES, MSG_AFTER
    End If
    
End Function

Public Function SetTextRTF(ByVal sTextRTF As String) As Boolean
Dim tSETTEXTEX As SETTEXTEX
    SetTextRTF = Not (SendMessage(m_hWnd, &H400 + 97, VarPtr(tSETTEXTEX), ByVal sTextRTF) = 0)
End Function


Public Function Move(ByVal lLeft As Long, ByVal Ltop As Long, ByVal lWidth As Long, ByVal lHeight As Long)
   Move = Not (MoveWindow(m_hWnd, lLeft, Ltop, lWidth, lHeight, True) = 0)
End Function
Public Sub Clear()
    If Not Readonly Then
        If m_hWnd <> 0 Then SendMessage m_hWnd, WM_CLEAR, 0, ByVal 0&
    Else
        Me.HideSelection = True
        Me.SetRange 0, Me.TextLen
        Me.SelText = vbNullString
        Me.HideSelection = False
    End If
End Sub
Property Let HideSelection(ByVal Value As Boolean)
    If m_hWnd <> 0 Then SendMessage m_hWnd, EM_HIDESELECTION, Abs(Value), ByVal 0&
End Property


Public Function LoadFile(ByVal sPath As String) As Boolean
Dim sText As String
Dim FF As Long
    
    On Error GoTo e
    
    FF = FreeFile
    Open sPath For Binary Access Read As FF
    sText = Space(LOF(FF))
    Get FF, , sText
    Close FF
    LoadFile = SetTextRTF(sText)
    
    Exit Function
e:
    Close FF
End Function

Public Function AddText(ByVal Text As String, Optional ByVal TextColor As Variant)
Dim RECF2 As CHARFORMAT2W

    SetRange -1, -1
     With RECF2
        .cbSize = LenB(RECF2)
        .dwMask = CFM_COLOR 'Or CFM_BACKCOLOR
        If Not IsMissing(TextColor) Then .TextColor = CLng(TextColor)
    End With
    SendMessage m_hWnd, EM_SETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)
    SendMessage m_hWnd, EM_REPLACESEL, 0, ByVal Text
End Function


Property Get hWnd() As Long: hWnd = m_hWnd: End Property
Property Get Text() As String
Dim gtlUnicode  As GETTEXTLENGTHEX
Dim gtUnicode   As GETTEXTEX
Dim lngChars    As Long

    With gtlUnicode
        .Flags = GTL_USECRLF Or GTL_PRECISE Or GTL_NUMCHARS
        .CodePage = CP_UNICODE
    End With
    lngChars = SendMessageWLng(m_hWnd, EM_GETTEXTLENGTHEX, VarPtr(gtlUnicode), 0)
    With gtUnicode
        .cb = (lngChars + 1) * 2
        .Flags = GT_USECRLF
        .CodePage = CP_UNICODE
    End With
    Text = String$(lngChars, 0)
    SendMessageWLng m_hWnd, EM_GETTEXTEX, VarPtr(gtUnicode), StrPtr(Text)

End Property
Property Let Text(ByVal Text As String)
    SetTextRTF Text
End Property

Property Get SelColor() As OLE_COLOR
Dim RECF2 As CHARFORMAT2W

    If m_hWnd = 0 Then Exit Property
    
    With RECF2
        .cbSize = LenB(RECF2)
        .dwMask = CFM_COLOR
        If (SendMessage(m_hWnd, EM_GETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)) And CFM_COLOR) <> 0 Then
            SelColor = .TextColor
        Else
            SelColor = -1
        End If
    End With

End Property
Property Let SelColor(ByVal Value As OLE_COLOR)
Dim RECF2 As CHARFORMAT2W
If m_hWnd = 0 Then Exit Property

    With RECF2
        .cbSize = LenB(RECF2)
        .dwMask = CFM_COLOR
        .TextColor = WinColor(Value)
        SendMessage m_hWnd, EM_SETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)
    End With

End Property

Property Get SelBkColor() As OLE_COLOR
Dim RECF2 As CHARFORMAT2W

    If m_hWnd = 0 Then Exit Property
    With RECF2
        .cbSize = LenB(RECF2)
        .dwMask = CFM_BACKCOLOR
        If (SendMessage(m_hWnd, EM_GETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)) And CFM_BACKCOLOR) <> 0 Then
            SelBkColor = .BackColor
        Else
            SelBkColor = -1
        End If
    End With

End Property
Property Let SelBkColor(ByVal Value As OLE_COLOR)
Dim RECF2 As CHARFORMAT2W

    If m_hWnd = 0 Then Exit Property
    With RECF2
        .cbSize = LenB(RECF2)
        .dwMask = CFM_BACKCOLOR
        .BackColor = WinColor(Value)
    End With
    SendMessage m_hWnd, EM_SETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)

End Property


Property Let SelBold(ByVal Value As Boolean)
Dim RECF2 As CHARFORMAT2W

    If m_hWnd = 0 Then Exit Property
    With RECF2
        .cbSize = LenB(RECF2)
        .dwMask = CFM_BOLD
        If Value = True Then .dwEffects = CFE_BOLD Else .dwEffects = 0
    End With
    SendMessage m_hWnd, EM_SETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)
End Property
Property Get SelBold() As Boolean
Dim RECF2 As CHARFORMAT2W

    If m_hWnd = 0 Then Exit Property
    With RECF2
        .cbSize = LenB(RECF2)
        .dwMask = CFM_BOLD
        If (SendMessage(m_hWnd, EM_GETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)) And CFM_BOLD) <> 0 Then
            SelBold = CBool((.dwEffects And CFE_BOLD) = CFE_BOLD)
        End If
    End With

End Property

Property Get SelItalic() As Boolean
Dim RECF2 As CHARFORMAT2W

    If m_hWnd = 0 Then Exit Property
    With RECF2
        .cbSize = LenB(RECF2)
        .dwMask = CFM_ITALIC
        If (SendMessage(m_hWnd, EM_GETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)) And CFM_ITALIC) <> 0 Then
            SelItalic = CBool((.dwEffects And CFE_ITALIC) = CFE_ITALIC)
        End If
    End With
End Property

Public Property Let SelItalic(ByVal Value As Boolean)
Dim RECF2 As CHARFORMAT2W

    If m_hWnd = 0 Then Exit Property
    With RECF2
        .cbSize = LenB(RECF2)
        .dwMask = CFM_ITALIC
        If Value = True Then .dwEffects = CFE_ITALIC Else .dwEffects = 0
    End With
    SendMessage m_hWnd, EM_SETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)
End Property

Public Property Get SelUnderline() As Boolean
Dim RECF2 As CHARFORMAT2W

    If m_hWnd = 0 Then Exit Property
    With RECF2
        .cbSize = LenB(RECF2)
        .dwMask = CFM_UNDERLINE
        If (SendMessage(m_hWnd, EM_GETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)) And CFM_UNDERLINE) <> 0 Then
            SelUnderline = CBool((.dwEffects And CFE_UNDERLINE) = CFE_UNDERLINE)
        End If
    End With

End Property
Public Property Let SelUnderline(ByVal Value As Boolean)
Dim RECF2 As CHARFORMAT2W

    If m_hWnd = 0 Then Exit Property
    With RECF2
        .cbSize = LenB(RECF2)
        .dwMask = CFM_UNDERLINE
        If Value = True Then .dwEffects = CFE_UNDERLINE Else .dwEffects = 0
    End With
    SendMessage m_hWnd, EM_SETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)

End Property

Property Get SelFontName() As String
Dim RECF2 As CHARFORMAT2W

    If m_hWnd = 0 Then Exit Property
    With RECF2
        .cbSize = LenB(RECF2)
        .dwMask = CFM_FACE
        If (SendMessage(m_hWnd, EM_GETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)) And CFM_FACE) <> 0 Then
            SelFontName = .FaceName()
        End If
        If InStr(SelFontName, Chr(0)) <> 0 Then SelFontName = Left$(SelFontName, InStr(SelFontName, Chr(0)) - 1)
    End With

End Property
Property Let SelFontName(ByVal Value As String)
Dim RECF2 As CHARFORMAT2W

    If m_hWnd = 0 Then Exit Property
    With RECF2
        .cbSize = LenB(RECF2)
        .dwMask = CFM_FACE

        Dim Length As Long, FontB() As Byte
        If Len(Value) > LF_FACESIZE Then
            Length = LF_FACESIZE * 2
        Else
            Length = LenB(Value)
        End If
        If Length > 0 Then
            FontB() = Value
            CopyMemory .FaceName(0), FontB(0), Length
        End If

    End With
    SendMessage m_hWnd, EM_SETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)

End Property

Public Property Get SelFontSize() As Single
Dim RECF2 As CHARFORMAT2W

    If m_hWnd = 0 Then Exit Property
    With RECF2
        .cbSize = LenB(RECF2)
        .dwMask = CFM_SIZE
        If (SendMessage(m_hWnd, EM_GETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)) And CFM_SIZE) <> 0 Then
            SelFontSize = CSng((.YHeight * 72) / 1440)
        End If
    End With

End Property
Public Property Let SelFontSize(ByVal Value As Single)
Dim RECF2 As CHARFORMAT2W

    If m_hWnd = 0 Then Exit Property
    With RECF2
        .cbSize = LenB(RECF2)
        .dwMask = CFM_SIZE
        .YHeight = (Value * 1440) / 72
    End With
    SendMessage m_hWnd, EM_SETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)

End Property

Property Get SelStart() As Long
Dim CR As CHARRANGE

    If m_hWnd = 0 Then Exit Property
    SendMessage m_hWnd, EM_EXGETSEL, 0, ByVal VarPtr(CR)
    SelStart = CR.Min
End Property
Property Let SelStart(ByVal Value As Long)
Dim CR As CHARRANGE
    If m_hWnd = 0 Then Exit Property
    If Not Value >= 0 Then Exit Property
    CR.Min = Value
    CR.Max = Value
    SendMessage m_hWnd, EM_EXSETSEL, 0, ByVal VarPtr(CR)
End Property


Property Get SelLength() As Long
Dim CR As CHARRANGE
    If m_hWnd = 0 Then Exit Property
    SendMessage m_hWnd, EM_EXGETSEL, 0, ByVal VarPtr(CR)
    SelLength = CR.Max - CR.Min
End Property
Property Let SelLength(ByVal Value As Long)
Dim CR As CHARRANGE

    If m_hWnd = 0 Then Exit Property
    If Not Value >= 0 Then Exit Property
        
    SendMessage m_hWnd, EM_EXGETSEL, 0, ByVal VarPtr(CR)
    CR.Max = CR.Min + Value
    SendMessage m_hWnd, EM_EXSETSEL, 0, ByVal VarPtr(CR)
End Property

Property Let SelText(ByVal Value As String)
    If m_hWnd <> 0 Then SendMessage m_hWnd, EM_REPLACESEL, 0, ByVal StrPtr(Value)
End Property

Public Sub SetRange(ByVal lStart As Long, ByVal lLen As Long)
Dim CR As CHARRANGE
    CR.Min = lStart
    CR.Max = lStart + lLen
    SendMessage m_hWnd, EM_EXSETSEL, 0&, CR
End Sub

Public Sub ScrollToLine(ByVal LineNumber As Long)
    If LineNumber < 0 Then Exit Sub
    If m_hWnd = 0 Then Exit Sub
    
    If SendMessage(m_hWnd, EM_LINESCROLL, 0, ByVal LineNumber - 1) <> 0 Then
        Dim FirstCharPos As Long
        FirstCharPos = SendMessage(m_hWnd, EM_LINEINDEX, LineNumber - 1, ByVal 0&)
        If FirstCharPos > -1 Then
            Me.SelStart = FirstCharPos
            SendMessage m_hWnd, EM_SCROLLCARET, 0, ByVal 0&
        End If
    End If

End Sub
Public Sub ScrollToCaret()
    If m_hWnd <> 0 Then SendMessage m_hWnd, EM_SCROLLCARET, 0, ByVal 0&
End Sub
Public Function Find(ByVal Text As String, Optional ByVal Min As Long, Optional ByVal Max As Long = -1, Optional ByVal Options As RTFFindOptions)
    Find = -1
    If m_hWnd = 0 Then Exit Function
    
    Dim REFTEX      As REFINDTEXTEX
    Dim dwOptions   As Long


    With REFTEX
    
        With .CHARRANGE
            If Min >= 0 Then .Min = Min
            If Max >= -1 Then .Max = Max
        End With
        .lpstrText = StrPtr(Text)
        dwOptions = &H1 'FR_DOWN
        
        If (Options And RTFFindWholeword) <> 0 Then dwOptions = dwOptions Or RTFFindWholeword
        If (Options And RTFFindMatchcase) <> 0 Then dwOptions = dwOptions Or RTFFindMatchcase
        
        Find = SendMessage(m_hWnd, EM_FINDTEXTEX, dwOptions, ByVal VarPtr(REFTEX))
        If (Options And RTFFindNoHighlight) = 0 And Find <> -1 Then SendMessage m_hWnd, EM_EXSETSEL, 0, ByVal VarPtr(.CharRangeText)
        
    End With

    
End Function
Public Sub SetFocus()
    If m_hWnd Then SetFocus2 m_hWnd
End Sub

Public Sub ScrollToBottom()
    SendMessage m_hWnd, WM_VSCROLL, 7&, 0    'SB_BOTTOM
End Sub
Public Sub ScrollToTop()
    SendMessage m_hWnd, WM_VSCROLL, 6&, 0    'SB_TOP
End Sub



Property Get Readonly() As Boolean
    If m_hWnd = 0 Then Exit Property
    Readonly = CBool((GetWindowLong(m_hWnd, GWL_STYLE) And ES_READONLY) <> 0)
End Property
Property Let Readonly(ByVal Value As Boolean)
    If m_hWnd = 0 Then Exit Property
    SendMessage m_hWnd, EM_SETREADONLY, IIf(Value = True, 1, 0), ByVal 0&
End Property

Property Get TextLen() As Long
    TextLen = SendMessage(m_hWnd, WM_GETTEXTLENGTH, 0&, ByVal 0&)
End Property
Property Get LineCount() As Long
    LineCount = SendMessage(m_hWnd, EM_GETLINECOUNT, 0, ByVal 0&)
End Property

Property Let LeftMargin(ByVal Value As Single)
Const EC_USEFONTINFO As Long = &HFFFF&
    If Value = EC_USEFONTINFO Or Value = -1 Then
        If m_hWnd <> 0 Then SendMessage m_hWnd, EM_SETMARGINS, &H1 Or EC_USEFONTINFO, ByVal 0&
    Else
        Value = Abs(Value)
        If m_hWnd <> 0 Then SendMessage m_hWnd, EM_SETMARGINS, &H1, ByVal MakeDWord(Value, 0)
    End If
End Property
Property Get LeftMargin() As Single
    LeftMargin = LoWord(SendMessage(m_hWnd, EM_GETMARGINS, 0, ByVal 0&))
End Property


Public Property Get RightMargin() As Single
    RightMargin = HiWord(SendMessage(m_hWnd, EM_GETMARGINS, 0, ByVal 0&))
End Property

Public Property Let RightMargin(ByVal Value As Single)
    If Value = &HFFFF& Or Value = -1 Then
        SendMessage m_hWnd, EM_SETMARGINS, &H2 Or &HFFFF&, ByVal 0&
    Else
        Value = Abs(Value)
        SendMessage m_hWnd, EM_SETMARGINS, &H2, ByVal MakeDWord(0, Value)
    End If
End Property



Private Function WinColor(ByVal Color As Long, Optional ByVal hPal As Long) As Long
    If OleTranslateColor(Color, hPal, WinColor) <> 0 Then WinColor = -1
End Function
Private Function MakeDWord(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
    MakeDWord = (CLng(HiWord) * &H10000) Or (LoWord And &HFFFF&)
End Function
Private Function LoWord(ByVal DWord As Long) As Integer
    If DWord And &H8000& Then
        LoWord = DWord Or &HFFFF0000
    Else
        LoWord = DWord And &HFFFF&
    End If
End Function
Private Function HiWord(ByVal DWord As Long) As Integer
    HiWord = (DWord And &HFFFF0000) \ &H10000
End Function
Private Function Get_X_lParam(ByVal lParam As Long) As Long
    Get_X_lParam = lParam And &H7FFF&
    If lParam And &H8000& Then Get_X_lParam = Get_X_lParam Or &HFFFF8000
End Function

Private Function Get_Y_lParam(ByVal lParam As Long) As Long
    Get_Y_lParam = (lParam And &H7FFF0000) \ &H10000
    If lParam And &H80000000 Then Get_Y_lParam = Get_Y_lParam Or &HFFFF8000
End Function


'- ordinal #1
Private Sub WndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal hWnd As Long, ByVal uMsg As ssc_eMsg, ByVal wParam As Long, ByVal lParam As Long, ByRef lParamUser As Long)

    'Debug.Print "Here " & Timer
    Select Case uMsg
        Case 20
        Case 46
        Case 47
        Case 18
        Case 85
        Case 3
        Case &HF
        Case WM_CONTEXTMENU:
            Dim PT As POINTAPI
            PT.x = Get_X_lParam(lParam)
            PT.y = Get_Y_lParam(lParam)

            ScreenToClient m_hWnd, PT
            RaiseEvent ContextMenu(PT.x, PT.y, bHandled)

        Case Else
            'Debug.Print Hex(uMsg)
    End Select

End Sub

