VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScrollBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SB_BOTTOM       As Integer = 7
Private Const SB_ENDSCROLL    As Integer = 8
Private Const SB_LEFT         As Integer = 6
Private Const SB_LINEDOWN     As Integer = 1
Private Const SB_LINELEFT     As Integer = 0
Private Const SB_LINERIGHT    As Integer = 1
Private Const SB_LINEUP       As Integer = 0
Private Const SB_PAGEDOWN     As Integer = 3
Private Const SB_PAGELEFT     As Integer = 2
Private Const SB_PAGERIGHT    As Integer = 3
Private Const SB_PAGEUP       As Integer = 2
Private Const SB_RIGHT        As Integer = 7
Private Const SB_THUMBTRACK   As Integer = 5
Private Const SB_TOP          As Integer = 6

Private Const SIF_ALL                           As Double = (&H1 Or &H2 Or &H4 Or &H10)

Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize     As Long
    dwMajorVersion          As Long
    dwMinorVersion          As Long
    dwBuildNumber           As Long
    dwPlatformId            As Long
    szCSDVersion            As String * 128
End Type

' Scroll bar:
Private Type SCROLLINFO
    cbSize  As Long    ' Size of structure
    fMask   As Long     ' Which value(s) you are changing
    nMin    As Long      ' Minimum value of the scroll bar
    nMax    As Long      ' Maximum value of the scroll bar
    nPage   As Long     ' Large-change amount
    nPos    As Long      ' Current value
    nTrackPos As Long ' Current scroll position
End Type

'/* Orientation
Public Enum EFSOrientationConstants
    efsoHorizontal
    efsoVertical
    efsoBoth
End Enum

'/* Style
Public Enum EFSStyleConstants
    efsRegular = 0
    efsEncarta = 1
    efsFlat = 2
End Enum

'/* Bars:
Public Enum EFSScrollBarConstants
    efsHorizontal = 0
    efsVertical = 1
End Enum

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Long, ByVal fuWinIni As Long) As Long
Private Declare Function InitialiseFlatSB Lib "comctl32.dll" Alias "InitializeFlatSB" (ByVal lhWnd As Long) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function EnableScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Private Declare Function FlatSB_EnableScrollBar Lib "comctl32.dll" (ByVal hwnd As Long, ByVal int2 As Long, ByVal UINT3 As Long) As Long
Private Declare Function FlatSB_ShowScrollBar Lib "comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_GetScrollInfo Lib "comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function FlatSB_SetScrollInfo Lib "comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollProp Lib "comctl32.dll" (ByVal hwnd As Long, ByVal Index As Long, ByVal NewValue As Long, ByVal fRedraw As Boolean) As Long

Private Declare Function UninitializeFlatSB Lib "comctl32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptX As Long, ByVal ptY As Long) As Long

Private Declare Function SetFocusEx Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long

Public Event ScrollClick(eBar As EFSScrollBarConstants, eButton As MouseButtonConstants)
Public Event Scroll(eBar As EFSScrollBarConstants)
Public Event Change(eBar As EFSScrollBarConstants)
Public Event MouseWheel(eBar As EFSScrollBarConstants, lAmount As Long)


Private cSubClass                               As cSubClass

Private m_lWheelScrollLines                     As Long
Private m_bNoFlatScrollBars                     As Boolean
Private m_hwnd                                  As Long
Private m_lSmallChangeHorz                      As Long
Private m_lSmallChangeVert                      As Long
Private m_bEnabledHorz                          As Boolean
Private m_bEnabledVert                          As Boolean
Private m_bVisibleHorz                          As Boolean
Private m_bVisibleVert                          As Boolean
Private m_bInitialised                          As Boolean
Private m_eStyle                                As EFSStyleConstants
Private m_eOrientation                          As EFSOrientationConstants

Public Property Get Visible(ByVal eBar As EFSScrollBarConstants) As Boolean
    If (eBar = efsHorizontal) Then
        Visible = m_bVisibleHorz
    Else
        Visible = m_bVisibleVert
    End If
End Property

Public Property Let Visible(ByVal eBar As EFSScrollBarConstants, ByVal bState As Boolean)

    If (eBar = efsHorizontal) Then
        m_bVisibleHorz = bState
    Else
        m_bVisibleVert = bState
    End If

    If (m_bNoFlatScrollBars) Then
        ShowScrollBar m_hwnd, eBar, Abs(bState)
    Else
        FlatSB_ShowScrollBar m_hwnd, eBar, Abs(bState)
    End If

End Property

Public Property Get Orientation() As EFSOrientationConstants
    Orientation = m_eOrientation
End Property

Public Property Let Orientation(ByVal eOrientation As EFSOrientationConstants)
    m_eOrientation = eOrientation
    pSetOrientation
End Property

Private Sub pSetOrientation()
    ShowScrollBar m_hwnd, 0, Abs((m_eOrientation = efsoBoth) Or (m_eOrientation = efsoHorizontal))
    ShowScrollBar m_hwnd, 1, Abs((m_eOrientation = efsoBoth) Or (m_eOrientation = efsoVertical))
End Sub

Private Sub pGetSI(ByVal eBar As EFSScrollBarConstants, ByRef tSI As SCROLLINFO, ByVal fMask As Long)
Dim lO      As Long

    lO = eBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)

    If (m_bNoFlatScrollBars) Then
        GetScrollInfo m_hwnd, lO, tSI
    Else
        FlatSB_GetScrollInfo m_hwnd, lO, tSI
    End If

End Sub

Private Sub pLetSI(ByVal eBar As EFSScrollBarConstants, ByRef tSI As SCROLLINFO, ByVal fMask As Long)
Dim lO      As Long

    lO = eBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)

    If (m_bNoFlatScrollBars) Then
        SetScrollInfo m_hwnd, lO, tSI, True
    Else
        FlatSB_SetScrollInfo m_hwnd, lO, tSI, True
    End If

End Sub

Public Property Get Style() As EFSStyleConstants
    Style = m_eStyle
End Property

Public Property Let Style(ByVal eStyle As EFSStyleConstants)
Dim lR      As Long

    If Not eStyle = efsRegular Then
        If (m_bNoFlatScrollBars) Then
            '/* can't do it..
            debugmsg "Can't set non-regular style mode on this system - XP or COMCTL32.DLL version < 4.71."
            Exit Property
        End If
    End If

    If (m_eOrientation = efsoHorizontal) Or (m_eOrientation = efsoBoth) Then
        lR = FlatSB_SetScrollProp(m_hwnd, &H200, eStyle, True)
    End If
    If (m_eOrientation = efsoVertical) Or (m_eOrientation = efsoBoth) Then
        lR = FlatSB_SetScrollProp(m_hwnd, &H100, eStyle, True)
    End If
    m_eStyle = eStyle

End Property

Public Property Get SmallChange(ByVal eBar As EFSScrollBarConstants) As Long
    If (eBar = efsHorizontal) Then
        SmallChange = m_lSmallChangeHorz
    Else
        SmallChange = m_lSmallChangeVert
    End If
End Property

Public Property Get Enabled(ByVal eBar As EFSScrollBarConstants) As Boolean
    If (eBar = efsHorizontal) Then
        Enabled = m_bEnabledHorz
    Else
        Enabled = m_bEnabledVert
    End If
End Property

Public Property Let Enabled(ByVal eBar As EFSScrollBarConstants, ByVal bEnabled As Boolean)
Dim lO      As Long
Dim lF      As Long

    lO = eBar
    If (bEnabled) Then
        lF = &H0
    Else
        lF = &H3
    End If
    If (m_bNoFlatScrollBars) Then
        EnableScrollBar m_hwnd, lO, lF
    Else
        FlatSB_EnableScrollBar m_hwnd, lO, lF
    End If

End Property

Public Property Get Min(ByVal eBar As EFSScrollBarConstants) As Long
Dim tSI         As SCROLLINFO
    pGetSI eBar, tSI, &H1
    Min = tSI.nMin
End Property

Public Property Get Max(ByVal eBar As EFSScrollBarConstants) As Long
Dim tSI         As SCROLLINFO

    pGetSI eBar, tSI, &H1 Or &H2
    Max = tSI.nMax - tSI.nPage

End Property

Public Property Get Value(ByVal eBar As EFSScrollBarConstants) As Long
Dim tSI         As SCROLLINFO

    pGetSI eBar, tSI, &H4
    Value = tSI.nPos

End Property

Public Property Get LargeChange(ByVal eBar As EFSScrollBarConstants) As Long
Dim tSI         As SCROLLINFO

    pGetSI eBar, tSI, &H2
    LargeChange = tSI.nPage
End Property

Public Property Let Min(ByVal eBar As EFSScrollBarConstants, ByVal iMin As Long)
Dim tSI         As SCROLLINFO

    tSI.nMin = iMin
    tSI.nMax = Max(eBar) + LargeChange(eBar)
    pLetSI eBar, tSI, &H1

End Property

Public Property Let Max(ByVal eBar As EFSScrollBarConstants, ByVal iMax As Long)
Dim tSI         As SCROLLINFO

    tSI.nMax = iMax + LargeChange(eBar)
    tSI.nMin = Min(eBar)
    pLetSI eBar, tSI, &H1
    pRaiseEvent eBar, False

End Property

Public Property Let Value(ByVal eBar As EFSScrollBarConstants, ByVal iValue As Long)
Dim tSI         As SCROLLINFO

    If Not iValue = Value(eBar) Then
        tSI.nPos = iValue
        pLetSI eBar, tSI, &H4
        pRaiseEvent eBar, False
    End If
End Property

Public Property Let LargeChange(ByVal eBar As EFSScrollBarConstants, ByVal iLargeChange As Long)
Dim tSI                     As SCROLLINFO
Dim lCurMax                 As Long
Dim lCurLargeChange         As Long

    pGetSI eBar, tSI, SIF_ALL
    tSI.nMax = tSI.nMax - tSI.nPage + iLargeChange
    tSI.nPage = iLargeChange
    pLetSI eBar, tSI, &H2 Or &H1

End Property

Public Property Let SmallChange(ByVal eBar As EFSScrollBarConstants, ByVal lSmallChange As Long)
    If (eBar = efsHorizontal) Then
        m_lSmallChangeHorz = lSmallChange
    Else
        m_lSmallChangeVert = lSmallChange
    End If
End Property

Public Property Get CanBeFlat() As Boolean
    CanBeFlat = Not (m_bNoFlatScrollBars)
End Property

Private Sub pCreateScrollBar()

Dim lR              As Long
Dim lStyle          As Long
Dim hParent         As Long
Dim lMajor          As Long
Dim lMinor          As Long
Dim isXp            As Boolean

    GetWindowsVersion lMajor, lMinor
    If (lMajor > 5) Then
        isXp = True
    ElseIf (lMajor = 5) And (lMinor >= 1) Then
        isXp = True
    End If

    If isXp Then
        m_bNoFlatScrollBars = True
    Else
        '/* Just checks for flag scroll bars...
        On Error Resume Next
            lR = InitialiseFlatSB(m_hwnd)
            If Not Err.Number = 0 Then
                '/* Can't find DLL entry point InitializeFlatSB in COMCTL32.DLL
                '/* Means we have version prior to 4.71
                '/* We get standard scroll bars.
                m_bNoFlatScrollBars = True
            Else
                Style = m_eStyle
            End If
        End If

        On Error GoTo 0

End Sub

Private Sub GetWindowsVersion(Optional ByRef lMajor = 0, Optional ByRef lMinor = 0, Optional ByRef lRevision = 0, Optional ByRef lBuildNumber = 0)
Dim lR      As Long

    lR = GetVersion()
    lBuildNumber = (lR And &H7F000000) \ &H1000000
    If (lR And &H80000000) Then lBuildNumber = lBuildNumber Or &H80
    lRevision = (lR And &HFF0000) \ &H10000
    lMinor = (lR And &HFF00&) \ &H100
    lMajor = (lR And &HFF)

End Sub

Public Sub Create(ByVal hWndA As Long)
    pClearUp
    m_hwnd = hWndA
    pCreateScrollBar
    pAttachMessages
End Sub

Private Sub pClearUp()
On Error Resume Next

        If Not m_hwnd = 0 Then
            '/* Stop flat scroll bar if we have it:
            If Not (m_bNoFlatScrollBars) Then
                UninitializeFlatSB m_hwnd
            End If

            With cSubClass
                .DelMsg m_hwnd, &H114, MSG_BEFORE
                .DelMsg m_hwnd, &H115, MSG_BEFORE
                .DelMsg m_hwnd, &H20A, MSG_BEFORE
                .DelMsg m_hwnd, &HA1, MSG_BEFORE
                .DelMsg m_hwnd, &HA7, MSG_BEFORE
                .DelMsg m_hwnd, &HA4, MSG_BEFORE
                
                .UnSubclass m_hwnd
            End With
        End If

        m_hwnd = 0
        m_bInitialised = False

    On Error GoTo 0

End Sub

Private Sub pAttachMessages()

    If Not m_hwnd = 0 Then
        With cSubClass
            If .Subclass(m_hwnd, , , Me) Then
                .AddMsg m_hwnd, &H114, MSG_BEFORE
                .AddMsg m_hwnd, &H115, MSG_BEFORE
                .AddMsg m_hwnd, &H20A, MSG_BEFORE
                .AddMsg m_hwnd, &HA1, MSG_BEFORE
                .AddMsg m_hwnd, &HA7, MSG_BEFORE
                .AddMsg m_hwnd, &HA4, MSG_BEFORE
            End If
            
        
        End With
        SystemParametersInfo &H68, 0, m_lWheelScrollLines, 0  'SPI_GETWHEELSCROLLLINES = 104

        If (m_lWheelScrollLines <= 0) Then
            m_lWheelScrollLines = 3
        End If
        m_bInitialised = True
    End If

End Sub

Private Sub debugmsg(ByVal sMsg As String)
#If DEBUGMODE = 1 Then
    MsgBox sMsg
#Else
    Debug.Print sMsg
#End If
End Sub

Private Sub Class_Initialize()

    Set cSubClass = New cSubClass
    
    m_lSmallChangeHorz = 1
    m_lSmallChangeVert = 1
    m_eStyle = efsRegular
    m_eOrientation = efsoBoth
End Sub

Private Sub Class_Terminate()
    pClearUp
    Set cSubClass = Nothing
End Sub

Private Function pRaiseEvent(ByVal eBar As EFSScrollBarConstants, ByVal bScroll As Boolean)
Static s_lLastValue(0 To 1)         As Long

    If Not Value(eBar) = s_lLastValue(eBar) Then
        If (bScroll) Then
            RaiseEvent Scroll(eBar)
        Else
            RaiseEvent Change(eBar)
        End If
        s_lLastValue(eBar) = Value(eBar)
    End If

End Function

Private Function LoWord(DWord As Long) As Long

' =====================================================================
' function to return the LoWord of a Long value
' =====================================================================

    If DWord And &H8000& Then
        LoWord = DWord Or &HFFFF0000
    Else
        LoWord = DWord And &HFFFF&
    End If

End Function

Private Function HiWord(DWord As Long) As Long

' =====================================================================
' function to return the HiWord of a Long value
' =====================================================================

    HiWord = (DWord And &HFFFF0000) \ &H10000

End Function

'- callback, usually ordinal #1, the last method in this source file----------------------

':) Ulli's VB Code Formatter V2.24.17 (2014-Apr-21 15:44)  Decl: 202  Code: 552  Total: 754 Lines
':) CommentOnly: 99 (13.1%)  Commented: 11 (1.5%)  Filled: 580 (76.9%)  Empty: 174 (23.1%)  Max Logic Depth: 5

'- ordinal #1
Private Sub WndProc(ByVal bBefore As Boolean, _
       ByRef bHandled As Boolean, _
       ByRef lReturn As Long, _
       ByVal hwnd As Long, _
       ByVal uMsg As Long, _
       ByVal wParam As Long, _
       ByVal lParam As Long, _
       ByRef lParamUser As Long)


Dim lScrollCode             As Long
Dim tSI                     As SCROLLINFO
Dim lv As Long, lSC         As Long
Dim eBar                    As EFSScrollBarConstants
Dim zDelta                  As Long
Dim lDelta                  As Long
Dim wMKeyFlags              As Long
Dim wRect As RECT

    Select Case uMsg
    
    Case &H20A  'WM_MOUSEWHEEL = &H20A

        GetWindowRect hwnd, wRect
        wRect.Right = wRect.Right + 1
        wRect.Bottom = wRect.Bottom + 1

        If PtInRect(wRect, LoWord(lParam), HiWord(lParam)) Then
            '/* Low-word of wParam indicates whether virtual keys
            '/* are down
            wMKeyFlags = wParam And &HFFFF&
            '/* High order word is the distance the wheel has been rotated,
            '/* in multiples of 120:
            If (wParam And &H8000000) Then
                '/* Towards the user:
                zDelta = &H8000& - (wParam And &H7FFF0000) \ &H10000
            Else
                '/* Away from the user:
                zDelta = -((wParam And &H7FFF0000) \ &H10000)
            End If
            'm_lWheelScrollLines=3
            'm_lSmallChangeVert=1
            lDelta = (zDelta \ 120) * SmallChange(efsVertical) * m_lWheelScrollLines
            eBar = efsVertical
            RaiseEvent MouseWheel(eBar, lDelta)
            If Not (lDelta = 0) Then
                If (Visible(eBar)) Then
                    Value(eBar) = Value(eBar) + lDelta
                    lReturn = 1
                End If
            End If
        End If

    Case &H115, &H114 'WM_VSCROLL = &H115
        'WM_HSCROLL = &H114
        If (uMsg = &H114) Then
            eBar = efsHorizontal
        Else
            eBar = efsVertical
        End If
        lScrollCode = (wParam And &HFFFF&) 'LoWord(wParam)
        Select Case lScrollCode
        Case SB_THUMBTRACK
            '/* Is vertical/horizontal?
            pGetSI eBar, tSI, &H10
            Value(eBar) = tSI.nTrackPos 'HiWord(wParam)
            pRaiseEvent eBar, True

        Case SB_LEFT, SB_TOP
            Value(eBar) = Min(eBar)
            pRaiseEvent eBar, False

        Case SB_RIGHT, SB_BOTTOM
            Value(eBar) = Max(eBar)
            pRaiseEvent eBar, False

        Case SB_LINELEFT, SB_LINEUP
            lv = Value(eBar)
            If (eBar = efsHorizontal) Then
                lSC = m_lSmallChangeHorz
            Else
                lSC = m_lSmallChangeVert
            End If
            If (lv - lSC < Min(eBar)) Then
                Value(eBar) = Min(eBar)
            Else
                Value(eBar) = lv - lSC
            End If
            pRaiseEvent eBar, False

        Case SB_LINERIGHT, SB_LINEDOWN
            lv = Value(eBar)
            If (eBar = efsHorizontal) Then
                lSC = m_lSmallChangeHorz
            Else
                lSC = m_lSmallChangeVert
            End If
            If (lv + lSC > Max(eBar)) Then
                Value(eBar) = Max(eBar)
            Else
                Value(eBar) = lv + lSC
            End If
            pRaiseEvent eBar, False

        Case SB_PAGELEFT, SB_PAGEUP
            Value(eBar) = Value(eBar) - LargeChange(eBar)
            pRaiseEvent eBar, False

        Case SB_PAGERIGHT, SB_PAGEDOWN
            Value(eBar) = Value(eBar) + LargeChange(eBar)
            pRaiseEvent eBar, False

        Case SB_ENDSCROLL
            pRaiseEvent eBar, False
        End Select

    Case &HA1, &HA4  'WM_NCLBUTTONDOWN = &HA1
        'WM_NCRBUTTONDOWN = &HA4
        Dim eBtn As MouseButtonConstants
        eBtn = IIf(uMsg = &HA1, vbLeftButton, vbRightButton)
        If wParam = 7 Then
            RaiseEvent ScrollClick(efsHorizontal, eBtn)
        ElseIf wParam = 6 Then
            RaiseEvent ScrollClick(efsVertical, eBtn)
        End If
    End Select
    
End Sub
