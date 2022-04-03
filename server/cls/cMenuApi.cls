VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMenuApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-----------------------------
'Autor: Leandro Ascierto
'Date:01/03/2021
'Web: www.leandroascierto.com
'-----------------------------
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Private Declare Function GetMenuItemInfo Lib "user32.dll" Alias "GetMenuItemInfoW" (ByVal hMenu As Long, ByVal un As Long, ByVal B As Boolean, ByRef lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal hWnd As Long, ByVal lptpm As Any) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
Private Declare Function DeleteMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemW" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function RemoveMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function SetMenuInfo Lib "user32" (ByVal hMenu As Long, MI As MENUINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32.dll" Alias "SetMenuItemInfoW" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long

Private Type POINTAPI
    x                   As Long
    y                   As Long
End Type

Private Type MENUITEMINFO
    cbSize              As Long
    fMask               As Long
    fType               As Long
    fState              As Long
    wID                 As Long
    hSubMenu            As Long
    hbmpChecked         As Long
    hbmpUnchecked       As Long
    dwItemData          As Long
    dwTypeData          As Long
    cch                 As Long
    hbmpItem            As Long
End Type

Private Type MENUINFO
    cbSize          As Long
    fMask           As Long
    dwStyle         As Long
    cyMax           As Long
    RhbrBack        As Long
    dwContextHelpID As Long
    dwMenuData      As Long
End Type

Private Const MF_BYCOMMAND      As Long = &H0&
Private Const MF_STRING         As Long = &H0&
Private Const MF_DISABLED       As Long = &H2&
Private Const MF_BITMAP         As Long = &H4&
Private Const MF_CHECKED        As Long = &H8&
Private Const MF_MENUBREAK      As Long = &H40&
Private Const MF_HILITE         As Long = &H80&
Private Const MF_BYPOSITION     As Long = &H400&
Private Const MF_SEPARATOR      As Long = &H800&
Private Const MF_DEFAULT        As Long = &H1000&

Private Const MIIM_STATE        As Long = &H1
Private Const MIIM_ID           As Long = &H2
Private Const MIIM_SUBMENU      As Long = &H4
Private Const MIIM_CHECKMARKS   As Long = &H8
Private Const MIIM_TYPE         As Long = &H10
Private Const MIIM_DATA         As Long = &H20
Private Const MIIM_STRING       As Long = &H40
Private Const MIIM_BITMAP       As Long = &H80

Private Const MFS_DEFAULT       As Long = &H1000&
Private Const MFT_RADIOCHECK    As Long = &H200&
Private Const MNS_NOCHECK       As Long = &H80000000
Private Const MIM_MAXHEIGHT     As Long = &H1
Private Const MIM_BACKGROUND    As Long = &H2
Private Const MIM_STYLE         As Long = &H10


Public Enum ShowMenuFlags
    TPM_BOTTOMALIGN = &H20&
    TPM_CENTERALIGN = &H4&
    TPM_HORIZONTAL = &H0&
    TPM_HORNEGANIMATION = &H800&
    TPM_HORPOSANIMATION = &H400&
    TPM_LEFTALIGN = &H0&
    TPM_LEFTBUTTON = &H0&
    TPM_NOANIMATION = &H4000&
    TPM_NONOTIFY = &H80&
    TPM_RECURSE = &H1&
    TPM_RETURNCMD = &H100&
    TPM_RIGHTALIGN = &H8&
    TPM_RIGHTBUTTON = &H2&
    TPM_TOPALIGN = &H0&
    TPM_VCENTERALIGN = &H10&
    TPM_VERNEGANIMATION = &H2000&
    TPM_VERPOSANIMATION = &H1000&
End Enum

Private cSubMenu As Collection
Private m_SubMenu As cMenuApi
Public ParentId As Long
Public hMenu As Long

Private Sub Class_Initialize()
    hMenu = CreatePopupMenu()
    Set cSubMenu = New Collection
End Sub

Public Function AddItem(Optional ByVal ID As Long, Optional ByVal Caption As String, Optional ByVal Separator As Boolean, Optional ByVal Checked As Boolean, Optional hBitmap As Long, Optional Disabled As Boolean, Optional BarBreak As Boolean, Optional Index As Long = -1) As Boolean
    Dim MII As MENUITEMINFO
    
    With MII
         .cbSize = Len(MII)
         .wID = ID
        If Separator Then
            .fMask = MIIM_TYPE
            .fType = MF_SEPARATOR
        Else
            If BarBreak Then
                .fMask = MIIM_TYPE
                .fType = MF_MENUBREAK
            Else
                .fMask = MIIM_STRING
            End If
            .fMask = .fMask Or MIIM_STATE Or MIIM_ID Or IIf(hBitmap <> 0, MIIM_STRING Or MIIM_BITMAP, 0)
            .cch = Len(Caption)
            .dwTypeData = StrPtr(Caption)
            .fState = -((Checked * MF_CHECKED) Or (Disabled * MF_DISABLED))
            .hbmpItem = hBitmap
        End If
    End With
    
    AddItem = InsertMenuItem(hMenu, Index, True, MII)
End Function

Public Function RemoveItem(ByVal ItemPos As Long) As Boolean
    RemoveItem = RemoveMenu(hMenu, ItemPos, MF_BYPOSITION)
End Function

Public Function AddSubMenu(Optional Key As Variant, Optional ByVal Caption As String, Optional hBitmap As Long, Optional Disabled As Boolean, Optional Index As Long = -1) As cMenuApi
    Dim cNewSubMenu As cMenuApi
    Dim Flags As String
    Dim MII As MENUITEMINFO
    
    Set cNewSubMenu = New cMenuApi

    With MII
        .wID = CLng(Rnd() * 1000)
        .cbSize = Len(MII)
        .fMask = MIIM_STRING Or MIIM_SUBMENU Or MIIM_STATE Or MIIM_ID Or IIf(hBitmap <> 0, MIIM_BITMAP, 0)
        .dwTypeData = StrPtr(Caption)
        .cch = Len(Caption)
        .fState = -(Disabled * MF_DISABLED)
        .hbmpItem = hBitmap
        .hSubMenu = cNewSubMenu.hMenu
    End With

    If InsertMenuItem(hMenu, Index, True, MII) Then
        cNewSubMenu.ParentId = MII.wID
        Set AddSubMenu = cNewSubMenu
        cSubMenu.Add cNewSubMenu, Key
    End If
End Function

Public Function RemoveSubMenu(Optional Key As Variant) As Boolean
    If RemoveMenu(hMenu, cSubMenu.Item(Key).ParentId, MF_BYCOMMAND) Then
        Call cSubMenu.Remove(Key)
    End If
End Function

Public Property Get SubMenu(Key As Variant) As cMenuApi
    Set SubMenu = cSubMenu.Item(Key)
End Property

Public Property Let ItemBitmap(ByVal ItemPos As Long, ByVal hBitmap As Long)
    Dim MII As MENUITEMINFO
    With MII
        .cbSize = Len(MII)
        .fMask = MIIM_BITMAP
        .hbmpItem = hBitmap
    End With
    
    Call SetMenuItemInfo(hMenu, ItemPos, True, MII)
End Property

Public Property Get ItemBitmap(ByVal ItemPos As Long) As Long
    Dim MII As MENUITEMINFO
    With MII
        .cbSize = Len(MII)
        .fMask = MIIM_BITMAP
    End With

    If GetMenuItemInfo(hMenu, ItemPos, True, MII) Then
        ItemBitmap = MII.hbmpItem
    End If
End Property

Public Property Let ItemCaption(ByVal ItemPos As Long, ByVal NewValue As String)
    Dim MII As MENUITEMINFO
    With MII
        .cbSize = Len(MII)
        .fMask = MIIM_STRING
        .cch = LenB(NewValue)
        .dwTypeData = StrPtr(NewValue)
    End With
    
    Call SetMenuItemInfo(hMenu, ItemPos, True, MII)
End Property

Public Property Get ItemCaption(ByVal ItemPos As Long) As String
    Dim MII As MENUITEMINFO
    Dim sBuff As String
    
    sBuff = Space$(256)
    With MII
          .cbSize = Len(MII)
          .dwTypeData = StrPtr(sBuff)
          .fMask = MIIM_STRING
          .cch = LenB(sBuff)
    End With

    If GetMenuItemInfo(hMenu, ItemPos, True, MII) Then
        ItemCaption = Left(sBuff, MII.cch - 1)
    End If
End Property

Public Property Let ItemRadioCheck(ByVal ItemPos As Long, ByVal NewValue As Boolean)
    Dim MII As MENUITEMINFO
    Dim sBuff As String
    sBuff = Space$(256)
    Dim s As String
    With MII
        .cbSize = Len(MII)
        .fMask = MIIM_TYPE
        .cch = Len(sBuff)
        .dwTypeData = StrPtr(sBuff)
    End With
    
    If GetMenuItemInfo(hMenu, ItemPos, True, MII) Then
        MII.fType = IIf(NewValue, MII.fType Or MFT_RADIOCHECK, MII.fType And Not MFT_RADIOCHECK)
        Call SetMenuItemInfo(hMenu, ItemPos, True, MII)
    End If
End Property

Public Property Get ItemRadioCheck(ByVal ItemPos As Long) As Boolean
    Dim MII As MENUITEMINFO
    With MII
        .cbSize = Len(MII)
        .fMask = MIIM_TYPE
    End With

    If GetMenuItemInfo(hMenu, ItemPos, True, MII) Then
        ItemRadioCheck = (MII.fType And MFT_RADIOCHECK = MFT_RADIOCHECK)
    End If
End Property

Public Property Let ItemCheck(ByVal ItemPos As Long, ByVal NewValue As Boolean)
    SetMenuItemState ItemPos, MF_CHECKED, NewValue
End Property

Public Property Get ItemCheck(ByVal ItemPos As Long) As Boolean
    ItemCheck = GetMenuItemState(ItemPos, MF_CHECKED)
End Property

Public Property Let ItemDisabled(ByVal ItemPos As Long, ByVal NewValue As Boolean)
    SetMenuItemState ItemPos, MF_DISABLED, NewValue
End Property

Public Property Get ItemDisabled(ByVal ItemPos As Long) As Boolean
    ItemDisabled = GetMenuItemState(ItemPos, MF_DISABLED)
End Property

Public Property Let ItemDefault(ByVal ItemPos As Long, ByVal NewValue As Boolean)
    SetMenuItemState ItemPos, MF_DEFAULT, NewValue
End Property

Public Property Get ItemDefault(ByVal ItemPos As Long) As Boolean
    ItemDefault = GetMenuItemState(ItemPos, MF_DEFAULT)
End Property

Public Property Let ItemHilite(ByVal ItemPos As Long, ByVal NewValue As Boolean)
    SetMenuItemState ItemPos, MF_HILITE, NewValue
End Property

Public Property Get ItemHilite(ByVal ItemPos As Long) As Boolean
    ItemHilite = GetMenuItemState(ItemPos, MF_HILITE)
End Property

Private Function SetMenuItemState(ItemPos As Long, Flags As Long, NewValue As Boolean)
    Dim MII As MENUITEMINFO
    With MII
        .cbSize = Len(MII)
        .fMask = MIIM_STATE
    End With
    Call GetMenuItemInfo(hMenu, ItemPos, True, MII)
    MII.fState = IIf(NewValue, MII.fState Or Flags, MII.fState And Not Flags)
    Call SetMenuItemInfo(hMenu, ItemPos, True, MII)
End Function

Public Property Get ItemsCount()
    ItemsCount = GetMenuItemCount(hMenu)
End Property

Public Property Let RemoveMenuCheck(NewValue As Boolean)
    Dim MI As MENUINFO
    
    With MI
        .cbSize = Len(MI)
        .fMask = MIM_STYLE
        .dwStyle = IIf(NewValue, MNS_NOCHECK, 0)
    End With

    SetMenuInfo hMenu, MI
End Property

Public Property Let MaxHeightMenu(NewValue As Long)
    Dim MI As MENUINFO
    
    With MI
        .cbSize = Len(MI)
        .fMask = MIM_MAXHEIGHT
        .cyMax = NewValue
    End With

    SetMenuInfo hMenu, MI
End Property

Public Property Let BackColorMenu(ByVal NewValue As OLE_COLOR)
    Dim MI As MENUINFO
    
    With MI
        .cbSize = Len(MI)
        .fMask = MIM_BACKGROUND
        .RhbrBack = CreateSolidBrush(NewValue)
    End With

    SetMenuInfo hMenu, MI
End Property


Public Function PopupMenu(Optional ByVal x As Long = -1, Optional ByVal y As Long = -1, Optional hWnd As Long, Optional Flags As ShowMenuFlags) As Long
    Dim PT As POINTAPI
    
    If x = -1 Or y = -1 Then
       GetCursorPos PT
       x = PT.x
       y = PT.y
    End If
    If hWnd = 0 Then hWnd = GetForegroundWindow
    PopupMenu = TrackPopupMenuEx(hMenu, Flags Or TPM_RETURNCMD, x, y, hWnd, ByVal 0&)

End Function

Private Function GetMenuItemState(ItemPos As Long, Flags As Long) As Boolean
    Dim MII As MENUITEMINFO
    With MII
          .cbSize = Len(MII)
          .fMask = MIIM_STATE
    End With

    If GetMenuItemInfo(hMenu, ItemPos, True, MII) Then
        GetMenuItemState = ((MII.fState And Flags) = Flags)
    End If
End Function

Private Sub Class_Terminate()
    DestroyMenu hMenu
    Set cSubMenu = Nothing
End Sub
