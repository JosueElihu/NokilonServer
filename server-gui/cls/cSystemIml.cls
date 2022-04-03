VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSystemIml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Autor          : J. Elihu
'----------------------------------------------------
'Version        : 1.0
'
'
'
'

Option Explicit
Private Const SHGFI_LARGEICON       As Long = &H0
Private Const SHGFI_SMALLICON       As Long = &H1
Private Const SHGFI_OPENICON        As Long = &H2
Private Const SHGFI_SHELLICONSIZE   As Long = &H4
Private Const SHGFI_PIDL            As Long = &H8
Private Const SHGFI_USEFILEATTRIBUTES   As Long = &H10
Private Const SHGFI_ICON            As Long = &H100
Private Const SHGFI_DISPLAYNAME     As Long = &H200
Private Const SHGFI_TYPENAME        As Long = &H400
Private Const SHGFI_ATTRIBUTES      As Long = &H800
Private Const SHGFI_ICONLOCATION    As Long = &H1000
Private Const SHGFI_EXETYPE         As Long = &H2000
Private Const SHGFI_SYSICONINDEX    As Long = &H4000
Private Const SHGFI_LINKOVERLAY     As Long = &H8000
Private Const SHGFI_SELECTED        As Long = &H10000
Private Const SHGFI_ATTR_SPECIFIED  As Long = &H20000

Private Const MAX_PATH              As Long = 260

Public Enum enmSysImlSize
    emSysIml_large = &H0        '32*32  || Si la opción Usar iconos grandes está seleccionada   -> 48x48
    emSysIml_small = &H1        '16*16  || Usuario puede personalizar el tamaño
    emSysIml_extra_large = &H2  '48*48 extra grande || Usuario puede personalizar el tamaño
    emSysIml_sys_small = &H3    'Tamaño especificado por GetSystemMetrics con SM_CXSMICON y SM_CYSMICON.
    emSysIml_jumbo = &H4        '256*256 -> Windows Vista y posterior
End Enum
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type SHFILEINFO
    hicon   As Long
    iIcon   As Long
    dwAttributes  As Long
    szDisplayName As String * MAX_PATH
    szTypeName    As String * 80
End Type

Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, Id As GUID) As Long
Private Declare Function SHGetImageList Lib "shell32.dll" (ByVal Flag As Long, Id As GUID, lResult As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As Any, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

'/ImageList
Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal Flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
Private Declare Function ImageList_GetIconSize Lib "comctl32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long
Private Declare Function ImageList_DrawEx Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long


Private m_iml   As Long
Private m_imlx  As Long
Private m_imly  As Long

Private Sub Class_Initialize()
'
End Sub
Private Sub Class_Terminate()
    m_iml = 0
End Sub

Property Get Height() As Long: Height = m_imly: End Property
Property Get Width() As Long: Width = m_imlx: End Property

Public Function CreateSystemImageList(ByVal mFlag As enmSysImlSize) As Boolean
On Error GoTo e
Dim eGUID  As GUID
Dim uSHFI  As SHFILEINFO
    
    If m_iml Then m_iml = 0
    
    CLSIDFromString StrPtr("{46EB5926-582E-4017-9FDF-E8998DAA0950}"), eGUID
    SHGetImageList mFlag, eGUID, m_iml
    If Not m_iml Then m_iml = SHGetFileInfo("C:\", 0, uSHFI, Len(uSHFI), SHGFI_SYSICONINDEX Or mFlag)
    
    If m_iml Then
        ImageList_GetIconSize m_iml, m_imlx, m_imly
        CreateSystemImageList = True
    Else
        m_imlx = 0: m_imly = 0
    End If
e:
End Function

Public Function FileImageIndex(pzp_path As String) As Long
On Error Resume Next
Dim uSHFI As SHFILEINFO
    
    If Not PathIsDirectory(pzp_path) Then
        Call SHGetFileInfo(pzp_path, ByVal 0&, uSHFI, Len(uSHFI), SHGFI_USEFILEATTRIBUTES Or SHGFI_SYSICONINDEX Or SHGFI_SMALLICON Or SHGFI_TYPENAME)
    Else
        Call SHGetFileInfo(pzp_path, ByVal 0&, uSHFI, Len(uSHFI), SHGFI_SYSICONINDEX Or SHGFI_SMALLICON Or SHGFI_TYPENAME Or SHGFI_DISPLAYNAME)
    End If
    FileImageIndex = uSHFI.iIcon
End Function

Public Function DrawFileIcon(dvc As Long, ByVal x As Long, ByVal y As Long, pzp_path As String) As Boolean
Dim i As Long
    
    i = FileImageIndex(pzp_path)
    ImageList_DrawEx m_iml, i, dvc, x, y, m_imlx, m_imly, &HFFFFFFFF, &HFF000000, 0
    
End Function
