Attribute VB_Name = "mPaths"
Option Explicit

Private Const OFS_MAXPATHNAME = 128

Public Enum SystemFolderID
  uMyPictures = &H27
  uSystem = &H25
  uAdminTools = &H30
  uAltStartUp = &H1D
  uAppData = &H1A
  uPrograms = &H2
  uRecent = &H8
  uSendTo = &H9
  uStartMenu = &HB
  uStartUp = &H7
  uSystemX86 = &H29
  uTemplates = &H15
  uWindows = &H24
  uBitBucket = &HA
  uCommonAdminTools = &H2F
  uCommonAltStartUp = &H1E
  uCommonAppData = &H23
  uCommonDesktopDirectory = &H19
  uCommonDocuments = &H2E
  uCommonFavorites = &H1F
  uCommonPrograms = &H17
  uCommonStartMenu = &H16
  uCommonStartUp = &H18
  uCommonTemplates = &H2D
  uConnections = &H31
  uControls = &H3
  uCookies = &H21
  uDesktop = &H0
  uMyMusic = &HD
  uMyVideo = &HE
  uDesktopDirectory = &H10
  uDrives = &H11
  uFavorites = &H6
  uFonts = &H14
  uInternet = &H1
  uHistory = &H22
  uInternetCache = &H20
  uLocalAppData = &H1C
  uNetHood = &H13
  uNetwork = &H12
  uPersonal = &H5
  uPrinters = &H4
  uPrintHood = &H1B
  uProfile = &H28
  uProgramFiles = &H26
  uProgramFilesX86 = &H2A
  uProgramFilesCommon = &H2B
  uProgramFilesCommonX86 = &H2C
End Enum

Private Type SHFILEINFO
  hIcon          As Long
  iIcon          As Long
  dwAttributes   As Long
  szDisplayName  As String * 260
  szTypeName     As String * 80
End Type

Private Type SHFILEOP
  hWnd        As Long
  wFunc       As Long
  pFrom       As String
  pTo         As String
  lflags      As Long
  lAnyOperationsAborted   As Boolean
  hNameMappings           As Long
  lpszProgressTitle       As String
End Type

Private Enum eFO
  FO_COPY = &H2&
  FO_DELETE = &H3&
  FO_MOVE = &H1&
  FO_RENAME = &H4&
  
  FOF_MULTIDESTFILES = &H1&
  FOF_CONFIRMMOUSE = &H2&
  FOF_SILENT = &H4&
  FOF_RENAMEONCOLLISION = &H8&
  FOF_NOCONFIRMATION = &H10&
  FOF_WANTMAPPINGHANDLE = &H20&
  
  FOF_ALLOWUNDO = &H40&
  FOF_FILESONLY = &H80&
  FOF_SIMPLEPROGRESS = &H100&
  FOF_NOCONFIRMMKDIR = &H200&
  FOF_NOERRORUI = &H400&
  FOF_NOCOPYSECURITYATTRIBS = &H800&
End Enum

Private Type OFSTRUCT
  cBytes      As Byte
  fFixedDisk  As Byte
  nErrCode    As Integer
  Reserved1   As Integer
  Reserved2   As Integer
  szPathName(OFS_MAXPATHNAME) As Byte
End Type
Private Type FILETIME
  dwLowDateTime   As Long
  dwHighDateTime  As Long
End Type
Private Type SYSTEMTIME
  wYear       As Integer
  wMonth      As Integer
  wDayOfWeek  As Integer
  wDay        As Integer
  wHour       As Integer
  wMinute     As Integer
  wSecond     As Integer
  wMilliseconds As Integer
End Type


Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, pIdl As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pIdl As Long, ByVal pszPath As String) As Long
Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long

Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOP) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As enmDriveType
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Any, lpLastAccessTime As Any, lpLastWriteTime As Any) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long

Private Declare Function GetShortPathNameW Lib "kernel32" (ByVal lpszLongPath As Long, ByVal lpszShortPath As Long, ByVal cchBuffer As Long) As Long

Private Enum enmDriveType
  DRIVE_UNKNOWN
  DRIVE_NO_ROOT_DIR
  DRIVE_REMOVABLE
  DRIVE_FIXED
  DRIVE_REMOTE
  DRIVE_CDROM
  DRIVE_RAMDISK
End Enum


Function Shell2(ByVal Path As String, Optional wStyle As Integer = 1)
    Call ShellExecute(0, "", Path, "", "", wStyle)
End Function

Public Function PathExist(FileName As String) As Boolean: PathExist = PathFileExists(FileName) <> 0: End Function
Public Function PathDirectory(pzpPath As String) As Boolean: PathDirectory = PathIsDirectory(pzpPath): End Function

Public Function GetFilePath(ByVal Path As String) As String
'On Error Resume Next
    GetFilePath = Left(Path, Len(Path) - Len(Right(Path, Len(Path) - InStrRev(Path, "\"))))
End Function
Function GetFileTitle(Path As String, Optional ByVal IncludeExt As Boolean = True) As String
Dim tmp As String
Dim ext    As Integer
     
    tmp = Right(Path, Len(Path) - InStrRev(Path, "\"))
    If Not IncludeExt Then ext = InStr(1, tmp, ".")
    If ext <> 0 Then tmp = Left(tmp, ext - 1)
    GetFileTitle = tmp
End Function
Function GetFileExt(Path As String) As String
Dim ext    As Long
    ext = InStr(1, Path, ".")
    If ext <> 0 Then GetFileExt = Right(Path, Len(Path) - InStrRev(Path, "."))
End Function
Function GetFileName(Path As String) As String
Dim lRet As String
        lRet = Right(Path, Len(Path) - InStrRev(Path, "\"))
        GetFileName = lRet
End Function

Public Function GetSafeFileName(sName As String) As String
    Dim s As String
    s = Replace(sName, "/", "")
    s = Replace(s, "\", "")
    s = Replace(s, ":", ";")
    s = Replace(s, "*", "")
    s = Replace(s, "?", "")
    s = Replace(s, Chr(34), "''")
    s = Replace(s, "<", "")
    s = Replace(s, ">", "")
    s = Replace(s, "|", "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    GetSafeFileName = s
End Function

Public Function RenamePath(sPath As String, sNewPath As String) As Boolean
On Error GoTo e
    Name sPath As sNewPath
    RenamePath = True
e:
End Function
Function FileDateTime2(Path As String) As String
On Error GoTo e
   FileDateTime2 = FileDateTime(Path)
e:
End Function
Public Function CustomFileDateTime(FileName As String) As String
On Local Error GoTo e
    CustomFileDateTime = Format(FileDateTime(FileName), "yyyymmddhhmmss")
    Exit Function
e:
    CustomFileDateTime = "19790101000000"
End Function

Public Function CreateDirectory(lpzPath As String) As Boolean
On Error GoTo e
    MkDir lpzPath
    CreateDirectory = True
e:
    If Err.Number Then Debug.Print "CreateDir::" & Err.Description: Err.Clear
End Function

Public Function GetShortPath(sLongPath As String) As String
Dim sShortPath As String
Dim lRet As Long

    lRet = GetShortPathNameW(StrPtr(sLongPath), 0, 0)
    If lRet Then
        sShortPath = Space$(lRet - 1)
        lRet = GetShortPathNameW(StrPtr(sLongPath), StrPtr(sShortPath), lRet)
        If lRet Then GetShortPath = sShortPath
    End If
End Function

Public Function SystemFolder(ByVal uPath As SystemFolderID) As String
Dim sPath   As String * 560
Dim Idl     As Long

    If SHGetSpecialFolderLocation(0, uPath, Idl) = 0 Then
        If SHGetPathFromIDList(ByVal Idl, ByVal sPath) Then
            SystemFolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
        End If
    End If
End Function

Public Function MoveSafeFile(oPath As String, dPath As String, Optional hWnd As Long) As Boolean
Dim uFoP As SHFILEOP
Dim cnt  As Long
Dim tmp  As String

    tmp = dPath
    Do While PathExist(tmp)
        cnt = cnt + 1
        tmp = dPath & "_" & cnt
        DoEvents
    Loop
    dPath = tmp
    
    With uFoP
        .hWnd = hWnd
        .wFunc = FO_MOVE
        .lflags = FOF_RENAMEONCOLLISION + FOF_SILENT + FOF_NOERRORUI + FOF_NOCONFIRMATION
        .pFrom = oPath & vbNullChar & vbNullChar
        .pTo = dPath & vbNullChar & vbNullChar
        .lpszProgressTitle = " "
    End With
    
    MoveSafeFile = SHFileOperation(uFoP) = 0
    
End Function


Public Function Drives() As Collection
Dim mDrv()  As String
Dim tmp     As String * 255
Dim vElm    As Variant
Dim i       As Long

    i = GetLogicalDriveStrings(255, tmp)
    mDrv = Split(Left$(tmp, i - 1), Chr$(0))
    
    Set Drives = New Collection
    ReDim vElm(1)
    
    For i = 0 To UBound(mDrv)
        vElm(0) = Left$(mDrv(i), 1)
        vElm(1) = DriveLabel(mDrv(i))
        Drives.Add vElm
    Next
    
End Function

Public Function DriveLabel(drive_path As String) As String
Dim tmp As String
    tmp = String$(255, Chr$(0))
    GetVolumeInformation Left(drive_path, 1) & ":\", tmp, 255, 0, 0, 0, "", 255
    tmp = Left$(tmp, InStr(1, tmp, Chr$(0)) - 1)
    If tmp = vbNullString Then
        Select Case GetDriveType(Left(drive_path, 1) & ":\")
            Case 2: tmp = "Unidad Extraible"
            Case 3: tmp = "Disco Local"
            Case 4: tmp = "Disco Remoto"
            Case 5: tmp = "CD-DVD" '"Unidad CD-DVD"
            Case 6: tmp = "RAMDISK"
            Case Else: tmp = "Unidad"
        End Select
    End If
    DriveLabel = tmp
End Function

Public Function ToRecycleBin(FileName As String, Optional Confirm As Boolean = False, Optional Silent As Boolean = True) As Boolean
'On Error Resume Next
Dim FO As SHFILEOP

    With FO
        .wFunc = &H3 'FO_DELETE
        .pFrom = FileName
        .lflags = True
        If Not Confirm Then .lflags = .lflags + &H10    'FOF_NOCONFIRMATION
        If Silent Then .lflags = .lflags + &H4          'FOF_SILENT
    End With
    DoEvents
    ToRecycleBin = SHFileOperation(FO) = 0
End Function

Public Function SetLastWriteFileDateTime(sPath As String, ByVal sTime As String) As Boolean
On Error GoTo e
Dim tTime As SYSTEMTIME
Dim lFile As Long
Dim tOF   As OFSTRUCT
Dim lTime As FILETIME

    With tTime
        .wYear = mvStripStr(sTime, 4)
        .wMonth = mvStripStr(sTime, 2)
        .wDay = mvStripStr(sTime, 2)
        .wHour = mvStripStr(sTime, 2)
        .wMinute = mvStripStr(sTime, 2)
        .wSecond = mvStripStr(sTime, 2)
    End With
    
    lFile = OpenFile(sPath, tOF, &H2)
    If lFile = &HFFFF Then Exit Function
    
    SystemTimeToFileTime tTime, lTime
    LocalFileTimeToFileTime lTime, lTime
    SetLastWriteFileDateTime = SetFileTime(lFile, 0&, 0&, lTime) <> 0
    CloseHandle lFile
e:
End Function

Public Function CreateShortCut(ByVal TargetPath As String, ByVal ShortCutPath As String, ByVal ShortCutname As String, Optional ByVal WorkPath As String, Optional ByVal Window_Style As Integer, Optional IconPath As String, Optional ByVal IconNum As Integer, Optional Args As String) As Boolean
On Error GoTo e
Dim MyShortcut  As Object
Dim VbsObj      As Object
    
    Set VbsObj = CreateObject("WScript.Shell")
    Set MyShortcut = VbsObj.CreateShortCut(ShortCutPath & "\" & ShortCutname & ".lnk")
    With MyShortcut
        .TargetPath = TargetPath
        .WorkingDirectory = WorkPath
        .WindowStyle = Window_Style
        If Trim(IconPath) <> "" Then .IconLocation = IconPath & "," & IconNum
        .Arguments = Args
        .Save
    End With
    CreateShortCut = True
    Exit Function
e:
End Function
Public Function RemoveFile(sPath As String) As Boolean
On Error GoTo e_
    Kill sPath
    RemoveFile = True
e_:
End Function
Public Function FileLen2(PathName As String) As Currency
On Error GoTo e
    FileLen2 = FileLen(PathName)
    If FileLen2 < 0 Then FileLen2 = CCur(FileLen2 And &H7FFFFFFF) + 2147483648# 'Else FileLen2 = CCur(FileLenL)
e:
    On Error GoTo 0
End Function

Public Function mvStripStr(sData As String, ln As Long) As String
    mvStripStr = Left$(sData, ln)
    If ln < Len(sData) Then sData = Right$(sData, Len(sData) - ln) Else sData = vbNullString
End Function

