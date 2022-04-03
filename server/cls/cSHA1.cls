VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : cSHA1
'    Autor      : J. Elihu
'    Description: SHA1 Hash
'--------------------------------------------------------------------------------


Option Explicit

Private Declare Function CryptAcquireContext Lib "advapi32" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long  'TRUE (<> 0) = success.  See Err.LastDLLError if FALSE.
Private Declare Function CryptCreateHash Lib "advapi32" (ByVal hProv As Long, ByVal algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long  'TRUE (<> 0) = success.  See Err.LastDLLError if FALSE.
Private Declare Function CryptDestroyHash Lib "advapi32" (ByVal hHash As Long) As Long  'TRUE (<> 0) = success.  See Err.LastDLLError if FALSE.
Private Declare Function CryptGetHashParam Lib "advapi32" (ByVal hHash As Long, ByVal dwParam As Long, ByRef pbData As Any, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptHashData Lib "advapi32" (ByVal hHash As Long, ByRef pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32" (ByVal hProv As Long, ByVal dwFlags As Long) As Long  'TRUE (<> 0) = success.  See Err.LastDLLError if FALSE.

Private Const ALG_TYPE_ANY          As Long = 0
Private Const ALG_CLASS_HASH        As Long = 32768

Private Const ALG_SID_MD2           As Long = 1
Private Const ALG_SID_MD4           As Long = 2
Private Const ALG_SID_MD5           As Long = 3
Private Const ALG_SID_SHA1          As Long = 4
Private Const ALG_SID_SHA_256       As Long = 12
Private Const ALG_SID_SHA_384       As Long = 13
Private Const ALG_SID_SHA_512       As Long = 14

Private Const CALG_SHA1             As Long = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1
Private Const HP_HASHVAL            As Long = 2
Private Const HP_HASHSIZE           As Long = 4
Private Const CHUNK                 As Long = 16384

Private m_hProvider As Long


Private Sub Class_Initialize()
Const CRYPT_VERIFYCONTEXT   As Long = &HF0000000

    CryptAcquireContext m_hProvider, vbNullString, "Microsoft Base Cryptographic Provider v1.0", 1&, CRYPT_VERIFYCONTEXT
    'cSha1.Hash("Hello world")
    'cSha1.Hash(StrConv("Hello world", vbFromUnicode))
End Sub

Private Sub Class_Terminate()
    If m_hProvider Then CryptReleaseContext m_hProvider, 0&
End Sub


Public Function Hash(ByRef data As String) As String
On Error GoTo e
Dim lHash   As Long
Dim lSize   As Long
Dim out()   As Byte
Dim i       As Integer

    If m_hProvider = 0 Then Exit Function
    If CryptCreateHash(m_hProvider, CALG_SHA1, 0&, 0&, lHash) = 0 Then GoTo e
    
    out() = data
    If CryptHashData(lHash, out(LBound(out)), UBound(out) - LBound(out) + 1, 0&) = 0 Then GoTo e
    If CryptGetHashParam(lHash, HP_HASHSIZE, lSize, 4&, 0&) = 0 Then GoTo e
    ReDim out(lSize - 1)
    If CryptGetHashParam(lHash, HP_HASHVAL, out(0), lSize, 0&) = 0 Then GoTo e
    For i = 0 To lSize - 1
        Hash = Hash & Right$("0" & Hex$(out(i)), 2)
    Next
    
e:  If lHash Then CryptDestroyHash lHash
End Function
Public Function HashFile(ByVal FileName As String) As String
On Error GoTo e
Dim lWholeChunks As Long
Dim lRemainder  As Long
Dim lngChunk    As Long
Dim lHash       As Long
Dim out()       As Byte
Dim lSize       As Long
Dim FF          As Integer
Dim i           As Long

    If m_hProvider = 0 Then Exit Function
    If Not Len(Dir$(FileName, vbNormal Or vbHidden Or vbReadOnly Or vbSystem)) > 0 Then Exit Function
    
    FF = FreeFile
    Open FileName For Binary Access Read As #FF
    lWholeChunks = LOF(FF) \ CHUNK
    lRemainder = LOF(FF) - (CHUNK * lWholeChunks)
        
    If CryptCreateHash(m_hProvider, CALG_SHA1, 0&, 0&, lHash) = 0 Then GoTo e
    ReDim out(CHUNK - 1)
    For i = 1 To lWholeChunks
        Get #FF, , out
        If CryptHashData(lHash, out(LBound(out)), UBound(out) - LBound(out) + 1, 0&) = 0 Then GoTo e
    Next
    
    If lRemainder > 0 Then
        ReDim out(lRemainder - 1)
        Get #FF, , out
        If CryptHashData(lHash, out(LBound(out)), UBound(out) - LBound(out) + 1, 0&) = 0 Then GoTo e
    End If
    Close #FF
    FF = 0
        
    If CryptGetHashParam(lHash, HP_HASHSIZE, lSize, 4&, 0&) = 0 Then GoTo e
    ReDim out(lSize - 1)
    If CryptGetHashParam(lHash, HP_HASHVAL, out(0), lSize, 0&) = 0 Then GoTo e
    For i = 0 To lSize - 1
        HashFile = HashFile & Right$("0" & Hex$(out(i)), 2)
    Next
    
e:
    On Error Resume Next
    If FF Then Close #FF
    If lHash Then CryptDestroyHash lHash
End Function

