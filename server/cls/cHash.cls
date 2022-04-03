VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'--------------------------------------------------------------------------------
'    Component  : cHASH
'    Autor      : J. Elihu
'    Description: advapi32 hash class
'--------------------------------------------------------------------------------

Option Explicit

Private Declare Function CryptAcquireContext Lib "advapi32" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long  'TRUE (<> 0) = success.  See Err.LastDLLError if FALSE.
Private Declare Function CryptCreateHash Lib "advapi32" (ByVal hProv As Long, ByVal algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long  'TRUE (<> 0) = success.  See Err.LastDLLError if FALSE.
Private Declare Function CryptDestroyHash Lib "advapi32" (ByVal hHash As Long) As Long  'TRUE (<> 0) = success.  See Err.LastDLLError if FALSE.
Private Declare Function CryptGetHashParam Lib "advapi32" (ByVal hHash As Long, ByVal dwParam As Long, ByRef pbData As Any, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptHashData Lib "advapi32" (ByVal hHash As Long, ByRef pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32" (ByVal hProv As Long, ByVal dwFlags As Long) As Long  'TRUE (<> 0) = success.  See Err.LastDLLError if FALSE.

Private Const PROV_RSA_FULL   As Long = 1
Private Const PROV_RSA_AES    As Long = 24
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000

Private Const HP_HASHVAL      As Long = 2
Private Const HP_HASHSIZE     As Long = 4

Private Const ALG_TYPE_ANY    As Long = 0
Private Const ALG_CLASS_HASH  As Long = 32768

Private Const ALG_SID_MD2     As Long = 1
Private Const ALG_SID_MD4     As Long = 2
Private Const ALG_SID_MD5     As Long = 3
Private Const ALG_SID_SHA_1   As Long = 4
Private Const ALG_SID_SHA_256 As Long = 12
Private Const ALG_SID_SHA_384 As Long = 13
Private Const ALG_SID_SHA_512 As Long = 14

Private Const CALG_MD2        As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2)
Private Const CALG_MD4        As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4)
Private Const CALG_MD5        As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5)
Private Const CALG_SHA_1      As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_1)
Private Const CALG_SHA_256    As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_256)
Private Const CALG_SHA_384    As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_384)
Private Const CALG_SHA_512    As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_512)

Public Enum enmHASH
    Md5 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5)
    Sha1 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_1)
    Sha256 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_256)
    Sha384 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_384)
    Sha512 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_512)
End Enum


Private Sub Class_Initialize()
'
End Sub
Private Sub Class_Terminate()
'
End Sub

Public Function Md5()
'
End Function
Public Function Sha1()
'
End Function

Public Function Hash(ByRef data As String, ByVal HashType As enmHASH)
Dim lProv   As Long
Dim lHash   As Long
Dim lSize   As Long
Dim out()   As Byte
Dim i       As Integer


    If CryptAcquireContext(lProv, vbNullString, vbNullString, IIf(HashType >= CALG_SHA_256, PROV_RSA_AES, PROV_RSA_FULL), &HF0000000) = 0 Then GoTo exit_
    If CryptCreateHash(lProv, HashType, 0&, 0&, lHash) = 0 Then GoTo exit_
    out() = data
    If CryptHashData(lHash, out(LBound(out)), UBound(out) - LBound(out) + 1, 0&) = 0 Then GoTo exit_
    If CryptGetHashParam(lHash, HP_HASHSIZE, lSize, 4&, 0&) = 0 Then GoTo exit_
    
    ReDim out(lSize - 1)
    If CryptGetHashParam(lHash, HP_HASHVAL, out(0), lSize, 0&) = 0 Then GoTo exit_
    For i = 0 To lSize - 1
        Hash = Hash & Right$("0" & Hex$(out(i)), 2)
    Next
exit_:
    If lHash Then CryptDestroyHash lHash
    If lProv Then CryptReleaseContext lProv, 0&
    
End Function




Private Function CreateHash(abytData() As Byte, ByVal lngAlgID As Long) As String
Dim hProv       As Long
Dim hHash       As Long
Dim abytHash(0 To 63) As Byte
Dim lngLength   As Long
Dim lngResult   As Long
Dim strHash     As String
Dim i           As Long

    strHash = ""
    If CryptAcquireContext(hProv, vbNullString, vbNullString, _
                           IIf(lngAlgID >= CALG_SHA_256, PROV_RSA_AES, PROV_RSA_FULL), _
                           CRYPT_VERIFYCONTEXT) <> 0& Then
        If CryptCreateHash(hProv, lngAlgID, 0&, 0&, hHash) <> 0& Then
            lngLength = UBound(abytData()) - LBound(abytData()) + 1
            If lngLength > 0 Then lngResult = CryptHashData(hHash, abytData(LBound(abytData())), lngLength, 0&) _
                             Else lngResult = CryptHashData(hHash, ByVal 0&, 0&, 0&)
            If lngResult <> 0& Then
                lngLength = UBound(abytHash()) - LBound(abytHash()) + 1
                If CryptGetHashParam(hHash, HP_HASHVAL, abytHash(LBound(abytHash())), lngLength, 0&) <> 0& Then
                    For i = 0 To lngLength - 1
                        strHash = strHash & Right$("0" & Hex$(abytHash(LBound(abytHash()) + i)), 2)
                    Next
                End If
            End If
            CryptDestroyHash hHash
        End If
        CryptReleaseContext hProv, 0&
    End If
    CreateHash = LCase$(strHash)
End Function

