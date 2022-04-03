Attribute VB_Name = "mEncrypt"
Option Explicit


Private Const ENCRYPT_KEY As String = "0elihu0"

Public Enum enmEncryptAction
    ACTION_ENCRYPT = 1
    ACTION_DECRYPT = 2
End Enum



Public Function EncryptStr(ByVal Text As String, Action As enmEncryptAction) As String
    If Len(Text) = 0 Then Exit Function
    If Action = ACTION_DECRYPT Then pp_Shrink Text
    ppSetKeys
    Call pp_Xor(Text)
    If Action = ACTION_ENCRYPT Then pp_Stretch Text
    EncryptStr = Text
End Function

Private Sub ppSetKeys()
Dim i As Long
    Randomize Rnd(-1)
    For i = 1 To Len(ENCRYPT_KEY): Randomize Rnd(-Rnd * Asc(Mid(ENCRYPT_KEY, i, 1))): Next
End Sub
Private Sub pp_Xor(pzp As String)
Dim c As Long
Dim B As Long
Dim i As Long
    
     For i = 1 To Len(pzp)
        c = Asc(Mid(pzp, i, 1))
        B = Int(Rnd * 256)
        Mid(pzp, i, 1) = Chr(c Xor B)
     Next
End Sub

'Printable String
Private Sub pp_Stretch(pzp As String)
'On Error Resume Next
Dim tmp As String
Dim c   As Long
Dim j   As Long
Dim k   As Long
Dim l   As Long
Dim i   As Long

    l = Len(pzp)
    tmp = Space(l + (l + 2) \ 3)
    
    For i = 1 To l
    
        c = Asc(Mid(pzp, i, 1))
        j = j + 1
        Mid(tmp, j, 1) = Chr((c And 63) + 59)
        Select Case i Mod 3
            Case 1: k = k Or ((c \ 64) * 16)
            Case 2: k = k Or ((c \ 64) * 4)
            Case 0
                k = k Or (c \ 64)
                j = j + 1
                Mid(tmp, j, 1) = Chr(k + 59)
                k = 0
        End Select
    Next
    
    If l Mod 3 Then
        j = j + 1
        Mid(tmp, j, 1) = Chr(k + 59)
    End If
    pzp = tmp
    
End Sub

'\ Stretch Inverse
Public Sub pp_Shrink(pzp As String)
Dim tmp As String
Dim c   As Long
Dim d   As Long
Dim e   As Long
Dim B   As Long
Dim j   As Long
Dim k   As Long
Dim l   As Long
Dim i   As Long

    If Len(pzp) = 0 Then Exit Sub
    l = Len(pzp)
    B = l - 1 - (l - 1) \ 4
    tmp = Space(B)

    For i = 1 To B
        j = j + 1
        c = Asc(Mid(pzp, j, 1)) - 59

        Select Case i Mod 3
            Case 1
                k = k + 4
                If k > l Then k = l
                e = Asc(Mid(pzp, k, 1)) - 59
                d = ((e \ 16) And 3) * 64
            Case 2
                d = ((e \ 4) And 3) * 64
            Case 0
                d = (e And 3) * 64
                j = j + 1
        End Select
        Mid(tmp, i, 1) = Chr(c Or d)
    Next
     pzp = tmp
End Sub


