Attribute VB_Name = "ModMain"
Public AsmXor As New ClsAsmXor
Public PrecTimer As New ClsPrecTimer

'pure Visual Basic function for the encryption
Public Sub XorEnDecrypt(Str As String, Password As String)
    Dim i As Long, j As Long
    Dim s, p
    j = 0
    If Len(Password) Then
        For i = 1 To Len(Str)
            s = Mid$(Str, i, 1)
            p = Mid$(Password, j + 1, 1)
            Mid$(Str, i, 1) = Chr$(Asc(Mid$(Str, i, 1)) Xor Asc(Mid$(Password, j + 1, 1)))
            j = (j + 1) Mod Len(Password)
        Next i
    End If
End Sub

'this function creates a string with the
'given size depending on a template string
Public Function CreateString(SampleText As String, length As Long) As String
    On Error Resume Next
    Dim Str As String
    Dim i As Long
    Str = SampleText
    For i = 5 To Int(Log(length) / Log(2)) + 1
        Str = Str & Str
    Next i
    Str = Left$(Str, length)
    CreateString = Str
End Function

