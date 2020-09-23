VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAsmXor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal ptrMC As Long, ByVal P1 As Long, ByVal P2 As Long, ByVal P3 As Long, ByVal P4 As Long) As Long
Const BinFile = "fastxor.bin"
Const CodeString = "5589E55356578B450C85C074468B451485C0743F8B4D0CD1E18B550801CAF7D98B5D14D1E38B451001D8894510F7DB895D148B040A035D1032032B5D1081C30200000075038B5D1489040A81C10200000075DF31C05F5E5B89EC5DC21000"
Private BinCode() As Byte ' Byte array to hold Machine Code
Private ptrMC As Long ' Pointer to Machine Code
Private Initialized As Boolean

Private Sub Class_Initialize()
    'while developing the asm-code you should specify the
    'BinFile constant and set the Codestring constant to
    'a nullstring, so the BinCode array is generated
    'dynamically from the binary file each time the
    'class is initialized. after developing you should
    'call the GenCodeString function and paste the data
    'into the Codestring constant to be independent
    'of the file with the binary data.
    If CodeString = vbNullString Then
        LoadMCode
    Else
        LoadCodeString
    End If
    If ptrMC <> 0 Then Initialized = True
End Sub

Public Function EnDeCrypt(Str As String, Password As String) As Boolean
    If Not Initialized Then Exit Function
    Dim ret As Long
    If (Len(Str) = 0) Or (Len(Password) = 0) Then Exit Function
    ret = CallWindowProc(ptrMC, StrPtr(Str), Len(Str), StrPtr(Password), Len(Password))
    EnDeCrypt = True
End Function

Public Function EnDeCryptFile(FileFrom As String, FileTo As String, Password As String, Optional OverWrite As Boolean = False) As Boolean
    If Not Initialized Then Exit Function
    Dim BuffSize As Long
    Dim FileSize As Long, FileNumFrom As Long
    Dim FileNumTo As Long, i As Long
    Dim BuffStr As String
    BuffSize = 8192
    'Buffsize mod Len(Password) should be zero
    'to be compatible to Buffers with another size
    BuffSize = BuffSize - (BuffSize Mod Len(Password))
    BuffStr = String(BuffSize, " ")
    If FileFrom = vbNullString Then Exit Function
    If FileTo = vbNullString Then Exit Function
    If Dir(FileFrom) = vbNullString Then Exit Function
    'delete existing files if overwrite is set to true
    If Dir(FileTo) <> vbNullString Then
        If OverWrite Then
            Kill FileTo
        Else
            Exit Function
        End If
    End If
    FileSize = FileLen(FileFrom)
    FileNumFrom = FreeFile
    Open FileFrom For Binary Access Read As FileNumFrom
    FileNumTo = FreeFile
    Open FileTo For Binary Access Write As FileNumTo
    'encyrpt while we have enough bytes
    'to fill the complete buffer
    For i = 1 To FileSize \ BuffSize
        Get FileNumFrom, , BuffStr
        EnDeCrypt BuffStr, Password
        Put FileNumTo, , BuffStr
    Next i
    'adapt the size of the buffer
    'and encrypt the rest
    If FileSize Mod BuffSize > 0 Then
        BuffStr = String(FileSize Mod BuffSize, " ")
        Get FileNumFrom, , BuffStr
        EnDeCrypt BuffStr, Password
        Put FileNumTo, , BuffStr
    End If
    Close FileNumFrom
    Close FileNumTo
    EnDeCryptFile = True
End Function

Private Sub LoadCodeString()
    'this sub loads the machine code from the CodeString
    'constant. it converts the hex-values to byte-values,
    'puts them into the BinCode array and lets ptrMC
    'point to the first byte.
    Dim i As Long
    If Len(CodeString) = 0 Then Exit Sub
    ReDim BinCode(Len(CodeString) \ 2 - 1)
    For i = 0 To Len(CodeString) \ 2 - 1
        BinCode(i) = CByte("&H" & Mid(CodeString, i * 2 + 1, 2))
    Next i
    ptrMC = VarPtr(BinCode(0))
    Initialized = True
End Sub

Private Sub LoadMCode()
    'this sub loads the machine code from the binary
    'file generated with nasm into a byte array
    'and lets ptrMC point to the first byte
    Dim MCSize As Long
    If Dir(BinFile) = "" Then
        MsgBox (BinFile & " missing")
        Exit Sub
    End If
    Open BinFile For Binary As #1
    MCSize = LOF(1)
    If MCSize = 0 Then
        MsgBox (InFile & " missing")
        Exit Sub
    End If
    ReDim BinCode(MCSize - 1)
    Get #1, , BinCode
    Close #1
    ptrMC = VarPtr(BinCode(0))
    Initialized = True
End Sub

Private Function CHex(Number As Byte) As String
    If Number <= 15 Then
        CHex = "0" & Hex(Number)
    Else
        CHex = Hex(Number)
    End If
End Function

Public Function GenCodeString(Optional CopyToClipboard As Boolean = True) As String
    'this function generates the constant used in the declarations,
    'it returns the string and copies it to the clipboard.
    'you can call "AsmXor.GenCodeString" from the debug window.
    Dim i As Long
    Dim Temp As String
    LoadMCode
    For i = 0 To UBound(BinCode)
        Temp = Temp & CHex(BinCode(i))
    Next i
    If CopyToClipboard Then
        Clipboard.Clear
        Clipboard.SetText Temp
    End If
    GenCodeString = Temp
End Function
