VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Ultra-Fast Xor Encryption"
   ClientHeight    =   3555
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6420
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   237
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   428
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComDlg.CommonDialog CDlg 
      Left            =   5880
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FraFile 
      Caption         =   "File En-/Decryption"
      Height          =   1455
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton CmdSelFileTo 
         Caption         =   "..."
         Height          =   255
         Left            =   3960
         TabIndex        =   21
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox TxtFileTo 
         Height          =   285
         Left            =   600
         TabIndex        =   19
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox TxtFile 
         Height          =   285
         Left            =   600
         TabIndex        =   6
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton CmdSelFile 
         Caption         =   "..."
         Height          =   255
         Left            =   3960
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox TxtFilePassword 
         Height          =   285
         Left            =   600
         TabIndex        =   8
         Text            =   "Password"
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton CmdCryptFile 
         Caption         =   "En-/Decrypt"
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton CmdCreateTestFile 
         Caption         =   "Testfile"
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label LblTo 
         Caption         =   "To:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   375
      End
      Begin VB.Label LblFile 
         Caption         =   "File:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   375
      End
      Begin VB.Label LblFilePassword 
         Caption         =   "Pass- word:"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   615
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame FraMethods 
      Caption         =   "Methods(for Comp.)"
      Height          =   1455
      Left            =   4680
      TabIndex        =   15
      Top             =   120
      Width           =   1575
      Begin VB.CheckBox ChkVBASM 
         Caption         =   "VB + ASM"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Value           =   1  'Aktiviert
         Width           =   1095
      End
      Begin VB.CheckBox ChkPureVB 
         Caption         =   "pure VB"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Value           =   1  'Aktiviert
         Width           =   1095
      End
   End
   Begin VB.TextBox TxtPassword 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Password"
      Top             =   1920
      Width           =   1335
   End
   Begin MSComctlLib.ListView LvwResults 
      Height          =   1215
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2143
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "Compare"
      Height          =   405
      Left            =   4440
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox TxtStrLen 
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Text            =   "100000"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox TxtIterations 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "1"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label LblPassword 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label LblStrLen 
      Caption         =   "String Length:"
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label LblIterations 
      Caption         =   "Iterations:"
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCreateTestFile_Click()
    Dim TestFile As String, InputString As String
    Dim Data As String
    Dim i As Long, FileSize As Long
    TestFile = App.Path
    If Right(TestFile, 1) <> "\" Then TestFile = TestFile & "\"
    TestFile = TestFile & "TestFile.txt"
    InputString = InputBox("Filesize in kb:", , "1024")
    If StrPtr(InputString) = 0 Then Exit Sub
    If Not IsNumeric(InputString) Then Exit Sub
    FileSize = CLng(InputString)
    If Dir(TestFile) <> vbNullString Then Kill TestFile
    Data = CreateString("This is a Testfile ", 1024)
    Open TestFile For Binary As #1
    For i = 1 To FileSize
        Put #1, , Data
    Next i
    Close #1
    TxtFile.Text = TestFile
    MsgBox "Testfile created."
End Sub

Private Sub CmdCryptFile_Click()
    Dim pt As New ClsPrecTimer
    Dim ret As Boolean
    Dim FileSize As Long
    If (TxtFile.Text = "") Or (Dir(TxtFile.Text) = vbNullString) Then
        MsgBox "Can't open file."
        Exit Sub
    End If
    If TxtFilePassword.Text = vbNullString Then
        MsgBox "No password specified."
        Exit Sub
    End If
    pt.ResetTimer
    Call AsmXor.EnDeCryptFile(TxtFile.Text, TxtFileTo.Text, TxtFilePassword.Text, True)
    pt.StopTimer
    FileSize = FileLen(TxtFile.Text)
    MsgBox IIf(FileSize >= 1048576, FileSize \ (1048576), _
            FileSize \ 1024) & " " & IIf(FileSize >= 1048576, "MB", "KB") _
            & " encrypted in " & Format$(pt.Elapsed / 1000, "###0.00 ms")
End Sub

Private Sub CmdSelFile_Click()
    CDlg.CancelError = False
    CDlg.InitDir = App.Path
    CDlg.FileName = ""
    CDlg.ShowOpen
    If CDlg.FileName <> vbNullString Then
        TxtFile.Text = CDlg.FileName
    End If
End Sub

Private Sub CmdSelFileTo_Click()
    CDlg.CancelError = False
    CDlg.InitDir = App.Path
    CDlg.FileName = ""
    CDlg.ShowOpen
    If CDlg.FileName <> vbNullString Then
        TxtFileTo.Text = CDlg.FileName
    End If
End Sub

Private Sub CmdStart_Click()
    Dim ret As Long, Iterations As Long, length As Long
    Dim TimeVB As Single, TimeASM As Single, Max As Single
    Dim Str As String, Pass As String
    Dim PrecTimer As New ClsPrecTimer
    Dim ctrl As Control
    If (TxtStrLen.Text <> vbNullString) And (IsNumeric(TxtStrLen.Text)) Then
        length = CLng(TxtStrLen.Text)
    Else
        MsgBox "Invalid string length."
        Exit Sub
    End If
    If (TxtIterations.Text <> vbNullString) And (IsNumeric(TxtIterations.Text)) Then
        Iterations = CLng(TxtIterations.Text)
    Else
        MsgBox "Invalid number of iterations."
        Exit Sub
    End If
    If TxtPassword.Text = "" Then
        MsgBox "No password specified."
        Exit Sub
    End If
    Pass = TxtPassword.Text
    For Each ctrl In Me.Controls
        If Not TypeOf ctrl Is CommonDialog Then
            ctrl.Enabled = False
        End If
    Next ctrl
    Me.Refresh
    Screen.MousePointer = vbHourglass
    LvwResults.ListItems.Clear
    Str = CreateString("This is a string", length)
    If ChkVBASM.Value = Checked Then
        PrecTimer.ResetTimer
        For i = 1 To Iterations
            AsmXor.EnDeCrypt Str, Pass
        Next i
        PrecTimer.StopTimer
        TimeASM = PrecTimer.Elapsed
    End If
    If ChkPureVB.Value = Checked Then
        PrecTimer.ResetTimer
        For i = 1 To Iterations
            XorEnDecrypt Str, Pass
        Next i
        PrecTimer.StopTimer
        TimeVB = PrecTimer.Elapsed
    End If
    Max = TimeASM
    If Max < TimeVB Then Max = TimeVB
    If ChkPureVB.Value = Checked Then
        LvwResults.ListItems.Add , , "pure VB"
        LvwResults.ListItems(LvwResults.ListItems.Count).SubItems(1) = Format$(TimeVB / 1000, "###0.00")
        LvwResults.ListItems(LvwResults.ListItems.Count).SubItems(2) = Format$(TimeVB / Max, "#0.0%")
        LvwResults.ListItems(LvwResults.ListItems.Count).SubItems(3) = Format$(Iterations * length / (TimeVB / 1000000) / (1048576), "###0.000 MB/s")
    End If
    If ChkVBASM.Value = Checked Then
        LvwResults.ListItems.Add , , "VB + ASM"
        LvwResults.ListItems(LvwResults.ListItems.Count).SubItems(1) = Format$(TimeASM / 1000, "###0.00")
        LvwResults.ListItems(LvwResults.ListItems.Count).SubItems(2) = Format$(TimeASM / Max, "#0.0%")
        LvwResults.ListItems(LvwResults.ListItems.Count).SubItems(3) = Format$(Iterations * length / (TimeASM / 1000000) / (1048576), "###0.000 MB/s")
    End If
    For Each ctrl In Me.Controls
        If Not TypeOf ctrl Is CommonDialog Then
            ctrl.Enabled = True
        End If
    Next ctrl
    Screen.MousePointer = vbDefault
End Sub

Private Sub TxtFile_Change()
    TxtFileTo.Text = Left(TxtFile.Text, InStrRev(TxtFile.Text, ".") - 1) _
            & "Enc" & Mid(TxtFile.Text, InStrRev(TxtFile.Text, "."))
End Sub

Private Sub Form_Load()
    LvwResults.ColumnHeaders.Add , , "Method"
    LvwResults.ColumnHeaders.Add , , "Time (ms)"
    LvwResults.ColumnHeaders.Add , , "Time (%)"
    LvwResults.ColumnHeaders.Add , , "Speed"
    LvwResults.GridLines = True
    LvwResults.View = lvwReport
End Sub

