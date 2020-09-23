VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPrecTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class was posted here:
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=31266&lngWId=1
Option Explicit
Private seqFreq As Currency
Private seqStart As Currency
Private seqStop As Currency
Private latency As Double
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private Sub Class_Initialize()
    Dim i As Long
    Dim lngTotal As Double
    ' Establish whether or not the system supports
    ' a performance counter.  If it does, retrieve
    ' the tick frequency.
    If QueryPerformanceFrequency(seqFreq) = 0 Then
        ' System does not support performance timer
        MsgBox "Performance Counter not available", vbExclamation
    Else
        ' Find average latency of timer calls
        For i = 1 To 1000
            ResetTimer
            StopTimer
            lngTotal = lngTotal + Elapsed
        Next i
        latency = lngTotal / i
    End If
End Sub

Friend Sub ResetTimer()
    ' Record the start-time
    QueryPerformanceCounter seqStart
End Sub

Friend Sub StopTimer()
    ' Record the end-time
    QueryPerformanceCounter seqStop
End Sub

Friend Function Elapsed() As Double
    ' Report the elapsed time in µs
    Elapsed = (10000 * (seqStop - seqStart)) - latency
End Function

Friend Function strElapsed() As String
    ' Report the elapsed time in µs
    strElapsed = Format$((10000 * (seqStop - seqStart)) - latency, "###,###.00") & Space$(1) & Chr$(181) & "s"
End Function


