VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTiming"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private seqFreq As Currency
Private seqStart As Currency
Private fAvailable As Boolean
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Class_Initialize()
' determine if system has Performance Counter and get its frequency returns ticks/sec
' note: the Currency var displays this as "119.318", gotta multiply by 10,000 to get the actual value
    If QueryPerformanceFrequency(seqFreq) = 0& Then
        'use alternative method
        seqFreq = 1000
    Else
        fAvailable = True
    End If
End Sub

Friend Sub Reset()
    If fAvailable Then
        QueryPerformanceCounter seqStart
    Else
        seqStart = GetTickCount
    End If
End Sub

Friend Function Elapsed(Optional DegreesOfPrecision As Long = 4) As Double 'more handy than Currency
' return elapsed time in seconds
Dim seqNow As Currency
    
    If fAvailable Then
        QueryPerformanceCounter seqNow
        Elapsed = Round(Abs(seqNow - seqStart) / seqFreq, DegreesOfPrecision)
    Else
        seqNow = GetTickCount
        Elapsed = Round(Abs(seqNow - seqStart) / seqFreq, DegreesOfPrecision)
    End If
    
End Function

Friend Function sElapsed() As String
' returns a nicely formatted string of seconds - only 1st 3 digits of decimal are of any value!
  sElapsed = Format$(Elapsed, "#,0.000")
End Function


