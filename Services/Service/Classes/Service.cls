VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Service"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim gApplication As HTE_ClientUtilities.Application
Dim WithEvents utilityTimer As ccrpTimers6.ccrpTimer
Attribute utilityTimer.VB_VarHelpID = -1

Private Sub InitializeApplication()
    Set gApplication = New HTE_ClientUtilities.Application
End Sub

Private Sub TerminateApplication()
    If Not gApplication Is Nothing Then
        Set gApplication = Nothing
    End If
End Sub

Private Sub Class_Initialize()
    InitializeApplication
    Set utilityTimer = New ccrpTimers6.ccrpTimer
    With utilityTimer
        .EventType = TimerPeriodic
        .Interval = 5000
        .Enabled = True
    End With
End Sub

Private Sub Class_Terminate()
    utilityTimer.Enabled = False
    Set utilityTimer = Nothing
    TerminateApplication
End Sub

Private Sub utilityTimer_Timer(ByVal Milliseconds As Long)
    If Not FindApp("HTE_CLIENTUTILITIES.EXE") Then
        TerminateApplication
        InitializeApplication
    End If
End Sub
