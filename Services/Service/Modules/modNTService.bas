Attribute VB_Name = "modNTService"
Option Explicit
Private Declare Function CreateThread Lib "kernel32" (ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private ServiceStatus As SERVICE_STATUS
Private hServiceStatus As Long

Function FncPtr(ByVal fnp As Long) As Long
' The FncPtr function returns function pointer.
    FncPtr = fnp
End Function

Public Function StartAsService() As Long
' The StartAsService function creates Service Dispatcher thread.
    Dim ThreadId As Long
    StartAsService = CreateThread(0&, 0&, AddressOf ServiceThread, 0&, 0&, ThreadId)
End Function

Private Sub ServiceThread(ByVal dummy As Long)
' The ServiceThread sub starts the service.
' This sub returns control only after service termination.
    Dim ServiceTableEntry As SERVICE_TABLE
    ServiceTableEntry.lpServiceName = ServiceNamePtr
    ServiceTableEntry.lpServiceProc = FncPtr(AddressOf ServiceMain)
    StartServiceCtrlDispatcher ServiceTableEntry
End Sub

Private Sub ServiceMain(ByVal dwArgc As Long, ByVal lpszArgv As Long)
    ' The ServiceMain sub - main service sub.
    ' It initializes service, sets event hStartEvent, and waits hStopEvent event.
    ' When hStopEvent fires, this sub exits and service stops.
    ServiceStatus.dwServiceType = SERVICE_WIN32_OWN_PROCESS
    ServiceStatus.dwControlsAccepted = SERVICE_ACCEPT_STOP Or SERVICE_ACCEPT_SHUTDOWN
    ServiceStatus.dwWin32ExitCode = 0&
    ServiceStatus.dwServiceSpecificExitCode = 0&
    ServiceStatus.dwCheckPoint = 0&
    ServiceStatus.dwWaitHint = 0&
    hServiceStatus = RegisterServiceCtrlHandler(SERVICE_NAME, AddressOf Handler)
    SetServiceState SERVICE_START_PENDING
    ' Set hStartEvent. It unlocks main application thread which allows to do some work in it
    SetEvent hStartEvent
    ' Wait until hStopEvent fires
    WaitForSingleObject hStopEvent, INFINITE
End Sub
   
Private Sub Handler(ByVal fdwControl As Long)
' The Handler sub processes commands from Service Dispatcher.
' It sets event hStopEvent when processes command SERVICE_CONTROL_STOP or SERVICE_CONTROL_SHUTDOWN.
    Select Case fdwControl
        Case SERVICE_CONTROL_SHUTDOWN, SERVICE_CONTROL_STOP
            SetServiceState SERVICE_STOP_PENDING
            SetEvent hStopPendingEvent
        Case Else
            SetServiceState
    End Select
End Sub

Public Sub SetServiceState(Optional ByVal NewState As SERVICE_STATE = 0&)
    ' The SetServiceState sub changes service state.
    ' If parameter omitted, it confirms previous state.
    If NewState <> 0& Then ServiceStatus.dwCurrentState = NewState
    SetServiceStatus hServiceStatus, ServiceStatus
End Sub
