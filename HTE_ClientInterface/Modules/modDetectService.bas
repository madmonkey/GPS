Attribute VB_Name = "modDetectService"
Option Explicit

Private Const SC_MANAGER_CONNECT = &H1&
Private Const SERVICE_QUERY_CONFIG = &H1&
Private Const ERROR_INSUFFICIENT_BUFFER = 122&
Private Const SERVICE_QUERY_STATUS = &H4&
Private Const SERVICE_START = &H10&
Private Const SERVICE_STOP = &H20&

Public Enum SERVICE_STATE
    SERVICE_NOTFOUND = &H0
    SERVICE_STOPPED = &H1
    SERVICE_START_PENDING = &H2
    SERVICE_STOP_PENDING = &H3
    SERVICE_RUNNING = &H4
    SERVICE_CONTINUE_PENDING = &H5
    SERVICE_PAUSE_PENDING = &H6
    SERVICE_PAUSED = &H7
End Enum

Private Type QUERY_SERVICE_CONFIG
    dwServiceType As Long
    dwStartType As Long
    dwErrorControl As Long
    lpBinaryPathName As Long
    lpLoadOrderGroup As Long
    dwTagId As Long
    lpDependencies As Long
    lpServiceStartName As Long
    lpDisplayName As Long
End Type

Private Type SERVICE_STATUS
    dwServiceType As Long
    dwCurrentState As Long
    dwControlsAccepted As Long
    dwWin32ExitCode As Long
    dwServiceSpecificExitCode As Long
    dwCheckPoint As Long
    dwWaitHint As Long
End Type

Private Enum SERVICE_CONTROL
    SERVICE_CONTROL_STOP = 1&
    SERVICE_CONTROL_PAUSE = 2&
    SERVICE_CONTROL_CONTINUE = 3&
    SERVICE_CONTROL_INTERROGATE = 4&
    SERVICE_CONTROL_SHUTDOWN = 5&
End Enum

Private Declare Function OpenSCManager Lib "advapi32" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function OpenService Lib "advapi32" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function QueryServiceConfig Lib "advapi32" Alias "QueryServiceConfigA" (ByVal hService As Long, lpServiceConfig As QUERY_SERVICE_CONFIG, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Private Declare Function CloseServiceHandle Lib "advapi32" (ByVal hSCObject As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function QueryServiceStatus Lib "advapi32" (ByVal hService As Long, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function ControlService Lib "advapi32" (ByVal hService As Long, ByVal dwControl As SERVICE_CONTROL, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function StartService Lib "advapi32" Alias "StartServiceA" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long

Public Function IsInstalledService() As Boolean
    'IsInstalledService = (GetServiceConfig() = 0)
    IsInstalledService = (GetServiceStatus <> SERVICE_NOTFOUND)
End Function

Private Function GetServiceConfig(Optional s As String) As Long
Dim hSCManager As Long, hService As Long
Dim r As Long, SCfg() As QUERY_SERVICE_CONFIG, r1 As Long

    hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    If hSCManager <> 0 Then
        hService = OpenService(hSCManager, SERVICE_NAME, SERVICE_QUERY_CONFIG)
        If hService <> 0 Then
            ReDim SCfg(1 To 1)
            If QueryServiceConfig(hService, SCfg(1), 36, r) = 0 Then
                If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
                    r1 = r \ 36 + 1
                    ReDim SCfg(1 To r1)
                    If QueryServiceConfig(hService, SCfg(1), r1 * 36, r) <> 0 Then
                        s = Space$(255)
                        lstrcpy s, SCfg(1).lpServiceStartName
                        s = Left$(s, lstrlen(s))
                    Else
                        GetServiceConfig = Err.LastDllError
                    End If
                Else
                    GetServiceConfig = Err.LastDllError
                End If
            End If
            CloseServiceHandle hService
        Else
            GetServiceConfig = Err.LastDllError
        End If
        CloseServiceHandle hSCManager
    Else
        GetServiceConfig = Err.LastDllError
    End If
    
End Function

Public Function GetServiceStatus() As SERVICE_STATE
' This function returns current service status or 0 on error
Dim hSCManager As Long, hService As Long, Status As SERVICE_STATUS
    
    hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    If hSCManager <> 0 Then
        hService = OpenService(hSCManager, SERVICE_NAME, SERVICE_QUERY_STATUS)
        If hService <> 0 Then
            If QueryServiceStatus(hService, Status) Then
                GetServiceStatus = Status.dwCurrentState
            End If
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSCManager
    End If
    
End Function

Public Function StopNTService() As Long
' This function stops service it returns nonzero value on error
Dim hSCManager As Long, hService As Long, Status As SERVICE_STATUS
    hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    If hSCManager <> 0 Then
        hService = OpenService(hSCManager, SERVICE_NAME, SERVICE_STOP)
        If hService <> 0 Then
            If ControlService(hService, SERVICE_CONTROL_STOP, Status) = 0 Then
                StopNTService = Err.LastDllError
            End If
        CloseServiceHandle hService
        Else
            StopNTService = Err.LastDllError
        End If
    CloseServiceHandle hSCManager
    Else
        StopNTService = Err.LastDllError
    End If
End Function

Public Function StartNTService() As Long
' This function starts service it returns nonzero value on error
Dim hSCManager As Long, hService As Long
    
    hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    If hSCManager <> 0 Then
        hService = OpenService(hSCManager, SERVICE_NAME, SERVICE_START)
        If hService <> 0 Then
            If StartService(hService, 0, 0) = 0 Then
                StartNTService = Err.LastDllError
            End If
        CloseServiceHandle hService
        Else
            StartNTService = Err.LastDllError
        End If
    CloseServiceHandle hSCManager
    Else
        StartNTService = Err.LastDllError
    End If
   
End Function
