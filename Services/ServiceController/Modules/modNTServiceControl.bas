Attribute VB_Name = "modNTServiceControl"
Option Explicit

Private Const ERROR_SERVICE_DOES_NOT_EXIST = 1060&
Private Const SERVICE_WIN32_OWN_PROCESS = &H10&
'Private Const SERVICE_WIN32_SHARE_PROCESS = &H20&
'Private Const SERVICE_WIN32 = SERVICE_WIN32_OWN_PROCESS + SERVICE_WIN32_SHARE_PROCESS
'Private Const SERVICE_ACCEPT_STOP = &H1
'Private Const SERVICE_ACCEPT_PAUSE_CONTINUE = &H2
'Private Const SERVICE_ACCEPT_SHUTDOWN = &H4
Private Const SC_MANAGER_CONNECT = &H1&
Private Const SC_MANAGER_CREATE_SERVICE = &H2&
'Private Const SC_MANAGER_ENUMERATE_SERVICE = &H4
'Private Const SC_MANAGER_LOCK = &H8
'Private Const SC_MANAGER_QUERY_LOCK_STATUS = &H10
'Private Const SC_MANAGER_MODIFY_BOOT_CONFIG = &H20
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SERVICE_QUERY_CONFIG = &H1&
Private Const SERVICE_CHANGE_CONFIG = &H2&
Private Const SERVICE_QUERY_STATUS = &H4&
Private Const SERVICE_ENUMERATE_DEPENDENTS = &H8&
Private Const SERVICE_START = &H10&
Private Const SERVICE_STOP = &H20&
Private Const SERVICE_PAUSE_CONTINUE = &H40&
Private Const SERVICE_INTERROGATE = &H80&
Private Const SERVICE_USER_DEFINED_CONTROL = &H100&
Private Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
                                       SERVICE_QUERY_CONFIG Or _
                                       SERVICE_CHANGE_CONFIG Or _
                                       SERVICE_QUERY_STATUS Or _
                                       SERVICE_ENUMERATE_DEPENDENTS Or _
                                       SERVICE_START Or _
                                       SERVICE_STOP Or _
                                       SERVICE_PAUSE_CONTINUE Or _
                                       SERVICE_INTERROGATE Or _
                                       SERVICE_USER_DEFINED_CONTROL)
'Private Const SERVICE_AUTO_START As Long = 2
Private Const SERVICE_DEMAND_START As Long = 3
Private Const SERVICE_ERROR_NORMAL As Long = 1
Private Const ERROR_INSUFFICIENT_BUFFER = 122&

Private Enum SERVICE_CONTROL
    SERVICE_CONTROL_STOP = 1&
    SERVICE_CONTROL_PAUSE = 2&
    SERVICE_CONTROL_CONTINUE = 3&
    SERVICE_CONTROL_INTERROGATE = 4&
    SERVICE_CONTROL_SHUTDOWN = 5&
End Enum

Public Enum SERVICE_STATE
    SERVICE_STOPPED = &H1
    SERVICE_START_PENDING = &H2
    SERVICE_STOP_PENDING = &H3
    SERVICE_RUNNING = &H4
    SERVICE_CONTINUE_PENDING = &H5
    SERVICE_PAUSE_PENDING = &H6
    SERVICE_PAUSED = &H7
End Enum

Private Type SERVICE_STATUS
    dwServiceType As Long
    dwCurrentState As Long
    dwControlsAccepted As Long
    dwWin32ExitCode As Long
    dwServiceSpecificExitCode As Long
    dwCheckPoint As Long
    dwWaitHint As Long
End Type

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

Private Declare Function OpenSCManager Lib "advapi32" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function CreateService Lib "advapi32" Alias "CreateServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal lpDisplayName As String, ByVal dwDesiredAccess As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, ByVal lpdwTagId As String, ByVal lpDependencies As String, ByVal lp As String, ByVal lpPassword As String) As Long
Private Declare Function DeleteService Lib "advapi32" (ByVal hService As Long) As Long
Private Declare Function CloseServiceHandle Lib "advapi32" (ByVal hSCObject As Long) As Long
'** Change SERVICE_NAME as needed
Private Declare Function OpenService Lib "advapi32" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function QueryServiceConfig Lib "advapi32" Alias "QueryServiceConfigA" (ByVal hService As Long, lpServiceConfig As QUERY_SERVICE_CONFIG, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Private Declare Function QueryServiceStatus Lib "advapi32" (ByVal hService As Long, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function ControlService Lib "advapi32" (ByVal hService As Long, ByVal dwControl As SERVICE_CONTROL, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function StartService Lib "advapi32" Alias "StartServiceA" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
Private Declare Function NetWkstaUserGetInfo Lib "Netapi32" (ByVal reserved As Any, ByVal Level As Long, lpBuffer As Any) As Long
Private Declare Function NetApiBufferFree Lib "Netapi32" (ByVal lpBuffer As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Public AppPath As String


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

' This function fills Service Account field in form.
' It returns nonzero value on error

Public Function GetServiceConfig() As Long
Dim hSCManager As Long, hService As Long
Dim r As Long, SCfg() As QUERY_SERVICE_CONFIG, r1 As Long, s As String

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
                        frmServiceControl.txtAccount = s
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

' This function installs service on local computer
' It returns nonzero value on error
Public Function SetNTService() As Long
Dim hSCManager As Long
Dim hService As Long, DomainName As String

    If frmServiceControl.txtAccount <> "LocalSystem" Then
    ' Add domain name to account string
        If InStr(1, frmServiceControl.txtAccount, "\") = 0 Then
            DomainName = GetDomainName()
            If DomainName = vbNullString Then DomainName = "."
            frmServiceControl.txtAccount.Text = DomainName & "\" & frmServiceControl.txtAccount.Text
        End If
    End If
    hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CREATE_SERVICE)
    If hSCManager <> 0 Then
        ' Install service to manual start. To set service to autostart replace SERVICE_DEMAND_START to SERVICE_AUTO_START
        hService = CreateService(hSCManager, SERVICE_NAME, SERVICE_DISPLAY_NAME, _
            SERVICE_ALL_ACCESS, SERVICE_WIN32_OWN_PROCESS, _
            SERVICE_DEMAND_START, SERVICE_ERROR_NORMAL, _
            AppPath & SERVICE_FILE_NAME, vbNullString, vbNullString, vbNullString, _
            frmServiceControl.txtAccount, frmServiceControl.txtPassword)
        If hService <> 0 Then
            CloseServiceHandle hService
        Else
            SetNTService = Err.LastDllError
        End If
        CloseServiceHandle hSCManager
    Else
        SetNTService = Err.LastDllError
    End If
    
End Function

' This function uninstalls service
' It returns nonzero value on error
Public Function DeleteNTService() As Long
Dim hSCManager As Long
Dim hService As Long, Status As SERVICE_STATUS

hSCManager = OpenSCManager(vbNullString, vbNullString, _
                       SC_MANAGER_CONNECT)
If hSCManager <> 0 Then
    hService = OpenService(hSCManager, SERVICE_NAME, _
                       SERVICE_ALL_ACCESS)
    If hService <> 0 Then
' Stop service if it is running
        ControlService hService, SERVICE_CONTROL_STOP, Status
        If DeleteService(hService) = 0 Then
            DeleteNTService = Err.LastDllError
        End If
        CloseServiceHandle hService
    Else
        DeleteNTService = Err.LastDllError
    End If
    CloseServiceHandle hSCManager
Else
    DeleteNTService = Err.LastDllError
End If

End Function

' This function returns local network domain name
' or zero-length string on error
Public Function GetDomainName() As String
Dim lpBuffer As Long, l As Long, p As Long
If NetWkstaUserGetInfo(0&, 1&, lpBuffer) = 0 Then
    CopyMemory p, ByVal lpBuffer + 4, 4
    l = lstrlenW(p)
    If l > 0 Then
        GetDomainName = Space$(l)
        CopyMemory ByVal StrPtr(GetDomainName), ByVal p, l * 2
    End If
    NetApiBufferFree lpBuffer
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

