Attribute VB_Name = "modNTServiceImpl"
Option Explicit

Public Const INFINITE = -1&      '  Infinite timeout
Public hStopEvent As Long, hStartEvent As Long, hStopPendingEvent As Long
Public ServiceNamePtr As Long

Private Const WAIT_TIMEOUT = 258&
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(1 To 128) As Byte      '  Maintenance string for PSS usage
End Type
Private Const VER_PLATFORM_WIN32_NT = 2&
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private IsNT As Boolean, IsNTService As Boolean
Private ServiceName() As Byte
Private m_Service As SunGardGPSService.Service

Private Sub Main()
    Dim hnd As Long
    Dim h(0 To 1) As Long
    ' Only one instance
    If App.PrevInstance Then Exit Sub
    ' Check OS type
    IsNT = CheckIsNT()
    ' Creating events
    hStopEvent = CreateEvent(0, 1, 0, vbNullString)
    hStopPendingEvent = CreateEvent(0, 1, 0, vbNullString)
    hStartEvent = CreateEvent(0, 1, 0, vbNullString)
    ServiceName = StrConv(SERVICE_NAME, vbFromUnicode)
    ServiceNamePtr = VarPtr(ServiceName(LBound(ServiceName)))
    If IsNT Then
        ' Trying to start service
        hnd = StartAsService
        h(0) = hnd
        h(1) = hStartEvent
        ' Waiting for one of two events: sucsessful service start (1) or
        ' terminaton of service thread (0)
        IsNTService = WaitForMultipleObjects(2&, h(0), 0&, INFINITE) = 1&
        If Not IsNTService Then
            CloseHandle hnd
            'MsgBox "This program must be started as service."
            MessageBox 0&, "This program must be started as a service.", App.Title, vbInformation Or vbOKOnly Or vbMsgBoxSetForeground
        End If
    Else
        MessageBox 0&, "This program is only for Windows NT/2000/XP.", App.Title, vbInformation Or vbOKOnly Or vbMsgBoxSetForeground
    End If
    
    If IsNTService Then
        ' ******************
        ' Here you may initialize and start service's objects
        ' These objects must be event-driven and must return control
        ' immediately after starting.
        ' ******************
        SetServiceState SERVICE_RUNNING
        App.LogEvent SERVICE_NAME & " started" '"VB6 Service Sample started"
        Set m_Service = New SunGardGPSService.Service
        Do
            ' ******************
            ' It is main service loop. Here you may place statements
            ' which perform useful functionality of this service.
            ' ******************
            ' Loop repeats every second. You may change this interval.
        Loop While WaitForSingleObject(hStopPendingEvent, 1000&) = WAIT_TIMEOUT
        ' ******************
        ' Here you may stop and destroy service's objects
        ' ******************
        Set m_Service = Nothing
        SetServiceState SERVICE_STOPPED
        App.LogEvent SERVICE_NAME & " stopped" '"VB6 Service Sample stopped"
        SetEvent hStopEvent
        ' Waiting for service thread termination
        WaitForSingleObject hnd, INFINITE
        CloseHandle hnd
    End If
    CloseHandle hStopEvent
    CloseHandle hStartEvent
    CloseHandle hStopPendingEvent
End Sub

Public Function CheckIsNT() As Boolean
    Dim OSVer As OSVERSIONINFO
    OSVer.dwOSVersionInfoSize = LenB(OSVer)
    GetVersionEx OSVer
    CheckIsNT = OSVer.dwPlatformId = VER_PLATFORM_WIN32_NT
End Function

