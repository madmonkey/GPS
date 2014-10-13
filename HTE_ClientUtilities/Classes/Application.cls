VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const cModuleName = "HTE_ClientUtilities.Application"

Implements HTE_GPS.Application
Private gProcesses As HTE_ClientUtilities.Processes
Private m_Status As HTE_GPS.GPS_HOST_STATUS
Private m_Notify As Collection 'subscribers to notification events
Public Type RouteConfig
    InMsg As HTE_GPS.GPS_MESSAGING_TYPES
    Path As String
End Type
Private bShowProcessed As Boolean

Private Sub Application_Cleanup()
On Error GoTo err_ApplicationCleanup
    Application_ShowLastMessage = False
    Application_Log cModuleName, "Application_Cleanup", "Clean-up of existing processes.", logInformation
    If Not gProcesses Is Nothing Then gProcesses.Unload
    Exit Sub
err_ApplicationCleanup:
UEH_LogError cModuleName, "Application_Cleanup", Err
End Sub

Private Property Get Application_HostStatus() As HTE_GPS.GPS_HOST_STATUS
    Application_HostStatus = m_Status
End Property

Private Function Application_IsVersionAtLeast(ByVal Version As String) As Boolean
    Application_IsVersionAtLeast = (StrComp(Version, App.Major & "." & App.Minor & "." & App.Revision, vbTextCompare) <> -1)
End Function

Private Sub Application_Log(ByVal Object As String, ByVal method As String, Optional ByVal Message As String = "", Optional ByVal LogDetail As HTE_GPS.GPS_LOG_DETAIL = 8&, Optional ByVal ErrorID As Long = 0&, Optional ByVal LogSource As String = "", Optional ByVal LogSourceDetail As HTE_GPS.GPS_LOG_SOURCE = 0&)
    UEH_Log Object, method, Message, LogDetail, ErrorID, LogSource, LogSourceDetail
End Sub

Private Sub Application_Processed(ByVal Message As HTE_GPS.GPSMessage, ByVal CurrentRoute As String)
Dim oStatus As HTE_GPS.AppStatus
Dim oObj As Variant
    If Application_ShowLastMessage Then
        If Not m_Notify Is Nothing Then
            Application_Log cModuleName, "Application_Processed", "Notifying subscriber(s) to processed message.", logInformation
            For Each oObj In m_Notify
                Set oStatus = oObj
                If Not oStatus Is Nothing Then
                    oStatus.Processed Message, CurrentRoute
                End If
            Next
        End If
    Else
        Application_Log cModuleName, "Application_Processed", "No notification taking place.", logInformation
    End If
End Sub

Private Function Application_ProcessStatus(ByVal CurrentRoute As String) As HTE_GPS.GPS_PROCESSOR_STATUS
    If Not gProcesses Is Nothing Then
        Application_ProcessStatus = gProcesses.ProcessStatus(CurrentRoute)
    End If
End Function

Private Function Application_PropertyPage(ByVal CurrentRoute As String) As HTE_GPS.PropertyPage
    If Not gProcesses Is Nothing Then
        Set Application_PropertyPage = gProcesses.PropertyPage(CurrentRoute)
    End If
End Function

Private Property Let Application_ShowLastMessage(RHS As Boolean)
    bShowProcessed = RHS
End Property

Private Property Get Application_ShowLastMessage() As Boolean
    'Try to limit out-of-process callbacks, only show messages if "asked"
    Application_ShowLastMessage = bShowProcessed
End Property

Private Property Get Application_Status() As HTE_GPS.GPS_HOST_STATUS
    Application_Status = m_Status
End Property

Private Sub Application_StatusCallback(Callback As HTE_GPS.AppStatus)
    Application_Log cModuleName, "Application_StatusCallback", "Adding subscriber to notification list.", logInformation
    m_Notify.Add Callback
    Application_Log cModuleName, "Application_StatusCallback", "Notifying subscriber to current status.", logInformation
    Callback.StatusChange m_Status
End Sub

Private Sub Application_StatusChange(statusCode As HTE_GPS.GPS_HOST_STATUS)
Dim oStatus As HTE_GPS.AppStatus
Dim oObj As Variant
    m_Status = statusCode
    If Not m_Notify Is Nothing Then
        Application_Log cModuleName, "Application_StatusChange", "Notifying subscriber(s) to changed status = " & hostStatusDesc(statusCode) & ".", logInformation
        For Each oObj In m_Notify
            Set oStatus = oObj
            oStatus.StatusChange m_Status
        Next
    End If
End Sub

Private Sub Class_Initialize()
Dim i As Long
    UEH_BeginLogging App.EXEName, App.EXEName
    If SoleInstance Then
        'This is solely if we wish to report errors in the future gracefully
        App.OleRequestPendingMsgTitle = "Global Positioning Processes"
        App.OleServerBusyMsgTitle = "Global Positioning Processes"
        App.OleServerBusyMsgText = "The Global Positioning Coordinate System is busy trying to complete a task." & vbCrLf & "If the situation persists, report the problem to your system administrator."
        App.OleRequestPendingTimeout = 60000
        App.OleServerBusyTimeout = 120000
        App.OleRequestPendingMsgText = "A Global Positioning Coordinate System Process is busy trying to complete a task." & vbCrLf & "If the situation persists, report the problem to your system administrator."
        'This should nullify any ActiveX message box error, and raise an error in the corresponding routine
        App.OleServerBusyRaiseError = True
        Application_Log cModuleName, "Class_Initialize", "Sole Instance, initializing processes.", logInformation
        Set gProcesses = New HTE_ClientUtilities.Processes
        Application_Log cModuleName, "Class_Initialize", "Processes loading....", logInformation
        gProcesses.Load Me
    Else
        Application_Log cModuleName, "Class_Initialize", "Instance already running...ignore and exit!", logWarning
    End If
    Set m_Notify = New Collection
End Sub

Private Sub Class_Terminate()
    Application_Log cModuleName, "Class_Terminate", "Unloaded processes.", logInformation
    While m_Notify.Count > 0: m_Notify.Remove 1: Wend
    Application_Log cModuleName, "Class_Terminate", "Removed notification subscribers.", logInformation
    Set m_Notify = Nothing
    Application_Log cModuleName, "Class_Terminate", "Clearing application mutex.", logInformation
    EndApp
    Application_Log cModuleName, "Class_Terminate", "Termination complete.", logInformation
    UEH_EndLogging
End Sub

