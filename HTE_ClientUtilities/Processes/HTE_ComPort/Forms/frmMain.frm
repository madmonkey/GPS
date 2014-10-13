VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   Caption         =   "ComPort WIndow"
   ClientHeight    =   1275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2565
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1275
   ScaleWidth      =   2565
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer statusTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   0
   End
   Begin SysInfoLib.SysInfo SystemInformation 
      Left            =   120
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer ReconnectTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1680
      Top             =   0
   End
   Begin VB.Timer ManualTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1200
      Top             =   0
   End
   Begin VB.Timer CommTime 
      Interval        =   1
      Left            =   720
      Top             =   0
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event ReceivedMessage(ByVal DataIn As String)
Public Event LogDetail(ByVal method As String, ByVal Message As String, ByVal LogDetail As HTE_GPS.GPS_LOG_DETAIL, ByVal ErrorID As Long, ByVal LogSource As String, ByVal LogSourceDetail As HTE_GPS.GPS_LOG_SOURCE)
Public Event Log(ByVal method As String, ByVal Message As String, ByVal LogDetail As HTE_GPS.GPS_LOG_DETAIL)
Public Event StatusChange(ByVal statusCode As HTE_GPS.GPS_PROCESSOR_STATUS)
Private m_GPSBuffer As stringBuilder 'String
Private m_PollingInfo As Collection
Private currentType As HTE_GPS.GPSConfiguration
'ADDED FOR SOFTWARE FLOWCONTROL - DEVICES THAT CANNOT BE CONFIGURED FOR TIME TO SEND (NMEA)
Private m_SecondsPerTransmission As Long 'send messages every x seconds
Private m_SecondsForTransmission As Long 'send messages for y seconds
Private m_Notify As ccrpTimers6.ccrpCountdown
Implements ccrpTimers6.ICcrpCountdownNotify
Private bTimeToWork As Boolean
Private dLastTransmission As Double
Private bUnloadingFromMemory As Boolean
Private m_PersistedSettings As CommSettings 'class
Private bReconnecting As Boolean 'is the timer thread actively trying to connect
Private bDeviceDisconnected As Boolean 'the system detected device is disconnected/we've verified it's ours
Private currStat As HTE_GPS.GPS_PROCESSOR_STATUS

'INFO: Troubleshooting Tips for the MSComm Control
    'http://support.microsoft.com/kb/192012/EN-US/
'PRB: Accessing WINMODEM with MSComm Control Can Hang Application
    'http://support.microsoft.com/kb/198128/
'Serial Port Device Connected Event?
    'http://forums.microsoft.com/MSDN/ShowPost.aspx?PostID=691074&SiteID=1
'MSCommControl
    'http://www.yes-tele.com/mscomm.html

Public Property Get ToString()
    With MSComm1
        ToString = "CommID [" & .CommID & _
            "] Break [" & CStr(.Break) & _
            "] CDHolding [" & CStr(.CDHolding) & _
            "] CTSHolding [" & CStr(.CTSHolding) & _
            "] DSRHolding [" & CStr(.DSRHolding) & _
            "] CommState [" & CommunicationState & _
            "] DTREnable [" & CStr(.DTREnable) & _
            "] RTSEnable [" & CStr(.RTSEnable) & "]"
    End With
End Property

Private Function CommunicationState() As String
    Select Case MSComm1.CommEvent
        Case comEvReceive: CommunicationState = "comEvReceive {Event}"
        Case comEvCTS: CommunicationState = "comEvCTS {Event}"
        Case comEventBreak: CommunicationState = "comEventBreak {Error}"
        Case comEventCDTO: CommunicationState = "comEventCDTO {Error}"
        Case comEventCTSTO: CommunicationState = "comEventCTSTO {Error}"
        Case comEventDSRTO: CommunicationState = "comEventDSRTO {Error}"
        Case comEventFrame: CommunicationState = "comEventFrame {Error}"
        Case comEventOverrun: CommunicationState = "comEventOverrun {Error}"
        Case comEventRxOver: CommunicationState = "comEventRxOver {Error}"
        Case comEventRxParity: CommunicationState = "comEventRxParity {Error}"
        Case comEventTxFull: CommunicationState = "comEventTxFull {Error}"
        Case comEventDCB: CommunicationState = "comEventDCB {Error}"
        Case comEvCD: CommunicationState = "comEvCD {Event}"
        Case comEvDSR: CommunicationState = "comEvDSR {Event}"
        Case comEvRing: CommunicationState = "comEvRing {Event}"
        Case comEvSend: CommunicationState = "comEvSend {Event}"
        Case comEvEOF: CommunicationState = "comEvEOF {Event}"
        Case Else: CommunicationState = "comUnknown {" & MSComm1.CommEvent & "}"
    End Select
End Function

Public Property Get IsPortOpen() As Boolean
    IsPortOpen = MSComm1.PortOpen And Not bDeviceDisconnected
    'RaiseEvent Log("IsPortOpen", "The answer is [" & CStr(MSComm1.PortOpen) & "]", GPS_LOG_VERBOSE)
End Property

Public Sub InitializeFromSettings(ByVal vSettings As CommSettings)
    Set m_PersistedSettings = vSettings
    If IsPortOpen Then MSComm1.PortOpen = False
    LoadFromSettings
    InitializeCommunications
End Sub

Private Sub LoadFromSettings()
    With MSComm1
        .CommPort = m_PersistedSettings.Port
        .Settings = m_PersistedSettings.Settings
        .RThreshold = m_PersistedSettings.RThreshold
        .InputLen = m_PersistedSettings.InputLen
        .InputMode = m_PersistedSettings.InputMode
        .DTREnable = m_PersistedSettings.DTREnable
        .EOFEnable = m_PersistedSettings.EOFEnable
        .Handshaking = m_PersistedSettings.Handshaking
        .InBufferSize = m_PersistedSettings.InBufferSize
        .NullDiscard = m_PersistedSettings.NullDiscard
        .RTSEnable = m_PersistedSettings.RTSEnable
        RaiseEvent Log("LoadFromSettings", m_PersistedSettings.ToString, GPS_LOG_INFORMATION)
    End With
End Sub

Public Function LimitMessages(ByVal SecondsPerTrans As Long, ByVal SecondsForTrans As Long)
    DisableCallback
    If SecondsPerTrans <> 0 And SecondsForTrans <> 0 Then
        m_SecondsPerTransmission = SecondsPerTrans
        m_SecondsForTransmission = SecondsForTrans
        bTimeToWork = True 'start polling right away!!!
        EnableCallback
    Else
        dLastTransmission = 0
        bTimeToWork = True
    End If
End Function

Private Function EnableCallback()
    If Not bUnloadingFromMemory Then
        Set m_Notify = New ccrpTimers6.ccrpCountdown
        With m_Notify
            .Duration = (m_SecondsForTransmission * 1000)
            .Interval = 100 'progress indicator
            Set .Notify = Me
            .Enabled = True
            dLastTransmission = Now
        End With
        RaiseEvent Log("EnableCallback", "Timer callback enabled", GPS_LOG_VERBOSE)
    End If
End Function

Private Function DisableCallback()
    If Not m_Notify Is Nothing Then
        m_Notify.Enabled = False
        Set m_Notify.Notify = Nothing
        Set m_Notify = Nothing
        RaiseEvent Log("DisableCallback", "Timer callback disabled", GPS_LOG_VERBOSE)
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    bUnloadingFromMemory = True
    DisableCallback
End Sub

Private Sub ICcrpCountdownNotify_Tick(ByVal TimeRemaining As Long)
    RaiseEvent Log("ICcrpCountdownNotify_Tick", "Tick - " & Format$(TimeRemaining / 1000, "0.0") & " sec till expiration.", GPS_LOG_VERBOSE)
End Sub

Private Sub ICcrpCountdownNotify_Timer()
    bTimeToWork = False
    dLastTransmission = Now
    DisableCallback
    RaiseEvent Log("ICcrpCountdownNotify_Timer", "Timer callback disabled", GPS_LOG_VERBOSE)
End Sub

Friend Function pollingBuffer(ByVal vData As String, ByVal vTime As String)
Dim sData() As String, sTime() As String
Dim x As Long
Dim this As Statement
Const cDefaultPoll = 30
    
    Set m_PollingInfo = New Collection
    If vData <> vbNullString Then
        sData = Split(vData, cSep)
        sTime = Split(vTime, cSep)
        For x = LBound(sData) To UBound(sData)
            Set this = New Statement
            With this
                .pollingBuffer = sData(x)
                .pollingInterval = cDefaultPoll
                If IsArray(sTime) Then
                    If x <= UBound(sTime) Then
                        If IsNumeric(sTime(x)) Then
                            .pollingInterval = CLng(sTime(x))
                        End If
                    End If
                End If
                m_PollingInfo.Add this
            End With
        Next
    End If
    ManualTime.Enabled = m_PollingInfo.Count > 0
    
End Function

Public Property Let bufferToSend(ByVal vData As String)
On Error GoTo err_bufferToSend
    RaiseEvent Log("bufferToSend", "CommState:" & ToString, GPS_LOG_INFORMATION)
    If MSComm1.CommID >= 0 Then
        If IsPortOpen Then
            'RaiseEvent Log("bufferToSend", "Sending data", GPS_LOG_INFORMATION)
            MSComm1.Output = currentType.SOM & vData & currentType.EOM
            RaiseEvent Log("bufferToSend", "Data sent.", GPS_LOG_INFORMATION)
            RaiseEvent LogDetail("bufferToSend", "Data sent.", GPS_LOG_VERBOSE, 0&, currentType.SOM & vData & currentType.EOM, GPS_SOURCE_BINARY)
        Else
            RaiseEvent Log("bufferToSend", "Port is not open data NOT sent.", GPS_LOG_ERROR)
            If Not ReconnectTimer.Enabled Then ReconnectTimer.Enabled = True
        End If
    Else
        RaiseEvent Log("bufferToSend", "Assigned device is unavailable!", GPS_LOG_ERROR)
    End If
    Exit Property
err_bufferToSend:
    RaiseEvent Log("bufferToSend", "[" & Err.Number & "] " & Err.Description, GPS_LOG_ERROR)
    Err.Clear
End Property

Public Property Get MessageType() As HTE_GPS.GPSConfiguration
    MessageType = currentType
End Property

Public Property Let MessageType(ByRef GPSType As HTE_GPS.GPSConfiguration)
    RaiseEvent Log("MessageType", "Assigned as " & GPSType.Desc, GPS_LOG_VERBOSE)
    currentType = GPSType
End Property

Private Sub CommTime_Timer()
'Dim bufferLenIn As Long
Dim bufferToSend As String
   
On Error GoTo err_CommTime
    If m_Notify Is Nothing Then
        If Not bTimeToWork Then
            RaiseEvent Log("CommTime_Timer", "Last transmission was " & CStr(CDate(dLastTransmission)) & "; " & DateDiff("s", dLastTransmission, Now) & " seconds ago; " & m_SecondsPerTransmission & " seconds for transactions.", GPS_LOG_VERBOSE)
            bTimeToWork = DateDiff("s", CDate(dLastTransmission), Now) >= (m_SecondsPerTransmission)
            If bTimeToWork Then EnableCallback
        End If
    End If

    If bTimeToWork Then
        'is it garbage day?
        If m_PersistedSettings.MaxCachedBytes < m_GPSBuffer.Length Then
            'yup
            RaiseEvent Log("CommTime_Timer", "Clearing buffer (bytes accumulated)[" & m_GPSBuffer.Length & "] > [" & m_PersistedSettings.MaxCachedBytes & "] bytes permitted.", GPS_LOG_INFORMATION)
            RaiseEvent LogDetail("CommTime_Timer", "Clearing buffer (data purged)", GPS_LOG_VERBOSE, 0, m_GPSBuffer.ToString, GPS_SOURCE_BINARY)
            m_GPSBuffer.Remove 0, m_GPSBuffer.Length - 1
        End If
        'bufferLenIn = MSComm1.InBufferCount + m_GPSBuffer.Length  'Len(m_GPSBuffer)
        'don't directly assign to long!
        If MSComm1.InBufferCount + m_GPSBuffer.Length > 0 Then
            m_GPSBuffer.Append MSComm1.Input
            bufferToSend = ParseGPSMessage
            If bufferToSend <> vbNullString Then
                If ValidGPSMessage(bufferToSend) Then
                    'Some units send so many different message types, don't flood airways with messages that will not be used
                    RaiseEvent ReceivedMessage(bufferToSend)
                    currentStatus GPS_STAT_READYANDWILLING
                End If
            End If
        End If
        CommTime.Enabled = (m_GPSBuffer.Find(currentType.SOM) > 0) And (m_GPSBuffer.Find(currentType.EOM) > 0) _
            And MSComm1.CommID >= 0 And Not bDeviceDisconnected 'InStr(1, m_GPSBuffer, currentType.SOM, vbTextCompare) > 0 And InStr(1, m_GPSBuffer, currentType.EOM, vbTextCompare) > 0 'False
    Else
        'I really, really don't like this, but since multiple transmissions will effectively "kill" the switch/network this
        'stop-gate is provided to limit the number of transmissions to the switch (specifically for NMEA devices)....
        'We just empty the buffer here and log that we've done so!!
        If MSComm1.InBufferCount > 0 Then
            'Only log IF we are actually dropping something!!
            bufferToSend = MSComm1.Input
            'Dropping message - but DON'T change status! Issue
            RaiseEvent LogDetail("CommTime_Timer", "Dropping message from delivery!", GPS_LOG_VERBOSE, 0, bufferToSend, GPS_SOURCE_BINARY)
'''            RaiseEvent StatusChange(GPS_STAT_READYANDWILLING) 'More consistant with RAWSOCKET Process
        End If
    End If
        Exit Sub
err_CommTime:
    RaiseEvent LogDetail("CommTime_Timer", Err.Description, GPS_LOG_ERROR, Err.Number, Err.Source, GPS_SOURCE_STRING)
    CommTime.Enabled = False
    If Not ReconnectTimer.Enabled Then ReconnectTimer.Enabled = True
End Sub

Private Function ParseGPSMessage() As String

Dim startPos As Long, endPos As Long

On Error GoTo err_ParseGPSMessage
    
    If Not m_GPSBuffer Is Nothing Then
        With m_GPSBuffer
            If .Length > 0 Then
                startPos = .Find(currentType.SOM)
                If startPos > 0 Then 'And endPos > startPos Then
                    If startPos > 1 Then
                        .Remove 0, startPos - 1
                        startPos = 1
                    End If
                    endPos = .Find(currentType.EOM, startPos + 1) 'Do we have an end tag?
                    If endPos > startPos Then
                        ParseGPSMessage = Mid$(.ToString, startPos + Len(currentType.SOM), (endPos - (startPos + Len(currentType.SOM))))    'Mid$(.ToString, startPos + Len(GPSConfig.SOM), (endPos - (startPos + Len(GPSConfig.SOM))))
                        .Remove startPos - 1, endPos
                    End If
                End If
            End If
        End With
    End If
    Exit Function
    
err_ParseGPSMessage:
    RaiseEvent Log("ParseGPSMessage", Err.Number & ":" & Err.Description, GPS_LOG_ERROR)
    Err.Clear
End Function

Private Function ValidGPSMessage(ByVal bufferToSend As String) As Boolean
'Left as a hook if needed, can configure from initstrings
    Select Case currentType.GPSType
        Case GPS_TYPE_0, GPS_TYPE_1, GPS_TYPE_2, GPS_TYPE_3
            ValidGPSMessage = True
        Case Else
            ValidGPSMessage = False
    End Select
End Function

Private Sub Form_Load()
    Set m_GPSBuffer = New stringBuilder
    bUnloadingFromMemory = False
    
    currentType.GPSType = GPS_TYPE_0
    currentType.EOM = "<"
    currentType.SOM = ">"
    currentType.Desc = "TAIP"
End Sub

Private Sub ManualTime_Timer()
'This function was retrofitted later to allow a programmatic way to request
'information from the device at somewhat regularly scheduled intervals, in case
'the reporting mode is "broken" in GPS device
Dim oObj As Statement
    If m_PollingInfo.Count > 0 Then
        For Each oObj In m_PollingInfo
            With oObj
                If DateDiff("s", Now, .NextAlarm) < 1 Then
                    RaiseEvent Log("ManualTime_Timer", "Sending timed manual message.", GPS_LOG_VERBOSE)
                    bufferToSend = .pollingBuffer
                    .NextAlarm = DateAdd("s", .pollingInterval, Now)
                    RaiseEvent Log("ManualTime_Timer", "Next message will be sent " & Format$(.NextAlarm, "hh:nn:ss") & ".", GPS_LOG_VERBOSE)
                End If
            End With
        Next
    Else
        ManualTime.Enabled = False
    End If
End Sub

Private Sub MSComm1_OnComm()
Dim bReinitialize As Boolean
On Error GoTo err_OnComm
    bReinitialize = False
    Select Case MSComm1.CommEvent
        Case comEvReceive
            If Not CommTime.Enabled Then CommTime.Enabled = MSComm1.CommID >= 0 And Not bDeviceDisconnected
            'ReportCommState "[comEvReceive] Data Recieved", GPS_LOG_VERBOSE
        Case comEvCTS
            If Not CommTime.Enabled And MSComm1.CTSHolding Then CommTime.Enabled = MSComm1.CTSHolding And MSComm1.CommID >= 0 And Not bDeviceDisconnected
            ReportCommState "[comEvCTS] Change in the CTS line.", GPS_LOG_VERBOSE
        'Errors - close/reopen???
        Case comEventBreak
            ReportCommState "[comEventBreak] A Break Was Recieved.", GPS_LOG_ERROR
            bReinitialize = True
        Case comEventCDTO
            ReportCommState "[comEventCDTO] CD (RLSD) Timeout.", GPS_LOG_ERROR
            bReinitialize = True
        Case comEventCTSTO
            ReportCommState "[comEventCTSTO] CTS Timeout.", GPS_LOG_ERROR
            bReinitialize = True
        Case comEventDSRTO   ' DSR Timeout.
            ReportCommState "[comEventDSRTO] DSR Timeout.", GPS_LOG_ERROR
            bReinitialize = True
        Case comEventFrame   ' Framing Error.
            ReportCommState "[comEventFrame] Framing Error.", GPS_LOG_ERROR
            bReinitialize = True
        Case comEventOverrun ' Data Lost.
            ReportCommState "[comEventOverrun] Data Lost.", GPS_LOG_ERROR
            bReinitialize = True
        Case comEventRxOver  ' Receive buffer overflow.
            ReportCommState "[comEventRxOver] Receive buffer overflow.", GPS_LOG_ERROR
            bReinitialize = True
        Case comEventRxParity   ' Parity Error.
            ReportCommState "[comEventRxParity] Parity Error.", GPS_LOG_ERROR
            bReinitialize = True
        Case comEventTxFull  ' Transmit buffer full.
            ReportCommState "[comEventTxFull] Transmit buffer full.", GPS_LOG_ERROR
            RaiseEvent Log("OnComm", "Clearing stale cached data.", GPS_LOG_INFORMATION)
            MSComm1.OutBufferCount = 0
            'bReinitialize = True
        Case comEventDCB     ' Unexpected error retrieving DCB]
            ReportCommState "[comEventDCB] Unexpected error retrieving DCB.", GPS_LOG_ERROR
            bReinitialize = True
         ' Events
        Case comEvCD   ' Change in the CD line.
            ReportCommState "[comEvCD] Change in the CD line.", GPS_LOG_VERBOSE
        Case comEvDSR  ' Change in the DSR line.
            ReportCommState "[comEvDSR] Change in the DSR line.", GPS_LOG_VERBOSE
        Case comEvRing ' Change in the Ring Indicator.
            ReportCommState "[comEvRing] Change in the Ring Indicator.", GPS_LOG_VERBOSE
        Case comEvSend ' There are SThreshold number of
            ReportCommState "[comEvSend] There are [" & MSComm1.SThreshold & "] byte(s) in transmit buffer.", GPS_LOG_VERBOSE
        Case comEvEOF  ' An EOF character was found in the input stream.
            ReportCommState "[comEvEOF] An EOF character was found in the input stream.", GPS_LOG_VERBOSE
    End Select
    Exit Sub
err_OnComm:
    RaiseEvent Log("OnComm", Err.Number & " - " & Err.Description, GPS_LOG_ERROR)
    Err.Clear
    If bReinitialize Then
        RaiseEvent Log("OnComm", "Attempting reinitialization of device", GPS_LOG_INFORMATION)
        On Local Error Resume Next
        MSComm1.PortOpen = False
        ReconnectTimer.Enabled = bReinitialize
    End If
'    If bReinitialize Then
'        If IsPortOpen Then MSComm1.PortOpen = False
'        ReconnectTimer_Timer
'    End If
End Sub
Private Sub ReportCommState(ByVal Description As String, ByVal Level As HTE_GPS.GPS_LOG_DETAIL)
    RaiseEvent Log("OnCommunicationsEvent", Description, Level)
End Sub
Private Sub ReconnectTimer_Timer()
On Error GoTo err_ReconnectTimer
    If m_PersistedSettings.IsAssignable Then
        If Not bReconnecting And Not bDeviceDisconnected Then
            RaiseEvent Log("ConnectionTimer", "bReconnecting = " & CStr(bReconnecting) & "; bDeviceDisconnected = " & CStr(bDeviceDisconnected), GPS_LOG_VERBOSE)
            bReconnecting = True
            If Not bUnloadingFromMemory Then
                If Not MSComm1.PortOpen Then
                    LoadFromSettings 'refresh any settings that may have been lost from close!
                    RaiseEvent Log("ConnectionTimer", "Opening port...", GPS_LOG_INFORMATION)
                    MSComm1.PortOpen = True
                    With ReconnectTimer
                        .Enabled = Not IsPortOpen
                        .Interval = Random(200&, 750&) 'takes somewhere in this range(ms) for hardware to initialize
                        If IsPortOpen Then
                            TransmitInitializationCommands
                        End If
                    End With
                Else
                    MSComm1.PortOpen = False
                    RaiseEvent Log("ConnectionTimer", "Closing port...", GPS_LOG_INFORMATION)
                    ReconnectTimer.Enabled = Not IsPortOpen
                End If
            Else
                RaiseEvent Log("ConnectionTimer", "Currently unloading from host -> shutting down.", GPS_LOG_VERBOSE)
                ReconnectTimer.Enabled = False
            End If
            bReconnecting = False
        End If
    Else
        RaiseEvent Log("ConnectionTimer", "Current assignment is port [" & m_PersistedSettings.Port & "] currently available port(s) [" & m_PersistedSettings.AvailablePorts & "]", GPS_LOG_INFORMATION)
        ReconnectTimer.Enabled = False
    End If
    Exit Sub
err_ReconnectTimer:
    RaiseEvent LogDetail("ReconnectTimer_Timer", "CommState: " & ToString, GPS_LOG_ERROR, Err.Number, Err.Description, GPS_SOURCE_STRING)
    bReconnecting = False
    Err.Clear
End Sub

Public Sub InitializeCommunications()

    If Not m_PersistedSettings Is Nothing Then
        If m_PersistedSettings.IsAssignable Then
            ReconnectTimer_Timer
            If Not IsPortOpen Then
                With ReconnectTimer
                    .Interval = Random(200&, 750&) 'takes somewhere in this range(ms) for hardware to initialize
                    .Enabled = True
                End With
            Else
                TransmitInitializationCommands
            End If
        Else
            RaiseEvent Log("InitializeCommunications", "Invalid port assigned (" & m_PersistedSettings.Port & ") - valid ports include(" & m_PersistedSettings.AvailablePorts & ")", GPS_LOG_ERROR)
        End If
    Else
        RaiseEvent Log("InitializeCommunications", "Persisted object settings are invalid!", GPS_LOG_ERROR)
    End If
End Sub
Private Sub TransmitInitializationCommands()
Dim i As Long
    'when we reopen port some devices don't cache settings - this is to reinitialize
    'RaiseEvent StatusChange(GPS_STAT_READYANDWILLING)
    currentStatus GPS_STAT_READYANDWILLING
    With m_PersistedSettings
        For i = LBound(.InitializationValues) To UBound(.InitializationValues)
            If IsPortOpen And MSComm1.CommID >= 0 Then
                RaiseEvent LogDetail("TransmitInitializationCommands", "Sending initialization cmd[" & CStr(i) & "]", GPS_LOG_INFORMATION, 0&, CStr(.InitializationValues(i)), GPS_SOURCE_BINARY)
                bufferToSend = CStr(.InitializationValues(i))
            Else
                RaiseEvent Log("TransmitInitializationCommands", "Device is unusable at this point - deferring processing commands!", GPS_LOG_ERROR)
            End If
        Next
        CommTime.Enabled = True
    End With
End Sub
Private Function Random(Lowerbound As Long, Upperbound As Long)
    Randomize Now
    Random = Int(Rnd * Upperbound) + Lowerbound
End Function

Private Sub statusTimer_Timer()
    statusTimer.Enabled = False
    RaiseEvent StatusChange(currStat)
End Sub

Private Sub SystemInformation_ConfigChangeCancelled()
    RaiseEvent Log("ConfigChangeCancelled", "Detected", GPS_LOG_INFORMATION)
End Sub

Private Sub SystemInformation_ConfigChanged(ByVal OldConfigNum As Long, ByVal NewConfigNum As Long)
    RaiseEvent Log("ConfigChanged", "OldConfigNum [" & OldConfigNum & "]; NewConfigNum [" & NewConfigNum & "]", GPS_LOG_INFORMATION)
End Sub

Private Sub SystemInformation_DeviceArrival(ByVal DeviceType As Long, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long)
Dim i As Long, vInit As Variant
    RaiseEvent Log("DeviceArrival", "DeviceType [" & DeviceType & "]; DeviceID [" & DeviceID & "]; DeviceName [" & DeviceName & "]; DeviceData [" & DeviceData & "]", GPS_LOG_INFORMATION)
    Select Case DeviceType
        Case 3 'Serial/Parallel port (we should be concerned).
            'is it ours? No other value passed have to tweak from devName
            If StrComp(DeviceName, "COM" & MSComm1.CommPort, vbTextCompare) = 0 Then
                bDeviceDisconnected = False
                currentStatus GPS_STAT_WARNING
                If Not IsPortOpen Then
                    ReconnectTimer.Enabled = True
                    RaiseEvent Log("DeviceArrival", "[COM" & MSComm1.CommPort & "] was signalled to activate.", GPS_LOG_INFORMATION)
                Else
                    RaiseEvent Log("DeviceArrival", "[COM" & MSComm1.CommPort & "] is reportedly already active.", GPS_LOG_INFORMATION)
                End If
            Else
                RaiseEvent Log("DeviceArrival", DeviceName & " added - currently using [COM" & MSComm1.CommPort & "]. No action taken.", GPS_LOG_INFORMATION)
            End If
        Case Else
            RaiseEvent Log("DeviceArrival", "Detected " & GetDeviceTypeDescription(DeviceType) & " added. No action taken.", GPS_LOG_INFORMATION)
    End Select
End Sub

Private Sub SystemInformation_DeviceOtherEvent(ByVal DeviceType As Long, ByVal EventName As String, ByVal DataPointer As Long)
    RaiseEvent Log("DeviceOtherEvent", "DeviceType [" & DeviceType & "]; EventName [" & EventName & "] DataPointer [" & DataPointer & "]", GPS_LOG_INFORMATION)
End Sub

Private Sub SystemInformation_DeviceQueryRemove(ByVal DeviceType As Long, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long, Cancel As Boolean)
    'I can cancel this if I want to! Yeah right...never gets triggerd!
    RaiseEvent Log("DeviceQueryRemove", "DeviceType [" & DeviceType & "]; DeviceID [" & DeviceID & "]; DeviceName [" & DeviceName & "]; DeviceData [" & DeviceData & "]", GPS_LOG_INFORMATION)
End Sub

Private Sub SystemInformation_DeviceQueryRemoveFailed(ByVal DeviceType As Long, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long)
    RaiseEvent Log("DeviceQueryRemoveFailed", "DeviceType [" & DeviceType & "]; DeviceID [" & DeviceID & "]; DeviceName [" & DeviceName & "]; DeviceData [" & DeviceData & "]", GPS_LOG_INFORMATION)
End Sub

Private Sub SystemInformation_DeviceRemoveComplete(ByVal DeviceType As Long, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long)
On Error GoTo err_SystemInformation_DeviceRemoveComplete
    RaiseEvent Log("DeviceRemoveComplete", "DeviceType [" & DeviceType & "]; DeviceID [" & DeviceID & "]; DeviceName [" & DeviceName & "]; DeviceData [" & DeviceData & "]", GPS_LOG_INFORMATION)
    'since the device is gone if I'm using it -  I should close the port
    Select Case DeviceType
        Case 3 'Serial/Parallel port (we should be concerned).
            'is it ours? No other value passed have to tweak from devName
            If StrComp(DeviceName, "COM" & MSComm1.CommPort, vbTextCompare) = 0 Then
                'RaiseEvent StatusChange(GPS_STAT_WARNING) 'put this further up the stack event though hasn't happened yet.
                If IsPortOpen Then
                    bDeviceDisconnected = True
                    MSComm1.PortOpen = False 'we use actual property - the IsPortOpen generally used for transmission
                    RaiseEvent Log("DeviceRemoveComplete", "[COM" & MSComm1.CommPort & "] was signalled to close.", GPS_LOG_INFORMATION)
                Else
                    bDeviceDisconnected = True
                    RaiseEvent Log("DeviceRemoveComplete", "[COM" & MSComm1.CommPort & "] is already reportedly closed.", GPS_LOG_INFORMATION)
                End If
                'RaiseEvent StatusChange(GPS_STAT_ERROR) 'show disconnected
                currentStatus GPS_STAT_ERROR 'marshalling between async of app to sync of comm is causing a problem
            Else
                RaiseEvent Log("DeviceRemoveComplete", DeviceName & " removed - currently using [COM" & MSComm1.CommPort & "]. No action taken.", GPS_LOG_INFORMATION)
            End If
        Case Else
            RaiseEvent Log("DeviceRemoveComplete", "Detected " & GetDeviceTypeDescription(DeviceType) & " removed. No action taken.", GPS_LOG_INFORMATION)
    End Select
    Exit Sub
err_SystemInformation_DeviceRemoveComplete:
    RaiseEvent Log("DeviceRemoveComplete", "[" & Err.Number & "] " & Err.Description, GPS_LOG_ERROR)
End Sub

Private Sub SystemInformation_DeviceRemovePending(ByVal DeviceType As Long, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long)
    RaiseEvent Log("DeviceRemovePending", "DeviceType [" & DeviceType & "]; DeviceID [" & DeviceID & "]; DeviceName [" & DeviceName & "]; DeviceData [" & DeviceData & "]", GPS_LOG_INFORMATION)
End Sub

Private Function GetDeviceTypeDescription(ByVal devType As Long) As String
    Select Case devType
        Case 0: GetDeviceTypeDescription = "{OEM-defined device}"
        'A device node refers to a device that can host other hardware, such as a SCSCI controller.
        Case 1: GetDeviceTypeDescription = "{Device node}"
        Case 2: GetDeviceTypeDescription = "{Logical volume (disk drive)}"
        Case 3: GetDeviceTypeDescription = "{Serial or parallel port}"
        Case Else: GetDeviceTypeDescription = "{Unsupported}"
    End Select
End Function
Private Sub SystemInformation_DevModeChanged()
    RaiseEvent Log("DevModeChanged", "Detected", GPS_LOG_INFORMATION)
End Sub

Private Sub currentStatus(ByVal currentStatCode As HTE_GPS.GPS_PROCESSOR_STATUS)
    currStat = currentStatCode
    statusTimer.Enabled = True
End Sub

Private Sub SystemInformation_DisplayChanged()
    RaiseEvent Log("DisplayChanged", "System display changed!", GPS_LOG_VERBOSE)
End Sub

Private Sub SystemInformation_PowerQuerySuspend(Cancel As Boolean)
    RaiseEvent Log("PowerQuerySuspend", "PowerQuerySuspend event Cancel handle:[" & CStr(Cancel) & "]", GPS_LOG_VERBOSE)
End Sub

Private Sub SystemInformation_PowerResume()
    RaiseEvent Log("PowerResume", "System notification of power resume", GPS_LOG_VERBOSE)
End Sub

Private Sub SystemInformation_PowerStatusChanged()
    RaiseEvent Log("PowerStatusChanged", "System notification of power status changed", GPS_LOG_VERBOSE)
End Sub

Private Sub SystemInformation_PowerSuspend()
    RaiseEvent Log("PowerSuspend", "System notification of power suspend", GPS_LOG_VERBOSE)
End Sub

Private Sub SystemInformation_QueryChangeConfig(Cancel As Boolean)
    RaiseEvent Log("QueryChangeConfig", "QueryChangeConfig event Cancel handle:[" & CStr(Cancel) & "]", GPS_LOG_VERBOSE)
End Sub

Private Sub SystemInformation_SettingChanged(ByVal Item As Integer)
    RaiseEvent Log("SettingChanged", "SettingChanged event event Item:[" & CStr(Item) & "]", GPS_LOG_VERBOSE)
End Sub

Private Sub SystemInformation_SysColorsChanged()
    RaiseEvent Log("SysColorsChanged", "System color changed event", GPS_LOG_VERBOSE)
End Sub

Private Sub SystemInformation_TimeChanged()
    RaiseEvent Log("TimeChanged", "System detected time change event", GPS_LOG_VERBOSE)
End Sub
