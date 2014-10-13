VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const cModuleName = "HTE_RawSocket.Process"
Private WithEvents fMain As frmMain
Attribute fMain.VB_VarHelpID = -1
Private gApp As HTE_GPS.Processes
Private m_Status As HTE_GPS.GPS_PROCESSOR_STATUS
Private localSettings As MSXML2.DOMDocument30
Private m_Settings As String
Implements HTE_GPS.Process
Implements HTE_GPS.PropPageCallback
Private m_Property As Object
Private Enum VersionRequired
    SOCKET_VERSION_11 = &H101
    SOCKET_VERSION_22 = &H202
End Enum
Private Const WSA_DescriptionSize = 257
Private Const WSA_SysStatusSize = 257
Private Type WSADataType
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DescriptionSize
    szSystemStatus As String * WSA_SysStatusSize
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type
Private WSAStartedUp As Boolean
Private bProcessOnSend As Boolean
Private Declare Function WSACancelBlockingCall Lib "wsock32.dll" () As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Function WSAIsBlocking Lib "wsock32.dll" () As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
Private currentTypes() As HTE_GPS.GPSConfiguration
'ADDED FOR EVER-CHANGING REQUIREMENTS (ALIASING) NOW - WHY IS LACK OF REQUIREMENTS ON THEIR POINT HEARTACHE FROM OUR POINT???
Private WithEvents AliasLookup As HTE_Entity.Alias 'resolve entity (ip/mac) class
Attribute AliasLookup.VB_VarHelpID = -1
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'added for slower machines where o/s has not released all resources!

Private Sub AliasLookup_LookupFailed(cMsg As HTE_GPS.GPSMessage)
    Log cModuleName, "AliasLookup_LookupFailed", "Unable to resolve entity according to settings.", GPS_LOG_VERBOSE, , cMsg.Entity, GPS_SOURCE_XML
    If Not gApp Is Nothing Then
        cMsg.MessageStatus = GPS_MSG_PROCESSED_WARNING
        gApp.ProcessComplete cMsg, Me
    End If
End Sub

Private Sub AliasLookup_LookupResolved(cMsg As HTE_GPS.GPSMessage)
    Log cModuleName, "AliasLookup_LookupResolved", "Resolved entity according to settings.", GPS_LOG_VERBOSE, , cMsg.Entity, GPS_SOURCE_XML
    If Not gApp Is Nothing Then gApp.ProcessComplete cMsg, Me
End Sub

Private Sub Class_Initialize()
    m_Status = GPS_STAT_INITIALIZED
    Set localSettings = New MSXML2.DOMDocument30
    StartWinsock App.EXEName ' just in case in same route as server socket...ensure that winsock is started!
    Set AliasLookup = New HTE_Entity.Alias
End Sub

Private Sub Class_Terminate()
    m_Status = GPS_STAT_UNINITIALIZED
    Set localSettings = Nothing
    EndWinsock
    Set AliasLookup = Nothing
End Sub

Private Sub fMain_Log(ByVal method As String, ByVal Message As String, ByVal LogDetail As HTE_GPS.GPS_LOG_DETAIL)
    Log cModuleName, method, Message, LogDetail
End Sub

Private Sub fMain_LogDetail(ByVal method As String, ByVal Message As String, ByVal LogDetail As HTE_GPS.GPS_LOG_DETAIL, ByVal ErrorID As Long, ByVal LogSource As String, ByVal LogSourceDetail As HTE_GPS.GPS_LOG_SOURCE)
    Log cModuleName, method, Message, LogDetail, ErrorID, LogSource, LogSourceDetail
End Sub

'''Private Sub fMain_ReceivedMessage(ByVal DataIn As String)
'''Dim msg As Message
'''On Error GoTo err_ReceivedMessage
'''    Log cModuleName, "fMain_ReceivedMessage", "Received message.", GPS_LOG_INFORMATION
'''    Log cModuleName, "fMain_ReceivedMessage", "Received message from listener.", GPS_LOG_VERBOSE, , DataIn, GPS_SOURCE_BINARY
'''    Set msg = New Message
'''    msg.MessageStatus = GPS_MSG_PROCESSED
'''    msg.MessageType = fMain.MessageType.GPSType
'''    msg.rawMessage = DataIn
'''    If Not gApp Is Nothing Then gApp.ProcessComplete msg, Me
'''    Exit Sub
'''err_ReceivedMessage:
'''    Err.Clear
'''End Sub

Private Sub fMain_ReceivedMessageFrom(ByVal DataIn As String, ByVal EndPoint As String)
Dim msg As Message
On Error GoTo err_ReceivedMessageFrom
    Log cModuleName, "fMain_ReceivedMessageFrom", "Received message from [" & EndPoint & "].", GPS_LOG_INFORMATION
    Log cModuleName, "fMain_ReceivedMessageFrom", "Received message from [" & EndPoint & "].", GPS_LOG_VERBOSE, , DataIn, GPS_SOURCE_BINARY
    Set msg = New Message
    msg.MessageStatus = GPS_MSG_PROCESSED
    msg.MessageType = fMain.MessageType.GPSType
    msg.rawMessage = DataIn
    If Not AliasLookup Is Nothing Then
        With AliasLookup
            .resolveEntity msg, EndPoint
        End With
    End If
    '''If Not gApp Is Nothing Then gApp.ProcessComplete msg, Me
    Exit Sub
err_ReceivedMessageFrom:
    Err.Clear
End Sub

Private Sub fMain_StatusChange(ByVal statusCode As HTE_GPS.GPS_PROCESSOR_STATUS)
    Process_Status = statusCode
End Sub

Private Sub Process_Connect(ByVal parentProcess As HTE_GPS.Processes, ByVal cookie As String)
On Error GoTo err_Connect
    If Not parentProcess Is Nothing Then
        Set gApp = parentProcess
        gApp.Connected Me, cookie
    End If
    Exit Sub
err_Connect:
    Err.Clear
End Sub

Private Sub Process_Disconnect()
    Log cModuleName, "Process_Disconnect", "Unhooking process", GPS_LOG_VERBOSE
    UnloadSafely
    checkForChanges
    Set m_Property = Nothing
On Error GoTo err_Disconnect
    Set gApp = Nothing
    Exit Sub
err_Disconnect:
    Err.Clear
End Sub

Private Sub UnloadSafely()
    If Not fMain Is Nothing Then
        If fMain.Winsock1.State = sckOpen And fMain.Winsock1.Protocol = sckTCPProtocol Then
            fMain.Winsock1.Close
        End If
        Unload fMain
        Set fMain = Nothing
    End If
End Sub

Private Property Get Process_FriendlyName() As String
    Process_FriendlyName = "Client socket connection"
End Property

Private Property Get Process_HostVersion() As String
    Process_HostVersion = "1.0.0"
End Property

Private Sub Process_ProcessMessage(ByVal PropData As HTE_GPS.GPSMessage)
Dim i As Long
Dim bFound As Boolean
    
    Log cModuleName, "Process_ProcessMessage", "Buffer arrived!", GPS_LOG_VERBOSE, , PropData.rawMessage, GPS_SOURCE_BINARY
    If Not fMain Is Nothing Then
        If fMain.ValidateMessage Then
            bFound = False
            'add appropriate start/end tags according to processed message
            For i = LBound(currentTypes) To UBound(currentTypes)
                If currentTypes(i).GPSType = PropData.Type Then
                    fMain.bufferToSend currentTypes(i).SOM & PropData.rawMessage & currentTypes(i).EOM, True
                    bFound = True
                    Exit For
                End If
            Next
            If Not bFound Then fMain.bufferToSend PropData.rawMessage, True 'in case NOT found
            'need to signal winsock to not "double validate"
        Else
            'just send raw
            fMain.bufferToSend PropData.rawMessage, True
        End If
        If Err.Number = 0 And Err.LastDllError = 0 Then
            PropData.MessageStatus = GPS_MSG_PROCESSED
            If bProcessOnSend Then
On Error Resume Next
                'the application object may be shutting down - we can try and guard against by testing
                'but on rare occasions you can "time" it just right...
                If Not gApp Is Nothing Then gApp.ProcessComplete PropData, Me
            End If
        Else
            PropData.MessageStatus = GPS_MSG_PROCESSED_ERROR
        End If
    Else
        PropData.MessageStatus = GPS_MSG_ERROR
    End If
    Exit Sub
err_ProcessMessage:
    Err.Clear
End Sub

Private Property Get Process_PropertyPage() As HTE_GPS.PropertyPage
Dim oObj As HTE_GPS.PropertyPage
On Error GoTo err_Process_PropertyPage
    Set m_Property = CreateObject(cHelperPage)
    Set oObj = m_Property
    oObj.Settings = m_Settings
    oObj.PropertyCallback = Me
    Set Process_PropertyPage = oObj
    Exit Property
err_Process_PropertyPage:
    Log cModuleName, "Process_PropertyPage", Err.Description, GPS_LOG_WARNING, Err.Number
End Property

Private Property Let Process_Settings(ByVal RHS As String)
    If Not loadLocalSettings(RHS) Then
        Log cModuleName, "Process_Settings", "Unable to load Instance Settings", GPS_LOG_ERROR, Err.Number, RHS, GPS_SOURCE_XML
    Else
        setProperties
    End If
    m_Settings = RHS
End Property

Private Property Get Process_Settings() As String
    Process_Settings = retrieveLocalSettings
End Property

Private Property Let Process_Status(RHS As HTE_GPS.GPS_PROCESSOR_STATUS)
    If m_Status <> RHS Then
        m_Status = RHS
On Error Resume Next
        'the application object may be shutting down - we can try and guard against by testing
        'but on rare occasions you can "time" it just right...
        If Not gApp Is Nothing Then gApp.StatusChanged m_Status, Me
    End If
    Exit Property
err_Status:
    Err.Clear
End Property

Private Property Get Process_Status() As HTE_GPS.GPS_PROCESSOR_STATUS
    Process_Status = m_Status
End Property

Private Function checkForChanges() As Boolean
Dim bBadNode As Boolean
On Error GoTo err_checkForChanges
    If Not gApp Is Nothing Then
        If modProperties.PropertiesChanged(m_Settings, localSettings, bBadNode) Then
            If bBadNode Then
                Log cModuleName, "checkForChanges", "Unable to verify against initial settings.", GPS_LOG_WARNING, , m_Settings, GPS_SOURCE_XML
            Else
                checkForChanges = True
                If Not gApp Is Nothing Then gApp.PropertiesChanged retrieveLocalSettings, Me
            End If
        End If
    End If
    Exit Function
err_checkForChanges:
    Err.Clear
End Function

Private Function loadLocalSettings(ByVal sXML As String, Optional bBaseline As Boolean = True) As Boolean
    If Not localSettings Is Nothing Then Set localSettings = Nothing
    Set localSettings = New MSXML2.DOMDocument30
    If bBaseline Then m_Settings = sXML
    loadLocalSettings = localSettings.loadXML(sXML)
End Function

Private Function retrieveLocalSettings() As String
    
    If Not localSettings Is Nothing Then
        retrieveLocalSettings = localSettings.xml
    Else
        retrieveLocalSettings = m_Settings
    End If
End Function

Private Sub Log(ByVal Object As String, ByVal method As String, Optional ByVal Message As String = "", Optional ByVal LogDetail As HTE_GPS.GPS_LOG_DETAIL = 8&, Optional ByVal ErrorID As Long = 0&, Optional ByVal LogSource As String = "", Optional ByVal LogSourceDetail As HTE_GPS.GPS_LOG_SOURCE = 0&)
Dim currentStat As HTE_GPS.GPS_PROCESSOR_STATUS
On Error GoTo err_Log
    If Not gApp Is Nothing Then
        gApp.Log Me, Object, method, Message, LogDetail, ErrorID, LogSource, LogSourceDetail
        Select Case LogDetail
            Case GPS_LOG_ERROR
                currentStat = GPS_STAT_ERROR
            Case GPS_LOG_WARNING
                currentStat = GPS_STAT_WARNING
            Case Else
                currentStat = Process_Status
        End Select
        If currentStat <> Process_Status Then Process_Status = currentStat
    End If
    Exit Sub
err_Log:
    Err.Clear
End Sub

Private Function getSettings(ByVal PropertyName As String, ByVal defaultValue As String, Optional ByVal bXML As Boolean = False) As String
Dim eReturn As getPropertyReturnCode
Dim sReturn As String
    If Not localSettings Is Nothing Then
        sReturn = getProperty(localSettings, PropertyName, defaultValue, bXML, eReturn)
        Select Case eReturn
            Case gprcFound
                Log cModuleName, "getSettings", "Property: " & PropertyName & " = " & sReturn, GPS_LOG_VERBOSE
            Case gprcDefault
                Log cModuleName, "getSettings", "Unable to find element - passing back default! Property: " & PropertyName & " = " & defaultValue, GPS_LOG_WARNING
            Case gprcBadNode
                Log cModuleName, "getSettings", "No child nodes in root of element - passing back default! Property: " & PropertyName & " = " & defaultValue, GPS_LOG_WARNING
        End Select
        getSettings = sReturn
    Else
        Log cModuleName, "getSettings", "No settings found - passing back default! Property: " & PropertyName & " = " & defaultValue, GPS_LOG_WARNING
        getSettings = defaultValue
    End If
End Function

Private Sub setProperties()

Dim currentType As HTE_GPS.GPSConfiguration
Dim initString As String, sHost As String, lPort As Long
Dim varArray As Variant
Dim inbound As HTE_GPS.GPS_MESSAGING_TYPES
Dim i As Long, x As Long
Dim bAttempt As Boolean, attemptNumber As Long
Const MAX_UDP_BIND_ATTEMPTS = 20

On Error GoTo err_setProperties
    If Not localSettings Is Nothing Then
        UnloadSafely
        With AliasLookup
            .EnableMacResolution = CBool(getSettings(cEnableMacResolution, cEnableMacResolutionValue))
            .MacLookupExpiration = CLng(Val(getSettings(cCacheMacLookupSeconds, cCacheMacLookupSecondsValue)))
        End With
        Set fMain = New frmMain
        With fMain
            .ValidateMessage = CBool(getSettings(cValidateMessage, cValidateMessageValue))
            .Winsock1.Protocol = Abs(UCase$(getSettings(cProtocol, cProtocolValue)) <> "TCP")
            Log cModuleName, "setProperties", "Protocol = " & CStr(.Winsock1.Protocol), GPS_LOG_VERBOSE
            sHost = getSettings(cAddress, ComputerName)
            Log cModuleName, "setProperties", "HostName = " & sHost, GPS_LOG_VERBOSE
            lPort = getSettings(cPort, cPortValue)
            Log cModuleName, "setProperties", "Port = " & lPort, GPS_LOG_VERBOSE
            bProcessOnSend = getSettings(cProcessOnSend, cProcessOnSendValue)
            Log cModuleName, "setProperties", "Setting matrix of available messages", GPS_LOG_VERBOSE
            currentTypes = SupportedMessages(localSettings)
            Log cModuleName, "setProperties", "Set matrix of available messages", GPS_LOG_VERBOSE
            inbound = gApp.InboundType(Me)
            For i = LBound(currentTypes) To UBound(currentTypes)
                If inbound = currentTypes(i).GPSType Then
                    .MessageType = currentTypes(i)
                    Exit For
                ElseIf i = UBound(currentTypes) Then
                    Log cModuleName, "setProperties", "No current type defined!", GPS_LOG_WARNING
                End If
            Next
            .Host = sHost
            .Port = lPort
            If .Winsock1.Protocol = sckUDPProtocol Then
                Log cModuleName, "setProperties", "Binding port...", GPS_LOG_VERBOSE
                bAttempt = True
                attemptNumber = 1
                On Error Resume Next
                Do While bAttempt
                    .Winsock1.LocalPort = getSettings(cLocalPort, cLocalPortValue)
                    .Winsock1.Bind
                    If Err.Number <> 0 Then
                        Log cModuleName, "setProperties", Err.Number & ": " & Err.Description, GPS_LOG_ERROR
                        Err.Clear
                        Log cModuleName, "setProperties", "Pausing for system resources...", GPS_LOG_WARNING
                        Sleep 500
                        DoEvents
                        Log cModuleName, "setProperties", "Binding port...", GPS_LOG_VERBOSE
                        If attemptNumber > MAX_UDP_BIND_ATTEMPTS Then
                            Log cModuleName, "setProperties", "Cannot bind after [" & MAX_UDP_BIND_ATTEMPTS & "] attempts!", GPS_LOG_ERROR
                            bAttempt = False
                            Process_Status = GPS_STAT_ERROR
                            Exit Sub
                        End If
                    Else
                        bAttempt = False
                        Log cModuleName, "setProperties", "Successful bind to port [" & getSettings(cLocalPort, cLocalPortValue) & "]", GPS_LOG_INFORMATION
                        On Error GoTo err_setProperties
                    End If
                    attemptNumber = attemptNumber + 1
                Loop
            Else
                Log cModuleName, "setProperties", "Connecting to host...", GPS_LOG_VERBOSE
                On Error Resume Next
                .Winsock1.Connect .Host, .Port
                If Err.Number <> 0 Then
                    Log cModuleName, "setProperties", Err.Number & ": " & Err.Description, GPS_LOG_ERROR
                    Err.Clear
                End If
                On Error GoTo err_setProperties
            End If
            initString = getSettings(cInitString, cInitStringValue)
            If initString <> vbNullString Then
                varArray = Split(initString, "|*|")
                For x = 0 To UBound(varArray)
                    .bufferToSend CStr(varArray(x)), False
                Next
            Else
                Log cModuleName, "setProperties", "No initialization string present!", GPS_LOG_VERBOSE
            End If
            If CBool(getSettings(cManualPoll, cManualPollValue)) Then
                .pollingBuffer UnscreenMessage(getSettings(cManualPollString, cManualPollStringValue)), getSettings(cManualPollInterval, cManualPollIntervalValue)
                .LimitMessages CLng(getSettings(cRelayInterval, cRelayIntervalValue)), CLng(getSettings(cDurationInterval, cDurationIntervalValue))
            Else
                .pollingBuffer vbNullString, vbNullString
                .LimitMessages 0, 0
            End If
            
        End With
        Process_Status = GPS_STAT_READYANDWILLING
    Else
        Process_Status = GPS_STAT_UNKNOWN
    End If
    Exit Sub
err_setProperties:
    Log cModuleName, "setProperties", Err.Description, GPS_LOG_ERROR, Err.Number
    Err.Clear
End Sub

Private Sub PropPageCallback_Exit()
    Set m_Property = Nothing
End Sub

Private Function PropPageCallback_SaveChanges(ByVal XMLDOMNode As String) As Boolean
On Error GoTo err_PropPageCallback
    If loadLocalSettings(XMLDOMNode, False) Then
        If checkForChanges Then
            setProperties
            m_Settings = retrieveLocalSettings
            PropPageCallback_SaveChanges = True
        End If
        PropPageCallback_SaveChanges = True
    Else
        Log cModuleName, "PropPageCallback_SaveChanges", "Unable to reload settings from string.", GPS_LOG_ERROR, , XMLDOMNode, GPS_SOURCE_XML
    End If
    Exit Function
err_PropPageCallback:
    Log cModuleName, "PropPageCallback_SaveChanges", Err.Description, GPS_LOG_WARNING, Err.Number
End Function

Private Function StartWinsock(sDescription As String, Optional Version As VersionRequired = SOCKET_VERSION_11) As Boolean
Dim StartupData As WSADataType
    If Not WSAStartedUp Then
        If Not WSAStartup(Version, StartupData) Then
            WSAStartedUp = True
            sDescription = StartupData.szDescription
        Else
            WSAStartedUp = False
        End If
    End If
    StartWinsock = WSAStartedUp
End Function

Public Sub EndWinsock()
Dim ret&
    If WSAIsBlocking() Then
        ret = WSACancelBlockingCall()
    End If
    ret = WSACleanup()
    WSAStartedUp = False
End Sub

Private Static Function ByteToHex(bytVal As Byte) As String
  ByteToHex = "00"
  Mid$(ByteToHex, 3 - Len(Hex$(bytVal))) = Hex$(bytVal)
End Function

Private Function ScreenMessage(ByVal Message As String) As String
'PRB: Error Message When an XML Document Contains Low-Order ASCII Characters
'http://support.microsoft.com/?kbid=315580
Dim x As Byte
    If Len(Message) > 0 Then
        For x = 0 To 31
            If InStr(1, Message, Chr$(x), vbBinaryCompare) > 0 Then
                Message = Replace(Message, Chr$(x), "/#*x" & ByteToHex(x) & ";")
            End If
        Next
    End If
    ScreenMessage = Message
    Log cModuleName, "ScreenMessage", "Message transformed for XML.", GPS_LOG_VERBOSE, , Message, GPS_SOURCE_STRING
End Function
Private Function UnscreenMessage(ByVal Message As String) As String
'PRB: Error Message When an XML Document Contains Low-Order ASCII Characters
'http://support.microsoft.com/?kbid=315580
Dim x As Byte
Dim sTemp As String * 2
    If Len(Message) > 0 Then
        For x = 0 To 31
            If InStr(1, Message, "/#*x", vbBinaryCompare) = 0 Then Exit For
            sTemp = ByteToHex(x)
            Message = Replace(Message, "/#*x" & sTemp & ";", Chr$(x))
        Next
    End If
    UnscreenMessage = Message
    Log cModuleName, "UnScreenMessage", "Message transformed from XML.", GPS_LOG_VERBOSE, , Message, GPS_SOURCE_STRING
End Function
