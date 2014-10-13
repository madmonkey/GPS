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
'*TODO - change names of projects and helper page variables to reflect your project
Private Const cModuleName = "HTE_UDP_Transport.Process"
Private gApp As HTE_GPS.Processes
Private m_Status As HTE_GPS.GPS_PROCESSOR_STATUS
Private localSettings As MSXML2.DOMDocument30
Private m_Settings As String
Private m_Property As Object
Implements HTE_GPS.Process
Implements HTE_GPS.PropPageCallback
Private WithEvents udpTransport As SunGard_Communications.CommunicationsEngine
Attribute udpTransport.VB_VarHelpID = -1
'Private m_Connected As Boolean 'do we have established communications?
Private m_ReceiveMessages As Boolean 'are we configured to receive
Private m_msgcounter As Long 'unique message identifier - used for send!
Private m_Identifier As String 'Instance ID for Host
Private m_RemoteAddr As String
Private m_RemotePort As Long
Private m_LocalAddr As String
Private m_LocalPort As Long
Private m_SegmentSize As Long
Private m_UseEncryption As Boolean
Private m_UseCompression As Boolean
Private m_ReceiveBufferSize As Long
Private m_KeepAliveInterval As Long
Private m_UseKeepAlive As Boolean
Private m_RetryInterval As Long
Private m_RetryRandomInterval As Long
Private m_KeepAliveFailureInterval As Long
Private m_MaxKeepAliveFailures As Long
Private m_MaxTimeToLive As Long 'in seconds
Private m_ProcessCompleteOnSend As Boolean
Private m_SerializeMsgStructure As Boolean
Private m_defaultRouteType As HTE_GPS.GPS_MESSAGING_TYPES
Private WithEvents retryTimer As ccrpTimers6.ccrpTimer
Attribute retryTimer.VB_VarHelpID = -1

'=====================================================================
'======================GPS PROCESS REGION=============================
'=====================================================================
Private Sub Class_Initialize()
    m_Status = GPS_STAT_INITIALIZED
    Set localSettings = New MSXML2.DOMDocument30
    Set retryTimer = New ccrpTimers6.ccrpTimer
    With retryTimer
        .EventType = TimerPeriodic
        .Interval = 1250
        .Enabled = False
    End With
End Sub

Private Sub Class_Terminate()
    m_Status = GPS_STAT_UNINITIALIZED
    Set localSettings = Nothing
    If Not retryTimer Is Nothing Then
        retryTimer.Enabled = False
        Set retryTimer = Nothing
    End If
End Sub

Private Sub Process_Connect(ByVal parentProcess As HTE_GPS.Processes, ByVal cookie As String)
    Set gApp = parentProcess
    gApp.Connected Me, cookie 'WHEN YOU CONNECT CALL CONNECTED METHOD TO RETRIEVE YOUR INSTANCE SETTINGS
End Sub

Private Sub Process_Disconnect()
    Log cModuleName, "Process_Disconnect", "Unhooking process", GPS_LOG_VERBOSE
    checkForChanges
    Set m_Property = Nothing
    Set gApp = Nothing
End Sub

Private Property Get Process_FriendlyName() As String
    Process_FriendlyName = "UDP Transport"
End Property

Private Property Get Process_HostVersion() As String
    Process_HostVersion = "1.0.0"
End Property

Private Sub Process_ProcessMessage(ByVal PropData As HTE_GPS.GPSMessage)
Dim oMsg As Message
    Log cModuleName, "Process_ProcessMessage", "Serialize Message Structure = " & CStr(m_SerializeMsgStructure), GPS_LOG_VERBOSE
    If m_SerializeMsgStructure Then
        Set oMsg = New Message
        SendMessage oMsg.Serialize(PropData)
    Else
        SendMessage PropData.rawMessage
    End If
    If m_ProcessCompleteOnSend Then gApp.ProcessComplete PropData, Me
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
        If Not gApp Is Nothing Then gApp.StatusChanged m_Status, Me
    End If
End Property

Private Property Get Process_Status() As HTE_GPS.GPS_PROCESSOR_STATUS
    Process_Status = m_Status
End Property

'=====================================================================
'=====================END GPS PROCESS REGION==========================
'=====================================================================

Private Function checkForChanges() As Boolean
   Dim bBadNode As Boolean
    If Not gApp Is Nothing Then
        If modProperties.PropertiesChanged(m_Settings, localSettings, bBadNode) Then
            If bBadNode Then
                Log cModuleName, "checkForChanges", "Unable to verify against initial settings.", GPS_LOG_WARNING, , m_Settings, GPS_SOURCE_XML
            Else
                checkForChanges = True
                gApp.PropertiesChanged retrieveLocalSettings, Me
            End If
        End If
    End If
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

Private Function getSettings(ByVal PropertyName As String, ByVal defaultValue As String, Optional ByVal bXML As Boolean = False) As String
Dim eReturn As getPropertyReturnCode
Dim sReturn As String
    If Not localSettings Is Nothing Then
        sReturn = modProperties.getProperty(localSettings, PropertyName, defaultValue, bXML, eReturn)
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

Private Sub Log(ByVal Object As String, ByVal method As String, Optional ByVal Message As String = "", Optional ByVal LogDetail As HTE_GPS.GPS_LOG_DETAIL = 8&, Optional ByVal ErrorID As Long = 0&, Optional ByVal LogSource As String = "", Optional ByVal LogSourceDetail As HTE_GPS.GPS_LOG_SOURCE = 0&)
Dim currentStat As HTE_GPS.GPS_PROCESSOR_STATUS
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
End Sub
Private Sub setProperties()
Dim sTemp As String
On Error GoTo err_setProperties
    If Not localSettings Is Nothing Then
        Disconnect
        'REMOTE ADDRESS
        m_RemoteAddr = getProperty(localSettings, cRemoteAddr, cRemoteAddrValue)
        Log cModuleName, "setProperties", "RemoteAddress = " & m_RemoteAddr, GPS_LOG_VERBOSE
        'REMOTE PORT
        sTemp = getProperty(localSettings, cRemotePort, cRemotePortValue)
        If ValidPort(sTemp) Then
            m_RemotePort = CLng(sTemp)
        Else
            Log cModuleName, "setProperties", "RemotePort is invalid - setting to default (" & CStr(cRemotePortValue) & ")!", GPS_LOG_WARNING
            m_RemotePort = cRemotePortValue
        End If
        'LOCAL ADDRESS
        m_LocalAddr = getProperty(localSettings, cLocalAddr, cLocalAddrValue)
        Log cModuleName, "setProperties", "LocalAddress = " & m_LocalAddr, GPS_LOG_VERBOSE
        'LOCAL PORT
        sTemp = getProperty(localSettings, cLocalPort, cLocalPortValue)
        If ValidPort(sTemp) Then
            m_ReceiveMessages = (CLng(sTemp) <> 0)
            m_LocalPort = CLng(sTemp)
            Log cModuleName, "setProperties", "LocalPort = " & CStr(m_LocalPort), GPS_LOG_VERBOSE
        Else
            Log cModuleName, "setProperties", "LocalPort is invalid - setting to default (" & CStr(cLocalPortValue) & ")!", GPS_LOG_WARNING
            m_LocalPort = cLocalPortValue
        End If
        'SEGMENT SIZE
        sTemp = getProperty(localSettings, cSegmentSize, cSegmentSizeValue)
        If ValidSize(sTemp) Then
            m_SegmentSize = CLng(sTemp)
            Log cModuleName, "setProperties", "SegmentSize = " & CStr(m_SegmentSize), GPS_LOG_VERBOSE
        Else
            Log cModuleName, "setProperties", "SegmentSize is invalid - setting to default (" & CStr(cSegmentSizeValue) & ")!", GPS_LOG_WARNING
            m_SegmentSize = cSegmentSizeValue
        End If
        'ENCRYPTION
        m_UseEncryption = CBool(getProperty(localSettings, cUseEncryption, cUseEncryptionValue))
        Log cModuleName, "setProperties", "UseEncryption = " & CStr(m_UseEncryption), GPS_LOG_VERBOSE
        'COMPRESSION
        m_UseCompression = CBool(getProperty(localSettings, cUseCompression, cUseCompressionValue))
        Log cModuleName, "setProperties", "UseCompression = " & CStr(m_UseCompression), GPS_LOG_VERBOSE
        'RECEIVEBUFFER
        sTemp = getProperty(localSettings, cReceiveBufferSize, cReceiveBufferSizeValue)
        If ValidSize(sTemp) Then
            m_ReceiveBufferSize = CLng(sTemp)
            Log cModuleName, "setProperties", "ReceiveBuffer = " & CStr(m_ReceiveBufferSize), GPS_LOG_VERBOSE
        Else
            Log cModuleName, "setProperties", "ReceiveBuffer is invalid - setting to default - (" & CStr(cReceiveBufferSizeValue) & ")!", GPS_LOG_WARNING
            m_ReceiveBufferSize = cReceiveBufferSizeValue
        End If
        'KEEPALIVE
        m_KeepAliveInterval = CLng(getProperty(localSettings, cKeepAliveInterval, cKeepAliveIntervalValue))
        Log cModuleName, "setProperties", "KeepAliveInterval = " & CStr(m_KeepAliveInterval), GPS_LOG_VERBOSE
        'USE KEEPALIVES
        m_UseKeepAlive = CBool(getProperty(localSettings, cUseKeepAlive, cUseKeepAliveValue))
        Log cModuleName, "setProperties", "UseKeepAlive = " & CStr(m_UseKeepAlive), GPS_LOG_VERBOSE
        'RETRY INTERVAL
        m_RetryInterval = CLng(getProperty(localSettings, cRetryInterval, cRetryIntervalValue))
        Log cModuleName, "setProperties", "RetryInterval = " & CStr(m_RetryInterval), GPS_LOG_VERBOSE
        'RANDOMIZE
        m_RetryRandomInterval = CLng(getProperty(localSettings, cRetryRandomInterval, cRetryRandomIntervalValue))
        Log cModuleName, "setProperties", "RetryRandomInterval = " & CStr(m_RetryRandomInterval), GPS_LOG_VERBOSE
        'KEEPALIVE FAILURE INTERVAL
        m_KeepAliveFailureInterval = CLng(getProperty(localSettings, cKeepAliveFailureInterval, cKeepAliveFailureIntervalValue))
        Log cModuleName, "setProperties", "KeepAliveFailureInterval = " & CStr(m_KeepAliveFailureInterval), GPS_LOG_VERBOSE
        'MAX KEEP ALIVE FAILURES
        m_MaxKeepAliveFailures = CLng(getProperty(localSettings, cMaxKeepAliveFailures, cMaxKeepAliveFailuresValue))
        Log cModuleName, "setProperties", "MaxKeepAliveFailures = " & CStr(m_MaxKeepAliveFailures), GPS_LOG_VERBOSE
        m_MaxTimeToLive = CLng(getProperty(localSettings, cTimeToLive, cTimeToLiveValue))
        Log cModuleName, "setProperties", "MaxTimeToLive = " & CStr(m_MaxTimeToLive), GPS_LOG_VERBOSE
        m_ProcessCompleteOnSend = CBool(getProperty(localSettings, cProcessOnSend, cProcessOnSendValue))
        Log cModuleName, "setProperties", "ProcessCompleteOnSend = " & CStr(m_ProcessCompleteOnSend), GPS_LOG_VERBOSE
        m_SerializeMsgStructure = CBool(getProperty(localSettings, cSerializeMsg, cSerializeMsgValue))
        m_defaultRouteType = gApp.InboundType(Me)
        If Connect Then
            Process_Status = GPS_STAT_READYANDWILLING
        Else
            Process_Status = GPS_STAT_ERROR
        End If
    Else
        Process_Status = GPS_STAT_UNKNOWN
    End If
    Exit Sub
err_setProperties:
    Log cModuleName, "setProperties", Err.Description, GPS_LOG_ERROR, Err.Number
End Sub

Private Function ValidPort(ByRef Port As Variant) As Boolean
    If IsNumeric(Port) Then
        ValidPort = CLng(Port) <= 32765 And CLng(Port) >= 0
    End If
End Function
Private Function ValidSize(ByRef Segment As Variant) As Boolean
    If IsNumeric(Segment) Then
        ValidSize = CLng(Segment) <= 4096 And CLng(Segment) > 0
    End If
End Function
'=====================================================================
'======================PROPERTY CALLBACK==============================
'=====================================================================
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
        Else
            PropPageCallback_SaveChanges = True
        End If
    Else
        Log cModuleName, "PropPageCallback_SaveChanges", "Unable to reload settings from string.", GPS_LOG_ERROR, , XMLDOMNode, GPS_SOURCE_XML
    End If
    Exit Function
err_PropPageCallback:
    Log cModuleName, "PropPageCallback_SaveChanges", Err.Description, GPS_LOG_WARNING, Err.Number
End Function

'=====================================================================
'=====================END PROPERTY CALLBACK===========================
'=====================================================================

'=====================================================================
'=======================UDP COMM ENGINE===============================
'=====================================================================

Private Sub udpTransport_OnKeepaliveElapsed(ByVal s As SunGard_Communications.ICommunicationSettings)
    Log cModuleName, "udpTransport_OnKeepaliveElapsed", _
        "No Keep-alive acknowledgement from " & s.address & _
        " [" & s.MacAddress & _
        "] unit may be out of range, encountering network difficulties or the hosting system may be experiencing problems." _
        , GPS_LOG_WARNING
End Sub

Private Sub udpTransport_OnKeepaliveReceived(ByVal s As SunGard_Communications.ICommunicationSettings)
    Log cModuleName, "udpTransport_OnKeepaliveReceived", "Keep-alive acknowledged from " & _
        s.address & " [" & s.MacAddress & "]", GPS_LOG_VERBOSE
    Process_Status = GPS_STAT_READYANDWILLING
End Sub

Private Sub udpTransport_OnMessageDeliverySuccess(ByVal s As SunGard_Communications.ICommunicationSettings, lookup() As Byte)
    Log cModuleName, "udpTransport_OnMessageDeliverySuccess", _
        "Message [" & StringFromByteArray(lookup) & _
        "] and all related packet(s) were acknowledged by intended destination " & _
        s.address & " [" & s.MacAddress & "].", GPS_LOG_VERBOSE
    Process_Status = GPS_STAT_READYANDWILLING
End Sub

Private Sub udpTransport_OnMessageTimeout(ByVal s As SunGard_Communications.ICommunicationSettings, lookup() As Byte, data() As Byte)
    Log cModuleName, "udpTransport_OnMessageTimeout", _
    "Message [" & StringFromByteArray(lookup) & "] timed out! No acknowledgement from " & _
    s.address & " [" & s.MacAddress & "]!", GPS_LOG_WARNING, , StringFromByteArray(data), GPS_SOURCE_BINARY
End Sub

Private Sub udpTransport_OnMessageTransmission(ByVal s As SunGard_Communications.ICommunicationSettings, lookup() As Byte)
    Log cModuleName, "udpTransport_OnMessageTransmission", _
        "Message [" & StringFromByteArray(lookup) & _
        "] and all related packet(s) were sent to network ONLY. Intended destination " & _
        s.address & " [" & s.MacAddress & "] has not yet acknowledged and may not acknowledge.", _
        GPS_LOG_VERBOSE
End Sub

Private Sub udpTransport_OnReceiveData(ByVal s As SunGard_Communications.ICommunicationSettings, data() As Byte)
Dim oMsg As Message
Dim oEnt As HTE_Entity.Entity
Dim sEnt As String
    Set oMsg = New Message
    Log cModuleName, "udpTransport_OnReceiveData", "Message received from " & s.address & " [" & s.MacAddress & "]!", GPS_LOG_VERBOSE, , StringFromByteArray(data), GPS_SOURCE_BINARY
    'DESERIALIZE AND RAISE UP TO APP IF VALID MESSAGE
    With oMsg
        Set oEnt = New HTE_Entity.Entity
        With oEnt
            .IPAddress = s.address
            .MacAddress = s.MacAddress
        End With
        sEnt = oEnt.Serialize
        Log cModuleName, "udpTransport_OnReceiveData", "Endpoint identified!", GPS_LOG_VERBOSE, , sEnt, GPS_SOURCE_XML
        Set oEnt = Nothing
        If m_SerializeMsgStructure Then
            If .Deserialize(StringFromByteArray(data)) Then
                oMsg.Entity = sEnt
                If Not gApp Is Nothing Then gApp.ProcessComplete oMsg, Me
            Else
                Log cModuleName, "udpTransport_OnReceiveData", "Message is NOT in expected format - dropping from delivery!", GPS_LOG_WARNING
            End If
        Else
            'don't really know - assume off the current Route configuration type
            .rawMessage = StringFromByteArray(data)
            .MessageType = m_defaultRouteType
            .MessageStatus = GPS_MSG_PROCESSED
            .Entity = sEnt
            If Not gApp Is Nothing Then gApp.ProcessComplete oMsg, Me
        End If
        Process_Status = GPS_STAT_READYANDWILLING
    End With
End Sub

'=====================================================================
'======================END UDP COMM ENGINE============================
'=====================================================================

Private Sub Disconnect()
    If Not udpTransport Is Nothing Then
        If m_ReceiveMessages Then udpTransport.FinishReceiving
        Set udpTransport = Nothing
        Log cModuleName, "Disconnect", "Finished receiving and disposed!", GPS_LOG_VERBOSE
    End If
End Sub

Private Function Connect() As Boolean
On Error GoTo err_Connect
    If udpTransport Is Nothing Then Set udpTransport = New SunGard_Communications.CommunicationsEngine
    InitializeEngine
    If Not udpTransport Is Nothing Then
        If m_ReceiveMessages Then udpTransport.StartReceiving
        Connect = ((Err.Number = 0) And (Err.LastDllError = 0))
    Else
        Log cModuleName, "Connect", "Uninitialized communication engine!", GPS_LOG_ERROR
    End If
    Exit Function

err_Connect:
    Log cModuleName, "Connect", "Critical error initializing!", GPS_LOG_ERROR, Err.Number, Err.Source, GPS_SOURCE_STRING
    retryTimer.Enabled = True
End Function

Private Sub retryTimer_Timer(ByVal Milliseconds As Long)
On Local Error Resume Next
    Log cModuleName, "retryTimer_Timer", "Retrying connection...", GPS_LOG_INFORMATION
    retryTimer.Enabled = False
    If Not Connect Then
        Log cModuleName, "retryTimer_Timer", "Reconnection failed!", GPS_LOG_WARNING
    Else
        Log cModuleName, "retryTimer_Timer", "Reconnection success!", GPS_LOG_INFORMATION
    End If
End Sub

Private Sub InitializeEngine()
Dim udpSettings As SunGard_Communications.CommunicationSettings
    
    'set-up local first
    Set udpSettings = udpTransport.Listener
    With udpSettings
        .address = m_LocalAddr
        .Port = m_LocalPort
        .ReceiveBufferSize = m_ReceiveBufferSize
    End With
    
    Set udpSettings = udpTransport.Item(m_RemoteAddr)
    With udpSettings
        .address = m_RemoteAddr
        .Port = m_RemotePort
        .SegmentSize = m_SegmentSize
        .KeepaliveFailureInterval = m_KeepAliveFailureInterval
        .KeepaliveInterval = m_KeepAliveInterval
        .MaxKeepaliveFailures = m_MaxKeepAliveFailures
        .RetryInterval = m_RetryInterval
        .RetryIntervalRandomizer = m_RetryRandomInterval
        .UseCompression = m_UseCompression
        .UseEncryption = m_UseEncryption
        .UseKeepalives = m_UseKeepAlive
        m_Identifier = .lookup
    End With
    
End Sub
Private Sub SendMessage(ByRef buffer As String)
Const cMAXLONG As Long = 2147483647
Dim messageParam As SunGard_Communications.MessageParameters

On Error GoTo err_SendMessage
    If Not udpTransport Is Nothing Then
        If m_msgcounter = cMAXLONG Then
            m_msgcounter = 0
            Log cModuleName, "SendMessage", "Reset message counter...", GPS_LOG_VERBOSE
        End If
        m_msgcounter = m_msgcounter + 1
        Set messageParam = New SunGard_Communications.MessageParameters
        If Not messageParam Is Nothing Then
            With messageParam
                .AddDestination ByteArrayFromString(m_Identifier)
                .SetData ByteArrayFromString(buffer)
                .SetMessageLookup ByteArrayFromString(CStr(m_msgcounter))
                .TimeToLive = m_MaxTimeToLive
                Log cModuleName, "SendMessage", "Sending Message (" & .TimeToLive & " second(s) to live.)", GPS_LOG_VERBOSE, , buffer, GPS_SOURCE_BINARY
            End With
            udpTransport.send messageParam 'ByteArrayFromString(m_Identifier), ByteArrayFromString(StatusTask), ByteArrayFromString(Buffer)
        Else
            Log cModuleName, "SendMessage", "Unable to instanciate message parameters object!", GPS_LOG_ERROR
        End If
    Else
        Log cModuleName, "SendMessage", HostName & " is NOT initialized.", GPS_LOG_WARNING
    End If
    Exit Sub
    
err_SendMessage:
    Log cModuleName, "SendMessage", Err.Number & ":" & Err.Description, GPS_LOG_ERROR, Err.LastDllError
    
End Sub

Private Property Get HostName() As String
    If Not udpTransport Is Nothing Then
        HostName = udpTransport.Item(m_Identifier).Item("RemoteHost")
    Else
        HostName = TypeName(udpTransport)
    End If
End Property

Private Function ByteArrayFromString(ByRef Source$) As Variant
   Dim Buf() As Byte
   Dim r$
   r$ = StrConv(Source$, vbFromUnicode)
   Buf() = r$
   ByteArrayFromString = Buf()
End Function

Private Function StringFromByteArray(vr As Variant) As String
   Dim Buf() As Byte
   Dim r$
   Buf() = vr
   r$ = Buf()

   StringFromByteArray = StrConv(r$, vbUnicode)
   
End Function


