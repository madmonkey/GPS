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

Private Const cModuleName = "HTE_Comport.Process"
Private gApp As HTE_GPS.Processes
Private m_Status As HTE_GPS.GPS_PROCESSOR_STATUS
Private localSettings As MSXML2.DOMDocument30
Private m_Settings As String
Private m_Property As Object
Implements HTE_GPS.Process
Implements HTE_GPS.PropPageCallback
Dim WithEvents fMain As frmMain
Attribute fMain.VB_VarHelpID = -1
Private ProcessOnSend As Boolean

Private Sub Class_Initialize()
    m_Status = GPS_STAT_INITIALIZED
    Set localSettings = New MSXML2.DOMDocument30
    Set fMain = New frmMain
    Load fMain
End Sub

Private Sub Class_Terminate()
    m_Status = GPS_STAT_UNINITIALIZED
    Set localSettings = Nothing
End Sub

Private Sub fMain_Log(ByVal method As String, ByVal Message As String, ByVal LogDetail As HTE_GPS.GPS_LOG_DETAIL)
    Log cModuleName, method, Message, LogDetail
End Sub

Private Sub fMain_LogDetail(ByVal method As String, ByVal Message As String, ByVal LogDetail As HTE_GPS.GPS_LOG_DETAIL, ByVal ErrorID As Long, ByVal LogSource As String, ByVal LogSourceDetail As HTE_GPS.GPS_LOG_SOURCE)
    Log cModuleName, method, Message, LogDetail, ErrorID, LogSource, LogSourceDetail
End Sub

Private Sub fMain_ReceivedMessage(ByVal DataIn As String)
Dim msg As Message
    Log cModuleName, "fMain_ReceivedMessage", "Received message.", GPS_LOG_INFORMATION
    Log cModuleName, "fMain_ReceivedMessage", "Received message from listener.", GPS_LOG_VERBOSE, , DataIn, GPS_SOURCE_BINARY
    Set msg = New Message
    msg.MessageStatus = GPS_MSG_PROCESSED
    msg.MessageType = fMain.MessageType.GPSType
    msg.rawMessage = DataIn
    If Not gApp Is Nothing Then gApp.ProcessComplete msg, Me
    'Process_Status GPS_STAT_READYANDWILLING 'receiving messages, we should be ok!
End Sub

Private Sub fMain_StatusChange(ByVal statusCode As HTE_GPS.GPS_PROCESSOR_STATUS)
    Process_Status = statusCode
End Sub

Private Sub Process_Connect(ByVal parentProcess As HTE_GPS.Processes, ByVal cookie As String)
    Set gApp = parentProcess
    gApp.Connected Me, cookie
End Sub

Private Sub Process_Disconnect()
    Log cModuleName, "Process_Disconnect", "Unhooking process", GPS_LOG_VERBOSE
    checkForChanges
    Unload fMain
    Set fMain = Nothing
    Set m_Property = Nothing
    Set gApp = Nothing
End Sub

Private Property Get Process_FriendlyName() As String
    Process_FriendlyName = "Communications Port Process"
End Property

Private Property Get Process_HostVersion() As String
    Process_HostVersion = "1.0.0"
End Property

Private Sub Process_ProcessMessage(ByVal PropData As HTE_GPS.GPSMessage)
    If Not fMain Is Nothing Then
        fMain.bufferToSend = PropData.rawMessage
        Log cModuleName, "Process_ProcessMessage", "Message processed and sent.", GPS_LOG_VERBOSE, , PropData.rawMessage, GPS_SOURCE_BINARY
        'if being used as a repeater, then "bubble-up" information as received
        If ProcessOnSend Then
            If Not gApp Is Nothing Then gApp.ProcessComplete PropData, Me
        End If
    Else
        Log cModuleName, "Process_ProcessMessage", "Unable to send message.", GPS_LOG_WARNING, , PropData.rawMessage, GPS_SOURCE_BINARY
    End If
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
Dim vVar As Variant
Dim i As Long, x As Long
Dim current As HTE_GPS.GPS_MESSAGING_TYPES
Dim currentTypes() As HTE_GPS.GPSConfiguration
Dim GPSTypeInfo As HTE_GPS.GPSConfiguration
Dim sValidPorts As String, bAssignedPort As Boolean, requestedPort As Long
Dim communicationSettings As CommSettings
Const MAX_LONG As Long = 2147483647

On Error GoTo err_setProperties
    If Not localSettings Is Nothing Then
        If Not fMain Is Nothing Then
            Set communicationSettings = New CommSettings
            Log cModuleName, "setProperties", "Available Port(s) are (" & communicationSettings.AvailablePorts & ")", GPS_LOG_INFORMATION
            requestedPort = CLng(getSettings(cComm, cCommValue))
            With communicationSettings
                .Port = requestedPort
                .Settings = getSettings(cSettings, cSettingsValue)
                .RThreshold = getSettings(cRThresh, cRThreshValue)
                .InputLen = CInt(Val(getSettings(cInputLen, cInputLenValue)))
                .InputMode = Val(getSettings(cInputMode, cInputModeValue))
                .DTREnable = CBool(getSettings(cDTREnable, cDTREnableValue))
                .EOFEnable = CBool(getSettings(cEOFEnable, cEOFEnableValue))
                .Handshaking = Val(getSettings(cHandshaking, cHandshakingValue))
                .InBufferSize = CInt(getSettings(cInBufferSize, cInBufferSizeValue))
                .NullDiscard = CBool(getSettings(cNullDiscard, cNullDiscardValue))
                .RTSEnable = CBool(getSettings(cRTSEnable, cRTSEnableValue))
                .InitializationString = getSettings(cInitString, cInitStringValue)
                ProcessOnSend = CBool(getSettings(cProcessOnSend, cProcessOnSendValue))
                If Val(getSettings(cMaxCacheBufferSize, cMaxCacheBufferSizeValue)) > MAX_LONG Or _
                    Val(getSettings(cMaxCacheBufferSize, cMaxCacheBufferSizeValue)) < 1 Then
                    Log cModuleName, "setProperties", "Invalid value [" & getSettings(cMaxCacheBufferSize, cMaxCacheBufferSizeValue) & _
                        "]  specified for MaxCachedBytes - using default[" & CStr(cMaxCacheBufferSizeValue) & "] !", GPS_LOG_WARNING
                    .MaxCachedBytes = cMaxCacheBufferSizeValue
                Else
                    .MaxCachedBytes = CLng(Val(getSettings(cMaxCacheBufferSize, cMaxCacheBufferSizeValue)))
                End If
                If .IsAssignable Then
                    Log cModuleName, "setProperties", "Port signalled to be opened.", GPS_LOG_VERBOSE
                Else
                    Log cModuleName, "setProperties", "Invalid Port specified port NOT opened. Available Port(s) are (" & .AvailablePorts & ")", GPS_LOG_ERROR
                End If
            End With
            current = gApp.InboundType(Me)
            Log cModuleName, "setProperties", "Set current message type", GPS_LOG_VERBOSE
            Log cModuleName, "setProperties", "Setting Inbound Type as " & CStr(current), GPS_LOG_VERBOSE
'''            #If DEBUGMODE = 1 Then
'''                Log cModuleName, "setProperties", "localSettings.documentElement.getElementsByTagName(TYPES) Is Nothing = " & CStr(localSettings.documentElement.getElementsByTagName("TYPES") Is Nothing)
'''                Log cModuleName, "setProperties", "localSettings.documentElement.getElementsByTagName(TYPES).Item(0) Is Nothing = " & CStr(localSettings.documentElement.getElementsByTagName("TYPES").Item(0) Is Nothing)
'''                Log cModuleName, "setProperties", "localSettings.documentElement.getElementsByTagName(TYPES).Item(0).hasChildNodes = " & CStr(localSettings.documentElement.getElementsByTagName("TYPES").Item(0).hasChildNodes)
'''                Log cModuleName, "setProperties", "locaSettings.documentElement.getElementsByTagName(TYPES).Item(0).childNodes.length = " & CStr(localSettings.documentElement.getElementsByTagName("TYPES").Item(0).childNodes.length)
'''            #End If
            currentTypes = SupportedMessages(localSettings)
            Log cModuleName, "setProperties", "Parsed supported messages: total = " & CStr(UBound(currentTypes) + 1), GPS_LOG_VERBOSE
            
            For i = LBound(currentTypes) To UBound(currentTypes)
                If current = currentTypes(i).GPSType Then
                    fMain.MessageType = currentTypes(i)
                    Log cModuleName, "setProperties", "Initializing communication settings!", GPS_LOG_VERBOSE
                    fMain.InitializeFromSettings communicationSettings
                    Log cModuleName, "setProperties", "Initialized communication settings!", GPS_LOG_VERBOSE
                    If CBool(getSettings(cManualPoll, cManualPollValue)) Then
                        fMain.pollingBuffer UnscreenMessage(getSettings(cManualPollString, cManualPollStringValue)), getSettings(cManualPollInterval, cManualPollIntervalValue)
                        fMain.LimitMessages CLng(getSettings(cRelayInterval, cRelayIntervalValue)), CLng(getSettings(cDurationInterval, cDurationIntervalValue))
                    Else
                        fMain.pollingBuffer vbNullString, vbNullString
                        fMain.LimitMessages 0, 0
                    End If
                    Process_Status = IIf(fMain.IsPortOpen, GPS_STAT_READYANDWILLING, GPS_STAT_ERROR)
                    Exit For
                End If
            Next
        End If
    Else
        Process_Status = GPS_STAT_UNKNOWN
    End If
    Exit Sub
err_setProperties:
    Log cModuleName, "setProperties", Err.Description, GPS_LOG_ERROR, Err.Number
    Err.Clear
    Resume Next
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
End Function

Private Static Function ByteToHex(bytVal As Byte) As String
    ByteToHex = "00"
    Mid$(ByteToHex, 3 - Len(Hex$(bytVal))) = Hex$(bytVal)
End Function
