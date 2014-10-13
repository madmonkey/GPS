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
Private Const cModuleName = "HTE_Emulator.Process"
Private gApp As HTE_GPS.Processes
Private m_Status As HTE_GPS.GPS_PROCESSOR_STATUS
Private localSettings As MSXML2.DOMDocument30
Private m_Settings As String
Private m_Property As Object
Implements HTE_GPS.Process
Implements HTE_GPS.PropPageCallback

Private WithEvents playbackTimer As ccrpTimers6.ccrpCountdown
Attribute playbackTimer.VB_VarHelpID = -1
Private m_EmulatedType As HTE_GPS.GPS_MESSAGING_TYPES
Private m_playbackType As Long
Private m_playbackInterval As Long
Private m_Loop As Boolean
Private m_Playback As MSXML2.DOMDocument30
Private m_currentNode As Long
Private Const cSourceNodeName = "Source"
Private Const cDateTimeNodeName = "DateTimeStamp"
Private m_ExtendedCharacterSet As Boolean
Private m_SourceConversion As stringBuilder
Private bShutdown As Boolean

Private Sub Class_Initialize()
    m_Status = GPS_STAT_INITIALIZED
    Set localSettings = New MSXML2.DOMDocument30
    Set playbackTimer = New ccrpTimers6.ccrpCountdown
    playbackTimer.Interval = 200 'we don't need millisec resolution - avail intervals are in secs - save some cycles
    Set m_SourceConversion = New stringBuilder
End Sub

Private Sub Class_Terminate()
    m_Status = GPS_STAT_UNINITIALIZED
    Set localSettings = Nothing
    If Not playbackTimer Is Nothing Then
        playbackTimer.Enabled = False
        Set playbackTimer = Nothing
    End If
    Set m_SourceConversion = Nothing
End Sub

Private Sub playbackTimer_Tick(ByVal TimeRemaining As Long)
    'Log cModuleName, "playbackTimer_Tick", "Time remaining = " & CStr(TimeRemaining), GPS_LOG_VERBOSE
End Sub

Private Sub playbackTimer_Timer()
Dim msg As Message
Dim iNode As MSXML2.IXMLDOMNode
    Log cModuleName, "playbackTimer_Timer", "Event fired!", GPS_LOG_VERBOSE
    If Not m_Playback Is Nothing Then
        Set iNode = m_Playback.getElementsByTagName(cSourceNodeName).Item(m_currentNode)
        If Not iNode Is Nothing Then
            Set msg = New Message
            With msg
                .MessageStatus = GPS_MSG_PROCESSED
                .MessageType = m_EmulatedType
                If m_ExtendedCharacterSet Then
                    .rawMessage = HexToBinary(iNode.nodeTypedValue)
                Else
                    .rawMessage = iNode.nodeTypedValue
                End If
            End With
            If Not gApp Is Nothing Then gApp.ProcessComplete msg, Me
            If m_currentNode = m_Playback.getElementsByTagName(cSourceNodeName).Length - 1 Then
                If m_Loop Then
                    Log cModuleName, "playbackTimer_Timer", "Reset from beginning of playback!", GPS_LOG_VERBOSE
                    m_currentNode = 0
                Else
                    Log cModuleName, "playbackTimer_Timer", "End of playback - Loop NOT Selected!", GPS_LOG_VERBOSE
                    Exit Sub
                End If
            Else
                m_currentNode = m_currentNode + 1
            End If
            If m_playbackType = 0 Then
                playbackTimer.Duration = GetRealTimeValue
            End If
            playbackTimer.Enabled = Not bShutdown 'True
        End If
    End If
End Sub

Private Sub Process_Connect(ByVal parentProcess As HTE_GPS.Processes, ByVal cookie As String)
    Set gApp = parentProcess
    gApp.Connected Me, cookie 'WHEN YOU CONNECT CALL CONNECTED METHOD TO RETRIEVE YOUR INSTANCE SETTINGS
End Sub

Private Sub Process_Disconnect()
    Log cModuleName, "Process_Disconnect", "Unhooking process", GPS_LOG_VERBOSE
    bShutdown = True
    If Not playbackTimer Is Nothing Then
        playbackTimer.Enabled = False
        Set playbackTimer = Nothing
    End If
    checkForChanges
    Set m_Property = Nothing
    Set gApp = Nothing
End Sub

Private Property Get Process_FriendlyName() As String
    '*TODO - change name to reflect your project
    Process_FriendlyName = "GPS Emulator Process"
End Property

Private Property Get Process_HostVersion() As String
    Process_HostVersion = "1.0.0"
End Property

Private Sub Process_ProcessMessage(ByVal PropData As HTE_GPS.GPSMessage)
    If Not gApp Is Nothing Then
        Log cModuleName, "ProcessMessage", "Message received from host - forwarding back!", GPS_LOG_VERBOSE, , PropData.rawMessage, GPS_SOURCE_BINARY
        gApp.ProcessComplete PropData, Me
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
    localSettings.async = False
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
'*TODO - initialize and ready your process, be sure to notify hosting application of your status

On Error GoTo err_setProperties
    If Not localSettings Is Nothing Then
        'Use the local getSettings function to retrieve your settings
        m_EmulatedType = getProperty(localSettings, cMessageType, cMessageTypeValue)
        Log cModuleName, "setProperties", "EmulatedType = " & CStr(m_EmulatedType), GPS_LOG_VERBOSE
        '0 - from actual file values, 1 for predetermined interval
        m_playbackType = getProperty(localSettings, cPlaybackIntervalType, cPlaybackIntervalTypeValue)
        Log cModuleName, "setProperties", "PlaybackType = " & CStr(m_playbackType), GPS_LOG_VERBOSE
        m_playbackInterval = getProperty(localSettings, cPlaybackInterval, cPlaybackIntervalValue)
        Log cModuleName, "setProperties", "PlaybackInterval = " & CStr(m_playbackInterval), GPS_LOG_VERBOSE
        m_Loop = CBool(getProperty(localSettings, cLoopbackType, cLoopbackTypeValue))
        Log cModuleName, "setProperties", "Loopback = " & CStr(m_Loop), GPS_LOG_VERBOSE
        m_ExtendedCharacterSet = CBool(getProperty(localSettings, cSourceType, cSourceTypeValue))
        Log cModuleName, "setProperties", "ExtendedCharacterSet = " & CStr(m_ExtendedCharacterSet), GPS_LOG_VERBOSE
        Set m_Playback = New MSXML2.DOMDocument30
        m_Playback.async = False
        Log cModuleName, "setProperties", "Playback", GPS_LOG_VERBOSE ', , getProperty(localSettings, cPlayback, cPlaybackValue), GPS_SOURCE_XML
        If m_Playback.loadXML(getProperty(localSettings, cPlayback, cPlaybackValue)) Then
            'playbackTimer.Interval = 20 '
            If m_playbackInterval = 1 Then
                Log cModuleName, "setProperties", "Playback on specified interval!", GPS_LOG_VERBOSE
                playbackTimer.Duration = m_playbackInterval * 1000
            Else
                Log cModuleName, "setProperties", "Playback from realtime values!", GPS_LOG_VERBOSE
                playbackTimer.Duration = GetRealTimeValue
            End If
            bShutdown = False
            playbackTimer.Enabled = True
            Process_Status = GPS_STAT_READYANDWILLING
        Else
            Process_Status = GPS_STAT_WARNING
        End If
    Else
        Process_Status = GPS_STAT_UNKNOWN
    End If
    Exit Sub
err_setProperties:
    Log cModuleName, "setProperties", Err.Description, GPS_LOG_ERROR, Err.Number
End Sub

Private Function GetRealTimeValue() As Long
Dim lValue As Long
Dim iNodes As MSXML2.IXMLDOMNodeList
Dim iNode As MSXML2.IXMLDOMElement
Dim iNode2 As MSXML2.IXMLDOMElement

On Local Error Resume Next
    lValue = m_playbackInterval * 1000
    If Not m_Playback Is Nothing Then
        If m_Playback.documentElement.hasChildNodes Then
            If (m_Playback.documentElement.childNodes.Length - 1) >= 1 Then
                Set iNodes = m_Playback.documentElement.getElementsByTagName("ROW")
                If Not iNodes Is Nothing Then
                    Set iNode = iNodes.Item(m_currentNode).selectSingleNode(cDateTimeNodeName)
                    If Not iNode Is Nothing Then
                        If m_currentNode > 0 And m_currentNode <= iNodes.Length - 1 Then
                            Set iNode2 = iNodes.Item(m_currentNode - 1).selectSingleNode(cDateTimeNodeName)
                            If Not iNode2 Is Nothing Then
                                If IsDate(iNode.nodeTypedValue) And IsDate(iNode2.nodeTypedValue) Then
                                    lValue = Abs(DateDiff("s", TimeValue(iNode.nodeTypedValue), TimeValue(iNode2.nodeTypedValue)))
                                    If lValue = 0 Then lValue = 1
                                    Log cModuleName, "GetRealTimeValue", "Calculated realtime value = " & lValue & " second(s).", GPS_LOG_VERBOSE
                                    lValue = lValue * 1000
                                End If
                            End If
                        Else
                            'end-of file no time to compare to give default
                            Log cModuleName, "GetRealTimeValue", "No node to calculate against - using default interval property " & CStr(m_playbackInterval) & " second(s)!", GPS_LOG_VERBOSE
                        End If
                    End If
                End If
            Else
                Log cModuleName, "GetRealTimeValue", "Only one NODE nothing to calculate against - using default interval property " & CStr(m_playbackInterval) & " second(s)!", GPS_LOG_VERBOSE
            End If
        End If
    End If
    GetRealTimeValue = lValue
End Function

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

Private Function HexToBinary(ByRef HexDump As String) As String
Dim sArray() As String
Dim i As Long
On Error GoTo err_HexToBinary
    m_SourceConversion.Remove 0, m_SourceConversion.Length
    sArray = Split(HexDump, Space$(1))
    If IsValidArray(sArray) Then
        For i = LBound(sArray) To UBound(sArray)
            m_SourceConversion.Append Chr$("&H" & Left$(sArray(i), 2))
        Next
    End If
    HexToBinary = m_SourceConversion.ToString
    Exit Function
err_HexToBinary:
    Log cModuleName, "HexToBinary", "Error decoding source from extended characterset.", GPS_LOG_ERROR, Err.Number
End Function
