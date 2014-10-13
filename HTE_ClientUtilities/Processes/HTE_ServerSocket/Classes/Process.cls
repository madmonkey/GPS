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

Private Const cModuleName = "HTE_ServerSocket.Process"
Private gApp As HTE_GPS.Processes
Private m_Status As HTE_GPS.GPS_PROCESSOR_STATUS
Private localSettings As MSXML2.DOMDocument30
Private m_Settings As String
Private m_Property As Object
Private m_Port  As Long
Implements HTE_GPS.Process
Implements HTE_GPS.PropPageCallback
Private gpsBuffers As Scripting.Dictionary
Private WithEvents fListener As SimpleSockets.Host
Attribute fListener.VB_VarHelpID = -1
Private WithEvents AliasLookup As HTE_Entity.Alias
Attribute AliasLookup.VB_VarHelpID = -1
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub AliasLookup_LookupFailed(cMsg As HTE_GPS.GPSMessage)
    Log cModuleName, "AliasLookup_LookupFailed", "Unable to resolve endpoint according to settings!", GPS_LOG_VERBOSE, , cMsg.Entity, GPS_SOURCE_XML
    If Not gApp Is Nothing Then
        cMsg.MessageStatus = GPS_MSG_PROCESSED_WARNING
        If Not gApp Is Nothing Then gApp.ProcessComplete cMsg, Me
    End If
End Sub

Private Sub AliasLookup_LookupResolved(cMsg As HTE_GPS.GPSMessage)
    Log cModuleName, "AliasLookup_LookupResolved", "Resolved endpoint according to settings!", GPS_LOG_VERBOSE, , cMsg.Entity, GPS_SOURCE_XML
    If Not gApp Is Nothing Then
        If Not gApp Is Nothing Then gApp.ProcessComplete cMsg, Me
    End If
End Sub

Private Sub Class_Initialize()
Dim bReturn As Boolean
    Set gpsBuffers = New Scripting.Dictionary
    Set AliasLookup = New HTE_Entity.Alias
    m_Status = GPS_STAT_INITIALIZED
    Set localSettings = New MSXML2.DOMDocument30
End Sub

Private Sub Class_Terminate()
    gpsBuffers.removeAll
    UnreadySocket
    Set gpsBuffers = Nothing
    Set AliasLookup = Nothing
    m_Status = GPS_STAT_UNINITIALIZED
    Set localSettings = Nothing
End Sub

Private Sub fListener_ClientAccepted(c As SimpleSockets.ConnectedClient, keepListening As Boolean)
    keepListening = True
    Process_Status = GPS_STAT_READYANDWILLING
    AddSocket c.address
    Log cModuleName, "ProcessMessage", "FD_ACCEPT - " & c.ToString & " connected.", GPS_LOG_VERBOSE
End Sub

Private Sub fListener_ClientDataReceived(c As SimpleSockets.IClient, data() As Byte, bytesRead As Long)
Dim lRet As Long, sTemp As stringBuilder, bytesRecv As Long, strData As String
    Process_Status = GPS_STAT_READYANDWILLING
    If gpsBuffers.Exists(c.address) Then
        strData = StringFromByteArray(data)
        Set sTemp = gpsBuffers.Item(c.address)
        sTemp.Append strData
        '!very important to tell engine how many bytes you've read in!
        bytesRead = Len(strData)
        Log cModuleName, "ProcessMessage", "FD_READ - " & c.ToString & " [" & sTemp.Length & "] bytes pending [" & bytesRead & "] bytes read.", GPS_LOG_VERBOSE, 0, sTemp.ToString, GPS_SOURCE_BINARY
        If sTemp.Length > 0 Then
            lRet = PrepareAndSendMessages(sTemp, c.address)
            Log cModuleName, "ProcessMessage", "Byte(s) remaining [" & sTemp.Length & "]", GPS_LOG_VERBOSE, , gpsBuffers.Item(c.address).ToString, GPS_SOURCE_BINARY
            Log cModuleName, "ProcessMessage", CStr(lRet) & " message" & IIf(lRet <> 1, "s", "") & " parsed.", GPS_LOG_INFORMATION
        Else
            Log cModuleName, "ProcessMessage", "Nothing to process - ignoring.", GPS_LOG_VERBOSE
        End If
    Else
        Log cModuleName, "ProcessMessage", "FD_READ - " & c.ToString & " is not in accepted list. Disconnecting...", GPS_LOG_INFORMATION
        RemoveSocket c.address
        c.Disconnect
    End If
End Sub

Private Sub fListener_ClientDisconnected(c As SimpleSockets.ConnectedClient)
    Process_Status = GPS_STAT_READYANDWILLING
    Log cModuleName, "ProcessMessage", "FD_CLOSE - " & c.ToString & " disconnected.", GPS_LOG_VERBOSE
    RemoveSocket c.address
End Sub

Private Sub Process_Connect(ByVal parentProcess As HTE_GPS.Processes, ByVal cookie As String)
    ReadySocket
    Set gApp = parentProcess
    If Not gApp Is Nothing Then gApp.Connected Me, cookie
End Sub

Private Sub ReadySocket()
Dim bReturn As Boolean
    On Local Error Resume Next
    If fListener Is Nothing Then Set fListener = New SimpleSockets.Host
    bReturn = Not (fListener Is Nothing)
    Log cModuleName, "ReadySocket", "Started Winsock = " & CStr(bReturn), IIf(bReturn, GPS_LOG_VERBOSE, GPS_LOG_ERROR)
End Sub
Private Sub UnreadySocket()
    If Not fListener Is Nothing Then
        With fListener
            If .IsListening Then .StopListening
            .ShutDown
            .Dispose
        End With
        Set fListener = Nothing
    End If
End Sub
Private Sub Process_Disconnect()
    Log cModuleName, "Process_Disconnect", "Unhooking process", GPS_LOG_VERBOSE
    UnreadySocket
    checkForChanges
    Set m_Property = Nothing
    Set gApp = Nothing
End Sub

Private Property Get Process_FriendlyName() As String
    Process_FriendlyName = "Server Socket Process"
End Property

Private Property Get Process_HostVersion() As String
    Process_HostVersion = "1.0.0"
End Property

Private Sub Process_ProcessMessage(ByVal PropData As HTE_GPS.GPSMessage)
Dim keys As Variant, i As Long
Dim iC As SimpleSockets.IClient
Dim msgTags As HTE_GPS.GPSConfiguration
    If Len(PropData.rawMessage) > 0 Then
        If Not gpsBuffers Is Nothing Then
            msgTags = getMessageType(PropData.Type)
            keys = gpsBuffers.keys 'if nothing added will be empty return - the ubound will be -1, instead of 0
            If Not IsEmpty(keys) Then
                Log cModuleName, "Process_ProcessMessage", "[" & UBound(keys) + 1 & "] clients set to receive message!", GPS_LOG_VERBOSE
                For i = 0 To UBound(keys)
                    Set iC = fListener.GetClientByIpAddress(CStr(keys(i)))
                    If iC.SendData(ByteArrayFromString(msgTags.SOM & PropData.rawMessage & msgTags.EOM)) Then
                        Log cModuleName, "Process_ProcessMessage", "Message received by client at address: " & iC.address & ".", GPS_LOG_INFORMATION
                        Process_Status = GPS_STAT_READYANDWILLING
                    Else
                        Log cModuleName, "Process_ProcessMessage", "Nothing received by client at address: " & iC.address & ". Invalid socket - disconnecting!", GPS_LOG_WARNING
                        RemoveSocket iC.address
                        iC.Disconnect
                    End If
                Next
            Else
                Log cModuleName, "Process_ProcessMessage", "Key list is empty!", GPS_LOG_INFORMATION
            End If
        Else
            Log cModuleName, "Process_ProcessMessage", "Buffer list is empty!", GPS_LOG_INFORMATION
        End If
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
                If Not gApp Is Nothing Then gApp.PropertiesChanged retrieveLocalSettings, Me
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

Dim bReinitialize As Boolean
Dim attempt As Long
Const MAX_ATTEMPTS As Long = 20
On Error GoTo err_setProperties
    If Not localSettings Is Nothing Then
        'Use the local getSettings function to retrieve your settings
        If Not gApp Is Nothing Then CurrentTypes SupportedMessages(localSettings)
        InitializeAlias
        Process_Status = GPS_STAT_ERROR
        Log cModuleName, "setProperties", "Port = [" & m_Port & "] Settings = [" & getSettings(cPortNumber, cPortNumberValue) & "]", GPS_LOG_VERBOSE
        m_Port = getSettings(cPortNumber, cPortNumberValue)
        DisconnectAll
        UnreadySocket
        ReadySocket
        If Not fListener Is Nothing Then
            With fListener
                .CaptureClientEvents = True
                .BindPort = m_Port
                .Listen
                Process_Status = IIf(.IsListening, GPS_STAT_READYANDWILLING, GPS_STAT_WARNING)
            End With
        Else
            Log cModuleName, "setProperties", "Message hWnd not instanciated!", GPS_LOG_ERROR
        End If
    Else
        Process_Status = GPS_STAT_UNKNOWN
    End If
    Exit Sub
err_setProperties:
    Log cModuleName, "setProperties", Err.Description, GPS_LOG_ERROR, Err.Number
End Sub
Private Sub InitializeAlias()
    If AliasLookup Is Nothing Then Set AliasLookup = New HTE_Entity.Alias
    With AliasLookup
        Log cModuleName, "setProperties", "Enable Mac Resolution = " & CStr(CBool(getSettings(cEnableMacResolution, cEnableMacResolutionValue))), GPS_LOG_VERBOSE
        .EnableMacResolution = CBool(getSettings(cEnableMacResolution, cEnableMacResolutionValue))
        Log cModuleName, "setProperties", "Cache Lookup for [" & Val(getSettings(cCacheMacLookupSeconds, cCacheMacLookupSecondsValue)) & "] second(s).", GPS_LOG_VERBOSE
        .MacLookupExpiration = CLng(Val(getSettings(cCacheMacLookupSeconds, cCacheMacLookupSecondsValue)))
    End With
    Process_Status = GPS_STAT_READYANDWILLING
End Sub

Private Function DisconnectAll()
    Log cModuleName, "DisconnectAll", "Removing any current connections from host socket", GPS_LOG_VERBOSE
    If Not fListener Is Nothing Then fListener.CloseAllConnections
    Log cModuleName, "DisconnectAll", "Flushing network buffers", GPS_LOG_VERBOSE
    If Not gpsBuffers Is Nothing Then gpsBuffers.removeAll
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

Private Function PrepareAndSendMessages(ByRef buffer As stringBuilder, ByVal sIPAddress As String) As Long
Dim msg As Message
Dim numMessages As Long
Dim sRaw As String
Dim myType As HTE_GPS.GPS_MESSAGING_TYPES
'Added to handle potentially multiple messages
On Error GoTo err_PrepareAndSendMessages
    numMessages = 0
    If Not App Is Nothing Then 'don't bother if you can't callback
        myType = InferType(buffer.ToString)
        sRaw = ParseGPSMessage(myType, buffer)
        Do While sRaw <> vbNullString
            Log cModuleName, "PrepareAndSendMessage", "Parsed Message!", , , sRaw, GPS_SOURCE_BINARY
            Set msg = New Message
            numMessages = numMessages + 1
            msg.MessageType = myType
            msg.rawMessage = sRaw
            msg.MessageStatus = GPS_MSG_PROCESSED
            If Not AliasLookup Is Nothing Then
                AliasLookup.resolveEntity msg, sIPAddress
            End If
            myType = InferType(buffer.ToString)
            sRaw = ParseGPSMessage(myType, buffer)
        Loop
        Log cModuleName, "PrepareAndSendMessage", "No " & IIf(numMessages > 0, "more", vbNullString) & " messages found in buffer.", GPS_LOG_VERBOSE
    End If

    PrepareAndSendMessages = numMessages
    Exit Function

err_PrepareAndSendMessages:
    Log cModuleName, "PrepareAndSendMessages", Err.Description, GPS_LOG_ERROR, Err.Number
    PrepareAndSendMessages = numMessages
End Function

Private Sub AddSocket(ByVal FromIP As String)
On Local Error Resume Next
Dim gpsBuffer As stringBuilder
    If Not gpsBuffers.Exists(FromIP) Then
        Set gpsBuffer = New stringBuilder
        gpsBuffer.ChunkSize = 1024 '1k start per endpoint
        gpsBuffers.Add FromIP, gpsBuffer
    End If
    Log cModuleName, "AddSocket", "[Current connections: " & fListener.GetClientCount & "]", GPS_LOG_INFORMATION
End Sub

Private Sub RemoveSocket(ByVal dicKey As String)
On Local Error Resume Next
    Log cModuleName, "RemoveSocket", "Removing client [" & dicKey & "]", GPS_LOG_VERBOSE
    If gpsBuffers.Exists(dicKey) Then gpsBuffers.Remove (dicKey)
    Log cModuleName, "RemoveSocket", "[Current connections: " & fListener.GetClientCount & "]", GPS_LOG_INFORMATION
End Sub

Private Function ByteArrayFromString(ByVal Source$) As Variant
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

Private Function ParseGPSMessage(ByVal msgType As HTE_GPS.GPS_MESSAGING_TYPES, ByRef buffer As stringBuilder) As String
Dim startPos As Long, endPos As Long
Dim GPSConfig As HTE_GPS.GPSConfiguration

    With buffer
        If .Length > 0 Then
            GPSConfig = getMessageType(msgType)
            Log cModuleName, "ParseGPSMessage", "GPSConfig.Desc = " & GPSConfig.Desc & "; GPSConfig.SOM = " & GPSConfig.SOM & "; GPSConfig.EOM = " & GPSConfig.EOM & "; GPSConfig.GPSType = " & GPSConfig.GPSType, GPS_LOG_VERBOSE
            If Len(GPSConfig.SOM) = 0 And Len(GPSConfig.EOM) = 0 Then 'for messages with no start/end
                ParseGPSMessage = .ToString
                .Remove 0, .Length
            Else
                startPos = .Find(GPSConfig.SOM)
                If startPos > 0 Then
                    If startPos > 1 Then
                        .Remove 0, startPos - 1
                        startPos = 1
                    End If
                    endPos = .Find(GPSConfig.EOM, startPos + 1)
                    If endPos > startPos Then
                        ParseGPSMessage = Mid$(.ToString, startPos + Len(GPSConfig.SOM), (endPos - (startPos + Len(GPSConfig.SOM))))    'Mid$(.ToString, startPos + Len(GPSConfig.SOM), (endPos - (startPos + Len(GPSConfig.SOM))))
                        .Remove startPos - 1, endPos
                    End If
                End If
            End If
        End If
    End With
End Function
