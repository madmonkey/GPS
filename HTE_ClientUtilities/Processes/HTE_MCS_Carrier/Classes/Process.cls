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

Private Const cModuleName = "HTE_MCS_CARRIER.Process"
Private gApp As HTE_GPS.Processes
Private m_Status As HTE_GPS.GPS_PROCESSOR_STATUS
Private localSettings As MSXML2.DOMDocument30
Private m_Settings As String
Private m_Property As Object

Implements HTE_GPS.PropPageCallback
Implements HTE_GPS.Process

Private WithEvents MCS As HTE_MCS.Transport 'UCS_MCS.MCSClient
Attribute MCS.VB_VarHelpID = -1
Private m_CAD As String
Private m_SesType As UCS_MCS.COMM_MSGTYPES
Private m_DestID As String
Private m_Session As String
Private m_Unit As String
Private m_MsgKey As String
Private m_MsgSend As String
Private m_MsgRecv() As String
Private m_TagRecv() As String
Private bOverrideUserID As Boolean
Private WithEvents currentSettings As cSettings
Attribute currentSettings.VB_VarHelpID = -1

Private Const SOH = &H1
Private Const STX = &H2
Private Const ETX = &H3
Private Const EOT = &H4
Private Const FS = &H1C

'ADDED FOR SOFTWARE FLOWCONTROL - MOSTLY TO CIRCUMVENT REPORTING TO CAD MORE FREQUENT THAN LOCAL
'ALSO TO POTENTIALLY SEND LOGIN MESSAGES AT PREDETERMINED INTERVALS
Private m_SecondsPerTransmission As Long 'send messages every x seconds
Private m_SecondsForTransmission As Long 'send messages for y seconds
Private m_Notify As ccrpTimers6.ccrpCountdown
Implements ccrpTimers6.ICcrpCountdownNotify
Private bTimeToWork As Boolean
Private dLastTransmission As Double
Private m_PollingInfo As Collection
Private WithEvents ManualTime As ccrpTimers6.ccrpTimer
Attribute ManualTime.VB_VarHelpID = -1
Private m_messageType As HTE_GPS.GPS_MESSAGING_TYPES

Private Sub Class_Initialize()
    m_Status = GPS_STAT_INITIALIZED
    Set localSettings = New MSXML2.DOMDocument30
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    'RESUME NEXT BECAUSE MCS MIGHT THROW ERROR WE WISH TO CONTINUE
    m_Status = GPS_STAT_UNINITIALIZED
    Set localSettings = Nothing
    If Not MCS Is Nothing Then
        MCS.Terminate
    End If
    Set MCS = Nothing
End Sub

Private Sub currentSettings_UserChanged(ByVal HTEUserID As String)
    If Not bOverrideUserID Then
        Log cModuleName, "currentSettings_UserChanged", "Changing user from " & m_Unit & " to " & HTEUserID, GPS_LOG_VERBOSE
        m_Unit = HTEUserID
    Else
        Log cModuleName, "currentSettings_UserChanged", "Overriden UserID NOT changing " & m_Unit & " to " & HTEUserID, GPS_LOG_VERBOSE
    End If
End Sub

Private Sub MCS_CommDeliveryFailure(ByVal MsgNumber As Long, ByVal EventType As UCS_MCS.COMM_EVENTS, ByVal Message As String)
    Log cModuleName, "MCS_CommDeliveryFailure", "Failure to deliver MCS message " & MsgNumber, GPS_LOG_WARNING, , Message, GPS_SOURCE_BINARY
End Sub

Private Sub MCS_CommEvent(ByVal EventType As UCS_MCS.COMM_EVENTS, ByVal EventDescription As String)
    Log cModuleName, "MCS_CommEvent", "COMM_EVENT - " & EventDescription, GPS_LOG_VERBOSE
    Select Case EventType
        Case COMM_HOST_DOWN, COMM_DATA_SVC_NOTAVAIL, COMM_HW_ERROR, COMM_NOT_AVAILABLE, _
                COMM_NOT_REGISTERED, COMM_SETUP_FAILED, COMM_SW_ERROR, COMM_TX_DISABLED, _
                COMM_XFAIL, COMM_NO_RPM_SIG, COMM_SELF_TEST_FAIL, COMM_SETUP_START
            Log cModuleName, "MCS_ReceivedData", "Critical event recognized as " & EventDescription, GPS_LOG_ERROR
            Process_Status = GPS_STAT_ERROR
        Case COMM_LOW_BATTERY, COMM_MODEM_SILENT, COMM_NO_RESPONSE, COMM_OUT_OF_RANGE, _
                COMM_TERM_FAILED, COMM_BUSY, COMM_NO_ACK, COMM_SETUP_COMPLETE
            Process_Status = GPS_STAT_WARNING
        Case Else
            Process_Status = GPS_STAT_READYANDWILLING
    End Select
End Sub

Private Sub MCS_CommMessage(ByVal Message As String)
    Log cModuleName, "MCS_CommMessage", "DATA_ARRIVAL", GPS_LOG_INFORMATION
    Process_Status = GPS_STAT_READYANDWILLING 'IF WE'VE RECV'D THEN WE'RE GETTING SOMETHING
    ProcessMCSMessage Message
    Log cModuleName, "MCS_CommMessage", "DATA_PROCESSED", GPS_LOG_INFORMATION
End Sub

Private Sub Process_Connect(ByVal parentProcess As HTE_GPS.Processes, ByVal cookie As String)
    Set MCS = New HTE_MCS.Transport 'UCS_MCS.MCSClient 'done on the connect instead of initialize because 'flickers' in systray when testing for good process interface
    Set currentSettings = New cSettings
    Set gApp = parentProcess
    gApp.Connected Me, cookie
End Sub

Private Sub Process_Disconnect()
On Error Resume Next
'RESUME NEXT BECAUSE MCS MIGHT THROW ERROR WE WISH TO CONTINUE
    Log cModuleName, "Process_Disconnect", "Unhooking process", GPS_LOG_VERBOSE
    checkForChanges
    If Not MCS Is Nothing Then
        Log cModuleName, "Process_Disconnect", "MCS is NOT Nothing - Terminate!", GPS_LOG_VERBOSE
        MCS.Terminate
        Log cModuleName, "Process_Disconnect", "MCS Terminated! Error Code = " & Err.Number, GPS_LOG_VERBOSE
    Else
        Log cModuleName, "Process_Disconnect", "MCS is Nothing - No need to Terminate!", GPS_LOG_VERBOSE
    End If
    Set MCS = Nothing
    Log cModuleName, "Process_Disconnect", "MCS object released!", GPS_LOG_VERBOSE
    Set m_Property = Nothing
    Log cModuleName, "Process_Disconnect", "m_Property released!", GPS_LOG_VERBOSE
    Log cModuleName, "Process_Disconnect", "currentSettings Terminate!", GPS_LOG_VERBOSE
    currentSettings.Terminate
    Set currentSettings = Nothing
    Log cModuleName, "Process_Disconnect", "currentSettings released!", GPS_LOG_VERBOSE
    Log cModuleName, "Process_Disconnect", "Application callback released!", GPS_LOG_VERBOSE
    Set gApp = Nothing
End Sub

Private Property Get Process_FriendlyName() As String
    Process_FriendlyName = "Client MCS Process"
End Property

Private Property Get Process_HostVersion() As String
    Process_HostVersion = "1.0.0"
End Property

Private Sub Process_ProcessMessage(ByVal PropData As HTE_GPS.GPSMessage)
    Log cModuleName, "ProcessMessage", "Received data.", GPS_LOG_VERBOSE, , PropData.rawMessage, GPS_SOURCE_BINARY
    BuildMessage PropData
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
Dim sDefault As String

On Error GoTo err_setProperties
    If Not localSettings Is Nothing Then
        m_DestID = getSettings(cDest, cDestValue)
        m_Session = getSettings(cSession, cSessionValue)
        m_MsgSend = getSettings(cMsgToSend, cMsgToSendValue)
        m_MsgRecv = Split(getSettings(cMsgsToRecv, cMsgsToRecvValue), cSep)
        m_TagRecv = Split(getSettings(cTagToRead, cTagToReadValue), cSep)
        CurrentTypes SupportedMessages(localSettings)
        With currentSettings
            .StopMonitorForChanges
            .SessionName = m_Session
            sDefault = IIf(StrComp(.GetValue(ikSessionType), "LECS", vbTextCompare) = 0, "CAD400", "CADV")
            m_CAD = getSettings(cCADType, sDefault)
            sDefault = IIf(StrComp(.GetValue(ikSessionType), "LECS", vbTextCompare) = 0, COMM_TFF, COMM_MESSAGE)
            m_SesType = IIf(.GetValue(ikSessionType) = "LECS", COMM_TFF, COMM_MESSAGE)
            Log cModuleName, "getProperty", "Property: Session Type = " & IIf(.GetValue(ikSessionType) = "LECS", "COMM_TFF", "COMM_MESSAGE"), GPS_LOG_VERBOSE
            m_MsgKey = Replace(.GetValue(ikSessionINIKey), "????,", vbNullString)
            bOverrideUserID = StrComp(getSettings(cAppIDTag, vbNullString), vbNullString) <> 0
            If Not bOverrideUserID Then
                m_Unit = .GetValue(ikNewUser)
                Log cModuleName, "getProperty", "Property: TerminalUnit = " & m_Unit
                Log cModuleName, "setProperties", m_CAD & " client, will monitor for Unit changes.", GPS_LOG_VERBOSE
                .MonitorForChanges
            Else
                m_Unit = getSettings(cAppIDTag, .GetValue(ikNewUser))
                Log cModuleName, "getProperty", "Property: TerminalUnit(OVERRIDE) = " & m_Unit
                Log cModuleName, "setProperties", m_CAD & " client, will NOT monitor for Unit changes - override in effect.", GPS_LOG_VERBOSE
            End If
            'FLOW CONTROL FOR MCS
            If CBool(getSettings(cManualPoll, cManualPollValue)) Then
                pollingBuffer UnscreenMessage(getSettings(cManualPollString, cManualPollStringValue)), getSettings(cManualPollInterval, cManualPollIntervalValue)
                LimitMessages CLng(getSettings(cRelayInterval, cRelayIntervalValue)), CLng(getSettings(cDurationInterval, cDurationIntervalValue))
            Else
                pollingBuffer vbNullString, vbNullString
                LimitMessages 0, 0
            End If
            'FLOW CONTROL FOR MCS
        End With
        If m_Session <> vbNullString Then
            If MCS.Session <> m_Session Then MCS.Terminate
            On Error Resume Next
            MCS.Initialize m_Session
            If Err.Number = 0 Then
                Process_Status = GPS_STAT_READYANDWILLING
            Else
                Log cModuleName, "setProperties", "Error initializing MCS session!", GPS_LOG_ERROR, Err.Number, Err.Source
            End If
        Else
            Log cModuleName, "setProperties", "Empty Session name found in configuration file...will be unable to send message!", GPS_LOG_WARNING
        End If
    Else
        Process_Status = GPS_STAT_UNKNOWN
    End If
    Exit Sub
err_setProperties:
    Log cModuleName, "setProperties", Err.Description, GPS_LOG_ERROR, Err.Number
    Resume Next
End Sub

Private Function EventDescription(ByVal UCS_Event As UCS_MCS.COMM_EVENTS) As String
Dim Ret$
    Select Case UCS_Event
        Case COMM_SUCCESS
            Ret = "COMM_SUCCESS"
        Case COMM_TRANSMITTING
            Ret = "COMM_TRANSMITTING"
        Case COMM_XFAIL
            Ret = "COMM_XFAIL"
        Case COMM_NO_ACK
            Ret = "COMM_NO_ACK"
        Case COMM_LOW_BATTERY
            Ret = "COMM_LOW_BATTERY"
        Case COMM_SETUP_START
            Ret = "COMM_SETUP_START"
        Case COMM_TERM_START
            Ret = "COMM_TERM_START"
        Case COMM_TERM_COMPLETE
            Ret = "COMM_TERM_COMPLETE"
        Case COMM_HOST_DOWN
            Ret = "COMM_HOST_DOWN"
        Case COMM_MODEM_SILENT
            Ret = "COMM_MODEM_SILENT"
        Case COMM_TERM_FAILED
            Ret = "COMM_TERM_FAILED"
        Case COMM_SETUP_FAILED
            Ret = "COMM_SETUP_FAILED"
        Case COMM_NOT_REGISTERED
            Ret = "COMM_NOT_REGISTERED"
        Case COMM_BUSY
            Ret = "COMM_BUSY"
        Case COMM_NOT_AVAILABLE
            Ret = "COMM_NOT_AVAILABLE"
        Case COMM_HW_ERROR
            Ret = "COMM_HW_ERROR"
        Case COMM_OUT_OF_RANGE
            Ret = "COMM_OUT_OF_RANGE"
        Case COMM_RX_IN_RANGE
            Ret = "COMM_RX_IN_RANGE"
        Case COMM_RPM_SIG_OK
            Ret = "COMM_RPM_SIG_OK"
        Case COMM_SETUP_COMPLETE
            Ret = "COMM_SETUP_COMPLETE"
        Case COMM_NO_RPM_SIG
            Ret = "COMM_NO_RPM_SIG"
        Case COMM_TOO_LONG
            Ret = "COMM_TOO_LONG"
        Case Else
            Ret = "UNKNOWN_COMM_EVENT"
    End Select
    EventDescription = Ret
End Function

Private Function BuildMessage(ByRef PropData As HTE_GPS.GPSMessage, Optional ByRef autoGenerated As Boolean = False) As String
Dim msgBuffer As stringBuilder
Dim msgLength As Long
Dim msgType As HTE_GPS.GPSConfiguration
    On Error GoTo errBuildMessage

    'Build message buffer
    msgType = getMessageType(PropData.Type)
    If (msgType.EOM <> vbNullString And msgType.SOM <> vbNullString And msgType.Desc <> vbNullString) Or autoGenerated Then
        Set msgBuffer = New stringBuilder
        If m_CAD <> "CAD400" Then
            msgBuffer.Append Chr(SOH) & String(8, 0) 'SOH, MN, PMN
            msgBuffer.Append Left$(m_Unit & Space$(8), 8)          'Source ID
            msgBuffer.Append Left$(m_DestID & Space$(8), 8)     'Dest ID
            msgBuffer.Append Chr(1)                                        'Message Priority
            msgBuffer.Append Chr(0) & Chr(0) & m_MsgSend & Chr(STX)           'Flags, Trans ID, STX
            If Not autoGenerated Then
                msgBuffer.Append msgType.Desc & Chr(FS) & msgType.SOM & PropData.rawMessage & msgType.EOM & Chr(FS)   'Data
            Else
                msgBuffer.Append PropData.rawMessage & Chr(FS)
            End If
            msgBuffer.Append Chr(ETX) & Chr(EOT)                           'ETX, EOT
            Log cModuleName, "BuildMessage", "Built position message for " & m_CAD & ":", GPS_LOG_VERBOSE, , msgBuffer.ToString, GPS_SOURCE_BINARY
        Else
            msgBuffer.Append Chr(SOH) & m_MsgSend & Chr(FS) & Left$(m_Unit & Space$(8), 8) & Chr(STX)
            If Not autoGenerated Then
                msgBuffer.Append msgType.Desc & Chr(FS) & msgType.SOM & PropData.rawMessage & msgType.EOM & Chr(FS)
            Else
                msgBuffer.Append PropData.rawMessage & Chr(FS)
            End If
            msgBuffer.Append Chr(ETX)
            Log cModuleName, "BuildMessage", "Built position message for " & m_CAD & ":", GPS_LOG_VERBOSE, , msgBuffer.ToString, GPS_SOURCE_BINARY
        End If
        If Not MCS Is Nothing Then
            MCS.send msgBuffer.ToString
            Log cModuleName, "BuildMessage", "Message was sent!", GPS_LOG_INFORMATION
            Process_Status = GPS_STAT_READYANDWILLING 'If we sent we are good!
        Else
            Log cModuleName, "BuildMessage", "MCS Client is nothing!!", GPS_LOG_WARNING
        End If
        Set msgBuffer = Nothing
    Else
        Log cModuleName, "BuildMessage", "Unable to retrieve settings for GPS Message type!", GPS_LOG_WARNING
    End If
    Exit Function

errBuildMessage:
    Log cModuleName, "BuildMessage", Err.Number & " - " & Err.Description, GPS_LOG_ERROR

End Function

Private Function ProcessMCSMessage(ByVal buffer As String)
    Dim Tokens As New Collection
    Dim Trans As String
    Dim msg As Message 'HTE_GPS.GPSMessage
    Dim i As Long, x As Long
    Dim bFound As Boolean
    
    bFound = False
    Log cModuleName, "ProcessMCSMessage", "Recv'd buffer!", GPS_LOG_VERBOSE, , buffer, GPS_SOURCE_BINARY
    Set msg = New Message
    msg.MessageStatus = GPS_MSG_ERROR
    If IsValidArray(m_MsgRecv) Then
        Trans = GetTransactionName(buffer, Tokens)
        If Trans <> vbNullString Then
            For i = LBound(m_MsgRecv) To UBound(m_MsgRecv)
                If StrComp(Trans, m_MsgRecv(i), vbTextCompare) = 0 Then
                    Log cModuleName, "ProcessMCSMessage", "Transaction found as " & Trans, GPS_LOG_VERBOSE
                    bFound = True
                    If IsValidArray(m_TagRecv) Then
                        'Message coming back is assumed to be first valid format unless otherwise specified in MCS message
                        msg.MessageType = GPS_TYPE_0
                        If Exists(Tokens, cMsgTypeTag) Then
                            msg.MessageType = CLng(Tokens(cMsgTypeTag))
                        End If
                        If i <= UBound(m_TagRecv) Then
                            If Exists(Tokens, m_TagRecv(i)) Then
                                msg.rawMessage = Chr(STX) & Tokens.Item(m_TagRecv(i)) & Chr(ETX)
                                msg.MessageStatus = GPS_MSG_PROCESSED
                                Exit For
                            Else
                                Log cModuleName, "ProcessMCSMessage", "Field token NOT found in transaction!", GPS_LOG_WARNING
                            End If
                        Else
                            Log cModuleName, "ProcessMCSMessage", "Field array NOT sync'd with transaction array...check configuration!", GPS_LOG_WARNING
                        End If
                    Else
                        Log cModuleName, "ProcessMCSMessage", "Fields array NOT a valid array...check configuration!", GPS_LOG_WARNING
                    End If
                End If
            Next
            If Not bFound Then
                'Don't flag as warning if not valid transaction
                Log cModuleName, "ProcessMCSMessage", "Transaction not found as one to respond to...ignoring!", GPS_LOG_VERBOSE
            Else
                'Only raise process complete if it is a valid configurable transaction...otherwise drop it!!!
                gApp.ProcessComplete msg, Me
            End If
            'If we are configured ONLY to receive, we should still update process state...
            'Process_Status = IIf(bFound, GPS_STAT_READYANDWILLING, GPS_STAT_WARNING)
        Else
            Log cModuleName, "ProcessMCSMessage", "No transaction name found!", GPS_LOG_WARNING
        End If
    Else
        Log cModuleName, "ProcessMCSMessage", "Transaction array NOT a valid array...check configuration!", GPS_LOG_WARNING
    End If
    
End Function
Private Function GetTransactionName(ByVal buffer As String, Optional ByRef myTokens As Collection) As String

    Dim vbSTX As String: vbSTX = Chr(2)
    Dim vbSOH As String: vbSOH = Chr(1)
    Dim vbFS As String: vbFS = Chr(&H1C)
    Dim Trans As String, myArray As Variant
    Dim x As Long

    On Error GoTo Err_GetTransName
        Select Case m_CAD
            Case "CAD400"
                If Len(buffer) >= 2 Then
                    If InStr(1, buffer, vbSOH) > 0 Then
                        buffer = Mid$(buffer, InStr(1, buffer, vbSOH), Len(buffer))
                        buffer = Mid$(buffer, 2, Len(buffer))
                        Trans = Mid$(buffer, 1, InStr(1, buffer, vbFS) - Len(vbFS))
                        buffer = Mid$(buffer, InStr(1, buffer, vbSTX) + Len(vbSTX))
                        buffer = Left$(buffer, InStr(1, buffer, Chr(3)) - Len(Chr(3)))
                    Else
                        Log cModuleName, "GetTransactionName", "Start of header not found.", GPS_LOG_WARNING, , buffer, GPS_SOURCE_BINARY
                    End If
                End If
            Case Else
                If Len(buffer) >= 29 Then
                    buffer = Mid$(buffer, 29)
                    Trans = Mid$(buffer, 1, InStr(1, buffer, vbSTX) - Len(vbSTX))
                End If
        End Select
        Log cModuleName, "GetTransactionName", "Transaction recognized as " & Trans, GPS_LOG_VERBOSE
        If Not IsMissing(myTokens) And Not myTokens Is Nothing Then
            myArray = Split(Mid$(buffer, InStr(1, buffer, vbSTX) + Len(vbSTX)), vbFS)
            If IsArray(myArray) Then
                For x = LBound(myArray) To UBound(myArray) - 1 Step 2
                    With myTokens
                        If Trim$(myArray(x)) <> vbNullString Then
                            .Add myArray(x + 1), myArray(x)
                        End If
                    End With
                Next
            End If
        End If
        GetTransactionName = Trans
        Exit Function
    
Err_GetTransName:
    GetTransactionName = Trans
    
End Function

Private Function Exists(ByRef m_Col As Collection, ByVal sKey As String) As Boolean
Dim vItem As Variant
On Local Error GoTo err_Exists
        If m_Col.Count > 0 Then
            If IsObject(m_Col.Item(sKey)) Then
                Set vItem = m_Col.Item(sKey)
            Else
                vItem = m_Col.Item(sKey)
            End If
            Exists = True
        End If
        Exit Function
err_Exists:
    Err.Clear
    Exists = False
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
        End If
        PropPageCallback_SaveChanges = True
    Else
        Log cModuleName, "PropPageCallback_SaveChanges", "Unable to reload settings from string.", GPS_LOG_ERROR, , XMLDOMNode, GPS_SOURCE_XML
    End If
    Exit Function
err_PropPageCallback:
    Log cModuleName, "PropPageCallback_SaveChanges", Err.Description, GPS_LOG_WARNING, Err.Number
End Function

'FLOW CONTROL
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


'window to send - begin
Private Function LimitMessages(ByVal SecondsPerTrans As Long, ByVal SecondsForTrans As Long)
On Error GoTo err_LimitMessages
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
    Exit Function
err_LimitMessages:
    Log cModuleName, "LimitMessages", Err.Number & ":" & Err.Description, GPS_LOG_ERROR
    Err.Clear
End Function

Private Function EnableCallback()
On Error GoTo err_EnableCallBack
    Set m_Notify = New ccrpTimers6.ccrpCountdown
    With m_Notify
        .Duration = (m_SecondsForTransmission * 1000)
        .Interval = 100 'progress indicator
        Set .Notify = Me
        .Enabled = True
        dLastTransmission = CDbl(Now)
    End With
    Log cModuleName, "EnableCallback", "Timer callback enabled", GPS_LOG_VERBOSE
    Exit Function
err_EnableCallBack:
    Log cModuleName, "EnableCallback", Err.Number & ":" & Err.Description, GPS_LOG_ERROR
    Err.Clear
End Function

Private Function DisableCallback()
On Error GoTo err_DisableCallback
    If Not m_Notify Is Nothing Then
        m_Notify.Enabled = False
        Set m_Notify.Notify = Nothing
        Set m_Notify = Nothing
        Log cModuleName, "DisableCallback", "Timer callback disabled", GPS_LOG_VERBOSE
    End If
    Exit Function
err_DisableCallback:
    Log cModuleName, "DisableCallback", Err.Number & ":" & Err.Description, GPS_LOG_ERROR
    Err.Clear
End Function

Private Sub ICcrpCountdownNotify_Tick(ByVal TimeRemaining As Long)

End Sub

Private Sub ICcrpCountdownNotify_Timer()
    bTimeToWork = False
    dLastTransmission = Now
    DisableCallback
    Log cModuleName, "ICcrpCountdownNotify_Timer", "Timer callback disabled", GPS_LOG_VERBOSE
End Sub
'window to send - end

'manual message send - begin
Private Function pollingBuffer(ByVal vData As String, ByVal vTime As String)
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

Private Sub ManualTime_Timer(ByVal Milliseconds As Long)
Dim oObj As Statement
Dim oMsg As Message
    If m_PollingInfo.Count > 0 Then
        For Each oObj In m_PollingInfo
            With oObj
                If DateDiff("s", Now, .NextAlarm) < 1 Then
                    Log cModuleName, "ManualTime_Timer", "Sending timed manual message.", GPS_LOG_VERBOSE
                    Set oMsg = New Message
                    oMsg.rawMessage = .pollingBuffer
                    oMsg.MessageStatus = GPS_MSG_PROCESSED
                    BuildMessage oMsg, True
                    'bufferToSend .pollingBuffer, False
                    .NextAlarm = DateAdd("s", .pollingInterval, Now)
                    Log cModuleName, "ManualTime_Timer", "Next message will be sent " & Format$(.NextAlarm, "hh:nn:ss") & ".", GPS_LOG_VERBOSE
                End If
            End With
        Next
    Else
        ManualTime.Enabled = False
    End If
End Sub
'manual message send - end
'FLOW CONTROL
