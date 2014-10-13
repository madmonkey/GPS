VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1905
   LinkTopic       =   "Form1"
   ScaleHeight     =   630
   ScaleWidth      =   1905
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer ManualTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   120
   End
   Begin VB.Timer CommTime 
      Interval        =   1
      Left            =   960
      Top             =   120
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   480
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event ReceivedMessage(ByVal DataIn As String)
Public Event ReceivedMessageFrom(ByVal DataIn As String, ByVal EndPoint As String)
Public Event Log(ByVal method As String, ByVal Message As String, ByVal LogDetail As HTE_GPS.GPS_LOG_DETAIL)
Public Event LogDetail(ByVal method As String, ByVal Message As String, ByVal LogDetail As HTE_GPS.GPS_LOG_DETAIL, ByVal ErrorID As Long, ByVal LogSource As String, ByVal LogSourceDetail As HTE_GPS.GPS_LOG_SOURCE)
Public Event StatusChange(ByVal statusCode As HTE_GPS.GPS_PROCESSOR_STATUS)
Private currentType As HTE_GPS.GPSConfiguration
'''Private m_GPSBuffer As stringBuilder 'String 'Incoming data
Private m_bufferToSend As stringBuilder 'String 'Outgoing data
Private bConnected As Boolean 'are we connected?
Private bValidateMessage As Boolean 'Should we Validate the message?
Private Const cConnectionTimeout = 15
Private Const cMsgSep = "@~|*|~@"
Private m_PollingInfo As Collection
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private m_Host As String
Private m_Port As Long
'ADDED FOR SOFTWARE FLOWCONTROL - DEVICES THAT CANNOT BE CONFIGURED FOR TIME TO SEND (NMEA)
Private m_SecondsPerTransmission As Long 'send messages every x seconds
Private m_SecondsForTransmission As Long 'send messages for y seconds
Private m_Notify As ccrpTimers6.ccrpCountdown
Implements ccrpTimers6.ICcrpCountdownNotify
Private bTimeToWork As Boolean
Private dLastTransmission As Double
'ADDED FOR EVER-CHANGING REQUIREMENTS (ALIASING) -
'WHY IS THE LACK OF REQUIREMENTS ON MANAGEMENT'S PART
'A CONTINUAL HEARTACHE FOR DEVELOPERS???
'we are keeping a stringbuilder for every address that sends to us
'since the intial requirement(s) dictated that the message would contain
'the identifier this is the only way to be sure that we are not
'quashing some else's buffer (the only way to be sure)
Private dicBuffers As Scripting.Dictionary
Private bUnloadingFromMemory As Boolean

Public Function LimitMessages(ByVal SecondsPerTrans As Long, ByVal SecondsForTrans As Long)
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
    RaiseEvent Log("LimitMessages", Err.Number & ":" & Err.Description, GPS_LOG_ERROR)
    Err.Clear
End Function

Private Function EnableCallback()
On Error GoTo err_EnableCallBack
    If Not bUnloadingFromMemory Then
        Set m_Notify = New ccrpTimers6.ccrpCountdown
        With m_Notify
            .Duration = (m_SecondsForTransmission * 1000)
            .Interval = 100 'progress indicator
            Set .Notify = Me
            .Enabled = True
            dLastTransmission = CDbl(Now)
        End With
        RaiseEvent Log("EnableCallback", "Timer callback enabled", GPS_LOG_VERBOSE)
        End If
    Exit Function
err_EnableCallBack:
    RaiseEvent Log("EnableCallback", Err.Number & ":" & Err.Description, GPS_LOG_ERROR)
    Err.Clear
End Function
Private Function DisableCallback()
On Error GoTo err_DisableCallback
    If Not m_Notify Is Nothing Then
        m_Notify.Enabled = False
        Set m_Notify.Notify = Nothing
        Set m_Notify = Nothing
        RaiseEvent Log("DisableCallback", "Timer callback disabled", GPS_LOG_VERBOSE)
    End If
    Exit Function
err_DisableCallback:
    RaiseEvent Log("DisableCallback", Err.Number & ":" & Err.Description, GPS_LOG_ERROR)
    Err.Clear
End Function

Private Sub Form_Load()
    Set dicBuffers = New Scripting.Dictionary
    bUnloadingFromMemory = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo err_Form_Unload
    bUnloadingFromMemory = True
    DisableCallback
    Set m_bufferToSend = Nothing
    With dicBuffers
        .removeAll
    End With
    Set dicBuffers = Nothing
    Exit Sub
err_Form_Unload:
    RaiseEvent Log("Form_Unload", Err.Number & ":" & Err.Description, GPS_LOG_ERROR)
    Err.Clear
End Sub

'WHEN SENDING UDP, IT IS POSSIBLE THAT THE WINSOCK CONTROL
'WILL "LOSE" THE INFORMATION INITIALLY ASSIGNED TO IT, WE NEED
'TO PERSIST THIS INOFRMATION SO WE MAY REASSIGN WHEN SENDING
Public Property Get Host() As String
    Host = m_Host
End Property

Public Property Let Host(ByVal vData As String)
    m_Host = vData
    If Winsock1.Protocol = sckUDPProtocol Then
        Winsock1.RemoteHost = m_Host
        bConnected = (Port <> 0) And (Host <> vbNullString)
    End If
End Property

Public Property Get Port() As Long
    Port = m_Port
End Property

Private Property Get ShouldSendData() As Boolean
    'NEED TO SET VARIABLE FOR REMOTEPORT - WHETHER OR NOT TO BROADCAST DATA
    'IF REMOTE SET TO "0" A RANDOM PORT IS ASSIGNED AND ATTEMPTS TO SEND
    'WE USE THIS TAG TO AFFIRM WHETHER OR NOT WE "SEND" TO THE REMOTE PORT OR NOT!
    ShouldSendData = CBool(Winsock1.Tag <> "FALSE")
End Property

Public Property Let Port(ByVal vData As Long)
    m_Port = vData
    If Winsock1.Protocol = sckUDPProtocol Then
        Winsock1.RemotePort = m_Port
        Winsock1.Tag = UCase$(vData <> 0)
        bConnected = (Port <> 0) And (Host <> vbNullString)
    Else
        Winsock1.RemotePort = m_Port
    End If
End Property

Public Property Get MessageType() As HTE_GPS.GPSConfiguration
    MessageType = currentType
End Property

Public Property Let MessageType(ByRef GPSType As HTE_GPS.GPSConfiguration)
    currentType = GPSType
End Property

Public Property Get ValidateMessage() As Boolean
    ValidateMessage = bValidateMessage
End Property

Public Property Let ValidateMessage(ByVal vData As Boolean)
    bValidateMessage = vData
End Property

Public Function bufferToSend(ByVal vData As String, Optional ByVal fromChain As Boolean = False)
On Local Error Resume Next
Static dMaxTime As Double
Dim x As Long
Dim arrToSend() As String
    
On Local Error Resume Next
    If ShouldSendData Then
        RaiseEvent Log("bufferToSend", "Ready to send received buffer.", GPS_LOG_VERBOSE) ', 0&, currentType.SOM & vData & currentType.EOM, GPS_SOURCE_BINARY)
        If m_bufferToSend Is Nothing Then Set m_bufferToSend = New stringBuilder
        If ValidateMessage And Not fromChain Then 'prepare for bad connection
            m_bufferToSend.Append currentType.SOM & vData & currentType.EOM & cMsgSep
        Else
            m_bufferToSend.Append vData & cMsgSep
        End If
    
Retry:
        If bConnected Then
            arrToSend = Split(m_bufferToSend.ToString, cMsgSep)
            For x = LBound(arrToSend) To UBound(arrToSend)
                If arrToSend(x) <> vbNullString Then
                    'If using UDP with Winsock, sending data may cause control to forget Remote Host/Port properties...<sigh>
                    If Winsock1.Protocol = sckUDPProtocol Then
                        Winsock1.RemoteHost = Host
                        Port = Port
                    End If
                    Winsock1.SendData (arrToSend(x))
                    If Err Then
                        bConnected = (Winsock1.State = sckOpen)
                        RaiseEvent Log("bufferToSend", "An error occured - " & Err.Description & " err number: " & Err.Number, GPS_LOG_ERROR)
                        dMaxTime = Now
                        bConnected = Reconnect
                        RaiseEvent Log("bufferToSend", "Reconnect = " & CStr(bConnected), GPS_LOG_WARNING)
                        Exit For
                    Else
                        RaiseEvent LogDetail("bufferToSend", "Buffer sent.", GPS_LOG_VERBOSE, 0&, arrToSend(x), GPS_SOURCE_BINARY)
                        RaiseEvent StatusChange(GPS_STAT_READYANDWILLING)
                        m_bufferToSend.Remove 1, m_bufferToSend.Find(cMsgSep) + Len(cMsgSep)
                    End If
                End If
            Next
        Else
            If Abs(DateDiff("s", dMaxTime, Now)) > cConnectionTimeout Then
                dMaxTime = Now
                bConnected = Reconnect
                RaiseEvent Log("bufferToSend", "Reconnect = " & CStr(bConnected), GPS_LOG_WARNING)
            End If
        End If
    Else
        RaiseEvent Log("bufferToSend", "Process configured NOT to send data => RemotePort: 0", GPS_LOG_INFORMATION)
    End If
    'If bConnected Then RaiseEvent StatusChange(GPS_STAT_READYANDWILLING)

End Function

Private Function Reconnect() As Boolean
On Local Error Resume Next
    RaiseEvent Log("bufferToSend", "Attempting reconnection....", GPS_LOG_ERROR)
    If Winsock1.Protocol = sckUDPProtocol Then
        Winsock1.Close
        Winsock1.RemoteHost = Host
        Port = Port
        RaiseEvent Log("bufferToSend", "Resetting UDP Parameters as Remote [" & Host & ":" & Port & "] / Local [" & Winsock1.LocalIP & ":" & Winsock1.LocalPort & "]....", GPS_LOG_VERBOSE)
        Winsock1.Bind
        Reconnect = (Err.Number = 0)
        RaiseEvent Log("Reconnect", "Reconnect = " & CStr(Err.Number = 0), GPS_LOG_WARNING)
    Else
        Winsock1.Close: DoEvents
        Winsock1.Connect Host, Port: DoEvents
        RaiseEvent Log("bufferToSend", "Resetting TCP/IP Parameters as [" & Host & ":" & Port & "]....", GPS_LOG_VERBOSE)
        Reconnect = (Winsock1.State = sckOpen) Or (Winsock1.State = sckConnected)
        RaiseEvent Log("Reconnect", "Reconnect = " & CStr((Winsock1.State = sckOpen) Or (Winsock1.State = sckConnected)), GPS_LOG_WARNING)
    End If
    
End Function

Private Sub ICcrpCountdownNotify_Tick(ByVal TimeRemaining As Long)
    RaiseEvent Log("ICcrpCountdownNotify_Tick", "Tick - " & Format$(TimeRemaining / 1000, "0.0") & " sec till expiration.", GPS_LOG_VERBOSE)
End Sub

Private Sub ICcrpCountdownNotify_Timer()
    bTimeToWork = False
    dLastTransmission = Now
    DisableCallback
    RaiseEvent Log("ICcrpCountdownNotify_Timer", "Timer callback disabled", GPS_LOG_VERBOSE)
End Sub

Private Sub Winsock1_Close()
    RaiseEvent Log("Winsock1_Close", "Socket closed.", GPS_LOG_WARNING)
    bConnected = False
End Sub

Private Sub Winsock1_Connect()
    RaiseEvent Log("Winsock1_Connect", "Socket connected.", GPS_LOG_VERBOSE)
    bConnected = True
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    RaiseEvent Log("Winsock1_ConnectionRequest", "Connection Request requestID = " & requestID & ".", GPS_LOG_VERBOSE)
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strData As String, strAddress As String
Dim gpsStream As stringBuilder
'260018 - BUG: Winsock Control Run-Time Error 10054 in DataArrival Event for UDP
'http://support.microsoft.com/default.aspx?scid=kb%3ben-us%3b260018
'Keywords: kbAPI kbbug kbCtrl kbIP kbnetwork kbpending kbWinsock KB260018
'This article was previously published under Q260018
On Local Error Resume Next
        Winsock1.GetData strData, vbString, bytesTotal
        strAddress = Winsock1.RemoteHostIP
        RaiseEvent Log("Winsock1_DataArrival", CStr(bytesTotal) & " byte(s) received from [" & strAddress & "]", GPS_LOG_VERBOSE)
        If Err Then
            Select Case Err.Number
                Case 10054
                    RaiseEvent Log("Winsock1_DataArrival", "Remote side is NOT listening! " & Err.Description, GPS_LOG_WARNING)
                Case Else
                    RaiseEvent Log("Winsock1_DataArrival", Err.Number & ": " & Err.Description, GPS_LOG_ERROR)
            End Select
        Else
            If m_Notify Is Nothing Then
                If Not bTimeToWork Then
                    RaiseEvent Log("Winsock1_DataArrival", "Last transmission was " & CStr(CDate(dLastTransmission)) & "; " & DateDiff("s", dLastTransmission, Now) & " seconds ago; " & m_SecondsPerTransmission & " seconds for transactions.", GPS_LOG_VERBOSE)
                    bTimeToWork = DateDiff("s", dLastTransmission, Now) >= (m_SecondsPerTransmission)
                    If bTimeToWork Then EnableCallback
                End If
            End If
            If bTimeToWork Then
                If Not dicBuffers.Exists(strAddress) Then
                    Set gpsStream = New stringBuilder
                    gpsStream.ChunkSize = 1024 'reduce the default size this will get big
                    dicBuffers.Add strAddress, gpsStream
                End If
                Set gpsStream = dicBuffers.Item(strAddress)
                gpsStream.Append strData
                CommTime.Enabled = True
            Else
                'I really, really don't like this, but since multiple transmissions will effectively "kill" the switch/network this
                'stop-gate is provided to limit the number of transmissions to the switch (specifically for NMEA devices)....
                'We just empty the buffer here and log that we've done so!!
                'Changed from Warning to Verbose - so not to "flicker" for NMEA
                RaiseEvent LogDetail("Winsock1_DataArrival", "Dropping message from delivery!", GPS_LOG_VERBOSE, 0, strData, GPS_SOURCE_BINARY)
            End If
            RaiseEvent StatusChange(GPS_STAT_READYANDWILLING)
        End If
    Exit Sub
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    CancelDisplay = True
    RaiseEvent LogDetail("Winsock1_Error", Description, GPS_LOG_ERROR, Number, Source, GPS_SOURCE_STRING)
End Sub

Private Sub CommTime_Timer()
Dim sBuffer As String, keys As Variant
Dim gpsStream As stringBuilder
Dim Length As Long
Dim i As Long
On Error GoTo err_CommTime
    If Not dicBuffers Is Nothing Then 'm_GPSBuffer Is Nothing Then
        If dicBuffers.Count > 0 Then
            keys = dicBuffers.keys
            For i = LBound(keys) To UBound(keys)
                Set gpsStream = dicBuffers.Item(keys(i))
                If gpsStream.Length > 0 Then
                    sBuffer = ParseGPSMessage(gpsStream)
                    While sBuffer <> vbNullString
                        'RaiseEvent ReceivedMessage(sBuffer)
                        RaiseEvent ReceivedMessageFrom(sBuffer, keys(i))
                        bufferToSend sBuffer, False
                        sBuffer = ParseGPSMessage(gpsStream)
                    Wend
                End If
            Next
            'check to see if we still have work to do...
            Length = 0
            keys = dicBuffers.keys
            For i = LBound(keys) To UBound(keys)
                Set gpsStream = dicBuffers.Item(keys(i))
                Length = Length + gpsStream.Length
                If Length > 0 Then Exit For
            Next
            CommTime.Enabled = Length > 0 'm_GPSBuffer.length > 0
        Else
            CommTime.Enabled = False
        End If
    Else
        CommTime.Enabled = False
    End If
    Exit Sub
    
err_CommTime:
    RaiseEvent Log("CommTime_Timer", Err.Number & ":" & Err.Description, GPS_LOG_ERROR)
    Err.Clear
End Sub

Private Function ParseGPSMessage(ByRef m_GPSBuffer As stringBuilder) As String
Dim startPos As Long, endPos As Long

On Error GoTo err_ParseGPSMessage
    If Not m_GPSBuffer Is Nothing Then
        With m_GPSBuffer
            If ValidateMessage Then
                If .Length > 0 Then
                    startPos = .Find(currentType.SOM)
                    If startPos > 0 Then
                        If startPos > 1 Then
                            .Remove 0, startPos - 1
                            startPos = 1
                        End If
                        endPos = .Find(currentType.EOM, startPos + 1)
                        If endPos > startPos Then
                            ParseGPSMessage = Mid$(.ToString, startPos + Len(currentType.SOM), (endPos - (startPos + Len(currentType.SOM))))    'Mid$(.ToString, startPos + Len(GPSConfig.SOM), (endPos - (startPos + Len(GPSConfig.SOM))))
                            .Remove startPos - 1, endPos
                        End If
                    End If
                End If
            Else
                ParseGPSMessage = .ToString
                .Remove 0, .Length
            End If
        End With
    End If
    Exit Function
err_ParseGPSMessage:
    RaiseEvent Log("ParseGPSMessage", Err.Number & ":" & Err.Description, GPS_LOG_ERROR)
    Err.Clear
End Function

Private Function ValidGPSMessage(ByVal bufferToSend As String) As Boolean
    'Some units send so many different message types, don't flood airways with messages that will not be used
    Select Case currentType.GPSType
        Case GPS_TYPE_0, GPS_TYPE_1, GPS_TYPE_2, GPS_TYPE_3
            ValidGPSMessage = True
        Case Else
            ValidGPSMessage = False
    End Select
End Function

Private Sub Winsock1_SendComplete()
    RaiseEvent Log("Winsock1_SendComplete", "Transfer completed.", GPS_LOG_VERBOSE)
End Sub

Private Sub Winsock1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    RaiseEvent Log("Winsock1_SendProgress", bytesSent & " bytesSent " & bytesRemaining & " bytes remaining.", GPS_LOG_VERBOSE)
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
                    bufferToSend .pollingBuffer, False
                    .NextAlarm = DateAdd("s", .pollingInterval, Now)
                    RaiseEvent Log("ManualTime_Timer", "Next message will be sent " & Format$(.NextAlarm, "hh:nn:ss") & ".", GPS_LOG_VERBOSE)
                End If
            End With
        Next
    Else
        ManualTime.Enabled = False
    End If
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
