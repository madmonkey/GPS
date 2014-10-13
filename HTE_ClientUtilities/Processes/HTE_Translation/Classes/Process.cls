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
Option Compare Binary

Private Const cModuleName = "HTE_Translation.Process"
Private gApp As HTE_GPS.Processes
Private m_Status As HTE_GPS.GPS_PROCESSOR_STATUS
Private localSettings As MSXML2.DOMDocument30
Private m_Settings As String
Private m_Property As Object
Private m_Transforms As Collection
Implements HTE_GPS.Process
Implements HTE_GPS.PropPageCallback
Implements ccrpTimers6.ICcrpTimerNotify

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private arrLookups(0 To 1, 0 To 35) As String
Private cReplace As stringReplace
Private m_AutoSaveEntity As Boolean
Private WithEvents aliasDataChanges As HTE_PubData.Subscriber 'used to be informed when dataset changes
Attribute aliasDataChanges.VB_VarHelpID = -1
Private m_aliasData As HTE_GPSData.DataSource 'common-used DAL functions
Private dataRetrievalTimer As ccrpTimers6.ccrpTimer  'retrieve database updates on individual thread
Attribute dataRetrievalTimer.VB_VarHelpID = -1

'Entity container - entities (each corresponding
Private applicationKeys As Scripting.Dictionary
Private deviceKeys As Scripting.Dictionary
Private addressKeys As Scripting.Dictionary
Private entityContainer As Scripting.Dictionary
Private processMutex As Mutex
'faster than vb6 strcomp for equality (even with binarycompare)
'http://www.xbeat.net/vbspeed/c_IsSameText.htm
'usage: fRet = (lstrcmpi(sDum1, sDum2) = 0)
Private Declare Function lstrcmpi Lib "kernel32" Alias "lstrcmpiA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Const cEntityDevice As String = "/@#{ENTITYDEV};"
Private Const cEntityUnit As String = "/@#{ENTITYUNIT};"
Private Const cEntityStat As String = "/@#{ENTITYSTAT};"
Private Const cEntityCode As String = "/@#{ENTITYCODE};"
Private bWorking As Boolean 'currently caching db
Private bRestarting As Boolean 'process is being initialized or reinitialized
Private bChangesPending As Boolean 'flag that critcial memory resource about to be altered
Private bTransactionTime As Boolean 'flag to indicate whether to keep track of time-elaspsed between transactions processed
Private Const cElapsedTime As String = "/@#{TIMEELAPSED};"
Dim elapsedTime As cTiming

Private Sub aliasDataChanges_OnReceiveString(ByVal TAG As String, ByVal Data As String)
    If Not dataRetrievalTimer Is Nothing Then
        If Not dataRetrievalTimer.Enabled Then
            Log cModuleName, "aliasDataChanges_OnReceiveString", "Enabling retrieval process...", GPS_LOG_VERBOSE
            dataRetrievalTimer.Enabled = True
            bChangesPending = True
            Log cModuleName, "aliasDataChanges_OnReceiveString", "Timer enabled...", GPS_LOG_VERBOSE
        End If
    End If
End Sub

Private Sub Class_Initialize()
    m_Status = GPS_STAT_INITIALIZED
    Set localSettings = New MSXML2.DOMDocument30
    Set m_Transforms = New Collection
    Set cReplace = New stringReplace
    LoadLookups
    Set m_aliasData = New HTE_GPSData.DataSource
    Set aliasDataChanges = m_aliasData.SubscribeToChanges
    Set elapsedTime = New cTiming
    elapsedTime.Reset
    InitializeDataTimer
    Set processMutex = New Mutex
    processMutex.ConstructMutex 'allow mutex to auto-generate an instance specific mutex - we only care about our instance memory
End Sub
Private Sub InitializeDataTimer()
    If dataRetrievalTimer Is Nothing Then
        Set dataRetrievalTimer = New ccrpTimers6.ccrpTimer
        With dataRetrievalTimer
            .EventType = TimerPeriodic
            .Interval = 20000
            .Enabled = False
            Set .Notify = Me
        End With
    End If
End Sub
Private Sub InitializeIMDB()
    Set entityContainer = New Scripting.Dictionary
    Set applicationKeys = New Scripting.Dictionary
    Set deviceKeys = New Scripting.Dictionary
    Set addressKeys = New Scripting.Dictionary
End Sub

Private Sub TerminateIMDB()
    If Not entityContainer Is Nothing Then entityContainer.removeAll: Set entityContainer = Nothing
    If Not applicationKeys Is Nothing Then applicationKeys.removeAll: Set applicationKeys = Nothing
    If Not deviceKeys Is Nothing Then deviceKeys.removeAll: Set deviceKeys = Nothing
    If Not addressKeys Is Nothing Then addressKeys.removeAll: Set addressKeys = Nothing
End Sub

Private Sub IndexIMDB()
Dim keys As Variant
Dim ent As HTE_Entity.Entity
Dim i As Long
Dim uniqueKey As String
    If Not entityContainer Is Nothing Then
        keys = entityContainer.keys
        For i = 0 To UBound(keys)
            uniqueKey = keys(i)
            Set ent = entityContainer.Item(uniqueKey)
            If Not ent Is Nothing Then
                With ent
                    'how do we expect the device to account for itself
                    If Len(.ApplicationID) > 0 Then
                        If Not applicationKeys.Exists(.ApplicationID) Then
                            applicationKeys.Add .ApplicationID, uniqueKey
                        End If
                    End If
                    'what-if any, physical characteristics do we have?
                    If Len(.MACAddress) > 0 Then
                        If Not deviceKeys.Exists(.MACAddress) Then
                            deviceKeys.Add .MACAddress, uniqueKey
                        End If
                    End If
                    'what is the address we received from
                    If Len(.IPAddress) > 0 Then
                        If Not addressKeys.Exists(.IPAddress) Then
                            addressKeys.Add .IPAddress, uniqueKey
                        End If
                    End If
                End With
            End If
        Next
        Log cModuleName, "IndexIMDB", "Indexed Distinctive Lookups: applicationKeys [" & applicationKeys.Count & "]; deviceKeys [" & deviceKeys.Count & "]; addressKeys [" & addressKeys.Count & "]", GPS_LOG_VERBOSE
    End If
End Sub

Private Sub LoadLookups()
    arrLookups(0, 0) = "/@#{ID};": arrLookups(1, 0) = "terminalunit"
    arrLookups(0, 1) = "/@#{LOGSTAT};": arrLookups(1, 1) = "logonstatus"
    arrLookups(0, 2) = "/@#{ALIAS};": arrLookups(1, 2) = "alias"
    arrLookups(0, 3) = "/@#{STATUS};": arrLookups(1, 3) = "status"
    arrLookups(0, 4) = "/@#{STATCODE};": arrLookups(1, 4) = "statuscode"
    arrLookups(0, 5) = "/@#{USER1};": arrLookups(1, 5) = "user1"
    arrLookups(0, 6) = "/@#{CAD};": arrLookups(1, 6) = "cadsystem"
    arrLookups(0, 7) = "/@#{ROLES};": arrLookups(1, 7) = "roles"
    arrLookups(0, 8) = "/@#{SHFT};": arrLookups(1, 8) = "shft"
    arrLookups(0, 9) = "/@#{AREA};": arrLookups(1, 9) = "area"
    arrLookups(0, 10) = "/@#{UNTT};": arrLookups(1, 10) = "untt"
    arrLookups(0, 11) = "/@#{SUPC};": arrLookups(1, 11) = "supc"
    arrLookups(0, 12) = "/@#{BEAT};": arrLookups(1, 12) = "beat"
    arrLookups(0, 13) = "/@#{USER2};": arrLookups(1, 13) = "user2"
    arrLookups(0, 14) = "/@#{USER3};": arrLookups(1, 14) = "user3"
    arrLookups(0, 15) = "/@#{USER4};": arrLookups(1, 15) = "user4"
    arrLookups(0, 16) = "/@#{BEAT1};": arrLookups(1, 16) = "beat1"
    arrLookups(0, 17) = "/@#{BEAT2};": arrLookups(1, 17) = "beat2"
    arrLookups(0, 18) = "/@#{BEAT3};": arrLookups(1, 18) = "beat3"
    arrLookups(0, 19) = "/@#{BEAT4};": arrLookups(1, 19) = "beat4"
    arrLookups(0, 20) = "/@#{SB01};": arrLookups(1, 20) = "sb01"
    arrLookups(0, 21) = "/@#{SB02};": arrLookups(1, 21) = "sb02"
    arrLookups(0, 22) = "/@#{SB03};": arrLookups(1, 22) = "sb03"
    arrLookups(0, 23) = "/@#{SB04};": arrLookups(1, 23) = "sb04"
    arrLookups(0, 24) = "/@#{SB05};": arrLookups(1, 24) = "sb05"
    arrLookups(0, 25) = "/@#{SB06};": arrLookups(1, 25) = "sb06"
    arrLookups(0, 26) = "/@#{SB07};": arrLookups(1, 26) = "sb07"
    arrLookups(0, 27) = "/@#{SB08};": arrLookups(1, 27) = "sb08"
    arrLookups(0, 28) = "/@#{SB09};": arrLookups(1, 28) = "sb09"
    arrLookups(0, 29) = "/@#{SB10};": arrLookups(1, 29) = "sb10"
    arrLookups(0, 30) = "/@#{SB11};": arrLookups(1, 30) = "sb11"
    arrLookups(0, 31) = "/@#{SB12};": arrLookups(1, 31) = "sb12"
    arrLookups(0, 32) = "/@#{SB13};": arrLookups(1, 32) = "sb13"
    arrLookups(0, 33) = "/@#{SB14};": arrLookups(1, 33) = "sb14"
    arrLookups(0, 34) = "/@#{SB15};": arrLookups(1, 34) = "sb15"
    arrLookups(0, 34) = "/@#{SB16};": arrLookups(1, 34) = "sb16"
    arrLookups(0, 35) = "/@#{DEVID};": arrLookups(1, 35) = "fauxdevid"
End Sub

Private Sub Class_Terminate()
    m_Status = GPS_STAT_UNINITIALIZED
    If Not dataRetrievalTimer Is Nothing Then
        dataRetrievalTimer.Enabled = False
        Set dataRetrievalTimer = Nothing
    End If
    Set localSettings = Nothing
    Set aliasDataChanges = Nothing
    Set m_aliasData = Nothing
    Set elapsedTime = Nothing
    clearTransforms
    processMutex.DestroyMutex
    Set processMutex = Nothing
End Sub

Private Sub ICcrpTimerNotify_Timer(ByVal Milliseconds As Long)
Dim bMutexObtained As Boolean
    '***MUTEX USAGE
    'non-blocking event-sink - lookup DB
    If Not bWorking Then
On Error Resume Next
        bWorking = True
        Do While Not processMutex.RequestMutex(10)
            If bRestarting Then
                bWorking = False
                Exit Sub
            End If
        Loop
        bMutexObtained = True
        Log cModuleName, "ICcrpCountdownNotify_Timer", "Mutex obtained retrieving/reindexing based on persistent store changes.", GPS_LOG_VERBOSE
        TerminateIMDB
        Log cModuleName, "ICcrpCountdownNotify_Timer", "Cleared memory store.", GPS_LOG_VERBOSE
        InitializeIMDB
        Log cModuleName, "ICcrpCountdownNotify_Timer", "Reinitialized memory store.", GPS_LOG_VERBOSE
        If m_aliasData Is Nothing Then
            Log cModuleName, "ICcrpCountdownNotify_Timer", "Reinitializing proxy.", GPS_LOG_VERBOSE
            Set m_aliasData = New HTE_GPSData.DataSource
            Log cModuleName, "ICcrpCountdownNotify_Timer", "Reinitialized proxy.", GPS_LOG_VERBOSE
            Set aliasDataChanges = m_aliasData.SubscribeToChanges
            Log cModuleName, "ICcrpCountdownNotify_Timer", "Reintialized subscriber.", GPS_LOG_VERBOSE
        End If
        Log cModuleName, "ICcrpCountdownNotify_Timer", "Obtaining entities memory store.", GPS_LOG_VERBOSE
        Set entityContainer = m_aliasData.entities
        Log cModuleName, "ICcrpCountdownNotify_Timer", "Retrieved configuration settings.", GPS_LOG_VERBOSE
        IndexIMDB
        Log cModuleName, "ICcrpCountdownNotify_Timer", "Indexing completed.", GPS_LOG_VERBOSE
        dataRetrievalTimer.Enabled = False
        bChangesPending = False
        If bMutexObtained Then processMutex.DiscardMutex
        Log cModuleName, "ICcrpCountdownNotify_Timer", "Mutex discarded.", GPS_LOG_VERBOSE
        bWorking = False
        
    End If
End Sub

Private Sub Process_Connect(ByVal parentProcess As HTE_GPS.Processes, ByVal cookie As String)
    Set gApp = parentProcess
    gApp.Connected Me, cookie
End Sub

Private Sub Process_Disconnect()
    Log cModuleName, "Process_Disconnect", "Unhooking process", GPS_LOG_VERBOSE
    checkForChanges
    Set m_Property = Nothing
    Set gApp = Nothing
End Sub

Private Property Get Process_FriendlyName() As String
    Process_FriendlyName = "Message Conversion Process"
End Property

Private Property Get Process_HostVersion() As String
    Process_HostVersion = "1.0.0"
End Property

Private Sub Process_ProcessMessage(ByVal propData As HTE_GPS.GPSMessage)
    transformMessage propData
    gApp.ProcessComplete propData, Me
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
    localSettings.async = False
    If bBaseline Then m_Settings = sXML
    loadLocalSettings = localSettings.LoadXML(sXML)
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

Private Function retrieveLocalSettings() As String
    If Not localSettings Is Nothing Then
        retrieveLocalSettings = localSettings.xml
    Else
        retrieveLocalSettings = m_Settings
    End If
End Function

Private Sub setProperties()
Dim iRoot As MSXML2.IXMLDOMNode
Dim iNode As MSXML2.IXMLDOMNode
Dim iAttribute As MSXML2.IXMLDOMNode
Dim iStylesheet As MSXML2.DOMDocument30

Dim i As Long
On Error GoTo err_setProperties
    Log cModuleName, "setProperties", "Begin.", GPS_LOG_VERBOSE
    If Not localSettings Is Nothing Then
        bRestarting = True
        Log cModuleName, "setProperties", "localSettings is something.", GPS_LOG_VERBOSE, 0, localSettings.xml, GPS_SOURCE_STRING
        m_AutoSaveEntity = CBool(StrComp(getSettings(cAutoInsert, cAutoInsertValue), "True", vbTextCompare) = 0)
        Set iRoot = localSettings.documentElement
        Log cModuleName, "setProperties", "iRoot assigned.", GPS_LOG_VERBOSE
        bRestarting = False
        bChangesPending = True 'flag that critcial memory resource about to be altered
        ICcrpTimerNotify_Timer 0&
        If Not iRoot Is Nothing Then
            If iRoot.hasChildNodes Then
                Log cModuleName, "setProperties", "iRoot has children", GPS_LOG_VERBOSE
                clearTransforms
                Log cModuleName, "setProperties", "Cleared existing transformations", GPS_LOG_VERBOSE
                For i = 0 To iRoot.childNodes.Length - 1
                    Log cModuleName, "setProperties", "i = " & CStr(i), GPS_LOG_VERBOSE
                    Set iStylesheet = New DOMDocument30
                    iStylesheet.async = False
                    Set iNode = localSettings.documentElement.childNodes(i)
                    If Not iNode Is Nothing Then
                        Log cModuleName, "setProperties", "iNode is something", GPS_LOG_VERBOSE
                        Set iAttribute = iNode.Attributes.getNamedItem(cAttrType)
                        If iStylesheet.LoadXML(iNode.xml) Then
                            Log cModuleName, "setProperties", "iStylesheet assigned.", GPS_LOG_VERBOSE
                            If iStylesheet.hasChildNodes Then
                                Log cModuleName, "setProperties", "iStylesheet has children.", GPS_LOG_VERBOSE
                                If iStylesheet.childNodes(0).hasChildNodes Then
                                    Log cModuleName, "setProperties", "iStylesheet has grandchildren.", GPS_LOG_VERBOSE
                                    If Not iAttribute Is Nothing Then
                                        Log cModuleName, "setProperties", "iAttribute is something.", GPS_LOG_VERBOSE
                                        If iStylesheet.LoadXML(iStylesheet.childNodes(0).nodeTypedValue) Then
                                            Log cModuleName, "setProperties", "Can load grand-child.", GPS_LOG_VERBOSE
                                            If Exists(m_Transforms, "xsl_" & iAttribute.nodeTypedValue) Then
                                                Log cModuleName, "setProperties", "Remove existing key.", GPS_LOG_VERBOSE
                                                m_Transforms.Remove "xsl_" & iAttribute.nodeTypedValue
                                            End If
                                            Log cModuleName, "setProperties", "Add as new.", GPS_LOG_VERBOSE
                                            m_Transforms.Add iStylesheet, "xsl_" & iAttribute.nodeTypedValue
                                        Else
                                            Log cModuleName, "setProperties", "Unable to load child process stylesheet!", GPS_LOG_ERROR
                                        End If
                                    Else
                                        Log cModuleName, "setProperties", "Unable to find stylesheet message identifier!", GPS_LOG_ERROR
                                    End If
                                Else
                                    Log cModuleName, "setProperties", "No sub-children for root process!", GPS_LOG_ERROR
                                End If
                            Else
                                Log cModuleName, "setProperties", "No children for root process!", GPS_LOG_ERROR
                            End If
                        Else
                            Log cModuleName, "setProperties", "Unable to load Node settings!", GPS_LOG_ERROR
                        End If
                    Else
                        Log cModuleName, "setProperties", "Unable to retrieve Node settings!", GPS_LOG_ERROR
                    End If
                Next
                If m_Transforms.Count > 0 Then
                    Process_Status = GPS_STAT_READYANDWILLING
                Else
                    Process_Status = GPS_STAT_WARNING
                End If
            Else
                Process_Status = GPS_STAT_ERROR
            End If
            CurrentTypes SupportedMessages(localSettings)
            Log cModuleName, "setProperties", "Assigned current types.", GPS_LOG_VERBOSE
        Else
            Log cModuleName, "setProperties", "No children present for current node!", GPS_LOG_ERROR
        End If
    Else
        Log cModuleName, "setProperties", "localSettings is nothing.", GPS_LOG_WARNING
        Process_Status = GPS_STAT_UNKNOWN
    End If
    
    Exit Sub
    
err_setProperties:
    Log cModuleName, "setProperties", Err.Description, GPS_LOG_ERROR, Err.Number
End Sub

Private Function transformMessage(ByRef propData As HTE_GPS.GPSMessage)
Dim iTransform As MSXML2.DOMDocument30, iResult As MSXML2.DOMDocument30
Dim iNode As MSXML2.IXMLDOMNode, iWorkspace As MSXML2.DOMDocument30
Dim iOutputQualifier As MSXML2.IXMLDOMNode
Dim sResult As String
On Error GoTo err_transformMessage
    Log cModuleName, "transformMessage", "Received message.", GPS_LOG_INFORMATION
    Log cModuleName, "transformMessage", "Message received: Type = " & getMessageType(propData.Type).Desc & " Status = " & messageStatusDesc(propData.MessageStatus) & ".", GPS_LOG_VERBOSE, , propData.rawMessage, GPS_SOURCE_BINARY
    'In case raw XML enters with no tranformation - for client-side identity
    propData.rawMessage = FindIdentities(propData.rawMessage)
    'In case raw XML enters with no transformation - for server-side reconciliation
    If Len(propData.Entity) > 0 Then propData.Entity = ResolveEntityFromMessage(propData).Serialize
    If Exists(m_Transforms, "xsl_" & propData.Type) Then
        If Not m_Transforms.Item("xsl_" & propData.Type) Is Nothing Then
            If IsXML(propData.rawMessage) Then
                Set iTransform = LoadXML("<?xml version=""1.0""?>" & "<GPSMessage><rawMessage>" & ScreenMessage(propData.rawMessage) & "</rawMessage></GPSMessage>")
            Else
                Set iTransform = getXML("GPSMessage")
                AddNode iTransform, "rawMessage", ScreenMessage(propData.rawMessage), NODE_ELEMENT
            End If
            Log cModuleName, "transformMessage", "QVXR Message received: ", GPS_LOG_VERBOSE, , iTransform.xml, GPS_SOURCE_BINARY
            If TypeName(m_Transforms.Item("xsl_" & propData.Type)) = "DOMDocument" Then
                'Workspace DOMNode added to account for fluctuations in size of data with replace functions!!!
                Set iWorkspace = New MSXML2.DOMDocument30
                iWorkspace.async = False
                iWorkspace.setProperty "SelectionLanguage", "XPath"
                If iWorkspace.LoadXML(FindEntityPlaceHolders(FindIdentities(m_Transforms.Item("xsl_" & propData.Type).xml), propData.Entity)) Then
                    Set iResult = New MSXML2.DOMDocument30
                    iResult.async = False
                    iResult.setProperty "SelectionLanguage", "XPath"
                    sResult = iTransform.transformNode(iWorkspace) 'm_Transforms.Item("xsl_" & PropData.Type))
                    If sResult <> vbNullString Then
                        If iResult.LoadXML(sResult) Then
                            Set iNode = iResult.getElementsByTagName(cMsgStatTag).Item(0)
                            If Not iNode Is Nothing Then propData.MessageStatus = iNode.nodeTypedValue
                            'reset the last processed message counter
                            If propData.MessageStatus = GPS_MSG_PROCESSED And Not elapsedTime Is Nothing Then
                                elapsedTime.Reset
                            End If
                            Set iNode = iResult.getElementsByTagName(cMsgTag).Item(0)
                            'In case XSL requires client-side identification
                            If Not iNode Is Nothing Then
                                'need to add ability to output xml as raw message!
                                If HasAttributeValue(iNode, "method") Then
                                    Set iOutputQualifier = iNode.Attributes.getNamedItem("method")
                                    Select Case iOutputQualifier.Text
                                        Case "xml", "html"
                                            propData.rawMessage = UnscreenMessage(iNode.firstChild.xml)
                                        Case Else
                                            propData.rawMessage = UnscreenMessage(iNode.nodeTypedValue)
                                    End Select
                                Else
                                    propData.rawMessage = UnscreenMessage(iNode.nodeTypedValue)
                                End If
                            End If
                            Set iNode = iResult.getElementsByTagName(cMsgType).Item(0)
                            If Not iNode Is Nothing Then propData.Type = iNode.nodeTypedValue
                            Log cModuleName, "transformMessage", "Message transformed: Type = " & getMessageType(propData.Type).Desc & " Status = " & messageStatusDesc(propData.MessageStatus) & ".", GPS_LOG_VERBOSE, , propData.rawMessage, GPS_SOURCE_BINARY
                            Process_Status = GPS_STAT_READYANDWILLING
                        Else
                            Log cModuleName, "transformMessage", "Unable to load from resulting transformation -" & iResult.parseError.reason & "!", GPS_LOG_WARNING
                        End If
                    Else
                        Log cModuleName, "transformMessage", "No resulting string returned from transformation.", GPS_LOG_WARNING
                    End If
                Else
                    Log cModuleName, "transformMessage", "Unable to load workspace file.", GPS_LOG_WARNING
                End If
            Else
                Log cModuleName, "transformMessage", "Invalid Object type stored in Transforms collection.", GPS_LOG_WARNING
            End If
        Else
            Log cModuleName, "transformMessage", "No transformation specified...letting message pass.", GPS_LOG_INFORMATION
        End If
    Else
         Log cModuleName, "transformMessage", "No transformation specified...letting message pass.", GPS_LOG_INFORMATION
    End If
    Exit Function
err_transformMessage:
    Log cModuleName, "transformMessage", Err.Description, GPS_LOG_ERROR, Err.Number
End Function

Private Static Function HasAttributeValue(ByRef oNode As MSXML2.IXMLDOMNode, ByRef AttributeName As String) As String
Dim iNode As MSXML2.IXMLDOMNode

Set iNode = oNode.Attributes.getNamedItem(AttributeName)

HasAttributeValue = Not iNode Is Nothing

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
        Message = FindIdentities(Message)
    End If
    UnscreenMessage = Message
    Log cModuleName, "UnScreenMessage", "Message transformed from XML.", GPS_LOG_VERBOSE, , Message, GPS_SOURCE_STRING
End Function

Private Static Function ByteToHex(bytVal As Byte) As String
  ByteToHex = "00"
  Mid$(ByteToHex, 3 - Len(Hex$(bytVal))) = Hex$(bytVal)
End Function

Private Sub PropPageCallback_Exit()
    Set m_Property = Nothing
End Sub
Private Function FindEntityPlaceHolders(ByVal sXML As String, ByVal sEnt As String) As String
Dim ent As HTE_Entity.Entity
    
    If InStr(1, sXML, cEntityDevice, vbBinaryCompare) > 0 Then
        Set ent = New HTE_Entity.Entity
        ent.Deserialize sEnt
        sXML = cReplace.Replace(sXML, cEntityDevice, ent.DeviceID)
    End If
    
    If InStr(1, sXML, cEntityUnit, vbBinaryCompare) > 0 Then
        If ent Is Nothing Then
            Set ent = New HTE_Entity.Entity
            ent.Deserialize sEnt
        End If
        sXML = cReplace.Replace(sXML, cEntityUnit, ent.Alias)
    End If
    
    If InStr(1, sXML, cEntityStat, vbBinaryCompare) > 0 Then
        If ent Is Nothing Then
            Set ent = New HTE_Entity.Entity
            ent.Deserialize sEnt
        End If
        sXML = cReplace.Replace(sXML, cEntityStat, ent.StatusDescription)
    End If
    
    If InStr(1, sXML, cEntityCode, vbBinaryCompare) > 0 Then
        If ent Is Nothing Then
            Set ent = New HTE_Entity.Entity
            ent.Deserialize sEnt
        End If
        sXML = cReplace.Replace(sXML, cEntityCode, ent.statusCode)
    End If
    
    FindEntityPlaceHolders = sXML
End Function

Private Function FindIdentities(ByVal sXML As String) As String
Dim l_rc As Long
Dim s_ret As String
Const cMaxLen = 255
Dim i As Long
Const NOT_FOUND = "{6CD1925A-975E-4ae3-A2D5-2F8CB6B822E1}"
    
    For i = LBound(arrLookups, 2) To UBound(arrLookups, 2)
        If InStr(1, sXML, arrLookups(0, i), vbBinaryCompare) > 0 Then
            s_ret = String(cMaxLen, 0)
            l_rc = GetPrivateProfileString("client", arrLookups(1, i), NOT_FOUND, s_ret, cMaxLen, Environ("WINDIR") & "\hte_common.ini")
            If l_rc <> 0 Then
                s_ret = Left$(s_ret, l_rc)
                If (lstrcmpi(NOT_FOUND, s_ret) = 0) Then
                'If StrComp(NOT_FOUND, s_ret, vbBinaryCompare) = 0 Then
                    'NO VALUE FOUND REPLACE TAG WITH EMPTY STRING
                    sXML = cReplace.Replace(sXML, arrLookups(0, i), vbNullString)
                Else
                    sXML = cReplace.Replace(sXML, arrLookups(0, i), s_ret)
                End If
            Else
                'sometimes o/s does not return default value - we already know tag is in here replace with nothing
                sXML = cReplace.Replace(sXML, arrLookups(0, i), vbNullString)
            End If
        End If
    Next
    
    'find last time processed element
    If InStr(1, sXML, cElapsedTime, vbBinaryCompare) > 0 Then
        If Not elapsedTime Is Nothing Then
            sXML = cReplace.Replace(sXML, cElapsedTime, elapsedTime.Elapsed)
        Else
            sXML = cReplace.Replace(sXML, cElapsedTime, "86400")
        End If
    End If
    
    FindIdentities = sXML
    
End Function
Private Function ResolveEntityFromMessage(ByVal propData As HTE_GPS.GPSMessage) As HTE_Entity.Entity
Dim ent As HTE_Entity.Entity
Dim pos As Long
Dim offset As Long
Dim TAG As String, PROXY As String
Dim msgDesc As String
    If Len(propData.Entity) > 0 Then
        Set ent = New HTE_Entity.Entity
        If ent.Deserialize(propData.Entity) Then
            Log cModuleName, "ResolveEntityFromMessage", "Communication identification", GPS_LOG_VERBOSE, , propData.Entity, GPS_SOURCE_XML
            'order of operation look for - applicationID [device] and [alias] in message
            'For TAIP messages device {ID=xxxxxx;} alias = {UID=xxxxxx;}
            'For MapLite messages device {position 20 for 16} alias {position 46 for 32}
            'For NMEA - no identifier
            'resolve to key lookups/indicators cached from database
            'Server - always gets final say on [device] Client gets final say on [alias]
            msgDesc = UCase$(getMessageType(propData.Type).Desc)
            Select Case msgDesc  'not pretty, since we used to be dynamic but...
                Case "TAIP", "TRIMBLE"
                    'has device identifier?
                    TAG = ";ID="
                    pos = InStr(1, propData.rawMessage, TAG, vbBinaryCompare)
                    If pos > 0 Then
                        offset = InStr(pos + Len(TAG), propData.rawMessage, ";", vbBinaryCompare)
                        If offset > 0 Then
                            'another tag - alias/unit perhaps
                            ent.ApplicationID = Mid$(propData.rawMessage, pos + Len(TAG), offset - (pos + Len(TAG)))                                                                 'the length of ;
                        Else
                            ent.ApplicationID = Mid$(propData.rawMessage, pos + Len(TAG))
                        End If
                    End If
                    'has alias identifier - (modified with a non-standard identifier?)
                    TAG = ";UID="
                    pos = InStr(1, propData.rawMessage, TAG, vbBinaryCompare)
                    If pos > 0 Then
                        offset = InStr(pos + Len(TAG), propData.rawMessage, ";", vbBinaryCompare)
                        If offset > 0 Then
                            'another tag - alias/unit perhaps
                            ent.Alias = Mid$(propData.rawMessage, pos + Len(TAG), offset - (pos + Len(TAG)))
                        Else
                            ent.Alias = Mid$(propData.rawMessage, pos + Len(TAG))
                        End If
                    End If
                    TAG = ";STAT="
                    pos = InStr(1, propData.rawMessage, TAG, vbBinaryCompare)
                    If pos > 0 Then
                        offset = InStr(pos + Len(TAG), propData.rawMessage, ";", vbBinaryCompare)
                        If offset > 0 Then
                            'another tag - alias/unit perhaps
                            ent.StatusDescription = Mid$(propData.rawMessage, pos + Len(TAG), offset - (pos + Len(TAG)))
                        Else
                            ent.StatusDescription = Mid$(propData.rawMessage, pos + Len(TAG))
                        End If
                    End If
                    TAG = ";CODE="
                    pos = InStr(1, propData.rawMessage, TAG, vbBinaryCompare)
                    If pos > 0 Then
                        offset = InStr(pos + Len(TAG), propData.rawMessage, ";", vbBinaryCompare)
                        If offset > 0 Then
                            'another tag - alias/unit perhaps
                            ent.statusCode = Mid$(propData.rawMessage, pos + Len(TAG), offset - (pos + Len(TAG)))
                        Else
                            ent.statusCode = Mid$(propData.rawMessage, pos + Len(TAG))
                        End If
                    End If
                    'added for messages received from switch
                    TAG = ";PROXY="
                    pos = InStr(1, propData.rawMessage, TAG, vbBinaryCompare)
                    ent.IsProxy = pos > 0
                    If ent.IsProxy Then
                    On Local Error Resume Next
                        offset = InStr(pos + Len(TAG), propData.rawMessage, ";", vbBinaryCompare)
                        If offset > 0 Then
                            PROXY = Mid$(propData.rawMessage, pos + Len(TAG), offset - (pos + Len(TAG)))
                        Else
                            PROXY = Mid$(propData.rawMessage, pos + Len(TAG))
                        End If
                        Log cModuleName, "ResolveEntityFromMessage", "Datastream reconciled to proxy [" & PROXY & "]", GPS_LOG_VERBOSE
                    End If
                
                Case "MAPLITE", "GTG" 'TODO: TEST THIS FOR ACCURACY!!!
                    If Len(propData.rawMessage) >= 21 Then
                        ent.ApplicationID = Trim$(Mid$(propData.rawMessage, 21, 16)) '"DeviceID: Start 21 for 16
                    End If
                    If Len(propData.rawMessage) >= 50 Then
                        ent.Alias = Trim$(Mid$(propData.rawMessage, 50, 31)) 'UnitID/Alias: Start 50 for 32 -1 for F,P, or E
                    End If
                    If Len(propData.rawMessage) >= 189 Then
                        ent.StatusDescription = Trim$(Mid$(propData.rawMessage, 189)) 'should always be last part
                    End If
                Case "SGAVL"
                    Dim xmlObj As MSXML2.DOMDocument30
                    Dim element As IXMLDOMNodeList
                    Set xmlObj = LoadXML(propData.rawMessage)
                    
                    Set element = xmlObj.getElementsByTagName("CMP")
                    If (element.Length > 0) Then
                        ent.ApplicationID = Trim$(element.Item(0).Text)
                    Else
                        Set element = xmlObj.getElementsByTagName("GDV")
                        If (element.Length > 0) Then
                            ent.ApplicationID = Trim$(element.Item(0).Text)
                            Log cModuleName, "ResolveEntityFromMessage", "ApplicationID aquired from GDV element as oppose to CMP element.", GPS_LOG_VERBOSE
                        End If
                    End If
                    
                    Set element = xmlObj.getElementsByTagName("UNT")
                    If (element.Length > 0) Then
                        ent.Alias = Trim$(element.Item(0).Text)
                    End If
                    
                    Set xmlObj = Nothing
                    Set element = Nothing
                Case Else '"NMEA" and others
                
            End Select
            Log cModuleName, "ResolveEntityFromMessage", "MessageType [" & msgDesc & "] AppId [" & ent.ApplicationID & "] Alias [" & ent.Alias & "]", GPS_LOG_VERBOSE
        End If
        Set ResolveEntityFromMessage = ValidateEntityFromMessage(ent)
    End If
    
End Function

Private Function ValidateEntityFromMessage(ByRef ent As HTE_Entity.Entity) As HTE_Entity.Entity
Dim storedEnt As HTE_Entity.Entity 'server configuration
Dim lookup As String, bMutexObtained As Boolean
    '***MUTEX USAGE
    'being very defensive here - shared memory resource
On Error Resume Next
    'flag that critcial memory resource about to be altered - only "set" when db about to be cached!
    If bChangesPending Then 'only "lock" resource when certain db/memory about to be reset
        Do While Not processMutex.RequestMutex(10)
            If bRestarting Then
                Exit Function
            End If
        Loop
        bMutexObtained = True
    End If
    With ent
        'Look for ApplicationID - how we expect device to uniquely identify itself
        Set storedEnt = GetEntityByKey(applicationKeys, .ApplicationID)
        'modified Entity class to support CLONE function
        'the reason is if
        If Not storedEnt Is Nothing Then ReconcileEntity ent, storedEnt
        'still here - try MAC
        If storedEnt Is Nothing Then Set storedEnt = GetEntityByKey(deviceKeys, .MACAddress)
        If Not storedEnt Is Nothing Then ReconcileEntity ent, storedEnt
        'still here - pray for static IP
        If storedEnt Is Nothing Then Set storedEnt = GetEntityByKey(addressKeys, .IPAddress)
        If Not storedEnt Is Nothing Then ReconcileEntity ent, storedEnt
        'still nothing? you are having a bad day...
        If storedEnt Is Nothing Then
            'no updates - insert as new only IF
            'autoconfigure turned on
            Set storedEnt = New HTE_Entity.Entity
            ReconcileEntity ent, storedEnt
            If m_AutoSaveEntity Then
                'we cannot validate based on any information - no AppID, IP or MAC - try and save as new!!
                If Not m_aliasData.AddEntity(ent) Then
                    Log cModuleName, "ValidateEntityFromMessage", "Unable to preserve entity to datastore - it may already exist or may be a resource contention!", GPS_LOG_VERBOSE
                End If
            End If
            
        End If
        Set ValidateEntityFromMessage = storedEnt
    End With
exit_ValidateEntityFromMessage:
    If bMutexObtained Then processMutex.DiscardMutex
    Log cModuleName, "ValidateEntityFromMessage", "DataStore reconciliation", GPS_LOG_VERBOSE, , storedEnt.Serialize, GPS_SOURCE_XML
End Function

Private Sub ReconcileEntity(ByRef remoteEnt As HTE_Entity.Entity, ByRef localEnt As HTE_Entity.Entity)
    'update address with where it really came from
    'your database doesn't match reality
    With remoteEnt
        'communication verification doesn't lie
        If Not (lstrcmpi(.IPAddress, localEnt.IPAddress) = 0) Then localEnt.IPAddress = .IPAddress
        If Len(.MACAddress) > 0 Then
            If Not (lstrcmpi(.MACAddress, localEnt.MACAddress) = 0) Then localEnt.MACAddress = .MACAddress
        End If
        
        'if remote is smart enough to tell us who they are then update - client ALWAYS wins Aliasing
        If Len(.Alias) > 0 Then localEnt.Alias = .Alias
        'same with status since we don't/shouldn't interface with GW, Switch, AS/400 etc.
        If Len(.statusCode) > 0 Then localEnt.statusCode = .statusCode
        If Len(.StatusDescription) > 0 Then localEnt.StatusDescription = .StatusDescription
        'keep how the client identifies itself with the message
        localEnt.ApplicationID = remoteEnt.ApplicationID
        'ALWAYS USE SERVER "DEVICE" UNLESS IT'S EMPTY
        If Len(localEnt.DeviceID) = 0 Then 'nothing entered
            If Len(.ApplicationID) > 0 Then
                localEnt.DeviceID = .ApplicationID
            Else
                localEnt.DeviceID = "????"
            End If
        End If
        'newly added - proxy property for use in determining from centralized host <sigh>
        localEnt.IsProxy = remoteEnt.IsProxy
    End With

End Sub

Private Function GetEntityByKey(ByRef keyDictionary As Scripting.Dictionary, ByRef lookup As String, Optional ByVal Clone As Boolean = True) As HTE_Entity.Entity
Dim rtn As HTE_Entity.Entity
    If lookup <> vbNullString Then
        If Not keyDictionary Is Nothing Then
            With keyDictionary
                If .Exists(lookup) Then
                    If Not entityContainer Is Nothing Then
                        If entityContainer.Exists(.Item(lookup)) Then
                            'modified to use clone so that "cached" settings don't get persisted - alias, status
                            Set rtn = entityContainer.Item(.Item(lookup))
                            If Clone Then
                                Set GetEntityByKey = rtn.Clone(True)
                            Else
                                Set GetEntityByKey = rtn
                            End If
                            Log cModuleName, "GetEntityByKey", "Resolved by [" & lookup & "]", GPS_LOG_VERBOSE
                        End If
                    End If
                End If
            End With
        End If
    End If
End Function

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

Private Function Exists(ByRef m_Col As Collection, ByVal key As String) As Boolean
On Local Error Resume Next
    If IsObject(m_Col.Item(key)) Then
        Exists = (Err.Number = 0)
    Else
        Exists = (Err.Number = 0)
    End If
    
End Function

Private Sub clearTransforms()
    While m_Transforms.Count > 0: m_Transforms.Remove (1): Wend
    Set m_Transforms = New Collection
End Sub

Private Function LoadXML(ByVal xml As String) As MSXML2.DOMDocument30
    Dim oXML As MSXML2.DOMDocument30
    Dim objPI As IXMLDOMProcessingInstruction
    
    Set oXML = New DOMDocument30
    
    oXML.LoadXML (xml)
    
    Set LoadXML = oXML
    Set oXML = Nothing
End Function

Private Function getXML(Optional ByVal rootName As String = "root") As MSXML2.DOMDocument30
    Dim oXML As MSXML2.DOMDocument30
    Dim objPI As IXMLDOMProcessingInstruction
    Dim rootElement As IXMLDOMElement
    
        Set oXML = New DOMDocument30
        Set objPI = oXML.createProcessingInstruction("xml", "version='1.0'")
        oXML.appendChild objPI
        Set rootElement = oXML.createElement(rootName)
        Set oXML.documentElement = rootElement
        Set getXML = oXML
        Set objPI = Nothing
        Set rootElement = Nothing
        Set oXML = Nothing
        
End Function

Private Function IsXML(ByVal valueToExamine As String) As Boolean
    Dim objRegExp As RegExp
    Dim colMatches As MatchCollection
    
    Set objRegExp = New RegExp
    objRegExp.Pattern = "^<(\w+).*<\/(.*)>$"
    Set colMatches = objRegExp.Execute(valueToExamine)

    If colMatches.Count > 0 Then
        IsXML = True
    Else
        IsXML = False
    End If
    
End Function

Private Function AddNode(ByRef poXML As MSXML2.DOMDocument30, ByRef nodeName As String, _
            ByRef nodeValue As String, ByVal nodeType As MSXML2.DOMNodeType) As MSXML2.IXMLDOMElement
    Dim rootElement As IXMLDOMElement
    Dim aElement As IXMLDOMNode

    Set rootElement = poXML.documentElement
    Set aElement = poXML.createNode(nodeType, nodeName, vbNullString)
    aElement.nodeTypedValue = nodeValue
    rootElement.appendChild aElement
    Set AddNode = aElement
End Function
