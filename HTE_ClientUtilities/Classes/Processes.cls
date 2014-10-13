VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Processes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_AppObject As HTE_GPS.Application
Private m_Processes As Collection 'A collection of process interfaces keyed by Instance MemoryPointer
Private m_CoordKeys As Dictionary 'A dictionary item:Routing Instance key: Instance MemoryPointer
Private m_Route As Collection 'Routing path for each Item:A delimited string with the current path key: Number of Routes
Private m_Types() As HTE_GPS.GPSConfiguration
Private configSettings As HTE_ClientConfiguration.Configuration
Private hostStatus As HTE_GPS.GPS_HOST_STATUS
Private bCatastrophic As Boolean
Private bInitializing As Boolean
Private bRunning As Boolean
Private iTypesNode As MSXML2.IXMLDOMNode
Implements HTE_GPS.Processes
Private Const cModuleName = "HTE_ClientUtilities.Processes"
Public Event Status(ByVal statusCode As HTE_GPS.GPS_HOST_STATUS)
Private Const cUPDATE As String = "Updates"


Private Sub ConnectProcesses()

Dim objectTest As Object, processTest As HTE_GPS.Process
Dim iRoute As MSXML2.IXMLDOMNode, iNode As MSXML2.IXMLDOMNode, iAttribute As MSXML2.IXMLDOMNode
Dim instanceID As String, routePath As String, progID As String
Dim i As Long, x As Long
Dim currentType As RouteConfig, InMsg As HTE_GPS.GPS_MESSAGING_TYPES
Dim bServerRoute As Boolean
    'The main initialization routine, responsible for loading instances, asserting errors and process routing
    bInitializing = True
    bRunning = True
    Set iRoute = configSettings.GetNode(UCase$(App.EXEName))
    If Not iRoute Is Nothing Then
    For x = 0 To (iRoute.childNodes.length - 1)
        routePath = vbNullString
        Set iNode = iRoute.childNodes(x)
        i = 0
        If Not iNode Is Nothing Then
            Set iAttribute = iNode.Attributes.getNamedItem("INBOUNDMSG")
            If iAttribute Is Nothing Then
                'Log cModuleName, "ConnectProcesses", "Inbound route message configuration attribute NOT found, defaulting to TAIP!", GPS_LOG_WARNING
                Log cModuleName, "ConnectProcesses", "Inbound route message configuration attribute NOT found, defaulting to TYPE_0!", GPS_LOG_WARNING
                InMsg = GPS_TYPE_0 'GPS_TYPE_TAIP
            Else
                Log cModuleName, "ConnectProcesses", "Inbound route message configuration attribute found!", GPS_LOG_VERBOSE
                InMsg = iAttribute.nodeTypedValue
            End If
            currentType.InMsg = InMsg
            currentType.Path = routePath
            bServerRoute = False
            Set iAttribute = iNode.Attributes.getNamedItem("SERVER")
            If Not iAttribute Is Nothing Then
                bServerRoute = CBool(iAttribute.nodeTypedValue)
            End If
            m_Route.Add currentType, "Route_" & CStr(m_Route.Count) 'MAINTAIN PROCESS ROUTING PATH
            Log cModuleName, "ConnectProcesses", iNode.childNodes.length & " processes found at first glance!", GPS_LOG_INFORMATION
            Do While i <= iNode.childNodes.length - 1
                If iNode.childNodes(i).nodeType = NODE_ELEMENT Then 'DON'T TRY TO LOAD FROM COMMENTS!
                    progID = iNode.childNodes(i).nodeName 'NODE NAME IS PROGID
                    Log cModuleName, "ConnectProcesses", "Creating application '" & progID & "'", GPS_LOG_VERBOSE
On Error GoTo CreateError
                    Set objectTest = CreateObject(progID)
                    Log cModuleName, "ConnectProcesses", "Created instance of application '" & progID & "'", GPS_LOG_VERBOSE
On Error GoTo InterfaceError
                    Set processTest = objectTest
                    Log cModuleName, "ConnectProcesses", "[" & progID & "] supports current PROCESS interface.", GPS_LOG_VERBOSE
                    If Not m_AppObject.IsVersionAtLeast(processTest.HostVersion) Then
                        processTest.Status = GPS_STAT_HOST_UNSUPPORTED 'OLD HOST
                        Log cModuleName, "ConnectProcesses", "Version conflict between Process '" & progID & "' and Host Application", GPS_LOG_ERROR
                    End If
                    Log cModuleName, "ConnectProcesses", "[" & progID & "] supports current HOST process.", GPS_LOG_VERBOSE
On Error GoTo InstanceError
                    '***INSTANCE MUST BE UNIQUE ACROSS THE BOARD, THIS IS VERY IMPORTANT***
                    instanceID = iNode.childNodes(i).Attributes.getNamedItem("INSTANCE").nodeTypedValue
                    Log cModuleName, "ConnectProcesses", "[" & progID & "] instance identifier found - " & instanceID, GPS_LOG_VERBOSE
                    If instanceID <> vbNullString Then
On Error GoTo ConnectError
                        'NEED TO ADD TO ROUTING FIRST FOR POTENTIAL ROUTING BACK TO PROCESS
                        If bServerRoute And i = iNode.childNodes.length - 1 Then
                            routePath = m_Route.Item("Route_" & CStr(m_Route.Count - 1)).Path & instanceID & "|*|" & iNode.childNodes(0).Attributes.getNamedItem("INSTANCE").nodeTypedValue
                        Else
                            routePath = m_Route.Item("Route_" & CStr(m_Route.Count - 1)).Path & instanceID & "|*|"
                        End If
                        Log cModuleName, "ConnectProcesses", "Modifying route path : " & routePath, GPS_LOG_VERBOSE
                        m_Route.Remove "Route_" & CStr(m_Route.Count - 1)
                        currentType.InMsg = InMsg
                        currentType.Path = routePath
                        m_Route.Add currentType, "Route_" & CStr(m_Route.Count)
                        DoEvents
                        processTest.Connect Me, instanceID
                        GoTo LoopContinue
                    Else
                        Log cModuleName, "ConnectProcesses", "[" & progID & "] instance identifier found and is empty!", GPS_LOG_WARNING
                    End If
InstanceError:
                    bCatastrophic = True
                    Log cModuleName, "ConnectProcesses", "Error retrieving instance attribute for : '" & progID & "'.", GPS_LOG_ERROR
                    Resume LoopContinue
ConnectError:
                    'MAY NEED TO ADJUST THE ROUTE HERE SINCE WE DIDN'T ADD TO THE COLLECTION!!!!
                    bCatastrophic = True
                    Log cModuleName, "ConnectProcesses", "Error calling Connect on Process object: '" & progID & "'.", GPS_LOG_ERROR
                    Resume LoopContinue
CreateError:
                    bCatastrophic = True
                    Log cModuleName, "ConnectProcesses", "Unable to load Process: '" & progID & "'.", GPS_LOG_ERROR
                    Resume LoopContinue
InterfaceError:
                    bCatastrophic = True
                    Log cModuleName, "ConnectProcesses", "Process object, '" & progID & "', does not implement the proper interface.", GPS_LOG_ERROR
                    On Error Resume Next 'IN CASE THEY DON'T SUPPORT THE "STATUS" INTERFACE <BAD IMPLEMENTS>
                    processTest.Status = GPS_STAT_BAD_INTERFACE 'TRY TO TELL COMPONENT THAT THEY AREN'T UP TO SPEC
                    Resume LoopContinue
LoopContinue:
                    Set objectTest = Nothing
                    Set processTest = Nothing
               Else
                    Log cModuleName, "ConnectProcesses", "Node Type is invalid for process list.", GPS_LOG_WARNING
               End If
               i = i + 1
            Loop
        Else
            Log cModuleName, "ConnectProcesses", "No Processes exist in configuration file!", GPS_LOG_ERROR
        End If
    Next
    Else
        bCatastrophic = True
        Log cModuleName, "ConnectProcesses", "Unable to set configuration file!", GPS_LOG_ERROR
    End If
ConnectProcesses:
   On Error Resume Next
   bInitializing = False
   StatusChanged
   Exit Sub
End Sub

Private Sub DisconnectProcesses()
Dim thisProcess As HTE_GPS.Process
Dim vVar As Variant
    'MUST BE CALLED TO ENSURE PROPER CLEAN-UP
    If Not m_Processes Is Nothing Then
        For Each vVar In m_Processes
            Set thisProcess = vVar
            If Not thisProcess Is Nothing Then
                Log cModuleName, "DisconnectProcesses", "Disconnecting Process Type:" & getObjectProcess(m_CoordKeys(ObjectKey(thisProcess))) & " Instance: " & m_CoordKeys(ObjectKey(thisProcess)), GPS_LOG_VERBOSE
                If Not thisProcess Is Nothing Then thisProcess.Disconnect
            End If
        Next
    End If
    Set iTypesNode = Nothing
End Sub

Private Function ObjectKey(ByRef AnyObject As IUnknown) As String

    If Not AnyObject Is Nothing Then
        ObjectKey = "O_" & Trim$(ObjPtr(AnyObject))
    End If
End Function

Public Sub Load(ByVal HostApp As HTE_GPS.Application)
    Set m_AppObject = HostApp
    Set m_Processes = New Collection
    Set m_Route = New Collection
    Set m_CoordKeys = New Dictionary
    Set configSettings = New HTE_ClientConfiguration.Configuration
    Log cModuleName, "Load", "Checking for updates... ", GPS_LOG_VERBOSE
    If CheckForUpdates Then
        If UpdateConfiguration Then
            Log cModuleName, "Load", "Update file applied. ", GPS_LOG_VERBOSE
            On Error Resume Next
            Kill GetUpdateFile
        Else
            Log cModuleName, "Load", "Update failed. ", GPS_LOG_VERBOSE
        End If
    Else
        Log cModuleName, "Load", "No updates available at this time.", GPS_LOG_VERBOSE
    End If
    Log cModuleName, "Load", "Configuration file " & getConfigurationPath, GPS_LOG_VERBOSE
    configSettings.ConfigurationFile = getConfigurationPath
    LoadConfigTypes
    ConnectProcesses
End Sub

Private Function UpdateConfiguration() As Boolean
    On Error Resume Next
    Log cModuleName, "UpdateConfiguration", "Backing up file..."
    FileCopy getConfigurationPath, GetPath(getConfigurationPath) & "\" & Format$(GetUTCEx, "YYYY-MM-DD HHNNSS") & ".bak"
    Log cModuleName, "UpdateConfiguration", "Copying update..."
    UpdateConfiguration = FileCopy(GetUpdateFile, getConfigurationPath)
End Function

Private Function GetUpdateFile() As String
    GetUpdateFile = GetPath(getConfigurationPath) & _
        "\" & cUPDATE & "\" & StripPath(getConfigurationPath)
End Function

Private Function CheckForUpdates() As Boolean
    CheckForUpdates = FileExists(GetUpdateFile)
End Function

Private Function LoadConfigTypes() As Boolean
'Dim iNode As MSXML2.IXMLDOMNode
Dim ichildNode As MSXML2.IXMLDOMNode
Dim iAttribute As MSXML2.IXMLDOMNode
Dim i As Long
    'WHAT ARE THE CURRENT GPS DATA TYPES AND THEIR RESPECTIVE MEANINGS
    Set iTypesNode = configSettings.GetNode("GENERAL", "TYPES")
    If Not iTypesNode Is Nothing Then
        If iTypesNode.hasChildNodes Then
            For i = 0 To iTypesNode.childNodes.length - 1
                Set ichildNode = iTypesNode.childNodes(i)
                ReDim Preserve m_Types(i)
                With m_Types(i)
                    .GPSType = ichildNode.nodeTypedValue
                    Set iAttribute = ichildNode.Attributes.getNamedItem("SOM")
                    If Not iAttribute Is Nothing Then .SOM = formatTag(iAttribute.nodeTypedValue)
                    Set iAttribute = ichildNode.Attributes.getNamedItem("EOM")
                    If Not iAttribute Is Nothing Then .EOM = formatTag(iAttribute.nodeTypedValue)
                    .Desc = ichildNode.nodeName
                End With
            Next
            Log cModuleName, "LoadConfigTypes", "Configuration types found and loaded. Count = " & UBound(m_Types) + 1, GPS_LOG_VERBOSE
        Else
            Log cModuleName, "LoadConfigTypes", "No configuration types found under general settings!", GPS_LOG_ERROR
            bCatastrophic = True
        End If
    Else
        Log cModuleName, "LoadConfigTypes", "No configuration types found under general settings!", GPS_LOG_ERROR
        bCatastrophic = True
    End If
End Function

Public Property Get ProcessCount() As Long
    If m_Processes Is Nothing Then
        ProcessCount = 0
    Else
        ProcessCount = m_Processes.Count
    End If
End Property

Private Sub Log(ByVal Object As String, ByVal method As String, Optional ByVal Message As String = vbNullString, Optional ByVal LogDetail As GPS_LOG_DETAIL, Optional ByVal ErrorID As Long = 0, Optional ByVal LogSource As String = vbNullString, Optional ByVal LogSourceDetail As GPS_LOG_SOURCE)
    If m_AppObject Is Nothing Then Exit Sub
    If bRunning Then m_AppObject.Log Object, method, Message, LogDetail, ErrorID, LogSource, LogSourceDetail
End Sub

Public Sub Unload()
    'SINCE PROCESSES CLASS MAY ALREADY BE "TERMINATED", START LOGGING AND END IT IN THIS ROUTINE
    bRunning = False
    Log cModuleName, "Unload", "Disconnecting existing processes...", GPS_LOG_INFORMATION
    DisconnectProcesses
    Log cModuleName, "Unload", "Disconnected existing processes...", GPS_LOG_INFORMATION
    While m_Processes.Count > 0: m_Processes.Remove 1: Wend
    Set m_Processes = Nothing
    Log cModuleName, "Unload", "Processes cleared and unloaded.", GPS_LOG_INFORMATION
    While m_Route.Count > 0: m_Route.Remove 1: Wend
    Set m_Route = Nothing
    Log cModuleName, "Unload", "Routes cleared and unloaded.", GPS_LOG_INFORMATION
    m_CoordKeys.removeAll
    Set m_CoordKeys = Nothing
    Log cModuleName, "Unload", "CoordKeys cleared and unloaded.", GPS_LOG_INFORMATION
    Set configSettings = Nothing
    Log cModuleName, "Unload", "Configuration object unloaded", GPS_LOG_INFORMATION
    Set m_AppObject = Nothing
    Log cModuleName, "Unload", "Application pointer unloaded", GPS_LOG_INFORMATION
End Sub

Private Sub Class_Initialize()
    hostStatus = GPS_HOST_UNINITIALIZED
End Sub

Private Sub Class_Terminate()
    hostStatus = GPS_HOST_UNINITIALIZED
End Sub

Private Sub Processes_Connected(ActiveInstance As HTE_GPS.Process, ByVal cookie As String)
    
    If m_Processes Is Nothing Then Set m_Processes = New Collection
    If m_CoordKeys Is Nothing Then Set m_CoordKeys = New Dictionary
On Error GoTo err_Processes_Connected
    Log cModuleName, "Processes_Connected", "Instance " & cookie & " properly acknowledged connection request.", GPS_LOG_INFORMATION
    If Not m_CoordKeys.Exists(ObjectKey(ActiveInstance)) Then
        m_Processes.Add ActiveInstance, ObjectKey(ActiveInstance)
        Log cModuleName, "Processes_Connected", "Added instance " & cookie & " to processes as '" & ObjectKey(ActiveInstance) & "'", GPS_LOG_VERBOSE
        m_CoordKeys.Add ObjectKey(ActiveInstance), cookie
        Log cModuleName, "Processes_Connected", "Added instance " & cookie & " to coordination with key: " & ObjectKey(ActiveInstance) & " and data: " & cookie, GPS_LOG_VERBOSE
        Log cModuleName, "Processes_Connected", "Retrieving settings for instance " & cookie & ".", GPS_LOG_VERBOSE
        ActiveInstance.Settings = getInstanceSettings(cookie) 'Instance is unique so this is OK but routing/description need to be found
        Log cModuleName, "Processes_Connected", "Passed settings instance to process instance: " & cookie, GPS_LOG_VERBOSE, , ActiveInstance.Settings, GPS_SOURCE_XML
    End If
    Exit Sub

err_Processes_Connected:
    Log cModuleName, "Processes_Connected", Err.Description, GPS_LOG_ERROR, Err.Number
    
End Sub

Private Function Processes_InboundType(ActiveInstance As HTE_GPS.Process) As HTE_GPS.GPS_MESSAGING_TYPES
    If ActiveInstance Is Nothing Then Exit Function
    Processes_InboundType = getInboundFormat(m_CoordKeys(ObjectKey(ActiveInstance)))
End Function

Private Sub Processes_Log(ActiveInstance As HTE_GPS.Process, ByVal Object As String, ByVal method As String, Optional ByVal Message As String = "", Optional ByVal LogDetail As HTE_GPS.GPS_LOG_DETAIL = 8&, Optional ByVal ErrorID As Long = 0&, Optional ByVal LogSource As String = "", Optional ByVal LogSourceDetail As HTE_GPS.GPS_LOG_SOURCE = 0&)
    If ActiveInstance Is Nothing Then Exit Sub
    Log Object & " (" & ObjectKey(ActiveInstance) & " - " & m_CoordKeys(ObjectKey(ActiveInstance)) & ")", method, Message, LogDetail, ErrorID, LogSource, LogSourceDetail
End Sub
Private Function ManageQueue(ByRef Message As HTE_GPS.GPSMessage, ByVal NextProcess As String, Optional bStoreOrProcess As Boolean = True) As Boolean
Static qMSG() As HTE_GPS.GPSMessage
Static qPROC() As String
Static inQueue As Long
Dim i As Long
Dim oProcess As HTE_GPS.Process
    'WHILE INITIALIZING, SOME PROCESSES MAY RECEIVE INFORMATION, WE COULD DROP THE MESSAGE
    'SINCE ITS NON-CRITICAL INFORMATION, BUT ITS NOT REALLY A BIG DEAL TO SAVE IN MEMORY
    'AND SEND WHEN ALL PROCESSES ARE FINISHED INITIALIZING
    If bStoreOrProcess Then
        inQueue = inQueue + 1
        Log cModuleName, "ManageQueue", "Adding to queue message (" & inQueue & ") .", GPS_LOG_VERBOSE
        ReDim Preserve qMSG(inQueue - 1)
        ReDim Preserve qPROC(inQueue - 1)
        Set qMSG(inQueue - 1) = Message
        qPROC(inQueue - 1) = NextProcess
    Else
        If IsValidArray(qMSG) Then
        For i = LBound(qMSG) To UBound(qMSG)
            Log cModuleName, "ManageQueue", "Drawing from queue message (" & CStr(i + 1) & ").", GPS_LOG_VERBOSE
On Error Resume Next
            Set oProcess = m_Processes.Item(qPROC(i))
            If Err.Number = 0 Then
                If Not oProcess Is Nothing Then oProcess.ProcessMessage qMSG(i)
            Else
                Log cModuleName, "ManageQueue", "Can't move message, dropping.", GPS_LOG_VERBOSE
            End If
        Next i
        End If
        Log cModuleName, "ManageQueue", "Clearing queue.", GPS_LOG_VERBOSE
        Erase qMSG
        Erase qPROC
        ManageQueue = True
    End If
End Function

Public Function ProcessStatus(ByVal CurrentRoute As String) As HTE_GPS.GPS_PROCESSOR_STATUS
    ProcessStatus = Processes_ProcessStatus(CurrentRoute)
End Function

Private Sub Processes_ProcessComplete(ByVal Message As HTE_GPS.GPSMessage, ActiveInstance As HTE_GPS.Process)
'ROUTE TO THE NEXT AVAILABLE PROCESS
Dim NextProcess As HTE_GPS.Process
Dim sNext As String
Dim i As Long
Static bRemoved As Boolean
On Error GoTo err_ProcessComplete
    
    If ActiveInstance Is Nothing Then Exit Sub
    
    Log cModuleName, "Processes_ProcessComplete", "Received message from " & ObjectKey(ActiveInstance) & " - (" & m_CoordKeys(ObjectKey(ActiveInstance)) & ").", GPS_LOG_VERBOSE
    If Not m_AppObject Is Nothing Then m_AppObject.Processed Message, getObjectProcess(m_CoordKeys(ObjectKey(ActiveInstance))) & " - " & m_CoordKeys(ObjectKey(ActiveInstance))
    If Message.MessageStatus <> GPS_MSG_ERROR Then
        If bRunning Then
            Log cModuleName, "Processes_ProcessComplete", "Message Status is " & messageStatusDesc(Message.MessageStatus), GPS_LOG_VERBOSE, , Message.rawMessage, GPS_SOURCE_BINARY
            If Not m_CoordKeys Is Nothing Then
                If ActiveInstance Is Nothing Then Exit Sub
                sNext = getNextProcess(m_CoordKeys(ObjectKey(ActiveInstance)))
                If sNext <> vbNullString Then
                    Log cModuleName, "Processes_ProcessComplete", "Forwarding message to " & sNext, GPS_LOG_VERBOSE
                    For i = 0 To m_CoordKeys.Count - 1 'Using a dictionary because we need access to keys as well as values, for routing!!
                        If StrComp(sNext, m_CoordKeys(m_CoordKeys.Keys(i)), vbTextCompare) = 0 Then
                            sNext = m_CoordKeys.Keys(i)
                            Exit For
                        End If
                    Next
                    If bInitializing Then
                        'IF WE HAVEN'T FINISHED INITIALIZING SAVE MESSAGE OFF AND PROCESS LATER
                        Log cModuleName, "Processes_ProcessComplete", "Processes still initializing storing for later.", GPS_LOG_INFORMATION
                        ManageQueue Message, sNext
                    Else
                        On Error GoTo err_ProcessNext
                        Set NextProcess = m_Processes.Item(sNext)
                        If Not bRemoved Then
                            'PROCESS SAVED MESSAGES FIRST...
                            bRemoved = True
                            ManageQueue Message, sNext, False
                        End If
                        NextProcess.ProcessMessage Message
                    End If
                Else
                    Log cModuleName, "Processes_ProcessComplete", "End of route...", GPS_LOG_VERBOSE
                End If
            End If
        Else
            Log cModuleName, "Processes_ProcessComplete", "Processes is shutting down, dropping message.", GPS_LOG_WARNING, , Message.rawMessage, GPS_SOURCE_BINARY
        End If
    Else
        Log cModuleName, "Processes_ProcessComplete", "Message received is in error, dropping message from delivery.", GPS_LOG_WARNING, , Message.rawMessage, GPS_SOURCE_BINARY
    End If
    Exit Sub
err_ProcessNext:
    Log cModuleName, "Processes_ProcessComplete", "Unable to locate process for " & sNext & ". Process may not have connected yet and is otherwise unavailable. " & Err.Description, GPS_LOG_ERROR, Err.Number
    Exit Sub
err_ProcessComplete:
    Log cModuleName, "Processes_ProcessComplete", Err.Description, GPS_LOG_ERROR, Err.Number
End Sub

Private Function Processes_ProcessStatus(ByVal CurrentRoute As String) As HTE_GPS.GPS_PROCESSOR_STATUS
Dim vItems As Variant
Dim i As Long
    Processes_ProcessStatus = GPS_STAT_UNKNOWN
    With m_CoordKeys
        For i = 0 To .Count - 1
            If StrComp(.Items(i), CurrentRoute, vbTextCompare) = 0 Then
                On Local Error Resume Next
                If Not m_Processes.Item(.Keys(i)) Is Nothing Then
                    Processes_ProcessStatus = m_Processes.Item(.Keys(i)).Status
                Else
                    Processes_ProcessStatus = GPS_STAT_UNINITIALIZED
                End If
                Exit For
            End If
        Next
    End With
End Function

Private Sub Processes_PropertiesChanged(XMLDOMNode As String, ActiveInstance As HTE_GPS.Process)
'NEED TO PERSIST INSTANCE SETTING CHANGES
Dim oDOM As DOMDocument30
Dim sProcess As String
Dim iNode As MSXML2.IXMLDOMNode
    Set oDOM = New DOMDocument30
    If oDOM.loadXML(XMLDOMNode) Then
        If ActiveInstance Is Nothing Then Exit Sub
        sProcess = m_CoordKeys(ObjectKey(ActiveInstance))
        Log cModuleName, "Process_PropertiesChanged", "Persisting settings for ActiveInstance.", GPS_LOG_VERBOSE, , XMLDOMNode, GPS_SOURCE_XML
        'Remove message types configuration from settings...don't save duplicate info
        Set oDOM.documentElement = removeTypesConfiguration(oDOM.documentElement)
        configSettings.LetNode oDOM.documentElement, getObjectProcess(sProcess)
    Else
        Log cModuleName, "Process_PropertiesChanged", "Unable to persist settings, format is incorrect!", GPS_LOG_WARNING, , XMLDOMNode, GPS_SOURCE_XML
    End If
End Sub

Public Function PropertyPage(ByVal CurrentRoute As String) As HTE_GPS.PropertyPage
    Set PropertyPage = Processes_PropertyPage(CurrentRoute)
End Function

Private Function Processes_PropertyPage(ByVal CurrentRoute As String) As HTE_GPS.PropertyPage
Dim vItems As Variant
Dim i As Long
    With m_CoordKeys
        For i = 0 To .Count - 1
            If StrComp(.Items(i), CurrentRoute, vbTextCompare) = 0 Then
                On Local Error Resume Next
                If Not m_Processes.Item(.Keys(i)) Is Nothing Then
                    Set Processes_PropertyPage = m_Processes.Item(.Keys(i)).PropertyPage
                End If
                Exit For
            End If
        Next
    End With
End Function

Private Sub Processes_StatusChanged(ByVal statusCode As HTE_GPS.GPS_PROCESSOR_STATUS, ActiveInstance As HTE_GPS.Process)
'NEED TO NOTIFY HOST APPLICATION OF OUR STATE
    Log cModuleName, "Processes_StatusChanged", "StatusCode = " & processorStatusDesc(statusCode) & " for instance " & m_CoordKeys.Item(ObjectKey(ActiveInstance)), GPS_LOG_VERBOSE
    If Not ActiveInstance Is Nothing Then StatusChanged
End Sub

Private Sub StatusChanged()
Dim thisOne As HTE_GPS.Process
Dim localStat As HTE_GPS.GPS_PROCESSOR_STATUS
Dim i As Long
    localStat = GPS_STAT_READYANDWILLING
    For i = 1 To m_Processes.Count
        Set thisOne = m_Processes.Item(i)
        If Not thisOne Is Nothing Then
            If thisOne.Status < localStat Then localStat = thisOne.Status
        End If
    Next
    Log cModuleName, "Processes_StatusChanged", "Local processes Status code considered " & processorStatusDesc(localStat), GPS_LOG_INFORMATION
    If bCatastrophic Then ' ERROR DURING INITIALIZATION
        hostStatus = GPS_HOST_ERROR
        If m_Processes.Count = 0 Then hostStatus = GPS_HOST_UNINITIALIZED
    Else
        Select Case localStat
            Case GPS_STAT_BAD_INTERFACE, GPS_STAT_ERROR, GPS_STAT_HOST_UNSUPPORTED, GPS_STAT_UNKNOWN
                hostStatus = GPS_HOST_ERROR
            Case GPS_STAT_READYANDWILLING
                hostStatus = GPS_HOST_GROOVY
            Case GPS_STAT_UNINITIALIZED
                hostStatus = GPS_HOST_UNINITIALIZED
            Case GPS_STAT_WARNING, GPS_STAT_INITIALIZED
                hostStatus = GPS_HOST_WARNING
        End Select
        If m_Processes.Count = 0 Then hostStatus = GPS_HOST_UNINITIALIZED
    End If
    Log cModuleName, "Processes_StatusChanged", "Host Status Code considered as " & hostStatusDesc(hostStatus), GPS_LOG_INFORMATION
    If Not m_AppObject Is Nothing Then m_AppObject.StatusChange hostStatus
End Sub

Private Function getInstanceSettings(ByVal processInstance As String) As String
'NEED TO PASS INDIVIDUAL SETTINGS HERE
Dim iNode As MSXML2.IXMLDOMNode
Dim iDOM As MSXML2.DOMDocument30
    Set iNode = configSettings.GetNode(processInstance)
    If Not iNode Is Nothing Then
        Set iNode = addTypesConfiguration(iNode)
        getInstanceSettings = iNode.xml
    Else
        'configuration defaults I suppose
        Set iDOM = getXML(processInstance)
        Set iDOM.documentElement = addTypesConfiguration(iDOM.documentElement)
        getInstanceSettings = iDOM.xml
    End If
End Function

Private Function addTypesConfiguration(ByRef iNode As MSXML2.IXMLDOMNode) As MSXML2.IXMLDOMNode
'Add message types configuration to settings, in case helper pages need to refer to...
'work off of 'clone' so as not to disrupt general settings
    If Not iTypesNode Is Nothing Then
        Set iNode = removeTypesConfiguration(iNode)
        iNode.appendChild iTypesNode.cloneNode(True)
        Set addTypesConfiguration = iNode
    End If
End Function

Private Function removeTypesConfiguration(ByRef iNode As MSXML2.IXMLDOMNode) As MSXML2.IXMLDOMNode
Dim oDOM As MSXML2.DOMDocument30
Dim iChild As MSXML2.IXMLDOMNode
    If Not iTypesNode Is Nothing Then
        Set oDOM = getXML
        If oDOM.loadXML(iNode.xml) Then
            Do
                Set iChild = oDOM.documentElement.selectSingleNode(iTypesNode.nodeName)
                If Not iChild Is Nothing Then oDOM.documentElement.removeChild iChild
            Loop While Not iChild Is Nothing
        End If
        Set iNode = oDOM.documentElement
        Set removeTypesConfiguration = iNode
    End If
End Function

Private Function getXML(Optional ByVal rootName As String = "configuration") As MSXML2.DOMDocument30
    Dim oXML As MSXML2.DOMDocument30
    Dim objPI As IXMLDOMProcessingInstruction
    Dim rootElement As IXMLDOMElement
    Dim newAtt As IXMLDOMAttribute
    Dim namedNodeMap As IXMLDOMNamedNodeMap

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

Private Function getNextProcess(ByVal processInstance As String) As String
'RETURNS THE NEXT PROCESS IN THE CHAIN, WHERE TO ROUTE THE MESSAGE TO!
Dim i As Long, pos As Long
Dim routePath As String
    'BASED ON PROCESS THAT JUST CALLED BACK, WHAT IS THE NEXT PROCESS IN THE LIST IF ANY
    If Not m_Route Is Nothing Then
        For i = 1 To m_Route.Count
            routePath = m_Route.Item(i).Path
            If routePath <> vbNullString Then
                pos = InStr(1, routePath, processInstance)
                If pos > 0 Then
                    Log cModuleName, "getNextProcess", "Route path found as " & routePath & ".", GPS_LOG_VERBOSE
                    routePath = Mid$(routePath, pos + Len(processInstance) + Len("|*|"))
                    If routePath <> vbNullString Then
                        getNextProcess = Split(routePath, "|*|")(0)
                        Exit Function
                    End If
                End If
            End If
        Next
    End If
End Function

Private Function getObjectProcess(ByVal processInstance As String) As String
'RETURNS THE ORIGINATING PROGID OF THE CALLING PROCESS, IMPORTANT FOR RETRIEVING SETTINGS!
Dim iNode As MSXML2.IXMLDOMNode
Dim routePath As String
Dim i As Long, pos As Long
Dim x As Long
    If Not m_Route Is Nothing Then
        configSettings.ConfigurationFile = getConfigurationPath
        For i = 1 To m_Route.Count
            routePath = m_Route.Item(i).Path
            If routePath <> vbNullString Then
                pos = InStr(1, routePath, processInstance)
                If pos > 0 Then
                    Set iNode = configSettings.GetNode(processInstance)
                    If Not iNode Is Nothing Then
                        getObjectProcess = iNode.parentNode.nodeName
                        Exit Function
                    End If
                End If
            End If
        Next
        
        'Made it this far didn't find Object Process in configuration file,
        'create empty settings with getXML and allow helper to run off of its defaults!!!
        
        'NO PROPERTIES IN FILE, RUNNING OFF OF DEFAULTS FIND ANOTHER WAY
        Set iNode = configSettings.GetNode(UCase$(App.EXEName))
        If Not iNode Is Nothing Then
            If iNode.hasChildNodes Then
                For x = 0 To iNode.childNodes.length - 1
                    If iNode.childNodes(x).nodeName <> "#text" Then
                        If iNode.childNodes(x).hasChildNodes Then
                            For i = 0 To iNode.childNodes(x).childNodes.length - 1
                                If Not iNode.childNodes(x).childNodes(i).Attributes.getNamedItem("INSTANCE") Is Nothing Then
                                    If StrComp(processInstance, iNode.childNodes(x).childNodes(i).Attributes.getNamedItem("INSTANCE").nodeTypedValue, vbTextCompare) = 0 Then
                                        getObjectProcess = Replace(iNode.childNodes(x).childNodes(i).nodeName, ".process", vbNullString, , , vbTextCompare)
                                        Exit Function
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            End If
        End If
    End If
End Function

Private Function getInboundFormat(ByVal processInstance As String) As HTE_GPS.GPS_MESSAGING_TYPES
Dim routePath As String
Dim i As Long, pos As Long
    'WHAT IS THE INBOUND MESSAGE FORMAT FOR THE ROUTE THE PROCESS IS PARTICIPATING IN?
    If Not m_Route Is Nothing Then
        For i = 1 To m_Route.Count
            routePath = m_Route.Item(i).Path
            If routePath <> vbNullString Then
                pos = InStr(1, routePath, processInstance)
                If pos > 0 Then
                    getInboundFormat = m_Route.Item(i).InMsg
                    Exit Function
                End If
            End If
        Next
    End If

End Function

Private Function Processes_SupportedMessages() As HTE_GPS.GPSConfiguration()
    'SAFEARRAY OF GPSConfiguration INFORMATION - NEEDED FOR SOME PROCESSES
    Processes_SupportedMessages = m_Types
End Function
