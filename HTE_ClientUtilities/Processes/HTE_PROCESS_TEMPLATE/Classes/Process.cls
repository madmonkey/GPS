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
Private Const cModuleName = "HTE_Process_Template.Process"
Private gApp As HTE_GPS.Processes
Private m_Status As HTE_GPS.GPS_PROCESSOR_STATUS
Private localSettings As MSXML2.DOMDocument30
Private m_Settings As String
Private m_Property As Object
Implements HTE_GPS.Process
Implements HTE_GPS.PropPageCallback

Private Sub Class_Initialize()
    m_Status = GPS_STAT_INITIALIZED
    Set localSettings = New MSXML2.DOMDocument30
End Sub

Private Sub Class_Terminate()
    m_Status = GPS_STAT_UNINITIALIZED
    Set localSettings = Nothing
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
    '*TODO - change name to reflect your project
    Process_FriendlyName = "GPS Template Process"
End Property

Private Property Get Process_HostVersion() As String
    Process_HostVersion = "1.0.0"
End Property

Private Sub Process_ProcessMessage(ByVal PropData As HTE_GPS.GPSMessage)
'*TODO - Write the code that your process will perform to process the message
    '*NOTE - if you are a bridge (netport - via TCP, UDP etc..) you will most likely send a
    '*NOTE - message to the remote piece instead of raising a process complete event
    '*NOTE - your processComplete would be raised WHEN you recieved a message via the socket etc.
    gApp.ProcessComplete PropData, Me
End Sub

'''Private Sub Process_ProcessMessage(PropData As HTE_GPS.GPSMessage)
'''    '*TODO - Write the code that your process will perform to process the message
'''    '*NOTE - if you are a bridge (netport - via TCP, UDP etc..) you will most likely send a
'''    '*NOTE - message to the remote piece instead of raising a process complete event
'''    '*NOTE - your processComplete would be raised WHEN you recieved a message via the socket etc.
'''    gApp.ProcessComplete PropData, Me
'''End Sub

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
        Process_Status = GPS_STAT_READYANDWILLING
    Else
        Process_Status = GPS_STAT_UNKNOWN
    End If
    Exit Sub
err_setProperties:
    Log cModuleName, "setProperties", Err.Description, GPS_LOG_ERROR, Err.Number
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
