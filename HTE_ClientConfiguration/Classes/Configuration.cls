VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Configuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim oConfig As MSXML2.DOMDocument30
Dim m_Path As String
Dim m_Root As String
Private Const cModuleName = "Configuration"
Private Const OFS_MAXPATHNAME = 128
Private Const OF_EXIST = &H4000
Private Const OF_READ = &H0
Private Type OFSTRUCT
   cBytes As Byte
   fFixedDisk As Byte
   nErrCode As Integer
   Reserved1 As Integer
   Reserved2 As Integer
   szPathName(OFS_MAXPATHNAME) As Byte
End Type
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private WithEvents Subscribe As HTE_PubData.Subscriber
Attribute Subscribe.VB_VarHelpID = -1
Private Publish As HTE_PubData.Publisher
Private Const HTE_TOPIC = "GPS_Configuration"
Private Const HTE_EVENT = "DataChanged"
Private dblFileDateTime As Double
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function FileExists(ByVal strSearchFile As String) As Boolean
    Dim strucFname As OFSTRUCT
    FileExists = (OpenFile(strSearchFile, strucFname, OF_EXIST) <> -1)
End Function

Public Property Let ConfigurationFile(ByVal fileName As String)
    m_Path = fileName
    While Not GetConfigurationFile
        Sleep 250
    Wend
End Property

Public Property Get ConfigurationFile() As String
    ConfigurationFile = m_Path
End Property

Public Property Let rootName(ByVal sRootName As String)
    m_Root = sRootName
End Property

Public Property Get rootName() As String
    If Not oConfig Is Nothing Then
        rootName = oConfig.documentElement.nodeName
    Else
        rootName = m_Root
    End If
End Property

Public Sub Save()
Dim sFile As String

On Error GoTo err_Save
    If Not oConfig Is Nothing Then
        If oConfig.url <> vbNullString Then
            oConfig.Save oConfig.url
            sFile = oConfig.url
        ElseIf ConfigurationFile <> vbNullString Then
            oConfig.Save ConfigurationFile
            sFile = ConfigurationFile
        End If
        If Not Publish Is Nothing Then Publish.SendString HTE_EVENT, sFile
    End If
    UEH_Log cModuleName, "Save", "Configuration saved!", logInformation
    Exit Sub
err_Save:
    UEH_LogError cModuleName, "Save", Err
    Err.Clear
On Error Resume Next
    
    oConfig.Save sFile
    If Err.Number <> 0 Then
        UEH_LogError cModuleName, "Save_Retry", Err
    Else
        UEH_Log cModuleName, "Save", "Retry succeeded!", logInformation
    End If

End Sub

Private Function GetConfigurationFile() As Boolean

On Error GoTo err_GetConfigurationFile
    
    If oConfig Is Nothing Then Set oConfig = getXML(rootName)
    oConfig.async = False
    If Not FileExists(m_Path) Then
        UEH_Log cModuleName, "GetConfigurationFile", "Configuration file not found, creating from scratch.", logWarning, , "Path = " & m_Path, logString
        oConfig.Save m_Path
        GetConfigurationFile = True
        dblFileDateTime = GetFileDateTime(m_Path)
    Else
        If GetFileDateTime(m_Path) > dblFileDateTime Then
            If Not oConfig.Load(m_Path) Then
                UEH_Log cModuleName, "GetConfigurationFile", "Configuration file is an invalid xml file.", logWarning, , "Path = " & m_Path, logString
                GetConfigurationFile = False
            Else
                UEH_Log cModuleName, "GetConfigurationFile", "Configuration file loaded.", logInformation
                GetConfigurationFile = True
                dblFileDateTime = GetFileDateTime(m_Path)
            End If
        Else
            GetConfigurationFile = True
            UEH_Log cModuleName, "GetConfigurationFile", "Configuration is already current."
        End If
    End If
    
    Exit Function

err_GetConfigurationFile:
    UEH_LogError cModuleName, "GetConfigurationFile", Err
End Function

Public Function GetNode(ByRef nodeName As String, Optional ByRef childNode As String) As MSXML2.IXMLDOMNode
Dim aElement As MSXML2.IXMLDOMNode
    If Not oConfig Is Nothing And nodeName <> vbNullString Then
        If oConfig.hasChildNodes Then
            Set aElement = oConfig.getElementsByTagName(nodeName).Item(0)
            If Not aElement Is Nothing Then
                If childNode <> vbNullString Then
                    Set GetNode = aElement.selectSingleNode(childNode)
                Else
                    Set GetNode = aElement
                End If
            Else
                UEH_Log cModuleName, "GetNode", "Element " & nodeName & " not found.", logWarning
            End If
        Else
            UEH_Log cModuleName, "GetNode", "No child nodes available for retrieval.", logWarning
        End If
    Else
        UEH_Log cModuleName, "GetNode", "Configuration file NOT loaded use Configuration File Property SET before attempting to retrieve a Node.", logWarning
    End If
    Exit Function
err_GetNode:
    UEH_Log cModuleName, "GetNode", Err.Description, logError, Err.Number
End Function

Public Function LetNode(Node As MSXML2.IXMLDOMNode, ByRef nodeName As String, Optional ByRef childNode As String) As Boolean
Dim aElement As MSXML2.IXMLDOMNode
Dim rootElement As IXMLDOMElement
    If Not oConfig Is Nothing And nodeName <> vbNullString Then
        Set aElement = GetNode(nodeName, vbNullString)
        If Not aElement Is Nothing Then
            If Not Node Is Nothing Then
                Set rootElement = aElement.selectNodes(Node.nodeName).Item(0)
                If Not rootElement Is Nothing Then
                    aElement.removeChild rootElement
                End If
                aElement.appendChild Node
''                its more dangerous to save too often!
''                Save
            Else
                If childNode <> vbNullString Then
                    Set Node = oConfig.createNode(NODE_ELEMENT, childNode, vbNullString)
                    Set Node = aElement.appendChild(Node)
                Else
                    Set Node = aElement
                End If
            End If
        Else
            UEH_Log cModuleName, "GetNode", "Element " & nodeName & " does not currently exist!", logVerbose
            If oConfig.documentElement Is Nothing Then Set oConfig = getXML 'not a valid XML document
            Set aElement = oConfig.createNode(NODE_ELEMENT, nodeName, vbNullString)
            If childNode <> vbNullString Then
                Set Node = oConfig.documentElement.appendChild(aElement)
                Set Node = Node.appendChild(oConfig.createNode(NODE_ELEMENT, childNode, vbNullString))
            Else
                Set Node = oConfig.documentElement.appendChild(aElement)
            End If
        End If
        Save
        LetNode = True
    Else
        UEH_Log cModuleName, "GetNode", "Configuration file NOT loaded use Configuration File Property SET before attempting to let a Node.", logWarning
    End If
    
End Function
Public Function addAttribute(ByRef aElement As MSXML2.IXMLDOMNode, ByVal attrName As String, ByVal attrValue As String) As MSXML2.IXMLDOMNode
Dim newAtt As IXMLDOMAttribute
Dim namedNodeMap As IXMLDOMNamedNodeMap
    If Not oConfig Is Nothing Then
        Set newAtt = oConfig.createAttribute(attrName)
        newAtt.nodeTypedValue = attrValue
        Set namedNodeMap = aElement.Attributes
        Set addAttribute = namedNodeMap.setNamedItem(newAtt)
    End If
End Function
Public Function AddChildNode(ByRef childNode As MSXML2.IXMLDOMElement, ByRef nodeName As String, _
            ByRef nodeValue As String, ByVal nodeType As MSXML2.DOMNodeType, Optional nodeAttribute As String = vbNullString, Optional nodeAttributeValue As String = vbNullString, Optional ByVal MIME As Boolean = False) As MSXML2.IXMLDOMElement
    
    Dim aElement As Object
    
    If Not oConfig Is Nothing Then
        Set aElement = oConfig.createNode(nodeType, nodeName, vbNullString)
        If nodeAttribute <> vbNullString Then addAttribute aElement, nodeAttribute, nodeAttributeValue
        If MIME Then
            aElement.dataType = "bin.base64"
            aElement.nodeTypedValue = ByteArrayFromString(nodeValue)
        Else
            aElement.nodeTypedValue = nodeValue
        End If
        Set AddChildNode = childNode.appendChild(aElement)
        Save
        
    End If
End Function
Public Function ByteArrayFromString(ByVal Source$) As Variant
   Dim Buf() As Byte
   Dim r$
   r$ = StrConv(Source$, vbFromUnicode)
   Buf() = r$
   ByteArrayFromString = Buf()
End Function

Public Function StringFromByteArray(vr As Variant) As String
   Dim Buf() As Byte
   Dim r$
   Buf() = vr
   r$ = Buf()

   StringFromByteArray = StrConv(r$, vbUnicode)
End Function

Private Function getXML(Optional ByVal rootName As String = "configuration") As MSXML2.DOMDocument30
    Dim oXML As MSXML2.DOMDocument30
    Dim objPI As IXMLDOMProcessingInstruction
    Dim rootElement As IXMLDOMElement
    Dim newAtt As IXMLDOMAttribute
    Dim namedNodeMap As IXMLDOMNamedNodeMap

        Set oXML = New DOMDocument30
        oXML.async = False
        Set objPI = oXML.createProcessingInstruction("xml", "version='1.0'")
        oXML.appendChild objPI
        Set rootElement = oXML.createElement(rootName)
        Set oXML.documentElement = rootElement
        Set getXML = oXML
        Set objPI = Nothing
        Set rootElement = Nothing
        Set oXML = Nothing
        
End Function

Private Sub Class_Initialize()
    UEH_BeginLogging App.EXEName, App.EXEName
    rootName = "configuration"
    Set Publish = New HTE_PubData.Publisher
    With Publish
        .Topic = HTE_TOPIC
        .Timeout = 15000
    End With
    Set Subscribe = New HTE_PubData.Subscriber
    With Subscribe
        .Topic = HTE_TOPIC
    End With
End Sub

Private Sub Class_Terminate()
    Set Subscribe = Nothing
    Set Publish = Nothing
    UEH_EndLogging
End Sub

Private Sub Subscribe_OnReceiveString(ByVal Tag As String, ByVal Data As String)
    Select Case Tag
        Case HTE_EVENT
            If StrComp(Data, ConfigurationFile, vbTextCompare) = 0 Then 'same file
                If FileExists(Data) Then 'fileexists(configurationfile)
                    If GetFileDateTime(Data) > dblFileDateTime Then
                        While Not GetConfigurationFile
                            'reload - somebody changed it
                            Sleep 300
                        Wend
                    End If
                End If
            End If
    End Select
End Sub

Private Function GetFileDateTime(fileName As String) As Double
Dim hHandle As Long
Dim FileTimeType1 As FILETIME, FileTimeType2 As FILETIME
Dim SysTime As SYSTEMTIME
Dim OF As OFSTRUCT
Dim chk As Integer
Dim rtn As Variant
    
    If FileExists(fileName) Then
        hHandle = OpenFile(fileName, OF, OF_READ)
        GetFileTime hHandle, FileTimeType1, FileTimeType1, FileTimeType2
        FileTimeToLocalFileTime FileTimeType2, FileTimeType1
        FileTimeToSystemTime FileTimeType1, SysTime
        CloseHandle hHandle
        With SysTime
            chk = .wHour
            Select Case True
                Case chk > 0
                    rtn = (1 * chk) / 24
                Case chk < 0
                    rtn = ((1 * chk) / 24) + 1
                Case Else
                    rtn = 0
            End Select
            If .wMinute > 0 Then
                rtn = rtn + ((1 * .wMinute) / 1440)
            End If
            If .wSecond > 0 Then
                rtn = rtn + ((1 * .wSecond) / 86400)
            End If
            If .wMilliseconds > 0 Then
                rtn = rtn + ((1 * .wMilliseconds) / 86400000)
            End If
            GetFileDateTime = CDec(CDbl(DateSerial(.wYear, .wMonth, .wDay)) + rtn)
        End With
    End If

End Function
