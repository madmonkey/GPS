VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_PrivateCollection As Collection
Dim m_ShowInTray As Boolean
Dim m_AllowExit As Boolean
Dim iTypesNode As MSXML2.IXMLDOMNode
Dim m_Config As HTE_ClientConfiguration.Configuration
Private Const cConfigProcessesHeading = "HTE_CLIENTUTILITIES"
Private bDirty As Boolean

Private Sub Class_Initialize()
    Set m_PrivateCollection = New Collection
    Set m_Config = New HTE_ClientConfiguration.Configuration
    m_Config.ConfigurationFile = getConfigurationPath
    initAppSettings
    Set iTypesNode = typesNode
    AddFromNode m_Config.GetNode(cConfigProcessesHeading)
End Sub
Public Sub Reload()
    'Set m_PrivateCollection = New Collection
    'Set m_Config = New HTE_ClientConfiguration.Configuration
    Clear
    AddFromNode m_Config.GetNode(cConfigProcessesHeading)
End Sub
Public Property Get IsDirty() As Boolean
    IsDirty = bDirty
End Property

Public Sub Add(newItem As cRoute, Optional Key As Variant)
Attribute Add.VB_Description = "Adds a member to a Collection object"
'''    flagChanges
    m_PrivateCollection.Add newItem, IIf(IsMissing(Key), newItem.RouteID, Key)
'    commitChanges
End Sub

'''Public Sub flagChanges()
'''    'reload as needed to update file changes
''''    m_Config.ConfigurationFile = getConfigurationPath
'''End Sub

Public Sub commitChanges()
Dim oRoute As cRoute, oProcess As cProcess
Dim iNode As MSXML2.IXMLDOMNode, iAttribute As MSXML2.IXMLDOMNode
Dim i As Long
Dim bOKtoAdd As Boolean
    'THIS ROUTINE GOES OVER THE ENTIRE IN-MEMORY REPRESENTATION (HTE_GPS) AND SAVES ALL SETTINGS
    'INTO THE CONFIGURATION FILE! CALL THIS ONLY WHEN YOU ARE ABSOULTELY SURE YOU HAVE EVERYTHING
    'AND YOU NEED TO PERSIST IT!
    'Remove Routes
    Set iNode = m_Config.GetNode(cConfigProcessesHeading)
    If iNode Is Nothing Then Set iNode = m_Config.AddChildNode(m_Config.GetNode("configuration"), cConfigProcessesHeading, vbNullString, NODE_ELEMENT)
    For i = 0 To iNode.childNodes.Length - 1
        iNode.removeChild iNode.childNodes.Item(0)
    Next
    For Each oRoute In Me
        Set iNode = m_Config.GetNode(cConfigProcessesHeading, oRoute.RouteID)
        If iNode Is Nothing Then
            Set iNode = m_Config.GetNode(cConfigProcessesHeading)
            Set iNode = m_Config.AddChildNode(m_Config.GetNode(cConfigProcessesHeading), oRoute.RouteID, vbNullString, NODE_ELEMENT)
        End If
        Set iAttribute = iNode.Attributes.getNamedItem("SERVER")
        bOKtoAdd = iAttribute Is Nothing
        If Not bOKtoAdd Then bOKtoAdd = iAttribute.nodeTypedValue <> CStr(oRoute.ServerProcess)
        If bOKtoAdd Then
            If oRoute.ServerProcess Then
                m_Config.addAttribute iNode, "SERVER", CStr(oRoute.ServerProcess)
            Else
                iNode.Attributes.removeNamedItem ("SERVER")
            End If
        End If
        Set iAttribute = iNode.Attributes.getNamedItem("INBOUNDMSG")
        bOKtoAdd = iAttribute Is Nothing
        If Not bOKtoAdd Then bOKtoAdd = iAttribute.nodeTypedValue <> oRoute.InboundType
        If bOKtoAdd Then m_Config.addAttribute iNode, "INBOUNDMSG", oRoute.InboundType
        'Remove all processes listed
        If iNode.hasChildNodes Then
            For i = 0 To iNode.childNodes.Length - 1
                iNode.removeChild iNode.childNodes.Item(0)
            Next
        End If
        For Each oProcess In oRoute
            Dim iChild As MSXML2.IXMLDOMNode
            Set iChild = m_Config.AddChildNode(iNode, oProcess.progID, vbNullString, NODE_ELEMENT, "INSTANCE", oProcess.InstanceID)
            If Not iChild Is Nothing Then
                m_Config.addAttribute iChild, "DESC", oProcess.friendlyName 'adding attributes don't auto-save
            End If
        Next
    Next
    m_Config.save
    bDirty = False
End Sub

Public Property Get AllowExit() As Boolean
    AllowExit = m_AllowExit
End Property

Public Property Let AllowExit(ByVal vData As Boolean)
Dim iNode As MSXML2.IXMLDOMNode
'''    flagChanges
    Set iNode = m_Config.GetNode("GENERAL", "ALLOWEXIT")
    If vData Then
        If iNode Is Nothing Then m_Config.LetNode iNode, "GENERAL", "ALLOWEXIT"
        If Not iNode Is Nothing Then
            iNode.nodeTypedValue = CStr(CBool(vData))
            m_AllowExit = vData 'should always be true here
        End If
    Else
        If Not iNode Is Nothing Then iNode.parentNode.removeChild iNode
        m_AllowExit = False
    End If
    bDirty = True
    commitChanges
End Property

Public Property Get ProcessSettings(ByVal sProgID As String, ByVal sInst As String) As String
Dim iNode As MSXML2.IXMLDOMNode
    If Not m_Config Is Nothing Then
        Set iNode = m_Config.GetNode(sProgID, sInst)
        If Not iNode Is Nothing Then
           ProcessSettings = addTypesConfiguration(iNode).xml
        Else
            ProcessSettings = addTypesConfiguration(getXML(sInst)).xml
        End If
    End If
End Property

Public Function LetProcessSettings(ByVal sProgID As String, ByVal sInst As String, ByVal XMLDOMNode As String) As Boolean
Dim iNode As MSXML2.IXMLDOMNode
Dim oDOM As MSXML2.DOMDocument30
    Set oDOM = New DOMDocument30
    oDOM.async = False
    If oDOM.loadXML(XMLDOMNode) Then
        'If bDirty Then commitChanges
        sProgID = Replace(sProgID, ".PROCESS", vbNullString, , , vbTextCompare)
        Set oDOM.documentElement = removeTypesConfiguration(oDOM.documentElement)
        Set iNode = m_Config.GetNode(m_Config.rootName, sProgID)
        If iNode Is Nothing Then
            Set iNode = m_Config.GetNode(m_Config.rootName)
            m_Config.LetNode iNode, sProgID
            If Not iNode Is Nothing Then
                iNode.appendChild oDOM.documentElement
                LetProcessSettings = True
                m_Config.save
            End If
        Else
            If Not iNode.selectSingleNode(sInst) Is Nothing Then
                If Not iNode.replaceChild(oDOM.documentElement, iNode.selectSingleNode(sInst)) Is Nothing Then
                    LetProcessSettings = True
                    m_Config.save
                End If
            Else
                iNode.appendChild oDOM.documentElement
                LetProcessSettings = True
                m_Config.save
            End If
        End If
        If bDirty Then commitChanges
        
    End If
End Function

Public Function DeleteProcessSettings(ByVal sProgID As String, ByVal sInst As String) As Boolean
Dim iNode As MSXML2.IXMLDOMNode
Dim iInstance As MSXML2.IXMLDOMNode
'''    flagChanges
    If Not m_Config Is Nothing Then
        Set iNode = m_Config.GetNode(m_Config.rootName, Replace(sProgID, ".PROCESS", vbNullString, , , vbTextCompare))
        If Not iNode Is Nothing Then
            Set iInstance = iNode.selectSingleNode(sInst)
            If Not iInstance Is Nothing Then
                iNode.removeChild iInstance
            End If
            If iNode.childNodes.Length = 0 Then
                iNode.parentNode.removeChild iNode
            End If
        End If
    End If
    commitChanges
End Function

Public Property Get ShowInTray() As Boolean
    ShowInTray = m_ShowInTray
End Property

Public Property Let ShowInTray(ByVal vData As Boolean)
Dim iNode As MSXML2.IXMLDOMNode
    Set iNode = m_Config.GetNode("GENERAL", "SYSTRAY")
    If vData Then
        If iNode Is Nothing Then m_Config.LetNode iNode, "GENERAL", "SYSTRAY"
        If Not iNode Is Nothing Then
            iNode.nodeTypedValue = CStr(CBool(vData))
            m_ShowInTray = vData
        End If
    Else
        If Not iNode Is Nothing Then iNode.parentNode.removeChild iNode
        m_ShowInTray = False
    End If
    bDirty = True
End Property

Private Sub initAppSettings()
Dim iNode As MSXML2.IXMLDOMNode
    Set iNode = m_Config.GetNode("GENERAL", "SYSTRAY")
    If Not iNode Is Nothing Then
        If iNode.nodeTypedValue <> vbNullString Then
            m_ShowInTray = CBool(iNode.nodeTypedValue)
        End If
        If Not m_ShowInTray Then iNode.parentNode.removeChild iNode
    End If
    Set iNode = m_Config.GetNode("GENERAL", "ALLOWEXIT")
    If Not iNode Is Nothing Then
        m_AllowExit = CBool(iNode.nodeTypedValue)
        If Not m_AllowExit Then iNode.parentNode.removeChild iNode
    End If
End Sub

Private Function AddFromNode(ByVal iNode As MSXML2.IXMLDOMNode) As Boolean
Dim i As Long
Dim oRoute As cRoute
    If Not iNode Is Nothing Then
        If iNode.hasChildNodes Then
            For i = 0 To iNode.childNodes.Length - 1
                Set oRoute = New cRoute
                If oRoute.AddFromNode(iNode.childNodes(i)) Then Add oRoute
            Next
        End If
    Else
        m_Config.AddChildNode m_Config.GetNode(m_Config.rootName), cConfigProcessesHeading, vbNullString, NODE_ELEMENT
    End If
End Function

Public Property Get typeDescription(ByVal msgType As HTE_GPS.GPS_MESSAGING_TYPES) As String
Dim i As Long
    If Not m_Config Is Nothing Then
        If Not iTypesNode Is Nothing Then
            If iTypesNode.hasChildNodes Then
                For i = 0 To (iTypesNode.childNodes.Length - 1)
                    If iTypesNode.childNodes(i).nodeTypedValue = msgType Then
                        typeDescription = iTypesNode.childNodes(i).nodeName
                        Exit For
                    End If
                Next
            End If
        End If
    End If
End Property

Public Property Get typesNode() As MSXML2.IXMLDOMNode
Dim iChild As MSXML2.IXMLDOMNode
    If Not m_Config Is Nothing Then
        With m_Config
            Set iTypesNode = .GetNode("GENERAL", "TYPES")
            If Not iTypesNode Is Nothing Then
                Set typesNode = iTypesNode
            Else
                If .LetNode(iTypesNode, "GENERAL", "TYPES") Then
                    Set iChild = .AddChildNode(iTypesNode, "TAIP", 0, NODE_ELEMENT, "SOM", "3E")
                    .addAttribute iChild, "EOM", "3C"
                    Set iChild = .AddChildNode(iTypesNode, "NMEA", 1, NODE_ELEMENT, "SOM", "24")
                    .addAttribute iChild, "EOM", "0D0A"
                    Set iChild = .AddChildNode(iTypesNode, "MAPLITE", 2, NODE_ELEMENT, "SOM", "02")
                    .addAttribute iChild, "EOM", "03"
                    .save
                End If
            End If
        End With
    Else
        'default
    End If
End Property

' Remove an item from the collection
Public Sub Remove(index As Variant)
Attribute Remove.VB_Description = "Removes a member from a Collection object"
'''    flagChanges
    m_PrivateCollection.Remove index
    commitChanges
End Sub

' Return a cRoute item from the collection
Function Item(index As Variant) As cRoute
Attribute Item.VB_Description = "Returns a specific member of a Collection "
Attribute Item.VB_UserMemId = 0
    Set Item = m_PrivateCollection.Item(index)
End Function

' Return the number of items in the collection
Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of members in a collection"
    Count = m_PrivateCollection.Count
End Property

' Remove all items from the collection
Public Sub Clear()
Attribute Clear.VB_Description = "Removes all members from a Collection object"
'''    flagChanges
    Set m_PrivateCollection = New Collection
'    commitChanges
End Sub

' Implement support for enumeration (For Each)
Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    ' delegate to the private collection
    Set NewEnum = m_PrivateCollection.[_NewEnum]
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

Private Function getXML(Optional ByVal rootName As String = "root", _
            Optional ByVal sAttribute As String = vbNullString) As MSXML2.DOMDocument30
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
        If sAttribute <> vbNullString Then
            Set newAtt = oXML.createAttribute("source")
            newAtt.nodeTypedValue = sAttribute
            Set namedNodeMap = rootElement.Attributes
            namedNodeMap.setNamedItem newAtt
        End If
        Set getXML = oXML
        Set objPI = Nothing
        Set rootElement = Nothing
        Set oXML = Nothing
        
End Function

