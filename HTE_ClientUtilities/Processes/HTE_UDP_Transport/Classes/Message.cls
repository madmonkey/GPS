VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements HTE_GPS.GPSMessage
Dim msgStatus As HTE_GPS.GPS_MESSAGE_STATUS
Dim msgType As HTE_GPS.GPS_MESSAGING_TYPES
Dim msg As String
Dim ent As String
Dim oXML As MSXML2.DOMDocument30
Public Property Let MessageStatus(RHS As HTE_GPS.GPS_MESSAGE_STATUS)
    msgStatus = RHS
End Property
Private Sub Class_Initialize()
    Set oXML = getXML("MODGPSMSG")
End Sub
Private Property Let GPSMessage_Entity(RHS As String)
    ent = RHS
End Property
Private Property Get GPSMessage_Entity() As String
    GPSMessage_Entity = ent
End Property
Public Property Get Entity() As String
    Entity = GPSMessage_Entity
End Property
Public Property Let Entity(RHS As String)
    GPSMessage_Entity = RHS
End Property
Private Property Let GPSMessage_MessageStatus(RHS As HTE_GPS.GPS_MESSAGE_STATUS)
    MessageStatus = RHS
End Property
Public Property Get MessageStatus() As HTE_GPS.GPS_MESSAGE_STATUS
    MessageStatus = msgStatus
End Property
Private Property Get GPSMessage_MessageStatus() As HTE_GPS.GPS_MESSAGE_STATUS
    GPSMessage_MessageStatus = MessageStatus
End Property
Public Property Let rawMessage(RHS As String)
    msg = RHS
End Property
Private Property Let GPSMessage_rawMessage(RHS As String)
    rawMessage = RHS
End Property
Public Property Get rawMessage() As String
    rawMessage = msg
End Property
Private Property Get GPSMessage_rawMessage() As String
    GPSMessage_rawMessage = rawMessage
End Property
Public Property Let MessageType(RHS As HTE_GPS.GPS_MESSAGING_TYPES)
    msgType = RHS
End Property
Private Property Let GPSMessage_Type(RHS As HTE_GPS.GPS_MESSAGING_TYPES)
    MessageType = RHS
End Property
Public Property Get MessageType() As HTE_GPS.GPS_MESSAGING_TYPES
    MessageType = msgType
End Property
Private Property Get GPSMessage_Type() As HTE_GPS.GPS_MESSAGING_TYPES
    GPSMessage_Type = MessageType
End Property
Public Function Serialize(Optional ByVal PropData As HTE_GPS.GPSMessage) As String

    If Not oXML Is Nothing Then
        If Not IsMissing(PropData) Then
            With PropData
                If Not AddNode(oXML, "TYPE", .Type, NODE_ELEMENT) Is Nothing Then
                    MessageType = .Type
                    If Not AddNode(oXML, "RAWMESSAGE", .rawMessage, NODE_ELEMENT) Is Nothing Then
                        rawMessage = .rawMessage
                        Serialize = oXML.xml
                    End If
                End If
            End With
        Else
            If Not AddNode(oXML, "TYPE", MessageType, NODE_ELEMENT) Is Nothing Then
                If Not AddNode(oXML, "RAWMESSAGE", rawMessage, NODE_ELEMENT) Is Nothing Then
                    Serialize = oXML.xml
                End If
            End If
        End If
    End If
End Function
Public Function Deserialize(ByRef Stream As String) As Boolean
Dim i As Long, iTest As Long
Dim iNode As MSXML2.IXMLDOMElement
Dim bGotType As Boolean, bGotMessage As Boolean
    If Not oXML Is Nothing Then
        If oXML.loadXML(Stream) Then
            'good xml
            If Not oXML.documentElement Is Nothing Then
                If oXML.documentElement.hasChildNodes Then
                    For i = 0 To oXML.documentElement.childNodes.length - 1
                        Set iNode = oXML.documentElement.childNodes.Item(i)
                        Select Case UCase$(iNode.nodeName)
                            Case "TYPE"
                                If IsNumeric(iNode.nodeTypedValue) Then
                                    MessageType = CLng(iNode.nodeTypedValue)
                                    bGotType = True
                                End If
                            Case "RAWMESSAGE"
                                rawMessage = iNode.nodeTypedValue
                                bGotMessage = True
                        End Select
                    Next
                End If
            End If
        End If
    End If
    Deserialize = (bGotType And bGotMessage)
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

Private Function AddNode(ByRef poXML As MSXML2.DOMDocument30, ByRef nodeName As String, _
            ByRef nodeValue As String, ByVal nodeType As MSXML2.DOMNodeType) As MSXML2.IXMLDOMElement
    Dim rootElement As IXMLDOMElement
    Dim aElement As Object

    Set rootElement = poXML.documentElement
    Set aElement = poXML.createNode(nodeType, nodeName, vbNullString)
    aElement.nodeTypedValue = nodeValue
    rootElement.appendChild aElement
    Set AddNode = aElement
End Function
