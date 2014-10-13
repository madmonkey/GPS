Attribute VB_Name = "modXML"
Option Explicit

Private Const cModuleName = "modXML"
Public Function getXML(Optional ByVal rootName As String = "configuration") As MSXML2.DOMDocument30
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
Public Function AddNode(ByRef poXML As MSXML2.DOMDocument30, ByRef nodeName As String, _
            ByRef nodeValue As String, ByVal nodeType As MSXML2.DOMNodeType, Optional nodeAttributeName As String = "default", Optional nodeAttribute As String = vbNullString, _
            Optional ByVal MIME As Boolean = False) As MSXML2.IXMLDOMElement
    Dim rootElement As IXMLDOMElement
    Dim aElement As Object
    Dim newAtt As IXMLDOMAttribute
    Dim namedNodeMap As IXMLDOMNamedNodeMap
    
    Set rootElement = poXML.documentElement
    Set aElement = poXML.createNode(nodeType, nodeName, vbNullString)
    If nodeAttribute <> vbNullString Then
        Set newAtt = poXML.createAttribute(nodeAttributeName)
        newAtt.nodeTypedValue = nodeAttribute
        Set namedNodeMap = aElement.Attributes
        namedNodeMap.setNamedItem newAtt
    End If
    If MIME Then
        aElement.dataType = "bin.base64"
        aElement.nodeTypedValue = ByteArrayFromString(nodeValue)
    Else
        aElement.nodeTypedValue = nodeValue
    End If
    rootElement.appendChild aElement
    Set AddNode = aElement
End Function

Public Function AddChildNode(ByRef poXML As MSXML2.DOMDocument30, ByRef childNode As MSXML2.IXMLDOMElement, ByRef nodeName As String, _
            ByRef nodeValue As String, ByVal nodeType As MSXML2.DOMNodeType, Optional nodeAttributeName As String = "default", Optional nodeAttribute As String = vbNullString, Optional ByVal MIME As Boolean = False) As MSXML2.IXMLDOMElement
    
    Dim aElement As Object
    Dim newAtt As IXMLDOMAttribute
    Dim namedNodeMap As IXMLDOMNamedNodeMap

    Set aElement = poXML.createNode(nodeType, nodeName, vbNullString)
    If nodeAttribute <> vbNullString Then
        Set newAtt = poXML.createAttribute(nodeAttributeName)
        newAtt.nodeTypedValue = nodeAttribute
        Set namedNodeMap = aElement.Attributes
        namedNodeMap.setNamedItem newAtt
    End If
    If MIME Then
        aElement.dataType = "bin.base64"
        aElement.nodeTypedValue = ByteArrayFromString(nodeValue)
    Else
        aElement.nodeTypedValue = nodeValue
    End If
    childNode.appendChild aElement
    Set AddChildNode = aElement
End Function
Public Function AddNodeAttribute(ByRef poXML As MSXML2.DOMDocument30, ByRef nodeName As String, attributeName As String, _
                attributeValue As String)
    
    Dim aElement As Object
    Dim newAtt As IXMLDOMAttribute
    Dim namedNodeMap As IXMLDOMNamedNodeMap
    
    Set aElement = poXML.nodeFromID(nodeName) 'createNode(nodeType, nodeName, vbNullString)
    If aElement Is Nothing Then Set aElement = AddNode(poXML, nodeName, vbNullString, NODE_ELEMENT)
    Set newAtt = aElement.Attributes.getNamedItem(attributeName)
    If newAtt Is Nothing Then Set newAtt = poXML.createAttribute(attributeName)
    newAtt.nodeTypedValue = attributeValue
    Set namedNodeMap = aElement.Attributes
    namedNodeMap.setNamedItem newAtt
    
End Function
Public Function AddChildNodeAttribute(ByRef poXML As MSXML2.DOMDocument30, ByRef childNode As MSXML2.IXMLDOMElement, ByRef nodeName As String, attributeName As String, _
                attributeValue As String)
    
    Dim aElement As Object
    Dim newAtt As IXMLDOMAttribute
    Dim namedNodeMap As IXMLDOMNamedNodeMap

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
