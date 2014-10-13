Attribute VB_Name = "modProperties"
Option Explicit
Option Compare Binary

Public Enum getPropertyReturnCode
    gprcFound = 0
    gprcDefault = 1
    gprcBadNode = 2
End Enum

Public Function letProperty(ByRef localSettings As MSXML2.DOMDocument30, ByVal PropertyName As String, ByVal propertyValue As String) As Boolean
Dim aElement As MSXML2.IXMLDOMNode
Dim rootElement As IXMLDOMElement

    If Not localSettings Is Nothing Then
        Set aElement = localSettings.getElementsByTagName(PropertyName).Item(0)
        If Not aElement Is Nothing Then
            letProperty = (aElement.nodeTypedValue <> propertyValue) 'same value
            aElement.nodeTypedValue = propertyValue
        Else
            Set rootElement = localSettings.documentElement
            Set aElement = localSettings.createNode(NODE_ELEMENT, PropertyName, vbNullString)
            aElement.nodeTypedValue = propertyValue
            rootElement.appendChild aElement
            letProperty = True
        End If
    End If
    
End Function

Public Function getProperty(ByRef localSettings As MSXML2.DOMDocument30, ByVal PropertyName As String, _
            ByVal defaultValue As String, Optional ByVal bAsXML As Boolean = False, _
            Optional ByRef ReturnCode As getPropertyReturnCode) As String
    
    If Not localSettings Is Nothing Then
        If localSettings.documentElement.hasChildNodes Then
            If localSettings.getElementsByTagName(PropertyName).length > 0 Then
                If Not bAsXML Then
                    getProperty = localSettings.getElementsByTagName(PropertyName).Item(0).nodeTypedValue
                Else
                    getProperty = localSettings.getElementsByTagName(PropertyName).Item(0).xml
                End If
                ReturnCode = gprcFound
                Exit Function
            Else
                ReturnCode = gprcDefault
            End If
        Else
            ReturnCode = gprcBadNode
        End If
    End If
    
    getProperty = defaultValue
    
End Function

Public Function PropertiesChanged(ByRef SettingString As String, ByRef localSettings As MSXML2.DOMDocument30, Optional ByRef bBadSettings As Boolean) As Boolean
Dim iRoot As MSXML2.DOMDocument30
    
    If Not localSettings Is Nothing Then
        Set iRoot = New MSXML2.DOMDocument30
        If iRoot.loadXML(SettingString) Then
            PropertiesChanged = (StrComp(CStr(localSettings.xml), iRoot.xml) <> 0)
        Else
            bBadSettings = True
        End If
    End If
    
End Function

Public Function SupportedMessages(ByRef Settings As MSXML2.DOMDocument30) As HTE_GPS.GPSConfiguration()
Dim returnArray() As HTE_GPS.GPSConfiguration
Dim x As Long
    If Settings.xml <> vbNullString Then
        If Not Settings.documentElement.getElementsByTagName("TYPES") Is Nothing Then
            If Not Settings.documentElement.getElementsByTagName("TYPES").Item(0) Is Nothing Then
                If Settings.documentElement.getElementsByTagName("TYPES").Item(0).hasChildNodes Then
                    If Settings.documentElement.getElementsByTagName("TYPES").Item(0).childNodes.length > 0 Then
                        ReDim returnArray(Settings.documentElement.getElementsByTagName("TYPES").Item(0).childNodes.length - 1)
                        For x = 0 To UBound(returnArray)
                            With Settings.documentElement.getElementsByTagName("TYPES").Item(0).childNodes(x)
                                returnArray(x).Desc = .nodeName
                                If Not .Attributes.getNamedItem("EOM") Is Nothing Then returnArray(x).EOM = formatTag(.Attributes.getNamedItem("EOM").nodeTypedValue)
                                If IsNumeric(.nodeTypedValue) Then returnArray(x).GPSType = CLng(.nodeTypedValue)
                                If Not .Attributes.getNamedItem("SOM") Is Nothing Then returnArray(x).SOM = formatTag(.Attributes.getNamedItem("SOM").nodeTypedValue)
                            End With
                        Next
                    SupportedMessages = returnArray
                    End If
                End If
            End If
        End If
    End If

End Function

Private Function formatTag(ByVal HexValue As String) As String
Dim i As Long
Dim sRtn As String
    i = 1
    sRtn = vbNullString
    On Error GoTo exit_formatTag
    If Len(HexValue) Mod 2 = 0 Then
        Do While i < Len(HexValue)
            sRtn = sRtn & Chr$("&H" & Mid$(HexValue, i, 2))
            i = i + 2
        Loop
    End If
exit_formatTag:
    formatTag = sRtn
End Function
