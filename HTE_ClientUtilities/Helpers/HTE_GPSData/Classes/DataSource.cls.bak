VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim oConfig As MSXML2.DOMDocument30
Dim m_Path As String
Dim m_Root As String
Private Const cModuleName = "DataSource"
Private appMutex As Mutex

Public Property Let ConfigurationFile(ByVal FileName As String)
    m_Path = FileName
    GetConfigurationFile
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

On Error GoTo err_Save

    If Not oConfig Is Nothing Then
        Do While Not appMutex.RequestMutex: Loop
        If oConfig.url <> vbNullString Then
            oConfig.Save oConfig.url
        ElseIf ConfigurationFile <> vbNullString Then
            oConfig.Save ConfigurationFile
        End If
        appMutex.DiscardMutex
    End If
    Exit Sub

err_Save:
    UEH_LogError cModuleName, "Save", Err
    appMutex.DiscardMutex
    
End Sub

Private Function GetConfigurationFile() As Boolean

On Error GoTo err_GetConfigurationFile
    
    If oConfig Is Nothing Then Set oConfig = getXML(rootName)
    oConfig.async = False
''''   http://www.15seconds.com/issue/010409.htm
'''    2) MSXML 3.0 Defaults Itself to Support XSL Pattern Matching Instead of XPath
'''    When using the selectNodes method in MSXML 3.0, the object will be defaulted to use
'''    simple XSL pattern matching. As a result, you will need to call the SetProperty method
'''    first to toggle the SelectionLanguage property.
'''    This property defaults to XSLPattern, however, you will need to set it to XPath.
'''    This will enable you to use the XPath functions listed in the documentation
'''    (such as contains(), name(), starts-with() etc).
    oConfig.setProperty "SelectionLanguage", "XPath"
    
    If Dir(m_Path) = vbNullString Then
        UEH_Log cModuleName, "GetConfigurationFile", "Configuration file not found, creating from scratch.", logWarning, , "Path = " & m_Path, logString
        oConfig.Save m_Path
    Else
        If Not oConfig.Load(m_Path) Then
            UEH_Log cModuleName, "GetConfigurationFile", "Configuration file not valid xml file, creating from scratch.", logWarning, , "Path = " & m_Path, logString
            If Dir(Replace(m_Path, ".xml", ".bak")) <> vbNullString Then Kill Replace(m_Path, ".xml", ".bak")
            FileCopy m_Path, Replace(m_Path, ".xml", ".bak")
            Kill m_Path
            oConfig.Save m_Path
        Else
            GetConfigurationFile = True
        End If
    End If
    Exit Function

err_GetConfigurationFile:
    UEH_LogError cModuleName, "GetConfigurationFile", Err
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
    rootName = "Aliases"
    Set appMutex = New Mutex
    appMutex.ConstructMutex "{9DF3034E-BF43-4e69-94A3-D410B890128E}"
End Sub

Private Sub Class_Terminate()
    appMutex.DestroyMutex
    Set appMutex = Nothing
    UEH_EndLogging
End Sub

Public Function CreateEntity(ByRef oEntity As HTE_Entity.Entity) As Entity
    
End Function

Public Function DeleteEntity(ByRef oEntity As HTE_Entity.Entity) As Entity

End Function

Public Function RetrieveEntity(ByVal queryString As String) As Entity

End Function

Public Function UpdateEntity(ByRef oEntity As HTE_Entity.Entity) As Entity

End Function


