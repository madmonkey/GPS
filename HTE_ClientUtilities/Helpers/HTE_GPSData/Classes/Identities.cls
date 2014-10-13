VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Identities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_xml As MSXML2.DOMDocument30
Implements IVSFlexDataSource
Public Enum gpsDataFields
    gafIdentityID = 0
    gafPhysicalLookup = 1
    gafAddressLookup = 2
    gafAppLookup = 3
    gafDevice = 4
    gafAlias = 5
    gafComments = 6
End Enum
Private appMutex As Mutex
Private m_DataSource As String
Private Const cModuleName = "Identities"
Private Const cDataSubFolder = "\Data"
Private Const cDataFile = "\identities.xml"
Implements HTE_Properties.IXMLPropertyBag
Private m_keys As Collection

Private Sub Class_Initialize()
    Set m_xml = getXML
    Set appMutex = New Mutex
    appMutex.ConstructMutex "{444EF90A-3133-41c7-8A14-A0DA2BD3837D}"
    m_DataSource = GetParentPath & cDataSubFolder & cDataFile 'set default
    SetDataSource m_DataSource
    Set m_keys = New Collection
End Sub

Private Sub Class_Terminate()
Dim i As Long
    Set m_xml = Nothing
    appMutex.DestroyMutex
    Set appMutex = Nothing
    For i = m_keys.Count To 1 Step -1
        Set m_keys(1) = Nothing
        m_keys.Remove 1
    Next
End Sub

Private Function IVSFlexDataSource_GetData(ByVal Field As Long, ByVal Record As Long) As String

End Function

Private Function IVSFlexDataSource_GetFieldCount() As Long
    IVSFlexDataSource_GetFieldCount = 7
End Function

Private Function IVSFlexDataSource_GetFieldName(ByVal Field As Long) As String
Select Case Field
        Case gafIdentityID: IVSFlexDataSource_GetFieldName = "AliasID"
        Case gafPhysicalLookup: IVSFlexDataSource_GetFieldName = "Mac Address"
        Case gafAddressLookup: IVSFlexDataSource_GetFieldName = "IP Address"
        Case gafAppLookup: IVSFlexDataSource_GetFieldName = "Application ID"
        Case gafDevice: IVSFlexDataSource_GetFieldName = "Device ID"
        Case gafAlias: IVSFlexDataSource_GetFieldName = "Unit ID/Alias"
        Case gafComments: IVSFlexDataSource_GetFieldName = "Comments"
    End Select
End Function

Private Function IVSFlexDataSource_GetRecordCount() As Long
    If Not m_xml Is Nothing Then
        If m_xml.documentElement.hasChildNodes Then
            IVSFlexDataSource_GetRecordCount = m_xml.documentElement.childNodes.length
        End If
    End If
End Function

Private Sub IVSFlexDataSource_SetData(ByVal Field As Long, ByVal Record As Long, ByVal newData As String)

End Sub

Public Function SetDataSource(ByVal xmlFile As String) As Boolean
    If m_xml Is Nothing Then Set m_xml = getXML
    m_DataSource = xmlFile
    m_xml.async = False
    m_xml.setProperty "SelectionLanguage", "XPath"
    If Dir(xmlFile) = vbNullString Then
        Save
    Else
        If Not m_xml.Load(xmlFile) Then
            UEH_Log cModuleName, "GetConfigurationFile", "Configuration file not valid xml file, creating from scratch.", logWarning, , "Path = " & xmlFile, logString
            If Dir(Replace(xmlFile, ".xml", ".bak")) <> vbNullString Then Kill Replace(xmlFile, ".xml", ".bak")
            FileCopy xmlFile, Replace(xmlFile, ".xml", ".bak")
            Kill xmlFile
            Save
        Else
            SetDataSource = True
        End If
    End If
    Exit Function

err_GetConfigurationFile:
    UEH_LogError cModuleName, "GetConfigurationFile", Err
End Function

Public Property Get DataSource() As String
    DataSource = m_DataSource
End Property

Public Sub Save()
On Error GoTo err_Save
    If Not m_xml Is Nothing Then
        Do While Not appMutex.RequestMutex: Loop
        If m_xml.url <> vbNullString Then
            m_xml.Save m_xml.url
        ElseIf DataSource <> vbNullString Then
            m_xml.Save DataSource
        End If
        appMutex.DiscardMutex
    End If
    Exit Sub
err_Save:
    UEH_LogError cModuleName, "Save", Err
    appMutex.DiscardMutex
End Sub

Private Function getXML(Optional ByVal rootName As String = "identities") As MSXML2.DOMDocument30
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

'''Public Function AddNew() As Identity
'''Dim childNode As MSXML2.IXMLDOMNode
'''Dim oID As Identity, sID As String
'''    If Not m_xml Is Nothing Then
'''        sID = "ID" & appMutex.CreateGUID()
'''        Set childNode = m_xml.createNode(NODE_ELEMENT, sID, vbNullString)
'''        Set oID = New Identity
'''        oID.AliasID = sID
'''        Set AddNew = oID
'''    End If
'''End Function

Private Sub IXMLPropertyBag_ReadProperties(ByVal PropertyBag As HTE_Properties.XMLPropertyBag)
Dim lIndex As Long, lCount As Long
Dim ent As HTE_Entity.Entity
    With PropertyBag
        lCount = .ReadProperty("Total", 0)
        For lIndex = 1 To lCount
            Set ent = GetNewDestination
             .ReadProperty "DEST" & lIndex, dest 'need unique name to retrieve by/reference later
            Add dest
        Next
    End With
End Sub

Private Sub IXMLPropertyBag_WriteProperties(ByVal PropertyBag As HTE_Properties.XMLPropertyBag)
Dim lIndex As Long
    With PropertyBag
        .WriteProperty "Total", Count
        .WriteProperty "MessageGuid", MessageGuid
        For lIndex = 1 To Count
             .WriteProperty "DEST" & lIndex, m_colDestinations.Item(lIndex)
        Next
    End With
End Sub
