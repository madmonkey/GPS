VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_File As String
Dim m_FriendlyName As String
Dim m_ProgID As String
Dim m_Type As String
Dim m_Status As HTE_GPS.GPS_PROCESSOR_STATUS
Dim m_InstanceID As String

Public Property Get fileName() As String
    fileName = m_File
End Property

Public Property Let fileName(ByVal vData As String)
    m_File = vData
End Property

Public Property Get friendlyName() As String
    friendlyName = m_FriendlyName
End Property

Public Property Let friendlyName(ByVal vData As String)
    m_FriendlyName = vData
End Property

Public Property Get progID() As String
    progID = m_ProgID
End Property

Public Property Let progID(ByVal vData As String)
    m_ProgID = vData
End Property

Public Property Get compiledType() As String
    compiledType = m_Type
End Property

Public Property Let compiledType(ByVal vData As String)
    m_Type = vData
End Property

Public Property Get Status() As HTE_GPS.GPS_PROCESSOR_STATUS
    Status = m_Status
End Property

Public Property Let Status(ByVal vData As HTE_GPS.GPS_PROCESSOR_STATUS)
    m_Status = vData
End Property

Public Property Get InstanceID() As String
    InstanceID = m_InstanceID
End Property

Public Property Let InstanceID(ByVal vData As String)
    m_InstanceID = vData
End Property

Private Sub Class_Initialize()
    Status = GPS_STAT_UNINITIALIZED
End Sub

Public Function AddFromNode(ByVal iNode As MSXML2.IXMLDOMNode) As Boolean
Dim iAttribute As MSXML2.IXMLDOMNode
    progID = iNode.nodeName
    Set iAttribute = iNode.Attributes.getNamedItem("INSTANCE")
    If Not iAttribute Is Nothing Then InstanceID = iAttribute.nodeValue
    Set iAttribute = iNode.Attributes.getNamedItem("DESC")
    If Not iAttribute Is Nothing Then friendlyName = iAttribute.nodeValue
    AddFromNode = (progID <> vbNullString) And (InstanceID <> vbNullString)
End Function

