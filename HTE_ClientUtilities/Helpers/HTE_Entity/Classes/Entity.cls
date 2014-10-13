VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Entity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements HTE_Properties.IXMLPropertyBag
Private m_IPAddress As String
Private m_MACAddress As String
Private m_DeviceID As String
Private m_Alias As String
Private m_Comments As String
Private m_AppID As String 'device id transmitted from device
Private m_StatusCode As String
Private m_StatusDesc As String
Private m_IsProxy As Boolean

Public Property Get IPAddress() As String
Attribute IPAddress.VB_Description = "IPAddress of the entity"
    IPAddress = m_IPAddress
End Property

Public Property Let IPAddress(ByVal Data As String)
    m_IPAddress = Data
End Property

Public Property Get ApplicationID() As String
    ApplicationID = m_AppID
End Property

Public Property Let ApplicationID(ByVal Data As String)
    m_AppID = Data
End Property

Public Property Get MACAddress() As String
Attribute MACAddress.VB_Description = "Physical Identifier of the entity"
    MACAddress = m_MACAddress
End Property

Public Property Let MACAddress(ByVal Data As String)
    m_MACAddress = Data
End Property

Public Property Get DeviceID() As String
Attribute DeviceID.VB_Description = "User-defined device identifier"
    DeviceID = m_DeviceID
End Property

Public Property Let DeviceID(ByVal Data As String)
    m_DeviceID = Data
End Property

Public Property Get Alias() As String
Attribute Alias.VB_Description = "User defined identifier relating to how the entity is identified"
    Alias = m_Alias
End Property

Public Property Let Alias(ByVal Data As String)
    m_Alias = Data
End Property

Public Property Get Comments() As String
Attribute Comments.VB_Description = "User-defined comments or notes regarding the entity"
    Comments = m_Comments
End Property

Public Property Let Comments(ByVal Data As String)
    m_Comments = Data
End Property

Public Property Get StatusCode() As String
    StatusCode = m_StatusCode
End Property

Public Property Let StatusCode(ByVal Data As String)
    m_StatusCode = Data
End Property

Public Property Get StatusDescription() As String
    StatusDescription = m_StatusDesc
End Property

Public Property Let StatusDescription(ByVal Data As String)
    m_StatusDesc = Data
End Property

Public Property Get IsProxy() As Boolean
    IsProxy = m_IsProxy
End Property

Public Property Let IsProxy(ByVal Data As Boolean)
    m_IsProxy = Data
End Property

Private Sub IXMLPropertyBag_ReadProperties(ByVal PropertyBag As HTE_Properties.XMLPropertyBag)
    With PropertyBag
        m_IPAddress = .ReadProperty("IPAddress")
        m_MACAddress = .ReadProperty("MACAddress")
        m_DeviceID = .ReadProperty("Device")
        m_Alias = .ReadProperty("Alias")
        m_Comments = .ReadProperty("Comments")
        m_AppID = .ReadProperty("AppId") 'how we expect the device to identify itself, grain of salt
        m_StatusCode = .ReadProperty("Code")
        m_StatusDesc = .ReadProperty("Desc")
        m_IsProxy = .ReadProperty("Proxy")
    End With
End Sub

Private Sub IXMLPropertyBag_WriteProperties(ByVal PropertyBag As HTE_Properties.XMLPropertyBag)
    With PropertyBag
        .WriteProperty "IPAddress", m_IPAddress
        .WriteProperty "MACAddress", m_MACAddress
        .WriteProperty "Device", m_DeviceID
        .WriteProperty "Alias", m_Alias
        .WriteProperty "Comments", m_Comments
        .WriteProperty "AppId", m_AppID
        .WriteProperty "Code", m_StatusCode
        .WriteProperty "Desc", m_StatusDesc
        .WriteProperty "Proxy", m_IsProxy
    End With
End Sub

Public Function Deserialize(ByVal XMLStream As String) As Boolean
Dim xmlProp As HTE_Properties.XMLPropertyBag
On Error GoTo err_Deserialize
    Set xmlProp = New HTE_Properties.XMLPropertyBag
    xmlProp.Contents = XMLStream
    xmlProp.RestoreState Me
    Deserialize = Len(xmlProp.Contents) > 0
    Set xmlProp = Nothing
    Exit Function
err_Deserialize:
    Err.Clear
    Deserialize = False
End Function

Public Function Serialize() As String
Dim xmlProp As HTE_Properties.XMLPropertyBag
    Set xmlProp = New HTE_Properties.XMLPropertyBag
    xmlProp.SaveState Me
    Serialize = xmlProp.Contents
    Set xmlProp = Nothing
End Function

Public Function Clone(Optional ByVal bDeepCopy As Boolean = False) As HTE_Entity.Entity
Dim tmp As HTE_Entity.Entity
    Set tmp = New HTE_Entity.Entity
    With tmp
        .ApplicationID = ApplicationID
        .IPAddress = IPAddress
        .MACAddress = MACAddress
        .DeviceID = DeviceID
        .Alias = Alias
        .StatusDescription = StatusDescription
        .IsProxy = IsProxy
        If bDeepCopy Then
            .Comments = Comments
            .StatusCode = StatusCode
        End If
    End With
    Set Clone = tmp
End Function
