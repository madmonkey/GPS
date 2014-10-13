VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRoute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_PrivateCollection As Collection
Dim m_ServerProcess As Boolean
Dim m_InboundType As Long
Dim m_RouteID As String

Private Sub Class_Initialize()
    Set m_PrivateCollection = New Collection
    m_RouteID = vbNullString
    m_InboundType = 0
    m_ServerProcess = False
End Sub

Public Sub Add(newItem As cProcess, Optional Key As Variant)
Attribute Add.VB_Description = "Adds a member to a Collection object"
    m_PrivateCollection.Add newItem, IIf(IsMissing(Key), newItem.InstanceID, Key)
End Sub
Public Function AddFromNode(ByVal iNode As MSXML2.IXMLDOMNode) As Boolean
Dim oProcess As cProcess
Dim iAttribute As MSXML2.IXMLDOMNode
Dim i As Long
    If Not iNode Is Nothing Then
        If Not iNode.nodeName = "#text" Then
            m_RouteID = iNode.nodeName
            Set iAttribute = iNode.Attributes.getNamedItem("SERVER")
            If Not iAttribute Is Nothing Then ServerProcess = CBool(iAttribute.nodeValue)
            Set iAttribute = iNode.Attributes.getNamedItem("INBOUNDMSG")
            If Not iAttribute Is Nothing Then InboundType = CLng(iAttribute.nodeValue)
            If iNode.hasChildNodes Then
                For i = 0 To iNode.childNodes.Length - 1
                    Set oProcess = New cProcess
                    If oProcess.AddFromNode(iNode.childNodes(i)) Then Add oProcess
                Next
            End If
        End If
    End If
    AddFromNode = m_RouteID <> vbNullString
End Function

Public Sub Remove(index As Variant)
Attribute Remove.VB_Description = "Removes a member from a Collection object"
    m_PrivateCollection.Remove index
End Sub

Function Item(index As Variant) As cProcess
Attribute Item.VB_Description = "Returns a specific member of a Collection "
Attribute Item.VB_UserMemId = 0
    Set Item = m_PrivateCollection.Item(index)
End Function

Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of members in a collection"
    Count = m_PrivateCollection.Count
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all members from a Collection object"
    Set m_PrivateCollection = New Collection
End Sub

Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_PrivateCollection.[_NewEnum]
End Function

Public Property Get ServerProcess() As Boolean
    ServerProcess = m_ServerProcess
End Property

Public Property Get RouteID() As String
    RouteID = m_RouteID
End Property

Public Property Let RouteID(ByVal vData As String)
    m_RouteID = vData
End Property

Public Property Let ServerProcess(ByVal vData As Boolean)
    m_ServerProcess = vData
End Property

Public Property Get InboundType() As Long
    InboundType = m_InboundType
End Property

Public Property Let InboundType(ByVal vData As Long)
    m_InboundType = vData
End Property
