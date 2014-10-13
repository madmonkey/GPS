VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRoutes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_PrivateCollection As Collection
Dim m_ShowInTray As Boolean
Dim m_AllowExit As Boolean
Dim m_Types As Collection
Private Const cConfigProcessesHeading = "HTE_CLIENTUTILITIES"

Private Sub Class_Initialize()
    Set m_PrivateCollection = New Collection
    Set m_Types = New Collection
End Sub

Private Sub Add(newItem As cRoute, Optional Key As Variant)
Attribute Add.VB_Description = "Adds a member to a Collection object"
    m_PrivateCollection.Add newItem, Key
End Sub

Public Sub AddFromNode(ByVal iNode As MSXML2.IXMLDOMNode)
Dim i As Long
Dim oRoute As cRoute
    If Not iNode Is Nothing Then
        If iNode.hasChildNodes Then
            For i = 0 To iNode.childNodes.Length - 1
                Set oRoute = New cRoute
                If oRoute.AddFromNode(iNode.childNodes(i).nodeName) Then
                    Add oRoute, iNode.childNodes(i).nodeName
                End If
            Next
        End If
    End If
End Sub

' Remove an item from the collection
Public Sub Remove(index As Variant)
Attribute Remove.VB_Description = "Removes a member from a Collection object"
    m_PrivateCollection.Remove index
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
    Set m_PrivateCollection = New Collection
End Sub

' Implement support for enumeration (For Each)
Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    ' delegate to the private collection
    Set NewEnum = m_PrivateCollection.[_NewEnum]
End Function
