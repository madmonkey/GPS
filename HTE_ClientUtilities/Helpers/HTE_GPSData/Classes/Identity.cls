VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Identity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_AliasID As String
Private m_Entity As HTE_Entity.Entity

Public Property Get AliasID() As String
    AliasID = m_AliasID
End Property

Public Property Let AliasID(ByVal Data As String)
    m_AliasID = Data
End Property

Public Property Get InstanceEntity() As HTE_Entity.Entity
    Set InstanceEntity = m_Entity
End Property

Public Property Let InstanceEntity(Data As HTE_Entity.Entity)
    Set m_Entity = Data
End Property

Private Sub Class_Initialize()
    Set m_Entity = New HTE_Entity.Entity
    Debug.Print m_Entity.Serialize
End Sub

Public Property Get Serialize() As String
        
End Property

Public Property Get Deserialize(ByVal Data As String) As Boolean
    
End Property
