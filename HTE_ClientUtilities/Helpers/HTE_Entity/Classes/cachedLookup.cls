VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cachedLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Time As Long
Private m_Alias As String
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Sub Class_Initialize()
    m_Time = timeGetTime
End Sub

Public Property Let Alias(ByRef vData As String)
    m_Alias = vData
End Property

Public Property Get Alias() As String
    Alias = m_Alias
End Property

Public Property Get Created() As Long
    Created = m_Time
End Property

