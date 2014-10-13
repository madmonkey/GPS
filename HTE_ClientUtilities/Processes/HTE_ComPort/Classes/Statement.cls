VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Statement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_Data As String
Dim m_Interval As Long
Dim m_NextAlarm As Date

Public Property Let pollingBuffer(ByVal vData As String)
    m_Data = vData
End Property
Public Property Get pollingBuffer() As String
    pollingBuffer = m_Data
End Property
Public Property Let pollingInterval(ByVal vData As Long)
    m_Interval = vData
End Property
Public Property Get pollingInterval() As Long
    pollingInterval = m_Interval
End Property
Public Property Let NextAlarm(ByVal vData As Date)
    m_NextAlarm = vData
End Property
Public Property Get NextAlarm() As Date
    NextAlarm = m_NextAlarm
End Property

