VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements HTE_GPS.GPSMessage
Dim msgStatus As HTE_GPS.GPS_MESSAGE_STATUS
Dim msgType As HTE_GPS.GPS_MESSAGING_TYPES
Dim msg As String
Dim ent As String 'version 1.1 implementation

Public Property Let MessageStatus(RHS As HTE_GPS.GPS_MESSAGE_STATUS)
    msgStatus = RHS
End Property
Private Property Let GPSMessage_Entity(RHS As String)
    ent = RHS
End Property
Private Property Get GPSMessage_Entity() As String
    GPSMessage_Entity = ent
End Property
Public Property Let Entity(RHS As String)
    GPSMessage_Entity = RHS
End Property
Public Property Get Entity() As String
    Entity = GPSMessage_Entity
End Property
Private Property Let GPSMessage_MessageStatus(RHS As HTE_GPS.GPS_MESSAGE_STATUS)
    MessageStatus = RHS
End Property
Public Property Get MessageStatus() As HTE_GPS.GPS_MESSAGE_STATUS
    MessageStatus = msgStatus
End Property
Private Property Get GPSMessage_MessageStatus() As HTE_GPS.GPS_MESSAGE_STATUS
    GPSMessage_MessageStatus = MessageStatus
End Property
Public Property Let rawMessage(RHS As String)
    msg = RHS
End Property
Private Property Let GPSMessage_rawMessage(RHS As String)
    rawMessage = RHS
End Property
Public Property Get rawMessage() As String
    rawMessage = msg
End Property
Private Property Get GPSMessage_rawMessage() As String
    GPSMessage_rawMessage = rawMessage
End Property
Public Property Let MessageType(RHS As HTE_GPS.GPS_MESSAGING_TYPES)
    msgType = RHS
End Property
Private Property Let GPSMessage_Type(RHS As HTE_GPS.GPS_MESSAGING_TYPES)
    MessageType = RHS
End Property
Public Property Get MessageType() As HTE_GPS.GPS_MESSAGING_TYPES
    MessageType = msgType
End Property
Private Property Get GPSMessage_Type() As HTE_GPS.GPS_MESSAGING_TYPES
    GPSMessage_Type = MessageType
End Property
