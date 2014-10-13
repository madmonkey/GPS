VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Port As Long
Private m_Settings As String
Private m_RThreshold As Integer
Private m_InputLen As Integer
Private m_InputMode As MSCommLib.InputModeConstants
Private m_DTREnable As Boolean
Private m_EOFEnable As Boolean
Private m_Handshaking As MSCommLib.HandshakeConstants
Private m_InBufferSize As Integer
Private m_NullDiscard As Boolean
Private m_RTSEnable As Boolean
Private m_ValidPorts As String
Private m_InitializationString As String
Private m_MaxBytesCached As Long 'added to take out the "garbage"

'if you repeatedly open/close a com port the control will eventually lose it's initial settings
'created a class to maintain the properties so we can reassign when doing so
Public Property Get Port() As Long
    Port = m_Port
End Property
Public Property Let Port(ByVal vData As Long)
    m_Port = vData
End Property
Public Property Get Settings() As String
    Settings = m_Settings
End Property
Public Property Let Settings(ByVal vData As String)
    m_Settings = vData
End Property
Public Property Get RThreshold() As Integer
    RThreshold = m_RThreshold
End Property
Public Property Let RThreshold(ByVal vData As Integer)
    m_RThreshold = vData
End Property
Public Property Get InputLen() As Integer
    InputLen = m_InputLen
End Property
Public Property Let InputLen(ByVal vData As Integer)
    m_InputLen = vData
End Property
Public Property Get InputMode() As MSCommLib.InputModeConstants
    InputMode = m_InputMode
End Property
Public Property Let InputMode(ByVal vData As MSCommLib.InputModeConstants)
    m_InputMode = vData
End Property
Public Property Get DTREnable() As Boolean
    DTREnable = m_DTREnable
End Property
Public Property Let DTREnable(ByVal vData As Boolean)
    m_DTREnable = vData
End Property
Public Property Get EOFEnable() As Boolean
    EOFEnable = m_EOFEnable
End Property
Public Property Let EOFEnable(ByVal vData As Boolean)
    m_EOFEnable = vData
End Property
Public Property Get Handshaking() As MSCommLib.HandshakeConstants
    Handshaking = m_Handshaking
End Property
Public Property Let Handshaking(ByVal vData As MSCommLib.HandshakeConstants)
    m_Handshaking = vData
End Property
Public Property Get InBufferSize() As Integer
    InBufferSize = m_InBufferSize
End Property
Public Property Let InBufferSize(ByVal vData As Integer)
    m_InBufferSize = vData
End Property
Public Property Get NullDiscard() As Boolean
    NullDiscard = m_NullDiscard
End Property
Public Property Let NullDiscard(ByVal vData As Boolean)
    m_NullDiscard = vData
End Property
Public Property Get RTSEnable() As Boolean
    RTSEnable = m_RTSEnable
End Property
Public Property Let RTSEnable(ByVal vData As Boolean)
    m_RTSEnable = vData
End Property
Private Sub Class_Initialize()
    Port = cCommValue 'default
    GetAvailablePorts
End Sub
Public Property Get AvailablePorts() As String
    AvailablePorts = m_ValidPorts
End Property
Public Sub GetAvailablePorts()
    m_ValidPorts = ListAvailablePorts
End Sub
Public Property Get IsAssignable() As Boolean
    'verifies whether or not port can be assigned to correctly
    GetAvailablePorts 'in case it changed
    IsAssignable = InStr(1, m_ValidPorts, Port) > 0
End Property
Public Property Get ToString()
    ToString = "CommPort [" & Port & "]; Settings [" & Settings & "]; RThreshold [" & _
        RThreshold & "]; InputLen [" & InputLen & "]; InputMode [" & Choose(Val(InputMode) + 1, "Text", "Binary") & "]; DTREnable [" & _
        CStr(DTREnable) & "]; EOFEnable [" & CStr(EOFEnable) & "]; Handshaking [" & Choose(Val(Handshaking) + 1, "None", "XOnXoff", "RTS", "RTSXOnXoff") & _
        "]; InBufferSize [" & InBufferSize & "]; NullDiscard [" & CStr(NullDiscard) & "]; RTSEnable [" & CStr(RTSEnable) & "] MaxCacheBytes [" & CStr(MaxCachedBytes) & "]"
End Property
Public Property Get InitializationValues() As Variant
Dim vVar As Variant
    InitializationValues = Split(m_InitializationString, cSep)
End Property
Public Property Let InitializationString(ByVal vData As String)
    m_InitializationString = vData
End Property
Public Property Let MaxCachedBytes(ByVal vData As Long)
    m_MaxBytesCached = vData
End Property
Public Property Get MaxCachedBytes() As Long
    MaxCachedBytes = m_MaxBytesCached
End Property
