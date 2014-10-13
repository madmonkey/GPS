Attribute VB_Name = "modConstants"
Option Explicit

Public Const cHelperPage = "HTE_ComPortConfig.CustomPropertyPage"
Public Const cSep = "|*|"

Public Const cComm = "COMPORT"
Public Const cCommValue = 1
Public Const cSettings = "SETTINGS"
Public Const cSettingsValue = "4800,N,8,1"
Public Const cRThresh = "RTHRESHOLD"
Public Const cRThreshValue = 1
Public Const cInputLen = "INPUTLEN"
Public Const cInputLenValue = 0
Public Const cInputMode = "INPUTMODE"
Public Const cInputModeValue = 0 'MSCommLib.comInputModeText - since library not referenced in Property OCX
Public Const cInitString = "INITSTRING"
Public Const cInitStringValue = "FLN00300000"
Public Const cDTREnable = "DTRENABLE"
Public Const cDTREnableValue = -1
Public Const cEOFEnable = "EOFENABLE"
Public Const cEOFEnableValue = 0
Public Const cHandshaking = "HANDSHAKING"
Public Const cHandshakingValue = 0 'MSCommLib.HandshakeConstants.comNone - since library not referenced in Property OCX
Public Const cInBufferSize = "INBUFFERSIZE"
Public Const cInBufferSizeValue = 1024
Public Const cNullDiscard = "NULLDISCARD"
Public Const cNullDiscardValue = 0
Public Const cRTSEnable = "RTSENABLE"
Public Const cRTSEnableValue = 0
Public Const cManualPoll = "MANUALPOLLING"
Public Const cManualPollValue = "False"
Public Const cManualPollString = "POLLINGSTRING"
Public Const cManualPollStringValue = vbNullString
Public Const cManualPollInterval = "POLLINGINTERVAL"
Public Const cManualPollIntervalValue = vbNullString
Public Const cRelayInterval = "RELAYINTERVAL"
Public Const cRelayIntervalValue = 0
Public Const cDurationInterval = "DURATIONINTERVAL"
Public Const cDurationIntervalValue = 0
Public Const cMaxCacheBufferSize = "MAXCACHEBUFFERBYTES"
Public Const cMaxCacheBufferSizeValue = 4096
Public Const cProcessOnSend = "PROCESSMSGONSEND"
Public Const cProcessOnSendValue = "False"
Public Const cMaxNumberOfPorts = 96

'API Declarations
'Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
'Modified to work under 95, 98 & ME
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'API Structures
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

'API constants
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_NORMAL = &H80

Public Function IsCOMPortAvailable(Port As Long) As Boolean
'Return TRUE if the COM exists, FALSE if the COM does not exist
    Dim hCOM As Long
    Dim ret As Long
    Dim sec As SECURITY_ATTRIBUTES

    'try to open the COM port
    hCOM = CreateFile("\\.\COM" & Port & "", 0&, FILE_SHARE_READ + FILE_SHARE_WRITE, sec, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
    If hCOM = -1 Then
        IsCOMPortAvailable = False
    Else
        IsCOMPortAvailable = True
        'close the COM port
        ret = CloseHandle(hCOM)
    End If
End Function

Public Function ListAvailablePorts() As String
Dim i As Long
Dim sReturn As String
    sReturn = vbNullString
    For i = 1 To 16
        If IsCOMPortAvailable(i) Then
            If sReturn <> vbNullString Then
                sReturn = sReturn & "," & CStr(i)
            Else
                sReturn = CStr(i)
            End If
        End If
    Next
    ListAvailablePorts = sReturn
End Function
