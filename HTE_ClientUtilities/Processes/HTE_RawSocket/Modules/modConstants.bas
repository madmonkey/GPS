Attribute VB_Name = "modConstants"
Option Explicit

Public Const cHelperPage = "HTE_SocketConfig.CustomPropertyPage"
Public Const cSep = "|*|"

Public Const cProtocol = "PROTOCOL"
Public Const cProtocolValue = "UDP"
Public Const cAddress = "IPADDRESS"
Public Const cPort = "PORT"
Public Const cLocalPort = "LOCALPORT"
Public Const cPortValue = 8669
Public Const cLocalPortValue = 21000
Public Const cInitString = "INITSTRING"
Public Const cInitStringValue = vbNullString
Public Const cValidateMessage = "VALIDATE"
Public Const cValidateMessageValue = "False"
Public Const cProcessOnSend = "PROCESSMSGONSEND"
Public Const cProcessOnSendValue = "False"
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
Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Const cEnableMacResolution = "RESOLVEMAC"
Public Const cEnableMacResolutionValue = "False"
Public Const cCacheMacLookupSeconds = "CACHEMACTIME"
Public Const cCacheMacLookupSecondsValue = 3600

Public Function ComputerName() As String
    Dim dwLen As Long
    Dim strString As String
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, Chr(0))
    GetComputerName strString, dwLen
    strString = Left(strString, dwLen)
    ComputerName = strString
End Function

