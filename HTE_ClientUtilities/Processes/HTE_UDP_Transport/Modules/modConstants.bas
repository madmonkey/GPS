Attribute VB_Name = "modConstants"
Option Explicit

Public Const cHelperPage = "HTE_UDP_Config.CustomPropertyPage"
Public Const cSep = "|*|"


Public Const cRemoteAddr = "REMOTEHOST"
Public Const cRemoteAddrValue = "0.0.0.0"
Public Const cRemotePort = "REMOTEPORT"
Public Const cRemotePortValue = 4104
Public Const cLocalAddr = "LOCALADDRESS"
Public Const cLocalAddrValue = "0.0.0.0"
Public Const cLocalPort = "LOCALPORT"
Public Const cLocalPortValue = 4105
Public Const cSegmentSize = "SEGMENTSIZE"
Public Const cSegmentSizeValue = 4096&
Public Const cUseEncryption = "ENCRYPTION"
Public Const cUseEncryptionValue = True
Public Const cUseCompression = "COMPRESSION"
Public Const cUseCompressionValue = True
Public Const cReceiveBufferSize = "RECEIVEBUFFER"
Public Const cReceiveBufferSizeValue = 4096&
Public Const cKeepAliveInterval = "KEEPALIVEINTERVAL" 'milliseconds
Public Const cKeepAliveIntervalValue = 30000&
Public Const cUseKeepAlive = "USEKEEPALIVE"
Public Const cUseKeepAliveValue = False
Public Const cRetryInterval = "RETRYINTERVAL" 'milliseconds
Public Const cRetryIntervalValue = 5000&
Public Const cRetryRandomInterval = "RETRYRANDOMINTERVAL" 'milliseconds
Public Const cRetryRandomIntervalValue = 5000&
Public Const cKeepAliveFailureInterval = "KEEPALIVEFAILUREINT" 'milliseconds
Public Const cKeepAliveFailureIntervalValue = 600000
Public Const cMaxKeepAliveFailures = "MAXKEEPALIVEFAILURES"
Public Const cMaxKeepAliveFailuresValue = 5
Public Const cTimeToLive = "TIMETOLIVE" 'seconds
Public Const cTimeToLiveValue = 0&
Public Const cProcessOnSend = "PROCESSMSGONSEND"
Public Const cProcessOnSendValue = False
Public Const cSerializeMsg = "SERIALIZEMSGSTRUCT"
Public Const cSerializeMsgValue = True
