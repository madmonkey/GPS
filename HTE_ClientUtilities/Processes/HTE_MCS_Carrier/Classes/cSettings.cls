VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim sINI As String
Dim sSection As String
Public Enum iniKeys
    ikSessionType = 0
    ikSessionINI = 1
    ikSessionINIKey = 2
    ikHTEUserID = 3
    ikConnectName = 4
    ikParagraph = 5
    ikSessions = 6
    ikNewUser = 7 'So we can keep the mess that is the MCS.INI file pretty
End Enum
'Regular ini expressions
Private m_Session As String
Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'All the insanity to monitor changes for CAD400 - terminal/Unit
'NOTE: 04/01/2004 - NOW MONITORS CAD400 & CADV
Public Event UserChanged(ByVal HTEUserID As String)
Private WithEvents monitorDirectory As CTimer
Attribute monitorDirectory.VB_VarHelpID = -1
Dim terminateFlag As Long
Dim currentUser As String
Dim ClientINI As String
Dim bMonitoring As Boolean
Dim defaultUser As String

Private Function ComputerName() As String
Dim dwLen As Long
Dim strString As String
    
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, Chr(0))
    GetComputerName strString, dwLen
    strString = Left(strString, dwLen)
    ComputerName = strString
End Function

Public Property Get INIFile() As String
    INIFile = sINI
End Property

Public Property Let INIFile(ByVal vData As String)
    sINI = vData
End Property

Public Property Let SessionName(ByVal vData As String)
    m_Session = vData
End Property

Public Property Get SessionName() As String
    SessionName = m_Session
End Property
Public Sub Terminate()
    'only flags the monitoring thread for termination..allows to be released from memory
    terminateFlag = True
End Sub
Public Sub MonitorForChanges()
    Set monitorDirectory = New CTimer
    monitorDirectory.Interval = 5000
    terminateFlag = False
End Sub
Public Sub StopMonitorForChanges()
    terminateFlag = True
    Set monitorDirectory = Nothing
End Sub
Private Function VerifyChange() As Boolean
Dim sCompare As String
Dim sTemp As String
    sTemp = currentUser
    sCompare = GetValue(ikNewUser)
    If StrComp(sTemp, sCompare, vbTextCompare) <> 0 Then
        RaiseEvent UserChanged(sCompare)
        VerifyChange = True
    End If
End Function

Private Sub Class_Initialize()
    INIFile = Environ("WINDIR") & "\mcs.ini"
    ClientINI = Environ("WINDIR") & "\hte_common.ini"
End Sub

Public Function GetValue(ByVal key As iniKeys) As Variant
Dim vDefault As Variant, NC As Long
Dim sKey As String, Ret As String, Section As String, OKtoWrite As Boolean
Dim DataType As VBA.VbVarType
Dim myINI As String
Const cMax = 2048
    Select Case key
        Case ikSessionType
            vDefault = "ANY_MSG"
        Case ikSessionINI
            vDefault = "MCS_MDB.INI"
        Case ikSessionINIKey
            vDefault = "????,MDB"
        Case ikHTEUserID
            vDefault = Left$(ComputerName, 4)
        Case ikConnectName
            vDefault = "MDB_CONNECTION"
        Case ikParagraph
            vDefault = "MDB_TCP_CONNECTION"
        Case ikSessions
            GetValue = ReturnSessionNames(GetINIFile(key))
            Exit Function
        Case ikNewUser 'read the default once since it is nested so deep...
            If defaultUser = vbNullString Then
                vDefault = GetValue(ikHTEUserID)
                defaultUser = vDefault
            Else
                vDefault = defaultUser
            End If
    End Select
    
    myINI = GetINIFile(key)
    sKey = GetKeyValue(key, DataType, OKtoWrite)
    Section = GetSection(key)
    Select Case DataType
        Case vbLong, vbBoolean
            GetValue = GetPrivateProfileInt(Section, sKey, CLng(vDefault), myINI)
            If DataType = vbBoolean Then
                GetValue = (GetValue = 1)
                WritePrivateProfileString Section, sKey, CStr(Abs(GetValue)), myINI
            End If
        Case Else
            Ret = String(cMax, 0)
            NC = GetPrivateProfileString(Section, sKey, CStr(vDefault), Ret, cMax, myINI)
            If NC <> 0 Then GetValue = Trim$(Left$(Ret, NC))
    End Select
    If key = ikNewUser Then currentUser = GetValue
    If DataType <> vbBoolean And OKtoWrite Then WritePrivateProfileString Section, sKey, Trim$(CStr(GetValue)), myINI
End Function

Private Function GetINIFile(ByVal key As iniKeys) As String
    
    Select Case key
        Case ikSessionINIKey
            GetINIFile = Environ("WINDIR") & "\" & GetValue(ikSessionINI)
        Case ikNewUser
            GetINIFile = ClientINI
        Case Else
            GetINIFile = Environ("WINDIR") & "\mcs.ini"
    End Select
    INIFile = GetINIFile
    
End Function

Public Function LetValue(ByVal key As iniKeys, ByVal Value As Variant) As Variant
Dim sKey As String, Section As String, myINI As String
Dim DataType As VBA.VbVarType
    
    myINI = GetINIFile(key)
    DataType = vbLong
    sKey = GetKeyValue(key, DataType)
    Section = GetSection(key)
    If DataType = vbBoolean Then Value = CLng(Abs(Abs(Value) = 1))
    WritePrivateProfileString Section, sKey, CStr(Value), myINI

End Function

Private Function GetKeyValue(ByVal key As iniKeys, Optional ByRef iniType As VBA.VbVarType, Optional bOKtoWrite As Boolean) As String

    iniType = vbLong
    Select Case key
        Case ikSessionType
            GetKeyValue = "SesType": iniType = vbString
        Case ikSessionINI
            GetKeyValue = "SesIni": iniType = vbString
        Case ikSessionINIKey
            GetKeyValue = "MessageKeys": iniType = vbString
        Case ikNewUser
            GetKeyValue = "TerminalUnit": iniType = vbString
        Case ikHTEUserID
            GetKeyValue = "HTEUserID": iniType = vbString
        Case ikConnectName
            GetKeyValue = "ConnectName": iniType = vbString
        Case ikParagraph
            GetKeyValue = "Paragraph": iniType = vbString
    End Select
    
End Function

Private Function GetSection(ByVal key As iniKeys) As String
    Select Case key
        Case ikSessionType, ikSessionINI, ikConnectName
            GetSection = SessionName
        Case ikSessionINIKey
            GetSection = "General"
        Case ikNewUser
            GetSection = "Client"
        Case ikParagraph
            GetSection = GetValue(ikConnectName)
        Case ikHTEUserID
            GetSection = GetValue(ikParagraph)
    End Select
End Function

Private Function ReturnSessionNames(ByVal sFileName As String) As Variant
Dim szBuf As String, Length As Integer
Dim SectionArr() As String, i As Integer, sReturn As String
Const cMax = 255
    szBuf = String$(cMax, 0)
    Length = GetPrivateProfileSectionNames(szBuf, cMax, sFileName)
    szBuf = Left$(szBuf, Length)
    SectionArr = Split(szBuf, vbNullChar)
    For i = 0 To UBound(SectionArr)
        szBuf = String$(cMax, 0) 'test for a connection name
        If GetPrivateProfileString(SectionArr(i), "ConnectName", vbNullString, szBuf, cMax, sFileName) > 0 Then
            sReturn = sReturn & SectionArr(i) & "|*|"
        End If
    Next
    ReturnSessionNames = sReturn
End Function

Private Sub Class_Terminate()
    Set monitorDirectory = Nothing
End Sub

Private Sub monitorDirectory_ThatTime()
'Since we are reading from our own file...rather than MCS' nested ini
'this is more expediant and less troubling, processwise to re-check on occasion
'than monitor directory and respond to changes where applicable.

'After restructuring for independant threads <ActiveX exe>, we noticed that MCS
'doesn't seem to unload properly even though we are closing with the same
'routine; setting terminateFlag and destroying timer...<sigh> so we aren't monitoring
'directory for changes, we are rechecking based on interval - to workaround MCS not
'be terminating appropriately - ie never returns from Terminate and stays in taskbar,
'for one reason or another.

'NOTE: This ONLY gets enabled for CAD400 format
'04/01/2004 - NOW MONITORS BOTH CAD400 AND CADV!
    If Not bMonitoring And Not terminateFlag Then
        bMonitoring = True
        VerifyChange
        bMonitoring = False
    End If
End Sub
