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
    ikPWHIPUserID = 7
    ikUseDynamicIP = 8
End Enum
Private m_Session As String
Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
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

Private Sub Class_Initialize()
    INIFile = Environ("WINDIR") & "\mcs.ini"
End Sub

Public Function GetValue(ByVal key As iniKeys) As Variant
Dim vDefault As Variant, NC As Long
Dim sKey As String, ret As String, Section As String, OKtoWrite As Boolean
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
        Case ikHTEUserID, ikPWHIPUserID
            vDefault = Left$(ComputerName, 4)
        Case ikConnectName
            vDefault = "MDB_CONNECTION"
        Case ikParagraph
            vDefault = "MDB_TCP_CONNECTION"
        Case ikSessions
            GetValue = ReturnSessionNames(GetINIFile(key))
            Exit Function
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
            ret = String(cMax, 0)
            NC = GetPrivateProfileString(Section, sKey, CStr(vDefault), ret, cMax, myINI)
            If NC <> 0 Then GetValue = Trim$(Left$(ret, NC))
    End Select
    If DataType <> vbBoolean And OKtoWrite Then WritePrivateProfileString Section, sKey, Trim$(CStr(GetValue)), myINI
End Function

Private Function GetINIFile(ByVal key As iniKeys) As String
    
    Select Case key
        Case ikSessionINIKey
            GetINIFile = Environ("WINDIR") & "\" & GetValue(ikSessionINI)
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
    bOKtoWrite = True
    Select Case key
        Case ikSessionType
            GetKeyValue = "SesType": iniType = vbString
            bOKtoWrite = False
        Case ikSessionINI
            GetKeyValue = "SesIni": iniType = vbString
            bOKtoWrite = False
        Case ikSessionINIKey
            GetKeyValue = "MessageKeys": iniType = vbString
            bOKtoWrite = False
        Case ikHTEUserID
            GetKeyValue = "HTEUserID": iniType = vbString
            bOKtoWrite = True
        Case ikConnectName
            GetKeyValue = "ConnectName": iniType = vbString
            bOKtoWrite = True
        Case ikParagraph
            GetKeyValue = "Paragraph": iniType = vbString
            bOKtoWrite = False
        Case ikPWHIPUserID
            GetKeyValue = "PWHIPUserID": iniType = vbString
            bOKtoWrite = False
    End Select
    
End Function

Private Function GetSection(ByVal key As iniKeys) As String
    Select Case key
        Case ikSessionType, ikSessionINI, ikConnectName
            GetSection = SessionName
        Case ikSessionINIKey
            GetSection = "General"
        Case ikParagraph
            GetSection = GetValue(ikConnectName)
        Case ikHTEUserID, ikPWHIPUserID
            GetSection = GetValue(ikParagraph)
    End Select
End Function
Private Function ReturnSessionNames(ByVal sFileName As String) As Variant
Dim szBuf As String, Length As Integer
Dim SectionArr() As String, i As Integer, sReturn As String
Const cMax = 8192
    szBuf = String$(cMax, 0)
    Length = GetPrivateProfileSectionNames(szBuf, cMax, sFileName)
    szBuf = Left$(szBuf, Length)
    SectionArr = Split(szBuf, vbNullChar)
    For i = 0 To UBound(SectionArr)
        'Test for connection name
        szBuf = String$(cMax, 0)
        If GetPrivateProfileString(SectionArr(i), "ConnectName", vbNullString, szBuf, cMax, sFileName) > 0 Then
            sReturn = sReturn & SectionArr(i) & "|*|"
        End If
    Next
    ReturnSessionNames = sReturn
End Function


