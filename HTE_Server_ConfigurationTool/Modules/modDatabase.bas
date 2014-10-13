Attribute VB_Name = "modDatabase"
Option Explicit

Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private m_INI As String

Public Function PrepareDatabaseConnection(Optional ReadOnly As Boolean = True) As ADODB.Connection
Dim m_Conn As ADODB.Connection
    Set m_Conn = New ADODB.Connection
    With m_Conn
        .ConnectionString = "Provider=" & getStoredValue("oledbProvider") & ";" & "Data Source=" & getStoredValue("Data Source") & ";"
        .ConnectionTimeout = getStoredValue("ConnectionTimeout")
        .Open
    End With
    
    Set PrepareDatabaseConnection = m_Conn
    
End Function

Public Function PrepareRecordset(ByRef SQL As String, ByRef Connection As ADODB.Connection, Optional ReadOnly As Boolean = True) As ADODB.Recordset
Dim rs As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    With rs
        .CursorLocation = adUseServer
        .CursorType = adOpenForwardOnly
        .ActiveConnection = Connection
        .Open SQL, Connection, adOpenForwardOnly, adLockReadOnly
    End With
    Set PrepareRecordset = rs
End Function
Private Function RetrieveSetting(ByRef Section As String, ByRef KeyName As String, Optional ByRef default As String = vbNullString, Optional ByRef bNumeric As Boolean = False) As String
Dim ret&, buff$
Const MAX_PATH = 265
    If m_INI = vbNullString Then m_INI = App.Path & IIf(Right$(App.Path, 1) = "\", vbNullString, "\") & "GPSConfig.ini"
    buff = String$(MAX_PATH, vbNullChar)

    ret = GetPrivateProfileString(Section, KeyName, default, buff, MAX_PATH, m_INI)
    If ret <> 0 Then RetrieveSetting = Left$(buff, ret)

    
End Function

Private Function getSection(ByRef KeyName As String) As String
    Select Case KeyName
        Case "oledbProvider", "Data Source", "ConnectionTimeout"
            getSection = "Database"
    End Select
End Function

Private Function getKeyName(ByRef KeyName As String) As String
Select Case KeyName
        Case "oledbProvider"
            getKeyName = "oledbProvider"
        Case "Data Source"
             getKeyName = "DataSource"
        Case "ConnectionTimeout"
            getKeyName = "Timeout"
    End Select
End Function

Private Function getDefault(ByRef KeyName As String, Optional ByRef bNumeric As Boolean = False) As String
    Select Case KeyName
        Case "oledbProvider"
            getDefault = "Microsoft.Jet.OLEDB.4.0"
        Case "Data Source"
            getDefault = App.Path & IIf(Right$(App.Path, 1) = "\", vbNullString, "\") & "Database\GPS.MDB"
        Case "ConnectionTimeout"
            getDefault = 120: bNumeric = True
    End Select
End Function

Private Function getStoredValue(ByRef KeyName As String) As String
Dim bNumeric As Boolean
    getStoredValue = RetrieveSetting(getSection(KeyName), getKeyName(KeyName), getDefault(KeyName, bNumeric), bNumeric)
End Function
