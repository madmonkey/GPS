Attribute VB_Name = "modCommmon"
Option Explicit

Private Const cSep = ";"
Private Const cEqual = "="
Private Const dbLangGeneral = ";LANGID=0x0409;CP=1252;COUNTRY=0"
Private Const dbEncrypt = 2&
Private Declare Sub CoFreeUnusedLibraries Lib "ole32" ()
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private m_PropertyBag As VBRUN.PropertyBag
Private Const OFS_MAXPATHNAME = 128
Private Const OF_EXIST = &H4000
    
Type OFSTRUCT
   cBytes As Byte
   fFixedDisk As Byte
   nErrCode As Integer
   Reserved1 As Integer
   Reserved2 As Integer
   szPathName(OFS_MAXPATHNAME) As Byte
End Type
Private Const cModuleName = "basMDB"
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function getTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long


Public Function CompactRepair(ByVal ConnectionString As String) As Boolean
Dim vCol As Collection
Dim oObj As Object
Dim sTmp As String
Dim sObjType As String
Dim iDBType As Integer
Dim sSource As String, sDest As String
Dim resArray() As Byte
Dim systemSecureFile As String
'CALLING ROUTINE IS RESPONSIBLE FOR INSURING THAT ACTIVE CONNECTIONS ARE CLOSED
'UPDATED TO ACCOUNT FOR SECURED DATABASES, REMOVED RETRY ROUTINE USING DAO FOR SAKE OF TIME...
'CAN REVIST AT A LATER DATE TO RETRY USING DAO REPAIR METHODS
On Error GoTo Err_ConnectionString
    UEH_Log "modCompact", "CompactRepair", "Begin", logVerbose
    iDBType = 0
    sObjType = "JRO.JetEngine"
    Set vCol = ConnectValues(ConnectionString)
Retry:
On Error GoTo Err_CompactRepair
    sTmp = ReplaceExtension(RetrieveCollectionValue(vCol, "DATA SOURCE"), "ldb")
    UEH_Log "modCompact", "CompactRepair", "Deleting Lock File: " & sTmp, logVerbose
    If Dir(sTmp) <> vbNullString Then Kill sTmp
    sTmp = ReplaceExtension(RetrieveCollectionValue(vCol, "DATA SOURCE"), "compact")
    UEH_Log "modCompact", "CompactRepair", "Creating Temp DB: " & sTmp, logVerbose
    If Dir(sTmp) <> vbNullString Then Kill sTmp
    UEH_Log "modCompact", "CompactRepair", "Creating Engine: " & sObjType, logVerbose
    Set oObj = CreateObject(sObjType)
    UEH_Log "modCompact", "CompactRepair", "Created Engine: " & sObjType, logVerbose
    sSource = "Provider" & cEqual & RetrieveCollectionValue(vCol, "PROVIDER") & ";Data Source" & cEqual & RetrieveCollectionValue(vCol, "DATA SOURCE")
    sDest = "Provider" & cEqual & RetrieveCollectionValue(vCol, "PROVIDER") & ";" & "Data Source=" & sTmp
    If GetProperty("UseAuth") Then
        '''sSource = sSource & ";User Id = " & vCol("User Id") & ";Password = " & _
                vCol("Password") & ";Jet OLEDB:System Database =" & vCol("Jet OLEDB:System Database")
        resArray = GetProperty("secFile")
        systemSecureFile = CreateSecurityFile(resArray)
        sSource = sSource & ";User Id = " & GetProperty("UID") & ";Password = " & _
                GetProperty("PWD") & ";Jet OLEDB:System Database =" & systemSecureFile
    End If
    If GetProperty("UseDBPass") Then
        sSource = sSource & ";Jet OLEDB:Database Password =" & GetProperty("dbpass") 'vCol("Jet OLEDB:Database Password")
        sDest = sDest & ";Jet OLEDB:Database Password =" & GetProperty("dbpass") 'vCol("Jet OLEDB:Database Password")
    End If
    If GetProperty("UseEncrypt") Then
        sDest = sDest & ";Jet OLEDB:Encrypt Database=True"
    End If
    
    Select Case sObjType
        Case "JRO.JetEngine"
            oObj.CompactDatabase sSource, sDest
    End Select
    UEH_Log "modCompact", "CompactRepair", "Successful Repair, Renaming Files", logVerbose
    CopyFile sTmp, RetrieveCollectionValue(vCol, "DATA SOURCE"), 0&
    CompactRepair = True
    GoTo Exit_Routine

Err_CompactRepair:
    If Err.Number = -2147467259 Or Err.Number = -2147221164 Or Err.Number = 429 Then 'Unspecified/DB Error OR Class Not Registered OR Can't Create Object
        UEH_Log "modCompact", "CompactRepair", "Repair Failed with Error(" & Err.Number & " : " & Err.Description & ")", logWarning
        Select Case iDBType
            Case Else
                'We're doomed....
                CompactRepair = False
                GoTo Exit_Routine
        End Select
        GoTo Retry
    Else 'Some other type Error Occurred
        UEH_LogError "modCompact Err_CompactRepair", "CompactRepair", Err
    End If
    CompactRepair = False
Exit_Routine:
    Set oObj = Nothing
    Set vCol = Nothing
    If FileExists(sTmp) Then Kill sTmp
    CoFreeUnusedLibraries
    UEH_Log "modCompact", "CompactRepair", "End", logVerbose
    Exit Function
Err_ConnectionString:
    CompactRepair = False
    UEH_LogError "modCompact Err_ConnectionString", "CompactRepair", Err
End Function

Private Function RetrieveCollectionValue(ByRef Col As Collection, ByRef Key As String) As String
On Error GoTo err_RetrieveCollectionValue
    RetrieveCollectionValue = Col(Key)
    Exit Function
err_RetrieveCollectionValue:
    UEH_Log "modCompact", "RetrieveCollectionValue", "Key value: " & Key & " NOT found in collection!", logError
    Err.Clear
    RetrieveCollectionValue = vbNullString
End Function

Public Function ConnectValues(ByVal ConnectionString As String) As Collection
'KEY VALUES ARE RETURNED IN UPPER CASE
Dim vArray As Variant
Dim vCol As New Collection
Dim X As Long
    
    vArray = Split(ConnectionString, cSep)
    For X = LBound(vArray) To UBound(vArray)
        If vArray(X) <> vbNullString Then
        vCol.Add Mid$(vArray(X), InStr(1, vArray(X), cEqual) + Len(cEqual), _
                Len(vArray(X)) - InStr(1, vArray(X), cEqual)), _
                UCase(Left$(vArray(X), InStr(1, vArray(X), cEqual) - Len(cEqual)))
        End If
    Next
    Set ConnectValues = vCol

End Function

Private Function ReplaceExtension(ByVal sPath As String, ByVal sExt As String) As String

ReplaceExtension = Left$(sPath, (InStrRev(sPath, "\"))) & _
        Mid$(sPath, InStrRev(sPath, "\") + Len("\"), InStrRev(sPath, ".") - InStrRev(sPath, "\")) & sExt

End Function


Public Function FileExists(ByVal strSearchFile As String) As Boolean
    Dim strucFname As OFSTRUCT
    
    FileExists = (OpenFile(strSearchFile, strucFname, OF_EXIST) <> -1)
    
End Function

Public Function LetPropertyBag(ByRef myBag() As Byte)
    Set m_PropertyBag = New VBRUN.PropertyBag
    If IsValidArray(myBag) Then
        If UBound(myBag) > 0 Then
        m_PropertyBag.Contents = myBag()
        End If
    End If
End Function

Public Function GetProperty(ByVal sName As String) As Variant
On Error GoTo err_GetProperty
    If Not m_PropertyBag Is Nothing Then 'no objects so no set/let
        GetProperty = m_PropertyBag.ReadProperty(sName, GetDefault(sName))
    End If
err_GetProperty:
End Function

Public Function LetProperty(ByVal sName As String, ByVal vValue As Variant) As Variant
On Error GoTo err_LetProperty
    If Not m_PropertyBag Is Nothing Then
        m_PropertyBag.WriteProperty sName, vValue
    End If
err_LetProperty:
End Function

Public Function GetPropertyBag(Optional clearContents As Boolean = False) As Byte()
    If Not m_PropertyBag Is Nothing Then
        UEH_Log cModuleName, "GetProperties", "Retrieving cached settings", logVerbose
        GetPropertyBag = m_PropertyBag.Contents
        If clearContents Then Set m_PropertyBag = Nothing
    End If
End Function

Private Function GetDefault(ByVal sName As String) As Variant
Dim fh As Long
Dim fileName As String
Dim arrByte() As Byte
    Select Case LCase$(sName)
        Case "useauth"
            GetDefault = False
        Case "usedbpass"
            GetDefault = False
        Case "useencrypt"
            GetDefault = True
        Case "uid"
            GetDefault = "Admin"
        Case "pwd"
            GetDefault = vbNullString
        Case "secfile"
            GetDefault = arrByte
        Case "dbpass"
            GetDefault = ""
    End Select
End Function

Private Function CreateSecuredConnection(ByVal databaseFile As String, ByVal oledbProvider As String) As Boolean
Dim cnnString As String
Dim bUseAuth As Boolean, bUseDBPass As Boolean, bUseAccessEncrypt As Boolean
Dim UserID As String, userPass As String, dbPass As String
Dim resArray() As Byte
Dim systemSecureFile As String
'''    UEH_Log cModuleName, "CreateSecuredConnection", "Method Start", logVerbose
    cnnString = "Provider=" & oledbProvider & ";" & "Data Source=" & databaseFile & ";"
    bUseAuth = GetProperty("UseAuth")
    If bUseAuth Then
        UserID = GetProperty("UID")
        userPass = GetProperty("PWD")
        resArray = GetProperty("secFile")
        systemSecureFile = CreateSecurityFile(resArray)
        cnnString = cnnString & "User Id=" & Chr$(34) & UserID & Chr$(34) & ";Password=" & Chr$(34) & userPass & Chr$(34) & ";Jet OLEDB:System Database=" & Chr$(34) & systemSecureFile & Chr$(34) & ";"
    End If
    bUseDBPass = GetProperty("UseDBPass")
    If bUseDBPass Then
        dbPass = GetProperty("DBPass")
        cnnString = cnnString & "Jet OLEDB:Database Password=" & Chr$(34) & dbPass & Chr$(34) & ";"
    End If
    bUseAccessEncrypt = GetProperty("UseEncrypt")
    
On Error GoTo err_CreateSecuredConnection
    Set g_Connection = New ADODB.Connection
    
    With g_Connection
        .Provider = oledbProvider
        If bUseAuth Then .Properties("Jet OLEDB:System database") = systemSecureFile
        .Open cnnString
    End With
    CreateSecuredConnection = True
    UEH_Log cModuleName, "CreateSecuredConnection", "Created Secure Connection!", logVerbose
    Exit Function

err_CreateSecuredConnection:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    UEH_Log cModuleName, "CreateSecuredConnection", "Method Error", logError
    CreateSecuredConnection = False
End Function

Public Function CreateSecurityFile(ByRef arrByte() As Byte) As String
Dim fh As Long
Dim fileName As String
Dim Buffer As String
    If IsValidArray(arrByte) Then
        fh = FreeFile
        fileName = getTemporaryFile
        Buffer = arrByte()
        Open fileName For Binary Access Write As #fh
        Put #fh, , Buffer
        Close #fh
        CreateSecurityFile = fileName
    End If
End Function

Private Function getTemporaryFile() As String
Dim sTemp As String
Const FILE_ATTRIBUTE_TEMPORARY = &H100
    sTemp = String(260, 0)
    'If you got to this point - with a defined temporary workspace that doesn't exist
    'It is only through the grace of God - we'll try to help out here
    If Dir(Environ("TEMP"), vbDirectory) = vbNullString Then MkDir Environ("TEMP")
    getTempFileName Environ("TEMP"), "GPS", 0, sTemp
    sTemp = Left$(sTemp, InStr(1, sTemp, Chr$(0)) - 1)
    SetFileAttributes sTemp, FILE_ATTRIBUTE_TEMPORARY
    getTemporaryFile = sTemp
End Function

Public Function IsValidArray(ByRef this As Variant) As Boolean
    If IsArray(this) Then
        IsValidArray = GetArrayDimensions(VarPtrArray(this)) >= 1
    Else
        IsValidArray = False
    End If
End Function
Private Function GetArrayDimensions(ByVal arrPtr As Long) As Integer
   Dim address As Long
   
   CopyMemory address, ByVal arrPtr, ByVal 4   'get the address of the SafeArray structure in memory
   If address <> 0 Then 'if there is a dimension, then address will point to the memory address of the array, otherwise the array isn't dimensioned
      CopyMemory GetArrayDimensions, ByVal address, 2 'fill the local variable with the first 2 bytes of the safearray structure. These first 2 bytes contain an integer describing the number of dimensions
   End If

End Function

Private Function VarPtrArray(arr As Variant) As Long

  'Function to get pointer to the array
   CopyMemory VarPtrArray, ByVal VarPtr(arr) + 8, ByVal 4
    
End Function
