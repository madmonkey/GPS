Attribute VB_Name = "modCompact"
Option Explicit
Private Const cSep = ";"
Private Const cEqual = "="
Private Const dbLangGeneral = ";LANGID=0x0409;CP=1252;COUNTRY=0"
Private Const dbEncrypt = 2&
Private Declare Sub CoFreeUnusedLibraries Lib "ole32" ()
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Const OFS_MAXPATHNAME = 128
Private Const OF_EXIST = &H4000
Private Type OFSTRUCT
   cBytes As Byte
   fFixedDisk As Byte
   nErrCode As Integer
   Reserved1 As Integer
   Reserved2 As Integer
   szPathName(OFS_MAXPATHNAME) As Byte
End Type
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long

Private Function FileExists(ByVal strSearchFile As String) As Boolean
    Dim strucFname As OFSTRUCT
    FileExists = (OpenFile(strSearchFile, strucFname, OF_EXIST) <> -1)
End Function
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
'''    UEH_Log "modCompact", "CompactRepair", "Deleting Lock File: " & sTmp, logVerbose
    If Not FileExists(sTmp) Then
        sTmp = ReplaceExtension(RetrieveCollectionValue(vCol, "DATA SOURCE"), "compact")
        UEH_Log "modCompact", "CompactRepair", "Creating Temp DB: " & sTmp, logVerbose
        If FileExists(sTmp) Then Kill sTmp
        UEH_Log "modCompact", "CompactRepair", "Creating Engine: " & sObjType, logVerbose
        Set oObj = CreateObject(sObjType)
        UEH_Log "modCompact", "CompactRepair", "Created Engine: " & sObjType, logVerbose
        sSource = "Provider" & cEqual & RetrieveCollectionValue(vCol, "PROVIDER") & ";Data Source" & cEqual & RetrieveCollectionValue(vCol, "DATA SOURCE")
        sDest = "Provider" & cEqual & RetrieveCollectionValue(vCol, "PROVIDER") & ";" & "Data Source=" & sTmp
    '''    If GetProperty("UseAuth") Then
    '''        '''sSource = sSource & ";User Id = " & vCol("User Id") & ";Password = " & _
    '''                vCol("Password") & ";Jet OLEDB:System Database =" & vCol("Jet OLEDB:System Database")
    '''        resArray = GetProperty("secFile")
    '''        systemSecureFile = CreateSecurityFile(resArray)
    '''        sSource = sSource & ";User Id = " & GetProperty("UID") & ";Password = " & _
    '''                GetProperty("PWD") & ";Jet OLEDB:System Database =" & systemSecureFile
    '''    End If
    '''    If GetProperty("UseDBPass") Then
        'Debug.Print sSource
'        sSource = sSource & ";" & RetrieveCollectionValue(vCol, (Ucase$("Jet OLEDB:Database Password")))
'        Debug.Print sSource
'        sDest = sDest & ";" & RetrieveCollectionValue(vCol, ("Jet OLEDB:Database Password"))
'        sSource = sSource & ";Jet OLEDB:Database Password =" & RetrieveCollectionValue(vCol, (UCase$("Jet OLEDB:Database Password"))) '& GetProperty("dbpass") 'vCol("Jet OLEDB:Database Password")
'        Debug.Print sSource
'        sDest = sDest & ";Jet OLEDB:Database Password =" & RetrieveCollectionValue(vCol, (UCase$("Jet OLEDB:Database Password"))) 'GetProperty("dbpass") 'vCol("Jet OLEDB:Database Password")
    '''    End If
    '''    If GetProperty("UseEncrypt") Then
    '''        sDest = sDest & ";Jet OLEDB:Encrypt Database=True"
    '''    End If
        
        Select Case sObjType
            Case "JRO.JetEngine"
                oObj.compactdatabase sSource, sDest
        End Select
        UEH_Log "modCompact", "CompactRepair", "Successful Repair, Renaming Files", logVerbose
        CopyFile sTmp, RetrieveCollectionValue(vCol, "DATA SOURCE"), 0&
        CompactRepair = True
    Else
        UEH_Log "modCompact", "CompactRepair", "Cannot obtain exclusive lock - aborting.", logInformation
    End If
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

Private Function ConnectValues(ByVal ConnectionString As String) As Collection
'KEY VALUES ARE RETURNED IN UPPER CASE
Dim vArray As Variant
Dim vCol As New Collection
Dim X As Long
    
    vArray = Split(ConnectionString, cSep)
    For X = LBound(vArray) To UBound(vArray)
        If vArray(X) <> vbNullString Then
        vCol.Add Mid$(vArray(X), InStr(1, vArray(X), cEqual) + Len(cEqual), _
                Len(vArray(X)) - InStr(1, vArray(X), cEqual)), _
                UCase$(Left$(vArray(X), InStr(1, vArray(X), cEqual) - Len(cEqual)))
        'Debug.Print "ADDED " & Mid$(vArray(X), InStr(1, vArray(X), cEqual) + Len(cEqual), _
                Len(vArray(X)) - InStr(1, vArray(X), cEqual)) & " with [KEY VALUE = " & UCase$(Left$(vArray(X), InStr(1, vArray(X), cEqual) - Len(cEqual))) & "]" & vArray(X)
        End If
    Next
    Set ConnectValues = vCol

End Function

Private Function ReplaceExtension(ByVal sPath As String, ByVal sExt As String) As String

ReplaceExtension = Left$(sPath, (InStrRev(sPath, "\"))) & _
        Mid$(sPath, InStrRev(sPath, "\") + Len("\"), InStrRev(sPath, ".") - InStrRev(sPath, "\")) & sExt

End Function


