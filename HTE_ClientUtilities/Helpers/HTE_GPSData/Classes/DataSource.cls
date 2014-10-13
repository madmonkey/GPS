VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Binary

Private Const cModuleName = "DataSource"
Private appMutex As Mutex 'treat the mdb database as a critical resource
Private m_ConnectionString As String
Public Enum gpsAliasField
    gafIdentityID = 0
    gafPhysicalLookup = 1
    gafAddressLookup = 2
    gafAppLookup = 3
    gafDevice = 4
    gafAlias = 5
    gafComments = 6
End Enum
Private Const cDataSubFolder = "\Data"
Private Const cDataFile = "\identity.mdb"
Private m_Publish As HTE_PubData.Publisher
Private Const cAliasUpdateTopic As String = "CommunityConnectAliasChanges"
Private m_DatabaseAvailable As Boolean
Private m_InstallPath As String
Private Const OF_EXIST = &H4000
Private Const OFS_MAXPATHNAME = 128
Private Type OFSTRUCT
   cBytes As Byte
   fFixedDisk As Byte
   nErrCode As Integer
   Reserved1 As Integer
   Reserved2 As Integer
   szPathName(OFS_MAXPATHNAME) As Byte
End Type
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private WithEvents datasourceWatcher As FileSystemWatch.Watcher
Attribute datasourceWatcher.VB_VarHelpID = -1
Private m_IgnoreSelfChanges As Boolean
Private Sub Class_Initialize()
    UEH_BeginLogging App.EXEName, App.EXEName
    m_InstallPath = GetInstalledPath
    m_ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
         m_InstallPath & cDataSubFolder & cDataFile & ";User Id=admin;Password=;" '& _
         "Jet OLEDB:Database Password= " & Chr$(34) & Chr$(190) & Chr$(169) & Chr$(112) & Chr$(213) & Chr$(166) & Chr$(121) & Chr$(189) & Chr$(239) & Chr$(63) & Chr$(63) & Chr$(91) & Chr$(24) & Chr$(157) & Chr$(12) & Chr$(73) & Chr$(143) & Chr$(53) & Chr$(63) & Chr$(69) & Chr$(89) & Chr$(34)
         '& Chr$(39) & Chr$(2) & Chr$(125) & Chr$(199) & Chr$(163) & Chr$(85) & Chr$(251) & Chr$(8) & Chr$(109) & Chr$(248) & Chr$(16) & Chr$(84) & Chr$(63) & Chr$(238) & Chr$(49) & Chr$(241) & Chr$(174) & Chr$(52) & Chr$(162) & Chr$(63) & Chr$(39)
    
    If IsDatabasePresent Then
        UEH_Log "DataSource", "Initialize", "Database Present = " & CStr(m_DatabaseAvailable), logDebug
        Set appMutex = New Mutex
        'lock-up occuring try instance specific
        appMutex.ConstructMutex "{9DF3034E-BF43-4e69-94A3-D410B890128E}"
        Set m_Publish = New HTE_PubData.Publisher
        With m_Publish
            .Topic = cAliasUpdateTopic
            .Timeout = 30000
        End With
        Set datasourceWatcher = New FileSystemWatch.Watcher
        UEH_Log "DataSource", "Initialize", "Monitoring [" & m_InstallPath & cDataSubFolder & "]", logVerbose
        datasourceWatcher.BeginMonitoring m_InstallPath & cDataSubFolder, "*.mdb"
    End If
    
End Sub

Private Function GetInstalledPath() As String
Dim reg As HTE_SystemUtility.Registry
    If Not InDevelopment Then
        Set reg = New HTE_SystemUtility.Registry
        With reg
            .ClassKey = HKEY_LOCAL_MACHINE
            .SectionKey = "Software\HTE\Modular GPS\"
            .ValueKey = "Install_Path"
            .ValueType = REG_SZ
            If .KeyExists Then
                GetInstalledPath = .Value
            Else
                GetInstalledPath = Environ("PROGRAMFILES") & "\HTE\Modular Gps"
            End If
        End With
    Else
        GetInstalledPath = GetParentPath
    End If
End Function

Private Function FileExists(ByVal strSearchFile As String) As Boolean
    Dim strucFname As OFSTRUCT
    FileExists = (OpenFile(strSearchFile, strucFname, OF_EXIST) <> -1)
End Function
Public Function IsDatabasePresent() As Boolean
    If m_InstallPath = vbNullString Then m_InstallPath = GetInstalledPath
    UEH_Log cModuleName, "IsDatabasePresent", "InstalledPath = " & m_InstallPath, logDebug
    m_DatabaseAvailable = FileExists(m_InstallPath & cDataSubFolder & cDataFile)
    IsDatabasePresent = m_DatabaseAvailable
End Function
Private Sub Class_Terminate()
    Cleanup
    Set m_Publish = Nothing
    If Not appMutex Is Nothing Then
        appMutex.DestroyMutex
        Set appMutex = Nothing
    End If
    If Not datasourceWatcher Is Nothing Then
        datasourceWatcher.StopMonitoring
        Set datasourceWatcher = Nothing
    End If
    UEH_EndLogging
End Sub

Private Function Cleanup()
'    If Not m_Conn Is Nothing Then
'        If m_Conn.State = adStateOpen Then m_Conn.Close
'        Set m_Conn = Nothing
'    End If
End Function
Private Function Connection() As ADODB.Connection
    Dim m_Conn As ADODB.Connection
    Set m_Conn = New ADODB.Connection
    With m_Conn
        .ConnectionTimeout = 15
        .CommandTimeout = 30
        .CursorLocation = adUseServer
        .IsolationLevel = adXactReadUncommitted
        .Mode = adModeReadWrite
        .ConnectionString = m_ConnectionString
    End With
    Set Connection = m_Conn
End Function
Public Property Get GridView(Optional Editable As Boolean = True, Optional SortBy As gpsAliasField = gafAddressLookup, _
    Optional Ascending As Boolean = True) As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim strSelect As String
    
On Local Error Resume Next
    If IsDatabasePresent Then
        Set rs = New ADODB.Recordset
        strSelect = "SELECT * FROM ALIAS " & _
            "ORDER BY " & Choose(SortBy + 1, "AliasID", "PhysicalLookup", "AddressLookup", "AppLookup", "Device", "Alias", "Comments") & _
            IIf(Ascending, " ASC", " DESC")
            If Editable Then
'                If appMutex.RequestMutex(14500) Then 'just short of default connection timeout
                    rs.Open strSelect, m_ConnectionString, adOpenKeyset, adLockBatchOptimistic
'                    If Not appMutex Is Nothing Then appMutex.DiscardMutex
'                Else
'                    UEH_Log cModuleName, "GridView", "Cannot obtain mutex - try again later.", logWarning
'                End If
            Else
                rs.Open strSelect, m_ConnectionString, adOpenForwardOnly, adLockReadOnly
            End If
        Set GridView = rs
    End If
End Property

Public Sub NotifyOfChanges(Optional ByVal ChangedSource = "GPSInterface")
On Local Error Resume Next
    If m_Publish Is Nothing Then Exit Sub
    With m_Publish
        .SendString ChangedSource, "Entities"
    End With
End Sub

Public Function SubscribeToChanges() As HTE_PubData.Subscriber
Dim m_Sub As HTE_PubData.Subscriber
    Set m_Sub = New HTE_PubData.Subscriber
    With m_Sub
        .Topic = cAliasUpdateTopic
    End With
    Set SubscribeToChanges = m_Sub
End Function

Public Function MaintainDB() As Boolean
Dim bClosed As Boolean
    If IsDatabasePresent Then
        If Not appMutex Is Nothing Then
            If appMutex.RequestMutex(250) Then
                Cleanup
                MaintainDB = CompactRepair(m_ConnectionString)
                appMutex.DiscardMutex
            Else
                UEH_Log cModuleName, "MaintainDB", "Cannot obtain mutex - try again later.", logWarning
            End If
        End If
    End If
End Function

Public Function Entities() As Scripting.Dictionary
Dim m_Dic As Scripting.Dictionary
Dim rs As ADODB.Recordset
Dim aliasInfoData As AliasData
Dim arrayList As Variant
Dim i As Integer
Dim ent As HTE_Entity.Entity
On Error GoTo err_Entities
    Set m_Dic = New Scripting.Dictionary
    arrayList = datasourceWatcher.ObtainAliasDataCollection

    For i = LBound(arrayList) To UBound(arrayList)
        Set ent = New HTE_Entity.Entity
        Set aliasInfoData = arrayList(i)

        ent.MACAddress = aliasInfoData.PhysicalLookUp
        ent.IPAddress = aliasInfoData.AddressLookUp
        ent.Alias = aliasInfoData.Alias
        ent.DeviceID = aliasInfoData.Device
        ent.Comments = aliasInfoData.Comments
        ent.ApplicationID = aliasInfoData.AppLookUp
        
        m_Dic.Add aliasInfoData.AliasId, ent
    Next
    
    Set Entities = m_Dic
    Exit Function
err_Entities:
    UEH_LogError cModuleName, "Entities", Err
End Function

Public Function AddEntity(ByRef ent As HTE_Entity.Entity) As Boolean
    'returns true if added or updated
    'would likely be called from translation (in which case
    'we should only add ApplicationID, Mac and/or IP IF they don't already exist!)
Dim cmd As ADODB.Command
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim bAdded As Boolean

On Error GoTo err_AddEntity
    If IsDatabasePresent Then
        m_IgnoreSelfChanges = True
        Set cn = Connection
        If Not cn Is Nothing Then
            cn.Open
            Set cmd = New ADODB.Command
            With cmd
                .ActiveConnection = cn
                If Not ent.IsProxy Then 'use device settings - otherwise useless
                    .CommandType = adCmdStoredProc
                    .CommandText = "spInsertEntity"
                    .Prepared = True
                    .Parameters.Append .CreateParameter("@PhysicalLookup", adVarChar, adParamInput, 17, IIf(Len(ent.MACAddress) = 0, Null, ent.MACAddress))
                    .Parameters.Append .CreateParameter("@AddressLookup", adVarChar, adParamInput, 39, IIf(Len(ent.IPAddress) = 0, Null, ent.IPAddress))
                    .Parameters.Append .CreateParameter("@AppLookup", adVarChar, adParamInput, 16, IIf(Len(ent.ApplicationID) = 0, Null, ent.ApplicationID))
                    .Execute , , adExecuteNoRecords
                    bAdded = True
                Else 'coming from centralized location (switch)
                    'check existence of AppID since NOT enforced unique in database!
                    'Add record only if not already added - doing it here since window of
                    'insert is between 20-30 seconds before process is made aware of it
                    'could modify database to account for (key unique) but am wary that
                    'would adversly affect (propagate) issue to other clients (i.e. "I have to change my database!")
                    If Len(ent.ApplicationID) > 0 Then
                        .CommandType = adCmdStoredProc
                        .CommandText = "spRetrieveByAppLookup"
                        .Prepared = True
                        .Parameters.Append .CreateParameter("@AppLookup", adVarChar, adParamInput, 16, ent.ApplicationID)
                        Set rs = New ADODB.Recordset
                        Set rs = .Execute
                        If rs.EOF Then 'doesn't exist
                            .CommandType = adCmdText 'thanks simon - no motorola - now adding dynamic SQL!
                            .CommandText = "INSERT INTO Alias (AppLookup, Device, Alias) " & _
                                "Values(" & ScrubDataValue(ent.ApplicationID) & _
                                    "," & ScrubDataValue(ent.ApplicationID) & _
                                    "," & ScrubDataValue(ent.Alias) & ")"
                            .Prepared = True
                            .Execute , , adExecuteNoRecords
                            bAdded = True
                        Else
                            'already exists!
                        End If
                        rs.Close
                        Set rs = Nothing
                    Else
                        'nothing to save from proxy!
                    End If
                End If
                If cn.State = adStateOpen Then cn.Close
            End With
            Set cn = Nothing
            Set cmd = Nothing
        End If
        If bAdded Then NotifyOfChanges "EntityAdded"
        AddEntity = bAdded
        m_IgnoreSelfChanges = False
    Else
        UEH_Log cModuleName, "AddEntity", "Unable to find database!", logError
    End If
    Exit Function
err_AddEntity:
    Dim cnError As ADODB.Error
    m_IgnoreSelfChanges = False
    For Each cnError In cn.Errors
        UEH_Log cModuleName, "AddEntity", cnError.Description, logError
    Next
    UEH_LogError cModuleName, "AddEntity", Err
On Error Resume Next
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    If Not cmd Is Nothing Then Set cmd = Nothing
End Function

Private Function ScrubDataValue(ByRef data As String) As String
    ScrubDataValue = "'" & Replace(data, "'", "''") & "'"
End Function

Private Sub datasourceWatcher_OnChangeNotification(ByVal sender As Variant, ByVal FileChangedEventArgs As FileSystemWatch.IFileChangedEventArgs)
UEH_Log cModuleName, "datasourceWatcher_OnChangeNotification", "Change detected", logVerbose
Select Case FileChangedEventArgs.ChangeType
    Case ChangeType.ChangeType_Changed
        If Not m_IgnoreSelfChanges Then
            NotifyOfChanges
            UEH_Log cModuleName, "datasourceWatcher_OnChangeNotification", "Notifying that database changes occured outside of application (.mdb file changed)", logVerbose
        Else
            UEH_Log cModuleName, "datasourceWatcher_OnChangeNotification", "Detected self-change - ignoring!", logVerbose
        End If
    Case Else
        UEH_Log cModuleName, "datasourceWatcher_OnChangeNotification", "[" & CStr(FileChangedEventArgs.ChangeType) + "] occured to [" & FileChangedEventArgs.Name & "]", logVerbose
End Select
End Sub

Private Sub datasourceWatcher_LogData(ByVal data As String)
    UEH_Log cModuleName, "datasourceWatcher_LogData", data, logError
End Sub
