VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Alias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Binary

Implements ccrpTimers6.ICcrpTimerNotify
Private m_TimeToVerify As Long 'HOW LONG TO CACHE MAC/ADDR BEFORE INQUIRING AGAIN
Private m_Tag As String 'HOW TO EMBED LOOKUP
Private m_Delimiter As String 'HOW TO FORMAT MAC
'LOOKUP STACK IMPLEMENTATION
Private WithEvents eventTimer As ccrpTimers6.ccrpTimer
Attribute eventTimer.VB_VarHelpID = -1
Private m_LookupQueue As cStack
Private LookupQueue As IStack
Private bLookupsWorking As Boolean
Public Event LookupResolved(ByRef cMsg As HTE_GPS.GPSMessage) ', ByRef LookupValue As String)
Attribute LookupResolved.VB_Description = "Event triggered after a succesful address resolution"
Public Event LookupFailed(ByRef cMsg As HTE_GPS.GPSMessage) ', ByRef LookupValue As String)
Attribute LookupFailed.VB_Description = "Event triggered after an unsuccessful resolution"
Private m_EnableMacResolution As Boolean
Private m_Dictionary As Scripting.Dictionary
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Property Let EnableMacResolution(ByVal Data As Boolean)
Attribute EnableMacResolution.VB_Description = "Whether or not to ARP remote Address to obtain a physical identifier "
    m_EnableMacResolution = Data
End Property

Public Property Get EnableMacResolution() As Boolean
    EnableMacResolution = m_EnableMacResolution
End Property

Public Property Let MacLookupExpiration(ByVal Data As Long)
Attribute MacLookupExpiration.VB_Description = "How long to cache a succesful ARP lookup (in seconds)"
    m_TimeToVerify = Data
End Property

Public Property Get MacLookupExpiration() As Long
    MacLookupExpiration = m_TimeToVerify
End Property

Private Sub Class_Initialize()

    Set eventTimer = New ccrpTimers6.ccrpTimer
    With eventTimer
        .EventType = TimerPeriodic
        .Interval = 1
        Set .Notify = Me
        .Enabled = False
    End With
    Set m_LookupQueue = New cStack
    Set LookupQueue = m_LookupQueue
    m_EnableMacResolution = False
    m_Delimiter = "-"
    m_TimeToVerify = 3600 'defaults to an hour
    Set m_Dictionary = New Scripting.Dictionary
    With m_Dictionary
        .CompareMode = BinaryCompare
    End With
End Sub

Private Sub Class_Terminate()
    
    eventTimer.Enabled = False
    Set eventTimer = Nothing
    Set LookupQueue = Nothing
    
End Sub
Public Sub ResolveEntity(ByRef cMsg As HTE_GPS.GPSMessage, ByRef IPAddress As String)
Attribute ResolveEntity.VB_Description = "Method used to modify entity field in the Message - will raise either LookupResolved or LookupFailed events"
    If Not m_EnableMacResolution Then
        ResolveAddrID cMsg, IPAddress
    Else
        ResolvePhyID cMsg, IPAddress
    End If
End Sub
Private Sub ResolvePhyID(ByRef cMsg As HTE_GPS.GPSMessage, ByRef Lookup As String)

    If Not m_LookupQueue Is Nothing Then
        cMsg.Entity = Lookup
        LookupQueue.Push cMsg
        eventTimer.Enabled = True
    End If

End Sub

Private Sub ResolveAddrID(ByRef cMsg As HTE_GPS.GPSMessage, ByRef Lookup As String)
Dim cEnt As Entity, xmlProp As HTE_Properties.XMLPropertyBag
    Set cEnt = New Entity
    With cEnt
        .IPAddress = Lookup
    End With
    Set xmlProp = New HTE_Properties.XMLPropertyBag
    With cMsg
        xmlProp.SaveState cEnt
        .Entity = xmlProp.Contents
    End With
    Set cEnt = Nothing
    Set xmlProp = Nothing
    RaiseEvent LookupResolved(cMsg)
End Sub

Private Sub ICcrpTimerNotify_Timer(ByVal Milliseconds As Long)
Dim cMsg As HTE_GPS.GPSMessage 'cMessage
Dim cachedAlias As cachedLookup
Dim sReturn As String
Dim bCachedValue As Boolean
Dim cEnt As Entity
Dim Lookup As String
    If Not bLookupsWorking Then
        bLookupsWorking = True
        Do While LookupQueue.Count > 0
            sReturn = vbNullString: bCachedValue = False
            Set cMsg = LookupQueue.Pop
            Set cEnt = New Entity
            With cMsg
                Lookup = .Entity
                If m_EnableMacResolution Then 'only lookup if nessesary
                    If m_Dictionary.Exists(Lookup) Then
                        Set cachedAlias = m_Dictionary.Item(Lookup)
                        If Abs(timeGetTime - cachedAlias.Created) / 1000 <= m_TimeToVerify Then
                            sReturn = cachedAlias.Alias: bCachedValue = True
                        Else
                            m_Dictionary.Remove Lookup
                            sReturn = GetPhysicalID(Lookup)
                        End If
                    Else
                        sReturn = GetPhysicalID(Lookup)
                    End If
                    cEnt.MACAddress = sReturn
                End If
                cEnt.IPAddress = Lookup
                .Entity = cEnt.Serialize
                Select Case Len(sReturn)
                    Case 0 '"{Not Available}"
                        RaiseEvent LookupFailed(cMsg)
                    Case Else
                        RaiseEvent LookupResolved(cMsg)
                        If Not bCachedValue Then
                            Set cachedAlias = New cachedLookup: cachedAlias.Alias = sReturn
                            m_Dictionary.Add Lookup, cachedAlias
                        End If
                End Select
            End With
        Loop
        bLookupsWorking = False
    End If
    eventTimer.Enabled = False
End Sub
