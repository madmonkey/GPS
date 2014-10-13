VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "macResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_TimeToVerify As Long 'HOW LONG TO CACHE MAC/ADDR BEFORE INQUIRING AGAIN
Private m_Tag As String 'HOW TO EMBED LOOKUP
Private m_Delimiter As String 'HOW TO FORMAT MAC
'LOOKUP STACK IMPLEMENTATION
Private WithEvents eventTimer As ccrpTimers6.ccrpTimer
Attribute eventTimer.VB_VarHelpID = -1
Private m_LookupQueue As cStack
Private LookupQueue As IStack
Private bLookupsWorking As Boolean
Public Event LookupResolved(ByRef cMsg As HTE_GPS.GPSMessage)
Public Event LookupFailed(ByRef cMsg As HTE_GPS.GPSMessage, ByRef LookupValue As String)
Private Sub Class_Initialize()

    Set eventTimer = New ccrpTimers6.ccrpTimer
    With eventTimer
        .EventType = TimerPeriodic
        .Interval = 1
        .Enabled = False
    End With
    Set m_LookupQueue = New cStack
    Set LookupQueue = m_LookupQueue
    m_Tag = ";LOOKUP="
    m_Delimiter = "-"
End Sub

Private Sub Class_Terminate()
    
    eventTimer.Enabled = False
    Set eventTimer = Nothing
    Set LookupQueue = Nothing
    
End Sub

Public Sub ResolvePhyID(ByRef cMsg As cMessage, ByRef Lookup As String)

    If Not m_LookupQueue Is Nothing Then
        cMsg.Lookup = Lookup
        LookupQueue.Push cMsg
        eventTimer.Enabled = True
    End If

End Sub

Public Sub ResolveAddrID(ByRef cMsg As cMessage, ByRef Lookup As String)
    With cMsg
        .rawMessage = .rawMessage & m_Tag & Lookup
    End With
    RaiseEvent LookupResolved(cMsg)
End Sub

Private Sub eventTimer_Timer(ByVal Milliseconds As Long)
Dim i As Long
Dim cMsg As cMessage
Dim sReturn As String
    If Not bLookupsWorking Then
        bLookupsWorking = True
        Do While LookupQueue.Count > 0
            sReturn = vbNullString
            Set cMsg = LookupQueue.Pop
            With cMsg
                sReturn = GetPhysicalID(.Lookup)
                Select Case sReturn
                    Case "{Not Available}"
                        RaiseEvent LookupFailed(cMsg, .Lookup)
                    Case Else
                        .rawMessage = .rawMessage '& m_Tag & sReturn
                        RaiseEvent LookupResolved(cMsg)
                End Select
            End With
        Loop
        bLookupsWorking = False
    End If
    eventTimer.Enabled = False
End Sub
