VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProcesses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Available Processes to select from
Private m_PrivateCollection As Collection

Private Sub Class_Initialize()
    ' explicit assignment is slightly faster than auto-instancing
    Set m_PrivateCollection = New Collection
    Initialize
End Sub

' Add a new cProcess item to the collection

Public Sub Add(newItem As cProcess, Optional Key As Variant)
Attribute Add.VB_Description = "Adds a member to a Collection object"
    m_PrivateCollection.Add newItem, IIf(IsMissing(Key), newItem.progID, Key)
End Sub
Private Sub Initialize()
Dim sPath As String, sFile As String, sName As String, sProgID As String
Dim myArray As Variant
Dim availProcess As cProcess
    
    If InDevelopment Then
        myArray = Split(App.Path, "\")
        ReDim Preserve myArray(UBound(myArray) - 1)
        sPath = Join(myArray, "\")
        sPath = sPath & "\HTE_ClientUtilities\Processes\"
    Else
        sPath = App.Path & "\Processes\"
    End If
    sFile = UCase$(Dir(sPath))
    While sFile <> vbNullString
        Set availProcess = New cProcess
        With availProcess
            .fileName = sFile
            .compiledType = UCase$(Right$(sFile, 4))
            Select Case .compiledType
                Case ".DLL", ".EXE"
                    If VerifyProcess(.fileName, .compiledType, sName, sProgID) Then
                        'ADD
                        .friendlyName = sName: .progID = sProgID
                        Add availProcess
                    End If
            End Select
            sFile = UCase$(Dir)
        End With
    Wend
        
End Sub

Private Function VerifyProcess(ByVal sFile As String, ByVal sExtn As String, Optional ByRef friendlyName As String, Optional ByRef progID As String) As Boolean
Dim oTest As HTE_GPS.Process
Dim oObj As Object
Dim sObjectName As String

On Local Error GoTo err_VerifyProcess
    sObjectName = Replace(sFile, sExtn, ".PROCESS", , , vbTextCompare)
    Set oObj = CreateObject(sObjectName)
    Set oTest = oObj
    friendlyName = oTest.friendlyName
    progID = sObjectName
    VerifyProcess = True
    Exit Function
err_VerifyProcess:
    Err.Clear
    VerifyProcess = False
End Function
' Remove an item from the collection

Public Sub Remove(index As Variant)
Attribute Remove.VB_Description = "Removes a member from a Collection object"
    m_PrivateCollection.Remove index
End Sub

' Return a cProcess item from the collection

Function Item(index As Variant) As cProcess
Attribute Item.VB_Description = "Returns a specific member of a Collection "
Attribute Item.VB_UserMemId = 0
    Set Item = m_PrivateCollection.Item(index)
End Function

' Return the number of items in the collection

Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of members in a collection"
    Count = m_PrivateCollection.Count
End Property

' Remove all items from the collection

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all members from a Collection object"
    Set m_PrivateCollection = New Collection
End Sub

' Implement support for enumeration (For Each)

Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    ' delegate to the private collection
    Set NewEnum = m_PrivateCollection.[_NewEnum]
End Function
