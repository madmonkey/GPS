VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mutex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private hMutex As Long
Private bLocked As Boolean
Private m_Name As String
Private Type GUID
     Data1 As Long
     Data2 As Long
     Data3 As Long
     Data4(8) As Byte
End Type
Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As Any, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function OpenMutex Lib "kernel32" Alias "OpenMutexA" (ByVal fAccess As Long, ByVal fInherit As Long, ByVal sMutexName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Const INFINITE = &HFFFF

Public Function CreateGUID(Optional strRemoveChars As String = "{}-") As String
Dim udtGUID As GUID
Dim strGUID As String
Dim bytGUID() As Byte
Dim lngLen As Long, lngRetVal As Long, lngPos As Long

    lngLen = 40
    bytGUID = String(lngLen, 0)
    CoCreateGuid udtGUID 'Create the GUID
    lngRetVal = StringFromGUID2(udtGUID, VarPtr(bytGUID(0)), lngLen) 'Convert the structure into a displayable string
    strGUID = bytGUID
    If (Asc(Mid$(strGUID, lngRetVal, 1)) = 0) Then lngRetVal = lngRetVal - 1
    strGUID = Left$(strGUID, lngRetVal) 'Trim the trailing characters
    For lngPos = 1 To Len(strRemoveChars) 'Remove the unwanted characters
        strGUID = Replace(strGUID, Mid(strRemoveChars, lngPos, 1), vbNullString)
    Next
    CreateGUID = strGUID
    
End Function

Private Sub Class_Initialize()
    m_Name = CreateGUID("{}")
End Sub

Public Function RequestMutex(Optional ByVal Timeout As Long = INFINITE) As Boolean
Dim lReturn As Long
Const WAIT_ABANDONED = &H80
Const WAIT_FAILED = &HFFFFFFFF
Const WAIT_OBJECT_0 = &H0
Const WAIT_TIMEOUT = &H102
    
    If hMutex <> 0 And Not bLocked Then
        lReturn = WaitForSingleObject(hMutex, Timeout)
        RequestMutex = (lReturn = WAIT_OBJECT_0)
        bLocked = (lReturn = WAIT_OBJECT_0)
    Else
        RequestMutex = False
    End If

End Function

Public Function ConstructMutex(Optional nameMutex As String = vbNullString) As Boolean
Const NO_ERROR = 0&
Const ERROR_ALREADY_EXISTS = 183&
Const SECURITYDESCRIPTOR = &HF0000
Const SYNCHRONIZE = &H100000
Const MUTANT = &H1
Const MUTEX_ALL_ACCESS = SECURITYDESCRIPTOR Or SYNCHRONIZE Or MUTANT
Dim lastErr As Long
    
    If nameMutex <> vbNullString Then m_Name = nameMutex
    hMutex = CreateMutex(ByVal 0&, 1, m_Name)
    lastErr = Err.LastDllError
    
    Select Case lastErr
        Case NO_ERROR
            ConstructMutex = True 'Created with NO_ERROR
        Case ERROR_ALREADY_EXISTS
            hMutex = OpenMutex(MUTEX_ALL_ACCESS, False, m_Name)
            If hMutex <> 0 Then
                ConstructMutex = True 'Created only NOT here!
            Else
                ConstructMutex = False 'Problems opening Mutex! - Err.LastDllError
            End If
        Case Else
            ConstructMutex = False '"Unable to create Mutex!"
    End Select
    
End Function

Public Function DiscardMutex()
    
    If hMutex <> 0 Then ReleaseMutex hMutex: bLocked = False
        
End Function

Public Function DestroyMutex()
    
    If hMutex <> 0 Then CloseHandle hMutex
    
End Function

Private Sub Class_Terminate()
    DestroyMutex
End Sub



