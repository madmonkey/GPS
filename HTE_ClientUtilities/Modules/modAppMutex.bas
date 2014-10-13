Attribute VB_Name = "modAppMutex"
Option Explicit
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Const ERROR_ALREADY_EXISTS = 183&
Private m_hMutex As Long
Private Function WeAreAlone(ByVal sMutex As String) As Boolean
   ' Don't call Mutex when in VBIDE because it will apply for the entire VB IDE session, not just the app's session.
   If InDevelopment Then
      WeAreAlone = Not (App.PrevInstance)
   Else
      ' Ensures we don't run a second instance even if the first instance is in the start-up phase
      m_hMutex = CreateMutex(ByVal 0&, 1, sMutex)
      If (Err.LastDllError = ERROR_ALREADY_EXISTS) Then
         CloseHandle m_hMutex
      Else
         WeAreAlone = True
      End If
   End If
End Function
Public Function SoleInstance() As Boolean
    SoleInstance = WeAreAlone(App.EXEName & "_APPLICATION_MUTEX")
End Function
Public Function EndApp()
    If (m_hMutex <> 0) Then CloseHandle m_hMutex
    m_hMutex = 0
End Function
