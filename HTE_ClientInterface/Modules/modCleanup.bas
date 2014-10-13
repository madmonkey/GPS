Attribute VB_Name = "modCleanup"
Option Explicit

Const MAX_PATH& = 260
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Const PROCESS_ALL_ACCESS = &H1F0FFF

Private Type LUID
   lowpart As Long
   highpart As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    LuidUDT As LUID
    Attributes As Long
End Type

Private Type PROCESSENTRY32
  dwSize As Long
  cntUsage As Long
  th32ProcessID As Long
  th32DefaultHeapID As Long
  th32ModuleID As Long
  cntThreads As Long
  th32ParentProcessID As Long
  pcPriClassBase As Long
  dwFlags As Long
  szexeFile As String * MAX_PATH
End Type

Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Any) As Long

Public Function KillApp(myName As String) As Boolean
    'THE NEED FOR THIS ROUTINE IS BEYOND REPREHENSIBLE - WHEN MCS_CARRIER IS INVOLVED IN A PROCESS
    'ROUTE - THE ACTIVEX EXE IS SOMETIMES RETAINED IN MEMORY - MAGICALLY, PROBABLY DUE TO AN INADEQUATE
    'CLEAN-UP ON MCS' PART. WHAT WE ARE IN FACT DOING IS FORCEFULLY REMOVING THE ACTIVEX EXE FROM MEMORY
    'WHEN WE CLEAN-UP (OURS NOT MCS) TO ATTEMPT TO CIRCUMVENT PROBLEMS THAT WILL ARISE BECAUSE OF THIS
    
    Const TH32CS_SNAPPROCESS As Long = 2&
    Const PROCESS_ALL_ACCESS = 0
   
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long, hSnapshot As Long, exitCode As Long, myProcess As Long
    Dim szExename As String
    Dim AppKill As Boolean
    Dim appCount As Integer, i As Integer
   
On Local Error GoTo exit_KillApp
    appCount = 0
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    Do While rProcessFound
        i = InStr(1, uProcess.szexeFile, Chr(0))
        szExename = LCase$(Left$(uProcess.szexeFile, i - 1))
        'If Right$(szExename, Len(myName)) = LCase$(myName) Then
        If InStr(1, LCase$(myName), Right$(szExename, Len(myName)), vbTextCompare) = 1 Then
            KillApp = True
            appCount = appCount + 1
            myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
            If KillProcess(uProcess.th32ProcessID, 0) Then
                'MsgBox "Instance no. " & appCount & " of " & szExename & " was terminated!"
            End If
        End If
        rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop
    Call CloseHandle(hSnapshot)
    Exit Function
exit_KillApp:
    
End Function

Private Function KillProcess(ByVal hProcessID As Long, Optional ByVal exitCode As Long) As Boolean

'Terminate any application and return an exit code to Windows.
Dim hToken As Long, hProcess As Long
Dim tp As TOKEN_PRIVILEGES
    
    If GetVersion() >= 0 Then
        If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) = 0 Then
            GoTo CleanUp
        End If
        If LookupPrivilegeValue("", "SeDebugPrivilege", tp.LuidUDT) = 0 Then
            GoTo CleanUp
        End If
        tp.PrivilegeCount = 1
        tp.Attributes = SE_PRIVILEGE_ENABLED
        If AdjustTokenPrivileges(hToken, False, tp, 0, ByVal 0&, ByVal 0&) = 0 Then
            GoTo CleanUp
        End If
    End If
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, hProcessID)
    If hProcess Then
        KillProcess = (TerminateProcess(hProcess, exitCode) <> 0)
        ' close the process handle
        CloseHandle hProcess
    End If
    
    If GetVersion() >= 0 Then
        ' under NT restore original privileges
        tp.Attributes = 0
        AdjustTokenPrivileges hToken, False, tp, 0, ByVal 0&, ByVal 0&

CleanUp:
        If hToken Then CloseHandle hToken
    End If
    
End Function
