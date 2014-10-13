Attribute VB_Name = "modMaintainService"
Option Explicit
Const MAX_PATH& = 260

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
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Function FindApp(myName As String) As Boolean
    Const TH32CS_SNAPPROCESS As Long = 2&
    Const PROCESS_ALL_ACCESS = 0
   
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long, hSnapshot As Long
    Dim szExename As String
    Dim i As Integer
   
On Local Error GoTo exit_FindApp
    FindApp = False
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    Do While rProcessFound
        i = InStr(1, uProcess.szexeFile, Chr(0))
        szExename = LCase$(Left$(uProcess.szexeFile, i - 1))
        If InStr(1, LCase$(myName), Right$(szExename, Len(myName)), vbTextCompare) = 1 Then
            FindApp = True
            Exit Do
        End If
        rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop
    Call CloseHandle(hSnapshot)
    Exit Function
exit_FindApp:
    
End Function
