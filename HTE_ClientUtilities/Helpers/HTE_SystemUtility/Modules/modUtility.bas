Attribute VB_Name = "modUtility"
Option Explicit

Private Const cModuleName = "modUtility"

Public Declare Function DeskSetCurrentScheme Lib "desk.cpl" (ByVal SchemeName As String) As Long
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformID As Long
   szCSDVersion As String * 128
End Type

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const GW_CHILD = 5
Private Const CB_FINDSTRING As Long = &H14C
Private Const CB_SETCURSEL As Long = &H14E
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_CLOSE = &H10
Private Const SW_HIDE As Long = 0
Private Const HWND_BOTTOM As Long = 1
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOSIZE As Long = &H1

Private Type DLLVersionInfo
    cbSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
End Type

Private Declare Function DllGetVersion Lib "Shlwapi.dll" (dwVersion As DLLVersionInfo) As Long

Private Type STARTUPINFO
  cb As Long
  lpReserved As String
  lpDesktop As String
  lpTitle As String
  dwX As Long
  dwY As Long
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwFlags As Long
  wShowWindow As Long
  cbReserved2 As Long
  lpReserved2 As Long
  hStdInput As Long
  hStdOutput As Long
  hStdError As Long
End Type

Private Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessId As Long
  dwThreadID As Long
End Type
Private Enum PROCESS_PRIORITY
    NORMAL_PRIORITY_CLASS = &H20&
    IDLE_PRIORITY_CLASS = &H40&
    HIGH_PRIORITY_CLASS = &H80&
End Enum
Private Const WAIT_INFINITE = -1&
Private Const WAIT_FAILED = &HFFFFFFFF
Private Const WAIT_ABANDONED = &H80
Private Const WAIT_TIMEOUT = 258&
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const STATUS_PENDING = &H103&
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpAppName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Function GetIEVersion() As String
    Dim VersionInfo As DLLVersionInfo
    
    VersionInfo.cbSize = Len(VersionInfo)
    Call DllGetVersion(VersionInfo)
    '"Internet Explorer " & VersionInfo.dwMajorVersion & "." & VersionInfo.dwMinorVersion & "." & VersionInfo.dwBuildNumber
    GetIEVersion = CStr(VersionInfo.dwMajorVersion & "." & VersionInfo.dwMinorVersion)
End Function
Public Function IsWinNT() As Boolean

  'Returns True if the current operating system is WinNT
   Dim osvi As OSVERSIONINFO
   osvi.dwOSVersionInfoSize = Len(osvi)
   GetVersionEx osvi
   IsWinNT = (osvi.dwPlatformID = VER_PLATFORM_WIN32_NT)
   
End Function
Public Function GetWinVersion(ByRef WIN As RGB_WINVER) As String
'returns a structure (RGB_WINVER) filled with OS information
   Dim OSV As OSVERSIONINFO
   Dim pos As Integer
   Dim sVer As String
   Dim sBuild As String

    OSV.dwOSVersionInfoSize = Len(OSV)
    If GetVersionEx(OSV) = 1 Then
        'PlatformId contains a value representing the OS
        WIN.PlatformID = OSV.dwPlatformID
        Select Case OSV.dwPlatformID
            Case VER_PLATFORM_WIN32s:   WIN.VersionName = "Win32s"
            Case VER_PLATFORM_WIN32_NT: WIN.VersionName = "Windows NT"
                Select Case OSV.dwMajorVersion
                    Case 4:  WIN.VersionName = "Windows NT"
                    Case 5:
                        Select Case OSV.dwMinorVersion
                            Case 0:  WIN.VersionName = "Windows 2000"
                            Case 1:  WIN.VersionName = "Windows XP"
                        End Select
                End Select
            Case VER_PLATFORM_WIN32_WINDOWS:
            'The dwVerMinor bit tells if its 95 or 98.
                Select Case OSV.dwMinorVersion
                    Case 0:    WIN.VersionName = "Windows 95"
                    Case 90:   WIN.VersionName = "Windows ME"
                    Case Else: WIN.VersionName = "Windows 98"
                End Select
        End Select
        'Get the version number
        WIN.VersionNo = OSV.dwMajorVersion & "." & OSV.dwMinorVersion
        'Get the build
        WIN.BuildNo = (OSV.dwBuildNumber And &HFFFF&)
        'Any additional info. In Win9x, this can be "any arbitrary string" provided by the manufacturer. In NT, this is the service pack.
        pos = InStr(OSV.szCSDVersion, Chr$(0))
        If pos Then
            WIN.ServicePack = Left$(OSV.szCSDVersion, pos - 1)
        End If
        GetWinVersion = WIN.VersionName & ":" & WIN.VersionNo & " " & WIN.ServicePack & " - Build " & WIN.BuildNo
    End If
End Function

Public Function IsVersionAtLeast(ByVal VerMajor As Long, ByVal VerMinor As Long, Optional ByVal BuildNumber As Long = 0) As Boolean
Dim OSV As OSVERSIONINFO
    OSV.dwOSVersionInfoSize = Len(OSV)
    If GetVersionEx(OSV) = 1 Then
        If OSV.dwPlatformID >= VER_PLATFORM_WIN32_WINDOWS Then
            Select Case True
                Case OSV.dwMajorVersion = VerMajor
                    IsVersionAtLeast = OSV.dwMinorVersion >= VerMinor And _
                        IIf(BuildNumber <> 0, OSV.dwBuildNumber >= BuildNumber, True)
                Case OSV.dwMajorVersion > VerMajor
                    IsVersionAtLeast = True
                Case Else
                    IsVersionAtLeast = False
            End Select
        Else
            IsVersionAtLeast = False
        End If
    End If
End Function
Public Function IsWin95() As Boolean
Dim OSV As OSVERSIONINFO
      OSV.dwOSVersionInfoSize = Len(OSV)
      If GetVersionEx(OSV) = 1 Then
         IsWin95 = (OSV.dwPlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
                (OSV.dwMajorVersion = 4 And OSV.dwMinorVersion = 0)
      End If
End Function
Public Function IsWin98() As Boolean
Dim OSV As OSVERSIONINFO
      OSV.dwOSVersionInfoSize = Len(OSV)
      If GetVersionEx(OSV) = 1 Then
         IsWin98 = (OSV.dwPlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
                   (OSV.dwMajorVersion > 4) Or _
                   (OSV.dwMajorVersion = 4 And OSV.dwMinorVersion > 0)
      End If
End Function
Public Function IsWinME() As Boolean
Dim OSV As OSVERSIONINFO
      OSV.dwOSVersionInfoSize = Len(OSV)
      If GetVersionEx(OSV) = 1 Then
         IsWinME = (OSV.dwPlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
                   (OSV.dwMajorVersion = 4 And OSV.dwMinorVersion = 90)
      End If
End Function
Public Function IsWinNT4() As Boolean
Dim OSV As OSVERSIONINFO
      OSV.dwOSVersionInfoSize = Len(OSV)
      If GetVersionEx(OSV) = 1 Then
        'PlatformId contains a value representing the OS.
        'If VER_PLATFORM_WIN32_NT and dwVerMajor is 4, return true
         IsWinNT4 = (OSV.dwPlatformID = VER_PLATFORM_WIN32_NT) And _
                    (OSV.dwMajorVersion = 4)
      End If
End Function
Public Function IsWin2000() As Boolean
'returns True if running Windows 2000 (NT5)
Dim OSV As OSVERSIONINFO
      OSV.dwOSVersionInfoSize = Len(OSV)
      If GetVersionEx(OSV) = 1 Then
         IsWin2000 = (OSV.dwPlatformID = VER_PLATFORM_WIN32_NT) And _
                     (OSV.dwMajorVersion = 5 And OSV.dwMinorVersion = 0)
      End If
End Function
Public Function IsWinXP() As Boolean
'returns True if running WinXP (NT5.1)
Dim OSV As OSVERSIONINFO
      OSV.dwOSVersionInfoSize = Len(OSV)
      If GetVersionEx(OSV) = 1 Then
         IsWinXP = (OSV.dwPlatformID = VER_PLATFORM_WIN32_NT) And _
                   (OSV.dwMajorVersion = 5 And OSV.dwMinorVersion = 1)
      End If
End Function
Public Function GetWinVer() As String
'returns a string representing the version,
Dim OSV As OSVERSIONINFO
Dim r As Long
Dim pos As Integer
Dim sVer As String
Dim sBuild As String
   
    OSV.dwOSVersionInfoSize = Len(OSV)
    If GetVersionEx(OSV) = 1 Then
        Select Case OSV.dwPlatformID     'PlatformId contains a value representing the OS
            Case VER_PLATFORM_WIN32s: GetWinVer = "32s"
            Case VER_PLATFORM_WIN32_NT:
            'dwVerMajor = NT version.
            'dwVerMinor = minor version
                Select Case OSV.dwMajorVersion
                    Case 3:
                        Select Case OSV.dwMinorVersion
                            Case 0:  GetWinVer = "NT3"
                            Case 1:  GetWinVer = "NT3.1"
                            Case 5:  GetWinVer = "NT3.5"
                            Case 51: GetWinVer = "NT3.51"
                        End Select
                    Case 4: GetWinVer = "NT 4"
                    Case 5:
                        Select Case OSV.dwMinorVersion
                            Case 0:  GetWinVer = "Win2000"
                            Case 1:  GetWinVer = "WinXP"
                        End Select
                    End Select
            Case VER_PLATFORM_WIN32_WINDOWS:
               'dwVerMinor bit tells if its 95 or 98.
                Select Case OSV.dwMinorVersion
                    Case 0:    GetWinVer = "95"
                    Case 90:   GetWinVer = "ME"
                    Case Else: GetWinVer = "98"
                End Select
          End Select
   End If
End Function
Public Function RunDLL(ByVal Property As String, Optional ByVal Pages As Long) As Boolean
    Dim cmd As String
    Dim r As Long, hProcess As Long, hWnd As Long
    Dim CPLSuccess As Boolean
    Dim cplRun As String, cplTitle As String, cplClassName As String
    Dim CPLCMD As Variant
    
    
    'TODO: I hate to do this BUT someone needs to verify the functionality for different platforms
    'and adjust accordingly. It's not going to be me today but probably will at some point in the future.
    'Hey dude how's it been - glad to see you are looking at this!
    UEH_Log cModuleName, "RunDLL", "Property = " & Property & " Pages = " & CStr(Pages), logVerbose
    CPLSuccess = True

    Select Case LCase$(Property)
        Case "display"
            cplRun = "desk.cpl": cplTitle = "Display Properties": cplClassName = "#32770"
        Case "addremove"
            cplRun = "appwiz.cpl": cplTitle = "Add or Remove Programs": cplClassName = "NativeHWNDHost"
        Case "internet"
            cplRun = "inetcpl.cpl": cplTitle = "Internet Properties": cplClassName = "#32770"
        Case "regional"
            cplRun = "intl.cpl": cplTitle = "Regional and Language Options": cplClassName = "#32770"
        Case "joystick"
            cplRun = "joy.cpl": cplTitle = "Game Controllers": cplClassName = "#32770"
        Case "mouse"
            cplRun = "main.cpl": cplTitle = "Mouse Properties": cplClassName = "#32770"
        Case "multimedia"
            'TODO: verify class and title
            cplRun = "mmsys.cpl": cplTitle = "Sounds and Audio Devices": cplClassName = "#32770"
        Case "network"
            'TODO: verify class and title
            cplRun = "netcpl.cpl": cplTitle = "Network Connections": cplClassName = "CabinetWClass"
        Case Is = "password"
            '???WIN9x only???
            cplRun = "password.cpl": cplTitle = ""
        Case "power"
            cplRun = "powercfg.cpl": cplTitle = "Power Options Properties": cplClassName = "#32770"
        Case "scanner"
            'TODO: verify class and title
            cplRun = "sticpl.cpl": cplTitle = ""
        Case "system"
            cplRun = "sysdm.cpl": cplTitle = "System Properties": cplClassName = "#32770"
        Case "timedate"
            cplRun = "timedate.cpl": cplTitle = "Date and Time Properties": cplClassName = "#32770"
        Case "access"
            cplRun = "access.cpl": cplTitle = "Accessibility Options": cplClassName = "#32770"
        Case "phone"
            cplRun = "telephon.cpl": cplTitle = "Phone and Modem Options": cplClassName = "#32770"
        Case "odbc"
            cplRun = "odbccp32.cpl": cplTitle = "ODBC Data Source Administrator": cplClassName = "#32770"
        Case "themes"
            If IsWin98 Then
                cplRun = "themes.cpl": cplTitle = "Display Properties": cplClassName = "#32770"
            Else
                cplRun = "desk.cpl": cplTitle = "Display Properties": cplClassName = "#32770"
            End If
        Case Else
            CPLSuccess = False
            GoTo CPLFail
        End Select
CPLProceed:
        If Not IsMissing(Pages) Then
            'THIS MAY HAVE TO BE LOOKED AT IN MORE DETAIL, SINCE THE ACTUAL CMDLINE MAY BE DIFFERENT
            'DEPENDING ON WHAT YOU ARE WANTING TO DO
            'CHECK http://herman.eldering.net/vb/cntrlpnl.htm for more details....
            CPLCMD = "rundll32.exe shell32.dll,Control_RunDLL " & cplRun & ",," & CStr(Pages)
        Else
            CPLCMD = "rundll32.exe shell32.dll,Control_RunDLL " & cplRun
        End If
        UEH_Log cModuleName, "RunDLL", "CPLCMD= " & CPLCMD, logVerbose
        r = Shell(CPLCMD, 1)
        RunDLL = (r <> 0)
        If r <> 0 Then
            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, r)
            DoEvents
            hWnd = FindSpecializedWindow(cplClassName, cplTitle, 100)
            EnsurePosition hWnd
            RunDLL = True
        Else
            RunDLL = False
        End If
        Exit Function
CPLFail:
        UEH_Log cModuleName, "RunDLL", "CPLFail", logError
        RunDLL = CPLSuccess
    End Function
Public Function HorribleWin9xHack(ByVal psScheme As String) As Boolean
Dim hWnd As Long 'MAIN WINDOW
Dim bWnd As Long 'CHILD WINDOW
Dim chWnd As Long 'ACTIVE CONTROL
Dim CPLCMD As String
Dim index As Long
Dim bReturn As Boolean
Dim PROC As PROCESS_INFORMATION
Dim ProcessID As Long
Dim hProcess As Long
Dim exitCode As Long
    UEH_Log cModuleName, "HorribleWin9xHack", "psScheme = " & psScheme, logVerbose
    bReturn = False
    psScheme = psScheme & Chr$(0)
    CPLCMD = "rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,2"
    ProcessID = Shell(CPLCMD)
    If ProcessID <> 0 Then 'CreateWinProcess(CPLCMD, PROC, HIGH_PRIORITY_CLASS) Then
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, ProcessID)
        DoEvents
        UEH_Log cModuleName, "HorribleWin9xHack", "Created system process ID = " & ProcessID & "; handle = " & hProcess, logVerbose
        DoEvents
        hWnd = FindSpecializedWindow("#32770", "Display Properties", 100)
        UEH_Log cModuleName, "HorribleWin9xHack", "hwnd = " & CStr(hWnd), logVerbose
        If hWnd <> 0 Then
            DoEvents
            ShowWindow hWnd, SW_HIDE
            UEH_Log cModuleName, "HorribleWin9xHack", "ShowWindow SW_HIDE", logVerbose
            SetWindowPos hWnd, HWND_BOTTOM, Screen.Height, Screen.Width, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE
            bWnd = GetWindow(hWnd, GW_CHILD)
            UEH_Log cModuleName, "HorribleWin9xHack", "bWnd =  " & CStr(bWnd), logVerbose
            If bWnd <> 0 Then
                chWnd = GetDlgItem(bWnd, &H578) 'IDENTIFIER OF THE SCHEME COMBO-BOX
                UEH_Log cModuleName, "HorribleWin9xHack", "COMBO chWnd =  " & CStr(chWnd), logVerbose
                If chWnd <> 0 Then
                    index = SendMessage(chWnd, CB_FINDSTRING, -1, ByVal psScheme)
                    UEH_Log cModuleName, "HorribleWin9xHack", "CB_FINDSTRING INDEX =  " & CStr(index), logVerbose
                    If index <> -1 Then
                        PostMessage chWnd, CB_SETCURSEL, index, 0
                        PostMessage chWnd, WM_LBUTTONDOWN, 0, ByVal 0&
                        PostMessage chWnd, WM_LBUTTONUP, 0, ByVal 0&
                        UEH_Log cModuleName, "HorribleWin9xHack", "COMBO Clicked", logVerbose
                        'GRAB OK BUTTON
                        chWnd = GetDlgItem(hWnd, &H1)
                        UEH_Log cModuleName, "HorribleWin9xHack", "OK chWnd =  " & CStr(chWnd), logVerbose
                        If chWnd <> 0 Then
                            PostMessage chWnd, WM_LBUTTONDOWN, 0, ByVal 0&
                            PostMessage chWnd, WM_LBUTTONUP, 0, ByVal 0&
                            UEH_Log cModuleName, "HorribleWin9xHack", "OK Clicked", logVerbose
                            bReturn = True
                            Do
                                Call GetExitCodeProcess(hProcess, exitCode)
                                DoEvents: DoEvents
                            Loop While exitCode = STATUS_PENDING
                            Call CloseHandle(hProcess)
'''                            Debug.Print WaitForSingleObject(PROC.hProcess, TIMEOUT)
'''                            bReturn = WaitForSingleObject(PROC.hProcess, TIMEOUT) = 0&
                        Else
                            UEH_Log cModuleName, "HorribleWin9xHack", "Unable to FindDefaultControl", logWarning
                        End If
                    Else
                        UEH_Log cModuleName, "HorribleWin9xHack", "Unable to FindString Composite", logWarning
                    End If
                Else
                    UEH_Log cModuleName, "HorribleWin9xHack", "Unable to GetDlgItem", logWarning
                End If
            Else
                UEH_Log cModuleName, "HorribleWin9xHack", "Unable to FindChildWindow", logWarning
            End If
        Else
            UEH_Log cModuleName, "HorribleWin9xHack", "Unable to FindWindow", logWarning
        End If
'''        UEH_Log cModuleName, "HorribleWin9xHack", "Clean up process thread handles (hProcess - " & PROC.hProcess & ", hThread - " & PROC.hThread & ")", logVerbose
'''        DestroyWinProcess PROC
    Else
        UEH_Log cModuleName, "HorribleWin9xHack", "Unable to spawn process", logWarning
    End If
    DoEvents
    UEH_Log cModuleName, "HorribleWin9xHack", "bReturn = " & CStr(bReturn), logVerbose
'''    If IsWindow(hWnd) <> 0 Then
'''        'It's possible that something bad happened, close-up
'''        UEH_Log cModuleName, "HorribleWin9xHack", "Window still open", logWarning
'''        PostMessage hWnd, WM_CLOSE, 0&, 0&
'''    End If
    HorribleWin9xHack = bReturn
    
End Function

Private Function WinSysDir() As String
Dim ls_Return As String, ll_ret As Long
    ls_Return = Space(255)
    ll_ret = GetSystemDirectory(ls_Return, 255)
    ls_Return = Left$(ls_Return, ll_ret)
    WinSysDir = ls_Return
End Function

Private Function CreateWinProcess(ByVal cmdline As String, ByRef PROC As PROCESS_INFORMATION, Optional ByVal Priority As PROCESS_PRIORITY = NORMAL_PRIORITY_CLASS) As Boolean

Dim start As STARTUPINFO
  
  'Initialize the STARTUPINFO structure by passing to start the size of the STARTUPINFO
  'type. Setting the .cb member is the only item of the structure needed to launch the program
   start.cb = Len(start)
  'Start the application
   CreateWinProcess = CreateProcess(0&, cmdline, 0&, 0&, 1&, Priority, 0&, 0&, start, PROC)
   'ret = WaitForSingleObject(proc.hProcess, Timeout) 'Wait for the application to finish
   
End Function

Private Function DestroyWinProcess(ByRef PROC As PROCESS_INFORMATION)
    Call CloseHandle(PROC.hProcess) 'Close the handle to the process
    Call CloseHandle(PROC.hThread) 'Close the handle to the thread created
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

