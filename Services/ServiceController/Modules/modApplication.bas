Attribute VB_Name = "modApplication"
Option Explicit

Dim fMain As frmServiceControl 'define your main application form here!
'GPF PROTECTION - ERROR ON CLOSE O/S
#If GPFProtect = 1 Then
    Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
    Private Const SEM_FAILCRITICALERRORS = &H1
    Private Const SEM_NOGPFAULTERRORBOX = &H2
    Private Const SEM_NOOPENFILEERRORBOX = &H8000&
#End If
'PROCESS NEEDS TO BE KILLED ON SHUTDOWN
#If AggressiveShutDown = 1 Then
    Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
    Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
    Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
#End If
'XP INITIALIZATION
#If UseThemes = 1 Then
    Private Type tagInitCommonControlsEx
       lngSize As Long
       lngICC As Long
    End Type
    Private Declare Function InitCommonControlsEx Lib "Comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
    Private Const ICC_USEREX_CLASSES = &H200
#End If
'SINGLETON APPLICATION
#If Singleton = 1 Then
    Private Const m_THISAPPID = "{56155A57-4D9E-4727-8294-2E9C1D2C36A4}"
    Private Const SW_SHOW As Long = 5&
    Private Const SW_RESTORE As Long = 9&
    Private APP_MUTEX As Mutex
    Private m_hWndPrevious As Long
    Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
    Private Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd&, lpdwProcessId&) As Long
    Private Declare Function ShowWindow Lib "user32" (ByVal hWnd&, ByVal nCmdShow&) As Long
    Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach&, ByVal idAttachTo&, ByVal fAttach&) As Long
    Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd&) As Long
    Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
    Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
    Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If
'DEVELOPMENT MODE
Private m_bInDevelopment As Boolean

Private Sub Main()
    If SingletonApp Then
        OLEApplicationPatch
        InitializeForThemes
        LoadMain
    End If
End Sub

Public Sub EndApp()
    ProtectErrorOnClose
    MutexCleanup
    Suicide
End Sub

Public Function InDevelopment() As Boolean
   ' Debug.Assert code not run in an EXE.  Therefore m_bInDevelopment variable is never set.
   Debug.Assert InDevelopmentHack() = True
   InDevelopment = m_bInDevelopment
End Function

Private Function InDevelopmentHack() As Boolean
   m_bInDevelopment = True
   InDevelopmentHack = m_bInDevelopment
End Function

#If Singleton = 1 Then
    Private Function WeAreAlone(ByVal sMutex As String) As Boolean
        ' Don't call Mutex when in VBIDE because it will apply for the entire VB IDE session, not just the app's session.
        If InDevelopment Then
            WeAreAlone = Not (App.PrevInstance)
        Else
            ' Ensures we don't run a second instance even if the first instance is in the start-up phase
            If APP_MUTEX Is Nothing Then Set APP_MUTEX = New Mutex
            If APP_MUTEX.ConstructMutex(sMutex) Then
                WeAreAlone = APP_MUTEX.RequestMutex
            End If
       End If
    End Function
    
    Private Function FindExistingApplicationInstance(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim bStop As Boolean
       ' Customised windows enumeration procedure.  Stops when it finds another application with the Window property set, or when all windows are exhausted.
       bStop = False
       If IsThisApp(hWnd) Then
          FindExistingApplicationInstance = 0
          m_hWndPrevious = hWnd
       Else
          FindExistingApplicationInstance = 1
       End If
    End Function
    
    Private Function EnumerateWindows() As Boolean
       ' Enumerate top-level windows:
       EnumWindows AddressOf FindExistingApplicationInstance, 0
    End Function
    
    Private Function IsThisApp(ByVal hWnd As Long) As Boolean
        IsThisApp = GetProp(hWnd, m_THISAPPID & "_APPLICATION") = 1
    End Function
    
    Public Function RestoreAndActivate(ByVal hWnd As Long) As Boolean
    Dim fShowWindowFlag&, hWndForeground&, nCurThreadID&, nNextThreadID&
    Const API_FALSE As Long = 0&
    Const API_TRUE As Long = 1&
        
        If IsWindowEnabled(hWnd) <> 0 Then
            If IsWindowVisible(hWnd) <> 0 Then
                ' tell the previous instance to restore itself incase it's minimized
                fShowWindowFlag = IIf(IsIconic(hWnd), SW_RESTORE, SW_SHOW)
                Call ShowWindow(hWnd, fShowWindowFlag)
                 hWndForeground = GetForegroundWindow()
                 ' see if the window is already the foreground window.
                 If hWnd <> hWndForeground Then
                     ' if not, we need to get the thread IDs for this app and the current foreground window
                     nCurThreadID = GetWindowThreadProcessId(hWndForeground, ByVal 0&)
                     nNextThreadID = GetWindowThreadProcessId(hWnd, ByVal 0&)
                     ' the active window.
                     If nCurThreadID <> nNextThreadID Then
                         ' if the thread IDs are not the same, attach to the foreground
                         ' thread long enough to make our window the foreground window
                         Call AttachThreadInput(nCurThreadID, nNextThreadID, API_TRUE)
                         Call BringWindowToTop(hWnd)
                         Call SetForegroundWindow(hWnd)
                         Call AttachThreadInput(nCurThreadID, nNextThreadID, API_FALSE)
                     Else
                         Call BringWindowToTop(hWnd)
                         Call SetForegroundWindow(hWnd)
                     End If
                End If
            End If
       End If
    End Function
    
    Private Sub TagWindow(ByVal hWnd As Long)
       ' Applies a window property to allow the window to be clearly identified.
       SetProp hWnd, m_THISAPPID & "_APPLICATION", 1
       SetProp hWnd, "ProcessID", GetCurrentProcessId
    End Sub

#End If

Private Sub InitializeForThemes()
    #If UseThemes Then
        Dim iccex As tagInitCommonControlsEx
            With iccex
               .lngSize = LenB(iccex)
               .lngICC = ICC_USEREX_CLASSES
           End With
           InitCommonControlsEx iccex
    #End If
End Sub

Private Function SingletonApp() As Boolean
    'returns TRUE if can continue, FALSE if NOT
    #If Singleton = 1 Then
        If Not (WeAreAlone(m_THISAPPID & "_APPLICATION_MUTEX")) Then
            EnumerateWindows
            If (m_hWndPrevious <> 0) Then
                RestoreAndActivate (m_hWndPrevious)
            End If
            SingletonApp = False
        Else
            SingletonApp = True
        End If
    #Else
        SingletonApp = True
    #End If
End Function
Private Sub OLEApplicationPatch()
    #If UseOLESettings Then
        With App
            .OleRequestPendingMsgText = App.FileDescription & " is still attempting to process your request." 'Returns/sets text of 'busy' message displayed while an Automation request is pending.
            .OleRequestPendingMsgTitle = App.FileDescription & " OLE Request Pending" 'Returns/sets title of 'busy' message displayed while an Automation request is pending.
            .OleRequestPendingTimeout = 3600000 'Returns/sets milliseconds Automation requests will run before user actions trigger a 'busy' message.
            .OleServerBusyMsgText = App.FileDescription & " is still attempting to process your request." 'Returns/sets text of 'busy' message displayed if an ActiveX component rejects a request.
            .OleServerBusyMsgTitle = App.FileDescription & " OLE Server Busy" 'Returns/sets title of 'busy' message displayed when an ActiveX component rejects a request.
            .OleServerBusyRaiseError = True 'Determines whether a rejected Automation request raises an error, instead of displaying a 'busy' message.
            .OleServerBusyTimeout = 3600000  'Returns/sets milliseconds during which an Automation request will continue to be retried.
            .OleServerBusyRaiseError = True
            End With
    #End If
End Sub


Private Sub LoadMain()
    Set fMain = New frmServiceControl
    Load fMain
    #If Singleton = 1 Then
        TagWindow fMain.hWnd
    #End If
    fMain.Show
    fMain.SetFocus
End Sub

Private Sub ProtectErrorOnClose()
    #If GPFProtect = 1 Then
        If Not InDevelopment Then
            SetErrorMode SEM_NOGPFAULTERRORBOX
        End If
    #End If
End Sub

Private Sub MutexCleanup()
    #If Singleton = 1 Then
        If Not APP_MUTEX Is Nothing Then
            APP_MUTEX.DiscardMutex
            APP_MUTEX.DestroyMutex
            Set APP_MUTEX = Nothing
        End If
    #End If
End Sub
Private Sub Suicide()
    #If AggressiveShutDown = 1 Then
        If Not InDevelopment Then
            ExitProcess GetExitCodeProcess(GetCurrentProcess, 0)
        End If
    #End If
End Sub
