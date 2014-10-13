Attribute VB_Name = "modXPTheme"
Option Explicit

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
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long

Private Function FileExists(ByVal strSearchFile As String) As Boolean
    Dim strucFname As OFSTRUCT
    FileExists = (OpenFile(strSearchFile, strucFname, OF_EXIST) <> -1)
End Function
Private Function AreThemesSupported() As Boolean
'Returns True if themes are supported
Dim hLib As Long
    hLib = LoadLibrary("uxtheme.dll")
    If hLib <> 0 Then FreeLibrary hLib
    AreThemesSupported = Not (hLib = 0)
End Function
Public Function InitializeForThemes() As Long
    InitializeForThemes = InitCommonControls
End Function
Public Function AreThemesActive() As Boolean
'Is theming currently active?
Dim hLib As Long, hProc As Long, hThemed As Long
    
    hThemed = 0
    hLib = LoadLibrary("uxtheme.dll")
    If hLib <> 0 Then
        hProc = GetProcAddress(hLib, "IsThemeActive")
        If hProc <> 0 Then
            hThemed = CallWindowProc(hProc, 0&, 0&, 0&, 0&)
        End If
        FreeLibrary hLib
    End If
    AreThemesActive = (hThemed <> 0)
    
End Function
Public Function EnsureManifest() As Boolean
'Returns 'true' if application needs to restart, to let changes take effect
Dim manifestFile As String
Dim fh As Long
Dim ret As Long
    manifestFile = App.Path & "\" & App.EXEName & ".exe.manifest"
    'If AreThemesSupported Then'next function tests for that by default
        If AreThemesActive Then
            If Not FileExists(manifestFile) Then
                fh = FreeFile
                Open manifestFile For Binary Access Write As #fh
                    Put fh, , getManifest
                Close #fh
                EnsureManifest = True
                RestartApp
            End If
        End If
    'End If
    
End Function
Public Sub RestartApp()
Dim parent As Long
Const SW_SHOWNORMAL = 1
    parent = GetDesktopWindow
    If IsWindow(parent) <> 0 Then
        ShellExecute parent, vbNullString, App.EXEName & ".exe", Command$, App.Path & IIf(Right$(App.Path, 1) = "\", vbNullString, "\"), SW_SHOWNORMAL
    End If
End Sub
Private Function getManifest() As String
    'XML Manifest required for XP Themes compliance
    getManifest = "<?xml version=" & Chr$(34) & "1.0" & Chr$(34) & " encoding=" & Chr$(34) & "UTF-8" & Chr$(34) & " standalone=" & Chr$(34) & "yes" & Chr$(34) & " ?> " & vbCrLf & _
                    "<assembly xmlns=" & Chr$(34) & "urn:schemas-microsoft-com:asm.v1" & Chr$(34) & " manifestVersion=" & Chr$(34) & "1.0" & Chr$(34) & "> " & vbCrLf & _
                    "<assemblyIdentity processorArchitecture=" & Chr$(34) & "*" & Chr$(34) & " version=" & Chr$(34) & "5.1.0.0" & Chr$(34) & " type=" & Chr$(34) & "win32" & Chr$(34) & vbCrLf & _
                    "name=" & Chr$(34) & "Microsoft.Windows.Shell.shell32" & Chr$(34) & "/> " & vbCrLf & _
                    "<description>Windows Shell</description> " & vbCrLf & _
                    "<dependency> " & vbCrLf & _
                    "<dependentAssembly> " & vbCrLf & _
                    "<assemblyIdentity type=" & Chr$(34) & "win32" & Chr$(34) & " name=" & Chr$(34) & "Microsoft.Windows.Common-Controls" & Chr$(34) & vbCrLf & _
                    "version=" & Chr$(34) & "6.0.0.0" & Chr$(34) & " publicKeyToken=" & Chr$(34) & "6595b64144ccf1df" & Chr$(34) & " language=" & Chr$(34) & "*" & Chr$(34) & vbCrLf & _
                    "processorArchitecture=" & Chr$(34) & "*" & Chr$(34) & " /> " & vbCrLf & _
                    "</dependentAssembly> " & vbCrLf & _
                    "</dependency> " & vbCrLf & _
                    "</assembly>"
End Function


