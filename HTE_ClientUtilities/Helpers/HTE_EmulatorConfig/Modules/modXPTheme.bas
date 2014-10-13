Attribute VB_Name = "modXPTheme"
Option Explicit

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long

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
