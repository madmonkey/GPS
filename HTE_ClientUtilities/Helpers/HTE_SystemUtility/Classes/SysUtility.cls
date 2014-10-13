VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SysUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum cpApplets
    Display = 0
    Accessibility = 1
    Date_Time = 2
    Mouse = 3
    System = 4
    Modem = 5
    Sounds = 6
    Add_Remove = 7
    Regional_Options = 8 'THE CORRECT SPELLING MAY CAUSE PROBLEMS IN CODE
    Joystick = 9
    Internet = 10
    Multimedia = 11
    Network = 12
    Password = 13
    Power = 14
    Scanner = 15
    Phone = 16
    ODBC = 17
    Themes = 18
End Enum
Public Type RGB_WINVER
  PlatformID      As Long
  VersionName     As String
  VersionNo       As String
  ServicePack     As String
  BuildNo         As String
End Type
Private Const cModuleName = "SysUtility"
Public Function GetExplorerVersion() As String
    GetExplorerVersion = GetIEVersion
End Function
Public Function IsWin2000() As Boolean
    IsWin2000 = modUtility.IsWin2000
End Function
Public Function IsWin95() As Boolean
    IsWin95 = modUtility.IsWin95
End Function
Public Function IsWin98() As Boolean
    IsWin98 = modUtility.IsWin98
End Function
Public Function IsWinME() As Boolean
    IsWinME = modUtility.IsWinME
End Function
Public Function IsWinNT() As Boolean
    IsWinNT = modUtility.IsWinNT
End Function
Public Function IsWinNT4() As Boolean
    IsWinNT4 = modUtility.IsWinNT4
End Function
Public Function IsWinXP() As Boolean
    IsWinXP = modUtility.IsWinXP
End Function
Public Function MinimumVersion(ByVal VerMajor As Long, ByVal VerMinor As Long, Optional BuildNumber As Long = 0)
    MinimumVersion = modUtility.IsVersionAtLeast(VerMajor, VerMinor, BuildNumber)
End Function
Public Function GetWinVersion(ByRef WIN_STRUCT As RGB_WINVER) As String
    GetWinVersion = modUtility.GetWinVersion(WIN_STRUCT)
End Function
Public Function GetWinVer() As String
    GetWinVer = modUtility.GetWinVer
End Function
Public Function getTemporaryFile(Optional ByVal sDesc As String = vbNullString) As String
    getTemporaryFile = modFileRoutines.getTemporaryFile(sDesc)
End Function
Public Function GetFileName(sPathIn As String) As String
    GetFileName = modFileRoutines.GetFileName(sPathIn)
End Function
Public Function GetExtension(sPathIn As String) As String
    GetExtension = modFileRoutines.GetExtension(sPathIn)
End Function
Public Function GetPath(sPathIn As String) As String
    GetPath = modFileRoutines.GetPath(sPathIn)
End Function
Public Function StripPath(sPathIn As String) As String
    StripPath = modFileRoutines.StripPath(sPathIn)
End Function
Public Function FileExists(ByVal strSearchFile As String) As Boolean
    FileExists = modFileRoutines.FileExists(strSearchFile)
End Function
Public Function FolderExists(ByVal strPath As String) As Boolean
    FolderExists = modFileRoutines.FolderExists(strPath)
End Function
Public Function FileCopy(ByVal sOrig As String, ByVal sDest As String, Optional ByVal bFailIfExist As Boolean = False) As Boolean
    FileCopy = modFileRoutines.FileCopy(sOrig, sDest, bFailIfExist)
End Function
Public Function IsValidArray(ByRef this As Variant) As Boolean
    IsValidArray = modUtility.IsValidArray(this)
End Function
Public Function GetWindowsSchemes() As String()
Dim oReg As New Registry
Dim sKeys() As String
Dim iKeyCount As Long
Const cScheme = "Control Panel\Appearance\Schemes"

On Error GoTo err_GetWindowsSchemes
    With oReg
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = cScheme
        .ValueType = REG_BINARY
        .EnumerateValues sKeys(), iKeyCount
    End With
    GoTo exit_GetWindowsSchemes
    
err_GetWindowsSchemes:
    UEH_LogError cModuleName, "GetWindowsSchemes", Err
    Err.Clear
exit_GetWindowsSchemes:
    Set oReg = Nothing
    GetWindowsSchemes = sKeys()
    
End Function
Private Function UpdateFromDefault(ByVal psScheme As String) As Boolean
Dim oReg As New Registry
Const cScheme = "Control Panel\Appearance\Schemes"
Dim bReturn As Boolean
Dim myArray As Variant

    bReturn = False
On Error GoTo err_UpdateFromDefault
    myArray = GetDefaultSchemeSettings(psScheme)
    If Not IsEmpty(myArray) Then
        UEH_Log cModuleName, "UpdateFromDefault", "Default Information Found for (" & psScheme & ")", logVerbose
        With oReg
            .ClassKey = HKEY_CURRENT_USER
            .SectionKey = cScheme
            .ValueKey = psScheme
            .ValueType = REG_BINARY
            .Value = myArray
        End With
        bReturn = True
    Else
        UEH_Log cModuleName, "UpdateFromDefault", "Default Information NOT Found for (" & psScheme & ")", logWarning
        bReturn = False
    End If
    GoTo exit_UpdateFromDefault
err_UpdateFromDefault:
    'LOG ERROR
    UEH_LogError cModuleName, "UpdateFromDefault", Err
    Err.Clear
exit_UpdateFromDefault:
    Set oReg = Nothing
    UpdateFromDefault = bReturn
End Function
Private Function GetDefaultSchemeSettings(ByVal psScheme As String) As Variant
Dim oReg As New Registry
Dim VR As Variant
Const cScheme = ".DEFAULT\Control Panel\Appearance\Schemes"
    UEH_Log cModuleName, "GetDefaultSchemeSettings", "Reading registry for .DEFAULT (" & psScheme & ")", logVerbose
    With oReg
        .ClassKey = HKEY_USERS
        .SectionKey = cScheme
        .ValueKey = psScheme
            If .KeyExists Then
                UEH_Log cModuleName, "GetDefaultSchemeSettings", "Found .DEFAULT settings for (" & psScheme & ")", logVerbose
                VR = .Value
            Else
                UEH_Log cModuleName, "GetDefaultSchemeSettings", "No .DEFAULT settings for (" & psScheme & ")", logError
            End If
        GetDefaultSchemeSettings = VR
    End With
    Set oReg = Nothing
End Function
Private Function SaveCurrentScheme(ByVal psScheme As String) As Boolean
Dim oReg As New Registry
Dim bReturn As Boolean
Const cScheme = "Control Panel\Current" 'CHANGES IN BOTH PLACES BY CHANGING IT HERE...(Control Panel\Appearance\Current)
bReturn = False
On Error GoTo err_SaveCurrentScheme
    UEH_Log cModuleName, "SaveCurrentScheme", "Reading registry for (" & psScheme & ")", logVerbose
    With oReg
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = cScheme
        .ValueKey = "Color Schemes" 'CHANGES IN BOTH PLACES BY CHANGING IT HERE...(Control Panel\Appearance\Current)
        .ValueType = REG_SZ
        .Value = psScheme
    End With
    bReturn = True
    GoTo exit_SaveCurrentScheme
err_SaveCurrentScheme:
    UEH_LogError cModuleName, "SaveCurrentScheme", Err
    Err.Clear
exit_SaveCurrentScheme:
    Set oReg = Nothing
    SaveCurrentScheme = bReturn
    UEH_Log cModuleName, "SaveCurrentScheme", "SaveCurrentScheme = " & CStr(bReturn), logVerbose
End Function
Public Function GetCurrentWindowsScheme() As String
Dim oReg As New Registry
Dim sReturn As String
Const cScheme = "Control Panel\Appearance" ' THIS IS WHERE WINDOWS MAKES THE CHANGE"Control Panel\Current" '
On Error GoTo err_GetCurrentWindowsScheme
    UEH_Log cModuleName, "GetCurrentWindowsScheme", "Reading registry...", logVerbose
    With oReg
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = cScheme
        .ValueKey = "Current" 'MATCH WHERE WINDOWS MAKES CHANGES"Color Schemes" '
        .ValueType = REG_SZ
        sReturn = .Value
    End With
    UEH_Log cModuleName, "GetCurrentWindowsScheme", "Current scheme = " & sReturn, logVerbose
    GoTo exit_GetCurrentWindowsScheme
err_GetCurrentWindowsScheme:
    UEH_LogError cModuleName, "GetCurrentWindowsScheme", Err
    Err.Clear
exit_GetCurrentWindowsScheme:
    Set oReg = Nothing
    GetCurrentWindowsScheme = sReturn
End Function
Public Function OpenSystemApplet(ByVal cpApplet As cpApplets, Optional ByVal Pages As Long) As Boolean
    OpenSystemApplet = RunDLL(TranslateApplet(cpApplet), Pages)
End Function
Private Function TranslateApplet(ByVal cpApplet As HTE_SystemUtility.cpApplets) As String
Select Case cpApplet
    Case Display
        TranslateApplet = "display"
    Case Add_Remove
        TranslateApplet = "addremove"
    Case Internet
        TranslateApplet = "internet"
    Case Regional_Options
        TranslateApplet = "regional"
    Case Joystick
        TranslateApplet = "joystick"
    Case Mouse
        TranslateApplet = "mouse"
    Case Multimedia
        TranslateApplet = "multimedia"
    Case Network
        TranslateApplet = "network"
    Case Password
        TranslateApplet = "password"
    Case Power
        TranslateApplet = "power"
    Case Scanner
        TranslateApplet = "scanner"
    Case System
        TranslateApplet = "system"
    Case Date_Time
        TranslateApplet = "timedate"
    Case Accessibility
        TranslateApplet = "access"
    Case Phone
        TranslateApplet = "phone"
    Case ODBC
        TranslateApplet = "odbc"
    Case Themes
        TranslateApplet = "themes"
    Case Else
        TranslateApplet = "unknown"
    End Select
    UEH_Log cModuleName, "TranslateApplet", "TranslateApplet = " & TranslateApplet, logVerbose
End Function
Public Function SetWindowsColorScheme(ByVal psScheme As String) As Boolean
'Need to change AND update registry
Dim sKeys() As String
Dim x As Long
Dim bReturn As Boolean
Dim bExists As Boolean
'Dim oObj As SYSTEMUTILITIESLib.HTE_SystemUtilities
    bReturn = False
    UEH_Log cModuleName, "SetWindowsColorScheme", "psScheme = " & psScheme, logVerbose
    On Error GoTo err_SetWindowsColorScheme
        sKeys = GetWindowsSchemes
        If IsArray(sKeys) Then
            bExists = False
            For x = 1 To UBound(sKeys)
                bExists = (StrComp(psScheme, sKeys(x), vbTextCompare) = 0)
                If bExists Then Exit For
            Next
            UEH_Log cModuleName, "SetWindowsColorScheme", "Existing HKEY_USER scheme = " & CStr(bExists), logVerbose
            If Not bExists Then 'CHECK UNDER DEFAULT SETTINGS AND SAVE UNDER CURRENT USER
                bExists = UpdateFromDefault(psScheme)
            End If
        End If
        If bExists Then
            If StrComp(GetCurrentWindowsScheme, psScheme, vbBinaryCompare) <> 0 Then
                UEH_Log cModuleName, "SetWindowsColorScheme", "Changing current scheme", logVerbose
                If IsWinNT Then
                    UEH_Log cModuleName, "SetWindowsColorScheme", "WinNT change", logVerbose
                    bReturn = (DeskSetCurrentScheme(psScheme) <> 0)
                Else
                    UEH_Log cModuleName, "SetWindowsColorScheme", "Win9x change", logVerbose
                    bReturn = HorribleWin9xHack(psScheme)
                End If
                UEH_Log cModuleName, "SetWindowsColorScheme", "Saving current scheme", logVerbose
                bReturn = bReturn And SaveCurrentScheme(psScheme)
            Else
                'NO SENSE IN CHANGING IF ALREADY SET
                UEH_Log cModuleName, "SetWindowsColorScheme", "Requested scheme is the same as current NOT changing", logVerbose
                bReturn = True
            End If
        Else
            UEH_Log cModuleName, "SetWindowsColorScheme", "No Existing HKEY_USER or Default scheme ", logWarning
        End If
        GoTo exit_SetWindowsColorScheme
err_SetWindowsColorScheme:
    UEH_LogError cModuleName, "SetWindowsColorScheme", Err
    Err.Clear
exit_SetWindowsColorScheme:
    UEH_Log cModuleName, "SetWindowsColorScheme", "SetWindowsColorScheme = " & CStr(bReturn), logVerbose
    SetWindowsColorScheme = bReturn

End Function

Private Sub Class_Initialize()
    UEH_BeginLogging , App.EXEName
End Sub

Private Sub Class_Terminate()
    UEH_EndLogging
End Sub
