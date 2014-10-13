VERSION 5.00
Begin VB.Form frmServiceControl 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modular GPS Service Controller Application"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   180
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServiceControl.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrCheck 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5520
      Top             =   1440
   End
   Begin VB.CheckBox chkSystem 
      Appearance      =   0  'Flat
      Caption         =   "System Account"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtAccount 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   3735
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Service"
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "Install Service"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblAccount 
      BackStyle       =   0  'Transparent
      Caption         =   "Account:"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   240
      X2              =   5910
      Y1              =   1910
      Y2              =   1910
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   5910
      Y1              =   1920
      Y2              =   1920
   End
End
Attribute VB_Name = "frmServiceControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(1 To 128) As Byte
End Type
Private Const VER_PLATFORM_WIN32_NT = 2&
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1&
Private ServState As SERVICE_STATE
Private Installed As Boolean

Private Sub chkSystem_Click()
    If chkSystem Then
        txtAccount = "LocalSystem"
        txtAccount.Enabled = False
        txtPassword.Enabled = False
        lblAccount.Enabled = False
        lblPassword.Enabled = False
    Else
        txtAccount = vbNullString
        txtAccount.Enabled = True
        txtPassword.Enabled = True
        lblAccount.Enabled = True
        lblPassword.Enabled = True
    End If
End Sub

Private Sub cmdInstall_Click()
    CheckService
    If Not cmdInstall.Enabled Then Exit Sub
    cmdInstall.Enabled = False
    If Installed Then
        DeleteNTService
    Else
        SetNTService
        txtPassword = vbNullString
    End If
    CheckService
End Sub

' This sub checks service status
Private Sub CheckService()
    If GetServiceConfig() = 0 Then
        Installed = True
        cmdInstall.Caption = "Uninstall Service"
        txtAccount.Enabled = False
        txtPassword.Enabled = False
        lblAccount.Enabled = False
        lblPassword.Enabled = False
        chkSystem.Enabled = False
        ServState = GetServiceStatus()
        Select Case ServState
            Case SERVICE_RUNNING
                cmdInstall.Enabled = False
                cmdStart.Caption = "Stop Service"
                cmdStart.Enabled = True
            Case SERVICE_STOPPED
                cmdInstall.Enabled = True
                cmdStart.Caption = "Start Service"
                cmdStart.Enabled = True
            Case Else
                cmdInstall.Enabled = False
                cmdStart.Enabled = False
        End Select
    Else
        Installed = False
        cmdInstall.Caption = "Install Service"
        txtAccount.Enabled = chkSystem = 0
        txtPassword.Enabled = chkSystem = 0
        lblAccount.Enabled = chkSystem = 0
        lblPassword.Enabled = chkSystem = 0
        chkSystem.Enabled = True
        cmdStart.Enabled = False
        cmdInstall.Enabled = True
    End If
End Sub

Private Sub cmdStart_Click()
    CheckService
    If Not cmdStart.Enabled Then Exit Sub
    cmdStart.Enabled = False
    If ServState = SERVICE_RUNNING Then
        StopNTService
        'M$ has changed the way the function call works (security?)
        'This hopefully will kill the process to allow it to be restarted
        'KillApp "HTE_ClientUtilities.exe"
    ElseIf ServState = SERVICE_STOPPED Then
        StartNTService
    End If
    CheckService
End Sub

Private Sub Form_Load()
    If Not CheckIsNT() Then
        MsgBox "This program requires Windows NT/2000/XP/2003", vbInformation, "Modular GPS Service Controller"
        Unload Me
        Exit Sub
    End If
    AppPath = App.Path
    If Right$(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
    chkSystem_Click
    CheckService
    tmrCheck.Enabled = True
End Sub

' This sub opens blank letter with filled address field
' in default e-mail client

'''Private Sub lblEmail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    If Button = 1 Then
'''        ShellExecute Me.hwnd, "open", "mailto:" & lblEmail.Caption, vbNullString, App.Path, SW_SHOWNORMAL
'''    End If
'''End Sub
'''
'''' This sub opens Web page in default browser
'''
'''Private Sub lblWeb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    If Button = 1 Then
'''        ShellExecute Me.hwnd, "open", lblWeb.Caption, vbNullString, App.Path, SW_SHOWNORMAL
'''    End If
'''End Sub


' CheckIsNT() returns True, if the program runs
' under Windows NT or Windows 2000, and False
' otherwise.

Private Function CheckIsNT() As Boolean
    Dim OSVer As OSVERSIONINFO
    OSVer.dwOSVersionInfoSize = LenB(OSVer)
    GetVersionEx OSVer
    CheckIsNT = (OSVer.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    EndApp
End Sub

Private Sub tmrCheck_Timer()
    CheckService
End Sub
