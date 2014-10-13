VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSysTray 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   645
   LinkTopic       =   "Form1"
   ScaleHeight     =   675
   ScaleWidth      =   645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ILStatusList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysTray.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysTray.frx":0395
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysTray.frx":0733
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysTray.frx":0AC9
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Const WM_CANCELMODE = &H1F
' Track popup menu constants:
Private Const TPM_CENTERALIGN = &H4&
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_LEFTBUTTON = &H0&
Private Const TPM_RIGHTALIGN = &H8&
Private Const TPM_RIGHTBUTTON = &H2&
Private Const TPM_NONOTIFY = &H80&           '/* Don't send any notification msgs */
Private Const TPM_RETURNCMD = &H100
Private Const TPM_HORIZONTAL = &H0          '/* Horz alignment matters more */
Private Const TPM_VERTICAL = &H40           '/* Vert alignment matters more */
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const MAX_TOOLTIP As Integer = 64
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * MAX_TOOLTIP
End Type
Private nfIconData As NOTIFYICONDATA
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As RECT) As Long

Public Event SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
Public Event SysTrayMouseUp(ByVal eButton As MouseButtonConstants)
Public Event SysTrayMouseMove()
Public Event SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)

Private m_bAddedMenuItem As Boolean
Private m_hWndOwner As Long

Public Sub Initialize(ByVal hWndOwner As Long)
   m_hWndOwner = hWndOwner
   With nfIconData
       .hwnd = Me.hwnd
       .uID = Me.Icon
       .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
       .uCallbackMessage = WM_MOUSEMOVE
       .hIcon = Me.Icon.Handle
       .szTip = App.FileDescription & Chr$(0)
       .cbSize = Len(nfIconData)
   End With
   Shell_NotifyIcon NIM_ADD, nfIconData
End Sub
Public Property Get ToolTip() As String
Dim sTip As String
Dim iPos As Long
    sTip = nfIconData.szTip
    iPos = InStr(sTip, Chr$(0))
    If (iPos <> 0) Then
        sTip = Left$(sTip, iPos - 1)
    End If
    ToolTip = sTip
End Property
Public Property Let ToolTip(ByVal sTip As String)
    If (sTip & Chr$(0) <> nfIconData.szTip) Then
        nfIconData.szTip = sTip & Chr$(0)
        nfIconData.uFlags = NIF_TIP
        Shell_NotifyIcon NIM_MODIFY, nfIconData
    End If
End Property
Public Property Get IconHandle() As Long
    IconHandle = nfIconData.hIcon
End Property
Public Property Let IconHandle(ByVal hIcon As Long)
    If (hIcon <> nfIconData.hIcon) Then
        nfIconData.hIcon = hIcon
        nfIconData.uFlags = NIF_ICON
        Shell_NotifyIcon NIM_MODIFY, nfIconData
    End If
End Property

Private Sub Form_Load()
    ILStatusList.MaskColor = vbMagenta
    ILStatusList.UseMaskColor = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lX As Long

   lX = Me.ScaleX(x, Me.ScaleMode, vbPixels)
   Select Case lX
   Case WM_MOUSEMOVE
      RaiseEvent SysTrayMouseMove
   Case WM_LBUTTONDOWN
      RaiseEvent SysTrayMouseDown(vbLeftButton)
   Case WM_LBUTTONUP
      RaiseEvent SysTrayMouseUp(vbLeftButton)
   Case WM_LBUTTONDBLCLK
      RaiseEvent SysTrayDoubleClick(vbLeftButton)
   Case WM_RBUTTONDOWN
      RaiseEvent SysTrayMouseDown(vbRightButton)
   Case WM_RBUTTONUP
      RaiseEvent SysTrayMouseUp(vbRightButton)
   Case WM_RBUTTONDBLCLK
      RaiseEvent SysTrayDoubleClick(vbRightButton)
   End Select

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Shell_NotifyIcon NIM_DELETE, nfIconData
End Sub
