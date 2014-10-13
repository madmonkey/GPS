VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmPreview 
   Appearance      =   0  'Flat
   Caption         =   "Preview Stylesheet"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   9870
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   9870
   StartUpPosition =   1  'CenterOwner
   Begin SHDocVwCtl.WebBrowser txtPreview 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      ExtentX         =   17383
      ExtentY         =   10610
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''
'''Private hw As Long
'''
'''Private Const GW_CHILD = 5
'''Private Const GW_HWNDNEXT = 2
'''
'''Private Const WM_RBUTTONDOWN = &H204
'''Private Const WM_RBUTTONUP = &H205
'''Private Const WM_RBUTTONDBLCLK = &H206
'''
'''Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'''Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
'''Private sc As Subclass
'''Implements ISubclass
'''
'''Private Sub Form_Load()
'''Dim h As Long, aClass As String, k As Long
'''    h = GetWindow(hWnd, GW_CHILD)
'''    aClass = Space$(128)
'''    Do While h
'''        k = GetClassName(h, aClass, 128)
'''        If Left$(aClass, k) = "Shell Embedding" Then hw = h: Exit Do
'''        h = GetWindow(h, GW_HWNDNEXT)
'''    Loop
'''    Set sc = New Subclass
'''    sc.AttachMessage (WM_RBUTTONDOWN)
'''    sc.AttachMessage (WM_RBUTTONUP)
'''    sc.AttachMessage (WM_RBUTTONDBLCLK)
'''    sc.Subclass hw, Me
'''End Sub
'''
'''Private Sub Form_Unload(Cancel As Integer)
'''    sc.UnSubclass
'''    Set sc = Nothing
'''End Sub
'''
'''Private Sub ISubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'''
'''End Sub
'''
'''Private Sub ISubclass_Before(lHandled As Long, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'''    Select Case uMsg
'''        Case WM_RBUTTONDOWN, WM_RBUTTONUP, WM_RBUTTONDBLCLK
'''
'''    End Select
'''End Sub
Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        If Me.Width > 110 And Me.Height > 385 Then
            txtPreview.Move 0, 0, Me.Width - 110, Me.Height - 385
        End If
    End If
End Sub

