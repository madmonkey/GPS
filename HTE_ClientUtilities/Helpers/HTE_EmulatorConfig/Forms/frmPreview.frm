VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmPreview 
   Appearance      =   0  'Flat
   Caption         =   "Preview Playback"
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
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private custDoc As ICustomDoc
Private WithEvents IDocProxy As docHostUIHandler
Attribute IDocProxy.VB_VarHelpID = -1
Private m_EnableDefaultBrowserKeys As Boolean
Private Const cBlankPage = "about:blank"

Private Sub Form_Load()
    PrepareHostHandler
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        If Me.Width > 110 And Me.Height > 385 Then
            txtPreview.Move 0, 0, Me.Width - 110, Me.Height - 510
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set IDocProxy = Nothing
End Sub

Private Sub txtPreview_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    txtPreview.Object.Document.body.Style.BorderStyle = "none" 'nice trick for flat browser  - MUCH simpler than IOLEClient sink
End Sub

Private Sub PrepareHostHandler()
    Set IDocProxy = New docHostUIHandler
    IDocProxy.IEHotKeysEnabled = m_EnableDefaultBrowserKeys
    txtPreview.Navigate cBlankPage
End Sub
