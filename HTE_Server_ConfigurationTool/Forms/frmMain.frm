VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Object = "{1C203F10-95AD-11D0-A84B-00A0247B735B}#1.0#0"; "SSTree.ocx"
Object = "{85202277-6C76-4228-BC56-7B3E69E8D5CA}#5.0#0"; "IGToolBars50.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Communitas © Server Maintenance Tool"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9075
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   1560
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      ToolBarsCount   =   2
      ToolsCount      =   29
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tools           =   "frmMain.frx":6852
      ToolBars        =   "frmMain.frx":1DAA8
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   1440
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      ResizeStyle     =   0
      ResizeFonts     =   0   'False
      DesignWidth     =   9075
      DesignHeight    =   6210
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Align           =   1  'Align Top
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   10398
      _Version        =   262144
      AutoSize        =   1
      SplitterBarJoinStyle=   0
      SplitterResizeStyle=   1
      SplitterBarAppearance=   1
      BorderStyle     =   0
      PaneTree        =   "frmMain.frx":1E22D
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   4545
         Left            =   2415
         TabIndex        =   2
         Top             =   0
         Width           =   6660
         ExtentX         =   11747
         ExtentY         =   8017
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
      Begin SSActiveTreeView.SSTree SSTree1 
         Height          =   5895
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   10398
         _Version        =   65536
         Appearance      =   0
         LineStyle       =   1
         Indentation     =   315
         Sorted          =   1
         PictureBackgroundUseMask=   0   'False
         HasFont         =   -1  'True
         HasMouseIcon    =   0   'False
         HasPictureBackground=   0   'False
         ImageList       =   "(None)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Webbrowser panel interface
Private custDoc As ICustomDoc
Private WithEvents IDocProxy As docHostUIHandler
Attribute IDocProxy.VB_VarHelpID = -1
Private m_EnableDefaultBrowserKeys As Boolean
Private Const cBlankPage = "about:blank"
Private m_Connection As ADODB.Connection

Private Sub PrepareHostHandler()
    Set IDocProxy = New docHostUIHandler
    IDocProxy.IEHotKeysEnabled = m_EnableDefaultBrowserKeys
    WebBrowser1.Navigate cBlankPage
End Sub

Private Sub PrepareConnection()
    Set m_Connection = PrepareDatabaseConnection
End Sub

Private Sub PrepareTree()
    If m_Connection.State = adStateOpen Then
        SSTree1.Nodes.Add , , "Agency", "Agencies"
        SSTree1.Nodes.Add , , "Department", "Departments"
        SSTree1.Nodes.Add , , "Groups", "Groups"
        SSTree1.Nodes.Add , , "Devices", "Devices"
        SSTree1.Nodes.Add , , "Locations", "Locations"
    End If
End Sub

Private Sub Form_Load()
    PrepareHostHandler
    PrepareConnection
    PrepareTree
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set IDocProxy = Nothing
    EndApp
End Sub

Private Sub SSTree1_NodeClick(Node As SSActiveTreeView.SSNode)
Dim rs As ADODB.Recordset
Dim keyValue As String, DescValue As String, SQL As String
    
    With Node
        If Node.children = 0 Then
            Select Case Node.Key
                Case "Agency"
                    keyValue = "AgencyID": DescValue = "Description": SQL = "Select * from AGENCY"
                Case "Devices"
                    keyValue = "DeviceID": DescValue = "Description": SQL = "Select * from Devices"
                Case "Departments"
                    keyValue = "DepartmentID": DescValue = "Description": SQL = "Select * from Departments"
                Case "Groups"
                    keyValue = "GroupID": DescValue = "Desc": SQL = "Select * from Groups"
                Case "Locations"
                    keyValue = "LocationID": DescValue = "Alias": SQL = "Select * from Location"
            End Select
        End If
        If SQL <> vbNullString Then
            Set rs = PrepareRecordset(SQL, m_Connection)
            Do Until rs.EOF
                SSTree1.Nodes.Add Node, ssatChild, rs.Fields(keyValue).Value, rs.Fields(DescValue).Value & vbNullString
                rs.MoveNext
            Loop
        End If
    End With
    
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    If Not IDocProxy Is Nothing Then
        If WebBrowser1.Document Is Nothing Or TypeOf WebBrowser1.Document Is HTMLDocument Then
            If Not InDevelopment Then
                Set IDocProxy.Document = WebBrowser1.Document
            End If
            WebBrowser1.Object.Document.body.Style.BorderStyle = "none" 'nice trick for flat browser  - MUCH simpler than IOLEClient sink
        End If
    End If
End Sub

