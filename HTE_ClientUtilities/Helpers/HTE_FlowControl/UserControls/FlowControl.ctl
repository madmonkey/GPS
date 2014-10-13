VERSION 5.00
Begin VB.UserControl FlowControl 
   ClientHeight    =   1410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   ScaleHeight     =   1410
   ScaleWidth      =   6435
   ToolboxBitmap   =   "FlowControl.ctx":0000
   Begin VB.Frame frPseudoMode 
      Appearance      =   0  'Flat
      Caption         =   "Software Reporting Mode"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6375
      Begin VB.TextBox txtDuration 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5640
         MaxLength       =   5
         TabIndex        =   2
         Text            =   "0"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtRelay 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3960
         MaxLength       =   5
         TabIndex        =   1
         Text            =   "0"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtInterval 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   960
         Width           =   4815
      End
      Begin VB.TextBox txtSentence 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   600
         Width           =   4815
      End
      Begin VB.CheckBox chkManual 
         Appearance      =   0  'Flat
         Caption         =   "Use SRM?"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Duration:"
         Height          =   255
         Left            =   4080
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblSendOnly 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Relay in sec(s)"
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Interval(s):"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Sentence(s):"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FlowControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Event SoftwareReportingClick()
Event RelayChange()
Event SentenceChange()
Event IntervalChange()
Event DurationChange()

Public Property Get UseSoftwareReporting() As Boolean
    UseSoftwareReporting = CBool(-chkManual.Value)
End Property

Public Property Let UseSoftwareReporting(ByVal vData As Boolean)
    chkManual.Value = Abs(vData)
End Property

Public Property Get ReportingSentence() As String
    ReportingSentence = txtSentence.Text
End Property

Public Property Let ReportingSentence(ByVal vData As String)
    txtSentence.Text = vData
End Property

Public Property Get ReportingInterval() As String
    ReportingInterval = txtInterval.Text
End Property

Public Property Let ReportingInterval(ByVal vData As String)
    txtInterval.Text = vData
End Property

Public Property Get RelayInterval() As Long
    If Not IsNumeric(txtRelay.Text) Then
        txtRelay.Text = 0
    End If
    RelayInterval = txtRelay.Text
End Property

Public Property Let RelayInterval(ByVal vData As Long)
    txtRelay.Text = vData
End Property

Public Property Get DurationInterval() As Long
    If Not IsNumeric(txtDuration.Text) Then
        txtDuration.Text = 0
    End If
    DurationInterval = txtDuration.Text
End Property

Public Property Let DurationInterval(ByVal vData As Long)
    txtDuration.Text = vData
End Property

Private Sub chkManual_Click()
    RaiseEvent SoftwareReportingClick
End Sub

Private Sub txtDuration_Change()
    RaiseEvent DurationChange
End Sub

Private Sub txtDuration_GotFocus()
    Highlight txtDuration
End Sub

Private Sub txtRelay_Change()
    RaiseEvent RelayChange
End Sub

Private Sub txtRelay_GotFocus()
    Highlight txtRelay
End Sub

Private Sub txtSentence_Change()
    RaiseEvent SentenceChange
End Sub

Private Sub txtSentence_GotFocus()
    Highlight txtSentence
End Sub

Private Sub txtInterval_Change()
    RaiseEvent IntervalChange
End Sub

Private Sub txtInterval_GotFocus()
    Highlight txtInterval
End Sub

Private Sub UserControl_Initialize()
    setStyle txtRelay.hWnd, esNumeric
    setStyle txtDuration.hWnd, esNumeric
End Sub

Private Sub UserControl_Resize()
    Size 6435, 1410
End Sub
