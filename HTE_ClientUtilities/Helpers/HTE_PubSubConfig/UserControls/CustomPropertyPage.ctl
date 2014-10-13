VERSION 5.00
Begin VB.UserControl CustomPropertyPage 
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
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
   ScaleHeight     =   2475
   ScaleWidth      =   5655
   ToolboxBitmap   =   "CustomPropertyPage.ctx":0000
   Begin VB.Frame frSubscriber 
      Caption         =   "Subscriber"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   5415
      Begin VB.TextBox txtSubTag 
         Height          =   285
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   3
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label4 
         Caption         =   "Tag:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame frPublisher 
      Caption         =   "Publisher"
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5415
      Begin VB.TextBox txtTag 
         Height          =   285
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   1
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox txtTimeout 
         Height          =   285
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   2
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox txtTopic 
         Height          =   285
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   0
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label3 
         Caption         =   "Tag:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Timeout (sec):"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Topic:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "CustomPropertyPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements HTE_GPS.PropertyPage
Dim m_Callback As HTE_GPS.PropPageCallback
Dim bChanged As Boolean
Dim m_Settings As MSXML2.DOMDocument30

Private Property Get PropertyPage_Changed() As Boolean
    PropertyPage_Changed = bChanged
End Property

Private Sub PropertyPage_Exit()
    Set m_Callback = Nothing
End Sub

Private Property Get PropertyPage_LicenseKey() As String
    PropertyPage_LicenseKey = vbNullString
End Property

Private Property Get PropertyPage_Name() As String
    PropertyPage_Name = App.EXEName & "." & UserControl.Name
End Property

Private Property Let PropertyPage_PropertyCallback(RHS As HTE_GPS.PropPageCallback)
    Set m_Callback = RHS
End Property

Private Property Get PropertyPage_PropertyCallback() As HTE_GPS.PropPageCallback
     Set PropertyPage_PropertyCallback = m_Callback
End Property

Private Function PropertyPage_SaveChanges() As Boolean
Dim bReturn As Boolean
    If Not m_Callback Is Nothing Then
        If Not m_Settings Is Nothing Then
            bReturn = m_Callback.SaveChanges(m_Settings.xml)
            bChanged = Not bReturn
            PropertyPage_SaveChanges = bReturn
        End If
    End If
End Function

Private Property Let PropertyPage_Settings(ByVal RHS As String)
    loadLocalSettings (RHS)
    bChanged = False
End Property

Private Property Get PropertyPage_Settings() As String
    If Not m_Settings Is Nothing Then
        PropertyPage_Settings = m_Settings.xml
    End If
End Property

Private Function loadLocalSettings(ByVal sXML As String) As Boolean
    Set m_Settings = New MSXML2.DOMDocument30
    loadLocalSettings = m_Settings.loadXML(sXML)
    If loadLocalSettings Then
        txtTopic.Text = getProperty(m_Settings, cPublishTopic, cPublishTopicValue)
        txtTag.Text = getProperty(m_Settings, cPublishField, cPublishFieldValue)
        txtTimeout.Text = CLng(Val(getProperty(m_Settings, cDefaultTimeout, cDefaultTimeoutValue)))
        txtSubTag.Text = getProperty(m_Settings, cSubcribeField, cSubcribeFieldValue)
        '*TODO - ADD CUSTOM PROPERTY GET FUNCTIONS HERE
        '*TODO - WHEN A PROPERTY CHANGES YOU WILL NEED TO FLAG YOUR CHANGED VARIABLE
        '* Example: bChanged = bChanged Or letProperty(m_Settings, cInBufferSize, txtInBufferSize.Text)
    End If
End Function

Private Sub txtSubTag_Change()
    bChanged = bChanged Or letProperty(m_Settings, cSubcribeField, txtSubTag.Text)
End Sub

Private Sub txtSubTag_GotFocus()
    Highlight txtSubTag
End Sub

Private Sub txtTag_Change()
    bChanged = bChanged Or letProperty(m_Settings, cPublishField, txtTag.Text)
End Sub

Private Sub txtTag_GotFocus()
    Highlight txtTag
End Sub

Private Sub txtTimeout_Change()
    bChanged = bChanged Or letProperty(m_Settings, cDefaultTimeout, txtTimeout.Text)
End Sub

Private Sub txtTimeout_GotFocus()
    Highlight txtTimeout
End Sub

Private Sub txtTopic_Change()
    bChanged = bChanged Or letProperty(m_Settings, cPublishTopic, txtTopic.Text)
End Sub

Private Sub txtTopic_GotFocus()
    Highlight txtTopic
End Sub

Private Sub UserControl_Initialize()
    txtTopic.Text = cPublishTopicValue
    txtTag.Text = cPublishFieldValue
    txtTimeout.Text = cDefaultTimeoutValue
    txtSubTag.Text = cSubcribeFieldValue
    setStyle txtTimeout.hWnd, esNumeric
End Sub

Private Sub UserControl_Resize()
    If UserControl.ScaleWidth > 240 Then
        frPublisher.Width = UserControl.ScaleWidth - 240
        frSubscriber.Width = UserControl.ScaleWidth - 240
    End If
End Sub
