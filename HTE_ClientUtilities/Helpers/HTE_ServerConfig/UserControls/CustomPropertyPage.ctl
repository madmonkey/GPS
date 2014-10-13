VERSION 5.00
Begin VB.UserControl CustomPropertyPage 
   ClientHeight    =   3075
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
   ScaleHeight     =   3075
   ScaleWidth      =   5655
   ToolboxBitmap   =   "CustomPropertyPage.ctx":0000
   Begin VB.CheckBox chkEnableMAC 
      Alignment       =   1  'Right Justify
      Caption         =   "Resolve MAC?"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Issue ARP command to attempt to obtain physical device identifier"
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txtCacheLookup 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   3
      ToolTipText     =   "Cache succesful MAC lookup (seconds)"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtPort 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      ToolTipText     =   "Port in which to listen for incoming messages"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cache Lookup:"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Port:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
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

Private Sub chkEnableMAC_Click()
    bChanged = bChanged Or letProperty(m_Settings, cEnableMacResolution, CStr(CBool(-chkEnableMAC.Value)))
End Sub

Private Sub txtCacheLookup_Change()
    bChanged = bChanged Or letProperty(m_Settings, cCacheMacLookupSeconds, txtCacheLookup.Text)
End Sub

Private Sub txtCacheLookup_GotFocus()
    Highlight txtCacheLookup
End Sub

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
        txtPort.Text = getProperty(m_Settings, cPortNumber, cPortNumberValue)
        chkEnableMAC.Value = Abs(CBool(getProperty(m_Settings, cEnableMacResolution, cEnableMacResolutionValue)))
        txtCacheLookup.Text = Val(getProperty(m_Settings, cCacheMacLookupSeconds, cCacheMacLookupSecondsValue))
        '*TODO - ADD CUSTOM PROPERTY GET FUNCTIONS HERE
        '*TODO - WHEN A PROPERTY CHANGES YOU WILL NEED TO FLAG YOUR CHANGED VARIABLE
        '* Example: bChanged = bChanged Or letProperty(m_Settings, cInBufferSize, txtInBufferSize.Text)
    End If
End Function

Private Sub txtPort_Change()
    bChanged = bChanged Or letProperty(m_Settings, cPortNumber, txtPort.Text)
End Sub

Private Sub txtPort_GotFocus()
    Highlight txtPort
End Sub

Private Sub UserControl_Initialize()
    setStyle txtPort.hWnd, esNumeric
    setStyle txtCacheLookup.hWnd, esNumeric
    txtCacheLookup.Text = cCacheMacLookupSecondsValue
End Sub
