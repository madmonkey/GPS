VERSION 5.00
Object = "{F9B2917F-12A4-4C2D-9B9B-4E1EDE8126BB}#1.1#0"; "HTE_FlowControl.ocx"
Begin VB.UserControl CustomPropertyPage 
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6585
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
   ScaleHeight     =   3840
   ScaleWidth      =   6585
   ToolboxBitmap   =   "CustomPropertyPage.ctx":0000
   Begin VB.TextBox txtCacheLookup 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4920
      MaxLength       =   5
      TabIndex        =   8
      ToolTipText     =   "Cache succesful MAC lookup (seconds)"
      Top             =   1920
      Width           =   615
   End
   Begin VB.CheckBox chkEnableMAC 
      Alignment       =   1  'Right Justify
      Caption         =   "Resolve MAC?"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      ToolTipText     =   "Issue ARP command to attempt to obtain physical device identifier"
      Top             =   1560
      Width           =   2055
   End
   Begin HTE_FlowControl.FlowControl FlowControl1 
      Height          =   1410
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   2487
   End
   Begin VB.TextBox txtLocalPort 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3480
      TabIndex        =   3
      ToolTipText     =   "Local Port to receive transmissions from."
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txtLocalIP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   3480
      TabIndex        =   13
      Text            =   "127.0.0.1"
      ToolTipText     =   "Local IP to receive transmissions from."
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txtIP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Text            =   "127.0.0.1"
      ToolTipText     =   "Remote Host to send transmissions to."
      Top             =   480
      Width           =   2055
   End
   Begin VB.CheckBox chkProcessSend 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Process complete on send?"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Raise a process complete on a send as well as a receive."
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CheckBox chkValidate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Validate message?"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Verify a valid message with proper start and end tags and keep tags with message"
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox txtInit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtPort 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      ToolTipText     =   "Remote Port to send transmissions to"
      Top             =   840
      Width           =   2055
   End
   Begin VB.ComboBox cmbProtocol 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cache Lookup:"
      Height          =   315
      Left            =   3480
      TabIndex        =   16
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblLocals 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Local Settings:"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3480
      TabIndex        =   14
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Init String:"
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Port:"
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Protocol:"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   855
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

Private Sub chkProcessSend_Click()
    bChanged = bChanged Or letProperty(m_Settings, cProcessOnSend, CStr(CBool(-chkProcessSend.Value)))
End Sub

Private Sub FlowControl1_DurationChange()
    bChanged = bChanged Or letProperty(m_Settings, cDurationInterval, FlowControl1.DurationInterval)
End Sub

'ADDED FOR DEVICE FLOWCONTROL
Private Sub FlowControl1_IntervalChange()
    bChanged = bChanged Or letProperty(m_Settings, cManualPollInterval, FlowControl1.ReportingInterval)
End Sub

Private Sub FlowControl1_RelayChange()
    bChanged = bChanged Or letProperty(m_Settings, cRelayInterval, FlowControl1.RelayInterval)
End Sub

Private Sub FlowControl1_SentenceChange()
    bChanged = bChanged Or letProperty(m_Settings, cManualPollString, ScreenMessage(FlowControl1.ReportingSentence))
End Sub

Private Sub FlowControl1_SoftwareReportingClick()
    bChanged = bChanged Or letProperty(m_Settings, cManualPoll, CStr(FlowControl1.UseSoftwareReporting))
End Sub
'ADDED FOR DEVICE FLOWCONTROL

Private Sub txtCacheLookup_Change()
    bChanged = bChanged Or letProperty(m_Settings, cCacheMacLookupSeconds, txtCacheLookup.Text)
End Sub

Private Sub txtCacheLookup_GotFocus()
    Highlight txtCacheLookup
End Sub

Private Sub txtLocalIP_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or _
        KeyAscii = 46 Or _
        KeyAscii = 8 Or _
        KeyAscii >= 37 And KeyAscii <= 40 Then
        'Leave it...numeric, backspace and period
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtLocalPort_Change()
    bChanged = bChanged Or letProperty(m_Settings, cLocalPort, txtLocalPort.Text)
End Sub

Private Sub txtLocalPort_GotFocus()
    Highlight txtLocalPort
End Sub

Private Sub chkValidate_Click()
    bChanged = bChanged Or letProperty(m_Settings, cValidateMessage, CStr(CBool(-chkValidate.Value)))
End Sub

Private Sub cmbProtocol_Click()
    bChanged = bChanged Or letProperty(m_Settings, cProtocol, cmbProtocol.List(cmbProtocol.ListIndex))
    AdditionalVisibles
End Sub

Private Sub cmbProtocol_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        bChanged = bChanged Or letProperty(m_Settings, cProtocol, cmbProtocol.List(cmbProtocol.ListIndex))
        AdditionalVisibles
    End If
End Sub

Private Sub AdditionalVisibles()
Dim bAssignAs As Boolean
    bAssignAs = StrComp(cmbProtocol.List(cmbProtocol.ListIndex), "UDP", vbTextCompare) = 0
    txtLocalIP.Visible = bAssignAs
    txtLocalPort.Visible = bAssignAs: If Not bAssignAs Then txtLocalPort.Text = cLocalPortValue
    lblLocals.Visible = bAssignAs
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
        cmbProtocol.ListIndex = Abs(getProperty(m_Settings, cProtocol, "UDP") <> "TCP")
        txtIP.Text = getProperty(m_Settings, cAddress, "127.0.0.1")
        txtPort.Text = getProperty(m_Settings, cPort, cPortValue)
        txtLocalPort.Text = getProperty(m_Settings, cLocalPort, cLocalPortValue)
        txtInit.Text = getProperty(m_Settings, cInitString, cInitStringValue)
        chkValidate.Value = Abs(CBool(getProperty(m_Settings, cValidateMessage, cValidateMessageValue)))
        chkProcessSend.Value = Abs(CBool(getProperty(m_Settings, cProcessOnSend, cProcessOnSendValue)))
        'ADDED FOR DEVICE FLOWCONTROL
        FlowControl1.UseSoftwareReporting = CBool(getProperty(m_Settings, cManualPoll, cManualPollValue))
        FlowControl1.ReportingSentence = UnscreenMessage(getProperty(m_Settings, cManualPollString, cManualPollStringValue))
        FlowControl1.ReportingInterval = getProperty(m_Settings, cManualPollInterval, cManualPollIntervalValue)
        FlowControl1.RelayInterval = CLng(getProperty(m_Settings, cRelayInterval, cRelayIntervalValue))
        FlowControl1.DurationInterval = CLng(getProperty(m_Settings, cDurationInterval, cDurationIntervalValue))
        'ADDED FOR DEVICE FLOWCONTROL
        'ADDED FOR ALIASING
        chkEnableMAC.Value = Abs(CBool(getProperty(m_Settings, cEnableMacResolution, cEnableMacResolutionValue)))
        txtCacheLookup.Text = CLng(Val(getProperty(m_Settings, cCacheMacLookupSeconds, cCacheMacLookupSecondsValue)))
        'ADDED FOR ALIASING
    End If
    
End Function

Private Sub txtInit_Change()
    bChanged = bChanged Or letProperty(m_Settings, cInitString, txtInit.Text)
    txtInit.ToolTipText = txtInit.Text
End Sub

Private Sub txtInit_GotFocus()
    Highlight txtInit
End Sub

Private Sub txtIP_Change()
    bChanged = bChanged Or letProperty(m_Settings, cAddress, txtIP.Text)
End Sub

Private Sub txtIP_GotFocus()
    Highlight txtIP
End Sub

Private Sub txtIP_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Or _
        KeyAscii = 46 Or _
        KeyAscii = 8 Or _
        KeyAscii >= 37 And KeyAscii <= 40 Then
        'Leave it...numeric, backspace and period
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtPort_Change()
    bChanged = bChanged Or letProperty(m_Settings, cPort, txtPort.Text)
End Sub

Private Sub txtPort_GotFocus()
    Highlight txtPort
End Sub

Private Sub UserControl_Initialize()
Dim bActive As Boolean
    bActive = AreThemesActive
    If Not bActive Then 'make it flat - looks better on non-theme enabled boxes
        'note don't change appearance on 3D - since it paints the control differently
        'on XP theme enabled machines
        cmbProtocol.Appearance = 0
    End If
    cmbProtocol.AddItem "TCP"
    cmbProtocol.AddItem "UDP"
    txtIP.Text = "127.0.0.1"
    txtLocalIP.Text = ComputerName
    txtPort.Text = cPortValue: setStyle txtPort.hWnd, esNumeric
    txtLocalPort.Text = cLocalPortValue: setStyle txtLocalPort.hWnd, esNumeric
    txtInit.Text = cInitStringValue
    txtCacheLookup.Text = cCacheMacLookupSecondsValue: setStyle txtCacheLookup.hWnd, esNumeric
    If Not bActive Then FixFlatComboboxes UserControl.Controls, False
End Sub

Private Static Function ByteToHex(bytVal As Byte) As String
  ByteToHex = "00"
  Mid$(ByteToHex, 3 - Len(Hex$(bytVal))) = Hex$(bytVal)
End Function

Private Function ScreenMessage(ByVal Message As String) As String
'PRB: Error Message When an XML Document Contains Low-Order ASCII Characters
'http://support.microsoft.com/?kbid=315580
Dim x As Byte
    If Len(Message) > 0 Then
        For x = 0 To 31
            If InStr(1, Message, Chr$(x), vbBinaryCompare) > 0 Then
                Message = Replace(Message, Chr$(x), "/#*x" & ByteToHex(x) & ";")
            End If
        Next
    End If
    ScreenMessage = Message
    'Log cModuleName, "ScreenMessage", "Message transformed for XML.", GPS_LOG_VERBOSE, , Message, GPS_SOURCE_STRING
End Function
Private Function UnscreenMessage(ByVal Message As String) As String
'PRB: Error Message When an XML Document Contains Low-Order ASCII Characters
'http://support.microsoft.com/?kbid=315580
Dim x As Byte
Dim sTemp As String * 2
    If Len(Message) > 0 Then
        For x = 0 To 31
            If InStr(1, Message, "/#*x", vbBinaryCompare) = 0 Then Exit For
            sTemp = ByteToHex(x)
            Message = Replace(Message, "/#*x" & sTemp & ";", Chr$(x))
        Next
    End If
    UnscreenMessage = Message
    'Log cModuleName, "UnScreenMessage", "Message transformed from XML.", GPS_LOG_VERBOSE, , Message, GPS_SOURCE_STRING
End Function
