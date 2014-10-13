VERSION 5.00
Object = "{F9B2917F-12A4-4C2D-9B9B-4E1EDE8126BB}#1.1#0"; "HTE_FlowControl.ocx"
Begin VB.UserControl CustomPropertyPage 
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6705
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
   ScaleHeight     =   4545
   ScaleWidth      =   6705
   ToolboxBitmap   =   "ppComPage.ctx":0000
   Begin VB.CheckBox chkProcessOnSend 
      Appearance      =   0  'Flat
      Caption         =   "Complete On Send"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   12
      ToolTipText     =   "Indicates that the component should bubble-up messages sent to device -> should only be checked if used as a repeater!"
      Top             =   2640
      Width           =   2175
   End
   Begin HTE_FlowControl.FlowControl FlowControl1 
      Height          =   1410
      Left            =   120
      TabIndex        =   21
      Top             =   3000
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   2487
   End
   Begin VB.ComboBox cmbPort 
      Height          =   315
      ItemData        =   "ppComPage.ctx":0312
      Left            =   2280
      List            =   "ppComPage.ctx":034D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtSettings 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txtInputLength 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2280
      TabIndex        =   7
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtRThreshold 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2280
      TabIndex        =   6
      Top             =   2280
      Width           =   2175
   End
   Begin VB.ComboBox cmbInputMode 
      Height          =   315
      ItemData        =   "ppComPage.ctx":03C8
      Left            =   2280
      List            =   "ppComPage.ctx":03D2
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtInitString 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2280
      TabIndex        =   4
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CheckBox chkDTR 
      Appearance      =   0  'Flat
      Caption         =   "DTR Enable"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.CheckBox chkEOF 
      Appearance      =   0  'Flat
      Caption         =   "EOF Enable"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   480
      Width           =   1695
   End
   Begin VB.CheckBox chkRTS 
      Appearance      =   0  'Flat
      Caption         =   "RTS Enable"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   840
      Width           =   1695
   End
   Begin VB.CheckBox chkNull 
      Appearance      =   0  'Flat
      Caption         =   "Null Discard"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtInBufferSize 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2280
      TabIndex        =   5
      Top             =   1920
      Width           =   2175
   End
   Begin VB.ComboBox cmbHandShake 
      Height          =   315
      ItemData        =   "ppComPage.ctx":03EE
      Left            =   2280
      List            =   "ppComPage.ctx":03FE
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "ComPort:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Settings:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Input Length:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "RThreshold:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Input Mode:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Init String:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "InBuffer Size:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "HandShaking:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   1215
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

Private Sub chkProcessOnSend_Click()
    bChanged = bChanged Or letProperty(m_Settings, cProcessOnSend, CStr(CBool(-chkProcessOnSend.Value)))
End Sub

'ADDED FOR DEVICE FLOW CONTROL
Private Sub FlowControl1_DurationChange()
    bChanged = bChanged Or letProperty(m_Settings, cDurationInterval, FlowControl1.DurationInterval)
End Sub

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
'ADDED FOR DEVICE FLOW CONTROL

Private Sub chkDTR_Click()
    bChanged = bChanged Or letProperty(m_Settings, cDTREnable, -chkDTR.Value)
End Sub

Private Sub chkEOF_Click()
    bChanged = bChanged Or letProperty(m_Settings, cEOFEnable, -chkEOF.Value)
End Sub

Private Sub chkNull_Click()
    bChanged = bChanged Or letProperty(m_Settings, cNullDiscard, -chkNull.Value)
End Sub

Private Sub chkRTS_Click()
    bChanged = bChanged Or letProperty(m_Settings, cRTSEnable, -chkRTS.Value)
End Sub

Private Sub cmbHandShake_Click()
     bChanged = bChanged Or letProperty(m_Settings, cHandshaking, cmbHandShake.ListIndex)
End Sub

Private Sub cmbHandShake_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        bChanged = bChanged Or letProperty(m_Settings, cHandshaking, cmbHandShake.ListIndex)
    End If
End Sub

Private Sub cmbInputMode_Click()
    bChanged = bChanged Or letProperty(m_Settings, cInputMode, cmbInputMode.ListIndex)
End Sub

Private Sub cmbInputMode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        bChanged = bChanged Or letProperty(m_Settings, cInputMode, cmbInputMode.ListIndex)
    End If
End Sub

Private Sub cmbPort_Click()
    bChanged = bChanged Or letProperty(m_Settings, cComm, cmbPort.ItemData(cmbPort.ListIndex))
End Sub

Private Sub cmbPort_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        bChanged = bChanged Or letProperty(m_Settings, cComm, cmbPort.ItemData(cmbPort.ListIndex))
    End If
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
Dim thisPort As Long
    Set m_Settings = New MSXML2.DOMDocument30
    loadLocalSettings = m_Settings.loadXML(sXML)
    If loadLocalSettings Then
        thisPort = CLng(getProperty(m_Settings, cComm, cCommValue))
        If thisPort > 0 And thisPort <= cMaxNumberOfPorts Then
            cmbPort.ListIndex = CLng(getProperty(m_Settings, cComm, cCommValue) - 1)
        End If
        cmbHandShake.ListIndex = getProperty(m_Settings, cHandshaking, cHandshakingValue)
        cmbInputMode.ListIndex = getProperty(m_Settings, cInputMode, cInputModeValue)
        txtSettings.Text = getProperty(m_Settings, cSettings, cSettingsValue)
        txtInitString.Text = getProperty(m_Settings, cInitString, cInitStringValue)
        txtInBufferSize.Text = getProperty(m_Settings, cInBufferSize, cInBufferSizeValue)
        txtRThreshold.Text = getProperty(m_Settings, cRThresh, cRThreshValue)
        txtInputLength.Text = getProperty(m_Settings, cInputLen, cInputLenValue)
        chkDTR.Value = Abs(getProperty(m_Settings, cDTREnable, cDTREnableValue))
        chkEOF.Value = Abs(getProperty(m_Settings, cEOFEnable, cEOFEnableValue))
        chkRTS.Value = Abs(getProperty(m_Settings, cRTSEnable, cRTSEnableValue))
        chkNull.Value = Abs(getProperty(m_Settings, cNullDiscard, cNullDiscardValue))
        chkProcessOnSend.Value = Abs(CBool(getProperty(m_Settings, cProcessOnSend, cProcessOnSendValue)))
        'ADDED FOR DEVICE FLOW CONTROL
        FlowControl1.UseSoftwareReporting = CBool(getProperty(m_Settings, cManualPoll, cManualPollValue))
        FlowControl1.ReportingSentence = UnscreenMessage(getProperty(m_Settings, cManualPollString, cManualPollStringValue))
        FlowControl1.ReportingInterval = getProperty(m_Settings, cManualPollInterval, cManualPollIntervalValue)
        FlowControl1.RelayInterval = CLng(getProperty(m_Settings, cRelayInterval, cRelayIntervalValue))
        FlowControl1.DurationInterval = CLng(getProperty(m_Settings, cDurationInterval, cDurationIntervalValue))
        'ADDED FOR DEVICE FLOW CONTROL
    End If
End Function

Private Sub txtInBufferSize_Change()
    bChanged = bChanged Or letProperty(m_Settings, cInBufferSize, txtInBufferSize.Text)
End Sub

Private Sub txtInBufferSize_GotFocus()
    Highlight txtInBufferSize
End Sub

Private Sub txtInitString_Change()
    bChanged = bChanged Or letProperty(m_Settings, cInitString, txtInitString.Text)
End Sub

Private Sub txtInitString_GotFocus()
    Highlight txtInitString
End Sub

Private Sub txtInputLength_Change()
    bChanged = bChanged Or letProperty(m_Settings, cInputLen, txtInputLength.Text)
End Sub

Private Sub txtInputLength_GotFocus()
    Highlight txtInputLength
End Sub

Private Sub txtRThreshold_Change()
    bChanged = bChanged Or letProperty(m_Settings, cRThresh, txtRThreshold.Text)
End Sub

Private Sub txtRThreshold_GotFocus()
    Highlight txtRThreshold
End Sub

Private Sub txtSettings_Change()
    bChanged = bChanged Or letProperty(m_Settings, cSettings, txtSettings.Text)
End Sub

Private Sub txtSettings_GotFocus()
    Highlight txtSettings
End Sub

Private Sub UserControl_Initialize()
Dim bActive As Boolean
    bActive = AreThemesActive
    If Not bActive Then 'make it flat - looks better on non-theme enabled boxes
        'note don't change appearance on 3D - since it paints the control differently
        'on XP theme enabled machines
        cmbPort.Appearance = 0
        cmbHandShake.Appearance = 0
        cmbInputMode.Appearance = 0
    End If
    LoadCOMPorts
    cmbPort.ListIndex = 0
    cmbHandShake.ListIndex = 0
    cmbInputMode.ListIndex = 0
    txtSettings.Text = cSettingsValue
    txtInitString.Text = cInitStringValue
    txtInBufferSize.Text = cInBufferSizeValue: setStyle txtInBufferSize.hWnd, esNumeric
    txtRThreshold.Text = cRThreshValue: setStyle txtRThreshold.hWnd, esNumeric
    txtInputLength.Text = cInputLenValue: setStyle txtInputLength.hWnd, esNumeric
    chkDTR.Value = Abs(cDTREnableValue)
    chkEOF.Value = Abs(cEOFEnableValue)
    chkRTS.Value = Abs(cRTSEnableValue)
    chkNull.Value = Abs(cNullDiscardValue)
    FlowControl1.RelayInterval = 0
    FlowControl1.DurationInterval = 0
    If Not bActive Then FixFlatComboboxes UserControl.Controls, False
End Sub

Private Function ScreenMessage(ByVal Message As String) As String
'PRB: Error Message When an XML Document Contains Low-Order ASCII Characters
'http://support.microsoft.com/?kbid=315580
Dim X As Byte
    If Len(Message) > 0 Then
        For X = 0 To 31
            If InStr(1, Message, Chr$(X), vbBinaryCompare) > 0 Then
                Message = Replace(Message, Chr$(X), "/#*x" & ByteToHex(X) & ";")
            End If
        Next
    End If
    ScreenMessage = Message
End Function
Private Function UnscreenMessage(ByVal Message As String) As String
'PRB: Error Message When an XML Document Contains Low-Order ASCII Characters
'http://support.microsoft.com/?kbid=315580
Dim X As Byte
Dim sTemp As String * 2
    If Len(Message) > 0 Then
        For X = 0 To 31
            If InStr(1, Message, "/#*x", vbBinaryCompare) = 0 Then Exit For
            sTemp = ByteToHex(X)
            Message = Replace(Message, "/#*x" & sTemp & ";", Chr$(X))
        Next
    End If
    UnscreenMessage = Message
End Function

Private Static Function ByteToHex(bytVal As Byte) As String
    ByteToHex = "00"
    Mid$(ByteToHex, 3 - Len(Hex$(bytVal))) = Hex$(bytVal)
End Function

Private Sub LoadCOMPorts()
Dim i As Long
    With cmbPort
        .Clear
        For i = 1 To cMaxNumberOfPorts
            .AddItem "COM:" & CStr(i)
            .ItemData(.NewIndex) = i
        Next
    End With
End Sub
