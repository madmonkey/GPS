VERSION 5.00
Begin VB.UserControl CustomPropertyPage 
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5505
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
   ScaleHeight     =   4215
   ScaleWidth      =   5505
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      Caption         =   "Time To Live Settings"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   27
      Top             =   2665
      Width           =   5295
      Begin VB.TextBox txtTTLRetryRnd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4440
         MaxLength       =   6
         TabIndex        =   12
         Text            =   "999999"
         ToolTipText     =   "Random retry variance (+/-) for a failed transmission (milliseconds)"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtTTLRetryInt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2760
         MaxLength       =   6
         TabIndex        =   11
         Text            =   "999999"
         ToolTipText     =   "Retry Interval for a failed transmission (milliseconds)"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtTimeToLive 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   960
         MaxLength       =   4
         TabIndex        =   10
         Text            =   "9999"
         ToolTipText     =   "Maximum Time-To-Live for a message being transmitted (seconds)"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "± Rand"
         Height          =   255
         Left            =   3600
         TabIndex        =   30
         Top             =   255
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "RetryInt "
         Height          =   255
         Left            =   1920
         TabIndex        =   29
         Top             =   255
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Max TTL"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   255
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      Caption         =   "KeepAlive Settings"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   24
      Top             =   3370
      Width           =   5295
      Begin VB.TextBox txtKeepAliveFails 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4440
         MaxLength       =   3
         TabIndex        =   15
         Text            =   "999"
         ToolTipText     =   "Maximum unacknowledged keepalive messages before indicating a problem"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtKeepAliveInt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2640
         MaxLength       =   7
         TabIndex        =   14
         Text            =   "9999999"
         ToolTipText     =   "Frequency interval to send keepalive message to host (milliseconds)"
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkKeepAlives 
         Caption         =   "Use KeepAlives"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "MaxFails"
         Height          =   255
         Left            =   3600
         TabIndex        =   26
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Interval"
         Height          =   255
         Left            =   1920
         TabIndex        =   25
         Top             =   255
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Local Settings"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      TabIndex        =   17
      Top             =   1330
      Width           =   5295
      Begin VB.CheckBox chkCompress 
         Caption         =   "Use Compression"
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         ToolTipText     =   "Use compression algorithom for message contents"
         Top             =   960
         Width           =   1815
      End
      Begin VB.CheckBox chkEncrypt 
         Caption         =   "Use Encryption"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Use AES encryption for data transmission"
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtLocalAddr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         MaxLength       =   15
         TabIndex        =   5
         Text            =   "255.255.255.255"
         ToolTipText     =   "Local IP Address to bind to - in order to receive messages (only needed if multiple network cards are present)"
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtLocalPort 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   6
         Text            =   "32767"
         ToolTipText     =   "Local port to receive messages on"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtReceiveBuffer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3360
         MaxLength       =   4
         TabIndex        =   7
         Text            =   "4096"
         ToolTipText     =   "Maximum receive segment size"
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Segment Size"
         Height          =   255
         Left            =   3360
         TabIndex        =   23
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Port"
         Height          =   255
         Left            =   2040
         TabIndex        =   22
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Address"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Remote Settings"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   5295
      Begin VB.CheckBox chkSerializeMessage 
         Caption         =   "Serialize structure"
         Height          =   255
         Left            =   2040
         TabIndex        =   4
         ToolTipText     =   "Serialize message structure, should be unchecked ONLY if using in emulation mode against another Host system."
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtSegmentSize 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3360
         MaxLength       =   4
         TabIndex        =   2
         Text            =   "4096"
         ToolTipText     =   "Maximum transmission segment size"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtRemotePort 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   1
         Text            =   "32767"
         ToolTipText     =   "Remote port to send messages to."
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtRemoteHost 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         MaxLength       =   15
         TabIndex        =   0
         Text            =   "255.255.255.255"
         ToolTipText     =   "Remote IP Address - to send messages to"
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox chkProcessComplete 
         Caption         =   "Process Complete"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Raise process complete event on send"
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Segment Size"
         Height          =   255
         Left            =   3360
         TabIndex        =   20
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Port"
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Address"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1815
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

Private Sub chkCompress_Click()
    bChanged = bChanged Or letProperty(m_Settings, cUseCompression, CStr(CBool(chkCompress.Value)))
End Sub

Private Sub chkEncrypt_Click()
    bChanged = bChanged Or letProperty(m_Settings, cUseEncryption, CStr(CBool(chkEncrypt.Value)))
End Sub

Private Sub chkKeepAlives_Click()
    bChanged = bChanged Or letProperty(m_Settings, cUseKeepAlive, CStr(CBool(chkKeepAlives.Value)))
    txtKeepAliveInt.Enabled = CBool(chkKeepAlives.Value): txtKeepAliveInt.Locked = Not CBool(chkKeepAlives.Value)
    txtKeepAliveFails.Enabled = CBool(chkKeepAlives.Value): txtKeepAliveFails.Locked = Not CBool(chkKeepAlives.Value)
End Sub

Private Sub chkProcessComplete_Click()
    bChanged = bChanged Or letProperty(m_Settings, cProcessOnSend, CStr(CBool(chkProcessComplete.Value)))
End Sub

Private Sub chkSerializeMessage_Click()
    bChanged = bChanged Or letProperty(m_Settings, cSerializeMsg, CStr(CBool(chkSerializeMessage.Value)))
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
        txtRemoteHost.Text = getProperty(m_Settings, cRemoteAddr, cRemoteAddrValue)
        txtRemotePort.Text = getProperty(m_Settings, cRemotePort, cRemotePortValue)
        txtSegmentSize.Text = getProperty(m_Settings, cSegmentSize, cSegmentSizeValue)
        chkProcessComplete.Value = Abs(CBool(getProperty(m_Settings, cProcessOnSend, cProcessOnSendValue)))
        txtLocalAddr.Text = getProperty(m_Settings, cLocalAddr, cLocalAddrValue)
        txtLocalPort.Text = getProperty(m_Settings, cLocalPort, cLocalPortValue)
        txtReceiveBuffer.Text = getProperty(m_Settings, cReceiveBufferSize, cReceiveBufferSizeValue)
        chkEncrypt.Value = Abs(CBool(getProperty(m_Settings, cUseEncryption, cUseEncryptionValue)))
        chkCompress.Value = Abs(CBool(getProperty(m_Settings, cUseCompression, cUseCompressionValue)))
        txtTimeToLive.Text = getProperty(m_Settings, cTimeToLive, cTimeToLiveValue)
        txtTTLRetryInt.Text = getProperty(m_Settings, cRetryInterval, cRetryIntervalValue)
        txtTTLRetryRnd.Text = getProperty(m_Settings, cRetryRandomInterval, cRetryRandomIntervalValue)
        chkKeepAlives.Value = Abs(CBool(getProperty(m_Settings, cUseKeepAlive, cUseKeepAliveValue)))
        txtKeepAliveInt.Text = getProperty(m_Settings, cKeepAliveInterval, cKeepAliveIntervalValue)
        txtKeepAliveFails.Text = getProperty(m_Settings, cMaxKeepAliveFailures, cMaxKeepAliveFailuresValue)
        txtKeepAliveInt.Enabled = CBool(chkKeepAlives.Value): txtKeepAliveInt.Locked = Not CBool(chkKeepAlives.Value)
        txtKeepAliveFails.Enabled = CBool(chkKeepAlives.Value): txtKeepAliveFails.Locked = Not CBool(chkKeepAlives.Value)
        chkSerializeMessage.Value = Abs(CBool(getProperty(m_Settings, cSerializeMsg, cSerializeMsgValue)))
    End If
End Function

Private Function IsValidIP(ByVal ipAddress As String) As Boolean
Dim regEx As VBScript_RegExp_55.RegExp
    Set regEx = New VBScript_RegExp_55.RegExp
    With regEx
        .IgnoreCase = True
        .MultiLine = False
        .Pattern = "^(25[0-5]|2[0-4]\d|[0-1]?\d?\d)(\.(25[0-5]|2[0-4]\d|[0-1]?\d?\d)){3}$"
        IsValidIP = .Test(ipAddress)
    End With
    Set regEx = Nothing
End Function

Private Function IsValidNumber(ByVal textVal As Variant, Optional ByVal Max As Long = 32767&, Optional ByVal Min As Long = 0&) As Boolean
    If IsNumeric(textVal) Then
        IsValidNumber = (textVal >= Min) And (textVal <= Max)
    End If
End Function

Private Sub txtKeepAliveFails_Change()
    bChanged = bChanged Or letProperty(m_Settings, cMaxKeepAliveFailures, txtKeepAliveFails.Text)
End Sub

Private Sub txtKeepAliveFails_GotFocus()
    Highlight txtKeepAliveFails
End Sub

Private Sub txtKeepAliveFails_Validate(Cancel As Boolean)
    Cancel = Not IsValidNumber(txtKeepAliveFails.Text, 999, 1)
End Sub

Private Sub txtKeepAliveInt_Change()
    bChanged = bChanged Or letProperty(m_Settings, cKeepAliveInterval, txtKeepAliveInt.Text)
End Sub

Private Sub txtKeepAliveInt_GotFocus()
    Highlight txtKeepAliveInt
End Sub

Private Sub txtKeepAliveInt_Validate(Cancel As Boolean)
    Cancel = Not IsValidNumber(txtKeepAliveInt.Text, 9999999, 1)
End Sub

Private Sub txtLocalAddr_Change()
    bChanged = bChanged Or letProperty(m_Settings, cLocalAddr, txtLocalAddr.Text)
End Sub

Private Sub txtLocalAddr_GotFocus()
    Highlight txtLocalAddr
End Sub

Private Sub txtLocalAddr_Validate(Cancel As Boolean)
    Cancel = Not IsValidIP(txtLocalAddr.Text)
End Sub

Private Sub txtLocalPort_Change()
    bChanged = bChanged Or letProperty(m_Settings, cLocalPort, txtLocalPort.Text)
End Sub

Private Sub txtLocalPort_GotFocus()
    Highlight txtLocalPort
End Sub

Private Sub txtLocalPort_Validate(Cancel As Boolean)
    Cancel = Not IsValidNumber(txtLocalPort.Text)
End Sub

Private Sub txtReceiveBuffer_Change()
    bChanged = bChanged Or letProperty(m_Settings, cReceiveBufferSize, txtReceiveBuffer.Text)
End Sub

Private Sub txtReceiveBuffer_GotFocus()
    Highlight txtReceiveBuffer
End Sub

Private Sub txtReceiveBuffer_Validate(Cancel As Boolean)
    Cancel = Not IsValidNumber(txtReceiveBuffer.Text, 4096&, 1&)
End Sub

Private Sub txtRemoteHost_Change()
    bChanged = bChanged Or letProperty(m_Settings, cRemoteAddr, txtRemoteHost.Text)
End Sub

Private Sub txtRemoteHost_GotFocus()
    Highlight txtRemoteHost
End Sub

Private Sub txtRemoteHost_Validate(Cancel As Boolean)
    Cancel = Not IsValidIP(txtRemoteHost.Text)
End Sub

Private Sub txtRemotePort_Change()
    bChanged = bChanged Or letProperty(m_Settings, cRemotePort, txtRemotePort.Text)
End Sub

Private Sub txtRemotePort_GotFocus()
    Highlight txtRemotePort
End Sub

Private Sub txtRemotePort_Validate(Cancel As Boolean)
    Cancel = Not IsValidNumber(txtRemotePort.Text)
End Sub

Private Sub txtSegmentSize_Change()
    bChanged = bChanged Or letProperty(m_Settings, cSegmentSize, txtSegmentSize.Text)
End Sub

Private Sub txtSegmentSize_GotFocus()
    Highlight txtSegmentSize
End Sub

Private Sub txtSegmentSize_Validate(Cancel As Boolean)
    Cancel = Not IsValidNumber(txtSegmentSize.Text, 4096&, 1&)
End Sub

Private Sub txtTimeToLive_Change()
    bChanged = bChanged Or letProperty(m_Settings, cTimeToLive, txtTimeToLive.Text)
End Sub

Private Sub txtTimeToLive_GotFocus()
    Highlight txtTimeToLive
End Sub

Private Sub txtTimeToLive_Validate(Cancel As Boolean)
    Cancel = Not IsValidNumber(txtTimeToLive.Text, 9999&, 0&)
End Sub

Private Sub txtTTLRetryInt_Change()
    bChanged = bChanged Or letProperty(m_Settings, cRetryInterval, txtTTLRetryInt.Text)
End Sub

Private Sub txtTTLRetryInt_GotFocus()
    Highlight txtTTLRetryInt
End Sub

Private Sub txtTTLRetryInt_Validate(Cancel As Boolean)
    If txtTTLRetryInt.Text = vbNullString Then
        txtTTLRetryInt.Text = 0
    End If
End Sub

Private Sub txtTTLRetryRnd_Change()
    bChanged = bChanged Or letProperty(m_Settings, cRetryRandomInterval, txtTTLRetryRnd.Text)
End Sub

Private Sub txtTTLRetryRnd_GotFocus()
    Highlight txtTTLRetryRnd
End Sub

Private Sub txtTTLRetryRnd_Validate(Cancel As Boolean)
    If txtTTLRetryRnd.Text = vbNullString Then
        txtTTLRetryRnd.Text = 0
    End If
End Sub

Private Sub UserControl_Initialize()
    setStyle txtRemotePort.hWnd, esNumeric
    setStyle txtSegmentSize.hWnd, esNumeric
    setStyle txtLocalPort.hWnd, esNumeric
    setStyle txtReceiveBuffer.hWnd, esNumeric
    setStyle txtKeepAliveInt.hWnd, esNumeric
    setStyle txtTimeToLive.hWnd, esNumeric
    setStyle txtTTLRetryInt.hWnd, esNumeric
    setStyle txtTTLRetryRnd.hWnd, esNumeric
    setStyle txtKeepAliveInt.hWnd, esNumeric
    setStyle txtKeepAliveFails.hWnd, esNumeric
End Sub
