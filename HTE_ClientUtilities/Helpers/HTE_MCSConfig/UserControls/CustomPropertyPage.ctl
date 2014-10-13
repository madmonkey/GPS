VERSION 5.00
Object = "{F9B2917F-12A4-4C2D-9B9B-4E1EDE8126BB}#1.1#0"; "HTE_FlowControl.ocx"
Begin VB.UserControl CustomPropertyPage 
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6540
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
   ScaleHeight     =   4425
   ScaleWidth      =   6540
   ToolboxBitmap   =   "CustomPropertyPage.ctx":0000
   Begin HTE_FlowControl.FlowControl FlowControl1 
      Height          =   1410
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   2487
   End
   Begin VB.ComboBox cmbCAD 
      Height          =   315
      Left            =   1920
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox txtRcvTags 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox txtRcvTrans 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox txtMTS 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1080
      Width           =   2895
   End
   Begin VB.ComboBox cmbSession 
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox txtDest 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label6 
      Caption         =   "CAD Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Received Tags:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Rcv Transaction(s):"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Send Transaction:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "MCS Session:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Destination ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
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
Dim oSettings As cSettings

Private Sub cmbCAD_Click()
    bChanged = bChanged Or letProperty(m_Settings, cCADType, cmbCAD.List(cmbCAD.ListIndex))
End Sub

Private Sub cmbCAD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        bChanged = bChanged Or letProperty(m_Settings, cCADType, cmbCAD.List(cmbCAD.ListIndex))
    End If
End Sub

Private Sub cmbSession_Click()
    bChanged = bChanged Or letProperty(m_Settings, cSession, cmbSession.List(cmbSession.ListIndex))
End Sub

Private Sub cmbSession_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        bChanged = bChanged Or letProperty(m_Settings, cSession, cmbSession.List(cmbSession.ListIndex))
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
Dim sReturn As String, i As Long
    
    Set m_Settings = New MSXML2.DOMDocument30
    Set oSettings = New cSettings
    loadLocalSettings = m_Settings.loadXML(sXML)
    If loadLocalSettings Then
        sReturn = getProperty(m_Settings, cSession, cSessionValue)
        oSettings.SessionName = sReturn
        For i = 0 To cmbSession.ListCount - 1
            If StrComp(sReturn, cmbSession.List(i), vbTextCompare) = 0 Then
                cmbSession.ListIndex = i
                Exit For
            End If
        Next
        sReturn = getProperty(m_Settings, cCADType, cmbCAD.List(Abs(oSettings.GetValue(ikSessionType) = "ANY_MSG")))
        For i = 0 To cmbCAD.ListCount - 1
            If StrComp(sReturn, cmbCAD.List(i), vbTextCompare) = 0 Then
                cmbCAD.ListIndex = i
                Exit For
            End If
        Next
        txtDest.Text = getProperty(m_Settings, cDest, cDestValue)
        txtMTS.Text = getProperty(m_Settings, cMsgToSend, cMsgToSendValue)
        txtRcvTrans.Text = getProperty(m_Settings, cMsgsToRecv, cMsgsToRecvValue)
        txtRcvTags.Text = getProperty(m_Settings, cTagToRead, cTagToReadValue)
        'ADDED FOR DEVICE FLOWCONTROL
        FlowControl1.UseSoftwareReporting = CBool(getProperty(m_Settings, cManualPoll, cManualPollValue))
        FlowControl1.ReportingSentence = UnscreenMessage(getProperty(m_Settings, cManualPollString, cManualPollStringValue))
        FlowControl1.ReportingInterval = getProperty(m_Settings, cManualPollInterval, cManualPollIntervalValue)
        FlowControl1.RelayInterval = CLng(getProperty(m_Settings, cRelayInterval, cRelayIntervalValue))
        FlowControl1.DurationInterval = CLng(getProperty(m_Settings, cDurationInterval, cDurationIntervalValue))
        'ADDED FOR DEVICE FLOWCONTROL
    End If
    
End Function

Private Sub txtDest_Change()
    bChanged = bChanged Or letProperty(m_Settings, cDest, txtDest.Text)
End Sub

Private Sub txtDest_GotFocus()
    Highlight txtDest
End Sub

Private Sub txtMTS_Change()
    bChanged = bChanged Or letProperty(m_Settings, cMsgToSend, txtMTS.Text)
End Sub

Private Sub txtMTS_GotFocus()
    Highlight txtMTS
End Sub

Private Sub txtRcvTags_Change()
    bChanged = bChanged Or letProperty(m_Settings, cTagToRead, txtRcvTags.Text)
End Sub

Private Sub txtRcvTags_GotFocus()
    Highlight txtRcvTags
End Sub

Private Sub txtRcvTrans_Change()
    bChanged = bChanged Or letProperty(m_Settings, cMsgsToRecv, txtRcvTrans.Text)
End Sub

Private Sub txtRcvTrans_GotFocus()
    Highlight txtRcvTrans
End Sub

Private Sub UserControl_Initialize()
Dim sReturn As String, aReturn() As String
Dim i As Long
Dim bActive As Boolean
    
    bActive = AreThemesActive
    If Not bActive Then 'make it flat - looks better on non-theme enabled boxes
        'note don't change appearance on 3D - since it paints the control differently
        'on XP theme enabled machines
        cmbSession.Appearance = 0
        cmbCAD.Appearance = 0
    End If

    cmbSession.Clear
    Set oSettings = New cSettings
    sReturn = oSettings.GetValue(ikSessions)
    aReturn = Split(sReturn, cSep)
    For i = 0 To UBound(aReturn)
        If Trim$(aReturn(i)) <> vbNullString Then
            cmbSession.AddItem aReturn(i)
            If StrComp(cmbSession.List(cmbSession.NewIndex), cSessionValue, vbTextCompare) Then
                cmbSession.ListIndex = i
            End If
        End If
    Next
    txtDest.Text = cDestValue
    cmbCAD.AddItem "CAD400"
    cmbCAD.AddItem "CADV"
    txtMTS.Text = cMsgToSendValue
    txtRcvTrans.Text = cMsgsToRecvValue
    txtRcvTags.Text = cTagToReadValue
    If Not bActive Then FixFlatComboboxes UserControl.Controls, False
End Sub

'ADDED FOR DEVICE FLOWCONTROL
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
End Function
'ADDED FOR DEVICE FLOWCONTROL
