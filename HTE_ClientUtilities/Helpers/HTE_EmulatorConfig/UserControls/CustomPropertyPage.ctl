VERSION 5.00
Begin VB.UserControl CustomPropertyPage 
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
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
   ScaleHeight     =   2160
   ScaleWidth      =   5790
   ToolboxBitmap   =   "CustomPropertyPage.ctx":0000
   Begin VB.CheckBox chkExtended 
      Appearance      =   0  'Flat
      Caption         =   "Extended Character Set"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Source data exists in hex representation because of extended characterset"
      Top             =   1320
      Width           =   4455
   End
   Begin VB.ComboBox cmbType 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "Combo1"
      ToolTipText     =   "Select message type of emulation"
      Top             =   120
      Width           =   2175
   End
   Begin VB.CheckBox chkLoopback 
      Appearance      =   0  'Flat
      Caption         =   "Loopback"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Continue playback from beginning"
      Top             =   1080
      Width           =   4455
   End
   Begin VB.TextBox txtInterval 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2880
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "5"
      ToolTipText     =   "Interval (in seconds) used for Interval or as default for Realtime"
      Top             =   600
      Width           =   735
   End
   Begin VB.OptionButton optInterval 
      Appearance      =   0  'Flat
      Caption         =   "Interval"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   5
      ToolTipText     =   "Fire playback at specific intervals"
      Top             =   720
      Value           =   -1  'True
      Width           =   3735
   End
   Begin VB.OptionButton optInterval 
      Appearance      =   0  'Flat
      Caption         =   "Realtime"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Infer time from datetime stamp of playback file"
      Top             =   720
      Width           =   4455
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblStatus 
      Caption         =   "Unassigned"
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   8
      Top             =   240
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
Dim styleSheet As String
Private Const cFileDescriptor = "emu"
'Private Const cPlayBackNode = "PLAYBACKNODE"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function getTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private m_localTypes() As HTE_GPS.GPSConfiguration

Private Function getTemporaryFile() As String
Dim sTemp As String
Dim sReturn As String
Const FILE_ATTRIBUTE_TEMPORARY = &H100
    sTemp = String(260, 0)
    getTempFileName Environ("TEMP"), cFileDescriptor, 0, sTemp
    sTemp = Left$(sTemp, InStr(1, sTemp, Chr$(0)) - 1)
    SetFileAttributes sTemp, FILE_ATTRIBUTE_TEMPORARY
    sReturn = Left$(sTemp, InStrRev(sTemp, ".")) & "xml"
    getTemporaryFile = sReturn
End Function

Private Sub chkExtended_Click()
    bChanged = bChanged Or letProperty(m_Settings, cSourceType, CBool(-chkExtended.Value))
End Sub

Private Sub chkLoopback_Click()
    bChanged = bChanged Or letProperty(m_Settings, cLoopbackType, CBool(-chkLoopback.Value))
End Sub

Private Sub cmbType_Click()
    bChanged = bChanged Or letProperty(m_Settings, cMessageType, cmbType.ItemData(cmbType.ListIndex))
End Sub

Private Sub optInterval_Click(Index As Integer)
    bChanged = bChanged Or letProperty(m_Settings, cPlaybackIntervalType, Index)
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
    m_Settings.async = False
    loadLocalSettings = m_Settings.loadXML(sXML)
    If loadLocalSettings Then
        optInterval(getProperty(m_Settings, cPlaybackIntervalType, cPlaybackIntervalTypeValue)).Value = True
        txtInterval.Text = getProperty(m_Settings, cPlaybackInterval, cPlaybackIntervalValue)
        chkLoopback.Value = Abs(CBool(getProperty(m_Settings, cLoopbackType, cLoopbackTypeValue)))
        chkExtended.Value = Abs(CBool(getProperty(m_Settings, cSourceType, cSourceTypeValue)))
        m_localTypes() = SupportedMessages(m_Settings)
        styleSheet = getProperty(m_Settings, cPlayback, cPlaybackValue)
        If validateStylesheet(styleSheet) Then
            lblStatus(0).Caption = "Assigned"
            lblStatus(0).ForeColor = vbBlack
        End If
        initTypesCombo getProperty(m_Settings, cMessageType, cMessageTypeValue)
    End If
End Function

Private Sub initTypesCombo(Optional ByVal currentMsgType As Long = 0)
Dim i As Long
    cmbType.Clear
    If IsValidArray(m_localTypes) Then
        For i = LBound(m_localTypes) To UBound(m_localTypes)
            cmbType.AddItem m_localTypes(i).Desc
            cmbType.ItemData(cmbType.NewIndex) = m_localTypes(i).GPSType
            If currentMsgType = m_localTypes(i).GPSType Then
                cmbType.ListIndex = cmbType.NewIndex
            End If
        Next
    End If
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
Dim oObj As HTE_SystemUtility.enhCommonDialog
Dim fileName As String, strFile As String, strCompare As String
Dim fh As Long
    
    Screen.MousePointer = vbHourglass
    Set oObj = New HTE_SystemUtility.enhCommonDialog
    strCompare = styleSheet
    If oObj.VBGetOpenFileName(fileName, "Select Playback File", True, , , True, "Playback File (*.pbf)|*.pbf", , App.Path, , , UserControl.hWnd) Then
        fh = FreeFile
        Open fileName For Binary Access Read Lock Read As #fh
        strFile = Input(LOF(fh), fh)
        'bizarre extraneous input from the logparser output from MS - shield to ensure no "bad" characters before XML root
        strFile = StrConv(Mid$(strFile, InStr(strFile, "<")), vbFromUnicode)
        Close #fh
        If validateStylesheet(strFile) Then
            styleSheet = strFile
            lblStatus(Index).Caption = "Assigned"
            lblStatus(Index).ForeColor = vbBlack
            addStyleSheet (Index)
        Else
            lblStatus(Index).Caption = "Unassigned"
            lblStatus(Index).ForeColor = vbRed
            styleSheet = vbNullString
            removeStylesheet (Index)
        End If
    End If
    
    bChanged = bChanged Or StrComp(strFile, strCompare, vbTextCompare) <> 0
    Screen.MousePointer = vbDefault
    Set oObj = Nothing
    
End Sub
Private Sub cmdPreview_Click(Index As Integer)
Dim fPreview As frmPreview
Dim sURL As String
Dim sWorkfile As String
Dim oXML As MSXML2.DOMDocument30
Const cBlankPage = "about:blank"
    
    Set fPreview = New frmPreview
    Load fPreview
    fPreview.Move UserControl.Parent.Left, UserControl.Parent.Top, UserControl.Parent.Width, UserControl.Parent.Height
    sWorkfile = styleSheet
    
    If sWorkfile <> vbNullString Then
        Set oXML = New MSXML2.DOMDocument
        oXML.async = False
        If oXML.loadXML(sWorkfile) Then
            sWorkfile = oXML.xml
            sURL = CreatePreview(sWorkfile)
        Else
            sURL = cBlankPage
        End If
    Else
        sURL = cBlankPage
    End If
    fPreview.txtPreview.Navigate sURL
    fPreview.Show vbModal
    If sURL <> cBlankPage Then
        If Dir(sURL) <> vbNullString Then Kill sURL
    End If
    Unload fPreview
    Set fPreview = Nothing
    
End Sub
Private Function CreatePreview(ByVal sXML As String) As String
Dim fh As Long
Dim fileName As String

    fileName = getTemporaryFile
    fh = FreeFile
    Open fileName For Binary Access Write As #fh
    Put #fh, , sXML
    Close #fh
    CreatePreview = fileName
    
End Function

Private Sub cmdClear_Click(Index As Integer)
Dim strCompare As String
    
    strCompare = styleSheet
    removeStylesheet (Index) 'This was never added in the clear for some reason....
    lblStatus(Index).Caption = "Unassigned"
    lblStatus(Index).ForeColor = vbBlack
    styleSheet = vbNullString
    bChanged = bChanged Or StrComp(vbNullString, strCompare, vbTextCompare) <> 0
    
End Sub

Private Function validateStylesheet(ByRef sXSL As String) As Boolean
Dim oDOM As MSXML2.DOMDocument30
Dim oNEWDOM As MSXML2.DOMDocument30
Dim iAttribute As MSXML2.IXMLDOMNode
Dim bReturn As Boolean, bHasTimeStamp As Boolean
Dim i As Long
On Local Error Resume Next
    Set oDOM = New MSXML2.DOMDocument30
    oDOM.async = False
    bReturn = oDOM.loadXML(sXSL)
    If bReturn Then
        bReturn = bReturn And StrComp(oDOM.documentElement.nodeName, "ROOT", vbTextCompare) = 0
        If bReturn Then
            'check to make sure created by parser
            bReturn = bReturn And (oDOM.documentElement.Attributes.length > 0)
            If bReturn Then
                Set iAttribute = oDOM.documentElement.Attributes.getNamedItem("CREATED_BY")
                bReturn = bReturn And Not iAttribute Is Nothing
                If bReturn Then
                    bReturn = bReturn And InStr(1, iAttribute.nodeTypedValue, "Microsoft Log Parser", vbTextCompare) > 0
                    If bReturn Then
                        'now check to make sure we have good elements
                        bReturn = bReturn And oDOM.documentElement.hasChildNodes
                        If bReturn Then
                            bReturn = bReturn And StrComp(oDOM.documentElement.firstChild.nodeName, "ROW", vbTextCompare) = 0
                            If bReturn Then
                                bReturn = bReturn And oDOM.documentElement.firstChild.hasChildNodes
                                If bReturn Then
                                    For i = 0 To oDOM.documentElement.firstChild.childNodes.length - 1
                                        Select Case UCase$(oDOM.documentElement.firstChild.childNodes.Item(i).nodeName)
                                            Case "SOURCE"
                                                bReturn = True
                                            Case "DATETIMESTAMP"
                                                'we can enable the realtime option button
                                                bHasTimeStamp = IsDate(oDOM.documentElement.firstChild.childNodes.Item(i).nodeTypedValue)
                                        End Select
                                    Next
                                    If bReturn Then
                                        Set iAttribute = oDOM.Attributes.getNamedItem("encoding")
                                        If Not iAttribute Is Nothing Then
                                            oDOM.Attributes.removeNamedItem "encoding"
                                            sXSL = oDOM.xml
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    optInterval(0).Enabled = bHasTimeStamp
    validateStylesheet = bReturn
    Set oDOM = Nothing
End Function

Private Function removeStylesheet(ByVal Index As Integer)
Dim iNodes As MSXML2.IXMLDOMNodeList, iNode As MSXML2.IXMLDOMNode ', iAttribute As MSXML2.IXMLDOMNode
Dim i As Long
    If Not m_Settings Is Nothing Then
        Set iNodes = m_Settings.documentElement.getElementsByTagName(cPlayback)
        If Not iNodes Is Nothing Then
            For i = 0 To iNodes.length - 1
                Set iNode = iNodes.Item(i)
                'Verify based on what we know, not what we think we know
                If Not iNode Is Nothing Then 'And Not iAttribute Is Nothing Then
                    m_Settings.documentElement.removeChild iNode
                    bChanged = True
                End If
            Next
        End If
    End If
End Function

Private Function addStyleSheet(ByVal Index As Integer)
Dim iNodes As MSXML2.IXMLDOMNodeList
Dim iNode As MSXML2.IXMLDOMNode, iAttribute As MSXML2.IXMLDOMNode
Dim newAtt As IXMLDOMAttribute, namedNodeMap As IXMLDOMNamedNodeMap
Dim sBefore As String
Dim i As Long
    If Not m_Settings Is Nothing Then
        sBefore = m_Settings.xml
        Set iNodes = m_Settings.documentElement.getElementsByTagName(cPlayback)
        If iNodes Is Nothing Then
            Set iNode = m_Settings.createElement(cPlayback)
            iNode.nodeTypedValue = styleSheet
        Else
            For i = 0 To iNodes.length - 1
                If iNodes.Item(i).parentNode.nodeName <> "TYPES" Then
                    Set iNode = m_Settings.getElementsByTagName(cPlayback).nextNode
                    Exit For
                End If
            Next
            If iNode Is Nothing Then
                Set iNode = m_Settings.createElement(cPlayback)
                If Not iNode Is Nothing Then iNode.nodeTypedValue = styleSheet
            Else
                iNode.nodeTypedValue = styleSheet
            End If
        End If
        m_Settings.documentElement.appendChild iNode
        bChanged = bChanged Or StrComp(m_Settings.xml, sBefore, vbTextCompare) <> 0
    End If
End Function

'Private Function ScreenMessage(ByVal Message As String) As String
''PRB: Error Message When an XML Document Contains Low-Order ASCII Characters
''http://support.microsoft.com/?kbid=315580
'Dim x As Byte
'    If Len(Message) > 0 Then
'        For x = 0 To 31
'            If InStr(1, Message, Chr$(x), vbBinaryCompare) > 0 Then
'                Message = Replace(Message, Chr$(x), "/#*x" & ByteToHex(x) & ";")
'            End If
'        Next
'    End If
'    ScreenMessage = Message
'End Function
'Private Function UnscreenMessage(ByVal Message As String) As String
''PRB: Error Message When an XML Document Contains Low-Order ASCII Characters
''http://support.microsoft.com/?kbid=315580
'Dim x As Byte
'Dim sTemp As String * 2
'    If Len(Message) > 0 Then
'        For x = 0 To 31
'            If InStr(1, Message, "/#*x", vbBinaryCompare) = 0 Then Exit For
'            sTemp = ByteToHex(x)
'            Message = Replace(Message, "/#*x" & sTemp & ";", Chr$(x))
'        Next
'    End If
'    UnscreenMessage = Message
'End Function
'
'Private Static Function ByteToHex(bytVal As Byte) As String
'    ByteToHex = "00"
'    Mid$(ByteToHex, 3 - Len(Hex$(bytVal))) = Hex$(bytVal)
'End Function

Private Sub txtInterval_Change()
    If IsNumeric(txtInterval.Text) Then
        bChanged = bChanged Or letProperty(m_Settings, cPlaybackInterval, txtInterval.Text)
    End If
End Sub

Private Sub txtInterval_GotFocus()
    Highlight txtInterval
End Sub

