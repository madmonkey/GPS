VERSION 5.00
Begin VB.UserControl CustomPropertyPage 
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5730
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
   ScaleHeight     =   4410
   ScaleWidth      =   5730
   ToolboxBitmap   =   "CustomPropertyPage.ctx":0000
   Begin VB.CheckBox chkRandomizeQuantity 
      Alignment       =   1  'Right Justify
      Caption         =   "Randomize Message Quantities?"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox txtQuantityMax 
      Height          =   375
      Left            =   2520
      MaxLength       =   5
      TabIndex        =   16
      Text            =   "1"
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtQuantityMin 
      Height          =   375
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   15
      Text            =   "1"
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox txtIntervalMax 
      Height          =   375
      Left            =   2520
      MaxLength       =   5
      TabIndex        =   12
      Text            =   "1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtIntervalMin 
      Height          =   375
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   11
      Text            =   "1"
      Top             =   2760
      Width           =   735
   End
   Begin VB.CheckBox chkRandomizeMsgOrder 
      Alignment       =   1  'Right Justify
      Caption         =   "Randomize Message Order?"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   3135
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
      Index           =   1
      Left            =   2040
      TabIndex        =   6
      Top             =   1920
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
      Index           =   1
      Left            =   2520
      TabIndex        =   7
      Top             =   1920
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
      Index           =   1
      Left            =   3000
      TabIndex        =   8
      Top             =   1920
      Width           =   375
   End
   Begin VB.OptionButton optInterval 
      Caption         =   "Sporadic"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   10
      Top             =   2400
      Width           =   1095
   End
   Begin VB.OptionButton optInterval 
      Caption         =   "Periodic"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   9
      Top             =   2400
      Value           =   -1  'True
      Width           =   975
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
      Left            =   2040
      TabIndex        =   3
      Top             =   1320
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
      Left            =   2520
      TabIndex        =   4
      Top             =   1320
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
      Left            =   3000
      TabIndex        =   5
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox txtDefaultSource 
      Height          =   375
      Left            =   2040
      MaxLength       =   8
      TabIndex        =   2
      Top             =   840
      Width           =   2895
   End
   Begin VB.ComboBox cmbCAD 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   480
      Width           =   2895
   End
   Begin VB.ComboBox cmbSession 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label10 
      Caption         =   "Quantity:"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lblQtyTo 
      Caption         =   "to"
      Height          =   255
      Left            =   2280
      TabIndex        =   27
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "Interval(s):"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Frequency:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblIntTo 
      Caption         =   "to"
      Height          =   255
      Left            =   2280
      TabIndex        =   24
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5640
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label5 
      Caption         =   "Message(s):"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label lblStatus 
      Caption         =   "Unassigned"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   22
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Logon Transaction(s):"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblStatus 
      Caption         =   "Unassigned"
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   20
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Application ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "CAD Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "MCS Session:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   1815
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
Dim styleSheet() As String
Private Const cFileDescriptor = "mcc"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function getTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type WINDOWPLACEMENT
        Length As Long
        flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
End Type
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

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

Private Sub chkRandomizeMsgOrder_Click()
    bChanged = bChanged Or letProperty(m_Settings, cRndOrder, CStr(CBool(-chkRandomizeMsgOrder.Value)))
End Sub

Private Sub chkRandomizeQuantity_Click()
    lblQtyTo.Visible = -chkRandomizeQuantity.Value
    txtQuantityMax.Visible = -chkRandomizeQuantity.Value
    bChanged = bChanged Or letProperty(m_Settings, cRndQuantity, CStr(CBool(-chkRandomizeQuantity.Value)))
End Sub

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

Private Sub optInterval_Click(Index As Integer)
    lblIntTo.Visible = optInterval(1).Value
    txtIntervalMax.Visible = optInterval(1).Value
    bChanged = bChanged Or letProperty(m_Settings, cFrequency, Index)
End Sub

Private Sub optInterval_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
        bChanged = bChanged Or letProperty(m_Settings, cFrequency, Index)
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
Dim iNode As MSXML2.IXMLDOMNode, iAttribute As MSXML2.IXMLDOMNode, iNodes As MSXML2.IXMLDOMNodeList

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
        
        txtDefaultSource.Text = getProperty(m_Settings, cAppID, oSettings.GetValue(ikPWHIPUserID))
        
        If IsValidArray(styleSheet) Then
            For i = LBound(styleSheet) To UBound(styleSheet)
                Set iNode = m_Settings.documentElement.selectSingleNode(IIf(i = 0, cLogonTrans, cMessageTrans))
                If Not iNode Is Nothing Then
                    If validateStylesheet(iNode.firstChild) Then
                        styleSheet(i) = iNode.firstChild.nodeTypedValue
                        lblStatus(i).Caption = "Assigned"
                    End If
                End If
            Next i
        End If
        
        sReturn = getProperty(m_Settings, cFrequency, cFreqDefault)
        
        For i = optInterval.LBound To optInterval.UBound
            If IsNumeric(sReturn) Then
                If CLng(sReturn) = i Then
                    optInterval(i).Value = True
                    Exit For
                End If
            End If
        Next
        
        txtIntervalMin.Text = getProperty(m_Settings, cMinInt, cMinIntDefault)
        txtIntervalMax.Text = getProperty(m_Settings, cMaxInt, cMaxIntDefault)
        
        chkRandomizeMsgOrder.Value = CBool(getProperty(m_Settings, cRndOrder, False))
        
        txtQuantityMin.Text = getProperty(m_Settings, cMinQty, cMinQtyDefault)
        txtQuantityMax.Text = getProperty(m_Settings, cMaxQty, cMaxQtyDefault)
        
        chkRandomizeQuantity.Value = CBool(getProperty(m_Settings, cRndQuantity, False))
        
    End If
    
    
End Function

Private Sub txtDefaultSource_Change()
    bChanged = bChanged Or letProperty(m_Settings, cAppID, txtDefaultSource.Text)
End Sub

Private Sub txtDefaultSource_GotFocus()
    Highlight txtDefaultSource
End Sub

Private Sub txtIntervalMax_Change()
    bChanged = bChanged Or letProperty(m_Settings, cMaxInt, txtIntervalMax.Text)
End Sub

Private Sub txtIntervalMax_GotFocus()
    Highlight txtIntervalMax
End Sub

Private Sub txtIntervalMin_Change()
    bChanged = bChanged Or letProperty(m_Settings, cMinInt, txtIntervalMin.Text)
End Sub

Private Sub txtIntervalMin_GotFocus()
    Highlight txtIntervalMin
End Sub

Private Sub txtQuantityMax_Change()
    bChanged = bChanged Or letProperty(m_Settings, cMaxQty, txtQuantityMax.Text)
End Sub

Private Sub txtQuantityMax_GotFocus()
    Highlight txtQuantityMax
End Sub

Private Sub txtQuantityMin_Change()
    bChanged = bChanged Or letProperty(m_Settings, cMinQty, txtQuantityMin.Text)
End Sub

Private Sub txtQuantityMin_GotFocus()
    Highlight txtQuantityMin
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
    cmbCAD.AddItem "CAD400"
    cmbCAD.AddItem "CADV"
    setStyle txtIntervalMin.hwnd, esNumeric
    setStyle txtIntervalMax.hwnd, esNumeric
    setStyle txtQuantityMin.hwnd, esNumeric
    setStyle txtQuantityMax.hwnd, esNumeric
    If Not bActive Then FixFlatComboboxes UserControl.Controls, False
    If Not IsValidArray(styleSheet) Then ReDim styleSheet(1) '0 - logon, 1 - messages
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
Dim oObj As HTE_SystemUtility.enhCommonDialog
Dim fileName As String, strFile As String, strCompare As String
Dim fh As Long
    If IsValidArray(styleSheet) Then
        If Index <= UBound(styleSheet) And Index >= LBound(styleSheet) Then
            Screen.MousePointer = vbHourglass
            Set oObj = New HTE_SystemUtility.enhCommonDialog
            strCompare = styleSheet(Index)
            If oObj.VBGetOpenFileName(fileName, "Select Transactional Data File", True, , , True, "XML File(*.xml)|*.xml", , App.Path, , , UserControl.hwnd) Then
                fh = FreeFile
                Open fileName For Binary Access Read Lock Read As #fh
                strFile = Input(LOF(fh), fh)
                Close #fh
                If validateStylesheet(strFile) Then
                    styleSheet(Index) = strFile
                    lblStatus(Index).Caption = "Assigned"
                    lblStatus(Index).foreColor = vbBlack
                    addStyleSheet (Index)
                Else
                    lblStatus(Index).Caption = "Unassigned"
                    lblStatus(Index).foreColor = vbRed
                    styleSheet(Index) = vbNullString
                    removeStylesheet (Index)
                End If
            End If
            
            bChanged = bChanged Or StrComp(strFile, strCompare, vbTextCompare) <> 0
            Screen.MousePointer = vbDefault
            Set oObj = Nothing
        End If
    End If
End Sub

Private Function validateStylesheet(ByVal sXSL As String) As Boolean
Dim oDOM As MSXML2.DOMDocument30
On Local Error Resume Next
    Set oDOM = New MSXML2.DOMDocument30
    validateStylesheet = oDOM.loadXML(sXSL)
    Set oDOM = Nothing
End Function

Private Function addStyleSheet(ByVal Index As Integer)
Dim iNodes As MSXML2.IXMLDOMNodeList
Dim iNode As MSXML2.IXMLDOMNode, iAttribute As MSXML2.IXMLDOMNode
Dim newAtt As IXMLDOMAttribute, namedNodeMap As IXMLDOMNamedNodeMap
Dim sBefore As String, sTag As String
Dim i As Long
    If Not m_Settings Is Nothing Then
        sBefore = m_Settings.xml
        Select Case Index
            Case 0
                sTag = cLogonTrans
            Case Else
                sTag = cMessageTrans
        End Select
        Set iNodes = m_Settings.documentElement.getElementsByTagName(sTag)
        If iNodes Is Nothing Then
            Set iNode = m_Settings.createElement(sTag)
            iNode.nodeTypedValue = styleSheet(Index)
        Else
            For i = 0 To iNodes.Length - 1
                If iNodes.Item(i).parentNode.nodeName <> "TYPES" Then
                    Set iNode = m_Settings.getElementsByTagName(sTag).nextNode
                    Exit For
                End If
            Next
            If iNode Is Nothing Then
                Set iNode = m_Settings.createElement(sTag)
                If Not iNode Is Nothing Then iNode.nodeTypedValue = styleSheet(Index)
            Else
                iNode.nodeTypedValue = styleSheet(Index)
            End If
        End If
'''        Set iAttribute = iNode.Attributes.getNamedItem("GPSTYPE")
'''        If iAttribute Is Nothing Then
'''            Set newAtt = m_Settings.createAttribute("GPSTYPE")
'''            newAtt.nodeTypedValue = txtType(Index).Tag
'''            Set namedNodeMap = iNode.Attributes
'''            Set iAttribute = namedNodeMap.setNamedItem(newAtt)
'''        Else
'''            iAttribute.nodeTypedValue = txtType(Index).Tag
'''        End If
        m_Settings.documentElement.appendChild iNode
        bChanged = bChanged Or StrComp(m_Settings.xml, sBefore, vbTextCompare) <> 0
    End If
End Function

Private Function removeStylesheet(ByVal Index As Integer)
Dim iNodes As MSXML2.IXMLDOMNodeList, iNode As MSXML2.IXMLDOMNode, iAttribute As MSXML2.IXMLDOMNode
Dim i As Long
Dim sTag As String

    If Not m_Settings Is Nothing Then
        Select Case Index
            Case 0
                sTag = cLogonTrans
            Case Else
                sTag = cMessageTrans
        End Select
        Set iNodes = m_Settings.documentElement.getElementsByTagName(sTag)
        If Not iNodes Is Nothing Then
            For i = 0 To iNodes.Length - 1
                Set iNode = iNodes.Item(i)
'                Set iAttribute = iNode.Attributes.getNamedItem(cAttrType)
                'Verify based on what we know, not what we think we know
                If Not iNode Is Nothing Then 'And Not iAttribute Is Nothing Then
                    m_Settings.documentElement.removeChild iNode
                    bChanged = True
                End If
            Next
        End If
    End If
End Function

Private Sub cmdPreview_Click(Index As Integer)
Dim fPreview As frmPreview
Dim sURL As String
Dim wp As WINDOWPLACEMENT
Dim rtn As Long
Const cBlankPage = "about:blank"
On Error Resume Next
    If IsValidArray(styleSheet) Then
        If Index <= UBound(styleSheet) And Index >= LBound(styleSheet) Then
            Set fPreview = New frmPreview
            Load fPreview
            wp.Length = Len(wp)
            rtn = GetWindowPlacement(UserControl.ContainerHwnd, wp)
            If rtn <> 0 Then
                With wp.rcNormalPosition
                    fPreview.Move .Left * Screen.TwipsPerPixelX, .Top * Screen.TwipsPerPixelY, (.Right - .Left) * Screen.TwipsPerPixelX, (.Bottom - .Top) * Screen.TwipsPerPixelY
                End With
            End If
            If styleSheet(Index) <> vbNullString Then
                sURL = CreatePreview(styleSheet(Index))
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
        End If
    End If
End Sub

Private Sub cmdClear_Click(Index As Integer)
Dim strCompare As String
    If IsValidArray(styleSheet) Then
        If Index <= UBound(styleSheet) And Index >= LBound(styleSheet) Then
            strCompare = styleSheet(Index)
            removeStylesheet (Index) 'This was never added in the clear for some reason....
            lblStatus(Index).Caption = "Unassigned"
            lblStatus(Index).foreColor = vbBlack
            styleSheet(Index) = vbNullString
            bChanged = bChanged Or StrComp(vbNullString, strCompare, vbTextCompare) <> 0
        End If
    End If
End Sub

Private Function IsValidArray(ByRef this As Variant) As Boolean
    If IsArray(this) Then
        IsValidArray = GetArrayDimensions(VarPtrArray(this)) >= 1
    Else
        IsValidArray = False
    End If
End Function

Private Function GetArrayDimensions(ByVal arrPtr As Long) As Integer
   Dim address As Long
   
   CopyMemory address, ByVal arrPtr, ByVal 4   'get the address of the SafeArray structure in memory
   If address <> 0 Then 'if there is a dimension, then address will point to the memory address of the array, otherwise the array isn't dimensioned
      CopyMemory GetArrayDimensions, ByVal address, 2 'fill the local variable with the first 2 bytes of the safearray structure. These first 2 bytes contain an integer describing the number of dimensions
   End If

End Function

Private Function VarPtrArray(arr As Variant) As Long

  'Function to get pointer to the array
   CopyMemory VarPtrArray, ByVal VarPtr(arr) + 8, ByVal 4
    
End Function
