VERSION 5.00
Begin VB.UserControl CustomPropertyPage 
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6480
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
   ScaleHeight     =   3075
   ScaleWidth      =   6480
   ToolboxBitmap   =   "CustomPropertyPage.ctx":0000
   Begin VB.CheckBox chkAutoInsert 
      Caption         =   "Auto Insert Entities"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   2895
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
      Left            =   4080
      TabIndex        =   4
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
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtType 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
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
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Unassigned"
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   3
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
Dim styleSheet() As String
Private Const cFileDescriptor = "htc"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function getTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

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

Private Sub chkAutoInsert_Click()
    bChanged = bChanged Or letProperty(m_Settings, cAutoInsert, CStr(CBool(chkAutoInsert.Value)))
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
            If oObj.VBGetOpenFileName(fileName, "Select Stylesheet", True, , , True, "XML Stylesheet (*.xsl)|*.xsl", , App.Path, , , UserControl.hWnd) Then
                fh = FreeFile
                Open fileName For Binary Access Read Lock Read As #fh
                strFile = Input(LOF(fh), fh)
                Close #fh
                If validateStylesheet(strFile) Then
                    styleSheet(Index) = strFile
                    lblStatus(Index).Caption = "Assigned"
                    lblStatus(Index).ForeColor = vbBlack
                    addStyleSheet (Index)
                Else
                    lblStatus(Index).Caption = "Unassigned"
                    lblStatus(Index).ForeColor = vbRed
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

Private Function addStyleSheet(ByVal Index As Integer)
Dim iNodes As MSXML2.IXMLDOMNodeList
Dim iNode As MSXML2.IXMLDOMNode, iAttribute As MSXML2.IXMLDOMNode
Dim newAtt As IXMLDOMAttribute, namedNodeMap As IXMLDOMNamedNodeMap
Dim sBefore As String
Dim i As Long
    If Not m_Settings Is Nothing Then
        sBefore = m_Settings.xml
        Set iNodes = m_Settings.documentElement.getElementsByTagName(txtType(Index).Text)
        If iNodes Is Nothing Then
            Set iNode = m_Settings.createElement(txtType(Index).Text)
            iNode.nodeTypedValue = styleSheet(Index)
        Else
            For i = 0 To iNodes.length - 1
                If iNodes.Item(i).parentNode.nodeName <> "TYPES" Then
                    Set iNode = m_Settings.getElementsByTagName(txtType(Index).Text).nextNode
                    Exit For
                End If
            Next
            If iNode Is Nothing Then
                Set iNode = m_Settings.createElement(txtType(Index).Text)
                If Not iNode Is Nothing Then iNode.nodeTypedValue = styleSheet(Index)
            Else
                iNode.nodeTypedValue = styleSheet(Index)
            End If
        End If
        Set iAttribute = iNode.Attributes.getNamedItem("GPSTYPE")
        If iAttribute Is Nothing Then
            Set newAtt = m_Settings.createAttribute("GPSTYPE")
            newAtt.nodeTypedValue = txtType(Index).Tag
            Set namedNodeMap = iNode.Attributes
            Set iAttribute = namedNodeMap.setNamedItem(newAtt)
        Else
            iAttribute.nodeTypedValue = txtType(Index).Tag
        End If
        m_Settings.documentElement.appendChild iNode
        bChanged = bChanged Or StrComp(m_Settings.xml, sBefore, vbTextCompare) <> 0
    End If
End Function

Private Function removeStylesheet(ByVal Index As Integer)
Dim iNodes As MSXML2.IXMLDOMNodeList, iNode As MSXML2.IXMLDOMNode, iAttribute As MSXML2.IXMLDOMNode
Dim i As Long
    If Not m_Settings Is Nothing Then
        Set iNodes = m_Settings.documentElement.getElementsByTagName(txtType(Index).Text)
        If Not iNodes Is Nothing Then
            For i = 0 To iNodes.length - 1
                Set iNode = iNodes.Item(i)
                Set iAttribute = iNode.Attributes.getNamedItem(cAttrType)
                'Verify based on what we know, not what we think we know
                If Not iNode Is Nothing And Not iAttribute Is Nothing Then
                    m_Settings.documentElement.removeChild iNode
                    bChanged = True
                End If
            Next
        End If
    End If
End Function
Private Sub cmdClear_Click(Index As Integer)
Dim strCompare As String
    If IsValidArray(styleSheet) Then
        If Index <= UBound(styleSheet) And Index >= LBound(styleSheet) Then
            strCompare = styleSheet(Index)
            removeStylesheet (Index) 'This was never added in the clear for some reason....
            lblStatus(Index).Caption = "Unassigned"
            lblStatus(Index).ForeColor = vbBlack
            styleSheet(Index) = vbNullString
            bChanged = bChanged Or StrComp(vbNullString, strCompare, vbTextCompare) <> 0
        End If
    End If
End Sub

Private Sub cmdPreview_Click(Index As Integer)
Dim fPreview As frmPreview
Dim sURL As String
Const cBlankPage = "about:blank"
    If IsValidArray(styleSheet) Then
        If Index <= UBound(styleSheet) And Index >= LBound(styleSheet) Then
            Set fPreview = New frmPreview
            Load fPreview
            fPreview.Move UserControl.Parent.Left, UserControl.Parent.Top, UserControl.Parent.Width, UserControl.Parent.Height
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
Dim iNode As MSXML2.IXMLDOMNode, iAttribute As MSXML2.IXMLDOMNode, iNodes As MSXML2.IXMLDOMNodeList
Dim i As Long, x As Long
Dim iChild As MSXML2.IXMLDOMNode
    Set m_Settings = New MSXML2.DOMDocument30
    m_Settings.async = False
    Erase styleSheet
    If sXML <> vbNullString Then
        loadLocalSettings = m_Settings.loadXML(sXML)
        If loadLocalSettings Then
            For i = 1 To txtType.UBound
                Unload txtType(i)
                Unload cmdBrowse(i)
                Unload cmdPreview(i)
                Unload lblStatus(i)
            Next
            'AUTO-CONFIGURE
            chkAutoInsert.Value = Abs(CBool(getProperty(m_Settings, cAutoInsert, cAutoInsertValue)))
            'FROM GENERAL/TYPES CONFIGURATION CREATE/LOAD PROPERTIES TO FILL
            If Not m_Settings.getElementsByTagName("TYPES") Is Nothing Then
                If m_Settings.getElementsByTagName("TYPES").length > 0 Then
                    Set iNode = m_Settings.getElementsByTagName("TYPES").Item(0)
                    If Not iNode Is Nothing Then
                        If iNode.hasChildNodes Then
                            ReDim styleSheet(iNode.childNodes.length - 1)
                            For i = 0 To iNode.childNodes.length - 1
                                If i > 0 Then
                                    Load txtType(i): txtType(i).Top = txtType(i - 1).Top + txtType(i - 1).Height + 50: txtType(i).Visible = True
                                    Load cmdBrowse(i): cmdBrowse(i).Top = cmdBrowse(i - 1).Top + cmdBrowse(i - 1).Height + 50: cmdBrowse(i).Visible = True
                                    Load cmdPreview(i): cmdPreview(i).Top = cmdPreview(i - 1).Top + cmdPreview(i - 1).Height + 50: cmdPreview(i).Visible = True
                                    Load cmdClear(i): cmdClear(i).Top = cmdClear(i - 1).Top + cmdClear(i - 1).Height + 50: cmdClear(i).Visible = True
                                    Load lblStatus(i): lblStatus(i).Top = lblStatus(i - 1).Top + lblStatus(i - 1).Height + 50: lblStatus(i).Visible = True
                                    'Resize control if needed...probably not in this lifetime
                                    If txtType(i).Height + txtType(i).Top > UserControl.Height Then
                                        UserControl.Height = (txtType(i).Height * 2) + txtType(i).Top
                                    End If
                                End If
                                'TAG CORRESPONDS TO GPSTYPE
                                txtType(i).Text = iNode.childNodes(i).nodeName
                                txtType(i).Tag = iNode.childNodes(i).nodeTypedValue
                                'FIND CORRESPONDING TYPE'S XSL CONFIGURATION IF ANY....
                                Set iNodes = m_Settings.documentElement.getElementsByTagName(iNode.childNodes(i).nodeName)
                                If Not iNodes Is Nothing Then
                                    styleSheet(i) = vbNullString
                                    lblStatus(i).Caption = "Unassigned"
                                    For x = 0 To iNodes.length - 1
                                        Set iChild = iNodes.Item(x)
                                        If iChild.parentNode.nodeName <> "TYPES" Then
                                            If validateStylesheet(iChild.nodeTypedValue) Then
                                                Debug.Print "Assigning " & iNode.childNodes(i).nodeName & " styleSheet:" & iChild.nodeTypedValue
                                                styleSheet(i) = iChild.nodeTypedValue
                                                lblStatus(i).Caption = "Assigned"
                                            End If
                                        End If
                                    Next
                                End If
                            Next
                            chkAutoInsert.Move txtType(txtType.UBound).Left, _
                                txtType(txtType.UBound).Top + txtType(txtType.UBound).Height + 105 'offset for space
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function
Private Function validateStylesheet(ByVal sXSL As String) As Boolean
Dim oDOM As MSXML2.DOMDocument30
On Local Error Resume Next
    Set oDOM = New MSXML2.DOMDocument30
    oDOM.async = False
    validateStylesheet = oDOM.loadXML(sXSL)
    Set oDOM = Nothing
End Function

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



