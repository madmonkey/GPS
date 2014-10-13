VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProperties 
   Appearance      =   0  'Flat
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   8385
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProperties.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frProperties 
      Caption         =   "   Process Description"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8175
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   360
         ScaleHeight     =   1095
         ScaleWidth      =   7695
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   360
         Width           =   7695
         Begin VB.TextBox txtFriendlyDesc 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   0
            TabIndex        =   0
            Top             =   720
            Width           =   7455
         End
         Begin VB.Label lblProcess 
            Caption         =   "lblProcess"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   7335
         End
         Begin VB.Label lblInstance 
            Caption         =   "lblInstance"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   240
            Width           =   7335
         End
         Begin VB.Label lblStatusDesc 
            Caption         =   "lblStatusDesc"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   11
            Top             =   480
            Width           =   7335
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   475
         Left            =   30
         ScaleHeight     =   480
         ScaleWidth      =   8055
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   5760
         Width           =   8055
         Begin VB.CommandButton cmdOK 
            Caption         =   "&OK"
            Default         =   -1  'True
            Height          =   375
            Left            =   4200
            TabIndex        =   2
            Top             =   75
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
            Caption         =   "Cance&l"
            Height          =   375
            Left            =   5520
            TabIndex        =   3
            Top             =   75
            Width           =   1215
         End
         Begin VB.CommandButton cmdApply 
            Caption         =   "&Apply"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6840
            TabIndex        =   4
            Top             =   75
            Width           =   1215
         End
      End
      Begin VB.Timer tmrApply 
         Interval        =   100
         Left            =   600
         Top             =   5160
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   3855
         Left            =   7800
         TabIndex        =   5
         Top             =   1560
         Width           =   220
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   220
         Left            =   360
         TabIndex        =   6
         Top             =   5520
         Width           =   7455
      End
      Begin VB.PictureBox picViewport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3975
         Left            =   360
         ScaleHeight     =   3945
         ScaleWidth      =   7425
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1560
         Width           =   7455
         Begin VB.Frame frWorkspace 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000C&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2000
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   3000
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "No Property Page Available"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   36
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   2295
               Left            =   360
               TabIndex        =   8
               Top             =   480
               Width           =   6975
            End
         End
      End
      Begin VB.Image imgStatus 
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   65
         Picture         =   "frmProperties.frx":000C
         Top             =   0
         Width           =   240
      End
   End
   Begin MSComctlLib.ImageList ILStatusList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProperties.frx":0391
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProperties.frx":0726
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProperties.frx":0AC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProperties.frx":0E5A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_PropPage As HTE_GPS.PropertyPage
'in some instances the node doesn't get saved from callback
'use callback to communicate with utilities, event for interface always!
Public Event SaveChanges(ByVal XMLDOMNode As String)
Dim WithEvents m_Extender As VB.VBControlExtender
Attribute m_Extender.VB_VarHelpID = -1
Private Const cModuleName = "frmProperties"
Private Const vpTop = 1560&
Private Const vpLeft = 360&
Private Const vpOffsetWidth = 685&
Private Const vpOffsetHeight = vpTop + 840&
Private lMinWidth As Long
Private lMinHeight As Long
Private initDesc As String
Public Event DescriptionChange(ByVal NewDesc As String)
Public Property Set propPage(ByRef vData As HTE_GPS.PropertyPage)
On Error GoTo err_PropPage
    If Not vData Is Nothing Then
        If vData.LicenseKey <> vbNullString Then Licenses.Add vData.Name, vData.LicenseKey
        Set m_Extender = Me.Controls.Add(vData.Name, Replace(vData.Name, ".", "_"), frWorkspace)
        Set m_PropPage = m_Extender.Object
        m_PropPage.Settings = vData.Settings
        m_PropPage.PropertyCallback = vData.PropertyCallback
        lMinWidth = IIf(m_Extender.Width > picViewport.Width, m_Extender.Width, picViewport.Width)
        lMinHeight = IIf(m_Extender.Height > picViewport.Height, m_Extender.Height, picViewport.Height)
        frWorkspace.Move 0, 0, lMinWidth, lMinHeight
        m_Extender.Move 0, 0, frWorkspace.Width, frWorkspace.Height
        m_Extender.Visible = True
        m_Extender.ZOrder
    Else
        UEH_Log cModuleName, "PropPage Set", "Property page is nothing!", logWarning
    End If
    Exit Property
err_PropPage:
    UEH_LogError cModuleName, "PropPage Set", Err
End Property

Private Sub cmdApply_Click()
Dim strSettings As String
On Error GoTo err_cmdApply
    If StrComp(initDesc, txtFriendlyDesc.Text, vbBinaryCompare) <> 0 Then RaiseEvent DescriptionChange(txtFriendlyDesc.Text)
    If Not m_PropPage Is Nothing Then
        If m_PropPage.Changed Then
            If m_PropPage.SaveChanges Then
                strSettings = m_PropPage.Settings
                UEH_Log cModuleName, "cmdApply", "Process settings", logVerbose, , strSettings, logXML
                If Len(strSettings) > 0 Then RaiseEvent SaveChanges(strSettings)
            Else
                UEH_Log cModuleName, "cmdApply", "Property interface callback returned false!", logWarning
            End If
        End If
    End If
    initDesc = txtFriendlyDesc.Text
    cmdApply.Enabled = False
    tmrApply.Enabled = True
    Exit Sub
err_cmdApply:
    UEH_LogError cModuleName, "cmdApply", Err
End Sub

Private Sub cmdCancel_Click()
    Cancel
End Sub
Public Sub Cancel()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim strSettings As String
On Local Error Resume Next

    If StrComp(initDesc, txtFriendlyDesc.Text, vbBinaryCompare) <> 0 Then
        RaiseEvent DescriptionChange(txtFriendlyDesc.Text)
    End If
    If Not m_PropPage Is Nothing Then
        If m_PropPage.Changed Then
            If m_PropPage.SaveChanges Then
                strSettings = m_PropPage.Settings
                UEH_Log cModuleName, "cmdOK", "Process settings", logVerbose, , strSettings, logXML
                If Len(strSettings) > 0 Then RaiseEvent SaveChanges(strSettings)
            End If
        Else
            UEH_Log cModuleName, "cmdOK", "Property interface callback returned false!", logWarning
        End If
    End If
    Unload Me
End Sub

Private Sub Form_Initialize()
    InitializeForThemes
End Sub

Private Sub Form_Resize()
Dim lWidth As Long
Dim lHeight As Long
Dim lvpW As Long, lvpH As Long
    If Me.WindowState <> vbMinimized Then
        'Form Frame
        frProperties.Move 0, 0, Me.ScaleWidth - 15, Me.ScaleHeight - 15
        'ViewPort
        lvpW = frProperties.Width - vpOffsetWidth
        lvpH = frProperties.Height - (vpOffsetHeight)
        picViewport.Move vpLeft, vpTop, IIf(lvpW < 0, 0, lvpW), IIf(lvpH < 0, 0, lvpH)
        If picViewport.Width > lMinWidth Then
            lWidth = picViewport.Width
        Else
            lWidth = lMinWidth
        End If
        If picViewport.Height > lMinHeight Then
            lHeight = picViewport.Height
        Else
            lHeight = lMinHeight
        End If
        'Workspace
        frWorkspace.Move 0, 0, lWidth, lHeight
        If Not m_Extender Is Nothing Then m_Extender.Move 0, 0, lWidth, lHeight
        'Scrollbars
        HScroll1.Move vpLeft, (picViewport.Height + vpTop), picViewport.Width
        VScroll1.Move (picViewport.Width + 130 + VScroll1.Width), vpTop, VScroll1.Width, picViewport.Height
        HScroll1.Max = frWorkspace.Width - picViewport.Width
        VScroll1.Max = frWorkspace.Height - picViewport.Height
        If HScroll1.Max > 0 Then
            HScroll1.LargeChange = Abs(HScroll1.Max) / 10
            HScroll1.SmallChange = Abs(CInt(HScroll1.Max) / 25) + 1
        End If
        If VScroll1.Max > 0 Then
            VScroll1.LargeChange = Abs(VScroll1.Max) / 10
            VScroll1.SmallChange = Abs(VScroll1.Max) / 25
        End If
        'Offset for scrollbars
        VScroll1.Visible = (picViewport.Height < lHeight)
        HScroll1.Visible = (picViewport.Width < lWidth)
        'Labels
        Picture2.Width = picViewport.Width
        txtFriendlyDesc.Width = picViewport.Width
        lblStatusDesc.Width = txtFriendlyDesc.Width
        lblInstance.Width = txtFriendlyDesc.Width
        lblProcess.Width = txtFriendlyDesc.Width
        'Buttons
        Picture1.Move Picture1.Left, (vpTop + picViewport.ScaleHeight) + 375, (frProperties.Width - Picture1.Left)
        'cmdApply.Move (frProperties.Width - picViewport.Width) + picViewport.Width - (picViewport.Left + cmdApply.Width), (vpTop + picViewport.ScaleHeight) + 375
        cmdApply.Move (frProperties.Width - picViewport.Width) + picViewport.Width - (picViewport.Left + cmdApply.Width), 50
        cmdCancel.Move (cmdApply.Left - (cmdCancel.Width + 105)), cmdApply.Top
        cmdOK.Move (cmdCancel.Left - (cmdOK.Width + 105)), cmdApply.Top
        
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_Extender = Nothing
    If Not m_PropPage Is Nothing Then
        Me.Controls.Remove Replace(m_PropPage.Name, ".", "_")
        m_PropPage.Exit
    End If
    Set m_PropPage = Nothing
End Sub

Private Sub tmrApply_Timer()
    ApplyEnabled
End Sub

Private Sub txtFriendlyDesc_Change()
    If initDesc = vbNullString Then initDesc = txtFriendlyDesc.Text
    Me.Caption = txtFriendlyDesc.Text
    ApplyEnabled
End Sub

Private Sub txtFriendlyDesc_GotFocus()
    txtFriendlyDesc.SelStart = 0
    txtFriendlyDesc.SelLength = Len(txtFriendlyDesc.Text)
End Sub

Private Sub VScroll1_Change()
    frWorkspace.Top = -VScroll1.Value
End Sub

Private Sub HScroll1_Change()
    frWorkspace.Left = -HScroll1.Value
End Sub

Private Sub ApplyEnabled()
Dim bEval As Boolean
    bEval = StrComp(initDesc, txtFriendlyDesc.Text, vbBinaryCompare) <> 0
    If Not m_PropPage Is Nothing Then
        bEval = bEval Or m_PropPage.Changed
    End If
    cmdApply.Enabled = bEval
    If bEval Then tmrApply.Enabled = False
End Sub
