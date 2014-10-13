VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   3600
      ScaleHeight     =   1665
      ScaleWidth      =   1785
      TabIndex        =   12
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Select Picture"
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Select Font"
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtFont 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Sample"
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtColor 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Select Color"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Win Version"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "IE Version"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   2880
      TabIndex        =   4
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open System Applet"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   0
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   0
      List            =   "Form1.frx":003D
      TabIndex        =   2
      Top             =   0
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set Windows Scheme"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.ComboBox schemes 
      Height          =   315
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "schemes"
      Top             =   480
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   1680
      Top             =   2880
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sa As New HTE_SystemUtility.SysUtility

Private Sub Command1_Click()
    If Text2 <> "" Then
        sa.OpenSystemApplet Combo1.ListIndex, CInt(Text2.Text)
    Else
        sa.OpenSystemApplet Combo1.ListIndex
    End If
End Sub

Private Sub Command2_Click()
    If sa.SetWindowsColorScheme(schemes.Text) Then
        Debug.Print "TRUE"
    Else
        Debug.Print "FALSE"
    End If
    
End Sub


Private Sub Command3_Click()
    MsgBox sa.GetExplorerVersion
End Sub

Private Sub Command4_Click()
    Dim mySTRUCT As HTE_SystemUtility.RGB_WINVER
    MsgBox sa.GetWinVersion(mySTRUCT)
    Debug.Print CStr(sa.MinimumVersion(4, 90))
End Sub

Private Sub Command5_Click()
Dim oObj As HTE_SystemUtility.enhCommonDialog
Dim lReturn As Long
Set oObj = New HTE_SystemUtility.enhCommonDialog
lReturn = txtColor.BackColor
With oObj
    If .VBChooseColor(lReturn, , , , Me.hWnd) Then
        txtColor.BackColor = lReturn
    End If
End With
Set oObj = Nothing
'''Dim oObj As HTE_SystemUtility.CommonDialog
'''Set oObj = New HTE_SystemUtility.CommonDialog
'''On Error GoTo err_Command5
'''With oObj
'''    .CancelError = True
'''    .hWnd = Me.hWnd
'''    .ShowColor
'''End With
'''    txtColor.BackColor = oObj.Color
'''err_Command5:
'''    Set oObj = Nothing
End Sub

Private Sub Command6_Click()
Dim oObj As HTE_SystemUtility.enhCommonDialog
Dim oReturn As StdFont
Dim lReturnColor As Long
Set oObj = New HTE_SystemUtility.enhCommonDialog
Set oReturn = txtFont.Font
lReturnColor = txtFont.ForeColor
With oObj
    If .VBChooseFont(oReturn, , Me.hWnd, lReturnColor) Then
        Set txtFont.Font = oReturn
        txtFont.ForeColor = lReturnColor
    End If
End With
Set oObj = Nothing
End Sub

Private Sub Command7_Click()
Dim oObj As CommonDialog
Dim IPic As IPersistPicture
Dim Pic As PersistPicture
Set oObj = New CommonDialog
With oObj
    .Filter = "Bitmaps (*.bmp)|*.bmp"
    .CancelError = False
    .ShowOpen
    If .FileName <> vbNullString Then
        Set Picture1.Picture = LoadPicture(.FileName)
        Set Pic = New PersistPicture
        Set IPic = Pic
        Set IPic.Picture = Picture1.Picture
        Image1.Picture = IPic.Picture
        Set Pic = Nothing
    End If
End With
Set oObj = Nothing
End Sub

Private Sub Form_Load()
    Combo1.ListIndex = 0
    Dim v As Variant
    Dim Item As Variant
    v = sa.GetWindowsSchemes
    For Each Item In v
        schemes.AddItem Item
    Next
    schemes.Text = sa.GetCurrentWindowsScheme
    
End Sub

