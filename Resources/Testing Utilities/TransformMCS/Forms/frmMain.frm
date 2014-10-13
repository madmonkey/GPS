VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8070
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2040
      Width           =   7935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convert"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmMain.frx":0000
      Top             =   120
      Width           =   7935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim sArray() As String
Dim i As Long
    txtOutput.Text = vbNullString
    sArray = ProcessLines(txtInput.Text)
    txtOutput.Text = "<?xml version=" & Chr$(34) & "1.0" & Chr$(34) & "?><?xml-stylesheet type=" & Chr$(34) & "text/xsl" & Chr$(34) & " ?>" & vbCrLf & vbTab & "<GPSMessage>" & vbCrLf & vbTab & vbTab & "<rawMessage>"
    For i = LBound(sArray) To UBound(sArray)
        txtOutput.Text = txtOutput.Text & ProcessLine(sArray(i))
    Next i
    txtOutput.Text = txtOutput.Text & "</rawMessage>" & vbCrLf & vbTab & "</GPSMessage>"
End Sub

Private Function ProcessLines(ByVal sText As String) As Variant
Dim sArray() As String
    
    If sText <> vbNullString Then
        sArray = Split(sText, vbCrLf)
        ProcessLines = sArray()
    End If
    
End Function

Private Function ProcessLine(ByVal sText As String) As String
Dim sArray() As String
Dim i As Long
Dim sRtn As String
Dim sTemp As String
    sRtn = vbNullString
    If IsNumeric(Left$(Trim$(sText), 3)) Then
        sArray = Split(sText, Space$(1))
        For i = LBound(sArray) To UBound(sArray)
            If IsNumeric("&H" & sArray(i)) Then
                If Len(sArray(i)) = 2 Then
                    sTemp = "&H" & sArray(i)
                    Select Case Val(sTemp)
                        Case Is < 32
                            sRtn = sRtn & "/#*x" & sArray(i) & ";"
                        Case Else
                            sRtn = sRtn & Chr$(Val(sTemp))
                    End Select
                End If
            End If
        Next
    End If
    sRtn = Replace(sRtn, "<", "&lt;")
    sRtn = Replace(sRtn, ">", "&gt;")
    ProcessLine = sRtn
End Function
Private Static Function ByteToHex(bytVal As Byte) As String
  ByteToHex = "00"
  Mid$(ByteToHex, 3 - Len(Hex$(bytVal))) = Hex$(bytVal)
End Function

Private Sub txtInput_Change()
    Command1.Enabled = Len(txtInput.Text) > 0
End Sub
