VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Embedded Tester..."
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
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
   ScaleHeight     =   3045
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstReply 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   2640
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Winsock1.SendData Text1.Text
End Sub
Private Sub Form_Load()
With Winsock1
    .RemoteHost = "127.0.0.1"
    .RemotePort = 21000
End With
Me.Caption = "Transmitting to " & Winsock1.RemoteHost & ":" & Winsock1.RemotePort
Debug.Print CStr((Winsock1.State = sckOpen) Or (Winsock1.State = sckConnected))
Me.Show
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
    On Error Resume Next
        Winsock1.GetData strData
        If Err Then
            Select Case Err.Number
                Case 10054
                    Debug.Print "Error"
                Case Else
                    lstReply.AddItem strData
            End Select
        Else
            lstReply.AddItem strData
        End If
End Sub

