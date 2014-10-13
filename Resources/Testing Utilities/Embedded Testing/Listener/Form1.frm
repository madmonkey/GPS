VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "What you said..."
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
With Winsock1
    .RemoteHost = "127.0.0.1"
    .Bind 21000
End With
Me.Caption = "Listening @ " & Winsock1.RemoteHost & ":" & Winsock1.LocalPort
'Debug.Print CStr((Winsock1.State = sckOpen) Or (Winsock1.State = sckConnected))
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
                    Debug.Print strData
            End Select
        Else
            Text1.Text = Text1.Text & strData
            Winsock1.SendData "recv'd @ " & CStr(Now)
        End If
End Sub

