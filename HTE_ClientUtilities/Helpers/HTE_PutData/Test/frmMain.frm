VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test HTE_PubData"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Initialize"
      Height          =   1275
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   4095
      Begin VB.CommandButton cmdSetTopic 
         Caption         =   "Set Topic"
         Height          =   375
         Left            =   2700
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtTopic 
         Height          =   315
         Left            =   180
         TabIndex        =   7
         Text            =   "HTE_MDB"
         Top             =   300
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Publish"
      Height          =   2115
      Left            =   120
      TabIndex        =   0
      Top             =   1500
      Width           =   4095
      Begin VB.CommandButton cmdSendString 
         Caption         =   "SendString"
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   1560
         Width           =   1275
      End
      Begin VB.TextBox txtData 
         Height          =   675
         Left            =   780
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "frmMain.frx":0000
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtTag 
         Height          =   315
         Left            =   780
         TabIndex        =   1
         Text            =   "TestTag"
         Top             =   300
         Width           =   3135
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Data"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tag"
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Pub As HTE_PubData.Publisher

Private Sub cmdSendString_Click()
    Pub.SendString txtTag.Text, txtData.Text
End Sub

Private Sub cmdSetTopic_Click()
    Pub.Topic = txtTopic.Text
End Sub

Private Sub Form_Initialize()
    Set Pub = New HTE_PubData.Publisher
End Sub

Private Sub Form_Terminate()
    Set Pub = Nothing
End Sub
