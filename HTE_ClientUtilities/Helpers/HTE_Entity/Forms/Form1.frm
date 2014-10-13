VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents AliasLookup As Alias
Attribute AliasLookup.VB_VarHelpID = -1

Private Sub AliasLookup_LookupFailed(cMsg As HTE_GPS.GPSMessage)
    Debug.Print cMsg.Entity
End Sub

Private Sub AliasLookup_LookupResolved(cMsg As HTE_GPS.GPSMessage)
    Debug.Print cMsg.Entity
End Sub

Private Sub Command1_Click()
Dim cMsg As cMessage
    Set cMsg = New cMessage
    cMsg.MessageType = GPS_TYPE_0
    cMsg.rawMessage = "test"
    cMsg.MessageStatus = GPS_MSG_PROCESSED
    
    AliasLookup.ResolveEntity cMsg, "10.255.103.102"
End Sub

Private Sub Form_Load()
    Set AliasLookup = New Alias
    AliasLookup.EnableMacResolution = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set AliasLookup = Nothing
End Sub
