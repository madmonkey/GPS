VERSION 5.00
Begin VB.Form frmListener 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "674D37A0-7B21-45c7-9087-0C877495E908"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1440
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   1440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sc As HTE_ServerSocket.Subclass

Implements HTE_ServerSocket.ISubclass

Public Event ProcessMessage(ByVal wParam As Long, lParam As Long)

Private Sub Form_Load()
    Set sc = New HTE_ServerSocket.Subclass
    sc.AttachMessage modWinsock.appMsg 'WINSOCK_MESSAGE
    sc.Subclass Me.hWnd, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not sc Is Nothing Then
        sc.UnSubclass
        Set sc = Nothing
    End If
End Sub

Private Sub ISubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    If uMsg = modWinsock.appMsg Then 'WINSOCK_MESSAGE Then
        RaiseEvent ProcessMessage(wParam, lParam)
    End If
End Sub

Private Sub ISubclass_Before(lHandled As Long, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

End Sub
