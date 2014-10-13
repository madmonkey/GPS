Attribute VB_Name = "modStyle"
Option Explicit

Private Const GWL_STYLE As Long = (-16)
 
Public Enum esWinStyles
    esUpperCase = &H8&
    esLowerCase = &H10&
    esNumeric = &H2000
End Enum
 
Public Enum emWinMessages
    emReadOnly = &HCF
    emWritable = 1
    emUndo = &HC7
    emScrollCaret = &HB7
    emSetSel = &HB1
End Enum
 
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 
Public Sub setStyle(hWnd As Long, esStyle As esWinStyles)
    
    Call SetWindowLong(hWnd, GWL_STYLE, GetWindowLong(hWnd, GWL_STYLE) Or esStyle)
 
End Sub
 
Public Sub setBehavior(hWnd As Long, emMessage As emWinMessages)
 
    Call SendMessage(hWnd, emMessage, (emMessage = emWritable), ByVal 0&)
 
End Sub

Public Sub Highlight(ByRef this As TextBox)
On Error Resume Next
    this.SelStart = 0
    this.SelLength = Len(this.Text)
End Sub
