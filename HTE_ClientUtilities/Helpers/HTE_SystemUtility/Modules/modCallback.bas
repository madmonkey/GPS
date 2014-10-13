Attribute VB_Name = "modCallback"
Option Explicit

Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private myClassName As String
Private myTitle As String
Private foundhWnd As Long

Public Sub EnsurePosition(ByVal hWnd As Long)
Const SW_SHOWNORMAL = 1&
Const SW_SHOWMINIMIZED = 2&
Const SW_TOPMOST = -1&
Const SWP_SHOWWINDOW = &H40
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const HWND_NOTOPMOST = -2

    SetWindowPos hWnd, SW_TOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub
Public Function FindSpecializedWindow(ByVal ClassName As String, ByVal Title As String, Optional Retries As Long = 20) As Long
Dim loopUntil As Long
    myClassName = vbNullString: myClassName = ClassName
    myTitle = vbNullString: myTitle = Title
    foundhWnd = 0
    For loopUntil = 0 To Retries
        Call EnumWindows(AddressOf EnumWindowProc, &H0)
        If foundhWnd <> 0 Then Exit For
    Next
    FindSpecializedWindow = foundhWnd
End Function
Public Function EnumWindowProc(ByVal hWnd As Long, ByVal lParam As Long) As Long

Dim sTitle As String
Dim sClass As String
Const MAX_PATH = 260
    
    sTitle = Space$(MAX_PATH)
    sClass = Space$(MAX_PATH)
    Call GetClassName(hWnd, sClass, MAX_PATH)
    Call GetWindowText(hWnd, sTitle, MAX_PATH)
    If StrComp(TrimNull(sClass), myClassName, vbTextCompare) = 0 And StrComp(TrimNull(sTitle), myTitle, vbTextCompare) = 0 Then
        'WE FOUND IT QUIT PROCESSING WINDOWS
        foundhWnd = hWnd
        EnumWindowProc = 0
    Else
        EnumWindowProc = 1
    End If

End Function


Private Function TrimNull(item As String)

  'remove string before the terminating null(s)
   Dim pos As Integer
   
   pos = InStr(item, Chr$(0))
   
   If pos Then
         TrimNull = Left$(item, pos - 1)
   Else: TrimNull = item
   End If
   
End Function

