VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "docHostUIHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDocHostUIHandler

Private Const cModuleName = "IDocHostUIHandler"
Private Const S_OK = &H0
Private Const S_FALSE = &H1
Private Const E_NOTIMPL = &H80004001
Private Const E_FAIL = &H80004005
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_CHAR = &H102
Private Const SYS_CMD = &H1B
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event ShowCustomMenu()
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private m_IEHotKeys As Boolean
Private custDoc As ICustomDoc
Private m_lastKeyStroke As Long
Public Property Get IEHotKeysEnabled() As Boolean
    IEHotKeysEnabled = m_IEHotKeys
End Property
Public Property Let IEHotKeysEnabled(ByVal vData As Boolean)
    m_IEHotKeys = vData
End Property
Public Property Set Document(ByRef vData As HTMLDocument)
    Do
        DoEvents
        Set custDoc = vData
    Loop While custDoc Is Nothing
    custDoc.SetUIHandler Me
End Property
Private Sub IDocHostUIHandler_ShowContextMenu(ByVal dwID As Long, ppt As WebBrowserInterfaces.tagPOINT, ByVal pcmdtReserved As Long, ByVal pdispReserved As Object)
On Error Resume Next
    'WE NEED TO SIGNAL TO THE CONTROL NOT TO SHOW THE MENU - THROUGH COM ERROR EVENT
    'THEN DISPLAY OUR OWN MENU IN ITS PLACE!
    Err.Raise E_FAIL
    RaiseEvent ShowCustomMenu
End Sub
Private Sub IDocHostUIHandler_EnableModeless(ByVal fEnable As Long)
    defaultDocHandlerBehavior "IDocHostUIHandler_EnableModeless"
End Sub
Private Sub IDocHostUIHandler_FilterDataObject(ByVal pDO As WebBrowserInterfaces.IDataObject, ppDORet As WebBrowserInterfaces.IDataObject)
    defaultDocHandlerBehavior "IDocHostUIHandler_FilterDataObject"
End Sub
Private Sub IDocHostUIHandler_GetDropTarget(ByVal pDropTarget As WebBrowserInterfaces.IDropTarget, ppDropTarget As WebBrowserInterfaces.IDropTarget)
    defaultDocHandlerBehavior "IDocHostUIHandler_GetDropTarget"
End Sub
Private Sub IDocHostUIHandler_GetExternal(ppDispatch As Object)
    defaultDocHandlerBehavior "IDocHostUIHandler_GetExternal"
End Sub
Private Sub IDocHostUIHandler_GetHostInfo(pInfo As WebBrowserInterfaces.DOCHOSTUIINFO)
    'defaultDocHandlerBehavior "IDocHostUIHandler_GetHostInfo"
    pInfo.dwFlags = pInfo.dwFlags Or DOCHOSTUIFLAG_NO3DBORDER
End Sub
Private Sub IDocHostUIHandler_GetOptionKeyPath(pchKey As String, ByVal dw As Long)
    defaultDocHandlerBehavior "IDocHostUIHandler_GetOptionKeyPath"
End Sub
Private Sub IDocHostUIHandler_HideUI()
    defaultDocHandlerBehavior "IDocHostUIHandler_HideUI"
End Sub
Private Sub IDocHostUIHandler_OnDocWindowActivate(ByVal fActivate As Long)
    defaultDocHandlerBehavior "IDocHostUIHandler_OnDocWindowActivate"
End Sub
Private Sub IDocHostUIHandler_OnFrameWindowActivate(ByVal fActivate As Long)
    defaultDocHandlerBehavior "IDocHostUIHandler_OnFrameWindowActivate"
End Sub
Private Sub IDocHostUIHandler_ResizeBorder(prcBorder As WebBrowserInterfaces.tagRECT, ByVal pUIWindow As WebBrowserInterfaces.IOleInPlaceUIWindow, ByVal fRameWindow As Long)
    defaultDocHandlerBehavior "IDocHostUIHandler_ResizeBorder"
End Sub
Private Sub IDocHostUIHandler_ShowUI(ByVal dwID As Long, ByVal pActiveObject As WebBrowserInterfaces.IOleInPlaceActiveObject, ByVal pCommandTarget As WebBrowserInterfaces.IOleCommandTarget, ByVal pFrame As WebBrowserInterfaces.IOleInPlaceFrame, ByVal pDoc As WebBrowserInterfaces.IOleInPlaceUIWindow)
    defaultDocHandlerBehavior "IDocHostUIHandler_ShowUI"
End Sub
Private Sub IDocHostUIHandler_TranslateAccelerator(lpmsg As WebBrowserInterfaces.tagMSG, pguidCmdGroup As WebBrowserInterfaces.GUID, ByVal nCmdID As Long)
    Select Case lpmsg.message
        Case WM_KEYDOWN, WM_KEYUP
            Select Case nCmdID
                Case SYS_CMD
                    If Not IEHotKeysEnabled Then
                        Select Case lpmsg.wParam
                            Case vbKeyN, vbKeyB, vbKeyE, vbKeyF, vbKeyH, vbKeyI, vbKeyL, vbKeyO, vbKeyP, vbKeyR, vbKeyW
                                RaiseEvent KeyDown(CInt(lpmsg.wParam), vbCtrlMask)
                                m_lastKeyStroke = lpmsg.wParam 'save last value for faux detection
                                Exit Sub
                        End Select
                    End If
                    m_lastKeyStroke = lpmsg.wParam 'save last value for faux detection
                    'If Not IgnoreSendKeyValue(lpmsg.wParam) Then
                    Err.Raise S_OK
                Case Else 'regular events
                    If Not IEHotKeysEnabled Then
                        If m_lastKeyStroke = vbKeyControl Then
                            Select Case lpmsg.wParam
                                Case vbKeyN, vbKeyB, vbKeyE, vbKeyF, vbKeyH, vbKeyI, vbKeyL, vbKeyO, vbKeyP, vbKeyR, vbKeyW
'                                    UEH_Log cModuleName, "IDocHostUIHandler_TranslateAccelerator", "Discarding faux IE command!"
                                    RaiseEvent KeyDown(CInt(lpmsg.wParam), vbCtrlMask)
                                    m_lastKeyStroke = lpmsg.wParam 'save last value for faux detection
                                    Exit Sub
                            End Select
                        'need to check for reverse as well!!!
                        End If
                    End If
                    m_lastKeyStroke = lpmsg.wParam 'save last value for faux detection
                    'If Not IgnoreSendKeyValue(lpmsg.wParam) Then
                    Err.Raise S_OK
            End Select
        Case WM_CHAR
            'This one is tricky...the wParam value is filled with the ASCII equivilent value NOT the keyboard value
            'Let this one pass through...
            Err.Raise E_NOTIMPL
        Case Else
            Err.Raise E_NOTIMPL
            'Shouldn't be here - I've never seen one anyway...
    End Select
        
End Sub
Private Sub IDocHostUIHandler_TranslateUrl(ByVal dwTranslate As Long, ByVal pchURLIn As String, ppchURLOut As String)
    defaultDocHandlerBehavior "IDocHostUIHandler_TranslateUrl"
End Sub
Private Sub IDocHostUIHandler_UpdateUI()
    defaultDocHandlerBehavior "IDocHostUIHandler_UpdateUI"
End Sub
Private Sub defaultDocHandlerBehavior(ByVal sRoutine As String)
On Error Resume Next
    'UEH_Log cModuleName, sRoutine
    Err.Raise E_NOTIMPL
End Sub
'Private Function IgnoreSendKeyValue(ByVal wParam As Long) As Boolean
'    Debug.Print "wParam = " & wParam
'    Select Case SendKey
'        Case vfGHK_Add
'            IgnoreSendKeyValue = CBool((wParam = vbKeyAdd) Or ((wParam = 187) And (GetAsyncKeyState(vbKeyShift) And &H8000)))
'            Debug.Print "IgnoreSendKeyValue - vfGHK_Add = " & CStr(CBool((wParam = vbKeyAdd) Or ((wParam = 187) And (GetAsyncKeyState(vbKeyShift) And &H8000))))
'        Case vfGHK_Apostrophe
'            IgnoreSendKeyValue = (wParam = 96)
'        Case vfGHK_Decimal
'            IgnoreSendKeyValue = (wParam = vbKeyDecimal) Or (wParam = 190)
'        Case vfGHK_Divide
'            IgnoreSendKeyValue = (wParam = vbKeyDivide) Or (wParam = 191)
'        Case vfGHK_Equals
'            IgnoreSendKeyValue = (wParam = 61)
'        Case vfGHK_Multiply
'            IgnoreSendKeyValue = (wParam = vbKeyMultiply) Or ((wParam = 56) And (GetAsyncKeyState(vbKeyShift) And &H8000))
'        Case vfGHK_Subtract
'            IgnoreSendKeyValue = (wParam = vbKeySubtract) Or (wParam = 109)
'        Case vfGHK_Tilde
'            IgnoreSendKeyValue = (wParam = 96) And (GetAsyncKeyState(vbKeyShift) And &H8000)
'        Case Else
'            IgnoreSendKeyValue = False
'    End Select
'
'End Function

