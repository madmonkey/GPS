Attribute VB_Name = "modFlatCombo"

Option Explicit

' APIs to install our subclassing routines
Private Const GWL_WNDPROC = (-4)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' Messages which we will be processing in our subclassing routines
Private Const WM_DESTROY        As Long = &H2
Private Const WM_PAINT          As Long = &HF
Private Const WM_GETFONT        As Long = &H31

Private Const CB_GETCURSEL      As Long = &H147
Private Const CB_GETLBTEXT      As Long = &H148
Private Const CB_GETLBTEXTLEN   As Long = &H149

Private Const BDR_RAISEDOUTER   As Long = &H1
Private Const BDR_SUNKENOUTER   As Long = &H2
Private Const BDR_RAISEDINNER   As Long = &H4
Private Const BDR_SUNKENINNER   As Long = &H8

Private Const BF_BOTTOM         As Long = &H8
Private Const BF_LEFT           As Long = &H1
Private Const BF_RIGHT          As Long = &H4
Private Const BF_TOP            As Long = &H2
Private Const BF_RECT           As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const DT_BOTTOM         As Long = &H8
Private Const DT_CENTER         As Long = &H1
Private Const DT_LEFT           As Long = &H0
Private Const DT_RIGHT          As Long = &H2
Private Const DT_TOP            As Long = &H0
Private Const DT_VCENTER        As Long = &H4
Private Const DT_SINGLELINE     As Long = &H20

' A POINT
Private Type POINTAPI
    X       As Long
    Y       As Long
End Type

' A rectangle.
Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

' A painting structure/UDT, used by BeginPaint and EndPaint
Private Type PAINTSTRUCT
    hDC                     As Long
    fErase                  As Long
    rcPaint                 As RECT
    fRestore                As Long
    fIncUpdate              As Long
    rgbReserved(1 To 32)    As Byte
End Type

' API used to convert a pointer to a referenceable object
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

' APIs used to keep track of brush handles and process addresses
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long

' APIs used in our subclassing routine to create the "flat" effect.
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Boolean, ByVal fdwUnderline As Boolean, ByVal fdwStrikeOut As Boolean, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal clr As Long, ByVal hpal As Long, ByRef lpcolorref As Long)
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

'''Public Function AreThemesActive() As Boolean
''''Is theming currently active?
'''Dim hLib As Long, hProc As Long, hThemed As Long
'''
'''    hThemed = 0
'''    hLib = LoadLibrary("uxtheme.dll")
'''    If hLib <> 0 Then
'''        hProc = GetProcAddress(hLib, "IsThemeActive")
'''        If hProc <> 0 Then
'''            hThemed = CallWindowProc(hProc, 0&, 0&, 0&, 0&)
'''        End If
'''        FreeLibrary hLib
'''    End If
'''    AreThemesActive = (hThemed <> 0)
'''
'''End Function

Public Function FixFlatComboboxes(aUserControl As Object, Optional RaisedFlat As Boolean)
    
    Dim aControl As Control
    
    ' Make sure we don't have any typos in our subclassing procedures.
    NewCboProc 0, 0, 0, 0
    FullCBOPaint 0, 0, 0, 0
    
    For Each aControl In aUserControl
        If TypeOf aControl Is ComboBox Then
            ' Subclass the ComboBox, IF NOT ALREADY subclassed
            If GetProp(aControl.hWnd, "OrigProcAddr") = 0 Then
                SetProp aControl.hWnd, "OrigProcAddr", SetWindowLong(aControl.hWnd, GWL_WNDPROC, AddressOf NewCboProc)
            End If
            ' Grab control props for quick access later
            SetProp aControl.hWnd, "FixedSingle", Abs(CInt(RaisedFlat))
            SetProp aControl.hWnd, "ControlPtr", ObjPtr(aControl)
        End If
    Next aControl
    
End Function

Private Function NewCboProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim origProc        As Long
    Dim aCbo            As ComboBox
    Dim cboStyle        As Long
    Dim cboAppearance   As Long
    Dim aPtr            As Long
    
    If hWnd = 0 Then Exit Function
    ' Get the original process address which we stored earlier.
    origProc = GetProp(hWnd, "OrigProcAddr")
    
    aPtr = GetProp(hWnd, "ControlPtr")
    If aPtr <> 0 Then
        CopyMemory aCbo, aPtr, 4
        cboAppearance = aCbo.Appearance
        cboStyle = aCbo.Style
        CopyMemory aCbo, 0&, 4
    End If
    
    If origProc <> 0 And cboAppearance = 0 Then
        ' We're subclassing! Which is silly, 'cause otherwise we wouldn't be in
        '  this function, however we double check the process address just in case.
        If uMsg = WM_PAINT Then
            NewCboProc = FullCBOPaint(hWnd, uMsg, wParam, lParam)
        ElseIf uMsg = WM_DESTROY Then
            
            ' The ComboBox's parent is closing / destroying, so we need to
            '  unhook our subclassing routine ... or bad things happen
            
            ' Remove our values we stored against the ComboBox's handle
            RemoveProp hWnd, "OrigProcAddr"
            RemoveProp hWnd, "ControlPtr"
            ' Replace the original process address
            SetWindowLong hWnd, GWL_WNDPROC, origProc
            ' Invoke the default "destroy" process
            NewCboProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
        Else
            ' We're not interested in this message, so we'll just let it truck
            '  right on thru... invoke the default process
            NewCboProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
        End If
    ElseIf origProc <> 0 Then
        NewCboProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
    Else
        ' A catch-all in case something freaky happens with the process addresses.
        NewCboProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
    End If
    
End Function

Private Function FullCBOPaint(hWnd As Long, uMsg As Long, wParam As Long, lParam As Long) As Long

    ' ****************************************************************************
    '  Here's the total repaint for the "Listbox" style combobox. Its pretty
    '   ugly, but what do you expect... we're totally overriding the default
    '   drawing that windows would have done for us.
    ' ****************************************************************************

    Dim aPS             As PAINTSTRUCT  ' A paint structure UDT
    Dim aRECT           As RECT         ' A rectangle - size of the back buffer
    Dim arrowRECT       As RECT         ' Another rectangle - location of the arrow
    Dim aDC             As Long         ' same as aPS.hDC
    Dim backBuffDC      As Long         ' back buffer's DC for drawing the control
    Dim backBuffBmp     As Long         ' back buffer's bitmap for drawing the control
    Dim aPen            As Long         ' Pen used to draw the fixed single border
    Dim aLen            As Long         ' Length of the current selection's string
    Dim aStr            As String       ' The current selection string
    Dim curInd          As Long         ' Index of the current selection
    Dim cboFont         As Long         ' handle to the font used by the combobox to draw the selected item
    Dim origFont        As Long         ' Original font created with the back buffer
    Dim arrowFont       As Long         ' handle to our Marlett font used to draw the down arrow
    Dim aBrush          As Long         ' a brush object... serves many purposes
    Dim aCbo            As ComboBox     ' Dummy variable used to reference the original from a handle
    Dim aPtr            As Long         ' Pointer to the object associated with this handle.
    
    ' Variables used for quick access. We could probably get these other
    '  ways, but instead we just stored them in the internal windows DB when
    '  we started the subclassing procedure.
    Dim cboStyle        As Long         ' combobox's style (Dropdown combo, etc)
    Dim clrTxt          As OLE_COLOR    ' color of the text
    Dim clrBack         As OLE_COLOR    ' color of the background
    Dim foreColor       As Long         ' color of the text as a RGB Long
    Dim backColor       As Long         ' color of the background as a RGB long
    Dim bFixedSingle    As Boolean      ' whether or not we're doing the pseudo-flat or fixed single style
        
    If hWnd = 0 Then Exit Function
    On Local Error Resume Next
    
    ' Get all of our properties from the combobox. Use our slimy copymemory hack to convert the
    '  object pointer to a combobox object we can reference properties from.
    aPtr = GetProp(hWnd, "ControlPtr")
    If aPtr <> 0 Then
        CopyMemory aCbo, aPtr, 4
        cboStyle = aCbo.Style
        clrTxt = aCbo.foreColor
        clrBack = aCbo.backColor
        CopyMemory aCbo, 0&, 4
    End If
    bFixedSingle = (GetProp(hWnd, "FixedSingle") = 0)
    
    ' Convert all system colors (&H8000000F, etc)... to long RGB equivalents.
    foreColor = TranslateColor(clrTxt)
    backColor = TranslateColor(clrBack)
    
    ' Begin the painting
    Call BeginPaint(hWnd, aPS)
    aDC = aPS.hDC
    
    ' Get the combobox's dimensions
    GetClientRect hWnd, aRECT
    ' Create a back buffer
    backBuffDC = CreateCompatibleDC(aDC)
    backBuffBmp = CreateCompatibleBitmap(aDC, aRECT.Right, aRECT.Bottom)
    DeleteObject SelectObject(backBuffDC, backBuffBmp)

    ' Fill in the control using our custom pattern brush
    aBrush = CreateSolidBrush(backColor)
    FillRect backBuffDC, aRECT, aBrush
    DeleteObject aBrush
  
    ' Draw the rectangle borders around it
    If bFixedSingle Then
        ' Null brush... allows Rectangle API to draw w/o filling in
        aBrush = SelectObject(backBuffDC, GetStockObject(5))
        ' Draw the outer fixed-single border around the entire control
        aPen = SelectObject(backBuffDC, CreatePen(0, 1, foreColor))
        Rectangle backBuffDC, aRECT.Left, aRECT.Top, aRECT.Right, aRECT.Bottom
        ' only draw the inner border if we're not using the "Simple Combo" style
        If cboStyle <> 1 Then
            DeleteObject SelectObject(backBuffDC, aBrush)
            ' Draw a rectangle with the button-face background for our flat "button"
            aBrush = SelectObject(backBuffDC, GetSysColorBrush(15))
            arrowRECT.Right = aRECT.Right
            arrowRECT.Left = aRECT.Right - 13
            arrowRECT.Top = aRECT.Top
            arrowRECT.Bottom = aRECT.Bottom
            Rectangle backBuffDC, arrowRECT.Left, arrowRECT.Top, arrowRECT.Right, arrowRECT.Bottom
        End If
        ' Replace original objects
        DeleteObject SelectObject(backBuffDC, aBrush)
        DeleteObject SelectObject(backBuffDC, aPen)
    Else
        ' We need to draw the 3D stuff our selves. fun fun fun!!!
        DrawEdge backBuffDC, aRECT, BDR_SUNKENOUTER, BF_RECT
        If cboStyle <> 1 Then
            ' Calculate the rectangle to use for our button
            arrowRECT.Right = aRECT.Right - 2
            arrowRECT.Left = aRECT.Right - 15
            arrowRECT.Top = aRECT.Top + 2
            arrowRECT.Bottom = aRECT.Bottom - 2
            ' Make the down arrow button the correct color
            aBrush = GetSysColorBrush(15)
            FillRect backBuffDC, arrowRECT, aBrush
            DeleteObject aBrush
            ' Draw the button edge
            DrawEdge backBuffDC, arrowRECT, BDR_RAISEDINNER, BF_RECT
        End If
    End If
    
    ' Make sure that we're not drawing for the "Simple Combo" style, which
    '  doesn't have a dropdown arrow.
    If cboStyle <> 1 Then
        ' **********************************************************
        '  DRAW IN THE DOWN ARROW.
        ' **********************************************************
        ' Create a font using the "Marlett" face the correct size
        arrowFont = CreateFont(-11, 0, 0, 0, 400, False, False, False, 1, 0, 0, 2, 0, "Marlett")
        ' Select our font into the device context
        SelectObject backBuffDC, arrowFont
        ' Make things the right color
        SetTextColor backBuffDC, IIf(bFixedSingle, foreColor, GetSysColor(8))
        SetBkMode backBuffDC, 1
        ' Here we go... drawing the arrow
        DrawText backBuffDC, "u", 1, arrowRECT, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
        ' Clean up our Marlett font resource
        DeleteObject arrowFont
    End If
    
    ' This little snippet is only required if we're drawing a combobox with
    '  the dropdown list style. This style doesn't have an embedded "Edit" control
    '  to store the item text, so we must draw it in manually.
    If cboStyle = 2 Then
        ' ********************************
        '  DRAW IN THE SELECTED ITEM TEXT.
        ' ********************************
    
        ' See if an item is currently selected.
        curInd = SendMessage(hWnd, CB_GETCURSEL, 0&, 0&)
        If curInd >= 0 Then
            ' Get the selected text for the item.
            aLen = SendMessage(hWnd, CB_GETLBTEXTLEN, curInd, 0&)
            aStr = Space$(aLen)
            SendMessage hWnd, CB_GETLBTEXT, curInd, aStr
            ' Set the bounding box for the text drawing
            aRECT.Left = aRECT.Left + 4
            aRECT.Right = aRECT.Right - (arrowRECT.Right - arrowRECT.Left) - 3
            ' get the combobox's font
            cboFont = SendMessage(hWnd, WM_GETFONT, 0&, 0&)
            ' Use the combobox's font in our back buffer
            origFont = SelectObject(backBuffDC, cboFont)
            ' Make the text the correct color
            SetTextColor backBuffDC, foreColor
            ' Draw in the text
            DrawText backBuffDC, aStr, aLen, aRECT, DT_VCENTER Or DT_LEFT Or DT_SINGLELINE
            ' Replace the original font
            SelectObject backBuffDC, origFont
        End If
    End If
    
    ' BitBlt the section of the control which needs to be updated from our
    '  backbuffer onto the control.
    With aPS.rcPaint
        BitBlt aDC, .Left, .Top, .Right - .Left, .Bottom - .Top, backBuffDC, .Left, .Top, vbSrcCopy
    End With
    
    ' Clean up our back buffer resources.
    DeleteDC backBuffDC
    DeleteObject backBuffBmp
    
    ' Release the control's drawing back to the system
    Call EndPaint(hWnd, aPS)
    
    FullCBOPaint = 0
End Function

Private Function TranslateColor(aColor As OLE_COLOR) As Long
    
    ' ***************************************************
    ' Converts system colors to Long type RGB equivalents
    ' ***************************************************
    Dim newcolor As Long
    OleTranslateColor aColor, 0&, newcolor
    TranslateColor = newcolor
    
End Function
