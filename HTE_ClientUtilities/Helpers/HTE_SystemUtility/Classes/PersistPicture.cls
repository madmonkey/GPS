VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PersistPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IPersistPicture

Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0
Private Const CBM_CREATEDIB = &H2
Private Const CBM_INIT = &H4
Private Const cMin = 4 'consists of long value - what TYPE of picture is it?

'Used by OleCreatePictureIndirect
Private Type CLSID
    id(15) As Byte
End Type

'RGBColor Type
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

'Length = 40 bytes ..always use len(bmInfo) for length
Private Type BitmapInfoHeader
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BitmapInfo
    bmiHeader As BitmapInfoHeader
    bmiColors(256) As RGBQUAD
End Type

'BitMap Data follows the BitMapINFOEX structure in m_DataArr Array
Private Type BitmapInfoEX
    bmiTypePicture As PictureTypeConstants
    bmiHeader As BitmapInfoHeader
    bmiColors(256) As RGBQUAD
End Type

Private Type BitMapType
    bmType As Long
    bm_lngWidth As Long
    bm_lngHeight As Long
    bm_lngWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

'Used to load BitMap into StdPicture
Private Type TypeBitmap
    cbSizeOfStruct As Long
    PicType As PictureTypeConstants
    hBitmap As Long
    hPal As Long
End Type

Private Type TypeMetafile
    cbSizeOfStruct As Long
    PicType As PictureTypeConstants
    hmeta As Long
    xExt As Long
    yExt As Long
End Type

Private Type TypeEnhMetafile
    cbSizeOfStruct As Long
    PicType As PictureTypeConstants
    henhmeta As Long
End Type

Private Type IconInfo
    fIcon As Long
    xHotSpot As Long
    yHotSpot As Long
    hbmMask As Long
    hbmColor As Long
End Type

'bmMask data stored after IconInfoEx structure
'LengthbmMask is Length of BitMapMask stored in m_DataArr array
Private Type ICONINFOEX
    PicType As PictureTypeConstants
    LengthbmMask As Long
    IconInfo As IconInfo
End Type

'Used to load Icon into StdPicture
Private Type TypeIcon
    cbSizeOfStruct As Long
    PicType As PictureTypeConstants
    hIcon As Long
End Type

'Windows API used by this class
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateDIBitmapByNum Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As Long, ByVal dwUsage As Long, lpInitBits As Any, ByVal lpInitInfo As Long, ByVal wUsage As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BitmapInfo, ByVal wUsage As Long) As Long
Private Declare Function GetDIBitsByNum Lib "gdi32" Alias "GetDIBits" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByVal lpBits As Long, ByVal lpBI As Long, ByVal wUsage As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function OleCreatePictureIndirect& Lib "oleaut32.dll" (pPicDesc As Any, riid As CLSID, ByVal fown As Long, lpUNK As Object)
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'Icon API's
Private Declare Function CreateIconIndirect Lib "user32" (piconinfo As IconInfo) As Long
Private Declare Function CreateIconIndirectByNum Lib "user32" Alias "CreateIconIndirect" (ByVal piconinfo As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As IconInfo) As Long
Private Declare Function GetIconInfoByNum Lib "user32" Alias "GetIconInfo" (ByVal hIcon As Long, ByVal piconinfo As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
'Private Members
Private m_blnChanged As Boolean
Private m_lngHeight As Long 'height in Hi_Metic
Private m_lngWidth As Long 'width in Hi_Metic
Private m_lngHeightPixels As Long 'height in Pixels
Private m_lngWidthPixels As Long 'width in Pixels
'Used to store BitMap and Icon data
Private m_DataArr() As Byte
'If BitMap
'Byte       Name                 Description
'0-3        PicType         PictureTypeConstants
'4-43      bmiHeader       BitmapINFOHEADER
'44-1067 bmiColors(256)  RGBQUAD
'1068 to UBound(m_DataArr)              BitMap Bytes

'If Icon
'Byte               Name                 Description
'0-3                PicType         PictureTypeConstants
'4-7            LengthbmMask        Length of Mask
'8-27           IconInfo            IconInfo type
'28-28 + lengthbmMask -1              BMMask Data
'28 + lengthbmMask - 1 to ubound(m_DataArr)    Icon MB data
'-------------------------------------------------------------------------------
Private m_lngHandle As Long
Private Const cModuleName = "PersistPicture"
Private m_ThrowExceptions As Boolean

'''Private Property Get IPersistPicture_Bitmap() As Long
'''    If IsValidArray(m_DataArr) Then
'''        IPersistPicture_Bitmap = BitmapFromDataArray(m_DataArr)
'''    End If
'''End Property

Private Property Let IPersistPicture_Bytes(RHS() As Byte)
Dim oPic As StdPicture
    IPersistPicture_Clear
    m_DataArr = RHS
    getPicturePointer oPic
End Property

Private Property Get IPersistPicture_Bytes() As Byte()
    IPersistPicture_Bytes = m_DataArr
End Property

Private Sub IPersistPicture_Clear()
    Clear
End Sub

Private Function IPersistPicture_Handle() As Long
    Dim PicType As Long

On Error GoTo err_IPersistPicture_Handle

    If m_lngHandle <> 0 Then
        IPersistPicture_Handle = m_lngHandle
        Exit Function
    End If
    
    If UBound(m_DataArr) < cMin Then Exit Function
    CopyMemory PicType, ByVal VarPtr(m_DataArr(0)), 4
    Select Case PicType
        Case vbPicTypeBitmap
            m_lngHandle = BitmapFromDataArray(m_DataArr())
            IPersistPicture_Handle = m_lngHandle
            m_blnChanged = True
        Case vbPicTypeIcon
            m_lngHandle = IconHandleFromDataArray(m_DataArr())
            IPersistPicture_Handle = m_lngHandle
            m_blnChanged = True
        Case Else
    End Select
    Exit Function

err_IPersistPicture_Handle:
    ErrRaise vbObjectError + 16, cModuleName & ":IPersistPicture_Handle", Err.Number & ":" & Err.Description
End Function

Private Property Get IPersistPicture_Height() As Long
    IPersistPicture_Height = m_lngHeight
End Property

Private Property Get IPersistPicture_HeightPixels() As Long
     IPersistPicture_HeightPixels = m_lngHeightPixels
End Property

'''Private Property Get IPersistPicture_Icon() As Long
'''If IsValidArray(m_DataArr) Then
'''    IPersistPicture_Icon = IconHandleFromDataArray(m_DataArr)
'''End If
'''End Property

Private Function IPersistPicture_Object() As Object
    Set IPersistPicture_Object = Me
End Function

Private Property Set IPersistPicture_Picture(Picture As IPictureDisp)
On Error GoTo err_IPersistPicture_Picture
    If Picture Is Nothing Then
        IPersistPicture_Clear
    Else
        Select Case Picture.Type
            Case vbPicTypeBitmap
                m_DataArr = BitmapToByteArray(Picture.Handle)
                If Not IsValidArray(m_DataArr) Then
                    ErrRaise vbObjectError + 20, cModuleName & ":IPersistPicture_Picture", "Invalid [BitMap/JPEG] array."
                End If
            
            Case vbPicTypeIcon
                m_DataArr = IconToByteArray(Picture.Handle)
                If Not IsValidArray(m_DataArr) Then
                    ErrRaise vbObjectError + 20, cModuleName & ":IPersistPicture_Picture", "Invalid [Icon] array."
                End If
            
            Case Else 'vbPicTypeEMetafile,vbPicTypeMetafile, vbPicTypeNone
                ErrRaise vbObjectError + 19, cModuleName & ":IPersistPicture_Picture", "Unsupported format."
        
        End Select
    End If
    Exit Property
err_IPersistPicture_Picture:
    ErrRaise vbObjectError + 18, cModuleName & ":IPersistPicture_Picture", Err.Number & ":" & Err.Description
End Property
Private Function getPicturePointer(Picture As IPictureDisp) As Boolean

    Dim hHandle As Long
    Dim pictureType As Long
    Dim picConst As PictureTypeConstants
    Dim Ret As Long

On Error GoTo err_getPicturePointer

    If IsValidArray(m_DataArr) Then
        If UBound(m_DataArr) - 1 > 8 Then
            CopyMemory pictureType, ByVal VarPtr(m_DataArr(0)), 4
            picConst = pictureType
            
            Select Case picConst
                
                Case vbPicTypeBitmap
                    hHandle = BitmapFromDataArray(m_DataArr())
                    If hHandle Then
                        If CreatePictureFromBitmap(hHandle, 0, Picture) Then
                            getPicturePointer = True
                        End If
                        Call DeleteObject(hHandle)
                    End If
        
                Case vbPicTypeIcon
                    hHandle = IconHandleFromDataArray(m_DataArr())
                    If hHandle <> 0 Then
                        If CreatePictureFromhIcon(hHandle, 0, Picture) Then
                            getPicturePointer = True
                        End If
                        Call DestroyIcon(hHandle)
                    End If
        
                Case vbPicTypeEMetafile, vbPicTypeMetafile
                    Exit Function
                    
                Case Else
                    Set Picture = New StdPicture
                    getPicturePointer = True
            End Select
            
            Exit Function
        End If
    End If
    getPicturePointer = False
    Exit Function
err_getPicturePointer:
    ErrRaise vbObjectError + 6, cModuleName & ":getPicturePointer", Err.Number & ":" & Err.Description

End Function
Private Property Get IPersistPicture_Picture() As stdole.Picture

On Error GoTo err_IPersistPicture_Picture
    
    getPicturePointer IPersistPicture_Picture
    Exit Property

err_IPersistPicture_Picture:
    ErrRaise vbObjectError + 5, cModuleName & ":IPersistPicture_Picture", Err.Number & ":" & Err.Description
End Property

Private Function IPersistPicture_PicType() As PictureTypeConstants
On Error GoTo err_IPersistPicture_PicType
Dim bmInfoEX As BitmapInfoEX
Dim lPicType As Long
        If IsValidArray(m_DataArr) Then
            If UBound(m_DataArr) - 1 > cMin Then
                CopyMemory lPicType, ByVal VarPtr(m_DataArr(0)), 4
                IPersistPicture_PicType = lPicType
                Exit Function
            End If
        End If
err_IPersistPicture_PicType:
    ErrRaise vbObjectError + 17, cModuleName & ":IPersistPicture_PicType", Err.Number & ":" & Err.Description
End Function

Private Function IPersistPicture_ReleaseHandle() As Boolean
On Error GoTo err_IPersistPicture_ReleaseHandle
Dim PicType As Long
    If m_lngHandle <> 0 Then
        If UBound(m_DataArr) - 1 < cMin Then Exit Function
        CopyMemory PicType, ByVal VarPtr(m_DataArr(0)), 4
        Select Case PicType
            Case vbPicTypeBitmap
                Call DeleteObject(m_lngHandle)
            Case vbPicTypeIcon
                Call DestroyIcon(m_lngHandle)
            Case Else
                Exit Function
        End Select
        m_lngHandle = 0
        m_lngHeightPixels = 0
        m_lngWidthPixels = 0
        m_lngHeight = 0
        m_lngWidth = 0
        m_blnChanged = True
        IPersistPicture_ReleaseHandle = True
    End If
    Exit Function
err_IPersistPicture_ReleaseHandle:
    ErrRaise vbObjectError + 2, cModuleName & ":IPersistPicture_ReleaseHandle", Err.Number & ":" & Err.Description
End Function

Private Function IPersistPicture_Render(hdc As Long, ByVal x As Long, ByVal y As Long, Optional ByVal WidthPixels As Long, Optional ByVal HeightPixels As Long, Optional ByVal xSrc As Long, Optional ByVal ySrc As Long, Optional ByVal cxSrc As Long, Optional ByVal cySrc As Long, Optional ByVal dwRop As Long) As Long

    Dim srcDC As Long, oldDC As Long, Ret As Long
    Dim PicType As Long, hHandle As Long
    
On Error GoTo err_IPersistPicture_Render
    
    If hdc = 0 Then Exit Function
    If m_lngWidthPixels <= 0 Or m_lngHeightPixels <= 0 Then Exit Function
    If UBound(m_DataArr) - 1 < cMin Then Exit Function
    If m_lngHandle <> 0 Then Exit Function

    IPersistPicture_Render = True
    CopyMemory PicType, ByVal VarPtr(m_DataArr(0)), 4
    Select Case PicType
        Case vbPicTypeBitmap
            hHandle = BitmapFromDataArray(m_DataArr())
            If hHandle = 0 Then Exit Function
            If dwRop = 0 Then dwRop = IPersistPicture_SRCCOPY
            If cxSrc = 0 Then cxSrc = m_lngWidthPixels
            If cySrc = 0 Then cySrc = m_lngHeightPixels
            If WidthPixels = 0 Then WidthPixels = m_lngWidthPixels
            If HeightPixels = 0 Then HeightPixels = m_lngHeightPixels
            If hHandle <> 0 Then
                RenderBitmap hHandle, hdc, x, y, WidthPixels, HeightPixels, xSrc, ySrc, cxSrc, cySrc, dwRop
                Call DeleteObject(hHandle)
            End If

        Case vbPicTypeIcon
            If m_lngHandle <> 0 Then
                IPersistPicture_Render = DrawIcon(hdc, x, y, m_lngHandle)
            Else
                If WidthPixels = 0 Then
                    hHandle = IconHandleFromDataArray(m_DataArr())
                    If hHandle <> 0 Then
                        IPersistPicture_Render = DrawIcon(hdc, x, y, hHandle)
                        Call DestroyIcon(hHandle)
                    End If
                Else
                    If cxSrc = 0 Then cxSrc = m_lngWidthPixels
                    If cySrc = 0 Then cySrc = m_lngHeightPixels
                    If WidthPixels = 0 Then WidthPixels = m_lngWidthPixels
                    If HeightPixels = 0 Then HeightPixels = m_lngHeightPixels
                    'Paint Mask on hDC
                    hHandle = BitmapMaskHandleFromIconDataArray(m_DataArr)
                    If hHandle <> 0 Then
                        RenderBitmap hHandle, hdc, x, y, WidthPixels, HeightPixels, xSrc, ySrc, cxSrc, cySrc, IPersistPicture_SRCAND
                        Call DeleteObject(hHandle)
                    End If
                    'Paint BitMap on hDC
                    hHandle = BitmapHandleFromIconDataArray(m_DataArr)
                    If hHandle <> 0 Then
                        Call RenderBitmap(hHandle, hdc, x, y, WidthPixels, HeightPixels, xSrc, ySrc, cxSrc, cySrc, IPersistPicture_SRCPAINT)
                        'Delete resource hHandle
                        Call DeleteObject(hHandle)
                    End If
                End If
            End If
        Case Else
            IPersistPicture_Render = 0
    End Select
    Exit Function
err_IPersistPicture_Render:
    IPersistPicture_Render = 0
    ErrRaise vbObjectError + 9, cModuleName & ":IPersistPicture_Render", Err.Number & ":" & Err.Description
End Function
Private Function RenderBitmap(hHandle As Long, hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidthPixels As Long, ByVal nHeightPixels As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal cxSrc As Long, ByVal cySrc As Long, ByVal dwRop As Long) As Long
On Error GoTo err_RenderBitmap
Dim srcDC As Long
Dim oldDC As Long
Dim Ret As Long
Dim PicType As Long

    If hdc = 0 Then Exit Function
    If hHandle = 0 Then Exit Function
    If UBound(m_DataArr) - 1 < cMin Then Exit Function

    'Get device context - set up source
    srcDC = CreateCompatibleDC(0)
    If srcDC <> 0 Then
        oldDC = SelectObject(srcDC, hHandle)
        If oldDC <> 0 Then
            'Now do the copy
            Ret = StretchBlt(hdc, x, y, nWidthPixels, nHeightPixels, srcDC, xSrc, ySrc, cxSrc, cySrc, dwRop)
            'Clean up the device contexts
            Call SelectObject(srcDC, oldDC)
        End If
        Call DeleteDC(srcDC)
    End If
    Exit Function
err_RenderBitmap:
    ErrRaise vbObjectError + 10, cModuleName & ":RenderBitmap", Err.Number & ":" & Err.Description
End Function
Private Function IPersistPicture_SRCAND() As Long
    IPersistPicture_SRCAND = &H8800C6
End Function

Private Function IPersistPicture_SRCCOPY() As Long
    'destination memory = source memory
    IPersistPicture_SRCCOPY = &HCC0020
End Function

Private Function IPersistPicture_SRCPAINT() As Long
    IPersistPicture_SRCPAINT = &HEE0086
End Function

Private Property Let IPersistPicture_ThrowExceptions(ByVal RHS As Boolean)
    m_ThrowExceptions = RHS
End Property

Private Property Get IPersistPicture_ThrowExceptions() As Boolean
    IPersistPicture_ThrowExceptions = m_ThrowExceptions
End Property

Private Property Get IPersistPicture_Width() As Long
    IPersistPicture_Width = m_lngWidth
End Property

Private Property Get IPersistPicture_WidthPixels() As Long
    IPersistPicture_WidthPixels = m_lngWidthPixels
End Property

Private Function BitmapToByteArray(hBitmap As Long) As Byte()
    Dim lReturn As Long, lDC As Long
    Dim m_tBI As BitmapInfoEX
    Dim bm As BitMapType
    Dim bBytes() As Byte

On Error GoTo err_BitmapToBytes
    
    If hBitmap <> 0 Then
        If GetObjectAPI(hBitmap, Len(bm), bm) <> 0 Then
            ReDim bBytes(0 To 0) As Byte
            IPersistPicture_ReleaseHandle
            m_tBI.bmiTypePicture = vbPicTypeBitmap
            With m_tBI.bmiHeader
                .biSize = Len(m_tBI.bmiHeader)
                .biWidth = bm.bm_lngWidth
                .biHeight = bm.bm_lngHeight
                m_lngHeightPixels = bm.bm_lngHeight
                m_lngWidthPixels = bm.bm_lngWidth
                m_lngWidth = ScaleX(m_lngWidthPixels) 'ScaleX(m_lngWidthPixels, False)
                m_lngHeight = ScaleY(m_lngHeightPixels)
                .biPlanes = bm.bmPlanes
                .biBitCount = bm.bmBitsPixel
                .biCompression = BI_RGB
                lDC = CreateCompatibleDC(0)
                If lDC <> 0 Then
                    lReturn = GetDIBitsByNum(lDC, hBitmap, 0, bm.bm_lngHeight, ByVal 0&, ByVal VarPtr(m_tBI.bmiHeader), DIB_RGB_COLORS)
                    If lReturn Then
                        ReDim bBytes(0 To .biSizeImage - 1 + Len(m_tBI))
                        CopyMemory ByVal VarPtr(bBytes(0)), m_tBI, Len(m_tBI)
                        lReturn = GetDIBitsByNum(lDC, hBitmap, 0, bm.bm_lngHeight, VarPtr(bBytes(Len(m_tBI))), VarPtr(bBytes(4)), DIB_RGB_COLORS)
                        BitmapToByteArray = bBytes
                    End If
                    Call DeleteDC(lDC)
                End If
            End With
            m_blnChanged = True
        End If
    End If
    
    Exit Function
    
err_BitmapToBytes:
    ErrRaise vbObjectError + 3, cModuleName & ":BitmapToBytes", Err.Number & ":" & Err.Description
End Function

Private Function BitmapFromDataArray(bBytes() As Byte) As Long
    
    Dim bmInfoEX As BitmapInfoEX
    Dim lDC As Long

On Error GoTo err_BitmapFromDataArray

    If UBound(bBytes) - 1 > Len(bmInfoEX) Then
        lDC = CreateCompatibleDC(0)
        If lDC <> 0 Then
            CopyMemory bmInfoEX, bBytes(0), Len(bmInfoEX)
            m_lngHeight = bmInfoEX.bmiHeader.biHeight
            m_lngHeightPixels = ScaleY(bmInfoEX.bmiHeader.biHeight)
            m_lngWidthPixels = bmInfoEX.bmiHeader.biWidth
            m_lngWidth = ScaleX(bmInfoEX.bmiHeader.biWidth)
            BitmapFromDataArray = CreateDIBitmapByNum(lDC, ByVal VarPtr(bBytes(Len(bmInfoEX.bmiTypePicture))), CBM_INIT Or CBM_CREATEDIB, ByVal VarPtr(bBytes(Len(bmInfoEX))), ByVal VarPtr(bBytes(4)), DIB_RGB_COLORS)
            Call DeleteDC(lDC)
        End If
    End If
    Exit Function
err_BitmapFromDataArray:
    ErrRaise vbObjectError + 8, cModuleName & ":BitmapFromDataArray", Err.Number & ":" & Err.Description
End Function

Private Function CreatePictureFromBitmap(ByVal hBitmap As Long, hPal As Long, Picture As IPictureDisp) As Boolean
On Error GoTo err_CreatePictureFromBitmap
Dim Ret As Long
Dim m_tB As TypeBitmap
Dim lpUNK As IUnknown
Dim m_UNKNOWN As CLSID
        
        If hBitmap <> 0 Then
            With m_UNKNOWN
                .id(15) = &H46
                .id(8) = &HC0
            End With
            m_tB.cbSizeOfStruct = Len(m_tB)
            m_tB.PicType = vbPicTypeBitmap
            'Need to copy Bitmaps just in case the bHandle has been selected into a Device Context.
            'Windows only permits the Bitmap to be loaded in one device at a time
            m_tB.hBitmap = CopyBitmap(hBitmap)
            If m_tB.hBitmap Then
                m_tB.hPal = hPal
                m_tB.PicType = vbPicTypeBitmap
                Ret = OleCreatePictureIndirect(m_tB, m_UNKNOWN, 1, lpUNK)
                If Ret = 0 Then
                    Set Picture = lpUNK
                    CreatePictureFromBitmap = True
                End If
            End If
        End If
        Exit Function

err_CreatePictureFromBitmap:
    ErrRaise vbObjectError + 4, cModuleName & ":CreatePictureFromBitmap", Err.Number & ":" & Err.Description
End Function

Private Function CopyBitmap(ByVal hBitmap As Long) As Long
On Error GoTo err_CopyBitmap
Dim bm As BitMapType
Dim newBM As Long
Dim srcDC As Long, DestDC As Long
Dim oldbmsrc As Long, oldbmdest As Long

    If GetObjectAPI(hBitmap, Len(bm), bm) = 0 Then Exit Function
    bm.bmBits = 0
    srcDC = CreateCompatibleDC(0)
    If srcDC <> 0 Then
        oldbmsrc = SelectObject(srcDC, hBitmap)
        If oldbmsrc <> 0 Then
            DestDC = CreateCompatibleDC(0)
            If DestDC <> 0 Then
                newBM = CreateCompatibleBitmap(srcDC, bm.bm_lngWidth, bm.bm_lngHeight)
                If newBM Then
                    oldbmdest = SelectObject(DestDC, newBM)
                    Call BitBlt(DestDC, 0, 0, bm.bm_lngWidth, bm.bm_lngHeight, srcDC, 0, 0, IPersistPicture_SRCCOPY)
                End If
                Call SelectObject(DestDC, oldbmdest)
                Call DeleteDC(DestDC)
            End If
            Call SelectObject(srcDC, oldbmsrc)
        End If
        Call DeleteDC(srcDC)
        CopyBitmap = newBM
    End If
    Exit Function
err_CopyBitmap:
    ErrRaise vbObjectError + 11, cModuleName & ":CopyBitmap", Err.Number & ":" & Err.Description
End Function
Private Function IconToByteArray(hIcon As Long) As Byte()
    
    Dim iInfo As ICONINFOEX
    Dim Buffer() As Byte, bBytes() As Byte
    Dim length As Long
    Dim Ret As Long, LenInfo As Long
    Dim bm As BitMapType

On Error GoTo err_IconToByteArray


        If hIcon = 0 Then Exit Function
        ReDim bBytes(0 To 0) As Byte
        If m_lngHandle <> 0 Then IPersistPicture_ReleaseHandle

        LenInfo = Len(iInfo)
        If GetIconInfoByNum(hIcon, ByVal VarPtr(iInfo.IconInfo)) <> 0 Then
            iInfo.PicType = vbPicTypeIcon
            Buffer = BitmapToByteArray(iInfo.IconInfo.hbmMask)
            If IsValidArray(Buffer) Then
            'If BitmapToBytes(iInfo.IconInfo.hbmMask, Buffer) Then
                length = UBound(Buffer) + 1
                iInfo.LengthbmMask = length
                ReDim Preserve bBytes(0 To (LenInfo + length - 1)) As Byte
                iInfo.PicType = vbPicTypeIcon 'need to reset because we know better
                CopyMemory ByVal VarPtr(bBytes(0)), iInfo, LenInfo
                If length > 0 Then
                    CopyMemory ByVal VarPtr(bBytes(Len(iInfo))), ByVal VarPtr(Buffer(0)), length
                    Erase Buffer
                    Buffer = BitmapToByteArray(iInfo.IconInfo.hbmColor)
                    If IsValidArray(Buffer) Then
                    'If BitmapToBytes(iInfo.IconInfo.hbmColor, Buffer) Then
                        length = UBound(Buffer) + 1
                        If GetObjectAPI(iInfo.IconInfo.hbmColor, Len(bm), bm) <> 0 Then
                            m_lngHeightPixels = bm.bm_lngHeight
                            m_lngWidthPixels = bm.bm_lngWidth
                            m_lngWidth = ScaleX(m_lngWidthPixels, False)
                            m_lngHeight = ScaleY(m_lngHeightPixels, False)
                        End If
                        ReDim Preserve bBytes(0 To (LenInfo + iInfo.LengthbmMask + length)) As Byte
                        CopyMemory ByVal (VarPtr(bBytes(LenInfo)) + iInfo.LengthbmMask), ByVal VarPtr(Buffer(0)), length
                        IconToByteArray = bBytes
                    Else
                        ReDim bBytes(0 To 0) As Byte
                    End If
                End If
            End If
        End If
        'Must destroy icons returned as these are use User32 resources owned by the operating system
        If iInfo.IconInfo.hbmMask Then Ret = DeleteObject(iInfo.IconInfo.hbmMask)
        If iInfo.IconInfo.hbmColor Then Ret = DeleteObject(iInfo.IconInfo.hbmColor)
        m_blnChanged = True
        Exit Function
err_IconToByteArray:
    ErrRaise vbObjectError + 7, cModuleName & ":IconToByteArray", Err.Number & ":" & Err.Description
End Function

Private Function IconHandleFromDataArray(Bytes() As Byte) As Long

On Error GoTo err_IconHandleFromDataArray
Dim iInfo As ICONINFOEX
Dim Ret As Long
Dim Buffer() As Byte
Dim LenInfo As Long
Dim newBM As Long

    LenInfo = Len(iInfo)
    If UBound(Bytes) > LenInfo - 1 Then
        CopyMemory iInfo, ByVal VarPtr(Bytes(0)), LenInfo
        ReDim Buffer(0 To iInfo.LengthbmMask - 1) As Byte
        CopyMemory ByVal VarPtr(Buffer(0)), ByVal VarPtr(Bytes(LenInfo)), iInfo.LengthbmMask
        newBM = BitmapFromDataArray(Buffer())
        Erase Buffer
        If newBM Then
            iInfo.IconInfo.hbmMask = newBM
            ReDim Buffer(0 To UBound(Bytes) - iInfo.LengthbmMask) As Byte
            CopyMemory ByVal VarPtr(Buffer(0)), ByVal VarPtr(Bytes(iInfo.LengthbmMask)) + LenInfo, UBound(Bytes()) - iInfo.LengthbmMask - LenInfo
            newBM = BitmapFromDataArray(Buffer())
            Erase Buffer
            If newBM Then
                iInfo.IconInfo.hbmColor = newBM
                IconHandleFromDataArray = CreateIconIndirect(iInfo.IconInfo)
            End If
            'Must delete Bitmaps otherwise memory leaks
            If iInfo.IconInfo.hbmMask Then Ret = DeleteObject(iInfo.IconInfo.hbmMask)
            If iInfo.IconInfo.hbmColor Then Ret = DeleteObject(iInfo.IconInfo.hbmColor)
        End If
    End If
    Exit Function
err_IconHandleFromDataArray:
    ErrRaise vbObjectError + 12, cModuleName & ":IconHandleFromDataArray", Err.Number & ":" & Err.Description
End Function
Private Function CreatePictureFromhIcon(ByVal hIcon As Long, hPal As Long, Picture As IPictureDisp) As Boolean
On Error GoTo err_CreatePictureFromhIcon
Dim Ret As Long
Dim newIcon As TypeIcon
Dim lpUNK As IUnknown
Dim m_UNKNOWN As CLSID
        
        With m_UNKNOWN
            .id(15) = &H46
            .id(8) = &HC0
        End With

        If hIcon = 0 Then Exit Function
        newIcon.cbSizeOfStruct = Len(newIcon)
        newIcon.PicType = vbPicTypeIcon
        newIcon.hIcon = hIcon
        If hIcon Then
            Ret = OleCreatePictureIndirect(newIcon, m_UNKNOWN, 1, lpUNK)
            If Ret = 0 Then
                Set Picture = lpUNK
                CreatePictureFromhIcon = True
            End If
        End If
        Exit Function
err_CreatePictureFromhIcon:
    ErrRaise vbObjectError + 13, cModuleName & ":CreatePictureFromhIcon", Err.Number & ":" & Err.Description
End Function
Private Function BitmapHandleFromIconDataArray(Bytes() As Byte) As Long

On Error GoTo err_BitmapHandleFromIconDataArray
Dim iInfo As ICONINFOEX
Dim Ret As Long
Dim Buffer() As Byte
Dim LenInfo As Long
Dim newBM As Long

    LenInfo = Len(iInfo)
    If UBound(Bytes) > LenInfo - 1 Then
        CopyMemory iInfo, ByVal VarPtr(Bytes(0)), LenInfo
        ReDim Buffer(0 To UBound(Bytes) - iInfo.LengthbmMask) As Byte
        CopyMemory ByVal VarPtr(Buffer(0)), ByVal VarPtr(Bytes(iInfo.LengthbmMask)) + LenInfo, UBound(Bytes()) - iInfo.LengthbmMask - LenInfo
        newBM = BitmapFromDataArray(Buffer())
        Erase Buffer
        If newBM Then
            BitmapHandleFromIconDataArray = newBM
        End If
    End If
    Exit Function
err_BitmapHandleFromIconDataArray:
    ErrRaise vbObjectError + 14, cModuleName & ":BitmapHandleFromIconDataArray", Err.Number & ":" & Err.Description
End Function
Private Function BitmapMaskHandleFromIconDataArray(Bytes() As Byte) As Long

On Error GoTo err_BitmapMaskHandleFromIconDataArray
Dim iInfo As ICONINFOEX
Dim Ret As Long
Dim Buffer() As Byte
Dim LenInfo As Long
Dim newBM As Long

    LenInfo = Len(iInfo)
    If UBound(Bytes) > LenInfo - 1 Then
        CopyMemory iInfo, ByVal VarPtr(Bytes(0)), LenInfo
        ReDim Buffer(0 To iInfo.LengthbmMask - 1) As Byte
        CopyMemory ByVal VarPtr(Buffer(0)), ByVal VarPtr(Bytes(LenInfo)), iInfo.LengthbmMask
        BitmapMaskHandleFromIconDataArray = BitmapFromDataArray(Buffer())
        Erase Buffer
    End If
    Exit Function
err_BitmapMaskHandleFromIconDataArray:
    ErrRaise vbObjectError + 15, cModuleName & ":BitmapMaskHandleFromIconDataArray", Err.Number & ":" & Err.Description
End Function
Private Sub Clear()
    If m_lngHandle > 0 Then IPersistPicture_ReleaseHandle
    m_blnChanged = False
End Sub

'***************************************************************
'Convert Pixels to HiMetric and visa versa
'True  - convert from  HiMetric to Pixels
'False - convert from  Pixels to HiMetric
'***************************************************************
Private Function ScaleX(ByVal dbl As Double, Optional ByVal bToPixels As Boolean = False) As Double
Dim x As Long
Dim dblValue As Double
    
    dblValue = (14.4 / 25.4)
    x = Screen.TwipsPerPixelX
    If bToPixels Then 'To Pixels
        ScaleX = (dbl * dblValue) \ x
    Else 'To HiMetric
        ScaleX = CLng((dbl / dblValue) * x)
    End If
End Function

Private Function ScaleY(ByVal dbl As Double, Optional ByVal bToPixels As Boolean = False) As Double
Dim y As Long
Dim dblValue As Double
    y = Screen.TwipsPerPixelY
    dblValue = (14.4 / 25.4)
    If bToPixels Then 'To Pixels
        ScaleY = (dbl * dblValue) \ y
    Else 'To HiMetric
        ScaleY = CLng((dbl / dblValue) * y)
    End If
End Function
Private Sub ErrRaise(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
    If IPersistPicture_ThrowExceptions Then
        Err.Raise Number, Source, Description
    End If
End Sub
Private Function IsValidArray(ByRef this As Variant) As Boolean
    If IsArray(this) Then
        IsValidArray = GetArrayDimensions(VarPtrArray(this)) >= 1
    Else
        IsValidArray = False
    End If
End Function

Private Function GetArrayDimensions(ByVal arrPtr As Long) As Integer
   Dim address As Long
   
   CopyMemory address, ByVal arrPtr, ByVal 4   'get the address of the SafeArray structure in memory
   If address <> 0 Then 'if there is a dimension, then address will point to the memory address of the array, otherwise the array isn't dimensioned
      CopyMemory GetArrayDimensions, ByVal address, 2 'fill the local variable with the first 2 bytes of the safearray structure. These first 2 bytes contain an integer describing the number of dimensions
   End If

End Function

Private Function VarPtrArray(arr As Variant) As Long

  'Function to get pointer to the array
   CopyMemory VarPtrArray, ByVal VarPtr(arr) + 8, ByVal 4
    
End Function
