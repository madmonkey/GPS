VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPersistPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get Bytes() As Byte()

End Property

Public Property Let Bytes(ByRef vData() As Byte)

End Property

Public Sub Clear()

End Sub

Public Function Handle() As Long

End Function

Public Property Get Height() As Long

End Property

Public Property Get HeightPixels() As Long

End Property

Public Function Object() As Object

End Function

Public Property Set Picture(Picture As StdPicture)
    
End Property

Public Property Get Picture() As StdPicture

End Property
'''Public Property Get Icon() As Long
'''
'''End Property
'''Public Property Get Bitmap() As Long
'''
'''End Property
Public Function PicType() As PictureTypeConstants

End Function

Public Function ReleaseHandle() As Boolean

End Function

Public Function Render(hdc As Long, ByVal x As Long, ByVal y As Long, Optional ByVal WidthPixels As Long, Optional ByVal HeightPixels As Long, Optional ByVal xSrc As Long, Optional ByVal ySrc As Long, Optional ByVal cxSrc As Long, Optional ByVal cySrc As Long, Optional ByVal dwRop As Long) As Long

End Function

Public Function SRCAND() As Long

End Function

Public Function SRCCOPY() As Long

End Function

Public Function SRCPAINT() As Long

End Function

Public Property Get ThrowExceptions() As Boolean

End Property
Public Property Let ThrowExceptions(ByVal vData As Boolean)

End Property

Public Property Get Width() As Long

End Property

Public Property Get WidthPixels() As Long

End Property
