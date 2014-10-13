VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IStack

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private m_lItem() As Long
Private m_lCount As Long
Private m_lArraySize As Long
Private m_lAllocationSize As Long

Private Property Get AllocationSize() As Long
   AllocationSize = m_lAllocationSize
End Property

Private Property Let AllocationSize(ByVal lAllocationSize As Long)
   m_lAllocationSize = lAllocationSize
End Property

Private Property Get Item(ByVal Index As Long) As Object
   ' Return the item at Index.  Note since this
   ' array does not automatically reduce size we
   ' need to check that index is in bounds:
   If Index > 0 And Index <= m_lCount Then
      Set Item = ObjectFromPtr(m_lItem(Index))
   Else
      ' Subscript out of range
      Err.Raise 9
   End If
End Property

Private Property Let Item(ByVal Index As Long, ByVal Value As Object)
Dim iU As IShellFolderEx_TLB.IUnknown

   If Index > 0 And Index <= m_lCount Then
      If (m_lItem(Index) <> 0) Then
         Set iU = ObjectFromPtr(m_lItem(Index))
         iU.Release
         Set iU = Nothing
      End If
   
      ' Add a reference to the cData object so it
      ' won't be freed until the list item
      ' is removed:
      Set iU = Value
      iU.AddRef
      Set iU = Nothing
      m_lItem(Index) = ObjPtr(Value)
   Else
      ' Subscript out of range
      Err.Raise 9
   End If
End Property

Private Sub Add(ByVal Value As Object, Optional ByVal Index As Long = -1)
Dim j As Long
Dim lTemp As Long
Dim iU As IShellFolderEx_TLB.IUnknown

   If Index <= 0 Or Index > m_lCount Then
      ' Add to end
      m_lCount = m_lCount + 1
      If m_lCount > m_lArraySize Then
         m_lArraySize = m_lCount
         ReDim Preserve m_lItem(1 To m_lArraySize) As Long
      End If
      Set iU = Value
      iU.AddRef
      Set iU = Nothing
      m_lItem(m_lCount) = ObjPtr(Value)
   Else
      lTemp = m_lItem(m_lCount)
      ' Increase array size as required:
      m_lCount = m_lCount + 1
      If m_lCount > m_lArraySize Then
         m_lArraySize = m_lArraySize + m_lAllocationSize
         ReDim Preserve m_lItem(1 To m_lArraySize) As Long
      End If
      ' Shift up from Index to m_lCount-1
      CopyMemory m_lItem(Index + 1), m_lItem(Index), (m_lCount - Index) * 4
      ' Increase array size:
      ' Set item at Index
      Set iU = Value
      iU.AddRef
      Set iU = Nothing
      m_lItem(Index) = ObjPtr(Value)
      ' Set last item:
      m_lItem(m_lCount) = lTemp
      
   End If
   
End Sub

Private Sub Remove(Optional ByVal Index As Long = -1)
Dim j As Long
   ' Reduce size:
   If Index <= 0 Or Index >= m_lCount Then
      If (m_lCount > 0) Then
         ' Call the object's IUnknown_Release method
         ' using the TLB
         Dim iU As IShellFolderEx_TLB.IUnknown
         Set iU = ObjectFromPtr(m_lItem(m_lCount))
         iU.Release
         Set iU = Nothing
      End If
      m_lCount = m_lCount - 1
      If m_lCount < 0 Then
         m_lCount = 0
      End If
   Else
      If (Index > 0) And (Index <= m_lCount) Then
         ' Call the object's IUnknown_Release method
         ' using the TLB
         Set iU = ObjectFromPtr(m_lItem(Index))
         iU.Release
         Set iU = Nothing
      End If
      m_lCount = m_lCount - 1
      If m_lCount > 0 Then
         ' shift down from Index to m_lCount-1
         CopyMemory m_lItem(Index), m_lItem(Index + 1), (m_lCount - Index + 1) * 4
      Else
         ' the end
         m_lCount = 0
      End If
   End If
   
End Sub

Private Sub HeapMinimize()
   ' Reduce the array storage size to
   ' match the number of items in it:
   If m_lArraySize > m_lCount Then
      If m_lCount <= 0 Then
         Erase m_lItem
         m_lArraySize = 0
      Else
         ReDim Preserve m_lItem(1 To m_lCount) As Long
         m_lArraySize = m_lCount
      End If
   End If
End Sub

Private Property Get Count() As Long
   ' Number of items in the array:
   Count = m_lCount
End Property

Private Property Get Exists(ByVal Index As Long)
   ' Does the item at Index exist?
   Exists = (Index > 0 And Index <= m_lCount)
End Property

Private Sub Clear()
   ' Release any object pointers:
   Dim i As Long
   Dim iU As IShellFolderEx_TLB.IUnknown

   For i = 1 To m_lCount
      ' Call the object's IUnknown_Release method
      ' using the TLB
      Set iU = ObjectFromPtr(m_lItem(i))
      iU.Release
      Set iU = Nothing
   Next i
   
   ' Clear down the array:
   m_lCount = 0
End Sub

Private Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
Dim objT As Object
   ' Bruce McKinney's code for getting an Object from the object pointer:
   CopyMemory objT, lPtr, 4
   Set ObjectFromPtr = objT
   CopyMemory objT, 0&, 4
End Property

Private Sub Class_Initialize()
   m_lAllocationSize = 1
End Sub

Private Sub Class_Terminate()
   Clear
End Sub

Private Property Get IStack_Count() As Long
    IStack_Count = Count
End Property

Private Function IStack_Pop() As HTE_GPS.GPSMessage
    Set IStack_Pop = Item(1)
    Remove 1
End Function

Private Sub IStack_Push(msgObj As HTE_GPS.GPSMessage)
    Add msgObj
End Sub

