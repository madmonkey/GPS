Attribute VB_Name = "modCommonProcess"
Option Explicit
Private m_bInDevelopment As Boolean
Private m_Types() As HTE_GPS.GPSConfiguration
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public Function InDevelopment() As Boolean
   ' Debug.Assert code not run in an EXE.  Therefore m_bInDevelopment variable is never set.
   Debug.Assert InDevelopmentHack() = True
   InDevelopment = m_bInDevelopment
End Function

Private Function InDevelopmentHack() As Boolean
   m_bInDevelopment = True
   InDevelopmentHack = m_bInDevelopment
End Function

Public Function messageStatusDesc(ByVal statusCode As HTE_GPS.GPS_MESSAGE_STATUS) As String
Dim sDesc As String
    Select Case statusCode
        Case GPS_MSG_PROCESSED_WARNING
            sDesc = "GPS_MSG_PROCESSED_WARNING"
        Case GPS_MSG_PROCESSED
            sDesc = "GPS_MSG_PROCESSED"
        Case GPS_MSG_PROCESSED_ERROR
            sDesc = "GPS_MSG_PROCESSED_ERROR"
        Case Else
            sDesc = "GPS_MSG_ERROR"
    End Select
    messageStatusDesc = sDesc
End Function

Public Function CurrentTypes(Types() As HTE_GPS.GPSConfiguration)
    m_Types = Types
End Function

Public Function getMessageType(ByVal GPSType As HTE_GPS.GPS_MESSAGING_TYPES) As HTE_GPS.GPSConfiguration
Dim x As Long
    If IsValidArray(m_Types) Then
        For x = LBound(m_Types) To UBound(m_Types)
            If m_Types(x).GPSType = GPSType Then
                getMessageType = m_Types(x)
                Exit Function
            End If
        Next
    End If
End Function

Public Function InferType(ByVal buffer As String) As HTE_GPS.GPS_MESSAGING_TYPES
Dim x As Long
    If IsValidArray(m_Types) Then
        For x = LBound(m_Types) To UBound(m_Types)
            If InStr(1, buffer, m_Types(x).SOM, vbBinaryCompare) > 0 And InStr(1, buffer, m_Types(x).EOM, vbBinaryCompare) > 0 Then
                InferType = m_Types(x).GPSType
                Exit Function
            End If
        Next
    End If
End Function

Public Function IsValidArray(ByRef this As Variant) As Boolean
    If IsArray(this) Then
        IsValidArray = GetArrayDimensions(VarPtrArray(this)) >= 1
    Else
        IsValidArray = False
    End If
End Function
Private Function GetArrayDimensions(ByVal ArrPtr As Long) As Integer
   Dim address As Long
   
   CopyMemory address, ByVal ArrPtr, ByVal 4   'get the address of the SafeArray structure in memory
   If address <> 0 Then 'if there is a dimension, then address will point to the memory address of the array, otherwise the array isn't dimensioned
      CopyMemory GetArrayDimensions, ByVal address, 2 'fill the local variable with the first 2 bytes of the safearray structure. These first 2 bytes contain an integer describing the number of dimensions
   End If

End Function

Private Function VarPtrArray(arr As Variant) As Long

  'Function to get pointer to the array
   CopyMemory VarPtrArray, ByVal VarPtr(arr) + 8, ByVal 4
    
End Function

