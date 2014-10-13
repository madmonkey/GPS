Attribute VB_Name = "basCommon"
Option Explicit
Private m_bInDevelopment As Boolean
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
Public Function getConfigurationPath() As String
Dim sPath As String
Dim myArray As Variant
If Not InDevelopment Then
    getConfigurationPath = App.Path & "\config.xml"
Else
    myArray = Split(App.Path, "\")
    ReDim Preserve myArray(UBound(myArray) - 1)
    getConfigurationPath = Join(myArray, "\") & "\HTE_ClientUtilities\config.xml"
End If
   
End Function
Public Function processorStatusDesc(ByVal statusCode As HTE_GPS.GPS_PROCESSOR_STATUS) As String
Dim sDesc As String
    Select Case statusCode
        Case GPS_STAT_UNINITIALIZED
            sDesc = "GPS_STAT_UNINITIALIZED"
        Case GPS_STAT_INITIALIZED
            sDesc = "GPS_STAT_INITIALIZED"
        Case GPS_STAT_BAD_INTERFACE
            sDesc = "GPS_STAT_BAD_INTERFACE"
        Case GPS_STAT_HOST_UNSUPPORTED
            sDesc = "GPS_STAT_HOST_UNSUPPORTED"
        Case GPS_STAT_ERROR
            sDesc = "GPS_STAT_ERROR"
        Case GPS_STAT_WARNING
            sDesc = "GPS_STAT_WARNING"
        Case GPS_STAT_READYANDWILLING
            sDesc = "GPS_STAT_READYANDWILLING"
        Case Else
            sDesc = "GPS_STAT_UNKNOWN"
    End Select
    processorStatusDesc = sDesc
End Function
Public Function hostStatusDesc(ByVal statusCode As HTE_GPS.GPS_HOST_STATUS) As String
Dim sDesc As String
    Select Case statusCode
        Case GPS_HOST_ERROR
            sDesc = "GPS_HOST_ERROR"
        Case GPS_HOST_WARNING
            sDesc = "GPS_HOST_WARNING"
        Case GPS_HOST_GROOVY
            sDesc = "GPS_HOST_GROOVY"
        Case Else
            sDesc = "GPS_HOST_UNINITIALIZED"
    End Select
    hostStatusDesc = sDesc
End Function
Public Function formatTag(ByVal HexValue As String) As String
Dim i As Long
Dim sRtn As String
    i = 1
    sRtn = vbNullString
    On Error GoTo exit_formatTag
    If Len(HexValue) Mod 2 = 0 Then
        Do While i < Len(HexValue)
            sRtn = sRtn & Chr$("&H" & Mid$(HexValue, i, 2))
            i = i + 2
        Loop
    End If
exit_formatTag:
    formatTag = sRtn
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
Public Function IsValidArray(ByRef this As Variant) As Boolean
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


