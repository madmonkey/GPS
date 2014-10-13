Attribute VB_Name = "modHexDump"
Option Explicit

Private Const TIME_ZONE_ID_DAYLIGHT As Long = 2
 
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
 
Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(63) As Byte
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(63) As Byte
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

Private Function GetClientTimeZoneInformation() As String
    Dim utTZ As TIME_ZONE_INFORMATION
    Dim dwBias As Long
    Dim sReturn As String
    
    Select Case GetTimeZoneInformation(utTZ)
        Case TIME_ZONE_ID_DAYLIGHT:
            dwBias = utTZ.Bias + utTZ.DaylightBias
        Case Else
            dwBias = utTZ.Bias + utTZ.StandardBias
    End Select
    sReturn = Format$(-dwBias \ 60, "00") & Format$(Abs(dwBias - (dwBias \ 60) * 60), "00")
    If InStr(sReturn, "-") = 0 Then
        sReturn = "+" & sReturn
    End If
    GetClientTimeZoneInformation = sReturn
    
End Function
Private Function getFullTimeFormat() As String
Dim myTime As SYSTEMTIME
Dim chk As Integer
    GetSystemTime myTime
    With myTime
        chk = myTime.wHour + (Val(GetClientTimeZoneInformation) / 100)
        Select Case True
            Case chk > 0
                myTime.wHour = chk
            Case chk < 0
                myTime.wHour = 24 + chk
            Case Else
                myTime.wHour = 0
        End Select
        getFullTimeFormat = Format$(.wHour, "0#") & ":" & Format$(.wMinute, "0#") & ":" & Format$(.wSecond, "0#") & "." & Format(.wMilliseconds, "00#")
    End With
End Function

Public Function HexDump(ByVal lpBuffer As Long, ByVal nBytes As Long) As String
   Dim i As Long, j As Long
   Dim ba() As Byte
   Dim sRet As stringBuilder
   Dim dBytes As Long
   Set sRet = New stringBuilder
   ' Size recieving buffer as requested, then sling memory block to buffer.
   ReDim ba(0 To nBytes - 1) As Byte
   Call CopyMemory(ba(0), ByVal lpBuffer, nBytes)
   sRet.Append " Total Bytes = " & nBytes 'String(85, "=") & vbCrLf & "Total Bytes = " & nBytes ' & "lpBuffer = &h" & Hex$(lpBuffer)
   ' Buffer may well not be even multiple of 16. If not, we need to round up.
   If nBytes Mod 16 = 0 Then
      dBytes = nBytes
   Else
      dBytes = ((nBytes \ 16) + 1) * 16
   End If
   
   ' Loop through buffer, displaying 16 bytes per row. Preface with offset, trail with ASCII.
   For i = 0 To (dBytes - 1)
      ' Add address and offset from beginning if at the start of new row.
      If (i Mod 16) = 0 Then
         sRet.Append vbCrLf & Space$(1) & getFullTimeFormat & "  " & Right$("000" & (i \ 16), 3) & "  "
      End If
      
      ' Append this byte.
      If i < nBytes Then
         sRet.Append Right$("0" & Hex(ba(i)), 2)
      Else
         sRet.Append "  "
      End If
      
      ' Special handling...
      If (i Mod 16) = 15 Then
         ' Display last 16 characters in
         ' ASCII if at end of row.
         sRet.Append "  "
         For j = (i - 15) To i
            If j < nBytes Then
               If ba(j) >= 32 And ba(j) <= 126 Then
                  sRet.Append Chr$(ba(j))
               Else
                  sRet.Append "."
               End If
            End If
         Next j
      ElseIf (i Mod 8) = 7 Then
         ' Insert hyphen between 8th and
         ' 9th bytes of hex display.
         sRet.Append "-"
      Else
         ' Insert space between other bytes.
         sRet.Append " "
      End If
   Next i
   HexDump = sRet.ToString & vbCrLf  'sRet & vbCrLf & String(85, "=") & vbCrLf
End Function
