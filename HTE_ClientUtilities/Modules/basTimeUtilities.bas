Attribute VB_Name = "basTimeUtilities"
Option Explicit
'Added for more accurate time reporting
Private Const TIME_ZONE_ID_UNKNOWN  As Long = 1
Private Const TIME_ZONE_ID_STANDARD As Long = 1
Private Const TIME_ZONE_ID_DAYLIGHT As Long = 2
Private Const TIME_ZONE_ID_INVALID  As Long = &HFFFFFFFF
 
Public Type SystemTime
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
    StandardDate As SystemTime
    StandardBias As Long
    DaylightName(63) As Byte
    DaylightDate As SystemTime
    DaylightBias As Long
End Type
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SystemTime)

Public Function GetClientTimeZoneInformation() As String
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
    If InStr(1, sReturn, "-", vbBinaryCompare) = 0 Then
        sReturn = "+" & sReturn
    End If
    GetClientTimeZoneInformation = sReturn
    
End Function
Public Function GetAccurateTime(Optional ByVal UseUTC As Boolean = False) As Variant
'Returns Variant since VB6 does NOT intrinsically support the decimal datatype
Dim myTime As SystemTime
Dim chk As Integer
Dim rtn As Variant
    GetSystemTime myTime
    If UseUTC Then
        chk = myTime.wHour
    Else
        chk = myTime.wHour + (Val(GetClientTimeZoneInformation) / 100)
    End If
    Select Case True
        Case chk > 0
            rtn = (1 * chk) / 24
        Case chk < 0
            rtn = ((1 * chk) / 24) + 1
        Case Else
            rtn = 0
    End Select
    If myTime.wMinute > 0 Then
        rtn = rtn + ((1 * myTime.wMinute) / 1440)
    End If
    If myTime.wSecond > 0 Then
        rtn = rtn + ((1 * myTime.wSecond) / 86400)
    End If
    If myTime.wMilliseconds > 0 Then
        rtn = rtn + ((1 * myTime.wMilliseconds) / 86400000)
    End If
    GetAccurateTime = CDec(CDbl(Date) + rtn)
    
End Function

Public Function NowEx() As Variant
    NowEx = GetAccurateTime
End Function

Public Function GetUTCEx() As Variant
    GetUTCEx = GetAccurateTime(True)
End Function
Private Function GetMilliseconds(ByVal varDateTime As Variant) As Long
Dim decConversionFactor As Variant
Dim decTime As Variant
 
'  The Decimal datatype can store decimal values exactly.
'  Variables cannot be directly declared as Decimal, so create a Variant then use CDec( ) to convert to Decimal.
     'K is used to convert a VB time unit back to seconds
     'K = 86400000 milliseconds per day
        decConversionFactor = CDec(86400000)
        decTime = CDec(varDateTime) 'Store the DateTime value in an exact decimal value called decTime
        decTime = Abs(decTime)     'Make sure the date/time value is positive
        decTime = decTime - Int(decTime) 'Get rid of the date (whole number), leaving time (decimal)
        decTime = (decTime * decConversionFactor)     'Convert to time to seconds
        GetMilliseconds = decTime Mod 1000     'Return the milliseconds
 
End Function
Public Function FormatSystemTime(SystemTime As Variant) As String
  FormatSystemTime = Format$(SystemTime, "YYYY-MM-DD HH:NN:SS") & "." & Format$(GetMilliseconds(SystemTime), "00#")
End Function

