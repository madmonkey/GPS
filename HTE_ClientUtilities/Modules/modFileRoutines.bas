Attribute VB_Name = "modFileRoutines"
Option Explicit

Public Const cFileDescriptor = "~CU"
Private Const cModuleName = "modFileRoutines"

Private Const IS_TEXT_UNICODE_UNICODE_MASK = &HF
Private Const OFS_MAXPATHNAME = 260
Private Const OF_EXIST = &H4000

Private Type OFSTRUCT
   cBytes As Byte
   fFixedDisk As Byte
   nErrCode As Integer
   Reserved1 As Integer
   Reserved2 As Integer
   szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function IsTextUnicode Lib "advapi32" (lpBuffer As Any, ByVal cb As Long, lpi As Long) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function SHGetFileName Lib "shell32" Alias "#34" (ByVal szPath As String) As Long
Private Declare Function SHGetExtension Lib "shell32" Alias "#31" (ByVal szPath As String) As Long
Private Declare Function SHGetPath Lib "shell32" Alias "#35" (ByVal szPath As String) As Long
Private Declare Function PathStripPath Lib "shlwapi" Alias "PathStripPathA" (ByVal pPath As String) As Long
Private Declare Function PathRemoveFileSpec Lib "shlwapi" Alias "PathRemoveFileSpecA" (ByVal pPath As String) As Long
Private Declare Function getTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private sIsNT As String

Public Function GetFileName(sPathIn As String) As String
    GetFileName = GetStrFromPtr(SHGetFileName(prepareStringforAPI(sPathIn)), OFS_MAXPATHNAME)
End Function

Public Function GetExtension(sPathIn As String) As String
Dim sPathOut As String
    sPathOut = prepareStringforAPI(sPathIn)
    GetExtension = GetStrFromPtr(SHGetExtension(sPathOut), Len(sPathOut))
End Function
Public Function GetPath(sPathIn As String) As String
Dim sPathOut As String
    PathRemoveFileSpec (sPathIn)
    sPathOut = GetStrFromBuffer(sPathIn)
    GetPath = sPathOut
End Function

Public Function StripPath(sPathIn As String) As String
   PathStripPath sPathIn
   StripPath = GetStrFromBuffer(sPathIn)
End Function
Public Function FileExists(ByVal strSearchFile As String) As Boolean
    Dim strucFname As OFSTRUCT
    
    FileExists = (OpenFile(strSearchFile, strucFname, OF_EXIST) <> -1)
    
End Function
Private Function prepareStringforAPI(sStr As String) As String
Dim sReturn As String
    sReturn = sStr & String$(OFS_MAXPATHNAME - Len(sStr), 0)
    If IsWinNT Then sReturn = StrConv(sReturn, vbUnicode)
    prepareStringforAPI = sReturn
End Function
Private Function GetStrFromPtr(lpszStr As Long, nBytes As Integer) As String
  
  'Returns string before first null charencountered (if any) from a string pointer.
  'lpszStr = memory address of first byte in string
  'nBytes = number of bytes to copy.
  'StrConv used for both ANSII and Unicode strings
  'BE CAREFUL!
   ReDim ab(nBytes) As Byte   'zero-based (nBytes + 1 elements)
   CopyMemory ab(0), ByVal lpszStr, nBytes
   GetStrFromPtr = GetStrFromBuffer(StrConv(ab(), vbUnicode))
  
End Function
Private Function GetStrFromBuffer(szStr As String) As String
   
  'Returns string before first null char encountered (if any) from either an ANSII or Unicode string buffer.
   If IsUnicodeStr(szStr) Then szStr = StrConv(szStr, vbFromUnicode)
   
   If InStr(szStr, vbNullChar) Then
         GetStrFromBuffer = Left$(szStr, InStr(szStr, vbNullChar) - 1)
   Else: GetStrFromBuffer = szStr
   End If

End Function

Private Function IsUnicodeStr(sBuffer As String) As Boolean
  
  'Returns True if sBuffer evaluates to a Unicode string
   Dim dwRtnFlags As Long
   dwRtnFlags = IS_TEXT_UNICODE_UNICODE_MASK
   IsUnicodeStr = IsTextUnicode(ByVal sBuffer, Len(sBuffer), dwRtnFlags)

End Function
Private Function IsWinNT() As Boolean
Dim oHelper As Object
On Error GoTo Err_IsWinNT
    If sIsNT = vbNullString Then
        Set oHelper = CreateObject("HTE_SystemUtility.SysUtility")
        sIsNT = CStr(oHelper.IsWinNT)
    End If
    IsWinNT = CBool(sIsNT)
    Exit Function
Err_IsWinNT:
    UEH_LogError cModuleName, "IsWinNT", Err
End Function

Public Function getTemporaryFile() As String
Dim sTemp As String
Const FILE_ATTRIBUTE_TEMPORARY = &H100
    sTemp = String(260, 0)
    getTempFileName Environ("TEMP"), cFileDescriptor, 0, sTemp
    sTemp = Left$(sTemp, InStr(1, sTemp, Chr$(0)) - 1)
    SetFileAttributes sTemp, FILE_ATTRIBUTE_TEMPORARY
    getTemporaryFile = sTemp
End Function

Public Function FileCopy(ByVal sOrig As String, ByVal sDest As String, Optional ByVal bFailIfExist As Boolean = False) As Boolean
    FileCopy = (CopyFile(sOrig, sDest, bFailIfExist) <> 0)
End Function
