VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stringReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' VB5 -> msvbvm50.dll
Private Declare Function ArrPtr& Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any)
Private Declare Sub RtlMoveMemory Lib "kernel32" (dst As Any, src As Any, ByVal nBytes&)
Private Declare Sub RtlZeroMemory Lib "kernel32" (dst As Any, ByVal nBytes&)
Private Declare Function CharLowerBuffA& Lib "user32" (lpsz As Any, ByVal cchLength&)
Private Declare Function CharLowerBuffW& Lib "user32" (lpsz As Any, ByVal cchLength&)

Private aSrc%(), saSrc As BStrAPI.SAFEARRAY1D
Private aNew%(), saNew As BStrAPI.SAFEARRAY1D
Private aOld%(), saOld As BStrAPI.SAFEARRAY1D
Private aDst%(), saDst As BStrAPI.SAFEARRAY1D
Private aPosFnd&(), ubPosFnd&
Private aLowChars%(&H8000 To &H7FFF)

Friend Function Replace(Text As String, sOld As String, sNew As String, _
    Optional ByVal Start As Long = 1, _
    Optional ByVal Count As Long = 2147483647, _
    Optional ByVal Compare As VbCompareMethod = vbBinaryCompare _
  ) As String

    Dim c&, i&, j&, cntCpy&, Fnd0%, ubFnd&, fSameLen As Boolean
    Dim cntFnd&, ptrSrc&, ptrDst&
    Dim lenFnd&, lenSrc&, lenNew&, lenNewB&
    Dim posFnd&, posOut&, posIn&
    
    lenSrc = Len(Text)
    lenNew = Len(sNew)
    lenFnd = Len(sOld)
    ubFnd = lenFnd - 1
    ptrSrc = StrPtr(Text)
    
    If lenSrc = 0 Then Exit Function
    If lenFnd = 0 Then Replace = BStrAPI.SysAllocStringLenPtr(ptrSrc, lenSrc): Exit Function
    If Start > 0 Then i = Start - 1
    
    saSrc.pvData = ptrSrc
    saOld.pvData = StrPtr(sOld)
    saNew.pvData = StrPtr(sNew)
    
    If lenFnd = lenNew Then
        fSameLen = True
        Replace = BStrAPI.SysAllocStringLenPtr(ptrSrc, lenSrc)

        saDst.pvData = StrPtr(Replace)
'        ptrDst = StrPtr(Replace)
'        saDst.pvData = ptrDst
    End If
    
    c = lenSrc - lenFnd
    If Compare = vbBinaryCompare Then
        
        Fnd0 = aOld(0)
    
        For i = i To c
            'Inline-Cascading for first Char
            If aSrc(i) <> Fnd0 Then i = i + 1: _
                If aSrc(i) <> Fnd0 Then i = i + 1: _
                If aSrc(i) <> Fnd0 Then i = i + 1: _
                If aSrc(i) <> Fnd0 Then i = i + 1: _
                If aSrc(i) <> Fnd0 Then i = i + 1: _
                If aSrc(i) <> Fnd0 Then i = i + 1: _
                If aSrc(i) <> Fnd0 Then i = i + 1: _
                If aSrc(i) <> Fnd0 Then GoTo loopNxt

            If i > c Then Exit For

            'Search all others
            j = ubFnd
            Do While j
                If aSrc(i + j) <> aOld(j) Then GoTo loopNxt
                j = j - 1
            Loop

            cntFnd = cntFnd + 1
            'Found at Position i (0 based)
            If fSameLen Then
                j = lenNew: Do: j = j - 1: aDst(i + j) = aNew(j): Loop While j
'                saDst.pvData = ptrDst + i * 2
'                j = lenNew: Do: j = j - 1: aDst(j) = aNew(j): Loop While j
            Else
                If cntFnd > ubPosFnd Then ubPosFnd = ubPosFnd + 512: ReDim Preserve aPosFnd(ubPosFnd)
                aPosFnd(cntFnd) = i * 2
            End If

            If cntFnd = Count Then Exit For
            i = i + ubFnd
loopNxt: Next i


    Else 'vbStringCompare
    
        Fnd0 = aLowChars(aOld(0))
        
        For i = i To c
            If aLowChars(aSrc(i)) <> Fnd0 Then i = i + 1: _
                If aLowChars(aSrc(i)) <> Fnd0 Then i = i + 1: _
                If aLowChars(aSrc(i)) <> Fnd0 Then i = i + 1: _
                If aLowChars(aSrc(i)) <> Fnd0 Then i = i + 1: _
                If aLowChars(aSrc(i)) <> Fnd0 Then i = i + 1: _
                If aLowChars(aSrc(i)) <> Fnd0 Then i = i + 1: _
                If aLowChars(aSrc(i)) <> Fnd0 Then i = i + 1: _
                If aLowChars(aSrc(i)) <> Fnd0 Then GoTo loopNxt2
            
            If i > c Then Exit For
            
            'Search all others
            j = ubFnd
            Do While j
                If aLowChars(aSrc(i + j)) <> aLowChars(aOld(j)) Then GoTo loopNxt2
                j = j - 1
            Loop
            'Found at Position i (0 based)
            cntFnd = cntFnd + 1
            If fSameLen Then
                j = lenNew: Do: j = j - 1: aDst(i + j) = aNew(j): Loop While j
'                saDst.pvData = ptrDst + i * 2
'                j = lenNew: Do: j = j - 1: aDst(j) = aNew(j): Loop While j
            Else
                If cntFnd > ubPosFnd Then ubPosFnd = ubPosFnd + 512: ReDim Preserve aPosFnd(ubPosFnd)
                aPosFnd(cntFnd) = i * 2
            End If
            If cntFnd = Count Then Exit For
            i = i + ubFnd
loopNxt2: Next i
    End If
  
    'Generate Output
    If Not fSameLen Then
        If cntFnd = 0 Then
            Replace = BStrAPI.SysAllocStringLenPtr(ptrSrc, lenSrc)
        Else
            c = lenSrc + (lenNew - lenFnd) * cntFnd
            Replace = BStrAPI.SysAllocStringLenPtr(ByVal 0, c)
            ptrDst = StrPtr(Replace)
            saDst.pvData = ptrDst
            
            lenFnd = lenFnd * 2
            If lenNew Then
                lenNewB = lenNew * 2
                For i = 1 To cntFnd
                    posFnd = aPosFnd(i)
                    cntCpy = posFnd - posIn
                    
                    If cntCpy > 50 Then
                        RtlMoveMemory ByVal saDst.pvData, ByVal saSrc.pvData, cntCpy
                        saDst.pvData = saDst.pvData + cntCpy
                    ElseIf cntCpy > 0 Then
                        j = cntCpy \ 2:  Do: j = j - 1: aDst(j) = aSrc(j): Loop While j
                        saDst.pvData = saDst.pvData + cntCpy
                    End If
                    
                    posIn = posFnd + lenFnd
                    saSrc.pvData = ptrSrc + posIn
                    
                    If lenNew > 50 Then
                        RtlMoveMemory ByVal saDst.pvData, ByVal saNew.pvData, lenNewB
                    Else
                        j = lenNew: Do: j = j - 1: aDst(j) = aNew(j): Loop While j
                    End If
                    saDst.pvData = saDst.pvData + lenNewB
                Next i
            Else
                For i = 1 To cntFnd
                    posFnd = aPosFnd(i)
                    cntCpy = posFnd - posIn
                    
                    If cntCpy > 50 Then
                        RtlMoveMemory ByVal saDst.pvData, ByVal saSrc.pvData, cntCpy
                        saDst.pvData = saDst.pvData + cntCpy
                    ElseIf cntCpy > 0 Then
                        j = cntCpy \ 2:  Do: j = j - 1: aDst(j) = aSrc(j): Loop While j
                        saDst.pvData = saDst.pvData + cntCpy
                    End If
                    
                    posIn = posFnd + lenFnd
                    saSrc.pvData = ptrSrc + posIn
                Next i
            End If
            
            c = lenSrc * 2 - posIn
            If c > 50 Then
                RtlMoveMemory aDst(0), aSrc(0), c
            ElseIf c > 0 Then
                c = c \ 2:  Do: c = c - 1: aDst(c) = aSrc(c): Loop While c
            End If
        End If
    End If
End Function


Private Sub Class_Initialize()
    Dim c&
    
    ubPosFnd = 512: ReDim Preserve aPosFnd(ubPosFnd)
  
    saSrc.cDims = 1
    saSrc.cbElements = 2
    saSrc.cElements1D = &H7FFFFFFF
                   RtlMoveMemory ByVal ArrPtr(aSrc), VarPtr(saSrc), 4
    saNew = saSrc: RtlMoveMemory ByVal ArrPtr(aNew), VarPtr(saNew), 4
    saOld = saSrc: RtlMoveMemory ByVal ArrPtr(aOld), VarPtr(saOld), 4
    saDst = saSrc: RtlMoveMemory ByVal ArrPtr(aDst), VarPtr(saDst), 4
    
    For c = -32768 To 32767: aLowChars(c) = c: Next c
    If CharLowerBuffW(aLowChars(-32768), &H10000) = 0 Then
      CharLowerBuffA aLowChars(65), (223 - 65) * 2
    End If
        
    ' added by donald, 20011210
    ' patch the stooges
    ' Š 138/352   š 154/353
    ' Œ 140/338   œ 156/339
    ' Ž 142/381   ž 158/382
    ' Ÿ 159/376   ÿ 255/255
    aLowChars(352) = 353
    aLowChars(338) = 339
    aLowChars(381) = 382
    aLowChars(376) = 255

End Sub

Private Sub Class_Terminate()
    RtlZeroMemory ByVal ArrPtr(aSrc), 4
    RtlZeroMemory ByVal ArrPtr(aNew), 4
    RtlZeroMemory ByVal ArrPtr(aOld), 4
    RtlZeroMemory ByVal ArrPtr(aDst), 4
End Sub



