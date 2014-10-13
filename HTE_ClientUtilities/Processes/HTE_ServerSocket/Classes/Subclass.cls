VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Subclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC       As Long = (-4)      'Get/Set the WndProc address with GetWindowLong/SetWindowLong
Private Const BUF_MAX           As Long = 512       'Maximum opcode bytes. Requires 125 + (# Messages * 11)
Private Const OFF_BEFORE        As Long = 3         'Offset into the opcode bytes for the BEFORE default processing
Private Const OFF_AFTER         As Long = 65        'Offset into the opcode bytes for the AFTER default processing
Private Const CODE_CMP_EAX      As String = "3D"    'Opcode for cmp eax,????????
Private Const CODE_JE           As String = "0F84"  'Opcode for je with a 4 byte relative offset.
Private Const CODE_WNDPROC      As String = "558BEC83C4F8EB6AC745FC000000008D45FC50C745F8000000008D45F850B8ptrOwner8BC88B0950FF5120837DF800753AFF7514FF7510FF750CFF7508E8wnd_procC9C21000E8wnd_proc8945FCFF7514FF7510FF750CFF75088D45FC50B8ptrOwner8BC88B0950FF511C8B45FCC9C210008B450CFF7514FF751050FF7508"
Private Const CODE_WNDPROC_ALL  As String = "558BEC83C4FCFF7514FF7510FF750CFF7508E8wnd_proc8945FCFF7514FF7510FF750CFF75088D45FC50B8ptrOwner8BC88B0950FF511C8B45FCC9C21000"

Private Type tCode
  Buf(0 To BUF_MAX) As Byte       'Opcode buffer
End Type

Private Type tCodeBuf
  Code              As tCode      'WndProc opcodes
  nBuf              As Long       'Opcode buffer index
End Type
Private Running     As Boolean    'Is the subclasser running?
Private All         As Boolean    'All messages?
Private hWnd        As Long       'Window being subclassed
Private WndProcPrev As Long       'The address of the existing WndProc
Private CodeBuf     As tCodeBuf   'Opcode buffer
Public Sub AttachMessage(MsgNum As Long, Optional BeforeEvent As Boolean = False)

'Add a message to those that will call back either before or after the existing WndProc.
    Debug.Assert (Running = False)
    With CodeBuf
        If .nBuf = 0 Then                                   'If the buffer is empty (first message to be added)
            Call AddCode(CODE_WNDPROC)                        'Add the WndProc opcodes
        End If
        Call AddCode(CODE_CMP_EAX & Hex8(htonl(MsgNum)))    'Append the opcodes to compare the MsgNum
        'Append the opcodes to jump if matched
        Call AddCode(CODE_JE & Hex8(htonl(Not (.nBuf - IIf(BeforeEvent, OFF_BEFORE, OFF_AFTER)))))
    End With
End Sub

Public Sub Subclass(hWndSub As Long, Owner As ISubclass, Optional AllMsgs As Boolean = False)
'Subclass the passed window handle.
  Dim ptrCode   As Long                                 'Address of the machine code
  Dim ptrOwner  As Long                                 'Object address of the owner
  Dim nPos      As Long

  Debug.Assert (Running = False)                        'Subclasser already running
  Debug.Assert (IsWindow(hWndSub))                      'Invalid hWnd
  Debug.Assert (Not AllMsgs And CodeBuf.nBuf > 0) Or _
               (AllMsgs And CodeBuf.nBuf = 0)           'Either filtered mode but no messages added OR All message mode but messages added.
  
  All = AllMsgs
  
  With CodeBuf
    hWnd = hWndSub                                      'Save the window handle
    WndProcPrev = GetWindowLong(hWnd, GWL_WNDPROC)      'Save the address of the current WndProc
    ptrOwner = ObjPtr(Owner)                            'Get the address of the owner
    ptrCode = VarPtr(.Code.Buf(0))                      'Get the address of our WndProc code
    
    If All Then
    
      Call AddCode(CODE_WNDPROC_ALL)
      Call CopyMemory(ByVal ptrCode + 19, (WndProcPrev - (ptrCode + 23)), 4)
      Call CopyMemory(ByVal ptrCode + 43, ptrOwner, 4)
    Else
    
      nPos = .nBuf + 1                                  'Save the buf pos for a later patch
      Call AddCode("E8wnd_procC9C21000")                'Add the trailing WndProcPrev call
      
      'Patch the opcodes with run-time values
      Call CopyMemory(ByVal ptrCode + 31, ptrOwner, 4)
      Call CopyMemory(ByVal ptrCode + 62, (WndProcPrev - (ptrCode + 66)), 4)
      Call CopyMemory(ByVal ptrCode + 71, (WndProcPrev - (ptrCode + 75)), 4)
      Call CopyMemory(ByVal ptrCode + 95, ptrOwner, 4)
      Call CopyMemory(ByVal ptrCode + nPos, (WndProcPrev - (ptrCode + nPos + 4)), 4)
    End If
  End With
  
  Call SetWindowLong(hWnd, GWL_WNDPROC, ptrCode)        'Set our WndProc
  Running = True
End Sub

Public Sub UnSubclass()
'Unsubclass the window
  Dim ptrCode As Long
  
    If Running Then
        ptrCode = VarPtr(CodeBuf.Code.Buf(0))
        Call CopyMemory(ByVal ptrCode + 6, &H10C2C9, 4)     'Block up the entrance
        If All Then
            Call CopyMemory(ByVal ptrCode + 23, &H10C2C9, 4)
        Else
            Call CopyMemory(ByVal ptrCode + 75, &H10C2C9, 4)  'Bail between iSuperClass_After and default processing
        End If
        Call SetWindowLong(hWnd, GWL_WNDPROC, WndProcPrev)  'Restore the previous WndProc
        CodeBuf.nBuf = 0                                    'Reset the opcode buffer
        Running = False                                     'Not running
    End If

End Sub
Private Sub Class_Terminate()
  If Running Then UnSubclass                            'Unsubclass if the Subclasser is running
End Sub
Private Sub AddCode(sOps As String)
'Translate the passed hex string character pairs to bytes and stuff into the opcode buffer.
Dim i As Long
Dim j As Long

    With CodeBuf
        j = Len(sOps)                                       'Get length of opcode string
        Debug.Assert (j Mod 2 = 0)                          'sOps length not evenly divisble by 2
        Debug.Assert (.nBuf + (j \ 2) <= BUF_MAX)           'Opcode buffer overflow, increase value of BUF_MAX
        For i = 1 To j Step 2                               'For each pair of hex chars
            .Code.Buf(.nBuf) = Val("&H" & Mid$(sOps, i, 2))   'Convert from hex to byte, append to buffer
            .nBuf = .nBuf + 1                                 'Bump the opcode buffer index
        Next i
    End With
    
End Sub
Private Function Hex8(lValue As Long) As String
'Return an 8 character hex representation of the passed 32 bit value
  Dim s As String
    s = Hex$(lValue)
    Hex8 = String$(8 - Len(s), "0") & s
End Function


