Attribute VB_Name = "modMediaAccessControl"
Option Explicit

Private Const NCBASTAT As Long = &H33
Private Const NCBNAMSZ As Long = 16
Private Const HEAP_ZERO_MEMORY As Long = &H8
Private Const HEAP_GENERATE_EXCEPTIONS As Long = &H4
Private Const NCBRESET As Long = &H32

Private Type NET_CONTROL_BLOCK  'NCB
   ncb_command    As Byte
   ncb_retcode    As Byte
   ncb_lsn        As Byte
   ncb_num        As Byte
   ncb_buffer     As Long
   ncb_length     As Integer
   ncb_callname   As String * NCBNAMSZ
   ncb_name       As String * NCBNAMSZ
   ncb_rto        As Byte
   ncb_sto        As Byte
   ncb_post       As Long
   ncb_lana_num   As Byte
   ncb_cmd_cplt   As Byte
   ncb_reserve(9) As Byte 'Reserved, must be 0
   ncb_event      As Long
End Type

Private Type ADAPTER_STATUS
   adapter_address(5) As Byte
   rev_major         As Byte
   reserved0         As Byte
   adapter_type      As Byte
   rev_minor         As Byte
   duration          As Integer
   frmr_recv         As Integer
   frmr_xmit         As Integer
   iframe_recv_err   As Integer
   xmit_aborts       As Integer
   xmit_success      As Long
   recv_success      As Long
   iframe_xmit_err   As Integer
   recv_buff_unavail As Integer
   t1_timeouts       As Integer
   ti_timeouts       As Integer
   Reserved1         As Long
   free_ncbs         As Integer
   max_cfg_ncbs      As Integer
   max_ncbs          As Integer
   xmit_buf_unavail  As Integer
   max_dgram_size    As Integer
   pending_sess      As Integer
   max_cfg_sess      As Integer
   max_sess          As Integer
   max_sess_pkt_size As Integer
   name_count        As Integer
End Type
   
Private Type NAME_BUFFER
   name        As String * NCBNAMSZ
   name_num    As Integer
   name_flags  As Integer
End Type

Private Type ASTAT
   adapt          As ADAPTER_STATUS
   NameBuff(30)   As NAME_BUFFER
End Type

Private Type HostEnt
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLen As Integer
    hAddrList As Long
End Type

Private Const WSA_DESCRIPTIONLEN = 256
Private Const WSA_DESCRIPTIONSIZE = WSA_DESCRIPTIONLEN + 1

Private Const WSA_SYS_STATUS_LEN = 128
Private Const WSA_SYSSTATUSSIZE = WSA_SYS_STATUS_LEN + 1

Private Type WSADataType
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DESCRIPTIONSIZE
    szSystemStatus As String * WSA_SYSSTATUSSIZE
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Const NO_ERROR = 0
Private Const SOCKET_ERROR = -1

Private Declare Function Netbios Lib "netapi32" (pncb As NET_CONTROL_BLOCK) As Byte
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryAny Lib "kernel32" Alias "RtlMoveMemory" (dst As Any, src As Any, ByVal bcount As Long)
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare Function SendARP Lib "iphlpapi.dll" (ByVal DestIP As Long, ByVal SrcIP As Long, pMacAddr As Long, PhyAddrLen As Long) As Long
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal s As String) As Long
Private Declare Function gethostname Lib "wsock32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long
Private Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal host_name As String) As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
Private Declare Function WSAIsBlocking Lib "wsock32.dll" () As Long
Private Declare Function WSACancelBlockingCall Lib "wsock32.dll" () As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private WSAStartedUp As Boolean     'Flag to keep track of whether winsock WSAStartup wascalled

Public Function GetPhysicalID(ByVal sIP As String) As String
Dim sReturn As String
    sReturn = GetIPAddress
    If InStr(1, sReturn, sIP, vbBinaryCompare) > 0 Then
        'its a local address
        GetPhysicalID = GetMACAddress
    Else
        If GetRemoteMACAddress(sIP, sReturn) Then
            GetPhysicalID = sReturn
        Else
            GetPhysicalID = vbNullString '"{Not Available}"
        End If
    End If
End Function
Private Function GetMACAddress(Optional sDelimiter As String = "-") As String

  'retrieve the MAC Address for the network controller installed, returning a formatted string
   
    Dim tmp As String
    Dim pASTAT As Long
    Dim NCB As NET_CONTROL_BLOCK
    Dim AST As ASTAT
    Dim cnt As Long

    'The IBM NetBIOS 3.0 specifications defines four basic
    'NetBIOS environments under the NCBRESET command. Win32
    'follows the OS/2 Dynamic Link Routine (DLR) environment.
    'This means that the first NCB issued by an application
    'must be a NCBRESET, with the exception of NCBENUM.
    'The Windows NT implementation differs from the IBM
    'NetBIOS 3.0 specifications in the NCB_CALLNAME field.
     NCB.ncb_command = NCBRESET
     Call Netbios(NCB)
   
    'To get the Media Access Control (MAC) address for an
    'ethernet adapter programmatically, use the Netbios()
    'NCBASTAT command and provide a "*" as the name in the
    'NCB.ncb_CallName field (in a 16-chr string).
    NCB.ncb_callname = "*               "
    NCB.ncb_command = NCBASTAT
   
    'For machines with multiple network adapters you need to
    'enumerate the LANA numbers and perform the NCBASTAT
    'command on each. Even when you have a single network
    'adapter, it is a good idea to enumerate valid LANA numbers
    'first and perform the NCBASTAT on one of the valid LANA
    'numbers. It is considered bad programming to hardcode the
    'LANA number to 0 (see the comments section below).
    NCB.ncb_lana_num = 0
    NCB.ncb_length = Len(AST)
   
    pASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS Or HEAP_ZERO_MEMORY, NCB.ncb_length)
    If pASTAT <> 0 Then
        NCB.ncb_buffer = pASTAT
        Call Netbios(NCB)
        CopyMemory AST, NCB.ncb_buffer, Len(AST)
        'convert the byte array to a string
        GetMACAddress = MakeMacAddress(AST.adapt.adapter_address(), sDelimiter)
        HeapFree GetProcessHeap(), 0, pASTAT
    Else
        Debug.Print "memory allocation failed!"
        Exit Function
    End If
   
End Function

Private Function MakeMacAddress(b() As Byte, Optional sDelim As String = "-") As String
    Dim cnt As Long
    Dim buff As String

On Local Error GoTo MakeMac_error
  
    If UBound(b) = 5 Then 'so far, MAC addresses are exactly 6 segments in size (0-5)
        For cnt = 0 To 4 'concatenate the first five values together and separate with the delimiter char
            buff = buff & Right$("00" & Hex(b(cnt)), 2) & sDelim
        Next
        buff = buff & Right$("00" & Hex(b(5)), 2) 'and append the last value
    End If  'UBound(b)
   MakeMacAddress = buff
   
MakeMac_exit:
   Exit Function
   
MakeMac_error:
   MakeMacAddress = "(error building MAC address)"
   Resume MakeMac_exit
   
End Function

Private Function GetRemoteMACAddress(ByVal sRemoteIP As String, sRemoteMacAddress As String, Optional sDelimiter As String = "-") As Boolean
    Dim dwRemoteIP As Long
    Dim pMacAddr As Long
    Dim bpMacAddr() As Byte
    Dim PhyAddrLen As Long
    
    'convert the string IP into an unsigned long value containing
    'a suitable binary representation of the Internet address given
    dwRemoteIP = ConvertIPtoLong(sRemoteIP)
    If dwRemoteIP <> 0 Then
        'must set this up first!
        PhyAddrLen = 6
        'assume failure
        GetRemoteMACAddress = False
        'retrieve the remote MAC address
        If SendARP(dwRemoteIP, 0&, pMacAddr, PhyAddrLen) = NO_ERROR Then
            If (pMacAddr <> 0) And (PhyAddrLen > 0) Then
                'returned value is a long pointer to the MAC address, so copy data to a byte array
                ReDim bpMacAddr(0 To PhyAddrLen - 1)
                CopyMemoryAny bpMacAddr(0), pMacAddr, ByVal PhyAddrLen
                'convert the byte array to a string and return success
                sRemoteMacAddress = MakeMacAddress(bpMacAddr(), sDelimiter)
                GetRemoteMACAddress = True
            End If 'pMacAddr
        Else
            sRemoteMacAddress = "Remote call failed"
        End If  'SendARP
    End If  'dwRemoteIP
      
End Function

Private Function ConvertIPtoLong(sIpAddress) As Long
    ConvertIPtoLong = inet_addr(sIpAddress)
End Function

Public Function GetIPAddress() As String
'Function to retrieve the IP address(es) returns pipe-delimited string
    Dim sHostName As String * 256
    Dim lpHost As Long
    Dim Host As HostEnt
    Dim dwIPAddr As Long
    Dim tmpIPAddr() As Byte
    Dim i As Integer
    Dim sIPAddr As String
    Static IPADDRESSLIST As String
    Const cSTARTUP As String = "{588D6A75-40BF-4421-8FB1-23D4A9AD861C}"

    If IPADDRESSLIST = vbNullString Then
        
        If Not WSAStartedUp Then StartWinsock (cSTARTUP)
        If gethostname(sHostName, 256) = SOCKET_ERROR Then
            GetIPAddress = vbNullString
            Debug.Print "Windows Sockets error " & Str$(WSAGetLastError()) & " has occurred. Unable to successfully get Host Name."
            Exit Function
        End If
        
        sHostName = Trim$(sHostName)
        lpHost = gethostbyname(sHostName)
        
        If lpHost = 0 Then
            GetIPAddress = vbNullString
            Debug.Print "Windows Sockets are not responding. " & "Unable to successfully get Host Name."
            Exit Function
        End If
        
        CopyMemory Host, lpHost, Len(Host)
        CopyMemory dwIPAddr, Host.hAddrList, 4
        
        Do Until dwIPAddr = 0
            ReDim tmpIPAddr(1 To Host.hLen)
            CopyMemory tmpIPAddr(1), dwIPAddr, Host.hLen
            For i = 1 To Host.hLen
                sIPAddr = sIPAddr & tmpIPAddr(i) & "."
            Next
            Host.hAddrList = Host.hAddrList + LenB(Host.hAddrList)
            CopyMemory dwIPAddr, Host.hAddrList, 4
            sIPAddr = Mid$(sIPAddr, 1, Len(sIPAddr) - 1) & "|"
        Loop
        
        IPADDRESSLIST = Mid$(sIPAddr, 1, Len(sIPAddr) - 1) 'remove last pipe
        If WSAStartedUp Then EndWinsock
            
    End If
    
    GetIPAddress = IPADDRESSLIST
    
End Function

Private Function StartWinsock(sDescription As String) As Boolean
    Dim StartupData As WSADataType
    If Not WSAStartedUp Then
        If Not WSAStartup(&H101, StartupData) Then
            WSAStartedUp = True
            sDescription = StartupData.szDescription
        Else
            WSAStartedUp = False
        End If
    End If
    StartWinsock = WSAStartedUp
End Function

Private Sub EndWinsock()
    Dim Ret&
    If WSAIsBlocking() Then
        Ret = WSACancelBlockingCall()
    End If
    Ret = WSACleanup()
    WSAStartedUp = False
End Sub
