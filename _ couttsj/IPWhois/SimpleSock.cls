VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimpleSock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' ---------------------------------------------------------------------------------
' File...........: SimpleSock.cls
' Author.........: J.A. Coutts
' Created........: 11/01/2016
' Version........: 1.0
' Website........: http://www.yellowhead.com
' Contact........: allecnarf@hotmail.com
'
'Copyright (c) 2016 by JAC Computing
'Vernon, BC, Canada
'
'Subclassing based on CSocketMaster by Emiliano Scavuzzo
'
'SimpleSock basically performs the same functions as NewSocket. Like NewSocket,
'it supports IPv6 as well as IPv4. As long as IPv4 is used, it will operate on
'pre Windows Vista systems that do not support dual stack. This was accomplished
'by using inet_ntoa for IPv4 functions instead of inet_ntop. Unlike NewSocket,
'it cannot be used as a Control Array because of the way it handles listening
'sockets.
'
'All error messages are returned to the calling program through the Error event.
'NewSocket and it's predecessors simply reported errors and caried on attempting
'communication. This sometimes made troubleshooting difficult as mutiple errors
'would get reported. SimpleSock attempts to exit the problem function on an error
'that would cause communication to fail.
'
'================================
'EVENTS
'================================
Public Event CloseSck()
Public Event Connect()
Public Event ConnectionRequest(ByVal requestID As Long)
Public Event DataArrival(ByVal bytesTotal As Long)
Public Event EncrDataArrival(ByVal bytesTotal As Long)
Public Event Error(ByVal Number As Long, Description As String, ByVal Source As String)
Public Event SendComplete()
Public Event SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)

'================================
'General Socket Constants
'================================
Private Const AF_UNSPEC As Long = 0
Private Const AF_INET As Long = 2
Private Const AF_INET6 As Long = 23
Private Const SOCK_STREAM As Long = 1
Private Const SOCK_DGRAM As Long = 2
Private Const INADDR_ANY As Long = 0
Private Const IPPROTO_TCP As Long = 6
Private Const IPPROTO_UDP As Long = 17
Private Const SOL_SOCKET As Long = 65535
Private Const SO_SNDBUF         As Long = &H1001&
Private Const SO_RCVBUF         As Long = &H1002&
Private Const SO_MAX_MSG_SIZE   As Long = &H2003
Private Const SO_BROADCAST      As Long = &H20
Private Const SOMAXCONN         As Long = 5
Private Const LOCAL_HOST_BUFF   As Integer = 256
Private Const FIONREAD          As Long = &H4004667F
Private Const SO_REUSEADDR As Long = &H4            'Allow local address reuse
Private Const SO_LINGER As Long = &H80              'Set TIME_WAIT
Private Const SCK_TCP As Long = 0
Private Const SCK_UDP As Long = 1
Private Const MSG_PEEK As Long = &H2
Private Const TF_REUSE_SOCKET As Long = &H2

' Length of string fields for IPv4 and IPv6
Private Const INET_ADDRSTRLEN   As Long = 16
Private Const INET6_ADDRSTRLEN  As Long = 46
Private Const AI_PASSIVE        As Long = 1

'WINSOCK CONTROL ERROR CODES
Private Const sckBadState = 40006
Private Const sckInvalidArg = 40014
Private Const sckUnsupported = 40018
Private Const sckInvalidOp = 40020

'================================
'Socket states
'================================
Private Enum SockState
    sckClosed = 0
    sckopen
    scklistening
    sckConnectionPending
    sckResolvingHost
    sckHostResolved
    sckconnecting
    sckConnected
    sckClosing
    sckError
End Enum

'================================
'MEMBER VARIABLES
'================================
Private m_hSocket           As Long    'socket handle
Private m_hListen           As Long    'listening socket handle
Private m_lMemoryPointer    As Long    'memory pointer used as buffer when resolving host
Private m_lMemoryHandle     As Long    'buffer memory handle
Private m_sRemoteHost       As String  'remote host
Private m_lRemotePort       As Long    'remote port
Private m_sRemotePort       As String  'remote port
Private m_sRemoteHostIP     As String  'remote host ip
Private m_lLocalPort        As Long    'local port
Private m_lPortListen       As Long    'listening port
Private m_sLocalIP          As String  'local IP
Private m_State             As SockState    'socket state
Private m_Protocol          As Long    'defaults to 0 = TCP (UDP = 1)
Private m_IPv               As Long    'IP Version
Private m_bAcceptClass      As Boolean 'if True then this is a Accept socket class
Private m_bEncr             As Boolean 'True for secure sockets
Private m_lSendBufferLen    As Long    'winsock buffer size for sends
Private m_lRecvBufferLen    As Long    'winsock buffer size for receives

Private Type sockaddr
    sa_family           As Integer  '2 bytes
    sa_data(25)         As Byte     '26 bytes
End Type                            'Total 28 bytes

'Basic IPv4 addressing structures.
Private Type in_addr
   s_addr   As Long
End Type
Private Type sockaddr_in
    sin_family          As Integer  '2 bytes
    sin_port            As Integer  '2 bytes
    sin_addr            As in_addr  '4 bytes
    sin_zero(0 To 7)    As Byte     '8 bytes
End Type                            'Total 16 bytes

'Basic IPv6 addressing structures.
Private Type in6_addr
    s6_addr(0 To 15)      As Byte
End Type
Private Type sockaddr_in6
    sin6_family         As Integer  '2 bytes
    sin6_port           As Integer  '2 bytes
    sin6_flowinfo       As Long     '4 bytes
    sin6_addr           As in6_addr '16 bytes
    sin6_scope_id       As Long     '4 bytes
End Type                            'Total 28 bytes

Private Type addrinfo
    ai_flags As Long
    ai_family As Long
    ai_socktype As Long
    ai_protocol As Long
    ai_addrlen As Long
    ai_canonname As Long 'strptr
    ai_addr As Long 'p sockaddr
    ai_next As Long 'p addrinfo
End Type

Private m_bSendBuffer() As Byte     'local outbound byte buffer
Private m_bRecvBuffer() As Byte     'local inbound byte buffer
Private m_bData()       As Byte     'Temporary store for recovered data
Private sa_UDPdest      As sockaddr 'UDP destination socket address

'================================
'API Functions
'================================
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare Function API_Socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal stype As Long, ByVal Protocol As Long) As Long
Private Declare Function API_CloseSocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long
Private Declare Function API_Connect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef name As sockaddr, ByVal namelen As Long) As Long
Private Declare Function API_Disconnect Lib "mswsock.dll" Alias "DisconnectEx" (ByVal hSocket As Long, ByVal lpOverlapped As Long, ByVal dwFlags As Long) As Long
Private Declare Function API_GetAddrInfo Lib "ws2_32.dll" Alias "getaddrinfo" (ByVal NodeName As String, ByVal ServName As String, ByVal lpHints As Long, lpResult As Long) As Long
Private Declare Function API_FreeAddrInfo Lib "ws2_32.dll" Alias "freeaddrinfo" (ByVal res As Long) As Long
Private Declare Function API_Send Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByRef buf As Byte, ByVal datalen As Long, ByVal Flags As Long) As Long
Private Declare Function API_SendTo Lib "ws2_32.dll" Alias "sendto" (ByVal s As Long, ByRef buf As Byte, ByVal datalen As Long, ByVal Flags As Long, ByRef toaddr As sockaddr, ByVal tolen As Long) As Long
Private Declare Function API_Bind Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, ByRef name As sockaddr, ByRef namelen As Long) As Long
Private Declare Function API_Listen Lib "ws2_32.dll" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function API_Accept Lib "ws2_32.dll" Alias "accept" (ByVal s As Long, ByRef addr As sockaddr, ByRef addrlen As Long) As Long
Private Declare Function API_GetSockOpt Lib "ws2_32.dll" Alias "getsockopt" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Private Declare Function API_SetSockOpt Lib "ws2_32.dll" Alias "setsockopt" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function API_GetPeerName Lib "ws2_32.dll" Alias "getpeername" (ByVal s As Long, ByRef name As sockaddr, ByRef namelen As Long) As Long
Private Declare Function API_GetSockName Lib "ws2_32.dll" Alias "getsockname" (ByVal s As Long, ByRef name As sockaddr, ByRef namelen As Long) As Long
Private Declare Function API_GetHostName Lib "ws2_32.dll" Alias "gethostname" (ByVal host_name As String, ByVal namelen As Long) As Long
Private Declare Function API_IoctlSocket Lib "ws2_32.dll" Alias "ioctlsocket" (ByVal s As Long, ByVal cmd As Long, ByRef argp As Long) As Long
Private Declare Function API_Recv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, ByRef buf As Byte, ByVal datalen As Long, ByVal Flags As Long) As Long
Private Declare Function API_RecvFrom Lib "ws2_32.dll" Alias "recvfrom" (ByVal s As Long, ByRef buf As Byte, ByVal datalen As Long, ByVal Flags As Long, ByRef fromaddr As sockaddr, ByRef fromlen As Long) As Long

Private Declare Function inet_ntop Lib "ws2_32.dll" (ByVal af As Long, ByRef ppAddr As Any, ByRef pStringBuf As Any, ByVal StringBufSize As Long) As Long
Private Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal iaddr As Long) As Long
Private Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Long) As Integer

'================================
'Generalized functions
'================================
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Sub WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    '  ****  WARNING WARNING WARNING WARNING ******
    'This sub MUST be the first on the class. DO NOT attempt
    'to change it's location or the code will CRASH.
    'This sub receives system messages from our WndProc.
    Dim LoWord As Long
    Dim HiWord As Long
    Select Case uMsg
    Case SOCKET_MESSAGE
        LoWord = lParam And &HFFFF&
        If (lParam And &H80000000) = &H80000000 Then
            HiWord = ((lParam And &H7FFF0000) \ &H10000) Or &H8000&
        Else
            HiWord = (lParam And &HFFFF0000) \ &H10000
        End If
        PostSocket LoWord, HiWord
    End Select
End Sub


Public Function Accept(requestID As Long, Optional NoClose As Boolean) As Boolean
    Const Routine As String = "SimpleSock.Accept"
    Call PrintDebug("STATE: " & CStr(m_State))
    If m_State <> scklistening Then
        RaiseEvent Error(sckInvalidOp, "Socket is not in the listening mode", Routine)
        API_CloseSocket requestID
        Exit Function
    End If
    m_hListen = m_hSocket 'Save listening socket handle
    m_hSocket = requestID 'Use the handle created by FD_ACCEPT message
    m_Protocol = SCK_TCP
    Call ProcessOptions
    If IsSocketRegistered(requestID) Then
        RaiseEvent Error(sckBadState, "Wrong protocol or connection state for the requested transaction or request", Routine)
        Exit Function
    Else
        m_bAcceptClass = True
        m_State = sckConnected
        Call PrintDebug("STATE: sckConnected")
        Call modSocket.RegisterSocket(m_hSocket, ObjPtr(Me), False)
    End If
    Accept = True
End Function


Public Property Get bInBuffer() As Byte()
    bInBuffer = m_bData
End Property

Public Property Let bOutBuffer(bNewValue() As Byte)
    Call AddByte(m_bSendBuffer, bNewValue)
End Property

Private Function BuildArray(ByVal lSize As Long, ByVal blnPeek As Boolean, ByRef lErrorCode As Long) As Byte()
    Dim lRet        As Long
    Dim tmpSa       As sockaddr
    Dim lFlags      As Long
    Dim bTmp() As Byte
    If m_Protocol = SCK_TCP Then 'TCP transfers data from m_bRecvBuffer
        BuildArray = m_bRecvBuffer  'lSize
        If Not blnPeek Then
            Call DeleteByte(m_bRecvBuffer, lSize)
        End If
    Else 'UDP recovers data from Winsock buffer
        If blnPeek Then lFlags = MSG_PEEK
        ReDim bTmp(lSize - 1)
        lRet = API_RecvFrom(m_hSocket, bTmp(0), lSize, lFlags, tmpSa, LenB(tmpSa))
        If lRet = SOCKET_ERROR Then
            lErrorCode = Err.LastDllError
        End If
        BuildArray = bTmp
        GetRemoteInfoFromSI tmpSa, m_lRemotePort, m_sRemoteHostIP, m_sRemoteHost
    End If
End Function

Public Property Get BytesReceived() As Long
    If m_Protocol = SCK_TCP Then
        BytesReceived = GetbSize(m_bRecvBuffer)
    Else
        BytesReceived = GetBufferLenUDP
    End If
End Property

Public Function CloseSocket() As Boolean
    Const Routine As String = "SimpleSock.CloseSocket"
    Dim lErrorCode As Long
    Dim lRet As Long
    Dim bTmp() As Byte
    If Not m_hSocket = SOCKET_ERROR Then
        lRet = API_CloseSocket(m_hSocket)
        If lRet = -1 Then
            m_State = sckError
            lErrorCode = Err.LastDllError
            RaiseEvent Error(lErrorCode, "Could not Close Socket!", Routine)
        Else
            Call modSocket.UnregisterSocket(m_hSocket)
            m_hSocket = SOCKET_ERROR
            RaiseEvent CloseSck
            CloseSocket = True
            If m_bAcceptClass Then
                m_hSocket = m_hListen
                m_State = scklistening
                Call PrintDebug(CStr(m_hSocket) & ": scklistening")
            Else
                m_State = sckClosed
                Call PrintDebug("STATE: sckClosed")
            End If
            m_sLocalIP = vbNullString
            m_sRemoteHostIP = ""
            m_bRecvBuffer = bTmp
            m_bSendBuffer = bTmp
            m_lSendBufferLen = 0
            m_lRecvBufferLen = 0
        End If
    End If
End Function



Private Sub EmptyBuffer()
    'Empty winsock incoming buffer from a UDP socket.
    Dim B As Byte
    API_Recv m_hSocket, B, Len(B), 0&
End Sub

Public Property Let EncrFlg(ByVal NewValue As Boolean)
    m_bEncr = NewValue
End Property

Private Function GetBufferLenUDP() As Long
    'Returns winsock incoming buffer length from a UDP socket.
    Dim lRet    As Long
    Dim lBuff   As Long
    lRet = API_IoctlSocket(m_hSocket, FIONREAD, lBuff)
    If lRet = SOCKET_ERROR Then
        GetBufferLenUDP = 0
    Else
        GetBufferLenUDP = lBuff
    End If
End Function

Public Function GetIPFromHost(sHost As String, sPort As String, Optional Options As Long) As String
    Const Routine As String = "SimpleSock.GetIPFromHost"
    Dim sa_dest As sockaddr
    Dim Sin4 As sockaddr_in
    Dim Sin6 As sockaddr_in6
    Dim ptrResult As Long
    Dim lRet As Long
    Dim aLen As Long
    Dim Hints As addrinfo
    Dim IPFamily As Long
    If m_IPv = 4 Then
        IPFamily = AF_INET
    Else
        IPFamily = AF_INET6
    End If
    'Recover info about the destination.
    Hints.ai_family = IPFamily
    If Len(sHost) = 0 Then
        Hints.ai_flags = AI_PASSIVE  ' fill in 0 IP for me
    End If
    lRet = API_GetAddrInfo(sHost, sPort, VarPtr(Hints), ptrResult)
    If lRet <> 0 Then
        m_State = sckError
        RaiseEvent Error(lRet, GetErrorDescription(lRet), Routine)
        Exit Function
    End If
    Hints.ai_next = ptrResult   'Point to first structure in linked list
    CopyMemory Hints, ByVal Hints.ai_next, LenB(Hints) 'Copy next address info to Hints
    CopyMemory sa_dest, ByVal Hints.ai_addr, LenB(sa_dest) 'Save sockaddr portion
    API_FreeAddrInfo (ptrResult) ' free the linked list
    Hints.ai_next = ptrResult   'Point to first structure in linked list
    If sa_dest.sa_family = AF_INET6 Then
         ReDim bBuffer(0 To INET6_ADDRSTRLEN - 1)    'Resize string buffer
         CopyMemory Sin6, sa_dest, LenB(Sin6)  'Save to sockaddr_in6
         'm_lRemotePort = modSocket.IntegerToUnsigned(ntohs(Sin6.sin6_port))
         lRet = inet_ntop(AF_INET6, Sin6.sin6_addr, bBuffer(0), INET6_ADDRSTRLEN)
     Else
         ReDim bBuffer(0 To INET_ADDRSTRLEN - 1)    'Resize string buffer
         CopyMemory Sin4, sa_dest, LenB(Sin4)  'Save to sockaddr_in
         'm_lRemotePort = modSocket.IntegerToUnsigned(ntohs(Sin4.sin_port))
         lRet = inet_ntoa(Sin4.sin_addr.s_addr)
     End If
     If lRet Then GetIPFromHost = StringFromPointer(lRet)
End Function

Public Function GetLocalHostName() As String
    Const Routine As String = "SimpleSock.GetLocalHostName"
    Dim HostNameBuf As String * LOCAL_HOST_BUFF
    Dim lRet As Long
    Dim lErrorCode As Long
    lRet = API_GetHostName(HostNameBuf, LOCAL_HOST_BUFF)
    If lRet = SOCKET_ERROR Then
        GetLocalHostName = vbNullString
        lErrorCode = Err.LastDllError
        RaiseEvent Error(lErrorCode, GetErrorDescription(lErrorCode), Routine)
    Else
        GetLocalHostName = Left(HostNameBuf, InStr(1, HostNameBuf, Chr(0)) - 1)
    End If
End Function

Private Function GetLocalPort(ByVal hSocket As Long) As Long
    'Returns local port from a connected or bound socket.
    'Returns SOCKET_ERROR if fails.
    Dim newSa As sockaddr
    Dim Sin4 As sockaddr_in
    Dim Sin6 As sockaddr_in6
    Dim lRet As Long
    lRet = API_GetSockName(hSocket, newSa, LenB(newSa))
    If lRet = SOCKET_ERROR Then
        GetLocalPort = SOCKET_ERROR
    Else
        If newSa.sa_family = AF_INET6 Then
            CopyMemory Sin6, newSa, LenB(Sin6)  'Save to sockaddr_in6
            GetLocalPort = modSocket.IntegerToUnsigned(ntohs(Sin6.sin6_port))
        Else
            CopyMemory Sin4, newSa, LenB(Sin4)  'Save to sockaddr_in
            GetLocalPort = modSocket.IntegerToUnsigned(ntohs(Sin4.sin_port))
        End If
    End If
End Function

Private Function GetRemoteInfo(ByVal hSocket As Long, ByRef lRemotePort As Long, ByRef strRemoteHostIP As String, ByRef strRemoteHost As String) As Boolean
    'Retrieves remote info from a connected socket.
    'If succeeds returns TRUE and loads the arguments.
    'If fails returns FALSE and arguments are not loaded.
    Dim lRet As Long
    Dim tmpSa As sockaddr
    GetRemoteInfo = False
    lRet = API_GetPeerName(hSocket, tmpSa, LenB(tmpSa))
    If lRet = 0 Then
        GetRemoteInfo = True
        GetRemoteInfoFromSI tmpSa, m_lRemotePort, m_sRemoteHostIP, m_sRemoteHost
    Else
       lRemotePort = 0
       strRemoteHostIP = ""
       strRemoteHost = ""
    End If
End Function

Private Sub GetRemoteInfoFromSI(ByRef newSa As sockaddr, ByRef lRemotePort As Long, ByRef sRemoteHostIP As String, ByRef sRemoteHost As String)
    'Gets remote info from a sockaddr_in structure.
    Dim lRet As Long
    Dim Sin4 As sockaddr_in
    Dim Sin6 As sockaddr_in6
    Dim aLen As Long
    Dim bBuffer() As Byte
    If newSa.sa_family = AF_INET6 Then
        aLen = INET6_ADDRSTRLEN
        CopyMemory Sin6, newSa, LenB(Sin6)  'Save to sockaddr_in6
        lRemotePort = IntegerToUnsigned(ntohs(Sin6.sin6_port))
        ReDim bBuffer(0 To aLen - 1)    'Resize string buffer
        'Get IPv6 address as string
        lRet = inet_ntop(AF_INET6, Sin6.sin6_addr, bBuffer(0), aLen)
    Else 'Must be IPv4
        aLen = INET_ADDRSTRLEN
        CopyMemory Sin4, newSa, LenB(Sin4)  'Save to sockaddr_in
        lRemotePort = IntegerToUnsigned(ntohs(Sin4.sin_port))
        ReDim bBuffer(0 To aLen - 1)  'Resize string buffer
        'Get IPv4 address as string
'        lRet = inet_ntop(AF_INET, Sin4.sin_addr, bBuffer(0), aLen)
        lRet = inet_ntoa(Sin4.sin_addr.s_addr)
    End If
    If lRet Then sRemoteHostIP = StringFromPointer(lRet)
    m_sRemoteHost = ""
End Sub

Private Function Initialize() As Boolean
    'socket's handle default value
    m_hSocket = SOCKET_ERROR
    'Set IPv Flag default value
    IPvFlg = 4 'Defaults to IPv4
    Initialize = modSocket.Initialize
End Function

Public Property Get IPvFlg() As Long
    IPvFlg = m_IPv
End Property

Public Property Let IPvFlg(ByVal IPvFlg As Long)
    m_IPv = IPvFlg
End Property

Public Function Listen(LocalPort As Long) As Boolean
    Const Routine As String = "SimpleSock.Listen"
    Dim sa_local As sockaddr
    Dim lRet As Long
    Dim lErrorCode As Long
    Dim IPFamily As Long
    If m_IPv = 4 Then
        IPFamily = AF_INET
    Else
        IPFamily = AF_INET6
    End If
    'Create TCP socket
    m_Protocol = SCK_TCP
    m_State = sckconnecting
    m_hSocket = API_Socket(IPFamily, SOCK_STREAM, IPPROTO_TCP)
    If m_hSocket = SOCKET_ERROR Then
        m_State = sckError
        RaiseEvent Error(m_hSocket, GetErrorDescription(m_hSocket), Routine)
        Exit Function
    Else
        ProcessOptions 'Get Send/Recv buffer lengths & register socket
        If Not modSocket.RegisterSocket(m_hSocket, ObjPtr(Me), True) Then Exit Function
    End If
    'Populate the local sockaddr
    sa_local.sa_family = IPFamily
    sa_local.sa_data(0) = PeekB(VarPtr(LocalPort) + 1)
    sa_local.sa_data(1) = PeekB(VarPtr(LocalPort))
    'Bind to socket
    lRet = API_Bind(m_hSocket, sa_local, LenB(sa_local))
    If lRet = SOCKET_ERROR Then
        m_State = sckError
        RaiseEvent Error(lRet, GetErrorDescription(lRet), Routine)
        Exit Function
    Else 'Set into listening mode
        m_State = sckopen
        lRet = API_Listen(m_hSocket, SOMAXCONN)
        If lRet = SOCKET_ERROR Then
            lErrorCode = Err.LastDllError
            RaiseEvent Error(lErrorCode, GetErrorDescription(lErrorCode), Routine)
        Else
            m_State = scklistening
            m_lPortListen = LocalPort
            Listen = True
            Call PrintDebug("STATE: sckListening")
        End If
    End If
End Function

Public Function ListenOnExisting(LocalPort As Long) As Boolean
    Const Routine As String = "SimpleSock.Listen"
    Dim sa_local As sockaddr
    Dim lRet As Long
    Dim lErrorCode As Long
    Dim IPFamily As Long
    If m_IPv = 4 Then
        IPFamily = AF_INET
    Else
        IPFamily = AF_INET6
    End If
    m_State = sckopen
    lRet = API_Listen(m_hSocket, SOMAXCONN)
    If lRet = SOCKET_ERROR Then
        lErrorCode = Err.LastDllError
        RaiseEvent Error(lErrorCode, GetErrorDescription(lErrorCode), Routine)
    Else
        m_State = scklistening
        m_lPortListen = LocalPort
        ListenOnExisting = True
        Call PrintDebug("STATE: sckListening")
    End If
End Function

Public Property Get LocalHostName() As String
    LocalHostName = GetLocalHostName
End Property

Public Property Get LocalPort() As Long
    LocalPort = GetLocalPort(m_hSocket)
End Property

Private Sub PostSocket(ByVal lEventID As Long, ByVal lErrorCode As Long)
    'This procedure is called by the WindowProc callback function
    'from the modSocket module. The lEventID argument is an ID of the
    'network event that occurred for the socket. The lErrorCode
    'argument contains an error code only if an error was occurred
    'during an asynchronous execution.
    Const Routine As String = "SimpleSock.PostSocket"
    Dim newSa As sockaddr
    Dim Sin4 As sockaddr_in
    Dim Sin6 As sockaddr_in6
    Dim lRet As Long
    Dim lBytesRecvd As Long
    If lErrorCode <> 0 Then
        m_State = sckError
        Call PrintDebug("STATE: sckError")
        RaiseEvent Error(lErrorCode, GetErrorDescription(lErrorCode), Routine)
        CloseSocket
        Exit Sub
    End If
    Select Case lEventID
    '======================================================================
    Case FD_CONNECT
        'Arrival of this message means that the connection initiated by the call
        'of the connect Winsock API function was successfully established.
        Call PrintDebug("FD_CONNECT " & CStr(m_hSocket))
        If m_State <> sckconnecting Then
            Call PrintDebug("WARNING: Omitting FD_CONNECT")
            Exit Sub
        End If
        'Get the connection local end-point parameters
        lRet = API_GetPeerName(m_hSocket, newSa, LenB(newSa))
        If lRet = 0 Then
            If newSa.sa_family = AF_INET6 Then
                ReDim bBuffer(0 To INET6_ADDRSTRLEN - 1)    'Resize string buffer
                CopyMemory Sin6, newSa, LenB(Sin6)  'Save to sockaddr_in6
                m_lRemotePort = modSocket.IntegerToUnsigned(ntohs(Sin6.sin6_port))
                lRet = inet_ntop(AF_INET6, Sin6.sin6_addr, bBuffer(0), INET6_ADDRSTRLEN)
            Else
                ReDim bBuffer(0 To INET_ADDRSTRLEN - 1)    'Resize string buffer
                CopyMemory Sin4, newSa, LenB(Sin4)  'Save to sockaddr_in
                m_lRemotePort = modSocket.IntegerToUnsigned(ntohs(Sin4.sin_port))
'                lRet = inet_ntop(AF_INET, Sin4.sin_addr, bBuffer(0), INET_ADDRSTRLEN)
                lRet = inet_ntoa(Sin4.sin_addr.s_addr)
            End If
            If lRet Then m_sRemoteHostIP = StringFromPointer(lRet)
        End If
        m_State = sckConnected
        Call PrintDebug("STATE: sckConnected")
        RaiseEvent Connect
    '======================================================================
    Case FD_WRITE
        'This message means that the socket in a write-able
        'state, that is, buffer for outgoing data of the transport
        'service is empty and ready to receive data to send through
        'the network.
        Call PrintDebug("FD_WRITE " & CStr(m_hSocket))
        If m_State <> sckConnected Then
            Call PrintDebug("WARNING: Omitting FD_WRITE")
            Exit Sub
        End If
        If GetbSize(m_bSendBuffer) > 0 Then
            If m_Protocol = SCK_TCP Then
                TCPSend
            Else
                UDPSend
            End If
        End If
    '======================================================================
    Case FD_READ
        'Some data has arrived for this socket.
        Call PrintDebug("FD_READ " & CStr(m_hSocket))
        If m_Protocol = SCK_TCP Then
            If m_State <> sckConnected Then
                Call PrintDebug("WARNING: Omitting FD_READ")
                Exit Sub
            End If
            'Call the RecvDataToBuffer function that move arrived data
            'from the Winsock buffer to the local one and returns number
            'of bytes received.
            lBytesRecvd = RecvDataToBuffer
            If lBytesRecvd > 0 Then
                If m_bEncr Then
                    RaiseEvent EncrDataArrival(GetbSize(m_bRecvBuffer))
                Else
                    RaiseEvent DataArrival(GetbSize(m_bRecvBuffer))
                End If
            End If
        Else 'UDP protocol
            If m_State <> sckopen Then
                Call PrintDebug("WARNING: Omitting FD_READ")
                Exit Sub
            End If
            'If we use UDP we don't remove data from winsock buffer.
            'We just let the user know the amount received so
            'he/she can decide what to do.
            lBytesRecvd = GetBufferLenUDP
            If lBytesRecvd > 0 Then
                If m_bEncr Then
                    RaiseEvent EncrDataArrival(lBytesRecvd)
                Else
                    RaiseEvent DataArrival(lBytesRecvd)
                End If
            End If
            'Now the buffer is emptied no matter what the user
            'dicided to do with the received data
            EmptyBuffer
        End If
    '======================================================================
    Case FD_ACCEPT
        'When the socket is in a listening state, arrival of this message
        'means that a connection request was received.
        Call PrintDebug("FD_ACCEPT " & CStr(m_hSocket))
        If m_State <> scklistening Then
            Call PrintDebug("WARNING: Omitting FD_ACCEPT")
            Exit Sub
        End If
        'Create a new socket for the requested connection.
        lRet = API_Accept(m_hSocket, newSa, LenB(newSa))
        If lRet = SOCKET_ERROR Then
            lErrorCode = Err.LastDllError
            RaiseEvent Error(lErrorCode, GetErrorDescription(lErrorCode), Routine)
        Else
            Call GetRemoteInfo(lRet, m_lRemotePort, m_sRemoteHostIP, m_sRemoteHost)
            Call PrintDebug("OK Accepted socket: " & CStr(lRet))
            RaiseEvent ConnectionRequest(lRet)
        End If
    '======================================================================
    Case FD_CLOSE
        'This message means that the remote host is closing the conection
        Call PrintDebug("FD_CLOSE " & CStr(m_hSocket))
        If m_State <> sckConnected Then
            Call PrintDebug("WARNING: Omitting FD_CLOSE")
            Exit Sub
        End If
        m_State = sckClosing
        Call PrintDebug("REMOTE: sckClosing")
        RaiseEvent CloseSck
    End Select
End Sub

Private Sub ProcessOptions()
    'Retrieves some socket options.
    'If it is an UDP socket also sets SO_BROADCAST option.
    Const Routine As String = "SimpleSock.ProcessOptions"
    Dim lRet As Long
    Dim lBuffer As Long
    Dim lErrorCode As Long
    If m_Protocol = SCK_TCP Then
        lRet = API_GetSockOpt(m_hSocket, SOL_SOCKET, SO_RCVBUF, lBuffer, LenB(lBuffer))
        If lRet = SOCKET_ERROR Then
            lErrorCode = Err.LastDllError
            RaiseEvent Error(lErrorCode, GetErrorDescription(lErrorCode), Routine)
        Else
            m_lRecvBufferLen = lBuffer
        End If
        lRet = API_GetSockOpt(m_hSocket, SOL_SOCKET, SO_SNDBUF, lBuffer, LenB(lBuffer))
        If lRet = SOCKET_ERROR Then
            lErrorCode = Err.LastDllError
            RaiseEvent Error(lErrorCode, GetErrorDescription(lErrorCode), Routine)
        Else
            m_lSendBufferLen = lBuffer
        End If
    Else
        lBuffer = 1
        lRet = API_SetSockOpt(m_hSocket, SOL_SOCKET, SO_BROADCAST, lBuffer, LenB(lBuffer))
        lRet = API_GetSockOpt(m_hSocket, SOL_SOCKET, SO_MAX_MSG_SIZE, lBuffer, LenB(lBuffer))
        If lRet = SOCKET_ERROR Then
            lErrorCode = Err.LastDllError
            RaiseEvent Error(lErrorCode, GetErrorDescription(lErrorCode), Routine)
        Else
            m_lRecvBufferLen = lBuffer
            m_lSendBufferLen = lBuffer
        End If
    End If
    Call PrintDebug("Winsock buffer size for sends: " & CStr(m_lRecvBufferLen))
    Call PrintDebug("Winsock buffer size for receives: " & CStr(m_lSendBufferLen))
End Sub

Public Property Get Protocol() As Long
    Protocol = m_Protocol
End Property

Public Sub RecoverData(Optional maxLen As Long)
    Const Routine As String = "SimpleSock.RecoverData"
    Dim lBytesRecv  As Long
    If m_Protocol = SCK_TCP Then
        If m_State <> sckConnected And Not m_bAcceptClass Then
            RaiseEvent Error(sckBadState, "Wrong protocol or connection state for the requested transaction or request", Routine)
            Exit Sub
        End If
    Else
        If m_State <> sckopen Then
            RaiseEvent Error(sckBadState, "Wrong protocol or connection state for the requested transaction or request", Routine)
            Exit Sub
        End If
        If GetBufferLenUDP = 0 Then Exit Sub
    End If
    If maxLen < 1 Then 'MaxLen not provided
        If m_Protocol = SCK_TCP Then
            maxLen = GetbSize(m_bRecvBuffer)
        Else
            maxLen = GetBufferLenUDP
        End If
    End If
    lBytesRecv = RecvData(m_bData, False, maxLen)
    Call PrintDebug("OK Bytes obtained from buffer: " & CStr(lBytesRecv))
 End Sub

Private Function RecvData(ByRef bData() As Byte, ByVal blnPeek As Boolean, Optional maxLen As Long) As Long
    Const Routine As String = "SimpleSock.RecvData"
    'Using TCP, the data is retrieved from a local buffer (m_bRecvBuffer).
    'Using UDP, the data is retrieved from the winsock buffer.
    'If blnPeek is TRUE, the function returns number of bytes in the
    'buffer, and copies the buffer data into the data argument.
    'If blnPeek is FALSE, then this function returns number of bytes in the
    'buffer and moves the buffer data into the data argument.
    Dim lBuffLen        As Long
    Dim lErrorCode      As Long
    Dim bTmp()          As Byte
    If m_Protocol = SCK_TCP Then
        lBuffLen = GetbSize(m_bRecvBuffer)
    Else
        lBuffLen = GetBufferLenUDP
    End If
    If maxLen = 0 Then 'if maxLen argument is missing
        RecvData = lBuffLen
        If RecvData = 0 Then
            bData = bTmp
            Exit Function
        Else
            bData = BuildArray(lBuffLen, blnPeek, lErrorCode)
        End If
    Else 'if maxLen argument is not missing
        If lBuffLen = 0 Then
            RecvData = 0
            bData = bTmp
            If m_Protocol = SCK_UDP Then
                EmptyBuffer
                RaiseEvent Error(WSAEMSGSIZE, GetErrorDescription(WSAEMSGSIZE), Routine)
            End If
            Exit Function
        ElseIf maxLen > lBuffLen Then
            RecvData = lBuffLen
            bData = BuildArray(lBuffLen, blnPeek, lErrorCode)
        Else
            RecvData = maxLen
            bData = BuildArray(maxLen, blnPeek, lErrorCode)
        End If
    End If
    'if BuildArray returns an error is handled here
    If lErrorCode <> 0 Then
        RaiseEvent Error(lErrorCode, GetErrorDescription(lErrorCode), Routine)
    End If
End Function

Private Function RecvDataToBuffer() As Long
    'This function retrieves data from the Winsock buffer
    'into the class local buffer. The function returns number
    'of bytes retrieved (received).
    Const Routine As String = "SimpleSock.RecvDataToBuffer"
    Dim bInArray()     As Byte
    Dim lBytesRecv      As Long
    Dim lErrorCode      As Long
    ReDim bInArray(m_lRecvBufferLen - 1)
    lBytesRecv = API_Recv(m_hSocket, bInArray(0), m_lRecvBufferLen, 0&)
    If lBytesRecv = SOCKET_ERROR Then
        m_State = sckError
        Call PrintDebug("STATE: sckError")
        lErrorCode = Err.LastDllError
        RaiseEvent Error(lErrorCode, GetErrorDescription(lErrorCode), Routine)
    ElseIf lBytesRecv > 0 Then
        Call AddByte(m_bRecvBuffer, bInArray, lBytesRecv)
        RecvDataToBuffer = lBytesRecv
    End If
End Function

Public Property Get RemoteHostIP() As String
    RemoteHostIP = m_sRemoteHostIP
End Property

Public Sub TCPSend()
    Const Routine As String = "SimpleSock.TCPSend"
    Dim bData()         As Byte
    Dim lBufLen         As Long
    Dim lSendLen        As Long
    Dim lRet            As Long
    Dim lTotalSent      As Long
    Dim lErrorCode      As Long
    Dim lTemp           As Long
    Dim bTmp() As Byte
    If m_State <> sckConnected Then
        RaiseEvent Error(sckBadState, "Wrong connection state for TCP Send!", Routine)
        Exit Sub
    End If
    lBufLen = GetbSize(m_bSendBuffer)
    Do Until lRet = SOCKET_ERROR Or lBufLen = 0
        If GetbSize(m_bSendBuffer) = 0 Then
            Exit Sub
        ElseIf lBufLen > m_lSendBufferLen Then
            lSendLen = m_lSendBufferLen
            ReDim bData(lSendLen - 1)
            CopyMemory bData(0), m_bSendBuffer(0), lSendLen
        Else
            lSendLen = lBufLen
            bData = m_bSendBuffer
        End If
        lRet = API_Send(m_hSocket, bData(0), lSendLen, 0&)
        If lRet = SOCKET_ERROR Then
            lErrorCode = Err.LastDllError
            If lErrorCode = WSAEWOULDBLOCK Then
                Call PrintDebug("WARNING: Send buffer full, waiting...")
                If lTotalSent > 0 Then RaiseEvent SendProgress(lTotalSent, lBufLen)
            Else
                m_State = sckError
                Call PrintDebug("STATE: sckError")
                RaiseEvent Error(lErrorCode, GetErrorDescription(lErrorCode), Routine)
            End If
        Else
            Call PrintDebug("OK Bytes sent: " & CStr(lRet))
            lTotalSent = lTotalSent + lRet
            If lBufLen > lRet Then
                Call DeleteByte(m_bSendBuffer, lRet)
            Else
                Call PrintDebug("OK Finished SENDING")
                m_bSendBuffer = bTmp
                lTemp = lTotalSent
                lTotalSent = 0
                RaiseEvent SendProgress(lTemp, 0)
                RaiseEvent SendComplete
            End If
        End If
        lBufLen = GetbSize(m_bSendBuffer)
    Loop
End Sub

Public Sub UDPSend()
    Const Routine As String = "SimpleSock.UDPSend"
    Dim bData()     As Byte
    Dim lBufLen     As Long
    Dim lRet        As Long
    Dim lErrorCode  As Long
    Dim lTotalSent  As Long
    Dim lTemp       As Long
    Dim bTmp()      As Byte
    If m_State <> sckopen Then
        RaiseEvent Error(sckBadState, "Wrong socket state for UDP Send!", Routine)
        Exit Sub
    End If
    lBufLen = GetbSize(m_bSendBuffer)
    bData = m_bSendBuffer
    m_bSendBuffer = bTmp
    lRet = API_SendTo(m_hSocket, bData(0), lBufLen, 0&, sa_UDPdest, LenB(sa_UDPdest))
    If lRet = SOCKET_ERROR Then
        lErrorCode = Err.LastDllError
        m_State = sckError
        Call PrintDebug("STATE: sckError")
        RaiseEvent Error(lErrorCode, GetErrorDescription(lErrorCode), Routine)
    Else
        Call PrintDebug("OK Bytes sent: " & CStr(lRet))
        lTotalSent = lTotalSent + lRet
        If lBufLen > lRet Then
            Call DeleteByte(m_bSendBuffer, lRet)
        Else
            Call PrintDebug("OK Finished SENDING")
            m_bSendBuffer = bTmp
            lTemp = lTotalSent
            lTotalSent = 0
            RaiseEvent SendProgress(lTemp, 0)
            RaiseEvent SendComplete
        End If
    End If
End Sub


Public Property Get sInBuffer() As String
    sInBuffer = ByteToStr(m_bData)
End Property

Public Property Let sOutBuffer(sNewValue As String)
    Dim bTmp() As Byte
    bTmp = StrToByte(sNewValue)
    Call AddByte(m_bSendBuffer, bTmp)
End Property

Public Property Get State() As Long
    State = m_State
End Property

Public Function TCPConnect(RemoteHost As String, RemotePort As Long, Optional ReUse As Boolean, Optional LocalPort As Long) As Boolean
    Const Routine As String = "SimpleSock.TCPConnect"
    Dim lErrorCode As Long
    Dim sa_local As sockaddr
    Dim sa_dest As sockaddr
    Dim lRet As Long
    Dim ptrResult As Long
    Dim Hints As addrinfo
    Dim IPFamily As Long
    If m_IPv = 4 Then
        IPFamily = AF_INET
    Else
        IPFamily = AF_INET6
    End If
    'Create TCP socket
    m_Protocol = SCK_TCP
    m_State = sckconnecting
    m_hSocket = API_Socket(IPFamily, SOCK_STREAM, IPPROTO_TCP)
    If m_hSocket = SOCKET_ERROR Then
        lErrorCode = Err.LastDllError
        m_State = sckError
        RaiseEvent Error(lErrorCode, GetErrorDescription(lErrorCode), Routine)
        Exit Function
    Else
        'Get Send/Recv buffer lengths & register socket
        ProcessOptions
        If Not modSocket.RegisterSocket(m_hSocket, ObjPtr(Me), True) Then Exit Function
    End If
    'Recover info about the destination.
    Hints.ai_family = IPFamily
    lRet = API_GetAddrInfo(RemoteHost, RemotePort, VarPtr(Hints), ptrResult)
    If lRet <> 0 Then
        m_State = sckError
        RaiseEvent Error(lRet, GetErrorDescription(lRet), Routine)
        Exit Function
    End If
    Hints.ai_next = ptrResult   'Point to first structure in linked list
    CopyMemory Hints, ByVal Hints.ai_next, LenB(Hints) 'Copy next address info to Hints
    CopyMemory sa_dest, ByVal Hints.ai_addr, LenB(sa_dest) 'Save sockaddr portion
    API_FreeAddrInfo (ptrResult) ' free the linked list
    If ReUse Then
        lRet = API_SetSockOpt(m_hSocket, SOL_SOCKET, SO_LINGER, 1&, 4&)
        If lRet = SOCKET_ERROR Then
            m_State = sckError
            RaiseEvent Error(lRet, GetErrorDescription(lRet), Routine)
            Exit Function
        End If
        lRet = API_SetSockOpt(m_hSocket, SOL_SOCKET, SO_REUSEADDR, 1&, 4&)
        If lRet = SOCKET_ERROR Then
            m_State = sckError
            RaiseEvent Error(lRet, GetErrorDescription(lRet), Routine)
            Exit Function
        End If
    End If
    If LocalPort > 0 Then
        'Populate the local sockaddr
        sa_local.sa_family = IPFamily
        sa_local.sa_data(0) = PeekB(VarPtr(LocalPort) + 1)
        sa_local.sa_data(1) = PeekB(VarPtr(LocalPort))
        lRet = API_Bind(m_hSocket, sa_local, LenB(sa_local))
        If lRet = SOCKET_ERROR Then
            lErrorCode = Err.LastDllError
            m_State = sckError
            RaiseEvent Error(lErrorCode, GetErrorDescription(lErrorCode), Routine)
            Exit Function
        End If
    End If
    'Connect to sockaddr
    m_State = sckconnecting
    lRet = API_Connect(m_hSocket, sa_dest, LenB(sa_dest))
    'Check and handle errors
    If lRet = SOCKET_ERROR Then
        lErrorCode = Err.LastDllError
        If lErrorCode <> WSAEWOULDBLOCK Then
            m_State = sckError
            RaiseEvent Error(lErrorCode, GetErrorDescription(lErrorCode), Routine)
            Exit Function
        End If
    End If
    TCPConnect = True
End Function

Public Function TCPConnectAgain(RemoteHost As String, RemotePort As Long) As Boolean
    Const Routine As String = "SimpleSock.TCPConnectAgain"
    Dim lErrorCode As Long
    Dim sa_dest As sockaddr
    Dim lRet As Long
    Dim ptrResult As Long
    Dim Hints As addrinfo
    Dim IPFamily As Long
    If m_IPv = 4 Then
        IPFamily = AF_INET
    Else
        IPFamily = AF_INET6
    End If
    'Recover info about the destination.
    Hints.ai_family = IPFamily
    lRet = API_GetAddrInfo(RemoteHost, RemotePort, VarPtr(Hints), ptrResult)
    If lRet <> 0 Then
        m_State = sckError
        RaiseEvent Error(lRet, GetErrorDescription(lRet), Routine)
        Exit Function
    End If
    Hints.ai_next = ptrResult   'Point to first structure in linked list
    CopyMemory Hints, ByVal Hints.ai_next, LenB(Hints) 'Copy next address info to Hints
    CopyMemory sa_dest, ByVal Hints.ai_addr, LenB(sa_dest) 'Save sockaddr portion
    API_FreeAddrInfo (ptrResult) ' free the linked list
    lRet = API_Connect(m_hSocket, sa_dest, LenB(sa_dest))
    'Check and handle errors
    If lRet = SOCKET_ERROR Then
        lErrorCode = Err.LastDllError
        If lErrorCode <> WSAEWOULDBLOCK Then
            m_State = sckError
            RaiseEvent Error(lErrorCode, GetErrorDescription(lErrorCode), Routine)
            Exit Function
        End If
    End If
    TCPConnectAgain = True
End Function

Public Function UDPInit(RemoteHost As String, RemotePort As Long, LocalPort As Long) As Boolean
    Const Routine As String = "SimpleSock.UDPInit"
    'Using UDP, we just bind the socket to a known port
    Dim lErrorCode As Long
    Dim sa_local As sockaddr
    Dim lRet As Long
    Dim ptrResult As Long
    Dim Hints As addrinfo
    Dim IPFamily As Long
    Dim SocketExists As Boolean
    If m_hSocket > 1 Then SocketExists = True
    If m_IPv = 4 Then
        IPFamily = AF_INET
    Else
        IPFamily = AF_INET6
    End If
    m_Protocol = SCK_UDP
    'If UDP socket does not exist, create and register it
    Debug.Print "UDP Socket " & CStr(m_hSocket)
    If Not SocketExists Then
        m_hSocket = API_Socket(IPFamily, SOCK_DGRAM, IPPROTO_UDP)
        If m_hSocket = SOCKET_ERROR Then
            m_State = sckError
            RaiseEvent Error(m_hSocket, GetErrorDescription(m_hSocket), Routine)
            Exit Function
        Else
            'Get Send/Recv buffer lengths & register socket
            ProcessOptions
            If Not modSocket.RegisterSocket(m_hSocket, ObjPtr(Me), True) Then Exit Function
        End If
    End If
    'Recover info about the destination.
    Hints.ai_family = IPFamily
    lRet = API_GetAddrInfo(RemoteHost, RemotePort, VarPtr(Hints), ptrResult)
    If lRet <> 0 Then
        m_State = sckError
        RaiseEvent Error(lRet, GetErrorDescription(lRet), Routine)
        Exit Function
    End If
    Hints.ai_next = ptrResult   'Point to first structure in linked list
    CopyMemory Hints, ByVal Hints.ai_next, LenB(Hints) 'Copy next address info to Hints
    CopyMemory sa_UDPdest, ByVal Hints.ai_addr, LenB(sa_UDPdest)    'Save sockaddr portion
    API_FreeAddrInfo (ptrResult) ' free the linked list
    If SocketExists Then
        UDPInit = True
        Exit Function
    End If
    'Populate the local sockaddr
    sa_local.sa_family = IPFamily
    sa_local.sa_data(0) = PeekB(VarPtr(LocalPort) + 1)
    sa_local.sa_data(1) = PeekB(VarPtr(LocalPort))
    'Bind to socket
    lRet = API_Bind(m_hSocket, sa_local, LenB(sa_local))
    If lRet = SOCKET_ERROR Then
        m_State = sckError
        RaiseEvent Error(lRet, GetErrorDescription(lRet), Routine)
    Else
        m_State = sckopen
        UDPInit = True
    End If
End Function

Private Sub Class_Initialize()
    Initialize
End Sub


Private Sub Class_Terminate()
    'clean processes and finish winsock service
    Call modSocket.Finalize
End Sub


