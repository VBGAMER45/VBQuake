Attribute VB_Name = "net_wins"
' net_wins.bas - the Winsock TCP/IP LAN driver (net_wins.c)
'
' C loads wsock32.dll with LoadLibrary and fills the function vectors
' (psocket, precvfrom, ...) with GetProcAddress, so that it runs on Windows
' 3.1 without Winsock; here they are Declare statements under the same
' names. A qsockaddr holds a sockaddr_in: sa_data(0 .. 1) is the port and
' sa_data(2 .. 5) the address, both in network byte order; myAddr and the
' other addresses are kept as C keeps s_addr (network order). Winsock's
' error codes come from Err.LastDllError (the VB runtime overwrites the
' thread's last error before WSAGetLastError could read it).
'
' Deviations: WINS_Read also takes WSAECONNRESET as "no data" (Windows 2000
' and later report a UDP port with nobody listening that way; C knows only
' WSAECONNREFUSED). The local name is looked up without the blocking hook
' (C cancels gethostbyname after 2 seconds). Test aid: under -hidden without
' -ip, every socket is bound to 127.0.0.1, so a test run opens no port on
' the network.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Declare Function pWSAStartup Lib "wsock32.dll" Alias "WSAStartup" (ByVal wVersionRequired As Integer, lpWSAData As Any) As Long
Private Declare Function pWSACleanup Lib "wsock32.dll" Alias "WSACleanup" () As Long
Private Declare Function psocket Lib "wsock32.dll" Alias "socket" (ByVal af As Long, ByVal stype As Long, ByVal protocol As Long) As Long
Private Declare Function pioctlsocket Lib "wsock32.dll" Alias "ioctlsocket" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long
Private Declare Function psetsockopt Lib "wsock32.dll" Alias "setsockopt" (ByVal s As Long, ByVal level As Long, _
    ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function precvfrom Lib "wsock32.dll" Alias "recvfrom" (ByVal s As Long, buf As Any, ByVal length_ As Long, _
    ByVal flags As Long, from As Any, fromlen As Any) As Long
Private Declare Function psendto Lib "wsock32.dll" Alias "sendto" (ByVal s As Long, buf As Any, ByVal length_ As Long, _
    ByVal flags As Long, to_ As Any, ByVal tolen As Long) As Long
Private Declare Function pclosesocket Lib "wsock32.dll" Alias "closesocket" (ByVal s As Long) As Long
Private Declare Function pgethostname Lib "wsock32.dll" Alias "gethostname" (name As Any, ByVal namelen As Long) As Long
Private Declare Function pgethostbyname Lib "wsock32.dll" Alias "gethostbyname" (ByVal name As String) As Long
Private Declare Function pgethostbyaddr Lib "wsock32.dll" Alias "gethostbyaddr" (addr As Any, ByVal length_ As Long, ByVal type_ As Long) As Long
Private Declare Function pgetsockname Lib "wsock32.dll" Alias "getsockname" (ByVal s As Long, name As Any, namelen As Long) As Long
Private Declare Function bind Lib "wsock32.dll" (ByVal s As Long, name As Any, ByVal namelen As Long) As Long
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long

Private Const AF_INET As Long = 2
Private Const PF_INET As Long = 2
Private Const SOCK_DGRAM As Long = 2
Private Const IPPROTO_UDP As Long = 17
Private Const FIONBIO As Long = &H8004667E
Private Const SOL_SOCKET As Long = &HFFFF&
Private Const SO_BROADCAST As Long = &H20
Private Const MSG_PEEK As Long = 2
Private Const SOCKET_ERROR As Long = -1
Private Const INADDR_ANY As Long = 0
Private Const INADDR_BROADCAST As Long = -1
Private Const INADDR_NONE As Long = -1
Private Const WSAEWOULDBLOCK As Long = 10035
Private Const WSAECONNRESET As Long = 10054
Private Const WSAECONNREFUSED As Long = 10061

Private Const MAXHOSTNAMELEN As Long = 256
Private Const QSOCKADDR_SIZE As Long = 16       ' sizeof(struct qsockaddr)

Private net_acceptsocket As Long         ' socket for fielding new connections (C: = -1)
Private net_controlsocket As Long
Private net_broadcastsocket As Long
Private broadcastaddr As qsockaddr

Private myAddr As Long

Private winsock_lib_initialized As Boolean

Private winsock_initialized As Long
Private winsockdata(0 To 399) As Byte    ' WSADATA

'=============================================================================
' (QuakeVB) sockaddr_in fields of a qsockaddr

Private Function SinAddr(addr As qsockaddr) As Long
    Dim a As Long
    CopyMemory a, addr.sa_data(2), 4
    SinAddr = a
End Function

Private Sub SetSinAddr(addr As qsockaddr, ByVal a As Long)
    CopyMemory addr.sa_data(2), a, 4
End Sub

' C: ntohs (sin_port)
Private Function SinPort(addr As qsockaddr) As Long
    SinPort = addr.sa_data(0) * 256& + addr.sa_data(1)
End Function

' C: sin_port = htons ((unsigned short)port)
Private Sub SetSinPort(addr As qsockaddr, ByVal port As Long)
    addr.sa_data(0) = (port And &HFF00&) \ &H100&
    addr.sa_data(1) = port And &HFF&
End Sub

' a network order address as "a.b.c.d" (C: inet_ntoa)
Public Function WINS_DottedAddr(ByVal a As Long) As String
    Dim b(0 To 3) As Byte
    CopyMemory b(0), a, 4
    WINS_DottedAddr = b(0) & "." & b(1) & "." & b(2) & "." & b(3)
End Function

' C: x << 8 on an int
Private Function Shl8(ByVal v As Long) As Long
    Shl8 = (v And &H7FFFFF) * &H100&
    If v And &H800000 Then Shl8 = Shl8 Or &H80000000
End Function

' the zero terminated string at a C pointer
Private Function CStringAt(ByVal p As Long) As String
    Dim n As Long, b() As Byte
    n = lstrlenA(p)
    If n = 0 Then Exit Function
    ReDim b(0 To n - 1)
    CopyMemory b(0), ByVal p, n
    CStringAt = BytesToStr(b, 0, n)
End Function

' the zero terminated string in a byte buffer
Private Function BufString(b() As Byte) As String
    Dim n As Long
    Do While n <= UBound(b)
        If b(n) = 0 Then Exit Do
        n = n + 1
    Loop
    BufString = BytesToStr(b, 0, n)
End Function

' C: *(int *)hostentry->h_addr_list[0]
Private Function HostentAddr(ByVal hostentry As Long) As Long
    Dim list As Long, p As Long, a As Long
    CopyMemory list, ByVal hostentry + 12, 4
    CopyMemory p, ByVal list, 4
    If p Then CopyMemory a, ByVal p, 4
    HostentAddr = a
End Function

'=============================================================================

Private Sub WINS_GetLocalAddress()
    Dim local_ As Long
    Dim buff(0 To MAXHOSTNAMELEN - 1) As Byte

    If myAddr <> INADDR_ANY Then Exit Sub

    If pgethostname(buff(0), MAXHOSTNAMELEN) = SOCKET_ERROR Then Exit Sub

    local_ = pgethostbyname(BufString(buff))
    If local_ = 0 Then Exit Sub

    myAddr = HostentAddr(local_)

    my_tcpip_address = WINS_DottedAddr(myAddr)
End Sub

Public Function WINS_Init() As Long
    Dim i As Long
    Dim buff(0 To MAXHOSTNAMELEN - 1) As Byte
    Dim name As String, p As Long, c As String
    Dim r As Long

    net_acceptsocket = -1

    ' (C loads wsock32.dll here and gets the function vectors; the Declare
    ' statements stand for them)
    winsock_lib_initialized = True

    If COM_CheckParm("-noudp") Then
        WINS_Init = -1
        Exit Function
    End If

    If winsock_initialized = 0 Then
        r = pWSAStartup(&H101, winsockdata(0))  ' MAKEWORD(1, 1)

        If r Then
            Con_SafePrintf "Winsock initialization failed.\n"
            WINS_Init = -1
            Exit Function
        End If
    End If
    winsock_initialized = winsock_initialized + 1

    ' determine my name
    If pgethostname(buff(0), MAXHOSTNAMELEN) = SOCKET_ERROR Then
        Con_DPrintf "Winsock TCP/IP Initialization failed.\n"
        winsock_initialized = winsock_initialized - 1
        If winsock_initialized = 0 Then pWSACleanup
        WINS_Init = -1
        Exit Function
    End If
    name = BufString(buff)

    ' if the quake hostname isn't set, set it to the machine name
    If cvars(cv_hostname).strval = "UNNAMED" Then
        ' see if it's a text IP address (well, close enough)
        For p = 1 To Len(name)
            c = Mid$(name, p, 1)
            If (c < "0" Or c > "9") And c <> "." Then Exit For
        Next

        ' if it is a real name, strip off the domain; we only want the host
        If p <= Len(name) Then
            For i = 0 To 14
                If Mid$(name, i + 1, 1) = "." Then Exit For
            Next
            name = Left$(name, i)
        End If
        Cvar_Set "hostname", name
    End If

    i = COM_CheckParm("-ip")
    If i Then
        If i < com_argc - 1 Then
            myAddr = inet_addr(com_argv(i + 1))
            If myAddr = INADDR_NONE Then Sys_Error "%s is not a valid IP address", com_argv(i + 1)
            my_tcpip_address = com_argv(i + 1)
        Else
            Sys_Error "NET_Init: you must specify an IP address after -ip"
        End If
    ElseIf sys_hidden Then
        ' (QuakeVB) a test run opens no port on the network
        myAddr = inet_addr("127.0.0.1")
        my_tcpip_address = "127.0.0.1"
    Else
        myAddr = INADDR_ANY
        my_tcpip_address = "INADDR_ANY"
    End If

    net_controlsocket = WINS_OpenSocket(0)
    If net_controlsocket = -1 Then
        Con_Printf "WINS_Init: Unable to open control socket\n"
        winsock_initialized = winsock_initialized - 1
        If winsock_initialized = 0 Then pWSACleanup
        WINS_Init = -1
        Exit Function
    End If

    broadcastaddr.sa_family = AF_INET
    SetSinAddr broadcastaddr, INADDR_BROADCAST
    SetSinPort broadcastaddr, net_hostport

    Con_Printf "Winsock TCP/IP Initialized\n"
    tcpipAvailable = True

    WINS_Init = net_controlsocket
End Function

'=============================================================================

Public Sub WINS_Shutdown()
    WINS_Listen False
    WINS_CloseSocket net_controlsocket
    winsock_initialized = winsock_initialized - 1
    If winsock_initialized = 0 Then pWSACleanup
End Sub

'=============================================================================

Public Sub WINS_Listen(ByVal state As Boolean)
    ' enable listening
    If state Then
        If net_acceptsocket <> -1 Then Exit Sub
        WINS_GetLocalAddress
        net_acceptsocket = WINS_OpenSocket(net_hostport)
        If net_acceptsocket = -1 Then Sys_Error "WINS_Listen: Unable to open accept socket\n"
        Exit Sub
    End If

    ' disable listening
    If net_acceptsocket = -1 Then Exit Sub
    WINS_CloseSocket net_acceptsocket
    net_acceptsocket = -1
End Sub

'=============================================================================

Public Function WINS_OpenSocket(ByVal port As Long) As Long
    Dim newsocket As Long
    Dim address As qsockaddr
    Dim true_ As Long

    true_ = 1

    newsocket = psocket(PF_INET, SOCK_DGRAM, IPPROTO_UDP)
    If newsocket = -1 Then
        WINS_OpenSocket = -1
        Exit Function
    End If

    If pioctlsocket(newsocket, FIONBIO, true_) = -1 Then GoTo ErrorReturn

    address.sa_family = AF_INET
    SetSinAddr address, myAddr
    SetSinPort address, port
    If bind(newsocket, address, QSOCKADDR_SIZE) = 0 Then
        WINS_OpenSocket = newsocket
        Exit Function
    End If

    Sys_Error "Unable to bind to %s", WINS_AddrToString(address)
ErrorReturn:
    pclosesocket newsocket
    WINS_OpenSocket = -1
End Function

'=============================================================================

Public Function WINS_CloseSocket(ByVal socket As Long) As Long
    If socket = net_broadcastsocket Then net_broadcastsocket = 0
    WINS_CloseSocket = pclosesocket(socket)
End Function

'=============================================================================
' PartialIPAddress
'
' this lets you type only as much of the net address as required, using
' the local network components to fill in the rest

Private Function PartialIPAddress(ByVal in_ As String, hostaddr As qsockaddr) As Long
    Dim buff As String, b As Long, c As String
    Dim addr As Long, num As Long, mask As Long, run As Long, port As Long

    buff = "." & in_
    b = 1
    If Mid$(buff, 2, 1) = "." Then b = b + 1

    addr = 0
    mask = -1
    Do While Mid$(buff, b, 1) = "."
        b = b + 1
        num = 0
        run = 0
        Do
            c = Mid$(buff, b, 1)
            If c < "0" Or c > "9" Or c = "" Then Exit Do
            num = num * 10 + Asc(c) - 48
            b = b + 1
            run = run + 1
            If run > 3 Then
                PartialIPAddress = -1
                Exit Function
            End If
        Loop
        c = Mid$(buff, b, 1)
        If c <> "." And c <> ":" And c <> "" Then
            PartialIPAddress = -1
            Exit Function
        End If
        If num < 0 Or num > 255 Then
            PartialIPAddress = -1
            Exit Function
        End If
        mask = Shl8(mask)
        addr = Shl8(addr) + num
    Loop

    If Mid$(buff, b, 1) = ":" Then
        port = Q_atoi(Mid$(buff, b + 1))
    Else
        port = net_hostport
    End If

    hostaddr.sa_family = AF_INET
    SetSinPort hostaddr, port
    SetSinAddr hostaddr, (myAddr And BigLong(mask)) Or BigLong(addr)

    PartialIPAddress = 0
End Function

'=============================================================================

Public Function WINS_Connect(ByVal socket As Long, addr As qsockaddr) As Long
    WINS_Connect = 0
End Function

'=============================================================================

Public Function WINS_CheckNewConnections() As Long
    Dim buf(0 To 4095) As Byte

    If net_acceptsocket = -1 Then
        WINS_CheckNewConnections = -1
        Exit Function
    End If

    If precvfrom(net_acceptsocket, buf(0), 4096, MSG_PEEK, ByVal 0&, ByVal 0&) > 0 Then
        WINS_CheckNewConnections = net_acceptsocket
        Exit Function
    End If
    WINS_CheckNewConnections = -1
End Function

'=============================================================================

Public Function WINS_Read(ByVal socket As Long, buf() As Byte, ByVal length_ As Long, addr As qsockaddr) As Long
    Dim addrlen As Long, ret As Long, errno As Long

    addrlen = QSOCKADDR_SIZE
    ret = precvfrom(socket, buf(0), length_, 0, addr, addrlen)
    If ret = -1 Then
        errno = Err.LastDllError

        If errno = WSAEWOULDBLOCK Or errno = WSAECONNREFUSED Or errno = WSAECONNRESET Then
            WINS_Read = 0
            Exit Function
        End If
    End If
    WINS_Read = ret
End Function

'=============================================================================

Private Function WINS_MakeSocketBroadcastCapable(ByVal socket As Long) As Long
    Dim i As Long

    i = 1

    ' make this socket broadcast capable
    If psetsockopt(socket, SOL_SOCKET, SO_BROADCAST, i, 4) < 0 Then
        WINS_MakeSocketBroadcastCapable = -1
        Exit Function
    End If
    net_broadcastsocket = socket

    WINS_MakeSocketBroadcastCapable = 0
End Function

'=============================================================================

Public Function WINS_Broadcast(ByVal socket As Long, buf() As Byte, ByVal length_ As Long) As Long
    Dim ret As Long

    If socket <> net_broadcastsocket Then
        If net_broadcastsocket <> 0 Then Sys_Error "Attempted to use multiple broadcasts sockets\n"
        WINS_GetLocalAddress
        ret = WINS_MakeSocketBroadcastCapable(socket)
        If ret = -1 Then
            Con_Printf "Unable to make socket broadcast capable\n"
            WINS_Broadcast = ret
            Exit Function
        End If
    End If

    WINS_Broadcast = WINS_Write(socket, buf, length_, broadcastaddr)
End Function

'=============================================================================

Public Function WINS_Write(ByVal socket As Long, buf() As Byte, ByVal length_ As Long, addr As qsockaddr) As Long
    Dim ret As Long

    ret = psendto(socket, buf(0), length_, 0, addr, QSOCKADDR_SIZE)
    If ret = -1 Then
        If Err.LastDllError = WSAEWOULDBLOCK Then
            WINS_Write = 0
            Exit Function
        End If
    End If

    WINS_Write = ret
End Function

'=============================================================================

Public Function WINS_AddrToString(addr As qsockaddr) As String
    WINS_AddrToString = WINS_DottedAddr(SinAddr(addr)) & ":" & SinPort(addr)
End Function

'=============================================================================

' (C: sscanf "%d.%d.%d.%d:%d"; the numbers it does not find are 0 here)
Public Function WINS_StringToAddr(ByVal s As String, addr As qsockaddr) As Long
    Dim ha(0 To 4) As Long, n As Long, p As Long, c As String
    Dim ipaddr As Long

    p = 1
    For n = 0 To 4
        ha(n) = Q_atoi(Mid$(s, p))
        Do While p <= Len(s)
            c = Mid$(s, p, 1)
            p = p + 1
            If c = "." Or c = ":" Then Exit Do
        Loop
    Next
    ipaddr = Shl8(Shl8(Shl8(ha(0) And 255) Or (ha(1) And 255)) Or (ha(2) And 255)) Or (ha(3) And 255)

    addr.sa_family = AF_INET
    SetSinAddr addr, BigLong(ipaddr)
    SetSinPort addr, ha(4)
    WINS_StringToAddr = 0
End Function

'=============================================================================

Public Function WINS_GetSocketAddr(ByVal socket As Long, addr As qsockaddr) As Long
    Dim addrlen As Long, a As Long

    addrlen = QSOCKADDR_SIZE
    addr.sa_family = 0
    ZeroMemory addr.sa_data(0), 14
    pgetsockname socket, addr, addrlen
    a = SinAddr(addr)
    If a = 0 Or a = inet_addr("127.0.0.1") Then SetSinAddr addr, myAddr

    WINS_GetSocketAddr = 0
End Function

'=============================================================================

Public Function WINS_GetNameFromAddr(addr As qsockaddr, name As String) As Long
    Dim hostentry As Long, p As Long

    hostentry = pgethostbyaddr(addr.sa_data(2), 4, AF_INET)
    If hostentry Then
        CopyMemory p, ByVal hostentry, 4         ' h_name
        name = Left$(CStringAt(p), NET_NAMELEN - 1)
        WINS_GetNameFromAddr = 0
        Exit Function
    End If

    name = WINS_AddrToString(addr)
    WINS_GetNameFromAddr = 0
End Function

'=============================================================================

Public Function WINS_GetAddrFromName(ByVal name As String, addr As qsockaddr) As Long
    Dim hostentry As Long, c As String

    c = Left$(name, 1)
    If c >= "0" And c <= "9" Then
        WINS_GetAddrFromName = PartialIPAddress(name, addr)
        Exit Function
    End If

    hostentry = pgethostbyname(name)
    If hostentry = 0 Then
        WINS_GetAddrFromName = -1
        Exit Function
    End If

    addr.sa_family = AF_INET
    SetSinPort addr, net_hostport
    SetSinAddr addr, HostentAddr(hostentry)

    WINS_GetAddrFromName = 0
End Function

'=============================================================================

Public Function WINS_AddrCompare(addr1 As qsockaddr, addr2 As qsockaddr) As Long
    If addr1.sa_family <> addr2.sa_family Then
        WINS_AddrCompare = -1
        Exit Function
    End If

    If SinAddr(addr1) <> SinAddr(addr2) Then
        WINS_AddrCompare = -1
        Exit Function
    End If

    If SinPort(addr1) <> SinPort(addr2) Then
        WINS_AddrCompare = 1
        Exit Function
    End If

    WINS_AddrCompare = 0
End Function

'=============================================================================

Public Function WINS_GetSocketPort(addr As qsockaddr) As Long
    WINS_GetSocketPort = SinPort(addr)
End Function

Public Function WINS_SetSocketPort(addr As qsockaddr, ByVal port As Long) As Long
    SetSinPort addr, port
    WINS_SetSocketPort = 0
End Function
