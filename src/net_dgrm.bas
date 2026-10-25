Attribute VB_Name = "net_dgrm"
' net_dgrm.bas - the datagram network driver: reliable and unreliable
' messages, connections and the server search over a LAN driver (net_dgrm.c)
'
' C calls the LAN drivers through net_landrivers[], a table of function
' pointers (dfunc is net_landrivers[net_landriverlevel], sfunc the socket's
' net_landrivers[sock->landriver]); here the Lan_ functions dispatch on the
' LAN driver number: 0 is Winsock TCP/IP (net_wins.bas). There is no IPX
' driver (net_wipx.c): Windows has had no IPX since Vista. packetBuffer is a
' byte array: the length and sequence words, in network byte order, then
' the data at NET_HEADERSIZE.
'
' Deviation: a packet whose header gives a length shorter than the header or
' longer than what arrived, or more reliable data than the receive buffer
' holds, is dropped as a short packet (C copies what the header says).
' Test aid: -netdrop n loses every n-th packet a connection reads (where C
' has its random drop commented out), to exercise the resends.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' This is enables a simple IP banning mechanism (C: BAN_TEST)
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long

Private Const AF_INET As Long = 2

Private net_landriverlevel As Long

' statistic counters
Private packetsSent As Long
Private packetsReSent As Long
Private packetsReceived As Long
Private receivedDuplicateCount As Long
Private shortPacketCount As Long
Private droppedDatagrams As Long

Private myDriverLevel As Long

Private packetBuffer(0 To NET_DATAGRAMSIZE - 1) As Byte

Private banAddr As Long                  ' = 0x00000000
Private banMask As Long                  ' = 0xffffffff (set in Datagram_Init)

Private testInProgress As Boolean
Private testPollCount As Long
Private testDriver As Long
Private testSocket As Long

Private test2InProgress As Boolean
Private test2Driver As Long
Private test2Socket As Long

Private netdrop As Long                  ' (QuakeVB) -netdrop n
Private netdrop_count As Long

' (QuakeVB) C: *(int *)buf = BigLong (l), and the other way
Private Sub PutBigLong(b() As Byte, ByVal ofs As Long, ByVal l As Long)
    PutLong b, ofs, BigLong(l)
End Sub

Private Function GetBigLong(b() As Byte, ByVal ofs As Long) As Long
    GetBigLong = BigLong(GetLong(b, ofs))
End Function

Private Sub NET_Ban_f()
    Dim addrStr As String, maskStr As String
    Dim toclient As Boolean              ' C: print is SV_ClientPrintf, else Con_Printf

    If cmd_source = src_command Then
        If Not sv.active Then
            Cmd_ForwardToServer
            Exit Sub
        End If
    Else
        If pr_global_struct(0).deathmatch <> 0 And Not svs.clients(host_client).privileged Then Exit Sub
        toclient = True
    End If

    Select Case Cmd_Argc()
    Case 1
        If banAddr Then
            addrStr = WINS_DottedAddr(banAddr)
            maskStr = WINS_DottedAddr(banMask)
            BanPrint toclient, va("Banning %s [%s]\n", addrStr, maskStr)
        Else
            BanPrint toclient, va("Banning not active\n")
        End If

    Case 2
        If Q_strcasecmp(Cmd_Argv(1), "off") = 0 Then
            banAddr = 0
        Else
            banAddr = inet_addr(Cmd_Argv(1))
        End If
        banMask = &HFFFFFFFF

    Case 3
        banAddr = inet_addr(Cmd_Argv(1))
        banMask = inet_addr(Cmd_Argv(2))

    Case Else
        BanPrint toclient, va("BAN ip_address [mask]\n")
    End Select
End Sub

Private Sub BanPrint(ByVal toclient As Boolean, ByVal s As String)
    If toclient Then
        SV_ClientPrintf "%s", s
    Else
        Con_Printf "%s", s
    End If
End Sub

Public Function Datagram_SendMessage(ByVal sock As Long, data As sizebuf_t) As Long
    Dim packetLen As Long, dataLen As Long, eom As Long

    With net_qsockets(sock)
        If data.cursize > 0 Then CopyMemory .sendMessage(0), data.data(0), data.cursize
        .sendMessageLength = data.cursize

        If data.cursize <= MAX_DATAGRAM Then
            dataLen = data.cursize
            eom = NETFLAG_EOM
        Else
            dataLen = MAX_DATAGRAM
            eom = 0
        End If
        packetLen = NET_HEADERSIZE + dataLen

        PutBigLong packetBuffer, 0, packetLen Or (NETFLAG_DATA Or eom)
        PutBigLong packetBuffer, 4, .sendSequence
        .sendSequence = .sendSequence + 1
        If dataLen > 0 Then CopyMemory packetBuffer(NET_HEADERSIZE), .sendMessage(0), dataLen

        .canSend = False

        If Lan_Write(.landriver, .socket, packetBuffer, packetLen, .addr) = -1 Then
            Datagram_SendMessage = -1
            Exit Function
        End If

        .lastSendTime = net_time
    End With
    packetsSent = packetsSent + 1
    Datagram_SendMessage = 1
End Function

Private Function SendMessageNext(ByVal sock As Long) As Long
    Dim packetLen As Long, dataLen As Long, eom As Long

    With net_qsockets(sock)
        If .sendMessageLength <= MAX_DATAGRAM Then
            dataLen = .sendMessageLength
            eom = NETFLAG_EOM
        Else
            dataLen = MAX_DATAGRAM
            eom = 0
        End If
        packetLen = NET_HEADERSIZE + dataLen

        PutBigLong packetBuffer, 0, packetLen Or (NETFLAG_DATA Or eom)
        PutBigLong packetBuffer, 4, .sendSequence
        .sendSequence = .sendSequence + 1
        If dataLen > 0 Then CopyMemory packetBuffer(NET_HEADERSIZE), .sendMessage(0), dataLen

        .sendNext = False

        If Lan_Write(.landriver, .socket, packetBuffer, packetLen, .addr) = -1 Then
            SendMessageNext = -1
            Exit Function
        End If

        .lastSendTime = net_time
    End With
    packetsSent = packetsSent + 1
    SendMessageNext = 1
End Function

Private Function ReSendMessage(ByVal sock As Long) As Long
    Dim packetLen As Long, dataLen As Long, eom As Long

    With net_qsockets(sock)
        If .sendMessageLength <= MAX_DATAGRAM Then
            dataLen = .sendMessageLength
            eom = NETFLAG_EOM
        Else
            dataLen = MAX_DATAGRAM
            eom = 0
        End If
        packetLen = NET_HEADERSIZE + dataLen

        PutBigLong packetBuffer, 0, packetLen Or (NETFLAG_DATA Or eom)
        PutBigLong packetBuffer, 4, .sendSequence - 1
        If dataLen > 0 Then CopyMemory packetBuffer(NET_HEADERSIZE), .sendMessage(0), dataLen

        .sendNext = False

        If Lan_Write(.landriver, .socket, packetBuffer, packetLen, .addr) = -1 Then
            ReSendMessage = -1
            Exit Function
        End If

        .lastSendTime = net_time
    End With
    packetsReSent = packetsReSent + 1
    ReSendMessage = 1
End Function

Public Function Datagram_CanSendMessage(ByVal sock As Long) As Boolean
    If net_qsockets(sock).sendNext Then SendMessageNext sock

    Datagram_CanSendMessage = net_qsockets(sock).canSend
End Function

Public Function Datagram_CanSendUnreliableMessage(ByVal sock As Long) As Boolean
    Datagram_CanSendUnreliableMessage = True
End Function

Public Function Datagram_SendUnreliableMessage(ByVal sock As Long, data As sizebuf_t) As Long
    Dim packetLen As Long

    packetLen = NET_HEADERSIZE + data.cursize

    With net_qsockets(sock)
        PutBigLong packetBuffer, 0, packetLen Or NETFLAG_UNRELIABLE
        PutBigLong packetBuffer, 4, .unreliableSendSequence
        .unreliableSendSequence = .unreliableSendSequence + 1
        If data.cursize > 0 Then CopyMemory packetBuffer(NET_HEADERSIZE), data.data(0), data.cursize

        If Lan_Write(.landriver, .socket, packetBuffer, packetLen, .addr) = -1 Then
            Datagram_SendUnreliableMessage = -1
            Exit Function
        End If
    End With

    packetsSent = packetsSent + 1
    Datagram_SendUnreliableMessage = 1
End Function

Public Function Datagram_GetMessage(ByVal sock As Long) As Long
    Dim length_ As Long, received As Long, flags As Long, ret As Long
    Dim readaddr As qsockaddr
    Dim sequence As Long, count As Long, p As Long

    With net_qsockets(sock)
        If Not .canSend Then
            If (net_time - .lastSendTime) > 1# Then ReSendMessage sock
        End If

        Do
            received = Lan_Read(.landriver, .socket, packetBuffer, NET_DATAGRAMSIZE, readaddr)

            ' if ((rand() & 255) > 220)
            '     continue;

            If received = 0 Then Exit Do

            If received = -1 Then
                Con_Printf "Read error\n"
                Datagram_GetMessage = -1
                Exit Function
            End If

            ' (QuakeVB) test aid: -netdrop n loses every n-th packet
            If netdrop > 0 Then
                netdrop_count = netdrop_count + 1
                If netdrop_count Mod netdrop = 0 Then GoTo continue_
            End If

            If Lan_AddrCompare(.landriver, readaddr, .addr) <> 0 Then GoTo continue_

            If received < NET_HEADERSIZE Then
                shortPacketCount = shortPacketCount + 1
                GoTo continue_
            End If

            length_ = GetBigLong(packetBuffer, 0)
            flags = length_ And (Not NETFLAG_LENGTH_MASK)
            length_ = length_ And NETFLAG_LENGTH_MASK

            If flags And NETFLAG_CTL Then GoTo continue_

            ' (QuakeVB) a length the packet does not hold
            If length_ < NET_HEADERSIZE Or length_ > received Then
                shortPacketCount = shortPacketCount + 1
                GoTo continue_
            End If

            sequence = GetBigLong(packetBuffer, 4)
            packetsReceived = packetsReceived + 1

            If flags And NETFLAG_UNRELIABLE Then
                If sequence < .unreliableReceiveSequence Then
                    Con_DPrintf "Got a stale datagram\n"
                    ret = 0
                    Exit Do
                End If
                If sequence <> .unreliableReceiveSequence Then
                    count = sequence - .unreliableReceiveSequence
                    droppedDatagrams = droppedDatagrams + count
                    Con_DPrintf "Dropped %u datagram(s)\n", count
                End If
                .unreliableReceiveSequence = sequence + 1

                length_ = length_ - NET_HEADERSIZE

                SZ_Clear net_message
                SZ_Write net_message, packetBuffer, NET_HEADERSIZE, length_

                ret = 2
                Exit Do
            End If

            If flags And NETFLAG_ACK Then
                If sequence <> (.sendSequence - 1) Then
                    Con_DPrintf "Stale ACK received\n"
                    GoTo continue_
                End If
                If sequence = .ackSequence Then
                    .ackSequence = .ackSequence + 1
                    If .ackSequence <> .sendSequence Then Con_DPrintf "ack sequencing error\n"
                Else
                    Con_DPrintf "Duplicate ACK received\n"
                    GoTo continue_
                End If
                .sendMessageLength = .sendMessageLength - MAX_DATAGRAM
                If .sendMessageLength > 0 Then
                    CopyMemory .sendMessage(0), .sendMessage(MAX_DATAGRAM), .sendMessageLength
                    .sendNext = True
                Else
                    .sendMessageLength = 0
                    .canSend = True
                End If
                GoTo continue_
            End If

            If flags And NETFLAG_DATA Then
                ' (QuakeVB) more than a message can hold
                If .receiveMessageLength + length_ - NET_HEADERSIZE > NET_MAXMESSAGE Then
                    shortPacketCount = shortPacketCount + 1
                    GoTo continue_
                End If

                PutBigLong packetBuffer, 0, NET_HEADERSIZE Or NETFLAG_ACK
                PutBigLong packetBuffer, 4, sequence
                Lan_Write .landriver, .socket, packetBuffer, NET_HEADERSIZE, readaddr

                If sequence <> .receiveSequence Then
                    receivedDuplicateCount = receivedDuplicateCount + 1
                    GoTo continue_
                End If
                .receiveSequence = .receiveSequence + 1

                length_ = length_ - NET_HEADERSIZE

                If flags And NETFLAG_EOM Then
                    SZ_Clear net_message
                    p = SZ_GetSpace(net_message, .receiveMessageLength)
                    If .receiveMessageLength > 0 Then CopyMemory net_message.data(p), .receiveMessage(0), .receiveMessageLength
                    SZ_Write net_message, packetBuffer, NET_HEADERSIZE, length_
                    .receiveMessageLength = 0

                    ret = 1
                    Exit Do
                End If

                If length_ > 0 Then CopyMemory .receiveMessage(.receiveMessageLength), packetBuffer(NET_HEADERSIZE), length_
                .receiveMessageLength = .receiveMessageLength + length_
                GoTo continue_
            End If
continue_:
        Loop

        If .sendNext Then SendMessageNext sock
    End With

    Datagram_GetMessage = ret
End Function

Private Sub PrintStats(ByVal s As Long)
    Con_Printf "canSend = %4u   \n", -CLng(net_qsockets(s).canSend)
    Con_Printf "sendSeq = %4u   ", net_qsockets(s).sendSequence
    Con_Printf "recvSeq = %4u   \n", net_qsockets(s).receiveSequence
    Con_Printf "\n"
End Sub

Private Sub NET_Stats_f()
    Dim s As Long

    If Cmd_Argc() = 1 Then
        Con_Printf "unreliable messages sent   = %i\n", unreliableMessagesSent
        Con_Printf "unreliable messages recv   = %i\n", unreliableMessagesReceived
        Con_Printf "reliable messages sent     = %i\n", messagesSent
        Con_Printf "reliable messages received = %i\n", messagesReceived
        Con_Printf "packetsSent                = %i\n", packetsSent
        Con_Printf "packetsReSent              = %i\n", packetsReSent
        Con_Printf "packetsReceived            = %i\n", packetsReceived
        Con_Printf "receivedDuplicateCount     = %i\n", receivedDuplicateCount
        Con_Printf "shortPacketCount           = %i\n", shortPacketCount
        Con_Printf "droppedDatagrams           = %i\n", droppedDatagrams
    ElseIf Cmd_Argv(1) = "*" Then
        s = net_activeSockets
        Do While s <> 0
            PrintStats s
            s = net_qsockets(s).nxt
        Loop
        s = net_freeSockets
        Do While s <> 0
            PrintStats s
            s = net_qsockets(s).nxt
        Loop
    Else
        s = net_activeSockets
        Do While s <> 0
            If Q_strcasecmp(Cmd_Argv(1), net_qsockets(s).address) = 0 Then Exit Do
            s = net_qsockets(s).nxt
        Loop
        If s = 0 Then
            s = net_freeSockets
            Do While s <> 0
                If Q_strcasecmp(Cmd_Argv(1), net_qsockets(s).address) = 0 Then Exit Do
                s = net_qsockets(s).nxt
            Loop
        End If
        If s = 0 Then Exit Sub
        PrintStats s
    End If
End Sub

Public Sub Test_Poll()
    Dim clientaddr As qsockaddr
    Dim control As Long, len_ As Long
    Dim name As String, address As String
    Dim colors As Long, frags As Long, connectTime As Long
    Dim playerNumber As Long

    net_landriverlevel = testDriver

    Do
        len_ = Lan_Read(net_landriverlevel, testSocket, net_message.data, net_message.maxsize, clientaddr)
        If len_ < 4 Then Exit Do

        net_message.cursize = len_

        MSG_BeginReading
        control = GetBigLong(net_message.data, 0)
        MSG_ReadLong
        If control = -1 Then Exit Do
        If (control And (Not NETFLAG_LENGTH_MASK)) <> NETFLAG_CTL Then Exit Do
        If (control And NETFLAG_LENGTH_MASK) <> len_ Then Exit Do

        If MSG_ReadByte() <> CCREP_PLAYER_INFO Then Sys_Error "Unexpected repsonse to Player Info request\n"

        playerNumber = MSG_ReadByte()
        name = MSG_ReadString()
        colors = MSG_ReadLong()
        frags = MSG_ReadLong()
        connectTime = MSG_ReadLong()
        address = MSG_ReadString()

        Con_Printf "%s\n  frags:%3i  colors:%u %u  time:%u\n  %s\n", name, frags, colors \ 16, colors And &HF&, connectTime \ 60, address
    Loop

    testPollCount = testPollCount - 1
    If testPollCount Then
        SchedulePollProcedure PP_TEST_POLL, 0.1
    Else
        Lan_CloseSocket net_landriverlevel, testSocket
        testInProgress = False
    End If
End Sub

Private Sub Test_f()
    Dim host As String
    Dim n As Long, max As Long
    Dim sendaddr As qsockaddr

    max = MAX_SCOREBOARD

    If testInProgress Then Exit Sub

    host = Cmd_Argv(1)

    If hostCacheCount Then
        For n = 0 To hostCacheCount - 1
            If Q_strcasecmp(host, hostcache(n).name) = 0 Then
                If hostcache(n).driver = myDriverLevel Then
                    net_landriverlevel = hostcache(n).ldriver
                    max = hostcache(n).maxusers
                    sendaddr = hostcache(n).addr
                    Exit For
                End If
            End If
        Next
        If n < hostCacheCount Then GoTo JustDoIt
    End If

    For net_landriverlevel = 0 To net_numlandrivers - 1
        If net_landrivers(net_landriverlevel).initialized Then
            ' see if we can resolve the host name
            If Lan_GetAddrFromName(net_landriverlevel, host, sendaddr) <> -1 Then Exit For
        End If
    Next
    If net_landriverlevel = net_numlandrivers Then Exit Sub

JustDoIt:
    testSocket = Lan_OpenSocket(net_landriverlevel, 0)
    If testSocket = -1 Then Exit Sub

    testInProgress = True
    testPollCount = 20
    testDriver = net_landriverlevel

    For n = 0 To max - 1
        SZ_Clear net_message
        ' save space for the header, filled in later
        MSG_WriteLong net_message, 0
        MSG_WriteByte net_message, CCREQ_PLAYER_INFO
        MSG_WriteByte net_message, n
        PutBigLong net_message.data, 0, NETFLAG_CTL Or (net_message.cursize And NETFLAG_LENGTH_MASK)
        Lan_Write net_landriverlevel, testSocket, net_message.data, net_message.cursize, sendaddr
    Next
    SZ_Clear net_message
    SchedulePollProcedure PP_TEST_POLL, 0.1
End Sub

Public Sub Test2_Poll()
    Dim clientaddr As qsockaddr
    Dim control As Long, len_ As Long
    Dim name As String, value As String

    net_landriverlevel = test2Driver
    name = ""

    len_ = Lan_Read(net_landriverlevel, test2Socket, net_message.data, net_message.maxsize, clientaddr)
    If len_ < 4 Then GoTo Reschedule

    net_message.cursize = len_

    MSG_BeginReading
    control = GetBigLong(net_message.data, 0)
    MSG_ReadLong
    If control = -1 Then GoTo Error_
    If (control And (Not NETFLAG_LENGTH_MASK)) <> NETFLAG_CTL Then GoTo Error_
    If (control And NETFLAG_LENGTH_MASK) <> len_ Then GoTo Error_

    If MSG_ReadByte() <> CCREP_RULE_INFO Then GoTo Error_

    name = MSG_ReadString()
    If Len(name) = 0 Then GoTo Done
    value = MSG_ReadString()

    Con_Printf "%-16.16s  %-16.16s\n", name, value

    SZ_Clear net_message
    ' save space for the header, filled in later
    MSG_WriteLong net_message, 0
    MSG_WriteByte net_message, CCREQ_RULE_INFO
    MSG_WriteString net_message, name
    PutBigLong net_message.data, 0, NETFLAG_CTL Or (net_message.cursize And NETFLAG_LENGTH_MASK)
    Lan_Write net_landriverlevel, test2Socket, net_message.data, net_message.cursize, clientaddr
    SZ_Clear net_message

Reschedule:
    SchedulePollProcedure PP_TEST2_POLL, 0.05
    Exit Sub

Error_:
    Con_Printf "Unexpected repsonse to Rule Info request\n"
Done:
    Lan_CloseSocket net_landriverlevel, test2Socket
    test2InProgress = False
End Sub

Private Sub Test2_f()
    Dim host As String
    Dim n As Long
    Dim sendaddr As qsockaddr

    If test2InProgress Then Exit Sub

    host = Cmd_Argv(1)

    If hostCacheCount Then
        For n = 0 To hostCacheCount - 1
            If Q_strcasecmp(host, hostcache(n).name) = 0 Then
                If hostcache(n).driver = myDriverLevel Then
                    net_landriverlevel = hostcache(n).ldriver
                    sendaddr = hostcache(n).addr
                    Exit For
                End If
            End If
        Next
        If n < hostCacheCount Then GoTo JustDoIt
    End If

    For net_landriverlevel = 0 To net_numlandrivers - 1
        If net_landrivers(net_landriverlevel).initialized Then
            ' see if we can resolve the host name
            If Lan_GetAddrFromName(net_landriverlevel, host, sendaddr) <> -1 Then Exit For
        End If
    Next
    If net_landriverlevel = net_numlandrivers Then Exit Sub

JustDoIt:
    test2Socket = Lan_OpenSocket(net_landriverlevel, 0)
    If test2Socket = -1 Then Exit Sub

    test2InProgress = True
    test2Driver = net_landriverlevel

    SZ_Clear net_message
    ' save space for the header, filled in later
    MSG_WriteLong net_message, 0
    MSG_WriteByte net_message, CCREQ_RULE_INFO
    MSG_WriteString net_message, ""
    PutBigLong net_message.data, 0, NETFLAG_CTL Or (net_message.cursize And NETFLAG_LENGTH_MASK)
    Lan_Write net_landriverlevel, test2Socket, net_message.data, net_message.cursize, sendaddr
    SZ_Clear net_message
    SchedulePollProcedure PP_TEST2_POLL, 0.05
End Sub

Public Function Datagram_Init() As Long
    Dim i As Long, csock As Long

    banMask = &HFFFFFFFF

    ' (net_win.c's LAN driver table)
    net_landrivers(0).name = "Winsock TCPIP"
    net_numlandrivers = 1

    myDriverLevel = net_driverlevel
    Cmd_AddCommand "net_stats", XC_NET_Stats_f

    i = COM_CheckParm("-netdrop")
    If i > 0 And i < com_argc - 1 Then netdrop = Q_atoi(com_argv(i + 1))

    If COM_CheckParm("-nolan") Then
        Datagram_Init = -1
        Exit Function
    End If

    For i = 0 To net_numlandrivers - 1
        csock = Lan_Init(i)
        If csock <> -1 Then
            net_landrivers(i).initialized = True
            net_landrivers(i).controlSock = csock
        End If
    Next

    Cmd_AddCommand "ban", XC_NET_Ban_f
    Cmd_AddCommand "test", XC_Test_f
    Cmd_AddCommand "test2", XC_Test2_f

    Datagram_Init = 0
End Function

Public Sub Datagram_Shutdown()
    Dim i As Long

    '
    ' shutdown the lan drivers
    '
    For i = 0 To net_numlandrivers - 1
        If net_landrivers(i).initialized Then
            Lan_Shutdown i
            net_landrivers(i).initialized = False
        End If
    Next
End Sub

Public Sub Datagram_Close(ByVal sock As Long)
    Lan_CloseSocket net_qsockets(sock).landriver, net_qsockets(sock).socket
End Sub

Public Sub Datagram_Listen(ByVal state As Boolean)
    Dim i As Long

    For i = 0 To net_numlandrivers - 1
        If net_landrivers(i).initialized Then Lan_Listen i, state
    Next
End Sub

' C: _Datagram_CheckNewConnections; returns a net_qsockets() index, 0 = NULL
Private Function DatagramCheckNewConnections() As Long
    Dim clientaddr As qsockaddr, newaddr As qsockaddr
    Dim newsock As Long, acceptsock As Long
    Dim sock As Long, s As Long
    Dim len_ As Long, command As Long, control As Long, ret As Long
    Dim playerNumber As Long, activeNumber As Long, clientNumber As Long
    Dim prevCvarName As String, var As Long
    Dim testAddr As Long

    acceptsock = Lan_CheckNewConnections(net_landriverlevel)
    If acceptsock = -1 Then Exit Function

    SZ_Clear net_message

    len_ = Lan_Read(net_landriverlevel, acceptsock, net_message.data, net_message.maxsize, clientaddr)
    If len_ < 4 Then Exit Function
    net_message.cursize = len_

    MSG_BeginReading
    control = GetBigLong(net_message.data, 0)
    MSG_ReadLong
    If control = -1 Then Exit Function
    If (control And (Not NETFLAG_LENGTH_MASK)) <> NETFLAG_CTL Then Exit Function
    If (control And NETFLAG_LENGTH_MASK) <> len_ Then Exit Function

    command = MSG_ReadByte()
    If command = CCREQ_SERVER_INFO Then
        If MSG_ReadString() <> "QUAKE" Then Exit Function

        SZ_Clear net_message
        ' save space for the header, filled in later
        MSG_WriteLong net_message, 0
        MSG_WriteByte net_message, CCREP_SERVER_INFO
        Lan_GetSocketAddr net_landriverlevel, acceptsock, newaddr
        MSG_WriteString net_message, Lan_AddrToString(net_landriverlevel, newaddr)
        MSG_WriteString net_message, cvars(cv_hostname).strval
        MSG_WriteString net_message, sv.name
        MSG_WriteByte net_message, net_activeconnections
        MSG_WriteByte net_message, svs.maxclients
        MSG_WriteByte net_message, NET_PROTOCOL_VERSION
        PutBigLong net_message.data, 0, NETFLAG_CTL Or (net_message.cursize And NETFLAG_LENGTH_MASK)
        Lan_Write net_landriverlevel, acceptsock, net_message.data, net_message.cursize, clientaddr
        SZ_Clear net_message
        Exit Function
    End If

    If command = CCREQ_PLAYER_INFO Then
        playerNumber = MSG_ReadByte()
        activeNumber = -1
        For clientNumber = 0 To svs.maxclients - 1
            If svs.clients(clientNumber).active Then
                activeNumber = activeNumber + 1
                If activeNumber = playerNumber Then Exit For
            End If
        Next
        If clientNumber = svs.maxclients Then Exit Function

        SZ_Clear net_message
        ' save space for the header, filled in later
        MSG_WriteLong net_message, 0
        MSG_WriteByte net_message, CCREP_PLAYER_INFO
        MSG_WriteByte net_message, playerNumber
        MSG_WriteString net_message, svs.clients(clientNumber).name
        MSG_WriteLong net_message, svs.clients(clientNumber).colors
        MSG_WriteLong net_message, ev(svs.clients(clientNumber).edict).frags
        MSG_WriteLong net_message, net_time - net_qsockets(svs.clients(clientNumber).netconnection).connecttime
        MSG_WriteString net_message, net_qsockets(svs.clients(clientNumber).netconnection).address
        PutBigLong net_message.data, 0, NETFLAG_CTL Or (net_message.cursize And NETFLAG_LENGTH_MASK)
        Lan_Write net_landriverlevel, acceptsock, net_message.data, net_message.cursize, clientaddr
        SZ_Clear net_message

        Exit Function
    End If

    If command = CCREQ_RULE_INFO Then
        ' find the search start location
        prevCvarName = MSG_ReadString()
        If Len(prevCvarName) Then
            var = Cvar_FindVar(prevCvarName)
            If var = 0 Then Exit Function
            var = var - 1                    ' C: var->next
        Else
            var = num_cvars                  ' C: cvar_vars
        End If

        ' search for the next server cvar
        Do While var > 0
            If cvars(var).server Then Exit Do
            var = var - 1
        Loop

        ' send the response

        SZ_Clear net_message
        ' save space for the header, filled in later
        MSG_WriteLong net_message, 0
        MSG_WriteByte net_message, CCREP_RULE_INFO
        If var > 0 Then
            MSG_WriteString net_message, cvars(var).name
            MSG_WriteString net_message, cvars(var).strval
        End If
        PutBigLong net_message.data, 0, NETFLAG_CTL Or (net_message.cursize And NETFLAG_LENGTH_MASK)
        Lan_Write net_landriverlevel, acceptsock, net_message.data, net_message.cursize, clientaddr
        SZ_Clear net_message

        Exit Function
    End If

    If command <> CCREQ_CONNECT Then Exit Function

    If MSG_ReadString() <> "QUAKE" Then Exit Function

    If MSG_ReadByte() <> NET_PROTOCOL_VERSION Then
        SZ_Clear net_message
        ' save space for the header, filled in later
        MSG_WriteLong net_message, 0
        MSG_WriteByte net_message, CCREP_REJECT
        MSG_WriteString net_message, "Incompatible version." & vbLf
        PutBigLong net_message.data, 0, NETFLAG_CTL Or (net_message.cursize And NETFLAG_LENGTH_MASK)
        Lan_Write net_landriverlevel, acceptsock, net_message.data, net_message.cursize, clientaddr
        SZ_Clear net_message
        Exit Function
    End If

    ' check for a ban
    If clientaddr.sa_family = AF_INET Then
        CopyMemory testAddr, clientaddr.sa_data(2), 4     ' sin_addr.s_addr
        If (testAddr And banMask) = banAddr Then
            SZ_Clear net_message
            ' save space for the header, filled in later
            MSG_WriteLong net_message, 0
            MSG_WriteByte net_message, CCREP_REJECT
            MSG_WriteString net_message, "You have been banned." & vbLf
            PutBigLong net_message.data, 0, NETFLAG_CTL Or (net_message.cursize And NETFLAG_LENGTH_MASK)
            Lan_Write net_landriverlevel, acceptsock, net_message.data, net_message.cursize, clientaddr
            SZ_Clear net_message
            Exit Function
        End If
    End If

    ' see if this guy is already connected
    s = net_activeSockets
    Do While s <> 0
        If net_qsockets(s).driver = net_driverlevel Then
            ret = Lan_AddrCompare(net_landriverlevel, clientaddr, net_qsockets(s).addr)
            If ret >= 0 Then
                ' is this a duplicate connection reqeust?
                If ret = 0 And net_time - net_qsockets(s).connecttime < 2# Then
                    ' yes, so send a duplicate reply
                    SZ_Clear net_message
                    ' save space for the header, filled in later
                    MSG_WriteLong net_message, 0
                    MSG_WriteByte net_message, CCREP_ACCEPT
                    Lan_GetSocketAddr net_landriverlevel, net_qsockets(s).socket, newaddr
                    MSG_WriteLong net_message, Lan_GetSocketPort(net_landriverlevel, newaddr)
                    PutBigLong net_message.data, 0, NETFLAG_CTL Or (net_message.cursize And NETFLAG_LENGTH_MASK)
                    Lan_Write net_landriverlevel, acceptsock, net_message.data, net_message.cursize, clientaddr
                    SZ_Clear net_message
                    Exit Function
                End If
                ' it's somebody coming back in from a crash/disconnect
                ' so close the old qsocket and let their retry get them back in
                NET_Close s
                Exit Function
            End If
        End If
        s = net_qsockets(s).nxt
    Loop

    ' allocate a QSocket
    sock = NET_NewQSocket()
    If sock = 0 Then
        ' no room; try to let him know
        SZ_Clear net_message
        ' save space for the header, filled in later
        MSG_WriteLong net_message, 0
        MSG_WriteByte net_message, CCREP_REJECT
        MSG_WriteString net_message, "Server is full." & vbLf
        PutBigLong net_message.data, 0, NETFLAG_CTL Or (net_message.cursize And NETFLAG_LENGTH_MASK)
        Lan_Write net_landriverlevel, acceptsock, net_message.data, net_message.cursize, clientaddr
        SZ_Clear net_message
        Exit Function
    End If

    ' allocate a network socket
    newsock = Lan_OpenSocket(net_landriverlevel, 0)
    If newsock = -1 Then
        NET_FreeQSocket sock
        Exit Function
    End If

    ' connect to the client
    If Lan_Connect(net_landriverlevel, newsock, clientaddr) = -1 Then
        Lan_CloseSocket net_landriverlevel, newsock
        NET_FreeQSocket sock
        Exit Function
    End If

    ' everything is allocated, just fill in the details
    net_qsockets(sock).socket = newsock
    net_qsockets(sock).landriver = net_landriverlevel
    net_qsockets(sock).addr = clientaddr
    net_qsockets(sock).address = Lan_AddrToString(net_landriverlevel, clientaddr)

    ' send him back the info about the server connection he has been allocated
    SZ_Clear net_message
    ' save space for the header, filled in later
    MSG_WriteLong net_message, 0
    MSG_WriteByte net_message, CCREP_ACCEPT
    Lan_GetSocketAddr net_landriverlevel, newsock, newaddr
    MSG_WriteLong net_message, Lan_GetSocketPort(net_landriverlevel, newaddr)
    ' MSG_WriteString(&net_message, dfunc.AddrToString(&newaddr));
    PutBigLong net_message.data, 0, NETFLAG_CTL Or (net_message.cursize And NETFLAG_LENGTH_MASK)
    Lan_Write net_landriverlevel, acceptsock, net_message.data, net_message.cursize, clientaddr
    SZ_Clear net_message

    DatagramCheckNewConnections = sock
End Function

' returns a net_qsockets() index, 0 = NULL
Public Function Datagram_CheckNewConnections() As Long
    Dim ret As Long

    For net_landriverlevel = 0 To net_numlandrivers - 1
        If net_landrivers(net_landriverlevel).initialized Then
            ret = DatagramCheckNewConnections()
            If ret <> 0 Then Exit For
        End If
    Next
    Datagram_CheckNewConnections = ret
End Function

' C: _Datagram_SearchForHosts
Private Sub DatagramSearchForHosts(ByVal xmit As Boolean)
    Dim ret As Long, n As Long, i As Long, l As Long
    Dim readaddr As qsockaddr, myaddr As qsockaddr
    Dim control As Long

    Lan_GetSocketAddr net_landriverlevel, net_landrivers(net_landriverlevel).controlSock, myaddr
    If xmit Then
        SZ_Clear net_message
        ' save space for the header, filled in later
        MSG_WriteLong net_message, 0
        MSG_WriteByte net_message, CCREQ_SERVER_INFO
        MSG_WriteString net_message, "QUAKE"
        MSG_WriteByte net_message, NET_PROTOCOL_VERSION
        PutBigLong net_message.data, 0, NETFLAG_CTL Or (net_message.cursize And NETFLAG_LENGTH_MASK)
        Lan_Broadcast net_landriverlevel, net_landrivers(net_landriverlevel).controlSock, net_message.data, net_message.cursize
        SZ_Clear net_message
    End If

    Do
        ret = Lan_Read(net_landriverlevel, net_landrivers(net_landriverlevel).controlSock, net_message.data, net_message.maxsize, readaddr)
        If ret <= 0 Then Exit Do

        If ret < 4 Then GoTo continue_
        net_message.cursize = ret

        ' don't answer our own query
        If Lan_AddrCompare(net_landriverlevel, readaddr, myaddr) >= 0 Then GoTo continue_

        ' is the cache full?
        If hostCacheCount = HOSTCACHESIZE Then GoTo continue_

        MSG_BeginReading
        control = GetBigLong(net_message.data, 0)
        MSG_ReadLong
        If control = -1 Then GoTo continue_
        If (control And (Not NETFLAG_LENGTH_MASK)) <> NETFLAG_CTL Then GoTo continue_
        If (control And NETFLAG_LENGTH_MASK) <> ret Then GoTo continue_

        If MSG_ReadByte() <> CCREP_SERVER_INFO Then GoTo continue_

        Lan_GetAddrFromName net_landriverlevel, MSG_ReadString(), readaddr
        ' search the cache for this server
        For n = 0 To hostCacheCount - 1
            If Lan_AddrCompare(net_landriverlevel, readaddr, hostcache(n).addr) = 0 Then Exit For
        Next

        ' is it already there?
        If n < hostCacheCount Then GoTo continue_

        ' add it
        hostCacheCount = hostCacheCount + 1
        hostcache(n).name = MSG_ReadString()
        hostcache(n).map = MSG_ReadString()
        hostcache(n).users = MSG_ReadByte()
        hostcache(n).maxusers = MSG_ReadByte()
        If MSG_ReadByte() <> NET_PROTOCOL_VERSION Then
            hostcache(n).cname = Left$(hostcache(n).name, 14)
            hostcache(n).name = "*" & hostcache(n).cname
        End If
        hostcache(n).addr = readaddr
        hostcache(n).driver = net_driverlevel
        hostcache(n).ldriver = net_landriverlevel
        hostcache(n).cname = Lan_AddrToString(net_landriverlevel, readaddr)

        ' check for a name conflict
        i = 0
        Do While i < hostCacheCount
            If i <> n Then
                If Q_strcasecmp(hostcache(n).name, hostcache(i).name) = 0 Then
                    l = Len(hostcache(n).name)
                    If l = 0 Then
                        hostcache(n).name = "0"          ' (C writes before the name)
                    ElseIf l < 15 And Right$(hostcache(n).name, 1) > "8" Then
                        hostcache(n).name = hostcache(n).name & "0"
                    Else
                        hostcache(n).name = Left$(hostcache(n).name, l - 1) & Chr$(Asc(Right$(hostcache(n).name, 1)) + 1)
                    End If
                    i = -1
                End If
            End If
            i = i + 1
        Loop
continue_:
    Loop
End Sub

Public Sub Datagram_SearchForHosts(ByVal xmit As Boolean)
    For net_landriverlevel = 0 To net_numlandrivers - 1
        If hostCacheCount = HOSTCACHESIZE Then Exit For
        If net_landrivers(net_landriverlevel).initialized Then DatagramSearchForHosts xmit
    Next
End Sub

' C: _Datagram_Connect; returns a net_qsockets() index, 0 = NULL
Private Function DatagramConnect(ByVal host As String) As Long
    Dim sendaddr As qsockaddr, readaddr As qsockaddr
    Dim sock As Long, newsock As Long, ret As Long, reps As Long
    Dim start_time As Double
    Dim control As Long
    Dim reason As String
    Dim name As String

    ' see if we can resolve the host name
    If Lan_GetAddrFromName(net_landriverlevel, host, sendaddr) = -1 Then Exit Function

    newsock = Lan_OpenSocket(net_landriverlevel, 0)
    If newsock = -1 Then Exit Function

    sock = NET_NewQSocket()
    If sock = 0 Then GoTo ErrorReturn2
    net_qsockets(sock).socket = newsock
    net_qsockets(sock).landriver = net_landriverlevel

    ' connect to the host
    If Lan_Connect(net_landriverlevel, newsock, sendaddr) = -1 Then GoTo ErrorReturn

    ' send the connection request
    Con_Printf "trying...\n": SCR_UpdateScreen
    start_time = net_time

    For reps = 0 To 2
        SZ_Clear net_message
        ' save space for the header, filled in later
        MSG_WriteLong net_message, 0
        MSG_WriteByte net_message, CCREQ_CONNECT
        MSG_WriteString net_message, "QUAKE"
        MSG_WriteByte net_message, NET_PROTOCOL_VERSION
        PutBigLong net_message.data, 0, NETFLAG_CTL Or (net_message.cursize And NETFLAG_LENGTH_MASK)
        Lan_Write net_landriverlevel, newsock, net_message.data, net_message.cursize, sendaddr
        SZ_Clear net_message
        Do
            ret = Lan_Read(net_landriverlevel, newsock, net_message.data, net_message.maxsize, readaddr)
            ' if we got something, validate it
            If ret > 0 Then
                ' is it from the right place?
                If Lan_AddrCompare(net_qsockets(sock).landriver, readaddr, sendaddr) <> 0 Then
                    ret = 0
                ElseIf ret < 4 Then
                    ret = 0
                Else
                    net_message.cursize = ret
                    MSG_BeginReading

                    control = GetBigLong(net_message.data, 0)
                    MSG_ReadLong
                    If control = -1 Then
                        ret = 0
                    ElseIf (control And (Not NETFLAG_LENGTH_MASK)) <> NETFLAG_CTL Then
                        ret = 0
                    ElseIf (control And NETFLAG_LENGTH_MASK) <> ret Then
                        ret = 0
                    End If
                End If
            End If
        Loop While ret = 0 And (SetNetTime() - start_time) < 2.5
        If ret Then Exit For
        Con_Printf "still trying...\n": SCR_UpdateScreen
        start_time = SetNetTime()
    Next

    If ret = 0 Then
        reason = "No Response"
        Con_Printf "%s\n", reason
        m_return_reason = reason
        GoTo ErrorReturn
    End If

    If ret = -1 Then
        reason = "Network Error"
        Con_Printf "%s\n", reason
        m_return_reason = reason
        GoTo ErrorReturn
    End If

    ret = MSG_ReadByte()
    If ret = CCREP_REJECT Then
        reason = MSG_ReadString()
        Con_Printf "%s", reason
        m_return_reason = Left$(reason, 31)
        GoTo ErrorReturn
    End If

    If ret = CCREP_ACCEPT Then
        net_qsockets(sock).addr = sendaddr
        Lan_SetSocketPort net_landriverlevel, net_qsockets(sock).addr, MSG_ReadLong()
    Else
        reason = "Bad Response"
        Con_Printf "%s\n", reason
        m_return_reason = reason
        GoTo ErrorReturn
    End If

    Lan_GetNameFromAddr net_landriverlevel, sendaddr, name
    net_qsockets(sock).address = name

    Con_Printf "Connection accepted\n"
    net_qsockets(sock).lastMessageTime = SetNetTime()

    ' switch the connection to the specified address
    If Lan_Connect(net_landriverlevel, newsock, net_qsockets(sock).addr) = -1 Then
        reason = "Connect to Game failed"
        Con_Printf "%s\n", reason
        m_return_reason = reason
        GoTo ErrorReturn
    End If

    m_return_onerror = False
    DatagramConnect = sock
    Exit Function

ErrorReturn:
    NET_FreeQSocket sock
ErrorReturn2:
    Lan_CloseSocket net_landriverlevel, newsock
    If m_return_onerror Then
        key_dest = key_menu
        m_state = m_return_state
        m_return_onerror = False
    End If
End Function

' returns a net_qsockets() index, 0 = NULL
Public Function Datagram_Connect(ByVal host As String) As Long
    Dim ret As Long

    For net_landriverlevel = 0 To net_numlandrivers - 1
        If net_landrivers(net_landriverlevel).initialized Then
            ret = DatagramConnect(host)
            If ret <> 0 Then Exit For
        End If
    Next
    Datagram_Connect = ret
End Function

' (QuakeVB) the console commands of this module, for Cmd_Call
Public Sub Datagram_CallCommand(ByVal fn As Long)
    Select Case fn
    Case XC_NET_Stats_f: NET_Stats_f
    Case XC_NET_Ban_f: NET_Ban_f
    Case XC_Test_f: Test_f
    Case XC_Test2_f: Test2_f
    Case Else: Sys_Error "Datagram_CallCommand: bad function %i", fn
    End Select
End Sub

'=============================================================================
' (QuakeVB) the LAN driver functions: C's dfunc.Read () etc.
' (dfunc = net_landrivers[net_landriverlevel], sfunc = net_landrivers[sock->landriver])

Private Function Lan_Init(ByVal ld As Long) As Long
    Select Case ld
    Case 0: Lan_Init = WINS_Init()
    Case Else: Sys_Error "Lan_Init: bad driver %i", ld
    End Select
End Function

Private Sub Lan_Shutdown(ByVal ld As Long)
    Select Case ld
    Case 0: WINS_Shutdown
    End Select
End Sub

Private Sub Lan_Listen(ByVal ld As Long, ByVal state As Boolean)
    Select Case ld
    Case 0: WINS_Listen state
    End Select
End Sub

Private Function Lan_OpenSocket(ByVal ld As Long, ByVal port As Long) As Long
    Select Case ld
    Case 0: Lan_OpenSocket = WINS_OpenSocket(port)
    End Select
End Function

Private Function Lan_CloseSocket(ByVal ld As Long, ByVal socket As Long) As Long
    Select Case ld
    Case 0: Lan_CloseSocket = WINS_CloseSocket(socket)
    End Select
End Function

Private Function Lan_Connect(ByVal ld As Long, ByVal socket As Long, addr As qsockaddr) As Long
    Select Case ld
    Case 0: Lan_Connect = WINS_Connect(socket, addr)
    End Select
End Function

Private Function Lan_CheckNewConnections(ByVal ld As Long) As Long
    Select Case ld
    Case 0: Lan_CheckNewConnections = WINS_CheckNewConnections()
    End Select
End Function

Private Function Lan_Read(ByVal ld As Long, ByVal socket As Long, buf() As Byte, ByVal length_ As Long, addr As qsockaddr) As Long
    Select Case ld
    Case 0: Lan_Read = WINS_Read(socket, buf, length_, addr)
    End Select
End Function

Private Function Lan_Write(ByVal ld As Long, ByVal socket As Long, buf() As Byte, ByVal length_ As Long, addr As qsockaddr) As Long
    Select Case ld
    Case 0: Lan_Write = WINS_Write(socket, buf, length_, addr)
    End Select
End Function

Private Function Lan_Broadcast(ByVal ld As Long, ByVal socket As Long, buf() As Byte, ByVal length_ As Long) As Long
    Select Case ld
    Case 0: Lan_Broadcast = WINS_Broadcast(socket, buf, length_)
    End Select
End Function

Private Function Lan_AddrToString(ByVal ld As Long, addr As qsockaddr) As String
    Select Case ld
    Case 0: Lan_AddrToString = WINS_AddrToString(addr)
    End Select
End Function

Private Function Lan_GetSocketAddr(ByVal ld As Long, ByVal socket As Long, addr As qsockaddr) As Long
    Select Case ld
    Case 0: Lan_GetSocketAddr = WINS_GetSocketAddr(socket, addr)
    End Select
End Function

Private Function Lan_GetNameFromAddr(ByVal ld As Long, addr As qsockaddr, name As String) As Long
    Select Case ld
    Case 0: Lan_GetNameFromAddr = WINS_GetNameFromAddr(addr, name)
    End Select
End Function

Private Function Lan_GetAddrFromName(ByVal ld As Long, ByVal name As String, addr As qsockaddr) As Long
    Select Case ld
    Case 0: Lan_GetAddrFromName = WINS_GetAddrFromName(name, addr)
    End Select
End Function

Private Function Lan_AddrCompare(ByVal ld As Long, addr1 As qsockaddr, addr2 As qsockaddr) As Long
    Select Case ld
    Case 0: Lan_AddrCompare = WINS_AddrCompare(addr1, addr2)
    End Select
End Function

Private Function Lan_GetSocketPort(ByVal ld As Long, addr As qsockaddr) As Long
    Select Case ld
    Case 0: Lan_GetSocketPort = WINS_GetSocketPort(addr)
    End Select
End Function

Private Function Lan_SetSocketPort(ByVal ld As Long, addr As qsockaddr, ByVal port As Long) As Long
    Select Case ld
    Case 0: Lan_SetSocketPort = WINS_SetSocketPort(addr, port)
    End Select
End Function
