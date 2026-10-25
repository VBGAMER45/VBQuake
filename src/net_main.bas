Attribute VB_Name = "net_main"
' net_main.bas - the network layer above the drivers (net_main.c)
'
' C calls the drivers through net_drivers[], a table of function pointers;
' here the Drv_ functions dispatch on the driver number: 0 is the loopback
' (net_loop.bas), 1 the datagram driver (net_dgrm.bas). The driver table
' itself is net_win.c's; NET_Init fills it in.
'
' Not ported: net_vcr (-record / -playback), and the serial and modem
' configuration (WinQuake has no serial driver).
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private listening As Boolean

Private slistStartTime As Double
Private slistLastShown As Long

Private cv_net_messagetimeout As Long

Private configRestored As Boolean
Private cv_config_com_port As Long
Private cv_config_com_irq As Long
Private cv_config_com_baud As Long
Private cv_config_com_modem As Long
Private cv_config_modem_dialtype As Long
Private cv_config_modem_clear As Long
Private cv_config_modem_init As Long
Private cv_config_modem_hangup As Long

' the poll procedures (C: PollProcedure structs in a linked list)
Private pollProcedureList As Long        ' PP_ number, 0 = NULL
Private poll_next(1 To MAX_POLL_PROCEDURES) As Long
Private poll_time(1 To MAX_POLL_PROCEDURES) As Double

Public Function SetNetTime() As Double
    net_time = Sys_FloatTime()
    SetNetTime = net_time
End Function

' Called by drivers when a new communications endpoint is required
' The sequence and buffer fields will be filled in properly
Public Function NET_NewQSocket() As Long
    Dim sock As Long

    If net_freeSockets = 0 Then Exit Function

    If net_activeconnections >= svs.maxclients Then Exit Function

    ' get one from free list
    sock = net_freeSockets
    net_freeSockets = net_qsockets(sock).nxt

    ' add it to active list
    net_qsockets(sock).nxt = net_activeSockets
    net_activeSockets = sock

    With net_qsockets(sock)
        .disconnected = False
        .connecttime = net_time
        .address = "UNSET ADDRESS"
        .driver = net_driverlevel
        .socket = 0
        .driverdata = 0
        .canSend = True
        .sendNext = False
        .lastMessageTime = net_time
        .ackSequence = 0
        .sendSequence = 0
        .unreliableSendSequence = 0
        .sendMessageLength = 0
        .receiveSequence = 0
        .unreliableReceiveSequence = 0
        .receiveMessageLength = 0
    End With

    NET_NewQSocket = sock
End Function

Public Sub NET_FreeQSocket(ByVal sock As Long)
    Dim s As Long

    ' remove it from active list
    If sock = net_activeSockets Then
        net_activeSockets = net_qsockets(net_activeSockets).nxt
    Else
        s = net_activeSockets
        Do While s <> 0
            If net_qsockets(s).nxt = sock Then
                net_qsockets(s).nxt = net_qsockets(sock).nxt
                Exit Do
            End If
            s = net_qsockets(s).nxt
        Loop
        If s = 0 Then Sys_Error "NET_FreeQSocket: not active\n"
    End If

    ' add it to free list
    net_qsockets(sock).nxt = net_freeSockets
    net_freeSockets = sock
    net_qsockets(sock).disconnected = True
End Sub

Private Sub NET_Listen_f()
    If Cmd_Argc() <> 2 Then
        If listening Then
            Con_Printf """listen"" is ""%u""\n", 1
        Else
            Con_Printf """listen"" is ""%u""\n", 0
        End If
        Exit Sub
    End If

    listening = Q_atoi(Cmd_Argv(1)) <> 0

    For net_driverlevel = 0 To net_numdrivers - 1
        If net_drivers(net_driverlevel).initialized Then Drv_Listen net_driverlevel, listening
    Next
End Sub

Private Sub MaxPlayers_f()
    Dim n As Long

    If Cmd_Argc() <> 2 Then
        Con_Printf """maxplayers"" is ""%u""\n", svs.maxclients
        Exit Sub
    End If

    If sv.active Then
        Con_Printf "maxplayers can not be changed while a server is running.\n"
        Exit Sub
    End If

    n = Q_atoi(Cmd_Argv(1))
    If n < 1 Then n = 1
    If n > svs.maxclientslimit Then
        n = svs.maxclientslimit
        Con_Printf """maxplayers"" set to ""%u""\n", n
    End If

    If n = 1 And listening Then Cbuf_AddText "listen 0" & vbLf

    If n > 1 And Not listening Then Cbuf_AddText "listen 1" & vbLf

    svs.maxclients = n
    If n = 1 Then
        Cvar_Set "deathmatch", "0"
    Else
        Cvar_Set "deathmatch", "1"
    End If
End Sub

Private Sub NET_Port_f()
    Dim n As Long

    If Cmd_Argc() <> 2 Then
        Con_Printf """port"" is ""%u""\n", net_hostport
        Exit Sub
    End If

    n = Q_atoi(Cmd_Argv(1))
    If n < 1 Or n > 65534 Then
        Con_Printf "Bad value, must be between 1 and 65534\n"
        Exit Sub
    End If

    DEFAULTnet_hostport = n
    net_hostport = n

    If listening Then
        ' force a change to the new port
        Cbuf_AddText "listen 0" & vbLf
        Cbuf_AddText "listen 1" & vbLf
    End If
End Sub

Private Sub PrintSlistHeader()
    Con_Printf "Server          Map             Users\n"
    Con_Printf "--------------- --------------- -----\n"
    slistLastShown = 0
End Sub

Private Sub PrintSlist()
    Dim n As Long

    For n = slistLastShown To hostCacheCount - 1
        If hostcache(n).maxusers Then
            Con_Printf "%-15.15s %-15.15s %2u/%2u\n", hostcache(n).name, hostcache(n).map, hostcache(n).users, hostcache(n).maxusers
        Else
            Con_Printf "%-15.15s %-15.15s\n", hostcache(n).name, hostcache(n).map
        End If
    Next
    slistLastShown = n
End Sub

Private Sub PrintSlistTrailer()
    If hostCacheCount Then
        Con_Printf "== end list ==\n\n"
    Else
        Con_Printf "No Quake servers found.\n\n"
    End If
End Sub

Public Sub NET_Slist_f()
    If slistInProgress Then Exit Sub

    If Not slistSilent Then
        Con_Printf "Looking for Quake servers...\n"
        PrintSlistHeader
    End If

    slistInProgress = True
    slistStartTime = Sys_FloatTime()

    SchedulePollProcedure PP_SLIST_SEND, 0#
    SchedulePollProcedure PP_SLIST_POLL, 0.1

    hostCacheCount = 0
End Sub

Private Sub Slist_Send()
    For net_driverlevel = 0 To net_numdrivers - 1
        If Not slistLocal And net_driverlevel = 0 Then GoTo nextdrv
        If Not net_drivers(net_driverlevel).initialized Then GoTo nextdrv
        Drv_SearchForHosts net_driverlevel, True
nextdrv:
    Next

    If (Sys_FloatTime() - slistStartTime) < 0.5 Then SchedulePollProcedure PP_SLIST_SEND, 0.75
End Sub

Private Sub Slist_Poll()
    For net_driverlevel = 0 To net_numdrivers - 1
        If Not slistLocal And net_driverlevel = 0 Then GoTo nextdrv
        If Not net_drivers(net_driverlevel).initialized Then GoTo nextdrv
        Drv_SearchForHosts net_driverlevel, False
nextdrv:
    Next

    If Not slistSilent Then PrintSlist

    If (Sys_FloatTime() - slistStartTime) < 1.5 Then
        SchedulePollProcedure PP_SLIST_POLL, 0.1
        Exit Sub
    End If

    If Not slistSilent Then PrintSlistTrailer
    slistInProgress = False
    slistSilent = False
    slistLocal = True
End Sub

' host "" stands for NULL; returns a net_qsockets() index, 0 if not able to
Public Function NET_Connect(ByVal host As String) As Long
    Dim ret As Long, n As Long, numdrivers As Long

    numdrivers = net_numdrivers

    SetNetTime

    If Len(host) Then
        If Q_strcasecmp(host, "local") = 0 Then
            numdrivers = 1
            GoTo JustDoIt
        End If

        If hostCacheCount Then
            For n = 0 To hostCacheCount - 1
                If Q_strcasecmp(host, hostcache(n).name) = 0 Then
                    host = hostcache(n).cname
                    Exit For
                End If
            Next
            If n < hostCacheCount Then GoTo JustDoIt
        End If
    End If

    slistSilent = Len(host) > 0
    NET_Slist_f

    Do While slistInProgress
        NET_Poll
    Loop

    If Len(host) = 0 Then
        If hostCacheCount <> 1 Then Exit Function
        host = hostcache(0).cname
        Con_Printf "Connecting to...\n%s @ %s\n\n", hostcache(0).name, host
    End If

    If hostCacheCount Then
        For n = 0 To hostCacheCount - 1
            If Q_strcasecmp(host, hostcache(n).name) = 0 Then
                host = hostcache(n).cname
                Exit For
            End If
        Next
    End If

JustDoIt:
    For net_driverlevel = 0 To numdrivers - 1
        If net_drivers(net_driverlevel).initialized Then
            ret = Drv_Connect(net_driverlevel, host)
            If ret Then
                NET_Connect = ret
                Exit Function
            End If
        End If
    Next

    If Len(host) Then
        Con_Printf "\n"
        PrintSlistHeader
        PrintSlist
        PrintSlistTrailer
    End If
End Function

' returns a new connection (net_qsockets() index) if there is one pending, else 0
Public Function NET_CheckNewConnections() As Long
    Dim ret As Long

    SetNetTime

    For net_driverlevel = 0 To net_numdrivers - 1
        If Not net_drivers(net_driverlevel).initialized Then GoTo nextdrv
        If net_driverlevel <> 0 And Not listening Then GoTo nextdrv
        ret = Drv_CheckNewConnections(net_driverlevel)
        If ret Then
            NET_CheckNewConnections = ret
            Exit Function
        End If
nextdrv:
    Next
End Function

Public Sub NET_Close(ByVal sock As Long)
    If sock = 0 Then Exit Sub

    If net_qsockets(sock).disconnected Then Exit Sub

    SetNetTime

    ' call the driver_Close function
    Drv_Close sock

    NET_FreeQSocket sock
End Sub

' If there is a complete message, return it in net_message
'
' returns 0 if no data is waiting
' returns 1 if a message was received
' returns -1 if connection is invalid
Public Function NET_GetMessage(ByVal sock As Long) As Long
    Dim ret As Long

    If sock = 0 Then
        NET_GetMessage = -1
        Exit Function
    End If

    If net_qsockets(sock).disconnected Then
        Con_Printf "NET_GetMessage: disconnected socket\n"
        NET_GetMessage = -1
        Exit Function
    End If

    SetNetTime

    ret = Drv_GetMessage(sock)

    ' see if this connection has timed out
    If ret = 0 And net_qsockets(sock).driver <> 0 Then
        If net_time - net_qsockets(sock).lastMessageTime > cvars(cv_net_messagetimeout).value Then
            NET_Close sock
            NET_GetMessage = -1
            Exit Function
        End If
    End If

    If ret > 0 Then
        If net_qsockets(sock).driver <> 0 Then
            net_qsockets(sock).lastMessageTime = net_time
            If ret = 1 Then
                messagesReceived = messagesReceived + 1
            ElseIf ret = 2 Then
                unreliableMessagesReceived = unreliableMessagesReceived + 1
            End If
        End If
    End If

    NET_GetMessage = ret
End Function

' Try to send a complete length+message unit over the reliable stream.
' returns 0 if the message cannot be delivered reliably, but the connection
'         is still considered valid
' returns 1 if the message was sent properly
' returns -1 if the connection died
Public Function NET_SendMessage(ByVal sock As Long, data As sizebuf_t) As Long
    Dim r As Long

    If sock = 0 Then
        NET_SendMessage = -1
        Exit Function
    End If

    If net_qsockets(sock).disconnected Then
        Con_Printf "NET_SendMessage: disconnected socket\n"
        NET_SendMessage = -1
        Exit Function
    End If

    SetNetTime
    r = Drv_SendMessage(sock, data)
    If r = 1 And net_qsockets(sock).driver <> 0 Then messagesSent = messagesSent + 1

    NET_SendMessage = r
End Function

Public Function NET_SendUnreliableMessage(ByVal sock As Long, data As sizebuf_t) As Long
    Dim r As Long

    If sock = 0 Then
        NET_SendUnreliableMessage = -1
        Exit Function
    End If

    If net_qsockets(sock).disconnected Then
        Con_Printf "NET_SendMessage: disconnected socket\n"
        NET_SendUnreliableMessage = -1
        Exit Function
    End If

    SetNetTime
    r = Drv_SendUnreliableMessage(sock, data)
    If r = 1 And net_qsockets(sock).driver <> 0 Then unreliableMessagesSent = unreliableMessagesSent + 1

    NET_SendUnreliableMessage = r
End Function

' Returns true or false if the given qsocket can currently accept a
' message to be transmitted.
Public Function NET_CanSendMessage(ByVal sock As Long) As Boolean
    If sock = 0 Then Exit Function

    If net_qsockets(sock).disconnected Then Exit Function

    SetNetTime

    NET_CanSendMessage = Drv_CanSendMessage(sock)
End Function

' This is a reliable *blocking* send to all attached clients.
Public Function NET_SendToAll(data As sizebuf_t, ByVal blocktime As Long) As Long
    Dim start As Double, i As Long, count As Long
    Dim state1(0 To MAX_SCOREBOARD - 1) As Boolean, state2(0 To MAX_SCOREBOARD - 1) As Boolean

    For i = 0 To svs.maxclients - 1
        host_client = i
        If svs.clients(i).netconnection = 0 Then
            ' (QuakeVB) C leaves state1 and state2 unset here (whatever is
            ' on the stack); a free slot counts as done, so the loop below
            ' does not wait the whole blocktime for it
            state1(i) = True
            state2(i) = True
            GoTo nexti
        End If
        If svs.clients(i).active Then
            If net_qsockets(svs.clients(i).netconnection).driver = 0 Then
                NET_SendMessage svs.clients(i).netconnection, data
                state1(i) = True
                state2(i) = True
                GoTo nexti
            End If
            count = count + 1
            state1(i) = False
            state2(i) = False
        Else
            state1(i) = True
            state2(i) = True
        End If
nexti:
    Next

    start = Sys_FloatTime()
    Do While count
        count = 0
        For i = 0 To svs.maxclients - 1
            host_client = i
            If Not state1(i) Then
                If NET_CanSendMessage(svs.clients(i).netconnection) Then
                    state1(i) = True
                    NET_SendMessage svs.clients(i).netconnection, data
                Else
                    NET_GetMessage svs.clients(i).netconnection
                End If
                count = count + 1
                GoTo nextj
            End If

            If Not state2(i) Then
                If NET_CanSendMessage(svs.clients(i).netconnection) Then
                    state2(i) = True
                Else
                    NET_GetMessage svs.clients(i).netconnection
                End If
                count = count + 1
                GoTo nextj
            End If
nextj:
        Next
        If (Sys_FloatTime() - start) > blocktime Then Exit Do
    Loop
    NET_SendToAll = count
End Function

'=============================================================================

Public Sub NET_Init()
    Dim i As Long, controlSocket As Long

    ' (net_win.c's driver table)
    net_drivers(0).name = "Loopback"
    net_drivers(1).name = "Datagram"
    net_numdrivers = 2
    DEFAULTnet_hostport = 26000
    slistLocal = True

    i = COM_CheckParm("-port")
    If i = 0 Then i = COM_CheckParm("-udpport")
    If i = 0 Then i = COM_CheckParm("-ipxport")

    If i Then
        If i < com_argc - 1 Then
            DEFAULTnet_hostport = Q_atoi(com_argv(i + 1))
        Else
            Sys_Error "NET_Init: you must specify a number after -port"
        End If
    End If
    net_hostport = DEFAULTnet_hostport

    If COM_CheckParm("-listen") <> 0 Or cls.state = ca_dedicated Then listening = True
    net_numsockets = svs.maxclientslimit
    If cls.state <> ca_dedicated Then net_numsockets = net_numsockets + 1

    SetNetTime

    ReDim net_qsockets(1 To net_numsockets)
    For i = 1 To net_numsockets
        net_qsockets(i).nxt = net_freeSockets
        net_freeSockets = i
        net_qsockets(i).disconnected = True
    Next

    ' allocate space for network message buffer
    SZ_Alloc net_message, NET_MAXMESSAGE

    cv_net_messagetimeout = Cvar_RegisterVariable("net_messagetimeout", "300")
    cv_hostname = Cvar_RegisterVariable("hostname", "UNNAMED")
    cv_config_com_port = Cvar_RegisterVariable("_config_com_port", "0x3f8", True)
    cv_config_com_irq = Cvar_RegisterVariable("_config_com_irq", "4", True)
    cv_config_com_baud = Cvar_RegisterVariable("_config_com_baud", "57600", True)
    cv_config_com_modem = Cvar_RegisterVariable("_config_com_modem", "1", True)
    cv_config_modem_dialtype = Cvar_RegisterVariable("_config_modem_dialtype", "T", True)
    cv_config_modem_clear = Cvar_RegisterVariable("_config_modem_clear", "ATZ", True)
    cv_config_modem_init = Cvar_RegisterVariable("_config_modem_init", "", True)
    cv_config_modem_hangup = Cvar_RegisterVariable("_config_modem_hangup", "AT H", True)

    Cmd_AddCommand "slist", XC_NET_Slist_f
    Cmd_AddCommand "listen", XC_NET_Listen_f
    Cmd_AddCommand "maxplayers", XC_MaxPlayers_f
    Cmd_AddCommand "port", XC_NET_Port_f

    ' initialize all the drivers
    For net_driverlevel = 0 To net_numdrivers - 1
        controlSocket = Drv_Init(net_driverlevel)
        If controlSocket <> -1 Then
            net_drivers(net_driverlevel).initialized = True
            net_drivers(net_driverlevel).controlSock = controlSocket
            If listening Then Drv_Listen net_driverlevel, True
        End If
    Next

    If Len(my_ipx_address) Then Con_DPrintf "IPX address %s\n", my_ipx_address
    If Len(my_tcpip_address) Then Con_DPrintf "TCP/IP address %s\n", my_tcpip_address
End Sub

Public Sub NET_Shutdown()
    Dim sock As Long, nxt As Long

    SetNetTime

    ' (QuakeVB: the next socket is read before NET_Close puts this one on the
    ' free list. C reads sock->next afterwards and so walks on into the free
    ' list: it closes only the first active socket.)
    sock = net_activeSockets
    Do While sock <> 0
        nxt = net_qsockets(sock).nxt
        NET_Close sock
        sock = nxt
    Loop

    '
    ' shutdown the drivers
    '
    For net_driverlevel = 0 To net_numdrivers - 1
        If net_drivers(net_driverlevel).initialized Then
            Drv_Shutdown net_driverlevel
            net_drivers(net_driverlevel).initialized = False
        End If
    Next
End Sub

Public Sub NET_Poll()
    Dim pp As Long

    If Not configRestored Then
        ' (serialAvailable is never set: WinQuake has no serial driver)
        configRestored = True
    End If

    SetNetTime

    pp = pollProcedureList
    Do While pp <> 0
        If poll_time(pp) > net_time Then Exit Do
        pollProcedureList = poll_next(pp)
        RunPollProcedure pp
        pp = poll_next(pp)
    Loop
End Sub

Public Sub SchedulePollProcedure(ByVal proc As Long, ByVal timeOffset As Double)
    Dim pp As Long, prev As Long

    poll_time(proc) = Sys_FloatTime() + timeOffset
    pp = pollProcedureList
    prev = 0
    Do While pp <> 0
        If poll_time(pp) >= poll_time(proc) Then Exit Do
        prev = pp
        pp = poll_next(pp)
    Loop

    If prev = 0 Then
        poll_next(proc) = pollProcedureList
        pollProcedureList = proc
        Exit Sub
    End If

    poll_next(proc) = pp
    poll_next(prev) = proc
End Sub

' (QuakeVB) C: pp->procedure (pp->arg)
Private Sub RunPollProcedure(ByVal pp As Long)
    Select Case pp
    Case PP_SLIST_SEND: Slist_Send
    Case PP_SLIST_POLL: Slist_Poll
    Case PP_TEST_POLL: Test_Poll
    Case PP_TEST2_POLL: Test2_Poll
    Case Else: Sys_Error "RunPollProcedure: bad procedure %i", pp
    End Select
End Sub

' (QuakeVB) the console commands of this module, for Cmd_Call
Public Sub NET_CallCommand(ByVal fn As Long)
    Select Case fn
    Case XC_NET_Slist_f: NET_Slist_f
    Case XC_NET_Listen_f: NET_Listen_f
    Case XC_MaxPlayers_f: MaxPlayers_f
    Case XC_NET_Port_f: NET_Port_f
    Case Else: Sys_Error "NET_CallCommand: bad function %i", fn
    End Select
End Sub

'=============================================================================
' (QuakeVB) the driver functions: C's net_drivers[d].Init () etc.
' (sfunc = net_drivers[sock->driver], dfunc = net_drivers[net_driverlevel])

Private Function Drv_Init(ByVal d As Long) As Long
    Select Case d
    Case 0: Drv_Init = Loop_Init()
    Case 1: Drv_Init = Datagram_Init()
    Case Else: Sys_Error "Drv_Init: bad driver %i", d
    End Select
End Function

Private Sub Drv_Listen(ByVal d As Long, ByVal state As Boolean)
    Select Case d
    Case 0: Loop_Listen state
    Case 1: Datagram_Listen state
    End Select
End Sub

Private Sub Drv_SearchForHosts(ByVal d As Long, ByVal xmit As Boolean)
    Select Case d
    Case 0: Loop_SearchForHosts xmit
    Case 1: Datagram_SearchForHosts xmit
    End Select
End Sub

Private Function Drv_Connect(ByVal d As Long, ByVal host As String) As Long
    Select Case d
    Case 0: Drv_Connect = Loop_Connect(host)
    Case 1: Drv_Connect = Datagram_Connect(host)
    End Select
End Function

Private Function Drv_CheckNewConnections(ByVal d As Long) As Long
    Select Case d
    Case 0: Drv_CheckNewConnections = Loop_CheckNewConnections()
    Case 1: Drv_CheckNewConnections = Datagram_CheckNewConnections()
    End Select
End Function

Private Function Drv_GetMessage(ByVal sock As Long) As Long
    Select Case net_qsockets(sock).driver
    Case 0: Drv_GetMessage = Loop_GetMessage(sock)
    Case 1: Drv_GetMessage = Datagram_GetMessage(sock)
    End Select
End Function

Private Function Drv_SendMessage(ByVal sock As Long, data As sizebuf_t) As Long
    Select Case net_qsockets(sock).driver
    Case 0: Drv_SendMessage = Loop_SendMessage(sock, data)
    Case 1: Drv_SendMessage = Datagram_SendMessage(sock, data)
    End Select
End Function

Private Function Drv_SendUnreliableMessage(ByVal sock As Long, data As sizebuf_t) As Long
    Select Case net_qsockets(sock).driver
    Case 0: Drv_SendUnreliableMessage = Loop_SendUnreliableMessage(sock, data)
    Case 1: Drv_SendUnreliableMessage = Datagram_SendUnreliableMessage(sock, data)
    End Select
End Function

Private Function Drv_CanSendMessage(ByVal sock As Long) As Boolean
    Select Case net_qsockets(sock).driver
    Case 0: Drv_CanSendMessage = Loop_CanSendMessage(sock)
    Case 1: Drv_CanSendMessage = Datagram_CanSendMessage(sock)
    End Select
End Function

Private Sub Drv_Close(ByVal sock As Long)
    Select Case net_qsockets(sock).driver
    Case 0: Loop_Close sock
    Case 1: Datagram_Close sock
    End Select
End Sub

Private Sub Drv_Shutdown(ByVal d As Long)
    Select Case d
    Case 0: Loop_Shutdown
    Case 1: Datagram_Shutdown
    End Select
End Sub
