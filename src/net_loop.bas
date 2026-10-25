Attribute VB_Name = "net_loop"
' net_loop.bas - the loopback driver: a local client talking to the local
' server (net_loop.c)
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private localconnectpending As Boolean
Private loop_client As Long              ' net_qsockets() index, 0 = NULL
Private loop_server As Long

Public Function Loop_Init() As Long
    If cls.state = ca_dedicated Then
        Loop_Init = -1
        Exit Function
    End If
    Loop_Init = 0
End Function

Public Sub Loop_Shutdown()
End Sub

Public Sub Loop_Listen(ByVal state As Boolean)
End Sub

Public Sub Loop_SearchForHosts(ByVal xmit As Boolean)
    If Not sv.active Then Exit Sub

    hostCacheCount = 1
    If cvars(cv_hostname).strval = "UNNAMED" Then
        hostcache(0).name = "local"
    Else
        hostcache(0).name = cvars(cv_hostname).strval
    End If
    hostcache(0).map = sv.name
    hostcache(0).users = net_activeconnections
    hostcache(0).maxusers = svs.maxclients
    hostcache(0).driver = net_driverlevel
    hostcache(0).cname = "local"
End Sub

Public Function Loop_Connect(ByVal host As String) As Long
    If host <> "local" Then Exit Function

    localconnectpending = True

    If loop_client = 0 Then
        loop_client = NET_NewQSocket()
        If loop_client = 0 Then
            Con_Printf "Loop_Connect: no qsocket available\n"
            Exit Function
        End If
        net_qsockets(loop_client).address = "localhost"
    End If
    net_qsockets(loop_client).receiveMessageLength = 0
    net_qsockets(loop_client).sendMessageLength = 0
    net_qsockets(loop_client).canSend = True

    If loop_server = 0 Then
        loop_server = NET_NewQSocket()
        If loop_server = 0 Then
            Con_Printf "Loop_Connect: no qsocket available\n"
            Exit Function
        End If
        net_qsockets(loop_server).address = "LOCAL"
    End If
    net_qsockets(loop_server).receiveMessageLength = 0
    net_qsockets(loop_server).sendMessageLength = 0
    net_qsockets(loop_server).canSend = True

    net_qsockets(loop_client).driverdata = loop_server
    net_qsockets(loop_server).driverdata = loop_client

    Loop_Connect = loop_client
End Function

Public Function Loop_CheckNewConnections() As Long
    If Not localconnectpending Then Exit Function

    localconnectpending = False
    net_qsockets(loop_server).sendMessageLength = 0
    net_qsockets(loop_server).receiveMessageLength = 0
    net_qsockets(loop_server).canSend = True
    net_qsockets(loop_client).sendMessageLength = 0
    net_qsockets(loop_client).receiveMessageLength = 0
    net_qsockets(loop_client).canSend = True
    Loop_CheckNewConnections = loop_server
End Function

Private Function IntAlign(ByVal value As Long) As Long
    IntAlign = (value + 3) And Not 3
End Function

Public Function Loop_GetMessage(ByVal sock As Long) As Long
    Dim ret As Long, length_ As Long, p As Long

    With net_qsockets(sock)
        If .receiveMessageLength = 0 Then Exit Function

        ret = .receiveMessage(0)
        length_ = .receiveMessage(1) + .receiveMessage(2) * 256&
        ' alignment byte skipped here
        SZ_Clear net_message
        p = SZ_GetSpace(net_message, length_)
        If length_ > 0 Then CopyMemory net_message.data(p), .receiveMessage(4), length_

        length_ = IntAlign(length_ + 4)
        .receiveMessageLength = .receiveMessageLength - length_

        If .receiveMessageLength Then CopyMemory .receiveMessage(0), .receiveMessage(length_), .receiveMessageLength

        If .driverdata <> 0 And ret = 1 Then net_qsockets(.driverdata).canSend = True
    End With

    Loop_GetMessage = ret
End Function

Public Function Loop_SendMessage(ByVal sock As Long, data As sizebuf_t) As Long
    Dim other As Long, buffer As Long

    other = net_qsockets(sock).driverdata
    If other = 0 Then
        Loop_SendMessage = -1
        Exit Function
    End If

    With net_qsockets(other)
        If (.receiveMessageLength + data.cursize + 4) > NET_MAXMESSAGE Then Sys_Error "Loop_SendMessage: overflow\n"

        buffer = .receiveMessageLength

        ' message type
        .receiveMessage(buffer) = 1

        ' length
        .receiveMessage(buffer + 1) = data.cursize And &HFF&
        .receiveMessage(buffer + 2) = data.cursize \ 256

        ' align
        ' message
        If data.cursize > 0 Then CopyMemory .receiveMessage(buffer + 4), data.data(0), data.cursize
        .receiveMessageLength = IntAlign(.receiveMessageLength + data.cursize + 4)
    End With

    net_qsockets(sock).canSend = False
    Loop_SendMessage = 1
End Function

Public Function Loop_SendUnreliableMessage(ByVal sock As Long, data As sizebuf_t) As Long
    Dim other As Long, buffer As Long

    other = net_qsockets(sock).driverdata
    If other = 0 Then
        Loop_SendUnreliableMessage = -1
        Exit Function
    End If

    With net_qsockets(other)
        ' (C checks for a 3 byte header, sizeof(byte) + sizeof(short), and
        ' writes 4: the last unreliable message may overrun the buffer by one
        ' byte there, and here hit the array bounds)
        If (.receiveMessageLength + data.cursize + 3) > NET_MAXMESSAGE Then Exit Function

        buffer = .receiveMessageLength

        ' message type
        .receiveMessage(buffer) = 2

        ' length
        .receiveMessage(buffer + 1) = data.cursize And &HFF&
        .receiveMessage(buffer + 2) = data.cursize \ 256

        ' align
        ' message
        If data.cursize > 0 Then CopyMemory .receiveMessage(buffer + 4), data.data(0), data.cursize
        .receiveMessageLength = IntAlign(.receiveMessageLength + data.cursize + 4)
    End With

    Loop_SendUnreliableMessage = 1
End Function

Public Function Loop_CanSendMessage(ByVal sock As Long) As Boolean
    If net_qsockets(sock).driverdata = 0 Then Exit Function
    Loop_CanSendMessage = net_qsockets(sock).canSend
End Function

Public Function Loop_CanSendUnreliableMessage(ByVal sock As Long) As Boolean
    Loop_CanSendUnreliableMessage = True
End Function

Public Sub Loop_Close(ByVal sock As Long)
    If net_qsockets(sock).driverdata Then net_qsockets(net_qsockets(sock).driverdata).driverdata = 0
    net_qsockets(sock).receiveMessageLength = 0
    net_qsockets(sock).sendMessageLength = 0
    net_qsockets(sock).canSend = True
    If sock = loop_client Then
        loop_client = 0
    Else
        loop_server = 0
    End If
End Sub
