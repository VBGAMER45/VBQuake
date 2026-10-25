Attribute VB_Name = "host"
' host.bas - coordinates spawning and killing of local servers (host.c).
'
' A server can allways be started, even if the system started out as a client
' to a remote system.
'
' A client can NOT be started if the system started as a dedicated server.
'
' Memory is cleared / released when a server or client begins, not when they end.
'
' Host_Error and Host_EndGame unwind to _Host_Frame by raising the VB error
' HOST_ABORTSERVER (C: longjmp (host_abortserver, 1)), so no procedure on
' the way may trap errors (On Error) around calls that can reach them.
'
' Host_InitVCR comes with net_vcr (cluster G).
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public host_time As Double
Public oldrealtime As Double         ' last frame run

Public host_frameran As Boolean      ' (QuakeVB) the last Host_Frame ran a frame

Public cv_host_framerate As Long     ' set for slow motion
Public cv_host_speeds As Long        ' set for running times
Public cv_serverprofile As Long
Public cv_fraglimit As Long
Public cv_timelimit As Long
Public cv_teamplay As Long
Public cv_samelevel As Long
Public cv_noexit As Long
Public cv_skill As Long              ' 0 - 3
Public cv_deathmatch As Long         ' 0, 1, or 2
Public cv_coop As Long               ' 0 or 1
Public cv_pausable As Long
Public cv_temp1 As Long

' C statics
Private hosterror_inerror As Boolean
Private time1 As Double
Private time2 As Double
Private time3 As Double
Private timetotal As Double
Private timecount As Long
Private shutdown_isdown As Boolean

Public Sub Host_EndGame(ByVal message As String, ParamArray args() As Variant)
    Dim a As Variant, s As String
    a = args
    s = vsprintf(message, a)
    Con_DPrintf "Host_EndGame: %s\n", s

    If sv.active Then Host_ShutdownServer False

    If cls.state = ca_dedicated Then Sys_Error "Host_EndGame: %s\n", s   ' dedicated servers exit

    If cls.demonum <> -1 Then
        CL_NextDemo
    Else
        CL_Disconnect
    End If

    Err.Raise HOST_ABORTSERVER
End Sub

' This shuts down both the client and server
Public Sub Host_Error(ByVal error_ As String, ParamArray args() As Variant)
    Dim a As Variant, s As String

    If hosterror_inerror Then Sys_Error "Host_Error: recursively entered"
    hosterror_inerror = True

    SCR_EndLoadingPlaque                 ' reenable screen updates

    a = args
    s = vsprintf(error_, a)
    Con_Printf "Host_Error: %s\n", s

    If sv.active Then Host_ShutdownServer False

    If cls.state = ca_dedicated Then Sys_Error "Host_Error: %s\n", s   ' dedicated servers exit

    CL_Disconnect
    cls.demonum = -1

    hosterror_inerror = False

    Err.Raise HOST_ABORTSERVER
End Sub

Private Sub Host_FindMaxClients()
    Dim i As Long

    svs.maxclients = 1

    i = COM_CheckParm("-dedicated")
    If i Then
        cls.state = ca_dedicated
        If i <> (com_argc - 1) Then
            svs.maxclients = Q_atoi(com_argv(i + 1))
        Else
            svs.maxclients = 8
        End If
    Else
        cls.state = ca_disconnected
    End If

    i = COM_CheckParm("-listen")
    If i Then
        If cls.state = ca_dedicated Then Sys_Error "Only one of -dedicated or -listen can be specified"
        If i <> (com_argc - 1) Then
            svs.maxclients = Q_atoi(com_argv(i + 1))
        Else
            svs.maxclients = 8
        End If
    End If
    If svs.maxclients < 1 Then
        svs.maxclients = 8
    ElseIf svs.maxclients > MAX_SCOREBOARD Then
        svs.maxclients = MAX_SCOREBOARD
    End If

    svs.maxclientslimit = svs.maxclients
    If svs.maxclientslimit < 4 Then svs.maxclientslimit = 4
    ReDim svs.clients(0 To svs.maxclientslimit - 1)

    If svs.maxclients > 1 Then
        Cvar_SetValue "deathmatch", 1#
    Else
        Cvar_SetValue "deathmatch", 0#
    End If
End Sub

Private Sub Host_InitLocal()
    Host_InitCommands

    cv_host_framerate = Cvar_RegisterVariable("host_framerate", "0")
    cv_host_speeds = Cvar_RegisterVariable("host_speeds", "0")

    cv_sys_ticrate = Cvar_RegisterVariable("sys_ticrate", "0.05")
    cv_serverprofile = Cvar_RegisterVariable("serverprofile", "0")

    cv_fraglimit = Cvar_RegisterVariable("fraglimit", "0", False, True)
    cv_timelimit = Cvar_RegisterVariable("timelimit", "0", False, True)
    cv_teamplay = Cvar_RegisterVariable("teamplay", "0", False, True)
    cv_samelevel = Cvar_RegisterVariable("samelevel", "0")
    cv_noexit = Cvar_RegisterVariable("noexit", "0", False, True)
    cv_skill = Cvar_RegisterVariable("skill", "1")
    cv_developer = Cvar_RegisterVariable("developer", "0")
    cv_deathmatch = Cvar_RegisterVariable("deathmatch", "0")
    cv_coop = Cvar_RegisterVariable("coop", "0")

    cv_pausable = Cvar_RegisterVariable("pausable", "1")

    cv_temp1 = Cvar_RegisterVariable("temp1", "0")

    Host_FindMaxClients

    host_time = 1#                       ' so a think at time 0 won't get called
End Sub

' Writes key bindings and archived cvars to config.cfg
Public Sub Host_WriteConfiguration()
    Dim f As Integer, ok As Boolean

    ' dedicated servers initialize the host but don't parse and set the
    ' config.cfg cvars
    If host_initialized And Not isDedicated Then
        f = FreeFile
        ok = OpenForOutput(com_gamedir & "/config.cfg", f)
        If Not ok Then
            Con_Printf "Couldn't write config.cfg.\n"
            Exit Sub
        End If

        Key_WriteBindings f
        Cvar_WriteVariables f

        Close #f
    End If
End Sub

' (QuakeVB) Opens a text file for writing; False if it cannot be created.
Public Function OpenForOutput(ByVal path As String, ByVal f As Integer) As Boolean
    On Error GoTo failed
    Open path For Output As #f
    OpenForOutput = True
    Exit Function
failed:
    OpenForOutput = False
End Function

' Sends text across to be displayed
' FIXME: make this just a stuffed echo?
Public Sub SV_ClientPrintf(ByVal fmt As String, ParamArray args() As Variant)
    Dim a As Variant, s As String
    a = args
    s = vsprintf(fmt, a)

    MSG_WriteByte svs.clients(host_client).message, svc_print
    MSG_WriteString svs.clients(host_client).message, s
End Sub

' Sends text to all active clients
Public Sub SV_BroadcastPrintf(ByVal fmt As String, ParamArray args() As Variant)
    Dim a As Variant, s As String, i As Long
    a = args
    s = vsprintf(fmt, a)

    For i = 0 To svs.maxclients - 1
        If svs.clients(i).active And svs.clients(i).spawned Then
            MSG_WriteByte svs.clients(i).message, svc_print
            MSG_WriteString svs.clients(i).message, s
        End If
    Next
End Sub

' Send text over to the client to be executed
Public Sub Host_ClientCommands(ByVal fmt As String, ParamArray args() As Variant)
    Dim a As Variant, s As String
    a = args
    s = vsprintf(fmt, a)

    MSG_WriteByte svs.clients(host_client).message, svc_stufftext
    MSG_WriteString svs.clients(host_client).message, s
End Sub

' Called when the player is getting totally kicked off the host
' if (crash = true), don't bother sending signofs
Public Sub SV_DropClient(ByVal crash As Boolean)
    Dim saveSelf As Long, i As Long

    If Not crash Then
        ' send any final messages (don't check for errors)
        If NET_CanSendMessage(svs.clients(host_client).netconnection) Then
            MSG_WriteByte svs.clients(host_client).message, svc_disconnect
            NET_SendMessage svs.clients(host_client).netconnection, svs.clients(host_client).message
        End If

        If svs.clients(host_client).edict <> 0 And svs.clients(host_client).spawned Then
            ' call the prog function for removing a client
            ' this will set the body to a dead frame, among other things
            saveSelf = pr_global_struct(0).self
            pr_global_struct(0).self = EDICT_TO_PROG(svs.clients(host_client).edict)
            PR_ExecuteProgram pr_global_struct(0).ClientDisconnect
            pr_global_struct(0).self = saveSelf
        End If

        Sys_Printf "Client %s removed\n", svs.clients(host_client).name
    End If

    ' break the net connection
    NET_Close svs.clients(host_client).netconnection
    svs.clients(host_client).netconnection = 0

    ' free the client (the body stays around)
    svs.clients(host_client).active = False
    SV_SetClientName host_client, ""
    svs.clients(host_client).old_frags = -999999
    net_activeconnections = net_activeconnections - 1

    ' send notification to all clients
    For i = 0 To svs.maxclients - 1
        If svs.clients(i).active Then
            MSG_WriteByte svs.clients(i).message, svc_updatename
            MSG_WriteByte svs.clients(i).message, host_client
            MSG_WriteString svs.clients(i).message, ""
            MSG_WriteByte svs.clients(i).message, svc_updatefrags
            MSG_WriteByte svs.clients(i).message, host_client
            MSG_WriteShort svs.clients(i).message, 0
            MSG_WriteByte svs.clients(i).message, svc_updatecolors
            MSG_WriteByte svs.clients(i).message, host_client
            MSG_WriteByte svs.clients(i).message, 0
        End If
    Next
End Sub

' This only happens at the end of a game, not between levels
Public Sub Host_ShutdownServer(ByVal crash As Boolean)
    Dim i As Long, count As Long, buf As sizebuf_t, start As Double
    Dim blank_sv As server_t, blank_client As client_t

    If Not sv.active Then Exit Sub

    sv.active = False

    ' stop all client sounds immediately
    If cls.state = ca_connected Then CL_Disconnect

    ' flush any pending messages - like the score!!!
    start = Sys_FloatTime()
    Do
        count = 0
        For i = 0 To svs.maxclients - 1
            host_client = i
            If svs.clients(i).active And svs.clients(i).message.cursize <> 0 Then
                If NET_CanSendMessage(svs.clients(i).netconnection) Then
                    NET_SendMessage svs.clients(i).netconnection, svs.clients(i).message
                    SZ_Clear svs.clients(i).message
                Else
                    NET_GetMessage svs.clients(i).netconnection
                    count = count + 1
                End If
            End If
        Next
        If (Sys_FloatTime() - start) > 3# Then Exit Do
    Loop While count <> 0

    ' make sure all the clients know we're disconnecting
    SZ_Alloc buf, 4
    MSG_WriteByte buf, svc_disconnect
    count = NET_SendToAll(buf, 5)
    If count Then Con_Printf "Host_ShutdownServer: NET_SendToAll failed for %u clients\n", count

    For i = 0 To svs.maxclients - 1
        host_client = i
        If svs.clients(i).active Then SV_DropClient crash
    Next

    '
    ' clear structures
    '
    sv = blank_sv
    For i = 0 To svs.maxclientslimit - 1
        svs.clients(i) = blank_client
    Next
End Sub

' This clears all the memory used by both the client and server, but does
' not reinitialize anything.
Public Sub Host_ClearMemory()
    Dim blank_sv As server_t, blank_cl As client_state_t

    Con_DPrintf "Clearing memory\n"
    D_FlushCaches
    Mod_ClearAll
    Mod_HunkFreeToMark                   ' C: if (host_hunklevel) Hunk_FreeToLowMark (host_hunklevel)

    cls.signon = 0
    sv = blank_sv
    cl = blank_cl
End Sub

'============================================================================

' Returns false if the time is too short to run a frame
Private Function Host_FilterTime(ByVal time_ As Single) As Boolean
    realtime = realtime + time_

    If Not cls.timedemo And realtime - oldrealtime < 1# / 72# Then
        Host_FilterTime = False          ' framerate is too high
        Exit Function
    End If

    host_frametime = realtime - oldrealtime
    oldrealtime = realtime

    If cvars(cv_host_framerate).value > 0 Then
        host_frametime = cvars(cv_host_framerate).value
    Else
        ' don't allow really long or short frames
        If host_frametime > 0.1 Then host_frametime = 0.1
        If host_frametime < 0.001 Then host_frametime = 0.001
    End If

    Host_FilterTime = True
End Function

' Add them exactly as if they had been typed at the console
Private Sub Host_GetConsoleCommands()
    Dim cmd_ As String
    Do
        cmd_ = Sys_ConsoleInput()
        If Len(cmd_) = 0 Then Exit Do
        Cbuf_AddText cmd_
    Loop
End Sub

Public Sub Host_ServerFrame()
    ' run the world state
    pr_global_struct(0).frametime = host_frametime

    ' set the time and clear the general datagram
    SV_ClearDatagram

    ' check for new clients
    SV_CheckForNewClients

    ' read client messages
    SV_RunClients

    ' move things around and think
    ' always pause in single player if in console or menus
    If Not sv.paused And (svs.maxclients > 1 Or key_dest = key_game) Then SV_Physics

    ' send all messages to the clients
    SV_SendClientMessages
End Sub

' Runs all active servers (C: _Host_Frame)
Private Sub Host_Frame_(ByVal time_ As Single)
    Dim pass1 As Long, pass2 As Long, pass3 As Long

    host_frameran = False
    Sys_ClearFPU                         ' (QuakeVB) see Sys_ClearFPU
    On Error GoTo abortserver            ' C: if (setjmp (host_abortserver)) return;

    ' keep the random time dependent
    rand

    ' decide the simulation time
    If Not Host_FilterTime(time_) Then Exit Sub   ' don't run too fast, or packets will flood out
    host_frameran = True

    ' get new key events
    Sys_SendKeyEvents

    ' allow mice or other external controllers to add commands
    IN_Commands

    ' process console commands
    Cbuf_Execute

    NET_Poll

    ' if running the server locally, make intentions now
    If sv.active Then CL_SendCmd

    '-------------------
    '
    ' server operations
    '
    '-------------------

    ' check for commands typed to the host
    Host_GetConsoleCommands

    If sv.active Then Host_ServerFrame

    '-------------------
    '
    ' client operations
    '
    '-------------------

    ' if running the server remotely, send intentions now after
    ' the incoming messages have been read
    If Not sv.active Then CL_SendCmd

    host_time = host_time + host_frametime

    ' fetch results from server
    If cls.state = ca_connected Then CL_ReadFromServer

    ' update video
    If cvars(cv_host_speeds).value <> 0 Then time1 = Sys_FloatTime()

    SCR_UpdateScreen

    If cvars(cv_host_speeds).value <> 0 Then time2 = Sys_FloatTime()

    ' update audio
    If cls.signon = SIGNONS Then
        S_Update r_origin, vpn, vright, vup
        CL_DecayLights
    Else
        S_Update vec3_origin, vec3_origin, vec3_origin, vec3_origin
    End If

    CDAudio_Update

    If cvars(cv_host_speeds).value <> 0 Then
        pass1 = Fix((time1 - time3) * 1000)
        time3 = Sys_FloatTime()
        pass2 = Fix((time2 - time1) * 1000)
        pass3 = Fix((time3 - time2) * 1000)
        Con_Printf "%3i tot %3i server %3i gfx %3i snd\n", pass1 + pass2 + pass3, pass1, pass2, pass3
    End If

    host_framecount = host_framecount + 1
    Exit Sub

abortserver:
    If Err.Number = HOST_ABORTSERVER Then Exit Sub   ' something bad happened, or the server disconnected
    Sys_RuntimeError "Host_Frame", Err.Number, Err.Description
End Sub

Public Sub Host_Frame(ByVal time_ As Single)
    Dim time1_ As Double, time2_ As Double, i As Long, c As Long, m As Long

    If cvars(cv_serverprofile).value = 0 Then
        Host_Frame_ time_
        Exit Sub
    End If

    time1_ = Sys_FloatTime()
    Host_Frame_ time_
    time2_ = Sys_FloatTime()

    timetotal = timetotal + time2_ - time1_
    timecount = timecount + 1

    If timecount < 1000 Then Exit Sub

    m = Fix(timetotal * 1000 / timecount)
    timecount = 0
    timetotal = 0
    c = 0
    For i = 0 To svs.maxclients - 1
        If svs.clients(i).active Then c = c + 1
    Next

    Con_Printf "serverprofile: %2i clients %2i msec\n", c, m
End Sub

'============================================================================

Public Sub Host_Init(parms As quakeparms_t)
    If standard_quake Then
        minimum_memory = MINIMUM_MEMORY_
    Else
        minimum_memory = MINIMUM_MEMORY_ + &H100000   ' MINIMUM_MEMORY_LEVELPAK
    End If

    ' (-minmemory and the heap size check go: VB allocates memory as needed)

    host_parms = parms

    Cbuf_Init
    Cmd_Init
    V_Init
    Chase_Init
    ' (Host_InitVCR comes with net_vcr, cluster G)
    COM_Init parms.basedir
    Host_InitLocal
    W_LoadWadFile "gfx.wad"
    Key_Init
    Con_Init
    M_Init
    PR_Init
    Mod_Init
    NET_Init
    SV_Init

    Con_Printf "Exe: %s\n", Sys_ExeTime()

    R_InitTextures                       ' needed even for dedicated servers

    If cls.state <> ca_dedicated Then
        If Not COM_LoadHunkFile("gfx/palette.lmp", host_basepal) Then Sys_Error "Couldn't load gfx/palette.lmp"
        If Not COM_LoadHunkFile("gfx/colormap.lmp", host_colormap) Then Sys_Error "Couldn't load gfx/colormap.lmp"

        VID_Init host_basepal
        Draw_Init
        SCR_Init
        R_Init
        ' (on Win32 S_Init is called by VID_Init)
        CDAudio_Init
        Sbar_Init
        CL_Init
        IN_Init
    End If

    Cbuf_InsertText "exec quake.rc" & vbLf

    Mod_HunkMark                         ' C: host_hunklevel = Hunk_LowMark ()

    host_initialized = True

    Sys_Printf "========Quake Initialized=========\n"
End Sub

' FIXME: this is a callback from Sys_Quit and Sys_Error.  It would be better
' to run quit through here before the final handoff to the sys code.
Public Sub Host_Shutdown()
    If shutdown_isdown Then
        Sys_Printf "recursive shutdown\n"
        Exit Sub
    End If
    shutdown_isdown = True

    ' keep Con_Printf from trying to update the screen
    scr_disabled_for_loading = True

    Host_WriteConfiguration

    CDAudio_Shutdown
    NET_Shutdown
    S_Shutdown
    IN_Shutdown

    If cls.state <> ca_dedicated Then VID_Shutdown
End Sub
