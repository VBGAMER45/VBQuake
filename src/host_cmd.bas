Attribute VB_Name = "host_cmd"
' host_cmd.bas - console commands of the host (host_cmd.c).
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const SAVEGAME_VERSION As Long = 5

Public Sub Host_Quit_f()
    If key_dest <> key_console And cls.state <> ca_dedicated Then
        M_Menu_Quit_f
        Exit Sub
    End If
    CL_Disconnect
    Host_ShutdownServer False

    Sys_Quit
End Sub

' (QuakeVB) C's print function pointer of Host_Status_f
Private Sub StatusPrint(ByVal toConsole As Boolean, ByVal fmt As String, ParamArray args() As Variant)
    Dim a As Variant, s As String
    a = args
    s = vsprintf(fmt, a)
    If toConsole Then
        Con_Printf "%s", s
    Else
        SV_ClientPrintf "%s", s
    End If
End Sub

Private Sub Host_Status_f()
    Dim seconds As Long, minutes As Long, hours As Long, j As Long, toConsole As Boolean

    If cmd_source = src_command Then
        If Not sv.active Then
            Cmd_ForwardToServer
            Exit Sub
        End If
        toConsole = True
    Else
        toConsole = False
    End If

    StatusPrint toConsole, "host:    %s\n", Cvar_VariableString("hostname")
    StatusPrint toConsole, "version: %4.2f\n", QUAKE_VERSION
    If tcpipAvailable Then StatusPrint toConsole, "tcp/ip:  %s\n", my_tcpip_address
    If ipxAvailable Then StatusPrint toConsole, "ipx:     %s\n", my_ipx_address
    StatusPrint toConsole, "map:     %s\n", sv.name
    StatusPrint toConsole, "players: %i active (%i max)\n\n", net_activeconnections, svs.maxclients
    For j = 0 To svs.maxclients - 1
        If svs.clients(j).active Then
            seconds = ToInt(net_time - net_qsockets(svs.clients(j).netconnection).connecttime)
            minutes = seconds \ 60
            If minutes Then
                seconds = seconds - (minutes * 60)
                hours = minutes \ 60
                If hours Then minutes = minutes - (hours * 60)
            Else
                hours = 0
            End If
            StatusPrint toConsole, "#%-2u %-16.16s  %3i  %2i:%02i:%02i\n", j + 1, svs.clients(j).name, ToInt(ev(svs.clients(j).edict).frags), hours, minutes, seconds
            StatusPrint toConsole, "   %s\n", net_qsockets(svs.clients(j).netconnection).address
        End If
    Next
End Sub

' Sets client to godmode
Private Sub Host_God_f()
    If cmd_source = src_command Then
        Cmd_ForwardToServer
        Exit Sub
    End If

    If pr_global_struct(0).deathmatch <> 0 And Not svs.clients(host_client).privileged Then Exit Sub

    ev(sv_player).flags = ToInt(ev(sv_player).flags) Xor FL_GODMODE
    If (ToInt(ev(sv_player).flags) And FL_GODMODE) = 0 Then
        SV_ClientPrintf "godmode OFF\n"
    Else
        SV_ClientPrintf "godmode ON\n"
    End If
End Sub

Private Sub Host_Notarget_f()
    If cmd_source = src_command Then
        Cmd_ForwardToServer
        Exit Sub
    End If

    If pr_global_struct(0).deathmatch <> 0 And Not svs.clients(host_client).privileged Then Exit Sub

    ev(sv_player).flags = ToInt(ev(sv_player).flags) Xor FL_NOTARGET
    If (ToInt(ev(sv_player).flags) And FL_NOTARGET) = 0 Then
        SV_ClientPrintf "notarget OFF\n"
    Else
        SV_ClientPrintf "notarget ON\n"
    End If
End Sub

Private Sub Host_Noclip_f()
    If cmd_source = src_command Then
        Cmd_ForwardToServer
        Exit Sub
    End If

    If pr_global_struct(0).deathmatch <> 0 And Not svs.clients(host_client).privileged Then Exit Sub

    If ev(sv_player).movetype <> MOVETYPE_NOCLIP Then
        noclip_anglehack = True
        ev(sv_player).movetype = MOVETYPE_NOCLIP
        SV_ClientPrintf "noclip ON\n"
    Else
        noclip_anglehack = False
        ev(sv_player).movetype = MOVETYPE_WALK
        SV_ClientPrintf "noclip OFF\n"
    End If
End Sub

' Sets client to flymode
Private Sub Host_Fly_f()
    If cmd_source = src_command Then
        Cmd_ForwardToServer
        Exit Sub
    End If

    If pr_global_struct(0).deathmatch <> 0 And Not svs.clients(host_client).privileged Then Exit Sub

    If ev(sv_player).movetype <> MOVETYPE_FLY Then
        ev(sv_player).movetype = MOVETYPE_FLY
        SV_ClientPrintf "flymode ON\n"
    Else
        ev(sv_player).movetype = MOVETYPE_WALK
        SV_ClientPrintf "flymode OFF\n"
    End If
End Sub

Private Sub Host_Ping_f()
    Dim i As Long, j As Long, total As Single

    If cmd_source = src_command Then
        Cmd_ForwardToServer
        Exit Sub
    End If

    SV_ClientPrintf "Client ping times:\n"
    For i = 0 To svs.maxclients - 1
        If svs.clients(i).active Then
            total = 0
            For j = 0 To NUM_PING_TIMES - 1
                total = total + svs.clients(i).ping_times(j)
            Next
            total = total / NUM_PING_TIMES
            SV_ClientPrintf "%4i %s\n", ToInt(total * 1000), svs.clients(i).name
        End If
    Next
End Sub

'===============================================================================
' SERVER TRANSITIONS

' handle a
' map <servername>
' command from the console.  Active clients are kicked off.
Private Sub Host_Map_f()
    Dim i As Long, name As String

    If cmd_source <> src_command Then Exit Sub

    cls.demonum = -1                     ' stop demo loop in case this fails

    CL_Disconnect
    Host_ShutdownServer False

    key_dest = key_game                  ' remove console or menu
    SCR_BeginLoadingPlaque

    cls.mapstring = ""
    For i = 0 To Cmd_Argc() - 1
        cls.mapstring = cls.mapstring & Cmd_Argv(i) & " "
    Next
    cls.mapstring = cls.mapstring & vbLf

    svs.serverflags = 0                  ' haven't completed an episode yet
    name = Cmd_Argv(1)
    SV_SpawnServer name
    If Not sv.active Then Exit Sub

    If cls.state <> ca_dedicated Then
        cls.spawnparms = ""

        For i = 2 To Cmd_Argc() - 1
            cls.spawnparms = cls.spawnparms & Cmd_Argv(i) & " "
        Next

        Cmd_ExecuteString "connect local", src_command
    End If
End Sub

' Goes to a new map, taking all clients along
Private Sub Host_Changelevel_f()
    Dim level As String

    If Cmd_Argc() <> 2 Then
        Con_Printf "changelevel <levelname> : continue game on a new level\n"
        Exit Sub
    End If
    If Not sv.active Or cls.demoplayback Then
        Con_Printf "Only the server may changelevel\n"
        Exit Sub
    End If
    SV_SaveSpawnparms
    level = Cmd_Argv(1)
    SV_SpawnServer level
End Sub

' Restarts the current server for a dead player
Private Sub Host_Restart_f()
    Dim mapname As String

    If cls.demoplayback Or Not sv.active Then Exit Sub

    If cmd_source <> src_command Then Exit Sub
    mapname = sv.name                    ' must copy out, because it gets cleared
                                         ' in sv_spawnserver
    SV_SpawnServer mapname
End Sub

' This command causes the client to wait for the signon messages again.
' This is sent just before a server changes levels
Private Sub Host_Reconnect_f()
    SCR_BeginLoadingPlaque
    cls.signon = 0                       ' need new connection messages
End Sub

' User command to connect to server
Private Sub Host_Connect_f()
    Dim name As String

    cls.demonum = -1                     ' stop demo loop in case this fails
    If cls.demoplayback Then
        CL_StopPlayback
        CL_Disconnect
    End If
    name = Cmd_Argv(1)
    CL_EstablishConnection name
    Host_Reconnect_f
End Sub

'===============================================================================
' LOAD / SAVE GAME

' Writes a SAVEGAME_COMMENT_LENGTH character comment describing the current
Private Function Host_SavegameComment() As String
    Dim text As String, kills As String

    text = Space$(SAVEGAME_COMMENT_LENGTH)
    Mid$(text, 1) = cl.levelname
    kills = va("kills:%3i/%3i", cl.stats(STAT_MONSTERS), cl.stats(STAT_TOTALMONSTERS))
    Mid$(text, 23) = kills
    ' convert space to _ to make stdio happy
    Host_SavegameComment = Replace(text, " ", "_")
End Function

Private Sub Host_Savegame_f()
    Dim name As String, f As Integer, i As Long

    If cmd_source <> src_command Then Exit Sub

    If Not sv.active Then
        Con_Printf "Not playing a local game.\n"
        Exit Sub
    End If

    If cl.intermission Then
        Con_Printf "Can't save in intermission.\n"
        Exit Sub
    End If

    If svs.maxclients <> 1 Then
        Con_Printf "Can't save multiplayer games.\n"
        Exit Sub
    End If

    If Cmd_Argc() <> 2 Then
        Con_Printf "save <savename> : save a game\n"
        Exit Sub
    End If

    If InStr(Cmd_Argv(1), "..") Then
        Con_Printf "Relative pathnames are not allowed.\n"
        Exit Sub
    End If

    For i = 0 To svs.maxclients - 1
        If svs.clients(i).active Then
            If ev(svs.clients(i).edict).health <= 0 Then
                Con_Printf "Can't savegame with a dead player\n"
                Exit Sub
            End If
        End If
    Next

    name = COM_DefaultExtension(com_gamedir & "/" & Cmd_Argv(1), ".sav")

    Con_Printf "Saving game to %s...\n", name
    f = FreeFile
    If Not OpenForOutput(name, f) Then
        Con_Printf "ERROR: couldn't open.\n"
        Exit Sub
    End If

    Print #f, va("%i", SAVEGAME_VERSION)
    Print #f, Host_SavegameComment()
    For i = 0 To NUM_SPAWN_PARMS - 1
        Print #f, va("%f", svs.clients(0).spawn_parms(i))
    Next
    Print #f, va("%d", current_skill)
    Print #f, sv.name
    Print #f, va("%f", sv.time)

    ' write the light styles

    For i = 0 To MAX_LIGHTSTYLES - 1
        If Len(sv.lightstyles(i)) Then
            Print #f, sv.lightstyles(i)
        Else
            Print #f, "m"
        End If
    Next

    ED_WriteGlobals f
    For i = 0 To sv.num_edicts - 1
        ED_Write f, EDICT_NUM(i)
    Next
    Close #f
    Con_Printf "done.\n"
End Sub

' (QuakeVB) fscanf (f, "%s\n", ...): the next whitespace delimited word of
' data() from p, which is moved past it and the whitespace that follows
' (also read by menu.bas's M_ScanSaves)
Public Function ScanWord(data() As Byte, p As Long) As String
    Dim start As Long

    Do While data(p) = 32 Or (data(p) >= 9 And data(p) <= 13)
        p = p + 1
    Loop
    start = p
    Do While data(p) > 32
        p = p + 1
    Loop
    ScanWord = BytesToStr(data, start, p - start)
    Do While data(p) = 32 Or (data(p) >= 9 And data(p) <= 13)
        p = p + 1
    Loop
End Function

Private Sub Host_Loadgame_f()
    Dim name As String, mapname As String, time_ As Single, tfloat As Single
    Dim data() As Byte, h As Long, l As Long, p As Long, i As Long
    Dim ent As Long, entnum As Long, version As Long, spawn_parms(0 To NUM_SPAWN_PARMS - 1) As Single

    If cmd_source <> src_command Then Exit Sub

    If Cmd_Argc() <> 2 Then
        Con_Printf "load <savename> : load a game\n"
        Exit Sub
    End If

    cls.demonum = -1                     ' stop demo loop in case this fails

    name = COM_DefaultExtension(com_gamedir & "/" & Cmd_Argv(1), ".sav")

    ' we can't call SCR_BeginLoadingPlaque, because too much stack space has
    ' been used.  The menu calls it before stuffing loadgame command
    ' SCR_BeginLoadingPlaque ();

    Con_Printf "Loading game from %s...\n", name
    l = Sys_FileOpenRead(name, h)
    If h = -1 Then
        Con_Printf "ERROR: couldn't open.\n"
        Exit Sub
    End If
    ' (QuakeVB: the whole file is read at once, with a 0 byte at the end)
    ReDim data(0 To l)
    If l > 0 Then Sys_FileRead h, data, 0, l
    Sys_FileClose h
    data(l) = 0
    p = 0

    version = C_atoi(ScanWord(data, p))
    If version <> SAVEGAME_VERSION Then
        Con_Printf "Savegame is version %i, not %i\n", version, SAVEGAME_VERSION
        Exit Sub
    End If
    ScanWord data, p                     ' the comment
    For i = 0 To NUM_SPAWN_PARMS - 1
        spawn_parms(i) = C_atof(ScanWord(data, p))
    Next
    ' this silliness is so we can load 1.06 save files, which have float skill values
    tfloat = C_atof(ScanWord(data, p))
    current_skill = ToInt(tfloat + 0.1)
    Cvar_SetValue "skill", current_skill

    mapname = ScanWord(data, p)
    time_ = C_atof(ScanWord(data, p))

    CL_Disconnect_f

    SV_SpawnServer mapname
    If Not sv.active Then
        Con_Printf "Couldn't load map\n"
        Exit Sub
    End If
    sv.paused = True                     ' pause until all clients connect
    sv.loadgame = True

    ' load the light styles

    For i = 0 To MAX_LIGHTSTYLES - 1
        sv.lightstyles(i) = ScanWord(data, p)
    Next

    ' load the edicts out of the savegame file
    ' (QuakeVB: C reads each block up to its '}' into a buffer first; the
    ' parsers stop at the '}' themselves)
    entnum = -1                          ' -1 is the globals
    Do
        p = COM_Parse(data, p)
        If Len(com_token) = 0 Then Exit Do   ' end of file
        If com_token <> "{" Then Sys_Error "First token isn't a brace"

        If entnum = -1 Then
            ' parse the global vars
            p = ED_ParseGlobals(data, p)
        Else
            ' parse an edict

            ent = EDICT_NUM(entnum)
            FillMemory ed_l(ent * EDICT_STRIDE), progs_entityfields * 4, 0
            edicts(ent).free = False
            p = ED_ParseEdict(data, p, ent)

            ' link it into the bsp tree
            If Not edicts(ent).free Then SV_LinkEdict ent, False
        End If

        entnum = entnum + 1
    Loop

    sv.num_edicts = entnum
    sv.time = time_

    For i = 0 To NUM_SPAWN_PARMS - 1
        svs.clients(0).spawn_parms(i) = spawn_parms(i)
    Next

    If cls.state <> ca_dedicated Then
        CL_EstablishConnection "local"
        Host_Reconnect_f
    End If
End Sub

'============================================================================

Private Sub Host_Name_f()
    Dim newName As String

    If Cmd_Argc() = 1 Then
        Con_Printf """name"" is ""%s""\n", cvars(cv_cl_name).strval
        Exit Sub
    End If
    If Cmd_Argc() = 2 Then
        newName = Cmd_Argv(1)
    Else
        newName = Cmd_Args()
    End If
    newName = Left$(newName, 15)

    If cmd_source = src_command Then
        If cvars(cv_cl_name).strval = newName Then Exit Sub
        Cvar_Set "_cl_name", newName
        If cls.state = ca_connected Then Cmd_ForwardToServer
        Exit Sub
    End If

    With svs.clients(host_client)
        If Len(.name) > 0 And .name <> "unconnected" Then
            If .name <> newName Then Con_Printf "%s renamed to %s\n", .name, newName
        End If
        SV_SetClientName host_client, newName
        ev(.edict).netname = PR_ClientNameString(host_client)

        ' send notification to all clients

        MSG_WriteByte sv.reliable_datagram, svc_updatename
        MSG_WriteByte sv.reliable_datagram, host_client
        MSG_WriteString sv.reliable_datagram, .name
    End With
End Sub

Public Sub Host_Version_f()
    Con_Printf "Version %4.2f\n", QUAKE_VERSION
    Con_Printf "Exe: %s\n", Sys_ExeTime()
End Sub

Private Sub Host_Say(ByVal teamonly As Boolean)
    Dim save As Long, j As Long, p As String, text As String, fromServer As Boolean

    If cmd_source = src_command Then
        If cls.state = ca_dedicated Then
            fromServer = True
            teamonly = False
        Else
            Cmd_ForwardToServer
            Exit Sub
        End If
    End If

    If Cmd_Argc() < 2 Then Exit Sub

    save = host_client

    p = Cmd_Args()
    ' remove quotes if present
    If Left$(p, 1) = """" Then
        p = Mid$(p, 2)
        If Len(p) Then p = Left$(p, Len(p) - 1)
    End If

    ' turn on color set 1
    If Not fromServer Then
        text = Chr$(1) & svs.clients(save).name & ": "
    Else
        text = Chr$(1) & "<" & cvars(cv_hostname).strval & "> "
    End If

    j = 64 - 2 - Len(text)               ' -2 for /n and null terminator
    If Len(p) > j Then p = Left$(p, j)

    text = text & p & vbLf

    For j = 0 To svs.maxclients - 1
        If svs.clients(j).active And svs.clients(j).spawned Then
            If cvars(cv_teamplay).value <> 0 And teamonly Then
                If ev(svs.clients(j).edict).team <> ev(svs.clients(save).edict).team Then GoTo nextj
            End If
            host_client = j
            SV_ClientPrintf "%s", text
        End If
nextj:
    Next
    host_client = save

    Sys_Printf "%s", Mid$(text, 2)
End Sub

Private Sub Host_Say_f()
    Host_Say False
End Sub

Private Sub Host_Say_Team_f()
    Host_Say True
End Sub

Private Sub Host_Tell_f()
    Dim save As Long, j As Long, p As String, text As String

    If cmd_source = src_command Then
        Cmd_ForwardToServer
        Exit Sub
    End If

    If Cmd_Argc() < 3 Then Exit Sub

    text = svs.clients(host_client).name & ": "

    p = Cmd_Args()

    ' remove quotes if present
    If Left$(p, 1) = """" Then
        p = Mid$(p, 2)
        If Len(p) Then p = Left$(p, Len(p) - 1)
    End If

    ' check length & truncate if necessary
    j = 64 - 2 - Len(text)               ' -2 for /n and null terminator
    If Len(p) > j Then p = Left$(p, j)

    text = text & p & vbLf

    save = host_client
    For j = 0 To svs.maxclients - 1
        If svs.clients(j).active And svs.clients(j).spawned Then
            If Q_strcasecmp(svs.clients(j).name, Cmd_Argv(1)) = 0 Then
                host_client = j
                SV_ClientPrintf "%s", text
                Exit For
            End If
        End If
    Next
    host_client = save
End Sub

Private Sub Host_Color_f()
    Dim top As Long, bottom As Long, playercolor As Long

    If Cmd_Argc() = 1 Then
        Con_Printf """color"" is ""%i %i""\n", (ToInt(cvars(cv_cl_color).value) And Not 15) \ 16, ToInt(cvars(cv_cl_color).value) And 15
        Con_Printf "color <0-13> [0-13]\n"
        Exit Sub
    End If

    If Cmd_Argc() = 2 Then
        top = C_atoi(Cmd_Argv(1))
        bottom = top
    Else
        top = C_atoi(Cmd_Argv(1))
        bottom = C_atoi(Cmd_Argv(2))
    End If

    top = top And 15
    If top > 13 Then top = 13
    bottom = bottom And 15
    If bottom > 13 Then bottom = 13

    playercolor = top * 16 + bottom

    If cmd_source = src_command Then
        Cvar_SetValue "_cl_color", playercolor
        If cls.state = ca_connected Then Cmd_ForwardToServer
        Exit Sub
    End If

    svs.clients(host_client).colors = playercolor
    ev(svs.clients(host_client).edict).team = bottom + 1

    ' send notification to all clients
    MSG_WriteByte sv.reliable_datagram, svc_updatecolors
    MSG_WriteByte sv.reliable_datagram, host_client
    MSG_WriteByte sv.reliable_datagram, svs.clients(host_client).colors
End Sub

Private Sub Host_Kill_f()
    If cmd_source = src_command Then
        Cmd_ForwardToServer
        Exit Sub
    End If

    If ev(sv_player).health <= 0 Then
        SV_ClientPrintf "Can't suicide -- allready dead!\n"
        Exit Sub
    End If

    pr_global_struct(0).time = sv.time
    pr_global_struct(0).self = EDICT_TO_PROG(sv_player)
    PR_ExecuteProgram pr_global_struct(0).ClientKill
End Sub

Private Sub Host_Pause_f()
    If cmd_source = src_command Then
        Cmd_ForwardToServer
        Exit Sub
    End If
    If cvars(cv_pausable).value = 0 Then
        SV_ClientPrintf "Pause not allowed.\n"
    Else
        sv.paused = Not sv.paused

        If sv.paused Then
            SV_BroadcastPrintf "%s paused the game\n", PR_GetString(ev(sv_player).netname)
        Else
            SV_BroadcastPrintf "%s unpaused the game\n", PR_GetString(ev(sv_player).netname)
        End If

        ' send notification to all clients
        MSG_WriteByte sv.reliable_datagram, svc_setpause
        If sv.paused Then
            MSG_WriteByte sv.reliable_datagram, 1
        Else
            MSG_WriteByte sv.reliable_datagram, 0
        End If
    End If
End Sub

'===========================================================================

Private Sub Host_PreSpawn_f()
    If cmd_source = src_command Then
        Con_Printf "prespawn is not valid from the console\n"
        Exit Sub
    End If

    If svs.clients(host_client).spawned Then
        Con_Printf "prespawn not valid -- allready spawned\n"
        Exit Sub
    End If

    SZ_Write svs.clients(host_client).message, sv.signon.data, 0, sv.signon.cursize
    MSG_WriteByte svs.clients(host_client).message, svc_signonnum
    MSG_WriteByte svs.clients(host_client).message, 2
    svs.clients(host_client).sendsignon = True
End Sub

Private Sub Host_Spawn_f()
    Dim i As Long, ent As Long

    If cmd_source = src_command Then
        Con_Printf "spawn is not valid from the console\n"
        Exit Sub
    End If

    If svs.clients(host_client).spawned Then
        Con_Printf "Spawn not valid -- allready spawned\n"
        Exit Sub
    End If

    ' run the entrance script
    If sv.loadgame Then
        ' loaded games are fully inited allready
        ' if this is the last client to be connected, unpause
        sv.paused = False
    Else
        ' set up the edict
        ent = svs.clients(host_client).edict

        FillMemory ed_l(ent * EDICT_STRIDE), progs_entityfields * 4, 0
        ev(ent).colormap = NUM_FOR_EDICT(ent)
        ev(ent).team = (svs.clients(host_client).colors And 15) + 1
        ev(ent).netname = PR_ClientNameString(host_client)

        ' copy spawn parms out of the client_t

        For i = 0 To NUM_SPAWN_PARMS - 1
            pr_gf(GOFS_parm1 + i) = svs.clients(host_client).spawn_parms(i)
        Next

        ' call the spawn function

        pr_global_struct(0).time = sv.time
        pr_global_struct(0).self = EDICT_TO_PROG(sv_player)
        PR_ExecuteProgram pr_global_struct(0).ClientConnect

        If (Sys_FloatTime() - net_qsockets(svs.clients(host_client).netconnection).connecttime) <= sv.time Then
            Sys_Printf "%s entered the game\n", svs.clients(host_client).name
        End If

        PR_ExecuteProgram pr_global_struct(0).PutClientInServer
    End If

    With svs.clients(host_client)
        ' send all current names, colors, and frag counts
        SZ_Clear .message

        ' send time of update
        MSG_WriteByte .message, svc_time
        MSG_WriteFloat .message, sv.time

        For i = 0 To svs.maxclients - 1
            MSG_WriteByte .message, svc_updatename
            MSG_WriteByte .message, i
            MSG_WriteString .message, svs.clients(i).name
            MSG_WriteByte .message, svc_updatefrags
            MSG_WriteByte .message, i
            MSG_WriteShort .message, svs.clients(i).old_frags
            MSG_WriteByte .message, svc_updatecolors
            MSG_WriteByte .message, i
            MSG_WriteByte .message, svs.clients(i).colors
        Next

        ' send all current light styles
        For i = 0 To MAX_LIGHTSTYLES - 1
            MSG_WriteByte .message, svc_lightstyle
            MSG_WriteByte .message, i
            MSG_WriteString .message, sv.lightstyles(i)
        Next

        '
        ' send some stats
        '
        MSG_WriteByte .message, svc_updatestat
        MSG_WriteByte .message, STAT_TOTALSECRETS
        MSG_WriteLong .message, pr_global_struct(0).total_secrets

        MSG_WriteByte .message, svc_updatestat
        MSG_WriteByte .message, STAT_TOTALMONSTERS
        MSG_WriteLong .message, pr_global_struct(0).total_monsters

        MSG_WriteByte .message, svc_updatestat
        MSG_WriteByte .message, STAT_SECRETS
        MSG_WriteLong .message, pr_global_struct(0).found_secrets

        MSG_WriteByte .message, svc_updatestat
        MSG_WriteByte .message, STAT_MONSTERS
        MSG_WriteLong .message, pr_global_struct(0).killed_monsters

        '
        ' send a fixangle
        ' Never send a roll angle, because savegames can catch the server
        ' in a state where it is expecting the client to correct the angle
        ' and it won't happen if the game was just loaded, so you wind up
        ' with a permanent head tilt
        ent = EDICT_NUM(1 + host_client)
        MSG_WriteByte .message, svc_setangle
        For i = 0 To 1
            MSG_WriteAngle .message, ev(ent).angles.v(i)
        Next
        MSG_WriteAngle .message, 0

        SV_WriteClientdataToMessage sv_player, .message

        MSG_WriteByte .message, svc_signonnum
        MSG_WriteByte .message, 3
        .sendsignon = True
    End With
End Sub

Private Sub Host_Begin_f()
    If cmd_source = src_command Then
        Con_Printf "begin is not valid from the console\n"
        Exit Sub
    End If

    svs.clients(host_client).spawned = True
End Sub

'===========================================================================

' Kicks a user off of the server
Private Sub Host_Kick_f()
    Dim who As String, message As String, hasMessage As Boolean, save As Long, i As Long
    Dim byNumber As Boolean, args As String, b() As Byte, p As Long

    If cmd_source = src_command Then
        If Not sv.active Then
            Cmd_ForwardToServer
            Exit Sub
        End If
    ElseIf pr_global_struct(0).deathmatch <> 0 And Not svs.clients(host_client).privileged Then
        Exit Sub
    End If

    save = host_client

    If Cmd_Argc() > 2 And Cmd_Argv(1) = "#" Then
        i = ToInt(Q_atof(Cmd_Argv(2)) - 1)
        If i < 0 Or i >= svs.maxclients Then Exit Sub
        If Not svs.clients(i).active Then Exit Sub
        host_client = i
        byNumber = True
    Else
        For i = 0 To svs.maxclients - 1
            host_client = i
            If svs.clients(i).active Then
                If Q_strcasecmp(svs.clients(i).name, Cmd_Argv(1)) = 0 Then Exit For
            End If
        Next
    End If

    If i < svs.maxclients Then
        If cmd_source = src_command Then
            If cls.state = ca_dedicated Then
                who = "Console"
            Else
                who = cvars(cv_cl_name).strval
            End If
        Else
            who = svs.clients(save).name
        End If

        ' can't kick yourself!
        If host_client = save Then Exit Sub

        If Cmd_Argc() > 2 Then
            args = Cmd_Args()
            StrToBytesZ args, b
            p = COM_Parse(b, 0)
            If p >= 0 Then
                hasMessage = True
                message = Mid$(args, p + 1)
                If byNumber Then
                    message = Mid$(message, 2)   ' skip the #
                    Do While Left$(message, 1) = " "   ' skip white space
                        message = Mid$(message, 2)
                    Loop
                    message = Mid$(message, Len(Cmd_Argv(2)) + 1)   ' skip the number
                End If
                Do While Left$(message, 1) = " "
                    message = Mid$(message, 2)
                Loop
            End If
        End If
        If hasMessage Then
            SV_ClientPrintf "Kicked by %s: %s\n", who, message
        Else
            SV_ClientPrintf "Kicked by %s\n", who
        End If
        SV_DropClient False
    End If

    host_client = save
End Sub

'===============================================================================
' DEBUGGING TOOLS

' (QuakeVB) GetEdictFieldValue (sv_player, field)->_float = v, for the
' rogue ammo; returns False if the progs have no such field
Private Function SetPlayerField(ByVal field As String, ByVal v As Long) As Boolean
    Dim val As Long
    val = GetEdictFieldValue(sv_player, field)
    If val <> -1 Then
        ed_f(val) = v
        SetPlayerField = True
    End If
End Function

Private Sub Host_Give_f()
    Dim t As String, c As String, v As Long

    If cmd_source = src_command Then
        Cmd_ForwardToServer
        Exit Sub
    End If

    If pr_global_struct(0).deathmatch <> 0 And Not svs.clients(host_client).privileged Then Exit Sub

    t = Cmd_Argv(1)
    v = C_atoi(Cmd_Argv(2))
    c = Left$(t, 1)

    With ev(sv_player)
        Select Case c
        Case "0" To "9"
            ' MED 01/04/97 added hipnotic give stuff
            If hipnotic Then
                If c = "6" Then
                    If Mid$(t, 2, 1) = "a" Then
                        .items = ToInt(.items) Or HIT_PROXIMITY_GUN
                    Else
                        .items = ToInt(.items) Or IT_GRENADE_LAUNCHER
                    End If
                ElseIf c = "9" Then
                    .items = ToInt(.items) Or HIT_LASER_CANNON
                ElseIf c = "0" Then
                    .items = ToInt(.items) Or HIT_MJOLNIR
                ElseIf c >= "2" Then
                    .items = ToInt(.items) Or CLng(IT_SHOTGUN * 2 ^ (AscW(c) - 50))
                End If
            Else
                If c >= "2" Then .items = ToInt(.items) Or CLng(IT_SHOTGUN * 2 ^ (AscW(c) - 50))
            End If

        Case "s"
            If rogue Then SetPlayerField "ammo_shells1", v

            .ammo_shells = v
        Case "n"
            If rogue Then
                If SetPlayerField("ammo_nails1", v) Then
                    If .weapon <= IT_LIGHTNING Then .ammo_nails = v
                End If
            Else
                .ammo_nails = v
            End If
        Case "l"
            If rogue Then
                If SetPlayerField("ammo_lava_nails", v) Then
                    If .weapon > IT_LIGHTNING Then .ammo_nails = v
                End If
            End If
        Case "r"
            If rogue Then
                If SetPlayerField("ammo_rockets1", v) Then
                    If .weapon <= IT_LIGHTNING Then .ammo_rockets = v
                End If
            Else
                .ammo_rockets = v
            End If
        Case "m"
            If rogue Then
                If SetPlayerField("ammo_multi_rockets", v) Then
                    If .weapon > IT_LIGHTNING Then .ammo_rockets = v
                End If
            End If
        Case "h"
            .health = v
        Case "c"
            If rogue Then
                If SetPlayerField("ammo_cells1", v) Then
                    If .weapon <= IT_LIGHTNING Then .ammo_cells = v
                End If
            Else
                .ammo_cells = v
            End If
        Case "p"
            If rogue Then
                If SetPlayerField("ammo_plasma", v) Then
                    If .weapon > IT_LIGHTNING Then .ammo_cells = v
                End If
            End If
        End Select
    End With
End Sub

' returns -1 if there is none (C: NULL)
Private Function FindViewthing() As Long
    Dim i As Long

    For i = 0 To sv.num_edicts - 1
        If PR_GetString(ev(EDICT_NUM(i)).classname) = "viewthing" Then
            FindViewthing = EDICT_NUM(i)
            Exit Function
        End If
    Next
    Con_Printf "No viewthing on map\n"
    FindViewthing = -1
End Function

Private Sub Host_Viewmodel_f()
    Dim e As Long, m As Long

    e = FindViewthing()
    If e = -1 Then Exit Sub

    m = Mod_ForName(Cmd_Argv(1), False)
    If m = 0 Then
        Con_Printf "Can't load %s\n", Cmd_Argv(1)
        Exit Sub
    End If

    ev(e).frame = 0
    cl.model_precache(ToInt(ev(e).modelindex)) = m
End Sub

Private Sub Host_Viewframe_f()
    Dim e As Long, f As Long, m As Long

    e = FindViewthing()
    If e = -1 Then Exit Sub
    m = cl.model_precache(ToInt(ev(e).modelindex))
    If m = 0 Then Exit Sub               ' (QuakeVB: C dereferences NULL)

    f = C_atoi(Cmd_Argv(1))
    If f >= mod_known(m).numframes Then f = mod_known(m).numframes - 1

    ev(e).frame = f
End Sub

Private Sub PrintFrameName(ByVal m As Long, ByVal frame As Long)
    Dim hdr As Long

    hdr = Mod_Extradata(m)
    If hdr = 0 Then Exit Sub

    Con_Printf "frame %i: %s\n", frame, aliashdrs(hdr).frames(frame).name
End Sub

Private Sub Host_Viewnext_f()
    Dim e As Long, m As Long

    e = FindViewthing()
    If e = -1 Then Exit Sub
    m = cl.model_precache(ToInt(ev(e).modelindex))
    If m = 0 Then Exit Sub               ' (QuakeVB: C dereferences NULL)

    ev(e).frame = ev(e).frame + 1
    If ev(e).frame >= mod_known(m).numframes Then ev(e).frame = mod_known(m).numframes - 1

    PrintFrameName m, ToInt(ev(e).frame)
End Sub

Private Sub Host_Viewprev_f()
    Dim e As Long, m As Long

    e = FindViewthing()
    If e = -1 Then Exit Sub

    m = cl.model_precache(ToInt(ev(e).modelindex))
    If m = 0 Then Exit Sub               ' (QuakeVB: C dereferences NULL)

    ev(e).frame = ev(e).frame - 1
    If ev(e).frame < 0 Then ev(e).frame = 0

    PrintFrameName m, ToInt(ev(e).frame)
End Sub

'===============================================================================
' DEMO LOOP CONTROL

Private Sub Host_Startdemos_f()
    Dim i As Long, c As Long

    If cls.state = ca_dedicated Then
        If Not sv.active Then Cbuf_AddText "map start" & vbLf
        Exit Sub
    End If

    c = Cmd_Argc() - 1
    If c > MAX_DEMOS Then
        Con_Printf "Max %i demos in demoloop\n", MAX_DEMOS
        c = MAX_DEMOS
    End If
    Con_Printf "%i demo(s) in loop\n", c

    For i = 1 To c
        cls.demos(i - 1) = Left$(Cmd_Argv(i), 15)
    Next

    If Not sv.active And cls.demonum <> -1 And Not cls.demoplayback Then
        cls.demonum = 0
        CL_NextDemo
    Else
        cls.demonum = -1
    End If
End Sub

' Return to looping demos
Private Sub Host_Demos_f()
    If cls.state = ca_dedicated Then Exit Sub
    If cls.demonum = -1 Then cls.demonum = 1
    CL_Disconnect_f
    CL_NextDemo
End Sub

' Return to looping demos
Private Sub Host_Stopdemo_f()
    If cls.state = ca_dedicated Then Exit Sub
    If Not cls.demoplayback Then Exit Sub
    CL_StopPlayback
    CL_Disconnect
End Sub

'=============================================================================

' (QuakeVB) the console commands of this module, for Cmd_Call
Public Sub Host_CallCommand(ByVal fn As Long)
    Select Case fn
    Case XC_Host_Status_f: Host_Status_f
    Case XC_Host_Quit_f: Host_Quit_f
    Case XC_Host_God_f: Host_God_f
    Case XC_Host_Notarget_f: Host_Notarget_f
    Case XC_Host_Fly_f: Host_Fly_f
    Case XC_Host_Map_f: Host_Map_f
    Case XC_Host_Restart_f: Host_Restart_f
    Case XC_Host_Changelevel_f: Host_Changelevel_f
    Case XC_Host_Connect_f: Host_Connect_f
    Case XC_Host_Reconnect_f: Host_Reconnect_f
    Case XC_Host_Name_f: Host_Name_f
    Case XC_Host_Noclip_f: Host_Noclip_f
    Case XC_Host_Version_f: Host_Version_f
    Case XC_Host_Say_f: Host_Say_f
    Case XC_Host_Say_Team_f: Host_Say_Team_f
    Case XC_Host_Tell_f: Host_Tell_f
    Case XC_Host_Color_f: Host_Color_f
    Case XC_Host_Kill_f: Host_Kill_f
    Case XC_Host_Pause_f: Host_Pause_f
    Case XC_Host_Spawn_f: Host_Spawn_f
    Case XC_Host_Begin_f: Host_Begin_f
    Case XC_Host_PreSpawn_f: Host_PreSpawn_f
    Case XC_Host_Kick_f: Host_Kick_f
    Case XC_Host_Ping_f: Host_Ping_f
    Case XC_Host_Loadgame_f: Host_Loadgame_f
    Case XC_Host_Savegame_f: Host_Savegame_f
    Case XC_Host_Give_f: Host_Give_f
    Case XC_Host_Startdemos_f: Host_Startdemos_f
    Case XC_Host_Demos_f: Host_Demos_f
    Case XC_Host_Stopdemo_f: Host_Stopdemo_f
    Case XC_Host_Viewmodel_f: Host_Viewmodel_f
    Case XC_Host_Viewframe_f: Host_Viewframe_f
    Case XC_Host_Viewnext_f: Host_Viewnext_f
    Case XC_Host_Viewprev_f: Host_Viewprev_f
    Case Else: Sys_Error "Host_CallCommand: bad function %i", fn
    End Select
End Sub

Public Sub Host_InitCommands()
    Cmd_AddCommand "status", XC_Host_Status_f
    Cmd_AddCommand "quit", XC_Host_Quit_f
    Cmd_AddCommand "god", XC_Host_God_f
    Cmd_AddCommand "notarget", XC_Host_Notarget_f
    Cmd_AddCommand "fly", XC_Host_Fly_f
    Cmd_AddCommand "map", XC_Host_Map_f
    Cmd_AddCommand "restart", XC_Host_Restart_f
    Cmd_AddCommand "changelevel", XC_Host_Changelevel_f
    Cmd_AddCommand "connect", XC_Host_Connect_f
    Cmd_AddCommand "reconnect", XC_Host_Reconnect_f
    Cmd_AddCommand "name", XC_Host_Name_f
    Cmd_AddCommand "noclip", XC_Host_Noclip_f
    Cmd_AddCommand "version", XC_Host_Version_f
    Cmd_AddCommand "say", XC_Host_Say_f
    Cmd_AddCommand "say_team", XC_Host_Say_Team_f
    Cmd_AddCommand "tell", XC_Host_Tell_f
    Cmd_AddCommand "color", XC_Host_Color_f
    Cmd_AddCommand "kill", XC_Host_Kill_f
    Cmd_AddCommand "pause", XC_Host_Pause_f
    Cmd_AddCommand "spawn", XC_Host_Spawn_f
    Cmd_AddCommand "begin", XC_Host_Begin_f
    Cmd_AddCommand "prespawn", XC_Host_PreSpawn_f
    Cmd_AddCommand "kick", XC_Host_Kick_f
    Cmd_AddCommand "ping", XC_Host_Ping_f
    Cmd_AddCommand "load", XC_Host_Loadgame_f
    Cmd_AddCommand "save", XC_Host_Savegame_f
    Cmd_AddCommand "give", XC_Host_Give_f

    Cmd_AddCommand "startdemos", XC_Host_Startdemos_f
    Cmd_AddCommand "demos", XC_Host_Demos_f
    Cmd_AddCommand "stopdemo", XC_Host_Stopdemo_f

    Cmd_AddCommand "viewmodel", XC_Host_Viewmodel_f
    Cmd_AddCommand "viewframe", XC_Host_Viewframe_f
    Cmd_AddCommand "viewnext", XC_Host_Viewnext_f
    Cmd_AddCommand "viewprev", XC_Host_Viewprev_f

    Cmd_AddCommand "mcache", XC_Mod_Print
End Sub
