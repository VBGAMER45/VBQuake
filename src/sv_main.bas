Attribute VB_Name = "sv_main"
' sv_main.bas - server main program (sv_main.c)
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private localmodels(0 To MAX_MODELS - 1) As String   ' inline model names for precache

Private fatbytes As Long
Private fatpvs(0 To MAX_MAP_LEAFS \ 8 - 1) As Byte

'============================================================================

Public Sub SV_Init()
    Dim i As Long

    cv_sv_maxvelocity = Cvar_RegisterVariable("sv_maxvelocity", "2000")
    cv_sv_gravity = Cvar_RegisterVariable("sv_gravity", "800", False, True)
    cv_sv_friction = Cvar_RegisterVariable("sv_friction", "4", False, True)
    cv_sv_edgefriction = Cvar_RegisterVariable("edgefriction", "2")
    cv_sv_stopspeed = Cvar_RegisterVariable("sv_stopspeed", "100")
    cv_sv_maxspeed = Cvar_RegisterVariable("sv_maxspeed", "320", False, True)
    cv_sv_accelerate = Cvar_RegisterVariable("sv_accelerate", "10")
    cv_sv_idealpitchscale = Cvar_RegisterVariable("sv_idealpitchscale", "0.8")
    cv_sv_aim = Cvar_RegisterVariable("sv_aim", "0.93")
    cv_sv_nostep = Cvar_RegisterVariable("sv_nostep", "0")

    For i = 0 To MAX_MODELS - 1
        localmodels(i) = "*" & i
    Next
End Sub

'=============================================================================
' EVENT MESSAGES

' Make sure the event gets sent to all clients
Public Sub SV_StartParticle(org As vec3_t, dir_ As vec3_t, ByVal color As Long, ByVal count As Long)
    Dim i As Long, v As Long

    If sv.datagram.cursize > MAX_DATAGRAM - 16 Then Exit Sub
    MSG_WriteByte sv.datagram, svc_particle
    MSG_WriteCoord sv.datagram, org.v(0)
    MSG_WriteCoord sv.datagram, org.v(1)
    MSG_WriteCoord sv.datagram, org.v(2)
    For i = 0 To 2
        v = ToInt(dir_.v(i) * 16)
        If v > 127 Then
            v = 127
        ElseIf v < -128 Then
            v = -128
        End If
        MSG_WriteChar sv.datagram, v
    Next
    MSG_WriteByte sv.datagram, count
    MSG_WriteByte sv.datagram, color
End Sub

' Each entity can have eight independant sound sources, like voice,
' weapon, feet, etc.
'
' Channel 0 is an auto-allocate channel, the others override anything
' allready running on that entity/channel pair.
'
' An attenuation of 0 will play full volume everywhere in the level.
' Larger attenuations will drop off.  (max 4 attenuation)
Public Sub SV_StartSound(ByVal entity As Long, ByVal channel As Long, ByVal sample As String, ByVal volume As Long, _
        ByVal attenuation As Single)
    Dim sound_num As Long, field_mask As Long, i As Long, ent As Long

    If volume < 0 Or volume > 255 Then Sys_Error "SV_StartSound: volume = %i", volume

    If attenuation < 0 Or attenuation > 4 Then Sys_Error "SV_StartSound: attenuation = %f", attenuation

    If channel < 0 Or channel > 7 Then Sys_Error "SV_StartSound: channel = %i", channel

    If sv.datagram.cursize > MAX_DATAGRAM - 16 Then Exit Sub

    ' find precache number for sound
    For sound_num = 1 To sv.num_sound_precache - 1
        If sample = sv.sound_precache(sound_num) Then Exit For
    Next

    If sound_num >= sv.num_sound_precache Then
        Con_Printf "SV_StartSound: %s not precacheed\n", sample
        Exit Sub
    End If

    ent = NUM_FOR_EDICT(entity)

    channel = (ent * 8) Or channel

    field_mask = 0
    If volume <> DEFAULT_SOUND_PACKET_VOLUME Then field_mask = field_mask Or SND_VOLUME
    If attenuation <> DEFAULT_SOUND_PACKET_ATTENUATION Then field_mask = field_mask Or SND_ATTENUATION

    ' directed messages go only to the entity the are targeted on
    MSG_WriteByte sv.datagram, svc_sound
    MSG_WriteByte sv.datagram, field_mask
    If field_mask And SND_VOLUME Then MSG_WriteByte sv.datagram, volume
    If field_mask And SND_ATTENUATION Then MSG_WriteByte sv.datagram, attenuation * 64
    MSG_WriteShort sv.datagram, channel
    MSG_WriteByte sv.datagram, sound_num
    For i = 0 To 2
        MSG_WriteCoord sv.datagram, ev(entity).origin.v(i) + 0.5 * (ev(entity).mins.v(i) + ev(entity).maxs.v(i))
    Next
End Sub

'==============================================================================
' CLIENT SPAWNING

' Sends the first message from the server to a connected client.
' This will be sent on the initial connection and upon each server load.
Public Sub SV_SendServerinfo(ByVal client As Long)
    Dim i As Long, message As String

    MSG_WriteByte svs.clients(client).message, svc_print
    message = va("%c\nVERSION %4.2f SERVER (%i CRC)", 2, QUAKE_VERSION, pr_crc)
    MSG_WriteString svs.clients(client).message, message

    MSG_WriteByte svs.clients(client).message, svc_serverinfo
    MSG_WriteLong svs.clients(client).message, PROTOCOL_VERSION
    MSG_WriteByte svs.clients(client).message, svs.maxclients

    If cvars(cv_coop).value = 0 And cvars(cv_deathmatch).value <> 0 Then
        MSG_WriteByte svs.clients(client).message, GAME_DEATHMATCH
    Else
        MSG_WriteByte svs.clients(client).message, GAME_COOP
    End If

    message = PR_GetString(ev(0).message)

    MSG_WriteString svs.clients(client).message, message

    For i = 1 To sv.num_model_precache - 1
        MSG_WriteString svs.clients(client).message, sv.model_precache(i)
    Next
    MSG_WriteByte svs.clients(client).message, 0

    For i = 1 To sv.num_sound_precache - 1
        MSG_WriteString svs.clients(client).message, sv.sound_precache(i)
    Next
    MSG_WriteByte svs.clients(client).message, 0

    ' send music
    MSG_WriteByte svs.clients(client).message, svc_cdtrack
    MSG_WriteByte svs.clients(client).message, ev(0).sounds
    MSG_WriteByte svs.clients(client).message, ev(0).sounds

    ' set view
    MSG_WriteByte svs.clients(client).message, svc_setview
    MSG_WriteShort svs.clients(client).message, NUM_FOR_EDICT(svs.clients(client).edict)

    MSG_WriteByte svs.clients(client).message, svc_signonnum
    MSG_WriteByte svs.clients(client).message, 1

    svs.clients(client).sendsignon = True
    svs.clients(client).spawned = False  ' need prespawn, spawn, etc
End Sub

' Initializes a client_t for a new net connection.  This will only be called
' once for a player each game, not once for each level change.
Public Sub SV_ConnectClient(ByVal clientnum As Long)
    Dim ent As Long, edictnum As Long, netconnection As Long, i As Long
    Dim spawn_parms(0 To NUM_SPAWN_PARMS - 1) As Single, blank As client_t

    Con_DPrintf "Client %s connected\n", net_qsockets(svs.clients(clientnum).netconnection).address

    edictnum = clientnum + 1

    ent = EDICT_NUM(edictnum)

    ' set up the client_t
    netconnection = svs.clients(clientnum).netconnection

    If sv.loadgame Then
        For i = 0 To NUM_SPAWN_PARMS - 1
            spawn_parms(i) = svs.clients(clientnum).spawn_parms(i)
        Next
    End If
    svs.clients(clientnum) = blank
    svs.clients(clientnum).netconnection = netconnection

    SV_SetClientName clientnum, "unconnected"
    svs.clients(clientnum).active = True
    svs.clients(clientnum).spawned = False
    svs.clients(clientnum).edict = ent
    SZ_Alloc svs.clients(clientnum).message, MAX_MSGLEN
    svs.clients(clientnum).message.allowoverflow = True  ' we can catch it

    svs.clients(clientnum).privileged = False

    If sv.loadgame Then
        For i = 0 To NUM_SPAWN_PARMS - 1
            svs.clients(clientnum).spawn_parms(i) = spawn_parms(i)
        Next
    Else
        ' call the progs to get default spawn parms for the new client
        PR_ExecuteProgram pr_global_struct(0).SetNewParms
        For i = 0 To NUM_SPAWN_PARMS - 1
            svs.clients(clientnum).spawn_parms(i) = pr_gf(GOFS_parm1 + i)
        Next
    End If

    SV_SendServerinfo clientnum
End Sub

Public Sub SV_CheckForNewClients()
    Dim ret As Long, i As Long

    '
    ' check for new connections
    '
    Do
        ret = NET_CheckNewConnections()
        If ret = 0 Then Exit Do

        '
        ' init a new client structure
        '
        For i = 0 To svs.maxclients - 1
            If Not svs.clients(i).active Then Exit For
        Next
        If i = svs.maxclients Then Sys_Error "Host_CheckForNewClients: no free clients"

        svs.clients(i).netconnection = ret
        SV_ConnectClient i
        net_activeconnections = net_activeconnections + 1
    Loop
End Sub

'===============================================================================
' FRAME UPDATES

Public Sub SV_ClearDatagram()
    SZ_Clear sv.datagram
End Sub

' The PVS must include a small area around the client to allow head bobbing
' or other small motion on the client side.  Otherwise, a bob might cause an
' entity that should be visible to not show up, especially when the bob
' crosses a waterline.
Private Sub SV_AddToFatPVS(org As vec3_t, ByVal node As Long)
    Dim i As Long, plane As Long, d As Single

    Do
        ' if this is a leaf, accumulate the pvs bits
        If mod_nodes(node).contents < 0 Then
            If mod_nodes(node).contents <> CONTENTS_SOLID Then
                Mod_LeafPVS node, sv.worldmodel
                For i = 0 To fatbytes - 1
                    fatpvs(i) = fatpvs(i) Or mod_pvs(i)
                Next
            End If
            Exit Sub
        End If

        plane = mod_nodes(node).plane
        d = DotProduct(org, mod_planes(plane).normal) - mod_planes(plane).dist
        If d > 8 Then
            node = mod_nodes(node).children(0)
        ElseIf d < -8 Then
            node = mod_nodes(node).children(1)
        Else
            ' go down both
            SV_AddToFatPVS org, mod_nodes(node).children(0)
            node = mod_nodes(node).children(1)
        End If
    Loop
End Sub

' Calculates a PVS that is the inclusive or of all leafs within 8 pixels of the
' given point.
Private Sub SV_FatPVS(org As vec3_t)
    fatbytes = (mod_known(sv.worldmodel).numleafs + 31) \ 8
    If fatbytes > 0 Then FillMemory fatpvs(0), fatbytes, 0
    SV_AddToFatPVS org, mod_known(sv.worldmodel).nodes
End Sub

'=============================================================================

Private Sub SV_WriteEntitiesToClient(ByVal clent As Long, msg As sizebuf_t)
    Dim e As Long, i As Long, bits As Long, org As vec3_t, miss As Single, ent As Long, l As Long

    ' find the client's PVS
    VectorAdd ev(clent).origin, ev(clent).view_ofs, org
    SV_FatPVS org

    ' send over all entities (excpet the client) that touch the pvs
    For e = 1 To sv.num_edicts - 1
        ent = e
        ' ignore if not touching a PV leaf
        If ent <> clent Then             ' clent is ALLWAYS sent
            ' ignore ents without visible models
            If ev(ent).modelindex = 0 Then GoTo nexte
            If PR_StringEmpty(ev(ent).model) Then GoTo nexte

            For i = 0 To edicts(ent).num_leafs - 1
                l = edicts(ent).leafnums(i)
                If fatpvs(l \ 8) And (2 ^ (l And 7)) Then Exit For
            Next

            If i = edicts(ent).num_leafs Then GoTo nexte   ' not visible
        End If

        If msg.maxsize - msg.cursize < 16 Then
            Con_Printf "packet overflow\n"
            Exit Sub
        End If

        ' send an update
        bits = 0

        For i = 0 To 2
            miss = ev(ent).origin.v(i) - edicts(ent).baseline.origin.v(i)
            If miss < -0.1 Or miss > 0.1 Then bits = bits Or (U_ORIGIN1 * (2 ^ i))
        Next

        If ev(ent).angles.v(0) <> edicts(ent).baseline.angles.v(0) Then bits = bits Or U_ANGLE1

        If ev(ent).angles.v(1) <> edicts(ent).baseline.angles.v(1) Then bits = bits Or U_ANGLE2

        If ev(ent).angles.v(2) <> edicts(ent).baseline.angles.v(2) Then bits = bits Or U_ANGLE3

        If ev(ent).movetype = MOVETYPE_STEP Then bits = bits Or U_NOLERP   ' don't mess up the step animation

        If edicts(ent).baseline.colormap <> ev(ent).colormap Then bits = bits Or U_COLORMAP

        If edicts(ent).baseline.skin <> ev(ent).skin Then bits = bits Or U_SKIN

        If edicts(ent).baseline.frame <> ev(ent).frame Then bits = bits Or U_FRAME

        If edicts(ent).baseline.effects <> ev(ent).effects Then bits = bits Or U_EFFECTS

        If edicts(ent).baseline.modelindex <> ev(ent).modelindex Then bits = bits Or U_MODEL

        If e >= 256 Then bits = bits Or U_LONGENTITY

        If bits >= 256 Then bits = bits Or U_MOREBITS

        '
        ' write the message
        '
        MSG_WriteByte msg, bits Or U_SIGNAL

        If bits And U_MOREBITS Then MSG_WriteByte msg, bits \ 256
        If bits And U_LONGENTITY Then
            MSG_WriteShort msg, e
        Else
            MSG_WriteByte msg, e
        End If

        If bits And U_MODEL Then MSG_WriteByte msg, ev(ent).modelindex
        If bits And U_FRAME Then MSG_WriteByte msg, ev(ent).frame
        If bits And U_COLORMAP Then MSG_WriteByte msg, ev(ent).colormap
        If bits And U_SKIN Then MSG_WriteByte msg, ev(ent).skin
        If bits And U_EFFECTS Then MSG_WriteByte msg, ev(ent).effects
        If bits And U_ORIGIN1 Then MSG_WriteCoord msg, ev(ent).origin.v(0)
        If bits And U_ANGLE1 Then MSG_WriteAngle msg, ev(ent).angles.v(0)
        If bits And U_ORIGIN2 Then MSG_WriteCoord msg, ev(ent).origin.v(1)
        If bits And U_ANGLE2 Then MSG_WriteAngle msg, ev(ent).angles.v(1)
        If bits And U_ORIGIN3 Then MSG_WriteCoord msg, ev(ent).origin.v(2)
        If bits And U_ANGLE3 Then MSG_WriteAngle msg, ev(ent).angles.v(2)
nexte:
    Next
End Sub

Private Sub SV_CleanupEnts()
    Dim e As Long
    For e = 1 To sv.num_edicts - 1
        ev(e).effects = ToInt(ev(e).effects) And Not EF_MUZZLEFLASH
    Next
End Sub

Public Sub SV_WriteClientdataToMessage(ByVal ent As Long, msg As sizebuf_t)
    Dim bits As Long, i As Long, other As Long, items As Long, val As Long

    '
    ' send a damage message
    '
    If ev(ent).dmg_take <> 0 Or ev(ent).dmg_save <> 0 Then
        other = PROG_TO_EDICT(ev(ent).dmg_inflictor)
        MSG_WriteByte msg, svc_damage
        MSG_WriteByte msg, ev(ent).dmg_save
        MSG_WriteByte msg, ev(ent).dmg_take
        For i = 0 To 2
            MSG_WriteCoord msg, ev(other).origin.v(i) + 0.5 * (ev(other).mins.v(i) + ev(other).maxs.v(i))
        Next

        ev(ent).dmg_take = 0
        ev(ent).dmg_save = 0
    End If

    '
    ' send the current viewpos offset from the view entity
    '
    SV_SetIdealPitch                     ' how much to look up / down ideally

    ' a fixangle might get lost in a dropped packet.  Oh well.
    If ev(ent).fixangle <> 0 Then
        MSG_WriteByte msg, svc_setangle
        For i = 0 To 2
            MSG_WriteAngle msg, ev(ent).angles.v(i)
        Next
        ev(ent).fixangle = 0
    End If

    bits = 0

    If ev(ent).view_ofs.v(2) <> DEFAULT_VIEWHEIGHT Then bits = bits Or SU_VIEWHEIGHT

    If ev(ent).idealpitch <> 0 Then bits = bits Or SU_IDEALPITCH

    ' stuff the sigil bits into the high bits of items for sbar, or else
    ' mix in items2
    val = GetEdictFieldValue(ent, "items2")

    If val <> -1 Then
        items = ToInt(ev(ent).items) Or (ToInt(ed_f(val)) * &H800000)
    Else
        items = ToInt(ev(ent).items) Or (ToInt(pr_global_struct(0).serverflags) * &H10000000)
    End If

    bits = bits Or SU_ITEMS

    If ToInt(ev(ent).flags) And FL_ONGROUND Then bits = bits Or SU_ONGROUND

    If ev(ent).waterlevel >= 2 Then bits = bits Or SU_INWATER

    For i = 0 To 2
        If ev(ent).punchangle.v(i) <> 0 Then bits = bits Or (SU_PUNCH1 * (2 ^ i))
        If ev(ent).velocity.v(i) <> 0 Then bits = bits Or (SU_VELOCITY1 * (2 ^ i))
    Next

    If ev(ent).weaponframe <> 0 Then bits = bits Or SU_WEAPONFRAME

    If ev(ent).armorvalue <> 0 Then bits = bits Or SU_ARMOR

    ' if (ent->v.weapon)
    bits = bits Or SU_WEAPON

    ' send the data

    MSG_WriteByte msg, svc_clientdata
    MSG_WriteShort msg, bits

    If bits And SU_VIEWHEIGHT Then MSG_WriteChar msg, ev(ent).view_ofs.v(2)

    If bits And SU_IDEALPITCH Then MSG_WriteChar msg, ev(ent).idealpitch

    For i = 0 To 2
        If bits And (SU_PUNCH1 * (2 ^ i)) Then MSG_WriteChar msg, ev(ent).punchangle.v(i)
        If bits And (SU_VELOCITY1 * (2 ^ i)) Then MSG_WriteChar msg, ev(ent).velocity.v(i) / 16
    Next

    ' [always sent]	if (bits & SU_ITEMS)
    MSG_WriteLong msg, items

    If bits And SU_WEAPONFRAME Then MSG_WriteByte msg, ev(ent).weaponframe
    If bits And SU_ARMOR Then MSG_WriteByte msg, ev(ent).armorvalue
    If bits And SU_WEAPON Then MSG_WriteByte msg, SV_ModelIndex(PR_GetString(ev(ent).weaponmodel))

    MSG_WriteShort msg, ev(ent).health
    MSG_WriteByte msg, ev(ent).currentammo
    MSG_WriteByte msg, ev(ent).ammo_shells
    MSG_WriteByte msg, ev(ent).ammo_nails
    MSG_WriteByte msg, ev(ent).ammo_rockets
    MSG_WriteByte msg, ev(ent).ammo_cells

    If standard_quake Then
        MSG_WriteByte msg, ev(ent).weapon
    Else
        For i = 0 To 31
            If ToInt(ev(ent).weapon) And (2 ^ i) Then
                MSG_WriteByte msg, i
                Exit For
            End If
        Next
    End If
End Sub

Private Function SV_SendClientDatagram(ByVal client As Long) As Boolean
    Dim msg As sizebuf_t

    SZ_Alloc msg, MAX_DATAGRAM

    MSG_WriteByte msg, svc_time
    MSG_WriteFloat msg, sv.time

    ' add the client specific data to the datagram
    SV_WriteClientdataToMessage svs.clients(client).edict, msg

    SV_WriteEntitiesToClient svs.clients(client).edict, msg

    ' copy the server datagram if there is space
    If msg.cursize + sv.datagram.cursize < msg.maxsize Then SZ_Write msg, sv.datagram.data, 0, sv.datagram.cursize

    ' send the datagram
    If NET_SendUnreliableMessage(svs.clients(client).netconnection, msg) = -1 Then
        SV_DropClient True               ' if the message couldn't send, kick off
        SV_SendClientDatagram = False
        Exit Function
    End If

    SV_SendClientDatagram = True
End Function

Private Sub SV_UpdateToReliableMessages()
    Dim i As Long, j As Long

    ' check for changes to be sent over the reliable streams
    For i = 0 To svs.maxclients - 1
        host_client = i
        If svs.clients(host_client).old_frags <> ev(svs.clients(host_client).edict).frags Then
            For j = 0 To svs.maxclients - 1
                If svs.clients(j).active Then
                    MSG_WriteByte svs.clients(j).message, svc_updatefrags
                    MSG_WriteByte svs.clients(j).message, i
                    MSG_WriteShort svs.clients(j).message, ev(svs.clients(host_client).edict).frags
                End If
            Next

            svs.clients(host_client).old_frags = ToInt(ev(svs.clients(host_client).edict).frags)
        End If
    Next

    For j = 0 To svs.maxclients - 1
        If svs.clients(j).active Then
            SZ_Write svs.clients(j).message, sv.reliable_datagram.data, 0, sv.reliable_datagram.cursize
        End If
    Next

    SZ_Clear sv.reliable_datagram
End Sub

' Send a nop message without trashing or sending the accumulated client
' message buffer
Private Sub SV_SendNop(ByVal client As Long)
    Dim msg As sizebuf_t

    SZ_Alloc msg, 4

    MSG_WriteChar msg, svc_nop

    If NET_SendUnreliableMessage(svs.clients(client).netconnection, msg) = -1 Then SV_DropClient True   ' if the message couldn't send, kick off
    svs.clients(client).last_message = realtime
End Sub

Public Sub SV_SendClientMessages()
    Dim i As Long

    ' update frags, names, etc
    SV_UpdateToReliableMessages

    ' build individual updates
    For i = 0 To svs.maxclients - 1
        host_client = i

        If Not svs.clients(i).active Then GoTo nexti

        If svs.clients(i).spawned Then
            If Not SV_SendClientDatagram(i) Then GoTo nexti
        Else
            ' the player isn't totally in the game yet
            ' send small keepalive messages if too much time has passed
            ' send a full message when the next signon stage has been requested
            ' some other message data (name changes, etc) may accumulate
            ' between signon stages
            If Not svs.clients(i).sendsignon Then
                If realtime - svs.clients(i).last_message > 5 Then SV_SendNop i
                GoTo nexti                ' don't send out non-signon messages
            End If
        End If

        ' check for an overflowed message.  Should only happen
        ' on a very fucked up connection that backs up a lot, then
        ' changes level
        If svs.clients(i).message.overflowed Then
            SV_DropClient True
            svs.clients(i).message.overflowed = False
            GoTo nexti
        End If

        If svs.clients(i).message.cursize <> 0 Or svs.clients(i).dropasap Then
            If Not NET_CanSendMessage(svs.clients(i).netconnection) Then GoTo nexti

            If svs.clients(i).dropasap Then
                SV_DropClient False      ' went to another level
            Else
                If NET_SendMessage(svs.clients(i).netconnection, svs.clients(i).message) = -1 Then SV_DropClient True   ' if the message couldn't send, kick off
                SZ_Clear svs.clients(i).message
                svs.clients(i).last_message = realtime
                svs.clients(i).sendsignon = False
            End If
        End If
nexti:
    Next

    ' clear muzzle flashes
    SV_CleanupEnts
End Sub

'==============================================================================
' SERVER SPAWNING

Public Function SV_ModelIndex(ByVal name As String) As Long
    Dim i As Long

    If Len(name) = 0 Then
        SV_ModelIndex = 0
        Exit Function
    End If

    For i = 0 To sv.num_model_precache - 1
        If sv.model_precache(i) = name Then
            SV_ModelIndex = i
            Exit Function
        End If
    Next
    Sys_Error "SV_ModelIndex: model %s not precached", name
End Function

Private Sub SV_CreateBaseline()
    Dim i As Long, svent As Long, entnum As Long

    For entnum = 0 To sv.num_edicts - 1
        ' get the current server version
        svent = EDICT_NUM(entnum)
        If edicts(svent).free Then GoTo nextent
        If entnum > svs.maxclients And ev(svent).modelindex = 0 Then GoTo nextent

        '
        ' create entity baseline
        '
        edicts(svent).baseline.origin = ev(svent).origin
        edicts(svent).baseline.angles = ev(svent).angles
        edicts(svent).baseline.frame = ToInt(ev(svent).frame)
        edicts(svent).baseline.skin = ToInt(ev(svent).skin)
        If entnum > 0 And entnum <= svs.maxclients Then
            edicts(svent).baseline.colormap = entnum
            edicts(svent).baseline.modelindex = SV_ModelIndex("progs/player.mdl")
        Else
            edicts(svent).baseline.colormap = 0
            edicts(svent).baseline.modelindex = SV_ModelIndex(PR_GetString(ev(svent).model))
        End If

        '
        ' add to the message
        '
        MSG_WriteByte sv.signon, svc_spawnbaseline
        MSG_WriteShort sv.signon, entnum

        MSG_WriteByte sv.signon, edicts(svent).baseline.modelindex
        MSG_WriteByte sv.signon, edicts(svent).baseline.frame
        MSG_WriteByte sv.signon, edicts(svent).baseline.colormap
        MSG_WriteByte sv.signon, edicts(svent).baseline.skin
        For i = 0 To 2
            MSG_WriteCoord sv.signon, edicts(svent).baseline.origin.v(i)
            MSG_WriteAngle sv.signon, edicts(svent).baseline.angles.v(i)
        Next
nextent:
    Next
End Sub

' Tell all the clients that the server is changing levels
Private Sub SV_SendReconnect()
    Dim msg As sizebuf_t

    SZ_Alloc msg, 128

    MSG_WriteChar msg, svc_stufftext
    MSG_WriteString msg, "reconnect" & vbLf
    NET_SendToAll msg, 5

    If cls.state <> ca_dedicated Then Cmd_ExecuteString "reconnect" & vbLf, src_command
End Sub

' Grabs the current state of each client for saving across the
' transition to another level
Public Sub SV_SaveSpawnparms()
    Dim i As Long, j As Long

    svs.serverflags = ToInt(pr_global_struct(0).serverflags)

    For i = 0 To svs.maxclients - 1
        host_client = i
        If svs.clients(i).active Then
            ' call the progs to get default spawn parms for the new client
            pr_global_struct(0).self = EDICT_TO_PROG(svs.clients(i).edict)
            PR_ExecuteProgram pr_global_struct(0).SetChangeParms
            For j = 0 To NUM_SPAWN_PARMS - 1
                svs.clients(i).spawn_parms(j) = pr_gf(GOFS_parm1 + j)
            Next
        End If
    Next
End Sub

' This is called at the start of each level
Public Sub SV_SpawnServer(ByVal server As String)
    Dim ent As Long, i As Long, blank As server_t

    ' let's not have any servers with no name
    If Len(cvars(cv_hostname).strval) = 0 Then Cvar_Set "hostname", "UNNAMED"
    scr_centertime_off = 0

    Con_DPrintf "SpawnServer: %s\n", server
    svs.changelevel_issued = False       ' now safe to issue another

    '
    ' tell all connected clients that we are going to a new level
    '
    If sv.active Then SV_SendReconnect

    '
    ' make cvars consistant
    '
    If cvars(cv_coop).value <> 0 Then Cvar_SetValue "deathmatch", 0
    current_skill = ToInt(cvars(cv_skill).value + 0.5)
    If current_skill < 0 Then current_skill = 0
    If current_skill > 3 Then current_skill = 3

    Cvar_SetValue "skill", current_skill

    '
    ' set up the new server
    '
    Host_ClearMemory

    sv = blank

    sv.name = server

    ' load progs to get entity field count
    PR_LoadProgs

    ' allocate server memory
    sv.max_edicts = MAX_EDICTS

    PR_AllocEdicts sv.max_edicts

    SZ_Alloc sv.datagram, MAX_DATAGRAM
    SZ_Alloc sv.reliable_datagram, MAX_DATAGRAM
    SZ_Alloc sv.signon, 8192

    ' leave slots at start for clients only
    sv.num_edicts = svs.maxclients + 1
    For i = 0 To svs.maxclients - 1
        ent = EDICT_NUM(i + 1)
        svs.clients(i).edict = ent
    Next

    sv.state = ss_loading
    sv.paused = False

    sv.time = 1#

    sv.name = server
    sv.modelname = va("maps/%s.bsp", server)
    sv.worldmodel = Mod_ForName(sv.modelname, False)
    If sv.worldmodel = 0 Then
        Con_Printf "Couldn't spawn server %s\n", sv.modelname
        sv.active = False
        Exit Sub
    End If
    sv.models(1) = sv.worldmodel

    '
    ' clear world interaction links
    '
    SV_ClearWorld

    sv.sound_precache(0) = ""            ' C: pr_strings
    sv.num_sound_precache = 1

    sv.model_precache(0) = ""
    sv.model_precache(1) = sv.modelname
    For i = 1 To mod_known(sv.worldmodel).numsubmodels - 1
        sv.model_precache(1 + i) = localmodels(i)
        sv.models(i + 1) = Mod_ForName(localmodels(i), False)
    Next
    sv.num_model_precache = mod_known(sv.worldmodel).numsubmodels + 1

    '
    ' load the rest of the entities
    '
    ent = EDICT_NUM(0)
    FillMemory ed_l(0), progs_entityfields * 4, 0
    edicts(ent).free = False
    PR_SetEngineString PRS_WORLDMODEL, 64, mod_known(sv.worldmodel).name
    ev(ent).model = PRS_WORLDMODEL
    ev(ent).modelindex = 1               ' world model
    ev(ent).solid = SOLID_BSP
    ev(ent).movetype = MOVETYPE_PUSH

    If cvars(cv_coop).value <> 0 Then
        pr_global_struct(0).coop = cvars(cv_coop).value
    Else
        pr_global_struct(0).deathmatch = cvars(cv_deathmatch).value
    End If

    PR_SetEngineString PRS_MAPNAME, 64, sv.name
    pr_global_struct(0).mapname = PRS_MAPNAME

    ' serverflags are for cross level information (sigils)
    pr_global_struct(0).serverflags = svs.serverflags

    ED_LoadFromFile mod_entdata, mod_known(sv.worldmodel).entities

    sv.active = True

    ' all setup is completed, any further precache statements are errors
    sv.state = ss_active

    ' run two frames to allow everything to settle
    host_frametime = 0.1
    SV_Physics
    SV_Physics

    ' create a baseline for more efficient communications
    SV_CreateBaseline

    ' send serverinfo to all connected clients
    For i = 0 To svs.maxclients - 1
        host_client = i
        If svs.clients(i).active Then SV_SendServerinfo i
    Next

    Con_DPrintf "Server spawned.\n"
End Sub
