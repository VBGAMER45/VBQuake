Attribute VB_Name = "pr_cmds"
' pr_cmds.bas - the QuakeC built-in functions (pr_cmds.c). PR_CallBuiltin
' is the pr_builtin[] table. Only the functions of the non-QUAKE2 build are
' ported (#60 .. #66 are PF_Fixme there).
'
' A builtin's parameters are the globals at OFS_PARM0 + 3 * n: floats in
' pr_gf(), the rest in pr_gl(), vectors as pr_global_struct(0).parm(n).
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public Const pr_numbuiltins As Long = 79

Public cv_sv_aim As Long

Private checkpvs(0 To MAX_MAP_LEAFS \ 8 - 1) As Byte

Private c_invis As Long
Private c_notvis As Long

Private Const MSG_BROADCAST As Long = 0  ' unreliable to all
Private Const MSG_ONE As Long = 1        ' reliable to one (msg_entity)
Private Const MSG_ALL As Long = 2        ' reliable to all
Private Const MSG_INIT As Long = 3       ' write to the init string

' what PF_Write* write (for WriteToDest)
Private Const WR_BYTE As Long = 0
Private Const WR_CHAR As Long = 1
Private Const WR_SHORT As Long = 2
Private Const WR_LONG As Long = 3
Private Const WR_ANGLE As Long = 4
Private Const WR_COORD As Long = 5
Private Const WR_STRING As Long = 6
Private Const WR_ENTITY As Long = 7

'===============================================================================
' BUILT-IN FUNCTIONS

' C: qboolean as a float
Private Function QB(ByVal b As Boolean) As Single
    If b Then QB = 1 Else QB = 0
End Function

' C: RETURN_EDICT(e)
Private Sub RETURN_EDICT(ByVal e As Long)
    pr_gl(OFS_RETURN) = EDICT_TO_PROG(e)
End Sub

Private Function PF_VarString(ByVal first As Long) As String
    Dim i As Long, out As String
    For i = first To pr_argc - 1
        out = out & G_STRING(OFS_PARM0 + i * 3)
    Next
    PF_VarString = out
End Function

' This is a TERMINAL error, which will kill off the entire server.
' Dumps self.
'
' error(value)
Private Sub PF_error()
    Dim s As String, ed As Long

    s = PF_VarString(0)
    Con_Printf "======SERVER ERROR in %s:\n%s\n", PR_GetString(pr_functions(pr_xfunction).s_name), s
    ed = PROG_TO_EDICT(pr_global_struct(0).self)
    ED_Print ed

    Host_Error "Program error"
End Sub

' Dumps out self, then an error message.  The program is aborted and self is
' removed, but the level can continue.
'
' objerror(value)
Private Sub PF_objerror()
    Dim s As String, ed As Long

    s = PF_VarString(0)
    Con_Printf "======OBJECT ERROR in %s:\n%s\n", PR_GetString(pr_functions(pr_xfunction).s_name), s
    ed = PROG_TO_EDICT(pr_global_struct(0).self)
    ED_Print ed
    ED_Free ed

    Host_Error "Program error"
End Sub

' Writes new values for v_forward, v_up, and v_right based on angles
' makevectors(vector)
Private Sub PF_makevectors()
    AngleVectors pr_global_struct(0).parm(0), pr_global_struct(0).v_forward, pr_global_struct(0).v_right, pr_global_struct(0).v_up
End Sub

' This is the only valid way to move an object without using the physics of
' the world (setting velocity and waiting).  Directly changing origin will
' not set internal links correctly, so clipping would be messed up.  This
' should be called when an object is spawned, and then only if it is
' teleported.
'
' setorigin (entity, origin)
Private Sub PF_setorigin()
    Dim e As Long

    e = G_EDICT(OFS_PARM0)
    ev(e).origin = pr_global_struct(0).parm(1)
    SV_LinkEdict e, False
End Sub

Private Sub SetMinMaxSize(ByVal e As Long, min As vec3_t, max As vec3_t, ByVal rotate As Boolean)
    Dim rmin As vec3_t, rmax As vec3_t, i As Long, minc As vec3_t, maxc As vec3_t

    For i = 0 To 2
        If min.v(i) > max.v(i) Then PR_RunError "backwards mins/maxs"
    Next

    rotate = False                       ' FIXME: implement rotation properly again
    ' (only the !rotate case is ever used)
    minc = min
    maxc = max
    rmin = minc
    rmax = maxc

    ' set derived values
    ev(e).mins = rmin
    ev(e).maxs = rmax
    VectorSubtract maxc, minc, ev(e).size

    SV_LinkEdict e, False
End Sub

' the size box is rotated by the current angle
'
' setsize (entity, minvector, maxvector)
Private Sub PF_setsize()
    Dim e As Long
    e = G_EDICT(OFS_PARM0)
    SetMinMaxSize e, pr_global_struct(0).parm(1), pr_global_struct(0).parm(2), False
End Sub

' setmodel(entity, model)
Private Sub PF_setmodel()
    Dim e As Long, m As String, i As Long, md As Long

    e = G_EDICT(OFS_PARM0)
    m = G_STRING(OFS_PARM1)

    ' check to see if model was properly precached
    For i = 0 To sv.num_model_precache - 1
        If sv.model_precache(i) = m Then Exit For
    Next

    If i = sv.num_model_precache Then PR_RunError "no precache: %s\n", m

    ev(e).model = pr_gl(OFS_PARM1)
    ev(e).modelindex = i                 ' SV_ModelIndex (m);

    md = sv.models(ToInt(ev(e).modelindex))  ' Mod_ForName (m, true);

    If md Then
        SetMinMaxSize e, mod_known(md).mins, mod_known(md).maxs, True
    Else
        SetMinMaxSize e, vec3_origin, vec3_origin, True
    End If
End Sub

' broadcast print to everyone on server
' bprint(value)
Private Sub PF_bprint()
    SV_BroadcastPrintf "%s", PF_VarString(0)
End Sub

' single print to a specific client
' sprint(clientent, value)
Private Sub PF_sprint()
    Dim s As String, entnum As Long

    entnum = G_EDICTNUM(OFS_PARM0)
    s = PF_VarString(1)

    If entnum < 1 Or entnum > svs.maxclients Then
        Con_Printf "tried to sprint to a non-client\n"
        Exit Sub
    End If

    MSG_WriteChar svs.clients(entnum - 1).message, svc_print
    MSG_WriteString svs.clients(entnum - 1).message, s
End Sub

' single print to a specific client
' centerprint(clientent, value)
Private Sub PF_centerprint()
    Dim s As String, entnum As Long

    entnum = G_EDICTNUM(OFS_PARM0)
    s = PF_VarString(1)

    If entnum < 1 Or entnum > svs.maxclients Then
        Con_Printf "tried to sprint to a non-client\n"
        Exit Sub
    End If

    MSG_WriteChar svs.clients(entnum - 1).message, svc_centerprint
    MSG_WriteString svs.clients(entnum - 1).message, s
End Sub

' vector normalize(vector)
Private Sub PF_normalize()
    Dim value1 As vec3_t, newvalue As vec3_t, nw As Single

    value1 = pr_global_struct(0).parm(0)

    nw = value1.v(0) * value1.v(0) + value1.v(1) * value1.v(1) + value1.v(2) * value1.v(2)
    nw = Sqr(nw)

    If nw <> 0 Then
        nw = 1 / nw
        newvalue.v(0) = value1.v(0) * nw
        newvalue.v(1) = value1.v(1) * nw
        newvalue.v(2) = value1.v(2) * nw
    End If

    pr_global_struct(0).ret = newvalue
End Sub

' scalar vlen(vector)
Private Sub PF_vlen()
    Dim value1 As vec3_t, nw As Single

    value1 = pr_global_struct(0).parm(0)

    nw = value1.v(0) * value1.v(0) + value1.v(1) * value1.v(1) + value1.v(2) * value1.v(2)
    nw = Sqr(nw)

    pr_gf(OFS_RETURN) = nw
End Sub

' float vectoyaw(vector)
Private Sub PF_vectoyaw()
    Dim value1 As vec3_t, yaw_ As Single

    value1 = pr_global_struct(0).parm(0)

    If value1.v(1) = 0 And value1.v(0) = 0 Then
        yaw_ = 0
    Else
        yaw_ = ToInt(atan2(value1.v(1), value1.v(0)) * 180 / M_PI)
        If yaw_ < 0 Then yaw_ = yaw_ + 360
    End If

    pr_gf(OFS_RETURN) = yaw_
End Sub

' vector vectoangles(vector)
Private Sub PF_vectoangles()
    Dim value1 As vec3_t, forward As Single, yaw_ As Single, pitch_ As Single

    value1 = pr_global_struct(0).parm(0)

    If value1.v(1) = 0 And value1.v(0) = 0 Then
        yaw_ = 0
        If value1.v(2) > 0 Then
            pitch_ = 90
        Else
            pitch_ = 270
        End If
    Else
        yaw_ = ToInt(atan2(value1.v(1), value1.v(0)) * 180 / M_PI)
        If yaw_ < 0 Then yaw_ = yaw_ + 360

        forward = Sqr(value1.v(0) * value1.v(0) + value1.v(1) * value1.v(1))
        pitch_ = ToInt(atan2(value1.v(2), forward) * 180 / M_PI)
        If pitch_ < 0 Then pitch_ = pitch_ + 360
    End If

    pr_gf(OFS_RETURN + 0) = pitch_
    pr_gf(OFS_RETURN + 1) = yaw_
    pr_gf(OFS_RETURN + 2) = 0
End Sub

' Returns a number from 0<= num < 1
' random()
Private Sub PF_random()
    Dim num As Single
    num = (rand() And &H7FFF) / CSng(&H7FFF)
    pr_gf(OFS_RETURN) = num
End Sub

' particle(origin, color, count)
Private Sub PF_particle()
    Dim color As Single, count As Single

    color = pr_gf(OFS_PARM2)
    count = pr_gf(OFS_PARM3)
    SV_StartParticle pr_global_struct(0).parm(0), pr_global_struct(0).parm(1), ToInt(color), ToInt(count)
End Sub

Private Sub PF_ambientsound()
    Dim samp As String, vol As Single, attenuation As Single, i As Long, soundnum As Long, pos As vec3_t

    pos = pr_global_struct(0).parm(0)
    samp = G_STRING(OFS_PARM1)
    vol = pr_gf(OFS_PARM2)
    attenuation = pr_gf(OFS_PARM3)

    ' check to see if samp was properly precached
    For soundnum = 0 To sv.num_sound_precache - 1
        If sv.sound_precache(soundnum) = samp Then Exit For
    Next

    If soundnum = sv.num_sound_precache Then
        Con_Printf "no precache: %s\n", samp
        Exit Sub
    End If

    ' add an svc_spawnambient command to the level signon packet
    MSG_WriteByte sv.signon, svc_spawnstaticsound
    For i = 0 To 2
        MSG_WriteCoord sv.signon, pos.v(i)
    Next

    MSG_WriteByte sv.signon, soundnum

    MSG_WriteByte sv.signon, vol * 255
    MSG_WriteByte sv.signon, attenuation * 64
End Sub

' Each entity can have eight independant sound sources, like voice,
' weapon, feet, etc.
'
' Channel 0 is an auto-allocate channel, the others override anything
' allready running on that entity/channel pair.
'
' An attenuation of 0 will play full volume everywhere in the level.
' Larger attenuations will drop off.
Private Sub PF_sound()
    Dim sample As String, channel As Long, entity As Long, volume As Long, attenuation As Single

    entity = G_EDICT(OFS_PARM0)
    channel = ToInt(pr_gf(OFS_PARM1))
    sample = G_STRING(OFS_PARM2)
    volume = ToInt(pr_gf(OFS_PARM3) * 255)
    attenuation = pr_gf(OFS_PARM4)

    If volume < 0 Or volume > 255 Then Sys_Error "SV_StartSound: volume = %i", volume

    If attenuation < 0 Or attenuation > 4 Then Sys_Error "SV_StartSound: attenuation = %f", attenuation

    If channel < 0 Or channel > 7 Then Sys_Error "SV_StartSound: channel = %i", channel

    SV_StartSound entity, channel, sample, volume, attenuation
End Sub

' break()
Private Sub PF_break()
    Con_Printf "break statement\n"
    ' (C writes to address -4 to drop into the debugger)
    Sys_Error "break statement"
End Sub

' Used for use tracing and shot targeting
' Traces are blocked by bbox and exact bsp entityes, and also slide box entities
' if the tryents flag is set.
'
' traceline (vector1, vector2, tryents)
Private Sub PF_traceline()
    Dim trace As trace_t, nomonsters As Long, ent As Long, v1 As vec3_t, v2 As vec3_t

    v1 = pr_global_struct(0).parm(0)
    v2 = pr_global_struct(0).parm(1)
    nomonsters = ToInt(pr_gf(OFS_PARM2))
    ent = G_EDICT(OFS_PARM3)

    trace = SV_Move(v1, vec3_origin, vec3_origin, v2, nomonsters, ent)

    With pr_global_struct(0)
        .trace_allsolid = QB(trace.allsolid)
        .trace_startsolid = QB(trace.startsolid)
        .trace_fraction = trace.fraction
        .trace_inwater = QB(trace.inwater)
        .trace_inopen = QB(trace.inopen)
        .trace_endpos = trace.endpos
        .trace_plane_normal = trace.plane.normal
        .trace_plane_dist = trace.plane.dist
        If trace.ent <> -1 Then
            .trace_ent = EDICT_TO_PROG(trace.ent)
        Else
            .trace_ent = EDICT_TO_PROG(0)
        End If
    End With
End Sub

' Returns true if the given entity can move to the given position from it's
' current position by walking or rolling.
' FIXME: make work...
' scalar checkpos (entity, vector)
Private Sub PF_checkpos()
End Sub

'============================================================================

Private Function PF_newcheckclient(ByVal check As Long) As Long
    Dim i As Long, ent As Long, leaf As Long, org As vec3_t

    ' cycle to the next one

    If check < 1 Then check = 1
    If check > svs.maxclients Then check = svs.maxclients

    If check = svs.maxclients Then
        i = 1
    Else
        i = check + 1
    End If

    Do
        If i = svs.maxclients + 1 Then i = 1

        ent = EDICT_NUM(i)

        If i = check Then Exit Do        ' didn't find anything else

        If Not edicts(ent).free Then
            If ev(ent).health > 0 Then
                If (ToInt(ev(ent).flags) And FL_NOTARGET) = 0 Then Exit Do   ' anything that is a client, or has a client as an enemy
            End If
        End If
        i = i + 1
    Loop

    ' get the PVS for the entity
    VectorAdd ev(ent).origin, ev(ent).view_ofs, org
    leaf = Mod_PointInLeaf(org, sv.worldmodel)
    Mod_LeafPVS leaf, sv.worldmodel
    CopyMemory checkpvs(0), mod_pvs(0), (mod_known(sv.worldmodel).numleafs + 7) \ 8

    PF_newcheckclient = i
End Function

' Returns a client (or object that has a client enemy) that would be a
' valid target.
'
' If there are more than one valid options, they are cycled each frame
'
' If (self.origin + self.viewofs) is not in the PVS of the current target,
' it is not returned at all.
'
' name checkclient ()
Private Sub PF_checkclient()
    Dim ent As Long, self As Long, leaf As Long, l As Long, view As vec3_t

    ' find a new check if on a new frame
    If sv.time - sv.lastchecktime >= 0.1 Then
        sv.lastcheck = PF_newcheckclient(sv.lastcheck)
        sv.lastchecktime = sv.time
    End If

    ' return check if it might be visible
    ent = EDICT_NUM(sv.lastcheck)
    If edicts(ent).free Or ev(ent).health <= 0 Then
        RETURN_EDICT 0
        Exit Sub
    End If

    ' if current entity can't possibly see the check entity, return 0
    self = PROG_TO_EDICT(pr_global_struct(0).self)
    VectorAdd ev(self).origin, ev(self).view_ofs, view
    leaf = Mod_PointInLeaf(view, sv.worldmodel)
    l = (leaf - mod_known(sv.worldmodel).leafs) - 1
    If l < 0 Then
        c_notvis = c_notvis + 1
        RETURN_EDICT 0
        Exit Sub
    End If
    If (checkpvs(l \ 8) And (2 ^ (l And 7))) = 0 Then
        c_notvis = c_notvis + 1
        RETURN_EDICT 0
        Exit Sub
    End If

    ' might be able to see it
    c_invis = c_invis + 1
    RETURN_EDICT ent
End Sub

'============================================================================

' Sends text over to the client's execution buffer
'
' stuffcmd (clientent, value)
Private Sub PF_stuffcmd()
    Dim entnum As Long, str_ As String, old As Long

    entnum = G_EDICTNUM(OFS_PARM0)
    If entnum < 1 Or entnum > svs.maxclients Then PR_RunError "Parm 0 not a client"
    str_ = G_STRING(OFS_PARM1)

    old = host_client
    host_client = entnum - 1
    Host_ClientCommands "%s", str_
    host_client = old
End Sub

' Sends text over to the client's execution buffer
'
' localcmd (string)
Private Sub PF_localcmd()
    Cbuf_AddText G_STRING(OFS_PARM0)
End Sub

' float cvar (string)
Private Sub PF_cvar()
    pr_gf(OFS_RETURN) = Cvar_VariableValue(G_STRING(OFS_PARM0))
End Sub

' float cvar (string)
Private Sub PF_cvar_set()
    Cvar_Set G_STRING(OFS_PARM0), G_STRING(OFS_PARM1)
End Sub

' Returns a chain of entities that have origins within a spherical area
'
' findradius (origin, radius)
Private Sub PF_findradius()
    Dim ent As Long, chain As Long, rad As Single, org As vec3_t, eorg As vec3_t, i As Long, j As Long

    chain = 0

    org = pr_global_struct(0).parm(0)
    rad = pr_gf(OFS_PARM1)

    For i = 1 To sv.num_edicts - 1
        ent = i
        If edicts(ent).free Then GoTo nexti
        If ev(ent).solid = SOLID_NOT Then GoTo nexti
        For j = 0 To 2
            eorg.v(j) = org.v(j) - (ev(ent).origin.v(j) + (ev(ent).mins.v(j) + ev(ent).maxs.v(j)) * 0.5)
        Next
        If Length(eorg) > rad Then GoTo nexti

        ev(ent).chain = EDICT_TO_PROG(chain)
        chain = ent
nexti:
    Next

    RETURN_EDICT chain
End Sub

Private Sub PF_dprint()
    Con_DPrintf "%s", PF_VarString(0)
End Sub

Private Sub PF_ftos()
    Dim v As Single

    v = pr_gf(OFS_PARM0)

    If v = ToInt(v) Then
        PR_SetEngineString PRS_TEMP, 128, va("%d", ToInt(v))
    Else
        PR_SetEngineString PRS_TEMP, 128, va("%5.1f", v)
    End If
    pr_gl(OFS_RETURN) = PRS_TEMP
End Sub

Private Sub PF_fabs()
    pr_gf(OFS_RETURN) = Abs(pr_gf(OFS_PARM0))
End Sub

Private Sub PF_vtos()
    PR_SetEngineString PRS_TEMP, 128, va("'%5.1f %5.1f %5.1f'", pr_gf(OFS_PARM0), pr_gf(OFS_PARM0 + 1), pr_gf(OFS_PARM0 + 2))
    pr_gl(OFS_RETURN) = PRS_TEMP
End Sub

Private Sub PF_Spawn()
    RETURN_EDICT ED_Alloc()
End Sub

Private Sub PF_Remove()
    ED_Free G_EDICT(OFS_PARM0)
End Sub

' entity (entity start, .string field, string match) find = #5;
Private Sub PF_Find()
    Dim e As Long, f As Long, s As Long

    e = G_EDICTNUM(OFS_PARM0)
    f = pr_gl(OFS_PARM1)
    s = pr_gl(OFS_PARM2)

    For e = e + 1 To sv.num_edicts - 1
        If Not edicts(e).free Then
            If PR_StringsEqual(ed_l(e * EDICT_STRIDE + f), s) Then
                RETURN_EDICT e
                Exit Sub
            End If
        End If
    Next

    RETURN_EDICT 0
End Sub

Private Sub PR_CheckEmptyString(ByVal s As Long)
    ' (C compares a signed char with ' ': bytes 128..255 fail too)
    If pr_strings(s) <= 32 Or pr_strings(s) >= 128 Then PR_RunError "Bad string"
End Sub

Private Sub PF_precache_file()
    ' precache_file is only used to copy files with qcc, it does nothing
    pr_gl(OFS_RETURN) = pr_gl(OFS_PARM0)
End Sub

Private Sub PF_precache_sound()
    Dim s As String, i As Long

    If sv.state <> ss_loading Then PR_RunError "PF_Precache_*: Precache can only be done in spawn functions"

    s = G_STRING(OFS_PARM0)
    pr_gl(OFS_RETURN) = pr_gl(OFS_PARM0)
    PR_CheckEmptyString pr_gl(OFS_PARM0)

    For i = 0 To MAX_SOUNDS - 1
        If i >= sv.num_sound_precache Then
            sv.sound_precache(i) = s
            sv.num_sound_precache = i + 1
            Exit Sub
        End If
        If sv.sound_precache(i) = s Then Exit Sub
    Next
    PR_RunError "PF_precache_sound: overflow"
End Sub

Private Sub PF_precache_model()
    Dim s As String, i As Long

    If sv.state <> ss_loading Then PR_RunError "PF_Precache_*: Precache can only be done in spawn functions"

    s = G_STRING(OFS_PARM0)
    pr_gl(OFS_RETURN) = pr_gl(OFS_PARM0)
    PR_CheckEmptyString pr_gl(OFS_PARM0)

    For i = 0 To MAX_MODELS - 1
        If i >= sv.num_model_precache Then
            sv.model_precache(i) = s
            sv.num_model_precache = i + 1
            sv.models(i) = Mod_ForName(s, True)
            Exit Sub
        End If
        If sv.model_precache(i) = s Then Exit Sub
    Next
    PR_RunError "PF_precache_model: overflow"
End Sub

Private Sub PF_coredump()
    ED_PrintEdicts
End Sub

Private Sub PF_traceon()
    pr_trace = True
End Sub

Private Sub PF_traceoff()
    pr_trace = False
End Sub

Private Sub PF_eprint()
    ED_PrintNum G_EDICTNUM(OFS_PARM0)
End Sub

' float(float yaw, float dist) walkmove
Private Sub PF_walkmove()
    Dim ent As Long, yaw_ As Single, dist As Single, move As vec3_t, oldf As Long, oldself As Long

    ent = PROG_TO_EDICT(pr_global_struct(0).self)
    yaw_ = pr_gf(OFS_PARM0)
    dist = pr_gf(OFS_PARM1)

    If (ToInt(ev(ent).flags) And (FL_ONGROUND Or FL_FLY Or FL_SWIM)) = 0 Then
        pr_gf(OFS_RETURN) = 0
        Exit Sub
    End If

    yaw_ = yaw_ * M_PI * 2 / 360

    move.v(0) = Cos(yaw_) * dist
    move.v(1) = Sin(yaw_) * dist
    move.v(2) = 0

    ' save program state, because SV_movestep may call other progs
    oldf = pr_xfunction
    oldself = pr_global_struct(0).self

    pr_gf(OFS_RETURN) = QB(SV_movestep(ent, move, True))

    ' restore program state
    pr_xfunction = oldf
    pr_global_struct(0).self = oldself
End Sub

' void() droptofloor
Private Sub PF_droptofloor()
    Dim ent As Long, end_ As vec3_t, trace As trace_t

    ent = PROG_TO_EDICT(pr_global_struct(0).self)

    end_ = ev(ent).origin
    end_.v(2) = end_.v(2) - 256

    trace = SV_Move(ev(ent).origin, ev(ent).mins, ev(ent).maxs, end_, 0, ent)

    If trace.fraction = 1 Or trace.allsolid Then
        pr_gf(OFS_RETURN) = 0
    Else
        ev(ent).origin = trace.endpos
        SV_LinkEdict ent, False
        ev(ent).flags = ToInt(ev(ent).flags) Or FL_ONGROUND
        ev(ent).groundentity = EDICT_TO_PROG(trace.ent)
        pr_gf(OFS_RETURN) = 1
    End If
End Sub

' void(float style, string value) lightstyle
Private Sub PF_lightstyle()
    Dim style As Long, val As String, j As Long

    style = ToInt(pr_gf(OFS_PARM0))
    val = G_STRING(OFS_PARM1)

    ' change the string in sv
    sv.lightstyles(style) = val

    ' send message to all clients on this server
    If sv.state <> ss_active Then Exit Sub

    For j = 0 To svs.maxclients - 1
        If svs.clients(j).active Or svs.clients(j).spawned Then
            MSG_WriteChar svs.clients(j).message, svc_lightstyle
            MSG_WriteChar svs.clients(j).message, style
            MSG_WriteString svs.clients(j).message, val
        End If
    Next
End Sub

Private Sub PF_rint()
    Dim f As Single
    f = pr_gf(OFS_PARM0)
    If f > 0 Then
        pr_gf(OFS_RETURN) = ToInt(f + 0.5)
    Else
        pr_gf(OFS_RETURN) = ToInt(f - 0.5)
    End If
End Sub

Private Sub PF_floor()
    pr_gf(OFS_RETURN) = Int(pr_gf(OFS_PARM0))
End Sub

Private Sub PF_ceil()
    pr_gf(OFS_RETURN) = -Int(-pr_gf(OFS_PARM0))
End Sub

Private Sub PF_checkbottom()
    pr_gf(OFS_RETURN) = QB(SV_CheckBottom(G_EDICT(OFS_PARM0)))
End Sub

Private Sub PF_pointcontents()
    pr_gf(OFS_RETURN) = SV_PointContents(pr_global_struct(0).parm(0))
End Sub

' entity nextent(entity)
Private Sub PF_nextent()
    Dim i As Long

    i = G_EDICTNUM(OFS_PARM0)
    Do
        i = i + 1
        If i = sv.num_edicts Then
            RETURN_EDICT 0
            Exit Sub
        End If
        If Not edicts(i).free Then
            RETURN_EDICT i
            Exit Sub
        End If
    Loop
End Sub

' Pick a vector for the player to shoot along
' vector aim(entity, missilespeed)
Private Sub PF_aim()
    Dim ent As Long, check As Long, bestent As Long, start As vec3_t, dir_ As vec3_t, end_ As vec3_t
    Dim bestdir As vec3_t, i As Long, j As Long, tr As trace_t, dist As Single, bestdist As Single
    Dim speed As Single, fwd As vec3_t

    ent = G_EDICT(OFS_PARM0)
    speed = pr_gf(OFS_PARM1)

    start = ev(ent).origin
    start.v(2) = start.v(2) + 20

    ' try sending a trace straight
    dir_ = pr_global_struct(0).v_forward
    VectorMA start, 2048, dir_, end_
    tr = SV_Move(start, vec3_origin, vec3_origin, end_, 0, ent)
    If tr.ent <> -1 Then
        If ev(tr.ent).takedamage = DAMAGE_AIM Then
            If cvars(cv_teamplay).value = 0 Or ev(ent).team <= 0 Or ev(ent).team <> ev(tr.ent).team Then
                pr_global_struct(0).ret = pr_global_struct(0).v_forward
                Exit Sub
            End If
        End If
    End If

    ' try all possible entities
    bestdir = dir_
    bestdist = cvars(cv_sv_aim).value
    bestent = -1

    For i = 1 To sv.num_edicts - 1
        check = i
        If ev(check).takedamage <> DAMAGE_AIM Then GoTo nexti
        If check = ent Then GoTo nexti
        If cvars(cv_teamplay).value <> 0 And ev(ent).team > 0 And ev(ent).team = ev(check).team Then GoTo nexti   ' don't aim at teammate
        For j = 0 To 2
            end_.v(j) = ev(check).origin.v(j) + 0.5 * (ev(check).mins.v(j) + ev(check).maxs.v(j))
        Next
        VectorSubtract end_, start, dir_
        VectorNormalize dir_
        fwd = pr_global_struct(0).v_forward
        dist = DotProduct(dir_, fwd)
        If dist < bestdist Then GoTo nexti   ' to far to turn
        tr = SV_Move(start, vec3_origin, vec3_origin, end_, 0, ent)
        If tr.ent = check Then
            ' can shoot at this one
            bestdist = dist
            bestent = check
        End If
nexti:
    Next

    If bestent <> -1 Then
        VectorSubtract ev(bestent).origin, ev(ent).origin, dir_
        fwd = pr_global_struct(0).v_forward
        dist = DotProduct(dir_, fwd)
        VectorScale fwd, dist, end_
        end_.v(2) = dir_.v(2)
        VectorNormalize end_
        pr_global_struct(0).ret = end_
    Else
        pr_global_struct(0).ret = bestdir
    End If
End Sub

' This was a major timewaster in progs, so it was converted to C
Public Sub PF_changeyaw()
    Dim ent As Long, ideal As Single, current As Single, move As Single, speed As Single

    ent = PROG_TO_EDICT(pr_global_struct(0).self)
    current = anglemod(ev(ent).angles.v(1))
    ideal = ev(ent).ideal_yaw
    speed = ev(ent).yaw_speed

    If current = ideal Then Exit Sub
    move = ideal - current
    If ideal > current Then
        If move >= 180 Then move = move - 360
    Else
        If move <= -180 Then move = move + 360
    End If
    If move > 0 Then
        If move > speed Then move = speed
    Else
        If move < -speed Then move = -speed
    End If

    ev(ent).angles.v(1) = anglemod(current + move)
End Sub

'===============================================================================
' MESSAGE WRITING

' which buffer: MSG_BROADCAST, MSG_ONE (client gets the svs.clients() index), MSG_ALL, MSG_INIT
Private Function WriteDest(client As Long) As Long
    Dim entnum As Long, dest As Long, ent As Long

    dest = ToInt(pr_gf(OFS_PARM0))
    Select Case dest
    Case MSG_BROADCAST, MSG_ALL, MSG_INIT
        WriteDest = dest

    Case MSG_ONE
        ent = PROG_TO_EDICT(pr_global_struct(0).msg_entity)
        entnum = NUM_FOR_EDICT(ent)
        If entnum < 1 Or entnum > svs.maxclients Then PR_RunError "WriteDest: not a client"
        client = entnum - 1
        WriteDest = dest

    Case Else
        PR_RunError "WriteDest: bad destination"
    End Select
End Function

Private Sub WriteTo(sb As sizebuf_t, ByVal kind As Long)
    Select Case kind
    Case WR_BYTE: MSG_WriteByte sb, pr_gf(OFS_PARM1)
    Case WR_CHAR: MSG_WriteChar sb, pr_gf(OFS_PARM1)
    Case WR_SHORT: MSG_WriteShort sb, pr_gf(OFS_PARM1)
    Case WR_LONG: MSG_WriteLong sb, pr_gf(OFS_PARM1)
    Case WR_ANGLE: MSG_WriteAngle sb, pr_gf(OFS_PARM1)
    Case WR_COORD: MSG_WriteCoord sb, pr_gf(OFS_PARM1)
    Case WR_STRING: MSG_WriteString sb, G_STRING(OFS_PARM1)
    Case WR_ENTITY: MSG_WriteShort sb, G_EDICTNUM(OFS_PARM1)
    End Select
End Sub

' C: MSG_Write<kind> (WriteDest(), <parm 1>)
Private Sub WriteToDest(ByVal kind As Long)
    Dim client As Long
    Select Case WriteDest(client)
    Case MSG_BROADCAST: WriteTo sv.datagram, kind
    Case MSG_ONE: WriteTo svs.clients(client).message, kind
    Case MSG_ALL: WriteTo sv.reliable_datagram, kind
    Case MSG_INIT: WriteTo sv.signon, kind
    End Select
End Sub

'==============================================================================

Private Sub PF_makestatic()
    Dim ent As Long, i As Long

    ent = G_EDICT(OFS_PARM0)

    MSG_WriteByte sv.signon, svc_spawnstatic

    MSG_WriteByte sv.signon, SV_ModelIndex(PR_GetString(ev(ent).model))

    MSG_WriteByte sv.signon, ev(ent).frame
    MSG_WriteByte sv.signon, ev(ent).colormap
    MSG_WriteByte sv.signon, ev(ent).skin
    For i = 0 To 2
        MSG_WriteCoord sv.signon, ev(ent).origin.v(i)
        MSG_WriteAngle sv.signon, ev(ent).angles.v(i)
    Next

    ' throw the entity away now
    ED_Free ent
End Sub

'==============================================================================

Private Sub PF_setspawnparms()
    Dim ent As Long, i As Long, client As Long

    ent = G_EDICT(OFS_PARM0)
    i = NUM_FOR_EDICT(ent)
    If i < 1 Or i > svs.maxclients Then PR_RunError "Entity is not a client"

    ' copy spawn parms out of the client_t
    client = i - 1

    For i = 0 To NUM_SPAWN_PARMS - 1
        pr_gf(GOFS_parm1 + i) = svs.clients(client).spawn_parms(i)
    Next
End Sub

Private Sub PF_changelevel()
    Dim s As String

    ' make sure we don't issue two changelevels
    If svs.changelevel_issued Then Exit Sub
    svs.changelevel_issued = True

    s = G_STRING(OFS_PARM0)
    Cbuf_AddText va("changelevel %s\n", s)
End Sub

Private Sub PF_Fixme()
    PR_RunError "unimplemented bulitin"
End Sub

' C: pr_builtins[i] ()
Public Sub PR_CallBuiltin(ByVal i As Long)
    Select Case i
    Case 1: PF_makevectors               ' void(entity e)	makevectors 		= #1;
    Case 2: PF_setorigin                 ' void(entity e, vector o) setorigin	= #2;
    Case 3: PF_setmodel                  ' void(entity e, string m) setmodel	= #3;
    Case 4: PF_setsize                   ' void(entity e, vector min, vector max) setsize = #4;
    Case 6: PF_break                     ' void() break						= #6;
    Case 7: PF_random                    ' float() random						= #7;
    Case 8: PF_sound                     ' void(entity e, float chan, string samp) sound = #8;
    Case 9: PF_normalize                 ' vector(vector v) normalize			= #9;
    Case 10: PF_error                    ' void(string e) error				= #10;
    Case 11: PF_objerror                 ' void(string e) objerror				= #11;
    Case 12: PF_vlen                     ' float(vector v) vlen				= #12;
    Case 13: PF_vectoyaw                 ' float(vector v) vectoyaw		= #13;
    Case 14: PF_Spawn                    ' entity() spawn						= #14;
    Case 15: PF_Remove                   ' void(entity e) remove				= #15;
    Case 16: PF_traceline                ' float(vector v1, vector v2, float tryents) traceline = #16;
    Case 17: PF_checkclient              ' entity() clientlist					= #17;
    Case 18: PF_Find                     ' entity(entity start, .string fld, string match) find = #18;
    Case 19: PF_precache_sound           ' void(string s) precache_sound		= #19;
    Case 20: PF_precache_model           ' void(string s) precache_model		= #20;
    Case 21: PF_stuffcmd                 ' void(entity client, string s)stuffcmd = #21;
    Case 22: PF_findradius               ' entity(vector org, float rad) findradius = #22;
    Case 23: PF_bprint                   ' void(string s) bprint				= #23;
    Case 24: PF_sprint                   ' void(entity client, string s) sprint = #24;
    Case 25: PF_dprint                   ' void(string s) dprint				= #25;
    Case 26: PF_ftos                     ' void(string s) ftos				= #26;
    Case 27: PF_vtos                     ' void(string s) vtos				= #27;
    Case 28: PF_coredump
    Case 29: PF_traceon
    Case 30: PF_traceoff
    Case 31: PF_eprint                   ' void(entity e) debug print an entire entity
    Case 32: PF_walkmove                 ' float(float yaw, float dist) walkmove
    Case 34: PF_droptofloor
    Case 35: PF_lightstyle
    Case 36: PF_rint
    Case 37: PF_floor
    Case 38: PF_ceil
    Case 40: PF_checkbottom
    Case 41: PF_pointcontents
    Case 43: PF_fabs
    Case 44: PF_aim
    Case 45: PF_cvar
    Case 46: PF_localcmd
    Case 47: PF_nextent
    Case 48: PF_particle
    Case 49: PF_changeyaw
    Case 51: PF_vectoangles
    Case 52: WriteToDest WR_BYTE         ' PF_WriteByte
    Case 53: WriteToDest WR_CHAR         ' PF_WriteChar
    Case 54: WriteToDest WR_SHORT        ' PF_WriteShort
    Case 55: WriteToDest WR_LONG         ' PF_WriteLong
    Case 56: WriteToDest WR_COORD        ' PF_WriteCoord
    Case 57: WriteToDest WR_ANGLE        ' PF_WriteAngle
    Case 58: WriteToDest WR_STRING       ' PF_WriteString
    Case 59: WriteToDest WR_ENTITY       ' PF_WriteEntity
    Case 67: SV_MoveToGoal
    Case 68: PF_precache_file
    Case 69: PF_makestatic
    Case 70: PF_changelevel
    Case 72: PF_cvar_set
    Case 73: PF_centerprint
    Case 74: PF_ambientsound
    Case 75: PF_precache_model
    Case 76: PF_precache_sound           ' precache_sound2 is different only for qcc
    Case 77: PF_precache_file
    Case 78: PF_setspawnparms
    Case Else: PF_Fixme                  ' 0, 5, 33, 39, 42, 50, 60 .. 66, 71
    End Select
End Sub
