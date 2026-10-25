Attribute VB_Name = "cl_main"
' cl_main.bas - client main loop (cl_main.c)
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public Sub CL_ClearState()
    Dim i As Long, blank As client_state_t, blankent As entity_t, blankdl As dlight_t
    Dim blankls As lightstyle_t, blankbeam As beam_t, blankef As efrag_t

    If Not sv.active Then Host_ClearMemory

    ' wipe the entire cl structure
    cl = blank
    cl_entities(CL_VIEWENT) = blankent  ' (C: cl.viewent is part of cl)

    SZ_Clear cls.message

    ' clear other arrays
    For i = 1 To MAX_EFRAGS
        cl_efrags(i) = blankef
    Next
    For i = 0 To MAX_EDICTS - 1
        cl_entities(i) = blankent
    Next
    For i = 0 To MAX_DLIGHTS - 1
        cl_dlights(i) = blankdl
    Next
    For i = 0 To MAX_LIGHTSTYLES - 1
        cl_lightstyle(i) = blankls
    Next
    For i = 0 To MAX_TEMP_ENTITIES - 1
        cl_entities(CL_TEMP_BASE + i) = blankent
    Next
    For i = 0 To MAX_BEAMS - 1
        cl_beams(i) = blankbeam
    Next

    '
    ' allocate the efrags and chain together into a free list
    '
    cl.free_efrags = 1
    For i = 1 To MAX_EFRAGS - 1
        cl_efrags(i).entnext = i + 1
    Next
    cl_efrags(i).entnext = 0
End Sub

' Sends a disconnect message to the server
' This is also called on Host_Error, so it shouldn't cause any errors
Public Sub CL_Disconnect()
    ' stop sounds (especially looping!)
    S_StopAllSounds True

    ' bring the console down and fade the colors back to normal
    ' SCR_BringDownConsole ();

    ' if running a local server, shut it down
    If cls.demoplayback Then
        CL_StopPlayback
    ElseIf cls.state = ca_connected Then
        If cls.demorecording Then CL_Stop_f

        Con_DPrintf "Sending clc_disconnect\n"
        SZ_Clear cls.message
        MSG_WriteByte cls.message, clc_disconnect
        NET_SendUnreliableMessage cls.netcon, cls.message
        SZ_Clear cls.message
        NET_Close cls.netcon

        cls.state = ca_disconnected
        If sv.active Then Host_ShutdownServer False
    End If

    cls.demoplayback = False
    cls.timedemo = False
    cls.signon = 0
End Sub

Public Sub CL_Disconnect_f()
    CL_Disconnect
    If sv.active Then Host_ShutdownServer False
End Sub

' Host should be either "local" or a net address to be passed on
Public Sub CL_EstablishConnection(ByVal host As String)
    If cls.state = ca_dedicated Then Exit Sub

    If cls.demoplayback Then Exit Sub

    CL_Disconnect

    cls.netcon = NET_Connect(host)
    If cls.netcon = 0 Then Host_Error "CL_Connect: connect failed\n"
    Con_DPrintf "CL_EstablishConnection: connected to %s\n", host

    cls.demonum = -1                     ' not in the demo loop now
    cls.state = ca_connected
    cls.signon = 0                       ' need all the signon messages before playing
End Sub

' An svc_signonnum has been received, perform a client side setup
Public Sub CL_SignonReply()
    Con_DPrintf "CL_SignonReply: %i\n", cls.signon

    Select Case cls.signon
    Case 1
        MSG_WriteByte cls.message, clc_stringcmd
        MSG_WriteString cls.message, "prespawn"

    Case 2
        MSG_WriteByte cls.message, clc_stringcmd
        MSG_WriteString cls.message, va("name ""%s""\n", cvars(cv_cl_name).strval)

        MSG_WriteByte cls.message, clc_stringcmd
        MSG_WriteString cls.message, va("color %i %i\n", (ToInt(cvars(cv_cl_color).value) And Not 15) \ 16, ToInt(cvars(cv_cl_color).value) And 15)

        MSG_WriteByte cls.message, clc_stringcmd
        MSG_WriteString cls.message, "spawn " & cls.spawnparms

    Case 3
        MSG_WriteByte cls.message, clc_stringcmd
        MSG_WriteString cls.message, "begin"
        ' Cache_Report ();  (print remaining memory: there is no cache)

    Case 4
        SCR_EndLoadingPlaque             ' allow normal screen updates
    End Select
End Sub

' Called to play the next demo in the demo loop
Public Sub CL_NextDemo()
    If cls.demonum = -1 Then Exit Sub    ' don't play demos

    SCR_BeginLoadingPlaque

    ' (C tests cls.demos[cls.demonum][0] first, a read past the array at MAX_DEMOS)
    If cls.demonum = MAX_DEMOS Then
        cls.demonum = 0
    ElseIf Len(cls.demos(cls.demonum)) = 0 Then
        cls.demonum = 0
    End If
    If Len(cls.demos(cls.demonum)) = 0 Then
        Con_Printf "No demos listed with startdemos\n"
        cls.demonum = -1
        Exit Sub
    End If

    Cbuf_InsertText "playdemo " & cls.demos(cls.demonum) & vbLf
    cls.demonum = cls.demonum + 1
End Sub

Private Sub CL_PrintEntities_f()
    Dim i As Long

    For i = 0 To cl.num_entities - 1
        Con_Printf "%3i:", i
        With cl_entities(i)
            If .model = 0 Then
                Con_Printf "EMPTY\n"
            Else
                Con_Printf "%s:%2i  (%5.1f,%5.1f,%5.1f) [%5.1f %5.1f %5.1f]\n", _
                    mod_known(.model).name, .frame, .origin.v(0), .origin.v(1), .origin.v(2), .angles.v(0), .angles.v(1), .angles.v(2)
            End If
        End With
    Next
End Sub

' Debugging tool, just flashes the screen (C: the body is #if 0)
Private Sub SetPal(ByVal i As Long)
End Sub

' returns a cl_dlights() index
Public Function CL_AllocDlight(ByVal key As Long) As Long
    Dim i As Long, blank As dlight_t

    ' first look for an exact key match
    If key Then
        For i = 0 To MAX_DLIGHTS - 1
            If cl_dlights(i).key = key Then
                cl_dlights(i) = blank
                cl_dlights(i).key = key
                CL_AllocDlight = i
                Exit Function
            End If
        Next
    End If

    ' then look for anything else
    For i = 0 To MAX_DLIGHTS - 1
        If cl_dlights(i).die < cl.time Then
            cl_dlights(i) = blank
            cl_dlights(i).key = key
            CL_AllocDlight = i
            Exit Function
        End If
    Next

    cl_dlights(0) = blank
    cl_dlights(0).key = key
    CL_AllocDlight = 0
End Function

Public Sub CL_DecayLights()
    Dim i As Long, time_ As Single

    time_ = cl.time - cl.oldtime

    For i = 0 To MAX_DLIGHTS - 1
        With cl_dlights(i)
            If .die < cl.time Or .radius = 0 Then GoTo nexti

            .radius = .radius - time_ * .decay
            If .radius < 0 Then .radius = 0
nexti:                                   ' (inside the With: see tools/check_with.py)
        End With
    Next
End Sub

' Determines the fraction between the last two messages that the objects
' should be put at.
Private Function CL_LerpPoint() As Single
    Dim f As Single, frac As Single

    f = cl.mtime(0) - cl.mtime(1)

    If f = 0 Or cvars(cv_cl_nolerp).value <> 0 Or cls.timedemo Or sv.active Then
        cl.time = cl.mtime(0)
        CL_LerpPoint = 1
        Exit Function
    End If

    If f > 0.1 Then
        ' dropped packet, or start of demo
        cl.mtime(1) = cl.mtime(0) - 0.1
        f = 0.1
    End If
    frac = (cl.time - cl.mtime(1)) / f
    ' Con_Printf ("frac: %f\n",frac);
    If frac < 0 Then
        If frac < -0.01 Then
            SetPal 1
            cl.time = cl.mtime(1)
            ' Con_Printf ("low frac\n");
        End If
        frac = 0
    ElseIf frac > 1 Then
        If frac > 1.01 Then
            SetPal 2
            cl.time = cl.mtime(0)
            ' Con_Printf ("high frac\n");
        End If
        frac = 1
    Else
        SetPal 0
    End If

    CL_LerpPoint = frac
End Function

Private Sub CL_RelinkEntities()
    Dim ent As Long, i As Long, j As Long, frac As Single, f As Single, d As Single
    Dim delta As vec3_t, bobjrotate As Single, oldorg As vec3_t, dl As Long
    Dim fv As vec3_t, rv As vec3_t, uv As vec3_t, skipent As Boolean

    ' determine partial update time
    frac = CL_LerpPoint()

    cl_numvisedicts = 0

    '
    ' interpolate player info
    '
    For i = 0 To 2
        cl.velocity.v(i) = cl.mvelocity(1).v(i) + frac * (cl.mvelocity(0).v(i) - cl.mvelocity(1).v(i))
    Next

    If cls.demoplayback Then
        ' interpolate the angles
        For j = 0 To 2
            d = cl.mviewangles(0).v(j) - cl.mviewangles(1).v(j)
            If d > 180 Then
                d = d - 360
            ElseIf d < -180 Then
                d = d + 360
            End If
            cl.viewangles.v(j) = cl.mviewangles(1).v(j) + frac * d
        Next
    End If

    bobjrotate = anglemod(100 * cl.time)

    ' start on the entity after the world
    For i = 1 To cl.num_entities - 1
        ent = i
        ' (QuakeVB) the continues jump to endent, inside the With block, and
        ' skipent takes them on to the next entity (see tools/check_with.py)
        skipent = True
        With cl_entities(ent)
            If .model = 0 Then
                ' empty slot
                If .forcelink Then R_RemoveEfrags ent   ' just became empty
                GoTo endent
            End If

            ' if the object wasn't included in the last packet, remove it
            If .msgtime <> cl.mtime(0) Then
                .model = 0
                GoTo endent
            End If

            oldorg = .origin

            If .forcelink Then
                ' the entity was not updated in the last message
                ' so move to the final spot
                .origin = .msg_origins(0)
                .angles = .msg_angles(0)
            Else
                ' if the delta is large, assume a teleport and don't lerp
                f = frac
                For j = 0 To 2
                    delta.v(j) = .msg_origins(0).v(j) - .msg_origins(1).v(j)
                    If delta.v(j) > 100 Or delta.v(j) < -100 Then f = 1   ' assume a teleportation, not a motion
                Next

                ' interpolate the origin and angles
                For j = 0 To 2
                    .origin.v(j) = .msg_origins(1).v(j) + f * delta.v(j)

                    d = .msg_angles(0).v(j) - .msg_angles(1).v(j)
                    If d > 180 Then
                        d = d - 360
                    ElseIf d < -180 Then
                        d = d + 360
                    End If
                    .angles.v(j) = .msg_angles(1).v(j) + f * d
                Next
            End If

            ' rotate binary objects locally
            If mod_known(.model).flags And EF_ROTATE Then .angles.v(1) = bobjrotate

            If .effects And EF_BRIGHTFIELD Then R_EntityParticles ent
            If .effects And EF_MUZZLEFLASH Then
                dl = CL_AllocDlight(i)
                cl_dlights(dl).origin = .origin
                cl_dlights(dl).origin.v(2) = cl_dlights(dl).origin.v(2) + 16
                AngleVectors .angles, fv, rv, uv

                VectorMA cl_dlights(dl).origin, 18, fv, cl_dlights(dl).origin
                cl_dlights(dl).radius = 200 + (rand() And 31)
                cl_dlights(dl).minlight = 32
                cl_dlights(dl).die = cl.time + 0.1
            End If
            If .effects And EF_BRIGHTLIGHT Then
                dl = CL_AllocDlight(i)
                cl_dlights(dl).origin = .origin
                cl_dlights(dl).origin.v(2) = cl_dlights(dl).origin.v(2) + 16
                cl_dlights(dl).radius = 400 + (rand() And 31)
                cl_dlights(dl).die = cl.time + 0.001
            End If
            If .effects And EF_DIMLIGHT Then
                dl = CL_AllocDlight(i)
                cl_dlights(dl).origin = .origin
                cl_dlights(dl).radius = 200 + (rand() And 31)
                cl_dlights(dl).die = cl.time + 0.001
            End If

            If mod_known(.model).flags And EF_GIB Then
                R_RocketTrail oldorg, .origin, 2
            ElseIf mod_known(.model).flags And EF_ZOMGIB Then
                R_RocketTrail oldorg, .origin, 4
            ElseIf mod_known(.model).flags And EF_TRACER Then
                R_RocketTrail oldorg, .origin, 3
            ElseIf mod_known(.model).flags And EF_TRACER2 Then
                R_RocketTrail oldorg, .origin, 5
            ElseIf mod_known(.model).flags And EF_ROCKET Then
                R_RocketTrail oldorg, .origin, 0
                dl = CL_AllocDlight(i)
                cl_dlights(dl).origin = .origin
                cl_dlights(dl).radius = 200
                cl_dlights(dl).die = cl.time + 0.01
            ElseIf mod_known(.model).flags And EF_GRENADE Then
                R_RocketTrail oldorg, .origin, 1
            ElseIf mod_known(.model).flags And EF_TRACER3 Then
                R_RocketTrail oldorg, .origin, 6
            End If

            .forcelink = False
            skipent = False
endent:
        End With
        If skipent Then GoTo nexti

        If i = cl.viewentity And cvars(cv_chase_active).value = 0 Then GoTo nexti

        If cl_numvisedicts < MAX_VISEDICTS Then
            cl_visedicts(cl_numvisedicts) = ent
            cl_numvisedicts = cl_numvisedicts + 1
        End If
nexti:
    Next
End Sub

' Read all incoming data from the server
Public Function CL_ReadFromServer() As Long
    Dim ret As Long

    cl.oldtime = cl.time
    cl.time = cl.time + host_frametime

    Do
        ret = CL_GetMessage()
        If ret = -1 Then Host_Error "CL_ReadFromServer: lost server connection"
        If ret = 0 Then Exit Do

        cl.last_received_message = realtime
        CL_ParseServerMessage
    Loop While ret <> 0 And cls.state = ca_connected

    If cvars(cv_cl_shownet).value <> 0 Then Con_Printf "\n"

    CL_RelinkEntities
    CL_UpdateTEnts

    '
    ' bring the links up to date
    '
    CL_ReadFromServer = 0
End Function

Public Sub CL_SendCmd()
    Dim cmd_ As usercmd_t

    If cls.state <> ca_connected Then Exit Sub

    If cls.signon = SIGNONS Then
        ' get basic movement from keyboard
        CL_BaseMove cmd_

        ' allow mice or other external controllers to add to the move
        IN_Move cmd_

        ' send the unreliable message
        CL_SendMove cmd_
    End If

    If cls.demoplayback Then
        SZ_Clear cls.message
        Exit Sub
    End If

    ' send the reliable message
    If cls.message.cursize = 0 Then Exit Sub   ' no message at all

    If Not NET_CanSendMessage(cls.netcon) Then
        Con_DPrintf "CL_WriteToServer: can't send\n"
        Exit Sub
    End If

    If NET_SendMessage(cls.netcon, cls.message) = -1 Then Host_Error "CL_WriteToServer: lost server connection"

    SZ_Clear cls.message
End Sub

Public Sub CL_Init()
    SZ_Alloc cls.message, 1024
    ReDim cl_entities(0 To CL_NUM_ENTITIES - 1)
    ReDim cl_efrags(1 To MAX_EFRAGS)

    CL_InitInput
    CL_InitTEnts

    '
    ' register our commands
    '
    cv_cl_name = Cvar_RegisterVariable("_cl_name", "player", True)
    cv_cl_color = Cvar_RegisterVariable("_cl_color", "0", True)
    cv_cl_upspeed = Cvar_RegisterVariable("cl_upspeed", "200")
    cv_cl_forwardspeed = Cvar_RegisterVariable("cl_forwardspeed", "200", True)
    cv_cl_backspeed = Cvar_RegisterVariable("cl_backspeed", "200", True)
    cv_cl_sidespeed = Cvar_RegisterVariable("cl_sidespeed", "350")
    cv_cl_movespeedkey = Cvar_RegisterVariable("cl_movespeedkey", "2.0")
    cv_cl_yawspeed = Cvar_RegisterVariable("cl_yawspeed", "140")
    cv_cl_pitchspeed = Cvar_RegisterVariable("cl_pitchspeed", "150")
    cv_cl_anglespeedkey = Cvar_RegisterVariable("cl_anglespeedkey", "1.5")
    cv_cl_shownet = Cvar_RegisterVariable("cl_shownet", "0")    ' can be 0, 1, or 2
    cv_cl_nolerp = Cvar_RegisterVariable("cl_nolerp", "0")
    cv_lookspring = Cvar_RegisterVariable("lookspring", "0", True)
    cv_lookstrafe = Cvar_RegisterVariable("lookstrafe", "0", True)
    cv_sensitivity = Cvar_RegisterVariable("sensitivity", "3", True)

    cv_m_pitch = Cvar_RegisterVariable("m_pitch", "0.022", True)
    cv_m_yaw = Cvar_RegisterVariable("m_yaw", "0.022", True)
    cv_m_forward = Cvar_RegisterVariable("m_forward", "1", True)
    cv_m_side = Cvar_RegisterVariable("m_side", "0.8", True)

    ' Cvar_RegisterVariable (&cl_autofire);

    Cmd_AddCommand "entities", XC_CL_PrintEntities_f
    Cmd_AddCommand "disconnect", XC_CL_Disconnect_f
    Cmd_AddCommand "record", XC_CL_Record_f
    Cmd_AddCommand "stop", XC_CL_Stop_f
    Cmd_AddCommand "playdemo", XC_CL_PlayDemo_f
    Cmd_AddCommand "timedemo", XC_CL_TimeDemo_f
End Sub

' (QuakeVB) the console commands registered by CL_Init, for Cmd_Call
Public Sub CL_CallCommand(ByVal fn As Long)
    Select Case fn
    Case XC_CL_PrintEntities_f: CL_PrintEntities_f
    Case XC_CL_Disconnect_f: CL_Disconnect_f
    Case XC_CL_Record_f: CL_Record_f
    Case XC_CL_Stop_f: CL_Stop_f
    Case XC_CL_PlayDemo_f: CL_PlayDemo_f
    Case XC_CL_TimeDemo_f: CL_TimeDemo_f
    Case Else: Sys_Error "CL_CallCommand: bad function %i", fn
    End Select
End Sub
