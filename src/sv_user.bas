Attribute VB_Name = "sv_user"
' sv_user.bas - server code for moving users (sv_user.c)
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public cv_sv_edgefriction As Long
Public cv_sv_idealpitchscale As Long
Public cv_sv_maxspeed As Long
Public cv_sv_accelerate As Long

Private sv_forward As vec3_t, sv_right As vec3_t, sv_up As vec3_t

Private wishdir As vec3_t
Private wishspeed As Single

' world
' (QuakeVB: C's angles / origin / velocity pointers are ev(sv_player).angles etc.)

Private onground As Boolean

Private sv_cmd As usercmd_t              ' C: cmd

Private Const MAX_FORWARD As Long = 6

Public Sub SV_SetIdealPitch()
    Dim angleval As Single, sinval As Single, cosval As Single, tr As trace_t
    Dim top As vec3_t, bottom As vec3_t, z(0 To MAX_FORWARD - 1) As Single
    Dim i As Long, j As Long, step_ As Long, dir_ As Long, steps As Long

    If (ToInt(ev(sv_player).flags) And FL_ONGROUND) = 0 Then Exit Sub

    angleval = ev(sv_player).angles.v(YAW) * M_PI * 2 / 360
    sinval = Sin(angleval)
    cosval = Cos(angleval)

    For i = 0 To MAX_FORWARD - 1
        top.v(0) = ev(sv_player).origin.v(0) + cosval * (i + 3) * 12
        top.v(1) = ev(sv_player).origin.v(1) + sinval * (i + 3) * 12
        top.v(2) = ev(sv_player).origin.v(2) + ev(sv_player).view_ofs.v(2)

        bottom.v(0) = top.v(0)
        bottom.v(1) = top.v(1)
        bottom.v(2) = top.v(2) - 160

        tr = SV_Move(top, vec3_origin, vec3_origin, bottom, MOVE_NOMONSTERS, sv_player)
        If tr.allsolid Then Exit Sub     ' looking at a wall, leave ideal the way is was

        If tr.fraction = 1 Then Exit Sub   ' near a dropoff

        z(i) = top.v(2) + tr.fraction * (bottom.v(2) - top.v(2))
    Next

    dir_ = 0
    steps = 0
    For j = 1 To i - 1
        step_ = ToInt(z(j) - z(j - 1))
        If step_ > -ON_EPSILON And step_ < ON_EPSILON Then GoTo nextj

        If dir_ <> 0 Then
            If step_ - dir_ > ON_EPSILON Or step_ - dir_ < -ON_EPSILON Then Exit Sub   ' mixed changes
        End If

        steps = steps + 1
        dir_ = step_
nextj:
    Next

    If dir_ = 0 Then
        ev(sv_player).idealpitch = 0
        Exit Sub
    End If

    If steps < 2 Then Exit Sub
    ev(sv_player).idealpitch = -dir_ * cvars(cv_sv_idealpitchscale).value
End Sub

Private Sub SV_UserFriction()
    Dim speed As Single, newspeed As Single, control As Single
    Dim start As vec3_t, stop_ As vec3_t, friction As Single, trace As trace_t

    With ev(sv_player)
        speed = Sqr(.velocity.v(0) * .velocity.v(0) + .velocity.v(1) * .velocity.v(1))
        If speed = 0 Then Exit Sub

        ' if the leading edge is over a dropoff, increase friction
        start.v(0) = .origin.v(0) + .velocity.v(0) / speed * 16
        stop_.v(0) = start.v(0)
        start.v(1) = .origin.v(1) + .velocity.v(1) / speed * 16
        stop_.v(1) = start.v(1)
        start.v(2) = .origin.v(2) + .mins.v(2)
        stop_.v(2) = start.v(2) - 34
    End With

    trace = SV_Move(start, vec3_origin, vec3_origin, stop_, MOVE_NOMONSTERS, sv_player)

    If trace.fraction = 1 Then
        friction = cvars(cv_sv_friction).value * cvars(cv_sv_edgefriction).value
    Else
        friction = cvars(cv_sv_friction).value
    End If

    ' apply friction
    If speed < cvars(cv_sv_stopspeed).value Then control = cvars(cv_sv_stopspeed).value Else control = speed
    newspeed = speed - host_frametime * control * friction

    If newspeed < 0 Then newspeed = 0
    newspeed = newspeed / speed

    With ev(sv_player)
        .velocity.v(0) = .velocity.v(0) * newspeed
        .velocity.v(1) = .velocity.v(1) * newspeed
        .velocity.v(2) = .velocity.v(2) * newspeed
    End With
End Sub

Private Sub SV_Accelerate()
    Dim i As Long, addspeed As Single, accelspeed As Single, currentspeed As Single

    currentspeed = DotProduct(ev(sv_player).velocity, wishdir)
    addspeed = wishspeed - currentspeed
    If addspeed <= 0 Then Exit Sub
    accelspeed = cvars(cv_sv_accelerate).value * host_frametime * wishspeed
    If accelspeed > addspeed Then accelspeed = addspeed

    For i = 0 To 2
        ev(sv_player).velocity.v(i) = ev(sv_player).velocity.v(i) + accelspeed * wishdir.v(i)
    Next
End Sub

Private Sub SV_AirAccelerate(wishveloc As vec3_t)
    Dim i As Long, addspeed As Single, wishspd As Single, accelspeed As Single, currentspeed As Single

    wishspd = VectorNormalize(wishveloc)
    If wishspd > 30 Then wishspd = 30
    currentspeed = DotProduct(ev(sv_player).velocity, wishveloc)
    addspeed = wishspd - currentspeed
    If addspeed <= 0 Then Exit Sub
    ' accelspeed = sv_accelerate.value * host_frametime;
    accelspeed = cvars(cv_sv_accelerate).value * wishspeed * host_frametime
    If accelspeed > addspeed Then accelspeed = addspeed

    For i = 0 To 2
        ev(sv_player).velocity.v(i) = ev(sv_player).velocity.v(i) + accelspeed * wishveloc.v(i)
    Next
End Sub

Private Sub DropPunchAngle()
    Dim len_ As Single

    len_ = VectorNormalize(ev(sv_player).punchangle)

    len_ = len_ - 10 * host_frametime
    If len_ < 0 Then len_ = 0
    VectorScale ev(sv_player).punchangle, len_, ev(sv_player).punchangle
End Sub

Private Sub SV_WaterMove()
    Dim i As Long, wishvel As vec3_t
    Dim speed As Single, newspeed As Single, wishspeed As Single, addspeed As Single, accelspeed As Single

    '
    ' user intentions
    '
    AngleVectors ev(sv_player).v_angle, sv_forward, sv_right, sv_up

    For i = 0 To 2
        wishvel.v(i) = sv_forward.v(i) * sv_cmd.forwardmove + sv_right.v(i) * sv_cmd.sidemove
    Next

    If sv_cmd.forwardmove = 0 And sv_cmd.sidemove = 0 And sv_cmd.upmove = 0 Then
        wishvel.v(2) = wishvel.v(2) - 60   ' drift towards bottom
    Else
        wishvel.v(2) = wishvel.v(2) + sv_cmd.upmove
    End If

    wishspeed = Length(wishvel)
    If wishspeed > cvars(cv_sv_maxspeed).value Then
        VectorScale wishvel, cvars(cv_sv_maxspeed).value / wishspeed, wishvel
        wishspeed = cvars(cv_sv_maxspeed).value
    End If
    wishspeed = wishspeed * 0.7

    '
    ' water friction
    '
    speed = Length(ev(sv_player).velocity)
    If speed <> 0 Then
        newspeed = speed - host_frametime * speed * cvars(cv_sv_friction).value
        If newspeed < 0 Then newspeed = 0
        VectorScale ev(sv_player).velocity, newspeed / speed, ev(sv_player).velocity
    Else
        newspeed = 0
    End If

    '
    ' water acceleration
    '
    If wishspeed = 0 Then Exit Sub

    addspeed = wishspeed - newspeed
    If addspeed <= 0 Then Exit Sub

    VectorNormalize wishvel
    accelspeed = cvars(cv_sv_accelerate).value * wishspeed * host_frametime
    If accelspeed > addspeed Then accelspeed = addspeed

    For i = 0 To 2
        ev(sv_player).velocity.v(i) = ev(sv_player).velocity.v(i) + accelspeed * wishvel.v(i)
    Next
End Sub

Private Sub SV_WaterJump()
    If sv.time > ev(sv_player).teleport_time Or ev(sv_player).waterlevel = 0 Then
        ev(sv_player).flags = ToInt(ev(sv_player).flags) And Not FL_WATERJUMP
        ev(sv_player).teleport_time = 0
    End If
    ev(sv_player).velocity.v(0) = ev(sv_player).movedir.v(0)
    ev(sv_player).velocity.v(1) = ev(sv_player).movedir.v(1)
End Sub

Private Sub SV_AirMove()
    Dim i As Long, wishvel As vec3_t, fmove As Single, smove As Single

    AngleVectors ev(sv_player).angles, sv_forward, sv_right, sv_up

    fmove = sv_cmd.forwardmove
    smove = sv_cmd.sidemove

    ' hack to not let you back into teleporter
    If sv.time < ev(sv_player).teleport_time And fmove < 0 Then fmove = 0

    For i = 0 To 2
        wishvel.v(i) = sv_forward.v(i) * fmove + sv_right.v(i) * smove
    Next

    If ToInt(ev(sv_player).movetype) <> MOVETYPE_WALK Then
        wishvel.v(2) = sv_cmd.upmove
    Else
        wishvel.v(2) = 0
    End If

    wishdir = wishvel
    wishspeed = VectorNormalize(wishdir)
    If wishspeed > cvars(cv_sv_maxspeed).value Then
        VectorScale wishvel, cvars(cv_sv_maxspeed).value / wishspeed, wishvel
        wishspeed = cvars(cv_sv_maxspeed).value
    End If

    If ev(sv_player).movetype = MOVETYPE_NOCLIP Then
        ' noclip
        ev(sv_player).velocity = wishvel
    ElseIf onground Then
        SV_UserFriction
        SV_Accelerate
    Else
        ' not on ground, so little effect on velocity
        SV_AirAccelerate wishvel
    End If
End Sub

' the move fields specify an intended velocity in pix/sec
' the angle fields specify an exact angular motion in degrees
Private Sub SV_ClientThink()
    Dim v_angle As vec3_t

    If ev(sv_player).movetype = MOVETYPE_NONE Then Exit Sub

    onground = (ToInt(ev(sv_player).flags) And FL_ONGROUND) <> 0

    DropPunchAngle

    '
    ' if dead, behave differently
    '
    If ev(sv_player).health <= 0 Then Exit Sub

    '
    ' angles
    ' show 1/3 the pitch angle and all the roll angle
    sv_cmd = svs.clients(host_client).cmd

    VectorAdd ev(sv_player).v_angle, ev(sv_player).punchangle, v_angle
    ev(sv_player).angles.v(ROLL) = V_CalcRoll(ev(sv_player).angles, ev(sv_player).velocity) * 4
    If ev(sv_player).fixangle = 0 Then
        ev(sv_player).angles.v(PITCH) = -v_angle.v(PITCH) / 3
        ev(sv_player).angles.v(YAW) = v_angle.v(YAW)
    End If

    If ToInt(ev(sv_player).flags) And FL_WATERJUMP Then
        SV_WaterJump
        Exit Sub
    End If
    '
    ' walk
    '
    If ev(sv_player).waterlevel >= 2 And ev(sv_player).movetype <> MOVETYPE_NOCLIP Then
        SV_WaterMove
        Exit Sub
    End If

    SV_AirMove
End Sub

Private Sub SV_ReadClientMove(move As usercmd_t)
    Dim i As Long, angle As vec3_t, bits As Long

    With svs.clients(host_client)
        ' read ping time
        .ping_times(.num_pings Mod NUM_PING_TIMES) = sv.time - MSG_ReadFloat()
        .num_pings = .num_pings + 1

        ' read current angles
        For i = 0 To 2
            angle.v(i) = MSG_ReadAngle()
        Next

        ev(.edict).v_angle = angle

        ' read movement
        move.forwardmove = MSG_ReadShort()
        move.sidemove = MSG_ReadShort()
        move.upmove = MSG_ReadShort()

        ' read buttons
        bits = MSG_ReadByte()
        ev(.edict).button0 = bits And 1
        ev(.edict).button2 = (bits And 2) \ 2

        i = MSG_ReadByte()
        If i Then ev(.edict).impulse = i
    End With
End Sub

' Returns false if the client should be killed
Private Function SV_ReadClientMessage() As Boolean
    Dim ret As Long, ccmd As Long, s As String

    ' (QuakeVB: C's "goto nextmsg" is the Exit Do; the inner loop is only left
    ' that way, so the outer do / while (ret == 1) never tests its condition)
    Do
        ret = NET_GetMessage(svs.clients(host_client).netconnection)
        If ret = -1 Then
            Sys_Printf "SV_ReadClientMessage: NET_GetMessage failed\n"
            Exit Function
        End If
        If ret = 0 Then
            SV_ReadClientMessage = True
            Exit Function
        End If

        MSG_BeginReading

        Do
            If Not svs.clients(host_client).active Then Exit Function   ' a command caused an error

            If msg_badread Then
                Sys_Printf "SV_ReadClientMessage: badread\n"
                Exit Function
            End If

            ccmd = MSG_ReadChar()

            Select Case ccmd
            Case -1
                Exit Do                  ' end of message

            Case clc_nop
                ' Sys_Printf ("clc_nop\n");

            Case clc_stringcmd
                s = MSG_ReadString()
                If svs.clients(host_client).privileged Then
                    ret = 2
                Else
                    ret = 0
                End If
                If Q_strncasecmp(s, "status", 6) = 0 Then
                    ret = 1
                ElseIf Q_strncasecmp(s, "god", 3) = 0 Then
                    ret = 1
                ElseIf Q_strncasecmp(s, "notarget", 8) = 0 Then
                    ret = 1
                ElseIf Q_strncasecmp(s, "fly", 3) = 0 Then
                    ret = 1
                ElseIf Q_strncasecmp(s, "name", 4) = 0 Then
                    ret = 1
                ElseIf Q_strncasecmp(s, "noclip", 6) = 0 Then
                    ret = 1
                ElseIf Q_strncasecmp(s, "say", 3) = 0 Then
                    ret = 1
                ElseIf Q_strncasecmp(s, "say_team", 8) = 0 Then
                    ret = 1
                ElseIf Q_strncasecmp(s, "tell", 4) = 0 Then
                    ret = 1
                ElseIf Q_strncasecmp(s, "color", 5) = 0 Then
                    ret = 1
                ElseIf Q_strncasecmp(s, "kill", 4) = 0 Then
                    ret = 1
                ElseIf Q_strncasecmp(s, "pause", 5) = 0 Then
                    ret = 1
                ElseIf Q_strncasecmp(s, "spawn", 5) = 0 Then
                    ret = 1
                ElseIf Q_strncasecmp(s, "begin", 5) = 0 Then
                    ret = 1
                ElseIf Q_strncasecmp(s, "prespawn", 8) = 0 Then
                    ret = 1
                ElseIf Q_strncasecmp(s, "kick", 4) = 0 Then
                    ret = 1
                ElseIf Q_strncasecmp(s, "ping", 4) = 0 Then
                    ret = 1
                ElseIf Q_strncasecmp(s, "give", 4) = 0 Then
                    ret = 1
                ElseIf Q_strncasecmp(s, "ban", 3) = 0 Then
                    ret = 1
                End If
                If ret = 2 Then
                    Cbuf_InsertText s
                ElseIf ret = 1 Then
                    Cmd_ExecuteString s, src_client
                Else
                    Con_DPrintf "%s tried to %s\n", svs.clients(host_client).name, s
                End If

            Case clc_disconnect
                ' Sys_Printf ("SV_ReadClientMessage: client disconnected\n");
                Exit Function

            Case clc_move
                SV_ReadClientMove svs.clients(host_client).cmd

            Case Else
                Sys_Printf "SV_ReadClientMessage: unknown command char\n"
                Exit Function
            End Select
        Loop
    Loop
End Function

Public Sub SV_RunClients()
    Dim i As Long, blank As usercmd_t

    For i = 0 To svs.maxclients - 1
        host_client = i
        If Not svs.clients(i).active Then GoTo nexti

        sv_player = svs.clients(i).edict

        If Not SV_ReadClientMessage() Then
            SV_DropClient False          ' client misbehaved...
            GoTo nexti
        End If

        If Not svs.clients(i).spawned Then
            ' clear client movement until a new packet is received
            svs.clients(i).cmd = blank
            GoTo nexti
        End If

        ' always pause in single player if in console or menus
        If Not sv.paused And (svs.maxclients > 1 Or key_dest = key_game) Then SV_ClientThink
nexti:
    Next
End Sub
