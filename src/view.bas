Attribute VB_Name = "view"
' view.bas - player eye positioning (view.c)
'
' The view is allowed to move slightly from it's true position for bobbing,
' but if it exceeds 8 pixels linear distance (spherical, not box), the list of
' entities sent from the server may not include everything in the pvs, especially
' when crossing a water boudnary.
'
' The GLQuake versions of V_CalcBlend and V_UpdatePalette are not ported.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' (cv_lcd_x is in quakedef.bas)
Private cv_lcd_yaw As Long

Private cv_scr_ofsx As Long
Private cv_scr_ofsy As Long
Private cv_scr_ofsz As Long

Private cv_cl_rollspeed As Long
Private cv_cl_rollangle As Long

Private cv_cl_bob As Long
Private cv_cl_bobcycle As Long
Private cv_cl_bobup As Long

Private cv_v_kicktime As Long
Private cv_v_kickroll As Long
Private cv_v_kickpitch As Long

Private cv_v_iyaw_cycle As Long
Private cv_v_iroll_cycle As Long
Private cv_v_ipitch_cycle As Long
Private cv_v_iyaw_level As Long
Private cv_v_iroll_level As Long
Private cv_v_ipitch_level As Long

Private cv_v_idlescale As Long

Public cv_crosshair As Long
Private cv_cl_crossx As Long
Private cv_cl_crossy As Long

Private cv_gl_cshiftpercent As Long

Private v_dmg_time As Single, v_dmg_roll As Single, v_dmg_pitch As Single

' C: vec3_t forward, right, up (also used by the lcd_x views)
Private v_forward As vec3_t, v_right As vec3_t, v_up As vec3_t

Private cv_v_centermove As Long
Private cv_v_centerspeed As Long

Private cshift_empty As cshift_t
Private cshift_water As cshift_t
Private cshift_slime As cshift_t
Private cshift_lava As cshift_t

Public cv_v_gamma As Long

Public gammatable(0 To 255) As Byte      ' palette is sent through this

Private oldgammavalue As Single          ' C: static of V_CheckGamma
Private oldyaw As Single, oldpitch As Single   ' C: statics of CalcGunAngle
Private oldz As Single                   ' C: static of V_CalcRefdef

' Used by view and sv_user
Public Function V_CalcRoll(angles As vec3_t, velocity As vec3_t) As Single
    Dim sign As Single, side As Single, value As Single

    AngleVectors angles, v_forward, v_right, v_up
    side = DotProduct(velocity, v_right)
    If side < 0 Then sign = -1 Else sign = 1
    side = Abs(side)

    value = cvars(cv_cl_rollangle).value
    ' if (cl.inwater)
    '     value *= 6;

    If side < cvars(cv_cl_rollspeed).value Then
        side = side * value / cvars(cv_cl_rollspeed).value
    Else
        side = value
    End If

    V_CalcRoll = side * sign
End Function

Private Function V_CalcBob() As Single
    Dim bob As Single, cycle_ As Single

    cycle_ = cl.time - ToInt(cl.time / cvars(cv_cl_bobcycle).value) * cvars(cv_cl_bobcycle).value
    cycle_ = cycle_ / cvars(cv_cl_bobcycle).value
    If cycle_ < cvars(cv_cl_bobup).value Then
        cycle_ = M_PI * cycle_ / cvars(cv_cl_bobup).value
    Else
        cycle_ = M_PI + M_PI * (cycle_ - cvars(cv_cl_bobup).value) / (1# - cvars(cv_cl_bobup).value)
    End If

    ' bob is proportional to velocity in the xy plane
    ' (don't count Z, or jumping messes it up)

    bob = Sqr(cl.velocity.v(0) * cl.velocity.v(0) + cl.velocity.v(1) * cl.velocity.v(1)) * cvars(cv_cl_bob).value
    ' Con_Printf ("speed: %5.1f\n", Length(cl.velocity));
    bob = bob * 0.3 + bob * 0.7 * Sin(cycle_)
    If bob > 4 Then
        bob = 4
    ElseIf bob < -7 Then
        bob = -7
    End If
    V_CalcBob = bob
End Function

'=============================================================================

Public Sub V_StartPitchDrift()
    If cl.laststop = cl.time Then
        Exit Sub                         ' something else is keeping it from drifting
    End If
    If cl.nodrift Or cl.pitchvel = 0 Then
        cl.pitchvel = cvars(cv_v_centerspeed).value
        cl.nodrift = False
        cl.driftmove = 0
    End If
End Sub

Public Sub V_StopPitchDrift()
    cl.laststop = cl.time
    cl.nodrift = True
    cl.pitchvel = 0
End Sub

' Moves the client pitch angle towards cl.idealpitch sent by the server.
'
' If the user is adjusting pitch manually, either with lookup/lookdown,
' mlook and mouse, or klook and keyboard, pitch drifting is constantly stopped.
'
' Drifting is enabled when the center view key is hit, mlook is released and
' lookspring is non 0, or when
Private Sub V_DriftPitch()
    Dim delta As Single, move As Single

    If noclip_anglehack Or Not cl.onground Or cls.demoplayback Then
        cl.driftmove = 0
        cl.pitchvel = 0
        Exit Sub
    End If

    ' don't count small mouse motion
    If cl.nodrift Then
        If Abs(cl.cmd.forwardmove) < cvars(cv_cl_forwardspeed).value Then
            cl.driftmove = 0
        Else
            cl.driftmove = cl.driftmove + host_frametime
        End If

        If cl.driftmove > cvars(cv_v_centermove).value Then V_StartPitchDrift
        Exit Sub
    End If

    delta = cl.idealpitch - cl.viewangles.v(PITCH)

    If delta = 0 Then
        cl.pitchvel = 0
        Exit Sub
    End If

    move = host_frametime * cl.pitchvel
    cl.pitchvel = cl.pitchvel + host_frametime * cvars(cv_v_centerspeed).value

    ' Con_Printf ("move: %f (%f)\n", move, host_frametime);

    If delta > 0 Then
        If move > delta Then
            cl.pitchvel = 0
            move = delta
        End If
        cl.viewangles.v(PITCH) = cl.viewangles.v(PITCH) + move
    ElseIf delta < 0 Then
        If move > -delta Then
            cl.pitchvel = 0
            move = -delta
        End If
        cl.viewangles.v(PITCH) = cl.viewangles.v(PITCH) - move
    End If
End Sub

'==============================================================================
' PALETTE FLASHES

Private Sub SetCshift(c As cshift_t, ByVal r As Long, ByVal g As Long, ByVal b As Long, ByVal percent As Long)
    c.destcolor(0) = r
    c.destcolor(1) = g
    c.destcolor(2) = b
    c.percent = percent
End Sub

Private Sub BuildGammaTable(ByVal g As Single)
    Dim i As Long, inf As Long

    If g = 1 Then
        For i = 0 To 255
            gammatable(i) = i
        Next
        Exit Sub
    End If

    For i = 0 To 255
        inf = ToInt(255 * ((i + 0.5) / 255.5) ^ g + 0.5)
        If inf < 0 Then inf = 0
        If inf > 255 Then inf = 255
        gammatable(i) = inf
    Next
End Sub

Private Function V_CheckGamma() As Boolean
    If cvars(cv_v_gamma).value = oldgammavalue Then Exit Function
    oldgammavalue = cvars(cv_v_gamma).value

    BuildGammaTable cvars(cv_v_gamma).value
    vid.recalc_refdef = 1                ' force a surface cache flush

    V_CheckGamma = True
End Function

Public Sub V_ParseDamage()
    Dim armor As Long, blood As Long, from As vec3_t, i As Long
    Dim forward As vec3_t, right_ As vec3_t, up As vec3_t, ent As Long, side As Single, count As Single

    armor = MSG_ReadByte()
    blood = MSG_ReadByte()
    For i = 0 To 2
        from.v(i) = MSG_ReadCoord()
    Next

    count = blood * 0.5 + armor * 0.5
    If count < 10 Then count = 10

    cl.faceanimtime = cl.time + 0.2      ' but sbar face into pain frame

    With cl.cshifts(CSHIFT_DAMAGE)
        .percent = ToInt(.percent + 3 * count)
        If .percent < 0 Then .percent = 0
        If .percent > 150 Then .percent = 150

        If armor > blood Then
            .destcolor(0) = 200
            .destcolor(1) = 100
            .destcolor(2) = 100
        ElseIf armor Then
            .destcolor(0) = 220
            .destcolor(1) = 50
            .destcolor(2) = 50
        Else
            .destcolor(0) = 255
            .destcolor(1) = 0
            .destcolor(2) = 0
        End If
    End With

    '
    ' calculate view angle kicks
    '
    ent = cl.viewentity

    VectorSubtract from, cl_entities(ent).origin, from
    VectorNormalize from

    AngleVectors cl_entities(ent).angles, forward, right_, up

    side = DotProduct(from, right_)
    v_dmg_roll = count * side * cvars(cv_v_kickroll).value

    side = DotProduct(from, forward)
    v_dmg_pitch = count * side * cvars(cv_v_kickpitch).value

    v_dmg_time = cvars(cv_v_kicktime).value
End Sub

Private Sub V_cshift_f()
    cshift_empty.destcolor(0) = C_atoi(Cmd_Argv(1))
    cshift_empty.destcolor(1) = C_atoi(Cmd_Argv(2))
    cshift_empty.destcolor(2) = C_atoi(Cmd_Argv(3))
    cshift_empty.percent = C_atoi(Cmd_Argv(4))
End Sub

' When you run over an item, the server sends this command
Private Sub V_BonusFlash_f()
    SetCshift cl.cshifts(CSHIFT_BONUS), 215, 186, 69, 50
End Sub

' Underwater, lava, etc each has a color shift
Public Sub V_SetContentsColor(ByVal contents As Long)
    Select Case contents
    Case CONTENTS_EMPTY, CONTENTS_SOLID
        cl.cshifts(CSHIFT_CONTENTS) = cshift_empty
    Case CONTENTS_LAVA
        cl.cshifts(CSHIFT_CONTENTS) = cshift_lava
    Case CONTENTS_SLIME
        cl.cshifts(CSHIFT_CONTENTS) = cshift_slime
    Case Else
        cl.cshifts(CSHIFT_CONTENTS) = cshift_water
    End Select
End Sub

Private Sub V_CalcPowerupCshift()
    If cl.items And IT_QUAD Then
        SetCshift cl.cshifts(CSHIFT_POWERUP), 0, 0, 255, 30
    ElseIf cl.items And IT_SUIT Then
        SetCshift cl.cshifts(CSHIFT_POWERUP), 0, 255, 0, 20
    ElseIf cl.items And IT_INVISIBILITY Then
        SetCshift cl.cshifts(CSHIFT_POWERUP), 100, 100, 100, 100
    ElseIf cl.items And IT_INVULNERABILITY Then
        SetCshift cl.cshifts(CSHIFT_POWERUP), 255, 255, 0, 30
    Else
        cl.cshifts(CSHIFT_POWERUP).percent = 0
    End If
End Sub

Public Sub V_UpdatePalette()
    Dim i As Long, j As Long, new_ As Boolean, pal(0 To 767) As Byte
    Dim r As Long, g As Long, b As Long, force As Boolean

    V_CalcPowerupCshift

    new_ = False

    For i = 0 To NUM_CSHIFTS - 1
        If cl.cshifts(i).percent <> cl.prev_cshifts(i).percent Then
            new_ = True
            cl.prev_cshifts(i).percent = cl.cshifts(i).percent
        End If
        For j = 0 To 2
            If cl.cshifts(i).destcolor(j) <> cl.prev_cshifts(i).destcolor(j) Then
                new_ = True
                cl.prev_cshifts(i).destcolor(j) = cl.cshifts(i).destcolor(j)
            End If
        Next
    Next

    ' drop the damage value
    cl.cshifts(CSHIFT_DAMAGE).percent = ToInt(cl.cshifts(CSHIFT_DAMAGE).percent - host_frametime * 150)
    If cl.cshifts(CSHIFT_DAMAGE).percent <= 0 Then cl.cshifts(CSHIFT_DAMAGE).percent = 0

    ' drop the bonus value
    cl.cshifts(CSHIFT_BONUS).percent = ToInt(cl.cshifts(CSHIFT_BONUS).percent - host_frametime * 100)
    If cl.cshifts(CSHIFT_BONUS).percent <= 0 Then cl.cshifts(CSHIFT_BONUS).percent = 0

    force = V_CheckGamma()
    If Not new_ And Not force Then Exit Sub

    For i = 0 To 255
        r = host_basepal(i * 3)
        g = host_basepal(i * 3 + 1)
        b = host_basepal(i * 3 + 2)

        ' (C's >> 8 of a negative value floors: (x And &HFFFFFF00) \ 256)
        For j = 0 To NUM_CSHIFTS - 1
            r = r + ((cl.cshifts(j).percent * (cl.cshifts(j).destcolor(0) - r)) And &HFFFFFF00) \ 256
            g = g + ((cl.cshifts(j).percent * (cl.cshifts(j).destcolor(1) - g)) And &HFFFFFF00) \ 256
            b = b + ((cl.cshifts(j).percent * (cl.cshifts(j).destcolor(2) - b)) And &HFFFFFF00) \ 256
        Next

        pal(i * 3) = gammatable(r)
        pal(i * 3 + 1) = gammatable(g)
        pal(i * 3 + 2) = gammatable(b)
    Next

    VID_ShiftPalette pal
End Sub

'==============================================================================
' VIEW RENDERING

Private Function angledelta(ByVal a As Single) As Single
    a = anglemod(a)
    If a > 180 Then a = a - 360
    angledelta = a
End Function

Private Sub CalcGunAngle()
    Dim yaw_ As Single, pitch_ As Single, move As Single

    yaw_ = r_refdef.viewangles.v(YAW)
    pitch_ = -r_refdef.viewangles.v(PITCH)

    yaw_ = angledelta(yaw_ - r_refdef.viewangles.v(YAW)) * 0.4
    If yaw_ > 10 Then yaw_ = 10
    If yaw_ < -10 Then yaw_ = -10
    pitch_ = angledelta(-pitch_ - r_refdef.viewangles.v(PITCH)) * 0.4
    If pitch_ > 10 Then pitch_ = 10
    If pitch_ < -10 Then pitch_ = -10
    move = host_frametime * 20
    If yaw_ > oldyaw Then
        If oldyaw + move < yaw_ Then yaw_ = oldyaw + move
    Else
        If oldyaw - move > yaw_ Then yaw_ = oldyaw - move
    End If

    If pitch_ > oldpitch Then
        If oldpitch + move < pitch_ Then pitch_ = oldpitch + move
    Else
        If oldpitch - move > pitch_ Then pitch_ = oldpitch - move
    End If

    oldyaw = yaw_
    oldpitch = pitch_

    With cl_entities(CL_VIEWENT)
        .angles.v(YAW) = r_refdef.viewangles.v(YAW) + yaw_
        .angles.v(PITCH) = -(r_refdef.viewangles.v(PITCH) + pitch_)

        .angles.v(ROLL) = .angles.v(ROLL) - cvars(cv_v_idlescale).value * Sin(cl.time * cvars(cv_v_iroll_cycle).value) * cvars(cv_v_iroll_level).value
        .angles.v(PITCH) = .angles.v(PITCH) - cvars(cv_v_idlescale).value * Sin(cl.time * cvars(cv_v_ipitch_cycle).value) * cvars(cv_v_ipitch_level).value
        .angles.v(YAW) = .angles.v(YAW) - cvars(cv_v_idlescale).value * Sin(cl.time * cvars(cv_v_iyaw_cycle).value) * cvars(cv_v_iyaw_level).value
    End With
End Sub

Private Sub V_BoundOffsets()
    Dim ent As Long

    ent = cl.viewentity

    ' absolutely bound refresh reletive to entity clipping hull
    ' so the view can never be inside a solid wall

    With cl_entities(ent)
        If r_refdef.vieworg.v(0) < .origin.v(0) - 14 Then
            r_refdef.vieworg.v(0) = .origin.v(0) - 14
        ElseIf r_refdef.vieworg.v(0) > .origin.v(0) + 14 Then
            r_refdef.vieworg.v(0) = .origin.v(0) + 14
        End If
        If r_refdef.vieworg.v(1) < .origin.v(1) - 14 Then
            r_refdef.vieworg.v(1) = .origin.v(1) - 14
        ElseIf r_refdef.vieworg.v(1) > .origin.v(1) + 14 Then
            r_refdef.vieworg.v(1) = .origin.v(1) + 14
        End If
        If r_refdef.vieworg.v(2) < .origin.v(2) - 22 Then
            r_refdef.vieworg.v(2) = .origin.v(2) - 22
        ElseIf r_refdef.vieworg.v(2) > .origin.v(2) + 30 Then
            r_refdef.vieworg.v(2) = .origin.v(2) + 30
        End If
    End With
End Sub

' Idle swaying
Private Sub V_AddIdle()
    r_refdef.viewangles.v(ROLL) = r_refdef.viewangles.v(ROLL) + cvars(cv_v_idlescale).value * Sin(cl.time * cvars(cv_v_iroll_cycle).value) * cvars(cv_v_iroll_level).value
    r_refdef.viewangles.v(PITCH) = r_refdef.viewangles.v(PITCH) + cvars(cv_v_idlescale).value * Sin(cl.time * cvars(cv_v_ipitch_cycle).value) * cvars(cv_v_ipitch_level).value
    r_refdef.viewangles.v(YAW) = r_refdef.viewangles.v(YAW) + cvars(cv_v_idlescale).value * Sin(cl.time * cvars(cv_v_iyaw_cycle).value) * cvars(cv_v_iyaw_level).value
End Sub

' Roll is induced by movement and damage
Private Sub V_CalcViewRoll()
    Dim side As Single

    side = V_CalcRoll(cl_entities(cl.viewentity).angles, cl.velocity)
    r_refdef.viewangles.v(ROLL) = r_refdef.viewangles.v(ROLL) + side

    If v_dmg_time > 0 Then
        r_refdef.viewangles.v(ROLL) = r_refdef.viewangles.v(ROLL) + v_dmg_time / cvars(cv_v_kicktime).value * v_dmg_roll
        r_refdef.viewangles.v(PITCH) = r_refdef.viewangles.v(PITCH) + v_dmg_time / cvars(cv_v_kicktime).value * v_dmg_pitch
        v_dmg_time = v_dmg_time - host_frametime
    End If

    If cl.stats(STAT_HEALTH) <= 0 Then
        r_refdef.viewangles.v(ROLL) = 80 ' dead view angle
        Exit Sub
    End If
End Sub

Private Sub V_CalcIntermissionRefdef()
    Dim ent As Long, view_ As Long, old As Single

    ' ent is the player model (visible when out of body)
    ent = cl.viewentity
    ' view is the weapon model (only visible from inside body)
    view_ = CL_VIEWENT

    r_refdef.vieworg = cl_entities(ent).origin
    r_refdef.viewangles = cl_entities(ent).angles
    cl_entities(view_).model = 0

    ' allways idle in intermission
    old = cvars(cv_v_idlescale).value
    cvars(cv_v_idlescale).value = 1
    V_AddIdle
    cvars(cv_v_idlescale).value = old
End Sub

Private Sub V_CalcRefdef()
    Dim ent As Long, view_ As Long, i As Long, forward As vec3_t, right_ As vec3_t, up As vec3_t
    Dim angles As vec3_t, bob As Single, steptime As Single

    V_DriftPitch

    ' ent is the player model (visible when out of body)
    ent = cl.viewentity
    ' view is the weapon model (only visible from inside body)
    view_ = CL_VIEWENT

    ' transform the view offset by the model's matrix to get the offset from
    ' model origin for the view
    cl_entities(ent).angles.v(YAW) = cl.viewangles.v(YAW)       ' the model should face
                                                                 ' the view dir
    cl_entities(ent).angles.v(PITCH) = -cl.viewangles.v(PITCH)  ' the model should face
                                                                 ' the view dir

    bob = V_CalcBob()

    ' refresh position
    r_refdef.vieworg = cl_entities(ent).origin
    r_refdef.vieworg.v(2) = r_refdef.vieworg.v(2) + cl.viewheight + bob

    ' never let it sit exactly on a node line, because a water plane can
    ' dissapear when viewed with the eye exactly on it.
    ' the server protocol only specifies to 1/16 pixel, so add 1/32 in each axis
    r_refdef.vieworg.v(0) = r_refdef.vieworg.v(0) + 1# / 32
    r_refdef.vieworg.v(1) = r_refdef.vieworg.v(1) + 1# / 32
    r_refdef.vieworg.v(2) = r_refdef.vieworg.v(2) + 1# / 32

    r_refdef.viewangles = cl.viewangles
    V_CalcViewRoll
    V_AddIdle

    ' offsets
    angles.v(PITCH) = -cl_entities(ent).angles.v(PITCH)   ' because entity pitches are
                                                           '  actually backward
    angles.v(YAW) = cl_entities(ent).angles.v(YAW)
    angles.v(ROLL) = cl_entities(ent).angles.v(ROLL)

    AngleVectors angles, forward, right_, up

    For i = 0 To 2
        r_refdef.vieworg.v(i) = r_refdef.vieworg.v(i) + cvars(cv_scr_ofsx).value * forward.v(i) _
            + cvars(cv_scr_ofsy).value * right_.v(i) _
            + cvars(cv_scr_ofsz).value * up.v(i)
    Next

    V_BoundOffsets

    ' set up gun position
    cl_entities(view_).angles = cl.viewangles

    CalcGunAngle

    With cl_entities(view_)
        .origin = cl_entities(ent).origin
        .origin.v(2) = .origin.v(2) + cl.viewheight

        For i = 0 To 2
            .origin.v(i) = .origin.v(i) + forward.v(i) * bob * 0.4
            ' view->origin[i] += right[i]*bob*0.4;
            ' view->origin[i] += up[i]*bob*0.8;
        Next
        .origin.v(2) = .origin.v(2) + bob

        ' fudge position around to keep amount of weapon visible
        ' roughly equal with different FOV

        If cvars(cv_scr_viewsize).value = 110 Then
            .origin.v(2) = .origin.v(2) + 1
        ElseIf cvars(cv_scr_viewsize).value = 100 Then
            .origin.v(2) = .origin.v(2) + 2
        ElseIf cvars(cv_scr_viewsize).value = 90 Then
            .origin.v(2) = .origin.v(2) + 1
        ElseIf cvars(cv_scr_viewsize).value = 80 Then
            .origin.v(2) = .origin.v(2) + 0.5
        End If

        .model = cl.model_precache(cl.stats(STAT_WEAPON))
        .frame = cl.stats(STAT_WEAPONFRAME)
        .colormap = 0                    ' vid.colormap
    End With

    ' set up the refresh position
    VectorAdd r_refdef.viewangles, cl.punchangle, r_refdef.viewangles

    ' smooth out stair step ups
    If cl.onground And cl_entities(ent).origin.v(2) - oldz > 0 Then
        steptime = cl.time - cl.oldtime
        If steptime < 0 Then
            ' FIXME I_Error ("steptime < 0");
            steptime = 0
        End If

        oldz = oldz + steptime * 80
        If oldz > cl_entities(ent).origin.v(2) Then oldz = cl_entities(ent).origin.v(2)
        If cl_entities(ent).origin.v(2) - oldz > 12 Then oldz = cl_entities(ent).origin.v(2) - 12
        r_refdef.vieworg.v(2) = r_refdef.vieworg.v(2) + oldz - cl_entities(ent).origin.v(2)
        cl_entities(view_).origin.v(2) = cl_entities(view_).origin.v(2) + oldz - cl_entities(ent).origin.v(2)
    Else
        oldz = cl_entities(ent).origin.v(2)
    End If

    If cvars(cv_chase_active).value <> 0 Then Chase_Update
End Sub

' The player's clipping box goes from (-16 -16 -24) to (16 16 32) from
' the entity origin, so any view position inside that will be valid
Public Sub V_RenderView()
    Dim i As Long

    If con_forcedup Then Exit Sub

    ' don't allow cheats in multiplayer
    If cl.maxclients > 1 Then
        Cvar_Set "scr_ofsx", "0"
        Cvar_Set "scr_ofsy", "0"
        Cvar_Set "scr_ofsz", "0"
    End If

    If cl.intermission Then
        ' intermission / finale rendering
        V_CalcIntermissionRefdef
    Else
        If Not cl.paused Then V_CalcRefdef   ' /* && (sv.maxclients > 1 || key_dest == key_game) */
    End If

    R_PushDlights

    If cvars(cv_lcd_x).value <> 0 Then
        '
        ' render two interleaved views
        '
        vid.rowbytes = vid.rowbytes * 2
        vid.aspect = vid.aspect * 0.5

        r_refdef.viewangles.v(YAW) = r_refdef.viewangles.v(YAW) - cvars(cv_lcd_yaw).value
        For i = 0 To 2
            r_refdef.vieworg.v(i) = r_refdef.vieworg.v(i) - v_right.v(i) * cvars(cv_lcd_x).value
        Next
        R_RenderView

        vid_bufferofs = vid_bufferofs + vid.rowbytes \ 2   ' C: vid.buffer += vid.rowbytes>>1

        R_PushDlights

        r_refdef.viewangles.v(YAW) = r_refdef.viewangles.v(YAW) + cvars(cv_lcd_yaw).value * 2
        For i = 0 To 2
            r_refdef.vieworg.v(i) = r_refdef.vieworg.v(i) + 2 * v_right.v(i) * cvars(cv_lcd_x).value
        Next
        R_RenderView

        vid_bufferofs = vid_bufferofs - vid.rowbytes \ 2

        r_refdef.vrect.height = r_refdef.vrect.height * 2

        vid.rowbytes = vid.rowbytes \ 2
        vid.aspect = vid.aspect * 2
    Else
        R_RenderView
    End If

    If cvars(cv_crosshair).value <> 0 Then
        Draw_Character ToInt(scr_vrect.x + scr_vrect.width \ 2 + cvars(cv_cl_crossx).value), _
            ToInt(scr_vrect.y + scr_vrect.height \ 2 + cvars(cv_cl_crossy).value), AscW("+")
    End If
End Sub

'============================================================================

Public Sub V_Init()
    SetCshift cshift_empty, 130, 80, 50, 0
    SetCshift cshift_water, 130, 80, 50, 128
    SetCshift cshift_slime, 0, 25, 5, 150
    SetCshift cshift_lava, 255, 80, 0, 150

    Cmd_AddCommand "v_cshift", XC_V_cshift_f
    Cmd_AddCommand "bf", XC_V_BonusFlash_f
    Cmd_AddCommand "centerview", XC_V_StartPitchDrift

    cv_lcd_x = Cvar_RegisterVariable("lcd_x", "0")
    cv_lcd_yaw = Cvar_RegisterVariable("lcd_yaw", "0")

    cv_v_centermove = Cvar_RegisterVariable("v_centermove", "0.15")
    cv_v_centerspeed = Cvar_RegisterVariable("v_centerspeed", "500")

    cv_v_iyaw_cycle = Cvar_RegisterVariable("v_iyaw_cycle", "2")
    cv_v_iroll_cycle = Cvar_RegisterVariable("v_iroll_cycle", "0.5")
    cv_v_ipitch_cycle = Cvar_RegisterVariable("v_ipitch_cycle", "1")
    cv_v_iyaw_level = Cvar_RegisterVariable("v_iyaw_level", "0.3")
    cv_v_iroll_level = Cvar_RegisterVariable("v_iroll_level", "0.1")
    cv_v_ipitch_level = Cvar_RegisterVariable("v_ipitch_level", "0.3")

    cv_v_idlescale = Cvar_RegisterVariable("v_idlescale", "0")
    cv_crosshair = Cvar_RegisterVariable("crosshair", "0", True)
    cv_cl_crossx = Cvar_RegisterVariable("cl_crossx", "0")
    cv_cl_crossy = Cvar_RegisterVariable("cl_crossy", "0")
    cv_gl_cshiftpercent = Cvar_RegisterVariable("gl_cshiftpercent", "100")

    cv_scr_ofsx = Cvar_RegisterVariable("scr_ofsx", "0")
    cv_scr_ofsy = Cvar_RegisterVariable("scr_ofsy", "0")
    cv_scr_ofsz = Cvar_RegisterVariable("scr_ofsz", "0")
    cv_cl_rollspeed = Cvar_RegisterVariable("cl_rollspeed", "200")
    cv_cl_rollangle = Cvar_RegisterVariable("cl_rollangle", "2.0")
    cv_cl_bob = Cvar_RegisterVariable("cl_bob", "0.02")
    cv_cl_bobcycle = Cvar_RegisterVariable("cl_bobcycle", "0.6")
    cv_cl_bobup = Cvar_RegisterVariable("cl_bobup", "0.5")

    cv_v_kicktime = Cvar_RegisterVariable("v_kicktime", "0.5")
    cv_v_kickroll = Cvar_RegisterVariable("v_kickroll", "0.6")
    cv_v_kickpitch = Cvar_RegisterVariable("v_kickpitch", "0.6")

    BuildGammaTable 1#                   ' no gamma yet
    cv_v_gamma = Cvar_RegisterVariable("gamma", "1", True)
End Sub

' (QuakeVB) the console commands of this module, for Cmd_Call
Public Sub V_CallCommand(ByVal fn As Long)
    Select Case fn
    Case XC_V_cshift_f: V_cshift_f
    Case XC_V_BonusFlash_f: V_BonusFlash_f
    Case XC_V_StartPitchDrift: V_StartPitchDrift
    Case Else: Sys_Error "V_CallCommand: bad function %i", fn
    End Select
End Sub
