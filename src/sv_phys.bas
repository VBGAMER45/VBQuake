Attribute VB_Name = "sv_phys"
' sv_phys.bas - entity physics (sv_phys.c)
'
' pushmove objects do not obey gravity, and do not interact with each other or
' trigger fields, but block normal movement and push normal objects when they move.
'
' onground is set for toss objects when they come to a complete rest.  it is set
' for steping or walking objects
'
' doors, plats, etc are SOLID_BSP, and MOVETYPE_PUSH
' bonus items are SOLID_TRIGGER touch, and MOVETYPE_TOSS
' corpses are SOLID_NOT and MOVETYPE_TOSS
' crates are SOLID_BBOX and MOVETYPE_TOSS
' walking monsters are SOLID_SLIDEBOX and MOVETYPE_STEP
' flying/floating monsters are SOLID_SLIDEBOX and MOVETYPE_FLY
'
' solid_edge items only clip against bsp models.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public cv_sv_friction As Long
Public cv_sv_stopspeed As Long
Public cv_sv_gravity As Long
Public cv_sv_maxvelocity As Long
Public cv_sv_nostep As Long

Private Const MOVE_EPSILON As Single = 0.01
Private Const STOP_EPSILON As Single = 0.1
Private Const MAX_CLIP_PLANES As Long = 5
Private Const STEPSIZE As Long = 18

Private Sub SV_CheckAllEnts()
    Dim e As Long

    ' see if any solid entities are inside the final position
    For e = 1 To sv.num_edicts - 1
        If Not edicts(e).free Then
            If ev(e).movetype <> MOVETYPE_PUSH And ev(e).movetype <> MOVETYPE_NONE And ev(e).movetype <> MOVETYPE_NOCLIP Then
                If SV_TestEntityPosition(e) <> -1 Then Con_Printf "entity in invalid position\n"
            End If
        End If
    Next
End Sub

Public Sub SV_CheckVelocity(ByVal ent As Long)
    Dim i As Long

    '
    ' bound velocity
    '
    For i = 0 To 2
        If IS_NAN(ev(ent).velocity.v(i)) Then
            Con_Printf "Got a NaN velocity on %s\n", PR_GetString(ev(ent).classname)
            ev(ent).velocity.v(i) = 0
        End If
        If IS_NAN(ev(ent).origin.v(i)) Then
            Con_Printf "Got a NaN origin on %s\n", PR_GetString(ev(ent).classname)
            ev(ent).origin.v(i) = 0
        End If
        If ev(ent).velocity.v(i) > cvars(cv_sv_maxvelocity).value Then
            ev(ent).velocity.v(i) = cvars(cv_sv_maxvelocity).value
        ElseIf ev(ent).velocity.v(i) < -cvars(cv_sv_maxvelocity).value Then
            ev(ent).velocity.v(i) = -cvars(cv_sv_maxvelocity).value
        End If
    Next
End Sub

' Runs thinking code if time.  There is some play in the exact time the think
' function will be called, because it is called before any movement is done
' in a frame.  Not used for pushmove objects, because they must be exact.
' Returns false if the entity removed itself.
Public Function SV_RunThink(ByVal ent As Long) As Boolean
    Dim thinktime As Single

    thinktime = ev(ent).nextthink
    If thinktime <= 0 Or thinktime > sv.time + host_frametime Then
        SV_RunThink = True
        Exit Function
    End If

    If thinktime < sv.time Then thinktime = sv.time   ' don't let things stay in the past.
                                                      ' it is possible to start that way
                                                      ' by a trigger with a local time.
    ev(ent).nextthink = 0
    pr_global_struct(0).time = thinktime
    pr_global_struct(0).self = EDICT_TO_PROG(ent)
    pr_global_struct(0).other = EDICT_TO_PROG(0)
    PR_ExecuteProgram ev(ent).think

    SV_RunThink = Not edicts(ent).free
End Function

' Two entities have touched, so run their touch functions
Private Sub SV_Impact(ByVal e1 As Long, ByVal e2 As Long)
    Dim old_self As Long, old_other As Long

    old_self = pr_global_struct(0).self
    old_other = pr_global_struct(0).other

    pr_global_struct(0).time = sv.time
    If ev(e1).touch <> 0 And ev(e1).solid <> SOLID_NOT Then
        pr_global_struct(0).self = EDICT_TO_PROG(e1)
        pr_global_struct(0).other = EDICT_TO_PROG(e2)
        PR_ExecuteProgram ev(e1).touch
    End If

    If ev(e2).touch <> 0 And ev(e2).solid <> SOLID_NOT Then
        pr_global_struct(0).self = EDICT_TO_PROG(e2)
        pr_global_struct(0).other = EDICT_TO_PROG(e1)
        PR_ExecuteProgram ev(e2).touch
    End If

    pr_global_struct(0).self = old_self
    pr_global_struct(0).other = old_other
End Sub

' Slide off of the impacting object
' returns the blocked flags (1 = floor, 2 = step / wall)
Private Function ClipVelocity(in_ As vec3_t, normal As vec3_t, out As vec3_t, ByVal overbounce As Single) As Long
    Dim backoff As Single, change As Single, i As Long, blocked As Long

    blocked = 0
    If normal.v(2) > 0 Then blocked = blocked Or 1   ' floor
    If normal.v(2) = 0 Then blocked = blocked Or 2   ' step

    backoff = DotProduct(in_, normal) * overbounce

    For i = 0 To 2
        change = normal.v(i) * backoff
        out.v(i) = in_.v(i) - change
        If out.v(i) > -STOP_EPSILON And out.v(i) < STOP_EPSILON Then out.v(i) = 0
    Next

    ClipVelocity = blocked
End Function

' The basic solid body movement clip that slides along multiple planes
' Returns the clipflags if the velocity was modified (hit something solid)
' 1 = floor
' 2 = wall / step
' 4 = dead stop
' If steptrace is not NULL, the trace of any vertical wall hit will be stored
' (QuakeVB: has_steptrace False stands for NULL)
Private Function SV_FlyMove(ByVal ent As Long, ByVal time_ As Single, steptrace As trace_t, ByVal has_steptrace As Boolean) As Long
    Dim bumpcount As Long, numbumps As Long, dir_ As vec3_t, d As Single, numplanes As Long
    Dim planes(0 To MAX_CLIP_PLANES - 1) As vec3_t, primal_velocity As vec3_t, original_velocity As vec3_t
    Dim new_velocity As vec3_t, i As Long, j As Long, trace As trace_t, end_ As vec3_t
    Dim time_left As Single, blocked As Long

    numbumps = 4

    blocked = 0
    original_velocity = ev(ent).velocity
    primal_velocity = ev(ent).velocity
    numplanes = 0

    time_left = time_

    For bumpcount = 0 To numbumps - 1
        If ev(ent).velocity.v(0) = 0 And ev(ent).velocity.v(1) = 0 And ev(ent).velocity.v(2) = 0 Then Exit For

        For i = 0 To 2
            end_.v(i) = ev(ent).origin.v(i) + time_left * ev(ent).velocity.v(i)
        Next

        trace = SV_Move(ev(ent).origin, ev(ent).mins, ev(ent).maxs, end_, 0, ent)

        If trace.allsolid Then
            ' entity is trapped in another solid
            ev(ent).velocity = vec3_origin
            SV_FlyMove = 3
            Exit Function
        End If

        If trace.fraction > 0 Then
            ' actually covered some distance
            ev(ent).origin = trace.endpos
            original_velocity = ev(ent).velocity
            numplanes = 0
        End If

        If trace.fraction = 1 Then Exit For      ' moved the entire distance

        If trace.ent = -1 Then Sys_Error "SV_FlyMove: !trace.ent"

        If trace.plane.normal.v(2) > 0.7 Then
            blocked = blocked Or 1               ' floor
            If ev(trace.ent).solid = SOLID_BSP Then
                ev(ent).flags = ToInt(ev(ent).flags) Or FL_ONGROUND
                ev(ent).groundentity = EDICT_TO_PROG(trace.ent)
            End If
        End If
        If trace.plane.normal.v(2) = 0 Then
            blocked = blocked Or 2               ' step
            If has_steptrace Then steptrace = trace   ' save for player extrafriction
        End If

        '
        ' run the impact function
        '
        SV_Impact ent, trace.ent
        If edicts(ent).free Then Exit For        ' removed by the impact function

        time_left = time_left - time_left * trace.fraction

        ' cliped to another plane
        If numplanes >= MAX_CLIP_PLANES Then
            ' this shouldn't really happen
            ev(ent).velocity = vec3_origin
            SV_FlyMove = 3
            Exit Function
        End If

        planes(numplanes) = trace.plane.normal
        numplanes = numplanes + 1

        '
        ' modify original_velocity so it parallels all of the clip planes
        '
        For i = 0 To numplanes - 1
            ClipVelocity original_velocity, planes(i), new_velocity, 1
            For j = 0 To numplanes - 1
                If j <> i Then
                    If DotProduct(new_velocity, planes(j)) < 0 Then Exit For   ' not ok
                End If
            Next
            If j = numplanes Then Exit For
        Next

        If i <> numplanes Then
            ' go along this plane
            ev(ent).velocity = new_velocity
        Else
            ' go along the crease
            If numplanes <> 2 Then
                ev(ent).velocity = vec3_origin
                SV_FlyMove = 7
                Exit Function
            End If
            CrossProduct planes(0), planes(1), dir_
            d = DotProduct(dir_, ev(ent).velocity)
            VectorScale dir_, d, ev(ent).velocity
        End If

        '
        ' if original velocity is against the original velocity, stop dead
        ' to avoid tiny occilations in sloping corners
        '
        If DotProduct(ev(ent).velocity, primal_velocity) <= 0 Then
            ev(ent).velocity = vec3_origin
            SV_FlyMove = blocked
            Exit Function
        End If
    Next

    SV_FlyMove = blocked
End Function

Private Sub SV_AddGravity(ByVal ent As Long)
    Dim ent_gravity As Single, val As Long

    val = GetEdictFieldValue(ent, "gravity")
    ent_gravity = 1#
    If val <> -1 Then
        If ed_f(val) <> 0 Then ent_gravity = ed_f(val)
    End If
    ev(ent).velocity.v(2) = ev(ent).velocity.v(2) - ent_gravity * cvars(cv_sv_gravity).value * host_frametime
End Sub

'===============================================================================
' PUSHMOVE

' Does not change the entities velocity at all
Private Function SV_PushEntity(ByVal ent As Long, push As vec3_t) As trace_t
    Dim trace As trace_t, end_ As vec3_t

    VectorAdd ev(ent).origin, push, end_

    If ev(ent).movetype = MOVETYPE_FLYMISSILE Then
        trace = SV_Move(ev(ent).origin, ev(ent).mins, ev(ent).maxs, end_, MOVE_MISSILE, ent)
    ElseIf ev(ent).solid = SOLID_TRIGGER Or ev(ent).solid = SOLID_NOT Then
        ' only clip against bmodels
        trace = SV_Move(ev(ent).origin, ev(ent).mins, ev(ent).maxs, end_, MOVE_NOMONSTERS, ent)
    Else
        trace = SV_Move(ev(ent).origin, ev(ent).mins, ev(ent).maxs, end_, MOVE_NORMAL, ent)
    End If

    ev(ent).origin = trace.endpos
    SV_LinkEdict ent, True

    If trace.ent <> -1 Then SV_Impact ent, trace.ent

    SV_PushEntity = trace
End Function

Private Sub SV_PushMove(ByVal pusher As Long, ByVal movetime As Single)
    Dim i As Long, e As Long, check As Long, block As Long, mins As vec3_t, maxs As vec3_t, move As vec3_t
    Dim entorig As vec3_t, pushorig As vec3_t, num_moved As Long
    Dim moved_edict(0 To MAX_EDICTS - 1) As Long, moved_from(0 To MAX_EDICTS - 1) As vec3_t

    If ev(pusher).velocity.v(0) = 0 And ev(pusher).velocity.v(1) = 0 And ev(pusher).velocity.v(2) = 0 Then
        ev(pusher).ltime = ev(pusher).ltime + movetime
        Exit Sub
    End If

    For i = 0 To 2
        move.v(i) = ev(pusher).velocity.v(i) * movetime
        mins.v(i) = ev(pusher).absmin.v(i) + move.v(i)
        maxs.v(i) = ev(pusher).absmax.v(i) + move.v(i)
    Next

    pushorig = ev(pusher).origin

    ' move the pusher to it's final position

    VectorAdd ev(pusher).origin, move, ev(pusher).origin
    ev(pusher).ltime = ev(pusher).ltime + movetime
    SV_LinkEdict pusher, False

    ' see if any solid entities are inside the final position
    num_moved = 0
    For e = 1 To sv.num_edicts - 1
        check = e
        If edicts(check).free Then GoTo nexte
        If ev(check).movetype = MOVETYPE_PUSH Or ev(check).movetype = MOVETYPE_NONE Or ev(check).movetype = MOVETYPE_NOCLIP Then GoTo nexte

        ' if the entity is standing on the pusher, it will definately be moved
        If Not ((ToInt(ev(check).flags) And FL_ONGROUND) <> 0 And PROG_TO_EDICT(ev(check).groundentity) = pusher) Then
            If ev(check).absmin.v(0) >= maxs.v(0) _
            Or ev(check).absmin.v(1) >= maxs.v(1) _
            Or ev(check).absmin.v(2) >= maxs.v(2) _
            Or ev(check).absmax.v(0) <= mins.v(0) _
            Or ev(check).absmax.v(1) <= mins.v(1) _
            Or ev(check).absmax.v(2) <= mins.v(2) Then GoTo nexte

            ' see if the ent's bbox is inside the pusher's final position
            If SV_TestEntityPosition(check) = -1 Then GoTo nexte
        End If

        ' remove the onground flag for non-players
        If ev(check).movetype <> MOVETYPE_WALK Then ev(check).flags = ToInt(ev(check).flags) And Not FL_ONGROUND

        entorig = ev(check).origin
        moved_from(num_moved) = ev(check).origin
        moved_edict(num_moved) = check
        num_moved = num_moved + 1

        ' try moving the contacted entity
        ev(pusher).solid = SOLID_NOT
        SV_PushEntity check, move
        ev(pusher).solid = SOLID_BSP

        ' if it is still inside the pusher, block
        block = SV_TestEntityPosition(check)
        If block <> -1 Then
            ' fail the move
            If ev(check).mins.v(0) = ev(check).maxs.v(0) Then GoTo nexte
            If ev(check).solid = SOLID_NOT Or ev(check).solid = SOLID_TRIGGER Then
                ' corpse
                ev(check).mins.v(0) = 0
                ev(check).mins.v(1) = 0
                ev(check).maxs = ev(check).mins
                GoTo nexte
            End If

            ev(check).origin = entorig
            SV_LinkEdict check, True

            ev(pusher).origin = pushorig
            SV_LinkEdict pusher, False
            ev(pusher).ltime = ev(pusher).ltime - movetime

            ' if the pusher has a "blocked" function, call it
            ' otherwise, just stay in place until the obstacle is gone
            If ev(pusher).blocked Then
                pr_global_struct(0).self = EDICT_TO_PROG(pusher)
                pr_global_struct(0).other = EDICT_TO_PROG(check)
                PR_ExecuteProgram ev(pusher).blocked
            End If

            ' move back any entities we already moved
            For i = 0 To num_moved - 1
                ev(moved_edict(i)).origin = moved_from(i)
                SV_LinkEdict moved_edict(i), False
            Next
            Exit Sub
        End If
nexte:
    Next
End Sub

Private Sub SV_Physics_Pusher(ByVal ent As Long)
    Dim thinktime As Single, oldltime As Single, movetime As Single

    oldltime = ev(ent).ltime

    thinktime = ev(ent).nextthink
    If thinktime < ev(ent).ltime + host_frametime Then
        movetime = thinktime - ev(ent).ltime
        If movetime < 0 Then movetime = 0
    Else
        movetime = host_frametime
    End If

    If movetime <> 0 Then
        SV_PushMove ent, movetime        ' advances ent->v.ltime if not blocked
    End If

    If thinktime > oldltime And thinktime <= ev(ent).ltime Then
        ev(ent).nextthink = 0
        pr_global_struct(0).time = sv.time
        pr_global_struct(0).self = EDICT_TO_PROG(ent)
        pr_global_struct(0).other = EDICT_TO_PROG(0)
        PR_ExecuteProgram ev(ent).think
        If edicts(ent).free Then Exit Sub
    End If
End Sub

'===============================================================================
' CLIENT MOVEMENT

' This is a big hack to try and fix the rare case of getting stuck in the world
' clipping hull.
Private Sub SV_CheckStuck(ByVal ent As Long)
    Dim i As Long, j As Long, z As Long, org As vec3_t

    If SV_TestEntityPosition(ent) = -1 Then
        ev(ent).oldorigin = ev(ent).origin
        Exit Sub
    End If

    org = ev(ent).origin
    ev(ent).origin = ev(ent).oldorigin
    If SV_TestEntityPosition(ent) = -1 Then
        Con_DPrintf "Unstuck.\n"
        SV_LinkEdict ent, True
        Exit Sub
    End If

    For z = 0 To 17
        For i = -1 To 1
            For j = -1 To 1
                ev(ent).origin.v(0) = org.v(0) + i
                ev(ent).origin.v(1) = org.v(1) + j
                ev(ent).origin.v(2) = org.v(2) + z
                If SV_TestEntityPosition(ent) = -1 Then
                    Con_DPrintf "Unstuck.\n"
                    SV_LinkEdict ent, True
                    Exit Sub
                End If
            Next
        Next
    Next

    ev(ent).origin = org
    Con_DPrintf "player is stuck.\n"
End Sub

Private Function SV_CheckWater(ByVal ent As Long) As Boolean
    Dim point As vec3_t, cont As Long

    point.v(0) = ev(ent).origin.v(0)
    point.v(1) = ev(ent).origin.v(1)
    point.v(2) = ev(ent).origin.v(2) + ev(ent).mins.v(2) + 1

    ev(ent).waterlevel = 0
    ev(ent).watertype = CONTENTS_EMPTY
    cont = SV_PointContents(point)
    If cont <= CONTENTS_WATER Then
        ev(ent).watertype = cont
        ev(ent).waterlevel = 1
        point.v(2) = ev(ent).origin.v(2) + (ev(ent).mins.v(2) + ev(ent).maxs.v(2)) * 0.5
        cont = SV_PointContents(point)
        If cont <= CONTENTS_WATER Then
            ev(ent).waterlevel = 2
            point.v(2) = ev(ent).origin.v(2) + ev(ent).view_ofs.v(2)
            cont = SV_PointContents(point)
            If cont <= CONTENTS_WATER Then ev(ent).waterlevel = 3
        End If
    End If

    SV_CheckWater = ev(ent).waterlevel > 1
End Function

Private Sub SV_WallFriction(ByVal ent As Long, trace As trace_t)
    Dim forward As vec3_t, right_ As vec3_t, up As vec3_t, d As Single, i As Single, into As vec3_t, side As vec3_t

    AngleVectors ev(ent).v_angle, forward, right_, up
    d = DotProduct(trace.plane.normal, forward)

    d = d + 0.5
    If d >= 0 Then Exit Sub

    ' cut the tangential velocity
    i = DotProduct(trace.plane.normal, ev(ent).velocity)
    VectorScale trace.plane.normal, i, into
    VectorSubtract ev(ent).velocity, into, side

    ev(ent).velocity.v(0) = side.v(0) * (1 + d)
    ev(ent).velocity.v(1) = side.v(1) * (1 + d)
End Sub

' Player has come to a dead stop, possibly due to the problem with limited
' float precision at some angle joins in the BSP hull.
'
' Try fixing by pushing one pixel in each direction.
'
' This is a hack, but in the interest of good gameplay...
Private Function SV_TryUnstick(ByVal ent As Long, oldvel As vec3_t) As Long
    Dim i As Long, oldorg As vec3_t, dir_ As vec3_t, clip As Long, steptrace As trace_t

    oldorg = ev(ent).origin
    dir_ = vec3_origin

    For i = 0 To 7
        ' try pushing a little in an axial direction
        Select Case i
        Case 0: dir_.v(0) = 2: dir_.v(1) = 0
        Case 1: dir_.v(0) = 0: dir_.v(1) = 2
        Case 2: dir_.v(0) = -2: dir_.v(1) = 0
        Case 3: dir_.v(0) = 0: dir_.v(1) = -2
        Case 4: dir_.v(0) = 2: dir_.v(1) = 2
        Case 5: dir_.v(0) = -2: dir_.v(1) = 2
        Case 6: dir_.v(0) = 2: dir_.v(1) = -2
        Case 7: dir_.v(0) = -2: dir_.v(1) = -2
        End Select

        SV_PushEntity ent, dir_

        ' retry the original move
        ev(ent).velocity.v(0) = oldvel.v(0)
        ev(ent).velocity.v(1) = oldvel.v(1)
        ev(ent).velocity.v(2) = 0
        clip = SV_FlyMove(ent, 0.1, steptrace, True)

        If Abs(oldorg.v(1) - ev(ent).origin.v(1)) > 4 Or Abs(oldorg.v(0) - ev(ent).origin.v(0)) > 4 Then
            SV_TryUnstick = clip
            Exit Function
        End If

        ' go back to the original pos and try again
        ev(ent).origin = oldorg
    Next

    ev(ent).velocity = vec3_origin
    SV_TryUnstick = 7                    ' still not moving
End Function

' Only used by players
Private Sub SV_WalkMove(ByVal ent As Long)
    Dim upmove As vec3_t, downmove As vec3_t, oldorg As vec3_t, oldvel As vec3_t
    Dim nosteporg As vec3_t, nostepvel As vec3_t, clip As Long, oldonground As Long
    Dim steptrace As trace_t, downtrace As trace_t

    '
    ' do a regular slide move unless it looks like you ran into a step
    '
    oldonground = ToInt(ev(ent).flags) And FL_ONGROUND
    ev(ent).flags = ToInt(ev(ent).flags) And Not FL_ONGROUND

    oldorg = ev(ent).origin
    oldvel = ev(ent).velocity

    clip = SV_FlyMove(ent, host_frametime, steptrace, True)

    If (clip And 2) = 0 Then Exit Sub    ' move didn't block on a step

    If oldonground = 0 And ev(ent).waterlevel = 0 Then Exit Sub   ' don't stair up while jumping

    If ev(ent).movetype <> MOVETYPE_WALK Then Exit Sub   ' gibbed by a trigger

    If cvars(cv_sv_nostep).value <> 0 Then Exit Sub

    If ToInt(ev(sv_player).flags) And FL_WATERJUMP Then Exit Sub

    nosteporg = ev(ent).origin
    nostepvel = ev(ent).velocity

    '
    ' try moving up and forward to go up a step
    '
    ev(ent).origin = oldorg              ' back to start pos

    upmove = vec3_origin
    downmove = vec3_origin
    upmove.v(2) = STEPSIZE
    downmove.v(2) = -STEPSIZE + oldvel.v(2) * host_frametime

    ' move up
    SV_PushEntity ent, upmove            ' FIXME: don't link?

    ' move forward
    ev(ent).velocity.v(0) = oldvel.v(0)
    ev(ent).velocity.v(1) = oldvel.v(1)
    ev(ent).velocity.v(2) = 0
    clip = SV_FlyMove(ent, host_frametime, steptrace, True)

    ' check for stuckness, possibly due to the limited precision of floats
    ' in the clipping hulls
    If clip Then
        If Abs(oldorg.v(1) - ev(ent).origin.v(1)) < 0.03125 And Abs(oldorg.v(0) - ev(ent).origin.v(0)) < 0.03125 Then
            ' stepping up didn't make any progress
            clip = SV_TryUnstick(ent, oldvel)
        End If
    End If

    ' extra friction based on view angle
    If clip And 2 Then SV_WallFriction ent, steptrace

    ' move down
    downtrace = SV_PushEntity(ent, downmove)   ' FIXME: don't link?

    If downtrace.plane.normal.v(2) > 0.7 Then
        If ev(ent).solid = SOLID_BSP Then
            ev(ent).flags = ToInt(ev(ent).flags) Or FL_ONGROUND
            ev(ent).groundentity = EDICT_TO_PROG(downtrace.ent)
        End If
    Else
        ' if the push down didn't end up on good ground, use the move without
        ' the step up.  This happens near wall / slope combinations, and can
        ' cause the player to hop up higher on a slope too steep to climb
        ev(ent).origin = nosteporg
        ev(ent).velocity = nostepvel
    End If
End Sub

' Player character actions
Private Sub SV_Physics_Client(ByVal ent As Long, ByVal num As Long)
    Dim dummy As trace_t

    If Not svs.clients(num - 1).active Then Exit Sub   ' unconnected slot

    '
    ' call standard client pre-think
    '
    pr_global_struct(0).time = sv.time
    pr_global_struct(0).self = EDICT_TO_PROG(ent)
    PR_ExecuteProgram pr_global_struct(0).PlayerPreThink

    '
    ' do a move
    '
    SV_CheckVelocity ent

    '
    ' decide which move function to call
    '
    Select Case ToInt(ev(ent).movetype)
    Case MOVETYPE_NONE
        If Not SV_RunThink(ent) Then Exit Sub

    Case MOVETYPE_WALK
        If Not SV_RunThink(ent) Then Exit Sub
        If Not SV_CheckWater(ent) And (ToInt(ev(ent).flags) And FL_WATERJUMP) = 0 Then SV_AddGravity ent
        SV_CheckStuck ent
        SV_WalkMove ent

    Case MOVETYPE_TOSS, MOVETYPE_BOUNCE
        SV_Physics_Toss ent

    Case MOVETYPE_FLY
        If Not SV_RunThink(ent) Then Exit Sub
        SV_FlyMove ent, host_frametime, dummy, False

    Case MOVETYPE_NOCLIP
        If Not SV_RunThink(ent) Then Exit Sub
        VectorMA ev(ent).origin, host_frametime, ev(ent).velocity, ev(ent).origin

    Case Else
        Sys_Error "SV_Physics_client: bad movetype %i", ToInt(ev(ent).movetype)
    End Select

    '
    ' call standard player post-think
    '
    SV_LinkEdict ent, True

    pr_global_struct(0).time = sv.time
    pr_global_struct(0).self = EDICT_TO_PROG(ent)
    PR_ExecuteProgram pr_global_struct(0).PlayerPostThink
End Sub

'============================================================================

' Non moving objects can only think
Private Sub SV_Physics_None(ByVal ent As Long)
    ' regular thinking
    SV_RunThink ent
End Sub

' A moving object that doesn't obey physics
Private Sub SV_Physics_Noclip(ByVal ent As Long)
    ' regular thinking
    If Not SV_RunThink(ent) Then Exit Sub

    VectorMA ev(ent).angles, host_frametime, ev(ent).avelocity, ev(ent).angles
    VectorMA ev(ent).origin, host_frametime, ev(ent).velocity, ev(ent).origin

    SV_LinkEdict ent, False
End Sub

'==============================================================================
' TOSS / BOUNCE

Private Sub SV_CheckWaterTransition(ByVal ent As Long)
    Dim cont As Long

    cont = SV_PointContents(ev(ent).origin)

    If ev(ent).watertype = 0 Then
        ' just spawned here
        ev(ent).watertype = cont
        ev(ent).waterlevel = 1
        Exit Sub
    End If

    If cont <= CONTENTS_WATER Then
        If ev(ent).watertype = CONTENTS_EMPTY Then
            ' just crossed into water
            SV_StartSound ent, 0, "misc/h2ohit1.wav", 255, 1
        End If
        ev(ent).watertype = cont
        ev(ent).waterlevel = 1
    Else
        If ev(ent).watertype <> CONTENTS_EMPTY Then
            ' just crossed into water
            SV_StartSound ent, 0, "misc/h2ohit1.wav", 255, 1
        End If
        ev(ent).watertype = CONTENTS_EMPTY
        ev(ent).waterlevel = cont
    End If
End Sub

' Toss, bounce, and fly movement.  When onground, do nothing.
Private Sub SV_Physics_Toss(ByVal ent As Long)
    Dim trace As trace_t, move As vec3_t, backoff As Single

    ' regular thinking
    If Not SV_RunThink(ent) Then Exit Sub

    ' if onground, return without moving
    If ToInt(ev(ent).flags) And FL_ONGROUND Then Exit Sub

    SV_CheckVelocity ent

    ' add gravity
    If ev(ent).movetype <> MOVETYPE_FLY And ev(ent).movetype <> MOVETYPE_FLYMISSILE Then SV_AddGravity ent

    ' move angles
    VectorMA ev(ent).angles, host_frametime, ev(ent).avelocity, ev(ent).angles

    ' move origin
    VectorScale ev(ent).velocity, host_frametime, move
    trace = SV_PushEntity(ent, move)
    If trace.fraction = 1 Then Exit Sub
    If edicts(ent).free Then Exit Sub

    If ev(ent).movetype = MOVETYPE_BOUNCE Then
        backoff = 1.5
    Else
        backoff = 1
    End If

    ClipVelocity ev(ent).velocity, trace.plane.normal, ev(ent).velocity, backoff

    ' stop if on ground
    If trace.plane.normal.v(2) > 0.7 Then
        If ev(ent).velocity.v(2) < 60 Or ev(ent).movetype <> MOVETYPE_BOUNCE Then
            ev(ent).flags = ToInt(ev(ent).flags) Or FL_ONGROUND
            ev(ent).groundentity = EDICT_TO_PROG(trace.ent)
            ev(ent).velocity = vec3_origin
            ev(ent).avelocity = vec3_origin
        End If
    End If

    ' check for in water
    SV_CheckWaterTransition ent
End Sub

'===============================================================================
' STEPPING MOVEMENT

' Monsters freefall when they don't have a ground entity, otherwise
' all movement is done with discrete steps.
'
' This is also used for objects that have become still on the ground, but
' will fall if the floor is pulled out from under them.
Private Sub SV_Physics_Step(ByVal ent As Long)
    Dim hitsound As Boolean, dummy As trace_t

    ' freefall if not onground
    If (ToInt(ev(ent).flags) And (FL_ONGROUND Or FL_FLY Or FL_SWIM)) = 0 Then
        If ev(ent).velocity.v(2) < cvars(cv_sv_gravity).value * -0.1 Then
            hitsound = True
        Else
            hitsound = False
        End If

        SV_AddGravity ent
        SV_CheckVelocity ent
        SV_FlyMove ent, host_frametime, dummy, False
        SV_LinkEdict ent, True

        If ToInt(ev(ent).flags) And FL_ONGROUND Then
            ' just hit ground
            If hitsound Then SV_StartSound ent, 0, "demon/dland2.wav", 255, 1
        End If
    End If

    ' regular thinking
    SV_RunThink ent

    SV_CheckWaterTransition ent
End Sub

'============================================================================

Public Sub SV_Physics()
    Dim i As Long, ent As Long

    ' let the progs know that a new frame has started
    pr_global_struct(0).self = EDICT_TO_PROG(0)
    pr_global_struct(0).other = EDICT_TO_PROG(0)
    pr_global_struct(0).time = sv.time
    PR_ExecuteProgram pr_global_struct(0).StartFrame

    '
    ' treat each object in turn
    '
    For i = 0 To sv.num_edicts - 1
        ent = i
        If edicts(ent).free Then GoTo nexti

        If pr_global_struct(0).force_retouch <> 0 Then SV_LinkEdict ent, True   ' force retouch even for stationary

        If i > 0 And i <= svs.maxclients Then
            SV_Physics_Client ent, i
        ElseIf ev(ent).movetype = MOVETYPE_PUSH Then
            SV_Physics_Pusher ent
        ElseIf ev(ent).movetype = MOVETYPE_NONE Then
            SV_Physics_None ent
        ElseIf ev(ent).movetype = MOVETYPE_NOCLIP Then
            SV_Physics_Noclip ent
        ElseIf ev(ent).movetype = MOVETYPE_STEP Then
            SV_Physics_Step ent
        ElseIf ev(ent).movetype = MOVETYPE_TOSS Or ev(ent).movetype = MOVETYPE_BOUNCE _
            Or ev(ent).movetype = MOVETYPE_FLY Or ev(ent).movetype = MOVETYPE_FLYMISSILE Then
            SV_Physics_Toss ent
        Else
            Sys_Error "SV_Physics: bad movetype %i", ToInt(ev(ent).movetype)
        End If
nexti:
    Next

    If pr_global_struct(0).force_retouch <> 0 Then pr_global_struct(0).force_retouch = pr_global_struct(0).force_retouch - 1

    sv.time = sv.time + host_frametime
End Sub
