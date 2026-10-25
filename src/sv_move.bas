Attribute VB_Name = "sv_move_c"
' sv_move.bas - monster movement (sv_move.c)
' (module sv_move_c: VB names are not case sensitive, and world has SV_Move)
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const STEPSIZE As Long = 18
Private Const DI_NODIR As Single = -1

Public c_yes As Long, c_no As Long

' Returns false if any part of the bottom of the entity is off an edge that
' is not a staircase.
Public Function SV_CheckBottom(ByVal ent As Long) As Boolean
    Dim mins As vec3_t, maxs As vec3_t, start As vec3_t, stop_ As vec3_t
    Dim trace As trace_t, x As Long, y As Long, mid As Single, bottom As Single

    VectorAdd ev(ent).origin, ev(ent).mins, mins
    VectorAdd ev(ent).origin, ev(ent).maxs, maxs

    ' if all of the points under the corners are solid world, don't bother
    ' with the tougher checks
    ' the corners must be within 16 of the midpoint
    start.v(2) = mins.v(2) - 1
    For x = 0 To 1
        For y = 0 To 1
            If x Then start.v(0) = maxs.v(0) Else start.v(0) = mins.v(0)
            If y Then start.v(1) = maxs.v(1) Else start.v(1) = mins.v(1)
            If SV_PointContents(start) <> CONTENTS_SOLID Then GoTo realcheck
        Next
    Next

    c_yes = c_yes + 1
    SV_CheckBottom = True                ' we got out easy
    Exit Function

realcheck:
    c_no = c_no + 1
    '
    ' check it for real...
    '
    start.v(2) = mins.v(2)

    ' the midpoint must be within 16 of the bottom
    start.v(0) = (mins.v(0) + maxs.v(0)) * 0.5
    stop_.v(0) = start.v(0)
    start.v(1) = (mins.v(1) + maxs.v(1)) * 0.5
    stop_.v(1) = start.v(1)
    stop_.v(2) = start.v(2) - 2 * STEPSIZE
    trace = SV_Move(start, vec3_origin, vec3_origin, stop_, MOVE_NOMONSTERS, ent)

    If trace.fraction = 1 Then Exit Function
    mid = trace.endpos.v(2)
    bottom = mid

    ' the corners must be within 16 of the midpoint
    For x = 0 To 1
        For y = 0 To 1
            If x Then start.v(0) = maxs.v(0) Else start.v(0) = mins.v(0)
            stop_.v(0) = start.v(0)
            If y Then start.v(1) = maxs.v(1) Else start.v(1) = mins.v(1)
            stop_.v(1) = start.v(1)

            trace = SV_Move(start, vec3_origin, vec3_origin, stop_, MOVE_NOMONSTERS, ent)

            If trace.fraction <> 1 And trace.endpos.v(2) > bottom Then bottom = trace.endpos.v(2)
            If trace.fraction = 1 Or mid - trace.endpos.v(2) > STEPSIZE Then Exit Function
        Next
    Next

    c_yes = c_yes + 1
    SV_CheckBottom = True
End Function

' Called by monster program code.
' The move will be adjusted for slopes and stairs, but if the move isn't
' possible, no move is done, false is returned, and
' pr_global_struct->trace_normal is set to the normal of the blocking wall
Public Function SV_movestep(ByVal ent As Long, move As vec3_t, ByVal relink As Boolean) As Boolean
    Dim dz As Single, oldorg As vec3_t, neworg As vec3_t, end_ As vec3_t
    Dim trace As trace_t, i As Long, enemy As Long

    ' try the move
    oldorg = ev(ent).origin
    VectorAdd ev(ent).origin, move, neworg

    ' flying monsters don't step up
    If ToInt(ev(ent).flags) And (FL_SWIM Or FL_FLY) Then
        ' try one move with vertical motion, then one without
        For i = 0 To 1
            VectorAdd ev(ent).origin, move, neworg
            enemy = PROG_TO_EDICT(ev(ent).enemy)
            If i = 0 And enemy <> 0 Then
                dz = ev(ent).origin.v(2) - ev(PROG_TO_EDICT(ev(ent).enemy)).origin.v(2)
                If dz > 40 Then neworg.v(2) = neworg.v(2) - 8
                If dz < 30 Then neworg.v(2) = neworg.v(2) + 8
            End If
            trace = SV_Move(ev(ent).origin, ev(ent).mins, ev(ent).maxs, neworg, MOVE_NORMAL, ent)

            If trace.fraction = 1 Then
                If (ToInt(ev(ent).flags) And FL_SWIM) <> 0 Then
                    If SV_PointContents(trace.endpos) = CONTENTS_EMPTY Then Exit Function   ' swim monster left water
                End If

                ev(ent).origin = trace.endpos
                If relink Then SV_LinkEdict ent, True
                SV_movestep = True
                Exit Function
            End If

            If enemy = 0 Then Exit For
        Next

        Exit Function
    End If

    ' push down from a step height above the wished position
    neworg.v(2) = neworg.v(2) + STEPSIZE
    end_ = neworg
    end_.v(2) = end_.v(2) - STEPSIZE * 2

    trace = SV_Move(neworg, ev(ent).mins, ev(ent).maxs, end_, MOVE_NORMAL, ent)

    If trace.allsolid Then Exit Function

    If trace.startsolid Then
        neworg.v(2) = neworg.v(2) - STEPSIZE
        trace = SV_Move(neworg, ev(ent).mins, ev(ent).maxs, end_, MOVE_NORMAL, ent)
        If trace.allsolid Or trace.startsolid Then Exit Function
    End If
    If trace.fraction = 1 Then
        ' if monster had the ground pulled out, go ahead and fall
        If ToInt(ev(ent).flags) And FL_PARTIALGROUND Then
            VectorAdd ev(ent).origin, move, ev(ent).origin
            If relink Then SV_LinkEdict ent, True
            ev(ent).flags = ToInt(ev(ent).flags) And Not FL_ONGROUND
            SV_movestep = True
            Exit Function
        End If

        Exit Function                    ' walked off an edge
    End If

    ' check point traces down for dangling corners
    ev(ent).origin = trace.endpos

    If Not SV_CheckBottom(ent) Then
        If ToInt(ev(ent).flags) And FL_PARTIALGROUND Then
            ' entity had floor mostly pulled out from underneath it
            ' and is trying to correct
            If relink Then SV_LinkEdict ent, True
            SV_movestep = True
            Exit Function
        End If
        ev(ent).origin = oldorg
        Exit Function
    End If

    If ToInt(ev(ent).flags) And FL_PARTIALGROUND Then
        ev(ent).flags = ToInt(ev(ent).flags) And Not FL_PARTIALGROUND
    End If
    ev(ent).groundentity = EDICT_TO_PROG(trace.ent)

    ' the move is ok
    If relink Then SV_LinkEdict ent, True
    SV_movestep = True
End Function

'============================================================================

' Turns to the movement direction, and walks the current distance if
' facing it.
Private Function SV_StepDirection(ByVal ent As Long, ByVal yaw_ As Single, ByVal dist As Single) As Boolean
    Dim move As vec3_t, oldorigin As vec3_t, delta As Single

    ' (QuakeVB: yaw_, since VB would take YAW below for the parameter)
    ev(ent).ideal_yaw = yaw_
    PF_changeyaw

    yaw_ = yaw_ * M_PI * 2 / 360
    move.v(0) = Cos(yaw_) * dist
    move.v(1) = Sin(yaw_) * dist
    move.v(2) = 0

    oldorigin = ev(ent).origin
    If SV_movestep(ent, move, False) Then
        delta = ev(ent).angles.v(YAW) - ev(ent).ideal_yaw
        If delta > 45 And delta < 315 Then
            ' not turned far enough, so don't take the step
            ev(ent).origin = oldorigin
        End If
        SV_LinkEdict ent, True
        SV_StepDirection = True
        Exit Function
    End If
    SV_LinkEdict ent, True
End Function

Private Sub SV_FixCheckBottom(ByVal ent As Long)
    ev(ent).flags = ToInt(ev(ent).flags) Or FL_PARTIALGROUND
End Sub

Private Sub SV_NewChaseDir(ByVal actor As Long, ByVal enemy As Long, ByVal dist As Single)
    Dim deltax As Single, deltay As Single, d(0 To 2) As Single
    Dim tdir As Single, olddir As Single, turnaround As Single

    olddir = anglemod(ToInt(ev(actor).ideal_yaw / 45) * 45)
    turnaround = anglemod(olddir - 180)

    deltax = ev(enemy).origin.v(0) - ev(actor).origin.v(0)
    deltay = ev(enemy).origin.v(1) - ev(actor).origin.v(1)
    If deltax > 10 Then
        d(1) = 0
    ElseIf deltax < -10 Then
        d(1) = 180
    Else
        d(1) = DI_NODIR
    End If
    If deltay < -10 Then
        d(2) = 270
    ElseIf deltay > 10 Then
        d(2) = 90
    Else
        d(2) = DI_NODIR
    End If

    ' try direct route
    If d(1) <> DI_NODIR And d(2) <> DI_NODIR Then
        If d(1) = 0 Then
            If d(2) = 90 Then tdir = 45 Else tdir = 315
        Else
            If d(2) = 90 Then tdir = 135 Else tdir = 215
        End If

        If tdir <> turnaround Then
            If SV_StepDirection(actor, tdir, dist) Then Exit Sub
        End If
    End If

    ' try other directions
    If ((rand() And 3) And 1) <> 0 Or Abs(ToInt(deltay)) > Abs(ToInt(deltax)) Then
        tdir = d(1)
        d(1) = d(2)
        d(2) = tdir
    End If

    If d(1) <> DI_NODIR And d(1) <> turnaround Then
        If SV_StepDirection(actor, d(1), dist) Then Exit Sub
    End If

    If d(2) <> DI_NODIR And d(2) <> turnaround Then
        If SV_StepDirection(actor, d(2), dist) Then Exit Sub
    End If

    ' there is no direct path to the player, so pick another direction

    If olddir <> DI_NODIR Then
        If SV_StepDirection(actor, olddir, dist) Then Exit Sub
    End If

    If rand() And 1 Then
        ' randomly determine direction of search
        For tdir = 0 To 315 Step 45
            If tdir <> turnaround Then
                If SV_StepDirection(actor, tdir, dist) Then Exit Sub
            End If
        Next
    Else
        For tdir = 315 To 0 Step -45
            If tdir <> turnaround Then
                If SV_StepDirection(actor, tdir, dist) Then Exit Sub
            End If
        Next
    End If

    If turnaround <> DI_NODIR Then
        If SV_StepDirection(actor, turnaround, dist) Then Exit Sub
    End If

    ev(actor).ideal_yaw = olddir         ' can't move

    ' if a bridge was pulled out from underneath a monster, it may not have
    ' a valid standing position at all

    If Not SV_CheckBottom(actor) Then SV_FixCheckBottom actor
End Sub

Private Function SV_CloseEnough(ByVal ent As Long, ByVal goal As Long, ByVal dist As Single) As Boolean
    Dim i As Long

    For i = 0 To 2
        If ev(goal).absmin.v(i) > ev(ent).absmax.v(i) + dist Then Exit Function
        If ev(goal).absmax.v(i) < ev(ent).absmin.v(i) - dist Then Exit Function
    Next
    SV_CloseEnough = True
End Function

Public Sub SV_MoveToGoal()
    Dim ent As Long, goal As Long, dist As Single

    ent = PROG_TO_EDICT(pr_global_struct(0).self)
    goal = PROG_TO_EDICT(ev(ent).goalentity)
    dist = pr_gf(OFS_PARM0)

    If (ToInt(ev(ent).flags) And (FL_ONGROUND Or FL_FLY Or FL_SWIM)) = 0 Then
        pr_gf(OFS_RETURN) = 0
        Exit Sub
    End If

    ' if the next step hits the enemy, return immediately
    If PROG_TO_EDICT(ev(ent).enemy) <> 0 Then
        If SV_CloseEnough(ent, goal, dist) Then Exit Sub
    End If

    ' bump around...
    If (rand() And 3) = 1 Then
        SV_NewChaseDir ent, goal, dist
    ElseIf Not SV_StepDirection(ent, ev(ent).ideal_yaw, dist) Then
        SV_NewChaseDir ent, goal, dist
    End If
End Sub
