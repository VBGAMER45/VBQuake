Attribute VB_Name = "chase"
' chase.bas - chase camera code (chase.c)
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public cv_chase_back As Long
Public cv_chase_up As Long
Public cv_chase_right As Long
Public cv_chase_active As Long

Public chase_pos As vec3_t
Public chase_angles As vec3_t

Public chase_dest As vec3_t
Public chase_dest_angles As vec3_t

Public Sub Chase_Init()
    cv_chase_back = Cvar_RegisterVariable("chase_back", "100")
    cv_chase_up = Cvar_RegisterVariable("chase_up", "16")
    cv_chase_right = Cvar_RegisterVariable("chase_right", "0")
    cv_chase_active = Cvar_RegisterVariable("chase_active", "0")
End Sub

Public Sub Chase_Reset()
    ' for respawning and teleporting
    ' start position 12 units behind head
End Sub

Private Sub TraceLine(start As vec3_t, end_ As vec3_t, impact As vec3_t)
    Dim trace As trace_t

    ' (memset 0: fraction 0, and endpos stays 0 0 0 if nothing is hit, as in C)
    SV_RecursiveHullCheck mod_known(cl.worldmodel).hulls(0), 0, 0, 1, start, end_, trace

    impact = trace.endpos
End Sub

Public Sub Chase_Update()
    Dim i As Long, dist As Single, forward As vec3_t, up As vec3_t, right_ As vec3_t
    Dim dest As vec3_t, stop_ As vec3_t

    ' if can't see player, reset
    AngleVectors cl.viewangles, forward, right_, up

    ' calc exact destination
    For i = 0 To 2
        chase_dest.v(i) = r_refdef.vieworg.v(i) - forward.v(i) * cvars(cv_chase_back).value - right_.v(i) * cvars(cv_chase_right).value
    Next
    chase_dest.v(2) = r_refdef.vieworg.v(2) + cvars(cv_chase_up).value

    ' find the spot the player is looking at
    VectorMA r_refdef.vieworg, 4096, forward, dest
    TraceLine r_refdef.vieworg, dest, stop_

    ' calculate pitch to look at the same spot from camera
    VectorSubtract stop_, r_refdef.vieworg, stop_
    dist = DotProduct(stop_, forward)
    If dist < 1 Then dist = 1
    r_refdef.viewangles.v(PITCH) = -Atn(stop_.v(2) / dist) / M_PI * 180

    ' move towards destination
    r_refdef.vieworg = chase_dest
End Sub
