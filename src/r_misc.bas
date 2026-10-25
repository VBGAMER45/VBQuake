Attribute VB_Name = "r_misc"
' r_misc.bas - frame setup, the view frustum, timing displays (r_misc.c)
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private oldbright As Single              ' C: static of R_CheckVariables

Private Const MAX_TIMINGS As Long = 100
Private timex As Long                    ' C: statics of R_TimeGraph
Private r_timings(0 To MAX_TIMINGS - 1) As Byte

Private Sub R_CheckVariables()
    If cvars(cv_r_fullbright).value <> oldbright Then
        oldbright = cvars(cv_r_fullbright).value
        D_FlushCaches                    ' so all lighting changes
    End If
End Sub

' Debugging use
Public Sub Show()
    Dim vr As vrect_t

    vr.x = 0
    vr.y = 0
    vr.width = vid.width
    vr.height = vid.height
    VID_Update vr
End Sub

' For program optimization
Public Sub R_TimeRefresh_f()
    Dim i As Long, start As Single, stop_ As Single, time_ As Single, startangle As Long, vr As vrect_t

    startangle = ToInt(r_refdef.viewangles.v(1))

    start = Sys_FloatTime()
    For i = 0 To 127
        r_refdef.viewangles.v(1) = i / 128# * 360#

        VID_LockBuffer

        R_RenderView

        VID_UnlockBuffer

        vr.x = r_refdef.vrect.x
        vr.y = r_refdef.vrect.y
        vr.width = r_refdef.vrect.width
        vr.height = r_refdef.vrect.height
        VID_Update vr
    Next
    stop_ = Sys_FloatTime()
    time_ = stop_ - start
    Con_Printf "%f seconds (%f fps)\n", time_, 128 / time_

    r_refdef.viewangles.v(1) = startangle
End Sub

' Only called by R_DisplayTime
Private Sub R_LineGraph(ByVal x As Long, ByVal y As Long, ByVal h As Long)
    Dim i As Long, dest As Long, s As Long

    ' FIXME: should be disabled on no-buffer adapters, or should be in the driver

    x = x + r_refdef.vrect.x
    y = y + r_refdef.vrect.y

    dest = vid_bufferofs + vid.rowbytes * y + x

    s = ToInt(cvars(cv_r_graphheight).value)

    If h > s Then h = s

    For i = 0 To h - 1
        vid_buffer(dest) = &HFF
        vid_buffer(dest - vid.rowbytes) = &H30
        dest = dest - vid.rowbytes * 2
    Next
    For i = i To s - 1
        vid_buffer(dest) = &H30
        vid_buffer(dest - vid.rowbytes) = &H30
        dest = dest - vid.rowbytes * 2
    Next
End Sub

' Performance monitoring tool
Public Sub R_TimeGraph()
    Dim a As Long, r_time2 As Single, x As Long

    r_time2 = Sys_FloatTime()

    a = ToInt((r_time2 - r_time1) / 0.01)
    ' a = fabs(mouse_y * 0.05);
    ' a = (int)((r_refdef.vieworg[2] + 1024)/1)%(int)r_graphheight.value;
    ' a = fabs(velocity[0])/20;
    ' a = ((int)fabs(origin[0])/8)%20;
    ' a = (cl.idealpitch + 30)/5;
    r_timings(timex) = a And &HFF&
    a = timex

    If r_refdef.vrect.width <= MAX_TIMINGS Then
        x = r_refdef.vrect.width - 1
    Else
        x = r_refdef.vrect.width - (r_refdef.vrect.width - MAX_TIMINGS) \ 2
    End If
    Do
        R_LineGraph x, r_refdef.vrect.height - 2, r_timings(a)
        If x = 0 Then Exit Do            ' screen too small to hold entire thing
        x = x - 1
        a = a - 1
        If a = -1 Then a = MAX_TIMINGS - 1
    Loop While a <> timex

    timex = (timex + 1) Mod MAX_TIMINGS
End Sub

Public Sub R_PrintTimes()
    Dim r_time2 As Single, ms As Single

    r_time2 = Sys_FloatTime()

    ms = 1000 * (r_time2 - r_time1)

    Con_Printf "%5.1f ms %3i/%3i/%3i poly %3i surf\n", ms, c_faceclip, r_polycount, r_drawnpolycount, c_surf
    c_surf = 0
End Sub

Public Sub R_PrintDSpeeds()
    Dim ms As Single, dp_time As Single, r_time2 As Single, rw_time As Single, db_time As Single
    Dim se_time As Single, de_time As Single, dv_time As Single

    r_time2 = Sys_FloatTime()

    dp_time = (dp_time2 - dp_time1) * 1000
    rw_time = (rw_time2 - rw_time1) * 1000
    db_time = (db_time2 - db_time1) * 1000
    se_time = (se_time2 - se_time1) * 1000
    de_time = (de_time2 - de_time1) * 1000
    dv_time = (dv_time2 - dv_time1) * 1000
    ms = (r_time2 - r_time1) * 1000

    Con_Printf "%3i %4.1fp %3iw %4.1fb %3is %4.1fe %4.1fv\n", ToInt(ms), dp_time, ToInt(rw_time), db_time, ToInt(se_time), de_time, dv_time
End Sub

Public Sub R_PrintAliasStats()
    Con_Printf "%3i polygon model drawn\n", r_amodels_drawn
End Sub

Public Sub WarpPalette()
    Dim i As Long, j As Long, newpalette(0 To 767) As Byte, basecolor(0 To 2) As Long

    basecolor(0) = 130
    basecolor(1) = 80
    basecolor(2) = 50

    ' pull the colors halfway to bright brown
    For i = 0 To 255
        For j = 0 To 2
            newpalette(i * 3 + j) = (host_basepal(i * 3 + j) + basecolor(j)) \ 2
        Next
    Next

    VID_ShiftPalette newpalette
End Sub

Public Sub R_TransformFrustum()
    Dim i As Long, v As vec3_t, v2 As vec3_t

    For i = 0 To 3
        v.v(0) = screenedge(i).normal.v(2)
        v.v(1) = -screenedge(i).normal.v(0)
        v.v(2) = screenedge(i).normal.v(1)

        v2.v(0) = v.v(1) * vright.v(0) + v.v(2) * vup.v(0) + v.v(0) * vpn.v(0)
        v2.v(1) = v.v(1) * vright.v(1) + v.v(2) * vup.v(1) + v.v(0) * vpn.v(1)
        v2.v(2) = v.v(1) * vright.v(2) + v.v(2) * vup.v(2) + v.v(0) * vpn.v(2)

        view_clipplanes(i).normal = v2

        view_clipplanes(i).dist = DotProduct(modelorg, v2)
    Next
End Sub

Public Sub TransformVector(in_ As vec3_t, out As vec3_t)
    out.v(0) = DotProduct(in_, vright)
    out.v(1) = DotProduct(in_, vup)
    out.v(2) = DotProduct(in_, vpn)
End Sub

Public Sub R_TransformPlane(p As mplane_t, normal As vec3_t, dist As Single)
    Dim d As Single

    d = DotProduct(r_origin, p.normal)
    dist = p.dist - d
    ' TODO: when we have rotating entities, this will need to use the view matrix
    TransformVector p.normal, normal
End Sub

Private Sub R_SetUpFrustumIndexes()
    Dim i As Long, j As Long, pindex As Long

    pindex = 0

    For i = 0 To 3
        For j = 0 To 2
            If view_clipplanes(i).normal.v(j) < 0 Then
                r_frustum_indexes(pindex + j) = j
                r_frustum_indexes(pindex + j + 3) = j + 3
            Else
                r_frustum_indexes(pindex + j) = j + 3
                r_frustum_indexes(pindex + j + 3) = j
            End If
        Next

        ' FIXME: do just once at start
        pfrustum_indexes(i) = pindex
        pindex = pindex + 6
    Next
End Sub

Public Sub R_SetupFrame()
    Dim edgecount As Long, vrect As vrect_t, w As Single, h As Single

    ' don't allow cheats in multiplayer
    If cl.maxclients > 1 Then
        Cvar_Set "r_draworder", "0"
        Cvar_Set "r_fullbright", "0"
        Cvar_Set "r_ambient", "0"
        Cvar_Set "r_drawflat", "0"
    End If

    If cvars(cv_r_numsurfs).value <> 0 Then
        If surface_p > r_maxsurfsseen Then r_maxsurfsseen = surface_p

        Con_Printf "Used %d of %d surfs; %d max\n", surface_p, surf_max, r_maxsurfsseen
    End If

    If cvars(cv_r_numedges).value <> 0 Then
        edgecount = edge_p - EDGES_BASE

        If edgecount > r_maxedgesseen Then r_maxedgesseen = edgecount

        Con_Printf "Used %d of %d edges; %d max\n", edgecount, r_numallocatededges, r_maxedgesseen
    End If

    r_refdef.ambientlight = ToInt(cvars(cv_r_ambient).value)

    If r_refdef.ambientlight < 0 Then r_refdef.ambientlight = 0

    If Not sv.active Then cvars(cv_r_draworder).value = 0   ' don't let cheaters look behind walls

    R_CheckVariables

    R_AnimateLight

    r_framecount = r_framecount + 1

    numbtofpolys = 0

    ' build the transformation matrix for the given view angles
    modelorg = r_refdef.vieworg
    r_origin = r_refdef.vieworg

    AngleVectors r_refdef.viewangles, vpn, vright, vup

    ' current viewleaf
    r_oldviewleaf = r_viewleaf
    r_viewleaf = Mod_PointInLeaf(r_origin, cl.worldmodel)

    r_dowarpold = r_dowarp
    r_dowarp = cvars(cv_r_waterwarp).value <> 0 And (mod_nodes(r_viewleaf).contents <= CONTENTS_WATER)

    If (r_dowarp <> r_dowarpold) Or r_viewchanged_ Or cvars(cv_lcd_x).value <> 0 Then
        If r_dowarp Then
            If (vid.width <= vid.maxwarpwidth) And (vid.height <= vid.maxwarpheight) Then
                vrect.x = 0
                vrect.y = 0
                vrect.width = vid.width
                vrect.height = vid.height

                R_ViewChanged vrect, sb_lines, vid.aspect
            Else
                w = vid.width
                h = vid.height

                If w > vid.maxwarpwidth Then
                    h = h * CSng(vid.maxwarpwidth) / w
                    w = vid.maxwarpwidth
                End If

                If h > vid.maxwarpheight Then
                    h = vid.maxwarpheight
                    w = w * CSng(vid.maxwarpheight) / h
                End If

                vrect.x = 0
                vrect.y = 0
                vrect.width = ToInt(w)
                vrect.height = ToInt(h)

                R_ViewChanged vrect, ToInt(CSng(sb_lines) * (h / CSng(vid.height))), _
                    vid.aspect * (h / w) * (CSng(vid.width) / CSng(vid.height))
            End If
        Else
            vrect.x = 0
            vrect.y = 0
            vrect.width = vid.width
            vrect.height = vid.height

            R_ViewChanged vrect, sb_lines, vid.aspect
        End If

        r_viewchanged_ = False
    End If

    ' start off with just the four screen edge clip planes
    R_TransformFrustum

    ' save base values
    base_vpn = vpn
    base_vright = vright
    base_vup = vup
    base_modelorg = modelorg

    R_SetSkyFrame

    R_SetUpFrustumIndexes

    r_cache_thrash = False

    ' clear frame counts
    c_faceclip = 0
    d_spanpixcount = 0
    r_polycount = 0
    r_drawnpolycount = 0
    r_wholepolycount = 0
    r_amodels_drawn = 0
    r_outofsurfaces = 0
    r_outofedges = 0

    D_SetupFrame
End Sub
