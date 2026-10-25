Attribute VB_Name = "r_light"
' r_light.bas - light styles, dynamic lights, light sampling (r_light.c)
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public r_dlightframecount As Long

Public Sub R_AnimateLight()
    Dim i As Long, j As Long, k As Long

    '
    ' light animations
    ' 'm' is normal light, 'a' is no light, 'z' is double bright
    i = ToInt(cl.time * 10)
    For j = 0 To MAX_LIGHTSTYLES - 1
        If cl_lightstyle(j).length = 0 Then
            d_lightstylevalue(j) = 256
        Else
            k = i Mod cl_lightstyle(j).length
            k = AscW(Mid$(cl_lightstyle(j).map, k + 1, 1)) - 97   ' - 'a'
            k = k * 22
            d_lightstylevalue(j) = k
        End If
    Next
End Sub

'=============================================================================
' DYNAMIC LIGHTS

' light: a cl_dlights() index
Public Sub R_MarkLights(ByVal light As Long, ByVal bit As Long, ByVal node As Long)
    Dim splitplane As Long, dist As Single, surf As Long, i As Long

    If mod_nodes(node).contents < 0 Then Exit Sub

    splitplane = mod_nodes(node).plane
    dist = DotProduct(cl_dlights(light).origin, mod_planes(splitplane).normal) - mod_planes(splitplane).dist

    If dist > cl_dlights(light).radius Then
        R_MarkLights light, bit, mod_nodes(node).children(0)
        Exit Sub
    End If
    If dist < -cl_dlights(light).radius Then
        R_MarkLights light, bit, mod_nodes(node).children(1)
        Exit Sub
    End If

    ' mark the polygons
    surf = mod_known(cl.worldmodel).surfaces + mod_nodes(node).firstsurface
    For i = 0 To mod_nodes(node).numsurfaces - 1
        If mod_surfaces(surf).dlightframe <> r_dlightframecount Then
            mod_surfaces(surf).dlightbits = 0
            mod_surfaces(surf).dlightframe = r_dlightframecount
        End If
        mod_surfaces(surf).dlightbits = mod_surfaces(surf).dlightbits Or bit
        surf = surf + 1
    Next

    R_MarkLights light, bit, mod_nodes(node).children(0)
    R_MarkLights light, bit, mod_nodes(node).children(1)
End Sub

Public Sub R_PushDlights()
    Dim i As Long

    r_dlightframecount = r_framecount + 1    ' because the count hasn't
                                             '  advanced yet for this frame

    For i = 0 To MAX_DLIGHTS - 1
        If cl_dlights(i).die < cl.time Or cl_dlights(i).radius = 0 Then GoTo nexti
        R_MarkLights i, BIT_(i), mod_known(cl.worldmodel).nodes
nexti:
    Next
End Sub

'=============================================================================
' LIGHT SAMPLING

Private Function RecursiveLightPoint(ByVal node As Long, start As vec3_t, end_ As vec3_t) As Long
    Dim r As Long, front As Single, back As Single, frac As Single, side As Long
    Dim plane As Long, mid As vec3_t, surf As Long, s As Long, t As Long, ds As Long, dt As Long
    Dim i As Long, tex As Long, lightmap As Long, scale_ As Long, maps As Long, hit As Boolean

    If mod_nodes(node).contents < 0 Then
        RecursiveLightPoint = -1         ' didn't hit anything
        Exit Function
    End If

    ' calculate mid point

    ' FIXME: optimize for axial
    plane = mod_nodes(node).plane
    front = DotProduct(start, mod_planes(plane).normal) - mod_planes(plane).dist
    back = DotProduct(end_, mod_planes(plane).normal) - mod_planes(plane).dist
    side = -(front < 0)

    If -(back < 0) = side Then
        RecursiveLightPoint = RecursiveLightPoint(mod_nodes(node).children(side), start, end_)
        Exit Function
    End If

    frac = front / (front - back)
    mid.v(0) = start.v(0) + (end_.v(0) - start.v(0)) * frac
    mid.v(1) = start.v(1) + (end_.v(1) - start.v(1)) * frac
    mid.v(2) = start.v(2) + (end_.v(2) - start.v(2)) * frac

    ' go down front side
    r = RecursiveLightPoint(mod_nodes(node).children(side), start, mid)
    If r >= 0 Then
        RecursiveLightPoint = r          ' hit something
        Exit Function
    End If

    If -(back < 0) = side Then
        RecursiveLightPoint = -1         ' didn't hit anuthing
        Exit Function
    End If

    ' check for impact on this node

    surf = mod_known(cl.worldmodel).surfaces + mod_nodes(node).firstsurface
    For i = 0 To mod_nodes(node).numsurfaces - 1
        ' (QuakeVB) the continues jump to endsurf, inside the With block, and
        ' hit says whether the surface was hit (see tools/check_with.py)
        hit = False
        With mod_surfaces(surf)
            If .flags And SURF_DRAWTILED Then GoTo endsurf   ' no lightmaps

            tex = .texinfo

            s = ToInt(mid.v(0) * mod_texinfo(tex).vecs(0, 0) + mid.v(1) * mod_texinfo(tex).vecs(0, 1) _
                + mid.v(2) * mod_texinfo(tex).vecs(0, 2) + mod_texinfo(tex).vecs(0, 3))
            t = ToInt(mid.v(0) * mod_texinfo(tex).vecs(1, 0) + mid.v(1) * mod_texinfo(tex).vecs(1, 1) _
                + mid.v(2) * mod_texinfo(tex).vecs(1, 2) + mod_texinfo(tex).vecs(1, 3))

            If s < .texturemins(0) Or t < .texturemins(1) Then GoTo endsurf

            ds = s - .texturemins(0)
            dt = t - .texturemins(1)

            If ds > .extents(0) Or dt > .extents(1) Then GoTo endsurf

            If .samples = -1 Then
                RecursiveLightPoint = 0
                Exit Function
            End If

            ds = ds \ 16
            dt = dt \ 16

            lightmap = .samples
            r = 0
            lightmap = lightmap + dt * ((.extents(0) \ 16) + 1) + ds

            maps = 0
            Do While maps < MAXLIGHTMAPS
                If .styles(maps) = 255 Then Exit Do
                scale_ = d_lightstylevalue(.styles(maps))
                r = r + mod_lightdata(lightmap) * scale_
                lightmap = lightmap + ((.extents(0) \ 16) + 1) * ((.extents(1) \ 16) + 1)
                maps = maps + 1
            Loop

            r = r \ 256
            hit = True
endsurf:
        End With

        If hit Then
            RecursiveLightPoint = r
            Exit Function
        End If
        surf = surf + 1
    Next

    ' go down back side
    RecursiveLightPoint = RecursiveLightPoint(mod_nodes(node).children(1 - side), mid, end_)
End Function

Public Function R_LightPoint(p As vec3_t) As Long
    Dim end_ As vec3_t, r As Long

    If mod_known(cl.worldmodel).lightdata = -1 Then
        R_LightPoint = 255
        Exit Function
    End If

    end_.v(0) = p.v(0)
    end_.v(1) = p.v(1)
    end_.v(2) = p.v(2) - 2048

    r = RecursiveLightPoint(mod_known(cl.worldmodel).nodes, p, end_)

    If r = -1 Then r = 0

    If r < r_refdef.ambientlight Then r = r_refdef.ambientlight

    R_LightPoint = r
End Function
