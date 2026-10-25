Attribute VB_Name = "r_surf"
' r_surf.bas - surface-related refresh code: builds the lit surfaces of the
' surface cache (r_surf.c)
'
' C has four R_DrawSurfaceBlock8_mip0..3, the same loop for blocks of 16, 8,
' 4 and 2 texels; here that is one routine with the block size. Not ported:
' the 16 bit drawers (R_DrawSurfaceBlock16, R_GenTurbTile16; r_pixbytes is
' always 1) and the tile generators R_GenTile and R_GenTurbTile, which
' nothing calls.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public r_drawsurf As drawsurf_t

Private lightleft As Long, sourcesstep As Long, blocksize As Long, sourcetstep As Long
Private lightdelta As Long, lightdeltastep As Long
Private lightright As Long, lightleftstep As Long, lightrightstep As Long, blockdivshift As Long
Private blockdivmask As Long
Private prowdestbase As Long             ' sc_bytes() offset
Private pbasesource As Long              ' texpixels() offset
Public surfrowbytes As Long              ' used by ASM files
Private r_lightptr As Long               ' blocklights() index
Private r_stepback As Long
Private r_lightwidth As Long
Private r_numhblocks As Long, r_numvblocks As Long
Private r_source As Long, r_sourcemax As Long   ' texpixels() offsets

Private blocklights(0 To 18 * 18 - 1) As Long    ' C: unsigned

Private Sub R_AddDynamicLights()
    Dim surf As Long, lnum As Long, sd As Long, td As Long, dist As Single, rad As Single, minlight As Single
    Dim impact As vec3_t, local_ As vec3_t, s As Long, t As Long, i As Long, smax As Long, tmax As Long
    Dim tex As Long, pl As Long

    surf = r_drawsurf.surf
    smax = (mod_surfaces(surf).extents(0) \ 16) + 1
    tmax = (mod_surfaces(surf).extents(1) \ 16) + 1
    tex = mod_surfaces(surf).texinfo
    pl = mod_surfaces(surf).plane

    For lnum = 0 To MAX_DLIGHTS - 1
        If (mod_surfaces(surf).dlightbits And BIT_(lnum)) = 0 Then GoTo nextlight   ' not lit by this light

        rad = cl_dlights(lnum).radius
        dist = DotProduct(cl_dlights(lnum).origin, mod_planes(pl).normal) - mod_planes(pl).dist
        rad = rad - Abs(dist)
        minlight = cl_dlights(lnum).minlight
        If rad < minlight Then GoTo nextlight
        minlight = rad - minlight

        For i = 0 To 2
            impact.v(i) = cl_dlights(lnum).origin.v(i) - mod_planes(pl).normal.v(i) * dist
        Next

        local_.v(0) = impact.v(0) * mod_texinfo(tex).vecs(0, 0) + impact.v(1) * mod_texinfo(tex).vecs(0, 1) _
            + impact.v(2) * mod_texinfo(tex).vecs(0, 2) + mod_texinfo(tex).vecs(0, 3)
        local_.v(1) = impact.v(0) * mod_texinfo(tex).vecs(1, 0) + impact.v(1) * mod_texinfo(tex).vecs(1, 1) _
            + impact.v(2) * mod_texinfo(tex).vecs(1, 2) + mod_texinfo(tex).vecs(1, 3)

        local_.v(0) = local_.v(0) - mod_surfaces(surf).texturemins(0)
        local_.v(1) = local_.v(1) - mod_surfaces(surf).texturemins(1)

        For t = 0 To tmax - 1
            td = ToInt(local_.v(1) - t * 16)
            If td < 0 Then td = -td
            For s = 0 To smax - 1
                sd = ToInt(local_.v(0) - s * 16)
                If sd < 0 Then sd = -sd
                If sd > td Then
                    dist = sd + (td \ 2)
                Else
                    dist = td + (sd \ 2)
                End If
                If dist < minlight Then
                    blocklights(t * smax + s) = ToInt(blocklights(t * smax + s) + (rad - dist) * 256)
                End If
            Next
        Next
nextlight:
    Next
End Sub

' Combine and scale multiple lightmaps into the 8.8 format in blocklights
Private Sub R_BuildLightMap()
    Dim smax As Long, tmax As Long, t As Long, i As Long, size As Long, lightmap As Long
    Dim scale_ As Long, maps As Long, surf As Long

    surf = r_drawsurf.surf

    smax = (mod_surfaces(surf).extents(0) \ 16) + 1
    tmax = (mod_surfaces(surf).extents(1) \ 16) + 1
    size = smax * tmax
    lightmap = mod_surfaces(surf).samples

    If cvars(cv_r_fullbright).value <> 0 Or mod_known(cl.worldmodel).lightdata = -1 Then
        For i = 0 To size - 1
            blocklights(i) = 0
        Next
        Exit Sub
    End If

    ' clear to ambient
    For i = 0 To size - 1
        blocklights(i) = r_refdef.ambientlight * 256
    Next

    ' add all the lightmaps
    If lightmap <> -1 Then
        maps = 0
        Do While maps < MAXLIGHTMAPS
            If mod_surfaces(surf).styles(maps) = 255 Then Exit Do
            scale_ = r_drawsurf.lightadj(maps)   ' 8.8 fraction
            For i = 0 To size - 1
                blocklights(i) = blocklights(i) + mod_lightdata(lightmap + i) * scale_
            Next
            lightmap = lightmap + size   ' skip to next lightmap
            maps = maps + 1
        Loop
    End If

    ' add all the dynamic lights
    If mod_surfaces(surf).dlightframe = r_framecount Then R_AddDynamicLights

    ' bound, invert, and shift
    For i = 0 To size - 1
        t = ((255& * 256 - blocklights(i)) And &HFFFFFFFC) \ 4   ' >> (8 - VID_CBITS)

        If t < 64 Then t = 64            ' (1 << 6)

        blocklights(i) = t
    Next
End Sub

' Returns the proper texture for a given time and base texture
' (textures() indices)
Public Function R_TextureAnimation(ByVal base As Long) As Long
    Dim reletive As Long, count As Long

    If cl_entities(currententity).frame Then
        If textures(base).alternate_anims Then base = textures(base).alternate_anims
    End If

    If textures(base).anim_total = 0 Then
        R_TextureAnimation = base
        Exit Function
    End If

    reletive = ToInt(cl.time * 10) Mod textures(base).anim_total

    count = 0
    Do While textures(base).anim_min > reletive Or textures(base).anim_max <= reletive
        base = textures(base).anim_next
        If base = 0 Then Sys_Error "R_TextureAnimation: broken cycle"
        count = count + 1
        If count > 100 Then Sys_Error "R_TextureAnimation: infinite cycle"
    Loop

    R_TextureAnimation = base
End Function

Public Sub R_DrawSurface()
    Dim basetptr As Long, smax As Long, tmax As Long, twidth As Long, u As Long
    Dim soffset As Long, basetoffset As Long, texwidth As Long, horzblockstep As Long
    Dim pcolumndest As Long, mt As Long, surfmip As Long

    ' calculate the lightings
    R_BuildLightMap

    surfrowbytes = r_drawsurf.rowbytes

    mt = r_drawsurf.texture
    surfmip = r_drawsurf.surfmip

    r_source = textures(mt).offsets(surfmip)

    ' the fractional light values should range from 0 to (VID_GRADES - 1) << 16
    ' from a source range of 0 - 255

    texwidth = textures(mt).width \ BIT_(surfmip)

    blocksize = 16 \ BIT_(surfmip)
    blockdivshift = 4 - surfmip
    blockdivmask = BIT_(blockdivshift) - 1

    r_lightwidth = (mod_surfaces(r_drawsurf.surf).extents(0) \ 16) + 1

    r_numhblocks = r_drawsurf.surfwidth \ BIT_(blockdivshift)
    r_numvblocks = r_drawsurf.surfheight \ BIT_(blockdivshift)

    '==============================

    ' (r_pixbytes is 1: pblockdrawer = surfmiptable[r_drawsurf.surfmip])
    ' TODO: only needs to be set when there is a display settings change
    horzblockstep = blocksize

    smax = textures(mt).width \ BIT_(surfmip)
    twidth = texwidth
    tmax = textures(mt).height \ BIT_(surfmip)
    sourcetstep = texwidth
    r_stepback = tmax * twidth

    r_sourcemax = r_source + (tmax * smax)

    soffset = mod_surfaces(r_drawsurf.surf).texturemins(0)
    basetoffset = mod_surfaces(r_drawsurf.surf).texturemins(1)

    ' << 16 components are to guarantee positive values for %
    soffset = (SAR(soffset, surfmip) + (smax * &H10000)) Mod smax
    basetptr = r_source + (((SAR(basetoffset, surfmip) + (tmax * &H10000)) Mod tmax) * twidth)

    pcolumndest = r_drawsurf.surfdat

    For u = 0 To r_numhblocks - 1
        r_lightptr = u

        prowdestbase = pcolumndest

        pbasesource = basetptr + soffset

        R_DrawSurfaceBlock8 blocksize

        soffset = soffset + blocksize
        If soffset >= smax Then soffset = 0

        pcolumndest = pcolumndest + horzblockstep
    Next
End Sub

'=============================================================================

' C: R_DrawSurfaceBlock8_mip0 (n 16, C shifts by 4), _mip1 (8, 3), _mip2 (4, 2),
' _mip3 (2, 1)
Private Sub R_DrawSurfaceBlock8(ByVal n As Long)
    Dim v As Long, i As Long, b As Long, lightstep As Long, lighttemp As Long, light As Long
    Dim pix As Byte, psource As Long, prowdest As Long, mask As Long

    mask = Not (n - 1)                   ' for C's >> shift of negative values

    psource = pbasesource
    prowdest = prowdestbase

    For v = 0 To r_numvblocks - 1
        ' FIXME: make these locals?
        ' FIXME: use delta rather than both right and left, like ASM?
        lightleft = blocklights(r_lightptr)
        lightright = blocklights(r_lightptr + 1)
        r_lightptr = r_lightptr + r_lightwidth
        lightleftstep = ((blocklights(r_lightptr) - lightleft) And mask) \ n
        lightrightstep = ((blocklights(r_lightptr + 1) - lightright) And mask) \ n

        For i = 0 To n - 1
            lighttemp = lightleft - lightright
            lightstep = (lighttemp And mask) \ n

            light = lightright

            For b = n - 1 To 0 Step -1
                pix = texpixels(psource + b)
                sc_bytes(prowdest + b) = vid_colormap((light And &HFF00&) + pix)
                light = light + lightstep
            Next

            psource = psource + sourcetstep
            lightright = lightright + lightrightstep
            lightleft = lightleft + lightleftstep
            prowdest = prowdest + surfrowbytes
        Next

        If psource >= r_sourcemax Then psource = psource - r_stepback
    Next
End Sub
