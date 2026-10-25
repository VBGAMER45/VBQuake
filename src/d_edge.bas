Attribute VB_Name = "d_edge"
' d_edge.bas - draws the surfaces' spans (d_edge.c)
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private miplevel As Long

Public scale_for_mip As Single
Public screenwidth As Long
Public ubasestep As Long, errorterm As Long, erroradjustup As Long, erroradjustdown As Long
Public vstartscan As Long

Public transformed_modelorg As vec3_t

Public Sub D_DrawPoly()
    ' this driver takes spans, not polygons
End Sub

Public Function D_MipLevelForScale(ByVal scale_ As Single) As Long
    Dim lmiplevel As Long

    If scale_ >= d_scalemip(0) Then
        lmiplevel = 0
    ElseIf scale_ >= d_scalemip(1) Then
        lmiplevel = 1
    ElseIf scale_ >= d_scalemip(2) Then
        lmiplevel = 2
    Else
        lmiplevel = 3
    End If

    If lmiplevel < d_minmip Then lmiplevel = d_minmip

    D_MipLevelForScale = lmiplevel
End Function

' FIXME: clean this up
Private Sub D_DrawSolidSurface(ByVal surf As Long, ByVal color As Long)
    Dim span As Long, pdest As Long, u As Long, u2 As Long, pix As Byte

    pix = color And &HFF&
    span = surfaces(surf).spans
    Do While span <> -1
        pdest = screenwidth * spans(span).v
        u = spans(span).u
        u2 = spans(span).u + spans(span).count - 1
        ' (C writes four bytes at a time where it can)
        For u = u To u2
            d_viewbuffer(pdest + u) = pix
        Next
        span = spans(span).pnext
    Loop
End Sub

Private Sub D_CalcGradients(ByVal pface As Long)
    Dim mipscale As Single, p_temp1 As vec3_t, p_saxis As vec3_t, p_taxis As vec3_t, t As Single
    Dim tex As Long, i As Long, sv As vec3_t, tv As vec3_t

    mipscale = 1# / BIT_(miplevel)

    tex = mod_surfaces(pface).texinfo
    For i = 0 To 2
        sv.v(i) = mod_texinfo(tex).vecs(0, i)
        tv.v(i) = mod_texinfo(tex).vecs(1, i)
    Next
    TransformVector sv, p_saxis
    TransformVector tv, p_taxis

    t = xscaleinv * mipscale
    d_sdivzstepu = p_saxis.v(0) * t
    d_tdivzstepu = p_taxis.v(0) * t

    t = yscaleinv * mipscale
    d_sdivzstepv = -p_saxis.v(1) * t
    d_tdivzstepv = -p_taxis.v(1) * t

    d_sdivzorigin = p_saxis.v(2) * mipscale - xcenter * d_sdivzstepu - ycenter * d_sdivzstepv
    d_tdivzorigin = p_taxis.v(2) * mipscale - xcenter * d_tdivzstepu - ycenter * d_tdivzstepv

    VectorScale transformed_modelorg, mipscale, p_temp1

    t = &H10000 * mipscale
    sadjust = ToInt(CDbl(ToInt(DotProduct(p_temp1, p_saxis) * &H10000 + 0.5)) _
        - SAR(mod_surfaces(pface).texturemins(0) * &H10000, miplevel) _
        + mod_texinfo(tex).vecs(0, 3) * t)
    tadjust = ToInt(CDbl(ToInt(DotProduct(p_temp1, p_taxis) * &H10000 + 0.5)) _
        - SAR(mod_surfaces(pface).texturemins(1) * &H10000, miplevel) _
        + mod_texinfo(tex).vecs(1, 3) * t)

    '
    ' -1 (-epsilon) so we never wander off the edge of the texture
    '
    bbextents = SAR(mod_surfaces(pface).extents(0) * &H10000, miplevel) - 1
    bbextentt = SAR(mod_surfaces(pface).extents(1) * &H10000, miplevel) - 1
End Sub

' (QuakeVB) restores the world's drawing state after a submodel surface
Private Sub RestoreWorldState(world_transformed_modelorg As vec3_t)
    '
    ' restore the old drawing state
    ' FIXME: we don't want to do this every time!
    ' TODO: speed up
    '
    currententity = 0
    transformed_modelorg = world_transformed_modelorg
    vpn = base_vpn
    vup = base_vup
    vright = base_vright
    modelorg = base_modelorg
    R_TransformFrustum
End Sub

' (QuakeVB) the submodel's drawing state
Private Sub SetSubmodelState(ByVal s As Long)
    Dim local_modelorg As vec3_t

    ' FIXME: we don't want to do all this for every polygon!
    ' TODO: store once at start of frame
    currententity = surfaces(s).entity   ' FIXME: make this passed in to
                                         ' R_RotateBmodel ()
    VectorSubtract r_origin, cl_entities(currententity).origin, local_modelorg
    TransformVector local_modelorg, transformed_modelorg

    R_RotateBmodel                       ' FIXME: don't mess with the frustum,
                                         ' make entity passed in
End Sub

Public Sub D_DrawSurfaces()
    Dim s As Long, pface As Long, pcurrentcache As Long, tx As Long, ofs As Long, n As Long
    Dim world_transformed_modelorg As vec3_t

    currententity = 0
    TransformVector modelorg, transformed_modelorg
    world_transformed_modelorg = transformed_modelorg

    ' TODO: could preset a lot of this at mode set time
    If cvars(cv_r_drawflat).value <> 0 Then
        For s = 1 To surface_p - 1
            If surfaces(s).spans <> -1 Then
                d_zistepu = surfaces(s).d_zistepu
                d_zistepv = surfaces(s).d_zistepv
                d_ziorigin = surfaces(s).d_ziorigin

                ' (C takes the color from the bits of the msurface_t pointer)
                D_DrawSolidSurface s, surfaces(s).data And &HFF&
                D_DrawZSpans surfaces(s).spans
            End If
        Next
    Else
        For s = 1 To surface_p - 1
            If surfaces(s).spans = -1 Then GoTo nexts

            r_drawnpolycount = r_drawnpolycount + 1

            d_zistepu = surfaces(s).d_zistepu
            d_zistepv = surfaces(s).d_zistepv
            d_ziorigin = surfaces(s).d_ziorigin

            If surfaces(s).flags And SURF_DRAWSKY Then
                If r_skymade = 0 Then R_MakeSky

                D_DrawSkyScans8 surfaces(s).spans
                D_DrawZSpans surfaces(s).spans
            ElseIf surfaces(s).flags And SURF_DRAWBACKGROUND Then
                ' set up a gradient for the background surface that places it
                ' effectively at infinity distance from the viewpoint
                d_zistepu = 0
                d_zistepv = 0
                d_ziorigin = -0.9

                D_DrawSolidSurface s, ToInt(cvars(cv_r_clearcolor).value) And &HFF&
                D_DrawZSpans surfaces(s).spans
            ElseIf surfaces(s).flags And SURF_DRAWTURB Then
                pface = surfaces(s).data
                miplevel = 0
                ' (C: cacheblock is the texture's pixels, 64 by 64)
                tx = mod_texinfo(mod_surfaces(pface).texinfo).texture
                ofs = textures(tx).offsets(0)
                n = UBound(texpixels) + 1 - ofs
                If n > 64 * 64 Then n = 64 * 64
                D_SetCacheBlock VarPtr(texpixels(ofs)), n
                cachewidth = 64

                If surfaces(s).insubmodel Then SetSubmodelState s

                D_CalcGradients pface
                Turbulent8 surfaces(s).spans
                D_DrawZSpans surfaces(s).spans

                If surfaces(s).insubmodel Then RestoreWorldState world_transformed_modelorg
            Else
                If surfaces(s).insubmodel Then SetSubmodelState s

                pface = surfaces(s).data
                miplevel = D_MipLevelForScale(surfaces(s).nearzi * scale_for_mip * mod_texinfo(mod_surfaces(pface).texinfo).mipadjust)

                ' FIXME: make this passed in to D_CacheSurface
                pcurrentcache = D_CacheSurface(pface, miplevel)

                cachewidth = D_CacheWidth(pcurrentcache)
                D_SetCacheBlock D_CacheDataPtr(pcurrentcache), D_CacheDataBytes(pcurrentcache)

                D_CalcGradients pface

                D_DrawSpans8 surfaces(s).spans

                D_DrawZSpans surfaces(s).spans

                If surfaces(s).insubmodel Then RestoreWorldState world_transformed_modelorg
            End If
nexts:
        Next
    End If
End Sub
