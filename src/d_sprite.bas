Attribute VB_Name = "d_sprite"
' d_sprite.bas - software top-level rasterization driver module for drawing
' sprites (d_sprite.c)
'
' The sprite spans are sprite_spans() (C: D_DrawSprite's local spans) and the
' sspan_t pointers indices there; cacheblock() is laid over the frame's
' pixels in sprite_pixels().
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private sprite_height As Long
Private minindex As Long, maxindex As Long
Private sprite_spans(0 To MAXHEIGHT) As sspan_t

' (QuakeVB) C: ceil
Private Function Ceil_(ByVal x As Double) As Double
    Ceil_ = -Int(-x)
End Function

Private Sub D_SpriteDrawSpans(ByVal pspan As Long)
    Dim count As Long, spancount As Long, izistep As Long, izi As Long, pdest As Long, pz As Long
    Dim s As Long, t As Long, snext As Long, tnext As Long, sstep As Long, tstep As Long
    Dim sdivz As Single, tdivz As Single, zi As Single, z As Single, du As Single, dv As Single, spancountminus1 As Single
    Dim sdivz8stepu As Single, tdivz8stepu As Single, zi8stepu As Single
    Dim btemp As Byte, cw As Long

    sstep = 0                            ' keep compiler happy
    tstep = 0                            ' ditto

    cw = cachewidth                      ' (C: pbase = cacheblock)

    sdivz8stepu = d_sdivzstepu * 8
    tdivz8stepu = d_tdivzstepu * 8
    zi8stepu = d_zistepu * 8

    ' we count on FP exceptions being turned off to avoid range problems
    izistep = ToInt(d_zistepu * &H8000& * &H10000)

    Do
        pdest = screenwidth * sprite_spans(pspan).v + sprite_spans(pspan).u
        pz = d_zwidth * sprite_spans(pspan).v + sprite_spans(pspan).u

        count = sprite_spans(pspan).count

        If count <= 0 Then GoTo NextSpan

        ' calculate the initial s/z, t/z, 1/z, s, and t and clamp
        du = sprite_spans(pspan).u
        dv = sprite_spans(pspan).v

        sdivz = d_sdivzorigin + dv * d_sdivzstepv + du * d_sdivzstepu
        tdivz = d_tdivzorigin + dv * d_tdivzstepv + du * d_tdivzstepu
        zi = d_ziorigin + dv * d_zistepv + du * d_zistepu
        z = 65536! / zi                  ' prescale to 16.16 fixed-point
        ' we count on FP exceptions being turned off to avoid range problems
        izi = ToInt(zi * &H8000& * &H10000)

        s = ToInt(sdivz * z) + sadjust
        If s > bbextents Then
            s = bbextents
        ElseIf s < 0 Then
            s = 0
        End If

        t = ToInt(tdivz * z) + tadjust
        If t > bbextentt Then
            t = bbextentt
        ElseIf t < 0 Then
            t = 0
        End If

        Do
            ' calculate s and t at the far end of the span
            If count >= 8 Then
                spancount = 8
            Else
                spancount = count
            End If

            count = count - spancount

            If count Then
                ' calculate s/z, t/z, zi->fixed s and t at far end of span,
                ' calculate s and t steps across span by shifting
                sdivz = sdivz + sdivz8stepu
                tdivz = tdivz + tdivz8stepu
                zi = zi + zi8stepu
                z = 65536! / zi          ' prescale to 16.16 fixed-point

                snext = ToInt(sdivz * z) + sadjust
                If snext > bbextents Then
                    snext = bbextents
                ElseIf snext < 8 Then
                    snext = 8            ' prevent round-off error on <0 steps from
                                         '  from causing overstepping & running off the
                                         '  edge of the texture
                End If

                tnext = ToInt(tdivz * z) + tadjust
                If tnext > bbextentt Then
                    tnext = bbextentt
                ElseIf tnext < 8 Then
                    tnext = 8            ' guard against round-off error on <0 steps
                End If

                sstep = ((snext - s) And &HFFFFFFF8) \ 8   ' >> 3
                tstep = ((tnext - t) And &HFFFFFFF8) \ 8
            Else
                ' calculate s/z, t/z, zi->fixed s and t at last pixel in span (so
                ' can't step off polygon), clamp, calculate s and t steps across
                ' span by division, biasing steps low so we don't run off the
                ' texture
                spancountminus1 = spancount - 1
                sdivz = sdivz + d_sdivzstepu * spancountminus1
                tdivz = tdivz + d_tdivzstepu * spancountminus1
                zi = zi + d_zistepu * spancountminus1
                z = 65536! / zi          ' prescale to 16.16 fixed-point
                snext = ToInt(sdivz * z) + sadjust
                If snext > bbextents Then
                    snext = bbextents
                ElseIf snext < 8 Then
                    snext = 8            ' prevent round-off error on <0 steps from
                                         '  from causing overstepping & running off the
                                         '  edge of the texture
                End If

                tnext = ToInt(tdivz * z) + tadjust
                If tnext > bbextentt Then
                    tnext = bbextentt
                ElseIf tnext < 8 Then
                    tnext = 8            ' guard against round-off error on <0 steps
                End If

                If spancount > 1 Then
                    sstep = (snext - s) \ (spancount - 1)
                    tstep = (tnext - t) \ (spancount - 1)
                End If
            End If

            ' (s and t stay in 0 .. bbextents: \ 65536 is C's >> 16 there)
            Do
                btemp = cacheblock(s \ 65536 + (t \ 65536) * cw)
                If btemp <> 255 Then
                    If d_pzbuffer(pz) <= (izi And &HFFFF0000) \ &H10000 Then
                        d_pzbuffer(pz) = (izi And &HFFFF0000) \ &H10000
                        d_viewbuffer(pdest) = btemp
                    End If
                End If

                izi = izi + izistep
                pdest = pdest + 1
                pz = pz + 1
                s = s + sstep
                t = t + tstep
                spancount = spancount - 1
            Loop While spancount > 0

            s = snext
            t = tnext
        Loop While count > 0

NextSpan:
        pspan = pspan + 1
    Loop While sprite_spans(pspan).count <> DS_SPAN_LIST_END
End Sub

Private Sub D_SpriteScanLeftEdge()
    Dim i As Long, v As Long, itop As Long, ibottom As Long, lmaxindex As Long
    Dim pvert As Long, pnext As Long, pspan As Long
    Dim du As Single, dv As Single, vtop As Single, vbottom As Single, slope As Single
    Dim u As Long, u_step As Long

    pspan = 0
    i = minindex
    If i = 0 Then i = r_spritedesc.nump

    lmaxindex = maxindex
    If lmaxindex = 0 Then lmaxindex = r_spritedesc.nump

    vtop = Ceil_(r_spriteverts(i).v)

    Do
        pvert = i
        pnext = pvert - 1

        vbottom = Ceil_(r_spriteverts(pnext).v)

        If vtop < vbottom Then
            du = r_spriteverts(pnext).u - r_spriteverts(pvert).u
            dv = r_spriteverts(pnext).v - r_spriteverts(pvert).v
            slope = du / dv
            u_step = ToInt(slope * &H10000)
            ' adjust u to ceil the integer portion
            u = ToInt((r_spriteverts(pvert).u + (slope * (vtop - r_spriteverts(pvert).v))) * &H10000) + (&H10000 - 1)
            itop = ToInt(vtop)
            ibottom = ToInt(vbottom)

            For v = itop To ibottom - 1
                sprite_spans(pspan).u = SAR(u, 16)
                sprite_spans(pspan).v = v
                u = u + u_step
                pspan = pspan + 1
            Next
        End If

        vtop = vbottom

        i = i - 1
        If i = 0 Then i = r_spritedesc.nump

    Loop While i <> lmaxindex
End Sub

Private Sub D_SpriteScanRightEdge()
    Dim i As Long, v As Long, itop As Long, ibottom As Long, pvert As Long, pnext As Long, pspan As Long
    Dim du As Single, dv As Single, vtop As Single, vbottom As Single, slope As Single
    Dim uvert As Single, unext As Single, vvert As Single, vnext As Single
    Dim u As Long, u_step As Long

    pspan = 0
    i = minindex

    vvert = r_spriteverts(i).v
    If vvert < r_refdef.fvrecty_adj Then vvert = r_refdef.fvrecty_adj
    If vvert > r_refdef.fvrectbottom_adj Then vvert = r_refdef.fvrectbottom_adj

    vtop = Ceil_(vvert)

    Do
        pvert = i
        pnext = pvert + 1

        vnext = r_spriteverts(pnext).v
        If vnext < r_refdef.fvrecty_adj Then vnext = r_refdef.fvrecty_adj
        If vnext > r_refdef.fvrectbottom_adj Then vnext = r_refdef.fvrectbottom_adj

        vbottom = Ceil_(vnext)

        If vtop < vbottom Then
            uvert = r_spriteverts(pvert).u
            If uvert < r_refdef.fvrectx_adj Then uvert = r_refdef.fvrectx_adj
            If uvert > r_refdef.fvrectright_adj Then uvert = r_refdef.fvrectright_adj

            unext = r_spriteverts(pnext).u
            If unext < r_refdef.fvrectx_adj Then unext = r_refdef.fvrectx_adj
            If unext > r_refdef.fvrectright_adj Then unext = r_refdef.fvrectright_adj

            du = unext - uvert
            dv = vnext - vvert
            slope = du / dv
            u_step = ToInt(slope * &H10000)
            ' adjust u to ceil the integer portion
            u = ToInt((uvert + (slope * (vtop - vvert))) * &H10000) + (&H10000 - 1)
            itop = ToInt(vtop)
            ibottom = ToInt(vbottom)

            For v = itop To ibottom - 1
                sprite_spans(pspan).count = SAR(u, 16) - sprite_spans(pspan).u
                u = u + u_step
                pspan = pspan + 1
            Next
        End If

        vtop = vbottom
        vvert = vnext

        i = i + 1
        If i = r_spritedesc.nump Then i = 0

    Loop While i <> maxindex

    sprite_spans(pspan).count = DS_SPAN_LIST_END   ' mark the end of the span list
End Sub

Private Sub D_SpriteCalculateGradients()
    Dim p_normal As vec3_t, p_saxis As vec3_t, p_taxis As vec3_t, p_temp1 As vec3_t, distinv As Single

    TransformVector r_spritedesc.vpn, p_normal
    TransformVector r_spritedesc.vright, p_saxis
    TransformVector r_spritedesc.vup, p_taxis
    VectorInverse p_taxis

    distinv = 1! / (-DotProduct(modelorg, r_spritedesc.vpn))

    d_sdivzstepu = p_saxis.v(0) * xscaleinv
    d_tdivzstepu = p_taxis.v(0) * xscaleinv

    d_sdivzstepv = -p_saxis.v(1) * yscaleinv
    d_tdivzstepv = -p_taxis.v(1) * yscaleinv

    d_zistepu = p_normal.v(0) * xscaleinv * distinv
    d_zistepv = -p_normal.v(1) * yscaleinv * distinv

    d_sdivzorigin = p_saxis.v(2) - xcenter * d_sdivzstepu - ycenter * d_sdivzstepv
    d_tdivzorigin = p_taxis.v(2) - xcenter * d_tdivzstepu - ycenter * d_tdivzstepv
    d_ziorigin = p_normal.v(2) * distinv - xcenter * d_zistepu - ycenter * d_zistepv

    TransformVector modelorg, p_temp1

    sadjust = ToInt(DotProduct(p_temp1, p_saxis) * &H10000 + 0.5) - (-(cachewidth \ 2) * &H10000)
    tadjust = ToInt(DotProduct(p_temp1, p_taxis) * &H10000 + 0.5) - (-(sprite_height \ 2) * &H10000)

    ' -1 (-epsilon) so we never wander off the edge of the texture
    bbextents = cachewidth * &H10000 - 1
    bbextentt = sprite_height * &H10000 - 1
End Sub

Public Sub D_DrawSprite()
    Dim i As Long, nump As Long, ymin As Single, ymax As Single, f As Long

    ' find the top and bottom vertices, and make sure there's at least one scan to
    ' draw
    ymin = 999999.9
    ymax = -999999.9

    For i = 0 To r_spritedesc.nump - 1
        If r_spriteverts(i).v < ymin Then
            ymin = r_spriteverts(i).v
            minindex = i
        End If

        If r_spriteverts(i).v > ymax Then
            ymax = r_spriteverts(i).v
            maxindex = i
        End If
    Next

    ymin = Ceil_(ymin)
    ymax = Ceil_(ymax)

    If ymin >= ymax Then Exit Sub        ' doesn't cross any scans at all

    f = r_spritedesc.pspriteframe
    cachewidth = sprite_frames(f).width
    sprite_height = sprite_frames(f).height
    D_SetCacheBlock VarPtr(sprite_pixels(sprite_frames(f).pixels)), cachewidth * sprite_height

    ' copy the first vertex to the last vertex, so we don't have to deal with
    ' wrapping
    nump = r_spritedesc.nump
    r_spriteverts(nump) = r_spriteverts(0)

    D_SpriteCalculateGradients
    D_SpriteScanLeftEdge
    D_SpriteScanRightEdge
    D_SpriteDrawSpans 0
End Sub
