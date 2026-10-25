Attribute VB_Name = "d_scan"
' d_scan.bas - portable C scan-level rasterization code, all pixel depths
' (d_scan.c)
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private r_turb_pdest As Long             ' d_viewbuffer() offset (C: r_turb_pbase is cacheblock)
Private r_turb_s As Long, r_turb_t As Long, r_turb_sstep As Long, r_turb_tstep As Long
Private r_turb_turb As Long              ' sintable() offset
Private r_turb_spancount As Long

' this performs a slight compression of the screen at the same time as
' the sine warp, to keep the edges from wrapping
Public Sub D_WarpScreen()
    Dim w As Long, h As Long, u As Long, v As Long, dest As Long, turb As Long
    Dim col As Long, row As Long, wratio As Single, hratio As Single
    Dim rowptr(0 To MAXHEIGHT + AMP2 * 2 - 1) As Long    ' d_viewbuffer() (warp buffer) offsets
    Dim column(0 To MAXWIDTH + AMP2 * 2 - 1) As Long

    w = r_refdef.vrect.width
    h = r_refdef.vrect.height

    wratio = w / CSng(scr_vrect.width)
    hratio = h / CSng(scr_vrect.height)

    For v = 0 To scr_vrect.height + AMP2 * 2 - 1
        rowptr(v) = (r_refdef.vrect.y * screenwidth) + (screenwidth * ToInt(CSng(v) * hratio * h / (h + AMP2 * 2)))
    Next

    For u = 0 To scr_vrect.width + AMP2 * 2 - 1
        column(u) = r_refdef.vrect.x + ToInt(CSng(u) * wratio * w / (w + AMP2 * 2))
    Next

    turb = ToInt(cl.time * SPEED_) And (CYCLE - 1)
    dest = vid_bufferofs + scr_vrect.y * vid.rowbytes + scr_vrect.x

    For v = 0 To scr_vrect.height - 1
        col = intsintable(turb + v)      ' &column[turb[v]]
        row = v                          ' &rowptr[v]
        For u = 0 To scr_vrect.width - 1
            vid_buffer(dest + u) = d_viewbuffer(rowptr(row + intsintable(turb + u)) + column(col + u))
        Next
        dest = dest + vid.rowbytes
    Next
End Sub

Private Sub D_DrawTurbulent8Span()
    Dim sturb As Long, tturb As Long

    Do
        sturb = (((r_turb_s + sintable(r_turb_turb + ((r_turb_t And &HFFFF0000) \ &H10000 And (CYCLE - 1)))) And &HFFFF0000) \ &H10000) And 63
        tturb = (((r_turb_t + sintable(r_turb_turb + ((r_turb_s And &HFFFF0000) \ &H10000 And (CYCLE - 1)))) And &HFFFF0000) \ &H10000) And 63
        d_viewbuffer(r_turb_pdest) = cacheblock(tturb * 64 + sturb)
        r_turb_pdest = r_turb_pdest + 1
        r_turb_s = r_turb_s + r_turb_sstep
        r_turb_t = r_turb_t + r_turb_tstep
        r_turb_spancount = r_turb_spancount - 1
    Loop While r_turb_spancount > 0
End Sub

Public Sub Turbulent8(ByVal pspan As Long)
    Dim count As Long, snext As Long, tnext As Long
    Dim sdivz As Single, tdivz As Single, zi As Single, z As Single, du As Single, dv As Single, spancountminus1 As Single
    Dim sdivz16stepu As Single, tdivz16stepu As Single, zi16stepu As Single

    r_turb_turb = ToInt(cl.time * SPEED_) And (CYCLE - 1)

    r_turb_sstep = 0                     ' keep compiler happy
    r_turb_tstep = 0                     ' ditto

    sdivz16stepu = d_sdivzstepu * 16
    tdivz16stepu = d_tdivzstepu * 16
    zi16stepu = d_zistepu * 16

    Do
        r_turb_pdest = screenwidth * spans(pspan).v + spans(pspan).u

        count = spans(pspan).count

        ' calculate the initial s/z, t/z, 1/z, s, and t and clamp
        du = spans(pspan).u
        dv = spans(pspan).v

        sdivz = d_sdivzorigin + dv * d_sdivzstepv + du * d_sdivzstepu
        tdivz = d_tdivzorigin + dv * d_tdivzstepv + du * d_tdivzstepu
        zi = d_ziorigin + dv * d_zistepv + du * d_zistepu
        z = 65536! / zi                  ' prescale to 16.16 fixed-point

        r_turb_s = ToInt(sdivz * z) + sadjust
        If r_turb_s > bbextents Then
            r_turb_s = bbextents
        ElseIf r_turb_s < 0 Then
            r_turb_s = 0
        End If

        r_turb_t = ToInt(tdivz * z) + tadjust
        If r_turb_t > bbextentt Then
            r_turb_t = bbextentt
        ElseIf r_turb_t < 0 Then
            r_turb_t = 0
        End If

        Do
            ' calculate s and t at the far end of the span
            If count >= 16 Then
                r_turb_spancount = 16
            Else
                r_turb_spancount = count
            End If

            count = count - r_turb_spancount

            If count Then
                ' calculate s/z, t/z, zi->fixed s and t at far end of span,
                ' calculate s and t steps across span by shifting
                sdivz = sdivz + sdivz16stepu
                tdivz = tdivz + tdivz16stepu
                zi = zi + zi16stepu
                z = 65536! / zi          ' prescale to 16.16 fixed-point

                snext = ToInt(sdivz * z) + sadjust
                If snext > bbextents Then
                    snext = bbextents
                ElseIf snext < 16 Then
                    snext = 16           ' prevent round-off error on <0 steps from
                                         '  from causing overstepping & running off the
                                         '  edge of the texture
                End If

                tnext = ToInt(tdivz * z) + tadjust
                If tnext > bbextentt Then
                    tnext = bbextentt
                ElseIf tnext < 16 Then
                    tnext = 16           ' guard against round-off error on <0 steps
                End If

                r_turb_sstep = ((snext - r_turb_s) And &HFFFFFFF0) \ 16   ' >> 4
                r_turb_tstep = ((tnext - r_turb_t) And &HFFFFFFF0) \ 16
            Else
                ' calculate s/z, t/z, zi->fixed s and t at last pixel in span (so
                ' can't step off polygon), clamp, calculate s and t steps across
                ' span by division, biasing steps low so we don't run off the
                ' texture
                spancountminus1 = r_turb_spancount - 1
                sdivz = sdivz + d_sdivzstepu * spancountminus1
                tdivz = tdivz + d_tdivzstepu * spancountminus1
                zi = zi + d_zistepu * spancountminus1
                z = 65536! / zi          ' prescale to 16.16 fixed-point
                snext = ToInt(sdivz * z) + sadjust
                If snext > bbextents Then
                    snext = bbextents
                ElseIf snext < 16 Then
                    snext = 16           ' prevent round-off error on <0 steps from
                                         '  from causing overstepping & running off the
                                         '  edge of the texture
                End If

                tnext = ToInt(tdivz * z) + tadjust
                If tnext > bbextentt Then
                    tnext = bbextentt
                ElseIf tnext < 16 Then
                    tnext = 16           ' guard against round-off error on <0 steps
                End If

                If r_turb_spancount > 1 Then
                    r_turb_sstep = (snext - r_turb_s) \ (r_turb_spancount - 1)
                    r_turb_tstep = (tnext - r_turb_t) \ (r_turb_spancount - 1)
                End If
            End If

            r_turb_s = r_turb_s And ((CYCLE * &H10000) - 1)
            r_turb_t = r_turb_t And ((CYCLE * &H10000) - 1)

            D_DrawTurbulent8Span

            r_turb_s = snext
            r_turb_t = tnext
        Loop While count > 0

        pspan = spans(pspan).pnext
    Loop While pspan <> -1
End Sub

Public Sub D_DrawSpans8(ByVal pspan As Long)
    Dim count As Long, spancount As Long, pdest As Long
    Dim s As Long, t As Long, snext As Long, tnext As Long, sstep As Long, tstep As Long
    Dim sdivz As Single, tdivz As Single, zi As Single, z As Single, du As Single, dv As Single, spancountminus1 As Single
    Dim sdivz8stepu As Single, tdivz8stepu As Single, zi8stepu As Single
    Dim cw As Long

    sstep = 0                            ' keep compiler happy
    tstep = 0                            ' ditto

    cw = cachewidth

    sdivz8stepu = d_sdivzstepu * 8
    tdivz8stepu = d_tdivzstepu * 8
    zi8stepu = d_zistepu * 8

    Do
        pdest = screenwidth * spans(pspan).v + spans(pspan).u

        count = spans(pspan).count

        ' calculate the initial s/z, t/z, 1/z, s, and t and clamp
        du = spans(pspan).u
        dv = spans(pspan).v

        sdivz = d_sdivzorigin + dv * d_sdivzstepv + du * d_sdivzstepu
        tdivz = d_tdivzorigin + dv * d_tdivzstepv + du * d_tdivzstepu
        zi = d_ziorigin + dv * d_zistepv + du * d_zistepu
        z = 65536! / zi                  ' prescale to 16.16 fixed-point

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
                d_viewbuffer(pdest) = cacheblock(s \ 65536 + (t \ 65536) * cw)
                pdest = pdest + 1
                s = s + sstep
                t = t + tstep
                spancount = spancount - 1
            Loop While spancount > 0

            s = snext
            t = tnext
        Loop While count > 0

        pspan = spans(pspan).pnext
    Loop While pspan <> -1
End Sub

Public Sub D_DrawZSpans(ByVal pspan As Long)
    Dim count As Long, izistep As Long, izi As Long, pdest As Long, zi As Double, du As Single, dv As Single

    ' FIXME: check for clamping/range problems
    ' we count on FP exceptions being turned off to avoid range problems
    izistep = ToInt(d_zistepu * &H8000& * &H10000)

    Do
        pdest = d_zwidth * spans(pspan).v + spans(pspan).u

        count = spans(pspan).count

        ' calculate the initial 1/z
        du = spans(pspan).u
        dv = spans(pspan).v

        zi = d_ziorigin + dv * d_zistepv + du * d_zistepu
        ' we count on FP exceptions being turned off to avoid range problems
        izi = ToInt(zi * &H8000& * &H10000)

        ' (C writes two shorts at a time where it can; the values are the same:
        ' (short)(izi >> 16) each)
        Do While count > 0
            d_pzbuffer(pdest) = (izi And &HFFFF0000) \ &H10000
            izi = izi + izistep
            pdest = pdest + 1
            count = count - 1
        Loop

        pspan = spans(pspan).pnext
    Loop While pspan <> -1
End Sub
