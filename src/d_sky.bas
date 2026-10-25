Attribute VB_Name = "d_sky"
' d_sky.bas - the sky spans (d_sky.c)
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const SKY_SPAN_SHIFT As Long = 5
Private Const SKY_SPAN_MAX As Long = 32  ' 1 << SKY_SPAN_SHIFT

Private Sub D_Sky_uv_To_st(ByVal u As Long, ByVal v As Long, s As Long, t As Long)
    Dim wu As Single, wv As Single, temp As Single, end_ As vec3_t

    If r_refdef.vrect.width >= r_refdef.vrect.height Then
        temp = r_refdef.vrect.width
    Else
        temp = r_refdef.vrect.height
    End If

    wu = 8192# * CSng(u - (vid.width \ 2)) / temp
    wv = 8192# * CSng((vid.height \ 2) - v) / temp

    end_.v(0) = 4096 * vpn.v(0) + wu * vright.v(0) + wv * vup.v(0)
    end_.v(1) = 4096 * vpn.v(1) + wu * vright.v(1) + wv * vup.v(1)
    end_.v(2) = 4096 * vpn.v(2) + wu * vright.v(2) + wv * vup.v(2)
    end_.v(2) = end_.v(2) * 3
    VectorNormalize end_

    temp = skytime * skyspeed            ' TODO: add D_SetupFrame & set this there
    s = ToInt((temp + 6 * (SKYSIZE \ 2 - 1) * end_.v(0)) * &H10000)
    t = ToInt((temp + 6 * (SKYSIZE \ 2 - 1) * end_.v(1)) * &H10000)
End Sub

Public Sub D_DrawSkyScans8(ByVal pspan As Long)
    Dim count As Long, spancount As Long, u As Long, v As Long, pdest As Long
    Dim s As Long, t As Long, snext As Long, tnext As Long, sstep As Long, tstep As Long
    Dim spancountminus1 As Long

    sstep = 0                            ' keep compiler happy
    tstep = 0                            ' ditto

    Do
        pdest = screenwidth * spans(pspan).v + spans(pspan).u

        count = spans(pspan).count

        ' calculate the initial s & t
        u = spans(pspan).u
        v = spans(pspan).v
        D_Sky_uv_To_st u, v, s, t

        Do
            If count >= SKY_SPAN_MAX Then
                spancount = SKY_SPAN_MAX
            Else
                spancount = count
            End If

            count = count - spancount

            If count Then
                u = u + spancount

                ' calculate s and t at far end of span,
                ' calculate s and t steps across span by shifting
                D_Sky_uv_To_st u, v, snext, tnext

                sstep = ((snext - s) And &HFFFFFFE0) \ SKY_SPAN_MAX   ' >> SKY_SPAN_SHIFT
                tstep = ((tnext - t) And &HFFFFFFE0) \ SKY_SPAN_MAX
            Else
                ' calculate s and t at last pixel in span,
                ' calculate s and t steps across span by division
                spancountminus1 = spancount - 1

                If spancountminus1 > 0 Then
                    u = u + spancountminus1
                    D_Sky_uv_To_st u, v, snext, tnext

                    sstep = (snext - s) \ spancountminus1
                    tstep = (tnext - t) \ spancountminus1
                End If
            End If

            Do
                d_viewbuffer(pdest) = newsky((t And R_SKY_TMASK) \ &H100& + (s And R_SKY_SMASK) \ &H10000)
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
