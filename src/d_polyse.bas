Attribute VB_Name = "d_polyse"
' d_polyse.bas - routines for drawing sets of polygons sharing the same
' texture (used for Alias models) (d_polyse.c)
'
' r_p0, r_p1 and r_p2 are one array, r_p(), at RP0, RP1 and RP2; the edge
' tables' vertex pointers are those offsets (-1 = NULL). A spanpackage_t's
' pointers are offsets: pdest into d_viewbuffer(), pz into d_pzbuffer(), ptex
' into alias_skins(), like skintable() and d_pcolormap (into acolormap()).
' D_PolysetRecursiveTriangle works on the vertices in rtv(), 6 Longs each,
' with a stack of the split points after the triangle's own three.
' adivtab is computed (D_PolysetInit) instead of included from adivtab.h.
' Not ported: D_PolysetFillSpans8, which nothing calls.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' TODO: put in span spilling to shrink list size
' !!! if this is changed, it must be changed in d_polysa.s too !!!
Private Const DPS_MAXSPANS As Long = MAXHEIGHT + 1   ' 1 extra for spanpackage that marks end

' !!! if this is changed, it must be changed in asm_draw.h too !!!
Private Type spanpackage_t
    pdest As Long                        ' d_viewbuffer() offset
    pz As Long                           ' d_pzbuffer() offset
    count As Long
    ptex As Long                         ' alias_skins() offset
    sfrac As Long
    tfrac As Long
    light As Long
    zi As Long
End Type

Private Type edgetable
    isflattop As Long
    numleftedges As Long
    pleftedgevert0 As Long               ' r_p() offsets (RP0, RP1, RP2), -1 = NULL
    pleftedgevert1 As Long
    pleftedgevert2 As Long
    numrightedges As Long
    prightedgevert0 As Long
    prightedgevert1 As Long
    prightedgevert2 As Long
End Type

Private Const RP0 As Long = 0
Private Const RP1 As Long = 6
Private Const RP2 As Long = 12

Private r_p(0 To 17) As Long             ' C: int r_p0[6], r_p1[6], r_p2[6]

Private d_pcolormap As Long              ' acolormap() offset

Public d_aflatcolor As Long
Private d_xdenom As Long

Private pedgetable As Long               ' edgetables() index

Private edgetables(0 To 11) As edgetable

' FIXME: some of these can become statics
Private a_sstepxfrac As Long, a_tstepxfrac As Long, r_lstepx As Long, a_ststepxwhole As Long
Private r_sstepx As Long, r_tstepx As Long, r_lstepy As Long, r_sstepy As Long, r_tstepy As Long
Private r_zistepx As Long, r_zistepy As Long
Private d_aspancount As Long, d_countextrastep As Long

Private a_spans() As spanpackage_t       ' C: D_PolysetDraw's local spans
Private d_pedgespanpackage As Long       ' a_spans() index
Private ystart As Long
Private d_pdest As Long, d_ptex As Long
Private d_pz As Long
Private d_sfrac As Long, d_tfrac As Long, d_light As Long, d_zi As Long
Private d_ptexextrastep As Long, d_sfracextrastep As Long
Private d_tfracextrastep As Long, d_lightextrastep As Long, d_pdestextrastep As Long
Private d_lightbasestep As Long, d_pdestbasestep As Long, d_ptexbasestep As Long
Private d_sfracbasestep As Long, d_tfracbasestep As Long
Private d_ziextrastep As Long, d_zibasestep As Long
Private d_pzextrastep As Long, d_pzbasestep As Long

Private adivtab_quotient(0 To 32 * 32 - 1) As Long    ' C: adivtab_t adivtab[32*32]
Private adivtab_remainder(0 To 32 * 32 - 1) As Long

Private skintable(0 To MAX_LBM_HEIGHT - 1) As Long   ' alias_skins() offsets
Private skinwidth As Long
Private skinstart As Long                ' alias_skins() offset

Private rtv() As Long                    ' D_PolysetRecursiveTriangle's vertices
Private rt_top As Long                   ' the next free vertex in rtv()

Private polyse_made As Boolean

Private Sub SetEdgeTable(ByVal i As Long, ByVal isflattop As Long, ByVal numleftedges As Long, _
        ByVal l0 As Long, ByVal l1 As Long, ByVal l2 As Long, ByVal numrightedges As Long, _
        ByVal r0 As Long, ByVal r1 As Long, ByVal r2 As Long)
    edgetables(i).isflattop = isflattop
    edgetables(i).numleftedges = numleftedges
    edgetables(i).pleftedgevert0 = l0
    edgetables(i).pleftedgevert1 = l1
    edgetables(i).pleftedgevert2 = l2
    edgetables(i).numrightedges = numrightedges
    edgetables(i).prightedgevert0 = r0
    edgetables(i).prightedgevert1 = r1
    edgetables(i).prightedgevert2 = r2
End Sub

' (QuakeVB) C's static tables and D_PolysetDraw's span buffer
Private Sub D_PolysetInit()
    Dim tm As Long, tn As Long, q As Long

    SetEdgeTable 0, 0, 1, RP0, RP2, -1, 2, RP0, RP1, RP2
    SetEdgeTable 1, 0, 2, RP1, RP0, RP2, 1, RP1, RP2, -1
    SetEdgeTable 2, 1, 1, RP0, RP2, -1, 1, RP1, RP2, -1
    SetEdgeTable 3, 0, 1, RP1, RP0, -1, 2, RP1, RP2, RP0
    SetEdgeTable 4, 0, 2, RP0, RP2, RP1, 1, RP0, RP1, -1
    SetEdgeTable 5, 0, 1, RP2, RP1, -1, 1, RP2, RP0, -1
    SetEdgeTable 6, 0, 1, RP2, RP1, -1, 2, RP2, RP0, RP1
    SetEdgeTable 7, 0, 2, RP2, RP1, RP0, 1, RP2, RP0, -1
    SetEdgeTable 8, 0, 1, RP1, RP0, -1, 1, RP1, RP2, -1
    SetEdgeTable 9, 1, 1, RP2, RP1, -1, 1, RP0, RP1, -1
    SetEdgeTable 10, 1, 1, RP1, RP0, -1, 1, RP2, RP0, -1
    SetEdgeTable 11, 0, 1, RP0, RP2, -1, 1, RP0, RP1, -1

    ' table of quotients and remainders for [-15...16] / [-15...16]
    ' (adivtab.h: floor division, the remainder takes the sign of the
    ' divisor, and 0 / 0 for a divisor of 0)
    For tm = -15 To 16
        For tn = -15 To 16
            If tn = 0 Then
                q = 0
            Else
                q = Int(tm / tn)
            End If
            adivtab_quotient((tm + 15) * 32 + (tn + 15)) = q
            adivtab_remainder((tm + 15) * 32 + (tn + 15)) = tm - q * tn
        Next
    Next

    ' one extra because of cache line pretouching
    ReDim a_spans(0 To DPS_MAXSPANS)
    ReDim rtv(0 To 6 * 256 - 1)

    polyse_made = True
End Sub

Public Sub D_PolysetDraw()
    If Not polyse_made Then D_PolysetInit

    If r_affinetridesc.drawtype Then
        D_DrawSubdiv
    Else
        D_DrawNonSubdiv
    End If
End Sub

' fv: an a_finalverts() index
Public Sub D_PolysetDrawFinalVerts(ByVal fv As Long, ByVal numverts As Long)
    Dim i As Long, z As Long, zbuf As Long, pix As Long

    For i = 0 To numverts - 1
        ' valid triangle coordinates for filling can include the bottom and
        ' right clip edges, due to the fill rule; these shouldn't be drawn
        If a_finalverts(fv).v(0) < r_refdef.vrectright And a_finalverts(fv).v(1) < r_refdef.vrectbottom Then
            z = (a_finalverts(fv).v(5) And &HFFFF0000) \ &H10000
            zbuf = zspantable(a_finalverts(fv).v(1)) + a_finalverts(fv).v(0)
            If z >= d_pzbuffer(zbuf) Then
                d_pzbuffer(zbuf) = z
                pix = alias_skins(skintable((a_finalverts(fv).v(3) And &HFFFF0000) \ &H10000) _
                    + (a_finalverts(fv).v(2) And &HFFFF0000) \ &H10000)
                pix = acolormap(pix + (a_finalverts(fv).v(4) And &HFF00&))
                d_viewbuffer(d_scantable(a_finalverts(fv).v(1)) + a_finalverts(fv).v(0)) = pix
            End If
        End If
        fv = fv + 1
    Next
End Sub

' (QuakeVB) the triangle's facesfront and vertex indices
Private Sub GetTriangle(ByVal ptri As Long, ffront As Long, i0 As Long, i1 As Long, i2 As Long)
    If ptri >= 0 Then
        ffront = alias_tris(ptri).facesfront
        i0 = alias_tris(ptri).vertindex(0)
        i1 = alias_tris(ptri).vertindex(1)
        i2 = alias_tris(ptri).vertindex(2)
    Else
        ffront = aclip_tri.facesfront
        i0 = aclip_tri.vertindex(0)
        i1 = aclip_tri.vertindex(1)
        i2 = aclip_tri.vertindex(2)
    End If
End Sub

' (QuakeVB) copies a finalvert's v[] to rtv() (or r_p()) at dst
Private Sub CopyVert(dst() As Long, ByVal d As Long, ByVal fv As Long)
    Dim k As Long

    For k = 0 To 5
        dst(d + k) = a_finalverts(fv).v(k)
    Next
End Sub

Private Sub D_DrawSubdiv()
    Dim ptri As Long, pfv As Long, index0 As Long, index1 As Long, index2 As Long
    Dim i As Long, lnumtriangles As Long, ffront As Long, i0 As Long, i1 As Long, i2 As Long

    pfv = r_affinetridesc.pfinalverts
    ptri = r_affinetridesc.ptriangles
    lnumtriangles = r_affinetridesc.numtriangles

    For i = 0 To lnumtriangles - 1
        If ptri >= 0 Then
            GetTriangle ptri + i, ffront, i0, i1, i2
        Else
            GetTriangle -1, ffront, i0, i1, i2
        End If
        index0 = pfv + i0
        index1 = pfv + i1
        index2 = pfv + i2

        If ((a_finalverts(index0).v(1) - a_finalverts(index1).v(1)) * _
            (a_finalverts(index0).v(0) - a_finalverts(index2).v(0)) - _
            (a_finalverts(index0).v(0) - a_finalverts(index1).v(0)) * _
            (a_finalverts(index0).v(1) - a_finalverts(index2).v(1))) >= 0 Then
            GoTo nexti
        End If

        d_pcolormap = a_finalverts(index0).v(4) And &HFF00&

        ' (C changes the vertices' s in place for the back side and restores
        ' it after; here the copies in rtv() get the seam fixup)
        CopyVert rtv, 0, index0
        CopyVert rtv, 6, index1
        CopyVert rtv, 12, index2
        rt_top = 18

        If ffront = 0 Then
            If a_finalverts(index0).flags And ALIAS_ONSEAM Then rtv(2) = rtv(2) + r_affinetridesc.seamfixupX16
            If a_finalverts(index1).flags And ALIAS_ONSEAM Then rtv(8) = rtv(8) + r_affinetridesc.seamfixupX16
            If a_finalverts(index2).flags And ALIAS_ONSEAM Then rtv(14) = rtv(14) + r_affinetridesc.seamfixupX16
        End If

        D_PolysetRecursiveTriangle 0, 6, 12
nexti:
    Next
End Sub

Private Sub D_DrawNonSubdiv()
    Dim ptri As Long, pfv As Long, index0 As Long, index1 As Long, index2 As Long
    Dim i As Long, lnumtriangles As Long, ffront As Long, i0 As Long, i1 As Long, i2 As Long

    pfv = r_affinetridesc.pfinalverts
    ptri = r_affinetridesc.ptriangles
    lnumtriangles = r_affinetridesc.numtriangles

    For i = 0 To lnumtriangles - 1
        GetTriangle ptri, ffront, i0, i1, i2
        index0 = pfv + i0
        index1 = pfv + i1
        index2 = pfv + i2

        d_xdenom = (a_finalverts(index0).v(1) - a_finalverts(index1).v(1)) * _
            (a_finalverts(index0).v(0) - a_finalverts(index2).v(0)) - _
            (a_finalverts(index0).v(0) - a_finalverts(index1).v(0)) * (a_finalverts(index0).v(1) - a_finalverts(index2).v(1))

        If d_xdenom >= 0 Then GoTo nexti

        CopyVert r_p, RP0, index0        ' u, v, s, t, light, iz
        CopyVert r_p, RP1, index1
        CopyVert r_p, RP2, index2

        If ffront = 0 Then
            If a_finalverts(index0).flags And ALIAS_ONSEAM Then r_p(RP0 + 2) = r_p(RP0 + 2) + r_affinetridesc.seamfixupX16
            If a_finalverts(index1).flags And ALIAS_ONSEAM Then r_p(RP1 + 2) = r_p(RP1 + 2) + r_affinetridesc.seamfixupX16
            If a_finalverts(index2).flags And ALIAS_ONSEAM Then r_p(RP2 + 2) = r_p(RP2 + 2) + r_affinetridesc.seamfixupX16
        End If

        D_PolysetSetEdgeTable
        D_RasterizeAliasPolySmooth
nexti:
        If ptri >= 0 Then ptri = ptri + 1
    Next
End Sub

' lp1, lp2, lp3: rtv() offsets of the corners
Private Sub D_PolysetRecursiveTriangle(ByVal lp1 As Long, ByVal lp2 As Long, ByVal lp3 As Long)
    Dim temp As Long, d As Long, new_ As Long, z As Long, zbuf As Long, pix As Long

    d = rtv(lp2) - rtv(lp1)
    If d < -1 Or d > 1 Then GoTo split
    d = rtv(lp2 + 1) - rtv(lp1 + 1)
    If d < -1 Or d > 1 Then GoTo split

    d = rtv(lp3) - rtv(lp2)
    If d < -1 Or d > 1 Then GoTo split2
    d = rtv(lp3 + 1) - rtv(lp2 + 1)
    If d < -1 Or d > 1 Then GoTo split2

    d = rtv(lp1) - rtv(lp3)
    If d < -1 Or d > 1 Then GoTo split3
    d = rtv(lp1 + 1) - rtv(lp3 + 1)
    If d < -1 Or d > 1 Then GoTo split3

    Exit Sub                             ' entire tri is filled

split3:
    temp = lp1
    lp1 = lp3
    lp3 = lp2
    lp2 = temp

    GoTo split

split2:
    temp = lp1
    lp1 = lp2
    lp2 = lp3
    lp3 = temp

split:
    ' (QuakeVB) C's local new[6]
    new_ = rt_top
    rt_top = rt_top + 6
    If rt_top > UBound(rtv) + 1 Then ReDim Preserve rtv(0 To 2 * (UBound(rtv) + 1) - 1)

    ' split this edge
    rtv(new_) = SAR(rtv(lp1) + rtv(lp2), 1)
    rtv(new_ + 1) = SAR(rtv(lp1 + 1) + rtv(lp2 + 1), 1)
    rtv(new_ + 2) = SAR(rtv(lp1 + 2) + rtv(lp2 + 2), 1)
    rtv(new_ + 3) = SAR(rtv(lp1 + 3) + rtv(lp2 + 3), 1)
    rtv(new_ + 5) = SAR(rtv(lp1 + 5) + rtv(lp2 + 5), 1)

    ' draw the point if splitting a leading edge
    If rtv(lp2 + 1) > rtv(lp1 + 1) Then GoTo nodraw
    If rtv(lp2 + 1) = rtv(lp1 + 1) And rtv(lp2) < rtv(lp1) Then GoTo nodraw

    z = (rtv(new_ + 5) And &HFFFF0000) \ &H10000
    zbuf = zspantable(rtv(new_ + 1)) + rtv(new_)
    If z >= d_pzbuffer(zbuf) Then
        d_pzbuffer(zbuf) = z
        pix = acolormap(d_pcolormap + alias_skins(skintable((rtv(new_ + 3) And &HFFFF0000) \ &H10000) _
            + (rtv(new_ + 2) And &HFFFF0000) \ &H10000))
        d_viewbuffer(d_scantable(rtv(new_ + 1)) + rtv(new_)) = pix
    End If

nodraw:
    ' recursively continue
    D_PolysetRecursiveTriangle lp3, lp1, new_
    D_PolysetRecursiveTriangle lp3, new_, lp2

    rt_top = rt_top - 6
End Sub

Public Sub D_PolysetUpdateTables()
    Dim i As Long, s As Long

    If r_affinetridesc.skinwidth <> skinwidth Or r_affinetridesc.pskin <> skinstart Then
        skinwidth = r_affinetridesc.skinwidth
        skinstart = r_affinetridesc.pskin
        s = skinstart
        For i = 0 To MAX_LBM_HEIGHT - 1
            skintable(i) = s
            s = s + skinwidth
        Next
    End If
End Sub

Private Sub D_PolysetScanLeftEdge(ByVal height As Long)
    Do
        a_spans(d_pedgespanpackage).pdest = d_pdest
        a_spans(d_pedgespanpackage).pz = d_pz
        a_spans(d_pedgespanpackage).count = d_aspancount
        a_spans(d_pedgespanpackage).ptex = d_ptex

        a_spans(d_pedgespanpackage).sfrac = d_sfrac
        a_spans(d_pedgespanpackage).tfrac = d_tfrac

        ' FIXME: need to clamp l, s, t, at both ends?
        a_spans(d_pedgespanpackage).light = d_light
        a_spans(d_pedgespanpackage).zi = d_zi

        d_pedgespanpackage = d_pedgespanpackage + 1

        errorterm = errorterm + erroradjustup
        If errorterm >= 0 Then
            d_pdest = d_pdest + d_pdestextrastep
            d_pz = d_pz + d_pzextrastep
            d_aspancount = d_aspancount + d_countextrastep
            d_ptex = d_ptex + d_ptexextrastep
            d_sfrac = d_sfrac + d_sfracextrastep
            d_ptex = d_ptex + d_sfrac \ &H10000   ' (d_sfrac >= 0)

            d_sfrac = d_sfrac And &HFFFF&
            d_tfrac = d_tfrac + d_tfracextrastep
            If d_tfrac And &H10000 Then
                d_ptex = d_ptex + r_affinetridesc.skinwidth
                d_tfrac = d_tfrac And &HFFFF&
            End If
            d_light = d_light + d_lightextrastep
            d_zi = d_zi + d_ziextrastep
            errorterm = errorterm - erroradjustdown
        Else
            d_pdest = d_pdest + d_pdestbasestep
            d_pz = d_pz + d_pzbasestep
            d_aspancount = d_aspancount + ubasestep
            d_ptex = d_ptex + d_ptexbasestep
            d_sfrac = d_sfrac + d_sfracbasestep
            d_ptex = d_ptex + d_sfrac \ &H10000
            d_sfrac = d_sfrac And &HFFFF&
            d_tfrac = d_tfrac + d_tfracbasestep
            If d_tfrac And &H10000 Then
                d_ptex = d_ptex + r_affinetridesc.skinwidth
                d_tfrac = d_tfrac And &HFFFF&
            End If
            d_light = d_light + d_lightbasestep
            d_zi = d_zi + d_zibasestep
        End If

        height = height - 1
    Loop While height <> 0
End Sub

Private Sub D_PolysetSetUpForLineScan(ByVal startvertu As Long, ByVal startvertv As Long, _
        ByVal endvertu As Long, ByVal endvertv As Long)
    Dim tm As Long, tn As Long, ptemp As Long

    ' TODO: implement x86 version

    errorterm = -1

    tm = endvertu - startvertu
    tn = endvertv - startvertv

    If (tm <= 16 And tm >= -15) And (tn <= 16 And tn >= -15) Then
        ptemp = (tm + 15) * 32 + (tn + 15)
        ubasestep = adivtab_quotient(ptemp)
        erroradjustup = adivtab_remainder(ptemp)
        erroradjustdown = tn
    Else
        FloorDivMod tm, tn, ubasestep, erroradjustup

        erroradjustdown = tn
    End If
End Sub

' (QuakeVB) C: ceil
Private Function Ceil_(ByVal x As Double) As Double
    Ceil_ = -Int(-x)
End Function

Private Sub D_PolysetCalcGradients(ByVal skinwidth_ As Long)
    Dim xstepdenominv As Single, ystepdenominv As Single, t0 As Single, t1 As Single
    Dim p01_minus_p21 As Single, p11_minus_p21 As Single, p00_minus_p20 As Single, p10_minus_p20 As Single

    p00_minus_p20 = r_p(RP0) - r_p(RP2)
    p01_minus_p21 = r_p(RP0 + 1) - r_p(RP2 + 1)
    p10_minus_p20 = r_p(RP1) - r_p(RP2)
    p11_minus_p21 = r_p(RP1 + 1) - r_p(RP2 + 1)

    xstepdenominv = 1! / CSng(d_xdenom)

    ystepdenominv = -xstepdenominv

    ' ceil () for light so positive steps are exaggerated, negative steps
    ' diminished,  pushing us away from underflow toward overflow. Underflow is
    ' very visible, overflow is very unlikely, because of ambient lighting
    t0 = r_p(RP0 + 4) - r_p(RP2 + 4)
    t1 = r_p(RP1 + 4) - r_p(RP2 + 4)
    r_lstepx = ToInt(Ceil_((t1 * p01_minus_p21 - t0 * p11_minus_p21) * xstepdenominv))
    r_lstepy = ToInt(Ceil_((t1 * p00_minus_p20 - t0 * p10_minus_p20) * ystepdenominv))

    t0 = r_p(RP0 + 2) - r_p(RP2 + 2)
    t1 = r_p(RP1 + 2) - r_p(RP2 + 2)
    r_sstepx = ToInt((t1 * p01_minus_p21 - t0 * p11_minus_p21) * xstepdenominv)
    r_sstepy = ToInt((t1 * p00_minus_p20 - t0 * p10_minus_p20) * ystepdenominv)

    t0 = r_p(RP0 + 3) - r_p(RP2 + 3)
    t1 = r_p(RP1 + 3) - r_p(RP2 + 3)
    r_tstepx = ToInt((t1 * p01_minus_p21 - t0 * p11_minus_p21) * xstepdenominv)
    r_tstepy = ToInt((t1 * p00_minus_p20 - t0 * p10_minus_p20) * ystepdenominv)

    t0 = r_p(RP0 + 5) - r_p(RP2 + 5)
    t1 = r_p(RP1 + 5) - r_p(RP2 + 5)
    r_zistepx = ToInt((t1 * p01_minus_p21 - t0 * p11_minus_p21) * xstepdenominv)
    r_zistepy = ToInt((t1 * p00_minus_p20 - t0 * p10_minus_p20) * ystepdenominv)

    a_sstepxfrac = r_sstepx And &HFFFF&
    a_tstepxfrac = r_tstepx And &HFFFF&

    a_ststepxwhole = skinwidth_ * SAR(r_tstepx, 16) + SAR(r_sstepx, 16)
End Sub

' pspanpackage: an a_spans() index
Private Sub D_PolysetDrawSpans8(ByVal pspanpackage As Long)
    Dim lcount As Long, lpdest As Long, lptex As Long, lsfrac As Long, ltfrac As Long
    Dim llight As Long, lzi As Long, lpz As Long, z As Long, sw As Long

    sw = r_affinetridesc.skinwidth

    Do
        lcount = d_aspancount - a_spans(pspanpackage).count

        errorterm = errorterm + erroradjustup
        If errorterm >= 0 Then
            d_aspancount = d_aspancount + d_countextrastep
            errorterm = errorterm - erroradjustdown
        Else
            d_aspancount = d_aspancount + ubasestep
        End If

        If lcount Then
            lpdest = a_spans(pspanpackage).pdest
            lptex = a_spans(pspanpackage).ptex
            lpz = a_spans(pspanpackage).pz
            lsfrac = a_spans(pspanpackage).sfrac
            ltfrac = a_spans(pspanpackage).tfrac
            llight = a_spans(pspanpackage).light
            lzi = a_spans(pspanpackage).zi

            Do
                z = (lzi And &HFFFF0000) \ &H10000
                If z >= d_pzbuffer(lpz) Then
                    d_viewbuffer(lpdest) = acolormap(alias_skins(lptex) + (llight And &HFF00&))
                    ' gel mapping                 *lpdest = gelmap[*lpdest];
                    d_pzbuffer(lpz) = z
                End If
                lpdest = lpdest + 1
                lzi = lzi + r_zistepx
                lpz = lpz + 1
                llight = llight + r_lstepx
                lptex = lptex + a_ststepxwhole
                lsfrac = lsfrac + a_sstepxfrac
                lptex = lptex + lsfrac \ &H10000   ' (lsfrac >= 0)
                lsfrac = lsfrac And &HFFFF&
                ltfrac = ltfrac + a_tstepxfrac
                If ltfrac And &H10000 Then
                    lptex = lptex + sw
                    ltfrac = ltfrac And &HFFFF&
                End If
                lcount = lcount - 1
            Loop While lcount <> 0
        End If

        pspanpackage = pspanpackage + 1
    Loop While a_spans(pspanpackage).count <> -999999
End Sub

' (QuakeVB) the left edge setup that C repeats for the top and bottom parts
Private Sub SetUpLeftEdgeSteps()
    Dim working_lstepx As Long

    d_pdestbasestep = screenwidth + ubasestep
    d_pdestextrastep = d_pdestbasestep + 1

    d_pzbasestep = d_zwidth + ubasestep
    d_pzextrastep = d_pzbasestep + 1

    ' for negative steps in x along left edge, bias toward overflow rather than
    ' underflow (sort of turning the floor () we did in the gradient calcs into
    ' ceil (), but plus a little bit)
    If ubasestep < 0 Then
        working_lstepx = r_lstepx - 1
    Else
        working_lstepx = r_lstepx
    End If

    d_countextrastep = ubasestep + 1
    d_ptexbasestep = SAR(r_sstepy + r_sstepx * ubasestep, 16) + _
        SAR(r_tstepy + r_tstepx * ubasestep, 16) * r_affinetridesc.skinwidth
    d_sfracbasestep = (r_sstepy + r_sstepx * ubasestep) And &HFFFF&
    d_tfracbasestep = (r_tstepy + r_tstepx * ubasestep) And &HFFFF&
    d_lightbasestep = r_lstepy + working_lstepx * ubasestep
    d_zibasestep = r_zistepy + r_zistepx * ubasestep

    d_ptexextrastep = SAR(r_sstepy + r_sstepx * d_countextrastep, 16) + _
        SAR(r_tstepy + r_tstepx * d_countextrastep, 16) * r_affinetridesc.skinwidth
    d_sfracextrastep = (r_sstepy + r_sstepx * d_countextrastep) And &HFFFF&
    d_tfracextrastep = (r_tstepy + r_tstepx * d_countextrastep) And &HFFFF&
    d_lightextrastep = d_lightbasestep + working_lstepx
    d_ziextrastep = d_zibasestep + r_zistepx
End Sub

' (QuakeVB) stores the current left edge state in the next span package
Private Sub EmitOneSpanPackage()
    a_spans(d_pedgespanpackage).pdest = d_pdest
    a_spans(d_pedgespanpackage).pz = d_pz
    a_spans(d_pedgespanpackage).count = d_aspancount
    a_spans(d_pedgespanpackage).ptex = d_ptex

    a_spans(d_pedgespanpackage).sfrac = d_sfrac
    a_spans(d_pedgespanpackage).tfrac = d_tfrac

    ' FIXME: need to clamp l, s, t, at both ends?
    a_spans(d_pedgespanpackage).light = d_light
    a_spans(d_pedgespanpackage).zi = d_zi

    d_pedgespanpackage = d_pedgespanpackage + 1
End Sub

Private Sub D_RasterizeAliasPolySmooth()
    Dim initialleftheight As Long, initialrightheight As Long
    Dim plefttop As Long, prighttop As Long, pleftbottom As Long, prightbottom As Long
    Dim originalcount As Long, height As Long, pstart As Long

    plefttop = edgetables(pedgetable).pleftedgevert0
    prighttop = edgetables(pedgetable).prightedgevert0

    pleftbottom = edgetables(pedgetable).pleftedgevert1
    prightbottom = edgetables(pedgetable).prightedgevert1

    initialleftheight = r_p(pleftbottom + 1) - r_p(plefttop + 1)
    initialrightheight = r_p(prightbottom + 1) - r_p(prighttop + 1)

    '
    ' set the s, t, and light gradients, which are consistent across the triangle
    ' because being a triangle, things are affine
    '
    D_PolysetCalcGradients r_affinetridesc.skinwidth

    '
    ' rasterize the polygon
    '

    '
    ' scan out the top (and possibly only) part of the left edge
    '
    d_pedgespanpackage = 0

    ystart = r_p(plefttop + 1)
    d_aspancount = r_p(plefttop) - r_p(prighttop)

    d_ptex = r_affinetridesc.pskin + SAR(r_p(plefttop + 2), 16) + _
        SAR(r_p(plefttop + 3), 16) * r_affinetridesc.skinwidth
    d_sfrac = r_p(plefttop + 2) And &HFFFF&
    d_tfrac = r_p(plefttop + 3) And &HFFFF&
    d_light = r_p(plefttop + 4)
    d_zi = r_p(plefttop + 5)

    d_pdest = ystart * screenwidth + r_p(plefttop)
    d_pz = ystart * d_zwidth + r_p(plefttop)

    If initialleftheight = 1 Then
        EmitOneSpanPackage
    Else
        D_PolysetSetUpForLineScan r_p(plefttop), r_p(plefttop + 1), r_p(pleftbottom), r_p(pleftbottom + 1)

        ' TODO: can reuse partial expressions here
        SetUpLeftEdgeSteps

        D_PolysetScanLeftEdge initialleftheight
    End If

    '
    ' scan out the bottom part of the left edge, if it exists
    '
    If edgetables(pedgetable).numleftedges = 2 Then
        plefttop = pleftbottom
        pleftbottom = edgetables(pedgetable).pleftedgevert2

        height = r_p(pleftbottom + 1) - r_p(plefttop + 1)

        ' TODO: make this a function; modularize this function in general

        ystart = r_p(plefttop + 1)
        d_aspancount = r_p(plefttop) - r_p(prighttop)
        d_ptex = r_affinetridesc.pskin + SAR(r_p(plefttop + 2), 16) + _
            SAR(r_p(plefttop + 3), 16) * r_affinetridesc.skinwidth
        d_sfrac = 0
        d_tfrac = 0
        d_light = r_p(plefttop + 4)
        d_zi = r_p(plefttop + 5)

        d_pdest = ystart * screenwidth + r_p(plefttop)
        d_pz = ystart * d_zwidth + r_p(plefttop)

        If height = 1 Then
            EmitOneSpanPackage
        Else
            D_PolysetSetUpForLineScan r_p(plefttop), r_p(plefttop + 1), r_p(pleftbottom), r_p(pleftbottom + 1)

            SetUpLeftEdgeSteps

            D_PolysetScanLeftEdge height
        End If
    End If

    ' scan out the top (and possibly only) part of the right edge, updating the
    ' count field
    d_pedgespanpackage = 0

    D_PolysetSetUpForLineScan r_p(prighttop), r_p(prighttop + 1), r_p(prightbottom), r_p(prightbottom + 1)
    d_aspancount = 0
    d_countextrastep = ubasestep + 1
    originalcount = a_spans(initialrightheight).count
    a_spans(initialrightheight).count = -999999   ' mark end of the spanpackages
    D_PolysetDrawSpans8 0

    ' scan out the bottom part of the right edge, if it exists
    If edgetables(pedgetable).numrightedges = 2 Then
        pstart = initialrightheight
        a_spans(pstart).count = originalcount

        d_aspancount = r_p(prightbottom) - r_p(prighttop)

        prighttop = prightbottom
        prightbottom = edgetables(pedgetable).prightedgevert2

        height = r_p(prightbottom + 1) - r_p(prighttop + 1)

        D_PolysetSetUpForLineScan r_p(prighttop), r_p(prighttop + 1), r_p(prightbottom), r_p(prightbottom + 1)

        d_countextrastep = ubasestep + 1
        a_spans(initialrightheight + height).count = -999999   ' mark end of the spanpackages
        D_PolysetDrawSpans8 pstart
    End If
End Sub

Private Sub D_PolysetSetEdgeTable()
    Dim edgetableindex As Long

    edgetableindex = 0                   ' assume the vertices are already in
                                         '  top to bottom order

    '
    ' determine which edges are right & left, and the order in which
    ' to rasterize them
    '
    If r_p(RP0 + 1) >= r_p(RP1 + 1) Then
        If r_p(RP0 + 1) = r_p(RP1 + 1) Then
            If r_p(RP0 + 1) < r_p(RP2 + 1) Then
                pedgetable = 2
            Else
                pedgetable = 5
            End If

            Exit Sub
        Else
            edgetableindex = 1
        End If
    End If

    If r_p(RP0 + 1) = r_p(RP2 + 1) Then
        If edgetableindex Then
            pedgetable = 8
        Else
            pedgetable = 9
        End If

        Exit Sub
    ElseIf r_p(RP1 + 1) = r_p(RP2 + 1) Then
        If edgetableindex Then
            pedgetable = 10
        Else
            pedgetable = 11
        End If

        Exit Sub
    End If

    If r_p(RP0 + 1) > r_p(RP2 + 1) Then edgetableindex = edgetableindex + 2

    If r_p(RP1 + 1) > r_p(RP2 + 1) Then edgetableindex = edgetableindex + 4

    pedgetable = edgetableindex
End Sub
