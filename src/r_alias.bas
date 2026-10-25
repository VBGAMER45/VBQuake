Attribute VB_Name = "r_alias"
' r_alias.bas - routines for setting up to draw alias models (r_alias.c)
'
' The finalvert_t and auxvert_t pointers are indices: a_finalverts() holds
' the model's vertices from 0 (C: R_AliasDrawModel's local finalverts) and
' r_aclip.bas's two clip buffers (C: fv[2][8]) from FV_CLIP; a_auxverts()
' is C's auxverts. paliashdr and pmdl are aliashdrs() indices (C's pmdl is
' the header's .mdl), r_apverts an alias_verts() index. acolormap() is a
' Byte array laid over the entity's colormap: vid_colormap() or a player's
' translations (SetView, common.bas).
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const LIGHT_MIN As Long = 5      ' lowest light value we'll allow, to avoid the
                                         '  need for inner-loop light clamping

Public Const FV_CLIP As Long = MAXALIASVERTS   ' r_aclip.bas's fv(0)(0); fv(1)(0) is FV_CLIP + 8

Public r_affinetridesc As affinetridesc_t

Public acolormap() As Byte               ' FIXME: should go away
Private sa_acolormap As SAFEARRAY1D

Public r_apverts As Long                 ' alias_verts() index

' TODO: these probably will go away with optimized rasterization
Private pmdl As Long                     ' aliashdrs() index
Private r_plightvec As vec3_t
Private r_ambientlight As Long
Private r_shadelight As Single
Private paliashdr As Long                ' aliashdrs() index
Public a_finalverts() As finalvert_t     ' C: pfinalverts (always element 0 here)
Public a_auxverts() As auxvert_t         ' C: pauxverts
Private aliasverts_made As Boolean
Private ziscale As Single
Private pmodel As Long                   ' mod_known() index

Private alias_forward As vec3_t, alias_right As vec3_t, alias_up As vec3_t

Public r_amodels_drawn As Long
Private a_skinwidth As Long
Private r_anumverts As Long

Public aliastransform(0 To 2, 0 To 3) As Single

Private aedges(0 To 11, 0 To 1) As Long  ' C: aedge_t aedges[12] (index0, index1)
Private aedges_made As Boolean

Private tmatrix(0 To 2, 0 To 3) As Single     ' C: statics of R_AliasSetUpTransform
Private viewmatrix(0 To 2, 0 To 3) As Single

Private Sub InitAedges()
    Dim a As Variant, i As Long

    a = Array(0, 1, 1, 2, 2, 3, 3, 0, 4, 5, 5, 6, 6, 7, 7, 4, 0, 5, 1, 4, 2, 7, 3, 6)
    For i = 0 To 11
        aedges(i, 0) = a(i * 2)
        aedges(i, 1) = a(i * 2 + 1)
    Next
    aedges_made = True
End Sub

Public Function R_AliasCheckBBox() As Boolean
    Dim i As Long, flags As Long, frame As Long, numv As Long, pahdr As Long
    Dim zi As Single, basepts(0 To 7) As vec3_t, v0 As Single, v1 As Single, frac As Single
    Dim viewflags(0 To 15) As Long       ' C: finalvert_t viewpts[16] (only the flags are used)
    Dim viewaux(0 To 15) As vec3_t       ' C: auxvert_t viewaux[16]
    Dim bmin As trivertx_t, bmax As trivertx_t
    Dim zclipped As Boolean, zfullyclipped As Boolean
    Dim anyclip As Long, allclip As Long, minz As Long, e0 As Long, e1 As Long

    If Not aedges_made Then InitAedges

    ' expand, rotate, and translate points into worldspace

    cl_entities(currententity).trivial_accept = 0
    pmodel = cl_entities(currententity).model
    pahdr = Mod_Extradata(pmodel)
    pmdl = pahdr

    R_AliasSetUpTransform 0

    ' construct the base bounding box for this frame
    frame = cl_entities(currententity).frame
    ' TODO: don't repeat this check when drawing?
    If frame >= aliashdrs(pmdl).mdl.numframes Or frame < 0 Then
        Con_DPrintf "No such frame %d %s\n", frame, mod_known(pmodel).name
        frame = 0
    End If

    bmin = aliashdrs(pahdr).frames(frame).bboxmin
    bmax = aliashdrs(pahdr).frames(frame).bboxmax

    ' x worldspace coordinates
    basepts(0).v(0) = bmin.v(0): basepts(1).v(0) = bmin.v(0)
    basepts(2).v(0) = bmin.v(0): basepts(3).v(0) = bmin.v(0)
    basepts(4).v(0) = bmax.v(0): basepts(5).v(0) = bmax.v(0)
    basepts(6).v(0) = bmax.v(0): basepts(7).v(0) = bmax.v(0)

    ' y worldspace coordinates
    basepts(0).v(1) = bmin.v(1): basepts(3).v(1) = bmin.v(1)
    basepts(5).v(1) = bmin.v(1): basepts(6).v(1) = bmin.v(1)
    basepts(1).v(1) = bmax.v(1): basepts(2).v(1) = bmax.v(1)
    basepts(4).v(1) = bmax.v(1): basepts(7).v(1) = bmax.v(1)

    ' z worldspace coordinates
    basepts(0).v(2) = bmin.v(2): basepts(1).v(2) = bmin.v(2)
    basepts(4).v(2) = bmin.v(2): basepts(5).v(2) = bmin.v(2)
    basepts(2).v(2) = bmax.v(2): basepts(3).v(2) = bmax.v(2)
    basepts(6).v(2) = bmax.v(2): basepts(7).v(2) = bmax.v(2)

    zclipped = False
    zfullyclipped = True

    minz = 9999
    For i = 0 To 7
        R_AliasTransformVector basepts(i), viewaux(i)

        If viewaux(i).v(2) < ALIAS_Z_CLIP_PLANE Then
            ' we must clip points that are closer than the near clip plane
            viewflags(i) = ALIAS_Z_CLIP
            zclipped = True
        Else
            If viewaux(i).v(2) < minz Then minz = ToInt(viewaux(i).v(2))
            viewflags(i) = 0
            zfullyclipped = False
        End If
    Next

    If zfullyclipped Then
        R_AliasCheckBBox = False         ' everything was near-z-clipped
        Exit Function
    End If

    numv = 8

    If zclipped Then
        ' organize points by edges, use edges to get new points (possible trivial
        ' reject)
        For i = 0 To 11
            ' edge endpoints
            e0 = aedges(i, 0)
            e1 = aedges(i, 1)

            ' if one end is clipped and the other isn't, make a new point
            If viewflags(e0) Xor viewflags(e1) Then
                frac = (ALIAS_Z_CLIP_PLANE - viewaux(e0).v(2)) / (viewaux(e1).v(2) - viewaux(e0).v(2))
                viewaux(numv).v(0) = viewaux(e0).v(0) + (viewaux(e1).v(0) - viewaux(e0).v(0)) * frac
                viewaux(numv).v(1) = viewaux(e0).v(1) + (viewaux(e1).v(1) - viewaux(e0).v(1)) * frac
                viewaux(numv).v(2) = ALIAS_Z_CLIP_PLANE
                viewflags(numv) = 0
                numv = numv + 1
            End If
        Next
    End If

    ' project the vertices that remain after clipping
    anyclip = 0
    allclip = ALIAS_XY_CLIP_MASK

    ' TODO: probably should do this loop in ASM, especially if we use floats
    For i = 0 To numv - 1
        ' we don't need to bother with vertices that were z-clipped
        If viewflags(i) And ALIAS_Z_CLIP Then GoTo nexti

        zi = 1! / viewaux(i).v(2)

        ' FIXME: do with chop mode in ASM, or convert to float
        v0 = (viewaux(i).v(0) * xscale * zi) + xcenter
        v1 = (viewaux(i).v(1) * yscale * zi) + ycenter

        flags = 0

        If v0 < r_refdef.fvrectx Then flags = flags Or ALIAS_LEFT_CLIP
        If v1 < r_refdef.fvrecty Then flags = flags Or ALIAS_TOP_CLIP
        If v0 > r_refdef.fvrectright Then flags = flags Or ALIAS_RIGHT_CLIP
        If v1 > r_refdef.fvrectbottom Then flags = flags Or ALIAS_BOTTOM_CLIP

        anyclip = anyclip Or flags
        allclip = allclip And flags
nexti:
    Next

    If allclip Then
        R_AliasCheckBBox = False         ' trivial reject off one side
        Exit Function
    End If

    If anyclip = 0 And Not zclipped Then cl_entities(currententity).trivial_accept = 1

    If cl_entities(currententity).trivial_accept Then
        If minz > (r_aliastransition + (aliashdrs(pmdl).mdl.size * r_resfudge)) Then
            cl_entities(currententity).trivial_accept = cl_entities(currententity).trivial_accept Or 2
        End If
    End If

    R_AliasCheckBBox = True
End Function

Private Sub R_AliasTransformVector(in_ As vec3_t, out As vec3_t)
    out.v(0) = in_.v(0) * aliastransform(0, 0) + in_.v(1) * aliastransform(0, 1) + in_.v(2) * aliastransform(0, 2) + aliastransform(0, 3)
    out.v(1) = in_.v(0) * aliastransform(1, 0) + in_.v(1) * aliastransform(1, 1) + in_.v(2) * aliastransform(1, 2) + aliastransform(1, 3)
    out.v(2) = in_.v(0) * aliastransform(2, 0) + in_.v(1) * aliastransform(2, 1) + in_.v(2) * aliastransform(2, 2) + aliastransform(2, 3)
End Sub

' General clipped case
Private Sub R_AliasPreparePoints()
    Dim i As Long, pstverts As Long, ptri As Long, f0 As Long, f1 As Long, f2 As Long

    pstverts = aliashdrs(paliashdr).stverts
    r_anumverts = aliashdrs(pmdl).mdl.numverts

    For i = 0 To r_anumverts - 1
        R_AliasTransformFinalVert a_finalverts(i), a_auxverts(i), r_apverts, pstverts
        If a_auxverts(i).fv(2) < ALIAS_Z_CLIP_PLANE Then
            a_finalverts(i).flags = a_finalverts(i).flags Or ALIAS_Z_CLIP
        Else
            R_AliasProjectFinalVert a_finalverts(i), a_auxverts(i)

            If a_finalverts(i).v(0) < r_refdef.aliasvrect.x Then a_finalverts(i).flags = a_finalverts(i).flags Or ALIAS_LEFT_CLIP
            If a_finalverts(i).v(1) < r_refdef.aliasvrect.y Then a_finalverts(i).flags = a_finalverts(i).flags Or ALIAS_TOP_CLIP
            If a_finalverts(i).v(0) > r_refdef.aliasvrectright Then a_finalverts(i).flags = a_finalverts(i).flags Or ALIAS_RIGHT_CLIP
            If a_finalverts(i).v(1) > r_refdef.aliasvrectbottom Then a_finalverts(i).flags = a_finalverts(i).flags Or ALIAS_BOTTOM_CLIP
        End If
        r_apverts = r_apverts + 1
        pstverts = pstverts + 1
    Next

    '
    ' clip and draw all triangles
    '
    r_affinetridesc.numtriangles = 1

    ptri = aliashdrs(paliashdr).triangles
    For i = 0 To aliashdrs(pmdl).mdl.numtris - 1
        f0 = a_finalverts(alias_tris(ptri).vertindex(0)).flags
        f1 = a_finalverts(alias_tris(ptri).vertindex(1)).flags
        f2 = a_finalverts(alias_tris(ptri).vertindex(2)).flags

        If (f0 And f1 And f2 And (ALIAS_XY_CLIP_MASK Or ALIAS_Z_CLIP)) = 0 Then   ' else completely clipped
            If ((f0 Or f1 Or f2) And (ALIAS_XY_CLIP_MASK Or ALIAS_Z_CLIP)) = 0 Then
                ' totally unclipped
                r_affinetridesc.pfinalverts = 0
                r_affinetridesc.ptriangles = ptri
                D_PolysetDraw
            Else
                ' partially clipped
                R_AliasClipTriangle ptri
            End If
        End If
        ptri = ptri + 1
    Next
End Sub

Private Sub R_AliasSetUpTransform(ByVal trivial_accept As Long)
    Dim i As Long, rotationmatrix(0 To 2, 0 To 3) As Single, t2matrix(0 To 2, 0 To 3) As Single
    Dim angles As vec3_t

    ' TODO: should really be stored with the entity instead of being reconstructed
    ' TODO: should use a look-up table
    ' TODO: could cache lazily, stored in the entity

    angles.v(ROLL) = cl_entities(currententity).angles.v(ROLL)
    angles.v(PITCH) = -cl_entities(currententity).angles.v(PITCH)
    angles.v(YAW) = cl_entities(currententity).angles.v(YAW)
    AngleVectors angles, alias_forward, alias_right, alias_up

    tmatrix(0, 0) = aliashdrs(pmdl).mdl.scale.v(0)
    tmatrix(1, 1) = aliashdrs(pmdl).mdl.scale.v(1)
    tmatrix(2, 2) = aliashdrs(pmdl).mdl.scale.v(2)

    tmatrix(0, 3) = aliashdrs(pmdl).mdl.scale_origin.v(0)
    tmatrix(1, 3) = aliashdrs(pmdl).mdl.scale_origin.v(1)
    tmatrix(2, 3) = aliashdrs(pmdl).mdl.scale_origin.v(2)

    ' TODO: can do this with simple matrix rearrangement

    For i = 0 To 2
        t2matrix(i, 0) = alias_forward.v(i)
        t2matrix(i, 1) = -alias_right.v(i)
        t2matrix(i, 2) = alias_up.v(i)
    Next

    t2matrix(0, 3) = -modelorg.v(0)
    t2matrix(1, 3) = -modelorg.v(1)
    t2matrix(2, 3) = -modelorg.v(2)

    ' FIXME: can do more efficiently than full concatenation
    R_ConcatTransforms t2matrix, tmatrix, rotationmatrix

    ' TODO: should be global, set when vright, etc., set
    For i = 0 To 2
        viewmatrix(0, i) = vright.v(i)
        viewmatrix(1, i) = -vup.v(i)     ' VectorInverse (viewmatrix[1])
        viewmatrix(2, i) = vpn.v(i)
    Next

    '   viewmatrix[0][3] = 0;
    '   viewmatrix[1][3] = 0;
    '   viewmatrix[2][3] = 0;

    R_ConcatTransforms viewmatrix, rotationmatrix, aliastransform

    ' do the scaling up of x and y to screen coordinates as part of the transform
    ' for the unclipped case (it would mess up clipping in the clipped case).
    ' Also scale down z, so 1/z is scaled 31 bits for free, and scale down x and y
    ' correspondingly so the projected x and y come out right
    ' FIXME: make this work for clipped case too?
    If trivial_accept Then
        For i = 0 To 3
            aliastransform(0, i) = aliastransform(0, i) * (aliasxscale * (1# / (32768! * 65536)))
            aliastransform(1, i) = aliastransform(1, i) * (aliasyscale * (1# / (32768! * 65536)))
            aliastransform(2, i) = aliastransform(2, i) * (1# / (32768! * 65536))
        Next
    End If
End Sub

Private Sub R_AliasTransformFinalVert(fv As finalvert_t, av As auxvert_t, ByVal pverts As Long, ByVal pstverts As Long)
    Dim temp As Long, lightcos As Single, n As Long, x As Single, y As Single, z As Single

    x = alias_verts(pverts).v(0)
    y = alias_verts(pverts).v(1)
    z = alias_verts(pverts).v(2)
    av.fv(0) = x * aliastransform(0, 0) + y * aliastransform(0, 1) + z * aliastransform(0, 2) + aliastransform(0, 3)
    av.fv(1) = x * aliastransform(1, 0) + y * aliastransform(1, 1) + z * aliastransform(1, 2) + aliastransform(1, 3)
    av.fv(2) = x * aliastransform(2, 0) + y * aliastransform(2, 1) + z * aliastransform(2, 2) + aliastransform(2, 3)

    fv.v(2) = alias_stverts(pstverts).s
    fv.v(3) = alias_stverts(pstverts).t

    fv.flags = alias_stverts(pstverts).onseam

    ' lighting
    n = alias_verts(pverts).lightnormalindex
    lightcos = r_avertexnormals(n, 0) * r_plightvec.v(0) + r_avertexnormals(n, 1) * r_plightvec.v(1) _
        + r_avertexnormals(n, 2) * r_plightvec.v(2)
    temp = r_ambientlight

    If lightcos < 0 Then
        temp = temp + ToInt(r_shadelight * lightcos)

        ' clamp; because we limited the minimum ambient and shading light, we
        ' don't have to clamp low light, just bright
        If temp < 0 Then temp = 0
    End If

    fv.v(4) = temp
End Sub

Private Sub R_AliasTransformAndProjectFinalVerts(ByVal fv As Long, ByVal pstverts As Long)
    Dim i As Long, temp As Long, lightcos As Single, zi As Single, pverts As Long, n As Long
    Dim x As Single, y As Single, z As Single

    pverts = r_apverts

    For i = 0 To r_anumverts - 1
        x = alias_verts(pverts).v(0)
        y = alias_verts(pverts).v(1)
        z = alias_verts(pverts).v(2)

        ' transform and project
        zi = 1! / (x * aliastransform(2, 0) + y * aliastransform(2, 1) + z * aliastransform(2, 2) + aliastransform(2, 3))

        ' x, y, and z are scaled down by 1/2**31 in the transform, so 1/z is
        ' scaled up by 1/2**31, and the scaling cancels out for x and y in the
        ' projection
        a_finalverts(fv).v(5) = ToInt(zi)

        a_finalverts(fv).v(0) = ToInt(((x * aliastransform(0, 0) + y * aliastransform(0, 1) + z * aliastransform(0, 2) _
            + aliastransform(0, 3)) * zi) + aliasxcenter)
        a_finalverts(fv).v(1) = ToInt(((x * aliastransform(1, 0) + y * aliastransform(1, 1) + z * aliastransform(1, 2) _
            + aliastransform(1, 3)) * zi) + aliasycenter)

        a_finalverts(fv).v(2) = alias_stverts(pstverts).s
        a_finalverts(fv).v(3) = alias_stverts(pstverts).t
        a_finalverts(fv).flags = alias_stverts(pstverts).onseam

        ' lighting
        n = alias_verts(pverts).lightnormalindex
        lightcos = r_avertexnormals(n, 0) * r_plightvec.v(0) + r_avertexnormals(n, 1) * r_plightvec.v(1) _
            + r_avertexnormals(n, 2) * r_plightvec.v(2)
        temp = r_ambientlight

        If lightcos < 0 Then
            temp = temp + ToInt(r_shadelight * lightcos)

            ' clamp; because we limited the minimum ambient and shading light, we
            ' don't have to clamp low light, just bright
            If temp < 0 Then temp = 0
        End If

        a_finalverts(fv).v(4) = temp

        fv = fv + 1
        pverts = pverts + 1
        pstverts = pstverts + 1
    Next
End Sub

Public Sub R_AliasProjectFinalVert(fv As finalvert_t, av As auxvert_t)
    Dim zi As Single

    ' project points
    zi = 1! / av.fv(2)

    fv.v(5) = ToInt(zi * ziscale)

    fv.v(0) = ToInt((av.fv(0) * aliasxscale * zi) + aliasxcenter)
    fv.v(1) = ToInt((av.fv(1) * aliasyscale * zi) + aliasycenter)
End Sub

Private Sub R_AliasPrepareUnclippedPoints()
    Dim pstverts As Long

    pstverts = aliashdrs(paliashdr).stverts
    r_anumverts = aliashdrs(pmdl).mdl.numverts
    ' FIXME: just use pfinalverts directly?

    R_AliasTransformAndProjectFinalVerts 0, pstverts

    If r_affinetridesc.drawtype Then D_PolysetDrawFinalVerts 0, r_anumverts

    r_affinetridesc.pfinalverts = 0
    r_affinetridesc.ptriangles = aliashdrs(paliashdr).triangles
    r_affinetridesc.numtriangles = aliashdrs(pmdl).mdl.numtris

    D_PolysetDraw
End Sub

Private Sub R_AliasSetupSkin()
    Dim skinnum As Long, i As Long, numskins As Long, paliasskingroup As Long, pskinintervals As Long
    Dim fullskininterval As Single, skintargettime As Single, skintime As Single
    Dim pskindesc As Long, skin As Long

    skinnum = cl_entities(currententity).skinnum
    If skinnum >= aliashdrs(pmdl).mdl.numskins Or skinnum < 0 Then
        Con_DPrintf "R_AliasSetupSkin: no such skin # %d\n", skinnum
        skinnum = 0
    End If

    pskindesc = aliashdrs(paliashdr).skindesc + skinnum   ' alias_skindescs() index
    a_skinwidth = aliashdrs(pmdl).mdl.skinwidth
    skin = alias_skindescs(pskindesc).skin

    If alias_skindescs(pskindesc).typ = ALIAS_SKIN_GROUP Then
        paliasskingroup = skin           ' alias_skingroups() index
        pskinintervals = alias_skingroups(paliasskingroup).intervals
        numskins = alias_skingroups(paliasskingroup).numskins
        fullskininterval = alias_intervals(pskinintervals + numskins - 1)

        skintime = cl.time + cl_entities(currententity).syncbase

        ' when loading in Mod_LoadAliasSkinGroup, we guaranteed all interval
        ' values are positive, so we don't have to worry about division by 0
        skintargettime = skintime - (ToInt(skintime / fullskininterval)) * fullskininterval

        For i = 0 To numskins - 2
            If alias_intervals(pskinintervals + i) > skintargettime Then Exit For
        Next

        skin = alias_skingroups(paliasskingroup).skindescs(i).skin
    End If

    r_affinetridesc.pskin = skin
    r_affinetridesc.skinwidth = a_skinwidth
    r_affinetridesc.seamfixupX16 = (a_skinwidth \ 2) * &H10000
    r_affinetridesc.skinheight = aliashdrs(pmdl).mdl.skinheight
End Sub

Private Sub R_AliasSetupLighting(plighting As alight_t)
    ' guarantee that no vertex will ever be lit below LIGHT_MIN, so we don't have
    ' to clamp off the bottom
    r_ambientlight = plighting.ambientlight

    If r_ambientlight < LIGHT_MIN Then r_ambientlight = LIGHT_MIN

    r_ambientlight = (255 - r_ambientlight) * VID_GRADES   ' << VID_CBITS

    If r_ambientlight < LIGHT_MIN Then r_ambientlight = LIGHT_MIN

    r_shadelight = plighting.shadelight

    If r_shadelight < 0 Then r_shadelight = 0

    r_shadelight = r_shadelight * VID_GRADES

    ' rotate the lighting vector into the model's frame of reference
    r_plightvec.v(0) = DotProduct(plighting.plightvec, alias_forward)
    r_plightvec.v(1) = -DotProduct(plighting.plightvec, alias_right)
    r_plightvec.v(2) = DotProduct(plighting.plightvec, alias_up)
End Sub

' set r_apverts
Private Sub R_AliasSetupFrame()
    Dim frame As Long, i As Long, numframes As Long, paliasgroup As Long, pintervals As Long
    Dim fullinterval As Single, targettime As Single, time_ As Single

    frame = cl_entities(currententity).frame
    If frame >= aliashdrs(pmdl).mdl.numframes Or frame < 0 Then
        Con_DPrintf "R_AliasSetupFrame: no such frame %d\n", frame
        frame = 0
    End If

    If aliashdrs(paliashdr).frames(frame).typ = ALIAS_SINGLE Then
        r_apverts = aliashdrs(paliashdr).frames(frame).frame
        Exit Sub
    End If

    paliasgroup = aliashdrs(paliashdr).frames(frame).frame   ' alias_groups() index
    pintervals = alias_groups(paliasgroup).intervals
    numframes = alias_groups(paliasgroup).numframes
    fullinterval = alias_intervals(pintervals + numframes - 1)

    time_ = cl.time + cl_entities(currententity).syncbase

    '
    ' when loading in Mod_LoadAliasGroup, we guaranteed all interval values
    ' are positive, so we don't have to worry about division by 0
    '
    targettime = time_ - (ToInt(time_ / fullinterval)) * fullinterval

    For i = 0 To numframes - 2
        If alias_intervals(pintervals + i) > targettime Then Exit For
    Next

    r_apverts = alias_groups(paliasgroup).frames(i).frame
End Sub

' (QuakeVB) acolormap = currententity->colormap
Private Sub R_AliasSetColormap()
    Dim cm As Long

    cm = cl_entities(currententity).colormap
    If cm = 0 Then
        SetView VarPtrArray(acolormap), sa_acolormap, VarPtr(vid_colormap(0)), UBound(vid_colormap) + 1, 1
    Else
        SetView VarPtrArray(acolormap), sa_acolormap, VarPtr(cl.scores(cm - 1).translations(0)), VID_GRADES * 256, 1
    End If
End Sub

' (QuakeVB) detaches acolormap() before the program ends (D_ClearViews)
Public Sub R_AliasClearViews()
    ClearView VarPtrArray(acolormap)
End Sub

Public Sub R_AliasDrawModel(plighting As alight_t)
    r_amodels_drawn = r_amodels_drawn + 1

    ' (C: the vertex arrays are locals, finalverts cache aligned)
    If Not aliasverts_made Then
        ReDim a_finalverts(0 To FV_CLIP + 2 * 8 - 1)
        ReDim a_auxverts(0 To MAXALIASVERTS - 1)
        aliasverts_made = True
    End If

    paliashdr = Mod_Extradata(cl_entities(currententity).model)
    pmdl = paliashdr

    R_AliasSetupSkin
    R_AliasSetUpTransform cl_entities(currententity).trivial_accept
    R_AliasSetupLighting plighting
    R_AliasSetupFrame

    ' (C stops with "R_AliasDrawModel: !currententity->colormap"; colormap 0
    ' is vid.colormap here)

    If cl_entities(currententity).trivial_accept = 3 And r_recursiveaffinetriangles Then
        r_affinetridesc.drawtype = 1
    Else
        r_affinetridesc.drawtype = 0
    End If

    If r_affinetridesc.drawtype Then
        D_PolysetUpdateTables            ' FIXME: precalc...
    End If

    R_AliasSetColormap

    If currententity <> CL_VIEWENT Then
        ziscale = 32768! * 65536!
    Else
        ziscale = 32768! * 65536! * 3!
    End If

    If cl_entities(currententity).trivial_accept Then
        R_AliasPrepareUnclippedPoints
    Else
        R_AliasPreparePoints
    End If
End Sub
