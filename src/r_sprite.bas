Attribute VB_Name = "r_sprite"
' r_sprite.bas - sets up sprites for drawing (r_sprite.c)
'
' A vec5_t of clip_verts is clip_verts(buffer, vertex, 0 .. 4); the float
' pointers that walk them are vertex numbers. The projected points go to
' r_spriteverts() (C: R_SetupAndDrawSprite's local outverts, which
' r_spritedesc.pverts points at).
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private clip_current As Long
Private clip_verts(0 To 1, 0 To MAXWORKINGVERTS - 1, 0 To 4) As Single
Private sprite_width As Long, sprite_height As Long

Public r_spritedesc As spritedesc_t
Public r_spriteverts(0 To MAXWORKINGVERTS) As emitpoint_t

Private Sub R_RotateSprite(ByVal beamlength As Single)
    Dim vec As vec3_t

    If beamlength = 0 Then Exit Sub

    VectorScale r_spritedesc.vpn, -beamlength, vec
    VectorAdd r_entorigin, vec, r_entorigin
    VectorSubtract modelorg, vec, modelorg
End Sub

' Clips the winding at clip_verts[clip_current] and changes clip_current
' Throws out the back side
' pclipplane: a view_clipplanes() index
Private Function R_ClipSpriteFace(ByVal nump As Long, ByVal pclipplane As Long) As Long
    Dim i As Long, k As Long, outcount As Long, dists(0 To MAXWORKINGVERTS) As Single
    Dim frac As Single, clipdist As Single, inb As Long, outb As Long
    Dim n0 As Single, n1 As Single, n2 As Single

    clipdist = view_clipplanes(pclipplane).dist
    n0 = view_clipplanes(pclipplane).normal.v(0)
    n1 = view_clipplanes(pclipplane).normal.v(1)
    n2 = view_clipplanes(pclipplane).normal.v(2)

    ' calc dists
    If clip_current Then
        inb = 1
        outb = 0
        clip_current = 0
    Else
        inb = 0
        outb = 1
        clip_current = 1
    End If

    For i = 0 To nump - 1
        dists(i) = clip_verts(inb, i, 0) * n0 + clip_verts(inb, i, 1) * n1 + clip_verts(inb, i, 2) * n2 - clipdist
    Next

    ' handle wraparound case
    dists(nump) = dists(0)
    For k = 0 To 4
        clip_verts(inb, nump, k) = clip_verts(inb, 0, k)
    Next

    ' clip the winding
    outcount = 0

    For i = 0 To nump - 1
        If dists(i) >= 0 Then
            For k = 0 To 4
                clip_verts(outb, outcount, k) = clip_verts(inb, i, k)
            Next
            outcount = outcount + 1
        End If

        If dists(i) = 0 Or dists(i + 1) = 0 Then GoTo nexti

        If (dists(i) > 0) = (dists(i + 1) > 0) Then GoTo nexti

        ' split it into a new vertex
        frac = dists(i) / (dists(i) - dists(i + 1))

        For k = 0 To 4
            clip_verts(outb, outcount, k) = clip_verts(inb, i, k) + frac * (clip_verts(inb, i + 1, k) - clip_verts(inb, i, k))
        Next

        outcount = outcount + 1
nexti:
    Next

    R_ClipSpriteFace = outcount
End Function

Private Sub R_SetupAndDrawSprite()
    Dim i As Long, nump As Long, dot As Single, scale_ As Single, f As Long
    Dim left_ As vec3_t, up As vec3_t, right_ As vec3_t, down As vec3_t, transformed As vec3_t, local_ As vec3_t
    Dim pv As vec3_t

    dot = DotProduct(r_spritedesc.vpn, modelorg)

    ' backface cull
    If dot >= 0 Then Exit Sub

    ' build the sprite poster in worldspace
    f = r_spritedesc.pspriteframe
    VectorScale r_spritedesc.vright, sprite_frames(f).right_, right_
    VectorScale r_spritedesc.vup, sprite_frames(f).up, up
    VectorScale r_spritedesc.vright, sprite_frames(f).left_, left_
    VectorScale r_spritedesc.vup, sprite_frames(f).down, down

    For i = 0 To 2
        clip_verts(0, 0, i) = r_entorigin.v(i) + up.v(i) + left_.v(i)
        clip_verts(0, 1, i) = r_entorigin.v(i) + up.v(i) + right_.v(i)
        clip_verts(0, 2, i) = r_entorigin.v(i) + down.v(i) + right_.v(i)
        clip_verts(0, 3, i) = r_entorigin.v(i) + down.v(i) + left_.v(i)
    Next
    clip_verts(0, 0, 3) = 0
    clip_verts(0, 0, 4) = 0
    clip_verts(0, 1, 3) = sprite_width
    clip_verts(0, 1, 4) = 0
    clip_verts(0, 2, 3) = sprite_width
    clip_verts(0, 2, 4) = sprite_height
    clip_verts(0, 3, 3) = 0
    clip_verts(0, 3, 4) = sprite_height

    ' clip to the frustum in worldspace
    nump = 4
    clip_current = 0

    For i = 0 To 3
        nump = R_ClipSpriteFace(nump, i)
        If nump < 3 Then Exit Sub
        If nump >= MAXWORKINGVERTS Then Sys_Error "R_SetupAndDrawSprite: too many points"
    Next

    ' transform vertices into viewspace and project
    r_spritedesc.nearzi = -999999

    For i = 0 To nump - 1
        pv.v(0) = clip_verts(clip_current, i, 0)
        pv.v(1) = clip_verts(clip_current, i, 1)
        pv.v(2) = clip_verts(clip_current, i, 2)
        VectorSubtract pv, r_origin, local_
        TransformVector local_, transformed

        If transformed.v(2) < NEAR_CLIP Then transformed.v(2) = NEAR_CLIP

        r_spriteverts(i).zi = 1! / transformed.v(2)
        If r_spriteverts(i).zi > r_spritedesc.nearzi Then r_spritedesc.nearzi = r_spriteverts(i).zi

        r_spriteverts(i).s = clip_verts(clip_current, i, 3)
        r_spriteverts(i).t = clip_verts(clip_current, i, 4)

        scale_ = xscale * r_spriteverts(i).zi
        r_spriteverts(i).u = (xcenter + scale_ * transformed.v(0))

        scale_ = yscale * r_spriteverts(i).zi
        r_spriteverts(i).v = (ycenter - scale_ * transformed.v(1))
    Next

    ' draw it
    r_spritedesc.nump = nump
    D_DrawSprite
End Sub

' psprite: an msprites() index; returns a sprite_frames() index
Private Function R_GetSpriteframe(ByVal psprite As Long) As Long
    Dim pspritegroup As Long, i As Long, numframes As Long, frame As Long, pintervals As Long
    Dim fullinterval As Single, targettime As Single, time_ As Single

    frame = cl_entities(currententity).frame

    If frame >= msprites(psprite).numframes Or frame < 0 Then
        Con_Printf "R_DrawSprite: no such frame %d\n", frame
        frame = 0
    End If

    If msprites(psprite).frames(frame).typ = SPR_SINGLE Then
        R_GetSpriteframe = msprites(psprite).frames(frame).frameptr
    Else
        pspritegroup = msprites(psprite).frames(frame).frameptr   ' sprite_groups() index
        pintervals = sprite_groups(pspritegroup).intervals
        numframes = sprite_groups(pspritegroup).numframes
        fullinterval = sprite_intervals(pintervals + numframes - 1)

        time_ = cl.time + cl_entities(currententity).syncbase

        ' when loading in Mod_LoadSpriteGroup, we guaranteed all interval values
        ' are positive, so we don't have to worry about division by 0
        targettime = time_ - (ToInt(time_ / fullinterval)) * fullinterval

        For i = 0 To numframes - 2
            If sprite_intervals(pintervals + i) > targettime Then Exit For
        Next

        R_GetSpriteframe = sprite_groups(pspritegroup).frames(i)
    End If
End Function

Public Sub R_DrawSprite()
    Dim i As Long, psprite As Long, tvec As vec3_t, dot As Single, angle As Single, sr As Single, cr As Single

    psprite = mod_known(cl_entities(currententity).model).cache

    r_spritedesc.pspriteframe = R_GetSpriteframe(psprite)

    sprite_width = sprite_frames(r_spritedesc.pspriteframe).width
    sprite_height = sprite_frames(r_spritedesc.pspriteframe).height

    ' TODO: make this caller-selectable
    Select Case msprites(psprite).typ
    Case SPR_FACING_UPRIGHT
        ' generate the sprite's axes, with vup straight up in worldspace, and
        ' r_spritedesc.vright perpendicular to modelorg.
        ' This will not work if the view direction is very close to straight up or
        ' down, because the cross product will be between two nearly parallel
        ' vectors and starts to approach an undefined state, so we don't draw if
        ' the two vectors are less than 1 degree apart
        tvec.v(0) = -modelorg.v(0)
        tvec.v(1) = -modelorg.v(1)
        tvec.v(2) = -modelorg.v(2)
        VectorNormalize tvec
        dot = tvec.v(2)                  ' same as DotProduct (tvec, r_spritedesc.vup) because
                                         '  r_spritedesc.vup is 0, 0, 1
        If dot > 0.999848 Or dot < -0.999848 Then Exit Sub   ' cos(1 degree) = 0.999848
        r_spritedesc.vup.v(0) = 0
        r_spritedesc.vup.v(1) = 0
        r_spritedesc.vup.v(2) = 1
        r_spritedesc.vright.v(0) = tvec.v(1)
                                         ' CrossProduct(r_spritedesc.vup, -modelorg,
        r_spritedesc.vright.v(1) = -tvec.v(0)
                                         '              r_spritedesc.vright)
        r_spritedesc.vright.v(2) = 0
        VectorNormalize r_spritedesc.vright
        r_spritedesc.vpn.v(0) = -r_spritedesc.vright.v(1)
        r_spritedesc.vpn.v(1) = r_spritedesc.vright.v(0)
        r_spritedesc.vpn.v(2) = 0
                    ' CrossProduct (r_spritedesc.vright, r_spritedesc.vup,
                    '  r_spritedesc.vpn)

    Case SPR_VP_PARALLEL
        ' generate the sprite's axes, completely parallel to the viewplane. There
        ' are no problem situations, because the sprite is always in the same
        ' position relative to the viewer
        For i = 0 To 2
            r_spritedesc.vup.v(i) = vup.v(i)
            r_spritedesc.vright.v(i) = vright.v(i)
            r_spritedesc.vpn.v(i) = vpn.v(i)
        Next

    Case SPR_VP_PARALLEL_UPRIGHT
        ' generate the sprite's axes, with vup straight up in worldspace, and
        ' r_spritedesc.vright parallel to the viewplane.
        ' This will not work if the view direction is very close to straight up or
        ' down, because the cross product will be between two nearly parallel
        ' vectors and starts to approach an undefined state, so we don't draw if
        ' the two vectors are less than 1 degree apart
        dot = vpn.v(2)                   ' same as DotProduct (vpn, r_spritedesc.vup) because
                                         '  r_spritedesc.vup is 0, 0, 1
        If dot > 0.999848 Or dot < -0.999848 Then Exit Sub   ' cos(1 degree) = 0.999848
        r_spritedesc.vup.v(0) = 0
        r_spritedesc.vup.v(1) = 0
        r_spritedesc.vup.v(2) = 1
        r_spritedesc.vright.v(0) = vpn.v(1)
                                         ' CrossProduct (r_spritedesc.vup, vpn,
        r_spritedesc.vright.v(1) = -vpn.v(0)   '  r_spritedesc.vright)
        r_spritedesc.vright.v(2) = 0
        VectorNormalize r_spritedesc.vright
        r_spritedesc.vpn.v(0) = -r_spritedesc.vright.v(1)
        r_spritedesc.vpn.v(1) = r_spritedesc.vright.v(0)
        r_spritedesc.vpn.v(2) = 0
                    ' CrossProduct (r_spritedesc.vright, r_spritedesc.vup,
                    '  r_spritedesc.vpn)

    Case SPR_ORIENTED
        ' generate the sprite's axes, according to the sprite's world orientation
        AngleVectors cl_entities(currententity).angles, r_spritedesc.vpn, r_spritedesc.vright, r_spritedesc.vup

    Case SPR_VP_PARALLEL_ORIENTED
        ' generate the sprite's axes, parallel to the viewplane, but rotated in
        ' that plane around the center according to the sprite entity's roll
        ' angle. So vpn stays the same, but vright and vup rotate
        angle = cl_entities(currententity).angles.v(ROLL) * (M_PI * 2 / 360)
        sr = Sin(angle)
        cr = Cos(angle)

        For i = 0 To 2
            r_spritedesc.vpn.v(i) = vpn.v(i)
            r_spritedesc.vright.v(i) = vright.v(i) * cr + vup.v(i) * sr
            r_spritedesc.vup.v(i) = vright.v(i) * -sr + vup.v(i) * cr
        Next

    Case Else
        Sys_Error "R_DrawSprite: Bad sprite type %d", msprites(psprite).typ
    End Select

    R_RotateSprite msprites(psprite).beamlength

    R_SetupAndDrawSprite
End Sub
