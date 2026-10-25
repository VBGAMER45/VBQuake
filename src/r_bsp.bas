Attribute VB_Name = "r_bsp"
' r_bsp.bas - walks the world BSP and clips brush models to it (r_bsp.c)
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

'
' current entity info
'
Public insubmodel As Boolean
Public currententity As Long             ' cl_entities() index
Public modelorg As vec3_t, base_modelorg As vec3_t
                                         ' modelorg is the viewpoint reletive to
                                         ' the currently rendering entity
Public r_entorigin As vec3_t             ' the currently rendering entity in world
                                         ' coordinates

Public entity_rotation(0 To 2, 0 To 2) As Single

Public r_worldmodelorg As vec3_t

Public r_currentbkey As Long

' solidstate_t
Private Const touchessolid As Long = 0
Private Const drawnode As Long = 1
Private Const nodrawnode As Long = 2

Private Const MAX_BMODEL_VERTS As Long = 500     ' 6K
Private Const MAX_BMODEL_EDGES As Long = 1000    ' 12K

' (C: pbverts and pbedges point at R_DrawSolidClippedSubmodelPolygons' local
' arrays; a bedge_t's vertex is a mod_vertexes() index, or -1 - n for
' r_bverts(n))
Public r_bverts(0 To MAX_BMODEL_VERTS - 1) As mvertex_t
Public r_bedges(0 To MAX_BMODEL_EDGES - 1) As bedge_t
Private numbverts As Long, numbedges As Long

Private pfrontenter As Long, pfrontexit As Long  ' vertex references, as in bedge_t

Private makeclippededge As Boolean

Private btofpolys() As btofpoly_t       ' C: R_RenderWorld's local (dynamic: VB's 64K of module data)
Private btofpolys_made As Boolean

'===========================================================================

Private Sub R_EntityRotate(vec As vec3_t)
    Dim tvec As vec3_t

    tvec = vec
    vec.v(0) = entity_rotation(0, 0) * tvec.v(0) + entity_rotation(0, 1) * tvec.v(1) + entity_rotation(0, 2) * tvec.v(2)
    vec.v(1) = entity_rotation(1, 0) * tvec.v(0) + entity_rotation(1, 1) * tvec.v(1) + entity_rotation(1, 2) * tvec.v(2)
    vec.v(2) = entity_rotation(2, 0) * tvec.v(0) + entity_rotation(2, 1) * tvec.v(1) + entity_rotation(2, 2) * tvec.v(2)
End Sub

Public Sub R_RotateBmodel()
    Dim angle As Single, s As Single, c As Single
    Dim temp1(0 To 2, 0 To 2) As Single, temp2(0 To 2, 0 To 2) As Single, temp3(0 To 2, 0 To 2) As Single

    ' TODO: should use a look-up table
    ' TODO: should really be stored with the entity instead of being reconstructed
    ' TODO: could cache lazily, stored in the entity
    ' TODO: share work with R_SetUpAliasTransform

    ' yaw
    angle = cl_entities(currententity).angles.v(YAW)
    angle = angle * M_PI * 2 / 360
    s = Sin(angle)
    c = Cos(angle)

    temp1(0, 0) = c
    temp1(0, 1) = s
    temp1(0, 2) = 0
    temp1(1, 0) = -s
    temp1(1, 1) = c
    temp1(1, 2) = 0
    temp1(2, 0) = 0
    temp1(2, 1) = 0
    temp1(2, 2) = 1

    ' pitch
    angle = cl_entities(currententity).angles.v(PITCH)
    angle = angle * M_PI * 2 / 360
    s = Sin(angle)
    c = Cos(angle)

    temp2(0, 0) = c
    temp2(0, 1) = 0
    temp2(0, 2) = -s
    temp2(1, 0) = 0
    temp2(1, 1) = 1
    temp2(1, 2) = 0
    temp2(2, 0) = s
    temp2(2, 1) = 0
    temp2(2, 2) = c

    R_ConcatRotations temp2, temp1, temp3

    ' roll
    angle = cl_entities(currententity).angles.v(ROLL)
    angle = angle * M_PI * 2 / 360
    s = Sin(angle)
    c = Cos(angle)

    temp1(0, 0) = 1
    temp1(0, 1) = 0
    temp1(0, 2) = 0
    temp1(1, 0) = 0
    temp1(1, 1) = c
    temp1(1, 2) = s
    temp1(2, 0) = 0
    temp1(2, 1) = -s
    temp1(2, 2) = c

    R_ConcatRotations temp1, temp3, entity_rotation

    '
    ' rotate modelorg and the transformation matrix
    '
    R_EntityRotate modelorg
    R_EntityRotate vpn
    R_EntityRotate vright
    R_EntityRotate vup

    R_TransformFrustum
End Sub

' (QuakeVB) the position of a vertex reference
Private Sub VertPos(ByVal ref As Long, out As vec3_t)
    If ref >= 0 Then
        out = mod_vertexes(ref).position
    Else
        out = r_bverts(-1 - ref).position
    End If
End Sub

Private Sub R_RecursiveClipBPoly(ByVal pedges As Long, ByVal pnode As Long, ByVal psurf As Long)
    Dim psideedges(0 To 1) As Long, pnextedge As Long, ptedge As Long
    Dim i As Long, side As Long, lastside As Long, dist As Single, frac As Single, lastdist As Single
    Dim splitplane As Long, tplane As mplane_t, pvert As Long, plastvert As Long, ptvert As Long, pn As Long
    Dim lastpos As vec3_t, pos As vec3_t

    psideedges(0) = -1
    psideedges(1) = -1

    makeclippededge = False

    ' transform the BSP plane into model space
    ' FIXME: cache these?
    splitplane = mod_nodes(pnode).plane
    tplane.dist = mod_planes(splitplane).dist - DotProduct(r_entorigin, mod_planes(splitplane).normal)
    tplane.normal.v(0) = entity_rotation(0, 0) * mod_planes(splitplane).normal.v(0) _
        + entity_rotation(0, 1) * mod_planes(splitplane).normal.v(1) + entity_rotation(0, 2) * mod_planes(splitplane).normal.v(2)
    tplane.normal.v(1) = entity_rotation(1, 0) * mod_planes(splitplane).normal.v(0) _
        + entity_rotation(1, 1) * mod_planes(splitplane).normal.v(1) + entity_rotation(1, 2) * mod_planes(splitplane).normal.v(2)
    tplane.normal.v(2) = entity_rotation(2, 0) * mod_planes(splitplane).normal.v(0) _
        + entity_rotation(2, 1) * mod_planes(splitplane).normal.v(1) + entity_rotation(2, 2) * mod_planes(splitplane).normal.v(2)

    ' clip edges to BSP plane
    Do While pedges <> -1
        pnextedge = r_bedges(pedges).pnext

        ' set the status for the last point as the previous point
        ' FIXME: cache this stuff somehow?
        plastvert = r_bedges(pedges).v(0)
        VertPos plastvert, lastpos
        lastdist = DotProduct(lastpos, tplane.normal) - tplane.dist

        If lastdist > 0 Then
            lastside = 0
        Else
            lastside = 1
        End If

        pvert = r_bedges(pedges).v(1)
        VertPos pvert, pos

        dist = DotProduct(pos, tplane.normal) - tplane.dist

        If dist > 0 Then
            side = 0
        Else
            side = 1
        End If

        If side <> lastside Then
            ' clipped
            If numbverts >= MAX_BMODEL_VERTS Then Exit Sub

            ' generate the clipped vertex
            frac = lastdist / (lastdist - dist)
            ptvert = -1 - numbverts
            With r_bverts(numbverts)
                .position.v(0) = lastpos.v(0) + frac * (pos.v(0) - lastpos.v(0))
                .position.v(1) = lastpos.v(1) + frac * (pos.v(1) - lastpos.v(1))
                .position.v(2) = lastpos.v(2) + frac * (pos.v(2) - lastpos.v(2))
            End With
            numbverts = numbverts + 1

            ' split into two edges, one on each side, and remember entering
            ' and exiting points
            ' FIXME: share the clip edge by having a winding direction flag?
            If numbedges >= (MAX_BMODEL_EDGES - 1) Then
                Con_Printf "Out of edges for bmodel\n"
                Exit Sub
            End If

            ptedge = numbedges
            r_bedges(ptedge).pnext = psideedges(lastside)
            psideedges(lastside) = ptedge
            r_bedges(ptedge).v(0) = plastvert
            r_bedges(ptedge).v(1) = ptvert

            ptedge = numbedges + 1
            r_bedges(ptedge).pnext = psideedges(side)
            psideedges(side) = ptedge
            r_bedges(ptedge).v(0) = ptvert
            r_bedges(ptedge).v(1) = pvert

            numbedges = numbedges + 2

            If side = 0 Then
                ' entering for front, exiting for back
                pfrontenter = ptvert
                makeclippededge = True
            Else
                pfrontexit = ptvert
                makeclippededge = True
            End If
        Else
            ' add the edge to the appropriate side
            r_bedges(pedges).pnext = psideedges(side)
            psideedges(side) = pedges
        End If

        pedges = pnextedge
    Loop

    ' if anything was clipped, reconstitute and add the edges along the clip
    ' plane to both sides (but in opposite directions)
    If makeclippededge Then
        If numbedges >= (MAX_BMODEL_EDGES - 2) Then
            Con_Printf "Out of edges for bmodel\n"
            Exit Sub
        End If

        ptedge = numbedges
        r_bedges(ptedge).pnext = psideedges(0)
        psideedges(0) = ptedge
        r_bedges(ptedge).v(0) = pfrontexit
        r_bedges(ptedge).v(1) = pfrontenter

        ptedge = numbedges + 1
        r_bedges(ptedge).pnext = psideedges(1)
        psideedges(1) = ptedge
        r_bedges(ptedge).v(0) = pfrontenter
        r_bedges(ptedge).v(1) = pfrontexit

        numbedges = numbedges + 2
    End If

    ' draw or recurse further
    For i = 0 To 1
        If psideedges(i) <> -1 Then
            ' draw if we've reached a non-solid leaf, done if all that's left is a
            ' solid leaf, and continue down the tree if it's not a leaf
            pn = mod_nodes(pnode).children(i)

            ' we're done with this branch if the node or leaf isn't in the PVS
            If mod_nodes(pn).visframe = r_visframecount Then
                If mod_nodes(pn).contents < 0 Then
                    If mod_nodes(pn).contents <> CONTENTS_SOLID Then
                        r_currentbkey = mod_nodes(pn).key
                        R_RenderBmodelFace psideedges(i), psurf
                    End If
                Else
                    R_RecursiveClipBPoly psideedges(i), mod_nodes(pnode).children(i), psurf
                End If
            End If
        End If
    Next
End Sub

' (QuakeVB) C's test of whether the viewer sees a surface's front
Private Function FacesViewer(ByVal psurf As Long) As Boolean
    Dim pplane As Long, dot As Single

    ' find which side of the node we are on
    pplane = mod_surfaces(psurf).plane

    dot = DotProduct(modelorg, mod_planes(pplane).normal) - mod_planes(pplane).dist

    ' draw the polygon
    If (mod_surfaces(psurf).flags And SURF_PLANEBACK) <> 0 Then
        FacesViewer = dot < -BACKFACE_EPSILON
    Else
        FacesViewer = dot > BACKFACE_EPSILON
    End If
End Function

Public Sub R_DrawSolidClippedSubmodelPolygons(ByVal pmodel As Long)
    Dim i As Long, j As Long, lindex As Long, psurf As Long, numsurfaces As Long
    Dim pbedge As Long, pedges As Long, pedge As Long

    ' FIXME: use bounding-box-based frustum clipping info?

    psurf = mod_known(pmodel).surfaces + mod_known(pmodel).firstmodelsurface
    numsurfaces = mod_known(pmodel).nummodelsurfaces
    pedges = mod_known(pmodel).edges

    For i = 0 To numsurfaces - 1
        If FacesViewer(psurf) Then
            ' FIXME: use bounding-box-based frustum clipping info?

            ' copy the edges to bedges, flipping if necessary so always
            ' clockwise winding
            ' FIXME: if edges and vertices get caches, these assignments must move
            ' outside the loop, and overflow checking must be done here
            numbverts = 0
            numbedges = 0

            If mod_surfaces(psurf).numedges > 0 Then
                pbedge = numbedges
                numbedges = numbedges + mod_surfaces(psurf).numedges

                For j = 0 To mod_surfaces(psurf).numedges - 1
                    lindex = mod_surfedges(mod_known(pmodel).surfedges + mod_surfaces(psurf).firstedge + j)

                    If lindex > 0 Then
                        pedge = pedges + lindex
                        r_bedges(pbedge + j).v(0) = r_pcurrentvertbase + mod_edges(pedge).v(0)
                        r_bedges(pbedge + j).v(1) = r_pcurrentvertbase + mod_edges(pedge).v(1)
                    Else
                        lindex = -lindex
                        pedge = pedges + lindex
                        r_bedges(pbedge + j).v(0) = r_pcurrentvertbase + mod_edges(pedge).v(1)
                        r_bedges(pbedge + j).v(1) = r_pcurrentvertbase + mod_edges(pedge).v(0)
                    End If

                    r_bedges(pbedge + j).pnext = pbedge + j + 1
                Next

                r_bedges(pbedge + j - 1).pnext = -1   ' mark end of edges

                R_RecursiveClipBPoly pbedge, cl_entities(currententity).topnode, psurf
            Else
                Sys_Error "no edges in bmodel"
            End If
        End If
        psurf = psurf + 1
    Next
End Sub

Public Sub R_DrawSubmodelPolygons(ByVal pmodel As Long, ByVal clipflags As Long)
    Dim i As Long, psurf As Long, numsurfaces As Long

    ' FIXME: use bounding-box-based frustum clipping info?

    psurf = mod_known(pmodel).surfaces + mod_known(pmodel).firstmodelsurface
    numsurfaces = mod_known(pmodel).nummodelsurfaces

    For i = 0 To numsurfaces - 1
        If FacesViewer(psurf) Then
            r_currentkey = mod_nodes(cl_entities(currententity).topnode).key

            ' FIXME: use bounding-box-based frustum clipping info?
            R_RenderFace psurf, clipflags
        End If
        psurf = psurf + 1
    Next
End Sub

Private Sub R_RecursiveWorldNode(ByVal node As Long, ByVal clipflags As Long)
    Dim i As Long, c As Long, side As Long, pindex As Long, acceptpt As vec3_t, rejectpt As vec3_t
    Dim plane As Long, surf As Long, mark As Long, d As Double, dot As Double

    If mod_nodes(node).contents = CONTENTS_SOLID Then Exit Sub   ' solid

    If mod_nodes(node).visframe <> r_visframecount Then Exit Sub

    ' cull the clipping planes if not trivial accept
    ' FIXME: the compiler is doing a lousy job of optimizing here; it could be
    '  twice as fast in ASM
    If clipflags Then
        For i = 0 To 3
            If (clipflags And BIT_(i)) = 0 Then GoTo nextplane   ' don't need to clip against it

            ' generate accept and reject points
            ' FIXME: do with fast look-ups or integer tests based on the sign bit
            ' of the floating point values

            pindex = pfrustum_indexes(i)

            rejectpt.v(0) = mod_nodes(node).minmaxs(r_frustum_indexes(pindex + 0))
            rejectpt.v(1) = mod_nodes(node).minmaxs(r_frustum_indexes(pindex + 1))
            rejectpt.v(2) = mod_nodes(node).minmaxs(r_frustum_indexes(pindex + 2))

            d = DotProduct(rejectpt, view_clipplanes(i).normal)
            d = d - view_clipplanes(i).dist

            If d <= 0 Then Exit Sub

            acceptpt.v(0) = mod_nodes(node).minmaxs(r_frustum_indexes(pindex + 3 + 0))
            acceptpt.v(1) = mod_nodes(node).minmaxs(r_frustum_indexes(pindex + 3 + 1))
            acceptpt.v(2) = mod_nodes(node).minmaxs(r_frustum_indexes(pindex + 3 + 2))

            d = DotProduct(acceptpt, view_clipplanes(i).normal)
            d = d - view_clipplanes(i).dist

            If d >= 0 Then clipflags = clipflags And Not BIT_(i)   ' node is entirely on screen
nextplane:
        Next
    End If

    ' if a leaf node, draw stuff
    If mod_nodes(node).contents < 0 Then
        mark = mod_nodes(node).firstmarksurface
        c = mod_nodes(node).nummarksurfaces

        Do While c > 0
            mod_surfaces(mod_marksurfaces(mark)).visframe = r_framecount
            mark = mark + 1
            c = c - 1
        Loop

        ' deal with model fragments in this leaf
        If mod_nodes(node).efrags Then R_StoreEfrags mod_nodes(node).efrags

        mod_nodes(node).key = r_currentkey
        r_currentkey = r_currentkey + 1  ' all bmodels in a leaf share the same key
    Else
        ' node is just a decision point, so go down the apropriate sides

        ' find which side of the node we are on
        plane = mod_nodes(node).plane

        Select Case mod_planes(plane).typ
        Case PLANE_X
            dot = modelorg.v(0) - mod_planes(plane).dist
        Case PLANE_Y
            dot = modelorg.v(1) - mod_planes(plane).dist
        Case PLANE_Z
            dot = modelorg.v(2) - mod_planes(plane).dist
        Case Else
            dot = DotProduct(modelorg, mod_planes(plane).normal) - mod_planes(plane).dist
        End Select

        If dot >= 0 Then
            side = 0
        Else
            side = 1
        End If

        ' recurse down the children, front side first
        R_RecursiveWorldNode mod_nodes(node).children(side), clipflags

        ' draw stuff
        c = mod_nodes(node).numsurfaces

        If c Then
            surf = mod_known(cl.worldmodel).surfaces + mod_nodes(node).firstsurface

            If dot < -BACKFACE_EPSILON Then
                Do
                    If (mod_surfaces(surf).flags And SURF_PLANEBACK) <> 0 And mod_surfaces(surf).visframe = r_framecount Then
                        If r_drawpolys Then
                            If r_worldpolysbacktofront Then
                                If numbtofpolys < MAX_BTOFPOLYS Then
                                    btofpolys(numbtofpolys).clipflags = clipflags
                                    btofpolys(numbtofpolys).psurf = surf
                                    numbtofpolys = numbtofpolys + 1
                                End If
                            Else
                                R_RenderPoly surf, clipflags
                            End If
                        Else
                            R_RenderFace surf, clipflags
                        End If
                    End If

                    surf = surf + 1
                    c = c - 1
                Loop While c > 0
            ElseIf dot > BACKFACE_EPSILON Then
                Do
                    If (mod_surfaces(surf).flags And SURF_PLANEBACK) = 0 And mod_surfaces(surf).visframe = r_framecount Then
                        If r_drawpolys Then
                            If r_worldpolysbacktofront Then
                                If numbtofpolys < MAX_BTOFPOLYS Then
                                    btofpolys(numbtofpolys).clipflags = clipflags
                                    btofpolys(numbtofpolys).psurf = surf
                                    numbtofpolys = numbtofpolys + 1
                                End If
                            Else
                                R_RenderPoly surf, clipflags
                            End If
                        Else
                            R_RenderFace surf, clipflags
                        End If
                    End If

                    surf = surf + 1
                    c = c - 1
                Loop While c > 0
            End If

            ' all surfaces on the same node share the same sequence number
            r_currentkey = r_currentkey + 1
        End If

        ' recurse down the back side
        R_RecursiveWorldNode mod_nodes(node).children(1 - side), clipflags
    End If
End Sub

Public Sub R_RenderWorld()
    Dim i As Long, clmodel As Long

    ' (C: pbtofpolys = btofpolys, a local array; here a module array)
    If Not btofpolys_made Then
        ReDim btofpolys(0 To MAX_BTOFPOLYS - 1)
        btofpolys_made = True
    End If

    currententity = 0
    modelorg = r_origin
    clmodel = cl_entities(currententity).model
    r_pcurrentvertbase = mod_known(clmodel).vertexes

    R_RecursiveWorldNode mod_known(clmodel).nodes, 15

    ' if the driver wants the polygons back to front, play the visible ones back
    ' in that order
    If r_worldpolysbacktofront Then
        For i = numbtofpolys - 1 To 0 Step -1
            R_RenderPoly btofpolys(i).psurf, btofpolys(i).clipflags
        Next
    End If
End Sub
