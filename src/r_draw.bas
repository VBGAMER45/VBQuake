Attribute VB_Name = "r_draw"
' r_draw.bas - clips and emits the edges of the visible polygons (r_draw.c)
'
' The vertex pointers of C (mvertex_t *) are ByRef mvertex_t parameters:
' R_ClipEdge's clipped vertex is a local, as in C. r_pedge is the mod_edges()
' index of the edge being emitted, -1 for C's dummy tedge.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const MAXLEFTCLIPEDGES As Long = 100

' !!! if these are changed, they must be changed in asm_draw.h too !!!
Private Const FULLY_CLIPPED_CACHED As Long = &H80000000
Private Const FRAMECOUNT_MASK As Long = &H7FFFFFFF

Private cacheoffset As Long              ' C: unsigned int

Public c_faceclip As Long                ' number of faces clipped

Public r_zpointdesc As zpointdesc_t

Public r_polydesc As polydesc_t
Public r_polyverts(0 To 99) As polyvert_t    ' C: r_polydesc.pverts, R_RenderPoly's pverts

Public view_clipplanes(0 To 3) As clipplane_t

Private r_pedge As Long

Public r_leftclipped As Boolean, r_rightclipped As Boolean
Private makeleftedge As Boolean, makerightedge As Boolean
Public r_nearzionly As Boolean

Public sintable(0 To SIN_BUFFER_SIZE - 1) As Long
Public intsintable(0 To SIN_BUFFER_SIZE - 1) As Long

Public r_leftenter As mvertex_t, r_leftexit As mvertex_t
Public r_rightenter As mvertex_t, r_rightexit As mvertex_t

Public r_emitted As Long
Public r_nearzi As Single
Public r_u1 As Single, r_v1 As Single, r_lzi1 As Single
Public r_ceilv1 As Long

Public r_lastvertvalid As Boolean

' (QuakeVB) C's (int) ceil (x)
Private Function CeilInt(ByVal x As Single) As Long
    CeilInt = ToInt(-Int(-x))
End Function

Public Sub R_EmitEdge(pv0 As mvertex_t, pv1 As mvertex_t)
    Dim edge As Long, pcheck As Long, u_check As Long, u As Single, u_step As Single
    Dim local_ As vec3_t, transformed As vec3_t, v As Long, v2 As Long, ceilv0 As Long
    Dim scale_ As Single, lzi0 As Single, u0 As Single, v0 As Single, side As Long

    If r_lastvertvalid Then
        u0 = r_u1
        v0 = r_v1
        lzi0 = r_lzi1
        ceilv0 = r_ceilv1
    Else
        ' transform and project
        VectorSubtract pv0.position, modelorg, local_
        TransformVector local_, transformed

        If transformed.v(2) < NEAR_CLIP Then transformed.v(2) = NEAR_CLIP

        lzi0 = 1# / transformed.v(2)

        ' FIXME: build x/yscale into transform?
        scale_ = xscale * lzi0
        u0 = (xcenter + scale_ * transformed.v(0))
        If u0 < r_refdef.fvrectx_adj Then u0 = r_refdef.fvrectx_adj
        If u0 > r_refdef.fvrectright_adj Then u0 = r_refdef.fvrectright_adj

        scale_ = yscale * lzi0
        v0 = (ycenter - scale_ * transformed.v(1))
        If v0 < r_refdef.fvrecty_adj Then v0 = r_refdef.fvrecty_adj
        If v0 > r_refdef.fvrectbottom_adj Then v0 = r_refdef.fvrectbottom_adj

        ceilv0 = CeilInt(v0)
    End If

    ' transform and project
    VectorSubtract pv1.position, modelorg, local_
    TransformVector local_, transformed

    If transformed.v(2) < NEAR_CLIP Then transformed.v(2) = NEAR_CLIP

    r_lzi1 = 1# / transformed.v(2)

    scale_ = xscale * r_lzi1
    r_u1 = (xcenter + scale_ * transformed.v(0))
    If r_u1 < r_refdef.fvrectx_adj Then r_u1 = r_refdef.fvrectx_adj
    If r_u1 > r_refdef.fvrectright_adj Then r_u1 = r_refdef.fvrectright_adj

    scale_ = yscale * r_lzi1
    r_v1 = (ycenter - scale_ * transformed.v(1))
    If r_v1 < r_refdef.fvrecty_adj Then r_v1 = r_refdef.fvrecty_adj
    If r_v1 > r_refdef.fvrectbottom_adj Then r_v1 = r_refdef.fvrectbottom_adj

    If r_lzi1 > lzi0 Then lzi0 = r_lzi1

    If lzi0 > r_nearzi Then r_nearzi = lzi0   ' for mipmap finding

    ' for right edges, all we want is the effect on 1/z
    If r_nearzionly Then Exit Sub

    r_emitted = 1

    r_ceilv1 = CeilInt(r_v1)

    ' create the edge
    If ceilv0 = r_ceilv1 Then
        ' we cache unclipped horizontal edges as fully clipped
        If cacheoffset <> &H7FFFFFFF Then cacheoffset = FULLY_CLIPPED_CACHED Or (r_framecount And FRAMECOUNT_MASK)

        Exit Sub                         ' horizontal edge
    End If

    side = -(ceilv0 > r_ceilv1)

    edge = edge_p
    edge_p = edge_p + 1

    With r_edges(edge)
        .owner = r_pedge

        .nearzi = lzi0

        If side = 0 Then
            ' trailing edge (go from p1 to p2)
            v = ceilv0
            v2 = r_ceilv1 - 1

            .surfs(0) = surface_p
            .surfs(1) = 0

            u_step = ((r_u1 - u0) / (r_v1 - v0))
            u = u0 + (CSng(v) - v0) * u_step
        Else
            ' leading edge (go from p2 to p1)
            v2 = ceilv0 - 1
            v = r_ceilv1

            .surfs(0) = 0
            .surfs(1) = surface_p

            u_step = ((u0 - r_u1) / (v0 - r_v1))
            u = r_u1 + (CSng(v) - r_v1) * u_step
        End If

        .u_step = ToInt(u_step * &H100000)
        .u = ToInt(u * &H100000 + &HFFFFF)

        ' we need to do this to avoid stepping off the edges if a very nearly
        ' horizontal edge is less than epsilon above a scan, and numeric error causes
        ' it to incorrectly extend to the scan, and the extension of the line goes off
        ' the edge of the screen
        ' FIXME: is this actually needed?
        If .u < r_refdef.vrect_x_adj_shift20 Then .u = r_refdef.vrect_x_adj_shift20
        If .u > r_refdef.vrectright_adj_shift20 Then .u = r_refdef.vrectright_adj_shift20

        '
        ' sort the edge in normally
        '
        u_check = .u
        If .surfs(0) Then u_check = u_check + 1   ' sort trailers after leaders
    End With

    pcheck = newedges(v)
    If pcheck = -1 Then
        r_edges(edge).nxt = newedges(v)
        newedges(v) = edge
    ElseIf r_edges(pcheck).u >= u_check Then
        r_edges(edge).nxt = newedges(v)
        newedges(v) = edge
    Else
        Do While r_edges(pcheck).nxt <> -1
            If r_edges(r_edges(pcheck).nxt).u >= u_check Then Exit Do
            pcheck = r_edges(pcheck).nxt
        Loop
        r_edges(edge).nxt = r_edges(pcheck).nxt
        r_edges(pcheck).nxt = edge
    End If

    r_edges(edge).nextremove = removeedges(v2)
    removeedges(v2) = edge
End Sub

' clip: a view_clipplanes() index, -1 = NULL
Public Sub R_ClipEdge(pv0 As mvertex_t, pv1 As mvertex_t, ByVal clip As Long)
    Dim d0 As Single, d1 As Single, f As Single, clipvert As mvertex_t

    If clip <> -1 Then
        Do
            d0 = DotProduct(pv0.position, view_clipplanes(clip).normal) - view_clipplanes(clip).dist
            d1 = DotProduct(pv1.position, view_clipplanes(clip).normal) - view_clipplanes(clip).dist

            If d0 >= 0 Then
                ' point 0 is unclipped
                If d1 >= 0 Then
                    ' both points are unclipped
                    GoTo continue_
                End If

                ' only point 1 is clipped

                ' we don't cache clipped edges
                cacheoffset = &H7FFFFFFF

                f = d0 / (d0 - d1)
                clipvert.position.v(0) = pv0.position.v(0) + f * (pv1.position.v(0) - pv0.position.v(0))
                clipvert.position.v(1) = pv0.position.v(1) + f * (pv1.position.v(1) - pv0.position.v(1))
                clipvert.position.v(2) = pv0.position.v(2) + f * (pv1.position.v(2) - pv0.position.v(2))

                If view_clipplanes(clip).leftedge Then
                    r_leftclipped = True
                    r_leftexit = clipvert
                ElseIf view_clipplanes(clip).rightedge Then
                    r_rightclipped = True
                    r_rightexit = clipvert
                End If

                R_ClipEdge pv0, clipvert, view_clipplanes(clip).nxt
                Exit Sub
            Else
                ' point 0 is clipped
                If d1 < 0 Then
                    ' both points are clipped
                    ' we do cache fully clipped edges
                    If Not r_leftclipped Then cacheoffset = FULLY_CLIPPED_CACHED Or (r_framecount And FRAMECOUNT_MASK)
                    Exit Sub
                End If

                ' only point 0 is clipped
                r_lastvertvalid = False

                ' we don't cache partially clipped edges
                cacheoffset = &H7FFFFFFF

                f = d0 / (d0 - d1)
                clipvert.position.v(0) = pv0.position.v(0) + f * (pv1.position.v(0) - pv0.position.v(0))
                clipvert.position.v(1) = pv0.position.v(1) + f * (pv1.position.v(1) - pv0.position.v(1))
                clipvert.position.v(2) = pv0.position.v(2) + f * (pv1.position.v(2) - pv0.position.v(2))

                If view_clipplanes(clip).leftedge Then
                    r_leftclipped = True
                    r_leftenter = clipvert
                ElseIf view_clipplanes(clip).rightedge Then
                    r_rightclipped = True
                    r_rightenter = clipvert
                End If

                R_ClipEdge clipvert, pv1, view_clipplanes(clip).nxt
                Exit Sub
            End If
continue_:
            clip = view_clipplanes(clip).nxt
        Loop While clip <> -1
    End If

    ' add the edge
    R_EmitEdge pv0, pv1
End Sub

Private Sub R_EmitCachedEdge()
    Dim pedge_t As Long

    pedge_t = EDGES_BASE + mod_edges(r_pedge).cachededgeoffset

    If r_edges(pedge_t).surfs(0) = 0 Then
        r_edges(pedge_t).surfs(0) = surface_p
    Else
        r_edges(pedge_t).surfs(1) = surface_p
    End If

    If r_edges(pedge_t).nearzi > r_nearzi Then r_nearzi = r_edges(pedge_t).nearzi   ' for mipmap finding

    r_emitted = 1
End Sub

' (QuakeVB) the edge cache test of R_RenderFace: True if the edge was emitted
' from the cache or is cached as fully clipped (C: continue)
Private Function EdgeCached() As Boolean
    Dim ofs As Long

    If insubmodel Then Exit Function

    ofs = mod_edges(r_pedge).cachededgeoffset
    If ofs And FULLY_CLIPPED_CACHED Then
        If (ofs And FRAMECOUNT_MASK) = r_framecount Then
            r_lastvertvalid = False
            EdgeCached = True
        End If
    Else
        If (edge_p - EDGES_BASE) > ofs Then
            If r_edges(EDGES_BASE + ofs).owner = r_pedge Then
                R_EmitCachedEdge
                r_lastvertvalid = False
                EdgeCached = True
            End If
        End If
    End If
End Function

' (QuakeVB) the end of R_RenderFace and R_RenderBmodelFace: posts the surface
Private Sub PostSurface(ByVal fa As Long, ByVal insub As Boolean, ByVal key As Long)
    Dim pplane As Long, distinv As Single, p_normal As vec3_t

    r_polycount = r_polycount + 1

    With surfaces(surface_p)
        .data = fa
        .nearzi = r_nearzi
        .flags = mod_surfaces(fa).flags
        .insubmodel = insub
        .spanstate = 0
        .entity = currententity
        .key = key
        .spans = -1

        pplane = mod_surfaces(fa).plane
        ' FIXME: cache this?
        TransformVector mod_planes(pplane).normal, p_normal
        ' FIXME: cache this?
        distinv = 1# / (mod_planes(pplane).dist - DotProduct(modelorg, mod_planes(pplane).normal))

        .d_zistepu = p_normal.v(0) * xscaleinv * distinv
        .d_zistepv = -p_normal.v(1) * yscaleinv * distinv
        .d_ziorigin = p_normal.v(2) * distinv - xcenter * .d_zistepu - ycenter * .d_zistepv
    End With

    ' JDC VectorCopy (r_worldmodelorg, surface_p->modelorg);
    surface_p = surface_p + 1
End Sub

Public Sub R_RenderFace(ByVal fa As Long, ByVal clipflags As Long)
    Dim i As Long, lindex As Long, mask As Long, pedges As Long, pclip As Long, vbase As Long

    ' skip out if no more surfs
    If surface_p >= surf_max Then
        r_outofsurfaces = r_outofsurfaces + 1
        Exit Sub
    End If

    ' ditto if not enough edges left, or switch to auxedges if possible
    If (edge_p + mod_surfaces(fa).numedges + 4) >= edge_max Then
        r_outofedges = r_outofedges + mod_surfaces(fa).numedges
        Exit Sub
    End If

    c_faceclip = c_faceclip + 1

    ' set up clip planes
    pclip = -1

    mask = &H8
    For i = 3 To 0 Step -1
        If clipflags And mask Then
            view_clipplanes(i).nxt = pclip
            pclip = i
        End If
        mask = mask \ 2
    Next

    ' push the edges through
    r_emitted = 0
    r_nearzi = 0
    r_nearzionly = False
    makeleftedge = False
    makerightedge = False
    pedges = mod_known(cl_entities(currententity).model).edges
    vbase = r_pcurrentvertbase
    r_lastvertvalid = False

    For i = 0 To mod_surfaces(fa).numedges - 1
        lindex = mod_surfedges(mod_known(cl_entities(currententity).model).surfedges + mod_surfaces(fa).firstedge + i)

        If lindex > 0 Then
            r_pedge = pedges + lindex

            ' if the edge is cached, we can just reuse the edge
            If EdgeCached() Then GoTo nexti

            ' assume it's cacheable
            cacheoffset = edge_p - EDGES_BASE
            r_leftclipped = False
            r_rightclipped = False
            R_ClipEdge mod_vertexes(vbase + mod_edges(r_pedge).v(0)), mod_vertexes(vbase + mod_edges(r_pedge).v(1)), pclip
            mod_edges(r_pedge).cachededgeoffset = cacheoffset

            If r_leftclipped Then makeleftedge = True
            If r_rightclipped Then makerightedge = True
            r_lastvertvalid = True
        Else
            lindex = -lindex
            r_pedge = pedges + lindex
            ' if the edge is cached, we can just reuse the edge
            ' (it's cached if the cached edge is valid and is owned
            ' by this medge_t)
            If EdgeCached() Then GoTo nexti

            ' assume it's cacheable
            cacheoffset = edge_p - EDGES_BASE
            r_leftclipped = False
            r_rightclipped = False
            R_ClipEdge mod_vertexes(vbase + mod_edges(r_pedge).v(1)), mod_vertexes(vbase + mod_edges(r_pedge).v(0)), pclip
            mod_edges(r_pedge).cachededgeoffset = cacheoffset

            If r_leftclipped Then makeleftedge = True
            If r_rightclipped Then makerightedge = True
            r_lastvertvalid = True
        End If
nexti:
    Next

    ' if there was a clip off the left edge, add that edge too
    ' FIXME: faster to do in screen space?
    ' FIXME: share clipped edges?
    If makeleftedge Then
        r_pedge = -1                     ' &tedge
        r_lastvertvalid = False
        R_ClipEdge r_leftexit, r_leftenter, view_clipplanes(pclip).nxt
    End If

    ' if there was a clip off the right edge, get the right r_nearzi
    If makerightedge Then
        r_pedge = -1                     ' &tedge
        r_lastvertvalid = False
        r_nearzionly = True
        R_ClipEdge r_rightexit, r_rightenter, view_clipplanes(1).nxt
    End If

    ' if no edges made it out, return without posting the surface
    If r_emitted = 0 Then Exit Sub

    PostSurface fa, insubmodel, r_currentkey
    r_currentkey = r_currentkey + 1
End Sub

' (QuakeVB) a vertex of a clipped bmodel edge (bedge_t.v)
Private Sub BVert(ByVal ref As Long, out As mvertex_t)
    If ref >= 0 Then
        out = mod_vertexes(ref)
    Else
        out = r_bverts(-1 - ref)
    End If
End Sub

Public Sub R_RenderBmodelFace(ByVal pedges As Long, ByVal psurf As Long)
    Dim i As Long, mask As Long, pclip As Long, v0 As mvertex_t, v1 As mvertex_t

    ' skip out if no more surfs
    If surface_p >= surf_max Then
        r_outofsurfaces = r_outofsurfaces + 1
        Exit Sub
    End If

    ' ditto if not enough edges left, or switch to auxedges if possible
    If (edge_p + mod_surfaces(psurf).numedges + 4) >= edge_max Then
        r_outofedges = r_outofedges + mod_surfaces(psurf).numedges
        Exit Sub
    End If

    c_faceclip = c_faceclip + 1

    ' this is a dummy to give the caching mechanism someplace to write to
    r_pedge = -1                         ' &tedge

    ' set up clip planes
    pclip = -1

    mask = &H8
    For i = 3 To 0 Step -1
        If r_clipflags And mask Then
            view_clipplanes(i).nxt = pclip
            pclip = i
        End If
        mask = mask \ 2
    Next

    ' push the edges through
    r_emitted = 0
    r_nearzi = 0
    r_nearzionly = False
    makeleftedge = False
    makerightedge = False
    ' FIXME: keep clipped bmodel edges in clockwise order so last vertex caching
    ' can be used?
    r_lastvertvalid = False

    Do While pedges <> -1
        r_leftclipped = False
        r_rightclipped = False
        BVert r_bedges(pedges).v(0), v0
        BVert r_bedges(pedges).v(1), v1
        R_ClipEdge v0, v1, pclip

        If r_leftclipped Then makeleftedge = True
        If r_rightclipped Then makerightedge = True

        pedges = r_bedges(pedges).pnext
    Loop

    ' if there was a clip off the left edge, add that edge too
    ' FIXME: faster to do in screen space?
    ' FIXME: share clipped edges?
    If makeleftedge Then
        r_pedge = -1                     ' &tedge
        R_ClipEdge r_leftexit, r_leftenter, view_clipplanes(pclip).nxt
    End If

    ' if there was a clip off the right edge, get the right r_nearzi
    If makerightedge Then
        r_pedge = -1                     ' &tedge
        r_nearzionly = True
        R_ClipEdge r_rightexit, r_rightenter, view_clipplanes(1).nxt
    End If

    ' if no edges made it out, return without posting the surface
    If r_emitted = 0 Then Exit Sub

    PostSurface psurf, True, r_currentbkey
End Sub

' (unused by the software driver: r_drawpolys is never set)
Public Sub R_RenderPoly(ByVal fa As Long, ByVal clipflags As Long)
    Dim i As Long, lindex As Long, lnumverts As Long, s_axis As Long, t_axis As Long
    Dim dist As Single, lastdist As Single, lzi As Single, scale_ As Single, u As Single, v As Single, frac As Single
    Dim mask As Long, local_ As vec3_t, transformed As vec3_t, pclip As Long, pedges As Long, pplane As Long
    Dim verts(0 To 1, 0 To 99) As mvertex_t  ' FIXME: do real number
    Dim vertpage As Long, newverts As Long, newpage As Long, lastvert As Long, visible As Boolean

    ' FIXME: clean this up and make it faster
    ' FIXME: guard against running out of vertices

    s_axis = 0                           ' keep compiler happy
    t_axis = 0

    ' set up clip planes
    pclip = -1

    mask = &H8
    For i = 3 To 0 Step -1
        If clipflags And mask Then
            view_clipplanes(i).nxt = pclip
            pclip = i
        End If
        mask = mask \ 2
    Next

    ' reconstruct the polygon
    ' FIXME: these should be precalculated and loaded off disk
    pedges = mod_known(cl_entities(currententity).model).edges
    lnumverts = mod_surfaces(fa).numedges
    vertpage = 0

    For i = 0 To lnumverts - 1
        lindex = mod_surfedges(mod_known(cl_entities(currententity).model).surfedges + mod_surfaces(fa).firstedge + i)

        If lindex > 0 Then
            r_pedge = pedges + lindex
            verts(0, i) = mod_vertexes(r_pcurrentvertbase + mod_edges(r_pedge).v(0))
        Else
            r_pedge = pedges - lindex
            verts(0, i) = mod_vertexes(r_pcurrentvertbase + mod_edges(r_pedge).v(1))
        End If
    Next

    ' clip the polygon, done if not visible
    Do While pclip <> -1
        lastvert = lnumverts - 1
        lastdist = DotProduct(verts(vertpage, lastvert).position, view_clipplanes(pclip).normal) - view_clipplanes(pclip).dist

        visible = False
        newverts = 0
        newpage = vertpage Xor 1

        For i = 0 To lnumverts - 1
            dist = DotProduct(verts(vertpage, i).position, view_clipplanes(pclip).normal) - view_clipplanes(pclip).dist

            If (lastdist > 0) <> (dist > 0) Then
                frac = dist / (dist - lastdist)
                verts(newpage, newverts).position.v(0) = verts(vertpage, i).position.v(0) _
                    + ((verts(vertpage, lastvert).position.v(0) - verts(vertpage, i).position.v(0)) * frac)
                verts(newpage, newverts).position.v(1) = verts(vertpage, i).position.v(1) _
                    + ((verts(vertpage, lastvert).position.v(1) - verts(vertpage, i).position.v(1)) * frac)
                verts(newpage, newverts).position.v(2) = verts(vertpage, i).position.v(2) _
                    + ((verts(vertpage, lastvert).position.v(2) - verts(vertpage, i).position.v(2)) * frac)
                newverts = newverts + 1
            End If

            If dist >= 0 Then
                verts(newpage, newverts) = verts(vertpage, i)
                newverts = newverts + 1
                visible = True
            End If

            lastvert = i
            lastdist = dist
        Next

        If Not visible Or newverts < 3 Then Exit Sub

        lnumverts = newverts
        vertpage = vertpage Xor 1
        pclip = view_clipplanes(pclip).nxt
    Loop

    ' transform and project, remembering the z values at the vertices and
    ' r_nearzi, and extract the s and t coordinates at the vertices
    pplane = mod_surfaces(fa).plane
    Select Case mod_planes(pplane).typ
    Case PLANE_X, PLANE_ANYX
        s_axis = 1
        t_axis = 2
    Case PLANE_Y, PLANE_ANYY
        s_axis = 0
        t_axis = 2
    Case PLANE_Z, PLANE_ANYZ
        s_axis = 0
        t_axis = 1
    End Select

    r_nearzi = 0

    For i = 0 To lnumverts - 1
        ' transform and project
        VectorSubtract verts(vertpage, i).position, modelorg, local_
        TransformVector local_, transformed

        If transformed.v(2) < NEAR_CLIP Then transformed.v(2) = NEAR_CLIP

        lzi = 1# / transformed.v(2)

        If lzi > r_nearzi Then r_nearzi = lzi   ' for mipmap finding

        ' FIXME: build x/yscale into transform?
        scale_ = xscale * lzi
        u = (xcenter + scale_ * transformed.v(0))
        If u < r_refdef.fvrectx_adj Then u = r_refdef.fvrectx_adj
        If u > r_refdef.fvrectright_adj Then u = r_refdef.fvrectright_adj

        scale_ = yscale * lzi
        v = (ycenter - scale_ * transformed.v(1))
        If v < r_refdef.fvrecty_adj Then v = r_refdef.fvrecty_adj
        If v > r_refdef.fvrectbottom_adj Then v = r_refdef.fvrectbottom_adj

        r_polyverts(i).u = u
        r_polyverts(i).v = v
        r_polyverts(i).zi = lzi
        r_polyverts(i).s = verts(vertpage, i).position.v(s_axis)
        r_polyverts(i).t = verts(vertpage, i).position.v(t_axis)
    Next

    ' build the polygon descriptor, including fa, r_nearzi, and u, v, s, t, and z
    ' for each vertex
    r_polydesc.numverts = lnumverts
    r_polydesc.nearzi = r_nearzi
    r_polydesc.pcurrentface = fa

    ' draw the polygon
    D_DrawPoly
End Sub

Public Sub R_ZDrawSubmodelPolys(ByVal pmodel As Long)
    Dim i As Long, numsurfaces As Long, psurf As Long, dot As Single, pplane As Long

    psurf = mod_known(pmodel).surfaces + mod_known(pmodel).firstmodelsurface
    numsurfaces = mod_known(pmodel).nummodelsurfaces

    For i = 0 To numsurfaces - 1
        ' find which side of the node we are on
        pplane = mod_surfaces(psurf).plane

        dot = DotProduct(modelorg, mod_planes(pplane).normal) - mod_planes(pplane).dist

        ' draw the polygon
        If ((mod_surfaces(psurf).flags And SURF_PLANEBACK) <> 0 And dot < -BACKFACE_EPSILON) Or _
           ((mod_surfaces(psurf).flags And SURF_PLANEBACK) = 0 And dot > BACKFACE_EPSILON) Then
            ' FIXME: use bounding-box-based frustum clipping info?
            R_RenderPoly psurf, 15
        End If
        psurf = psurf + 1
    Next
End Sub
