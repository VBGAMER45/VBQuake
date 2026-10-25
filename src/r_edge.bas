Attribute VB_Name = "r_edge"
' r_edge.bas - the active edge table: turns the edges into spans per surface
' (r_edge.c)
'
' Edges, surfaces and spans are arrays and their pointers indices (-1 =
' NULL). r_edges() starts with the four fixed edges of C (edge_head,
' edge_tail, edge_aftertail, edge_sentinel), then the frame's edges from
' EDGES_BASE, which is C's r_edges[0]. The spans are spans(), C's local
' basespans of R_ScanEdges.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public Const EDGE_HEAD As Long = 0
Public Const EDGE_TAIL As Long = 1
Public Const EDGE_AFTERTAIL As Long = 2
Public Const EDGE_SENTINEL As Long = 3
Public Const EDGES_BASE As Long = 4

Public r_edges() As edge_t
Public edge_p As Long, edge_max As Long

Public surfaces() As surf_t
Public surface_p As Long, surf_max As Long

' surfaces are generated in back to front order by the bsp, so if a surf
' pointer is greater than another one, it should be drawn in front
' surfaces[1] is the background, and is used as the active surface stack

Public newedges(0 To MAXHEIGHT - 1) As Long
Public removeedges(0 To MAXHEIGHT - 1) As Long

Public spans() As espan_t                ' (dynamic: VB's 64K of module data)
Private spans_made As Boolean
Public span_p As Long, max_span_p As Long

Public r_currentkey As Long

Private current_iv As Long

Private edge_head_u_shift20 As Long, edge_tail_u_shift20 As Long

Private drawbackward As Boolean          ' C: pdrawfunc is R_GenerateSpansBackward

Private fv As Single

'=============================================================================

Private Sub R_DrawCulledPolys_()
    Dim s As Long, pface As Long

    currententity = 0

    If r_worldpolysbacktofront Then
        For s = surface_p - 1 To 2 Step -1
            If surfaces(s).spans <> -1 Then
                If (surfaces(s).flags And SURF_DRAWBACKGROUND) = 0 Then
                    pface = surfaces(s).data
                    R_RenderPoly pface, 15
                End If
            End If
        Next
    Else
        For s = 1 To surface_p - 1
            If surfaces(s).spans <> -1 Then
                If (surfaces(s).flags And SURF_DRAWBACKGROUND) = 0 Then
                    pface = surfaces(s).data
                    R_RenderPoly pface, 15
                End If
            End If
        Next
    End If
End Sub

Public Sub R_BeginEdgeFrame()
    Dim v As Long

    edge_p = EDGES_BASE
    edge_max = EDGES_BASE + r_numallocatededges

    surface_p = 2                        ' background is surface 1,
                                         '  surface 0 is a dummy
    surfaces(1).spans = -1               ' no background spans yet
    surfaces(1).flags = SURF_DRAWBACKGROUND

    ' put the background behind everything in the world
    If cvars(cv_r_draworder).value <> 0 Then
        drawbackward = True
        surfaces(1).key = 0
        r_currentkey = 1
    Else
        drawbackward = False
        surfaces(1).key = &H7FFFFFFF
        r_currentkey = 0
    End If

    ' FIXME: set with memset
    For v = r_refdef.vrect.y To r_refdef.vrectbottom - 1
        newedges(v) = -1
        removeedges(v) = -1
    Next
End Sub

' Adds the edges in the linked list edgestoadd, adding them to the edges in the
' linked list edgelist.  edgestoadd is assumed to be sorted on u, and non-empty (this is actually newedges[v]).  edgelist is assumed to be sorted on u, with a
' sentinel at the end (actually, this is the active edge table starting at
' edge_head.next).
Private Sub R_InsertNewEdges(ByVal edgestoadd As Long, ByVal edgelist As Long)
    Dim next_edge As Long

    Do
        next_edge = r_edges(edgestoadd).nxt
        ' (C's edgesearch loop, unrolled four times)
        Do While r_edges(edgelist).u < r_edges(edgestoadd).u
            edgelist = r_edges(edgelist).nxt
        Loop

        ' insert edgestoadd before edgelist
        r_edges(edgestoadd).nxt = edgelist
        r_edges(edgestoadd).prev = r_edges(edgelist).prev
        r_edges(r_edges(edgelist).prev).nxt = edgestoadd
        r_edges(edgelist).prev = edgestoadd

        edgestoadd = next_edge
    Loop While edgestoadd <> -1
End Sub

Private Sub R_RemoveEdges(ByVal pedge As Long)
    Do
        r_edges(r_edges(pedge).nxt).prev = r_edges(pedge).prev
        r_edges(r_edges(pedge).prev).nxt = r_edges(pedge).nxt
        pedge = r_edges(pedge).nextremove
    Loop While pedge <> -1
End Sub

Private Sub R_StepActiveU(ByVal pedge As Long)
    Dim pnext_edge As Long, pwedge As Long

    ' (C's nextedge loop, unrolled four times)
    Do
        r_edges(pedge).u = r_edges(pedge).u + r_edges(pedge).u_step
        If r_edges(pedge).u < r_edges(r_edges(pedge).prev).u Then
            ' pushback
            If pedge = EDGE_AFTERTAIL Then Exit Sub

            ' push it back to keep it sorted
            pnext_edge = r_edges(pedge).nxt

            ' pull the edge out of the edge list
            r_edges(r_edges(pedge).nxt).prev = r_edges(pedge).prev
            r_edges(r_edges(pedge).prev).nxt = r_edges(pedge).nxt

            ' find out where the edge goes in the edge list
            pwedge = r_edges(r_edges(pedge).prev).prev

            Do While r_edges(pwedge).u > r_edges(pedge).u
                pwedge = r_edges(pwedge).prev
            Loop

            ' put the edge back into the edge list
            r_edges(pedge).nxt = r_edges(pwedge).nxt
            r_edges(pedge).prev = pwedge
            r_edges(r_edges(pedge).nxt).prev = pedge
            r_edges(pwedge).nxt = pedge

            pedge = pnext_edge
            If pedge = EDGE_TAIL Then Exit Sub
        Else
            pedge = r_edges(pedge).nxt
        End If
    Loop
End Sub

' (QuakeVB) emits a span from surf's last_u to iu (C repeats this inline)
Private Sub EmitSpan(ByVal surf As Long, ByVal iu As Long)
    Dim span As Long

    span = span_p
    span_p = span_p + 1
    spans(span).u = surfaces(surf).last_u
    spans(span).count = iu - spans(span).u
    spans(span).v = current_iv
    spans(span).pnext = surfaces(surf).spans
    surfaces(surf).spans = span
End Sub

Private Sub R_CleanupSpan()
    Dim surf As Long, iu As Long

    ' now that we've reached the right edge of the screen, we're done with any
    ' unfinished surfaces, so emit a span for whatever's on top
    surf = surfaces(1).nxt
    iu = edge_tail_u_shift20
    If iu > surfaces(surf).last_u Then EmitSpan surf, iu

    ' reset spanstate for all surfaces in the surface stack
    Do
        surfaces(surf).spanstate = 0
        surf = surfaces(surf).nxt
    Loop While surf <> 1
End Sub

Private Sub R_LeadingEdgeBackwards(ByVal edge As Long)
    Dim surf As Long, surf2 As Long, iu As Long

    ' it's adding a new surface in, so find the correct place
    surf = r_edges(edge).surfs(1)

    ' don't start a span if this is an inverted span, with the end
    ' edge preceding the start edge (that is, we've already seen the
    ' end edge)
    surfaces(surf).spanstate = surfaces(surf).spanstate + 1
    If surfaces(surf).spanstate = 1 Then
        surf2 = surfaces(1).nxt

        If surfaces(surf).key > surfaces(surf2).key Then GoTo newtop

        ' if it's two surfaces on the same plane, the one that's already
        ' active is in front, so keep going unless it's a bmodel
        If surfaces(surf).insubmodel And (surfaces(surf).key = surfaces(surf2).key) Then
            ' must be two bmodels in the same leaf; don't care, because they'll
            ' never be farthest anyway
            GoTo newtop
        End If

continue_search:

        Do
            surf2 = surfaces(surf2).nxt
        Loop While surfaces(surf).key < surfaces(surf2).key

        If surfaces(surf).key = surfaces(surf2).key Then
            ' if it's two surfaces on the same plane, the one that's already
            ' active is in front, so keep going unless it's a bmodel
            If Not surfaces(surf).insubmodel Then GoTo continue_search

            ' must be two bmodels in the same leaf; don't care which is really
            ' in front, because they'll never be farthest anyway
        End If

        GoTo gotposition

newtop:
        ' emit a span (obscures current top)
        iu = (r_edges(edge).u And &HFFF00000) \ &H100000   ' >> 20

        If iu > surfaces(surf2).last_u Then EmitSpan surf2, iu

        ' set last_u on the new span
        surfaces(surf).last_u = iu

gotposition:
        ' insert before surf2
        surfaces(surf).nxt = surf2
        surfaces(surf).prev = surfaces(surf2).prev
        surfaces(surfaces(surf2).prev).nxt = surf
        surfaces(surf2).prev = surf
    End If
End Sub

Private Sub R_TrailingEdge(ByVal surf As Long, ByVal edge As Long)
    Dim iu As Long

    ' don't generate a span if this is an inverted span, with the end
    ' edge preceding the start edge (that is, we haven't seen the
    ' start edge yet)
    surfaces(surf).spanstate = surfaces(surf).spanstate - 1
    If surfaces(surf).spanstate = 0 Then
        If surfaces(surf).insubmodel Then r_bmodelactive = r_bmodelactive - 1

        If surf = surfaces(1).nxt Then
            ' emit a span (current top going away)
            iu = (r_edges(edge).u And &HFFF00000) \ &H100000   ' >> 20
            If iu > surfaces(surf).last_u Then EmitSpan surf, iu

            ' set last_u on the surface below
            surfaces(surfaces(surf).nxt).last_u = iu
        End If

        surfaces(surfaces(surf).prev).nxt = surfaces(surf).nxt
        surfaces(surfaces(surf).nxt).prev = surfaces(surf).prev
    End If
End Sub

' (QuakeVB) C's two bmodels in the same leaf: True if surf is in front of
' surf2 at the edge (sorted on 1/z)
Private Function BmodelInFront(ByVal surf As Long, ByVal surf2 As Long, ByVal edge As Long) As Boolean
    Dim fu As Double, newzi As Double, testzi As Double, newzitop As Double, newzibottom As Double

    fu = CSng(r_edges(edge).u - &HFFFFF) * (1# / &H100000)
    newzi = surfaces(surf).d_ziorigin + fv * surfaces(surf).d_zistepv + fu * surfaces(surf).d_zistepu
    newzibottom = newzi * 0.99

    testzi = surfaces(surf2).d_ziorigin + fv * surfaces(surf2).d_zistepv + fu * surfaces(surf2).d_zistepu

    If newzibottom >= testzi Then
        BmodelInFront = True
        Exit Function
    End If

    newzitop = newzi * 1.01
    If newzitop >= testzi Then
        If surfaces(surf).d_zistepu >= surfaces(surf2).d_zistepu Then BmodelInFront = True
    End If
End Function

Private Sub R_LeadingEdge(ByVal edge As Long)
    Dim surf As Long, surf2 As Long, iu As Long

    If r_edges(edge).surfs(1) Then
        ' it's adding a new surface in, so find the correct place
        surf = r_edges(edge).surfs(1)

        ' don't start a span if this is an inverted span, with the end
        ' edge preceding the start edge (that is, we've already seen the
        ' end edge)
        surfaces(surf).spanstate = surfaces(surf).spanstate + 1
        If surfaces(surf).spanstate = 1 Then
            If surfaces(surf).insubmodel Then r_bmodelactive = r_bmodelactive + 1

            surf2 = surfaces(1).nxt

            If surfaces(surf).key < surfaces(surf2).key Then GoTo newtop

            ' if it's two surfaces on the same plane, the one that's already
            ' active is in front, so keep going unless it's a bmodel
            If surfaces(surf).insubmodel And (surfaces(surf).key = surfaces(surf2).key) Then
                ' must be two bmodels in the same leaf; sort on 1/z
                If BmodelInFront(surf, surf2, edge) Then GoTo newtop
            End If

continue_search:

            Do
                surf2 = surfaces(surf2).nxt
            Loop While surfaces(surf).key > surfaces(surf2).key

            If surfaces(surf).key = surfaces(surf2).key Then
                ' if it's two surfaces on the same plane, the one that's already
                ' active is in front, so keep going unless it's a bmodel
                If Not surfaces(surf).insubmodel Then GoTo continue_search

                ' must be two bmodels in the same leaf; sort on 1/z
                If BmodelInFront(surf, surf2, edge) Then GoTo gotposition

                GoTo continue_search
            End If

            GoTo gotposition

newtop:
            ' emit a span (obscures current top)
            iu = (r_edges(edge).u And &HFFF00000) \ &H100000   ' >> 20

            If iu > surfaces(surf2).last_u Then EmitSpan surf2, iu

            ' set last_u on the new span
            surfaces(surf).last_u = iu

gotposition:
            ' insert before surf2
            surfaces(surf).nxt = surf2
            surfaces(surf).prev = surfaces(surf2).prev
            surfaces(surfaces(surf2).prev).nxt = surf
            surfaces(surf2).prev = surf
        End If
    End If
End Sub

Private Sub R_GenerateSpans()
    Dim edge As Long, surf As Long

    r_bmodelactive = 0

    ' clear active surfaces to just the background surface
    surfaces(1).nxt = 1
    surfaces(1).prev = 1
    surfaces(1).last_u = edge_head_u_shift20

    ' generate spans
    edge = r_edges(EDGE_HEAD).nxt
    Do While edge <> EDGE_TAIL
        If r_edges(edge).surfs(0) Then
            ' it has a left surface, so a surface is going away for this span
            surf = r_edges(edge).surfs(0)

            R_TrailingEdge surf, edge

            If r_edges(edge).surfs(1) = 0 Then GoTo nextedge
        End If

        R_LeadingEdge edge
nextedge:
        edge = r_edges(edge).nxt
    Loop

    R_CleanupSpan
End Sub

Private Sub R_GenerateSpansBackward()
    Dim edge As Long

    r_bmodelactive = 0

    ' clear active surfaces to just the background surface
    surfaces(1).nxt = 1
    surfaces(1).prev = 1
    surfaces(1).last_u = edge_head_u_shift20

    ' generate spans
    edge = r_edges(EDGE_HEAD).nxt
    Do While edge <> EDGE_TAIL
        If r_edges(edge).surfs(0) Then R_TrailingEdge r_edges(edge).surfs(0), edge

        If r_edges(edge).surfs(1) Then R_LeadingEdgeBackwards edge

        edge = r_edges(edge).nxt
    Loop

    R_CleanupSpan
End Sub

' Input:
' newedges[] array
'     this has links to edges, which have links to surfaces
'
' Output:
' Each surface has a linked list of its visible spans
Public Sub R_ScanEdges()
    Dim iv As Long, bottom As Long, s As Long

    If Not spans_made Then
        ReDim spans(0 To MAXSPANS - 1)
        spans_made = True
    End If

    max_span_p = MAXSPANS - r_refdef.vrect.width

    span_p = 0

    ' clear active edges to just the background edges around the whole screen
    ' FIXME: most of this only needs to be set up once
    With r_edges(EDGE_HEAD)
        .u = r_refdef.vrect.x * &H100000
        edge_head_u_shift20 = (.u And &HFFF00000) \ &H100000
        .u_step = 0
        .prev = -1
        .nxt = EDGE_TAIL
        .surfs(0) = 0
        .surfs(1) = 1
    End With

    With r_edges(EDGE_TAIL)
        .u = r_refdef.vrectright * &H100000 + &HFFFFF
        edge_tail_u_shift20 = (.u And &HFFF00000) \ &H100000
        .u_step = 0
        .prev = EDGE_HEAD
        .nxt = EDGE_AFTERTAIL
        .surfs(0) = 1
        .surfs(1) = 0
    End With

    With r_edges(EDGE_AFTERTAIL)
        .u = -1                          ' force a move
        .u_step = 0
        .nxt = EDGE_SENTINEL
        .prev = EDGE_TAIL
    End With

    ' FIXME: do we need this now that we clamp x in r_draw.c?
    r_edges(EDGE_SENTINEL).u = &HD0000000    ' make sure nothing sorts past this
                                             ' (C: 2000 << 24, which wraps to this in 32 bits)
    r_edges(EDGE_SENTINEL).prev = EDGE_AFTERTAIL

    '
    ' process all scan lines
    '
    bottom = r_refdef.vrectbottom - 1

    For iv = r_refdef.vrect.y To bottom - 1
        current_iv = iv
        fv = iv

        ' mark that the head (background start) span is pre-included
        surfaces(1).spanstate = 1

        If newedges(iv) <> -1 Then R_InsertNewEdges newedges(iv), r_edges(EDGE_HEAD).nxt

        If drawbackward Then
            R_GenerateSpansBackward
        Else
            R_GenerateSpans
        End If

        ' flush the span list if we can't be sure we have enough spans left for
        ' the next scan
        If span_p >= max_span_p Then
            VID_UnlockBuffer
            S_ExtraUpdate                ' don't let sound get messed up if going slow
            VID_LockBuffer

            If r_drawculledpolys Then
                R_DrawCulledPolys_
            Else
                D_DrawSurfaces
            End If

            ' clear the surface span pointers
            For s = 1 To surface_p - 1
                surfaces(s).spans = -1
            Next

            span_p = 0
        End If

        If removeedges(iv) <> -1 Then R_RemoveEdges removeedges(iv)

        If r_edges(EDGE_HEAD).nxt <> EDGE_TAIL Then R_StepActiveU r_edges(EDGE_HEAD).nxt
    Next

    ' do the last scan (no need to step or sort or remove on the last scan)

    current_iv = iv
    fv = iv

    ' mark that the head (background start) span is pre-included
    surfaces(1).spanstate = 1

    If newedges(iv) <> -1 Then R_InsertNewEdges newedges(iv), r_edges(EDGE_HEAD).nxt

    If drawbackward Then
        R_GenerateSpansBackward
    Else
        R_GenerateSpans
    End If

    ' draw whatever's left in the span list
    If r_drawculledpolys Then
        R_DrawCulledPolys_
    Else
        D_DrawSurfaces
    End If
End Sub
