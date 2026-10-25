Attribute VB_Name = "world"
' world.bas - world query functions (world.c)
'
' entities never clip against themselves, or their owner
'
' line of sight checks trace->crosscontent, but bullets don't
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Type moveclip_t
    boxmins As vec3_t                    ' enclose the test object along entire move
    boxmaxs As vec3_t
    mins As vec3_t                       ' size of the moving object (C: pointers)
    maxs As vec3_t
    mins2 As vec3_t                      ' size when clipping against mosnters
    maxs2 As vec3_t
    start As vec3_t                      ' (C: pointers)
    end_ As vec3_t
    trace As trace_t
    typ As Long
    passedict As Long                    ' -1 = NULL
End Type

'===============================================================================
' HULL BOXES

Private box_hull As hull_t               ' its clipnodes and planes are in model.bas's pools

'===============================================================================
' ENTITY AREA CHECKING

Private Type areanode_t
    axis As Long                         ' -1 = leaf node
    dist As Single
    children(0 To 1) As Long             ' sv_areanodes() indices, -1 = NULL
    ' trigger_edicts: link AREA_LINKS + n * 2, solid_edicts: AREA_LINKS + n * 2 + 1
End Type

Private Const AREA_DEPTH As Long = 4
Private Const AREA_NODES As Long = 32
Private Const AREA_LINKS As Long = MAX_EDICTS    ' the first areanode link

Private sv_areanodes(0 To AREA_NODES - 1) As areanode_t
Private sv_numareanodes As Long
Private links_made As Boolean

' 1/32 epsilon to keep floating point happy
Private Const DIST_EPSILON As Single = 0.03125

Private Function TRIGGER_EDICTS(ByVal node As Long) As Long
    TRIGGER_EDICTS = AREA_LINKS + node * 2
End Function

Private Function SOLID_EDICTS(ByVal node As Long) As Long
    SOLID_EDICTS = AREA_LINKS + node * 2 + 1
End Function

' (QuakeVB) an edict's area link is NULL until it is linked (C: the edicts
' are zeroed when allocated)
Public Sub SV_ClearAreaLink(ByVal ent As Long)
    If Not links_made Then
        InitLinks AREA_LINKS + AREA_NODES * 2
        links_made = True
    End If
    link_prev(ent) = -1
    link_next(ent) = -1
End Sub

' Set up the planes and clipnodes so that the six floats of a bounding box
' can just be stored out and get a proper hull_t structure.
Private Sub SV_InitBoxHull()
    Dim i As Long, side As Long, c As Long, p As Long

    box_hull.clipnodes = Mod_AllocClipnodes(6)
    box_hull.planes = Mod_AllocPlanes(6)
    box_hull.firstclipnode = 0
    box_hull.lastclipnode = 5

    For i = 0 To 5
        c = box_hull.clipnodes + i
        p = box_hull.planes + i
        mod_clipnodes(c).planenum = i

        side = i And 1

        mod_clipnodes(c).children(side) = CONTENTS_EMPTY
        If i <> 5 Then
            mod_clipnodes(c).children(side Xor 1) = i + 1
        Else
            mod_clipnodes(c).children(side Xor 1) = CONTENTS_SOLID
        End If

        mod_planes(p).typ = i \ 2
        mod_planes(p).normal = vec3_origin
        mod_planes(p).normal.v(i \ 2) = 1
        mod_planes(p).dist = 0
        mod_planes(p).signbits = 0
    Next
End Sub

' To keep everything totally uniform, bounding boxes are turned into small
' BSP trees instead of being compared directly.
Private Function SV_HullForBox(mins As vec3_t, maxs As vec3_t) As hull_t
    mod_planes(box_hull.planes + 0).dist = maxs.v(0)
    mod_planes(box_hull.planes + 1).dist = mins.v(0)
    mod_planes(box_hull.planes + 2).dist = maxs.v(1)
    mod_planes(box_hull.planes + 3).dist = mins.v(1)
    mod_planes(box_hull.planes + 4).dist = maxs.v(2)
    mod_planes(box_hull.planes + 5).dist = mins.v(2)

    SV_HullForBox = box_hull
End Function

' Returns a hull that can be used for testing or clipping an object of mins/maxs
' size.
' Offset is filled in to contain the adjustment that must be added to the
' testing object's origin to get a point to use with the returned hull.
Private Function SV_HullForEntity(ByVal ent As Long, mins As vec3_t, maxs As vec3_t, offset As vec3_t) As hull_t
    Dim m As Long, size As vec3_t, hullmins As vec3_t, hullmaxs As vec3_t, hull As hull_t

    ' decide which clipping hull to use, based on the size
    If ev(ent).solid = SOLID_BSP Then
        ' explicit hulls in the BSP model
        If ev(ent).movetype <> MOVETYPE_PUSH Then Sys_Error "SOLID_BSP without MOVETYPE_PUSH"

        m = sv.models(ToInt(ev(ent).modelindex))

        If m = 0 Then Sys_Error "MOVETYPE_PUSH with a non bsp model"
        If mod_known(m).typ <> mod_brush Then Sys_Error "MOVETYPE_PUSH with a non bsp model"

        VectorSubtract maxs, mins, size
        If size.v(0) < 3 Then
            hull = mod_known(m).hulls(0)
        ElseIf size.v(0) <= 32 Then
            hull = mod_known(m).hulls(1)
        Else
            hull = mod_known(m).hulls(2)
        End If

        ' calculate an offset value to center the origin
        VectorSubtract hull.clip_mins, mins, offset
        VectorAdd offset, ev(ent).origin, offset
    Else
        ' create a temp hull from bounding box sizes
        VectorSubtract ev(ent).mins, maxs, hullmins
        VectorSubtract ev(ent).maxs, mins, hullmaxs
        hull = SV_HullForBox(hullmins, hullmaxs)

        offset = ev(ent).origin
    End If

    SV_HullForEntity = hull
End Function

Private Function SV_CreateAreaNode(ByVal depth As Long, mins As vec3_t, maxs As vec3_t) As Long
    Dim anode As Long, size As vec3_t, mins1 As vec3_t, maxs1 As vec3_t, mins2 As vec3_t, maxs2 As vec3_t
    Dim c As Long

    anode = sv_numareanodes
    sv_numareanodes = sv_numareanodes + 1

    ClearLink TRIGGER_EDICTS(anode)
    ClearLink SOLID_EDICTS(anode)

    If depth = AREA_DEPTH Then
        sv_areanodes(anode).axis = -1
        sv_areanodes(anode).children(0) = -1
        sv_areanodes(anode).children(1) = -1
        SV_CreateAreaNode = anode
        Exit Function
    End If

    VectorSubtract maxs, mins, size
    If size.v(0) > size.v(1) Then
        sv_areanodes(anode).axis = 0
    Else
        sv_areanodes(anode).axis = 1
    End If

    sv_areanodes(anode).dist = 0.5 * (maxs.v(sv_areanodes(anode).axis) + mins.v(sv_areanodes(anode).axis))
    mins1 = mins
    mins2 = mins
    maxs1 = maxs
    maxs2 = maxs

    maxs1.v(sv_areanodes(anode).axis) = sv_areanodes(anode).dist
    mins2.v(sv_areanodes(anode).axis) = sv_areanodes(anode).dist

    c = SV_CreateAreaNode(depth + 1, mins2, maxs2)
    sv_areanodes(anode).children(0) = c
    c = SV_CreateAreaNode(depth + 1, mins1, maxs1)
    sv_areanodes(anode).children(1) = c

    SV_CreateAreaNode = anode
End Function

' called after the world model has been loaded, before linking any entities
Public Sub SV_ClearWorld()
    Dim i As Long, blank As areanode_t

    SV_InitBoxHull

    For i = 0 To AREA_NODES - 1
        sv_areanodes(i) = blank
    Next
    sv_numareanodes = 0
    SV_CreateAreaNode 0, mod_known(sv.worldmodel).mins, mod_known(sv.worldmodel).maxs
End Sub

' call before removing an entity, and before trying to move one,
' so it doesn't clip against itself
' flags ent->v.modified
Public Sub SV_UnlinkEdict(ByVal ent As Long)
    If link_prev(ent) = -1 Then Exit Sub ' not linked in anywhere
    RemoveLink ent
    link_prev(ent) = -1
    link_next(ent) = -1
End Sub

Private Sub SV_TouchLinks(ByVal ent As Long, ByVal node As Long)
    Dim l As Long, nxt As Long, touch As Long, old_self As Long, old_other As Long, head As Long

    ' touch linked edicts
    head = TRIGGER_EDICTS(node)
    l = link_next(head)
    Do While l <> head
        nxt = link_next(l)
        touch = l                        ' EDICT_FROM_AREA(l)
        If touch = ent Then GoTo nextl
        If ev(touch).touch = 0 Or ev(touch).solid <> SOLID_TRIGGER Then GoTo nextl
        If ev(ent).absmin.v(0) > ev(touch).absmax.v(0) _
        Or ev(ent).absmin.v(1) > ev(touch).absmax.v(1) _
        Or ev(ent).absmin.v(2) > ev(touch).absmax.v(2) _
        Or ev(ent).absmax.v(0) < ev(touch).absmin.v(0) _
        Or ev(ent).absmax.v(1) < ev(touch).absmin.v(1) _
        Or ev(ent).absmax.v(2) < ev(touch).absmin.v(2) Then GoTo nextl

        old_self = pr_global_struct(0).self
        old_other = pr_global_struct(0).other

        pr_global_struct(0).self = EDICT_TO_PROG(touch)
        pr_global_struct(0).other = EDICT_TO_PROG(ent)
        pr_global_struct(0).time = sv.time
        PR_ExecuteProgram ev(touch).touch

        pr_global_struct(0).self = old_self
        pr_global_struct(0).other = old_other
nextl:
        l = nxt
    Loop

    ' recurse down both sides
    If sv_areanodes(node).axis = -1 Then Exit Sub

    If ev(ent).absmax.v(sv_areanodes(node).axis) > sv_areanodes(node).dist Then SV_TouchLinks ent, sv_areanodes(node).children(0)
    If ev(ent).absmin.v(sv_areanodes(node).axis) < sv_areanodes(node).dist Then SV_TouchLinks ent, sv_areanodes(node).children(1)
End Sub

' node: a mod_nodes() index
Private Sub SV_FindTouchedLeafs(ByVal ent As Long, ByVal node As Long)
    Dim splitplane As Long, sides As Long, leafnum As Long

    If mod_nodes(node).contents = CONTENTS_SOLID Then Exit Sub

    ' add an efrag if the node is a leaf
    If mod_nodes(node).contents < 0 Then
        If edicts(ent).num_leafs = MAX_ENT_LEAFS Then Exit Sub

        leafnum = node - mod_known(sv.worldmodel).leafs - 1

        edicts(ent).leafnums(edicts(ent).num_leafs) = leafnum
        edicts(ent).num_leafs = edicts(ent).num_leafs + 1
        Exit Sub
    End If

    ' NODE_MIXED
    splitplane = mod_nodes(node).plane
    sides = BOX_ON_PLANE_SIDE(ev(ent).absmin, ev(ent).absmax, mod_planes(splitplane))

    ' recurse down the contacted sides
    If sides And 1 Then SV_FindTouchedLeafs ent, mod_nodes(node).children(0)

    If sides And 2 Then SV_FindTouchedLeafs ent, mod_nodes(node).children(1)
End Sub

' Needs to be called any time an entity changes origin, mins, maxs, or solid
' flags ent->v.modified
' sets ent->v.absmin and ent->v.absmax
' if touchtriggers, calls prog functions for the intersected triggers
Public Sub SV_LinkEdict(ByVal ent As Long, ByVal touch_triggers As Boolean)
    Dim node As Long

    If link_prev(ent) <> -1 Then SV_UnlinkEdict ent  ' unlink from old position

    If ent = 0 Then Exit Sub             ' don't add the world

    If edicts(ent).free Then Exit Sub

    ' set the abs box
    VectorAdd ev(ent).origin, ev(ent).mins, ev(ent).absmin
    VectorAdd ev(ent).origin, ev(ent).maxs, ev(ent).absmax

    '
    ' to make items easier to pick up and allow them to be grabbed off
    ' of shelves, the abs sizes are expanded
    '
    If ToInt(ev(ent).flags) And FL_ITEM Then
        ev(ent).absmin.v(0) = ev(ent).absmin.v(0) - 15
        ev(ent).absmin.v(1) = ev(ent).absmin.v(1) - 15
        ev(ent).absmax.v(0) = ev(ent).absmax.v(0) + 15
        ev(ent).absmax.v(1) = ev(ent).absmax.v(1) + 15
    Else
        ' because movement is clipped an epsilon away from an actual edge,
        ' we must fully check even when bounding boxes don't quite touch
        ev(ent).absmin.v(0) = ev(ent).absmin.v(0) - 1
        ev(ent).absmin.v(1) = ev(ent).absmin.v(1) - 1
        ev(ent).absmin.v(2) = ev(ent).absmin.v(2) - 1
        ev(ent).absmax.v(0) = ev(ent).absmax.v(0) + 1
        ev(ent).absmax.v(1) = ev(ent).absmax.v(1) + 1
        ev(ent).absmax.v(2) = ev(ent).absmax.v(2) + 1
    End If

    ' link to PVS leafs
    edicts(ent).num_leafs = 0
    If ev(ent).modelindex Then SV_FindTouchedLeafs ent, mod_known(sv.worldmodel).nodes

    If ev(ent).solid = SOLID_NOT Then Exit Sub

    ' find the first node that the ent's box crosses
    node = 0
    Do
        If sv_areanodes(node).axis = -1 Then Exit Do
        If ev(ent).absmin.v(sv_areanodes(node).axis) > sv_areanodes(node).dist Then
            node = sv_areanodes(node).children(0)
        ElseIf ev(ent).absmax.v(sv_areanodes(node).axis) < sv_areanodes(node).dist Then
            node = sv_areanodes(node).children(1)
        Else
            Exit Do                      ' crosses the node
        End If
    Loop

    ' link it in

    If ev(ent).solid = SOLID_TRIGGER Then
        InsertLinkBefore ent, TRIGGER_EDICTS(node)
    Else
        InsertLinkBefore ent, SOLID_EDICTS(node)
    End If

    ' if touch_triggers, touch all entities at this node and decend for more
    If touch_triggers Then SV_TouchLinks ent, 0
End Sub

'===============================================================================
' POINT TESTING IN HULLS

Public Function SV_HullPointContents(hull As hull_t, ByVal num As Long, p As vec3_t) As Long
    Dim d As Single, node As Long, plane As Long

    Do While num >= 0
        If num < hull.firstclipnode Or num > hull.lastclipnode Then Sys_Error "SV_HullPointContents: bad node number"

        node = hull.clipnodes + num
        plane = hull.planes + mod_clipnodes(node).planenum

        If mod_planes(plane).typ < 3 Then
            d = p.v(mod_planes(plane).typ) - mod_planes(plane).dist
        Else
            d = DotProduct(mod_planes(plane).normal, p) - mod_planes(plane).dist
        End If
        If d < 0 Then
            num = mod_clipnodes(node).children(1)
        Else
            num = mod_clipnodes(node).children(0)
        End If
    Loop

    SV_HullPointContents = num
End Function

' returns the CONTENTS_* value from the world at the given point.
' does not check any entities at all
' the non-true version remaps the water current contents to content_water
Public Function SV_PointContents(p As vec3_t) As Long
    Dim cont As Long

    cont = SV_HullPointContents(mod_known(sv.worldmodel).hulls(0), 0, p)
    If cont <= CONTENTS_CURRENT_0 And cont >= CONTENTS_CURRENT_DOWN Then cont = CONTENTS_WATER
    SV_PointContents = cont
End Function

Public Function SV_TruePointContents(p As vec3_t) As Long
    SV_TruePointContents = SV_HullPointContents(mod_known(sv.worldmodel).hulls(0), 0, p)
End Function

'===========================================================================

' This could be a lot more efficient...
' returns the world (0) if the entity is stuck, -1 (NULL) if not
Public Function SV_TestEntityPosition(ByVal ent As Long) As Long
    Dim trace As trace_t

    trace = SV_Move(ev(ent).origin, ev(ent).mins, ev(ent).maxs, ev(ent).origin, 0, ent)

    If trace.startsolid Then
        SV_TestEntityPosition = 0
    Else
        SV_TestEntityPosition = -1
    End If
End Function

'===============================================================================
' LINE TESTING IN HULLS

Public Function SV_RecursiveHullCheck(hull As hull_t, ByVal num As Long, ByVal p1f As Single, ByVal p2f As Single, _
        p1 As vec3_t, p2 As vec3_t, trace As trace_t) As Boolean
    Dim node As Long, plane As Long, t1 As Single, t2 As Single, frac As Single
    Dim i As Long, mid As vec3_t, side As Long, midf As Single

    ' check for empty
    If num < 0 Then
        If num <> CONTENTS_SOLID Then
            trace.allsolid = False
            If num = CONTENTS_EMPTY Then
                trace.inopen = True
            Else
                trace.inwater = True
            End If
        Else
            trace.startsolid = True
        End If
        SV_RecursiveHullCheck = True     ' empty
        Exit Function
    End If

    If num < hull.firstclipnode Or num > hull.lastclipnode Then Sys_Error "SV_RecursiveHullCheck: bad node number"

    '
    ' find the point distances
    '
    node = hull.clipnodes + num
    plane = hull.planes + mod_clipnodes(node).planenum

    If mod_planes(plane).typ < 3 Then
        t1 = p1.v(mod_planes(plane).typ) - mod_planes(plane).dist
        t2 = p2.v(mod_planes(plane).typ) - mod_planes(plane).dist
    Else
        t1 = DotProduct(mod_planes(plane).normal, p1) - mod_planes(plane).dist
        t2 = DotProduct(mod_planes(plane).normal, p2) - mod_planes(plane).dist
    End If

    If t1 >= 0 And t2 >= 0 Then
        SV_RecursiveHullCheck = SV_RecursiveHullCheck(hull, mod_clipnodes(node).children(0), p1f, p2f, p1, p2, trace)
        Exit Function
    End If
    If t1 < 0 And t2 < 0 Then
        SV_RecursiveHullCheck = SV_RecursiveHullCheck(hull, mod_clipnodes(node).children(1), p1f, p2f, p1, p2, trace)
        Exit Function
    End If

    ' put the crosspoint DIST_EPSILON pixels on the near side
    If t1 < 0 Then
        frac = (t1 + DIST_EPSILON) / (t1 - t2)
    Else
        frac = (t1 - DIST_EPSILON) / (t1 - t2)
    End If
    If frac < 0 Then frac = 0
    If frac > 1 Then frac = 1

    midf = p1f + (p2f - p1f) * frac
    For i = 0 To 2
        mid.v(i) = p1.v(i) + frac * (p2.v(i) - p1.v(i))
    Next

    side = -(t1 < 0)

    ' move up to the node
    If Not SV_RecursiveHullCheck(hull, mod_clipnodes(node).children(side), p1f, midf, p1, mid, trace) Then
        SV_RecursiveHullCheck = False
        Exit Function
    End If

    If SV_HullPointContents(hull, mod_clipnodes(node).children(side Xor 1), mid) <> CONTENTS_SOLID Then
        ' go past the node
        SV_RecursiveHullCheck = SV_RecursiveHullCheck(hull, mod_clipnodes(node).children(side Xor 1), midf, p2f, mid, p2, trace)
        Exit Function
    End If

    If trace.allsolid Then
        SV_RecursiveHullCheck = False    ' never got out of the solid area
        Exit Function
    End If

    '==================
    ' the other side of the node is solid, this is the impact point
    '==================
    If side = 0 Then
        trace.plane.normal = mod_planes(plane).normal
        trace.plane.dist = mod_planes(plane).dist
    Else
        VectorSubtract vec3_origin, mod_planes(plane).normal, trace.plane.normal
        trace.plane.dist = -mod_planes(plane).dist
    End If

    Do While SV_HullPointContents(hull, hull.firstclipnode, mid) = CONTENTS_SOLID
        ' shouldn't really happen, but does occasionally
        frac = frac - 0.1
        If frac < 0 Then
            trace.fraction = midf
            trace.endpos = mid
            Con_DPrintf "backup past 0\n"
            SV_RecursiveHullCheck = False
            Exit Function
        End If
        midf = p1f + (p2f - p1f) * frac
        For i = 0 To 2
            mid.v(i) = p1.v(i) + frac * (p2.v(i) - p1.v(i))
        Next
    Loop

    trace.fraction = midf
    trace.endpos = mid

    SV_RecursiveHullCheck = False
End Function

' Handles selection or creation of a clipping hull, and offseting (and
' eventually rotation) of the end points
Private Function SV_ClipMoveToEntity(ByVal ent As Long, start As vec3_t, mins As vec3_t, maxs As vec3_t, end_ As vec3_t) As trace_t
    Dim trace As trace_t, offset As vec3_t, start_l As vec3_t, end_l As vec3_t, hull As hull_t

    ' fill in a default trace
    trace.fraction = 1
    trace.allsolid = True
    trace.ent = -1
    trace.endpos = end_

    ' get the clipping hull
    hull = SV_HullForEntity(ent, mins, maxs, offset)

    VectorSubtract start, offset, start_l
    VectorSubtract end_, offset, end_l

    ' trace a line through the apropriate clipping hull
    SV_RecursiveHullCheck hull, hull.firstclipnode, 0, 1, start_l, end_l, trace

    ' fix trace up by the offset
    If trace.fraction <> 1 Then VectorAdd trace.endpos, offset, trace.endpos

    ' did we clip the move?
    If trace.fraction < 1 Or trace.startsolid Then trace.ent = ent

    SV_ClipMoveToEntity = trace
End Function

'===========================================================================

' Mins and maxs enclose the entire area swept by the move
Private Sub SV_ClipToLinks(ByVal node As Long, clip As moveclip_t)
    Dim l As Long, nxt As Long, touch As Long, trace As trace_t, head As Long

    ' touch linked edicts
    head = SOLID_EDICTS(node)
    l = link_next(head)
    Do While l <> head
        nxt = link_next(l)
        touch = l                        ' EDICT_FROM_AREA(l)
        If ev(touch).solid = SOLID_NOT Then GoTo nextl
        If touch = clip.passedict Then GoTo nextl
        If ev(touch).solid = SOLID_TRIGGER Then Sys_Error "Trigger in clipping list"

        If clip.typ = MOVE_NOMONSTERS And ev(touch).solid <> SOLID_BSP Then GoTo nextl

        If clip.boxmins.v(0) > ev(touch).absmax.v(0) _
        Or clip.boxmins.v(1) > ev(touch).absmax.v(1) _
        Or clip.boxmins.v(2) > ev(touch).absmax.v(2) _
        Or clip.boxmaxs.v(0) < ev(touch).absmin.v(0) _
        Or clip.boxmaxs.v(1) < ev(touch).absmin.v(1) _
        Or clip.boxmaxs.v(2) < ev(touch).absmin.v(2) Then GoTo nextl

        If clip.passedict <> -1 Then
            If ev(clip.passedict).size.v(0) <> 0 And ev(touch).size.v(0) = 0 Then GoTo nextl  ' points never interact
        End If

        ' might intersect, so do an exact clip
        If clip.trace.allsolid Then Exit Sub
        If clip.passedict <> -1 Then
            If PROG_TO_EDICT(ev(touch).owner) = clip.passedict Then GoTo nextl  ' don't clip against own missiles
            If PROG_TO_EDICT(ev(clip.passedict).owner) = touch Then GoTo nextl  ' don't clip against owner
        End If

        If ToInt(ev(touch).flags) And FL_MONSTER Then
            trace = SV_ClipMoveToEntity(touch, clip.start, clip.mins2, clip.maxs2, clip.end_)
        Else
            trace = SV_ClipMoveToEntity(touch, clip.start, clip.mins, clip.maxs, clip.end_)
        End If
        If trace.allsolid Or trace.startsolid Or trace.fraction < clip.trace.fraction Then
            trace.ent = touch
            If clip.trace.startsolid Then
                clip.trace = trace
                clip.trace.startsolid = True
            Else
                clip.trace = trace
            End If
        ElseIf trace.startsolid Then
            clip.trace.startsolid = True
        End If
nextl:
        l = nxt
    Loop

    ' recurse down both sides
    If sv_areanodes(node).axis = -1 Then Exit Sub

    If clip.boxmaxs.v(sv_areanodes(node).axis) > sv_areanodes(node).dist Then SV_ClipToLinks sv_areanodes(node).children(0), clip
    If clip.boxmins.v(sv_areanodes(node).axis) < sv_areanodes(node).dist Then SV_ClipToLinks sv_areanodes(node).children(1), clip
End Sub

Private Sub SV_MoveBounds(start As vec3_t, mins As vec3_t, maxs As vec3_t, end_ As vec3_t, boxmins As vec3_t, boxmaxs As vec3_t)
    Dim i As Long

    For i = 0 To 2
        If end_.v(i) > start.v(i) Then
            boxmins.v(i) = start.v(i) + mins.v(i) - 1
            boxmaxs.v(i) = end_.v(i) + maxs.v(i) + 1
        Else
            boxmins.v(i) = end_.v(i) + mins.v(i) - 1
            boxmaxs.v(i) = start.v(i) + maxs.v(i) + 1
        End If
    Next
End Sub

' mins and maxs are reletive
'
' if the entire move stays in a solid volume, trace.allsolid will be set
'
' if the starting point is in a solid, it will be allowed to move out
' to an open area
'
' nomonsters is used for line of sight or edge testing, where mosnters
' shouldn't be considered solid objects
'
' passedict is explicitly excluded from clipping checks (-1 = NULL)
Public Function SV_Move(start As vec3_t, mins As vec3_t, maxs As vec3_t, end_ As vec3_t, ByVal typ As Long, ByVal passedict As Long) As trace_t
    Dim clip As moveclip_t, i As Long

    ' clip to world
    clip.trace = SV_ClipMoveToEntity(0, start, mins, maxs, end_)

    clip.start = start
    clip.end_ = end_
    clip.mins = mins
    clip.maxs = maxs
    clip.typ = typ
    clip.passedict = passedict

    If typ = MOVE_MISSILE Then
        For i = 0 To 2
            clip.mins2.v(i) = -15
            clip.maxs2.v(i) = 15
        Next
    Else
        clip.mins2 = mins
        clip.maxs2 = maxs
    End If

    ' create the bounding box of the entire move
    SV_MoveBounds clip.start, clip.mins2, clip.maxs2, clip.end_, clip.boxmins, clip.boxmaxs

    ' clip to entities
    SV_ClipToLinks 0, clip

    SV_Move = clip.trace
End Function
