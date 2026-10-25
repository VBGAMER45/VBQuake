Attribute VB_Name = "r_efrag"
' r_efrag.bas - entity fragments: the leafs a static entity is in (r_efrag.c)
'
' C keeps efrag_t ** pointers to the link being built; here lastlink says
' which link that is: 0 the entity's efrag field, k > 0 cl_efrags(k).entnext.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public r_pefragtopnode As Long           ' mod_nodes() index, -1 = NULL

'===========================================================================
' ENTITY FRAGMENT FUNCTIONS

Private lastlink As Long
Public r_emins As vec3_t, r_emaxs As vec3_t

Private r_addent As Long                 ' cl_entities() index

' (QuakeVB) *lastlink = ef
Private Sub SetLastLink(ByVal ef As Long)
    If lastlink = 0 Then
        cl_entities(r_addent).efrag = ef
    Else
        cl_efrags(lastlink).entnext = ef
    End If
End Sub

' Call when removing an object from the world or moving it to another position
Public Sub R_RemoveEfrags(ByVal ent As Long)
    Dim ef As Long, old As Long, walk As Long, prev As Long, leaf As Long

    ef = cl_entities(ent).efrag

    Do While ef <> 0
        ' (prev: 0 is the leaf's efrags field, k > 0 cl_efrags(k).leafnext)
        leaf = cl_efrags(ef).leaf
        prev = 0
        Do
            If prev = 0 Then
                walk = mod_nodes(leaf).efrags
            Else
                walk = cl_efrags(prev).leafnext
            End If
            If walk = 0 Then Exit Do
            If walk = ef Then
                ' remove this fragment
                If prev = 0 Then
                    mod_nodes(leaf).efrags = cl_efrags(ef).leafnext
                Else
                    cl_efrags(prev).leafnext = cl_efrags(ef).leafnext
                End If
                Exit Do
            Else
                prev = walk
            End If
        Loop

        old = ef
        ef = cl_efrags(ef).entnext

        ' put it on the free list
        cl_efrags(old).entnext = cl.free_efrags
        cl.free_efrags = old
    Loop

    cl_entities(ent).efrag = 0
End Sub

Private Sub R_SplitEntityOnNode(ByVal node As Long)
    Dim ef As Long, sides As Long

    If mod_nodes(node).contents = CONTENTS_SOLID Then Exit Sub

    ' add an efrag if the node is a leaf

    If mod_nodes(node).contents < 0 Then
        If r_pefragtopnode = -1 Then r_pefragtopnode = node

        ' grab an efrag off the free list
        ef = cl.free_efrags
        If ef = 0 Then
            Con_Printf "Too many efrags!\n"
            Exit Sub                     ' no free fragments...
        End If
        cl.free_efrags = cl_efrags(cl.free_efrags).entnext

        cl_efrags(ef).entity = r_addent

        ' add the entity link
        SetLastLink ef
        lastlink = ef
        cl_efrags(ef).entnext = 0

        ' set the leaf links
        cl_efrags(ef).leaf = node
        cl_efrags(ef).leafnext = mod_nodes(node).efrags
        mod_nodes(node).efrags = ef

        Exit Sub
    End If

    ' NODE_MIXED

    sides = BOX_ON_PLANE_SIDE(r_emins, r_emaxs, mod_planes(mod_nodes(node).plane))

    If sides = 3 Then
        ' split on this plane
        ' if this is the first splitter of this bmodel, remember it
        If r_pefragtopnode = -1 Then r_pefragtopnode = node
    End If

    ' recurse down the contacted sides
    If sides And 1 Then R_SplitEntityOnNode mod_nodes(node).children(0)

    If sides And 2 Then R_SplitEntityOnNode mod_nodes(node).children(1)
End Sub

Public Sub R_SplitEntityOnNode2(ByVal node As Long)
    Dim sides As Long

    If mod_nodes(node).visframe <> r_visframecount Then Exit Sub

    If mod_nodes(node).contents < 0 Then
        If mod_nodes(node).contents <> CONTENTS_SOLID Then r_pefragtopnode = node   ' we've reached a non-solid leaf, so it's
                                                                                   '  visible and not BSP clipped
        Exit Sub
    End If

    sides = BOX_ON_PLANE_SIDE(r_emins, r_emaxs, mod_planes(mod_nodes(node).plane))

    If sides = 3 Then
        ' remember first splitter
        r_pefragtopnode = node
        Exit Sub
    End If

    ' not split yet; recurse down the contacted side
    If sides And 1 Then
        R_SplitEntityOnNode2 mod_nodes(node).children(0)
    Else
        R_SplitEntityOnNode2 mod_nodes(node).children(1)
    End If
End Sub

Public Sub R_AddEfrags(ByVal ent As Long)
    Dim entmodel As Long, i As Long

    If cl_entities(ent).model = 0 Then Exit Sub

    If ent = 0 Then Exit Sub             ' never add the world

    r_addent = ent

    lastlink = 0                         ' &ent->efrag
    r_pefragtopnode = -1

    entmodel = cl_entities(ent).model
    For i = 0 To 2
        r_emins.v(i) = cl_entities(ent).origin.v(i) + mod_known(entmodel).mins.v(i)
        r_emaxs.v(i) = cl_entities(ent).origin.v(i) + mod_known(entmodel).maxs.v(i)
    Next

    R_SplitEntityOnNode mod_known(cl.worldmodel).nodes

    cl_entities(ent).topnode = r_pefragtopnode
End Sub

' FIXME: a lot of this goes away with edge-based
' (C: R_StoreEfrags (efrag_t **ppefrag), called with &leaf->efrags; it only
' reads the list, so it takes the first efrag)
Public Sub R_StoreEfrags(ByVal pefrag As Long)
    Dim pent As Long, clmodel As Long

    Do While pefrag <> 0
        pent = cl_efrags(pefrag).entity
        clmodel = cl_entities(pent).model

        Select Case mod_known(clmodel).typ
        Case mod_alias, mod_brush, mod_sprite
            pent = cl_efrags(pefrag).entity

            If cl_entities(pent).visframe <> r_framecount And cl_numvisedicts < MAX_VISEDICTS Then
                cl_visedicts(cl_numvisedicts) = pent
                cl_numvisedicts = cl_numvisedicts + 1

                ' mark that we've recorded this entity for this frame
                cl_entities(pent).visframe = r_framecount
            End If

            pefrag = cl_efrags(pefrag).leafnext

        Case Else
            Sys_Error "R_StoreEfrags: Bad entity type %d\n", mod_known(clmodel).typ
        End Select
    Loop
End Sub
