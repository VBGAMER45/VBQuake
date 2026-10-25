Attribute VB_Name = "r_main"
' r_main.bas - the refresh: sets up the view and draws the world and the
' entities (r_main.c)
'
' Not ported: the stack and alignment checks of R_RenderView, and
' Sys_LowFPPrecision / Sys_HighFPPrecision (x87 precision control for the
' asm drawers; VB's floating point stays as it is).
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' define PASSAGES

Public viewlightvec As vec3_t
Public r_viewlighting As alight_t        ' = {128, 192, viewlightvec}, set by R_Init
Public r_time1 As Single
Public r_numallocatededges As Long
Public r_drawpolys As Boolean
Public r_drawculledpolys As Boolean
Public r_worldpolysbacktofront As Boolean
Public r_recursiveaffinetriangles As Boolean   ' = true, set by D_Init
Public r_pixbytes As Long                ' = 1, set by D_Init
Public r_aliasuvscale As Single          ' = 1.0, set by D_Init
Public r_outofsurfaces As Long
Public r_outofedges As Long

Public r_dowarp As Boolean, r_dowarpold As Boolean, r_viewchanged_ As Boolean

Public numbtofpolys As Long
Public r_pcurrentvertbase As Long        ' mod_vertexes() index (C: mvertex_t *)

Public c_surf As Long
Public r_maxsurfsseen As Long, r_maxedgesseen As Long, r_cnumsurfs As Long
Public r_surfsonstack As Boolean
Public r_clipflags As Long

Public r_warpbuffer() As Byte            ' C: R_RenderView_'s local warpbuffer

Public r_fov_greater_than_90 As Boolean

'
' view origin
'
' (vup, vpn, vright and r_origin are in quakedef.bas)
Public base_vup As vec3_t
Public base_vpn As vec3_t
Public base_vright As vec3_t

'
' screen size info
'
' (r_refdef is in quakedef.bas)
Public xcenter As Single, ycenter As Single
Public xscale As Single, yscale As Single
Public xscaleinv As Single, yscaleinv As Single
Public xscaleshrink As Single, yscaleshrink As Single
Public aliasxscale As Single, aliasyscale As Single, aliasxcenter As Single, aliasycenter As Single

Public pixelAspect As Single
Public screenAspect As Single
Public verticalFieldOfView As Single
Public xOrigin As Single, yOrigin As Single

Public screenedge(0 To 3) As mplane_t

'
' refresh flags
'
Public r_framecount As Long              ' = 1 so frame counts initialized to 0 don't match (R_Init)
Public r_visframecount As Long
Public d_spanpixcount As Long
Public r_polycount As Long
Public r_drawnpolycount As Long
Public r_wholepolycount As Long

Private Const VIEWMODNAME_LENGTH As Long = 256
Public viewmodname As String
Public modcount As Long

Public pfrustum_indexes(0 To 3) As Long  ' r_frustum_indexes() offsets
Public r_frustum_indexes(0 To 4 * 6 - 1) As Long

Public reinit_surfcache As Long          ' = 1 (R_Init): if 1, surface cache is currently empty and
                                         ' must be reinitialized for current cache size

Public r_viewleaf As Long, r_oldviewleaf As Long   ' mod_nodes() indices, -1 = NULL

' (r_notexture_mip is in quakedef.bas)

Public r_aliastransition As Single, r_resfudge As Single

Public d_lightstylevalue(0 To 255) As Long   ' 8.8 fraction of base light value

Public dp_time1 As Single, dp_time2 As Single, db_time1 As Single, db_time2 As Single, rw_time1 As Single, rw_time2 As Single
Public se_time1 As Single, se_time2 As Single, de_time1 As Single, de_time2 As Single, dv_time1 As Single, dv_time2 As Single

Public cv_r_draworder As Long
Public cv_r_speeds As Long
Public cv_r_timegraph As Long
Public cv_r_graphheight As Long
Public cv_r_clearcolor As Long
Public cv_r_waterwarp As Long
Public cv_r_fullbright As Long
Public cv_r_drawentities As Long
Public cv_r_drawviewmodel As Long
Public cv_r_aliasstats As Long
Public cv_r_dspeeds As Long
Public cv_r_drawflat As Long
Public cv_r_ambient As Long
Public cv_r_reportsurfout As Long
Public cv_r_maxsurfs As Long
Public cv_r_numsurfs As Long
Public cv_r_reportedgeout As Long
Public cv_r_maxedges As Long
Public cv_r_numedges As Long
Public cv_r_aliastransbase As Long
Public cv_r_aliastransadj As Long

Public Sub R_InitTextures()
    Dim x As Long, y As Long, m As Long, dest As Long, pixofs As Long, sz As Long

    ' create a simple checkerboard texture for the default
    r_notexture_mip = Mod_AllocTexture(16 * 16 + 8 * 8 + 4 * 4 + 2 * 2, pixofs)

    With textures(r_notexture_mip)
        .name = "notexture"
        .width = 16
        .height = 16
        .offsets(0) = pixofs
        .offsets(1) = .offsets(0) + 16 * 16
        .offsets(2) = .offsets(1) + 8 * 8
        .offsets(3) = .offsets(2) + 4 * 4

        For m = 0 To 3
            dest = .offsets(m)
            sz = 16 \ BIT_(m)
            For y = 0 To sz - 1
                For x = 0 To sz - 1
                    If (y < sz \ 2) Xor (x < sz \ 2) Then
                        texpixels(dest) = 0
                    Else
                        texpixels(dest) = &HFF
                    End If
                    dest = dest + 1
                Next
            Next
        Next
    End With
End Sub

Public Sub R_Init()
    ' (C remembers the stack position here, for R_RenderView's check)
    r_framecount = 1
    reinit_surfcache = 1
    r_viewlighting.ambientlight = 128
    r_viewlighting.shadelight = 192
    r_viewleaf = -1
    r_oldviewleaf = -1
    ReDim r_warpbuffer(0 To WARP_WIDTH * WARP_HEIGHT - 1)

    R_InitTurb

    Cmd_AddCommand "timerefresh", XC_R_TimeRefresh_f
    Cmd_AddCommand "pointfile", XC_R_ReadPointFile_f

    cv_r_draworder = Cvar_RegisterVariable("r_draworder", "0")
    cv_r_speeds = Cvar_RegisterVariable("r_speeds", "0")
    cv_r_timegraph = Cvar_RegisterVariable("r_timegraph", "0")
    cv_r_graphheight = Cvar_RegisterVariable("r_graphheight", "10")
    cv_r_drawflat = Cvar_RegisterVariable("r_drawflat", "0")
    cv_r_ambient = Cvar_RegisterVariable("r_ambient", "0")
    cv_r_clearcolor = Cvar_RegisterVariable("r_clearcolor", "2")
    cv_r_waterwarp = Cvar_RegisterVariable("r_waterwarp", "1")
    cv_r_fullbright = Cvar_RegisterVariable("r_fullbright", "0")
    cv_r_drawentities = Cvar_RegisterVariable("r_drawentities", "1")
    cv_r_drawviewmodel = Cvar_RegisterVariable("r_drawviewmodel", "1")
    cv_r_aliasstats = Cvar_RegisterVariable("r_polymodelstats", "0")
    cv_r_dspeeds = Cvar_RegisterVariable("r_dspeeds", "0")
    cv_r_reportsurfout = Cvar_RegisterVariable("r_reportsurfout", "0")
    cv_r_maxsurfs = Cvar_RegisterVariable("r_maxsurfs", "0")
    cv_r_numsurfs = Cvar_RegisterVariable("r_numsurfs", "0")
    cv_r_reportedgeout = Cvar_RegisterVariable("r_reportedgeout", "0")
    cv_r_maxedges = Cvar_RegisterVariable("r_maxedges", "0")
    cv_r_numedges = Cvar_RegisterVariable("r_numedges", "0")
    cv_r_aliastransbase = Cvar_RegisterVariable("r_aliastransbase", "200")
    cv_r_aliastransadj = Cvar_RegisterVariable("r_aliastransadj", "100")

    Cvar_SetValue "r_maxedges", NUMSTACKEDGES
    Cvar_SetValue "r_maxsurfs", NUMSTACKSURFACES

    view_clipplanes(0).leftedge = 1
    view_clipplanes(1).rightedge = 1
    view_clipplanes(1).leftedge = 0
    view_clipplanes(2).leftedge = 0
    view_clipplanes(3).leftedge = 0
    view_clipplanes(0).rightedge = 0
    view_clipplanes(2).rightedge = 0
    view_clipplanes(3).rightedge = 0

    r_refdef.xOrigin = XCENTERING
    r_refdef.yOrigin = YCENTERING

    R_InitVertexNormals                  ' (QuakeVB) C's static r_avertexnormals

    R_InitParticles

    D_Init
End Sub

Public Sub R_NewMap()
    Dim i As Long, blanksurf As surf_t, blankedge As edge_t

    ' clear out efrags in case the level hasn't been reloaded
    ' FIXME: is this one short?
    For i = 0 To mod_known(cl.worldmodel).numleafs - 1
        mod_nodes(mod_known(cl.worldmodel).leafs + i).efrags = 0
    Next

    r_viewleaf = -1
    R_ClearParticles

    r_cnumsurfs = ToInt(cvars(cv_r_maxsurfs).value)

    If r_cnumsurfs <= MINSURFACES Then r_cnumsurfs = MINSURFACES

    ' (C puts NUMSTACKSURFACES on the stack and more on the hunk; both are
    ' indexed from surfaces[1], with surfaces[0] the dummy before them, so
    ' surf_max is one past surfaces[r_cnumsurfs])
    r_surfsonstack = r_cnumsurfs <= NUMSTACKSURFACES
    ReDim surfaces(0 To r_cnumsurfs)
    surf_max = r_cnumsurfs + 1
    surface_p = 1
    ' surface 0 doesn't really exist; it's just a dummy because index 0
    ' is used to indicate no edge attached to surface
    R_SurfacePatch

    r_maxedgesseen = 0
    r_maxsurfsseen = 0

    r_numallocatededges = ToInt(cvars(cv_r_maxedges).value)

    If r_numallocatededges < MINEDGES Then r_numallocatededges = MINEDGES

    ReDim r_edges(0 To EDGES_BASE + r_numallocatededges - 1)

    r_dowarpold = False
    r_viewchanged_ = False
End Sub

Public Sub R_SetVrect(pvrectin As vrect_t, pvrect As vrect_t, ByVal lineadj As Long)
    Dim h As Long, size As Single

    If cvars(cv_scr_viewsize).value > 100 Then size = 100 Else size = cvars(cv_scr_viewsize).value
    If cl.intermission Then
        size = 100
        lineadj = 0
    End If
    size = size / 100

    h = pvrectin.height - lineadj
    pvrect.width = ToInt(pvrectin.width * size)
    If pvrect.width < 96 Then
        size = 96# / pvrectin.width
        pvrect.width = 96                ' min for icons
    End If
    pvrect.width = pvrect.width And Not 7
    pvrect.height = ToInt(pvrectin.height * size)
    If pvrect.height > pvrectin.height - lineadj Then pvrect.height = pvrectin.height - lineadj

    pvrect.height = pvrect.height And Not 1

    pvrect.x = (pvrectin.width - pvrect.width) \ 2
    pvrect.y = (h - pvrect.height) \ 2

    If cvars(cv_lcd_x).value <> 0 Then
        pvrect.y = SAR(pvrect.y, 1)
        pvrect.height = SAR(pvrect.height, 1)
    End If
End Sub

' Called every time the vid structure or r_refdef changes.
' Guaranteed to be called before the first refresh
Public Sub R_ViewChanged(pvrect As vrect_t, ByVal lineadj As Long, ByVal aspect As Single)
    Dim i As Long, res_scale As Single

    r_viewchanged_ = True

    R_SetVrect pvrect, r_refdef.vrect, lineadj

    With r_refdef
        .horizontalFieldOfView = 2# * Tan(.fov_x / 360 * M_PI)
        .fvrectx = .vrect.x
        .fvrectx_adj = .vrect.x - 0.5
        .vrect_x_adj_shift20 = (.vrect.x * &H100000) + &H80000 - 1
        .fvrecty = .vrect.y
        .fvrecty_adj = .vrect.y - 0.5
        .vrectright = .vrect.x + .vrect.width
        .vrectright_adj_shift20 = (.vrectright * &H100000) + &H80000 - 1
        .fvrectright = .vrectright
        .fvrectright_adj = .vrectright - 0.5
        .vrectrightedge = .vrectright - 0.99
        .vrectbottom = .vrect.y + .vrect.height
        .fvrectbottom = .vrectbottom
        .fvrectbottom_adj = .vrectbottom - 0.5

        .aliasvrect.x = ToInt(.vrect.x * r_aliasuvscale)
        .aliasvrect.y = ToInt(.vrect.y * r_aliasuvscale)
        .aliasvrect.width = ToInt(.vrect.width * r_aliasuvscale)
        .aliasvrect.height = ToInt(.vrect.height * r_aliasuvscale)
        .aliasvrectright = .aliasvrect.x + .aliasvrect.width
        .aliasvrectbottom = .aliasvrect.y + .aliasvrect.height
    End With

    pixelAspect = aspect
    xOrigin = r_refdef.xOrigin
    yOrigin = r_refdef.yOrigin

    screenAspect = r_refdef.vrect.width * pixelAspect / r_refdef.vrect.height
    ' 320*200 1.0 pixelAspect = 1.6 screenAspect
    ' 320*240 1.0 pixelAspect = 1.3333 screenAspect
    ' proper 320*200 pixelAspect = 0.8333333

    verticalFieldOfView = r_refdef.horizontalFieldOfView / screenAspect

    ' values for perspective projection
    ' if math were exact, the values would range from 0.5 to to range+0.5
    ' hopefully they wll be in the 0.000001 to range+.999999 and truncate
    ' the polygon rasterization will never render in the first row or column
    ' but will definately render in the [range] row and column, so adjust the
    ' buffer origin to get an exact edge to edge fill
    xcenter = (CSng(r_refdef.vrect.width) * XCENTERING) + r_refdef.vrect.x - 0.5
    aliasxcenter = xcenter * r_aliasuvscale
    ycenter = (CSng(r_refdef.vrect.height) * YCENTERING) + r_refdef.vrect.y - 0.5
    aliasycenter = ycenter * r_aliasuvscale

    xscale = r_refdef.vrect.width / r_refdef.horizontalFieldOfView
    aliasxscale = xscale * r_aliasuvscale
    xscaleinv = 1# / xscale
    yscale = xscale * pixelAspect
    aliasyscale = yscale * r_aliasuvscale
    yscaleinv = 1# / yscale
    xscaleshrink = (r_refdef.vrect.width - 6) / r_refdef.horizontalFieldOfView
    yscaleshrink = xscaleshrink * pixelAspect

    ' left side clip
    screenedge(0).normal.v(0) = -1# / (xOrigin * r_refdef.horizontalFieldOfView)
    screenedge(0).normal.v(1) = 0
    screenedge(0).normal.v(2) = 1
    screenedge(0).typ = PLANE_ANYZ

    ' right side clip
    screenedge(1).normal.v(0) = 1# / ((1# - xOrigin) * r_refdef.horizontalFieldOfView)
    screenedge(1).normal.v(1) = 0
    screenedge(1).normal.v(2) = 1
    screenedge(1).typ = PLANE_ANYZ

    ' top side clip
    screenedge(2).normal.v(0) = 0
    screenedge(2).normal.v(1) = -1# / (yOrigin * verticalFieldOfView)
    screenedge(2).normal.v(2) = 1
    screenedge(2).typ = PLANE_ANYZ

    ' bottom side clip
    screenedge(3).normal.v(0) = 0
    screenedge(3).normal.v(1) = 1# / ((1# - yOrigin) * verticalFieldOfView)
    screenedge(3).normal.v(2) = 1
    screenedge(3).typ = PLANE_ANYZ

    For i = 0 To 3
        VectorNormalize screenedge(i).normal
    Next

    res_scale = Sqr(CDbl(r_refdef.vrect.width * r_refdef.vrect.height) / (320# * 152#)) * (2# / r_refdef.horizontalFieldOfView)
    r_aliastransition = cvars(cv_r_aliastransbase).value * res_scale
    r_resfudge = cvars(cv_r_aliastransadj).value * res_scale

    If cvars(cv_scr_fov).value <= 90# Then
        r_fov_greater_than_90 = False
    Else
        r_fov_greater_than_90 = True
    End If

    D_ViewChanged
End Sub

Private Sub R_MarkLeaves()
    Dim i As Long, node As Long, wm As Long

    If r_oldviewleaf = r_viewleaf Then Exit Sub

    r_visframecount = r_visframecount + 1
    r_oldviewleaf = r_viewleaf

    wm = cl.worldmodel
    Mod_LeafPVS r_viewleaf, wm           ' vis is mod_pvs()

    For i = 0 To mod_known(wm).numleafs - 1
        If mod_pvs(i \ 8) And BIT_(i And 7) Then
            node = mod_known(wm).leafs + i + 1
            Do
                If mod_nodes(node).visframe = r_visframecount Then Exit Do
                mod_nodes(node).visframe = r_visframecount
                node = mod_nodes(node).parent
            Loop While node <> -1
        End If
    Next
End Sub

' (QuakeVB) the dynamic lights' share of an entity's ambient light
Private Function DlightsAdd(ByVal ambient As Long, ent As vec3_t) As Long
    Dim lnum As Long, dist As vec3_t, add As Single

    For lnum = 0 To MAX_DLIGHTS - 1
        If cl_dlights(lnum).die >= cl.time Then
            VectorSubtract ent, cl_dlights(lnum).origin, dist
            add = cl_dlights(lnum).radius - Length(dist)

            If add > 0 Then ambient = ToInt(ambient + add)
        End If
    Next
    DlightsAdd = ambient
End Function

Private Sub R_DrawEntitiesOnList()
    Dim i As Long, j As Long, lighting As alight_t, lightvec As vec3_t

    ' FIXME: remove and do real lighting
    lightvec.v(0) = -1
    lightvec.v(1) = 0
    lightvec.v(2) = 0

    If cvars(cv_r_drawentities).value = 0 Then Exit Sub

    For i = 0 To cl_numvisedicts - 1
        currententity = cl_visedicts(i)

        If currententity = cl.viewentity Then GoTo nexti   ' don't draw the player

        Select Case mod_known(cl_entities(currententity).model).typ
        Case mod_sprite
            r_entorigin = cl_entities(currententity).origin
            VectorSubtract r_origin, r_entorigin, modelorg
            R_DrawSprite

        Case mod_alias
            r_entorigin = cl_entities(currententity).origin
            VectorSubtract r_origin, r_entorigin, modelorg

            ' see if the bounding box lets us trivially reject, also sets
            ' trivial accept status
            If R_AliasCheckBBox() Then
                j = R_LightPoint(cl_entities(currententity).origin)

                lighting.ambientlight = j
                lighting.shadelight = j

                lighting.plightvec = lightvec

                lighting.ambientlight = DlightsAdd(lighting.ambientlight, cl_entities(currententity).origin)

                ' clamp lighting so it doesn't overbright as much
                If lighting.ambientlight > 128 Then lighting.ambientlight = 128
                If lighting.ambientlight + lighting.shadelight > 192 Then lighting.shadelight = 192 - lighting.ambientlight

                R_AliasDrawModel lighting
            End If
        End Select
nexti:
    Next
End Sub

Private Sub R_DrawViewModel()
    Dim lightvec As vec3_t, j As Long, lnum As Long, dist As vec3_t, add As Single

    ' FIXME: remove and do real lighting
    lightvec.v(0) = -1
    lightvec.v(1) = 0
    lightvec.v(2) = 0

    If cvars(cv_r_drawviewmodel).value = 0 Or r_fov_greater_than_90 Then Exit Sub

    If cl.items And IT_INVISIBILITY Then Exit Sub

    If cl.stats(STAT_HEALTH) <= 0 Then Exit Sub

    currententity = CL_VIEWENT
    If cl_entities(currententity).model = 0 Then Exit Sub

    r_entorigin = cl_entities(currententity).origin
    VectorSubtract r_origin, r_entorigin, modelorg

    viewlightvec = vup
    VectorInverse viewlightvec

    j = R_LightPoint(cl_entities(currententity).origin)

    If j < 24 Then j = 24                ' allways give some light on gun
    r_viewlighting.ambientlight = j
    r_viewlighting.shadelight = j

    ' add dynamic lights
    For lnum = 0 To MAX_DLIGHTS - 1
        If cl_dlights(lnum).radius = 0 Then GoTo nextl
        If cl_dlights(lnum).radius = 0 Then GoTo nextl
        If cl_dlights(lnum).die < cl.time Then GoTo nextl

        VectorSubtract cl_entities(currententity).origin, cl_dlights(lnum).origin, dist
        add = cl_dlights(lnum).radius - Length(dist)
        If add > 0 Then r_viewlighting.ambientlight = ToInt(r_viewlighting.ambientlight + add)
nextl:
    Next

    ' clamp lighting so it doesn't overbright as much
    If r_viewlighting.ambientlight > 128 Then r_viewlighting.ambientlight = 128
    If r_viewlighting.ambientlight + r_viewlighting.shadelight > 192 Then r_viewlighting.shadelight = 192 - r_viewlighting.ambientlight

    r_viewlighting.plightvec = lightvec

    R_AliasDrawModel r_viewlighting
End Sub

Private Function R_BmodelCheckBBox(ByVal clmodel As Long, minmaxs() As Single) As Long
    Dim i As Long, pindex As Long, clipflags As Long, acceptpt As vec3_t, rejectpt As vec3_t, d As Double

    clipflags = 0

    With cl_entities(currententity)
        If .angles.v(0) <> 0 Or .angles.v(1) <> 0 Or .angles.v(2) <> 0 Then
            For i = 0 To 3
                d = DotProduct(.origin, view_clipplanes(i).normal)
                d = d - view_clipplanes(i).dist

                If d <= -mod_known(clmodel).radius Then
                    R_BmodelCheckBBox = BMODEL_FULLY_CLIPPED
                    Exit Function
                End If

                If d <= mod_known(clmodel).radius Then clipflags = clipflags Or BIT_(i)
            Next
        Else
            For i = 0 To 3
                ' generate accept and reject points
                ' FIXME: do with fast look-ups or integer tests based on the sign bit
                ' of the floating point values

                pindex = pfrustum_indexes(i)

                rejectpt.v(0) = minmaxs(r_frustum_indexes(pindex + 0))
                rejectpt.v(1) = minmaxs(r_frustum_indexes(pindex + 1))
                rejectpt.v(2) = minmaxs(r_frustum_indexes(pindex + 2))

                d = DotProduct(rejectpt, view_clipplanes(i).normal)
                d = d - view_clipplanes(i).dist

                If d <= 0 Then
                    R_BmodelCheckBBox = BMODEL_FULLY_CLIPPED
                    Exit Function
                End If

                acceptpt.v(0) = minmaxs(r_frustum_indexes(pindex + 3 + 0))
                acceptpt.v(1) = minmaxs(r_frustum_indexes(pindex + 3 + 1))
                acceptpt.v(2) = minmaxs(r_frustum_indexes(pindex + 3 + 2))

                d = DotProduct(acceptpt, view_clipplanes(i).normal)
                d = d - view_clipplanes(i).dist

                If d <= 0 Then clipflags = clipflags Or BIT_(i)
            Next
        End If
    End With

    R_BmodelCheckBBox = clipflags
End Function

Private Sub R_DrawBEntitiesOnList()
    Dim i As Long, j As Long, k As Long, clipflags As Long, oldorigin As vec3_t, clmodel As Long
    Dim minmaxs(0 To 5) As Single

    If cvars(cv_r_drawentities).value = 0 Then Exit Sub

    oldorigin = modelorg
    insubmodel = True
    r_dlightframecount = r_framecount

    For i = 0 To cl_numvisedicts - 1
        currententity = cl_visedicts(i)

        If mod_known(cl_entities(currententity).model).typ <> mod_brush Then GoTo nexti

        clmodel = cl_entities(currententity).model

        ' see if the bounding box lets us trivially reject, also sets
        ' trivial accept status
        For j = 0 To 2
            minmaxs(j) = cl_entities(currententity).origin.v(j) + mod_known(clmodel).mins.v(j)
            minmaxs(3 + j) = cl_entities(currententity).origin.v(j) + mod_known(clmodel).maxs.v(j)
        Next

        clipflags = R_BmodelCheckBBox(clmodel, minmaxs)

        If clipflags <> BMODEL_FULLY_CLIPPED Then
            r_entorigin = cl_entities(currententity).origin
            VectorSubtract r_origin, r_entorigin, modelorg
            ' FIXME: is this needed?
            r_worldmodelorg = modelorg

            r_pcurrentvertbase = mod_known(clmodel).vertexes

            ' FIXME: stop transforming twice
            R_RotateBmodel

            ' calculate dynamic lighting for bmodel if it's not an
            ' instanced model
            If mod_known(clmodel).firstmodelsurface <> 0 Then
                For k = 0 To MAX_DLIGHTS - 1
                    If cl_dlights(k).die < cl.time Or cl_dlights(k).radius = 0 Then GoTo nextk

                    R_MarkLights k, BIT_(k), mod_known(clmodel).nodes + mod_known(clmodel).hulls(0).firstclipnode
nextk:
                Next
            End If

            ' if the driver wants polygons, deliver those. Z-buffering is on
            ' at this point, so no clipping to the world tree is needed, just
            ' frustum clipping
            If r_drawpolys Or r_drawculledpolys Then
                R_ZDrawSubmodelPolys clmodel
            Else
                r_pefragtopnode = -1

                For j = 0 To 2
                    r_emins.v(j) = minmaxs(j)
                    r_emaxs.v(j) = minmaxs(3 + j)
                Next

                R_SplitEntityOnNode2 mod_known(cl.worldmodel).nodes

                If r_pefragtopnode <> -1 Then
                    cl_entities(currententity).topnode = r_pefragtopnode

                    If mod_nodes(r_pefragtopnode).contents >= 0 Then
                        ' not a leaf; has to be clipped to the world BSP
                        r_clipflags = clipflags
                        R_DrawSolidClippedSubmodelPolygons clmodel
                    Else
                        ' falls entirely in one leaf, so we just put all the
                        ' edges in the edge list and let 1/z sorting handle
                        ' drawing order
                        R_DrawSubmodelPolygons clmodel, clipflags
                    End If

                    cl_entities(currententity).topnode = -1
                End If
            End If

            ' put back world rotation and frustum clipping
            ' FIXME: R_RotateBmodel should just work off base_vxx
            vpn = base_vpn
            vup = base_vup
            vright = base_vright
            modelorg = base_modelorg
            modelorg = oldorigin
            R_TransformFrustum
        End If
nexti:
    Next

    insubmodel = False
End Sub

Private Sub R_EdgeDrawing()
    ' (C puts the edges and surfaces on the stack unless r_maxedges and
    ' r_maxsurfs ask for more; R_NewMap allocates them here)

    If r_surfsonstack Then R_SurfacePatch

    R_BeginEdgeFrame

    If cvars(cv_r_dspeeds).value <> 0 Then rw_time1 = Sys_FloatTime()

    R_RenderWorld

    If r_drawculledpolys Then R_ScanEdges

    ' only the world can be drawn back to front with no z reads or compares, just
    ' z writes, so have the driver turn z compares on now
    D_TurnZOn

    If cvars(cv_r_dspeeds).value <> 0 Then
        rw_time2 = Sys_FloatTime()
        db_time1 = rw_time2
    End If

    R_DrawBEntitiesOnList

    If cvars(cv_r_dspeeds).value <> 0 Then
        db_time2 = Sys_FloatTime()
        se_time1 = db_time2
    End If

    If cvars(cv_r_dspeeds).value = 0 Then
        VID_UnlockBuffer
        S_ExtraUpdate                    ' don't let sound get messed up if going slow
        VID_LockBuffer
    End If

    If Not (r_drawpolys Or r_drawculledpolys) Then R_ScanEdges
End Sub

' r_refdef must be set before the first call
Private Sub R_RenderView_()
    ' (C: byte warpbuffer[WARP_WIDTH * WARP_HEIGHT]; r_warpbuffer = warpbuffer;
    ' r_warpbuffer() is allocated by R_Init)

    If cvars(cv_r_timegraph).value <> 0 Or cvars(cv_r_speeds).value <> 0 Or cvars(cv_r_dspeeds).value <> 0 Then r_time1 = Sys_FloatTime()

    R_SetupFrame

    R_MarkLeaves                         ' done here so we know if we're in water

    ' make FDIV fast. This reduces timing precision after we've been running for a
    ' while, so we don't do it globally.  This also sets chop mode, and we do it
    ' here so that setup stuff like the refresh area calculations match what's
    ' done in screen.c
    ' Sys_LowFPPrecision ();  (not ported)

    If cl_entities(0).model = 0 Or cl.worldmodel = 0 Then Sys_Error "R_RenderView: NULL worldmodel"

    If cvars(cv_r_dspeeds).value = 0 Then
        VID_UnlockBuffer
        S_ExtraUpdate                    ' don't let sound get messed up if going slow
        VID_LockBuffer
    End If

    R_EdgeDrawing

    If cvars(cv_r_dspeeds).value = 0 Then
        VID_UnlockBuffer
        S_ExtraUpdate                    ' don't let sound get messed up if going slow
        VID_LockBuffer
    End If

    If cvars(cv_r_dspeeds).value <> 0 Then
        se_time2 = Sys_FloatTime()
        de_time1 = se_time2
    End If

    R_DrawEntitiesOnList

    If cvars(cv_r_dspeeds).value <> 0 Then
        de_time2 = Sys_FloatTime()
        dv_time1 = de_time2
    End If

    R_DrawViewModel

    If cvars(cv_r_dspeeds).value <> 0 Then
        dv_time2 = Sys_FloatTime()
        dp_time1 = Sys_FloatTime()
    End If

    R_DrawParticles

    If cvars(cv_r_dspeeds).value <> 0 Then dp_time2 = Sys_FloatTime()

    If r_dowarp Then D_WarpScreen

    V_SetContentsColor mod_nodes(r_viewleaf).contents

    If cvars(cv_r_timegraph).value <> 0 Then R_TimeGraph

    If cvars(cv_r_aliasstats).value <> 0 Then R_PrintAliasStats

    If cvars(cv_r_speeds).value <> 0 Then R_PrintTimes

    If cvars(cv_r_dspeeds).value <> 0 Then R_PrintDSpeeds

    If cvars(cv_r_reportsurfout).value <> 0 And r_outofsurfaces <> 0 Then Con_Printf "Short %d surfaces\n", r_outofsurfaces

    If cvars(cv_r_reportedgeout).value <> 0 And r_outofedges <> 0 Then Con_Printf "Short roughly %d edges\n", r_outofedges * 2 \ 3

    ' back to high floating-point precision
    ' Sys_HighFPPrecision ();  (not ported)
End Sub

Public Sub R_RenderView()
    ' (C checks here that there is enough stack, and that the hunk, the stack
    ' and the globals are aligned)
    R_RenderView_
End Sub

Public Sub R_InitTurb()
    Dim i As Long

    For i = 0 To SIN_BUFFER_SIZE - 1
        sintable(i) = ToInt(AMP + Sin(i * 3.14159 * 2 / CYCLE) * AMP)
        intsintable(i) = ToInt(AMP2 + Sin(i * 3.14159 * 2 / CYCLE) * AMP2)   ' AMP2, not 20
    Next
End Sub

' (QuakeVB) the console commands of the renderer, for Cmd_Call
Public Sub R_CallCommand(ByVal fn As Long)
    Select Case fn
    Case XC_R_TimeRefresh_f: R_TimeRefresh_f
    Case XC_R_ReadPointFile_f: R_ReadPointFile_f
    Case Else: Sys_Error "R_CallCommand: bad function %i", fn
    End Select
End Sub
