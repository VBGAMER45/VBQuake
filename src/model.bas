Attribute VB_Name = "model"
' model.bas - model loading and caching (model.c).
' models are the only shared resource between a client and server running
' on the same machine.
'
' A model_t* is an index into mod_known(1 .. MAX_MOD_KNOWN), 0 = NULL.
'
' The data of all loaded brush models and sprites is in the pools below; a
' model_t holds the index of its first element in each (C: hunk memory and
' pointers). C frees the hunk back to host_hunklevel at every level change
' (Host_ClearMemory); Mod_HunkFreeToMark resets these pools the same way.
' Alias models are cached by C (they survive level changes); their pools
' are never reset.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const MAX_MOD_KNOWN As Long = 256

' values for model_t's needload
Private Const NL_PRESENT As Long = 0
Private Const NL_NEEDS_LOADED As Long = 1
Private Const NL_UNREFERENCED As Long = 2

Private Const IDPOLYHEADER As Long = &H4F504449     ' little-endian "IDPO"
Private Const IDSPRITEHEADER As Long = &H50534449   ' little-endian "IDSP"
Private Const ANIM_CYCLE As Long = 2

Public mod_known() As model_t
Private mod_numknown As Long

Private loadmodel As Long
Private loadname As String               ' for hunk tags

Public mod_novis(0 To MAX_MAP_LEAFS \ 8 - 1) As Byte
Public mod_pvs(0 To MAX_MAP_LEAFS \ 8 - 1) As Byte  ' (QuakeVB) Mod_LeafPVS's result (C: returns a pointer)

Private mod_base() As Byte               ' the file being loaded

' brush model and sprite pools ("the hunk")
Public mod_vertexes() As mvertex_t
Public mod_edges() As medge_t
Public mod_surfedges() As Long
Public mod_texlist() As Long             ' C: texture_t **textures, entries are textures() indices
Public textures() As texture_t           ' (1 ..), 0 = NULL
Public texpixels() As Byte
Public mod_lightdata() As Byte
Public mod_visdata() As Byte
Public mod_entdata() As Byte
Public mod_planes() As mplane_t
Public mod_texinfo() As mtexinfo_t
Public mod_surfaces() As msurface_t
Public mod_marksurfaces() As Long        ' C: msurface_t **, entries are mod_surfaces() indices
Public mod_nodes() As mnode_t            ' nodes and leafs
Public mod_clipnodes() As dclipnode_t
Public mod_submodels() As dmodel_t
Public msprites() As msprite_t           ' (1 ..)
Public sprite_frames() As mspriteframe_t ' (1 ..)
Public sprite_groups() As mspritegroup_t ' (1 ..)
Public sprite_intervals() As Single
Public sprite_pixels() As Byte

' how much of each pool is in use (and the mark that Mod_HunkFreeToMark returns to)
Private Const NUM_POOLS As Long = 21
Private Const P_VERTEXES As Long = 0
Private Const P_EDGES As Long = 1
Private Const P_SURFEDGES As Long = 2
Private Const P_TEXLIST As Long = 3
Private Const P_TEXTURES As Long = 4
Private Const P_TEXPIXELS As Long = 5
Private Const P_LIGHTDATA As Long = 6
Private Const P_VISDATA As Long = 7
Private Const P_ENTDATA As Long = 8
Private Const P_PLANES As Long = 9
Private Const P_TEXINFO As Long = 10
Private Const P_SURFACES As Long = 11
Private Const P_MARKSURFACES As Long = 12
Private Const P_NODES As Long = 13
Private Const P_CLIPNODES As Long = 14
Private Const P_SUBMODELS As Long = 15
Private Const P_SPRITES As Long = 16
Private Const P_SPRITEFRAMES As Long = 17
Private Const P_SPRITEGROUPS As Long = 18
Private Const P_SPRITEINTERVALS As Long = 19
Private Const P_SPRITEPIXELS As Long = 20
Private pool_used(0 To NUM_POOLS - 1) As Long
Private pool_mark(0 To NUM_POOLS - 1) As Long

' alias model pools (the cache: never freed)
Public aliashdrs() As aliashdr_t         ' (1 ..)
Public alias_stverts() As stvert_t
Public alias_tris() As mtriangle_t
Public alias_verts() As trivertx_t
Public alias_skins() As Byte
Public alias_skindescs() As maliasskindesc_t
Public alias_skingroups() As maliasskingroup_t   ' (1 ..)
Public alias_groups() As maliasgroup_t   ' (1 ..)
Public alias_intervals() As Single
Private num_aliashdrs As Long
Private num_alias_stverts As Long
Private num_alias_tris As Long
Private num_alias_verts As Long
Private num_alias_skins As Long
Private num_alias_skindescs As Long
Private num_alias_skingroups As Long
Private num_alias_groups As Long
Private num_alias_intervals As Long

'=============================================================================
' pools

' C: Hunk_AllocName: returns the index of n new elements of pool p
Private Function PoolAlloc(ByVal p As Long, ByVal n As Long) As Long
    Dim first As Long, last As Long

    first = pool_used(p)
    If p = P_TEXTURES Or p = P_SPRITES Or p = P_SPRITEFRAMES Or p = P_SPRITEGROUPS Then
        first = first + 1                ' 1-based: 0 is NULL
    End If
    pool_used(p) = pool_used(p) + n
    last = first + n                     ' one spare, so n = 0 works

    Select Case p
    Case P_VERTEXES: ReDim Preserve mod_vertexes(0 To last)
    Case P_EDGES: ReDim Preserve mod_edges(0 To last)
    Case P_SURFEDGES: ReDim Preserve mod_surfedges(0 To last)
    Case P_TEXLIST: ReDim Preserve mod_texlist(0 To last)
    Case P_TEXTURES: ReDim Preserve textures(0 To last)
    Case P_TEXPIXELS: ReDim Preserve texpixels(0 To last)
    Case P_LIGHTDATA: ReDim Preserve mod_lightdata(0 To last)
    Case P_VISDATA: ReDim Preserve mod_visdata(0 To last)
    Case P_ENTDATA: ReDim Preserve mod_entdata(0 To last)
    Case P_PLANES: ReDim Preserve mod_planes(0 To last)
    Case P_TEXINFO: ReDim Preserve mod_texinfo(0 To last)
    Case P_SURFACES: ReDim Preserve mod_surfaces(0 To last)
    Case P_MARKSURFACES: ReDim Preserve mod_marksurfaces(0 To last)
    Case P_NODES: ReDim Preserve mod_nodes(0 To last)
    Case P_CLIPNODES: ReDim Preserve mod_clipnodes(0 To last)
    Case P_SUBMODELS: ReDim Preserve mod_submodels(0 To last)
    Case P_SPRITES: ReDim Preserve msprites(0 To last)
    Case P_SPRITEFRAMES: ReDim Preserve sprite_frames(0 To last)
    Case P_SPRITEGROUPS: ReDim Preserve sprite_groups(0 To last)
    Case P_SPRITEINTERVALS: ReDim Preserve sprite_intervals(0 To last)
    Case P_SPRITEPIXELS: ReDim Preserve sprite_pixels(0 To last)
    End Select
    PoolAlloc = first
End Function

Public Function Mod_AllocPlanes(ByVal n As Long) As Long
    Mod_AllocPlanes = PoolAlloc(P_PLANES, n)
End Function

Public Function Mod_AllocClipnodes(ByVal n As Long) As Long
    Mod_AllocClipnodes = PoolAlloc(P_CLIPNODES, n)
End Function

Public Function Mod_AllocTexture(ByVal pixels As Long, pixofs As Long) As Long
    Mod_AllocTexture = PoolAlloc(P_TEXTURES, 1)
    pixofs = PoolAlloc(P_TEXPIXELS, pixels)
End Function

' C: host_hunklevel = Hunk_LowMark () at the end of Host_Init
Public Sub Mod_HunkMark()
    Dim i As Long
    For i = 0 To NUM_POOLS - 1
        pool_mark(i) = pool_used(i)
    Next
End Sub

' C: Hunk_FreeToLowMark (host_hunklevel)
Public Sub Mod_HunkFreeToMark()
    Dim i As Long
    For i = 0 To NUM_POOLS - 1
        pool_used(i) = pool_mark(i)
    Next
End Sub

'=============================================================================

Public Sub Mod_Init()
    ReDim mod_known(1 To MAX_MOD_KNOWN)
    FillMemory mod_novis(0), MAX_MAP_LEAFS \ 8, &HFF
End Sub

' Caches the data if needed
Public Function Mod_Extradata(ByVal m As Long) As Long
    If mod_known(m).cache Then
        Mod_Extradata = mod_known(m).cache
        Exit Function
    End If

    Mod_LoadModel m, True

    If mod_known(m).cache = 0 Then Sys_Error "Mod_Extradata: caching failed"
    Mod_Extradata = mod_known(m).cache
End Function

' returns a mod_nodes() index (a leaf)
Public Function Mod_PointInLeaf(p As vec3_t, ByVal m As Long) As Long
    Dim node As Long, d As Single, plane As Long

    If m = 0 Then Sys_Error "Mod_PointInLeaf: bad model"
    If mod_known(m).numnodes = 0 Then Sys_Error "Mod_PointInLeaf: bad model"

    node = mod_known(m).nodes
    Do
        If mod_nodes(node).contents < 0 Then
            Mod_PointInLeaf = node
            Exit Function
        End If
        plane = mod_nodes(node).plane
        d = DotProduct(p, mod_planes(plane).normal) - mod_planes(plane).dist
        If d > 0 Then
            node = mod_nodes(node).children(0)
        Else
            node = mod_nodes(node).children(1)
        End If
    Loop
End Function

' in: an offset in mod_visdata(), -1 = NULL. The result is in mod_pvs().
Private Sub Mod_DecompressVis(ByVal in_ As Long, ByVal m As Long)
    Dim c As Long, out As Long, row As Long

    row = (mod_known(m).numleafs + 7) \ 8
    out = 0

    If in_ = -1 Then
        ' no vis info, so make all visible
        Do While row
            mod_pvs(out) = &HFF
            out = out + 1
            row = row - 1
        Loop
        Exit Sub
    End If

    Do
        If mod_visdata(in_) Then
            mod_pvs(out) = mod_visdata(in_)
            out = out + 1
            in_ = in_ + 1
        Else
            c = mod_visdata(in_ + 1)
            in_ = in_ + 2
            Do While c
                mod_pvs(out) = 0
                out = out + 1
                c = c - 1
            Loop
        End If
    Loop While out < row
End Sub

' leaf: a mod_nodes() index. The pvs is in mod_pvs() afterwards.
Public Sub Mod_LeafPVS(ByVal leaf As Long, ByVal m As Long)
    If leaf = mod_known(m).leafs Then
        CopyMemory mod_pvs(0), mod_novis(0), MAX_MAP_LEAFS \ 8
        Exit Sub
    End If
    Mod_DecompressVis mod_nodes(leaf).compressed_vis, m
End Sub

Public Sub Mod_ClearAll()
    Dim i As Long
    For i = 1 To mod_numknown
        mod_known(i).needload = NL_UNREFERENCED
        ' FIX FOR CACHE_ALLOC ERRORS:
        If mod_known(i).typ = mod_sprite Then mod_known(i).cache = 0
    Next
End Sub

Public Function Mod_FindName(ByVal name As String) As Long
    Dim i As Long, avail As Long, m As Long

    If Len(name) = 0 Then Sys_Error "Mod_ForName: NULL name"

    '
    ' search the currently loaded models
    '
    For i = 1 To mod_numknown
        If mod_known(i).name = name Then Exit For
        If mod_known(i).needload = NL_UNREFERENCED Then
            If avail = 0 Or mod_known(i).typ <> mod_alias Then avail = i
        End If
    Next

    m = i
    If i = mod_numknown + 1 Then
        If mod_numknown = MAX_MOD_KNOWN Then
            If avail Then
                m = avail
                ' (an alias model in the cache is simply dropped here)
                If mod_known(m).typ = mod_alias Then mod_known(m).cache = 0
            Else
                Sys_Error "mod_numknown == MAX_MOD_KNOWN"
            End If
        Else
            mod_numknown = mod_numknown + 1
        End If
        mod_known(m).name = name
        mod_known(m).needload = NL_NEEDS_LOADED
    End If

    Mod_FindName = m
End Function

Public Sub Mod_TouchModel(ByVal name As String)
    Mod_FindName name
    ' (C touches the cache entry of a loaded alias model)
End Sub

' Loads a model into the cache
Public Function Mod_LoadModel(ByVal m As Long, ByVal crash As Boolean) As Long
    If mod_known(m).typ = mod_alias Then
        If mod_known(m).cache Then
            mod_known(m).needload = NL_PRESENT
            Mod_LoadModel = m
            Exit Function
        End If
    Else
        If mod_known(m).needload = NL_PRESENT Then
            Mod_LoadModel = m
            Exit Function
        End If
    End If

    '
    ' because the world is so huge, load it one piece at a time
    '

    '
    ' load the file
    '
    If Not COM_LoadFile(mod_known(m).name, mod_base) Then
        If crash Then Sys_Error "Mod_NumForName: %s not found", mod_known(m).name
        Mod_LoadModel = 0
        Exit Function
    End If

    '
    ' allocate a new model
    '
    loadname = COM_FileBase(mod_known(m).name)

    loadmodel = m

    '
    ' fill it in
    '

    ' call the apropriate loader
    mod_known(m).needload = NL_PRESENT

    Select Case GetLong(mod_base, 0)
    Case IDPOLYHEADER
        Mod_LoadAliasModel m
    Case IDSPRITEHEADER
        Mod_LoadSpriteModel m
    Case Else
        Mod_LoadBrushModel m
    End Select

    Erase mod_base
    Mod_LoadModel = m
End Function

' Loads in a model for the given name
Public Function Mod_ForName(ByVal name As String, ByVal crash As Boolean) As Long
    Dim m As Long
    m = Mod_FindName(name)
    Mod_ForName = Mod_LoadModel(m, crash)
End Function

'===============================================================================
' BRUSHMODEL LOADING

Private Function LumpOfs(ByVal l As Long) As Long
    LumpOfs = GetLong(mod_base, 4 + l * 8)
End Function

Private Function LumpLen(ByVal l As Long) As Long
    LumpLen = GetLong(mod_base, 4 + l * 8 + 4)
End Function

Private Sub FunnyLump(ByVal l As Long, ByVal size As Long)
    If LumpLen(l) Mod size Then Sys_Error "MOD_LoadBmodel: funny lump size in %s", mod_known(loadmodel).name
End Sub

' 1 if the strings of the textures' names from character 2 on are equal
' (C: strcmp (tx2->name+2, tx->name+2))
Private Function TexNameChar(ByVal t As Long, ByVal i As Long) As Long
    If Len(textures(t).name) > i Then TexNameChar = AscW(Mid$(textures(t).name, i + 1, 1))
End Function

Private Sub Mod_LoadTextures(ByVal l As Long)
    Dim i As Long, j As Long, pixels As Long, num As Long, max As Long, altmax As Long
    Dim mt As Long, tx As Long, tx2 As Long, m As Long, nummiptex As Long
    Dim dataofs As Long, pixofs As Long, mtw As Long, mth As Long
    Dim anims(0 To 9) As Long, altanims(0 To 9) As Long

    If LumpLen(l) = 0 Then
        mod_known(loadmodel).textures = -1
        Exit Sub
    End If
    m = LumpOfs(l)
    nummiptex = GetLong(mod_base, m)

    mod_known(loadmodel).numtextures = nummiptex
    mod_known(loadmodel).textures = PoolAlloc(P_TEXLIST, nummiptex)

    For i = 0 To nummiptex - 1
        dataofs = GetLong(mod_base, m + 4 + i * 4)
        mod_texlist(mod_known(loadmodel).textures + i) = 0
        If dataofs <> -1 Then
            mt = m + dataofs
            mtw = GetLong(mod_base, mt + 16)
            mth = GetLong(mod_base, mt + 20)

            If (mtw And 15) Or (mth And 15) Then Sys_Error "Texture %s is not 16 aligned", BufToStrN(mod_base, mt, 16)
            pixels = (mtw * mth \ 64) * 85   ' (VB's \ binds looser than *: C's w*h/64*85)
            tx = Mod_AllocTexture(pixels, pixofs)
            mod_texlist(mod_known(loadmodel).textures + i) = tx

            textures(tx).name = BufToStrN(mod_base, mt, 16)
            textures(tx).width = mtw
            textures(tx).height = mth
            textures(tx).anim_total = 0
            textures(tx).anim_min = 0
            textures(tx).anim_max = 0
            textures(tx).anim_next = 0
            textures(tx).alternate_anims = 0
            ' the pixels immediately follow the structures (C: the miptex_t
            ' offsets, minus its 40 bytes, plus the texture_t's size)
            For j = 0 To MIPLEVELS - 1
                textures(tx).offsets(j) = pixofs + GetLong(mod_base, mt + 24 + j * 4) - 40
            Next
            If pixels > 0 Then CopyMemory texpixels(pixofs), mod_base(mt + 40), pixels

            If Left$(textures(tx).name, 3) = "sky" Then R_InitSky tx
        End If
    Next

    '
    ' sequence the animations
    '
    For i = 0 To nummiptex - 1
        tx = mod_texlist(mod_known(loadmodel).textures + i)
        If tx = 0 Then GoTo nexttex
        If Left$(textures(tx).name, 1) <> "+" Then GoTo nexttex
        If textures(tx).anim_next Then GoTo nexttex  ' allready sequenced

        ' find the number of frames in the animation
        For j = 0 To 9
            anims(j) = 0
            altanims(j) = 0
        Next

        max = TexNameChar(tx, 1)
        altmax = 0
        If max >= 97 And max <= 122 Then max = max - (97 - 65)
        If max >= 48 And max <= 57 Then
            max = max - 48
            altmax = 0
            anims(max) = tx
            max = max + 1
        ElseIf max >= 65 And max <= 74 Then
            altmax = max - 65
            max = 0
            altanims(altmax) = tx
            altmax = altmax + 1
        Else
            Sys_Error "Bad animating texture %s", textures(tx).name
        End If

        For j = i + 1 To nummiptex - 1
            tx2 = mod_texlist(mod_known(loadmodel).textures + j)
            If tx2 = 0 Then GoTo nextj
            If Left$(textures(tx2).name, 1) <> "+" Then GoTo nextj
            If Mid$(textures(tx2).name, 3) <> Mid$(textures(tx).name, 3) Then GoTo nextj

            num = TexNameChar(tx2, 1)
            If num >= 97 And num <= 122 Then num = num - (97 - 65)
            If num >= 48 And num <= 57 Then
                num = num - 48
                anims(num) = tx2
                If num + 1 > max Then max = num + 1
            ElseIf num >= 65 And num <= 74 Then
                num = num - 65
                altanims(num) = tx2
                If num + 1 > altmax Then altmax = num + 1
            Else
                Sys_Error "Bad animating texture %s", textures(tx).name
            End If
nextj:
        Next

        ' link them all together
        For j = 0 To max - 1
            tx2 = anims(j)
            If tx2 = 0 Then Sys_Error "Missing frame %i of %s", j, textures(tx).name
            textures(tx2).anim_total = max * ANIM_CYCLE
            textures(tx2).anim_min = j * ANIM_CYCLE
            textures(tx2).anim_max = (j + 1) * ANIM_CYCLE
            textures(tx2).anim_next = anims((j + 1) Mod max)
            If altmax Then textures(tx2).alternate_anims = altanims(0)
        Next
        For j = 0 To altmax - 1
            tx2 = altanims(j)
            If tx2 = 0 Then Sys_Error "Missing frame %i of %s", j, textures(tx).name
            textures(tx2).anim_total = altmax * ANIM_CYCLE
            textures(tx2).anim_min = j * ANIM_CYCLE
            textures(tx2).anim_max = (j + 1) * ANIM_CYCLE
            textures(tx2).anim_next = altanims((j + 1) Mod altmax)
            If max Then textures(tx2).alternate_anims = anims(0)
        Next
nexttex:
    Next
End Sub

Private Sub Mod_LoadLighting(ByVal l As Long)
    Dim p As Long
    If LumpLen(l) = 0 Then
        mod_known(loadmodel).lightdata = -1
        Exit Sub
    End If
    p = PoolAlloc(P_LIGHTDATA, LumpLen(l))
    mod_known(loadmodel).lightdata = p
    CopyMemory mod_lightdata(p), mod_base(LumpOfs(l)), LumpLen(l)
End Sub

Private Sub Mod_LoadVisibility(ByVal l As Long)
    Dim p As Long
    If LumpLen(l) = 0 Then
        mod_known(loadmodel).visdata = -1
        Exit Sub
    End If
    p = PoolAlloc(P_VISDATA, LumpLen(l))
    mod_known(loadmodel).visdata = p
    CopyMemory mod_visdata(p), mod_base(LumpOfs(l)), LumpLen(l)
End Sub

Private Sub Mod_LoadEntities(ByVal l As Long)
    Dim p As Long
    If LumpLen(l) = 0 Then
        mod_known(loadmodel).entities = -1
        Exit Sub
    End If
    ' (one more byte, so the text is zero-terminated for COM_Parse)
    p = PoolAlloc(P_ENTDATA, LumpLen(l) + 1)
    mod_known(loadmodel).entities = p
    CopyMemory mod_entdata(p), mod_base(LumpOfs(l)), LumpLen(l)
    mod_entdata(p + LumpLen(l)) = 0
End Sub

Private Sub Mod_LoadVertexes(ByVal l As Long)
    Dim count As Long, out As Long

    FunnyLump l, 12
    count = LumpLen(l) \ 12
    out = PoolAlloc(P_VERTEXES, count)

    mod_known(loadmodel).vertexes = out
    mod_known(loadmodel).numvertexes = count

    ' (mvertex_t has the layout of dvertex_t: three floats)
    If count > 0 Then CopyMemory mod_vertexes(out), mod_base(LumpOfs(l)), count * 12
End Sub

Private Sub Mod_LoadSubmodels(ByVal l As Long)
    Dim i As Long, j As Long, count As Long, out As Long, in_ As Long

    FunnyLump l, 64
    count = LumpLen(l) \ 64
    out = PoolAlloc(P_SUBMODELS, count)

    mod_known(loadmodel).submodels = out
    mod_known(loadmodel).numsubmodels = count

    in_ = LumpOfs(l)
    For i = 0 To count - 1
        For j = 0 To 2
            ' spread the mins / maxs by a pixel
            mod_submodels(out + i).mins.v(j) = GetFloat(mod_base, in_ + j * 4) - 1
            mod_submodels(out + i).maxs.v(j) = GetFloat(mod_base, in_ + 12 + j * 4) + 1
            mod_submodels(out + i).origin.v(j) = GetFloat(mod_base, in_ + 24 + j * 4)
        Next
        For j = 0 To MAX_MAP_HULLS - 1
            mod_submodels(out + i).headnode(j) = GetLong(mod_base, in_ + 36 + j * 4)
        Next
        mod_submodels(out + i).visleafs = GetLong(mod_base, in_ + 52)
        mod_submodels(out + i).firstface = GetLong(mod_base, in_ + 56)
        mod_submodels(out + i).numfaces = GetLong(mod_base, in_ + 60)
        in_ = in_ + 64
    Next
End Sub

Private Sub Mod_LoadEdges(ByVal l As Long)
    Dim i As Long, count As Long, out As Long, in_ As Long

    FunnyLump l, 4
    count = LumpLen(l) \ 4
    out = PoolAlloc(P_EDGES, count + 1)

    mod_known(loadmodel).edges = out
    mod_known(loadmodel).numedges = count

    in_ = LumpOfs(l)
    For i = 0 To count - 1
        mod_edges(out + i).v(0) = GetShort(mod_base, in_) And &HFFFF&
        mod_edges(out + i).v(1) = GetShort(mod_base, in_ + 2) And &HFFFF&
        mod_edges(out + i).cachededgeoffset = 0
        in_ = in_ + 4
    Next
End Sub

Private Sub Mod_LoadTexinfo(ByVal l As Long)
    Dim i As Long, j As Long, count As Long, miptex As Long, out As Long, in_ As Long
    Dim len1 As Single, len2 As Single, tx As Long

    FunnyLump l, 40
    count = LumpLen(l) \ 40
    out = PoolAlloc(P_TEXINFO, count)

    mod_known(loadmodel).texinfo = out
    mod_known(loadmodel).numtexinfo = count

    in_ = LumpOfs(l)
    For i = 0 To count - 1
        For j = 0 To 7
            mod_texinfo(out + i).vecs(j \ 4, j And 3) = GetFloat(mod_base, in_ + j * 4)
        Next
        len1 = Sqr(mod_texinfo(out + i).vecs(0, 0) * mod_texinfo(out + i).vecs(0, 0) + _
                   mod_texinfo(out + i).vecs(0, 1) * mod_texinfo(out + i).vecs(0, 1) + _
                   mod_texinfo(out + i).vecs(0, 2) * mod_texinfo(out + i).vecs(0, 2))
        len2 = Sqr(mod_texinfo(out + i).vecs(1, 0) * mod_texinfo(out + i).vecs(1, 0) + _
                   mod_texinfo(out + i).vecs(1, 1) * mod_texinfo(out + i).vecs(1, 1) + _
                   mod_texinfo(out + i).vecs(1, 2) * mod_texinfo(out + i).vecs(1, 2))
        len1 = (len1 + len2) / 2
        If len1 < 0.32 Then
            mod_texinfo(out + i).mipadjust = 4
        ElseIf len1 < 0.49 Then
            mod_texinfo(out + i).mipadjust = 3
        ElseIf len1 < 0.99 Then
            mod_texinfo(out + i).mipadjust = 2
        Else
            mod_texinfo(out + i).mipadjust = 1
        End If

        miptex = GetLong(mod_base, in_ + 32)
        mod_texinfo(out + i).flags = GetLong(mod_base, in_ + 36)

        If mod_known(loadmodel).textures = -1 Then
            mod_texinfo(out + i).texture = r_notexture_mip   ' checkerboard texture
            mod_texinfo(out + i).flags = 0
        Else
            If miptex >= mod_known(loadmodel).numtextures Then Sys_Error "miptex >= loadmodel->numtextures"
            tx = mod_texlist(mod_known(loadmodel).textures + miptex)
            mod_texinfo(out + i).texture = tx
            If tx = 0 Then
                mod_texinfo(out + i).texture = r_notexture_mip   ' texture not found
                mod_texinfo(out + i).flags = 0
            End If
        End If
        in_ = in_ + 40
    Next
End Sub

' Fills in s->texturemins[] and s->extents[]
Private Sub CalcSurfaceExtents(ByVal s As Long)
    Dim mins(0 To 1) As Single, maxs(0 To 1) As Single, val As Single
    Dim i As Long, j As Long, e As Long, v As Long, tex As Long
    Dim bmins(0 To 1) As Long, bmaxs(0 To 1) As Long

    mins(0) = 999999: mins(1) = 999999
    maxs(0) = -99999: maxs(1) = -99999

    tex = mod_surfaces(s).texinfo

    For i = 0 To mod_surfaces(s).numedges - 1
        e = mod_surfedges(mod_known(loadmodel).surfedges + mod_surfaces(s).firstedge + i)
        If e >= 0 Then
            v = mod_known(loadmodel).vertexes + mod_edges(mod_known(loadmodel).edges + e).v(0)
        Else
            v = mod_known(loadmodel).vertexes + mod_edges(mod_known(loadmodel).edges - e).v(1)
        End If

        For j = 0 To 1
            val = mod_vertexes(v).position.v(0) * mod_texinfo(tex).vecs(j, 0) + _
                  mod_vertexes(v).position.v(1) * mod_texinfo(tex).vecs(j, 1) + _
                  mod_vertexes(v).position.v(2) * mod_texinfo(tex).vecs(j, 2) + _
                  mod_texinfo(tex).vecs(j, 3)
            If val < mins(j) Then mins(j) = val
            If val > maxs(j) Then maxs(j) = val
        Next
    Next

    For i = 0 To 1
        bmins(i) = Int(mins(i) / 16)
        bmaxs(i) = -Int(-maxs(i) / 16)    ' ceil

        mod_surfaces(s).texturemins(i) = bmins(i) * 16
        mod_surfaces(s).extents(i) = (bmaxs(i) - bmins(i)) * 16
        If (mod_texinfo(tex).flags And TEX_SPECIAL) = 0 And mod_surfaces(s).extents(i) > 256 Then
            Sys_Error "Bad surface extents"
        End If
    Next
End Sub

Private Sub Mod_LoadFaces(ByVal l As Long)
    Dim i As Long, count As Long, surfnum As Long, planenum As Long, side As Long
    Dim out As Long, in_ As Long, s As Long, tname As String

    FunnyLump l, 20
    count = LumpLen(l) \ 20
    out = PoolAlloc(P_SURFACES, count)

    mod_known(loadmodel).surfaces = out
    mod_known(loadmodel).numsurfaces = count

    in_ = LumpOfs(l)
    For surfnum = 0 To count - 1
        s = out + surfnum
        mod_surfaces(s).firstedge = GetLong(mod_base, in_ + 4)
        mod_surfaces(s).numedges = GetShort(mod_base, in_ + 8)
        mod_surfaces(s).flags = 0
        mod_surfaces(s).visframe = 0
        mod_surfaces(s).dlightframe = 0
        mod_surfaces(s).dlightbits = 0
        For i = 0 To MIPLEVELS - 1
            mod_surfaces(s).cachespots(i) = 0
        Next

        planenum = GetShort(mod_base, in_)
        side = GetShort(mod_base, in_ + 2)
        If side Then mod_surfaces(s).flags = mod_surfaces(s).flags Or SURF_PLANEBACK

        mod_surfaces(s).plane = mod_known(loadmodel).planes + planenum

        mod_surfaces(s).texinfo = mod_known(loadmodel).texinfo + GetShort(mod_base, in_ + 10)

        CalcSurfaceExtents s

        ' lighting info
        For i = 0 To MAXLIGHTMAPS - 1
            mod_surfaces(s).styles(i) = mod_base(in_ + 12 + i)
        Next
        i = GetLong(mod_base, in_ + 16)
        If i = -1 Or mod_known(loadmodel).lightdata = -1 Then
            mod_surfaces(s).samples = -1
        Else
            mod_surfaces(s).samples = mod_known(loadmodel).lightdata + i
        End If

        ' set the drawing flags flag
        tname = textures(mod_texinfo(mod_surfaces(s).texinfo).texture).name
        If Left$(tname, 3) = "sky" Then
            ' sky
            mod_surfaces(s).flags = mod_surfaces(s).flags Or (SURF_DRAWSKY Or SURF_DRAWTILED)
        ElseIf Left$(tname, 1) = "*" Then
            ' turbulent
            mod_surfaces(s).flags = mod_surfaces(s).flags Or (SURF_DRAWTURB Or SURF_DRAWTILED)
            For i = 0 To 1
                mod_surfaces(s).extents(i) = 16384
                mod_surfaces(s).texturemins(i) = -8192
            Next
        End If
        in_ = in_ + 20
    Next
End Sub

Private Sub Mod_SetParent(ByVal node As Long, ByVal parent As Long)
    mod_nodes(node).parent = parent
    If mod_nodes(node).contents < 0 Then Exit Sub
    Mod_SetParent mod_nodes(node).children(0), node
    Mod_SetParent mod_nodes(node).children(1), node
End Sub

Private Sub Mod_LoadNodes(ByVal l As Long)
    Dim i As Long, j As Long, count As Long, p As Long, out As Long, in_ As Long, n As Long

    FunnyLump l, 24
    count = LumpLen(l) \ 24
    out = PoolAlloc(P_NODES, count)

    mod_known(loadmodel).nodes = out
    mod_known(loadmodel).numnodes = count

    in_ = LumpOfs(l)
    For i = 0 To count - 1
        n = out + i
        For j = 0 To 2
            mod_nodes(n).minmaxs(j) = GetShort(mod_base, in_ + 8 + j * 2)
            mod_nodes(n).minmaxs(3 + j) = GetShort(mod_base, in_ + 14 + j * 2)
        Next

        p = GetLong(mod_base, in_)
        mod_nodes(n).plane = mod_known(loadmodel).planes + p

        mod_nodes(n).firstsurface = GetShort(mod_base, in_ + 20) And &HFFFF&
        mod_nodes(n).numsurfaces = GetShort(mod_base, in_ + 22) And &HFFFF&
        mod_nodes(n).contents = 0
        mod_nodes(n).visframe = 0
        mod_nodes(n).compressed_vis = -1
        mod_nodes(n).efrags = 0

        For j = 0 To 1
            p = GetShort(mod_base, in_ + 4 + j * 2)
            If p >= 0 Then
                mod_nodes(n).children(j) = mod_known(loadmodel).nodes + p
            Else
                mod_nodes(n).children(j) = mod_known(loadmodel).leafs + (-1 - p)
            End If
        Next
        in_ = in_ + 24
    Next

    Mod_SetParent mod_known(loadmodel).nodes, -1     ' sets nodes and leafs
End Sub

Private Sub Mod_LoadLeafs(ByVal l As Long)
    Dim i As Long, j As Long, count As Long, p As Long, out As Long, in_ As Long, n As Long

    FunnyLump l, 28
    count = LumpLen(l) \ 28
    out = PoolAlloc(P_NODES, count)

    mod_known(loadmodel).leafs = out
    mod_known(loadmodel).numleafs = count

    in_ = LumpOfs(l)
    For i = 0 To count - 1
        n = out + i
        For j = 0 To 2
            mod_nodes(n).minmaxs(j) = GetShort(mod_base, in_ + 8 + j * 2)
            mod_nodes(n).minmaxs(3 + j) = GetShort(mod_base, in_ + 14 + j * 2)
        Next

        p = GetLong(mod_base, in_)
        mod_nodes(n).contents = p
        mod_nodes(n).visframe = 0

        mod_nodes(n).firstmarksurface = mod_known(loadmodel).marksurfaces + GetShort(mod_base, in_ + 20)
        mod_nodes(n).nummarksurfaces = GetShort(mod_base, in_ + 22) And &HFFFF&

        p = GetLong(mod_base, in_ + 4)
        If p = -1 Then
            mod_nodes(n).compressed_vis = -1
        Else
            mod_nodes(n).compressed_vis = mod_known(loadmodel).visdata + p
        End If
        mod_nodes(n).efrags = 0

        For j = 0 To 3
            mod_nodes(n).ambient_sound_level(j) = mod_base(in_ + 24 + j)
        Next
        in_ = in_ + 28
    Next
End Sub

Private Sub Mod_LoadClipnodes(ByVal l As Long)
    Dim i As Long, count As Long, out As Long, in_ As Long

    FunnyLump l, 8
    count = LumpLen(l) \ 8
    out = PoolAlloc(P_CLIPNODES, count)

    mod_known(loadmodel).clipnodes = out
    mod_known(loadmodel).numclipnodes = count

    With mod_known(loadmodel).hulls(1)
        .clipnodes = out
        .firstclipnode = 0
        .lastclipnode = count - 1
        .planes = mod_known(loadmodel).planes
        .clip_mins.v(0) = -16
        .clip_mins.v(1) = -16
        .clip_mins.v(2) = -24
        .clip_maxs.v(0) = 16
        .clip_maxs.v(1) = 16
        .clip_maxs.v(2) = 32
    End With

    With mod_known(loadmodel).hulls(2)
        .clipnodes = out
        .firstclipnode = 0
        .lastclipnode = count - 1
        .planes = mod_known(loadmodel).planes
        .clip_mins.v(0) = -32
        .clip_mins.v(1) = -32
        .clip_mins.v(2) = -24
        .clip_maxs.v(0) = 32
        .clip_maxs.v(1) = 32
        .clip_maxs.v(2) = 64
    End With

    in_ = LumpOfs(l)
    For i = 0 To count - 1
        mod_clipnodes(out + i).planenum = GetLong(mod_base, in_)
        mod_clipnodes(out + i).children(0) = GetShort(mod_base, in_ + 4)
        mod_clipnodes(out + i).children(1) = GetShort(mod_base, in_ + 6)
        in_ = in_ + 8
    Next
End Sub

' Deplicate the drawing hull structure as a clipping hull
Private Sub Mod_MakeHull0()
    Dim in_ As Long, child As Long, out As Long, i As Long, j As Long, count As Long

    in_ = mod_known(loadmodel).nodes
    count = mod_known(loadmodel).numnodes
    out = PoolAlloc(P_CLIPNODES, count)

    With mod_known(loadmodel).hulls(0)
        .clipnodes = out
        .firstclipnode = 0
        .lastclipnode = count - 1
        .planes = mod_known(loadmodel).planes
    End With

    For i = 0 To count - 1
        mod_clipnodes(out + i).planenum = mod_nodes(in_ + i).plane - mod_known(loadmodel).planes
        For j = 0 To 1
            child = mod_nodes(in_ + i).children(j)
            If mod_nodes(child).contents < 0 Then
                mod_clipnodes(out + i).children(j) = mod_nodes(child).contents
            Else
                mod_clipnodes(out + i).children(j) = child - mod_known(loadmodel).nodes
            End If
        Next
    Next
End Sub

Private Sub Mod_LoadMarksurfaces(ByVal l As Long)
    Dim i As Long, j As Long, count As Long, out As Long, in_ As Long

    FunnyLump l, 2
    count = LumpLen(l) \ 2
    out = PoolAlloc(P_MARKSURFACES, count)

    mod_known(loadmodel).marksurfaces = out
    mod_known(loadmodel).nummarksurfaces = count

    in_ = LumpOfs(l)
    For i = 0 To count - 1
        j = GetShort(mod_base, in_ + i * 2)
        If j >= mod_known(loadmodel).numsurfaces Then Sys_Error "Mod_ParseMarksurfaces: bad surface number"
        mod_marksurfaces(out + i) = mod_known(loadmodel).surfaces + j
    Next
End Sub

Private Sub Mod_LoadSurfedges(ByVal l As Long)
    Dim count As Long, out As Long

    FunnyLump l, 4
    count = LumpLen(l) \ 4
    out = PoolAlloc(P_SURFEDGES, count)

    mod_known(loadmodel).surfedges = out
    mod_known(loadmodel).numsurfedges = count

    If count > 0 Then CopyMemory mod_surfedges(out), mod_base(LumpOfs(l)), count * 4
End Sub

Private Sub Mod_LoadPlanes(ByVal l As Long)
    Dim i As Long, j As Long, count As Long, bits As Long, out As Long, in_ As Long

    FunnyLump l, 20
    count = LumpLen(l) \ 20
    out = PoolAlloc(P_PLANES, count * 2)

    mod_known(loadmodel).planes = out
    mod_known(loadmodel).numplanes = count

    in_ = LumpOfs(l)
    For i = 0 To count - 1
        bits = 0
        For j = 0 To 2
            mod_planes(out + i).normal.v(j) = GetFloat(mod_base, in_ + j * 4)
            If mod_planes(out + i).normal.v(j) < 0 Then bits = bits Or (2 ^ j)
        Next

        mod_planes(out + i).dist = GetFloat(mod_base, in_ + 12)
        mod_planes(out + i).typ = GetLong(mod_base, in_ + 16) And 255
        mod_planes(out + i).signbits = bits
        in_ = in_ + 20
    Next
End Sub

Private Function RadiusFromBounds(mins As vec3_t, maxs As vec3_t) As Single
    Dim i As Long, corner As vec3_t

    For i = 0 To 2
        If Abs(mins.v(i)) > Abs(maxs.v(i)) Then corner.v(i) = Abs(mins.v(i)) Else corner.v(i) = Abs(maxs.v(i))
    Next

    RadiusFromBounds = Length(corner)
End Function

Private Sub Mod_LoadBrushModel(ByVal m As Long)
    Dim i As Long, j As Long, bm As Long, name As String

    mod_known(loadmodel).typ = mod_brush

    i = GetLong(mod_base, 0)
    If i <> BSPVERSION Then
        Sys_Error "Mod_LoadBrushModel: %s has wrong version number (%i should be %i)", mod_known(m).name, i, BSPVERSION
    End If

    ' load into heap
    Mod_LoadVertexes LUMP_VERTEXES
    Mod_LoadEdges LUMP_EDGES
    Mod_LoadSurfedges LUMP_SURFEDGES
    Mod_LoadTextures LUMP_TEXTURES
    Mod_LoadLighting LUMP_LIGHTING
    Mod_LoadPlanes LUMP_PLANES
    Mod_LoadTexinfo LUMP_TEXINFO
    Mod_LoadFaces LUMP_FACES
    Mod_LoadMarksurfaces LUMP_MARKSURFACES
    Mod_LoadVisibility LUMP_VISIBILITY
    Mod_LoadLeafs LUMP_LEAFS
    Mod_LoadNodes LUMP_NODES
    Mod_LoadClipnodes LUMP_CLIPNODES
    Mod_LoadEntities LUMP_ENTITIES
    Mod_LoadSubmodels LUMP_MODELS

    Mod_MakeHull0

    mod_known(m).numframes = 2           ' regular and alternate animation
    mod_known(m).flags = 0

    '
    ' set up the submodels (FIXME: this is confusing)
    '
    For i = 0 To mod_known(m).numsubmodels - 1
        bm = mod_known(m).submodels + i

        mod_known(m).hulls(0).firstclipnode = mod_submodels(bm).headnode(0)
        For j = 1 To MAX_MAP_HULLS - 1
            mod_known(m).hulls(j).firstclipnode = mod_submodels(bm).headnode(j)
            mod_known(m).hulls(j).lastclipnode = mod_known(m).numclipnodes - 1
        Next

        mod_known(m).firstmodelsurface = mod_submodels(bm).firstface
        mod_known(m).nummodelsurfaces = mod_submodels(bm).numfaces

        mod_known(m).maxs = mod_submodels(bm).maxs
        mod_known(m).mins = mod_submodels(bm).mins
        mod_known(m).radius = RadiusFromBounds(mod_known(m).mins, mod_known(m).maxs)

        mod_known(m).numleafs = mod_submodels(bm).visleafs

        If i < mod_known(m).numsubmodels - 1 Then
            ' duplicate the basic information
            name = "*" & (i + 1)
            loadmodel = Mod_FindName(name)
            mod_known(loadmodel) = mod_known(m)
            mod_known(loadmodel).name = name
            m = loadmodel
        End If
    Next
End Sub

'==============================================================================
' ALIAS MODELS

' pin: the frame's daliasframe_t in mod_base(); returns the offset after it
Private Function Mod_LoadAliasFrame(ByVal pin As Long, pframeindex As Long, ByVal numv As Long, _
        pbboxmin As trivertx_t, pbboxmax As trivertx_t, name As String) As Long
    Dim i As Long, j As Long, k As Long, pframe As Long, pinframe As Long

    name = BufToStrN(mod_base, pin + 8, 16)

    For i = 0 To 2
        ' these are byte values, so we don't have to worry about
        ' endianness
        pbboxmin.v(i) = mod_base(pin + i)
        pbboxmax.v(i) = mod_base(pin + 4 + i)
    Next

    pinframe = pin + 24
    pframe = num_alias_verts
    num_alias_verts = num_alias_verts + numv
    ReDim Preserve alias_verts(0 To num_alias_verts)

    pframeindex = pframe

    For j = 0 To numv - 1
        ' these are all byte values, so no need to deal with endianness
        alias_verts(pframe + j).lightnormalindex = mod_base(pinframe + j * 4 + 3)
        For k = 0 To 2
            alias_verts(pframe + j).v(k) = mod_base(pinframe + j * 4 + k)
        Next
    Next

    Mod_LoadAliasFrame = pinframe + numv * 4
End Function

Private Function Mod_LoadAliasGroup(ByVal pin As Long, pframeindex As Long, ByVal numv As Long, _
        pbboxmin As trivertx_t, pbboxmax As trivertx_t, name As String) As Long
    Dim i As Long, numframes As Long, g As Long, pin_intervals As Long, ptemp As Long

    numframes = GetLong(mod_base, pin)

    num_alias_groups = num_alias_groups + 1
    ReDim Preserve alias_groups(0 To num_alias_groups)
    g = num_alias_groups
    alias_groups(g).numframes = numframes
    ReDim alias_groups(g).frames(0 To numframes - 1)

    For i = 0 To 2
        ' these are byte values, so we don't have to worry about endianness
        pbboxmin.v(i) = mod_base(pin + 4 + i)
        pbboxmax.v(i) = mod_base(pin + 8 + i)
    Next

    pframeindex = g

    pin_intervals = pin + 12

    alias_groups(g).intervals = num_alias_intervals
    num_alias_intervals = num_alias_intervals + numframes
    ReDim Preserve alias_intervals(0 To num_alias_intervals)

    For i = 0 To numframes - 1
        alias_intervals(alias_groups(g).intervals + i) = GetFloat(mod_base, pin_intervals)
        If alias_intervals(alias_groups(g).intervals + i) <= 0# Then Sys_Error "Mod_LoadAliasGroup: interval<=0"
        pin_intervals = pin_intervals + 4
    Next

    ptemp = pin_intervals

    For i = 0 To numframes - 1
        ptemp = Mod_LoadAliasFrame(ptemp, alias_groups(g).frames(i).frame, numv, _
                    alias_groups(g).frames(i).bboxmin, alias_groups(g).frames(i).bboxmax, name)
    Next

    Mod_LoadAliasGroup = ptemp
End Function

Private Function Mod_LoadAliasSkin(ByVal pin As Long, pskinindex As Long, ByVal skinsize As Long) As Long
    Dim pskin As Long

    pskin = num_alias_skins
    num_alias_skins = num_alias_skins + skinsize
    ReDim Preserve alias_skins(0 To num_alias_skins)

    pskinindex = pskin

    ' (r_pixbytes is 1)
    CopyMemory alias_skins(pskin), mod_base(pin), skinsize

    Mod_LoadAliasSkin = pin + skinsize
End Function

Private Function Mod_LoadAliasSkinGroup(ByVal pin As Long, pskinindex As Long, ByVal skinsize As Long) As Long
    Dim i As Long, numskins As Long, g As Long, pinskinintervals As Long, ptemp As Long

    numskins = GetLong(mod_base, pin)

    num_alias_skingroups = num_alias_skingroups + 1
    ReDim Preserve alias_skingroups(0 To num_alias_skingroups)
    g = num_alias_skingroups
    alias_skingroups(g).numskins = numskins
    ReDim alias_skingroups(g).skindescs(0 To numskins - 1)

    pskinindex = g

    pinskinintervals = pin + 4

    alias_skingroups(g).intervals = num_alias_intervals
    num_alias_intervals = num_alias_intervals + numskins
    ReDim Preserve alias_intervals(0 To num_alias_intervals)

    For i = 0 To numskins - 1
        alias_intervals(alias_skingroups(g).intervals + i) = GetFloat(mod_base, pinskinintervals)
        If alias_intervals(alias_skingroups(g).intervals + i) <= 0 Then Sys_Error "Mod_LoadAliasSkinGroup: interval<=0"
        pinskinintervals = pinskinintervals + 4
    Next

    ptemp = pinskinintervals

    For i = 0 To numskins - 1
        ptemp = Mod_LoadAliasSkin(ptemp, alias_skingroups(g).skindescs(i).skin, skinsize)
    Next

    Mod_LoadAliasSkinGroup = ptemp
End Function

Private Sub Mod_LoadAliasModel(ByVal m As Long)
    Dim i As Long, j As Long, version As Long, numframes As Long, numskins As Long
    Dim h As Long, pskintype As Long, pskindesc As Long, skinsize As Long, skintype As Long
    Dim pinstverts As Long, pintriangles As Long, pframetype As Long, frametype As Long

    version = GetLong(mod_base, 4)
    If version <> ALIAS_VERSION Then
        Sys_Error "%s has wrong version number (%i should be %i)", mod_known(m).name, version, ALIAS_VERSION
    End If

    '
    ' allocate space for a working header
    '
    num_aliashdrs = num_aliashdrs + 1
    ReDim Preserve aliashdrs(0 To num_aliashdrs)
    h = num_aliashdrs

    mod_known(m).flags = GetLong(mod_base, 76)

    '
    ' endian-adjust and copy the data, starting with the alias model header
    '
    With aliashdrs(h).mdl
        .ident = GetLong(mod_base, 0)
        .version = version
        .boundingradius = GetFloat(mod_base, 32)
        .numskins = GetLong(mod_base, 48)
        .skinwidth = GetLong(mod_base, 52)
        .skinheight = GetLong(mod_base, 56)

        If .skinheight > MAX_LBM_HEIGHT Then
            Sys_Error "model %s has a skin taller than %d", mod_known(m).name, MAX_LBM_HEIGHT
        End If

        .numverts = GetLong(mod_base, 60)

        If .numverts <= 0 Then Sys_Error "model %s has no vertices", mod_known(m).name

        If .numverts > MAXALIASVERTS Then Sys_Error "model %s has too many vertices", mod_known(m).name

        .numtris = GetLong(mod_base, 64)

        If .numtris <= 0 Then Sys_Error "model %s has no triangles", mod_known(m).name

        .numframes = GetLong(mod_base, 68)
        .size = GetFloat(mod_base, 80) * ALIAS_BASE_SIZE_RATIO
        .synctype = GetLong(mod_base, 72)
        .flags = mod_known(m).flags
        mod_known(m).synctype = .synctype
        mod_known(m).numframes = .numframes

        For i = 0 To 2
            .scale.v(i) = GetFloat(mod_base, 8 + i * 4)
            .scale_origin.v(i) = GetFloat(mod_base, 20 + i * 4)
            .eyeposition.v(i) = GetFloat(mod_base, 36 + i * 4)
        Next

        numskins = .numskins
        numframes = .numframes

        If .skinwidth And &H3 Then Sys_Error "Mod_LoadAliasModel: skinwidth not multiple of 4"

        skinsize = .skinheight * .skinwidth
    End With

    '
    ' load the skins
    '
    If numskins < 1 Then Sys_Error "Mod_LoadAliasModel: Invalid # of skins: %d\n", numskins

    pskintype = 84

    pskindesc = num_alias_skindescs
    num_alias_skindescs = num_alias_skindescs + numskins
    ReDim Preserve alias_skindescs(0 To num_alias_skindescs)

    aliashdrs(h).skindesc = pskindesc

    For i = 0 To numskins - 1
        skintype = GetLong(mod_base, pskintype)
        alias_skindescs(pskindesc + i).typ = skintype

        If skintype = ALIAS_SKIN_SINGLE Then
            pskintype = Mod_LoadAliasSkin(pskintype + 4, alias_skindescs(pskindesc + i).skin, skinsize)
        Else
            pskintype = Mod_LoadAliasSkinGroup(pskintype + 4, alias_skindescs(pskindesc + i).skin, skinsize)
        End If
    Next

    '
    ' set base s and t vertices
    '
    pinstverts = pskintype
    aliashdrs(h).stverts = num_alias_stverts
    num_alias_stverts = num_alias_stverts + aliashdrs(h).mdl.numverts
    ReDim Preserve alias_stverts(0 To num_alias_stverts)

    For i = 0 To aliashdrs(h).mdl.numverts - 1
        alias_stverts(aliashdrs(h).stverts + i).onseam = GetLong(mod_base, pinstverts + i * 12)
        ' put s and t in 16.16 format
        alias_stverts(aliashdrs(h).stverts + i).s = GetLong(mod_base, pinstverts + i * 12 + 4) * &H10000
        alias_stverts(aliashdrs(h).stverts + i).t = GetLong(mod_base, pinstverts + i * 12 + 8) * &H10000
    Next

    '
    ' set up the triangles
    '
    pintriangles = pinstverts + aliashdrs(h).mdl.numverts * 12
    aliashdrs(h).triangles = num_alias_tris
    num_alias_tris = num_alias_tris + aliashdrs(h).mdl.numtris
    ReDim Preserve alias_tris(0 To num_alias_tris)

    For i = 0 To aliashdrs(h).mdl.numtris - 1
        alias_tris(aliashdrs(h).triangles + i).facesfront = GetLong(mod_base, pintriangles + i * 16)
        For j = 0 To 2
            alias_tris(aliashdrs(h).triangles + i).vertindex(j) = GetLong(mod_base, pintriangles + i * 16 + 4 + j * 4)
        Next
    Next

    '
    ' load the frames
    '
    If numframes < 1 Then Sys_Error "Mod_LoadAliasModel: Invalid # of frames: %d\n", numframes

    ReDim aliashdrs(h).frames(0 To numframes - 1)
    pframetype = pintriangles + aliashdrs(h).mdl.numtris * 16

    For i = 0 To numframes - 1
        frametype = GetLong(mod_base, pframetype)
        aliashdrs(h).frames(i).typ = frametype

        If frametype = ALIAS_SINGLE Then
            pframetype = Mod_LoadAliasFrame(pframetype + 4, aliashdrs(h).frames(i).frame, aliashdrs(h).mdl.numverts, _
                aliashdrs(h).frames(i).bboxmin, aliashdrs(h).frames(i).bboxmax, aliashdrs(h).frames(i).name)
        Else
            pframetype = Mod_LoadAliasGroup(pframetype + 4, aliashdrs(h).frames(i).frame, aliashdrs(h).mdl.numverts, _
                aliashdrs(h).frames(i).bboxmin, aliashdrs(h).frames(i).bboxmax, aliashdrs(h).frames(i).name)
        End If
    Next

    mod_known(m).typ = mod_alias

    ' FIXME: do this right
    mod_known(m).mins.v(0) = -16: mod_known(m).mins.v(1) = -16: mod_known(m).mins.v(2) = -16
    mod_known(m).maxs.v(0) = 16: mod_known(m).maxs.v(1) = 16: mod_known(m).maxs.v(2) = 16

    '
    ' move the complete, relocatable alias model to the cache
    '
    mod_known(m).cache = h
End Sub

'=============================================================================

' pin: the frame's dspriteframe_t in mod_base(); returns the offset after it
Private Function Mod_LoadSpriteFrame(ByVal pin As Long, ppframe As Long) As Long
    Dim width As Long, height As Long, size As Long, f As Long, origin0 As Long, origin1 As Long

    width = GetLong(mod_base, pin + 8)
    height = GetLong(mod_base, pin + 12)
    size = width * height

    f = PoolAlloc(P_SPRITEFRAMES, 1)
    ppframe = f

    sprite_frames(f).width = width
    sprite_frames(f).height = height
    sprite_frames(f).pcachespot = 0
    origin0 = GetLong(mod_base, pin)
    origin1 = GetLong(mod_base, pin + 4)

    sprite_frames(f).up = origin1
    sprite_frames(f).down = origin1 - height
    sprite_frames(f).left_ = origin0
    sprite_frames(f).right_ = width + origin0

    ' (r_pixbytes is 1)
    sprite_frames(f).pixels = PoolAlloc(P_SPRITEPIXELS, size)
    If size > 0 Then CopyMemory sprite_pixels(sprite_frames(f).pixels), mod_base(pin + 16), size

    Mod_LoadSpriteFrame = pin + 16 + size
End Function

Private Function Mod_LoadSpriteGroup(ByVal pin As Long, ppframe As Long) As Long
    Dim i As Long, numframes As Long, g As Long, pin_intervals As Long, ptemp As Long

    numframes = GetLong(mod_base, pin)

    g = PoolAlloc(P_SPRITEGROUPS, 1)
    sprite_groups(g).numframes = numframes
    ReDim sprite_groups(g).frames(0 To numframes - 1)

    ppframe = g

    pin_intervals = pin + 4

    sprite_groups(g).intervals = PoolAlloc(P_SPRITEINTERVALS, numframes)

    For i = 0 To numframes - 1
        sprite_intervals(sprite_groups(g).intervals + i) = GetFloat(mod_base, pin_intervals)
        If sprite_intervals(sprite_groups(g).intervals + i) <= 0# Then Sys_Error "Mod_LoadSpriteGroup: interval<=0"
        pin_intervals = pin_intervals + 4
    Next

    ptemp = pin_intervals

    For i = 0 To numframes - 1
        ptemp = Mod_LoadSpriteFrame(ptemp, sprite_groups(g).frames(i))
    Next

    Mod_LoadSpriteGroup = ptemp
End Function

Private Sub Mod_LoadSpriteModel(ByVal m As Long)
    Dim i As Long, version As Long, numframes As Long, ps As Long, pframetype As Long, frametype As Long

    version = GetLong(mod_base, 4)
    If version <> SPRITE_VERSION Then
        Sys_Error "%s has wrong version number (%i should be %i)", mod_known(m).name, version, SPRITE_VERSION
    End If

    numframes = GetLong(mod_base, 24)

    ps = PoolAlloc(P_SPRITES, 1)

    mod_known(m).cache = ps

    msprites(ps).typ = GetLong(mod_base, 8)
    msprites(ps).maxwidth = GetLong(mod_base, 16)
    msprites(ps).maxheight = GetLong(mod_base, 20)
    msprites(ps).beamlength = GetFloat(mod_base, 28)
    mod_known(m).synctype = GetLong(mod_base, 32)
    msprites(ps).numframes = numframes

    mod_known(m).mins.v(0) = -msprites(ps).maxwidth \ 2
    mod_known(m).mins.v(1) = mod_known(m).mins.v(0)
    mod_known(m).maxs.v(0) = msprites(ps).maxwidth \ 2
    mod_known(m).maxs.v(1) = mod_known(m).maxs.v(0)
    mod_known(m).mins.v(2) = -msprites(ps).maxheight \ 2
    mod_known(m).maxs.v(2) = msprites(ps).maxheight \ 2

    '
    ' load the frames
    '
    If numframes < 1 Then Sys_Error "Mod_LoadSpriteModel: Invalid # of frames: %d\n", numframes

    mod_known(m).numframes = numframes
    mod_known(m).flags = 0

    ReDim msprites(ps).frames(0 To numframes - 1)
    pframetype = 36

    For i = 0 To numframes - 1
        frametype = GetLong(mod_base, pframetype)
        msprites(ps).frames(i).typ = frametype

        If frametype = SPR_SINGLE Then
            pframetype = Mod_LoadSpriteFrame(pframetype + 4, msprites(ps).frames(i).frameptr)
        Else
            pframetype = Mod_LoadSpriteGroup(pframetype + 4, msprites(ps).frames(i).frameptr)
        End If
    Next

    mod_known(m).typ = mod_sprite
End Sub

'=============================================================================

Public Sub Mod_Print()
    Dim i As Long

    Con_Printf "Cached models:\n"
    For i = 1 To mod_numknown
        Con_Printf "%8p : %s", mod_known(i).cache, mod_known(i).name
        If mod_known(i).needload And NL_UNREFERENCED Then Con_Printf " (!R)"
        If mod_known(i).needload And NL_NEEDS_LOADED Then Con_Printf " (!P)"
        Con_Printf "\n"
    Next
End Sub
