Attribute VB_Name = "d_surf"
' d_surf.bas - rasterization driver surface heap manager (d_surf.c)
'
' The surface cache is one block of memory, sc_l(), with the views sc_bytes()
' (Byte) and sc_f() (Single) over it. A surfcache_t is a header followed by
' its pixels, in C's layout (data at byte 48, sizeof 52), so the sizes come
' out as in C. A block (C: surfcache_t *) is its byte offset + 1, 0 = NULL;
' an owner (C: surfcache_t **, the msurface_t's cachespots[miplevel]) is
' surface * MIPLEVELS + miplevel + 1, 0 = NULL.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public surfscale As Single
' (r_cache_thrash, set if surface cache is thrashing, is in quakedef.bas)

Public sc_size As Long
Public sc_rover As Long, sc_base As Long ' blocks, 0 = NULL

Private sc_l() As Long                   ' the memory (C: vid_surfcache)
Public sc_bytes() As Byte                ' views of it
Private sc_f() As Single
Private sa_scbytes As SAFEARRAY1D
Private sa_scf As SAFEARRAY1D

Private Const GUARDSIZE As Long = 4

' surfcache_t: the Long offsets of the header fields
Private Const SCF_NEXT As Long = 0        ' struct surfcache_s *next
Private Const SCF_OWNER As Long = 1       ' struct surfcache_s **owner (NULL is an empty chunk of memory)
Private Const SCF_LIGHTADJ As Long = 2    ' int lightadj[MAXLIGHTMAPS] (checked for strobe flush)
Private Const SCF_DLIGHT As Long = 6
Private Const SCF_SIZE As Long = 7        ' including header
Private Const SCF_WIDTH As Long = 8
Private Const SCF_HEIGHT As Long = 9      ' DEBUG only needed for debug
Private Const SCF_MIPSCALE As Long = 10   ' float
Private Const SCF_TEXTURE As Long = 11    ' struct texture_s *texture (checked for animating textures)
Public Const SC_DATA As Long = 48        ' byte data[4] (width*height elements)
Private Const SC_SIZEOF As Long = 52     ' sizeof (surfcache_t)

' (QuakeVB) the Long index of a block's header
Private Function H(ByVal c As Long) As Long
    H = (c - 1) \ 4
End Function

' (QuakeVB) *owner = NULL
Private Sub ClearOwner(ByVal owner As Long)
    owner = owner - 1
    mod_surfaces(owner \ MIPLEVELS).cachespots(owner Mod MIPLEVELS) = 0
End Sub

' (QuakeVB) a block's width, and the address of its pixels (for cacheblock)
Public Function D_CacheWidth(ByVal c As Long) As Long
    D_CacheWidth = sc_l(H(c) + SCF_WIDTH)
End Function

Public Function D_CacheDataOfs(ByVal c As Long) As Long
    D_CacheDataOfs = c - 1 + SC_DATA
End Function

Public Function D_CacheDataPtr(ByVal c As Long) As Long
    D_CacheDataPtr = VarPtr(sc_l(0)) + c - 1 + SC_DATA
End Function

' (QuakeVB) the bytes of pixels a block has room for
Public Function D_CacheDataBytes(ByVal c As Long) As Long
    D_CacheDataBytes = sc_l(H(c) + SCF_SIZE) - SC_DATA
End Function

Public Function D_SurfaceCacheForRes(ByVal width As Long, ByVal height As Long) As Long
    Dim size As Long, pix As Long

    If COM_CheckParm("-surfcachesize") Then
        size = Q_atoi(com_argv(COM_CheckParm("-surfcachesize") + 1)) * 1024
        D_SurfaceCacheForRes = size
        Exit Function
    End If

    size = SURFCACHE_SIZE_AT_320X200

    pix = width * height
    If pix > 64000 Then size = size + (pix - 64000) * 3

    D_SurfaceCacheForRes = size
End Function

Private Sub D_CheckCacheGuard()
    Dim i As Long

    For i = 0 To GUARDSIZE - 1
        If sc_bytes(sc_size + i) <> i Then Sys_Error "D_CheckCacheGuard: failed"
    Next
End Sub

Private Sub D_ClearCacheGuard()
    Dim i As Long

    For i = 0 To GUARDSIZE - 1
        sc_bytes(sc_size + i) = i
    Next
End Sub

' (QuakeVB) C passes the buffer VID_AllocBuffers made; here the cache
' allocates its own
Public Sub D_InitCaches(ByVal size As Long)
    If Not msg_suppress_1 Then Con_Printf "%ik surface cache\n", size \ 1024

    D_ClearCacheViews
    ReDim sc_l(0 To (size + 3) \ 4)
    SetView VarPtrArray(sc_bytes), sa_scbytes, VarPtr(sc_l(0)), size, 1
    SetView VarPtrArray(sc_f), sa_scf, VarPtr(sc_l(0)), (size + 3) \ 4, 4

    sc_size = size - GUARDSIZE
    sc_base = 1
    sc_rover = sc_base

    sc_l(H(sc_base) + SCF_NEXT) = 0
    sc_l(H(sc_base) + SCF_OWNER) = 0
    sc_l(H(sc_base) + SCF_SIZE) = sc_size

    D_ClearCacheGuard
End Sub

' (QuakeVB) detaches the views (d_vars.bas's D_ClearViews)
Public Sub D_ClearCacheViews()
    ClearView VarPtrArray(sc_bytes)
    ClearView VarPtrArray(sc_f)
End Sub

Public Sub D_FlushCaches()
    Dim c As Long

    If sc_base = 0 Then Exit Sub

    c = sc_base
    Do While c <> 0
        If sc_l(H(c) + SCF_OWNER) Then ClearOwner sc_l(H(c) + SCF_OWNER)
        c = sc_l(H(c) + SCF_NEXT)
    Loop

    sc_rover = sc_base
    sc_l(H(sc_base) + SCF_NEXT) = 0
    sc_l(H(sc_base) + SCF_OWNER) = 0
    sc_l(H(sc_base) + SCF_SIZE) = sc_size
End Sub

Private Function D_SCAlloc(ByVal width As Long, ByVal size As Long) As Long
    Dim new_ As Long, wrapped_this_time As Boolean

    If width < 0 Or width > 256 Then Sys_Error "D_SCAlloc: bad cache width %d\n", width

    If size <= 0 Or size > &H10000 Then Sys_Error "D_SCAlloc: bad cache size %d\n", size

    size = SC_DATA + size                ' (int)&((surfcache_t *)0)->data[size]
    size = (size + 3) And Not 3
    If size > sc_size Then Sys_Error "D_SCAlloc: %i > cache size", size

    ' if there is not size bytes after the rover, reset to the start
    wrapped_this_time = False

    If sc_rover = 0 Then
        sc_rover = sc_base
    ElseIf sc_rover - sc_base > sc_size - size Then
        wrapped_this_time = True
        sc_rover = sc_base
    End If

    ' colect and free surfcache_t blocks until the rover block is large enough
    new_ = sc_rover
    If sc_l(H(sc_rover) + SCF_OWNER) Then ClearOwner sc_l(H(sc_rover) + SCF_OWNER)

    Do While sc_l(H(new_) + SCF_SIZE) < size
        ' free another
        sc_rover = sc_l(H(sc_rover) + SCF_NEXT)
        If sc_rover = 0 Then Sys_Error "D_SCAlloc: hit the end of memory"
        If sc_l(H(sc_rover) + SCF_OWNER) Then ClearOwner sc_l(H(sc_rover) + SCF_OWNER)

        sc_l(H(new_) + SCF_SIZE) = sc_l(H(new_) + SCF_SIZE) + sc_l(H(sc_rover) + SCF_SIZE)
        sc_l(H(new_) + SCF_NEXT) = sc_l(H(sc_rover) + SCF_NEXT)
    Loop

    ' create a fragment out of any leftovers
    If sc_l(H(new_) + SCF_SIZE) - size > 256 Then
        sc_rover = new_ + size
        sc_l(H(sc_rover) + SCF_SIZE) = sc_l(H(new_) + SCF_SIZE) - size
        sc_l(H(sc_rover) + SCF_NEXT) = sc_l(H(new_) + SCF_NEXT)
        sc_l(H(sc_rover) + SCF_WIDTH) = 0
        sc_l(H(sc_rover) + SCF_OWNER) = 0
        sc_l(H(new_) + SCF_NEXT) = sc_rover
        sc_l(H(new_) + SCF_SIZE) = size
    Else
        sc_rover = sc_l(H(new_) + SCF_NEXT)
    End If

    sc_l(H(new_) + SCF_WIDTH) = width
    ' DEBUG
    If width > 0 Then sc_l(H(new_) + SCF_HEIGHT) = (size - SC_SIZEOF + 4) \ width

    sc_l(H(new_) + SCF_OWNER) = 0         ' should be set properly after return

    If d_roverwrapped Then
        If wrapped_this_time Or sc_rover >= d_initial_rover Then r_cache_thrash = True
    ElseIf wrapped_this_time Then
        d_roverwrapped = True
    End If

    D_CheckCacheGuard                    ' DEBUG
    D_SCAlloc = new_
End Function

Public Sub D_SCDump()
    Dim test As Long

    test = sc_base
    Do While test <> 0
        If test = sc_rover Then Sys_Printf "ROVER:\n"
        Sys_Printf "%p : %i bytes     %i width\n", test - 1, sc_l(H(test) + SCF_SIZE), sc_l(H(test) + SCF_WIDTH)
        test = sc_l(H(test) + SCF_NEXT)
    Loop
End Sub

'=============================================================================

' if the num is not a power of 2, assume it will not repeat
Public Function MaskForNum(ByVal num As Long) As Long
    If num = 128 Then
        MaskForNum = 127
    ElseIf num = 64 Then
        MaskForNum = 63
    ElseIf num = 32 Then
        MaskForNum = 31
    ElseIf num = 16 Then
        MaskForNum = 15
    Else
        MaskForNum = 255
    End If
End Function

Public Function D_log2(ByVal num As Long) As Long
    Dim c As Long

    c = 0

    num = num \ 2
    Do While num <> 0
        c = c + 1
        num = num \ 2
    Loop
    D_log2 = c
End Function

'=============================================================================

Public Function D_CacheSurface(ByVal surface As Long, ByVal miplevel As Long) As Long
    Dim cache As Long, i As Long

    '
    ' if the surface is animating or flashing, flush the cache
    '
    r_drawsurf.texture = R_TextureAnimation(mod_texinfo(mod_surfaces(surface).texinfo).texture)
    For i = 0 To MAXLIGHTMAPS - 1
        r_drawsurf.lightadj(i) = d_lightstylevalue(mod_surfaces(surface).styles(i))
    Next

    '
    ' see if the cache holds apropriate data
    '
    cache = mod_surfaces(surface).cachespots(miplevel)

    If cache <> 0 Then
        If sc_l(H(cache) + SCF_DLIGHT) = 0 And mod_surfaces(surface).dlightframe <> r_framecount _
            And sc_l(H(cache) + SCF_TEXTURE) = r_drawsurf.texture _
            And sc_l(H(cache) + SCF_LIGHTADJ) = r_drawsurf.lightadj(0) _
            And sc_l(H(cache) + SCF_LIGHTADJ + 1) = r_drawsurf.lightadj(1) _
            And sc_l(H(cache) + SCF_LIGHTADJ + 2) = r_drawsurf.lightadj(2) _
            And sc_l(H(cache) + SCF_LIGHTADJ + 3) = r_drawsurf.lightadj(3) Then
            D_CacheSurface = cache
            Exit Function
        End If
    End If

    '
    ' determine shape of surface
    '
    surfscale = 1# / BIT_(miplevel)
    r_drawsurf.surfmip = miplevel
    r_drawsurf.surfwidth = mod_surfaces(surface).extents(0) \ BIT_(miplevel)
    r_drawsurf.rowbytes = r_drawsurf.surfwidth
    r_drawsurf.surfheight = mod_surfaces(surface).extents(1) \ BIT_(miplevel)

    '
    ' allocate memory if needed
    '
    If cache = 0 Then                    ' if a texture just animated, don't reallocate it
        cache = D_SCAlloc(r_drawsurf.surfwidth, r_drawsurf.surfwidth * r_drawsurf.surfheight)
        mod_surfaces(surface).cachespots(miplevel) = cache
        sc_l(H(cache) + SCF_OWNER) = surface * MIPLEVELS + miplevel + 1
        sc_f(H(cache) + SCF_MIPSCALE) = surfscale
    End If

    If mod_surfaces(surface).dlightframe = r_framecount Then
        sc_l(H(cache) + SCF_DLIGHT) = 1
    Else
        sc_l(H(cache) + SCF_DLIGHT) = 0
    End If

    r_drawsurf.surfdat = D_CacheDataOfs(cache)

    sc_l(H(cache) + SCF_TEXTURE) = r_drawsurf.texture
    For i = 0 To MAXLIGHTMAPS - 1
        sc_l(H(cache) + SCF_LIGHTADJ + i) = r_drawsurf.lightadj(i)
    Next

    '
    ' draw and light the surface texture
    '
    r_drawsurf.surf = surface

    c_surf = c_surf + 1
    R_DrawSurface

    D_CacheSurface = mod_surfaces(surface).cachespots(miplevel)
End Function
