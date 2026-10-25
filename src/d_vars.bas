Attribute VB_Name = "d_vars"
' d_vars.bas - global refresh variables of the rasterization driver (d_vars.c)
'
' cacheblock and d_viewbuffer are pointers in C. Here they are Byte arrays
' laid over other memory (SetView, common.bas): cacheblock() over the
' surface cache block, the texture or the sprite frame being drawn
' (D_SetCacheBlock), d_viewbuffer() over vid_buffer() or the warp buffer
' (D_SetupFrame). D_ClearViews detaches them (and the surface cache views
' and r_alias.bas's acolormap()) before the program ends.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' all global and static refresh variables are collected in a contiguous block
' to avoid cache conflicts.

'-------------------------------------------------------
' global refresh variables
'-------------------------------------------------------

' FIXME: make into one big structure, like cl or sv
' FIXME: do separately for refresh engine and driver

Public d_sdivzstepu As Single, d_tdivzstepu As Single, d_zistepu As Single
Public d_sdivzstepv As Single, d_tdivzstepv As Single, d_zistepv As Single
Public d_sdivzorigin As Single, d_tdivzorigin As Single, d_ziorigin As Single

Public sadjust As Long, tadjust As Long, bbextents As Long, bbextentt As Long

Public cacheblock() As Byte
Public cachewidth As Long

Public d_viewbuffer() As Byte
Public d_pzbuffer() As Integer           ' (allocated by VID_Init, as WinQuake's VID_AllocBuffers)
Public d_zrowbytes As Long
Public d_zwidth As Long

Private sa_cacheblock As SAFEARRAY1D
Private sa_viewbuffer As SAFEARRAY1D

' (QuakeVB) points cacheblock() at count bytes at pData
Public Sub D_SetCacheBlock(ByVal pData As Long, ByVal count As Long)
    SetView VarPtrArray(cacheblock), sa_cacheblock, pData, count, 1
End Sub

' (QuakeVB) points d_viewbuffer() at count bytes at pData
Public Sub D_SetViewBuffer(ByVal pData As Long, ByVal count As Long)
    SetView VarPtrArray(d_viewbuffer), sa_viewbuffer, pData, count, 1
End Sub

' (QuakeVB) detaches the renderer's views before the program ends
Public Sub D_ClearViews()
    ClearView VarPtrArray(cacheblock)
    ClearView VarPtrArray(d_viewbuffer)
    D_ClearCacheViews
    R_AliasClearViews
End Sub
