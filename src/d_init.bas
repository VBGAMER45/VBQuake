Attribute VB_Name = "d_init_c"
' d_init.bas - rasterization driver initialization (d_init.c)
' (module d_init_c: VB names are not case sensitive, and this has D_Init)
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const NUM_MIPS As Long = 4

Public cv_d_subdiv16 As Long
Private cv_d_mipcap As Long
Private cv_d_mipscale As Long

Public d_initial_rover As Long           ' surface cache block (d_surf.bas), 0 = NULL
Public d_roverwrapped As Boolean
Public d_minmip As Long
Public d_scalemip(0 To NUM_MIPS - 2) As Single

' (C: static float basemip[NUM_MIPS-1] = {1.0, 0.5*0.8, 0.25*0.8})
Private Function basemip(ByVal i As Long) As Single
    Select Case i
    Case 0: basemip = 1#
    Case 1: basemip = 0.5 * 0.8
    Case Else: basemip = 0.25 * 0.8
    End Select
End Function

' (C: void (*d_drawspans) (espan_t *pspan) is always D_DrawSpans8 without
' id386; d_edge.bas calls that directly)

Public Sub D_Init()
    r_skydirect = 1

    cv_d_subdiv16 = Cvar_RegisterVariable("d_subdiv16", "1")
    cv_d_mipcap = Cvar_RegisterVariable("d_mipcap", "0")
    cv_d_mipscale = Cvar_RegisterVariable("d_mipscale", "1")

    r_drawpolys = False
    r_worldpolysbacktofront = False
    r_recursiveaffinetriangles = True
    r_pixbytes = 1
    r_aliasuvscale = 1#
End Sub

Public Sub D_CopyRects(prects As vrect_t, ByVal transparent As Long)
    ' this function is only required if the CPU doesn't have direct access to the
    ' back buffer, and there's some driver interface function that the driver
    ' doesn't support and requires Quake to do in software (such as drawing the
    ' console); Quake will then draw into wherever the driver points vid.buffer
    ' and will call this function before swapping buffers
End Sub

Public Sub D_EnableBackBufferAccess()
    VID_LockBuffer
End Sub

Public Sub D_TurnZOn()
    ' not needed for software version
End Sub

Public Sub D_DisableBackBufferAccess()
    VID_UnlockBuffer
End Sub

Public Sub D_SetupFrame()
    Dim i As Long

    If r_dowarp Then
        D_SetViewBuffer VarPtr(r_warpbuffer(0)), WARP_WIDTH * WARP_HEIGHT
    Else
        D_SetViewBuffer VarPtr(vid_buffer(vid_bufferofs)), UBound(vid_buffer) + 1 - vid_bufferofs
    End If

    If r_dowarp Then
        screenwidth = WARP_WIDTH
    Else
        screenwidth = vid.rowbytes
    End If

    d_roverwrapped = False
    d_initial_rover = sc_rover

    d_minmip = ToInt(cvars(cv_d_mipcap).value)
    If d_minmip > 3 Then
        d_minmip = 3
    ElseIf d_minmip < 0 Then
        d_minmip = 0
    End If

    For i = 0 To NUM_MIPS - 2
        d_scalemip(i) = basemip(i) * cvars(cv_d_mipscale).value
    Next

    d_aflatcolor = 0
End Sub

Public Sub D_UpdateRects(prect As vrect_t)
    ' the software driver draws these directly to the vid buffer
End Sub
