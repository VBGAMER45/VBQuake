Attribute VB_Name = "screen"
' screen.bas - master for refresh, status bar, console, chat, notify, etc
' (screen.c).
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private oldscreensize As Single
Private oldfov As Single

Public cv_scr_fov As Long            ' 10 - 170
Public cv_scr_conspeed As Long
Public cv_scr_centertime As Long
Public cv_scr_showram As Long
Public cv_scr_showturtle As Long
Public cv_scr_showpause As Long
Public cv_scr_printspeed As Long

Private scr_initialized As Boolean   ' ready to draw

Private scr_ram As Long
Private scr_net As Long
Private scr_turtle As Long

Private clearconsole As Long

Public scr_drawloading As Boolean
Public scr_disabled_time As Single

' C statics of SCR_DrawTurtle and SCR_UpdateScreen
Private turtle_count As Long
Private oldscr_viewsize As Single
Private oldlcd_x As Single

Private scr_notifystring As String
Private scr_drawdialog As Boolean

'===============================================================================
' CENTER PRINTING

Private scr_centerstring As String
Private scr_centertime_start As Single   ' for slow victory printing
Public scr_centertime_off As Single
Private scr_center_lines As Long
Private scr_erase_lines As Long
Private scr_erase_center As Long

' Called for important messages that should stay in the center of the screen
' for a few moments
Public Sub SCR_CenterPrint(ByVal str_ As String)
    Dim i As Long

    scr_centerstring = Left$(str_, 1024 - 1)
    scr_centertime_off = cvars(cv_scr_centertime).value
    scr_centertime_start = cl.time

    ' count the number of lines for centering
    scr_center_lines = 1
    For i = 1 To Len(str_)
        If Mid$(str_, i, 1) = vbLf Then scr_center_lines = scr_center_lines + 1
    Next
End Sub

Private Sub SCR_EraseCenterString()
    Dim y As Long, t As Long

    t = scr_erase_center
    scr_erase_center = scr_erase_center + 1
    If t > vid.numpages Then
        scr_erase_lines = 0
        Exit Sub
    End If

    If scr_center_lines <= 4 Then
        y = Fix(vid.height * 0.35)
    Else
        y = 48
    End If

    scr_copytop = 1
    Draw_TileClear 0, y, vid.width, 8 * scr_erase_lines
End Sub

Private Sub SCR_DrawCenterString()
    Dim start As Long, l As Long, j As Long, x As Long, y As Long, remaining As Long
    Dim s As String, n As Long

    ' the finale prints the characters one at a time
    If cl.intermission Then
        remaining = Fix(cvars(cv_scr_printspeed).value * (cl.time - scr_centertime_start))
    Else
        remaining = 9999
    End If

    scr_erase_center = 0
    s = scr_centerstring
    n = Len(s)
    start = 1

    If scr_center_lines <= 4 Then
        y = Fix(vid.height * 0.35)
    Else
        y = 48
    End If

    Do
        ' scan the width of the line
        For l = 0 To 39
            If start + l > n Then Exit For
            If Mid$(s, start + l, 1) = vbLf Then Exit For
        Next
        x = (vid.width - l * 8) \ 2
        For j = 0 To l - 1
            Draw_Character x, y, AscW(Mid$(s, start + j, 1))
            If remaining = 0 Then Exit Sub
            remaining = remaining - 1
            x = x + 8
        Next

        y = y + 8

        Do While start <= n
            If Mid$(s, start, 1) = vbLf Then Exit Do
            start = start + 1
        Loop

        If start > n Then Exit Do
        start = start + 1                ' skip the \n
    Loop
End Sub

Private Sub SCR_CheckDrawCenterString()
    scr_copytop = 1
    If scr_center_lines > scr_erase_lines Then scr_erase_lines = scr_center_lines

    scr_centertime_off = scr_centertime_off - host_frametime

    If scr_centertime_off <= 0 And cl.intermission = 0 Then Exit Sub
    If key_dest <> key_game Then Exit Sub

    SCR_DrawCenterString
End Sub

'=============================================================================

Private Function CalcFov(ByVal fov_x As Single, ByVal width As Single, ByVal height As Single) As Single
    Dim a As Single, x As Single

    If fov_x < 1 Or fov_x > 179 Then Sys_Error "Bad fov: %f", fov_x

    ' (QuakeVB) the first call comes before R_ViewChanged has set the vrect;
    ' the C computes atan (0/0) then. That NaN would leave the FPU's invalid
    ' flag set, and VB's next float to integer conversion would fail with
    ' error 6. fov_y is not used by the software renderer.
    If width = 0 Or height = 0 Then
        CalcFov = 0
        Exit Function
    End If

    x = width / Tan(fov_x / 360 * M_PI)

    a = Atn(height / x)

    a = a * 360 / M_PI

    CalcFov = a
End Function

' Must be called whenever vid changes
' Internal use only
Private Sub SCR_CalcRefdef()
    Dim vrect As vrect_t, size As Single

    scr_fullupdate = 0                   ' force a background redraw
    vid.recalc_refdef = 0

    ' force the status bar to redraw
    Sbar_Changed

    '========================================

    ' bound viewsize
    If cvars(cv_scr_viewsize).value < 30 Then Cvar_Set "viewsize", "30"
    If cvars(cv_scr_viewsize).value > 120 Then Cvar_Set "viewsize", "120"

    ' bound field of view
    If cvars(cv_scr_fov).value < 10 Then Cvar_Set "fov", "10"
    If cvars(cv_scr_fov).value > 170 Then Cvar_Set "fov", "170"

    r_refdef.fov_x = cvars(cv_scr_fov).value
    r_refdef.fov_y = CalcFov(r_refdef.fov_x, r_refdef.vrect.width, r_refdef.vrect.height)

    ' intermission is always full screen
    If cl.intermission Then
        size = 120
    Else
        size = cvars(cv_scr_viewsize).value
    End If

    If size >= 120 Then
        sb_lines = 0                     ' no status bar at all
    ElseIf size >= 110 Then
        sb_lines = 24                    ' no inventory
    Else
        sb_lines = 24 + 16 + 8
    End If

    ' these calculations mirror those in R_Init() for r_refdef, but take no
    ' account of water warping
    vrect.x = 0
    vrect.y = 0
    vrect.width = vid.width
    vrect.height = vid.height

    R_SetVrect vrect, scr_vrect, sb_lines

    ' guard against going from one mode to another that's less than half the
    ' vertical resolution
    If scr_con_current > vid.height Then scr_con_current = vid.height

    ' notify the refresh of the change
    R_ViewChanged vrect, sb_lines, vid.aspect
End Sub

' Keybinding command
Public Sub SCR_SizeUp_f()
    Cvar_SetValue "viewsize", cvars(cv_scr_viewsize).value + 10
    vid.recalc_refdef = 1
End Sub

' Keybinding command
Public Sub SCR_SizeDown_f()
    Cvar_SetValue "viewsize", cvars(cv_scr_viewsize).value - 10
    vid.recalc_refdef = 1
End Sub

'============================================================================

Public Sub SCR_Init()
    cv_scr_fov = Cvar_RegisterVariable("fov", "90")
    cv_scr_viewsize = Cvar_RegisterVariable("viewsize", "100", True)
    cv_scr_conspeed = Cvar_RegisterVariable("scr_conspeed", "300")
    cv_scr_showram = Cvar_RegisterVariable("showram", "1")
    cv_scr_showturtle = Cvar_RegisterVariable("showturtle", "0")
    cv_scr_showpause = Cvar_RegisterVariable("showpause", "1")
    cv_scr_centertime = Cvar_RegisterVariable("scr_centertime", "2")
    cv_scr_printspeed = Cvar_RegisterVariable("scr_printspeed", "8")

    '
    ' register our commands
    '
    Cmd_AddCommand "screenshot", XC_SCR_ScreenShot_f
    Cmd_AddCommand "sizeup", XC_SCR_SizeUp_f
    Cmd_AddCommand "sizedown", XC_SCR_SizeDown_f

    scr_ram = Draw_PicFromWad("ram")
    scr_net = Draw_PicFromWad("net")
    scr_turtle = Draw_PicFromWad("turtle")

    scr_initialized = True
End Sub

Private Sub SCR_DrawRam()
    If cvars(cv_scr_showram).value = 0 Then Exit Sub

    If Not r_cache_thrash Then Exit Sub

    Draw_Pic scr_vrect.x + 32, scr_vrect.y, scr_ram
End Sub

Private Sub SCR_DrawTurtle()
    If cvars(cv_scr_showturtle).value = 0 Then Exit Sub

    If host_frametime < 0.1 Then
        turtle_count = 0
        Exit Sub
    End If

    turtle_count = turtle_count + 1
    If turtle_count < 3 Then Exit Sub

    Draw_Pic scr_vrect.x, scr_vrect.y, scr_turtle
End Sub

Private Sub SCR_DrawNet()
    If realtime - cl.last_received_message < 0.3 Then Exit Sub
    If cls.demoplayback Then Exit Sub

    Draw_Pic scr_vrect.x + 64, scr_vrect.y, scr_net
End Sub

Private Sub SCR_DrawPause()
    Dim pic As Long

    If cvars(cv_scr_showpause).value = 0 Then Exit Sub   ' turn off for screenshots

    If Not cl.paused Then Exit Sub

    pic = Draw_CachePic("gfx/pause.lmp")
    Draw_Pic (vid.width - pics(pic).width) \ 2, (vid.height - 48 - pics(pic).height) \ 2, pic
End Sub

Private Sub SCR_DrawLoading_()
    Dim pic As Long

    If Not scr_drawloading Then Exit Sub

    pic = Draw_CachePic("gfx/loading.lmp")
    Draw_Pic (vid.width - pics(pic).width) \ 2, (vid.height - 48 - pics(pic).height) \ 2, pic
End Sub

'=============================================================================

Private Sub SCR_SetUpToDrawConsole()
    Dim t As Long

    Con_CheckResize

    If scr_drawloading Then Exit Sub     ' never a console with loading plaque

    ' decide on the height of the console
    con_forcedup = (cl.worldmodel = 0) Or (cls.signon <> SIGNONS)

    If con_forcedup Then
        scr_conlines = vid.height        ' full screen
        scr_con_current = scr_conlines
    ElseIf key_dest = key_console Then
        scr_conlines = vid.height \ 2    ' half screen
    Else
        scr_conlines = 0                 ' none visible
    End If

    If scr_conlines < scr_con_current Then
        scr_con_current = scr_con_current - cvars(cv_scr_conspeed).value * host_frametime
        If scr_conlines > scr_con_current Then scr_con_current = scr_conlines
    ElseIf scr_conlines > scr_con_current Then
        scr_con_current = scr_con_current + cvars(cv_scr_conspeed).value * host_frametime
        If scr_conlines < scr_con_current Then scr_con_current = scr_conlines
    End If

    t = clearconsole
    clearconsole = clearconsole + 1
    If t < vid.numpages Then
        scr_copytop = 1
        Draw_TileClear 0, Fix(scr_con_current), vid.width, vid.height - Fix(scr_con_current)
        Sbar_Changed
    Else
        t = clearnotify
        clearnotify = clearnotify + 1
        If t < vid.numpages Then
            scr_copytop = 1
            Draw_TileClear 0, 0, vid.width, con_notifylines
        Else
            con_notifylines = 0
        End If
    End If
End Sub

Private Sub SCR_DrawConsole()
    If scr_con_current <> 0 Then
        scr_copyeverything = 1
        Con_DrawConsole Fix(scr_con_current), True
        clearconsole = 0
    Else
        If key_dest = key_game Or key_dest = key_message Then Con_DrawNotify   ' only draw notify in game
    End If
End Sub

'==============================================================================
' SCREEN SHOTS

Private Sub WritePCXfile(ByVal filename As String, data() As Byte, ByVal width As Long, ByVal height As Long, _
        ByVal rowbytes As Long, palette() As Byte)
    Dim i As Long, j As Long, pack As Long, src As Long, pcx() As Byte, b As Byte

    ReDim pcx(0 To width * height * 2 + 1000 - 1)

    pcx(0) = &HA                         ' manufacturer: PCX id
    pcx(1) = 5                           ' version: 256 color
    pcx(2) = 1                           ' encoding: uncompressed
    pcx(3) = 8                           ' bits_per_pixel: 256 color
    PutShort pcx, 4, 0                   ' xmin
    PutShort pcx, 6, 0                   ' ymin
    PutShort pcx, 8, width - 1           ' xmax
    PutShort pcx, 10, height - 1         ' ymax
    PutShort pcx, 12, width              ' hres
    PutShort pcx, 14, height             ' vres
    ' palette[48] and reserved stay 0
    pcx(65) = 1                          ' color_planes: chunky image
    PutShort pcx, 66, width              ' bytes_per_line
    PutShort pcx, 68, 2                  ' palette_type: not a grey scale
    ' filler[58] stays 0

    ' pack the image
    pack = 128
    src = 0
    For i = 0 To height - 1
        For j = 0 To width - 1
            b = data(src)
            If (b And &HC0) <> &HC0 Then
                pcx(pack) = b
                pack = pack + 1
            Else
                pcx(pack) = &HC1
                pcx(pack + 1) = b
                pack = pack + 2
            End If
            src = src + 1
        Next
        src = src + rowbytes - width
    Next

    ' write the palette
    pcx(pack) = &HC                      ' palette ID byte
    pack = pack + 1
    CopyMemory pcx(pack), palette(0), 768
    pack = pack + 768

    ' write output file
    COM_WriteFile filename, pcx, pack
End Sub

Public Sub SCR_ScreenShot_f()
    Dim i As Long, pcxname As String, checkname As String

    '
    ' find a file name to save it to
    '
    For i = 0 To 99
        pcxname = "quake" & (i \ 10) & (i Mod 10) & ".pcx"
        checkname = com_gamedir & "/" & pcxname
        If Sys_FileTime(checkname) = -1 Then Exit For   ' file doesn't exist
    Next
    If i = 100 Then
        Con_Printf "SCR_ScreenShot_f: Couldn't create a PCX file\n"
        Exit Sub
    End If

    '
    ' save the pcx file
    '
    D_EnableBackBufferAccess             ' enable direct drawing of console to back
                                         ' buffer

    WritePCXfile pcxname, vid_buffer, vid.width, vid.height, vid.rowbytes, host_basepal

    D_DisableBackBufferAccess            ' for adapters that can't stay mapped in
                                         ' for linear writes all the time

    Con_Printf "Wrote %s\n", pcxname
End Sub

'=============================================================================

Public Sub SCR_BeginLoadingPlaque()
    S_StopAllSounds True

    If cls.state <> ca_connected Then Exit Sub
    If cls.signon <> SIGNONS Then Exit Sub

    ' redraw with no console and the loading plaque
    Con_ClearNotify
    scr_centertime_off = 0
    scr_con_current = 0

    scr_drawloading = True
    scr_fullupdate = 0
    Sbar_Changed
    SCR_UpdateScreen
    scr_drawloading = False

    scr_disabled_for_loading = True
    scr_disabled_time = realtime
    scr_fullupdate = 0
End Sub

Public Sub SCR_EndLoadingPlaque()
    scr_disabled_for_loading = False
    scr_fullupdate = 0
    Con_ClearNotify
End Sub

'=============================================================================

Private Sub SCR_DrawNotifyString()
    Dim start As Long, l As Long, j As Long, x As Long, y As Long, s As String, n As Long

    s = scr_notifystring
    n = Len(s)
    start = 1

    y = Fix(vid.height * 0.35)

    Do
        ' scan the width of the line
        For l = 0 To 39
            If start + l > n Then Exit For
            If Mid$(s, start + l, 1) = vbLf Then Exit For
        Next
        x = (vid.width - l * 8) \ 2
        For j = 0 To l - 1
            Draw_Character x, y, AscW(Mid$(s, start + j, 1))
            x = x + 8
        Next

        y = y + 8

        Do While start <= n
            If Mid$(s, start, 1) = vbLf Then Exit Do
            start = start + 1
        Loop

        If start > n Then Exit Do
        start = start + 1                ' skip the \n
    Loop
End Sub

' Displays a text string in the center of the screen and waits for a Y or N
' keypress.
Public Function SCR_ModalMessage(ByVal text As String) As Boolean
    If cls.state = ca_dedicated Then
        SCR_ModalMessage = True
        Exit Function
    End If

    scr_notifystring = text

    ' draw a fresh screen
    scr_fullupdate = 0
    scr_drawdialog = True
    SCR_UpdateScreen
    scr_drawdialog = False

    S_ClearBuffer                        ' so dma doesn't loop current sound

    Do
        key_count = -1                   ' wait for a key down and up
        Sys_SendKeyEvents
        Sys_Sleep
    Loop While key_lastpress <> 121 And key_lastpress <> 110 And key_lastpress <> K_ESCAPE

    scr_fullupdate = 0
    SCR_UpdateScreen

    SCR_ModalMessage = (key_lastpress = 121)
End Function

'=============================================================================

' Brings the console down and fades the palettes back to normal
Public Sub SCR_BringDownConsole()
    Dim i As Long

    scr_centertime_off = 0

    For i = 0 To 19
        If scr_conlines = scr_con_current Then Exit For
        SCR_UpdateScreen
    Next

    cl.cshifts(0).percent = 0            ' no area contents palette on next frame
    VID_SetPalette host_basepal
End Sub

' This is called every frame, and can also be called explicitly to flush
' text to the screen.
Public Sub SCR_UpdateScreen()
    Dim vrect As vrect_t, t As Long

    If scr_skipupdate Or block_drawing Then Exit Sub

    scr_copytop = 0
    scr_copyeverything = 0

    If scr_disabled_for_loading Then
        If realtime - scr_disabled_time > 60 Then
            scr_disabled_for_loading = False
            Con_Printf "load failed.\n"
        Else
            Exit Sub
        End If
    End If

    If cls.state = ca_dedicated Then Exit Sub   ' stdout only

    If Not scr_initialized Or Not con_initialized Then Exit Sub   ' not initialized yet

    If cvars(cv_scr_viewsize).value <> oldscr_viewsize Then
        oldscr_viewsize = cvars(cv_scr_viewsize).value
        vid.recalc_refdef = 1
    End If

    '
    ' check for vid changes
    '
    If oldfov <> cvars(cv_scr_fov).value Then
        oldfov = cvars(cv_scr_fov).value
        vid.recalc_refdef = 1
    End If

    If oldlcd_x <> cvars(cv_lcd_x).value Then
        oldlcd_x = cvars(cv_lcd_x).value
        vid.recalc_refdef = 1
    End If

    If oldscreensize <> cvars(cv_scr_viewsize).value Then
        oldscreensize = cvars(cv_scr_viewsize).value
        vid.recalc_refdef = 1
    End If

    If vid.recalc_refdef Then
        ' something changed, so reorder the screen
        SCR_CalcRefdef
    End If

    '
    ' do 3D refresh drawing, and then update the screen
    '
    D_EnableBackBufferAccess             ' of all overlay stuff if drawing directly

    t = scr_fullupdate
    scr_fullupdate = scr_fullupdate + 1
    If t < vid.numpages Then
        ' clear the entire screen
        scr_copyeverything = 1
        Draw_TileClear 0, 0, vid.width, vid.height
        Sbar_Changed
    End If

    SCR_SetUpToDrawConsole
    SCR_EraseCenterString

    D_DisableBackBufferAccess            ' for adapters that can't stay mapped in
                                         ' for linear writes all the time

    VID_LockBuffer
    V_RenderView
    VID_UnlockBuffer

    D_EnableBackBufferAccess             ' of all overlay stuff if drawing directly

    If scr_drawdialog Then
        Sbar_Draw
        Draw_FadeScreen
        SCR_DrawNotifyString
        scr_copyeverything = 1
    ElseIf scr_drawloading Then
        SCR_DrawLoading_
        Sbar_Draw
    ElseIf cl.intermission = 1 And key_dest = key_game Then
        Sbar_IntermissionOverlay
    ElseIf cl.intermission = 2 And key_dest = key_game Then
        Sbar_FinaleOverlay
        SCR_CheckDrawCenterString
    ElseIf cl.intermission = 3 And key_dest = key_game Then
        SCR_CheckDrawCenterString
    Else
        SCR_DrawRam
        SCR_DrawNet
        SCR_DrawTurtle
        SCR_DrawPause
        SCR_CheckDrawCenterString
        Sbar_Draw
        SCR_DrawConsole
        M_Draw
    End If

    D_DisableBackBufferAccess            ' for adapters that can't stay mapped in
                                         ' for linear writes all the time
    ' (pconupdate is never set: D_UpdateRects is not called)

    V_UpdatePalette

    '
    ' update one of three areas
    '
    If scr_copyeverything Then
        vrect.x = 0
        vrect.y = 0
        vrect.width = vid.width
        vrect.height = vid.height
        VID_Update vrect
    ElseIf scr_copytop Then
        vrect.x = 0
        vrect.y = 0
        vrect.width = vid.width
        vrect.height = vid.height - sb_lines
        VID_Update vrect
    Else
        vrect.x = scr_vrect.x
        vrect.y = scr_vrect.y
        vrect.width = scr_vrect.width
        vrect.height = scr_vrect.height
        VID_Update vrect
    End If
End Sub

Public Sub SCR_UpdateWholeScreen()
    scr_fullupdate = 0
    SCR_UpdateScreen
End Sub
