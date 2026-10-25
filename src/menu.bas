Attribute VB_Name = "menu"
' menu.bas - the menus (menu.c)
'
' The qpic_t pointers are pics() indices. Not ported: the serial and modem
' configuration menus (WinQuake has no serial driver, so serialAvailable is
' never set and the net menu skips those entries) and the video modes menu
' (vid_menudrawfn: QuakeVB's window has no modes to list, so "Video
' Options" stays hidden, as in C without a video menu). The window is
' always a window: C's modestate is MS_WINDOWED, so "Use Mouse" is shown.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private vid_menudrawfn As Boolean        ' C: void (*vid_menudrawfn)(void), never set here

' m_state
Private Const m_none As Long = 0
Private Const m_main As Long = 1
Private Const m_singleplayer As Long = 2
Private Const m_load As Long = 3
Private Const m_save As Long = 4
Private Const m_multiplayer As Long = 5
Private Const m_setup As Long = 6
Private Const m_net As Long = 7
Private Const m_options As Long = 8
Private Const m_video As Long = 9
Private Const m_keys As Long = 10
Private Const m_help As Long = 11
Private Const m_quit As Long = 12
Private Const m_serialconfig As Long = 13
Private Const m_modemconfig As Long = 14
Private Const m_lanconfig As Long = 15
Private Const m_gameoptions As Long = 16
Private Const m_search As Long = 17
Private Const m_slist As Long = 18

Public m_state As Long                   ' (net_dgrm sets it when a connect fails)

Private m_entersound As Boolean          ' play after drawing a frame, so caching
                                         ' won't disrupt the sound
Private m_recursiveDraw As Boolean

Public m_return_state As Long
Public m_return_onerror As Boolean
Public m_return_reason As String         ' up to 31 characters

Private identityTable(0 To 255) As Byte
Private translationTable(0 To 255) As Byte

Private m_save_demonum As Long

' MAIN MENU
Private m_main_cursor As Long
Private Const MAIN_ITEMS As Long = 5

' SINGLE PLAYER MENU
Private m_singleplayer_cursor As Long
Private Const SINGLEPLAYER_ITEMS As Long = 3

' LOAD/SAVE MENU
Private load_cursor As Long              ' 0 < load_cursor < MAX_SAVEGAMES
Private Const MAX_SAVEGAMES As Long = 12
Private m_filenames(0 To MAX_SAVEGAMES - 1) As String   ' up to SAVEGAME_COMMENT_LENGTH characters
Private loadable(0 To MAX_SAVEGAMES - 1) As Boolean

' MULTIPLAYER MENU
Private m_multiplayer_cursor As Long
Private Const MULTIPLAYER_ITEMS As Long = 3

' SETUP MENU
Private setup_cursor As Long             ' = 4 (M_Init)
Private setup_hostname As String         ' up to 15 characters
Private setup_myname As String           ' up to 15 characters
Private setup_oldtop As Long
Private setup_oldbottom As Long
Private setup_top As Long
Private setup_bottom As Long
Private Const NUM_SETUP_CMDS As Long = 5

' NET MENU
Private m_net_cursor As Long
Private m_net_items As Long

' OPTIONS MENU
Private Const OPTIONS_ITEMS As Long = 14 ' (_WIN32)
Private Const SLIDER_RANGE As Long = 10
Private options_cursor As Long

' KEYS MENU
Private Const NUMCOMMANDS As Long = 18
Private bindnames(0 To NUMCOMMANDS - 1, 0 To 1) As String   ' (InitBindnames)
Private keys_cursor As Long
Private bind_grab As Boolean

' HELP MENU
Private help_page As Long
Private Const NUM_HELP_PAGES As Long = 6

' QUIT MENU
Private msgNumber As Long
Private m_quit_prevstate As Long
Private wasInMenus As Boolean

' LAN CONFIG MENU
Private lanConfig_cursor As Long         ' = -1 (M_Init)
Private Const NUM_LANCONFIG_CMDS As Long = 3
Private lanConfig_port As Long
Private lanConfig_portname As String     ' up to 5 characters
Private lanConfig_joinname As String     ' up to 21 characters

' GAME OPTIONS MENU
Private Type level_t
    name As String
    description As String
End Type

Private Type episode_t
    description As String
    firstLevel As Long
    levels As Long
End Type

Private levels() As level_t              ' (InitLevels)
Private hipnoticlevels() As level_t
Private roguelevels() As level_t
Private episodes() As episode_t
Private hipnoticepisodes() As episode_t
Private rogueepisodes() As episode_t

Private startepisode As Long
Private startlevel As Long
Private maxplayers As Long
Private m_serverInfoMessage As Boolean
Private m_serverInfoMessageTime As Double

Private gameoptions_cursor As Long
Private Const NUM_GAMEOPTIONS As Long = 9

' SEARCH MENU
Private searchComplete As Boolean
Private searchCompleteTime As Double

' SLIST MENU
Private slist_cursor As Long
Private slist_sorted As Boolean

' C: #define StartingGame (m_multiplayer_cursor == 1) and the others
Private Function StartingGame() As Boolean
    StartingGame = (m_multiplayer_cursor = 1)
End Function

Private Function JoiningGame() As Boolean
    JoiningGame = (m_multiplayer_cursor = 0)
End Function

Private Function IPXConfig() As Boolean
    IPXConfig = (m_net_cursor = 2)
End Function

Private Function TCPIPConfig() As Boolean
    TCPIPConfig = (m_net_cursor = 3)
End Function

' Draws one solid graphics character
Private Sub M_DrawCharacter(ByVal cx As Long, ByVal line_ As Long, ByVal num As Long)
    Draw_Character cx + ((vid.width - 320) \ 2), line_, num
End Sub

Private Sub M_Print(ByVal cx As Long, ByVal cy As Long, ByVal str_ As String)
    Dim i As Long

    For i = 1 To Len(str_)
        M_DrawCharacter cx, cy, AscW(Mid$(str_, i, 1)) + 128
        cx = cx + 8
    Next
End Sub

Private Sub M_PrintWhite(ByVal cx As Long, ByVal cy As Long, ByVal str_ As String)
    Dim i As Long

    For i = 1 To Len(str_)
        M_DrawCharacter cx, cy, AscW(Mid$(str_, i, 1))
        cx = cx + 8
    Next
End Sub

Private Sub M_DrawTransPic(ByVal x As Long, ByVal y As Long, ByVal pic As Long)
    Draw_TransPic x + ((vid.width - 320) \ 2), y, pic
End Sub

Public Sub M_DrawPic(ByVal x As Long, ByVal y As Long, ByVal pic As Long)
    Draw_Pic x + ((vid.width - 320) \ 2), y, pic
End Sub

Private Sub M_BuildTranslationTable(ByVal top As Long, ByVal bottom As Long)
    Dim j As Long

    For j = 0 To 255
        identityTable(j) = j
    Next
    For j = 0 To 255
        translationTable(j) = identityTable(j)
    Next

    If top < 128 Then                    ' the artists made some backwards ranges.  sigh.
        For j = 0 To 15
            translationTable(TOP_RANGE + j) = identityTable(top + j)
        Next
    Else
        For j = 0 To 15
            translationTable(TOP_RANGE + j) = identityTable(top + 15 - j)
        Next
    End If

    If bottom < 128 Then
        For j = 0 To 15
            translationTable(BOTTOM_RANGE + j) = identityTable(bottom + j)
        Next
    Else
        For j = 0 To 15
            translationTable(BOTTOM_RANGE + j) = identityTable(bottom + 15 - j)
        Next
    End If
End Sub

Private Sub M_DrawTransPicTranslate(ByVal x As Long, ByVal y As Long, ByVal pic As Long)
    Draw_TransPicTranslate x + ((vid.width - 320) \ 2), y, pic, translationTable
End Sub

Private Sub M_DrawTextBox(ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal lines As Long)
    Dim p As Long, cx As Long, cy As Long, n As Long

    ' draw left side
    cx = x
    cy = y
    p = Draw_CachePic("gfx/box_tl.lmp")
    M_DrawTransPic cx, cy, p
    p = Draw_CachePic("gfx/box_ml.lmp")
    For n = 0 To lines - 1
        cy = cy + 8
        M_DrawTransPic cx, cy, p
    Next
    p = Draw_CachePic("gfx/box_bl.lmp")
    M_DrawTransPic cx, cy + 8, p

    ' draw middle
    cx = cx + 8
    Do While width > 0
        cy = y
        p = Draw_CachePic("gfx/box_tm.lmp")
        M_DrawTransPic cx, cy, p
        p = Draw_CachePic("gfx/box_mm.lmp")
        For n = 0 To lines - 1
            cy = cy + 8
            If n = 1 Then p = Draw_CachePic("gfx/box_mm2.lmp")
            M_DrawTransPic cx, cy, p
        Next
        p = Draw_CachePic("gfx/box_bm.lmp")
        M_DrawTransPic cx, cy + 8, p
        width = width - 2
        cx = cx + 16
    Loop

    ' draw right side
    cy = y
    p = Draw_CachePic("gfx/box_tr.lmp")
    M_DrawTransPic cx, cy, p
    p = Draw_CachePic("gfx/box_mr.lmp")
    For n = 0 To lines - 1
        cy = cy + 8
        M_DrawTransPic cx, cy, p
    Next
    p = Draw_CachePic("gfx/box_br.lmp")
    M_DrawTransPic cx, cy + 8, p
End Sub

' (QuakeVB) C: 12+((int)(realtime*4)&1), the blinking cursor
Private Function BlinkChar(ByVal c As Long) As Long
    BlinkChar = c + (ToInt(realtime * 4) And 1)
End Function

' (QuakeVB) C: va("gfx/menudot%i.lmp", f+1) with f = (int)(host_time * 10)%6
Private Function MenuDot() As Long
    MenuDot = Draw_CachePic(va("gfx/menudot%i.lmp", (ToInt(host_time * 10) Mod 6) + 1))
End Function

'=============================================================================

Public Sub M_ToggleMenu_f()
    m_entersound = True

    If key_dest = key_menu Then
        If m_state <> m_main Then
            M_Menu_Main_f
            Exit Sub
        End If
        key_dest = key_game
        m_state = m_none
        Exit Sub
    End If
    If key_dest = key_console Then
        Con_ToggleConsole_f
    Else
        M_Menu_Main_f
    End If
End Sub

'=============================================================================
' MAIN MENU

Public Sub M_Menu_Main_f()
    If key_dest <> key_menu Then
        m_save_demonum = cls.demonum
        cls.demonum = -1
    End If
    key_dest = key_menu
    m_state = m_main
    m_entersound = True
End Sub

Private Sub M_Main_Draw()
    Dim p As Long

    M_DrawTransPic 16, 4, Draw_CachePic("gfx/qplaque.lmp")
    p = Draw_CachePic("gfx/ttl_main.lmp")
    M_DrawPic (320 - pics(p).width) \ 2, 4, p
    M_DrawTransPic 72, 32, Draw_CachePic("gfx/mainmenu.lmp")

    M_DrawTransPic 54, 32 + m_main_cursor * 20, MenuDot()
End Sub

Private Sub M_Main_Key(ByVal key As Long)
    Select Case key
    Case K_ESCAPE
        key_dest = key_game
        m_state = m_none
        cls.demonum = m_save_demonum
        If cls.demonum <> -1 And Not cls.demoplayback And cls.state <> ca_connected Then CL_NextDemo

    Case K_DOWNARROW
        S_LocalSound "misc/menu1.wav"
        m_main_cursor = m_main_cursor + 1
        If m_main_cursor >= MAIN_ITEMS Then m_main_cursor = 0

    Case K_UPARROW
        S_LocalSound "misc/menu1.wav"
        m_main_cursor = m_main_cursor - 1
        If m_main_cursor < 0 Then m_main_cursor = MAIN_ITEMS - 1

    Case K_ENTER
        m_entersound = True

        Select Case m_main_cursor
        Case 0: M_Menu_SinglePlayer_f
        Case 1: M_Menu_MultiPlayer_f
        Case 2: M_Menu_Options_f
        Case 3: M_Menu_Help_f
        Case 4: M_Menu_Quit_f
        End Select
    End Select
End Sub

'=============================================================================
' SINGLE PLAYER MENU

Private Sub M_Menu_SinglePlayer_f()
    key_dest = key_menu
    m_state = m_singleplayer
    m_entersound = True
End Sub

Private Sub M_SinglePlayer_Draw()
    Dim p As Long

    M_DrawTransPic 16, 4, Draw_CachePic("gfx/qplaque.lmp")
    p = Draw_CachePic("gfx/ttl_sgl.lmp")
    M_DrawPic (320 - pics(p).width) \ 2, 4, p
    M_DrawTransPic 72, 32, Draw_CachePic("gfx/sp_menu.lmp")

    M_DrawTransPic 54, 32 + m_singleplayer_cursor * 20, MenuDot()
End Sub

Private Sub M_SinglePlayer_Key(ByVal key As Long)
    Select Case key
    Case K_ESCAPE
        M_Menu_Main_f

    Case K_DOWNARROW
        S_LocalSound "misc/menu1.wav"
        m_singleplayer_cursor = m_singleplayer_cursor + 1
        If m_singleplayer_cursor >= SINGLEPLAYER_ITEMS Then m_singleplayer_cursor = 0

    Case K_UPARROW
        S_LocalSound "misc/menu1.wav"
        m_singleplayer_cursor = m_singleplayer_cursor - 1
        If m_singleplayer_cursor < 0 Then m_singleplayer_cursor = SINGLEPLAYER_ITEMS - 1

    Case K_ENTER
        m_entersound = True

        Select Case m_singleplayer_cursor
        Case 0
            If sv.active Then
                If Not SCR_ModalMessage("Are you sure you want to" & vbLf & "start a new game?" & vbLf) Then Exit Sub
            End If
            key_dest = key_game
            If sv.active Then Cbuf_AddText "disconnect" & vbLf
            Cbuf_AddText "maxplayers 1" & vbLf
            Cbuf_AddText "map start" & vbLf

        Case 1
            M_Menu_Load_f

        Case 2
            M_Menu_Save_f
        End Select
    End Select
End Sub

'=============================================================================
' LOAD/SAVE MENU

Private Sub M_ScanSaves()
    Dim i As Long, name As String, h As Long, l As Long, version As Long, data() As Byte, p As Long

    For i = 0 To MAX_SAVEGAMES - 1
        m_filenames(i) = "--- UNUSED SLOT ---"
        loadable(i) = False
        name = va("%s/s%i.sav", com_gamedir, i)
        l = Sys_FileOpenRead(name, h)
        If h = -1 Then GoTo nexti
        ' (QuakeVB: the header is read into a buffer and scanned as fscanf
        ' does: "%i\n" for the version, "%79s\n" for the comment)
        If l > 256 Then l = 256
        ReDim data(0 To l)
        If l > 0 Then Sys_FileRead h, data, 0, l
        data(l) = 0
        p = 0
        version = C_atoi(ScanWord(data, p))
        name = Left$(ScanWord(data, p), 79)
        m_filenames(i) = Left$(name, SAVEGAME_COMMENT_LENGTH)

        ' change _ back to space
        m_filenames(i) = Replace(m_filenames(i), "_", " ")
        loadable(i) = True
        Sys_FileClose h
nexti:
    Next
End Sub

Private Sub M_Menu_Load_f()
    m_entersound = True
    m_state = m_load
    key_dest = key_menu
    M_ScanSaves
End Sub

Private Sub M_Menu_Save_f()
    If Not sv.active Then Exit Sub
    If cl.intermission Then Exit Sub
    If svs.maxclients <> 1 Then Exit Sub
    m_entersound = True
    m_state = m_save
    key_dest = key_menu
    M_ScanSaves
End Sub

Private Sub M_Load_Draw()
    Dim i As Long, p As Long

    p = Draw_CachePic("gfx/p_load.lmp")
    M_DrawPic (320 - pics(p).width) \ 2, 4, p

    For i = 0 To MAX_SAVEGAMES - 1
        M_Print 16, 32 + 8 * i, m_filenames(i)
    Next

    ' line cursor
    M_DrawCharacter 8, 32 + load_cursor * 8, BlinkChar(12)
End Sub

Private Sub M_Save_Draw()
    Dim i As Long, p As Long

    p = Draw_CachePic("gfx/p_save.lmp")
    M_DrawPic (320 - pics(p).width) \ 2, 4, p

    For i = 0 To MAX_SAVEGAMES - 1
        M_Print 16, 32 + 8 * i, m_filenames(i)
    Next

    ' line cursor
    M_DrawCharacter 8, 32 + load_cursor * 8, BlinkChar(12)
End Sub

Private Sub M_Load_Key(ByVal k As Long)
    Select Case k
    Case K_ESCAPE
        M_Menu_SinglePlayer_f

    Case K_ENTER
        S_LocalSound "misc/menu2.wav"
        If Not loadable(load_cursor) Then Exit Sub
        m_state = m_none
        key_dest = key_game

        ' Host_Loadgame_f can't bring up the loading plaque because too much
        ' stack space has been used, so do it now
        SCR_BeginLoadingPlaque

        ' issue the load command
        Cbuf_AddText va("load s%i\n", load_cursor)
        Exit Sub

    Case K_UPARROW, K_LEFTARROW
        S_LocalSound "misc/menu1.wav"
        load_cursor = load_cursor - 1
        If load_cursor < 0 Then load_cursor = MAX_SAVEGAMES - 1

    Case K_DOWNARROW, K_RIGHTARROW
        S_LocalSound "misc/menu1.wav"
        load_cursor = load_cursor + 1
        If load_cursor >= MAX_SAVEGAMES Then load_cursor = 0
    End Select
End Sub

Private Sub M_Save_Key(ByVal k As Long)
    Select Case k
    Case K_ESCAPE
        M_Menu_SinglePlayer_f

    Case K_ENTER
        m_state = m_none
        key_dest = key_game
        Cbuf_AddText va("save s%i\n", load_cursor)
        Exit Sub

    Case K_UPARROW, K_LEFTARROW
        S_LocalSound "misc/menu1.wav"
        load_cursor = load_cursor - 1
        If load_cursor < 0 Then load_cursor = MAX_SAVEGAMES - 1

    Case K_DOWNARROW, K_RIGHTARROW
        S_LocalSound "misc/menu1.wav"
        load_cursor = load_cursor + 1
        If load_cursor >= MAX_SAVEGAMES Then load_cursor = 0
    End Select
End Sub

'=============================================================================
' MULTIPLAYER MENU

Private Sub M_Menu_MultiPlayer_f()
    key_dest = key_menu
    m_state = m_multiplayer
    m_entersound = True
End Sub

Private Sub M_MultiPlayer_Draw()
    Dim p As Long

    M_DrawTransPic 16, 4, Draw_CachePic("gfx/qplaque.lmp")
    p = Draw_CachePic("gfx/p_multi.lmp")
    M_DrawPic (320 - pics(p).width) \ 2, 4, p
    M_DrawTransPic 72, 32, Draw_CachePic("gfx/mp_menu.lmp")

    M_DrawTransPic 54, 32 + m_multiplayer_cursor * 20, MenuDot()

    If serialAvailable Or ipxAvailable Or tcpipAvailable Then Exit Sub
    M_PrintWhite (320 \ 2) - ((27 * 8) \ 2), 148, "No Communications Available"
End Sub

Private Sub M_MultiPlayer_Key(ByVal key As Long)
    Select Case key
    Case K_ESCAPE
        M_Menu_Main_f

    Case K_DOWNARROW
        S_LocalSound "misc/menu1.wav"
        m_multiplayer_cursor = m_multiplayer_cursor + 1
        If m_multiplayer_cursor >= MULTIPLAYER_ITEMS Then m_multiplayer_cursor = 0

    Case K_UPARROW
        S_LocalSound "misc/menu1.wav"
        m_multiplayer_cursor = m_multiplayer_cursor - 1
        If m_multiplayer_cursor < 0 Then m_multiplayer_cursor = MULTIPLAYER_ITEMS - 1

    Case K_ENTER
        m_entersound = True
        Select Case m_multiplayer_cursor
        Case 0
            If serialAvailable Or ipxAvailable Or tcpipAvailable Then M_Menu_Net_f

        Case 1
            If serialAvailable Or ipxAvailable Or tcpipAvailable Then M_Menu_Net_f

        Case 2
            M_Menu_Setup_f
        End Select
    End Select
End Sub

'=============================================================================
' SETUP MENU

Private Sub M_Menu_Setup_f()
    key_dest = key_menu
    m_state = m_setup
    m_entersound = True
    ' (C copies into char[16] buffers)
    setup_myname = Left$(cvars(cv_cl_name).strval, 15)
    setup_hostname = Left$(cvars(cv_hostname).strval, 15)
    setup_oldtop = ToInt(cvars(cv_cl_color).value) \ 16
    setup_top = setup_oldtop
    setup_oldbottom = ToInt(cvars(cv_cl_color).value) And 15
    setup_bottom = setup_oldbottom
End Sub

Private Sub M_Setup_Draw()
    Dim p As Long

    M_DrawTransPic 16, 4, Draw_CachePic("gfx/qplaque.lmp")
    p = Draw_CachePic("gfx/p_multi.lmp")
    M_DrawPic (320 - pics(p).width) \ 2, 4, p

    M_Print 64, 40, "Hostname"
    M_DrawTextBox 160, 32, 16, 1
    M_Print 168, 40, setup_hostname

    M_Print 64, 56, "Your name"
    M_DrawTextBox 160, 48, 16, 1
    M_Print 168, 56, setup_myname

    M_Print 64, 80, "Shirt color"
    M_Print 64, 104, "Pants color"

    M_DrawTextBox 64, 140 - 8, 14, 1
    M_Print 72, 140, "Accept Changes"

    p = Draw_CachePic("gfx/bigbox.lmp")
    M_DrawTransPic 160, 64, p
    p = Draw_CachePic("gfx/menuplyr.lmp")
    M_BuildTranslationTable setup_top * 16, setup_bottom * 16
    M_DrawTransPicTranslate 172, 72, p

    M_DrawCharacter 56, SetupCursorY(setup_cursor), BlinkChar(12)

    If setup_cursor = 0 Then M_DrawCharacter 168 + 8 * Len(setup_hostname), SetupCursorY(setup_cursor), BlinkChar(10)

    If setup_cursor = 1 Then M_DrawCharacter 168 + 8 * Len(setup_myname), SetupCursorY(setup_cursor), BlinkChar(10)
End Sub

' C: int setup_cursor_table[] = {40, 56, 80, 104, 140}
Private Function SetupCursorY(ByVal i As Long) As Long
    Select Case i
    Case 0: SetupCursorY = 40
    Case 1: SetupCursorY = 56
    Case 2: SetupCursorY = 80
    Case 3: SetupCursorY = 104
    Case 4: SetupCursorY = 140
    End Select
End Function

' (QuakeVB) the "forward" step of M_Setup_Key (C jumps to it from K_ENTER)
Private Sub SetupForward()
    S_LocalSound "misc/menu3.wav"
    If setup_cursor = 2 Then setup_top = setup_top + 1
    If setup_cursor = 3 Then setup_bottom = setup_bottom + 1
End Sub

Private Sub M_Setup_Key(ByVal k As Long)
    Select Case k
    Case K_ESCAPE
        M_Menu_MultiPlayer_f

    Case K_UPARROW
        S_LocalSound "misc/menu1.wav"
        setup_cursor = setup_cursor - 1
        If setup_cursor < 0 Then setup_cursor = NUM_SETUP_CMDS - 1

    Case K_DOWNARROW
        S_LocalSound "misc/menu1.wav"
        setup_cursor = setup_cursor + 1
        If setup_cursor >= NUM_SETUP_CMDS Then setup_cursor = 0

    Case K_LEFTARROW
        If setup_cursor < 2 Then Exit Sub
        S_LocalSound "misc/menu3.wav"
        If setup_cursor = 2 Then setup_top = setup_top - 1
        If setup_cursor = 3 Then setup_bottom = setup_bottom - 1

    Case K_RIGHTARROW
        If setup_cursor < 2 Then Exit Sub
        SetupForward

    Case K_ENTER
        If setup_cursor = 0 Or setup_cursor = 1 Then Exit Sub

        If setup_cursor = 2 Or setup_cursor = 3 Then
            SetupForward
        Else
            ' setup_cursor == 4 (OK)
            If cvars(cv_cl_name).strval <> setup_myname Then Cbuf_AddText va("name ""%s""\n", setup_myname)
            If cvars(cv_hostname).strval <> setup_hostname Then Cvar_Set "hostname", setup_hostname
            If setup_top <> setup_oldtop Or setup_bottom <> setup_oldbottom Then
                Cbuf_AddText va("color %i %i\n", setup_top, setup_bottom)
            End If
            m_entersound = True
            M_Menu_MultiPlayer_f
        End If

    Case K_BACKSPACE
        If setup_cursor = 0 Then
            If Len(setup_hostname) Then setup_hostname = Left$(setup_hostname, Len(setup_hostname) - 1)
        End If

        If setup_cursor = 1 Then
            If Len(setup_myname) Then setup_myname = Left$(setup_myname, Len(setup_myname) - 1)
        End If

    Case Else
        If k < 32 Or k > 127 Then GoTo done
        If setup_cursor = 0 Then
            If Len(setup_hostname) < 15 Then setup_hostname = setup_hostname & ChrW$(k)
        End If
        If setup_cursor = 1 Then
            If Len(setup_myname) < 15 Then setup_myname = setup_myname & ChrW$(k)
        End If
    End Select
done:

    If setup_top > 13 Then setup_top = 0
    If setup_top < 0 Then setup_top = 13
    If setup_bottom > 13 Then setup_bottom = 0
    If setup_bottom < 0 Then setup_bottom = 13
End Sub

'=============================================================================
' NET MENU

Private Function NetHelpMessage(ByVal i As Long) As String
    Select Case i
    '                         .........1.........2....
    Case 0, 3, 4, 7: NetHelpMessage = "                        "
    Case 1, 5: NetHelpMessage = " Two computers connected"
    Case 2: NetHelpMessage = "   through two modems.  "
    Case 6: NetHelpMessage = " by a null-modem cable. "
    Case 8: NetHelpMessage = " Novell network LANs    "
    Case 9: NetHelpMessage = " or Windows 95 DOS-box. "
    Case 10: NetHelpMessage = "                        "
    Case 11: NetHelpMessage = "(LAN=Local Area Network)"
    Case 12: NetHelpMessage = " Commonly used to play  "
    Case 13: NetHelpMessage = " over the Internet, but "
    Case 14: NetHelpMessage = " also used on a Local   "
    Case 15: NetHelpMessage = " Area Network.          "
    End Select
End Function

Private Sub M_Menu_Net_f()
    key_dest = key_menu
    m_state = m_net
    m_entersound = True
    m_net_items = 4

    If m_net_cursor >= m_net_items Then m_net_cursor = 0
    m_net_cursor = m_net_cursor - 1
    M_Net_Key K_DOWNARROW
End Sub

Private Sub M_Net_Draw()
    Dim f As Long, p As Long

    M_DrawTransPic 16, 4, Draw_CachePic("gfx/qplaque.lmp")
    p = Draw_CachePic("gfx/p_multi.lmp")
    M_DrawPic (320 - pics(p).width) \ 2, 4, p

    f = 32

    If serialAvailable Then
        p = Draw_CachePic("gfx/netmen1.lmp")
    Else
        p = 0                            ' (_WIN32)
    End If

    If p Then M_DrawTransPic 72, f, p

    f = f + 19

    If serialAvailable Then
        p = Draw_CachePic("gfx/netmen2.lmp")
    Else
        p = 0                            ' (_WIN32)
    End If

    If p Then M_DrawTransPic 72, f, p

    f = f + 19
    If ipxAvailable Then
        p = Draw_CachePic("gfx/netmen3.lmp")
    Else
        p = Draw_CachePic("gfx/dim_ipx.lmp")
    End If
    M_DrawTransPic 72, f, p

    f = f + 19
    If tcpipAvailable Then
        p = Draw_CachePic("gfx/netmen4.lmp")
    Else
        p = Draw_CachePic("gfx/dim_tcp.lmp")
    End If
    M_DrawTransPic 72, f, p

    If m_net_items = 5 Then              ' JDC, could just be removed
        f = f + 19
        p = Draw_CachePic("gfx/netmen5.lmp")
        M_DrawTransPic 72, f, p
    End If

    f = (320 - 26 * 8) \ 2
    M_DrawTextBox f, 134, 24, 4
    f = f + 8
    M_Print f, 142, NetHelpMessage(m_net_cursor * 4 + 0)
    M_Print f, 150, NetHelpMessage(m_net_cursor * 4 + 1)
    M_Print f, 158, NetHelpMessage(m_net_cursor * 4 + 2)
    M_Print f, 166, NetHelpMessage(m_net_cursor * 4 + 3)

    M_DrawTransPic 54, 32 + m_net_cursor * 20, MenuDot()
End Sub

Private Sub M_Net_Key(ByVal k As Long)
    Do                                   ' C: again:
        Select Case k
        Case K_ESCAPE
            M_Menu_MultiPlayer_f

        Case K_DOWNARROW
            S_LocalSound "misc/menu1.wav"
            m_net_cursor = m_net_cursor + 1
            If m_net_cursor >= m_net_items Then m_net_cursor = 0

        Case K_UPARROW
            S_LocalSound "misc/menu1.wav"
            m_net_cursor = m_net_cursor - 1
            If m_net_cursor < 0 Then m_net_cursor = m_net_items - 1

        Case K_ENTER
            m_entersound = True

            Select Case m_net_cursor
            Case 0, 1
                ' (M_Menu_SerialConfig_f: never reached, serialAvailable is
                ' always false)
            Case 2
                M_Menu_LanConfig_f
            Case 3
                M_Menu_LanConfig_f
            Case 4
                ' multiprotocol
            End Select
        End Select

        If m_net_cursor = 0 And Not serialAvailable Then GoTo again
        If m_net_cursor = 1 And Not serialAvailable Then GoTo again
        If m_net_cursor = 2 And Not ipxAvailable Then GoTo again
        If m_net_cursor = 3 And Not tcpipAvailable Then GoTo again
        Exit Do
again:
    Loop
End Sub

'=============================================================================
' OPTIONS MENU

Private Sub M_Menu_Options_f()
    key_dest = key_menu
    m_state = m_options
    m_entersound = True

    ' (C resets options_cursor 13 when not windowed; the window always is)
End Sub

Private Sub M_AdjustSliders(ByVal dir_ As Long)
    S_LocalSound "misc/menu3.wav"

    Select Case options_cursor
    Case 3                               ' screen size
        cvars(cv_scr_viewsize).value = cvars(cv_scr_viewsize).value + dir_ * 10
        If cvars(cv_scr_viewsize).value < 30 Then cvars(cv_scr_viewsize).value = 30
        If cvars(cv_scr_viewsize).value > 120 Then cvars(cv_scr_viewsize).value = 120
        Cvar_SetValue "viewsize", cvars(cv_scr_viewsize).value

    Case 4                               ' gamma
        cvars(cv_v_gamma).value = cvars(cv_v_gamma).value - dir_ * 0.05
        If cvars(cv_v_gamma).value < 0.5 Then cvars(cv_v_gamma).value = 0.5
        If cvars(cv_v_gamma).value > 1 Then cvars(cv_v_gamma).value = 1
        Cvar_SetValue "gamma", cvars(cv_v_gamma).value

    Case 5                               ' mouse speed
        cvars(cv_sensitivity).value = cvars(cv_sensitivity).value + dir_ * 0.5
        If cvars(cv_sensitivity).value < 1 Then cvars(cv_sensitivity).value = 1
        If cvars(cv_sensitivity).value > 11 Then cvars(cv_sensitivity).value = 11
        Cvar_SetValue "sensitivity", cvars(cv_sensitivity).value

    Case 6                               ' music volume
        cvars(cv_bgmvolume).value = cvars(cv_bgmvolume).value + dir_ * 1#   ' (_WIN32)
        If cvars(cv_bgmvolume).value < 0 Then cvars(cv_bgmvolume).value = 0
        If cvars(cv_bgmvolume).value > 1 Then cvars(cv_bgmvolume).value = 1
        Cvar_SetValue "bgmvolume", cvars(cv_bgmvolume).value

    Case 7                               ' sfx volume
        cvars(cv_volume).value = cvars(cv_volume).value + dir_ * 0.1
        If cvars(cv_volume).value < 0 Then cvars(cv_volume).value = 0
        If cvars(cv_volume).value > 1 Then cvars(cv_volume).value = 1
        Cvar_SetValue "volume", cvars(cv_volume).value

    Case 8                               ' allways run
        If cvars(cv_cl_forwardspeed).value > 200 Then
            Cvar_SetValue "cl_forwardspeed", 200
            Cvar_SetValue "cl_backspeed", 200
        Else
            Cvar_SetValue "cl_forwardspeed", 400
            Cvar_SetValue "cl_backspeed", 400
        End If

    Case 9                               ' invert mouse
        Cvar_SetValue "m_pitch", -cvars(cv_m_pitch).value

    Case 10                              ' lookspring
        Cvar_SetValue "lookspring", -(cvars(cv_lookspring).value = 0)

    Case 11                              ' lookstrafe
        Cvar_SetValue "lookstrafe", -(cvars(cv_lookstrafe).value = 0)

    Case 13                              ' _windowed_mouse
        Cvar_SetValue "_windowed_mouse", -(cvars(cv_windowed_mouse).value = 0)
    End Select
End Sub

Private Sub M_DrawSlider(ByVal x As Long, ByVal y As Long, ByVal range As Single)
    Dim i As Long

    If range < 0 Then range = 0
    If range > 1 Then range = 1
    M_DrawCharacter x - 8, y, 128
    For i = 0 To SLIDER_RANGE - 1
        M_DrawCharacter x + i * 8, y, 129
    Next
    M_DrawCharacter x + i * 8, y, 130
    M_DrawCharacter ToInt(x + (SLIDER_RANGE - 1) * 8 * range), y, 131
End Sub

Private Sub M_DrawCheckbox(ByVal x As Long, ByVal y As Long, ByVal on_ As Long)
    If on_ Then
        M_Print x, y, "on"
    Else
        M_Print x, y, "off"
    End If
End Sub

Private Sub M_Options_Draw()
    Dim r As Single, p As Long

    M_DrawTransPic 16, 4, Draw_CachePic("gfx/qplaque.lmp")
    p = Draw_CachePic("gfx/p_option.lmp")
    M_DrawPic (320 - pics(p).width) \ 2, 4, p

    M_Print 16, 32, "    Customize controls"
    M_Print 16, 40, "         Go to console"
    M_Print 16, 48, "     Reset to defaults"

    M_Print 16, 56, "           Screen size"
    r = (cvars(cv_scr_viewsize).value - 30) / (120 - 30)
    M_DrawSlider 220, 56, r

    M_Print 16, 64, "            Brightness"
    r = (1# - cvars(cv_v_gamma).value) / 0.5
    M_DrawSlider 220, 64, r

    M_Print 16, 72, "           Mouse Speed"
    r = (cvars(cv_sensitivity).value - 1) / 10
    M_DrawSlider 220, 72, r

    M_Print 16, 80, "       CD Music Volume"
    r = cvars(cv_bgmvolume).value
    M_DrawSlider 220, 80, r

    M_Print 16, 88, "          Sound Volume"
    r = cvars(cv_volume).value
    M_DrawSlider 220, 88, r

    M_Print 16, 96, "            Always Run"
    M_DrawCheckbox 220, 96, -(cvars(cv_cl_forwardspeed).value > 200)

    M_Print 16, 104, "          Invert Mouse"
    M_DrawCheckbox 220, 104, -(cvars(cv_m_pitch).value < 0)

    M_Print 16, 112, "            Lookspring"
    M_DrawCheckbox 220, 112, ToInt(cvars(cv_lookspring).value)

    M_Print 16, 120, "            Lookstrafe"
    M_DrawCheckbox 220, 120, ToInt(cvars(cv_lookstrafe).value)

    If vid_menudrawfn Then M_Print 16, 128, "         Video Options"

    ' (modestate == MS_WINDOWED)
    M_Print 16, 136, "             Use Mouse"
    M_DrawCheckbox 220, 136, ToInt(cvars(cv_windowed_mouse).value)

    ' cursor
    M_DrawCharacter 200, 32 + options_cursor * 8, BlinkChar(12)
End Sub

Private Sub M_Options_Key(ByVal k As Long)
    Select Case k
    Case K_ESCAPE
        M_Menu_Main_f

    Case K_ENTER
        m_entersound = True
        Select Case options_cursor
        Case 0
            M_Menu_Keys_f
        Case 1
            m_state = m_none
            Con_ToggleConsole_f
        Case 2
            Cbuf_AddText "exec default.cfg" & vbLf
        Case 12
            M_Menu_Video_f
        Case Else
            M_AdjustSliders 1
        End Select
        Exit Sub

    Case K_UPARROW
        S_LocalSound "misc/menu1.wav"
        options_cursor = options_cursor - 1
        If options_cursor < 0 Then options_cursor = OPTIONS_ITEMS - 1

    Case K_DOWNARROW
        S_LocalSound "misc/menu1.wav"
        options_cursor = options_cursor + 1
        If options_cursor >= OPTIONS_ITEMS Then options_cursor = 0

    Case K_LEFTARROW
        M_AdjustSliders -1

    Case K_RIGHTARROW
        M_AdjustSliders 1
    End Select

    If options_cursor = 12 And Not vid_menudrawfn Then
        If k = K_UPARROW Then
            options_cursor = 11
        Else
            options_cursor = 0
        End If
    End If

    ' (C moves off option 13 when not windowed; the window always is)
End Sub

'=============================================================================
' KEYS MENU

Private Sub InitBindnames()
    Dim a As Variant, i As Long

    a = Array("+attack", "attack", _
        "impulse 10", "change weapon", _
        "+jump", "jump / swim up", _
        "+forward", "walk forward", _
        "+back", "backpedal", _
        "+left", "turn left", _
        "+right", "turn right", _
        "+speed", "run", _
        "+moveleft", "step left", _
        "+moveright", "step right", _
        "+strafe", "sidestep", _
        "+lookup", "look up", _
        "+lookdown", "look down", _
        "centerview", "center view", _
        "+mlook", "mouse look", _
        "+klook", "keyboard look", _
        "+moveup", "swim up", _
        "+movedown", "swim down")
    For i = 0 To NUMCOMMANDS - 1
        bindnames(i, 0) = a(i * 2)
        bindnames(i, 1) = a(i * 2 + 1)
    Next
End Sub

Private Sub M_Menu_Keys_f()
    key_dest = key_menu
    m_state = m_keys
    m_entersound = True
End Sub

Private Sub M_FindKeysForCommand(ByVal command As String, twokeys() As Long)
    Dim count As Long, j As Long, l As Long, b As String

    twokeys(0) = -1
    twokeys(1) = -1
    l = Len(command)
    count = 0

    For j = 0 To 255
        b = keybindings(j)
        If Len(b) = 0 Then GoTo nextj
        If Left$(b, l) = command Then
            twokeys(count) = j
            count = count + 1
            If count = 2 Then Exit For
        End If
nextj:
    Next
End Sub

Private Sub M_UnbindCommand(ByVal command As String)
    Dim j As Long, l As Long, b As String

    l = Len(command)

    For j = 0 To 255
        b = keybindings(j)
        If Len(b) = 0 Then GoTo nextj
        If Left$(b, l) = command Then Key_SetBinding j, ""
nextj:
    Next
End Sub

Private Sub M_Keys_Draw()
    Dim i As Long, keys(0 To 1) As Long, name As String, x As Long, y As Long, p As Long

    p = Draw_CachePic("gfx/ttl_cstm.lmp")
    M_DrawPic (320 - pics(p).width) \ 2, 4, p

    If bind_grab Then
        M_Print 12, 32, "Press a key or button for this action"
    Else
        M_Print 18, 32, "Enter to change, backspace to clear"
    End If

    ' search for known bindings
    For i = 0 To NUMCOMMANDS - 1
        y = 48 + 8 * i

        M_Print 16, y, bindnames(i, 1)

        M_FindKeysForCommand bindnames(i, 0), keys

        If keys(0) = -1 Then
            M_Print 140, y, "???"
        Else
            name = Key_KeynumToString(keys(0))
            M_Print 140, y, name
            x = Len(name) * 8
            If keys(1) <> -1 Then
                M_Print 140 + x + 8, y, "or"
                M_Print 140 + x + 32, y, Key_KeynumToString(keys(1))
            End If
        End If
    Next

    If bind_grab Then
        M_DrawCharacter 130, 48 + keys_cursor * 8, 61   ' '='
    Else
        M_DrawCharacter 130, 48 + keys_cursor * 8, BlinkChar(12)
    End If
End Sub

Private Sub M_Keys_Key(ByVal k As Long)
    Dim keys(0 To 1) As Long

    If bind_grab Then
        ' defining a key
        S_LocalSound "misc/menu1.wav"
        If k = K_ESCAPE Then
            bind_grab = False
        ElseIf k <> 96 Then              ' '`'
            Cbuf_InsertText va("bind ""%s"" ""%s""\n", Key_KeynumToString(k), bindnames(keys_cursor, 0))
        End If

        bind_grab = False
        Exit Sub
    End If

    Select Case k
    Case K_ESCAPE
        M_Menu_Options_f

    Case K_LEFTARROW, K_UPARROW
        S_LocalSound "misc/menu1.wav"
        keys_cursor = keys_cursor - 1
        If keys_cursor < 0 Then keys_cursor = NUMCOMMANDS - 1

    Case K_DOWNARROW, K_RIGHTARROW
        S_LocalSound "misc/menu1.wav"
        keys_cursor = keys_cursor + 1
        If keys_cursor >= NUMCOMMANDS Then keys_cursor = 0

    Case K_ENTER                         ' go into bind mode
        M_FindKeysForCommand bindnames(keys_cursor, 0), keys
        S_LocalSound "misc/menu2.wav"
        If keys(1) <> -1 Then M_UnbindCommand bindnames(keys_cursor, 0)
        bind_grab = True

    Case K_BACKSPACE, K_DEL              ' delete bindings
        S_LocalSound "misc/menu2.wav"
        M_UnbindCommand bindnames(keys_cursor, 0)
    End Select
End Sub

'=============================================================================
' VIDEO MENU

Private Sub M_Menu_Video_f()
    key_dest = key_menu
    m_state = m_video
    m_entersound = True
End Sub

Private Sub M_Video_Draw()
    ' (C: (*vid_menudrawfn) (); there is no video menu)
End Sub

Private Sub M_Video_Key(ByVal key As Long)
    ' (C: (*vid_menukeyfn) (key); there is no video menu)
End Sub

'=============================================================================
' HELP MENU

Private Sub M_Menu_Help_f()
    key_dest = key_menu
    m_state = m_help
    m_entersound = True
    help_page = 0
End Sub

Private Sub M_Help_Draw()
    M_DrawPic 0, 0, Draw_CachePic(va("gfx/help%i.lmp", help_page))
End Sub

Private Sub M_Help_Key(ByVal key As Long)
    Select Case key
    Case K_ESCAPE
        M_Menu_Main_f

    Case K_UPARROW, K_RIGHTARROW
        m_entersound = True
        help_page = help_page + 1
        If help_page >= NUM_HELP_PAGES Then help_page = 0

    Case K_DOWNARROW, K_LEFTARROW
        m_entersound = True
        help_page = help_page - 1
        If help_page < 0 Then help_page = NUM_HELP_PAGES - 1
    End Select
End Sub

'=============================================================================
' QUIT MENU

Public Sub M_Menu_Quit_f()
    If m_state = m_quit Then Exit Sub
    wasInMenus = (key_dest = key_menu)
    key_dest = key_menu
    m_quit_prevstate = m_state
    m_state = m_quit
    m_entersound = True
    msgNumber = rand() And 7
End Sub

Private Sub M_Quit_Key(ByVal key As Long)
    Select Case key
    Case K_ESCAPE, 110, 78               ' 'n', 'N'
        If wasInMenus Then
            m_state = m_quit_prevstate
            m_entersound = True
        Else
            key_dest = key_game
            m_state = m_none
        End If

    Case 89, 121                         ' 'Y', 'y'
        key_dest = key_console
        Host_Quit_f
    End Select
End Sub

Private Sub M_Quit_Draw()
    If wasInMenus Then
        m_state = m_quit_prevstate
        m_recursiveDraw = True
        M_Draw
        m_state = m_quit
    End If

    ' (_WIN32)
    M_DrawTextBox 0, 0, 38, 23
    M_PrintWhite 16, 12, "  Quake version 1.09 by id Software" & vbLf & vbLf
    M_PrintWhite 16, 28, "Programming        Art " & vbLf
    M_Print 16, 36, " John Carmack       Adrian Carmack" & vbLf
    M_Print 16, 44, " Michael Abrash     Kevin Cloud" & vbLf
    M_Print 16, 52, " John Cash          Paul Steed" & vbLf
    M_Print 16, 60, " Dave 'Zoid' Kirsch" & vbLf
    M_PrintWhite 16, 68, "Design             Biz" & vbLf
    M_Print 16, 76, " John Romero        Jay Wilbur" & vbLf
    M_Print 16, 84, " Sandy Petersen     Mike Wilson" & vbLf
    M_Print 16, 92, " American McGee     Donna Jackson" & vbLf
    M_Print 16, 100, " Tim Willits        Todd Hollenshead" & vbLf
    M_PrintWhite 16, 108, "Support            Projects" & vbLf
    M_Print 16, 116, " Barrett Alexander  Shawn Green" & vbLf
    M_PrintWhite 16, 124, "Sound Effects" & vbLf
    M_Print 16, 132, " Trent Reznor and Nine Inch Nails" & vbLf & vbLf
    M_PrintWhite 16, 140, "Quake is a trademark of Id Software," & vbLf
    M_PrintWhite 16, 148, "inc., (c)1996 Id Software, inc. All" & vbLf
    M_PrintWhite 16, 156, "rights reserved. NIN logo is a" & vbLf
    M_PrintWhite 16, 164, "registered trademark licensed to" & vbLf
    M_PrintWhite 16, 172, "Nothing Interactive, Inc. All rights" & vbLf
    M_PrintWhite 16, 180, "reserved. Press y to exit" & vbLf
End Sub

'=============================================================================
' LAN CONFIG MENU

' C: int lanConfig_cursor_table [] = {72, 92, 124}
Private Function LanCursorY(ByVal i As Long) As Long
    Select Case i
    Case 0: LanCursorY = 72
    Case 1: LanCursorY = 92
    Case 2: LanCursorY = 124
    End Select
End Function

Private Sub M_Menu_LanConfig_f()
    key_dest = key_menu
    m_state = m_lanconfig
    m_entersound = True
    If lanConfig_cursor = -1 Then
        If JoiningGame() And TCPIPConfig() Then
            lanConfig_cursor = 2
        Else
            lanConfig_cursor = 1
        End If
    End If
    If StartingGame() And lanConfig_cursor = 2 Then lanConfig_cursor = 1
    lanConfig_port = DEFAULTnet_hostport
    lanConfig_portname = CStr(lanConfig_port)

    m_return_onerror = False
    m_return_reason = ""
End Sub

Private Sub M_LanConfig_Draw()
    Dim p As Long, basex As Long, startJoin As String, protocol As String

    M_DrawTransPic 16, 4, Draw_CachePic("gfx/qplaque.lmp")
    p = Draw_CachePic("gfx/p_multi.lmp")
    basex = (320 - pics(p).width) \ 2
    M_DrawPic basex, 4, p

    If StartingGame() Then
        startJoin = "New Game"
    Else
        startJoin = "Join Game"
    End If
    If IPXConfig() Then
        protocol = "IPX"
    Else
        protocol = "TCP/IP"
    End If
    M_Print basex, 32, va("%s - %s", startJoin, protocol)
    basex = basex + 8

    M_Print basex, 52, "Address:"
    If IPXConfig() Then
        M_Print basex + 9 * 8, 52, my_ipx_address
    Else
        M_Print basex + 9 * 8, 52, my_tcpip_address
    End If

    M_Print basex, LanCursorY(0), "Port"
    M_DrawTextBox basex + 8 * 8, LanCursorY(0) - 8, 6, 1
    M_Print basex + 9 * 8, LanCursorY(0), lanConfig_portname

    If JoiningGame() Then
        M_Print basex, LanCursorY(1), "Search for local games..."
        M_Print basex, 108, "Join game at:"
        M_DrawTextBox basex + 8, LanCursorY(2) - 8, 22, 1
        M_Print basex + 16, LanCursorY(2), lanConfig_joinname
    Else
        M_DrawTextBox basex, LanCursorY(1) - 8, 2, 1
        M_Print basex + 8, LanCursorY(1), "OK"
    End If

    M_DrawCharacter basex - 8, LanCursorY(lanConfig_cursor), BlinkChar(12)

    If lanConfig_cursor = 0 Then M_DrawCharacter basex + 9 * 8 + 8 * Len(lanConfig_portname), LanCursorY(0), BlinkChar(10)

    If lanConfig_cursor = 2 Then M_DrawCharacter basex + 16 + 8 * Len(lanConfig_joinname), LanCursorY(2), BlinkChar(10)

    If Len(m_return_reason) Then M_PrintWhite basex, 148, m_return_reason
End Sub

Private Sub M_LanConfig_Key(ByVal key As Long)
    Dim l As Long

    Select Case key
    Case K_ESCAPE
        M_Menu_Net_f

    Case K_UPARROW
        S_LocalSound "misc/menu1.wav"
        lanConfig_cursor = lanConfig_cursor - 1
        If lanConfig_cursor < 0 Then lanConfig_cursor = NUM_LANCONFIG_CMDS - 1

    Case K_DOWNARROW
        S_LocalSound "misc/menu1.wav"
        lanConfig_cursor = lanConfig_cursor + 1
        If lanConfig_cursor >= NUM_LANCONFIG_CMDS Then lanConfig_cursor = 0

    Case K_ENTER
        If lanConfig_cursor = 0 Then GoTo done

        m_entersound = True

        M_ConfigureNetSubsystem

        If lanConfig_cursor = 1 Then
            If StartingGame() Then
                M_Menu_GameOptions_f
                GoTo done
            End If
            M_Menu_Search_f
            GoTo done
        End If

        If lanConfig_cursor = 2 Then
            m_return_state = m_state
            m_return_onerror = True
            key_dest = key_game
            m_state = m_none
            Cbuf_AddText va("connect ""%s""\n", lanConfig_joinname)
        End If

    Case K_BACKSPACE
        If lanConfig_cursor = 0 Then
            If Len(lanConfig_portname) Then lanConfig_portname = Left$(lanConfig_portname, Len(lanConfig_portname) - 1)
        End If

        If lanConfig_cursor = 2 Then
            If Len(lanConfig_joinname) Then lanConfig_joinname = Left$(lanConfig_joinname, Len(lanConfig_joinname) - 1)
        End If

    Case Else
        If key < 32 Or key > 127 Then GoTo done

        If lanConfig_cursor = 2 Then
            If Len(lanConfig_joinname) < 21 Then lanConfig_joinname = lanConfig_joinname & ChrW$(key)
        End If

        If key < 48 Or key > 57 Then GoTo done   ' '0' .. '9'
        If lanConfig_cursor = 0 Then
            If Len(lanConfig_portname) < 5 Then lanConfig_portname = lanConfig_portname & ChrW$(key)
        End If
    End Select
done:

    If StartingGame() And lanConfig_cursor = 2 Then
        If key = K_UPARROW Then
            lanConfig_cursor = 1
        Else
            lanConfig_cursor = 0
        End If
    End If

    l = Q_atoi(lanConfig_portname)
    If l > 65535 Then
        l = lanConfig_port
    Else
        lanConfig_port = l
    End If
    lanConfig_portname = CStr(lanConfig_port)
End Sub

'=============================================================================
' GAME OPTIONS MENU

Private Sub SetLevels(lv() As level_t, ByVal a As Variant)
    Dim i As Long

    ReDim lv(0 To (UBound(a) + 1) \ 2 - 1)
    For i = 0 To UBound(lv)
        lv(i).name = a(i * 2)
        lv(i).description = a(i * 2 + 1)
    Next
End Sub

Private Sub SetEpisodes(ep() As episode_t, ByVal a As Variant)
    Dim i As Long

    ReDim ep(0 To (UBound(a) + 1) \ 3 - 1)
    For i = 0 To UBound(ep)
        ep(i).description = a(i * 3)
        ep(i).firstLevel = a(i * 3 + 1)
        ep(i).levels = a(i * 3 + 2)
    Next
End Sub

' (QuakeVB) C's static level and episode tables
Private Sub InitLevels()
    SetLevels levels, Array( _
        "start", "Entrance", _
        "e1m1", "Slipgate Complex", "e1m2", "Castle of the Damned", "e1m3", "The Necropolis", _
        "e1m4", "The Grisly Grotto", "e1m5", "Gloom Keep", "e1m6", "The Door To Chthon", _
        "e1m7", "The House of Chthon", "e1m8", "Ziggurat Vertigo", _
        "e2m1", "The Installation", "e2m2", "Ogre Citadel", "e2m3", "Crypt of Decay", _
        "e2m4", "The Ebon Fortress", "e2m5", "The Wizard's Manse", "e2m6", "The Dismal Oubliette", _
        "e2m7", "Underearth", _
        "e3m1", "Termination Central", "e3m2", "The Vaults of Zin", "e3m3", "The Tomb of Terror", _
        "e3m4", "Satan's Dark Delight", "e3m5", "Wind Tunnels", "e3m6", "Chambers of Torment", _
        "e3m7", "The Haunted Halls", _
        "e4m1", "The Sewage System", "e4m2", "The Tower of Despair", "e4m3", "The Elder God Shrine", _
        "e4m4", "The Palace of Hate", "e4m5", "Hell's Atrium", "e4m6", "The Pain Maze", _
        "e4m7", "Azure Agony", "e4m8", "The Nameless City", _
        "end", "Shub-Niggurath's Pit", _
        "dm1", "Place of Two Deaths", "dm2", "Claustrophobopolis", "dm3", "The Abandoned Base", _
        "dm4", "The Bad Place", "dm5", "The Cistern", "dm6", "The Dark Zone")

    'MED 01/06/97 added hipnotic levels
    SetLevels hipnoticlevels, Array( _
        "start", "Command HQ", _
        "hip1m1", "The Pumping Station", "hip1m2", "Storage Facility", "hip1m3", "The Lost Mine", _
        "hip1m4", "Research Facility", "hip1m5", "Military Complex", _
        "hip2m1", "Ancient Realms", "hip2m2", "The Black Cathedral", "hip2m3", "The Catacombs", _
        "hip2m4", "The Crypt", "hip2m5", "Mortum's Keep", "hip2m6", "The Gremlin's Domain", _
        "hip3m1", "Tur Torment", "hip3m2", "Pandemonium", "hip3m3", "Limbo", "hip3m4", "The Gauntlet", _
        "hipend", "Armagon's Lair", _
        "hipdm1", "The Edge of Oblivion")

    'PGM 01/07/97 added rogue levels
    'PGM 03/02/97 added dmatch level
    SetLevels roguelevels, Array( _
        "start", "Split Decision", _
        "r1m1", "Deviant's Domain", "r1m2", "Dread Portal", "r1m3", "Judgement Call", _
        "r1m4", "Cave of Death", "r1m5", "Towers of Wrath", "r1m6", "Temple of Pain", _
        "r1m7", "Tomb of the Overlord", _
        "r2m1", "Tempus Fugit", "r2m2", "Elemental Fury I", "r2m3", "Elemental Fury II", _
        "r2m4", "Curse of Osiris", "r2m5", "Wizard's Keep", "r2m6", "Blood Sacrifice", _
        "r2m7", "Last Bastion", "r2m8", "Source of Evil", _
        "ctf1", "Division of Change")

    SetEpisodes episodes, Array( _
        "Welcome to Quake", 0, 1, _
        "Doomed Dimension", 1, 8, _
        "Realm of Black Magic", 9, 7, _
        "Netherworld", 16, 7, _
        "The Elder World", 23, 8, _
        "Final Level", 31, 1, _
        "Deathmatch Arena", 32, 6)

    'MED 01/06/97  added hipnotic episodes
    SetEpisodes hipnoticepisodes, Array( _
        "Scourge of Armagon", 0, 1, _
        "Fortress of the Dead", 1, 5, _
        "Dominion of Darkness", 6, 6, _
        "The Rift", 12, 4, _
        "Final Level", 16, 1, _
        "Deathmatch Arena", 17, 1)

    'PGM 01/07/97 added rogue episodes
    'PGM 03/02/97 added dmatch episode
    SetEpisodes rogueepisodes, Array( _
        "Introduction", 0, 1, _
        "Hell's Fortress", 1, 7, _
        "Corridors of Time", 8, 8, _
        "Deathmatch Arena", 16, 1)
End Sub

Private Sub M_Menu_GameOptions_f()
    key_dest = key_menu
    m_state = m_gameoptions
    m_entersound = True
    If maxplayers = 0 Then maxplayers = svs.maxclients
    If maxplayers < 2 Then maxplayers = svs.maxclientslimit
End Sub

' C: int gameoptions_cursor_table[] = {40, 56, 64, 72, 80, 88, 96, 112, 120}
Private Function GameOptionsCursorY(ByVal i As Long) As Long
    Select Case i
    Case 0: GameOptionsCursorY = 40
    Case 1: GameOptionsCursorY = 56
    Case 2: GameOptionsCursorY = 64
    Case 3: GameOptionsCursorY = 72
    Case 4: GameOptionsCursorY = 80
    Case 5: GameOptionsCursorY = 88
    Case 6: GameOptionsCursorY = 96
    Case 7: GameOptionsCursorY = 112
    Case 8: GameOptionsCursorY = 120
    End Select
End Function

' (QuakeVB) the level the episode and level cursors pick, in the table of the game
Private Sub GetStartLevel(name As String, description As String)
    Dim l As Long

    If hipnotic Then
        l = hipnoticepisodes(startepisode).firstLevel + startlevel
        name = hipnoticlevels(l).name
        description = hipnoticlevels(l).description
    ElseIf rogue Then
        l = rogueepisodes(startepisode).firstLevel + startlevel
        name = roguelevels(l).name
        description = roguelevels(l).description
    Else
        l = episodes(startepisode).firstLevel + startlevel
        name = levels(l).name
        description = levels(l).description
    End If
End Sub

Private Sub M_GameOptions_Draw()
    Dim p As Long, x As Long, msg As String, name As String, description As String

    M_DrawTransPic 16, 4, Draw_CachePic("gfx/qplaque.lmp")
    p = Draw_CachePic("gfx/p_multi.lmp")
    M_DrawPic (320 - pics(p).width) \ 2, 4, p

    M_DrawTextBox 152, 32, 10, 1
    M_Print 160, 40, "begin game"

    M_Print 0, 56, "      Max players"
    M_Print 160, 56, va("%i", maxplayers)

    M_Print 0, 64, "        Game Type"
    If cvars(cv_coop).value Then
        M_Print 160, 64, "Cooperative"
    Else
        M_Print 160, 64, "Deathmatch"
    End If

    M_Print 0, 72, "        Teamplay"
    If rogue Then
        Select Case ToInt(cvars(cv_teamplay).value)
        Case 1: msg = "No Friendly Fire"
        Case 2: msg = "Friendly Fire"
        Case 3: msg = "Tag"
        Case 4: msg = "Capture the Flag"
        Case 5: msg = "One Flag CTF"
        Case 6: msg = "Three Team CTF"
        Case Else: msg = "Off"
        End Select
        M_Print 160, 72, msg
    Else
        Select Case ToInt(cvars(cv_teamplay).value)
        Case 1: msg = "No Friendly Fire"
        Case 2: msg = "Friendly Fire"
        Case Else: msg = "Off"
        End Select
        M_Print 160, 72, msg
    End If

    M_Print 0, 80, "            Skill"
    If cvars(cv_skill).value = 0 Then
        M_Print 160, 80, "Easy difficulty"
    ElseIf cvars(cv_skill).value = 1 Then
        M_Print 160, 80, "Normal difficulty"
    ElseIf cvars(cv_skill).value = 2 Then
        M_Print 160, 80, "Hard difficulty"
    Else
        M_Print 160, 80, "Nightmare difficulty"
    End If

    M_Print 0, 88, "       Frag Limit"
    If cvars(cv_fraglimit).value = 0 Then
        M_Print 160, 88, "none"
    Else
        M_Print 160, 88, va("%i frags", ToInt(cvars(cv_fraglimit).value))
    End If

    M_Print 0, 96, "       Time Limit"
    If cvars(cv_timelimit).value = 0 Then
        M_Print 160, 96, "none"
    Else
        M_Print 160, 96, va("%i minutes", ToInt(cvars(cv_timelimit).value))
    End If

    M_Print 0, 112, "         Episode"
    'MED 01/06/97 added hipnotic episodes
    If hipnotic Then
        M_Print 160, 112, hipnoticepisodes(startepisode).description
    'PGM 01/07/97 added rogue episodes
    ElseIf rogue Then
        M_Print 160, 112, rogueepisodes(startepisode).description
    Else
        M_Print 160, 112, episodes(startepisode).description
    End If

    M_Print 0, 120, "           Level"
    GetStartLevel name, description
    M_Print 160, 120, description
    M_Print 160, 128, name

    ' line cursor
    M_DrawCharacter 144, GameOptionsCursorY(gameoptions_cursor), BlinkChar(12)

    If m_serverInfoMessage Then
        If (realtime - m_serverInfoMessageTime) < 5# Then
            x = (320 - 26 * 8) \ 2
            M_DrawTextBox x, 138, 24, 4
            x = x + 8
            M_Print x, 146, "  More than 4 players   "
            M_Print x, 154, " requires using command "
            M_Print x, 162, "line parameters; please "
            M_Print x, 170, "   see techinfo.txt.    "
        Else
            m_serverInfoMessage = False
        End If
    End If
End Sub

Private Sub M_NetStart_Change(ByVal dir_ As Long)
    Dim count As Long

    Select Case gameoptions_cursor
    Case 1
        maxplayers = maxplayers + dir_
        If maxplayers > svs.maxclientslimit Then
            maxplayers = svs.maxclientslimit
            m_serverInfoMessage = True
            m_serverInfoMessageTime = realtime
        End If
        If maxplayers < 2 Then maxplayers = 2

    Case 2
        If cvars(cv_coop).value Then
            Cvar_SetValue "coop", 0
        Else
            Cvar_SetValue "coop", 1
        End If

    Case 3
        If rogue Then
            count = 6
        Else
            count = 2
        End If

        Cvar_SetValue "teamplay", cvars(cv_teamplay).value + dir_
        If cvars(cv_teamplay).value > count Then
            Cvar_SetValue "teamplay", 0
        ElseIf cvars(cv_teamplay).value < 0 Then
            Cvar_SetValue "teamplay", count
        End If

    Case 4
        Cvar_SetValue "skill", cvars(cv_skill).value + dir_
        If cvars(cv_skill).value > 3 Then Cvar_SetValue "skill", 0
        If cvars(cv_skill).value < 0 Then Cvar_SetValue "skill", 3

    Case 5
        Cvar_SetValue "fraglimit", cvars(cv_fraglimit).value + dir_ * 10
        If cvars(cv_fraglimit).value > 100 Then Cvar_SetValue "fraglimit", 0
        If cvars(cv_fraglimit).value < 0 Then Cvar_SetValue "fraglimit", 100

    Case 6
        Cvar_SetValue "timelimit", cvars(cv_timelimit).value + dir_ * 5
        If cvars(cv_timelimit).value > 60 Then Cvar_SetValue "timelimit", 0
        If cvars(cv_timelimit).value < 0 Then Cvar_SetValue "timelimit", 60

    Case 7
        startepisode = startepisode + dir_
        'MED 01/06/97 added hipnotic count
        If hipnotic Then
            count = 6
        'PGM 01/07/97 added rogue count
        'PGM 03/02/97 added 1 for dmatch episode
        ElseIf rogue Then
            count = 4
        ElseIf cvars(cv_registered).value Then
            count = 7
        Else
            count = 2
        End If

        If startepisode < 0 Then startepisode = count - 1

        If startepisode >= count Then startepisode = 0

        startlevel = 0

    Case 8
        startlevel = startlevel + dir_
        'MED 01/06/97 added hipnotic episodes
        If hipnotic Then
            count = hipnoticepisodes(startepisode).levels
        'PGM 01/06/97 added hipnotic episodes
        ElseIf rogue Then
            count = rogueepisodes(startepisode).levels
        Else
            count = episodes(startepisode).levels
        End If

        If startlevel < 0 Then startlevel = count - 1

        If startlevel >= count Then startlevel = 0
    End Select
End Sub

Private Sub M_GameOptions_Key(ByVal key As Long)
    Dim name As String, description As String

    Select Case key
    Case K_ESCAPE
        M_Menu_Net_f

    Case K_UPARROW
        S_LocalSound "misc/menu1.wav"
        gameoptions_cursor = gameoptions_cursor - 1
        If gameoptions_cursor < 0 Then gameoptions_cursor = NUM_GAMEOPTIONS - 1

    Case K_DOWNARROW
        S_LocalSound "misc/menu1.wav"
        gameoptions_cursor = gameoptions_cursor + 1
        If gameoptions_cursor >= NUM_GAMEOPTIONS Then gameoptions_cursor = 0

    Case K_LEFTARROW
        If gameoptions_cursor = 0 Then Exit Sub
        S_LocalSound "misc/menu3.wav"
        M_NetStart_Change -1

    Case K_RIGHTARROW
        If gameoptions_cursor = 0 Then Exit Sub
        S_LocalSound "misc/menu3.wav"
        M_NetStart_Change 1

    Case K_ENTER
        S_LocalSound "misc/menu2.wav"
        If gameoptions_cursor = 0 Then
            If sv.active Then Cbuf_AddText "disconnect" & vbLf
            Cbuf_AddText "listen 0" & vbLf   ' so host_netport will be re-examined
            Cbuf_AddText va("maxplayers %u\n", maxplayers)
            SCR_BeginLoadingPlaque

            GetStartLevel name, description
            Cbuf_AddText va("map %s\n", name)

            Exit Sub
        End If

        M_NetStart_Change 1
    End Select
End Sub

'=============================================================================
' SEARCH MENU

Private Sub M_Menu_Search_f()
    key_dest = key_menu
    m_state = m_search
    m_entersound = False
    slistSilent = True
    slistLocal = False
    searchComplete = False
    NET_Slist_f
End Sub

Private Sub M_Search_Draw()
    Dim p As Long, x As Long

    p = Draw_CachePic("gfx/p_multi.lmp")
    M_DrawPic (320 - pics(p).width) \ 2, 4, p
    x = (320 \ 2) - ((12 * 8) \ 2) + 4
    M_DrawTextBox x - 8, 32, 12, 1
    M_Print x, 40, "Searching..."

    If slistInProgress Then
        NET_Poll
        Exit Sub
    End If

    If Not searchComplete Then
        searchComplete = True
        searchCompleteTime = realtime
    End If

    If hostCacheCount Then
        M_Menu_ServerList_f
        Exit Sub
    End If

    M_PrintWhite (320 \ 2) - ((22 * 8) \ 2), 64, "No Quake servers found"
    If (realtime - searchCompleteTime) < 3# Then Exit Sub

    M_Menu_LanConfig_f
End Sub

Private Sub M_Search_Key(ByVal key As Long)
End Sub

'=============================================================================
' SLIST MENU

Private Sub M_Menu_ServerList_f()
    key_dest = key_menu
    m_state = m_slist
    m_entersound = True
    slist_cursor = 0
    m_return_onerror = False
    m_return_reason = ""
    slist_sorted = False
End Sub

Private Sub M_ServerList_Draw()
    Dim n As Long, string_ As String, p As Long, i As Long, j As Long, temp As hostcache_t

    If Not slist_sorted Then
        If hostCacheCount > 1 Then
            For i = 0 To hostCacheCount - 1
                For j = i + 1 To hostCacheCount - 1
                    If StrComp(hostcache(j).name, hostcache(i).name, vbBinaryCompare) < 0 Then
                        temp = hostcache(j)
                        hostcache(j) = hostcache(i)
                        hostcache(i) = temp
                    End If
                Next
            Next
        End If
        slist_sorted = True
    End If

    p = Draw_CachePic("gfx/p_multi.lmp")
    M_DrawPic (320 - pics(p).width) \ 2, 4, p
    For n = 0 To hostCacheCount - 1
        If hostcache(n).maxusers Then
            string_ = va("%-15.15s %-15.15s %2u/%2u\n", hostcache(n).name, hostcache(n).map, hostcache(n).users, hostcache(n).maxusers)
        Else
            string_ = va("%-15.15s %-15.15s\n", hostcache(n).name, hostcache(n).map)
        End If
        M_Print 16, 32 + 8 * n, string_
    Next
    M_DrawCharacter 0, 32 + slist_cursor * 8, BlinkChar(12)

    If Len(m_return_reason) Then M_PrintWhite 16, 148, m_return_reason
End Sub

Private Sub M_ServerList_Key(ByVal k As Long)
    Select Case k
    Case K_ESCAPE
        M_Menu_LanConfig_f

    Case K_SPACE
        M_Menu_Search_f

    Case K_UPARROW, K_LEFTARROW
        S_LocalSound "misc/menu1.wav"
        slist_cursor = slist_cursor - 1
        If slist_cursor < 0 Then slist_cursor = hostCacheCount - 1

    Case K_DOWNARROW, K_RIGHTARROW
        S_LocalSound "misc/menu1.wav"
        slist_cursor = slist_cursor + 1
        If slist_cursor >= hostCacheCount Then slist_cursor = 0

    Case K_ENTER
        S_LocalSound "misc/menu2.wav"
        m_return_state = m_state
        m_return_onerror = True
        slist_sorted = False
        key_dest = key_game
        m_state = m_none
        Cbuf_AddText va("connect ""%s""\n", hostcache(slist_cursor).cname)
    End Select
End Sub

'=============================================================================
' Menu Subsystem

Public Sub M_Init()
    ' (C: static initializers)
    setup_cursor = 4
    lanConfig_cursor = -1
    InitBindnames
    InitLevels

    Cmd_AddCommand "togglemenu", XC_M_ToggleMenu_f

    Cmd_AddCommand "menu_main", XC_M_Menu_Main_f
    Cmd_AddCommand "menu_singleplayer", XC_M_Menu_SinglePlayer_f
    Cmd_AddCommand "menu_load", XC_M_Menu_Load_f
    Cmd_AddCommand "menu_save", XC_M_Menu_Save_f
    Cmd_AddCommand "menu_multiplayer", XC_M_Menu_MultiPlayer_f
    Cmd_AddCommand "menu_setup", XC_M_Menu_Setup_f
    Cmd_AddCommand "menu_options", XC_M_Menu_Options_f
    Cmd_AddCommand "menu_keys", XC_M_Menu_Keys_f
    Cmd_AddCommand "menu_video", XC_M_Menu_Video_f
    Cmd_AddCommand "help", XC_M_Menu_Help_f
    Cmd_AddCommand "menu_quit", XC_M_Menu_Quit_f
End Sub

' (QuakeVB) the console commands of this module, for Cmd_Call
Public Sub M_CallCommand(ByVal fn As Long)
    Select Case fn
    Case XC_M_ToggleMenu_f: M_ToggleMenu_f
    Case XC_M_Menu_Main_f: M_Menu_Main_f
    Case XC_M_Menu_SinglePlayer_f: M_Menu_SinglePlayer_f
    Case XC_M_Menu_Load_f: M_Menu_Load_f
    Case XC_M_Menu_Save_f: M_Menu_Save_f
    Case XC_M_Menu_MultiPlayer_f: M_Menu_MultiPlayer_f
    Case XC_M_Menu_Setup_f: M_Menu_Setup_f
    Case XC_M_Menu_Options_f: M_Menu_Options_f
    Case XC_M_Menu_Keys_f: M_Menu_Keys_f
    Case XC_M_Menu_Video_f: M_Menu_Video_f
    Case XC_M_Menu_Help_f: M_Menu_Help_f
    Case XC_M_Menu_Quit_f: M_Menu_Quit_f
    End Select
End Sub

Public Sub M_Draw()
    If m_state = m_none Or key_dest <> key_menu Then Exit Sub

    If Not m_recursiveDraw Then
        scr_copyeverything = 1

        If scr_con_current Then
            Draw_ConsoleBackground vid.height
            VID_UnlockBuffer
            S_ExtraUpdate
            VID_LockBuffer
        Else
            Draw_FadeScreen
        End If

        scr_fullupdate = 0
    Else
        m_recursiveDraw = False
    End If

    Select Case m_state
    Case m_main: M_Main_Draw
    Case m_singleplayer: M_SinglePlayer_Draw
    Case m_load: M_Load_Draw
    Case m_save: M_Save_Draw
    Case m_multiplayer: M_MultiPlayer_Draw
    Case m_setup: M_Setup_Draw
    Case m_net: M_Net_Draw
    Case m_options: M_Options_Draw
    Case m_keys: M_Keys_Draw
    Case m_video: M_Video_Draw
    Case m_help: M_Help_Draw
    Case m_quit: M_Quit_Draw
    Case m_lanconfig: M_LanConfig_Draw
    Case m_gameoptions: M_GameOptions_Draw
    Case m_search: M_Search_Draw
    Case m_slist: M_ServerList_Draw
    End Select

    If m_entersound Then
        S_LocalSound "misc/menu2.wav"
        m_entersound = False
    End If

    VID_UnlockBuffer
    S_ExtraUpdate
    VID_LockBuffer
End Sub

Public Sub M_Keydown(ByVal key As Long)
    Select Case m_state
    Case m_none: Exit Sub
    Case m_main: M_Main_Key key
    Case m_singleplayer: M_SinglePlayer_Key key
    Case m_load: M_Load_Key key
    Case m_save: M_Save_Key key
    Case m_multiplayer: M_MultiPlayer_Key key
    Case m_setup: M_Setup_Key key
    Case m_net: M_Net_Key key
    Case m_options: M_Options_Key key
    Case m_keys: M_Keys_Key key
    Case m_video: M_Video_Key key
    Case m_help: M_Help_Key key
    Case m_quit: M_Quit_Key key
    Case m_lanconfig: M_LanConfig_Key key
    Case m_gameoptions: M_GameOptions_Key key
    Case m_search: M_Search_Key key
    Case m_slist: M_ServerList_Key key
    End Select
End Sub

Private Sub M_ConfigureNetSubsystem()
    ' enable/disable net systems to match desired config

    Cbuf_AddText "stopdemo" & vbLf
    If m_net_cursor = 0 Or m_net_cursor = 1 Then   ' SerialConfig || DirectConfig
        Cbuf_AddText "com1 enable" & vbLf
    End If

    If IPXConfig() Or TCPIPConfig() Then net_hostport = lanConfig_port
End Sub
