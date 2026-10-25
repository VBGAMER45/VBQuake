Attribute VB_Name = "quakedef"
' quakedef.bas - types, constants and globals shared by every module: quakedef.h,
' common.h, cvar.h, cmd.h, mathlib.h, vid.h, wad.h, draw.h, keys.h, console.h,
' screen.h, render.h, sys.h, and the parts of client.h, server.h and net.h that
' the ported clusters use so far (those types grow as their clusters arrive).
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

'=============================================================================
' Win32 helpers used everywhere

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal n As Long)
Public Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (dest As Any, ByVal n As Long, ByVal fill As Byte)
Public Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (arr() As Any) As Long

'=============================================================================
' quakedef.h

Public Const QUAKE_VERSION As Single = 1.09      ' C: VERSION (clashes with "version")
Public Const GAMENAME As String = "id1"

Public Const MINIMUM_MEMORY_ As Long = &H550000
Public Const MINIMUM_MEMORY_LEVELPAK As Long = &H650000
Public Const MAX_NUM_ARGVS As Long = 50

' up / down, left / right, fall over
Public Const PITCH As Long = 0
Public Const YAW As Long = 1
Public Const ROLL As Long = 2

Public Const MAX_QPATH As Long = 64              ' max length of a quake game pathname
Public Const MAX_OSPATH As Long = 128            ' max length of a filesystem pathname
Public Const ON_EPSILON As Single = 0.1          ' point on plane side epsilon
Public Const MAX_MSGLEN As Long = 8000           ' max length of a reliable message
Public Const MAX_DATAGRAM As Long = 1024         ' max length of unreliable message

' per-level limits
Public Const MAX_EDICTS As Long = 600
Public Const MAX_LIGHTSTYLES As Long = 64
Public Const MAX_MODELS As Long = 256            ' these are sent over the net as bytes
Public Const MAX_SOUNDS As Long = 256            ' so they cannot be blindly increased
Public Const SAVEGAME_COMMENT_LENGTH As Long = 39
Public Const MAX_STYLESTRING As Long = 64

' stats are integers communicated to the client by the server
Public Const MAX_CL_STATS As Long = 32
Public Const STAT_HEALTH As Long = 0
Public Const STAT_FRAGS As Long = 1
Public Const STAT_WEAPON As Long = 2
Public Const STAT_AMMO As Long = 3
Public Const STAT_ARMOR As Long = 4
Public Const STAT_WEAPONFRAME As Long = 5
Public Const STAT_SHELLS As Long = 6
Public Const STAT_NAILS As Long = 7
Public Const STAT_ROCKETS As Long = 8
Public Const STAT_CELLS As Long = 9
Public Const STAT_ACTIVEWEAPON As Long = 10
Public Const STAT_TOTALSECRETS As Long = 11
Public Const STAT_TOTALMONSTERS As Long = 12
Public Const STAT_SECRETS As Long = 13           ' bumped on client side by svc_foundsecret
Public Const STAT_MONSTERS As Long = 14          ' bumped by svc_killedmonster

' stock defines
Public Const IT_SHOTGUN As Long = 1
Public Const IT_SUPER_SHOTGUN As Long = 2
Public Const IT_NAILGUN As Long = 4
Public Const IT_SUPER_NAILGUN As Long = 8
Public Const IT_GRENADE_LAUNCHER As Long = 16
Public Const IT_ROCKET_LAUNCHER As Long = 32
Public Const IT_LIGHTNING As Long = 64
Public Const IT_SUPER_LIGHTNING As Long = 128
Public Const IT_SHELLS As Long = 256
Public Const IT_NAILS As Long = 512
Public Const IT_ROCKETS As Long = 1024
Public Const IT_CELLS As Long = 2048
Public Const IT_AXE As Long = 4096
Public Const IT_ARMOR1 As Long = 8192
Public Const IT_ARMOR2 As Long = 16384
Public Const IT_ARMOR3 As Long = 32768
Public Const IT_SUPERHEALTH As Long = 65536
Public Const IT_KEY1 As Long = 131072
Public Const IT_KEY2 As Long = 262144
Public Const IT_INVISIBILITY As Long = 524288
Public Const IT_INVULNERABILITY As Long = 1048576
Public Const IT_SUIT As Long = 2097152
Public Const IT_QUAD As Long = 4194304
Public Const IT_SIGIL1 As Long = &H10000000
Public Const IT_SIGIL2 As Long = &H20000000
Public Const IT_SIGIL3 As Long = &H40000000
Public Const IT_SIGIL4 As Long = &H80000000

' rogue changed and added defines
Public Const RIT_SHELLS As Long = 128
Public Const RIT_NAILS As Long = 256
Public Const RIT_ROCKETS As Long = 512
Public Const RIT_CELLS As Long = 1024
Public Const RIT_AXE As Long = 2048
Public Const RIT_LAVA_NAILGUN As Long = 4096
Public Const RIT_LAVA_SUPER_NAILGUN As Long = 8192
Public Const RIT_MULTI_GRENADE As Long = 16384
Public Const RIT_MULTI_ROCKET As Long = 32768
Public Const RIT_PLASMA_GUN As Long = 65536
Public Const RIT_ARMOR1 As Long = 8388608
Public Const RIT_ARMOR2 As Long = 16777216
Public Const RIT_ARMOR3 As Long = 33554432
Public Const RIT_LAVA_NAILS As Long = 67108864
Public Const RIT_PLASMA_AMMO As Long = 134217728
Public Const RIT_MULTI_ROCKETS As Long = 268435456
Public Const RIT_SHIELD As Long = 536870912
Public Const RIT_ANTIGRAV As Long = 1073741824
Public Const RIT_SUPERHEALTH As Long = &H80000000

' hipnotic added defines
Public Const HIT_PROXIMITY_GUN_BIT As Long = 16
Public Const HIT_MJOLNIR_BIT As Long = 7
Public Const HIT_LASER_CANNON_BIT As Long = 23
Public Const HIT_PROXIMITY_GUN As Long = &H10000     ' 1 << 16
Public Const HIT_MJOLNIR As Long = &H80             ' 1 << 7
Public Const HIT_LASER_CANNON As Long = &H800000    ' 1 << 23
Public Const HIT_WETSUIT As Long = &H2000000        ' 1 << (23+2)
Public Const HIT_EMPATHY_SHIELDS As Long = &H4000000   ' 1 << (23+3)

Public Const MAX_SCOREBOARD As Long = 16
Public Const MAX_SCOREBOARDNAME As Long = 32
Public Const SOUND_CHANNELS As Long = 8

' the host system specifies the base of the directory tree and the command
' line parms passed to the program (the memory fields are not used: VB
' allocates as it goes)
Public Type quakeparms_t
    basedir As String
    cachedir As String
    argc As Long
End Type

' Host_Error and Host_EndGame unwind to _Host_Frame with this VB error
' (C: longjmp (host_abortserver, 1))
Public Const HOST_ABORTSERVER As Long = vbObjectError + &H1066

'=============================================================================
' common.h

Public Type sizebuf_t
    allowoverflow As Boolean     ' if false, do a Sys_Error
    overflowed As Boolean        ' set to true if the buffer size failed
    data() As Byte
    maxsize As Long
    cursize As Long
End Type

Public Const Q_MAXINT As Long = &H7FFFFFFF
Public Const Q_MININT As Long = &H80000000

'=============================================================================
' mathlib.h

Public Type vec3_t
    v(0 To 2) As Single
End Type

Public Const M_PI As Double = 3.14159265358979

' mplane_t (model.h), needed by BoxOnPlaneSide
Public Type mplane_t
    normal As vec3_t
    dist As Single
    typ As Byte                  ' for texture axis selection and fast side tests
    signbits As Byte             ' signx + signy<<1 + signz<<1
End Type

'=============================================================================
' cvar.h: cvars live in cvars(1..num_cvars); a C cvar global is the index of
' its entry (cv_<name>); 0 is an unregistered dummy whose value is 0

Public Type cvar_t
    name As String
    strval As String             ' C: string
    archive As Boolean           ' set to true to cause it to be saved to vars.rc
    server As Boolean            ' notifies players when changed
    value As Single
End Type

Public cvars() As cvar_t
Public num_cvars As Long

'=============================================================================
' cmd.h

Public Const src_client As Long = 0      ' came in over a net connection as a clc_stringcmd
Public Const src_command As Long = 1     ' from the command buffer

' xcommand_t: the functions behind console commands (dispatched by Cmd_Call in cmd.bas)
Public Const XC_Cmd_StuffCmds_f As Long = 1
Public Const XC_Cmd_Exec_f As Long = 2
Public Const XC_Cmd_Echo_f As Long = 3
Public Const XC_Cmd_Alias_f As Long = 4
Public Const XC_Cmd_ForwardToServer As Long = 5
Public Const XC_Cmd_Wait_f As Long = 6
Public Const XC_COM_Path_f As Long = 7
Public Const XC_Con_ToggleConsole_f As Long = 8
Public Const XC_Con_MessageMode_f As Long = 9
Public Const XC_Con_MessageMode2_f As Long = 10
Public Const XC_Con_Clear_f As Long = 11
Public Const XC_Key_Bind_f As Long = 12
Public Const XC_Key_Unbind_f As Long = 13
Public Const XC_Key_Unbindall_f As Long = 14
Public Const XC_SCR_ScreenShot_f As Long = 15
Public Const XC_SCR_SizeUp_f As Long = 16
Public Const XC_SCR_SizeDown_f As Long = 17
Public Const XC_Host_Quit_f As Long = 18
Public Const XC_Host_Version_f As Long = 19
Public Const XC_Host_Status_f As Long = 20
Public Const XC_Host_God_f As Long = 21
Public Const XC_Host_Notarget_f As Long = 22
Public Const XC_Host_Fly_f As Long = 23
Public Const XC_Host_Map_f As Long = 24
Public Const XC_Host_Restart_f As Long = 25
Public Const XC_Host_Changelevel_f As Long = 26
Public Const XC_Host_Connect_f As Long = 27
Public Const XC_Host_Reconnect_f As Long = 28
Public Const XC_Host_Name_f As Long = 29
Public Const XC_Host_Noclip_f As Long = 30
Public Const XC_Host_Say_f As Long = 31
Public Const XC_Host_Say_Team_f As Long = 32
Public Const XC_Host_Tell_f As Long = 33
Public Const XC_Host_Color_f As Long = 34
Public Const XC_Host_Kill_f As Long = 35
Public Const XC_Host_Pause_f As Long = 36
Public Const XC_Host_Spawn_f As Long = 37
Public Const XC_Host_Begin_f As Long = 38
Public Const XC_Host_PreSpawn_f As Long = 39
Public Const XC_Host_Kick_f As Long = 40
Public Const XC_Host_Ping_f As Long = 41
Public Const XC_Host_Loadgame_f As Long = 42
Public Const XC_Host_Savegame_f As Long = 43
Public Const XC_Host_Give_f As Long = 44
Public Const XC_Host_Startdemos_f As Long = 45
Public Const XC_Host_Demos_f As Long = 46
Public Const XC_Host_Stopdemo_f As Long = 47
Public Const XC_Host_Viewmodel_f As Long = 48
Public Const XC_Host_Viewframe_f As Long = 49
Public Const XC_Host_Viewnext_f As Long = 50
Public Const XC_Host_Viewprev_f As Long = 51
' (host_cmd.bas dispatches XC_Host_Quit_f .. XC_Host_Viewprev_f)
Public Const XC_Mod_Print As Long = 52
Public Const XC_ED_PrintEdict_f As Long = 53
Public Const XC_ED_PrintEdicts As Long = 54
Public Const XC_ED_Count As Long = 55
Public Const XC_PR_Profile_f As Long = 56
' (net_main.bas dispatches these)
Public Const XC_NET_Slist_f As Long = 57
Public Const XC_NET_Listen_f As Long = 58
Public Const XC_MaxPlayers_f As Long = 59
Public Const XC_NET_Port_f As Long = 60
' (cl_main.bas (and cl_demo's) dispatches these)
Public Const XC_CL_PrintEntities_f As Long = 61
Public Const XC_CL_Disconnect_f As Long = 62
Public Const XC_CL_Record_f As Long = 63
Public Const XC_CL_Stop_f As Long = 64
Public Const XC_CL_PlayDemo_f As Long = 65
Public Const XC_CL_TimeDemo_f As Long = 66
' (cl_input.bas dispatches these)
Public Const XC_IN_KLookDown As Long = 67
Public Const XC_IN_KLookUp As Long = 68
Public Const XC_IN_MLookDown As Long = 69
Public Const XC_IN_MLookUp As Long = 70
Public Const XC_IN_UpDown As Long = 71
Public Const XC_IN_UpUp As Long = 72
Public Const XC_IN_DownDown As Long = 73
Public Const XC_IN_DownUp As Long = 74
Public Const XC_IN_LeftDown As Long = 75
Public Const XC_IN_LeftUp As Long = 76
Public Const XC_IN_RightDown As Long = 77
Public Const XC_IN_RightUp As Long = 78
Public Const XC_IN_ForwardDown As Long = 79
Public Const XC_IN_ForwardUp As Long = 80
Public Const XC_IN_BackDown As Long = 81
Public Const XC_IN_BackUp As Long = 82
Public Const XC_IN_LookupDown As Long = 83
Public Const XC_IN_LookupUp As Long = 84
Public Const XC_IN_LookdownDown As Long = 85
Public Const XC_IN_LookdownUp As Long = 86
Public Const XC_IN_MoveleftDown As Long = 87
Public Const XC_IN_MoveleftUp As Long = 88
Public Const XC_IN_MoverightDown As Long = 89
Public Const XC_IN_MoverightUp As Long = 90
Public Const XC_IN_SpeedDown As Long = 91
Public Const XC_IN_SpeedUp As Long = 92
Public Const XC_IN_StrafeDown As Long = 93
Public Const XC_IN_StrafeUp As Long = 94
Public Const XC_IN_AttackDown As Long = 95
Public Const XC_IN_AttackUp As Long = 96
Public Const XC_IN_UseDown As Long = 97
Public Const XC_IN_UseUp As Long = 98
Public Const XC_IN_JumpDown As Long = 99
Public Const XC_IN_JumpUp As Long = 100
Public Const XC_IN_Impulse As Long = 101
' (view.bas dispatches these)
Public Const XC_V_cshift_f As Long = 102
Public Const XC_V_BonusFlash_f As Long = 103
Public Const XC_V_StartPitchDrift As Long = 104
' (in_win.bas dispatches these)
Public Const XC_Force_CenterView_f As Long = 105
Public Const XC_Joy_AdvancedUpdate_f As Long = 106
' (r_main.bas dispatches these)
Public Const XC_R_TimeRefresh_f As Long = 107
Public Const XC_R_ReadPointFile_f As Long = 108
' (sbar.bas dispatches XC_Sbar_ShowScores .. XC_Sbar_DontShowScores)
Public Const XC_Sbar_ShowScores As Long = 109
Public Const XC_Sbar_DontShowScores As Long = 110
' (menu.bas dispatches XC_M_ToggleMenu_f .. XC_M_Menu_Quit_f)
Public Const XC_M_ToggleMenu_f As Long = 111
Public Const XC_M_Menu_Main_f As Long = 112
Public Const XC_M_Menu_SinglePlayer_f As Long = 113
Public Const XC_M_Menu_Load_f As Long = 114
Public Const XC_M_Menu_Save_f As Long = 115
Public Const XC_M_Menu_MultiPlayer_f As Long = 116
Public Const XC_M_Menu_Setup_f As Long = 117
Public Const XC_M_Menu_Options_f As Long = 118
Public Const XC_M_Menu_Keys_f As Long = 119
Public Const XC_M_Menu_Video_f As Long = 120
Public Const XC_M_Menu_Help_f As Long = 121
Public Const XC_M_Menu_Quit_f As Long = 122
' (snd_dma.bas dispatches XC_S_Play .. XC_CD_f)
Public Const XC_S_Play As Long = 123
Public Const XC_S_PlayVol As Long = 124
Public Const XC_S_StopAllSoundsC As Long = 125
Public Const XC_S_SoundList As Long = 126
Public Const XC_S_SoundInfo_f As Long = 127
Public Const XC_CD_f As Long = 128
' (net_dgrm.bas dispatches XC_NET_Stats_f .. XC_Test2_f)
Public Const XC_NET_Stats_f As Long = 129
Public Const XC_NET_Ban_f As Long = 130
Public Const XC_Test_f As Long = 131
Public Const XC_Test2_f As Long = 132

'=============================================================================
' vid.h

Public Const VID_CBITS As Long = 6
Public Const VID_GRADES As Long = 64     ' 1 << VID_CBITS

Public Type vrect_t
    x As Long
    y As Long
    width As Long
    height As Long
End Type

' vid.buffer and vid.conbuffer are the same array, vid_buffer() (offset 0)
Public Type viddef_t
    fullbright As Long           ' index of first fullbright color
    rowbytes As Long             ' may be > width if displayed in a window
    width As Long
    height As Long
    aspect As Single             ' width / height -- < 0 is taller than wide
    numpages As Long
    recalc_refdef As Long        ' if true, recalc vid-based stuff
    conrowbytes As Long
    conwidth As Long
    conheight As Long
    maxwarpwidth As Long
    maxwarpheight As Long
End Type

Public vid As viddef_t           ' global video state
Public vid_buffer() As Byte      ' C: vid.buffer, vid.conbuffer (invisible buffer)
Public vid_bufferofs As Long     ' (QuakeVB) where C's vid.buffer starts in vid_buffer() (lcd_x moves it)
Public vid_colormap() As Byte    ' C: vid.colormap (256 * VID_GRADES)

' d_iface.h
Public Const WARP_WIDTH As Long = 320
Public Const WARP_HEIGHT As Long = 200
Public Const TRANSPARENT_COLOR As Long = &HFF

'=============================================================================
' wad.h

Public Const CMP_NONE As Long = 0
Public Const CMP_LZSS As Long = 1

Public Const TYP_NONE As Long = 0
Public Const TYP_LABEL As Long = 1
Public Const TYP_LUMPY As Long = 64      ' 64 + grab command number
Public Const TYP_PALETTE As Long = 64
Public Const TYP_QTEX As Long = 65
Public Const TYP_QPIC As Long = 66
Public Const TYP_SOUND As Long = 67
Public Const TYP_MIPTEX As Long = 68

Public Type lumpinfo_t
    filepos As Long
    disksize As Long
    size As Long                 ' uncompressed
    typ As Byte
    compression As Byte
    name As String               ' cleaned up: lowercase, at most 16 chars
End Type

'=============================================================================
' draw.h: a qpic_t* is an index into pics(); the pixels are in draw_pool()

Public Type qpic_t
    width As Long
    height As Long
    ofs As Long                  ' C: data, as an offset into draw_pool()
End Type

Public pics() As qpic_t
Public draw_pool() As Byte

'=============================================================================
' keys.h

Public Const K_TAB As Long = 9
Public Const K_ENTER As Long = 13
Public Const K_ESCAPE As Long = 27
Public Const K_SPACE As Long = 32

' normal keys should be passed as lowercased ascii
Public Const K_BACKSPACE As Long = 127
Public Const K_UPARROW As Long = 128
Public Const K_DOWNARROW As Long = 129
Public Const K_LEFTARROW As Long = 130
Public Const K_RIGHTARROW As Long = 131
Public Const K_ALT As Long = 132
Public Const K_CTRL As Long = 133
Public Const K_SHIFT As Long = 134
Public Const K_F1 As Long = 135
Public Const K_F2 As Long = 136
Public Const K_F3 As Long = 137
Public Const K_F4 As Long = 138
Public Const K_F5 As Long = 139
Public Const K_F6 As Long = 140
Public Const K_F7 As Long = 141
Public Const K_F8 As Long = 142
Public Const K_F9 As Long = 143
Public Const K_F10 As Long = 144
Public Const K_F11 As Long = 145
Public Const K_F12 As Long = 146
Public Const K_INS As Long = 147
Public Const K_DEL As Long = 148
Public Const K_PGDN As Long = 149
Public Const K_PGUP As Long = 150
Public Const K_HOME As Long = 151
Public Const K_END As Long = 152
Public Const K_PAUSE As Long = 255

' mouse buttons generate virtual keys
Public Const K_MOUSE1 As Long = 200
Public Const K_MOUSE2 As Long = 201
Public Const K_MOUSE3 As Long = 202

' joystick buttons
Public Const K_JOY1 As Long = 203
Public Const K_JOY2 As Long = 204
Public Const K_JOY3 As Long = 205
Public Const K_JOY4 As Long = 206

' aux keys are for multi-buttoned joysticks to generate so they can use
' the normal binding process (K_AUX1 .. K_AUX32 = 207 .. 238)
Public Const K_AUX1 As Long = 207

Public Const K_MWHEELUP As Long = 239
Public Const K_MWHEELDOWN As Long = 240

' keydest_t
Public Const key_game As Long = 0
Public Const key_console As Long = 1
Public Const key_message As Long = 2
Public Const key_menu As Long = 3

Public key_dest As Long
Public keybindings(0 To 255) As String
Public keybinding_set(0 To 255) As Boolean       ' C: keybindings[i] != NULL
Public key_repeats(0 To 255) As Long             ' if > 1, it is autorepeating
Public key_count As Long                         ' incremented every key event
Public key_lastpress As Long

Public Const MAXCMDLINE As Long = 256
Public key_lines(0 To 31, 0 To MAXCMDLINE - 1) As Byte
Public key_linepos As Long
Public edit_line As Long

'=============================================================================
' console.h

Public con_totallines As Long        ' total lines in console scrollback
Public con_backscroll As Long        ' lines up from bottom to display
Public con_forcedup As Boolean       ' because no entities to refresh
Public con_initialized As Boolean
Public con_notifylines As Long       ' scan lines to clear for notify lines

'=============================================================================
' screen.h

Public scr_con_current As Single
Public scr_conlines As Single        ' lines of console to display
Public scr_fullupdate As Long        ' set to 0 to force full redraw
Public sb_lines As Long              ' (sbar.c)
Public clearnotify As Long           ' set to 0 whenever notify text is drawn
Public scr_disabled_for_loading As Boolean
Public scr_skipupdate As Boolean
Public scr_copytop As Long
Public scr_copyeverything As Long
Public block_drawing As Boolean
Public scr_vrect As vrect_t

'=============================================================================
' render.h

Public Type refdef_t
    vrect As vrect_t                     ' subwindow in video for refresh
    aliasvrect As vrect_t                ' scaled Alias version
    vrectright As Long                   ' right & bottom screen coords
    vrectbottom As Long
    aliasvrectright As Long              ' scaled Alias versions
    aliasvrectbottom As Long
    vrectrightedge As Single             ' rightmost right edge we care about, for use in edge list
    fvrectx As Single                    ' for floating-point compares
    fvrecty As Single
    fvrectx_adj As Single                ' left and top edges, for clamping
    fvrecty_adj As Single
    vrect_x_adj_shift20 As Long          ' (vrect.x + 0.5 - epsilon) << 20
    vrectright_adj_shift20 As Long       ' (vrectright + 0.5 - epsilon) << 20
    fvrectright_adj As Single            ' right and bottom edges, for clamping
    fvrectbottom_adj As Single
    fvrectright As Single                ' rightmost edge, for Alias clamping
    fvrectbottom As Single               ' bottommost edge, for Alias clamping
    horizontalFieldOfView As Single      ' at Z = 1.0, this many X is visible; 2.0 = 90 degrees
    xOrigin As Single                    ' should probably allways be 0.5
    yOrigin As Single                    ' between be around 0.3 to 0.5
    vieworg As vec3_t
    viewangles As vec3_t
    fov_x As Single
    fov_y As Single
    ambientlight As Long
End Type

Public r_refdef As refdef_t
Public r_cache_thrash As Boolean     ' (r_main.c / d_surf.c)
Public r_notexture_mip As Long       ' (r_main.c) textures() index of the checkerboard
Public r_origin As vec3_t            ' (r_main.c)
Public vpn As vec3_t
Public vright As vec3_t
Public vup As vec3_t

'=============================================================================
' protocol.h

Public Const PROTOCOL_VERSION As Long = 15

' if the high bit of the servercmd is set, the low bits are fast update flags:
Public Const U_MOREBITS As Long = &H1
Public Const U_ORIGIN1 As Long = &H2
Public Const U_ORIGIN2 As Long = &H4
Public Const U_ORIGIN3 As Long = &H8
Public Const U_ANGLE2 As Long = &H10
Public Const U_NOLERP As Long = &H20     ' don't interpolate movement
Public Const U_FRAME As Long = &H40
Public Const U_SIGNAL As Long = &H80     ' just differentiates from other updates
' svc_update can pass all of the fast update bits, plus more
Public Const U_ANGLE1 As Long = &H100
Public Const U_ANGLE3 As Long = &H200
Public Const U_MODEL As Long = &H400
Public Const U_COLORMAP As Long = &H800
Public Const U_SKIN As Long = &H1000
Public Const U_EFFECTS As Long = &H2000
Public Const U_LONGENTITY As Long = &H4000

Public Const SU_VIEWHEIGHT As Long = &H1
Public Const SU_IDEALPITCH As Long = &H2
Public Const SU_PUNCH1 As Long = &H4
Public Const SU_PUNCH2 As Long = &H8
Public Const SU_PUNCH3 As Long = &H10
Public Const SU_VELOCITY1 As Long = &H20
Public Const SU_VELOCITY2 As Long = &H40
Public Const SU_VELOCITY3 As Long = &H80
Public Const SU_ITEMS As Long = &H200
Public Const SU_ONGROUND As Long = &H400     ' no data follows, the bit is it
Public Const SU_INWATER As Long = &H800      ' no data follows, the bit is it
Public Const SU_WEAPONFRAME As Long = &H1000
Public Const SU_ARMOR As Long = &H2000
Public Const SU_WEAPON As Long = &H4000

' a sound with no channel is a local only sound
Public Const SND_VOLUME As Long = &H1        ' a byte
Public Const SND_ATTENUATION As Long = &H2   ' a byte
Public Const SND_LOOPING As Long = &H4       ' a long

' defaults for clientinfo messages
Public Const DEFAULT_VIEWHEIGHT As Long = 22

' game types sent by serverinfo
' these determine which intermission screen plays
Public Const GAME_COOP As Long = 0
Public Const GAME_DEATHMATCH As Long = 1

' server to client
Public Const svc_bad As Long = 0
Public Const svc_nop As Long = 1
Public Const svc_disconnect As Long = 2
Public Const svc_updatestat As Long = 3      ' [byte] [long]
Public Const svc_version As Long = 4         ' [long] server version
Public Const svc_setview As Long = 5         ' [short] entity number
Public Const svc_sound As Long = 6           ' <see code>
Public Const svc_time As Long = 7            ' [float] server time
Public Const svc_print As Long = 8           ' [string] null terminated string
Public Const svc_stufftext As Long = 9       ' [string] stuffed into client's console buffer
Public Const svc_setangle As Long = 10       ' [angle3] set the view angle to this absolute value
Public Const svc_serverinfo As Long = 11     ' [long] version [string] signon string ...
Public Const svc_lightstyle As Long = 12     ' [byte] [string]
Public Const svc_updatename As Long = 13     ' [byte] [string]
Public Const svc_updatefrags As Long = 14    ' [byte] [short]
Public Const svc_clientdata As Long = 15     ' <shortbits + data>
Public Const svc_stopsound As Long = 16      ' <see code>
Public Const svc_updatecolors As Long = 17   ' [byte] [byte]
Public Const svc_particle As Long = 18       ' [vec3] <variable>
Public Const svc_damage As Long = 19
Public Const svc_spawnstatic As Long = 20
Public Const svc_spawnbaseline As Long = 22
Public Const svc_temp_entity As Long = 23
Public Const svc_setpause As Long = 24       ' [byte] on / off
Public Const svc_signonnum As Long = 25      ' [byte]  used for the signon sequence
Public Const svc_centerprint As Long = 26    ' [string] to put in center of the screen
Public Const svc_killedmonster As Long = 27
Public Const svc_foundsecret As Long = 28
Public Const svc_spawnstaticsound As Long = 29   ' [coord3] [byte] samp [byte] vol [byte] aten
Public Const svc_intermission As Long = 30   ' [string] music
Public Const svc_finale As Long = 31         ' [string] music [string] text
Public Const svc_cdtrack As Long = 32        ' [byte] track [byte] looptrack
Public Const svc_sellscreen As Long = 33
Public Const svc_cutscene As Long = 34

' client to server
Public Const clc_bad As Long = 0
Public Const clc_nop As Long = 1
Public Const clc_disconnect As Long = 2
Public Const clc_move As Long = 3            ' [usercmd_t]
Public Const clc_stringcmd As Long = 4       ' [string] message

' temp entity events
Public Const TE_SPIKE As Long = 0
Public Const TE_SUPERSPIKE As Long = 1
Public Const TE_GUNSHOT As Long = 2
Public Const TE_EXPLOSION As Long = 3
Public Const TE_TAREXPLOSION As Long = 4
Public Const TE_LIGHTNING1 As Long = 5
Public Const TE_LIGHTNING2 As Long = 6
Public Const TE_WIZSPIKE As Long = 7
Public Const TE_KNIGHTSPIKE As Long = 8
Public Const TE_LIGHTNING3 As Long = 9
Public Const TE_LAVASPLASH As Long = 10
Public Const TE_TELEPORT As Long = 11
Public Const TE_EXPLOSION2 As Long = 12
Public Const TE_BEAM As Long = 13

' sound.h
Public Const DEFAULT_SOUND_PACKET_VOLUME As Long = 255
Public Const DEFAULT_SOUND_PACKET_ATTENUATION As Single = 1!

' sound.h (cluster F). An sfx_t * is a known_sfx() index (snd_dma.bas),
' 1-based, 0 = NULL; its cache is the sfxcaches() entry of the same number.

Public Type sfx_t
    name As String                       ' up to MAX_QPATH - 1 characters
    cache As Long                        ' C: cache_user_t: the sfxcaches() index once loaded, 0 = not loaded
End Type

Public Type sfxcache_t
    length As Long
    loopstart As Long
    speed As Long
    width As Long
    stereo As Long
    ' C: byte data[1], variable sized: one of these, by width
    data8() As Byte                      ' width 1: the samples as signed char bit patterns
    data16() As Integer                  ' width 2
End Type

Public Type dma_t
    gamealive As Boolean
    soundalive As Boolean
    splitbuffer As Boolean
    channels As Long
    samples As Long                      ' mono samples in buffer
    submission_chunk As Long             ' don't mix less than this #
    samplepos As Long                    ' in mono samples
    samplebits As Long
    speed As Long
    buffer As Long                       ' C: unsigned char *: the address of snd_win.bas's dma_buffer()
End Type

Public Type channel_t
    sfx As Long                          ' known_sfx() index, 0 = NULL
    leftvol As Long                      ' 0-255 volume
    rightvol As Long                     ' 0-255 volume
    end_ As Long                         ' C: end; end time in global paintsamples
    pos As Long                          ' sample position in sfx
    looping As Long                      ' where to loop, -1 = no looping
    entnum As Long                       ' to allow overriding a specific sound
    entchannel As Long
    origin As vec3_t                     ' origin of sound effect
    dist_mult As Single                  ' distance multiplier (attenuation/clipK)
    master_vol As Long                   ' 0-255 master volume
End Type

Public Type wavinfo_t
    rate As Long
    width As Long
    channels As Long
    loopstart As Long
    samples As Long
    dataofs As Long                      ' chunk starts this many bytes from file start
End Type

Public Const MAX_CHANNELS As Long = 128
Public Const MAX_DYNAMIC_CHANNELS As Long = 8

' quakedef.h
Public Type entity_state_t
    origin As vec3_t
    angles As vec3_t
    modelindex As Long
    frame As Long
    colormap As Long
    skin As Long
    effects As Long
End Type

'=============================================================================
' bspfile.h

Public Const MAX_MAP_HULLS As Long = 4
Public Const MAX_MAP_MODELS As Long = 256
Public Const MAX_MAP_LEAFS As Long = 8192
Public Const BSPVERSION As Long = 29
Public Const MIPLEVELS As Long = 4
Public Const MAXLIGHTMAPS As Long = 4
Public Const NUM_AMBIENTS As Long = 4    ' automatic ambient sounds
Public Const AMBIENT_WATER As Long = 0
Public Const AMBIENT_SKY As Long = 1
Public Const AMBIENT_SLIME As Long = 2
Public Const AMBIENT_LAVA As Long = 3

Public Const LUMP_ENTITIES As Long = 0
Public Const LUMP_PLANES As Long = 1
Public Const LUMP_TEXTURES As Long = 2
Public Const LUMP_VERTEXES As Long = 3
Public Const LUMP_VISIBILITY As Long = 4
Public Const LUMP_NODES As Long = 5
Public Const LUMP_TEXINFO As Long = 6
Public Const LUMP_FACES As Long = 7
Public Const LUMP_LIGHTING As Long = 8
Public Const LUMP_CLIPNODES As Long = 9
Public Const LUMP_LEAFS As Long = 10
Public Const LUMP_MARKSURFACES As Long = 11
Public Const LUMP_EDGES As Long = 12
Public Const LUMP_SURFEDGES As Long = 13
Public Const LUMP_MODELS As Long = 14
Public Const HEADER_LUMPS As Long = 15

' 0-2 are axial planes, 3-5 are non-axial planes snapped to the nearest
Public Const PLANE_X As Long = 0
Public Const PLANE_Y As Long = 1
Public Const PLANE_Z As Long = 2
Public Const PLANE_ANYX As Long = 3
Public Const PLANE_ANYY As Long = 4
Public Const PLANE_ANYZ As Long = 5

Public Const CONTENTS_EMPTY As Long = -1
Public Const CONTENTS_SOLID As Long = -2
Public Const CONTENTS_WATER As Long = -3
Public Const CONTENTS_SLIME As Long = -4
Public Const CONTENTS_LAVA As Long = -5
Public Const CONTENTS_SKY As Long = -6
Public Const CONTENTS_ORIGIN As Long = -7    ' removed at csg time
Public Const CONTENTS_CLIP As Long = -8      ' changed to contents_solid
Public Const CONTENTS_CURRENT_0 As Long = -9
Public Const CONTENTS_CURRENT_90 As Long = -10
Public Const CONTENTS_CURRENT_180 As Long = -11
Public Const CONTENTS_CURRENT_270 As Long = -12
Public Const CONTENTS_CURRENT_UP As Long = -13
Public Const CONTENTS_CURRENT_DOWN As Long = -14

Public Const TEX_SPECIAL As Long = 1     ' sky or slime, no lightmap or 256 subdivision

Public Type dmodel_t
    mins As vec3_t
    maxs As vec3_t
    origin As vec3_t
    headnode(0 To MAX_MAP_HULLS - 1) As Long
    visleafs As Long                     ' not including the solid leaf 0
    firstface As Long
    numfaces As Long
End Type

Public Type dclipnode_t
    planenum As Long
    children(0 To 1) As Long             ' C: short; negative numbers are contents
End Type

'=============================================================================
' model.h: all brush model data lives in pools (model.bas) that every loaded
' brush model shares; a model_t holds the index of its first element in
' each pool (C: a pointer to its own array). C pointers between these
' structures are pool indices; C's int indices (planenum, firstedge, clip
' node children) stay relative to the model's arrays as in C.

Public Const SIDE_FRONT As Long = 0
Public Const SIDE_BACK As Long = 1
Public Const SIDE_ON As Long = 2

Public Type mvertex_t
    position As vec3_t
End Type

' textures(1 ..), 0 = NULL. The pixels of the four mip levels are in
' texpixels() at offsets(0 .. 3) (C: offsets from the texture_t)
Public Type texture_t
    name As String
    width As Long
    height As Long
    anim_total As Long                   ' total tenths in sequence ( 0 = no)
    anim_min As Long                     ' time for this frame min <=time< max
    anim_max As Long
    anim_next As Long                    ' in the animation sequence
    alternate_anims As Long              ' bmodels in frmae 1 use these
    offsets(0 To MIPLEVELS - 1) As Long  ' four mip maps stored
End Type

Public Const SURF_PLANEBACK As Long = 2
Public Const SURF_DRAWSKY As Long = 4
Public Const SURF_DRAWSPRITE As Long = 8
Public Const SURF_DRAWTURB As Long = &H10
Public Const SURF_DRAWTILED As Long = &H20
Public Const SURF_DRAWBACKGROUND As Long = &H40

Public Type medge_t
    v(0 To 1) As Long                    ' C: unsigned short
    cachededgeoffset As Long
End Type

Public Type mtexinfo_t
    vecs(0 To 1, 0 To 3) As Single
    mipadjust As Single
    texture As Long                      ' textures() index
    flags As Long
End Type

Public Type msurface_t
    visframe As Long                     ' should be drawn when node is crossed
    dlightframe As Long
    dlightbits As Long
    plane As Long                        ' mod_planes() index
    flags As Long
    firstedge As Long                    ' look up in model->surfedges[], negative numbers
    numedges As Long                     ' are backwards edges
    ' surface generation data
    cachespots(0 To MIPLEVELS - 1) As Long   ' surface cache index, 0 = NULL (cluster C)
    texturemins(0 To 1) As Long          ' C: short
    extents(0 To 1) As Long              ' C: short
    texinfo As Long                      ' mod_texinfo() index
    ' lighting info
    styles(0 To MAXLIGHTMAPS - 1) As Byte
    samples As Long                      ' offset in mod_lightdata(), -1 = NULL
End Type

' mnode_t and mleaf_t share their first fields in C, and node children
' point to either; here both live in mod_nodes(), with the fields of both.
' contents is 0 for a node, negative for a leaf, as in C.
Public Type mnode_t
    ' common with leaf
    contents As Long                     ' 0, to differentiate from leafs
    visframe As Long                     ' node needs to be traversed if current
    minmaxs(0 To 5) As Long              ' for bounding box culling (C: short)
    parent As Long                       ' mod_nodes() index, -1 = NULL
    ' node specific
    plane As Long                        ' mod_planes() index
    children(0 To 1) As Long             ' mod_nodes() indices
    firstsurface As Long                 ' C: unsigned short
    numsurfaces As Long
    ' leaf specific
    compressed_vis As Long               ' offset in mod_visdata(), -1 = NULL
    efrags As Long                       ' (cluster C) 0 = NULL
    firstmarksurface As Long             ' mod_marksurfaces() index
    nummarksurfaces As Long
    key As Long                          ' BSP sequence number for leaf's contents
    ambient_sound_level(0 To NUM_AMBIENTS - 1) As Byte
End Type

Public Type hull_t
    clipnodes As Long                    ' mod_clipnodes() index of node 0
    planes As Long                       ' mod_planes() index of plane 0
    firstclipnode As Long
    lastclipnode As Long
    clip_mins As vec3_t
    clip_maxs As vec3_t
End Type

' sprites (spritegn.h, model.h): msprites(1 ..), sprite_frames(1 ..), sprite_groups(1 ..)

Public Const SPRITE_VERSION As Long = 1
Public Const SPR_VP_PARALLEL_UPRIGHT As Long = 0
Public Const SPR_FACING_UPRIGHT As Long = 1
Public Const SPR_VP_PARALLEL As Long = 2
Public Const SPR_ORIENTED As Long = 3
Public Const SPR_VP_PARALLEL_ORIENTED As Long = 4
Public Const SPR_SINGLE As Long = 0
Public Const SPR_GROUP As Long = 1

Public Type mspriteframe_t
    width As Long
    height As Long
    pcachespot As Long                   ' remove?
    up As Single
    down As Single
    left_ As Single
    right_ As Single
    pixels As Long                       ' offset in sprite_pixels()
End Type

Public Type mspritegroup_t
    numframes As Long
    intervals As Long                    ' index of the first in sprite_intervals()
    frames() As Long                     ' sprite_frames() indices
End Type

Public Type mspriteframedesc_t
    typ As Long                          ' spriteframetype_t
    frameptr As Long                     ' SPR_SINGLE: sprite_frames() index, SPR_GROUP: sprite_groups() index
End Type

Public Type msprite_t
    typ As Long
    maxwidth As Long
    maxheight As Long
    numframes As Long
    beamlength As Single                 ' remove?
    cachespot As Long                    ' remove?
    frames() As mspriteframedesc_t
End Type

' alias models (modelgen.h, model.h): aliashdrs(1 ..). C keeps a model in one
' relocatable block addressed by byte offsets; here the parts are in pools
' (model.bas) and the offsets are pool indices.

Public Const ALIAS_VERSION As Long = 6
Public Const ALIAS_ONSEAM As Long = &H20
Public Const ALIAS_SINGLE As Long = 0
Public Const ALIAS_GROUP As Long = 1
Public Const ALIAS_SKIN_SINGLE As Long = 0
Public Const ALIAS_SKIN_GROUP As Long = 1
Public Const DT_FACES_FRONT As Long = &H10
Public Const ST_SYNC As Long = 0
Public Const ST_RAND As Long = 1

Public Type mdl_t
    ident As Long
    version As Long
    scale As vec3_t
    scale_origin As vec3_t
    boundingradius As Single
    eyeposition As vec3_t
    numskins As Long
    skinwidth As Long
    skinheight As Long
    numverts As Long
    numtris As Long
    numframes As Long
    synctype As Long
    flags As Long
    size As Single
End Type

Public Type stvert_t
    onseam As Long
    s As Long
    t As Long
End Type

Public Type mtriangle_t
    facesfront As Long
    vertindex(0 To 2) As Long
End Type

Public Type trivertx_t
    v(0 To 2) As Byte
    lightnormalindex As Byte
End Type

Public Type maliasframedesc_t
    typ As Long                          ' aliasframetype_t
    bboxmin As trivertx_t
    bboxmax As trivertx_t
    frame As Long                        ' ALIAS_SINGLE: alias_verts() index, ALIAS_GROUP: alias_groups() index
    name As String
End Type

Public Type maliasskindesc_t
    typ As Long                          ' aliasskintype_t
    pcachespot As Long
    skin As Long                         ' ALIAS_SKIN_SINGLE: alias_skins() offset, ALIAS_SKIN_GROUP: alias_skingroups() index
End Type

Public Type maliasgroupframedesc_t
    bboxmin As trivertx_t
    bboxmax As trivertx_t
    frame As Long                        ' alias_verts() index
End Type

Public Type maliasgroup_t
    numframes As Long
    intervals As Long                    ' index of the first in alias_intervals()
    frames() As maliasgroupframedesc_t
End Type

Public Type maliasskingroup_t
    numskins As Long
    intervals As Long                    ' index of the first in alias_intervals()
    skindescs() As maliasskindesc_t
End Type

Public Type aliashdr_t
    mdl As mdl_t                         ' C: model, the offset of the mdl_t
    stverts As Long                      ' alias_stverts() index
    skindesc As Long                     ' alias_skindescs() index
    triangles As Long                    ' alias_tris() index
    frames() As maliasframedesc_t
End Type

' Whole model: mod_known(1 .. MAX_MOD_KNOWN), 0 = NULL

' modtype_t
Public Const mod_brush As Long = 0
Public Const mod_sprite As Long = 1
Public Const mod_alias As Long = 2

Public Const EF_ROCKET As Long = 1       ' leave a trail
Public Const EF_GRENADE As Long = 2      ' leave a trail
Public Const EF_GIB As Long = 4          ' leave a trail
Public Const EF_ROTATE As Long = 8       ' rotate (bonus items)
Public Const EF_TRACER As Long = 16      ' green split trail
Public Const EF_ZOMGIB As Long = 32      ' small blood trail
Public Const EF_TRACER2 As Long = 64     ' orange split trail + rotate
Public Const EF_TRACER3 As Long = 128    ' purple trail

Public Type model_t
    name As String
    needload As Long                     ' bmodels and sprites don't cache normally
    typ As Long                          ' modtype_t
    numframes As Long
    synctype As Long
    flags As Long
    ' volume occupied by the model
    mins As vec3_t
    maxs As vec3_t
    radius As Single
    ' brush model: pool indices of the first element (C: pointers)
    firstmodelsurface As Long
    nummodelsurfaces As Long
    numsubmodels As Long
    submodels As Long
    numplanes As Long
    planes As Long
    numleafs As Long                     ' number of visible leafs, not counting 0
    leafs As Long                        ' mod_nodes() index of leaf 0
    numvertexes As Long
    vertexes As Long
    numedges As Long
    edges As Long
    numnodes As Long
    nodes As Long                        ' mod_nodes() index of node 0
    numtexinfo As Long
    texinfo As Long
    numsurfaces As Long
    surfaces As Long
    numsurfedges As Long
    surfedges As Long
    numclipnodes As Long
    clipnodes As Long
    nummarksurfaces As Long
    marksurfaces As Long
    hulls(0 To MAX_MAP_HULLS - 1) As hull_t
    numtextures As Long
    textures As Long                     ' mod_texlist() index; -1 = NULL
    visdata As Long                      ' mod_visdata() offset, -1 = NULL
    lightdata As Long                    ' mod_lightdata() offset, -1 = NULL
    entities As Long                     ' mod_entdata() offset, -1 = NULL
    ' additional model data (C: cache, only access through Mod_Extradata):
    ' aliashdrs() index for alias models, msprites() index for sprites, 0 = none
    cache As Long
End Type

'=============================================================================
' world.h

Public Type plane_t
    normal As vec3_t
    dist As Single
End Type

Public Type trace_t
    allsolid As Boolean                  ' if true, plane is not valid
    startsolid As Boolean                ' if true, the initial point was in a solid area
    inopen As Boolean
    inwater As Boolean
    fraction As Single                   ' time completed, 1.0 = didn't hit anything
    endpos As vec3_t                     ' final position
    plane As plane_t                     ' surface normal at impact
    ent As Long                          ' entity the surface is on; -1 = NULL
End Type

Public Const MOVE_NORMAL As Long = 0
Public Const MOVE_NOMONSTERS As Long = 1
Public Const MOVE_MISSILE As Long = 2

'=============================================================================
' progs.h: an edict_t* is the edict number, -1 = NULL (0 is the world).
' The engine's part of edict n is edicts(n); its progs fields are ev(n)
' (entvars_t), laid over the Long array ed_l() (pr_edict.bas).

Public Const MAX_ENT_LEAFS As Long = 16

' pr_comp.h
Public Const OFS_NULL As Long = 0
Public Const OFS_RETURN As Long = 1
Public Const OFS_PARM0 As Long = 4       ' leave 3 ofs for each parm to hold vectors
Public Const OFS_PARM1 As Long = 7
Public Const OFS_PARM2 As Long = 10
Public Const OFS_PARM3 As Long = 13
Public Const OFS_PARM4 As Long = 16
Public Const OFS_PARM5 As Long = 19
Public Const OFS_PARM6 As Long = 22
Public Const OFS_PARM7 As Long = 25
Public Const RESERVED_OFS As Long = 28

' etype_t
Public Const ev_void As Long = 0
Public Const ev_string As Long = 1
Public Const ev_float As Long = 2
Public Const ev_vector As Long = 3
Public Const ev_entity As Long = 4
Public Const ev_field As Long = 5
Public Const ev_function As Long = 6
Public Const ev_pointer As Long = 7

Public Const DEF_SAVEGLOBAL As Long = &H8000&
Public Const MAX_PARMS As Long = 8
Public Const PROG_VERSION As Long = 6

Public Type dstatement_t
    op As Long                           ' C: unsigned short
    a As Long                            ' C: short
    b As Long
    c As Long
End Type

Public Type ddef_t
    typ As Long                          ' if DEF_SAVEGLOBAL bit is set
                                         ' the variable needs to be saved in savegames
    ofs As Long                          ' C: unsigned short
    s_name As Long
End Type

Public Type dfunction_t
    first_statement As Long              ' negative numbers are builtins
    parm_start As Long
    locals As Long                       ' total ints of parms + locals
    profile As Long                      ' runtime
    s_name As Long
    s_file As Long                       ' source file defined in
    numparms As Long
    parm_size(0 To MAX_PARMS - 1) As Long    ' C: byte
End Type

Public Type edict_t
    free As Boolean
    ' area: the link is link_prev(n) / link_next(n) in world.bas
    num_leafs As Long
    leafnums(0 To MAX_ENT_LEAFS - 1) As Long   ' C: short
    baseline As entity_state_t
    freetime As Single                   ' sv.time when the object was freed
End Type

'=============================================================================
' net.h

Public Const NET_NAMELEN As Long = 64

Public Const NET_MAXMESSAGE As Long = 8192
Public Const NET_HEADERSIZE As Long = 8          ' 2 * sizeof(unsigned int)
Public Const NET_DATAGRAMSIZE As Long = MAX_DATAGRAM + NET_HEADERSIZE

' NetHeader flags
Public Const NETFLAG_LENGTH_MASK As Long = &HFFFF&
Public Const NETFLAG_DATA As Long = &H10000
Public Const NETFLAG_ACK As Long = &H20000
Public Const NETFLAG_NAK As Long = &H40000
Public Const NETFLAG_EOM As Long = &H80000
Public Const NETFLAG_UNRELIABLE As Long = &H100000
Public Const NETFLAG_CTL As Long = &H80000000

Public Const NET_PROTOCOL_VERSION As Long = 3

' This is the network info/connection protocol.  It is used to find Quake
' servers, get info about them, and connect to them.  Once connected, the
' Quake game protocol (documented elsewhere) is used.
' (the message layouts are documented in net.h)
Public Const CCREQ_CONNECT As Long = &H1
Public Const CCREQ_SERVER_INFO As Long = &H2
Public Const CCREQ_PLAYER_INFO As Long = &H3
Public Const CCREQ_RULE_INFO As Long = &H4

Public Const CCREP_ACCEPT As Long = &H81
Public Const CCREP_REJECT As Long = &H82
Public Const CCREP_SERVER_INFO As Long = &H83
Public Const CCREP_PLAYER_INFO As Long = &H84
Public Const CCREP_RULE_INFO As Long = &H85

Public Type qsockaddr
    sa_family As Integer
    sa_data(0 To 13) As Byte
End Type

' net_qsockets(1 .. net_numsockets), 0 = NULL
Public Type qsocket_t
    nxt As Long                          ' C: next (net_qsockets() index, 0 = NULL)
    connecttime As Double
    lastMessageTime As Double
    lastSendTime As Double

    disconnected As Boolean
    canSend As Boolean
    sendNext As Boolean

    driver As Long
    landriver As Long
    socket As Long
    driverdata As Long                   ' loopback: the qsocket at the other end, 0 = NULL

    ackSequence As Long
    sendSequence As Long
    unreliableSendSequence As Long
    sendMessageLength As Long
    sendMessage(0 To NET_MAXMESSAGE - 1) As Byte

    receiveSequence As Long
    unreliableReceiveSequence As Long
    receiveMessageLength As Long
    receiveMessage(0 To NET_MAXMESSAGE - 1) As Byte

    addr As qsockaddr
    address As String                    ' up to NET_NAMELEN - 1 characters
End Type

Public net_qsockets() As qsocket_t
Public net_activeSockets As Long         ' net_qsockets() index, 0 = NULL
Public net_freeSockets As Long
Public net_numsockets As Long

' net_drivers(): the dispatch is a Select Case on the driver number in
' net_main.bas (0 Loopback, 1 Datagram)
Public Const MAX_NET_DRIVERS As Long = 8

Public Type net_driver_t
    name As String
    initialized As Boolean
    controlSock As Long
End Type

Public net_numdrivers As Long
Public net_drivers(0 To MAX_NET_DRIVERS - 1) As net_driver_t

' net_landrivers(): the datagram driver's LAN drivers, dispatched in
' net_dgrm.bas (0 Winsock TCPIP; there is no Winsock IPX)
Public Type net_landriver_t
    name As String
    initialized As Boolean
    controlSock As Long
End Type

Public net_numlandrivers As Long
Public net_landrivers(0 To MAX_NET_DRIVERS - 1) As net_landriver_t

Public DEFAULTnet_hostport As Long
Public net_hostport As Long

Public net_driverlevel As Long
Public cv_hostname As Long

Public messagesSent As Long
Public messagesReceived As Long
Public unreliableMessagesSent As Long
Public unreliableMessagesReceived As Long

Public Const HOSTCACHESIZE As Long = 8

Public Type hostcache_t
    name As String                       ' up to 15 characters
    map As String                        ' up to 15
    cname As String                      ' up to 31
    users As Long
    maxusers As Long
    driver As Long
    ldriver As Long
    addr As qsockaddr
End Type

Public hostCacheCount As Long
Public hostcache(0 To HOSTCACHESIZE - 1) As hostcache_t

' PollProcedure: C keeps a linked list of structs with a function pointer.
' Here each procedure is a PP_ number (net_main.bas dispatches them), and
' poll_next() / poll_time() hold the list.
Public Const PP_SLIST_SEND As Long = 1
Public Const PP_SLIST_POLL As Long = 2
Public Const PP_TEST_POLL As Long = 3    ' net_dgrm.bas
Public Const PP_TEST2_POLL As Long = 4
Public Const MAX_POLL_PROCEDURES As Long = 8

Public net_time As Double
Public net_message As sizebuf_t
Public net_activeconnections As Long

Public serialAvailable As Boolean
Public ipxAvailable As Boolean
Public tcpipAvailable As Boolean
Public my_ipx_address As String
Public my_tcpip_address As String

Public slistInProgress As Boolean
Public slistSilent As Boolean
Public slistLocal As Boolean

'=============================================================================
' render.h (the entity parts)

Public Const MAXCLIPPLANES As Long = 11

Public Const TOP_RANGE As Long = 16      ' soldier uniform colors
Public Const BOTTOM_RANGE As Long = 96

' cl_efrags(1 .. MAX_EFRAGS), 0 = NULL
Public Type efrag_t
    leaf As Long                         ' mod_nodes() index
    leafnext As Long                     ' cl_efrags() index
    entity As Long                       ' cl_entities() index
    entnext As Long                      ' cl_efrags() index
End Type

' C's entity_t * points into cl_entities, cl_static_entities,
' cl_temp_entities or at cl.viewent; here all of them are in one array,
' cl_entities(), and an entity is its index there (-1 = NULL)
Public Type entity_t
    forcelink As Boolean                 ' model changed

    update_type As Long

    baseline As entity_state_t           ' to fill in defaults in updates

    msgtime As Double                    ' time of last update
    msg_origins(0 To 1) As vec3_t        ' last two updates (0 is newest)
    origin As vec3_t
    msg_angles(0 To 1) As vec3_t         ' last two updates (0 is newest)
    angles As vec3_t
    model As Long                        ' mod_known() index, 0 = no model
    efrag As Long                        ' linked list of efrags (cl_efrags() index, 0 = NULL)
    frame As Long
    syncbase As Single                   ' for client-side animations
    colormap As Long                     ' 0 = vid.colormap, n = cl.scores(n - 1).translations
    effects As Long                      ' light, particals, etc
    skinnum As Long                      ' for Alias models
    visframe As Long                     ' last frame this entity was
                                         '  found in an active leaf

    dlightframe As Long                  ' dynamic lighting
    dlightbits As Long

    ' FIXME: could turn these into a union
    trivial_accept As Long
    topnode As Long                      ' for bmodels, first world node
                                         '  that splits bmodel (mod_nodes() index), or
                                         '  NULL if not split (cluster C decides the NULL)
End Type

'=============================================================================
' r_shared.h, r_local.h, d_local.h, d_iface.h

Public Const MAXVERTS As Long = 16               ' max points in a surface polygon
Public Const MAXWORKINGVERTS As Long = MAXVERTS + 4   ' max points in an intermediate
                                                 '  polygon (while processing)
Public Const MAXHEIGHT As Long = 1024
Public Const MAXWIDTH As Long = 1280
Public Const MAXDIMENSION As Long = 1280         ' (MAXHEIGHT > MAXWIDTH) ? MAXHEIGHT : MAXWIDTH

Public Const CYCLE As Long = 128                 ' turbulent cycle size
Public Const SIN_BUFFER_SIZE As Long = MAXDIMENSION + CYCLE

Public Const INFINITE_DISTANCE As Long = &H10000 ' distance that's always guaranteed to
                                                 '  be farther away than anything in
                                                 '  the scene

Public Const NUMSTACKEDGES As Long = 2400
Public Const MINEDGES As Long = NUMSTACKEDGES
Public Const NUMSTACKSURFACES As Long = 800
Public Const MINSURFACES As Long = NUMSTACKSURFACES
Public Const MAXSPANS As Long = 3000

' spans: one array (d_edge's spans), -1 = NULL
Public Type espan_t
    u As Long
    v As Long
    count As Long
    pnext As Long                        ' espan_t index
End Type

' surfaces(): surfaces are generated in back to front order by the bsp, so
' if a surf index is greater than another one, it should be drawn in front.
' surfaces(1) is the background, and is used as the active surface stack.
' surfaces(0) is a dummy, because index 0 is used to indicate no surface
'  attached to an edge_t
Public Type surf_t
    nxt As Long                          ' C: next; active surface stack in r_edge.c
    prev As Long                         ' used in r_edge.c for active surf stack
    spans As Long                        ' linked list of spans to draw (espan_t index, -1 = NULL)
    key As Long                          ' sorting key (BSP order)
    last_u As Long                       ' set during tracing
    spanstate As Long                    ' 0 = not in span
                                         ' 1 = in span
                                         ' -1 = in inverted span (end before
                                         '  start)
    flags As Long                        ' currentface flags
    data As Long                         ' associated data: the mod_surfaces() index
    entity As Long                       ' cl_entities() index
    nearzi As Single                     ' nearest 1/z on surface, for mipmapping
    insubmodel As Boolean
    d_ziorigin As Single
    d_zistepu As Single
    d_zistepv As Single
End Type

' flags in finalvert_t.flags
Public Const ALIAS_LEFT_CLIP As Long = &H1
Public Const ALIAS_TOP_CLIP As Long = &H2
Public Const ALIAS_RIGHT_CLIP As Long = &H4
Public Const ALIAS_BOTTOM_CLIP As Long = &H8
Public Const ALIAS_Z_CLIP As Long = &H10
' (ALIAS_ONSEAM is with modelgen.h's definitions above)
Public Const ALIAS_XY_CLIP_MASK As Long = &HF

' r_edges(): the four fixed edges first (EDGE_HEAD .. EDGE_SENTINEL), then
' the frame's edges from EDGES_BASE (C: r_edges[0]); -1 = NULL
Public Type edge_t
    u As Long                            ' fixed16_t
    u_step As Long
    prev As Long
    nxt As Long                          ' C: next
    surfs(0 To 1) As Long                ' C: unsigned short, surfaces() indices
    nextremove As Long
    nearzi As Single
    owner As Long                        ' mod_edges() index (-1: a dummy medge_t)
End Type

Public Const ALIAS_BASE_SIZE_RATIO As Double = 1# / 11#  ' normalizing factor so player model works out to about
                                                 '  1 pixel per triangle

Public Const BMODEL_FULLY_CLIPPED As Long = &H10 ' value returned by R_BmodelCheckBBox ()
                                                 '  if bbox is trivially rejected

' viewmodel lighting
Public Type alight_t
    ambientlight As Long
    shadelight As Long
    plightvec As vec3_t                  ' C: float * to the light vector
End Type

' clipped bmodel edges: v() >= 0 is a mod_vertexes() index, v() < 0 the
' clipped vertex r_bverts(-1 - v)
Public Type bedge_t
    v(0 To 1) As Long
    pnext As Long                        ' bedge_t index, -1 = NULL
End Type

Public Const XCENTERING As Double = 1# / 2#
Public Const YCENTERING As Double = 1# / 2#

Public Const CLIP_EPSILON As Double = 0.001

Public Const BACKFACE_EPSILON As Double = 0.01

Public Const DIST_NOT_SET As Long = 98765

' view_clipplanes(0 .. 3); next is an index there, -1 = NULL
Public Type clipplane_t
    normal As vec3_t
    dist As Single
    nxt As Long                          ' C: next
    leftedge As Byte
    rightedge As Byte
End Type

Public Const NEAR_CLIP As Double = 0.01

Public Const MAXBVERTINDEXES As Long = 1000      ' new clipped vertices when clipping bmodels
                                                 '  to the world BSP

Public Type btofpoly_t
    clipflags As Long
    psurf As Long                        ' mod_surfaces() index
End Type

Public Const MAX_BTOFPOLYS As Long = 5000        ' FIXME: tune this

' turbulence stuff
Public Const AMP As Long = 8 * &H10000
Public Const AMP2 As Long = 3
Public Const SPEED_ As Long = 20                ' C: SPEED (the locals speed would hide it)

' d_local.h

' TODO: fine-tune this; it's based on providing some overage even if there
' is a 2k-wide scan, with subdivision every 8, for 256 spans of 12 bytes each
Public Const SCANBUFFERPAD As Long = &H1000

Public Const R_SKY_SMASK As Long = &H7F0000
Public Const R_SKY_TMASK As Long = &H7F0000

Public Const DS_SPAN_LIST_END As Long = -128

Public Const SURFCACHE_SIZE_AT_320X200 As Long = 600& * 1024

' surfcache_t: the blocks live in d_surf.bas's surface cache memory, a
' header followed by the pixels; a block is its byte offset + 1, 0 = NULL
' (see d_surf.bas)

' d_iface.h

Public Const MAX_LBM_HEIGHT As Long = 480

Public Type emitpoint_t
    u As Single
    v As Single
    s As Single
    t As Single
    zi As Single
End Type

Public Const PARTICLE_Z_CLIP As Single = 8!

Public Type polyvert_t
    u As Single
    v As Single
    zi As Single
    s As Single
    t As Single
End Type

Public Type polydesc_t
    numverts As Long
    nearzi As Single
    pcurrentface As Long                 ' mod_surfaces() index
    ' C: polyvert_t *pverts (r_draw.bas's r_polyverts())
End Type

Public Type zpointdesc_t
    u As Long
    v As Long
    zi As Single
    color As Long
End Type

' transparency types for D_DrawRect ()
Public Const DR_SOLID As Long = 0
Public Const DR_TRANSPARENT As Long = 1

' callbacks to Quake
Public Type drawsurf_t
    surfdat As Long                      ' destination for generated surface (surface cache byte offset)
    rowbytes As Long                     ' destination logical width in bytes
    surf As Long                         ' description for surface to generate (mod_surfaces() index)
    lightadj(0 To MAXLIGHTMAPS - 1) As Long  ' adjust for lightmap levels for dynamic lighting
    texture As Long                      ' corrected for animating textures (textures() index)
    surfmip As Long                      ' mipmapped ratio of surface texels / world pixels
    surfwidth As Long                    ' in mipmapped texels
    surfheight As Long                   ' in mipmapped texels
End Type

Public Const TURB_TEX_SIZE As Long = 64          ' base turbulent texture size

Public Const TILE_SIZE As Long = 128             ' size of textures generated by R_GenTiledSurf

Public Const SKYSHIFT As Long = 7
Public Const SKYSIZE As Long = 128               ' 1 << SKYSHIFT
Public Const SKYMASK As Long = SKYSIZE - 1

' models, sprites and particles (cluster D)

Public Const MAXALIASVERTS As Long = 2000        ' TODO: tune this
Public Const ALIAS_Z_CLIP_PLANE As Long = 5

' ptype_t
Public Const pt_static As Long = 0
Public Const pt_grav As Long = 1
Public Const pt_slowgrav As Long = 2
Public Const pt_fire As Long = 3
Public Const pt_explode As Long = 4
Public Const pt_explode2 As Long = 5
Public Const pt_blob As Long = 6
Public Const pt_blob2 As Long = 7

' particles(): r_part.bas's pool; nxt is a particles() index, -1 = NULL
Public Type particle_t
    ' driver-usable fields
    org As vec3_t
    color As Single
    ' drivers never touch the following fields
    nxt As Long                          ' C: next
    vel As vec3_t
    ramp As Single
    die As Single
    typ As Long                          ' ptype_t
End Type

' a_finalverts(): r_alias.bas's model vertices, then r_aclip.bas's clip buffers
Public Type finalvert_t
    v(0 To 5) As Long                    ' u, v, s, t, l, 1/z
    flags As Long
    reserved As Single
End Type

Public Type auxvert_t
    fv(0 To 2) As Single                 ' viewspace x, y
End Type

Public Type affinetridesc_t
    pskin As Long                        ' alias_skins() offset
    ' (C's pskindesc is set but never read)
    skinwidth As Long
    skinheight As Long
    ptriangles As Long                   ' alias_tris() index; -1 = r_aclip.bas's aclip_tri
    pfinalverts As Long                  ' a_finalverts() index
    numtriangles As Long
    drawtype As Long
    seamfixupX16 As Long
End Type

' sprite spans (d_local.h)
Public Type sspan_t
    u As Long
    v As Long
    count As Long
End Type

Public Type spritedesc_t
    nump As Long
    ' C: emitpoint_t *pverts (r_sprite.bas's r_spriteverts(); there's room for
    ' an extra element at [nump], if the driver wants to duplicate element
    ' [0] at element [nump] to avoid dealing with wrapping)
    pspriteframe As Long                 ' sprite_frames() index
    vup As vec3_t                        ' in worldspace
    vright As vec3_t
    vpn As vec3_t
    nearzi As Single
End Type

' a VB array whose memory is another array's (pr_edict.bas, d_vars.bas)
Public Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    cElements As Long
    lLbound As Long
End Type

'=============================================================================
' client.h

Public Type usercmd_t
    viewangles As vec3_t
    ' intended velocities
    forwardmove As Single
    sidemove As Single
    upmove As Single
End Type

Public Type lightstyle_t
    length As Long
    map As String                        ' up to MAX_STYLESTRING - 1 characters
End Type

Public Type scoreboard_t
    name As String                       ' up to MAX_SCOREBOARDNAME - 1 characters
    entertime As Single
    frags As Long
    colors As Long                       ' two 4 bit fields
    translations(0 To VID_GRADES * 256 - 1) As Byte
End Type

Public Type cshift_t
    destcolor(0 To 2) As Long
    percent As Long                      ' 0-256
End Type

Public Const CSHIFT_CONTENTS As Long = 0
Public Const CSHIFT_DAMAGE As Long = 1
Public Const CSHIFT_BONUS As Long = 2
Public Const CSHIFT_POWERUP As Long = 3
Public Const NUM_CSHIFTS As Long = 4

Public Const NAME_LENGTH As Long = 64

' client_state_t should hold all pieces of the client state

Public Const SIGNONS As Long = 4         ' signon messages to receive before connected

Public Const MAX_DLIGHTS As Long = 32

Public Type dlight_t
    origin As vec3_t
    radius As Single
    die As Single                        ' stop lighting after this time
    decay As Single                      ' drop this each second
    minlight As Single                   ' don't add when contributing less
    key As Long
End Type

Public Const MAX_BEAMS As Long = 24

Public Type beam_t
    entity As Long
    model As Long                        ' mod_known() index, 0 = NULL
    endtime As Single
    start As vec3_t
    end_ As vec3_t
End Type

Public Const MAX_EFRAGS As Long = 640

Public Const MAX_MAPSTRING As Long = 2048
Public Const MAX_DEMOS As Long = 8
Public Const MAX_DEMONAME As Long = 16

' cactive_t
Public Const ca_dedicated As Long = 0    ' a dedicated server with no ability to start a client
Public Const ca_disconnected As Long = 1 ' full screen console with no connection
Public Const ca_connected As Long = 2    ' valid netcon, talking to a server

' the client_static_t structure is persistant through an arbitrary number
' of server connections
Public Type client_static_t
    state As Long

    ' personalization data sent to server
    mapstring As String
    spawnparms As String                 ' to restart a level

    ' demo loop control
    demonum As Long                      ' -1 = don't play demos
    demos(0 To MAX_DEMOS - 1) As String  ' when not playing

    ' demo recording info must be here, because record is started before
    ' entering a map (and clearing client_state_t)
    demorecording As Boolean
    demoplayback As Boolean
    timedemo As Boolean
    forcetrack As Long                   ' -1 = use normal cd track
    demofile As Long                     ' a Win32 file handle, 0 = NULL
    td_lastframe As Long                 ' to meter out one message a frame
    td_startframe As Long                ' host_framecount at start
    td_starttime As Single               ' realtime at second frame of timedemo

    ' connection information
    signon As Long                       ' 0 to SIGNONS
    netcon As Long                       ' net_qsockets() index, 0 = NULL
    message As sizebuf_t                 ' writing buffer to send to server
End Type

Public cls As client_static_t

' the client_state_t structure is wiped completely at every
' server signon
Public Type client_state_t
    movemessages As Long                 ' since connecting to this server
                                         ' throw out the first couple, so the player
                                         ' doesn't accidentally do something the
                                         ' first frame
    cmd As usercmd_t                     ' last command sent to the server

    ' information for local display
    stats(0 To MAX_CL_STATS - 1) As Long ' health, etc
    items As Long                        ' inventory bit flags
    item_gettime(0 To 31) As Single      ' cl.time of aquiring item, for blinking
    faceanimtime As Single               ' use anim frame if cl.time < this

    cshifts(0 To NUM_CSHIFTS - 1) As cshift_t       ' color shifts for damage, powerups
    prev_cshifts(0 To NUM_CSHIFTS - 1) As cshift_t  ' and content types

    ' the client maintains its own idea of view angles, which are
    ' sent to the server each frame.  The server sets punchangle when
    ' the view is temporarliy offset, and an angle reset commands at the start
    ' of each level and after teleporting.
    mviewangles(0 To 1) As vec3_t        ' during demo playback viewangles is lerped
                                         ' between these
    viewangles As vec3_t

    mvelocity(0 To 1) As vec3_t          ' update by server, used for lean+bob
                                         ' (0 is newest)
    velocity As vec3_t                   ' lerped between mvelocity[0] and [1]

    punchangle As vec3_t                 ' temporary offset

    ' pitch drifting vars
    idealpitch As Single
    pitchvel As Single
    nodrift As Boolean
    driftmove As Single
    laststop As Double

    viewheight As Single
    crouch As Single                     ' local amount for smoothing stepups

    paused As Boolean                    ' send over by server
    onground As Boolean
    inwater As Boolean

    intermission As Long                 ' don't change view angle, full screen, etc
    completed_time As Long               ' latched at intermission start

    mtime(0 To 1) As Double              ' the timestamp of last two messages
    time As Double                       ' clients view of time, should be between
                                         ' servertime and oldservertime to generate
                                         ' a lerp point for other data
    oldtime As Double                    ' previous cl.time, time-oldtime is used
                                         ' to decay light values and smooth step ups

    last_received_message As Single      ' (realtime) for net trouble icon

    '
    ' information that is static for the entire time connected to a server
    '
    model_precache(0 To MAX_MODELS - 1) As Long  ' mod_known() indices, 0 = NULL
    sound_precache(0 To MAX_SOUNDS - 1) As Long  ' (cluster F) sfx handles, 0 = NULL

    levelname As String                  ' for display on solo scoreboard (up to 39 characters)
    viewentity As Long                   ' cl_entitites[cl.viewentity] = player
    maxclients As Long
    gametype As Long

    ' refresh related state
    worldmodel As Long                   ' cl_entitites[0].model
    free_efrags As Long                  ' cl_efrags() index, 0 = NULL
    num_entities As Long                 ' held in cl_entities array
    num_statics As Long                  ' held in cl_staticentities array
    ' C: entity_t viewent, the gun model: here cl_entities(CL_VIEWENT)

    cdtrack As Long                      ' cd audio
    looptrack As Long

    ' frag scoreboard
    scores() As scoreboard_t             ' [cl.maxclients]
End Type

Public cl As client_state_t

' cvars
Public cv_cl_name As Long                ' (cl_main.c)
Public cv_cl_color As Long

Public cv_cl_upspeed As Long             ' (cl_input.c)
Public cv_cl_forwardspeed As Long
Public cv_cl_backspeed As Long
Public cv_cl_sidespeed As Long

Public cv_cl_movespeedkey As Long

Public cv_cl_yawspeed As Long
Public cv_cl_pitchspeed As Long

Public cv_cl_anglespeedkey As Long

Public cv_cl_autofire As Long

Public cv_cl_shownet As Long             ' (cl_main.c)
Public cv_cl_nolerp As Long

Public cv_cl_pitchdriftspeed As Long
Public cv_lookspring As Long
Public cv_lookstrafe As Long
Public cv_sensitivity As Long

Public cv_m_pitch As Long
Public cv_m_yaw As Long
Public cv_m_forward As Long
Public cv_m_side As Long

Public Const MAX_TEMP_ENTITIES As Long = 64      ' lightning bolts, etc
Public Const MAX_STATIC_ENTITIES As Long = 128   ' torches, etc

' (QuakeVB) where C's entity arrays start in cl_entities()
Public Const CL_STATIC_BASE As Long = MAX_EDICTS                            ' cl_static_entities[0]
Public Const CL_TEMP_BASE As Long = CL_STATIC_BASE + MAX_STATIC_ENTITIES    ' cl_temp_entities[0]
Public Const CL_VIEWENT As Long = CL_TEMP_BASE + MAX_TEMP_ENTITIES          ' cl.viewent
Public Const CL_NUM_ENTITIES As Long = CL_VIEWENT + 1

' FIXME, allocate dynamically
Public cl_efrags() As efrag_t            ' (1 .. MAX_EFRAGS)
Public cl_entities() As entity_t         ' (0 .. CL_NUM_ENTITIES - 1)
Public cl_lightstyle(0 To MAX_LIGHTSTYLES - 1) As lightstyle_t
Public cl_dlights(0 To MAX_DLIGHTS - 1) As dlight_t
Public cl_beams(0 To MAX_BEAMS - 1) As beam_t

Public Const MAX_VISEDICTS As Long = 256
Public cl_numvisedicts As Long
Public cl_visedicts(0 To MAX_VISEDICTS - 1) As Long  ' cl_entities() indices

' cl_input
Public Type kbutton_t
    down(0 To 1) As Long                 ' key nums holding it down
    state As Long                        ' low bit is down state
End Type

Public in_mlook As kbutton_t
Public in_klook As kbutton_t
Public in_strafe As kbutton_t
Public in_speed As kbutton_t

'=============================================================================
' server.h

Public Const NUM_PING_TIMES As Long = 16
Public Const NUM_SPAWN_PARMS As Long = 16

Public Type client_t
    active As Boolean                    ' false = client is free
    spawned As Boolean                   ' false = don't send datagrams
    dropasap As Boolean                  ' has been told to go to another level
    privileged As Boolean                ' can execute any host command
    sendsignon As Boolean                ' only valid before spawned
    last_message As Double               ' reliable messages must be sent
                                         ' periodically
    netconnection As Long                ' communications handle (net_qsockets() index, 0 = NULL)
    cmd As usercmd_t                     ' movement
    wishdir As vec3_t                    ' intended motion calced from cmd
    message As sizebuf_t                 ' can be added to at any time,
                                         ' copied and clear once per frame
    edict As Long                        ' EDICT_NUM(clientnum+1)
    name As String                       ' for printing to other people (SV_SetClientName sets it)
    colors As Long
    ping_times(0 To NUM_PING_TIMES - 1) As Single
    num_pings As Long                    ' ping_times[num_pings%NUM_PING_TIMES]
    ' spawn parms are carried from level to level
    spawn_parms(0 To NUM_SPAWN_PARMS - 1) As Single
    ' client known data for deltas
    old_frags As Long
End Type

Public Type server_static_t
    maxclients As Long
    maxclientslimit As Long
    clients() As client_t                ' [maxclientslimit]
    serverflags As Long                  ' episode completion information
    changelevel_issued As Boolean        ' cleared when at SV_SpawnServer
End Type

Public svs As server_static_t            ' persistant server info

' server_state_t
Public Const ss_loading As Long = 0
Public Const ss_active As Long = 1

Public Type server_t
    active As Boolean                    ' false if only a net client
    paused As Boolean
    loadgame As Boolean                  ' handle connections specially
    time As Double
    lastcheck As Long                    ' used by PF_checkclient
    lastchecktime As Double
    name As String                       ' map name
    modelname As String                  ' maps/<name>.bsp, for model_precache[0]
    worldmodel As Long                   ' mod_known() index, 0 = NULL
    ' C: NULL terminated lists of string pointers; here the first
    ' num_model_precache / num_sound_precache entries are set
    model_precache(0 To MAX_MODELS - 1) As String
    num_model_precache As Long
    models(0 To MAX_MODELS - 1) As Long  ' mod_known() indices, 0 = NULL
    sound_precache(0 To MAX_SOUNDS - 1) As String
    num_sound_precache As Long
    lightstyles(0 To MAX_LIGHTSTYLES - 1) As String  ' "" = NULL
    num_edicts As Long
    max_edicts As Long
    state As Long                        ' some actions are only valid during load
    datagram As sizebuf_t
    reliable_datagram As sizebuf_t       ' copied to all clients at end of frame
    signon As sizebuf_t
End Type

Public sv As server_t                    ' local server

' edict->movetype values
Public Const MOVETYPE_NONE As Long = 0           ' never moves
Public Const MOVETYPE_ANGLENOCLIP As Long = 1
Public Const MOVETYPE_ANGLECLIP As Long = 2
Public Const MOVETYPE_WALK As Long = 3           ' gravity
Public Const MOVETYPE_STEP As Long = 4           ' gravity, special edge handling
Public Const MOVETYPE_FLY As Long = 5
Public Const MOVETYPE_TOSS As Long = 6           ' gravity
Public Const MOVETYPE_PUSH As Long = 7           ' no clip to world, push and crush
Public Const MOVETYPE_NOCLIP As Long = 8
Public Const MOVETYPE_FLYMISSILE As Long = 9     ' extra size to monsters
Public Const MOVETYPE_BOUNCE As Long = 10

' edict->solid values
Public Const SOLID_NOT As Long = 0               ' no interaction with other objects
Public Const SOLID_TRIGGER As Long = 1           ' touch on edge, but not blocking
Public Const SOLID_BBOX As Long = 2              ' touch on edge, block
Public Const SOLID_SLIDEBOX As Long = 3          ' touch on edge, but not an onground
Public Const SOLID_BSP As Long = 4               ' bsp clip, touch on edge, block

' edict->deadflag values
Public Const DEAD_NO As Long = 0
Public Const DEAD_DYING As Long = 1
Public Const DEAD_DEAD As Long = 2

Public Const DAMAGE_NO As Long = 0
Public Const DAMAGE_YES As Long = 1
Public Const DAMAGE_AIM As Long = 2

' edict->flags
Public Const FL_FLY As Long = 1
Public Const FL_SWIM As Long = 2
Public Const FL_CONVEYOR As Long = 4
Public Const FL_CLIENT As Long = 8
Public Const FL_INWATER As Long = 16
Public Const FL_MONSTER As Long = 32
Public Const FL_GODMODE As Long = 64
Public Const FL_NOTARGET As Long = 128
Public Const FL_ITEM As Long = 256
Public Const FL_ONGROUND As Long = 512
Public Const FL_PARTIALGROUND As Long = 1024     ' not all corners are valid
Public Const FL_WATERJUMP As Long = 2048         ' player jumping out of water
Public Const FL_JUMPRELEASED As Long = 4096      ' for jump debouncing

' entity effects
Public Const EF_BRIGHTFIELD As Long = 1
Public Const EF_MUZZLEFLASH As Long = 2
Public Const EF_BRIGHTLIGHT As Long = 4
Public Const EF_DIMLIGHT As Long = 8

Public Const SPAWNFLAG_NOT_EASY As Long = 256
Public Const SPAWNFLAG_NOT_MEDIUM As Long = 512
Public Const SPAWNFLAG_NOT_HARD As Long = 1024
Public Const SPAWNFLAG_NOT_DEATHMATCH As Long = 2048

Public host_client As Long               ' current client (svs.clients() index, -1 = NULL)
Public sv_player As Long                 ' (sv_user.c) edict of host_client
Public current_skill As Long             ' skill level for currently loaded level
Public noclip_anglehack As Boolean

'=============================================================================
' host (quakedef.h)

Public host_parms As quakeparms_t
Public host_initialized As Boolean       ' true if into command execution
Public host_frametime As Double
Public host_basepal() As Byte
Public host_colormap() As Byte
Public host_framecount As Long           ' incremented every frame, never reset
Public realtime As Double                ' not bounded in any way, changed at
                                         ' start of every frame, never reset
Public isDedicated As Boolean
Public minimum_memory As Long

' host and other cvars that headers make global
Public cv_sys_ticrate As Long
Public cv_developer As Long
Public cv_registered As Long
Public cv_scr_viewsize As Long
Public cv_lcd_x As Long                  ' (view.c, not ported yet: the dummy)
