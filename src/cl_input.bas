Attribute VB_Name = "cl_input"
' cl_input.bas - builds an intended movement command to send to the server
' (cl_input.c)
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

'===============================================================================
' KEY BUTTONS
'
' Continuous button event tracking is complicated by the fact that two different
' input sources (say, mouse button 1 and the control key) can both press the
' same button, but the button should only be released when both of the
' pressing key have been released.
'
' When a key event issues a button command (+forward, +attack, etc), it appends
' its key number as a parameter to the command so it can be matched up with
' the release.
'
' state bit 0 is the current state of the key
' state bit 1 is edge triggered on the up to down transition
' state bit 2 is edge triggered on the down to up transition
'===============================================================================

' (in_mlook, in_klook, in_strafe and in_speed are in quakedef.bas)
Private in_left As kbutton_t, in_right As kbutton_t, in_forward As kbutton_t, in_back As kbutton_t
Private in_lookup As kbutton_t, in_lookdown As kbutton_t, in_moveleft As kbutton_t, in_moveright As kbutton_t
Private in_use As kbutton_t, in_jump As kbutton_t, in_attack As kbutton_t
Private in_up As kbutton_t, in_down As kbutton_t

Private in_impulse As Long

Private Sub KeyDown(b As kbutton_t)
    Dim k As Long, c As String

    c = Cmd_Argv(1)
    If Len(c) Then
        k = C_atoi(c)
    Else
        k = -1                           ' typed manually at the console for continuous down
    End If

    If k = b.down(0) Or k = b.down(1) Then Exit Sub   ' repeating key

    If b.down(0) = 0 Then
        b.down(0) = k
    ElseIf b.down(1) = 0 Then
        b.down(1) = k
    Else
        Con_Printf "Three keys down for a button!\n"
        Exit Sub
    End If

    If b.state And 1 Then Exit Sub       ' still down
    b.state = b.state Or (1 + 2)         ' down + impulse down
End Sub

Private Sub KeyUp(b As kbutton_t)
    Dim k As Long, c As String

    c = Cmd_Argv(1)
    If Len(c) Then
        k = C_atoi(c)
    Else
        ' typed manually at the console, assume for unsticking, so clear all
        b.down(0) = 0
        b.down(1) = 0
        b.state = 4                      ' impulse up
        Exit Sub
    End If

    If b.down(0) = k Then
        b.down(0) = 0
    ElseIf b.down(1) = k Then
        b.down(1) = 0
    Else
        Exit Sub                         ' key up without coresponding down (menu pass through)
    End If
    If b.down(0) <> 0 Or b.down(1) <> 0 Then Exit Sub   ' some other key is still holding it down

    If (b.state And 1) = 0 Then Exit Sub ' still up (this should not happen)
    b.state = b.state And Not 1          ' now up
    b.state = b.state Or 4               ' impulse up
End Sub

Private Sub IN_MLookUp()
    KeyUp in_mlook
    If (in_mlook.state And 1) = 0 And cvars(cv_lookspring).value <> 0 Then V_StartPitchDrift
End Sub

' (QuakeVB: IN_Impulse_, as VB takes in_impulse for the same name)
Private Sub IN_Impulse_()
    in_impulse = Q_atoi(Cmd_Argv(1))
End Sub

' Returns 0.25 if a key was pressed and released during the frame,
' 0.5 if it was pressed and held
' 0 if held then released, and
' 1.0 if held for the entire time
Public Function CL_KeyState(key As kbutton_t) As Single
    Dim val As Single, impulsedown As Boolean, impulseup As Boolean, down As Boolean

    impulsedown = (key.state And 2) <> 0
    impulseup = (key.state And 4) <> 0
    down = (key.state And 1) <> 0
    val = 0

    If impulsedown And Not impulseup Then
        If down Then
            val = 0.5                    ' pressed and held this frame
        Else
            val = 0                      '  I_Error ();
        End If
    End If
    If impulseup And Not impulsedown Then
        If down Then
            val = 0                      '  I_Error ();
        Else
            val = 0                      ' released this frame
        End If
    End If
    If Not impulsedown And Not impulseup Then
        If down Then
            val = 1                      ' held the entire frame
        Else
            val = 0                      ' up the entire frame
        End If
    End If
    If impulsedown And impulseup Then
        If down Then
            val = 0.75                   ' released and re-pressed this frame
        Else
            val = 0.25                   ' pressed and released this frame
        End If
    End If

    key.state = key.state And 1          ' clear impulses

    CL_KeyState = val
End Function

'==========================================================================
' (the cvars of cl_input.c are registered by CL_Init)

' Moves the local angle positions
Private Sub CL_AdjustAngles()
    Dim speed As Single, up As Single, down As Single

    If in_speed.state And 1 Then
        speed = host_frametime * cvars(cv_cl_anglespeedkey).value
    Else
        speed = host_frametime
    End If

    If (in_strafe.state And 1) = 0 Then
        cl.viewangles.v(YAW) = cl.viewangles.v(YAW) - speed * cvars(cv_cl_yawspeed).value * CL_KeyState(in_right)
        cl.viewangles.v(YAW) = cl.viewangles.v(YAW) + speed * cvars(cv_cl_yawspeed).value * CL_KeyState(in_left)
        cl.viewangles.v(YAW) = anglemod(cl.viewangles.v(YAW))
    End If
    If in_klook.state And 1 Then
        V_StopPitchDrift
        cl.viewangles.v(PITCH) = cl.viewangles.v(PITCH) - speed * cvars(cv_cl_pitchspeed).value * CL_KeyState(in_forward)
        cl.viewangles.v(PITCH) = cl.viewangles.v(PITCH) + speed * cvars(cv_cl_pitchspeed).value * CL_KeyState(in_back)
    End If

    up = CL_KeyState(in_lookup)
    down = CL_KeyState(in_lookdown)

    cl.viewangles.v(PITCH) = cl.viewangles.v(PITCH) - speed * cvars(cv_cl_pitchspeed).value * up
    cl.viewangles.v(PITCH) = cl.viewangles.v(PITCH) + speed * cvars(cv_cl_pitchspeed).value * down

    If up <> 0 Or down <> 0 Then V_StopPitchDrift

    If cl.viewangles.v(PITCH) > 80 Then cl.viewangles.v(PITCH) = 80
    If cl.viewangles.v(PITCH) < -70 Then cl.viewangles.v(PITCH) = -70

    If cl.viewangles.v(ROLL) > 50 Then cl.viewangles.v(ROLL) = 50
    If cl.viewangles.v(ROLL) < -50 Then cl.viewangles.v(ROLL) = -50
End Sub

' Send the intended movement message to the server
Public Sub CL_BaseMove(cmd_ As usercmd_t)
    Dim blank As usercmd_t

    If cls.signon <> SIGNONS Then Exit Sub

    CL_AdjustAngles

    cmd_ = blank

    If in_strafe.state And 1 Then
        cmd_.sidemove = cmd_.sidemove + cvars(cv_cl_sidespeed).value * CL_KeyState(in_right)
        cmd_.sidemove = cmd_.sidemove - cvars(cv_cl_sidespeed).value * CL_KeyState(in_left)
    End If

    cmd_.sidemove = cmd_.sidemove + cvars(cv_cl_sidespeed).value * CL_KeyState(in_moveright)
    cmd_.sidemove = cmd_.sidemove - cvars(cv_cl_sidespeed).value * CL_KeyState(in_moveleft)

    cmd_.upmove = cmd_.upmove + cvars(cv_cl_upspeed).value * CL_KeyState(in_up)
    cmd_.upmove = cmd_.upmove - cvars(cv_cl_upspeed).value * CL_KeyState(in_down)

    If (in_klook.state And 1) = 0 Then
        cmd_.forwardmove = cmd_.forwardmove + cvars(cv_cl_forwardspeed).value * CL_KeyState(in_forward)
        cmd_.forwardmove = cmd_.forwardmove - cvars(cv_cl_backspeed).value * CL_KeyState(in_back)
    End If

    '
    ' adjust for speed key
    '
    If in_speed.state And 1 Then
        cmd_.forwardmove = cmd_.forwardmove * cvars(cv_cl_movespeedkey).value
        cmd_.sidemove = cmd_.sidemove * cvars(cv_cl_movespeedkey).value
        cmd_.upmove = cmd_.upmove * cvars(cv_cl_movespeedkey).value
    End If
End Sub

Public Sub CL_SendMove(cmd_ As usercmd_t)
    Dim i As Long, bits As Long, buf As sizebuf_t

    SZ_Alloc buf, 128

    cl.cmd = cmd_

    '
    ' send the movement message
    '
    MSG_WriteByte buf, clc_move

    MSG_WriteFloat buf, cl.mtime(0)      ' so server can get ping times

    For i = 0 To 2
        MSG_WriteAngle buf, cl.viewangles.v(i)
    Next

    MSG_WriteShort buf, cmd_.forwardmove
    MSG_WriteShort buf, cmd_.sidemove
    MSG_WriteShort buf, cmd_.upmove

    '
    ' send button bits
    '
    bits = 0

    If in_attack.state And 3 Then bits = bits Or 1
    in_attack.state = in_attack.state And Not 2

    If in_jump.state And 3 Then bits = bits Or 2
    in_jump.state = in_jump.state And Not 2

    MSG_WriteByte buf, bits

    MSG_WriteByte buf, in_impulse
    in_impulse = 0

    '
    ' deliver the message
    '
    If cls.demoplayback Then Exit Sub

    '
    ' allways dump the first two message, because it may contain leftover inputs
    ' from the last level
    '
    cl.movemessages = cl.movemessages + 1
    If cl.movemessages <= 2 Then Exit Sub

    If NET_SendUnreliableMessage(cls.netcon, buf) = -1 Then
        Con_Printf "CL_SendMove: lost server connection\n"
        CL_Disconnect
    End If
End Sub

Public Sub CL_InitInput()
    Cmd_AddCommand "+moveup", XC_IN_UpDown
    Cmd_AddCommand "-moveup", XC_IN_UpUp
    Cmd_AddCommand "+movedown", XC_IN_DownDown
    Cmd_AddCommand "-movedown", XC_IN_DownUp
    Cmd_AddCommand "+left", XC_IN_LeftDown
    Cmd_AddCommand "-left", XC_IN_LeftUp
    Cmd_AddCommand "+right", XC_IN_RightDown
    Cmd_AddCommand "-right", XC_IN_RightUp
    Cmd_AddCommand "+forward", XC_IN_ForwardDown
    Cmd_AddCommand "-forward", XC_IN_ForwardUp
    Cmd_AddCommand "+back", XC_IN_BackDown
    Cmd_AddCommand "-back", XC_IN_BackUp
    Cmd_AddCommand "+lookup", XC_IN_LookupDown
    Cmd_AddCommand "-lookup", XC_IN_LookupUp
    Cmd_AddCommand "+lookdown", XC_IN_LookdownDown
    Cmd_AddCommand "-lookdown", XC_IN_LookdownUp
    Cmd_AddCommand "+strafe", XC_IN_StrafeDown
    Cmd_AddCommand "-strafe", XC_IN_StrafeUp
    Cmd_AddCommand "+moveleft", XC_IN_MoveleftDown
    Cmd_AddCommand "-moveleft", XC_IN_MoveleftUp
    Cmd_AddCommand "+moveright", XC_IN_MoverightDown
    Cmd_AddCommand "-moveright", XC_IN_MoverightUp
    Cmd_AddCommand "+speed", XC_IN_SpeedDown
    Cmd_AddCommand "-speed", XC_IN_SpeedUp
    Cmd_AddCommand "+attack", XC_IN_AttackDown
    Cmd_AddCommand "-attack", XC_IN_AttackUp
    Cmd_AddCommand "+use", XC_IN_UseDown
    Cmd_AddCommand "-use", XC_IN_UseUp
    Cmd_AddCommand "+jump", XC_IN_JumpDown
    Cmd_AddCommand "-jump", XC_IN_JumpUp
    Cmd_AddCommand "impulse", XC_IN_Impulse
    Cmd_AddCommand "+klook", XC_IN_KLookDown
    Cmd_AddCommand "-klook", XC_IN_KLookUp
    Cmd_AddCommand "+mlook", XC_IN_MLookDown
    Cmd_AddCommand "-mlook", XC_IN_MLookUp
End Sub

' (QuakeVB) the console commands of this module, for Cmd_Call
' (C: IN_UpDown () { KeyDown (&in_up); } and so on)
Public Sub CL_InputCommand(ByVal fn As Long)
    Select Case fn
    Case XC_IN_KLookDown: KeyDown in_klook
    Case XC_IN_KLookUp: KeyUp in_klook
    Case XC_IN_MLookDown: KeyDown in_mlook
    Case XC_IN_MLookUp: IN_MLookUp
    Case XC_IN_UpDown: KeyDown in_up
    Case XC_IN_UpUp: KeyUp in_up
    Case XC_IN_DownDown: KeyDown in_down
    Case XC_IN_DownUp: KeyUp in_down
    Case XC_IN_LeftDown: KeyDown in_left
    Case XC_IN_LeftUp: KeyUp in_left
    Case XC_IN_RightDown: KeyDown in_right
    Case XC_IN_RightUp: KeyUp in_right
    Case XC_IN_ForwardDown: KeyDown in_forward
    Case XC_IN_ForwardUp: KeyUp in_forward
    Case XC_IN_BackDown: KeyDown in_back
    Case XC_IN_BackUp: KeyUp in_back
    Case XC_IN_LookupDown: KeyDown in_lookup
    Case XC_IN_LookupUp: KeyUp in_lookup
    Case XC_IN_LookdownDown: KeyDown in_lookdown
    Case XC_IN_LookdownUp: KeyUp in_lookdown
    Case XC_IN_MoveleftDown: KeyDown in_moveleft
    Case XC_IN_MoveleftUp: KeyUp in_moveleft
    Case XC_IN_MoverightDown: KeyDown in_moveright
    Case XC_IN_MoverightUp: KeyUp in_moveright
    Case XC_IN_SpeedDown: KeyDown in_speed
    Case XC_IN_SpeedUp: KeyUp in_speed
    Case XC_IN_StrafeDown: KeyDown in_strafe
    Case XC_IN_StrafeUp: KeyUp in_strafe
    Case XC_IN_AttackDown: KeyDown in_attack
    Case XC_IN_AttackUp: KeyUp in_attack
    Case XC_IN_UseDown: KeyDown in_use
    Case XC_IN_UseUp: KeyUp in_use
    Case XC_IN_JumpDown: KeyDown in_jump
    Case XC_IN_JumpUp: KeyUp in_jump
    Case XC_IN_Impulse: IN_Impulse_
    Case Else: Sys_Error "CL_InputCommand: bad function %i", fn
    End Select
End Sub
