Attribute VB_Name = "in_win"
' in_win.bas - windows 95 mouse and joystick code (in_win.c)
' 02/21/97 JCB Added extended DirectInput code to support external controllers.
'
' Not ported: DirectInput (-dinput); the mouse is read with GetCursorPos,
' C's path without DirectInput. Under -hidden (test runs) the mouse and the
' joystick stay off, so the real devices of the PC cannot get in.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type JOYINFOEX
    dwSize As Long
    dwFlags As Long
    dwXpos As Long
    dwYpos As Long
    dwZpos As Long
    dwRpos As Long
    dwUpos As Long
    dwVpos As Long
    dwButtons As Long
    dwButtonNumber As Long
    dwPOV As Long
    dwReserved1 As Long
    dwReserved2 As Long
End Type

Private Type JOYCAPS
    wMid As Integer
    wPid As Integer
    szPname(0 To 31) As Byte
    wXmin As Long
    wXmax As Long
    wYmin As Long
    wYmax As Long
    wZmin As Long
    wZmax As Long
    wNumButtons As Long
    wPeriodMin As Long
    wPeriodMax As Long
    wRmin As Long
    wRmax As Long
    wUmin As Long
    wUmax As Long
    wVmin As Long
    wVmax As Long
    wCaps As Long
    wMaxAxes As Long
    wNumAxes As Long
    wMaxButtons As Long
    szRegKey(0 To 31) As Byte
    szOEMVxD(0 To 259) As Byte
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As RECT) As Long
Private Declare Function ClipCursorNull Lib "user32" Alias "ClipCursor" (ByVal lpRect As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function joyGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function joyGetPosEx Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFOEX) As Long
Private Declare Function joyGetDevCaps Lib "winmm.dll" Alias "joyGetDevCapsA" (ByVal id As Long, lpCaps As JOYCAPS, ByVal uSize As Long) As Long

Private Const SPI_GETMOUSE As Long = 3
Private Const SPI_SETMOUSE As Long = 4

Private Const JOYERR_NOERROR As Long = 0
Private Const JOY_RETURNX As Long = &H1
Private Const JOY_RETURNY As Long = &H2
Private Const JOY_RETURNZ As Long = &H4
Private Const JOY_RETURNR As Long = &H8
Private Const JOY_RETURNU As Long = &H10
Private Const JOY_RETURNV As Long = &H20
Private Const JOY_RETURNPOV As Long = &H40
Private Const JOY_RETURNBUTTONS As Long = &H80
Private Const JOY_RETURNCENTERED As Long = &H400
Private Const JOYCAPS_HASPOV As Long = &H10
Private Const JOY_POVCENTERED As Long = &HFFFF&
Private Const JOY_POVFORWARD As Long = 0
Private Const JOY_POVRIGHT As Long = 9000
Private Const JOY_POVBACKWARD As Long = 18000
Private Const JOY_POVLEFT As Long = 27000

' mouse variables
Private cv_m_filter As Long

Private mouse_buttons As Long
Private mouse_oldbuttonstate As Long
Private current_pos As POINTAPI
Private mouse_x As Long, mouse_y As Long, old_mouse_x As Long, old_mouse_y As Long
Private mx_accum As Long, my_accum As Long

Private restore_spi As Boolean
Private originalmouseparms(0 To 2) As Long, newmouseparms(0 To 2) As Long

Public uiWheelMessage As Long
Public mouseactive As Boolean
Public mouseinitialized As Boolean
Private mouseparmsvalid As Boolean, mouseactivatetoggle As Boolean
Private mousehidden As Boolean           ' C: !mouseshowtoggle (which starts at 1)

' joystick defines and variables
' where should defines be moved?
Private Const JOY_ABSOLUTE_AXIS As Long = &H0    ' control like a joystick
Private Const JOY_RELATIVE_AXIS As Long = &H10   ' control like a mouse, spinner, trackball
Private Const JOY_MAX_AXES As Long = 6           ' X, Y, Z, R, U, V
Private Const JOY_AXIS_X As Long = 0
Private Const JOY_AXIS_Y As Long = 1
Private Const JOY_AXIS_Z As Long = 2
Private Const JOY_AXIS_R As Long = 3
Private Const JOY_AXIS_U As Long = 4
Private Const JOY_AXIS_V As Long = 5

' _ControlList
Private Const AxisNada As Long = 0
Private Const AxisForward As Long = 1
Private Const AxisLook As Long = 2
Private Const AxisSide As Long = 3
Private Const AxisTurn As Long = 4

Private dwAxisFlags(0 To JOY_MAX_AXES - 1) As Long
Private dwAxisMap(0 To JOY_MAX_AXES - 1) As Long
Private dwControlMap(0 To JOY_MAX_AXES - 1) As Long
Private pdwRawValue(0 To JOY_MAX_AXES - 1) As Long   ' C: pointers to the ji fields; here the JOY_AXIS_ number

' none of these cvars are saved over a session
' this means that advanced controller configuration needs to be executed
' each time.  this avoids any problems with getting back to a default usage
' or when changing from one controller to another.  this way at least something
' works.
Private cv_in_joystick As Long
Private cv_joy_name As Long
Private cv_joy_advanced As Long
Private cv_joy_advaxisx As Long
Private cv_joy_advaxisy As Long
Private cv_joy_advaxisz As Long
Private cv_joy_advaxisr As Long
Private cv_joy_advaxisu As Long
Private cv_joy_advaxisv As Long
Private cv_joy_forwardthreshold As Long
Private cv_joy_sidethreshold As Long
Private cv_joy_pitchthreshold As Long
Private cv_joy_yawthreshold As Long
Private cv_joy_forwardsensitivity As Long
Private cv_joy_sidesensitivity As Long
Private cv_joy_pitchsensitivity As Long
Private cv_joy_yawsensitivity As Long
Private cv_joy_wwhack1 As Long
Private cv_joy_wwhack2 As Long

Private joy_avail As Boolean, joy_advancedinit As Boolean, joy_haspov As Boolean
Private joy_oldbuttonstate As Long, joy_oldpovstate As Long

Private joy_id As Long
Private joy_flags As Long
Private joy_numbuttons As Long

Private ji As JOYINFOEX

Private Sub Force_CenterView_f()
    cl.viewangles.v(PITCH) = 0
End Sub

' (QuakeVB) C's window_rect
Private Sub GetWindowRect_(r As RECT)
    r.Left = window_x
    r.Top = window_y
    r.Right = window_x + window_width
    r.Bottom = window_y + window_height
End Sub

Public Sub IN_UpdateClipCursor()
    Dim r As RECT

    If mouseinitialized And mouseactive Then
        GetWindowRect_ r
        ClipCursor r
    End If
End Sub

Public Sub IN_ShowMouse()
    If mousehidden Then
        ShowCursor 1
        mousehidden = False
    End If
End Sub

Public Sub IN_HideMouse()
    If Not mousehidden Then
        ShowCursor 0
        mousehidden = True
    End If
End Sub

Public Sub IN_ActivateMouse()
    Dim r As RECT

    mouseactivatetoggle = True

    If mouseinitialized Then
        If mouseparmsvalid Then restore_spi = SystemParametersInfo(SPI_SETMOUSE, 0, newmouseparms(0), 0) <> 0

        SetCursorPos window_center_x, window_center_y
        SetCapture mainwindow
        GetWindowRect_ r
        ClipCursor r

        mouseactive = True
    End If
End Sub

Public Sub IN_SetQuakeMouseState()
    If mouseactivatetoggle Then IN_ActivateMouse
End Sub

Public Sub IN_DeactivateMouse()
    mouseactivatetoggle = False

    If mouseinitialized Then
        If restore_spi Then SystemParametersInfo SPI_SETMOUSE, 0, originalmouseparms(0), 0

        ClipCursorNull 0
        ReleaseCapture

        mouseactive = False
    End If
End Sub

Public Sub IN_RestoreOriginalMouseState()
    If mouseactivatetoggle Then
        IN_DeactivateMouse
        mouseactivatetoggle = True
    End If

    ' try to redraw the cursor so it gets reinitialized, because sometimes it
    ' has garbage after the mode switch
    ShowCursor 1
    ShowCursor 0
End Sub

Private Sub IN_StartupMouse()
    If COM_CheckParm("-nomouse") Then Exit Sub
    If sys_hidden Then Exit Sub          ' (QuakeVB)

    mouseinitialized = True

    mouseparmsvalid = SystemParametersInfo(SPI_GETMOUSE, 0, originalmouseparms(0), 0) <> 0

    If mouseparmsvalid Then
        If COM_CheckParm("-noforcemspd") Then newmouseparms(2) = originalmouseparms(2)

        If COM_CheckParm("-noforcemaccel") Then
            newmouseparms(0) = originalmouseparms(0)
            newmouseparms(1) = originalmouseparms(1)
        End If

        If COM_CheckParm("-noforcemparms") Then
            newmouseparms(0) = originalmouseparms(0)
            newmouseparms(1) = originalmouseparms(1)
            newmouseparms(2) = originalmouseparms(2)
        End If
    End If

    mouse_buttons = 3

    ' if a fullscreen video mode was set before the mouse was initialized,
    ' set the mouse state appropriately
    If mouseactivatetoggle Then IN_ActivateMouse
End Sub

Public Sub IN_Init()
    ' C: static int newmouseparms[3] = {0, 0, 1}
    newmouseparms(0) = 0
    newmouseparms(1) = 0
    newmouseparms(2) = 1
    dwAxisFlags(0) = JOY_RETURNX
    dwAxisFlags(1) = JOY_RETURNY
    dwAxisFlags(2) = JOY_RETURNZ
    dwAxisFlags(3) = JOY_RETURNR
    dwAxisFlags(4) = JOY_RETURNU
    dwAxisFlags(5) = JOY_RETURNV

    ' mouse variables
    cv_m_filter = Cvar_RegisterVariable("m_filter", "0")

    ' joystick variables
    cv_in_joystick = Cvar_RegisterVariable("joystick", "0", True)
    cv_joy_name = Cvar_RegisterVariable("joyname", "joystick")
    cv_joy_advanced = Cvar_RegisterVariable("joyadvanced", "0")
    cv_joy_advaxisx = Cvar_RegisterVariable("joyadvaxisx", "0")
    cv_joy_advaxisy = Cvar_RegisterVariable("joyadvaxisy", "0")
    cv_joy_advaxisz = Cvar_RegisterVariable("joyadvaxisz", "0")
    cv_joy_advaxisr = Cvar_RegisterVariable("joyadvaxisr", "0")
    cv_joy_advaxisu = Cvar_RegisterVariable("joyadvaxisu", "0")
    cv_joy_advaxisv = Cvar_RegisterVariable("joyadvaxisv", "0")
    cv_joy_forwardthreshold = Cvar_RegisterVariable("joyforwardthreshold", "0.15")
    cv_joy_sidethreshold = Cvar_RegisterVariable("joysidethreshold", "0.15")
    cv_joy_pitchthreshold = Cvar_RegisterVariable("joypitchthreshold", "0.15")
    cv_joy_yawthreshold = Cvar_RegisterVariable("joyyawthreshold", "0.15")
    cv_joy_forwardsensitivity = Cvar_RegisterVariable("joyforwardsensitivity", "-1.0")
    cv_joy_sidesensitivity = Cvar_RegisterVariable("joysidesensitivity", "-1.0")
    cv_joy_pitchsensitivity = Cvar_RegisterVariable("joypitchsensitivity", "1.0")
    cv_joy_yawsensitivity = Cvar_RegisterVariable("joyyawsensitivity", "-1.0")
    cv_joy_wwhack1 = Cvar_RegisterVariable("joywwhack1", "0.0")
    cv_joy_wwhack2 = Cvar_RegisterVariable("joywwhack2", "0.0")

    Cmd_AddCommand "force_centerview", XC_Force_CenterView_f
    Cmd_AddCommand "joyadvancedupdate", XC_Joy_AdvancedUpdate_f

    uiWheelMessage = RegisterWindowMessage("MSWHEEL_ROLLMSG")

    IN_StartupMouse
    IN_StartupJoystick
End Sub

Public Sub IN_Shutdown()
    IN_DeactivateMouse
    IN_ShowMouse
End Sub

Public Sub IN_MouseEvent(ByVal mstate As Long)
    Dim i As Long

    If mouseactive Then
        ' perform button actions
        For i = 0 To mouse_buttons - 1
            If (mstate And BIT_(i)) <> 0 And (mouse_oldbuttonstate And BIT_(i)) = 0 Then Key_Event K_MOUSE1 + i, True

            If (mstate And BIT_(i)) = 0 And (mouse_oldbuttonstate And BIT_(i)) <> 0 Then Key_Event K_MOUSE1 + i, False
        Next

        mouse_oldbuttonstate = mstate
    End If
End Sub

Private Sub IN_MouseMove(cmd_ As usercmd_t)
    Dim mx As Long, my As Long

    If Not mouseactive Then Exit Sub

    GetCursorPos current_pos
    mx = current_pos.x - window_center_x + mx_accum
    my = current_pos.y - window_center_y + my_accum
    mx_accum = 0
    my_accum = 0

    ' if (mx ||  my)
    '     Con_DPrintf("mx=%d, my=%d\n", mx, my);

    If cvars(cv_m_filter).value <> 0 Then
        mouse_x = ToInt((mx + old_mouse_x) * 0.5)
        mouse_y = ToInt((my + old_mouse_y) * 0.5)
    Else
        mouse_x = mx
        mouse_y = my
    End If

    old_mouse_x = mx
    old_mouse_y = my

    mouse_x = ToInt(mouse_x * cvars(cv_sensitivity).value)
    mouse_y = ToInt(mouse_y * cvars(cv_sensitivity).value)

    ' add mouse X/Y movement to cmd
    If (in_strafe.state And 1) <> 0 Or (cvars(cv_lookstrafe).value <> 0 And (in_mlook.state And 1) <> 0) Then
        cmd_.sidemove = cmd_.sidemove + cvars(cv_m_side).value * mouse_x
    Else
        cl.viewangles.v(YAW) = cl.viewangles.v(YAW) - cvars(cv_m_yaw).value * mouse_x
    End If

    If in_mlook.state And 1 Then V_StopPitchDrift

    If (in_mlook.state And 1) <> 0 And (in_strafe.state And 1) = 0 Then
        cl.viewangles.v(PITCH) = cl.viewangles.v(PITCH) + cvars(cv_m_pitch).value * mouse_y
        If cl.viewangles.v(PITCH) > 80 Then cl.viewangles.v(PITCH) = 80
        If cl.viewangles.v(PITCH) < -70 Then cl.viewangles.v(PITCH) = -70
    Else
        If (in_strafe.state And 1) <> 0 And noclip_anglehack Then
            cmd_.upmove = cmd_.upmove - cvars(cv_m_forward).value * mouse_y
        Else
            cmd_.forwardmove = cmd_.forwardmove - cvars(cv_m_forward).value * mouse_y
        End If
    End If

    ' if the mouse has moved, force it to the center, so there's room to move
    If mx <> 0 Or my <> 0 Then SetCursorPos window_center_x, window_center_y
End Sub

Public Sub IN_Move(cmd_ As usercmd_t)
    If VID_IsActive() And Not VID_IsMinimized() Then
        IN_MouseMove cmd_
        IN_JoyMove cmd_
    End If
End Sub

Public Sub IN_Accumulate()
    If mouseactive Then
        GetCursorPos current_pos

        mx_accum = mx_accum + current_pos.x - window_center_x
        my_accum = my_accum + current_pos.y - window_center_y

        ' force the mouse to the center, so there's room to move
        SetCursorPos window_center_x, window_center_y
    End If
End Sub

Public Sub IN_ClearStates()
    If mouseactive Then
        mx_accum = 0
        my_accum = 0
        mouse_oldbuttonstate = 0
    End If
End Sub

Private Sub IN_StartupJoystick()
    Dim numdevs As Long, jc As JOYCAPS, mmr As Long, blank As JOYINFOEX

    ' assume no joystick
    joy_avail = False

    ' abort startup if user requests no joystick
    If COM_CheckParm("-nojoy") Then Exit Sub
    If sys_hidden Then Exit Sub          ' (QuakeVB)

    ' verify joystick driver is present
    numdevs = joyGetNumDevs()
    If numdevs = 0 Then
        Con_Printf "\njoystick not found -- driver not present\n\n"
        Exit Sub
    End If

    ' cycle through the joystick ids for the first valid one
    For joy_id = 0 To numdevs - 1
        ji = blank
        ji.dwSize = Len(ji)
        ji.dwFlags = JOY_RETURNCENTERED

        mmr = joyGetPosEx(joy_id, ji)
        If mmr = JOYERR_NOERROR Then Exit For
    Next

    ' abort startup if we didn't find a valid joystick
    If mmr <> JOYERR_NOERROR Then
        Con_Printf "\njoystick not found -- no valid joysticks (%x)\n\n", mmr
        Exit Sub
    End If

    ' get the capabilities of the selected joystick
    ' abort startup if command fails
    mmr = joyGetDevCaps(joy_id, jc, Len(jc))
    If mmr <> JOYERR_NOERROR Then
        Con_Printf "\njoystick not found -- invalid joystick capabilities (%x)\n\n", mmr
        Exit Sub
    End If

    ' save the joystick's number of buttons and POV status
    joy_numbuttons = jc.wNumButtons
    joy_haspov = (jc.wCaps And JOYCAPS_HASPOV) <> 0

    ' old button and POV states default to no buttons pressed
    joy_oldbuttonstate = 0
    joy_oldpovstate = 0

    ' mark the joystick as available and advanced initialization not completed
    ' this is needed as cvars are not available during initialization

    joy_avail = True
    joy_advancedinit = False

    Con_Printf "\njoystick detected\n\n"
End Sub

' (QuakeVB) C: *pdwRawValue[i], the ji field of the axis
Private Function RawValue(ByVal axis As Long) As Long
    Select Case axis
    Case JOY_AXIS_X: RawValue = ji.dwXpos
    Case JOY_AXIS_Y: RawValue = ji.dwYpos
    Case JOY_AXIS_Z: RawValue = ji.dwZpos
    Case JOY_AXIS_R: RawValue = ji.dwRpos
    Case JOY_AXIS_U: RawValue = ji.dwUpos
    Case JOY_AXIS_V: RawValue = ji.dwVpos
    End Select
End Function

Private Sub Joy_AdvancedUpdate_f()
    ' called once by IN_ReadJoystick and by user whenever an update is needed
    ' cvars are now available
    Dim i As Long, dwTemp As Long

    ' initialize all the maps
    For i = 0 To JOY_MAX_AXES - 1
        dwAxisMap(i) = AxisNada
        dwControlMap(i) = JOY_ABSOLUTE_AXIS
        pdwRawValue(i) = i               ' RawValuePointer(i)
    Next

    If cvars(cv_joy_advanced).value = 0 Then
        ' default joystick initialization
        ' 2 axes only with joystick control
        dwAxisMap(JOY_AXIS_X) = AxisTurn
        ' dwControlMap[JOY_AXIS_X] = JOY_ABSOLUTE_AXIS;
        dwAxisMap(JOY_AXIS_Y) = AxisForward
        ' dwControlMap[JOY_AXIS_Y] = JOY_ABSOLUTE_AXIS;
    Else
        If cvars(cv_joy_name).strval <> "joystick" Then
            ' notify user of advanced controller
            Con_Printf "\n%s configured\n\n", cvars(cv_joy_name).strval
        End If

        ' advanced initialization here
        ' data supplied by user via joy_axisn cvars
        dwTemp = ToInt(cvars(cv_joy_advaxisx).value)
        dwAxisMap(JOY_AXIS_X) = dwTemp And &HF
        dwControlMap(JOY_AXIS_X) = dwTemp And JOY_RELATIVE_AXIS
        dwTemp = ToInt(cvars(cv_joy_advaxisy).value)
        dwAxisMap(JOY_AXIS_Y) = dwTemp And &HF
        dwControlMap(JOY_AXIS_Y) = dwTemp And JOY_RELATIVE_AXIS
        dwTemp = ToInt(cvars(cv_joy_advaxisz).value)
        dwAxisMap(JOY_AXIS_Z) = dwTemp And &HF
        dwControlMap(JOY_AXIS_Z) = dwTemp And JOY_RELATIVE_AXIS
        dwTemp = ToInt(cvars(cv_joy_advaxisr).value)
        dwAxisMap(JOY_AXIS_R) = dwTemp And &HF
        dwControlMap(JOY_AXIS_R) = dwTemp And JOY_RELATIVE_AXIS
        dwTemp = ToInt(cvars(cv_joy_advaxisu).value)
        dwAxisMap(JOY_AXIS_U) = dwTemp And &HF
        dwControlMap(JOY_AXIS_U) = dwTemp And JOY_RELATIVE_AXIS
        dwTemp = ToInt(cvars(cv_joy_advaxisv).value)
        dwAxisMap(JOY_AXIS_V) = dwTemp And &HF
        dwControlMap(JOY_AXIS_V) = dwTemp And JOY_RELATIVE_AXIS
    End If

    ' compute the axes to collect from DirectInput
    joy_flags = JOY_RETURNCENTERED Or JOY_RETURNBUTTONS Or JOY_RETURNPOV
    For i = 0 To JOY_MAX_AXES - 1
        If dwAxisMap(i) <> AxisNada Then joy_flags = joy_flags Or dwAxisFlags(i)
    Next
End Sub

Public Sub IN_Commands()
    Dim i As Long, key_index As Long, buttonstate As Long, povstate As Long

    If Not joy_avail Then Exit Sub

    ' loop through the joystick buttons
    ' key a joystick event or auxillary event for higher number buttons for each state change
    buttonstate = ji.dwButtons
    For i = 0 To joy_numbuttons - 1
        If (buttonstate And BIT_(i)) <> 0 And (joy_oldbuttonstate And BIT_(i)) = 0 Then
            If i < 4 Then key_index = K_JOY1 Else key_index = K_AUX1
            Key_Event key_index + i, True
        End If

        If (buttonstate And BIT_(i)) = 0 And (joy_oldbuttonstate And BIT_(i)) <> 0 Then
            If i < 4 Then key_index = K_JOY1 Else key_index = K_AUX1
            Key_Event key_index + i, False
        End If
    Next
    joy_oldbuttonstate = buttonstate

    If joy_haspov Then
        ' convert POV information into 4 bits of state information
        ' this avoids any potential problems related to moving from one
        ' direction to another without going through the center position
        povstate = 0
        If ji.dwPOV <> JOY_POVCENTERED Then
            If ji.dwPOV = JOY_POVFORWARD Then povstate = povstate Or &H1
            If ji.dwPOV = JOY_POVRIGHT Then povstate = povstate Or &H2
            If ji.dwPOV = JOY_POVBACKWARD Then povstate = povstate Or &H4
            If ji.dwPOV = JOY_POVLEFT Then povstate = povstate Or &H8
        End If
        ' determine which bits have changed and key an auxillary event for each change
        For i = 0 To 3
            If (povstate And BIT_(i)) <> 0 And (joy_oldpovstate And BIT_(i)) = 0 Then Key_Event K_AUX1 + 28 + i, True   ' K_AUX29

            If (povstate And BIT_(i)) = 0 And (joy_oldpovstate And BIT_(i)) <> 0 Then Key_Event K_AUX1 + 28 + i, False
        Next
        joy_oldpovstate = povstate
    End If
End Sub

Private Function IN_ReadJoystick() As Boolean
    Dim blank As JOYINFOEX

    ji = blank
    ji.dwSize = Len(ji)
    ji.dwFlags = joy_flags

    If joyGetPosEx(joy_id, ji) = JOYERR_NOERROR Then
        ' this is a hack -- there is a bug in the Logitech WingMan Warrior DirectInput Driver
        ' rather than having 32768 be the zero point, they have the zero point at 32668
        ' go figure -- anyway, now we get the full resolution out of the device
        If cvars(cv_joy_wwhack1).value <> 0 Then ji.dwUpos = ji.dwUpos + 100
        IN_ReadJoystick = True
    Else
        ' read error occurred
        ' turning off the joystick seems too harsh for 1 read error,\
        ' but what should be done?
        ' Con_Printf ("IN_ReadJoystick: no response\n");
        ' joy_avail = false;
        IN_ReadJoystick = False
    End If
End Function

Private Sub IN_JoyMove(cmd_ As usercmd_t)
    Dim speed As Single, aspeed As Single, fAxisValue As Single, fTemp As Single, i As Long

    ' complete initialization if first time in
    ' this is needed as cvars are not available at initialization time
    If Not joy_advancedinit Then
        Joy_AdvancedUpdate_f
        joy_advancedinit = True
    End If

    ' verify joystick is available and that the user wants to use it
    If Not joy_avail Or cvars(cv_in_joystick).value = 0 Then Exit Sub

    ' collect the joystick data, if possible
    If Not IN_ReadJoystick() Then Exit Sub

    If in_speed.state And 1 Then
        speed = cvars(cv_cl_movespeedkey).value
    Else
        speed = 1
    End If
    aspeed = speed * host_frametime

    ' loop through the axes
    For i = 0 To JOY_MAX_AXES - 1
        ' get the floating point zero-centered, potentially-inverted data for the current axis
        fAxisValue = RawValue(pdwRawValue(i))
        ' move centerpoint to zero
        fAxisValue = fAxisValue - 32768#

        If cvars(cv_joy_wwhack2).value <> 0 Then
            If dwAxisMap(i) = AxisTurn Then
                ' this is a special formula for the Logitech WingMan Warrior
                ' y=ax^b; where a = 300 and b = 1.3
                ' also x values are in increments of 800 (so this is factored out)
                ' then bounds check result to level out excessively high spin rates
                ' (C's abs is the integer one)
                fTemp = 300# * (Abs(ToInt(fAxisValue)) / 800#) ^ 1.3
                If fTemp > 14000# Then fTemp = 14000#
                ' restore direction information
                If fAxisValue > 0 Then fAxisValue = fTemp Else fAxisValue = -fTemp
            End If
        End If

        ' convert range from -32768..32767 to -1..1
        fAxisValue = fAxisValue / 32768#

        Select Case dwAxisMap(i)
        Case AxisForward
            If cvars(cv_joy_advanced).value = 0 And (in_mlook.state And 1) <> 0 Then
                ' user wants forward control to become look control
                If Abs(fAxisValue) > cvars(cv_joy_pitchthreshold).value Then
                    ' if mouse invert is on, invert the joystick pitch value
                    ' only absolute control support here (joy_advanced is false)
                    If cvars(cv_m_pitch).value < 0 Then
                        cl.viewangles.v(PITCH) = cl.viewangles.v(PITCH) - (fAxisValue * cvars(cv_joy_pitchsensitivity).value) * aspeed * cvars(cv_cl_pitchspeed).value
                    Else
                        cl.viewangles.v(PITCH) = cl.viewangles.v(PITCH) + (fAxisValue * cvars(cv_joy_pitchsensitivity).value) * aspeed * cvars(cv_cl_pitchspeed).value
                    End If
                    V_StopPitchDrift
                Else
                    ' no pitch movement
                    ' disable pitch return-to-center unless requested by user
                    ' *** this code can be removed when the lookspring bug is fixed
                    ' *** the bug always has the lookspring feature on
                    If cvars(cv_lookspring).value = 0 Then V_StopPitchDrift
                End If
            Else
                ' user wants forward control to be forward control
                If Abs(fAxisValue) > cvars(cv_joy_forwardthreshold).value Then
                    cmd_.forwardmove = cmd_.forwardmove + (fAxisValue * cvars(cv_joy_forwardsensitivity).value) * speed * cvars(cv_cl_forwardspeed).value
                End If
            End If

        Case AxisSide
            If Abs(fAxisValue) > cvars(cv_joy_sidethreshold).value Then
                cmd_.sidemove = cmd_.sidemove + (fAxisValue * cvars(cv_joy_sidesensitivity).value) * speed * cvars(cv_cl_sidespeed).value
            End If

        Case AxisTurn
            If (in_strafe.state And 1) <> 0 Or (cvars(cv_lookstrafe).value <> 0 And (in_mlook.state And 1) <> 0) Then
                ' user wants turn control to become side control
                If Abs(fAxisValue) > cvars(cv_joy_sidethreshold).value Then
                    cmd_.sidemove = cmd_.sidemove - (fAxisValue * cvars(cv_joy_sidesensitivity).value) * speed * cvars(cv_cl_sidespeed).value
                End If
            Else
                ' user wants turn control to be turn control
                If Abs(fAxisValue) > cvars(cv_joy_yawthreshold).value Then
                    If dwControlMap(i) = JOY_ABSOLUTE_AXIS Then
                        cl.viewangles.v(YAW) = cl.viewangles.v(YAW) + (fAxisValue * cvars(cv_joy_yawsensitivity).value) * aspeed * cvars(cv_cl_yawspeed).value
                    Else
                        cl.viewangles.v(YAW) = cl.viewangles.v(YAW) + (fAxisValue * cvars(cv_joy_yawsensitivity).value) * speed * 180#
                    End If
                End If
            End If

        Case AxisLook
            If in_mlook.state And 1 Then
                If Abs(fAxisValue) > cvars(cv_joy_pitchthreshold).value Then
                    ' pitch movement detected and pitch movement desired by user
                    If dwControlMap(i) = JOY_ABSOLUTE_AXIS Then
                        cl.viewangles.v(PITCH) = cl.viewangles.v(PITCH) + (fAxisValue * cvars(cv_joy_pitchsensitivity).value) * aspeed * cvars(cv_cl_pitchspeed).value
                    Else
                        cl.viewangles.v(PITCH) = cl.viewangles.v(PITCH) + (fAxisValue * cvars(cv_joy_pitchsensitivity).value) * speed * 180#
                    End If
                    V_StopPitchDrift
                Else
                    ' no pitch movement
                    ' disable pitch return-to-center unless requested by user
                    ' *** this code can be removed when the lookspring bug is fixed
                    ' *** the bug always has the lookspring feature on
                    If cvars(cv_lookspring).value = 0 Then V_StopPitchDrift
                End If
            End If

        Case Else
        End Select
    Next

    ' bounds check pitch
    If cl.viewangles.v(PITCH) > 80# Then cl.viewangles.v(PITCH) = 80#
    If cl.viewangles.v(PITCH) < -70# Then cl.viewangles.v(PITCH) = -70#
End Sub

' (QuakeVB) the console commands of this module, for Cmd_Call
Public Sub IN_CallCommand(ByVal fn As Long)
    Select Case fn
    Case XC_Force_CenterView_f: Force_CenterView_f
    Case XC_Joy_AdvancedUpdate_f: Joy_AdvancedUpdate_f
    Case Else: Sys_Error "IN_CallCommand: bad function %i", fn
    End Select
End Sub
