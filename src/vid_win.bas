Attribute VB_Name = "vid_win"
' vid_win.bas - Win32 video (replaces vid_win.c, which used MGL, DirectDraw
' and VESA modes): the 8-bit view buffer vid_buffer() is copied into a
' palettized DIB section and stretched onto frmQuake.
'
' The internal resolution is -width x -height (default 320 x 240, WinQuake's
' windowed mode 0) and the window is -scale times that (default 2).
'
' Keyboard, mouse wheel and mouse button messages are queued by the window
' procedure and handed to Key_Event / IN_MouseEvent by Sys_SendKeyEvents
' (VID_ProcessEvents), so no engine code runs inside a Windows callback,
' where a VB error could not unwind. C calls Key_Event from MainWndProc,
' which runs inside Sys_SendKeyEvents too, so the timing is the same.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO256
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As Long
End Type

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO256, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal hdc As Long, ByVal un1 As Long, ByVal un2 As Long, pcRGBQuad As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function AdjustWindowRect Lib "user32" (lpRect As RECT, ByVal dwStyle As Long, ByVal bMenu As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Private Const GWL_WNDPROC As Long = -4
Private Const GWL_STYLE As Long = -16
Private Const WM_MOVE As Long = &H3
Private Const WM_SIZE As Long = &H5
Private Const WM_ACTIVATE As Long = &H6
Private Const WM_CLOSE As Long = &H10
Private Const WM_ERASEBKGND As Long = &H14
Private Const MM_MCINOTIFY As Long = &H3B9
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_CHAR As Long = &H102
Private Const WM_SYSKEYDOWN As Long = &H104
Private Const WM_SYSKEYUP As Long = &H105
Private Const WM_SYSCHAR As Long = &H106
Private Const WM_SYSCOMMAND As Long = &H112
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const SC_KEYMENU As Long = &HF100&
Private Const WA_INACTIVE As Long = 0
Private Const SIZE_MINIMIZED As Long = 1
Private Const MK_LBUTTON As Long = 1
Private Const MK_RBUTTON As Long = 2
Private Const MK_MBUTTON As Long = &H10
Private Const VK_F4 As Long = &H73
Private Const SWP_NOZORDER As Long = 4
Private Const COLORONCOLOR As Long = 3
Private Const SRCCOPY As Long = &HCC0020

' r_shared.h

Public mainwindow As Long

' the client area on the screen (C: window_x etc. and window_rect)
Public window_center_x As Long, window_center_y As Long
Public window_x As Long, window_y As Long, window_width As Long, window_height As Long
Private windowed_mouse As Long
Private oldWndProc As Long
Private memDC As Long
Private hDib As Long
Private hOldBmp As Long
Private dibBits As Long
Private bmi As BITMAPINFO256
Private palQuads(0 To 255) As Long
Private vid_initialized As Boolean
Private windowHidden As Boolean

Private ActiveApp As Boolean
Private sound_active As Boolean          ' C: static of AppActivate
Private Minimized As Boolean
Private quit_requested As Boolean
Private states_to_clear As Boolean

' events posted by the window procedure: a key (down or up) or a mouse state
Private Const EVQ_SIZE As Long = 256
Private Const EV_KEY As Long = 1
Private Const EV_MOUSE As Long = 2
Private evq_type(0 To EVQ_SIZE - 1) As Long
Private evq_data(0 To EVQ_SIZE - 1) As Long
Private evq_down(0 To EVQ_SIZE - 1) As Boolean
Private evq_head As Long
Private evq_tail As Long

Private scantokey(0 To 127) As Long

Public cv_vid_mode As Long
Public cv_windowed_mouse As Long     ' C: _windowed_mouse

' (QuakeVB) -key <frame>[:<n>] <key>: scripted key presses for test runs. The
' key goes down at host_framecount <frame> and up n frames later (default 1).
' <key> is a key name as bind takes it (a lowercase letter, ENTER, SPACE,
' UPARROW, ...).
Private Type scriptkey_t
    frame As Long
    holdframes As Long
    key As Long
    pressed As Boolean
    released As Boolean
End Type
Private scriptkeys() As scriptkey_t
Private num_scriptkeys As Long

'==============================================================================

Private Sub InitScanToKey()
    Dim t As Variant, i As Long
    ' 0x00 .. 0x7f, as vid_win.c's scantokey
    t = Array( _
        0, 27, 49, 50, 51, 52, 53, 54, 55, 56, 57, 48, 45, 61, K_BACKSPACE, 9, _
        113, 119, 101, 114, 116, 121, 117, 105, 111, 112, 91, 93, 13, K_CTRL, 97, 115, _
        100, 102, 103, 104, 106, 107, 108, 59, 39, 96, K_SHIFT, 92, 122, 120, 99, 118, _
        98, 110, 109, 44, 46, 47, K_SHIFT, 42, K_ALT, 32, 0, K_F1, K_F2, K_F3, K_F4, K_F5, _
        K_F6, K_F7, K_F8, K_F9, K_F10, K_PAUSE, 0, K_HOME, K_UPARROW, K_PGUP, 45, K_LEFTARROW, 53, K_RIGHTARROW, 43, K_END, _
        K_DOWNARROW, K_PGDN, K_INS, K_DEL, 0, 0, 0, K_F11, K_F12, 0, 0, 0, 0, 0, 0, 0)
    For i = 0 To 127
        If i <= UBound(t) Then scantokey(i) = t(i) Else scantokey(i) = 0
    Next
End Sub

' Map from windows to quake keynums
Private Function MapKey(ByVal key As Long) As Long
    key = (key And &HFF0000) \ &H10000
    If key > 127 Then
        MapKey = 0
        Exit Function
    End If
    MapKey = scantokey(key)
End Function

Private Sub PostEvent(ByVal typ As Long, ByVal data As Long, ByVal down As Boolean)
    Dim nxt As Long
    nxt = (evq_head + 1) And (EVQ_SIZE - 1)
    If nxt = evq_tail Then Exit Sub      ' full: drop it
    evq_type(evq_head) = typ
    evq_data(evq_head) = data
    evq_down(evq_head) = down
    evq_head = nxt
End Sub

Private Sub AddScriptKey(ByVal spec As String, ByVal keyname As String)
    Dim p As Long, k As scriptkey_t
    p = InStr(spec, ":")
    If p > 0 Then
        k.frame = Q_atoi(Left$(spec, p - 1))
        k.holdframes = Q_atoi(Mid$(spec, p + 1))
    Else
        k.frame = Q_atoi(spec)
        k.holdframes = 1
    End If
    If k.holdframes < 1 Then k.holdframes = 1
    k.key = Key_StringToKeynum(keyname)
    If k.key = -1 Then Sys_Error "-key: unknown key %s", keyname
    num_scriptkeys = num_scriptkeys + 1
    ReDim Preserve scriptkeys(1 To num_scriptkeys)
    scriptkeys(num_scriptkeys) = k
End Sub

' -key <frame>[:<n>] <key>, and -keyfile <file> with one "<frame>[:<n>] <key>"
' per line (the command line keeps only MAX_NUM_ARGVS arguments)
Private Sub ParseScriptKeys()
    Dim i As Long, f As Integer, ln As String, p As Long
    For i = 1 To com_argc - 3
        If com_argv(i) = "-key" Then AddScriptKey com_argv(i + 1), com_argv(i + 2)
    Next
    i = COM_CheckParm("-keyfile")
    If i > 0 And i < com_argc - 1 Then
        If Sys_FileTime(com_argv(i + 1)) = -1 Then Sys_Error "-keyfile: can't open %s", com_argv(i + 1)
        f = FreeFile
        Open com_argv(i + 1) For Input As #f
        Do While Not EOF(f)
            Line Input #f, ln
            ln = Trim$(ln)
            p = InStr(ln, " ")
            If p > 0 Then AddScriptKey Left$(ln, p - 1), Trim$(Mid$(ln, p + 1))
        Loop
        Close #f
    End If
End Sub

Private Function RunScriptKeys() As Boolean
    Dim i As Long
    For i = 1 To num_scriptkeys
        If Not scriptkeys(i).pressed And host_framecount >= scriptkeys(i).frame Then
            scriptkeys(i).pressed = True
            Key_Event scriptkeys(i).key, True
            RunScriptKeys = True
        End If
        If scriptkeys(i).pressed And Not scriptkeys(i).released Then
            If host_framecount >= scriptkeys(i).frame + scriptkeys(i).holdframes Then
                scriptkeys(i).released = True
                Key_Event scriptkeys(i).key, False
                RunScriptKeys = True
            End If
        End If
    Next
End Function

' Sends the queued window events to the engine; True if there were any.
Public Function VID_ProcessEvents() As Boolean
    Dim i As Long

    ' (C: AppActivate) enable/disable sound on focus gain/loss. Checked on
    ' every call, so a -hidden run, which counts as active, unblocks it at once.
    If Not VID_IsActive() And sound_active Then
        S_BlockSound
        S_ClearBuffer
        sound_active = False
    ElseIf VID_IsActive() And Not sound_active Then
        S_UnblockSound
        S_ClearBuffer
        sound_active = True
    End If

    If num_scriptkeys Then
        If RunScriptKeys() Then VID_ProcessEvents = True
    End If

    If quit_requested Then
        quit_requested = False
        Sys_Quit
    End If

    If states_to_clear Then
        ' fix the leftover Alt from any Alt-Tab or the like that switched us away
        ' (C: ClearAllStates: send an up event for each key, to make sure the
        ' server clears them all)
        states_to_clear = False
        For i = 0 To 255
            Key_Event i, False
        Next
        Key_ClearStates
        IN_ClearStates
        VID_ProcessEvents = True

        ' (C: AppActivate) mouse-capture normal windows on demand
        If cvars(cv_windowed_mouse).value <> 0 And Not windowHidden Then
            If ActiveApp Then
                IN_ActivateMouse
                IN_HideMouse
            Else
                IN_DeactivateMouse
                IN_ShowMouse
            End If
        End If
    End If

    Do While evq_tail <> evq_head
        VID_ProcessEvents = True
        Select Case evq_type(evq_tail)
        Case EV_KEY
            Key_Event evq_data(evq_tail), evq_down(evq_tail)
        Case EV_MOUSE
            IN_MouseEvent evq_data(evq_tail)
        End Select
        evq_tail = (evq_tail + 1) And (EVQ_SIZE - 1)
    Loop
End Function

Public Function VID_IsActive() As Boolean
    VID_IsActive = ActiveApp Or windowHidden
End Function

Public Function VID_IsMinimized() As Boolean
    VID_IsMinimized = Minimized And Not windowHidden
End Function

' Form_QueryUnload: the window's close box
Public Sub VID_RequestQuit()
    quit_requested = True
End Sub

Public Function QuakeWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim temp As Long

    Select Case uMsg
    Case WM_KEYDOWN, WM_SYSKEYDOWN
        ' (QuakeVB) Alt+F4 quits
        If uMsg = WM_SYSKEYDOWN And wParam = VK_F4 Then
            quit_requested = True
            Exit Function
        End If
        ' a -hidden run takes its input from the command line only: the
        ' invisible window can still own the keyboard focus
        If Not windowHidden Then PostEvent EV_KEY, MapKey(lParam), True
        Exit Function

    Case WM_KEYUP, WM_SYSKEYUP
        If Not windowHidden Then PostEvent EV_KEY, MapKey(lParam), False
        Exit Function

    Case WM_CHAR, WM_SYSCHAR
        ' keep Alt-Space from happening
        Exit Function

    Case WM_SYSCOMMAND
        If (wParam And &HFFF0&) = SC_KEYMENU Then Exit Function

    Case WM_SIZE
        Minimized = (wParam = SIZE_MINIMIZED)
        VID_UpdateWindowStatus

    Case WM_MOVE
        VID_UpdateWindowStatus

    Case WM_ACTIVATE
        ActiveApp = ((wParam And &HFFFF&) <> WA_INACTIVE)
        Minimized = ((wParam And &HFFFF0000) <> 0)
        states_to_clear = True

    ' this is complicated because Win32 seems to pack multiple mouse events into
    ' one update sometimes, so we always check all states and look for events
    Case WM_LBUTTONDOWN, WM_LBUTTONUP, WM_RBUTTONDOWN, WM_RBUTTONUP, WM_MBUTTONDOWN, WM_MBUTTONUP, WM_MOUSEMOVE
        If Not windowHidden Then
            temp = 0
            If wParam And MK_LBUTTON Then temp = temp Or 1
            If wParam And MK_RBUTTON Then temp = temp Or 2
            If wParam And MK_MBUTTON Then temp = temp Or 4
            PostEvent EV_MOUSE, temp, False
        End If

    ' JACK: This is the mouse wheel with the Intellimouse
    ' Its delta is either positive or neg, and we generate the proper
    ' Event.
    Case WM_MOUSEWHEEL
        If Not windowHidden Then
            If (wParam And &H80000000) = 0 Then
                PostEvent EV_KEY, K_MWHEELUP, True
                PostEvent EV_KEY, K_MWHEELUP, False
            Else
                PostEvent EV_KEY, K_MWHEELDOWN, True
                PostEvent EV_KEY, K_MWHEELDOWN, False
            End If
        End If
        Exit Function

    Case WM_ERASEBKGND
        QuakeWndProc = 1
        Exit Function

    Case WM_CLOSE
        quit_requested = True
        Exit Function

    Case MM_MCINOTIFY
        ' (C: CDAudio_MessageHandler; queued, CDAudio_Update handles it)
        CDAudio_QueueNotify wParam, lParam
        Exit Function
    End Select

    QuakeWndProc = CallWindowProc(oldWndProc, hWnd, uMsg, wParam, lParam)
End Function

'==============================================================================

Public Sub VID_Init(palette() As Byte)
    Dim w As Long, h As Long, p As Long, scale_ As Long
    Dim r As RECT, hdc As Long

    InitScanToKey
    ParseScriptKeys

    cv_vid_mode = Cvar_RegisterVariable("vid_mode", "0")
    ' (QuakeVB) on by default (C: "0"): the game is always a window, so
    ' without it the mouse would do nothing
    cv_windowed_mouse = Cvar_RegisterVariable("_windowed_mouse", "1", True)

    ' internal resolution: -width (default 320) and -height (default 3/4 of
    ' the width, 240 for 320); the width is a multiple of 8
    w = 320
    p = COM_CheckParm("-width")
    If p > 0 And p < com_argc - 1 Then w = Q_atoi(com_argv(p + 1))
    If w < 320 Then w = 320
    If w > MAXWIDTH Then w = MAXWIDTH
    w = w And Not 7
    h = (w * 3) \ 4
    p = COM_CheckParm("-height")
    If p > 0 And p < com_argc - 1 Then h = Q_atoi(com_argv(p + 1))
    If h < 200 Then h = 200
    If h > MAXHEIGHT Then h = MAXHEIGHT

    ' window size: -scale times the resolution (default 2 up to 400 pixels wide)
    If w <= 400 Then scale_ = 2 Else scale_ = 1
    p = COM_CheckParm("-scale")
    If p > 0 And p < com_argc - 1 Then scale_ = Q_atoi(com_argv(p + 1))
    If scale_ < 1 Then scale_ = 1

    vid.width = w
    vid.height = h
    vid.conwidth = w
    vid.conheight = h
    vid.rowbytes = w                     ' DIB rows are 4-byte aligned: w is a multiple of 8
    vid.conrowbytes = w
    vid.numpages = 1
    vid.maxwarpwidth = WARP_WIDTH
    vid.maxwarpheight = WARP_HEIGHT
    vid.aspect = (CSng(h) / CSng(w)) * (320! / 240!)
    vid_colormap = host_colormap
    vid.fullbright = 256 - GetLong(host_colormap, 2048 * 4)
    vid.recalc_refdef = 1
    ReDim vid_buffer(0 To vid.rowbytes * vid.height - 1)
    vid_bufferofs = 0

    ' (C: VID_AllocBuffers and D_InitCaches)
    ReDim d_pzbuffer(0 To vid.width * vid.height - 1)
    D_InitCaches D_SurfaceCacheForRes(vid.width, vid.height)

    windowHidden = sys_hidden

    Load frmQuake
    mainwindow = frmQuake.hWnd

    r.Left = 0: r.Top = 0: r.Right = w * scale_: r.Bottom = h * scale_
    AdjustWindowRect r, GetWindowLong(mainwindow, GWL_STYLE), 0
    SetWindowPos mainwindow, 0, (GetSystemMetrics(0) - (r.Right - r.Left)) \ 2, _
        (GetSystemMetrics(1) - (r.Bottom - r.Top)) \ 2, r.Right - r.Left, r.Bottom - r.Top, SWP_NOZORDER

    ' palettized top-down DIB section
    bmi.bmiHeader.biSize = Len(bmi.bmiHeader)
    bmi.bmiHeader.biWidth = w
    bmi.bmiHeader.biHeight = -h
    bmi.bmiHeader.biPlanes = 1
    bmi.bmiHeader.biBitCount = 8
    bmi.bmiHeader.biCompression = 0
    bmi.bmiHeader.biClrUsed = 256
    bmi.bmiHeader.biClrImportant = 256

    hdc = GetDC(mainwindow)
    memDC = CreateCompatibleDC(hdc)
    hDib = CreateDIBSection(memDC, bmi, 0, dibBits, 0, 0)
    ReleaseDC mainwindow, hdc
    If hDib = 0 Or dibBits = 0 Then Sys_Error "VID_Init: CreateDIBSection failed"
    hOldBmp = SelectObject(memDC, hDib)

    oldWndProc = SetWindowLong(mainwindow, GWL_WNDPROC, AddressOf QuakeWndProc)
    VID_UpdateWindowStatus
    vid_initialized = True

    VID_SetPalette palette

    ' sound initialization has to go here, so there's a window for the sound
    ' code to work with (as vid_win.c does)
    S_Init

    ActiveApp = True
    If Not windowHidden Then
        frmQuake.Show
        frmQuake.SetFocus
    End If
    DoEvents
End Sub

Public Sub VID_Shutdown()
    If Not vid_initialized Then Exit Sub
    vid_initialized = False
    If oldWndProc <> 0 Then SetWindowLong mainwindow, GWL_WNDPROC, oldWndProc
    oldWndProc = 0
    If memDC <> 0 Then
        SelectObject memDC, hOldBmp
        DeleteObject hDib
        DeleteDC memDC
    End If
    memDC = 0
    Unload frmQuake
End Sub

Public Sub VID_SetPalette(palette() As Byte)
    Dim i As Long
    For i = 0 To 255
        palQuads(i) = palette(i * 3 + 2) Or (palette(i * 3 + 1) * 256&) Or (palette(i * 3) * 65536)
    Next
    If memDC <> 0 Then SetDIBColorTable memDC, 0, 256, palQuads(0)
End Sub

' called for bonus and pain flashes, and for underwater color changes
Public Sub VID_ShiftPalette(palette() As Byte)
    VID_SetPalette palette
End Sub

' flushes the given rectangles from the view buffer to the screen
' (the whole buffer is copied and shown: it costs nothing worth saving)
Public Sub VID_Update(rects As vrect_t)
    If Not vid_initialized Then Exit Sub
    CopyMemory ByVal dibBits, vid_buffer(0), vid.rowbytes * vid.height
    If windowHidden Then Exit Sub
    BlitToWindow

    ' handle the mouse state when windowed if that's changed
    If ToInt(cvars(cv_windowed_mouse).value) <> windowed_mouse Then
        If cvars(cv_windowed_mouse).value <> 0 Then
            IN_ActivateMouse
            IN_HideMouse
        Else
            IN_DeactivateMouse
            IN_ShowMouse
        End If

        windowed_mouse = ToInt(cvars(cv_windowed_mouse).value)
    End If
End Sub

' the client area's place on the screen, for the mouse
Public Sub VID_UpdateWindowStatus()
    Dim r As RECT, pt As POINTAPI

    If mainwindow = 0 Then Exit Sub
    GetClientRect mainwindow, r
    ClientToScreen mainwindow, pt
    window_x = pt.x
    window_y = pt.y
    window_width = r.Right
    window_height = r.Bottom
    window_center_x = window_x + window_width \ 2
    window_center_y = window_y + window_height \ 2

    IN_UpdateClipCursor
End Sub

Private Sub BlitToWindow()
    Dim hdc As Long, r As RECT
    GetClientRect mainwindow, r
    hdc = GetDC(mainwindow)
    SetStretchBltMode hdc, COLORONCOLOR
    StretchBlt hdc, 0, 0, r.Right, r.Bottom, memDC, 0, 0, vid.width, vid.height, SRCCOPY
    ReleaseDC mainwindow, hdc
End Sub

' Form_Paint and Form_Resize: show the last frame again
Public Sub VID_RepaintWindow()
    If vid_initialized And Not windowHidden Then BlitToWindow
End Sub

' called only on Win32, when pause happens, so the mouse can be released
Public Sub VID_HandlePause(ByVal pause As Boolean)
    If cvars(cv_windowed_mouse).value <> 0 And Not windowHidden Then
        If pause Then
            IN_DeactivateMouse
            IN_ShowMouse
        Else
            IN_ActivateMouse
            IN_HideMouse
        End If
    End If
End Sub

Public Sub VID_LockBuffer()
End Sub

Public Sub VID_UnlockBuffer()
End Sub

' the loading disc is drawn straight to the screen in C; not shown here yet
Public Sub D_BeginDirectRect(ByVal x As Long, ByVal y As Long, ByVal pbitmap As Long, ByVal width As Long, ByVal height As Long)
End Sub

Public Sub D_EndDirectRect(ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long)
End Sub
