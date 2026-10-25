Attribute VB_Name = "console"
' console.bas - the scrollback console, notify lines and the print functions
' (console.c).
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public con_linewidth As Long
Private Const con_cursorspeed As Single = 4

Private Const CON_TEXTSIZE As Long = 16384

Public con_current As Long           ' where next message will be printed
Public con_x As Long                 ' offset in current line for next print
Private con_text() As Byte

Public cv_con_notifytime As Long     ' seconds

Private Const NUM_CON_TIMES As Long = 4
Private con_times(0 To NUM_CON_TIMES - 1) As Single  ' realtime time the line was generated
                                                     ' for transparent notify lines
Public con_vislines As Long

Public con_debuglog_ As Boolean

Private con_cr As Boolean            ' C: static int cr in Con_Print
Private con_inupdate As Boolean      ' C: static qboolean inupdate in Con_Printf

Public Sub Con_ToggleConsole_f()
    If key_dest = key_console Then
        If cls.state = ca_connected Then
            key_dest = key_game
            key_lines(edit_line, 1) = 0  ' clear any typing
            key_linepos = 1
        Else
            M_Menu_Main_f
        End If
    Else
        key_dest = key_console
    End If

    SCR_EndLoadingPlaque
    Con_ClearNotify
End Sub

Public Sub Con_Clear_f()
    If con_initialized Then FillMemory con_text(0), CON_TEXTSIZE, 32
End Sub

Public Sub Con_ClearNotify()
    Dim i As Long
    For i = 0 To NUM_CON_TIMES - 1
        con_times(i) = 0
    Next
End Sub

Public Sub Con_MessageMode_f()
    key_dest = key_message
    team_message = False
End Sub

Public Sub Con_MessageMode2_f()
    key_dest = key_message
    team_message = True
End Sub

' If the line width has changed, reformat the buffer.
Public Sub Con_CheckResize()
    Dim i As Long, j As Long, width As Long, oldwidth As Long, oldtotallines As Long
    Dim numlines As Long, numchars As Long
    Dim tbuf(0 To CON_TEXTSIZE - 1) As Byte

    width = (vid.width \ 8) - 2

    If width = con_linewidth Then Exit Sub

    If width < 1 Then
        ' video hasn't been initialized yet
        width = 38
        con_linewidth = width
        con_totallines = CON_TEXTSIZE \ con_linewidth
        FillMemory con_text(0), CON_TEXTSIZE, 32
    Else
        oldwidth = con_linewidth
        con_linewidth = width
        oldtotallines = con_totallines
        con_totallines = CON_TEXTSIZE \ con_linewidth
        numlines = oldtotallines

        If con_totallines < numlines Then numlines = con_totallines

        numchars = oldwidth

        If con_linewidth < numchars Then numchars = con_linewidth

        CopyMemory tbuf(0), con_text(0), CON_TEXTSIZE
        FillMemory con_text(0), CON_TEXTSIZE, 32

        For i = 0 To numlines - 1
            For j = 0 To numchars - 1
                con_text((con_totallines - 1 - i) * con_linewidth + j) = _
                    tbuf(((con_current - i + oldtotallines) Mod oldtotallines) * oldwidth + j)
            Next
        Next

        Con_ClearNotify
    End If

    con_backscroll = 0
    con_current = con_totallines - 1
End Sub

Public Sub Con_Init()
    Dim temp As String

    con_debuglog_ = (COM_CheckParm("-condebug") <> 0)

    If con_debuglog_ Then
        temp = com_gamedir & "/qconsole.log"
        On Error Resume Next
        Kill temp
        On Error GoTo 0
    End If

    ReDim con_text(0 To CON_TEXTSIZE - 1)
    FillMemory con_text(0), CON_TEXTSIZE, 32
    con_linewidth = -1
    Con_CheckResize

    Con_Printf "Console initialized.\n"

    '
    ' register our commands
    '
    cv_con_notifytime = Cvar_RegisterVariable("con_notifytime", "3")

    Cmd_AddCommand "toggleconsole", XC_Con_ToggleConsole_f
    Cmd_AddCommand "messagemode", XC_Con_MessageMode_f
    Cmd_AddCommand "messagemode2", XC_Con_MessageMode2_f
    Cmd_AddCommand "clear", XC_Con_Clear_f
    con_initialized = True
End Sub

Private Sub Con_Linefeed()
    con_x = 0
    con_current = con_current + 1
    FillMemory con_text((con_current Mod con_totallines) * con_linewidth), con_linewidth, 32
End Sub

' Handles cursor positioning, line wrapping, etc
' All console printing must go through this in order to be logged to disk
' If no console is visible, the notify window will pop up.
Public Sub Con_Print(ByVal txt As String)
    Dim y As Long, c As Long, l As Long, mask As Long
    Dim t() As Byte, p As Long

    con_backscroll = 0

    StrToBytesZ txt, t
    p = 0

    If t(0) = 1 Then
        mask = 128                       ' go to colored text
        S_LocalSound "misc/talk.wav"
        ' play talk wav
        p = p + 1
    ElseIf t(0) = 2 Then
        mask = 128                       ' go to colored text
        p = p + 1
    Else
        mask = 0
    End If

    Do
        c = t(p)
        If c = 0 Then Exit Do

        ' count word length (C compares signed chars: 128..255 end a word too)
        For l = 0 To con_linewidth - 1
            If t(p + l) <= 32 Or t(p + l) >= 128 Then Exit For
        Next

        ' word wrap
        If l <> con_linewidth And (con_x + l > con_linewidth) Then con_x = 0

        p = p + 1

        If con_cr Then
            con_current = con_current - 1
            con_cr = False
        End If

        If con_x = 0 Then
            Con_Linefeed
            ' mark time for transparent overlay
            If con_current >= 0 Then con_times(con_current Mod NUM_CON_TIMES) = realtime
        End If

        Select Case c
        Case 10                          ' \n
            con_x = 0
        Case 13                          ' \r
            con_x = 0
            con_cr = True
        Case Else                        ' display character and advance
            y = con_current Mod con_totallines
            con_text(y * con_linewidth + con_x) = c Or mask
            con_x = con_x + 1
            If con_x >= con_linewidth Then con_x = 0
        End Select
    Loop
End Sub

Public Sub Con_DebugLog(ByVal file As String, ByVal fmt As String, ParamArray args() As Variant)
    Dim a As Variant, f As Integer
    a = args
    On Error Resume Next
    f = FreeFile
    Open file For Append As #f
    Print #f, Replace(vsprintf(fmt, a), vbLf, vbCrLf);
    Close #f
End Sub

' Handles cursor positioning, line wrapping, etc
Public Sub Con_Printf(ByVal fmt As String, ParamArray args() As Variant)
    Dim a As Variant, msg As String
    a = args
    msg = vsprintf(fmt, a)

    ' also echo to debugging console
    Sys_Printf "%s", msg

    ' log all messages to file
    If con_debuglog_ Then Con_DebugLog com_gamedir & "/qconsole.log", "%s", msg

    If Not con_initialized Then Exit Sub

    If cls.state = ca_dedicated Then Exit Sub    ' no graphics mode

    ' write it to the scrollable buffer
    Con_Print msg

    ' update the screen if the console is displayed
    If cls.signon <> SIGNONS And Not scr_disabled_for_loading Then
        ' protect against infinite loop if something in SCR_UpdateScreen calls
        ' Con_Printd
        If Not con_inupdate Then
            con_inupdate = True
            SCR_UpdateScreen
            con_inupdate = False
        End If
    End If
End Sub

' A Con_Printf that only shows up if the "developer" cvar is set
Public Sub Con_DPrintf(ByVal fmt As String, ParamArray args() As Variant)
    Dim a As Variant

    If cvars(cv_developer).value = 0 Then Exit Sub   ' don't confuse non-developers with techie stuff...

    a = args
    Con_Printf "%s", vsprintf(fmt, a)
End Sub

' Okay to call even when the screen can't be updated
Public Sub Con_SafePrintf(ByVal fmt As String, ParamArray args() As Variant)
    Dim a As Variant, temp As Boolean
    a = args

    temp = scr_disabled_for_loading
    scr_disabled_for_loading = True
    Con_Printf "%s", vsprintf(fmt, a)
    scr_disabled_for_loading = temp
End Sub

'==============================================================================
' DRAWING

' The input line scrolls horizontally if typing goes beyond the right edge
Private Sub Con_DrawInput()
    Dim i As Long, text As Long

    If key_dest <> key_console And Not con_forcedup Then Exit Sub    ' don't draw anything

    ' add the cursor frame
    key_lines(edit_line, key_linepos) = 10 + (CLng(Fix(realtime * con_cursorspeed)) And 1)

    ' fill out remainder with spaces
    For i = key_linepos + 1 To con_linewidth - 1
        key_lines(edit_line, i) = 32
    Next

    ' prestep if horizontally scrolling
    text = 0
    If key_linepos >= con_linewidth Then text = 1 + key_linepos - con_linewidth

    ' draw it
    For i = 0 To con_linewidth - 1
        Draw_Character (i + 1) * 8, con_vislines - 16, key_lines(edit_line, text + i)
    Next

    ' remove cursor
    key_lines(edit_line, key_linepos) = 0
End Sub

' Draws the last few lines of output transparently over the game top
Public Sub Con_DrawNotify()
    Dim x As Long, v As Long, text As Long, i As Long, tm As Single

    v = 0
    For i = con_current - NUM_CON_TIMES + 1 To con_current
        If i < 0 Then GoTo nexti
        tm = con_times(i Mod NUM_CON_TIMES)
        If tm = 0 Then GoTo nexti
        tm = realtime - tm
        If tm > cvars(cv_con_notifytime).value Then GoTo nexti
        text = (i Mod con_totallines) * con_linewidth

        clearnotify = 0
        scr_copytop = 1

        For x = 0 To con_linewidth - 1
            Draw_Character (x + 1) * 8, v, con_text(text + x)
        Next

        v = v + 8
nexti:
    Next

    If key_dest = key_message Then
        clearnotify = 0
        scr_copytop = 1

        Draw_String 8, v, "say:"
        For x = 0 To Len(chat_buffer) - 1
            Draw_Character (x + 5) * 8, v, AscW(Mid$(chat_buffer, x + 1, 1))
        Next
        Draw_Character (x + 5) * 8, v, 10 + (CLng(Fix(realtime * con_cursorspeed)) And 1)
        v = v + 8
    End If

    If v > con_notifylines Then con_notifylines = v
End Sub

' Draws the console with the solid background
' The typing input line at the bottom should only be drawn if typing is allowed
Public Sub Con_DrawConsole(ByVal lines As Long, ByVal drawinput As Boolean)
    Dim i As Long, x As Long, y As Long, rows As Long, text As Long, j As Long

    If lines <= 0 Then Exit Sub

    ' draw the background
    Draw_ConsoleBackground lines

    ' draw the text
    con_vislines = lines

    rows = (lines - 16) \ 8                  ' rows of text to draw
    y = lines - 16 - (rows * 8)              ' may start slightly negative

    i = con_current - rows + 1
    Do While i <= con_current
        j = i - con_backscroll
        If j < 0 Then j = 0
        text = (j Mod con_totallines) * con_linewidth

        For x = 0 To con_linewidth - 1
            Draw_Character (x + 1) * 8, y, con_text(text + x)
        Next
        i = i + 1
        y = y + 8
    Loop

    ' draw the input prompt, user text, and cursor if desired
    If drawinput Then Con_DrawInput
End Sub

Public Sub Con_NotifyBox(ByVal text As String)
    Dim t1 As Double, t2 As Double

    ' during startup for sound / cd warnings
    Con_Printf "\n\n\35\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\37\n"

    Con_Printf "%s", text
    Con_Printf "Press a key.\n"
    Con_Printf "\35\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\37\n"

    key_count = -2                       ' wait for a key down and up
    key_dest = key_console

    Do
        t1 = Sys_FloatTime()
        SCR_UpdateScreen
        Sys_SendKeyEvents
        t2 = Sys_FloatTime()
        realtime = realtime + t2 - t1    ' make the cursor blink
    Loop While key_count < 0

    Con_Printf "\n"
    key_dest = key_game
    realtime = 0                         ' put the cursor back to invisible
End Sub
