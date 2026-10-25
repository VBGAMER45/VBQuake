Attribute VB_Name = "keys"
' keys.bas - key bindings, console line editing and chat input (keys.c).
' key up events are sent even if in console mode
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private shift_down As Boolean
Private history_line As Long

Private consolekeys(0 To 255) As Boolean ' if true, can't be rebound while in console
Private menubound(0 To 255) As Boolean   ' if true, can't be rebound while in menu
Private keyshift(0 To 255) As Long       ' key to map to if shift held down in console
Private keydown(0 To 255) As Boolean

Private keyname_names() As String
Private keyname_nums() As Long
Private num_keynames As Long

Public chat_buffer As String
Public team_message As Boolean
Private chat_bufferlen As Long

'==============================================================================
' key_lines rows are C strings: these read and copy them

' the zero-terminated text of key_lines(line) from position start
Public Function KeyLineStr(ByVal line_ As Long, ByVal start As Long) As String
    Dim e As Long, s As String
    e = start
    Do While e < MAXCMDLINE
        If key_lines(line_, e) = 0 Then Exit Do
        s = s & ChrW$(key_lines(line_, e))
        e = e + 1
    Loop
    KeyLineStr = s
End Function

' C: strlen (key_lines[line])
Private Function KeyLineLen(ByVal line_ As Long) As Long
    Dim e As Long
    Do While e < MAXCMDLINE
        If key_lines(line_, e) = 0 Then Exit Do
        e = e + 1
    Loop
    KeyLineLen = e
End Function

' C: Q_strcpy (key_lines[dst], key_lines[src])
Private Sub KeyLineCopy(ByVal dst As Long, ByVal src As Long)
    Dim i As Long
    For i = 0 To MAXCMDLINE - 1
        key_lines(dst, i) = key_lines(src, i)
        If key_lines(src, i) = 0 Then Exit For
    Next
End Sub

Private Sub AddKeyName(ByVal name As String, ByVal keynum As Long)
    num_keynames = num_keynames + 1
    ReDim Preserve keyname_names(1 To num_keynames)
    ReDim Preserve keyname_nums(1 To num_keynames)
    keyname_names(num_keynames) = name
    keyname_nums(num_keynames) = keynum
End Sub

Private Sub InitKeyNames()
    Dim i As Long
    AddKeyName "TAB", K_TAB
    AddKeyName "ENTER", K_ENTER
    AddKeyName "ESCAPE", K_ESCAPE
    AddKeyName "SPACE", K_SPACE
    AddKeyName "BACKSPACE", K_BACKSPACE
    AddKeyName "UPARROW", K_UPARROW
    AddKeyName "DOWNARROW", K_DOWNARROW
    AddKeyName "LEFTARROW", K_LEFTARROW
    AddKeyName "RIGHTARROW", K_RIGHTARROW

    AddKeyName "ALT", K_ALT
    AddKeyName "CTRL", K_CTRL
    AddKeyName "SHIFT", K_SHIFT

    For i = 1 To 12
        AddKeyName "F" & i, K_F1 + i - 1
    Next

    AddKeyName "INS", K_INS
    AddKeyName "DEL", K_DEL
    AddKeyName "PGDN", K_PGDN
    AddKeyName "PGUP", K_PGUP
    AddKeyName "HOME", K_HOME
    AddKeyName "END", K_END

    AddKeyName "MOUSE1", K_MOUSE1
    AddKeyName "MOUSE2", K_MOUSE2
    AddKeyName "MOUSE3", K_MOUSE3

    AddKeyName "JOY1", K_JOY1
    AddKeyName "JOY2", K_JOY2
    AddKeyName "JOY3", K_JOY3
    AddKeyName "JOY4", K_JOY4

    For i = 1 To 32
        AddKeyName "AUX" & i, K_AUX1 + i - 1
    Next

    AddKeyName "PAUSE", K_PAUSE

    AddKeyName "MWHEELUP", K_MWHEELUP
    AddKeyName "MWHEELDOWN", K_MWHEELDOWN

    AddKeyName "SEMICOLON", 59           ' because a raw semicolon seperates commands
End Sub

'==============================================================================
' LINE TYPING INTO THE CONSOLE

' Interactive line editing and console scrollback
Private Sub Key_Console_(ByVal key As Long)
    Dim cmd_ As String, i As Long

    If key = K_ENTER Then
        Cbuf_AddText KeyLineStr(edit_line, 1)    ' skip the >
        Cbuf_AddText vbLf
        Con_Printf "%s\n", KeyLineStr(edit_line, 0)
        edit_line = (edit_line + 1) And 31
        history_line = edit_line
        key_lines(edit_line, 0) = 93             ' ']'
        key_linepos = 1
        If cls.state = ca_disconnected Then SCR_UpdateScreen  ' force an update, because the command
                                                              ' may take some time
        Exit Sub
    End If

    If key = K_TAB Then
        ' command completion
        cmd_ = Cmd_CompleteCommand(KeyLineStr(edit_line, 1))
        If Len(cmd_) = 0 Then cmd_ = Cvar_CompleteVariable(KeyLineStr(edit_line, 1))
        If Len(cmd_) Then
            For i = 1 To Len(cmd_)
                key_lines(edit_line, i) = AscW(Mid$(cmd_, i, 1)) And 255
            Next
            key_linepos = Len(cmd_) + 1
            key_lines(edit_line, key_linepos) = 32
            key_linepos = key_linepos + 1
            key_lines(edit_line, key_linepos) = 0
            Exit Sub
        End If
    End If

    If key = K_BACKSPACE Or key = K_LEFTARROW Then
        If key_linepos > 1 Then key_linepos = key_linepos - 1
        Exit Sub
    End If

    If key = K_UPARROW Then
        Do
            history_line = (history_line - 1) And 31
            If history_line = edit_line Then Exit Do
        Loop While key_lines(history_line, 1) = 0
        If history_line = edit_line Then history_line = (edit_line + 1) And 31
        KeyLineCopy edit_line, history_line
        key_linepos = KeyLineLen(edit_line)
        Exit Sub
    End If

    If key = K_DOWNARROW Then
        If history_line = edit_line Then Exit Sub
        Do
            history_line = (history_line + 1) And 31
            If history_line = edit_line Then Exit Do
        Loop While key_lines(history_line, 1) = 0
        If history_line = edit_line Then
            key_lines(edit_line, 0) = 93
            key_linepos = 1
        Else
            KeyLineCopy edit_line, history_line
            key_linepos = KeyLineLen(edit_line)
        End If
        Exit Sub
    End If

    If key = K_PGUP Or key = K_MWHEELUP Then
        con_backscroll = con_backscroll + 2
        If con_backscroll > con_totallines - (vid.height \ 8) - 1 Then con_backscroll = con_totallines - (vid.height \ 8) - 1
        Exit Sub
    End If

    If key = K_PGDN Or key = K_MWHEELDOWN Then
        con_backscroll = con_backscroll - 2
        If con_backscroll < 0 Then con_backscroll = 0
        Exit Sub
    End If

    If key = K_HOME Then
        con_backscroll = con_totallines - (vid.height \ 8) - 1
        Exit Sub
    End If

    If key = K_END Then
        con_backscroll = 0
        Exit Sub
    End If

    If key < 32 Or key > 127 Then Exit Sub   ' non printable

    If key_linepos < MAXCMDLINE - 1 Then
        key_lines(edit_line, key_linepos) = key
        key_linepos = key_linepos + 1
        key_lines(edit_line, key_linepos) = 0
    End If
End Sub

'============================================================================

Private Sub Key_Message_(ByVal key As Long)
    If key = K_ENTER Then
        If team_message Then
            Cbuf_AddText "say_team """
        Else
            Cbuf_AddText "say """
        End If
        Cbuf_AddText chat_buffer
        Cbuf_AddText """" & vbLf

        key_dest = key_game
        chat_bufferlen = 0
        chat_buffer = ""
        Exit Sub
    End If

    If key = K_ESCAPE Then
        key_dest = key_game
        chat_bufferlen = 0
        chat_buffer = ""
        Exit Sub
    End If

    If key < 32 Or key > 127 Then Exit Sub   ' non printable

    If key = K_BACKSPACE Then
        If chat_bufferlen Then
            chat_bufferlen = chat_bufferlen - 1
            chat_buffer = Left$(chat_buffer, chat_bufferlen)
        End If
        Exit Sub
    End If

    If chat_bufferlen = 31 Then Exit Sub     ' all full

    chat_buffer = chat_buffer & ChrW$(key)
    chat_bufferlen = chat_bufferlen + 1
End Sub

'============================================================================

' Returns a key number to be used to index keybindings[] by looking at
' the given string.  Single ascii characters return themselves, while
' the K_* names are matched up.
Public Function Key_StringToKeynum(ByVal str_ As String) As Long
    Dim kn As Long

    If Len(str_) = 0 Then
        Key_StringToKeynum = -1
        Exit Function
    End If
    If Len(str_) = 1 Then
        Key_StringToKeynum = AscW(str_) And 255
        Exit Function
    End If

    For kn = 1 To num_keynames
        If Q_strcasecmp(str_, keyname_names(kn)) = 0 Then
            Key_StringToKeynum = keyname_nums(kn)
            Exit Function
        End If
    Next
    Key_StringToKeynum = -1
End Function

' Returns a string (either a single ascii char, or a K_* name) for the
' given keynum.
Public Function Key_KeynumToString(ByVal keynum As Long) As String
    Dim kn As Long

    If keynum = -1 Then
        Key_KeynumToString = "<KEY NOT FOUND>"
        Exit Function
    End If
    If keynum > 32 And keynum < 127 Then
        ' printable ascii
        Key_KeynumToString = ChrW$(keynum)
        Exit Function
    End If

    For kn = 1 To num_keynames
        If keynum = keyname_nums(kn) Then
            Key_KeynumToString = keyname_names(kn)
            Exit Function
        End If
    Next

    Key_KeynumToString = "<UNKNOWN KEYNUM>"
End Function

Public Sub Key_SetBinding(ByVal keynum As Long, ByVal binding As String)
    If keynum = -1 Then Exit Sub
    keybindings(keynum) = binding
    keybinding_set(keynum) = True
End Sub

Public Sub Key_Unbind_f()
    Dim b As Long

    If Cmd_Argc() <> 2 Then
        Con_Printf "unbind <key> : remove commands from a key\n"
        Exit Sub
    End If

    b = Key_StringToKeynum(Cmd_Argv(1))
    If b = -1 Then
        Con_Printf """%s"" isn't a valid key\n", Cmd_Argv(1)
        Exit Sub
    End If

    Key_SetBinding b, ""
End Sub

Public Sub Key_Unbindall_f()
    Dim i As Long
    For i = 0 To 255
        If keybinding_set(i) Then Key_SetBinding i, ""
    Next
End Sub

Public Sub Key_Bind_f()
    Dim i As Long, c As Long, b As Long, cmd_ As String

    c = Cmd_Argc()

    If c <> 2 And c <> 3 Then
        Con_Printf "bind <key> [command] : attach a command to a key\n"
        Exit Sub
    End If
    b = Key_StringToKeynum(Cmd_Argv(1))
    If b = -1 Then
        Con_Printf """%s"" isn't a valid key\n", Cmd_Argv(1)
        Exit Sub
    End If

    If c = 2 Then
        If keybinding_set(b) Then
            Con_Printf """%s"" = ""%s""\n", Cmd_Argv(1), keybindings(b)
        Else
            Con_Printf """%s"" is not bound\n", Cmd_Argv(1)
        End If
        Exit Sub
    End If

    ' copy the rest of the command line
    cmd_ = ""                            ' start out with a null string
    For i = 2 To c - 1
        If i > 2 Then cmd_ = cmd_ & " "
        cmd_ = cmd_ & Cmd_Argv(i)
    Next

    Key_SetBinding b, cmd_
End Sub

' Writes lines containing "bind key value" (f: an open VB file number)
Public Sub Key_WriteBindings(ByVal f As Integer)
    Dim i As Long
    For i = 0 To 255
        If keybinding_set(i) Then
            If Len(keybindings(i)) Then Print #f, va("bind ""%s"" ""%s""", Key_KeynumToString(i), keybindings(i))
        End If
    Next
End Sub

Public Sub Key_Init()
    Dim i As Long

    InitKeyNames

    For i = 0 To 31
        key_lines(i, 0) = 93             ' ']'
        key_lines(i, 1) = 0
    Next
    key_linepos = 1

    '
    ' init ascii characters in console mode
    '
    For i = 32 To 127
        consolekeys(i) = True
    Next
    consolekeys(K_ENTER) = True
    consolekeys(K_TAB) = True
    consolekeys(K_LEFTARROW) = True
    consolekeys(K_RIGHTARROW) = True
    consolekeys(K_UPARROW) = True
    consolekeys(K_DOWNARROW) = True
    consolekeys(K_BACKSPACE) = True
    consolekeys(K_PGUP) = True
    consolekeys(K_PGDN) = True
    consolekeys(K_SHIFT) = True
    consolekeys(K_MWHEELUP) = True
    consolekeys(K_MWHEELDOWN) = True
    consolekeys(96) = False              ' '`'
    consolekeys(126) = False             ' '~'

    For i = 0 To 255
        keyshift(i) = i
    Next
    For i = 97 To 122                    ' 'a' .. 'z'
        keyshift(i) = i - 97 + 65
    Next
    keyshift(Asc("1")) = Asc("!")
    keyshift(Asc("2")) = Asc("@")
    keyshift(Asc("3")) = Asc("#")
    keyshift(Asc("4")) = Asc("$")
    keyshift(Asc("5")) = Asc("%")
    keyshift(Asc("6")) = Asc("^")
    keyshift(Asc("7")) = Asc("&")
    keyshift(Asc("8")) = Asc("*")
    keyshift(Asc("9")) = Asc("(")
    keyshift(Asc("0")) = Asc(")")
    keyshift(Asc("-")) = Asc("_")
    keyshift(Asc("=")) = Asc("+")
    keyshift(Asc(",")) = Asc("<")
    keyshift(Asc(".")) = Asc(">")
    keyshift(Asc("/")) = Asc("?")
    keyshift(Asc(";")) = Asc(":")
    keyshift(Asc("'")) = Asc("""")
    keyshift(Asc("[")) = Asc("{")
    keyshift(Asc("]")) = Asc("}")
    keyshift(Asc("`")) = Asc("~")
    keyshift(Asc("\")) = Asc("|")

    menubound(K_ESCAPE) = True
    For i = 0 To 11
        menubound(K_F1 + i) = True
    Next

    '
    ' register our functions
    '
    Cmd_AddCommand "bind", XC_Key_Bind_f
    Cmd_AddCommand "unbind", XC_Key_Unbind_f
    Cmd_AddCommand "unbindall", XC_Key_Unbindall_f
End Sub

' Called by the system between frames for both key up and key down events
' Should NOT be called during an interrupt!
Public Sub Key_Event(ByVal key As Long, ByVal down As Boolean)
    Dim kb As String

    keydown(key) = down

    If Not down Then key_repeats(key) = 0

    key_lastpress = key
    key_count = key_count + 1
    If key_count <= 0 Then Exit Sub      ' just catching keys for Con_NotifyBox

    ' update auto-repeat status
    If down Then
        key_repeats(key) = key_repeats(key) + 1
        If key <> K_BACKSPACE And key <> K_PAUSE And key_repeats(key) > 1 Then
            Exit Sub                     ' ignore most autorepeats
        End If

        If key >= 200 And Not keybinding_set(key) Then
            Con_Printf "%s is unbound, hit F4 to set.\n", Key_KeynumToString(key)
        End If
    End If

    If key = K_SHIFT Then shift_down = down

    '
    ' handle escape specialy, so the user can never unbind it
    '
    If key = K_ESCAPE Then
        If Not down Then Exit Sub
        Select Case key_dest
        Case key_message
            Key_Message_ key
        Case key_menu
            M_Keydown key
        Case key_game, key_console
            M_ToggleMenu_f
        Case Else
            Sys_Error "Bad key_dest"
        End Select
        Exit Sub
    End If

    '
    ' key up events only generate commands if the game key binding is
    ' a button command (leading + sign).  These will occur even in console mode,
    ' to keep the character from continuing an action started before a console
    ' switch.  Button commands include the kenum as a parameter, so multiple
    ' downs can be matched with ups
    '
    If Not down Then
        If keybinding_set(key) Then
            kb = keybindings(key)
            If Left$(kb, 1) = "+" Then Cbuf_AddText va("-%s %i\n", Mid$(kb, 2), key)
        End If
        If keyshift(key) <> key Then
            If keybinding_set(keyshift(key)) Then
                kb = keybindings(keyshift(key))
                If Left$(kb, 1) = "+" Then Cbuf_AddText va("-%s %i\n", Mid$(kb, 2), key)
            End If
        End If
        Exit Sub
    End If

    '
    ' during demo playback, most keys bring up the main menu
    '
    If cls.demoplayback And down And consolekeys(key) And key_dest = key_game Then
        M_ToggleMenu_f
        Exit Sub
    End If

    '
    ' if not a consolekey, send to the interpreter no matter what mode is
    '
    If (key_dest = key_menu And menubound(key)) _
       Or (key_dest = key_console And Not consolekeys(key)) _
       Or (key_dest = key_game And (Not con_forcedup Or Not consolekeys(key))) Then
        If keybinding_set(key) Then
            kb = keybindings(key)
            If Left$(kb, 1) = "+" Then
                ' button commands add keynum as a parm
                Cbuf_AddText va("%s %i\n", kb, key)
            Else
                Cbuf_AddText kb
                Cbuf_AddText vbLf
            End If
        End If
        Exit Sub
    End If

    If Not down Then Exit Sub            ' other systems only care about key down events

    If shift_down Then key = keyshift(key)

    Select Case key_dest
    Case key_message
        Key_Message_ key
    Case key_menu
        M_Keydown key
    Case key_game, key_console
        Key_Console_ key
    Case Else
        Sys_Error "Bad key_dest"
    End Select
End Sub

Public Sub Key_ClearStates()
    Dim i As Long
    For i = 0 To 255
        keydown(i) = False
        key_repeats(i) = 0
    Next
End Sub
