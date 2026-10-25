Attribute VB_Name = "cmd"
' cmd.bas - Quake script command processing module (cmd.c).
'
' Commands and aliases are kept in arrays; C links new ones in at the head of
' their lists, so lookups walk the arrays from the end. A command's function
' is an XC_ id (quakedef.bas) that Cmd_Call dispatches.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const MAX_ALIAS_NAME As Long = 32

Private alias_names() As String
Private alias_values() As String
Private num_aliases As Long

Private cmd_wait As Boolean

Private cmd_text As sizebuf_t

Private Const MAX_ARGS As Long = 80

Private cmd_argc_ As Long
Private cmd_argv_(0 To MAX_ARGS - 1) As String
Private cmd_args_ As String
Public cmd_source As Long

Private cmd_names() As String            ' possible commands to execute
Private cmd_funcs() As Long
Private num_cmds As Long

'=============================================================================

' Causes execution of the remainder of the command buffer to be delayed until
' next frame.  This allows commands like:
' bind g "impulse 5 ; +attack ; wait ; -attack ; impulse 2"
Public Sub Cmd_Wait_f()
    cmd_wait = True
End Sub

'=============================================================================
' COMMAND BUFFER

Public Sub Cbuf_Init()
    SZ_Alloc cmd_text, 8192              ' space for commands and script files
End Sub

' Adds command text at the end of the buffer
Public Sub Cbuf_AddText(ByVal text As String)
    Dim l As Long, p As Long

    l = Len(text)

    If cmd_text.cursize + l >= cmd_text.maxsize Then
        Con_Printf "Cbuf_AddText: overflow\n"
        Exit Sub
    End If

    p = SZ_GetSpace(cmd_text, l)
    StrToBuf text, cmd_text.data, p
End Sub

' Adds command text immediately after the current command
' Adds a \n to the text
Public Sub Cbuf_InsertText(ByVal text As String)
    Dim temp() As Byte, templen As Long

    ' copy off any commands still remaining in the exec buffer
    templen = cmd_text.cursize
    If templen Then
        ReDim temp(0 To templen - 1)
        CopyMemory temp(0), cmd_text.data(0), templen
        SZ_Clear cmd_text
    End If

    ' add the entire text of the file
    Cbuf_AddText text

    ' add the copied off data
    If templen Then SZ_Write cmd_text, temp, 0, templen
End Sub

Public Sub Cbuf_Execute()
    Dim i As Long, c As Long, quotes As Long, line_ As String

    Do While cmd_text.cursize
        ' find a \n or ; line break
        quotes = 0
        For i = 0 To cmd_text.cursize - 1
            c = cmd_text.data(i)
            If c = 34 Then quotes = quotes + 1
            If (quotes And 1) = 0 And c = 59 Then Exit For   ' don't break if inside a quoted string
            If c = 10 Then Exit For
        Next

        line_ = BytesToStr(cmd_text.data, 0, i)

        ' delete the text from the command buffer and move remaining commands down
        ' this is necessary because commands (exec, alias) can insert data at the
        ' beginning of the text buffer
        If i = cmd_text.cursize Then
            cmd_text.cursize = 0
        Else
            i = i + 1
            cmd_text.cursize = cmd_text.cursize - i
            If cmd_text.cursize > 0 Then CopyMemory cmd_text.data(0), cmd_text.data(i), cmd_text.cursize
        End If

        ' execute the command line
        Cmd_ExecuteString line_, src_command

        If cmd_wait Then
            ' skip out while text still remains in buffer, leaving it
            ' for next frame
            cmd_wait = False
            Exit Do
        End If
    Loop
End Sub

'==============================================================================
' SCRIPT COMMANDS

' Adds command line parameters as script statements
' Commands lead with a +, and continue until a - or another +
' quake +prog jctest.qp +cmd amlev1
' quake -nosound +cmd amlev1
Public Sub Cmd_StuffCmds_f()
    Dim i As Long, j As Long, text As String, build As String, c As String

    If Cmd_Argc() <> 1 Then
        Con_Printf "stuffcmds : execute command line parameters\n"
        Exit Sub
    End If

    ' build the combined string to parse from
    For i = 1 To com_argc - 1
        text = text & com_argv(i)
        If i <> com_argc - 1 Then text = text & " "
    Next
    If Len(text) = 0 Then Exit Sub

    ' pull out the commands (0-based positions as in C)
    i = 0
    Do While i < Len(text)
        If Mid$(text, i + 1, 1) = "+" Then
            i = i + 1
            j = i
            Do While j < Len(text)
                c = Mid$(text, j + 1, 1)
                If c = "+" Or c = "-" Then Exit Do
                j = j + 1
            Loop
            build = build & Mid$(text, i + 1, j - i) & vbLf
            i = j - 1
        End If
        i = i + 1
    Loop

    If Len(build) Then Cbuf_InsertText build
End Sub

Public Sub Cmd_Exec_f()
    Dim f() As Byte

    If Cmd_Argc() <> 2 Then
        Con_Printf "exec <filename> : execute a script file\n"
        Exit Sub
    End If

    If Not COM_LoadHunkFile(Cmd_Argv(1), f) Then
        Con_Printf "couldn't exec %s\n", Cmd_Argv(1)
        Exit Sub
    End If
    Con_Printf "execing %s\n", Cmd_Argv(1)

    Cbuf_InsertText BufToStr(f, 0)
End Sub

' Just prints the rest of the line to the console
Public Sub Cmd_Echo_f()
    Dim i As Long
    For i = 1 To Cmd_Argc() - 1
        Con_Printf "%s ", Cmd_Argv(i)
    Next
    Con_Printf "\n"
End Sub

' Creates a new command that executes a command string (possibly ; seperated)
Public Sub Cmd_Alias_f()
    Dim a As Long, cmd_ As String, i As Long, c As Long, s As String

    If Cmd_Argc() = 1 Then
        Con_Printf "Current alias commands:\n"
        For a = num_aliases To 1 Step -1
            Con_Printf "%s : %s\n", alias_names(a), alias_values(a)
        Next
        Exit Sub
    End If

    s = Cmd_Argv(1)
    If Len(s) >= MAX_ALIAS_NAME Then
        Con_Printf "Alias name is too long\n"
        Exit Sub
    End If

    ' if the alias allready exists, reuse it
    For a = num_aliases To 1 Step -1
        If s = alias_names(a) Then Exit For
    Next

    If a = 0 Then
        num_aliases = num_aliases + 1
        ReDim Preserve alias_names(1 To num_aliases)
        ReDim Preserve alias_values(1 To num_aliases)
        a = num_aliases
    End If
    alias_names(a) = s

    ' copy the rest of the command line
    cmd_ = ""                            ' start out with a null string
    c = Cmd_Argc()
    For i = 2 To c - 1
        cmd_ = cmd_ & Cmd_Argv(i)
        If i <> c Then cmd_ = cmd_ & " "
    Next
    cmd_ = cmd_ & vbLf

    alias_values(a) = cmd_
End Sub

'=============================================================================
' COMMAND EXECUTION

Public Sub Cmd_Init()
    '
    ' register our commands
    '
    Cmd_AddCommand "stuffcmds", XC_Cmd_StuffCmds_f
    Cmd_AddCommand "exec", XC_Cmd_Exec_f
    Cmd_AddCommand "echo", XC_Cmd_Echo_f
    Cmd_AddCommand "alias", XC_Cmd_Alias_f
    Cmd_AddCommand "cmd", XC_Cmd_ForwardToServer
    Cmd_AddCommand "wait", XC_Cmd_Wait_f
End Sub

Public Function Cmd_Argc() As Long
    Cmd_Argc = cmd_argc_
End Function

Public Function Cmd_Argv(ByVal arg As Long) As String
    If arg < 0 Or arg >= cmd_argc_ Then
        Cmd_Argv = ""
        Exit Function
    End If
    Cmd_Argv = cmd_argv_(arg)
End Function

Public Function Cmd_Args() As String
    Cmd_Args = cmd_args_
End Function

' Parses the given string into command line tokens.
Public Sub Cmd_TokenizeString(ByVal text As String)
    Dim b() As Byte, p As Long

    ' clear the args from the last string
    cmd_argc_ = 0
    cmd_args_ = ""

    StrToBytesZ text, b
    p = 0

    Do
        ' skip whitespace up to a /n
        Do While b(p) <> 0 And b(p) <= 32 And b(p) <> 10
            p = p + 1
        Loop

        If b(p) = 10 Then
            ' a newline seperates commands in the buffer
            p = p + 1
            Exit Do
        End If

        If b(p) = 0 Then Exit Sub

        If cmd_argc_ = 1 Then cmd_args_ = Mid$(text, p + 1)

        p = COM_Parse(b, p)
        If p = -1 Then Exit Sub

        If cmd_argc_ < MAX_ARGS Then
            cmd_argv_(cmd_argc_) = com_token
            cmd_argc_ = cmd_argc_ + 1
        End If
    Loop
End Sub

Public Sub Cmd_AddCommand(ByVal cmd_name As String, ByVal fn As Long)
    Dim i As Long

    If host_initialized Then Sys_Error "Cmd_AddCommand after host_initialized"   ' because hunk allocation would get stomped

    ' fail if the command is a variable name
    If Len(Cvar_VariableString(cmd_name)) > 0 Then
        Con_Printf "Cmd_AddCommand: %s already defined as a var\n", cmd_name
        Exit Sub
    End If

    ' fail if the command already exists
    For i = num_cmds To 1 Step -1
        If cmd_name = cmd_names(i) Then
            Con_Printf "Cmd_AddCommand: %s already defined\n", cmd_name
            Exit Sub
        End If
    Next

    num_cmds = num_cmds + 1
    ReDim Preserve cmd_names(1 To num_cmds)
    ReDim Preserve cmd_funcs(1 To num_cmds)
    cmd_names(num_cmds) = cmd_name
    cmd_funcs(num_cmds) = fn
End Sub

Public Function Cmd_Exists(ByVal cmd_name As String) As Boolean
    Dim i As Long
    For i = num_cmds To 1 Step -1
        If cmd_name = cmd_names(i) Then
            Cmd_Exists = True
            Exit Function
        End If
    Next
    Cmd_Exists = False
End Function

' "" if nothing matches (C: NULL)
Public Function Cmd_CompleteCommand(ByVal partial As String) As String
    Dim i As Long, l As Long

    l = Len(partial)
    If l = 0 Then Exit Function

    ' check functions
    For i = num_cmds To 1 Step -1
        If Q_strncmp(partial, cmd_names(i), l) = 0 Then
            Cmd_CompleteCommand = cmd_names(i)
            Exit Function
        End If
    Next
End Function

' A complete command line has been parsed, so try to execute it
Public Sub Cmd_ExecuteString(ByVal text As String, ByVal src As Long)
    Dim i As Long, a As Long

    cmd_source = src
    Cmd_TokenizeString text

    ' execute the command line
    If Cmd_Argc() = 0 Then Exit Sub      ' no tokens

    ' check functions
    For i = num_cmds To 1 Step -1
        If Q_strcasecmp(cmd_argv_(0), cmd_names(i)) = 0 Then
            Cmd_Call cmd_funcs(i)
            Exit Sub
        End If
    Next

    ' check alias
    For a = num_aliases To 1 Step -1
        If Q_strcasecmp(cmd_argv_(0), alias_names(a)) = 0 Then
            Cbuf_InsertText alias_values(a)
            Exit Sub
        End If
    Next

    ' check cvars
    If Not Cvar_Command() Then Con_Printf "Unknown command ""%s""\n", Cmd_Argv(0)
End Sub

' Sends the entire command line over to the server
Public Sub Cmd_ForwardToServer()
    If cls.state <> ca_connected Then
        Con_Printf "Can't ""%s"", not connected\n", Cmd_Argv(0)
        Exit Sub
    End If

    If cls.demoplayback Then Exit Sub    ' not really connected

    MSG_WriteByte cls.message, clc_stringcmd
    If Q_strcasecmp(Cmd_Argv(0), "cmd") <> 0 Then
        SZ_Print cls.message, Cmd_Argv(0)
        SZ_Print cls.message, " "
    End If
    If Cmd_Argc() > 1 Then
        SZ_Print cls.message, Cmd_Args()
    Else
        SZ_Print cls.message, vbLf
    End If
End Sub

' Returns the position (1 to argc-1) in the command's argument list
' where the given parameter apears, or 0 if not present
Public Function Cmd_CheckParm(ByVal parm As String) As Long
    Dim i As Long
    For i = 1 To Cmd_Argc() - 1
        If Q_strcasecmp(parm, Cmd_Argv(i)) = 0 Then
            Cmd_CheckParm = i
            Exit Function
        End If
    Next
    Cmd_CheckParm = 0
End Function

' C: cmd->function () -- the xcommand_t of every registered command
Private Sub Cmd_Call(ByVal fn As Long)
    Select Case fn
    Case XC_Cmd_StuffCmds_f: Cmd_StuffCmds_f
    Case XC_Cmd_Exec_f: Cmd_Exec_f
    Case XC_Cmd_Echo_f: Cmd_Echo_f
    Case XC_Cmd_Alias_f: Cmd_Alias_f
    Case XC_Cmd_ForwardToServer: Cmd_ForwardToServer
    Case XC_Cmd_Wait_f: Cmd_Wait_f
    Case XC_COM_Path_f: COM_Path_f
    Case XC_Con_ToggleConsole_f: Con_ToggleConsole_f
    Case XC_Con_MessageMode_f: Con_MessageMode_f
    Case XC_Con_MessageMode2_f: Con_MessageMode2_f
    Case XC_Con_Clear_f: Con_Clear_f
    Case XC_Key_Bind_f: Key_Bind_f
    Case XC_Key_Unbind_f: Key_Unbind_f
    Case XC_Key_Unbindall_f: Key_Unbindall_f
    Case XC_SCR_ScreenShot_f: SCR_ScreenShot_f
    Case XC_SCR_SizeUp_f: SCR_SizeUp_f
    Case XC_SCR_SizeDown_f: SCR_SizeDown_f
    Case XC_Host_Quit_f To XC_Host_Viewprev_f: Host_CallCommand fn
    Case XC_Mod_Print: Mod_Print
    Case XC_ED_PrintEdict_f: ED_PrintEdict_f
    Case XC_ED_PrintEdicts: ED_PrintEdicts
    Case XC_ED_Count: ED_Count
    Case XC_PR_Profile_f: PR_Profile_f
    Case XC_NET_Slist_f To XC_NET_Port_f: NET_CallCommand fn
    Case XC_CL_PrintEntities_f To XC_CL_TimeDemo_f: CL_CallCommand fn
    Case XC_IN_KLookDown To XC_IN_Impulse: CL_InputCommand fn
    Case XC_V_cshift_f To XC_V_StartPitchDrift: V_CallCommand fn
    Case XC_Force_CenterView_f To XC_Joy_AdvancedUpdate_f: IN_CallCommand fn
    Case XC_R_TimeRefresh_f To XC_R_ReadPointFile_f: R_CallCommand fn
    Case XC_Sbar_ShowScores To XC_Sbar_DontShowScores: Sbar_CallCommand fn
    Case XC_M_ToggleMenu_f To XC_M_Menu_Quit_f: M_CallCommand fn
    Case XC_S_Play To XC_S_SoundInfo_f: S_CallCommand fn
    Case XC_CD_f: CD_f
    Case XC_NET_Stats_f To XC_Test2_f: Datagram_CallCommand fn
    Case Else: Sys_Error "Cmd_Call: bad function %i", fn
    End Select
End Sub
