Attribute VB_Name = "cl_demo"
' cl_demo.bas - demo recording and playback (cl_demo.c)
'
' When a demo is playing back, all NET_SendMessages are skipped, and
' NET_GetMessages are read from the demo file.
'
' Whenever cl.time gets past the last received message, another message is
' read from the demo file.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' Called when a demo file runs out, or the user starts a game
Public Sub CL_StopPlayback()
    If Not cls.demoplayback Then Exit Sub

    Sys_FileClose cls.demofile
    cls.demoplayback = False
    cls.demofile = 0
    cls.state = ca_disconnected

    If cls.timedemo Then CL_FinishTimeDemo
End Sub

' Dumps the current net message, prefixed by the length and view angles
Private Sub CL_WriteDemoMessage()
    Dim header(0 To 15) As Byte, i As Long, f As Single

    CopyMemory header(0), net_message.cursize, 4
    For i = 0 To 2
        f = cl.viewangles.v(i)
        CopyMemory header(4 + i * 4), f, 4
    Next
    Sys_FileWrite cls.demofile, header, 0, 16
    Sys_FileWrite cls.demofile, net_message.data, 0, net_message.cursize
End Sub

' Handles recording and playback of demos, on top of NET_ code
Public Function CL_GetMessage() As Long
    Dim r As Long, i As Long, f As Single, b(0 To 3) As Byte

    If cls.demoplayback Then
        ' decide if it is time to grab the next message
        If cls.signon = SIGNONS Then     ' allways grab until fully connected
            If cls.timedemo Then
                If host_framecount = cls.td_lastframe Then Exit Function   ' allready read this frame's message
                cls.td_lastframe = host_framecount
                ' if this is the second frame, grab the real td_starttime
                ' so the bogus time on the first frame doesn't count
                If host_framecount = cls.td_startframe + 1 Then cls.td_starttime = realtime
            ElseIf cl.time <= cl.mtime(0) Then   ' /* cl.time > 0 && */
                Exit Function            ' don't need another message yet
            End If
        End If

        ' get the next message
        ' (C's fread leaves cursize and the angles as they were at the end of the file)
        If Sys_FileRead(cls.demofile, b, 0, 4) = 4 Then CopyMemory net_message.cursize, b(0), 4
        cl.mviewangles(1) = cl.mviewangles(0)
        For i = 0 To 2
            If Sys_FileRead(cls.demofile, b, 0, 4) = 4 Then
                CopyMemory f, b(0), 4
                cl.mviewangles(0).v(i) = f
            End If
        Next

        If net_message.cursize > MAX_MSGLEN Then Sys_Error "Demo message > MAX_MSGLEN"
        ' (C: r = fread (data, cursize, 1, ...): 0 for a short read, and for
        ' cursize 0)
        r = 0
        If net_message.cursize > 0 Then
            If Sys_FileRead(cls.demofile, net_message.data, 0, net_message.cursize) = net_message.cursize Then r = 1
        End If
        If r <> 1 Then
            CL_StopPlayback
            Exit Function
        End If

        CL_GetMessage = 1
        Exit Function
    End If

    Do
        r = NET_GetMessage(cls.netcon)

        If r <> 1 And r <> 2 Then
            CL_GetMessage = r
            Exit Function
        End If

        ' discard nop keepalive message
        If net_message.cursize = 1 And net_message.data(0) = svc_nop Then
            Con_Printf "<-- server to client keepalive\n"
        Else
            Exit Do
        End If
    Loop

    If cls.demorecording Then CL_WriteDemoMessage

    CL_GetMessage = r
End Function

' stop recording a demo
Public Sub CL_Stop_f()
    If cmd_source <> src_command Then Exit Sub

    If Not cls.demorecording Then
        Con_Printf "Not recording a demo.\n"
        Exit Sub
    End If

    ' write a disconnect message to the demo file
    SZ_Clear net_message
    MSG_WriteByte net_message, svc_disconnect
    CL_WriteDemoMessage

    ' finish up
    Sys_FileClose cls.demofile
    cls.demofile = 0
    cls.demorecording = False
    Con_Printf "Completed demo\n"
End Sub

' record <demoname> <map> [cd track]
Public Sub CL_Record_f()
    Dim c As Long, name As String, track As Long, h As Long, b() As Byte

    If cmd_source <> src_command Then Exit Sub

    c = Cmd_Argc()
    If c <> 2 And c <> 3 And c <> 4 Then
        Con_Printf "record <demoname> [<map> [cd track]]\n"
        Exit Sub
    End If

    If InStr(Cmd_Argv(1), "..") Then
        Con_Printf "Relative pathnames are not allowed.\n"
        Exit Sub
    End If

    If c = 2 And cls.state = ca_connected Then
        Con_Printf "Can not record - already connected to server\nClient demo recording must be started before connecting\n"
        Exit Sub
    End If

    ' write the forced cd track number, or -1
    If c = 4 Then
        track = C_atoi(Cmd_Argv(3))
        Con_Printf "Forcing CD track to %i\n", cls.forcetrack
    Else
        track = -1
    End If

    name = com_gamedir & "/" & Cmd_Argv(1)

    '
    ' start the map up
    '
    If c > 2 Then Cmd_ExecuteString va("map %s", Cmd_Argv(2)), src_command

    '
    ' open the demo file
    '
    name = COM_DefaultExtension(name, ".dem")

    Con_Printf "recording to %s.\n", name
    h = Sys_FileCreate(name)
    If h = -1 Then
        Con_Printf "ERROR: couldn't open.\n"
        Exit Sub
    End If
    cls.demofile = h

    cls.forcetrack = track
    StrToBytesZ CStr(cls.forcetrack) & vbLf, b
    Sys_FileWrite cls.demofile, b, 0, UBound(b)

    cls.demorecording = True
End Sub

' play [demoname]
Public Sub CL_PlayDemo_f()
    Dim name As String, c As Long, neg As Boolean, h As Long, b(0 To 0) As Byte

    If cmd_source <> src_command Then Exit Sub

    If Cmd_Argc() <> 2 Then
        Con_Printf "play <demoname> : plays a demo\n"
        Exit Sub
    End If

    '
    ' disconnect from server
    '
    CL_Disconnect

    '
    ' open the demo file
    '
    name = COM_DefaultExtension(Cmd_Argv(1), ".dem")

    Con_Printf "Playing demo from %s.\n", name
    COM_FOpenFile name, h
    If h = -1 Then
        Con_Printf "ERROR: couldn't open.\n"
        cls.demonum = -1                 ' stop demo loop
        Exit Sub
    End If
    cls.demofile = h

    cls.demoplayback = True
    cls.state = ca_connected
    cls.forcetrack = 0

    Do
        ' (QuakeVB: stops at the end of the file too, where C's getc loop
        ' would run forever)
        If Sys_FileRead(cls.demofile, b, 0, 1) <> 1 Then Exit Do
        c = b(0)
        If c = 10 Then Exit Do
        If c = 45 Then                   ' '-'
            neg = True
        Else
            cls.forcetrack = cls.forcetrack * 10 + (c - 48)
        End If
    Loop

    If neg Then cls.forcetrack = -cls.forcetrack
    ' ZOID, fscanf is evil
    ' fscanf (cls.demofile, "%i\n", &cls.forcetrack);
End Sub

Private Sub CL_FinishTimeDemo()
    Dim frames As Long, time_ As Single

    cls.timedemo = False

    ' the first frame didn't count
    frames = (host_framecount - cls.td_startframe) - 1
    time_ = realtime - cls.td_starttime
    If time_ = 0 Then time_ = 1
    Con_Printf "%i frames %5.1f seconds %5.1f fps\n", frames, time_, frames / time_
End Sub

' timedemo [demoname]
Public Sub CL_TimeDemo_f()
    If cmd_source <> src_command Then Exit Sub

    If Cmd_Argc() <> 2 Then
        Con_Printf "timedemo <demoname> : gets demo speeds\n"
        Exit Sub
    End If

    CL_PlayDemo_f

    ' cls.td_starttime will be grabbed at the second frame of the demo, so
    ' all the loading time doesn't get counted

    cls.timedemo = True
    cls.td_startframe = host_framecount
    cls.td_lastframe = -1                ' get a new message this frame
End Sub
