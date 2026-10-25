Attribute VB_Name = "cl_parse"
' cl_parse.bas - parse a message received from the server (cl_parse.c)
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private svc_strings(0 To 34) As String
Private svc_strings_made As Boolean

Private bitcounts(0 To 15) As Long

Private keepalive_lastmsg As Single      ' C: static float lastmsg of CL_KeepaliveMessage

Private Function SvcString(ByVal cmd_ As Long) As String
    If Not svc_strings_made Then
        svc_strings(0) = "svc_bad"
        svc_strings(1) = "svc_nop"
        svc_strings(2) = "svc_disconnect"
        svc_strings(3) = "svc_updatestat"
        svc_strings(4) = "svc_version"           ' [long] server version
        svc_strings(5) = "svc_setview"           ' [short] entity number
        svc_strings(6) = "svc_sound"             ' <see code>
        svc_strings(7) = "svc_time"              ' [float] server time
        svc_strings(8) = "svc_print"             ' [string] null terminated string
        svc_strings(9) = "svc_stufftext"         ' [string] stuffed into client's console buffer
                                                 ' the string should be \n terminated
        svc_strings(10) = "svc_setangle"         ' [vec3] set the view angle to this absolute value
        svc_strings(11) = "svc_serverinfo"       ' [long] version
                                                 ' [string] signon string
                                                 ' [string]..[0]model cache [string]...[0]sounds cache
                                                 ' [string]..[0]item cache
        svc_strings(12) = "svc_lightstyle"       ' [byte] [string]
        svc_strings(13) = "svc_updatename"       ' [byte] [string]
        svc_strings(14) = "svc_updatefrags"      ' [byte] [short]
        svc_strings(15) = "svc_clientdata"       ' <shortbits + data>
        svc_strings(16) = "svc_stopsound"        ' <see code>
        svc_strings(17) = "svc_updatecolors"     ' [byte] [byte]
        svc_strings(18) = "svc_particle"         ' [vec3] <variable>
        svc_strings(19) = "svc_damage"           ' [byte] impact [byte] blood [vec3] from
        svc_strings(20) = "svc_spawnstatic"
        svc_strings(21) = "OBSOLETE svc_spawnbinary"
        svc_strings(22) = "svc_spawnbaseline"
        svc_strings(23) = "svc_temp_entity"      ' <variable>
        svc_strings(24) = "svc_setpause"
        svc_strings(25) = "svc_signonnum"
        svc_strings(26) = "svc_centerprint"
        svc_strings(27) = "svc_killedmonster"
        svc_strings(28) = "svc_foundsecret"
        svc_strings(29) = "svc_spawnstaticsound"
        svc_strings(30) = "svc_intermission"
        svc_strings(31) = "svc_finale"           ' [string] music [string] text
        svc_strings(32) = "svc_cdtrack"          ' [byte] track [byte] looptrack
        svc_strings(33) = "svc_sellscreen"
        svc_strings(34) = "svc_cutscene"
        svc_strings_made = True
    End If
    ' (C reads past the array for a bad command, before it reports it)
    If cmd_ >= 0 And cmd_ <= 34 Then SvcString = svc_strings(cmd_)
End Function

Private Sub SHOWNET(ByVal x As String)
    If cvars(cv_cl_shownet).value = 2 Then Con_Printf "%3i:%s\n", msg_readcount - 1, x
End Sub

'=============================================================================

' This error checks and tracks the total number of entities
' (returns a cl_entities() index)
Private Function CL_EntityNum(ByVal num As Long) As Long
    If num >= cl.num_entities Then
        If num >= MAX_EDICTS Then Host_Error "CL_EntityNum: %i is an invalid number", num
        Do While cl.num_entities <= num
            cl_entities(cl.num_entities).colormap = 0   ' vid.colormap
            cl.num_entities = cl.num_entities + 1
        Loop
    End If

    CL_EntityNum = num
End Function

Private Sub CL_ParseStartSoundPacket()
    Dim pos As vec3_t, channel As Long, ent As Long, sound_num As Long, volume As Long
    Dim field_mask As Long, attenuation As Single, i As Long

    field_mask = MSG_ReadByte()

    If field_mask And SND_VOLUME Then
        volume = MSG_ReadByte()
    Else
        volume = DEFAULT_SOUND_PACKET_VOLUME
    End If

    If field_mask And SND_ATTENUATION Then
        attenuation = MSG_ReadByte() / 64#
    Else
        attenuation = DEFAULT_SOUND_PACKET_ATTENUATION
    End If

    channel = MSG_ReadShort()
    sound_num = MSG_ReadByte()

    ent = (channel And &HFFFFFFF8) \ 8   ' channel >> 3
    channel = channel And 7

    If ent > MAX_EDICTS Then Host_Error "CL_ParseStartSoundPacket: ent = %i", ent

    For i = 0 To 2
        pos.v(i) = MSG_ReadCoord()
    Next

    S_StartSound ent, channel, cl.sound_precache(sound_num), pos, volume / 255#, attenuation
End Sub

' When the client is taking a long time to load stuff, send keepalive messages
' so the server doesn't disconnect.
Private Sub CL_KeepaliveMessage()
    Dim time_ As Single, ret As Long, old As sizebuf_t

    If sv.active Then Exit Sub           ' no need if server is local
    If cls.demoplayback Then Exit Sub

    ' read messages from server, should just be nops
    old = net_message                    ' (copies the data too)

    Do
        ret = CL_GetMessage()
        Select Case ret
        Case 0
            ' nothing waiting
        Case 1
            Host_Error "CL_KeepaliveMessage: received a message"
        Case 2
            If MSG_ReadByte() <> svc_nop Then Host_Error "CL_KeepaliveMessage: datagram wasn't a nop"
        Case Else
            Host_Error "CL_KeepaliveMessage: CL_GetMessage failed"
        End Select
    Loop While ret <> 0

    net_message = old

    ' check time
    time_ = Sys_FloatTime()
    If time_ - keepalive_lastmsg < 5 Then Exit Sub
    keepalive_lastmsg = time_

    ' write out a nop
    Con_Printf "--> client to server keepalive\n"

    MSG_WriteByte cls.message, clc_nop
    NET_SendMessage cls.netcon, cls.message
    SZ_Clear cls.message
End Sub

Private Sub CL_ParseServerInfo()
    Dim str_ As String, i As Long, nummodels As Long, numsounds As Long
    Dim model_precache(0 To MAX_MODELS - 1) As String, sound_precache(0 To MAX_SOUNDS - 1) As String

    Con_DPrintf "Serverinfo packet received.\n"
    '
    ' wipe the client_state_t struct
    '
    CL_ClearState

    ' parse protocol version number
    i = MSG_ReadLong()
    If i <> PROTOCOL_VERSION Then
        Con_Printf "Server returned version %i, not %i", i, PROTOCOL_VERSION
        Exit Sub
    End If

    ' parse maxclients
    cl.maxclients = MSG_ReadByte()
    If cl.maxclients < 1 Or cl.maxclients > MAX_SCOREBOARD Then
        Con_Printf "Bad maxclients (%u) from server\n", cl.maxclients
        Exit Sub
    End If
    ReDim cl.scores(0 To cl.maxclients - 1)

    ' parse gametype
    cl.gametype = MSG_ReadByte()

    ' parse signon message
    str_ = MSG_ReadString()
    cl.levelname = Left$(str_, 39)

    ' seperate the printfs so the server message can have a color
    Con_Printf "\n\n\35\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\36\37\n\n"
    Con_Printf "%c%s\n", 2, str_

    '
    ' first we go through and touch all of the precache data that still
    ' happens to be in the cache, so precaching something else doesn't
    ' needlessly purge it
    '

    ' precache models
    For i = 0 To MAX_MODELS - 1
        cl.model_precache(i) = 0
    Next
    nummodels = 1
    Do
        str_ = MSG_ReadString()
        If Len(str_) = 0 Then Exit Do
        If nummodels = MAX_MODELS Then
            Con_Printf "Server sent too many model precaches\n"
            Exit Sub
        End If
        model_precache(nummodels) = str_
        Mod_TouchModel str_
        nummodels = nummodels + 1
    Loop

    ' precache sounds
    For i = 0 To MAX_SOUNDS - 1
        cl.sound_precache(i) = 0
    Next
    numsounds = 1
    Do
        str_ = MSG_ReadString()
        If Len(str_) = 0 Then Exit Do
        If numsounds = MAX_SOUNDS Then
            Con_Printf "Server sent too many sound precaches\n"
            Exit Sub
        End If
        sound_precache(numsounds) = str_
        S_TouchSound str_
        numsounds = numsounds + 1
    Loop

    '
    ' now we try to load everything else until a cache allocation fails
    '

    For i = 1 To nummodels - 1
        cl.model_precache(i) = Mod_ForName(model_precache(i), False)
        If cl.model_precache(i) = 0 Then
            Con_Printf "Model %s not found\n", model_precache(i)
            Exit Sub
        End If
        CL_KeepaliveMessage
    Next

    S_BeginPrecaching
    For i = 1 To numsounds - 1
        cl.sound_precache(i) = S_PrecacheSound(sound_precache(i))
        CL_KeepaliveMessage
    Next
    S_EndPrecaching

    ' local state
    cl.worldmodel = cl.model_precache(1)
    cl_entities(0).model = cl.worldmodel

    R_NewMap

    ' Hunk_Check ();  (make sure nothing is hurt: there is no hunk)

    noclip_anglehack = False             ' noclip is turned off at start
End Sub

' Parse an entity update message from the server
' If an entities model or origin changes from frame to frame, it must be
' relinked.  Other attributes can change without relinking.
Private Sub CL_ParseUpdate(ByVal bits As Long)
    Dim i As Long, model As Long, modnum As Long, forcelink As Boolean, ent As Long, num As Long

    If cls.signon = SIGNONS - 1 Then
        ' first update is the final signon stage
        cls.signon = SIGNONS
        CL_SignonReply
    End If

    If bits And U_MOREBITS Then
        i = MSG_ReadByte()
        bits = bits Or (i * 256&)
    End If

    If bits And U_LONGENTITY Then
        num = MSG_ReadShort()
    Else
        num = MSG_ReadByte()
    End If

    ent = CL_EntityNum(num)

    For i = 0 To 15
        If bits And BIT_(i) Then bitcounts(i) = bitcounts(i) + 1
    Next

    With cl_entities(ent)
        If .msgtime <> cl.mtime(1) Then
            forcelink = True             ' no previous frame to lerp from
        Else
            forcelink = False
        End If

        .msgtime = cl.mtime(0)

        If bits And U_MODEL Then
            modnum = MSG_ReadByte()
            If modnum >= MAX_MODELS Then Host_Error "CL_ParseModel: bad modnum"
        Else
            modnum = .baseline.modelindex
        End If

        model = cl.model_precache(modnum)
        If model <> .model Then
            .model = model
            ' automatic animation (torches, etc) can be either all together
            ' or randomized
            If model Then
                If mod_known(model).synctype = ST_RAND Then
                    .syncbase = CSng(rand() And &H7FFF) / &H7FFF
                Else
                    .syncbase = 0
                End If
            Else
                forcelink = True         ' hack to make null model players work
            End If
        End If

        If bits And U_FRAME Then
            .frame = MSG_ReadByte()
        Else
            .frame = .baseline.frame
        End If

        If bits And U_COLORMAP Then
            i = MSG_ReadByte()
        Else
            i = .baseline.colormap
        End If
        If i = 0 Then
            .colormap = 0                ' vid.colormap
        Else
            If i > cl.maxclients Then Sys_Error "i >= cl.maxclients"
            .colormap = i                ' cl.scores[i-1].translations
        End If

        If bits And U_SKIN Then
            .skinnum = MSG_ReadByte()
        Else
            .skinnum = .baseline.skin
        End If

        If bits And U_EFFECTS Then
            .effects = MSG_ReadByte()
        Else
            .effects = .baseline.effects
        End If

        ' shift the known values for interpolation
        .msg_origins(1) = .msg_origins(0)
        .msg_angles(1) = .msg_angles(0)

        If bits And U_ORIGIN1 Then
            .msg_origins(0).v(0) = MSG_ReadCoord()
        Else
            .msg_origins(0).v(0) = .baseline.origin.v(0)
        End If
        If bits And U_ANGLE1 Then
            .msg_angles(0).v(0) = MSG_ReadAngle()
        Else
            .msg_angles(0).v(0) = .baseline.angles.v(0)
        End If

        If bits And U_ORIGIN2 Then
            .msg_origins(0).v(1) = MSG_ReadCoord()
        Else
            .msg_origins(0).v(1) = .baseline.origin.v(1)
        End If
        If bits And U_ANGLE2 Then
            .msg_angles(0).v(1) = MSG_ReadAngle()
        Else
            .msg_angles(0).v(1) = .baseline.angles.v(1)
        End If

        If bits And U_ORIGIN3 Then
            .msg_origins(0).v(2) = MSG_ReadCoord()
        Else
            .msg_origins(0).v(2) = .baseline.origin.v(2)
        End If
        If bits And U_ANGLE3 Then
            .msg_angles(0).v(2) = MSG_ReadAngle()
        Else
            .msg_angles(0).v(2) = .baseline.angles.v(2)
        End If

        If bits And U_NOLERP Then .forcelink = True

        If forcelink Then
            ' didn't have an update last message
            .msg_origins(1) = .msg_origins(0)
            .origin = .msg_origins(0)
            .msg_angles(1) = .msg_angles(0)
            .angles = .msg_angles(0)
            .forcelink = True
        End If
    End With
End Sub

Private Sub CL_ParseBaseline(ByVal ent As Long)
    Dim i As Long

    With cl_entities(ent).baseline
        .modelindex = MSG_ReadByte()
        .frame = MSG_ReadByte()
        .colormap = MSG_ReadByte()
        .skin = MSG_ReadByte()
        For i = 0 To 2
            .origin.v(i) = MSG_ReadCoord()
            .angles.v(i) = MSG_ReadAngle()
        Next
    End With
End Sub

' Server information pertaining to this client only
Private Sub CL_ParseClientdata(ByVal bits As Long)
    Dim i As Long, j As Long

    If bits And SU_VIEWHEIGHT Then
        cl.viewheight = MSG_ReadChar()
    Else
        cl.viewheight = DEFAULT_VIEWHEIGHT
    End If

    If bits And SU_IDEALPITCH Then
        cl.idealpitch = MSG_ReadChar()
    Else
        cl.idealpitch = 0
    End If

    cl.mvelocity(1) = cl.mvelocity(0)
    For i = 0 To 2
        If bits And (SU_PUNCH1 * BIT_(i)) Then
            cl.punchangle.v(i) = MSG_ReadChar()
        Else
            cl.punchangle.v(i) = 0
        End If
        If bits And (SU_VELOCITY1 * BIT_(i)) Then
            cl.mvelocity(0).v(i) = MSG_ReadChar() * 16
        Else
            cl.mvelocity(0).v(i) = 0
        End If
    Next

    ' [always sent]  if (bits & SU_ITEMS)
    i = MSG_ReadLong()

    If cl.items <> i Then
        ' set flash times
        Sbar_Changed
        For j = 0 To 31
            If (i And BIT_(j)) <> 0 And (cl.items And BIT_(j)) = 0 Then cl.item_gettime(j) = cl.time
        Next
        cl.items = i
    End If

    cl.onground = (bits And SU_ONGROUND) <> 0
    cl.inwater = (bits And SU_INWATER) <> 0

    If bits And SU_WEAPONFRAME Then
        cl.stats(STAT_WEAPONFRAME) = MSG_ReadByte()
    Else
        cl.stats(STAT_WEAPONFRAME) = 0
    End If

    If bits And SU_ARMOR Then
        i = MSG_ReadByte()
    Else
        i = 0
    End If
    If cl.stats(STAT_ARMOR) <> i Then
        cl.stats(STAT_ARMOR) = i
        Sbar_Changed
    End If

    If bits And SU_WEAPON Then
        i = MSG_ReadByte()
    Else
        i = 0
    End If
    If cl.stats(STAT_WEAPON) <> i Then
        cl.stats(STAT_WEAPON) = i
        Sbar_Changed
    End If

    i = MSG_ReadShort()
    If cl.stats(STAT_HEALTH) <> i Then
        cl.stats(STAT_HEALTH) = i
        Sbar_Changed
    End If

    i = MSG_ReadByte()
    If cl.stats(STAT_AMMO) <> i Then
        cl.stats(STAT_AMMO) = i
        Sbar_Changed
    End If

    For i = 0 To 3
        j = MSG_ReadByte()
        If cl.stats(STAT_SHELLS + i) <> j Then
            cl.stats(STAT_SHELLS + i) = j
            Sbar_Changed
        End If
    Next

    i = MSG_ReadByte()

    If standard_quake Then
        If cl.stats(STAT_ACTIVEWEAPON) <> i Then
            cl.stats(STAT_ACTIVEWEAPON) = i
            Sbar_Changed
        End If
    Else
        ' (C's 1 << i on x86 shifts by i & 31)
        If cl.stats(STAT_ACTIVEWEAPON) <> BIT_(i And 31) Then
            cl.stats(STAT_ACTIVEWEAPON) = BIT_(i And 31)
            Sbar_Changed
        End If
    End If
End Sub

Public Sub CL_NewTranslation(ByVal slot As Long)
    Dim i As Long, j As Long, top As Long, bottom As Long, dest As Long, source As Long

    If slot > cl.maxclients Then Sys_Error "CL_NewTranslation: slot > cl.maxclients"
    With cl.scores(slot)
        CopyMemory .translations(0), vid_colormap(0), VID_GRADES * 256
        top = .colors And &HF0
        bottom = (.colors And 15) * 16

        For i = 0 To VID_GRADES - 1
            dest = i * 256
            source = i * 256
            If top < 128 Then            ' the artists made some backwards ranges.  sigh.
                CopyMemory .translations(dest + TOP_RANGE), vid_colormap(source + top), 16
            Else
                For j = 0 To 15
                    .translations(dest + TOP_RANGE + j) = vid_colormap(source + top + 15 - j)
                Next
            End If

            If bottom < 128 Then
                CopyMemory .translations(dest + BOTTOM_RANGE), vid_colormap(source + bottom), 16
            Else
                For j = 0 To 15
                    .translations(dest + BOTTOM_RANGE + j) = vid_colormap(source + bottom + 15 - j)
                Next
            End If
        Next
    End With
End Sub

Private Sub CL_ParseStatic()
    Dim ent As Long, i As Long

    i = cl.num_statics
    If i >= MAX_STATIC_ENTITIES Then Host_Error "Too many static entities"
    ent = CL_STATIC_BASE + i
    cl.num_statics = cl.num_statics + 1
    CL_ParseBaseline ent

    ' copy it to the current state
    With cl_entities(ent)
        .model = cl.model_precache(.baseline.modelindex)
        .frame = .baseline.frame
        .colormap = 0                    ' vid.colormap
        .skinnum = .baseline.skin
        .effects = .baseline.effects

        .origin = .baseline.origin
        .angles = .baseline.angles
    End With
    R_AddEfrags ent
End Sub

Private Sub CL_ParseStaticSound()
    Dim org As vec3_t, sound_num As Long, vol As Long, atten As Long, i As Long

    For i = 0 To 2
        org.v(i) = MSG_ReadCoord()
    Next
    sound_num = MSG_ReadByte()
    vol = MSG_ReadByte()
    atten = MSG_ReadByte()

    S_StaticSound cl.sound_precache(sound_num), org, vol, atten
End Sub

Public Sub CL_ParseServerMessage()
    Dim cmd_ As Long, i As Long

    '
    ' if recording demos, copy the message out
    '
    If cvars(cv_cl_shownet).value = 1 Then
        Con_Printf "%i ", net_message.cursize
    ElseIf cvars(cv_cl_shownet).value = 2 Then
        Con_Printf "------------------\n"
    End If

    cl.onground = False                  ' unless the server says otherwise
    '
    ' parse the message
    '
    MSG_BeginReading

    Do
        If msg_badread Then Host_Error "CL_ParseServerMessage: Bad server message"

        cmd_ = MSG_ReadByte()

        If cmd_ = -1 Then
            SHOWNET "END OF MESSAGE"
            Exit Sub                     ' end of message
        End If

        ' if the high bit of the command byte is set, it is a fast update
        If cmd_ And 128 Then
            SHOWNET "fast update"
            CL_ParseUpdate cmd_ And 127
            GoTo nextcmd
        End If

        SHOWNET SvcString(cmd_)

        ' other commands
        Select Case cmd_
        Case svc_nop
            ' Con_Printf ("svc_nop\n");

        Case svc_time
            cl.mtime(1) = cl.mtime(0)
            cl.mtime(0) = MSG_ReadFloat()

        Case svc_clientdata
            i = MSG_ReadShort()
            CL_ParseClientdata i

        Case svc_version
            i = MSG_ReadLong()
            If i <> PROTOCOL_VERSION Then Host_Error "CL_ParseServerMessage: Server is protocol %i instead of %i\n", i, PROTOCOL_VERSION

        Case svc_disconnect
            Host_EndGame "Server disconnected\n"

        Case svc_print
            Con_Printf "%s", MSG_ReadString()

        Case svc_centerprint
            SCR_CenterPrint MSG_ReadString()

        Case svc_stufftext
            Cbuf_AddText MSG_ReadString()

        Case svc_damage
            V_ParseDamage

        Case svc_serverinfo
            CL_ParseServerInfo
            vid.recalc_refdef = 1        ' leave intermission full screen

        Case svc_setangle
            For i = 0 To 2
                cl.viewangles.v(i) = MSG_ReadAngle()
            Next

        Case svc_setview
            cl.viewentity = MSG_ReadShort()

        Case svc_lightstyle
            i = MSG_ReadByte()
            If i >= MAX_LIGHTSTYLES Then Sys_Error "svc_lightstyle > MAX_LIGHTSTYLES"
            cl_lightstyle(i).map = MSG_ReadString()
            cl_lightstyle(i).length = Len(cl_lightstyle(i).map)

        Case svc_sound
            CL_ParseStartSoundPacket

        Case svc_stopsound
            i = MSG_ReadShort()
            S_StopSound (i And &HFFFFFFF8) \ 8, i And 7

        Case svc_updatename
            Sbar_Changed
            i = MSG_ReadByte()
            If i >= cl.maxclients Then Host_Error "CL_ParseServerMessage: svc_updatename > MAX_SCOREBOARD"
            cl.scores(i).name = MSG_ReadString()

        Case svc_updatefrags
            Sbar_Changed
            i = MSG_ReadByte()
            If i >= cl.maxclients Then Host_Error "CL_ParseServerMessage: svc_updatefrags > MAX_SCOREBOARD"
            cl.scores(i).frags = MSG_ReadShort()

        Case svc_updatecolors
            Sbar_Changed
            i = MSG_ReadByte()
            If i >= cl.maxclients Then Host_Error "CL_ParseServerMessage: svc_updatecolors > MAX_SCOREBOARD"
            cl.scores(i).colors = MSG_ReadByte()
            CL_NewTranslation i

        Case svc_particle
            R_ParseParticleEffect

        Case svc_spawnbaseline
            i = MSG_ReadShort()
            ' must use CL_EntityNum() to force cl.num_entities up
            CL_ParseBaseline CL_EntityNum(i)

        Case svc_spawnstatic
            CL_ParseStatic

        Case svc_temp_entity
            CL_ParseTEnt

        Case svc_setpause
            cl.paused = MSG_ReadByte() <> 0

            If cl.paused Then
                CDAudio_Pause
                VID_HandlePause True
            Else
                CDAudio_Resume
                VID_HandlePause False
            End If

        Case svc_signonnum
            i = MSG_ReadByte()
            If i <= cls.signon Then Host_Error "Received signon %i when at %i", i, cls.signon
            cls.signon = i
            CL_SignonReply

        Case svc_killedmonster
            cl.stats(STAT_MONSTERS) = cl.stats(STAT_MONSTERS) + 1

        Case svc_foundsecret
            cl.stats(STAT_SECRETS) = cl.stats(STAT_SECRETS) + 1

        Case svc_updatestat
            i = MSG_ReadByte()
            If i < 0 Or i >= MAX_CL_STATS Then Sys_Error "svc_updatestat: %i is invalid", i
            cl.stats(i) = MSG_ReadLong()

        Case svc_spawnstaticsound
            CL_ParseStaticSound

        Case svc_cdtrack
            cl.cdtrack = MSG_ReadByte()
            cl.looptrack = MSG_ReadByte()
            If (cls.demoplayback Or cls.demorecording) And cls.forcetrack <> -1 Then
                CDAudio_Play cls.forcetrack And 255, True
            Else
                CDAudio_Play cl.cdtrack And 255, True
            End If

        Case svc_intermission
            cl.intermission = 1
            cl.completed_time = ToInt(cl.time)
            vid.recalc_refdef = 1        ' go to full screen

        Case svc_finale
            cl.intermission = 2
            cl.completed_time = ToInt(cl.time)
            vid.recalc_refdef = 1        ' go to full screen
            SCR_CenterPrint MSG_ReadString()

        Case svc_cutscene
            cl.intermission = 3
            cl.completed_time = ToInt(cl.time)
            vid.recalc_refdef = 1        ' go to full screen
            SCR_CenterPrint MSG_ReadString()

        Case svc_sellscreen
            Cmd_ExecuteString "help", src_command

        Case Else
            Host_Error "CL_ParseServerMessage: Illegible server message\n"
        End Select
nextcmd:
    Loop
End Sub
