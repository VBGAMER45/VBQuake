Attribute VB_Name = "snd_dma"
' snd_dma.bas - main control for any streaming sound output device (snd_dma.c)
'
' An sfx_t pointer is a known_sfx() index, 1-based, 0 = NULL; a channel_t
' pointer a channels() index, -1 = NULL. shm is C's sn, and shm_valid says
' whether C's shm points at it. There is no cache allocator: a loaded sound
' stays loaded, so Cache_Check is "sfx.cache <> 0". Not ported: -simsound
' (fakedma), which WinQuake's snd_win.c has no DMA position for, and the
' memsize check that forces loadas8bit on small heaps (there is no hunk).
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' =======================================================================
' Internal sound data & structures
' =======================================================================

Public channels(0 To MAX_CHANNELS - 1) As channel_t
Public total_channels As Long

Public snd_blocked As Long
Private snd_ambient As Boolean           ' = 1 (S_Init)
Public snd_initialized As Boolean

' pointer should go away
Public shm As dma_t                      ' C: sn
Public shm_valid As Boolean              ' C: shm != NULL

Public listener_origin As vec3_t
Public listener_forward As vec3_t
Public listener_right As vec3_t
Public listener_up As vec3_t
Public sound_nominal_clip_dist As Single ' = 1000.0 (S_Init)

Private soundtime As Long                ' sample PAIRS
Public paintedtime As Long               ' sample PAIRS

Private Const MAX_SFX As Long = 512
Public known_sfx(1 To MAX_SFX) As sfx_t  ' C: hunk allocated [MAX_SFX]
Public sfxcaches(1 To MAX_SFX) As sfxcache_t
Private num_sfx As Long

Private ambient_sfx(0 To NUM_AMBIENTS - 1) As Long

Private sound_started As Long

Public cv_bgmvolume As Long
Public cv_volume As Long

Private cv_nosound As Long
Private cv_precache As Long
Public cv_loadas8bit As Long
Private cv_bgmbuffer As Long
Private cv_ambient_level As Long
Private cv_ambient_fade As Long
Private cv_snd_noextraupdate As Long
Private cv_snd_show As Long
Private cv__snd_mixahead As Long

Private play_hash As Long                ' C: static of S_Play
Private playvol_hash As Long             ' C: static of S_PlayVol

' GetSoundtime's statics
Private gst_buffers As Long
Private gst_oldsamplepos As Long

Public Sub S_AmbientOff()
    snd_ambient = False
End Sub

Public Sub S_AmbientOn()
    snd_ambient = True
End Sub

Private Sub S_SoundInfo_f()
    If sound_started = 0 Or Not shm_valid Then
        Con_Printf "sound system not started\n"
        Exit Sub
    End If

    Con_Printf "%5d stereo\n", shm.channels - 1
    Con_Printf "%5d samples\n", shm.samples
    Con_Printf "%5d samplepos\n", shm.samplepos
    Con_Printf "%5d samplebits\n", shm.samplebits
    Con_Printf "%5d submission_chunk\n", shm.submission_chunk
    Con_Printf "%5d speed\n", shm.speed
    Con_Printf "0x%x dma buffer\n", shm.buffer
    Con_Printf "%5d total_channels\n", total_channels
End Sub

Public Sub S_Startup()
    Dim rc As Long

    If Not snd_initialized Then Exit Sub

    rc = SNDDMA_Init()

    If rc = 0 Then
        sound_started = 0
        Exit Sub
    End If

    sound_started = 1
End Sub

Public Sub S_Init()
    Con_Printf "\nSound Initialization\n"

    ' (C: static initializers)
    snd_ambient = True
    sound_nominal_clip_dist = 1000!
    play_hash = 345
    playvol_hash = 543

    If COM_CheckParm("-nosound") Then Exit Sub

    Cmd_AddCommand "play", XC_S_Play
    Cmd_AddCommand "playvol", XC_S_PlayVol
    Cmd_AddCommand "stopsound", XC_S_StopAllSoundsC
    Cmd_AddCommand "soundlist", XC_S_SoundList
    Cmd_AddCommand "soundinfo", XC_S_SoundInfo_f

    cv_nosound = Cvar_RegisterVariable("nosound", "0")
    cv_volume = Cvar_RegisterVariable("volume", "0.7", True)
    cv_precache = Cvar_RegisterVariable("precache", "1")
    cv_loadas8bit = Cvar_RegisterVariable("loadas8bit", "0")
    cv_bgmvolume = Cvar_RegisterVariable("bgmvolume", "1", True)
    cv_bgmbuffer = Cvar_RegisterVariable("bgmbuffer", "4096")
    cv_ambient_level = Cvar_RegisterVariable("ambient_level", "0.3")
    cv_ambient_fade = Cvar_RegisterVariable("ambient_fade", "100")
    cv_snd_noextraupdate = Cvar_RegisterVariable("snd_noextraupdate", "0")
    cv_snd_show = Cvar_RegisterVariable("snd_show", "0")
    cv__snd_mixahead = Cvar_RegisterVariable("_snd_mixahead", "0.1", True)

    ' (C forces loadas8bit with less than 8 MB of heap)

    snd_initialized = True

    S_Startup

    SND_InitScaletable

    num_sfx = 0

    ' (C prints shm->speed whether or not a device started; shm is NULL then)
    If shm_valid Then Con_Printf "Sound sampling rate: %i\n", shm.speed

    ' provides a tick sound until washed clean

    '   if (shm->buffer)
    '       shm->buffer[4] = shm->buffer[5] = 0x7f;  // force a pop for debugging

    ambient_sfx(AMBIENT_WATER) = S_PrecacheSound("ambience/water1.wav")
    ambient_sfx(AMBIENT_SKY) = S_PrecacheSound("ambience/wind2.wav")

    S_StopAllSounds True
End Sub

' =======================================================================
' Shutdown sound engine
' =======================================================================

Public Sub S_Shutdown()
    If sound_started = 0 Then Exit Sub

    If shm_valid Then shm.gamealive = False

    shm_valid = False
    sound_started = 0

    SNDDMA_Shutdown
End Sub

' =======================================================================
' Load a sound
' =======================================================================

Private Function S_FindName(ByVal name As String) As Long
    Dim i As Long

    If Len(name) = 0 Then Sys_Error "S_FindName: NULL\n"

    If Len(name) >= MAX_QPATH Then Sys_Error "Sound name too long: %s", name

    ' see if already loaded
    For i = 1 To num_sfx
        If known_sfx(i).name = name Then
            S_FindName = i
            Exit Function
        End If
    Next

    If num_sfx = MAX_SFX Then Sys_Error "S_FindName: out of sfx_t"

    known_sfx(i).name = name

    num_sfx = num_sfx + 1

    S_FindName = i
End Function

Public Sub S_TouchSound(ByVal name As String)
    Dim sfx As Long

    If sound_started = 0 Then Exit Sub

    sfx = S_FindName(name)
    ' (C: Cache_Check (&sfx->cache); nothing is ever flushed here)
End Sub

Public Function S_PrecacheSound(ByVal name As String) As Long
    Dim sfx As Long

    If sound_started = 0 Or cvars(cv_nosound).value <> 0 Then Exit Function

    sfx = S_FindName(name)

    ' cache it in
    If cvars(cv_precache).value Then S_LoadSound sfx

    S_PrecacheSound = sfx
End Function

'=============================================================================

Private Function SND_PickChannel(ByVal entnum As Long, ByVal entchannel As Long) As Long
    Dim ch_idx As Long, first_to_die As Long, life_left As Long

    ' Check for replacement sound, or find the best one to replace
    first_to_die = -1
    life_left = &H7FFFFFFF
    For ch_idx = NUM_AMBIENTS To NUM_AMBIENTS + MAX_DYNAMIC_CHANNELS - 1
        If entchannel <> 0 And channels(ch_idx).entnum = entnum And _
           (channels(ch_idx).entchannel = entchannel Or entchannel = -1) Then
            ' allways override sound from same entity
            first_to_die = ch_idx
            Exit For
        End If

        ' don't let monster sounds override player sounds
        If channels(ch_idx).entnum = cl.viewentity And entnum <> cl.viewentity And channels(ch_idx).sfx <> 0 Then GoTo nextch

        If channels(ch_idx).end_ - paintedtime < life_left Then
            life_left = channels(ch_idx).end_ - paintedtime
            first_to_die = ch_idx
        End If
nextch:
    Next

    If first_to_die = -1 Then
        SND_PickChannel = -1
        Exit Function
    End If

    If channels(first_to_die).sfx Then channels(first_to_die).sfx = 0

    SND_PickChannel = first_to_die
End Function

Private Sub SND_Spatialize(ByVal ch As Long)
    Dim dot As Single, dist As Single, lscale As Single, rscale As Single, scale_ As Single
    Dim source_vec As vec3_t

    ' anything coming from the view entity will allways be full volume
    If channels(ch).entnum = cl.viewentity Then
        channels(ch).leftvol = channels(ch).master_vol
        channels(ch).rightvol = channels(ch).master_vol
        Exit Sub
    End If

    ' calculate stereo seperation and distance attenuation

    VectorSubtract channels(ch).origin, listener_origin, source_vec

    dist = VectorNormalize(source_vec) * channels(ch).dist_mult

    dot = DotProduct(listener_right, source_vec)

    If shm.channels = 1 Then
        rscale = 1!
        lscale = 1!
    Else
        rscale = 1! + dot
        lscale = 1! - dot
    End If

    ' add in distance effect
    scale_ = (1! - dist) * rscale
    channels(ch).rightvol = ToInt(channels(ch).master_vol * scale_)
    If channels(ch).rightvol < 0 Then channels(ch).rightvol = 0

    scale_ = (1! - dist) * lscale
    channels(ch).leftvol = ToInt(channels(ch).master_vol * scale_)
    If channels(ch).leftvol < 0 Then channels(ch).leftvol = 0
End Sub

' =======================================================================
' Start a sound effect
' =======================================================================

Public Sub S_StartSound(ByVal entnum As Long, ByVal entchannel As Long, ByVal sfx As Long, origin As vec3_t, _
        ByVal fvol As Single, ByVal attenuation As Single)
    Dim target_chan As Long, check As Long, sc As Long, vol As Long, skip As Long, blank As channel_t

    If sound_started = 0 Then Exit Sub

    If sfx = 0 Then Exit Sub

    If cvars(cv_nosound).value Then Exit Sub

    vol = ToInt(fvol * 255)

    ' pick a channel to play on
    target_chan = SND_PickChannel(entnum, entchannel)
    If target_chan = -1 Then Exit Sub

    ' spatialize
    channels(target_chan) = blank
    channels(target_chan).origin = origin
    channels(target_chan).dist_mult = attenuation / sound_nominal_clip_dist
    channels(target_chan).master_vol = vol
    channels(target_chan).entnum = entnum
    channels(target_chan).entchannel = entchannel
    SND_Spatialize target_chan

    If channels(target_chan).leftvol = 0 And channels(target_chan).rightvol = 0 Then Exit Sub   ' not audible at all

    ' new channel
    sc = S_LoadSound(sfx)
    If sc = 0 Then
        channels(target_chan).sfx = 0
        Exit Sub                         ' couldn't load the sound's data
    End If

    channels(target_chan).sfx = sfx
    channels(target_chan).pos = 0
    channels(target_chan).end_ = paintedtime + sfxcaches(sc).length

    ' if an identical sound has also been started this frame, offset the pos
    ' a bit to keep it from just making the first one louder
    For check = NUM_AMBIENTS To NUM_AMBIENTS + MAX_DYNAMIC_CHANNELS - 1
        If check = target_chan Then GoTo nextcheck
        If channels(check).sfx = sfx And channels(check).pos = 0 Then
            skip = rand() Mod ToInt(0.1 * shm.speed)
            If skip >= channels(target_chan).end_ Then skip = channels(target_chan).end_ - 1
            channels(target_chan).pos = channels(target_chan).pos + skip
            channels(target_chan).end_ = channels(target_chan).end_ - skip
            Exit For
        End If
nextcheck:
    Next
End Sub

Public Sub S_StopSound(ByVal entnum As Long, ByVal entchannel As Long)
    Dim i As Long

    For i = 0 To MAX_DYNAMIC_CHANNELS - 1
        If channels(i).entnum = entnum And channels(i).entchannel = entchannel Then
            channels(i).end_ = 0
            channels(i).sfx = 0
            Exit Sub
        End If
    Next
End Sub

Public Sub S_StopAllSounds(ByVal clear As Boolean)
    Dim i As Long, blank As channel_t

    If sound_started = 0 Then Exit Sub

    total_channels = MAX_DYNAMIC_CHANNELS + NUM_AMBIENTS   ' no statics

    For i = 0 To MAX_CHANNELS - 1
        channels(i) = blank
    Next

    If clear Then S_ClearBuffer
End Sub

Private Sub S_StopAllSoundsC()
    S_StopAllSounds True
End Sub

Public Sub S_ClearBuffer()
    Dim clear As Byte

    If sound_started = 0 Or Not shm_valid Then Exit Sub
    If shm.buffer = 0 Then Exit Sub

    If shm.samplebits = 8 Then
        clear = &H80
    Else
        clear = 0
    End If

    FillMemory ByVal shm.buffer, shm.samples * shm.samplebits \ 8, clear
End Sub

Public Sub S_StaticSound(ByVal sfx As Long, origin As vec3_t, ByVal vol As Single, ByVal attenuation As Single)
    Dim ss As Long, sc As Long

    If sfx = 0 Then Exit Sub

    If total_channels = MAX_CHANNELS Then
        Con_Printf "total_channels == MAX_CHANNELS\n"
        Exit Sub
    End If

    ss = total_channels
    total_channels = total_channels + 1

    sc = S_LoadSound(sfx)
    If sc = 0 Then Exit Sub

    If sfxcaches(sc).loopstart = -1 Then
        Con_Printf "Sound %s not looped\n", known_sfx(sfx).name
        Exit Sub
    End If

    channels(ss).sfx = sfx
    channels(ss).origin = origin
    channels(ss).master_vol = ToInt(vol)
    channels(ss).dist_mult = (attenuation / 64) / sound_nominal_clip_dist
    channels(ss).end_ = paintedtime + sfxcaches(sc).length

    SND_Spatialize ss
End Sub

'=============================================================================

Private Sub S_UpdateAmbientSounds()
    Dim l As Long, vol As Single, ambient_channel As Long

    If Not snd_ambient Then Exit Sub

    ' calc ambient sound levels
    If cl.worldmodel = 0 Then Exit Sub

    l = Mod_PointInLeaf(listener_origin, cl.worldmodel)
    ' (C also tests !l; Mod_PointInLeaf always returns a leaf)
    If cvars(cv_ambient_level).value = 0 Then
        For ambient_channel = 0 To NUM_AMBIENTS - 1
            channels(ambient_channel).sfx = 0
        Next
        Exit Sub
    End If

    For ambient_channel = 0 To NUM_AMBIENTS - 1
        channels(ambient_channel).sfx = ambient_sfx(ambient_channel)

        vol = cvars(cv_ambient_level).value * mod_nodes(l).ambient_sound_level(ambient_channel)
        If vol < 8 Then vol = 0

        ' don't adjust volume too fast
        If channels(ambient_channel).master_vol < vol Then
            channels(ambient_channel).master_vol = ToInt(channels(ambient_channel).master_vol + host_frametime * cvars(cv_ambient_fade).value)
            If channels(ambient_channel).master_vol > vol Then channels(ambient_channel).master_vol = ToInt(vol)
        ElseIf channels(ambient_channel).master_vol > vol Then
            channels(ambient_channel).master_vol = ToInt(channels(ambient_channel).master_vol - host_frametime * cvars(cv_ambient_fade).value)
            If channels(ambient_channel).master_vol < vol Then channels(ambient_channel).master_vol = ToInt(vol)
        End If

        channels(ambient_channel).leftvol = channels(ambient_channel).master_vol
        channels(ambient_channel).rightvol = channels(ambient_channel).master_vol
    Next
End Sub

' Called once each time through the main loop
Public Sub S_Update(origin As vec3_t, forward As vec3_t, right_ As vec3_t, up As vec3_t)
    Dim i As Long, j As Long, total As Long, ch As Long, combine As Long

    If sound_started = 0 Or (snd_blocked > 0) Then Exit Sub

    listener_origin = origin
    listener_forward = forward
    listener_right = right_
    listener_up = up

    ' update general area ambient sound sources
    S_UpdateAmbientSounds

    combine = -1

    ' update spatialization for static and dynamic sounds
    For i = NUM_AMBIENTS To total_channels - 1
        ch = i
        If channels(ch).sfx = 0 Then GoTo nexti
        SND_Spatialize ch                ' respatialize channel
        If channels(ch).leftvol = 0 And channels(ch).rightvol = 0 Then GoTo nexti

        ' try to combine static sounds with a previous channel of the same
        ' sound effect so we don't mix five torches every frame

        If i >= MAX_DYNAMIC_CHANNELS + NUM_AMBIENTS Then
            ' see if it can just use the last one
            If combine <> -1 Then
                If channels(combine).sfx = channels(ch).sfx Then
                    channels(combine).leftvol = channels(combine).leftvol + channels(ch).leftvol
                    channels(combine).rightvol = channels(combine).rightvol + channels(ch).rightvol
                    channels(ch).leftvol = 0
                    channels(ch).rightvol = 0
                    GoTo nexti
                End If
            End If
            ' search for one
            j = MAX_DYNAMIC_CHANNELS + NUM_AMBIENTS
            Do While j < i
                If channels(j).sfx = channels(ch).sfx Then Exit Do
                j = j + 1
            Loop
            combine = j

            If j = total_channels Then
                combine = -1
            Else
                If combine <> ch Then
                    channels(combine).leftvol = channels(combine).leftvol + channels(ch).leftvol
                    channels(combine).rightvol = channels(combine).rightvol + channels(ch).rightvol
                    channels(ch).leftvol = 0
                    channels(ch).rightvol = 0
                End If
                GoTo nexti
            End If
        End If
nexti:
    Next

    '
    ' debugging output
    '
    If cvars(cv_snd_show).value Then
        total = 0
        For i = 0 To total_channels - 1
            If channels(i).sfx <> 0 And (channels(i).leftvol <> 0 Or channels(i).rightvol <> 0) Then
                'Con_Printf ("%3i %3i %s\n", ch->leftvol, ch->rightvol, ch->sfx->name);
                total = total + 1
            End If
        Next

        Con_Printf "----(%i)----\n", total
    End If

    ' mix some sound
    S_Update_
End Sub

Private Sub GetSoundtime()
    Dim samplepos As Long, fullsamples As Long

    fullsamples = shm.samples \ shm.channels

    ' it is possible to miscount buffers if it has wrapped twice between
    ' calls to S_Update.  Oh well.
    samplepos = SNDDMA_GetDMAPos()

    If samplepos < gst_oldsamplepos Then
        gst_buffers = gst_buffers + 1    ' buffer wrapped

        If paintedtime > &H40000000 Then
            ' time to chop things off to avoid 32 bit limits
            gst_buffers = 0
            paintedtime = fullsamples
            S_StopAllSounds True
        End If
    End If
    gst_oldsamplepos = samplepos

    soundtime = gst_buffers * fullsamples + samplepos \ shm.channels
End Sub

Public Sub S_ExtraUpdate()
    IN_Accumulate                        ' (_WIN32)

    If cvars(cv_snd_noextraupdate).value Then Exit Sub   ' don't pollute timings
    S_Update_
End Sub

Private Sub S_Update_()
    Dim endtime As Long, samps As Long

    If sound_started = 0 Or (snd_blocked > 0) Then Exit Sub

    ' Updates DMA time
    GetSoundtime

    ' check to make sure that we haven't overshot
    If paintedtime < soundtime Then
        'Con_Printf ("S_Update_ : overflow\n");
        paintedtime = soundtime
    End If

    ' mix ahead of current position
    endtime = ToInt(soundtime + cvars(cv__snd_mixahead).value * shm.speed)
    samps = SAR(shm.samples, shm.channels - 1)
    If endtime - soundtime > samps Then endtime = soundtime + samps

    ' (the DirectSound buffer restore of _WIN32: there is no DirectSound)

    S_PaintChannels endtime

    SNDDMA_Submit
End Sub

'===============================================================================
' console functions

' (QuakeVB) C: if (!Q_strrchr(name, '.')) strcat (name, ".wav")
Private Function WavName(ByVal name As String) As String
    If InStrRev(name, ".") = 0 Then
        WavName = name & ".wav"
    Else
        WavName = name
    End If
End Function

Private Sub S_Play()
    Dim i As Long, sfx As Long

    i = 1
    Do While i < Cmd_Argc()
        sfx = S_PrecacheSound(WavName(Cmd_Argv(i)))
        S_StartSound play_hash, 0, sfx, listener_origin, 1!, 1!
        play_hash = play_hash + 1
        i = i + 1
    Loop
End Sub

Private Sub S_PlayVol()
    Dim i As Long, vol As Single, sfx As Long

    i = 1
    Do While i < Cmd_Argc()
        sfx = S_PrecacheSound(WavName(Cmd_Argv(i)))
        vol = Q_atof(Cmd_Argv(i + 1))
        S_StartSound playvol_hash, 0, sfx, listener_origin, vol, 1!
        playvol_hash = playvol_hash + 1
        i = i + 2
    Loop
End Sub

Private Sub S_SoundList()
    Dim i As Long, sc As Long, size As Long, total As Long

    total = 0
    For i = 1 To num_sfx
        sc = known_sfx(i).cache          ' C: Cache_Check (&sfx->cache)
        If sc = 0 Then GoTo nexti
        size = sfxcaches(sc).length * sfxcaches(sc).width * (sfxcaches(sc).stereo + 1)
        total = total + size
        If sfxcaches(sc).loopstart >= 0 Then
            Con_Printf "L"
        Else
            Con_Printf " "
        End If
        Con_Printf "(%2db) %6i : %s\n", sfxcaches(sc).width * 8, size, known_sfx(i).name
nexti:
    Next
    Con_Printf "Total resident: %i\n", total
End Sub

' (QuakeVB) the console commands of this module, for Cmd_Call
Public Sub S_CallCommand(ByVal fn As Long)
    Select Case fn
    Case XC_S_Play: S_Play
    Case XC_S_PlayVol: S_PlayVol
    Case XC_S_StopAllSoundsC: S_StopAllSoundsC
    Case XC_S_SoundList: S_SoundList
    Case XC_S_SoundInfo_f: S_SoundInfo_f
    End Select
End Sub

Public Sub S_LocalSound(ByVal sound As String)
    Dim sfx As Long

    If cvars(cv_nosound).value Then Exit Sub
    If sound_started = 0 Then Exit Sub

    sfx = S_PrecacheSound(sound)
    If sfx = 0 Then
        Con_Printf "S_LocalSound: can't cache %s\n", sound
        Exit Sub
    End If
    S_StartSound cl.viewentity, -1, sfx, vec3_origin, 1!, 1!
End Sub

Public Sub S_ClearPrecache()
End Sub

Public Sub S_BeginPrecaching()
End Sub

Public Sub S_EndPrecaching()
End Sub
