Attribute VB_Name = "snd_mem"
' snd_mem.bas - sound caching (snd_mem.c)
'
' A sound's cache is sfxcaches(sfx) (snd_dma.bas), loaded once and kept.
' The WAV parser's byte pointers (data_p, iff_end, last_chunk, iff_data) are
' offsets into the file's bytes, which the helpers take as a parameter;
' data_p = -1 is C's NULL. Reads past the end of the file give 0 (C reads
' whatever follows the buffer).
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private data_p As Long
Private iff_end As Long
Private last_chunk As Long
Private iff_data As Long
Private iff_chunk_len As Long

' (QuakeVB) the byte at p, 0 past the end
Private Function ByteAt(wav() As Byte, ByVal p As Long) As Long
    If p >= 0 And p <= UBound(wav) Then ByteAt = wav(p)
End Function

' data: the file's bytes, from ofs
Private Sub ResampleSfx(ByVal sfx As Long, ByVal inrate As Long, ByVal inwidth As Long, data() As Byte, ByVal ofs As Long)
    Dim outcount As Long, srcsample As Long, stepscale As Single, i As Long
    Dim sample As Long, samplefrac As Long, fracstep As Long, sc As Long

    sc = known_sfx(sfx).cache            ' C: Cache_Check (&sfx->cache)
    If sc = 0 Then Exit Sub

    stepscale = CSng(inrate) / shm.speed   ' this is usually 0.5, 1, or 2

    outcount = ToInt(sfxcaches(sc).length / stepscale)
    sfxcaches(sc).length = outcount
    If sfxcaches(sc).loopstart <> -1 Then sfxcaches(sc).loopstart = ToInt(sfxcaches(sc).loopstart / stepscale)

    sfxcaches(sc).speed = shm.speed
    If cvars(cv_loadas8bit).value Then
        sfxcaches(sc).width = 1
    Else
        sfxcaches(sc).width = inwidth
    End If
    sfxcaches(sc).stereo = 0

    ' (C allocated the data with the cache entry in S_LoadSound)
    If sfxcaches(sc).width = 2 Then
        If outcount > 0 Then ReDim sfxcaches(sc).data16(0 To outcount - 1)
    Else
        If outcount > 0 Then ReDim sfxcaches(sc).data8(0 To outcount - 1)
    End If

    ' resample / decimate to the current source rate

    If stepscale = 1 And inwidth = 1 And sfxcaches(sc).width = 1 Then
        ' fast special case
        For i = 0 To outcount - 1
            sfxcaches(sc).data8(i) = ByteAt(data, ofs + i) Xor &H80   ' (signed char)(data[i] - 128)
        Next
    Else
        ' general case
        samplefrac = 0
        fracstep = ToInt(stepscale * 256)
        For i = 0 To outcount - 1
            srcsample = samplefrac \ 256     ' (samplefrac >= 0)
            samplefrac = samplefrac + fracstep
            If inwidth = 2 Then
                sample = ByteAt(data, ofs + srcsample * 2) + ByteAt(data, ofs + srcsample * 2 + 1) * 256&
                If sample > 32767 Then sample = sample - 65536   ' LittleShort
            Else
                sample = (ByteAt(data, ofs + srcsample) - 128) * 256&
            End If
            If sfxcaches(sc).width = 2 Then
                sfxcaches(sc).data16(i) = sample
            Else
                sfxcaches(sc).data8(i) = SAR(sample, 8) And &HFF&   ' (signed char)(sample >> 8)
            End If
        Next
    End If
End Sub

'=============================================================================

' returns the sfxcaches() index, 0 if the sound cannot be loaded
Public Function S_LoadSound(ByVal s As Long) As Long
    Dim namebuffer As String, data() As Byte, info As wavinfo_t, sc As Long, blank As sfxcache_t

    ' see if still in memory
    sc = known_sfx(s).cache
    If sc Then
        S_LoadSound = sc
        Exit Function
    End If

    'Con_Printf ("S_LoadSound: %x\n", (int)stackbuf);
    ' load it in
    namebuffer = "sound/" & known_sfx(s).name

    '   Con_Printf ("loading %s\n",namebuffer);

    If Not COM_LoadFile(namebuffer, data) Then
        Con_Printf "Couldn't load %s\n", namebuffer
        Exit Function
    End If

    info = GetWavinfo(known_sfx(s).name, data, com_filesize)
    If info.channels <> 1 Then
        Con_Printf "%s is a stereo sample\n", known_sfx(s).name
        Exit Function
    End If

    ' (C: Cache_Alloc (&s->cache, len + sizeof(sfxcache_t), s->name), with
    ' len the resampled size; ResampleSfx allocates the samples here)
    sc = s
    sfxcaches(sc) = blank
    known_sfx(s).cache = sc

    sfxcaches(sc).length = info.samples
    sfxcaches(sc).loopstart = info.loopstart
    sfxcaches(sc).speed = info.rate
    sfxcaches(sc).width = info.width
    sfxcaches(sc).stereo = info.channels

    ResampleSfx s, sfxcaches(sc).speed, sfxcaches(sc).width, data, info.dataofs

    S_LoadSound = sc
End Function

'===============================================================================
' WAV loading

Private Function GetLittleShort(wav() As Byte) As Long
    Dim val As Long

    val = ByteAt(wav, data_p) + ByteAt(wav, data_p + 1) * 256&
    If val > 32767 Then val = val - 65536    ' (short)
    data_p = data_p + 2
    GetLittleShort = val
End Function

Private Function GetLittleLong(wav() As Byte) As Long
    Dim val As Long, b3 As Long

    b3 = ByteAt(wav, data_p + 3)
    If b3 And &H80 Then b3 = b3 - 256
    val = ByteAt(wav, data_p) + ByteAt(wav, data_p + 1) * 256& + ByteAt(wav, data_p + 2) * 65536 + b3 * 16777216
    data_p = data_p + 4
    GetLittleLong = val
End Function

' (QuakeVB) C: !Q_strncmp (data_p + ofs, name, 4)
Private Function ChunkIs(wav() As Byte, ByVal p As Long, ByVal name As String) As Boolean
    Dim i As Long

    For i = 0 To 3
        If ByteAt(wav, p + i) <> AscW(Mid$(name, i + 1, 1)) Then Exit Function
    Next
    ChunkIs = True
End Function

Private Sub FindNextChunk(wav() As Byte, ByVal name As String)
    Do
        data_p = last_chunk

        If data_p >= iff_end Then
            ' didn't find the chunk
            data_p = -1
            Exit Sub
        End If

        data_p = data_p + 4
        iff_chunk_len = GetLittleLong(wav)
        If iff_chunk_len < 0 Then
            data_p = -1
            Exit Sub
        End If
        '       if (iff_chunk_len > 1024*1024)
        '           Sys_Error ("FindNextChunk: %i length is past the 1 meg sanity limit", iff_chunk_len);
        data_p = data_p - 8
        last_chunk = data_p + 8 + ((iff_chunk_len + 1) And Not 1)
        If ChunkIs(wav, data_p, name) Then Exit Sub
    Loop
End Sub

Private Sub FindChunk(wav() As Byte, ByVal name As String)
    last_chunk = iff_data
    FindNextChunk wav, name
End Sub

Private Function GetWavinfo(ByVal name As String, wav() As Byte, ByVal wavlength As Long) As wavinfo_t
    Dim info As wavinfo_t, i As Long, format_ As Long, samples As Long

    iff_data = 0
    iff_end = wavlength

    ' find "RIFF" chunk
    FindChunk wav, "RIFF"
    If data_p = -1 Then
        Con_Printf "Missing RIFF/WAVE chunks\n"
        GetWavinfo = info
        Exit Function
    End If
    If Not ChunkIs(wav, data_p + 8, "WAVE") Then
        Con_Printf "Missing RIFF/WAVE chunks\n"
        GetWavinfo = info
        Exit Function
    End If

    ' get "fmt " chunk
    iff_data = data_p + 12
    ' DumpChunks ();

    FindChunk wav, "fmt "
    If data_p = -1 Then
        Con_Printf "Missing fmt chunk\n"
        GetWavinfo = info
        Exit Function
    End If
    data_p = data_p + 8
    format_ = GetLittleShort(wav)
    If format_ <> 1 Then
        Con_Printf "Microsoft PCM format only\n"
        GetWavinfo = info
        Exit Function
    End If

    info.channels = GetLittleShort(wav)
    info.rate = GetLittleLong(wav)
    data_p = data_p + 4 + 2
    info.width = GetLittleShort(wav) \ 8

    ' get cue chunk
    FindChunk wav, "cue "
    If data_p <> -1 Then
        data_p = data_p + 32
        info.loopstart = GetLittleLong(wav)
        '       Con_Printf("loopstart=%d\n", sfx->loopstart);

        ' if the next chunk is a LIST chunk, look for a cue length marker
        FindNextChunk wav, "LIST"
        If data_p <> -1 Then
            If ChunkIs(wav, data_p + 28, "mark") Then
                ' this is not a proper parse, but it works with cooledit...
                data_p = data_p + 24
                i = GetLittleLong(wav)   ' samples in loop
                info.samples = info.loopstart + i
                '               Con_Printf("looped length: %i\n", i);
            End If
        End If
    Else
        info.loopstart = -1
    End If

    ' find data chunk
    FindChunk wav, "data"
    If data_p = -1 Then
        Con_Printf "Missing data chunk\n"
        GetWavinfo = info
        Exit Function
    End If

    data_p = data_p + 4
    samples = GetLittleLong(wav) \ info.width

    If info.samples Then
        If samples < info.samples Then Sys_Error "Sound %s has a bad loop length", name
    Else
        info.samples = samples
    End If

    info.dataofs = data_p

    GetWavinfo = info
End Function
