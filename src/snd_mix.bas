Attribute VB_Name = "snd_mix"
' snd_mix.bas - portable code to mix sounds for snd_dma.c (snd_mix.c)
'
' paintbuffer is a flat Long array: sample pair i is (2i, 2i+1), so C's int
' pointer snd_p is an index there. The DMA buffer is snd_win.bas's
' dma_buffer() (shm.buffer is its address); snd_out is an index there.
' As in C, a channel paints from paintbuffer[0] even when it restarts its
' loop in the middle of a block.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const PAINTBUFFER_SIZE As Long = 512
Private paintbuffer(0 To PAINTBUFFER_SIZE * 2 - 1) As Long   ' C: portable_samplepair_t[PAINTBUFFER_SIZE]
Private snd_scaletable(0 To 31, 0 To 255) As Long
Private snd_p As Long, snd_linear_count As Long, snd_vol As Long
Private snd_out As Long

Private Sub Snd_WriteLinearBlastStereo16()
    Dim i As Long, val As Long

    For i = 0 To snd_linear_count - 1 Step 2
        val = SAR(paintbuffer(snd_p + i) * snd_vol, 8)
        If val > &H7FFF& Then
            dma_buffer(snd_out + i) = &H7FFF
        ElseIf val < -32768 Then
            dma_buffer(snd_out + i) = -32768
        Else
            dma_buffer(snd_out + i) = val
        End If

        val = SAR(paintbuffer(snd_p + i + 1) * snd_vol, 8)
        If val > &H7FFF& Then
            dma_buffer(snd_out + i + 1) = &H7FFF
        ElseIf val < -32768 Then
            dma_buffer(snd_out + i + 1) = -32768
        Else
            dma_buffer(snd_out + i + 1) = val
        End If
    Next
End Sub

Private Sub S_TransferStereo16(ByVal endtime As Long)
    Dim lpos As Long, lpaintedtime As Long

    snd_vol = ToInt(cvars(cv_volume).value * 256)

    snd_p = 0
    lpaintedtime = paintedtime

    ' (C: pbuf = shm->buffer, or the locked DirectSound buffer)

    Do While lpaintedtime < endtime
        ' handle recirculating buffer issues
        lpos = lpaintedtime And ((shm.samples \ 2) - 1)

        snd_out = lpos * 2

        snd_linear_count = (shm.samples \ 2) - lpos
        If lpaintedtime + snd_linear_count > endtime Then snd_linear_count = endtime - lpaintedtime

        snd_linear_count = snd_linear_count * 2

        ' write a linear blast of samples
        Snd_WriteLinearBlastStereo16

        snd_p = snd_p + snd_linear_count
        lpaintedtime = lpaintedtime + (snd_linear_count \ 2)
    Loop
End Sub

Private Sub S_TransferPaintBuffer(ByVal endtime As Long)
    Dim out_idx As Long, count As Long, out_mask As Long, p As Long, step_ As Long, val As Long, vol As Long

    If shm.samplebits = 16 And shm.channels = 2 Then
        S_TransferStereo16 endtime
        Exit Sub
    End If

    p = 0
    count = (endtime - paintedtime) * shm.channels
    out_mask = shm.samples - 1
    out_idx = (paintedtime * shm.channels) And out_mask
    step_ = 3 - shm.channels
    vol = ToInt(cvars(cv_volume).value * 256)

    If shm.samplebits = 16 Then
        Do While count > 0
            count = count - 1
            val = SAR(paintbuffer(p) * vol, 8)
            p = p + step_
            If val > &H7FFF& Then
                val = &H7FFF&
            ElseIf val < -32768 Then
                val = -32768
            End If
            dma_buffer(out_idx) = val
            out_idx = (out_idx + 1) And out_mask
        Loop
    ElseIf shm.samplebits = 8 Then
        Do While count > 0
            count = count - 1
            val = SAR(paintbuffer(p) * vol, 8)
            p = p + step_
            If val > &H7FFF& Then
                val = &H7FFF&
            ElseIf val < -32768 Then
                val = -32768
            End If
            ' (C: out[out_idx] = (val>>8) + 128, a byte of the buffer)
            DmaPutByte out_idx, SAR(val, 8) + 128
            out_idx = (out_idx + 1) And out_mask
        Loop
    End If
End Sub

'===============================================================================
' CHANNEL MIXING

Public Sub S_PaintChannels(ByVal endtime As Long)
    Dim i As Long, end_ As Long, ch As Long, sc As Long, ltime As Long, count As Long

    Do While paintedtime < endtime
        ' if paintbuffer is smaller than DMA buffer
        end_ = endtime
        If endtime - paintedtime > PAINTBUFFER_SIZE Then end_ = paintedtime + PAINTBUFFER_SIZE

        ' clear the paint buffer
        ZeroMemory paintbuffer(0), (end_ - paintedtime) * 8

        ' paint in the channels.
        For ch = 0 To total_channels - 1
            If channels(ch).sfx = 0 Then GoTo nextch
            If channels(ch).leftvol = 0 And channels(ch).rightvol = 0 Then GoTo nextch
            sc = S_LoadSound(channels(ch).sfx)
            If sc = 0 Then GoTo nextch

            ltime = paintedtime

            Do While ltime < end_
                ' paint up to end
                If channels(ch).end_ < end_ Then
                    count = channels(ch).end_ - ltime
                Else
                    count = end_ - ltime
                End If

                If count > 0 Then
                    If sfxcaches(sc).width = 1 Then
                        SND_PaintChannelFrom8 ch, sc, count
                    Else
                        SND_PaintChannelFrom16 ch, sc, count
                    End If

                    ltime = ltime + count
                End If

                ' if at end of loop, restart
                If ltime >= channels(ch).end_ Then
                    If sfxcaches(sc).loopstart >= 0 Then
                        channels(ch).pos = sfxcaches(sc).loopstart
                        channels(ch).end_ = ltime + sfxcaches(sc).length - channels(ch).pos
                    Else
                        ' channel just stopped
                        channels(ch).sfx = 0
                        Exit Do
                    End If
                End If
            Loop
nextch:
        Next

        ' transfer out according to DMA format
        S_TransferPaintBuffer end_
        paintedtime = end_
    Loop
End Sub

Public Sub SND_InitScaletable()
    Dim i As Long, j As Long, sj As Long

    For i = 0 To 31
        For j = 0 To 255
            sj = j
            If sj > 127 Then sj = sj - 256   ' (signed char)j
            snd_scaletable(i, j) = sj * i * 8
        Next
    Next
End Sub

Private Sub SND_PaintChannelFrom8(ByVal ch As Long, ByVal sc As Long, ByVal count As Long)
    Dim data As Long, lscale As Long, rscale As Long, i As Long, pos As Long

    If channels(ch).leftvol > 255 Then channels(ch).leftvol = 255
    If channels(ch).rightvol > 255 Then channels(ch).rightvol = 255

    lscale = channels(ch).leftvol \ 8    ' (C: the snd_scaletable rows)
    rscale = channels(ch).rightvol \ 8
    pos = channels(ch).pos

    For i = 0 To count - 1
        data = sfxcaches(sc).data8(pos + i)
        paintbuffer(i * 2) = paintbuffer(i * 2) + snd_scaletable(lscale, data)
        paintbuffer(i * 2 + 1) = paintbuffer(i * 2 + 1) + snd_scaletable(rscale, data)
    Next

    channels(ch).pos = channels(ch).pos + count
End Sub

Private Sub SND_PaintChannelFrom16(ByVal ch As Long, ByVal sc As Long, ByVal count As Long)
    Dim data As Long, left_ As Long, right_ As Long, leftvol As Long, rightvol As Long, i As Long, pos As Long

    leftvol = channels(ch).leftvol
    rightvol = channels(ch).rightvol
    pos = channels(ch).pos

    For i = 0 To count - 1
        data = sfxcaches(sc).data16(pos + i)
        left_ = SAR(data * leftvol, 8)
        right_ = SAR(data * rightvol, 8)
        paintbuffer(i * 2) = paintbuffer(i * 2) + left_
        paintbuffer(i * 2 + 1) = paintbuffer(i * 2 + 1) + right_
    Next

    channels(ch).pos = channels(ch).pos + count
End Sub
