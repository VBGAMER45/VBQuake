Attribute VB_Name = "snd_win"
' snd_win.bas - the Windows sound device (snd_win.c)
'
' Only WinQuake's waveOut path is ported (C: -wavonly, or when DirectSound
' fails); there is no DirectSound. The waveform memory is dma_buffer() and
' the headers lpWaveHdr(), VB arrays that stay where they are while the
' device is open (C: GlobalAlloc'd blocks). As in C the device is polled:
' SNDDMA_Submit checks WHDR_DONE, there is no callback.
' Test aids (no device is opened): -sndcapture <file.wav> writes what would
' be played to a WAV file, with the blocks completing in real time; under
' -hidden without -sndcapture no device is opened at all.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Type WAVEFORMATEX
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
    cbSize As Integer
End Type

Private Type WAVEHDR
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long
    reserved As Long
End Type

Private Declare Function waveOutOpen Lib "winmm.dll" (lphWaveOut As Long, ByVal uDeviceID As Long, _
    lpFormat As WAVEFORMATEX, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function waveOutPrepareHeader Lib "winmm.dll" (ByVal hwo As Long, ByVal pwh As Long, ByVal cbwh As Long) As Long
Private Declare Function waveOutUnprepareHeader Lib "winmm.dll" (ByVal hwo As Long, ByVal pwh As Long, ByVal cbwh As Long) As Long
Private Declare Function waveOutWrite Lib "winmm.dll" (ByVal hwo As Long, ByVal pwh As Long, ByVal cbwh As Long) As Long
Private Declare Function waveOutReset Lib "winmm.dll" (ByVal hwo As Long) As Long
Private Declare Function waveOutClose Lib "winmm.dll" (ByVal hwo As Long) As Long

Private Const WAVE_MAPPER As Long = -1
Private Const WAVE_FORMAT_PCM As Long = 1
Private Const CALLBACK_NULL As Long = 0
Private Const MMSYSERR_NOERROR As Long = 0
Private Const MMSYSERR_ALLOCATED As Long = 4
Private Const WHDR_DONE As Long = 1

Private Const WAV_BUFFERS As Long = 64
Private Const WAV_MASK As Long = &H3F
Private Const WAV_BUFFER_SIZE As Long = &H400

Private dsound_init As Boolean           ' (never set: there is no DirectSound)
Private wav_init As Boolean
Private snd_firsttime As Boolean         ' = true (SNDDMA_Init)
Private snd_firsttime_set As Boolean
Private snd_isdirect As Boolean, snd_iswave As Boolean

Private sample16 As Long
Private snd_sent As Long, snd_completed As Long

Private hWaveOut As Long
Private lpWaveHdr(0 To WAV_BUFFERS - 1) As WAVEHDR
Public dma_buffer() As Integer           ' C: lpData, the samples (16 bit stereo)
Private dma_made As Boolean
Private gSndBufSize As Long

' -sndcapture
Private capture_file As Long             ' file handle, 0 = not capturing
Private capture_bytes As Long            ' bytes of samples written
Private capture_start As Double          ' realtime of the first block
Private capture_block(0 To WAV_BUFFER_SIZE - 1) As Byte

Public Sub S_BlockSound()
    ' DirectSound takes care of blocking itself
    If snd_iswave Then
        snd_blocked = snd_blocked + 1

        If snd_blocked = 1 Then
            If hWaveOut Then waveOutReset hWaveOut
        End If
    End If
End Sub

Public Sub S_UnblockSound()
    ' DirectSound takes care of blocking itself
    If snd_iswave Then snd_blocked = snd_blocked - 1
End Sub

Private Sub FreeSound()
    Dim i As Long

    If hWaveOut Then
        waveOutReset hWaveOut

        For i = 0 To WAV_BUFFERS - 1
            waveOutUnprepareHeader hWaveOut, VarPtr(lpWaveHdr(i)), LenB(lpWaveHdr(i))
        Next

        waveOutClose hWaveOut
    End If

    If capture_file Then CaptureClose

    hWaveOut = 0
    If dma_made Then
        Erase dma_buffer
        dma_made = False
    End If
    shm.buffer = 0
    dsound_init = False
    wav_init = False
End Sub

' (QuakeVB) -sndcapture: the WAV header, with the sizes patched at the end
Private Sub CaptureWriteHeader(ByVal datasize As Long)
    Dim h(0 To 43) As Byte

    PutLE h, 0, &H46464952               ' "RIFF"
    PutLE h, 4, 36 + datasize
    PutLE h, 8, &H45564157               ' "WAVE"
    PutLE h, 12, &H20746D66              ' "fmt "
    PutLE h, 16, 16
    PutLE h, 20, WAVE_FORMAT_PCM + shm.channels * &H10000
    PutLE h, 24, shm.speed
    PutLE h, 28, shm.speed * shm.channels * (shm.samplebits \ 8)
    PutLE h, 32, shm.channels * (shm.samplebits \ 8) + shm.samplebits * &H10000
    PutLE h, 36, &H61746164              ' "data"
    PutLE h, 40, datasize
    Sys_FileSeek capture_file, 0
    Sys_FileWrite capture_file, h, 0, 44
End Sub

Private Sub PutLE(b() As Byte, ByVal ofs As Long, ByVal v As Long)
    CopyMemory b(ofs), v, 4
End Sub

Private Sub CaptureClose()
    CaptureWriteHeader capture_bytes
    Sys_FileClose capture_file
    capture_file = 0
End Sub

' Crappy windows multimedia base
Private Function SNDDMA_InitWav() As Boolean
    Dim format_ As WAVEFORMATEX, i As Long, hr As Long, p As Long

    snd_sent = 0
    snd_completed = 0

    shm_valid = True                     ' shm = &sn

    shm.channels = 2
    shm.samplebits = 16
    shm.speed = 11025

    format_.wFormatTag = WAVE_FORMAT_PCM
    format_.nChannels = shm.channels
    format_.wBitsPerSample = shm.samplebits
    format_.nSamplesPerSec = shm.speed
    format_.nBlockAlign = format_.nChannels * format_.wBitsPerSample \ 8
    format_.cbSize = 0
    format_.nAvgBytesPerSec = format_.nSamplesPerSec * format_.nBlockAlign

    p = COM_CheckParm("-sndcapture")
    If p <> 0 And p < com_argc - 1 Then
        ' (QuakeVB) no device: the blocks go to the file
        capture_file = Sys_FileCreate(com_argv(p + 1))
        If capture_file = -1 Then
            capture_file = 0
            Con_SafePrintf "-sndcapture: can't create %s\n", com_argv(p + 1)
            Exit Function
        End If
        capture_bytes = 0
        capture_start = realtime
        CaptureWriteHeader 0
    Else
        ' Open a waveform device for output using window callback.
        Do
            hr = waveOutOpen(hWaveOut, WAVE_MAPPER, format_, 0, 0, CALLBACK_NULL)
            If hr = MMSYSERR_NOERROR Then Exit Do

            If hr <> MMSYSERR_ALLOCATED Then
                Con_SafePrintf "waveOutOpen failed\n"
                Exit Function
            End If

            If MsgBox("The sound hardware is in use by another app." & vbLf & vbLf & _
                      "Select Retry to try to start sound again or Cancel to run Quake with no sound.", _
                      vbRetryCancel Or vbExclamation, "Sound not available") <> vbRetry Then
                Con_SafePrintf "waveOutOpen failure;\n  hardware already in use\n"
                Exit Function
            End If
        Loop
    End If

    '
    ' Allocate and lock memory for the waveform data. The memory
    ' for waveform data must be globally allocated with
    ' GMEM_MOVEABLE and GMEM_SHARE flags.
    ' (QuakeVB: a VB array, never resized while the device is open)
    '
    gSndBufSize = WAV_BUFFERS * WAV_BUFFER_SIZE
    ReDim dma_buffer(0 To gSndBufSize \ 2 - 1)
    dma_made = True

    '
    ' Allocate and lock memory for the header. This memory must
    ' also be globally allocated with GMEM_MOVEABLE and
    ' GMEM_SHARE flags.
    '
    ZeroMemory lpWaveHdr(0), LenB(lpWaveHdr(0)) * WAV_BUFFERS

    ' After allocation, set up and prepare headers.
    For i = 0 To WAV_BUFFERS - 1
        lpWaveHdr(i).dwBufferLength = WAV_BUFFER_SIZE
        lpWaveHdr(i).lpData = VarPtr(dma_buffer(0)) + i * WAV_BUFFER_SIZE

        If capture_file = 0 Then
            If waveOutPrepareHeader(hWaveOut, VarPtr(lpWaveHdr(i)), LenB(lpWaveHdr(i))) <> MMSYSERR_NOERROR Then
                Con_SafePrintf "Sound: failed to prepare wave headers\n"
                FreeSound
                Exit Function
            End If
        End If
    Next

    shm.soundalive = True
    shm.splitbuffer = False
    shm.samples = gSndBufSize \ (shm.samplebits \ 8)
    shm.samplepos = 0
    shm.submission_chunk = 1
    shm.buffer = VarPtr(dma_buffer(0))
    sample16 = (shm.samplebits \ 8) - 1

    wav_init = True

    SNDDMA_InitWav = True
End Function

' Try to find a sound device to mix for.
' Returns false if nothing is found.
Public Function SNDDMA_Init() As Long
    If Not snd_firsttime_set Then
        snd_firsttime = True             ' (C: static initializer)
        snd_firsttime_set = True
    End If

    ' (QuakeVB) a -hidden test run opens no sound device unless it captures
    If sys_hidden And COM_CheckParm("-sndcapture") = 0 Then
        snd_firsttime = False
        Exit Function
    End If

    dsound_init = False
    wav_init = False

    ' (no DirectSound: C tries it first unless -wavonly is given)

    ' if DirectSound didn't succeed in initializing, try to initialize
    ' waveOut sound, unless DirectSound failed because the hardware is
    ' already allocated (in which case the user has already chosen not
    ' to have sound)
    If Not dsound_init Then
        If snd_firsttime Or snd_iswave Then
            snd_iswave = SNDDMA_InitWav()

            If snd_iswave Then
                If snd_firsttime Then
                    If capture_file Then
                        Con_SafePrintf "Wave sound initialized (captured to a file)\n"
                    Else
                        Con_SafePrintf "Wave sound initialized\n"
                    End If
                End If
            Else
                Con_SafePrintf "Wave sound failed to init\n"
            End If
        End If
    End If

    snd_firsttime = False

    If Not dsound_init And Not wav_init Then
        If snd_firsttime Then Con_SafePrintf "No sound device initialized\n"

        Exit Function
    End If

    SNDDMA_Init = 1
End Function

' return the current sample position (in mono samples read)
' inside the recirculating dma buffer, so the mixing code will know
' how many sample are required to fill it up.
Public Function SNDDMA_GetDMAPos() As Long
    Dim s As Long

    If wav_init Then s = snd_sent * WAV_BUFFER_SIZE

    s = SAR(s, sample16)

    s = s And (shm.samples - 1)

    SNDDMA_GetDMAPos = s
End Function

' (QuakeVB) -sndcapture: the blocks that would have been played by now
Private Function CaptureBlockDone(ByVal n As Long) As Boolean
    Dim played As Double

    played = (realtime - capture_start) * shm.speed * shm.channels * (shm.samplebits \ 8) / WAV_BUFFER_SIZE
    CaptureBlockDone = (n < played)
End Function

' Send sound to device if buffer isn't really the dma buffer
Public Sub SNDDMA_Submit()
    Dim h As Long, wResult As Long

    If Not wav_init Then Exit Sub

    '
    ' find which sound blocks have completed
    '
    Do
        If snd_completed = snd_sent Then
            Con_DPrintf "Sound overrun\n"
            Exit Do
        End If

        If capture_file Then
            If Not CaptureBlockDone(snd_completed) Then Exit Do
        Else
            If (lpWaveHdr(snd_completed And WAV_MASK).dwFlags And WHDR_DONE) = 0 Then Exit Do
        End If

        snd_completed = snd_completed + 1   ' this buffer has been played
    Loop

    '
    ' submit two new sound blocks
    '
    Do While SAR(snd_sent - snd_completed, sample16) < 4
        h = snd_sent And WAV_MASK

        snd_sent = snd_sent + 1
        '
        ' Now the data block can be sent to the output device. The
        ' waveOutWrite function returns immediately and waveform
        ' data is sent to the output device in the background.
        '
        If capture_file Then
            CopyMemory capture_block(0), ByVal lpWaveHdr(h).lpData, WAV_BUFFER_SIZE
            Sys_FileWrite capture_file, capture_block, 0, WAV_BUFFER_SIZE
            capture_bytes = capture_bytes + WAV_BUFFER_SIZE
        Else
            wResult = waveOutWrite(hWaveOut, VarPtr(lpWaveHdr(h)), LenB(lpWaveHdr(h)))

            If wResult <> MMSYSERR_NOERROR Then
                Con_SafePrintf "Failed to write block to device\n"
                FreeSound
                Exit Sub
            End If
        End If
    Loop
End Sub

' (QuakeVB) C's 8 bit transfer writes bytes of the buffer: out[idx] = b
Public Sub DmaPutByte(ByVal idx As Long, ByVal b As Long)
    Dim v As Byte

    v = b And &HFF&
    CopyMemory ByVal shm.buffer + idx, v, 1
End Sub

' Reset the sound device for exiting
Public Sub SNDDMA_Shutdown()
    FreeSound
End Sub
