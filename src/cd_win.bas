Attribute VB_Name = "cd_win"
' cd_win.bas - CD audio through MCI (cd_win.c)
'
' The window procedure cannot run engine code, so it only queues the
' MM_MCINOTIFY messages (CDAudio_QueueNotify); CDAudio_Update hands them to
' CDAudio_MessageHandler (C calls it from the window procedure). Under
' -hidden the CD is not opened, like the sound device.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Type MCI_OPEN_PARMS
    dwCallback As Long
    wDeviceID As Long
    lpstrDeviceType As Long
    lpstrElementName As Long
    lpstrAlias As Long
End Type

Private Type MCI_SET_PARMS
    dwCallback As Long
    dwTimeFormat As Long
    dwAudio As Long
End Type

Private Type MCI_STATUS_PARMS
    dwCallback As Long
    dwReturn As Long
    dwItem As Long
    dwTrack As Long
End Type

Private Type MCI_PLAY_PARMS
    dwCallback As Long
    dwFrom As Long
    dwTo As Long
End Type

Private Type MCI_GENERIC_PARMS
    dwCallback As Long
End Type

Private Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, _
    ByVal uMessage As Long, ByVal dwParam1 As Long, dwParam2 As Any) As Long

Private Const MCI_OPEN As Long = &H803
Private Const MCI_CLOSE As Long = &H804
Private Const MCI_PLAY As Long = &H806
Private Const MCI_STOP As Long = &H808
Private Const MCI_PAUSE As Long = &H809
Private Const MCI_SET As Long = &H80D
Private Const MCI_STATUS As Long = &H814

Private Const MCI_NOTIFY As Long = &H1
Private Const MCI_WAIT As Long = &H2
Private Const MCI_FROM As Long = &H4
Private Const MCI_TO As Long = &H8
Private Const MCI_TRACK As Long = &H10
Private Const MCI_OPEN_SHAREABLE As Long = &H100
Private Const MCI_OPEN_TYPE As Long = &H2000
Private Const MCI_STATUS_ITEM As Long = &H100
Private Const MCI_SET_DOOR_OPEN As Long = &H100
Private Const MCI_SET_DOOR_CLOSED As Long = &H200
Private Const MCI_SET_TIME_FORMAT As Long = &H400

Private Const MCI_STATUS_LENGTH As Long = 1
Private Const MCI_STATUS_NUMBER_OF_TRACKS As Long = 3
Private Const MCI_STATUS_READY As Long = 7
Private Const MCI_CDA_STATUS_TYPE_TRACK As Long = &H4001
Private Const MCI_CDA_TRACK_AUDIO As Long = &H440   ' MCI_CD_OFFSET + 0
Private Const MCI_FORMAT_TMSF As Long = 10

Private Const MCI_NOTIFY_SUCCESSFUL As Long = 1
Private Const MCI_NOTIFY_SUPERSEDED As Long = 2
Private Const MCI_NOTIFY_ABORTED As Long = 4
Private Const MCI_NOTIFY_FAILURE As Long = 8

Private cdValid As Boolean
Private playing As Boolean
Private wasPlaying As Boolean
Private initialized As Boolean
Private enabled As Boolean
Private playLooping As Boolean
Private cdvolume As Single
Private remap(0 To 99) As Byte
Private cdrom As Byte
Private playTrack As Byte
Private maxTrack As Byte

Private wDeviceID As Long

Private cdaudio_ansi As String           ' "cdaudio" as ANSI bytes, for lpstrDeviceType

' the queued MM_MCINOTIFY messages
Private Const NOTIFY_QUEUE As Long = 16
Private notify_wparam(0 To NOTIFY_QUEUE - 1) As Long
Private notify_lparam(0 To NOTIFY_QUEUE - 1) As Long
Private notify_head As Long, notify_tail As Long

' C: MCI_MAKE_TMSF
Private Function MCI_MAKE_TMSF(ByVal t As Long, ByVal m As Long, ByVal s As Long, ByVal f As Long) As Long
    MCI_MAKE_TMSF = (t And &HFF&) Or (m And &HFF&) * &H100& Or (s And &HFF&) * &H10000
    If f And &H80& Then
        MCI_MAKE_TMSF = MCI_MAKE_TMSF Or (f And &H7F&) * &H1000000 Or &H80000000
    Else
        MCI_MAKE_TMSF = MCI_MAKE_TMSF Or (f And &H7F&) * &H1000000
    End If
End Function

Private Sub CDAudio_Eject()
    Dim dwReturn As Long

    dwReturn = mciSendCommand(wDeviceID, MCI_SET, MCI_SET_DOOR_OPEN, ByVal 0&)
    If dwReturn Then Con_DPrintf "MCI_SET_DOOR_OPEN failed (%i)\n", dwReturn
End Sub

Private Sub CDAudio_CloseDoor()
    Dim dwReturn As Long

    dwReturn = mciSendCommand(wDeviceID, MCI_SET, MCI_SET_DOOR_CLOSED, ByVal 0&)
    If dwReturn Then Con_DPrintf "MCI_SET_DOOR_CLOSED failed (%i)\n", dwReturn
End Sub

Private Function CDAudio_GetAudioDiskInfo() As Long
    Dim dwReturn As Long, mciStatusParms As MCI_STATUS_PARMS

    cdValid = False

    mciStatusParms.dwItem = MCI_STATUS_READY
    dwReturn = mciSendCommand(wDeviceID, MCI_STATUS, MCI_STATUS_ITEM Or MCI_WAIT, mciStatusParms)
    If dwReturn Then
        Con_DPrintf "CDAudio: drive ready test - get status failed\n"
        CDAudio_GetAudioDiskInfo = -1
        Exit Function
    End If
    If mciStatusParms.dwReturn = 0 Then
        Con_DPrintf "CDAudio: drive not ready\n"
        CDAudio_GetAudioDiskInfo = -1
        Exit Function
    End If

    mciStatusParms.dwItem = MCI_STATUS_NUMBER_OF_TRACKS
    dwReturn = mciSendCommand(wDeviceID, MCI_STATUS, MCI_STATUS_ITEM Or MCI_WAIT, mciStatusParms)
    If dwReturn Then
        Con_DPrintf "CDAudio: get tracks - status failed\n"
        CDAudio_GetAudioDiskInfo = -1
        Exit Function
    End If
    If mciStatusParms.dwReturn < 1 Then
        Con_DPrintf "CDAudio: no music tracks\n"
        CDAudio_GetAudioDiskInfo = -1
        Exit Function
    End If

    cdValid = True
    maxTrack = mciStatusParms.dwReturn And &HFF&

    CDAudio_GetAudioDiskInfo = 0
End Function

Public Sub CDAudio_Play(ByVal track As Long, ByVal looping As Boolean)
    Dim dwReturn As Long, mciPlayParms As MCI_PLAY_PARMS, mciStatusParms As MCI_STATUS_PARMS

    track = track And &HFF&              ' (byte)

    If Not enabled Then Exit Sub

    If Not cdValid Then
        CDAudio_GetAudioDiskInfo
        If Not cdValid Then Exit Sub
    End If

    track = remap(track)

    If track < 1 Or track > maxTrack Then
        Con_DPrintf "CDAudio: Bad track number %u.\n", track
        Exit Sub
    End If

    ' don't try to play a non-audio track
    mciStatusParms.dwItem = MCI_CDA_STATUS_TYPE_TRACK
    mciStatusParms.dwTrack = track
    dwReturn = mciSendCommand(wDeviceID, MCI_STATUS, MCI_STATUS_ITEM Or MCI_TRACK Or MCI_WAIT, mciStatusParms)
    If dwReturn Then
        Con_DPrintf "MCI_STATUS failed (%i)\n", dwReturn
        Exit Sub
    End If
    If mciStatusParms.dwReturn <> MCI_CDA_TRACK_AUDIO Then
        Con_Printf "CDAudio: track %i is not audio\n", track
        Exit Sub
    End If

    ' get the length of the track to be played
    mciStatusParms.dwItem = MCI_STATUS_LENGTH
    mciStatusParms.dwTrack = track
    dwReturn = mciSendCommand(wDeviceID, MCI_STATUS, MCI_STATUS_ITEM Or MCI_TRACK Or MCI_WAIT, mciStatusParms)
    If dwReturn Then
        Con_DPrintf "MCI_STATUS failed (%i)\n", dwReturn
        Exit Sub
    End If

    If playing Then
        If playTrack = track Then Exit Sub
        CDAudio_Stop
    End If

    mciPlayParms.dwFrom = MCI_MAKE_TMSF(track, 0, 0, 0)
    mciPlayParms.dwTo = (mciStatusParms.dwReturn * &H100&) Or track
    mciPlayParms.dwCallback = mainwindow
    dwReturn = mciSendCommand(wDeviceID, MCI_PLAY, MCI_NOTIFY Or MCI_FROM Or MCI_TO, mciPlayParms)
    If dwReturn Then
        Con_DPrintf "CDAudio: MCI_PLAY failed (%i)\n", dwReturn
        Exit Sub
    End If

    playLooping = looping
    playTrack = track
    playing = True

    If cdvolume = 0 Then CDAudio_Pause
End Sub

Public Sub CDAudio_Stop()
    Dim dwReturn As Long

    If Not enabled Then Exit Sub

    If Not playing Then Exit Sub

    dwReturn = mciSendCommand(wDeviceID, MCI_STOP, 0, ByVal 0&)
    If dwReturn Then Con_DPrintf "MCI_STOP failed (%i)", dwReturn

    wasPlaying = False
    playing = False
End Sub

Public Sub CDAudio_Pause()
    Dim dwReturn As Long, mciGenericParms As MCI_GENERIC_PARMS

    If Not enabled Then Exit Sub

    If Not playing Then Exit Sub

    mciGenericParms.dwCallback = mainwindow
    dwReturn = mciSendCommand(wDeviceID, MCI_PAUSE, 0, mciGenericParms)
    If dwReturn Then Con_DPrintf "MCI_PAUSE failed (%i)", dwReturn

    wasPlaying = playing
    playing = False
End Sub

Public Sub CDAudio_Resume()
    Dim dwReturn As Long, mciPlayParms As MCI_PLAY_PARMS

    If Not enabled Then Exit Sub

    If Not cdValid Then Exit Sub

    If Not wasPlaying Then Exit Sub

    mciPlayParms.dwFrom = MCI_MAKE_TMSF(playTrack, 0, 0, 0)
    mciPlayParms.dwTo = MCI_MAKE_TMSF(playTrack + 1, 0, 0, 0)
    mciPlayParms.dwCallback = mainwindow
    dwReturn = mciSendCommand(wDeviceID, MCI_PLAY, MCI_TO Or MCI_NOTIFY, mciPlayParms)
    If dwReturn Then
        Con_DPrintf "CDAudio: MCI_PLAY failed (%i)\n", dwReturn
        Exit Sub
    End If
    playing = True
End Sub

Public Sub CD_f()
    Dim command As String, ret As Long, n As Long

    If Cmd_Argc() < 2 Then Exit Sub

    command = Cmd_Argv(1)

    If Q_strcasecmp(command, "on") = 0 Then
        enabled = True
        Exit Sub
    End If

    If Q_strcasecmp(command, "off") = 0 Then
        If playing Then CDAudio_Stop
        enabled = False
        Exit Sub
    End If

    If Q_strcasecmp(command, "reset") = 0 Then
        enabled = True
        If playing Then CDAudio_Stop
        For n = 0 To 99
            remap(n) = n
        Next
        CDAudio_GetAudioDiskInfo
        Exit Sub
    End If

    If Q_strcasecmp(command, "remap") = 0 Then
        ret = Cmd_Argc() - 2
        If ret <= 0 Then
            For n = 1 To 99
                If remap(n) <> n Then Con_Printf "  %u -> %u\n", n, remap(n)
            Next
            Exit Sub
        End If
        For n = 1 To ret
            remap(n) = Q_atoi(Cmd_Argv(n + 1)) And &HFF&
        Next
        Exit Sub
    End If

    If Q_strcasecmp(command, "close") = 0 Then
        CDAudio_CloseDoor
        Exit Sub
    End If

    If Not cdValid Then
        CDAudio_GetAudioDiskInfo
        If Not cdValid Then
            Con_Printf "No CD in player.\n"
            Exit Sub
        End If
    End If

    If Q_strcasecmp(command, "play") = 0 Then
        CDAudio_Play Q_atoi(Cmd_Argv(2)) And &HFF&, False
        Exit Sub
    End If

    If Q_strcasecmp(command, "loop") = 0 Then
        CDAudio_Play Q_atoi(Cmd_Argv(2)) And &HFF&, True
        Exit Sub
    End If

    If Q_strcasecmp(command, "stop") = 0 Then
        CDAudio_Stop
        Exit Sub
    End If

    If Q_strcasecmp(command, "pause") = 0 Then
        CDAudio_Pause
        Exit Sub
    End If

    If Q_strcasecmp(command, "resume") = 0 Then
        CDAudio_Resume
        Exit Sub
    End If

    If Q_strcasecmp(command, "eject") = 0 Then
        If playing Then CDAudio_Stop
        CDAudio_Eject
        cdValid = False
        Exit Sub
    End If

    If Q_strcasecmp(command, "info") = 0 Then
        Con_Printf "%u tracks\n", maxTrack
        If playing Then
            If playLooping Then
                Con_Printf "Currently %s track %u\n", "looping", playTrack
            Else
                Con_Printf "Currently %s track %u\n", "playing", playTrack
            End If
        ElseIf wasPlaying Then
            If playLooping Then
                Con_Printf "Paused %s track %u\n", "looping", playTrack
            Else
                Con_Printf "Paused %s track %u\n", "playing", playTrack
            End If
        End If
        Con_Printf "Volume is %f\n", cdvolume
        Exit Sub
    End If
End Sub

' (QuakeVB) called by the window procedure for MM_MCINOTIFY: only queues it
Public Sub CDAudio_QueueNotify(ByVal wParam As Long, ByVal lParam As Long)
    Dim n As Long

    n = (notify_head + 1) And (NOTIFY_QUEUE - 1)
    If n = notify_tail Then Exit Sub     ' full
    notify_wparam(notify_head) = wParam
    notify_lparam(notify_head) = lParam
    notify_head = n
End Sub

Private Function CDAudio_MessageHandler(ByVal wParam As Long, ByVal lParam As Long) As Long
    If lParam <> wDeviceID Then
        CDAudio_MessageHandler = 1
        Exit Function
    End If

    Select Case wParam
    Case MCI_NOTIFY_SUCCESSFUL
        If playing Then
            playing = False
            If playLooping Then CDAudio_Play playTrack, True
        End If

    Case MCI_NOTIFY_ABORTED, MCI_NOTIFY_SUPERSEDED

    Case MCI_NOTIFY_FAILURE
        Con_DPrintf "MCI_NOTIFY_FAILURE\n"
        CDAudio_Stop
        cdValid = False

    Case Else
        Con_DPrintf "Unexpected MM_MCINOTIFY type (%i)\n", wParam
        CDAudio_MessageHandler = 1
        Exit Function
    End Select

    CDAudio_MessageHandler = 0
End Function

Public Sub CDAudio_Update()
    ' (QuakeVB) the notifications the window procedure queued
    Do While notify_tail <> notify_head
        CDAudio_MessageHandler notify_wparam(notify_tail), notify_lparam(notify_tail)
        notify_tail = (notify_tail + 1) And (NOTIFY_QUEUE - 1)
    Loop

    If Not enabled Then Exit Sub

    If cvars(cv_bgmvolume).value <> cdvolume Then
        If cdvolume Then
            Cvar_SetValue "bgmvolume", 0!
            cdvolume = cvars(cv_bgmvolume).value
            CDAudio_Pause
        Else
            Cvar_SetValue "bgmvolume", 1!
            cdvolume = cvars(cv_bgmvolume).value
            CDAudio_Resume
        End If
    End If
End Sub

Public Function CDAudio_Init() As Long
    Dim dwReturn As Long, mciOpenParms As MCI_OPEN_PARMS, mciSetParms As MCI_SET_PARMS, n As Long

    If cls.state = ca_dedicated Then
        CDAudio_Init = -1
        Exit Function
    End If

    If COM_CheckParm("-nocdaudio") Then
        CDAudio_Init = -1
        Exit Function
    End If

    ' (QuakeVB) a -hidden test run leaves the CD alone
    If sys_hidden Then
        CDAudio_Init = -1
        Exit Function
    End If

    cdaudio_ansi = StrConv("cdaudio" & vbNullChar, vbFromUnicode)
    mciOpenParms.lpstrDeviceType = StrPtr(cdaudio_ansi)
    dwReturn = mciSendCommand(0, MCI_OPEN, MCI_OPEN_TYPE Or MCI_OPEN_SHAREABLE, mciOpenParms)
    If dwReturn Then
        Con_Printf "CDAudio_Init: MCI_OPEN failed (%i)\n", dwReturn
        CDAudio_Init = -1
        Exit Function
    End If
    wDeviceID = mciOpenParms.wDeviceID

    ' Set the time format to track/minute/second/frame (TMSF).
    mciSetParms.dwTimeFormat = MCI_FORMAT_TMSF
    dwReturn = mciSendCommand(wDeviceID, MCI_SET, MCI_SET_TIME_FORMAT, mciSetParms)
    If dwReturn Then
        Con_Printf "MCI_SET_TIME_FORMAT failed (%i)\n", dwReturn
        mciSendCommand wDeviceID, MCI_CLOSE, 0, ByVal 0&
        CDAudio_Init = -1
        Exit Function
    End If

    For n = 0 To 99
        remap(n) = n
    Next
    initialized = True
    enabled = True

    If CDAudio_GetAudioDiskInfo() Then
        Con_Printf "CDAudio_Init: No CD in player.\n"
        cdValid = False
    End If

    Cmd_AddCommand "cd", XC_CD_f

    Con_Printf "CD Audio Initialized\n"

    CDAudio_Init = 0
End Function

Public Sub CDAudio_Shutdown()
    If Not initialized Then Exit Sub
    CDAudio_Stop
    If mciSendCommand(wDeviceID, MCI_CLOSE, MCI_WAIT, ByVal 0&) Then Con_DPrintf "CDAudio_Shutdown: MCI_CLOSE failed\n"
End Sub
