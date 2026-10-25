Attribute VB_Name = "sys_win"
' sys_win.bas - Win32 system interface code (sys_win.c): the program entry
' point, file IO on Win32 handles, the timer, errors and quitting.
'
' A C file handle is a Win32 file handle here (-1 = none, as in C). Sys_Printf
' goes to quakevb.log in the exe folder (C: stdout of a dedicated server).
' The dedicated server console (Sys_ConsoleInput, conproc.c) and the startup
' dialog are not ported.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Declare Function CreateFileA Lib "kernel32" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, _
    ByVal lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, ByVal lpFileSizeHigh As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileAttributesA Lib "kernel32" (ByVal lpFileName As String) As Long
Private Declare Function CreateDirectoryA Lib "kernel32" (ByVal lpPathName As String, ByVal lpSecurityAttributes As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private Declare Function clearfp Lib "msvcrt.dll" Alias "_clearfp" () As Long

Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = 1
Private Const OPEN_EXISTING As Long = 3
Private Const CREATE_ALWAYS As Long = 2
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Private Const INVALID_HANDLE_VALUE As Long = -1

Private pfreq As Double
Private perf_start As Currency
Private curtime As Double
Private lastcurtime As Double
Private sametimecount As Long
Private timeinit As Boolean

Private logfile As Long
Private logopen As Boolean

' C statics of Sys_Error
Private in_sys_error0 As Boolean
Private in_sys_error1 As Boolean

Private holdrand As Long             ' C library rand () state

Public sys_hidden As Boolean         ' (QuakeVB) -hidden: no window, no MessageBox

'===============================================================================
' C library replacements

' MSVC's rand (): 15 random bits
Public Function rand() As Long
    holdrand = holdrand * 214013 + 2531011
    rand = (holdrand And &H7FFF0000) \ &H10000
End Function

Public Sub srand(ByVal seed As Long)
    holdrand = seed
End Sub

' (QuakeVB) Clears the sticky exception flags of the FPU status word.
' Compiled VB code tests them after its own float operations: once a C style
' operation has left invalid, divide by zero or overflow set (a 0/0 in the
' progs, a NaN compare), the next float to integer conversion in VB fails
' with run time error 6 or 11. Called before every builtin, at the end of
' PR_ExecuteProgram, by ToInt after a NaN, and once a frame.
Public Sub Sys_ClearFPU()
    clearfp
End Sub

'===============================================================================
' FILE IO

' returns the file length, or -1 if it cannot be opened (hndl = -1 then)
Public Function Sys_FileOpenRead(ByVal path As String, hndl As Long) As Long
    Dim h As Long

    h = CreateFileA(path, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If h = INVALID_HANDLE_VALUE Then
        hndl = -1
        Sys_FileOpenRead = -1
    Else
        hndl = h
        Sys_FileOpenRead = GetFileSize(h, 0)
    End If
End Function

Public Function Sys_FileOpenWrite(ByVal path As String) As Long
    Dim h As Long

    h = CreateFileA(path, GENERIC_WRITE, 0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If h = INVALID_HANDLE_VALUE Then Sys_Error "Error opening %s: error %i", path, Err.LastDllError
    Sys_FileOpenWrite = h
End Function

' (QuakeVB) like Sys_FileOpenWrite, but returns -1 if the file cannot be
' created (C: fopen (name, "wb") returning NULL)
Public Function Sys_FileCreate(ByVal path As String) As Long
    Sys_FileCreate = CreateFileA(path, GENERIC_WRITE, 0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
End Function

Public Sub Sys_FileClose(ByVal handle As Long)
    CloseHandle handle
End Sub

Public Sub Sys_FileSeek(ByVal handle As Long, ByVal position As Long)
    SetFilePointer handle, position, 0, 0
End Sub

' reads count bytes into dest(ofs ..)
Public Function Sys_FileRead(ByVal handle As Long, dest() As Byte, ByVal ofs As Long, ByVal count As Long) As Long
    Dim x As Long
    If count <= 0 Then Exit Function
    ReadFile handle, dest(ofs), count, x, 0
    Sys_FileRead = x
End Function

' writes count bytes from data(ofs ..)
Public Function Sys_FileWrite(ByVal handle As Long, data() As Byte, ByVal ofs As Long, ByVal count As Long) As Long
    Dim x As Long
    If count <= 0 Then Exit Function
    WriteFile handle, data(ofs), count, x, 0
    Sys_FileWrite = x
End Function

' 1 if the file exists, else -1
Public Function Sys_FileTime(ByVal path As String) As Long
    Dim a As Long
    a = GetFileAttributesA(path)
    If a <> -1 And (a And FILE_ATTRIBUTE_DIRECTORY) = 0 Then
        Sys_FileTime = 1
    Else
        Sys_FileTime = -1
    End If
End Function

Public Sub Sys_mkdir(ByVal path As String)
    CreateDirectoryA path, 0
End Sub

'===============================================================================
' SYSTEM IO

Public Sub Sys_Init()
    Dim f As Currency
    If QueryPerformanceFrequency(f) = 0 Then Sys_Error "No hardware timer available"
    pfreq = 1# / CDbl(f)
    Sys_InitFloatTime
End Sub

Public Sub Sys_Error(ByVal error_ As String, ParamArray args() As Variant)
    Dim a As Variant, text As String

    a = args
    text = vsprintf(error_, a)

    Sys_Printf "Error: %s\n", text

    If Not sys_hidden Then
        ' C: VID_SetDefaultMode so the message box is visible
        If Not in_sys_error0 Then
            in_sys_error0 = True
            MsgBox text, vbCritical Or vbSystemModal, "Quake Error"
        Else
            MsgBox text, vbCritical Or vbSystemModal, "Double Quake Error"
        End If
    End If

    If Not in_sys_error1 Then
        in_sys_error1 = True
        Host_Shutdown
    End If

    Sys_CloseLog
    timeEndPeriod 1
    PR_ClearViews                        ' (QuakeVB) before VB frees the arrays
    D_ClearViews
    End
End Sub

' (QuakeVB) a VB run time error: reported like a Sys_Error. The caller passes
' Err.Number and Err.Description (a procedure with an On Error statement,
' as in the debug build, would clear Err before they could be read here).
Public Sub Sys_RuntimeError(ByVal where As String, ByVal num As Long, ByVal desc As String)
    Sys_Error "Run time error %i (%s) in %s", num, desc, where
End Sub

Public Sub Sys_Printf(ByVal fmt As String, ParamArray args() As Variant)
    Dim a As Variant
    a = args
    Sys_Log vsprintf(fmt, a)
End Sub

' written with WriteFile, so nothing is lost if the process dies
Private Sub Sys_Log(ByVal text As String)
    Dim b() As Byte, x As Long
    If Not logopen Then
        logfile = CreateFileA(App.Path & "\quakevb.log", GENERIC_WRITE, FILE_SHARE_READ, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
        logopen = True
    End If
    If logfile = INVALID_HANDLE_VALUE Or Len(text) = 0 Then Exit Sub
    StrToBytesZ Replace(text, vbLf, vbCrLf), b
    WriteFile logfile, b(0), UBound(b), x, 0
End Sub

Private Sub Sys_CloseLog()
    If logopen Then
        If logfile <> INVALID_HANDLE_VALUE Then CloseHandle logfile
        logopen = False
    End If
End Sub

Public Sub Sys_Quit()
    Host_Shutdown
    Sys_CloseLog
    timeEndPeriod 1
    PR_ClearViews                        ' (QuakeVB) before VB frees the arrays
    D_ClearViews
    End
End Sub

' seconds since Sys_InitFloatTime, never going backwards
' (C: 32 of the 64 counter bits, shifted for about 1 microsecond resolution)
Public Function Sys_FloatTime() As Double
    Dim c As Currency, t As Double

    QueryPerformanceCounter c
    If Not timeinit Then
        perf_start = c
        timeinit = True
    End If
    t = CDbl(c - perf_start) * pfreq

    ' curtime is the base: the -starttime value, plus 1 for every 100000 calls
    ' that saw no change (as the C does)
    Sys_FloatTime = curtime + t
    If Sys_FloatTime = lastcurtime Then
        sametimecount = sametimecount + 1
        If sametimecount > 100000 Then
            curtime = curtime + 1#
            sametimecount = 0
        End If
    Else
        sametimecount = 0
    End If
    lastcurtime = Sys_FloatTime
End Function

Private Sub Sys_InitFloatTime()
    Dim j As Long

    Sys_FloatTime

    j = COM_CheckParm("-starttime")

    If j Then
        curtime = Q_atof(com_argv(j + 1))
    Else
        curtime = 0#
    End If
    lastcurtime = curtime
End Sub

' the dedicated server console is not ported: there is never input
Public Function Sys_ConsoleInput() As String
    Sys_ConsoleInput = ""
End Function

Public Sub Sys_Sleep()
    Sleep 1
End Sub

Public Sub Sys_SendKeyEvents()
    ' we always update if there are any event, even if we're paused
    DoEvents
    If VID_ProcessEvents() Then scr_skipupdate = False
End Sub

' (QuakeVB) the build time, from the exe file (C: __TIME__ " " __DATE__)
Public Function Sys_ExeTime() As String
    Dim d As Date
    On Error Resume Next
    d = FileDateTime(App.Path & "\" & App.EXEName & ".exe")
    Sys_ExeTime = Format$(d, "hh:nn:ss") & " " & Format$(d, "mmm dd yyyy")
End Function

'==============================================================================
' WINDOWS CRAP

' Program entry point (C: WinMain)
Public Sub Main()
    Dim parms As quakeparms_t, argv() As String, argc As Long
    Dim cmdline As String, i As Long, c As Long, start As Long
    Dim time_ As Double, oldtime As Double, newtime As Double

    On Error GoTo crashed

    ReDim cvars(0 To 0)                  ' the unregistered dummy cvar
    holdrand = 1

    parms.basedir = App.Path
    If Right$(parms.basedir, 1) = "\" Or Right$(parms.basedir, 1) = "/" Then
        parms.basedir = Left$(parms.basedir, Len(parms.basedir) - 1)
    End If
    parms.cachedir = ""

    ' split the command line at blanks (and characters outside 33..126)
    ReDim argv(0 To MAX_NUM_ARGVS)
    argc = 1
    argv(0) = ""
    cmdline = Command$
    i = 1
    Do While i <= Len(cmdline) And argc < MAX_NUM_ARGVS
        c = AscW(Mid$(cmdline, i, 1))
        If c <= 32 Or c > 126 Then
            i = i + 1
        Else
            start = i
            Do While i <= Len(cmdline)
                c = AscW(Mid$(cmdline, i, 1))
                If c <= 32 Or c > 126 Then Exit Do
                i = i + 1
            Loop
            argv(argc) = Mid$(cmdline, start, i - start)
            argc = argc + 1
        End If
    Loop
    parms.argc = argc

    COM_InitArgv argc, argv
    parms.argc = com_argc

    isDedicated = (COM_CheckParm("-dedicated") <> 0)
    If isDedicated Then Sys_Error "QuakeVB has no dedicated server yet"

    sys_hidden = (COM_CheckParm("-hidden") <> 0)

    timeBeginPeriod 1
    Sys_Init

    ' because sound is off until we become active
    S_BlockSound

    Sys_Printf "Host_Init\n"
    Host_Init parms

    oldtime = Sys_FloatTime()

    ' main window message loop
    Do
        ' yield the CPU for a little while when paused, minimized, or not the focus
        ' (QuakeVB: and when the last frame was skipped for running too fast)
        If (cl.paused And Not VID_IsActive()) Or VID_IsMinimized() Or block_drawing Then
            Sleep 50                     ' PAUSE_SLEEP
            scr_skipupdate = True        ' no point in bothering to draw
        ElseIf Not VID_IsActive() Then
            Sleep 20                     ' NOT_FOCUS_SLEEP
        ElseIf Not host_frameran And Not cls.timedemo Then
            Sleep 1
        End If
        newtime = Sys_FloatTime()
        time_ = newtime - oldtime

        Host_Frame time_
        oldtime = newtime
    Loop

crashed:
    Sys_RuntimeError "Main", Err.Number, Err.Description
End Sub
