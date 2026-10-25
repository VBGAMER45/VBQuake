Attribute VB_Name = "wad"
' wad.bas - the gfx.wad lump directory (wad.c). W_GetLumpName returns the
' lump's offset in wad_base() (C: a pointer into it).
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public wad_numlumps As Long
Public wad_lumps() As lumpinfo_t
Public wad_base() As Byte

' Lowercases name and pads with spaces and a terminating 0 to the length of
' lumpinfo_t->name (here: lowercases and cuts the name at 16 characters).
Public Function W_CleanupName(ByVal in_ As String) As String
    Dim i As Long, c As Long, out As String
    For i = 1 To 16
        If i > Len(in_) Then Exit For
        c = AscW(Mid$(in_, i, 1))
        If c = 0 Then Exit For
        If c >= 65 And c <= 90 Then c = c + 32
        out = out & ChrW$(c)
    Next
    W_CleanupName = out
End Function

Public Sub W_LoadWadFile(ByVal filename As String)
    Dim i As Long, infotableofs As Long, p As Long

    If Not COM_LoadHunkFile(filename, wad_base) Then Sys_Error "W_LoadWadFile: couldn't load %s", filename

    If wad_base(0) <> 87 Or wad_base(1) <> 65 Or wad_base(2) <> 68 Or wad_base(3) <> 50 Then   ' "WAD2"
        Sys_Error "Wad file %s doesn't have WAD2 id\n", filename
    End If

    wad_numlumps = GetLong(wad_base, 4)
    infotableofs = GetLong(wad_base, 8)
    ReDim wad_lumps(0 To wad_numlumps)

    For i = 0 To wad_numlumps - 1
        p = infotableofs + i * 32
        wad_lumps(i).filepos = GetLong(wad_base, p)
        wad_lumps(i).disksize = GetLong(wad_base, p + 4)
        wad_lumps(i).size = GetLong(wad_base, p + 8)
        wad_lumps(i).typ = wad_base(p + 12)
        wad_lumps(i).compression = wad_base(p + 13)
        wad_lumps(i).name = W_CleanupName(BufToStrN(wad_base, p + 16, 16))
        ' (TYP_QPIC lumps need no SwapPic: Windows is little-endian)
    Next
End Sub

' index into wad_lumps()
Public Function W_GetLumpinfo(ByVal name As String) As Long
    Dim i As Long, clean As String

    clean = W_CleanupName(name)

    For i = 0 To wad_numlumps - 1
        If clean = wad_lumps(i).name Then
            W_GetLumpinfo = i
            Exit Function
        End If
    Next

    Sys_Error "W_GetLumpinfo: %s not found", name
End Function

Public Function W_GetLumpName(ByVal name As String) As Long
    W_GetLumpName = wad_lumps(W_GetLumpinfo(name)).filepos
End Function

Public Function W_GetLumpNum(ByVal num As Long) As Long
    If num < 0 Or num > wad_numlumps Then Sys_Error "W_GetLumpNum: bad number: %i", num
    W_GetLumpNum = wad_lumps(num).filepos
End Function
