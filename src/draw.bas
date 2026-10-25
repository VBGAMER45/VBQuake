Attribute VB_Name = "draw"
' draw.bas - this is the only file outside the refresh that touches the
' vid buffer (draw.c). Only the 8-bit paths are ported (r_pixbytes is 1).
'
' A qpic_t* is an index into pics() (0 = NULL); the pixels of every pic live
' in draw_pool() at pics(n).ofs. The console characters are draw_chars().
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Type rectdesc_t
    rect As vrect_t
    width As Long
    height As Long
    ptexbytes As Long                ' offset into draw_pool()
    rowbytes As Long
End Type

Private r_rectdesc As rectdesc_t

Private draw_chars() As Byte         ' 8*8 graphic characters (128 x 128)
Public draw_disc As Long
Public draw_backtile As Long

Private num_pics As Long
Private pool_used As Long

Private Const MAX_CACHED_PICS As Long = 128
Private menu_cachepics(0 To MAX_CACHED_PICS - 1) As String
Private menu_cachepic_pic(0 To MAX_CACHED_PICS - 1) As Long
Private menu_numcachepics As Long

Private wad_picnames() As String     ' lumps already made into pics by Draw_PicFromWad
Private wad_picnums() As Long
Private num_wadpics As Long

'=============================================================================
' Support Routines

' Makes a pic from width * height pixels at src(srcofs).
Private Function NewPic(ByVal width As Long, ByVal height As Long, src() As Byte, ByVal srcofs As Long) As Long
    Dim size As Long, newsize As Long

    size = width * height
    If num_pics = 0 Then
        ReDim pics(0 To 15)
        ReDim draw_pool(0 To 65535)
    End If
    num_pics = num_pics + 1
    If num_pics > UBound(pics) Then ReDim Preserve pics(0 To UBound(pics) * 2 + 1)
    If pool_used + size > UBound(draw_pool) + 1 Then
        newsize = (UBound(draw_pool) + 1) * 2
        Do While newsize < pool_used + size
            newsize = newsize * 2
        Loop
        ReDim Preserve draw_pool(0 To newsize - 1)
    End If

    pics(num_pics).width = width
    pics(num_pics).height = height
    pics(num_pics).ofs = pool_used
    If size > 0 Then CopyMemory draw_pool(pool_used), src(srcofs), size
    pool_used = pool_used + size
    NewPic = num_pics
End Function

Public Function Draw_PicFromWad(ByVal name As String) As Long
    Dim i As Long, p As Long

    ' C returns a pointer into the wad, the same one every time
    For i = 1 To num_wadpics
        If wad_picnames(i) = name Then
            Draw_PicFromWad = wad_picnums(i)
            Exit Function
        End If
    Next

    p = W_GetLumpName(name)
    num_wadpics = num_wadpics + 1
    ReDim Preserve wad_picnames(1 To num_wadpics)
    ReDim Preserve wad_picnums(1 To num_wadpics)
    wad_picnames(num_wadpics) = name
    wad_picnums(num_wadpics) = NewPic(GetLong(wad_base, p), GetLong(wad_base, p + 4), wad_base, p + 8)
    Draw_PicFromWad = wad_picnums(num_wadpics)
End Function

Public Function Draw_CachePic(ByVal path As String) As Long
    Dim i As Long, dat() As Byte

    For i = 0 To menu_numcachepics - 1
        If path = menu_cachepics(i) Then Exit For
    Next

    If i = menu_numcachepics Then
        If menu_numcachepics = MAX_CACHED_PICS Then Sys_Error "menu_numcachepics == MAX_CACHED_PICS"
        menu_numcachepics = menu_numcachepics + 1
        menu_cachepics(i) = path
    End If

    If menu_cachepic_pic(i) Then
        ' (the cache never throws a pic out here)
        Draw_CachePic = menu_cachepic_pic(i)
        Exit Function
    End If

    '
    ' load the pic from disk
    '
    If Not COM_LoadFile(path, dat) Then Sys_Error "Draw_CachePic: failed to load %s", path

    menu_cachepic_pic(i) = NewPic(GetLong(dat, 0), GetLong(dat, 4), dat, 8)
    Draw_CachePic = menu_cachepic_pic(i)
End Function

Public Sub Draw_Init()
    Dim p As Long

    p = W_GetLumpName("conchars")
    ReDim draw_chars(0 To 128& * 128 - 1)
    CopyMemory draw_chars(0), wad_base(p), 128& * 128

    draw_disc = Draw_PicFromWad("disc")
    draw_backtile = Draw_PicFromWad("backtile")

    r_rectdesc.width = pics(draw_backtile).width
    r_rectdesc.height = pics(draw_backtile).height
    r_rectdesc.ptexbytes = pics(draw_backtile).ofs
    r_rectdesc.rowbytes = pics(draw_backtile).width
End Sub

' Draws one 8*8 graphics character with 0 being transparent.
' It can be clipped to the top of the screen to allow the console to be
' smoothly scrolled off.
Public Sub Draw_Character(ByVal x As Long, ByVal y As Long, ByVal num As Long)
    Dim dest As Long, source As Long, drawline As Long, row As Long, col As Long, i As Long, b As Byte

    num = num And 255

    If y <= -8 Then Exit Sub             ' totally off screen

    row = num \ 16
    col = num And 15
    source = row * 1024 + col * 8

    If y < 0 Then
        ' clipped
        drawline = 8 + y
        source = source - 128 * y
        y = 0
    Else
        drawline = 8
    End If

    dest = y * vid.conrowbytes + x
    Do While drawline > 0
        For i = 0 To 7
            b = draw_chars(source + i)
            If b Then vid_buffer(dest + i) = b
        Next
        source = source + 128
        dest = dest + vid.conrowbytes
        drawline = drawline - 1
    Loop
End Sub

Public Sub Draw_String(ByVal x As Long, ByVal y As Long, ByVal str_ As String)
    Dim i As Long
    For i = 1 To Len(str_)
        Draw_Character x, y, AscW(Mid$(str_, i, 1))
        x = x + 8
    Next
End Sub

' Draws a single character directly to the upper right corner of the screen.
' This is for debugging lockups by drawing different chars in different parts
' of the code. (vid.direct is NULL for a window, so it never draws.)
Public Sub Draw_DebugChar(ByVal num As Long)
End Sub

Public Sub Draw_Pic(ByVal x As Long, ByVal y As Long, ByVal pic As Long)
    Dim dest As Long, source As Long, v As Long, w As Long, h As Long

    w = pics(pic).width
    h = pics(pic).height
    If x < 0 Or x + w > vid.width Or y < 0 Or y + h > vid.height Then
        Sys_Error "Draw_Pic: bad coordinates"
    End If

    source = pics(pic).ofs
    dest = y * vid.rowbytes + x
    If w > 0 Then
        For v = 0 To h - 1
            CopyMemory vid_buffer(dest), draw_pool(source), w
            dest = dest + vid.rowbytes
            source = source + w
        Next
    End If
End Sub

Public Sub Draw_TransPic(ByVal x As Long, ByVal y As Long, ByVal pic As Long)
    Dim dest As Long, source As Long, tbyte As Byte, v As Long, u As Long, w As Long, h As Long
    Dim u0 As Long, u1 As Long, v0 As Long, v1 As Long

    w = pics(pic).width
    h = pics(pic).height

    ' (QuakeVB) the part off the screen is cut, where C stops with "bad
    ' coordinates": LibreQuake's net menu pictures are 264 pixels wide
    ' (transparent on the right) and drawn at x 72 of 320
    u1 = w
    v1 = h
    If x < 0 Then u0 = -x
    If y < 0 Then v0 = -y
    If x + w > vid.width Then u1 = vid.width - x
    If y + h > vid.height Then v1 = vid.height - y
    If u0 >= u1 Or v0 >= v1 Then Exit Sub

    source = pics(pic).ofs + v0 * w
    dest = (y + v0) * vid.rowbytes + x
    ' (the C unwinds the loop by 8 for widths that are multiples of 8)
    For v = v0 To v1 - 1
        For u = u0 To u1 - 1
            tbyte = draw_pool(source + u)
            If tbyte <> TRANSPARENT_COLOR Then vid_buffer(dest + u) = tbyte
        Next
        dest = dest + vid.rowbytes
        source = source + w
    Next
End Sub

Public Sub Draw_TransPicTranslate(ByVal x As Long, ByVal y As Long, ByVal pic As Long, translation() As Byte)
    Dim dest As Long, source As Long, tbyte As Byte, v As Long, u As Long, w As Long, h As Long

    w = pics(pic).width
    h = pics(pic).height
    If x < 0 Or x + w > vid.width Or y < 0 Or y + h > vid.height Then
        Sys_Error "Draw_TransPic: bad coordinates"
    End If

    source = pics(pic).ofs
    dest = y * vid.rowbytes + x
    For v = 0 To h - 1
        For u = 0 To w - 1
            tbyte = draw_pool(source + u)
            If tbyte <> TRANSPARENT_COLOR Then vid_buffer(dest + u) = translation(tbyte)
        Next
        dest = dest + vid.rowbytes
        source = source + w
    Next
End Sub

' dest: an offset into draw_pool() (the conback pic)
Private Sub Draw_CharToConback(ByVal num As Long, ByVal dest As Long)
    Dim row As Long, col As Long, source As Long, drawline As Long, x As Long

    row = num \ 16
    col = num And 15
    source = row * 1024 + col * 8

    drawline = 8

    Do While drawline > 0
        For x = 0 To 7
            If draw_chars(source + x) Then draw_pool(dest + x) = &H60 + draw_chars(source + x)
        Next
        source = source + 128
        dest = dest + 320
        drawline = drawline - 1
    Loop
End Sub

Public Sub Draw_ConsoleBackground(ByVal lines As Long)
    Dim x As Long, y As Long, v As Long, src As Long, dest As Long
    Dim f As Long, fstep As Long, conback As Long, ver As String, data As Long

    conback = Draw_CachePic("gfx/conback.lmp")
    data = pics(conback).ofs

    ' hack the version number directly into the pic
    ' (QuakeVB: C prints "(WinQuake) %4.2f")
    ver = va("(QuakeVB) %4.2f", QUAKE_VERSION)
    dest = data + 320& * 186 + 320 - 11 - 8 * Len(ver)

    For x = 0 To Len(ver) - 1
        Draw_CharToConback AscW(Mid$(ver, x + 1, 1)), dest + x * 8
    Next

    ' draw the pic
    dest = 0
    For y = 0 To lines - 1
        v = ((vid.conheight - lines + y) * 200) \ vid.conheight
        src = data + v * 320
        If vid.conwidth = 320 Then
            CopyMemory vid_buffer(dest), draw_pool(src), vid.conwidth
        Else
            f = 0
            fstep = (320& * &H10000) \ vid.conwidth
            For x = 0 To vid.conwidth - 1
                vid_buffer(dest + x) = draw_pool(src + f \ &H10000)
                f = f + fstep
            Next
        End If
        dest = dest + vid.conrowbytes
    Next
End Sub

' psrc: an offset into draw_pool()
Private Sub R_DrawRect8(prect As vrect_t, ByVal rowbytes As Long, ByVal psrc As Long, ByVal transparent As Long)
    Dim t As Byte, i As Long, j As Long, srcdelta As Long, destdelta As Long, pdest As Long

    pdest = (prect.y * vid.rowbytes) + prect.x

    srcdelta = rowbytes - prect.width
    destdelta = vid.rowbytes - prect.width

    If transparent Then
        For i = 0 To prect.height - 1
            For j = 0 To prect.width - 1
                t = draw_pool(psrc)
                If t <> TRANSPARENT_COLOR Then vid_buffer(pdest) = t
                psrc = psrc + 1
                pdest = pdest + 1
            Next
            psrc = psrc + srcdelta
            pdest = pdest + destdelta
        Next
    Else
        If prect.width <= 0 Then Exit Sub
        For i = 0 To prect.height - 1
            CopyMemory vid_buffer(pdest), draw_pool(psrc), prect.width
            psrc = psrc + rowbytes
            pdest = pdest + vid.rowbytes
        Next
    End If
End Sub

' This repeats a 64*64 tile graphic to fill the screen around a sized down
' refresh window.
Public Sub Draw_TileClear(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    Dim width As Long, height As Long, tileoffsetx As Long, tileoffsety As Long
    Dim psrc As Long, vr As vrect_t

    r_rectdesc.rect.x = x
    r_rectdesc.rect.y = y
    r_rectdesc.rect.width = w
    r_rectdesc.rect.height = h

    vr.y = r_rectdesc.rect.y
    height = r_rectdesc.rect.height

    tileoffsety = vr.y Mod r_rectdesc.height

    Do While height > 0
        vr.x = r_rectdesc.rect.x
        width = r_rectdesc.rect.width

        If tileoffsety <> 0 Then
            vr.height = r_rectdesc.height - tileoffsety
        Else
            vr.height = r_rectdesc.height
        End If

        If vr.height > height Then vr.height = height

        tileoffsetx = vr.x Mod r_rectdesc.width

        Do While width > 0
            If tileoffsetx <> 0 Then
                vr.width = r_rectdesc.width - tileoffsetx
            Else
                vr.width = r_rectdesc.width
            End If

            If vr.width > width Then vr.width = width

            psrc = r_rectdesc.ptexbytes + (tileoffsety * r_rectdesc.rowbytes) + tileoffsetx

            R_DrawRect8 vr, r_rectdesc.rowbytes, psrc, 0

            vr.x = vr.x + vr.width
            width = width - vr.width
            tileoffsetx = 0              ' only the left tile can be left-clipped
        Loop

        vr.y = vr.y + vr.height
        height = height - vr.height
        tileoffsety = 0                  ' only the top tile can be top-clipped
    Loop
End Sub

' Fills a box of pixels with a single color
Public Sub Draw_Fill(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal c As Long)
    Dim dest As Long, v As Long
    dest = y * vid.rowbytes + x
    If w <= 0 Then Exit Sub
    For v = 0 To h - 1
        FillMemory vid_buffer(dest), w, c And 255
        dest = dest + vid.rowbytes
    Next
End Sub

'=============================================================================

Public Sub Draw_FadeScreen()
    Dim x As Long, y As Long, pbuf As Long, t As Long

    S_ExtraUpdate

    For y = 0 To vid.height - 1
        pbuf = vid.rowbytes * y
        t = (y And 1) * 2
        For x = 0 To vid.width - 1
            If (x And 3) <> t Then vid_buffer(pbuf + x) = 0
        Next
    Next

    S_ExtraUpdate
End Sub

'=============================================================================

' Draws the little blue disc in the corner of the screen.
' Call before beginning any disc IO.
Public Sub Draw_BeginDisc()
    If draw_disc = 0 Then Exit Sub       ' (C passes NULL->data before Draw_Init)
    D_BeginDirectRect vid.width - 24, 0, pics(draw_disc).ofs, 24, 24
End Sub

' Erases the disc icon.
' Call after completing any disc IO
Public Sub Draw_EndDisc()
    If draw_disc = 0 Then Exit Sub
    D_EndDirectRect vid.width - 24, 0, 24, 24
End Sub
