Attribute VB_Name = "r_sky"
' r_sky.bas - the moving sky (r_sky.c)
'
' Not ported: R_GenSkyTile and R_GenSkyTile16, which nothing calls (the 16
' bit one draws for r_pixbytes 2, which is never set).
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const iskyspeed As Long = 8
Private Const iskyspeed2 As Long = 2
Public skyspeed As Single, skyspeed2 As Single

Public skytime As Single

' (C: byte *r_skysource, which is always newsky; d_sky.bas reads newsky())

Public r_skymade As Long
Public r_skydirect As Long               ' not used?

' TODO: clean up these routines

' (dynamic: as fixed arrays the three are over VB's 64K of module data)
Private bottomsky() As Byte              ' 128*131
Private bottommask() As Byte             ' 128*131
Public newsky() As Byte                  ' 128*256: newsky and topsky both pack in here, 128 bytes
                                         '  of newsky on the left of each scan, 128 bytes
                                         '  of topsky on the right, because the low-level
                                         '  drawers need 256-byte scan widths

Private xlast As Long, ylast As Long     ' C: statics of R_MakeSky, = -1 at the start
Private xylast_set As Boolean
Private sky_allocated As Boolean

Private Sub AllocSky()
    If sky_allocated Then Exit Sub
    ReDim bottomsky(0 To 128& * 131 - 1)
    ReDim bottommask(0 To 128& * 131 - 1)
    ReDim newsky(0 To 128& * 256 - 1)
    sky_allocated = True
End Sub

' A sky texture is 256*128, with the right side being a masked overlay
Public Sub R_InitSky(ByVal tx As Long)
    Dim i As Long, j As Long, src As Long, p As Byte

    AllocSky
    src = textures(tx).offsets(0)

    For i = 0 To 127
        For j = 0 To 127
            newsky((i * 256) + j + 128) = texpixels(src + i * 256 + j + 128)
        Next
    Next

    For i = 0 To 127
        For j = 0 To 130
            p = texpixels(src + i * 256 + (j And &H7F))
            If p Then
                bottomsky((i * 131) + j) = p
                bottommask((i * 131) + j) = 0
            Else
                bottomsky((i * 131) + j) = 0
                bottommask((i * 131) + j) = &HFF
            End If
        Next
    Next
End Sub

Public Sub R_MakeSky()
    Dim x As Long, y As Long, ofs As Long, baseofs As Long, xshift As Long, yshift As Long, pnewsky As Long

    AllocSky
    If Not xylast_set Then
        xlast = -1
        ylast = -1
        xylast_set = True
    End If

    xshift = ToInt(skytime * skyspeed)
    yshift = ToInt(skytime * skyspeed)

    If xshift = xlast And yshift = ylast Then Exit Sub

    xlast = xshift
    ylast = yshift

    ' (the byte by byte version of C's loop; UNALIGNED_OK does the same four
    ' bytes at a time)
    pnewsky = 0

    For y = 0 To SKYSIZE - 1
        baseofs = ((y + yshift) And SKYMASK) * 131

        For x = 0 To SKYSIZE - 1
            ofs = baseofs + ((x + xshift) And SKYMASK)

            newsky(pnewsky) = (newsky(pnewsky + 128) And bottommask(ofs)) Or bottomsky(ofs)
            pnewsky = pnewsky + 1
        Next

        pnewsky = pnewsky + 128
    Next

    r_skymade = 1
End Sub

Public Sub R_SetSkyFrame()
    Dim g As Long, s1 As Long, s2 As Long, temp As Single

    skyspeed = iskyspeed
    skyspeed2 = iskyspeed2

    g = GreatestCommonDivisor(iskyspeed, iskyspeed2)
    s1 = iskyspeed \ g
    s2 = iskyspeed2 \ g
    temp = SKYSIZE * s1 * s2

    skytime = cl.time - (ToInt(cl.time / temp) * temp)

    r_skymade = 0
End Sub
