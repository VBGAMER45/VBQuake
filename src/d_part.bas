Attribute VB_Name = "d_part"
' d_part.bas - software driver module for drawing particles (d_part.c)
'
' C unrolls the pixel loop for particles 1 to 4 pixels wide; here the
' general loop draws them all (same pixels).
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public Sub D_EndParticles()
    ' not used by software driver
End Sub

Public Sub D_StartParticles()
    ' not used by software driver
End Sub

' pparticle: a particles() index
Public Sub D_DrawParticle(ByVal pparticle As Long)
    Dim local_ As vec3_t, transformed As vec3_t, zi As Single, pdest As Long, pz As Long
    Dim i As Long, izi As Long, pix As Long, count As Long, u As Long, v As Long, color As Byte

    ' transform point
    VectorSubtract particles(pparticle).org, r_origin, local_

    transformed.v(0) = DotProduct(local_, r_pright)
    transformed.v(1) = DotProduct(local_, r_pup)
    transformed.v(2) = DotProduct(local_, r_ppn)

    If transformed.v(2) < PARTICLE_Z_CLIP Then Exit Sub

    ' project the point
    ' FIXME: preadjust xcenter and ycenter
    zi = 1! / transformed.v(2)
    u = ToInt(xcenter + zi * transformed.v(0) + 0.5)
    v = ToInt(ycenter - zi * transformed.v(1) + 0.5)

    If v > d_vrectbottom_particle Or u > d_vrectright_particle Or v < d_vrecty Or u < d_vrectx Then
        Exit Sub
    End If

    pz = (d_zwidth * v) + u
    pdest = d_scantable(v) + u
    izi = ToInt(zi * &H8000&)

    pix = SAR(izi, d_pix_shift)

    If pix < d_pix_min Then
        pix = d_pix_min
    ElseIf pix > d_pix_max Then
        pix = d_pix_max
    End If

    color = ToInt(particles(pparticle).color) And &HFF&   ' (byte)

    count = pix * BIT_(d_y_aspect_shift)

    Do While count
        For i = 0 To pix - 1
            If d_pzbuffer(pz + i) <= izi Then
                d_pzbuffer(pz + i) = izi
                d_viewbuffer(pdest + i) = color
            End If
        Next
        count = count - 1
        pz = pz + d_zwidth
        pdest = pdest + screenwidth
    Loop
End Sub
