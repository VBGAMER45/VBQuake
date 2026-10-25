Attribute VB_Name = "d_zpoint"
' d_zpoint.bas - software driver module for drawing z-buffered points
' (d_zpoint.c)
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public Sub D_DrawZPoint()
    Dim pdest As Long, pz As Long, izi As Long

    pz = d_zwidth * r_zpointdesc.v + r_zpointdesc.u
    pdest = d_scantable(r_zpointdesc.v) + r_zpointdesc.u
    izi = ToInt(r_zpointdesc.zi * &H8000&)

    If d_pzbuffer(pz) <= izi Then
        d_pzbuffer(pz) = ((izi + &H8000&) And &HFFFF&) - &H8000&   ' (short)
        d_viewbuffer(pdest) = r_zpointdesc.color And &HFF&
    End If
End Sub
