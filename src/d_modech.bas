Attribute VB_Name = "d_modech"
' d_modech.bas - called when mode has just changed (d_modech.c)
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public d_vrectx As Long, d_vrecty As Long, d_vrectright_particle As Long, d_vrectbottom_particle As Long

Public d_y_aspect_shift As Long, d_pix_min As Long, d_pix_max As Long, d_pix_shift As Long

Public d_scantable(0 To MAXHEIGHT - 1) As Long
Public zspantable(0 To MAXHEIGHT - 1) As Long    ' d_pzbuffer() offsets (C: short *)

Private Sub D_Patch()
    ' (only patches code on Intel)
End Sub

Public Sub D_ViewChanged()
    Dim rowbytes As Long, i As Long

    If r_dowarp Then
        rowbytes = WARP_WIDTH
    Else
        rowbytes = vid.rowbytes
    End If

    scale_for_mip = xscale
    If yscale > xscale Then scale_for_mip = yscale

    d_zrowbytes = vid.width * 2
    d_zwidth = vid.width

    d_pix_min = r_refdef.vrect.width \ 320
    If d_pix_min < 1 Then d_pix_min = 1

    d_pix_max = ToInt(CSng(r_refdef.vrect.width) / (320# / 4#) + 0.5)
    d_pix_shift = 8 - ToInt(CSng(r_refdef.vrect.width) / 320# + 0.5)
    If d_pix_max < 1 Then d_pix_max = 1

    If pixelAspect > 1.4 Then
        d_y_aspect_shift = 1
    Else
        d_y_aspect_shift = 0
    End If

    d_vrectx = r_refdef.vrect.x
    d_vrecty = r_refdef.vrect.y
    d_vrectright_particle = r_refdef.vrectright - d_pix_max
    d_vrectbottom_particle = r_refdef.vrectbottom - (d_pix_max * BIT_(d_y_aspect_shift))

    For i = 0 To vid.height - 1
        d_scantable(i) = i * rowbytes
        zspantable(i) = i * d_zwidth
    Next

    D_Patch
End Sub
