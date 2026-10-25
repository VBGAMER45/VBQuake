Attribute VB_Name = "d_fill"
' d_fill.bas - clears a specified rectangle to the specified color (d_fill.c)
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public Sub D_FillRect(rect As vrect_t, ByVal color As Long)
    Dim rx As Long, ry As Long, rwidth As Long, rheight As Long, dest As Long

    rx = rect.x
    ry = rect.y
    rwidth = rect.width
    rheight = rect.height

    If rx < 0 Then
        rwidth = rwidth + rx
        rx = 0
    End If
    If ry < 0 Then
        rheight = rheight + ry
        ry = 0
    End If
    If rx + rwidth > vid.width Then rwidth = vid.width - rx
    If ry + rheight > vid.height Then rheight = vid.height - rx   ' (sic, as in C)

    If rwidth < 1 Or rheight < 1 Then Exit Sub

    dest = vid_bufferofs + ry * vid.rowbytes + rx

    ' (C clears aligned rectangles a dword at a time; the bytes are the same)
    For ry = 0 To rheight - 1
        For rx = 0 To rwidth - 1
            vid_buffer(dest + rx) = color And &HFF&
        Next
        dest = dest + vid.rowbytes
    Next
End Sub
