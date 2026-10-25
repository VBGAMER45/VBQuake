Attribute VB_Name = "r_aclip"
' r_aclip.bas - clip routines for drawing Alias models directly to the screen
' (r_aclip.c)
'
' C's fv[2][8] are a_finalverts(FV_CLIP ..) (r_alias.bas): fv[n][i] is
' a_finalverts(FV_CLIP + 8 * n + i), and the finalvert_t pointers are
' indices there. R_AliasClip takes the clip flag where C takes the clip
' function as well; the flag picks the function.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private av(0 To 7) As auxvert_t

Public aclip_tri As mtriangle_t          ' C: R_AliasClipTriangle's local mtri

' pfv0 is the unclipped vertex, pfv1 is the z-clipped vertex
Private Sub R_Alias_clip_z(ByVal pfv0 As Long, ByVal pfv1 As Long, ByVal out As Long)
    Dim scale_ As Single, pav0 As Long, pav1 As Long, avout As auxvert_t

    pav0 = pfv0 - FV_CLIP
    pav1 = pfv1 - FV_CLIP

    If a_finalverts(pfv0).v(1) >= a_finalverts(pfv1).v(1) Then
        scale_ = (ALIAS_Z_CLIP_PLANE - av(pav0).fv(2)) / (av(pav1).fv(2) - av(pav0).fv(2))

        avout.fv(0) = av(pav0).fv(0) + (av(pav1).fv(0) - av(pav0).fv(0)) * scale_
        avout.fv(1) = av(pav0).fv(1) + (av(pav1).fv(1) - av(pav0).fv(1)) * scale_
        avout.fv(2) = ALIAS_Z_CLIP_PLANE

        a_finalverts(out).v(2) = ToInt(a_finalverts(pfv0).v(2) + (a_finalverts(pfv1).v(2) - a_finalverts(pfv0).v(2)) * scale_)
        a_finalverts(out).v(3) = ToInt(a_finalverts(pfv0).v(3) + (a_finalverts(pfv1).v(3) - a_finalverts(pfv0).v(3)) * scale_)
        a_finalverts(out).v(4) = ToInt(a_finalverts(pfv0).v(4) + (a_finalverts(pfv1).v(4) - a_finalverts(pfv0).v(4)) * scale_)
    Else
        scale_ = (ALIAS_Z_CLIP_PLANE - av(pav1).fv(2)) / (av(pav0).fv(2) - av(pav1).fv(2))

        avout.fv(0) = av(pav1).fv(0) + (av(pav0).fv(0) - av(pav1).fv(0)) * scale_
        avout.fv(1) = av(pav1).fv(1) + (av(pav0).fv(1) - av(pav1).fv(1)) * scale_
        avout.fv(2) = ALIAS_Z_CLIP_PLANE

        a_finalverts(out).v(2) = ToInt(a_finalverts(pfv1).v(2) + (a_finalverts(pfv0).v(2) - a_finalverts(pfv1).v(2)) * scale_)
        a_finalverts(out).v(3) = ToInt(a_finalverts(pfv1).v(3) + (a_finalverts(pfv0).v(3) - a_finalverts(pfv1).v(3)) * scale_)
        a_finalverts(out).v(4) = ToInt(a_finalverts(pfv1).v(4) + (a_finalverts(pfv0).v(4) - a_finalverts(pfv1).v(4)) * scale_)
    End If

    R_AliasProjectFinalVert a_finalverts(out), avout

    If a_finalverts(out).v(0) < r_refdef.aliasvrect.x Then a_finalverts(out).flags = a_finalverts(out).flags Or ALIAS_LEFT_CLIP
    If a_finalverts(out).v(1) < r_refdef.aliasvrect.y Then a_finalverts(out).flags = a_finalverts(out).flags Or ALIAS_TOP_CLIP
    If a_finalverts(out).v(0) > r_refdef.aliasvrectright Then a_finalverts(out).flags = a_finalverts(out).flags Or ALIAS_RIGHT_CLIP
    If a_finalverts(out).v(1) > r_refdef.aliasvrectbottom Then a_finalverts(out).flags = a_finalverts(out).flags Or ALIAS_BOTTOM_CLIP
End Sub

' (QuakeVB) out = pfv0 + (pfv1 - pfv0) * scale + 0.5 in all six components,
' the common part of the four edge clippers
Private Sub ClipLerp(ByVal pfv0 As Long, ByVal pfv1 As Long, ByVal out As Long, ByVal scale_ As Single)
    Dim i As Long

    For i = 0 To 5
        a_finalverts(out).v(i) = ToInt(a_finalverts(pfv0).v(i) + (a_finalverts(pfv1).v(i) - a_finalverts(pfv0).v(i)) * scale_ + 0.5)
    Next
End Sub

Private Sub R_Alias_clip_left(ByVal pfv0 As Long, ByVal pfv1 As Long, ByVal out As Long)
    Dim scale_ As Single

    If a_finalverts(pfv0).v(1) >= a_finalverts(pfv1).v(1) Then
        scale_ = CSng(r_refdef.aliasvrect.x - a_finalverts(pfv0).v(0)) / (a_finalverts(pfv1).v(0) - a_finalverts(pfv0).v(0))
        ClipLerp pfv0, pfv1, out, scale_
    Else
        scale_ = CSng(r_refdef.aliasvrect.x - a_finalverts(pfv1).v(0)) / (a_finalverts(pfv0).v(0) - a_finalverts(pfv1).v(0))
        ClipLerp pfv1, pfv0, out, scale_
    End If
End Sub

Private Sub R_Alias_clip_right(ByVal pfv0 As Long, ByVal pfv1 As Long, ByVal out As Long)
    Dim scale_ As Single

    If a_finalverts(pfv0).v(1) >= a_finalverts(pfv1).v(1) Then
        scale_ = CSng(r_refdef.aliasvrectright - a_finalverts(pfv0).v(0)) / (a_finalverts(pfv1).v(0) - a_finalverts(pfv0).v(0))
        ClipLerp pfv0, pfv1, out, scale_
    Else
        scale_ = CSng(r_refdef.aliasvrectright - a_finalverts(pfv1).v(0)) / (a_finalverts(pfv0).v(0) - a_finalverts(pfv1).v(0))
        ClipLerp pfv1, pfv0, out, scale_
    End If
End Sub

Private Sub R_Alias_clip_top(ByVal pfv0 As Long, ByVal pfv1 As Long, ByVal out As Long)
    Dim scale_ As Single

    If a_finalverts(pfv0).v(1) >= a_finalverts(pfv1).v(1) Then
        scale_ = CSng(r_refdef.aliasvrect.y - a_finalverts(pfv0).v(1)) / (a_finalverts(pfv1).v(1) - a_finalverts(pfv0).v(1))
        ClipLerp pfv0, pfv1, out, scale_
    Else
        scale_ = CSng(r_refdef.aliasvrect.y - a_finalverts(pfv1).v(1)) / (a_finalverts(pfv0).v(1) - a_finalverts(pfv1).v(1))
        ClipLerp pfv1, pfv0, out, scale_
    End If
End Sub

Private Sub R_Alias_clip_bottom(ByVal pfv0 As Long, ByVal pfv1 As Long, ByVal out As Long)
    Dim scale_ As Single

    If a_finalverts(pfv0).v(1) >= a_finalverts(pfv1).v(1) Then
        scale_ = CSng(r_refdef.aliasvrectbottom - a_finalverts(pfv0).v(1)) / (a_finalverts(pfv1).v(1) - a_finalverts(pfv0).v(1))
        ClipLerp pfv0, pfv1, out, scale_
    Else
        scale_ = CSng(r_refdef.aliasvrectbottom - a_finalverts(pfv1).v(1)) / (a_finalverts(pfv0).v(1) - a_finalverts(pfv1).v(1))
        ClipLerp pfv1, pfv0, out, scale_
    End If
End Sub

' in_, out: a_finalverts() indices
Private Function R_AliasClip(ByVal in_ As Long, ByVal out As Long, ByVal flag As Long, ByVal count As Long) As Long
    Dim i As Long, j As Long, k As Long, flags As Long, oldflags As Long

    j = count - 1
    k = 0
    For i = 0 To count - 1
        oldflags = a_finalverts(in_ + j).flags And flag
        flags = a_finalverts(in_ + i).flags And flag

        If flags <> 0 And oldflags <> 0 Then GoTo continue_
        If oldflags Xor flags Then
            Select Case flag
            Case ALIAS_Z_CLIP: R_Alias_clip_z in_ + j, in_ + i, out + k
            Case ALIAS_LEFT_CLIP: R_Alias_clip_left in_ + j, in_ + i, out + k
            Case ALIAS_RIGHT_CLIP: R_Alias_clip_right in_ + j, in_ + i, out + k
            Case ALIAS_BOTTOM_CLIP: R_Alias_clip_bottom in_ + j, in_ + i, out + k
            Case ALIAS_TOP_CLIP: R_Alias_clip_top in_ + j, in_ + i, out + k
            End Select
            a_finalverts(out + k).flags = 0
            If a_finalverts(out + k).v(0) < r_refdef.aliasvrect.x Then a_finalverts(out + k).flags = a_finalverts(out + k).flags Or ALIAS_LEFT_CLIP
            If a_finalverts(out + k).v(1) < r_refdef.aliasvrect.y Then a_finalverts(out + k).flags = a_finalverts(out + k).flags Or ALIAS_TOP_CLIP
            If a_finalverts(out + k).v(0) > r_refdef.aliasvrectright Then a_finalverts(out + k).flags = a_finalverts(out + k).flags Or ALIAS_RIGHT_CLIP
            If a_finalverts(out + k).v(1) > r_refdef.aliasvrectbottom Then a_finalverts(out + k).flags = a_finalverts(out + k).flags Or ALIAS_BOTTOM_CLIP
            k = k + 1
        End If
        If flags = 0 Then
            a_finalverts(out + k) = a_finalverts(in_ + i)
            k = k + 1
        End If
continue_:
        j = i
    Next

    R_AliasClip = k
End Function

' ptri: an alias_tris() index
Public Sub R_AliasClipTriangle(ByVal ptri As Long)
    Dim i As Long, k As Long, pingpong As Long, clipflags As Long, fv0 As Long, fv1 As Long

    fv0 = FV_CLIP                        ' C: fv[0]

    ' copy vertexes and fix seam texture coordinates
    If alias_tris(ptri).facesfront Then
        a_finalverts(fv0) = a_finalverts(alias_tris(ptri).vertindex(0))
        a_finalverts(fv0 + 1) = a_finalverts(alias_tris(ptri).vertindex(1))
        a_finalverts(fv0 + 2) = a_finalverts(alias_tris(ptri).vertindex(2))
    Else
        For i = 0 To 2
            a_finalverts(fv0 + i) = a_finalverts(alias_tris(ptri).vertindex(i))

            If alias_tris(ptri).facesfront = 0 And (a_finalverts(fv0 + i).flags And ALIAS_ONSEAM) <> 0 Then
                a_finalverts(fv0 + i).v(2) = a_finalverts(fv0 + i).v(2) + r_affinetridesc.seamfixupX16
            End If
        Next
    End If

    ' clip
    clipflags = a_finalverts(fv0).flags Or a_finalverts(fv0 + 1).flags Or a_finalverts(fv0 + 2).flags

    If clipflags And ALIAS_Z_CLIP Then
        For i = 0 To 2
            av(i) = a_auxverts(alias_tris(ptri).vertindex(i))
        Next

        k = R_AliasClip(FV_CLIP, FV_CLIP + 8, ALIAS_Z_CLIP, 3)
        If k = 0 Then Exit Sub

        pingpong = 1
        fv1 = FV_CLIP + 8
        clipflags = a_finalverts(fv1).flags Or a_finalverts(fv1 + 1).flags Or a_finalverts(fv1 + 2).flags
    Else
        pingpong = 0
        k = 3
    End If

    If clipflags And ALIAS_LEFT_CLIP Then
        k = R_AliasClip(FV_CLIP + 8 * pingpong, FV_CLIP + 8 * (pingpong Xor 1), ALIAS_LEFT_CLIP, k)
        If k = 0 Then Exit Sub

        pingpong = pingpong Xor 1
    End If

    If clipflags And ALIAS_RIGHT_CLIP Then
        k = R_AliasClip(FV_CLIP + 8 * pingpong, FV_CLIP + 8 * (pingpong Xor 1), ALIAS_RIGHT_CLIP, k)
        If k = 0 Then Exit Sub

        pingpong = pingpong Xor 1
    End If

    If clipflags And ALIAS_BOTTOM_CLIP Then
        k = R_AliasClip(FV_CLIP + 8 * pingpong, FV_CLIP + 8 * (pingpong Xor 1), ALIAS_BOTTOM_CLIP, k)
        If k = 0 Then Exit Sub

        pingpong = pingpong Xor 1
    End If

    If clipflags And ALIAS_TOP_CLIP Then
        k = R_AliasClip(FV_CLIP + 8 * pingpong, FV_CLIP + 8 * (pingpong Xor 1), ALIAS_TOP_CLIP, k)
        If k = 0 Then Exit Sub

        pingpong = pingpong Xor 1
    End If

    fv1 = FV_CLIP + 8 * pingpong
    For i = 0 To k - 1
        If a_finalverts(fv1 + i).v(0) < r_refdef.aliasvrect.x Then
            a_finalverts(fv1 + i).v(0) = r_refdef.aliasvrect.x
        ElseIf a_finalverts(fv1 + i).v(0) > r_refdef.aliasvrectright Then
            a_finalverts(fv1 + i).v(0) = r_refdef.aliasvrectright
        End If

        If a_finalverts(fv1 + i).v(1) < r_refdef.aliasvrect.y Then
            a_finalverts(fv1 + i).v(1) = r_refdef.aliasvrect.y
        ElseIf a_finalverts(fv1 + i).v(1) > r_refdef.aliasvrectbottom Then
            a_finalverts(fv1 + i).v(1) = r_refdef.aliasvrectbottom
        End If

        a_finalverts(fv1 + i).flags = 0
    Next

    ' draw triangles
    aclip_tri.facesfront = alias_tris(ptri).facesfront
    r_affinetridesc.ptriangles = -1      ' aclip_tri
    r_affinetridesc.pfinalverts = fv1

    ' FIXME: do all at once as trifan?
    aclip_tri.vertindex(0) = 0
    For i = 1 To k - 2
        aclip_tri.vertindex(1) = i
        aclip_tri.vertindex(2) = i + 1
        D_PolysetDraw
    Next
End Sub
