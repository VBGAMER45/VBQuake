Attribute VB_Name = "mathlib"
' mathlib.bas - math primitives (mathlib.c, and the macros of mathlib.h).
' A vec3_t is a UDT with v(0 To 2); a float m[3][3] or [3][4] is a Single
' array with the same bounds.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public vec3_origin As vec3_t             ' never written
Public Const nanmask As Long = &H7F800000    ' 255<<23

'-----------------------------------------------------------------
' mathlib.h macros

Public Function DotProduct(x As vec3_t, y As vec3_t) As Single
    DotProduct = x.v(0) * y.v(0) + x.v(1) * y.v(1) + x.v(2) * y.v(2)
End Function

Public Sub VectorSubtract(a As vec3_t, b As vec3_t, c As vec3_t)
    c.v(0) = a.v(0) - b.v(0)
    c.v(1) = a.v(1) - b.v(1)
    c.v(2) = a.v(2) - b.v(2)
End Sub

Public Sub VectorAdd(a As vec3_t, b As vec3_t, c As vec3_t)
    c.v(0) = a.v(0) + b.v(0)
    c.v(1) = a.v(1) + b.v(1)
    c.v(2) = a.v(2) + b.v(2)
End Sub

Public Sub VectorCopy(a As vec3_t, b As vec3_t)
    b.v(0) = a.v(0)
    b.v(1) = a.v(1)
    b.v(2) = a.v(2)
End Sub

' C library atan2 (VB has only Atn)
Public Function atan2(ByVal y As Double, ByVal x As Double) As Double
    If x > 0 Then
        atan2 = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then atan2 = Atn(y / x) + M_PI Else atan2 = Atn(y / x) - M_PI
    Else
        If y > 0 Then
            atan2 = M_PI / 2
        ElseIf y < 0 Then
            atan2 = -M_PI / 2
        Else
            atan2 = 0
        End If
    End If
End Function

Public Function IS_NAN(ByVal x As Single) As Boolean
    Dim l As Long
    CopyMemory l, x, 4
    IS_NAN = ((l And nanmask) = nanmask)
End Function

'-----------------------------------------------------------------

Private Function DEG2RAD(ByVal a As Single) As Single
    DEG2RAD = (a * M_PI) / 180!
End Function

Public Sub ProjectPointOnPlane(dst As vec3_t, p As vec3_t, normal As vec3_t)
    Dim d As Single, n As vec3_t, inv_denom As Single

    inv_denom = 1! / DotProduct(normal, normal)

    d = DotProduct(normal, p) * inv_denom

    n.v(0) = normal.v(0) * inv_denom
    n.v(1) = normal.v(1) * inv_denom
    n.v(2) = normal.v(2) * inv_denom

    dst.v(0) = p.v(0) - d * n.v(0)
    dst.v(1) = p.v(1) - d * n.v(1)
    dst.v(2) = p.v(2) - d * n.v(2)
End Sub

' assumes "src" is normalized
Public Sub PerpendicularVector(dst As vec3_t, src As vec3_t)
    Dim pos As Long, i As Long, minelem As Single, tempvec As vec3_t

    minelem = 1!

    ' find the smallest magnitude axially aligned vector
    pos = 0
    For i = 0 To 2
        If Abs(src.v(i)) < minelem Then
            pos = i
            minelem = Abs(src.v(i))
        End If
    Next
    tempvec.v(0) = 0!: tempvec.v(1) = 0!: tempvec.v(2) = 0!
    tempvec.v(pos) = 1!

    ' project the point onto the plane defined by src
    ProjectPointOnPlane dst, tempvec, src

    ' normalize the result
    VectorNormalize dst
End Sub

Public Sub RotatePointAroundVector(dst As vec3_t, dir_ As vec3_t, point_ As vec3_t, ByVal degrees As Single)
    Dim m(0 To 2, 0 To 2) As Single, im(0 To 2, 0 To 2) As Single
    Dim zrot(0 To 2, 0 To 2) As Single, tmpmat(0 To 2, 0 To 2) As Single
    Dim rot(0 To 2, 0 To 2) As Single
    Dim i As Long, j As Long, vr As vec3_t, vup_ As vec3_t, vf As vec3_t

    vf.v(0) = dir_.v(0)
    vf.v(1) = dir_.v(1)
    vf.v(2) = dir_.v(2)

    PerpendicularVector vr, dir_
    CrossProduct vr, vf, vup_

    m(0, 0) = vr.v(0)
    m(1, 0) = vr.v(1)
    m(2, 0) = vr.v(2)

    m(0, 1) = vup_.v(0)
    m(1, 1) = vup_.v(1)
    m(2, 1) = vup_.v(2)

    m(0, 2) = vf.v(0)
    m(1, 2) = vf.v(1)
    m(2, 2) = vf.v(2)

    For i = 0 To 2
        For j = 0 To 2
            im(i, j) = m(i, j)
        Next
    Next

    im(0, 1) = m(1, 0)
    im(0, 2) = m(2, 0)
    im(1, 0) = m(0, 1)
    im(1, 2) = m(2, 1)
    im(2, 0) = m(0, 2)
    im(2, 1) = m(1, 2)

    zrot(2, 2) = 1!
    zrot(0, 0) = Cos(DEG2RAD(degrees))
    zrot(0, 1) = Sin(DEG2RAD(degrees))
    zrot(1, 0) = -Sin(DEG2RAD(degrees))
    zrot(1, 1) = Cos(DEG2RAD(degrees))

    R_ConcatRotations m, zrot, tmpmat
    R_ConcatRotations tmpmat, im, rot

    For i = 0 To 2
        dst.v(i) = rot(i, 0) * point_.v(0) + rot(i, 1) * point_.v(1) + rot(i, 2) * point_.v(2)
    Next
End Sub

'-----------------------------------------------------------------

Public Function anglemod(ByVal a As Single) As Single
    a = (360# / 65536) * (CLng(Fix(a * (65536 / 360#))) And 65535)
    anglemod = a
End Function

' Split out like this for ASM to call.
Private Sub BOPS_Error()
    Sys_Error "BoxOnPlaneSide:  Bad signbits"
End Sub

' Returns 1, 2, or 1 + 2
Public Function BoxOnPlaneSide(emins As vec3_t, emaxs As vec3_t, p As mplane_t) As Long
    Dim dist1 As Single, dist2 As Single, sides As Long

    ' general case
    Select Case p.signbits
    Case 0
        dist1 = p.normal.v(0) * emaxs.v(0) + p.normal.v(1) * emaxs.v(1) + p.normal.v(2) * emaxs.v(2)
        dist2 = p.normal.v(0) * emins.v(0) + p.normal.v(1) * emins.v(1) + p.normal.v(2) * emins.v(2)
    Case 1
        dist1 = p.normal.v(0) * emins.v(0) + p.normal.v(1) * emaxs.v(1) + p.normal.v(2) * emaxs.v(2)
        dist2 = p.normal.v(0) * emaxs.v(0) + p.normal.v(1) * emins.v(1) + p.normal.v(2) * emins.v(2)
    Case 2
        dist1 = p.normal.v(0) * emaxs.v(0) + p.normal.v(1) * emins.v(1) + p.normal.v(2) * emaxs.v(2)
        dist2 = p.normal.v(0) * emins.v(0) + p.normal.v(1) * emaxs.v(1) + p.normal.v(2) * emins.v(2)
    Case 3
        dist1 = p.normal.v(0) * emins.v(0) + p.normal.v(1) * emins.v(1) + p.normal.v(2) * emaxs.v(2)
        dist2 = p.normal.v(0) * emaxs.v(0) + p.normal.v(1) * emaxs.v(1) + p.normal.v(2) * emins.v(2)
    Case 4
        dist1 = p.normal.v(0) * emaxs.v(0) + p.normal.v(1) * emaxs.v(1) + p.normal.v(2) * emins.v(2)
        dist2 = p.normal.v(0) * emins.v(0) + p.normal.v(1) * emins.v(1) + p.normal.v(2) * emaxs.v(2)
    Case 5
        dist1 = p.normal.v(0) * emins.v(0) + p.normal.v(1) * emaxs.v(1) + p.normal.v(2) * emins.v(2)
        dist2 = p.normal.v(0) * emaxs.v(0) + p.normal.v(1) * emins.v(1) + p.normal.v(2) * emaxs.v(2)
    Case 6
        dist1 = p.normal.v(0) * emaxs.v(0) + p.normal.v(1) * emins.v(1) + p.normal.v(2) * emins.v(2)
        dist2 = p.normal.v(0) * emins.v(0) + p.normal.v(1) * emaxs.v(1) + p.normal.v(2) * emaxs.v(2)
    Case 7
        dist1 = p.normal.v(0) * emins.v(0) + p.normal.v(1) * emins.v(1) + p.normal.v(2) * emins.v(2)
        dist2 = p.normal.v(0) * emaxs.v(0) + p.normal.v(1) * emaxs.v(1) + p.normal.v(2) * emaxs.v(2)
    Case Else
        dist1 = 0
        dist2 = 0                        ' shut up compiler
        BOPS_Error
    End Select

    sides = 0
    If dist1 >= p.dist Then sides = 1
    If dist2 < p.dist Then sides = sides Or 2

    BoxOnPlaneSide = sides
End Function

' C: the BOX_ON_PLANE_SIDE macro (fast axial cases first)
Public Function BOX_ON_PLANE_SIDE(emins As vec3_t, emaxs As vec3_t, p As mplane_t) As Long
    If p.typ < 3 Then
        If p.dist <= emins.v(p.typ) Then
            BOX_ON_PLANE_SIDE = 1
        ElseIf p.dist >= emaxs.v(p.typ) Then
            BOX_ON_PLANE_SIDE = 2
        Else
            BOX_ON_PLANE_SIDE = 3
        End If
    Else
        BOX_ON_PLANE_SIDE = BoxOnPlaneSide(emins, emaxs, p)
    End If
End Function

Public Sub AngleVectors(angles As vec3_t, forward As vec3_t, right_ As vec3_t, up As vec3_t)
    Dim angle As Single, sr As Single, sp As Single, sy As Single, cr As Single, cp As Single, cy As Single

    angle = angles.v(YAW) * (M_PI * 2 / 360)
    sy = Sin(angle)
    cy = Cos(angle)
    angle = angles.v(PITCH) * (M_PI * 2 / 360)
    sp = Sin(angle)
    cp = Cos(angle)
    angle = angles.v(ROLL) * (M_PI * 2 / 360)
    sr = Sin(angle)
    cr = Cos(angle)

    forward.v(0) = cp * cy
    forward.v(1) = cp * sy
    forward.v(2) = -sp
    right_.v(0) = (-1 * sr * sp * cy + -1 * cr * -sy)
    right_.v(1) = (-1 * sr * sp * sy + -1 * cr * cy)
    right_.v(2) = -1 * sr * cp
    up.v(0) = (cr * sp * cy + -sr * -sy)
    up.v(1) = (cr * sp * sy + -sr * cy)
    up.v(2) = cr * cp
End Sub

Public Function VectorCompare(v1 As vec3_t, v2 As vec3_t) As Long
    Dim i As Long
    For i = 0 To 2
        If v1.v(i) <> v2.v(i) Then
            VectorCompare = 0
            Exit Function
        End If
    Next
    VectorCompare = 1
End Function

Public Sub VectorMA(veca As vec3_t, ByVal scale_ As Single, vecb As vec3_t, vecc As vec3_t)
    vecc.v(0) = veca.v(0) + scale_ * vecb.v(0)
    vecc.v(1) = veca.v(1) + scale_ * vecb.v(1)
    vecc.v(2) = veca.v(2) + scale_ * vecb.v(2)
End Sub

Public Sub CrossProduct(v1 As vec3_t, v2 As vec3_t, cross As vec3_t)
    cross.v(0) = v1.v(1) * v2.v(2) - v1.v(2) * v2.v(1)
    cross.v(1) = v1.v(2) * v2.v(0) - v1.v(0) * v2.v(2)
    cross.v(2) = v1.v(0) * v2.v(1) - v1.v(1) * v2.v(0)
End Sub

Public Function Length(v As vec3_t) As Single
    Dim i As Long, l As Single
    l = 0
    For i = 0 To 2
        l = l + v.v(i) * v.v(i)
    Next
    l = Sqr(l)                           ' FIXME
    Length = l
End Function

' returns vector length
Public Function VectorNormalize(v As vec3_t) As Single
    Dim l As Single, ilength As Single

    l = v.v(0) * v.v(0) + v.v(1) * v.v(1) + v.v(2) * v.v(2)
    l = Sqr(l)                           ' FIXME

    If l <> 0 Then
        ilength = 1 / l
        v.v(0) = v.v(0) * ilength
        v.v(1) = v.v(1) * ilength
        v.v(2) = v.v(2) * ilength
    End If

    VectorNormalize = l
End Function

Public Sub VectorInverse(v As vec3_t)
    v.v(0) = -v.v(0)
    v.v(1) = -v.v(1)
    v.v(2) = -v.v(2)
End Sub

Public Sub VectorScale(in_ As vec3_t, ByVal scale_ As Single, out As vec3_t)
    out.v(0) = in_.v(0) * scale_
    out.v(1) = in_.v(1) * scale_
    out.v(2) = in_.v(2) * scale_
End Sub

Public Function Q_log2(ByVal val As Long) As Long
    Dim answer As Long
    Do
        val = val \ 2
        If val = 0 Then Exit Do
        answer = answer + 1
    Loop
    Q_log2 = answer
End Function

Public Sub R_ConcatRotations(in1() As Single, in2() As Single, out() As Single)
    out(0, 0) = in1(0, 0) * in2(0, 0) + in1(0, 1) * in2(1, 0) + in1(0, 2) * in2(2, 0)
    out(0, 1) = in1(0, 0) * in2(0, 1) + in1(0, 1) * in2(1, 1) + in1(0, 2) * in2(2, 1)
    out(0, 2) = in1(0, 0) * in2(0, 2) + in1(0, 1) * in2(1, 2) + in1(0, 2) * in2(2, 2)
    out(1, 0) = in1(1, 0) * in2(0, 0) + in1(1, 1) * in2(1, 0) + in1(1, 2) * in2(2, 0)
    out(1, 1) = in1(1, 0) * in2(0, 1) + in1(1, 1) * in2(1, 1) + in1(1, 2) * in2(2, 1)
    out(1, 2) = in1(1, 0) * in2(0, 2) + in1(1, 1) * in2(1, 2) + in1(1, 2) * in2(2, 2)
    out(2, 0) = in1(2, 0) * in2(0, 0) + in1(2, 1) * in2(1, 0) + in1(2, 2) * in2(2, 0)
    out(2, 1) = in1(2, 0) * in2(0, 1) + in1(2, 1) * in2(1, 1) + in1(2, 2) * in2(2, 1)
    out(2, 2) = in1(2, 0) * in2(0, 2) + in1(2, 1) * in2(1, 2) + in1(2, 2) * in2(2, 2)
End Sub

Public Sub R_ConcatTransforms(in1() As Single, in2() As Single, out() As Single)
    out(0, 0) = in1(0, 0) * in2(0, 0) + in1(0, 1) * in2(1, 0) + in1(0, 2) * in2(2, 0)
    out(0, 1) = in1(0, 0) * in2(0, 1) + in1(0, 1) * in2(1, 1) + in1(0, 2) * in2(2, 1)
    out(0, 2) = in1(0, 0) * in2(0, 2) + in1(0, 1) * in2(1, 2) + in1(0, 2) * in2(2, 2)
    out(0, 3) = in1(0, 0) * in2(0, 3) + in1(0, 1) * in2(1, 3) + in1(0, 2) * in2(2, 3) + in1(0, 3)
    out(1, 0) = in1(1, 0) * in2(0, 0) + in1(1, 1) * in2(1, 0) + in1(1, 2) * in2(2, 0)
    out(1, 1) = in1(1, 0) * in2(0, 1) + in1(1, 1) * in2(1, 1) + in1(1, 2) * in2(2, 1)
    out(1, 2) = in1(1, 0) * in2(0, 2) + in1(1, 1) * in2(1, 2) + in1(1, 2) * in2(2, 2)
    out(1, 3) = in1(1, 0) * in2(0, 3) + in1(1, 1) * in2(1, 3) + in1(1, 2) * in2(2, 3) + in1(1, 3)
    out(2, 0) = in1(2, 0) * in2(0, 0) + in1(2, 1) * in2(1, 0) + in1(2, 2) * in2(2, 0)
    out(2, 1) = in1(2, 0) * in2(0, 1) + in1(2, 1) * in2(1, 1) + in1(2, 2) * in2(2, 1)
    out(2, 2) = in1(2, 0) * in2(0, 2) + in1(2, 1) * in2(1, 2) + in1(2, 2) * in2(2, 2)
    out(2, 3) = in1(2, 0) * in2(0, 3) + in1(2, 1) * in2(1, 3) + in1(2, 2) * in2(2, 3) + in1(2, 3)
End Sub

' Returns mathematically correct (floor-based) quotient and remainder for
' numer and denom, both of which should contain no fractional part. The
' quotient must fit in 32 bits.
Public Sub FloorDivMod(ByVal numer As Double, ByVal denom As Double, quotient As Long, rem_ As Long)
    Dim q As Long, r As Long, x As Double

    If denom <= 0# Then Sys_Error "FloorDivMod: bad denominator %d\n", denom

    If numer >= 0# Then
        x = Int(numer / denom)
        q = x
        r = Int(numer - (x * denom))
    Else
        '
        ' perform operations with positive values, and fix mod to make floor-based
        '
        x = Int(-numer / denom)
        q = -x
        r = Int(-numer - (x * denom))
        If r <> 0 Then
            q = q - 1
            r = Fix(denom) - r
        End If
    End If

    quotient = q
    rem_ = r
End Sub

Public Function GreatestCommonDivisor(ByVal i1 As Long, ByVal i2 As Long) As Long
    If i1 > i2 Then
        If i2 = 0 Then
            GreatestCommonDivisor = i1
            Exit Function
        End If
        GreatestCommonDivisor = GreatestCommonDivisor(i2, i1 Mod i2)
    Else
        If i1 = 0 Then
            GreatestCommonDivisor = i2
            Exit Function
        End If
        GreatestCommonDivisor = GreatestCommonDivisor(i1, i2 Mod i1)
    End If
End Function

' Inverts an 8.24 value to a 16.16 value
Public Function Invert24To16(ByVal val As Long) As Long
    Dim d As Double
    If val < 256 Then
        Invert24To16 = &HFFFFFFFF
        Exit Function
    End If
    d = Fix((65536# * 16777216# / val) + 0.5)
    ' an x86 conversion of an out of range value gives 0x80000000
    If d > 2147483647# Then Invert24To16 = &H80000000 Else Invert24To16 = d
End Function

' (QuakeVB) C's 1 << n for n = 0 .. 31 (2 ^ 31 overflows a Long)
Public Function BIT_(ByVal n As Long) As Long
    If n = 31 Then
        BIT_ = &H80000000
    Else
        BIT_ = 2 ^ n
    End If
End Function

' (QuakeVB) C's x >> n of a signed int: an arithmetic shift, which floors
' (VB's \ truncates toward zero). Inline (x And &HFFFFFFF0) \ 16 where it is hot.
Public Function SAR(ByVal x As Long, ByVal n As Long) As Long
    Dim p As Long
    If n <= 0 Then
        SAR = x
    ElseIf n >= 31 Then
        If x < 0 Then SAR = -1 Else SAR = 0
    Else
        p = BIT_(n)
        SAR = (x And Not (p - 1)) \ p
    End If
End Function
