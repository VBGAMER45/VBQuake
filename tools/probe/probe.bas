Attribute VB_Name = "probe"
' Feasibility probe for a VB6 Quake port:
'  1. a Single array aliased over a Long array (QuakeC globals/edicts are a float/int union)
'  2. QuakeC-style interpreter dispatch speed (Select Case over all 66 opcodes)
'  3. D_DrawSpans8 + D_DrawZSpans speed at 320x200
'  4. float -> int conversion semantics (C truncates, CLng rounds)
Option Explicit

Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    cElements As Long
    lLbound As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal n As Long)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (arr() As Any) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (c As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (f As Currency) As Long

Private gl() As Long
Private gf() As Single
Private saF As SAFEARRAY1D

Private st_op() As Long, st_a() As Long, st_b() As Long, st_c() As Long

Private logf As Integer

' QuakeC opcodes (pr_comp.h) used by the loop
Private Const OP_MUL_F = 1, OP_ADD_F = 6, OP_LT = 22, OP_IFNOT = 50, OP_GOTO = 61, OP_DONE = 0

Private Sub LogLine(ByVal s As String)
    Print #logf, s
End Sub

Private Function Now_() As Double
    Dim c As Currency, f As Currency
    QueryPerformanceCounter c
    QueryPerformanceFrequency f
    Now_ = CDbl(c) / CDbl(f)
End Function

Private Sub AliasSingle(target() As Single, sa As SAFEARRAY1D, ByVal pData As Long, ByVal count As Long)
    Dim p As Long
    sa.cDims = 1
    sa.fFeatures = &H11            ' FADF_AUTO Or FADF_FIXEDSIZE: VB never frees or resizes it
    sa.cbElements = 4
    sa.pvData = pData
    sa.cElements = count
    sa.lLbound = 0
    p = VarPtr(sa)
    CopyMemory ByVal VarPtrArray(target), p, 4
End Sub

Private Sub UnaliasSingle(target() As Single)
    Dim p As Long
    CopyMemory ByVal VarPtrArray(target), p, 4
End Sub

Private Sub TestAlias()
    ReDim gl(0 To 1023)
    AliasSingle gf, saF, VarPtr(gl(0)), 1024
    gf(5) = 1.5
    LogLine "alias: gf(5)=1.5 -> gl(5)=&H" & Hex$(gl(5)) & IIf(gl(5) = &H3FC00000, "  OK", "  FAIL")
    gl(7) = &H40490FDB
    LogLine "alias: gl(7)=&H40490FDB -> gf(7)=" & gf(7) & IIf(Abs(gf(7) - 3.141593) < 0.00001, "  OK", "  FAIL")
    ReDim Preserve gl(0 To 1023)   ' same size: memory must stay put for the alias to remain valid
    LogLine "alias: UBound(gf)=" & UBound(gf)
End Sub

Private Sub SetStmt(ByVal i As Long, ByVal op As Long, ByVal a As Long, ByVal b As Long, ByVal c As Long)
    st_op(i) = op: st_a(i) = a: st_b(i) = b: st_c(i) = c
End Sub

Private Function RunVM(ByVal iterations As Long) As Long
    ' globals: 10 = i, 11 = n, 12 = one, 13 = two, 14 = t, 15 = cond
    Dim s As Long, a As Long, b As Long, c As Long, count As Long
    ReDim st_op(0 To 15): ReDim st_a(0 To 15): ReDim st_b(0 To 15): ReDim st_c(0 To 15)
    gf(10) = 0: gf(11) = iterations: gf(12) = 1: gf(13) = 2
    SetStmt 0, OP_ADD_F, 10, 12, 10      ' i = i + 1
    SetStmt 1, OP_MUL_F, 10, 13, 14      ' t = i * 2
    SetStmt 2, OP_LT, 10, 11, 15         ' cond = i < n
    SetStmt 3, OP_IFNOT, 15, 2, 0        ' if !cond goto 5
    SetStmt 4, OP_GOTO, -4, 0, 0         ' goto 0
    SetStmt 5, OP_DONE, 0, 0, 0
    s = -1
    Do
        s = s + 1
        count = count + 1
        a = st_a(s): b = st_b(s): c = st_c(s)
        Select Case st_op(s)
        Case 0: Exit Do
        Case 1: gf(c) = gf(a) * gf(b)
        Case 2: gf(c) = gf(a) * gf(b) + gf(a + 1) * gf(b + 1) + gf(a + 2) * gf(b + 2)
        Case 3: gf(c) = gf(a) * gf(b): gf(c + 1) = gf(a) * gf(b + 1): gf(c + 2) = gf(a) * gf(b + 2)
        Case 4: gf(c) = gf(b) * gf(a): gf(c + 1) = gf(b) * gf(a + 1): gf(c + 2) = gf(b) * gf(a + 2)
        Case 5: gf(c) = gf(a) / gf(b)
        Case 6: gf(c) = gf(a) + gf(b)
        Case 7: gf(c) = gf(a) + gf(b): gf(c + 1) = gf(a + 1) + gf(b + 1): gf(c + 2) = gf(a + 2) + gf(b + 2)
        Case 8: gf(c) = gf(a) - gf(b)
        Case 9: gf(c) = gf(a) - gf(b): gf(c + 1) = gf(a + 1) - gf(b + 1): gf(c + 2) = gf(a + 2) - gf(b + 2)
        Case 10: gf(c) = -(gf(a) = gf(b))
        Case 11: gf(c) = -((gf(a) = gf(b)) And (gf(a + 1) = gf(b + 1)) And (gf(a + 2) = gf(b + 2)))
        Case 12: gf(c) = -(gl(a) = gl(b))
        Case 13: gf(c) = -(gl(a) = gl(b))
        Case 14: gf(c) = -(gl(a) = gl(b))
        Case 15: gf(c) = -(gf(a) <> gf(b))
        Case 16: gf(c) = -((gf(a) <> gf(b)) Or (gf(a + 1) <> gf(b + 1)) Or (gf(a + 2) <> gf(b + 2)))
        Case 17: gf(c) = -(gl(a) <> gl(b))
        Case 18: gf(c) = -(gl(a) <> gl(b))
        Case 19: gf(c) = -(gl(a) <> gl(b))
        Case 20: gf(c) = -(gf(a) <= gf(b))
        Case 21: gf(c) = -(gf(a) >= gf(b))
        Case 22: gf(c) = -(gf(a) < gf(b))
        Case 23: gf(c) = -(gf(a) > gf(b))
        Case 24, 25, 26, 27, 28, 29: gl(c) = gl(gl(a) + gl(b))
        Case 30: gl(c) = gl(a) + gl(b)
        Case 31, 33, 34, 35, 36: gl(b) = gl(a)
        Case 32: gl(b) = gl(a): gl(b + 1) = gl(a + 1): gl(b + 2) = gl(a + 2)
        Case 37, 39, 40, 41, 42: gl(gl(b)) = gl(a)
        Case 38: gl(gl(b)) = gl(a): gl(gl(b) + 1) = gl(a + 1): gl(gl(b) + 2) = gl(a + 2)
        Case 43: gl(1) = gl(a): Exit Do
        Case 44: gf(c) = -(gf(a) = 0)
        Case 45: gf(c) = -((gf(a) = 0) And (gf(a + 1) = 0) And (gf(a + 2) = 0))
        Case 46, 47, 48: gf(c) = -(gl(a) = 0)
        Case 49: If gl(a) <> 0 Then s = s + b - 1
        Case 50: If gl(a) = 0 Then s = s + b - 1
        Case 51, 52, 53, 54, 55, 56, 57, 58, 59: gl(2) = c
        Case 60: gl(3) = a
        Case 61: s = s + a - 1
        Case 62: gf(c) = -((gf(a) <> 0) And (gf(b) <> 0))
        Case 63: gf(c) = -((gf(a) <> 0) Or (gf(b) <> 0))
        Case 64: gf(c) = CLng(gf(a)) And CLng(gf(b))
        Case 65: gf(c) = CLng(gf(a)) Or CLng(gf(b))
        End Select
    Loop
    RunVM = count
End Function

Private Sub TestVM()
    Dim t0 As Double, t1 As Double, n As Long
    t0 = Now_()
    n = RunVM(4000000)
    t1 = Now_()
    LogLine "vm: " & n & " statements in " & Format$(t1 - t0, "0.000") & " s = " & _
            Format$(n / (t1 - t0) / 1000000#, "0.0") & " M statements/s  (i=" & gf(10) & ")"
End Sub

Private Sub TestSpans()
    Const W As Long = 320, H As Long = 200, CACHEW As Long = 256
    Dim scrn() As Byte, zbuf() As Integer, tex() As Byte
    ReDim scrn(0 To W * H - 1): ReDim zbuf(0 To W * H - 1): ReDim tex(0 To CACHEW * 256& - 1)
    Dim i As Long
    For i = 0 To UBound(tex): tex(i) = (i * 7) And 255: Next

    Dim d_sdivzorigin As Single, d_sdivzstepu As Single, d_sdivzstepv As Single
    Dim d_tdivzorigin As Single, d_tdivzstepu As Single, d_tdivzstepv As Single
    Dim d_ziorigin As Single, d_zistepu As Single, d_zistepv As Single
    d_ziorigin = 0.01: d_zistepu = 0.000001: d_zistepv = 0.000002
    d_sdivzorigin = 0: d_sdivzstepu = 0.007: d_sdivzstepv = 0.0003
    d_tdivzorigin = 0: d_tdivzstepu = 0.0002: d_tdivzstepv = 0.011
    Dim bbextents As Long, bbextentt As Long
    bbextents = 255& * 65536: bbextentt = 255& * 65536

    Dim frame As Long, v As Long, u As Long, count As Long, spancount As Long, pdest As Long
    Dim s As Long, t As Long, snext As Long, tnext As Long, sstep As Long, tstep As Long
    Dim sdivz As Single, tdivz As Single, zi As Single, z As Single, du As Single, dv As Single
    Dim sdivz8stepu As Single, tdivz8stepu As Single, zi8stepu As Single
    Dim izi As Long, izistep As Long, zdest As Long
    Dim t0 As Double, t1 As Double
    Const FRAMES = 300

    sdivz8stepu = d_sdivzstepu * 8: tdivz8stepu = d_tdivzstepu * 8: zi8stepu = d_zistepu * 8
    t0 = Now_()
    For frame = 1 To FRAMES
        For v = 0 To H - 1
            ' D_DrawSpans8: one full-width span per row (the edge list covers the screen once)
            u = 0: count = W
            pdest = v * W + u
            du = u: dv = v
            sdivz = d_sdivzorigin + dv * d_sdivzstepv + du * d_sdivzstepu
            tdivz = d_tdivzorigin + dv * d_tdivzstepv + du * d_tdivzstepu
            zi = d_ziorigin + dv * d_zistepv + du * d_zistepu
            z = 65536! / zi
            s = Int(sdivz * z): If s > bbextents Then s = bbextents Else If s < 0 Then s = 0
            t = Int(tdivz * z): If t > bbextentt Then t = bbextentt Else If t < 0 Then t = 0
            Do
                If count >= 8 Then spancount = 8 Else spancount = count
                count = count - spancount
                If count Then
                    sdivz = sdivz + sdivz8stepu: tdivz = tdivz + tdivz8stepu: zi = zi + zi8stepu
                    z = 65536! / zi
                    snext = Int(sdivz * z): If snext > bbextents Then snext = bbextents Else If snext < 8 Then snext = 8
                    tnext = Int(tdivz * z): If tnext > bbextentt Then tnext = bbextentt Else If tnext < 8 Then tnext = 8
                    sstep = (snext - s) \ 8: tstep = (tnext - t) \ 8
                Else
                    sdivz = sdivz + d_sdivzstepu * (spancount - 1)
                    tdivz = tdivz + d_tdivzstepu * (spancount - 1)
                    zi = zi + d_zistepu * (spancount - 1)
                    z = 65536! / zi
                    snext = Int(sdivz * z): If snext > bbextents Then snext = bbextents Else If snext < 8 Then snext = 8
                    tnext = Int(tdivz * z): If tnext > bbextentt Then tnext = bbextentt Else If tnext < 8 Then tnext = 8
                    If spancount > 1 Then sstep = (snext - s) \ (spancount - 1): tstep = (tnext - t) \ (spancount - 1)
                End If
                Do
                    scrn(pdest) = tex((s \ 65536) + (t \ 65536) * CACHEW)
                    pdest = pdest + 1
                    s = s + sstep
                    t = t + tstep
                    spancount = spancount - 1
                Loop While spancount > 0
                s = snext: t = tnext
            Loop While count > 0

            ' D_DrawZSpans: 1/z in 16-bit fixed point per pixel
            zi = d_ziorigin + dv * d_zistepv
            izi = Int(zi * &H8000& * &H10000)
            izistep = Int(d_zistepu * &H8000& * &H10000)
            zdest = v * W
            For u = 0 To W - 1
                zbuf(zdest + u) = (izi \ 65536) And &H7FFF&
                izi = izi + izistep
            Next
        Next
    Next
    t1 = Now_()
    LogLine "spans: " & FRAMES & " frames of 320x200 texture + z spans in " & Format$(t1 - t0, "0.000") & _
            " s = " & Format$(FRAMES / (t1 - t0), "0") & " fps  (checksum " & scrn(12345) & "," & zbuf(23456) & ")"
End Sub

Private Sub TestConversions()
    Dim x As Single, l As Long
    x = 2.5: l = CLng(x): LogLine "CLng(2.5)=" & l & "  Int(2.5)=" & Int(x) & "  Fix(2.5)=" & Fix(x)
    x = -1.5: l = CLng(x): LogLine "CLng(-1.5)=" & l & "  Int(-1.5)=" & Int(x) & "  Fix(-1.5)=" & Fix(x) & "   (C (int)-1.5 = -1)"
    Dim t0 As Double, i As Long, acc As Long, r As Long
    x = 1234.75
    t0 = Now_(): For i = 1 To 10000000: r = CLng(x): acc = acc Xor r: Next
    LogLine "CLng:      " & Format$((Now_() - t0) * 100, "0.0") & " ns/op"
    t0 = Now_(): For i = 1 To 10000000: r = Int(x): acc = acc Xor r: Next
    LogLine "Int:       " & Format$((Now_() - t0) * 100, "0.0") & " ns/op"
    t0 = Now_(): For i = 1 To 10000000: r = Fix(x): acc = acc Xor r: Next
    LogLine "Fix:       " & Format$((Now_() - t0) * 100, "0.0") & " ns/op  (" & acc & ")"
End Sub

Sub Main()
    logf = FreeFile
    Open App.Path & "\probe_" & Command$ & ".log" For Output As #logf
    On Error GoTo fail
    TestAlias
    TestConversions
    TestVM
    TestSpans
done:
    UnaliasSingle gf
    Close #logf
    Exit Sub
fail:
    LogLine "run time error " & Err.Number & ": " & Err.Description
    Resume done
End Sub
