Attribute VB_Name = "pr_exec"
' pr_exec.bas - the QuakeC interpreter (pr_exec.c).
'
' The globals are pr_gl() / pr_gf() (pr_edict.bas): an operand a, b or c
' is a global's index (C: (eval_t *)&pr_globals[st->a]). Stores and returns
' copy the raw Longs, as C copies the union. A field load reads ed_l() at
' the edict's byte offset / 4 plus the field's offset.
'
' Before a builtin runs, before the bitwise operators convert floats to
' ints, and when control goes back to the engine, Sys_ClearFPU clears the
' FPU's exception flags: a division by zero or a NaN in QuakeC is harmless
' in C, but its flag would make VB's next float to integer conversion fail.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Type prstack_t
    s As Long
    f As Long                            ' pr_functions() index, -1 = NULL
End Type

Private Const MAX_STACK_DEPTH As Long = 32
Private pr_stack(0 To MAX_STACK_DEPTH) As prstack_t
Public pr_depth As Long

Private Const LOCALSTACK_SIZE As Long = 2048
Private localstack(0 To LOCALSTACK_SIZE - 1) As Long
Private localstack_used As Long

Public pr_trace As Boolean
Public pr_xfunction As Long              ' pr_functions() index, -1 = NULL
Public pr_xstatement As Long
Public pr_argc As Long

Private Const OP_DONE As Long = 0
Private Const OP_MUL_F As Long = 1
Private Const OP_MUL_V As Long = 2
Private Const OP_MUL_FV As Long = 3
Private Const OP_MUL_VF As Long = 4
Private Const OP_DIV_F As Long = 5
Private Const OP_ADD_F As Long = 6
Private Const OP_ADD_V As Long = 7
Private Const OP_SUB_F As Long = 8
Private Const OP_SUB_V As Long = 9
Private Const OP_EQ_F As Long = 10
Private Const OP_EQ_V As Long = 11
Private Const OP_EQ_S As Long = 12
Private Const OP_EQ_E As Long = 13
Private Const OP_EQ_FNC As Long = 14
Private Const OP_NE_F As Long = 15
Private Const OP_NE_V As Long = 16
Private Const OP_NE_S As Long = 17
Private Const OP_NE_E As Long = 18
Private Const OP_NE_FNC As Long = 19
Private Const OP_LE As Long = 20
Private Const OP_GE As Long = 21
Private Const OP_LT As Long = 22
Private Const OP_GT As Long = 23
Private Const OP_LOAD_F As Long = 24
Private Const OP_LOAD_V As Long = 25
Private Const OP_LOAD_S As Long = 26
Private Const OP_LOAD_ENT As Long = 27
Private Const OP_LOAD_FLD As Long = 28
Private Const OP_LOAD_FNC As Long = 29
Private Const OP_ADDRESS As Long = 30
Private Const OP_STORE_F As Long = 31
Private Const OP_STORE_V As Long = 32
Private Const OP_STORE_S As Long = 33
Private Const OP_STORE_ENT As Long = 34
Private Const OP_STORE_FLD As Long = 35
Private Const OP_STORE_FNC As Long = 36
Private Const OP_STOREP_F As Long = 37
Private Const OP_STOREP_V As Long = 38
Private Const OP_STOREP_S As Long = 39
Private Const OP_STOREP_ENT As Long = 40
Private Const OP_STOREP_FLD As Long = 41
Private Const OP_STOREP_FNC As Long = 42
Private Const OP_RETURN As Long = 43
Private Const OP_NOT_F As Long = 44
Private Const OP_NOT_V As Long = 45
Private Const OP_NOT_S As Long = 46
Private Const OP_NOT_ENT As Long = 47
Private Const OP_NOT_FNC As Long = 48
Private Const OP_IF As Long = 49
Private Const OP_IFNOT As Long = 50
Private Const OP_CALL0 As Long = 51
Private Const OP_CALL8 As Long = 59
Private Const OP_STATE As Long = 60
Private Const OP_GOTO As Long = 61
Private Const OP_AND As Long = 62
Private Const OP_OR As Long = 63
Private Const OP_BITAND As Long = 64
Private Const OP_BITOR As Long = 65

Private pr_opnames As Variant

'=============================================================================

Private Sub InitOpnames()
    pr_opnames = Array("DONE", "MUL_F", "MUL_V", "MUL_FV", "MUL_VF", "DIV", "ADD_F", "ADD_V", "SUB_F", "SUB_V", _
        "EQ_F", "EQ_V", "EQ_S", "EQ_E", "EQ_FNC", "NE_F", "NE_V", "NE_S", "NE_E", "NE_FNC", _
        "LE", "GE", "LT", "GT", "INDIRECT", "INDIRECT", "INDIRECT", "INDIRECT", "INDIRECT", "INDIRECT", _
        "ADDRESS", "STORE_F", "STORE_V", "STORE_S", "STORE_ENT", "STORE_FLD", "STORE_FNC", _
        "STOREP_F", "STOREP_V", "STOREP_S", "STOREP_ENT", "STOREP_FLD", "STOREP_FNC", _
        "RETURN", "NOT_F", "NOT_V", "NOT_S", "NOT_ENT", "NOT_FNC", "IF", "IFNOT", _
        "CALL0", "CALL1", "CALL2", "CALL3", "CALL4", "CALL5", "CALL6", "CALL7", "CALL8", _
        "STATE", "GOTO", "AND", "OR", "BITAND", "BITOR")
End Sub

Public Sub PR_PrintStatement(ByVal s As Long)
    Dim i As Long, op As Long

    If IsEmpty(pr_opnames) Then InitOpnames

    op = pr_statements(s).op
    If op >= 0 And op <= UBound(pr_opnames) Then
        Con_Printf "%s ", pr_opnames(op)
        i = Len(pr_opnames(op))
        Do While i < 10
            Con_Printf " "
            i = i + 1
        Loop
    End If

    If op = OP_IF Or op = OP_IFNOT Then
        Con_Printf "%sbranch %i", PR_GlobalString(pr_statements(s).a), pr_statements(s).b
    ElseIf op = OP_GOTO Then
        Con_Printf "branch %i", pr_statements(s).a
    ElseIf op - OP_STORE_F >= 0 And op - OP_STORE_F < 6 Then
        Con_Printf "%s", PR_GlobalString(pr_statements(s).a)
        Con_Printf "%s", PR_GlobalStringNoContents(pr_statements(s).b)
    Else
        If pr_statements(s).a Then Con_Printf "%s", PR_GlobalString(pr_statements(s).a)
        If pr_statements(s).b Then Con_Printf "%s", PR_GlobalString(pr_statements(s).b)
        If pr_statements(s).c Then Con_Printf "%s", PR_GlobalStringNoContents(pr_statements(s).c)
    End If
    Con_Printf "\n"
End Sub

Public Sub PR_StackTrace()
    Dim f As Long, i As Long

    If pr_depth = 0 Then
        Con_Printf "<NO STACK>\n"
        Exit Sub
    End If

    pr_stack(pr_depth).f = pr_xfunction
    For i = pr_depth To 0 Step -1
        f = pr_stack(i).f

        If f = -1 Then
            Con_Printf "<NO FUNCTION>\n"
        Else
            Con_Printf "%12s : %s\n", PR_GetString(pr_functions(f).s_file), PR_GetString(pr_functions(f).s_name)
        End If
    Next
End Sub

Public Sub PR_Profile_f()
    Dim f As Long, best As Long, max As Long, num As Long, i As Long

    num = 0
    Do
        max = 0
        best = -1
        For i = 0 To progs_numfunctions - 1
            f = i
            If pr_functions(f).profile > max Then
                max = pr_functions(f).profile
                best = f
            End If
        Next
        If best <> -1 Then
            If num < 10 Then Con_Printf "%7i %s\n", pr_functions(best).profile, PR_GetString(pr_functions(best).s_name)
            num = num + 1
            pr_functions(best).profile = 0
        End If
    Loop While best <> -1
End Sub

' Aborts the currently executing function
Public Sub PR_RunError(ByVal error_ As String, ParamArray args() As Variant)
    Dim a As Variant, s As String
    a = args
    s = vsprintf(error_, a)

    PR_PrintStatement pr_xstatement
    PR_StackTrace
    Con_Printf "%s\n", s

    pr_depth = 0                         ' dump the stack so host_error can shutdown functions

    Host_Error "Program error"
End Sub

'============================================================================
' PR_ExecuteProgram
' The interpretation main loop

' Returns the new program statement counter
Private Function PR_EnterFunction(ByVal f As Long) As Long
    Dim i As Long, j As Long, c As Long, o As Long

    pr_stack(pr_depth).s = pr_xstatement
    pr_stack(pr_depth).f = pr_xfunction
    pr_depth = pr_depth + 1
    If pr_depth >= MAX_STACK_DEPTH Then PR_RunError "stack overflow"

    ' save off any locals that the new function steps on
    c = pr_functions(f).locals
    If localstack_used + c > LOCALSTACK_SIZE Then PR_RunError "PR_ExecuteProgram: locals stack overflow\n"

    For i = 0 To c - 1
        localstack(localstack_used + i) = pr_gl(pr_functions(f).parm_start + i)
    Next
    localstack_used = localstack_used + c

    ' copy parameters
    o = pr_functions(f).parm_start
    For i = 0 To pr_functions(f).numparms - 1
        For j = 0 To pr_functions(f).parm_size(i) - 1
            pr_gl(o) = pr_gl(OFS_PARM0 + i * 3 + j)
            o = o + 1
        Next
    Next

    pr_xfunction = f
    PR_EnterFunction = pr_functions(f).first_statement - 1   ' offset the s++
End Function

Private Function PR_LeaveFunction() As Long
    Dim i As Long, c As Long

    If pr_depth <= 0 Then Sys_Error "prog stack underflow"

    ' restore locals from the stack
    c = pr_functions(pr_xfunction).locals
    localstack_used = localstack_used - c
    If localstack_used < 0 Then PR_RunError "PR_ExecuteProgram: locals stack underflow\n"

    For i = 0 To c - 1
        pr_gl(pr_functions(pr_xfunction).parm_start + i) = localstack(localstack_used + i)
    Next

    ' up stack
    pr_depth = pr_depth - 1
    pr_xfunction = pr_stack(pr_depth).f
    PR_LeaveFunction = pr_stack(pr_depth).s
End Function

Public Sub PR_ExecuteProgram(ByVal fnum As Long)
    Dim a As Long, b As Long, c As Long, s As Long, op As Long
    Dim newf As Long, runaway As Long, i As Long, ed As Long, exitdepth As Long, ptr As Long

    If fnum = 0 Or fnum >= progs_numfunctions Then
        If pr_global_struct(0).self Then ED_Print PROG_TO_EDICT(pr_global_struct(0).self)
        Host_Error "PR_ExecuteProgram: NULL function"
    End If

    runaway = 100000
    pr_trace = False

    ' make a stack frame
    exitdepth = pr_depth

    s = PR_EnterFunction(fnum)

    Do
        s = s + 1                        ' next statement

        op = pr_statements(s).op
        a = pr_statements(s).a
        b = pr_statements(s).b
        c = pr_statements(s).c

        runaway = runaway - 1
        If runaway = 0 Then PR_RunError "runaway loop error"

        pr_functions(pr_xfunction).profile = pr_functions(pr_xfunction).profile + 1
        pr_xstatement = s

        If pr_trace Then PR_PrintStatement s

        Select Case op
        Case OP_ADD_F
            pr_gf(c) = pr_gf(a) + pr_gf(b)
        Case OP_ADD_V
            pr_gf(c) = pr_gf(a) + pr_gf(b)
            pr_gf(c + 1) = pr_gf(a + 1) + pr_gf(b + 1)
            pr_gf(c + 2) = pr_gf(a + 2) + pr_gf(b + 2)

        Case OP_SUB_F
            pr_gf(c) = pr_gf(a) - pr_gf(b)
        Case OP_SUB_V
            pr_gf(c) = pr_gf(a) - pr_gf(b)
            pr_gf(c + 1) = pr_gf(a + 1) - pr_gf(b + 1)
            pr_gf(c + 2) = pr_gf(a + 2) - pr_gf(b + 2)

        Case OP_MUL_F
            pr_gf(c) = pr_gf(a) * pr_gf(b)
        Case OP_MUL_V
            pr_gf(c) = pr_gf(a) * pr_gf(b) + pr_gf(a + 1) * pr_gf(b + 1) + pr_gf(a + 2) * pr_gf(b + 2)
        Case OP_MUL_FV
            pr_gf(c) = pr_gf(a) * pr_gf(b)
            pr_gf(c + 1) = pr_gf(a) * pr_gf(b + 1)
            pr_gf(c + 2) = pr_gf(a) * pr_gf(b + 2)
        Case OP_MUL_VF
            pr_gf(c) = pr_gf(b) * pr_gf(a)
            pr_gf(c + 1) = pr_gf(b) * pr_gf(a + 1)
            pr_gf(c + 2) = pr_gf(b) * pr_gf(a + 2)

        Case OP_DIV_F
            pr_gf(c) = pr_gf(a) / pr_gf(b)

        Case OP_BITAND
            Sys_ClearFPU
            pr_gf(c) = ToInt(pr_gf(a)) And ToInt(pr_gf(b))
        Case OP_BITOR
            Sys_ClearFPU
            pr_gf(c) = ToInt(pr_gf(a)) Or ToInt(pr_gf(b))

        Case OP_GE
            pr_gf(c) = -(pr_gf(a) >= pr_gf(b))
        Case OP_LE
            pr_gf(c) = -(pr_gf(a) <= pr_gf(b))
        Case OP_GT
            pr_gf(c) = -(pr_gf(a) > pr_gf(b))
        Case OP_LT
            pr_gf(c) = -(pr_gf(a) < pr_gf(b))
        Case OP_AND
            pr_gf(c) = -((pr_gf(a) <> 0) And (pr_gf(b) <> 0))
        Case OP_OR
            pr_gf(c) = -((pr_gf(a) <> 0) Or (pr_gf(b) <> 0))

        Case OP_NOT_F
            pr_gf(c) = -(pr_gf(a) = 0)
        Case OP_NOT_V
            pr_gf(c) = -((pr_gf(a) = 0) And (pr_gf(a + 1) = 0) And (pr_gf(a + 2) = 0))
        Case OP_NOT_S
            If pr_gl(a) = 0 Then
                pr_gf(c) = 1
            Else
                pr_gf(c) = -(pr_strings(pr_gl(a)) = 0)
            End If
        Case OP_NOT_FNC
            pr_gf(c) = -(pr_gl(a) = 0)
        Case OP_NOT_ENT
            pr_gf(c) = -(pr_gl(a) = 0)

        Case OP_EQ_F
            pr_gf(c) = -(pr_gf(a) = pr_gf(b))
        Case OP_EQ_V
            pr_gf(c) = -((pr_gf(a) = pr_gf(b)) And (pr_gf(a + 1) = pr_gf(b + 1)) And (pr_gf(a + 2) = pr_gf(b + 2)))
        Case OP_EQ_S
            pr_gf(c) = -PR_StringsEqual(pr_gl(a), pr_gl(b))
        Case OP_EQ_E
            pr_gf(c) = -(pr_gl(a) = pr_gl(b))
        Case OP_EQ_FNC
            pr_gf(c) = -(pr_gl(a) = pr_gl(b))

        Case OP_NE_F
            pr_gf(c) = -(pr_gf(a) <> pr_gf(b))
        Case OP_NE_V
            pr_gf(c) = -((pr_gf(a) <> pr_gf(b)) Or (pr_gf(a + 1) <> pr_gf(b + 1)) Or (pr_gf(a + 2) <> pr_gf(b + 2)))
        Case OP_NE_S
            pr_gf(c) = -(Not PR_StringsEqual(pr_gl(a), pr_gl(b)))
        Case OP_NE_E
            pr_gf(c) = -(pr_gl(a) <> pr_gl(b))
        Case OP_NE_FNC
            pr_gf(c) = -(pr_gl(a) <> pr_gl(b))

        '==================
        Case OP_STORE_F, OP_STORE_ENT, OP_STORE_FLD, OP_STORE_S, OP_STORE_FNC   ' integers, pointers
            pr_gl(b) = pr_gl(a)
        Case OP_STORE_V
            pr_gl(b) = pr_gl(a)
            pr_gl(b + 1) = pr_gl(a + 1)
            pr_gl(b + 2) = pr_gl(a + 2)

        Case OP_STOREP_F, OP_STOREP_ENT, OP_STOREP_FLD, OP_STOREP_S, OP_STOREP_FNC  ' integers, pointers
            ptr = pr_gl(b) \ 4
            ed_l(ptr) = pr_gl(a)
        Case OP_STOREP_V
            ptr = pr_gl(b) \ 4
            ed_l(ptr) = pr_gl(a)
            ed_l(ptr + 1) = pr_gl(a + 1)
            ed_l(ptr + 2) = pr_gl(a + 2)

        Case OP_ADDRESS
            If pr_gl(a) = 0 And sv.state = ss_active Then PR_RunError "assignment to world entity"
            pr_gl(c) = pr_gl(a) + pr_gl(b) * 4

        Case OP_LOAD_F, OP_LOAD_FLD, OP_LOAD_ENT, OP_LOAD_S, OP_LOAD_FNC
            pr_gl(c) = ed_l(pr_gl(a) \ 4 + pr_gl(b))

        Case OP_LOAD_V
            ptr = pr_gl(a) \ 4 + pr_gl(b)
            pr_gl(c) = ed_l(ptr)
            pr_gl(c + 1) = ed_l(ptr + 1)
            pr_gl(c + 2) = ed_l(ptr + 2)

        '==================

        Case OP_IFNOT
            If pr_gl(a) = 0 Then s = s + b - 1   ' offset the s++

        Case OP_IF
            If pr_gl(a) <> 0 Then s = s + b - 1  ' offset the s++

        Case OP_GOTO
            s = s + a - 1                        ' offset the s++

        Case OP_CALL0 To OP_CALL8
            pr_argc = op - OP_CALL0
            If pr_gl(a) = 0 Then PR_RunError "NULL function"

            newf = pr_gl(a)

            If pr_functions(newf).first_statement < 0 Then
                ' negative statements are built in functions
                i = -pr_functions(newf).first_statement
                If i >= pr_numbuiltins Then PR_RunError "Bad builtin call number"
                Sys_ClearFPU
                PR_CallBuiltin i
            Else
                s = PR_EnterFunction(newf)
            End If

        Case OP_DONE, OP_RETURN
            pr_gl(OFS_RETURN) = pr_gl(a)
            pr_gl(OFS_RETURN + 1) = pr_gl(a + 1)
            pr_gl(OFS_RETURN + 2) = pr_gl(a + 2)

            s = PR_LeaveFunction()
            If pr_depth = exitdepth Then
                Sys_ClearFPU
                Exit Sub                 ' all done
            End If

        Case OP_STATE
            ed = PROG_TO_EDICT(pr_global_struct(0).self)
            ev(ed).nextthink = pr_global_struct(0).time + 0.1
            If pr_gf(a) <> ev(ed).frame Then ev(ed).frame = pr_gf(a)
            ev(ed).think = pr_gl(b)

        Case Else
            PR_RunError "Bad opcode %i", op
        End Select
    Loop
End Sub
