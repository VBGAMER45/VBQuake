Attribute VB_Name = "pr_edict"
' pr_edict.bas - entity dictionary (pr_edict.c), and the progs memory.
'
' The globals are the Long array pr_gl(); pr_gf() (Single) and
' pr_global_struct(0) (globalvars_t, progdefs.bas) are laid over the same
' memory, as C's pr_globals and pr_global_struct are one block of memory.
' The fields of all edicts are the Long array ed_l(), EDICT_STRIDE Longs
' per edict, with ed_f() and ev() (entvars_t) laid over it: ev(n).origin
' is C's EDICT_NUM(n)->v.origin. Never ReDim pr_gl() or ed_l() but through
' PR_LoadProgs / PR_AllocEdicts, which lay the views over them again.
'
' An edict as QuakeC sees it (EDICT_TO_PROG) is a byte offset into ed_l(),
' n * pr_edict_size; a field pointer (OP_ADDRESS) is a byte offset too.
'
' string_t is an offset into pr_strings(): the progs' strings, then the
' engine's strings at fixed places (PRS_*: ftos/vtos/etos's buffer, the
' map name, the world model's name, the client names; C points to engine
' memory relative to pr_strings), then ED_NewString's strings.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' progs header (dprograms_t)
Public progs_numstatements As Long
Public progs_numglobaldefs As Long
Public progs_numfielddefs As Long
Public progs_numfunctions As Long
Public progs_numstrings As Long
Public progs_numglobals As Long
Public progs_entityfields As Long

Public pr_functions() As dfunction_t
Public pr_strings() As Byte
Private pr_strings_used As Long
Public pr_globaldefs() As ddef_t
Public pr_fielddefs() As ddef_t
Public pr_statements() As dstatement_t

Public pr_gl() As Long                   ' C: (int *)pr_globals
Public pr_gf() As Single                 ' C: pr_globals
Public pr_global_struct() As globalvars_t   ' (0): C: *pr_global_struct
Public pr_edict_size As Long             ' in bytes (as the progs see it)
Public pr_crc As Long

Public ed_l() As Long                    ' the fields of all edicts
Public ed_f() As Single
Public ev() As entvars_t                 ' ev(n): C: EDICT_NUM(n)->v
Public edicts() As edict_t               ' the engine's part of each edict

Private sa_gf As SAFEARRAY1D
Private sa_pgs As SAFEARRAY1D
Private sa_ef As SAFEARRAY1D
Private sa_ev As SAFEARRAY1D
Private global_views_set As Boolean
Private edict_views_set As Boolean

' engine strings in pr_strings()
Public PRS_TEMP As Long                  ' C: pr_string_temp[128] (pr_cmds.c)
Public PRS_MAPNAME As Long               ' C: sv.name
Public PRS_WORLDMODEL As Long            ' C: sv.worldmodel->name
Public PRS_CLIENTNAMES As Long           ' C: svs.clients[i].name: + 32 * i
Private Const PRS_ENGINE_SIZE As Long = 128 + 64 + 64 + MAX_SCOREBOARD * 32

Public type_size(0 To 7) As Long

Public cv_nomonsters As Long
Public cv_gamecfg As Long
Public cv_scratch1 As Long
Public cv_scratch2 As Long
Public cv_scratch3 As Long
Public cv_scratch4 As Long
Public cv_savedgamecfg As Long
Public cv_saved1 As Long
Public cv_saved2 As Long
Public cv_saved3 As Long
Public cv_saved4 As Long

Private Const MAX_FIELD_LEN As Long = 64
Private Const GEFV_CACHESIZE As Long = 2
Private gefvCache_def(0 To GEFV_CACHESIZE - 1) As Long   ' pr_fielddefs() index, -1 = NULL
Private gefvCache_field(0 To GEFV_CACHESIZE - 1) As String
Private gefv_rep As Long

'===========================================================================
' the views over the progs memory


' Takes the views down; called before the program ends, so VB does not try
' to free memory it does not own.
Public Sub PR_ClearViews()
    ClearGlobalViews
    ClearEdictViews
End Sub

Private Sub ClearGlobalViews()
    If Not global_views_set Then Exit Sub
    ClearView VarPtrArray(pr_gf)
    ClearView VarPtrArray(pr_global_struct)
    global_views_set = False
End Sub

Private Sub ClearEdictViews()
    If Not edict_views_set Then Exit Sub
    ClearView VarPtrArray(ed_f)
    ClearView VarPtrArray(ev)
    edict_views_set = False
End Sub

Private Sub SetGlobalViews()
    Dim g As globalvars_t
    SetView VarPtrArray(pr_gf), sa_gf, VarPtr(pr_gl(0)), UBound(pr_gl) + 1, 4
    SetView VarPtrArray(pr_global_struct), sa_pgs, VarPtr(pr_gl(0)), 1, LenB(g)
    global_views_set = True
End Sub

' C: sv.edicts = Hunk_AllocName (sv.max_edicts*pr_edict_size, "edicts")
Public Sub PR_AllocEdicts(ByVal max_edicts_ As Long)
    Dim e As entvars_t, i As Long

    If LenB(e) <> EDICT_STRIDE * 4 Then Sys_Error "entvars_t is %i bytes, not %i", LenB(e), EDICT_STRIDE * 4

    ClearEdictViews
    ReDim ed_l(0 To max_edicts_ * EDICT_STRIDE - 1)
    ReDim edicts(0 To max_edicts_ - 1)
    SetView VarPtrArray(ed_f), sa_ef, VarPtr(ed_l(0)), max_edicts_ * EDICT_STRIDE, 4
    SetView VarPtrArray(ev), sa_ev, VarPtr(ed_l(0)), max_edicts_, EDICT_STRIDE * 4
    edict_views_set = True

    ' the layout must be C's: the progs address the fields by offset
    If VarPtr(ev(0).origin) - VarPtr(ev(0)) <> EOFS_origin * 4 Or _
       VarPtr(ev(1).modelindex) - VarPtr(ev(0)) <> EDICT_STRIDE * 4 Or _
       VarPtr(ev(0).noise3) - VarPtr(ev(0)) <> EOFS_noise3 * 4 Then
        Sys_Error "entvars_t does not have the layout of progdefs.h"
    End If

    For i = 0 To max_edicts_ - 1
        SV_ClearAreaLink i
    Next
End Sub

'===========================================================================
' strings

' C: pr_strings + ofs
Public Function PR_GetString(ByVal ofs As Long) As String
    PR_GetString = BufToStr(pr_strings, ofs)
End Function

' C: !*(pr_strings + ofs)
Public Function PR_StringEmpty(ByVal ofs As Long) As Boolean
    PR_StringEmpty = (pr_strings(ofs) = 0)
End Function

' C: strcmp (pr_strings + a, pr_strings + b) == 0
Public Function PR_StringsEqual(ByVal a As Long, ByVal b As Long) As Boolean
    Do
        If pr_strings(a) <> pr_strings(b) Then
            PR_StringsEqual = False
            Exit Function
        End If
        If pr_strings(a) = 0 Then
            PR_StringsEqual = True
            Exit Function
        End If
        a = a + 1
        b = b + 1
    Loop
End Function

' Writes text into an engine string slot (zero-terminated, at most size - 1 characters).
Public Sub PR_SetEngineString(ByVal ofs As Long, ByVal size As Long, ByVal text As String)
    text = Left$(text, size - 1)
    StrToBuf text, pr_strings, ofs
    pr_strings(ofs + Len(text)) = 0
End Sub

' C: client->name - pr_strings
Public Function PR_ClientNameString(ByVal client As Long) As Long
    PR_ClientNameString = PRS_CLIENTNAMES + client * 32
End Function

' Sets svs.clients[client].name, as the progs see it too (C: the netname of
' a client's edict points to its name)
Public Sub SV_SetClientName(ByVal client As Long, ByVal name As String)
    svs.clients(client).name = Left$(name, 31)
    If PRS_CLIENTNAMES Then PR_SetEngineString PRS_CLIENTNAMES + client * 32, 32, svs.clients(client).name
End Sub

'===========================================================================

' Sets everything to NULL
Public Sub ED_ClearEdict(ByVal e As Long)
    FillMemory ed_l(e * EDICT_STRIDE), progs_entityfields * 4, 0
    edicts(e).free = False
End Sub

' Either finds a free edict, or allocates a new one.
' Try to avoid reusing an entity that was recently freed, because it
' can cause the client to think the entity morphed into something else
' instead of being removed and recreated, which can cause interpolated
' angles and bad trails.
Public Function ED_Alloc() As Long
    Dim i As Long

    For i = svs.maxclients + 1 To sv.num_edicts - 1
        ' the first couple seconds of server time can involve a lot of
        ' freeing and allocating, so relax the replacement policy
        If edicts(i).free Then
            If edicts(i).freetime < 2 Or sv.time - edicts(i).freetime > 0.5 Then
                ED_ClearEdict i
                ED_Alloc = i
                Exit Function
            End If
        End If
    Next

    If i = MAX_EDICTS Then Sys_Error "ED_Alloc: no free edicts"

    sv.num_edicts = sv.num_edicts + 1
    ED_ClearEdict i

    ED_Alloc = i
End Function

' Marks the edict as free
' FIXME: walk all entities and NULL out references to this entity
Public Sub ED_Free(ByVal ed As Long)
    SV_UnlinkEdict ed                    ' unlink from world bsp

    edicts(ed).free = True
    ev(ed).model = 0
    ev(ed).takedamage = 0
    ev(ed).modelindex = 0
    ev(ed).colormap = 0
    ev(ed).skin = 0
    ev(ed).frame = 0
    ev(ed).origin = vec3_origin
    ev(ed).angles = vec3_origin
    ev(ed).nextthink = -1
    ev(ed).solid = 0

    edicts(ed).freetime = sv.time
End Sub

'===========================================================================

' pr_globaldefs() index, -1 = NULL
Public Function ED_GlobalAtOfs(ByVal ofs As Long) As Long
    Dim i As Long
    For i = 0 To progs_numglobaldefs - 1
        If pr_globaldefs(i).ofs = ofs Then
            ED_GlobalAtOfs = i
            Exit Function
        End If
    Next
    ED_GlobalAtOfs = -1
End Function

' pr_fielddefs() index, -1 = NULL
Public Function ED_FieldAtOfs(ByVal ofs As Long) As Long
    Dim i As Long
    For i = 0 To progs_numfielddefs - 1
        If pr_fielddefs(i).ofs = ofs Then
            ED_FieldAtOfs = i
            Exit Function
        End If
    Next
    ED_FieldAtOfs = -1
End Function

Public Function ED_FindField(ByVal name As String) As Long
    Dim i As Long
    For i = 0 To progs_numfielddefs - 1
        If PR_GetString(pr_fielddefs(i).s_name) = name Then
            ED_FindField = i
            Exit Function
        End If
    Next
    ED_FindField = -1
End Function

Public Function ED_FindGlobal(ByVal name As String) As Long
    Dim i As Long
    For i = 0 To progs_numglobaldefs - 1
        If PR_GetString(pr_globaldefs(i).s_name) = name Then
            ED_FindGlobal = i
            Exit Function
        End If
    Next
    ED_FindGlobal = -1
End Function

' pr_functions() index, -1 = NULL
Public Function ED_FindFunction(ByVal name As String) As Long
    Dim i As Long
    For i = 0 To progs_numfunctions - 1
        If PR_GetString(pr_functions(i).s_name) = name Then
            ED_FindFunction = i
            Exit Function
        End If
    Next
    ED_FindFunction = -1
End Function

' returns the field's index in ed_l() / ed_f(), -1 = NULL
Public Function GetEdictFieldValue(ByVal ed As Long, ByVal field As String) As Long
    Dim def As Long, i As Long

    def = -1
    For i = 0 To GEFV_CACHESIZE - 1
        If field = gefvCache_field(i) Then
            def = gefvCache_def(i)
            GoTo Done
        End If
    Next

    def = ED_FindField(field)

    If Len(field) < MAX_FIELD_LEN Then
        gefvCache_def(gefv_rep) = def
        gefvCache_field(gefv_rep) = field
        gefv_rep = gefv_rep Xor 1
    End If

Done:
    If def = -1 Then
        GetEdictFieldValue = -1
        Exit Function
    End If

    GetEdictFieldValue = ed * EDICT_STRIDE + pr_fielddefs(def).ofs
End Function

Private Function L2F(ByVal l As Long) As Single
    Dim f As Single
    CopyMemory f, l, 4
    L2F = f
End Function

' Returns a string describing the value at mem(idx) (pr_gl or ed_l) in a
' type specific manner
Public Function PR_ValueString(ByVal typ As Long, mem() As Long, ByVal idx As Long) As String
    Dim def As Long

    typ = typ And Not DEF_SAVEGLOBAL
    Select Case typ
    Case ev_string
        PR_ValueString = va("%s", PR_GetString(mem(idx)))
    Case ev_entity
        PR_ValueString = va("entity %i", NUM_FOR_EDICT(PROG_TO_EDICT(mem(idx))))
    Case ev_function
        PR_ValueString = va("%s()", PR_GetString(pr_functions(mem(idx)).s_name))
    Case ev_field
        def = ED_FieldAtOfs(mem(idx))
        PR_ValueString = va(".%s", PR_GetString(pr_fielddefs(def).s_name))
    Case ev_void
        PR_ValueString = "void"
    Case ev_float
        PR_ValueString = va("%5.1f", L2F(mem(idx)))
    Case ev_vector
        PR_ValueString = va("'%5.1f %5.1f %5.1f'", L2F(mem(idx)), L2F(mem(idx + 1)), L2F(mem(idx + 2)))
    Case ev_pointer
        PR_ValueString = "pointer"
    Case Else
        PR_ValueString = va("bad type %i", typ)
    End Select
End Function

' Returns a string describing the value in a type specific manner
' Easier to parse than PR_ValueString
Public Function PR_UglyValueString(ByVal typ As Long, mem() As Long, ByVal idx As Long) As String
    Dim def As Long

    typ = typ And Not DEF_SAVEGLOBAL
    Select Case typ
    Case ev_string
        PR_UglyValueString = va("%s", PR_GetString(mem(idx)))
    Case ev_entity
        PR_UglyValueString = va("%i", NUM_FOR_EDICT(PROG_TO_EDICT(mem(idx))))
    Case ev_function
        PR_UglyValueString = va("%s", PR_GetString(pr_functions(mem(idx)).s_name))
    Case ev_field
        def = ED_FieldAtOfs(mem(idx))
        PR_UglyValueString = va("%s", PR_GetString(pr_fielddefs(def).s_name))
    Case ev_void
        PR_UglyValueString = "void"
    Case ev_float
        PR_UglyValueString = va("%f", L2F(mem(idx)))
    Case ev_vector
        PR_UglyValueString = va("%f %f %f", L2F(mem(idx)), L2F(mem(idx + 1)), L2F(mem(idx + 2)))
    Case Else
        PR_UglyValueString = va("bad type %i", typ)
    End Select
End Function

' Returns a string with a description and the contents of a global,
' padded to 20 field width
Public Function PR_GlobalString(ByVal ofs As Long) As String
    Dim s As String, def As Long, line_ As String

    def = ED_GlobalAtOfs(ofs)
    If def = -1 Then
        line_ = va("%i(???)", ofs)
    Else
        s = PR_ValueString(pr_globaldefs(def).typ, pr_gl, ofs)
        line_ = va("%i(%s)%s", ofs, PR_GetString(pr_globaldefs(def).s_name), s)
    End If

    If Len(line_) < 20 Then line_ = line_ & Space$(20 - Len(line_))
    PR_GlobalString = line_ & " "
End Function

Public Function PR_GlobalStringNoContents(ByVal ofs As Long) As String
    Dim def As Long, line_ As String

    def = ED_GlobalAtOfs(ofs)
    If def = -1 Then
        line_ = va("%i(???)", ofs)
    Else
        line_ = va("%i(%s)", ofs, PR_GetString(pr_globaldefs(def).s_name))
    End If

    If Len(line_) < 20 Then line_ = line_ & Space$(20 - Len(line_))
    PR_GlobalStringNoContents = line_ & " "
End Function

' For debugging
Public Sub ED_Print(ByVal ed As Long)
    Dim l As Long, d As Long, v As Long, i As Long, j As Long, name As String, typ As Long

    If edicts(ed).free Then
        Con_Printf "FREE\n"
        Exit Sub
    End If

    Con_Printf "\nEDICT %i:\n", NUM_FOR_EDICT(ed)
    For i = 1 To progs_numfielddefs - 1
        d = i
        name = PR_GetString(pr_fielddefs(d).s_name)
        If Len(name) >= 2 Then
            If Mid$(name, Len(name) - 1, 1) = "_" Then GoTo nexti   ' skip _x, _y, _z vars
        End If

        v = ed * EDICT_STRIDE + pr_fielddefs(d).ofs

        ' if the value is still all 0, skip the field
        typ = pr_fielddefs(d).typ And Not DEF_SAVEGLOBAL
        For j = 0 To type_size(typ) - 1
            If ed_l(v + j) Then Exit For
        Next
        If j = type_size(typ) Then GoTo nexti

        Con_Printf "%s", name
        l = Len(name)
        Do While l < 15
            Con_Printf " "
            l = l + 1
        Loop

        Con_Printf "%s\n", PR_ValueString(pr_fielddefs(d).typ, ed_l, v)
nexti:
    Next
End Sub

' For savegames (f: an open VB file number)
Public Sub ED_Write(ByVal f As Integer, ByVal ed As Long)
    Dim d As Long, v As Long, i As Long, j As Long, name As String, typ As Long

    Print #f, "{"

    If edicts(ed).free Then
        Print #f, "}"
        Exit Sub
    End If

    For i = 1 To progs_numfielddefs - 1
        d = i
        name = PR_GetString(pr_fielddefs(d).s_name)
        If Len(name) >= 2 Then
            If Mid$(name, Len(name) - 1, 1) = "_" Then GoTo nexti   ' skip _x, _y, _z vars
        End If

        v = ed * EDICT_STRIDE + pr_fielddefs(d).ofs

        ' if the value is still all 0, skip the field
        typ = pr_fielddefs(d).typ And Not DEF_SAVEGLOBAL
        For j = 0 To type_size(typ) - 1
            If ed_l(v + j) Then Exit For
        Next
        If j = type_size(typ) Then GoTo nexti

        Print #f, va("""%s"" ""%s""", name, PR_UglyValueString(pr_fielddefs(d).typ, ed_l, v))
nexti:
    Next

    Print #f, "}"
End Sub

Public Sub ED_PrintNum(ByVal ent As Long)
    ED_Print EDICT_NUM(ent)
End Sub

' For debugging, prints all the entities in the current server
Public Sub ED_PrintEdicts()
    Dim i As Long
    Con_Printf "%i entities\n", sv.num_edicts
    For i = 0 To sv.num_edicts - 1
        ED_PrintNum i
    Next
End Sub

' For debugging, prints a single edicy
Public Sub ED_PrintEdict_f()
    Dim i As Long

    i = Q_atoi(Cmd_Argv(1))
    If i >= sv.num_edicts Then
        Con_Printf "Bad edict number\n"
        Exit Sub
    End If
    ED_PrintNum i
End Sub

' For debugging
Public Sub ED_Count()
    Dim i As Long, active As Long, models As Long, solid As Long, stepc As Long

    For i = 0 To sv.num_edicts - 1
        If Not edicts(i).free Then
            active = active + 1
            If ev(i).solid Then solid = solid + 1
            If ev(i).model Then models = models + 1
            If ev(i).movetype = MOVETYPE_STEP Then stepc = stepc + 1
        End If
    Next

    Con_Printf "num_edicts:%3i\n", sv.num_edicts
    Con_Printf "active    :%3i\n", active
    Con_Printf "view      :%3i\n", models
    Con_Printf "touch     :%3i\n", solid
    Con_Printf "step      :%3i\n", stepc
End Sub

'==============================================================================
' ARCHIVING GLOBALS
' FIXME: need to tag constants, doesn't really work

Public Sub ED_WriteGlobals(ByVal f As Integer)
    Dim i As Long, name As String, typ As Long

    Print #f, "{"
    For i = 0 To progs_numglobaldefs - 1
        typ = pr_globaldefs(i).typ
        If (typ And DEF_SAVEGLOBAL) = 0 Then GoTo nexti
        typ = typ And Not DEF_SAVEGLOBAL

        If typ <> ev_string And typ <> ev_float And typ <> ev_entity Then GoTo nexti

        name = PR_GetString(pr_globaldefs(i).s_name)
        Print #f, va("""%s"" ""%s""", name, PR_UglyValueString(typ, pr_gl, pr_globaldefs(i).ofs))
nexti:
    Next
    Print #f, "}"
End Sub

' data(p): the text after the opening brace; returns the position after the
' closing brace
Public Function ED_ParseGlobals(data() As Byte, ByVal p As Long) As Long
    Dim keyname As String, key As Long

    Do
        ' parse key
        p = COM_Parse(data, p)
        If Left$(com_token, 1) = "}" Then Exit Do
        If p = -1 Then Sys_Error "ED_ParseEntity: EOF without closing brace"

        keyname = com_token

        ' parse value
        p = COM_Parse(data, p)
        If p = -1 Then Sys_Error "ED_ParseEntity: EOF without closing brace"

        If Left$(com_token, 1) = "}" Then Sys_Error "ED_ParseEntity: closing brace without data"

        key = ED_FindGlobal(keyname)
        If key = -1 Then
            Con_Printf "'%s' is not a global\n", keyname
        Else
            If Not ED_ParseEpair(True, 0, key, com_token) Then Host_Error "ED_ParseGlobals: parse error"
        End If
    Loop

    ED_ParseGlobals = p
End Function

'============================================================================

' returns the new string's offset in pr_strings() (C: a copy on the hunk)
Public Function ED_NewString(ByVal s As String) As Long
    Dim i As Long, l As Long, b() As Byte, newp As Long, c As Long, need As Long

    StrToBytesZ s, b
    l = Len(s) + 1

    need = pr_strings_used + l
    If need > UBound(pr_strings) + 1 Then ReDim Preserve pr_strings(0 To need * 2)

    ED_NewString = pr_strings_used
    newp = pr_strings_used
    i = 0
    Do While i < l
        c = b(i)
        If c = 92 And i < l - 1 Then     ' '\\'
            i = i + 1
            If b(i) = 110 Then           ' 'n'
                pr_strings(newp) = 10
            Else
                pr_strings(newp) = 92
            End If
        Else
            pr_strings(newp) = c
        End If
        newp = newp + 1
        i = i + 1
    Loop
    pr_strings_used = newp
End Function

' Can parse either fields or globals: the value goes to pr_gl(key's ofs) or
' to ed_l(baseidx + key's ofs). returns false if error
Public Function ED_ParseEpair(ByVal isglobal As Boolean, ByVal baseidx As Long, ByVal key As Long, ByVal s As String) As Boolean
    Dim i As Long, d As Long, def As Long, func As Long, parts() As String, w As String, v As Long

    d = baseidx + pr_globaldefs_or_fields(isglobal, key)

    Select Case pr_def_type(isglobal, key) And Not DEF_SAVEGLOBAL
    Case ev_string
        SetMem isglobal, d, ED_NewString(s)

    Case ev_float
        SetMemF isglobal, d, C_atof(s)

    Case ev_vector
        ' (C copies to a 128 byte buffer and splits at spaces)
        w = Left$(s, 127)
        For i = 0 To 2
            v = InStr(w, " ")
            If v = 0 Then
                SetMemF isglobal, d + i, C_atof(w)
                w = ""
            Else
                SetMemF isglobal, d + i, C_atof(Left$(w, v - 1))
                w = Mid$(w, v + 1)
            End If
        Next

    Case ev_entity
        SetMem isglobal, d, EDICT_TO_PROG(EDICT_NUM(C_atoi(s)))

    Case ev_field
        def = ED_FindField(s)
        If def = -1 Then
            Con_Printf "Can't find field %s\n", s
            ED_ParseEpair = False
            Exit Function
        End If
        SetMem isglobal, d, pr_gl(pr_fielddefs(def).ofs)

    Case ev_function
        func = ED_FindFunction(s)
        If func = -1 Then
            Con_Printf "Can't find function %s\n", s
            ED_ParseEpair = False
            Exit Function
        End If
        SetMem isglobal, d, func
    End Select
    ED_ParseEpair = True
End Function

Private Function pr_globaldefs_or_fields(ByVal isglobal As Boolean, ByVal key As Long) As Long
    If isglobal Then pr_globaldefs_or_fields = pr_globaldefs(key).ofs Else pr_globaldefs_or_fields = pr_fielddefs(key).ofs
End Function

Private Function pr_def_type(ByVal isglobal As Boolean, ByVal key As Long) As Long
    If isglobal Then pr_def_type = pr_globaldefs(key).typ Else pr_def_type = pr_fielddefs(key).typ
End Function

Private Sub SetMem(ByVal isglobal As Boolean, ByVal idx As Long, ByVal l As Long)
    If isglobal Then pr_gl(idx) = l Else ed_l(idx) = l
End Sub

Private Sub SetMemF(ByVal isglobal As Boolean, ByVal idx As Long, ByVal f As Single)
    If isglobal Then pr_gf(idx) = f Else ed_f(idx) = f
End Sub

' Parses an edict out of the given text, returning the new position
' ed should be a properly initialized empty edict.
' Used for initial level load and for savegames.
Public Function ED_ParseEdict(data() As Byte, ByVal p As Long, ByVal ent As Long) As Long
    Dim key As Long, anglehack As Boolean, init As Boolean, keyname As String, token As String

    init = False

    ' clear it
    If ent <> 0 Then FillMemory ed_l(ent * EDICT_STRIDE), progs_entityfields * 4, 0   ' hack

    ' go through all the dictionary pairs
    Do
        ' parse key
        p = COM_Parse(data, p)
        If Left$(com_token, 1) = "}" Then Exit Do
        If p = -1 Then Sys_Error "ED_ParseEntity: EOF without closing brace"

        ' anglehack is to allow QuakeEd to write single scalar angles
        ' and allow them to be turned into vectors. (FIXME...)
        If com_token = "angle" Then
            com_token = "angles"
            anglehack = True
        Else
            anglehack = False
        End If

        ' FIXME: change light to _light to get rid of this hack
        If com_token = "light" Then com_token = "light_lev"   ' hack for single light def

        keyname = com_token

        ' another hack to fix heynames with trailing spaces
        Do While Len(keyname) > 0
            If Right$(keyname, 1) <> " " Then Exit Do
            keyname = Left$(keyname, Len(keyname) - 1)
        Loop

        ' parse value
        p = COM_Parse(data, p)
        If p = -1 Then Sys_Error "ED_ParseEntity: EOF without closing brace"

        If Left$(com_token, 1) = "}" Then Sys_Error "ED_ParseEntity: closing brace without data"

        init = True

        ' keynames with a leading underscore are used for utility comments,
        ' and are immediately discarded by quake
        If Left$(keyname, 1) = "_" Then GoTo nextpair

        key = ED_FindField(keyname)
        If key = -1 Then
            Con_Printf "'%s' is not a field\n", keyname
            GoTo nextpair
        End If

        token = com_token
        If anglehack Then token = va("0 %s 0", Left$(token, 31))

        If Not ED_ParseEpair(False, ent * EDICT_STRIDE, key, token) Then Host_Error "ED_ParseEdict: parse error"
nextpair:
    Loop

    If Not init Then edicts(ent).free = True

    ED_ParseEdict = p
End Function

' The entities are directly placed in the array, rather than allocated with
' ED_Alloc, because otherwise an error loading the map would have entity
' number references out of order.
'
' Creates a server's entity / program execution context by
' parsing textual entity definitions out of an ent file.
'
' Used for both fresh maps and savegame loads.  A fresh map would also need
' to call ED_CallSpawnFunctions () to let the objects initialize themselves.
Public Sub ED_LoadFromFile(data() As Byte, ByVal p As Long)
    Dim ent As Long, inhibit As Long, func As Long, flags As Long

    ent = -1
    inhibit = 0
    pr_global_struct(0).time = sv.time

    ' parse ents
    Do
        ' parse the opening brace
        p = COM_Parse(data, p)
        If p = -1 Then Exit Do
        If Left$(com_token, 1) <> "{" Then Sys_Error "ED_LoadFromFile: found %s when expecting {", com_token

        If ent = -1 Then
            ent = EDICT_NUM(0)
        Else
            ent = ED_Alloc()
        End If
        p = ED_ParseEdict(data, p, ent)

        ' remove things from different skill levels or deathmatch
        flags = ToInt(ev(ent).spawnflags)
        If cvars(cv_deathmatch).value <> 0 Then
            If (flags And SPAWNFLAG_NOT_DEATHMATCH) Then
                ED_Free ent
                inhibit = inhibit + 1
                GoTo nextent
            End If
        ElseIf (current_skill = 0 And (flags And SPAWNFLAG_NOT_EASY) <> 0) _
            Or (current_skill = 1 And (flags And SPAWNFLAG_NOT_MEDIUM) <> 0) _
            Or (current_skill >= 2 And (flags And SPAWNFLAG_NOT_HARD) <> 0) Then
            ED_Free ent
            inhibit = inhibit + 1
            GoTo nextent
        End If

        '
        ' immediately call spawn function
        '
        If ev(ent).classname = 0 Then
            Con_Printf "No classname for:\n"
            ED_Print ent
            ED_Free ent
            GoTo nextent
        End If

        ' look for the spawn function
        func = ED_FindFunction(PR_GetString(ev(ent).classname))

        If func = -1 Then
            Con_Printf "No spawn function for:\n"
            ED_Print ent
            ED_Free ent
            GoTo nextent
        End If

        pr_global_struct(0).self = EDICT_TO_PROG(ent)
        PR_ExecuteProgram func
nextent:
    Loop

    Con_DPrintf "%i entities inhibited\n", inhibit
End Sub

Public Sub PR_LoadProgs()
    Dim i As Long, buf() As Byte, ofs As Long, n As Long
    Dim ofs_statements As Long, ofs_globaldefs As Long, ofs_fielddefs As Long
    Dim ofs_functions As Long, ofs_strings As Long, ofs_globals As Long, version As Long, crc As Long

    ' flush the non-C variable lookup cache
    For i = 0 To GEFV_CACHESIZE - 1
        gefvCache_field(i) = ""
    Next

    CRC_Init pr_crc

    If Not COM_LoadHunkFile("progs.dat", buf) Then Sys_Error "PR_LoadProgs: couldn't load progs.dat"
    Con_DPrintf "Programs occupy %iK.\n", com_filesize \ 1024

    For i = 0 To com_filesize - 1
        CRC_ProcessByte pr_crc, buf(i)
    Next

    version = GetLong(buf, 0)
    crc = GetLong(buf, 4)
    ofs_statements = GetLong(buf, 8)
    progs_numstatements = GetLong(buf, 12)
    ofs_globaldefs = GetLong(buf, 16)
    progs_numglobaldefs = GetLong(buf, 20)
    ofs_fielddefs = GetLong(buf, 24)
    progs_numfielddefs = GetLong(buf, 28)
    ofs_functions = GetLong(buf, 32)
    progs_numfunctions = GetLong(buf, 36)
    ofs_strings = GetLong(buf, 40)
    progs_numstrings = GetLong(buf, 44)
    ofs_globals = GetLong(buf, 48)
    progs_numglobals = GetLong(buf, 52)
    progs_entityfields = GetLong(buf, 56)

    If version <> PROG_VERSION Then Sys_Error "progs.dat has wrong version number (%i should be %i)", version, PROG_VERSION
    If crc <> PROGHEADER_CRC Then Sys_Error "progs.dat system vars have been modified, progdefs.h is out of date"
    If progs_entityfields > EDICT_STRIDE Then
        Sys_Error "progs.dat has %i entity fields, QuakeVB has room for %i", progs_entityfields, EDICT_STRIDE
    End If

    pr_edict_size = EDICT_STRIDE * 4

    ' the lumps
    ReDim pr_statements(0 To progs_numstatements)
    ofs = ofs_statements
    For i = 0 To progs_numstatements - 1
        pr_statements(i).op = GetShort(buf, ofs) And &HFFFF&
        pr_statements(i).a = GetShort(buf, ofs + 2)
        pr_statements(i).b = GetShort(buf, ofs + 4)
        pr_statements(i).c = GetShort(buf, ofs + 6)
        ofs = ofs + 8
    Next

    ReDim pr_functions(0 To progs_numfunctions)
    ofs = ofs_functions
    For i = 0 To progs_numfunctions - 1
        pr_functions(i).first_statement = GetLong(buf, ofs)
        pr_functions(i).parm_start = GetLong(buf, ofs + 4)
        pr_functions(i).locals = GetLong(buf, ofs + 8)
        pr_functions(i).profile = GetLong(buf, ofs + 12)
        pr_functions(i).s_name = GetLong(buf, ofs + 16)
        pr_functions(i).s_file = GetLong(buf, ofs + 20)
        pr_functions(i).numparms = GetLong(buf, ofs + 24)
        For n = 0 To MAX_PARMS - 1
            pr_functions(i).parm_size(n) = buf(ofs + 28 + n)
        Next
        ofs = ofs + 36
    Next

    ReDim pr_globaldefs(0 To progs_numglobaldefs)
    ofs = ofs_globaldefs
    For i = 0 To progs_numglobaldefs - 1
        pr_globaldefs(i).typ = GetShort(buf, ofs) And &HFFFF&
        pr_globaldefs(i).ofs = GetShort(buf, ofs + 2) And &HFFFF&
        pr_globaldefs(i).s_name = GetLong(buf, ofs + 4)
        ofs = ofs + 8
    Next

    ReDim pr_fielddefs(0 To progs_numfielddefs)
    ofs = ofs_fielddefs
    For i = 0 To progs_numfielddefs - 1
        pr_fielddefs(i).typ = GetShort(buf, ofs) And &HFFFF&
        If pr_fielddefs(i).typ And DEF_SAVEGLOBAL Then Sys_Error "PR_LoadProgs: pr_fielddefs[i].type & DEF_SAVEGLOBAL"
        pr_fielddefs(i).ofs = GetShort(buf, ofs + 2) And &HFFFF&
        pr_fielddefs(i).s_name = GetLong(buf, ofs + 4)
        ofs = ofs + 8
    Next

    ' the strings, the engine's string slots, and room for ED_NewString
    ReDim pr_strings(0 To progs_numstrings + PRS_ENGINE_SIZE + 65536)
    CopyMemory pr_strings(0), buf(ofs_strings), progs_numstrings
    PRS_TEMP = progs_numstrings
    PRS_MAPNAME = PRS_TEMP + 128
    PRS_WORLDMODEL = PRS_MAPNAME + 64
    PRS_CLIENTNAMES = PRS_WORLDMODEL + 64
    pr_strings_used = PRS_CLIENTNAMES + MAX_SCOREBOARD * 32
    For i = 0 To svs.maxclients - 1
        PR_SetEngineString PRS_CLIENTNAMES + i * 32, 32, svs.clients(i).name
    Next

    ' the globals (at least as many as globalvars_t has)
    n = progs_numglobals
    If n < GLOBALVARS_SIZE Then n = GLOBALVARS_SIZE
    ClearGlobalViews
    ReDim pr_gl(0 To n - 1)
    CopyMemory pr_gl(0), buf(ofs_globals), progs_numglobals * 4
    SetGlobalViews
End Sub

Public Sub PR_Init()
    type_size(0) = 1: type_size(1) = 1: type_size(2) = 1: type_size(3) = 3
    type_size(4) = 1: type_size(5) = 1: type_size(6) = 1: type_size(7) = 1

    Cmd_AddCommand "edict", XC_ED_PrintEdict_f
    Cmd_AddCommand "edicts", XC_ED_PrintEdicts
    Cmd_AddCommand "edictcount", XC_ED_Count
    Cmd_AddCommand "profile", XC_PR_Profile_f
    cv_nomonsters = Cvar_RegisterVariable("nomonsters", "0")
    cv_gamecfg = Cvar_RegisterVariable("gamecfg", "0")
    cv_scratch1 = Cvar_RegisterVariable("scratch1", "0")
    cv_scratch2 = Cvar_RegisterVariable("scratch2", "0")
    cv_scratch3 = Cvar_RegisterVariable("scratch3", "0")
    cv_scratch4 = Cvar_RegisterVariable("scratch4", "0")
    cv_savedgamecfg = Cvar_RegisterVariable("savedgamecfg", "0", True)
    cv_saved1 = Cvar_RegisterVariable("saved1", "0", True)
    cv_saved2 = Cvar_RegisterVariable("saved2", "0", True)
    cv_saved3 = Cvar_RegisterVariable("saved3", "0", True)
    cv_saved4 = Cvar_RegisterVariable("saved4", "0", True)
End Sub

Public Function EDICT_NUM(ByVal n As Long) As Long
    If n < 0 Or n >= sv.max_edicts Then Sys_Error "EDICT_NUM: bad number %i", n
    EDICT_NUM = n
End Function

Public Function NUM_FOR_EDICT(ByVal e As Long) As Long
    If e < 0 Or e >= sv.num_edicts Then Sys_Error "NUM_FOR_EDICT: bad pointer"
    NUM_FOR_EDICT = e
End Function

' C: (byte *)e - (byte *)sv.edicts
Public Function EDICT_TO_PROG(ByVal e As Long) As Long
    EDICT_TO_PROG = e * pr_edict_size
End Function

' C: (edict_t *)((byte *)sv.edicts + e)
Public Function PROG_TO_EDICT(ByVal e As Long) As Long
    PROG_TO_EDICT = e \ pr_edict_size
End Function

' C: G_EDICT(o), G_EDICTNUM(o)
Public Function G_EDICT(ByVal o As Long) As Long
    G_EDICT = pr_gl(o) \ pr_edict_size
End Function

Public Function G_EDICTNUM(ByVal o As Long) As Long
    G_EDICTNUM = NUM_FOR_EDICT(G_EDICT(o))
End Function

' C: G_STRING(o)
Public Function G_STRING(ByVal o As Long) As String
    G_STRING = BufToStr(pr_strings, pr_gl(o))
End Function
