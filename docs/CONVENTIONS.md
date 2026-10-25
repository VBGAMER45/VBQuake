# QuakeVB porting conventions

How the C code of WinQuake (in `ref/WinQuake`) is translated to VB6. Most of
the pitfalls come from DoomVB (`C:\Users\Owner\Desktop\websites\doom\docs\CONVENTIONS.md`),
where each one was hit at least once; the Quake-specific ones were hit in
this port. Items marked *(decide in cluster N)* are still open: settle them
in that cluster and update this file.

## Structure

* One `.bas` module per C file, same name (`sv_phys.c` -> `src/sv_phys.bas`).
  A module name must not equal a procedure name; add `_c` when it would
  (`sv_move.bas` is module `sv_move_c`, because of world's `SV_Move`).
* Shared types, constants and globals from the headers live in
  `src/quakedef.bas`. A global that is `extern` in a header goes there, even
  while the C file that defines it is not ported; file-local C globals
  become `Private`. The partial `client_state_t`, `server_t` etc. there grow
  as their clusters arrive.
* Clusters not yet ported are stood in for by `stubs/stubs_<cluster>.bas`
  (empty procedures with the real signatures). `tools/build.py` writes
  `QuakeVB.vbp` from `src/` and adds the stub of every cluster none of whose
  files are in `src/` yet.
* A C file that needs a later cluster may be ported in part: its header
  comment lists what is missing. The other way round, a stub module may hold
  one real function that an earlier cluster needs: `stubs_D.bas` had the
  real `R_ParseParticleEffect` (r_part.c) until cluster D, since it reads
  message bytes (see the stub pitfall below).
* Only the software renderer and the Windows platform are ported: `id386` is
  0 (the C fallbacks are used, never the `.s` files); GLQuake, QuakeWorld,
  DOS, Linux, IPX and serial code are not.
* The Windows layers (`sys_win`, `vid_win`, `in_win`, `snd_win`, `cd_win`)
  are rewritten in the style of DoomVB's `i_*.bas`, keeping the C function
  names that the rest of the engine calls.
* Generated code (field offsets from `progdefs.h`, test cases) comes from a
  script in `tools/`; never edit generated files by hand.

## Memory model: pointers become indices

| C | VB |
|---|----|
| `vec3_t` | UDT `vec3_t` with `v(0 To 2) As Single`: `org.v(i)`, so C's `for (i=0 ; i<3 ; i++)` loops stay loops and vectors pass ByRef. The mathlib.h macros are procedures (`DotProduct`, `VectorCopy`, ...) |
| `float m[3][3]`, `[3][4]` | `Single` arrays with the same bounds, passed ByRef |
| `cvar_t foo` + `Cvar_RegisterVariable (&foo)` | `cvars()` array (quakedef.bas); the global is the index `cv_foo = Cvar_RegisterVariable("foo", "0", archive, server)`, read as `cvars(cv_foo).value` / `.strval` (C: `.string`). Index 0 is an unregistered dummy with value 0 |
| command and alias lists, cvar list | arrays; C inserts at the head of a list, so "walk the list" is "walk the array from the end" (completion and `path` print in the C order) |
| `xcommand_t` (`Cmd_AddCommand`) | `XC_<function>` constant (quakedef.bas), dispatched by `Cmd_Call` in cmd.bas. A module with many commands gets its own dispatcher (`Host_CallCommand`, `CL_InputCommand`, `NET_CallCommand`, ...) over a consecutive range of constants, and `Cmd_Call` routes the range (`Case XC_IN_KLookDown To XC_IN_Impulse`), so the command functions stay Private |
| other function pointers (`net_drivers[d].Connect`, `PollProcedure.procedure`) | a number and a `Select Case`: net_main.bas's `Drv_` functions on the driver number (0 loopback, 1 datagram), net_dgrm.bas's `Lan_` functions on the LAN driver number (`net_landrivers[]`: 0 Winsock TCP/IP), `RunPollProcedure` on a `PP_` constant |
| `char *` strings | VB Strings with one character per byte: make characters with `ChrW$`, read them with `AscW` (never `Chr$`/`Asc`/`StrConv` on bytes 128..255: they go through the code page). `BytesToStr`, `BufToStr`, `StrToBytesZ` (common.bas) convert |
| `char` buffers that C edits in place (`key_lines`, `con_text`, files, messages) | Byte arrays |
| `COM_Parse (char *data)` | `p = COM_Parse(bytes, p)` over a zero-terminated Byte array; -1 = NULL |
| `byte *buf = COM_LoadFile (...)` | `If COM_LoadFile(path, buf) Then` (buf gets the file plus a 0 byte); hunk/temp/cache variants are the same |
| C file handles, `FILE *` | Win32 file handles (`Sys_FileOpenRead` etc. in sys_win.bas; -1 = none). `COM_FOpenFile` opens a new handle positioned at the file in its pak |
| `sizebuf_t` | UDT with `data() As Byte`; `SZ_GetSpace` returns an offset |
| `qpic_t *` | index into `pics()` (0 = NULL); the pixels are in `draw_pool()` at `pics(n).ofs` |
| `vid.buffer`, `vid.conbuffer` | `vid_buffer()`; C's `vid.buffer` starts at `vid_bufferofs` (0, except while `lcd_x` draws its second view a row down), `vid.conbuffer` at 0 |
| wad lumps (`W_GetLumpName`) | offset into `wad_base()` |
| `qboolean` | `Boolean` (True is -1: write C's `x = (a == b)` used as an int as `-(a = b)`) |
| `longjmp (host_abortserver)` | `Err.Raise HOST_ABORTSERVER`, caught in `Host_Frame_` (C: `_Host_Frame`) |
| `pr_globals` (float/int/string/function union) | `pr_gl() As Long` owns the memory; `pr_gf() As Single` and `pr_global_struct(0) As globalvars_t` are views over it. `pr_global_struct(0).self` is C's `pr_global_struct->self`, `pr_gf(OFS_PARM0)` is `G_FLOAT(OFS_PARM0)` |
| entity fields `ent->v.x`, `E_FLOAT(e, o)` | `ed_l() As Long` owns all edicts' fields, `EDICT_STRIDE` (512) Longs per edict; `ed_f() As Single` and `ev() As entvars_t` are views over it, so `ev(e).origin` is `ent->v.origin` and passes ByRef straight into progs memory. `edicts()` holds the rest of `edict_t` (free, area links, leafnums, baseline, freetime). `progdefs.bas` (generated by `tools/gen_progdefs.py` from `progdefs.q1`) has the types and the `GOFS_`/`EOFS_` offsets |
| `edict_t *` | edict number, 0-based (0 is the world), NULL = -1 |
| edict pointer as QuakeC sees it (`EDICT_TO_PROG`) | stays a byte offset: `e * pr_edict_size` (2048); a field pointer (`OP_ADDRESS`) is a byte offset into `ed_l` (index `ptr \ 4`) |
| `eval_t *GetEdictFieldValue` | `ed_l` index, -1 = NULL |
| `string_t` | offset into `pr_strings()`: the progs' strings, then fixed engine slots (`PRS_TEMP`, `PRS_MAPNAME`, `PRS_WORLDMODEL`, the client names `PRS_CLIENTNAMES`), then `ED_NewString` strings. `PR_GetString(ofs)` reads one; a client's name is set with `SV_SetClientName`, which updates its slot too |
| `model_t *` and the other handles (`texture_t*`, sprites, `aliashdr_t*`, `qsocket_t*`) | 1-based index into `mod_known()`, `textures()`, `aliashdrs()`, `net_qsockets()`, ...; **0 = NULL**, so a zeroed UDT holds NULLs |
| brush model data (`msurface_t*`, `mplane_t*`, `mtexinfo_t*`, ...) | 0-based index into one pool per type in model.bas (`mod_planes()`, `mod_surfaces()`, ...); `model_t` keeps the index of its first element where C keeps a pointer, and C's int indices into those arrays stay relative to it (a hull's clipnode children and plane numbers, as in C) |
| `mnode_t *` / `mleaf_t *` | one `mnode_t` UDT for both, in `mod_nodes()`; `contents < 0` means a leaf, as C tells them apart. `model_t.nodes` and `.leafs` are the indices of node 0 and leaf 0 |
| hunk (`Hunk_FreeToLowMark (host_hunklevel)`) | the pools: `Mod_HunkMark` at the end of `Host_Init`, `Mod_HunkFreeToMark` in `Host_ClearMemory`. Alias and sprite data stay (C keeps them in the cache) |
| `entity_t *` (client) | index into the one `cl_entities()` array, -1 = NULL. C's four entity arrays are ranges of it: `cl_entities[n]` is `cl_entities(n)`, `cl_static_entities[n]` is `cl_entities(CL_STATIC_BASE + n)`, `cl_temp_entities[n]` is `cl_entities(CL_TEMP_BASE + n)`, `cl.viewent` is `cl_entities(CL_VIEWENT)`. `cl_visedicts()` holds such indices |
| `entity_t.colormap` (`byte *`) | 0 = `vid.colormap` (`vid_colormap()`), n = `cl.scores(n - 1).translations` |
| `efrag_t *` | `cl_efrags()` index, 1-based, 0 = NULL |
| `qsocket_t *`, `cls.netcon`, `client_t.netconnection` | `net_qsockets()` index, 0 = NULL; the active and free lists chain through `.nxt`; a loopback socket's `driverdata` is the socket at the other end |
| `struct qsockaddr` / `sockaddr_in` | `qsockaddr` (16 bytes, passed straight to Winsock): `sa_data(0..1)` the port, `sa_data(2..5)` the address, both in network byte order; net_wins.bas keeps addresses as C keeps `s_addr` (`myAddr`, ban addresses) |
| net_dgrm.c's `packetBuffer` struct | a byte array: the length and sequence words in network byte order (`PutBigLong` / `GetBigLong`), the data from `NET_HEADERSIZE` |
| `sfx_t *` | `known_sfx()` index (snd_dma.bas), 1-based, 0 = NULL (S_PrecacheSound returns 0 while there is no sound) |
| `sfxcache_t *` | `sfxcaches()` index, the same number as its sfx (0 = not loaded); the samples are its own `data8()` (signed char bit patterns) or `data16()`. Loaded once and kept: no cache allocator |
| `channel_t *` | `channels()` index, -1 = NULL |
| `dma_t *shm`, `shm->buffer` | `shm` is C's `sn`, `shm_valid` says whether C's `shm` points at it; `shm.buffer` is the address of snd_win.bas's `dma_buffer()` (Integer), which the mixer writes |
| `portable_samplepair_t paintbuffer[]`, `int *snd_p` | one Long array, pair i at (2i, 2i+1); `snd_p` is an index there |
| `FILE *cls.demofile` | Win32 file handle, 0 = NULL (`COM_FOpenFile` opens a new one, positioned in the pak) |
| `link_t` (area lists) | `link_prev()`/`link_next()` index arrays (common.bas): 0 .. MAX_EDICTS-1 are the edicts' area links, an areanode's trigger and solid heads come after them |
| zone, cache (`zone.c`) | not ported: VB arrays and Strings |
| renderer pointers into the view buffer, z-buffer, textures | offsets: `d_viewbuffer()` (Byte, see the views below), `d_pzbuffer()` (Integer: C's `short`, so a z value is stored as `(izi And &HFFFF0000) \ &H10000`), `texpixels()` (all textures' mips; `textures(n).offsets(mip)`). `d_scantable()` and `zspantable()` hold row offsets |
| `byte *cacheblock`, `byte *d_viewbuffer` | Byte arrays laid over other memory with `SetView` (common.bas): `D_SetCacheBlock` points `cacheblock()` at a surface cache block or a texture, `D_SetupFrame` points `d_viewbuffer()` at `vid_buffer()` or `r_warpbuffer()`. `D_ClearViews` detaches them before `End` |
| `surfcache_t *` (surface cache) | the cache is one Long array `sc_l()` with the views `sc_bytes()` and `sc_f()`; a block keeps C's header layout (pixels at byte 48, sizeof 52) and is its byte offset + 1, **0 = NULL**. Its owner (`surfcache_t **`, an `msurface_t`'s `cachespots[mip]`) is `surface * MIPLEVELS + mip + 1`. d_surf.bas's helpers (`D_CacheWidth`, `D_CacheDataOfs`, `D_CacheDataBytes`) read a block |
| `edge_t *` | `r_edges()` index: C's fixed edges are `EDGE_HEAD` 0, `EDGE_TAIL` 1, `EDGE_AFTERTAIL` 2, `EDGE_SENTINEL` 3, the frame's edges start at `EDGES_BASE` (C's `r_edges[0]`), -1 = NULL. `cachededgeoffset` is `edge - EDGES_BASE`, as in C |
| `surf_t *` | `surfaces()` index, 0 to `r_cnumsurfs` (`surf_max` is `r_cnumsurfs + 1`): 0 is C's unused `surfaces[0]`, 1 the background |
| `espan_t *` | `spans()` index (C: R_ScanEdges' local `basespans`), -1 = NULL |
| `mvertex_t *` | a ByRef `mvertex_t` parameter; a `bedge_t` vertex reference is a `mod_vertexes()` index when >= 0 and `r_bverts(-1 - ref)` (clipped submodel vertices) when < 0 |
| `mnode_t *` as NULL-able (`entity_t.topnode`, `r_pefragtopnode`, `r_viewleaf`) | `mod_nodes()` index, -1 = NULL |
| `efrag_t **lastlink` (r_efrag.c) | 0 = the entity's `efrag` field, k > 0 = `cl_efrags(k).entnext` |
| `aliashdr_t *`, `mdl_t *` (r_alias.c) | `aliashdrs()` index (`paliashdr`, `pmdl`; C's pmdl is the header's `.mdl`). Its data are pool indices: `alias_verts()` (frames, `r_apverts`), `alias_stverts()`, `alias_tris()`, `alias_skins()` (skin pixels: `r_affinetridesc.pskin`, `skintable()`), `alias_skindescs()`, groups and intervals |
| `finalvert_t *`, `auxvert_t *` | `a_finalverts()` index: the model's vertices from 0 (C: R_AliasDrawModel's local array), r_aclip's two clip buffers (C: `fv[2][8]`) from `FV_CLIP`; `a_auxverts()` alongside. `r_affinetridesc.ptriangles` is an `alias_tris()` index, -1 = r_aclip's `aclip_tri` |
| `int *` vertex pointers in d_polyse.c (`r_p0`..`r_p2`, the edge tables, `D_PolysetRecursiveTriangle`) | offsets into one Long array: `r_p()` (RP0, RP1, RP2; -1 = NULL), `rtv()` for the recursive triangles (6 Longs per vertex, the split points stacked after the corners) |
| `acolormap` (`void *`) | Byte array laid over `vid_colormap()` or `cl.scores(n - 1).translations` with `SetView`, per model drawn |
| `particle_t *` | `particles()` index, -1 = NULL (`active_particles`, `free_particles`, `.nxt`) |
| `mspriteframe_t *`, `emitpoint_t *` (sprites) | `sprite_frames()` index; the projected points are `r_spriteverts()`, a `sspan_t *` a `sprite_spans()` index, `cacheblock()` is laid over the frame's `sprite_pixels()` |

**Aliased arrays** (`tools/probe/probe.bas` has the working technique): a
module-level `SAFEARRAY1D` descriptor with `fFeatures = &H11` (auto, fixed
size) and `pvData` pointing at the Long array's first element, copied into
the Single array variable through `VarPtrArray`. Never `ReDim` the Long
array while the alias exists (re-point it after every allocation), clear the
alias before the program ends, and keep `cElements` equal to the Long
array's size. `SetView` and `ClearView` (common.bas) do it. pr_edict.bas
keeps the progs views: `PR_LoadProgs` and `PR_AllocEdicts` re-point them,
and `PR_ClearViews` clears them; the renderer's views (the surface cache,
`cacheblock()`, `d_viewbuffer()`) are cleared by `D_ClearViews`. `Sys_Quit`
and `Sys_Error` call both before `End`.

## printf and strings

* The printf family (`va`, `Con_Printf`, `Con_DPrintf`, `Con_SafePrintf`,
  `Sys_Printf`, `Sys_Error`, `Host_Error`, `Host_EndGame`, `SV_ClientPrintf`,
  ...) take the C format string with its backslash escapes and a
  `ParamArray`; `vsprintf` (common.bas) interprets `\n`, `\t`, `\r`, octal
  `\ooo`, `\\` and C's conversions (`%[-+ 0#][width][.prec][hlL]d i u x X o
  c s f e E g G p %`). Its output matches C's printf
  (`python -I tools/vbtest.py` checks 300 cases).
* **VB string literals have no escapes.** A quote inside a string is `""`,
  never `\"` (that ends the string). So C's `"\"%s\" is \"%s\"\n"` becomes
  `"""%s"" is ""%s""\n"`: the quotes in VB style, the `\n` left for
  vsprintf. A single backslash is `"\"`.
* Only the printf family interprets escapes. `Con_Print`, `Cbuf_AddText`,
  `Cbuf_InsertText`, `SZ_Print`, `MSG_WriteString` and the rest take plain
  text: write a newline as `vbLf` there.
* When C passes a variable as the format (`Con_Printf (text)`), write
  `Con_Printf "%s", text`, so a backslash in the text is not taken for an
  escape.
* A ParamArray goes on to another procedure as a Variant: `Dim a As Variant:
  a = args: s = vsprintf(fmt, a)`.

## VB6 pitfalls (each one bit DoomVB or this port)

* **Integer literals are 16-bit.** `10240 * 4` overflows (silently, with
  overflow checks off), and so does `Const W = 320, H = 200` used as `W * H`
  (the feasibility probe hit this). Write `10240& * 4`, `Const W As Long`.
  Hex `&H8000`..`&HFFFF` are negative Integers: use `&H8000&`, `&HFFFF&`.
* **Operator precedence:** comparisons bind tighter than `And`/`Or`:
  `If flags And FL_ONGROUND <> 0` is wrong; write `If (flags And FL_ONGROUND) <> 0`.
  `\` binds tighter than `And`: write `(x And &HFFFF0000) \ &H10000`.
* **`\` binds looser than `*` and `/`:** VB reads `a * b \ c * d` as
  `(a * b) \ (c * d)`, where C's `a * b / c * d` goes left to right. C's
  `w*h/64*85` (model.bas's texture size) must be `(w * h \ 64) * 85`; as
  written first it made the texture pool tiny and showed up only as error 9
  in the renderer. `python -I tools\check_intdiv.py` finds every `\` that a
  later `*` or `/` takes apart.
* **No short-circuit:** `If a >= 0 And arr(a) = 1` evaluates `arr(-1)`. Nest Ifs.
* **`/` is floating point.** C integer `/` is `\` (truncates toward zero, like C).
* **Shifts:** `x >> 16` on a possibly negative value is `(x And &HFFFF0000) \ &H10000`
  (`x \ 65536` truncates instead of flooring). `SAR(x, n)` (mathlib.bas) is
  the general form; hot loops inline the mask and the division.
* **Unsigned:** VB has no unsigned 32-bit type. Compare with `ULT/UGT/ULE/UGE`
  helpers where C compares unsigned values.
* **`Not` is bitwise:** for a Long used as a boolean write `x = 0`, not `Not x`.
* **Module-level declarations must come before the first procedure.**
* **Reserved-ish names** to avoid as identifiers: `scale`, `line`, `name`,
  `type`, `next`, `step`, `len`, `left`, `right`, `time`, `date`, `point`,
  `dir`, `str`, `print`, `open`, `close`, `get`, `put`, `me`, `Empty`,
  `string`, `width`/`height` as globals. C field `type` is `typ`, `next` is
  `nxt`, `string` is `strval`; other names get a trailing underscore
  (`right_`, `time_`, `str_`).
* **Case-insensitive names:** `Cmd_Argc` and `cmd_argc`, `Key_Console` and
  `key_console`, `SCR_DrawLoading` and `scr_drawloading` are one name to
  VB, also inside one module. Rule: the C name used in more places keeps
  it, the other gets a trailing underscore (`cmd_argc_`, `Key_Console_`,
  `SCR_DrawLoading_`, `MINIMUM_MEMORY_`, `con_debuglog_`). Cvars avoid the
  problem with their `cv_` prefix. Run **`python -I tools\check_names.py`**
  before building: it finds them all in one pass (VB reports one per build).
  `python -I tools\case_clashes.py ref\WinQuake` lists the C side.
* **Case-insensitive locals:** a parameter or local hides a global of the
  same name in any case, and the compiler says nothing. sv_move.c's
  `SV_StepDirection (edict_t *ent, float yaw, float dist)` uses both `yaw`
  and the macro `YAW`; in VB `angles.v(YAW)` indexed with the parameter (by
  then in radians) and failed with error 9. check_names.py reports every
  local spelled like a global in another case; rename the local (`yaw_`).
* **Editing the sources by script:** the files are CRLF. Python's text mode
  turns them into `\n`, so never split on `\r\n` after a text-mode read, and
  compare the line count before writing a file back. A Python heredoc in
  bash also turns `\n` and `\35` inside the script's strings into real
  characters: use the Edit tool for VB lines that contain backslashes.
  Git Bash's `sed -i` can write the whole file back with LF line ends (it
  did on a test copy), and `\r` in its pattern does not match: for scripted
  edits read and write the file in binary mode in Python, and check with
  `grep -c $'\r'` that every line still ends in CRLF.
* **API error codes:** read `Err.LastDllError`, not `GetLastError` /
  `WSAGetLastError`.
* Pass scalars `ByVal` (VB defaults to ByRef). UDTs and arrays must be ByRef.
  `As Any` exists only in Declare statements: a procedure that reads into
  memory takes `buf() As Byte, ByVal ofs As Long`.
* `IIf` evaluates both branches: use If blocks.
* Never `ReDim` an array inside a `With` block on one of its elements, or
  while one of its elements is passed ByRef (the array is locked).
* **Never `GoTo` out of a `With` block on an array element.** `End With`
  unlocks the array; a GoTo to a label outside the block skips that, and
  the lock stays after the procedure returns (a probe confirmed it;
  `Exit Sub` / `Exit Function` inside the block are fine). The next `ReDim`
  of the array then fails with error 10 ("This array is fixed or
  temporarily locked"), possibly much later: a `continue` written as
  `GoTo nexti` in RecursiveLightPoint broke the next map after a demo. Put
  the label just before `End With` (with a flag if code follows the
  block). `python -I tools\check_with.py` finds them.
* `Not Not arr` (the "is it allocated" trick) upsets VB's floating point
  stack: keep a flag instead.
* **64 KB of fixed-size data per module.** A module's fixed arrays and
  UDTs together must stay under 64 KB ("Fixed or static data can't be
  larger than 64K"). Big arrays (`spans()`, `btofpolys()`, the sky
  textures) are dynamic and `ReDim`ed once, behind a flag.
* Executing an `On Error` statement clears `Err`: read `Err.Number` and
  `Err.Description` before calling anything (`Sys_RuntimeError` takes them
  as parameters).

## Quake-specific pitfalls

* **C `(int)` truncates, VB `CLng` rounds** (to even: `CLng(2.5) = 2`,
  `CLng(-1.5) = -2`; C gives 2 and -1). Every `(int)` cast of a float is
  `ToInt(x)` (common.bas: truncates, and gives C's 0x80000000 for values out
  of range and NaN); C `floor` is `Int(x)`. Entity fields are floats, so
  `(int)ent->v.flags & FL_ONGROUND` is `ToInt(ev(e).flags) And FL_ONGROUND`,
  and C's integer `abs` of a float is `Abs(ToInt(x))`. `MSG_WriteByte`,
  `MSG_WriteChar`, `MSG_WriteShort` and `MSG_WriteLong` take a Double and
  truncate inside, as C's implicit conversion does.
* **Float to Long out of range raises error 6** even with overflow checks off
  (C gives 0x80000000). `ToInt` handles it; clamp elsewhere.
* **VB's unary minus of 0 gives +0**, where C's gives -0, and C's printf
  prints -0 as `-0.000000`. `C_atof` negates with `* -1#`, and vsprintf
  takes the sign of `%f`/`%e`/`%g` from the sign bit, so savegames and
  `edict` dumps print as in C.
* **A NaN poisons the next conversion.** With floating point checks off,
  `0/0` gives NaN silently, but it leaves the FPU's sticky "invalid" flag
  set, and VB's next float-to-Long conversion then fails with error 6, on a
  perfectly good value, somewhere else (measured: `x = Atn(0/0)` then
  `l = Fix(240!)` raises error 6; `_clearfp` from msvcrt.dll clears the
  flag). C ignores such NaNs. So any C computation that can make a NaN
  (`0/0`, `inf - inf`, `0 * inf`, comparisons with a NaN) needs a guard in
  VB, noted in a comment: `CalcFov` with an empty vrect was the first. A
  run time error 6 on a harmless line means: look for an earlier NaN.
  QuakeC can make NaNs and divisions by zero at will, so `Sys_ClearFPU`
  (sys_win.bas, `_clearfp`) clears the flags before every builtin, when
  `PR_ExecuteProgram` returns, in `ToInt` after a NaN, and at the start of
  every frame; the VM's own float to int conversions (`OP_BITAND`,
  `OP_BITOR`) clear them first as well.
* **Run-time errors open a modal dialog.** Hidden test runs must never reach
  one: `Sys_Error` skips the MessageBox under `-hidden`, and `Host_Frame_`
  and `Main` trap VB errors and report them through `Sys_Error`.
* **Host_Error unwinds with a VB error**, so no procedure between it and
  `Host_Frame_` may trap errors: no `On Error Resume Next` in engine code,
  only in leaf system functions that call nothing that can reach
  `Host_Error`.
* **No engine code inside Windows callbacks.** The window procedure queues
  key and mouse events; `Sys_SendKeyEvents` hands them to `Key_Event` after
  `DoEvents` (a VB error cannot unwind through a Windows callback).
* Quake's `float` is VB `Single`; write float literals with `!` where the
  precision matters (`65536!`). `double` in C stays `Double` (host timing).
* `char` in the network protocol is signed: `MSG_ReadChar` converts a Byte
  to -128..127. C code that compares `char` text with `<= ' '` also counts
  bytes 128..255 (see `Con_Print`).
* **memset of a struct** is the assignment of a blank UDT
  (`Dim blank As server_t: sv = blank`). Then set the fields whose NULL is
  -1 (`trace.ent`, a node's `parent`) explicitly. `memset (&ent->v, 0, ...)`
  is `FillMemory ed_l(e * EDICT_STRIDE), progs_entityfields * 4, 0`.
* C's NULL-terminated lists (`sv.model_precache`, `sv.sound_precache`)
  become arrays with a count (`num_model_precache`).
* **A stub must keep the message stream in step.** A stand-in for a
  function that reads from `net_message` has to read the same bytes, or
  everything after it in the message is parsed wrong. `R_ParseParticleEffect`
  (svc_particle) was such a function, so stubs_D.bas had the real one; the
  other particle and sound calls get their arguments already parsed.
* `1 << n` for n up to 31 is `BIT_(n)` (mathlib.bas): `2 ^ 31` overflows a
  Long. C's `1 << n` for n >= 32 shifts by `n & 31` on x86.
* Keep C's loop quirks when they change behavior: `CL_UpdateTEnts` reuses
  its beam counter `i` in an inner loop, so after a beam that made entities
  the last beams are skipped that frame; the VB keeps `i` and the beam index
  as separate counters to do the same.
* Checking a doubtful server result: a trace or a QuakeC warning ("walkmonster
  in wall", "fell out of level") is not a port bug just because it looks
  odd. Replay it in a reference: a straight Python port of world.c's
  `SV_RecursiveHullCheck` over the map's clipnodes (floats rounded to 32 bit)
  showed every such warning of the LibreQuake maps to be C behavior (hull2
  monsters under low ceilings, ammo boxes in corners, an explobox2 touching
  another one exactly on its mins face, which a box hull counts as inside).

## Compiler settings

As in DoomVB: native code, **No Optimization** (VB6's "Optimize for Fast Code"
miscompiled a DoomVB routine). Integer overflow and floating point checks are
removed (the code relies on wrap-around). Array bounds checks stay on unless
`--release` is given. The feasibility probe measured the unoptimized,
bounds-checked build as fast enough (see RESUME.md).

## Tools

| Command | Does |
|---|---|
| `python -I tools\build.py [--release]` | writes `QuakeVB.vbp`, compiles `QuakeVB.exe` (VB6 reports the first error with file and line) |
| `python -I tools\check_names.py` | finds VB name clashes (case-insensitive) in one pass, including locals that hide a global |
| `python -I tools\check_intdiv.py` | finds integer divisions `\` that a later `*` or `/` takes apart (VB precedence) |
| `python -I tools\check_with.py` | finds `GoTo` statements that leave a `With` block on an array element (leaked array locks) |
| `python -I tools\gen_progdefs.py` | regenerates `src/progdefs.bas` (globalvars_t, entvars_t, offsets) from `ref/WinQuake/progdefs.q1` |
| `python -I tools\gen_anorms.py` | regenerates `src/anorms.bas` (r_avertexnormals) from `ref/WinQuake/anorms.h` |
| `python -I tools\debugbuild.py` | `QuakeVB_debug.exe`: every procedure logs `module.procedure line N` to quakevb.log on a run time error |
| `python -I tools\vbtest.py` | builds and runs the unit tests in `tests/` (vsprintf against C printf, parsing, CRC) |
| `python -I tools\gen_printf_tests.py` | regenerates `tests/printf_cases.bas` |
| `python -I tools\typekeys.py [--file f] <frame> <step> <text or <KEY>> ...` | scripted typing for `-key` / `-keyfile` |
| `python -I tools\pcx2png.py in.pcx out.png [scale]` | views a `screenshot` |
| `python -I tools\inspect_paks.py id1` | checks game data against the vanilla limits |
| `python -I tools\case_clashes.py ref\WinQuake` | C names that clash in VB |
