# QuakeVB: resume notes

Handoff for continuing the port in a fresh session. Read this file first, then
`docs/CONVENTIONS.md` (VB6 pitfalls and the C-to-VB translation rules).

## Goal

A Visual Basic 6 port of id Software's GPL Quake source (`ref/WinQuake`, the
software-rendered WinQuake; GPL v2, the port stays GPL), done the way DoomVB
(`C:\Users\Owner\Desktop\websites\doom`) was: one VB module per C file,
translated statement by statement, built from the command line, tested with
hidden scripted runs. The game data is LibreQuake (BSD art, GPL `progs.dat`),
so the result is a completely free game.

Sources:

* Engine: https://github.com/id-software/quake at
  `bf4ac424ce754894ac8f1dae6a3981954bc9852d` (2012-01-31). Only `WinQuake/*.c`
  and `*.h` are copied to `ref/WinQuake` (plus `readme.txt`, and
  `progdefs.q1` / `progdefs.q2`, which `progdefs.h` includes; `gnu.txt` is
  `LICENSE.TXT`). The QuakeWorld, GLQuake, DOS, Linux and asm paths are not
  ported.
* Data: https://github.com/lavenderdotpet/LibreQuake release `v0.09-beta`,
  **lite** package: `id1/pak0.pak` (43.9 MB) and `id1/pak1.pak` (5.9 MB),
  license files in `librequake/`. The git repo itself is 1.5 GB of raw
  sources (maps, textures, QuakeC) that need its own build tools; the
  release paks are what an engine loads.

## Status

**Cluster 0 (foundation) is done**: about 6,500 lines of VB in `src/`
(common, cmd, cvar, crc, mathlib, wad, draw, console, keys, screen, host
and host_cmd in part, sys_win, vid_win, frmQuake). QuakeVB.exe opens its
window, loads LibreQuake's paks, checks `gfx/pop.lmp` ("Playing registered
version"), runs `quake.rc`, `default.cfg` and `config.cfg`, and shows the
console over `conback` with `(QuakeVB) 1.09` in the corner, at 320x240
(default, window x2) or any `-width`/`-height` up to 1280x1024.

Verified with hidden scripted runs (see Build and test) and screenshots:
typing at the console, TAB completion (`ver` -> `version `), history (up
arrow), `echo`, `alias` and running an alias, cvar query and set, `path`,
`bind` / `unbind` / `unbindall` from default.cfg, `exec`, `stuffcmds`
(`+cmd` on the command line), `wait`, `screenshot` (PCX), `toggleconsole`,
`quit` (writes config.cfg with the 47 bindings and the archived cvars),
`-condebug` (id1/qconsole.log), word wrap at 38 and 78 columns, the scaled
console background at 640x480, and the visible window (palette, x2 scale,
blinking cursor) captured from the screen. `python -I tools\vbtest.py`:
330 unit tests pass (vsprintf against C printf output, Q_atoi / Q_atof,
COM_Parse, COM_FileBase and friends, the CRC).

"Unknown command" lines at startup are expected: default.cfg sets settings
of modern engines (FTE, Ironwail, QuakeSpasm: `gl_texturemode`,
`r_wateralpha`, `contrast`, `topcolor`, `cl_alwaysrun`, `set`, `if`, ...).

**Cluster A (server and QuakeC) is done**: about 8,700 more lines (model,
pr_edict, pr_exec, pr_cmds, sv_main, sv_phys, sv_move, sv_user, world, the
generated progdefs, the server half of host.c and all of host_cmd.c), about
16,300 lines of VB in all. `map <name>` loads the BSP and its models, runs
LibreQuake's `progs.dat` (CRC 5927, all 78 builtins), spawns the level and
runs server frames with physics, monster thinking, doors and trains. No
client connects yet (`connect local` is a stub until cluster B), so
monsters stand and nothing is drawn.

Verified with hidden scripted runs:

* All ten maps (`start`, `lq_e0m1`..`lq_e0m8`, `dev`) spawn and run 1,100
  server frames without errors under `developer 1` (up to 309 edicts and 81
  MOVETYPE_STEP monsters on lq_e0m2). On lq_e0m2, 1,100 frames take 16.3 s
  after startup (about 67 fps: the frame limiter's pace), and the same at
  `host_framerate 0.1` (7 times the simulated time per frame), so the server
  is far from being the bottleneck.
* The QuakeC warnings these maps print ("Bonus item fell out of level",
  "walkmonster in wall", a `misc_explobox2` that does not drop) were replayed
  in a Python port of world.c's hull trace and are all C behavior
  (CONVENTIONS.md, last Quake pitfall).
* `save` / `load` round trip on `start`: the second save equals the first
  except `frametime`, and the `absmin`/`absmax` of entities that were never
  linked (loading links every edict, as in C). The one free edict more after
  the load is the player slot with all-zero fields (`ED_ParseEdict` frees an
  empty block, as in C).
* `status`, `edict`, `edicts`, `edictcount`, `profile`, `mcache`,
  `changelevel`, `restart`, `map` of a missing map ("Couldn't spawn server"),
  and the client commands (`god` etc.: "Can't "god", not connected").

**Cluster B (client and loopback) is done**: about 4,700 more lines
(net_main, net_loop, cl_main, cl_parse, cl_input, cl_demo, cl_tent, chase,
view, in_win), about 21,400 lines of VB in all. Single player now runs
end to end except for the pictures: `map` starts the server, the local
client connects over the loopback driver, signs on (serverinfo, prespawn,
spawn, begin) and plays; demos play back and record. Nothing is drawn yet
(R_RenderView is a stub), so a visible window shows the tiled backdrop.

Verified with hidden scripted runs:

* Signon on `start`: "LQ Player entered the game", `status` shows the
  client, `quit` sends clc_disconnect. Holding `w` for 100 frames walks the
  player 317 units (input, clc_move, SV_ClientThink, physics).
* `god`, `notarget`, `noclip`, `fly`, `give`, `name`, `color`, `say`,
  `ping`, `pause`, `kill` (respawn restarts the level), `entities`,
  `slist` (finds the local server), `cl_shownet 2`, `changelevel` with the
  client following, `disconnect`, `map` again.
* On lq_e0m1 the player walks into the level and gets killed by the
  monsters ("was gored by a Gore Hound"): AI, damage, death, no errors.
* `save` / `load` with the client connected: the player's origin and angles
  come back exactly; the client reconnects and the server unpauses.
* `timedemo` of all three LibreQuake demos parses every message: demo1
  4527 frames, demo2 4570, demo3 3237 (14,700 to 23,000 fps without
  drawing). The startup demo loop plays demo1 in real time (entities
  interpolated between messages).
* `record vbtest start`, walking, `stop`, then `timedemo vbtest` plays the
  recorded demo back (565 frames).

**Cluster C (the world renderer) is done**: about 5,000 more lines (r_main,
r_bsp, r_edge, r_surf, r_draw, r_light, r_efrag, r_misc, r_sky, r_vars,
d_init, d_modech, d_edge, d_scan, d_surf, d_sky, d_vars, d_fill, d_zpoint,
nonintel), about 26,700 lines of VB in all. The world draws: BSP walk and
PVS, the edge-sorted span rasterizer with the z-buffer, the surface cache
with light styles and dynamic lights, brush models (doors, plats), turbulent
water, slime and lava, the scrolling two-layer sky, the underwater screen
warp, `lcd_x`. Models, sprites and particles are still stubs (cluster D),
and the status bar too (cluster E), so the bottom 48 rows show the tiled
backdrop.

Verified with hidden scripted runs and screenshots:

* The first frame of `start` (the welcome sign), demo1, demo2 and demo3 at
  intervals (lightmaps, brush models, sky, water, slime, a waterfall),
  `r_drawflat 1`, `r_drawentities 0`.
* All ten maps load and render from their spawn point without errors.
* 640x480 renders correctly (`-width 640 -height 480`).
* Under water in lq_e0m1 (wade in, `noclip`, `+movedown`): the warp bends
  the lake bed and the surface seen from below.
* `timedemo` at 320x240 without models, bounds checks on / `--release`:
  demo1 180 / 285 fps, demo2 216 / 297, demo3 267 / 381; demo1 at 640x480
  (release) 132 fps.
* After a timedemo of demo3 the last 3D frame stays on screen with a
  textured band across rows 83-90 and no console. That is C behavior:
  LibreQuake's demos are cutscenes (`cl.intermission` 3, never reset on
  disconnect), so SCR_UpdateScreen draws only the center string branch, and
  SCR_EraseCenterString tile-clears its rows every frame; V_RenderView
  returns early while the console is forced up.
* `noclip` while looking down climbs instead of sinking, and a noclip
  player next to a bank in the water rises for 2 seconds: both C behavior
  (SV_AirMove takes the pitch from `v.angles`, which is `-v_angle / 3`;
  QuakeC's CheckWaterJump sets FL_WATERJUMP, and noclip never updates
  `waterlevel`).

**Cluster D (models, sprites, particles) is done**: about 3,100 more lines
(r_alias, r_aclip, d_polyse, r_sprite, d_sprite, r_part, d_part, and the
generated anorms), about 29,800 lines of VB in all. Monsters, items, the
view weapon, sprites (explosions) and particles (rocket trails, blood,
gunshot puffs, explosions, teleport splashes) are drawn: alias models with
lighting, skin and frame groups, the near-plane and screen-edge clipper, the
affine triangle rasterizer and the recursive subdivision for distant
models, and the player color translation.

Verified with hidden scripted runs and screenshots:

* demo1, demo2 and demo3 to the end without errors: soldiers, Gore Hounds,
  zombies, a scrag, flame models, ammo boxes, models clipped at the screen
  edge.
* On `start`: the view weapon; `give 7`, `give r 100`, `impulse 7` and
  firing: muzzle flash, rocket trail, the explosion sprite, the explosion
  particles fading to smoke, the dynamic light on the wall.
* On lq_e0m1: shotgun pellet puffs on a wall, then killed by a Gore Hound.
* Player translations: a test build that gave every alias model colormap
  1 turned the 16-31 and 96-111 colors of the models gray (the demo
  player's colors), and only those pixels.
* `timedemo` at 320x240 with models, bounds checks on / `--release`: demo1
  134 / 190 fps, demo2 184 / 285, demo3 227 / 335; demo1 at 640x480
  (release) 91 fps.
* `impulse 9` answers "Invalid weapon" in LibreQuake's QuakeC; the engine's
  `give` works. `chase_active 1` shows the floating view weapon, not the
  player: C's R_DrawEntitiesOnList skips the view entity.
* Not tested: `pointfile` (no `.pts` files in LibreQuake).

**Cluster E (status bar and menus) is done**: about 3,000 more lines
(sbar, menu; screen and view were complete already), about 32,900 lines of
VB in all. Single player is playable from the main menu: LibreQuake's menu
art, new game, save and load, options, customize controls, help, quit,
the status bar with inventory, face and the solo scoreboard, the
intermission and finale screens.

Verified with hidden scripted runs (keys only) and screenshots:

* From startup: Escape opens the main menu over the attract demo; SOLO,
  NEW GAME starts `start` with the full status bar; OPTIONS (sliders,
  on/off), MANUAL (LibreQuake's help pages), QUIT and `y` quit.
* TAB shows the solo scoreboard (monsters, secrets, time, level name).
* SAVE to slot 0, then LOAD lists it ("Tainted Tech Threshold" and the
  kills) and loads it.
* ONLINE says "No Communications Available" (no TCP/IP until cluster G);
  SETUP shows the hostname, the name and the translated player picture;
  CUSTOMIZE lists the bindings.
* `viewsize 110` drops the inventory, `viewsize 120` the status bar; at
  640x480 the bar is centered with the backdrop tiled beside it.
* A test build that turned `notarget` into an intermission and `fly` into
  the finale drew COMPLETED with time, secrets and kills, and YOU WIN! with
  the center text typed out slowly.
* timedemo with the status bar: demo1 142, demo2 193, demo3 231 fps.
* Found on the way: `map` after a demo failed with run time error 10. A
  `GoTo` out of a `With mod_surfaces(surf)` block in RecursiveLightPoint
  (lighting of the models, first used by cluster D) left the array locked,
  so the next map's `ReDim Preserve` failed; CL_DecayLights and
  CL_RelinkEntities had the same pattern. All fixed;
  `tools\check_with.py` finds the pattern (CONVENTIONS.md).

**Cluster F (sound) is done**: about 2,100 more lines (snd_dma, snd_mem,
snd_mix, snd_win, cd_win), about 35,100 lines of VB in all. Sound effects
play through waveOut: WAV loading at any rate and width (resampled to 11025
Hz), the 8 dynamic channels, ambient and static sounds with
spatialization, the mixer, `play`, `playvol`, `stopsound`, `soundlist`,
`soundinfo`, and the menu sounds. CD audio is ported (MCI) but has nothing
to play here.

Verified with hidden scripted runs that capture the mix to a WAV file
(`-sndcapture`, then a Python analysis of the file) and one visible run:

* On `start`: the shotgun about every 0.5 s while `+attack` is held, then
  the menu sounds; `soundinfo` (11025 Hz, 16 bit, stereo, 32768 samples),
  `soundlist` (16 bit, looped ones marked L).
* On lq_e0m1: shooting, the Gore Hound and the player's death, with 16-bit
  sounds and with `loadas8bit 1` (the 8-bit mixer): the same loudness.
* demo2 and demo3 to the end with sound, no load errors. demo1 is silent,
  correctly: its three camera positions are far from its one static sound
  (a probe printed every channel's distance and volume).
* A visible run with `volume 0`: the real device opens ("Wave sound
  initialized"), blocks go out, one "Sound overrun" at startup as in C.
  `CDAudio_Init: MCI_OPEN failed (266)`: no CD drive.
* timedemo with the mixer running (capture): demo1 144, demo2 198, demo3
  221 fps, the same as without sound.
* LibreQuake lacks `sound/player/slimbrn2.wav` ("Couldn't load"), as C
  would say.

**Cluster G (network) is done**: about 1,850 more lines (net_dgrm,
net_wins), about 37,000 lines of VB in all. The datagram driver is network
driver 1 over Winsock UDP (LAN driver 0; there is no IPX): reliable
messages split into 1,024-byte packets with ACKs and resends, unreliable
datagrams with sequence numbers, the connect handshake, the server search
by broadcast, `net_stats`, `ban`, `test` and `test2`. The multiplayer
menus now find TCP/IP.

Verified with two hidden copies on this PC (a listen server and a client,
see Build and test):

* `connect 127.0.0.1`: "trying...", "Connection accepted", the client
  signs on and plays on lq_e0m1 (moving, firing, chat in both directions,
  `kill` with QuakeC's -2 frags, respawn). `status` on the server lists
  both players with the client at 127.0.0.1:port; on the client, `ping`
  (17 ms), `status` (forwarded), `test` (player info) and `test2` (the
  server's rules) answer. `net_stats`: 1,814 packets, no resends,
  duplicates or drops. The client quits: "Client vbclient removed".
* With `-netdrop 7` on the server and `-netdrop 5` on the client: resends
  (3 and 2), a duplicate thrown away, 151 and 218 unreliable datagrams
  counted as dropped; the game goes on. The server quits first: the
  client drops to the console (C's Host_EndGame, silent).
* Through the menus only: Multiplayer, New Game, TCP/IP (the net menu with
  IPX dimmed), OK, game options (4 players, deathmatch), begin game on
  `start`; the client's Join Game, TCP/IP, Search for local games ("No
  Quake servers found", see below), then `127.0.0.1` in "Join game at:".
* `connect` with no server: three tries, "No Response". `ban 127.0.0.1`
  on the server: the client gets "You have been banned.", `ban off`
  clears it.
* timedemo with the control socket open and with `-nolan`: the same speed
  (demo1 137 and 134 fps; demo2 167 to 186 either way, this PC was busy).
* Found on the way: a server quit with free player slots waited the whole
  5 s of NET_SendToAll ("failed for 3 clients"): C leaves the free slots'
  flags unset; they count as done now. LibreQuake's net menu pictures are
  264 pixels wide at x 72, so Draw_TransPic stopped with "bad
  coordinates" at 320 pixels (C would too); it cuts them now.
* As in WinQuake, the server search ignores answers from our own address,
  so a server on the same PC is joined by address, not found by Search.
  A second client from the same address replaces the first (C's
  "coming back in from a crash" rule).

## Assessment: how hard compared to DoomVB

Size is similar; the difficulty is not.

| | linuxdoom 1.10 | WinQuake (software path) |
|---|---|---|
| C + headers to port | 54,640 lines | about 63,000 lines (73 C files + headers) |
| code lines (no blanks or comment lines) | about 31,300 | about 39,300 (5,100 of them Windows drivers that get replaced by new, smaller layers) |
| game logic | in C (p_*.c, about 12k lines) | in QuakeC: `progs.dat` runs on a VM, nothing to port |

Harder than DOOM:

1. **Memory model.** Quake leans on pointer arithmetic and type punning far
   more: QuakeC globals and entity fields are one float/int/string union
   addressed by offset, the hunk/zone/cache allocators hand out raw memory,
   `mnode_t` and `mleaf_t` share a header and are told apart by a field, and
   the renderer walks `byte*` / `short*` buffers (view, z-buffer, surface
   cache). All of this becomes arrays plus indices, as in DoomVB, but there
   is more of it.
2. **Floating point.** Quake is float throughout, and C `(int)` truncates
   where VB `CLng` rounds to even. Every cast site needs `Fix`/`Int`, and
   conversions out of Long range raise error 6 in VB (C wraps).
3. **Client/server even in single player.** The local game runs a server and
   a client talking over a loopback message protocol. Nothing is playable
   until server, client and renderer all work (DOOM was playable after the
   renderer and play core).
4. **Renderer.** An edge-sorted span rasterizer with a surface cache,
   perspective-correct texture mapping, an affine triangle rasterizer with a
   z-buffer for models, turbulent water and sky. Many tightly coupled globals.
5. **More system code.** Console, cvars, command buffer, key bindings,
   aliases, a 3,200-line menu with network setup, two network drivers.

Easier than DOOM:

1. **No game logic to port** (monsters, weapons, items, triggers are QuakeC).
   LibreQuake's `progs.dat` is plain vanilla: version 6, CRC 5927 (what
   WinQuake checks), only opcodes 0..65, only builtins 1..78.
2. **Demos are recorded server messages**, not input. Playback does not need
   bit-exact physics, so float rounding differences cannot desync it (DoomVB
   had to match vanilla tic for tic).
3. **Speed headroom is large** (measured below).
4. DoomVB's tooling and lessons carry over: build script, GDI window and
   blit, waveOut mixer, Winsock UDP, scripted hidden test runs, and every
   VB6 pitfall already found.

Estimate: about 35,000 to 45,000 lines of VB, roughly **2.5 to 3 times the
work of DoomVB**. Biggest risks: the renderer's pointer-heavy span and edge
code, and the memory model (decided up front, see CONVENTIONS.md).

## Feasibility measurements (tools/probe)

`python -I tools\probe\build_probe.py tools\probe` compiles `probe.bas` three
ways and runs it. Results on the development PC:

| Test | Bounds checks on, no optimization | Checks off, no optimization | Checks off, fast code |
|---|---|---|---|
| Single array aliased over a Long array (QuakeC union) | works | works | works |
| QuakeC-style dispatch, `Select Case` over all 66 opcodes | 49 M statements/s | 138 M/s | 260 M/s |
| D_DrawSpans8 + D_DrawZSpans, 320x200 full screen | 943 fps | 1,484 fps | 2,458 fps |
| `CLng` / `Int` / `Fix` (Single to Long) | 3.7 / 4.8 / 4.9 ns | | |

A game frame runs some tens of thousands of QuakeC statements, so the VM is
not a concern. The span fill is the core of the world renderer; the full
frame (edges, surface cache, models, particles) costs several times that,
which still leaves a few hundred fps at 320x200 and playable rates at 640x480.
DoomVB's rule (no optimization, because "fast code" miscompiled a routine)
can stay.

## Game data facts (tools/inspect_paks.py)

`python -I tools\inspect_paks.py id1` checks a game folder against the
vanilla engine limits.

* lite: 23 BSPs (`start`, `lq_e0m1`..`lq_e0m8`, `dev`, item boxes), **all
  within vanilla limits** (BSP version 29, under 600 entities, 8192 leafs,
  32767 clipnodes, 256 models, +-4096 map extent).
* `demo1.dem`, `demo2.dem` and `demo3.dem` (4533, 4576 and 3243 messages)
  are all protocol 15 (vanilla) and play back completely: the automated
  playback tests. (An earlier note here said demo3 was protocol 999; its
  serverinfo says 15, and it plays.)
* Sounds: 8-bit and 16-bit mono WAV at 11025 to 64000 Hz (vanilla resamples
  any of these). Eleven item sounds have a `JUNK` chunk before `fmt `, which
  vanilla's chunk search handles.
* `.lit` files (coloured light for GL engines) are ignored, as vanilla does.
* Music: OGG Vorbis tracks in `id1/music/` (not in the paks, not in git).
  WinQuake plays CD audio only, so there is no music until a Vorbis decoder
  is ported (stb_vorbis, about 5,000 lines of C) or the tracks are converted.
* The **full** package (pak0 94.7 MB + pak1 100.6 MB, episodes 1 to 4) has
  13 maps beyond vanilla: up to 1,849 entities, 16,753 leafs, 62,541
  clipnodes (needs unsigned clipnode children), 400 models (needs a protocol
  with 16-bit model indices), extents to 6,271 (needs float coordinates), and
  `lq_e2m4` is BSP2. Its 13 deathmatch maps (`lqdm1`..`lqdm13`, loose files)
  fit vanilla limits (check lqdm5's 2,691 map entities at run time: lights
  without a targetname remove themselves). The full paks stay out of git
  (pak1 is near GitHub's 100 MB file limit).

## Plan: clusters (one at a time, each compiled and tested before the next)

| Cluster | C files | Milestone |
|---|---|---|
| 0 Foundation | quakedef/common (COM_, file system with paks, SZ_/MSG_), zone (replaced), cmd, cvar, crc, mathlib, wad, sys_win, new GDI vid layer, in_win (keyboard), draw, console, keys, host (frame loop), screen (console parts) | window shows the console over `conback`; `quake.rc` and `default.cfg` exec; `path`, `cvarlist`, `cmdlist`, `bind`, `exec`, `quit` work |
| A Server + QuakeC | model, pr_edict, pr_exec, pr_cmds, sv_main, sv_phys, sv_move, sv_user, world, host_cmd | `map start` spawns the level and runs server frames without errors; `edicts` dump is sane; save/load round trip |
| B Client + loopback | net_main, net_loop, cl_main, cl_parse, cl_input, cl_demo, cl_tent, chase, view (non-drawing parts) | local client signs on over loopback; `timedemo demo1` parses every message (no drawing yet) |
| C Renderer: world | r_main r_bsp r_edge r_surf r_draw r_light r_efrag r_misc r_sky r_vars, d_init d_modech d_edge d_scan d_surf d_sky d_vars d_fill d_zpoint, nonintel | demo1 draws the world; screenshots; timedemo fps |
| D Renderer: models | r_alias r_aclip d_polyse, r_sprite d_sprite, r_part d_part | monsters, view weapon, sprites, particles in demo1 |
| E HUD, menus, screens | sbar, menu, screen (rest), view (blends, bob, roll) | full single player from the main menu: new game, save/load, intermission, finale |
| F Sound | snd_dma, snd_mem, snd_mix, new waveOut snd_win (as DoomVB) | effects in game; capture test without a device |
| G Network | net_dgrm, net_wins (Winsock UDP, as DoomVB's i_net), multiplayer menus | two copies on one PC play deathmatch on lqdm1 |
| H Beyond vanilla (optional) | raised limits, unsigned clipnodes, FitzQuake protocol 666, BSP2, Vorbis music | the full LibreQuake release, episodes 1 to 4 |

Clusters 0, A and B can be tested without a renderer (console output, hidden
runs, logs), so the first picture of a level comes with cluster C.

## Environment

* Project: `C:\Users\Owner\desktop\websites\quake` (git, branch `main`)
* VB6: `C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE` (`/make`
  works from the command line); resource compiler `VB98\Wizards\RC.EXE`
* Python 3.12 runs the tools (`python -I tools\...`)
* DoomVB for reference and reuse: `C:\Users\Owner\Desktop\websites\doom`
  (`src/i_video.bas` GDI window, `src/i_sound.bas` waveOut mixer,
  `src/i_net.bas` Winsock, `tools/build.py`, `docs/CONVENTIONS.md`)

## Build and test

```
python -I tools\check_names.py     VB name clashes, all at once (run before building)
python -I tools\build.py           QuakeVB.exe (QuakeVB.vbp is written from src\ and stubs\)
python -I tools\debugbuild.py      QuakeVB_debug.exe: a run time error logs a call trace
python -I tools\vbtest.py          unit tests (tests\*.bas)
```

VB6 reports only the first compile error (file and line in build.log); its
line numbers match the file's lines. Run `python -I tools\check_intdiv.py`
and `python -I tools\check_with.py` as well (the `\` precedence and the
With-block lock pitfalls in CONVENTIONS.md).

If QuakeVB.exe is running from the project folder (the user playing it),
the build cannot replace the exe and the running copy holds quakevb.log. Do
not stop it: copy `src`, `stubs` (and `res`) to `build_alt\`, run
`python -I tools\build.py <full path of build_alt>` and test with
`build_alt\QuakeVB.exe -basedir <project folder> ...`; its log is
`build_alt\quakevb.log`.

QuakeVB writes `quakevb.log` (everything Sys_Printf prints: startup, every
console line, every file found in a pak, errors) in the exe folder, through
WriteFile, so it is complete even if the process is killed.

Hidden scripted runs (no window, no MessageBox, real keyboard ignored):

```
QuakeVB.exe -hidden +wait +wait +screenshot +toggleconsole +quit
QuakeVB.exe -hidden -key 30 e -key 32 c -key 34 h -key 36 o -key 38 ENTER
python -I tools\typekeys.py --file keys.txt 20 2 "echo hi" "<ENTER>" "`" "quit" "<ENTER>"
QuakeVB.exe -hidden -keyfile keys.txt
python -I tools\pcx2png.py id1\quake00.pcx shot.png 2
```

* `+cmd` arguments run after quake.rc's `stuffcmds`; `wait` defers the rest
  of the command buffer by one frame.
* `-key <frame>[:<n>] <key>` presses a key at host_framecount `<frame>` and
  releases it n frames later (default 1). Keys are bind names: a lowercase
  letter, ENTER, SPACE, TAB, UPARROW, BACKSPACE, ESCAPE, SEMICOLON, `` ` ``.
  The command line keeps only 50 arguments (16 keys), so longer scripts go
  in a `-keyfile` (one `<frame> <key>` per line; typekeys.py writes it).
  A keyfile written by hand needs CRLF line ends: VB's `Line Input` does not
  end a line at a bare LF, so the whole file reads as one bad key.
* `F12` takes a screenshot (LibreQuake's binding). Screenshots go to the
  first free `id1\quakeNN.pcx`, so move old ones away first.
* Sound: `-sndcapture <file.wav>` (11025 Hz, 16 bit, stereo) records what
  would be played; analyze it with Python's `wave` module. A hidden run
  never opens the sound device or the CD.
* Network: run the second copy from a copy of QuakeVB.exe in another
  folder (quakevb.log takes one writer) and give each copy its own
  `-basedir` (an `id1` with hard links to the paks), so neither writes the
  user's config.cfg. Start the listen server (`+maxplayers 4 +map
  lq_e0m1`), then the client a few seconds later (`+connect 127.0.0.1`).
  A hidden run binds to 127.0.0.1 only; a visible run opens UDP sockets on
  the network (Windows Firewall may ask once), so visible tests that do
  not need the network take `-nolan`. `-netdrop n` loses every n-th
  packet a connection reads.
* A test that quits writes `id1\config.cfg`: archived cvars it changed
  (`viewsize`, `volume`, ...) stay changed for the user. Put them back.
* A test quits with `` ` `` (toggleconsole), then `quit` and ENTER at the
  console; through the menus it is ESCAPE, the QUIT entry, then `y`.
* Track where the console is: `map` (and `record ... <map>`) switch to the
  game, so the next script step opens the console with `` ` ``; `load`,
  `changelevel`, `restart` and `disconnect` leave it open, and a `` ` ``
  there closes it, so the rest is typed into the game (LibreQuake binds `t`
  to messagemode: the log then shows `say` lines). In the game, keys are
  LibreQuake's bindings: `w` is `+forward`, not UPARROW. A hold is a
  keyfile line `<frame>:<n> <key>`.
* A run under `timeout 60` that ends with exit code 124 was killed: it hung
  or never reached its quit.
* For the visible window, `tools\capture_window.ps1` starts QuakeVB with a
  keyfile, finds the form by its class `ThunderRT6FormDC` (VB's main window
  handle is a hidden 0x0 owner window) and saves it with CopyFromScreen.

## Deviations from WinQuake (all marked in the source)

* New Windows layers: `vid_win.bas` is a window with a palettized DIB
  section (no MGL, DirectDraw or VESA modes; `-width`, `-height`, `-scale`),
  `sys_win.bas` uses Win32 file handles and QueryPerformanceCounter. No
  startup dialog, no dedicated server console yet.
* The window procedure queues key and mouse events for `Sys_SendKeyEvents`
  instead of calling Key_Event itself (same timing). Alt+F4 and the close
  box quit without the "Are you sure?" box.
* The conback says `(QuakeVB) 1.09` (C: `(WinQuake) 1.09`); "Exe:" shows the
  exe's file time (C: `__TIME__ __DATE__`). The "megabyte heap" line and the
  memory checks are gone (no hunk).
* `CalcFov` returns 0 for an empty vrect instead of a NaN (see the NaN
  pitfall in CONVENTIONS.md; fov_y is unused by the software renderer).
* The main loop sleeps 1 ms when a frame was skipped for running faster
  than 72 fps (C spins), and `SCR_ModalMessage` sleeps while it waits.
* `Draw_BeginDisc` draws nothing yet (C draws the disc directly to the
  screen during file loads).
* Test options: `-hidden`, `-key`, `-keyfile`, `-sndcapture`, `-netdrop`.
* Savegames are written with VB's `Print #` (CRLF lines, as C's text mode
  on Windows); `load` reads the whole file and parses it with a word reader
  for the header (C: `fscanf`) and `COM_Parse` for the blocks (C reads each
  block up to its `}` into a buffer first; the parsers stop there anyway).
* Commands that dereference a NULL model in C (`viewframe`, `viewnext`,
  `viewprev` without a viewthing model) return instead.
* No net_vcr (`-record` / `-playback` of network sessions) and no serial
  or modem configuration (WinQuake has no serial driver either).
* Network (cluster G): no IPX driver (net_wipx.c; Windows has had no IPX
  since Vista), so `net_landrivers` has Winsock TCP/IP only. Winsock is
  called through Declare statements named as C's function vectors
  (`psocket`, `precvfrom`, ...; C loads wsock32.dll with GetProcAddress),
  error codes come from `Err.LastDllError`. WINS_Read also takes
  WSAECONNRESET as "no data" (how Windows 2000 and later report a UDP
  port nobody listens on; C knows WSAECONNREFUSED). The local name is
  looked up without C's 2-second blocking hook. A packet whose header
  length does not fit what arrived, or that would overflow the receive
  buffer, is dropped as a short packet (C copies what the header says).
  NET_SendToAll counts free player slots as done (C leaves their flags
  unset). Under `-hidden` without `-ip` every socket is bound to
  127.0.0.1, so a test run opens no port on the network.
* `Draw_TransPic` cuts the part of a picture that is off the screen (C
  stops with "bad coordinates"): LibreQuake's net menu pictures are 264
  pixels wide and drawn at x 72.
* No DirectInput (`-dinput`): the mouse is read with GetCursorPos, C's path
  without it. Under `-hidden` the mouse and the joystick are not started, so
  a test run cannot grab the real cursor or read a real joystick.
  `_windowed_mouse` ("Use Mouse") defaults to 1 (C: 0, since WinQuake's
  mouse was for fullscreen modes): the game is always a window, so a visible
  run takes the mouse pointer at startup.
* `NET_Shutdown` closes every active socket (C closes only the first: its
  loop follows `sock->next` after NET_FreeQSocket moved the socket).
* `playdemo` stops reading the CD track line at the end of the file (C's
  getc loop never ends there).
* Renderer (cluster C): one `R_DrawSurfaceBlock8` with the block size
  instead of C's four `_mip0`..`_mip3` copies. Not ported: the 16-bit
  drawers (`r_pixbytes` is always 1), the tile generators nothing calls
  (`R_GenTile`, `R_GenTurbTile`, `R_GenSkyTile`), the stack and alignment
  checks of `R_RenderView`, and `Sys_LowFPPrecision` /
  `Sys_HighFPPrecision` (x87 precision for the asm drawers).
  `D_InitCaches` allocates the surface cache itself (C gets a buffer from
  `VID_AllocBuffers`); `VID_Init` sizes it with `D_SurfaceCacheForRes`.
  `r_drawflat` colors a surface by the low byte of its `mod_surfaces()`
  index (C uses the low byte of the `msurface_t` pointer), so the colors
  differ from C's.
* Models (cluster D): `adivtab` is computed at the first draw (the same
  1,024 entries as adivtab.h, checked); `r_avertexnormals` is filled by
  `R_InitVertexNormals` (generated from anorms.h by `tools/gen_anorms.py`),
  called from `R_Init`. `R_AliasClip` picks the clip function from the clip
  flag (C passes a function pointer). D_PolysetRecursiveTriangle works on
  copies of the vertices, so the seam fixup is applied to the copies (C
  changes the finalverts and restores them). One particle pixel loop for
  every size (C unrolls sizes 1 to 4). `pointfile` reads the `.pts` file
  whole (C: fscanf). Not ported: D_PolysetFillSpans8, which nothing calls.
* Menus (cluster E): no serial or modem configuration menus (WinQuake has
  no serial driver, so the net menu never offers them) and no video modes
  menu (the window has no modes; "Video Options" stays hidden as in C
  without `vid_menudrawfn`). The window is always windowed, so "Use Mouse"
  is always offered. The setup menu cuts the name and hostname to 15
  characters (C copies them into 16-byte buffers). M_ScanSaves reads the
  first 256 bytes of each savegame for its version and comment (C:
  fscanf). Not ported: Sbar_UpdateScoreboard, which only C's `#if 0` code
  calls.
* Sound (cluster F): waveOut only (C's `-wavonly` path; there is no
  DirectSound). Under `-hidden` no sound device and no CD are opened
  (`-sndcapture <file.wav>` writes the mix to a file instead, with blocks
  completing in real time). The sound memory is per sound (no cache
  allocator; a loaded sound stays), so `Cache_Check` is "loaded". C's
  AppActivate sound blocking is checked on every VID_ProcessEvents call
  (a hidden run counts as active). The window procedure queues
  MM_MCINOTIFY for CDAudio_Update (C handles it in the window procedure).
  Reads past the end of a WAV file give 0. Not ported: `-simsound`
  (fakedma, which snd_win.c has no DMA position for) and the small-heap
  check that forces `loadas8bit`.
* `pr_edict_size` is a fixed 2048 bytes (512 Longs per edict: LibreQuake's
  105 entity fields with room for mods that add fields; `PR_LoadProgs`
  stops with an error if a progs needs more). C makes it `entityfields * 4`
  plus the edict_t header. QuakeC only sees it through edict references,
  which it never takes apart, and savegames store edict numbers.

## Next steps

1. The vanilla engine is complete (clusters 0 and A to G): single player
   and network play. Next in the plan: **cluster H, beyond vanilla**: the
   full LibreQuake release (episodes 1 to 4: raised limits, unsigned
   clipnodes, FitzQuake protocol 666, BSP2) and Vorbis music.
2. Not ported yet, and not in a cluster: the dedicated server
   (`-dedicated`, with its console window: conproc.c and sys_win.c's
   console input).
