# QuakeVB

**Quake, ported line by line to Visual Basic 6.** The 1996 engine is
complete: single player with sound, and network play over UDP. Next is
optional work beyond the original's limits, for the full LibreQuake
release.

QuakeVB is a port of id Software's GPL Quake source release (the
software-rendered WinQuake) to classic Visual Basic 6, done the same way as
DoomVB: one VB module for every C file of the original, translated
statement by statement and compiled to a native Windows executable. It plays
[LibreQuake](https://github.com/lavenderdotpet/LibreQuake), a free
replacement for the Quake data, which is included in `id1/`.

## Status

| Part | State |
|---|---|
| Project, plan, feasibility tests | done |
| 0 Foundation: file system, console, cvars, commands, window | done |
| A Server and QuakeC virtual machine | done |
| B Client and local connection | done |
| C Renderer: world | done |
| D Renderer: models, sprites, particles | done |
| E Status bar, menus, screens | done |
| F Sound | done |
| G Network play | done |
| H Beyond vanilla limits (full LibreQuake release) | optional |

The plan, the assessment and the measurements are in
[docs/RESUME.md](docs/RESUME.md); the translation rules are in
[docs/CONVENTIONS.md](docs/CONVENTIONS.md).

## Playing

Run `QuakeVB.exe`. The attract demos play until you press **Esc**, which
opens the main menu; choose **Single Player**, then **New Game**. You start
in the hub map, where the hallways choose the skill and lead to episode 0.

The game runs in a window, 320×240 shown at twice that size. Options for
the command line:

| Option | Effect |
|---|---|
| `-width N` | Render N pixels wide (the height follows at 4:3) |
| `-height N` | Render N pixels high |
| `-scale N` | Show the window at N times the render size |
| `-nosound`, `-nocdaudio` | No sound effects, no CD music |
| `-nomouse`, `-nojoy` | Leave the mouse or the joystick alone |
| `-nolan` | No network at all |
| `-port N` | Use UDP port N instead of 26000 |
| `-ip A.B.C.D` | Use this network address of the PC |

### Controls

These are LibreQuake's default keys. Change them in **Options → Customize
Controls**; your keys and settings are saved to `id1/config.cfg` when you
quit.

The window holds the mouse pointer while the game is in front; press
**Pause** or **Alt+Tab** to let go of it. To play with the keyboard alone,
switch off **Use Mouse**, the last line of the Options menu (press **Up**
from the top line to reach it), or type `_windowed_mouse 0` in the console.

| Movement | |
|---|---|
| W / S | Forward / back |
| A / D | Step left / right |
| Mouse | Look and turn (mouse look is on) |
| Space or right mouse button | Jump; swim up |
| E | Swim up |
| C or Ctrl | Swim down |
| Shift (hold) | Run (or turn on **Always Run** in Options) |
| Alt (hold) | Step sideways with the mouse |

| Weapons | |
|---|---|
| Left mouse button | Fire |
| 1 to 8 | Axe, Shotgun, Super Shotgun, Nailgun, Super Nailgun, Grenade Launcher, Rocket Launcher, Lightning Gun |
| Q, / or mouse wheel up | Next weapon |
| Mouse wheel down | Previous weapon |
| ↑ | Shotgun; press again for the Super Shotgun |
| → | Nailgun; press again for the Super Nailgun |
| ← | Grenade Launcher; press again for the Rocket Launcher |
| ↓ | Lightning Gun |

| Game | |
|---|---|
| Esc | Menu |
| Tab (hold) | Scores, kills, secrets and time |
| F1 | Help |
| F2 / F3 | Save / load menu |
| F4 | Options |
| F6 / F9 | Quick save / quick load |
| Pause | Pause |
| = or + | Bigger view |
| - | Smaller view |
| F12 | Screenshot (`id1/quakeNN.pcx`) |
| \` (backquote) | Console |
| T or Enter | Chat message (network play) |

LibreQuake's config also binds mouse button 5 (melee), Print Screen, F10
(graphics style) and gamepad buttons for newer engines. The 1996 engine has
no such keys or commands, so those do nothing here.

The classic console cheats work in single player: `god`, `noclip`, `fly`,
`notarget`, `impulse 9` (all weapons and ammo), `map start` (back to the
hub) and `skill 0` to `skill 3` (takes effect at the next map).

### Network play

QuakeVB plays deathmatch and cooperative games over UDP (port 26000) with
WinQuake's network protocol. So far it has been tested between two copies
on one PC.

* **Host a game:** Multiplayer → New Game → TCP/IP → OK, set the players,
  game type and level, then **begin game**.
* **Join a game:** Multiplayer → Join a Game → TCP/IP, type the host's
  address under **Join game at:** and press Enter, or type `connect
  <address>` in the console. **Search for local games** looks for
  servers on your network; like WinQuake's, it skips your own PC.
* **Which address:** like WinQuake, the host listens on the first address
  Windows gives for the PC's name. On a PC with virtual network adapters
  (VirtualBox, Hyper-V, WSL, a VPN) that may not be the LAN address, so
  start the host with `-ip` and its LAN address (`ipconfig` lists it), for
  example `QuakeVB.exe -ip 192.168.1.20`. The console command `status`
  shows the address in use.
* **Two copies on one PC:** start both with `-ip 127.0.0.1`, host on one
  and join `127.0.0.1` on the other. The second copy keeps no log file. A
  third copy would push the second one out (WinQuake's rule for two
  players from one address).

The first time QuakeVB opens a network socket, Windows Firewall may ask
whether to allow it. A dedicated server (`-dedicated`) is not ported yet.

## Game data

`id1/pak0.pak` and `id1/pak1.pak` are the **lite** package of LibreQuake
v0.09-beta: the start map and episode 0, which fit the limits of the 1996
engine. The full release (episodes 1 to 4) needs engine limits raised beyond
vanilla, planned as the last part. LibreQuake's art, models and sounds are
under the BSD license and its game code (`progs.dat`) under the GPL; see
[librequake/](librequake/).

## Folders

| Folder | Contents |
|---|---|
| `src/` | The VB6 source, one module per C file |
| `stubs/` | Stand-ins for the clusters not ported yet |
| `tests/` | VB unit tests (`python -I tools\vbtest.py`) |
| `tools/` | `build.py`, `debugbuild.py`, `check_names.py`, `vbtest.py` and the test helpers (listed in docs/CONVENTIONS.md), `probe/` (the feasibility benchmark) |
| `docs/` | Plan and status (RESUME.md), translation rules (CONVENTIONS.md) |
| `id1/` | Game data (LibreQuake lite) |
| `librequake/` | LibreQuake's licenses and credits |
| `ref/WinQuake/` | The original C source, for reference |

## License

QuakeVB is free software, licensed under the **GNU General Public License,
version 2**; see [LICENSE.TXT](LICENSE.TXT). It is a derivative of the Quake
source code by id Software, which id released under the GPL in 1999.

The included game data is LibreQuake, Copyright © 2019-2023 Contributors to
the LibreQuake project: a separate work, distributed under the BSD license
(art, models, sounds) and the GPL (`progs.dat`), as described in
[librequake/README-IMPORTANT-LICENCE-INFO](librequake/README-IMPORTANT-LICENCE-INFO).

QuakeVB is not affiliated with or endorsed by id Software, ZeniMax Media or
Bethesda. Quake is a registered trademark of id Software LLC.
