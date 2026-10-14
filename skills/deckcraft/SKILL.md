---
name: deckcraft
description: Create and edit presentations (.pptx) with DeckCraft (PowerPoint-style, ArtCraft suite) through MCP, the CLI or the control channel, and set up, connect or troubleshoot its MCP server on Windows and macOS.
---

# DeckCraft for Claude

DeckCraft (ArtCraft suite, storytold/deckcraft) is a clean-room, open-source, pure-Rust reimplementation of **Microsoft PowerPoint**: ribbon, masters and layouts with placeholders, 8 original themes, 150+ preset shapes, connectors, tables, charts, basic SmartArt, audio/video, transitions (including Morph), animations, full-screen presenting and presenter view. It opens and saves `.pptx`, exports PDF (slides, notes pages, handouts), PNG/JPEG and outline. Status: **alpha**.

**Every action is a command** with a stable id (PowerPoint-ribbon style: `slide.new`, `format.bold`, `arrange.group`, `transition.set`, ...) and JSON parameters. UI, CLI, control channel and MCP all go through `Session::execute`, so they do exactly the same thing; programmatic calls **never open dialogs**. The docs speak of 226 commands (README: "200+"). If something disagrees with the installed version, trust `deckcraft-cli describe <id>` / `deckcraft-cli commands <filter>`.

## Quick check: is the MCP server connected?

1. Look at the tools you have. If tools from the MCP server `deckcraft` are available (for example `list_commands`, `execute` or `inspect_*` tools) the server is connected: skip to "Operating knowledge" below and use them directly.
2. If they are missing, check whether the user installed it under the name `deckcraft` or `deckcraft-headless`. In Claude Desktop the server may show as **disconnected**.
3. **Most common cause of "disconnected":** the live server (`mcp --connect {port}`) exits when DeckCraft is not running at the moment Claude starts. Fix: open DeckCraft with the **"DeckCraft (Claude)" launcher** (it starts the app with its control channel on), then re-enable the server in Claude's settings or fully quit and reopen Claude.
4. If the user does not need to watch the app live, **headless mode** always starts (no window, works on files). See "Live, headless or both".
5. If nothing is configured yet, follow "Set up the MCP server" below. Ask the user which OS they use if you do not know.

## Set up the MCP server (guide the user)

Walk the user through this step by step. Ask for the OS (Windows or macOS) first. If you have a shell on the user's machine you may run the steps yourself, but **ask before editing any config file and always back it up first**. Never print tokens or keys in chat.

### 1. Prerequisites
- Install DeckCraft from its upstream releases page: https://github.com/storytold/deckcraft/releases (Windows: the `.msi` installer; macOS: the `.dmg`, then drag the app to `/Applications`).
- Install Claude Desktop and/or Claude Code. For Claude Desktop, fully quit and reopen it after any config change.
- The MCP server is the **CLI** that ships with the app: `deckcraft-cli`. Windows installs it next to the app (`C:\Program Files\DeckCraft\deckcraft-cli.exe`, not on `PATH`). On macOS the CLI is a separate release download (asset `deckcraft-cli-<version>-macos-universal.zip`) unless it is bundled in the app (unverified): the setup script below fetches it into `~/.artcraft/bin/` and clears the macOS quarantine flag.
- This kit uses port **7990** (the upstream default), which does not clash with the other apps of the kit.

### 2. Option A: automatic setup (recommended)
The setup script finds the app and CLI, creates the launcher, backs up and updates Claude Desktop's config, registers the server in Claude Code (if the `claude` command exists) and installs this skill. Preview first with a dry run: nothing is changed.

Windows (PowerShell):
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "& ([scriptblock]::Create((irm https://raw.githubusercontent.com/Vinciux99/artcraft-claude-skills/main/install.ps1))) -Apps deckcraft"
```
macOS (Terminal):
```sh
curl -fsSL https://raw.githubusercontent.com/Vinciux99/artcraft-claude-skills/main/install.sh | bash -s -- --apps deckcraft
```
Or clone/download this repository and double-click `setup-windows.cmd` (Windows) or `setup-macos.command` (macOS; first time: right-click, Open). Useful options: `-Mode live|headless|both` / `--mode ...` (default `live`; `both` registers `deckcraft` live and `deckcraft-headless`), `-DryRun` / `--dry-run`, `-NoLauncher`, `-NoClaudeDesktop`, `-NoClaudeCode`, `-NoSkill` (macOS: `--no-launcher`, `--no-claude-desktop`, `--no-claude-code`, `--no-skill`).

What it changes: Claude Desktop config (`%APPDATA%\Claude\claude_desktop_config.json` plus the Microsoft Store/MSIX copy under `%LOCALAPPDATA%\Packages\Claude_*\LocalCache\Roaming\Claude\`; macOS `~/Library/Application Support/Claude/claude_desktop_config.json`; backup saved next to it as `<file>.backup-<timestamp>`), the launcher "DeckCraft (Claude)" (Windows: Desktop shortcut; macOS: `~/Applications/DeckCraft (Claude).app` plus a Desktop alias), Claude Code's user-scope MCP list, and `~/.claude/skills/deckcraft/`.

### 3. Option B: manual setup
**Windows** - edit `%APPDATA%\Claude\claude_desktop_config.json` (if Claude was installed from the Microsoft Store, also the copy under `%LOCALAPPDATA%\Packages\Claude_*\LocalCache\Roaming\Claude\`). Keep other servers; add this inside `mcpServers`. Backslashes must be doubled in JSON.
```json
{
  "mcpServers": {
    "deckcraft": {
      "command": "C:\\Program Files\\DeckCraft\\deckcraft-cli.exe",
      "args": [
        "mcp",
        "--connect",
        "7990"
      ]
    },
    "deckcraft-headless": {
      "command": "C:\\Program Files\\DeckCraft\\deckcraft-cli.exe",
      "args": [
        "mcp"
      ]
    }
  }
}
```
**macOS** - edit `~/Library/Application Support/Claude/claude_desktop_config.json`. Use the absolute CLI path (replace `YOUR_NAME`; JSON does not expand `~`). The CLI path is unverified if your install differs: check with `ls ~/.artcraft/bin /Applications/DeckCraft.app/Contents/MacOS`.
```json
{
  "mcpServers": {
    "deckcraft": {
      "command": "/Users/YOUR_NAME/.artcraft/bin/deckcraft-cli",
      "args": [
        "mcp",
        "--connect",
        "7990"
      ]
    },
    "deckcraft-headless": {
      "command": "/Users/YOUR_NAME/.artcraft/bin/deckcraft-cli",
      "args": [
        "mcp"
      ]
    }
  }
}
```
**Launcher** (starts the app with the control channel; needed before the live server can connect):
- Windows: create a Desktop shortcut named "DeckCraft (Claude)" with target `"C:\Program Files\DeckCraft\deckcraft.exe" --control 7990`
- macOS: run `open -n -a "/Applications/DeckCraft.app" --args --control 7990`, or save it as an app with Script Editor/`osacompile` (the setup script does this for you).

**Claude Code** (any OS; run once):
Windows (PowerShell):
```powershell
claude mcp add --scope user deckcraft -- "C:\Program Files\DeckCraft\deckcraft-cli.exe" mcp --connect 7990
claude mcp add --scope user deckcraft-headless -- "C:\Program Files\DeckCraft\deckcraft-cli.exe" mcp
```
macOS (Terminal):
```sh
claude mcp add --scope user deckcraft -- "$HOME/.artcraft/bin/deckcraft-cli" mcp --connect 7990
claude mcp add --scope user deckcraft-headless -- "$HOME/.artcraft/bin/deckcraft-cli" mcp
```
Check with `claude mcp list`.

### 4. Authentication
DeckCraft needs no token: the control channel listens on loopback (`127.0.0.1`) only and has no authentication. Anyone on the machine can drive the app while it is open, so close it when you are done.

### 5. Finish and verify
1. **Fully quit and reopen Claude** (Desktop: quit from the tray/menu bar, not just close the window).
2. Open DeckCraft with the **"DeckCraft (Claude)" launcher** before asking Claude to use it.
3. In Claude, confirm the server `deckcraft` is connected and call its `list_commands` tool (or equivalent). A non-empty command list means it works.

### Live, headless or both
- **Live** (`deckcraft`): drives the open app; the user sees every change and screenshots/UI tools work. Needs the launcher first.
- **Headless** (`deckcraft-headless`): in-process engine, no window; always starts; works on files; UI/screenshot tools are not available.
- **Both**: register both and pick per task. Switch with `-Mode` / `--mode`, or by editing the config entries above.

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| Server shows "disconnected" right after Claude starts | Live server exits when DeckCraft is not running | Open DeckCraft with the "DeckCraft (Claude)" launcher, then re-enable the server or restart Claude; or use headless mode |
| "Port already in use" / app opens but Claude cannot connect | Another app or an old DeckCraft instance holds port 7990 | Quit all DeckCraft windows, check the launcher uses the kit port, make sure no other ArtCraft app uses the same port |
| `command not found` / spawn error | CLI path wrong or the CLI is not installed (Windows: not on `PATH`; macOS: separate download) | Use the full absolute path in the config; macOS: run the setup script (it downloads the CLI to `~/.artcraft/bin/`) |
| macOS: "cannot be opened because the developer cannot be verified" | Gatekeeper/quarantine on the app or CLI | Right-click the app, Open; for the CLI: `xattr -dr com.apple.quarantine ~/.artcraft/bin/deckcraft-cli` |
| Anyone else can drive the app | The control channel has no authentication (loopback only) | Start the app with the control channel only when needed and quit it when done |
| Config edited but nothing changed (Windows) | Claude installed from the Microsoft Store reads the MSIX copy of the config | Edit/update `%LOCALAPPDATA%\Packages\Claude_*\LocalCache\Roaming\Claude\claude_desktop_config.json` too |
| JSON config error / server not listed | Unescaped backslashes or trailing comma | Double the backslashes on Windows; validate the JSON; restore the `.backup-*` file if needed |
| App crashes on start or window is blank | Graphics backend problem | Windows: set `WGPU_BACKEND=vulkan` (or `dx12`) for one launch; avoid `gl` on some AMD drivers |
| Changes do not appear in the window | You are on the headless server | Use the live server (launcher first) |
| Quotes disappear in PowerShell JSON arguments | Windows PowerShell 5.1 quoting for `.exe` | Use `--script`/files where the CLI supports them, or single-quote the whole JSON |

## Operating knowledge

## App, settings, environment
- **App flags** (from the source): `--control <port>`, `--sample` (opens the "Northern Lights" sample deck), `--show` (start presenting at once), `--version`; files as positional arguments. The "DeckCraft (Claude)" launcher runs the app with `--control 7990`.
- **Settings folder:** Windows `%APPDATA%\DeckCraft`; macOS `~/Library/Application Support/DeckCraft` (unverified): `ui.json`, `prefs.json`, `logs/`.
- **Log:** `<settings>/logs/deckcraft.log`; at each start the previous one becomes `deckcraft.1.log` (then `.2.log`); max 16 MiB. Contains panics, failed recoveries, PPTX parts skipped at import, audio problems. `--version` writes no log.

| Variable | Effect |
|---|---|
| `RUST_LOG` | Log levels (default `info` for DeckCraft, `warn` for the rest), e.g. `RUST_LOG=warn,deckcraft_pptx=trace` |
| `WGPU_BACKEND` | Force the graphics backend (`dx12`, `vulkan`, `metal`); Windows default = DirectX 12 only. Avoid `gl` (AMD driver crash `atio6axx.dll`) |
| `DECKCRAFT_CONTROL_PORT` | Same as `--control <port>` |
| `DECKCRAFT_NO_RECOVERY` | Disable AutoRecover at startup |
| `DECKCRAFT_NO_PREFS` | Do not read/write `ui.json`/`prefs.json` |
| `DECKCRAFT_NO_AUDIO` | Silent media |

- If the app flashes and closes (known GPU crashes), from PowerShell in the exe folder: `$env:WGPU_BACKEND = "vulkan"; & .\deckcraft.exe`, then `Remove-Item Env:WGPU_BACKEND`.
- **AutoRecover:** at startup the app reopens what the previous session had not saved (`file.recovery.open`); manage with `file.recovery.list` / `file.recovery.discard`.

## Connecting

| Mode | Start | What it drives |
|---|---|---|
| MCP headless | `deckcraft-cli mcp` (optional `--sample`) | In-process engine session, no window; **starts with no presentation**: call `new_presentation` or `open_presentation` first |
| MCP attached to the app | `deckcraft-cli mcp --connect 7990` | The app started with `deckcraft --control 7990`; agent and person work on the same deck |
| CLI | `deckcraft-cli <subcommand>` | Headless `run`/`render`/`convert`/`info`; `app --port 7990 ...` toward the live app |
| Control channel | `deckcraft --control 7990` (or `DECKCRAFT_CONTROL_PORT`) | JSON lines over TCP `127.0.0.1:7990` |

- **Port:** 7990 is the one in all the upstream examples and does not collide with the other ArtCraft apps in this kit. **Careful:** the internal default of `deckcraft-cli app` (and the example in the MCP source) is **7979**, which here belongs to VectorCraft: **always** pass `--port 7990` / `--connect 7990`, or you end up talking to the wrong app.
- **`--connect` accepts** a port (→ `127.0.0.1:<port>`) or `host:port`. The MCP server connects **at startup**: if the app is not already open with `--control` it exits with "can't reach the app". Afterwards, if the app restarts, it retries once with a new connection. (So open the app with the launcher *before* starting Claude.)
- **Authentication: none** (no token, no file). The server listens on loopback only, but anyone on the machine can command the app: start `--control` only when needed.
- **Read/write roots: none.** Paths in `file.open`, `file.export`, `insert.picture`, `ui.screenshot {path}`, etc. are read/written wherever the user can, on the app's machine. Choose paths carefully.
- **Timeouts:** the app answers within 60 s (then `{"ok":false,"error":"timeout"}`); the MCP client waits up to 90 s.

### Choosing the channel
1. **A DeckCraft MCP server is available** → use it (tools below): `render_slide`/`screenshot` return images to look at.
2. **A shell on the user's machine** → `deckcraft-cli` with its full path; for the open app `deckcraft-cli app --port 7990 ...`.
3. **A cloud sandbox** cannot run the user's app or reach its ports; it can only prepare files (images, outline `.txt`/`.md` for `slide.fromOutline`, lists of `--cmd`, parameter JSON).
4. **Last resort: computer use on the GUI.** Non-US keyboards: shortcuts with `[` `]` (`arrange.bringForward` Ctrl+], `arrange.sendBackward` Ctrl+[) and `` ` `` may misfire: use the ribbon/menus. In the docs `Cmd` = `Ctrl` on Windows.

## Ground rules
- **Units:** slide points (1/72 in), y down; a 16:9 slide is **960 × 540**. `rect` = `[x, y, w, h]` in pt. The coordinates of `ui.click/move/drag` are instead **window** points (use `ui.inspect` → `slideRect`, `canvasScale` to convert).
- **Indices:** slides by **0-based index** (`slide.go {index}`, `render --slide 0` is the first; `render --all` writes `slide-01.png`...). Slides also have an `id` (`slide.go {id}`, `slide.delete {ids}`). Shapes: **integer numeric ids** read from `inspect_slide`/`slide.inspect` or returned by commands (`shape.insert` → `{"id": 258}`).
- **Current slide:** shape commands act on the current slide: first `slide.go {index}` / `go_to_slide`.
- **Targets:** nearly all shape/format commands accept `ids` (or `id`); without them they act on the **selection**. If a command is disabled for "nothing selected" but you pass `ids`/`id`, it runs anyway. `edit.select {ids:[...], add?, toggle?}` to select. Locked shapes (`shape.lock`) refuse move/resize/rotate/flip/crop.
- **Colors:** `"#RRGGBB"`, theme slots (`accent1`...`accent6`, `tx1`...), preset names, or `{"scheme":"accent2","lumMod":75000}`.
- **Text:** in `slide.new`/`text.set` one line = one paragraph, leading tabs = list level; in `text.insert` `\n` = new paragraph, `\u000b` = line break within a paragraph.
- **Layouts** (`slide.new`, `slide.layout`): `title`, `titleAndContent`, `sectionHeader`, `twoContent`, `comparison`, `titleOnly`, `blank`, `contentWithCaption`, `pictureWithCaption` (or the master's layout name).
- **Themes** (`file.new {theme}`, `design.theme {name}`): DeckCraft, Harbor, Ember, Meadow, Nocturne, Paper, Coral Reef, Slate (`design.themes` lists them with color schemes/fonts).
- **Undo:** every modifying command is **one** undo step (`edit.undo`/`edit.redo`, 500-entry history); `query` commands (inspect, list, get, render...) do not touch history; file, navigation, selection and playback are not undoable.
- **Parameters:** the engine reads them (optional keys with defaults); the source shows no rejection of unknown keys, so a typo may be **silently ignored**: always verify with `inspect_slide`/`render_slide`. Exact format: `describe <id>`.
- **Batch:** `batch` runs in order and stops at the first error (unless `keepGoing`). The docs do **not** provide `$N` references to earlier results: if you need the id of a just-created shape, make one call, read the id, then continue.
- **Long jobs:** no async jobs or progress; heavy PDF/PNG exports run inside the call (60 s app-side limit). For many images prefer headless `deckcraft-cli render`.
- **Traps:** `file.save` without `path` **overwrites the open file** (even an original `.pptx`); `app.quit` closes **without asking** (`{"save":true}` saves first); `edit.paste`/pasting images from the clipboard does not work on Windows; `ui.*`, `view.*`, `app.*`, `screenshot`, `call_app` need the app (`--connect`), in headless they error.

## MCP tools (25)

| Tool | Arguments | Notes |
|---|---|---|
| `list_commands` | `filter?` | id, label, menu, shortcut, params, enabled (+ reason) |
| `run_command` | `command`, `params?` | Any command: the general route |
| `batch` | `commands:[{command, params?}]`, `keepGoing?` | Returns each result |
| `new_presentation` | `theme?`, `blank?` | = `file.new` (one title slide; `blank` = none) |
| `open_presentation` | `path` | `.deckcraft` or `.pptx` |
| `save_presentation` | `path?` | Format from extension; without path overwrites |
| `export` | `path`, `format?`, `slide?`, `all?`, `scale?`, `layout?` (slides/notes/handouts), `perPage?` (1,2,3,4,6,9), `dpi?` (36-600), `slides?`, `includeHidden?`, `textLayer?`, `frame?` | = `file.export`; formats png, jpeg, pptx, deckcraft, outline, pdf |
| `inspect_document` | — | Slides (index, id, title, layout, transition, animations, notes), sections, theme, layouts, media, selection |
| `inspect_slide` | `index?` | Shapes with id, type, box `[x,y,w,h]`, placeholders, text, fills, tables, charts |
| `render_slide` | `slide?`, `scale?` (1 = 960×540), `edit?` | PNG to look at; works headless too |
| `add_slide` | `layout?`, `title?`, `body?` | = `slide.new`, after the current one |
| `go_to_slide` | `index` | = `slide.go` |
| `add_shape` | `preset`, `x`, `y`, `w`, `h`, `text?`, `fill?`, `line?` (color or `none`) | `shape.insert` + `shape.fill`/`shape.line` |
| `add_text_box` | `x`, `y`, `w`, `text`, `h?`, `size?`, `color?`, `align?` | Height grows with the text |
| `set_text` | `id`, `text` | = `text.set`, keeps the first run's format |
| `insert_picture` | `path` or `data` (base64) + `name?`, `x? y? w? h?` | = `insert.picture` |
| `insert_table` | `rows`, `cols`, `data?` (rows of strings), `x? y? w? h?` | = `insert.table` |
| `insert_chart` | `type?`, `title?`, `categories?`, `series?:[{name, values}]` | = `insert.chart` |
| `select` | `ids` | = `edit.select` |
| `pointer` | `events:[{kind, x, y, mods?}]`, `mods?` | Gestures of the active tool, in slide pt |
| `select_tool` | `tool`, `preset?` | select, textBox, shape (+preset), pen, highlighter, eraser or preset name |
| `key` | `key`, `mods?` (`shift`, `alt`, `cmd`) | Keyboard (also while presenting) |
| `type_text` | `text` | Into the text being edited |
| `screenshot` | — | Whole window; **app only** |
| `call_app` | `method`, `params?` | Any control-channel method; **app only** |

**Resources:** `deckcraft://document` (= `document.inspect`) and `deckcraft://commands` (= `engine.commands`), JSON. No prompts. MCP protocol 2025-06-18 (also accepts 2025-03-26 and 2024-11-05).

## Recipes

New deck, slide with layout, text, shape, check (MCP):
```json
new_presentation {"theme": "Harbor"}
add_slide {"layout": "titleAndContent", "title": "Q3 results", "body": "Revenue +12%\n\tInternational +20%\nNew customers: 34"}
add_shape {"preset": "roundRect", "x": 600, "y": 380, "w": 300, "h": 100, "text": "Target reached", "fill": "accent2", "line": "none"}
inspect_slide {}
render_slide {"scale": 1}
```

The same via commands (`run_command`/`batch`):
```json
{"command": "file.new", "params": {"theme": "Harbor"}}
{"command": "slide.new", "params": {"layout": "titleOnly", "title": "Hello"}}
{"command": "shape.insert", "params": {"preset": "star5", "rect": [100, 100, 200, 200], "text": "Hi"}}
{"command": "shape.fill", "params": {"ids": [258], "gradient": {"stops": [[0, "#1E3A8A"], [1, "accent1"]], "angle": 90}}}
{"command": "shape.line", "params": {"ids": [258], "none": true}}
{"command": "shape.effects", "params": {"ids": [258], "shadow": "outer"}}
```

- **Text and format:** `text.set {"id": 12, "text": "Title"}`; then `edit.select {"ids":[12]}` and `format.font {"family": "Inter"}`, `format.size {"size": 40}`, `format.bold {"on": true}`, `format.color {"color": "accent1"}`, `format.align {"align": "center"}`, `format.bullets {"on": true, "char": "➢"}`, `format.autofit {"mode": "shrink"}`.
- **Picture:** `insert.picture {"path": "/Users/me/Pictures/logo.png", "rect": [40, 40, 200, 120]}`; `picture.crop {"left": 0.1, "right": 0.1}` (fractions 0..1), `picture.adjust {"brightness": 0.1}`.
- **Table:** `insert_table {"rows": 3, "cols": 2, "data": [["Item","Value"],["A","10"],["B","20"]]}`; `table.style {"style": "medium2-accent1", "id": ...}`, `table.merge {"from": [0,0], "to": [0,1]}`.
- **Chart:** `insert.chart {"type": "column", "categories": ["Q1","Q2","Q3"], "series": [{"name": "Revenue", "values": [10, 14, 18]}]}`; then `chart.options {"title": "Revenue", "legend": "b", "dataLabels": true}`, `chart.data {...}`.
- **Connector:** `shape.connect {"from": 258, "to": 260, "preset": "bentConnector3"}` (stays glued to shapes; sites via `shape.sites {id}`).
- **Theme and background:** `design.theme {"name": "Nocturne"}`, `design.colors {"colors": {"accent1": "#E4572E"}}`, `design.fonts {"major": "...", "minor": "..."}`, `design.background {"color": "#101820", "all": true}`, `design.slideSize {"preset": "standard"}`, `design.headerFooter {"slideNumber": true, "footer": true, "footerText": "Acme Inc."}`.
- **Notes and sections:** `slide.notes {"text": "...", "index": 2}`, `section.add {"name": "Intro", "at": 0}`.
- **From an outline:** `slide.fromOutline {"text": "..."}` adds slides after the current one and returns `{"slides": n}`. The outline text format is not documented (presumably one line per title, tabs for levels as in `slide.new`): try it on a copy and verify with `inspect_document`.
- **Transitions/animations:** `transition.set {"kind": "morph", "duration": 1000}`, `transition.applyAll {}`; `animation.add {"effect": "fade", "class": "entrance", "start": "afterPrevious", "ids": [258]}`, `animation.options {"textBuild": "byParagraph"}`. Lists: `transition.list`, `animation.list`.
- **Export:** `file.export {"path": "/Users/me/Desktop/deck.pdf", "format": "pdf"}`; handouts `{"path": ".../handout.pdf", "layout": "handouts", "perPage": 3}`; PPTX `{"path": ".../deck.pptx", "format": "pptx"}`; all slides as PNG `{"path": ".../out.png", "format": "png", "all": true, "scale": 2}`.
- **Quality control:** `review.accessibility {}`, `review.spelling {}`, `file.info {}`.

## CLI
```text
deckcraft-cli info FILE                        # slides, titles, layouts (JSON)
deckcraft-cli render (FILE | --sample) [--slide N | --all] [--scale S] [--edit] OUT.png|OUT_DIR
deckcraft-cli render-transition FILE --slide N --t T[,T...] [--width PX] OUT.png
deckcraft-cli convert IN OUT                   # .deckcraft <-> .pptx, outline .txt, .png (first slide)
deckcraft-cli run [--in FILE | --sample] [--cmd ID[=JSON]]... [--save OUT] [--export OUT] [--print]
deckcraft-cli commands [FILTER]                # JSON list
deckcraft-cli describe ID                      # one command: where it lives, shortcut, params
deckcraft-cli app [--port PORT] COMMAND [JSON] | --method METHOD [JSON]
deckcraft-cli mcp [--connect PORT] [--sample]
deckcraft-cli links | --version
```
- `run` without `--in`/`--sample` starts from `file.new`; `--cmd`s run in order and the first error stops; `--print` shows each result; `--save` = `file.saveAs`, `--export` = `file.export`.
- `render`: `--scale` default 2 (px per pt); `--all` needs a folder. `--edit` shows placeholder prompts as in the editor.
```sh
deckcraft-cli commands format.
deckcraft-cli describe shape.fill
deckcraft-cli run --sample --cmd 'slide.new={"layout":"titleOnly","title":"Hello"}' --save hello.deckcraft
deckcraft-cli run --in draft.pptx --cmd 'design.theme={"name":"Slate"}' --save draft-slate.pptx --export draft.pdf
deckcraft-cli render deck.pptx --all --scale 1 out/
deckcraft-cli convert deck.deckcraft deck.pptx
deckcraft-cli app --port 7990 --method ui.inspect
deckcraft-cli app --port 7990 slide.new '{"title":"From the terminal"}'
```
(Windows: `& "C:\Program Files\DeckCraft\deckcraft-cli.exe" ...`.) In Windows PowerShell 5.1 the quotes inside JSON passed to an exe can get lost: if the JSON arrives broken ("bad JSON"), write `\"` inside the string or use cmd/PowerShell 7.

## Control channel
- **Transport:** TCP on `127.0.0.1:<port>` only (`--control <port>`), JSON lines: one request object per line, one response per line, several simultaneous clients possible. No auth.
- **Format:** `{"id": 1, "method": "...", "params": {...}}` → `{"id": 1, "ok": true, "result": ...}` or `{"id": 1, "ok": false, "error": "..."}`. Malformed JSON → `bad JSON: ...`; a response taking over 60 s → `"timeout"`.

| Method | Params | What it does |
|---|---|---|
| `engine.execute` / `ui.menu.invoke` / `command` | `{command, params}` | Run any command (engine or UI) |
| `engine.commands` | — | All commands with enablement (UI commands have `"ui": true`) |
| `document.inspect` | — | Slides, sections, layouts, shapes, selection |
| `ui.inspect` | — | UI state, tool, window/canvas/slide rects, `canvasScale`, dialog, `home`, `backstage`, running show, documents (title, dirty, path), status, perf |
| `ui.tool.select` | `{tool, preset?}` | `select`, `text`/`textBox`, `shape`+preset, preset name, `curve`, `freeform`, `scribble`, `pen`, `highlighter`, `eraser` |
| `ui.pointer` | `{events:[{kind, x, y, mods?}], mods?}` | kind `down`/`drag`/`up`/`move`/`doubleclick`/`tripleclick`, in slide pt |
| `ui.key` / `ui.text` | `{key, shift?, alt?, cmd?}` / `{text}` | Keyboard (also drives the show) |
| `ui.ime` | `{preedit, active?}` / `{commit}` | IME composition |
| `ui.move` / `ui.click` / `ui.drag` | `{x, y}`; click `button?` (`right`), `count?` 1-3; drag `toX, toY, steps?` | Real egui input in window points: reaches every widget |
| `ui.set` | `{tab?, view?, pane?, notes?, zoom?, dark?, ribbonCollapsed?, formatTab?}` | UI state (returns the full state); `view`: `sorter`/`outline`/`notesPage`/normal |
| `ui.dialog.open` / `ui.dialog.close` | `{id, fields?}` | Open/close a dialog by id |
| `ui.screenshot` | `{path?}` | Window PNG (base64 without path) |
| `ui.render` | `{path?, slide?, scale?, edit?}` | One rendered slide (base64 without path) |
| `ui.resize` / `ui.focus` | `{width, height}` / — | Window; only `ui.focus` steals focus |
| `show.start` / `show.end` | `{from?, reading?}` | Presentation |
| `app.open` / `app.save` / `app.export` / `app.quit` | `{path, ...}` | Files; export takes `file.export` params; quit `{save?}` |

```json
{"id":1,"method":"engine.execute","params":{"command":"shape.insert","params":{"preset":"roundRect","rect":[100,100,240,120],"text":"Hello"}}}
{"id":2,"method":"ui.screenshot","params":{"path":"/Users/me/Desktop/shot.png"}}
{"id":3,"method":"app.export","params":{"path":"/Users/me/Desktop/deck.pdf","layout":"handouts","perPage":3}}
```

**UI-only commands** (via `engine.execute` with the app): `view.normal`, `view.outline`, `view.sorter`, `view.notesPage`, `view.reading`, `view.notes`, `view.ruler`, `view.gridlines`, `view.guides`, `view.zoomIn`, `view.zoomOut`, `view.zoom {percent}`, `view.fit`, `view.pane {pane: format|animation|selection|comments|designer|background|null}`, `view.tab`, `view.collapseRibbon`, `view.dark {on?}`, `view.grayscale`, `show.start {from?}`, `show.presenter {on?}`, `show.end`, `app.home`, `app.backstage {page?}`, `app.closePresentation`, `app.print`, `app.about`, `app.palette`, `app.preferences`, `app.openDialog`, `app.insertPictureDialog`, `app.insertAudioDialog`, `app.insertVideoDialog`, `app.saveAsDialog`, `app.exportDialog`, `app.dialog {id}`, `app.links`. The "...Dialog" ones open native dialogs: avoid them in automation, use commands with `path`.

## Commands by family
Real ids from the registry. For the rest and for parameters: `list_commands {filter}` / `deckcraft-cli commands <filter>` / `describe <id>`.

- **file.** `new`, `open`, `openBytes`, `save`, `saveAs`, `saveTemplate`, `saveBytes`, `export`, `info`, `render`, `close`, `properties`, `revert`, `recovery.save|list|open|discard`
- **edit.** `undo`, `redo`, `cut`, `copy`, `paste`, `pasteText`, `duplicate`, `delete`, `selectAll`, `deselect`, `select` · **selection.** `select`, `showAll`, `hideAll`
- **slide.** `new`, `duplicate`, `delete`, `move`, `go`, `next`, `previous`, `first`, `last`, `selectSlides`, `hide`, `layout`, `reset`, `notes`, `rename`, `fromOutline`, `inspect` · **section.** `add`, `rename`, `remove`, `removeAll`, `move`
- **insert.** `textBox`, `picture`, `table`, `chart`, `audio`, `video`, `wordArt`, `slideNumber`, `dateTime`, `symbol`, `hyperlink`, `smartArt` (`list|process|cycle|hierarchy|pyramid|matrix`), `actionButton`
- **shape.** `insert`, `presets`, `connect`, `freeform`, `merge` (`union|combine|fragment|intersect|subtract`), `sites`, `move`, `resize`, `setBounds`, `rotate`, `flip`, `adjust`, `change`, `fill`, `line`, `effects`, `quickStyle`, `rename`, `altText`, `visible`, `lock`, `setDefault`, `inspect`
- **picture.** `crop`, `adjust`, `reset`, `change` · **media.** `play`, `pause`, `toggle`, `stop`, `seek`, `info`, `posterFrame`, `gifPlay`, `options`
- **arrange.** `bringToFront`, `bringForward`, `sendBackward`, `sendToBack`, `reorder`, `group`, `ungroup`, `regroup`, `align`, `distribute`, `rotateLeft`, `rotateRight`, `flipHorizontal`, `flipVertical`, `nudge`
- **text.** `edit`, `exit`, `insert`, `delete`, `commit`, `move`, `select`, `selectAll`, `selectWord`, `selectParagraph`, `set`, `get`
- **format.** `bold`, `italic`, `underline`, `strikethrough`, `superscript`, `subscript`, `font`, `size`, `grow`, `shrink`, `color`, `highlight`, `case`, `caps`, `spacing`, `clear`, `textShadow`, `textOutline`, `align`, `alignLeft|Center|Right`, `justify`, `rtl`, `bullets`, `numbering`, `indent`, `outdent`, `lineSpacing`, `paragraph`, `anchor`, `direction`, `columns`, `autofit`, `wrap`, `margins`, `painter`, `painterApply`, `state`
- **design.** `theme`, `themes`, `colors`, `fonts`, `background`, `hideBackgroundGraphics`, `slideSize`, `headerFooter` · **view.** `slideMaster`, `closeMaster` · **master.** `insertLayout`, `renameLayout`, `deleteLayout`, `insertPlaceholder`
- **table.** `insertRowAbove|Below`, `insertColumnLeft|Right`, `deleteRow`, `deleteColumn`, `merge`, `split`, `style`, `options`, `cellFill`, `distributeRows|Columns`, `columnWidth`, `rowHeight`, `selectCells` · **chart.** `type`, `data`, `options`
- **transition.** `set`, `options`, `timing`, `applyAll`, `list` · **animation.** `set`, `add`, `remove`, `move`, `timing`, `options`, `clear`, `painter`, `painterApply`, `list`, `get`
- **show.** `fromStart`, `fromCurrent`, `setup`, `customShow` · **comment.** `add`, `reply`, `delete`, `resolve`, `list` · **review.** `accessibility`, `spelling`
- **inspect:** `document.inspect`, `slide.inspect`, `shape.inspect`, `commands.list` · **window.** `next`

## Recommended workflow
1. **Orient:** `inspect_document` (or `ui.inspect` with the app) to see what is open, how many slides, which is current and whether the document is `dirty`.
2. **Work on a copy:** open the original and immediately `file.saveAs` to a new path (e.g. `.../deck-v2.pptx`), or `deckcraft-cli run --in orig.pptx ... --save copy.pptx`. Never `file.save` without a path on a user's file.
3. **Find commands:** `list_commands {filter}` and `describe <id>`; check `enabled` / `disabled_reason` before insisting.
4. **Edit in small steps:** `slide.go`, then a few commands (or a short `batch`), reading ids from `inspect_slide`. If you slip, `edit.undo`.
5. **Verify visually** after each significant step: `render_slide` (also headless) or `screenshot` / `ui.screenshot` with the app; check overflowing text, overlaps, contrast.
6. **Save/export only with the user's OK**, to a new path; never overwrite original `.pptx`/`.deckcraft`.
7. **Close:** the control channel has no authentication, so when done close the app started with `--control` (or ask the user to); `app.quit` closes without asking, use it only after saving.
