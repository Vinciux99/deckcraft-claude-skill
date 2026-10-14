# DeckCraft skill for Claude

DeckCraft (ArtCraft suite, storytold/deckcraft) is a clean-room, open-source, pure-Rust reimplementation of **Microsoft PowerPoint**: ribbon, masters and layouts with placeholders, 8 original themes, 150+ preset shapes, connectors, tables, charts, basic SmartArt, audio/video, transitions (including Morph), animations, full-screen presenting and presenter view. It opens and saves `.pptx`, exports PDF (slides, notes pages, handouts), PNG/JPEG and outline. Status: **alpha**.

This repository packages a **skill** that teaches Claude to drive DeckCraft (part of the [ArtCraft suite](https://github.com/storytold/deckcraft) by storytold) and that **knows how to create the MCP server and guide you through it**, on **Windows and macOS**. It is unofficial and not affiliated with storytold.

> Part of the [ArtCraft Claude skills hub](https://github.com/Vinciux99/artcraft-claude-skills): all 13 skills, a picker and a one-command installer.

## What it does
- Operating manual for DeckCraft: connection modes, ground rules, every MCP tool, recipes, CLI, control channel, command families and a safe workflow (work on copies, verify, save only with your OK).
- Built-in MCP setup assistant: Claude checks whether the server is connected, then walks you through installing it (automatic script or manual config), the launcher, authentication and verification.
- Troubleshooting table for the usual problems (server "disconnected", ports, macOS quarantine, Windows Store config location...).

## What you can ask Claude
- "Build a 6-slide deck on our Q3 results with the Harbor theme: title, three bullet slides, a chart of revenue by quarter, and a closing slide."
- "Open `pitch.pptx` as a copy, switch it to the Slate theme and fix any text that overflows."
- "Add a table of the three pricing tiers on slide 4 and style it."
- "Turn this outline into slides, one per heading."
- "Add speaker notes to every slide and export a handout PDF with 3 slides per page."
- "Add a Morph transition between slides 2 and 3 and fade the bullets in one by one."
- "Export all slides as PNG at 2×."
- "Run an accessibility and spelling check and tell me what to fix."

## Highlights
- Headless and attached (live) MCP, 25 tools, plus `call_app` for the whole control channel.
- Slide-point coordinate model, ids/indices conventions, themes, layouts and colors reference.
- Recipes for shapes, text, pictures, tables, charts, connectors, transitions/animations and export.
- CLI `run`/`render`/`convert` for fully headless pipelines; PPTX in and out.
- Built-in guidance for creating and troubleshooting the MCP server on Windows and macOS.

## App-specific notes
- Kit port **7990** (same as upstream docs). Always pass `--connect 7990`: the CLI's own default (7979) belongs to VectorCraft in this kit.
- The control channel has no authentication, loopback only; no read/write roots.
- Alpha software; `file.save` without a path overwrites the open file.

## Get the skill

| Where | How |
|---|---|
| **Claude app** (claude.ai / Desktop) | Download [`deckcraft-skill.zip](https://github.com/Vinciux99/deckcraft-claude-skill/releases/latest/download/deckcraft-skill.zip)`, then Customize > Skills > "+" > Create skill > Upload a skill. Needs Settings > Capabilities > "Code execution and file creation" on. |
| **Claude Code, plugin** | `/plugin marketplace add Vinciux99/artcraft-claude-skills` then `/plugin install deckcraft@artcraft` |
| **Claude Code, manual** | copy `skills/deckcraft/` to `~/.claude/skills/deckcraft/` (the setup script below does it) |

## Create the MCP server

Prerequisite: install DeckCraft from [upstream releases](https://github.com/storytold/deckcraft/releases) (Windows `.msi`, macOS `.dmg`) and Claude Desktop and/or Claude Code.

**One command (recommended)**

Windows (PowerShell):
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "& ([scriptblock]::Create((irm https://raw.githubusercontent.com/Vinciux99/artcraft-claude-skills/main/install.ps1))) -Apps deckcraft"
```
macOS (Terminal):
```sh
curl -fsSL https://raw.githubusercontent.com/Vinciux99/artcraft-claude-skills/main/install.sh | bash -s -- --apps deckcraft
```
Add `-DryRun` (Windows) or `--dry-run` (macOS) to preview, and `-Mode live|headless|both` / `--mode ...` to choose the mode.

**Or from a clone:** double-click `setup-windows.cmd` (Windows) or `setup-macos.command` (macOS; first time right-click > Open).

**Or just ask Claude:** once the skill is installed, say "set up the DeckCraft MCP server" and Claude will guide you step by step.

After setup: fully quit and reopen Claude, open DeckCraft with the **"DeckCraft (Claude)"** launcher, then ask Claude to use it. If the server shows "disconnected", the app was not running when Claude started: open it with the launcher and re-enable the server, or use headless mode.

What the setup changes (all backed up or reversible): Claude Desktop config (`mcpServers.deckcraft`, backup `<file>.backup-<timestamp>`), the launcher, Claude Code's user-scope MCP list and `~/.claude/skills/deckcraft/`.

## Repository layout
```
skills/deckcraft/SKILL.md      the skill
app.json                   machine-readable settings (ports, paths, MCP args) used by the scripts
scripts/                   artcraft-setup.ps1 (Windows), artcraft-setup.sh (macOS), json-tool.js
setup-windows.cmd          double-click wrapper
setup-macos.command        double-click wrapper
.claude-plugin/plugin.json Claude Code plugin manifest
.github/workflows/ci.yml   validation, setup tests on Windows and macOS, release zip
```

## CI and releases
Every push validates `app.json`, `SKILL.md` and the scripts, then runs the setup scripts against a fake install on **Windows** and **macOS** runners. Pushing a tag `v*` builds `deckcraft-skill.zip` and publishes a GitHub release.

## Status and caveats
- Windows paths and arguments come from a working install. The macOS paths (`/Applications/DeckCraft.app`, CLI location and release asset name) follow the upstream packaging but have **not** been verified on a Mac; the skill marks unverified items and the setup script reports what it could not find.
- DeckCraft is a young open-source project; behaviour can change between releases. `list_commands` / `deckcraft-cli commands` is always the source of truth.
- Unofficial community packaging, MIT licensed (see `LICENSE`).
