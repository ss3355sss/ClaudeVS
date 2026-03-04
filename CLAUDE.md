# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

ClaudeVS is a Visual Studio 2022/2026 extension (VSIX) that embeds AI coding CLIs (Claude Code, Copilot, Codex, Gemini) as native terminal windows inside the IDE. It uses Windows ConPTY for terminal emulation and Microsoft.Terminal.Wpf for rendering.

## Build

```
cmd /c "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe" ClaudeVS.csproj -t:Build -p:Configuration=Debug -v:minimal
```

Or just run `Build.bat`. **Build, don't Rebuild.** There are no automated tests.

To test: launch the VS Experimental Instance (F5 from VS with `/rootsuffix Exp`).

## Code Conventions

- Don't add new comments. Don't remove existing comments unless instructed.
- Don't add log statements unless explicitly told to.
- Don't use braces where not necessary (single-line `if`, single-line loops, etc.).
- Tabs for indentation, not spaces.
- CRLF line endings.
- Don't ever try to install or uninstall VS extensions.
- Git: read-only operations only (no stage, commit, push, pull).

## Architecture

**Entry point:** `ClaudeVSPackage.cs` — the VS package that registers all commands on async initialization.

**Terminal stack (bottom-up):**
1. `ConPtyTerminal.cs` — P/Invoke wrapper around Windows ConPTY API (CreatePseudoConsole, process pipes, resize). Discovers CLI paths and spawns agent processes.
2. `ConPtyTerminalConnection.cs` — implements `ITerminalConnection` to bridge ConPTY I/O with Microsoft.Terminal.Wpf's `TerminalControl`. Handles output buffering and pause/resume.
3. `ClaudeTerminalControl.xaml/.cs` — WPF user control with a tab bar; each tab owns one ConPtyTerminal + TerminalConnection pair. Manages theming (Light/Dark/System).
4. `ClaudeTerminal.cs` — VS tool window pane hosting the WPF control. Implements `IOleCommandTarget` to intercept VS key bindings and forward them to the terminal.

**Command handlers** (each registered in `ClaudeVSPackage.InitializeAsync`):
- `ClaudeTerminalCommand` — opens/focuses the tool window
- `SendFileLocationCommand` — sends active file path + line + selection to agent
- `SendCommentLineCommand` — sends the current comment line as a task instruction
- `SendDebuggerExceptionCommand` — captures debugger exceptions via `IDebugEventCallback2`
- `AgentActionCommand` — forwards hotkeys (Ctrl+B/O/R, Alt+1-4 QuickSwitch) to terminal
- `SpeechCommand` — Windows Speech Recognition voice input

**Settings:** `SettingsManager.cs` persists user preferences (font size, theme, last command, QuickSwitch presets) via VS `ShellSettingsManager`.

**External dependencies:**
- `Lib/Microsoft.Terminal.Control.{Debug,Release}.dll` — pre-built Microsoft.Terminal.Wpf binaries
- NuGet: Microsoft.VisualStudio.SDK, Microsoft.Windows.SDK.Contracts

## Key Patterns

- All terminal input uses **bracketed paste mode** (`\x1b[200~...\x1b[201~`) for sending multi-line content.
- Command definitions live in `VSCommandTable.vsct` (VSCT XML format defining menus, groups, buttons, and GUIDs/IDs).
- The extension targets .NET Framework 4.7.2.
