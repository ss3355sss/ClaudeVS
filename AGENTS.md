# AGENTS.md

This file provides guidance to Codex (Codex.ai/code) when working with code in this repository.

## Project

ClaudeVS is a Visual Studio 2022/2026 extension (VSIX) that embeds AI coding CLIs (Codex, Copilot, Codex, Gemini) as native terminal windows inside the IDE. It uses Windows ConPTY for terminal emulation and Microsoft.Terminal.Wpf for rendering. Targets .NET Framework 4.7.2.

## Build

```
cmd /c "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe" ClaudeVS.csproj -t:Build -p:Configuration=Debug -v:minimal
```

`Build.bat` runs **Release** config. **Build, don't Rebuild.** There are no automated tests.

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

**Entry point:** `ClaudeVSPackage.cs` — the VS package. `InitializeAsync` registers `SendFileLocationCommand` and `SendDebuggerExceptionCommand`. Other commands (`AgentActionCommand`, `SpeechCommand`, `SendCommentLineCommand`, `ClaudeTerminalCommand`) are initialized from `ClaudeTerminalControl` when the tool window loads.

**Terminal stack (bottom-up):**
1. `ConPtyTerminal.cs` — P/Invoke wrapper around Windows ConPTY API (CreatePseudoConsole, process pipes, resize). CLI path discovery: checks `%APPDATA%/npm/<cli>.cmd` first, then `PATH`. Spawns via `cmd.exe /c` for `.cmd`/`.bat` scripts, direct exec for `.exe`.
2. `ConPtyTerminalConnection.cs` — implements `ITerminalConnection` to bridge ConPTY I/O with Microsoft.Terminal.Wpf's `TerminalControl`. `WriteInput` auto-wraps multi-char non-escape input in bracketed paste (`\x1b[200~...\x1b[201~`). Supports output pause/resume with buffer.
3. `ClaudeTerminalControl.xaml/.cs` — WPF user control with a tab bar and toolbar. Each tab is an `AgentTab` struct holding its own ConPtyTerminal + Connection + TerminalControl. Manages theming (Light/Dark/System), font size, solution change detection, and multi-agent tabs.
4. `ClaudeTerminal.cs` — VS tool window pane hosting the WPF control. Implements `IOleCommandTarget` and `IVsWindowFrameNotify3`. `PreProcessMessage` forwards all key messages (except Escape) to the terminal instead of VS.

**Command handlers:**
- `SendFileLocationCommand` (0x0102) — copies file path + line + selection to clipboard, opens VS terminal
- `SendCommentLineCommand` (0x0103) — sends comment line as a task instruction via bracketed paste
- `SendDebuggerExceptionCommand` (0x0104) — captures exceptions via `IDebugEventCallback2`, copies rich debug context to clipboard
- `AgentActionCommand` (0x0105-0x0112) — forwards hotkeys to terminal, handles QuickSwitch (tab switching), NewAgent, and clipboard paste with bracketed paste mode
- `SpeechCommand` (0x010B) — Windows Speech Recognition voice input (unavailable in admin mode)

All commands share `CommandSet` GUID `a7c8e9d0-1234-5678-9abc-def012345678`. Package GUID is `b7d90b76-b34d-46e0-ab4f-888666287245`.

**Settings:** `SettingsManager.cs` persists user preferences (font size, theme, last command, QuickSwitch presets) via VS `ShellSettingsManager` under collection path `"ClaudeVS"`.

**External dependencies:**
- `Lib/Microsoft.Terminal.Control.{Debug,Release}.dll` — pre-built Microsoft.Terminal.Wpf binaries (native DLL, P/Invoked for selection/scroll)
- NuGet: Microsoft.VisualStudio.SDK, Microsoft.Windows.SDK.Contracts (for WinRT Speech API)

## Key Patterns

- **Bracketed paste mode** (`\x1b[200~...\x1b[201~`): `ConPtyTerminalConnection.WriteInput` applies this automatically for multi-char non-escape input. `AgentActionCommand` applies it explicitly for clipboard paste. `SendCommentLineCommand` wraps manually.
- **Send commands use clipboard**: `SendFileLocationCommand` and `SendDebuggerExceptionCommand` copy to clipboard and open the VS terminal — they do NOT write directly to the ConPTY terminal.
- **VSCT is minimal**: `VSCommandTable.vsct` declares only 2 buttons (SendFileLocation, SendDebuggerException) in the View menu. All other commands (AgentActions, Speech, QuickSwitch, NewAgent) are registered purely in code.
- **Solution awareness**: `ClaudeTerminalControl` listens to `SolutionEvents.Opened`/`AfterClosing`. On solution open, all tabs restart with the new solution directory. On close, all terminals stop.
- **Tab lifecycle**: Tabs are created via `CreateNewAgentTab` but initialized lazily (`EnsureTabInitialized` → `InitializeConPtyTerminal`) when first activated. Each tab tracks its own `CurrentSolutionPath`.
