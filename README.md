# ClaudeVS

**Use top AI coding CLIs directly inside Visual Studio.**

ClaudeVS integrates **Claude Code CLI, Copilot CLI, Codex CLI, and Gemini CLI** into Visual Studio 2022 and 2026, so you can work with your preferred agent without leaving the IDE.

## Why ClaudeVS

- **Multi-agent productivity with tabs**  
  Run multiple agents at the same time, each in its own tab, to compare outputs, split tasks, and switch contexts quickly.
- **Native CLI experience inside Visual Studio**  
  Agents run in integrated console windows, preserving their original behavior and controls.
- **Faster development loop**  
  Send editor context, task lines/comments, and debugger exceptions straight to an agent.
- **Hands-free input**  
  Press `Alt+S` to use Windows speech recognition for voice commands.
- **Theme-aware UI**  
  Supports **Light**, **Dark**, and **System** theme modes.
- **Privacy-first architecture**  
  The extension does **not** save credentials, conversation history, or other user data.

## Core Features

- Multiple concurrent agents via tabs
- Voice command input (`Alt+S`)
- Send active file path, line number, and selected text
- Execute current line/comment as a command
- Send debugger exception details (error, call stack, etc.)
- Support for custom programs in addition to supported CLIs
- Supports Light, Dark, and System themes

## Usage

| Menu Item | Description |
| --- | --- |
| **View → ClaudeVS** | Launch Claude Code CLI (requires an open project/solution) |
| **View → Send Location to Agent** | Send active file path, line number, and selected text |
| **View → Send Task to Agent** | Execute current line as a command |
| **View → Send Exception to Agent** | Send debugger exception/error details from debugger |

For voice input: press `Alt+S`, speak your command, then press Enter to send. Recording stops automatically when Windows speech recognition detects you are done speaking.

## Keyboard Shortcuts

All actions are hotkey-configurable:

- `ClaudeVS.SpeechCommand`
- `View.ClaudeVS`
- `View.SendLocationtoAgent`
- `View.SendTasktoAgent`
- `View.SendDebuggerExceptiontoAgent`

ClaudeVS can also capture and forward key combinations commonly used by Claude Code CLI (for example `Ctrl+B`, `Ctrl+O`, `Ctrl+R`). Configure these in:

**Tools → Options → Environment → Keyboard**

- `ClaudeVS.AgentAction1`
- `ClaudeVS.AgentAction2`
- `ClaudeVS.AgentAction3`
- `ClaudeVS.AgentAction4`

![claudevs_dark](https://github.com/user-attachments/assets/9fbe406b-b329-4589-9a8e-771e5e801789)

![claudevs_light](https://github.com/user-attachments/assets/ad602517-e77a-4c95-a198-52b6d3786282)
