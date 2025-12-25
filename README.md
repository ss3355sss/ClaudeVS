# ClaudeVS

[Free on the Visual Studio Marketplace](https://marketplace.visualstudio.com/items?itemName=GlassBeaver.ClaudeVS)

_NEW FEATURE:_ Certain hotkeys used by Claude Code CLI like Ctrl+B, Ctrl+O, Ctrl+R are now captured and forwarded.

Simple integration of Claude Code CLI, Copilot CLI, Codex CLI and Gemini CLI into Visual Studio 2026 and 2022. Works by launching an integrated console window with the desired agent (or any custom program) in it, so the agents/programs retain all of their native look & feel and functionality while VS is able to communicate with them.

Currently, three actions are implemented:
- Send active file path, line number and selected text to the agent
- Have the agent execute a comment in code, similar to how Copilot's tab completion works but simpler
- Send exception & error to the agent from the debugger (error message, callstack, etc.)

Nothing is saved by the extension: no credentials, conversations or anything since it just embeds the actual CLI programs inside Visual Studio.
The project was written entirely by AI so it's got lots of useless comments and debug logging.

**Usage**

| Menu Item  | Description |
| --- | --- |
| View -> ClaudeVS | Launch Claude Code CLI (requires an open project/solution) |
| View -> Send Location to Agent | Send file path and line number along with any text that's selected |
| View -> Send Task to Agent | Execute the current line as a command for code generation |
| View -> Send Exception to Agent | Send exception & error to the agent from the debugger |

All of these are hotkeyable:
- View.ClaudeVS
- View.SendLocationtoAgent
- View.SendTasktoAgent
- View.SendDebuggerExceptiontoAgent

Certain hotkeys used by Claude Code CLI like Ctrl+B, Ctrl+O, Ctrl+R are now captured and forwarded.
These can be configured in Options -> Environment -> Keyboard under the following commands:
- ClaudeVS.AgentAction1
- ClaudeVS.AgentAction2
- ClaudeVS.AgentAction3
- ClaudeVS.AgentAction4

![ClaudeVS](https://github.com/user-attachments/assets/b472dd7b-3c2a-45cb-8ab6-2026d1a5f0a0)
