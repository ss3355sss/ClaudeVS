# ClaudeVS

Free on the Visual Studio Marketplace: [https://marketplace.visualstudio.com/items?itemName=GlassBeaver.ClaudeVS]

Simple integration of Claude Code CLI, Copilot CLI, Codex CLI and Gemini CLI into Visual Studio 2026 and 2022. Works by launching an integrated console window with the desired agent (or any custom program) in it, so the agents/programs retain all of their native look & fell and functionality while VS is able to communicate with it.

Currently, two actions are implemented:
- Send active file path, line number and selected text to the agent
- Have the agent execute a comment in code, similar to how Copilot's tab completion works but simpler

Uses a custom-built integrated Windows Terminal because the one up on nuget wasn't taking the Esc key, which is a common and non-remappable keybind in many CLI agents.

Nothing is saved by the extension: no credentials, conversations or anything since it just embeds the actual CLI programs inside Visual Studio.
The project was written entirely by Sonnet 4.5 so it's got lots of useless comments and debug logging.

**Usage**

| Menu Item  | Description |
| --- | --- |
| View -> ClaudeVS | Launch Claude Code CLI (requires an open project/solution) |
| View -> Send Location to Agent | Send file path and line number along with any text that's selected |
| View -> Send Task to Agent | Execute the current line as a command for code generation |

All of these are hotkeyable:
- View.ClaudeVS
- View.SendLocationtoAgent
- View.SendTasktoAgent

![ClaudeVS](https://github.com/user-attachments/assets/b472dd7b-3c2a-45cb-8ab6-2026d1a5f0a0)
