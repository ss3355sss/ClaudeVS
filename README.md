# ClaudeVS

[Free on the Visual Studio Marketplace](https://marketplace.visualstudio.com/items?itemName=GlassBeaver.ClaudeVS)

_NEW FEATURE:_ Voice input. Press Alt+S to activate speech recognition, speak into your mic and have Claude execute your command.

Straightforward integration of Claude Code CLI, Copilot CLI, Codex CLI and Gemini CLI into Visual Studio 2026 and 2022.
Works by launching an integrated console window with the desired agent (or any custom program), so the agents/programs retain their native look & feel and functionality while VS is able to communicate with them.

Available actions:
- Voice input: press Alt+S to activate speech recognition, speak into your mic and have Claude execute your command
- Send active file path, line number and selected text to agent
- Execute comment in code as a command
- Send exception/error details from debugger to agent (callstack, error message, etc.)

Nothing is saved by the extension: no credentials, conversations or anything else since it just embeds the actual CLI programs inside Visual Studio.

**Usage**

If you want to use your voice instead of typing, press Alt+S to record your command and hit Enter to send. The recording stops automatically when Windows speech recognition detects you've stopped speaking.

| Menu Item  | Description |
| --- | --- |
| View -> ClaudeVS | Launch Claude Code CLI (requires an open project/solution) |
| View -> Send Location to Agent | Send active file path, line number and selected text to agent |
| View -> Send Task to Agent | Execute current line as a command |
| View -> Send Exception to Agent | Send exception/error details to agent from debugger |

All of these are hotkeyable:
- ClaudeVS.SpeechCommand
- View.ClaudeVS
- View.SendLocationtoAgent
- View.SendTasktoAgent
- View.SendDebuggerExceptiontoAgent

Certain hotkeys used by Claude Code CLI like Ctrl+B, Ctrl+O, Ctrl+R are captured and forwarded.
These can be configured in Options -> Environment -> Keyboard under the following commands:
* ClaudeVS.AgentAction1
* ClaudeVS.AgentAction2
* ClaudeVS.AgentAction3
* ClaudeVS.AgentAction4

![claudevs_dark](https://github.com/user-attachments/assets/9fbe406b-b329-4589-9a8e-771e5e801789)

![claudevs_light](https://github.com/user-attachments/assets/ad602517-e77a-4c95-a198-52b6d3786282)
