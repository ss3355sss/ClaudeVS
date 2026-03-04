# Plan: Improve CLAUDE.md

## Context

The existing CLAUDE.md has good build instructions and terminal project details, but is missing code conventions from `.github/copilot-instructions.md` and lacks an architecture overview that would help future Claude Code instances navigate the codebase efficiently.

## Changes

### 1. Add required prefix
Replace the first line with the standard `/init` prefix:
```
# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.
```

### 2. Add missing code conventions (from copilot-instructions.md)
Append to the Code Conventions section:
- `Don't use braces where not necessary, e.g. single-line if statements, single-line loops.`
- `Use tabs, not spaces. Use CRLF, not LF.`

### 3. Expand Project Overview
Add two key architectural facts: the extension embeds Windows Terminal WPF, and all CLI communication is via terminal I/O (no API layer). Without these, an agent might try to create API calls.

### 4. Add Architecture section (between Project Overview and Build Command)
A concise section covering:
- The six command handlers and their async singleton pattern
- Component ownership chain (indented to show nesting):
  ```
  ClaudeVSPackage → ClaudeTerminalCommand → ClaudeTerminal (ToolWindowPane)
    → ClaudeTerminalControl (manages AgentTab instances)
      → AgentTab (per-tab: ConPtyTerminal + ConPtyTerminalConnection + TerminalControl)
  ```
- Key patterns: ThreadHelper guards, bracketed paste mode, SettingsManager

### 5. Preserve all existing content
Build Command, Terminal Project section, and terminal Architecture subsection remain unchanged.

### What was deliberately excluded
- File listings (discoverable via solution)
- Test instructions (no test framework exists)
- Git conventions from copilot-instructions.md (Copilot-specific; would conflict with Claude Code's own git workflow)
- `Build.bat` reference (uses VS 2022 path; CLAUDE.md build command correctly uses VS 2026/18)
- Generic development advice

## Files to modify
- `C:\ws\ClaudeVS\CLAUDE.md`

## Verification
- Read the updated file to confirm formatting and completeness
- Build the project to ensure no accidental changes to code files
