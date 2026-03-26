# Claude Code Context Menu for Windows 11

Adds a **Claude Code** dropdown to Windows 11's modern Explorer right-click menu (the top-level menu, not "Show more options").

![Windows 11 Context Menu](https://img.shields.io/badge/Windows_11-Context_Menu-0078D6?style=flat&logo=windows11)

<img width="429" height="339" alt="image" src="https://github.com/user-attachments/assets/f15d4b86-1ed9-4788-a076-98b45966bc6d" />


## Menu Items

| Item | Command | Risk Level |
|------|---------|------------|
| **Open (Default)** | `claude` in the current folder | Safe |
| **Open (Auto)** | `claude --enable-auto-mode` | Moderate |
| **Open (YOLO)** | `claude --dangerously-skip-permissions` | High |

> **Auto Mode** (new in March 2026) is the sweet spot -- an AI classifier reviews each action before it runs, auto-approving safe operations and blocking risky ones. Requires Claude Code on a Team plan with Sonnet 4.6 or Opus 4.6.

> **Warning:** The YOLO option launches Claude Code with all permission checks disabled. Use at your own risk.

## Prerequisites

- **Windows 11** (22H2 or later)
- **[Claude Code CLI](https://docs.anthropic.com/en/docs/claude-code)** installed and on PATH
- **Visual Studio** with the **C++ desktop development** workload (for compiling the native COM DLL)
- **Windows 10/11 SDK** (for MSIX packaging -- typically installed with Visual Studio)

## Install

Run PowerShell **as Administrator**:

```
.\add-claude-context-menu.ps1
```

The script will:

1. Auto-detect your `claude.exe` location (or pass `-ClaudePath "C:\path\to\claude.exe"`)
2. Compile the native COM DLL
3. Create a self-signed certificate and signed sparse MSIX package
4. Register the package and restart Explorer

## Uninstall

```
.\add-claude-context-menu.ps1 -Uninstall
```

Removes the package, COM registration, certificates, and installed files.

## How It Works

A native C COM DLL implements [`IExplorerCommand`](https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/nn-shobjidl_core-iexplorercommand) with `ECF_HASSUBCOMMANDS`, returning sub-commands via `IEnumExplorerCommand`. It is registered through a [sparse AppX package](https://learn.microsoft.com/en-us/windows/apps/desktop/modernize/grant-identity-to-nonpackaged-apps) with `desktop5:FileExplorerContextMenus` -- the only supported way to add items to Windows 11's modern context menu.

## Adding More Menu Items

Edit `ClaudeCodeContextMenu.c`:

1. Add a new `case` in `Enum_Next()` with a `CreateSubCommand()` call
2. Bump `NUM_SUBCMDS`
3. Rebuild and reinstall:

```
.\add-claude-context-menu.ps1
```

## License

MIT
