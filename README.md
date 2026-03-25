# Claude Code Context Menu for Windows 11

Adds a **Claude Code** dropdown to Windows 11's modern Explorer right-click menu (the top-level menu, not "Show more options").

![Context Menu Preview](https://img.shields.io/badge/Windows_11-Context_Menu-0078D6?style=flat&logo=windows11)

<img width="519" height="480" alt="image" src="https://github.com/user-attachments/assets/e1abec12-4082-4bab-8406-3d40abf0fc95" />


## Menu Items

| Item | Command |
|------|---------|
| **Open Claude Code Here** | `claude` in the current folder |
| **YOLO Claude, YOLO!** | `claude --dangerously-skip-permissions` |

> **Warning:** The YOLO option launches Claude Code with all permission checks disabled. Use at your own risk.

## Prerequisites

- **Windows 11** (22H2 or later)
- **[Claude Code CLI](https://docs.anthropic.com/en/docs/claude-code)** installed and on PATH
- **Visual Studio** with the **C++ desktop development** workload (for compiling the native COM DLL)
- **Windows 10/11 SDK** (for MSIX packaging — typically installed with Visual Studio)

## Install

Run PowerShell **as Administrator**:

```powershell
.\add-claude-context-menu.ps1
```

The script will:
1. Auto-detect your `claude.exe` location (or pass `-ClaudePath "C:\path\to\claude.exe"`)
2. Compile the native COM DLL
3. Create a self-signed certificate and signed sparse MSIX package
4. Register the package and restart Explorer

## Uninstall

```powershell
.\add-claude-context-menu.ps1 -Uninstall
```

Removes the package, COM registration, certificates, and installed files.

## How It Works

A native C COM DLL implements [`IExplorerCommand`](https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/nn-shobjidl_core-iexplorercommand) with `ECF_HASSUBCOMMANDS`, returning sub-commands via `IEnumExplorerCommand`. It is registered through a [sparse AppX package](https://learn.microsoft.com/en-us/windows/apps/desktop/modernize/grant-identity-to-nonpackaged-apps) with `desktop5:FileExplorerContextMenus` — the only supported way to add items to Windows 11's modern context menu.

## Adding More Menu Items

Edit `ClaudeCodeContextMenu.c`:

1. Add a new `case` in `Enum_Next()` with a `CreateSubCommand()` call
2. Bump `NUM_SUBCMDS`
3. Rebuild and reinstall:

```powershell
.\add-claude-context-menu.ps1
```

## License

MIT
