# Powershell.SilentRemoveClickOnce

A PowerShell toolkit for silently uninstalling ClickOnce applications from the current user's profile — no GUI prompts, no `rundll32` required.

## Overview

ClickOnce applications store their installation state in per-user registry keys and a local file cache (`%LOCALAPPDATA%\Apps\2.0`). The standard uninstall path launches an interactive dialog. This toolkit bypasses that dialog by directly removing all ClickOnce artifacts:

| Step | What it does |
|------|--------------|
| 1 | Reads the app's uninstall information from `HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall` |
| 2 | Forcefully closes the running application process (if any) |
| 3 | Deletes the application files from `%LOCALAPPDATA%\Apps\2.0` |
| 4 | Removes Start Menu and Desktop shortcuts |
| 5 | Removes ClickOnce-specific registry keys under `HKCU\Software\Classes\…\Deployment\SideBySide\2.0` |
| 6 | Removes the uninstall registry entry |

## Files

| File | Purpose |
|------|---------|
| `main.ps1` | Entry point — accepts the application name and orchestrates the uninstall |
| `Models.ps1` | All supporting classes (`Uninstaller`, `UninstallInfo`, `ClickOnceRegistry`, etc.) |
| `Detection.ps1` | Intune/SCCM detection script that exits with code `1` (installed) or `0` (not installed) |

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- Must be run **as the target user** (all registry paths are under `HKCU`; no elevation required)
- The ClickOnce application must be registered in the current user's uninstall registry

## Usage

### Uninstall a ClickOnce application

```powershell
.\main.ps1 -AppName "Your Application Name"
```

`-AppName` must match the **Display Name** value stored in the registry (the same name that appears in *Add or Remove Programs*).

**Example:**

```powershell
.\main.ps1 -AppName "ATS Remote"
```

**Example output:**

```
Uninstalling ATS Remote
Components to remove:
    - <component key>
    ...
Uninstall complete
```

If the application cannot be found the script throws an exception and exits:

```
Exception: Could not find application: ATS Remote
```

### Detection script (Intune / SCCM)

`Detection.ps1` checks whether the hardcoded application (`ATS Remote`) is installed for the current user. It is intended to be used as a detection rule in Microsoft Intune or Configuration Manager.

```powershell
.\Detection.ps1
```

| Exit code | Meaning |
|-----------|---------|
| `0` | Application is **not** installed |
| `1` | Application **is** installed |

To adapt the detection script for a different application, update the `$appName` variable at the top of `Detection.ps1`.

## How it works

### `main.ps1`

1. Accepts the mandatory `-AppName` string parameter.
2. Dot-sources `Models.ps1` to load all classes.
3. Calls `[UninstallInfo]::Find($AppName)` to locate the app in the registry.
4. Creates an `[Uninstaller]` instance and calls `Uninstall($uninstallInfo)`.

### `Models.ps1`

The file is structured in four regions:

**ClickOnce Registry** (`ClickOnceRegistry`, `Component`, `Mark`, `Implication`, `RegistryMarker`)  
Reads the ClickOnce side-by-side component and marks registry trees so the uninstaller knows which keys to delete.

**Uninstall Info** (`UninstallInfo`)  
Reads `DisplayName`, `UninstallString`, and shortcut metadata from the standard uninstall registry hive.

**Uninstall Actions** — each class implements `Prepare()`, `PrintDebugInformation()`, and `Execute()`:

| Class | Responsibility |
|-------|---------------|
| `CloseOpenApplication` | Locates and force-stops the running process |
| `RemoveFiles` | Deletes component folders/files from `%LOCALAPPDATA%\Apps\2.0` |
| `RemoveStartMenuEntry` | Removes `.appref-ms` and `.url` shortcuts from Start Menu and Desktop |
| `RemoveRegistryKeys` | Deletes ClickOnce component, mark, package metadata, and state-manager registry keys |
| `RemoveUninstallEntry` | Removes the entry from the standard Windows uninstall registry |

**Uninstaller** (`Uninstaller`)  
Orchestrates all actions: reads the ClickOnce registry, resolves which components belong to the target app (via the public key token), then runs each action's `Prepare → PrintDebugInformation → Execute` pipeline.

## License

See [LICENSE](LICENSE).
