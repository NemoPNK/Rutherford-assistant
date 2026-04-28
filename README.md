# Rutherford-assistant

Portable Windows launcher for running the Rutherford preparation scripts from a USB key. The launcher window stays open while scripts run, shows live logs, network status, and audits the modifications applied by `LaRoche.ps1`.

## Files

- `LaRocheLauncher.bat` starts the graphical launcher (fallback when the EXE is absent)
- `RutherfordAssistant.exe` is the preferred packaged launcher when built
- `assets\RutherfordLauncher.ps1` contains the WPF interface and the engine
- `assets\<Name>.ps1` is a script that the launcher can run
- `assets\<Name>.manifest.json` registers the script as a button in the launcher
- `assets\config\network-profiles.json` stores the expected network names and IP values
- `assets\checks\*.check.ps1` contains the LaRoche audit checks (one per modification)
- `assets\preinstall` contains the installation files used by setup

## Operator Flow

1. Plug in the USB key on the Windows PC.
2. Double-click `RutherfordAssistant.exe` if available, otherwise `LaRocheLauncher.bat`.
3. Accept the administrator prompt.
4. Click the action button you want (`Run Setup`, `Run Network`, etc.).
5. Keep the window open while the script runs.
6. Open the generated HTML report if needed.

## Behavior

- the launcher stays open until the operator closes it (asks for confirmation if a task is still running)
- logs are displayed live in the window
- one action runs at a time
- a colored HTML report is saved in the `reports` folder next to `LaRocheLauncher.bat` after each run
- the launcher shows live network cards with green / red status based on `Network.ps1` expectations, refreshed every 5 seconds
- after `LaRoche.ps1` finishes, the LaRoche Audit panel is automatically refreshed and shows green / orange / red badges per modification
- the per-script status (`Done` / `Error` / `Not done`) is persisted to `C:\ProgramData\Rutherford\launcher-state.json` so re-opening the launcher on the same PC restores the latest state

## Adding a new script (no EXE rebuild)

Drop two files in `assets\` next to the existing scripts:

- `MyScript.ps1`
- `MyScript.manifest.json`

Manifest format:

```json
{
  "key": "myscript",
  "label": "Run My Script",
  "description": "What this script does",
  "primary": false,
  "order": 30,
  "auditAfterRun": false
}
```

The launcher discovers manifests automatically at startup and renders one button per script, sorted by `order`.

## Adding or modifying an audit check (no EXE rebuild)

Drop a `*.check.ps1` file in `assets\checks\`. Each check returns a single hashtable:

```powershell
@{
    Id       = "my-check"
    Label    = "What is being verified"
    Category = "Cleanup"
    Order    = 100
    Test     = {
        # ... return one of:
        return @{ Status = "ok";      Detail = "All good" }
        return @{ Status = "missing"; Detail = "Not applied" }
        return @{ Status = "partial"; Detail = "Partially applied" }
        return @{ Status = "unknown"; Detail = "Cannot verify" }
    }
}
```

Errors thrown inside `Test` are caught and rendered as `unknown` instead of crashing the launcher.

## Build EXE

You can build the Windows EXE on GitHub Actions with the workflow in `.github/workflows/build-exe.yml`.
The workflow produces a portable zip containing:

- `RutherfordAssistant.exe`
- `LaRocheLauncher.bat`
- `assets\...` (scripts, manifests, checks, config, preinstall, wallpaper)

You can also run the build locally on Windows with:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-exe.ps1
```

Only `RutherfordLauncher.ps1` is compiled into the EXE. All other files (scripts, manifests, audit checks, network config, wallpaper, preinstall folder) are shipped as-is in `assets\` and can be modified on the USB key without rebuilding.
