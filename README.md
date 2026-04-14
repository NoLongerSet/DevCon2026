# Access DevCon Vienna 2026 — GBLauncher Demo

Companion repository for the [Access DevCon Vienna 2026](https://nolongerset.com/devcon-2026/) presentation on deploying and updating Microsoft Access applications using GBLauncher. No web server required.

> This repository does not accept pull requests. It is a distribution channel for GBLauncher demo files.

## Prerequisites

- **Microsoft Access** (any version — 32-bit or 64-bit)
- **Windows 10 or later**

## Quick Start (Dropbox)

1. **Download** this repository as a ZIP: click the green **Code** button above, then **Download ZIP**
2. **Unblock** the ZIP: right-click the downloaded file → **Properties** → check **Unblock** → **OK**
3. **Extract** the ZIP to any folder
4. **Double-click** `gblauncher-demo.bat`

The launcher will download the demo Access application from Dropbox, verify its integrity, cache it locally, and open it in Microsoft Access. A form will display the current version number (e.g., **v1.0.0**).

### During the Presentation

1. Close Microsoft Access
2. Wait for the presenter to publish a new version
3. Double-click `gblauncher-demo.bat` again
4. The launcher detects the new version, downloads it, and opens it — the form now shows the updated version

This demonstrates the full update cycle: the presenter publishes a new version to Dropbox with a single command, and every attendee gets the update automatically on their next launch.

## Local Network Deployment

GBLauncher also supports file-based deployment over local or network drives — no cloud service needed.

### Publishing a version

Publish an Access file to a shared folder:

```cmd
GBLauncher.exe --publish-local --app myapp --file "C:\path\to\MyApp.accdb" --version 1.0.0 --output "\\server\share\deploy"
```

This creates two files in the output folder:
- `MyApp.accdb` — a copy of the application
- `myapp.version` — a JSON sidecar with version, SHA-256 hash, and filename

### Launching from a network share

Point the client launcher at the sidecar file:

```cmd
GBLauncher.exe --app myapp --source "\\server\share\deploy\myapp.version"
```

The launcher reads the sidecar, copies the `.accdb`, verifies the SHA-256 hash, caches it locally to `%LOCALAPPDATA%\GBLauncher\cache\`, and opens it in Access. On subsequent runs, it only downloads if a new version is available.

## Troubleshooting

### "Open File - Security Warning" dialog

If Windows shows a security warning when you double-click the batch file, you extracted the ZIP without unblocking it first. To fix: right-click the original ZIP → **Properties** → check **Unblock** → **OK**, then re-extract. Alternatively, click **Run** to proceed.

### Windows SmartScreen warning

Because this executable is not code-signed, Windows may show a SmartScreen warning the first time you run it. Click **More info**, then **Run anyway**.

### Antivirus blocking the executable

Some antivirus software may flag unfamiliar executables. If `GBLauncher.exe` is quarantined, add an exception for the folder where you extracted the ZIP.

### "Microsoft Access is not installed"

The demo requires any edition of Microsoft Access. If you don't have Access installed, you can still follow along during the presentation.

## What's in This Repository

| File | Description |
|------|-------------|
| `gblauncher-demo.bat` | Batch file that launches the demo (pre-configured with the Dropbox URL) |
| `GBLauncher.exe` | The launcher client — downloads, verifies, caches, and opens Access apps |
| `GBLauncher.twinproj` | Full twinBASIC source code for the launcher (open in [twinBASIC IDE](https://twinbasic.com/)) |
| `templates/modSampleApp.bas` | VBA module used to create the demo Access application |
| `CHANGELOG.md` | Version history in [Keep a Changelog](https://keepachangelog.com/) format |
| `LICENSE` | MIT license |

Versioned downloads are also available on the [Releases](https://github.com/NoLongerSet/DevCon2026/releases) page.

## License

[MIT](LICENSE) — see the LICENSE file for details.
