# Access DevCon Vienna 2026 — GBLauncher Demo

Companion repository for the [Access DevCon Vienna 2026](https://nolongerset.com/devcon-2026/) presentation on deploying and updating Microsoft Access applications using GBLauncher. No web server required.

> This repository does not accept pull requests. It is a distribution channel for GBLauncher demo files.

## Prerequisites

- **Microsoft Access** (any version — 32-bit or 64-bit)
- **Windows 10 or later**

## Features

- **SHA-256 integrity verification** — every download is hash-checked against the version sidecar before caching, protecting against corrupted or tampered files
- **Automatic Trusted Location registration** — the local cache folder is registered as an Access Trusted Location in the registry, so users are never prompted to enable content
- **No admin privileges required** — all data is stored under `%LOCALAPPDATA%\GBLauncher\` and only `HKCU` registry keys are touched; no `HKLM` writes, no elevation prompts
- **Auto-detects Microsoft Access** — finds the correct `msaccess.exe` via the registry, works with both full Access and Access Runtime, any version (32-bit or 64-bit)
- **Smart caching** — only downloads when a new version is available; cached versions launch instantly without network calls
- **RDS / Terminal Server compatible** — the front-end is cached in each user's local profile (`%LOCALAPPDATA%`), so multiple users on the same server each get their own copy
- **Single executable, no installer** — just drop `GBLauncher.exe` in a folder and run it; nothing to install, nothing to configure
- **Offline fallback** — if the network is unavailable, the launcher opens the last cached version instead of failing
- **Multiple deployment methods** — serve files from a local folder, a network share (UNC path), Dropbox, or any HTTP URL

## Quick Start (Dropbox)

1. **Download** this repository as a ZIP: click the green **Code** button above, then **Download ZIP**
2. **Unblock** the ZIP: right-click the downloaded file → **Properties** → check **Unblock** → **OK**
3. **Extract** the ZIP to any folder
4. **Double-click** `install-gblauncher-demo.bat`

The launcher will install itself, download the GBLauncher Demo app from Dropbox, create a desktop shortcut with a custom icon, and open the app in Microsoft Access. A form will display the current version number (e.g., **v1.0.0**).

5. Close Access and **double-click** `install-folder-watcher.bat`

This installs a second app, Folder Watcher, with its own desktop shortcut and icon. You now have two independent apps on your desktop.

### During the Presentation

1. Close Microsoft Access
2. Wait for the presenter to publish a new version of **GBLauncher Demo**
3. Click the **GBLauncher Demo** desktop shortcut
4. The launcher detects the new version, downloads it, and opens it — the form now shows the updated version
5. Click the **Folder Watcher** desktop shortcut — it still shows the original version, proving apps update independently

This demonstrates the full install and update cycle: the presenter publishes a new version to Dropbox with a single command, and every attendee gets the update automatically on their next launch — without touching Folder Watcher.

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
| `install-gblauncher-demo.bat` | Installs the GBLauncher Demo app — creates a desktop shortcut and launches it |
| `install-folder-watcher.bat` | Installs the Folder Watcher app — creates a desktop shortcut and launches it |
| `GBLauncher.exe` | The launcher client — installs, downloads, verifies, caches, and opens Access apps |
| `GBLauncher.twinproj` | Full twinBASIC source code for the launcher (open in [twinBASIC IDE](https://twinbasic.com/)) |
| `CHANGELOG.md` | Version history in [Keep a Changelog](https://keepachangelog.com/) format |
| `LICENSE` | MIT license |

The FolderWatcher twinBASIC source is available at [github.com/NoLongerSet/tb-folder-watcher](https://github.com/NoLongerSet/tb-folder-watcher).

Versioned downloads are also available on the [Releases](https://github.com/NoLongerSet/DevCon2026/releases) page.

## License

[MIT](LICENSE) — see the LICENSE file for details.
