# Appx Backup Manager

A Windows desktop tool for backing up installed Microsoft Store apps to a
portable folder, and restoring them later — on the same PC or a different
one. Wraps two PowerShell scripts (`Backup-AppxApps.ps1` /
`Restore-AppxApps.ps1`) with a GUI for selecting apps, tracking backup
status, and managing the resulting backup folder over time.

Built with the assistance of Claude (Anthropic).

## Purpose & legitimate use

This tool operates only on apps **already installed on your own machine**.
Backing up repackages an installed app's own files into a new `.appx` and
can sign it with a **self-signed certificate** so it can be sideloaded
again later — the same general mechanism Windows itself uses for
enterprise/developer sideloading, not a way to bypass any purchase, license
check, or DRM. It doesn't download, crack, or redistribute anyone else's app.

Restoring a backup onto a different machine may still be subject to that
app's own license terms — that's between the app's publisher and whoever
installs it, same as moving any of your own files between your own
machines. This is a personal backup/restore utility, not a piracy tool.

## Requirements

- Windows 10 or 11 (partial compatibility work was also done for
  Windows 8.1).
- Python 3.9+ (Tkinter included with a standard Windows install).
- PowerShell (ships with Windows).
- Administrator rights (needed for `Get-AppxPackage`/`Remove-AppxPackage`
  and sideloading registry changes).
- Optional: the Windows SDK (`MakeAppx.exe`, `SignTool.exe`) for producing
  portable, signed `.appx` files during backup. Without it, backup still
  works, just as a raw-folder-only backup restorable on the same PC. If you
  don't already have the full SDK, `MakeAppx.exe`/`SignTool.exe` can also be
  obtained by installing just the "Windows SDK Signing Tools for Desktop
  Apps" component via the standalone SDK installer, without the rest of
  Visual Studio.

## Running it

Keep `appx_backup_gui.py` in the same folder as `Backup-AppxApps.ps1` and
`Restore-AppxApps.ps1`, then:

```
python appx_backup_gui.py
```

## Backup types

- **Raw** — an exact copy of the installed app's folder. Restorable on the
  same PC (Register mode).
- **Packed** — a signed `.appx` built from Raw. Portable to other PCs
  (Install mode). In "Packed only" mode the temporary Raw folder is removed
  after packing, but a Raw backup you made earlier is always kept.
- **Both** — Raw and Packed.

Raw can be recreated later from Packed ("Create Raw Backups from Packed…")
and Packed from Raw ("Create Packed Backups from Raw…"), both under
**Special Actions** on the Restore tab. A failed re-pack never replaces an
existing working `.appx`.

Every backup is tracked in `backup-summary.csv` in the backup folder. Use
**Open Summary CSV** (next to the Backup / Restore tabs) to open a
timestamped copy of it — viewing or editing the copy never affects the real
file the scripts use.

## Restoring

- **Install** (recommended) — installs the packed `.appx` together with its
  dependencies and trusts the backup's certificate.
- **Register** — registers the Raw copy in place. Dependencies from the
  backup that aren't installed yet are registered first automatically.

Both scripts exit with code 1 when any package fails, so the GUI's
succeeded/failed summary is accurate; the Notes column of
`backup-summary.csv` and the Log show the reason (including the
MakeAppx/SignTool error text for packing and signing failures).

## Building a standalone .exe (optional)

```
pip install pyinstaller
pyinstaller --onefile --noconsole --name AppxBackupManager appx_backup_gui.py
```

Ships as a single `.exe` with both PowerShell scripts embedded inside it —
the two `.ps1` files don't need to travel alongside the built executable,
only alongside the raw `.py` source when running it directly. Runs without
Python installed, but only on machines matching the same architecture
(32-bit vs 64-bit) as whichever Python built it.

## Known limitations

A few are inherent to Windows itself rather than bugs in this tool:

- **Register-mode restore** is unreliable for Store-signed packages;
  Install mode is recommended instead.
- **Split / resource packages** (language/scale/architecture variants) are
  now included in the "Installed" check, but Windows often removes them
  together with their main app — so after uninstalling an app, its
  resource packages correctly show as not installed.
- **Same version already installed:** because backups are re-signed with
  your own certificate, installing one over the identical Store-installed
  version fails (`0x80073CFB`). Uninstall the existing copy first.
- **Unusual publisher names:** if the publisher name can't be used as-is
  for the certificate, signing fails and is reported as a failure (the
  package stays unsigned rather than being marked as signed).
- Some apps are simply incompatible with a different Windows version than
  they were built for (e.g. a manifest schema newer than the target OS
  understands) — restoring them there will fail regardless of the backup's
  own integrity.

See each tool's in-app Help (a button in the GUI itself) for further detail
on specific behaviors and options.

## License

GPL-3.0. See the `LICENSE` file (or add one from
[gnu.org](https://www.gnu.org/licenses/gpl-3.0.txt) if you haven't yet) —
note that GPL-3.0 is copyleft: anything you build on top of this and
distribute must also be released under GPL-3.0.
