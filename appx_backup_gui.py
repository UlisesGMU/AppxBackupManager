#!/usr/bin/env python3
"""
Appx Backup Manager
-------------------
A GUI wrapper around Backup-AppxApps.ps1 and Restore-AppxApps.ps1: lists
your installed Windows Store apps with a checkbox per app (and an estimated
size) for backup, and lets you pick exactly which apps from an existing
backup to restore - each with its own search box to filter a long list.

Windows-only. Must run as Administrator (same requirement as the two .ps1
scripts themselves - this GUI just launches them, inheriting elevation).

Usage:
    python appx_backup_gui.py
"""

import os
import sys
import csv
import json
import datetime
import base64
import glob
import queue
import ctypes
import shutil
import hashlib
import functools
import io
import re
import zipfile
import xml.etree.ElementTree as ET
import subprocess
import tempfile
import threading
import random
import time

if os.name != 'nt':
    print("This tool only works on Windows (it wraps Windows AppX management commands).")
    sys.exit(1)

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

SCRIPT_VERSION = "1.88.0"

# The set of per-line color tags _classify_log_line ever hands out - used
# by _copy_log_content_with_tags to know which tag (if any) to carry over
# when moving log content between the attached and detached Text widgets.
LOG_LINE_TAGS = ('ok_line', 'fail_line', 'warn_line', 'header_line', 'summary_header')

# The two PowerShell scripts are embedded below so this can be packaged into
# a single portable .exe (via PyInstaller) with nothing else needed
# alongside it. If a Backup-AppxApps.ps1 / Restore-AppxApps.ps1 file happens
# to sit next to this program (or the built .exe), that copy is used
# instead - handy if you want to tweak the script without rebuilding.

def _resolve_script(filename, embedded_text):
    if getattr(sys, 'frozen', False):
        here = os.path.dirname(sys.executable)
    else:
        here = os.path.dirname(os.path.abspath(__file__))
    companion = os.path.join(here, filename)
    if os.path.isfile(companion):
        return companion
    tmp_dir = os.path.join(tempfile.gettempdir(), 'appx-backup-manager')
    os.makedirs(tmp_dir, exist_ok=True)
    extracted = os.path.join(tmp_dir, filename)
    with open(extracted, 'w', encoding='utf-8') as f:
        f.write(embedded_text)
    return extracted


# Small embedded fallback so "Show useful tips" still works even if
# tips.json isn't sitting alongside the script/exe (e.g. it was built
# into a single .exe without also copying this file over) - load_tips()
# always prefers the actual file when it's there, since the whole point
# of a separate file is that it's yours to add more tips to over time.
DEFAULT_TIPS = {
    "backup": [
        "Dependencies and frameworks are backed up automatically whenever the main app is backed up - you don't need to select them yourself. This means the space needed will increase beyond just the app itself."
    ],
    "restore": [
        "Dependencies and frameworks are restored automatically whenever the main app is restored - you don't need to select them yourself. This means the space needed will increase beyond just the app itself."
    ],
}


def load_tips():
    """Reads tips.json from beside the script/exe (same folder
    _resolve_script looks in), falling back to DEFAULT_TIPS if it's
    missing or malformed - a tip failing to load should never be loud
    enough to interrupt anything else. Returns {'backup': [...], 'restore': [...]}."""
    if getattr(sys, 'frozen', False):
        here = os.path.dirname(sys.executable)
    else:
        here = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(here, 'tips.json')
    try:
        with open(path, encoding='utf-8') as f:
            data = json.load(f)
        if isinstance(data, dict):
            return {
                'backup': [t for t in data.get('backup', []) if isinstance(t, str)] or DEFAULT_TIPS['backup'],
                'restore': [t for t in data.get('restore', []) if isinstance(t, str)] or DEFAULT_TIPS['restore'],
            }
    except Exception:
        pass
    return DEFAULT_TIPS


BACKUP_PS1_TEXT = r"""<#
.SYNOPSIS
    Backs up installed Windows Store ("Modern"/UWP-style) apps so they can be
    reinstalled later - either on the SAME PC (e.g. after a Windows reset) or
    on ANOTHER PC (as a signed, portable .appx you can hand to Add-AppxPackage).

.DESCRIPTION
    For each selected installed package (plus everything it depends on, resolved
    automatically):
      1. Copies the raw install folder from C:\Program Files\WindowsApps\<...>
         using robocopy's backup-mode flag (/ZB), which uses the Windows backup
         privilege to read the files even though that folder is normally locked
         down to TrustedInstaller. No permissions on the original folder are
         changed - this is read-only and non-invasive.
      2. Repackages that folder into a clean .appx with MakeAppx.exe (part of the
         Windows SDK - the same tool family as makepri.exe).
      3. Self-signs each packed .appx (one certificate per distinct publisher,
         reused across all their apps) so it can be installed on another PC with
         sideloading enabled. Same-PC restores don't need this - see
         Restore-AppxApps.ps1's -Mode Register.
      4. Writes backup-summary.csv recording each package's dependencies, so the
         restore script can wire up -DependencyPath automatically.

.PARAMETER BackupRoot
    Folder to store everything in. Created if it doesn't exist.

.PARAMETER Include
    One or more wildcard patterns matched against each app's Name or
    PackageFullName. Default '*' (everything, minus Exclude, minus frameworks
    unless -IncludeFrameworks / they're a dependency of something selected).

.PARAMETER IncludeFile
    Path to a text file with one pattern/exact name per line, used instead
    of -Include. Preferred for long lists passed in from another program.

.PARAMETER ExcludeFile
    Same idea as -IncludeFile, but for -Exclude.

.PARAMETER Exclude
    Wildcard patterns to skip - handy for built-in Microsoft apps you don't
    care about (e.g. 'Microsoft.BingWeather', 'Microsoft.WindowsCalculator').

.PARAMETER AllUsers
    Also include packages registered for other user accounts / provisioned
    for the whole machine. Requires this script to be run elevated.

.PARAMETER IncludeFrameworks
    Also treat framework/runtime packages as directly selectable via
    -Include (they're always backed up anyway if something else depends on
    them - this only affects whether they show up as "selected" on their own).

.PARAMETER SkipPack
    Only do the raw folder backup - skip MakeAppx packing and signing
    entirely. Useful if you only care about same-PC restore and want this to
    run faster / not require the Windows SDK.

.PARAMETER SkipSign
    Pack into .appx but don't sign - the files won't be installable via
    Add-AppxPackage anywhere until you sign them yourself.

.PARAMETER DeleteRawAfterPack
    Once an app is successfully packed (and signed, unless -SkipSign),
    delete its Raw\<PackageFullName> folder, keeping only the packed
    .appx. Saves disk space for a "packed only" backup, but trades away
    two things that depend specifically on the Raw folder: Register-mode
    restore (which only ever reads Raw, never the packed .appx) won't work
    for this app afterward, and the file-level integrity checks (Verify
    Files / Verify All Backups) will have nothing left to compare against.
    Install-mode restore is unaffected, since it only ever reads the
    packed .appx. Ignored if -SkipPack is also set, since there'd be
    nothing packed to justify deleting Raw for.

.PARAMETER RepackFromRaw
    Instead of enumerating installed packages, reads backup-summary.csv
    for apps matching -Include/-IncludeFile that already have a Raw
    backup (RawBackupOK), and packs (and signs, unless -SkipSign)
    directly from that existing Raw folder - skipping the robocopy step
    entirely, since Raw is already there. Works even if the app is no
    longer installed. An app that already has a packed file is skipped
    the same way a normal backup run would skip anything already good -
    pass -Force too to overwrite it instead.

.PARAMETER CreateRawFromPacked
    The reverse of -RepackFromRaw: reads backup-summary.csv for apps
    matching -Include/-IncludeFile that already have a Packed backup
    but no Raw one, and adds Raw without touching the existing packed
    file (skips packing/signing entirely, whether or not -SkipPack/
    -SkipSign are also given). For each, if the app is currently
    installed, copies from its live install location exactly like a
    normal backup would (the more accurate source when it's
    available); if it isn't installed, extracts Raw directly from the
    packed .appx instead, since a .appx file is itself a zip archive -
    this never requires installing the app just to back it up. An app
    that already has a Raw backup is skipped the same way a normal
    backup run would skip anything already good - pass -Force too to
    overwrite it instead.

.PARAMETER MakeAppxPath
    Explicit path to makeappx.exe, if auto-detection (PATH, then the default
    Windows SDK install folder) doesn't find it.

.PARAMETER SignToolPath
    Explicit path to signtool.exe, if auto-detection doesn't find it.

.PARAMETER Force
    Normally, running this script again into a BackupRoot that already has a
    successfully-backed-up copy of a package (raw copy, and packing/signing
    if those weren't skipped) leaves that package alone rather than
    redoing the work - this is what makes it safe and fast to back up one
    app at a time into the same folder, since shared dependencies only get
    copied/packed/signed once. Pass -Force to redo everything regardless.

.EXAMPLE
    .\Backup-AppxApps.ps1 -BackupRoot D:\AppxBackup

.EXAMPLE
    .\Backup-AppxApps.ps1 -BackupRoot D:\AppxBackup -Include 'Contoso.*','Outfit7.*' -Exclude 'Microsoft.*'

.NOTES
    - Must be run as Administrator (needed for robocopy's backup-mode privilege,
      -AllUsers enumeration, and certificate installation to the machine store).
    - Requires the Windows SDK (MakeAppx.exe, SignTool.exe) for packing/signing -
      download at https://developer.microsoft.com/windows/downloads/windows-sdk
      if not already installed. Raw-folder-only backup (-SkipPack) does not
      need the SDK at all.
    - This only touches apps already installed on THIS machine, for your own
      personal backup/reinstall use.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$BackupRoot,

    [string[]]$Include = @('*'),
    [string[]]$Exclude = @(),
    [string]$IncludeFile,
    [string]$ExcludeFile,
    [switch]$AllUsers,
    [switch]$IncludeFrameworks,
    [switch]$SkipPack,
    [switch]$SkipSign,
    [switch]$DeleteRawAfterPack,
    [switch]$RepackFromRaw,
    [switch]$CreateRawFromPacked,
    [string]$MakeAppxPath,
    [string]$SignToolPath,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# -IncludeFile / -ExcludeFile: one pattern/name per line. Preferred over
# passing many values to -Include/-Exclude on the command line, which can
# behave inconsistently across PowerShell versions when a long list is
# involved - a file sidesteps that entirely.
if ($IncludeFile) {
    if (-not (Test-Path $IncludeFile)) {
        Write-Error "-IncludeFile path not found: $IncludeFile"
        exit 1
    }
    $Include = @(Get-Content -Path $IncludeFile | Where-Object { $_.Trim() -ne '' })
}
if ($ExcludeFile) {
    if (-not (Test-Path $ExcludeFile)) {
        Write-Error "-ExcludeFile path not found: $ExcludeFile"
        exit 1
    }
    $Exclude = @(Get-Content -Path $ExcludeFile | Where-Object { $_.Trim() -ne '' })
}

function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsAdmin)) {
    Write-Error "This script must be run as Administrator (needed for robocopy's backup privilege and, with -AllUsers, package enumeration). Right-click PowerShell -> Run as administrator, then re-run this script."
    exit 1
}

function Find-SdkTool {
    param([string]$Name)
    $onPath = Get-Command $Name -ErrorAction SilentlyContinue
    if ($onPath) { return $onPath.Source }
    $candidates = Get-ChildItem -Path 'C:\Program Files (x86)\Windows Kits\10\bin\*\x64' -Filter $Name -ErrorAction SilentlyContinue
    if (-not $candidates) {
        $candidates = Get-ChildItem -Path 'C:\Program Files (x86)\Windows Kits\10\bin\*\x86' -Filter $Name -ErrorAction SilentlyContinue
    }
    if ($candidates) { return ($candidates | Sort-Object FullName -Descending | Select-Object -First 1).FullName }
    return $null
}

New-Item -ItemType Directory -Path $BackupRoot -Force | Out-Null
$rawDir = Join-Path $BackupRoot 'Raw'
# /MIR makes Raw an exact copy (removes leftovers); /ZB uses backup mode for
# files the normal permissions would block.
$roboOptions = @('/MIR', '/ZB', '/R:2', '/W:2', '/NFL', '/NDL', '/NJH', '/NJS')
$packedDir = Join-Path $BackupRoot 'Packed'
$certsDir = Join-Path $BackupRoot 'Certs'
$manifestsDir = Join-Path $BackupRoot 'Manifests'
New-Item -ItemType Directory -Path $rawDir -Force | Out-Null
New-Item -ItemType Directory -Path $manifestsDir -Force | Out-Null

if ($CreateRawFromPacked) {
    # Loaded once, here, rather than inside the per-app loop below - a
    # .appx file is itself a zip archive, and this is what lets Raw be
    # extracted straight from one without installing the app at all.
    Add-Type -AssemblyName System.IO.Compression.FileSystem
}

$makeAppx = $null
$signTool = $null
# -CreateRawFromPacked never packs or signs, so it doesn't need either tool
# (and shouldn't warn that they're missing).
if (-not $SkipPack -and -not $CreateRawFromPacked) {
    if ($MakeAppxPath -and (Test-Path $MakeAppxPath)) {
        $makeAppx = $MakeAppxPath
    } else {
        $makeAppx = Find-SdkTool 'makeappx.exe'
    }
    if (-not $makeAppx) {
        Write-Warning "makeappx.exe not found - falling back to raw-folder-only backup (same-PC restore will still work; portable .appx creation will not). Pass -MakeAppxPath if you know where it is."
        $SkipPack = $true
    } else {
        New-Item -ItemType Directory -Path $packedDir -Force | Out-Null
        Write-Host "Using MakeAppx.exe: $makeAppx"
    }
}
if (-not $SkipPack -and -not $SkipSign -and -not $CreateRawFromPacked) {
    if ($SignToolPath -and (Test-Path $SignToolPath)) {
        $signTool = $SignToolPath
    } else {
        $signTool = Find-SdkTool 'signtool.exe'
    }
    if (-not $signTool) {
        Write-Warning "signtool.exe not found - packed .appx files will be left UNSIGNED and won't install anywhere until signed manually. Pass -SignToolPath if you know where it is."
        $SkipSign = $true
    } else {
        New-Item -ItemType Directory -Path $certsDir -Force | Out-Null
        Write-Host "Using SignTool.exe: $signTool"
    }
}

# ---- Enumerate and select packages ----

function Test-NameMatch {
    param($pkg, [string[]]$patterns)
    foreach ($p in $patterns) {
        if ($pkg.Name -like $p -or $pkg.PackageFullName -like $p) { return $true }
    }
    return $false
}

# Builds a package-like object from a backup-summary.csv row, in the same shape
# Get-AppxPackage returns, so the CSV-driven modes (-RepackFromRaw,
# -CreateRawFromPacked) can share the main loop with the normal mode.
function ConvertTo-PackageFromRow {
    param($row, $installLocation)
    $depNames = @()
    if ($row.Dependencies) { $depNames = $row.Dependencies -split ';' | Where-Object { $_ } }
    [PSCustomObject]@{
        Name            = $row.Name
        PackageFullName = $row.PackageFullName
        Version         = $row.Version
        Publisher       = $row.Publisher
        Architecture    = $row.Architecture
        IsFramework     = ($row.IsFramework -eq 'True')
        Dependencies    = $depNames | ForEach-Object { [PSCustomObject]@{ PackageFullName = $_ } }
        InstallLocation = $installLocation
    }
}

# Last few non-empty lines of a command-line tool's output, for the Notes
# column - so a MakeAppx/SignTool failure says why, not just an exit code.
function Get-ToolErrorText {
    param($output)
    return (@($output | ForEach-Object { "$_".Trim() } | Where-Object { $_ }) | Select-Object -Last 3) -join ' | '
}

if ($RepackFromRaw) {
    Write-Host "Reading existing backup-summary.csv to find apps to repack from Raw..."
    $summaryPathEarly = Join-Path $BackupRoot 'backup-summary.csv'
    if (-not (Test-Path $summaryPathEarly)) {
        Write-Error "No backup-summary.csv found in $BackupRoot - nothing to repack from."
        exit 1
    }
    $csvRowsEarly = Import-Csv -Path $summaryPathEarly
    $candidateRows = @($csvRowsEarly | Where-Object {
        $_.RawBackupOK -eq 'True' -and
        (Test-NameMatch $_ $Include) -and
        (-not (Test-NameMatch $_ $Exclude))
    })
    if (-not $candidateRows) {
        Write-Error "No rows in backup-summary.csv matched -Include $($Include -join ',') with an existing Raw backup. Nothing to repack."
        exit 1
    }
    Write-Host "Found $($candidateRows.Count) app(s) with an existing Raw backup to repack."
    $packages = $candidateRows | ForEach-Object { ConvertTo-PackageFromRow $_ $_.InstallLocation } | Sort-Object PackageFullName
} elseif ($CreateRawFromPacked) {
    Write-Host "Reading existing backup-summary.csv to find apps to add Raw for..."
    $summaryPathEarly = Join-Path $BackupRoot 'backup-summary.csv'
    if (-not (Test-Path $summaryPathEarly)) {
        Write-Error "No backup-summary.csv found in $BackupRoot - nothing to add Raw to."
        exit 1
    }
    $csvRowsEarly = Import-Csv -Path $summaryPathEarly
    $candidateRows = @($csvRowsEarly | Where-Object {
        $_.PackedOK -eq 'True' -and
        (Test-NameMatch $_ $Include) -and
        (-not (Test-NameMatch $_ $Exclude))
    })
    if (-not $candidateRows) {
        Write-Error "No rows in backup-summary.csv matched -Include $($Include -join ',') with an existing Packed backup. Nothing to add Raw to."
        exit 1
    }
    Write-Host "Found $($candidateRows.Count) app(s) with an existing Packed backup to add Raw for."
    # Checked once per app in the loop below, against whichever of these
    # actually correspond to something currently installed - not
    # assumed here, since the CSV alone can't say whether that's still
    # true right now.
    $liveInstalled = if ($AllUsers) { Get-AppxPackage -AllUsers } else { Get-AppxPackage }
    $liveInstalledByName = @{}
    foreach ($p in $liveInstalled) { $liveInstalledByName[$p.PackageFullName] = $p }
    # The live location if it's currently installed - a reinstalled or
    # updated app can live somewhere different from what the CSV recorded.
    $packages = $candidateRows | ForEach-Object {
        $liveLocation = if ($liveInstalledByName.ContainsKey($_.PackageFullName)) { $liveInstalledByName[$_.PackageFullName].InstallLocation } else { $null }
        ConvertTo-PackageFromRow $_ $liveLocation
    } | Sort-Object PackageFullName
} else {

Write-Host "Enumerating installed packages..."
$allPkgs = if ($AllUsers) { Get-AppxPackage -AllUsers } else { Get-AppxPackage }

$directlySelected = @($allPkgs | Where-Object {
    (-not $_.IsFramework -or $IncludeFrameworks) -and
    (Test-NameMatch $_ $Include) -and
    (-not (Test-NameMatch $_ $Exclude))
})

if (-not $directlySelected) {
    Write-Error "No installed packages matched -Include $($Include -join ',') after applying -Exclude. Nothing to back up. Run 'Get-AppxPackage | Select Name,PackageFullName' to see what's installed."
    exit 1
}

Write-Host "Directly selected: $($directlySelected.Count) package(s). Resolving dependencies..."
# Used below for AddedAsDependency - recorded once at backup time (was
# this package matched by -Include directly, or only pulled in via the
# dependency walk below?) rather than recomputed later, since "was this
# ever explicitly asked for" is a historical fact that shouldn't change
# just because something that used to depend on it got removed later.
$directlySelectedNames = @{}
foreach ($p in $directlySelected) { $directlySelectedNames[$p.PackageFullName] = $true }

# Walk dependencies transitively, deduping by PackageFullName.
# Skip anything ending in ".CBS" - these are Component-Based Servicing
# variants that Windows Update manages entirely on its own, completely
# independent of the Store app ecosystem. Windows can silently replace the
# exact installed version at any time (including mid-way through a long
# backup run), so their PackageFullName is inherently unstable to depend
# on - and there's no legitimate scenario needing to restore an old CBS
# component anyway, since Windows always maintains its own current one.
$worklist = @{}
function Add-ToWorklist {
    param($pkg)
    if ($worklist.ContainsKey($pkg.PackageFullName)) { return }
    if ($pkg.Name -like '*.CBS') {
        Write-Host "  Skipping $($pkg.PackageFullName) - a Component-Based Servicing package Windows manages on its own, not a real app dependency to preserve."
        return
    }
    $worklist[$pkg.PackageFullName] = $pkg
    foreach ($dep in $pkg.Dependencies) {
        Add-ToWorklist -pkg $dep
    }
}
foreach ($pkg in $directlySelected) { Add-ToWorklist -pkg $pkg }

$packages = $worklist.Values | Sort-Object PackageFullName
Write-Host "Total package(s) to back up, including dependencies: $($packages.Count)"

}

# ---- Load any existing summary, so repeated runs into the same folder ----
# ---- (e.g. one app at a time) MERGE instead of overwriting each other. ----

# Only the normal (enumerate-installed) mode fills this in; -RepackFromRaw and
# -CreateRawFromPacked still read it in the "already backed up" skip path, so it
# must exist as an empty table there instead of being $null (which throws).
if ($null -eq $directlySelectedNames) { $directlySelectedNames = @{} }

$summaryPath = Join-Path $BackupRoot 'backup-summary.csv'
$existingRows = @{}
if (Test-Path $summaryPath) {
    try {
        Import-Csv -Path $summaryPath | ForEach-Object { $existingRows[$_.PackageFullName] = $_ }
        Write-Host "Found an existing backup-summary.csv with $($existingRows.Count) package(s) - merging rather than overwriting."
    } catch {
        Write-Warning "Could not read the existing backup-summary.csv ($($_.Exception.Message)) - it will be replaced."
    }
}

function ConvertTo-NormalizedRow {
    # Guarantees every row has the full, current column set (older backups
    # made before a column existed - e.g. RawSizeBytes - would otherwise
    # break Export-Csv when merged with newer rows that do have it).
    param($row)
    $props = $row.PSObject.Properties.Name
    [PSCustomObject][ordered]@{
        Name             = $row.Name
        PackageFullName  = $row.PackageFullName
        Version          = $row.Version
        Publisher        = $row.Publisher
        Architecture     = $row.Architecture
        IsFramework      = $row.IsFramework
        AddedAsDependency = if ($props -contains 'AddedAsDependency') { $row.AddedAsDependency } else { '' }
        Dependencies     = $row.Dependencies
        InstallLocation  = $row.InstallLocation
        RawBackupOK      = $row.RawBackupOK
        RawDeletedAfterPack = if ($props -contains 'RawDeletedAfterPack') { $row.RawDeletedAfterPack } else { '' }
        RawSizeBytes     = if ($props -contains 'RawSizeBytes') { $row.RawSizeBytes } else { '' }
        PackedPath       = $row.PackedPath
        PackedOK         = $row.PackedOK
        SignedOK         = $row.SignedOK
        Notes            = $row.Notes
    }
}

# ---- Back up each package ----

$results = @()

# Adds a finished row to $results. If this run didn't end up with a working
# .appx for the app (it didn't pack at all - Raw-only mode, Create Raw from
# Packed - or the pack/Raw step failed), but an earlier run's .appx is still
# on disk, that earlier Packed backup is kept in the summary instead of being
# forgotten. Without this, a fresh row (PackedOK = False) replaced the old one
# and the CSV lost track of a perfectly good .appx.
function Add-ResultRow {
    param($row, $existing)
    if (-not $row.PackedOK -and $existing -and $existing.PackedOK -eq 'True' -and
        $existing.PackedPath -and (Test-Path -LiteralPath $existing.PackedPath)) {
        $row.PackedPath = $existing.PackedPath
        $row.PackedOK   = $true
        $row.SignedOK   = ($existing.SignedOK -eq 'True')
    }
    $script:results += [PSCustomObject]$row
}
$i = 0
foreach ($pkg in $packages) {
    $i++
    $existing = $existingRows[$pkg.PackageFullName]
    # Raw counts as "fine, nothing to do" either the normal way (it's
    # actually there), or when it was intentionally removed by an
    # earlier -DeleteRawAfterPack run AND this run is also
    # -DeleteRawAfterPack - i.e. the app is already sitting in exactly
    # the packed-only state this run would put it in anyway. Without
    # that second case, an app deliberately kept packed-only would look
    # "not fine" (RawBackupOK is correctly false for it) on every future
    # run and get needlessly robocopied/packed/signed/deleted again each
    # time, even though nothing about it actually needs to change. If
    # this run ISN'T -DeleteRawAfterPack, that second case doesn't apply
    # and Raw is correctly treated as missing - since the caller now
    # wants it restored, not left the way a past run intentionally left it.
    $rawIsFine = ($existing.RawBackupOK -eq 'True') -or
        ($DeleteRawAfterPack -and $existing.RawDeletedAfterPack -eq 'True' -and $existing.PackedOK -eq 'True')
    $alreadyGood = $existing -and $rawIsFine -and
        ($SkipPack -or $existing.PackedOK -eq 'True') -and
        ($SkipPack -or $SkipSign -or $existing.SignedOK -eq 'True')

    if ($alreadyGood -and -not $Force) {
        Write-Host "[$i/$($packages.Count)] $($pkg.PackageFullName) - already backed up, skipping (use -Force to redo)."
        $row = [ordered]@{
            Name              = $existing.Name
            PackageFullName   = $existing.PackageFullName
            Version           = $existing.Version
            Publisher         = $existing.Publisher
            Architecture      = $existing.Architecture
            IsFramework       = $existing.IsFramework
            # "Sticky" the same way DirectlySelected used to be: once
            # explicitly selected (either previously recorded as such, or
            # this run's own -Include matches it directly), that fact is
            # kept even though this app itself is being skipped as
            # already-good rather than reprocessed.
            AddedAsDependency = -not (($existing.AddedAsDependency -eq 'False') -or $directlySelectedNames.ContainsKey($pkg.PackageFullName))
            Dependencies      = $existing.Dependencies
            InstallLocation   = $existing.InstallLocation
            # Reflect whether Raw really exists now (it may have been removed
            # on purpose by -DeleteRawAfterPack) - never blindly $true, or the
            # cleanup step below tries to delete it again and the failure
            # check then counts it as a failure.
            RawBackupOK       = ($existing.RawBackupOK -eq 'True')
            # Internal only (not written to the CSV): see the -DeleteRawAfterPack step.
            KeepRaw           = ($existing.RawBackupOK -eq 'True')
            RawDeletedAfterPack = $existing.RawDeletedAfterPack
            RawSizeBytes      = $existing.RawSizeBytes
            PackedPath        = $existing.PackedPath
            PackedOK          = ($existing.PackedOK -eq 'True')
            SignedOK          = ($existing.SignedOK -eq 'True')
            Notes             = 'Already backed up in a previous run - skipped (use -Force to redo).'
        }
        Add-ResultRow $row $existing
        continue
    }

    Write-Host "[$i/$($packages.Count)] $($pkg.PackageFullName)"

    $row = [ordered]@{
        Name              = $pkg.Name
        PackageFullName   = $pkg.PackageFullName
        Version           = $pkg.Version
        Publisher         = $pkg.Publisher
        Architecture      = $pkg.Architecture
        IsFramework       = $pkg.IsFramework
        AddedAsDependency = if ($RepackFromRaw -or $CreateRawFromPacked) { $existing.AddedAsDependency -eq 'True' } else { -not $directlySelectedNames.ContainsKey($pkg.PackageFullName) }
        Dependencies      = if ($RepackFromRaw -or $CreateRawFromPacked) { $existing.Dependencies } else { ($pkg.Dependencies | ForEach-Object { $_.PackageFullName }) -join ';' }
        InstallLocation   = $pkg.InstallLocation
        RawBackupOK       = $false
        RawDeletedAfterPack = $false
        RawSizeBytes      = ''
        PackedPath        = ''
        PackedOK          = $false
        SignedOK          = $false
        Notes             = ''
        # Internal only (not written to the CSV): Raw already existed before
        # this run, so -DeleteRawAfterPack must leave it alone.
        KeepRaw           = [bool]($existing -and $existing.RawBackupOK -eq 'True')
    }

    $dest = Join-Path $rawDir $pkg.PackageFullName

    if ($RepackFromRaw) {
        # Raw is expected to already be there - that's the whole point of
        # this mode - so just confirm it actually still is (not deleted,
        # e.g. via a "Packed only" cleanup since the CSV row was written)
        # rather than assuming and only finding out when MakeAppx fails
        # with a much less clear error further down.
        if (-not (Test-Path $dest) -or -not (Get-ChildItem -LiteralPath $dest -File -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1)) {
            $row.Notes = "Raw folder missing or empty - nothing to repack from (was it removed since the backup summary was written?)."
            Add-ResultRow $row $existing
            Write-Warning "  $($row.Notes)"
            continue
        }
        $row.RawBackupOK = $true
        # Raw itself isn't being touched in this mode, so its previously
        # recorded size is still accurate - carried forward from $existing
        # rather than left at the fresh row's default empty value.
        $row.RawSizeBytes = $existing.RawSizeBytes
    } elseif ($CreateRawFromPacked) {
        if ($pkg.InstallLocation -and (Test-Path $pkg.InstallLocation)) {
            # Currently installed - copy from the live install, the same
            # way a normal backup would, since that's the more accurate
            # source when it's actually available.
            New-Item -ItemType Directory -Path $dest -Force | Out-Null
            $roboArgs = @($pkg.InstallLocation, $dest) + $roboOptions
            & robocopy @roboArgs | Out-Null
            $roboExit = $LASTEXITCODE
            if ($roboExit -ge 8) {
                $lockedHint = if ($roboExit -band 8) { " - likely one or more files were locked by another program; retrying later sometimes succeeds" } else { "" }
                $row.Notes = "robocopy failed (exit code $roboExit)$lockedHint."
                Write-Warning "  $($row.Notes)"
            } else {
                $row.RawBackupOK = $true
            }
        } else {
            # Not currently installed - extract straight from the packed
            # .appx instead, since a .appx file is itself a zip archive.
            # Never requires installing the app just to back it up.
            $packedPathForExtract = Join-Path $packedDir "$($pkg.PackageFullName).appx"
            if (-not (Test-Path $packedPathForExtract)) {
                $row.Notes = "Packed .appx missing - nothing to extract Raw from (was it removed since the backup summary was written?)."
                Add-ResultRow $row $existing
                Write-Warning "  $($row.Notes)"
                continue
            }
            try {
                # Start from an empty folder so nothing left over from an
                # earlier, partial attempt ends up mixed into this Raw copy.
                if (Test-Path -LiteralPath $dest) { Remove-Item -LiteralPath $dest -Recurse -Force }
                $destFull = [System.IO.Path]::GetFullPath($dest).TrimEnd('\') + '\'
                [System.IO.Directory]::CreateDirectory($destFull) | Out-Null
                # Extract entry by entry instead of ExtractToDirectory: .appx
                # stores file names URL-encoded (e.g. %20 for a space), so each
                # name is decoded back to what the installed app actually has.
                # [Content_Types].xml is packaging metadata that never exists in
                # an installed app's folder, so it's left out to match a live copy.
                $zip = [System.IO.Compression.ZipFile]::OpenRead($packedPathForExtract)
                try {
                    foreach ($entry in $zip.Entries) {
                        if (-not $entry.Name) { continue }  # directory entry
                        $relPath = [System.Uri]::UnescapeDataString($entry.FullName)
                        if ($relPath -eq '[Content_Types].xml') { continue }
                        $target = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($destFull, ($relPath -replace '/', '\')))
                        if (-not $target.StartsWith($destFull, [System.StringComparison]::OrdinalIgnoreCase)) {
                            throw "Refusing to extract '$relPath' - it would land outside the Raw folder."
                        }
                        [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($target)) | Out-Null
                        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $target, $true)
                    }
                } finally {
                    $zip.Dispose()
                }
                $row.RawBackupOK = $true
            } catch {
                $row.Notes = "Extracting Raw from the packed .appx failed: $($_.Exception.Message)"
                Write-Warning "  $($row.Notes)"
            }
        }
        # Packed already exists in this mode and isn't being touched -
        # carried forward from $existing rather than left at the fresh
        # row's default false/empty values.
        if ($row.RawBackupOK) {
            $row.PackedPath = $existing.PackedPath
            $row.PackedOK = ($existing.PackedOK -eq 'True')
            $row.SignedOK = ($existing.SignedOK -eq 'True')
        }
    } else {
        if (-not $pkg.InstallLocation -or -not (Test-Path $pkg.InstallLocation)) {
            $row.Notes = "InstallLocation missing or inaccessible - skipped."
            Add-ResultRow $row $existing
            Write-Warning "  $($row.Notes)"
            continue
        }

        New-Item -ItemType Directory -Path $dest -Force | Out-Null

        # /ZB: try normal copy, fall back to backup-mode (uses the backup privilege
        # to read past the restrictive ACL) if access is denied. /E copies subdirs
        # including empty ones. /R:2 /W:2 gives a locked file a couple of
        # short retries (a transient lock - antivirus scanning it,
        # OneDrive syncing it - can clear in a second or two) without
        # turning a genuinely, permanently locked file into a long hang.
        $roboArgs = @($pkg.InstallLocation, $dest) + $roboOptions
        & robocopy @roboArgs | Out-Null
        $roboExit = $LASTEXITCODE
        if ($roboExit -ge 8) {
            # Exit codes are a bitmask - 8 specifically means "some
            # file(s) couldn't be copied even after retrying" (usually a
            # lock: something else has the file open, or it's mid-scan/
            # mid-sync) - worth saying plainly, since "exit code 9" or
            # "exit code 8" alone doesn't tell you that's what happened
            # or that retrying later (once whatever's holding it closes)
            # is often all that's needed.
            $lockedHint = if ($roboExit -band 8) { " - likely one or more files were locked by another program; retrying later sometimes succeeds" } else { "" }
            $row.Notes = "robocopy failed (exit code $roboExit)$lockedHint."
            Write-Warning "  $($row.Notes)"
        } else {
            $row.RawBackupOK = $true
        }
    }

    if ($row.RawBackupOK -and -not $SkipPack -and -not $CreateRawFromPacked) {
        $packedPath = Join-Path $packedDir "$($pkg.PackageFullName).appx"
        $row.PackedPath = $packedPath
        try {
            # Pack to a temporary file and only replace the existing .appx once
            # the new one is known good - so a failed re-pack never destroys an
            # earlier working backup, and a stale file is never mistaken for a
            # new one.
            $tmpPackedPath = Join-Path $packedDir "$($pkg.PackageFullName).tmp.appx"
            if (Test-Path -LiteralPath $tmpPackedPath) { Remove-Item -LiteralPath $tmpPackedPath -Force }
            # 'Continue' just for this call: with 'Stop', Windows PowerShell turns
            # the tool's first stderr line into an exception, hiding the exit
            # code and the rest of its message.
            $prevEap = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
            $makeOut = & $makeAppx pack /d $dest /p $tmpPackedPath /o /nv /l 2>&1
            $makeExit = $LASTEXITCODE
            $ErrorActionPreference = $prevEap
            if ($makeExit -ne 0) {
                $row.Notes += " MakeAppx exited with code $($makeExit): $(Get-ToolErrorText $makeOut)"
                if (Test-Path -LiteralPath $tmpPackedPath) { Remove-Item -LiteralPath $tmpPackedPath -Force -ErrorAction SilentlyContinue }
            } elseif (Test-Path -LiteralPath $tmpPackedPath) {
                Move-Item -LiteralPath $tmpPackedPath -Destination $packedPath -Force
                $row.PackedOK = $true
            } else {
                $row.Notes += " MakeAppx did not produce an output file."
            }
        } catch {
            $row.Notes += " MakeAppx failed: $($_.Exception.Message)"
        }
    }

    if ($row.RawBackupOK) {
        # Recorded AFTER packing, not before - MakeAppx regenerates
        # AppxBlockMap.xml (and evidently touches other files too) directly
        # inside $dest as part of producing the packed .appx, so a manifest
        # taken before that step would show those exact files as "changed"
        # on every future verification even though nothing is actually
        # wrong - it'd just be comparing against a snapshot taken too
        # early. Recording it here instead captures the true final state
        # of the Raw folder once the whole backup process is done with it.
        try {
            $manifestEntries = Get-ChildItem -LiteralPath $dest -Recurse -File | ForEach-Object {
                $rel = $_.FullName.Substring($dest.Length + 1)
                [PSCustomObject]@{
                    rel  = $rel
                    hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
                    size = $_.Length
                }
            }
            # Recorded so the GUI can still show a real size for this app
            # even once Raw itself is gone (e.g. after a "Packed only"
            # backup, or a later "Raw Only" removal) - reusing the sizes
            # already computed above for the manifest, rather than
            # walking $dest a second time just for this.
            $row.RawSizeBytes = [int64](($manifestEntries | Measure-Object -Property size -Sum).Sum)
            $manifestPath = Join-Path $manifestsDir "$($pkg.PackageFullName).json"
            $manifestEntries | ConvertTo-Json -Depth 3 | Set-Content -Path $manifestPath -Encoding UTF8
        } catch {
            Write-Warning "  Could not write file manifest: $($_.Exception.Message)"
        }
    }

    Add-ResultRow $row $existing
}

# ---- Sign packed packages, one certificate per distinct publisher ----

if (-not $SkipSign -and $signTool) {
    $signable = @($results | Where-Object { $_.PackedOK -and -not $_.SignedOK })
    $publisherGroups = $signable | Group-Object Publisher
    if ($signable.Count -gt 0) {
        Write-Host "`nSigning $($signable.Count) packed package(s) ($(@($publisherGroups).Count) distinct publisher(s))..."
    }

    foreach ($group in $publisherGroups) {
        $publisher = $group.Name
        $safeName = ($publisher -replace '[^a-zA-Z0-9]', '_')
        $pfxPath = Join-Path $certsDir "$safeName.pfx"
        $cerPath = Join-Path $certsDir "$safeName.cer"
        $pwFilePath = Join-Path $certsDir "$safeName.password.txt"

        if (Test-Path $pfxPath) {
            # Reuse the SAME password an earlier run already created this
            # .pfx with - generating a fresh random one here (like earlier
            # versions of this script did) would sign against the wrong
            # password and fail every time the .pfx already exists.
            if (Test-Path $pwFilePath) {
                $pwText = (Get-Content -Path $pwFilePath -Raw).Trim()
            } else {
                Write-Warning "  Found $safeName.pfx but no matching $safeName.password.txt - can't sign with it. Delete Certs\$safeName.pfx (and .cer) to have this regenerate them, or restore the password file if you still have it."
                continue
            }
        } else {
            $pwText = [Guid]::NewGuid().ToString('N')
            try {
                # Built via certreq.exe + an INF file, and exported via raw
                # .NET (X509Certificate2.Export), instead of
                # New-SelfSignedCertificate / Export-PfxCertificate. Those
                # PowerShell PKI-module cmdlets vary a lot across Windows
                # versions - Windows 8.1's built-in copy lacks -Type,
                # -TextExtension, and possibly more. certreq.exe and .NET's
                # certificate classes have been stable and unchanged since
                # Windows XP, sidestepping this whole class of problem.
                $tmpDir = Join-Path $env:TEMP "appxcert_$([Guid]::NewGuid().ToString('N'))"
                New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null

                # $publisher is whatever Windows reports for this package -
                # an X.500 Distinguished Name string we didn't construct
                # ourselves. Some publishers' DNs contain characters (a
                # comma or "+" inside an Organization name, for instance)
                # that need RFC 1779 escaping certreq.exe doesn't tolerate
                # well even when present - rather than trying to correctly
                # re-escape an arbitrary DN we can't fully validate, try it
                # as-is first, and fall back to a guaranteed-safe
                # CN-only subject (built from the same sanitized string
                # already used for the file names) if that fails.
                function New-CertViaCertreq {
                    param([string]$SubjectValue, [string]$InfPath, [string]$OutCerPath)
                    $infContent = @"
[Version]
Signature="`$Windows NT`$"

[NewRequest]
Subject = "$SubjectValue"
KeySpec = 1
KeyLength = 2048
Exportable = TRUE
MachineKeySet = FALSE
SMIME = False
PrivateKeyArchive = FALSE
UserProtected = FALSE
UseExistingKeySet = FALSE
ProviderName = "Microsoft RSA SChannel Cryptographic Provider"
ProviderType = 12
RequestType = Cert
KeyUsage = 0xa0
HashAlgorithm = SHA256
ValidityPeriod = Years
ValidityPeriodUnits = 5

[EnhancedKeyUsageExtension]
OID=1.3.6.1.5.5.7.3.3
"@
                    Set-Content -Path $InfPath -Value $infContent -Encoding ASCII
                    & certreq.exe -new $InfPath $OutCerPath 2>&1 | Out-Null
                    return (Test-Path $OutCerPath)
                }

                $infPath = Join-Path $tmpDir 'request.inf'
                $subjectUsed = $publisher
                if (-not (New-CertViaCertreq -SubjectValue $publisher -InfPath $infPath -OutCerPath $cerPath)) {
                    Write-Warning "  certreq.exe rejected this publisher's certificate subject as-is (likely an unescaped special character in its name) - retrying with a simplified subject instead."
                    $subjectUsed = "CN=$safeName"
                    if (-not (New-CertViaCertreq -SubjectValue $subjectUsed -InfPath $infPath -OutCerPath $cerPath)) {
                        throw "certreq.exe did not produce a certificate file even with a simplified subject - see $infPath for the request that was used."
                    }
                }

                $newCert = Get-ChildItem 'Cert:\CurrentUser\My' |
                    Where-Object { $_.Subject -eq $subjectUsed } |
                    Sort-Object NotBefore -Descending |
                    Select-Object -First 1
                if (-not $newCert) {
                    throw "certreq.exe reported success but the new certificate wasn't found in Cert:\CurrentUser\My."
                }

                $pfxBytes = $newCert.Export([Security.Cryptography.X509Certificates.X509ContentType]::Pfx, $pwText)
                [System.IO.File]::WriteAllBytes($pfxPath, $pfxBytes)
                Set-Content -Path $pwFilePath -Value $pwText
                Remove-Item "Cert:\CurrentUser\My\$($newCert.Thumbprint)" -Force
                Remove-Item $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
            } catch {
                Write-Warning "  Could not create/export certificate for publisher '$publisher': $($_.Exception.Message)"
                continue
            }
        }

        foreach ($row in $group.Group) {
            try {
                $prevEap = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
                $signOut = & $signTool sign /fd SHA256 /f $pfxPath /p $pwText $row.PackedPath 2>&1
                $signExit = $LASTEXITCODE
                $ErrorActionPreference = $prevEap
                # Native tools don't throw on failure - check the exit code, or a
                # failed signature (e.g. cert subject not matching the manifest
                # Publisher) would be recorded as SignedOK.
                if ($signExit -ne 0) { throw "SignTool exited with code $($signExit): $(Get-ToolErrorText $signOut) (does the certificate subject exactly match the package Publisher?)" }
                $row.SignedOK = $true
            } catch {
                $row.Notes += " Signing failed: $($_.Exception.Message)"
            }
        }
    }
}

# ---- Delete Raw\ for anything successfully packed, if requested ----
# ---- (skipped entirely if -SkipPack, since nothing would be packed) ----

if ($DeleteRawAfterPack -and -not $SkipPack -and -not $CreateRawFromPacked) {
    # Only Raw folders this run created just to pack from are removed. A Raw
    # backup that already existed (made on purpose earlier) is kept, so
    # "Raw, then Packed" ends with both - the same as "Packed, then Raw".
    $toClean = @($results | Where-Object { $_.PackedOK -and $_.RawBackupOK -and -not $_.KeepRaw })
    if ($toClean.Count -gt 0) {
        Write-Host "`nRemoving Raw\ folder for $($toClean.Count) successfully packed app(s)..."
    }
    foreach ($row in $toClean) {
        $rawDirForRow = Join-Path $rawDir $row.PackageFullName
        try {
            if (Test-Path $rawDirForRow) {
                Remove-Item -LiteralPath $rawDirForRow -Recurse -Force -ErrorAction Stop
            }
            # RawBackupOK is set to false here too, not just
            # RawDeletedAfterPack - Raw genuinely doesn't exist anymore,
            # and RawBackupOK needs to reflect that current reality for
            # anything that reads it directly to decide whether Raw is
            # actually there right now (this script's own "already good,
            # skip" check on a later run, or the GUI's own candidate
            # search for "Create Raw Backups from Packed") - leaving it
            # at $true forever as a purely historical "did this ever
            # succeed" record, with RawDeletedAfterPack as the only place
            # the deletion was recorded, meant every one of those checks
            # would need to separately know to also check
            # RawDeletedAfterPack (or the disk itself) - and the ones
            # that didn't would wrongly treat Raw as still fine.
            # RawDeletedAfterPack remains a separate, permanent record of
            # why it's false - "genuinely never made" and "made, then
            # intentionally removed" still look different to anything
            # that cares (see the GUI's own display logic for exactly
            # that distinction).
            $row.RawBackupOK = $false
            $row | Add-Member -Force -NotePropertyName RawDeletedAfterPack -NotePropertyValue $true
        } catch {
            Write-Warning "  Could not remove Raw\$($row.PackageFullName): $($_.Exception.Message)"
        }
    }
}

# ---- Write summary: merge this run's results with any existing entries ----
# ---- that this run didn't touch, instead of overwriting the whole file. ----

$resultsByName = @{}
foreach ($r in $results) { $resultsByName[$r.PackageFullName] = ConvertTo-NormalizedRow $r }
foreach ($name in $existingRows.Keys) {
    if (-not $resultsByName.ContainsKey($name)) {
        $resultsByName[$name] = ConvertTo-NormalizedRow $existingRows[$name]
    }
}
$finalResults = $resultsByName.Values | Sort-Object PackageFullName
$finalResults | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8

$readmePath = Join-Path $BackupRoot 'README.txt'
@"
Appx Backup - $(Get-Date -Format 'yyyy-MM-dd HH:mm')

This run backed up $($results.Count) package(s). The backup folder now
contains $($finalResults.Count) package(s) in total (including anything
backed up in earlier runs into this same folder).

  Raw folder backup succeeded:  $(@($finalResults | Where-Object { "$($_.RawBackupOK)" -eq 'True' }).Count)
  Packed into .appx:            $(@($finalResults | Where-Object { "$($_.PackedOK)" -eq 'True' }).Count)
  Signed (portable-ready):      $(@($finalResults | Where-Object { "$($_.SignedOK)" -eq 'True' }).Count)

Folders:
  Raw\<PackageFullName>\      - exact copy of the installed app's files.
                                 Use Restore-AppxApps.ps1 -Mode Register on
                                 THIS SAME PC (e.g. after a Windows reset) to
                                 bring these back with no signing needed.
  Packed\<PackageFullName>.appx - repackaged, signed .appx files.
                                 Use Restore-AppxApps.ps1 -Mode Install on
                                 ANY PC (this one or another) - it will import
                                 the certificate(s) below automatically.
  Certs\*.cer / *.pfx          - one self-signed certificate per publisher
                                 found among your apps. The matching
                                 *.password.txt holds the .pfx password.
                                 Restore-AppxApps.ps1 imports the .cer
                                 automatically in -Mode Install; keep the .pfx
                                 only if you want to re-sign something later.
  Manifests\<PackageFullName>.json - hash + size of every file as it was at
                                 backup time, so you can later check whether
                                 anything in Raw\ was deleted or changed since.

See backup-summary.csv for full per-package details, including each
package's Dependencies (used automatically by Restore-AppxApps.ps1).
"@ | Set-Content -Path $readmePath -Encoding UTF8

Write-Host "`nDone. Summary: $summaryPath"
Write-Host "Read $readmePath for how to restore."
$failed = @($results | Where-Object {
    # Raw removed on purpose after a successful pack is not a failure.
    $rawGoneOnPurpose = ("$($_.RawDeletedAfterPack)" -eq 'True') -and $_.PackedOK
    (-not $_.RawBackupOK -and -not $rawGoneOnPurpose) -or
    (-not $SkipPack -and $_.RawBackupOK -and -not $_.PackedOK) -or
    (-not $SkipPack -and -not $SkipSign -and $_.PackedOK -and -not $_.SignedOK)
})
if ($failed) {
    Write-Warning "$($failed.Count) package(s) had a backup failure - see the Notes column in backup-summary.csv."
    # A non-zero exit code here is what lets the GUI (or any other
    # caller) actually notice something went wrong - without it, this
    # invocation still exits 0 like any other successful run, even
    # though one of the packages it covered never got backed up. The
    # per-app Write-Warning above is easy to miss in a scrolling log;
    # this is what a summary counting successes/failures by exit code
    # actually depends on.
    exit 1
}
"""

RESTORE_PS1_TEXT = r"""<#
.SYNOPSIS
    Restores apps backed up by Backup-AppxApps.ps1.

.PARAMETER BackupRoot
    The folder passed to Backup-AppxApps.ps1's -BackupRoot.

.PARAMETER Mode
    'Register' - SAME PC ONLY (e.g. after a Windows reset/reinstall). Copies
                 each app from Raw\ into a permanent local folder under
                 %ProgramData%\AppxRestoredApps\ first (registering directly
                 from BackupRoot is unreliable, especially from removable
                 media - Windows can reject it outright, and even when it
                 doesn't, -Register never copies files itself, so the app
                 keeps depending on that exact path forever). No certificate
                 trust needed either way, because these are the original,
                 already-trusted package files.
    'Install'  - ANY PC (this one or another). Imports the certificate(s) from
                 Certs\ into Trusted People, then installs each Packed\*.appx
                 with its dependencies wired up via -DependencyPath. Requires
                 sideloading / Developer Mode to be enabled on the target PC.

.PARAMETER Include
    Wildcard patterns matched against Name or PackageFullName (from
    backup-summary.csv). Default '*' (everything in the backup).

.PARAMETER IncludeFile
    Path to a text file with one pattern/exact name per line, used instead
    of -Include. Preferred for long lists passed in from another program.

.PARAMETER ExcludeFile
    Same idea as -IncludeFile, but for -Exclude.

.PARAMETER Exclude
    Wildcard patterns to skip.

.PARAMETER EnableSideloading
    (Install mode only) Sets the registry keys both Windows 10/11 ("Developer
    Mode") and Windows 8/8.1 ("Allow all trusted apps to install") use for
    sideloading, since they're different keys entirely - safe to set both
    regardless of which OS you're on. Modifies HKLM, requires elevation. On
    Windows 8/8.1 Pro editions that aren't Enterprise or domain-joined, this
    registry setting alone may not be enough - Microsoft also requires a
    separate sideloading product activation key via slmgr.exe for those.
    On Windows 10/11 you can instead enable this by hand: Settings > Update
    & Security > For developers > Sideload apps (or Developer Mode).

.EXAMPLE
    .\Restore-AppxApps.ps1 -BackupRoot D:\AppxBackup -Mode Register

.EXAMPLE
    .\Restore-AppxApps.ps1 -BackupRoot D:\AppxBackup -Mode Install -EnableSideloading
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$BackupRoot,

    [Parameter(Mandatory = $true)]
    [ValidateSet('Register', 'Install')]
    [string]$Mode,

    [string[]]$Include = @('*'),
    [string[]]$Exclude = @(),
    [string]$IncludeFile,
    [string]$ExcludeFile,
    [switch]$EnableSideloading
)

$ErrorActionPreference = 'Stop'

# -IncludeFile / -ExcludeFile: one pattern/name per line - preferred over
# passing many values to -Include/-Exclude on the command line for the same
# reason as Backup-AppxApps.ps1 (avoids any risk of inconsistent array
# binding across PowerShell versions with a long list of values).
if ($IncludeFile) {
    if (-not (Test-Path $IncludeFile)) {
        Write-Error "-IncludeFile path not found: $IncludeFile"
        exit 1
    }
    $Include = @(Get-Content -Path $IncludeFile | Where-Object { $_.Trim() -ne '' })
}
if ($ExcludeFile) {
    if (-not (Test-Path $ExcludeFile)) {
        Write-Error "-ExcludeFile path not found: $ExcludeFile"
        exit 1
    }
    $Exclude = @(Get-Content -Path $ExcludeFile | Where-Object { $_.Trim() -ne '' })
}

function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsAdmin)) {
    Write-Error "This script must be run as Administrator. Right-click PowerShell -> Run as administrator, then re-run this script."
    exit 1
}

$summaryPath = Join-Path $BackupRoot 'backup-summary.csv'
if (-not (Test-Path $summaryPath)) {
    Write-Error "backup-summary.csv not found under $BackupRoot - point -BackupRoot at a folder created by Backup-AppxApps.ps1."
    exit 1
}
$allRows = Import-Csv -Path $summaryPath

function Test-NameMatch {
    param($row, [string[]]$patterns)
    foreach ($p in $patterns) {
        if ($row.Name -like $p -or $row.PackageFullName -like $p) { return $true }
    }
    return $false
}
$selected = @($allRows | Where-Object { (Test-NameMatch $_ $Include) -and (-not (Test-NameMatch $_ $Exclude)) })
if (-not $selected) {
    Write-Error "No entries in backup-summary.csv matched -Include $($Include -join ',') after -Exclude."
    exit 1
}
Write-Host "Selected $($selected.Count) package(s) to restore."

# ---- Pre-flight: make sure every dependency a selected app needs is ----
# ---- actually present (and in good shape) somewhere in this backup.  ----

$allByName = @{}
foreach ($r in $allRows) { $allByName[$r.PackageFullName] = $r }

# Dependency names of a row, minus Component-Based Servicing (*.CBS) packages:
# Windows manages those itself and the backup deliberately never saves them,
# so listing them here only produced false "missing dependency" warnings.
function Get-DepNames {
    param($row)
    if (-not $row.Dependencies) { return @() }
    return @($row.Dependencies -split ';' | Where-Object { $_ -and (($_ -split '_')[0] -notlike '*.CBS') })
}

# Counted so the script can exit with code 1 when anything failed - the GUI's
# succeeded/failed summary depends on the exit code, and this script used to
# always exit 0.
$failures = 0

$missingDepWarnings = 0
foreach ($row in $selected) {
    if (-not $row.Dependencies) { continue }
    foreach ($depName in (Get-DepNames $row)) {
        if (-not $allByName.ContainsKey($depName)) {
            Write-Warning "$($row.PackageFullName) depends on $depName, which is not present anywhere in this backup at all."
            $missingDepWarnings++
            continue
        }
        $dep = $allByName[$depName]
        if ($Mode -eq 'Register' -and $dep.RawBackupOK -ne 'True') {
            Write-Warning "$($row.PackageFullName) depends on $depName, whose raw backup did not succeed - Register may fail for this app."
            $missingDepWarnings++
        }
        if ($Mode -eq 'Install' -and ($dep.PackedOK -ne 'True' -or $dep.SignedOK -ne 'True')) {
            Write-Warning "$($row.PackageFullName) depends on $depName, which was not successfully packed/signed - Install may fail for this app."
            $missingDepWarnings++
        }
    }
}
if ($missingDepWarnings -eq 0) {
    Write-Host "Dependency check passed - everything the selected app(s) need is present in this backup."
} else {
    Write-Host "Dependency check found $missingDepWarnings potential issue(s) above - continuing anyway, but keep an eye out for related failures below."
}

# ================= Mode: Register (same PC) =================

if ($Mode -eq 'Register') {
    $rawDir = Join-Path $BackupRoot 'Raw'
    # Unlike Install (which gets dependencies via -DependencyPath), Register
    # doesn't bring dependencies along - so every dependency of the selected
    # app(s) that's in this backup and not already installed is added too,
    # ordered so dependencies are registered before the apps that need them.
    $installedNow = @{}
    foreach ($p in (Get-AppxPackage -PackageTypeFilter Main,Framework,Resource,Optional)) { $installedNow[$p.PackageFullName] = $true }
    $ordered = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    function Add-WithDeps {
        param($r)
        if ($seen.ContainsKey($r.PackageFullName)) { return }
        $seen[$r.PackageFullName] = $true
        foreach ($d in (Get-DepNames $r)) {
            if ($allByName.ContainsKey($d) -and -not $installedNow.ContainsKey($d)) { Add-WithDeps $allByName[$d] }
        }
        $ordered.Add($r)
    }
    foreach ($r in $selected) { Add-WithDeps $r }
    $extraDeps = $ordered.Count - $selected.Count
    if ($extraDeps -gt 0) {
        Write-Host "Also registering $extraDeps dependency package(s) of the selected app(s) that aren't installed on this PC yet."
    }

    $toRegister = @($ordered | Where-Object { $_.RawBackupOK -eq 'True' })
    $skipped = $ordered.Count - $toRegister.Count
    if ($skipped -gt 0) {
        Write-Warning "$skipped package(s) have no Raw backup and will be skipped (if they were backed up as Packed only, use Install mode for them instead)."
        $failures += $skipped
    }

    # -Register never copies anything - it tells Windows "this exact folder,
    # forever, IS the installed app", and the app keeps running directly from
    # wherever you register it from. Registering straight out of BackupRoot
    # is fragile: if it's removable media (a USB drive), Windows commonly
    # REJECTS the registration outright ("manifest not located in the
    # package root" even though it plainly is); and even when it doesn't,
    # moving/renaming the backup folder later, or a USB drive getting a
    # different letter next time it's plugged in, breaks the "installed" app.
    # So: copy each app to a dedicated, permanent LOCAL folder first, and
    # register from THAT stable copy instead - decoupling the portable
    # backup archive from the live, registered app data.
    $localStoreDir = Join-Path $env:ProgramData 'AppxRestoredApps'
    New-Item -ItemType Directory -Path $localStoreDir -Force | Out-Null
    Write-Host "Registering from local copies under: $localStoreDir"

    # Plain arrays (not [List[object]]::new()) - the ::new() static
    # constructor syntax needs PowerShell 5.0+, but Windows 8.1 ships with
    # PowerShell 4.0 by default, where that syntax doesn't exist at all.
    $pending = @($toRegister)

    # Some apps may fail to register before their dependencies are registered.
    # Two passes: whatever still fails on pass 1 gets one more try on pass 2,
    # by which point earlier dependencies in the same run should be in place.
    for ($pass = 1; $pass -le 2; $pass++) {
        Write-Host "`n--- Register pass $pass ---"
        $stillPending = @()
        foreach ($row in $pending) {
            $sourceDir = Join-Path $rawDir $row.PackageFullName
            $localDir = Join-Path $localStoreDir $row.PackageFullName

            # Synced on the first pass every run (not only when the manifest is
            # missing), so a partial copy from an earlier run gets completed.
            if ($pass -eq 1) {
                Write-Host "  Copying $($row.PackageFullName) to a permanent local folder..."
                New-Item -ItemType Directory -Path $localDir -Force | Out-Null
                & robocopy $sourceDir $localDir /MIR /R:1 /W:1 /NFL /NDL /NJH /NJS | Out-Null
                if ($LASTEXITCODE -ge 8) {
                    Write-Warning "  FAILED: $($row.PackageFullName) - copying to the local folder failed (robocopy exit code $LASTEXITCODE; is the app running?)."
                    $failures++
                    continue
                }
            }

            $manifestPath = Join-Path $localDir 'AppxManifest.xml'
            if (-not (Test-Path $manifestPath)) {
                Write-Warning "  $($row.PackageFullName): AppxManifest.xml not found after copying locally, skipping."
                $failures++
                continue
            }
            try {
                Add-AppxPackage -Register $manifestPath -DisableDevelopmentMode -ErrorAction Stop
                Write-Host "  OK: $($row.PackageFullName)"
            } catch {
                if ($pass -eq 1) {
                    $stillPending += $row
                } else {
                    Write-Warning "  FAILED: $($row.PackageFullName) - $($_.Exception.Message)"
                    $failures++
                    if ($_.Exception.Message -match '0x80073CF9') {
                        Write-Warning "    'manifest not in package root' even from a local copy usually means the manifest itself is malformed or the folder is missing required files (AppxBlockMap.xml, AppxSignature.p7x) - check with 'Show Backed-Up Files' in the GUI."
                    }
                }
            }
        }
        $pending = $stillPending
        if ($pending.Count -eq 0) { break }
    }
    Write-Host "`nDone."
    if ($failures -gt 0) {
        Write-Warning "$failures package(s) could not be restored - see the messages above."
        exit 1
    }
    exit 0
}

# ================= Mode: Install (this PC or another) =================

if ($EnableSideloading) {
    Write-Host "Enabling sideloading via registry..."

    # Each key is wrapped separately - one failing (e.g. a security policy
    # blocking writes under SOFTWARE\Policies on this particular system)
    # must not stop the whole script before it even gets to installing
    # anything. Both are best-effort; the actual install below is the real
    # test of whether sideloading trust is sufficient.

    # Windows 10/11 "Developer Mode" mechanism.
    try {
        $win10Key = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock'
        New-Item -Path $win10Key -Force -ErrorAction Stop | Out-Null
        Set-ItemProperty -Path $win10Key -Name 'AllowAllTrustedApps' -Value 1 -Type DWord -ErrorAction Stop
        Set-ItemProperty -Path $win10Key -Name 'AllowDevelopmentWithoutDevLicense' -Value 1 -Type DWord -ErrorAction Stop
        Write-Host "  Set Windows 10/11 Developer Mode sideloading keys."
    } catch {
        Write-Warning "  Could not set the Windows 10/11 sideloading registry key: $($_.Exception.Message)"
    }

    # Windows 8/8.1 mechanism - a DIFFERENT registry location entirely (the
    # "Allow all trusted apps to install" Group Policy setting writes here).
    # Not needed at all on Windows 10/11, and some systems block direct
    # writes under SOFTWARE\Policies even for an elevated Administrator -
    # a failure here is expected on many machines and safe to ignore.
    try {
        $win81Key = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Appx'
        New-Item -Path $win81Key -Force -ErrorAction Stop | Out-Null
        Set-ItemProperty -Path $win81Key -Name 'AllowAllTrustedApps' -Value 1 -Type DWord -ErrorAction Stop
        Write-Host "  Set Windows 8/8.1 'Allow all trusted apps' registry key."
        Write-Host "  Note for Windows 8/8.1 Pro (non-Enterprise, non-domain-joined) editions: this registry setting alone may not be enough - Microsoft also requires a separate sideloading product activation key entered via slmgr.exe for those specific editions. Windows 8.1 Enterprise, or a domain-joined PC, does not need that extra key."
    } catch {
        Write-Warning "  Could not set the Windows 8/8.1 sideloading registry key (expected/harmless on Windows 10/11): $($_.Exception.Message)"
    }
}

$certsDir = Join-Path $BackupRoot 'Certs'
if (Test-Path $certsDir) {
    # Only import certs for the PUBLISHERS actually being restored - both
    # the selected apps themselves AND their recorded dependencies (a
    # third-party framework dependency needs its own cert trusted too;
    # Microsoft's own frameworks are already trusted by Windows regardless,
    # but checking dependencies' publishers as well costs nothing and
    # covers that case correctly) - not every cert this backup folder has
    # ever accumulated across every app ever backed up into it, which is
    # confusing and unnecessary when restoring just one or two apps.
    $dependencyFullNames = @()
    foreach ($row in $selected) {
        if ($row.Dependencies) {
            $dependencyFullNames += ($row.Dependencies -split ';' | Where-Object { $_ })
        }
    }
    $dependencyRows = @($allRows | Where-Object { $dependencyFullNames -contains $_.PackageFullName })
    $neededSafeNames = @(($selected + $dependencyRows) | ForEach-Object { ($_.Publisher -replace '[^a-zA-Z0-9]', '_') } | Select-Object -Unique)
    $cerFiles = @(Get-ChildItem -Path $certsDir -Filter '*.cer' | Where-Object {
        $neededSafeNames -contains [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
    })
    if ($cerFiles.Count -eq 0) {
        Write-Warning "No matching certificate found in Certs\ for the selected app(s)' publisher(s) - they may have been left unsigned at backup time, or the Certs folder is from a different backup."
    } else {
        Write-Host "`nImporting $($cerFiles.Count) certificate(s) for the selected app(s)' publisher(s) into Trusted People..."
        foreach ($cer in $cerFiles) {
            try {
                Import-Certificate -FilePath $cer.FullName -CertStoreLocation 'Cert:\LocalMachine\TrustedPeople' | Out-Null
                Write-Host "  Imported: $($cer.Name)"
            } catch {
                Write-Warning "  Failed to import $($cer.Name): $($_.Exception.Message)"
            }
        }
    }
} else {
    Write-Warning "No Certs folder found - packages may have been left unsigned at backup time. Add-AppxPackage will fail unless they're already trusted some other way."
}

$packedDir = Join-Path $BackupRoot 'Packed'

function Get-PackageFamilyKey {
    # Name and PublisherId are the first and last underscore-separated
    # segments of a PackageFullName regardless of version/architecture/split
    # qualifiers in between - the same grouping Windows itself uses to know
    # which resource/split packages belong to which main app.
    param([string]$FullName)
    $parts = $FullName -split '_'
    if ($parts.Count -lt 2) { return $FullName }
    return "$($parts[0])_$($parts[-1])"
}

$packedSelected = @($selected | Where-Object { $_.PackedOK -eq 'True' })

# Resource/split packages (language, scale, architecture variants) can NEVER
# be installed on their own - Windows requires them to be installed TOGETHER
# with their main app in the same operation, not as an independent app. Pull
# them out of the main install loop and instead attach each one to its main
# app's -DependencyPath below (grouped by package family).
$splitPackages = @($packedSelected | Where-Object { $_.PackageFullName -match '_split\.' })
$splitByFamily = @{}
foreach ($sp in $splitPackages) {
    $fam = Get-PackageFamilyKey $sp.PackageFullName
    if (-not $splitByFamily.ContainsKey($fam)) { $splitByFamily[$fam] = @() }
    $splitByFamily[$fam] += $sp
}

$apps = @($packedSelected | Where-Object { $_.IsFramework -eq 'False' -and $_.PackageFullName -notmatch '_split\.' })
$skippedFrameworks = @($selected | Where-Object { $_.IsFramework -eq 'True' }).Count
if ($skippedFrameworks -gt 0) {
    Write-Host "$skippedFrameworks framework package(s) in the selection will be installed automatically as dependencies where needed, not installed directly."
}
if ($splitPackages.Count -gt 0) {
    Write-Host "$($splitPackages.Count) resource/split package(s) (language, scale, or architecture variants) will be installed together with their main app, not as their own separate install."
}

$notPacked = @($selected | Where-Object { $_.PackedOK -ne 'True' -and $_.IsFramework -ne 'True' })
if ($notPacked.Count -gt 0) {
    Write-Warning "$($notPacked.Count) selected package(s) have no Packed backup and can't be installed (use Register mode for them if they have a Raw backup):"
    $notPacked | ForEach-Object { Write-Warning "  $($_.PackageFullName)" }
    $failures += $notPacked.Count
}

Write-Host "`nInstalling $($apps.Count) app(s)..."
foreach ($row in $apps) {
    $mainPath = Join-Path $packedDir "$($row.PackageFullName).appx"
    if (-not (Test-Path $mainPath)) {
        Write-Warning "  $($row.PackageFullName): packed .appx not found, skipping."
        $failures++
        continue
    }

    $depPaths = @()
    if ($row.Dependencies) {
        foreach ($depFullName in (Get-DepNames $row)) {
            $depPacked = Join-Path $packedDir "$depFullName.appx"
            if (Test-Path $depPacked) {
                $depPaths += $depPacked
            } else {
                Write-Warning "    Dependency $depFullName has no packed .appx - install may fail if it's missing on this PC."
            }
        }
    }

    $fam = Get-PackageFamilyKey $row.PackageFullName
    if ($splitByFamily.ContainsKey($fam)) {
        foreach ($sp in $splitByFamily[$fam]) {
            $spPacked = Join-Path $packedDir "$($sp.PackageFullName).appx"
            if (Test-Path $spPacked) {
                $depPaths += $spPacked
            } else {
                Write-Warning "    Resource package $($sp.PackageFullName) has no packed .appx - skipping it for this install."
            }
        }
    }

    # A resource/split package can end up in BOTH the recorded Dependencies
    # list AND the family-grouping above (Windows' own package graph
    # sometimes already lists resource packages as formal dependencies) -
    # Add-AppxPackage rejects a -DependencyPath with the same package
    # listed twice ("Each specified package must be unique"), so dedupe.
    $depPaths = @($depPaths | Select-Object -Unique)

    try {
        if ($depPaths.Count -gt 0) {
            Add-AppxPackage -Path $mainPath -DependencyPath $depPaths -ForceApplicationShutdown -ErrorAction Stop
        } else {
            Add-AppxPackage -Path $mainPath -ForceApplicationShutdown -ErrorAction Stop
        }
        Write-Host "  OK: $($row.PackageFullName)"
    } catch {
        Write-Warning "  FAILED: $($row.PackageFullName) - $($_.Exception.Message)"
        $failures++
        if ($_.Exception.Message -match '0x80073CFB') {
            Write-Warning "    The same version is already installed with different contents (the backup was re-signed with your own certificate). Uninstall the existing copy first, then restore."
        }
        if ($_.Exception.Message -match '0x800B0100|0x80073CF9|0x80073CFF|trust') {
            Write-Warning "    This usually means sideloading isn't enabled, or the certificate wasn't trusted. Re-run with -EnableSideloading."
        }
        if ($_.Exception.Message -match '0x80080204') {
            Write-Warning "    This means the app's manifest uses a schema newer than this Windows version understands - a genuine OS compatibility limit, not something a re-try or re-signing can fix."
        }
    }
}

Write-Host "`nDone."
if ($failures -gt 0) {
    Write-Warning "$failures package(s) could not be restored - see the messages above."
    exit 1
}
"""

BACKUP_PS1 = _resolve_script('Backup-AppxApps.ps1', BACKUP_PS1_TEXT)
RESTORE_PS1 = _resolve_script('Restore-AppxApps.ps1', RESTORE_PS1_TEXT)

# Avoids a black console window flashing up for every PowerShell call once
# this is packaged as a windowed (--noconsole) .exe.
CREATE_NO_WINDOW = getattr(subprocess, 'CREATE_NO_WINDOW', 0)

HELP_TEXT = """APPX BACKUP MANAGER - WHAT EVERYTHING DOES

COPYING DATA (works everywhere - the Backup list, the Restore list, and the
per-app file verifier)
  Right-click any row for "Copy Cell" (just the column you clicked), "Copy
  Row" (every column for that row), or "Copy All Rows"/"Copy All Visible
  Rows" (everything currently shown, tab-separated - pastes cleanly into a
  spreadsheet or a chat message). The Log pane has its own right-click ->
  "Copy All Log Text", though its text also supports normal click-drag
  selection and Ctrl+C like any text box.

SHARED "Backup / restore folder" (top of window)
  Both tabs use this same folder. For Backup, it's where a new backup gets
  created - running a backup into a folder that already has one MERGES with
  it (adds new apps, skips ones already successfully backed up) rather than
  replacing it, so it's safe to back up one app at a time into the same
  folder over multiple sessions. For Restore, it's where an existing backup
  (made by this tool) lives - it needs a backup-summary.csv in it, which
  the Backup tab creates/updates automatically. This is a dropdown of your
  5 most recently used folders - pick one instead of browsing again, or
  type/browse a new one as usual. Picking a recent folder that already has
  a backup in it also auto-loads the Restore tab's contents - and the most
  recently used folder is pre-filled automatically the next time you open
  this app too, with the same auto-load behavior, so there's no need to
  re-pick it every session.

BACKUP TAB
  "Show:" (All / Apps / Frameworks)
      Switches which kind of installed package the list below shows -
      instant, since everything installed is only ever fetched once (on
      "Refresh List" or when this tab first loads) and this just filters
      what's already there rather than re-scanning. Whatever's checked
      elsewhere stays checked when you switch views.
        All - Apps and Frameworks together (see both below). Not every
          installed package - a non-framework package signed with the
          same certificate as Windows itself (deeply OS-integrated
          components, not really "apps" in the usual sense) is excluded
          here the same as under either of the other two views.
        Apps - real, launchable apps: not a framework, not OS-signed, and
          (optionally) Start Menu-visible. What you want almost all the
          time.
        Frameworks - VCLibs, UI.Xaml, .NET Native, and similar. Normally
          never independently selected - Backup-AppxApps.ps1 already
          pulls in whatever a real app needs, via a fresh live lookup, not
          a name captured here which could go stale in a long batch -
          this exists so you can still find and pick one directly
          yourself if you specifically want to.

  "Include all users (-AllUsers)"
      Changes which apps show up as candidates to back up - checked, it
      also lists apps installed only for other Windows accounts on this
      PC, or provisioned machine-wide; unchecked, only your own account's
      apps appear. This only affects what's listed here - it doesn't
      change how the backup itself runs. Needs admin (which this whole
      app already requires) to see other accounts' apps at all. Unlike
      "Show:" and "Only show Start Menu apps" below, this one DOES need a
      fresh PowerShell query to take effect - click "Refresh List" after
      changing it.

  "Only show Start Menu apps" (on by default)
      Cross-references against the same list Windows itself uses for Start
      Menu/search (Get-StartApps), hiding invisible background/shell
      components that aren't frameworks but also aren't anything you'd
      recognize as an app to back up (Microsoft.Windows.ShellExperienceHost,
      Windows.CBSPreview, and similar). Genuinely launchable apps - first-
      party or third-party - still show up either way. Turn this off if you
      suspect it's hiding something you actually want (it's a heuristic,
      not a guarantee); has no effect when "Include all users" is checked,
      since Get-StartApps only sees the current user's own Start Menu.
      Instant, like "Show:" above - no re-fetch needed to take effect.

  "Refresh List"
      Re-scans installed apps and re-calculates their sizes. Can take a
      minute since it has to measure every app's folder. Clicking it again
      before the first scan finishes doesn't stack up two competing scans -
      only the most recently started one's result is ever applied.

  Search box
      Type to filter the list by app name or publisher as you type. Your
      checked apps stay checked even while hidden by a search filter.

  Per-column filters ("Filters:" row) and sortable column headers
      Free-text columns get their own small filter box on the first row;
      columns with only a fixed set of values (Type, Installed, Backup,
      Signed, etc.) get a dropdown of exactly those values on the
      second row, plus "Clear Filters" - so there's nothing to type or
      misspell for those. All active filters combine with each other and
      the search box above (everything must match), and each row wraps onto
      more lines as needed rather than overflowing the window. Click any
      column header to sort by it (▲/▼ shows the current direction); click
      the same header again to reverse it. Both tabs work the same way.

  "Show filters"
      Shows or hides the "Filters:" row(s) above - purely a space-saver for
      the list below; whatever filters are already set keep working even
      while hidden; they just take effect visibly again once shown. Both
      tabs work the same way.

  "Selected first"
      Groups every checked app at the top of the list, keeping whatever
      column sort is active as the order within each group - handy for
      reviewing exactly what you're about to back up or restore without
      hunting through a long, filtered list. The live "Selected: N app(s),
      ~X total" line below the list (next to the size estimate) always
      reflects what's actually checked, filter or no filter.

  "Select All" / "Select None"
      Bulk-check or uncheck whatever is currently visible (i.e. matching
      your search, if any). Only real, non-framework apps are ever listed
      here at all - shared runtime packages (VCLibs, .NET Native, DirectX
      Runtime, etc.) never appear as independently selectable items, since
      the backup automatically pulls in whatever frameworks your selected
      apps actually depend on, using a live lookup at the exact moment each
      app is backed up. Trying to independently back up a framework by
      itself used to be both redundant and fragile - a framework's exact
      installed version can change on its own (Windows Update manages
      these independent of anything you do), which could make an older,
      already-captured name for it fail to match by the time its own turn
      came up in a long backup - so this was removed rather than worked
      around case by case.

  "Force re-backup"
      Normally, backing up into a folder that already has a given app or
      dependency successfully backed up just skips it (fast, and avoids
      re-copying shared dependencies every time you back up one more app
      into the same folder). Check this to redo everything anyway.

  Size column / "Selected: N app(s), ~X" line
      An estimate of each app's on-disk size, and a running total for
      everything currently checked - use this to gauge how much backup
      space you'll need. The total does not include dependencies, since
      those are only resolved once the backup actually runs.

  "Show packing/signing tools"
      Shows or hides the box below once its paths are set correctly and
      don't need frequent attention - purely a space-saver, doesn't
      change whether packing/signing itself happens.

  "Packing / signing tools" box (MakeAppx.exe / SignTool.exe)
      Auto-filled if these Windows SDK tools are found on PATH or in the
      default SDK install folder - or, failing that, from whatever you set
      last time, since both paths are remembered across sessions once you
      set them (typing, browsing, or letting auto-detect fill one in all
      count), so a custom SDK install location only needs pointing at once.
      Click "Test" to actually confirm a path runs (not just that a file
      exists there), or "Browse…" to point at a copy in a nonstandard
      location. Leave both blank if you don't have the SDK - backup still
      works, it just produces a raw-folder-only backup (same-PC restore
      only, no portable .appx files).

  "Raw folder only" / "Packed only" / "Both" (default: Both)
      Controls what a backup actually produces for each app:
        Raw folder only - just the raw copy of the app's files (-SkipPack).
          Fastest, needs nothing beyond what's always required. Only usable
          for same-PC Register-mode restore afterward, since Install mode
          needs a packed .appx that was never created.
        Packed only - packs (and signs, unless SkipSign) as normal, then
          deletes the Raw folder afterward, keeping only the packed .appx
          (-DeleteRawAfterPack). Saves disk space, but trades away two
          things that specifically depend on the Raw folder still being
          there: Register-mode restore (which only ever reads Raw, never
          the packed .appx) won't work for this app anymore, and "Verify
          Files"/"Verify All Backups" will have nothing left to hash and
          compare against. Install-mode restore is unaffected, since it
          only ever reads the packed .appx. Shown afterward in the Restore
          tab's "Backup" column as "Packed-Only" - this is expected, not
          a sign anything went wrong; it's just what you asked this mode
          to leave behind.
        Both - the default, and the only way to keep every restore mode
          and integrity check available for an app later. Costs the most
          disk space of the three, since it's keeping everything.
      Applies to the whole batch you're about to start - not a per-app
      choice within one run. Switching it doesn't affect apps already
      backed up in a previous run; it only changes what a NEW run of
      Backup-AppxApps.ps1 (Start Backup) produces from here.

  "Start Backup"
      Runs Backup-AppxApps.ps1 once PER checked app (each pulling in its own
      dependencies automatically), one after another - not one big run
      covering everything at once. Checks the MakeAppx.exe/SignTool.exe
      paths above first if either is filled in, warning (with a chance to
      cancel) if one points at a file that doesn't actually exist there,
      rather than letting every single app in the batch fail to pack/sign
      the same way before you'd otherwise notice. Each app's actual command
      is only built right before that app runs, not all upfront - so fixing
      a wrong tool path partway through a multi-app batch (after seeing the
      first one or two fail) is picked up by every app still to come,
      rather than every already-queued app keeping the old, broken path
      that was in the field when the batch started. Progress for each app
      streams into the Log pane as it happens.

RESTORE TAB
  "Refresh List"
      Reads backup-summary.csv from the shared folder above and lists every
      app that was backed up. Apps NOT currently installed are checked by
      default; apps already installed start unchecked, since you usually
      only want to restore what's actually missing - use "Select Already
      Installed" if you specifically want to reinstall over an existing
      copy too. Also runs automatically whenever you pick a folder (via
      "Browse…" or the recent-folders dropdown) that already has a backup
      in it - you only need to click this yourself to re-check for changes
      made outside this tool since it last loaded. Always starts from a
      clean slate, so pointing this at a different folder never leaves
      stale rows from a previous one behind.

  Right-click a row -> "Show Backed-Up Files…"
      Lists every file actually present in that app's raw backup folder
      (with total size and whether a manifest file was found - if not, that
      backup is incomplete and restoring it will likely fail), plus whether
      a packed .appx exists. Worth checking before restoring something you
      haven't looked at in a while.

  Right-click a row -> "Verify Files (checklist + hashes)…"
      Opens a per-file checklist for that app. Every file's SHA256 hash
      inside the backup is computed automatically as soon as the window
      opens - no button needed for that part. Three ways to check it:
        "Check All" / "Check Selected" - compare against the SAME files in
          the app's CURRENT install (if it's still installed), by making a
          fresh backup-mode copy of it and hashing that - confirms your
          backup genuinely matches what's actually installed right now,
          file by file (Match / MISMATCH / Missing). A handful of files are
          EXPECTED to differ if the app was restored in Install mode and
          shown as "Expected diff (re-signed)" or "OS-generated (not app
          content)" instead of a real MISMATCH: AppxBlockMap.xml,
          AppxSignature.p7x, and CodeIntegrity.cat necessarily change when
          re-signed with this tool's own certificate, and the
          microsoft.system.package.metadata\\ files are Windows' own
          per-user bookkeeping, never part of the app's actual content -
          neither counts as a sign of a bad backup. Nothing to compare
          against if the app isn't currently installed - it'll say so.
        "Check Backup Integrity" - compares the backup folder against a
          record made at the moment it was originally backed up, so it
          catches files deleted or changed in the backup ITSELF since then
          (e.g. by accident) - this works even for apps no longer installed
          anywhere. Only available for backups made with this version or
          later (older backups have no such record to check against).

      Note on split packages (names containing "_split.", e.g. a language,
      scale, or architecture variant of an app): Windows doesn't expose
      these as independently queryable installed packages, so "Installed"
      will show No and "Check All"/"Check Selected" will say there's
      nothing to compare against - even when the app genuinely is
      installed. This is a Windows limitation, not a bug here. Check the
      app's main (non-split) package instead for a live comparison; "Check
      Backup Integrity" still works fine for split packages either way,
      since it doesn't depend on anything being installed.

      Note on non-framework dependencies: some of these have the same kind
      of Windows-side limitation as split packages above - not
      independently, reliably queryable via a direct installed-package
      lookup, which could otherwise show "Installed": No for one that's
      genuinely present. Since Windows enforces the dependency graph (it
      refuses to let a dependency be removed while something that still
      needs it remains installed), if any app in the list that depends on
      it shows as installed, this is treated as proof the dependency is
      too, regardless of what the direct lookup found. Frameworks aren't
      affected by this - they're already reliably queryable on their own,
      so this only ever applies to non-framework dependencies.

      "Copy Flagged Rows" button, or right-click any row for "Copy Cell" /
      "Copy Row" / "Copy All Rows": copies file details to the clipboard as
      plain text, for pasting somewhere else - handy if a file's status
      looks wrong and you want to note down exactly which one and why.

  "Size" column
      The raw backup folder's total on-disk size for that app, calculated
      live when you load the backup contents where that folder still
      exists. Falls back to the size recorded at backup time (from
      backup-summary.csv) if Raw itself is gone but was measured before
      it went - e.g. a "Packed only" backup, or Raw removed afterward -
      since that's the app's real installed footprint, which the packed
      .appx's own (often smaller, sometimes compressed) size would
      otherwise understate. Only falls back to the packed file's own size
      as a last resort, when neither of those is available; shows "?"
      when nothing at all is.

  "Check all users' installs (-AllUsers)"
      Only affects how the "Installed" column is detected, and therefore
      the two Select buttons below that rely on it (checked, it also
      counts something as "installed" if any other Windows account on
      this PC has it, not just your own). It does NOT change who the
      actual restore installs for or how - that's determined separately
      by Restore-AppxApps.ps1's own Install/Register mode logic. Needs
      admin (already required to run this tool at all).

  "Select:" combobox + "Confirm" (Apps / Dependencies / Frameworks / Installed / Not Installed / Raw Backup / Packed Backup / Missing)
      Consolidates what used to be several separate buttons - pick an
      option, then click "Confirm" to actually apply it. Deliberately not
      instant on picking a value: this overwrites your own manual
      checkbox selections in the list, and applying that the moment the
      dropdown closes - with nothing visibly different in the combobox
      itself afterward - made it look like the picked value just sat
      there unapplied. A backup contains both the apps you actually
      checked when backing up, and whatever frameworks/dependencies those
      needed - the "Type" column shows which is which:
        Apps - packages nothing else in this backup depends on (not based
          on how it was originally selected - see the "Type" column help
          below for why that's no longer how this is decided).
        Dependencies - non-framework dependencies only; frameworks have
          their own separate option below, so this doesn't lump them in
          together.
        Frameworks - only framework/runtime packages (VCLibs, etc.).
        Installed / Not Installed - based on the "Installed" column
          (checked live against this PC every time you load a backup).
          "Installed" is mostly there for completeness/symmetry - less
          commonly useful, but there if you want it.
        Raw Backup / Packed Backup / Missing - based on the "Backup"
          column - which apps only have a raw folder, only a packed
          .appx, or (rare, and worth investigating if you see it) neither
          currently on disk.
      Handy for restoring only the app(s) you actually wanted without also
      reinstalling shared runtimes you might already have. When a backup
      first loads, only not-installed rows classified as an App (Type)
      get checked by default - never a framework or other dependency on
      its own, since restoring an app already pulls in whatever it needs
      automatically.

  "Special Actions" menu
      Holds "Clean Up Orphaned Frameworks…", "Clean Up Orphaned
      Dependencies…", "Create Packed Backups from Raw…", and "Create Raw
      Backups from Packed…" (all below) - a single menu button rather
      than a growing row of standalone buttons, with room for more
      actions like these later.

  "Show tips on startup" and "Open Summary CSV" (right of the Backup /
  Restore tabs, visible on both)
      "Open Summary CSV" opens a timestamped copy of backup-summary.csv from
      your temp folder, so viewing or editing it (e.g. in Excel) never locks
      or changes the real file the scripts use.
      "Show tips on startup" is one shared setting for both tabs. Controls the small "Tip" dialog shown once, shortly after the
      app itself finishes opening - unrelated to any specific backup or
      restore run. Tips themselves come from tips.json, alongside this
      script - add more lines there any time; no code changes needed.

  "Create Packed Backups from Raw…"
      Scans the whole backup for apps with a raw-only backup (no packed
      .appx yet) - the same way "Clean Up Orphaned Frameworks" scans
      for its own candidates, rather than only working on whatever's
      currently checked - and packs (and signs, unless SkipSign) each one
      directly from its existing Raw folder, skipping the robocopy step
      entirely since Raw is already there. Works even for an app no
      longer installed, since it never touches the live install at all -
      only what's already sitting in this backup. Shows a confirmation
      listing exactly which apps qualify before doing anything. Only ever
      adds a packed file where one doesn't already exist - an app that
      already has one (from a normal backup, or a previous run of this)
      is always left alone, with no override option.

  "Create Raw Backups from Packed…"
      The reverse: scans the whole backup for apps with a packed-only
      backup (no Raw yet) and adds Raw for each, without touching the
      existing packed file at all. Never installs anything to do this -
      for whichever of these apps are currently installed, copies from
      the live install exactly like a normal backup would (the more
      accurate source when it's actually available); for the rest,
      extracts Raw directly from the packed .appx instead, since a .appx
      file is itself a zip archive underneath. "Check all users' installs"
      (further up this tab) also applies here, for deciding whether an
      app installed only under a different account counts as installed
      for this purpose. Shows a confirmation listing exactly which apps
      qualify before doing anything. Only ever adds Raw where it doesn't
      already exist - an app that already has it is always left alone,
      with no override option.

  "Remove" and its mode combobox (Both / Raw / Packed)
      Removes every currently CHECKED app's backup at once - one combined
      confirmation covering all of them, rather than right-clicking and
      confirming each one individually. If any of them depend on shared
      frameworks, it correctly checks whether those would still be needed
      by something OUTSIDE this batch (removing two apps that share a
      framework together can make that framework genuinely orphaned even
      though removing just one of them wouldn't have).
      The combobox controls what actually gets deleted for each app:
        Both (default) - Raw folder, Packed .appx, and the Manifests JSON
          (which records Raw's own file hashes, so it's meaningless once
          Raw itself is gone) - its backup-summary.csv row is dropped
          entirely, same as always.
        Raw - just the Raw folder (and its now-meaningless Manifests
          JSON) - keeps the packed .appx, and updates the row rather than
          removing it, so the "Backup" column correctly shows "Packed-Only"
          afterward.
        Packed - just the Packed .appx - keeps Raw and its Manifests
          JSON (both still valid), updating the row so "Backup" shows
          "Raw-Only" afterward.
      This same combobox applies to "Clean Up Orphaned Frameworks…" too,
      not just this button - one shared setting for whatever you remove
      next there, not tied to one specific action. A single row's own
      "Remove Backup…" (right-click a row) has its own copy of this same
      combobox instead, right there in that dialog - it starts at whatever
      this one currently shows, but from that point on is that dialog's
      own independent choice, letting you pick a different mode for just
      that one removal without changing this one.
      Checked apps that don't actually have anything for the current mode
      to remove (e.g. mode is "Raw" but an app is packed-only, with
      no Raw at all) are left out before anything happens - if that
      leaves nothing left to do, you just see a short message and nothing
      is touched, rather than a confirmation dialog for an operation that
      would have done nothing anyway.
      Removing something updates just the affected row(s) in the list
      directly afterward - not a full "Refresh List" (which also re-runs
      the installed-status check against every remaining row, real time
      for a large backup) - so a partial removal that leaves something
      behind (Raw against a "Both" backup, say) correctly stays in
      the list with its "Backup" column updated, rather than needing a
      manual refresh to reappear.
      The confirmation dialog (here, "Clean Up Orphaned Frameworks…",
      and a single row's own "Remove Backup…") lists any dependencies the
      app(s) being removed also use, each marked whether removing it
      would affect anything else. "Select all" checks every one of them;
      "Select all safe to remove" only checks the ones marked "not used
      elsewhere" - leaving anything flagged as still used by another app
      unchecked, for you to decide on individually.

  "Clean Up Orphaned Frameworks…"
      Scans the whole backup for frameworks no longer referenced by any
      app's own Dependencies field - the kind of thing that can
      accumulate after removing several apps one at a time over time -
      and offers to remove them all in one pass. Pre-checked by default
      since these are genuinely unused by definition. Frameworks only,
      deliberately - a non-framework package with no current dependents
      is indistinguishable from an ordinary standalone app that simply
      has no dependents either (which describes most apps), so this
      never touches anything outside "Type" = Framework, where that
      ambiguity doesn't exist. See "Clean Up Orphaned Dependencies…"
      below for non-framework packages instead.

  "Clean Up Orphaned Dependencies…"
      The same idea, for non-framework packages - possible without the
      ambiguity above because Backup-AppxApps.ps1 now records, once, at
      backup time, whether a package was ever explicitly asked for or
      only ever added because something else needed it
      (AddedAsDependency) - a fact that doesn't change later just
      because whatever depended on it gets removed, unlike a live
      "is anything currently using this" check on its own. Only offers a
      package that's both currently unreferenced AND was never
      explicitly selected - an ordinary standalone app you asked for
      directly is never touched, no matter how few or many dependents it
      has. Pre-checked by default, same as the framework version, though
      slightly less absolute: a row from a backup made before this field
      existed has no recorded answer and defaults to being treated as
      fair game if currently unreferenced, rather than being silently
      excluded.

  "Check files aren't locked before modifying them" (off by default)
      Before a removal or a restore touches a backup's files, checks each
      one can actually be opened for writing right now - catching a file
      that's mid-scan by antivirus, still syncing via OneDrive or similar,
      open in another program, or marked read-only, any of which could
      otherwise cause the operation to fail partway through with a
      confusing error. If any look locked, you're shown which ones and
      can choose to proceed anyway or cancel.
      Off by default since scanning every file first adds time on top of
      the operation itself - proportional to how many files are involved,
      so negligible for one app but potentially several seconds for a
      large bulk removal. The scan itself always runs in the background
      regardless, the same way the removal/restore it's checking for
      already does, so even with it on this never freezes the window
      while it works.
      Also writes a small marker file in the backup folder for the
      duration of the operation, and checks for one before starting -
      not a true file lock (this tool's own file access during the
      operation still needs to go through normally), but a safety net
      against a genuinely concurrent operation on the same folder, most
      plausibly another window of this app. A marker older than 6 hours
      is treated as left over from a crash rather than a live operation,
      so it can never block you forever.

  "Export Summary…" / "Open Last Summary"
      Writes a plain-text (or CSV) list of everything in this backup - name,
      version, publisher, and whether it's an App, Framework, or
      Dependency (see the "Type" column above) - independent of
      backup-summary.csv's own internal format, for record-keeping or
      sharing with someone else.
      Defaults to a timestamped filename so repeated exports don't
      overwrite each other. Logged to the main log pane when it completes.
      "Open Last Summary" reopens whichever one you exported most recently
      in this session, in its default application.

  "Verify All Backups…" / "Open Last Full Report"
      Re-hashes every backed-up package's Raw\\ files and compares them
      against what was recorded at the moment it was backed up (hash
      comparison is case-insensitive, since PowerShell's Get-FileHash and
      Python's own hashing report the same hash in different letter case -
      a real difference in content is still caught either way). Runs in
      the background (can take a while for a large backup); logs a start
      line and a one-line summary to the main log when done, plus which
      packages had issues if any did (never the full per-file detail
      though, to keep the log itself short), and shows the
      complete pass/fail list in its own results window, where a long
      detail can be read in full and the whole thing can be exported with
      "Export Full Report…". "Open Last Full Report"
      reopens whichever one was written most recently in this session
      (auto-saved or manually exported), in its default application.
      Independent of whether anything is still installed to compare
      against - this only checks whether the backup itself, sitting on
      disk, still matches what it should. Worth running before archiving
      a backup or moving it to another drive.

  Search box / Select All / Select None
      Same idea as the Backup tab, filtering and bulk-selecting the backup's
      contents instead of your installed apps.

  Column: "Type" (App / Framework / Dependency)
      What kind of package this row is - Framework for a framework/
      runtime package; Dependency for anything else that at least one
      other package in this backup still lists in its own Dependencies;
      App for everything else. Not based on how it was originally
      selected for backup (there's no reliable way to know that - picking
      a dependency directly on the Backup tab makes it just as
      "explicitly selected" as anything else) - purely on what it
      actually is and what currently depends on it, using data already
      in the backup. See the "Select:" combobox above for picking by
      this column.

  Column: "Backup" (Both / Raw-Only / Packed-Only / Missing)
      What's actually present on disk for that app right now - reflects
      the backup mode used (see "Raw folder only" / "Packed only" / "Both"
      above) and stays accurate even if a "Packed only" backup's Raw
      folder was deliberately removed after packing, or if either turned
      out to genuinely be missing/corrupted since backup time. "Raw-Only"
      usually means Install-mode restore isn't possible for that app,
      since it needs a packed .appx; "Packed-Only" means Register mode
      isn't possible, since it only ever reads the Raw folder. "Missing"
      means neither is currently present - it won't be restorable in
      either mode at all right now.

  Column: "Signed" (True / False / NA)
      Whether the packed .appx was successfully signed at backup time.
      Shows "NA" rather than False for a Raw-Only backup - there's no
      packed file to sign in that case, so False would misleadingly look
      like signing was attempted and failed rather than never applicable.
      Signed = False (with a packed file present) means Install mode
      likely won't work for it even though Register mode might.

  Mode: "Register (same PC)"
      Meant for the exact same Windows installation the backup was made
      from (e.g. after a Windows reset/reinstall), and lighter-weight (no
      certificates needed) - but in practice has turned out unreliable for
      real signed Store packages on at least some Windows/PowerShell
      combinations, failing with a generic "manifest not in package root"
      error even from a clean local copy. Kept available in case it works
      fine for you, but not the default anymore.

  Mode: "Install (this or another PC)" - the default
      Works for the same PC or a different one, and has proven reliable in
      testing so far. Imports the backup's self-signed certificate(s) into
      Trusted People, then installs the packaged .appx files with their
      dependencies (and any resource/split packages) wired up automatically.
      Many Windows 10/11 PCs already have sideloading trust on by default
      (e.g. if Developer Mode was ever turned on before) - try it once
      WITHOUT "Enable sideloading" first, and only turn that on if the
      install actually fails on a trust-related error.

  "Enable sideloading (Install mode)"
      Sets the sideloading registry keys for BOTH Windows 10/11 and Windows
      8/8.1 (they're different keys entirely). On Windows 8/8.1 Pro editions
      that aren't Enterprise or domain-joined, this alone may not be enough -
      Microsoft also requires a separate sideloading product activation key
      via slmgr.exe for those specific editions. On Windows 10/11 you can
      also do this by hand under Settings > Update & Security > For
      developers.

  A successful install/registration is not a guarantee the app will actually
  RUN. Some apps (particularly paid games) perform their own internal
  license/entitlement check against Microsoft's Store servers when they
  launch, separate from the sideloading trust this tool sets up - a
  sideloaded copy can fail to open with a generic error even though
  Add-AppxPackage reported success. That's the app's own logic, not
  something this tool (or any backup/restore approach) can affect.

  "Restore"
      Runs Restore-AppxApps.ps1 for exactly the apps you checked, in
      whichever mode is selected. Before doing anything, it checks that
      every dependency your selected apps need is actually present (and in
      good shape) in this backup, and warns in the Log pane about any gaps
      it finds - restoring still proceeds, but this tells you up front why
      something might fail rather than leaving you to guess afterwards.

  "Cancel" (Backup and Restore tabs)
      Stops the batch after whichever app is currently being processed
      finishes - not mid-way through it. The running app's own operation
      is left to complete normally either way, so nothing is interrupted
      partway through and potentially left in an inconsistent state; only
      whatever was still queued behind it is skipped. Only does anything
      while a backup or restore is actually running.

LOG PANE (bottom of window)
  Live output from whichever PowerShell script is currently running. Errors
  and warnings show up here even if the operation as a whole succeeds for
  most apps - scroll up if something looks like it silently failed.

  Every line is also written to a file on disk in real time, as soon as it
  appears - not just kept in memory - so nothing is lost if the app closes
  unexpectedly mid-operation. "Saving to:" shows the current session's log
  file path; "Open Folder" opens it in File Explorer. A fresh log file
  is started each time you launch the app.

  "Clear" - wipes the visible log AND starts a genuinely new session
  log file (shown in "Saving to:") rather than just clearing the display -
  whatever was already written to the old file stays exactly as it was,
  untouched, in its own file.

  "Detach Log" - pops the log out into its own resizable window,
  useful on a small monitor where the main window doesn't have room to show
  much log text at once. "Reattach" brings it back; closing the detached
  window's X button does the same thing. Content carries over both ways, and
  new output keeps appearing wherever the log currently is. The detached
  window has its own "Find:" search bar - highlights every match, the count
  label shows which one you're on, and Enter/Shift+Enter or the ▲/▼ buttons
  step between them.
"""


def decode_process_output(raw: bytes) -> str:
    """
    Decode subprocess output that may be UTF-16 (common for Windows console
    tools/PowerShell when their output is redirected/piped, as it is here)
    or a normal single-byte encoding. Naively decoding UTF-16 bytes as
    CP1252/UTF-8 turns "Microsoft" into "M\\x00i\\x00c\\x00..." which looks
    like garbage or just "M".
    """
    if not raw:
        return ''
    if raw[:2] in (b'\xff\xfe', b'\xfe\xff') or raw.count(b'\x00') > len(raw) // 4:
        try:
            return raw.decode('utf-16')
        except UnicodeError:
            pass
    for enc in ('utf-8', 'cp1252', 'latin-1'):
        try:
            return raw.decode(enc)
        except UnicodeError:
            continue
    return raw.decode('utf-8', errors='replace')


def select_in_explorer(path):
    """Opens the containing folder with this exact file/folder highlighted,
    via the actual Windows Shell API (SHParseDisplayName +
    SHOpenFolderAndSelectItems) - the same mechanism Explorer, browsers,
    and download managers use internally for "Show in folder". This avoids
    shelling out to explorer.exe's own "/select," command-line syntax
    entirely, which turned out to be unreliable via subprocess regardless
    of how that argument was quoted. Raises OSError on failure."""
    path = os.path.normpath(os.path.abspath(path))
    shell32 = ctypes.windll.shell32
    ole32 = ctypes.windll.ole32
    ole32.CoInitialize(None)
    try:
        pidl = ctypes.c_void_p()
        hr = shell32.SHParseDisplayName(path, None, ctypes.byref(pidl), 0, None)
        if hr != 0 or not pidl:
            raise OSError(f"SHParseDisplayName failed (0x{hr & 0xFFFFFFFF:08X}) for: {path}")
        try:
            hr2 = shell32.SHOpenFolderAndSelectItems(pidl, 0, None, 0)
            if hr2 != 0:
                raise OSError(f"SHOpenFolderAndSelectItems failed (0x{hr2 & 0xFFFFFFFF:08X}) for: {path}")
        finally:
            ole32.CoTaskMemFree(pidl)
    finally:
        ole32.CoUninitialize()


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False


def relaunch_as_admin():
    params = subprocess.list2cmdline(sys.argv)
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)


def log_ts():
    """Date+time timestamp for a log line - used on operation-boundary
    lines (started/finished this app, this batch, this verify run, etc.)
    so a single session log spanning several operations over a long
    period (potentially days, if the app's left open) shows exactly when
    each one actually happened, not just when the session itself started.
    Plain datetime, safe to call from any thread (unlike anything
    touching a Tkinter variable)."""
    return datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def humansize(num_bytes):
    if num_bytes is None:
        return "?"
    n = float(num_bytes)
    for unit in ('B', 'KB', 'MB', 'GB', 'TB'):
        if n < 1024 or unit == 'TB':
            return f"{n:.1f} {unit}" if unit != 'B' else f"{int(n)} B"
        n /= 1024


def run_powershell_encoded(script, timeout=180):
    """Run a PowerShell snippet via -EncodedCommand (sidesteps quoting/escaping
    entirely) and return decoded stdout/stderr text."""
    encoded = base64.b64encode(script.encode('utf-16-le')).decode('ascii')
    proc = subprocess.run(
        ['powershell', '-NoProfile', '-NonInteractive', '-EncodedCommand', encoded],
        capture_output=True, timeout=timeout, creationflags=CREATE_NO_WINDOW
    )
    return decode_process_output(proc.stdout), decode_process_output(proc.stderr)


def get_installed_apps(all_users=False):
    """Returns EVERY installed package, with IsFramework, IsSystemSigned,
    and IsStartMenuVisible attached to each row - enough for the Backup
    tab's "Show:" combobox and "Only show Start Menu apps" checkbox to
    filter entirely client-side (see AppxBackupGUI._on_backup_category_changed)
    instead of re-querying PowerShell every time either one changes, which
    used to mean a fresh wait on every single switch.

    IsSystemSigned reflects SignatureKind=System per Microsoft's own docs -
    packages "signed by a certificate that's also used to sign the Windows
    Operating System" (their own example: Windows Settings) - deeply
    OS-integrated components, not really "apps" in the usual sense.
    SignatureKind was only introduced in Windows 10 version 1607, so this
    is accessed defensively (checked for existence first) and simply can't
    tell these apart on Windows 8.1/early Windows 10 - IsSystemSigned just
    comes back False there for everything, never throwing.

    IsStartMenuVisible cross-references Get-StartApps - the same list
    Windows itself uses for Start Menu/search. Confirmed available since
    Windows 8.1 (Microsoft's own "Scripting Guys" blog demonstrates it
    there); if it's ever unavailable or fails for any reason, this comes
    back True for everything (degrading to "don't filter by this" rather
    than hiding everything or erroring out). Not meaningful with
    all_users=True (Get-StartApps is scoped to the current user only), so
    it's forced True for every row in that case too."""
    start_apps_setup = (
        "$startFamilies = @(); "
        "try { "
        "  Import-Module StartLayout -ErrorAction SilentlyContinue; "
        "  $startFamilies = @((Get-StartApps -ErrorAction Stop).AppID | ForEach-Object { ($_ -split '!')[0] } | Select-Object -Unique) "
        "} catch { $startFamilies = @() }; "
        if not all_users else "$startFamilies = @(); "
    )
    start_menu_expr = (
        "$true" if all_users
        else "($startFamilies.Count -eq 0 -or $startFamilies -contains $_.PackageFamilyName)"
    )

    script = (
        "[Console]::OutputEncoding = [Text.Encoding]::UTF8; "
        f"{start_apps_setup}"
        f"$pkgs = Get-AppxPackage{' -AllUsers' if all_users else ''}; "
        "$pkgs | ForEach-Object { "
        "  $sigKind = if ($_.PSObject.Properties['SignatureKind']) { $_.SignatureKind } else { $null }; "
        "  $sz = 0; "
        "  if ($_.InstallLocation -and (Test-Path $_.InstallLocation)) { "
        "    $sz = (Get-ChildItem -LiteralPath $_.InstallLocation -Recurse -File -ErrorAction SilentlyContinue "
        "           | Measure-Object -Property Length -Sum).Sum "
        "  }; "
        "  [PSCustomObject]@{ "
        "    Name = $_.Name; PackageFullName = $_.PackageFullName; "
        "    Version = $_.Version.ToString(); Publisher = $_.Publisher; "
        "    IsFramework = $_.IsFramework; SizeBytes = $sz; "
        "    InstallLocation = $_.InstallLocation; "
        "    IsSystemSigned = ($sigKind -eq 'System'); "
        f"    IsStartMenuVisible = {start_menu_expr} "
        "  } "
        "} | ConvertTo-Json -Depth 3"
    )
    stdout, stderr = run_powershell_encoded(script)
    stdout = stdout.strip()
    if not stdout:
        raise RuntimeError(stderr.strip() or "No output from PowerShell (no apps found?).")
    data = json.loads(stdout)
    if isinstance(data, dict):
        data = [data]
    return data


def get_installed_package_fullnames(all_users=False):
    """Lightweight check (no size calculation) - just the set of PackageFullName
    strings currently installed, for annotating a restore list with whether
    each app is already present on this machine."""
    script = (
        "[Console]::OutputEncoding = [Text.Encoding]::UTF8; "
        f"(Get-AppxPackage{' -AllUsers' if all_users else ''} -PackageTypeFilter Main,Framework,Resource,Optional,Bundle).PackageFullName"
    )
    stdout, stderr = run_powershell_encoded(script)
    if not stdout.strip() and stderr.strip():
        raise RuntimeError(stderr.strip())
    return {line.strip() for line in stdout.splitlines() if line.strip()}


def load_backup_summary(backup_root):
    """Returns (rows, error). rows is a list of dicts straight from backup-summary.csv."""
    path = os.path.join(backup_root, 'backup-summary.csv')
    if not os.path.isfile(path):
        return None, f"No backup-summary.csv found in:\n{backup_root}\n\nRun a backup into this folder first, or browse to a different one."
    try:
        with open(path, encoding='utf-8-sig', newline='') as f:
            reader = csv.DictReader(f)
            rows = list(reader)
        return rows, None
    except Exception as e:
        return None, f"Could not read backup-summary.csv: {e}"


def get_raw_backup_size(backup_root, package_full_name):
    """Total on-disk size (bytes) of a package's raw backup folder, or None if missing."""
    raw_dir = os.path.join(backup_root, 'Raw', package_full_name)
    if not os.path.isdir(raw_dir):
        return None
    total = 0
    for dirpath, _, filenames in os.walk(raw_dir):
        for fn in filenames:
            try:
                total += os.path.getsize(os.path.join(dirpath, fn))
            except OSError:
                pass
    return total


def get_backup_size(backup_root, package_full_name, recorded_raw_size_bytes=None):
    """Size to show in the restore list's "Size" column - the raw folder's
    live total size where that's available (most accurate, since it
    reflects whatever's actually on disk right now), then the size
    recorded in backup-summary.csv's RawSizeBytes at backup time (passed
    in by the caller, which has the row) if Raw itself is gone but this
    was recorded before it went (e.g. a "Packed only" backup, or Raw
    removed afterward) - this is the app's real installed footprint,
    which a packed .appx's own (often compressed, always smaller) size
    would otherwise understate. Falls back to the packed .appx's own size
    only if neither of those is available, and returns None (shown as
    "?") only if nothing at all is."""
    size = get_raw_backup_size(backup_root, package_full_name)
    if size is not None:
        return size
    if recorded_raw_size_bytes:
        try:
            return int(recorded_raw_size_bytes)
        except (TypeError, ValueError):
            pass
    packed_path = os.path.join(backup_root, 'Packed', f'{package_full_name}.appx')
    try:
        return os.path.getsize(packed_path)
    except OSError:
        return None


def compute_file_hash(path, algo='sha256', chunk_size=1024 * 1024):
    h = hashlib.new(algo)
    with open(path, 'rb') as f:
        while True:
            chunk = f.read(chunk_size)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def get_raw_backup_file_list(backup_root, package_full_name):
    """Returns (raw_dir, [{'rel':, 'size':}, ...]) for a package's raw backup folder."""
    raw_dir = os.path.join(backup_root, 'Raw', package_full_name)
    files = []
    if os.path.isdir(raw_dir):
        for dirpath, _, filenames in os.walk(raw_dir):
            for fn in filenames:
                full = os.path.join(dirpath, fn)
                rel = os.path.relpath(full, raw_dir)
                try:
                    size = os.path.getsize(full)
                except OSError:
                    size = 0
                files.append({'rel': rel, 'size': size})
    return raw_dir, files


def get_recent_folders_path():
    base = os.environ.get('APPDATA') or tempfile.gettempdir()
    d = os.path.join(base, 'AppxBackupManager')
    try:
        os.makedirs(d, exist_ok=True)
    except OSError:
        pass
    return os.path.join(d, 'recent_folders.json')


def get_settings_path():
    """A small persisted-settings file, separate from recent_folders.json
    (which is a plain list, not a dict, and already has its own users out
    there - reusing it for a different shape risks silently losing
    someone's existing recent-folders list on upgrade). Currently just the
    MakeAppx.exe/SignTool.exe paths, so a custom Windows SDK install only
    needs pointing at once rather than every session."""
    base = os.environ.get('APPDATA') or tempfile.gettempdir()
    d = os.path.join(base, 'AppxBackupManager')
    try:
        os.makedirs(d, exist_ok=True)
    except OSError:
        pass
    return os.path.join(d, 'settings.json')


def load_settings():
    path = get_settings_path()
    if not os.path.isfile(path):
        return {}
    try:
        with open(path, encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_setting(key, value):
    """Updates one key in the persisted settings file, leaving any others
    already there untouched. Best-effort - a failure to write here isn't
    worth interrupting the user over, so it's silently ignored, same as
    remember_recent_folder."""
    settings = load_settings()
    settings[key] = value
    try:
        with open(get_settings_path(), 'w', encoding='utf-8') as f:
            json.dump(settings, f)
    except Exception:
        pass


def get_new_session_log_path():
    """A fresh log file path for this run of the app, under the same config
    folder as recent_folders.json - written to in real time so nothing is
    lost if the app closes unexpectedly mid-operation."""
    base = os.environ.get('APPDATA') or tempfile.gettempdir()
    d = os.path.join(base, 'AppxBackupManager', 'logs')
    try:
        os.makedirs(d, exist_ok=True)
    except OSError:
        pass
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    return os.path.join(d, f'session_{timestamp}.log')


def load_recent_folders():
    path = get_recent_folders_path()
    if not os.path.isfile(path):
        return []
    try:
        with open(path, encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception:
        return []


def remember_recent_folder(folder):
    """Adds `folder` to the front of the recent-folders list (deduped),
    trims to 5, and persists it. Returns the updated list."""
    folders = [f for f in load_recent_folders() if f != folder]
    folders.insert(0, folder)
    folders = folders[:5]
    try:
        with open(get_recent_folders_path(), 'w', encoding='utf-8') as f:
            json.dump(folders, f)
    except Exception:
        pass
    return folders


def find_child(root, tagname):
    """Find the first descendant element with this local tag name, ignoring
    XML namespaces (manifests declare several, and the tag name alone is
    enough to identify Properties/Logo)."""
    for el in root.iter():
        tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
        if tag == tagname:
            return el
    return None


def _find_logo_hint(root):
    """Finds the logo/tile image path hint from a parsed AppxManifest.xml
    root element. Tries <Properties><Logo> first (present in virtually all
    manifests), then falls back to the Square150x150Logo/Square44x44Logo
    attributes on <uap:VisualElements> - some manifests, especially
    converted or unusual ones, only declare it that way."""
    props = find_child(root, 'Properties')
    if props is not None:
        logo_el = find_child(props, 'Logo')
        if logo_el is not None and logo_el.text and logo_el.text.strip():
            return logo_el.text.strip()

    for el in root.iter():
        tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
        if tag == 'VisualElements':
            for attr in ('Square150x150Logo', 'Square44x44Logo'):
                val = el.attrib.get(attr)
                if val and val.strip():
                    return val.strip()
    return None


def _resolve_logo_bytes(z, names, logo_hint):
    """Given an open ZipFile, its namelist, and a logo hint path from the
    manifest, finds and returns the best-matching image's bytes. The hint
    is only a HINT, not necessarily an exact filename - modern apps ship
    several scaled variants (Logo.scale-100.png, Logo.scale-200.png, ...)
    rather than one file at that exact path, so this falls back to a
    best-match search alongside an exact-path check."""
    logo_hint_norm = logo_hint.replace('\\', '/')

    for n in names:
        if n.replace('\\', '/').lower() == logo_hint_norm.lower():
            return z.read(n)

    base_dir = os.path.dirname(logo_hint_norm)
    base_name = os.path.basename(logo_hint_norm)
    stem = base_name.split('.')[0]
    ext = os.path.splitext(base_name)[1] or '.png'

    candidates = [
        n for n in names
        if os.path.dirname(n.replace('\\', '/')).lower() == base_dir.lower()
        and os.path.basename(n).lower().startswith(stem.lower())
        and n.lower().endswith(ext.lower())
    ]
    if not candidates:
        candidates = [
            n for n in names
            if os.path.basename(n).lower().startswith(stem.lower())
            and n.lower().endswith(ext.lower())
        ]
    if candidates:
        for c in candidates:
            if 'scale-100' in c.lower():
                return z.read(c)
        return z.read(candidates[0])
    return None


@functools.lru_cache(maxsize=256)
def find_manifest_icon_bytes(zip_path):
    """Extract the app's logo/tile image bytes from a packed .appx (or
    .appxbundle) file, for the hover-preview tooltip. Returns bytes or
    None. Cached by path since the same file gets hovered repeatedly.

    A .appxbundle's OUTER zip never contains AppxManifest.xml directly - it
    only has an AppxBundleManifest.xml plus several NESTED .appx files
    (each themselves a full zip), and the real per-app manifest (with the
    actual Logo reference) lives inside one of those, so bundles need an
    extra step of opening each nested .appx as its own zip. This tool's own
    backup/restore pipeline only ever produces plain .appx files, but this
    still guards against being pointed at a bundle some other way.
    """
    try:
        with zipfile.ZipFile(zip_path) as z:
            names = z.namelist()

            manifest_name = next((n for n in names if n.lower() == 'appxmanifest.xml'), None)
            if manifest_name:
                root = ET.fromstring(z.read(manifest_name))
                hint = _find_logo_hint(root)
                if hint:
                    icon = _resolve_logo_bytes(z, names, hint)
                    if icon:
                        return icon

            for nested_name in [n for n in names if n.lower().endswith('.appx')]:
                try:
                    with zipfile.ZipFile(io.BytesIO(z.read(nested_name))) as nz:
                        nnames = nz.namelist()
                        nmanifest = next((n for n in nnames if n.lower() == 'appxmanifest.xml'), None)
                        if not nmanifest:
                            continue
                        nroot = ET.fromstring(nz.read(nmanifest))
                        nhint = _find_logo_hint(nroot)
                        if nhint:
                            icon = _resolve_logo_bytes(nz, nnames, nhint)
                            if icon:
                                return icon
                except Exception:
                    continue

            return None
    except Exception:
        return None


@functools.lru_cache(maxsize=256)
def find_folder_icon_bytes(folder_path):
    """Same idea as find_manifest_icon_bytes, but for an on-disk app folder
    (an installed app's InstallLocation, or a Raw\\<PackageFullName> backup
    folder) rather than a packed .appx. Returns bytes or None. Only checks
    the manifest's own folder for the best-match fallback (not a full
    recursive walk) - app folders can be huge, and a slow first hover would
    be worse than occasionally not finding a scaled variant."""
    try:
        manifest_path = os.path.join(folder_path, 'AppxManifest.xml')
        if not os.path.isfile(manifest_path):
            return None
        with open(manifest_path, 'rb') as f:
            root = ET.fromstring(f.read())
        logo_hint = _find_logo_hint(root)
        if not logo_hint:
            return None
        logo_hint = logo_hint.replace('\\', '/')

        exact_path = os.path.join(folder_path, *logo_hint.split('/'))
        if os.path.isfile(exact_path):
            with open(exact_path, 'rb') as f:
                return f.read()

        base_dir = os.path.join(folder_path, os.path.dirname(logo_hint))
        base_name = os.path.basename(logo_hint)
        stem = base_name.split('.')[0]
        ext = os.path.splitext(base_name)[1] or '.png'
        if not os.path.isdir(base_dir):
            return None

        candidates = [
            os.path.join(base_dir, fn) for fn in os.listdir(base_dir)
            if fn.lower().startswith(stem.lower()) and fn.lower().endswith(ext.lower())
        ]
        if candidates:
            for c in candidates:
                if 'scale-100' in c.lower():
                    with open(c, 'rb') as f:
                        return f.read()
            with open(candidates[0], 'rb') as f:
                return f.read()
        return None
    except Exception:
        return None


class WrappingFrame(ttk.Frame):
    """A frame that lays out its children left to right, wrapping to a new
    row when they'd overflow the frame's current width - used for the
    per-column filter row, since a list with many columns would otherwise
    overflow the window instead of wrapping."""

    def __init__(self, parent, **kw):
        super().__init__(parent, **kw)
        self._items = []
        self.bind('<Configure>', self._reflow)

    def add_widget(self, widget):
        self._items.append(widget)
        self._reflow()

    def _reflow(self, event=None):
        width = self.winfo_width()
        if width <= 1:
            self.after(10, self._reflow)
            return
        x = 0
        row = 0
        col = 0
        for w in self._items:
            w.update_idletasks()
            w_width = w.winfo_reqwidth()
            if x + w_width > width and col > 0:
                row += 1
                col = 0
                x = 0
            w.grid(in_=self, row=row, column=col, sticky='w', padx=(0, 4), pady=2)
            x += w_width + 4
            col += 1


class IconTooltip:
    """Shows a small floating icon preview near the cursor when hovering a
    Treeview row. icon_provider(row_id) should return image bytes (PNG) or
    None; the byte-fetching itself does its own caching by path since row
    ids are regenerated on every re-render (e.g. search filtering).
    text_provider(row_id), if given, should return a string (or None) shown
    above the icon - e.g. the row's full package name, handy for reading
    in full without widening that column permanently in the list itself."""

    def __init__(self, tree, icon_provider, text_provider=None):
        self.tree = tree
        self.icon_provider = icon_provider
        self.text_provider = text_provider
        self.tip_win = None
        self.tip_image = None  # keep a reference so Tk doesn't garbage-collect it
        self.current_row = None
        tree.bind('<Motion>', self._on_motion, add='+')
        tree.bind('<Leave>', self._on_leave, add='+')

    def _on_motion(self, event):
        row_id = self.tree.identify_row(event.y)
        if row_id != self.current_row:
            self._hide()
            self.current_row = row_id
            if row_id:
                self._maybe_show(row_id, event)
        elif row_id and self.tip_win:
            self._reposition(event)

    def _on_leave(self, event):
        self._hide()
        self.current_row = None

    def _maybe_show(self, row_id, event):
        try:
            raw_bytes = self.icon_provider(row_id)
        except Exception:
            raw_bytes = None
        text = None
        if self.text_provider:
            try:
                text = self.text_provider(row_id)
            except Exception:
                text = None

        img = None
        if raw_bytes:
            try:
                img = tk.PhotoImage(data=raw_bytes)
                w = img.width()
                target = 128
                # Some icons come in noticeably smaller than others (a
                # 44px tile vs a 150px one), which looked inconsistent -
                # zoom (integer upscale) handles the small ones, subsample
                # (integer downscale) the large ones, so everything lands
                # close to the same target size either way. Neither op
                # supports a fractional/exact factor, so this gets close
                # rather than pixel-perfect, but it's a real improvement
                # over only ever shrinking, never enlarging.
                if w > target:
                    factor = max(1, w // target)
                    img = img.subsample(factor, factor)
                elif 0 < w < target:
                    factor = max(1, target // w)
                    img = img.zoom(factor, factor)
            except Exception:
                img = None  # not a format Tk's built-in PhotoImage can decode (e.g. JPEG)

        if not img and not text:
            return

        self.tip_image = img
        self.tip_win = tk.Toplevel(self.tree)
        self.tip_win.wm_overrideredirect(True)
        try:
            self.tip_win.attributes('-topmost', True)
        except tk.TclError:
            pass
        if text:
            tk.Label(
                self.tip_win, text=text, borderwidth=1, relief='solid', background='white',
                font=('TkDefaultFont', 8), wraplength=400, justify='left'
            ).pack(fill='x')
        if img:
            tk.Label(self.tip_win, image=img, borderwidth=1, relief='solid', background='white').pack()
        self._reposition(event)

    def _reposition(self, event):
        if self.tip_win:
            x = self.tree.winfo_rootx() + event.x + 18
            y = self.tree.winfo_rooty() + event.y + 18
            self.tip_win.wm_geometry(f'+{x}+{y}')

    def _hide(self):
        if self.tip_win:
            try:
                self.tip_win.destroy()
            except tk.TclError:
                pass
            self.tip_win = None
        self.tip_image = None


def find_locked_files(paths):
    """Returns the subset of `paths` that appear to be locked by another
    process or otherwise not currently modifiable - opening each briefly
    in read+write mode without truncating, which fails with a
    PermissionError/OSError if something else holds an exclusive lock on
    it (or it's marked read-only). Not a perfect test (a file locked only
    against writes by a shared-read handle would still show up here, and
    some legitimate shared-read locks may not), but it catches the
    realistic, common case: antivirus mid-scan, cloud sync (OneDrive,
    etc.) still writing it, Explorer or another program having it open,
    or a stale read-only flag - any of which would otherwise cause a
    restore or removal to fail partway through with a confusing error.
    Missing files are skipped entirely - nothing to lock there."""
    locked = []
    for p in paths:
        if not os.path.isfile(p):
            continue
        try:
            with open(p, 'r+b'):
                pass
        except (PermissionError, OSError):
            locked.append(p)
    return locked


def gather_backup_files_for(backup_root, full_names):
    """Every file under Raw\\<name>\\ plus the Packed .appx (if present)
    for each given package - the actual external files a restore or a
    removal is about to read from or delete."""
    paths = []
    for name in full_names:
        raw_dir = os.path.join(backup_root, 'Raw', name)
        if os.path.isdir(raw_dir):
            for root_dir, _, files in os.walk(raw_dir):
                for f in files:
                    paths.append(os.path.join(root_dir, f))
        packed_path = os.path.join(backup_root, 'Packed', f'{name}.appx')
        if os.path.isfile(packed_path):
            paths.append(packed_path)
    return paths


LOCK_MARKER_NAME = '.appx-backup-manager.lock'


def get_backup_manifest(backup_root, package_full_name):
    """
    Returns {rel_path: {'hash':, 'size':}} recorded by Backup-AppxApps.ps1 at
    the moment this package was actually backed up, or None if there isn't
    one (e.g. a backup made with an older version of the script). This is
    the fixed reference point for detecting files deleted or changed in the
    Raw\\ folder itself since the backup was made - independent of whether
    the app is even still installed to compare against.
    """
    path = os.path.join(backup_root, 'Manifests', f'{package_full_name}.json')
    if not os.path.isfile(path):
        return None
    try:
        with open(path, encoding='utf-8-sig') as f:
            data = json.load(f)
        if isinstance(data, dict):
            data = [data]
        return {entry['rel']: entry for entry in data}
    except Exception:
        return None


def get_installed_file_hashes(package_full_name, rel_paths, all_users=False):
    """
    If `package_full_name` is currently installed, robocopies it (backup-mode,
    same technique Backup-AppxApps.ps1 uses) into a temp folder and computes
    SHA256 for the requested relative paths there, for comparing against a
    backup without needing to touch the locked-down install folder directly.
    Returns ({rel_path: hash_or_None}, None) on success, or (None, message)
    if the app isn't currently installed. `all_users` must match whatever
    was used to determine the "Installed" column, or an app only visible
    under -AllUsers scope will wrongly look not-installed here.
    """
    script_check = (
        "[Console]::OutputEncoding = [Text.Encoding]::UTF8; "
        f"(Get-AppxPackage{' -AllUsers' if all_users else ''} -PackageTypeFilter Main,Framework,Resource,Optional,Bundle | Where-Object {{ $_.PackageFullName -eq "
        f"'{package_full_name}' }}).InstallLocation"
    )
    stdout, stderr = run_powershell_encoded(script_check)
    install_location = stdout.strip()
    if not install_location:
        lower_name = package_full_name.lower()
        if '_split.' in lower_name or '.split.' in lower_name:
            return None, (
                "This is a resource/split package (a language, scale, or architecture variant, e.g. "
                "'...neutral_split.language-es_...') - Windows doesn't expose these as independently "
                "queryable installed packages via Get-AppxPackage, so there's nothing to directly compare "
                "against here even though the app is genuinely installed. Check the main package instead "
                "(the one without '_split.' in its name, e.g. the plain x86/x64/neutral one) - that's the "
                "one Windows actually reports as installed."
            )
        return None, "This app is not currently installed on this PC - nothing to compare against."

    tmp_dir = tempfile.mkdtemp(prefix='appx_verify_')
    try:
        subprocess.run(
            ['robocopy', install_location, tmp_dir, '/E', '/ZB', '/R:1', '/W:1', '/NFL', '/NDL', '/NJH', '/NJS'],
            capture_output=True, timeout=300, creationflags=CREATE_NO_WINDOW
        )
        hashes = {}
        for rel in rel_paths:
            full = os.path.join(tmp_dir, rel)
            if os.path.isfile(full):
                try:
                    hashes[rel] = compute_file_hash(full)
                except Exception:
                    hashes[rel] = None
            else:
                hashes[rel] = None
        return hashes, None
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def get_app_backup_files(backup_root, package_full_name):
    """
    Human-readable listing of what's actually on disk for one package in a
    backup - the raw folder's full file listing (with a manifest-present
    check) and whether a packed .appx exists. Meant to be checked before
    restoring, so an incomplete/corrupted backup shows up before you try to
    install it rather than partway through.
    """
    lines = []
    raw_dir = os.path.join(backup_root, 'Raw', package_full_name)
    packed_path = os.path.join(backup_root, 'Packed', f'{package_full_name}.appx')

    lines.append(f"Raw folder: {raw_dir}")
    if os.path.isdir(raw_dir):
        entries = []
        for dirpath, _, filenames in os.walk(raw_dir):
            for fn in filenames:
                full = os.path.join(dirpath, fn)
                rel = os.path.relpath(full, raw_dir)
                try:
                    size = os.path.getsize(full)
                except OSError:
                    size = 0
                entries.append((rel, size))

        total_size = sum(size for _, size in entries)
        lines.append(f"  {len(entries)} file(s), {humansize(total_size)} total")

        has_manifest = any(
            os.path.basename(rel).lower() in ('appxmanifest.xml', 'wmappmanifest.xml')
            for rel, _ in entries
        )
        if has_manifest:
            lines.append("  Manifest present: Yes")
        else:
            lines.append("  Manifest present: NO - this backup looks INCOMPLETE, restoring it will likely fail.")

        lines.append("")
        for rel, size in sorted(entries):
            lines.append(f"  {size:>10,} B   {rel}")
    else:
        lines.append("  NOT FOUND - this app's raw backup is missing entirely.")

    lines.append("")
    if os.path.isfile(packed_path):
        lines.append(f"Packed .appx: {packed_path} ({humansize(os.path.getsize(packed_path))})")
    else:
        lines.append(f"Packed .appx: NOT FOUND at {packed_path}")

    return "\n".join(lines)


def find_sdk_tool(name):
    """
    Mirrors Backup-AppxApps.ps1's own Find-SdkTool: check PATH first, then
    the default Windows SDK install folder. Just for pre-filling the GUI's
    path fields - the PowerShell script does its own detection regardless.
    """
    for directory in os.environ.get('PATH', '').split(os.pathsep):
        candidate = os.path.join(directory, name)
        if os.path.isfile(candidate):
            return candidate
    candidates = glob.glob(rf'C:\Program Files (x86)\Windows Kits\10\bin\*\x64\{name}')
    if not candidates:
        candidates = glob.glob(rf'C:\Program Files (x86)\Windows Kits\10\bin\*\x86\{name}')
    return sorted(candidates)[-1] if candidates else None


def test_sdk_tool(path):
    """Try to locate and actually run a tool (makeappx.exe / signtool.exe). Returns (ok, message)."""
    if not path:
        return False, "No path given and nothing was auto-detected.\n\nInstall the Windows SDK, or browse to the .exe manually."
    if not os.path.isfile(path):
        return False, f"Not found at:\n{path}"
    try:
        proc = subprocess.run([path, '/?'], capture_output=True, timeout=15, creationflags=CREATE_NO_WINDOW)
        raw = proc.stdout or proc.stderr or b''
        output = decode_process_output(raw).strip()
        preview = '\n'.join(output.splitlines()[:6]) if output else '(no output captured)'
        return True, f"Found and ran successfully:\n{path}\n\nOutput preview:\n{preview}"
    except subprocess.TimeoutExpired:
        return False, f"Found at:\n{path}\n\nBut it did not respond within 15 seconds."
    except Exception as e:
        return False, f"Found at:\n{path}\n\nBut running it failed: {e}"


class CheckableAppList:
    """
    A reusable checkbox-list widget built on ttk.Treeview: search box, bulk
    select buttons, and a checkbox column, backed by a persistent
    (key -> checked) set so filtering never loses a selection.
    """

    def __init__(self, parent, columns, key_field, on_change=None, column_options=None, single_filter_row=False):
        """
        columns: list of (field, heading, width, anchor) tuples. `field` must
        be a key present in each row dict passed to set_rows().
        key_field: the field used as each row's unique identity for checking.
        column_options: optional {field: [value, ...]} - for a column whose
        values are always one of a small fixed set (e.g. True/False, Yes/No),
        this shows a dropdown of exactly those values (plus "(All)") instead
        of a free-text filter box, so there's nothing to type or misspell.
        single_filter_row: with few enough columns, splitting filters into a
        free-text row and a dropdown row (the default) wastes vertical space
        - set True to put everything in one wrapping row instead.
        """
        self.columns = columns
        self.key_field = key_field
        self.on_change = on_change
        self.column_options = column_options or {}
        self.single_filter_row = single_filter_row
        # Optional extra predicate the owning GUI can set (a row-dict ->
        # bool callable), consulted fresh on every _apply_filter() call - a
        # deliberately independent filtering layer, not proxied through any
        # of the per-column filter StringVars above, so a combobox driving
        # it can never drift out of sync with "Clear Filters" or a manual
        # edit to one of those per-column filters.
        self.extra_filter_fn = None
        self.all_rows = []
        self.checked = set()
        self.row_id_to_key = {}
        self.sort_col = None
        self.sort_reverse = False
        # A column whose displayed text isn't the right thing to sort by
        # (e.g. "size_display" is a formatted string like "12.3 MB", which
        # sorts wrong alphabetically) can list the row field to sort by
        # instead here.
        self.sort_key_overrides = {'size_display': 'SizeBytes'}

        self.frame = ttk.Frame(parent)

        search_bar = ttk.Frame(self.frame)
        search_bar.pack(fill='x', pady=(0, 4))
        ttk.Label(search_bar, text="Search:").pack(side='left')
        self.search_var = tk.StringVar()
        self.search_var.trace_add('write', lambda *a: self._apply_filter())
        ttk.Entry(search_bar, textvariable=self.search_var).pack(side='left', fill='x', expand=True, padx=4)
        ttk.Button(search_bar, text="✕", width=2, command=lambda: self.search_var.set('')).pack(side='left')
        ttk.Button(search_bar, text="Select All", command=lambda: self._bulk_check(True)).pack(side='left', padx=(10, 2))
        ttk.Button(search_bar, text="Select None", command=lambda: self._bulk_check(False)).pack(side='left', padx=2)
        self.selected_first_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            search_bar, text="Selected first", variable=self.selected_first_var, command=self._apply_filter
        ).pack(side='left', padx=(10, 0))
        self.filters_visible_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            search_bar, text="Show filters", variable=self.filters_visible_var, command=self._toggle_filters
        ).pack(side='left', padx=(10, 0))

        # Two rows: fixed-choice (combobox) columns first, then free-text
        # columns plus "Clear Filters" - keeps similar controls grouped
        # together rather than intermixed. Each row wraps independently if
        # it has more columns than fit on one line.
        combo_fields = [c for c in columns if c[0] in self.column_options]
        text_fields = [c for c in columns if c[0] not in self.column_options]

        filter_row1 = WrappingFrame(self.frame)
        label1 = ttk.Frame(filter_row1)
        ttk.Label(label1, text="Filters:").pack(side='left')
        filter_row1.add_widget(label1)

        if self.single_filter_row:
            # Few enough columns that a second row would just waste vertical
            # space - everything (text fields, dropdowns, Clear Filters)
            # goes in the one row, still wrapping if the window is narrow.
            filter_row2 = filter_row1
        else:
            filter_row1.pack(fill='x', pady=(0, 2))
            filter_row2 = WrappingFrame(self.frame)
            filter_row2.pack(fill='x', pady=(0, 4))

        self.column_filter_vars = {}

        for field, heading, width, anchor, stretch in text_fields:
            pair = ttk.Frame(filter_row1)
            ttk.Label(pair, text=f"{heading}:").pack(side='left', padx=(0, 2))
            var = tk.StringVar()
            var.trace_add('write', lambda *a: self._apply_filter())
            entry_width = max(6, min(12, width // 10))
            ttk.Entry(pair, textvariable=var, width=entry_width).pack(side='left')
            self.column_filter_vars[field] = var
            filter_row1.add_widget(pair)

        for field, heading, width, anchor, stretch in combo_fields:
            pair = ttk.Frame(filter_row2)
            ttk.Label(pair, text=f"{heading}:").pack(side='left', padx=(0, 2))
            values = ['(All)'] + list(self.column_options[field])
            # Set the initial value BEFORE registering the trace - setting
            # it after would fire _apply_filter() immediately, before
            # self.tree even exists yet (it's built further below).
            var = tk.StringVar(value='(All)')
            var.trace_add('write', lambda *a: self._apply_filter())
            combo_width = max(6, min(12, width // 10))
            ttk.Combobox(pair, textvariable=var, values=values, state='readonly', width=combo_width).pack(side='left')
            self.column_filter_vars[field] = var
            filter_row2.add_widget(pair)

        clear_pair = ttk.Frame(filter_row2)
        ttk.Button(clear_pair, text="Clear Filters", command=self._clear_column_filters).pack(side='left')
        filter_row2.add_widget(clear_pair)

        if self.single_filter_row:
            filter_row1.pack(fill='x', pady=(0, 4))

        # Stored so _toggle_filters can show/hide these rows (and, when
        # re-showing, put them back directly before the tree rather than
        # wherever pack() would otherwise re-append them).
        self.filter_rows = [filter_row1] if self.single_filter_row else [filter_row1, filter_row2]
        # Matches filters_visible_var's own default (off) - without this,
        # the rows built above would stay visibly packed even though the
        # checkbox itself shows unchecked, until manually toggled once.
        self._toggle_filters()

        tree_frame = ttk.Frame(self.frame)
        self.tree_frame = tree_frame
        tree_frame.pack(fill='both', expand=True)
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        col_ids = ['select'] + [c[0] for c in columns]
        self.tree = ttk.Treeview(tree_frame, columns=col_ids, show='headings', selectmode='browse')
        self.tree.heading('select', text='✓')
        self.tree.column('select', width=30, anchor='center', stretch=False)
        for field, heading, width, anchor, stretch in columns:
            self.tree.heading(field, text=heading, command=lambda f=field: self._sort_by(f))
            self.tree.column(field, width=width, anchor=anchor, stretch=stretch)

        vsb = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        self.tree.bind('<Button-1>', self._on_click)
        self.tree.bind('<Button-3>', self._on_right_click)

        # Shown in place of the tree whenever a filter/search/category
        # selection matches nothing, so an empty list reads as "nothing
        # matches this filter" rather than looking broken - occupies the
        # same cell as the tree and is only raised above it when needed.
        self.empty_label = ttk.Label(
            tree_frame, text="", foreground='#888', justify='center', anchor='center'
        )
        self.empty_label.grid(row=0, column=0, sticky='nsew')
        self.empty_label.grid_remove()

    def pack(self, **kw):
        self.frame.pack(**kw)

    def set_rows(self, rows, default_checked=False, default_checked_fn=None, preserve_checked=False):
        """default_checked_fn, if given, takes a row dict and returns whether
        it should start checked - overrides default_checked when provided,
        for cases like "only check apps not already installed".
        preserve_checked keeps whatever's currently checked (dropping only
        keys that no longer exist in the new rows) instead of recomputing
        from scratch - for a plain refresh of the same underlying data
        (not a switch to a genuinely different backup/folder), where
        losing the user's own manual selections as a side effect of an
        unrelated refresh would be a real workflow problem, not just a
        cosmetic one."""
        self.all_rows = rows
        valid_keys = {r[self.key_field] for r in rows}
        if preserve_checked:
            self.checked = self.checked & valid_keys
        elif default_checked_fn:
            self.checked = {r[self.key_field] for r in rows if default_checked_fn(r)}
        elif default_checked:
            self.checked = valid_keys
        else:
            self.checked = set()
        self._apply_filter()

    def _clear_column_filters(self):
        for field, var in self.column_filter_vars.items():
            var.set('(All)' if field in self.column_options else '')

    def _toggle_filters(self):
        """Shows or hides the per-column filter row(s) - purely a
        space-saver for the app list below, doesn't touch whatever
        filters are currently set (they just take effect again as soon
        as the rows are shown). before=self.tree_frame keeps them in
        their original position when re-shown, since pack() would
        otherwise just re-append them after the tree instead of putting
        them back above it - and packing them in the same order as they
        were originally added keeps that relative order intact too."""
        if self.filters_visible_var.get():
            for i, row in enumerate(self.filter_rows):
                pady = (0, 4) if i == len(self.filter_rows) - 1 else (0, 2)
                row.pack(fill='x', pady=pady, before=self.tree_frame)
        else:
            for row in self.filter_rows:
                row.pack_forget()

    def _sort_by(self, field):
        if self.sort_col == field:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_col = field
            self.sort_reverse = False
        for f, heading, _width, _anchor, _stretch in self.columns:
            if f == self.sort_col:
                arrow = ' ▼' if self.sort_reverse else ' ▲'
                self.tree.heading(f, text=heading + arrow)
            else:
                self.tree.heading(f, text=heading)
        self._apply_filter()

    def _sort_key_for(self, row, field):
        raw = row.get(self.sort_key_overrides.get(field, field), '')
        if isinstance(raw, (int, float)):
            return (0, raw)
        s = str(raw)
        try:
            return (0, float(s))
        except (TypeError, ValueError):
            return (1, s.lower())

    def _apply_filter(self):
        query = self.search_var.get().strip().lower()
        column_queries = {}
        for f, v in self.column_filter_vars.items():
            val = v.get().strip()
            if not val:
                continue
            if f in self.column_options and val == '(All)':
                continue
            column_queries[f] = val
        self.tree.delete(*self.tree.get_children())
        self.row_id_to_key.clear()

        rows = self.all_rows
        if self.extra_filter_fn is not None:
            rows = [r for r in rows if self.extra_filter_fn(r)]
        if query:
            rows = [r for r in rows if query in ' '.join(str(r.get(f, '')) for f, *_ in self.columns).lower()]
        for field, needle in column_queries.items():
            if field in self.column_options:
                rows = [r for r in rows if str(r.get(field, '')).lower() == needle.lower()]
            else:
                rows = [r for r in rows if needle.lower() in str(r.get(field, '')).lower()]
        if self.sort_col:
            rows = sorted(rows, key=lambda r: self._sort_key_for(r, self.sort_col), reverse=self.sort_reverse)
        if self.selected_first_var.get():
            # A second, stable sort pass - Python's sort preserves the
            # column-sort order established above WITHIN each of the two
            # groups (selected, then not), rather than needing to combine
            # both into one compound key (which would fight the column
            # sort's own reverse direction).
            rows = sorted(rows, key=lambda r: 0 if r[self.key_field] in self.checked else 1)

        if not rows:
            if self.all_rows:
                # Data has loaded, but the current filter/search/category
                # selection matches nothing - say so plainly rather than
                # leaving an unexplained blank list that could read as
                # broken.
                active = []
                if self.extra_filter_fn is not None:
                    active.append('the "Show:" selection')
                if query:
                    active.append('the search box')
                if column_queries:
                    active.append('a column filter')
                reason = ' and '.join(active) if active else 'the current filter'
                self.empty_label.config(
                    text=f"No matches - {reason} excluded every item here.\nClear or adjust it to see more."
                )
            else:
                self.empty_label.config(text="No items to show.")
            self.empty_label.grid()
            self.empty_label.tkraise()
        else:
            self.empty_label.grid_remove()

        for row in rows:
            key = row[self.key_field]
            mark = '☑' if key in self.checked else '☐'
            values = [mark] + [row.get(f, '') for f, *_ in self.columns]
            row_id = self.tree.insert('', 'end', values=values)
            self.row_id_to_key[row_id] = key
        if self.on_change:
            self.on_change()

    def _on_click(self, event):
        region = self.tree.identify('region', event.x, event.y)
        col = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if region == 'cell' and col == '#1' and row_id in self.row_id_to_key:
            key = self.row_id_to_key[row_id]
            if key in self.checked:
                self.checked.discard(key)
                self.tree.set(row_id, 'select', '☐')
            else:
                self.checked.add(key)
                self.tree.set(row_id, 'select', '☑')
            if self.selected_first_var.get():
                self._apply_filter()  # handles reordering and on_change together
            else:
                if self.on_change:
                    self.on_change()

    def _bulk_check(self, checked):
        mark = '☑' if checked else '☐'
        for row_id, key in self.row_id_to_key.items():
            self.tree.set(row_id, 'select', mark)
            if checked:
                self.checked.add(key)
            else:
                self.checked.discard(key)
        if self.selected_first_var.get():
            self._apply_filter()
        else:
            if self.on_change:
                self.on_change()

    def get_checked_rows(self):
        by_key = {r[self.key_field]: r for r in self.all_rows}
        return [by_key[k] for k in self.checked if k in by_key]

    def get_row_by_id(self, row_id):
        key = self.row_id_to_key.get(row_id)
        if key is None:
            return None
        return next((r for r in self.all_rows if r[self.key_field] == key), None)

    def _on_right_click(self, event):
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        self.tree.selection_set(row_id)
        col = self.tree.identify_column(event.x)
        values = self.tree.item(row_id, 'values')
        col_index = None
        if col and col.startswith('#'):
            try:
                col_index = int(col[1:]) - 1
            except ValueError:
                pass
        cell_value = values[col_index] if col_index is not None and 0 <= col_index < len(values) else ''

        menu = tk.Menu(self.tree, tearoff=0)
        menu.add_command(label="Copy Cell", command=lambda: self._copy_to_clipboard(str(cell_value)))
        menu.add_command(label="Copy Row (all columns)", command=lambda: self._copy_to_clipboard('\t'.join(str(v) for v in values)))
        menu.add_command(label="Copy All Visible Rows", command=self.copy_all_visible_rows)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _copy_to_clipboard(self, text):
        self.tree.clipboard_clear()
        self.tree.clipboard_append(text)

    def copy_all_visible_rows(self):
        """Copies the header + every row currently matching the search filter
        (tab-separated, so it pastes cleanly into a spreadsheet or a message)."""
        header = '\t'.join(['Select'] + [heading for _, heading, _, _, _ in self.columns])
        rows_text = [header]
        for row_id in self.tree.get_children():
            values = self.tree.item(row_id, 'values')
            rows_text.append('\t'.join(str(v) for v in values))
        if len(rows_text) <= 1:
            return
        self._copy_to_clipboard('\n'.join(rows_text))


def is_known_benign_diff(rel):
    """These files are EXPECTED to differ between the original raw backup
    and a currently-installed copy that went through this tool's own
    repackage-and-resign pipeline (Install mode), or are OS-generated
    per-user bookkeeping that was never part of the app's actual content -
    flagging them as MISMATCH/Missing is a false alarm, not a sign of a
    bad backup:
      - AppxBlockMap.xml / AppxSignature.p7x / CodeIntegrity.cat: these are
        the package's own signing/integrity metadata. Re-signing with this
        tool's self-signed certificate necessarily changes all three, since
        the signature (and everything derived from it) is different by
        design - not because any content is missing or corrupted.
      - microsoft.system.package.metadata\\...: Windows' own per-user
        dependency-tracking bookkeeping, generated after install and tied
        to whichever SID installed it - never part of the app's content.
    """
    rel_norm = rel.replace('\\', '/').lower()
    basename = rel_norm.rsplit('/', 1)[-1]
    if basename in ('appxblockmap.xml', 'appxsignature.p7x'):
        return True
    if rel_norm.endswith('appxmetadata/codeintegrity.cat'):
        return True
    if rel_norm.startswith('microsoft.system.package.metadata/'):
        return True
    return False


class FileVerifierDialog:
    """
    Per-app file checklist: shows every file in the raw backup with its size
    and SHA256 (computed automatically, no button needed), plus a live
    "Installed hash" column filled in on demand - "Check All" / "Check
    Selected" robocopy the currently-installed app (if present) into a temp
    folder and hash the same relative paths there, so you can confirm the
    backup genuinely matches what's/was actually installed.
    """

    def __init__(self, app, package_full_name, display_name, files):
        self.app = app
        self.package_full_name = package_full_name
        self.files = files  # list of {'rel':, 'size':}
        self.checked = {f['rel'] for f in files}  # all checked by default
        self.row_id_to_rel = {}
        self.backup_hashes = {}
        self.installed_hashes = {}
        self.manifest = get_backup_manifest(app.folder_var.get().strip(), package_full_name)

        self.win = tk.Toplevel(app.root)
        self.win.title(f"Verify Files — {display_name}")
        self.win.geometry("950x580")
        self.win.minsize(600, 350)

        btn_frame = ttk.Frame(self.win)
        btn_frame.pack(side='bottom', fill='x', pady=6, padx=8)
        self.status_var = tk.StringVar(value="Computing backup file hashes...")
        ttk.Label(btn_frame, textvariable=self.status_var).pack(side='left')
        ttk.Button(btn_frame, text="Close", command=self.win.destroy).pack(side='right')
        ttk.Button(btn_frame, text="Check Selected", command=self.check_selected).pack(side='right', padx=6)
        ttk.Button(btn_frame, text="Check All", command=self.check_all).pack(side='right')
        ttk.Button(btn_frame, text="Check Backup Integrity", command=self.check_backup_integrity).pack(side='right', padx=6)

        sel_frame = ttk.Frame(self.win)
        sel_frame.pack(side='top', fill='x', padx=8, pady=(8, 0))
        ttk.Button(sel_frame, text="Select All", command=lambda: self._bulk_check(True)).pack(side='left')
        ttk.Button(sel_frame, text="Select None", command=lambda: self._bulk_check(False)).pack(side='left', padx=4)
        ttk.Button(sel_frame, text="Copy Flagged Rows", command=self.copy_flagged_rows).pack(side='left', padx=(12, 0))
        ttk.Label(sel_frame, text="(right-click a row to copy just that one)", foreground='#666').pack(side='left', padx=6)

        search_row = ttk.Frame(self.win)
        search_row.pack(side='top', fill='x', padx=8, pady=(8, 0))
        ttk.Label(search_row, text="Search:").pack(side='left')
        self.search_var = tk.StringVar()
        self.search_var.trace_add('write', lambda *a: self._refresh_view())
        ttk.Entry(search_row, textvariable=self.search_var).pack(side='left', fill='x', expand=True, padx=4)
        ttk.Button(search_row, text="✕", width=2, command=lambda: self.search_var.set('')).pack(side='left')

        tree_frame = ttk.Frame(self.win)
        tree_frame.pack(side='top', fill='both', expand=True, padx=8, pady=8)
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        columns = ('select', 'file', 'size', 'backup_hash', 'installed_hash', 'status')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', selectmode='browse')
        self.tree.heading('select', text='✓')
        self.tree.heading('file', text='File')
        self.tree.heading('size', text='Size')
        self.tree.heading('backup_hash', text='Backup SHA256')
        self.tree.heading('installed_hash', text='Installed SHA256')
        self.tree.heading('status', text='Status')
        self.tree.column('select', width=30, anchor='center', stretch=False)
        self.tree.column('file', width=280)
        self.tree.column('size', width=90, anchor='e')
        self.tree.column('backup_hash', width=180)
        self.tree.column('installed_hash', width=180)
        self.tree.column('status', width=140, anchor='center')

        vsb = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        self.tree.bind('<Button-1>', self._on_click)
        self.tree.bind('<Button-3>', self._on_right_click)

        for f in sorted(files, key=lambda x: x['rel']):
            row_id = self.tree.insert('', 'end', values=('☑', f['rel'], humansize(f['size']), '…', '', ''))
            self.row_id_to_rel[row_id] = f['rel']

        threading.Thread(target=self._compute_backup_hashes_worker, daemon=True).start()

    def _refresh_view(self):
        """Filters which rows are visible per the search box - detach/move
        rather than delete/reinsert, since rows here live directly in the
        tree (inserted once, updated via .set() afterward) rather than a
        separate list re-rendered each time."""
        query = self.search_var.get().strip().lower()
        all_ids = list(self.row_id_to_rel.keys())
        value_cols = ('file', 'size', 'backup_hash', 'installed_hash', 'status')
        if not query:
            visible_ids = all_ids
        else:
            visible_ids = [
                rid for rid in all_ids
                if query in ' '.join(str(self.tree.set(rid, c)) for c in value_cols).lower()
            ]
        hidden_ids = set(all_ids) - set(visible_ids)
        for rid in hidden_ids:
            self.tree.detach(rid)
        for index, rid in enumerate(visible_ids):
            self.tree.move(rid, '', index)

    def _bulk_check(self, checked):
        mark = '☑' if checked else '☐'
        for row_id in self.tree.get_children():
            rel = self.row_id_to_rel.get(row_id)
            if rel is None:
                continue
            self.tree.set(row_id, 'select', mark)
            if checked:
                self.checked.add(rel)
            else:
                self.checked.discard(rel)

    def _on_click(self, event):
        region = self.tree.identify('region', event.x, event.y)
        col = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if region == 'cell' and col == '#1' and row_id in self.row_id_to_rel:
            rel = self.row_id_to_rel[row_id]
            if rel in self.checked:
                self.checked.discard(rel)
                self.tree.set(row_id, 'select', '☐')
            else:
                self.checked.add(rel)
                self.tree.set(row_id, 'select', '☑')

    def _on_right_click(self, event):
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        self.tree.selection_set(row_id)
        col = self.tree.identify_column(event.x)
        values = self.tree.item(row_id, 'values')
        col_index = None
        if col and col.startswith('#'):
            try:
                col_index = int(col[1:]) - 1
            except ValueError:
                pass
        cell_value = values[col_index] if col_index is not None and 0 <= col_index < len(values) else ''

        menu = tk.Menu(self.tree, tearoff=0)
        menu.add_command(label="Copy Cell", command=lambda: self._copy_to_clipboard(str(cell_value)))
        menu.add_command(label="Copy Row (all columns)", command=lambda: self._copy_to_clipboard('\t'.join(str(v) for v in values)))
        menu.add_command(label="Copy All Rows", command=self.copy_all_rows)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _copy_to_clipboard(self, text):
        self.win.clipboard_clear()
        self.win.clipboard_append(text)

    def copy_all_rows(self):
        header = "Select\tFile\tSize\tBackup SHA256\tInstalled SHA256\tStatus"
        rows_text = [header]
        for row_id in self.tree.get_children():
            values = self.tree.item(row_id, 'values')
            rows_text.append('\t'.join(str(v) for v in values))
        if len(rows_text) <= 1:
            return
        self._copy_to_clipboard('\n'.join(rows_text))

    def copy_flagged_rows(self):
        """Copies every row whose Status isn't OK/Match/blank - meant for
        pasting into a bug report when something like 'Not in original
        manifest' shows up unexpectedly."""
        header = "Select\tFile\tSize\tBackup SHA256\tInstalled SHA256\tStatus"
        flagged = []
        for row_id in self.tree.get_children():
            values = self.tree.item(row_id, 'values')
            status = values[5] if len(values) > 5 else ''
            if status and status not in ('OK', 'Match', 'Expected diff (re-signed)', 'OS-generated (not app content)'):
                flagged.append('\t'.join(str(v) for v in values))
        if not flagged:
            messagebox.showinfo("Nothing flagged", "No rows currently show a problem status.", parent=self.win)
            return
        self._copy_to_clipboard(header + '\n' + '\n'.join(flagged))
        messagebox.showinfo("Copied", f"Copied {len(flagged)} flagged row(s) to the clipboard.", parent=self.win)

    def _compute_backup_hashes_worker(self):
        raw_dir = os.path.join(self.app.folder_var.get().strip(), 'Raw', self.package_full_name)
        for row_id, rel in list(self.row_id_to_rel.items()):
            full = os.path.join(raw_dir, rel)
            try:
                h = compute_file_hash(full)
            except Exception:
                h = None
            self.backup_hashes[rel] = h
            display = h[:16] + '…' if h else 'ERROR'
            self.win.after(0, lambda r=row_id, d=display: self.tree.set(r, 'backup_hash', d))
        self.win.after(0, lambda: self.status_var.set("Backup hashes computed. Use Check All/Selected to compare against the current install."))

    def check_all(self):
        self._run_check(set(self.row_id_to_rel.values()))

    def check_selected(self):
        if not self.checked:
            messagebox.showinfo("Nothing selected", "Check at least one file first.")
            return
        self._run_check(set(self.checked))

    def check_backup_integrity(self):
        """
        Compares the CURRENT contents of the Raw\\ folder against the
        manifest recorded at backup time - catches files deleted or changed
        in the backup itself (e.g. by accident), independent of whether the
        app is even still installed. Reuses the already-computed backup
        hashes rather than re-hashing, so this is effectively instant once
        those finish.
        """
        if self.manifest is None:
            messagebox.showinfo(
                "No manifest",
                "This backup doesn't have a recorded file manifest (it may have been made with an "
                "older version of the script) - can't check for deletions/changes this way.",
                parent=self.win
            )
            return

        # Match paths in a normalized form (lowercase, forward slashes) rather
        # than exact string equality - a manifest entry and the matching
        # current file should always refer to the same file, but relative
        # path strings computed by two different processes (PowerShell at
        # backup time, Python here) can differ in casing or separator for a
        # handful of entries (e.g. deeply nested paths that trip Windows'
        # long-path handling), which would otherwise show up as false
        # "Not in original manifest" mismatches for files that are actually fine.
        def normalize(rel):
            return rel.lower().replace('\\', '/')

        manifest_by_norm = {normalize(rel): (rel, entry) for rel, entry in self.manifest.items()}
        current_norms = {normalize(rel) for rel in self.row_id_to_rel.values()}
        deleted = sorted(set(manifest_by_norm.keys()) - current_norms)

        pending = 0
        for row_id, rel in self.row_id_to_rel.items():
            match = manifest_by_norm.get(normalize(rel))
            if match is None:
                status = 'Not in original manifest'
            else:
                _, expected = match
                actual_hash = self.backup_hashes.get(rel)
                if actual_hash is None:
                    status, pending = 'Hash pending…', pending + 1
                elif actual_hash.lower() == expected.get('hash', '').lower():
                    status = 'OK'
                else:
                    status = 'CHANGED since backup'
            self.tree.set(row_id, 'status', status)

        for norm_rel in deleted:
            orig_rel, expected = manifest_by_norm[norm_rel]
            self.tree.insert('', 'end', values=(
                '', orig_rel, humansize(expected.get('size', 0)), expected.get('hash', '')[:16] + '…',
                '(file missing)', 'DELETED FROM BACKUP'
            ))

        msg = f"Backup integrity check: {len(deleted)} file(s) deleted since backup."
        if pending:
            msg += f" ({pending} still waiting on the initial hash pass - try again in a moment.)"
        self.status_var.set(msg)

    def _run_check(self, rels):
        self.status_var.set(f"Checking {len(rels)} file(s) against the current install...")
        threading.Thread(target=self._run_check_worker, args=(rels,), daemon=True).start()

    def _run_check_worker(self, rels):
        hashes, error = get_installed_file_hashes(
            self.package_full_name, list(rels), all_users=self.app.restore_all_users_var.get()
        )
        if error:
            self.win.after(0, lambda: self.status_var.set(error))
            self.win.after(0, lambda: messagebox.showinfo("Not installed", error, parent=self.win))
            return
        self.installed_hashes.update(hashes)
        for row_id, rel in self.row_id_to_rel.items():
            if rel not in rels:
                continue
            installed_hash = hashes.get(rel)
            backup_hash = self.backup_hashes.get(rel)
            if installed_hash is None:
                if is_known_benign_diff(rel):
                    status, display = 'OS-generated (not app content)', ''
                else:
                    status, display = 'Missing from install', ''
            elif backup_hash is None:
                status, display = 'Backup hash pending', installed_hash[:16] + '…'
            elif installed_hash == backup_hash:
                status, display = 'Match', installed_hash[:16] + '…'
            else:
                if is_known_benign_diff(rel):
                    status, display = 'Expected diff (re-signed)', installed_hash[:16] + '…'
                else:
                    status, display = 'MISMATCH', installed_hash[:16] + '…'
            self.win.after(0, lambda r=row_id, d=display, s=status: (
                self.tree.set(r, 'installed_hash', d), self.tree.set(r, 'status', s)
            ))
        self.win.after(0, lambda: self.status_var.set(f"Checked {len(rels)} file(s) against the current install."))


class AppxBackupGUI:
    def __init__(self, root):
        self.root = root
        root.title(f"Appx Backup Manager v{SCRIPT_VERSION}")
        root.geometry("1080x720")
        root.minsize(860, 540)

        self.log_queue = queue.Queue()
        self.current_proc = None
        self._batch_running = False  # True for the whole duration of _run_commands_worker, unlike
                                      # current_proc which is briefly None between items in a batch
        self._cancel_requested = False
        self.restore_load_token = 0  # guards against a stale background load overwriting a newer one
        self.backup_load_token = 0  # same idea, for refresh_installed_apps on the Backup tab
        self.last_summary_path = None  # most recently exported backup summary, for "Open Last Summary"
        self.last_verify_report_path = None  # most recently written verify report, for "Open Last Full Report"
        self._backup_folder_busy = False  # guards against overlapping removals/verify-all racing on the same files/CSV
        self._active_lock_root = None  # set by start_restore when it acquires a folder lock; released when the run finishes

        self._build_ui()
        self.root.after(100, self._poll_log_queue)
        self.root.after(300, self._maybe_show_tip_dialog)
        self.refresh_installed_apps()
        # Auto-load restore contents on startup too, the same as picking a
        # folder via Browse… or the recent-folders dropdown already does -
        # only if the pre-filled folder above (the most recently used one)
        # actually has a backup in it, same condition as those two paths.
        initial_folder = self.folder_var.get().strip()
        if initial_folder and os.path.isfile(os.path.join(initial_folder, 'backup-summary.csv')):
            self.load_restore_list()

    # ---------------- UI construction ----------------

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=8)
        top.pack(fill='x')
        ttk.Label(top, text="Backup / restore folder:").pack(side='left')
        self.recent_folders = load_recent_folders()
        self.tips = load_tips()
        self.show_tips_var = tk.BooleanVar(value=load_settings().get('show_tips', True))
        self.show_tips_var.trace_add('write', lambda *a: save_setting('show_tips', self.show_tips_var.get()))
        # Pre-filled with the most recently used folder (recent_folders[0],
        # since remember_recent_folder always inserts at the front) rather
        # than starting empty every time - saves re-picking or re-typing
        # the same folder each session, and the auto-load-if-it-has-a-
        # backup-already check at the end of __init__ (below, once the
        # restore tab and its widgets actually exist) covers this the same
        # way picking it manually from Browse… or the dropdown already does.
        self.folder_var = tk.StringVar(value=self.recent_folders[0] if self.recent_folders else '')
        self.folder_combo = ttk.Combobox(top, textvariable=self.folder_var, values=self.recent_folders)
        self.folder_combo.pack(side='left', fill='x', expand=True, padx=4)
        self.folder_combo.bind('<<ComboboxSelected>>', self._on_recent_folder_selected)
        ttk.Button(top, text="Browse…", command=self._browse_folder).pack(side='left')
        ttk.Button(top, text="Open Folder", command=self._open_backup_root_folder).pack(side='left', padx=(6, 0))
        ttk.Button(top, text="Help", command=self._show_help).pack(side='right')

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=8, pady=4)

        # Shared controls placed on the empty right side of the tab strip
        # (next to the Backup / Restore tabs), so they show on both tabs.
        tab_bar_extras = ttk.Frame(self.root)
        tab_bar_extras.place(in_=self.notebook, relx=1.0, x=-2, y=-1, anchor='ne')
        ttk.Checkbutton(
            tab_bar_extras, text="Show tips on startup", variable=self.show_tips_var
        ).pack(side='left', padx=(0, 10))
        ttk.Button(
            tab_bar_extras, text="Open Summary CSV", command=self._open_summary_csv_copy
        ).pack(side='left')

        self._build_backup_tab()
        self._build_restore_tab()

        self.log_frame = ttk.LabelFrame(self.root, text="Log", padding=4)
        self.log_frame.pack(fill='both', expand=False, padx=8, pady=(0, 8))
        log_header = ttk.Frame(self.log_frame)
        log_header.pack(fill='x', pady=(0, 4))
        self.log_file_path = get_new_session_log_path()
        self.log_saving_to_var = tk.StringVar(value=f"Saving to: {self.log_file_path}")
        ttk.Label(log_header, textvariable=self.log_saving_to_var, foreground='#666').pack(side='left')
        ttk.Button(log_header, text="Open Folder", command=self._open_log_folder).pack(side='right', padx=(6, 0))
        ttk.Button(log_header, text="Detach Log", command=self._detach_log).pack(side='right', padx=(6, 0))
        ttk.Button(log_header, text="Clear", command=self._clear_log).pack(side='right')

        self.log_text = tk.Text(self.log_frame, height=9, wrap='word', state='disabled', font=('Consolas', 9))
        log_vsb = ttk.Scrollbar(self.log_frame, orient='vertical', command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_vsb.set)
        self.log_text.pack(side='left', fill='both', expand=True)
        log_vsb.pack(side='left', fill='y')
        self._bind_log_copy_menu(self.log_text)
        self._configure_log_tags(self.log_text)

        self.log_detached = False
        self.detached_log_win = None
        self.detached_log_text = None

        self._write_log_file(f"=== Appx Backup Manager v{SCRIPT_VERSION} - session started {datetime.datetime.now()} ===\n")

    def _clear_log(self):
        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', 'end')
        self.log_text.configure(state='disabled')
        if self.detached_log_text is not None:
            self.detached_log_text.delete('1.0', 'end')

        # A genuinely NEW session log file, not just a cleared display - so
        # anything already written to disk from before stays exactly as it
        # was, in its own file, rather than being erased.
        self.log_file_path = get_new_session_log_path()
        self.log_saving_to_var.set(f"Saving to: {self.log_file_path}")
        self._write_log_file(f"=== Appx Backup Manager v{SCRIPT_VERSION} - session started {datetime.datetime.now()} ===\n")

    def _open_log_folder(self):
        try:
            os.startfile(os.path.dirname(self.log_file_path))
        except Exception as e:
            messagebox.showerror("Could not open folder", str(e))

    def _write_log_file(self, text):
        try:
            with open(self.log_file_path, 'a', encoding='utf-8') as f:
                f.write(text)
        except Exception:
            pass  # best-effort - never let disk logging break the UI

    def _bind_log_copy_menu(self, widget):
        def on_right_click(event):
            menu = tk.Menu(widget, tearoff=0)
            menu.add_command(label="Copy All Log Text", command=lambda: self._copy_widget_all(widget))
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()
        widget.bind('<Button-3>', on_right_click)

    def _copy_widget_all(self, widget):
        text = widget.get('1.0', 'end-1c')
        widget.clipboard_clear()
        widget.clipboard_append(text)

    def _add_log_search_bar(self, win, text_widget):
        """Adds a Ctrl+F-style find bar above a detached log window's text
        widget: highlights every match and lets you jump between them."""
        text_widget.tag_configure('search_highlight', background='#ffe066')
        text_widget.tag_configure('search_current', background='#ff9800')

        search_bar = ttk.Frame(win)
        search_bar.pack(side='top', fill='x', padx=8, pady=(8, 0))
        ttk.Label(search_bar, text="Find:").pack(side='left')
        search_var = tk.StringVar()
        entry = ttk.Entry(search_bar, textvariable=search_var)
        entry.pack(side='left', fill='x', expand=True, padx=4)
        ttk.Button(search_bar, text="✕", width=2, command=lambda: search_var.set('')).pack(side='left')
        match_var = tk.StringVar(value="")
        ttk.Label(search_bar, textvariable=match_var, foreground='#666', width=8).pack(side='left')

        state = {'matches': [], 'current': -1}

        def do_search(*_args):
            text_widget.tag_remove('search_highlight', '1.0', 'end')
            text_widget.tag_remove('search_current', '1.0', 'end')
            query = search_var.get()
            state['matches'] = []
            state['current'] = -1
            if not query:
                match_var.set("")
                return
            start = '1.0'
            while True:
                pos = text_widget.search(query, start, stopindex='end', nocase=True)
                if not pos:
                    break
                end_pos = f"{pos}+{len(query)}c"
                text_widget.tag_add('search_highlight', pos, end_pos)
                state['matches'].append((pos, end_pos))
                start = end_pos
            match_var.set(f"0/{len(state['matches'])}")
            if state['matches']:
                go_to(0)

        def go_to(index):
            if not state['matches']:
                return
            text_widget.tag_remove('search_current', '1.0', 'end')
            index = index % len(state['matches'])
            state['current'] = index
            pos, end_pos = state['matches'][index]
            text_widget.tag_add('search_current', pos, end_pos)
            text_widget.see(pos)
            match_var.set(f"{index + 1}/{len(state['matches'])}")

        def find_next(event=None):
            if not state['matches']:
                do_search()
            else:
                go_to(state['current'] + 1)
            return 'break'

        def find_prev(event=None):
            if not state['matches']:
                do_search()
            else:
                go_to(state['current'] - 1)
            return 'break'

        search_var.trace_add('write', lambda *a: do_search())
        entry.bind('<Return>', find_next)
        entry.bind('<Shift-Return>', find_prev)
        ttk.Button(search_bar, text="▲ Prev", command=find_prev).pack(side='left', padx=(4, 0))
        ttk.Button(search_bar, text="▼ Next", command=find_next).pack(side='left', padx=(2, 0))

    def _copy_log_content_with_tags(self, src_widget, dst_widget):
        """Copies a log Text widget's content to another, preserving the
        per-line color tags - a plain .get()/.insert() (what this used to
        do) only copies the text itself, silently losing all the coloring
        in the process, which is exactly what was happening on every
        detach and reattach."""
        line_count = int(src_widget.index('end-1c').split('.')[0])
        for line_num in range(1, line_count + 1):
            line_start = f"{line_num}.0"
            line_text = src_widget.get(line_start, f"{line_num}.end")
            tag = next((t for t in src_widget.tag_names(line_start) if t in LOG_LINE_TAGS), None)
            if tag:
                dst_widget.insert('end', line_text, tag)
            else:
                dst_widget.insert('end', line_text)
            if line_num < line_count:
                dst_widget.insert('end', '\n')

    def _detach_log(self):
        if self.log_detached:
            self.detached_log_win.lift()
            return
        self.log_detached = True
        self.log_frame.pack_forget()

        win = tk.Toplevel(self.root)
        win.title("Log — Appx Backup Manager")
        win.geometry("800x400")
        win.minsize(400, 200)
        win.protocol("WM_DELETE_WINDOW", self._reattach_log)

        btn_frame = ttk.Frame(win)
        btn_frame.pack(side='bottom', fill='x', pady=6, padx=8)
        ttk.Button(btn_frame, text="Reattach", command=self._reattach_log).pack(side='right')

        text_frame = ttk.Frame(win)
        new_text = tk.Text(text_frame, wrap='word', font=('Consolas', 9))
        vsb = ttk.Scrollbar(text_frame, orient='vertical', command=new_text.yview)
        new_text.configure(yscrollcommand=vsb.set)
        new_text.pack(side='left', fill='both', expand=True)
        vsb.pack(side='left', fill='y')
        self._configure_log_tags(new_text)
        self._copy_log_content_with_tags(self.log_text, new_text)
        new_text.see('end')
        self._bind_log_copy_menu(new_text)
        # Packed (search bar) before text_frame (expand=True) below,
        # deliberately - packing text_frame first would let it greedily
        # claim all remaining space immediately, leaving the search bar
        # packed afterward squeezed into whatever sliver was left over.
        self._add_log_search_bar(win, new_text)
        text_frame.pack(fill='both', expand=True, padx=8, pady=(8, 0))

        self.detached_log_win = win
        self.detached_log_text = new_text

    def _reattach_log(self):
        if not self.log_detached:
            return
        if self.detached_log_win:
            self.log_text.configure(state='normal')
            self.log_text.delete('1.0', 'end')
            if self.detached_log_text:
                self._copy_log_content_with_tags(self.detached_log_text, self.log_text)
            self.detached_log_win.destroy()
        self.detached_log_win = None
        self.detached_log_text = None
        self.log_detached = False

        self.log_text.see('end')
        self.log_text.configure(state='disabled')
        self.log_frame.pack(fill='both', expand=False, padx=8, pady=(0, 8))

    def _open_summary_csv_copy(self):
        """Opens a timestamped COPY of backup-summary.csv (in %TEMP%), so
        viewing or editing it in Excel can never lock or change the real
        file the backup/restore scripts read and write."""
        backup_root = self.folder_var.get().strip()
        src = os.path.join(backup_root, 'backup-summary.csv') if backup_root else ''
        if not src or not os.path.isfile(src):
            messagebox.showinfo(
                "Open Summary CSV",
                f"No backup-summary.csv found in:\n{backup_root or '(no folder selected)'}"
            )
            return
        try:
            dest_dir = os.path.join(tempfile.gettempdir(), 'AppxBackupSummaryCopies')
            os.makedirs(dest_dir, exist_ok=True)
            stamp = datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S')
            dest = os.path.join(dest_dir, f"backup-summary (copy {stamp}).csv")
            shutil.copy2(src, dest)
            os.startfile(dest)
            self.log_queue.put(f"Opened a copy of the summary (changes to it don't affect the backup): {dest}\n")
        except Exception as e:
            messagebox.showerror("Open Summary CSV", f"Could not open a copy of backup-summary.csv:\n{e}")

    def _build_backup_tab(self):
        tab = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(tab, text="Backup")

        controls = ttk.Frame(tab)
        controls.pack(fill='x', pady=(0, 4))
        ttk.Button(controls, text="Refresh List", command=self.refresh_installed_apps).pack(side='left')
        ttk.Label(controls, text="Show:").pack(side='left', padx=(12, 0))
        self.backup_category_var = tk.StringVar(value="Apps")
        backup_category_combo = ttk.Combobox(
            controls, textvariable=self.backup_category_var,
            values=["All", "Apps", "Frameworks"], state='readonly', width=15
        )
        backup_category_combo.pack(side='left', padx=(4, 12))
        backup_category_combo.bind('<<ComboboxSelected>>', lambda e: self._on_backup_category_changed())
        self.all_users_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(controls, text="Include all users (-AllUsers)", variable=self.all_users_var).pack(side='left')
        self.only_start_apps_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            controls, text="Only show Start Menu apps", variable=self.only_start_apps_var,
            command=self._on_backup_category_changed
        ).pack(side='left', padx=(12, 0))
        self.backup_status_var = tk.StringVar(value="Loading installed apps...")
        ttk.Label(controls, textvariable=self.backup_status_var).pack(side='right')

        force_row = ttk.Frame(tab)
        force_row.pack(fill='x', pady=(0, 4))
        self.force_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            force_row, text="Force re-backup", variable=self.force_var
        ).pack(side='left')
        self.show_tools_frame_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            force_row, text="Show packing/signing tools", variable=self.show_tools_frame_var,
            command=self._toggle_tools_frame
        ).pack(side='left', padx=(16, 0))

        self.tools_frame = tools_frame = ttk.LabelFrame(tab, text="Packing / signing tools (optional - needed for portable .appx creation)", padding=6)
        tools_frame.pack(fill='x', pady=(0, 6))

        # Persisted settings take priority over auto-detection - once set
        # (typed, browsed to, or auto-detected and then implicitly kept),
        # a custom Windows SDK install location only needs pointing at
        # once, not every session. trace_add saves on every change so
        # there's no separate "Save" step to remember.
        saved_settings = load_settings()

        self.makeappx_path_var = tk.StringVar(value=saved_settings.get('makeappx_path') or find_sdk_tool('makeappx.exe') or '')
        # Plain Python string, kept in sync by the trace below (which only
        # ever fires on the main thread, since it's driven by a UI event) -
        # _build_backup_cmd reads THIS, not the StringVar directly, because
        # it can be called from _run_commands_worker's background thread,
        # and calling .get() on a Tkinter variable from a non-main thread
        # is unsafe (Tcl's interpreter state isn't thread-safe the way
        # plain Python attribute access is).
        self._makeappx_path = self.makeappx_path_var.get().strip()

        def _on_makeappx_path_changed(*_a):
            value = self.makeappx_path_var.get().strip()
            self._makeappx_path = value
            save_setting('makeappx_path', value)
        self.makeappx_path_var.trace_add('write', _on_makeappx_path_changed)
        ttk.Label(tools_frame, text="MakeAppx.exe:").grid(row=0, column=0, sticky='w')
        ttk.Entry(tools_frame, textvariable=self.makeappx_path_var).grid(row=0, column=1, sticky='ew', padx=4)
        ttk.Button(tools_frame, text="Browse…", command=lambda: self._browse_tool(self.makeappx_path_var, 'makeappx.exe')).grid(row=0, column=2)
        ttk.Button(tools_frame, text="Test", command=lambda: self._test_tool(self.makeappx_path_var, 'MakeAppx.exe')).grid(row=0, column=3, padx=(4, 0))

        self.signtool_path_var = tk.StringVar(value=saved_settings.get('signtool_path') or find_sdk_tool('signtool.exe') or '')
        self._signtool_path = self.signtool_path_var.get().strip()  # see the makeappx equivalent above for why

        def _on_signtool_path_changed(*_a):
            value = self.signtool_path_var.get().strip()
            self._signtool_path = value
            save_setting('signtool_path', value)
        self.signtool_path_var.trace_add('write', _on_signtool_path_changed)
        ttk.Label(tools_frame, text="SignTool.exe:").grid(row=1, column=0, sticky='w', pady=(4, 0))
        ttk.Entry(tools_frame, textvariable=self.signtool_path_var).grid(row=1, column=1, sticky='ew', padx=4, pady=(4, 0))
        ttk.Button(tools_frame, text="Browse…", command=lambda: self._browse_tool(self.signtool_path_var, 'signtool.exe')).grid(row=1, column=2, pady=(4, 0))
        ttk.Button(tools_frame, text="Test", command=lambda: self._test_tool(self.signtool_path_var, 'SignTool.exe')).grid(row=1, column=3, padx=(4, 0), pady=(4, 0))
        tools_frame.columnconfigure(1, weight=1)

        ttk.Label(
            tools_frame,
            text="Leave blank to auto-detect (checks PATH, then the default Windows SDK install folder). "
                 "If left blank and not found, backup still works - it just skips packing/signing and does a "
                 "raw-folder-only, same-PC-restore backup instead.",
            wraplength=700, justify='left', foreground='#555'
        ).grid(row=2, column=0, columnspan=4, sticky='w', pady=(4, 0))

        columns = [
            ('Name', 'Name', 260, 'w', True),
            ('Version', 'Version', 90, 'w', False),
            ('Publisher', 'Publisher', 260, 'w', True),
            ('size_display', 'Size', 90, 'e', False),
        ]
        self.backup_list = CheckableAppList(
            tab, columns, key_field='PackageFullName', on_change=self._update_backup_total,
            single_filter_row=True
        )
        self.backup_list.pack(fill='both', expand=True)
        IconTooltip(self.backup_list.tree, self._get_backup_row_icon_bytes)
        self.backup_list.tree.bind('<Button-3>', self._on_backup_right_click)

        total_row = ttk.Frame(tab)
        total_row.pack(fill='x', pady=(4, 4))
        self.backup_total_var = tk.StringVar(value="Selected: 0 app(s), ~0 B total")
        ttk.Label(total_row, textvariable=self.backup_total_var).pack(side='left')

        self.backup_mode_var = tk.StringVar(value='both')
        ttk.Radiobutton(
            total_row, text="Raw folder only", variable=self.backup_mode_var, value='raw'
        ).pack(side='left', padx=(16, 0))
        ttk.Radiobutton(
            total_row, text="Packed only", variable=self.backup_mode_var, value='packed'
        ).pack(side='left', padx=(8, 0))
        ttk.Radiobutton(
            total_row, text="Both", variable=self.backup_mode_var, value='both'
        ).pack(side='left', padx=(8, 0))

        ttk.Button(total_row, text="Start Backup", command=self.start_backup).pack(side='right')
        ttk.Button(total_row, text="Cancel", command=self.request_cancel_run).pack(side='right', padx=(0, 8))

    def _get_backup_row_icon_bytes(self, row_id):
        row = self.backup_list.get_row_by_id(row_id)
        if not row:
            return None
        install_location = row.get('InstallLocation')
        if install_location and os.path.isdir(install_location):
            return find_folder_icon_bytes(install_location)
        return None

    def _get_restore_row_icon_bytes(self, row_id):
        row = self.restore_list.get_row_by_id(row_id)
        if not row:
            return None
        full_name = row.get('PackageFullName')
        backup_root = self.folder_var.get().strip()
        if not full_name or not backup_root:
            return None
        raw_dir = os.path.join(backup_root, 'Raw', full_name)
        if os.path.isdir(raw_dir):
            icon = find_folder_icon_bytes(raw_dir)
            if icon:
                return icon
        packed_path = os.path.join(backup_root, 'Packed', f'{full_name}.appx')
        if os.path.isfile(packed_path):
            return find_manifest_icon_bytes(packed_path)
        return None

    def _get_restore_row_full_name(self, row_id):
        row = self.restore_list.get_row_by_id(row_id)
        return row.get('PackageFullName') if row else None

    def _browse_tool(self, path_var, default_name):
        path = filedialog.askopenfilename(
            title=f"Locate {default_name}",
            filetypes=[(default_name, default_name), ("Executables", "*.exe"), ("All files", "*.*")]
        )
        if path:
            path_var.set(path)

    def _toggle_tools_frame(self):
        """Shows or hides the packing/signing tools box - it's optional
        (backup works fine without it, just skipping packing/signing), so
        hiding it frees up vertical space for the app list once its paths
        are set correctly and don't need frequent attention. before=
        keeps it in its original position when re-shown, since pack()
        would otherwise just re-append it to the end (after the app
        list) instead of putting it back where it was."""
        if self.show_tools_frame_var.get():
            self.tools_frame.pack(fill='x', pady=(0, 6), before=self.backup_list)
        else:
            self.tools_frame.pack_forget()

    def _test_tool(self, path_var, label):
        path = path_var.get().strip() or find_sdk_tool(label.lower())
        ok, message = test_sdk_tool(path)
        self.log_queue.put(f"\n--- [{log_ts()}] Tested {label}: {'OK' if ok else 'WARNING'} - {message} ---\n")
        if ok:
            messagebox.showinfo(f"{label} check", message)
        else:
            messagebox.showwarning(f"{label} check", message)

    def _build_restore_tab(self):
        tab = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(tab, text="Restore")

        # Both filter rows live in one container: side by side when the
        # window is wide enough, stacked otherwise (see _layout_restore_filters).
        filters = ttk.Frame(tab)
        filters.pack(fill='x', pady=(0, 4))
        controls = ttk.Frame(filters)
        ttk.Button(controls, text="Refresh List", command=self.load_restore_list).pack(side='left')
        ttk.Label(controls, text="Show:").pack(side='left', padx=(12, 4))
        self.restore_category_var = tk.StringVar(value="Apps")
        restore_category_combo = ttk.Combobox(
            controls, textvariable=self.restore_category_var,
            values=["All", "Apps", "Frameworks", "Dependencies"], state='readonly', width=18
        )
        restore_category_combo.pack(side='left')
        restore_category_combo.bind('<<ComboboxSelected>>', self._on_restore_category_changed)
        self.restore_all_users_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            controls, text="Check all users' installs (-AllUsers)", variable=self.restore_all_users_var
        ).pack(side='left', padx=8)
        self.restore_status_var = tk.StringVar(
            value="Choose a folder above - if it already has a backup, its contents load automatically."
        )
        ttk.Label(controls, textvariable=self.restore_status_var).pack(side='right')

        controls2 = ttk.Frame(filters)
        ttk.Label(controls2, text="Select:").pack(side='left')
        self.restore_select_var = tk.StringVar(value="Apps")
        restore_select_combo = ttk.Combobox(
            controls2, textvariable=self.restore_select_var,
            values=["Apps", "Dependencies", "Frameworks", "Installed", "Not Installed",
                    "Raw Backup", "Packed Backup", "Missing"],
            state='readonly', width=15
        )
        restore_select_combo.pack(side='left', padx=4)
        # Deliberately not bound to <<ComboboxSelected>> - picking a value
        # here overwrites your own manual checkbox selections in the list,
        # and applying that the instant the dropdown closes (with no
        # visible sign afterward that anything happened) made it look like
        # the picked value just sits there unapplied. "Confirm" makes the
        # action an explicit, visible step instead.
        ttk.Button(controls2, text="Confirm", command=self._on_select_combo_changed).pack(side='left', padx=(4, 0))

        self.check_files_locked_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            controls2, text="Check files aren't locked before modifying them", variable=self.check_files_locked_var
        ).pack(side='right')
        special_actions_btn = ttk.Menubutton(controls2, text="Special Actions")
        special_actions_menu = tk.Menu(special_actions_btn, tearoff=0)
        special_actions_menu.add_command(label="Clean Up Orphaned Frameworks…", command=self._cleanup_orphaned_dependencies)
        special_actions_menu.add_command(label="Clean Up Orphaned Dependencies…", command=self._cleanup_orphaned_non_framework_dependencies)
        special_actions_menu.add_separator()
        special_actions_menu.add_command(label="Create Packed Backups from Raw…", command=self._create_packed_from_raw)
        special_actions_menu.add_command(label="Create Raw Backups from Packed…", command=self._create_raw_from_packed)
        special_actions_btn.configure(menu=special_actions_menu)
        special_actions_btn.pack(side='right', padx=(0, 16))

        self._restore_filters = (filters, controls, controls2)
        self._restore_filters_one_row = None
        filters.bind('<Configure>', lambda e: self._layout_restore_filters())
        self.restore_status_var.trace_add('write', lambda *a: self.root.after_idle(self._layout_restore_filters))
        self._layout_restore_filters()
        # (Export Summary and Verify All Backups moved to the same row as
        # Start Restore below, opposite corner)

        columns = [
            ('Name', 'Name', 170, 'w', True),
            ('PackageFullName', 'Package Full Name', 320, 'w', True),
            ('Version', 'Version', 90, 'w', False),
            ('Publisher', 'Publisher', 170, 'w', True),
            ('Type', 'Type', 90, 'center', False),
            ('InstalledStatus', 'Installed', 80, 'center', False),
            ('size_display', 'Size', 90, 'e', False),
            ('BackupType', 'Backup', 100, 'center', False),
            ('SignedOK', 'Signed', 70, 'center', False),
        ]
        self.restore_list = CheckableAppList(
            tab, columns, key_field='PackageFullName', on_change=self._update_restore_total,
            column_options={
                'Type': ['App', 'Framework', 'Dependency'],
                'InstalledStatus': ['Yes', 'No'],
                'BackupType': ['Both', 'Raw-Only', 'Packed-Only', 'Missing'],
                'SignedOK': ['True', 'False', 'NA'],
            }
        )
        self.restore_list.pack(fill='both', expand=True)
        self.restore_list.tree.bind('<Button-3>', self._on_restore_right_click)
        IconTooltip(self.restore_list.tree, self._get_restore_row_icon_bytes, self._get_restore_row_full_name)

        options = ttk.Frame(tab)
        options.pack(fill='x', pady=(4, 4))
        self.restore_total_var = tk.StringVar(value="Selected: 0 app(s), ~0 B total")
        ttk.Label(options, textvariable=self.restore_total_var).pack(side='left')

        mode_frame = ttk.Frame(options)
        mode_frame.pack(side='right')
        self.restore_mode_var = tk.StringVar(value='Install')
        ttk.Radiobutton(
            mode_frame, text="Register (same PC)", variable=self.restore_mode_var, value='Register'
        ).pack(side='left')
        ttk.Radiobutton(
            mode_frame, text="Install (this or another PC)", variable=self.restore_mode_var, value='Install'
        ).pack(side='left', padx=12)
        self.enable_sideloading_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            mode_frame, text="Enable sideloading (Install mode)", variable=self.enable_sideloading_var
        ).pack(side='left', padx=(12, 0))

        bottom_row = ttk.Frame(tab)
        bottom_row.pack(fill='x', pady=(4, 0))
        ttk.Button(bottom_row, text="Export Summary…", command=self._export_backup_summary).pack(side='left')
        ttk.Button(
            bottom_row, text="Open Last Summary", command=self._open_last_summary
        ).pack(side='left', padx=(8, 0))
        ttk.Button(
            bottom_row, text="Verify All Backups…", command=self._verify_all_backups
        ).pack(side='left', padx=(16, 0))
        ttk.Button(
            bottom_row, text="Open Last Full Report", command=self._open_last_verify_report
        ).pack(side='left', padx=(8, 0))
        ttk.Button(bottom_row, text="Restore", command=self.start_restore).pack(side='right')
        ttk.Button(bottom_row, text="Cancel", command=self.request_cancel_run).pack(side='right', padx=(0, 8))
        ttk.Button(bottom_row, text="Remove", command=self._remove_selected_backups).pack(side='right', padx=(0, 8))
        self.remove_mode_var = tk.StringVar(value='Both')
        ttk.Combobox(
            bottom_row, textvariable=self.remove_mode_var, values=['Both', 'Raw', 'Packed'],
            state='readonly', width=9
        ).pack(side='right', padx=(0, 4))

    def _layout_restore_filters(self):
        """Puts the Restore tab's two filter rows on one line when they fit
        the current width, or stacks them when they don't."""
        filters, row1, row2 = self._restore_filters
        width = filters.winfo_width()
        one_row = width > 1 and (row1.winfo_reqwidth() + row2.winfo_reqwidth() + 16) <= width
        if one_row == self._restore_filters_one_row:
            return
        self._restore_filters_one_row = one_row
        row1.grid_forget()
        row2.grid_forget()
        filters.columnconfigure(0, weight=1)
        filters.columnconfigure(1, weight=0)
        if one_row:
            row1.grid(row=0, column=0, sticky='ew')
            row2.grid(row=0, column=1, sticky='ew', padx=(16, 0))
        else:
            row1.grid(row=0, column=0, columnspan=2, sticky='ew')
            row2.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(4, 0))

    def _open_last_summary(self):
        if not self.last_summary_path or not os.path.isfile(self.last_summary_path):
            messagebox.showinfo("No summary yet", "Export a backup summary first (\"Export Summary…\").")
            return
        try:
            os.startfile(self.last_summary_path)
        except OSError as e:
            messagebox.showerror("Could not open", str(e))

    def _open_last_verify_report(self):
        if not self.last_verify_report_path or not os.path.isfile(self.last_verify_report_path):
            messagebox.showinfo(
                "No report yet",
                "Run \"Verify All Backups…\" first, then use \"Export Full Report…\" in the results window."
            )
            return
        try:
            os.startfile(self.last_verify_report_path)
        except OSError as e:
            messagebox.showerror("Could not open", str(e))

    def _select_not_installed(self):
        matched = 0
        for row_id, key in self.restore_list.row_id_to_key.items():
            row = next((r for r in self.restore_list.all_rows if r['PackageFullName'] == key), None)
            if row is None:
                continue
            installed = row.get('InstalledStatus') == 'Yes'
            self.restore_list.tree.set(row_id, 'select', '☐' if installed else '☑')
            if installed:
                self.restore_list.checked.discard(key)
            else:
                self.restore_list.checked.add(key)
                matched += 1
        self.restore_list._apply_filter()
        if matched == 0:
            messagebox.showinfo(
                "Nothing to select",
                "No not-installed items are currently visible."
            )

    def _select_already_installed(self):
        matched = 0
        for row_id, key in self.restore_list.row_id_to_key.items():
            row = next((r for r in self.restore_list.all_rows if r['PackageFullName'] == key), None)
            if row is None:
                continue
            installed = row.get('InstalledStatus') == 'Yes'
            self.restore_list.tree.set(row_id, 'select', '☑' if installed else '☐')
            if installed:
                self.restore_list.checked.add(key)
                matched += 1
            else:
                self.restore_list.checked.discard(key)
        self.restore_list._apply_filter()
        if matched == 0:
            messagebox.showinfo(
                "Nothing to select",
                "No already-installed items are currently visible."
            )

    def _on_select_combo_changed(self, event=None):
        """Dispatches to the matching selection action for whichever
        option is picked - consolidates what used to be several separate
        buttons into one combobox, applied only when "Confirm" is
        actually clicked (not the instant the dropdown closes)."""
        action = self.restore_select_var.get()
        dispatch = {
            "Apps": self._select_direct_only,
            "Frameworks": self._select_frameworks_only,
            "Dependencies": self._select_dependencies_only,
            "Not Installed": self._select_not_installed,
            "Installed": self._select_already_installed,
            "Raw Backup": self._select_raw_backup_only,
            "Packed Backup": self._select_packed_backup_only,
            "Missing": self._select_missing_backup,
        }
        fn = dispatch.get(action)
        if fn:
            fn()

    def _select_direct_only(self):
        matched = 0
        for row_id, key in self.restore_list.row_id_to_key.items():
            row = next((r for r in self.restore_list.all_rows if r['PackageFullName'] == key), None)
            if row is None:
                continue
            match = row.get('Type') == 'App'
            self.restore_list.tree.set(row_id, 'select', '☑' if match else '☐')
            if match:
                self.restore_list.checked.add(key)
                matched += 1
            else:
                self.restore_list.checked.discard(key)
        self.restore_list._apply_filter()
        if matched == 0:
            messagebox.showinfo(
                "Nothing to select",
                "No apps are currently visible."
            )

    def _select_dependencies_only(self):
        """Non-framework dependencies only - frameworks have their own
        separate "Frameworks" option now, so this no longer lumps
        them in together."""
        matched = 0
        for row_id, key in self.restore_list.row_id_to_key.items():
            row = next((r for r in self.restore_list.all_rows if r['PackageFullName'] == key), None)
            if row is None:
                continue
            match = row.get('Type') == 'Dependency'
            self.restore_list.tree.set(row_id, 'select', '☑' if match else '☐')
            if match:
                self.restore_list.checked.add(key)
                matched += 1
            else:
                self.restore_list.checked.discard(key)
        self.restore_list._apply_filter()
        if matched == 0:
            messagebox.showinfo(
                "Nothing to select",
                "No non-framework dependencies are currently visible."
            )

    def _select_frameworks_only(self):
        matched = 0
        for row_id, key in self.restore_list.row_id_to_key.items():
            row = next((r for r in self.restore_list.all_rows if r['PackageFullName'] == key), None)
            if row is None:
                continue
            match = row.get('Type') == 'Framework'
            self.restore_list.tree.set(row_id, 'select', '☑' if match else '☐')
            if match:
                self.restore_list.checked.add(key)
                matched += 1
            else:
                self.restore_list.checked.discard(key)
        self.restore_list._apply_filter()
        if matched == 0:
            messagebox.showinfo(
                "Nothing to select",
                "No frameworks are currently visible."
            )

    def _select_by_backup_type(self, backup_type, none_message):
        """Shared body for the three "Backup" status selections below -
        same logic each time, just matching a different BackupType value
        and a different empty-list message."""
        matched = 0
        for row_id, key in self.restore_list.row_id_to_key.items():
            row = next((r for r in self.restore_list.all_rows if r['PackageFullName'] == key), None)
            if row is None:
                continue
            match = row.get('BackupType') == backup_type
            self.restore_list.tree.set(row_id, 'select', '☑' if match else '☐')
            if match:
                self.restore_list.checked.add(key)
                matched += 1
            else:
                self.restore_list.checked.discard(key)
        self.restore_list._apply_filter()
        if matched == 0:
            messagebox.showinfo("Nothing to select", none_message)

    def _select_raw_backup_only(self):
        self._select_by_backup_type('Raw-Only', "No raw-only backups are currently visible.")

    def _select_packed_backup_only(self):
        self._select_by_backup_type('Packed-Only', "No packed-only backups are currently visible.")

    def _select_missing_backup(self):
        self._select_by_backup_type('Missing', "No missing backups are currently visible.")

    def _show_text_dialog(self, title, text, geometry="700x560"):
        win = tk.Toplevel(self.root)
        win.title(title)
        width, height = (int(n) for n in geometry.split('x'))
        self._center_window(win, width=width, height=height)
        win.minsize(400, 300)
        ttk.Button(win, text="Close", command=win.destroy).pack(side='bottom', pady=6)
        text_frame = ttk.Frame(win)
        text_frame.pack(side='top', fill='both', expand=True, padx=8, pady=(8, 0))
        text_widget = tk.Text(text_frame, wrap='word')
        vsb = ttk.Scrollbar(text_frame, orient='vertical', command=text_widget.yview)
        text_widget.configure(yscrollcommand=vsb.set)
        text_widget.pack(side='left', fill='both', expand=True)
        vsb.pack(side='left', fill='y')
        text_widget.insert('1.0', text)
        text_widget.configure(state='disabled')

    def _maybe_show_tip_dialog(self):
        """Shown once, shortly after the main window itself has already
        rendered (scheduled via root.after, not called directly from
        __init__) - a startup "tip of the day" rather than anything tied
        to a specific backup/restore run, so it draws from every category
        in tips.json pooled together rather than just one. Silently does
        nothing if there are no tips to show at all, or if disabled."""
        if not self.show_tips_var.get():
            return
        all_tips = [t for tips_list in self.tips.values() for t in tips_list]
        if not all_tips:
            return
        tip = random.choice(all_tips)

        win = tk.Toplevel(self.root)
        win.title("Tip")
        win.resizable(False, False)
        win.transient(self.root)

        ttk.Label(
            win, text=tip, wraplength=380, justify='left', padding=(16, 16, 16, 8)
        ).pack()

        # Bound directly to the same shared variable as the two checkboxes
        # elsewhere (Backup tab, Restore tab) - the trace on it already
        # persists any change, from wherever it happens, so there's
        # nothing extra to sync or save here.
        ttk.Checkbutton(
            win, text="Show tips when this app starts", variable=self.show_tips_var
        ).pack(anchor='w', padx=16, pady=(0, 8))
        ttk.Button(win, text="OK", command=win.destroy).pack(pady=(0, 16))
        win.protocol("WM_DELETE_WINDOW", win.destroy)
        self._center_window(win)

    def _show_help(self):
        self._show_text_dialog("Help", HELP_TEXT)

    def _on_restore_right_click(self, event):
        row_id = self.restore_list.tree.identify_row(event.y)
        col = self.restore_list.tree.identify_column(event.x)
        values = self.restore_list.tree.item(row_id, 'values') if row_id else ()
        col_index = None
        if col and col.startswith('#'):
            try:
                col_index = int(col[1:]) - 1
            except ValueError:
                pass
        cell_value = values[col_index] if col_index is not None and 0 <= col_index < len(values) else ''

        row = None
        if row_id in self.restore_list.row_id_to_key:
            self.restore_list.tree.selection_set(row_id)
            key = self.restore_list.row_id_to_key[row_id]
            row = next((r for r in self.restore_list.all_rows if r['PackageFullName'] == key), None)

        menu = tk.Menu(self.restore_list.tree, tearoff=0)
        if row is not None:
            menu.add_command(label="Show Backed-Up Files…", command=lambda: self.show_app_files(row))
            menu.add_command(label="Verify Files (checklist + hashes)…", command=lambda: self.show_app_file_verifier(row))
            menu.add_command(label="Open Raw Folder", command=lambda: self._open_restore_raw_folder(row))
            menu.add_command(label="Find Packed File", command=lambda: self._find_restore_packed_file(row))
            menu.add_separator()
            menu.add_command(label="Show Dependencies…", command=lambda: self._show_dependencies_dialog(row))
            menu.add_command(label="Show Which Apps Use This…", command=lambda: self._show_used_by_dialog(row))
            menu.add_command(label="Remove Backup…", command=lambda: self._remove_backup_for_app(row))
            menu.add_separator()
        if row_id:
            menu.add_command(label="Copy Cell", command=lambda: self.restore_list._copy_to_clipboard(str(cell_value)))
            menu.add_command(label="Copy Row (all columns)", command=lambda: self.restore_list._copy_to_clipboard('\t'.join(str(v) for v in values)))
        menu.add_command(label="Copy All Visible Rows", command=self.restore_list.copy_all_visible_rows)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _open_restore_raw_folder(self, row):
        backup_root = self.folder_var.get().strip()
        if not backup_root:
            messagebox.showinfo("Folder needed", "Choose a backup/restore folder first.")
            return
        full_name = row.get('PackageFullName')
        raw_dir = os.path.join(backup_root, 'Raw', full_name)
        if not os.path.isdir(raw_dir):
            messagebox.showwarning("Not found", f"No Raw backup folder found for:\n{full_name}")
            return
        self._open_folder_in_explorer(raw_dir)

    def _find_restore_packed_file(self, row):
        """Opens the Packed\\ folder with this app's .appx highlighted -
        useful for grabbing the file directly (to copy elsewhere, attach to
        an email, etc.) without digging through the folder yourself."""
        backup_root = self.folder_var.get().strip()
        if not backup_root:
            messagebox.showinfo("Folder needed", "Choose a backup/restore folder first.")
            return
        full_name = row.get('PackageFullName')
        packed_path = os.path.join(backup_root, 'Packed', f'{full_name}.appx')
        if not os.path.isfile(packed_path):
            messagebox.showwarning("Not found", f"No packed .appx found for:\n{full_name}")
            return
        self._show_in_explorer(packed_path)

    def _read_backup_summary_csv(self, backup_root):
        """Returns (rows_by_package_full_name, fieldnames). Empty dict/list
        if the file doesn't exist yet."""
        summary_path = os.path.join(backup_root, 'backup-summary.csv')
        if not os.path.isfile(summary_path):
            return {}, []
        with open(summary_path, newline='', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            fieldnames = reader.fieldnames or []
            rows = {row['PackageFullName']: row for row in reader if row.get('PackageFullName')}
        return rows, fieldnames

    def _write_backup_summary_csv(self, backup_root, rows, fieldnames):
        summary_path = os.path.join(backup_root, 'backup-summary.csv')
        with open(summary_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for row in rows.values():
                writer.writerow(row)

    def _prepare_and_run(self, backup_root, full_names, action_label, lock_description, on_ready, on_cancel=None):
        """Runs the (potentially slow, since it has to touch every relevant
        file on disk) locked-files scan in a background thread - never on
        the caller's thread - so a bulk operation over many apps/files
        never freezes the UI while it's checking, exactly the problem
        the removal operations were made async to avoid in the first
        place. When the scan is done (or immediately, if "Check files
        aren't locked" is off - no thread is spawned at all in that
        case, so there's zero added time), handles any needed
        confirmation and lock acquisition on the main thread, then calls
        on_ready() to actually start the real operation. on_cancel, if
        given, is called instead when the user declines to proceed."""
        if not self.check_files_locked_var.get():
            on_ready()
            return
        threading.Thread(
            target=self._locked_files_scan_worker,
            args=(backup_root, full_names, action_label, lock_description, on_ready, on_cancel),
            daemon=True
        ).start()

    def _locked_files_scan_worker(self, backup_root, full_names, action_label, lock_description, on_ready, on_cancel):
        try:
            paths = gather_backup_files_for(backup_root, full_names)
            locked = find_locked_files(paths) if paths else []
        except Exception:
            # This pre-check is a safety net, not the real operation -
            # never let a problem in the check itself block the actual
            # removal/restore from being attempted.
            locked = []
        self.root.after(0, lambda: self._on_locked_files_scanned(
            backup_root, locked, action_label, lock_description, on_ready, on_cancel
        ))

    def _on_locked_files_scanned(self, backup_root, locked, action_label, lock_description, on_ready, on_cancel):
        if locked:
            preview = '\n'.join(f"  {p}" for p in locked[:10])
            more = f"\n  ...and {len(locked) - 10} more" if len(locked) > 10 else ""
            if not messagebox.askyesno(
                "Some files appear locked",
                f"{len(locked)} file(s) look like they might be open or locked by another "
                f"program (or are marked read-only) - {action_label} could fail partway "
                f"through:\n\n{preview}{more}\n\nProceed anyway?"
            ):
                if on_cancel:
                    on_cancel()
                return
        if not self._try_acquire_backup_lock(backup_root, lock_description):
            if on_cancel:
                on_cancel()
            return
        on_ready()

    def _try_acquire_backup_lock(self, backup_root, description):
        """Writes a small marker file recording that an operation is in
        progress against this backup folder - a lightweight, advisory
        lock against a genuinely concurrent operation from elsewhere
        (most plausibly another window of this same app pointed at the
        same folder), not an OS-level file lock: this tool's own
        subprocess-based file access (robocopy, MakeAppx, PowerShell)
        still needs unrestricted access to these same files, so actually
        holding them open here would conflict with the very operations
        this is meant to protect. A lock older than 6 hours is treated as
        left behind by a crash rather than a live operation, so it never
        blocks forever. Returns True to proceed (lock acquired or the
        option is off), False if the user chose to cancel after seeing a
        live lock. Safe to call repeatedly - a failure to write the
        marker itself is non-fatal, since this is a safety net, not a
        hard requirement."""
        if not self.check_files_locked_var.get():
            return True
        lock_path = os.path.join(backup_root, LOCK_MARKER_NAME)
        if os.path.isfile(lock_path):
            try:
                age_seconds = time.time() - os.path.getmtime(lock_path)
            except OSError:
                age_seconds = 0
            if age_seconds < 6 * 3600:
                try:
                    with open(lock_path, encoding='utf-8') as f:
                        existing = f.read().strip()
                except OSError:
                    existing = ""
                age_minutes = int(age_seconds // 60)
                if not messagebox.askyesno(
                    "Backup folder appears busy",
                    f"This backup folder already has a lock from {existing or 'another operation'}, "
                    f"started {age_minutes} minute(s) ago - it may still be running, possibly in "
                    f"another window of this app.\n\nProceed anyway?"
                ):
                    return False
        try:
            with open(lock_path, 'w', encoding='utf-8') as f:
                f.write(f"{description} (PID {os.getpid()}) at {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        except OSError:
            pass
        return True

    def _release_backup_lock(self, backup_root):
        try:
            lock_path = os.path.join(backup_root, LOCK_MARKER_NAME)
            if os.path.isfile(lock_path):
                os.remove(lock_path)
        except OSError:
            pass

    def _delete_backup_files_for(self, backup_root, full_name, remove_mode='Both'):
        """Just the raw deletion for one package. No confirmation, no CSV
        update (the caller handles that once for the whole batch).
        Returns a list of error strings; empty means fully successful.

        remove_mode controls what actually gets deleted:
          'Both'   - Raw folder, Packed .appx, and the Manifests JSON
                     (which records Raw's own file hashes, so it's
                     meaningless once Raw itself is gone) - the
                     original, full removal.
          'Raw'    - just the Raw folder (and its Manifests JSON, for
                     the same reason) - keeps the packed .appx.
          'Packed' - just the Packed .appx - keeps Raw and its
                     Manifests JSON, since both are still valid."""
        raw_dir = os.path.join(backup_root, 'Raw', full_name)
        packed_path = os.path.join(backup_root, 'Packed', f'{full_name}.appx')
        manifest_path = os.path.join(backup_root, 'Manifests', f'{full_name}.json')
        errors = []
        if remove_mode in ('Both', 'Raw') and os.path.isdir(raw_dir):
            try:
                shutil.rmtree(raw_dir)
            except OSError as e:
                errors.append(f"{full_name} - Raw folder: {e}")
        if remove_mode in ('Both', 'Packed') and os.path.isfile(packed_path):
            try:
                os.remove(packed_path)
            except OSError as e:
                errors.append(f"{full_name} - Packed .appx: {e}")
        if remove_mode in ('Both', 'Raw') and os.path.isfile(manifest_path):
            try:
                os.remove(manifest_path)
            except OSError as e:
                errors.append(f"{full_name} - Manifests JSON: {e}")
        return errors

    def _verify_all_backups(self):
        """Bulk sanity check: for every app in this backup with a recorded
        manifest, re-hashes its Raw\\ folder's files and compares against
        what was recorded at backup time - independent of whether the app
        is even still installed. Runs in the background since hashing
        everything can take a while for a large backup; logs a start line
        and a one-line summary when done (not the full per-file detail,
        to keep the main log from filling up) - the results dialog itself
        still shows everything, and can export the full detail to a file
        on demand ("Export Full Report…")."""
        if self._backup_folder_busy:
            messagebox.showinfo("Busy", "A removal or verification is already in progress - please wait for it to finish first.")
            return
        backup_root = self.folder_var.get().strip()
        if not backup_root:
            messagebox.showinfo("Folder needed", "Choose a backup/restore folder first.")
            return
        rows, _ = self._read_backup_summary_csv(backup_root)
        if not rows:
            messagebox.showinfo("Nothing to verify", "No backup is currently loaded here.")
            return

        self._backup_folder_busy = True
        self.log_queue.put(f"\n--- [{log_ts()}] Verifying {len(rows)} package(s) against their recorded manifests... ---\n")
        self.restore_status_var.set(f"Verifying {len(rows)} package(s) against their recorded manifests...")
        threading.Thread(target=self._verify_all_backups_worker, args=(backup_root, list(rows.keys())), daemon=True).start()

    def _verify_all_backups_worker(self, backup_root, full_names):
        results = []  # (name, status, detail)
        for full_name in full_names:
            try:
                manifest = get_backup_manifest(backup_root, full_name)
                if manifest is None:
                    results.append((full_name, 'NO MANIFEST', 'made with an older version of this tool, or Raw backup failed'))
                    continue
                raw_dir = os.path.join(backup_root, 'Raw', full_name)
                if not os.path.isdir(raw_dir):
                    results.append((full_name, 'MISSING', 'Raw folder not found on disk at all'))
                    continue
                mismatches = []
                for rel, info in manifest.items():
                    if is_known_benign_diff(rel):
                        # Backups made before this tool recorded the manifest
                        # AFTER packing (rather than before it) can show these
                        # specific files as "changed" even though nothing is
                        # actually wrong - MakeAppx regenerates them as part of
                        # packing, so an earlier-recorded manifest is comparing
                        # against a snapshot taken too soon, not detecting a
                        # real problem. Newer backups won't hit this at all.
                        continue
                    full_path = os.path.join(raw_dir, rel)
                    if not os.path.isfile(full_path):
                        mismatches.append(f"{rel} - missing")
                        continue
                    try:
                        actual_hash = compute_file_hash(full_path)
                    except OSError as e:
                        mismatches.append(f"{rel} - couldn't read: {e}")
                        continue
                    # PowerShell's Get-FileHash (used when the manifest was
                    # recorded) always returns uppercase hex, while Python's
                    # hashlib.hexdigest() always returns lowercase - comparing
                    # them directly as strings would make every single file
                    # "mismatch" regardless of whether the actual content
                    # matches at all, which is exactly what was happening here.
                    if actual_hash.lower() != (info.get('hash') or '').lower():
                        mismatches.append(f"{rel} - hash mismatch")
                if mismatches:
                    results.append((full_name, 'MISMATCH', '; '.join(mismatches[:5]) + (f" (+{len(mismatches) - 5} more)" if len(mismatches) > 5 else '')))
                else:
                    results.append((full_name, 'OK', f"{len(manifest)} file(s) verified"))
            except Exception as e:
                # One app hitting something genuinely unexpected (a
                # permissions error, a path issue, anything not already
                # handled above) shouldn't silently kill the whole batch
                # and leave the rest unchecked with no explanation - and
                # critically, it must not prevent _show_verify_all_results
                # from ever running, which is what resets the "verify in
                # progress" guard; an uncaught exception here would leave
                # that guard stuck forever, permanently blocking every
                # future verify attempt until the app is restarted.
                results.append((full_name, 'ERROR', f"Unexpected error during verification: {e}"))
        self.root.after(0, lambda: self._show_verify_all_results(backup_root, results))

    def _write_verify_report(self, backup_root, results, path):
        """Writes the full, detailed verification report (every package,
        every mismatch detail) to a plain text file via the "Export Full
        Report" button, so the main log itself never has to hold all of
        this."""
        ok_count = sum(1 for _, status, _ in results if status == 'OK')
        lines = [
            f"Verification report for: {backup_root}",
            f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"{ok_count} OK, {len(results) - ok_count} with issues (out of {len(results)} checked)",
            "",
        ]
        for name, status, detail in sorted(results, key=lambda r: r[1] == 'OK'):
            lines.append(f"[{status}] {name}")
            if detail:
                lines.append(f"    {detail}")
        with open(path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines) + '\n')

    def _show_verify_all_results(self, backup_root, results):
        self._backup_folder_busy = False
        ok_count = sum(1 for _, status, _ in results if status == 'OK')
        problem_count = len(results) - ok_count
        summary_line = f"{ok_count} OK, {problem_count} with issues (out of {len(results)} checked)"
        self.restore_status_var.set(f"Verification complete - {summary_line}.")
        # Summary only in the main log - the full per-file detail (which
        # can genuinely be a lot of text for a large, problem-heavy
        # backup) stays in this results dialog and the optional exported
        # report instead of flooding the log. Which PACKAGES had issues
        # (not which files within them) is still worth a line here
        # though - enough to know at a glance whether to go look, without
        # needing to open the results dialog just to find that out.
        self.log_queue.put(f"--- [{log_ts()}] Verification finished - {summary_line} ---\n")
        if problem_count:
            problem_names = [name for name, status, _ in results if status != 'OK']
            preview = ', '.join(problem_names[:10])
            if len(problem_names) > 10:
                preview += f", and {len(problem_names) - 10} more"
            self.log_queue.put(f"    Affected: {preview}\n")

        win = tk.Toplevel(self.root)
        win.title("Backup verification results")
        win.minsize(500, 300)

        ttk.Label(
            win, text=f"{summary_line}:",
            padding=(10, 10, 10, 4), font=('TkDefaultFont', 9, 'bold')
        ).pack(anchor='w')

        tree_frame = ttk.Frame(win)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=(0, 6))
        tree = ttk.Treeview(tree_frame, columns=('status', 'detail'), show='tree headings', selectmode='browse')
        tree.heading('#0', text='Package')
        tree.heading('status', text='Status')
        tree.heading('detail', text='Detail (select a row to read the full text below)')
        tree.column('#0', width=280)
        tree.column('status', width=100, anchor='center')
        tree.column('detail', width=320)
        vsb = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='left', fill='y')
        tree.tag_configure('ok', foreground='#0a7d0a')
        tree.tag_configure('problem', foreground='#c00000')

        detail_by_id = {}
        # Problems first, so they're not buried below a long list of OKs.
        for name, status, detail in sorted(results, key=lambda r: r[1] == 'OK'):
            row_id = tree.insert('', 'end', text=name, values=(status, detail), tags=('ok' if status == 'OK' else 'problem',))
            detail_by_id[row_id] = detail

        # A Treeview cell doesn't wrap, so a long detail (several missing
        # files, a long path) just gets visually cut off with no way to
        # read the rest - this shows the full, unwrapped text for
        # whichever row is currently selected, and it's a normal Text
        # widget so it can be selected/copied too.
        ttk.Label(win, text="Full detail for selected row:", padding=(10, 4, 10, 2)).pack(anchor='w')
        detail_frame = ttk.Frame(win)
        detail_frame.pack(fill='x', padx=10, pady=(0, 8))
        detail_text = tk.Text(detail_frame, height=4, wrap='word')
        detail_vsb = ttk.Scrollbar(detail_frame, orient='vertical', command=detail_text.yview)
        detail_text.configure(yscrollcommand=detail_vsb.set, state='disabled')
        detail_text.pack(side='left', fill='both', expand=True)
        detail_vsb.pack(side='left', fill='y')
        self._bind_log_copy_menu(detail_text)

        def on_select(event=None):
            sel = tree.selection()
            detail_text.configure(state='normal')
            detail_text.delete('1.0', 'end')
            if sel:
                detail_text.insert('1.0', detail_by_id.get(sel[0], ''))
            detail_text.configure(state='disabled')

        tree.bind('<<TreeviewSelect>>', on_select)

        def on_export():
            default_name = f"verify-report-{datetime.datetime.now().strftime('%Y%m%d-%H%M%S')}.txt"
            path = filedialog.asksaveasfilename(
                title="Export verification report",
                defaultextension=".txt",
                filetypes=[("Text file", "*.txt"), ("All files", "*.*")],
                initialfile=default_name
            )
            if not path:
                return
            try:
                self._write_verify_report(backup_root, results, path)
                self.last_verify_report_path = path
                self.log_queue.put(f"    Full report exported to: {path}\n")
                messagebox.showinfo("Exported", f"Report written to:\n{path}")
            except OSError as e:
                messagebox.showerror("Could not export", str(e))

        btn_frame = ttk.Frame(win)
        btn_frame.pack(fill='x', padx=10, pady=(0, 10))
        ttk.Button(btn_frame, text="Export Full Report…", command=on_export).pack(side='left')
        ttk.Button(btn_frame, text="Close", command=win.destroy).pack(side='right')
        self._center_window(win, width=780, height=480)

    def _filter_rows_for_remove_mode(self, rows, remove_mode):
        """Given the current remove-mode combobox value ('Both' / 'Raw' /
        'Packed'), returns only the rows that actually have
        something for that mode to remove - e.g. "Raw" against a
        row with no raw backup would otherwise be a silent no-op, still
        showing a full confirmation dialog as if it would do something.
        Works against raw CSV rows and the GUI's own display-enriched
        rows alike, since it reads RawBackupOK/PackedOK directly rather
        than the display-only "BackupType" field (which is derived from
        exactly these two anyway, and isn't present on a raw CSV row at
        all). Returns (valid_rows, skipped_names)."""
        def has_something(r):
            has_raw = r.get('RawBackupOK') == 'True'
            has_packed = r.get('PackedOK') == 'True'
            if remove_mode == 'Raw':
                return has_raw
            elif remove_mode == 'Packed':
                return has_packed
            else:
                # "Both" always qualifies, even for a row with neither
                # piece left (BackupType "Missing") - there's still its
                # now-meaningless backup-summary.csv row to clean up, and
                # that's exactly what "Both" removal does either way.
                # Only "Raw"/"Packed" can genuinely have
                # nothing to act on, since each targets one specific
                # piece that might not exist.
                return True
        valid, skipped = [], []
        for r in rows:
            (valid if has_something(r) else skipped).append(r)
        return valid, [r.get('PackageFullName') for r in skipped]

    def _build_dependency_checklist(
        self, parent, next_row, dep_info,
        extra_warnings=None, intro_label_text=None, header_prefix_text=None,
        max_canvas_height=280, canvas_base_offset=30, wraplength=600
    ):
        """Builds the scrollable "which dependencies to also remove"
        section shared by the bulk Remove and single-row Remove Backup
        dialogs: a header row with "Select all" / "Select all safe to
        remove" (the latter only shown if at least one dependency
        genuinely qualifies), and a scrollable list below it, each
        dependency shown with a checkbox (disabled if it isn't actually
        backed up) and a status line - "not backed up", "used by N other
        app(s) not in this removal: ...", or "not used elsewhere - safe
        to remove".
        extra_warnings, if given, is a list of already-formatted warning
        strings (e.g. "AppX is itself still used by...") shown above the
        dependency rows, inside the same scrollable area - for warnings
        about the removal batch itself rather than about a specific
        dependency (only the bulk dialog has these; single-row removal
        passes None).
        parent must use grid layout; next_row is where this starts
        placing rows, and the returned next_row is where the caller
        should continue placing whatever comes after. Adds nothing (and
        returns next_row unchanged) if there's neither a dependency nor
        an extra warning to show.
        Returns (next_row, dep_vars, safe_dep_vars) - read
        dep_vars[dep].get() / safe_dep_vars after confirmation for which
        dependencies ended up checked."""
        dep_vars = {}
        safe_dep_vars = {}
        total_extra_rows = len(extra_warnings or []) + len(dep_info)
        if not total_extra_rows:
            return next_row, dep_vars, safe_dep_vars

        header_frame = ttk.Frame(parent)
        header_frame.grid(row=next_row, column=0, sticky='ew', padx=10, pady=(4, 2))
        next_row += 1
        if header_prefix_text:
            ttk.Label(header_frame, text=header_prefix_text).pack(side='left')

        select_all_var = tk.BooleanVar(value=False)

        def on_select_all_deps():
            state = select_all_var.get()
            for v in dep_vars.values():
                v.set(state)

        select_safe_var = tk.BooleanVar(value=False)

        def on_select_safe_deps():
            state = select_safe_var.get()
            for v in safe_dep_vars.values():
                v.set(state)

        list_frame = ttk.Frame(parent)
        list_frame.grid(row=next_row, column=0, sticky='nsew', padx=10, pady=(0, 4))
        parent.rowconfigure(next_row, weight=1)
        next_row += 1
        canvas = tk.Canvas(
            list_frame, highlightthickness=0,
            height=min(max_canvas_height, canvas_base_offset + total_extra_rows * 46)
        )
        vsb = ttk.Scrollbar(list_frame, orient='vertical', command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side='left', fill='both', expand=True)
        vsb.pack(side='left', fill='y')

        for warning_text in (extra_warnings or []):
            ttk.Label(
                inner, text=warning_text, foreground='#a15c00', wraplength=600, justify='left'
            ).pack(fill='x', anchor='w', pady=(4, 4))

        if dep_info and intro_label_text:
            ttk.Label(inner, text=intro_label_text).pack(anchor='w', pady=(6, 2))

        for dep, other_users, is_backed_up in dep_info:
            row_frame = ttk.Frame(inner)
            row_frame.pack(fill='x', anchor='w', pady=(4, 4))
            name_row = ttk.Frame(row_frame)
            name_row.pack(fill='x', anchor='w')
            var = tk.BooleanVar(value=False)
            dep_vars[dep] = var
            # Disabled (nothing to check) when this dependency has no
            # backup of its own to remove in the first place - it's
            # still listed here rather than hidden, so what this app (or
            # these apps) actually need stays fully visible either way.
            cb_state = 'normal' if is_backed_up else 'disabled'
            ttk.Checkbutton(name_row, variable=var, state=cb_state).pack(side='left')
            # Full name, wrapped rather than clipped - stacked above its
            # status line (below) rather than squeezed beside it, since a
            # real package name can easily run 60+ characters and a
            # fixed-width label beside the status text was cutting it off
            # outright rather than wrapping it.
            ttk.Label(
                name_row, text=dep, wraplength=wraplength, justify='left', font=('TkDefaultFont', 9, 'bold')
            ).pack(side='left', anchor='w')
            if not is_backed_up:
                status_text = "not backed up - nothing to remove"
                status_color = '#888'
            elif other_users:
                status_text = f"⚠ used by {len(other_users)} other app(s) not in this removal: {', '.join(other_users)}"
                status_color = '#a15c00'
            else:
                status_text = "not used elsewhere - safe to remove"
                status_color = '#0a7d0a'
                safe_dep_vars[dep] = var
            ttk.Label(
                row_frame, text=status_text, foreground=status_color, wraplength=wraplength, justify='left'
            ).pack(anchor='w', padx=(24, 0))

        ttk.Checkbutton(
            header_frame, text="Select all", variable=select_all_var, command=on_select_all_deps
        ).pack(side='right')
        if safe_dep_vars:
            ttk.Checkbutton(
                header_frame, text="Select all safe to remove", variable=select_safe_var,
                command=on_select_safe_deps
            ).pack(side='right', padx=(0, 8))

        return next_row, dep_vars, safe_dep_vars

    def _remove_selected_backups(self):
        """Bulk version of _remove_backup_for_app - removes every currently
        checked app's backup in one operation, with a single combined
        confirmation dialog (rather than showing one dialog per item) and
        one list refresh at the end instead of one per item."""
        backup_root = self.folder_var.get().strip()
        if not backup_root:
            messagebox.showinfo("Folder needed", "Choose a backup/restore folder first.")
            return
        checked_rows = self.restore_list.get_checked_rows()
        if not checked_rows:
            messagebox.showinfo("Nothing checked", "Check one or more apps in the list first.")
            return

        remove_mode = self.remove_mode_var.get()
        checked_rows, skipped_names = self._filter_rows_for_remove_mode(checked_rows, remove_mode)
        if not checked_rows:
            messagebox.showinfo(
                "Nothing to remove",
                f"None of the checked app(s) currently have a \"{remove_mode}\" backup to remove."
            )
            return

        main_names = [r.get('PackageFullName') for r in checked_rows]
        main_names_set = set(main_names)
        rows, fieldnames = self._read_backup_summary_csv(backup_root)

        # Any of the main apps being removed that are themselves still
        # needed by something OUTSIDE this batch (a dependency of one
        # being removed by another in the same batch doesn't count as
        # "outside" - that one's going away too).
        main_warnings = []
        for name in main_names:
            users = [
                n for n, r in rows.items()
                if n not in main_names_set and name in (r.get('Dependencies') or '').split(';')
            ]
            if users:
                main_warnings.append((name, users))

        # Union of every dependency across all apps being removed, minus
        # anything that's already one of the main apps itself.
        combined_deps = set()
        for name in main_names:
            r = rows.get(name, {})
            combined_deps.update(d for d in (r.get('Dependencies') or '').split(';') if d)
        combined_deps -= main_names_set

        dep_info = []
        for dep in sorted(combined_deps):
            is_backed_up = dep in rows
            other_users = [
                n for n, r in rows.items()
                if n not in main_names_set and dep in (r.get('Dependencies') or '').split(';')
            ]
            dep_info.append((dep, other_users, is_backed_up))

        width = 700
        win = tk.Toplevel(self.root)
        win.title("Remove selected backups")
        win.grab_set()
        win.resizable(True, True)
        win.columnconfigure(0, weight=1)
        next_row = 0

        ttk.Label(
            win, text=f"Remove the backups of {len(main_names)} app(s):",
            padding=(10, 10, 10, 4), justify='left', font=('TkDefaultFont', 9, 'bold')
        ).grid(row=next_row, column=0, sticky='w')
        next_row += 1
        # Capped rather than joining every name unconditionally - with a
        # large selection this single label could otherwise wrap to many,
        # many lines and push the buttons below off the bottom of the
        # window regardless of window height.
        names_preview = ', '.join(main_names[:15])
        if len(main_names) > 15:
            names_preview += f", and {len(main_names) - 15} more"
        ttk.Label(
            win, text=names_preview, padding=(10, 0, 10, 4), wraplength=650, justify='left'
        ).grid(row=next_row, column=0, sticky='w')
        next_row += 1
        ttk.Label(
            win,
            text="Deletes the backup files for each one (per the selected Remove mode), and "
                 "updates or removes its row in backup-summary.csv to match. This cannot be undone.",
            padding=(10, 0, 10, 6), wraplength=650, justify='left', foreground='#666'
        ).grid(row=next_row, column=0, sticky='w')
        next_row += 1

        # main_warnings and dep_info share ONE scrollable area below
        # (rather than main_warnings being separate, unbounded labels) -
        # this is what actually keeps the whole dialog's height bounded
        # regardless of how many apps are selected or how many warnings
        # or dependencies that produces.
        extra_warnings = [
            f"⚠ {name} is itself still listed as a dependency of {len(users)} other app(s) "
            f"not in this removal: {', '.join(users)} - removing it could break restoring those apps later."
            for name, users in main_warnings
        ]
        next_row, dep_vars, safe_dep_vars = self._build_dependency_checklist(
            win, next_row, dep_info,
            extra_warnings=extra_warnings,
            intro_label_text="These apps also depend on the following - check any you want removed too:",
            header_prefix_text="Details:",
            max_canvas_height=280, canvas_base_offset=30, wraplength=560
        )

        result = {'confirmed': False}

        def on_confirm():
            result['confirmed'] = True
            win.destroy()

        btn_frame = ttk.Frame(win)
        btn_frame.grid(row=next_row, column=0, sticky='ew', padx=10, pady=8)
        ttk.Button(btn_frame, text="Remove", command=on_confirm).pack(side='right')
        ttk.Button(btn_frame, text="Cancel", command=win.destroy).pack(side='right', padx=(0, 6))

        # Natural size, now safely clamped by _center_window itself if it
        # would otherwise end up taller than the screen.
        self._center_window(win, width=width)
        win.wait_window()
        if not result['confirmed']:
            return

        to_remove = main_names + [dep for dep, var in dep_vars.items() if var.get()]
        self._perform_removal_async(backup_root, to_remove, f"Removed the backups of {len(to_remove)} item(s).")

    def _export_backup_summary(self):
        """Writes a plain-text summary of everything in this backup (name,
        version, publisher, App/Framework/Dependency) - independent
        of backup-summary.csv's own internal format, for record-keeping
        or sharing."""
        backup_root = self.folder_var.get().strip()
        if not backup_root:
            messagebox.showinfo("Folder needed", "Choose a backup/restore folder first.")
            return
        rows, _ = self._read_backup_summary_csv(backup_root)
        all_rows = list(rows.values())
        if not all_rows:
            messagebox.showinfo("Nothing to export", "No backup is currently loaded here.")
            return

        default_name = f"backup-summary-{datetime.datetime.now().strftime('%Y%m%d-%H%M%S')}.txt"
        path = filedialog.asksaveasfilename(
            title="Export backup summary",
            defaultextension=".txt",
            filetypes=[("Text file", "*.txt"), ("CSV file", "*.csv"), ("All files", "*.*")],
            initialfile=default_name
        )
        if not path:
            return

        # Raw CSV rows here, not the GUI's own display-enriched ones, so
        # Type isn't already on them the way it would be after loading
        # the Restore tab - computed fresh, the same way, right here.
        def kind_of(r):
            if str(r.get('IsFramework', '')).lower() == 'true':
                return 'Framework'
            elif str(r.get('AddedAsDependency', '')).lower() == 'true':
                return 'Dependency'
            else:
                return 'App'

        apps = sorted(
            (r for r in all_rows if kind_of(r) == 'App'),
            key=lambda r: (r.get('Name') or r.get('PackageFullName') or '').lower()
        )
        deps = sorted(
            (r for r in all_rows if kind_of(r) != 'App'),
            key=lambda r: (r.get('Name') or r.get('PackageFullName') or '').lower()
        )

        try:
            if path.lower().endswith('.csv'):
                with open(path, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow(['Name', 'PackageFullName', 'Version', 'Publisher', 'Type'])
                    for r in apps + deps:
                        writer.writerow([r.get('Name', ''), r.get('PackageFullName', ''), r.get('Version', ''), r.get('Publisher', ''), kind_of(r)])
            else:
                lines = [
                    f"Backup summary for: {backup_root}",
                    f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                    f"Total: {len(all_rows)} package(s) - {len(apps)} app(s), {len(deps)} framework/dependency",
                    "",
                    f"=== Apps ({len(apps)}) ===",
                ]
                for r in apps:
                    lines.append(f"  {r.get('Name', r.get('PackageFullName', '?'))} ({r.get('Version', '?')}) - {r.get('Publisher', '?')}")
                lines.append("")
                lines.append(f"=== Frameworks / Dependencies ({len(deps)}) ===")
                for r in deps:
                    lines.append(f"  {r.get('Name', r.get('PackageFullName', '?'))} ({r.get('Version', '?')})")
                with open(path, 'w', encoding='utf-8') as f:
                    f.write('\n'.join(lines) + '\n')
        except OSError as e:
            messagebox.showerror("Could not export", str(e))
            return

        self.last_summary_path = path
        self.log_queue.put(f"\n--- [{log_ts()}] Exported backup summary to: {path} ---\n")
        messagebox.showinfo("Exported", f"Summary written to:\n{path}")

    def _create_packed_from_raw(self):
        """Packs (and signs) directly from an existing Raw backup for
        every app in the whole backup that genuinely qualifies (Raw
        present, no packed file yet) - scans backup-summary.csv itself,
        the same way "Clean Up Orphaned Frameworks" does, rather than
        only working on whatever's currently checked in the list (which
        would need first switching Select: to "Raw Backup" and clicking
        Confirm just to find them). Skips the robocopy step entirely
        since Raw is already there, so this works even for an app no
        longer installed. Uses -RepackFromRaw in Backup-AppxApps.ps1; see
        that switch's own docstring for exactly how it changes the
        script's behavior."""
        backup_root = self.folder_var.get().strip()
        if not backup_root:
            messagebox.showinfo("No folder", "Choose a backup folder first.")
            return
        rows, _ = self._read_backup_summary_csv(backup_root)
        if not rows:
            messagebox.showinfo("Nothing to check", "No backup is currently loaded here.")
            return
        candidates = [r for r in rows.values() if r.get('RawBackupOK') == 'True' and r.get('PackedOK') != 'True']
        if not candidates:
            messagebox.showinfo("Nothing to repack", "No app in this backup currently has a raw-only backup.")
            return

        win = tk.Toplevel(self.root)
        win.title("Create Packed Backups from Raw")
        win.transient(self.root)

        names = [r['PackageFullName'] for r in candidates]
        preview = ', '.join(names[:15])
        if len(names) > 15:
            preview += f", and {len(names) - 15} more"

        ttk.Label(
            win, text=f"Pack {len(candidates)} app(s) directly from their existing Raw backup:\n{preview}",
            wraplength=420, justify='left', padding=(16, 16, 16, 8)
        ).pack()

        def _on_confirm():
            win.destroy()
            include_file = self._write_include_file(names)
            cmd = ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', BACKUP_PS1,
                   '-BackupRoot', backup_root, '-IncludeFile', include_file, '-RepackFromRaw']
            if self._makeappx_path:
                cmd += ['-MakeAppxPath', self._makeappx_path]
            if self._signtool_path:
                cmd += ['-SignToolPath', self._signtool_path]
            self._append_log(f"\n=== [{log_ts()}] Creating packed backups from Raw for {len(candidates)} app(s) ===\n")
            self._run_commands_async([cmd], [f"Packing {len(candidates)} app(s) from Raw"])

        btn_row = ttk.Frame(win)
        btn_row.pack(pady=(0, 16))
        ttk.Button(btn_row, text="Cancel", command=win.destroy).pack(side='left', padx=4)
        ttk.Button(btn_row, text="Confirm", command=_on_confirm).pack(side='left', padx=4)
        self._center_window(win)

    def _create_raw_from_packed(self):
        """The reverse of _create_packed_from_raw: adds Raw for every app
        in the whole backup that has a Packed backup but no Raw one yet -
        scans backup-summary.csv itself, the same way. Never installs
        anything: for each app currently installed, copies from its live
        install location exactly like a normal backup would (the more
        accurate source when it's available); for one that isn't
        installed, extracts Raw directly from the packed .appx instead,
        since a .appx file is itself a zip archive. Uses
        -CreateRawFromPacked in Backup-AppxApps.ps1; see that switch's
        own docstring for exactly how it changes the script's behavior."""
        backup_root = self.folder_var.get().strip()
        if not backup_root:
            messagebox.showinfo("No folder", "Choose a backup folder first.")
            return
        rows, _ = self._read_backup_summary_csv(backup_root)
        if not rows:
            messagebox.showinfo("Nothing to check", "No backup is currently loaded here.")
            return
        candidates = [r for r in rows.values() if r.get('PackedOK') == 'True' and r.get('RawBackupOK') != 'True']
        if not candidates:
            messagebox.showinfo("Nothing to do", "No app in this backup currently has a packed-only backup.")
            return

        win = tk.Toplevel(self.root)
        win.title("Create Raw Backups from Packed")
        win.transient(self.root)

        names = [r['PackageFullName'] for r in candidates]
        preview = ', '.join(names[:15])
        if len(names) > 15:
            preview += f", and {len(names) - 15} more"

        ttk.Label(
            win,
            text=f"Add a Raw backup for {len(candidates)} app(s) that currently only have a "
                 f"packed backup:\n{preview}\n\nFor whichever of these are currently installed, "
                 "copies from the live install (most accurate); for the rest, extracts Raw "
                 "directly from the packed .appx instead - nothing is ever installed just to do this.",
            wraplength=420, justify='left', padding=(16, 16, 16, 8)
        ).pack()

        def _on_confirm():
            win.destroy()
            include_file = self._write_include_file(names)
            cmd = ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', BACKUP_PS1,
                   '-BackupRoot', backup_root, '-IncludeFile', include_file, '-CreateRawFromPacked']
            if self.restore_all_users_var.get():
                cmd.append('-AllUsers')
            self._append_log(f"\n=== [{log_ts()}] Creating Raw backups from Packed for {len(candidates)} app(s) ===\n")
            self._run_commands_async([cmd], [f"Adding Raw for {len(candidates)} app(s) from Packed"])

        btn_row = ttk.Frame(win)
        btn_row.pack(pady=(0, 16))
        ttk.Button(btn_row, text="Cancel", command=win.destroy).pack(side='left', padx=4)
        ttk.Button(btn_row, text="Confirm", command=_on_confirm).pack(side='left', padx=4)
        self._center_window(win)

    def _show_orphan_cleanup_dialog(self, backup_root, rows, orphans, dialog_title, header_text, success_message_fn, pre_checked=True):
        """Builds and shows the confirmation dialog shared by both orphan-
        cleanup actions (frameworks, non-framework dependencies) - a
        scrollable, checkable list of candidates with "Select all",
        followed by the same remove-mode-aware removal flow as
        everywhere else. Performs the removal itself if confirmed;
        the caller is only responsible for finding orphans (empty means
        nothing to do - callers show their own "nothing orphaned"
        message before ever reaching this, so it's never called with one)."""
        width = 640
        win = tk.Toplevel(self.root)
        win.title(dialog_title)
        win.grab_set()
        win.resizable(True, True)
        win.columnconfigure(0, weight=1)
        next_row = 0

        ttk.Label(
            win, text=header_text,
            padding=(10, 10, 10, 4), justify='left', font=('TkDefaultFont', 9, 'bold'), wraplength=620
        ).grid(row=next_row, column=0, sticky='w')
        next_row += 1
        ttk.Label(
            win,
            text="Check any you want removed - deletes the backup files for each one (per the "
                 "selected Remove mode), and updates or removes its row in backup-summary.csv to "
                 "match. This cannot be undone.",
            padding=(10, 0, 10, 6), wraplength=620, justify='left', foreground='#666'
        ).grid(row=next_row, column=0, sticky='w')
        next_row += 1

        header_frame = ttk.Frame(win)
        header_frame.grid(row=next_row, column=0, sticky='ew', padx=10, pady=(0, 2))
        next_row += 1
        orphan_vars = {}

        select_all_var = tk.BooleanVar(value=pre_checked)

        def on_select_all():
            state = select_all_var.get()
            for v in orphan_vars.values():
                v.set(state)

        ttk.Label(header_frame, text="Orphaned items:").pack(side='left')
        ttk.Checkbutton(header_frame, text="Select all", variable=select_all_var, command=on_select_all).pack(side='right')

        list_frame = ttk.Frame(win)
        list_frame.grid(row=next_row, column=0, sticky='nsew', padx=10, pady=(0, 4))
        win.rowconfigure(next_row, weight=1)
        next_row += 1
        canvas = tk.Canvas(list_frame, highlightthickness=0, height=min(240, 30 + len(orphans) * 26))
        vsb = ttk.Scrollbar(list_frame, orient='vertical', command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side='left', fill='both', expand=True)
        vsb.pack(side='left', fill='y')

        for name in orphans:
            row_frame = ttk.Frame(inner)
            row_frame.pack(fill='x', anchor='w', pady=2)
            var = tk.BooleanVar(value=pre_checked)
            orphan_vars[name] = var
            ttk.Checkbutton(row_frame, variable=var).pack(side='left')
            ttk.Label(row_frame, text=name, wraplength=560, justify='left').pack(side='left', anchor='w')

        result = {'confirmed': False}

        def on_confirm():
            result['confirmed'] = True
            win.destroy()

        btn_frame = ttk.Frame(win)
        btn_frame.grid(row=next_row, column=0, sticky='ew', padx=10, pady=8)
        ttk.Button(btn_frame, text="Remove Checked", command=on_confirm).pack(side='right')
        ttk.Button(btn_frame, text="Cancel", command=win.destroy).pack(side='right', padx=(0, 6))

        self._center_window(win, width=width)
        win.wait_window()
        if not result['confirmed']:
            return

        to_remove = [name for name, var in orphan_vars.items() if var.get()]
        if not to_remove:
            return

        remove_mode = self.remove_mode_var.get()
        to_remove_rows, skipped_names = self._filter_rows_for_remove_mode(
            [rows[name] for name in to_remove if name in rows], remove_mode
        )
        if not to_remove_rows:
            messagebox.showinfo(
                "Nothing to remove",
                f"None of the selected orphan(s) currently have a \"{remove_mode}\" backup to remove."
            )
            return
        to_remove = [r.get('PackageFullName') for r in to_remove_rows]

        self._perform_removal_async(backup_root, to_remove, success_message_fn(len(to_remove)))

    def _cleanup_orphaned_dependencies(self):
        """Scans the whole backup for frameworks no longer referenced by
        any app's Dependencies field, and offers to remove them in one
        batch - useful after several individual app removals leave
        orphaned frameworks behind with nothing left needing them.
        Deliberately limited to frameworks only (IsFramework) - without a
        reliable "was this actually chosen as its own main app" marker
        (a live IsFramework/currently-referenced check alone can't tell
        "orphaned dependency" apart from "ordinary standalone app with no
        dependents", since those look identical the moment nothing
        references it - see AddedAsDependency and
        _cleanup_orphaned_non_framework_dependencies below for how that's
        actually solved for non-framework packages instead), a
        non-framework package with no current dependents is
        indistinguishable from a perfectly normal standalone app that
        simply has no dependents either - which describes most apps. A
        framework carries no such ambiguity; nothing legitimately wants
        one for its own sake, so an unused one is safe to always treat
        as orphaned."""
        backup_root = self.folder_var.get().strip()
        if not backup_root:
            messagebox.showinfo("Folder needed", "Choose a backup/restore folder first.")
            return
        rows, fieldnames = self._read_backup_summary_csv(backup_root)
        if not rows:
            messagebox.showinfo("Nothing to check", "No backup is currently loaded here.")
            return

        orphans = []
        for name, row in rows.items():
            is_framework = str(row.get('IsFramework', '')).lower() == 'true'
            if not is_framework:
                continue  # only frameworks are unambiguously "orphanable" - see docstring above
            used_by = [
                n for n, r in rows.items()
                if n != name and name in (r.get('Dependencies') or '').split(';')
            ]
            if not used_by:
                orphans.append(name)
        orphans.sort()

        if not orphans:
            messagebox.showinfo(
                "Nothing orphaned",
                "Every framework in this backup is still used by at least one app."
            )
            return

        self._show_orphan_cleanup_dialog(
            backup_root, rows, orphans,
            dialog_title="Clean up orphaned frameworks",
            header_text=f"Found {len(orphans)} framework(s) no longer used by any app in this backup:",
            success_message_fn=lambda n: f"Removed {n} orphaned framework backup(s).",
            pre_checked=True
        )

    def _cleanup_orphaned_non_framework_dependencies(self):
        """Scans the whole backup for non-framework dependencies no
        longer referenced by any app's Dependencies field AND that were
        only ever added to this backup as a dependency in the first
        place - never directly, explicitly selected (AddedAsDependency,
        recorded once at backup time in Backup-AppxApps.ps1; see its own
        docstring there). Unlike a live IsFramework check, "was this ever
        explicitly selected" is a historical fact fixed at backup time,
        so it stays reliable even after whatever used to depend on this
        one gets removed later - a live-recomputed check couldn't tell
        "orphaned dependency" apart from "ordinary standalone app with no
        dependents" (the same thing a plain unreferenced check runs into
        for _cleanup_orphaned_dependencies above), since both look
        identical the instant nothing references them; this field is
        what actually breaks that tie. Not quite as absolute as the
        framework version, though: a row from a backup made before this
        field existed has no recorded answer, and defaults to treating a
        currently-unreferenced one as fair game here (see
        AddedAsDependency's own docstring for why that specific default
        was chosen) rather than silently excluding it."""
        backup_root = self.folder_var.get().strip()
        if not backup_root:
            messagebox.showinfo("Folder needed", "Choose a backup/restore folder first.")
            return
        rows, fieldnames = self._read_backup_summary_csv(backup_root)
        if not rows:
            messagebox.showinfo("Nothing to check", "No backup is currently loaded here.")
            return

        orphans = []
        for name, row in rows.items():
            is_framework = str(row.get('IsFramework', '')).lower() == 'true'
            if is_framework:
                continue  # frameworks have their own separate action above
            # Missing/empty (a backup made before this field existed)
            # defaults to "yes, treat as added-only-as-a-dependency" -
            # see the field's own docstring in Backup-AppxApps.ps1.
            added_as_dependency = str(row.get('AddedAsDependency', '')).strip().lower() != 'false'
            if not added_as_dependency:
                continue  # explicitly selected at some point - never orphanable
            used_by = [
                n for n, r in rows.items()
                if n != name and name in (r.get('Dependencies') or '').split(';')
            ]
            if not used_by:
                orphans.append(name)
        orphans.sort()

        if not orphans:
            messagebox.showinfo(
                "Nothing orphaned",
                "Every non-framework dependency in this backup is still used by at least one "
                "app (or was explicitly chosen on its own)."
            )
            return

        self._show_orphan_cleanup_dialog(
            backup_root, rows, orphans,
            dialog_title="Clean up orphaned dependencies",
            header_text=f"Found {len(orphans)} non-framework dependency(ies) no longer used by any app in this backup:",
            success_message_fn=lambda n: f"Removed {n} orphaned dependency backup(s).",
            pre_checked=True
        )

    def _remove_backup_for_app(self, row):
        """Confirms and removes one app's backup (its Raw folder, Packed
        .appx, and Manifests JSON, or a subset of those per the selected
        Remove mode - and its backup-summary.csv row, updated or removed
        to match). If the app has its
        own dependencies, the same dialog lists each one with a checkbox to
        remove it too - shown regardless of whether it's still shared with
        another app, since that's your call to make, not something to
        silently block; a shared one is just clearly marked as such so you
        know what removing it would affect."""
        backup_root = self.folder_var.get().strip()
        if not backup_root:
            messagebox.showinfo("Folder needed", "Choose a backup/restore folder first.")
            return
        full_name = row.get('PackageFullName')
        rows, fieldnames = self._read_backup_summary_csv(backup_root)
        my_row = rows.get(full_name, row)

        deps = [d for d in (my_row.get('Dependencies') or '').split(';') if d]

        dep_info = []
        for dep in deps:
            is_backed_up = dep in rows
            other_users = [
                n for n, r in rows.items()
                if n != full_name and dep in (r.get('Dependencies') or '').split(';')
            ]
            dep_info.append((dep, other_users, is_backed_up))

        # The item being removed might itself be listed as a dependency of
        # some other app - e.g. removing a framework directly (via the
        # "Frameworks" category view) rather than through "Show
        # Dependencies", which is the only other place this was being
        # checked before. Checked here too so this warning always shows
        # regardless of which path led to this dialog.
        users_of_this = [
            n for n, r in rows.items()
            if n != full_name and full_name in (r.get('Dependencies') or '').split(';')
        ]

        # Sized to what's actually being shown, computed after the widgets
        # exist (see _center_window below) rather than pre-guessed - a
        # manual estimate here previously ended up too small once there
        # were dependencies to list, clipping the button row.
        width = 680
        win = tk.Toplevel(self.root)
        win.title("Remove backup")
        # Deliberately NOT calling win.transient(self.root) here - per
        # Tcl/Tk's own documented behavior, a transient window is expected
        # to lose its minimize button and, on Windows specifically, often
        # its maximize button too (treated as a "dialog-style" window by
        # the window manager) - which defeated the whole point of making
        # this resizable in the first place. grab_set() below still makes
        # it modal on its own, independent of transient().
        win.grab_set()
        win.resizable(True, True)

        # Grid with row weights, rather than pack + a manual <Configure>
        # binding to sync the canvas's height - binding to <Configure>
        # directly turned out fragile across a minimize/restore cycle
        # (Windows can report a transient, incorrect size mid-transition,
        # which then got baked into the canvas's height and pushed the
        # button row out of view). Grid's own row-weight mechanism handles
        # "this row expands, these don't" natively, without needing to
        # track pixel sizes by hand at all.
        win.columnconfigure(0, weight=1)
        next_row = 0

        ttk.Label(
            win, text=f"Remove the backup of:  {full_name}",
            padding=(10, 10, 10, 4), justify='left', font=('TkDefaultFont', 9, 'bold')
        ).grid(row=next_row, column=0, sticky='w')
        next_row += 1

        mode_row = ttk.Frame(win)
        mode_row.grid(row=next_row, column=0, sticky='w', padx=10, pady=(0, 2))
        next_row += 1
        ttk.Label(mode_row, text="Remove:").pack(side='left')
        # Shares the same variable as the Restore tab's own Remove combobox,
        # so changing either one changes the other.
        dialog_remove_mode_var = self.remove_mode_var
        ttk.Combobox(
            mode_row, textvariable=dialog_remove_mode_var, values=['Both', 'Raw', 'Packed'],
            state='readonly', width=9
        ).pack(side='left', padx=(6, 0))

        ttk.Label(
            win,
            text="Deletes its backup files (per the mode above), and updates or "
                 "removes its row in backup-summary.csv to match. This cannot be undone.",
            padding=(10, 0, 10, 6), wraplength=650, justify='left', foreground='#666'
        ).grid(row=next_row, column=0, sticky='w')
        next_row += 1

        if users_of_this:
            ttk.Label(
                win,
                text=f"⚠ {full_name} is itself still listed as a dependency of "
                     f"{len(users_of_this)} other app(s): {', '.join(users_of_this)} - "
                     f"removing it could break restoring those apps later.",
                padding=(10, 0, 10, 6), wraplength=650, justify='left', foreground='#a15c00'
            ).grid(row=next_row, column=0, sticky='w')
            next_row += 1

        next_row, dep_vars, safe_dep_vars = self._build_dependency_checklist(
            win, next_row, dep_info,
            header_prefix_text="This app also depends on the following - check any you want removed too:",
            max_canvas_height=240, canvas_base_offset=40, wraplength=600
        )

        result = {'confirmed': False}

        def on_confirm():
            remove_mode = dialog_remove_mode_var.get()
            valid_rows, _ = self._filter_rows_for_remove_mode([my_row], remove_mode)
            if not valid_rows:
                messagebox.showinfo(
                    "Nothing to remove",
                    f"{full_name} doesn't currently have a \"{remove_mode}\" backup to remove."
                )
                return
            result['confirmed'] = True
            result['remove_mode'] = remove_mode
            win.destroy()

        btn_frame = ttk.Frame(win)
        btn_frame.grid(row=next_row, column=0, sticky='ew', padx=10, pady=8)
        ttk.Button(btn_frame, text="Remove", command=on_confirm).pack(side='right')
        ttk.Button(btn_frame, text="Cancel", command=win.destroy).pack(side='right', padx=(0, 6))

        self._center_window(win, width=width)
        win.wait_window()
        if not result['confirmed']:
            return

        to_remove = [full_name] + [dep for dep, var in dep_vars.items() if var.get()]
        self._perform_removal_async(
            backup_root, to_remove, "Removed the backup of:\n" + "\n".join(to_remove),
            remove_mode=result['remove_mode']
        )

    def _show_dependencies_dialog(self, row):
        """Lists what this app depends on, and for each dependency, which
        OTHER backed-up apps also depend on it - so you can tell whether
        it's safe to remove that dependency's own backup separately.
        Each dependency is a parent row with its users as child rows
        underneath (rather than a comma-joined list crammed into one
        column) - a Treeview cell doesn't wrap, so joining several
        (often long) package names together just gets visually cut off
        once there's more than one or two."""
        backup_root = self.folder_var.get().strip()
        if not backup_root:
            messagebox.showinfo("Folder needed", "Choose a backup/restore folder first.")
            return
        full_name = row.get('PackageFullName')
        rows, _ = self._read_backup_summary_csv(backup_root)
        my_row = rows.get(full_name)
        # NOT filtered against rows here - a dependency stays listed even
        # after its own backup is removed (removing it doesn't rewrite
        # this app's own Dependencies field, and that's fine: the list
        # below shows it as "not backed up" rather than making it vanish,
        # so you can still see everything this app is meant to need).
        deps = [d for d in (my_row.get('Dependencies') if my_row else '').split(';') if d] if my_row else []

        if not deps:
            messagebox.showinfo("Dependencies", f"{full_name} has no recorded dependencies in this backup.")
            return

        win = tk.Toplevel(self.root)
        win.title(f"Dependencies of {full_name}")
        self._center_window(win, width=760, height=360)
        win.minsize(500, 250)

        label_var = tk.StringVar(value=f"Dependencies of {full_name} (expand one to see who else uses it):")
        ttk.Label(win, textvariable=label_var, padding=(8, 8, 8, 0)).pack(anchor='w')

        tree_frame = ttk.Frame(win)
        tree_frame.pack(fill='both', expand=True, padx=8, pady=8)
        tree = ttk.Treeview(tree_frame, columns=('status',), show='tree headings', selectmode='browse')
        tree.heading('#0', text='Dependency / used by')
        tree.heading('status', text='Status')
        tree.column('#0', width=460)
        tree.column('status', width=200)
        vsb = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='left', fill='y')

        dep_users = {}
        dep_backed_up = {}

        def refresh_tree():
            """Re-reads the backup and repopulates the tree in place -
            called for the initial population and again after removing a
            dependency, so this dialog can stay open to remove several in
            a row instead of closing after every single one. Returns
            False only if there are no dependency names left at all."""
            tree.delete(*tree.get_children())
            dep_users.clear()
            dep_backed_up.clear()
            current_rows, _ = self._read_backup_summary_csv(backup_root)
            current_row = current_rows.get(full_name)
            # NOT filtered against current_rows - a dependency stays
            # listed even after its own backup is removed, shown as "not
            # backed up" below instead of disappearing, so the full set of
            # what this app is meant to need stays visible.
            current_deps = [
                d for d in (current_row.get('Dependencies') if current_row else '').split(';') if d
            ] if current_row else []
            if not current_deps:
                return False
            for dep in current_deps:
                is_backed_up = dep in current_rows
                dep_backed_up[dep] = is_backed_up
                users = [
                    other_name for other_name, other_row in current_rows.items()
                    if other_name != full_name and dep in (other_row.get('Dependencies') or '').split(';')
                ]
                dep_users[dep] = users
                if not is_backed_up:
                    status = "NOT BACKED UP - nothing to remove"
                elif users:
                    status = f"Used by {len(users)} other app(s)"
                else:
                    status = "Not used elsewhere - safe to remove"
                tree.insert('', 'end', iid=dep, text=dep, values=(status,), open=bool(users))
                for user in users:
                    tree.insert(dep, 'end', text=f"    {user}", values=('',))
            return True

        # Belt-and-suspenders alongside the check above (which should
        # already prevent getting here at all with nothing to show) - if
        # it somehow still happens, close immediately rather than leave a
        # blank dialog open with nothing in it.
        if not refresh_tree():
            win.destroy()
            return

        btn_frame = ttk.Frame(win)
        btn_frame.pack(fill='x', padx=8, pady=(0, 8))

        def remove_selected_dependency():
            sel = tree.selection()
            if not sel:
                messagebox.showinfo("Nothing selected", "Select a dependency in the list first.")
                self._bring_window_forward(win)
                return
            item = sel[0]
            if tree.parent(item):
                messagebox.showinfo(
                    "Select the dependency, not the app",
                    "This is one of the apps using the dependency, not the dependency itself. "
                    "Select the dependency's own row (the one above it) to remove its backup."
                )
                self._bring_window_forward(win)
                return
            dep = item
            if not dep_backed_up.get(dep, True):
                messagebox.showinfo(
                    "Not backed up",
                    f"{dep} isn't backed up - there's nothing to remove for it."
                )
                self._bring_window_forward(win)
                return
            # Deliberately doesn't close this window first -
            # _remove_backup_for_app shows its own confirmation (including
            # a warning if dep is still used elsewhere, and a checklist
            # for any of its own dependencies too); once that's done, the
            # tree just refreshes in place so several dependencies can be
            # removed one after another without reopening this dialog
            # each time.
            self._remove_backup_for_app({'PackageFullName': dep})
            # The confirmation dialog isn't transient to this window (it's
            # transient to nothing at all, deliberately, to keep its own
            # maximize button - see _remove_backup_for_app), so closing it
            # doesn't automatically return focus here the way a transient
            # child normally would - confirmed this fix works for this
            # specific path.
            self._bring_window_forward(win)
            if not refresh_tree():
                win.destroy()

        ttk.Button(
            btn_frame, text="Remove Selected Dependency's Backup", command=remove_selected_dependency
        ).pack(side='left')
        ttk.Label(
            btn_frame, text="The next dialog will warn you if it's still used elsewhere.", foreground='#666'
        ).pack(side='left', padx=8)
        ttk.Button(btn_frame, text="Close", command=win.destroy).pack(side='right')

    def _show_used_by_dialog(self, row):
        """The reverse of "Show Dependencies…" - for a dependency or
        framework row, shows which OTHER backed-up apps actually need it.
        Useful when looking at a dependency directly (e.g. after filtering
        to "Frameworks") without already knowing which app to check from."""
        backup_root = self.folder_var.get().strip()
        if not backup_root:
            messagebox.showinfo("Folder needed", "Choose a backup/restore folder first.")
            return
        full_name = row.get('PackageFullName')
        rows, _ = self._read_backup_summary_csv(backup_root)
        users = [
            other_name for other_name, other_row in rows.items()
            if other_name != full_name and full_name in (other_row.get('Dependencies') or '').split(';')
        ]

        if not users:
            messagebox.showinfo(
                "Not used by anything else",
                f"{full_name} isn't listed as a dependency of any other app in this backup."
            )
            return

        win = tk.Toplevel(self.root)
        win.title(f"Apps that use {full_name}")
        self._center_window(win, width=640, height=320)
        win.minsize(400, 200)

        ttk.Label(
            win, text=f"{full_name} is a dependency of {len(users)} app(s) in this backup:",
            padding=(8, 8, 8, 0)
        ).pack(anchor='w')

        list_frame = ttk.Frame(win)
        list_frame.pack(fill='both', expand=True, padx=8, pady=8)
        listbox = tk.Listbox(list_frame)
        vsb = ttk.Scrollbar(list_frame, orient='vertical', command=listbox.yview)
        listbox.configure(yscrollcommand=vsb.set)
        listbox.pack(side='left', fill='both', expand=True)
        vsb.pack(side='left', fill='y')
        for user in sorted(users):
            listbox.insert('end', user)

        ttk.Button(win, text="Close", command=win.destroy).pack(pady=(0, 8))

    def show_app_files(self, row):
        backup_root = self.folder_var.get().strip()
        if not backup_root:
            messagebox.showinfo("Folder needed", "Choose a backup/restore folder first.")
            return
        full_name = row['PackageFullName']
        text = get_app_backup_files(backup_root, full_name)
        display_name = row.get('Name', full_name)
        self._show_text_dialog(f"Files — {display_name}", text, geometry="750x600")

    def show_app_file_verifier(self, row):
        backup_root = self.folder_var.get().strip()
        if not backup_root:
            messagebox.showinfo("Folder needed", "Choose a backup/restore folder first.")
            return
        full_name = row['PackageFullName']
        display_name = row.get('Name', full_name)
        raw_dir, files = get_raw_backup_file_list(backup_root, full_name)
        if not files:
            messagebox.showinfo("No files", f"No raw backup files found for '{display_name}' at:\n{raw_dir}")
            return
        FileVerifierDialog(self, full_name, display_name, files)

    def _on_backup_right_click(self, event):
        row_id = self.backup_list.tree.identify_row(event.y)
        col = self.backup_list.tree.identify_column(event.x)
        values = self.backup_list.tree.item(row_id, 'values') if row_id else ()
        col_index = None
        if col and col.startswith('#'):
            try:
                col_index = int(col[1:]) - 1
            except ValueError:
                pass
        cell_value = values[col_index] if col_index is not None and 0 <= col_index < len(values) else ''

        row = None
        if row_id in self.backup_list.row_id_to_key:
            self.backup_list.tree.selection_set(row_id)
            key = self.backup_list.row_id_to_key[row_id]
            row = next((r for r in self.backup_list.all_rows if r['PackageFullName'] == key), None)

        menu = tk.Menu(self.backup_list.tree, tearoff=0)
        if row is not None:
            menu.add_command(label="Open Folder", command=lambda: self._open_backup_app_folder(row))
            menu.add_separator()
        if row_id:
            menu.add_command(label="Copy Cell", command=lambda: self.backup_list._copy_to_clipboard(str(cell_value)))
            menu.add_command(label="Copy Row (all columns)", command=lambda: self.backup_list._copy_to_clipboard('\t'.join(str(v) for v in values)))
        menu.add_command(label="Copy All Visible Rows", command=self.backup_list.copy_all_visible_rows)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _open_backup_app_folder(self, row):
        install_location = row.get('InstallLocation')
        if not install_location:
            messagebox.showwarning("Not found", "This app has no recorded install location.")
            return
        self._open_folder_in_explorer(install_location)

    def _center_window(self, win, width=None, height=None):
        """Centers a Toplevel over the main window. Pass width/height for a
        fixed-size dialog; omit both to let Tkinter's own natural layout
        determine the size instead - the more robust choice whenever a
        dialog's content varies (e.g. a variable-length list), since a
        manually guessed fixed size can end up too small and clip
        something like the button row below it. Either way, the result is
        clamped to a fraction of the screen - a dialog whose content grew
        large enough (e.g. many items selected in a bulk operation) could
        otherwise end up taller than the screen itself, which is what
        actually hides buttons at the bottom regardless of any internal
        scrolling a dialog does on its own."""
        win.update_idletasks()
        if width is None:
            width = win.winfo_reqwidth()
        if height is None:
            height = win.winfo_reqheight()
        height = min(height, int(win.winfo_screenheight() * 0.85))
        width = min(width, int(win.winfo_screenwidth() * 0.9))
        self.root.update_idletasks()
        root_x, root_y = self.root.winfo_rootx(), self.root.winfo_rooty()
        root_w, root_h = self.root.winfo_width(), self.root.winfo_height()
        x = max(0, root_x + (root_w - width) // 2)
        y = max(0, root_y + (root_h - height) // 2)
        win.geometry(f"{width}x{height}+{x}+{y}")

    def _bring_window_forward(self, win):
        """Forces a Toplevel back to the front/focus after a modal child
        dialog (that isn't transient to it - see _remove_backup_for_app)
        closes. lift()/focus_force() alone didn't reliably work, likely
        Windows' own foreground-lock behavior refusing a plain focus
        request depending on timing - toggling -topmost briefly is a more
        forceful, commonly-used workaround for exactly that."""
        win.lift()
        win.attributes('-topmost', True)
        win.attributes('-topmost', False)
        win.focus_force()

    def _open_backup_root_folder(self):
        folder = self.folder_var.get().strip()
        if not folder:
            messagebox.showinfo("Folder needed", "Choose a backup/restore folder first.")
            return
        self._open_folder_in_explorer(folder)

    def _open_folder_in_explorer(self, folder):
        if not os.path.isdir(folder):
            messagebox.showwarning("Folder not found", f"This folder doesn't exist (any more?):\n{folder}")
            return
        try:
            os.startfile(folder)
        except OSError as e:
            messagebox.showerror("Could not open folder", f"{folder}\n\n{e}")

    def _show_in_explorer(self, path):
        """Opens the containing folder with this exact file/folder
        highlighted, via the actual Windows Shell API (the same one
        Explorer/browsers/download managers use for "Show in folder") -
        this sidesteps explorer.exe's own command-line parsing entirely,
        which turned out to be unreliable via subprocess regardless of how
        the /select, argument was quoted."""
        if not os.path.exists(path):
            messagebox.showwarning("Not found", f"This no longer exists:\n{path}")
            return
        try:
            select_in_explorer(path)
        except OSError as e:
            # Fall back to just opening the containing folder - still
            # useful even without the specific file highlighted.
            try:
                os.startfile(os.path.dirname(path))
            except OSError:
                messagebox.showerror("Could not open Explorer", str(e))

    def _browse_folder(self):
        folder = filedialog.askdirectory(title="Choose the backup/restore folder")
        if folder:
            self.folder_var.set(folder)
            self._remember_folder(folder)
            # Auto-load restore contents if this folder already has a backup
            # in it - stays silent if it doesn't (e.g. picking a fresh empty
            # folder to back up INTO for the first time).
            if os.path.isfile(os.path.join(folder, 'backup-summary.csv')):
                self.load_restore_list()

    def _remember_folder(self, folder):
        self.recent_folders = remember_recent_folder(folder)
        self.folder_combo.configure(values=self.recent_folders)

    def _on_recent_folder_selected(self, event=None):
        folder = self.folder_var.get().strip()
        if folder and os.path.isfile(os.path.join(folder, 'backup-summary.csv')):
            self.load_restore_list()

    # ---------------- Backup tab logic ----------------

    def refresh_installed_apps(self):
        """Fetches every installed package (see get_installed_apps) and
        recalculates sizes - can take a minute, which invites an
        impatient re-click while one is already running. A token
        (matching the same pattern already used for load_restore_list)
        makes an older, now-superseded refresh's result get silently
        discarded when it eventually completes, rather than briefly
        overwriting a newer, already-applied one - since this is a
        read-only fetch, there's no data-corruption risk in letting more
        than one run at once, just a UI flicker/revert this avoids."""
        self.backup_load_token += 1
        my_token = self.backup_load_token
        self.backup_status_var.set("Loading installed apps and calculating sizes... this can take a minute.")
        threading.Thread(target=self._refresh_installed_apps_worker, args=(my_token,), daemon=True).start()

    def _refresh_installed_apps_worker(self, token):
        try:
            apps = get_installed_apps(self.all_users_var.get())
        except Exception as e:
            self.log_queue.put(f"ERROR loading apps: {e}\n")
            self.root.after(0, lambda: self.backup_status_var.set("Failed to load apps - see log."))
            return
        for a in apps:
            a['size_display'] = humansize(a.get('SizeBytes'))
        apps.sort(key=lambda a: (a.get('Name') or '').lower())
        self.root.after(0, lambda: self._on_installed_apps_loaded(apps, token))

    def _on_backup_category_changed(self):
        """Client-side filter, applied to whatever's already loaded -
        deliberately not a re-fetch, since Get-AppxPackage/Get-StartApps
        now only ever run once (see get_installed_apps), on "Refresh
        List" or the initial load. Switching "Show:" or toggling "Only
        show Start Menu apps" used to mean waiting for a fresh PowerShell
        call every single time; now it's instant, and (since this only
        changes what's visible, not the underlying row list) whatever's
        checked elsewhere in the list stays checked too."""
        category = self.backup_category_var.get()
        only_start_apps = self.only_start_apps_var.get()

        def is_app(row):
            if row.get('IsFramework') or row.get('IsSystemSigned'):
                return False
            if only_start_apps and not row.get('IsStartMenuVisible', True):
                return False
            return True

        if category == "Frameworks":
            self.backup_list.extra_filter_fn = lambda r: bool(r.get('IsFramework'))
        elif category == "All":
            # Apps and Frameworks together - deliberately not every
            # installed package: a system-signed, non-framework package
            # is neither, and stays hidden here the same as under either
            # of the other two options, not just lumped into "All" by
            # virtue of not being excluded elsewhere.
            self.backup_list.extra_filter_fn = lambda r: is_app(r) or bool(r.get('IsFramework'))
        else:  # "Apps"
            self.backup_list.extra_filter_fn = is_app
        self.backup_list._apply_filter()

    def _on_installed_apps_loaded(self, apps, token):
        if token != self.backup_load_token:
            return  # a newer refresh has since started - this one is stale, ignore it
        self.backup_list.set_rows(apps, preserve_checked=True)
        self._on_backup_category_changed()  # apply whatever "Show:" is currently set to
        self.backup_status_var.set(f"{len(apps)} package(s) loaded.")
        self._update_backup_total()

    def _update_backup_total(self):
        selected = self.backup_list.get_checked_rows()
        total_size = sum((a.get('SizeBytes') or 0) for a in selected)
        visible_count = len(self.backup_list.row_id_to_key)
        self.backup_total_var.set(
            f"{visible_count} in list - Selected: {len(selected)} app(s), ~{humansize(total_size)} total"
        )

    def _update_restore_total(self):
        selected = self.restore_list.get_checked_rows()
        total_size = sum((a.get('SizeBytes') or 0) for a in selected)
        visible_count = len(self.restore_list.row_id_to_key)
        self.restore_total_var.set(
            f"{visible_count} in list - Selected: {len(selected)} app(s), ~{humansize(total_size)} total"
        )

    def _write_include_file(self, names):
        """Write names (one per line) to a fresh temp file and return its path.
        Used instead of passing many values on the command line to -Include,
        which is the safer, version-independent way to hand a long list to
        the PowerShell scripts."""
        fd, path = tempfile.mkstemp(prefix='appx_include_', suffix='.txt')
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            f.write('\n'.join(names))
        return path

    def start_backup(self):
        if not os.path.isfile(BACKUP_PS1):
            messagebox.showerror("Script not found", f"Backup-AppxApps.ps1 could not be located:\n{BACKUP_PS1}")
            return
        backup_root = self.folder_var.get().strip()
        if not backup_root:
            messagebox.showinfo("Folder needed", "Choose a backup/restore folder first.")
            return
        selected = self.backup_list.get_checked_rows()
        if not selected:
            messagebox.showinfo("Nothing selected", "Check at least one app in the Backup tab's list first.")
            return

        # Checked once, upfront, so a wrong/missing path is caught before
        # any work starts rather than discovered app-by-app partway
        # through a long batch. Not a hard block - backup still works
        # without either, just as a raw-folder-only backup - so this is a
        # confirm/cancel rather than an error.
        problems = []
        if self._makeappx_path and not os.path.isfile(self._makeappx_path):
            problems.append(f"MakeAppx.exe path doesn't exist:\n  {self._makeappx_path}")
        if self._signtool_path and not os.path.isfile(self._signtool_path):
            problems.append(f"SignTool.exe path doesn't exist:\n  {self._signtool_path}")
        if problems:
            if not messagebox.askyesno(
                "Tool path problem",
                "\n\n".join(problems) +
                "\n\nPacking and/or signing will fail for every app in this batch until this is fixed "
                "(you'll still get raw-folder backups). Continue anyway?"
            ):
                return

        self._remember_folder(backup_root)

        # One invocation per checked app (each pulling in its own dependencies
        # automatically) rather than one invocation covering all of them -
        # this keeps each run's -Include down to a single name, which is the
        # simplest possible case for PowerShell's parameter binding to get
        # right regardless of version, and the merge/skip logic already in
        # Backup-AppxApps.ps1 means shared dependencies still only get
        # copied/packed/signed once across the whole sequence.
        #
        # Each app's command is built lazily (functools.partial, called by
        # _run_commands_worker right before that app actually runs) rather
        # than all at once here - building every command upfront would bake
        # in whatever the MakeAppx/SignTool path fields say at this exact
        # moment as a literal string, so fixing a wrong path mid-batch (after
        # noticing app 1 failed to pack) would never reach apps 2, 3, ... in
        # the same run, since their command was already fixed in place
        # before the fix ever happened.
        commands = []
        labels = []
        # all_users/force/backup_mode captured once, here, on the main
        # thread, rather than re-read inside the builder below - unlike
        # the tool paths, there's no reported need for these to be
        # "fixable mid-batch", and capturing them as plain values now
        # means the builder never needs to touch a Tkinter variable at
        # all when it actually runs (which happens on a background
        # thread - see _build_backup_cmd).
        all_users = self.all_users_var.get()
        force = self.force_var.get()
        backup_mode = self.backup_mode_var.get()
        for a in selected:
            name = a['PackageFullName']
            include_file = self._write_include_file([name])
            is_framework = bool(a.get('IsFramework'))
            commands.append(functools.partial(
                self._build_backup_cmd, backup_root, name, include_file, all_users, force, backup_mode, is_framework
            ))
            labels.append(f"Backing up {name}")

        self._append_log(f"\n=== [{log_ts()}] Starting backup of {len(commands)} app(s), one at a time (dependencies added automatically per app) ===\n")
        self._run_commands_async(commands, labels)

    def _build_backup_cmd(self, backup_root, name, include_file, all_users, force, backup_mode, is_framework=False):
        """Builds one app's Backup-AppxApps.ps1 command. Called from
        _run_commands_worker's background thread, right before this app's
        turn - so the MakeAppx/SignTool paths below are read from the
        plain Python attributes _makeappx_path/_signtool_path (kept in
        sync by a trace on the main thread - see where those StringVars
        are created), never from the Tkinter variables directly, since
        calling .get() on a Tkinter variable from a non-main thread is
        unsafe. all_users/force/backup_mode/is_framework are passed in as
        plain values, captured once upfront in start_backup, for the
        same reason.

        backup_mode is one of:
          'raw'    - -SkipPack: Raw folder only, no packed .appx at all.
          'packed' - -DeleteRawAfterPack: packs (and signs, unless
                     -SkipSign) as normal, then deletes the Raw folder,
                     keeping only the packed .appx. Trades away
                     Register-mode restore and the Raw-folder-based
                     integrity checks for this app - see that switch's
                     own docstring in Backup-AppxApps.ps1 for why.
          'both'   - the default either way: neither flag, keeps both.

        is_framework matters because the script's own -Include filter
        excludes frameworks unless -IncludeFrameworks is also passed
        (mirroring how a normal, no-Include backup run wouldn't back up
        every framework on the system just because something depends on
        one) - without this, explicitly checking a framework here and
        backing it up on its own would fail with "no installed packages
        matched", since the very framework you asked for would be
        filtered straight back out."""
        cmd = ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', BACKUP_PS1,
               '-BackupRoot', backup_root, '-IncludeFile', include_file]
        if all_users:
            cmd.append('-AllUsers')
        if force:
            cmd.append('-Force')
        if is_framework:
            cmd.append('-IncludeFrameworks')
        if backup_mode == 'raw':
            cmd.append('-SkipPack')
        elif backup_mode == 'packed':
            cmd.append('-DeleteRawAfterPack')
        if self._makeappx_path:
            cmd += ['-MakeAppxPath', self._makeappx_path]
        if self._signtool_path:
            cmd += ['-SignToolPath', self._signtool_path]
        return cmd

    # ---------------- Restore tab logic ----------------

    def _on_restore_category_changed(self, event=None):
        """A fully independent filtering layer via CheckableAppList's
        extra_filter_fn, reading the already-computed "Type" field
        (see _compute_backup_display_fields) rather than a per-column
        filter StringVar - keeps this from drifting out of sync with
        "Clear Filters" or a manual edit to that column's own filter
        widget."""
        category_to_type = {"Apps": "App", "Frameworks": "Framework", "Dependencies": "Dependency"}
        wanted = category_to_type.get(self.restore_category_var.get())
        if wanted is None:  # "All"
            self.restore_list.extra_filter_fn = None
        else:
            self.restore_list.extra_filter_fn = lambda r: r.get('Type') == wanted
        self.restore_list._apply_filter()

    def _perform_removal_async(self, backup_root, to_remove, success_message, remove_mode=None):
        """Runs the actual file deletion + CSV update for a removal in a
        background thread - deleting many files/folders synchronously on
        the main thread (e.g. a bulk removal of a lot of apps) would
        freeze the whole UI until it finished. Logged to the main log
        pane (start and a one-line result) so it's visible there too, not
        just in a message box you might miss.

        Guarded against overlapping runs (shares _backup_folder_busy with
        Verify All Backups): since this no longer blocks the UI, nothing
        else would otherwise stop a second removal (single, bulk, or
        orphan cleanup - they all funnel through here) from starting
        while the first is still running, and both would then read/write
        backup-summary.csv concurrently - whichever finishes last would
        silently overwrite the other's CSV change, even though that first
        operation's files were already deleted from disk. Sharing the
        flag with Verify All Backups also stops it from hashing files a
        removal is deleting out from under it at the same moment.

        The "check files aren't locked" pre-check (if on) runs in the
        background, not here on the caller's thread - a bulk removal over
        many apps could mean scanning a lot of files, and that must never
        block the UI any more than the removal itself does. _backup_folder_busy
        is claimed immediately though, before that scan even starts - not
        after - so a second removal attempt during the scan is still
        correctly refused rather than sneaking in through the gap."""
        if self._backup_folder_busy:
            messagebox.showinfo("Busy", "A removal or verification is already in progress - please wait for it to finish first.")
            return
        self._backup_folder_busy = True
        # Captured now, on the main thread, rather than read inside the
        # worker (which runs on a background thread) - same reasoning as
        # the MakeAppx/SignTool paths elsewhere: a Tkinter variable should
        # never be read from a non-main thread. A caller with its own,
        # already-confirmed mode (e.g. the right-click "Remove backup…"
        # dialog's own combobox) passes it directly instead of falling
        # back to the shared one here.
        if remove_mode is None:
            remove_mode = self.remove_mode_var.get()
        self._prepare_and_run(
            backup_root, to_remove, "removing them", "a removal",
            on_ready=lambda: self._start_removal_thread(backup_root, to_remove, success_message, remove_mode),
            on_cancel=self._clear_backup_folder_busy
        )

    def _clear_backup_folder_busy(self):
        self._backup_folder_busy = False

    def _start_removal_thread(self, backup_root, to_remove, success_message, remove_mode):
        self.log_queue.put(f"\n--- [{log_ts()}] Removing {len(to_remove)} item(s) ({remove_mode}) ---\n")
        self.restore_status_var.set(f"Removing {len(to_remove)} item(s)...")
        threading.Thread(
            target=self._perform_removal_worker, args=(backup_root, to_remove, success_message, remove_mode), daemon=True
        ).start()

    def _perform_removal_worker(self, backup_root, to_remove, success_message, remove_mode='Both'):
        all_errors = []
        try:
            for name in to_remove:
                all_errors.extend(self._delete_backup_files_for(backup_root, name, remove_mode))
            try:
                rows2, fieldnames2 = self._read_backup_summary_csv(backup_root)
                changed = False
                if remove_mode == 'Both':
                    # Nothing left at all for this app - drop its row
                    # entirely, same as always.
                    for name in to_remove:
                        if name in rows2:
                            del rows2[name]
                            changed = True
                else:
                    # Something is still there (the piece not removed) -
                    # keep the row, but correct the field for whichever
                    # piece is now actually gone, so a future backup run's
                    # own "already good, skip" check (which reads this
                    # CSV directly, not this app's live disk-state
                    # adjustment) doesn't wrongly think that piece is
                    # still fine.
                    for name in to_remove:
                        if name not in rows2:
                            continue
                        if remove_mode == 'Raw':
                            rows2[name]['RawBackupOK'] = 'False'
                            rows2[name]['RawDeletedAfterPack'] = ''
                        elif remove_mode == 'Packed':
                            rows2[name]['PackedOK'] = 'False'
                            rows2[name]['SignedOK'] = 'False'
                            rows2[name]['PackedPath'] = ''
                        changed = True
                        # This removal might be the second of two separate,
                        # partial removals done at different times (e.g.
                        # "Packed" today, "Raw" weeks from now,
                        # on the same app) - if NEITHER piece is left after
                        # this one, the row is now fully stale and would
                        # otherwise sit there forever (nothing else ever
                        # revisits it), showing "Missing" with a Packed
                        # path pointing at a file that no longer exists.
                        if rows2[name].get('RawBackupOK') != 'True' and rows2[name].get('PackedOK') != 'True':
                            del rows2[name]
                if changed:
                    self._write_backup_summary_csv(backup_root, rows2, fieldnames2)
                elif not fieldnames2:
                    all_errors.append("backup-summary.csv not found - nothing to update there.")
            except Exception as e:
                all_errors.append(f"backup-summary.csv: {e}")
        except Exception as e:
            # Same reasoning as the verify-all worker: something genuinely
            # unexpected here must still not prevent _on_removal_finished
            # from running, since that's what resets the "removal in
            # progress" guard - an uncaught exception would otherwise
            # leave it stuck forever, permanently blocking every future
            # removal until the app is restarted.
            all_errors.append(f"Unexpected error during removal: {e}")
        self.root.after(0, lambda: self._on_removal_finished(backup_root, to_remove, all_errors, success_message))

    def _on_removal_finished(self, backup_root, to_remove, all_errors, success_message):
        self._backup_folder_busy = False
        self._release_backup_lock(backup_root)
        names_preview = ', '.join(to_remove[:10])
        if len(to_remove) > 10:
            names_preview += f", and {len(to_remove) - 10} more"
        if all_errors:
            self.log_queue.put(f"--- [{log_ts()}] Removal finished - {len(to_remove)} item(s), {len(all_errors)} issue(s) ---\n")
            self.log_queue.put(f"    Affected: {names_preview}\n")
            do_full_refresh = messagebox.askyesno(
                "Removed with some issues",
                f"Removed {len(to_remove)} item(s), but ran into:\n\n" + "\n".join(all_errors) +
                "\n\nRefresh the full list to make sure everything shown is still accurate?"
            )
        else:
            self.log_queue.put(f"--- [{log_ts()}] Removal finished - {len(to_remove)} item(s) removed, no issues ---\n")
            self.log_queue.put(f"    Affected: {names_preview}\n")
            messagebox.showinfo("Removed", success_message)
            do_full_refresh = False
        if do_full_refresh:
            self.load_restore_list(reset_filters=False)
        else:
            self._refresh_specific_restore_rows(to_remove)

    def _refresh_specific_restore_rows(self, names):
        """Updates just the given (by PackageFullName) rows already shown
        in the restore list, re-reading their current backup-summary.csv
        data and recomputing display fields - without re-running the
        installed-status PowerShell check against every row the way a
        full load_restore_list does, since an operation that's known to
        affect only these specific rows (a removal, a repack) can't have
        changed what's installed on this PC at all. InstalledStatus is
        simply kept as already known instead. A name no longer in the
        CSV at all (fully removed - e.g. a "Both" removal, or a partial
        one that happened to leave nothing behind) drops out of the
        displayed list entirely; one still there (a partial removal that
        left something behind) has its row updated in place instead."""
        backup_root = self.folder_var.get().strip()
        if not backup_root:
            return
        names_set = set(names)
        fresh_rows, _ = self._read_backup_summary_csv(backup_root)

        updated_rows = []
        for r in self.restore_list.all_rows:
            full_name = r.get('PackageFullName')
            if full_name not in names_set:
                updated_rows.append(r)
                continue
            fresh = fresh_rows.get(full_name)
            if fresh is None:
                continue  # no longer in the CSV at all - drop it from the display too
            fresh['InstalledStatus'] = r.get('InstalledStatus')
            fresh['InstalledInferred'] = r.get('InstalledInferred', False)
            self._compute_backup_display_fields(fresh, backup_root)
            updated_rows.append(fresh)

        self.restore_list.set_rows(updated_rows, preserve_checked=True)
        already = sum(1 for r in updated_rows if r.get('InstalledStatus') == 'Yes')
        self.restore_status_var.set(
            f"{len(updated_rows)} package(s) backed up - {already} installed, {len(updated_rows) - already} not installed."
        )

    def load_restore_list(self, reset_filters=True):
        backup_root = self.folder_var.get().strip()
        if not backup_root:
            messagebox.showinfo("Folder needed", "Choose a backup/restore folder first.")
            return
        self._remember_folder(backup_root)

        # Always start from a clean slate of ROWS, whatever was shown
        # before - a folder that previously had a different (or no) backup
        # shouldn't ever leave stale rows behind once this points at a
        # real backup. Clearing the search/filter state too only makes
        # sense for that same "switched to a different folder" case
        # though (reset_filters=True, the default here) - a plain refresh
        # after removing a backup (reset_filters=False) keeps whatever you
        # were searching for, since the folder itself hasn't changed.
        if reset_filters:
            self.restore_list.search_var.set('')
        self.restore_list.set_rows([], default_checked=False)
        self.restore_status_var.set("Loading...")

        self.restore_load_token += 1
        my_token = self.restore_load_token

        rows, error = load_backup_summary(backup_root)
        if error:
            messagebox.showerror("Could not load backup", error)
            self.restore_status_var.set("Failed to load - see message.")
            return
        self.restore_status_var.set(f"Loaded {len(rows)} package(s) - checking which are already installed...")
        threading.Thread(target=self._check_installed_worker, args=(rows, my_token, reset_filters), daemon=True).start()

    def _compute_backup_display_fields(self, r, backup_root):
        """Computes everything about a restore-list row that depends only
        on the backup folder's own contents - size, the Raw/PackedOK
        "still actually there?" adjustment, and the derived Type/Backup/
        Signed display fields - but deliberately NOT InstalledStatus,
        which depends on a live PowerShell query against this PC and is
        unrelated to anything a backup-folder operation (removal, a new
        backup, a repack) could itself have changed. Used both for a full
        list load (after that separate installed-check) and for a
        lightweight, targeted refresh of just a few known-affected rows
        (see _refresh_specific_restore_rows) where redoing that same
        installed-check for every remaining row would be needless, slow
        work for a large backup.

        Type reads AddedAsDependency directly (recorded once, at backup
        time, in Backup-AppxApps.ps1) rather than checking whether
        anything currently references this package - a live check would
        make Type flip from "Dependency" to "App" the moment whatever
        used to depend on it gets removed, which is confusing (it hasn't
        actually changed what this package is) and also happens to be
        exactly why it can't be used to find orphaned dependencies later
        (see _cleanup_orphaned_non_framework_dependencies) - a stable,
        historical fact doesn't have either problem."""
        full_name = r.get('PackageFullName')
        size = get_backup_size(backup_root, full_name, r.get('RawSizeBytes')) if full_name else None
        r['size_display'] = humansize(size) if size is not None else '?'
        r['SizeBytes'] = size if size is not None else 0

        # backup-summary.csv reflects what happened AT BACKUP TIME - if
        # someone deleted files from Raw\/Packed\ afterward (by hand,
        # outside this tool), that wouldn't show up unless checked
        # against the actual current disk state right now.
        if full_name:
            raw_dir = os.path.join(backup_root, 'Raw', full_name)
            if r.get('RawBackupOK') == 'True' and not os.path.isdir(raw_dir):
                # Distinguished from a genuinely missing/corrupted raw
                # folder - this one is gone on purpose ("Packed only"
                # mode, -DeleteRawAfterPack), not a problem to flag the
                # same alarming way as an actually unexpected loss. A
                # current backup already sets RawBackupOK itself to
                # false the moment Raw is actually removed this way, so
                # this specific branch is mainly a safety net for a
                # backup made before that fix, where RawBackupOK could
                # still be stuck at true.
                if str(r.get('RawDeletedAfterPack', '')).strip().lower() == 'true':
                    r['RawBackupOK'] = 'Removed (packed-only)'
                else:
                    r['RawBackupOK'] = 'MISSING NOW'
            packed_path = os.path.join(backup_root, 'Packed', f'{full_name}.appx')
            if r.get('PackedOK') == 'True' and not os.path.isfile(packed_path):
                r['PackedOK'] = 'MISSING NOW'

        # Display-only derived field - the underlying IsFramework/
        # AddedAsDependency/RawBackupOK/PackedOK fields stay exactly as
        # they are (category filtering, dependency inference, and export
        # summary all still key off those directly), this just combines
        # them into a single "Type" column.
        is_framework = str(r.get('IsFramework', '')).strip().lower() == 'true'
        added_as_dependency = str(r.get('AddedAsDependency', '')).strip().lower() == 'true'
        if is_framework:
            r['Type'] = 'Framework'
        elif added_as_dependency:
            r['Type'] = 'Dependency'
        else:
            r['Type'] = 'App'

        raw_present = r.get('RawBackupOK') == 'True'
        packed_present = r.get('PackedOK') == 'True'
        if raw_present and packed_present:
            r['BackupType'] = 'Both'
        elif raw_present:
            r['BackupType'] = 'Raw-Only'
        elif packed_present:
            r['BackupType'] = 'Packed-Only'
        else:
            r['BackupType'] = 'Missing'

        # Signing only ever applies to a packed .appx - showing
        # True/False for a raw-only backup would look like signing
        # was attempted and failed, when really there was never
        # anything to sign in the first place.
        if not packed_present:
            r['SignedOK'] = 'NA'

    def _check_installed_worker(self, rows, token, reset_filters=True):
        backup_root = self.folder_var.get().strip()
        try:
            installed = get_installed_package_fullnames(self.restore_all_users_var.get())
        except Exception as e:
            self.log_queue.put(f"ERROR checking installed apps: {e}\n")
            installed = set()
        for r in rows:
            full_name = r.get('PackageFullName')
            r['InstalledStatus'] = 'Yes' if full_name in installed else 'No'
            r['InstalledInferred'] = False
            self._compute_backup_display_fields(r, backup_root)

        # A non-framework dependency isn't always independently queryable
        # via Get-AppxPackage the same reliable way a normal app or a
        # framework package is - some resource/optional packages get
        # folded into their main app's own registration rather than
        # listed separately - which can make a dependency that's
        # genuinely present show up as "Not Installed" above. Windows
        # enforces the dependency graph though (Remove-AppxPackage
        # refuses to remove something while a package that still needs
        # it is installed), so an app that actually depends on this one
        # being installed is proof this one is too, regardless of what
        # the direct lookup found. Scoped to non-framework dependencies
        # only - frameworks are already reliably queryable on their own
        # and don't need this fallback, and this never applies to a row
        # classified as an App (Type) rather than a Dependency.
        for r in rows:
            if r.get('InstalledStatus') == 'Yes':
                continue
            if r.get('Type') != 'Dependency':
                continue
            full_name = r.get('PackageFullName')
            if not full_name:
                continue
            depended_on_by_installed_app = any(
                other is not r and full_name in (other.get('Dependencies') or '').split(';')
                and other.get('InstalledStatus') == 'Yes'
                for other in rows
            )
            if depended_on_by_installed_app:
                r['InstalledStatus'] = 'Yes'
                r['InstalledInferred'] = True

        self.root.after(0, lambda: self._on_restore_list_loaded(rows, token, reset_filters))

    def _on_restore_list_loaded(self, rows, token, reset_filters=True):
        if token != self.restore_load_token:
            return  # a newer load has since started - this one is stale, ignore it
        if reset_filters:
            # A genuinely different backup/folder just got loaded - only
            # auto-check not-installed items classified as an App (Type),
            # never frameworks or other auto-pulled dependencies.
            # Restore's own Install mode already resolves and includes
            # whatever a checked app needs via -DependencyPath, so
            # independently auto-checking a framework/dependency here
            # would be redundant at best - and, worse, since it happens
            # before any "Show:" filtering is applied, it could silently
            # check something that never becomes visible if the current
            # view happens to hide that category.
            self.restore_list.set_rows(
                rows,
                default_checked_fn=lambda r: (
                    r.get('InstalledStatus') != 'Yes' and r.get('Type') == 'App'
                )
            )
        else:
            # Just a refresh of the same backup (e.g. after removing
            # something) - keep whatever was actually checked rather than
            # silently resetting it back to the default logic as a side
            # effect of an unrelated action.
            self.restore_list.set_rows(rows, preserve_checked=True)
        self._on_restore_category_changed()  # apply whatever "Show:" is currently set to
        already = sum(1 for r in rows if r.get('InstalledStatus') == 'Yes')
        not_installed_count = len(rows) - already
        # "Selected in current view" isn't repeated here - it already
        # shows live, right next to the list itself, via _update_restore_total.
        status_line = f"{len(rows)} package(s) backed up - {already} installed, {not_installed_count} not installed."
        self.restore_status_var.set(status_line)
        if reset_filters:
            self.log_queue.put(f"\n--- [{log_ts()}] Loaded backup: {status_line} ---\n")

    def start_restore(self):
        if not os.path.isfile(RESTORE_PS1):
            messagebox.showerror("Script not found", f"Restore-AppxApps.ps1 could not be located:\n{RESTORE_PS1}")
            return
        # Checked here, before acquiring the folder lock below, rather
        # than only relying on _run_commands_async's own copy of this
        # same check - if that were the only place it happened, a lock
        # already written to disk by this method would be stranded there
        # forever, since _run_commands_worker (the only place that
        # releases it) would never actually run to release it.
        if self.current_proc is not None:
            messagebox.showinfo("Busy", "A backup or restore is already running - wait for it to finish first.")
            return
        restore_root = self.folder_var.get().strip()
        if not restore_root:
            messagebox.showinfo("Folder needed", "Choose a backup/restore folder first.")
            return
        selected = self.restore_list.get_checked_rows()
        if not selected:
            messagebox.showinfo("Nothing selected", "Load the backup contents and check at least one app first.")
            return
        mode = self.restore_mode_var.get()

        if mode == 'Install' and not messagebox.askyesno(
            "Confirm Install mode",
            "Install mode imports a self-signed certificate into Trusted People and installs the packaged "
            "apps here. If this is a different PC than the one you backed up from, make sure you trust "
            "having done that. Continue?"
        ):
            return

        full_names = [a['PackageFullName'] for a in selected]
        self._prepare_and_run(
            restore_root, full_names, "restoring them", "a restore",
            on_ready=lambda: self._start_restore_run(restore_root, mode, full_names)
        )

    def _start_restore_run(self, restore_root, mode, full_names):
        # Only set when a lock was actually acquired (i.e. the checkbox
        # was on) - when it's off, _prepare_and_run calls straight
        # through to here without ever calling _try_acquire_backup_lock,
        # so no marker file was written and there's nothing to release
        # later. Setting this unconditionally wouldn't cause a visible
        # bug today (_release_backup_lock is a safe no-op when there's
        # nothing to release), but it'd be tracking a lock that was never
        # actually taken.
        if self.check_files_locked_var.get():
            self._active_lock_root = restore_root
        include_file = self._write_include_file(full_names)
        cmd = ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', RESTORE_PS1,
               '-BackupRoot', restore_root, '-Mode', mode, '-IncludeFile', include_file]
        if mode == 'Install' and self.enable_sideloading_var.get():
            cmd.append('-EnableSideloading')

        self._append_log(f"\n=== [{log_ts()}] Starting restore ({mode}) of {len(full_names)} selected app(s) from {restore_root} ===\n")
        self._run_commands_async([cmd])

    # ---------------- Process running / log ----------------

    def _run_commands_async(self, commands, labels=None):
        """Run a list of [cmd, ...] one at a time, in order, streaming output
        for each into the log. labels, if given, is a matching list of short
        descriptions printed before each command starts."""
        if self._batch_running:
            messagebox.showinfo("Busy", "A backup or restore is already running - wait for it to finish first.")
            return
        self._cancel_requested = False
        threading.Thread(target=self._run_commands_worker, args=(commands, labels), daemon=True).start()

    def request_cancel_run(self):
        """Asks the current batch to stop after whichever item is running
        right now finishes, rather than killing it mid-way - the running
        subprocess is left alone either way, so nothing already in
        progress is interrupted partway through and potentially left in
        an inconsistent state; only the items still queued behind it are
        skipped."""
        if not self._batch_running:
            return
        self._cancel_requested = True
        self.log_queue.put(f"\n--- [{log_ts()}] Cancel requested - will stop after the current item finishes ---\n")

    def _run_commands_worker(self, commands, labels):
        """
        Each entry in commands is either a plain, already-built [cmd, ...]
        list (what start_restore passes - a single command, nothing to
        gain from deferring), or a zero-argument callable that builds and
        returns one (what start_backup passes - one call per checked app).
        The callable form matters here: building every app's command
        upfront, before the batch even starts, would bake in whatever the
        MakeAppx/SignTool path fields say at that moment as a literal
        string - so fixing a wrong path partway through a multi-app run
        would have no effect on the apps still queued behind the one that
        just failed, since their command was already fixed in place
        before the fix ever happened. Building each one right before it
        actually runs means a fix made between apps is picked up by every
        app processed after it.
        """
        total = len(commands)
        successes = []  # per-app names, from "OK: name" lines - mainly Restore's format
        failures = []   # per-app (name, reason), from "FAILED: name - reason" lines - mainly Restore's format
        cmd_ok_count = 0
        cmd_fail_labels = []  # per-COMMAND labels, tracked by exit code - works for Backup too,
                               # which never prints OK:/FAILED: text at all, unlike Restore
        self._batch_running = True
        cancelled = False
        try:
            for idx, cmd_or_builder in enumerate(commands, start=1):
                if self._cancel_requested:
                    cancelled = True
                    self.log_queue.put(
                        f"\n--- [{log_ts()}] Cancelled - {total - idx + 1} remaining item(s) skipped ---\n"
                    )
                    break
                cmd = cmd_or_builder() if callable(cmd_or_builder) else cmd_or_builder
                label = labels[idx - 1] if labels else None
                if label:
                    self.log_queue.put(f"\n--- [{log_ts()}] [{idx}/{total}] {label} ---\n")
                try:
                    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, creationflags=CREATE_NO_WINDOW)
                    self.current_proc = proc
                    for line in iter(proc.stdout.readline, b''):
                        text = decode_process_output(line)
                        self.log_queue.put(text)
                        m_ok = re.match(r'\s*OK:\s*(\S+)', text)
                        if m_ok:
                            successes.append(m_ok.group(1))
                        m_fail = re.search(r'FAILED:\s*(\S+)\s*-\s*(.+)', text)
                        if m_fail:
                            failures.append((m_fail.group(1), m_fail.group(2).strip()))
                    proc.stdout.close()
                    proc.wait()
                    self.log_queue.put(f"--- [{log_ts()}] [{idx}/{total}] finished (exit code {proc.returncode}) ---\n")
                    if label:
                        if proc.returncode == 0:
                            cmd_ok_count += 1
                        else:
                            cmd_fail_labels.append(label)
                except Exception as e:
                    self.log_queue.put(f"\nERROR running command {idx}/{total}: {e}\n")
                    if label:
                        cmd_fail_labels.append(label)
                finally:
                    self.current_proc = None
        finally:
            self._batch_running = False
            self._cancel_requested = False
        if total > 1 and not cancelled:
            self.log_queue.put(f"\n=== [{log_ts()}] All {total} operation(s) complete ===\n")

        if cmd_ok_count or cmd_fail_labels or successes or failures:
            summary = [f"\n=== Summary: {cmd_ok_count} succeeded, {len(cmd_fail_labels)} failed (of {cmd_ok_count + len(cmd_fail_labels)} run) ===\n"]
            if cmd_fail_labels:
                summary.append(f"Failed ({len(cmd_fail_labels)}):\n")
                for lbl in cmd_fail_labels:
                    summary.append(f"  FAILED: {lbl}\n")
            if successes or failures:
                # Restore's OK:/FAILED: lines give a more granular per-app
                # breakdown than the per-command one above (a single restore
                # command can cover several apps at once) - shown in
                # addition when available, not instead of.
                summary.append(f"\nPer-app detail: {len(successes)} succeeded, {len(failures)} failed\n")
                if failures:
                    for name, reason in failures:
                        # Deliberately phrased "FAILED: name - reason" here
                        # too, matching the live per-app log lines above, so
                        # the same red highlighting in _append_log applies.
                        summary.append(f"  FAILED: {name} - {reason}\n")
            self.log_queue.put(''.join(summary))

        # Only ever set by start_restore (never start_backup), and only
        # when the "check files aren't locked" option was on - releasing
        # it here, in the one worker shared by both backup and restore
        # runs, covers the run regardless of how many commands it took
        # rather than needing each individual command's completion to
        # know whether it was the last one.
        if self._active_lock_root:
            self._release_backup_lock(self._active_lock_root)
            self._active_lock_root = None

    def _configure_log_tags(self, widget):
        widget.tag_configure('ok_line', foreground='#0a7d0a')
        widget.tag_configure('fail_line', foreground='#c00000', font=('Consolas', 9, 'bold'))
        widget.tag_configure('warn_line', foreground='#a15c00')
        widget.tag_configure('header_line', foreground='#1a3d6d', font=('Consolas', 9, 'bold'), spacing1=6)
        widget.tag_configure('summary_header', foreground='#1a3d6d', font=('Consolas', 9, 'bold'),
                              background='#eef3fa', spacing1=8, spacing3=2)

    def _classify_log_line(self, text):
        """Which visual style a log line gets, checked in priority order
        since a line can match more than one pattern (a warning line often
        also contains "FAILED:", for instance)."""
        if 'Summary:' in text and text.strip().startswith('==='):
            return 'summary_header'
        if re.match(r'\s*OK[\s:(]', text):
            return 'ok_line'
        if 'FAILED:' in text:
            return 'fail_line'
        if re.search(r'\b(ADVERTENCIA|WARNING)\b', text):
            return 'warn_line'
        stripped = text.strip()
        if stripped.startswith('---') or stripped.startswith('==='):
            return 'header_line'
        return None

    def _append_log(self, text):
        self._write_log_file(text)
        tag = self._classify_log_line(text)

        if self.log_detached and self.detached_log_text is not None:
            target = self.detached_log_text
            if tag:
                target.insert('end', text, tag)
            else:
                target.insert('end', text)
            target.see('end')
        else:
            self.log_text.configure(state='normal')
            if tag:
                self.log_text.insert('end', text, tag)
            else:
                self.log_text.insert('end', text)
            self.log_text.see('end')
            self.log_text.configure(state='disabled')

    def _poll_log_queue(self):
        try:
            while True:
                text = self.log_queue.get_nowait()
                self._append_log(text)
        except queue.Empty:
            pass
        self.root.after(150, self._poll_log_queue)


def main():
    if not is_admin():
        tmp = tk.Tk()
        tmp.withdraw()
        proceed = messagebox.askyesno(
            "Administrator required",
            "This tool needs to run as Administrator (it launches PowerShell scripts that read "
            "backup-protected folders and install certificates).\n\nRestart elevated now?"
        )
        tmp.destroy()
        if proceed:
            relaunch_as_admin()
        sys.exit(0)

    root = tk.Tk()
    AppxBackupGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()
