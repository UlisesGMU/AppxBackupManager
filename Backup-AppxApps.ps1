<#
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
