<#
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
