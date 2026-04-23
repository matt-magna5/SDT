<#
.SYNOPSIS
    SDT installer + updater - sets up a permanent 'sdt' command that
    auto-updates itself on every launch.

.DESCRIPTION
    Installs to %LOCALAPPDATA%\Magna5\SDT\ (no admin required).
    Adds %LOCALAPPDATA%\Magna5\SDT\bin\ to the current user's PATH.
    Creates 'sdt' as a shim command that:
      - checks GitHub for a newer release every launch
      - auto-updates the app folder in place (preserving portable Python + pip packages)
      - dispatches subcommands: sdt | sdt invoke | sdt cli | sdt update | sdt uninstall
    Also installs portable Python 3.12 and pip-installs pyVmomi + requests so the
    hypervisor scan works out of the box.

.NOTES
    One-time install one-liner:
        [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; iwr https://raw.githubusercontent.com/matt-magna5/SDT/main/install.ps1 -UseBasicParsing | iex

    After install, in any new terminal:
        sdt           -> browser GUI (default)
        sdt invoke    -> single-host Invoke-ServerDiscovery (run on the target)
        sdt cli       -> legacy console Start-DiscoverySession wizard
        sdt update    -> force-refresh from GitHub now
        sdt uninstall -> remove the install
        sdt version   -> show paths + installed version
#>
[CmdletBinding()]
param(
    [string] $Version = 'latest',
    [switch] $Quiet,
    [switch] $NoLaunch
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

function Say([string]$m, [string]$c='White') { if (-not $Quiet) { Write-Host "  $m" -ForegroundColor $c } }

Say ""
Say "================================================================" DarkMagenta
Say "  MAGNA5 SDT - installer & auto-updater" Magenta
Say "================================================================" DarkMagenta
Say ""

# ----- Paths -----------------------------------------------------------------
$Root    = Join-Path $env:LOCALAPPDATA 'Magna5\SDT'
$BinDir  = Join-Path $Root 'bin'
$AppDir  = Join-Path $Root 'app'
$PyDir   = Join-Path $AppDir 'python'
$VerFile = Join-Path $Root 'VERSION'
$ShimPS  = Join-Path $BinDir 'sdt.ps1'
$ShimCmd = Join-Path $BinDir 'sdt.cmd'

foreach ($d in @($Root, $BinDir, $AppDir)) {
    New-Item -ItemType Directory -Force -Path $d | Out-Null
}

# ----- Resolve latest tag ----------------------------------------------------
if ($Version -eq 'latest') {
    Say "Resolving latest release..." DarkCyan
    $Version = $null
    try {
        $rel = Invoke-WebRequest 'https://api.github.com/repos/matt-magna5/SDT/releases?per_page=20' -UseBasicParsing -TimeoutSec 10
        $releases = $rel.Content | ConvertFrom-Json
        # Prefer v4+ / alpha / beta (GUI-capable), else newest overall
        $guiTag = $releases | Where-Object { $_.tag_name -match '^v4' -or $_.tag_name -match 'alpha|beta|rc' } | Select-Object -First 1
        $Version = if ($guiTag) { $guiTag.tag_name } else { $releases[0].tag_name }
    } catch {
        Say "Releases API failed: $($_.Exception.Message)" Yellow
        foreach ($try in @('v4.0-alpha','v3.11','v3.10')) {
            try {
                $h = Invoke-WebRequest "https://github.com/matt-magna5/SDT/archive/refs/tags/$try.zip" -Method Head -UseBasicParsing -TimeoutSec 5
                if ($h.StatusCode -eq 200) { $Version = $try; break }
            } catch { continue }
        }
        if (-not $Version) { throw "Could not determine SDT version" }
    }
    Say "Latest: $Version" DarkGreen
}

# ----- Read existing version -------------------------------------------------
$existing = if (Test-Path $VerFile) { (Get-Content $VerFile -Raw -EA 0).Trim() } else { '' }
if ($existing -eq $Version) {
    Say "Already at $Version - nothing to download." DarkGreen
} else {
    # ----- Download + extract ------------------------------------------------
    $url    = "https://github.com/matt-magna5/SDT/archive/refs/tags/$Version.zip"
    $zipTmp = Join-Path $env:TEMP "sdt-install-$Version.zip"
    $extTmp = Join-Path $env:TEMP ("sdt-install-" + [guid]::NewGuid().ToString('N').Substring(0,8))
    Say "Downloading $Version ..." DarkCyan
    Invoke-WebRequest -Uri $url -OutFile $zipTmp -UseBasicParsing -TimeoutSec 120
    New-Item -ItemType Directory -Force -Path $extTmp | Out-Null
    Expand-Archive -Path $zipTmp -DestinationPath $extTmp -Force
    Remove-Item $zipTmp -Force -EA 0
    $src = Get-ChildItem $extTmp -Directory | Select-Object -First 1
    if (-not $src) { throw "Extraction produced no folder" }

    # ----- Copy release files into AppDir (preserve python/ + _output/) -----
    Say "Installing to $AppDir ..." DarkCyan
    # Delete old files but skip preserved folders
    $preserve = @('python','_output','_archive')
    Get-ChildItem $AppDir -Force -EA 0 | Where-Object { $_.Name -notin $preserve } | ForEach-Object {
        Remove-Item $_.FullName -Recurse -Force -EA 0
    }
    # Copy everything from the extracted release
    Copy-Item -Path (Join-Path $src.FullName '*') -Destination $AppDir -Recurse -Force
    Remove-Item $extTmp -Recurse -Force -EA 0

    # ----- Stamp version -----------------------------------------------------
    Set-Content -Path $VerFile -Value $Version -Encoding ASCII
    Say "Installed $Version" DarkGreen
}

# ----- Portable Python + pip deps -------------------------------------------
$pyExe = Join-Path $PyDir 'python.exe'
if (-not (Test-Path $pyExe)) {
    Say "Fetching portable Python 3.12 + plink (~10 MB)..." DarkCyan
    $getPy = Join-Path $AppDir 'Get-PortablePython.ps1'
    if (Test-Path $getPy) {
        try { Push-Location $AppDir; & $getPy | Out-Null } catch { Say "Portable Python fetch failed: $($_.Exception.Message)" Yellow } finally { Pop-Location }
    }
}
if (Test-Path $pyExe) {
    # Install required packages for the hypervisor scan (idempotent)
    Say "Ensuring Python deps (pyVmomi, requests)..." DarkCyan
    try {
        $pipOut = & $pyExe -m pip install --quiet --disable-pip-version-check pyVmomi requests urllib3 2>&1 | Out-String
        Say "Python deps ready." DarkGreen
    } catch {
        Say "pip install failed (non-fatal): $($_.Exception.Message)" Yellow
    }
} else {
    Say "Portable Python not available - hypervisor scan will need system Python on PATH." Yellow
}

# ----- Write sdt.ps1 shim ----------------------------------------------------
$shimBody = @"
# SDT launcher shim - auto-generated by install.ps1
# Checks GitHub for a newer release on every launch, self-updates in place,
# then dispatches the requested subcommand.
param(
    [Parameter(Position=0)][string]`$Mode = 'gui',
    [Parameter(ValueFromRemainingArguments=`$true)]`$Rest
)

`$Root    = '$Root'
`$AppDir  = '$AppDir'
`$VerFile = '$VerFile'

`$ErrorActionPreference = 'Continue'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

# ---- Loud auto-update - always prints status on every launch ----
function Invoke-SdtAutoUpdate {
    `$local = if (Test-Path `$VerFile) { (Get-Content `$VerFile -Raw -EA 0).Trim() } else { '' }
    if (`$env:SDT_NO_AUTOUPDATE -eq '1') {
        Write-Host ("  [sdt] update check skipped (SDT_NO_AUTOUPDATE=1, local v{0})" -f `$local) -ForegroundColor DarkGray
        return
    }
    Write-Host ("  [sdt] checking GitHub for newer release (local v{0})..." -f `$local) -ForegroundColor DarkCyan
    `$latest = `$null
    try {
        `$ProgressPreference = 'SilentlyContinue'
        `$rel = Invoke-WebRequest 'https://api.github.com/repos/matt-magna5/SDT/releases?per_page=20' -UseBasicParsing -TimeoutSec 6 -EA Stop
        `$releases = `$rel.Content | ConvertFrom-Json
        `$guiTag = `$releases | Where-Object { `$_.tag_name -match '^v4' -or `$_.tag_name -match 'alpha|beta|rc' } | Select-Object -First 1
        `$latest = if (`$guiTag) { `$guiTag.tag_name } else { `$releases[0].tag_name }
    } catch {
        Write-Host ("  [sdt] update check failed: {0} - running local v{1}" -f `$_.Exception.Message, `$local) -ForegroundColor DarkYellow
        return
    }
    if (-not `$latest) {
        Write-Host "  [sdt] couldn't determine latest release - running local" -ForegroundColor DarkYellow
        return
    }
    if (`$latest -eq `$local) {
        Write-Host ("  [sdt] up to date (v{0})" -f `$local) -ForegroundColor DarkGreen
        return
    }
    Write-Host ("  [sdt] v{0} available (local v{1}) - updating..." -f `$latest, `$local) -ForegroundColor Yellow
    try {
        `$inst = Invoke-WebRequest 'https://raw.githubusercontent.com/matt-magna5/SDT/main/install.ps1' -UseBasicParsing -TimeoutSec 30
        `$sb = [ScriptBlock]::Create(`$inst.Content)
        & `$sb -Version `$latest -Quiet -NoLaunch | Out-Null
        Write-Host ("  [sdt] updated to v{0}" -f `$latest) -ForegroundColor Green
    } catch {
        Write-Host ("  [sdt] update failed: {0} - running local v{1}" -f `$_.Exception.Message, `$local) -ForegroundColor Red
    }
}

# ---- Dispatch ----
switch -Regex (`$Mode) {
    '^(version|-v|--version)$' {
        `$v = if (Test-Path `$VerFile) { Get-Content `$VerFile -Raw -EA 0 } else { 'unknown' }
        Write-Host "SDT installed at: `$AppDir"
        Write-Host "Version: `$(`$v.Trim())"
        return
    }
    '^(update|upgrade)$' {
        Write-Host "Forcing update..." -ForegroundColor Cyan
        try {
            `$inst = Invoke-WebRequest 'https://raw.githubusercontent.com/matt-magna5/SDT/main/install.ps1' -UseBasicParsing -TimeoutSec 30
            `$sb = [ScriptBlock]::Create(`$inst.Content)
            & `$sb -NoLaunch
        } catch { Write-Host "Update failed: `$(`$_.Exception.Message)" -ForegroundColor Red }
        return
    }
    '^(uninstall|remove)$' {
        `$path = Split-Path `$AppDir -Parent
        Write-Host "Removing `$path ..." -ForegroundColor Yellow
        Remove-Item `$path -Recurse -Force -EA 0
        # Clean PATH
        `$userPath = [Environment]::GetEnvironmentVariable('Path','User')
        `$new = (`$userPath -split ';') | Where-Object { `$_ -and `$_ -notmatch 'Magna5\\SDT\\bin' } | ForEach-Object { `$_ }
        [Environment]::SetEnvironmentVariable('Path', (`$new -join ';'), 'User')
        Write-Host "Uninstalled." -ForegroundColor Green
        return
    }
}

Invoke-SdtAutoUpdate

# Normalize `$Rest so empty splat is safe on PS 5.1 + PS 7
if (`$null -eq `$Rest) { `$Rest = @() }
`$RestArr = @(`$Rest)

switch -Regex (`$Mode) {
    '^(cli|console|tui)$'   {
        if (`$RestArr.Count -gt 0) { & (Join-Path `$AppDir 'Start-DiscoverySession.ps1') @RestArr }
        else { & (Join-Path `$AppDir 'Start-DiscoverySession.ps1') }
        return
    }
    '^(invoke|local|bare)$' {
        if (`$RestArr.Count -gt 0) { & (Join-Path `$AppDir 'Invoke-ServerDiscovery.ps1') @RestArr }
        else { & (Join-Path `$AppDir 'Invoke-ServerDiscovery.ps1') }
        return
    }
    default {
        if (`$RestArr.Count -gt 0) { & (Join-Path `$AppDir 'Start-DiscoverySessionGUI.ps1') @RestArr }
        else { & (Join-Path `$AppDir 'Start-DiscoverySessionGUI.ps1') }
        return
    }
}
"@
Set-Content -Path $ShimPS -Value $shimBody -Encoding UTF8

# ----- Write sdt.cmd (CMD wrapper) ------------------------------------------
@"
@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$ShimPS" %*
"@ | Set-Content -Path $ShimCmd -Encoding ASCII

Say "Shim scripts written." DarkGreen

# ----- PATH (user scope) -----------------------------------------------------
$userPath = [Environment]::GetEnvironmentVariable('Path','User')
if (-not $userPath) { $userPath = '' }
$already  = ($userPath -split ';') | Where-Object { $_ -ieq $BinDir }
if (-not $already) {
    $newPath = if ($userPath) { "$userPath;$BinDir" } else { $BinDir }
    [Environment]::SetEnvironmentVariable('Path', $newPath, 'User')
    # Also make it available in THIS session so we can launch immediately
    $env:Path = $env:Path + ';' + $BinDir
    Say "Added $BinDir to user PATH." DarkGreen
} else {
    Say "PATH already contains $BinDir" DarkGray
}

Say ""
Say "================================================================" DarkMagenta
Say "  Install complete - version $Version" Green
Say "================================================================" DarkMagenta
Say ""
Say "Commands (in ANY new terminal):" Cyan
Say "  sdt             browser GUI (default)" DarkGray
Say "  sdt invoke      local per-host Invoke-ServerDiscovery" DarkGray
Say "  sdt cli         legacy console session wizard" DarkGray
Say "  sdt update      force re-install from GitHub" DarkGray
Say "  sdt uninstall   remove everything" DarkGray
Say "  sdt version     show install info" DarkGray
Say ""
Say "Auto-update runs on every 'sdt' launch - you'll never paste a" DarkGray
Say "download one-liner again." DarkGray
Say ""

# ----- Launch GUI immediately so user sees something ------------------------
if (-not $NoLaunch) {
    Say "Launching GUI now..." Yellow
    & $ShimPS
}
