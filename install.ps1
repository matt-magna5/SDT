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
    [switch] $NoLaunch,
    [switch] $Force
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
if ($existing -eq $Version -and -not $Force) {
    Say "Already at $Version - nothing to download." DarkGreen
    Say "(Use -Force to re-download the tag even if the version string matches.)" DarkGray
} else {
    if ($Force -and $existing -eq $Version) {
        Say "Force re-download requested - downloading $Version fresh..." DarkYellow
    }
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
if (-not (Test-Path $pyExe)) {
    Say "[X] Portable Python missing. Hypervisor scan will fail." Red
    Say "  Expected: $pyExe" DarkGray
} else {
    Say "Python found: $pyExe" DarkGray

    # === BELT-AND-SUSPENDERS: Python dep setup can never terminate the install ===
    # Multiple layers:
    #   1. Force local $ErrorActionPreference=Continue so native stderr under
    #      $global ='Stop' cannot halt us.
    #   2. All helper calls wrapped in try/catch.
    #   3. Skip ensurepip entirely (embeddable Python never has it).
    #   4. get-pip.py attempted twice (bootstrap.pypa.io then pypa/get-pip GitHub mirror).
    #   5. Final state reported but NEVER throws - install always finishes.
    $prevEAP = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {

    # --- Helpers: silent, non-throwing python probes ---
    # Use call-operator (&) rather than Start-Process: ExitCode is reliable,
    # no Start-Process quoting/handle-close quirks that gave false negatives
    # on quick python -c checks. Safe here because the outer try/finally has
    # already set $ErrorActionPreference='Continue'.
    function Test-PyImport([string]$py, [string]$module) {
        try {
            $null = & $py -c "import $module" 2>&1
            return ($LASTEXITCODE -eq 0)
        } catch { return $false }
    }
    function Test-PyPipVersion([string]$py) {
        try {
            $out = & $py -m pip --version 2>&1 | Out-String
            return ($LASTEXITCODE -eq 0 -and $out -match 'pip \d')
        } catch { return $false }
    }

    # Step 1: ensure 'import site' in ._pth (embeddable Python disables it by default)
    $pthFile = Get-ChildItem $PyDir -Filter 'python*._pth' -EA 0 | Select-Object -First 1
    if ($pthFile) {
        $pthContent = Get-Content $pthFile.FullName -Raw
        if ($pthContent -match '(?m)^\s*#\s*import\s+site\s*$') {
            ($pthContent -replace '(?m)^\s*#\s*import\s+site\s*$', 'import site') | Set-Content $pthFile.FullName -Encoding ASCII
            Say "Enabled 'import site' in $($pthFile.Name)" DarkGray
        } else {
            Say "'import site' already enabled in $($pthFile.Name)" DarkGray
        }
    } else {
        Say "No python*._pth found in $PyDir - skipping site-packages enable" DarkYellow
    }

    # Step 2: bootstrap pip via get-pip.py. We SKIP ensurepip entirely -
    # embeddable Python never has it, and under $ErrorActionPreference='Stop'
    # its "No module named ensurepip" stderr becomes a NativeCommandError
    # that killed older installs.
    if (Test-PyPipVersion $pyExe) {
        Say "pip already present." DarkGreen
    } else {
        Say "Bootstrapping pip via get-pip.py..." DarkCyan
        $getPipPy = Join-Path $PyDir 'get-pip.py'
        $havePip  = $false
        # Source 1 (preferred): bundled get-pip.py shipped with the SDT repo -
        # zero network dependency, works behind any proxy / ThreatLocker.
        $bundled = Join-Path $AppDir 'get-pip.py'
        if (Test-Path $bundled) {
            try {
                Copy-Item -Path $bundled -Destination $getPipPy -Force
                $havePip = $true
                Say "Using bundled get-pip.py ($([int]((Get-Item $getPipPy).Length/1024)) KB)" DarkGray
            } catch {
                Say "Bundled get-pip.py copy failed: $($_.Exception.Message)" DarkYellow
            }
        }
        # Source 2+: network mirrors (only if bundled missing or copy failed)
        if (-not $havePip) {
            $mirrors = @(
                'https://bootstrap.pypa.io/get-pip.py',
                'https://raw.githubusercontent.com/pypa/get-pip/main/public/get-pip.py',
                'https://github.com/matt-magna5/SDT/raw/main/get-pip.py'
            )
            foreach ($url in $mirrors) {
                try {
                    Invoke-WebRequest -Uri $url -OutFile $getPipPy -UseBasicParsing -TimeoutSec 60 -EA Stop
                    if ((Test-Path $getPipPy) -and (Get-Item $getPipPy).Length -gt 100000) {
                        $havePip = $true
                        Say "Downloaded get-pip.py from $url ($([int]((Get-Item $getPipPy).Length/1024)) KB)" DarkGray
                        break
                    }
                } catch {
                    Say "Mirror failed ($url): $($_.Exception.Message)" DarkYellow
                }
            }
        }
        if (-not $havePip) {
            Say "[X] Could not obtain get-pip.py (bundled + $(($mirrors|Measure).Count) mirror(s) all failed)." Red
            Say "    Network/proxy/ThreatLocker blocking outbound HTTPS?" DarkYellow
        } else {
            Say "Running get-pip.py (up to 90s; hang = python.exe blocked from pypi.org)..." DarkCyan
            $logFile = Join-Path $env:TEMP "sdt-getpip-$([guid]::NewGuid().ToString('N').Substring(0,6)).log"
            try {
                $proc = Start-Process -FilePath $pyExe -ArgumentList @($getPipPy, '--disable-pip-version-check') -NoNewWindow -PassThru -RedirectStandardOutput $logFile -RedirectStandardError "$logFile.err" -EA SilentlyContinue
                $finished = if ($proc) { $proc.WaitForExit(90000) } else { $false }
                if (-not $finished -and $proc) {
                    try { $proc.Kill() } catch { }
                    Say "[X] get-pip.py timed out after 90s - python.exe blocked from pypi.org" Red
                    Say "   Whitelist python.exe outbound to: files.pythonhosted.org, pypi.org, bootstrap.pypa.io" DarkYellow
                    Say "   Path: $pyExe" DarkGray
                } else {
                    $getPipLog = ''
                    if (Test-Path $logFile) { $getPipLog = Get-Content $logFile -Raw -EA 0 }
                    if (Test-Path "$logFile.err") { $getPipLog += "`n--- STDERR ---`n" + (Get-Content "$logFile.err" -Raw -EA 0) }
                    if (Test-PyPipVersion $pyExe) {
                        Say "[OK] pip bootstrapped via get-pip.py." DarkGreen
                    } else {
                        Say "[X] get-pip.py ran but pip still not working:" Red
                        Say $getPipLog Yellow
                    }
                }
            } catch {
                Say "[X] get-pip.py execution failed: $($_.Exception.Message)" Red
            } finally {
                Remove-Item $logFile, "$logFile.err" -EA 0
            }
            Remove-Item $getPipPy -EA 0
        }
    }

    # Step 3: pip install the required packages.
    # Trust pip's own output over exit-code fiddling: "Successfully installed"
    # or "Requirement already satisfied" are the only signals that matter.
    if (Test-PyPipVersion $pyExe) {
        Say "Installing Python deps (pyVmomi, requests, urllib3)..." DarkCyan
        $pipLog = Join-Path $env:TEMP "sdt-pipinst-$([guid]::NewGuid().ToString('N').Substring(0,6)).log"
        $proc = Start-Process -FilePath $pyExe -ArgumentList @('-m','pip','install','--disable-pip-version-check','pyVmomi','requests','urllib3') -NoNewWindow -PassThru -RedirectStandardOutput $pipLog -RedirectStandardError "$pipLog.err" -EA SilentlyContinue
        $finished = if ($proc) { $proc.WaitForExit(120000) } else { $false }
        $pipOut = ''
        if (Test-Path $pipLog)       { $pipOut = Get-Content $pipLog -Raw -EA 0 }
        if (Test-Path "$pipLog.err") { $pipOut += "`n" + (Get-Content "$pipLog.err" -Raw -EA 0) }
        Remove-Item $pipLog, "$pipLog.err" -EA 0
        if (-not $finished) {
            try { $proc.Kill() } catch { }
            Say "[X] pip install timed out after 120s." Red
        } elseif ($pipOut -match '(?m)Successfully installed|Requirement already satisfied') {
            Say "[OK] Python deps installed." DarkGreen
        } else {
            Say "[X] pip install did not complete cleanly:" Red
            Say $pipOut Yellow
        }
    } else {
        Say "[X] pip not available - skipping deps install. Manual fix:" Red
        Say "    & '$pyExe' '<repo>\get-pip.py'" DarkYellow
        Say "    & '$pyExe' -m pip install pyVmomi requests urllib3" DarkYellow
    }

    # === FINAL PRE-FLIGHT: import verification is the real signal ===
    Say "" DarkGray
    Say "--- Python dep pre-flight ---" DarkCyan
    $checks = @{
        'python.exe'      = (Test-Path $pyExe)
        'pip module'      = (Test-PyPipVersion $pyExe)
        'import requests' = (Test-PyImport $pyExe 'requests')
        'import pyVmomi'  = (Test-PyImport $pyExe 'pyVmomi')
        'import urllib3'  = (Test-PyImport $pyExe 'urllib3')
    }
    $allOk = $true
    foreach ($k in 'python.exe','pip module','import requests','import pyVmomi','import urllib3') {
        if ($checks[$k]) { Say ("  [OK]  {0}" -f $k) DarkGreen }
        else             { Say ("  [X]   {0}" -f $k) Red; $allOk = $false }
    }
    if ($allOk) {
        Say "[OK] All Python deps verified - hypervisor scan will work." Green
    } else {
        Say "[!] Some deps missing. If the hypervisor scan later works, the pre-flight check was over-strict and you can ignore this." Yellow
    }

    } finally {
        $ErrorActionPreference = $prevEAP
    }
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
            & `$sb -NoLaunch -Force
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
