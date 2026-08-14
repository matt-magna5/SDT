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

# -- PowerShell version check --------------------------------------------------
$script:PSMaj = $PSVersionTable.PSVersion.Major
if ($script:PSMaj -lt 3) {
    Write-Host ""
    Write-Host "  [X] PowerShell $($PSVersionTable.PSVersion) is not supported." -ForegroundColor Red
    Write-Host "      Minimum required: PS 3.0  |  Recommended: PS 5.1 or PS 7+" -ForegroundColor DarkRed
    Write-Host "      Download PS 7: https://aka.ms/powershell" -ForegroundColor DarkGray
    exit 1
}
if ($script:PSMaj -eq 3 -or $script:PSMaj -eq 4) {
    Write-Host ""
    Write-Host "  [!] PowerShell $($PSVersionTable.PSVersion) detected -- using .NET ZipFile fallback (Expand-Archive not available)." -ForegroundColor Yellow
    Write-Host "      Everything will work. Recommend upgrading to PS 5.1+ eventually." -ForegroundColor DarkYellow
    Write-Host ""
} elseif ($script:PSMaj -ge 7) {
    Write-Host "  [OK] PowerShell $($PSVersionTable.PSVersion) -- full compatibility." -ForegroundColor DarkGreen
} else {
    Write-Host "  [OK] PowerShell $($PSVersionTable.PSVersion) -- compatible." -ForegroundColor DarkGreen
}

function Say([string]$m, [string]$c='White') { if (-not $Quiet) { Write-Host "  $m" -ForegroundColor $c } }

function Expand-ZipCompat([string]$ZipPath, [string]$Destination) {
    if (-not (Test-Path $Destination)) { New-Item -ItemType Directory -Force -Path $Destination | Out-Null }
    if (Get-Command Expand-Archive -ErrorAction SilentlyContinue) {
        Expand-Archive -Path $ZipPath -DestinationPath $Destination -Force
    } else {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $Destination)
    }
}

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
# Query /tags FIRST (returns ALL tags including those without formal Release
# objects). Falls back to /releases if /tags fails. Previously we only
# queried /releases which meant tag-only updates (no Release object) were
# invisible to auto-discovery and the install would pin to the last formally
# released version. /tags is the authoritative source.
if ($Version -eq 'latest') {
    Say "Resolving latest tag..." DarkCyan
    $Version = $null
    try {
        $tagsResp = Invoke-WebRequest 'https://api.github.com/repos/matt-magna5/SDT/tags?per_page=50' -UseBasicParsing -TimeoutSec 10
        $tags = ($tagsResp.Content | ConvertFrom-Json) | ForEach-Object { $_.name }
        # Filter to semver-ish v4.x.y first (GUI-capable era), then newest by
        # semver-aware sort. Fall back to whatever the API returned in order.
        $semverTags = $tags | Where-Object { $_ -match '^v(\d+)\.(\d+)\.(\d+)(?:-(.+))?$' }
        $v4Tags = $semverTags | Where-Object { $_ -match '^v4' }
        if ($v4Tags) {
            # Sort v4.x.y by numeric components (descending)
            $sorted = $v4Tags | Sort-Object @{
                Expression = {
                    if ($_ -match '^v(\d+)\.(\d+)\.(\d+)') {
                        [int]$matches[1] * 1000000 + [int]$matches[2] * 1000 + [int]$matches[3]
                    } else { 0 }
                }
                Descending = $true
            }
            $Version = $sorted | Select-Object -First 1
        } else {
            $Version = $tags | Select-Object -First 1
        }
    } catch {
        Say "Tags API failed: $($_.Exception.Message). Falling back to /releases..." Yellow
        try {
            $rel = Invoke-WebRequest 'https://api.github.com/repos/matt-magna5/SDT/releases?per_page=20' -UseBasicParsing -TimeoutSec 10
            $releases = $rel.Content | ConvertFrom-Json
            $guiTag = $releases | Where-Object { $_.tag_name -match '^v4' -or $_.tag_name -match 'alpha|beta|rc' } | Select-Object -First 1
            $Version = if ($guiTag) { $guiTag.tag_name } else { $releases[0].tag_name }
        } catch {
            Say "Releases API also failed. Probing known tags..." Yellow
            foreach ($try in @('v4.1.13','v4.1.12','v4.1.11','v4.1.10','v4.1.9','v4.1.8')) {
                try {
                    $h = Invoke-WebRequest "https://github.com/matt-magna5/SDT/archive/refs/tags/$try.zip" -Method Head -UseBasicParsing -TimeoutSec 5
                    if ($h.StatusCode -eq 200) { $Version = $try; break }
                } catch { continue }
            }
            if (-not $Version) { throw "Could not determine SDT version" }
        }
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
    Expand-ZipCompat $zipTmp $extTmp
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

# ---- Self-heal: parse every PS1 in AppDir. If any has a syntax error,
# the install is broken and we MUST update before launching anything. ----
function Test-SdtHealth {
    `$ps1s = Get-ChildItem `$AppDir -Filter '*.ps1' -EA SilentlyContinue
    if (-not `$ps1s) { return @{ ok=`$false; reason='no .ps1 files in app dir' } }
    foreach (`$f in `$ps1s) {
        `$errs = `$null
        try {
            [System.Management.Automation.Language.Parser]::ParseFile(`$f.FullName, [ref]`$null, [ref]`$errs) | Out-Null
            if (`$errs -and `$errs.Count -gt 0) {
                return @{ ok=`$false; reason="parse error in `$(`$f.Name) at line `$(`$errs[0].Extent.StartLineNumber): `$(`$errs[0].Message)" }
            }
        } catch {
            return @{ ok=`$false; reason="cannot parse `$(`$f.Name): `$(`$_.Exception.Message)" }
        }
    }
    return @{ ok=`$true }
}

function Invoke-SdtRepair {
    Write-Host "  [sdt] FORCE REPAIR - re-downloading from main..." -ForegroundColor Yellow
    try {
        `$inst = Invoke-WebRequest 'https://raw.githubusercontent.com/matt-magna5/SDT/main/install.ps1' -UseBasicParsing -TimeoutSec 30
        `$sb = [ScriptBlock]::Create(`$inst.Content)
        & `$sb -NoLaunch -Force
        Write-Host "  [sdt] repair complete" -ForegroundColor Green
        return `$true
    } catch {
        Write-Host "  [sdt] repair failed: `$(`$_.Exception.Message)" -ForegroundColor Red
        return `$false
    }
}

# ---- Loud auto-update - always prints status on every launch ----
function Invoke-SdtAutoUpdate {
    `$localRaw = if (Test-Path `$VerFile) { (Get-Content `$VerFile -Raw -EA 0).Trim() } else { '' }
    # `$local is the bare version for display (format strings add 'v').
    `$local = `$localRaw
    if (`$local -match '^v') { `$local = `$local.Substring(1) }
    if (`$env:SDT_NO_AUTOUPDATE -eq '1') {
        Write-Host ("  [sdt] update check skipped (SDT_NO_AUTOUPDATE=1, local v{0})" -f `$local) -ForegroundColor DarkGray
        return
    }
    Write-Host ("  [sdt] checking GitHub for newer tag (local v{0})..." -f `$local) -ForegroundColor DarkCyan
    `$latest = `$null
    try {
        `$ProgressPreference = 'SilentlyContinue'
        # Use /tags (includes tag-only updates without formal Release objects).
        `$tagsResp = Invoke-WebRequest 'https://api.github.com/repos/matt-magna5/SDT/tags?per_page=50' -UseBasicParsing -TimeoutSec 6 -EA Stop
        `$tags = (`$tagsResp.Content | ConvertFrom-Json) | ForEach-Object { `$_.name }
        `$v4Tags = `$tags | Where-Object { `$_ -match '^v4\.(\d+)\.(\d+)' }
        if (`$v4Tags) {
            `$sorted = `$v4Tags | Sort-Object @{
                Expression = {
                    if (`$_ -match '^v(\d+)\.(\d+)\.(\d+)') {
                        [int]`$matches[1] * 1000000 + [int]`$matches[2] * 1000 + [int]`$matches[3]
                    } else { 0 }
                }
                Descending = `$true
            }
            `$latest = `$sorted | Select-Object -First 1
        } else {
            `$latest = `$tags | Select-Object -First 1
        }
    } catch {
        Write-Host ("  [sdt] update check failed: {0} - running local v{1}" -f `$_.Exception.Message, `$local) -ForegroundColor DarkYellow
        return
    }
    if (-not `$latest) {
        Write-Host "  [sdt] couldn't determine latest release - running local" -ForegroundColor DarkYellow
        return
    }
    # Normalize BOTH sides before comparison: 'v4.2.11' vs '4.2.11' must match.
    `$latestDisplay = `$latest; if (`$latestDisplay -match '^v') { `$latestDisplay = `$latestDisplay.Substring(1) }
    if (`$latestDisplay -eq `$local) {
        Write-Host ("  [sdt] up to date (v{0})" -f `$local) -ForegroundColor DarkGreen
        return
    }
    Write-Host ("  [sdt] v{0} available (local v{1}) - updating..." -f `$latestDisplay, `$local) -ForegroundColor Yellow
    try {
        `$inst = Invoke-WebRequest 'https://raw.githubusercontent.com/matt-magna5/SDT/main/install.ps1' -UseBasicParsing -TimeoutSec 30
        `$sb = [ScriptBlock]::Create(`$inst.Content)
        & `$sb -Version `$latest -Quiet -NoLaunch | Out-Null
        Write-Host ("  [sdt] updated to v{0} - relaunching" -f `$latestDisplay) -ForegroundColor Green
        return 'updated'
    } catch {
        Write-Host ("  [sdt] update failed: {0} - running local v{1}" -f `$_.Exception.Message, `$local) -ForegroundColor Red
    }
}

# ---- Dispatch ----
switch -Regex (`$Mode) {
    '^(help|-h|--help|/\?|\?)$' {
        Write-Host ""
        Write-Host "SDT - Magna5 Server Discovery Tool" -ForegroundColor Cyan
        Write-Host "==================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "USAGE:" -ForegroundColor Yellow
        Write-Host "  sdt [<command>] [args...]"
        Write-Host ""
        Write-Host "COMMANDS:" -ForegroundColor Yellow
        Write-Host "  (no args)            Launch the browser GUI (default)"
        Write-Host "  gui                  Launch the browser GUI explicitly"
        Write-Host "  invoke / run-local   Run Invoke-ServerDiscovery.ps1 on this box only"
        Write-Host "  scan / scan-local    (aliases - all do the same thing)"
        Write-Host "                       Any extra args are passed to the script."
        Write-Host "  cli / console / tui  Legacy console wizard (Start-DiscoverySession.ps1)"
  Write-Host "  newui                Launch the dossier-style UI (test/preview)"
        Write-Host ""
        Write-Host "  version / -v         Show installed version + path"
        Write-Host "  folder / where       Show the install folder paths"
        Write-Host "  open                 Open the install folder in Explorer"
        Write-Host "  stop / kill          Stop any running SDT GUI process"
        Write-Host "  update / upgrade     Force-pull the latest release from GitHub"
        Write-Host "  repair / reinstall   Re-download the install (use if 'sdt' won't launch)"
        Write-Host "  uninstall / remove   Remove SDT entirely (cleans up PATH)"
        Write-Host ""
        Write-Host "  help / -h / /?       Show this help"
        Write-Host ""
        Write-Host "EXAMPLES:" -ForegroundColor Yellow
        Write-Host "  sdt                  # browser GUI"
        Write-Host "  sdt invoke           # single-host scan, no UI"
        Write-Host "  sdt folder           # print install paths"
        Write-Host "  sdt update           # force-fetch latest from GitHub"
        Write-Host ""
        Write-Host "ENV VARS:" -ForegroundColor Yellow
        Write-Host "  SDT_NO_AUTOUPDATE=1  Skip the auto-update check on launch"
        Write-Host "  SDT_VERBOSE_THREADJOB=1  Show ThreadJob fallback details"
        Write-Host ""
        Write-Host "DOCS:  https://github.com/matt-magna5/SDT"
        Write-Host ""
        return
    }
    '^(version|-v|--version)$' {
        `$v = if (Test-Path `$VerFile) { Get-Content `$VerFile -Raw -EA 0 } else { 'unknown' }
        Write-Host "SDT installed at: `$AppDir"
        Write-Host "Version: `$(`$v.Trim())"
        return
    }
    '^(folder|folders|where|paths|installdir)$' {
        Write-Host ""
        Write-Host "SDT install paths:" -ForegroundColor Cyan
        Write-Host "  Root:    `$Root"
        Write-Host "  App:     `$AppDir"
        Write-Host "  Bin:     `$BinDir"
        Write-Host "  Python:  `$(Join-Path `$AppDir 'python')"
        Write-Host "  Version: `$VerFile"
        Write-Host ""
        Write-Host "Tip: 'sdt open' opens the install folder in Explorer."
        Write-Host ""
        return
    }
    '^(open|reveal|explorer)$' {
        if (Test-Path `$AppDir) {
            Write-Host "Opening `$AppDir ..." -ForegroundColor DarkGray
            Start-Process explorer.exe `$AppDir
        } else {
            Write-Host "App dir not found: `$AppDir" -ForegroundColor Red
        }
        return
    }
    '^(stop|kill|quit)$' {
        # Find PowerShell processes running the GUI script and stop them.
        `$guiPath = Join-Path `$AppDir 'Start-DiscoverySessionGUI.ps1'
        `$killed = 0
        Get-CimInstance Win32_Process -EA SilentlyContinue |
            Where-Object { `$_.Name -match 'powershell\.exe' -and `$_.CommandLine -and `$_.CommandLine.Contains(`$guiPath) } |
            ForEach-Object {
                Write-Host ("  [sdt] stopping GUI PID {0}" -f `$_.ProcessId) -ForegroundColor Yellow
                try { Stop-Process -Id `$_.ProcessId -Force -EA Stop; `$killed++ } catch { Write-Host "  [sdt] failed: `$_" -ForegroundColor Red }
            }
        if (`$killed -eq 0) { Write-Host "  [sdt] no SDT GUI process found." -ForegroundColor DarkGray }
        else { Write-Host ("  [sdt] stopped {0} process(es)." -f `$killed) -ForegroundColor Green }
        return
    }
    '^gui$' {
        # Explicit GUI launch - falls through to the default dispatch below.
        `$Mode = ''
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
    '^(repair|reinstall|heal)$' {
        # Same as update but bypasses every version check - always reinstalls.
        Invoke-SdtRepair | Out-Null
        return
    }
    '^(uninstall|remove)$' {
        # Full removal with an explicit, itemized checklist so the SE can prove
        # to a client that the jump box was returned to its original state.
        # 'sdt uninstall -y' / '-force' skips the confirmation prompt.
        `$force = @(`$Rest) -match '^-{0,2}(y|yes|f|force)`$'
        `$root  = Split-Path `$AppDir -Parent

        function _Step(`$done, `$label, `$detail) {
            `$mark = if (`$done) { '[x]' } else { '[ ]' }
            `$col  = if (`$done) { 'Green' } else { 'DarkGray' }
            Write-Host ("  {0} {1}" -f `$mark, `$label) -ForegroundColor `$col
            if (`$detail) { Write-Host ("      {0}" -f `$detail) -ForegroundColor DarkGray }
        }

        Write-Host ""
        Write-Host "  SDT UNINSTALL" -ForegroundColor Cyan
        Write-Host "  ---------------------------------------------------------" -ForegroundColor DarkGray

        # ---- Survey what exists BEFORE deleting, so we can report accurately ----
        `$pyDir      = Join-Path `$AppDir 'python'
        `$hasPython  = Test-Path `$pyDir
        `$tempPats   = @('sdt-install-*','sdt-getpip-*.log','sdt-pipinst-*.log','sdt-session-creds.xml')
        `$tempItems  = @()
        foreach (`$p in `$tempPats) {
            `$tempItems += @(Get-ChildItem -Path `$env:TEMP -Filter `$p -Force -EA 0)
        }
        `$credCache  = Join-Path `$env:TEMP 'sdt-session-creds.xml'
        `$hasCred    = Test-Path `$credCache

        # Discovery sessions / generated reports = CLIENT DATA. Enumerate before
        # touching anything so the operator sees exactly what would be destroyed.
        `$sessionDirs = @()
        foreach (`$base in @(`$root, (Join-Path `$AppDir '_output'), `$PWD.Path, [Environment]::GetFolderPath('Desktop'))) {
            if (`$base -and (Test-Path `$base)) {
                `$sessionDirs += @(Get-ChildItem -Path `$base -Directory -Filter 'Discovery-Session-*' -EA 0)
            }
        }
        `$sessionDirs = @(`$sessionDirs | Sort-Object FullName -Unique)
        `$reportFiles = @()
        foreach (`$d in `$sessionDirs) {
            `$reportFiles += @(Get-ChildItem -Path `$d.FullName -Filter '*DiscoveryReport*.html' -Recurse -EA 0)
        }

        if (`$sessionDirs.Count -gt 0) {
            Write-Host ""
            Write-Host "  CLIENT DATA FOUND - `$(`$sessionDirs.Count) discovery session folder(s), `$(`$reportFiles.Count) report(s):" -ForegroundColor Yellow
            foreach (`$d in (`$sessionDirs | Select-Object -First 12)) {
                Write-Host "      `$(`$d.FullName)" -ForegroundColor DarkYellow
            }
            if (`$sessionDirs.Count -gt 12) { Write-Host "      ... +`$(`$sessionDirs.Count - 12) more" -ForegroundColor DarkYellow }
            Write-Host "  These contain collected client inventory data and generated reports." -ForegroundColor Gray
            if (-not `$force) {
                Write-Host ""
                `$ans = Read-Host "  Delete these too? [y]es / [n]o - keep them / [c]ancel uninstall"
                if (`$ans -imatch '^[Cc]') { Write-Host "  Cancelled - nothing removed." -ForegroundColor Yellow; return }
                `$killData = (`$ans -imatch '^[Yy]')
            } else { `$killData = `$true }
        } else { `$killData = `$false }

        Write-Host ""
        Write-Host "  Removing:" -ForegroundColor White

        # 1. Session data / reports (only if authorized)
        `$dataRemoved = 0
        `$rescuedTo   = ''
        if (`$killData) {
            foreach (`$d in `$sessionDirs) {
                try { Remove-Item `$d.FullName -Recurse -Force -EA Stop; `$dataRemoved++ } catch { }
            }
            _Step (`$dataRemoved -eq `$sessionDirs.Count) "Discovery sessions + generated reports" "`$dataRemoved of `$(`$sessionDirs.Count) folder(s) deleted"
        } elseif (`$sessionDirs.Count -gt 0) {
            # The whole SDT root gets deleted below. Any session folder living
            # INSIDE it would be destroyed regardless of the user's choice, so
            # move those out to the Desktop first and say where they went.
            `$rootPrefix = `$root.TrimEnd('\') + '\'
            `$inRoot = @(`$sessionDirs | Where-Object { `$_.FullName.StartsWith(`$rootPrefix, 'OrdinalIgnoreCase') })
            if (`$inRoot.Count -gt 0) {
                `$rescuedTo = Join-Path ([Environment]::GetFolderPath('Desktop')) 'SDT-Discovery-Sessions'
                try {
                    if (-not (Test-Path `$rescuedTo)) { New-Item -ItemType Directory -Path `$rescuedTo -Force | Out-Null }
                    foreach (`$d in `$inRoot) {
                        try { Move-Item `$d.FullName -Destination `$rescuedTo -Force -EA Stop } catch { }
                    }
                    _Step `$true "Kept session data MOVED out of the install folder" `$rescuedTo
                } catch {
                    _Step `$false "Could not move kept session data out of the install folder" "`$rescuedTo"
                    `$rescuedTo = ''
                }
            }
            _Step `$false "Discovery sessions + reports - KEPT at user request" "`$(`$sessionDirs.Count) folder(s) preserved"
        }

        # 2. Bundled Python + plink that SDT downloaded
        if (`$hasPython) {
            try { Remove-Item `$pyDir -Recurse -Force -EA Stop } catch { }
            _Step (-not (Test-Path `$pyDir)) "Bundled Python 3.12 + plink.exe" "`$pyDir"
        } else {
            _Step `$true "Bundled Python - none was installed" ""
        }

        # 3. Temp files (installer zips, pip logs, and the CREDENTIAL CACHE)
        `$tmpKilled = 0
        foreach (`$t in `$tempItems) {
            try { Remove-Item `$t.FullName -Recurse -Force -EA Stop; `$tmpKilled++ } catch { }
        }
        _Step (`$tmpKilled -eq @(`$tempItems).Count) "Temporary files (installer zips, logs)" "`$tmpKilled item(s) from `$env:TEMP"
        if (`$hasCred) {
            _Step (-not (Test-Path `$credCache)) "Cached session CREDENTIALS purged" "`$credCache"
        } else {
            _Step `$true "Cached session credentials - none present" ""
        }

        # 4. The install tree itself - the entire SDT root, not just its contents
        try { Remove-Item `$root -Recurse -Force -EA Stop } catch { }
        `$rootGone = -not (Test-Path `$root)
        _Step `$rootGone "SDT root folder removed (app, bin, sdt command)" `$root
        if (-not `$rootGone) {
            Write-Host "      NOTE: some files were locked. Close any running SDT window ('sdt stop') and re-run." -ForegroundColor Yellow
        }
        # Drop the Magna5 parent too when SDT was the only thing in it, so the
        # uninstall leaves no empty folders behind.
        `$parent = Split-Path `$root -Parent
        if (`$rootGone -and (Test-Path `$parent)) {
            try {
                if (-not (Get-ChildItem `$parent -Force -EA SilentlyContinue)) {
                    Remove-Item `$parent -Force -Recurse -EA Stop
                    _Step (-not (Test-Path `$parent)) "Empty parent folder removed" `$parent
                }
            } catch { }
        }

        # 5. PATH entry
        `$pathCleaned = `$false
        try {
            `$userPath = [Environment]::GetEnvironmentVariable('Path','User')
            `$new = (`$userPath -split ';') | Where-Object { `$_ -and `$_ -notmatch 'Magna5\\SDT\\bin' }
            [Environment]::SetEnvironmentVariable('Path', (`$new -join ';'), 'User')
            `$verify = [Environment]::GetEnvironmentVariable('Path','User')
            `$pathCleaned = (`$verify -notmatch 'Magna5\\SDT\\bin')
        } catch { }
        _Step `$pathCleaned "PATH entry removed (user environment)" ""

        # 6. Registry - SDT never writes any; state this explicitly so the
        #    operator does not have to wonder.
        _Step `$true "Registry - no keys were ever created by SDT" "nothing to remove"

        # 7. Target servers - remind that WinRM/state is restored per-session
        Write-Host ""
        Write-Host "  Target servers:" -ForegroundColor White
        _Step `$true "WinRM restored to original state after each discovery run" "handled per-session; nothing installed on targets"

        Write-Host ""
        Write-Host "  ---------------------------------------------------------" -ForegroundColor DarkGray
        if (`$rootGone -and `$pathCleaned) {
            Write-Host "  UNINSTALL COMPLETE - system returned to original state." -ForegroundColor Green
        } else {
            Write-Host "  UNINSTALL INCOMPLETE - see unchecked items above." -ForegroundColor Yellow
        }
        if (-not `$killData -and `$sessionDirs.Count -gt 0) {
            Write-Host "  Client data was intentionally kept. Delete manually when done:" -ForegroundColor Yellow
            if (`$rescuedTo) {
                Write-Host "    `$rescuedTo   (moved here out of the install folder)" -ForegroundColor DarkYellow
            }
            foreach (`$d in (`$sessionDirs | Select-Object -First 5)) {
                if (-not `$d.FullName.StartsWith(`$root.TrimEnd('\') + '\', 'OrdinalIgnoreCase')) { Write-Host "    `$(`$d.FullName)" -ForegroundColor DarkYellow }
            }
        }
        Write-Host "  Close and reopen your terminal for the PATH change to apply." -ForegroundColor DarkGray
        Write-Host ""
        return
    }
}

# BEFORE auto-update: validate the existing install. If broken, force repair.
# This makes broken installs self-heal without the user needing to know about
# 'sdt repair' or pasting one-liners.
`$health = Test-SdtHealth
if (-not `$health.ok) {
    Write-Host "  [sdt] BROKEN INSTALL detected: `$(`$health.reason)" -ForegroundColor Red
    Write-Host "  [sdt] auto-repairing..." -ForegroundColor Yellow
    if (Invoke-SdtRepair) {
        # Re-check after repair
        `$health = Test-SdtHealth
        if (-not `$health.ok) {
            Write-Host "  [sdt] STILL broken after repair: `$(`$health.reason)" -ForegroundColor Red
            Write-Host "  [sdt] aborting launch. Report this." -ForegroundColor Red
            return
        }
        Write-Host "  [sdt] self-heal succeeded." -ForegroundColor Green
    } else {
        Write-Host "  [sdt] could not repair. Try: sdt uninstall; then reinstall." -ForegroundColor Red
        return
    }
}

# Start-Process -ArgumentList does NOT quote array elements, so any path
# containing a space (e.g. C:\Users\John Smith\AppData\...) is split into
# separate tokens and the child process dies with "the file does not have a
# '.ps1' extension". Every path handed to powershell.exe must be pre-quoted.
function _Q { param([string]`$p) if (`$p -match '\s') { '"' + `$p + '"' } else { `$p } }

`$updateResult = Invoke-SdtAutoUpdate
if (`$updateResult -eq 'updated') {
    # Relaunch with the new version - this process has old code in memory
    `$fwd = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (_Q `$PSCommandPath))
    if (`$Mode -and `$Mode -ne 'gui') { `$fwd += `$Mode }
    if (`$Rest) { `$fwd += @(`$Rest) }
    Start-Process powershell.exe -ArgumentList `$fwd
    exit 0
}

# Normalize `$Rest so empty splat is safe on PS 5.1 + PS 7
if (`$null -eq `$Rest) { `$Rest = @() }
`$RestArr = @(`$Rest)

switch -Regex (`$Mode) {
    '^(cli|console|tui)$'   {
        if (`$RestArr.Count -gt 0) { & (Join-Path `$AppDir 'Start-DiscoverySession.ps1') @RestArr }
        else { & (Join-Path `$AppDir 'Start-DiscoverySession.ps1') }
        return
    }
    '^(invoke|local|bare|run-local|runlocal|scan-local|scan)$' {
        if (`$RestArr.Count -gt 0) { & (Join-Path `$AppDir 'Invoke-ServerDiscovery.ps1') @RestArr }
        else { & (Join-Path `$AppDir 'Invoke-ServerDiscovery.ps1') }
        return
    }
    '^(newui|new-ui|ui2|dossier)$' {
        # Launch GUI with the dossier-style light UI (test/preview flag)
        `$guiScript = Join-Path `$AppDir 'Start-DiscoverySessionGUI.ps1'
        `$psArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (_Q `$guiScript), '-NewUI')
        try {
            `$child = Start-Process powershell.exe -ArgumentList `$psArgs -PassThru
            Write-Host ("  [sdt] New UI launched in background (PID {0}). Browser will open shortly." -f `$child.Id) -ForegroundColor DarkGray
        } catch {
            Write-Host "  [sdt] Detached launch failed - falling back to in-process." -ForegroundColor Yellow
            & `$guiScript -NewUI
        }
        return
    }
    default {
        # Detached GUI launch - spawn the listener in a NEW PowerShell process
        # so the user's current prompt is freed immediately. Without this the
        # listener's blocking GetContext() call wedges the calling shell and
        # Ctrl+C doesn't escape (only X works). The GUI script self-minimizes
        # its own console after startup, so the user just sees the browser.
        `$guiScript = Join-Path `$AppDir 'Start-DiscoverySessionGUI.ps1'
        `$psArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (_Q `$guiScript))
        if (`$RestArr.Count -gt 0) { `$psArgs += `$RestArr }
        try {
            `$child = Start-Process powershell.exe -ArgumentList `$psArgs -PassThru
            Write-Host ("  [sdt] GUI launched in background (PID {0}). Browser will open shortly." -f `$child.Id) -ForegroundColor DarkGray
            Write-Host "  [sdt] Your prompt is yours. To stop the GUI: close the browser tab + run 'sdt stop' or kill the PID above." -ForegroundColor DarkGray
        } catch {
            Write-Host "  [sdt] Detached launch failed (`$(`$_.Exception.Message)) - falling back to in-process." -ForegroundColor Yellow
            if (`$RestArr.Count -gt 0) { & `$guiScript @RestArr }
            else { & `$guiScript }
        }
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
Say "  sdt invoke      local per-host Invoke-ServerDiscovery (also: sdt run-local, sdt scan)" DarkGray
Say "  sdt cli         legacy console session wizard" DarkGray
Say "  sdt update      pull latest tag from GitHub" DarkGray
Say "  sdt repair      force re-download (use if sdt won't launch)" DarkGray
Say "  sdt uninstall   remove everything" DarkGray
Say "  sdt version     show install info" DarkGray
Say "  sdt folder      show install paths" DarkGray
Say "  sdt open        open install folder in Explorer" DarkGray
Say "  sdt stop        stop a running SDT GUI process" DarkGray
Say "  sdt help        show all commands" DarkGray
Say ""
Say "Auto-update runs on every 'sdt' launch - you'll never paste a" DarkGray
Say "download one-liner again." DarkGray
Say ""

# ----- Launch GUI immediately so user sees something ------------------------
if (-not $NoLaunch) {
    Say "Launching GUI now..." Yellow
    & $ShimPS
}
