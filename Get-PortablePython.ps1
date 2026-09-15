#Requires -Version 3.0
<#
.SYNOPSIS
    Downloads SDT dependencies: portable Python (for HTML reports) and
    plink.exe (for Linux SSH discovery). Run once per machine.
.EXAMPLE
    .\Get-PortablePython.ps1
#>

param(
    # Folder holding prepackaged dependencies: python-*-embed-amd64.zip,
    # plink.exe and SHA256SUMS.txt. A local path, USB drive or UNC share.
    # Defaults to $env:SDT_DEPS_PATH so it can be set before `iwr | iex`.
    [string]$DepsPath = $env:SDT_DEPS_PATH,
    # Never touch the network; prepackaged dependencies only.
    [switch]$Offline
)

# .NET Framework < 4.5 has no Tls12 member; guard it so the script still loads
# and can fall back to prepackaged dependencies.
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

$script:NoNetwork = [bool]$Offline -or ($env:SDT_OFFLINE -match '^(1|true|yes)$')

# SHA256 manifest shipped next to prepackaged deps. Every local candidate is
# checked against it, so a corrupt copy or a tampered share is never trusted.
function Get-DepsManifest([string]$dir) {
    $map = @{}
    $f = Join-Path $dir 'SHA256SUMS.txt'
    if (-not (Test-Path $f)) { return $map }
    foreach ($line in (Get-Content $f -ErrorAction SilentlyContinue)) {
        if ($line -match '^\s*([0-9a-fA-F]{64})\s+\*?(.+?)\s*$') {
            $map[($Matches[2] -replace '\\','/').ToLower()] = $Matches[1].ToLower()
        }
    }
    return $map
}

function Get-FileSha256([string]$path) {
    # Get-FileHash is PS 4+; use .NET directly on PS 3 (Server 2008 R2 / 2012).
    if (Get-Command Get-FileHash -ErrorAction SilentlyContinue) {
        return (Get-FileHash -Path $path -Algorithm SHA256).Hash.ToLower()
    }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $fs  = [System.IO.File]::OpenRead($path)
    try { return (($sha.ComputeHash($fs) | ForEach-Object { $_.ToString('x2') }) -join '') }
    finally { $fs.Close(); $sha.Dispose() }
}

# Locate a SHA256-verified prepackaged copy of $name. Search order:
#   1. -DepsPath / SDT_DEPS_PATH   (explicit local folder, USB, UNC share)
#   2. .\deps beside this script    (bundled inside every SDT release)
# Returns the full path, or $null when no verified copy exists.
function Find-PrepackagedDep([string]$name) {
    $dirs = @()
    if ($DepsPath) { $dirs += $DepsPath }
    $dirs += (Join-Path $PSScriptRoot 'deps')
    foreach ($d in $dirs) {
        if (-not $d -or -not (Test-Path $d)) { continue }
        $cand = Join-Path $d $name
        if (-not (Test-Path $cand)) { continue }
        $want = (Get-DepsManifest $d)[$name.ToLower()]
        if (-not $want) {
            Write-Host "  [WARN] $name in $d is not listed in SHA256SUMS.txt - not trusted" -ForegroundColor Yellow
            continue
        }
        $got = ''
        try { $got = Get-FileSha256 $cand } catch { }
        if ($got -eq $want) { return $cand }
        Write-Host "  [WARN] $name in $d failed SHA256 verification - ignoring it" -ForegroundColor Yellow
    }
    return $null
}


function Expand-ZipCompat([string]$ZipPath, [string]$Destination) {
    if (-not (Test-Path $Destination)) { New-Item -ItemType Directory -Force -Path $Destination | Out-Null }
    if (Get-Command Expand-Archive -ErrorAction SilentlyContinue) {
        Expand-Archive -Path $ZipPath -DestinationPath $Destination -Force
    } else {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $Destination)
    }
}

Write-Host ""
Write-Host "  SDT -- Dependencies Setup" -ForegroundColor Cyan
Write-Host ("  " + "=" * 40) -ForegroundColor DarkCyan
Write-Host ""

function Get-FileWithProgress {
    param([string]$Url, [string]$Dest, [string]$Label, [int]$TimeoutSec = 120,
          [int]$MinBytes = 0)

    # A partially-written file is worse than no file: a truncated python.exe or
    # plink.exe fails later with a cryptic "not a valid Win32 application".
    # Every method routes its result through this so a short file is rejected
    # and the next method gets a turn.
    function _Validate([string]$path, [int]$min, [string]$lbl) {
        if (-not (Test-Path $path)) { return $false }
        $len = (Get-Item $path -EA SilentlyContinue).Length
        if ($min -gt 0 -and $len -lt $min) {
            Write-Host ("`r  {0}  TRUNCATED ({1} KB < {2} KB expected) - discarding  " -f `
                        $lbl, [math]::Round($len/1KB,0), [math]::Round($min/1KB,0)) -ForegroundColor Yellow
            Remove-Item $path -Force -EA SilentlyContinue
            return $false
        }
        return ($len -gt 0)
    }

    try {
        $allTls = [enum]::GetValues([Net.SecurityProtocolType]) | Where-Object { $_ -match 'Tls' }
        [Net.ServicePointManager]::SecurityProtocol = $allTls -as [Net.SecurityProtocolType]
    } catch {
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
    }
    try { [Net.WebRequest]::DefaultWebProxy = [Net.WebRequest]::GetSystemWebProxy()
          [Net.WebRequest]::DefaultWebProxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials } catch { }

    # Method 1: HttpClient chunked with progress
    Write-Host "  [1/5] HttpClient - up to ${TimeoutSec}s" -ForegroundColor DarkGray
    try {
        $client   = New-Object System.Net.Http.HttpClient
        $client.Timeout = [TimeSpan]::FromSeconds($TimeoutSec)
        $response = $client.GetAsync($Url, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).Result
        $response.EnsureSuccessStatusCode() | Out-Null
        $total    = $response.Content.Headers.ContentLength
        $inStream = $response.Content.ReadAsStreamAsync().Result
        # A stalled body read would otherwise block forever; the client Timeout
        # only covers the headers when ResponseHeadersRead is used.
        try { if ($inStream.CanTimeout) { $inStream.ReadTimeout = 30000 } } catch { }
        $outFile  = New-Object System.IO.FileStream($Dest, [System.IO.FileMode]::Create,
                        [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        $buf = New-Object byte[] 65536; $totalRead = [long]0
        $lastRead = [long]0; $lastTime = [DateTime]::Now; $startTime = [DateTime]::Now; $read = 0
        while (($read = $inStream.Read($buf, 0, $buf.Length)) -gt 0) {
            $outFile.Write($buf, 0, $read); $totalRead += $read
            $now = [DateTime]::Now
            if (($now - $lastTime).TotalSeconds -ge 0.4) {
                $speed = [math]::Round(($totalRead-$lastRead)/1KB/($now-$lastTime).TotalSeconds,0)
                $lastRead = $totalRead; $lastTime = $now
                $readMB = [math]::Round($totalRead/1MB,1)
                $line = if ($total -gt 0) {
                    "`r  {0}  {1}%  {2} / {3} MB  {4} KB/s   " -f $Label,
                        [math]::Round($totalRead/$total*100,0),$readMB,[math]::Round($total/1MB,1),$speed
                } else { "`r  {0}  {1} MB  {2} KB/s   " -f $Label,$readMB,$speed }
                Write-Host $line -NoNewline -ForegroundColor Cyan
            }
        }
        $outFile.Close(); $inStream.Close(); $client.Dispose()
        $avg = [math]::Round($totalRead/1KB/([DateTime]::Now-$startTime).TotalSeconds,0)
        Write-Host ("`r  {0}  Done  {1} MB  avg {2} KB/s                    " -f $Label,[math]::Round($totalRead/1MB,1),$avg) -ForegroundColor Green
        if (_Validate $Dest $MinBytes $Label) { return $true }
    } catch { Write-Host "`r  Method 1 failed - trying WebClient...                 " -ForegroundColor DarkGray }

    # Tear a download job down WITHOUT ever blocking the installer.
    # Stop-Job / Remove-Job can hang indefinitely on a Start-BitsTransfer job,
    # because the BITS service - not this process - owns the transfer. A job that
    # already finished is removed normally; one still running is stopped on a
    # side runspace with a hard 3s wait, and abandoned if that hangs (the orphan
    # dies when this short-lived installer process exits).
    function _Reap($job) {
        if (-not $job) { return }
        try {
            if ($job.State -ne 'Running' -and $job.State -ne 'NotStarted') {
                Remove-Job $job -Force -EA SilentlyContinue
                return
            }
            $rs = [runspacefactory]::CreateRunspace(); $rs.Open()
            $ps = [powershell]::Create(); $ps.Runspace = $rs
            [void]$ps.AddScript({ param($j) try { $j.StopJob() } catch { } }).AddArgument($job)
            $h = $ps.BeginInvoke()
            if ($h.AsyncWaitHandle.WaitOne(3000)) {
                try { $ps.Dispose(); $rs.Dispose() } catch { }
                Remove-Job $job -Force -EA SilentlyContinue
            }
        } catch { }
    }

    # True when the file can be opened for reading, i.e. the downloader has
    # released it. BITS can hold the handle briefly after the bytes land, and
    # extracting a still-locked zip fails with a sharing violation.
    function _Unlocked([string]$path) {
        try {
            $fs = [System.IO.File]::Open($path, 'Open', 'Read', 'ReadWrite')
            $fs.Close()
            return $true
        } catch { return $false }
    }

    # Watch a background download job and decide how it ended, with a live
    # countdown so the operator can see exactly when the next method starts.
    #
    #   success  - the job completed, OR the file reached the expected size and
    #              has been stable and unlocked for $settleSec. The second case
    #              matters: a finished download does not always flip the job to
    #              Completed promptly, and treating a complete file as a "stall"
    #              threw away good downloads.
    #   failure  - no bytes within $noRespSec, OR growth stopped while still
    #              BELOW the expected size for $stallSec, OR $maxSec exceeded.
    #
    # $endWriter: BITS and certutil write the destination only when finished, so
    # zero bytes mid-transfer is normal for them - they are judged on job state
    # and the $maxSec ceiling, never on "no response".
    #
    # NOTE: loop on $job.State. PowerShell job objects have no HasExited
    # property (always $null), so the previous `while (-not $job.HasExited)`
    # could never observe completion. That was the root of the hang.
    function _spin($job, $destPath, $lbl, $minBytes = 0, [switch]$endWriter,
                   $stallSec = 30, $noRespSec = 20, $maxSec = 180, $settleSec = 3,
                   [string]$nextLabel = 'next method') {
        $sp = @('|','/','-','\'); $i = 0
        $lastSz = -1; $lastGrowth = [DateTime]::Now; $everHadBytes = $false
        $started = [DateTime]::Now
        if ($endWriter) { $noRespSec = $maxSec }
        while ($job.State -eq 'Running' -or $job.State -eq 'NotStarted') {
            $sz = [long]0
            if (Test-Path $destPath) { $sz = [long](Get-Item $destPath -EA SilentlyContinue).Length }
            if ($sz -gt 0) { $everHadBytes = $true }
            if ($sz -gt $lastSz) { $lastSz = $sz; $lastGrowth = [DateTime]::Now }
            $idle    = ([DateTime]::Now - $lastGrowth).TotalSeconds
            $elapsed = ([DateTime]::Now - $started).TotalSeconds
            $mb      = [math]::Round($sz/1MB,1)

            # Complete, stable, released -> done even if the job has not reported
            # Completed yet.
            if ($minBytes -gt 0 -and $sz -ge $minBytes -and $idle -ge $settleSec -and (_Unlocked $destPath)) {
                Write-Host ("`r  {0}  Done  {1} MB  ({2}s)                              " -f $lbl, $mb, [int]$elapsed) -ForegroundColor Green
                return $true
            }
            if (-not $everHadBytes -and $elapsed -gt $noRespSec) {
                Write-Host ("`r  {0}  no response after {1}s - {2}...                 " -f $lbl, [int]$elapsed, $nextLabel) -ForegroundColor DarkGray
                return $false
            }
            if ($everHadBytes -and ($minBytes -le 0 -or $sz -lt $minBytes) -and $idle -gt $stallSec) {
                Write-Host ("`r  {0}  stalled at {1} MB for {2}s - {3}...             " -f $lbl, $mb, [int]$idle, $nextLabel) -ForegroundColor DarkGray
                return $false
            }
            if ($elapsed -gt $maxSec) {
                Write-Host ("`r  {0}  exceeded {1}s limit at {2} MB - {3}...          " -f $lbl, $maxSec, $mb, $nextLabel) -ForegroundColor DarkGray
                return $false
            }

            # Countdown to whichever limit will fire first from here.
            $deadlines = @($maxSec - $elapsed)
            if (-not $everHadBytes) { $deadlines += ($noRespSec - $elapsed) }
            elseif ($minBytes -le 0 -or $sz -lt $minBytes) { $deadlines += ($stallSec - $idle) }
            $left = [int][math]::Ceiling(($deadlines | Measure-Object -Minimum).Minimum)
            if ($left -lt 0) { $left = 0 }
            $pct = ''
            if ($minBytes -gt 0 -and $sz -gt 0) { $pct = ' ~{0}%' -f [int][math]::Min(99, ($sz / $minBytes) * 100) }
            Write-Host ("`r  {0}  {1}  {2} MB{3}   elapsed {4}s   {5} in {6}s     " -f `
                        $lbl, $sp[$i%4], $mb, $pct, [int]$elapsed, $nextLabel, $left) -NoNewline -ForegroundColor Cyan
            $i++; Start-Sleep -Milliseconds 250
        }
        # Job left the Running state on its own.
        $elapsed = [int]([DateTime]::Now - $started).TotalSeconds
        $mb = 0
        if (Test-Path $destPath) { $mb = [math]::Round((Get-Item $destPath).Length/1MB,1) }
        if ($job.State -ne 'Completed') {
            Write-Host ("`r  {0}  method ended ({1}) at {2} MB after {3}s              " -f $lbl, $job.State, $mb, $elapsed) -ForegroundColor DarkGray
        } else {
            Write-Host ("`r  {0}  Done  {1} MB  ({2}s)                              " -f $lbl, $mb, $elapsed) -ForegroundColor Green
        }
        return $true   # caller still validates the size before accepting
    }

    # Method 2: WebClient
    Write-Host "  [2/5] WebClient - up to 180s" -ForegroundColor DarkGray
    try {
        if (Test-Path $Dest) { Remove-Item $Dest -Force -EA SilentlyContinue }
        $job = Start-Job -ScriptBlock {
            param($u,$d)
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $wc = New-Object System.Net.WebClient
            $wc.Proxy = [Net.WebRequest]::GetSystemWebProxy()
            $wc.Proxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials
            $wc.DownloadFile($u, $d)
        } -ArgumentList $Url, $Dest
        $ok = _spin $job $Dest $Label $MinBytes -nextLabel 'next: Invoke-WebRequest'
        _Reap $job
        if ($ok -and (_Validate $Dest $MinBytes $Label)) { return $true }
    } catch { }
    Write-Host "`r  Method 2 failed - trying Invoke-WebRequest...         " -ForegroundColor DarkGray

    # Method 3: Invoke-WebRequest
    Write-Host "  [3/5] Invoke-WebRequest - up to 180s" -ForegroundColor DarkGray
    try {
        if (Test-Path $Dest) { Remove-Item $Dest -Force -EA SilentlyContinue }
        $job = Start-Job -ScriptBlock {
            param($u,$d,$t)
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $u -OutFile $d -UseBasicParsing -TimeoutSec $t -EA Stop
        } -ArgumentList $Url, $Dest, $TimeoutSec
        $afterIwr = 'next: certutil'
        if (Get-Command Start-BitsTransfer -EA SilentlyContinue) { $afterIwr = 'next: BITS' }
        $ok = _spin $job $Dest $Label $MinBytes -nextLabel $afterIwr
        _Reap $job
        if ($ok -and (_Validate $Dest $MinBytes $Label)) { return $true }
    } catch { }
    Write-Host "`r  Method 3 failed - trying BITS...                      " -ForegroundColor DarkGray

    # Method 4: BITS Transfer
    if (Get-Command Start-BitsTransfer -EA SilentlyContinue) {
        Write-Host "  [4/5] BITS - up to 180s" -ForegroundColor DarkGray
        try {
            if (Test-Path $Dest) { Remove-Item $Dest -Force -EA SilentlyContinue }
            $job = Start-Job -ScriptBlock { param($u,$d) Start-BitsTransfer -Source $u -Destination $d -EA Stop } -ArgumentList $Url, $Dest
            # BITS writes the destination only when finished; judge it on job
            # state and the time ceiling, never on zero bytes mid-transfer.
            $ok = _spin $job $Dest $Label $MinBytes -endWriter -nextLabel 'next: certutil'
            _Reap $job
            if ($ok -and (_Validate $Dest $MinBytes $Label)) { return $true }
        } catch { }
        Write-Host "`r  Method 4 failed - trying certutil...                  " -ForegroundColor DarkGray
    }

    # Method 5: certutil (every Windows version including 2008 R2)
    Write-Host "  [5/5] certutil - up to 180s (last method)" -ForegroundColor DarkGray
    try {
        if (Test-Path $Dest) { Remove-Item $Dest -Force -EA SilentlyContinue }
        $job = Start-Job -ScriptBlock { param($u,$d) & certutil.exe -urlcache -split -f $u $d 2>&1 } -ArgumentList $Url, $Dest
        # certutil also writes the file only at the very end.
        [void](_spin $job $Dest $Label $MinBytes -endWriter -nextLabel 'giving up')
        _Reap $job
        & certutil.exe -urlcache -f $Url delete 2>&1 | Out-Null
        if (_Validate $Dest $MinBytes $Label) { return $true }
    } catch { }

    Write-Host "  [FAIL] All download methods failed for: $Label" -ForegroundColor Red
    return $false
}

# --- PORTABLE PYTHON ---

$PY_VERSION  = "3.12.6"
$PY_URL      = "https://www.python.org/ftp/python/$PY_VERSION/python-$PY_VERSION-embed-amd64.zip"
$DEST_FOLDER = Join-Path $PSScriptRoot "python"
$ZIP_TMP     = Join-Path $env:TEMP "sdt-py-embed.zip"

$PY_ZIP_NAME = "python-$PY_VERSION-embed-amd64.zip"
if (Test-Path (Join-Path $DEST_FOLDER "python.exe")) {
    Write-Host "  [OK] Portable Python already installed." -ForegroundColor Green
} else {
    # Prepackaged copy first: every SDT release ships this exact zip with a
    # pinned SHA256, so a network that blocks python.org costs nothing.
    $pyZip = Find-PrepackagedDep $PY_ZIP_NAME
    $fromBundle = [bool]$pyZip
    if ($fromBundle) {
        Write-Host "  [OK] Using prepackaged Python $PY_VERSION - SHA256 verified, no download" -ForegroundColor Green
    } elseif ($script:NoNetwork) {
        Write-Host "  [FAIL] Offline mode and no verified $PY_ZIP_NAME found." -ForegroundColor Red
        Write-Host "         Point -DepsPath (or SDT_DEPS_PATH) at a folder holding it plus SHA256SUMS.txt." -ForegroundColor DarkGray
    } else {
        Write-Host "  No prepackaged Python found - downloading from python.org" -ForegroundColor DarkGray
        if (Get-FileWithProgress -Url $PY_URL -Dest $ZIP_TMP -Label "Python $PY_VERSION" -MinBytes 8000000) {
            $pyZip = $ZIP_TMP
        }
    }
    if ($pyZip) {
        Expand-ZipCompat $pyZip $DEST_FOLDER
        if (-not $fromBundle) { Remove-Item $ZIP_TMP -Force -ErrorAction SilentlyContinue }
        if (Test-Path (Join-Path $DEST_FOLDER "python.exe")) {
            Write-Host "  [OK] Portable Python ready." -ForegroundColor Green
        } else {
            Write-Host "  [WARN] Extracted but python.exe not found. Check: $DEST_FOLDER" -ForegroundColor Yellow
        }
    }
}

# --- PLINK.EXE (PuTTY SSH client for Linux discovery) ---

$PLINK_URL  = "https://the.earth.li/~sgtatham/putty/latest/w64/plink.exe"
$PLINK_DEST = Join-Path $PSScriptRoot "plink.exe"

if (Test-Path $PLINK_DEST) {
    Write-Host "  [OK] plink.exe already present." -ForegroundColor Green
} else {
    $ok = $false
    $plinkSrc = Find-PrepackagedDep 'plink.exe'
    if ($plinkSrc) {
        try {
            Copy-Item -Path $plinkSrc -Destination $PLINK_DEST -Force -ErrorAction Stop
            $ok = $true
            Write-Host "  [OK] Using prepackaged plink.exe - SHA256 verified, no download" -ForegroundColor Green
        } catch {
            Write-Host "  [WARN] Could not copy prepackaged plink.exe: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    if (-not $ok -and -not $script:NoNetwork) {
        Write-Host "  No prepackaged plink.exe found - downloading" -ForegroundColor DarkGray
        $ok = Get-FileWithProgress -Url $PLINK_URL -Dest $PLINK_DEST -Label "plink.exe" -MinBytes 300000
    }
    if (-not $ok) {
        Write-Host "         Linux SSH discovery will not be available without it." -ForegroundColor DarkGray
    }
}

Write-Host ""
Write-Host "  Setup complete. Run Start-DiscoverySession.ps1 to begin." -ForegroundColor Cyan
Write-Host ""
