#Requires -Version 5.1
<#
.SYNOPSIS
    Downloads SDT dependencies: portable Python (for HTML reports) and
    plink.exe (for Linux SSH discovery). Run once per machine.
.EXAMPLE
    .\Get-PortablePython.ps1
#>

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Write-Host ""
Write-Host "  SDT -- Dependencies Setup" -ForegroundColor Cyan
Write-Host ("  " + "=" * 40) -ForegroundColor DarkCyan
Write-Host ""

function Get-FileWithProgress {
    param([string]$Url, [string]$Dest, [string]$Label, [int]$TimeoutSec = 120)

    try {
        $allTls = [enum]::GetValues([Net.SecurityProtocolType]) | Where-Object { $_ -match 'Tls' }
        [Net.ServicePointManager]::SecurityProtocol = $allTls -as [Net.SecurityProtocolType]
    } catch {
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
    }
    try { [Net.WebRequest]::DefaultWebProxy = [Net.WebRequest]::GetSystemWebProxy()
          [Net.WebRequest]::DefaultWebProxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials } catch { }

    # Method 1: HttpClient chunked with progress
    try {
        $client   = New-Object System.Net.Http.HttpClient
        $client.Timeout = [TimeSpan]::FromSeconds($TimeoutSec)
        $response = $client.GetAsync($Url, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).Result
        $response.EnsureSuccessStatusCode() | Out-Null
        $total    = $response.Content.Headers.ContentLength
        $inStream = $response.Content.ReadAsStreamAsync().Result
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
        return $true
    } catch { Write-Host "`r  Method 1 failed — trying WebClient...                 " -ForegroundColor DarkGray }

    function _spin($job, $destPath, $lbl) {
        $sp = @('|','/','-','\'); $i = 0
        $lastSz = -1; $lastGrowth = [DateTime]::Now; $everHadBytes = $false
        while (-not $job.HasExited) {
            $sz = if (Test-Path $destPath) { (Get-Item $destPath -EA SilentlyContinue).Length } else { 0 }
            if ($sz -gt 0) { $everHadBytes = $true }
            if ($sz -gt $lastSz) { $lastSz = $sz; $lastGrowth = [DateTime]::Now }
            $waited = ([DateTime]::Now - $lastGrowth).TotalSeconds
            if ($everHadBytes -and $sz -gt 102400 -and $waited -gt 5) {
                Stop-Job $job -EA SilentlyContinue
                Write-Host ("`r  {0}  Done  {1} MB                              " -f $lbl, [math]::Round($sz/1MB,1)) -ForegroundColor Green
                return $true
            }
            if (-not $everHadBytes -and $waited -gt 10) {
                Stop-Job $job -EA SilentlyContinue
                Write-Host ("`r  {0}  no response — skipping...                 " -f $lbl) -ForegroundColor DarkGray
                return $false
            }
            if ($everHadBytes -and $waited -gt 60) {
                Stop-Job $job -EA SilentlyContinue
                Write-Host ("`r  {0}  stalled mid-transfer — skipping...        " -f $lbl) -ForegroundColor DarkGray
                return $false
            }
            Write-Host ("`r  {0}  {1}  {2} MB   " -f $lbl, $sp[$i%4], [math]::Round($sz/1MB,1)) -NoNewline -ForegroundColor Cyan
            $i++; Start-Sleep -Milliseconds 150
        }
        $sz = if (Test-Path $destPath) { [math]::Round((Get-Item $destPath).Length/1MB,1) } else { 0 }
        Write-Host ("`r  {0}  Done  {1} MB                              " -f $lbl, $sz) -ForegroundColor Green
        return $true
    }

    # Method 2: WebClient
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
        $ok = _spin $job $Dest $Label
        Receive-Job $job -EA SilentlyContinue | Out-Null
        Remove-Job $job -Force -EA SilentlyContinue
        if ($ok -and (Test-Path $Dest)) { return $true }
    } catch { }
    Write-Host "`r  Method 2 failed — trying Invoke-WebRequest...         " -ForegroundColor DarkGray

    # Method 3: Invoke-WebRequest
    try {
        if (Test-Path $Dest) { Remove-Item $Dest -Force -EA SilentlyContinue }
        $job = Start-Job -ScriptBlock {
            param($u,$d,$t)
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $u -OutFile $d -UseBasicParsing -TimeoutSec $t -EA Stop
        } -ArgumentList $Url, $Dest, $TimeoutSec
        $ok = _spin $job $Dest $Label
        Receive-Job $job -EA SilentlyContinue | Out-Null
        Remove-Job $job -Force -EA SilentlyContinue
        if ($ok -and (Test-Path $Dest)) { return $true }
    } catch { }
    Write-Host "`r  Method 3 failed — trying BITS...                      " -ForegroundColor DarkGray

    # Method 4: BITS Transfer
    if (Get-Command Start-BitsTransfer -EA SilentlyContinue) {
        try {
            if (Test-Path $Dest) { Remove-Item $Dest -Force -EA SilentlyContinue }
            $job = Start-Job -ScriptBlock { param($u,$d) Start-BitsTransfer -Source $u -Destination $d -EA Stop } -ArgumentList $Url, $Dest
            $ok = _spin $job $Dest $Label
            Receive-Job $job -EA SilentlyContinue | Out-Null
            Remove-Job $job -Force -EA SilentlyContinue
            if ($ok -and (Test-Path $Dest)) { return $true }
        } catch { }
        Write-Host "`r  Method 4 failed — trying certutil...                  " -ForegroundColor DarkGray
    }

    # Method 5: certutil (every Windows version including 2008 R2)
    try {
        if (Test-Path $Dest) { Remove-Item $Dest -Force -EA SilentlyContinue }
        $job = Start-Job -ScriptBlock { param($u,$d) & certutil.exe -urlcache -split -f $u $d 2>&1 } -ArgumentList $Url, $Dest
        _spin $job $Dest $Label -stallSec 20 | Out-Null
        Remove-Job $job -Force -EA SilentlyContinue
        if (Test-Path $Dest) {
            & certutil.exe -urlcache -f $Url delete 2>&1 | Out-Null
            return $true
        }
    } catch { }

    Write-Host "  [FAIL] All download methods failed for: $Label" -ForegroundColor Red
    return $false
}

# --- PORTABLE PYTHON ---

$PY_VERSION  = "3.12.6"
$PY_URL      = "https://www.python.org/ftp/python/$PY_VERSION/python-$PY_VERSION-embed-amd64.zip"
$DEST_FOLDER = Join-Path $PSScriptRoot "python"
$ZIP_TMP     = Join-Path $env:TEMP "sdt-py-embed.zip"

if (Test-Path (Join-Path $DEST_FOLDER "python.exe")) {
    Write-Host "  [OK] Portable Python already installed." -ForegroundColor Green
} else {
    $ok = Get-FileWithProgress -Url $PY_URL -Dest $ZIP_TMP -Label "Python $PY_VERSION"
    if ($ok) {
        New-Item -ItemType Directory -Path $DEST_FOLDER -Force | Out-Null
        Expand-Archive -Path $ZIP_TMP -DestinationPath $DEST_FOLDER -Force
        Remove-Item $ZIP_TMP -Force -ErrorAction SilentlyContinue
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
    $ok = Get-FileWithProgress -Url $PLINK_URL -Dest $PLINK_DEST -Label "plink.exe"
    if (-not $ok) {
        Write-Host "         Linux SSH discovery will not be available without it." -ForegroundColor DarkGray
    }
}

Write-Host ""
Write-Host "  Setup complete. Run Start-DiscoverySession.ps1 to begin." -ForegroundColor Cyan
Write-Host ""
