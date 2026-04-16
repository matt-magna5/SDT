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
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # Primary: HttpClient chunked stream with live progress
    try {
        $client   = New-Object System.Net.Http.HttpClient
        $client.Timeout = [TimeSpan]::FromSeconds($TimeoutSec)
        $response = $client.GetAsync($Url, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).Result
        $response.EnsureSuccessStatusCode() | Out-Null
        $total    = $response.Content.Headers.ContentLength
        $inStream = $response.Content.ReadAsStreamAsync().Result
        $outFile  = New-Object System.IO.FileStream($Dest, [System.IO.FileMode]::Create,
                        [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        $buffer    = New-Object byte[] 65536
        $totalRead = [long]0; $lastRead = [long]0
        $lastTime  = [DateTime]::Now; $startTime = [DateTime]::Now
        $read      = 0
        while (($read = $inStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
            $outFile.Write($buffer, 0, $read)
            $totalRead += $read
            $now     = [DateTime]::Now
            $elapsed = ($now - $lastTime).TotalSeconds
            if ($elapsed -ge 0.4) {
                $speed    = [math]::Round(($totalRead - $lastRead) / 1KB / $elapsed, 0)
                $lastRead = $totalRead; $lastTime = $now
                $readMB   = [math]::Round($totalRead / 1MB, 1)
                $line = if ($total -gt 0) {
                    $pct = [math]::Round($totalRead / $total * 100, 0)
                    "`r  {0}  {1}%  {2} / {3} MB  {4} KB/s   " -f $Label, $pct, $readMB, [math]::Round($total/1MB,1), $speed
                } else {
                    "`r  {0}  {1} MB  {2} KB/s   " -f $Label, $readMB, $speed
                }
                Write-Host $line -NoNewline -ForegroundColor Cyan
            }
        }
        $outFile.Close(); $inStream.Close(); $client.Dispose()
        $avgSpd = [math]::Round($totalRead / 1KB / ([DateTime]::Now - $startTime).TotalSeconds, 0)
        Write-Host ("`r  {0}  Done  {1} MB  avg {2} KB/s                    " -f $Label, [math]::Round($totalRead/1MB,1), $avgSpd) -ForegroundColor Green
        return $true
    } catch {
        Write-Host "`r  HttpClient failed, trying fallback...                " -ForegroundColor DarkGray
        try {
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $Url -OutFile $Dest -UseBasicParsing `
                -TimeoutSec $TimeoutSec -ErrorAction Stop
            Write-Host "  [OK] $Label downloaded (fallback method)." -ForegroundColor Green
            return $true
        } catch {
            Write-Host "  [FAIL] $Label download failed: $_" -ForegroundColor Red
            return $false
        }
    }
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
Write-Host "  Setup complete. Run Start-DiscoverySession_2.0.ps1 to begin." -ForegroundColor Cyan
Write-Host ""
