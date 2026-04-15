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

# --- PORTABLE PYTHON ---

$PY_VERSION  = "3.12.6"
$PY_URL      = "https://www.python.org/ftp/python/$PY_VERSION/python-$PY_VERSION-embed-amd64.zip"
$DEST_FOLDER = Join-Path $PSScriptRoot "python"
$ZIP_TMP     = Join-Path $env:TEMP "sdt-py-embed.zip"

if (Test-Path (Join-Path $DEST_FOLDER "python.exe")) {
    Write-Host "  [OK] Portable Python already installed." -ForegroundColor Green
} else {
    Write-Host "  Downloading Python $PY_VERSION embeddable package..." -ForegroundColor DarkGray
    try {
        $spin = @('|','/','-','\'); $si = 0
        $dlJob = Start-Job -ScriptBlock {
            param($u,$o)
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri $u -OutFile $o -UseBasicParsing
        } -ArgumentList $PY_URL, $ZIP_TMP
        while (-not $dlJob.HasExited) {
            Write-Host ("`r  Downloading Python $PY_VERSION...  $($spin[$si % 4])") -NoNewline -ForegroundColor DarkGray
            $si++; Start-Sleep -Milliseconds 150
        }
        Write-Host "`r  Download complete.                              " -ForegroundColor DarkGray
        Receive-Job $dlJob -ErrorAction Stop | Out-Null
        Remove-Job $dlJob -Force -ErrorAction SilentlyContinue
        New-Item -ItemType Directory -Path $DEST_FOLDER -Force | Out-Null
        Expand-Archive -Path $ZIP_TMP -DestinationPath $DEST_FOLDER -Force
        Remove-Item $ZIP_TMP -Force -ErrorAction SilentlyContinue
        if (Test-Path (Join-Path $DEST_FOLDER "python.exe")) {
            Write-Host "  [OK] Portable Python ready." -ForegroundColor Green
        } else {
            Write-Host "  [WARN] Python extracted but python.exe not found. Check: $DEST_FOLDER" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "  [FAIL] Python download failed: $_" -ForegroundColor Red
    }
}

# --- PLINK.EXE (PuTTY SSH client for Linux discovery) ---

$PLINK_URL  = "https://the.earth.li/~sgtatham/putty/latest/w64/plink.exe"
$PLINK_DEST = Join-Path $PSScriptRoot "plink.exe"

if (Test-Path $PLINK_DEST) {
    Write-Host "  [OK] plink.exe already present." -ForegroundColor Green
} else {
    Write-Host "  Downloading plink.exe (PuTTY SSH client)..." -ForegroundColor DarkGray
    try {
        Invoke-WebRequest -Uri $PLINK_URL -OutFile $PLINK_DEST -UseBasicParsing
        if (Test-Path $PLINK_DEST) {
            Write-Host "  [OK] plink.exe ready." -ForegroundColor Green
        } else {
            Write-Host "  [WARN] plink.exe download succeeded but file not found." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "  [FAIL] plink.exe download failed: $_" -ForegroundColor Red
        Write-Host "         Linux SSH discovery will not be available without it." -ForegroundColor DarkGray
    }
}

Write-Host ""
Write-Host "  Setup complete. Run Start-DiscoverySession_2.0.ps1 to begin." -ForegroundColor Cyan
Write-Host ""
