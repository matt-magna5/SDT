#Requires -Version 5.1
<#
.SYNOPSIS
    Downloads the Python embeddable package into the SDT folder.
    Run once on any machine where Python is not installed system-wide.
    After this runs, Start-DiscoverySession will auto-generate HTML reports
    without any additional setup.
.EXAMPLE
    .\Get-PortablePython.ps1
#>

$PY_VERSION  = "3.12.6"
$PY_URL      = "https://www.python.org/ftp/python/$PY_VERSION/python-$PY_VERSION-embed-amd64.zip"
$DEST_FOLDER = Join-Path $PSScriptRoot "python"
$ZIP_TMP     = Join-Path $env:TEMP "sdt-py-embed.zip"

Write-Host ""
Write-Host "  SDT -- Portable Python Setup" -ForegroundColor Cyan
Write-Host ("  " + "=" * 40) -ForegroundColor DarkCyan
Write-Host ""

if (Test-Path (Join-Path $DEST_FOLDER "python.exe")) {
    Write-Host "  Portable Python already installed at:" -ForegroundColor Green
    Write-Host "  $DEST_FOLDER" -ForegroundColor White
    Write-Host ""
    exit 0
}

Write-Host "  Downloading Python $PY_VERSION embeddable package..." -ForegroundColor DarkGray
Write-Host "  $PY_URL" -ForegroundColor DarkGray
Write-Host ""

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $PY_URL -OutFile $ZIP_TMP -UseBasicParsing
} catch {
    Write-Host "  Download failed: $_" -ForegroundColor Red
    Write-Host ""
    exit 1
}

Write-Host "  Extracting to: $DEST_FOLDER" -ForegroundColor DarkGray
New-Item -ItemType Directory -Path $DEST_FOLDER -Force | Out-Null
Expand-Archive -Path $ZIP_TMP -DestinationPath $DEST_FOLDER -Force
Remove-Item $ZIP_TMP -Force -ErrorAction SilentlyContinue

if (Test-Path (Join-Path $DEST_FOLDER "python.exe")) {
    Write-Host ""
    Write-Host "  Done. Portable Python ready." -ForegroundColor Green
    Write-Host "  Start-DiscoverySession will now auto-generate HTML reports." -ForegroundColor DarkGray
    Write-Host ""
} else {
    Write-Host ""
    Write-Host "  Extraction completed but python.exe not found." -ForegroundColor Yellow
    Write-Host "  Check: $DEST_FOLDER" -ForegroundColor DarkGray
    Write-Host ""
    exit 1
}
