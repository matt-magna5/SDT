<#
.SYNOPSIS
    Builds the self-contained offline SDT bundle (SDT-Offline-vX.Y.Z.zip).

.DESCRIPTION
    Produces a single zip that installs SDT on a machine with no internet
    access: the scripts, the report generator, and the prepackaged Python /
    plink / wheel dependencies from deps\, plus RUN-SDT.cmd as a double-click
    entry point and README-OFFLINE.txt.

    Client data, local output, and developer-only folders are excluded, so the
    bundle is shippable as-is.

    Run from the repo root:
        .\Make-OfflineBundle.ps1
        .\Make-OfflineBundle.ps1 -OutDir C:\Temp -Version v4.6.0

.NOTES
    The version is read from $script:SessionVersion in
    Start-DiscoverySession.ps1 unless -Version is given. A VERSION file is
    written into the bundle so install.ps1 can stamp it without a tag lookup.
#>
[CmdletBinding()]
param(
    [string] $OutDir,
    [string] $Version
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

$RepoRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
if (-not $OutDir) { $OutDir = Join-Path $RepoRoot '_dist' }

function Say($m, $c = 'Gray') { Write-Host "  $m" -ForegroundColor $c }

Write-Host ""
Write-Host "  MAGNA5 SDT - offline bundle builder" -ForegroundColor Magenta
Write-Host "  ===================================" -ForegroundColor Magenta
Write-Host ""

# ----- Version ---------------------------------------------------------------
if (-not $Version) {
    $sds = Join-Path $RepoRoot 'Start-DiscoverySession.ps1'
    if (-not (Test-Path $sds)) { throw "Start-DiscoverySession.ps1 not found - run this from the repo root." }
    $m = [regex]::Match((Get-Content $sds -Raw), "SessionVersion\s*=\s*'([^']+)'")
    if (-not $m.Success) { throw "Could not read SessionVersion from Start-DiscoverySession.ps1" }
    $Version = 'v' + $m.Groups[1].Value
}
if ($Version -notmatch '^v') { $Version = 'v' + $Version }
Say "Version: $Version" DarkGreen

# ----- What ships ------------------------------------------------------------
# Explicit allow-list. An exclude-list would quietly ship the next folder of
# client data someone drops in the repo.
$files = @(
    'install.ps1'
    'Start-DiscoverySession.ps1'
    'Start-DiscoverySessionGUI.ps1'
    'Invoke-ServerDiscovery.ps1'
    'Get-PortablePython.ps1'
    'gen_report.py'
    '_report_lib.py'
    'clixml_to_json.py'
    'collect_vsphere_perf.py'
    'parse_ntnx_collector.py'
    'detection_rules.json'
    'hardware_eol.json'
    'get-pip.py'
    'README.md'
    'RUN-SDT.cmd'
    'README-OFFLINE.txt'
)
$dirs = @('deps', 'wiki')

$missing = @($files | Where-Object { -not (Test-Path (Join-Path $RepoRoot $_)) })
if ($missing.Count -gt 0) { throw "Missing required file(s): $($missing -join ', ')" }
if (-not (Test-Path (Join-Path $RepoRoot 'deps'))) {
    throw "deps\ is missing - the bundle would not be offline-capable."
}

# ----- Stage -----------------------------------------------------------------
$stage = Join-Path ([IO.Path]::GetTempPath()) ("sdt-bundle-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
$name  = "SDT-Offline-$Version"
$dest  = Join-Path $stage $name
New-Item -ItemType Directory -Force -Path $dest | Out-Null

try {
    foreach ($f in $files) {
        Copy-Item (Join-Path $RepoRoot $f) -Destination $dest -Force
    }
    foreach ($d in $dirs) {
        $sp = Join-Path $RepoRoot $d
        if (Test-Path $sp) { Copy-Item $sp -Destination $dest -Recurse -Force }
    }

    Set-Content -Path (Join-Path $dest 'VERSION') -Value $Version -Encoding ASCII

    # Nothing generated or client-specific may ride along.
    Get-ChildItem $dest -Recurse -Force -EA 0 |
        Where-Object { $_.Name -eq '__pycache__' -or $_.Extension -in @('.pyc', '.html', '.xlsx') } |
        Remove-Item -Recurse -Force -EA 0

    # ----- Verify the staged tree is actually self-sufficient ----------------
    $checks = @(
        @{ N = 'install.ps1';         P = (Join-Path $dest 'install.ps1') }
        @{ N = 'RUN-SDT.cmd';         P = (Join-Path $dest 'RUN-SDT.cmd') }
        @{ N = 'README-OFFLINE.txt';  P = (Join-Path $dest 'README-OFFLINE.txt') }
        @{ N = 'VERSION';             P = (Join-Path $dest 'VERSION') }
        @{ N = 'deps\SHA256SUMS.txt'; P = (Join-Path $dest 'deps\SHA256SUMS.txt') }
        @{ N = 'deps\plink.exe';      P = (Join-Path $dest 'deps\plink.exe') }
    )
    $bad = @()
    foreach ($c in $checks) { if (-not (Test-Path $c.P)) { $bad += $c.N } }
    $py = @(Get-ChildItem (Join-Path $dest 'deps') -Filter 'python-*-embed-*.zip' -EA 0)
    if ($py.Count -eq 0) { $bad += 'deps\python-*-embed-*.zip' }
    $wh = @(Get-ChildItem (Join-Path $dest 'deps\wheels') -Filter '*.whl' -EA 0)
    if ($wh.Count -eq 0) { $bad += 'deps\wheels\*.whl' }
    if ($bad.Count -gt 0) { throw "Bundle is incomplete, refusing to ship: $($bad -join ', ')" }
    Say "Contents verified ($($wh.Count) wheel(s), $($py[0].Name))" DarkGreen

    # ----- Zip ---------------------------------------------------------------
    New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
    $zip = Join-Path $OutDir "$name.zip"
    if (Test-Path $zip) { Remove-Item $zip -Force }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [IO.Compression.ZipFile]::CreateFromDirectory(
        $dest, $zip, [IO.Compression.CompressionLevel]::Optimal, $true)

    $mb  = [math]::Round((Get-Item $zip).Length / 1MB, 1)
    $sha = (Get-FileHash $zip -Algorithm SHA256).Hash

    Write-Host ""
    Say "Bundle: $zip" Green
    Say "Size:   $mb MB" DarkGray
    Say "SHA256: $sha" DarkGray
    Write-Host ""
    Say "Ship this one file. On the target: unblock, extract, run RUN-SDT.cmd." DarkCyan
    Write-Host ""
    return $zip
}
finally {
    Remove-Item $stage -Recurse -Force -EA 0
}
