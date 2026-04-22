<#
.SYNOPSIS
    SDT - zero-install launcher. Paste-to-run.

.DESCRIPTION
    Downloads the latest SDT release into %TEMP%, launches the browser
    GUI. No PATH changes, no persistent install, no Start Menu entries.
    When you close the PowerShell window, the temp copy can be deleted
    or will be cleaned up by Windows on the next temp sweep.

    One-liner to paste on any box:
        iwr https://raw.githubusercontent.com/matt-magna5/SDT/main/run.ps1 -UseBasicParsing | iex

.NOTES
    No admin required. Runs as the logged-in user.
#>
[CmdletBinding()]
param(
    [int]    $Port = 8080,
    [switch] $NoOpenBrowser,
    [string] $Version = 'latest'
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

function Say([string]$m,[string]$c='White') { Write-Host "  $m" -ForegroundColor $c }

Say ""
Say "================================================================" DarkMagenta
Say "  MAGNA5 SDT  -  zero-install launcher" Magenta
Say "================================================================" DarkMagenta
Say ""

# ----- Resolve latest tag ----------------------------------------------------
if ($Version -eq 'latest') {
    Say "Resolving latest release..." DarkCyan
    $Version = $null
    # Query /releases (not /latest) so prereleases/alphas are included.
    try {
        $rel = Invoke-WebRequest 'https://api.github.com/repos/matt-magna5/SDT/releases?per_page=20' -UseBasicParsing -TimeoutSec 10
        $tags = ($rel.Content | ConvertFrom-Json) | ForEach-Object { $_.tag_name }
        # Prefer tags that contain Start-DiscoverySessionGUI.ps1 (v4+). Sort by creation order (first = newest).
        $guiTag = $tags | Where-Object { $_ -match '^v4' -or $_ -match 'alpha' -or $_ -match 'beta' } | Select-Object -First 1
        if ($guiTag) { $Version = $guiTag } else { $Version = $tags[0] }
    } catch {
        Say "Releases API failed: $($_.Exception.Message)" Yellow
        # Probe fallback
        foreach ($try in @('v4.0-alpha','v3.11','v3.10','v3.9','v3.8')) {
            try {
                $h = Invoke-WebRequest "https://github.com/matt-magna5/SDT/archive/refs/tags/$try.zip" -Method Head -UseBasicParsing -TimeoutSec 5
                if ($h.StatusCode -eq 200) { $Version = $try; break }
            } catch { continue }
        }
        if (-not $Version) { throw "Could not determine SDT version" }
    }
    Say "Version: $Version" DarkGreen
}

# ----- Stage to temp ---------------------------------------------------------
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$workDir = Join-Path $env:TEMP "sdt-$stamp"
New-Item -ItemType Directory -Path $workDir -Force | Out-Null
$zipPath = Join-Path $workDir 'sdt.zip'

Say "Downloading $Version to $workDir ..." DarkCyan
Invoke-WebRequest "https://github.com/matt-magna5/SDT/archive/refs/tags/$Version.zip" -OutFile $zipPath -UseBasicParsing -TimeoutSec 120
Expand-Archive -Path $zipPath -DestinationPath $workDir -Force
Remove-Item $zipPath -Force

$appDir = Get-ChildItem $workDir -Directory | Select-Object -First 1
if (-not $appDir) { throw "Extraction failed" }
Set-Location $appDir.FullName
Say "Staged at $($appDir.FullName)" DarkGreen
Say ""

# ----- Launch the GUI --------------------------------------------------------
$gui = Join-Path $appDir.FullName 'Start-DiscoverySessionGUI.ps1'
if (-not (Test-Path $gui)) {
    throw "Start-DiscoverySessionGUI.ps1 not found in release $Version. Try a newer tag with -Version v4.0-alpha (or later)."
}

Say "Launching browser GUI at http://localhost:$Port ..." Yellow
Say "(Close the browser tab + press Ctrl+C here to exit.)" DarkGray
Say ""

$guiArgs = @{ Port = $Port }
if ($NoOpenBrowser) { $guiArgs.NoOpenBrowser = $true }
& $gui @guiArgs
