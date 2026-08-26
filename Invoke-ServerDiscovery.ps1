<#
.SYNOPSIS
    Magna5 Solutions Engineering - Full Server Discovery v2.0
.DESCRIPTION
    Comprehensive server discovery for lift-and-shift, migration, and upgrade
    assessments. Collects OS, roles, SQL, Exchange, applications, shares,
    network, and more. Outputs structured JSON for HTML report generation.

    Safe to run on any Windows Server 2008 R2+. Read-only. No changes made.

.PARAMETER ComputerName
    Target server. Default = local machine.
    Remote requires WinRM enabled on target and admin credentials.
    To enable WinRM on target: winrm quickconfig -y

.PARAMETER OutputPath
    Directory for JSON output. Default = script directory.

.PARAMETER Credential
    PSCredential for remote execution. Prompts if not provided and remote.

.EXAMPLE
    # Local
    .\Invoke-ServerDiscovery.ps1

    # Remote (you'll be prompted for creds)
    .\Invoke-ServerDiscovery.ps1 -ComputerName SRV-APP01

    # Remote with creds pre-loaded
    $cred = Get-Credential
    .\Invoke-ServerDiscovery.ps1 -ComputerName SRV-APP01 -Credential $cred

    # Multi-server loop
    @("SRV-APP01","SRV-SQL01","SRV-DC01") | ForEach-Object {
        .\Invoke-ServerDiscovery.ps1 -ComputerName $_
    }
#>
[CmdletBinding()]
param(
    [string]       $ComputerName = $env:COMPUTERNAME,
    [string]       $OutputPath   = $PSScriptRoot,
    [PSCredential] $Credential,
    [switch]       $NonInteractive,
    # Hyper-V PowerShell Direct transport: when set, the collection block is sent
    # via Invoke-Command -VMId/-VMName (in-band via VMBus) instead of WinRM. The
    # orchestrator uses this as the first leg of the cascade for guests on the
    # local Hyper-V host (no network, no firewall, no per-guest WinRM trust).
    # Requires Hyper-V host 2016+ AND Windows guest 2016+/Win10+ with integration
    # services. Credentials are still required (PS Direct authenticates AS A
    # LOCAL guest account -- domain creds work if the guest is domain-joined).
    [string]       $HyperVGuestVMId,
    [string]       $HyperVGuestVMName
)

$ErrorActionPreference = 'Continue'
# When invoked from the GUI (inside a ThreadJob), there's no interactive
# host - any Read-Host / Get-Credential call throws "host does not support
# user interaction" and kills the scan before it even starts. Detect both
# the explicit flag and the lack of a real user interactive session.
if ($NonInteractive -or -not [Environment]::UserInteractive) {
    $script:IsNonInteractive = $true
} else {
    $script:IsNonInteractive = $false
}
$script:ScriptVersion  = '4.2.35'
$script:StartTime      = Get-Date
$script:CollectErrors  = [System.Collections.ArrayList]@()

# Fix 1 -- TLS 1.2 for older .NET (ensures HTTPS downloads work on .NET 4.0 / PS 3/4)
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

function Expand-ZipCompat([string]$ZipPath, [string]$Destination) {
    if (-not (Test-Path $Destination)) { New-Item -ItemType Directory -Force -Path $Destination | Out-Null }
    if (Get-Command Expand-Archive -ErrorAction SilentlyContinue) {
        Expand-Archive -Path $ZipPath -DestinationPath $Destination -Force
    } else {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $Destination)
    }
}

$script:IsHVGuest = [bool]($HyperVGuestVMId -or $HyperVGuestVMName)
$script:IsRemote  = $script:IsHVGuest -or (
                     $ComputerName -ne $env:COMPUTERNAME -and
                     $ComputerName -ne 'localhost'         -and
                     $ComputerName -ne '127.0.0.1'         -and
                     $ComputerName -ne '.')

# -- AUTO-UPDATE (LOUD) --------------------------------------------------------
# Check GitHub for a newer tag EVERY run. If newer, download, replace this
# script, refresh detection rules, and re-launch. Prints status on every run
# so failures are visible. Opt out: set $env:SDT_NO_AUTOUPDATE = '1'.
function Invoke-SelfUpdate {
    if ($script:IsRemote) {
        Write-Host ("  [auto-update] skipped (remote mode)") -ForegroundColor DarkGray
        return
    }
    if ($env:SDT_NO_AUTOUPDATE -eq '1') {
        Write-Host ("  [auto-update] skipped (SDT_NO_AUTOUPDATE=1)") -ForegroundColor DarkGray
        return
    }
    if ($env:SDT_UPDATED_CHILD -eq '1') { return }   # prevent re-launch loop
    Write-Host ("  [auto-update] checking GitHub for newer version (local v{0})..." -f $script:ScriptVersion) -ForegroundColor DarkCyan
    $latest = $null
    try {
        $ProgressPreference = 'SilentlyContinue'
        # Try releases API first (auth'd requests have higher limits; anon 60/hr)
        $resp = Invoke-WebRequest `
            -Uri 'https://api.github.com/repos/matt-magna5/SDT/releases/latest' `
            -UseBasicParsing -TimeoutSec 8 -ErrorAction Stop
        $latest = ((($resp.Content | ConvertFrom-Json).tag_name) -replace '^v','').Trim()
    } catch {
        Write-Host ("  [auto-update] releases API failed: {0}" -f $_.Exception.Message) -ForegroundColor DarkYellow
        # Fallback: probe common next-version zip URLs (covers rate-limit case)
        try {
            $cur = [version]$script:ScriptVersion
            for ($minor = $cur.Minor + 1; $minor -lt $cur.Minor + 10; $minor++) {
                $try = "{0}.{1}" -f $cur.Major, $minor
                $zu  = "https://github.com/matt-magna5/SDT/archive/refs/tags/v$try.zip"
                try {
                    $h = Invoke-WebRequest -Uri $zu -Method Head -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop
                    if ($h.StatusCode -eq 200) { $latest = $try }
                } catch { break }
            }
            if ($latest) { Write-Host ("  [auto-update] fallback found v{0}" -f $latest) -ForegroundColor DarkCyan }
        } catch { }
    }
    if (-not $latest) {
        Write-Host "  [auto-update] unable to determine latest version - proceeding with current" -ForegroundColor DarkYellow
        return
    }
    try { $newer = ([version]$latest -gt [version]$script:ScriptVersion) } catch { $newer = $false }
    if (-not $newer) {
        Write-Host ("  [auto-update] up to date (v{0})" -f $script:ScriptVersion) -ForegroundColor DarkGreen
        return
    }
    Write-Host ("  [auto-update] v{0} available (local v{1}) - downloading..." -f $latest, $script:ScriptVersion) -ForegroundColor Yellow
    try {
        $zipUrl = "https://github.com/matt-magna5/SDT/archive/refs/tags/v$latest.zip"
        $zipTmp = Join-Path $env:TEMP "sdt-inv-update.zip"
        $extTmp = Join-Path $env:TEMP ("sdt-inv-upd-{0}-{1}" -f $latest, [guid]::NewGuid().ToString('N').Substring(0,6))
        Invoke-WebRequest -Uri $zipUrl -OutFile $zipTmp -UseBasicParsing -TimeoutSec 60 -ErrorAction Stop
        if (Test-Path $extTmp) { Remove-Item $extTmp -Recurse -Force -EA SilentlyContinue }
        Expand-ZipCompat $zipTmp $extTmp
        Remove-Item $zipTmp -Force -ErrorAction SilentlyContinue

        $srcDir = Get-ChildItem $extTmp -Directory | Select-Object -First 1
        if (-not $srcDir) { throw "extracted folder not found" }
        $newInv = Join-Path $srcDir.FullName 'Invoke-ServerDiscovery.ps1'
        if (-not (Test-Path $newInv)) { throw "new Invoke-ServerDiscovery.ps1 not found in zip" }
        $size = (Get-Item $newInv).Length
        $body = Get-Content $newInv -Raw -ErrorAction SilentlyContinue
        if ($size -lt 10240) { throw "new file too small ($size bytes)" }
        if ($body -notmatch "ScriptVersion\s*=\s*'$latest'") { throw "version string v$latest not present in downloaded script" }

        Copy-Item $newInv $PSCommandPath -Force
        foreach ($aux in @('detection_rules.json','hardware_eol.json','gen_report.py')) {
            $src = Join-Path $srcDir.FullName $aux
            $dst = Join-Path (Split-Path -Parent $PSCommandPath) $aux
            if (Test-Path $src) { Copy-Item $src $dst -Force }
        }
        Remove-Item $extTmp -Recurse -Force -EA SilentlyContinue

        Write-Host ("  [auto-update] updated to v{0} - re-launching..." -f $latest) -ForegroundColor Green
        $env:SDT_UPDATED_CHILD = '1'
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath @PSBoundParameters
        $ec = $LASTEXITCODE
        $env:SDT_UPDATED_CHILD = $null
        exit $ec
    } catch {
        Write-Host ("  [auto-update] update failed: {0} - running current v{1}" -f $_.Exception.Message, $script:ScriptVersion) -ForegroundColor Red
    }
}
Invoke-SelfUpdate

# -- PS VERSION & CAPABILITY PROBE ---------------------------------------------

$localPSMajor = $PSVersionTable.PSVersion.Major
$localPSMinor = $PSVersionTable.PSVersion.Minor
$localPSStr   = "$localPSMajor.$localPSMinor"

# -- BANNER --------------------------------------------------------------------

Write-Host ""
Write-Host ("=" * 72) -ForegroundColor DarkMagenta
Write-Host "  MAGNA5 SERVER DISCOVERY  v$script:ScriptVersion" -ForegroundColor Magenta
Write-Host ("=" * 72) -ForegroundColor DarkMagenta
Write-Host ("  Target    : " + $ComputerName.ToUpper()) -ForegroundColor Cyan
Write-Host ("  Started   : " + $script:StartTime.ToString("yyyy-MM-dd HH:mm:ss")) -ForegroundColor Gray
Write-Host ("  PS Local  : $localPSStr  |  Mode: " + $(if ($script:IsRemote) { "REMOTE via WinRM" } else { "LOCAL" })) -ForegroundColor $(if ($localPSMajor -ge 4) { "Green" } elseif ($localPSMajor -ge 3) { "Yellow" } else { "Red" })
Write-Host ("=" * 72) -ForegroundColor DarkMagenta
Write-Host ""

# -- PS COMPAT CHECK + CERT BYPASS ---------------------------------------------

if ($localPSMajor -lt 3) {
    Write-Host "  (x_x)  PowerShell $localPSStr is not supported." -ForegroundColor Red
    Write-Host "         Minimum: PS 3.0  |  Recommended: PS 5.1 or PS 7+" -ForegroundColor DarkRed
    exit 1
}
if ($localPSMajor -eq 3 -or $localPSMajor -eq 4) {
    Write-Host "  (>_<)  PowerShell $localPSStr ? limited support. Date conversions may degrade." -ForegroundColor Yellow
    Write-Host "         Recommend PS 5.1+ for full functionality." -ForegroundColor DarkYellow
    Write-Host ""
}
if ($localPSMajor -ge 5 -and $localPSMajor -lt 6) {
    Write-Host "  (^_^)  PowerShell $localPSStr ? compatible (PS 5.1)." -ForegroundColor DarkGreen
} elseif ($localPSMajor -ge 7) {
    Write-Host "  (^_^)  PowerShell $localPSStr ? full compatibility." -ForegroundColor DarkGreen
}

# Cert bypass for PS 5.1 ? self-signed certs on ESXi/vCenter won't block REST calls
if ($localPSMajor -lt 6) {
    try {
        Add-Type -TypeDefinition @"
using System.Net; using System.Security.Cryptography.X509Certificates;
public class M5TrustAllDisc : ICertificatePolicy {
    public bool CheckValidationResult(ServicePoint s, X509Certificate c, WebRequest r, int p) { return true; }
}
"@ -ErrorAction SilentlyContinue
        [System.Net.ServicePointManager]::CertificatePolicy = New-Object M5TrustAllDisc
        [System.Net.ServicePointManager]::SecurityProtocol  = [System.Net.SecurityProtocolType]::Tls12
    } catch { }
}

# -- BUDDY SYSTEM --------------------------------------------------------------

$buddyFrames = @(
    "(^_^) ", "(^_^)>", "(o_o) ", "(o_o)>",
    "(-_-) ", "(>_<) ", "(*_*) ", "(^_-) ",
    "(._.) ", "(T_T) ", "(^o^) ", "(x_x) "
)

$buddyPhaseLines = @{
    System      = "figuring out what OS this actually is..."
    Hardware    = "counting cores and sticks of RAM..."
    Disks       = "looking at these very full drives..."
    Network     = "untangling the network config..."
    Roles       = "cataloging the chaos they've enabled..."
    AD          = "asking Active Directory nicely for info..."
    DNS         = "querying DNS (someone had to)..."
    DHCP        = "counting leases like a landlord at rent day..."
    Shares      = "snooping through all the file shares..."
    NPS         = "interrogating the RADIUS server..."
    IIS         = "checking which websites are still alive..."
    SQL         = "finding all the databases they forgot about..."
    Exchange    = "checking if Exchange is still breathing..."
    HyperV      = "counting all the VMs they spun up and forgot..."
    Apps        = "cataloging the software graveyard..."
    Tasks       = "reading the scheduler's diary..."
    Services    = "checking which services actually showed up today..."
    EventLog    = "reading the error log (this may hurt)..."
    Netstat     = "snooping on open connections..."
    Printers    = "looking for printers from 2004..."
    Remote      = "waiting for remote server to respond..."
}

function Write-BuddyPhase {
    param([string]$Phase, [string]$Override = "")
    $frame   = $buddyFrames[(Get-Random -Maximum $buddyFrames.Count)]
    $comment = if ($Override) { $Override } elseif ($buddyPhaseLines.ContainsKey($Phase)) { $buddyPhaseLines[$Phase] } else { "working on it..." }
    Write-Host ""
    Write-Host ("  -- " + $Phase.ToUpper() + " " + ("-" * [Math]::Max(1, 56 - $Phase.Length))) -ForegroundColor DarkMagenta
    Write-Host ("  $frame  $comment") -ForegroundColor DarkCyan
}

function Write-BuddyOK   { param([string]$msg) Write-Host ("  (^_^)  $msg") -ForegroundColor DarkGreen  }
function Write-BuddyWarn { param([string]$msg) Write-Host ("  (>_<)  $msg") -ForegroundColor DarkYellow }
function Write-BuddyErr  { param([string]$ctx, [string]$msg)
    Write-Host ("  (x_x)  [$ctx] $msg") -ForegroundColor DarkRed
    [void]$script:CollectErrors.Add("[$ctx] $msg")
}

# -- REMOTE PREFLIGHT ----------------------------------------------------------

if ($script:IsRemote) {
    Write-BuddyPhase "Remote" "remote mode - testing WinRM on $ComputerName..."

    try {
        Test-WSMan -ComputerName $ComputerName -ErrorAction Stop | Out-Null
        Write-BuddyOK "WinRM reachable on $ComputerName"
    } catch {
        Write-BuddyErr "RemotePreflight" "WinRM not reachable: $_"
        if ($script:IsNonInteractive) {
            Write-Host "  (non-interactive mode) proceeding anyway - actual connect will tell us" -ForegroundColor DarkGray
        } else {
            Write-Host ""
            Write-Host "  Run these on the TARGET server to enable WinRM:" -ForegroundColor Yellow
            Write-Host "    winrm quickconfig -y" -ForegroundColor DarkGray
            Write-Host "    Enable-PSRemoting -Force" -ForegroundColor DarkGray
            Write-Host "    # If workgroup (not domain-joined):" -ForegroundColor DarkGray
            Write-Host "    Set-Item WSMan:\localhost\Client\TrustedHosts -Value '*' -Force" -ForegroundColor DarkGray
            Write-Host ""
            $ans = Read-Host "  Press Enter to try anyway, or Ctrl+C to abort"
        }
    }

    # -- CREDENTIAL PROMPT + VALIDATION LOOP ------------------------------------
    $credValid = $false
    $credAttempts = 0
    while (-not $credValid) {
        $credAttempts++
        if ($credAttempts -gt 5) {
            Write-BuddyErr "RemotePreflight" "Too many failed credential attempts - aborting."
            exit 1
        }

        if (-not $Credential) {
            if ($script:IsNonInteractive) {
                Write-BuddyErr "RemotePreflight" "No credentials passed and non-interactive mode - GUI should have supplied -Credential. Skipping."
                exit 1
            }
            Write-Host ""
            Write-Host "  Credentials needed for $ComputerName" -ForegroundColor Yellow
            Write-Host "  (format: DOMAIN\user  or  .\localadmin)" -ForegroundColor DarkGray
            try {
                $Credential = Get-Credential -Message "Admin credentials for $ComputerName"
                if (-not $Credential) { throw "No credentials entered" }
            } catch {
                Write-BuddyErr "RemotePreflight" "No credentials provided - cannot continue."
                exit 1
            }
        }

        Write-Host ("  (^_^)  Testing credentials against $ComputerName...") -ForegroundColor DarkCyan
        try {
            $testParams = @{ ComputerName = $ComputerName; Credential = $Credential; ErrorAction = 'Stop'; ScriptBlock = { $env:COMPUTERNAME } }
            $testResult = Invoke-Command @testParams
            Write-BuddyOK "Credentials verified - connected as $($Credential.UserName) -> $testResult responded."
            $credValid = $true
        } catch {
            $errMsg = $_.ToString()

            # -- AUTO-FIX: TrustedHosts (IP address / non-domain target) ----------
            # This fires when connecting by IP or to a non-domain machine.
            # WinRM requires the target be in TrustedHosts for NTLM auth.
            # Nothing wrong with the password - just need to whitelist the target.
            if ($errMsg -match 'TrustedHosts' -or $errMsg -match 'authentication scheme') {
                Write-Host ""
                Write-BuddyWarn "WinRM TrustedHosts issue detected (not a bad password)."
                Write-Host "  Adding $ComputerName to local WinRM TrustedHosts..." -ForegroundColor DarkCyan
                try {
                    $current = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction Stop).Value
                    if ($current -notmatch [regex]::Escape($ComputerName)) {
                        $newVal = if ($current -and $current.Trim() -ne '') { "$current,$ComputerName" } else { $ComputerName }
                        Set-Item WSMan:\localhost\Client\TrustedHosts -Value $newVal -Force -ErrorAction Stop
                        Write-BuddyOK "TrustedHosts updated. Retrying connection..."
                    } else {
                        Write-BuddyWarn "$ComputerName already in TrustedHosts. Retrying anyway..."
                    }
                    # Retry the test immediately - don't ask the user anything
                    $testResult = Invoke-Command @testParams
                    Write-BuddyOK "Credentials verified after TrustedHosts fix -> $testResult responded."
                    $credValid = $true
                } catch {
                    $errMsg2 = $_.ToString()
                    Write-Host ""
                    Write-Host ("  (x_x)  Still failing after TrustedHosts fix: $errMsg2") -ForegroundColor Red
                    Write-Host "  This may be an actual credential error now." -ForegroundColor DarkGray
                    Write-Host ""
                    if ($script:IsNonInteractive) {
                        Write-BuddyErr "RemotePreflight" "WinRM auth failed after TrustedHosts fix: $errMsg2 (non-interactive - not retrying)"
                        exit 1
                    }
                    Write-Host "    [R] Retry with different credentials" -ForegroundColor DarkGray
                    Write-Host "    [Q] Quit" -ForegroundColor DarkGray
                    $ans2 = Read-Host "  Choice (R/Q)"
                    if ($ans2.Trim().ToUpper() -eq 'R') { $Credential = $null }
                    else { Write-Host "  Exiting." -ForegroundColor Gray; exit 1 }
                }
                continue
            }

            # -- ACCESS DENIED = bad password / wrong user ------------------------
            if ($errMsg -match 'Access is denied' -or $errMsg -match 'AccessDenied' -or $errMsg -match 'LogonFailure' -or $errMsg -match '0x80070005') {
                Write-Host ""
                Write-Host ("  (x_x)  ACCESS DENIED - wrong username or password.") -ForegroundColor Red
                Write-Host ("         User: $($Credential.UserName)") -ForegroundColor DarkRed
                if ($script:IsNonInteractive) {
                    Write-BuddyErr "RemotePreflight" "Access denied for $($Credential.UserName) on $ComputerName (non-interactive - not retrying)"
                    exit 1
                }
                Write-Host ""
                Write-Host "    [R] Retry with different credentials" -ForegroundColor DarkGray
                Write-Host "    [Q] Quit" -ForegroundColor DarkGray
                $ans = Read-Host "  Choice (R/Q)"
                if ($ans.Trim().ToUpper() -eq 'R') { $Credential = $null }
                else { Write-Host "  Exiting." -ForegroundColor Gray; exit 1 }
                continue
            }

            # -- ALL OTHER FAILURES -----------------------------------------------
            Write-Host ""
            Write-Host ("  (x_x)  Connection test failed:") -ForegroundColor Red
            Write-Host ("         $errMsg") -ForegroundColor DarkRed
            if ($script:IsNonInteractive) {
                Write-BuddyErr "RemotePreflight" "Connection test failed on ${ComputerName}: $errMsg (non-interactive - aborting this target)"
                exit 1
            }
            Write-Host ""
            Write-Host "  What would you like to do?" -ForegroundColor Yellow
            Write-Host "    [R] Retry with different credentials" -ForegroundColor DarkGray
            Write-Host "    [S] Skip test and try anyway (risky)" -ForegroundColor DarkGray
            Write-Host "    [Q] Quit" -ForegroundColor DarkGray
            $ans = Read-Host "  Choice (R/S/Q)"
            switch ($ans.Trim().ToUpper()) {
                'R' { $Credential = $null; continue }
                'S' {
                    Write-BuddyWarn "Skipping credential validation - proceeding anyway."
                    $credValid = $true
                }
                default { Write-Host "  Exiting." -ForegroundColor Gray; exit 1 }
            }
        }
    }
    Write-Host "  (This runs as a single remote job - buddy will animate while waiting)" -ForegroundColor DarkGray
}

# -----------------------------------------------------------------------------
# COLLECTION SCRIPTBLOCK - self-contained, runs local or remote unchanged
# -----------------------------------------------------------------------------

$CollectionBlock = {

    $ErrorActionPreference = 'Continue'

    # -- INTERNAL HELPERS ------------------------------------------------------

    $cbErrors = [System.Collections.ArrayList]@()
    $cbFlags  = [System.Collections.ArrayList]@()

    function cb-Log  { param([string]$ctx,[string]$msg) [void]$cbErrors.Add("[$ctx] $msg") }
    function cb-Flag { param([string]$sev,[string]$title,[string]$detail)
        [void]$cbFlags.Add([PSCustomObject]@{ Severity=$sev; Title=$title; Detail=$detail })
    }

    # PS version inside the target
    $PSMaj = $PSVersionTable.PSVersion.Major
    $PSMin = $PSVersionTable.PSVersion.Minor

    # -- LOOKUP TABLES (embedded - safe for remote) -----------------------------

    $OSEOLMap = @{
        '5.2' = @{ Name='Windows Server 2003';    EOL='2015-07-14'; Status='EOL'       }
        '6.0' = @{ Name='Windows Server 2008';    EOL='2020-01-14'; Status='EOL'       }
        '6.1' = @{ Name='Windows Server 2008 R2'; EOL='2020-01-14'; Status='EOL'       }
        '6.2' = @{ Name='Windows Server 2012';    EOL='2023-10-10'; Status='EOL'       }
        '6.3' = @{ Name='Windows Server 2012 R2'; EOL='2023-10-10'; Status='EOL'       }
    }
    $OSBuildMap = @{
        '10.0.14393' = @{ Name='Windows Server 2016'; EOL='2027-01-12'; Status='Supported'  }
        '10.0.17763' = @{ Name='Windows Server 2019'; EOL='2029-01-09'; Status='Supported'  }
        '10.0.20348' = @{ Name='Windows Server 2022'; EOL='2031-10-14'; Status='Supported'  }
        '10.0.26100' = @{ Name='Windows Server 2025'; EOL='2034-10-10'; Status='Supported'  }
    }
    $SQLEOLMap = @{
        '8.00'  = @{ Name='SQL Server 2000';     EOL='2013-04-09'; Status='EOL' }
        '9.00'  = @{ Name='SQL Server 2005';     EOL='2016-04-12'; Status='EOL' }
        '10.00' = @{ Name='SQL Server 2008';     EOL='2019-07-09'; Status='EOL' }
        '10.50' = @{ Name='SQL Server 2008 R2';  EOL='2019-07-09'; Status='EOL' }
        '11.0'  = @{ Name='SQL Server 2012';     EOL='2022-07-12'; Status='EOL' }
        '12.0'  = @{ Name='SQL Server 2014';     EOL='2024-07-09'; Status='EOL' }
        '13.0'  = @{ Name='SQL Server 2016';     EOL='2026-07-14'; Status='Supported' }
        '14.0'  = @{ Name='SQL Server 2017';     EOL='2027-10-12'; Status='Supported' }
        '15.0'  = @{ Name='SQL Server 2019';     EOL='2030-01-08'; Status='Supported' }
        '16.0'  = @{ Name='SQL Server 2022';     EOL='2033-01-11'; Status='Supported' }
    }
    $ExchEOLMap = @{
        '6'    = @{ Name='Exchange 2003'; EOL='2009-04-14'; Status='EOL'      }
        '8'    = @{ Name='Exchange 2007'; EOL='2017-04-11'; Status='EOL'      }
        '14'   = @{ Name='Exchange 2010'; EOL='2020-10-13'; Status='EOL'      }
        '15.0' = @{ Name='Exchange 2013'; EOL='2023-04-11'; Status='EOL'      }
        '15.1' = @{ Name='Exchange 2016'; EOL='2025-10-14'; Status='Near EOL' }
        '15.2' = @{ Name='Exchange 2019'; EOL='2025-10-14'; Status='Near EOL' }
    }
    $SqlDbAppMap = @{
        'SUSDB'              = 'Windows Server Update Services (WSUS)'
        'ReportServer'       = 'SQL Server Reporting Services (SSRS)'
        'ReportServerTempDB' = 'SSRS temp database'
        'AutotaskAPI'        = 'Autotask PSA'
        'Kaseya'             = 'Kaseya VSA RMM'
        'SWNetPerfMon'       = 'SolarWinds NPM'
        'SWRemote'           = 'SolarWinds Remote Monitoring'
        'ACT7'               = 'ACT! CRM v7'
        'ACT9'               = 'ACT! CRM v9'
        'QBPOS'              = 'QuickBooks Point of Sale'
        'NAVData'            = 'Microsoft Dynamics NAV'
        'DYNAMICS'           = 'Microsoft Dynamics'
        'TimeMatters'        = 'LexisNexis Time Matters (Legal)'
        'PCLaw'              = 'PCLaw Legal Billing'
        'Medisoft'           = 'Medisoft Medical Billing'
        'Kareo'              = 'Kareo Medical Practice Mgmt'
        'Avimark'            = 'Avimark Veterinary Practice'
        'Cornerstone'        = 'IDEXX Cornerstone Veterinary'
        'JobBOSS'            = 'JobBOSS Manufacturing ERP'
        'EPICOR'             = 'Epicor ERP'
        'MAS90'              = 'Sage 100 (MAS 90)'
        'MAS200'             = 'Sage 200 (MAS 200)'
        'Peachtree'          = 'Peachtree/Sage 50'
        'ConnectWise'        = 'ConnectWise Manage/PSA'
        'NinjaRMM'           = 'NinjaRMM'
        'LabTech'            = 'ConnectWise Automate (LabTech)'
        'GWAVA'              = 'GWAVA Email Security'
    }
    $AppFlagRules = @(
        @{ K=@('veeam');                      Cat='Backup';        Sev='red';    Note='Veeam agent - confirm jobs, schedule, and offsite/cloud copy' }
        @{ K=@('acronis');                    Cat='Backup';        Sev='red';    Note='Acronis backup agent' }
        @{ K=@('backup exec');                Cat='Backup';        Sev='red';    Note='Veritas Backup Exec' }
        @{ K=@('datto');                      Cat='Backup';        Sev='red';    Note='DATTO agent installed' }
        @{ K=@('shadowprotect','storagecraft');Cat='Backup';        Sev='red';    Note='StorageCraft ShadowProtect' }
        @{ K=@('carbonite');                  Cat='Backup';        Sev='red';    Note='Carbonite backup' }
        @{ K=@('commvault');                  Cat='Backup';        Sev='red';    Note='Commvault agent' }
        @{ K=@('windows server backup');      Cat='Backup';        Sev='yellow'; Note='Windows Server Backup (built-in) - minimal BDR protection' }
        @{ K=@('teamviewer');                 Cat='RemoteAccess';  Sev='yellow'; Note='TeamViewer - verify authorized and inventoried' }
        @{ K=@('anydesk');                    Cat='RemoteAccess';  Sev='yellow'; Note='AnyDesk - verify authorized' }
        @{ K=@('logmein');                    Cat='RemoteAccess';  Sev='yellow'; Note='LogMeIn remote access' }
        @{ K=@('screenconnect');              Cat='RemoteAccess';  Sev='yellow'; Note='ConnectWise Control (ScreenConnect)' }
        @{ K=@('quickbooks');                 Cat='Accounting';    Sev='red';    Note='QuickBooks - locate .qbw data files and confirm backup coverage' }
        @{ K=@('sage 50','peachtree');        Cat='Accounting';    Sev='red';    Note='Sage 50/Peachtree accounting' }
        @{ K=@('great plains','dynamics gp'); Cat='ERP';           Sev='red';    Note='Microsoft Dynamics GP - confirm SQL database and backup' }
        @{ K=@('dynamics nav','navision');    Cat='ERP';           Sev='red';    Note='Microsoft Dynamics NAV - confirm SQL database and backup' }
        @{ K=@('office 2007');                Cat='Legacy';        Sev='red';    Note='Microsoft Office 2007 - EOL, no security patches since 2017' }
        @{ K=@('office 2010');                Cat='Legacy';        Sev='red';    Note='Microsoft Office 2010 - EOL October 2020' }
        @{ K=@('office 2013');                Cat='Legacy';        Sev='yellow'; Note='Microsoft Office 2013 - EOL April 2023' }
        @{ K=@('mysql');                      Cat='Database';      Sev='yellow'; Note='MySQL Server - catalog databases and confirm backup' }
        @{ K=@('postgresql');                 Cat='Database';      Sev='yellow'; Note='PostgreSQL - catalog databases' }
        @{ K=@('oracle database');            Cat='Database';      Sev='yellow'; Note='Oracle DB - confirm version and support status' }
        @{ K=@('kaseya');                     Cat='RMM';           Sev='info';   Note='Kaseya VSA RMM agent' }
        @{ K=@('labtech','connectwise automate'); Cat='RMM';       Sev='info';   Note='ConnectWise Automate (LabTech) RMM agent' }
        @{ K=@('n-central','n-able');         Cat='RMM';           Sev='info';   Note='N-able/N-central RMM agent' }
        @{ K=@('crowdstrike');                Cat='EDR';           Sev='info';   Note='CrowdStrike Falcon EDR' }
        @{ K=@('sentinelone');                Cat='EDR';           Sev='info';   Note='SentinelOne EDR' }
        @{ K=@('cylance');                    Cat='EDR';           Sev='info';   Note='Cylance/BlackBerry Protect EDR' }
        @{ K=@('vmware tools');               Cat='VM';            Sev='info';   Note='VMware Tools - server is a VMware VM' }
        @{ K=@('symantec endpoint','norton endpoint'); Cat='AV';   Sev='yellow'; Note='Symantec/Norton AV - confirm still licensed and active' }
        @{ K=@('mcafee');                     Cat='AV';            Sev='yellow'; Note='McAfee AV - confirm still licensed' }
        @{ K=@('trend micro');                Cat='AV';            Sev='yellow'; Note='Trend Micro AV' }
        @{ K=@('citrix');                     Cat='Virtualization'; Sev='yellow'; Note='Citrix component - confirm active use' }
        @{ K=@('cisco anyconnect');           Cat='VPN';           Sev='info';   Note='Cisco AnyConnect VPN client' }
        @{ K=@('globalprotect');              Cat='VPN';           Sev='info';   Note='Palo Alto GlobalProtect VPN' }
    )

    # -- HELPER: SAFE WMI/CIM QUERY --------------------------------------------

    function Safe-Wmi {
        param([string]$Class, [string]$Context = "WMI")
        try {
            if ($PSMaj -ge 3) {
                Get-CimInstance -ClassName $Class -ErrorAction Stop
            } else {
                Get-WmiObject -Class $Class -ErrorAction Stop
            }
        } catch {
            cb-Log $Context "WMI query failed for ${Class}: $_"
            return $null
        }
    }

    function Safe-WmiQuery {
        param([string]$Query, [string]$Context = "WMI")
        try {
            if ($PSMaj -ge 3) {
                Get-CimInstance -Query $Query -ErrorAction Stop
            } else {
                Get-WmiObject -Query $Query -ErrorAction Stop
            }
        } catch {
            cb-Log $Context "WMI query failed: $_"
            return $null
        }
    }

    # -- SYSTEM INFO -----------------------------------------------------------

    function Collect-SystemInfo {
        Write-Host "  [System] Collecting OS info..." -ForegroundColor Gray
        $result = @{
            Hostname       = $env:COMPUTERNAME
            Domain         = $env:USERDNSDOMAIN
            RunAsUser      = "$env:USERDOMAIN\$env:USERNAME"
            PSVersion      = "$PSMaj.$PSMin"
            OSName         = 'Unknown'
            OSBuild        = 'Unknown'
            OSVersion      = 'Unknown'
            OSInstallDate  = 'Unknown'
            OSEOLDate      = 'Unknown'
            OSEOLStatus    = 'Unknown'
            LastBoot       = 'Unknown'
            UptimeDays     = 0
            Timezone       = 'Unknown'
            Partial        = $false
        }
        try {
            $os = Safe-Wmi Win32_OperatingSystem System
            if ($os) {
                $result.OSName    = $os.Caption
                $result.OSBuild   = $os.BuildNumber
                $result.OSVersion = $os.Version

                # EOL lookup - try build map first, then major.minor
                $verKey = ($os.Version -split '\.' | Select-Object -First 3) -join '.'
                $majMin = ($os.Version -split '\.' | Select-Object -First 2) -join '.'
                if ($OSBuildMap.ContainsKey($verKey)) {
                    $eol = $OSBuildMap[$verKey]
                } elseif ($OSBuildMap.ContainsKey($majMin)) {
                    $eol = $OSBuildMap[$majMin]
                } elseif ($OSEOLMap.ContainsKey($majMin)) {
                    $eol = $OSEOLMap[$majMin]
                } else {
                    $eol = @{ Name=$os.Caption; EOL='Unknown'; Status='Unknown' }
                }
                $result.OSEOLDate   = $eol.EOL
                $result.OSEOLStatus = $eol.Status

                if ($eol.Status -eq 'EOL') {
                    cb-Flag 'critical' "EOL Operating System: $($os.Caption)" "EOL date: $($eol.EOL). No security patches. Immediate upgrade required before migration."
                } elseif ($eol.Status -eq 'Near EOL') {
                    cb-Flag 'warning' "Near-EOL OS: $($os.Caption)" "EOL date: $($eol.EOL). Plan upgrade within 12 months."
                }

                try {
                    if ($os.InstallDate) {
                        if ($PSMaj -ge 3) {
                            $result.OSInstallDate = $os.InstallDate.ToString("yyyy-MM-dd")
                        } else {
                            $result.OSInstallDate = [System.Management.ManagementDateTimeConverter]::ToDateTime($os.InstallDate).ToString("yyyy-MM-dd")
                        }
                    }
                } catch { $result.Partial = $true }

                try {
                    if ($os.LastBootUpTime) {
                        $boot = if ($PSMaj -ge 3) { $os.LastBootUpTime } else { [System.Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime) }
                        $result.LastBoot    = $boot.ToString("yyyy-MM-dd HH:mm:ss")
                        $result.UptimeDays  = [math]::Round(((Get-Date) - $boot).TotalDays, 1)
                        if ($result.UptimeDays -gt 365) {
                            cb-Flag 'warning' "Server Uptime: $($result.UptimeDays) days" "This server has not been rebooted in over a year. Pending Windows updates and possible deferred reboots."
                        }
                    }
                } catch { $result.Partial = $true }
            }
        } catch {
            cb-Log "SystemInfo" "Outer error: $_"
            $result.Partial = $true
        }
        try { $result.Timezone = (Get-TimeZone -ErrorAction SilentlyContinue).DisplayName } catch { }
        try {
            $domain = (Safe-Wmi Win32_ComputerSystem System)
            if ($domain) {
                $result.Domain = $domain.Domain
            }
        } catch { }
        return $result
    }

    # -- HARDWARE --------------------------------------------------------------

    function Collect-Hardware {
        Write-Host "  [Hardware] Collecting CPU, RAM, VM info..." -ForegroundColor Gray
        $result = @{ CPUName='Unknown'; CPUCores=0; RAMTotalGB=0; RAMAvailGB=0; IsVM=$false; VMPlatform='Physical'; Partial=$false }
        try {
            $cpu = Safe-Wmi Win32_Processor Hardware
            if ($cpu) {
                $first = @($cpu)[0]
                $result.CPUName  = $first.Name.Trim()
                $result.CPUCores = ($cpu | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum
            }
        } catch { $result.Partial = $true }
        try {
            $cs = Safe-Wmi Win32_ComputerSystem Hardware
            if ($cs) {
                $result.RAMTotalGB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
                # VM detection
                $model = $cs.Model + ' ' + $cs.Manufacturer
                if     ($model -match 'VMware')          { $result.IsVM = $true; $result.VMPlatform = 'VMware'    }
                elseif ($model -match 'Virtual Machine')  { $result.IsVM = $true; $result.VMPlatform = 'Hyper-V'  }
                elseif ($model -match 'VirtualBox')       { $result.IsVM = $true; $result.VMPlatform = 'VirtualBox' }
                elseif ($model -match 'KVM|QEMU')         { $result.IsVM = $true; $result.VMPlatform = 'KVM/QEMU' }
                elseif ($model -match 'Xen')              { $result.IsVM = $true; $result.VMPlatform = 'Xen'      }
            }
        } catch { $result.Partial = $true }
        try {
            $os = Safe-Wmi Win32_OperatingSystem Hardware
            if ($os) { $result.RAMAvailGB = [math]::Round($os.FreePhysicalMemory / 1MB, 2) }
        } catch { $result.Partial = $true }
        # Fallback VM check via bios/service
        if (-not $result.IsVM) {
            try {
                $bios = Safe-Wmi Win32_BIOS Hardware
                if ($bios -and $bios.SerialNumber -match 'VMware|Xen') { $result.IsVM = $true; $result.VMPlatform = 'VMware' }
            } catch { }
        }
        # Manufacturer / Identity data (bare-metal focus)
        try {
            $cs2  = Safe-Wmi Win32_ComputerSystem Hardware
            $bios2 = Safe-Wmi Win32_BIOS Hardware
            $board = Safe-Wmi Win32_BaseBoard Hardware
            if ($cs2) {
                $result.Manufacturer = [string]$cs2.Manufacturer
                $result.Model        = [string]$cs2.Model
            }
            if ($bios2) {
                $result.SerialNumber = [string]$bios2.SerialNumber
                $result.BIOSVersion  = [string]$bios2.SMBIOSBIOSVersion
                try {
                    $bd = $bios2.ReleaseDate
                    if ($bd) { $result.BIOSDate = [System.Management.ManagementDateTimeConverter]::ToDateTime($bd).ToString('yyyy-MM-dd') }
                } catch { $result.BIOSDate = [string]$bios2.ReleaseDate }
            }
            if ($board) {
                $result.BoardProduct = [string]$board.Product
                $result.BoardSerial  = [string]$board.SerialNumber
            }
        } catch { }
        return $result
    }

    # -- DISKS -----------------------------------------------------------------

    function Collect-Disks {
        Write-Host "  [Disks] Collecting drive info..." -ForegroundColor Gray
        $disks = [System.Collections.ArrayList]@()
        try {
            $vols = Safe-Wmi Win32_LogicalDisk Disks
            if ($vols) {
                foreach ($v in ($vols | Where-Object { $_.DriveType -eq 3 })) {
                    try {
                        $totalGB = [math]::Round($v.Size / 1GB, 2)
                        $freeGB  = [math]::Round($v.FreeSpace / 1GB, 2)
                        $usedPct = if ($totalGB -gt 0) { [math]::Round((($totalGB - $freeGB) / $totalGB) * 100, 1) } else { 0 }
                        $disk = @{
                            Drive     = $v.DeviceID
                            Label     = $v.VolumeName
                            TotalGB   = $totalGB
                            FreeGB    = $freeGB
                            UsedPct   = $usedPct
                            Filesystem = $v.FileSystem
                        }
                        if ($usedPct -ge 90) {
                            cb-Flag 'critical' "Disk $($v.DeviceID) Critical ($usedPct% full)" "$($v.FreeSpace / 1GB | % { [math]::Round($_,1) }) GB free of $totalGB GB. Migration will fail if disk is full."
                        } elseif ($usedPct -ge 80) {
                            cb-Flag 'warning' "Disk $($v.DeviceID) Near Full ($usedPct%)" "$freeGB GB free of $totalGB GB. Monitor before migration."
                        }
                        [void]$disks.Add($disk)
                    } catch { cb-Log "Disks" "Error on volume $($v.DeviceID): $_" }
                }
            }
        } catch { cb-Log "Disks" "Outer error: $_" }
        return ,$disks
    }

    # -- NETWORK ---------------------------------------------------------------

    function Collect-Network {
        Write-Host "  [Network] Collecting adapters and connections..." -ForegroundColor Gray
        $result = @{ Adapters=@(); ListeningPorts=@(); EstablishedConns=@(); Partial=$false }
        try {
            $adapters = [System.Collections.ArrayList]@()
            $nics = Safe-Wmi Win32_NetworkAdapterConfiguration Network
            if ($nics) {
                foreach ($n in ($nics | Where-Object { $_.IPEnabled -eq $true })) {
                    try {
                        [void]$adapters.Add(@{
                            Description  = $n.Description
                            IPAddresses  = ($n.IPAddress -join ', ')
                            SubnetMasks  = ($n.IPSubnet -join ', ')
                            Gateway      = ($n.DefaultIPGateway -join ', ')
                            DNS          = ($n.DNSServerSearchOrder -join ', ')
                            MAC          = $n.MACAddress
                            DHCPEnabled  = $n.DHCPEnabled
                        })
                    } catch { }
                }
            }
            $result.Adapters = $adapters
        } catch {
            cb-Log "Network" "Adapter collection failed: $_"
            $result.Partial = $true
        }
        try {
            $rawNetstat = netstat -ano 2>&1
            $listening  = [System.Collections.ArrayList]@()
            $established= [System.Collections.ArrayList]@()
            foreach ($line in $rawNetstat) {
                if ($line -match '^\s+(TCP|UDP)\s+(\S+):(\d+)\s+(\S+)\s+(LISTENING|ESTABLISHED|TIME_WAIT)\s+(\d*)') {
                    $proto    = $Matches[1]
                    $localIP  = $Matches[2]
                    $port     = [int]$Matches[3]
                    $remoteEP = $Matches[4]
                    $state    = $Matches[5]
                    $netPid   = $Matches[6]   # $PID is a PS reserved variable - use $netPid
                    # Try to resolve process name
                    $procName = 'Unknown'
                    try {
                        if ($netPid -and $netPid -ne '0') {
                            $p = Get-Process -Id ([int]$netPid) -ErrorAction SilentlyContinue
                            if ($p) { $procName = $p.ProcessName }
                        }
                    } catch { }
                    $entry = @{ Proto=$proto; LocalIP=$localIP; Port=$port; Remote=$remoteEP; State=$state; PID=$netPid; Process=$procName }
                    if ($state -eq 'LISTENING')    { [void]$listening.Add($entry)   }
                    elseif ($state -eq 'ESTABLISHED') { [void]$established.Add($entry) }
                }
            }
            $result.ListeningPorts    = $listening   | Sort-Object Port | Select-Object -First 100
            $result.EstablishedConns  = $established | Sort-Object Port | Select-Object -First 100
        } catch {
            cb-Log "Network" "netstat collection failed: $_"
            $result.Partial = $true
        }
        return $result
    }

    # -- ROLES & FEATURES ------------------------------------------------------

    function Collect-Roles {
        Write-Host "  [Roles] Enumerating server roles and features..." -ForegroundColor Gray
        $result = @{ InstalledRoles=@(); InstalledFeatures=@(); Method='Unknown'; Partial=$false }
        try {
            # Modern path: PS4+ with ServerManager module
            if ($PSMaj -ge 4) {
                try {
                    Import-Module ServerManager -ErrorAction Stop
                    $allFeatures = Get-WindowsFeature -ErrorAction Stop | Where-Object { $_.Installed -eq $true }
                    $result.InstalledRoles    = @($allFeatures | Where-Object { $_.FeatureType -eq 'Role'    } | Select-Object Name, DisplayName, Description)
                    $result.InstalledFeatures = @($allFeatures | Where-Object { $_.FeatureType -eq 'Feature' } | Select-Object Name, DisplayName, Description)
                    $result.Method = 'Get-WindowsFeature'
                    return $result
                } catch {
                    cb-Log "Roles" "Get-WindowsFeature failed, falling back to WMI: $_"
                }
            }
            # Fallback: WMI Win32_ServerFeature (2008+, PS 2+)
            $wmiFeatures = Safe-Wmi Win32_ServerFeature Roles
            if ($wmiFeatures) {
                # IDs for common roles (partial list for WMI fallback)
                $roleIDMap = @{
                    10='AD DS'; 12='DNS Server'; 11='DHCP Server'; 51='IIS'; 6='Hyper-V'
                    33='File Services'; 60='NPS'; 14='Print Services'; 35='AD RMS'
                    13='Streaming Media Services'; 8='Terminal Services / RDS'; 41='WSUS'
                }
                $roles = @($wmiFeatures | ForEach-Object {
                    $name = if ($roleIDMap.ContainsKey($_.ID)) { $roleIDMap[$_.ID] } else { "Feature ID $($_.ID)" }
                    @{ Name=$name; DisplayName=$_.Name; ID=$_.ID }
                })
                $result.InstalledRoles = $roles
                $result.Method = 'Win32_ServerFeature (WMI)'
            }
        } catch {
            cb-Log "Roles" "Outer error: $_"
            $result.Partial = $true
        }
        return $result
    }

    # -- ACTIVE DIRECTORY ------------------------------------------------------

    function Collect-ADDetails {
        Write-Host "  [AD] Querying Active Directory..." -ForegroundColor Gray
        $result = @{ Installed=$false; DomainName=''; ForestName=''; DomainFL=''; ForestFL=''; UserCount=0; ComputerCount=0; OUCount=0; DCCount=0; StaleUsers=@(); StaleComputers=@(); Partial=$false }
        try {
            # Check if AD module is available
            $adAvail = Get-Module -ListAvailable -Name ActiveDirectory -ErrorAction SilentlyContinue
            if (-not $adAvail) {
                # Try via ADSI fallback
                try {
                    $rootDSE = [ADSI]"LDAP://RootDSE"
                    if ($rootDSE.dnsHostName) {
                        $result.Installed = $true
                        $result.DomainName = $rootDSE.defaultNamingContext
                        # Try Get-ADDomain via WMI/nltest
                        try {
                            $nltest = nltest /dclist: 2>&1
                            $dcs = @($nltest | Where-Object { $_ -match '\\' } | ForEach-Object { ($_ -split '\\')[1].Trim() })
                            $result.DCCount = $dcs.Count
                        } catch { }
                        $result.Partial = $true
                        return $result
                    }
                } catch { }
                return $result  # AD not installed/accessible
            }
            Import-Module ActiveDirectory -ErrorAction Stop
            $result.Installed = $true
            try {
                $domain = Get-ADDomain -ErrorAction Stop
                $result.DomainName = $domain.DNSRoot
                $result.DomainFL   = $domain.DomainMode
            } catch { cb-Log "AD" "Get-ADDomain failed: $_"; $result.Partial = $true }
            try {
                $forest = Get-ADForest -ErrorAction Stop
                $result.ForestName = $forest.RootDomain
                $result.ForestFL   = $forest.ForestMode
            } catch { cb-Log "AD" "Get-ADForest failed: $_" }
            try { $result.DCCount       = (Get-ADDomainController -Filter * -ErrorAction Stop).Count      } catch { }
            # FSMO roles held by this DC
            try {
                $domain = Get-ADDomain -ErrorAction Stop
                $forest = Get-ADForest -ErrorAction Stop
                $fsmo = [System.Collections.ArrayList]@()
                $myFQDN = "$env:COMPUTERNAME.$env:USERDNSDOMAIN".ToLower()
                if ($domain.PDCEmulator.ToLower()          -eq $myFQDN) { [void]$fsmo.Add('PDC Emulator') }
                if ($domain.RIDMaster.ToLower()            -eq $myFQDN) { [void]$fsmo.Add('RID Master') }
                if ($domain.InfrastructureMaster.ToLower() -eq $myFQDN) { [void]$fsmo.Add('Infrastructure Master') }
                if ($forest.SchemaMaster.ToLower()         -eq $myFQDN) { [void]$fsmo.Add('Schema Master') }
                if ($forest.DomainNamingMaster.ToLower()   -eq $myFQDN) { [void]$fsmo.Add('Domain Naming Master') }
                $result.FSMORoles       = @($fsmo)
                $result.DomainFL        = [string]$domain.DomainMode
                $result.ForestFL        = [string]$forest.ForestMode
                $result.ForestName      = [string]$forest.RootDomain
                $result.PDCEmulator     = [string]$domain.PDCEmulator
                $result.RIDMaster       = [string]$domain.RIDMaster
                $result.SchemaMaster    = [string]$forest.SchemaMaster
            } catch { }
            # Force array + [int] so the count is always a plain integer. A bare
            # .Count on a scalar/AD result can serialize as an ADPropertyValueCollection
            # object (renders as the .NET type name in the report) instead of a number.
            try { $result.UserCount     = [int]@(Get-ADUser              -Filter * -ErrorAction Stop).Count } catch { }
            try { $result.ComputerCount = [int]@(Get-ADComputer          -Filter * -ErrorAction Stop).Count } catch { }
            try { $result.OUCount       = [int]@(Get-ADOrganizationalUnit -Filter * -ErrorAction Stop).Count } catch { }
            # Stale accounts (90 days inactive). CRITICAL: convert each result to
            # a plain hashtable BEFORE it enters $result. AD module returns .NET
            # types with PSParameterizedProperty indexer members that break JSON.
            try {
                $cutoff = (Get-Date).AddDays(-90)
                $staleUsersRaw = Get-ADUser -Filter { LastLogonDate -lt $cutoff -and Enabled -eq $true } -Properties LastLogonDate -ErrorAction Stop | Select-Object -First 50
                $staleUsersClean = @()
                foreach ($u in $staleUsersRaw) {
                    $ll = ''
                    try { if ($u.LastLogonDate) { $ll = $u.LastLogonDate.ToString('yyyy-MM-dd HH:mm:ss') } } catch { }
                    $staleUsersClean += @{
                        Name           = [string]$u.Name
                        SamAccountName = [string]$u.SamAccountName
                        LastLogon      = $ll
                    }
                }
                $result.StaleUsers = $staleUsersClean
                if ($staleUsersClean.Count -gt 0) {
                    cb-Flag 'warning' "Stale AD Accounts: $($staleUsersClean.Count)" "$($staleUsersClean.Count) enabled users haven't logged in for 90+ days. Review before migration."
                }
            } catch { }
            try {
                $cutoff = (Get-Date).AddDays(-90)
                $staleCompsRaw = Get-ADComputer -Filter { LastLogonDate -lt $cutoff -and Enabled -eq $true } -Properties LastLogonDate -ErrorAction Stop | Select-Object -First 50
                $staleCompsClean = @()
                foreach ($c in $staleCompsRaw) {
                    $ll = ''
                    try { if ($c.LastLogonDate) { $ll = $c.LastLogonDate.ToString('yyyy-MM-dd HH:mm:ss') } } catch { }
                    $staleCompsClean += @{
                        Name           = [string]$c.Name
                        SamAccountName = [string]$c.SamAccountName
                        LastLogon      = $ll
                    }
                }
                $result.StaleComputers = $staleCompsClean
            } catch { }
        } catch {
            cb-Log "AD" "Outer error: $_"
            $result.Partial = $true
        }
        return $result
    }

    # -- DNS -------------------------------------------------------------------

    function Collect-DNSDetails {
        Write-Host "  [DNS] Querying DNS zones..." -ForegroundColor Gray
        $result = @{ Installed=$false; Zones=@(); Forwarders=@(); Partial=$false }
        try {
            if (Get-Module -ListAvailable -Name DnsServer -ErrorAction SilentlyContinue) {
                Import-Module DnsServer -ErrorAction Stop
                $result.Installed = $true
                try { $result.Zones      = @(Get-DnsServerZone -ErrorAction Stop | Select-Object ZoneName, ZoneType, IsDsIntegrated, IsReverseLookupZone) } catch { }
                try { $result.Forwarders = @(Get-DnsServerForwarder -ErrorAction Stop | Select-Object -ExpandProperty IPAddress) } catch { }
            } else {
                # Fallback: check if DNS service is running
                $dnsSvc = Get-Service -Name DNS -ErrorAction SilentlyContinue
                if ($dnsSvc -and $dnsSvc.Status -eq 'Running') {
                    $result.Installed = $true
                    $result.Partial   = $true
                    # Try dnscmd fallback
                    try {
                        $dnscmdOut = dnscmd /enumzones 2>&1
                        $result.Zones = @($dnscmdOut | Where-Object { $_ -match '^\s+\w' } | ForEach-Object { @{ ZoneName=$_.Trim() } })
                    } catch { }
                }
            }
        } catch {
            cb-Log "DNS" "Outer error: $_"
            $result.Partial = $true
        }
        return $result
    }

    # -- DHCP ------------------------------------------------------------------

    function Collect-DHCPDetails {
        Write-Host "  [DHCP] Querying DHCP scopes..." -ForegroundColor Gray
        $result = @{ Installed=$false; Scopes=@(); Partial=$false }
        try {
            if (Get-Module -ListAvailable -Name DhcpServer -ErrorAction SilentlyContinue) {
                Import-Module DhcpServer -ErrorAction Stop
                $result.Installed = $true
                try {
                    $scopes = Get-DhcpServerv4Scope -ErrorAction Stop
                    $result.Scopes = @($scopes | ForEach-Object {
                        try {
                            $stats = Get-DhcpServerv4ScopeStatistics -ScopeId $_.ScopeId -ErrorAction SilentlyContinue
                            @{
                                ScopeId     = $_.ScopeId.ToString()
                                Name        = $_.Name
                                SubnetMask  = $_.SubnetMask.ToString()
                                State       = $_.State
                                StartRange  = $_.StartRange.ToString()
                                EndRange    = $_.EndRange.ToString()
                                InUse       = if ($stats) { $stats.InUse } else { '?' }
                                Available   = if ($stats) { $stats.Free } else { '?' }
                            }
                        } catch { @{ ScopeId=$_.ScopeId.ToString(); Name=$_.Name; Partial=$true } }
                    })
                } catch { cb-Log "DHCP" "Get-DhcpServerv4Scope failed: $_"; $result.Partial = $true }
            } else {
                $dhcpSvc = Get-Service -Name DHCPServer -ErrorAction SilentlyContinue
                if ($dhcpSvc -and $dhcpSvc.Status -eq 'Running') {
                    $result.Installed = $true; $result.Partial = $true
                }
            }
        } catch {
            cb-Log "DHCP" "Outer error: $_"
            $result.Partial = $true
        }
        return $result
    }

    # -- FILE SHARES -----------------------------------------------------------

    function Collect-FileShares {
        Write-Host "  [Shares] Enumerating file shares..." -ForegroundColor Gray
        $result = @{ Shares=@(); OpenSessions=0; Partial=$false }
        try {
            # Try Get-SmbShare (PS4+) first
            if ($PSMaj -ge 4) {
                try {
                    $allShares = @(Get-SmbShare -ErrorAction Stop | Where-Object { $_.Name -notmatch '^\w\$$|^ADMIN\$$|^IPC\$$|^PRINT\$$' })
                    Write-Host ("  [Shares] Found {0} non-admin share(s); collecting ACLs..." -f $allShares.Count) -ForegroundColor Gray
                    $collected = New-Object System.Collections.ArrayList
                    $idx = 0
                    foreach ($s in $allShares) {
                        $idx++
                        # Emit per-share progress so the UI shows which share we are on
                        # (Get-SmbShareAccess can hang for minutes on stale DFS or offline backends)
                        Write-Host ("  [Shares] ({0}/{1}) {2} -> {3}" -f $idx, $allShares.Count, $s.Name, $s.Path) -ForegroundColor DarkGray
                        try {
                            $acl = Get-SmbShareAccess -Name $s.Name -ErrorAction SilentlyContinue

                            # Enumerate folder size + file count (30s timeout per share via background job)
                            $sizeGB = $null; $fileCount = $null; $subfolders = @()
                            if ($s.Path -and (Test-Path $s.Path -ErrorAction SilentlyContinue)) {
                                try {
                                    $sizeJob = Start-Job -ScriptBlock {
                                        param($p)
                                        $items = Get-ChildItem $p -Recurse -Force -ErrorAction SilentlyContinue
                                        $bytes = ($items | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                                        $count = ($items | Where-Object {-not $_.PSIsContainer} | Measure-Object).Count
                                        # Top-level subfolder sizes (one level deep only)
                                        $subs = Get-ChildItem $p -Directory -Force -ErrorAction SilentlyContinue | ForEach-Object {
                                            $sb = (Get-ChildItem $_.FullName -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                                            @{ Name=$_.Name; SizeGB=[math]::Round(($sb/1GB),2) }
                                        }
                                        @{ Bytes=$bytes; Count=$count; Subs=$subs }
                                    } -ArgumentList $s.Path
                                    $done = Wait-Job $sizeJob -Timeout 30
                                    if ($done) {
                                        $r = Receive-Job $sizeJob -ErrorAction SilentlyContinue
                                        if ($r) {
                                            $sizeGB    = [math]::Round(($r.Bytes / 1GB), 2)
                                            $fileCount = $r.Count
                                            $subfolders = @($r.Subs)
                                        }
                                    } else {
                                        Write-Host ("  [Shares] ({0}/{1}) {2} size timed out" -f $idx, $allShares.Count, $s.Name) -ForegroundColor DarkYellow
                                    }
                                    Remove-Job $sizeJob -Force -ErrorAction SilentlyContinue
                                } catch { }
                            }

                            [void]$collected.Add(@{
                                Name        = $s.Name
                                Path        = $s.Path
                                Description = $s.Description
                                Permissions = @($acl | Select-Object AccountName, AccessControlType, AccessRight)
                                SizeGB      = $sizeGB
                                FileCount   = $fileCount
                                Subfolders  = $subfolders
                            })
                        } catch {
                            Write-Host ("  [Shares] ({0}/{1}) {2} ACL failed: {3}" -f $idx, $allShares.Count, $s.Name, $_.Exception.Message) -ForegroundColor DarkYellow
                            [void]$collected.Add(@{ Name=$s.Name; Path=$s.Path; Description=$s.Description })
                        }
                    }
                    $result.Shares = @($collected)
                    try {
                        $sessions = Get-SmbSession -ErrorAction SilentlyContinue
                        $result.OpenSessions = if ($sessions) { @($sessions).Count } else { 0 }
                    } catch { }
                    Write-Host ("  [Shares] Done - {0} share(s) collected" -f $result.Shares.Count) -ForegroundColor Gray
                    return $result
                } catch { }
            }
            # Fix 8 -- WMI Win32_Share fallback (PS 3 or when Get-SmbShare is unavailable)
            # Maps Name, Path, Description, Type to match the Get-SmbShare output structure
            $wmiShares = Safe-Wmi Win32_Share FileShares
            if ($wmiShares) {
                $result.Shares = @($wmiShares |
                    Where-Object { $_.Type -eq 0 -and $_.Name -notmatch '^\w\$$|^IPC\$$|^ADMIN\$$|^PRINT\$$' } |
                    ForEach-Object {
                        @{
                            Name        = [string]$_.Name
                            Path        = [string]$_.Path
                            Description = [string]$_.Description
                            Type        = [int]$_.Type
                            Permissions = @()   # Win32_Share has no ACL method -- left empty for fallback
                        }
                    })
                $result.Partial = $true
            }
        } catch {
            cb-Log "FileShares" "Outer error: $_"
            $result.Partial = $true
        }
        return $result
    }

    # -- NPS / RADIUS ----------------------------------------------------------

    function Collect-NPSDetails {
        Write-Host "  [NPS] Querying NPS/RADIUS configuration..." -ForegroundColor Gray
        $result = @{ Installed=$false; Clients=@(); Policies=@(); Partial=$false }
        try {
            $npsSvc = Get-Service -Name IAS -ErrorAction SilentlyContinue  # NPS = IAS service
            if (-not $npsSvc -or $npsSvc.Status -ne 'Running') { return $result }
            $result.Installed = $true
            # Try NPS module (2012+)
            try {
                if (Get-Module -ListAvailable -Name NPS -ErrorAction SilentlyContinue) {
                    Import-Module NPS -ErrorAction Stop
                    $result.Clients  = @(Get-NpsRadiusClient -ErrorAction Stop | Select-Object Name, Address, Enabled)
                    $result.Policies = @(Get-NpsNetworkPolicy -ErrorAction Stop | Select-Object Name, Enabled, ProcessingOrder)
                    return $result
                }
            } catch { }
            # Fix 5 -- robust netsh nps fallback parsing
            # Parses 'Name = ' and 'Address = ' fields from netsh output into structured objects
            try {
                $npsClientOut = netsh nps show client 2>&1
                $npsNpOut     = netsh nps show np 2>&1

                # Parse clients: each block has Name = ... and Address = ... lines
                $parsedClients = [System.Collections.ArrayList]@()
                $curName = ''; $curAddr = ''
                foreach ($ln in $npsClientOut) {
                    if ($ln -match '^\s*Name\s*=\s*(.+)$')    { $curName = $matches[1].Trim() }
                    if ($ln -match '^\s*Address\s*=\s*(.+)$') { $curAddr = $matches[1].Trim() }
                    if ($curName -and $curAddr) {
                        [void]$parsedClients.Add([ordered]@{ Name=$curName; Address=$curAddr })
                        $curName = ''; $curAddr = ''
                    }
                }
                $result.Clients = @($parsedClients)

                # Parse network policies: Name = ... lines
                $parsedPolicies = [System.Collections.ArrayList]@()
                foreach ($ln in $npsNpOut) {
                    if ($ln -match '^\s*Name\s*=\s*(.+)$') {
                        [void]$parsedPolicies.Add([ordered]@{ Name=$matches[1].Trim(); Address='' })
                    }
                }
                $result.Policies = @($parsedPolicies)

                $result.Partial = $true
                cb-Log "NPS" "Used netsh fallback - $($parsedClients.Count) client(s), $($parsedPolicies.Count) polic(ies) found"
            } catch { cb-Log "NPS" "netsh nps fallback failed: $_" }
        } catch {
            cb-Log "NPS" "Outer error: $_"
            $result.Partial = $true
        }
        return $result
    }

    # -- IIS -------------------------------------------------------------------

    function Collect-IISDetails {
        Write-Host "  [IIS] Querying IIS sites and app pools..." -ForegroundColor Gray
        $result = @{ Installed=$false; Sites=@(); AppPools=@(); Partial=$false }
        try {
            $w3svc = Get-Service -Name W3SVC -ErrorAction SilentlyContinue
            if (-not $w3svc -or $w3svc.Status -ne 'Running') { return $result }
            $result.Installed = $true
            # Try WebAdministration module
            try {
                Import-Module WebAdministration -ErrorAction Stop
                $result.Sites = @(Get-Website -ErrorAction Stop | Select-Object Name, State, PhysicalPath,
                    @{N='Bindings';E={ ($_.Bindings.Collection | ForEach-Object { $_.BindingInformation }) -join '; ' }})
                $result.AppPools = @(Get-WebConfiguration machine/webroot/apphost/*/[system.applicationHost/applicationPools/add] -ErrorAction SilentlyContinue |
                    Select-Object Name, State, ManagedRuntimeVersion, ManagedPipelineMode)
                return $result
            } catch { cb-Log "IIS" "WebAdministration module unavailable: $_" }
            # Fallback: appcmd.exe
            try {
                $appcmd = "$env:SystemRoot\System32\inetsrv\appcmd.exe"
                if (Test-Path $appcmd) {
                    $sitesRaw  = & $appcmd list site 2>&1
                    $result.Sites = @($sitesRaw | ForEach-Object { @{ Raw=$_ } })
                    $result.Partial = $true
                }
            } catch { }
        } catch {
            cb-Log "IIS" "Outer error: $_"
            $result.Partial = $true
        }
        return $result
    }

    # -- SQL SERVER ------------------------------------------------------------

    function Collect-SQLDetails {
        Write-Host "  [SQL] Scanning for SQL Server instances..." -ForegroundColor Gray
        $result = @{ Instances=@(); Partial=$false }
        # Find all instances via registry (works for all SQL versions, no module needed)
        $instances = [System.Collections.ArrayList]@()
        $regPaths = @(
            'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Microsoft SQL Server'
        )
        $instanceNames = @()
        foreach ($reg in $regPaths) {
            try {
                $insKey = Get-ItemProperty -Path "$reg" -Name InstalledInstances -ErrorAction SilentlyContinue
                if ($insKey) { $instanceNames += $insKey.InstalledInstances }
            } catch { }
        }
        $instanceNames = @($instanceNames | Select-Object -Unique)
        if ($instanceNames.Count -eq 0) { return $result }

        foreach ($instName in $instanceNames) {
            Write-Host ("    [SQL] Instance: $instName") -ForegroundColor DarkGray
            $inst = @{ InstanceName=$instName; Version='Unknown'; Edition='Unknown'; ProductName='Unknown'; ProductLevel='Unknown'; ProductVersion='Unknown'; EngineEdition=$null; ServiceAccount='Unknown'; EOLStatus='Unknown'; EOLDate='Unknown'; Databases=@(); Logins=@(); Partial=$false }
            try {
                # Get version from registry
                $vKey = if ($instName -eq 'MSSQLSERVER') { 'MSSQL' } else { "MSSQL.$instName" }
                foreach ($reg in $regPaths) {
                    try {
                        # Find the actual versioned key
                        $sqlKey = Get-ChildItem "$reg" -ErrorAction SilentlyContinue |
                            Where-Object { $_.Name -match 'MSSQL\d+' } |
                            ForEach-Object {
                                $kv = Get-ItemProperty -Path "$($_.PSPath)\MSSQLServer\CurrentVersion" -ErrorAction SilentlyContinue
                                if ($kv) { return $kv }
                            } | Select-Object -First 1
                        if ($sqlKey) {
                            $inst.Version = $sqlKey.CurrentVersion
                            break
                        }
                    } catch { }
                }
                # Match version to EOL table
                $verShort = ($inst.Version -split '\.' | Select-Object -First 2) -join '.'
                $verParts  = $inst.Version -split '\.'
                $verKey10  = $verParts[0] + '.' + $verParts[1]
                foreach ($k in $SQLEOLMap.Keys) {
                    if ($inst.Version.StartsWith($k)) {
                        $inst.EOLStatus   = $SQLEOLMap[$k].Status
                        $inst.EOLDate     = $SQLEOLMap[$k].EOL
                        # ProductName = friendly version (e.g. "SQL Server 2017").
                        # Edition gets the SKU (Standard/Enterprise/...) below
                        # if the SQL connection succeeds.
                        $inst.ProductName = $SQLEOLMap[$k].Name
                        $inst.Edition     = $SQLEOLMap[$k].Name  # provisional - real SKU overwrites below
                        if ($inst.EOLStatus -eq 'EOL') {
                            cb-Flag 'critical' "EOL SQL Server: $($SQLEOLMap[$k].Name) ($instName)" "EOL: $($SQLEOLMap[$k].EOL). No security patches. Must upgrade or migrate before production use."
                        }
                        break
                    }
                }
                # Get service account
                try {
                    $svcName = if ($instName -eq 'MSSQLSERVER') { 'MSSQLSERVER' } else { "MSSQL`$$instName" }
                    $svc = Get-WmiObject Win32_Service -Filter "Name='$svcName'" -ErrorAction SilentlyContinue
                    if ($svc) { $inst.ServiceAccount = $svc.StartName }
                } catch { }
                # SQL deep connect - pull database-level details via integrated security
                $inst.Connected = $false
                try {
                    $sqlServer = if ($instName -eq 'MSSQLSERVER') { '.' } else { ".\$instName" }
                    $conn = New-Object System.Data.SqlClient.SqlConnection
                    $conn.ConnectionString = "Server=$sqlServer;Integrated Security=True;Connect Timeout=10;Application Name=M5Discovery"
                    $conn.Open()
                    $inst.Connected = $true

                    # SERVERPROPERTY pull: capture the actual SKU edition and
                    # patch level, not just the friendly product name. The
                    # registry path only gives us the version number; without
                    # this query we can't distinguish Standard/Enterprise/
                    # Developer/Web/Express.
                    try {
                        $cmdSp = $conn.CreateCommand()
                        $cmdSp.CommandText = @"
SELECT
    CAST(SERVERPROPERTY('Edition')        AS nvarchar(256)) AS [Edition],
    CAST(SERVERPROPERTY('EngineEdition')  AS int)           AS [EngineEdition],
    CAST(SERVERPROPERTY('ProductLevel')   AS nvarchar(64))  AS [ProductLevel],
    CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(64))  AS [ProductVersion],
    CAST(SERVERPROPERTY('ProductUpdateLevel') AS nvarchar(64)) AS [ProductUpdateLevel],
    CAST(SERVERPROPERTY('ProductBuild')   AS nvarchar(64))  AS [ProductBuild]
"@
                        $rdrSp = $cmdSp.ExecuteReader()
                        if ($rdrSp.Read()) {
                            $skuEdition = "$($rdrSp['Edition'])"
                            if ($skuEdition) { $inst.Edition = $skuEdition }
                            $eEdition = $rdrSp['EngineEdition']
                            if ($eEdition -isnot [DBNull]) { $inst.EngineEdition = [int]$eEdition }
                            $pLevel = "$($rdrSp['ProductLevel'])"
                            if ($pLevel) { $inst.ProductLevel = $pLevel }
                            $pVer = "$($rdrSp['ProductVersion'])"
                            if ($pVer) { $inst.ProductVersion = $pVer }
                            $pUpd = "$($rdrSp['ProductUpdateLevel'])"
                            if ($pUpd -and $pUpd -ne 'NULL') { $inst.ProductUpdateLevel = $pUpd }
                            $pBld = "$($rdrSp['ProductBuild'])"
                            if ($pBld) { $inst.ProductBuild = $pBld }
                        }
                        $rdrSp.Close()
                    } catch {
                        $inst.Partial = $true
                        Write-Host ("    [SQL]   SERVERPROPERTY query failed: $($_.Exception.Message)") -ForegroundColor DarkYellow
                    }

                    $cmd = $conn.CreateCommand()
                    $cmd.CommandText = @"
SELECT
    d.name,
    d.state_desc,
    d.recovery_model_desc,
    d.compatibility_level,
    CONVERT(varchar, d.create_date, 23) AS create_date,
    ISNULL(SUM(CASE WHEN mf.type = 0 THEN CAST(mf.size AS bigint) * 8 / 1024 ELSE 0 END), 0) AS data_size_mb,
    ISNULL(SUM(CASE WHEN mf.type = 1 THEN CAST(mf.size AS bigint) * 8 / 1024 ELSE 0 END), 0) AS log_size_mb,
    MAX(CONVERT(varchar, b.backup_finish_date, 23)) AS last_full_backup
FROM sys.databases d
LEFT JOIN sys.master_files mf ON d.database_id = mf.database_id
LEFT JOIN msdb.dbo.backupset b ON b.database_name = d.name AND b.type = 'D'
WHERE d.name NOT IN ('master','model','msdb','tempdb','resource','distribution')
GROUP BY d.name, d.state_desc, d.recovery_model_desc, d.compatibility_level, d.create_date
ORDER BY data_size_mb DESC
"@
                    $reader = $cmd.ExecuteReader()
                    $dbs = [System.Collections.ArrayList]@()
                    while ($reader.Read()) {
                        $dbName        = $reader["name"].ToString()
                        $lastFullBackup = if ($reader["last_full_backup"] -is [DBNull] -or $reader["last_full_backup"].ToString() -eq '') { $null } else { $reader["last_full_backup"].ToString() }
                        $dataSizeMB    = if ($reader["data_size_mb"] -is [DBNull]) { 0 } else { [long]$reader["data_size_mb"] }
                        $logSizeMB     = if ($reader["log_size_mb"]  -is [DBNull]) { 0 } else { [long]$reader["log_size_mb"]  }
                        # Lookup app name
                        $appName = ''
                        foreach ($k in $SqlDbAppMap.Keys) {
                            if ($dbName -match $k) { $appName = $SqlDbAppMap[$k]; break }
                        }
                        $db = @{
                            Name           = $dbName
                            State          = $reader["state_desc"].ToString()
                            RecoveryModel  = $reader["recovery_model_desc"].ToString()
                            CompatLevel    = [int]$reader["compatibility_level"]
                            CreateDate     = $reader["create_date"].ToString()
                            DataSizeMB     = $dataSizeMB
                            LogSizeMB      = $logSizeMB
                            LastFullBackup = $lastFullBackup
                            AppGuess       = $appName
                        }
                        # Flag backup health
                        if ($null -eq $lastFullBackup) {
                            cb-Flag 'critical' "SQL DB Never Backed Up: $dbName ($instName)" "No full backup record in msdb. Confirm alternate backup method."
                        } else {
                            try {
                                $daysSince = ((Get-Date) - [datetime]::Parse($lastFullBackup)).TotalDays
                                if ($daysSince -gt 30) {
                                    cb-Flag 'warning' "SQL DB Stale Backup: $dbName" "Last full backup: $lastFullBackup ($([math]::Round($daysSince,0)) days ago)"
                                }
                            } catch { }
                        }
                        [void]$dbs.Add($db)
                    }
                    $reader.Close()

                    # ---- WHO CAN GET IN: server logins + role membership ----
                    # Migration and security both need this: service accounts to
                    # recreate, sysadmins to justify, and any BUILTIN or broad
                    # domain group that effectively opens the instance to
                    # everyone. Read-only, one query, degrades quietly when the
                    # discovery account cannot read server principals.
                    $logins = [System.Collections.ArrayList]@()
                    try {
                        $lcmd = $conn.CreateCommand()
                        $lcmd.CommandTimeout = 30
                        $lcmd.CommandText = @"
SELECT sp.name AS login_name,
       sp.type_desc,
       sp.is_disabled,
       STUFF((SELECT ', ' + r.name
              FROM sys.server_role_members rm
              JOIN sys.server_principals r ON r.principal_id = rm.role_principal_id
              WHERE rm.member_principal_id = sp.principal_id
              FOR XML PATH('')), 1, 2, '') AS server_roles
FROM sys.server_principals sp
WHERE sp.type IN ('S','U','G')
  AND sp.name NOT LIKE '##%'
  AND sp.name NOT LIKE 'NT SERVICE\%'
ORDER BY sp.name
"@
                        $lrdr = $lcmd.ExecuteReader()
                        while ($lrdr.Read()) {
                            $lname  = [string]$lrdr['login_name']
                            $ltype  = [string]$lrdr['type_desc']
                            $ldis   = [bool]$lrdr['is_disabled']
                            $lroles = if ($lrdr['server_roles'] -is [DBNull]) { '' } else { [string]$lrdr['server_roles'] }
                            $isSa   = ($lroles -match 'sysadmin')
                            # A login that effectively grants everyone access.
                            $broad  = ($lname -match '(?i)^BUILTIN\\(Administrators|Users)$' -or
                                       $lname -match '(?i)\\(Domain Users|Authenticated Users|Everyone)$')
                            [void]$logins.Add(@{
                                Name       = $lname
                                Type       = $ltype
                                Disabled   = $ldis
                                Roles      = $lroles
                                IsSysadmin = $isSa
                                IsBroad    = $broad
                            })
                            if ($isSa -and $broad -and -not $ldis) {
                                cb-Flag 'critical' "SQL wide-open sysadmin: $lname ($instName)" "$lname holds sysadmin on $instName. Every member of that group is a full administrator of the SQL instance."
                            }
                        }
                        $lrdr.Close()
                    } catch { cb-Log "SQL-$instName" "Login enumeration failed (needs VIEW ANY DEFINITION): $($_.Exception.Message)" }
                    $inst.Logins = @($logins)

                    # ---- Per-database users. Bounded: a server with hundreds of
                    # databases would otherwise turn this into a long serial walk.
                    $MAX_DB_ACL = 40
                    $dbIdx = 0
                    foreach ($dbrec in $dbs) {
                        if ($dbIdx -ge $MAX_DB_ACL) { $inst.Partial = $true; break }
                        $dbIdx++
                        $dbUsers = [System.Collections.ArrayList]@()
                        try {
                            $ucmd = $conn.CreateCommand()
                            $ucmd.CommandTimeout = 20
                            $safeDb = ($dbrec.Name -replace ']', ']]')
                            $ucmd.CommandText = @"
USE [$safeDb];
SELECT dp.name AS user_name,
       dp.type_desc,
       STUFF((SELECT ', ' + r.name
              FROM sys.database_role_members drm
              JOIN sys.database_principals r ON r.principal_id = drm.role_principal_id
              WHERE drm.member_principal_id = dp.principal_id
              FOR XML PATH('')), 1, 2, '') AS db_roles
FROM sys.database_principals dp
WHERE dp.type IN ('S','U','G')
  AND dp.name NOT IN ('dbo','guest','sys','INFORMATION_SCHEMA','public')
ORDER BY dp.name
"@
                            $urdr = $ucmd.ExecuteReader()
                            while ($urdr.Read()) {
                                $uroles = if ($urdr['db_roles'] -is [DBNull]) { '' } else { [string]$urdr['db_roles'] }
                                [void]$dbUsers.Add(@{
                                    Name  = [string]$urdr['user_name']
                                    Type  = [string]$urdr['type_desc']
                                    Roles = $uroles
                                })
                            }
                            $urdr.Close()
                        } catch { cb-Log "SQL-$instName" "DB user enumeration failed for $($dbrec.Name): $($_.Exception.Message)" }
                        $dbrec['Users'] = @($dbUsers)
                    }

                    $conn.Close()
                    $inst.Databases = $dbs
                } catch {
                    # Fix 6 -- capture connect error detail so report can show why it failed
                    $connErrMsg = $_.Exception.Message
                    if ($connErrMsg.Length -gt 200) { $connErrMsg = $connErrMsg.Substring(0, 200) }
                    cb-Log "SQL-$instName" "Deep connect failed: $_"
                    $inst.Connected     = $false
                    $inst.ConnectError  = $connErrMsg
                    $inst.Databases     = @()
                    $inst.Partial       = $true
                }
            } catch {
                cb-Log "SQL-$instName" "Instance collection error: $_"
                $inst.Partial = $true
            }
            [void]$instances.Add($inst)
        }
        $result.Instances = $instances
        return $result
    }

    # -- EXCHANGE --------------------------------------------------------------

    function Collect-ExchangeDetails {
        Write-Host "  [Exchange] Checking for Exchange Server..." -ForegroundColor Gray
        $result = @{ Installed=$false; Version=''; VersionName=''; EOLStatus=''; EOLDate=''; MailboxCount=0; DatabaseSizes=@(); TransportServiceRunning=$false; Partial=$false }
        try {
            # Check registry for Exchange install
            $exchKey = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup' -ErrorAction SilentlyContinue
            if (-not $exchKey) { $exchKey = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup' -ErrorAction SilentlyContinue }
            if (-not $exchKey) { $exchKey = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Exchange\Setup' -ErrorAction SilentlyContinue }
            if (-not $exchKey) { return $result }
            $result.Installed = $true
            $result.Version   = $exchKey.MsiProductMajor
            # Check transport service
            $transpSvc = Get-Service -Name MSExchangeTransport -ErrorAction SilentlyContinue
            $result.TransportServiceRunning = ($transpSvc -and $transpSvc.Status -eq 'Running')
            if (-not $result.TransportServiceRunning) {
                cb-Flag 'warning' 'Exchange Transport Service Not Running' 'MSExchangeTransport is stopped - mail flow may be interrupted or Exchange is decommissioned.'
            }
            # EOL lookup
            $majVer = $result.Version.ToString()
            if     ($majVer -eq '15') {
                # Need minor to distinguish 2013/2016/2019
                try {
                    $minKey = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup' -ErrorAction Stop
                    $full   = "$($minKey.MsiProductMajor).$($minKey.MsiProductMinor)"
                    if     ($full -match '^15\.2') { $eolKey = '15.2' }
                    elseif ($full -match '^15\.1') { $eolKey = '15.1' }
                    else                           { $eolKey = '15.0' }
                } catch { $eolKey = '15.0' }
            } elseif ($majVer -eq '14') { $eolKey = '14'  }
            elseif ($majVer -eq '8')  { $eolKey = '8'   }
            elseif ($majVer -eq '6')  { $eolKey = '6'   }
            else                       { $eolKey = $majVer }
            if ($ExchEOLMap.ContainsKey($eolKey)) {
                $eol = $ExchEOLMap[$eolKey]
                $result.VersionName = $eol.Name
                $result.EOLStatus   = $eol.Status
                $result.EOLDate     = $eol.EOL
                if ($eol.Status -eq 'EOL') {
                    cb-Flag 'critical' "EOL Exchange Server: $($eol.Name)" "EOL: $($eol.EOL). No security patches. Migrate mailboxes to Exchange Online or newer version immediately."
                } elseif ($eol.Status -eq 'Near EOL') {
                    cb-Flag 'warning' "Near-EOL Exchange: $($eol.Name)" "EOL: $($eol.EOL). Plan migration to Exchange Online."
                }
            }
            # Try Exchange management shell for mailbox count
            try {
                $exchBin = $exchKey.MsiInstallPath + '\bin\RemoteExchange.ps1'
                if (Test-Path $exchBin) {
                    # Load just enough to query
                    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop 2>$null
                    $result.MailboxCount = (Get-Mailbox -ResultSize Unlimited -ErrorAction Stop).Count
                    $dbs = Get-MailboxDatabase -Status -ErrorAction SilentlyContinue
                    $result.DatabaseSizes = @($dbs | Select-Object Name,
                        @{N='SizeGB';E={ if($_.DatabaseSize) { [math]::Round($_.DatabaseSize.ToBytes()/1GB,2) } else { '?' } }},
                        @{N='EdbFilePath';E={$_.EdbFilePath}})
                }
            } catch { $result.Partial = $true; cb-Log "Exchange" "EMS query failed - partial data: $_" }
        } catch {
            cb-Log "Exchange" "Outer error: $_"
            $result.Partial = $true
        }
        return $result
    }

    # -- HYPER-V ---------------------------------------------------------------

    function Collect-HyperVDetails {
        Write-Host "  [Hyper-V] Querying virtual machines..." -ForegroundColor Gray
        $result = @{ Installed=$false; VMs=@(); Partial=$false }
        try {
            $hvSvc = Get-Service -Name vmms -ErrorAction SilentlyContinue
            if (-not $hvSvc -or $hvSvc.Status -ne 'Running') { return $result }
            $result.Installed = $true
            # Fix 3 -- Hyper-V WMI namespace fallback (root\virtualization\v2 = Server 2012+;
            # root\virtualization = Server 2008 R2 / early HV; try v2 first, fall back gracefully)
            $hvNS = 'root\virtualization\v2'
            try {
                Get-WmiObject -Namespace 'root\virtualization\v2' -Class Msvm_ComputerSystem -ErrorAction Stop | Out-Null
            } catch {
                try {
                    Get-WmiObject -Namespace 'root\virtualization' -Class Msvm_ComputerSystem -ErrorAction Stop | Out-Null
                    $hvNS = 'root\virtualization'
                    cb-Log "HyperV" "root\virtualization\v2 not present - using legacy root\virtualization namespace"
                } catch {
                    cb-Log "HyperV" "No Hyper-V WMI namespace found - HV may not be fully installed"
                }
            }
            try {
                if (Get-Module -ListAvailable -Name Hyper-V -ErrorAction SilentlyContinue) {
                    Import-Module Hyper-V -ErrorAction Stop
                    $vms = Get-VM -ErrorAction Stop
                    $result.VMs = @($vms | ForEach-Object {
                        @{
                            Name          = $_.Name
                            State         = $_.State.ToString()
                            Generation    = $_.Generation
                            MemoryGB      = [math]::Round($_.MemoryAssigned / 1GB, 2)
                            CPUCount      = $_.ProcessorCount
                            Uptime        = $_.Uptime.ToString()
                            Checkpoints   = ($_ | Get-VMCheckpoint -ErrorAction SilentlyContinue | Measure-Object).Count
                        }
                    })
                    # Virtual switches
                    try {
                        $result.VirtualSwitches = @(Get-VMSwitch -ErrorAction Stop | ForEach-Object {
                            @{ Name=$_.Name; SwitchType=$_.SwitchType.ToString(); NetAdapterName=[string]$_.NetAdapterInterfaceDescription }
                        })
                    } catch { }
                    # Default VM/VHD store paths
                    try {
                        $hvSettings = Get-VMHost -ErrorAction Stop
                        $result.DefaultVMPath  = [string]$hvSettings.VirtualMachinePath
                        $result.DefaultVHDPath = [string]$hvSettings.VirtualHardDiskPath
                        $result.LiveMigrationEnabled = [bool]$hvSettings.VirtualMachineMigrationEnabled
                        $result.NumaSpanningEnabled  = [bool]$hvSettings.NumaSpanningEnabled
                    } catch { }
                }
            } catch { $result.Partial = $true; cb-Log "HyperV" "Get-VM failed: $_" }
        } catch {
            cb-Log "HyperV" "Outer error: $_"
            $result.Partial = $true
        }

        # -- HYPER-V PERFORMANCE CHECK (PerfMon .blg log files) ----------------
        # Look for Performance Monitor Data Collector Set log files with 30+ days
        # of history. If found, pull per-VM CPU% and memory counters from the log.
        $result.HyperVPerfAvailable = $false
        $result.HyperVPerfNote      = ''
        $result.HyperVPerf          = $null

        try {
            $blgFiles = [System.Collections.ArrayList]@()
            $perfLogRoots = @('C:\PerfLogs\Admin', 'C:\PerfLogs')
            foreach ($root in $perfLogRoots) {
                if (Test-Path $root) {
                    try {
                        $found = Get-ChildItem -Path $root -Filter '*.blg' -Recurse -ErrorAction SilentlyContinue
                        if ($found) { foreach ($f in $found) { [void]$blgFiles.Add($f) } }
                    } catch { }
                }
            }

            # Find the best candidate: longest span that is >= 30 days
            $bestFile  = $null
            $bestSpan  = 0
            foreach ($blg in $blgFiles) {
                try {
                    $spanDays = ($blg.LastWriteTime - $blg.CreationTime).TotalDays
                    if ($spanDays -ge 30 -and $spanDays -gt $bestSpan) {
                        $bestSpan = $spanDays
                        $bestFile = $blg
                    }
                } catch { }
            }

            if ($null -eq $bestFile) {
                $result.HyperVPerfAvailable = $false
                $result.HyperVPerfNote      = 'No 30-day PerfMon history found - performance data unavailable'
            } else {
                # Attempt to read Hyper-V counters from the log file
                try {
                    $blgPath   = $bestFile.FullName
                    $startDate = $bestFile.CreationTime.ToString('yyyy-MM-dd')
                    $endDate   = $bestFile.LastWriteTime.ToString('yyyy-MM-dd')
                    $spanDays  = [math]::Round($bestSpan, 1)

                    $cpuCounter = '\Hyper-V Hypervisor Virtual Processor(*)\% Guest Run Time'
                    $ramCounter = '\Hyper-V Dynamic Memory VM(*)\Physical Memory'

                    # Import-Counter reads from a .blg binary log file
                    $cpuData = Import-Counter -Path $blgPath -Counter $cpuCounter -MaxSamples 9999999 -ErrorAction Stop |
                               Select-Object -ExpandProperty CounterSamples

                    # Group by instance (VM name) and calculate avg/max
                    $vmStats = @{}
                    foreach ($sample in $cpuData) {
                        # Counter instance is like "vm name:0" - strip the colon-suffix
                        $instLabel = ($sample.InstanceName -replace ':\d+$', '').Trim()
                        if ($instLabel -eq '_total' -or $instLabel -eq '') { continue }
                        if (-not $vmStats.ContainsKey($instLabel)) {
                            $vmStats[$instLabel] = @{ CpuSamples=[System.Collections.ArrayList]@(); RamSamples=[System.Collections.ArrayList]@() }
                        }
                        [void]$vmStats[$instLabel].CpuSamples.Add([double]$sample.CookedValue)
                    }

                    # Pull RAM samples - wrap in try so missing counter doesn't abort
                    try {
                        $ramData = Import-Counter -Path $blgPath -Counter $ramCounter -MaxSamples 9999999 -ErrorAction Stop |
                                   Select-Object -ExpandProperty CounterSamples
                        foreach ($sample in $ramData) {
                            $instLabel = ($sample.InstanceName -replace ':\d+$', '').Trim()
                            if ($instLabel -eq '_total' -or $instLabel -eq '') { continue }
                            if (-not $vmStats.ContainsKey($instLabel)) {
                                $vmStats[$instLabel] = @{ CpuSamples=[System.Collections.ArrayList]@(); RamSamples=[System.Collections.ArrayList]@() }
                            }
                            [void]$vmStats[$instLabel].RamSamples.Add([double]$sample.CookedValue)
                        }
                    } catch { }

                    # Build per-VM result list
                    $vmPerfList = [System.Collections.ArrayList]@()
                    foreach ($vmName in $vmStats.Keys) {
                        $cpu  = $vmStats[$vmName].CpuSamples
                        $ram  = $vmStats[$vmName].RamSamples
                        $cpuAvg = if ($cpu.Count -gt 0) { [math]::Round(($cpu | Measure-Object -Average).Average, 1) } else { $null }
                        $cpuMax = if ($cpu.Count -gt 0) { [math]::Round(($cpu | Measure-Object -Maximum).Maximum, 1) } else { $null }
                        $ramAvg = if ($ram.Count -gt 0) { [math]::Round(($ram | Measure-Object -Average).Average, 0) } else { $null }
                        $ramMax = if ($ram.Count -gt 0) { [math]::Round(($ram | Measure-Object -Maximum).Maximum, 0) } else { $null }
                        [void]$vmPerfList.Add(@{
                            Name        = $vmName
                            CPU_Avg_Pct = $cpuAvg
                            CPU_Max_Pct = $cpuMax
                            RAM_Avg_MB  = $ramAvg
                            RAM_Max_MB  = $ramMax
                        })
                    }

                    $result.HyperVPerfAvailable = $true
                    $result.HyperVPerf = @{
                        Available  = $true
                        SpanDays   = $spanDays
                        LogPath    = $blgPath
                        StartDate  = $startDate
                        EndDate    = $endDate
                        VMs        = $vmPerfList
                    }
                } catch {
                    # Get-Counter failed on this log (counters missing, format issue, etc.)
                    $result.HyperVPerfAvailable = $false
                    $result.HyperVPerf = @{
                        Available = $false
                        LogPath   = $bestFile.FullName
                        SpanDays  = [math]::Round($bestSpan, 1)
                        Note      = "Get-Counter failed reading log: $_"
                        VMs       = @()
                    }
                    cb-Log "HyperV-Perf" "Get-Counter failed on $($bestFile.FullName): $_"
                }
            }
        } catch {
            # Entire perf check failed - non-fatal
            $result.HyperVPerfAvailable = $false
            $result.HyperVPerfNote      = "PerfMon check error: $_"
            cb-Log "HyperV-Perf" "Outer perf check error: $_"
        }

        return $result
    }

    # -- INSTALLED APPLICATIONS ------------------------------------------------

    function Collect-InstalledApps {
        Write-Host "  [Apps] Cataloging installed software..." -ForegroundColor Gray
        $apps = [System.Collections.ArrayList]@()
        $regPaths = @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
        )
        foreach ($path in $regPaths) {
            try {
                $entries = Get-ItemProperty $path -ErrorAction SilentlyContinue |
                    Where-Object { $_.DisplayName -and $_.DisplayName.Trim() -ne '' }
                foreach ($e in $entries) {
                    try {
                        $nameLC = $e.DisplayName.ToLower()
                        # Apply flag rules
                        $flagMatch = $null
                        foreach ($rule in $AppFlagRules) {
                            $matched = $false
                            foreach ($kw in $rule.K) { if ($nameLC -match [regex]::Escape($kw)) { $matched = $true; break } }
                            if ($matched) { $flagMatch = $rule; break }
                        }
                        if ($flagMatch) {
                            cb-Flag $flagMatch.Sev "$($flagMatch.Cat): $($e.DisplayName)" $flagMatch.Note
                        }
                        [void]$apps.Add(@{
                            Name         = $e.DisplayName
                            Version      = $e.DisplayVersion
                            InstallDate  = $e.InstallDate
                            Publisher    = $e.Publisher
                            Category     = if ($flagMatch) { $flagMatch.Cat } else { 'Other' }
                            FlagSeverity = if ($flagMatch) { $flagMatch.Sev } else { 'none' }
                        })
                    } catch { }
                }
            } catch { cb-Log "Apps" "Registry path failed ($path): $_" }
        }
        # Deduplicate by name
        $seen  = @{}
        $dedup = [System.Collections.ArrayList]@()
        foreach ($a in $apps) {
            $key = $a.Name.ToLower()
            if (-not $seen.ContainsKey($key)) {
                $seen[$key] = $true
                [void]$dedup.Add($a)
            }
        }
        return ,$dedup
    }

    # -- SCHEDULED TASKS -------------------------------------------------------

    function Collect-ScheduledTasks {
        Write-Host "  [Tasks] Enumerating non-Microsoft scheduled tasks..." -ForegroundColor Gray
        $tasks = [System.Collections.ArrayList]@()
        try {
            if ($PSMaj -ge 4) {
                $allTasks = Get-ScheduledTask -ErrorAction Stop | Where-Object {
                    $_.TaskPath -notmatch '\\Microsoft\\' -and $_.State -ne 'Disabled'
                }
                foreach ($t in $allTasks) {
                    try {
                        $info = Get-ScheduledTaskInfo -TaskName $t.TaskName -TaskPath $t.TaskPath -ErrorAction SilentlyContinue
                        [void]$tasks.Add(@{
                            Name        = $t.TaskName
                            Path        = $t.TaskPath
                            State       = $t.State.ToString()
                            LastRun     = if ($info) { $info.LastRunTime.ToString("yyyy-MM-dd HH:mm:ss") } else { 'Unknown' }
                            LastResult  = if ($info) { "0x{0:X}" -f $info.LastTaskResult } else { 'Unknown' }
                            NextRun     = if ($info -and $info.NextRunTime -gt (Get-Date)) { $info.NextRunTime.ToString("yyyy-MM-dd HH:mm:ss") } else { 'Not scheduled' }
                            Action      = ($t.Actions | ForEach-Object { $_.Execute + ' ' + $_.Arguments }) -join '; '
                        })
                    } catch { }
                }
            } else {
                # Fallback: schtasks.exe
                $schtasks = schtasks /query /fo CSV /nh 2>&1 | ConvertFrom-Csv -Header 'Name','NextRun','Status' -ErrorAction SilentlyContinue
                if ($schtasks) {
                    foreach ($t in ($schtasks | Where-Object { $_.Status -ne 'Disabled' })) {
                        [void]$tasks.Add(@{ Name=$t.Name; NextRun=$t.NextRun; State=$t.Status })
                    }
                }
            }
        } catch { cb-Log "Tasks" "Outer error: $_" }
        return ,$tasks
    }

    # -- SERVICES --------------------------------------------------------------

    function Collect-Services {
        Write-Host "  [Services] Enumerating non-standard running services..." -ForegroundColor Gray
        $services = [System.Collections.ArrayList]@()
        $msPublishers = @('Microsoft','Windows','NT AUTHORITY')
        try {
            $wmiSvcs = Safe-Wmi Win32_Service Services
            if ($wmiSvcs) {
                foreach ($s in ($wmiSvcs | Where-Object { $_.State -eq 'Running' })) {
                    try {
                        $isMS = $false
                        foreach ($pub in $msPublishers) { if ($s.PathName -match $pub -or $s.Description -match $pub) { $isMS = $true; break } }
                        if ($s.PathName -match 'system32|syswow64') { $isMS = $true }
                        [void]$services.Add(@{
                            Name        = $s.Name
                            DisplayName = $s.DisplayName
                            State       = $s.State
                            StartMode   = $s.StartMode
                            StartName   = $s.StartName
                            Path        = $s.PathName
                            Description = $s.Description
                            IsMS        = $isMS
                        })
                    } catch { }
                }
            }
        } catch { cb-Log "Services" "Outer error: $_" }
        return ,$services
    }

    # -- EVENT LOG SUMMARY -----------------------------------------------------

    function Collect-EventLogSummary {
        Write-Host "  [EventLog] Reading last 24h critical/error events..." -ForegroundColor Gray
        $result = @{ CriticalCount=0; ErrorCount=0; TopSources=@(); RecentCritical=@(); Partial=$false }
        try {
            $cutoff  = (Get-Date).AddHours(-24)
            $logs    = @('System','Application')
            $allEvts = [System.Collections.ArrayList]@()
            # Fix 11 -- Get-EventLog wrapped in Start-Job/Wait-Job with 15s timeout
            # Prevents the call from hanging indefinitely on large or corrupt event logs
            foreach ($log in $logs) {
                try {
                    $evtJob = Start-Job -ScriptBlock {
                        param($logName, $after)
                        Get-EventLog -LogName $logName -EntryType Error,Warning,Information -After $after -Newest 200 -ErrorAction SilentlyContinue |
                            Select-Object TimeGenerated, EntryType, Source,
                                @{ N='Message'; E={ ($_.Message -replace '\r?\n',' ').Substring(0, [Math]::Min(200, ($_.Message -replace '\r?\n',' ').Length)) } }
                    } -ArgumentList $log, $cutoff
                    $evtDone = Wait-Job $evtJob -Timeout 15
                    if ($evtDone) {
                        $evts = Receive-Job $evtJob -ErrorAction SilentlyContinue
                        if ($evts) { foreach ($e in $evts) { [void]$allEvts.Add($e) } }
                    } else {
                        cb-Log "EventLog" "Get-EventLog timed out on $log log (>15s) - partial results"
                        $result.Partial = $true
                    }
                    Remove-Job $evtJob -Force -ErrorAction SilentlyContinue
                } catch { cb-Log "EventLog" "Job error on $log`: $_" }
            }
            $result.CriticalCount = @($allEvts | Where-Object { $_.EntryType -eq 'Error' }).Count
            $result.TopSources    = @($allEvts | Group-Object Source | Sort-Object Count -Descending | Select-Object -First 10 |
                ForEach-Object { @{ Source=$_.Name; Count=$_.Count } })
            $result.RecentCritical = @($allEvts | Where-Object { $_.EntryType -eq 'Error' } |
                Sort-Object TimeGenerated -Descending | Select-Object -First 10 |
                ForEach-Object {
                    # TimeGenerated may be a DateTime (live) or already a string (deserialized from job)
                    $timeStr = try { if ($_.TimeGenerated -is [datetime]) { $_.TimeGenerated.ToString("yyyy-MM-dd HH:mm") } else { [string]$_.TimeGenerated } } catch { '' }
                    # Message is pre-truncated to 200 chars in the job Select-Object; just use it as-is
                    $msgStr  = [string]$_.Message
                    @{ Time=$timeStr; Source=$_.Source; Message=$msgStr }
                })
        } catch {
            cb-Log "EventLog" "Outer error: $_"
            $result.Partial = $true
        }
        return $result
    }

    # -- PRINTERS --------------------------------------------------------------

    function Collect-Printers {
        Write-Host "  [Printers] Looking for print services..." -ForegroundColor Gray
        $printers = [System.Collections.ArrayList]@()
        try {
            $wmiPrinters = Safe-Wmi Win32_Printer Printers
            if ($wmiPrinters) {
                foreach ($p in ($wmiPrinters | Where-Object { $_.Name -notmatch 'Fax|PDF|XPS|OneNote|Microsoft' })) {
                    [void]$printers.Add(@{
                        Name       = $p.Name
                        PortName   = $p.PortName
                        DriverName = $p.DriverName
                        Shared     = $p.Shared
                        ShareName  = $p.ShareName
                        Status     = $p.Status
                    })
                }
            }
        } catch { cb-Log "Printers" "Outer error: $_" }
        return ,$printers
    }

    # -- PRINT SERVER (shared queues + access ACL + GPP deploy cross-ref) -------

    # Scan SYSVOL for Group Policy Preference printer deployments. Returns a
    # map of printer sharename/path/name -> deploying GPO display name. Any
    # domain member can read \\<domain>\SYSVOL, so this works off a DC too.
    function Get-GPPrinterDeployments {
        $map = @{}
        try {
            $dom = $env:USERDNSDOMAIN
            if (-not $dom) { return $map }
            $polRoot = "\\$dom\SYSVOL\$dom\Policies"
            if (-not (Test-Path $polRoot -ErrorAction SilentlyContinue)) { return $map }
            # GUID -> display name (best effort; falls back to the raw GUID)
            $nameCache = @{}
            try {
                if (Get-Command Get-GPO -ErrorAction SilentlyContinue) {
                    foreach ($g in (Get-GPO -All -ErrorAction Stop)) {
                        $nameCache[("{$($g.Id)}").ToUpper()] = [string]$g.DisplayName
                    }
                }
            } catch { }
            # A recursive scan of \\<domain>\SYSVOL is a NETWORK walk that can run
            # for many minutes on a large or slowly-replicating SYSVOL (WAN-linked
            # DC, years of accumulated GPOs). Unbounded, it hangs the entire remote
            # discovery job and every server queued behind it. Time-box it the same
            # way the file-share size scan and event-log read are boxed.
            $xmlFiles = @()
            try {
                $sysvolJob = Start-Job -ScriptBlock {
                    param($root)
                    Get-ChildItem -Path $root -Recurse -Filter 'Printers.xml' -ErrorAction SilentlyContinue |
                        Select-Object -ExpandProperty FullName
                } -ArgumentList $polRoot
                if (Wait-Job $sysvolJob -Timeout 45) {
                    $paths = Receive-Job $sysvolJob -ErrorAction SilentlyContinue
                    if ($paths) { $xmlFiles = @($paths | ForEach-Object { [PSCustomObject]@{ FullName = $_ } }) }
                } else {
                    Stop-Job $sysvolJob -ErrorAction SilentlyContinue
                    cb-Log "PrintServer" "SYSVOL GPP scan timed out after 45s - GPO deployment cross-reference skipped"
                }
                Remove-Job $sysvolJob -Force -ErrorAction SilentlyContinue
            } catch { cb-Log "PrintServer" "SYSVOL scan job failed: $_" }

            foreach ($xf in $xmlFiles) {
                $guid = ''
                if ($xf.FullName -match '\\Policies\\(\{[0-9A-Fa-f-]+\})\\') { $guid = $Matches[1].ToUpper() }
                $gpoName = if ($guid -and $nameCache.ContainsKey($guid)) { $nameCache[$guid] } elseif ($guid) { $guid } else { '(unknown GPO)' }
                try {
                    [xml]$x = Get-Content -Path $xf.FullName -ErrorAction Stop
                    foreach ($node in $x.SelectNodes("//*[local-name()='SharedPrinter' or local-name()='PortPrinter' or local-name()='LocalPrinter']")) {
                        $props = $node.SelectSingleNode("*[local-name()='Properties']")
                        $pPath = if ($props) { [string]$props.getAttribute('path') } else { '' }
                        $pName = [string]$node.getAttribute('name')
                        foreach ($key in @($pPath, $pName)) {
                            if ($key) { $map[$key.ToLower().TrimStart('\')] = $gpoName }
                        }
                    }
                } catch { }
            }
        } catch { cb-Log "PrintServer" "GPP scan error: $_" }
        return $map
    }

    # Classify a printer's access from its security descriptor.
    # Returns @{ Access='open'|'restricted'|'unknown'; Groups=@(named trustees) }.
    # "open"  = Everyone / Authenticated Users can print.
    # "restricted" = only named security group(s) hold Print rights.
    function Get-PrinterAccessSummary {
        <#
            Work out WHO can print to a queue, and with what rights.

            Returns:
              Access      'open' | 'restricted' | 'unknown'
              Groups      flat list of named (non-well-known) trustees
              Permissions [ @{ Trustee; Rights; Kind; IsOpen } ] - the full ACL,
                          used by the report to render a per-printer access list.

            "Open" means a well-known everyone-style trustee (Everyone,
            Authenticated Users, Domain Users) holds Print. In that case the
            report says "Everyone" rather than enumerating individual accounts,
            because every domain user inherits access and listing them is noise.
            When only named groups/users hold Print, each one is listed - those
            are what has to be recreated in Entra / Universal Print.
        #>
        param($Printer, [bool]$UseGetPrinter)

        $out = @{ Access='unknown'; Groups=@(); Permissions=@() }
        $openTrustees   = @('everyone','authenticated users','domain users','all application packages')
        $adminTrustees  = @('administrators','system','creator owner','print operators',
                            'server operators','nt service\trustedinstaller','trustedinstaller',
                            'nt authority\system','builtin\administrators','power users')
        $named = [System.Collections.ArrayList]@()
        $perms = [System.Collections.ArrayList]@()
        $sawOpen = $false

        # Translate a printer access mask / SDDL rights string into the three
        # permissions the Windows printer security UI actually shows.
        $rightsLabel = {
            param([string]$r)
            if (-not $r) { return 'Print' }
            $t = $r.ToLower()
            $isManage = ($t -match 'writedacl|writeowner|genericall|fullcontrol|changepermissions|takeownership|sd|wd|wo')
            $isDocs   = ($t -match 'genericwrite|delete|createchild|deletechild|dc|cc')
            if ($isManage) { return 'Manage this printer' }
            if ($isDocs)   { return 'Manage documents' }
            return 'Print'
        }

        # Record one ACE, bucketing the trustee so the report can decide what to
        # show without re-deriving any of this.
        $addPerm = {
            param([string]$account, [string]$rights)
            if (-not $account) { return }
            $short = ($account -replace '^.*\\', '').Trim()
            if (-not $short) { return }
            $lc = $short.ToLower()
            $lcFull = $account.ToLower()
            $label = & $rightsLabel $rights

            if ($openTrustees -contains $lc) {
                $script:_pOpen = $true
                if (-not (@($perms | Where-Object { $_.Trustee -eq $short }).Count)) {
                    [void]$perms.Add(@{ Trustee=$short; Rights=$label; Kind='Everyone'; IsOpen=$true })
                }
                return
            }
            if (($adminTrustees -contains $lc) -or ($adminTrustees -contains $lcFull)) {
                if (-not (@($perms | Where-Object { $_.Trustee -eq $short }).Count)) {
                    [void]$perms.Add(@{ Trustee=$short; Rights=$label; Kind='Administrative'; IsOpen=$false })
                }
                return
            }
            # A real named principal. Only 'Print' rights gate who can actually
            # print, but record management rights too - useful for scoping.
            $kind = if ($account -match '\\') { 'Domain group or user' } else { 'Local group or user' }
            if (-not (@($perms | Where-Object { $_.Trustee -eq $short }).Count)) {
                [void]$perms.Add(@{ Trustee=$account; Rights=$label; Kind=$kind; IsOpen=$false })
            }
            if ($named -notcontains $short) { [void]$named.Add($short) }
        }

        try {
            $script:_pOpen = $false
            $sddl = ''
            if ($UseGetPrinter) {
                try {
                    $full = Get-Printer -Name $Printer.Name -Full -ErrorAction Stop
                    $sddl = [string]$full.PermissionSDDL
                } catch { cb-Log "PrintServer" "Get-Printer -Full failed for $($Printer.Name): $_" }
            }

            if ($sddl -and (Get-Command ConvertFrom-SddlString -ErrorAction SilentlyContinue)) {
                # PS 5.1+ - friendly account names and right names for free.
                $parsed = ConvertFrom-SddlString -Sddl $sddl -ErrorAction Stop
                foreach ($ace in @($parsed.DiscretionaryAcl)) {
                    $s = [string]$ace
                    $acct = ($s -split ':')[0].Trim()
                    $rights = ''
                    $mr = [regex]::Match($s, '\(([^)]*)\)')
                    if ($mr.Success) { $rights = $mr.Groups[1].Value }
                    & $addPerm $acct $rights
                }
            }
            elseif ($sddl) {
                # Server 2012/2012R2 (PS 3/4): ConvertFrom-SddlString does not
                # exist, so parse the DACL and translate SIDs by hand.
                $dacl = ''
                $mD = [regex]::Match($sddl, 'D:(.*?)(?=S:|$)')
                if ($mD.Success) { $dacl = $mD.Groups[1].Value }
                foreach ($ace in [regex]::Matches($dacl, '\(([^)]*)\)')) {
                    $parts = $ace.Groups[1].Value -split ';'
                    if ($parts.Count -lt 6) { continue }
                    $rightsRaw = $parts[2]
                    $trustee   = $parts[5]
                    if (-not $trustee) { continue }
                    $acct = $trustee
                    switch ($trustee) {
                        'WD' { $acct = 'Everyone' }
                        'AU' { $acct = 'Authenticated Users' }
                        'BA' { $acct = 'Administrators' }
                        'SY' { $acct = 'SYSTEM' }
                        'PU' { $acct = 'Power Users' }
                        'CO' { $acct = 'CREATOR OWNER' }
                        'DU' { $acct = 'Domain Users' }
                        'PO' { $acct = 'Print Operators' }
                        default {
                            if ($trustee -match '^S-1-') {
                                try {
                                    $sidObj = New-Object System.Security.Principal.SecurityIdentifier($trustee)
                                    $acct = $sidObj.Translate([System.Security.Principal.NTAccount]).Value
                                } catch { $acct = $trustee }
                            }
                        }
                    }
                    & $addPerm $acct $rightsRaw
                }
            }
            else {
                # No SDDL available (WMI-only path). Read the security descriptor
                # straight off the WMI object - works on every Windows version.
                try {
                    $sd = $null
                    if ($PSMaj -ge 3) {
                        $sd = (Invoke-CimMethod -InputObject $Printer -MethodName GetSecurityDescriptor -ErrorAction Stop).Descriptor
                    } else {
                        $sd = $Printer.GetSecurityDescriptor().Descriptor
                    }
                    foreach ($ace in @($sd.DACL)) {
                        $t = $ace.Trustee
                        if (-not $t) { continue }
                        $acct = if ($t.Domain) { "$($t.Domain)\$($t.Name)" } else { [string]$t.Name }
                        & $addPerm $acct ([string]$ace.AccessMask)
                    }
                } catch { cb-Log "PrintServer" "WMI SD read failed for $($Printer.Name): $_" }
            }
            $sawOpen = $script:_pOpen
        } catch { cb-Log "PrintServer" "ACL parse error for $($Printer.Name): $_" }

        $out.Groups      = @($named)
        $out.Permissions = @($perms)
        if ($named.Count -gt 0) { $out.Access = 'restricted' }
        elseif ($sawOpen)       { $out.Access = 'open' }
        else                    { $out.Access = 'unknown' }
        return $out
    }

    function Collect-PrintServer {
        Write-Host "  [PrintServer] Enumerating shared print queues + ACLs..." -ForegroundColor Gray
        $result = @{ IsPrintServer=$false; DetectedBy=''; Queues=@(); Partial=$false
                     Method=''; Diagnostics=@() }
        $pdiag = [System.Collections.ArrayList]@()
        $pnote = { param($m) [void]$pdiag.Add([string]$m); cb-Log "PrintServer" $m }
        try {
            # ---- Role signal. Get-WindowsFeature needs the ServerManager module
            # and is absent on client OS / some Core installs, so treat it as a
            # hint only, never a gate.
            $hasRole = $false
            try {
                if ($PSMaj -ge 4) { Import-Module ServerManager -ErrorAction SilentlyContinue }
                $feat = Get-WindowsFeature -Name 'Print-Server' -ErrorAction SilentlyContinue
                if ($feat -and $feat.Installed) { $hasRole = $true; $result.DetectedBy = 'Print-Server role' }
            } catch { }
            # Registry role hint - works with no modules at all
            if (-not $hasRole) {
                try {
                    if (Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers' -ErrorAction SilentlyContinue) {
                        $spooler = Get-Service -Name Spooler -ErrorAction SilentlyContinue
                        if ($spooler -and $spooler.Status -eq 'Running') { $hasRole = $false }  # hint only
                    }
                } catch { }
            }

            # ---- Shared queue enumeration: three independent methods so this
            # works on 2012 (no PrintManagement) through 2025.
            $shared = @()
            $useGetPrinter = $false

            # Method 1: Get-Printer (PrintManagement module, 2012+)
            if (Get-Command Get-Printer -ErrorAction SilentlyContinue) {
                try {
                    $shared = @(Get-Printer -ErrorAction Stop | Where-Object { $_.Shared -eq $true })
                    $useGetPrinter = $true
                    $result.Method = 'Get-Printer'
                } catch { & $pnote "Get-Printer failed ($($_.Exception.Message.Trim())) - trying WMI" }
            } else {
                & $pnote "Get-Printer not available on this OS - using WMI"
            }

            # Method 2: WMI Win32_Printer (every version incl. 2008 R2)
            if (-not $shared.Count) {
                $wmi = Safe-WmiQuery "SELECT * FROM Win32_Printer WHERE Shared=TRUE" "PrintServer"
                if (-not $wmi) { $wmi = Safe-Wmi Win32_Printer "PrintServer" | Where-Object { $_.Shared } }
                if ($wmi) {
                    $shared = @($wmi | Where-Object { $_.Name -notmatch 'Fax|PDF|XPS|OneNote' })
                    $useGetPrinter = $false
                    if ($shared.Count) { $result.Method = 'WMI Win32_Printer' }
                }
            }

            # Method 3: registry (spooler present but WMI/cmdlet both blocked)
            if (-not $shared.Count) {
                try {
                    $regBase = 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers'
                    if (Test-Path $regBase -ErrorAction SilentlyContinue) {
                        $regPrinters = Get-ChildItem $regBase -ErrorAction SilentlyContinue
                        $fromReg = @()
                        foreach ($rp in $regPrinters) {
                            $props = Get-ItemProperty $rp.PSPath -ErrorAction SilentlyContinue
                            $attr  = [int]($props.Attributes | Select-Object -First 1)
                            # bit 3 (0x8) = shared
                            if ($attr -band 8) {
                                $fromReg += [PSCustomObject]@{
                                    Name       = $rp.PSChildName
                                    ShareName  = [string]$props.'Share Name'
                                    DriverName = [string]$props.'Printer Driver'
                                    PortName   = [string]$props.Port
                                    Shared     = $true
                                }
                            }
                        }
                        if ($fromReg.Count) {
                            $shared = $fromReg
                            $result.Method = 'Registry (WMI + cmdlet unavailable)'
                            & $pnote "Enumerated $($fromReg.Count) shared queue(s) from the registry"
                        }
                    }
                } catch { & $pnote "Registry printer enumeration failed: $($_.Exception.Message.Trim())" }
            }

            if (-not $hasRole -and $shared.Count -eq 0) {
                & $pnote "No shared print queues and no Print-Server role - not a print server"
                $result.Diagnostics = @($pdiag)
                return $result
            }
            if (-not $result.DetectedBy) { $result.DetectedBy = 'Shared queues + Spooler' }
            $result.IsPrintServer = $true
            & $pnote "Found $($shared.Count) shared queue(s) via $($result.Method)"

            $depMap = Get-GPPrinterDeployments

            foreach ($p in $shared) {
                $pName  = [string]$p.Name
                $share  = [string]$p.ShareName
                $driver = if ($useGetPrinter) { [string]$p.DriverName } else { [string]$p.DriverName }
                $port   = if ($useGetPrinter) { [string]$p.PortName }   else { [string]$p.PortName }

                $acc = Get-PrinterAccessSummary -Printer $p -UseGetPrinter $useGetPrinter

                # Cross-reference GPP deployment by share name or UNC path.
                $deployGpo = ''
                $unc = if ($share) { "\\$env:COMPUTERNAME\$share" } else { '' }
                foreach ($cand in @($share, $pName, $unc)) {
                    if (-not $cand) { continue }
                    $k = $cand.ToLower().TrimStart('\')
                    if ($depMap.ContainsKey($k)) { $deployGpo = $depMap[$k]; break }
                }

                if ($acc.Access -eq 'restricted' -and $acc.Groups.Count -gt 0) {
                    cb-Flag 'info' "Restricted printer: $pName" "Print access limited to: $($acc.Groups -join ', '). These groups must be recreated in Entra/Intune (Universal Print) during migration."
                }

                $result.Queues += @{
                    Name        = $pName
                    ShareName   = $share
                    Driver      = $driver
                    Port        = $port
                    Access      = $acc.Access
                    AccessGroups = @($acc.Groups)
                    Permissions  = @($acc.Permissions)
                    DeployedByGPO = $deployGpo
                }
            }
        } catch { & $pnote "Outer error: $($_.Exception.Message.Trim())"; $result.Partial = $true }
        $result.Diagnostics = @($pdiag)
        return $result
    }

    # -- GROUP POLICY (linked GPOs only, deep-dive) ----------------------------

    # Parse a GPO report XML into category tokens (for the English synopsis) and
    # raw setting rows (for the expand view). Namespace-agnostic via local-name().
    function Parse-GPOReport {
        param([xml]$Report)
        $cats = [System.Collections.ArrayList]@()
        $raw  = [System.Collections.ArrayList]@()
        $addCat = { param($t) if ($t -and ($cats -notcontains $t)) { [void]$cats.Add($t) } }
        # Null-safe child-node text: returns '' when the node is absent so a
        # missing element never aborts the whole parse.
        $nodeText = {
            param($node, $xpath)
            try { $n = $node.SelectSingleNode($xpath); if ($n) { return [string]$n.InnerText } } catch { }
            return ''
        }
        try {
            # ADMX / registry-based administrative templates
            $admxCats = @{}
            foreach ($pol in $Report.SelectNodes("//*[local-name()='Policy']")) {
                $nm  = & $nodeText $pol "*[local-name()='Name']"
                $st  = & $nodeText $pol "*[local-name()='State']"
                $cat = & $nodeText $pol "*[local-name()='Category']"
                if ($nm) {
                    & $addCat 'admx'
                    if ($cat) { $admxCats[$cat] = $true }
                    if ($raw.Count -lt 400) {
                        [void]$raw.Add(@{ Area='Administrative Templates'; Category=$cat; Name=$nm; Setting=$st })
                    }
                }
            }
            foreach ($ac in $admxCats.Keys) { & $addCat ("admx:" + $ac) }

            # Group Policy Preferences: drive maps
            foreach ($d in $Report.SelectNodes("//*[local-name()='DriveMapSettings']//*[local-name()='Drive']")) {
                & $addCat 'drive_maps'
                if ($raw.Count -lt 400) {
                    $pr = $d.SelectSingleNode("*[local-name()='Properties']")
                    $letter = if ($pr) { [string]$pr.getAttribute('letter') } else { '' }
                    $path   = if ($pr) { [string]$pr.getAttribute('path') } else { '' }
                    [void]$raw.Add(@{ Area='Drive Maps'; Category=''; Name=$letter; Setting=$path })
                }
            }

            # Folder redirection
            foreach ($f in $Report.SelectNodes("//*[local-name()='Folder' and ancestor::*[local-name()='FolderRedirection']]")) {
                & $addCat 'folder_redirection'
                if ($raw.Count -lt 400) {
                    $id = [string]$f.getAttribute('Id')
                    $loc = & $nodeText $f ".//*[local-name()='DestinationPath']"
                    [void]$raw.Add(@{ Area='Folder Redirection'; Category=''; Name=$id; Setting=$loc })
                }
            }

            # Printer deployment (inside a GPO)
            foreach ($p in $Report.SelectNodes("//*[local-name()='SharedPrinter' or local-name()='PortPrinter']")) {
                & $addCat 'printers'
                if ($raw.Count -lt 400) {
                    [void]$raw.Add(@{ Area='Printer Deployment'; Category=''; Name=[string]$p.getAttribute('name'); Setting='' })
                }
            }

            # Scripts (logon/logoff/startup/shutdown)
            foreach ($s in $Report.SelectNodes("//*[local-name()='Script']")) {
                & $addCat 'scripts'
                if ($raw.Count -lt 400) {
                    $cmd = & $nodeText $s "*[local-name()='Command']"
                    [void]$raw.Add(@{ Area='Scripts'; Category=''; Name=$cmd; Setting='' })
                }
            }

            # Software installation (assigned/published MSI)
            foreach ($m in $Report.SelectNodes("//*[local-name()='MsiApplication']")) {
                & $addCat 'software_install'
                if ($raw.Count -lt 400) {
                    $nm = & $nodeText $m "*[local-name()='Name']"
                    [void]$raw.Add(@{ Area='Software Installation'; Category=''; Name=$nm; Setting='' })
                }
            }

            # Security options + account/password policy
            if ($Report.SelectNodes("//*[local-name()='SecurityOptions']").Count -gt 0) { & $addCat 'security_options' }
            if ($Report.SelectNodes("//*[local-name()='Account']").Count -gt 0)         { & $addCat 'account_policy' }
            foreach ($so in $Report.SelectNodes("//*[local-name()='SecurityOptions']")) {
                if ($raw.Count -lt 400) {
                    $kn = & $nodeText $so "*[local-name()='KeyName']"
                    $dn = & $nodeText $so "*[local-name()='Display']/*[local-name()='Name']"
                    $soName = if ($dn) { $dn } else { $kn }
                    [void]$raw.Add(@{ Area='Security Options'; Category=''; Name=$soName; Setting='' })
                }
            }
            foreach ($ac in $Report.SelectNodes("//*[local-name()='Account']")) {
                if ($raw.Count -lt 400) {
                    $nm = [string]$ac.getAttribute('Name')
                    $st = & $nodeText $ac "*[local-name()='SettingNumber']"
                    [void]$raw.Add(@{ Area='Account Policy'; Category=''; Name=$nm; Setting=$st })
                }
            }
        } catch { cb-Log "GPO" "Report parse error: $_" }
        return @{ Categories = @($cats); Raw = @($raw) }
    }

    # Extended-right GUID for "Apply Group Policy". An ACE granting this is what
    # actually decides whether a GPO takes effect for a principal - the OU link
    # only decides where it is *considered*.
    $script:APPLY_GP_RIGHT = 'edacfd8f-ffb3-11d1-b41d-00a0c968f939'

    function Get-GPOSecurityFiltering {
        <#
            Answer "who does this GPO actually apply to?".

            A GPO linked to an OU does NOT necessarily hit everything in that
            OU - security filtering narrows it. By default Authenticated Users
            holds Apply Group Policy, which means everyone; admins frequently
            replace that with a specific group. That distinction is exactly what
            has to be reproduced as an Intune assignment group, so it is worth
            surfacing next to the links.

            Returns:
              Scope       'everyone' | 'filtered' | 'unknown'
              Principals  [ @{ Trustee; Kind } ]  - who holds Apply Group Policy
              WmiFilter   name/text of any WMI filter, or ''
        #>
        param([string]$Guid, [string]$DomainDN, [bool]$HaveModule)

        $out = @{ Scope='unknown'; Principals=@(); WmiFilter='' }
        $everyoneTrustees = @('authenticated users','domain computers','domain users','everyone')
        $principals = [System.Collections.ArrayList]@()
        $sawEveryone = $false

        $addP = {
            param([string]$account)
            if (-not $account) { return }
            $short = ($account -replace '^.*\\', '').Trim()
            if (-not $short) { return }
            if ($everyoneTrustees -contains $short.ToLower()) {
                $script:_gpEveryone = $true
                if (-not (@($principals | Where-Object { $_.Trustee -eq $short }).Count)) {
                    [void]$principals.Add(@{ Trustee=$short; Kind='Everyone' })
                }
                return
            }
            $kind = if ($account -match '\\') { 'Domain group or user' } else { 'Group or user' }
            if (-not (@($principals | Where-Object { $_.Trustee -eq $account }).Count)) {
                [void]$principals.Add(@{ Trustee=$account; Kind=$kind })
            }
        }

        try {
            $script:_gpEveryone = $false

            # ---- Method A: GroupPolicy module ----
            $gotAny = $false
            if ($HaveModule -and (Get-Command Get-GPPermission -ErrorAction SilentlyContinue)) {
                try {
                    $bare = $Guid.Trim('{','}')
                    foreach ($perm in (Get-GPPermission -Guid $bare -All -ErrorAction Stop)) {
                        if ([string]$perm.Permission -eq 'GpoApply') {
                            $t = $perm.Trustee
                            $acct = if ($t.Domain) { "$($t.Domain)\$($t.Name)" } else { [string]$t.Name }
                            & $addP $acct
                            $gotAny = $true
                        }
                    }
                } catch { cb-Log "GPO" "Get-GPPermission failed for $Guid - falling back to LDAP ACL" }
            }

            # ---- Method B: LDAP ACL on the groupPolicyContainer object ----
            # Works with no modules: look for ACEs granting the Apply Group
            # Policy extended right.
            if (-not $gotAny -and $DomainDN) {
                try {
                    $gpoDN = "CN=$Guid,CN=Policies,CN=System,$DomainDN"
                    $de = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$gpoDN")
                    $sec = $de.ObjectSecurity
                    if ($sec) {
                        foreach ($ace in $sec.GetAccessRules($true, $true, [System.Security.Principal.NTAccount])) {
                            if ([string]$ace.AccessControlType -ne 'Allow') { continue }
                            $rightOk = $false
                            try {
                                if ($ace.ObjectType -and ([string]$ace.ObjectType).ToLower() -eq $script:APPLY_GP_RIGHT) { $rightOk = $true }
                            } catch { }
                            if ($rightOk) {
                                & $addP ([string]$ace.IdentityReference)
                                $gotAny = $true
                            }
                        }
                    }
                    $de.Close()
                } catch { cb-Log "GPO" "LDAP ACL read failed for $Guid : $($_.Exception.Message.Trim())" }
            }

            # ---- WMI filter (further narrows applicability) ----
            if ($DomainDN) {
                try {
                    $gpoDN = "CN=$Guid,CN=Policies,CN=System,$DomainDN"
                    $de2 = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$gpoDN")
                    $wq = [string]$de2.Properties['gPCWQLFilter'].Value
                    if ($wq) {
                        # Format: [domain;{GUID};0] - resolve to the filter's name
                        $mg = [regex]::Match($wq, '\{[0-9A-Fa-f-]+\}')
                        if ($mg.Success) {
                            $fDN = "CN=$($mg.Value),CN=SOM,CN=WMIPolicy,CN=System,$DomainDN"
                            try {
                                $fde = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$fDN")
                                $fname = [string]$fde.Properties['msWMI-Name'].Value
                                $out.WmiFilter = if ($fname) { $fname } else { $mg.Value }
                                $fde.Close()
                            } catch { $out.WmiFilter = $mg.Value }
                        }
                    }
                    $de2.Close()
                } catch { }
            }

            $sawEveryone = $script:_gpEveryone
            if (-not $gotAny)      { $out.Scope = 'unknown' }
            elseif ($sawEveryone -and @($principals | Where-Object { $_.Kind -ne 'Everyone' }).Count -eq 0) {
                $out.Scope = 'everyone'
            }
            elseif ($sawEveryone)  { $out.Scope = 'everyone' }
            else                   { $out.Scope = 'filtered' }
        } catch { cb-Log "GPO" "Security filtering read failed for $Guid : $_" }

        $out.Principals = @($principals)
        return $out
    }

    function Collect-GPODetails {
        <#
            Collect every LINKED GPO, with a self-healing method chain so this
            works on Server 2012 through 2025 and on DCs, member servers, and
            Server Core alike.

              Method A  GroupPolicy module (Get-GPO / Get-GPInheritance /
                        Get-GPOReport). Richest data. Needs RSAT-GPMC, which
                        is present on most DCs but NOT on member servers or
                        many Core installs.
              Method B  Raw ADSI/LDAP. Reads the groupPolicyContainer objects
                        under CN=Policies,CN=System and the gPLink attribute on
                        the domain root / every OU / every site. Requires NO
                        modules at all - any domain member can do this.
              Method C  SYSVOL folder inspection per GPO GUID. Infers which
                        policy areas a GPO configures from the folders/files
                        present (Scripts, Preferences\Drives, Registry.pol...).
                        Used to fill in settings detail when Get-GPOReport is
                        unavailable.

            The previous version bailed out entirely when Get-GPO was missing,
            which is why real engagements came back with an empty GPO section
            and no explanation. Whatever happens now, Diagnostics records WHY.
        #>
        Write-Host "  [GPO] Enumerating linked GPOs..." -ForegroundColor Gray
        $result = @{
            Available=$false; TotalCount=0; LinkedCount=0; GPOs=@(); Partial=$false
            Method=''; Diagnostics=@()
        }
        $diag = [System.Collections.ArrayList]@()
        $note = { param($m) [void]$diag.Add([string]$m); cb-Log "GPO" $m }

        $MAX_OUS      = 750
        $GPO_BUDGET_S = 120
        $gpoSw        = [System.Diagnostics.Stopwatch]::StartNew()

        # ---------- domain context (needed by every method) ----------
        $domainDN = ''
        $domainDNS = ''
        try {
            $rootDSE   = [ADSI]"LDAP://RootDSE"
            $domainDN  = [string]$rootDSE.defaultNamingContext
            $cfgDN     = [string]$rootDSE.configurationNamingContext
        } catch { $cfgDN = '' }
        if (-not $domainDN) {
            & $note "Not domain-joined or LDAP RootDSE unreachable - no Group Policy to collect"
            $result.Diagnostics = @($diag); return $result
        }
        try { $domainDNS = $env:USERDNSDOMAIN } catch { }
        if (-not $domainDNS) {
            $domainDNS = (($domainDN -split ',') | Where-Object { $_ -match '^DC=' } |
                          ForEach-Object { $_.Substring(3) }) -join '.'
        }

        # ---------- helper: parse a gPLink attribute string ----------
        # Format: [LDAP://cn={GUID},cn=policies,cn=system,DC=x;FLAGS] repeated.
        # FLAGS: 0 = enabled, 1 = disabled, 2 = enforced, 3 = disabled+enforced
        $parseGPLink = {
            param([string]$gpLink, [string]$targetDN)
            $out = @()
            if (-not $gpLink) { return $out }
            foreach ($m in [regex]::Matches($gpLink, '\[LDAP://([^;]+);(\d+)\]')) {
                $dn    = $m.Groups[1].Value
                $flags = [int]$m.Groups[2].Value
                $g = [regex]::Match($dn, '\{[0-9A-Fa-f-]+\}')
                if (-not $g.Success) { continue }
                $out += @{
                    Guid     = $g.Value.ToUpper()
                    Target   = $targetDN
                    Enabled  = (($flags -band 1) -eq 0)
                    Enforced = (($flags -band 2) -ne 0)
                }
            }
            return $out
        }

        # ---------- METHOD A: GroupPolicy module ----------
        $haveGPModule = $false
        if (Get-Command Get-GPO -ErrorAction SilentlyContinue) { $haveGPModule = $true }
        else {
            try { Import-Module GroupPolicy -ErrorAction Stop; $haveGPModule = $true }
            catch { & $note "GroupPolicy module not available ($($_.Exception.Message.Trim())) - falling back to LDAP" }
        }

        # ---------- Build GUID -> display name, and GUID -> links ----------
        $nameMap = @{}     # {GUID} -> display name
        $linkMap = @{}     # {GUID} -> list of link hashtables

        if ($haveGPModule) {
            try {
                $all = @(Get-GPO -All -ErrorAction Stop)
                foreach ($g in $all) { $nameMap[("{$($g.Id)}").ToUpper()] = [string]$g.DisplayName }
                $result.TotalCount = $all.Count
                $result.Method = 'GroupPolicy module'
            } catch {
                & $note "Get-GPO -All failed: $($_.Exception.Message.Trim())"
                $haveGPModule = $false
            }
        }

        # LDAP enumeration of GPOs - always run when the module path didn't
        # produce names, and used as the authoritative link source either way
        # (gPLink is cheap to read and needs no modules).
        if (-not $nameMap.Count) {
            try {
                $polDN = "CN=Policies,CN=System,$domainDN"
                $searcher = New-Object System.DirectoryServices.DirectorySearcher
                $searcher.SearchRoot = [ADSI]"LDAP://$polDN"
                $searcher.Filter     = '(objectClass=groupPolicyContainer)'
                $searcher.PageSize   = 200
                [void]$searcher.PropertiesToLoad.Add('displayname')
                [void]$searcher.PropertiesToLoad.Add('cn')
                foreach ($r in $searcher.FindAll()) {
                    $cn = [string]$r.Properties['cn'][0]
                    $dn = if ($r.Properties['displayname'].Count) { [string]$r.Properties['displayname'][0] } else { $cn }
                    if ($cn) { $nameMap[$cn.ToUpper()] = $dn }
                }
                $searcher.Dispose()
                $result.TotalCount = $nameMap.Count
                if (-not $result.Method) { $result.Method = 'LDAP (no GroupPolicy module)' }
                & $note "LDAP enumeration found $($nameMap.Count) GPO object(s)"
            } catch {
                & $note "LDAP GPO enumeration failed: $($_.Exception.Message.Trim())"
            }
        }

        # ---------- Link discovery: gPLink on domain root + OUs + sites ----------
        # This is module-free and is the piece that makes 'which GPOs actually
        # apply' work on a member server or Core install.
        $ouCount = 0
        try {
            # Domain root
            try {
                $rootObj = [ADSI]"LDAP://$domainDN"
                foreach ($lk in (& $parseGPLink ([string]$rootObj.gPLink) $domainDN)) {
                    if (-not $linkMap.ContainsKey($lk.Guid)) { $linkMap[$lk.Guid] = @() }
                    $linkMap[$lk.Guid] += $lk
                }
            } catch { & $note "Could not read gPLink on domain root: $($_.Exception.Message.Trim())" }

            # Every OU
            $ouSearcher = New-Object System.DirectoryServices.DirectorySearcher
            $ouSearcher.SearchRoot = [ADSI]"LDAP://$domainDN"
            $ouSearcher.Filter     = '(objectClass=organizationalUnit)'
            $ouSearcher.PageSize   = 200
            [void]$ouSearcher.PropertiesToLoad.Add('gplink')
            [void]$ouSearcher.PropertiesToLoad.Add('distinguishedname')
            foreach ($r in $ouSearcher.FindAll()) {
                if ($ouCount -ge $MAX_OUS) { & $note "OU cap ($MAX_OUS) reached - link data may be partial"; $result.Partial = $true; break }
                if ($gpoSw.Elapsed.TotalSeconds -gt $GPO_BUDGET_S) { & $note "OU link walk exceeded ${GPO_BUDGET_S}s - stopping early"; $result.Partial = $true; break }
                $ouCount++
                $odn = [string]$r.Properties['distinguishedname'][0]
                $gpl = if ($r.Properties['gplink'].Count) { [string]$r.Properties['gplink'][0] } else { '' }
                foreach ($lk in (& $parseGPLink $gpl $odn)) {
                    if (-not $linkMap.ContainsKey($lk.Guid)) { $linkMap[$lk.Guid] = @() }
                    $linkMap[$lk.Guid] += $lk
                }
            }
            $ouSearcher.Dispose()

            # Sites (live in the configuration partition)
            if ($cfgDN) {
                try {
                    $siteSearcher = New-Object System.DirectoryServices.DirectorySearcher
                    $siteSearcher.SearchRoot = [ADSI]"LDAP://CN=Sites,$cfgDN"
                    $siteSearcher.Filter     = '(objectClass=site)'
                    [void]$siteSearcher.PropertiesToLoad.Add('gplink')
                    [void]$siteSearcher.PropertiesToLoad.Add('distinguishedname')
                    foreach ($r in $siteSearcher.FindAll()) {
                        $sdn = [string]$r.Properties['distinguishedname'][0]
                        $gpl = if ($r.Properties['gplink'].Count) { [string]$r.Properties['gplink'][0] } else { '' }
                        foreach ($lk in (& $parseGPLink $gpl $sdn)) {
                            if (-not $linkMap.ContainsKey($lk.Guid)) { $linkMap[$lk.Guid] = @() }
                            $linkMap[$lk.Guid] += $lk
                        }
                    }
                    $siteSearcher.Dispose()
                } catch { }
            }
            & $note "Scanned $ouCount OU(s); found links for $($linkMap.Keys.Count) GPO(s)"
        } catch {
            & $note "OU/site link enumeration failed: $($_.Exception.Message.Trim())"
            $result.Partial = $true
        }

        if (-not $nameMap.Count -and -not $linkMap.Keys.Count) {
            & $note "No GPO data obtainable by any method"
            $result.Diagnostics = @($diag)
            return $result
        }
        $result.Available   = $true
        $result.LinkedCount = $linkMap.Keys.Count
        if (-not $result.TotalCount) { $result.TotalCount = $nameMap.Count }

        # ---------- Per-GPO detail ----------
        $reportBudgetS = $GPO_BUDGET_S + 180
        $sysvolRoot = "\\$domainDNS\SYSVOL\$domainDNS\Policies"

        foreach ($guid in $linkMap.Keys) {
            $links = @($linkMap[$guid] | ForEach-Object {
                @{ Target=$_.Target; Enabled=$_.Enabled; Enforced=$_.Enforced }
            })
            $dispName = if ($nameMap.ContainsKey($guid)) { $nameMap[$guid] } else { $guid }

            $cats = @(); $raw = @(); $status = ''
            $gotDetail = $false

            # Preferred: full settings report via the module
            if ($haveGPModule -and $gpoSw.Elapsed.TotalSeconds -le $reportBudgetS) {
                try {
                    $bare = $guid.Trim('{','}')
                    [xml]$rpt = Get-GPOReport -Guid $bare -ReportType Xml -ErrorAction Stop
                    $parsed = Parse-GPOReport -Report $rpt
                    $cats = $parsed.Categories
                    $raw  = $parsed.Raw
                    $gotDetail = $true
                } catch {
                    & $note "Get-GPOReport failed for $dispName - using SYSVOL inspection"
                }
            }

            # Fallback: infer configured areas from the GPO's SYSVOL folder.
            # No modules needed and works on every Windows Server version.
            if (-not $gotDetail) {
                try {
                    $gpDir = Join-Path $sysvolRoot $guid
                    if (Test-Path $gpDir -ErrorAction SilentlyContinue) {
                        $probe = @(
                            @{ Path='Machine\Registry.pol';               Cat='admx';               Area='Administrative Templates (Computer)' },
                            @{ Path='User\Registry.pol';                  Cat='admx';               Area='Administrative Templates (User)' },
                            @{ Path='Machine\Preferences\Drives';         Cat='drive_maps';         Area='Drive Maps' },
                            @{ Path='User\Preferences\Drives';            Cat='drive_maps';         Area='Drive Maps' },
                            @{ Path='Machine\Preferences\Printers';       Cat='printers';           Area='Printer Deployment' },
                            @{ Path='User\Preferences\Printers';          Cat='printers';           Area='Printer Deployment' },
                            @{ Path='Machine\Scripts';                    Cat='scripts';            Area='Scripts' },
                            @{ Path='User\Scripts';                       Cat='scripts';            Area='Scripts' },
                            @{ Path='Machine\Applications';               Cat='software_install';   Area='Software Installation' },
                            @{ Path='Machine\Microsoft\Windows NT\SecEdit\GptTmpl.inf'; Cat='security_options'; Area='Security Settings' },
                            @{ Path='Machine\Documents & Settings\fdeploy1.ini';        Cat='folder_redirection'; Area='Folder Redirection' },
                            @{ Path='User\Documents & Settings\fdeploy1.ini';           Cat='folder_redirection'; Area='Folder Redirection' }
                        )
                        foreach ($pr in $probe) {
                            $full = Join-Path $gpDir $pr.Path
                            if (Test-Path $full -ErrorAction SilentlyContinue) {
                                if ($cats -notcontains $pr.Cat) { $cats += $pr.Cat }
                                $raw += @{ Area=$pr.Area; Category=''; Name='(present in SYSVOL)'; Setting='detected' }
                            }
                        }
                        if ($cats.Count) { $gotDetail = $true }
                    }
                } catch { }
            }

            # Who the policy actually applies to (security filtering). A link
            # only says where the GPO is considered; Apply Group Policy says who
            # it hits. Both are needed to reproduce scope in Intune.
            $secf = Get-GPOSecurityFiltering -Guid $guid -DomainDN $domainDN -HaveModule $haveGPModule

            $result.GPOs += @{
                Name        = [string]$dispName
                Id          = [string]$guid
                Status      = [string]$status
                Links       = $links
                Categories  = @($cats)
                RawSettings = @($raw)
                AppliesToScope = [string]$secf.Scope
                AppliesTo      = @($secf.Principals)
                WmiFilter      = [string]$secf.WmiFilter
                DetailLevel = $(if ($gotDetail -and $haveGPModule) { 'full' } elseif ($gotDetail) { 'sysvol-inferred' } else { 'links-only' })
            }
        }

        if (-not $result.GPOs.Count) {
            & $note "GPO objects found but none are linked anywhere"
        }
        $result.Diagnostics = @($diag)
        return $result
    }


    # -- ASSEMBLE & RETURN -----------------------------------------------------

    $Discovery = [ordered]@{
        Meta = @{
            ScriptVersion  = '1.0'
            CollectedAt    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            PSVersionTarget = "$PSMaj.$PSMin"
        }
        System     = Collect-SystemInfo
        Hardware   = Collect-Hardware
        Disks      = Collect-Disks
        Network    = Collect-Network
        Roles      = Collect-Roles
        AD         = Collect-ADDetails
        DNS        = Collect-DNSDetails
        DHCP       = Collect-DHCPDetails
        FileShares = Collect-FileShares
        NPS        = Collect-NPSDetails
        IIS        = Collect-IISDetails
        SQL        = Collect-SQLDetails
        Exchange   = Collect-ExchangeDetails
        HyperV     = Collect-HyperVDetails
        Apps       = Collect-InstalledApps
        Tasks      = Collect-ScheduledTasks
        Services   = Collect-Services
        EventLog   = Collect-EventLogSummary
        Printers   = Collect-Printers
        PrintServer = Collect-PrintServer
        GPO        = Collect-GPODetails
        Flags      = $cbFlags
        Errors     = $cbErrors
    }

    # Fix 7 -- ensure array fields are never $null (prevents JSON key omission on PS 3/4)
    # NESTED array fields: section is a hashtable, inner field is a list.
    $arrayFields = @{
        'Network'  = @('Adapters','ListeningPorts','EstablishedConns')
        'Roles'    = @('InstalledRoles','InstalledFeatures')
        'AD'       = @('StaleUsers','StaleComputers','FSMORoles')
        'DNS'      = @('Zones','Forwarders')
        'DHCP'     = @('Scopes')
        'FileShares' = @('Shares')
        'NPS'      = @('Clients','Policies')
        'IIS'      = @('Sites','AppPools')
        'SQL'      = @('Instances')
        'Exchange' = @('DatabaseSizes')
        'HyperV'   = @('VMs','VirtualSwitches')
        'EventLog' = @('TopSources','RecentCritical')
        'PrintServer' = @('Queues')
        'GPO'      = @('GPOs')
    }
    foreach ($section in $arrayFields.Keys) {
        $sec = $Discovery[$section]
        # Only walk nested fields when the section is a dictionary. A plain
        # array section (Disks, Apps, Tasks, ...) is handled separately below.
        if ($sec -is [System.Collections.IDictionary]) {
            foreach ($field in $arrayFields[$section]) {
                if ($null -eq $sec[$field]) { $sec[$field] = @() }
            }
        }
    }
    # Top-level array sections returned directly (flat lists, no subkeys)
    foreach ($topArr in @('Disks','Apps','Tasks','Services','Printers')) {
        if ($null -eq $Discovery[$topArr]) { $Discovery[$topArr] = @() }
    }

    return $Discovery
}

# -----------------------------------------------------------------------------
# EXECUTE COLLECTION
# -----------------------------------------------------------------------------

Write-BuddyPhase "Collection" "starting full discovery - this takes 2-5 minutes..."
Write-Host ""

$discoveryResult = $null

if ($script:IsRemote) {
    $transportLabel =
        if ($script:IsHVGuest) {
            $tgt = if ($HyperVGuestVMId) { "VMId=$HyperVGuestVMId" } else { "VMName=$HyperVGuestVMName" }
            "Hyper-V PowerShell Direct ($tgt)"
        } else { "WinRM to $ComputerName" }
    Write-Host ("  Sending collection block via $transportLabel...") -ForegroundColor DarkCyan
    Write-Host ("  Remote execution is synchronous - buddy will animate while waiting.") -ForegroundColor DarkGray
    Write-Host ""

    # Spinner while remote job runs
    $job = $null
    try {
        # IMPORTANT: Must use Invoke-Command -AsJob, NOT Start-Job { Invoke-Command }.
        # In PS5.1, passing a Hashtable with a ScriptBlock through Start-Job -ArgumentList
        # serializes the ScriptBlock to a string. Invoke-Command then returns that string
        # as output instead of executing it. -AsJob bypasses serialization entirely.
        $icParams = @{
            ScriptBlock  = $CollectionBlock
            AsJob        = $true
            ErrorAction  = 'Stop'
        }
        if ($script:IsHVGuest) {
            # PowerShell Direct -- routes through Hyper-V VMBus, no network.
            # Requires Credential (PS Direct does not support implicit creds).
            if (-not $Credential) {
                throw "PowerShell Direct requires -Credential (no implicit auth via VMBus). Pass a local-guest or domain credential."
            }
            if ($HyperVGuestVMId)   { $icParams.VMId   = $HyperVGuestVMId }
            elseif ($HyperVGuestVMName) { $icParams.VMName = $HyperVGuestVMName }
            $icParams.Credential = $Credential
        } else {
            $icParams.ComputerName = $ComputerName
            if ($Credential) { $icParams.Credential = $Credential }
        }

        $job = Invoke-Command @icParams

        $spinChars = @('|','/','-','\')
        $spinIdx   = 0
        while ($job.State -eq 'Running') {
            $frame = $buddyFrames[$spinIdx % $buddyFrames.Count]
            if ($script:IsHVGuest) {
                $gid = if ($HyperVGuestVMName) { $HyperVGuestVMName } else { $HyperVGuestVMId }
                $waitTarget = "$gid (PS Direct)"
            } else { $waitTarget = $ComputerName }
            Write-Host ("`r  $frame  waiting for $waitTarget...  $($spinChars[$spinIdx % 4])   ") -NoNewline -ForegroundColor DarkCyan
            Start-Sleep -Milliseconds 400
            $spinIdx++
        }
        Write-Host "`r  (^_^)  Remote collection complete.                              " -ForegroundColor DarkGreen
        $discoveryResult = Receive-Job -Job $job -ErrorAction Stop
        Remove-Job -Job $job -Force
    } catch {
        $remoteErr = $_.ToString()
        Write-Host ""
        Write-Host ("  (x_x)  Remote collection FAILED on $ComputerName") -ForegroundColor Red
        Write-Host ("         $remoteErr") -ForegroundColor DarkRed
        if ($script:IsNonInteractive) {
            if ($job) { Remove-Job -Job $job -Force -EA SilentlyContinue }
            Write-BuddyErr "RemoteCollect" "$ComputerName collection failed (non-interactive - moving on)"
            exit 1
        }
        Write-Host ""
        Write-Host "  What would you like to do?" -ForegroundColor Yellow
        Write-Host "    [R] Retry this server" -ForegroundColor DarkGray
        Write-Host "    [S] Skip this server and continue" -ForegroundColor DarkGray
        Write-Host "    [Q] Quit" -ForegroundColor DarkGray
        $ans = Read-Host "  Choice (R/S/Q)"
        if ($job) { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue }
        switch ($ans.Trim().ToUpper()) {
            'R' {
                Write-Host "  Re-running discovery against $ComputerName..." -ForegroundColor Cyan
                & $PSCommandPath -ComputerName $ComputerName -OutputPath $OutputPath -Credential $Credential
                exit $LASTEXITCODE
            }
            'S' {
                Write-BuddyWarn "Skipping $ComputerName. No JSON saved for this server."
                exit 0
            }
            default {
                Write-Host "  Exiting." -ForegroundColor Gray
                exit 1
            }
        }
    }
} else {
    # Local - run inline, progress prints from within the block
    try {
        $discoveryResult = & $CollectionBlock
    } catch {
        Write-BuddyErr "LocalExecution" "Collection block failed: $_"
        exit 1
    }
}

# -----------------------------------------------------------------------------
# OUTPUT
# -----------------------------------------------------------------------------

# Strip complex .NET/CIM/WMI types before ConvertTo-Json.
# Select-Object on objects like Get-WindowsFeature, Get-DnsServerZone, Get-Website, etc.
# retains the original .NET base object - ConvertTo-Json then walks into SubFeatures
# trees / circular property chains and either throws or hangs.
# This function converts everything to plain hashtables / arrays / primitives first.
function ConvertTo-SafeObject {
    param($Obj, [int]$Depth = 0)
    if ($Depth -gt 20) { return '[max depth]' }
    if ($null -eq $Obj) { return $null }
    if ($Obj -is [string])      { return $Obj }
    if ($Obj -is [bool])        { return $Obj }
    if ($Obj -is [datetime])    { return $Obj.ToString('yyyy-MM-dd HH:mm:ss') }
    if ($Obj -is [System.Enum]) { return $Obj.ToString() }
    # PS metadata objects - never recurse into these, they carry self-refs
    $typeName = ''
    try { $typeName = $Obj.GetType().FullName } catch {}
    if ($typeName -match 'System\.Management\.Automation\.PS(Parameterized|Method|Event|DynamicMember)') {
        return '[ps-member]'
    }
    # CIM/WMI instances - extract only the property bag, ignore metadata
    if ($typeName -match 'Microsoft\.Management\.Infrastructure\.CimInstance' -or
        $typeName -match 'System\.Management\.ManagementBaseObject') {
        $h = [ordered]@{}
        try {
            foreach ($p in $Obj.CimInstanceProperties) {
                try { $h[$p.Name] = ConvertTo-SafeObject $p.Value ($Depth+1) } catch { $h[$p.Name] = $null }
            }
        } catch {
            # Fall back to dict-like enumeration
            try { foreach ($k in $Obj.Keys) { $h["$k"] = ConvertTo-SafeObject $Obj[$k] ($Depth+1) } } catch { }
        }
        return $h
    }
    $t = $Obj.GetType()
    if ($t.IsPrimitive -or $t.IsValueType) { return $Obj }
    # Unwrap PSObject so we operate on the base
    $base = if ($Obj -is [System.Management.Automation.PSObject]) { $Obj.PSObject.BaseObject } else { $Obj }
    if ($base -is [string]) { return [string]$base }
    if ($base -is [System.Collections.IDictionary]) {
        $h = [ordered]@{}
        foreach ($k in @($base.Keys)) {
            try { $h["$k"] = ConvertTo-SafeObject $base[$k] ($Depth+1) } catch { $h["$k"] = $null }
        }
        return $h
    }
    if ($base -is [System.Collections.IEnumerable]) {
        $arr = [System.Collections.ArrayList]@()
        foreach ($item in $base) {
            try { [void]$arr.Add((ConvertTo-SafeObject $item ($Depth+1))) } catch { [void]$arr.Add($null) }
        }
        return ,$arr
    }
    # PSCustomObject or complex .NET object ? enumerate visible PS properties only
    if ($Obj -is [System.Management.Automation.PSObject]) {
        $h = [ordered]@{}
        foreach ($prop in $Obj.PSObject.Properties) {
            # Skip indexer-style properties (e.g. Item[int] on CIM/WMI objects).
            # ConvertTo-Json can't serialize them and they cause circular walks.
            if ($prop -is [System.Management.Automation.PSParameterizedProperty]) { continue }
            if ($prop.MemberType -eq 'ParameterizedProperty') { continue }
            # Skip known noisy / circular-prone property names from CIM/WMI/AD
            if ($prop.Name -in @('PSAdapted','PSBase','PSObject','PSComputerName','PSShowComputerName','CimClass','CimInstanceProperties','CimSystemProperties','SyncRoot','Site','Parent','SubFeatures','ParentFeature','PropertyNames')) { continue }
            $v = $null
            try { $v = $prop.Value } catch { continue }
            try { $h[$prop.Name] = ConvertTo-SafeObject $v ($Depth+1) } catch { $h[$prop.Name] = $null }
        }
        return $h
    }
    # Last-resort string conversion with guard
    try { return $Obj.ToString() } catch { return $null }
}

Write-BuddyPhase "Output" "assembling and saving JSON..."

if (-not $discoveryResult) {
    Write-BuddyErr "Output" "No discovery result returned. Nothing to save."
    exit 1
}

# Add any main-script-level errors
# NOTE: $discoveryResult.Errors may be a deserialized ArrayList from a remote job.
# Deserialized collections are read-only - .Add() throws. Wrap each call.
foreach ($e in $script:CollectErrors) {
    try { $discoveryResult.Errors.Add($e) | Out-Null } catch { }
}

# Guard: $OutputPath might be empty if the script was re-launched via splatting
# with no bound params. Fall back to the script directory, then cwd.
if (-not $OutputPath) { $OutputPath = $PSScriptRoot }
if (-not $OutputPath) { $OutputPath = (Get-Location).Path }

# Validate output path exists before attempting write
if (-not (Test-Path $OutputPath)) {
    try {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        Write-BuddyOK "Created output directory: $OutputPath"
    } catch {
        Write-BuddyErr "Output" "OutputPath does not exist and could not be created: $OutputPath - $_"
        # Fall back to script directory
        $OutputPath = $PSScriptRoot
        Write-BuddyWarn "Falling back to script directory: $OutputPath"
    }
}

$hostname  = $discoveryResult.System.Hostname
if (-not $hostname) { $hostname = $env:COMPUTERNAME }
$dateStr   = (Get-Date).ToString("yyyy-MM-dd")
$filename  = "${hostname}-discovery-${dateStr}.json"
$outputFile = Join-Path $OutputPath $filename

# Sanitize first ? ConvertTo-SafeObject strips all CIM/WMI/.NET types to plain
# hashtables/arrays/primitives so ConvertTo-Json never walks into circular graphs.
Write-Host "  Sanitizing result object..." -ForegroundColor DarkGray
$safeResult = ConvertTo-SafeObject $discoveryResult

# Second pass: strict allow-list. Only null/string/bool/numeric/hashtable/array
# survive. Anything else gets stringified. If the first pass missed an exotic
# type (PSParameterizedProperty, CimClass, etc.) this pass guarantees JSON-safe.
function Invoke-StrictSanitize {
    param($Obj, [int]$Depth = 0)
    if ($Depth -gt 30) { return '[deep]' }
    if ($null -eq $Obj) { return $null }
    if ($Obj -is [string] -or $Obj -is [bool]) { return $Obj }
    if ($Obj -is [int] -or $Obj -is [long] -or $Obj -is [double] -or $Obj -is [decimal] -or $Obj -is [single] -or $Obj -is [byte]) { return $Obj }
    if ($Obj -is [datetime]) { return $Obj.ToString('yyyy-MM-dd HH:mm:ss') }
    if ($Obj -is [System.Collections.IDictionary]) {
        $h = [ordered]@{}
        foreach ($k in @($Obj.Keys)) {
            try { $h["$k"] = Invoke-StrictSanitize $Obj[$k] ($Depth+1) } catch { $h["$k"] = $null }
        }
        return $h
    }
    if ($Obj -is [System.Collections.IEnumerable]) {
        $a = [System.Collections.ArrayList]@()
        foreach ($i in $Obj) {
            try { [void]$a.Add((Invoke-StrictSanitize $i ($Depth+1))) } catch { [void]$a.Add($null) }
        }
        return ,$a
    }
    try { return [string]$Obj } catch { return $null }
}
$safeResult = Invoke-StrictSanitize $safeResult

# Serialize to JSON
$json = $null
$jsonErr = ''
if ($localPSMajor -ge 3) {
    foreach ($depth in @(20, 10, 6)) {
        try {
            $json = $safeResult | ConvertTo-Json -Depth $depth -ErrorAction Stop
            Write-BuddyOK "Serialized at depth $depth"
            break
        } catch {
            $jsonErr = "depth $depth failed: $_"
            Write-BuddyWarn "ConvertTo-Json depth $depth failed - trying lower..."
        }
    }
} else {
    # PS 2.0 fallback - use New-Object (::new() syntax requires PS 5.0+)
    try {
        Add-Type -AssemblyName System.Web.Extensions -ErrorAction Stop
        $serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
        $serializer.MaxJsonLength = 50MB
        $json = $serializer.Serialize($safeResult)
    } catch { $jsonErr = $_.ToString() }
}

if ($json) {
    try {
        [System.IO.File]::WriteAllText($outputFile, $json, [System.Text.Encoding]::UTF8)
        Write-BuddyOK "Saved: $outputFile"
    } catch {
        Write-BuddyErr "Output" "File write failed: $_ -- Path: $outputFile"
        exit 1
    }
} else {
    Write-BuddyWarn "ConvertTo-Json failed: $jsonErr"
    Write-Host "  Auto-recovery: exporting to CLI-XML and re-importing to strip problem types..." -ForegroundColor DarkYellow

    # AUTO-RECOVERY: Export-Clixml -> Import-Clixml roundtrip strips
    # PSParameterizedProperty members (they don't survive CLI-XML deserialization).
    # Then ConvertTo-Json on the deserialized object works cleanly.
    $xmlFile = Join-Path $OutputPath "${hostname}-discovery-${dateStr}.clixml"
    $recovered = $false
    try {
        Export-Clixml -InputObject $discoveryResult -Path $xmlFile -Depth 20 -Force -ErrorAction Stop
        $xmlSize = (Get-Item $xmlFile -ErrorAction SilentlyContinue).Length
        Write-Host ("  CLI-XML saved: {0} bytes" -f $xmlSize) -ForegroundColor DarkGray

        # Re-import the deserialized object - PS strips the indexer members on roundtrip
        $reimported = Import-Clixml -Path $xmlFile -ErrorAction Stop

        # Run it through our sanitizers again for good measure
        $safeResult2 = Invoke-StrictSanitize (ConvertTo-SafeObject $reimported)

        foreach ($depth in @(20, 10, 6)) {
            try {
                $json = $safeResult2 | ConvertTo-Json -Depth $depth -ErrorAction Stop
                Write-BuddyOK "Serialized at depth $depth (via CLI-XML roundtrip)"
                break
            } catch {
                $jsonErr = "roundtrip depth $depth failed: $_"
            }
        }

        if ($json) {
            [System.IO.File]::WriteAllText($outputFile, $json, [System.Text.Encoding]::UTF8)
            Write-BuddyOK "Saved: $outputFile"
            # Keep the .clixml as an audit/forensic backup
            Write-Host "  (CLI-XML backup retained: $xmlFile)" -ForegroundColor DarkGray
            $recovered = $true
        }
    } catch {
        Write-BuddyWarn "CLI-XML recovery failed: $($_.Exception.Message)"
    }

    if (-not $recovered) {
        Write-BuddyErr "Output" "All serialization paths failed. Emergency dumps:"
        Write-Host "    $xmlFile" -ForegroundColor DarkGray
        try {
            $txtFile = Join-Path $OutputPath "discovery-error-$(Get-Date -f yyyyMMdd-HHmmss).txt"
            $discoveryResult | Out-File $txtFile -ErrorAction SilentlyContinue
            Write-Host "    $txtFile" -ForegroundColor DarkGray
        } catch { }
        Write-Host ""
        Write-Host "  Manual recovery (paste on this server):" -ForegroundColor Cyan
        Write-Host "    `$d=Import-Clixml '$xmlFile'; `$d|ConvertTo-Json -Depth 20|Set-Content '$outputFile' -Encoding utf8" -ForegroundColor DarkCyan
        exit 1
    }
}

# -- VERIFY OUTPUT FILE EXISTS AND HAS CONTENT --------------------------------
if (-not (Test-Path $outputFile)) {
    Write-Host ""
    Write-Host ("  (x_x)  OUTPUT FILE MISSING: $outputFile") -ForegroundColor Red
    Write-Host "         The discovery ran but nothing was written to disk." -ForegroundColor DarkRed
    Write-Host ""
    Write-Host "  Possible causes:" -ForegroundColor Yellow
    Write-Host "    - OutputPath does not exist or is not writable: $OutputPath" -ForegroundColor DarkGray
    Write-Host "    - Remote job returned empty/null data" -ForegroundColor DarkGray
    Write-Host "    - Hostname from remote was null (network/auth issue)" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Try:" -ForegroundColor Yellow
    Write-Host "    - Verify OutputPath exists: Test-Path '$OutputPath'" -ForegroundColor DarkGray
    Write-Host "    - Re-run with explicit path: -OutputPath C:\Temp" -ForegroundColor DarkGray
    if (-not $script:IsNonInteractive) {
        $open = Read-Host "  Open output folder in Explorer? (Y/N)"
        if ($open -match '^[Yy]') { Start-Process explorer.exe $OutputPath }
    }
    exit 1
}

$fileSize = (Get-Item $outputFile).Length
if ($fileSize -lt 100) {
    Write-Host ""
    Write-Host ("  (>_<)  Output file is suspiciously small ($fileSize bytes): $outputFile") -ForegroundColor DarkYellow
    Write-Host "         The file exists but may be empty or corrupt." -ForegroundColor DarkGray
    Write-Host "         Check it before generating a report." -ForegroundColor DarkGray
    Write-Host ""
}

# -----------------------------------------------------------------------------
# SUMMARY
# -----------------------------------------------------------------------------

$elapsed = [math]::Round(((Get-Date) - $script:StartTime).TotalSeconds, 1)
$flagCounts = @{ critical=0; warning=0; info=0 }
foreach ($f in $discoveryResult.Flags) {
    $sev = $f.Severity.ToLower()
    if ($flagCounts.ContainsKey($sev)) { $flagCounts[$sev]++ }
}

Write-Host ""
Write-Host ("=" * 72) -ForegroundColor DarkMagenta
Write-Host "  DISCOVERY COMPLETE" -ForegroundColor Magenta
Write-Host ("=" * 72) -ForegroundColor DarkMagenta
Write-Host ("  Server    : " + $discoveryResult.System.Hostname) -ForegroundColor Cyan
Write-Host ("  OS        : " + $discoveryResult.System.OSName + " (" + $discoveryResult.System.OSEOLStatus + ")") -ForegroundColor $(if ($discoveryResult.System.OSEOLStatus -eq 'EOL') { "Red" } elseif ($discoveryResult.System.OSEOLStatus -eq 'Near EOL') { "Yellow" } else { "Green" })
Write-Host ("  SQL       : " + $(if ($discoveryResult.SQL.Instances.Count -gt 0) { "$($discoveryResult.SQL.Instances.Count) instance(s)" } else { "None detected" })) -ForegroundColor Gray
Write-Host ("  Apps      : " + $discoveryResult.Apps.Count + " applications cataloged") -ForegroundColor Gray
Write-Host ("  Flags     : " + $flagCounts.critical + " critical  |  " + $flagCounts.warning + " warnings  |  " + $flagCounts.info + " info") -ForegroundColor $(if ($flagCounts.critical -gt 0) { "Red" } elseif ($flagCounts.warning -gt 0) { "Yellow" } else { "Green" })
Write-Host ("  Errors    : " + $discoveryResult.Errors.Count + " collection error(s)") -ForegroundColor $(if ($discoveryResult.Errors.Count -gt 0) { "Yellow" } else { "Gray" })
Write-Host ("  Runtime   : ${elapsed}s") -ForegroundColor DarkGray
Write-Host ("  Output    : $outputFile") -ForegroundColor White
Write-Host ("=" * 72) -ForegroundColor DarkMagenta
Write-Host ""
Write-Host "  (^_^)>  Send that JSON to Claude and ask for the HTML report." -ForegroundColor DarkCyan
Write-Host ""

if ($discoveryResult.Errors.Count -gt 0) {
    Write-Host "  Collection errors (non-fatal - partial data in JSON):" -ForegroundColor DarkYellow
    foreach ($e in $discoveryResult.Errors) {
        Write-Host ("    " + $e) -ForegroundColor DarkGray
    }
    Write-Host ""
}
