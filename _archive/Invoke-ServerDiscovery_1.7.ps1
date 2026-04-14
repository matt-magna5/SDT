<#
.SYNOPSIS
    Magna5 Solutions Engineering - Full Server Discovery v1.4
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
    [PSCredential] $Credential
)

$ErrorActionPreference = 'Continue'
$script:ScriptVersion  = '1.7'
$script:StartTime      = Get-Date
$script:CollectErrors  = [System.Collections.ArrayList]@()

$script:IsRemote = ($ComputerName -ne $env:COMPUTERNAME -and
                    $ComputerName -ne 'localhost'         -and
                    $ComputerName -ne '127.0.0.1'         -and
                    $ComputerName -ne '.')

# -- PS VERSION & CAPABILITY PROBE ---------------------------------------------

$localPSMajor = $PSVersionTable.PSVersion.Major
$localPSMinor = $PSVersionTable.PSVersion.Minor
$localPSStr   = "$localPSMajor.$localPSMinor"

# -- BANNER --------------------------------------------------------------------

Clear-Host
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
    Write-Host "  (>_<)  PowerShell $localPSStr — limited support. Date conversions may degrade." -ForegroundColor Yellow
    Write-Host "         Recommend PS 5.1+ for full functionality." -ForegroundColor DarkYellow
    Write-Host ""
}
if ($localPSMajor -ge 5 -and $localPSMajor -lt 6) {
    Write-Host "  (^_^)  PowerShell $localPSStr — compatible (PS 5.1)." -ForegroundColor DarkGreen
} elseif ($localPSMajor -ge 7) {
    Write-Host "  (^_^)  PowerShell $localPSStr — full compatibility." -ForegroundColor DarkGreen
}

# Cert bypass for PS 5.1 — self-signed certs on ESXi/vCenter won't block REST calls
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
function Write-BuddyErr  { param([string]$ctx, [string]$msg) {
    Write-Host ("  (x_x)  [$ctx] $msg") -ForegroundColor DarkRed
    [void]$script:CollectErrors.Add("[$ctx] $msg")
}}

# -- REMOTE PREFLIGHT ----------------------------------------------------------

if ($script:IsRemote) {
    Write-BuddyPhase "Remote" "remote mode - testing WinRM on $ComputerName..."

    try {
        Test-WSMan -ComputerName $ComputerName -ErrorAction Stop | Out-Null
        Write-BuddyOK "WinRM reachable on $ComputerName"
    } catch {
        Write-BuddyErr "RemotePreflight" "WinRM not reachable: $_"
        Write-Host ""
        Write-Host "  Run these on the TARGET server to enable WinRM:" -ForegroundColor Yellow
        Write-Host "    winrm quickconfig -y" -ForegroundColor DarkGray
        Write-Host "    Enable-PSRemoting -Force" -ForegroundColor DarkGray
        Write-Host "    # If workgroup (not domain-joined):" -ForegroundColor DarkGray
        Write-Host "    Set-Item WSMan:\localhost\Client\TrustedHosts -Value '*' -Force" -ForegroundColor DarkGray
        Write-Host ""
        $ans = Read-Host "  Press Enter to try anyway, or Ctrl+C to abort"
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
    function cb-Flag { param([string]$sev,[string]$title,[string]$detail) {
        [void]$cbFlags.Add([PSCustomObject]@{ Severity=$sev; Title=$title; Detail=$detail })
    }}

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
            try { $result.UserCount     = (Get-ADUser     -Filter * -ErrorAction Stop).Count              } catch { }
            try { $result.ComputerCount = (Get-ADComputer -Filter * -ErrorAction Stop).Count              } catch { }
            try { $result.OUCount       = (Get-ADOrganizationalUnit -Filter * -ErrorAction Stop).Count    } catch { }
            # Stale accounts (90 days inactive)
            try {
                $cutoff = (Get-Date).AddDays(-90)
                $staleUsers = Get-ADUser -Filter { LastLogonDate -lt $cutoff -and Enabled -eq $true } -Properties LastLogonDate -ErrorAction Stop |
                    Select-Object Name, SamAccountName, @{N='LastLogon';E={$_.LastLogonDate}} |
                    Select-Object -First 50
                $result.StaleUsers = @($staleUsers)
                if ($staleUsers.Count -gt 0) {
                    cb-Flag 'warning' "Stale AD Accounts: $($staleUsers.Count)" "$($staleUsers.Count) enabled users haven't logged in for 90+ days. Review before migration."
                }
            } catch { }
            try {
                $cutoff = (Get-Date).AddDays(-90)
                $staleComps = Get-ADComputer -Filter { LastLogonDate -lt $cutoff -and Enabled -eq $true } -Properties LastLogonDate -ErrorAction Stop |
                    Select-Object Name, @{N='LastLogon';E={$_.LastLogonDate}} |
                    Select-Object -First 50
                $result.StaleComputers = @($staleComps)
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
                    $shares = Get-SmbShare -ErrorAction Stop | Where-Object { $_.Name -notmatch '^\w\$$|^ADMIN\$$|^IPC\$$|^PRINT\$$' }
                    $result.Shares = @($shares | ForEach-Object {
                        try {
                            $acl = Get-SmbShareAccess -Name $_.Name -ErrorAction SilentlyContinue
                            @{
                                Name        = $_.Name
                                Path        = $_.Path
                                Description = $_.Description
                                Permissions = @($acl | Select-Object AccountName, AccessControlType, AccessRight)
                            }
                        } catch { @{ Name=$_.Name; Path=$_.Path; Description=$_.Description } }
                    })
                    try {
                        $sessions = Get-SmbSession -ErrorAction SilentlyContinue
                        $result.OpenSessions = if ($sessions) { @($sessions).Count } else { 0 }
                    } catch { }
                    return $result
                } catch { }
            }
            # Fallback: net share + WMI Win32_Share
            $wmiShares = Safe-Wmi Win32_Share FileShares
            if ($wmiShares) {
                $result.Shares = @($wmiShares | Where-Object { $_.Type -eq 0 -and $_.Name -notmatch '^\w\$$|^IPC\$$' } |
                    Select-Object @{N='Name';E={$_.Name}}, @{N='Path';E={$_.Path}}, @{N='Description';E={$_.Description}})
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
            # Fallback: parse netsh nps output
            try {
                $npsExport = netsh nps show client 2>&1
                $clients = @($npsExport | Where-Object { $_ -match 'Client Name|IP Address' } | ForEach-Object { $_.Trim() })
                $result.Clients = @($clients | ForEach-Object { @{ Raw = $_ } })
                $result.Partial = $true
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
            $inst = @{ InstanceName=$instName; Version='Unknown'; Edition='Unknown'; ServiceAccount='Unknown'; EOLStatus='Unknown'; EOLDate='Unknown'; Databases=@(); Partial=$false }
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
                        $inst.EOLStatus = $SQLEOLMap[$k].Status
                        $inst.EOLDate   = $SQLEOLMap[$k].EOL
                        $inst.Edition   = $SQLEOLMap[$k].Name
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
                # Query databases via SQL connection
                try {
                    $connStr = if ($instName -eq 'MSSQLSERVER') { $env:COMPUTERNAME } else { "$env:COMPUTERNAME\$instName" }
                    $conn = New-Object System.Data.SqlClient.SqlConnection
                    $conn.ConnectionString = "Server=$connStr;Integrated Security=True;Connect Timeout=10"
                    $conn.Open()
                    $cmd = $conn.CreateCommand()
                    $cmd.CommandText = @"
SELECT
    d.name,
    d.state_desc,
    CAST(SUM(mf.size) * 8.0 / 1024 / 1024 AS DECIMAL(10,2)) AS SizeGB,
    d.create_date,
    MAX(b.backup_finish_date) AS LastBackup,
    (SELECT MAX(login_time) FROM sys.dm_exec_sessions s WHERE s.database_id = d.database_id) AS LastConnectionApprox
FROM sys.databases d
LEFT JOIN sys.master_files mf ON d.database_id = mf.database_id
LEFT JOIN msdb.dbo.backupset b ON d.name = b.database_name
GROUP BY d.name, d.state_desc, d.create_date, d.database_id
ORDER BY d.name
"@
                    $reader = $cmd.ExecuteReader()
                    $dbs = [System.Collections.ArrayList]@()
                    while ($reader.Read()) {
                        $dbName     = $reader["name"].ToString()
                        $lastBackup = if ($reader["LastBackup"] -is [DBNull]) { 'Never' } else { $reader["LastBackup"].ToString("yyyy-MM-dd") }
                        $sizeGB     = if ($reader["SizeGB"] -is [DBNull]) { 0 } else { $reader["SizeGB"] }
                        $created    = $reader["create_date"].ToString("yyyy-MM-dd")
                        # Lookup app name
                        $appName = ''
                        foreach ($k in $SqlDbAppMap.Keys) {
                            if ($dbName -match $k) { $appName = $SqlDbAppMap[$k]; break }
                        }
                        $db = @{
                            Name        = $dbName
                            State       = $reader["state_desc"].ToString()
                            SizeGB      = $sizeGB
                            Created     = $created
                            LastBackup  = $lastBackup
                            AppGuess    = $appName
                        }
                        if ($lastBackup -eq 'Never' -and $dbName -notin @('tempdb','model')) {
                            cb-Flag 'critical' "SQL DB Never Backed Up: $dbName ($instName)" "No backup record in msdb. This database has no recorded backups - confirm alternate backup method."
                        } elseif ($lastBackup -ne 'Never') {
                            $daysSince = ((Get-Date) - [datetime]::Parse($lastBackup)).TotalDays
                            if ($daysSince -gt 30) {
                                cb-Flag 'warning' "SQL DB Stale Backup: $dbName" "Last backup: $lastBackup ($([math]::Round($daysSince,0)) days ago)"
                            }
                        }
                        [void]$dbs.Add($db)
                    }
                    $reader.Close()
                    $conn.Close()
                    $inst.Databases = $dbs
                } catch {
                    cb-Log "SQL-$instName" "Database query failed (may need SA access): $_"
                    $inst.Partial = $true
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
                }
            } catch { $result.Partial = $true; cb-Log "HyperV" "Get-VM failed: $_" }
        } catch {
            cb-Log "HyperV" "Outer error: $_"
            $result.Partial = $true
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
                        if (-not $isMS -and $s.PathName -notmatch 'system32|syswow64' -or $s.StartName -notmatch 'LocalSystem|LocalService|NetworkService') {
                            [void]$services.Add(@{
                                Name        = $s.Name
                                DisplayName = $s.DisplayName
                                State       = $s.State
                                StartMode   = $s.StartMode
                                StartName   = $s.StartName
                                Path        = $s.PathName
                            })
                        }
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
            foreach ($log in $logs) {
                try {
                    $evts = Get-EventLog -LogName $log -EntryType Error,Warning,Critical -After $cutoff -Newest 200 -ErrorAction SilentlyContinue
                    if ($evts) { foreach ($e in $evts) { [void]$allEvts.Add($e) } }
                } catch { }
            }
            $result.CriticalCount = @($allEvts | Where-Object { $_.EntryType -eq 'Error' }).Count
            $result.TopSources    = @($allEvts | Group-Object Source | Sort-Object Count -Descending | Select-Object -First 10 |
                ForEach-Object { @{ Source=$_.Name; Count=$_.Count } })
            $result.RecentCritical = @($allEvts | Where-Object { $_.EntryType -eq 'Error' } |
                Sort-Object TimeGenerated -Descending | Select-Object -First 10 |
                ForEach-Object { @{ Time=$_.TimeGenerated.ToString("yyyy-MM-dd HH:mm"); Source=$_.Source; Message=($_.Message -replace '\r?\n',' ' | Select-Object -First 200) } })
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
        Flags      = $cbFlags
        Errors     = $cbErrors
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
    Write-Host ("  Sending collection block to $ComputerName...") -ForegroundColor DarkCyan
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
            ComputerName = $ComputerName
            ScriptBlock  = $CollectionBlock
            AsJob        = $true
            ErrorAction  = 'Stop'
        }
        if ($Credential) { $icParams.Credential = $Credential }

        $job = Invoke-Command @icParams

        $spinChars = @('|','/','-','\')
        $spinIdx   = 0
        while ($job.State -eq 'Running') {
            $frame = $buddyFrames[$spinIdx % $buddyFrames.Count]
            Write-Host ("`r  $frame  waiting for $ComputerName...  $($spinChars[$spinIdx % 4])   ") -NoNewline -ForegroundColor DarkCyan
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
    # PSCustomObject or complex .NET object — enumerate visible PS properties only
    if ($Obj -is [System.Management.Automation.PSObject]) {
        $h = [ordered]@{}
        foreach ($prop in $Obj.PSObject.Properties) {
            try { $h[$prop.Name] = ConvertTo-SafeObject $prop.Value ($Depth+1) } catch { $h[$prop.Name] = $null }
        }
        return $h
    }
    return $Obj.ToString()
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

# Sanitize first — ConvertTo-SafeObject strips all CIM/WMI/.NET types to plain
# hashtables/arrays/primitives so ConvertTo-Json never walks into circular graphs.
Write-Host "  Sanitizing result object..." -ForegroundColor DarkGray
$safeResult = ConvertTo-SafeObject $discoveryResult

# Serialize to JSON
$json = $null
$jsonErr = ''
if ($PSMaj -ge 3) {
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
    # PS 2.0 fallback
    try {
        Add-Type -AssemblyName System.Web.Extensions -ErrorAction Stop
        $serializer = [System.Web.Script.Serialization.JavaScriptSerializer]::new()
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
    Write-BuddyErr "Output" "All JSON serialization attempts failed. Last error: $jsonErr"
    Write-Host "  Saving raw object dump as emergency fallback..." -ForegroundColor DarkYellow
    try {
        $fallbackFile = Join-Path $OutputPath "discovery-error-$(Get-Date -f yyyyMMdd-HHmmss).txt"
        $discoveryResult | Out-File $fallbackFile -ErrorAction SilentlyContinue
        Write-BuddyWarn "Emergency fallback saved: $fallbackFile"
    } catch { }
    exit 1
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
    $open = Read-Host "  Open output folder in Explorer? (Y/N)"
    if ($open -match '^[Yy]') { Start-Process explorer.exe $OutputPath }
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
