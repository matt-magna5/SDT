<#
.SYNOPSIS
    Magna5 Server Discovery - Session Launcher v2.3

.DESCRIPTION
    Entry point for multi-server discovery. Enumerates servers from your
    hypervisor (vCenter, ESXi, or Hyper-V), lets you add manual targets,
    probes every server, builds a full plan table, and asks for one explicit
    "go" before running anything.

    WinRM safety:
      - WinRM OFF on a target? Enables it via WMI, runs discovery, shuts it back down.
      - WinRM ON already? Left exactly as found.
      - Cleanup handler fires even on Ctrl+C - no server gets left with WinRM on.

    Run from a domain-joined jump box with domain admin credentials.
    Invoke-ServerDiscovery_1.4.ps1 must be in the same folder.

.EXAMPLE
    .\Start-DiscoverySession.ps1
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Continue'
$script:SessionVersion  = '2.3'
$script:SessionStart    = Get-Date
$script:WinRMRestoreMap = @{}
$script:PendingInventories = [System.Collections.ArrayList]@()
$script:SessionErrors   = [System.Collections.ArrayList]@()
$script:DomainCred      = $null
$script:vSphereSources = [System.Collections.ArrayList]@()

# -- COMPANION SOURCE PATHS ---------------------------------------------------
# Ordered list of locations to search for Invoke-ServerDiscovery when the
# local copy is missing or truncated (RDP/SMB clipboard truncation is common).
#
# Add a UNC path or corp PC share to let the launcher self-heal from one copy:
#   $env:M5_DISCOVERY_SOURCE = '\\CORPPC\C$\Tools\M5Discovery'
#   $env:M5_DISCOVERY_SOURCE = '\\fileserver\shared\M5Scripts'
#
# Set that env var permanently on your machine and you only ever need to
# transfer Start-DiscoverySession to a client  --  it fetches its companion.

$script:CompanionSources = @(
    $PSScriptRoot,
    'C:\Temp\M5Discovery',
    $env:M5_DISCOVERY_SOURCE
) | Where-Object { $_ -and $_.Trim() -ne '' }

# -- DEPENDENCY CHECK + INTEGRITY VALIDATION -----------------------------------
# Validates syntax of companion script. If local copy is missing or truncated,
# searches all CompanionSources and auto-copies the first valid one found.

function Test-ScriptIntegrity {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $false }
    $sz = (Get-Item $Path).Length
    if ($sz -lt 50000) { return $false }   # file should be ~79KB - truncated if smaller
    $errs = $null
    $null = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$errs)
    return ($errs.Count -eq 0)
}

$invokeScriptName = 'Invoke-ServerDiscovery_2.0.ps1'
$DiscoveryScript  = Join-Path $PSScriptRoot $invokeScriptName

if (Test-ScriptIntegrity $DiscoveryScript) {
    Write-Host "  (^_^)  Companion script OK." -ForegroundColor DarkGreen
} else {
    $localSz = if (Test-Path $DiscoveryScript) { (Get-Item $DiscoveryScript).Length } else { 0 }
    Write-Host ""
    Write-Host "  (>_<)  $invokeScriptName is missing or corrupt (${localSz} bytes)." -ForegroundColor Yellow
    Write-Host "         Searching source locations..." -ForegroundColor DarkGray

    $found = $null
    foreach ($src in $script:CompanionSources) {
        $candidate = Join-Path $src $invokeScriptName
        Write-Host "         Checking: $candidate" -ForegroundColor DarkGray
        if (Test-ScriptIntegrity $candidate) { $found = $candidate; break }
    }

    if ($found) {
        Write-Host "  (^_^)  Found valid copy at: $found" -ForegroundColor Green
        try {
            Copy-Item -Path $found -Destination (Join-Path $PSScriptRoot $invokeScriptName) -Force -ErrorAction Stop
            $DiscoveryScript = Join-Path $PSScriptRoot $invokeScriptName
            Write-Host "  (^_^)  Auto-copied to $PSScriptRoot  --  will be local next run." -ForegroundColor DarkGreen
        } catch {
            $DiscoveryScript = $found
            Write-Host "         (Could not write to script folder  --  using source path directly.)" -ForegroundColor DarkGray
        }
    } else {
        Write-Host ""
        Write-Host "  (x_x)  No valid copy of $invokeScriptName found in any location." -ForegroundColor Red
        Write-Host ""
        Write-Host "  Locations checked:" -ForegroundColor DarkGray
        foreach ($src in $script:CompanionSources) { Write-Host "    - $src" -ForegroundColor DarkGray }
        Write-Host ""
        Write-Host "  Fix options:" -ForegroundColor Yellow
        Write-Host "    1.  Copy $invokeScriptName to any location above." -ForegroundColor DarkGray
        Write-Host "    2.  Set `$env:M5_DISCOVERY_SOURCE to a folder that has the file:" -ForegroundColor DarkGray
        Write-Host "           `$env:M5_DISCOVERY_SOURCE = '\\CORPPC\C`$\Tools\M5Discovery'" -ForegroundColor DarkGray
        Write-Host "        Then re-run  --  this script will pull and cache it automatically." -ForegroundColor DarkGray
        exit 1
    }
}

# -- SUGGESTED HYPERVISOR SCAN (AD + DNS) ------------------------------------
function Get-SuggestedHypervisors {
    $hvPatterns = @('*VH*','*HV*','*ESX*','*ESXI*','*VCENTER*','*HYPERV*','*HYP*',
                    '*VMWARE*','*VMW*','*NUTANIX*','*NTX*','*PRISM*',
                    '*XEN*','*PROXMOX*','*PVE*','*VHOST*','*VIRT*')
    $found = [System.Collections.ArrayList]@()

    # Try AD via ActiveDirectory module — run in job with 10s timeout to avoid hangs
    $adComps = $null
    try {
        $adJob = Start-Job -ScriptBlock {
            Import-Module ActiveDirectory -ErrorAction Stop
            Get-ADComputer -Filter * -Properties Name,IPv4Address,Description,OperatingSystem -ErrorAction Stop
        }
        if (Wait-Job $adJob -Timeout 10) {
            $adComps = Receive-Job $adJob
        } else {
            Stop-Job $adJob
        }
        Remove-Job $adJob -Force -ErrorAction SilentlyContinue
    } catch { }

    if ($adComps) {
        foreach ($c in $adComps) {
            $n = $c.Name.ToUpper()
            if ($hvPatterns | Where-Object { $n -like $_ }) {
                if (-not ($found | Where-Object { $_.Name -ieq $c.Name })) {
                    [void]$found.Add([PSCustomObject]@{
                        Name = $c.Name; IP = $c.IPv4Address; Source = 'AD'
                        Description = $c.Description; OS = $c.OperatingSystem
                    })
                }
            }
        }
    } else {
        # Fallback: ADSI/LDAP with explicit timeout
        try {
            $root    = [ADSI]'LDAP://RootDSE'
            $base    = $root.defaultNamingContext
            $filter  = '(&(objectClass=computer)(|(name=*VH*)(name=*HV*)(name=*ESX*)(name=*VCENTER*)(name=*HYPERV*)(name=*VMWARE*)(name=*VMW*)(name=*NUTANIX*)(name=*NTX*)(name=*PRISM*)(name=*XEN*)(name=*PROXMOX*)(name=*PVE*)(name=*VHOST*)(name=*VIRT*)))'
            $srch    = New-Object DirectoryServices.DirectorySearcher([ADSI]"LDAP://$base", $filter)
            $srch.ClientTimeout  = [TimeSpan]::FromSeconds(8)
            $srch.ServerTimeLimit = [TimeSpan]::FromSeconds(8)
            @('name','dNSHostName','description') | ForEach-Object { [void]$srch.PropertiesToLoad.Add($_) }
            $srch.FindAll() | ForEach-Object {
                $n   = if ($_.Properties['name'].Count)        { $_.Properties['name'][0] }        else { '' }
                $dns = if ($_.Properties['dNSHostName'].Count)  { $_.Properties['dNSHostName'][0] }  else { $n }
                $dsc = if ($_.Properties['description'].Count)  { $_.Properties['description'][0] }  else { '' }
                if ($n -and -not ($found | Where-Object { $_.Name -ieq $n })) {
                    [void]$found.Add([PSCustomObject]@{ Name=$n; IP=$dns; Source='AD (LDAP)'; Description=$dsc; OS='' })
                }
            }
        } catch { }
    }

    # DNS forward-lookup sweep — fire all candidates in parallel, collect within 4s total
    if ($found.Count -eq 0) {
        $prefixes = @('VH','HV','ESX','ESXI','VCENTER','VC','HYPERV',
                      'VMW','NTX','PRISM','XEN','PVE','VHOST')
        $candidates = @()
        foreach ($p in $prefixes) {
            foreach ($i in 1..5) {
                foreach ($fmt in @("$p$i","${p}0$i","${p}-$i","${p}_$i")) {
                    $candidates += $fmt
                }
            }
        }
        # Start all async lookups simultaneously
        $handles = [ordered]@{}
        foreach ($c in $candidates) {
            try { $handles[$c] = [System.Net.Dns]::BeginGetHostEntry($c, $null, $null) } catch { }
        }
        # Collect results within a 4s window
        $deadline = [DateTime]::Now.AddSeconds(4)
        foreach ($c in $handles.Keys) {
            $remaining = [int]($deadline - [DateTime]::Now).TotalMilliseconds
            if ($remaining -le 0) { break }
            try {
                if ($handles[$c].AsyncWaitHandle.WaitOne([Math]::Max($remaining, 1))) {
                    $r = [System.Net.Dns]::EndGetHostEntry($handles[$c])
                    if (-not ($found | Where-Object { $_.Name -ieq $c })) {
                        [void]$found.Add([PSCustomObject]@{
                            Name=$c; IP=($r.AddressList[0].ToString())
                            Source='DNS'; Description=''; OS=''
                        })
                    }
                }
            } catch { }
        }
    }
    return ,$found   # comma forces array return
}

# -- SUGGESTED SERVER SCAN (AD - by OS type, bare metal path) -----------------
function Get-SuggestedServers {
    $found = [System.Collections.ArrayList]@()

    # Try AD module first
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        $adComps = Get-ADComputer -Filter {OperatingSystem -like "*Server*"} `
            -Properties Name,IPv4Address,Description,OperatingSystem -ErrorAction Stop
        foreach ($c in $adComps) {
            if (-not ($found | Where-Object { $_.Name -ieq $c.Name })) {
                [void]$found.Add([PSCustomObject]@{
                    Name=$c.Name; IP=$c.IPv4Address; Source='AD'
                    Description=$c.Description; OS=$c.OperatingSystem
                })
            }
        }
    } catch {
        # ADSI/LDAP fallback (no AD module needed)
        try {
            $root   = [ADSI]'LDAP://RootDSE'
            $base   = $root.defaultNamingContext
            $filter = '(&(objectClass=computer)(operatingSystem=*Server*))'
            $srch   = New-Object DirectoryServices.DirectorySearcher([ADSI]"LDAP://$base", $filter)
            @('name','dNSHostName','description','operatingSystem') | ForEach-Object { [void]$srch.PropertiesToLoad.Add($_) }
            $srch.FindAll() | ForEach-Object {
                $n   = if ($_.Properties['name'].Count)            { $_.Properties['name'][0] }            else { '' }
                $dns = if ($_.Properties['dNSHostName'].Count)     { $_.Properties['dNSHostName'][0] }     else { $n }
                $os  = if ($_.Properties['operatingSystem'].Count) { $_.Properties['operatingSystem'][0] } else { '' }
                $dsc = if ($_.Properties['description'].Count)     { $_.Properties['description'][0] }     else { '' }
                if ($n -and -not ($found | Where-Object { $_.Name -ieq $n })) {
                    [void]$found.Add([PSCustomObject]@{ Name=$n; IP=$dns; Source='AD (LDAP)'; Description=$dsc; OS=$os })
                }
            }
        } catch { }
    }
    return ,$found
}

# -- CERT BYPASS (PS 5.1 - self-signed certs on ESXi/vCenter) -----------------

$PSMaj = $PSVersionTable.PSVersion.Major
if ($PSMaj -lt 6) {
    try {
        Add-Type -TypeDefinition @"
using System.Net; using System.Security.Cryptography.X509Certificates;
public class M5TrustAll : ICertificatePolicy {
    public bool CheckValidationResult(ServicePoint s, X509Certificate c, WebRequest r, int p) { return true; }
}
"@ -ErrorAction SilentlyContinue
        [System.Net.ServicePointManager]::CertificatePolicy   = New-Object M5TrustAll
        [System.Net.ServicePointManager]::SecurityProtocol    = [System.Net.SecurityProtocolType]::Tls12
    } catch { }
}


# -- LINUX OS DETECTION --------------------------------------------------------
# Uses the GuestOS hint from vCenter/ESXi. These strings come back from the
# vSphere REST API guest_OS field (e.g. "CENTOS 64", "PHOTON OS 64", etc.)
function Test-IsLinux {
    param([string]$GuestOSHint)
    if (-not $GuestOSHint -or $GuestOSHint -eq 'Unknown') { return $false }
    $linuxPatterns = @('LINUX','CENTOS','UBUNTU','RHEL','DEBIAN','FEDORA',
                       'SUSE','ORACLE LINUX','PHOTON','COREOS','ROCKY',
                       'ALMA','AMAZON LINUX','OTHER 64','OTHER LINUX',
                       'FREEBSD','SOLARIS','DARWIN','UNIX','OPENBSD',
                       'OTHER (64','OTHER GUEST','VMWARE PHOTON',
                       'OPENSOLARIS','NETBSD','MANJARO','ARCH LINUX')
    foreach ($p in $linuxPatterns) {
        if ($GuestOSHint.ToUpper().Contains($p)) { return $true }
    }
    return $false
}

# -- LINUX SSH DISCOVERY -------------------------------------------------------

function Parse-LinuxOutput {
    param([string]$Raw, [string]$IP, [string]$GuestOS)

    $sec = @{}; $cur = $null; $buf = [System.Collections.ArrayList]@()
    foreach ($ln in ($Raw -split "`r?`n")) {
        if ($ln -match '^---([A-Z]+)---$') {
            if ($cur) { $sec[$cur] = @($buf) }
            $cur = $matches[1]; $buf.Clear()
        } elseif ($cur -and $ln.Trim()) {
            [void]$buf.Add($ln.TrimEnd())
        }
    }
    if ($cur) { $sec[$cur] = @($buf) }

    $hostname = if ($sec['HOSTNAME']) { $sec['HOSTNAME'][0].Trim() } else { $IP }

    $os = @{ PrettyName='Unknown'; ID='unknown'; VersionID=''; Kernel=''; Architecture='' }
    foreach ($ln in ($sec['OSRELEASE'] -as [array])) {
        if ($ln -match '^PRETTY_NAME="?([^"]+)"?')  { $os.PrettyName  = $matches[1] }
        if ($ln -match '^ID="?([^"]+)"?')            { $os.ID          = $matches[1] }
        if ($ln -match '^VERSION_ID="?([^"]+)"?')    { $os.VersionID   = $matches[1] }
    }
    $uname = if ($sec['UNAME']) { $sec['UNAME'][0] } else { '' }
    if ($uname -match '\S+\s+(\S+)\s+(\S+)') { $os.Kernel = $matches[1]; $os.Architecture = $matches[2] }

    $cores = 1; $cpuModel = 'Unknown'
    foreach ($ln in ($sec['CPU'] -as [array])) {
        if ($ln -match 'CORES=(\d+)')                { $cores    = [int]$matches[1] }
        if ($ln -match 'Model name\s*:\s*(.+)')      { $cpuModel = $matches[1].Trim() }
    }

    $memTotal = 0; $memUsed = 0; $memFree = 0
    foreach ($ln in ($sec['MEMORY'] -as [array])) {
        if ($ln -match '^Mem:\s+(\d+)\s+(\d+)\s+(\d+)') {
            $memTotal = [int]$matches[1]; $memUsed = [int]$matches[2]; $memFree = [int]$matches[3]
        }
    }

    $disks = [System.Collections.ArrayList]@()
    foreach ($ln in ($sec['DISKS'] -as [array])) {
        if ($ln -match '(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\d+)%\s+(\S+)') {
            [void]$disks.Add([PSCustomObject]@{
                Source=$matches[1]; Size=$matches[2]; Used=$matches[3]
                Free=$matches[4]; UsePct=[int]$matches[5]; Mount=$matches[6]
            })
        }
    }

    $ifMap = [ordered]@{}
    foreach ($ln in ($sec['NETWORK'] -as [array])) {
        if ($ln -match '^\d+:\s+(\S+)\s+inet\s+(\S+)') {
            $iface = ($matches[1] -replace '@.*','').Trim()
            if (-not $ifMap[$iface]) { $ifMap[$iface] = [System.Collections.ArrayList]@() }
            [void]$ifMap[$iface].Add($matches[2])
        } elseif ($ln -match 'inet\s+(\d+\.\d+\.\d+\.\d+)') {
            if (-not $ifMap['eth0']) { $ifMap['eth0'] = [System.Collections.ArrayList]@() }
            [void]$ifMap['eth0'].Add($matches[1])
        }
    }
    $netList = @($ifMap.GetEnumerator() | ForEach-Object {
        [PSCustomObject]@{ Interface=$_.Key; Addresses=@($_.Value) }
    })

    $svcs = @(($sec['SERVICES'] -as [array]) | Where-Object { $_ -and $_ -ne 'none' } |
              ForEach-Object { [PSCustomObject]@{ Name=$_.Trim() } })

    return [PSCustomObject]@{
        _type       = 'LinuxDiscovery'
        CollectedAt = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')
        Hostname    = $hostname
        IP          = $IP
        GuestOS     = $GuestOS
        OS          = $os
        CPU         = [PSCustomObject]@{ Cores=$cores; ModelName=$cpuModel }
        Memory      = [PSCustomObject]@{ TotalMB=$memTotal; UsedMB=$memUsed; FreeMB=$memFree }
        Disks       = @($disks)
        Network     = $netList
        Services    = $svcs
    }
}

function Invoke-LinuxDiscovery {
    param(
        [PSCustomObject]$Row,
        [string]$PlinkPath,
        [string]$OutputFolder,
        [string]$DateStr
    )

    $target = $Row.Address
    $name   = $Row.DisplayName
    $guest  = if ($Row.ResolvedOS -and $Row.ResolvedOS -ne 'Unknown') { $Row.ResolvedOS } else { 'Linux' }

    Write-Host ""
    Write-Host ("  [{0}]  {1}  ({2})" -f $name, $target, $guest) -ForegroundColor Cyan
    $sshUser = (Read-Host "    SSH Username [root]").Trim()
    if (-not $sshUser) { $sshUser = 'root' }
    $sshPassSec = Read-Host "    SSH Password" -AsSecureString
    $sshPass    = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sshPassSec))

    # Bash discovery command — section-delimited, single line
    $bashCmd = "echo '---HOSTNAME---'; hostname 2>/dev/null; echo '---OSRELEASE---'; cat /etc/os-release 2>/dev/null || echo 'ID=unknown'; echo '---UNAME---'; uname -srm 2>/dev/null; echo '---CPU---'; echo CORES=\$(nproc 2>/dev/null || grep -c ^processor /proc/cpuinfo 2>/dev/null || echo 1); lscpu 2>/dev/null | grep 'Model name' || echo 'Model name: Unknown'; echo '---MEMORY---'; free -m 2>/dev/null; echo '---DISKS---'; df -h 2>/dev/null | grep -v tmpfs | grep -v udev | grep -v overlay | grep -v '^Filesystem'; echo '---NETWORK---'; ip -o addr show 2>/dev/null | grep ' inet ' || ifconfig 2>/dev/null | grep 'inet '; echo '---SERVICES---'; systemctl list-units --type=service --state=running --no-pager --no-legend 2>/dev/null | awk '{print \$1}' | head -40; echo '---DONE---'"

    Write-Host "    Connecting..." -NoNewline -ForegroundColor DarkGray

    try {
        # Try with -batch first (works if host key cached in PuTTY registry)
        $output = & $PlinkPath -ssh -l $sshUser -pw $sshPass -batch $target $bashCmd 2>&1
        $outStr = $output -join "`n"

        if ($outStr -notmatch '---DONE---') {
            # Host key not cached — pipe 'y' to accept it, then retry
            $output = "y`n" | & $PlinkPath -ssh -l $sshUser -pw $sshPass $target $bashCmd 2>&1
            $outStr = $output -join "`n"
        }

        if ($outStr -notmatch '---DONE---') {
            Write-Host " FAILED" -ForegroundColor Red
            Write-Host "    Check: credentials, SSH enabled, firewall allows port 22" -ForegroundColor DarkGray
            return $null
        }

        Write-Host " OK" -ForegroundColor Green
        $result  = Parse-LinuxOutput -Raw $outStr -IP $target -GuestOS $guest
        $outFile = Join-Path $OutputFolder "$name-linux-$DateStr.json"
        $result | ConvertTo-Json -Depth 6 | Out-File $outFile -Encoding UTF8 -Force
        B-OK "Saved: $(Split-Path $outFile -Leaf)"
        return $outFile

    } catch {
        Write-Host " ERROR: $_" -ForegroundColor Red
        return $null
    } finally {
        if ($sshPass) { $sshPass = [string]::Empty }
    }
}

# -- PS VERSION COMPAT CHECK ---------------------------------------------------

if ($PSMaj -lt 3) {
    Write-Host ""
    Write-Host "  (x_x)  PowerShell $($PSVersionTable.PSVersion) is not supported." -ForegroundColor Red
    Write-Host "         Minimum: PS 3.0  |  Recommended: PS 5.1 or PS 7+" -ForegroundColor DarkRed
    Write-Host "         Upgrade PowerShell on this machine before running discovery." -ForegroundColor DarkGray
    exit 1
}
if ($PSMaj -eq 3 -or $PSMaj -eq 4) {
    Write-Host ""
    Write-Host "  (>_<)  PowerShell $($PSVersionTable.PSVersion)  --  limited compatibility." -ForegroundColor Yellow
    Write-Host "         Some WMI method calls may degrade. Recommend PS 5.1+." -ForegroundColor DarkYellow
    Write-Host ""
}
if ($PSMaj -ge 7) {
    Write-Host "  (^_^)  PowerShell $($PSVersionTable.PSVersion)  --  full CIM/WinRM compatibility." -ForegroundColor DarkGreen
} elseif ($PSMaj -ge 5) {
    Write-Host "  (^_^)  PowerShell $($PSVersionTable.PSVersion)  --  compatible." -ForegroundColor DarkGreen
}

# -- BUDDY ---------------------------------------------------------------------

$buddyFrames = @(
    "(^_^) ","(^_^)>","(o_o) ","(o_o)>",
    "(-_-) ","(>_<) ","(*_*) ","(^_-) ",
    "(._.) ","(T_T) ","(^o^) ","(x_x) "
)
function B-Line { param([string]$m,[string]$c='DarkCyan')
    Write-Host ("  " + $buddyFrames[(Get-Random -Max $buddyFrames.Count)] + "  $m") -ForegroundColor $c }
function B-OK   { param([string]$m) Write-Host "  (^_^)  $m" -ForegroundColor DarkGreen  }
function B-Warn { param([string]$m) Write-Host "  (>_<)  $m" -ForegroundColor DarkYellow }
function B-Err  { param([string]$m) Write-Host "  (x_x)  $m" -ForegroundColor Red; [void]$script:SessionErrors.Add($m) }
function Write-Phase { param([string]$t)
    Write-Host ""
    Write-Host ("  -- " + $t.ToUpper() + " " + ("-" * [Math]::Max(1, 56 - $t.Length))) -ForegroundColor DarkMagenta }
function Write-Divider { Write-Host ("  " + ("-" * 68)) -ForegroundColor DarkGray }
# Null-safe Read-Host wrapper - Read-Host returns $null when stdin is redirected
# or user hits Ctrl+C; calling .Trim() on $null crashes. This prevents that.
function Read-Safe { param([string]$prompt) $r = Read-Host $prompt; if ($null -eq $r) { return '' }; $r.Trim() }

# -- BANNER --------------------------------------------------------------------

Clear-Host
Write-Host ""
Write-Host "   __  __    _    ____  _   _    _    ____  " -ForegroundColor Magenta
Write-Host "  |  \/  |  / \  / ___|| \ | |  / \  | ___|" -ForegroundColor Magenta
Write-Host "  | |\/| | / _ \| |  _ |  \| | / _ \ |___ \" -ForegroundColor Magenta
Write-Host "  | |  | |/ ___ \ |_| || |\  |/ ___ \ ___) |" -ForegroundColor Magenta
Write-Host "  |_|  |_/_/   \_\____||_| \_/_/   \_\____/ " -ForegroundColor Magenta
Write-Host "                                       v$script:SessionVersion" -ForegroundColor DarkMagenta
Write-Host ""
Write-Host ("=" * 72) -ForegroundColor DarkMagenta
Write-Host "    SERVER DISCOVERY  --  SESSION LAUNCHER" -ForegroundColor Magenta
Write-Host ("=" * 72) -ForegroundColor DarkMagenta
Write-Host ("  Started  : " + $script:SessionStart.ToString("yyyy-MM-dd HH:mm:ss") + "  |  Host: $env:COMPUTERNAME") -ForegroundColor Gray
Write-Host ("  PS Ver   : " + $PSVersionTable.PSVersion.ToString()) -ForegroundColor Gray
Write-Host ("=" * 72) -ForegroundColor DarkMagenta
Write-Host ""

# -----------------------------------------------------------------------------
# PORTABLE PYTHON — AUTO-SETUP
# -----------------------------------------------------------------------------

$portablePy     = Join-Path $PSScriptRoot 'python\python.exe'
$setupScript    = Join-Path $PSScriptRoot 'Get-PortablePython.ps1'

if (-not (Test-Path $portablePy)) {
    $hasSysPython = $false
    foreach ($pyCandidate in @('python','python3','py')) {
        try {
            $pyOut = & $pyCandidate --version 2>&1
            if ($pyOut -match 'Python [23]') { $hasSysPython = $true; break }
        } catch { }
    }

    if (-not $hasSysPython) {
        if (Test-Path $setupScript) {
            Write-Host "  Portable Python not found -- running setup..." -ForegroundColor Cyan
            try {
                & $setupScript
                if (Test-Path $portablePy) {
                    Write-Host ""
                } else {
                    Write-Host "  Setup completed but python.exe not found -- report will need manual generation." -ForegroundColor Yellow
                    Write-Host ""
                }
            } catch {
                Write-Host "  Python setup failed: $_ -- report will need manual generation." -ForegroundColor Yellow
                Write-Host ""
            }
        } else {
            Write-Host "  Python not found and Get-PortablePython.ps1 is missing." -ForegroundColor Yellow
            Write-Host "  Discovery will run normally -- HTML report will need manual generation." -ForegroundColor DarkGray
            Write-Host ""
        }
    }
}

# -----------------------------------------------------------------------------
# HYPERVISOR HELPERS
# -----------------------------------------------------------------------------

# -- WMI / CIM COMPAT HELPERS -------------------------------------------------
# PS3+  : Get-CimInstance + Invoke-CimMethod (modern, preferred)
# PS2   : Get-WmiObject + direct .Method() calls (legacy fallback only)
# CimInstance objects have NO direct methods; always use Invoke-WmiOrCimMethod.

function Get-WmiOrCim {
    # -UseDCOM forces DCOM transport (CimSession with Protocol=Dcom).
    # Use this for pre-WinRM probes (Test-WMIAccess, WinRM state checks, etc.)
    # where WinRM may not yet be running on the target.
    # Default (no -UseDCOM) uses WS-MAN / WinRM transport.
    param(
        [string]$Class,
        [string]$Namespace    = 'root\cimv2',
        [string]$Filter       = '',
        [string]$ComputerName = '',
        [PSCredential]$Cred   = $null,
        [string]$EA           = 'Stop',
        [switch]$UseDCOM
    )
    if ($PSMaj -ge 3) {
        $p = @{ ClassName=$Class; Namespace=$Namespace; ErrorAction=$EA }
        if ($Filter) { $p.Filter = $Filter }
        if ($ComputerName) {
            if ($UseDCOM) {
                # DCOM transport  --  works without WinRM, same as legacy Get-WmiObject
                $sessOpt  = New-CimSessionOption -Protocol Dcom
                $sessArgs = @{ ComputerName=$ComputerName; SessionOption=$sessOpt; ErrorAction=$EA }
                if ($Cred) { $sessArgs.Credential = $Cred }
                $cimSess  = New-CimSession @sessArgs
                $p.CimSession = $cimSess
            } else {
                $p.ComputerName = $ComputerName
                if ($Cred) { $p.Credential = $Cred }
            }
        }
        Get-CimInstance @p
    } else {
        $p = @{ Class=$Class; Namespace=$Namespace; ErrorAction=$EA }
        if ($Filter)       { $p.Filter       = $Filter }
        if ($ComputerName) { $p.ComputerName = $ComputerName }
        if ($Cred)         { $p.Credential   = $Cred }
        Get-WmiObject @p
    }
}

function Invoke-WmiOrCimMethod {
    # Instance method : -Instance (from Get-WmiOrCim), -MethodName, -Arguments
    # Static/class    : -ClassName, -ComputerName, -MethodName, -Arguments
    # -UseDCOM        : force DCOM transport for static/class calls (pre-WinRM)
    param(
        [string]$MethodName,
        [hashtable]$Arguments  = @{},
        [object]$Instance      = $null,
        [string]$ClassName     = '',
        [string]$ComputerName  = '',
        [PSCredential]$Cred    = $null,
        [string]$Namespace     = 'root\cimv2',
        [switch]$UseDCOM
    )
    if ($PSMaj -ge 3) {
        if ($Instance) {
            # Instance already has transport baked in via its CimSession
            Invoke-CimMethod -InputObject $Instance -MethodName $MethodName -Arguments $Arguments -ErrorAction Stop
        } else {
            if ($UseDCOM -and $ComputerName) {
                $sessOpt  = New-CimSessionOption -Protocol Dcom
                $sessArgs = @{ ComputerName=$ComputerName; SessionOption=$sessOpt; ErrorAction='Stop' }
                if ($Cred) { $sessArgs.Credential = $Cred }
                $cimSess = New-CimSession @sessArgs
                Invoke-CimMethod -ClassName $ClassName -MethodName $MethodName -Arguments $Arguments `
                                 -Namespace $Namespace -CimSession $cimSess -ErrorAction Stop
            } else {
                $p = @{ ClassName=$ClassName; MethodName=$MethodName; Arguments=$Arguments; Namespace=$Namespace; ErrorAction='Stop' }
                if ($ComputerName) { $p.ComputerName = $ComputerName }
                if ($Cred)         { $p.Credential   = $Cred }
                Invoke-CimMethod @p
            }
        }
    } else {
        # PS2 WMI fallback  --  direct method invocation on WmiObject
        if ($Instance) {
            $Instance.InvokeMethod($MethodName, @($Arguments.Values))
        } else {
            $ns = if ($Namespace) { $Namespace } else { 'root\cimv2' }
            ([wmiclass]"\\$ComputerName\${ns}:${ClassName}").InvokeMethod($MethodName, @($Arguments.Values))
        }
    }
}

# -- vSphere REST API ----------------------------------------------------------

function Invoke-VSphereRest {
    param([string]$Uri, [string]$Method='GET', [hashtable]$Headers, [object]$Body=$null)
    try {
        $p = @{ Uri=$Uri; Method=$Method; Headers=$Headers; ErrorAction='Stop' }
        if ($Body)       { $p.Body        = ($Body | ConvertTo-Json); $p.ContentType = 'application/json' }
        if ($PSMaj -ge 6) { $p.SkipCertificateCheck = $true }
        Invoke-RestMethod @p
    } catch { throw $_ }
}

function Connect-VSphere {
    param([string]$Server, [PSCredential]$Cred)
    $user = $Cred.UserName
    $pass = $Cred.GetNetworkCredential().Password
    $b64  = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("${user}:${pass}"))
    $authHeader = @{ Authorization = "Basic $b64" }
    # Try vSphere API v7+ first, then v6.x fallback
    $lastErr = ''
    foreach ($apiPath in @('/api/session', '/rest/com/vmware/cis/session')) {
        try {
            $uri   = "https://$Server$apiPath"
            $token = Invoke-VSphereRest -Uri $uri -Method POST -Headers $authHeader
            $apiVer = if ($apiPath -match '^/api') { 'v7' } else { 'v6' }
            B-OK "Connected to vSphere ($apiVer API) at $Server"
            return @{ OK=$true; Token=$token; Server=$Server; APIVer=$apiVer }
        } catch { $lastErr = $_.Exception.Message }
    }
    return @{ OK=$false; Error="Could not connect to vSphere REST API on $Server — $lastErr" }
}

function Get-VSphereVMs {
    param([hashtable]$Conn)
    $tokenHeader = if ($Conn.APIVer -eq 'v7') {
        @{ 'vmware-api-session-id' = $Conn.Token }
    } else {
        @{ 'vmware-api-session-id' = $Conn.Token.value }
    }
    $vmListUri = if ($Conn.APIVer -eq 'v7') {
        "https://$($Conn.Server)/api/vcenter/vm"
    } else {
        "https://$($Conn.Server)/rest/vcenter/vm"
    }
    try {
        $raw = Invoke-VSphereRest -Uri $vmListUri -Headers $tokenHeader
        $vms = if ($Conn.APIVer -eq 'v7') { $raw } else { $raw.value }
        $list = [System.Collections.ArrayList]@()
        foreach ($vm in $vms) {
            $name       = $vm.name
            $vmId       = $vm.vm
            $powerState = $vm.power_state
            $guestOS    = $vm.guest_OS -replace '_',' '
            # Try to get IP via guest networking (requires VMware Tools running)
            $ip = ''
            try {
                $netUri = if ($Conn.APIVer -eq 'v7') {
                    "https://$($Conn.Server)/api/vcenter/vm/$vmId/guest/networking/interfaces"
                } else {
                    "https://$($Conn.Server)/rest/vcenter/vm/$vmId/guest/networking/interfaces"
                }
                $netRaw  = Invoke-VSphereRest -Uri $netUri -Headers $tokenHeader
                $netData = if ($Conn.APIVer -eq 'v7') { $netRaw } else { $netRaw.value }
                $ips     = $netData | ForEach-Object {
                    $_.ip.ip_addresses | Where-Object { $_.prefix_length -lt 33 -and $_.ip_address -notmatch ':' }
                } | Select-Object -ExpandProperty ip_address
                $ip = ($ips | Select-Object -First 2) -join ', '
            } catch { }
            [void]$list.Add([PSCustomObject]@{
                Name       = $name
                IP         = $ip
                PowerState = $powerState
                GuestOS    = $guestOS
                Source     = 'vCenter'
                VMID       = $vmId
            })
        }
        return $list
    } catch {
        B-Err "Failed to enumerate VMs from vCenter: $_"
        return @()
    }
}

function Disconnect-VSphere {
    param([hashtable]$Conn)
    try {
        $tokenHeader = if ($Conn.APIVer -eq 'v7') {
            @{ 'vmware-api-session-id' = $Conn.Token }
        } else {
            @{ 'vmware-api-session-id' = $Conn.Token.value }
        }
        $uri = "https://$($Conn.Server)/" + $(if ($Conn.APIVer -eq 'v7') { 'api/session' } else { 'rest/com/vmware/cis/session' })
        Invoke-VSphereRest -Uri $uri -Method DELETE -Headers $tokenHeader | Out-Null
    } catch { }
}


function Get-VSphereInventory {
    # Deep vCenter/ESXi inventory  --  VMs with full hardware, datastores, hosts.
    # Called after Get-VSphereVMs; uses the same open $Conn.
    param([hashtable]$Conn, [array]$VMList)
    B-Line "  Collecting vSphere inventory from $($Conn.Server) ($($VMList.Count) VMs)..."
    $hdr = if ($Conn.APIVer -eq 'v7') {
        @{ 'vmware-api-session-id' = $Conn.Token }
    } else {
        @{ 'vmware-api-session-id' = $Conn.Token.value }
    }
    $base = "https://$($Conn.Server)"
    $v7   = ($Conn.APIVer -eq 'v7')

    # Datastores
    $datastores = @()
    try {
        $dsUri = "$base/" + $(if ($v7) { 'api/vcenter/datastore' } else { 'rest/vcenter/datastore' })
        $dsRaw = Invoke-VSphereRest -Uri $dsUri -Headers $hdr
        $dsData = if ($v7) { $dsRaw } else { $dsRaw.value }
        $datastores = @($dsData | ForEach-Object {
            [PSCustomObject]@{
                Name       = $_.name
                Type       = $_.type
                CapacityGB = if ($_.capacity)   { [math]::Round($_.capacity/1GB, 1)    } else { $null }
                FreeGB     = if ($_.free_space) { [math]::Round($_.free_space/1GB, 1)  } else { $null }
            }
        })
    } catch { B-Warn "  Datastores unavailable: $_" }

    # Hosts (vCenter only  --  ESXi returns itself)
    $esxHosts = @()
    try {
        $hUri  = "$base/" + $(if ($v7) { 'api/vcenter/host' } else { 'rest/vcenter/host' })
        $hRaw  = Invoke-VSphereRest -Uri $hUri -Headers $hdr
        $hData = if ($v7) { $hRaw } else { $hRaw.value }
        $esxHosts = @($hData | ForEach-Object {
            [PSCustomObject]@{
                Name       = $_.name
                State      = $_.connection_state
                PowerState = $_.power_state
                MOID       = $_.host
            }
        })
    } catch { B-Warn "  Hosts endpoint unavailable: $_" }

    # Per-VM detail
    $vmDetails = [System.Collections.ArrayList]@()
    $tot = $VMList.Count; $done = 0
    foreach ($vm in $VMList) {
        $done++
        Write-Host ("`r    VM details $done/$tot  ") -NoNewline -ForegroundColor DarkCyan
        $det = $null
        try {
            $vUri = "$base/" + $(if ($v7) { "api/vcenter/vm/$($vm.VMID)" } else { "rest/vcenter/vm/$($vm.VMID)" })
            $vRaw = Invoke-VSphereRest -Uri $vUri -Headers $hdr
            $det  = if ($v7) { $vRaw } else { $vRaw.value }
        } catch { }
        $disks = @()
        $nics  = @()
        $cpu   = $null; $ramGB = $null
        if ($det) {
            $cpu   = $det.cpu.count
            $ramGB = if ($det.memory.size_MiB) { [math]::Round($det.memory.size_MiB/1024, 2) } else { $null }
            if ($det.disks.PSObject.Properties) {
                $disks = @($det.disks.PSObject.Properties | ForEach-Object {
                    [PSCustomObject]@{
                        Label      = $_.Value.label
                        CapacityGB = if ($_.Value.capacity) { [math]::Round($_.Value.capacity/1GB,1) } else { $null }
                        BackingType= $_.Value.backing.type
                    }
                })
            }
            if ($det.nics.PSObject.Properties) {
                $nics = @($det.nics.PSObject.Properties | ForEach-Object {
                    [PSCustomObject]@{
                        Label = $_.Value.label
                        MAC   = $_.Value.mac_address
                        Type  = $_.Value.type
                    }
                })
            }
        }
        [void]$vmDetails.Add([PSCustomObject]@{
            Name       = $vm.Name
            VMID       = $vm.VMID
            PowerState = $vm.PowerState
            GuestOS    = $vm.GuestOS
            IPs        = $vm.IP
            vCPU       = $cpu
            RAMgb      = $ramGB
            Disks      = $disks
            NICs       = $nics
        })
    }
    Write-Host "`r                          " -NoNewline; Write-Host ""
    B-OK "  vSphere inventory done: $($vmDetails.Count) VMs, $($datastores.Count) datastores, $($esxHosts.Count) hosts"

    return [PSCustomObject]@{
        _type      = 'vSphereInventory'
        CollectedAt= (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Server     = $Conn.Server
        APIVersion = $Conn.APIVer
        ESXHosts   = $esxHosts
        Datastores = $datastores
        VMs        = $vmDetails
    }
}

# -- Hyper-V enumeration via WMI -----------------------------------------------

function Get-HyperVVMs {
    # Primary: Get-VM inline (local) or Invoke-Command (remote)
    # Fallback: WMI Msvm_ComputerSystem (if WinRM unavailable on remote HV host)
    param([string]$HVHost, [PSCredential]$Cred)
    $isLocal = ($HVHost -match '^(localhost|127\.0\.0\.1)$' -or $HVHost -ieq $env:COMPUTERNAME)
    try {
        $vmCollect = {
            Get-VM | ForEach-Object {
                $vm  = $_
                $ads = Get-VMNetworkAdapter -VM $vm -EA SilentlyContinue
                $ips = $ads | ForEach-Object {
                    $_.IPAddresses | Where-Object { $_ -notmatch ':' -and $_ -ne '0.0.0.0' }
                }
                [PSCustomObject]@{
                    Name   = $vm.Name
                    State  = $vm.State.ToString()
                    IPs    = (($ips | Select-Object -First 4) -join ', ')
                }
            }
        }
        $vmData = if ($isLocal) {
            B-OK "  HV host is local -- running Get-VM inline (no WinRM needed)"
            & $vmCollect
        } else {
            $icArgs = @{ ComputerName = $HVHost; ErrorAction = 'Stop' }
            if ($Cred) { $icArgs.Credential = $Cred }
            Invoke-Command @icArgs -ScriptBlock $vmCollect
        }
        $list = [System.Collections.ArrayList]@()
        foreach ($vm in $vmData) {
            $stateMap = @{ 'Running'='Running'; 'Off'='Off'; 'Paused'='Paused'; 'Saved'='Saved'; 'Starting'='Starting' }
            $ps = if ($stateMap.ContainsKey($vm.State)) { $stateMap[$vm.State] } else { $vm.State }
            [void]$list.Add([PSCustomObject]@{
                Name       = $vm.Name
                IP         = $vm.IPs
                PowerState = $ps
                GuestOS    = 'Windows (Hyper-V guest)'
                Source     = "Hyper-V ($HVHost)"
                VMID       = $vm.Name
            })
        }
        return $list
    } catch {
        B-Warn "Invoke-Command failed on $HVHost (WinRM unavailable or workgroup host -- domain creds won't work on non-domain machines)"
        B-Warn "Falling back to WMI/DCOM  --  VM list may be partial and IPs will be blank"
        B-Warn "If host is WORKGROUP: re-run and enter LOCAL admin creds (./Administrator) for this HV host"
        return Get-HyperVVMs-WMI -HVHost $HVHost -Cred $Cred
    }
}

function Get-HyperVVMs-WMI {
    # WMI fallback when Invoke-Command / WinRM is unavailable on the HV host.
    param([string]$HVHost, [PSCredential]$Cred)
    $isLocal = ($HVHost -match '^(localhost|127\.0\.0\.1)$' -or $HVHost -ieq $env:COMPUTERNAME)
    try {
        $wmiConn  = if ($isLocal) { @{} } else { @{ ComputerName=$HVHost; Cred=$Cred } }
        $vms = Get-WmiOrCim -Class 'Msvm_ComputerSystem' -Namespace 'root\virtualization\v2' `
                            -Filter "Caption='Virtual Machine'" -EA 'Stop' @wmiConn
        # Get all adapter configs once, then match by VM GUID (Msvm_ComputerSystem.Name IS the GUID)
        $allAdapters = Get-WmiOrCim -Class 'Msvm_GuestNetworkAdapterConfiguration' `
                                    -Namespace 'root\virtualization\v2' `
                                    -EA 'SilentlyContinue' @wmiConn
        $list = [System.Collections.ArrayList]@()
        $stateMap = @{ 2='Running'; 3='Off'; 9='Paused'; 6='Saved'; 10='Starting' }
        foreach ($vm in $vms) {
            $state   = if ($stateMap.ContainsKey([int]$vm.EnabledState)) { $stateMap[[int]$vm.EnabledState] } else { $vm.EnabledState }
            $vmGuid  = $vm.Name   # .Name on Msvm_ComputerSystem is the VM GUID
            $adapters = $allAdapters | Where-Object { $_.InstanceID -match [regex]::Escape($vmGuid) }
            $ip = ''
            if ($adapters) {
                $ips = $adapters | ForEach-Object {
                    $_.IPAddresses | Where-Object { $_ -notmatch ':' -and $_ -ne '0.0.0.0' }
                }
                $ip = ($ips | Select-Object -First 4) -join ', '
            }
            [void]$list.Add([PSCustomObject]@{
                Name       = $vm.ElementName
                IP         = $ip
                PowerState = $state
                GuestOS    = 'Windows (Hyper-V guest)'
                Source     = "Hyper-V ($HVHost)"
                VMID       = $vmGuid
            })
        }
        return $list
    } catch {
        B-Err "Hyper-V WMI query failed on $HVHost`: $_"
        return @()
    }
}

function Get-HyperVInventory {
    # Deep HV host dump -- VMs with full config, VHD usage, switches, host hardware.
    # Produces the -hv-inventory-.json file used for the Virtualization tab.
    param([string]$HVHost, [PSCredential]$Cred)
    $isLocal = ($HVHost -match '^(localhost|127\.0\.0\.1)$' -or $HVHost -ieq $env:COMPUTERNAME)
    B-Line "  Collecting Hyper-V host inventory from $HVHost..."

    $invBlock = {
        $cs  = Get-WmiObject Win32_ComputerSystem
        $cpu = Get-WmiObject Win32_Processor | Select-Object -First 1
        $vols = Get-WmiObject Win32_LogicalDisk -Filter 'DriveType=3' | ForEach-Object {
            [PSCustomObject]@{
                Drive   = $_.DeviceID
                Label   = $_.VolumeName
                TotalGB = [math]::Round($_.Size/1GB, 1)
                FreeGB  = [math]::Round($_.FreeSpace/1GB, 1)
                UsedPct = if ($_.Size -gt 0) { [math]::Round(100 - ($_.FreeSpace/$_.Size*100), 1) } else { 0 }
            }
        }
        $switches = Get-VMSwitch | ForEach-Object {
            [PSCustomObject]@{
                Name       = $_.Name
                Type       = $_.SwitchType.ToString()
                NetAdapter = $_.NetAdapterInterfaceDescription
            }
        }
        $vms = Get-VM | ForEach-Object {
            $vm  = $_
            $ads = Get-VMNetworkAdapter -VM $vm -EA SilentlyContinue
            $ips = $ads | ForEach-Object {
                $_.IPAddresses | Where-Object { $_ -notmatch ':' -and $_ -ne '0.0.0.0' }
            }
            $disks = Get-VMHardDiskDrive -VM $vm -EA SilentlyContinue | ForEach-Object {
                $vhd = Get-VHD $_.Path -EA SilentlyContinue
                [PSCustomObject]@{
                    Path           = $_.Path
                    ControllerType = $_.ControllerType.ToString()
                    SizeGB  = if ($vhd) { [math]::Round($vhd.Size/1GB, 1) }     else { $null }
                    UsedGB  = if ($vhd) { [math]::Round($vhd.FileSize/1GB, 1) } else { $null }
                    VHDType = if ($vhd) { $vhd.VhdType.ToString() }              else { $null }
                }
            }
            $nets = $ads | ForEach-Object {
                [PSCustomObject]@{
                    Name       = $_.Name
                    SwitchName = $_.SwitchName
                    MacAddress = $_.MacAddress
                    IPs        = (($_.IPAddresses | Where-Object { $_ -notmatch ':' -and $_ -ne '0.0.0.0' }) -join ', ')
                }
            }
            $snaps  = @(Get-VMSnapshot -VM $vm -EA SilentlyContinue).Count
            $intSvc = (Get-VMIntegrationService -VM $vm -EA SilentlyContinue |
                       Where-Object { $_.Enabled } | Select-Object -ExpandProperty Name) -join ', '
            [PSCustomObject]@{
                Name          = $vm.Name
                State         = $vm.State.ToString()
                Generation    = $vm.Generation
                Version       = $vm.Version
                vCPU          = $vm.ProcessorCount
                RAMgb         = [math]::Round($vm.MemoryAssigned/1GB, 2)
                RAMMinGB      = [math]::Round($vm.MemoryMinimum/1GB, 2)
                RAMMaxGB      = [math]::Round($vm.MemoryMaximum/1GB, 2)
                DynamicMemory = $vm.DynamicMemoryEnabled
                UptimeHours   = [math]::Round($vm.Uptime.TotalHours, 1)
                Snapshots     = $snaps
                AutoStart     = $vm.AutomaticStartAction.ToString()
                IPs           = (($ips | Select-Object -First 4) -join ', ')
                Disks         = $disks
                NetworkAdapters = $nets
                IntegrationServices = $intSvc
            }
        }
        [PSCustomObject]@{
            _type       = 'HyperVInventory'
            HVHost      = $env:COMPUTERNAME
            HostSummary = [PSCustomObject]@{
                Manufacturer = $cs.Manufacturer
                Model        = $cs.Model
                TotalRAMgb   = [math]::Round($cs.TotalPhysicalMemory/1GB, 1)
                CPUModel     = $cpu.Name.Trim()
                CPUCores     = $cpu.NumberOfCores
                CPULogical   = $cpu.NumberOfLogicalProcessors
                Volumes      = $vols
            }
            VirtualSwitches = $switches
            VMs             = $vms
        }
    }

    $inv = $null
    if ($isLocal) {
        try   { $inv = & $invBlock }
        catch { B-Err "  HV inventory failed (local): $_" }
    } else {
        try {
            $icArgs = @{ ComputerName = $HVHost; ErrorAction = 'Stop' }
            if ($Cred) { $icArgs.Credential = $Cred }
            $inv = Invoke-Command @icArgs -ScriptBlock $invBlock
        } catch { B-Err "  HV inventory failed on $HVHost`: $_" }
    }
    if ($inv) { B-OK "  HV inventory: $($inv.VMs.Count) VMs, $($inv.VirtualSwitches.Count) switches" }
    return $inv
}



# -----------------------------------------------------------------------------
# WMI / WINRM HELPERS (same as before)
# -----------------------------------------------------------------------------

function Get-SubnetHosts {
    param([string]$CIDR)
    try {
        if ($CIDR -notmatch '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}/\d{1,2}$') { throw "Bad CIDR" }
        $parts  = $CIDR -split '/'; $prefix = [int]$parts[1]
        if ($prefix -lt 16 -or $prefix -gt 30) { throw "Prefix must be /16-/30" }
        $ipBytes = [System.Net.IPAddress]::Parse($parts[0]).GetAddressBytes()
        [Array]::Reverse($ipBytes)
        $ipInt   = [BitConverter]::ToUInt32($ipBytes,0)
        $mask    = [uint32]([uint32]::MaxValue -shl (32-$prefix))
        $netInt  = $ipInt -band $mask
        $bcastInt = $netInt -bor (-bnot $mask -band [uint32]::MaxValue)
        $hosts = for ($i = $netInt+1; $i -lt $bcastInt; $i++) {
            $b = [BitConverter]::GetBytes([uint32]$i); [Array]::Reverse($b)
            [System.Net.IPAddress]::new($b).ToString()
        }
        return $hosts
    } catch { B-Err "Subnet parse: $_"; return @() }
}

function Invoke-PingSweep {
    param([string[]]$IPs)
    B-Line "Ping sweeping $($IPs.Count) addresses (~10-20 seconds)..."
    # Use .NET async pings -- avoids spawning one PS process per host (254 Start-Jobs = freeze)
    $tasks = $IPs | ForEach-Object {
        $p = [System.Net.NetworkInformation.Ping]::new()
        [PSCustomObject]@{ IP=$_; Ping=$p; Task=$p.SendPingAsync($_, 1500) }
    }
    $spin = 0; $spinC = @('|','/','-','\')
    $timeout = (Get-Date).AddSeconds(25)
    while (($tasks | Where-Object { -not $_.Task.IsCompleted }).Count -gt 0 -and (Get-Date) -lt $timeout) {
        $done = ($tasks | Where-Object { $_.Task.IsCompleted }).Count
        Write-Host ("`r  (o_o)  scanning... $($spinC[$spin%4])  $done/$($tasks.Count)   ") -NoNewline -ForegroundColor DarkCyan
        Start-Sleep -Milliseconds 200; $spin++
    }
    Write-Host "`r                                          " -NoNewline; Write-Host ""
    $live = [System.Collections.ArrayList]@()
    foreach ($t in $tasks) {
        try {
            if ($t.Task.IsCompleted -and $t.Task.Result.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
                [void]$live.Add($t.IP)
            }
        } catch { }
        $t.Ping.Dispose()
    }
    B-OK "$($live.Count) live hosts found"
    return ,$live.ToArray()
}

function Test-WMIAccess {
    param([string]$Target, [PSCredential]$Cred)
    try {
        $cs = Get-WmiOrCim -Class 'Win32_ComputerSystem' -ComputerName $Target -Cred $Cred -EA 'Stop' -UseDCOM
        return @{ OK=$true; Hostname=$cs.Name; OS=$cs.Caption }
    } catch { return @{ OK=$false; Error=$_.Exception.Message } }
}

function Get-WinRMStateViaWMI {
    param([string]$Target, [PSCredential]$Cred)
    try {
        $svc = Get-WmiOrCim -Class 'Win32_Service' -Filter "Name='WinRM'" -ComputerName $Target -Cred $Cred -EA 'Stop' -UseDCOM
        if (-not $svc) { return @{ OK=$false; Error='WinRM service not found' } }
        return @{ OK=$true; Running=($svc.State -eq 'Running'); StartMode=$svc.StartMode }
    } catch { return @{ OK=$false; Error=$_.Exception.Message } }
}

function Enable-WinRMViaWMI {
    param([string]$Target, [PSCredential]$Cred)
    try {
        $svc = Get-WmiOrCim -Class 'Win32_Service' -Filter "Name='WinRM'" -ComputerName $Target -Cred $Cred -EA 'Stop' -UseDCOM
        Invoke-WmiOrCimMethod -Instance $svc -MethodName 'ChangeStartMode' -Arguments @{ StartMode='Automatic' } | Out-Null
        Invoke-WmiOrCimMethod -Instance $svc -MethodName 'StartService'    -Arguments @{} | Out-Null
        Start-Sleep -Seconds 3
        $psRemCmd = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Enable-PSRemoting -Force -SkipNetworkProfileCheck 2>&1|Out-Null"'
        Invoke-WmiOrCimMethod -ClassName 'Win32_Process' -ComputerName $Target -Cred $Cred `
                              -MethodName 'Create' -Arguments @{ CommandLine=$psRemCmd } -UseDCOM | Out-Null
        Start-Sleep -Seconds 6
        for ($i=0; $i -lt 3; $i++) {
            try {
                $wp = @{ ComputerName=$Target; EA='Stop' }
                if ($Cred) { $wp.Credential=$Cred }
                Test-WSMan @wp | Out-Null; return $true
            } catch { Start-Sleep -Seconds 3 }
        }
        B-Err "WinRM started on $Target but port 5985 still not answering (firewall?)"
        return $false
    } catch { B-Err "Enable-WinRM failed on $Target`: $_"; return $false }
}

function Restore-WinRMState {
    param([string]$Target, [PSCredential]$Cred, [hashtable]$Orig)
    if ($Orig.Running) { return }
    Write-Host "    [WinRM] Restoring $Target to OFF / $($Orig.StartMode)..." -ForegroundColor DarkGray
    try {
        $cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " +
               "`"Disable-PSRemoting -Force 2>&1|Out-Null;" +
               "Stop-Service WinRM -Force 2>&1|Out-Null;" +
               "Set-Service WinRM -StartupType $($Orig.StartMode) 2>&1|Out-Null`""
        Invoke-WmiOrCimMethod -ClassName 'Win32_Process' -ComputerName $Target -Cred $Cred `
                              -MethodName 'Create' -Arguments @{ CommandLine=$cmd } -UseDCOM | Out-Null
        Start-Sleep -Seconds 4
        Write-Host "    [WinRM] Restored on $Target" -ForegroundColor DarkGray
    } catch { B-Err "WinRM restore failed on $Target`: $_" }
}

# -- Cleanup on exit -----------------------------------------------------------

$script:CleanupDone = $false
function Invoke-Cleanup {
    if ($script:CleanupDone) { return }; $script:CleanupDone = $true
    if ($script:WinRMRestoreMap.Count -gt 0) {
        Write-Host ""; B-Warn "Restoring WinRM states before exit..."
        foreach ($e in $script:WinRMRestoreMap.GetEnumerator()) {
            if (-not $e.Value.WasAlreadyOn) {
                Restore-WinRMState -Target $e.Key -Cred $script:DomainCred -Orig $e.Value.OriginalState
            }
        }
        B-OK "WinRM cleanup complete."
    }
}
Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action { Invoke-Cleanup } | Out-Null

# -----------------------------------------------------------------------------
# PHASE 0 - ENVIRONMENT TYPE
# -----------------------------------------------------------------------------

Write-Phase "Step 1 - Environment Type"
B-Line "what's running under these servers?"
Write-Host ""
Write-Host "  What hypervisor (or combination) is in use at this client?" -ForegroundColor White
Write-Host ""
Write-Host "  [1]  VMware vCenter  (manages multiple ESXi hosts)" -ForegroundColor Cyan
Write-Host "  [2]  VMware ESXi     (standalone host, no vCenter)" -ForegroundColor Cyan
Write-Host "  [3]  Microsoft Hyper-V" -ForegroundColor Cyan
Write-Host "  [4]  Bare metal / mix  (no hypervisor, or I'll add servers manually)" -ForegroundColor Cyan
Write-Host "  [5]  Multiple of the above" -ForegroundColor Cyan
Write-Host ""

$envType = ''
while ($envType -notin @('1','2','3','4','5')) { $envType = Read-Safe "  Choice [1-5]" }

$hypervisorSources = [System.Collections.ArrayList]@()

# Helper to collect a hypervisor entry
function Add-HypervisorSource {
    param([string]$Type)
    $src = @{ Type=$Type; Host=''; Cred=$null; VMs=@() }
    switch ($Type) {
        'vCenter' {
            $src.Host = Read-Safe "  vCenter hostname or IP"
            Write-Host "  vCenter credentials (user@vsphere.local or DOMAIN\user):" -ForegroundColor Gray
            $src.Cred = Get-Credential -Message "vCenter credentials for $($src.Host)"
        }
        'ESXi' {
            $src.Host = Read-Safe "  ESXi hostname or IP"
            Write-Host "  ESXi credentials (root or admin account):" -ForegroundColor Gray
            $src.Cred = Get-Credential -Message "ESXi credentials for $($src.Host)"
        }
        'HyperV' {
            Write-Host "  Hyper-V host hostname or IP" -ForegroundColor Gray
            Write-Host "  (press Enter or type 'localhost' if running this script ON the HV host)" -ForegroundColor DarkGray
            $src.Host = Read-Safe "  HV host"
            if (-not $src.Host -or $src.Host -eq '') { $src.Host = 'localhost' }

            $isLocalHV = ($src.Host -match '^(localhost|127\.0\.0\.1)$' -or
                          $src.Host -ieq $env:COMPUTERNAME)

            if ($isLocalHV) {
                B-OK "Running ON the HV host ($($src.Host)) -- no credentials needed"
                # cred stays null; Get-HyperVVMs will run Get-VM inline
            } else {
                Write-Host ""
                Write-Host "  HYPER-V HOST credentials:" -ForegroundColor Yellow
                Write-Host "    - Domain-joined host  ->  same domain creds as the servers  [Y]" -ForegroundColor DarkGray
                Write-Host "    - Standalone/workgroup host  ->  local admin  (.\Administrator)  [N]" -ForegroundColor DarkGray
                Write-Host ""
                $ans = Read-Safe "  Same as domain creds? [Y/N]"
                if ($ans -match '^[Nn]') {
                    Write-Host "  Enter LOCAL admin credentials for the Hyper-V host $($src.Host):" -ForegroundColor Gray
                    Write-Host "  (format: .\Administrator  or  .\localadmin)" -ForegroundColor DarkGray
                    $src.Cred = Get-Credential -Message "LOCAL admin for Hyper-V host $($src.Host) -- format: .\Administrator"
                }
                # else: cred stays null - will use $script:DomainCred at runtime
            }
        }
    }
    return $src
}

switch ($envType) {
    '1' { [void]$hypervisorSources.Add((Add-HypervisorSource 'vCenter')) }
    '2' { [void]$hypervisorSources.Add((Add-HypervisorSource 'ESXi')) }
    '3' { [void]$hypervisorSources.Add((Add-HypervisorSource 'HyperV')) }
    '4' { B-Line "bare metal / manual mode - will collect server list directly" }
    '5' {
        Write-Host ""
        Write-Host "  Add hypervisors one at a time. Press Enter on a blank line when done." -ForegroundColor Gray
        while ($true) {
            Write-Host ""
            Write-Host "  Next hypervisor: [1] vCenter  [2] ESXi  [3] Hyper-V  [Enter] Done" -ForegroundColor Gray
            $hv = Read-Safe "  >"
            if ($hv -eq '') { break }
            switch ($hv) {
                '1' { [void]$hypervisorSources.Add((Add-HypervisorSource 'vCenter')) }
                '2' { [void]$hypervisorSources.Add((Add-HypervisorSource 'ESXi')) }
                '3' { [void]$hypervisorSources.Add((Add-HypervisorSource 'HyperV')) }
            }
        }
    }
}

# -----------------------------------------------------------------------------
# PHASE 1 - DOMAIN CREDENTIALS
# -----------------------------------------------------------------------------

Write-Phase "Step 2 - Domain Credentials (for the SERVERS being discovered)"
B-Line "these creds run discovery on each Windows server - needs local or domain admin..."
Write-Host ""
Write-Host "  These are the credentials used to CONNECT TO and RUN DISCOVERY on each target server." -ForegroundColor White
Write-Host "  NOT for the hypervisor host - that was handled above (or will be asked separately)." -ForegroundColor DarkGray
Write-Host ""
Write-Host "  Format:  DOMAIN\username   (domain-joined servers - most common)" -ForegroundColor Cyan
Write-Host "           .\Administrator   (local admin, workgroup servers only)" -ForegroundColor DarkGray
Write-Host ""
try {
    $script:DomainCred = Get-Credential -Message "SERVER credentials -- DOMAIN\username or .\Administrator"
    if (-not $script:DomainCred) { throw "No credentials" }
    B-OK "Server creds loaded: $($script:DomainCred.UserName)"
} catch { B-Err "No credentials entered - cannot proceed"; exit 1 }

# -----------------------------------------------------------------------------
# PHASE 2 - ENUMERATE FROM HYPERVISORS
# -----------------------------------------------------------------------------

$hypervisorVMs = [System.Collections.ArrayList]@()

if ($hypervisorSources.Count -gt 0) {
    Write-Phase "Step 3 - Hypervisor Enumeration"
    B-Line "connecting to hypervisor(s) and pulling the VM list..."
    Write-Host ""

    foreach ($src in $hypervisorSources) {
        Write-Host ("  Connecting to $($src.Type): $($src.Host)...") -ForegroundColor Gray
        switch ($src.Type) {
            'vCenter' {
                $conn = Connect-VSphere -Server $src.Host -Cred $src.Cred
                if ($conn.OK) {
                    $vms = Get-VSphereVMs -Conn $conn
                    $vms | ForEach-Object { [void]$hypervisorVMs.Add($_) }
                    B-OK "  Got $($vms.Count) VMs from vCenter $($src.Host)"
                    $vInv = Get-VSphereInventory -Conn $conn -VMList $vms
                    Disconnect-VSphere -Conn $conn
                    if ($vInv) { [void]$script:PendingInventories.Add(@{ Type='vSphere'; Host=$src.Host; Data=$vInv }) }
                    [void]$script:vSphereSources.Add(@{ Host=$src.Host; Cred=$src.Cred; Type='vCenter' })
                } else {
                    B-Err "  vCenter connection failed: $($conn.Error)"
                }
            }
            'ESXi' {
                # ESXi direct - same REST API path as vCenter
                $conn = Connect-VSphere -Server $src.Host -Cred $src.Cred
                if ($conn.OK) {
                    $vms = Get-VSphereVMs -Conn $conn
                    $vms | ForEach-Object { $_.Source = "ESXi ($($src.Host))"; [void]$hypervisorVMs.Add($_) }
                    B-OK "  Got $($vms.Count) VMs from ESXi $($src.Host)"
                    $vInv = Get-VSphereInventory -Conn $conn -VMList $vms
                    Disconnect-VSphere -Conn $conn
                    if ($vInv) { [void]$script:PendingInventories.Add(@{ Type='vSphere'; Host=$src.Host; Data=$vInv }) }
                    [void]$script:vSphereSources.Add(@{ Host=$src.Host; Cred=$src.Cred; Type='ESXi' })
                    [void]$script:vSphereSources.Add(@{ Host=$src.Host; Cred=$src.Cred; Type='vCenter' })
                } else {
                    B-Err "  ESXi connection failed: $($conn.Error)"
                }
            }
            'HyperV' {
                $hvCred = if ($src.Cred) { $src.Cred } else { $script:DomainCred }
                $vms = Get-HyperVVMs -HVHost $src.Host -Cred $hvCred
                $vms | ForEach-Object { [void]$hypervisorVMs.Add($_) }
                B-OK "  Got $($vms.Count) VMs from Hyper-V $($src.Host)"
                $hvInv = Get-HyperVInventory -HVHost $src.Host -Cred $hvCred
                if ($hvInv) { [void]$script:PendingInventories.Add(@{ Type='HyperV'; Host=$src.Host; Data=$hvInv }) }
                # Auto-add the HV host itself as a discovery target (produces a server tab in the report)
                $isLocalHVSrc = ($src.Host -match '^(localhost|127\.0\.0\.1)$' -or $src.Host -ieq $env:COMPUTERNAME)
                if (-not $isLocalHVSrc) {
                    [void]$hypervisorVMs.Add([PSCustomObject]@{
                        Name       = $src.Host.ToUpper()
                        IP         = $src.Host
                        PowerState = 'Running'
                        GuestOS    = 'Windows (Hyper-V Host)'
                        Source     = "HyperV Host ($($src.Host))"
                        VMID       = $src.Host
                    })
                    B-OK "  Added $($src.Host) (HV host) as a server discovery target"
                }
            }
        }
    }

    Write-Host ""
    if ($hypervisorVMs.Count -gt 0) {
        Write-Host ("  VMs from hypervisor(s) ({0} total):" -f $hypervisorVMs.Count) -ForegroundColor White
        Write-Host ("  {0,-28} {1,-12} {2,-18} {3}" -f "Name","State","IP","Source") -ForegroundColor DarkMagenta
        Write-Divider
        foreach ($vm in $hypervisorVMs) {
            $stateColor = if ($vm.PowerState -match 'Running|POWERED_ON') { 'Green' }
            elseif ($vm.PowerState -match 'Off|Saved|Paused')    { 'Yellow' }
            else                                                 { 'Cyan' }
              $ipDisplay = if ($vm.IP -and $vm.IP -ne '') { $vm.IP -replace ',.*','' } else { '(no IP)' }
            Write-Host ("  {0,-28} {1,-12} {2,-18} {3}" -f $vm.Name, $vm.PowerState, $ipDisplay, $vm.Source) -ForegroundColor $stateColor
        }
        Write-Host ""
        Write-Host "  Only powered-on Windows VMs will be included. Use manual step to add/exclude." -ForegroundColor DarkGray
    }
}

# -----------------------------------------------------------------------------
# PHASE 3 - MANUAL TARGETS
# -----------------------------------------------------------------------------

$manualTargets = [System.Collections.ArrayList]@()

if ($hypervisorSources.Count -gt 0 -and $hypervisorVMs.Count -gt 0) {
    # ?? HV / vCenter mode: VMs already known  --  just ask for outliers ??????????

    # ── SUGGESTED SERVERS — AD/DNS SCAN ──────────────────────────────────────
    Write-Phase "Step 3b - Suggested Servers"
    Write-Host "  Scan AD/DNS for servers not already in your discovery list?" -ForegroundColor White
    Write-Host "  (useful for bare-metal boxes or servers outside the hypervisor)" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  [Y]  Scan  (AD first, DNS fallback, ~10s max)" -ForegroundColor Cyan
    Write-Host "  [N]  Skip  (recommended if all servers are VMs in vCenter/Hyper-V)" -ForegroundColor Cyan
    Write-Host ""
    $scanChoice = ''
    while ($scanChoice -notin @('Y','N')) {
        $scanChoice = (Read-Host "  Scan? [Y/N]").Trim().ToUpper()
    }

    $suggestedAll = @()
    if ($scanChoice -eq 'Y') {
        B-Line "scanning..."
        $scanJob = Start-Job -ScriptBlock ([ScriptBlock]::Create("function Get-SuggestedHypervisors {`n" + (Get-Command Get-SuggestedHypervisors).Definition + "`n}`nGet-SuggestedHypervisors"))
        $skipped = $false
        while (-not $scanJob.HasExited) {
            if ([Console]::KeyAvailable) {
                [Console]::ReadKey($true) | Out-Null
                Stop-Job $scanJob
                $skipped = $true
                break
            }
            Start-Sleep -Milliseconds 200
        }
        if ($skipped) {
            B-Warn "Scan cancelled."
        } else {
            $suggestedAll = @(Receive-Job $scanJob)
        }
        Remove-Job $scanJob -Force -ErrorAction SilentlyContinue
    } else {
        B-Line "Skipped."
    }
    $knownNames    = @($hypervisorVMs | ForEach-Object { $_.Name.ToUpper() }) +
                     @($hypervisorSources | ForEach-Object { $_.Host.ToUpper() })
    $newSuggested  = @($suggestedAll | Where-Object { $n = $_.Name.ToUpper(); $knownNames -notcontains $n })

    if ($newSuggested.Count -gt 0) {
        Write-Host ("  Found {0} candidate(s) not yet in your discovery list:" -f $newSuggested.Count) -ForegroundColor White
        Write-Host ""
        $suggMap = @{}
        for ($si = 0; $si -lt $newSuggested.Count; $si++) {
            $s    = $newSuggested[$si]
            $num  = $si + 1
            $ipStr = if ($s.IP) { $s.IP } else { 'IP unknown' }
            Write-Host ("  [{0}]  {1,-22} {2,-18} [{3}]" -f $num, $s.Name, $ipStr, $s.Source) -ForegroundColor Cyan
            if ($s.Description) { Write-Host ("        {0}" -f $s.Description) -ForegroundColor DarkGray }
            $suggMap["$num"] = $s
        }
        Write-Host ""
        Write-Host "  Enter numbers to add to discovery (comma-separated), or Enter to skip:" -ForegroundColor Gray
        $picks = Read-Safe "  >"
        if ($picks.Trim()) {
            $picks -split '[,\s]+' | Where-Object { $_.Trim() } | ForEach-Object {
                if ($suggMap.ContainsKey($_.Trim())) {
                    $s = $suggMap[$_.Trim()]
                    $ip = if ($s.IP) { ($s.IP -split ',')[0].Trim() } else { $s.Name }
                    [void]$manualTargets.Add([PSCustomObject]@{
                        Name       = $s.Name
                        IP         = $ip
                        PowerState = 'Unknown'
                        GuestOS    = if ($s.OS) { $s.OS } else { 'Unknown' }
                        Source     = "Suggested ($($s.Source))"
                        VMID       = $s.Name
                    })
                    B-OK "Added: $($s.Name) ($ip)"
                }
            }
        }
    } else {
        B-Line "No additional hypervisor/server candidates found in AD or DNS."
    }
    Write-Host ""

    Write-Phase "Step 4 - Additional Targets (Outliers)"
    B-Line "$($hypervisorVMs.Count) VM(s) discovered from hypervisor(s). Any standalone servers NOT on it?"
    Write-Host ""
    Write-Host "  Physical servers, management boxes, or anything outside the hypervisor." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  [1]  Add by hostname or IP" -ForegroundColor White
    Write-Host "  [2]  Load from text file (one per line)" -ForegroundColor White
    Write-Host "  [3]  Skip  --  hypervisor list is complete" -ForegroundColor White
    Write-Host ""
    $addMode = Read-Safe "  Choice [1/2/3]"
    switch ($addMode) {
        '1' {
            Write-Host ""
            Write-Host "  Enter hostnames/IPs (comma-separated or one per line). Blank line = done." -ForegroundColor Gray
            while ($true) {
                $line = Read-Host "  > "; if ($line.Trim() -eq '') { break }
                $line -split '[,\s]+' | Where-Object { $_.Trim() } | ForEach-Object {
                    [void]$manualTargets.Add([PSCustomObject]@{ Name=$_.Trim(); IP=$_.Trim(); PowerState='Unknown'; GuestOS='Unknown'; Source='Manual' })
                }
            }
        }
        '2' {
            $fp = (Read-Safe "  File path").Trim('"')
            if (Test-Path $fp) {
                Get-Content $fp | Where-Object { $_.Trim() -and -not $_.StartsWith('#') } | ForEach-Object {
                    [void]$manualTargets.Add([PSCustomObject]@{ Name=$_.Trim(); IP=$_.Trim(); PowerState='Unknown'; GuestOS='Unknown'; Source='File' })
                }
                B-OK "$($manualTargets.Count) targets loaded from file"
            } else { B-Err "File not found: $fp" }
        }
        '3' { B-Line "Using hypervisor list only  --  no outlier targets added" }
        default { B-Line "Skipping additional targets" }
    }
} else {
    # ?? Bare metal / no hypervisor: full target entry ?????????????????????????
    Write-Phase "Step 3b - Suggested Servers"
    B-Line "scanning Active Directory for server candidates..."
    Write-Host ""
    $suggestedServers = Get-SuggestedServers
    if ($suggestedServers.Count -gt 0) {
        Write-Host ("  Found {0} server(s) in Active Directory:" -f $suggestedServers.Count) -ForegroundColor White
        Write-Host ""
        $srvMap = @{}
        for ($si = 0; $si -lt $suggestedServers.Count; $si++) {
            $s = $suggestedServers[$si]; $num = $si + 1
            $ipStr = if ($s.IP) { $s.IP } else { 'IP unknown' }
            Write-Host ("  [{0}]  {1,-28} {2,-18} [{3}]" -f $num, $s.Name, $ipStr, $s.Source) -ForegroundColor Cyan
            if ($s.OS) { Write-Host ("        {0}" -f $s.OS) -ForegroundColor DarkGray }
            $srvMap["$num"] = $s
        }
        Write-Host ""
        Write-Host "  Enter numbers to add (comma-separated), A for all, or Enter to skip:" -ForegroundColor Gray
        $picks = Read-Safe "  >"
        if ($picks.Trim() -ieq 'A') {
            foreach ($s in $suggestedServers) {
                $ip = if ($s.IP) { ($s.IP -split ',')[0].Trim() } else { $s.Name }
                [void]$manualTargets.Add([PSCustomObject]@{
                    Name=$s.Name; IP=$ip; PowerState='Unknown'
                    GuestOS=if ($s.OS) { $s.OS } else { 'Unknown' }; Source="Suggested ($($s.Source))"
                })
            }
            B-OK "$($manualTargets.Count) server(s) added from AD"
        } elseif ($picks.Trim()) {
            $picks -split '[,\s]+' | Where-Object { $_.Trim() } | ForEach-Object {
                if ($srvMap.ContainsKey($_.Trim())) {
                    $s  = $srvMap[$_.Trim()]
                    $ip = if ($s.IP) { ($s.IP -split ',')[0].Trim() } else { $s.Name }
                    [void]$manualTargets.Add([PSCustomObject]@{
                        Name=$s.Name; IP=$ip; PowerState='Unknown'
                        GuestOS=if ($s.OS) { $s.OS } else { 'Unknown' }; Source="Suggested ($($s.Source))"
                    })
                }
            }
            B-OK "$($manualTargets.Count) server(s) added from AD"
        }
    } else {
        B-Line "No servers found in Active Directory — or no AD access from this machine."
    }

    Write-Phase "Step 4 - Server Targets"
    B-Line "add any remaining targets not covered above..."
    Write-Host ""
    Write-Host "  [1]  Enter hostnames or IPs manually" -ForegroundColor White
    Write-Host "  [2]  Load from text file (one per line)" -ForegroundColor White
    Write-Host "  [3]  Subnet scan  (~10-20s per /24, ICMP must be open)" -ForegroundColor Yellow
    Write-Host "  [4]  Skip" -ForegroundColor DarkGray
    Write-Host ""
    $addMode = Read-Safe "  Choice [1/2/3/4]"
    switch ($addMode) {
        '1' {
            Write-Host ""
            Write-Host "  Enter hostnames/IPs (comma-separated or one per line). Blank line = done." -ForegroundColor Gray
            while ($true) {
                $line = Read-Host "  > "; if ($line.Trim() -eq '') { break }
                $line -split '[,\s]+' | Where-Object { $_.Trim() } | ForEach-Object {
                    [void]$manualTargets.Add([PSCustomObject]@{ Name=$_.Trim(); IP=$_.Trim(); PowerState='Unknown'; GuestOS='Unknown'; Source='Manual' })
                }
            }
        }
        '2' {
            $fp = (Read-Safe "  File path").Trim('"')
            if (Test-Path $fp) {
                Get-Content $fp | Where-Object { $_.Trim() -and -not $_.StartsWith('#') } | ForEach-Object {
                    [void]$manualTargets.Add([PSCustomObject]@{ Name=$_.Trim(); IP=$_.Trim(); PowerState='Unknown'; GuestOS='Unknown'; Source='File' })
                }
                B-OK "$($manualTargets.Count) targets loaded from file"
            } else { B-Err "File not found: $fp" }
        }
        '3' {
            $cidr = Read-Safe "  CIDR (e.g. 10.0.1.0/24)"
            $allIPs = Get-SubnetHosts -CIDR $cidr
            if ($allIPs.Count -gt 0) {
                $liveIPs = Invoke-PingSweep -IPs $allIPs
                Write-Host ""; Write-Host "  Live hosts:" -ForegroundColor Gray
                $liveIPs | ForEach-Object { Write-Host "    $_" -ForegroundColor White }
                Write-Host ""
                $excl = Read-Safe "  Exclude any? (comma-separated IPs, or Enter for none)"
                $excludes = if ($excl) { $excl -split ',' | ForEach-Object { $_.Trim() } } else { @() }
                $liveIPs | Where-Object { $_ -notin $excludes } | ForEach-Object {
                    [void]$manualTargets.Add([PSCustomObject]@{ Name=$_; IP=$_; PowerState='Unknown'; GuestOS='Unknown'; Source='SubnetScan' })
                }
            }
        }
        '4' { B-Line "Skipping target entry" }
        default { B-Line "Skipping" }
    }
}

# -- Exclusions from hypervisor list ------------------------------------------

if ($hypervisorVMs.Count -gt 0) {
    Write-Host ""
    $excAns = Read-Safe "  Exclude any VMs from the hypervisor list? (Y/N)"
    if ($excAns -match '^[Yy]') {
        Write-Host "  Enter VM names to exclude (comma-separated):" -ForegroundColor Gray
        $excNames = (Read-Host "  > ") -split ',' | ForEach-Object { $_.Trim() }
        $hypervisorVMs = [System.Collections.ArrayList]@($hypervisorVMs | Where-Object { $_.Name -notin $excNames })
        B-OK "Excluded $($excNames.Count) VM(s)"
    }
}

# -- Build master target list --------------------------------------------------

$allVMTargets = [System.Collections.ArrayList]@()
# Add powered-on VMs from hypervisor
foreach ($vm in $hypervisorVMs) {
    if ($vm.PowerState -match 'Running|POWERED_ON|running') { [void]$allVMTargets.Add($vm) }
    elseif ($vm.PowerState -eq 'Unknown') { [void]$allVMTargets.Add($vm) }
    # Off VMs: show in plan but mark as skipped
    else {
        $vm | Add-Member -NotePropertyName '_Offline' -NotePropertyValue $true -Force
        [void]$allVMTargets.Add($vm)
    }
}
foreach ($m in $manualTargets) { [void]$allVMTargets.Add($m) }

if ($allVMTargets.Count -eq 0) {
    B-Err "No targets in list. Add servers manually or check hypervisor connection."
    exit 1
}

# ?? Target Summary ????????????????????????????????????????????????????????????
Write-Host ""
Write-Host ("=" * 72) -ForegroundColor DarkCyan
Write-Host ("  DISCOVERED TARGETS  ({0} total)" -f $allVMTargets.Count) -ForegroundColor Cyan
Write-Host ("=" * 72) -ForegroundColor DarkCyan
Write-Host ""
Write-Host ("  {0,-30} {1,-20} {2,-12} {3}" -f "Name","IP(s)","Source","Power") -ForegroundColor DarkCyan
Write-Host ("  " + ("-" * 68)) -ForegroundColor DarkCyan
foreach ($t in $allVMTargets) {
    $ips      = if ($t.IP -and $t.IP -ne '') { $t.IP } else { '(pending)' }
    $ipsShort = if ($ips.Length -gt 19) { $ips.Substring(0,17) + '..' } else { $ips }
    $srcShort = if ($t.Source -and $t.Source.Length -gt 11) { $t.Source.Substring(0,9) + '..' } else { if ($t.Source) { $t.Source } else { '-' } }
    $pwr      = if ($t._Offline)                           { 'OFF'     }
                elseif ($t.PowerState -match 'Running|POWERED_ON|on') { 'Running' }
                elseif ($t.PowerState -and $t.PowerState -ne 'Unknown') { $t.PowerState }
                else { '?' }
    $osHint   = if ($t.GuestOS -and $t.GuestOS -ne 'Unknown') {
                    if (Test-IsLinux $t.GuestOS) { '[Linux]' } else { '[Win]' }
                } else { '[?]' }
    $clr = if ($t._Offline) { 'DarkGray' } elseif ($osHint -eq '[Linux]') { 'Cyan' } else { 'White' }
    Write-Host ("  {0,-30} {1,-20} {2,-12} {3,-8} {4}" -f $t.Name, $ipsShort, $srcShort, $pwr, $osHint) -ForegroundColor $clr
}
Write-Host ""
Write-Host "  READ-ONLY POLICY:" -ForegroundColor DarkGreen
Write-Host "    All collection uses Get-* cmdlets only. Nothing is written, modified," -ForegroundColor DarkGreen
Write-Host "    or deleted on any server." -ForegroundColor DarkGreen
Write-Host "    WinRM bootstrap: if WinRM is OFF we enable it via WMI, collect data," -ForegroundColor DarkGreen
Write-Host "    then restore the original OFF state. That is the only state change." -ForegroundColor DarkGreen
Write-Host "    Choose [S] at the plan prompt to skip bootstrap entirely (strict mode)." -ForegroundColor DarkGreen
Write-Host ("=" * 72) -ForegroundColor DarkCyan
Write-Host ""
$preAns = (Read-Safe "  Proceed to pre-flight check? [Y/N]").ToUpper()
if ($preAns -ne 'Y') { Write-Host "  Aborted." -ForegroundColor Red; exit 0 }

do {
# -----------------------------------------------------------------------------
# PHASE 4 - PROBE ALL TARGETS
# -----------------------------------------------------------------------------

Write-Phase "Step 5 - Probing All Targets"
B-Line "WMI-probing every target to check OS, WinRM state, and connectivity..."
Write-Host ""

$planRows = [System.Collections.ArrayList]@()

foreach ($vm in $allVMTargets) {
    # Pick the best address to connect to
    $addr = if ($vm.IP -and $vm.IP -ne '') { ($vm.IP -split ',')[0].Trim() } else { $vm.Name }
    $displayName = $vm.Name

    Write-Host ("  Probing {0,-28} ({1})..." -f $displayName, $addr) -NoNewline -ForegroundColor Gray

    $row = [PSCustomObject]@{
        DisplayName   = $displayName
        Address       = $addr
        Source        = $vm.Source
        GuestOSHint   = $vm.GuestOS
        PowerState    = $vm.PowerState
        Ping          = $false
        WMI           = $false
        ResolvedOS    = ''
        ResolvedHost  = ''
        WinRMState    = 'Unknown'
        WinRMPort     = $false
        ConnectMethod = ''
        Action        = 'SKIP'
        SkipReason    = ''
        _VMObj        = $vm
    }

    # Skip offline VMs immediately
    if ($vm._Offline) {
        Write-Host " POWERED OFF" -ForegroundColor DarkGray
        $row.Action = 'SKIP'; $row.SkipReason = 'Powered off'
        [void]$planRows.Add($row); continue
    }

    # Linux pre-screen: skip before any WMI/WinRM attempt
    if (Test-IsLinux $vm.GuestOS) {
        Write-Host " LINUX (skipped)" -ForegroundColor Cyan
        $row.Action     = 'LINUX'
        $row.SkipReason = "Linux guest ($($vm.GuestOS)) -- WinRM N/A"
        $row.ConnectMethod = 'N/A (Linux)'
        $row.ResolvedOS = $vm.GuestOS
        [void]$planRows.Add($row); continue
    }

    # Ping
    try { $row.Ping = Test-Connection -ComputerName $addr -Count 1 -Quiet -EA SilentlyContinue } catch { }
    if (-not $row.Ping) {
        Write-Host " NO PING" -ForegroundColor Red
        $row.Action = 'MANUAL'; $row.SkipReason = 'No ping response'
        [void]$planRows.Add($row); continue
    }

    # WMI
    $wmi = Test-WMIAccess -Target $addr -Cred $script:DomainCred
    if (-not $wmi.OK) {
        $row.Ping = $true
        # If vCenter guest hint is non-Windows, treat as Linux/appliance instead of MANUAL
        $guestHint  = (($vm.GuestOS -replace '_',' ') + '').ToUpper().Trim()
        $looksWin   = ($guestHint -match 'WINDOWS|WIN') -or [string]::IsNullOrEmpty($guestHint) -or $guestHint -eq 'UNKNOWN'
        if (-not $looksWin) {
            Write-Host " LINUX/APPLIANCE" -ForegroundColor Cyan
            $row.Action        = 'LINUX'
            $row.SkipReason    = "Non-Windows guest ($($vm.GuestOS)) -- WinRM N/A"
            $row.ConnectMethod = 'N/A (Linux/Appliance)'
            $row.ResolvedOS    = $vm.GuestOS
        } else {
            Write-Host " WMI FAILED" -ForegroundColor Red
            $row.Action    = 'MANUAL'
            $row.SkipReason = "WMI failed - $($wmi.Error)"
        }
        [void]$planRows.Add($row); continue
    }
    $row.WMI          = $true
    $row.ResolvedHost = $wmi.Hostname
    $row.ResolvedOS   = $wmi.OS

    # WinRM state (read-only WMI query)
    $ws = Get-WinRMStateViaWMI -Target $addr -Cred $script:DomainCred
    if ($ws.OK) {
        $row.WinRMState = if ($ws.Running) { 'ON' } else { "OFF ($($ws.StartMode))" }
        $row.ConnectMethod = if ($ws.Running) { 'WinRM (already on)' } else { 'WinRM (WMI bootstrap)' }
    } else {
        $row.WinRMState    = 'Unknown'
        $row.ConnectMethod = 'WinRM (will attempt)'
    }

    # Port 5985 reachability (read-only  --  Test-NetConnection)
    try {
        $row.WinRMPort = [bool](Test-NetConnection -ComputerName $addr -Port 5985 `
            -InformationLevel Quiet -EA SilentlyContinue -WarningAction SilentlyContinue)
    } catch { $row.WinRMPort = $false }

    $row.Action = 'AUTO'
    $portTag  = if ($row.WinRMPort) { '5985:open' } else { '5985:closed' }
    $portClr  = if ($row.WinRMPort) { 'Green' } else { 'Yellow' }
    $winrmTag = "WinRM:$($row.WinRMState)"
    Write-Host (" OK") -NoNewline -ForegroundColor Green
    Write-Host ("  | $($row.ResolvedHost)") -NoNewline -ForegroundColor White
    Write-Host ("  | $winrmTag") -NoNewline -ForegroundColor $(if ($row.WinRMState -eq 'ON') { 'Green' } else { 'Yellow' })
    Write-Host ("  | $portTag") -ForegroundColor $portClr
    [void]$planRows.Add($row)
}

# -----------------------------------------------------------------------------
# LOCAL MACHINE HANDLING
# Always discover the machine this script is running on - no creds needed.
# Rules:
#   1. If the local machine already appears in $planRows (e.g. HV host enumerated
#      itself as a VM somehow), force its ConnectMethod to LOCAL so we never try
#      to WinRM to ourselves.
#   2. If the local machine is not in $planRows at all, prepend it as LOCAL.
# This guarantees the script host is always discovered, always locally.
# -----------------------------------------------------------------------------
$localHostname = $env:COMPUTERNAME.ToUpper()
Write-Host ""
B-OK "Script is running on: $localHostname - will always be included as LOCAL (no creds)"

$existingLocalRow = $planRows | Where-Object {
    $_.ResolvedHost  -eq $localHostname -or
    $_.DisplayName   -eq $localHostname -or
    ($_.Address -ne 'localhost' -and $_.Address -eq $env:COMPUTERNAME)
} | Select-Object -First 1

if ($existingLocalRow) {
    # Already in the plan - force it to LOCAL so WinRM is never attempted against ourselves
    $existingLocalRow.ConnectMethod = 'LOCAL'
    $existingLocalRow.Address       = 'localhost'
    $existingLocalRow.WinRMState    = 'N/A (local)'
    $existingLocalRow.Action        = 'AUTO'
    B-OK "$localHostname found in plan - forced to LOCAL (was: $($existingLocalRow.ConnectMethod))"
} else {
    # Not in the plan at all - add it
    $localRow = [PSCustomObject]@{
        DisplayName   = $localHostname
        Address       = 'localhost'
        Source        = 'Local (script host)'
        GuestOSHint   = ''
        PowerState    = 'Running'
        Ping          = $true
        WMI           = $true
        ResolvedOS    = [System.Environment]::OSVersion.VersionString
        ResolvedHost  = $localHostname
        WinRMState    = 'N/A (local)'
        ConnectMethod = 'LOCAL'
        Action        = 'AUTO'
        SkipReason    = ''
        _VMObj        = $null
    }
    $planRows.Insert(0, $localRow)
    B-OK "$localHostname added to plan as LOCAL target"
}

# -----------------------------------------------------------------------------
# PHASE 5 - PRE-RUN PLAN TABLE
# -----------------------------------------------------------------------------

$autoRows   = @($planRows | Where-Object { $_.Action -eq 'AUTO'   })
$manualRows = @($planRows | Where-Object { $_.Action -eq 'MANUAL' })
$skipRows   = @($planRows | Where-Object { $_.Action -eq 'SKIP'   })
$linuxRows  = @($planRows | Where-Object { $_.Action -eq 'LINUX'  })

Write-Host ""
Write-Host ""
Write-Host ("=" * 72) -ForegroundColor DarkMagenta
Write-Host "  DISCOVERY PLAN - Review this carefully before running" -ForegroundColor Magenta
Write-Host ("=" * 72) -ForegroundColor DarkMagenta
Write-Host ""

# Header
Write-Host ("  {0,-26} {1,-12} {2,-18} {3,-8} {4,-20} {5}" -f "Server","Source","OS","WinRM","Connect Method","Action") -ForegroundColor DarkMagenta
Write-Divider

foreach ($r in $planRows) {
    $osDisplay  = if ($r.ResolvedOS) { $r.ResolvedOS } elseif ($r.GuestOSHint -and $r.GuestOSHint -ne 'Unknown') { $r.GuestOSHint + ' (hint)' } else { '(unknown)' }
    $osShort    = if ($osDisplay.Length -gt 17) { $osDisplay.Substring(0,15) + '..' } else { $osDisplay }
    $srcShort   = if ($r.Source -and $r.Source.Length -gt 11) { $r.Source.Substring(0,9) + '..' } else { if ($r.Source) { $r.Source } else { '-' } }
    $methShort  = if ($r.ConnectMethod) { $r.ConnectMethod } elseif ($r.SkipReason) { $r.SkipReason } else { '-' }
    if ($methShort.Length -gt 19) { $methShort = $methShort.Substring(0,17) + '..' }
    $winrmShort = if ($r.WinRMState) {
                      $s = $r.WinRMState -replace '\(.*\)',''  # strip "(Manual)" etc
                      $p = if ($r.WinRMPort) { '+' } else { '' }  # + means port open too
                      ($s.Trim() + $p).Trim()
                  } else { '?' }

    $actionLabel = switch ($r.Action) {
        'AUTO'   { '[AUTO]  ' }
        'MANUAL' { '[MANUAL]' }
        'SKIP'   { '[SKIP]  ' }
        'LINUX'  { '[LINUX] ' }
        default  { '[?]     ' }
    }
    $color = switch ($r.Action) {
        'AUTO'   { 'White'    }
        'MANUAL' { 'Yellow'   }
        'SKIP'   { 'DarkGray' }
        'LINUX'  { 'Cyan'     }
        default  { 'Gray'     }
    }
    Write-Host ("  {0,-26} {1,-12} {2,-18} {3,-8} {4,-20} {5}" -f $r.DisplayName, $srcShort, $osShort, $winrmShort, $methShort, $actionLabel) -ForegroundColor $color
}

Write-Divider
Write-Host ""
$winrmOffCount = ($autoRows | Where-Object { $_.WinRMState -ne 'ON' -and $_.WinRMState -ne 'N/A (local)' }).Count
Write-Host ("  AUTO   ({0,2}) - Will run without you touching anything" -f $autoRows.Count) -ForegroundColor Green
if ($winrmOffCount -gt 0) {
    Write-Host ("           ({0} need WinRM bootstrap - enabled via WMI, restored after)" -f $winrmOffCount) -ForegroundColor DarkYellow
}
if ($manualRows.Count -gt 0) {
    Write-Host ("  MANUAL ({0,2}) - Unreachable; you must run the script locally on these:" -f $manualRows.Count) -ForegroundColor Yellow
    foreach ($r in $manualRows) {
        Write-Host ("           $($r.DisplayName)  -  $($r.SkipReason)") -ForegroundColor DarkYellow
    }
}
if ($skipRows.Count -gt 0) {
    Write-Host ("  SKIP   ({0,2}) - Powered off or excluded" -f $skipRows.Count) -ForegroundColor DarkGray
if ($linuxRows.Count -gt 0) {
    Write-Host ("  LINUX  ({0,2}) - Linux guest OS  --  logged but no WinRM collection" -f $linuxRows.Count) -ForegroundColor Cyan
    $linuxRows | ForEach-Object { Write-Host ("           $($_.DisplayName)  --  $($_.ResolvedOS)") -ForegroundColor DarkCyan }
}
}
Write-Host ""

# Connection method legend
Write-Host "  Connection method key:" -ForegroundColor DarkGray
Write-Host "    WinRM (already on)   - WinRM found running. Nothing changed."    -ForegroundColor DarkGray
Write-Host "    WinRM (WMI bootstrap)- WinRM is OFF. Will enable via WMI, run discovery, then turn it back OFF." -ForegroundColor DarkGray
Write-Host "    WinRM (will attempt) - WinRM state unclear; will try anyway."    -ForegroundColor DarkGray
Write-Host ""

if ($autoRows.Count -eq 0) {
    B-Warn "No servers reachable automatically. All would require manual runs."
    B-Line "Check WMI/firewall access and re-run, or proceed to get manual instructions."
}

Write-Host ("=" * 72) -ForegroundColor DarkMagenta
Write-Host ""

# -- Final confirmation --------------------------------------------------------

Write-Host "  Does this plan look right?" -ForegroundColor White
Write-Host ""
Write-Host "  [Y]  Run - discovery only; WinRM bootstrap where needed (state restored after)" -ForegroundColor Green
Write-Host "  [S]  Strict read-only - skip any server where WinRM is currently OFF (no bootstrap)" -ForegroundColor Cyan
Write-Host "  [N]  Abort - nothing changed on any server" -ForegroundColor Red
Write-Host "  [R]  Re-probe - re-test connectivity with current credentials" -ForegroundColor DarkCyan
Write-Host "  [C]  Re-enter credentials - change account and re-probe" -ForegroundColor Yellow
Write-Host ""

$goAns = ''
while ($goAns -notin @('Y','S','N','R','C')) { $goAns = (Read-Safe "  Ready to run? [Y/S/N/R/C]").ToUpper() }

# Strict read-only mode: drop any server that would need WinRM bootstrap
if ($goAns -eq 'S') {
    $bootstrapRows = @($autoRows | Where-Object { $_.ConnectMethod -match 'bootstrap' })
    if ($bootstrapRows.Count -gt 0) {
        B-Warn "Strict mode: skipping $($bootstrapRows.Count) server(s) where WinRM is OFF:"
        $bootstrapRows | ForEach-Object { Write-Host "    - $($_.DisplayName)  ($($_.WinRMState))" -ForegroundColor DarkYellow }
        $autoRows = @($autoRows | Where-Object { $_.ConnectMethod -notmatch 'bootstrap' })
        Write-Host ""
    } else {
        B-OK "Strict mode: all AUTO targets already have WinRM ON. No change to plan."
    }
    $goAns = 'Y'
}

if ($goAns -eq 'N') {
    B-Line "Aborted. Nothing was changed on any server."; Invoke-Cleanup; exit 0
}
if ($goAns -eq 'C') {
    Write-Host ""
    Write-Host "  Re-entering server credentials..." -ForegroundColor Yellow
    try {
        $newCred = Get-Credential -Message "SERVER credentials -- DOMAIN\username or .\Administrator"
        if ($newCred) {
            $script:DomainCred = $newCred
            B-OK "Credentials updated: $($script:DomainCred.UserName) - re-probing..."
        } else {
            B-Err "No credentials entered - keeping existing: $($script:DomainCred.UserName)"
        }
    } catch { B-Err "Credential prompt failed - keeping existing: $($script:DomainCred.UserName)" }
}
} while ($goAns -in @('R','C'))

# -----------------------------------------------------------------------------
# PHASE 6 - DISCOVERY LOOP
# -----------------------------------------------------------------------------

$sessionLabel  = "Discovery-Session-" + $script:SessionStart.ToString("yyyy-MM-dd-HHmm")
$sessionFolder = Join-Path $PSScriptRoot $sessionLabel
try { New-Item -ItemType Directory -Path $sessionFolder -Force | Out-Null } catch {
    B-Err "Cannot create output folder: $_"; exit 1
}

B-OK "Output folder: $sessionFolder"

# Session log file - append timestamped entries for every event
$script:SessionLog = Join-Path $sessionFolder "session-log.txt"
function Write-Log {
    param([string]$level, [string]$target, [string]$msg)
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "[$ts] [$level] [$target] $msg"
    Add-Content -Path $script:SessionLog -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
    # Also print failures loudly to terminal
    if ($level -eq 'FAIL') { Write-Host "  (x_x)  $line" -ForegroundColor Red }
}
Write-Log 'INFO' 'Session' "Started - $($autoRows.Count) AUTO targets"

$successCount = 0; $failCount = 0
$outputFiles  = [System.Collections.ArrayList]@()

# Write any pending hypervisor inventory JSONs
if ($script:PendingInventories.Count -gt 0) {
    $invDateStr = $script:SessionStart.ToString("yyyy-MM-dd")
    foreach ($inv in $script:PendingInventories) {
        $invLabel = $inv.Host -replace '[\\/:*?"<>|]', '_'
        $invFile  = Join-Path $sessionFolder "$invLabel-inventory-$invDateStr.json"
        try {
            $inv.Data | ConvertTo-Json -Depth 12 | Out-File $invFile -Encoding UTF8 -Force
            B-OK "Inventory JSON: $(Split-Path $invFile -Leaf)"
            [void]$outputFiles.Add($invFile)
            Write-Log 'OK' $inv.Host "Inventory JSON: $(Split-Path $invFile -Leaf)"
        } catch { B-Err "Failed to write inventory JSON for $($inv.Host): $_" }
    }
}

foreach ($row in $autoRows) {
    $target = $row.Address
    Write-Phase "Discovering: $($row.DisplayName)"

    # -- LOCAL vs REMOTE -------------------------------------------------------
    $isLocal = ($row.ConnectMethod -eq 'LOCAL')

    if ($isLocal) {
        # Running on the machine itself - no WinRM, no creds
        B-OK "LOCAL target - running discovery in-process (no WinRM needed)"
        Write-Log 'INFO' $row.DisplayName "LOCAL run - no WinRM needed"
    } else {
        # -- WinRM Bootstrap ---------------------------------------------------
        $ws = Get-WinRMStateViaWMI -Target $target -Cred $script:DomainCred
        $weEnabled = $false

        if ($ws.OK -and -not $ws.Running) {
            B-Warn "WinRM OFF on $target - enabling temporarily via WMI..."
            Write-Log 'INFO' $row.DisplayName "WinRM was OFF - attempting WMI bootstrap"
            if (Enable-WinRMViaWMI -Target $target -Cred $script:DomainCred) {
                $weEnabled = $true
                $script:WinRMRestoreMap[$target] = @{ WasAlreadyOn=$false; OriginalState=$ws }
                B-OK "WinRM enabled on $target - will restore when done"
                Write-Log 'INFO' $row.DisplayName "WinRM bootstrap succeeded"
            } else {
                B-Err "Could not enable WinRM on $target - skipping"
                Write-Log 'FAIL' $row.DisplayName "WinRM bootstrap failed - skipped. WMI error: $($ws.Error)"
                $failCount++; continue
            }
        } elseif (-not $ws.OK) {
            Write-Log 'WARN' $row.DisplayName "WMI probe returned error before WinRM check: $($ws.Error)"
            $script:WinRMRestoreMap[$target] = @{ WasAlreadyOn=$true; OriginalState=$ws }
            Write-Host "    [WinRM] State unknown - proceeding anyway" -ForegroundColor DarkYellow
        } else {
            $script:WinRMRestoreMap[$target] = @{ WasAlreadyOn=$true; OriginalState=$ws }
            Write-Host "    [WinRM] Already running on $target" -ForegroundColor DarkGreen
            Write-Log 'INFO' $row.DisplayName "WinRM already ON"
        }
    }

    # -- Run Discovery ---------------------------------------------------------
    Write-Host "    [Discovery] Running against $target..." -ForegroundColor Gray
    Write-Host ""
    $ok = $false
    $discError = ''
    try {
        if ($isLocal) {
            & $DiscoveryScript -OutputPath $sessionFolder
        } else {
            & $DiscoveryScript -ComputerName $target -Credential $script:DomainCred -OutputPath $sessionFolder
        }
        $ok = $true
    } catch {
        $discError = $_.ToString()
        B-Err "Discovery script exception on $target`: $discError"
        Write-Log 'FAIL' $row.DisplayName "Discovery script threw exception: $discError"
    }

    # Find output JSON
    $hostname = if ($row.ResolvedHost) { $row.ResolvedHost } else { $row.DisplayName }
    $dateStr  = (Get-Date).ToString("yyyy-MM-dd")
    $jsonPath = Join-Path $sessionFolder "$hostname-discovery-$dateStr.json"
    if (-not (Test-Path $jsonPath)) {
        $jsonPath = Get-ChildItem $sessionFolder -Filter "*discovery*$dateStr*.json" -EA SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1 | Select-Object -ExpandProperty FullName
    }
    if ($jsonPath -and (Test-Path $jsonPath)) {
        [void]$outputFiles.Add($jsonPath)
        $successCount++
        B-OK "JSON saved: $(Split-Path $jsonPath -Leaf)"
        Write-Log 'OK' $row.DisplayName "Success - $(Split-Path $jsonPath -Leaf)"
    } else {
        B-Err "No JSON output found for $($row.DisplayName) ($target)"
        Write-Log 'FAIL' $row.DisplayName "No JSON file produced. Script exit ok=$ok. Prior error: $(if($discError){"$discError"}else{'none - check Invoke-ServerDiscovery output above'})"
        $failCount++
    }

    # -- Restore WinRM ---------------------------------------------------------
    if (-not $isLocal -and $weEnabled) {
        Write-Host ""
        Restore-WinRMState -Target $target -Cred $script:DomainCred -Orig $ws
        $script:WinRMRestoreMap.Remove($target)
    }
    Write-Divider
}


# -----------------------------------------------------------------------------
# PHASE 6b - VSPHERE PERFORMANCE COLLECTION
# -----------------------------------------------------------------------------

if ($script:vSphereSources.Count -gt 0) {
    Write-Phase "Performance Data Collection"
    B-Line "collecting 90-day CPU/RAM/IOPS history from vSphere..."
    Write-Host ""

    # Locate Python — portable bundle takes priority over system install
    $pythonCmd = $null
    $portablePy = Join-Path $PSScriptRoot 'python\python.exe'
    if (Test-Path $portablePy) {
        $pythonCmd = $portablePy
    } else {
        foreach ($pyCandidate in @('python','python3','py')) {
            try {
                $pyOut = & $pyCandidate --version 2>&1
                if ($pyOut -match 'Python [23]') { $pythonCmd = $pyCandidate; break }
            } catch { }
        }
    }

    $collectorScript = Join-Path $PSScriptRoot 'collect_vsphere_perf.py'

    if (-not $pythonCmd) {
        B-Warn "Python not found in PATH -- skipping vSphere performance collection"
        B-Warn "Install Python 3.x and add it to PATH, then re-run, or run manually:"
        B-Warn "  python collect_vsphere_perf.py --vcenter <ip> --user <u> --pass <p> --output `"$sessionFolder`""
    } elseif (-not (Test-Path $collectorScript)) {
        B-Warn "collect_vsphere_perf.py not found alongside this script -- skipping"
        B-Warn "Expected: $collectorScript"
    } else {
        foreach ($vSrc in $script:vSphereSources) {
            Write-Host ""
            Write-Host ("  [$($vSrc.Type)] $($vSrc.Host) -- 90-day perf collection starting...") -ForegroundColor Cyan
            Write-Host "  (this can take a few minutes for large environments)" -ForegroundColor DarkGray
            $vUser = $vSrc.Cred.UserName
            $vPass = $vSrc.Cred.GetNetworkCredential().Password
            try {
                & $pythonCmd $collectorScript `
                    --vcenter $vSrc.Host `
                    --user    $vUser     `
                    --pass    $vPass     `
                    --days    90         `
                    --output  $sessionFolder
                $perfFile = Get-ChildItem $sessionFolder -Filter 'vsphere-perf*.json' -EA SilentlyContinue |
                            Sort-Object LastWriteTime -Descending | Select-Object -First 1
                if ($perfFile) {
                    [void]$outputFiles.Add($perfFile.FullName)
                    B-OK "Performance data: $($perfFile.Name)"
                    Write-Log 'OK' $vSrc.Host "vSphere perf JSON: $($perfFile.Name)"
                } else {
                    B-Warn "Perf collection ran but no vsphere-perf JSON found in $sessionFolder"
                }
            } catch {
                B-Err "vSphere perf collection failed for $($vSrc.Host): $_"
                Write-Log 'FAIL' $vSrc.Host "vSphere perf collection error: $_"
            }
        }
    }
}

# -----------------------------------------------------------------------------
# PHASE 6c - LINUX / APPLIANCE SSH DISCOVERY
# -----------------------------------------------------------------------------

if ($linuxRows.Count -gt 0) {
    Write-Phase "Linux / Appliance Discovery (SSH)"
    Write-Host ("  {0} Linux/appliance box(es) detected." -f $linuxRows.Count) -ForegroundColor Cyan
    Write-Host "  Each box prompts for its own SSH credentials." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  [Y]  Attempt SSH discovery on each box" -ForegroundColor Cyan
    Write-Host "  [N]  Skip (boxes still get placeholder tabs in the report)" -ForegroundColor Cyan
    Write-Host ""
    $linuxChoice = ''
    while ($linuxChoice -notin @('Y','N')) {
        $linuxChoice = (Read-Host "  Discover Linux? [Y/N]").Trim().ToUpper()
    }

    if ($linuxChoice -eq 'Y') {
        $plinkExe = Join-Path $PSScriptRoot 'plink.exe'
        if (-not (Test-Path $plinkExe)) {
            B-Warn "plink.exe not found. Run Get-PortablePython.ps1 to download it."
            B-Warn "Skipping SSH discovery — boxes will appear as placeholders."
        } else {
            $linuxDateStr = (Get-Date).ToString('yyyy-MM-dd')
            foreach ($lr in $linuxRows) {
                $lFile = Invoke-LinuxDiscovery -Row $lr -PlinkPath $plinkExe `
                             -OutputFolder $sessionFolder -DateStr $linuxDateStr
                if ($lFile) { [void]$outputFiles.Add($lFile) }
            }
        }
    } else {
        B-Line "Skipped. Linux boxes will appear as placeholder tabs in the report."
    }
    Write-Host ""
}

# -----------------------------------------------------------------------------
# CLEANUP + SUMMARY
# -----------------------------------------------------------------------------

Invoke-Cleanup

$elapsed = [math]::Round(((Get-Date) - $script:SessionStart).TotalSeconds, 0)

Write-Host ""
Write-Host ("=" * 72) -ForegroundColor DarkMagenta
Write-Host "  SESSION COMPLETE" -ForegroundColor Magenta
Write-Host ("=" * 72) -ForegroundColor DarkMagenta
Write-Host ("  Targets total  : " + $planRows.Count + "  (auto: $($autoRows.Count)  manual: $($manualRows.Count)  skip: $($skipRows.Count))") -ForegroundColor Gray
Write-Host ("  Succeeded      : $successCount") -ForegroundColor Green
if ($failCount -gt 0) { Write-Host ("  Failed         : $failCount") -ForegroundColor Red }
Write-Host ("  Runtime        : ${elapsed}s") -ForegroundColor DarkGray
Write-Host ("  Output folder  : $sessionFolder") -ForegroundColor White
Write-Host ("=" * 72) -ForegroundColor DarkMagenta

if ($outputFiles.Count -gt 0) {
    $dateStr = $script:SessionStart.ToString("yyyy-MM-dd")

    Write-Host ""
    Write-Host "  JSON files collected:" -ForegroundColor White
    $outputFiles | ForEach-Object { Write-Host ("    " + (Split-Path $_ -Leaf)) -ForegroundColor Cyan }

    # ── MANIFEST GENERATION ─────────────────────────────────────────────────
    Write-Host ""
    $clientShort = (Read-Host "  Client name (appears in report header — e.g. 'Acme Corporation')").Trim()
    if (-not $clientShort) { $clientShort = "CLIENT" }
    $clientFull  = $clientShort

    $inventoryFile = $outputFiles |
                     Where-Object { (Split-Path $_ -Leaf) -match '-inventory-' } |
                     Select-Object -First 1 |
                     ForEach-Object { Split-Path $_ -Leaf }

    $serverEntries = [System.Collections.ArrayList]@()

    # Windows discovery entries
    foreach ($f in ($outputFiles | Where-Object { (Split-Path $_ -Leaf) -match '-discovery-' })) {
        $leaf    = Split-Path $f -Leaf
        $srvName = [System.IO.Path]::GetFileNameWithoutExtension($leaf) -replace '-discovery-.*$', ''
        $srvId   = ($srvName.ToLower() -replace '[^a-z0-9]', '')
        $matchedRow = $autoRows | Where-Object {
            $_.DisplayName -ieq $srvName -or
            ($_.ResolvedHost -and $_.ResolvedHost -ieq $srvName)
        } | Select-Object -First 1
        $srvIP = if ($matchedRow -and $matchedRow.Address -and $matchedRow.Address -ne 'localhost') {
            $matchedRow.Address
        } else { '' }
        [void]$serverEntries.Add([ordered]@{
            id       = $srvId
            file     = $leaf
            name     = $srvName
            ip       = $srvIP
            in_scope = $true
        })
    }

    # Linux/appliance entries — SSH-discovered and placeholder-only
    foreach ($lr in $linuxRows) {
        $lName   = $lr.DisplayName
        $lId     = ($lName.ToLower() -replace '[^a-z0-9]', '') + 'lnx'
        $lIP     = $lr.Address
        $lGuest  = if ($lr.ResolvedOS -and $lr.ResolvedOS -ne 'Unknown') { $lr.ResolvedOS } else { 'Linux' }
        # Check if SSH discovery produced a file for this box
        $lFile   = $outputFiles | Where-Object { (Split-Path $_ -Leaf) -match "^$([regex]::Escape($lName))-linux-" } |
                   Select-Object -First 1
        $lLeaf   = if ($lFile) { Split-Path $lFile -Leaf } else { '' }
        [void]$serverEntries.Add([ordered]@{
            id       = $lId
            file     = $lLeaf
            name     = $lName
            ip       = $lIP
            in_scope = $true
            os_type  = 'linux'
            guest_os = $lGuest
        })
    }

    $manifest = [ordered]@{
        client         = $clientShort
        client_full    = $clientFull
        date           = $dateStr
        session_dir    = "."
        output_dir     = "."
        inventory_file = if ($inventoryFile) { $inventoryFile } else { "" }
        logo_file      = ""
        servers        = @($serverEntries)
    }

    $manifestFile = Join-Path $sessionFolder "$clientShort-manifest-$dateStr.json"
    $manifest | ConvertTo-Json -Depth 5 | Out-File $manifestFile -Encoding UTF8 -Force
    B-OK "Manifest: $(Split-Path $manifestFile -Leaf)"

    # ── AUTO-GENERATE HTML REPORT ────────────────────────────────────────────
    $genReportScript = Join-Path $PSScriptRoot 'gen_report.py'
    $portablePy      = Join-Path $PSScriptRoot 'python\python.exe'

    $reportPython = $null
    if (Test-Path $portablePy) {
        $reportPython = $portablePy
    } else {
        foreach ($pyCandidate in @('python','python3','py')) {
            try {
                $pyOut = & $pyCandidate --version 2>&1
                if ($pyOut -match 'Python [23]') { $reportPython = $pyCandidate; break }
            } catch { }
        }
    }

    Write-Host ""
    Write-Host ("=" * 72) -ForegroundColor DarkCyan

    if ($reportPython -and (Test-Path $genReportScript)) {
        Write-Host "  GENERATING HTML REPORT" -ForegroundColor Cyan
        Write-Host ("=" * 72) -ForegroundColor DarkCyan
        Write-Host ""
        Write-Host "  Role labels are auto-detected from installed roles and services." -ForegroundColor DarkGray
        Write-Host "  To exclude a server post-run: set  in_scope: false  in the manifest." -ForegroundColor DarkGray
        Write-Host ""
        try {
            & $reportPython $genReportScript $manifestFile
            Write-Host ""
            B-OK "Report saved to: $sessionFolder"
        } catch {
            B-Err "Report generation failed: $_"
            Write-Host ""
            Write-Host "  Run manually:" -ForegroundColor DarkGray
            Write-Host "  python `"$genReportScript`" `"$manifestFile`"" -ForegroundColor Green
        }
    } else {
        Write-Host "  NEXT STEP -- Generate HTML Report" -ForegroundColor Yellow
        Write-Host ("=" * 72) -ForegroundColor DarkCyan
        Write-Host ""
        Write-Host "  Python not found. Run Get-PortablePython.ps1 once to set up:" -ForegroundColor Yellow
        Write-Host "  .\Get-PortablePython.ps1" -ForegroundColor Green
        Write-Host ""
        Write-Host "  Then re-run, or generate the report manually:" -ForegroundColor DarkGray
        Write-Host "  python `"$genReportScript`" `"$manifestFile`"" -ForegroundColor Green
        Write-Host ""
        Write-Host "  HTML drops into: $sessionFolder" -ForegroundColor DarkGray
    }
    Write-Host ""
}

if ($manualRows.Count -gt 0) {
    Write-Host ""
    Write-Host "  MANUAL RUNS NEEDED - copy $invokeScriptName to these servers and run it locally:" -ForegroundColor Yellow
    foreach ($r in $manualRows) {
        Write-Host ("    $($r.DisplayName)  ($($r.SkipReason))") -ForegroundColor DarkYellow
    }
    Write-Host ""
    Write-Host "  On each: powershell.exe -ExecutionPolicy Bypass -File .\$invokeScriptName" -ForegroundColor DarkGray
    Write-Host "  Then copy the JSON back to: $sessionFolder" -ForegroundColor DarkGray
}

if ($script:SessionErrors.Count -gt 0) {
    Write-Host ""
    Write-Host "  Session errors:" -ForegroundColor DarkYellow
    $script:SessionErrors | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
}
Write-Host ""
