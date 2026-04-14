<#
.SYNOPSIS
    Magna5 Server Discovery - Session Launcher v1.4

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
$script:SessionVersion  = '1.6'
$script:SessionStart    = Get-Date
$script:WinRMRestoreMap = @{}
$script:SessionErrors   = [System.Collections.ArrayList]@()
$script:DomainCred      = $null

# -- COMPANION SOURCE PATHS ---------------------------------------------------
# Ordered list of locations to search for Invoke-ServerDiscovery when the
# local copy is missing or truncated (RDP/SMB clipboard truncation is common).
#
# Add a UNC path or corp PC share to let the launcher self-heal from one copy:
#   $env:M5_DISCOVERY_SOURCE = '\\CORPPC\C$\Tools\M5Discovery'
#   $env:M5_DISCOVERY_SOURCE = '\\fileserver\shared\M5Scripts'
#
# Set that env var permanently on your machine and you only ever need to
# transfer Start-DiscoverySession to a client — it fetches its companion.

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

$invokeScriptName = 'Invoke-ServerDiscovery_1.6.ps1'
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
            Write-Host "  (^_^)  Auto-copied to $PSScriptRoot — will be local next run." -ForegroundColor DarkGreen
        } catch {
            $DiscoveryScript = $found
            Write-Host "         (Could not write to script folder — using source path directly.)" -ForegroundColor DarkGray
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
        Write-Host "        Then re-run — this script will pull and cache it automatically." -ForegroundColor DarkGray
        exit 1
    }
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
    Write-Host "  (>_<)  PowerShell $($PSVersionTable.PSVersion) — limited compatibility." -ForegroundColor Yellow
    Write-Host "         Some WMI method calls may degrade. Recommend PS 5.1+." -ForegroundColor DarkYellow
    Write-Host ""
}
if ($PSMaj -ge 7) {
    Write-Host "  (^_^)  PowerShell $($PSVersionTable.PSVersion) — full CIM/WinRM compatibility." -ForegroundColor DarkGreen
} elseif ($PSMaj -ge 5) {
    Write-Host "  (^_^)  PowerShell $($PSVersionTable.PSVersion) — compatible." -ForegroundColor DarkGreen
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
function B-Err  { param([string]$m) {
    Write-Host "  (x_x)  $m" -ForegroundColor Red
    [void]$script:SessionErrors.Add($m) }}
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
Write-Host ("=" * 72) -ForegroundColor DarkMagenta
Write-Host "  MAGNA5 SERVER DISCOVERY - SESSION LAUNCHER  v$script:SessionVersion" -ForegroundColor Magenta
Write-Host ("=" * 72) -ForegroundColor DarkMagenta
Write-Host ("  Started  : " + $script:SessionStart.ToString("yyyy-MM-dd HH:mm:ss") + "  |  Host: $env:COMPUTERNAME") -ForegroundColor Gray
Write-Host ("  PS Ver   : " + $PSVersionTable.PSVersion.ToString()) -ForegroundColor Gray
Write-Host ("=" * 72) -ForegroundColor DarkMagenta
Write-Host ""

# -----------------------------------------------------------------------------
# HYPERVISOR HELPERS
# -----------------------------------------------------------------------------

# -- WMI / CIM COMPAT HELPERS -------------------------------------------------
# PS3+  : Get-CimInstance + Invoke-CimMethod (modern, preferred)
# PS2   : Get-WmiObject + direct .Method() calls (legacy fallback only)
# CimInstance objects have NO direct methods; always use Invoke-WmiOrCimMethod.

function Get-WmiOrCim {
    param(
        [string]$Class,
        [string]$Namespace    = 'root\cimv2',
        [string]$Filter       = '',
        [string]$ComputerName = '',
        [PSCredential]$Cred   = $null,
        [string]$EA           = 'Stop'
    )
    if ($PSMaj -ge 3) {
        $p = @{ ClassName=$Class; Namespace=$Namespace; ErrorAction=$EA }
        if ($Filter)       { $p.Filter       = $Filter }
        if ($ComputerName) { $p.ComputerName = $ComputerName }
        if ($Cred)         { $p.Credential   = $Cred }
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
    param(
        [string]$MethodName,
        [hashtable]$Arguments  = @{},
        [object]$Instance      = $null,
        [string]$ClassName     = '',
        [string]$ComputerName  = '',
        [PSCredential]$Cred    = $null,
        [string]$Namespace     = 'root\cimv2'
    )
    if ($PSMaj -ge 3) {
        if ($Instance) {
            Invoke-CimMethod -InputObject $Instance -MethodName $MethodName -Arguments $Arguments -ErrorAction Stop
        } else {
            $p = @{ ClassName=$ClassName; MethodName=$MethodName; Arguments=$Arguments; Namespace=$Namespace; ErrorAction='Stop' }
            if ($ComputerName) { $p.ComputerName = $ComputerName }
            if ($Cred)         { $p.Credential   = $Cred }
            Invoke-CimMethod @p
        }
    } else {
        # PS2 WMI fallback — direct method invocation on WmiObject
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
    foreach ($apiPath in @('/api/session', '/rest/com/vmware/cis/session')) {
        try {
            $uri   = "https://$Server$apiPath"
            $token = Invoke-VSphereRest -Uri $uri -Method POST -Headers $authHeader
            $apiVer = if ($apiPath -match '^/api') { 'v7' } else { 'v6' }
            B-OK "Connected to vSphere ($apiVer API) at $Server"
            return @{ OK=$true; Token=$token; Server=$Server; APIVer=$apiVer }
        } catch { }
    }
    return @{ OK=$false; Error="Could not connect to vSphere REST API on $Server. Check hostname and credentials." }
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

# -- Hyper-V enumeration via WMI -----------------------------------------------

function Get-HyperVVMs {
    param([string]$HVHost, [PSCredential]$Cred)
    try {
        $vms = Get-WmiOrCim -Class 'Msvm_ComputerSystem' -Namespace 'root\virtualization\v2' `
                            -Filter "Caption='Virtual Machine'" -ComputerName $HVHost -Cred $Cred -EA 'Stop'
        $list = [System.Collections.ArrayList]@()
        foreach ($vm in $vms) {
            $ip = ''
            # State: 2=Running, 3=Off, 9=Paused, 6=Saved
            $stateMap = @{ 2='Running'; 3='Off'; 9='Paused'; 6='Saved'; 10='Starting' }
            $state = if ($stateMap.ContainsKey([int]$vm.EnabledState)) { $stateMap[[int]$vm.EnabledState] } else { $vm.EnabledState }
            # Try to get IP from guest network adapter (requires Integration Services)
            try {
                $adapters = Get-WmiOrCim -Class 'Msvm_GuestNetworkAdapterConfiguration' `
                                         -Namespace 'root\virtualization\v2' `
                                         -ComputerName $HVHost -Cred $Cred -EA 'SilentlyContinue' |
                                         Where-Object { $_.InstanceID -match $vm.Name }
                if ($adapters) {
                    $ips = $adapters | ForEach-Object { $_.IPAddresses | Where-Object { $_ -notmatch ':' -and $_ -ne '0.0.0.0' } }
                    $ip = ($ips | Select-Object -First 2) -join ', '
                }
            } catch { }
            [void]$list.Add([PSCustomObject]@{
                Name       = $vm.ElementName
                IP         = $ip
                PowerState = $state
                GuestOS    = 'Windows (Hyper-V guest)'
                Source     = "Hyper-V ($HVHost)"
                VMID       = $vm.Name
            })
        }
        return $list
    } catch {
        B-Err "Hyper-V WMI query failed on $HVHost`: $_"
        return @()
    }
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
    B-Line "Ping sweeping $($IPs.Count) addresses (~15-30 seconds)..."
    $jobs = $IPs | ForEach-Object {
        $ip = $_
        Start-Job -ScriptBlock { param($h)
            [PSCustomObject]@{ IP=$h; Alive=(Test-Connection $h -Count 1 -Quiet -EA SilentlyContinue) }
        } -ArgumentList $ip -EA SilentlyContinue
    }
    $spin = 0; $spinC = @('|','/','-','\')
    $timeout = (Get-Date).AddSeconds(30)
    while (($jobs|Where-Object{$_.State -eq 'Running'}).Count -gt 0 -and (Get-Date) -lt $timeout) {
        $done = ($jobs|Where-Object{$_.State -ne 'Running'}).Count
        Write-Host ("`r  (o_o)  scanning... $($spinC[$spin%4])  $done/$($jobs.Count)   ") -NoNewline -ForegroundColor DarkCyan
        Start-Sleep -Milliseconds 250; $spin++
    }
    Write-Host "`r                                          " -NoNewline; Write-Host ""
    $results = $jobs | Wait-Job -Timeout 5 | Receive-Job -EA SilentlyContinue
    $jobs    | Remove-Job -Force -EA SilentlyContinue
    $live    = @($results | Where-Object { $_.Alive } | Select-Object -ExpandProperty IP)
    B-OK "$($live.Count) live hosts found"
    return $live
}

function Test-WMIAccess {
    param([string]$Target, [PSCredential]$Cred)
    try {
        $cs = Get-WmiOrCim -Class 'Win32_ComputerSystem' -ComputerName $Target -Cred $Cred -EA 'Stop'
        return @{ OK=$true; Hostname=$cs.Name; OS=$cs.Caption }
    } catch { return @{ OK=$false; Error=$_.Exception.Message } }
}

function Get-WinRMStateViaWMI {
    param([string]$Target, [PSCredential]$Cred)
    try {
        $svc = Get-WmiOrCim -Class 'Win32_Service' -Filter "Name='WinRM'" -ComputerName $Target -Cred $Cred -EA 'Stop'
        if (-not $svc) { return @{ OK=$false; Error='WinRM service not found' } }
        return @{ OK=$true; Running=($svc.State -eq 'Running'); StartMode=$svc.StartMode }
    } catch { return @{ OK=$false; Error=$_.Exception.Message } }
}

function Enable-WinRMViaWMI {
    param([string]$Target, [PSCredential]$Cred)
    try {
        $svc = Get-WmiOrCim -Class 'Win32_Service' -Filter "Name='WinRM'" -ComputerName $Target -Cred $Cred -EA 'Stop'
        Invoke-WmiOrCimMethod -Instance $svc -MethodName 'ChangeStartMode' -Arguments @{ StartMode='Automatic' } | Out-Null
        Invoke-WmiOrCimMethod -Instance $svc -MethodName 'StartService'    -Arguments @{} | Out-Null
        Start-Sleep -Seconds 3
        $psRemCmd = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Enable-PSRemoting -Force -SkipNetworkProfileCheck 2>&1|Out-Null"'
        Invoke-WmiOrCimMethod -ClassName 'Win32_Process' -ComputerName $Target -Cred $Cred `
                              -MethodName 'Create' -Arguments @{ CommandLine=$psRemCmd } | Out-Null
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
                              -MethodName 'Create' -Arguments @{ CommandLine=$cmd } | Out-Null
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
            $src.Host = Read-Safe "  Hyper-V host hostname or IP"
            Write-Host ""
            Write-Host "  HYPER-V HOST credentials:" -ForegroundColor Yellow
            Write-Host "    Is the Hyper-V HOST itself domain-joined, or does it use a local admin account?" -ForegroundColor Gray
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
                    Disconnect-VSphere -Conn $conn
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
                    Disconnect-VSphere -Conn $conn
                } else {
                    B-Err "  ESXi connection failed: $($conn.Error)"
                }
            }
            'HyperV' {
                $hvCred = if ($src.Cred) { $src.Cred } else { $script:DomainCred }
                $vms = Get-HyperVVMs -HVHost $src.Host -Cred $hvCred
                $vms | ForEach-Object { [void]$hypervisorVMs.Add($_) }
                B-OK "  Got $($vms.Count) VMs from Hyper-V $($src.Host)"
            }
        }
    }

    Write-Host ""
    if ($hypervisorVMs.Count -gt 0) {
        Write-Host ("  VMs from hypervisor(s) ({0} total):" -f $hypervisorVMs.Count) -ForegroundColor White
        Write-Host ("  {0,-28} {1,-12} {2,-18} {3}" -f "Name","State","IP","Source") -ForegroundColor DarkMagenta
        Write-Divider
        foreach ($vm in $hypervisorVMs) {
            $stateColor = if ($vm.PowerState -match 'Running|POWERED_ON') { 'Green' } else { 'DarkGray' }
            Write-Host ("  {0,-28} {1,-12} {2,-18} {3}" -f $vm.Name, $vm.PowerState, ($vm.IP -replace ',.*',''), $vm.Source) -ForegroundColor $stateColor
        }
        Write-Host ""
        Write-Host "  Only powered-on Windows VMs will be included. Use manual step to add/exclude." -ForegroundColor DarkGray
    }
}

# -----------------------------------------------------------------------------
# PHASE 3 - MANUAL TARGETS
# -----------------------------------------------------------------------------

Write-Phase "Step 4 - Manual Targets"
B-Line "physical servers, off-hypervisor boxes, or anything to add or exclude..."
Write-Host ""
Write-Host "  [1]  Add servers manually (hostnames or IPs)" -ForegroundColor White
Write-Host "  [2]  Subnet scan (find live hosts on a subnet)" -ForegroundColor White
Write-Host "  [3]  Load from text file (one per line)" -ForegroundColor White
Write-Host "  [4]  Skip - use hypervisor list only" -ForegroundColor White
Write-Host ""

$manualTargets = [System.Collections.ArrayList]@()
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
        $cidr  = Read-Safe "  CIDR (e.g. 10.0.1.0/24)"
        $allIPs = Get-SubnetHosts -CIDR $cidr
        if ($allIPs.Count -gt 0) {
            $liveIPs = Invoke-PingSweep -IPs $allIPs
            Write-Host ""; Write-Host "  Live hosts:" -ForegroundColor Gray
            $liveIPs | ForEach-Object { Write-Host "    $_" -ForegroundColor White }
            Write-Host ""
            $excl = Read-Safe "  Exclude any? (comma-separated IPs, or Enter for none)"
            $excludes = if ($excl) { $excl -split ',' | ForEach-Object { $_.Trim() } } else { @() }
            $liveIPs | Where-Object { $_ -notin $excludes } | ForEach-Object {
                [void]$manualTargets.Add([PSCustomObject]@{ Name=$_; IP=$_; PowerState='Unknown'; GuestOS='Unknown'; Source='Manual (scan)' })
            }
        }
    }
    '3' {
        $fp = Read-Safe "  File path".Trim('"')
        if (Test-Path $fp) {
            Get-Content $fp | Where-Object { $_.Trim() -and -not $_.StartsWith('#') } | ForEach-Object {
                [void]$manualTargets.Add([PSCustomObject]@{ Name=$_.Trim(); IP=$_.Trim(); PowerState='Unknown'; GuestOS='Unknown'; Source='File' })
            }
        } else { B-Err "File not found: $fp" }
    }
    '4' { B-Line "skipping manual entry - using hypervisor list only" }
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
        Write-Host " WMI FAILED" -ForegroundColor Red
        $row.Ping   = $true
        $row.Action = 'MANUAL'; $row.SkipReason = "WMI failed - $($wmi.Error -replace '.{0,60}$','')"
        [void]$planRows.Add($row); continue
    }
    $row.WMI          = $true
    $row.ResolvedHost = $wmi.Hostname
    $row.ResolvedOS   = $wmi.OS

    # WinRM state
    $ws = Get-WinRMStateViaWMI -Target $addr -Cred $script:DomainCred
    if ($ws.OK) {
        $row.WinRMState = if ($ws.Running) { 'ON' } else { "OFF ($($ws.StartMode))" }
        $row.ConnectMethod = if ($ws.Running) { 'WinRM (already on)' } else { 'WinRM (WMI bootstrap)' }
    } else {
        $row.WinRMState    = 'Unknown'
        $row.ConnectMethod = 'WinRM (will attempt)'
    }
    $row.Action = 'AUTO'
    Write-Host (" OK  |  $($row.ResolvedHost)  |  WinRM: $($row.WinRMState)") -ForegroundColor Green
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

Write-Host ""
Write-Host ""
Write-Host ("=" * 72) -ForegroundColor DarkMagenta
Write-Host "  DISCOVERY PLAN - Review this carefully before running" -ForegroundColor Magenta
Write-Host ("=" * 72) -ForegroundColor DarkMagenta
Write-Host ""

# Header
Write-Host ("  {0,-26} {1,-14} {2,-22} {3,-20} {4}" -f "Server","Source","OS","Connect Method","Action") -ForegroundColor DarkMagenta
Write-Divider

foreach ($r in $planRows) {
    $osDisplay  = if ($r.ResolvedOS) { $r.ResolvedOS } elseif ($r.GuestOSHint -and $r.GuestOSHint -ne 'Unknown') { $r.GuestOSHint + ' (hint)' } else { '(unknown)' }
    $osShort    = if ($osDisplay.Length -gt 21) { $osDisplay.Substring(0,19) + '..' } else { $osDisplay }
    $srcShort   = if ($r.Source.Length -gt 13)  { $r.Source.Substring(0,11) + '..' } else { $r.Source }
    $methShort  = if ($r.ConnectMethod) { $r.ConnectMethod } elseif ($r.SkipReason) { $r.SkipReason } else { '-' }
    if ($methShort.Length -gt 19) { $methShort = $methShort.Substring(0,17) + '..' }

    $actionLabel = switch ($r.Action) {
        'AUTO'   { '[AUTO]  ' }
        'MANUAL' { '[MANUAL]' }
        'SKIP'   { '[SKIP]  ' }
    }
    $color = switch ($r.Action) {
        'AUTO'   { 'White'    }
        'MANUAL' { 'Yellow'   }
        'SKIP'   { 'DarkGray' }
    }
    Write-Host ("  {0,-26} {1,-14} {2,-22} {3,-20} {4}" -f $r.DisplayName, $srcShort, $osShort, $methShort, $actionLabel) -ForegroundColor $color
}

Write-Divider
Write-Host ""
Write-Host ("  AUTO   ({0,2}) - Will run without you touching anything" -f $autoRows.Count) -ForegroundColor Green
if ($manualRows.Count -gt 0) {
    Write-Host ("  MANUAL ({0,2}) - Unreachable; you must run the script locally on these:" -f $manualRows.Count) -ForegroundColor Yellow
    foreach ($r in $manualRows) {
        Write-Host ("           $($r.DisplayName)  -  $($r.SkipReason)") -ForegroundColor DarkYellow
    }
}
if ($skipRows.Count -gt 0) {
    Write-Host ("  SKIP   ({0,2}) - Powered off or excluded" -f $skipRows.Count) -ForegroundColor DarkGray
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
Write-Host "  [Y]  Yes - run discovery on all AUTO targets now" -ForegroundColor Green
Write-Host "  [N]  No  - abort (nothing has been changed on any server)" -ForegroundColor Red
Write-Host "  [R]  Re-probe - go back and re-test targets" -ForegroundColor Cyan
Write-Host ""

$goAns = ''
while ($goAns -notin @('Y','N','R')) { $goAns = Read-Safe "  Ready to run? [Y/N/R]".ToUpper() }

if ($goAns -eq 'N') {
    B-Line "Aborted. Nothing was changed on any server."; Invoke-Cleanup; exit 0
}
if ($goAns -eq 'R') {
    B-Line "Re-running the probe phase - restart the script to go through setup again."
    Invoke-Cleanup; exit 0
}

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
    Write-Host ""
    Write-Host "  JSON files ready for report:" -ForegroundColor White
    $outputFiles | ForEach-Object { Write-Host ("    " + (Split-Path $_ -Leaf)) -ForegroundColor Cyan }
    Write-Host ""
    Write-Host "  (^_^)>  Send those JSON files to Claude:" -ForegroundColor DarkCyan
    Write-Host "          'Here are my server discovery JSONs - build the HTML report'" -ForegroundColor DarkGray
}

if ($manualRows.Count -gt 0) {
    Write-Host ""
    Write-Host "  MANUAL RUNS NEEDED - copy Invoke-ServerDiscovery_1.4.ps1 to these servers:" -ForegroundColor Yellow
    foreach ($r in $manualRows) {
        Write-Host ("    $($r.DisplayName)  ($($r.SkipReason))") -ForegroundColor DarkYellow
    }
    Write-Host ""
    Write-Host "  On each: powershell.exe -ExecutionPolicy Bypass -File .\Invoke-ServerDiscovery_1.4.ps1" -ForegroundColor DarkGray
    Write-Host "  Then copy the JSON back to: $sessionFolder" -ForegroundColor DarkGray
}

if ($script:SessionErrors.Count -gt 0) {
    Write-Host ""
    Write-Host "  Session errors:" -ForegroundColor DarkYellow
    $script:SessionErrors | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
}
Write-Host ""
