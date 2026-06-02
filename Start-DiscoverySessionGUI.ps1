<#
.SYNOPSIS
    Start-DiscoverySessionGUI.ps1 - Browser-based UI for SDT discovery sessions.

.DESCRIPTION
    Spins up a local HTTP listener on http://localhost:8080, serves a
    self-contained HTML wizard, opens the user's default browser.

    The wizard collects client info, hypervisor target, server list, and
    credentials, then runs discovery with live progress updates.

    Reuses existing Invoke-ServerDiscovery.ps1 per target and gen_report.py
    for the final HTML report. All credentials live in PowerShell memory
    only - never written to disk.

    Console-mode Start-DiscoverySession.ps1 is untouched. This script is
    additive.

.NOTES
    v4.0-alpha  |  2026-04-21
    Requires: PowerShell 5.1+
#>
param(
    [int]    $Port = 8080,
    [switch] $NoOpenBrowser
)

$ErrorActionPreference = 'Stop'

# Resolve version dynamically. Priority:
#   1. VERSION file at ../VERSION (written by install.ps1 at %LOCALAPPDATA%\Magna5\SDT\VERSION)
#   2. VERSION file alongside the script (zero-install / dev mode)
#   3. Containing folder name like 'SDT-4.1.20' (zip-extract default)
#   4. Hardcoded fallback (kept in sync with the most recent tag)
function Get-SdtVersion {
    $candidates = @(
        (Join-Path (Split-Path $PSScriptRoot -Parent) 'VERSION'),
        (Join-Path $PSScriptRoot 'VERSION')
    )
    foreach ($p in $candidates) {
        if ($p -and (Test-Path $p)) {
            try {
                $v = (Get-Content $p -Raw -EA Stop).Trim()
                if ($v) { return ($v -replace '^v','') }
            } catch { }
        }
    }
    # Folder-name probe: 'SDT-4.1.20' or 'sdt-4.1.20'
    $leaf = Split-Path $PSScriptRoot -Leaf
    if ($leaf -match '^[Ss][Dd][Tt][-_]v?(\d+\.\d+\.\d+)') { return $matches[1] }
    # Last-ditch fallback - this gets bumped at every release
    return '4.1.20'
}
$script:Version   = Get-SdtVersion
$script:ScriptDir = $PSScriptRoot
$script:BaseUrl   = "http://localhost:$Port"

# Session state shared between HTTP handlers and discovery worker jobs.
$script:Session = [hashtable]::Synchronized(@{
    Status     = 'idle'                       # idle | running | complete | error
    Client     = ''
    OutputDir  = ''
    SessionDir = ''
    Targets    = @()                          # list of { Name; Address; State; Phase; Buddy; Started; Finished }
    LogTail    = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]@())
    ReportPath = ''
    ReportZipPath = ''
    MissingTargets = @()
    StartedAt  = $null
    FinishedAt = $null
})

$script:BuddyFrames = @('(^_^) ','(^_^)>','(o_o) ','(o_o)>','(-_-) ','(>_<) ','(*_*) ','(^_-) ','(._.) ','(T_T) ','(^o^) ','(x_x) ')

function Add-Log([string]$msg) {
    $stamp = (Get-Date).ToString('HH:mm:ss')
    [void]$script:Session.LogTail.Add("[$stamp] $msg")
    # Keep only last 2000 lines (higher now that parallel scans stream live)
    while ($script:Session.LogTail.Count -gt 2000) {
        $script:Session.LogTail.RemoveAt(0)
    }
}

# Runs collect_vsphere_perf.py with auto-retry across username formats.
# Returns hashtable: @{ ok=$true/$false; user=<succeeded-format>; log=<combined-log>; file=<json-path>; error=<short>; }
function Test-IsLocalHyperVHost {
    # GUI-side mirror of Test-IsHyperVHost in Start-DiscoverySession.ps1.
    # Returns {IsHost, PassCount, Reasons[], Detail{}} -- 2-of-4 quorum.
    [CmdletBinding()] param()
    $sig = [ordered]@{ vmms_service=$false; hyperv_feature=$false; hypervisor_present=$false; get_vm_cmdlet=$false }
    $reasons = New-Object System.Collections.ArrayList

    try {
        $svc = Get-Service -Name vmms -ErrorAction Stop
        $sig.vmms_service = ($svc.Status -eq 'Running')
        if (-not $sig.vmms_service) { [void]$reasons.Add("vmms service is '$($svc.Status)' (Start-Service vmms)") }
    } catch { [void]$reasons.Add("vmms service not found -- Hyper-V role not installed") }

    try {
        if (Get-Command Get-WindowsFeature -EA SilentlyContinue) {
            $f = Get-WindowsFeature -Name Hyper-V -EA SilentlyContinue
            if ($f) { $sig.hyperv_feature = ($f.Installed -eq $true) }
        } elseif (Get-Command Get-WindowsOptionalFeature -EA SilentlyContinue) {
            $f = Get-WindowsOptionalFeature -FeatureName Microsoft-Hyper-V -Online -EA SilentlyContinue
            if ($f) { $sig.hyperv_feature = ($f.State -eq 'Enabled') }
        }
        if (-not $sig.hyperv_feature) { [void]$reasons.Add("Hyper-V Windows feature not enabled") }
    } catch { [void]$reasons.Add("could not query Windows features: $($_.Exception.Message)") }

    try {
        $cs = Get-CimInstance Win32_ComputerSystem -EA Stop
        $sig.hypervisor_present = [bool]$cs.HypervisorPresent
        if (-not $sig.hypervisor_present) { [void]$reasons.Add("HypervisorPresent=false (CPU virt may be off in BIOS)") }
    } catch { [void]$reasons.Add("could not query Win32_ComputerSystem") }

    $sig.get_vm_cmdlet = [bool](Get-Command Get-VM -EA SilentlyContinue | Where-Object { $_.ModuleName -eq 'Hyper-V' })
    if (-not $sig.get_vm_cmdlet) { [void]$reasons.Add("Get-VM (Hyper-V module) missing -- Install-WindowsFeature Hyper-V-PowerShell") }

    $pass = ($sig.Values | Where-Object { $_ -eq $true }).Count
    [pscustomobject]@{ IsHost=($pass -ge 2); PassCount=$pass; Reasons=$reasons; Detail=$sig }
}

function Invoke-LocalHyperVInventory {
    # Runs Get-VM / Get-VMHost / Get-VMSwitch locally. Returns HyperVInventory
    # object (same shape as Start-DiscoverySession.ps1 produces) + a flattened
    # VM list for the UI. Self-heals: if Get-VM fails, falls back to WMI.
    param([string]$OutputDir, [switch]$Force)

    $pf = Test-IsLocalHyperVHost
    if (-not $pf.IsHost -and -not $Force) {
        return @{ ok=$false; preflight=$pf; error=("Not a Hyper-V host: " + ($pf.Reasons -join '; ')) }
    }

    $vms = @()
    $invErr = ''
    $hvInv = $null
    try {
        $hostCS  = Get-CimInstance Win32_ComputerSystem -EA Stop
        $hostCPU = (Get-CimInstance Win32_Processor -EA SilentlyContinue | Select-Object -First 1)
        $vols    = Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' -EA SilentlyContinue | ForEach-Object {
            [pscustomobject]@{
                Drive   = $_.DeviceID
                Label   = $_.VolumeName
                TotalGB = if ($_.Size) { [math]::Round($_.Size/1GB,1) } else { 0 }
                FreeGB  = if ($_.FreeSpace) { [math]::Round($_.FreeSpace/1GB,1) } else { 0 }
                UsedPct = if ($_.Size -gt 0) { [math]::Round(100 - ($_.FreeSpace/$_.Size*100),1) } else { 0 }
            }
        }

        $switches = @()
        try {
            $switches = Get-VMSwitch -EA Stop | ForEach-Object {
                [pscustomobject]@{ Name=$_.Name; Type=$_.SwitchType.ToString(); NetAdapter=$_.NetAdapterInterfaceDescription }
            }
        } catch { Add-Log "Get-VMSwitch failed: $($_.Exception.Message)" }

        # Primary: Get-VM (Hyper-V module). Fallback: WMI Msvm_ComputerSystem.
        $rawVms = $null
        try { $rawVms = Get-VM -EA Stop } catch {
            Add-Log "Get-VM failed ($($_.Exception.Message)) -- trying WMI fallback"
            try {
                $rawVms = Get-CimInstance -Namespace 'root\virtualization\v2' -ClassName Msvm_ComputerSystem `
                    -Filter "Caption='Virtual Machine'" -EA Stop
            } catch { $invErr = "Both Get-VM and WMI failed: $($_.Exception.Message)" }
        }

        $vmList = @()
        if ($rawVms) {
            foreach ($vm in $rawVms) {
                $isWmi = ($vm.PSObject.TypeNames -join ',') -match 'Msvm_ComputerSystem'
                if ($isWmi) {
                    $stateMap = @{ 2='Running'; 3='Off'; 9='Paused'; 6='Saved'; 10='Starting' }
                    $state = if ($stateMap.ContainsKey([int]$vm.EnabledState)) { $stateMap[[int]$vm.EnabledState] } else { "$($vm.EnabledState)" }
                    $vmList += [pscustomobject]@{
                        Name=$vm.ElementName; State=$state; vCPU=$null; RAMgb=$null
                        UptimeHours=$null; Snapshots=0; IPs=''; Disks=@(); NetworkAdapters=@()
                        IntegrationServices=''
                    }
                } else {
                    $ads = $null; $ips = @(); $disks = @(); $nets = @(); $snaps = 0; $intSvc = ''
                    try { $ads = Get-VMNetworkAdapter -VM $vm -EA SilentlyContinue } catch {}
                    if ($ads) {
                        $ips = $ads | ForEach-Object { $_.IPAddresses } | Where-Object { $_ -and ($_ -notmatch ':') -and $_ -ne '0.0.0.0' }
                        $nets = $ads | ForEach-Object {
                            [pscustomobject]@{
                                Name=$_.Name; SwitchName=$_.SwitchName; MacAddress=$_.MacAddress
                                IPs=(($_.IPAddresses | Where-Object { $_ -notmatch ':' -and $_ -ne '0.0.0.0' }) -join ', ')
                            }
                        }
                    }
                    try {
                        $disks = Get-VMHardDiskDrive -VM $vm -EA SilentlyContinue | ForEach-Object {
                            $vhd = Get-VHD $_.Path -EA SilentlyContinue
                            [pscustomobject]@{
                                Path=$_.Path; ControllerType="$($_.ControllerType)"
                                SizeGB  = if ($vhd) { [math]::Round($vhd.Size/1GB,1) } else { $null }
                                UsedGB  = if ($vhd) { [math]::Round($vhd.FileSize/1GB,1) } else { $null }
                                VHDType = if ($vhd) { "$($vhd.VhdType)" } else { $null }
                            }
                        }
                    } catch {}
                    try { $snaps = @(Get-VMSnapshot -VM $vm -EA SilentlyContinue).Count } catch {}
                    try {
                        $intSvc = (Get-VMIntegrationService -VM $vm -EA SilentlyContinue |
                                   Where-Object { $_.Enabled } | Select-Object -ExpandProperty Name) -join ', '
                    } catch {}
                    $vmList += [pscustomobject]@{
                        Name=$vm.Name; State="$($vm.State)"; vCPU=$vm.ProcessorCount
                        RAMgb=[math]::Round($vm.MemoryAssigned/1GB,2)
                        UptimeHours=[math]::Round($vm.Uptime.TotalHours,1)
                        Snapshots=$snaps
                        IPs=(($ips | Select-Object -First 4) -join ', ')
                        Disks=$disks; NetworkAdapters=$nets; IntegrationServices=$intSvc
                    }
                }
            }
        }

        $hvInv = [pscustomobject]@{
            _type='HyperVInventory'
            HVHost=$env:COMPUTERNAME
            HostSummary=[pscustomobject]@{
                Manufacturer=$hostCS.Manufacturer; Model=$hostCS.Model
                TotalRAMgb=[math]::Round($hostCS.TotalPhysicalMemory/1GB,1)
                CPUModel    = $(if ($hostCPU) { $hostCPU.Name.Trim() } else { '' })
                CPUCores    = $(if ($hostCPU) { $hostCPU.NumberOfCores } else { 0 })
                CPULogical  = $(if ($hostCPU) { $hostCPU.NumberOfLogicalProcessors } else { 0 })
                Volumes=$vols
            }
            VirtualSwitches=$switches
            VMs=$vmList
            Preflight=$pf
        }

        if ($OutputDir) {
            try {
                if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
                $outFile = Join-Path $OutputDir ("{0}-hv-inventory-{1}.json" -f $env:COMPUTERNAME, (Get-Date -f 'yyyy-MM-dd'))
                $hvInv | ConvertTo-Json -Depth 8 | Set-Content -Path $outFile -Encoding UTF8
                return @{ ok=$true; inventory=$hvInv; vms=$vmList; file=$outFile; preflight=$pf }
            } catch { Add-Log "Failed to write HV inventory: $($_.Exception.Message)" }
        }
        return @{ ok=$true; inventory=$hvInv; vms=$vmList; file=''; preflight=$pf }
    } catch {
        return @{ ok=$false; error=("Local Hyper-V inventory failed: " + $_.Exception.Message); preflight=$pf }
    }
}

function Test-PyImport {
    # Returns hashtable @{ ok=$bool; diag=<stderr text if fail> }
    # PS5 Start-Process -ArgumentList does NOT quote args with spaces, so
    # passing @('-c','import requests') makes python see "-c import requests"
    # and chokes on a bare "import". Workaround: write the import statement
    # to a temp .py file and run python against the file path.
    param([string]$PyExe, [string]$Module)
    $tmpPy = [System.IO.Path]::GetTempFileName() + '.py'
    $outFile = [System.IO.Path]::GetTempFileName()
    $errFile = [System.IO.Path]::GetTempFileName()
    try {
        [System.IO.File]::WriteAllText($tmpPy, "import $Module`r`n", [System.Text.Encoding]::ASCII)
        $p = Start-Process -FilePath $PyExe -ArgumentList @($tmpPy) `
            -NoNewWindow -PassThru -Wait `
            -RedirectStandardOutput $outFile -RedirectStandardError $errFile -EA Stop
        $ok = ($p.ExitCode -eq 0)
        $diag = ''
        if (-not $ok) {
            $stderr = if (Test-Path $errFile) { (Get-Content $errFile -Raw) } else { '' }
            $diag = "import ${Module} failed (exit=$($p.ExitCode)): $stderr"
        }
        return @{ ok=$ok; diag=$diag }
    } catch {
        return @{ ok=$false; diag="Start-Process failed for import ${Module}: $($_.Exception.Message)" }
    }
    finally {
        Remove-Item $tmpPy, $outFile, $errFile -EA SilentlyContinue
    }
}

function Ensure-PyDeps {
    # Self-heal portable Python: install requests + pyVmomi + urllib3 if any missing.
    # Safety net in case install.ps1's pip bootstrap was skipped or failed.
    # Returns @{ ok=$true/$false; log=<combined>; missing=@(...) }
    param([string]$PyExe, [string]$ScriptDir)

    $required = @('requests','pyVmomi','urllib3')
    $missing = @()
    foreach ($m in $required) {
        $r = Test-PyImport -PyExe $PyExe -Module $m
        if (-not $r.ok) { $missing += $m }
    }
    if ($missing.Count -eq 0) { return @{ ok=$true; log=''; missing=@() } }

    Add-Log "Missing Python deps: $($missing -join ', '). Bootstrapping pip + installing..."
    $log = "Missing: $($missing -join ', ')`n"
    $pyDir = Split-Path $PyExe -Parent

    # Step 1: Fix ._pth so embeddable Python actually picks up site-packages.
    #
    # ROOT CAUSE FROM PRIOR FAILURE: embeddable Python's ._pth file OVERRIDES
    # PYTHONPATH entirely. Just uncommenting 'import site' doesn't add
    # Lib\site-packages to sys.path reliably across all builds. pip happily
    # installs to Lib\site-packages and then 'import requests' fails because
    # that dir isn't on sys.path.
    #
    # Fix: ensure ._pth contains BOTH 'import site' AND a literal
    # 'Lib\site-packages' line. That guarantees pip's install target is on
    # sys.path regardless of site.py's behavior.
    try {
        $pthFile = Get-ChildItem $pyDir -Filter 'python*._pth' -EA 0 | Select-Object -First 1
        if ($pthFile) {
            $pthContent = Get-Content $pthFile.FullName -Raw
            $modified = $false
            # Uncomment 'import site' if commented
            if ($pthContent -match '(?m)^\s*#\s*import\s+site\s*$') {
                $pthContent = $pthContent -replace '(?m)^\s*#\s*import\s+site\s*$','import site'
                $modified = $true
                $log += "Enabled 'import site' in $($pthFile.Name)`n"
            } elseif ($pthContent -notmatch '(?m)^\s*import\s+site\s*$') {
                # Not present at all - append it
                $pthContent = $pthContent.TrimEnd() + "`r`nimport site`r`n"
                $modified = $true
                $log += "Appended 'import site' to $($pthFile.Name)`n"
            }
            # Ensure 'Lib\site-packages' is on sys.path
            if ($pthContent -notmatch '(?im)^\s*Lib\\site-packages\s*$') {
                $pthContent = $pthContent.TrimEnd() + "`r`nLib\site-packages`r`n"
                $modified = $true
                $log += "Added 'Lib\site-packages' to $($pthFile.Name)`n"
            }
            if ($modified) {
                # Write WITHOUT BOM (ASCII encoding) - Python embeddable launcher
                # parses ._pth byte-by-byte and BOM characters break it.
                [System.IO.File]::WriteAllText($pthFile.FullName, $pthContent, (New-Object System.Text.ASCIIEncoding))
            }
        }
    } catch { $log += "._pth tweak failed: $($_.Exception.Message)`n" }

    # Step 2: bootstrap pip via bundled get-pip.py, fall back to network mirrors
    $getPipPy = Join-Path $pyDir 'get-pip.py'
    if (-not (Test-Path $getPipPy)) {
        $bundled = Join-Path $ScriptDir 'get-pip.py'
        if (Test-Path $bundled) {
            Copy-Item $bundled $getPipPy -Force
            $log += "Using bundled get-pip.py`n"
        } else {
            $mirrors = @(
                'https://bootstrap.pypa.io/get-pip.py',
                'https://raw.githubusercontent.com/pypa/get-pip/main/public/get-pip.py'
            )
            foreach ($url in $mirrors) {
                try {
                    Invoke-WebRequest -Uri $url -OutFile $getPipPy -UseBasicParsing -TimeoutSec 30 -EA Stop
                    if ((Get-Item $getPipPy).Length -gt 100000) {
                        $log += "Downloaded get-pip.py from $url`n"
                        break
                    }
                } catch { $log += "Mirror $url failed: $($_.Exception.Message)`n" }
            }
        }
    }

    if (-not (Test-Path $getPipPy)) {
        return @{ ok=$false; log=$log + "Could not obtain get-pip.py"; missing=$missing }
    }

    # Step 2: run get-pip.py to install pip into the embeddable Python.
    # NOTE: we do NOT gate on Test-PyImport 'pip' - embeddable Python's ._pth
    # quirk means `import pip` can fail even after a successful install
    # (site-packages not picked up reliably until next invocation). We just run
    # get-pip.py unconditionally - it's idempotent and fast if pip is already
    # there. Cost ~2 sec, benefit: no false-negative gate.
    try {
        $stdOut = [System.IO.Path]::GetTempFileName()
        $stdErr = [System.IO.Path]::GetTempFileName()
        $p = Start-Process -FilePath $PyExe -ArgumentList @($getPipPy,'--disable-pip-version-check') `
            -NoNewWindow -PassThru -Wait -RedirectStandardOutput $stdOut -RedirectStandardError $stdErr -EA Stop
        $log += "get-pip.py exit=$($p.ExitCode)`n"
        $log += (Get-Content $stdOut -Raw) + (Get-Content $stdErr -Raw)
        Remove-Item $stdOut, $stdErr -EA SilentlyContinue
    } catch { $log += "get-pip.py failed: $($_.Exception.Message)`n" }

    # Belt-and-suspenders: force Lib\site-packages onto sys.path so subsequent
    # python invocations can find pip even if ._pth wasn't updated correctly.
    $sitePackages = Join-Path $pyDir 'Lib\site-packages'
    $prevPyPath = $env:PYTHONPATH
    if (Test-Path $sitePackages) {
        if ($env:PYTHONPATH) { $env:PYTHONPATH = "$sitePackages;$env:PYTHONPATH" }
        else                 { $env:PYTHONPATH = $sitePackages }
        $log += "Set PYTHONPATH to include $sitePackages`n"
    }

    # Step 3: pip install required packages. Run unconditionally - if pip isn't
    # available, python will say so in stderr and we capture it.
    $pipArgs = @('-m','pip','install','--disable-pip-version-check') + $required
    try {
        $stdOut = [System.IO.Path]::GetTempFileName()
        $stdErr = [System.IO.Path]::GetTempFileName()
        $p = Start-Process -FilePath $PyExe -ArgumentList $pipArgs `
            -NoNewWindow -PassThru -Wait -RedirectStandardOutput $stdOut -RedirectStandardError $stdErr -EA Stop
        $log += "pip install exit=$($p.ExitCode)`n"
        $log += (Get-Content $stdOut -Raw) + (Get-Content $stdErr -Raw)
        Remove-Item $stdOut, $stdErr -EA SilentlyContinue
    } catch { $log += "pip install failed: $($_.Exception.Message)`n" }

    # Re-check the actual required modules (NOT pip - these are what we care
    # about). If they import, we're good.
    $stillMissing = @()
    foreach ($m in $required) {
        $r = Test-PyImport -PyExe $PyExe -Module $m
        if (-not $r.ok) {
            $stillMissing += $m
            if ($r.diag) { $log += "Post-install check: $($r.diag)`n" }
        }
    }
    if ($stillMissing.Count -eq 0) {
        Add-Log "Python deps installed OK."
        return @{ ok=$true; log=$log; missing=@() }
    }

    # FALLBACK LAYER 2: pip install succeeded but imports still fail. Most
    # common cause: site-packages not on sys.path. Dump diagnostics so we know
    # what python actually sees, then try the wheel-direct path.
    # IMPORTANT: Use a temp .py file because PS5 Start-Process -ArgumentList
    # doesn't quote args containing spaces. Passing -c "code with spaces"
    # makes python see "-c code with spaces" unquoted and breaks.
    $log += "`n=== Diagnostic dump (sys.path / site-packages contents) ===`n"
    try {
        $diagOut = [System.IO.Path]::GetTempFileName()
        $diagErr = [System.IO.Path]::GetTempFileName()
        $diagPy  = [System.IO.Path]::GetTempFileName() + '.py'
        $diagCode = @'
import sys, os
print("--- sys.path ---")
for p in sys.path: print(p)
sp = os.path.join(os.path.dirname(sys.executable), "Lib", "site-packages")
print("--- site-packages exists:", os.path.isdir(sp))
print("--- contents:")
if os.path.isdir(sp):
    for x in os.listdir(sp): print(" ", x)
'@
        [System.IO.File]::WriteAllText($diagPy, $diagCode, [System.Text.Encoding]::ASCII)
        $p = Start-Process -FilePath $PyExe -ArgumentList @($diagPy) `
            -NoNewWindow -PassThru -Wait -RedirectStandardOutput $diagOut -RedirectStandardError $diagErr -EA Stop
        $log += (Get-Content $diagOut -Raw) + (Get-Content $diagErr -Raw)
        Remove-Item $diagOut, $diagErr, $diagPy -EA SilentlyContinue
    } catch { $log += "Diagnostic dump failed: $($_.Exception.Message)`n" }

    # FALLBACK LAYER 3: wheel-direct install. Download pure-python wheels from
    # PyPI's JSON API and extract directly into Lib\site-packages. Works even
    # if pip itself is broken or sys.path is wrong.
    $log += "`n=== Fallback: direct wheel install ===`n"
    Add-Log "pip install didn't make modules importable. Trying wheel-direct fallback..."
    Add-Type -AssemblyName System.IO.Compression.FileSystem -EA SilentlyContinue
    if (-not (Test-Path $sitePackages)) {
        try { New-Item -ItemType Directory -Path $sitePackages -Force | Out-Null } catch { }
    }
    foreach ($pkg in $stillMissing) {
        try {
            $log += "Downloading wheel for $pkg from PyPI...`n"
            $info = Invoke-RestMethod "https://pypi.org/pypi/$pkg/json" -UseBasicParsing -TimeoutSec 20
            $version = $info.info.version
            $files = $info.releases.$version
            # Prefer py3-none-any; fall back to cp312-win_amd64 then cp311
            $wheel = $files | Where-Object { $_.packagetype -eq 'bdist_wheel' -and $_.filename -match 'py3-none-any\.whl$' } | Select-Object -First 1
            if (-not $wheel) {
                $wheel = $files | Where-Object { $_.packagetype -eq 'bdist_wheel' -and $_.filename -match 'cp312.*win_amd64\.whl$' } | Select-Object -First 1
            }
            if (-not $wheel) {
                $wheel = $files | Where-Object { $_.packagetype -eq 'bdist_wheel' -and $_.filename -match 'cp311.*win_amd64\.whl$' } | Select-Object -First 1
            }
            if (-not $wheel) {
                $log += "  No compatible wheel for $pkg (need py3-none-any or cp31x win_amd64)`n"
                continue
            }
            $whlPath = Join-Path $env:TEMP $wheel.filename
            Invoke-WebRequest $wheel.url -OutFile $whlPath -UseBasicParsing -TimeoutSec 60
            # Extract into site-packages (wheel is a zip)
            try {
                [System.IO.Compression.ZipFile]::ExtractToDirectory($whlPath, $sitePackages)
            } catch [System.IO.IOException] {
                # File already exists - extract individual entries, overwriting
                $zip = [System.IO.Compression.ZipFile]::OpenRead($whlPath)
                foreach ($entry in $zip.Entries) {
                    $target = Join-Path $sitePackages $entry.FullName
                    $targetDir = Split-Path $target -Parent
                    if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
                    if (-not $entry.FullName.EndsWith('/')) {
                        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $target, $true)
                    }
                }
                $zip.Dispose()
            }
            Remove-Item $whlPath -EA SilentlyContinue
            $log += "  Extracted $($wheel.filename) -> $sitePackages`n"
        } catch {
            $log += "  Wheel-direct for $pkg failed: $($_.Exception.Message)`n"
        }
    }

    # Re-check after wheel-direct
    $finalMissing = @()
    foreach ($m in $required) {
        $r = Test-PyImport -PyExe $PyExe -Module $m
        if (-not $r.ok) {
            $finalMissing += $m
            if ($r.diag) { $log += "Final check: $($r.diag)`n" }
        }
    }
    if ($finalMissing.Count -eq 0) {
        Add-Log "Python deps installed via wheel-direct fallback OK."
        return @{ ok=$true; log=$log; missing=@() }
    }
    return @{ ok=$false; log=$log + "FAILED after pip + wheel-direct fallback. Still missing: $($finalMissing -join ', ')`n"; missing=$finalMissing }
}

function Invoke-VsphereCollect {
    param(
        [string] $PyExe,
        [string] $ScriptPath,
        [string] $VCenterHost,
        [string] $UserRaw,
        [string] $PassRaw,
        [string] $OutputDir
    )

    # Self-heal: portable Python may be missing requests/pyVmomi/urllib3 when
    # the SDT was launched via the zero-install run.ps1 path (which skips
    # install.ps1's pip bootstrap). Ensure deps are present before invoking.
    $sdtDir = Split-Path $ScriptPath -Parent
    $dep = Ensure-PyDeps -PyExe $PyExe -ScriptDir $sdtDir
    if (-not $dep.ok) {
        $depLogFile = ''
        try {
            if ($OutputDir -and (Test-Path $OutputDir)) {
                $stamp = (Get-Date).ToString('yyyy-MM-dd-HHmmss')
                $depLogFile = Join-Path $OutputDir "vsphere-pydeps-$stamp.log"
                [System.IO.File]::WriteAllText($depLogFile, $dep.log, [System.Text.Encoding]::UTF8)
            }
        } catch { }
        return @{
            ok=$false
            log=$dep.log
            error=("Python dependency setup failed. Missing modules: " + ($dep.missing -join ', ') + ". See log for pip/get-pip output.")
            pyError=("ModuleNotFoundError: " + ($dep.missing -join ', '))
            logPath=$depLogFile
            triedUsers=@()
        }
    }

    # Build username variants to try in order
    $variants = [System.Collections.ArrayList]@()
    [void]$variants.Add($UserRaw)

    if ($UserRaw -notmatch '[\\@]') {
        [void]$variants.Add("$UserRaw@vsphere.local")
        [void]$variants.Add("VSPHERE.LOCAL\$UserRaw")
    }
    if ($UserRaw -match '@') {
        $parts = $UserRaw -split '@', 2
        [void]$variants.Add("$($parts[1].ToUpper())\$($parts[0])")
    }
    if ($UserRaw -match '\\') {
        $parts = $UserRaw -split '\\', 2
        [void]$variants.Add("$($parts[1])@$($parts[0].ToLower())")
    }
    $uniqueVariants = @($variants | Select-Object -Unique)

    $env:SDT_HV_PASS = $PassRaw
    $combined = ''
    $succeeded = $null
    $successFile = $null
    $nonAuthFail = $false
    try {
        foreach ($u in $uniqueVariants) {
            # Track files in output dir BEFORE running so we know which one is new
            $beforeSet = @()
            try { $beforeSet = (Get-ChildItem $OutputDir -Filter '*inventory*.json' -EA 0 | ForEach-Object { $_.FullName }) } catch { }

            $pyArgs = @(
                $ScriptPath,
                '--vcenter',  $VCenterHost,
                '--user',     $u,
                '--pass-env', 'SDT_HV_PASS',
                '--output',   $OutputDir
            )
            # Use Start-Process with separate stdout/stderr files so Python's
            # full traceback survives. Direct `2>&1 | Out-String` truncates
            # after the first stderr line when $ErrorActionPreference is Stop
            # somewhere up the call stack.
            $stdOutFile = Join-Path $env:TEMP ("sdt-py-out-" + [guid]::NewGuid().ToString('N').Substring(0,8) + ".txt")
            $stdErrFile = Join-Path $env:TEMP ("sdt-py-err-" + [guid]::NewGuid().ToString('N').Substring(0,8) + ".txt")
            $out = ''
            $exitCode = -1
            try {
                $proc = Start-Process -FilePath $PyExe -ArgumentList $pyArgs `
                    -NoNewWindow -PassThru -Wait `
                    -RedirectStandardOutput $stdOutFile `
                    -RedirectStandardError  $stdErrFile `
                    -ErrorAction Stop
                $exitCode = $proc.ExitCode
                $stdoutTxt = if (Test-Path $stdOutFile) { [System.IO.File]::ReadAllText($stdOutFile) } else { '' }
                $stderrTxt = if (Test-Path $stdErrFile) { [System.IO.File]::ReadAllText($stdErrFile) } else { '' }
                $out = "[exit $exitCode]`n--- STDOUT ---`n$stdoutTxt`n--- STDERR ---`n$stderrTxt"
            } catch {
                $out = "[Start-Process failed before Python ran]`n$($_.Exception.Message)`n$($_.ScriptStackTrace)"
            } finally {
                Remove-Item $stdOutFile -EA SilentlyContinue
                Remove-Item $stdErrFile -EA SilentlyContinue
            }
            $combined += "`n=== attempt with --user '$u' (exit=$exitCode) ===`n$out`n"

            # Find any NEW inventory file written during this attempt
            $afterSet = @()
            try { $afterSet = (Get-ChildItem $OutputDir -Filter '*inventory*.json' -EA 0 | ForEach-Object { $_.FullName }) } catch { }
            $new = $afterSet | Where-Object { $_ -notin $beforeSet } | Select-Object -First 1
            if (-not $new) {
                # Also accept the newest vsphere-perf-*.json if present
                $new = Get-ChildItem $OutputDir -Filter 'vsphere-perf-*.json' -EA 0 | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | ForEach-Object { $_.FullName }
            }
            if ($new -and (Test-Path $new) -and (Get-Item $new).Length -gt 1000) {
                $succeeded = $u
                $successFile = $new
                break
            }

            # If the error isn't auth-related, stop - retrying won't help
            if ($out -notmatch '(?i)(not\s+authenticated|InvalidLogin|Login failed|Incorrect user|password|credentials|SAML|401|403|unauthorized)') {
                $nonAuthFail = $true
                break
            }
        }
    } finally {
        $env:SDT_HV_PASS = $null
    }

    if ($succeeded) {
        return @{ ok=$true; user=$succeeded; log=$combined; file=$successFile }
    }

    # Extract last meaningful Python error line(s) so the UI shows something
    # actionable WITHOUT making the SE expand a collapsed log block.
    $pyError = ''
    $logLines = ($combined -split "`r?`n") | Where-Object { $_ -and $_.Trim() }
    # Look for typical Python exception terminators (line that starts with
    # "<Type>Error:", "<Type>Exception:", or "Errno N", or known vSphere errors)
    $errPatterns = @(
        '^\s*\w+(Error|Exception):\s*.+',
        '^\s*pyVmomi\.vim\.fault\..+',
        '^\s*ssl\.SSL\w+:.+',
        '^\s*socket\.\w+:.+',
        '^\s*urllib\d?\..+:.+',
        '^\s*requests\.exceptions\..+',
        '^\s*ConnectionRefusedError:.+',
        '^\s*\[(SSL|Errno).+'
    )
    foreach ($pat in $errPatterns) {
        $hit = $logLines | Where-Object { $_ -match $pat } | Select-Object -Last 1
        if ($hit) { $pyError = $hit.Trim(); break }
    }
    if (-not $pyError) {
        # Fallback: last non-empty line that isn't a banner
        $pyError = ($logLines | Where-Object { $_ -notmatch '^={3,}|^\s*===\s*attempt' } | Select-Object -Last 1)
        if ($pyError) { $pyError = $pyError.Trim() }
    }

    # Dump the full combined log to disk so it survives the response cycle
    $logFile = ''
    try {
        if ($OutputDir -and (Test-Path $OutputDir)) {
            $stamp = (Get-Date).ToString('yyyy-MM-dd-HHmmss')
            $logFile = Join-Path $OutputDir "vsphere-scan-$stamp.log"
            [System.IO.File]::WriteAllText($logFile, $combined, [System.Text.Encoding]::UTF8)
        }
    } catch { }

    $errBase = if ($nonAuthFail) { 'non-auth failure - check connectivity / SSL / script output' } else { "auth failed for all $($uniqueVariants.Count) username format(s) tried" }
    $errMsg = if ($pyError) { "$errBase. Python said: $pyError" } else { $errBase }
    return @{ ok=$false; log=$combined; error=$errMsg; pyError=$pyError; logPath=$logFile; triedUsers=$uniqueVariants }
}

# -----------------------------------------------------------------------------
# HTML UI - self-contained single page, served as a here-string
# -----------------------------------------------------------------------------
$script:HtmlUI = @'
<!DOCTYPE html><html lang="en"><head>
<meta charset="utf-8"><title>Magna5 SDT - Discovery Session</title>
<style>
:root {
  --bg:#0B1220; --surface:#131A2B; --elevated:#1A2340; --elevated-2:#222C4A;
  --border:rgba(255,255,255,0.07); --border-2:rgba(255,255,255,0.11);
  --text:#E6EAF2; --muted:#8B95A8; --dim:#5A6478;
  --accent:#4F8CFF; --accent-2:#8B5CF6;
  --ok:#22C55E; --warn:#F59E0B; --crit:#EF4444; --info:#38BDF8;
  --mono:"Cascadia Code","Consolas",ui-monospace,monospace;
  --sans:"Segoe UI Variable Display","Segoe UI",system-ui,-apple-system,sans-serif;
}
*{box-sizing:border-box;}html,body{margin:0;padding:0;}
body{background:var(--bg);color:var(--text);font-family:var(--sans);font-size:14px;line-height:1.5;
  background-image:radial-gradient(1200px 600px at 80% -200px,rgba(79,140,255,0.12),transparent 60%),
                   radial-gradient(900px 400px at -200px 200px,rgba(139,92,246,0.08),transparent 50%);
  background-attachment:fixed;min-height:100vh;}
.wrap{max-width:1280px;margin:0 auto;padding:0 28px 60px;}
.hdr{background:rgba(11,18,32,0.75);backdrop-filter:blur(18px);
  border-bottom:1px solid var(--border);padding:14px 28px;display:flex;justify-content:space-between;align-items:center;
  position:sticky;top:0;z-index:100;}
.brand{font-size:16px;font-weight:700;letter-spacing:.5px;
  background:linear-gradient(92deg,#fff 0%,#a7b3ca 100%);
  -webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent;}
.sub{color:var(--muted);font-size:12px;}
.ver-chip{background:var(--elevated);border:1px solid var(--border);color:var(--muted);
  font-size:11px;font-weight:600;padding:4px 10px;border-radius:99px;}
.tab-nav{display:flex;gap:4px;margin:18px 0 22px;background:var(--surface);border:1px solid var(--border);
  border-radius:14px;padding:6px;width:fit-content;}
.tab-btn{background:transparent;border:none;color:var(--muted);padding:10px 22px;font-size:13px;font-weight:600;
  border-radius:10px;cursor:pointer;font-family:var(--sans);}
.tab-btn:hover{color:var(--text);background:var(--elevated);}
.tab-btn.active{background:linear-gradient(135deg,#1c2540 0%,#222c4f 100%);color:var(--text);
  box-shadow:0 0 0 1px var(--border-2),inset 0 1px 0 rgba(255,255,255,0.04);}
.tab-btn:disabled{opacity:.4;cursor:not-allowed;}
.tab-pane{display:none;}
.tab-pane.active{display:block;animation:fadeIn .2s ease;}
@keyframes fadeIn{from{opacity:0;transform:translateY(4px);}to{opacity:1;transform:translateY(0);}}
.card{background:linear-gradient(180deg,var(--surface) 0%,rgba(19,26,43,0.94) 100%);
  border:1px solid var(--border);border-radius:14px;padding:22px 24px;margin-bottom:18px;}
.card-title{font-size:14px;font-weight:700;letter-spacing:-.01em;margin-bottom:2px;}
.card-sub{font-size:12px;color:var(--muted);margin-bottom:14px;}
.section-hdr{font-size:11px;font-weight:700;color:var(--accent);text-transform:uppercase;letter-spacing:.8px;
  margin:18px 0 10px;}
label{display:block;font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;
  letter-spacing:.6px;margin-bottom:6px;}
input,textarea,select{width:100%;background:var(--elevated);border:1px solid var(--border);
  border-radius:8px;padding:10px 14px;color:var(--text);font-family:var(--sans);font-size:13px;
  outline:none;transition:border-color .12s;}
input:focus,textarea:focus,select:focus{border-color:var(--accent);}
textarea{font-family:var(--mono);min-height:120px;resize:vertical;}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px;}
.grid3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px;}
.field{margin-bottom:14px;}
.help{font-size:11px;color:var(--dim);margin-top:4px;}
.btn{background:linear-gradient(135deg,var(--accent) 0%,var(--accent-2) 100%);border:none;color:#fff;
  font-weight:700;font-size:13px;padding:12px 28px;border-radius:10px;cursor:pointer;letter-spacing:.3px;}
.btn:hover{filter:brightness(1.12);}
.btn:disabled{opacity:.4;cursor:not-allowed;filter:none;}
.btn-secondary{background:var(--elevated);color:var(--text);border:1px solid var(--border-2);}
.pill{display:inline-flex;align-items:center;gap:6px;padding:3px 10px;border-radius:99px;
  font-size:11px;font-weight:600;white-space:nowrap;}
.pill .dot{width:6px;height:6px;border-radius:50%;}
.pill.idle{background:rgba(139,149,168,.10);color:var(--muted);border:1px solid var(--border-2);}
.pill.running{background:rgba(79,140,255,.14);color:#93b9ff;border:1px solid rgba(79,140,255,.25);}
.pill.running .dot{background:#4f8cff;box-shadow:0 0 8px rgba(79,140,255,.7);animation:pulse 1s infinite;}
.pill.ok{background:rgba(34,197,94,.12);color:#86efac;border:1px solid rgba(34,197,94,.22);}
.pill.err{background:rgba(239,68,68,.14);color:#fca5a5;border:1px solid rgba(239,68,68,.25);}
@keyframes pulse{0%,100%{opacity:1;}50%{opacity:.4;}}
.pbar{background:rgba(255,255,255,.06);border-radius:4px;height:8px;overflow:hidden;margin:10px 0;}
.pbar > div{height:100%;background:linear-gradient(90deg,var(--accent),var(--accent-2));transition:width .3s;}
.target-row{display:grid;grid-template-columns:28px 1fr auto auto;gap:14px;align-items:center;
  padding:12px 14px;border-bottom:1px solid var(--border);font-size:13px;}
.target-row:last-child{border-bottom:none;}
.buddy{font-family:var(--mono);font-size:12px;color:var(--info);width:48px;}
.target-name{font-family:var(--mono);font-weight:600;color:var(--text);}
.target-phase{color:var(--muted);font-size:12px;}
.logbox{background:#07101f;border:1px solid var(--border);border-radius:10px;padding:14px 16px;
  font-family:var(--mono);font-size:11.5px;color:#c7d1df;height:300px;overflow-y:auto;white-space:pre-wrap;
  line-height:1.55;}
.footer{text-align:center;padding:30px 0 10px;color:var(--dim);font-size:11px;}
.callout{background:rgba(245,158,11,0.08);border-left:3px solid var(--warn);padding:10px 14px;
  border-radius:6px;font-size:12px;color:#fcd34d;margin:12px 0;}
.pw-wrap{position:relative;}
.pw-wrap input{padding-right:40px;}
.pw-toggle{position:absolute;right:8px;top:50%;transform:translateY(-50%);background:transparent;border:none;
  color:var(--muted);cursor:pointer;padding:4px 6px;border-radius:4px;display:flex;align-items:center;justify-content:center;}
.pw-toggle:hover{color:var(--text);background:var(--elevated-2);}
.pw-toggle svg{width:18px;height:18px;}
.hint{display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;border-radius:50%;
  background:var(--elevated-2);color:var(--muted);font-size:10px;font-weight:700;margin-left:8px;cursor:help;
  border:1px solid var(--border-2);vertical-align:middle;user-select:none;position:relative;}
.hint:hover{background:var(--accent);color:#fff;border-color:var(--accent);}
.hint::after{content:attr(data-tip);position:absolute;top:calc(100% + 8px);left:50%;transform:translateX(-50%);
  background:#0b1220;color:var(--text);border:1px solid var(--border-2);border-radius:8px;padding:10px 14px;
  font-size:11.5px;font-weight:500;white-space:normal;width:280px;text-align:left;line-height:1.5;letter-spacing:.1px;
  box-shadow:0 8px 24px rgba(0,0,0,.35);opacity:0;pointer-events:none;transition:opacity .15s;z-index:150;}
.hint:hover::after{opacity:1;}
.hint.right::after{left:auto;right:-8px;transform:none;}
.banner-run{background:linear-gradient(125deg,rgba(79,140,255,0.18) 0%,rgba(139,92,246,0.14) 100%);
  border:1px solid var(--border-2);border-radius:14px;padding:20px 24px;margin-bottom:18px;
  display:flex;justify-content:space-between;align-items:center;}
.banner-run .title{font-size:17px;font-weight:700;}
.banner-run .meta{color:var(--muted);font-size:12px;margin-top:3px;}
</style></head><body>
<div class="hdr">
  <div style="display:flex;align-items:center;gap:14px;">
    <!-- M5 logo mark -->
    <svg width="34" height="34" viewBox="0 0 34 34" xmlns="http://www.w3.org/2000/svg" style="flex-shrink:0;">
      <defs><linearGradient id="lg" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stop-color="#4F8CFF"/>
        <stop offset="100%" stop-color="#8B5CF6"/>
      </linearGradient></defs>
      <rect x="1" y="1" width="32" height="32" rx="8" fill="url(#lg)"/>
      <text x="17" y="22" text-anchor="middle" font-family="Segoe UI Variable Display, Segoe UI, system-ui" font-size="13" font-weight="700" fill="#fff" letter-spacing="-0.5">M5</text>
    </svg>
    <div class="brand">MAGNA5</div>
    <div style="color:var(--dim);">/</div>
    <div class="sub">SDT - Discovery Session</div>
  </div>
  <span class="ver-chip">SDT GUI v__VERSION__</span>
</div>
<div class="wrap">
<div class="tab-nav">
  <button class="tab-btn active" id="tb-setup" onclick="setTab('setup')">Setup</button>
  <button class="tab-btn" id="tb-run" onclick="setTab('run')">Run</button>
  <button class="tab-btn" id="tb-report" onclick="setTab('report')" disabled>Report</button>
</div>

<!-- SETUP TAB -->
<div id="tab-setup" class="tab-pane active">

<!-- Quick action: scan THIS machine only (no HV, no creds, no remote) -->
<div id="localOnlyBar" style="display:flex;align-items:center;gap:10px;padding:8px 12px;margin-bottom:10px;font-size:12px;color:var(--muted);background:rgba(80,160,255,0.05);border:1px solid var(--border);border-radius:8px;">
<span style="flex:1;">No HV / remote access? Quick-scan this machine only:</span>
<button type="button" class="btn btn-secondary" id="localOnlyBtn" onclick="runLocalOnly()" style="padding:4px 10px;font-size:12px;white-space:nowrap;">Scan this box</button>
<span id="localOnlyStatus" style="font-size:11px;"></span>
</div>

<form id="setupForm" novalidate onsubmit="submitSetup(event)">

<div class="card">
<div class="card-title">Session <span class="hint" data-tip="Client name appears in the final HTML report header. Output folder is where every JSON, log, and final report for this discovery run gets written.">i</span></div>
<div class="card-sub">Client name and output folder. Credentials live in memory only - never saved to disk.</div>
<div class="grid2">
<div class="field"><label>Client name <span class="hint" data-tip="Appears in the report header. Use the sales-facing name (e.g. Acme Corporation).">i</span></label>
<input name="client" required placeholder="Acme Corporation"></div>
<div class="field"><label>Output folder <span class="hint" data-tip="Discovery JSONs + HTML report land here. Default is fine; a per-client subfolder is created automatically.">i</span></label>
<input name="outputDir" value="C:\Temp\sdt\sessions" required></div>
</div>
<div class="grid2" style="margin-top:10px;">
<div class="field"><label>Parallel scans <span class="hint" data-tip="How many servers to scan at once. 4 is a good default. Go higher (6-8) on powerful boxes, lower if you see WinRM errors from load. Each scan takes ~3-5 minutes on busy servers.">i</span></label>
<select name="parallel"><option value="1" selected>1 (sequential - default)</option><option value="2">2</option><option value="4">4</option><option value="6">6</option><option value="8">8</option></select></div>
<div class="field"></div>
</div>
</div>

<div class="card">
<div class="card-title">Admin credentials <span class="hint" data-tip="Used to remotely sign in to every ticked VM + every manual target. Needs WinRM remoting permissions (domain admin works; local admin works for workgroup hosts).">i</span></div>
<div class="card-sub">Domain admin or local admin that can log into the Windows targets. Held in memory only - never written to disk.</div>
<div class="grid2">
<div class="field"><label>Domain / Local admin <span class="hint" data-tip="Format: DOMAIN\\username for a domain admin, user@domain.local for UPN, or .\\username for a local admin on workgroup hosts.">i</span></label>
<input name="winrmUser" placeholder="DOMAIN\administrator or admin@contoso.local"></div>
<div class="field"><label>Password <span class="hint" data-tip="Password for the admin account. Held in memory only for this session - no file, no registry, no persistence.">i</span></label>
<div class="pw-wrap">
<input name="winrmPass" type="password" autocomplete="off">
<button type="button" class="pw-toggle" onclick="togglePw(this)" title="Show/hide password" aria-label="Toggle password visibility">
<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>
</button>
</div>
</div>
</div>
<div style="display:flex;align-items:center;gap:12px;margin-top:10px;">
<button type="button" class="btn btn-secondary" id="testCredsBtn" onclick="testCreds()">Test creds</button>
<span id="credStatus" style="font-size:12px;color:var(--muted);">Validates against the domain (DOMAIN\user or UPN). Local accounts (.\user) require a manual target.</span>
</div>
</div>

<div class="card">
<div class="card-title">Hypervisor (optional) <span class="hint" data-tip="Hit 'Scan Hypervisor' to connect, list every VM, and tick which ones to collect Windows data from. Skip this whole section if you're running against bare-metal servers only.">i</span></div>
<div class="card-sub">Connect to vCenter or an ESXi host. Hit <strong>Scan Hypervisor</strong> to discover VMs, then tick which ones to collect Windows data from. Or leave "None" for bare-metal-only runs.</div>
<div class="grid3">
<div class="field"><label>Type <span class="hint" data-tip="vCenter/ESXi uses the vSphere SOAP API (remote). Hyper-V is LOCAL-ONLY: must be run on the Hyper-V host itself (host/user/pass fields are ignored). 'None' skips hypervisor inventory entirely.">i</span></label>
<select name="hvType" onchange="onHvTypeChange(this.value)"><option value="none">None</option><option value="vsphere">vCenter / ESXi (remote)</option><option value="hyperv">Hyper-V Host (local-only)</option></select></div>
<div class="field"><label>IP / FQDN <span class="hint" data-tip="Hostname or IP of your vCenter server (or single ESXi host). Port 443 must be reachable.">i</span></label>
<input name="hvHost" placeholder="192.168.10.75"></div>
<div class="field"><label>User <span class="hint" data-tip="vCenter SSO account (e.g. administrator@vsphere.local) or local ESXi root. Needs read access to inventory + perf counters.">i</span></label>
<input name="hvUser" placeholder="administrator@vsphere.local"></div>
</div>
<div class="grid2">
<div class="field"><label>Password <span class="hint" data-tip="Password for the hypervisor account. Held in memory only for this session.">i</span></label>
<div class="pw-wrap">
<input name="hvPass" type="password" autocomplete="off">
<button type="button" class="pw-toggle" onclick="togglePw(this)" title="Show/hide password" aria-label="Toggle password visibility">
<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>
</button>
</div>
</div>
<div class="field" style="display:flex;align-items:flex-end;">
<button type="button" class="btn btn-secondary" id="scanHvBtn" onclick="scanHv()">Scan Hypervisor</button>
</div>
</div>
<div id="scanStatus" style="margin-top:10px;font-size:12px;color:var(--muted);"></div>
</div>

<!-- Discovered VMs (appears after hypervisor scan) -->
<div class="card" id="discoveredCard" style="display:none;">
<div class="card-title">Discovered VMs <span class="hint" data-tip="Every VM found on the hypervisor. Ticked rows get per-server WinRM/Windows discovery in addition to the HV inventory. Use the Filter input to narrow by name, IP, or OS.">i</span></div>
<div class="card-sub">Tick rows to include in per-server Windows discovery. Linux / appliance / vCenter boxes are auto-unchecked.</div>
<div style="display:flex;gap:10px;align-items:center;margin:10px 0;">
<input id="vmFilter" placeholder="Filter by name, IP, OS..." oninput="renderVmTable()" style="flex:1;">
<button type="button" class="btn btn-secondary" onclick="toggleAllVms(true)">Select all</button>
<button type="button" class="btn btn-secondary" onclick="toggleAllVms(false)">Clear</button>
</div>
<div id="vmTableWrap" style="max-height:360px;overflow-y:auto;border:1px solid var(--border);border-radius:10px;">
<table class="dt" id="vmTable" style="width:100%;font-size:12.5px;"><thead><tr>
<th style="width:40px;text-align:center;padding:8px 10px;"><input type="checkbox" id="vmAllCbx" onclick="toggleAllVms(this.checked)"></th>
<th style="padding:8px 10px;">Name</th>
<th style="padding:8px 10px;">IP</th>
<th style="padding:8px 10px;">Guest OS</th>
<th style="padding:8px 10px;">Power</th>
</tr></thead><tbody id="vmTableBody"></tbody></table>
</div>
<div id="vmCount" style="font-size:11px;color:var(--muted);margin-top:8px;">0 VMs discovered</div>
</div>

<div class="card">
<div class="card-title">Windows Targets (manual additions) <span class="hint" data-tip="Only needed for hosts that aren't in the hypervisor scan above - e.g. physical boxes, DMZ VMs, or anything the hypervisor can't see. Leave blank if everything's in the HV.">i</span></div>
<div class="card-sub">Hosts not in the hypervisor above. One per line. IP or hostname. Leave blank if HV scan covers everything.
<br>Example: <code style="font-family:var(--mono);color:var(--info);">192.168.10.4</code> or <code style="font-family:var(--mono);color:var(--info);">QES-OFFICE-DC</code></div>
<div class="field"><label>Manual targets <span class="hint" data-tip="One host per line. IP or hostname. The script will try remote WinRM first; if that fails, you'll see an error row in the Run tab.">i</span></label>
<textarea name="targets" placeholder="Optional - only if a host isn't in the HV"></textarea></div>
<div class="callout">
<strong>Coming soon:</strong> subnet auto-discovery (scan button). For now, list targets manually or use AD/DNS export.
</div>
</div>

<div style="display:flex;justify-content:flex-end;gap:10px;">
<button type="submit" id="runBtn" class="btn" onclick="submitSetup(event)">Run Discovery</button>
</div>
</form>
</div>

<!-- RUN TAB -->
<div id="tab-run" class="tab-pane">
<div class="banner-run">
  <div>
    <div class="title" id="runTitle">Waiting to start...</div>
    <div class="meta" id="runMeta">Submit the setup form to begin.</div>
  </div>
  <span id="runPill" class="pill idle"><span class="dot"></span><span id="runPillText">Idle</span></span>
</div>

<div class="card">
<div class="card-title">Progress</div>
<div class="pbar"><div id="overallBar" style="width:0%"></div></div>
<div id="progressText" style="font-size:12px;color:var(--muted);">0 / 0 targets complete</div>
<div class="section-hdr">Targets</div>
<div id="targetList">
<div style="color:var(--muted);font-style:italic;padding:20px;text-align:center;">No run in progress.</div>
</div>
</div>

<div class="card">
<div class="card-title">Log</div>
<div class="logbox" id="logbox">Waiting for discovery to start...</div>
</div>

<div style="margin-top:18px;text-align:center;">
  <button type="button" class="btn btn-secondary" onclick="setTab('report')" style="padding:10px 24px;">
    Jump to Report Tab &rarr;
  </button>
</div>
</div>

<!-- RESULTS TAB - minimal: 'Scan complete' + 2 buttons. All diagnostics
     collapsed behind a small Details disclosure. -->
<div id="tab-report" class="tab-pane">
<div id="reportContent">
<div style="color:var(--muted);font-style:italic;padding:40px 20px;text-align:center;">Run a discovery session first.</div>
</div>
</div>

<div class="footer" style="line-height:1.8;">
  <div style="font-weight:600;color:var(--muted);">Intellectual property of Magna5, Inc. &middot; All rights reserved.</div>
  <div>Written by <strong style="color:var(--text);">Matthew Kelly</strong> &middot; Magna5 Solutions Engineering</div>
  <div>Questions: <a href="mailto:matthew.kelly@magna5.com" style="color:var(--accent);">matthew.kelly@magna5.com</a></div>
  <div style="margin-top:6px;color:var(--dim);">SDT GUI v__VERSION__</div>
</div>
</div>

<script>
function setTab(t){
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.toggle('active', b.id==='tb-'+t));
  document.querySelectorAll('.tab-pane').forEach(p=>p.classList.toggle('active', p.id==='tab-'+t));
  // Track current pane on body so CSS can hide chrome + show back button.
  document.body.classList.remove('tab-setup','tab-run','tab-report');
  document.body.classList.add('tab-' + t);
  window.scrollTo(0,0);
}

// Holds the last hypervisor scan result so submit can pass checked VMs.
window._discoveredVMs = [];

function onHvTypeChange(v){
  // Hyper-V is local-only: dim host/user/pass since they're ignored.
  const dim = (v === 'hyperv');
  ['hvHost','hvUser','hvPass'].forEach(n => {
    const el = document.querySelector('[name="'+n+'"]');
    if (!el) return;
    el.disabled = dim;
    el.style.opacity = dim ? '0.45' : '1';
    if (dim) el.placeholder = '(not used in Hyper-V local mode)';
  });
  const ss = document.getElementById('scanStatus');
  if (ss) {
    if (v === 'hyperv') {
      ss.style.color = 'var(--muted)';
      ss.innerHTML = 'Hyper-V mode is LOCAL-ONLY -- the tool must run on the Hyper-V host. Click <strong>Scan Hypervisor</strong> to run preflight + Get-VM on this machine.';
    } else if (v === 'none') {
      ss.textContent = '';
    } else {
      ss.style.color = 'var(--muted)';
      ss.textContent = 'vCenter / ESXi -- enter host, user, and password, then Scan Hypervisor.';
    }
  }
}

function getHvFields(){
  const form = document.getElementById('setupForm');
  const fd = new FormData(form);
  return {
    hvType: fd.get('hvType') || 'none',
    hvHost: (fd.get('hvHost')||'').trim(),
    hvUser: (fd.get('hvUser')||'').trim(),
    hvPass: fd.get('hvPass') || ''
  };
}

async function scanHv(){
  const btn = document.getElementById('scanHvBtn');
  const status = document.getElementById('scanStatus');
  const hv = getHvFields();
  if (hv.hvType === 'none') { alert('Pick a hypervisor type first.'); return; }

  // Hyper-V is LOCAL-ONLY: no host/creds prompt; runs preflight + Get-VM on this machine.
  if (hv.hvType === 'hyperv') { return scanHvLocal(false); }

  if (!hv.hvHost || !hv.hvUser || !hv.hvPass) { alert('Hypervisor host, user, and password required.'); return; }

  btn.disabled = true; btn.textContent = 'Scanning...';
  status.style.color = 'var(--muted)';
  status.innerHTML = 'Connecting to ' + escapeHtml(hv.hvHost) + '. This can take 30-60 seconds for a mid-size vCenter...';

  try {
    const resp = await fetch('/api/hv-scan', {
      method: 'POST', headers: {'Content-Type':'application/json'},
      body: JSON.stringify(hv)
    });
    const data = await resp.json();
    if (!resp.ok || !data.ok) {
      const err = data.error || 'scan failed';
      // Auto-expand the log block. Saved-log path shown so SE can grab it later.
      if (data.log) {
        status.style.color = 'var(--crit)';
        const pathLine = data.logPath
          ? '<div style="margin-top:8px;font-size:11px;color:var(--muted);">Full log saved to: <code style="color:#a7b3ca;">' + escapeHtml(data.logPath) + '</code></div>'
          : '';
        status.innerHTML = '<strong>Scan failed:</strong> ' + escapeHtml(err) + pathLine +
          '<details open style="margin-top:10px;"><summary style="cursor:pointer;color:var(--muted);">Python/collector output (auto-expanded)</summary>' +
          '<pre style="background:#07101f;color:#c7d1df;padding:10px;border-radius:6px;font-family:var(--mono);font-size:11px;overflow:auto;white-space:pre-wrap;word-break:break-word;">' +
          escapeHtml(data.log) + '</pre></details>';
        return;
      }
      throw new Error(err);
    }

    window._discoveredVMs = (data.vms || []).map(v => ({...v, _checked: isWinVM(v)}));
    document.getElementById('discoveredCard').style.display = '';
    renderVmTable();
    status.style.color = 'var(--ok)';
    status.textContent = 'Scanned ' + window._discoveredVMs.length + ' VMs. Tick the ones to include in Windows discovery below.';
  } catch(err) {
    status.style.color = 'var(--crit)';
    status.textContent = 'Scan failed: ' + err.message;
  } finally {
    btn.disabled = false; btn.textContent = 'Scan Hypervisor';
  }
}

// Hyper-V LOCAL mode: no remote host, no creds. Calls /api/hv-local-scan which runs
// the preflight + Get-VM on the box this GUI is running on. If preflight fails, the
// user sees the specific signal failures and can: (R)etry, (F)orce, or skip.
async function scanHvLocal(force){
  const btn = document.getElementById('scanHvBtn');
  const status = document.getElementById('scanStatus');
  btn.disabled = true; btn.textContent = force ? 'Forcing local scan...' : 'Running local Hyper-V scan...';
  status.style.color = 'var(--muted)';
  status.innerHTML = 'Hyper-V local mode -- preflight + Get-VM on this machine. No remote connection, no creds.';
  try {
    const resp = await fetch('/api/hv-local-scan', {
      method: 'POST', headers: {'Content-Type':'application/json'},
      body: JSON.stringify({ force: !!force })
    });
    const data = await resp.json();
    if (!data.ok && data.preflight) {
      // Preflight failed -- show signal table + retry/force/skip controls
      status.style.color = 'var(--warn)';
      const sig = data.detail || {};
      const sigRow = (k,v) => '<tr><td style="padding:4px 10px;font-family:var(--mono);">' + k + '</td><td style="padding:4px 10px;">' + (v ? '<span style="color:var(--ok)">PASS</span>' : '<span style="color:var(--crit)">FAIL</span>') + '</td></tr>';
      const reasons = (data.reasons || []).map(r => '<li>' + escapeHtml(r) + '</li>').join('');
      status.innerHTML =
        '<strong>This machine is not a Hyper-V host (' + (data.passCount || 0) + '/4 signals).</strong>' +
        '<table style="margin-top:8px;font-size:11.5px;">' +
          sigRow('vmms service running',        !!sig.vmms_service) +
          sigRow('Hyper-V feature enabled',     !!sig.hyperv_feature) +
          sigRow('HypervisorPresent flag',      !!sig.hypervisor_present) +
          sigRow('Get-VM cmdlet available',     !!sig.get_vm_cmdlet) +
        '</table>' +
        '<ul style="margin-top:8px;font-size:12px;">' + reasons + '</ul>' +
        '<div style="margin-top:10px;display:flex;gap:8px;">' +
          '<button type="button" class="btn btn-secondary" onclick="scanHvLocal(false)">Retry preflight</button>' +
          '<button type="button" class="btn btn-secondary" onclick="scanHvLocal(true)">Force scan anyway</button>' +
        '</div>';
      return;
    }
    if (!data.ok) { throw new Error(data.error || 'local Hyper-V scan failed'); }

    window._discoveredVMs = (data.vms || []).map(v => ({...v, _checked: isWinVM(v)}));
    document.getElementById('discoveredCard').style.display = '';
    renderVmTable();
    status.style.color = 'var(--ok)';
    const forcedNote = data.forced ? ' (forced past preflight)' : '';
    status.textContent = 'Hyper-V local scan: ' + window._discoveredVMs.length + ' VMs on ' + (data.host || 'this host') + forcedNote + '. Tick which to include in Windows discovery below.';
  } catch(err) {
    status.style.color = 'var(--crit)';
    status.textContent = 'Scan failed: ' + err.message;
  } finally {
    btn.disabled = false; btn.textContent = 'Scan Hypervisor';
  }
}

function isWinVM(v){
  const os = (v.GuestOS || v.guestOs || '').toLowerCase();
  const nm = (v.Name || v.name || '').toLowerCase();
  // Uncheck Linux, Photon, vCenter appliances, and anything with 'vcsa'/'vcenter' in the name
  if (/linux|photon|ubuntu|debian|centos|redhat|bsd|coreos/.test(os)) return false;
  if (/vcsa|vcenter|esxi\b/.test(nm)) return false;
  return true;
}

function renderVmTable(){
  const tbody = document.getElementById('vmTableBody');
  const count = document.getElementById('vmCount');
  const filter = (document.getElementById('vmFilter').value || '').toLowerCase();
  const vms = window._discoveredVMs || [];
  const filtered = vms.filter(v => {
    if (!filter) return true;
    const hay = [v.Name, v.name, v.IPs, v.ips, v.GuestOS, v.guestOs, v.PowerState, v.powerState].filter(Boolean).join(' ').toLowerCase();
    return hay.includes(filter);
  });
  tbody.innerHTML = filtered.map((v, i) => {
    const nm  = v.Name || v.name || '';
    const ip  = v.IPs || v.ips || '';
    const os  = v.GuestOS || v.guestOs || '';
    const ps  = v.PowerState || v.powerState || '';
    const idx = vms.indexOf(v);
    const powCls = /on|POWERED_ON/i.test(ps) ? 'ok' : 'neutral';
    return `<tr>
      <td style="text-align:center;padding:6px 10px;"><input type="checkbox" data-vm-idx="${idx}" ${v._checked ? 'checked':''} onchange="window._discoveredVMs[${idx}]._checked=this.checked;updateVmCount();"></td>
      <td style="padding:6px 10px;font-family:var(--mono);font-weight:600;">${escapeHtml(nm)}</td>
      <td style="padding:6px 10px;font-family:var(--mono);font-size:11px;">${escapeHtml(ip)}</td>
      <td style="padding:6px 10px;font-size:11.5px;color:var(--muted);">${escapeHtml(os)}</td>
      <td style="padding:6px 10px;"><span class="pill ${powCls}"><span class="dot"></span>${escapeHtml(ps)}</span></td>
    </tr>`;
  }).join('') || '<tr><td colspan="5" style="color:var(--muted);padding:14px;text-align:center;">No matches.</td></tr>';
  updateVmCount();
}

function updateVmCount(){
  const vms = window._discoveredVMs || [];
  const checked = vms.filter(v => v._checked).length;
  document.getElementById('vmCount').textContent = checked + ' of ' + vms.length + ' selected';
}

function toggleAllVms(on){
  (window._discoveredVMs || []).forEach(v => v._checked = !!on);
  renderVmTable();
}

async function runLocalOnly(){
  const btn = document.getElementById('localOnlyBtn');
  const status = document.getElementById('localOnlyStatus');
  if (!confirm('Run discovery against THIS machine only? No HV, no remote targets.')) return;
  btn.disabled = true; btn.textContent = 'Starting...';
  status.textContent = '';
  status.style.color = 'var(--muted)';
  try {
    const r = await fetch('/api/local-only', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' });
    const d = await r.json();
    if (!d.ok) {
      status.textContent = 'Failed: ' + (d.error || 'unknown error');
      status.style.color = 'var(--danger, #f55)';
      btn.disabled = false; btn.textContent = 'Scan this box';
      return;
    }
    status.textContent = '';
    setTab('run');
    // CRITICAL: kick off the status poller so the Run tab updates as the
    // discovery progresses. Without this the UI stalls on "Waiting...".
    startPolling();
  } catch (e) {
    status.textContent = 'Request failed: ' + e.message;
    status.style.color = 'var(--danger, #f55)';
    btn.disabled = false; btn.textContent = 'Scan this box';
  }
}

async function submitSetup(e){
  e.preventDefault();
  const form = document.getElementById('setupForm');
  const data = Object.fromEntries(new FormData(form).entries());
  // Start with manual targets
  let targets = (data.targets||'').split(/\r?\n/).map(s=>s.trim()).filter(Boolean);
  // Add checked VMs (prefer IP, fall back to hostname)
  const checked = (window._discoveredVMs || []).filter(v => v._checked);
  const vmTargets = checked.map(v => (v.IPs || v.ips || v.Name || v.name || '').toString().split(',')[0].trim()).filter(Boolean);
  targets = Array.from(new Set([...targets, ...vmTargets]));
  data.targets = targets;

  // Up-front sanity checks so Run Discovery never appears to silently do nothing
  if (!data.client || !data.client.trim()) { alert('Client name is required.'); return; }
  if (!data.outputDir || !data.outputDir.trim()) { alert('Output folder is required.'); return; }
  const hvConfigured = data.hvType && data.hvType !== 'none' && data.hvHost;
  if (targets.length === 0 && !hvConfigured) {
    alert('No targets. Enter manual targets, tick VMs from a hypervisor scan, or configure a hypervisor.'); return;
  }

  const btn = document.getElementById('runBtn') || form.querySelector('button[type="submit"]');
  if (btn) { btn.disabled = true; btn.textContent = 'Starting...'; }

  try {
    const resp = await fetch('/api/start', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify(data)
    });
    const respText = await resp.text();
    if (!resp.ok) { throw new Error(respText || ('HTTP ' + resp.status)); }
    setTab('run');
    const tbs = document.getElementById('tb-setup'); if (tbs) tbs.disabled = true;
    startPolling();
  } catch(err) {
    alert('Failed to start: ' + err.message);
    if (btn) { btn.disabled = false; btn.textContent = 'Run Discovery'; }
  }
}

let pollTimer = null;
function startPolling(){
  if (pollTimer) return;
  pollTimer = setInterval(async () => {
    try {
      const resp = await fetch('/api/status');
      const s = await resp.json();
      renderStatus(s);
      if (s.Status === 'complete' || s.Status === 'error') {
        clearInterval(pollTimer); pollTimer = null;
        if (s.Status === 'complete') {
          document.getElementById('tb-report').disabled = false;
          renderReport(s);
          // Auto-jump to Results tab once the run finishes so the user
          // lands directly on the report. Setup/Run tabs remain navigable
          // with their last state preserved.
          setTab('report');
        }
      }
    } catch(e) { /* transient network error; keep polling */ }
  }, 1000);
}

function renderStatus(s){
  const pill = document.getElementById('runPill');
  const pillText = document.getElementById('runPillText');
  pill.className = 'pill ' + (s.Status==='running'?'running':s.Status==='complete'?'ok':s.Status==='error'?'err':'idle');
  pillText.textContent = s.Status.charAt(0).toUpperCase()+s.Status.slice(1);
  document.getElementById('runTitle').textContent = s.Client ? 'Discovering ' + s.Client : 'Discovery session';
  document.getElementById('runMeta').textContent = (s.SessionDir || '-');
  const tot = s.Targets.length;
  const done = s.Targets.filter(t=>t.State==='done'||t.State==='error').length;
  const pct = tot ? Math.round(done/tot*100) : 0;
  document.getElementById('overallBar').style.width = pct + '%';
  document.getElementById('progressText').textContent = done + ' / ' + tot + ' targets complete (' + pct + '%)';
  const list = document.getElementById('targetList');
  if (tot === 0) { list.innerHTML = '<div style="color:var(--muted);padding:14px;">No targets.</div>'; }
  else {
    list.innerHTML = s.Targets.map(t => {
      const stateCls = t.State==='done'?'ok':t.State==='error'?'err':t.State==='running'?'running':'idle';
      const stateLabel = t.State.charAt(0).toUpperCase()+t.State.slice(1);
      return `<div class="target-row">
        <span class="buddy">${escapeHtml(t.Buddy || '')}</span>
        <div><div class="target-name">${escapeHtml(t.Name)}</div>
        <div class="target-phase">${escapeHtml(t.Phase || '')}</div></div>
        <span class="pill ${stateCls}"><span class="dot"></span>${stateLabel}</span>
      </div>`;
    }).join('');
  }
  const lb = document.getElementById('logbox');
  const atBottom = lb.scrollTop + lb.clientHeight >= lb.scrollHeight - 20;
  lb.textContent = (s.LogTail || []).join('\n');
  if (atBottom) lb.scrollTop = lb.scrollHeight;
}

function renderReport(s){
  const el = document.getElementById('reportContent');
  const sessionDir = s.SessionDir || '';
  const openFolderBtn = sessionDir
    ? `<button class="btn btn-secondary" onclick="openSessionFolder()" style="margin-top:10px;margin-left:8px;">Open session folder</button>`
    : '';
  const viewLogBtn = sessionDir
    ? `<button class="btn btn-secondary" onclick="viewGenLog()" style="margin-top:10px;margin-left:8px;">View gen_report log</button>`
    : '';
  const zipBtn = s.ReportZipPath
    ? `<a href="/api/download-zip" class="btn" style="margin-top:10px;margin-left:8px;display:inline-block;text-decoration:none;">Download zip</a>`
    : '';
  const copyLogsBtn = `<button class="btn btn-secondary" onclick="copyLogs(this)" style="margin-top:10px;margin-left:8px;">Copy logs</button>`;

  // Missing-JSON audit block (only if anything is missing)
  let missingHtml = '';
  if (Array.isArray(s.MissingTargets) && s.MissingTargets.length) {
    const rows = s.MissingTargets.map(m =>
      `<li><code>${escapeHtml(m.Address)}</code> <span style="color:var(--muted);">(${escapeHtml(m.Kind||'')})</span> - <span style="color:var(--warn);">${escapeHtml(m.Reason||'')}</span></li>`
    ).join('');
    missingHtml = `<div style="margin-top:14px;padding:10px 14px;background:#2a1b0f;border-left:3px solid var(--warn);border-radius:6px;">
      <div style="font-size:11px;color:var(--warn);text-transform:uppercase;letter-spacing:.6px;margin-bottom:6px;">Missing JSON (${s.MissingTargets.length})</div>
      <ul style="margin:0;padding-left:18px;font-size:12.5px;line-height:1.55;">${rows}</ul>
    </div>`;
  }

  if (s.ReportPath) {
    // SUCCESS: blank page, "Scan complete" headline, two big buttons. All
    // diagnostics collapsed behind small Details disclosure.
    el.innerHTML = `<div style="text-align:center;padding:80px 20px 20px;">
        <h1 style="font-size:28px;font-weight:700;margin-bottom:32px;">Scan complete</h1>
        <div style="display:flex;gap:16px;justify-content:center;flex-wrap:wrap;">
          <a href="/api/report-html" target="_blank" class="btn"
             style="text-decoration:none;font-size:15px;padding:18px 36px;">
            Open HTML Report
          </a>
          <button type="button" class="btn btn-secondary" onclick="openSessionFolder()"
             style="font-size:15px;padding:18px 36px;">
            Open Session Folder
          </button>
        </div>
      </div>
      <details style="margin-top:48px;max-width:700px;margin-left:auto;margin-right:auto;">
        <summary style="cursor:pointer;color:var(--muted);font-size:11px;text-align:center;">Details &amp; logs</summary>
        <div style="margin-top:12px;font-size:12px;">
          <p>Report path: <code style="font-family:var(--mono);color:var(--info);word-break:break-all;">${escapeHtml(s.ReportPath)}</code></p>
          <p>${zipBtn}${viewLogBtn}${copyLogsBtn}</p>
          ${missingHtml}
          <div id="genLogBox"></div>
        </div>
      </details>`;
  } else {
    // FAILURE: blank page, warning + retry button. Diagnostics collapsed.
    el.innerHTML = `<div style="text-align:center;padding:80px 20px 20px;">
        <h1 style="font-size:28px;font-weight:700;margin-bottom:12px;color:var(--warn);">Scan complete (no report)</h1>
        <p style="color:var(--muted);font-size:13px;margin-bottom:32px;">The discovery scans finished but the HTML report failed to generate.</p>
        <div style="display:flex;gap:16px;justify-content:center;flex-wrap:wrap;">
          <button type="button" class="btn" onclick="retryReportGen(this)"
             style="font-size:15px;padding:18px 36px;">
            Retry Report Generation
          </button>
          <button type="button" class="btn btn-secondary" onclick="openSessionFolder()"
             style="font-size:15px;padding:18px 36px;">
            Open Session Folder
          </button>
        </div>
      </div>
      <details style="margin-top:48px;max-width:700px;margin-left:auto;margin-right:auto;">
        <summary style="cursor:pointer;color:var(--muted);font-size:11px;text-align:center;">Diagnostics</summary>
        <div style="margin-top:12px;font-size:12px;">
          <p style="color:var(--muted);">Session dir: <code style="font-family:var(--mono);">${escapeHtml(sessionDir)}</code></p>
          <p>${viewLogBtn}${copyLogsBtn}</p>
          ${missingHtml}
          <div id="genLogBox"></div>
        </div>
      </details>`;
  }
}

async function retryReportGen(btn){
  const orig = btn.textContent;
  btn.disabled = true;
  btn.textContent = 'Regenerating...';
  try {
    const resp = await fetch('/api/regenerate-report', { method: 'POST' });
    const data = await resp.json();
    if (data.ok && data.reportPath) {
      // Refresh the Report tab content via /api/status
      const s = await (await fetch('/api/status')).json();
      window._lastStatus = s;
      renderReportPane(s);
      return;
    }
    btn.textContent = data.error ? ('Failed: ' + data.error.substring(0, 60)) : 'Failed';
    setTimeout(()=>{ btn.textContent = orig; btn.disabled = false; }, 4000);
  } catch(e) {
    btn.textContent = 'Failed: ' + e.message.substring(0, 60);
    setTimeout(()=>{ btn.textContent = orig; btn.disabled = false; }, 4000);
  }
}

async function copyLogs(btn){
  const orig = btn.textContent;
  btn.textContent = 'Copying...';
  btn.disabled = true;
  try {
    const resp = await fetch('/api/combined-logs');
    const data = await resp.json();
    const text = data.content || '';
    try {
      await navigator.clipboard.writeText(text);
      btn.textContent = 'Copied!';
    } catch(e) {
      // Fallback for non-HTTPS/localhost contexts
      const ta = document.createElement('textarea');
      ta.value = text; document.body.appendChild(ta); ta.select();
      document.execCommand('copy'); document.body.removeChild(ta);
      btn.textContent = 'Copied!';
    }
    setTimeout(()=>{ btn.textContent = orig; btn.disabled = false; }, 1500);
  } catch(e) {
    btn.textContent = 'Failed';
    setTimeout(()=>{ btn.textContent = orig; btn.disabled = false; }, 1500);
  }
}

async function openSessionFolder(){
  try {
    const resp = await fetch('/api/open-folder', { method: 'POST' });
    if (!resp.ok) { alert('Failed to open folder: ' + await resp.text()); }
  } catch(e) { alert('Error: ' + e.message); }
}

async function viewGenLog(){
  const box = document.getElementById('genLogBox');
  box.innerHTML = '<p style="color:var(--muted);font-size:12px;margin-top:14px;">Loading log...</p>';
  try {
    const resp = await fetch('/api/gen-log');
    const data = await resp.json();
    if (data.content) {
      box.innerHTML = `<div style="margin-top:14px;"><div style="font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.6px;margin-bottom:6px;">gen_report.log</div>
        <pre style="background:#07101f;color:#c7d1df;padding:14px;border-radius:8px;font-family:var(--mono);font-size:11.5px;max-height:420px;overflow:auto;white-space:pre-wrap;">${escapeHtml(data.content)}</pre></div>`;
    } else {
      box.innerHTML = `<p style="color:var(--muted);font-size:12px;margin-top:10px;">No log file: ${escapeHtml(data.error || 'unknown')}</p>`;
    }
  } catch(e) {
    box.innerHTML = `<p style="color:var(--crit);font-size:12px;margin-top:10px;">Error loading log: ${escapeHtml(e.message)}</p>`;
  }
}

function escapeHtml(s){
  return String(s||'').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function togglePw(btn){
  const input = btn.parentElement.querySelector('input');
  if (!input) return;
  input.type = (input.type === 'password') ? 'text' : 'password';
}

async function testCreds(){
  const form = document.getElementById('setupForm');
  const user = (form.winrmUser.value || '').trim();
  const pass = form.winrmPass.value || '';
  const btn  = document.getElementById('testCredsBtn');
  const st   = document.getElementById('credStatus');
  if (!user || !pass) {
    st.style.color = 'var(--warn)';
    st.textContent = 'Enter a username and password first.';
    return;
  }
  btn.disabled = true;
  const orig = btn.textContent;
  btn.textContent = 'Checking...';
  st.style.color = 'var(--muted)';
  st.textContent = 'Contacting domain controller...';
  try {
    const resp = await fetch('/api/test-creds', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ winrmUser: user, winrmPass: pass })
    });
    const data = await resp.json();
    if (data.ok) {
      st.style.color = 'var(--ok)';
      st.textContent = 'OK - ' + (data.message || 'credentials valid');
    } else if (data.soft) {
      // Soft warning: we couldnt reach a DC from this collector. Not an error.
      st.style.color = 'var(--warn)';
      st.textContent = 'Warning: ' + (data.error || 'could not validate from this collector - will test during scan');
    } else {
      st.style.color = 'var(--crit)';
      st.textContent = data.error || 'validation failed';
    }
  } catch(e) {
    st.style.color = 'var(--crit)';
    st.textContent = 'Error: ' + e.message;
  } finally {
    btn.textContent = orig;
    btn.disabled = false;
  }
}
</script>
</body></html>
'@

# Inject version into template
$script:HtmlUI = $script:HtmlUI -replace '__VERSION__', $script:Version

# -----------------------------------------------------------------------------
# HTTP LISTENER
# -----------------------------------------------------------------------------
function Start-HttpListener {
    # Defensive: if caller botched the port, snap back to default 8080
    if ($Port -le 0 -or $Port -gt 65535) {
        Write-Host "  [warn] Invalid port $Port - falling back to 8080" -ForegroundColor DarkYellow
        $script:Port = 8080
    }
    # Try requested port, then scan a few adjacent ports if it's busy
    $triedPorts = @()
    foreach ($p in @($Port, 8080, 8081, 8082, 8888, 9090)) {
        if ($triedPorts -contains $p) { continue }
        $triedPorts += $p
        $listener = New-Object System.Net.HttpListener
        $prefix   = "http://localhost:$p/"
        $listener.Prefixes.Add($prefix)
        try {
            $listener.Start()
            $script:Port    = $p
            $script:BaseUrl = "http://localhost:$p"
            Write-Host ""
            Write-Host "  SDT GUI running at: $prefix" -ForegroundColor Green
            Write-Host "  (Press Ctrl+C to stop)" -ForegroundColor DarkGray
            Write-Host ""
            return $listener
        } catch {
            try { $listener.Close() } catch { }
            Write-Host "  [skip] port $p busy ($($_.Exception.Message.Split('.')[0]))" -ForegroundColor DarkGray
        }
    }
    throw "Could not bind any of: $($triedPorts -join ', '). Free one up or use -Port <n>."
}

function Send-Response {
    param(
        [System.Net.HttpListenerResponse] $Response,
        [string] $Body,
        [string] $ContentType = 'text/html; charset=utf-8',
        [int]    $StatusCode  = 200
    )
    $Response.StatusCode  = $StatusCode
    $Response.ContentType = $ContentType
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Body)
    $Response.ContentLength64 = $bytes.Length
    $Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $Response.OutputStream.Close()
}

function Send-Json {
    param([System.Net.HttpListenerResponse] $Response, $Data, [int] $StatusCode = 200)
    $json = $Data | ConvertTo-Json -Depth 10 -Compress
    Send-Response -Response $Response -Body $json -ContentType 'application/json; charset=utf-8' -StatusCode $StatusCode
}

function Read-RequestBody {
    param([System.Net.HttpListenerRequest] $Request)
    if (-not $Request.HasEntityBody) { return '' }
    $reader = New-Object System.IO.StreamReader($Request.InputStream, [System.Text.Encoding]::UTF8)
    try { return $reader.ReadToEnd() } finally { $reader.Close() }
}

# -----------------------------------------------------------------------------
# DISCOVERY WORKER - spawns Invoke-ServerDiscovery.ps1 per target, updates state
# -----------------------------------------------------------------------------
function Start-DiscoveryRun {
    param($Payload)

    $script:Session.Status     = 'running'
    $script:Session.StartedAt  = Get-Date
    $script:Session.Client     = $Payload.client
    $client                    = if ($Payload.client) { ($Payload.client -replace '[^A-Za-z0-9_-]+','_') } else { 'CLIENT' }
    $stamp                     = (Get-Date).ToString('yyyy-MM-dd-HHmm')
    $outRoot                   = if ($Payload.outputDir) { $Payload.outputDir } else { 'C:\Temp\sdt\sessions' }
    $sessionDir                = Join-Path $outRoot ("$client-$stamp")
    New-Item -ItemType Directory -Path $sessionDir -Force | Out-Null
    $script:Session.SessionDir = $sessionDir
    $script:Session.Targets    = @()
    $script:Session.LogTail.Clear()
    $script:Session.ReportPath = ''
    $script:Session.ReportZipPath = ''
    $script:Session.MissingTargets = @()
    Add-Log "Session started. Output: $sessionDir"

    # Initialize target rows
    foreach ($t in $Payload.targets) {
        $script:Session.Targets += [ordered]@{
            Name=$t; Address=$t; State='pending'; Phase=''; Buddy=''; Started=$null; Finished=$null; Kind='server'
        }
    }
    # If hypervisor provided, add a synthetic "Hypervisor" row at the top.
    $hvType = "$($Payload.hvType)"
    $hvHost = "$($Payload.hvHost)"
    if ($hvType -and $hvType -ne 'none' -and $hvHost) {
        $alreadyScanned = $script:Session.HvStagingFile -and (Test-Path $script:Session.HvStagingFile)
        $hvRow = [ordered]@{
            Name="$hvType`: $hvHost"
            Address=$hvHost
            State='pending'
            Phase=$(if ($alreadyScanned) { 'inventory already scanned' } else { 'queued' })
            Buddy=''
            Started=$null
            Finished=$null
            Kind='hypervisor'
            HvType=$hvType
            AlreadyScanned=$alreadyScanned
        }
        # Prepend
        $script:Session.Targets = @($hvRow) + @($script:Session.Targets)
        # If we already scanned, move the staging file into the session dir NOW
        if ($alreadyScanned) {
            try {
                $destName = Split-Path -Leaf $script:Session.HvStagingFile
                $dest = Join-Path $sessionDir $destName
                Move-Item -Path $script:Session.HvStagingFile -Destination $dest -Force
                if ($script:Session.HvStagingDir -and (Test-Path $script:Session.HvStagingDir)) {
                    Remove-Item $script:Session.HvStagingDir -Recurse -Force -EA SilentlyContinue
                }
                $script:Session.HvStagingFile = ''
                $script:Session.HvStagingDir = ''
                Add-Log "Hypervisor inventory moved into session from prior scan: $destName"
            } catch {
                Add-Log "Failed to move staged HV inventory: $($_.Exception.Message)"
            }
        }
    }

    # Kick off the worker scriptblock in a runspace so the HTTP listener stays responsive.
    $job = Start-ThreadJob -ScriptBlock {
        param($Session, $Payload, $ScriptDir, $BuddyFrames)

        $invoke = Join-Path $ScriptDir 'Invoke-ServerDiscovery.ps1'
        if (-not (Test-Path $invoke)) {
            $Session.Status = 'error'
            return
        }

        $winrmUser = $Payload.winrmUser
        $winrmPass = $Payload.winrmPass
        $cred = $null
        if ($winrmUser -and $winrmPass) {
            $sec  = ConvertTo-SecureString $winrmPass -AsPlainText -Force
            $cred = New-Object System.Management.Automation.PSCredential($winrmUser, $sec)
        }

        # Pick Python. Prefer portable Python the SDT ships; auto-fetch if missing
        # (only when a hypervisor target is present, since that's the only thing
        # that needs Python right now).
        $pyExe = Join-Path $ScriptDir 'python\python.exe'
        $needPy = $Session.Targets | Where-Object { $_.Kind -eq 'hypervisor' } | Select-Object -First 1
        if ($needPy -and -not (Test-Path $pyExe)) {
            $getPy = Join-Path $ScriptDir 'Get-PortablePython.ps1'
            if (Test-Path $getPy) {
                try { & $getPy 2>&1 | Out-Null } catch { }
            }
        }
        if (-not (Test-Path $pyExe)) { $pyExe = 'python' }

        # Scriptblock that does one server's work end-to-end. Runs in a child
        # ThreadJob so we can fan out across multiple servers at once.
        $serverWorker = {
            param($Session, $Payload, $ScriptDir, $BuddyFrames, $i, $cred, $invoke)
            $t = $Session.Targets[$i]
            $t.State   = 'running'
            $t.Started = (Get-Date).ToString('HH:mm:ss')
            $t.Buddy   = $BuddyFrames[(Get-Random -Maximum $BuddyFrames.Count)]
            $t.Phase   = 'connecting...'
            $Session.Targets[$i] = $t
            $safeName = ($t.Address -replace '[^A-Za-z0-9_.-]+','_')
            $perLog = Join-Path $Session.SessionDir ("server-{0}.log" -f $safeName)
            $t.LogPath = $perLog
            $Session.Targets[$i] = $t
            $lines = New-Object System.Collections.ArrayList

            try {
                # ---- Preflight reachability check ---------------------
                # If the target IS this machine, skip the WinRM probe entirely.
                # Invoke-ServerDiscovery auto-detects local execution (no WinRM
                # needed) and the probe would just print noisy "inconclusive"
                # warnings. Mark reachable so the orchestrator proceeds straight
                # to the scan.
                $isLocalTarget = ($t.Address -ieq $env:COMPUTERNAME) -or ($t.Address -ieq 'localhost') -or ($t.Address -eq '127.0.0.1') -or ($t.Address -eq '.')
                if ($isLocalTarget) {
                    $t.Phase = 'local discovery (no WinRM needed)'
                    $Session.Targets[$i] = $t
                    [void]$lines.Add("[preflight] target is local machine - skipping WinRM probe")
                    $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] $($t.Address): local target - skipping WinRM probe") | Out-Null
                    $pingOk = $true; $winrmOk = $true; $sshOk = $false
                } else {
                # 5s per probe - 1.5s was too aggressive across VPN + NetBIOS
                # name resolution. DNS alone can chew 3s.
                $t.Phase = 'preflight...'
                $Session.Targets[$i] = $t
                $pingOk  = $false; $winrmOk = $false; $sshOk = $false
                try { $pingOk = Test-Connection -ComputerName $t.Address -Count 1 -Quiet -EA SilentlyContinue } catch {}
                try {
                    $tcp = New-Object System.Net.Sockets.TcpClient
                    $ia  = $tcp.BeginConnect($t.Address, 5985, $null, $null)
                    if ($ia.AsyncWaitHandle.WaitOne(5000)) { $winrmOk = $tcp.Connected; $tcp.EndConnect($ia) }
                    $tcp.Close()
                } catch {}
                if (-not $winrmOk) {
                    try {
                        $tcp = New-Object System.Net.Sockets.TcpClient
                        $ia  = $tcp.BeginConnect($t.Address, 5986, $null, $null)
                        if ($ia.AsyncWaitHandle.WaitOne(5000)) { $winrmOk = $tcp.Connected; $tcp.EndConnect($ia) }
                        $tcp.Close()
                    } catch {}
                }
                # Second-chance check: Test-WSMan uses its own discovery/auth
                # stack and sometimes succeeds where raw TCP connects fail
                # (e.g. HTTPS-only with a custom listener cert).
                if (-not $winrmOk) {
                    try {
                        $null = Test-WSMan -ComputerName $t.Address -ErrorAction Stop
                        $winrmOk = $true
                    } catch {}
                }
                try {
                    $tcp = New-Object System.Net.Sockets.TcpClient
                    $ia  = $tcp.BeginConnect($t.Address, 22, $null, $null)
                    if ($ia.AsyncWaitHandle.WaitOne(5000)) { $sshOk = $tcp.Connected; $tcp.EndConnect($ia) }
                    $tcp.Close()
                } catch {}
                [void]$lines.Add("[preflight] ping=$pingOk winrm=$winrmOk ssh=$sshOk")
                } # end else (non-local preflight)

                # Only skip the scan in ONE case: truly unreachable (no ping,
                # no WinRM, no SSH). Every other combination - including
                # "up but WinRM probe failed" - gets attempted anyway,
                # because probes can false-negative (firewall drops SYN
                # silently, name resolution slow, etc.) and the actual
                # WinRM call will give us the real error.
                if (-not $pingOk -and -not $winrmOk -and -not $sshOk) {
                    $t.State='error'; $t.Phase='unreachable (no ping, no WinRM, no SSH)'
                    $t.Finished = (Get-Date).ToString('HH:mm:ss')
                    $Session.Targets[$i] = $t
                    try { [System.IO.File]::WriteAllText($perLog, ($lines -join "`r`n"), [System.Text.Encoding]::UTF8) } catch {}
                    $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] $($t.Address): SKIPPED - fully unreachable") | Out-Null
                    return
                }
                if ($sshOk -and -not $winrmOk -and -not $pingOk) {
                    # SSH-only, no ICMP, no WinRM - very likely Linux appliance
                    $t.State='error'; $t.Phase='Linux/SSH-only appliance - skip for WinRM run'
                    $t.Finished = (Get-Date).ToString('HH:mm:ss')
                    $Session.Targets[$i] = $t
                    try { [System.IO.File]::WriteAllText($perLog, ($lines -join "`r`n"), [System.Text.Encoding]::UTF8) } catch {}
                    $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] $($t.Address): SKIPPED - SSH-only (Linux)") | Out-Null
                    return
                }
                # Preflight couldn't confirm WinRM but the box isnt dead.
                # Press on and let Invoke-ServerDiscovery tell us for real.
                if (-not $winrmOk) {
                    [void]$lines.Add("[preflight] WinRM probe inconclusive - attempting scan anyway")
                    $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] $($t.Address): preflight inconclusive - attempting scan") | Out-Null
                }

                # ---- Actual WinRM discovery ---------------------------
                # Build a ranked list of target addresses to try. First the
                # user-provided one, then FQDN variants (derived from the
                # hypervisor host suffix when the target is short-NetBIOS),
                # then the resolved IP. First one that produces a JSON wins.
                $tryAddrs = New-Object System.Collections.ArrayList
                [void]$tryAddrs.Add($t.Address)
                $hvHost = "$($Payload.hvHost)"
                if ($t.Address -notmatch '\.' -and $hvHost -match '\.') {
                    $hvSuffix = $hvHost -replace '^[^\.]+\.',''
                    if ($hvSuffix) {
                        $fqdnGuess = "$($t.Address).$hvSuffix"
                        if ($tryAddrs -notcontains $fqdnGuess) { [void]$tryAddrs.Add($fqdnGuess) }
                    }
                }
                # Try IP as last resort
                try {
                    $ips = [System.Net.Dns]::GetHostAddresses($t.Address) 2>$null
                    $ipStr = ($ips | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1).IPAddressToString
                    if ($ipStr -and ($tryAddrs -notcontains $ipStr)) { [void]$tryAddrs.Add($ipStr) }
                } catch {}

                $addrShort = ($t.Address -split '\.',2)[0]
                $ownJson = $null
                foreach ($addr in $tryAddrs) {
                    [void]$lines.Add("[attempt] ComputerName=$addr")
                    $invokeArgs = @{
                        ComputerName   = $addr
                        OutputPath     = $Session.SessionDir
                        NonInteractive = $true
                    }
                    if ($cred) { $invokeArgs.Credential = $cred }
                    & $invoke @invokeArgs *>&1 | ForEach-Object {
                    $line = "$_"
                    [void]$lines.Add($line)
                    $stripped = $line.Trim()
                    # Stream every substantive line to the shared session log
                    # so the UI can show what a slow phase is actually doing.
                    # Filter out banner/separator/decor lines.
                    if ($stripped -and $stripped -notmatch '^[=\-_]{4,}$' -and $stripped -notmatch '^\([\^xo_\.\*\->T]+\)$') {
                        $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] [$($t.Address)] $stripped") | Out-Null
                        while ($Session.LogTail.Count -gt 2000) { try { $Session.LogTail.RemoveAt(0) } catch { break } }
                    }
                    if ($line -match '^\s*\[(\w+)\]') {
                        $t.Phase = $matches[1]
                        $t.Buddy = $BuddyFrames[(Get-Random -Maximum $BuddyFrames.Count)]
                        $Session.Targets[$i] = $t
                    }
                }
                    # End of per-attempt ForEach-Object pipeline
                    # Check for this target's own JSON after this attempt
                    $probe = @(Get-ChildItem $Session.SessionDir -Filter "$addrShort-discovery-*.json" -EA 0 |
                        Sort-Object LastWriteTime -Descending | Select-Object -First 1)
                    if ($probe) { $ownJson = $probe; break }
                    [void]$lines.Add("[attempt] no JSON yet - trying next address")
                }
                try { [System.IO.File]::WriteAllText($perLog, ($lines -join "`r`n"), [System.Text.Encoding]::UTF8) } catch {}
                if ($ownJson) {
                    $jsonFile = $ownJson[0].FullName
                    $t.JsonPath = $jsonFile
                    $t.State    = 'done'
                    $t.Phase    = "json: $(Split-Path -Leaf $jsonFile)"
                    $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] $($t.Address): JSON written - $(Split-Path -Leaf $jsonFile)") | Out-Null
                } else {
                    $tail = ($lines | Select-Object -Last 20) -join ' | '
                    $reason = 'no discovery JSON written'
                    if     ($tail -match '(?i)Access is denied')           { $reason = 'WinRM auth denied (check creds/domain)' }
                    elseif ($tail -match '(?i)cannot find')                { $reason = 'host unreachable or name not resolvable' }
                    elseif ($tail -match '(?i)WinRM|5985|5986')            { $reason = 'WinRM not reachable (port 5985/5986)' }
                    elseif ($tail -match '(?i)timed? ?out')                { $reason = 'connection timed out' }
                    elseif ($tail -match '(?i)ConvertTo-Json')             { $reason = 'data collected but JSON serialize failed - see per-server log' }
                    $t.State    = 'error'
                    $t.Phase    = $reason
                    $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] $($t.Address): NO JSON - $reason (see $(Split-Path -Leaf $perLog))") | Out-Null
                }
            } catch {
                $t.State = 'error'
                $t.Phase = $_.Exception.Message
            }
            $t.Finished = (Get-Date).ToString('HH:mm:ss')
            $Session.Targets[$i] = $t
        }

        # ---- Run hypervisor targets sequentially (usually 0 or 1), then
        # fan out server targets across $parallel ThreadJobs.
        $parallel = 4
        try { $parallel = [int]("$($Payload.parallel)") } catch {}
        if ($parallel -lt 1) { $parallel = 1 }
        if ($parallel -gt 16) { $parallel = 16 }
        $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] Discovery starting - parallelism=$parallel") | Out-Null

        for ($i = 0; $i -lt $Session.Targets.Count; $i++) {
            $t = $Session.Targets[$i]
            if ($t.Kind -ne 'hypervisor') { continue }
            $t.State   = 'running'
            $t.Started = (Get-Date).ToString('HH:mm:ss')
            $t.Buddy   = $BuddyFrames[(Get-Random -Maximum $BuddyFrames.Count)]
            $t.Phase   = 'connecting...'
            $Session.Targets[$i] = $t

            try {
                if ($t.Kind -eq 'hypervisor') {
                    # If already scanned during Setup, we already moved the file - mark done and skip
                    if ($t.AlreadyScanned) {
                        $t.State = 'done'
                        $t.Phase = 'reused scan from setup'
                        $t.Finished = (Get-Date).ToString('HH:mm:ss')
                        $Session.Targets[$i] = $t
                        continue
                    }
                    # Otherwise do a fresh hypervisor discovery via collect_vsphere_perf.py
                    $hvScript = Join-Path $ScriptDir 'collect_vsphere_perf.py'
                    if (-not (Test-Path $hvScript)) { throw "vSphere collector not found at $hvScript" }
                    $t.Phase = 'connecting to vCenter/ESXi'
                    $Session.Targets[$i] = $t
                    # Auto-retry with username format variants
                    $hvResult = Invoke-VsphereCollect -PyExe $pyExe -ScriptPath $hvScript `
                        -VCenterHost $t.Address -UserRaw $Payload.hvUser -PassRaw $Payload.hvPass -OutputDir $Session.SessionDir
                    if ($hvResult.ok -and $hvResult.user -ne $Payload.hvUser) {
                        $t.Phase = "auth OK as $($hvResult.user)"
                        $Session.Targets[$i] = $t
                    } elseif (-not $hvResult.ok) {
                        $t.State='error'
                        $t.Phase = $hvResult.error
                        $Session.Targets[$i] = $t
                    }
                    # Look for any *-inventory-*.json written into the session dir
                    $written = Get-ChildItem $Session.SessionDir -Filter '*inventory*.json' -EA 0 | Select-Object -First 1
                    if ($written) {
                        $t.State='done'; $t.Phase=("inventory saved: {0}" -f $written.Name)
                    } else {
                        $t.State='error'; $t.Phase='inventory not written (check hypervisor creds / reachability)'
                    }
                }
                $t.Finished = (Get-Date).ToString('HH:mm:ss')
            } catch {
                $t.State    = 'error'
                $t.Phase    = $_.Exception.Message
                $t.Finished = (Get-Date).ToString('HH:mm:ss')
            }
            $Session.Targets[$i] = $t
        }

        # ---- Fan out server targets across $parallel ThreadJobs ----
        $serverIndexes = @()
        for ($i = 0; $i -lt $Session.Targets.Count; $i++) {
            if ($Session.Targets[$i].Kind -ne 'hypervisor') { $serverIndexes += $i }
        }
        $jobs = [System.Collections.ArrayList]@()
        foreach ($idx in $serverIndexes) {
            # Throttle: wait until under the limit before launching another
            while (($jobs | Where-Object { $_.State -eq 'Running' }).Count -ge $parallel) {
                Start-Sleep -Milliseconds 400
            }
            $j = Start-ThreadJob -ScriptBlock $serverWorker `
                -ArgumentList $Session, $Payload, $ScriptDir, $BuddyFrames, $idx, $cred, $invoke
            [void]$jobs.Add($j)
        }
        # Wait for every server job to finish. 30-minute safety cap per job.
        foreach ($j in $jobs) {
            try {
                Wait-Job -Job $j -Timeout 1800 | Out-Null
                if ($j.State -eq 'Running') {
                    Stop-Job -Job $j -EA SilentlyContinue
                    $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] Worker job $($j.Id) exceeded 30 min timeout - stopped") | Out-Null
                }
                Receive-Job -Job $j -EA SilentlyContinue | Out-Null
            } catch {}
            try { Remove-Job -Job $j -Force -EA SilentlyContinue } catch {}
        }

        # ---------------------------------------------------------------
        # Post-session verification: confirm each target produced an
        # expected JSON. Log a compact audit so the SE sees what landed
        # and what didn't without having to dig through the folder.
        # ---------------------------------------------------------------
        $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] --- Post-session JSON audit ---") | Out-Null
        $missing = @()
        for ($i = 0; $i -lt $Session.Targets.Count; $i++) {
            $t = $Session.Targets[$i]
            if ($t.Kind -eq 'hypervisor') {
                $inv = Get-ChildItem $Session.SessionDir -Filter '*inventory*.json' -EA 0 | Select-Object -First 1
                if ($inv) {
                    $Session.LogTail.Add("   [OK] $($t.Address) (hypervisor) -> $($inv.Name)") | Out-Null
                } else {
                    $Session.LogTail.Add("   [MISS] $($t.Address) (hypervisor): $($t.Phase)") | Out-Null
                    $missing += $t
                }
            } else {
                if ($t.JsonPath -and (Test-Path $t.JsonPath)) {
                    $Session.LogTail.Add("   [OK] $($t.Address) -> $(Split-Path -Leaf $t.JsonPath)") | Out-Null
                } else {
                    $Session.LogTail.Add("   [MISS] $($t.Address): $($t.Phase)") | Out-Null
                    $missing += $t
                }
            }
        }
        if ($missing.Count -eq 0) {
            $Session.LogTail.Add("   All $($Session.Targets.Count) target(s) produced JSON. Proceeding to report.") | Out-Null
        } else {
            $Session.LogTail.Add("   $($missing.Count) of $($Session.Targets.Count) target(s) missing JSON. Report will include only what was collected.") | Out-Null
            $Session.MissingTargets = @($missing | ForEach-Object { @{ Address=$_.Address; Reason=$_.Phase; Kind=$_.Kind } })
        }

        # Attempt to generate the HTML report. Capture stdout+stderr to gen_report.log
        # in the session dir so failures are diagnosable without restarting.
        $gen = Join-Path $ScriptDir 'gen_report.py'
        $py  = Join-Path $ScriptDir 'python\python.exe'
        if (-not (Test-Path $py)) { $py = 'python' }
        $reportLog = Join-Path $Session.SessionDir 'gen_report.log'
        try {
            $mf = Join-Path $Session.SessionDir 'manifest.json'
            # gen_report.py needs a `servers` list of {file:'<discovery json name>'}
            # plus an optional inventory_file pointing at one HV inventory.
            $serversList = @()
            foreach ($t in $Session.Targets) {
                if ($t.Kind -eq 'hypervisor') { continue }
                if ($t.JsonPath -and (Test-Path $t.JsonPath)) {
                    $jname = Split-Path -Leaf $t.JsonPath
                    # Derive name + id slug from the discovery filename so gen_report.py
                    # never hits a KeyError on missing 'id'.
                    $tBase = ($jname -replace '-discovery-\d{4}-\d{2}-\d{2}.*$','') -replace '\.json$',''
                    $tSlug = ($tBase.ToLower() -replace '[^a-z0-9]+','-').Trim('-')
                    if (-not $tSlug) { $tSlug = 'server' }
                    $serversList += @{ file = $jname; name = $tBase; id = $tSlug; in_scope = $true }
                }
            }
            $invFile = ''
            $firstInv = Get-ChildItem $Session.SessionDir -Filter '*inventory*.json' -EA 0 | Select-Object -First 1
            if ($firstInv) { $invFile = $firstInv.Name }
            $manifest = @{
                client      = $Payload.client
                client_full = $Payload.client
                date        = (Get-Date).ToString('yyyy-MM-dd')
                session_dir = '.'
                output_dir  = '.'
                inventory_file = $invFile
                logo_file   = ''
                servers     = $serversList
            } | ConvertTo-Json -Depth 6
            [System.IO.File]::WriteAllText($mf, $manifest, [System.Text.Encoding]::UTF8)
            $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] manifest: $($serversList.Count) server JSON(s), inventory=$invFile") | Out-Null

            # ============================================================
            # Bulletproof report generation. 3 attempts, verify HTML after
            # each. If one approach fails, try the next before giving up.
            # ============================================================
            function Find-ReportHtml([string]$dir) {
                Get-ChildItem $dir -Filter '*DiscoveryReport*.html' -EA 0 |
                    Where-Object { $_.Length -gt 1024 } |
                    Select-Object -First 1
            }

            $combinedLog = New-Object System.Text.StringBuilder
            $Session.ReportPath = ''

            # ---- Attempt 1: portable python + manifest path as-is ----
            [void]$combinedLog.AppendLine("===== ATTEMPT 1: portable python + manifest =====")
            try {
                $out1 = & $py $gen $mf 2>&1 | Out-String
                [void]$combinedLog.AppendLine($out1)
            } catch {
                [void]$combinedLog.AppendLine("EXCEPTION: $($_.Exception.Message)")
            }
            $html = Find-ReportHtml $Session.SessionDir
            if ($html) {
                $Session.ReportPath = $html.FullName
                [void]$combinedLog.AppendLine("ATTEMPT 1 OK - HTML: $($html.Name)")
                $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] Report generated (attempt 1): $($html.Name)") | Out-Null
            }

            # ---- Attempt 2: self-heal manifest from session dir JSONs ----
            if (-not $Session.ReportPath) {
                [void]$combinedLog.AppendLine("")
                [void]$combinedLog.AppendLine("===== ATTEMPT 2: self-heal manifest + rerun =====")
                $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] No HTML yet - rebuilding manifest from session JSONs...") | Out-Null
                try {
                    $autoServers = @(Get-ChildItem $Session.SessionDir -Filter '*-discovery-*.json' -EA 0 |
                        Where-Object { $_.Length -gt 100 } |
                        ForEach-Object {
                            # Derive server name + id slug from the discovery filename.
                            # File pattern: <NAME>-discovery-<DATE>.json
                            $base = $_.BaseName -replace '-discovery-\d{4}-\d{2}-\d{2}.*$',''
                            $slug = ($base.ToLower() -replace '[^a-z0-9]+','-').Trim('-')
                            if (-not $slug) { $slug = 'server' }
                            @{ file = $_.Name; name = $base; id = $slug; in_scope = $true }
                        })
                    $autoInv = ''
                    $invf = Get-ChildItem $Session.SessionDir -Filter '*inventory*.json' -EA 0 | Where-Object { $_.Length -gt 100 } | Select-Object -First 1
                    if ($invf) { $autoInv = $invf.Name }
                    $manifest2 = @{
                        client      = $Payload.client
                        client_full = $Payload.client
                        date        = (Get-Date).ToString('yyyy-MM-dd')
                        session_dir = '.'
                        output_dir  = '.'
                        inventory_file = $autoInv
                        logo_file   = ''
                        servers     = $autoServers
                    } | ConvertTo-Json -Depth 6
                    [System.IO.File]::WriteAllText($mf, $manifest2, [System.Text.Encoding]::UTF8)
                    [void]$combinedLog.AppendLine("Rebuilt manifest: $($autoServers.Count) server(s), inventory='$autoInv'")
                    $out2 = & $py $gen $mf 2>&1 | Out-String
                    [void]$combinedLog.AppendLine($out2)
                } catch {
                    [void]$combinedLog.AppendLine("EXCEPTION: $($_.Exception.Message)")
                }
                $html = Find-ReportHtml $Session.SessionDir
                if ($html) {
                    $Session.ReportPath = $html.FullName
                    [void]$combinedLog.AppendLine("ATTEMPT 2 OK - HTML: $($html.Name)")
                    $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] Report generated after manifest self-heal: $($html.Name)") | Out-Null
                }
            }

            # ---- Attempt 3: run from session dir cwd, try system python fallback ----
            if (-not $Session.ReportPath) {
                [void]$combinedLog.AppendLine("")
                [void]$combinedLog.AppendLine("===== ATTEMPT 3: system python + cwd=session dir =====")
                $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] Still no HTML - trying system python with cwd=session dir...") | Out-Null
                $altPy = 'python'
                try {
                    $pyCmd = Get-Command python -EA SilentlyContinue
                    if ($pyCmd) { $altPy = $pyCmd.Source }
                    [void]$combinedLog.AppendLine("Using python at: $altPy")
                    $prevCwd = Get-Location
                    try {
                        Set-Location $Session.SessionDir
                        $out3 = & $altPy $gen 'manifest.json' 2>&1 | Out-String
                        [void]$combinedLog.AppendLine($out3)
                    } finally {
                        Set-Location $prevCwd
                    }
                } catch {
                    [void]$combinedLog.AppendLine("EXCEPTION: $($_.Exception.Message)")
                }
                $html = Find-ReportHtml $Session.SessionDir
                if ($html) {
                    $Session.ReportPath = $html.FullName
                    [void]$combinedLog.AppendLine("ATTEMPT 3 OK - HTML: $($html.Name)")
                    $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] Report generated via fallback: $($html.Name)") | Out-Null
                }
            }

            # Final outcome
            if (-not $Session.ReportPath) {
                [void]$combinedLog.AppendLine("")
                [void]$combinedLog.AppendLine("===== ALL 3 ATTEMPTS FAILED =====")
                $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] gen_report.py failed after 3 attempts - see gen_report.log") | Out-Null
            }
            [System.IO.File]::WriteAllText($reportLog, $combinedLog.ToString(), [System.Text.Encoding]::UTF8)
        } catch {
            $errMsg = $_.Exception.Message
            [System.IO.File]::WriteAllText($reportLog, "Exception invoking gen_report.py: $errMsg", [System.Text.Encoding]::UTF8)
            $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] gen_report.py exception: $errMsg") | Out-Null
        }

        # If the report generated successfully, bundle the whole session into
        # a zip for easy handoff. Sits next to the session folder.
        if ($Session.ReportPath) {
            try {
                $clientSlug = ($Payload.client -replace '[^A-Za-z0-9_-]+','_')
                $stampNow = (Get-Date).ToString('yyyy-MM-dd-HHmm')
                $zipName = "$clientSlug-sdt-$stampNow.zip"
                $zipPath = Join-Path (Split-Path $Session.SessionDir -Parent) $zipName
                # Compress everything in the session dir
                Compress-Archive -Path (Join-Path $Session.SessionDir '*') -DestinationPath $zipPath -Force
                if (Test-Path $zipPath) {
                    $Session.ReportZipPath = $zipPath
                    $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] Session bundled: $zipPath") | Out-Null
                }
            } catch {
                $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] Zip bundle failed: $($_.Exception.Message)") | Out-Null
            }
        }

        $Session.Status     = 'complete'
        $Session.FinishedAt = Get-Date

    } -ArgumentList $script:Session, $Payload, $script:ScriptDir, $script:BuddyFrames

    return $job
}

# -----------------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------------

# -----------------------------------------------------------------------------
# ThreadJob availability + ASYNC RUNSPACE fallback (v4.2.10+)
# -----------------------------------------------------------------------------
# ThreadJob 2.1.0+ depends on netstandard 2.0; some Windows PowerShell 5.1 hosts
# fail Import-Module with "Could not load file or assembly 'netstandard, Version=2.0.0.0'".
#
# Resolution ladder (each step wrapped in try/catch; never throws):
#   1. Preload netstandard if .NET Framework 4.7.2+ provides it.
#   2. Try Import-Module ThreadJob.
#   3. If that fails, install pinned ThreadJob 2.0.3 after removing any
#      broken 2.1.0+ copy.
#   4. If installation also fails (offline, gallery blocked, ThreatLocker),
#      install runspace-based ASYNC shims. Same shape as ThreadJob:
#      Start-ThreadJob spawns a fresh runspace in this AppDomain (so $Session
#      stays shared-by-reference with the HTTP listener) and returns IMMEDIATELY.
#      Wait-Job / Receive-Job / Stop-Job / Remove-Job proxy synthetic jobs and
#      delegate real jobs back to Microsoft.PowerShell.Core.
#
# CRITICAL: shim MUST be async. A synchronous shim blocks the HTTP listener
# thread and the GUI stalls at "Waiting for discovery to start...".
# -----------------------------------------------------------------------------
$Global:SDT_ThreadJobMode = 'unknown'

if (-not (Get-Command Start-ThreadJob -EA 0)) {
    try { Add-Type -AssemblyName 'netstandard' -EA SilentlyContinue } catch {}

    $tjPinned  = '2.0.3'
    $verboseTJ = ($env:SDT_VERBOSE_THREADJOB -eq '1')
    $threadJobOk = $false

    try {
        Import-Module ThreadJob -EA Stop
        $threadJobOk = $true
    } catch {
        if ($verboseTJ) {
            $firstLine = ($_.Exception.Message -split "`r?`n")[0]
            Write-Host "  [warn] ThreadJob import failed ($firstLine)." -ForegroundColor Yellow
            Write-Host "  [warn] Attempting install of ThreadJob v$tjPinned..." -ForegroundColor Yellow
        }
        try {
            Get-Module -ListAvailable ThreadJob | ForEach-Object {
                try { Uninstall-Module ThreadJob -RequiredVersion $_.Version -Force -EA SilentlyContinue } catch {}
            }
            try {
                if (-not (Get-PackageProvider -Name NuGet -EA SilentlyContinue)) {
                    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -EA SilentlyContinue | Out-Null
                }
                Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -EA SilentlyContinue
            } catch {}
            Install-Module ThreadJob -RequiredVersion $tjPinned -Scope CurrentUser -Force -AllowClobber -EA Stop 2>$null
            Import-Module ThreadJob -RequiredVersion $tjPinned -EA Stop
            $threadJobOk = $true
        } catch {
            if ($verboseTJ) {
                Write-Host "  [warn] ThreadJob install failed - falling back to async runspace shims." -ForegroundColor Yellow
            }
        }
    }

    if ($threadJobOk) {
        $Global:SDT_ThreadJobMode = 'parallel'
    } else {
        # ----- ASYNC RUNSPACE SHIMS -----
        # Each Start-ThreadJob call spawns a fresh runspace in the current
        # AppDomain. $Session (a synchronized hashtable) stays shared by
        # reference, so the HTTP listener sees live updates.
        $Global:SDT_ThreadJobMode = 'asyncShim'

        # Source text of the shim functions. Stored globally so child runspaces
        # can re-inject it when they recursively call Start-ThreadJob.
        $Global:SDT_ShimSource = @'
function Start-ThreadJob {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=0)][scriptblock]$ScriptBlock,
        [object[]]$ArgumentList,
        [int]$ThrottleLimit = 5,
        [string]$Name = 'SdtAsyncJob',
        [scriptblock]$InitializationScript,
        [object[]]$InputObject,
        [object]$StreamingHost
    )
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions  = 'ReuseThread'
    try { $rs.Open() } catch {
        # Runspace creation failed - degrade further to inline execution rather
        # than throw. The user sees the work happen but the listener will block
        # for the duration. Last-resort path.
        $out = $null; $errs = @()
        try {
            if ($ArgumentList) { $out = & $ScriptBlock @ArgumentList } else { $out = & $ScriptBlock }
        } catch { $errs += $_ }
        $obj = [pscustomobject]@{ Id = [Math]::Abs([Guid]::NewGuid().GetHashCode()); Name = $Name; HasMoreData = $true; _InlineOutput = $out; _InlineErrors = $errs; _InlineCompleted = $true }
        $obj | Add-Member -MemberType ScriptProperty -Name State -Value { if ($this._InlineErrors.Count) { 'Failed' } else { 'Completed' } }
        $obj.PSObject.TypeNames.Insert(0, 'Sdt.InlineJob')
        return $obj
    }
    # Push the shim source into the child runspace's globals so recursive
    # Start-ThreadJob calls inside the user's scriptblock work.
    try { $rs.SessionStateProxy.SetVariable('SDT_ShimSource', $SDT_ShimSource) } catch {}

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript('try { . ([scriptblock]::Create($SDT_ShimSource)) } catch {}')
    [void]$ps.AddStatement()
    [void]$ps.AddScript($ScriptBlock.ToString())
    if ($ArgumentList) {
        foreach ($a in $ArgumentList) { [void]$ps.AddArgument($a) }
    }
    $handle = $ps.BeginInvoke()

    $obj = [pscustomobject]@{
        Id          = [Math]::Abs([Guid]::NewGuid().GetHashCode())
        Name        = $Name
        HasMoreData = $true
        PSBeginTime = Get-Date
        _PS         = $ps
        _Handle     = $handle
        _Runspace   = $rs
    }
    $obj | Add-Member -MemberType ScriptProperty -Name State -Value {
        try {
            if ($this._Handle -and $this._Handle.IsCompleted) {
                if ($this._PS.HadErrors -or $this._PS.InvocationStateInfo.State -eq 'Failed') { 'Failed' } else { 'Completed' }
            } elseif ($this._PS.InvocationStateInfo.State -eq 'Stopped') { 'Stopped' }
            else { 'Running' }
        } catch { 'Unknown' }
    }
    $obj.PSObject.TypeNames.Insert(0, 'Sdt.AsyncJob')
    return $obj
}

function Wait-Job {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline=$true, Position=0)][object[]]$Job,
        [int]$Timeout = -1,
        [switch]$Any,
        [switch]$Force
    )
    process {
        foreach ($j in @($Job)) {
            if ($j -and ($j.PSObject.TypeNames -contains 'Sdt.AsyncJob')) {
                try {
                    if ($Timeout -gt 0) {
                        [void]$j._Handle.AsyncWaitHandle.WaitOne([int]($Timeout * 1000))
                    } else {
                        [void]$j._Handle.AsyncWaitHandle.WaitOne()
                    }
                } catch {}
                $j
            } elseif ($j -and ($j.PSObject.TypeNames -contains 'Sdt.InlineJob')) {
                $j
            } else {
                Microsoft.PowerShell.Core\Wait-Job -Job $j -Timeout $Timeout -Any:$Any -Force:$Force
            }
        }
    }
}

function Receive-Job {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline=$true, Position=0)][object[]]$Job,
        [switch]$Keep,
        [switch]$Wait,
        [switch]$AutoRemoveJob
    )
    process {
        foreach ($j in @($Job)) {
            if ($j -and ($j.PSObject.TypeNames -contains 'Sdt.AsyncJob')) {
                try {
                    if (-not $j._Handle.IsCompleted -and $Wait) {
                        [void]$j._Handle.AsyncWaitHandle.WaitOne()
                    }
                    if ($j._Handle.IsCompleted) {
                        $out = $null
                        try { $out = $j._PS.EndInvoke($j._Handle) } catch { Write-Error $_ }
                        foreach ($e in $j._PS.Streams.Error) { Write-Error $e }
                        $out
                        if (-not $Keep) { $j.HasMoreData = $false }
                    }
                } catch {}
            } elseif ($j -and ($j.PSObject.TypeNames -contains 'Sdt.InlineJob')) {
                foreach ($e in $j._InlineErrors) { Write-Error $e }
                $j._InlineOutput
                if (-not $Keep) { $j.HasMoreData = $false }
            } else {
                Microsoft.PowerShell.Core\Receive-Job -Job $j -Keep:$Keep -Wait:$Wait -AutoRemoveJob:$AutoRemoveJob
            }
        }
    }
}

function Stop-Job {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true, Position=0)][object[]]$Job, [switch]$PassThru)
    process {
        foreach ($j in @($Job)) {
            if ($j -and ($j.PSObject.TypeNames -contains 'Sdt.AsyncJob')) {
                try { if ($j._PS -and -not $j._Handle.IsCompleted) { $j._PS.Stop() } } catch {}
                if ($PassThru) { $j }
            } elseif ($j -and ($j.PSObject.TypeNames -contains 'Sdt.InlineJob')) {
                if ($PassThru) { $j }
            } else {
                Microsoft.PowerShell.Core\Stop-Job -Job $j -PassThru:$PassThru
            }
        }
    }
}

function Remove-Job {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true, Position=0)][object[]]$Job, [switch]$Force)
    process {
        foreach ($j in @($Job)) {
            if ($j -and ($j.PSObject.TypeNames -contains 'Sdt.AsyncJob')) {
                try { if ($j._PS -and -not $j._Handle.IsCompleted -and $Force) { $j._PS.Stop() } } catch {}
                try { $j._PS.Dispose() } catch {}
                try { $j._Runspace.Close() } catch {}
                try { $j._Runspace.Dispose() } catch {}
            } elseif ($j -and ($j.PSObject.TypeNames -contains 'Sdt.InlineJob')) {
                # no-op
            } else {
                Microsoft.PowerShell.Core\Remove-Job -Job $j -Force:$Force
            }
        }
    }
}
'@

        # Dot-source the shim into the parent's global scope.
        try {
            . ([scriptblock]::Create($Global:SDT_ShimSource))
        } catch {
            Write-Host "  [error] Could not load runspace shim: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "  [error] Discovery will not work. Reinstall ThreadJob manually:" -ForegroundColor Red
            Write-Host "          Install-Module ThreadJob -RequiredVersion 2.0.3 -Scope CurrentUser -Force" -ForegroundColor Yellow
            $Global:SDT_ThreadJobMode = 'broken'
        }

        # Self-check: spawn a tiny async job, wait with HARD timeout. If the
        # runspace hangs (e.g. ThreatLocker blocking .NET API), we abort the
        # PowerShell handle and fall back to inline execution rather than
        # letting EndInvoke block the startup forever.
        if ($Global:SDT_ThreadJobMode -eq 'asyncShim') {
            $selfTestPassed = $false
            $testJob = $null
            try {
                $testJob = Start-ThreadJob -ScriptBlock { param($x) "ok-$x" } -ArgumentList 'shim'
                $completed = $testJob._Handle.AsyncWaitHandle.WaitOne(3000)
                if (-not $completed) {
                    throw "self-test timed out after 3s (runspace hung)"
                }
                $testOut = $testJob._PS.EndInvoke($testJob._Handle)
                if ("$testOut" -ne 'ok-shim') {
                    throw "self-test produced '$testOut' (expected 'ok-shim')"
                }
                $selfTestPassed = $true
            } catch {
                Write-Host "  [warn] Runspace shim self-test FAILED: $($_.Exception.Message)" -ForegroundColor Yellow
                Write-Host "  [warn] Falling back to inline execution (discovery will block listener)." -ForegroundColor Yellow
                $Global:SDT_ThreadJobMode = 'inlineLastResort'
            } finally {
                # Always try to clean up the test runspace, even on timeout.
                if ($testJob) {
                    try { if (-not $testJob._Handle.IsCompleted) { $testJob._PS.Stop() } } catch {}
                    try { $testJob._PS.Dispose() } catch {}
                    try { $testJob._Runspace.Dispose() } catch {}
                }
            }
            if (-not $selfTestPassed -and $Global:SDT_ThreadJobMode -eq 'inlineLastResort') {
                # Replace the runspace-based Start-ThreadJob with an inline
                # synchronous version. This blocks the caller but never hangs.
                # Wait-Job / Receive-Job / etc. handle the synthetic InlineJob.
                Remove-Item Function:\Start-ThreadJob -EA SilentlyContinue
                function global:Start-ThreadJob {
                    [CmdletBinding()]
                    param(
                        [Parameter(Mandatory=$true, Position=0)][scriptblock]$ScriptBlock,
                        [object[]]$ArgumentList,
                        [int]$ThrottleLimit = 1,
                        [string]$Name = 'SdtInlineJob',
                        [scriptblock]$InitializationScript
                    )
                    $out = $null; $errs = @()
                    try {
                        if ($InitializationScript) { . $InitializationScript }
                        if ($ArgumentList -and $ArgumentList.Count -gt 0) {
                            $out = & $ScriptBlock @ArgumentList
                        } else {
                            $out = & $ScriptBlock
                        }
                    } catch { $errs += $_ }
                    $obj = [pscustomobject]@{
                        Id = [Math]::Abs([Guid]::NewGuid().GetHashCode())
                        Name = $Name; HasMoreData = $true
                        _InlineOutput = $out; _InlineErrors = $errs; _InlineCompleted = $true
                    }
                    $obj | Add-Member -MemberType ScriptProperty -Name State -Value { if ($this._InlineErrors.Count) { 'Failed' } else { 'Completed' } }
                    $obj.PSObject.TypeNames.Insert(0, 'Sdt.InlineJob')
                    return $obj
                }
            }
        }
    }
}

# Single quiet line: ThreadJob status.
if ($Global:SDT_ThreadJobMode -eq 'asyncShim') {
    Write-Host "  [sdt] ThreadJob unavailable - using runspace shim (in-process async)" -ForegroundColor DarkGray
} elseif ($Global:SDT_ThreadJobMode -eq 'inlineLastResort') {
    Write-Host "  [sdt] WARNING: runspaces also blocked - using inline mode (UI may freeze during scan)." -ForegroundColor Yellow
} elseif ($Global:SDT_ThreadJobMode -eq 'broken') {
    Write-Host "  [sdt] WARNING: discovery primitives broken. Fix: Install-Module ThreadJob -RequiredVersion 2.0.3 -Scope CurrentUser -Force" -ForegroundColor Red
}

# Auto-minimize the host PowerShell window so the GUI takes focus.
try {
    Add-Type -Name _SdtWin -Namespace _Sdt -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool ShowWindow(System.IntPtr h, int s);
[System.Runtime.InteropServices.DllImport("kernel32.dll")]
public static extern System.IntPtr GetConsoleWindow();
'@ -ErrorAction Stop
    $h = [_Sdt._SdtWin]::GetConsoleWindow()
    if ($h -ne [System.IntPtr]::Zero) {
        [void][_Sdt._SdtWin]::ShowWindow($h, 6)  # SW_MINIMIZE
    }
} catch { }

$listener = Start-HttpListener

# Open browser - prefer Edge, then Chrome, fall back to default.
# Use --app mode for a clean window without browser chrome.
function Find-ModernBrowser {
    $candidates = @(
        # Edge (Chromium) - standard install locations
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe",
        # Chrome - standard install locations
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "$env:ProgramFiles(x86)\Google\Chrome\Application\chrome.exe",
        "$env:LocalAppData\Google\Chrome\Application\chrome.exe"
    )
    foreach ($p in $candidates) {
        if ($p -and (Test-Path $p)) { return $p }
    }
    # Registry fallback for App Paths
    foreach ($reg in @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe'
    )) {
        try {
            $v = (Get-ItemProperty -Path $reg -Name '(Default)' -EA Stop).'(Default)'
            if ($v -and (Test-Path $v)) { return $v }
        } catch { }
    }
    return $null
}

if (-not $NoOpenBrowser) {
    $browser = Find-ModernBrowser
    try {
        if ($browser) {
            Write-Host "  Launching: $browser" -ForegroundColor DarkGray
            # --app mode strips tab bar/URL bar for a native-feel window
            Start-Process -FilePath $browser -ArgumentList "--app=$script:BaseUrl" | Out-Null
        } else {
            Write-Host "  No Edge/Chrome found; using system default browser." -ForegroundColor DarkYellow
            Start-Process $script:BaseUrl | Out-Null
        }
    } catch {
        # Last-resort fallback
        try { Start-Process $script:BaseUrl | Out-Null } catch { }
    }
}

# Request loop
try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $req     = $context.Request
        $resp    = $context.Response
        $path    = $req.Url.AbsolutePath
        $method  = $req.HttpMethod

        try {
            switch -Regex ("$method $path") {
                '^GET /api/status$' {
                    # Wrapped so a single malformed field (e.g. XML in a Phase
                    # string) can never 500 the status poll and break the UI.
                    try {
                        $safeTargets = @()
                        foreach ($t in $script:Session.Targets) {
                            $copy = [ordered]@{}
                            foreach ($k in $t.Keys) {
                                $v = $t[$k]
                                if ($v -is [string] -and $v.Length -gt 500) { $v = $v.Substring(0,497) + '...' }
                                $copy[$k] = $v
                            }
                            $safeTargets += $copy
                        }
                        Send-Json -Response $resp -Data @{
                            Status         = $script:Session.Status
                            Client         = $script:Session.Client
                            SessionDir     = $script:Session.SessionDir
                            Targets        = $safeTargets
                            LogTail        = @($script:Session.LogTail)
                            ReportPath     = $script:Session.ReportPath
                            ReportZipPath  = $script:Session.ReportZipPath
                            MissingTargets = @($script:Session.MissingTargets)
                        }
                    } catch {
                        # Minimal fallback so polling never breaks
                        Send-Json -Response $resp -Data @{
                            Status = $script:Session.Status
                            error  = "status serialize failed: $($_.Exception.Message)"
                        }
                    }
                    break
                }
                '^POST /api/hv-scan$' {
                    $body = Read-RequestBody -Request $req
                    $hv = $null
                    try { $hv = $body | ConvertFrom-Json -ErrorAction Stop } catch {
                        Send-Json -Response $resp -Data @{ ok=$false; error="JSON parse failed" } -StatusCode 400
                        break
                    }
                    if (-not $hv.hvHost -or -not $hv.hvUser -or -not $hv.hvPass) {
                        Send-Json -Response $resp -Data @{ ok=$false; error="hvHost, hvUser, hvPass required" } -StatusCode 400
                        break
                    }
                    # Scan into a temp staging dir so we don't commit to a session folder yet
                    $stageDir = Join-Path $env:TEMP ("sdt-hvscan-" + [guid]::NewGuid().ToString('N').Substring(0,8))
                    New-Item -ItemType Directory -Path $stageDir -Force | Out-Null
                    try {
                        $hvScript = Join-Path $script:ScriptDir 'collect_vsphere_perf.py'
                        if (-not (Test-Path $hvScript)) { throw "collect_vsphere_perf.py not found" }
                        $pyExe = Join-Path $script:ScriptDir 'python\python.exe'
                        if (-not (Test-Path $pyExe)) {
                            $getPy = Join-Path $script:ScriptDir 'Get-PortablePython.ps1'
                            if (Test-Path $getPy) { try { & $getPy 2>&1 | Out-Null } catch {} }
                        }
                        if (-not (Test-Path $pyExe)) { $pyExe = 'python' }
                        # Invoke collector with auto-retry across username formats
                        $collectResult = Invoke-VsphereCollect -PyExe $pyExe -ScriptPath $hvScript `
                            -VCenterHost $hv.hvHost -UserRaw $hv.hvUser -PassRaw $hv.hvPass -OutputDir $stageDir

                        if (-not $collectResult.ok) {
                            $errMsg = "Collector failed: $($collectResult.error)"
                            if ($collectResult.triedUsers) {
                                $errMsg += ". Tried usernames: $($collectResult.triedUsers -join ', ')"
                            }
                            Add-Log "vSphere scan FAILED: $($collectResult.pyError -or $collectResult.error). Log: $($collectResult.logPath)"
                            Send-Json -Response $resp -Data @{
                                ok=$false; error=$errMsg
                                log=$collectResult.log
                                pyError=$collectResult.pyError
                                logPath=$collectResult.logPath
                            } -StatusCode 500
                            break
                        }
                        $outFile = Get-Item $collectResult.file
                        if ($collectResult.user -ne $hv.hvUser) {
                            Add-Log ("Hypervisor auth succeeded with adjusted username: {0}" -f $collectResult.user)
                        }
                        $raw = [System.IO.File]::ReadAllText($outFile.FullName, [System.Text.Encoding]::UTF8)
                        $doc = $raw | ConvertFrom-Json
                        # Remember staging so /api/start can move it into the final session dir
                        $script:Session.HvStagingFile = $outFile.FullName
                        $script:Session.HvStagingDir  = $stageDir
                        # Shape VMs for UI
                        $vms = @()
                        foreach ($v in $doc.VMs) {
                            $vms += @{
                                Name       = $v.Name
                                IPs        = $v.IPs
                                GuestOS    = $v.GuestOS
                                PowerState = $v.PowerState
                            }
                        }
                        Send-Json -Response $resp -Data @{ ok=$true; vms=$vms; count=$vms.Count; stagingFile=$outFile.FullName }
                    } catch {
                        # Always write a log file even when the catch fires before the collector runs.
                        $errLog = ''
                        try {
                            if (Test-Path $stageDir) {
                                $stamp = (Get-Date).ToString('yyyy-MM-dd-HHmmss')
                                $errLog = Join-Path $stageDir "vsphere-scan-error-$stamp.log"
                                $errBody = "Endpoint exception:`n$($_.Exception.Message)`n`nStackTrace:`n$($_.ScriptStackTrace)`n`nFullRecord:`n$($_ | Out-String)"
                                [System.IO.File]::WriteAllText($errLog, $errBody, [System.Text.Encoding]::UTF8)
                            }
                        } catch { }
                        Add-Log "vSphere scan endpoint FAILED: $($_.Exception.Message). Log: $errLog"
                        Send-Json -Response $resp -Data @{
                            ok=$false
                            error=$_.Exception.Message
                            log="$($_.Exception.Message)`n`n$($_.ScriptStackTrace)"
                            logPath=$errLog
                        } -StatusCode 500
                    }
                    break
                }
                '^POST /api/hv-local-scan$' {
                    # Hyper-V LOCAL mode -- preflight + Get-VM on this machine.
                    # No remote host, no creds.  Body: { force: bool } (optional).
                    $body = Read-RequestBody -Request $req
                    $req2 = $null
                    try { if ($body) { $req2 = $body | ConvertFrom-Json -EA Stop } } catch {}
                    $force = $false
                    if ($req2 -and $req2.PSObject.Properties.Name -contains 'force') { $force = [bool]$req2.force }

                    $pf = Test-IsLocalHyperVHost
                    if (-not $pf.IsHost -and -not $force) {
                        Send-Json -Response $resp -Data @{
                            ok=$false; preflight=$true
                            isHost=$false; passCount=$pf.PassCount
                            reasons=@($pf.Reasons); detail=$pf.Detail
                            error="This machine is not a Hyper-V host. Fix the listed signals and retry, or pass force=true to try anyway."
                        } -StatusCode 200   # 200 so JS can render the retry UI without treating it as a network error
                        break
                    }

                    $stageDir = Join-Path $env:TEMP ("sdt-hvlocal-" + [guid]::NewGuid().ToString('N').Substring(0,8))
                    New-Item -ItemType Directory -Path $stageDir -Force | Out-Null
                    try {
                        $res = Invoke-LocalHyperVInventory -OutputDir $stageDir -Force:$force
                        if (-not $res.ok) {
                            Send-Json -Response $resp -Data @{
                                ok=$false; preflight=$false
                                isHost=[bool]$pf.IsHost; passCount=$pf.PassCount
                                error=$res.error
                                reasons=@($pf.Reasons); detail=$pf.Detail
                            } -StatusCode 500
                            break
                        }
                        $script:Session.HvStagingFile = $res.file
                        $script:Session.HvStagingDir  = $stageDir
                        $vmsOut = @()
                        foreach ($v in $res.vms) {
                            $guestOsVal = 'Windows (Hyper-V guest)'
                            try { if ($v.PSObject.Properties.Name -contains 'GuestOS' -and $v.GuestOS) { $guestOsVal = $v.GuestOS } } catch {}
                            $vmsOut += @{
                                Name       = $v.Name
                                IPs        = $v.IPs
                                GuestOS    = $guestOsVal
                                PowerState = $v.State
                            }
                        }
                        Send-Json -Response $resp -Data @{
                            ok=$true; mode='hyperv-local'
                            host=$env:COMPUTERNAME
                            passCount=$pf.PassCount
                            forced=[bool]$force
                            vms=$vmsOut; count=$vmsOut.Count
                            stagingFile=$res.file
                        }
                    } catch {
                        Send-Json -Response $resp -Data @{ ok=$false; error=$_.Exception.Message } -StatusCode 500
                    }
                    break
                }
                '^POST /api/local-only$' {
                    # No-HV, no-creds, single-target = THIS machine.
                    # Used when the SE is RDP-ed into a server and just needs that
                    # one box scanned + a report. Bypasses the form entirely.
                    if ($script:Session.Status -eq 'running') {
                        Send-Json -Response $resp -Data @{ ok = $false; error = 'A discovery session is already running.' } -StatusCode 409
                        break
                    }
                    $localTarget = $env:COMPUTERNAME
                    $payload = [pscustomobject]@{
                        client    = $localTarget
                        outputDir = 'C:\Temp\sdt\sessions'
                        parallel  = 1
                        targets   = @($localTarget)
                        winrmUser = ''
                        winrmPass = ''
                        hvType    = 'none'
                        hvHost    = ''
                        hvUser    = ''
                        hvPass    = ''
                        localOnly = $true
                    }
                    try {
                        $null = Start-DiscoveryRun -Payload $payload
                        Send-Json -Response $resp -Data @{ ok = $true; sessionDir = $script:Session.SessionDir; target = $localTarget }
                    } catch {
                        Send-Json -Response $resp -Data @{ ok = $false; error = $_.Exception.Message } -StatusCode 500
                    }
                    break
                }
                '^POST /api/start$' {
                    if ($script:Session.Status -eq 'running') {
                        Send-Json -Response $resp -Data @{ error = 'A discovery session is already running.' } -StatusCode 409
                        break
                    }
                    $body = Read-RequestBody -Request $req
                    $payload = $null
                    try { $payload = $body | ConvertFrom-Json -ErrorAction Stop } catch {
                        Send-Json -Response $resp -Data @{ error = "JSON parse failed: $($_.Exception.Message)"; bodyLen = $body.Length } -StatusCode 400
                        break
                    }
                    # Normalize targets: PS ConvertFrom-Json may return $null, string,
                    # or Object[]. Force into an array of non-empty strings.
                    $raw = $null
                    if ($payload) {
                        try { $raw = $payload.targets } catch { $raw = $null }
                    }
                    $targets = @()
                    if ($raw -is [string] -and $raw.Trim()) {
                        $targets = @($raw -split '\r?\n' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                    } elseif ($raw) {
                        $targets = @(@($raw) | ForEach-Object { "$_".Trim() } | Where-Object { $_ })
                    }
                    # Blank targets OK if a hypervisor is set - we'll discover VMs from it.
                    $hvType = ''
                    try { $hvType = "$($payload.hvType)" } catch { $hvType = '' }
                    $hvHost = ''
                    try { $hvHost = "$($payload.hvHost)" } catch { $hvHost = '' }
                    # Hyper-V local mode: no remote host required -- staging file from
                    # /api/hv-local-scan is the proof of a valid hypervisor input.
                    $hvLocalReady = ($hvType -eq 'hyperv') -and $script:Session.HvStagingFile -and (Test-Path $script:Session.HvStagingFile)
                    if ($targets.Count -eq 0 -and ($hvType -eq '' -or $hvType -eq 'none' -or (-not $hvHost -and -not $hvLocalReady))) {
                        Send-Json -Response $resp -Data @{
                            error = 'No targets and no hypervisor. Provide at least one target host OR hypervisor details (run Scan Hypervisor first for Hyper-V local mode).'
                        } -StatusCode 400
                        break
                    }
                    if ($hvType -eq 'hyperv' -and -not $hvHost) {
                        $payload | Add-Member -NotePropertyName hvHost -NotePropertyValue $env:COMPUTERNAME -Force
                    }
                    # Replace parsed targets back into payload for worker
                    $payload | Add-Member -NotePropertyName targets -NotePropertyValue $targets -Force
                    $null = Start-DiscoveryRun -Payload $payload
                    Send-Json -Response $resp -Data @{ ok = $true; sessionDir = $script:Session.SessionDir; targetCount = $targets.Count }
                    break
                }
                '^POST /api/test-creds$' {
                    $body = Read-RequestBody -Request $req
                    $c = $null
                    try { $c = $body | ConvertFrom-Json -ErrorAction Stop } catch {
                        Send-Json -Response $resp -Data @{ ok=$false; error='JSON parse failed' } -StatusCode 400; break
                    }
                    $u = "$($c.winrmUser)".Trim()
                    $p = "$($c.winrmPass)"
                    if (-not $u -or -not $p) {
                        Send-Json -Response $resp -Data @{ ok=$false; error='username + password required' } -StatusCode 400; break
                    }
                    $dom = $null; $justUser = $null
                    if ($u -match '^\.\\(.+)$') {
                        Send-Json -Response $resp -Data @{ ok=$false; soft=$true; error='Local account (.\user) - will be tested during discovery.' }; break
                    } elseif ($u -match '^([^\\]+)\\(.+)$') {
                        $dom = $matches[1]; $justUser = $matches[2]
                    } elseif ($u -like '*@*') {
                        $dom = ''; $justUser = $u   # UPN: LogonUser wants domain="", user=UPN
                    } else {
                        Send-Json -Response $resp -Data @{ ok=$false; error='Username needs a domain: DOMAIN\user or user@domain.local' }; break
                    }

                    # --- Primary: Win32 LogonUser (advapi32.dll) --------------
                    # Works anywhere the box can talk to a DC for that domain -
                    # on the DC itself, on any domain-joined member, or across a
                    # trust. No ADWS, no DC discovery dance, no LDAP bind quirks.
                    $usedAPI = 'LogonUser'
                    try {
                        if (-not ('Win32.NativeCreds' -as [type])) {
                            Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace Win32 {
    public static class NativeCreds {
        [DllImport("advapi32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
        public static extern bool LogonUserW(string user, string domain, string password,
            int logonType, int logonProvider, out IntPtr token);
        [DllImport("kernel32.dll", SetLastError=true)]
        public static extern bool CloseHandle(IntPtr h);
    }
}
'@ -ErrorAction Stop
                        }
                        $token = [IntPtr]::Zero
                        # LOGON32_LOGON_NETWORK=3, LOGON32_PROVIDER_DEFAULT=0
                        $ok = [Win32.NativeCreds]::LogonUserW($justUser, $dom, $p, 3, 0, [ref]$token)
                        $err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
                        if ($ok) {
                            [void][Win32.NativeCreds]::CloseHandle($token)
                            Send-Json -Response $resp -Data @{ ok=$true; message=("Valid on domain {0} (via {1})" -f ($(if($dom){$dom}else{'UPN'})), $usedAPI) }
                            break
                        }
                        # Map common errors to friendly messages
                        $hard = $true; $soft = $false
                        $why = switch ($err) {
                            1326 { 'Wrong password (LOGON_FAILURE)' }
                            1327 { 'Account restriction (may need password change, login hours, etc.)' }
                            1328 { 'Login not permitted at this time (login hours restriction)' }
                            1329 { 'Account cannot log on from this workstation' }
                            1330 { 'Password expired' }
                            1331 { 'Account disabled' }
                            1909 { 'Account locked out' }
                            1789 { $hard=$false; $soft=$true; 'Cannot reach a DC for domain from this box' }
                            1355 { $hard=$false; $soft=$true; 'Specified domain does not exist / no DC reachable' }
                            1317 { 'Unknown user name' }
                            default { "LogonUser error $err" }
                        }
                        if ($soft) {
                            Send-Json -Response $resp -Data @{ ok=$false; soft=$true; error="$why - will be tested during scan"; winErr=$err }
                        } else {
                            Send-Json -Response $resp -Data @{ ok=$false; error=$why; winErr=$err }
                        }
                        break
                    } catch {
                        # LogonUser itself threw - fall through to LDAP as backup
                    }

                    # --- Backup 2: AccountManagement LDAP bind -----------------
                    $lookup = if ($dom) { $dom } else { ($u -split '@',2)[1] }
                    try {
                        Add-Type -AssemblyName System.DirectoryServices.AccountManagement -ErrorAction Stop
                        $ctx = New-Object System.DirectoryServices.AccountManagement.PrincipalContext(
                            [System.DirectoryServices.AccountManagement.ContextType]::Domain, $lookup)
                        $valid = $ctx.ValidateCredentials($u, $p)
                        $ctx.Dispose()
                        if ($valid) {
                            Send-Json -Response $resp -Data @{ ok=$true; message="Valid on $lookup (via LDAP PrincipalContext)" }
                            break
                        } else {
                            Send-Json -Response $resp -Data @{ ok=$false; error="Rejected by $lookup (wrong password, expired, or locked)" }
                            break
                        }
                    } catch {
                        # fall through to DirectoryEntry
                    }

                    # --- Backup 3: Direct LDAP bind via DirectoryEntry ---------
                    try {
                        $domForDn = $lookup
                        if ($domForDn -notmatch '\.') { $domForDn = "$domForDn.local" }
                        $dn = "DC=" + ($domForDn -split '\.' -join ',DC=')
                        $de = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$dn", $u, $p)
                        $null = $de.NativeObject   # triggers the bind
                        $de.Close()
                        Send-Json -Response $resp -Data @{ ok=$true; message="Valid on $lookup (via DirectoryEntry LDAP bind)" }
                        break
                    } catch {
                        $em = $_.Exception.Message
                        if ($em -match '(?i)invalid credentials|logon failure|8009030C|525|52e') {
                            Send-Json -Response $resp -Data @{ ok=$false; error="Rejected by $lookup (wrong password)" }
                            break
                        }
                        # Otherwise treat as unreachable -> soft
                        Send-Json -Response $resp -Data @{
                            ok=$false; soft=$true
                            error="Cannot reach any DC for '$lookup' from this box (tried LogonUser, PrincipalContext, DirectoryEntry). Creds will be tested during the actual scan."
                        }
                    }
                    break
                }
                '^POST /api/open-folder$' {
                    $dir = $script:Session.SessionDir
                    if (-not $dir -or -not (Test-Path $dir)) {
                        Send-Json -Response $resp -Data @{ error = "Session folder does not exist: $dir" } -StatusCode 404
                        break
                    }
                    try {
                        Start-Process explorer.exe $dir | Out-Null
                        Send-Json -Response $resp -Data @{ ok = $true; path = $dir }
                    } catch {
                        Send-Json -Response $resp -Data @{ error = $_.Exception.Message } -StatusCode 500
                    }
                    break
                }
                '^GET /api/report-html$' {
                    # Stream the generated HTML report through the local HTTP
                    # server. Chrome blocks file:// from http:// origins, so we
                    # must serve it ourselves.
                    $rp = $script:Session.ReportPath
                    if (-not $rp -or -not (Test-Path $rp)) {
                        Send-Response -Response $resp -Body '<h1>No report generated yet</h1>' -StatusCode 404
                        break
                    }
                    try {
                        $bytes = [System.IO.File]::ReadAllBytes($rp)
                        $resp.StatusCode = 200
                        $resp.ContentType = 'text/html; charset=utf-8'
                        $resp.ContentLength64 = $bytes.Length
                        $resp.OutputStream.Write($bytes, 0, $bytes.Length)
                        $resp.OutputStream.Close()
                    } catch {
                        Send-Response -Response $resp -Body "<h1>Error reading report: $($_.Exception.Message)</h1>" -StatusCode 500
                    }
                    break
                }
                '^POST /api/regenerate-report$' {
                    # Self-heal: rebuild manifest from any *-discovery-*.json files
                    # in the session dir and re-run gen_report.py. Used by the
                    # "Retry Report Generation" button on the Report tab.
                    $dir = $script:Session.SessionDir
                    if (-not $dir -or -not (Test-Path $dir)) {
                        Send-Json -Response $resp -Data @{ ok=$false; error='no session dir' } -StatusCode 400
                        break
                    }
                    try {
                        $gen = Join-Path $script:ScriptDir 'gen_report.py'
                        $py  = Join-Path $script:ScriptDir 'python\python.exe'
                        if (-not (Test-Path $py)) { $py = 'python' }
                        $mf  = Join-Path $dir 'manifest.json'
                        $autoServers = @(Get-ChildItem $dir -Filter '*-discovery-*.json' -EA 0 |
                            Where-Object { $_.Length -gt 100 } |
                            ForEach-Object {
                                $base = $_.BaseName -replace '-discovery-\d{4}-\d{2}-\d{2}.*$',''
                                $slug = ($base.ToLower() -replace '[^a-z0-9]+','-').Trim('-')
                                if (-not $slug) { $slug = 'server' }
                                @{ file = $_.Name; name = $base; id = $slug; in_scope = $true }
                            })
                        $autoInv = ''
                        $invf = Get-ChildItem $dir -Filter '*inventory*.json' -EA 0 | Where-Object { $_.Length -gt 100 } | Select-Object -First 1
                        if ($invf) { $autoInv = $invf.Name }
                        $client = if ($script:Session.Client) { $script:Session.Client } else { 'CLIENT' }
                        $manifestObj = @{
                            client = $client; client_full = $client
                            date = (Get-Date).ToString('yyyy-MM-dd')
                            session_dir = '.'; output_dir = '.'
                            inventory_file = $autoInv; logo_file = ''
                            servers = $autoServers
                        } | ConvertTo-Json -Depth 6
                        [System.IO.File]::WriteAllText($mf, $manifestObj, [System.Text.Encoding]::UTF8)
                        $stdOut = [System.IO.Path]::GetTempFileName()
                        $stdErr = [System.IO.Path]::GetTempFileName()
                        $p = Start-Process -FilePath $py -ArgumentList @($gen, $mf) `
                            -NoNewWindow -PassThru -Wait `
                            -RedirectStandardOutput $stdOut -RedirectStandardError $stdErr `
                            -WorkingDirectory $dir -EA Stop
                        $log = "[exit $($p.ExitCode)]`n" + (Get-Content $stdOut -Raw) + "`n--- STDERR ---`n" + (Get-Content $stdErr -Raw)
                        Remove-Item $stdOut, $stdErr -EA SilentlyContinue
                        # Append to gen_report.log so user can see the retry
                        $reportLog = Join-Path $dir 'gen_report.log'
                        Add-Content -Path $reportLog -Value "`n===== RETRY ($(Get-Date -f 'yyyy-MM-dd HH:mm:ss')) =====`n$log" -EA SilentlyContinue
                        $html = Get-ChildItem $dir -Filter '*DiscoveryReport*.html' -EA 0 | Where-Object { $_.Length -gt 1024 } | Select-Object -First 1
                        if ($html) {
                            $script:Session.ReportPath = $html.FullName
                            Add-Log "Report regenerated: $($html.Name)"
                            Send-Json -Response $resp -Data @{ ok=$true; reportPath=$html.FullName }
                        } else {
                            Send-Json -Response $resp -Data @{ ok=$false; error='gen_report.py ran but no HTML produced'; log=$log } -StatusCode 500
                        }
                    } catch {
                        Send-Json -Response $resp -Data @{ ok=$false; error=$_.Exception.Message } -StatusCode 500
                    }
                    break
                }
                '^GET /api/gen-log$' {
                    $dir = $script:Session.SessionDir
                    if (-not $dir) {
                        Send-Json -Response $resp -Data @{ error = 'no session dir' } -StatusCode 404
                        break
                    }
                    $log = Join-Path $dir 'gen_report.log'
                    if (-not (Test-Path $log)) {
                        Send-Json -Response $resp -Data @{ error = "gen_report.log not found in $dir" } -StatusCode 404
                        break
                    }
                    try {
                        $content = [System.IO.File]::ReadAllText($log, [System.Text.Encoding]::UTF8)
                        Send-Json -Response $resp -Data @{ ok = $true; content = $content; path = $log }
                    } catch {
                        Send-Json -Response $resp -Data @{ error = $_.Exception.Message } -StatusCode 500
                    }
                    break
                }
                '^GET /api/combined-logs$' {
                    $dir = $script:Session.SessionDir
                    $sb  = New-Object System.Text.StringBuilder
                    [void]$sb.AppendLine("===== SDT Combined Logs =====")
                    [void]$sb.AppendLine("Client      : $($script:Session.Client)")
                    [void]$sb.AppendLine("SessionDir  : $dir")
                    [void]$sb.AppendLine("Status      : $($script:Session.Status)")
                    [void]$sb.AppendLine("ReportPath  : $($script:Session.ReportPath)")
                    [void]$sb.AppendLine("ReportZip   : $($script:Session.ReportZipPath)")
                    [void]$sb.AppendLine("Targets     :")
                    foreach ($t in $script:Session.Targets) {
                        $js = if ($t.JsonPath) { Split-Path -Leaf $t.JsonPath } else { '<none>' }
                        [void]$sb.AppendLine(("   {0,-30} state={1,-7} phase={2} json={3}" -f $t.Address, $t.State, $t.Phase, $js))
                    }
                    if ($script:Session.MissingTargets -and $script:Session.MissingTargets.Count -gt 0) {
                        [void]$sb.AppendLine("Missing JSONs:")
                        foreach ($m in $script:Session.MissingTargets) {
                            [void]$sb.AppendLine("   $($m.Address) ($($m.Kind)) - $($m.Reason)")
                        }
                    }
                    [void]$sb.AppendLine("")
                    [void]$sb.AppendLine("----- Session Log -----")
                    foreach ($ln in $script:Session.LogTail) { [void]$sb.AppendLine($ln) }
                    if ($dir -and (Test-Path $dir)) {
                        $gen = Join-Path $dir 'gen_report.log'
                        if (Test-Path $gen) {
                            [void]$sb.AppendLine("")
                            [void]$sb.AppendLine("----- gen_report.log -----")
                            [void]$sb.AppendLine([System.IO.File]::ReadAllText($gen, [System.Text.Encoding]::UTF8))
                        }
                        Get-ChildItem $dir -Filter 'server-*.log' -EA 0 | Sort-Object Name | ForEach-Object {
                            [void]$sb.AppendLine("")
                            [void]$sb.AppendLine("----- $($_.Name) -----")
                            try { [void]$sb.AppendLine([System.IO.File]::ReadAllText($_.FullName, [System.Text.Encoding]::UTF8)) } catch {}
                        }
                    }
                    Send-Json -Response $resp -Data @{ ok = $true; content = $sb.ToString() }
                    break
                }
                '^GET /api/download-zip$' {
                    $zip = $script:Session.ReportZipPath
                    if (-not $zip -or -not (Test-Path $zip)) {
                        Send-Response -Response $resp -Body 'Zip not available yet.' -StatusCode 404 -ContentType 'text/plain'
                        break
                    }
                    try {
                        $bytes = [System.IO.File]::ReadAllBytes($zip)
                        $fname = [System.IO.Path]::GetFileName($zip)
                        $resp.StatusCode = 200
                        $resp.ContentType = 'application/zip'
                        $resp.Headers.Add('Content-Disposition', "attachment; filename=""$fname""")
                        $resp.ContentLength64 = $bytes.Length
                        $resp.OutputStream.Write($bytes, 0, $bytes.Length)
                        $resp.OutputStream.Close()
                    } catch {
                        try { Send-Response -Response $resp -Body "Zip stream error: $($_.Exception.Message)" -StatusCode 500 -ContentType 'text/plain' } catch {}
                    }
                    break
                }
                '^GET /$|^GET /index\.html$' {
                    Send-Response -Response $resp -Body $script:HtmlUI
                    break
                }
                default {
                    Send-Response -Response $resp -Body '404' -StatusCode 404 -ContentType 'text/plain'
                }
            }
        } catch {
            try { Send-Response -Response $resp -Body "Server error: $($_.Exception.Message)" -StatusCode 500 -ContentType 'text/plain' } catch { }
        }
    }
} finally {
    if ($listener) { $listener.Stop(); $listener.Close() }
}
