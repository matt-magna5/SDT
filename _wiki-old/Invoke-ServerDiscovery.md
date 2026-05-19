# Invoke-ServerDiscovery

**File:** `Invoke-ServerDiscovery_1.17.ps1`  
**Version:** 1.10 (internal script version)  
**Role:** Per-server deep discovery agent

---

## Overview

Runs on a single target server — locally or remotely via WinRM. Completely read-only. No changes made to the target. Compatible with Windows Server 2008 R2 and later.

Output is a structured JSON file saved to the session directory. The filename format is `SERVERNAME-discovery-DATE.json`.

---

## Invocation

**Local:**
```powershell
.\Invoke-ServerDiscovery.ps1
```

**Remote (called by Start-DiscoverySession):**
```powershell
.\Invoke-ServerDiscovery.ps1 -ComputerName SRV-DC01 -Credential $cred -OutputPath "C:\Temp\SDT\Session-2026-04-13"
```

**Parameters:**

| Parameter | Default | Description |
|---|---|---|
| `ComputerName` | Local machine | Target server hostname or IP |
| `OutputPath` | Script directory | Where to save the JSON output |
| `Credential` | Prompted if remote | PSCredential for WinRM |

---

## Data Collected

### System
| Field | Source |
|---|---|
| OS Name | `Win32_OperatingSystem.Caption` |
| OS Build | `Win32_OperatingSystem.BuildNumber` |
| OS Version | `Win32_OperatingSystem.Version` |
| EOL Date | Hardcoded lookup table by OS build |
| EOL Status | Derived from EOL date vs. current date |
| Install Date | `Win32_OperatingSystem.InstallDate` |
| Last Boot | `Win32_OperatingSystem.LastBootUpTime` |
| Uptime (Days) | Calculated from last boot |
| Domain | `Win32_ComputerSystem.Domain` |
| Hostname | `$env:COMPUTERNAME` |
| PowerShell Version | `$PSVersionTable.PSVersion` |
| Timezone | `[TimeZoneInfo]::Local.DisplayName` |
| Run As User | `[Security.Principal.WindowsIdentity]::GetCurrent().Name` |

### Hardware
| Field | Source |
|---|---|
| CPU Model | `Win32_Processor.Name` |
| CPU Cores | `Win32_Processor.NumberOfCores` |
| CPU Logical | `Win32_Processor.NumberOfLogicalProcessors` |
| RAM Total GB | `Win32_ComputerSystem.TotalPhysicalMemory` |
| RAM Available GB | `Win32_OperatingSystem.FreePhysicalMemory` |
| Manufacturer | `Win32_ComputerSystem.Manufacturer` |
| Model | `Win32_ComputerSystem.Model` |
| Serial Number | `Win32_BIOS.SerialNumber` |
| BIOS Version | `Win32_BIOS.SMBIOSBIOSVersion` |
| BIOS Date | `Win32_BIOS.ReleaseDate` |
| Board Product | `Win32_BaseBoard.Product` |
| VM Platform | Detected from manufacturer/model (VMware, Hyper-V, Nutanix, AWS, Azure, GCP) |

### Disks
- Drive letter, volume label
- Total GB, Free GB, Used %
- Source: `Win32_LogicalDisk` (local fixed drives only)

### Network
- All adapters with IP addresses, subnet masks, MAC addresses
- DNS servers, default gateway
- Adapter description
- All active TCP listening ports via `Get-NetTCPConnection` or `netstat`

### Roles & Features
- Full list of installed Windows Server roles and features
- Source: `Get-WindowsFeature` (Server 2008 R2+)
- Captures both `Installed` state and `Display Name`

### Active Directory (DCs only)
| Field | Source |
|---|---|
| Domain Name | `Get-ADDomain.Name` |
| Forest Name | `Get-ADForest.Name` |
| Domain Functional Level | `Get-ADDomain.DomainMode` |
| Forest Functional Level | `Get-ADForest.ForestMode` |
| PDC Emulator | `Get-ADDomain.PDCEmulator` |
| RID Master | `Get-ADDomain.RIDMaster` |
| Infrastructure Master | `Get-ADDomain.InfrastructureMaster` |
| Schema Master | `Get-ADForest.SchemaMaster` |
| Domain Naming Master | `Get-ADForest.DomainNamingMaster` |
| FSMO Roles (this DC) | `netdom query fsmo` parsed |
| DC Count | `(Get-ADDomainController -Filter *).Count` |
| User Count | `(Get-ADUser -Filter *).Count` |
| Computer Count | `(Get-ADComputer -Filter *).Count` |
| OU Count | `(Get-ADOrganizationalUnit -Filter *).Count` |
| Stale Users | Users with LastLogonDate > 90 days ago |
| Stale Computers | Computers with LastLogonDate > 90 days ago |

### DNS (if DNS role installed)
- All zones (name, type, dynamic update setting)
- Configured forwarders

### DHCP (if DHCP role installed)
- All scopes with: ScopeId, Name, State, StartRange, EndRange, AddressesInUse, AddressesFree
- Source: `Get-DhcpServerv4Scope`

### NPS / RADIUS (if NPS role installed)
- Network policies count
- RADIUS clients count

### SQL Server (if installed)
- Instance name
- SQL Server version string
- Edition (e.g. SQL Server 2016, SQL Server 2019)
- EOL date and status
- Service account
- Database list (names, sizes where accessible)
- Source: Registry + SQL WMI provider

### Exchange (if installed)
- Installed, version, server roles
- Source: Exchange management snap-in or registry

### IIS (if installed)
- Sites (name, bindings, state, physical path)
- Application pools
- Source: `Microsoft.Web.Administration`

### Hyper-V (if Hyper-V role installed)
- VM list with: Name, State, Generation, Version, vCPU, RAM (assigned/min/max), Dynamic Memory, Uptime Hours, Snapshot count, AutoStart action
- VM IPs (via Get-VMNetworkAdapter)
- VM disk files (path, controller type, SizeGB, UsedGB, VHD type)
- Network adapters (switch name, MAC)
- Integration services state
- Source: `Get-VM`, `Get-VHD`, `Get-VMNetworkAdapter`

### File Shares
- All non-admin shares (excludes `$` hidden shares)
- Share name, path, description
- Permissions: AccountName, AccessRight, AccessControlType
- Source: `Get-SmbShare`, `Get-SmbShareAccess`

### Services
- All Windows services: DisplayName, Name, State, StartMode, StartName (account)
- Stopped services that are set to auto-start (flagged)
- Source: `Get-Service` + `Win32_Service`

### Installed Applications
- All installed software from both 32-bit and 64-bit registry hives:
  - `HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*`
  - `HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*`
- Fields: Name, Publisher, Version, InstallDate, InstallLocation

### Event Log
- Recent critical and error events from System and Application logs
- Last 50 events, last 7 days
- Fields: TimeCreated, Id, LevelDisplayName, ProviderName, Message (truncated)

### Scheduled Tasks
- All non-Microsoft scheduled tasks
- Fields: TaskName, TaskPath, State, LastRunTime, LastTaskResult, NextRunTime
- Source: `Get-ScheduledTask`

### Printers
- Installed printers and print servers
- Source: `Get-Printer`

---

## Auto-Generated Flags

The script automatically generates flags based on collected data:

| Flag | Severity | Condition |
|---|---|---|
| Windows Server EOL | Critical | OS EOL date has passed or is within 180 days |
| SMB 1.0 Enabled | Critical | `FS-SMB1` feature installed and enabled |
| SQL Server EOL | Critical | SQL EOL date has passed |
| Stale User Accounts | Warning | Any users inactive 90+ days |
| Stopped Auto-Start Services | Warning | Services set to Automatic that are Stopped |
| Single Point of Failure | Warning | AD + DNS + DHCP all on one VM |
| Disk Usage Critical | Warning | Any drive >85% full |
| Domain Functional Level | Warning | Domain or Forest FL below Windows 2016 |

---

## JSON Output Structure

```json
{
  "System": { "OSName": "...", "Domain": "...", ... },
  "Hardware": { "CPUModel": "...", "RAMTotalGB": 32, ... },
  "Disks": [ { "Drive": "C:", "TotalGB": 100, "FreeGB": 45, "UsedPct": 55 } ],
  "Network": { "Adapters": [...], "ListeningPorts": [...] },
  "Roles": { "InstalledRoles": [...], "InstalledFeatures": [...] },
  "AD": [ { "DomainName": "corp.local", "UserCount": 250, ... } ],
  "DNS": { "Installed": true, "Zones": [...] },
  "DHCP": { "Installed": true, "Scopes": [...] },
  "SQL": { "Instances": { "InstanceName": "...", "Version": "...", ... } },
  "HyperV": { "VMs": [...] },
  "FileShares": { "Shares": [...] },
  "Services": [...],
  "Apps": [...],
  "EventLog": [...],
  "Tasks": [...],
  "Flags": [...],
  "Meta": { "CollectedAt": "2026-04-13 16:53:28", "ScriptVersion": "1.10" }
}
```

---

## Compatibility Notes

| OS | Support Level |
|---|---|
| Windows Server 2022 | Full |
| Windows Server 2019 | Full |
| Windows Server 2016 | Full |
| Windows Server 2012 R2 | Full |
| Windows Server 2008 R2 | Partial — some CIM cmdlets fall back to WMI |
| Windows 10/11 | Works for workstation discovery |
