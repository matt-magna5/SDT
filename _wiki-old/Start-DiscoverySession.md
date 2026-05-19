# Start-DiscoverySession

**File:** `Start-DiscoverySession.ps1`  
**Version:** 2.0  
**Role:** Interactive session orchestrator — run this first

---

## Overview

Entry point for the entire discovery workflow. Handles hypervisor enumeration, target discovery, credential management, WinRM safety, and orchestrates per-server discovery runs. Interactive menu-driven — prompts before doing anything destructive.

---

## Phase 1 — Environment Type

Prompts for the hypervisor type(s) in use:

| Option | Type | Notes |
|---|---|---|
| [1] | VMware vCenter | Manages multiple ESXi hosts; connects via vSphere REST API |
| [2] | VMware ESXi | Standalone host; connects via vSphere REST API |
| [3] | Microsoft Hyper-V | Connects via WinRM / Invoke-Command; WMI/DCOM fallback |
| [4] | Bare metal / manual | No hypervisor; manually enter server targets |
| [5] | Multiple | Add multiple hypervisors one at a time |

For each Hyper-V host entered, the script:
- Connects via `Invoke-Command` using domain or local admin credentials
- Runs `Get-VM` to enumerate all VMs with name, state, IP, vCPU, RAM, disk info
- Falls back to WMI/DCOM if WinRM is unavailable (VM list may be partial, IPs blank)
- Pulls a full host inventory (hardware, storage volumes, switch config) via `Get-HyperVInventory`
- **Automatically adds the remote HV host itself as a server discovery target** so it gets a full tab in the report

For each ESXi/vCenter host:
- Authenticates via vSphere REST API (tries v7 API path, falls back to v6)
- Pulls VM list with name, power state, guest OS, IPs via VMware Tools
- Pulls host inventory and datastores
- Saves vSphere connection for optional perf collection later

---

## Phase 2 — Domain Credentials

Prompts for credentials used to connect to and run discovery on each target server. These are separate from the hypervisor credentials (which are asked per-host).

- Format: `DOMAIN\username` for domain-joined environments
- Format: `.\Administrator` for workgroup/local admin

Credentials are stored in `$script:DomainCred` and reused for all targets.

---

## Phase 3 — Suggested Servers (AD + DNS Scan)

Scans the environment for hypervisor/server candidates not already in the target list.

### AD Scan (Primary)
Uses the `ActiveDirectory` module if available:
```powershell
Get-ADComputer -Filter * -Properties Name, IPv4Address, Description, OperatingSystem
```
Filters for names matching any of these patterns:
`*VH*`, `*HV*`, `*ESX*`, `*ESXI*`, `*VCENTER*`, `*HYPERV*`, `*HYP*`, `*VMWARE*`, `*VMW*`, `*NUTANIX*`, `*NTX*`, `*PRISM*`, `*XEN*`, `*PROXMOX*`, `*PVE*`, `*VHOST*`, `*VIRT*`

### ADSI/LDAP Fallback
If the AD module isn't installed, falls back to a direct LDAP query via `DirectoryServices.DirectorySearcher` — no module required, works on any domain-joined machine.

### DNS Sweep Fallback
If AD returns nothing (non-domain or access denied), sweeps DNS with forward lookups for common naming patterns:
`VH`, `HV`, `ESX`, `ESXI`, `VCENTER`, `VC`, `HYPERV`, `VMW`, `NTX`, `PRISM`, `XEN`, `PVE`, `VHOST`

Tries suffixes 1–10 with formats: `VH1`, `VH01`, `VH-1`, `VH_1`

### Result
Presents a numbered list of candidates not already in the target list. User enters comma-separated numbers to add.

---

## Phase 4 — Additional Manual Targets

After hypervisor enumeration, asks if there are standalone servers not on any hypervisor:
- **[1]** Enter hostnames/IPs manually (comma-separated or one per line)
- **[2]** Load from a text file (one hostname/IP per line, `#` for comments)
- **[3]** Skip

---

## Phase 5 — Plan Review & Confirmation

Displays a full table of all targets:

```
  Name                         State        IP                 Source
  DUSODC1                      Running      192.168.100.11     Hyper-V (DUSOVH3)
  DUSOAPPS1                    Running      192.168.100.15     Hyper-V (DUSOVH3)
  DUSOVH5                      Running      192.168.100.152    HyperV Host (DUSOVH5)
  ...
```

User types `GO` to proceed. Anything else cancels.

---

## Phase 6 — Per-Server Discovery

For each target in the plan:

1. **WinRM state check** — queries current WinRM state via WMI/DCOM (without needing WinRM)
2. **If WinRM is OFF** — enables it via `winrm quickconfig` through WMI, runs discovery, then disables it again
3. **If WinRM is already ON** — leaves it exactly as found
4. **Cleanup handler** — `Register-EngineEvent PowerShell.Exiting` fires even on `Ctrl+C` so no server is ever left with WinRM accidentally enabled
5. Calls `Invoke-ServerDiscovery` via `Invoke-Command` for remote targets, or directly for localhost
6. Logs success/failure per server to `session-log.txt`

---

## Phase 7 — Session Output

After all targets complete, writes the session manifest and invokes the report generator:

```json
{
  "client": "ClientName",
  "date": "2026-04-13",
  "session_dir": ".",
  "inventory_file": "DUSOVH3-inventory-2026-04-13.json",
  "servers": [
    { "id": "dusodc1", "file": "DUSODC1-discovery-2026-04-13.json", "name": "DUSODC1", "ip": "192.168.100.11", "in_scope": true },
    ...
  ]
}
```

Automatically runs `gen_report.py` to produce the HTML — uses bundled portable Python (`python\python.exe`) first, falls back to system Python, degrades gracefully with manual command if neither found.

---

## WinRM Safety Details

| Scenario | Behavior |
|---|---|
| WinRM OFF on target | Enable via WMI → run discovery → disable |
| WinRM ON on target | Run discovery → leave ON (don't touch) |
| Ctrl+C during session | Cleanup handler fires → restore all WinRM states |
| WMI unavailable | Skip target, log error, continue with next |
| Workgroup host | Prompt for local admin creds (`.\Administrator`) |

---

## Key Variables

| Variable | Purpose |
|---|---|
| `$script:DomainCred` | Domain credentials for target servers |
| `$script:WinRMRestoreMap` | Tracks original WinRM state per host for cleanup |
| `$script:PendingInventories` | HV/vSphere inventory data to save after discovery |
| `$script:vSphereSources` | vSphere connections for perf collection |
| `$script:SessionVersion` | Current version string (2.0) |
