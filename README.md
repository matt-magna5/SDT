# Magna5 Server Discovery Tool (SDT)

PowerShell-based server and VM discovery tool for presales cloud migration scoping. Connects to Hyper-V, ESXi, and vCenter environments, runs deep per-server discovery, and generates a branded HTML report for client presentations and internal SE use.

---

## What It Does

### 1. Session Launcher — `Start-DiscoverySession_1.17.ps1`

Interactive session manager. Run this first on any domain-joined jump box or directly on a Hyper-V host.

**Step 1 — Hypervisor Detection**
- Select environment type: vCenter, ESXi (standalone), Hyper-V, bare metal, or any combination
- Connects remotely to each hypervisor and pulls the full VM list
- For Hyper-V: uses `Invoke-Command` (WinRM) with domain or local creds; falls back to WMI/DCOM if WinRM is unavailable
- For ESXi/vCenter: uses the vSphere REST API (v6 and v7 supported)
- Automatically adds each remote HV host itself as a discovery target alongside its VMs

**Step 2 — Suggested Servers (AD + DNS Scan)**
- Scans Active Directory for computer accounts matching hypervisor naming patterns
- Patterns: `VH*`, `HV*`, `ESX*`, `ESXI*`, `VCENTER*`, `VC*`, `HYPERV*`, `VMWARE*`, `VMW*`, `NUTANIX*`, `NTX*`, `PRISM*`, `XEN*`, `PROXMOX*`, `PVE*`, `VHOST*`, `VIRT*`
- Tries the `ActiveDirectory` module first; falls back to ADSI/LDAP (no module required)
- If AD returns nothing, falls back to DNS forward-lookup sweep against common naming patterns
- Presents discovered candidates as a numbered list — user picks which to add to the session

**Step 3 — Target Review & Confirmation**
- Builds a full plan table of all targets: name, IP, power state, OS hint, source
- Shows what will be discovered before running anything
- User types `GO` to proceed — no accidental runs

**Step 4 — Per-Server Discovery**
- Calls `Invoke-ServerDiscovery` on each target
- WinRM management: if WinRM is off, enables it via WMI, runs discovery, disables it again
- Cleanup handler fires on `Ctrl+C` — no server gets left with WinRM enabled
- Runs discovery on powered-on Windows VMs only; Linux VMs are skipped gracefully

**Step 5 — Report Generation**
- Saves per-server JSON to a timestamped session folder
- Saves HV host inventory JSONs (hardware, VM list, storage volumes)
- Writes a manifest JSON pointing to all output files
- Calls `gen_report.py` automatically if Python is available

---

### 2. Server Discovery Agent — `Invoke-ServerDiscovery_1.17.ps1`

Runs on each target server (locally or via WinRM). Read-only. Safe on Windows Server 2008 R2+. Collects:

| Category | Data Collected |
|---|---|
| **System** | OS name, build, version, EOL date/status, install date, uptime, last boot, domain, PowerShell version, timezone, run-as account |
| **Hardware** | CPU model, core count, RAM total/available, manufacturer, model, serial number, BIOS version/date, motherboard |
| **Disks** | Drive letter, label, total GB, free GB, used %, health status |
| **Network** | Adapters, IPs, MACs, DNS servers, default gateway, listening ports |
| **Roles & Features** | All installed Windows roles and features via `Get-WindowsFeature` |
| **Active Directory** | Domain name, forest name, functional levels, DC count, user/computer/OU counts, FSMO roles, stale users (90+ days), stale computers, PDC Emulator, RID Master, Schema Master |
| **DNS** | Installed zones, forwarders |
| **DHCP** | Scopes, address pools, in-use/free counts, start/end ranges |
| **SQL Server** | Instance name, version, edition, EOL date/status, service account, databases |
| **Exchange** | Installed, version, roles |
| **IIS** | Installed, sites, app pools |
| **Hyper-V** | VM list, vCPU/RAM/disk per VM, network adapters, snapshots, uptime, integration services |
| **File Shares** | Share names, paths, permissions (flags Everyone:Full only) |
| **Services** | All running and stopped-auto services, categorized by type |
| **Applications** | Installed software from registry (32-bit and 64-bit) |
| **Event Log** | Recent critical/error events from System and Application logs |
| **Scheduled Tasks** | Non-Microsoft tasks, last run time, status |
| **Printers** | Installed printers and print servers |
| **Flags** | Auto-generated critical/warning flags: EOL OS, SMB1 enabled, stale users, single-point-of-failure roles, SQL EOL, stopped auto-start services, disk usage >85% |

---

### 3. Report Generator — `gen_report.py`

Python script. Takes a session manifest JSON and produces a single self-contained HTML report.

```bash
python gen_report.py <path-to-manifest.json>
```

**Report structure:**

- **Tab per server** — one tab for each discovered server
- **Hyper-V Hosts tab** — aggregated view of all HV hosts with per-host hardware, VM inventory, storage volumes, and cloud sizing inputs (total vCPU/RAM allocated)
- **Views** — Advanced (default, full technical detail) and SBR (executive summary sidebar)

**Per-server tab sections:**

| Section | Contents |
|---|---|
| System Overview | OS, uptime, last boot, domain stats + Security & Protection panel |
| Security & Protection | 🛡️ EDR/XDR, ⚙️ RMM, 🔗 Remote Access, 💾 Backup, 🔑 PAM — auto-detected from installed apps and services |
| Alerts | Auto-generated critical and warning flags |
| Hardware | CPU, RAM, platform, manufacturer, model, serial, BIOS, Dell warranty link |
| Installed Applications | Categorized: Security, Backup, RMM, Database, Web, Other |
| Roles & Features | All installed Windows roles and features |
| Role Configuration | Detailed config for AD, DNS, DHCP, NPS, Hyper-V, IIS, SQL Server, Exchange |
| Disk Storage | Per-drive usage bars |
| Network | Adapters, IPs, listening ports |
| Services | Running services categorized by type; stopped auto-start services flagged |
| File Shares | Share name, path, Everyone:Full flag |

**Security & Protection auto-detection covers:**

- **EDR/XDR**: CrowdStrike Falcon, SentinelOne, Sophos, Huntress, Cylance, Cortex XDR, Carbon Black, Trellix/McAfee, ESET, Kaspersky, Webroot, Bitdefender, Malwarebytes, Cybereason, Elastic Security, Darktrace, Adlumin, Arctic Wolf, Cisco Secure Endpoint, Microsoft Defender for Endpoint, and more
- **RMM**: N-able, NinjaOne, Kaseya VSA, ConnectWise Automate, Datto RMM, Syncro, Atera, Pulseway, Splashtop RMM, ManageEngine, and more
- **Remote Access**: ScreenConnect, N-able Take Control, TeamViewer, AnyDesk, BeyondTrust Remote Support, Dameware, VNC (TightVNC/UltraVNC/RealVNC), Zoho Assist, RustDesk, Chrome Remote Desktop, Parsec, Radmin, and more
- **Backup**: Veeam, Acronis, Commvault, Datto BCDR, Axcient, Azure Backup (MARS), Druva, Rubrik, Cohesity, Barracuda Backup, Windows Server Backup, and more
- **PAM**: CyberArk, BeyondTrust, Delinea/Thycotic, WALLIX, HashiCorp Vault, and more

---

### 4. Supporting Scripts

| Script | Purpose |
|---|---|
| `collect_vsphere_perf.py` | Pulls 95th percentile CPU/RAM utilization history from vCenter/ESXi via vSphere REST API — used for right-sizing cloud migration quotes |
| `parse_ntnx_collector.py` | Parses Nutanix Collector XLSX output for VM inventory and sizing inputs |

---

## Quick Start

**On the target environment (any domain-joined Windows machine with network access to the servers):**

```powershell
# Download and run
iwr https://raw.githubusercontent.com/trophyscar-bit/sdt/main/Start-DiscoverySession_1.17.ps1 -OutFile C:\Temp\Start-DiscoverySession.ps1
iwr https://raw.githubusercontent.com/trophyscar-bit/sdt/main/Invoke-ServerDiscovery_1.17.ps1 -OutFile C:\Temp\Invoke-ServerDiscovery_1.17.ps1
cd C:\Temp
.\Start-DiscoverySession.ps1
```

**Generate report (requires Python 3.8+):**

```bash
python gen_report.py "C:\Temp\Discovery-Session-YYYY-MM-DD-HHMM\manifest.json"
```

---

## Requirements

| Component | Requirement |
|---|---|
| PowerShell | 5.1+ recommended (3.0 minimum) |
| Permissions | Domain admin or local admin on each target server |
| Network | WinRM (port 5985) or WMI/DCOM access to target servers |
| Python | 3.8+ with `openpyxl` for report generation and Nutanix parsing |
| vSphere | REST API access (port 443) for ESXi/vCenter environments |

---

## Output

Each session produces a timestamped folder:

```
Discovery-Session-2026-04-13-1647/
  manifest.json                    # Session index — pass this to gen_report.py
  session-log.txt                  # Run log with timestamps and status
  SERVERNAME-discovery-DATE.json   # Per-server deep discovery data
  HVHOST-inventory-DATE.json       # Per-HV-host inventory (VMs, hardware, storage)
  ClientName-DiscoveryReport-DATE.html  # Final HTML report
```

---

## Version History

| Version | Notes |
|---|---|
| v1.17 | Multi-hypervisor session launcher, vSphere REST API, WinRM safety, Nutanix Collector parser |
| v1.18 | HV host auto-added as discovery targets; AD/DNS suggested servers scan; report fixes |
| v1.19 | Initial report generator release |
| v1.20 | Full security product detection (EDR/RMM/Remote Access/Backup/PAM); System Overview layout; Basic tab removed; AD/SQL/file share rendering fixes |
