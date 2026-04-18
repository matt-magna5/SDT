# Magna5 Server Discovery Tool (SDT)

PowerShell-based server and VM discovery tool for presales cloud migration scoping. Connects to Hyper-V, ESXi, and vCenter environments, runs deep per-server discovery, and generates a branded HTML report for client presentations and internal SE use.

**Download and extract in one shot — paste this into PowerShell on the target machine:**

```powershell
iwr https://github.com/matt-magna5/SDT/archive/refs/tags/v3.0.zip -OutFile C:\Temp\sdt.zip; Expand-Archive C:\Temp\sdt.zip C:\Temp -Force; cd C:\Temp\sdt-3.0
```

Then run `.\Start-DiscoverySession.ps1` — everything else is automatic.

---

## What It Does

### 1. Session Launcher — `Start-DiscoverySession.ps1` 

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
- Calls `Invoke-ServerDiscovery` on each target with a live progress counter (`[3/12]`)
- WinRM management: if WinRM is off, enables it via WMI, runs discovery, re-disables it — 15s grace period + 3-attempt retry eliminates the race condition
- Cleanup handler fires on `Ctrl+C` — credential cache purged, WinRM never left enabled
- **Linux / appliance detection**: SSH port probe on each target — if port 22 answers, prompts for SSH credentials and runs a lightweight Linux inventory (OS, RAM, CPU, disk, uptime, open ports, running services). Windows WinRM is skipped for Linux hosts.
- Separate credential flow for Linux: one set of SSH creds re-used across all Linux targets, with `[B]` to restart and change creds mid-session

**Step 5 — Auto-Update Check**
- On startup, checks GitHub for a newer release tag (5s timeout — skipped silently if offline)
- If update available: shows version delta, download link, and one-liner to apply it
- 3-step verification after download: file size check, version string match, SHA256 hash

**Step 6 — Report Generation (automatic)**
- Saves per-server JSON to a timestamped session folder
- Saves HV host inventory JSONs (hardware, VM list, storage volumes)
- Writes a manifest JSON pointing to all output files
- Calls `gen_report.py` automatically — no second step needed
- Uses bundled portable Python (`python\python.exe`) if present; falls back to system Python; warns at session end if Python was never confirmed working

---

### 2. Server Discovery Agent — `Invoke-ServerDiscovery.ps1` 

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

Python script. Runs automatically at the end of each session via the bundled portable Python. Can also be called manually:

```bash
python gen_report.py <path-to-manifest.json>
```

**Report structure:**

- **Tab per server** — one tab for each discovered Windows server (full deep discovery)
- **Linux tab** — lightweight tab for SSH-discovered Linux hosts (OS, CPU, RAM, disk, services, open ports)
- **Hyper-V Hosts tab** — aggregated view of all HV hosts with per-host hardware, VM inventory, storage volumes, and cloud sizing inputs (total vCPU/RAM allocated)
- **SQL tab** — all SQL Server instances across all servers consolidated in one view (databases, sizes, recovery models, compat levels, last backup dates)
- **EOL tab** — cross-server EOL risk summary: Windows OS, SQL Server, Exchange, and any EOL-trackable Microsoft products found in installed apps. Color-coded by days remaining. Pulled live from endoflife.date API with `detection_rules.json` as offline fallback.
- **Private Cloud tab** — auto-generated cloud migration sizing table: merges WinRM and hypervisor data, deduplicates, excludes vCenter/VCSA management infra, totals vCPU/RAM/disk, and includes Commvault backup sizing estimate (1:1 front-end match + 20% growth buffer)
- **Views** — Advanced (default, full technical detail) and SBR (executive summary sidebar)

**Per-server tab sections (in order):**

| Section | Contents |
|---|---|
| Alerts | Auto-generated critical and warning flags |
| System Overview | OS, uptime, last boot, domain stats + Security & Protection panel |
| Hardware | CPU, RAM, platform, manufacturer, model, serial, BIOS, Dell warranty link |
| Installed Applications | Categorized: Security & EDR, Management & RMM, **Line of Business / ERP / CRM**, Browser, Other |
| Roles & Features | All installed Windows roles and features (auto-expanded) |
| Role Configuration | Detailed config for AD, DNS, DHCP, NPS/RADIUS, Hyper-V, IIS, SQL Server, Exchange |
| File Shares | Share name, path, Everyone:Full flag — shown immediately after Role Config |
| Disk Storage | Per-drive usage bars |
| Network | Adapters, IPs, listening ports |
| Services | Running services categorized by type; stopped auto-start services flagged |
| Service Security Flags | Services running from suspicious paths or as custom domain accounts |

**Security & Protection auto-detection covers:**

- **EDR/XDR**: CrowdStrike Falcon, SentinelOne, Sophos, Huntress, Cylance, Cortex XDR, Carbon Black, Trellix/McAfee, ESET, Kaspersky, Webroot, Bitdefender, Malwarebytes, Cybereason, Elastic Security, Darktrace, Adlumin, Arctic Wolf, Cisco Secure Endpoint, Microsoft Defender for Endpoint, and more
- **RMM**: N-able, NinjaOne, Kaseya VSA, ConnectWise Automate, Datto RMM, Syncro, Atera, Pulseway, Splashtop RMM, ManageEngine, and more
- **Remote Access**: ScreenConnect, N-able Take Control, TeamViewer, AnyDesk, BeyondTrust Remote Support, Dameware, VNC (TightVNC/UltraVNC/RealVNC), Zoho Assist, RustDesk, Chrome Remote Desktop, Parsec, Radmin, and more
- **Backup**: Veeam, Acronis, Commvault, Datto BCDR, Axcient, Azure Backup (MARS), Druva, Rubrik, Cohesity, Barracuda Backup, Windows Server Backup, and more
- **PAM**: CyberArk, BeyondTrust, Delinea/Thycotic, WALLIX, HashiCorp Vault, and more

**Line of Business / ERP / CRM auto-detection covers:**

- **ERP**: SAP Business One, SAP HANA, Epicor, Sage (100/200/300/500/X3/Intacct), Oracle NetSuite, Acumatica, SYSPRO, Infor, JobBOSS, E2 Shop, Fishbowl, Global Shop, Aptean, IQMS, ProShop, Spire, and more
- **Accounting**: QuickBooks, Sage 50/Peachtree, MYOB, AccountEdge, Dynamics GP/NAV, Business Central, Dynamics 365
- **CRM**: Redtail CRM, Act!, GoldMine, Vtiger
- **Healthcare**: Medisoft, AdvancedMD, eClinicalWorks, NextGen, Greenway Health, Kareo, PointClickCare
- **Legal**: Tabs3, Time Matters, PCLaw, Clio
- **Construction / Field Service**: Procore, Viewpoint, Jonas Premier, ServiceTitan, simPRO, FieldEdge, WennSoft
- **Property / Real Estate**: Yardi, MRI Software, AppFolio, Rent Manager, Buildium
- **Financial / Wealth**: Orion Advisor, Tamarac, Junxure CRM, Wealthbox, Advent Portfolio Exchange
- **HR / Payroll**: ADP Workforce Now, Paylocity, Paycom, Paychex, UKG Pro, Kronos, Ceridian Dayforce
- **POS / Distribution**: NCR Counterpoint, Lightspeed POS
- **ITSM**: ServiceNow, Freshservice, Cherwell, Ivanti Service Manager

All LOB app keywords are stored in `detection_rules.json` and can be extended without editing Python code.

---

### 4. Supporting Scripts & Config

| File | Purpose |
|---|---|
| `collect_vsphere_perf.py` | Connects to vCenter/ESXi via **vSphere SOAP API** and pulls 120-day CPU/RAM/IOPS/throughput history at 95th percentile for every VM and host. Outputs JSON in the same schema as `parse_ntnx_collector.py`. Usage: `python collect_vsphere_perf.py --vcenter <IP> --user <user> --pass <pass> [--days 120] [--output ./session/]` |
| `parse_ntnx_collector.py` | **Standalone — not part of the session launcher.** Run the Nutanix Collector GUI manually against a client's vCenter, export the XLSX, then run this script to convert it into the same JSON schema as `collect_vsphere_perf.py`. Feed that JSON to `gen_report.py` manually. Usage: `python parse_ntnx_collector.py <xlsx_path> [output_dir]` then `python gen_report.py <manifest.json>` |
| `detection_rules.json` | Central config file for all detection logic. Edit this to add/tune detection without touching Python or PowerShell. Contains: security product keywords (EDR/RMM/PAM/backup/remote access), LOB app keywords (80+ ERP/CRM/vertical apps), service category keywords, OS/SQL/Exchange EOL lifecycle dates, VMware EOL data, and safe/suspicious path lists for the service security scanner. |
| `Get-PortablePython.ps1` | **Called automatically by the launcher** if `python\python.exe` isn't present. Downloads Python 3.12 embeddable + `plink.exe` (~10MB) into `python\`. No installer, no AV exposure. You never need to run this manually. |

---

## Quick Start

### What YOU do (3 things total)

**1. Download to the jump box — paste this into PowerShell:**

```powershell
iwr https://github.com/trophyscar-bit/sdt/archive/refs/tags/v3.0.zip -OutFile C:\Temp\sdt.zip; Expand-Archive C:\Temp\sdt.zip C:\Temp -Force; cd C:\Temp\sdt-3.0
```

**2. Launch:**

```powershell
.\Start-DiscoverySession.ps1
```

**3. Answer the prompts:**
- Select environment type (vCenter / ESXi / Hyper-V / bare metal / combination)
- Enter domain admin credentials (for WinRM to the Windows servers)
- Enter vCenter/hypervisor IP + credentials (if applicable)
- Review the target list and type `GO` to start

That's it. Type `GO` and walk away.

---

### What the launcher does automatically (you don't touch this)

| Step | What happens |
|---|---|
| Startup | Checks GitHub for a newer version — shows update one-liner if behind |
| Python setup | If `python\python.exe` not found, auto-runs `Get-PortablePython.ps1` — downloads Python 3.12 embeddable + `plink.exe` (~10MB, no installer) |
| Target scan | Queries AD + DNS for suggested hypervisors and servers; you pick which to add |
| WinRM discovery | Calls `Invoke-ServerDiscovery.ps1` on each Windows target via WinRM; enables/disables WinRM automatically if off; shows live `[3/12]` progress counter |
| Linux discovery | SSH port probe on every target — if port 22 answers, runs lightweight Linux inventory via `plink.exe` |
| vSphere perf | Runs `collect_vsphere_perf.py` against vCenter to collect 120-day CPU/RAM/IOPS history at 95th percentile for every VM |
| Report | Calls `gen_report.py` automatically — HTML report saved to the session folder |

---

### Files in the folder

**Session launcher bundle — these all work together:**
```
Start-DiscoverySession.ps1    # YOU run this — everything else below is called by it
Invoke-ServerDiscovery.ps1    # Auto: per-server WinRM discovery agent
gen_report.py                 # Auto: HTML report generator
detection_rules.json          # Config: edit to extend LOB/security detection (no code changes needed)
Get-PortablePython.ps1        # Auto: downloads Python 3.12 + plink.exe on first run
collect_vsphere_perf.py       # Auto: vSphere SOAP perf collector (called for vCenter/ESXi targets)
```

**Standalone tools — independent of the session launcher:**
```
parse_ntnx_collector.py       # Run manually after using the Nutanix Collector GUI on a client site.
                              # Takes the XLSX export and converts it to the same JSON format
                              # as collect_vsphere_perf.py so gen_report.py can consume it.
                              # Usage: python parse_ntnx_collector.py <xlsx_path> [output_dir]
```

---

## Requirements

| Component | Requirement |
|---|---|
| **PowerShell** | 5.1+ recommended (3.0 minimum) — run from jump box or any domain-joined Windows machine |
| **Permissions** | Domain admin (or local admin) on each Windows target server |
| **Network — Windows** | WinRM port 5985 reachable on targets (script enables it via WMI if off, then re-disables) |
| **Network — Linux** | SSH port 22 reachable on Linux targets (plink used — no agent install on target) |
| **Network — vSphere** | HTTPS port 443 to vCenter/ESXi for session API + SOAP perf collection |
| **Python** | Auto-downloaded via `Get-PortablePython.ps1` on first run — no manual install |
| **plink.exe** | Auto-downloaded alongside Python — required only if Linux/appliance targets present |

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
| v1.18 | HV host auto-added as discovery targets; AD/DNS suggested servers scan; report fixes; async ping sweep; B-Err silent error swallow fix |
| v1.19 | Initial report generator release |
| v1.20 | Full security product detection (EDR/RMM/Remote Access/Backup/PAM); System Overview layout; Basic tab removed; AD/SQL/file share rendering fixes |
| **v2.0** | **Self-contained workflow — portable Python bundle eliminates manual report generation step. `Get-PortablePython.ps1` downloads official signed Python 3.12 embeddable (~10MB). HTML report auto-generates at session end. No system Python install required.** |
| **v2.1** | **Reliability hardening + report improvements.** PS fixes: TLS 1.2 forced at startup; Hyper-V WMI namespace auto-detects v2/v1 (Server 2008 R2 compat); WinRM post-enable race condition fixed (15s sleep + 3-attempt retry); NPS netsh parser rewritten to produce structured objects; SQL connect error detail captured; EventLog wrapped in 15s timeout job; Win32_Share fallback maps to same schema as Get-SmbShare; credential cache cleanup in `finally`/`trap`; WinRM restore in `finally`. vSphere fixes: PS 5.1 self-signed cert bypass confirmed; API version (v6/v7) locked to responding endpoint; VMware Tools missing → IP shows `(VMware Tools not running)`; ping timeout 1500ms→3000ms. Report fixes: LOB/ERP/CRM app auto-detection (80+ products, extensible via `detection_rules.json`); Service Anomalies replaced with Service Security Flags (suspicious paths + custom domain accounts only); File Shares card moved up after Role Config; Roles & Features auto-expanded; NPS section now shows RADIUS clients table and policy list. |
| **v2.2** | Step 3B opt-in prompt; parallel async DNS sweep for suggested server scan. |
| **v2.3** | **Linux SSH discovery** — SSH port probe on each target; if port 22 answers, lightweight Linux inventory runs (OS, RAM, CPU, disk, uptime, services, open ports). Linux VMs get their own report tab. |
| **v2.4** | Disk Storage table gains Used (GB) column; Used % and utilization bar merged into single column. |
| **v2.5** | **Auto-update check** at startup — compares local version against GitHub release tag, shows update one-liner if behind. Smart stall detection on all Python/portable downloads (no-bytes 10s → skip; 100KB+ 5s → done). `[B]` back-to-start with saved creds during Linux discovery. SSH credential flow separated from Windows credential flow. |
| **v2.6** | 5-method download chain (BITS, HttpClient chunked stream, .NET WebClient, certutil, curl) with per-method buddy spinner. |
| **v2.7–2.8** | Download stall timeouts tightened; pre-commit hook strips all Unicode box-drawing/em-dash chars from PS1 files to prevent PS5 parse errors. |
| **v2.9** | All detection rules extracted to `detection_rules.json` — security products, LOB apps, service categories, safe/suspicious paths all config-driven. Scripts renamed to remove version suffix. 3-step update verification: file size + version string + SHA256 hash. |
| **v3.0** | **EOL tab** — cross-server EOL risk summary for OS, SQL, Exchange, and installed MS products; live lookup via endoflife.date API. **Private Cloud tab** — auto-generated migration sizing: merges WinRM + hypervisor data, deduplicates, excludes vCenter management infra, totals vCPU/RAM/disk, includes Commvault sizing (1:1 front-end + 20% growth). **Progress counter** on discovery (`[3/12] SERVERNAME…`). **120-day perf default** for `collect_vsphere_perf.py`. Live VMware ESXi/vCenter EOL lookups via endoflife.date. |
