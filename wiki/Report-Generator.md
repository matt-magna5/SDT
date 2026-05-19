# Report Generator

**File:** `gen_report.py`  
**Version:** 2.0  
**Role:** Converts session JSON output into a self-contained HTML report

---

## Usage

As of v2.0, `gen_report.py` runs automatically at the end of every discovery session via the bundled portable Python. No manual step needed.

To run manually:

```bash
python gen_report.py <path-to-manifest.json>
```

**Requirements:** Python 3.8+ (bundled via `Get-PortablePython.ps1` — no system install required)

Output is written to the session folder as `ClientName-DiscoveryReport-DATE.html`.

---

## Manifest Input

The manifest JSON tells the generator where to find all the data:

```json
{
  "client": "Precision Corr",
  "client_full": "Precision Corr FL",
  "date": "2026-04-13",
  "session_dir": ".",
  "output_dir": ".",
  "inventory_file": "DUSOVH3-inventory-2026-04-13.json",
  "logo_file": "",
  "servers": [
    { "id": "dusodc1", "file": "DUSODC1-discovery-2026-04-13.json", "name": "DUSODC1", "ip": "192.168.100.11", "in_scope": true }
  ]
}
```

- `inventory_file` — primary HV host inventory (used as seed for Hyper-V tab)
- All `*-inventory-*.json` files in the session directory are auto-loaded for the HV tab
- Servers with `"in_scope": false` are shown with a hollow dot indicator in the tab nav

---

## Report Structure

### Views

The report has two views toggled by buttons in the sticky top bar:

| View | Contents |
|---|---|
| **Advanced** (default) | Full technical detail — all cards visible |
| **SBR** | Executive summary sidebar — key metrics, flags, security status only |

CSS classes control visibility:
- `.sbr-only` — visible in SBR view only
- `.hide-sbr` — hidden in SBR view, visible in Advanced

---

### Tab Navigation

One tab per server, plus a **Hyper-V Hosts** tab:

```
[DUSODC1] [DUSOAPPS1] [DUSOAPPS2] ... [Hyper-V Hosts]
```

- Tabs with critical flags get a red top border
- Tabs with warnings get a yellow top border
- Out-of-scope servers show a `◦` indicator
- Clicking a tab shows that server's content; all others hide

---

### Per-Server Tab — Advanced View

Cards displayed top to bottom:

#### 1. SBR Sidebar (hidden in Advanced, shown in SBR)
- Server name, role label, status badge (CRITICAL / ATTENTION / HEALTHY)
- Left column: OS & System, Active Directory (DCs only), Storage, Security & Protection
- Right column: Security flags, Installed Roles, Network Shares

#### 2. Quick Navigation
- Links to each section within the tab
- Role configuration deep-links (AD, DNS, DHCP, Hyper-V, SQL, etc.)

#### 3. System Overview
Two-column layout:
- **Left:** OS stat boxes (OS, Uptime, Last Boot, Domain) + detail lines (full OS name, EOL date/status, install date, PowerShell version, run-as account, collection timestamp)
- **Right:** Security & Protection panel — see [[Security Detection]]

#### 4. Alerts
Auto-generated flags from the discovery data. Critical flags (red) and warnings (yellow). If no flags: green "No critical alerts" pill.

#### 5. Hardware
- Stat boxes: CPU Cores, RAM Total, RAM Available, Platform (VM/Physical)
- CPU model string
- Hardware identity table: Manufacturer, Model, Serial Number (with Dell warranty link), BIOS version/date, Board product

#### 6. Installed Applications
Categorized by type:
- Security (EDR, AV, firewall agents)
- Backup agents
- RMM agents
- Database (SQL, Oracle, MySQL)
- Web (IIS, Apache, nginx components)
- Other (everything else, collapsed if >10)

Microsoft system components filtered out to reduce noise.

#### 7. Roles & Features
All installed Windows roles and features as badges. Clickable badges jump to the Role Configuration section.

#### 8. Role Configuration
Per-role detail sections (only shown if role is installed):
- **Active Directory** — stat boxes (Users, Computers, OUs, FSMO count), FSMO roles as pills, domain/forest details table (functional levels, PDC, RID, Schema Master), stale user count
- **DNS** — zone count, forwarders, zone table (name, type, dynamic update)
- **DHCP** — scope table (ID, name, state, in-use, free, start/end range)
- **NPS / RADIUS** — policy count, RADIUS client count
- **Hyper-V** — VM table within the server tab (name, state, vCPU, RAM, uptime, snapshots, IPs)
- **IIS** — site table (name, bindings, state, path), app pool table
- **SQL Server** — instance name, version, edition, EOL date/status pill, service account, database list
- **Exchange** — version, server roles

#### 9. Disk Storage
Per-drive rows with usage bar, free GB, total GB, label. Drives >85% flagged red, >70% flagged yellow.

#### 10. Network
- Adapter table: name, IPs, MACs, gateway, DNS servers
- Listening ports table (collapsed by default): port, protocol, process

#### 11. Services
- Stopped auto-start services flagged at top
- Running services grouped by category: EDR/Endpoint Protection (shows detected product name), PAM, RMM, Virtualization, Windows Core, Print, Other
- "Other" group collapsed if >10 services

#### 12. File Shares
- Share name and path
- Flag column: red "⚠️ Everyone: Full" if any permission grants Everyone full access; green checkmark otherwise
- Admin shares (`$`) excluded

---

### Hyper-V Hosts Tab

Aggregated view of all Hyper-V hosts discovered in the session.

**Auto-loading:** All `*-inventory-*.json` files in the session directory with `"_type": "HyperVInventory"` and file size >100 bytes are automatically loaded — no manifest change needed.

**Tab contents:**
- Anchor nav pills at top — one per host, click to jump
- **Cloud Sizing Inputs** banner — total vCPU, RAM, VMs across all hosts (allocated figures, with right-size warning)
- Per-host section (with `↑ Top` link at bottom):
  - Host header: name, manufacturer, model, CPU model, core/logical count, total RAM
  - VM table: Name, State, vCPU, RAM, IP, Disk (size/used), Snapshots
  - Host Storage Volumes: drive letter, label, total GB, free GB, usage bar

---

## Flags Auto-Generation

The report re-evaluates all collected data and generates flags for display:

| Flag | Trigger |
|---|---|
| Windows Server EOL | EOL date < today |
| SMB 1.0 Enabled | `FS-SMB1` in installed features |
| SQL Server EOL/Near-EOL | SQL EOL date within 180 days or past |
| Domain/Forest FL upgrade | Functional level < Windows 2016 |
| Stale User Accounts | StaleUsers list not empty |
| Single Point of Failure | AD + DNS + DHCP all on same server |
| Disk Critical | Any drive >85% used |

---

## Report Versioning

Footer shows the SDT version:
```
Generated 2026-04-13 ET · Magna5 Solutions Engineering · SDT v2.0
```
