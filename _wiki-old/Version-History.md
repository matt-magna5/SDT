# Version History

---

## v2.0 (2026-04-15)

**Self-contained portable Python workflow — no manual report generation step.**

- `Start-DiscoverySession` auto-generates HTML report at session end — no second script needed
- Portable Python bundle: `Get-PortablePython.ps1` downloads official Python 3.12 embeddable package (~10MB, signed) into `SDT\python\`
- Python resolution order: `python\python.exe` (portable) → system Python → graceful fallback with manual command
- Same portable Python detection applied to vSphere perf collector
- `python\` added to `.gitignore` — bundle is set up locally, not committed
- All version strings unified to v2.0 across all components
- Report footer simplified to `SDT v2.0`

---

## Start-DiscoverySession

### v1.18 (2026-04-13)
- Auto-add remote Hyper-V hosts as server discovery targets alongside their VMs
- New Step 3b: AD + DNS scan for suggested hypervisor/server candidates
  - Tries `ActiveDirectory` module, falls back to ADSI/LDAP, then DNS sweep
  - Patterns: VH, HV, ESX, ESXI, VCENTER, HYPERV, VMWARE, NUTANIX, NTX, PRISM, XEN, PROXMOX, PVE, VHOST, VIRT
  - Presents numbered list — user picks which to add
- ScreenConnect display name normalized to "ScreenConnect"
- Version bump: `$script:SessionVersion = '1.18'`

### v1.17 (initial release)
- Multi-hypervisor session launcher (vCenter, ESXi, Hyper-V, bare metal, mixed)
- vSphere REST API v6/v7 support for ESXi and vCenter enumeration
- WinRM safety: enable via WMI → discover → restore; cleanup on Ctrl+C
- Suggested servers: manual entry or file load
- Plan table with explicit GO confirmation before running
- Companion script auto-locate and integrity validation
- Self-heal: if `Invoke-ServerDiscovery` is missing/corrupt, searches companion source paths and auto-copies

---

## Invoke-ServerDiscovery

### v1.10 (current)
- Full data collection: System, Hardware, Disks, Network, Roles, AD, DNS, DHCP, NPS, SQL, Exchange, IIS, Hyper-V, File Shares, Services, Apps, Event Log, Tasks, Printers
- Auto-generated flags for EOL OS, SMB1, SQL EOL, stale users, stopped auto-start services, SPOF roles, disk usage
- PS 3.0 compatibility mode with WMI fallback where CIM cmdlets unavailable
- Windows Server 2008 R2 support

---

## Report Generator (gen_report.py)

### v1.20 (2026-04-13)
- Security & Protection panel in System Overview card (Advanced view)
- Full security product detection: 35+ EDR products, 15+ RMM tools, 20+ remote access tools, 20+ backup agents, 10+ PAM solutions
- Remote Access added as 5th security category (🔗)
- EDR service bucket expanded with all major product service names
- Services card EDR label shows detected product name: "EDR / Endpoint Protection — Sophos"
- Basic tab removed — Advanced is default, SBR only other view
- `scroll-padding-top: 100px` fixes anchor links scrolling behind sticky header
- ScreenConnect normalized to "ScreenConnect" in display

### v1.19 (2026-04-13 — initial release)
- Per-server tabs with Advanced and SBR views
- Hyper-V Hosts tab: all HV inventories auto-loaded, per-host VM tables, storage volumes, cloud sizing inputs
- Anchor nav pills in HV tab with `↑ Top` links per host
- AD rendering fixes: `—` escaping, DN→FQDN conversion for domain display, FSMORoles list used directly
- SQL deep-dive card: instance name, version, edition, EOL, service account
- File shares: removed permissions column, flag Everyone:Full only
- `null` Publisher/Name fields handled gracefully in app categorization
