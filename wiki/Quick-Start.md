# Quick Start

## Requirements

| Requirement | Details |
|---|---|
| **PowerShell** | 5.1+ recommended, 3.0 minimum |
| **Permissions** | Domain admin (or local admin) on each Windows target server |
| **Where to run** | Any domain-joined Windows machine with network access to the environment |
| **Network — Windows targets** | WinRM port 5985 (auto-enabled by script if off, re-disabled after) |
| **Network — Linux targets** | SSH port 22 (plink used — no agent install on target) |
| **Network — vSphere** | HTTPS port 443 to vCenter/ESXi |
| **Python / plink** | Auto-downloaded on first run — no manual install needed |

---

## What YOU do

### Step 1 — Install to the jump box

Paste this into PowerShell on the target machine:

```powershell
[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; iwr https://raw.githubusercontent.com/matt-magna5/SDT/main/install.ps1 -UseBasicParsing | iex
```

### Step 2 — Launch

```powershell
.\Start-DiscoverySession.ps1
```

### Step 3 — Answer the prompts

1. **Select environment type** — vCenter, ESXi, Hyper-V, bare metal, or a combination
2. **Enter domain credentials** — used for WinRM to all Windows targets
3. **Enter hypervisor IP + credentials** — if vCenter/ESXi/Hyper-V selected
4. **Pick from suggested servers** — AD/DNS scan presents candidates; select which to add
5. **Add any manual targets** — physical boxes, outliers, anything not auto-discovered
6. **Review the full plan table** — all targets listed with name, IP, OS hint, source
7. **Type `GO`** — discovery starts

That's it. Walk away after GO.

---

## What happens automatically (you don't touch this)

| | What the launcher does |
|---|---|
| **Startup** | Checks GitHub for a newer version — prints update one-liner if behind |
| **Python setup** | If `python\python.exe` not found, auto-runs `Get-PortablePython.ps1` — downloads Python 3.12 embeddable + `plink.exe` (~10MB, no installer, no system changes) |
| **Windows discovery** | Calls `Invoke-ServerDiscovery.ps1` on each Windows target via WinRM; shows live progress counter `[3/12]`; enables WinRM via WMI if off, re-disables when done |
| **Linux discovery** | SSH port probe on every target — if port 22 answers, runs lightweight Linux inventory via plink (OS, CPU, RAM, disk, uptime, services, open ports) |
| **vSphere perf** | Runs `collect_vsphere_perf.py` against vCenter — collects 120-day CPU/RAM/IOPS history at 95th percentile for every VM and host |
| **Report** | Calls `gen_report.py` — HTML report auto-generated and saved to the session folder |

---

## Output

All output lands in a timestamped session folder next to the scripts:

```
Discovery-Session-2026-04-13-1647\
  ClientName-manifest-2026-04-13.json       ← session index
  session-log.txt                            ← run log with timestamps
  SERVERNAME-discovery-2026-04-13.json       ← one per Windows server
  HVHOST-inventory-2026-04-13.json           ← one per HV host
  ClientName-DiscoveryReport-2026-04-13.html ← final HTML report
```

Open the HTML in any browser — fully self-contained, no external dependencies.

---

## Files in the folder

**Session launcher bundle — these all work together:**

```
Start-DiscoverySession.ps1    ← YOU run this. Everything else is called by it.
Invoke-ServerDiscovery.ps1    ← Auto: per-server WinRM discovery agent
gen_report.py                 ← Auto: HTML report generator
detection_rules.json          ← Config: edit to add LOB/security detection keywords
Get-PortablePython.ps1        ← Auto: downloads Python 3.12 + plink.exe on first run
collect_vsphere_perf.py       ← Auto: vSphere SOAP perf collector for vCenter/ESXi
```

**Standalone tool — independent of the session launcher:**

```
parse_ntnx_collector.py       ← Run manually if you used the Nutanix Collector GUI
                                 on a client site. Converts the XLSX export to JSON
                                 so gen_report.py can consume it.
                                 Usage: python parse_ntnx_collector.py <xlsx> [outdir]
```

---

## Execution Policy

If you get an execution policy error:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

This only affects the current PowerShell session — safe to run on any machine.

---

## If Python Can't Download

If the jump box has no internet access, discovery still runs and completes normally. At the end, the script prints the manual report command:

```powershell
python "C:\Temp\sdt-3.0\gen_report.py" "C:\Temp\sdt-3.0\Discovery-Session-...\manifest.json"
```

Run this from any machine with Python installed after copying the session folder over.
