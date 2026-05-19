# Magna5 Server Discovery Tool (SDT) — Wiki

The SDT is a self-contained PowerShell + Python toolset for running deep server and VM discovery against client environments. It produces a branded HTML report used for cloud migration scoping, presales assessments, and internal SE documentation.

As of v2.0, the HTML report generates automatically at session end — no manual Python step, no system dependencies. Portable Python is bundled via `Get-PortablePython.ps1`.

---

## Wiki Pages

| Page | Description |
|---|---|
| [[Quick Start]] | Download, run, generate your first report |
| [[Start-DiscoverySession]] | Session launcher — hypervisor detection, AD/DNS scan, target management |
| [[Invoke-ServerDiscovery]] | Per-server discovery agent — all data collected |
| [[Report-Generator]] | gen_report.py — HTML report structure, tabs, views |
| [[Security-Detection]] | Full product lists for EDR, RMM, Remote Access, Backup, PAM |
| [[Hyper-V-Inventory]] | HV host inventory collection and HV tab in the report |
| [[vSphere-Performance]] | collect_vsphere_perf.py — 95th percentile utilization from vCenter |
| [[Nutanix-Collector]] | parse_ntnx_collector.py — parsing Nutanix Collector XLSX output |
| [[Manifest-Format]] | Session manifest JSON schema |
| [[Troubleshooting]] | WinRM errors, AD access, Python issues, common failures |
| [[Version-History]] | Changelog across all script versions |

---

## Architecture Overview

```
Get-PortablePython.ps1              ← Run once (downloads Python bundle)

Start-DiscoverySession.ps1          ← Run this (interactive)
    │
    ├── Connects to hypervisors (HV / ESXi / vCenter)
    ├── Scans AD + DNS for suggested targets
    ├── Adds HV hosts + VMs to target list
    ├── Prompts user to confirm
    │
    └── Calls Invoke-ServerDiscovery.ps1 on each target
            │
            └── Outputs JSON per server → session folder
                    │
                    └── gen_report.py runs automatically → HTML report
```

## Quick Setup (v2.0)

Copy the SDT folder to the target machine, then:

```powershell
# One-time setup
.\Get-PortablePython.ps1

# Run discovery (report auto-generates at the end)
.\Start-DiscoverySession.ps1
```
