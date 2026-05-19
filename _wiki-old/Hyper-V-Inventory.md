# Hyper-V Inventory

---

## What Gets Collected

When `Start-DiscoverySession` connects to a Hyper-V host, it runs `Get-HyperVInventory` which collects full host-level data — separate from the per-server deep discovery that runs on each VM.

### Host Summary

| Field | Source |
|---|---|
| Manufacturer | `Win32_ComputerSystem.Manufacturer` |
| Model | `Win32_ComputerSystem.Model` |
| CPU Model | `Win32_Processor.Name` |
| CPU Cores | `Win32_Processor.NumberOfCores` |
| CPU Logical | `Win32_Processor.NumberOfLogicalProcessors` |
| Total RAM (GB) | `Win32_ComputerSystem.TotalPhysicalMemory` |
| Storage Volumes | `Win32_Volume` — drive letter, label, total GB, free GB, used % |

### VM List

For each VM on the host via `Get-VM`:

| Field | Source |
|---|---|
| Name | `VM.Name` |
| State | `VM.State` (Running, Off, Saved, Paused) |
| Generation | `VM.Generation` (1 or 2) |
| Hyper-V Version | `VM.Version` |
| vCPU | `VM.ProcessorCount` |
| RAM Assigned (GB) | `VM.MemoryAssigned / 1GB` |
| RAM Min/Max | From memory settings (dynamic memory) |
| Dynamic Memory | `VM.DynamicMemoryEnabled` |
| Uptime Hours | `VM.Uptime.TotalHours` |
| Snapshots | `(Get-VMSnapshot $vm).Count` |
| Auto Start | `VM.AutomaticStartAction` |
| IPs | `(Get-VMNetworkAdapter $vm).IPAddresses` |
| Disks | `Get-VMHardDiskDrive` → `Get-VHD` for SizeGB and UsedGB |
| Network Adapters | Switch name, MAC address, IPs |
| Integration Services | `(Get-VMIntegrationService $vm).Name` joined |

### Inventory JSON Format

```json
{
  "_type": "HyperVInventory",
  "HVHost": "DUSOVH3",
  "HostSummary": {
    "Manufacturer": "Dell Inc.",
    "Model": "PowerEdge R630",
    "CPUModel": "Intel Xeon E5-2650L v4",
    "CPUCores": 14,
    "CPULogical": 28,
    "TotalRAMgb": 256.0,
    "Volumes": [
      { "Drive": "C:", "Label": "OS", "TotalGB": 371.4, "FreeGB": 101.8, "UsedPct": 72.6 }
    ]
  },
  "VMs": [
    {
      "Name": "DUSODC1",
      "State": "Running",
      "Generation": 1,
      "vCPU": 6,
      "RAMgb": 16,
      "DynamicMemory": false,
      "UptimeHours": 548.4,
      "Snapshots": 0,
      "IPs": "192.168.100.11",
      "Disks": [
        { "Path": "E:\\VHD\\DUSODC1\\DUSODC1.vhdx", "SizeGB": 127, "UsedGB": 125.6, "VHDType": "Dynamic" }
      ],
      "NetworkAdapters": { "SwitchName": "10gbps-team", "MacAddress": "00155D64150C", "IPs": "192.168.100.11" }
    }
  ]
}
```

---

## Hyper-V Tab in the Report

The report auto-loads **all** `*-inventory-*.json` files in the session directory that have `"_type": "HyperVInventory"` and are non-empty (>100 bytes). No manifest change needed — just run discovery against additional HV hosts and the tab expands automatically.

### Tab Layout

```
[DUSOVH3] [DUSOVH5] [KPVH1]          ← anchor nav pills

┌─ Cloud Sizing Inputs ────────────────────────────────┐
│  3 Hosts   12 VMs   140 vCPU   456 GB RAM            │
│  (allocated figures — right-size before quoting)     │
└──────────────────────────────────────────────────────┘

DUSOVH3                                                  ← id="hv-host-dusovh3"
Dell Inc. PowerEdge R630 · Intel Xeon E5-2650L v4 · 14 cores / 28 logical · 256 GB RAM
                                          4 VMs · 4 running · 48 vCPU · 112 GB RAM

┌──────────────┬───────┬──────┬───────┬─────────────────┬──────────────────┬───────┐
│ VM Name      │ State │ vCPU │ RAM   │ IP              │ Disk (Size/Used) │ Snaps │
├──────────────┼───────┼──────┼───────┼─────────────────┼──────────────────┼───────┤
│ DUSOAPPS1    │ ●     │  8   │ 16 GB │ 192.168.100.15  │ 127 GB / 120 GB  │   0   │
│ ...          │       │      │       │                 │                  │       │

HOST STORAGE VOLUMES
Drive  Label          Total    Free      Usage
C:     OS             371 GB   101 GB    [████████░░] 73%
D:     Data           2979 GB  300 GB    [█████████░] 90% ⚠

                                                          ↑ Top
```

---

## Why RAM Shows as Allocated, Not Used

Hyper-V does not store historical memory utilization. The `RAMgb` field in the inventory is the **assigned/allocated** RAM at the moment the inventory ran — not actual usage.

To get real utilization data you would need:
- Windows Performance Monitor data collector sets running on the HV host (rarely configured)
- RMM/monitoring platform historical data (e.g., N-able, NinjaOne)
- A live `Get-Counter` session during the discovery window (captures only the current moment)

**Implication for cloud sizing:** Always flag allocated RAM figures as a ceiling estimate. Right-size based on actual utilization from RMM or monitoring before finalizing a quote.

---

## Missing Hosts in HV Tab

Common reasons a host doesn't appear:

| Reason | Fix |
|---|---|
| Inventory file is 0 bytes | WinRM/WMI connection to HV host failed; check credentials and network |
| Host is ESXi (TTL=64) | ESXi can't be inventoried as Hyper-V; use `collect_vsphere_perf.py` instead |
| Dead/stale DNS entry | Hostname resolves but machine doesn't exist; skip it |
| Inventory file not in session dir | File landed in wrong location; check output path setting |
