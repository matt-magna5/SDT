# Nutanix Collector Parser

**File:** `parse_ntnx_collector.py`  
**Purpose:** Parse the XLSX output from the Nutanix Collector tool into VM inventory and sizing data

---

## Background

The [Nutanix Collector](https://collector.nutanix.com) is a free tool that runs against VMware vCenter or ESXi environments and collects:
- VM inventory (name, power state, vCPU, RAM, disk, OS)
- Performance data (CPU/RAM/IOPS at various percentiles over a collection window)
- Host hardware inventory
- Datastore usage

It outputs a ZIP containing an XLSX file with multiple worksheets. This script parses that XLSX and extracts the data needed for cloud migration sizing.

---

## Usage

```bash
python parse_ntnx_collector.py <path-to-collector-zip>
```

Or pass the extracted XLSX directly:
```bash
python parse_ntnx_collector.py ntnxcollector_2026_4_6_11_51_51.xlsx
```

---

## XLSX Worksheets Parsed

| Sheet | Data Extracted |
|---|---|
| `vInfo` | VM name, power state, guest OS, IOPS 95th pct, host name |
| `vCPU` | vCPU count, CPU MHz, peak/average/95th pct utilization % |
| `vMemory` | Provisioned RAM, peak/average/95th pct utilization % |
| `vDisk` | Disk provisioned size per disk per VM |
| `vPartition` | Actual consumed disk per partition (requires VMware Tools) |
| `vHosts` | Host hardware (model, CPU, RAM) |

---

## VMware Tools Dependency

The `vPartition` sheet (actual consumed disk) only has data for VMs with **VMware Tools installed and running**. VMs without Tools show `null` for partition consumption — the script falls back to provisioned disk from `vDisk` for those.

This is a common issue and should be noted in any sizing output.

---

## Output

The parser produces a summary suitable for cloud sizing:

```
=== VM Sizing Summary ===

VM Name              State     vCPU  RAM (GB)  RAM 95%   Disk (GB)  IOPS 95%  OS
QES-RDS-01           On          8     272       51 GB    1,831       657      Windows Server 2019
QES_DC02             On          2       4        2.4 GB    100       520      Windows Server 2019
...

Active Totals: 18 vCPU | 72 GB RAM (95th pct used) | 1,700 GB consumed | 665 peak IOPS
```

---

## Using for Private Cloud Sizing

After parsing, use the VM totals to build an M5 Private Cloud quote:

| Metric | Use |
|---|---|
| vCPU total | `CLOUD-CPU` SKU qty |
| RAM 95th pct total | `CLOUD-RAM` SKU qty (right-sized) |
| Disk consumed total | `CLOUD-DISK` SKU qty (GB) |

Always use **95th percentile RAM** rather than provisioned RAM — provisioned is often 2–10x the actual requirement.

---

## Common Issues

**Collector file won't parse**  
Ensure you're using the XLSX inside the ZIP, not the ZIP itself. The ZIP typically contains `ntnxcollector_DATE/ntnxcollector_DATE.xlsx`.

**All memory utilization columns are null**  
The collector didn't run long enough to capture performance samples, or vCenter didn't have performance statistics enabled. In this case, provisioned RAM is the only option — flag this in your sizing.

**VM missing from output**  
Powered-off VMs have no performance data. The script includes them in inventory but excludes from sizing totals by default.
