# vSphere Performance Collection

**File:** `collect_vsphere_perf.py`  
**Purpose:** Pull historical CPU and RAM utilization from vCenter or ESXi via the vSphere REST API

---

## Why This Matters

The Nutanix Collector and this script solve the same problem: **allocated vs. actual utilization**.

A VM with 32 GB RAM allocated may only use 4 GB at 95th percentile. Quoting the full 32 GB means you're presenting a 8x oversized cloud cost. The 95th percentile metric captures peak actual usage over time — the right number to size from.

vCenter stores historical performance data automatically at multiple rollup intervals:

| Interval | Retention |
|---|---|
| Real-time (20s samples) | 1 hour |
| 5-minute rollups | 1 day |
| 30-minute rollups | 1 week |
| 2-hour rollups | 1 month |
| 1-day rollups | 1 year |

---

## Usage

```bash
python collect_vsphere_perf.py <vcenter-host> <username> <password> [--days 30]
```

**Examples:**
```bash
# Pull 30 days of perf data from vCenter
python collect_vsphere_perf.py vcenter.corp.local administrator@vsphere.local MyPass --days 30

# Pull from standalone ESXi host
python collect_vsphere_perf.py 192.168.100.22 root MyRootPass --days 7
```

Output: `vsphere-perf-DATE.json` in the current directory.

---

## Data Collected

For each powered-on VM:

| Metric | Description |
|---|---|
| CPU Usage (%) | Average CPU utilization across all vCPUs |
| CPU Usage (MHz) | Absolute MHz consumed |
| Memory Active (MB) | Active memory — what the guest is actually using |
| Memory Consumed (MB) | Total host memory consumed by the VM |
| Memory Balloon (MB) | Memory reclaimed by VMware balloon driver |
| Disk Read (KB/s) | Average disk read throughput |
| Disk Write (KB/s) | Average disk write throughput |
| IOPS | Combined read + write I/O operations per second |
| Network Rx/Tx (KB/s) | Network throughput per adapter |

From these, the script calculates:
- **95th percentile** for CPU %, Memory Active, IOPS
- **Average** for the same metrics
- **Peak** observed value

---

## Output Format

```json
{
  "vcenter": "vcenter.corp.local",
  "collected": "2026-04-11",
  "days": 30,
  "vms": [
    {
      "name": "QES-RDS-01",
      "vm_id": "vm-123",
      "cpu_pct_95th": 12.4,
      "cpu_mhz_95th": 890,
      "mem_active_gb_95th": 48.2,
      "mem_consumed_gb_95th": 51.0,
      "iops_95th": 657,
      "disk_read_kbps_95th": 4200,
      "disk_write_kbps_95th": 3100
    }
  ]
}
```

---

## When to Use It

Use `collect_vsphere_perf.py` whenever:
- The client environment runs VMware ESXi or vCenter
- You need actual utilization data for accurate cloud sizing
- You see large allocated RAM values on VMs (especially RDS, SQL, application servers)

You need:
- vCenter or ESXi credentials (read-only is sufficient)
- Network access to vCenter/ESXi on port 443
- Python 3.8+ with `requests` library

---

## Hyper-V Equivalent

Hyper-V does **not** store historical performance data by default. If PerfMon data collector sets are running on the HV host, the script could query those logs — but this is rarely configured in client environments. For Hyper-V environments, use the RMM platform (N-able, NinjaOne, etc.) to pull historical utilization if available.
