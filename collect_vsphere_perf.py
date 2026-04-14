"""
collect_vsphere_perf.py  —  Full vSphere SOAP Collector
Nutanix Collector equivalent: inventory + 90-day perf + storage + licenses.
No PowerCLI, no Nutanix Collector tool required.
Output: vsphere-perf-<date>.json  (same schema as parse_ntnx_collector.py)

What it collects (matches Nutanix Collector XLS sheets):
  vHosts   — host model, service tag, CPU, RAM, hypervisor, NIC count,
              CPU/Mem usage %, IOPS P95, disk throughput P95
  vCluster — aggregate CPU/Mem %, IOPS P95, storage capacity/consumed
  vCPU     — per-VM: avg/peak/median/P95 CPU usage, CPU Readiness P95
  vMemory  — per-VM: avg/peak/median/P95 memory usage
  vInfo    — per-VM: power state, guest OS, tool status, IOPS P95
  vmList   — per-VM: datastore assignment, capacity (GB), consumed (GB)
  vPart    — per-VM: disk partitions (path, capacity, free) via guest tools
  Datastore— per-datastore: name, type, capacity, free, used %
  vLicense — VMware license keys, edition, used/total, expiry

Usage:
  python collect_vsphere_perf.py --vcenter 10.200.1.12 --user admin@vsphere.local --pass secret
  python collect_vsphere_perf.py --vcenter 10.200.1.12 --user admin@vsphere.local --pass secret --days 90 --output ./session/
"""
import requests, json, sys, os, re, argparse
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone
from statistics import mean, median as stat_median

requests.packages.urllib3.disable_warnings()

SOAP_URL = "https://{host}/sdk/vimService"
NS       = "urn:vim25"
XSI      = "http://www.w3.org/2001/XMLSchema-instance"

# ── SOAP plumbing ─────────────────────────────────────────────────────────────

def soap_req(sess, host, body_xml):
    env = (f'<?xml version="1.0" encoding="UTF-8"?>'
           f'<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"'
           f' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"'
           f' xmlns:vim25="urn:vim25"><soapenv:Body>{body_xml}</soapenv:Body></soapenv:Envelope>')
    r = sess.post(SOAP_URL.format(host=host), data=env.encode(),
                  headers={"Content-Type": "text/xml; charset=utf-8",
                           "SOAPAction": "urn:vim25/6.0"},
                  verify=False, timeout=180)
    if r.status_code not in (200, 500):
        raise RuntimeError(f"HTTP {r.status_code}: {r.text[:400]}")
    root = ET.fromstring(r.content)
    fault = root.find('.//{http://schemas.xmlsoap.org/soap/envelope/}Fault')
    if fault is not None:
        raise RuntimeError(f"SOAP fault {fault.findtext('faultcode','?')}: {fault.findtext('faultstring','?')}")
    return root

def xt(e, *path):
    """Walk element by tag names, return .text or None."""
    cur = e
    for t in path:
        if cur is None: return None
        cur = cur.find(f'{{{NS}}}{t}') or cur.find(t)
    return cur.text if cur is not None else None

def xall(e, tag):
    return e.findall(f'{{{NS}}}{tag}') + e.findall(tag)

# ── Standard full-inventory traversal (rootFolder → everything) ───────────────
_TRAV = '''
<vim25:selectSet xsi:type="vim25:TraversalSpec">
  <vim25:name>tsFolderChild</vim25:name><vim25:type>Folder</vim25:type>
  <vim25:path>childEntity</vim25:path><vim25:skip>false</vim25:skip>
  <vim25:selectSet><vim25:name>tsFolderChild</vim25:name></vim25:selectSet>
  <vim25:selectSet><vim25:name>tsDCvmFolder</vim25:name></vim25:selectSet>
  <vim25:selectSet><vim25:name>tsDChostFolder</vim25:name></vim25:selectSet>
  <vim25:selectSet><vim25:name>tsDCdsFolder</vim25:name></vim25:selectSet>
  <vim25:selectSet><vim25:name>tsClusterRP</vim25:name></vim25:selectSet>
  <vim25:selectSet><vim25:name>tsRP</vim25:name></vim25:selectSet>
</vim25:selectSet>
<vim25:selectSet xsi:type="vim25:TraversalSpec">
  <vim25:name>tsDCvmFolder</vim25:name><vim25:type>Datacenter</vim25:type>
  <vim25:path>vmFolder</vim25:path><vim25:skip>false</vim25:skip>
  <vim25:selectSet><vim25:name>tsFolderChild</vim25:name></vim25:selectSet>
</vim25:selectSet>
<vim25:selectSet xsi:type="vim25:TraversalSpec">
  <vim25:name>tsDChostFolder</vim25:name><vim25:type>Datacenter</vim25:type>
  <vim25:path>hostFolder</vim25:path><vim25:skip>false</vim25:skip>
  <vim25:selectSet><vim25:name>tsFolderChild</vim25:name></vim25:selectSet>
</vim25:selectSet>
<vim25:selectSet xsi:type="vim25:TraversalSpec">
  <vim25:name>tsDCdsFolder</vim25:name><vim25:type>Datacenter</vim25:type>
  <vim25:path>datastoreFolder</vim25:path><vim25:skip>false</vim25:skip>
  <vim25:selectSet><vim25:name>tsFolderChild</vim25:name></vim25:selectSet>
</vim25:selectSet>
<vim25:selectSet xsi:type="vim25:TraversalSpec">
  <vim25:name>tsClusterRP</vim25:name><vim25:type>ComputeResource</vim25:type>
  <vim25:path>resourcePool</vim25:path><vim25:skip>false</vim25:skip>
  <vim25:selectSet><vim25:name>tsRP</vim25:name></vim25:selectSet>
</vim25:selectSet>
<vim25:selectSet xsi:type="vim25:TraversalSpec">
  <vim25:name>tsRP</vim25:name><vim25:type>ResourcePool</vim25:type>
  <vim25:path>vm</vim25:path><vim25:skip>false</vim25:skip>
</vim25:selectSet>'''

def retrieve_all(sess, host, pc_ref, root_folder, obj_type, paths):
    """
    Retrieve named properties for all objects of obj_type under rootFolder.
    Returns list of dicts: {'_moid': ..., '_elem_<prop>': xml_element, prop: text_or_None, ...}
    """
    ps = ''.join(f'<vim25:pathSet>{p}</vim25:pathSet>' for p in paths)
    body = f'''<vim25:RetrievePropertiesEx>
      <vim25:_this type="PropertyCollector">{pc_ref}</vim25:_this>
      <vim25:specSet>
        <vim25:propSpec>
          <vim25:type>{obj_type}</vim25:type>
          <vim25:all>false</vim25:all>
          {ps}
        </vim25:propSpec>
        <vim25:objectSpec>
          <vim25:obj type="Folder">{root_folder}</vim25:obj>
          <vim25:skip>false</vim25:skip>
          {_TRAV}
        </vim25:objectSpec>
      </vim25:specSet>
      <vim25:options><vim25:maxObjects>0</vim25:maxObjects></vim25:options>
    </vim25:RetrievePropertiesEx>'''
    root = soap_req(sess, host, body)
    results = []
    for oc in list(root.iter(f'{{{NS}}}objects')) + list(root.iter('objects')):
        moid_e = oc.find(f'{{{NS}}}obj') or oc.find('obj')
        rec = {'_moid': moid_e.text if moid_e is not None else ''}
        for ps_e in xall(oc, 'propSet'):
            n   = xt(ps_e, 'name') or ''
            v_e = ps_e.find(f'{{{NS}}}val') or ps_e.find('val')
            rec[n]            = v_e.text if v_e is not None else None
            rec[f'_e_{n}']    = v_e           # keep element for complex parsing
        results.append(rec)
    return results

# ── Step 1: Login  ───────────────────────────────────────────────────────────

def login(sess, host, user, pw):
    """Returns (pc_ref, perf_ref, root_folder, lic_mgr_ref)."""
    root = soap_req(sess, host,
        '<vim25:RetrieveServiceContent>'
        '<vim25:_this type="ServiceInstance">ServiceInstance</vim25:_this>'
        '</vim25:RetrieveServiceContent>')
    rv = root.find(f'.//{{{NS}}}returnval') or root.find('.//returnval')

    def _ref(name):
        e = rv.find(f'{{{NS}}}{name}') or rv.find(name) if rv is not None else None
        return e.text if e is not None else None

    sm   = _ref('sessionManager')
    pc   = _ref('propertyCollector')
    perf = _ref('perfManager')
    rf   = _ref('rootFolder')
    lic  = _ref('licenseManager')

    soap_req(sess, host,
        f'<vim25:Login>'
        f'<vim25:_this type="SessionManager">{sm}</vim25:_this>'
        f'<vim25:userName>{user}</vim25:userName>'
        f'<vim25:password>{pw}</vim25:password>'
        f'</vim25:Login>')
    return pc, perf, rf, lic

# ── Step 2: VM inventory ──────────────────────────────────────────────────────

def get_vms(sess, host, pc_ref, root_folder):
    """Return list of VM dicts with config + storage + partition data."""
    rows = retrieve_all(sess, host, pc_ref, root_folder, 'VirtualMachine', [
        'name',
        'runtime.powerState',
        'config.guestFullName',
        'guest.guestFullName',
        'config.hardware.numCPU',
        'config.hardware.memoryMB',
        'guest.toolsStatus',
        'summary.config.vmPathName',        # "[DatastoreName] vm/vm.vmx"
        'summary.storage.committed',         # bytes consumed on disk
        'summary.storage.uncommitted',       # bytes provisioned but not used
        'guest.disk',                        # partition detail (requires Tools running)
    ])

    vms = []
    for r in rows:
        if not r.get('name'):
            continue

        # Extract datastore from vmPathName: "[SAS01] QES-RDS-01/QES-RDS-01.vmx"
        vmx = r.get('summary.config.vmPathName', '') or ''
        ds_match = re.search(r'\[(.+?)\]', vmx)
        datastore = ds_match.group(1) if ds_match else ''

        # Disk sizing (bytes)
        committed   = int(r.get('summary.storage.committed')   or 0)
        uncommitted = int(r.get('summary.storage.uncommitted') or 0)
        cap_gb      = round((committed + uncommitted) / (1024**3), 1)
        cons_gb     = round(committed / (1024**3), 1)

        # Guest disk partitions from guest.disk element
        partitions = []
        gdisk_e = r.get('_e_guest.disk')
        if gdisk_e is not None:
            for gdi in list(gdisk_e.iter(f'{{{NS}}}GuestDiskInfo')) + list(gdisk_e.iter('GuestDiskInfo')):
                disk_path  = gdi.findtext(f'{{{NS}}}diskPath')  or gdi.findtext('diskPath')  or ''
                disk_cap   = int(gdi.findtext(f'{{{NS}}}capacity')  or gdi.findtext('capacity')  or 0)
                disk_free  = int(gdi.findtext(f'{{{NS}}}freeSpace') or gdi.findtext('freeSpace') or 0)
                disk_used  = disk_cap - disk_free
                partitions.append({
                    "Path":        disk_path,
                    "CapacityMiB": round(disk_cap / (1024**2), 1),
                    "ConsumedMiB": round(disk_used / (1024**2), 1),
                })

        ram_mb = int(r.get('config.hardware.memoryMB', 0) or 0)
        guest_os = r.get('guest.guestFullName') or r.get('config.guestFullName') or ''
        is_linux = bool(guest_os and any(k in guest_os.lower() for k in
                   ['linux','ubuntu','centos','rhel','debian','photon','suse','oracle',
                    'rocky','alma','amazon','coreos','fedora']))
        # Also detect Linux by partition paths (/ = Linux mount)
        if not is_linux and partitions:
            is_linux = any(p['Path'].startswith('/') for p in partitions)

        vms.append({
            'Name':       r['name'],
            'MOID':       r['_moid'],
            'PowerState': r.get('runtime.powerState', ''),
            'GuestOS':    guest_os,
            'IsLinux':    is_linux,
            'vCPUs':      int(r.get('config.hardware.numCPU', 0) or 0),
            'RAMmb':      ram_mb,
            'RAMgb':      round(ram_mb / 1024, 1),
            'ToolStatus': r.get('guest.toolsStatus', ''),
            'Datastore':  datastore,
            'DiskCapGB':  cap_gb,
            'DiskConsumedGB': cons_gb,
            'Partitions': partitions,
        })
    return vms

# ── Step 3: Host info + host-level perf ──────────────────────────────────────

def get_host_info(sess, host_addr, pc_ref, root_folder):
    """Return host hardware dict including quickStats CPU/Mem usage."""
    rows = retrieve_all(sess, host_addr, pc_ref, root_folder, 'HostSystem', [
        'name',
        'hardware.systemInfo.model',
        'hardware.systemInfo.vendor',
        'hardware.systemInfo.serialNumber',
        'hardware.cpuInfo.numCpuCores',
        'hardware.cpuInfo.hz',
        'hardware.memorySize',
        'config.product.fullName',
        'summary.hardware.numNics',
        'summary.hardware.cpuModel',
        'summary.quickStats.overallCpuUsage',       # MHz currently used
        'summary.quickStats.overallMemoryUsage',    # MB currently used
        'summary.hardware.cpuMhz',                  # MHz per core
        'summary.config.name',
        'summary.managementServerIp',
    ])
    if not rows:
        return {}
    r = rows[0]
    total_mhz  = (int(r.get('hardware.cpuInfo.numCpuCores', 0) or 0) *
                  int(r.get('summary.hardware.cpuMhz', 0) or 0))
    used_mhz   = int(r.get('summary.quickStats.overallCpuUsage', 0) or 0)
    total_mem_mb = int(r.get('hardware.memorySize', 0) or 0) // (1024**2)
    used_mem_mb  = int(r.get('summary.quickStats.overallMemoryUsage', 0) or 0)

    cpu_pct = round(used_mhz  / total_mhz   * 100, 2) if total_mhz  else 0
    mem_pct = round(used_mem_mb / total_mem_mb * 100, 2) if total_mem_mb else 0

    return {
        "Name":        r.get('name') or r.get('summary.config.name', ''),
        "Model":       r.get('hardware.systemInfo.model', ''),
        "Vendor":      r.get('hardware.systemInfo.vendor', ''),
        "ServiceTag":  r.get('hardware.systemInfo.serialNumber', ''),
        "Cores":       int(r.get('hardware.cpuInfo.numCpuCores', 0) or 0),
        "CPUModel":    r.get('summary.hardware.cpuModel', ''),
        "CPUSpeedMHz": int(int(r.get('hardware.cpuInfo.hz', 0) or 0) / 1_000_000),
        "RAMgb":       round(int(r.get('hardware.memorySize', 0) or 0) / (1024**3), 2),
        "Hypervisor":  r.get('config.product.fullName', ''),
        "NICs":        int(r.get('summary.hardware.numNics', 0) or 0),
        "CPUUsagePct": cpu_pct,
        "MemUsagePct": mem_pct,
        # IOPS_95th filled later from host perf query
        "IOPS_95th":   0,
        "DiskKBps_95th": 0,
    }

# ── Step 4: Datastores ───────────────────────────────────────────────────────

def get_datastores(sess, host, pc_ref, root_folder):
    """Return list of datastore dicts."""
    rows = retrieve_all(sess, host, pc_ref, root_folder, 'Datastore', [
        'name',
        'summary.type',
        'summary.capacity',
        'summary.freeSpace',
        'summary.accessible',
    ])
    ds_list = []
    for r in rows:
        if not r.get('name'):
            continue
        cap  = int(r.get('summary.capacity', 0)  or 0)
        free = int(r.get('summary.freeSpace', 0) or 0)
        used = cap - free
        ds_list.append({
            "Name":       r['name'],
            "Type":       r.get('summary.type', 'VMFS'),
            "CapacityGB": round(cap  / (1024**3), 1),
            "FreeGB":     round(free / (1024**3), 1),
            "ConsumedGB": round(used / (1024**3), 1),
            "UsedPct":    round(used / cap * 100, 1) if cap else 0,
            "CapacityMiB": round(cap  / (1024**2), 1),
            "ConsumedMiB": round(used / (1024**2), 1),
        })
    return ds_list

# ── Step 5: Cluster info ──────────────────────────────────────────────────────

def get_cluster_info(sess, host, pc_ref, root_folder):
    """Return cluster aggregate dict. Falls back to empty if standalone host."""
    rows = retrieve_all(sess, host, pc_ref, root_folder, 'ClusterComputeResource', [
        'name',
        'summary.currentCpuUsage',      # MHz
        'summary.totalCpu',             # MHz total
        'summary.currentMemoryUsage',   # MB
        'summary.totalMemory',          # bytes
        'summary.numEffectiveHosts',
        'summary.numHosts',
    ])
    if not rows:
        return {}
    r = rows[0]
    total_mhz = int(r.get('summary.totalCpu', 0) or 0)
    used_mhz  = int(r.get('summary.currentCpuUsage', 0) or 0)
    total_mem_b = int(r.get('summary.totalMemory', 0) or 0)
    used_mem_mb = int(r.get('summary.currentMemoryUsage', 0) or 0)
    total_mem_mb = total_mem_b // (1024**2)

    return {
        "Name":        r.get('name', ''),
        "CPUUsagePct": round(used_mhz / total_mhz * 100, 2) if total_mhz else 0,
        "MemUsagePct": round(used_mem_mb / total_mem_mb * 100, 2) if total_mem_mb else 0,
        "CapacityMiB": 0,   # filled from datastore sum
        "ConsumedMiB": 0,
        "IOPS_95th":   0,   # filled from perf
        "DiskKBps_95th": 0,
        "NumHosts":    int(r.get('summary.numHosts', 0) or 0),
    }

# ── Step 6: Licenses ──────────────────────────────────────────────────────────

def get_licenses(sess, host, pc_ref, lic_ref):
    """Return list of license dicts from LicenseManager."""
    if not lic_ref:
        return []
    body = f'''<vim25:RetrievePropertiesEx>
      <vim25:_this type="PropertyCollector">{pc_ref}</vim25:_this>
      <vim25:specSet>
        <vim25:propSpec>
          <vim25:type>LicenseManager</vim25:type>
          <vim25:all>false</vim25:all>
          <vim25:pathSet>licenses</vim25:pathSet>
        </vim25:propSpec>
        <vim25:objectSpec>
          <vim25:obj type="LicenseManager">{lic_ref}</vim25:obj>
        </vim25:objectSpec>
      </vim25:specSet>
      <vim25:options/>
    </vim25:RetrievePropertiesEx>'''
    try:
        root = soap_req(sess, host, body)
    except Exception as e:
        print(f"  [warn] License query failed: {e}")
        return []

    licenses = []
    # licenses is ArrayOfLicenseManagerLicenseInfo
    for lic_e in list(root.iter(f'{{{NS}}}LicenseManagerLicenseInfo')) + list(root.iter('LicenseManagerLicenseInfo')):
        lic_key  = lic_e.findtext(f'{{{NS}}}licenseKey') or lic_e.findtext('licenseKey') or ''
        name     = lic_e.findtext(f'{{{NS}}}name')       or lic_e.findtext('name')       or ''
        total    = lic_e.findtext(f'{{{NS}}}total')      or lic_e.findtext('total')      or '0'
        used     = lic_e.findtext(f'{{{NS}}}used')       or lic_e.findtext('used')       or '0'
        # Expiry is in properties array as expirationHours or expirationDate
        expiry = ''
        for prop_e in list(lic_e.iter(f'{{{NS}}}properties')) + list(lic_e.iter('properties')):
            k = prop_e.findtext(f'{{{NS}}}key') or prop_e.findtext('key') or ''
            if 'expiration' in k.lower():
                v_e = prop_e.find(f'{{{NS}}}value') or prop_e.find('value')
                if v_e is not None:
                    expiry = v_e.findtext(f'{{{NS}}}expirationDate') or v_e.findtext('expirationDate') or v_e.text or ''
        licenses.append({
            "Name":   name,
            "Key":    lic_key[:8] + '...' if len(lic_key) > 8 else lic_key,
            "Total":  total,
            "Used":   int(used or 0),
            "Expiry": str(expiry),
        })
    return licenses

# ── Step 7: Performance counter IDs ──────────────────────────────────────────

def get_counter_ids(sess, host, perf_ref):
    """Return dict: 'group.name.rollup' -> counter_id"""
    body = f'''<vim25:QueryPerfCounterByLevel>
      <vim25:_this type="PerformanceManager">{perf_ref}</vim25:_this>
      <vim25:level>4</vim25:level>
    </vim25:QueryPerfCounterByLevel>'''
    root = soap_req(sess, host, body)
    counters = {}
    for pc in list(root.iter(f'{{{NS}}}returnval')) + list(root.iter('returnval')):
        grp  = pc.findtext(f'{{{NS}}}groupInfo/{{{NS}}}key') or pc.findtext('groupInfo/key') or ''
        nm   = pc.findtext(f'{{{NS}}}nameInfo/{{{NS}}}key')  or pc.findtext('nameInfo/key')  or ''
        roll = pc.findtext(f'{{{NS}}}rollupType') or pc.findtext('rollupType') or ''
        cid  = pc.findtext(f'{{{NS}}}key')        or pc.findtext('key')        or ''
        if grp and nm and cid:
            counters[f'{grp}.{nm}.{roll}'] = int(cid)
    return counters

# ── Step 8: Query performance (VM or Host) ───────────────────────────────────

_VM_COUNTERS = [
    'cpu.usage.average',           # 0-100% of allocated vCPU (x100 in vSphere)
    'cpu.ready.summation',         # ms VM was ready but not scheduled
    'mem.active.average',          # KB of actively used memory
    'mem.consumed.average',        # KB of consumed memory
    'disk.numberRead.summation',   # IOPS reads per sample
    'disk.numberWrite.summation',  # IOPS writes per sample
    'disk.read.average',           # KB/s read throughput
    'disk.write.average',          # KB/s write throughput
]
_HOST_COUNTERS = [
    'cpu.usage.average',
    'mem.usage.average',           # % of host memory used (0-10000)
    'disk.numberRead.summation',
    'disk.numberWrite.summation',
    'disk.read.average',
    'disk.write.average',
]

def query_perf(sess, host, perf_ref, moid, moid_type, counter_ids, days=90):
    """Query historical perf for a VM or HostSystem. Returns {counter_key: [values]}."""
    end_dt   = datetime.now(timezone.utc)
    start_dt = end_dt - timedelta(days=days)
    INTERVAL = 86400  # daily samples for 90-day history

    wanted = {k: counter_ids.get(k) for k in
              (_VM_COUNTERS if moid_type == 'VirtualMachine' else _HOST_COUNTERS)}
    wanted = {k: v for k, v in wanted.items() if v}
    if not wanted:
        return {}

    metrics_xml = ''.join(
        f'<vim25:metricId><vim25:counterId>{cid}</vim25:counterId>'
        f'<vim25:instance></vim25:instance></vim25:metricId>'
        for cid in wanted.values())

    body = f'''<vim25:QueryPerf>
      <vim25:_this type="PerformanceManager">{perf_ref}</vim25:_this>
      <vim25:querySpec>
        <vim25:entity type="{moid_type}">{moid}</vim25:entity>
        <vim25:startTime>{start_dt.strftime("%Y-%m-%dT%H:%M:%SZ")}</vim25:startTime>
        <vim25:endTime>{end_dt.strftime("%Y-%m-%dT%H:%M:%SZ")}</vim25:endTime>
        <vim25:intervalId>{INTERVAL}</vim25:intervalId>
        <vim25:format>normal</vim25:format>
        {metrics_xml}
      </vim25:querySpec>
    </vim25:QueryPerf>'''

    root = soap_req(sess, host, body)
    result = {}
    for entity_metric in list(root.iter(f'{{{NS}}}returnval')) + list(root.iter('returnval')):
        for series in list(entity_metric.iter(f'{{{NS}}}value')) + list(entity_metric.iter('value')):
            cid_e = series.find(f'{{{NS}}}id/{{{NS}}}counterId') or series.find('id/counterId')
            if cid_e is None:
                id_e = series.find(f'{{{NS}}}id') or series.find('id')
                cid_e = id_e.find(f'{{{NS}}}counterId') if id_e is not None else None
            if cid_e is None:
                continue
            cid_val = int(cid_e.text)
            ckey = next((k for k, v in wanted.items() if v == cid_val), str(cid_val))
            vals = []
            for v_e in series.findall(f'{{{NS}}}value') + series.findall('value'):
                try:
                    v = int(v_e.text)
                    if v >= 0:
                        vals.append(v)
                except (TypeError, ValueError):
                    pass
            if vals:
                result[ckey] = vals
    return result

# ── Step 9: Stats helpers ─────────────────────────────────────────────────────

def p95(vals):
    if not vals: return None
    s = sorted(vals)
    return round(s[min(int(len(s) * 0.95), len(s)-1)], 2)

def stats4(fvals):
    if not fvals: return {"Average": None, "Peak": None, "Median": None, "P95": None}
    return {
        "Average": round(mean(fvals), 2),
        "Peak":    round(max(fvals),  2),
        "Median":  round(stat_median(fvals), 2),
        "P95":     p95(fvals),
    }

def cpu_stats(vals):
    """vSphere cpu.usage.average is 0-10000 (= 0-100%)."""
    return stats4([v / 100.0 for v in vals]) if vals else stats4([])

def cpu_ready_p95(vals, vcpus, interval_sec=86400):
    """
    cpu.ready.summation: ms the VM was ready but not running per interval.
    % = (summation_ms / (interval_ms * vcpus)) * 100
    """
    if not vals or not vcpus: return None
    interval_ms = interval_sec * 1000
    pcts = [v / (interval_ms * vcpus) * 100 for v in vals]
    return p95(pcts)

def mem_stats_pct(vals):
    """mem.usage.average (host) is 0-10000 (= 0-100%)."""
    return stats4([v / 100.0 for v in vals]) if vals else stats4([])

def mem_stats_kb(vals, ram_mb):
    """mem.active/consumed in KB → % of provisioned RAM."""
    if not vals or not ram_mb: return stats4([])
    ram_kb = ram_mb * 1024
    return stats4([min(v / ram_kb * 100, 100) for v in vals])

def iops_vals(perf):
    reads  = perf.get('disk.numberRead.summation', [])
    writes = perf.get('disk.numberWrite.summation', [])
    if reads and writes:
        return [r + w for r, w in zip(reads, writes)]
    return reads or writes

def throughput_kbps_p95(perf):
    r = perf.get('disk.read.average', [])
    w = perf.get('disk.write.average', [])
    if r and w: combined = [a + b for a, b in zip(r, w)]
    else:       combined = r or w
    return p95(combined)

# ── Step 10: Build VM output record ──────────────────────────────────────────

def build_vm_output(vm, perf):
    cpu_v   = perf.get('cpu.usage.average', [])
    ready_v = perf.get('cpu.ready.summation', [])
    mem_v   = perf.get('mem.active.average', perf.get('mem.consumed.average', []))
    iops_v  = iops_vals(perf)

    cpu_s = cpu_stats(cpu_v)
    cpu_s['Ready_P95'] = cpu_ready_p95(ready_v, vm['vCPUs'])

    return {
        "Name":           vm['Name'],
        "MOID":           vm['MOID'],
        "PowerState":     vm['PowerState'],
        "GuestOS":        vm['GuestOS'],
        "IsLinux":        vm['IsLinux'],
        "vCPUs":          vm['vCPUs'],
        "RAMgb":          vm['RAMgb'],
        "DiskCapGB":      vm['DiskCapGB'],
        "DiskConsumedGB": vm['DiskConsumedGB'],
        "Datastore":      vm['Datastore'],
        "ToolStatus":     vm['ToolStatus'],
        "CPU":            cpu_s,
        "Memory":         mem_stats_kb(mem_v, vm['RAMmb']),
        "IOPS_P95":       p95(iops_v),
        "DiskKBps_P95":   throughput_kbps_p95(perf),
        "Partitions":     vm['Partitions'],
    }

# ── Main ──────────────────────────────────────────────────────────────────────

def run(vcenter, username, password, days=90, output_dir='.'):
    sess = requests.Session()
    sess.verify = False

    print(f"[1/8] Connecting to vCenter {vcenter}...")
    pc_ref, perf_ref, root_folder, lic_ref = login(sess, vcenter, username, password)
    print(f"      PropertyCollector={pc_ref}  PerfManager={perf_ref}")

    print(f"[2/8] VM inventory + storage + partition data...")
    vms = get_vms(sess, vcenter, pc_ref, root_folder)
    print(f"      {len(vms)} VMs found")

    print(f"[3/8] Host hardware + quickStats...")
    host_info = get_host_info(sess, vcenter, pc_ref, root_folder)
    host_name = host_info.get('Name', vcenter)
    print(f"      Host: {host_name}  Model: {host_info.get('Model','')}  "
          f"CPU: {host_info.get('CPUUsagePct',0):.1f}%  Mem: {host_info.get('MemUsagePct',0):.1f}%")

    print(f"[4/8] Datastores...")
    datastores = get_datastores(sess, vcenter, pc_ref, root_folder)
    print(f"      {len(datastores)} datastores")
    for ds in datastores:
        print(f"      {ds['Name']}  {ds['CapacityGB']:.0f} GB total  {ds['UsedPct']:.0f}% used")

    print(f"[5/8] Cluster info...")
    cluster = get_cluster_info(sess, vcenter, pc_ref, root_folder)
    if cluster:
        print(f"      {cluster.get('Name','')}  CPU:{cluster['CPUUsagePct']:.1f}%  Mem:{cluster['MemUsagePct']:.1f}%")
    else:
        # Standalone host — use host quickStats
        cluster = {
            "Name":          host_name,
            "CPUUsagePct":   host_info.get('CPUUsagePct', 0),
            "MemUsagePct":   host_info.get('MemUsagePct', 0),
            "IOPS_95th":     0,
            "DiskKBps_95th": 0,
            "CapacityMiB":   sum(d['CapacityMiB'] for d in datastores),
            "ConsumedMiB":   sum(d['ConsumedMiB'] for d in datastores),
        }
        print(f"      Standalone host (no cluster object)")

    print(f"[6/8] VMware licenses...")
    licenses = get_licenses(sess, vcenter, pc_ref, lic_ref)
    print(f"      {len(licenses)} license(s) found")

    print(f"[7/8] Performance counter definitions...")
    counter_ids = get_counter_ids(sess, vcenter, perf_ref)
    print(f"      {len(counter_ids)} counters available")
    for k in ['cpu.usage.average','cpu.ready.summation','mem.active.average',
              'disk.numberRead.summation','disk.read.average']:
        print(f"      {k}: ID={counter_ids.get(k,'N/A')}")

    print(f"[8/8] {days}-day perf history for {len(vms)} VMs + host...")
    vm_results = []
    for vm in vms:
        if vm['PowerState'] not in ('poweredOn', 'POWERED_ON'):
            print(f"  {vm['Name']}: powered off — skipping perf")
            vm_results.append(build_vm_output(vm, {}))
            continue
        if vm['IsLinux']:
            print(f"  {vm['Name']}: Linux — skipping Windows perf metrics")
            vm_results.append(build_vm_output(vm, {}))
            continue
        print(f"  {vm['Name']} ({vm['MOID']})...", end='', flush=True)
        try:
            perf = query_perf(sess, vcenter, perf_ref, vm['MOID'], 'VirtualMachine', counter_ids, days)
            pts = sum(len(v) for v in perf.values())
            print(f" {pts} data pts")
        except Exception as e:
            print(f" ERR: {e}")
            perf = {}
        vm_results.append(build_vm_output(vm, perf))

    # Host-level perf (for IOPS/throughput P95 on host and cluster)
    print(f"  Host {host_name} perf...", end='', flush=True)
    try:
        host_moid = ''
        # Get host MOID from one of our HostSystem queries
        host_rows = retrieve_all(sess, vcenter, pc_ref, root_folder, 'HostSystem', ['name'])
        host_moid = host_rows[0]['_moid'] if host_rows else ''
        host_perf = query_perf(sess, vcenter, perf_ref, host_moid, 'HostSystem', counter_ids, days) if host_moid else {}
        host_iops = iops_vals(host_perf)
        host_tp   = throughput_kbps_p95(host_perf)
        host_info['IOPS_95th']    = p95(host_iops) or 0
        host_info['DiskKBps_95th'] = host_tp or 0
        cluster['IOPS_95th']     = host_info['IOPS_95th']
        cluster['DiskKBps_95th'] = host_info['DiskKBps_95th']
        print(f" IOPS P95={host_info['IOPS_95th']}  DiskKBps P95={host_info['DiskKBps_95th']}")
    except Exception as e:
        print(f" ERR: {e}")

    # Fill cluster storage from datastores
    if not cluster.get('CapacityMiB'):
        cluster['CapacityMiB'] = round(sum(d['CapacityMiB'] for d in datastores), 1)
        cluster['ConsumedMiB'] = round(sum(d['ConsumedMiB'] for d in datastores), 1)

    # Assemble output
    date_str = datetime.now().strftime('%Y-%m-%d')
    out = {
        "_type":        "vSpherePerf",
        "_source":      f"collect_vsphere_perf.py — direct SOAP — {days} days",
        "CollectedAt":  datetime.now().isoformat(),
        "DurationDays": days,
        "Host":         host_info,
        "Cluster":      cluster,
        "VMs":          vm_results,
        "Datastores":   datastores,
        "Licenses":     licenses,
    }

    os.makedirs(output_dir, exist_ok=True)
    out_path = os.path.join(output_dir, f"vsphere-perf-{date_str}.json")
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(out, f, indent=2, default=str)

    print(f"\nWritten: {out_path}  ({os.path.getsize(out_path):,} bytes)")
    print(f"VMs:        {len(vm_results)}   ({sum(1 for v in vm_results if v['CPU']['P95'] is not None)} with perf data)")
    print(f"Datastores: {len(datastores)}")
    print(f"Licenses:   {len(licenses)}")
    return out_path


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='vSphere full SOAP collector — Nutanix Collector replacement')
    parser.add_argument('--vcenter', required=True)
    parser.add_argument('--user',    required=True)
    parser.add_argument('--pass',    required=True, dest='password')
    parser.add_argument('--days',    type=int, default=90)
    parser.add_argument('--output',  default='.')
    args = parser.parse_args()
    run(args.vcenter, args.user, args.password, args.days, args.output)
