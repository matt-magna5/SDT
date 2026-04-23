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
    # PropertyFilterSpec fields are propSet (collection of PropertySpec) and
    # objectSet (collection of ObjectSpec). ESXi 7 was lenient with the
    # singular '*Spec' tags but vCenter 7/8 enforces the schema strictly.
    body = f'''<vim25:RetrievePropertiesEx>
      <vim25:_this type="PropertyCollector">{pc_ref}</vim25:_this>
      <vim25:specSet>
        <vim25:propSet>
          <vim25:type>{obj_type}</vim25:type>
          <vim25:all>false</vim25:all>
          {ps}
        </vim25:propSet>
        <vim25:objectSet>
          <vim25:obj type="Folder">{root_folder}</vim25:obj>
          <vim25:skip>false</vim25:skip>
          {_TRAV}
        </vim25:objectSet>
      </vim25:specSet>
      <vim25:options/>
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

# Global pyVmomi ServiceInstance - set by login(), used by inventory helpers
_SI = None

# ── Step 1: Login  ───────────────────────────────────────────────────────────

def login(sess, host, user, pw):
    """Returns (pc_ref, perf_ref, root_folder, lic_mgr_ref).

    Uses pyVmomi's SmartConnect to do the auth handshake, then extracts the
    authenticated SOAP session cookie and splices it into our custom `sess`
    so the rest of this script (which talks raw SOAP for perf collection
    performance) reuses the authenticated session.

    pyVmomi handles: SAML/SSO, identity federation, vCenter 7/8 enhanced
    linked mode, XML escaping, certificate trust, session renewal — all the
    modern-vCenter complexity that hand-rolled SOAP Login can't.
    """
    # Lazy imports so the script still runs for users who install pyVmomi later
    try:
        from pyVim.connect import SmartConnect, Disconnect
        import ssl as _ssl
    except ImportError as e:
        raise RuntimeError(f"pyVmomi not installed: {e}. Run: python -m pip install pyVmomi") from e

    # Build an insecure SSL context — mirrors our requests.verify=False elsewhere
    _ctx = _ssl.create_default_context()
    _ctx.check_hostname = False
    _ctx.verify_mode = _ssl.CERT_NONE

    # Parse host:port
    _host = host
    _port = 443
    if ':' in host:
        _host, _p = host.rsplit(':', 1)
        try: _port = int(_p)
        except: pass

    print(f"      pyVmomi SmartConnect to {_host}:{_port} as {user} ...")
    try:
        si = SmartConnect(host=_host, port=_port, user=user, pwd=pw,
                          sslContext=_ctx, disableSslCertValidation=True)
    except Exception as e:
        # Surface the underlying cause clearly - vim.fault.InvalidLogin, etc.
        err_name = type(e).__name__
        raise RuntimeError(f"vCenter login failed ({err_name}): {e}") from e

    # Extract authenticated vmware_soap_session cookie from pyVmomi's stub
    try:
        cookie_hdr = si._stub.cookie or ''
    except Exception:
        cookie_hdr = ''
    m = re.search(r'vmware_soap_session="?([^";]+)"?', cookie_hdr)
    if m:
        sess.cookies.set('vmware_soap_session', m.group(1))
    else:
        # Fallback: save the entire cookie header directly
        sess.headers['Cookie'] = cookie_hdr

    # Pull the Managed Object References we need for the rest of the script
    content = si.RetrieveContent()
    pc_ref   = content.propertyCollector._moId if content.propertyCollector else None
    perf_ref = content.perfManager._moId       if content.perfManager       else None
    rf_ref   = content.rootFolder._moId        if content.rootFolder        else None
    lic_ref  = content.licenseManager._moId    if content.licenseManager    else None

    # Save ServiceInstance for inventory helpers to use pyVmomi directly
    global _SI
    _SI = si

    # Don't Disconnect - we want to keep the session alive for subsequent SOAP calls
    return pc_ref, perf_ref, rf_ref, lic_ref


# ── pyVmomi inventory helpers - used by all get_* functions below ────────────

def _pv_get_all(obj_type):
    """Return list of pyVmomi managed objects of obj_type under rootFolder."""
    if _SI is None:
        raise RuntimeError("pyVmomi session missing - login() must run first")
    from pyVmomi import vim
    type_map = {
        'VirtualMachine':         [vim.VirtualMachine],
        'HostSystem':             [vim.HostSystem],
        'Datastore':              [vim.Datastore],
        'ClusterComputeResource': [vim.ClusterComputeResource],
        'ComputeResource':        [vim.ComputeResource],
        'Datacenter':             [vim.Datacenter],
        'Network':                [vim.Network],
    }
    vtype = type_map.get(obj_type)
    if vtype is None:
        return []
    content = _SI.RetrieveContent()
    view = content.viewManager.CreateContainerView(content.rootFolder, vtype, True)
    try:
        return list(view.view)
    finally:
        view.Destroy()


def _pv_attr(obj, path, default=None):
    """Walk dotted attribute path on a pyVmomi object, returning default on any miss."""
    try:
        cur = obj
        for part in path.split('.'):
            cur = getattr(cur, part, None)
            if cur is None:
                return default
        return cur
    except Exception:
        return default

# ── Step 2: VM inventory ──────────────────────────────────────────────────────

def get_vms(sess, host, pc_ref, root_folder):
    """Return list of VM dicts with config + storage + partition data (pyVmomi-based)."""
    vms = []
    for vm in _pv_get_all('VirtualMachine'):
        name = _pv_attr(vm, 'name')
        if not name:
            continue

        # Extract datastore from vmPathName: "[DatastoreName] vm/vm.vmx"
        vmx = _pv_attr(vm, 'summary.config.vmPathName', '') or ''
        ds_match = re.search(r'\[(.+?)\]', vmx)
        datastore = ds_match.group(1) if ds_match else ''

        committed   = int(_pv_attr(vm, 'summary.storage.committed',   0) or 0)
        uncommitted = int(_pv_attr(vm, 'summary.storage.uncommitted', 0) or 0)
        cap_gb  = round((committed + uncommitted) / (1024**3), 1)
        cons_gb = round(committed / (1024**3), 1)

        # Guest partitions - list of vim.vm.GuestInfo.DiskInfo
        partitions = []
        gdisks = _pv_attr(vm, 'guest.disk') or []
        try:
            for gdi in gdisks:
                disk_cap  = int(getattr(gdi, 'capacity',  0) or 0)
                disk_free = int(getattr(gdi, 'freeSpace', 0) or 0)
                partitions.append({
                    "Path":        getattr(gdi, 'diskPath', '') or '',
                    "CapacityMiB": round(disk_cap / (1024**2), 1),
                    "ConsumedMiB": round((disk_cap - disk_free) / (1024**2), 1),
                })
        except Exception:
            pass

        ram_mb   = int(_pv_attr(vm, 'config.hardware.memoryMB', 0) or 0)
        vcpu     = int(_pv_attr(vm, 'config.hardware.numCPU', 0) or 0)
        guest_os = (_pv_attr(vm, 'guest.guestFullName') or
                    _pv_attr(vm, 'config.guestFullName') or '')
        is_linux = bool(guest_os and any(k in guest_os.lower() for k in
                   ['linux','ubuntu','centos','rhel','debian','photon','suse','oracle',
                    'rocky','alma','amazon','coreos','fedora']))
        if not is_linux and partitions:
            is_linux = any(p['Path'].startswith('/') for p in partitions)

        # Cast enum-ish fields to strings
        power_state = str(_pv_attr(vm, 'runtime.powerState', '') or '')
        tool_status = str(_pv_attr(vm, 'guest.toolsStatus', '') or '')

        vms.append({
            'Name':       name,
            'MOID':       vm._moId,
            'PowerState': power_state,
            'GuestOS':    guest_os,
            'IsLinux':    is_linux,
            'vCPUs':      vcpu,
            'RAMmb':      ram_mb,
            'RAMgb':      round(ram_mb / 1024, 1),
            'ToolStatus': tool_status,
            'Datastore':  datastore,
            'DiskCapGB':  cap_gb,
            'DiskConsumedGB': cons_gb,
            'Partitions': partitions,
        })
    return vms

# ── Step 3: Host info + host-level perf ──────────────────────────────────────

def get_host_info(sess, host_addr, pc_ref, root_folder):
    """Return FIRST host hardware dict (pyVmomi-based)."""
    hosts = _pv_get_all('HostSystem')
    if not hosts:
        return {}
    h = hosts[0]
    cores   = int(_pv_attr(h, 'hardware.cpuInfo.numCpuCores', 0) or 0)
    mhz_per = int(_pv_attr(h, 'summary.hardware.cpuMhz', 0) or 0)
    total_mhz  = cores * mhz_per
    used_mhz   = int(_pv_attr(h, 'summary.quickStats.overallCpuUsage', 0) or 0)
    total_mem_mb = int(_pv_attr(h, 'hardware.memorySize', 0) or 0) // (1024**2)
    used_mem_mb  = int(_pv_attr(h, 'summary.quickStats.overallMemoryUsage', 0) or 0)
    cpu_pct = round(used_mhz / total_mhz * 100, 2) if total_mhz else 0
    mem_pct = round(used_mem_mb / total_mem_mb * 100, 2) if total_mem_mb else 0
    return {
        "Name":        _pv_attr(h, 'name') or _pv_attr(h, 'summary.config.name', ''),
        "Model":       _pv_attr(h, 'hardware.systemInfo.model', ''),
        "Vendor":      _pv_attr(h, 'hardware.systemInfo.vendor', ''),
        "ServiceTag":  _pv_attr(h, 'hardware.systemInfo.serialNumber', ''),
        "Cores":       cores,
        "CPUModel":    _pv_attr(h, 'summary.hardware.cpuModel', ''),
        "CPUSpeedMHz": int(int(_pv_attr(h, 'hardware.cpuInfo.hz', 0) or 0) / 1_000_000),
        "RAMgb":       round(int(_pv_attr(h, 'hardware.memorySize', 0) or 0) / (1024**3), 2),
        "Hypervisor":  _pv_attr(h, 'config.product.fullName', ''),
        "NICs":        int(_pv_attr(h, 'summary.hardware.numNics', 0) or 0),
        "CPUUsagePct": cpu_pct,
        "MemUsagePct": mem_pct,
        "IOPS_95th":   0,
        "DiskKBps_95th": 0,
    }

# ── Step 4: Datastores ───────────────────────────────────────────────────────

def get_datastores(sess, host, pc_ref, root_folder):
    """Return list of datastore dicts (pyVmomi-based)."""
    ds_list = []
    for ds in _pv_get_all('Datastore'):
        name = _pv_attr(ds, 'name')
        if not name:
            continue
        cap  = int(_pv_attr(ds, 'summary.capacity',  0) or 0)
        free = int(_pv_attr(ds, 'summary.freeSpace', 0) or 0)
        used = cap - free
        ds_list.append({
            "Name":        name,
            "Type":        str(_pv_attr(ds, 'summary.type', 'VMFS') or 'VMFS'),
            "CapacityGB":  round(cap  / (1024**3), 1),
            "FreeGB":      round(free / (1024**3), 1),
            "ConsumedGB":  round(used / (1024**3), 1),
            "UsedPct":     round(used / cap * 100, 1) if cap else 0,
            "CapacityMiB": round(cap  / (1024**2), 1),
            "ConsumedMiB": round(used / (1024**2), 1),
        })
    return ds_list

# ── Step 5: Cluster info ──────────────────────────────────────────────────────

def get_cluster_info(sess, host, pc_ref, root_folder):
    """Return first-cluster aggregate dict (pyVmomi-based). Empty for standalone host."""
    clusters = _pv_get_all('ClusterComputeResource')
    if not clusters:
        return {}
    c = clusters[0]
    total_mhz = int(_pv_attr(c, 'summary.totalCpu', 0) or 0)
    used_mhz  = int(_pv_attr(c, 'summary.currentCpuUsage', 0) or 0)
    total_mem_b  = int(_pv_attr(c, 'summary.totalMemory', 0) or 0)
    used_mem_mb  = int(_pv_attr(c, 'summary.currentMemoryUsage', 0) or 0)
    total_mem_mb = total_mem_b // (1024**2)
    return {
        "Name":        _pv_attr(c, 'name', ''),
        "CPUUsagePct": round(used_mhz / total_mhz * 100, 2) if total_mhz else 0,
        "MemUsagePct": round(used_mem_mb / total_mem_mb * 100, 2) if total_mem_mb else 0,
        "CapacityMiB": 0,
        "ConsumedMiB": 0,
        "IOPS_95th":   0,
        "DiskKBps_95th": 0,
        "NumHosts":    int(_pv_attr(c, 'summary.numHosts', 0) or 0),
    }

# ── Step 6: Licenses ──────────────────────────────────────────────────────────

def get_licenses(sess, host, pc_ref, lic_ref):
    """Return list of license dicts (pyVmomi-based)."""
    if _SI is None:
        return []
    try:
        content = _SI.RetrieveContent()
        lic_mgr = content.licenseManager
        if lic_mgr is None:
            return []
    except Exception as e:
        print(f"  [warn] License manager access failed: {e}")
        return []

    licenses = []
    try:
        for lic in (lic_mgr.licenses or []):
            lic_key = getattr(lic, 'licenseKey', '') or ''
            expiry = ''
            try:
                for prop in (getattr(lic, 'properties', []) or []):
                    k = getattr(prop, 'key', '') or ''
                    if 'expiration' in k.lower():
                        val = getattr(prop, 'value', None)
                        if val is not None:
                            expiry = str(getattr(val, 'expirationDate', '') or val)
            except Exception:
                pass
            licenses.append({
                "Name":   getattr(lic, 'name', '') or '',
                "Key":    lic_key[:8] + '...' if len(lic_key) > 8 else lic_key,
                "Total":  str(getattr(lic, 'total', 0) or 0),
                "Used":   int(getattr(lic, 'used', 0) or 0),
                "Expiry": expiry,
            })
    except Exception as e:
        print(f"  [warn] License enumeration failed: {e}")
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

    # Assemble output. Schema + filename match what gen_report.py picks up
    # as a vSphereInventory (filename contains 'inventory' + top-level _type).
    date_str = datetime.now().strftime('%Y-%m-%d')
    # Sanitize host for filename
    host_safe = re.sub(r'[^A-Za-z0-9_.-]', '_', vcenter)
    out = {
        "_type":        "vSphereInventory",
        "_source":      f"collect_vsphere_perf.py via pyVmomi + SOAP - {days} days",
        "Server":       vcenter,
        "Version":      host_info.get('Hypervisor', '') if isinstance(host_info, dict) else '',
        "APIVersion":   host_info.get('Hypervisor', '') if isinstance(host_info, dict) else '',
        "CollectedAt":  datetime.now().isoformat(),
        "DurationDays": days,
        "Host":         host_info,
        "ESXHosts":     [host_info] if host_info else [],
        "Cluster":      cluster,
        "VMs":          vm_results,
        "Datastores":   datastores,
        "Licenses":     licenses,
    }

    os.makedirs(output_dir, exist_ok=True)
    out_path = os.path.join(output_dir, f"{host_safe}-inventory-{date_str}.json")
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(out, f, indent=2, default=str)

    print(f"\nWritten: {out_path}  ({os.path.getsize(out_path):,} bytes)")
    print(f"VMs:        {len(vm_results)}   ({sum(1 for v in vm_results if v['CPU']['P95'] is not None)} with perf data)")
    print(f"Datastores: {len(datastores)}")
    print(f"Licenses:   {len(licenses)}")
    return out_path


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='vSphere full SOAP collector - Nutanix Collector replacement')
    parser.add_argument('--vcenter', required=True)
    parser.add_argument('--user',    required=True)
    parser.add_argument('--pass',    dest='password', default=None,
                        help='Password (or set SDT_HV_PASS env var to avoid CLI escaping issues)')
    parser.add_argument('--pass-env', dest='pass_env', default=None,
                        help='Name of env var holding the password (safer than --pass for special chars)')
    parser.add_argument('--days',    type=int, default=120)
    parser.add_argument('--output',  default='.')
    args = parser.parse_args()

    # Resolve password: explicit env var name > default SDT_HV_PASS env var > --pass arg
    password = None
    if args.pass_env:
        password = os.environ.get(args.pass_env)
        if password is None:
            print(f"[error] env var {args.pass_env} is not set", file=sys.stderr)
            sys.exit(2)
    elif os.environ.get('SDT_HV_PASS'):
        password = os.environ['SDT_HV_PASS']
    elif args.password is not None:
        password = args.password
    else:
        print("[error] no password provided (use --pass, --pass-env, or set SDT_HV_PASS)", file=sys.stderr)
        sys.exit(2)

    run(args.vcenter, args.user, password, args.days, args.output)
