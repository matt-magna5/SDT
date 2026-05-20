"""Shared loader/utilities for SDT report mockups.
Loads a session dir (manifest.json + per-server JSONs) and normalizes the data.
"""
import json, os, re, html
from pathlib import Path

# Flags we treat as universal noise - hide from per-server view, surface
# once as an environment-wide rollup instead.
NOISE_FLAGS = {
    'SMB 1.0',
    'Windows Defender',  # default AV everywhere; only flag if no third-party EDR found
}

def _safe(d, *keys, default=None):
    cur = d
    for k in keys:
        if not isinstance(cur, dict): return default
        cur = cur.get(k)
        if cur is None: return default
    return cur

def load_session(session_dir):
    session_dir = Path(session_dir)
    manifest = json.load(open(session_dir / 'manifest.json', encoding='utf-8-sig'))
    client = manifest.get('client_full') or manifest.get('client') or 'Client'
    date = manifest.get('date', '')
    servers = []
    for s in manifest.get('servers', []):
        fp = session_dir / s.get('file', '')
        if not fp.exists(): continue
        try:
            d = json.load(open(fp, encoding='utf-8-sig'))
        except Exception:
            continue
        srv = normalize_server(d, s, fp.name)
        if srv: servers.append(srv)
    # Load vSphere inventory if present
    vcenter = None
    inv_name = manifest.get('inventory_file')
    if inv_name:
        ip = session_dir / inv_name
        if ip.exists():
            try:
                vcenter = json.load(open(ip, encoding='utf-8-sig'))
            except Exception: pass
    return {'client': client, 'date': date, 'servers': servers, 'vcenter': vcenter}

def normalize_server(d, manifest_entry, filename):
    """Pull a stable set of fields from a raw discovery JSON."""
    sys_ = d.get('System', {}) or {}
    hw   = d.get('Hardware', {}) or {}
    name = manifest_entry.get('name') or sys_.get('Hostname') or filename.split('-discovery')[0]
    flags_raw = d.get('Flags', [])
    # Some JSONs (CLI-XML round-tripped) emit an empty Hashtable shape - skip those
    if not isinstance(flags_raw, list):
        flags_raw = []
    flags = [f for f in flags_raw if isinstance(f, dict) and f.get('Title')]
    return {
        'name': name,
        'id':   manifest_entry.get('id') or re.sub(r'[^a-z0-9]+','-', name.lower()).strip('-'),
        'os':   sys_.get('OSName', ''),
        'os_eol': sys_.get('OSEOLStatus', ''),
        'uptime_days': sys_.get('UptimeDays'),
        'domain': sys_.get('Domain', ''),
        'cpu':  hw.get('CPUName') or hw.get('CPUModel') or '',
        'cores': hw.get('CPUCores'),
        'ram_gb': hw.get('RAMTotalGB'),
        'platform': hw.get('VMPlatform') or ('VM' if hw.get('IsVM') else 'Physical'),
        'mfr':  hw.get('Manufacturer', ''),
        'model': hw.get('Model', ''),
        'disks': d.get('Disks', []) if isinstance(d.get('Disks'), list) else [],
        'flags': flags,
        'apps':  d.get('Apps', []) if isinstance(d.get('Apps'), list) else [],
        'services_count': len(d.get('Services', [])) if isinstance(d.get('Services'), list) else 0,
        'sql':   d.get('SQL', {}) if isinstance(d.get('SQL'), dict) else {},
        'shares': _safe(d, 'FileShares', 'Shares', default=[]) or [],
        'iis':   d.get('IIS', {}) if isinstance(d.get('IIS'), dict) else {},
        'ad':    d.get('AD') if isinstance(d.get('AD'), dict) else None,
        'roles': _safe(d, 'Roles', 'InstalledRoles', default=[]) or [],
        'hyperv_vms': _safe(d, 'HyperV', 'VMs', default=[]) or [],
        'network_adapters': _safe(d, 'Network', 'Adapters', default=[]) or [],
    }

def detect_security(srv):
    """Scan apps + services for known EDR/RMM products. Returns dict of category -> product name."""
    apps_str = ' '.join((str(a.get('Name',''))+' '+str(a.get('Publisher',''))).lower() for a in srv['apps'])
    found = {}
    edr_map = {
        'sentinelone':'SentinelOne', 'crowdstrike':'CrowdStrike', 'huntress':'Huntress',
        'sophos':'Sophos', 'cylance':'Cylance', 'trellix':'Trellix', 'mcafee':'McAfee',
        'eset':'ESET', 'kaspersky':'Kaspersky', 'webroot':'Webroot', 'bitdefender':'Bitdefender',
        'malwarebytes':'Malwarebytes', 'cortex xdr':'Cortex XDR', 'carbon black':'Carbon Black',
        'cybereason':'Cybereason', 'darktrace':'Darktrace', 'adlumin':'Adlumin',
        'arctic wolf':'Arctic Wolf', 'defender for endpoint':'Defender for Endpoint',
    }
    rmm_map = {
        'n-able':'N-able', 'solarwinds':'N-able (SolarWinds)', 'kaseya':'Kaseya',
        'connectwise':'ConnectWise', 'labtech':'ConnectWise Automate',
        'ninjarmm':'NinjaOne', 'ninjaone':'NinjaOne', 'datto rmm':'Datto RMM',
        'syncro':'Syncro', 'atera':'Atera', 'pulseway':'Pulseway',
    }
    backup_map = {
        'veeam':'Veeam', 'commvault':'Commvault', 'acronis':'Acronis',
        'datto bcdr':'Datto BCDR', 'axcient':'Axcient', 'arcserve':'Arcserve',
        'shadowprotect':'ShadowProtect', 'backup exec':'Veritas Backup Exec',
        'cohesity':'Cohesity', 'rubrik':'Rubrik', 'azure backup':'Azure Backup',
    }
    rmt_map = {
        'screenconnect':'ScreenConnect', 'connectwise control':'ScreenConnect',
        'teamviewer':'TeamViewer', 'anydesk':'AnyDesk', 'splashtop':'Splashtop',
        'logmein':'LogMeIn', 'bomgar':'BeyondTrust', 'dameware':'Dameware',
    }
    def scan(m): return [v for k,v in m.items() if k in apps_str]
    found['edr']    = scan(edr_map)
    found['rmm']    = scan(rmm_map)
    found['backup'] = scan(backup_map)
    found['remote'] = scan(rmt_map)
    return found

def env_rollup(servers):
    """Compute environment-wide stats + identify universal-vs-spotty findings."""
    total = len(servers)
    if not total: return {}
    os_counts = {}
    eol_at_risk = []
    no_edr = []
    high_disk = []
    sql_servers = []
    has_smb1 = []
    has_veeam = []
    by_platform = {}
    for s in servers:
        os_counts[s['os']] = os_counts.get(s['os'], 0) + 1
        by_platform[s.get('platform') or '?'] = by_platform.get(s.get('platform') or '?', 0) + 1
        if s.get('os_eol') and 'eol' in str(s['os_eol']).lower():
            eol_at_risk.append(s['name'])
        sec = detect_security(s)
        if not sec['edr']:
            no_edr.append(s['name'])
        for d in s['disks']:
            if d.get('UsedPct') and d['UsedPct'] >= 85:
                high_disk.append(f"{s['name']}:{d.get('Drive')} ({d['UsedPct']}%)")
        if isinstance(s['sql'], dict) and s['sql'].get('Instances', {}).get('InstanceName'):
            sql_servers.append(s['name'])
        for f in s['flags']:
            if 'smb' in str(f.get('Title','')).lower() and '1.0' in str(f.get('Title','')):
                has_smb1.append(s['name'])
        if 'veeam' in ' '.join(a.get('Name','').lower() for a in s['apps']):
            has_veeam.append(s['name'])
    return {
        'total': total,
        'os_counts': os_counts,
        'by_platform': by_platform,
        'eol_at_risk': eol_at_risk,
        'no_edr': no_edr,
        'high_disk': high_disk,
        'sql_servers': sql_servers,
        'has_smb1': has_smb1,
        'has_veeam': has_veeam,
    }

def filtered_flags(srv):
    """Return per-server flags with universal noise removed."""
    out = []
    for f in srv['flags']:
        t = str(f.get('Title','')).lower()
        if 'smb' in t and '1.0' in t: continue
        out.append(f)
    return out

def h(s):
    return html.escape(str(s)) if s is not None else ''
