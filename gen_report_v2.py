#!/usr/bin/env python3
"""
gen_report_v2.py — Magna5 SDT Discovery Report (v2 redesign)

Surfaces EVERYTHING in SDT's per-server + hypervisor JSONs. 4 tabs:
  Summary | Servers | Environment | Risks

Usage:
    python gen_report_v2.py <session_dir> [--client "Client Name"] [--out path.html]
"""
from __future__ import annotations
import argparse, datetime, html, json, re, sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
TODAY = datetime.date.today()
TODAY_STR = TODAY.strftime('%Y-%m-%d')
VERSION = 'v2.1-beta'

# ─── JSON loading + normalization ──────────────────────────────────────
def _read_json(p: Path, default=None):
    if not p.exists(): return default
    raw = p.read_bytes()
    for enc in ('utf-8-sig','utf-16','utf-16-le','utf-8'):
        try: return json.loads(raw.decode(enc))
        except Exception: continue
    return default

def _is_arraylist_empty(v) -> bool:
    """Detect PowerShell ArrayList serialized as dict (empty collection)."""
    if not isinstance(v, dict): return False
    keys = set(v.keys())
    return keys.issuperset({'Capacity','Count','IsFixedSize','IsReadOnly','IsSynchronized','SyncRoot'}) and v.get('Count',0) == 0

def _unwrap_script_prefixed_list(v):
    """PS sometimes writes [script_fragment_str, real_data_dict]. Find the dict."""
    if isinstance(v, list) and v:
        # Filter out PS script fragments; return first real dict, or the filtered list
        dicts = [x for x in v if isinstance(x, dict)]
        non_str = [x for x in v if not (isinstance(x, str) and ('[void]' in x or '$cb' in x or 'PSCustomObject' in x))]
        if dicts and len(dicts) == 1 and len(v) <= 3:
            return dicts[0]
        return non_str
    return v

def normalize(v):
    """Recursively clean ArrayList+script-prefix artifacts."""
    if _is_arraylist_empty(v):
        return []
    v = _unwrap_script_prefixed_list(v)
    if isinstance(v, dict):
        return {k: normalize(val) for k, val in v.items()}
    if isinstance(v, list):
        return [normalize(x) for x in v]
    return v

DETECTION_RULES = _read_json(SCRIPT_DIR / 'detection_rules.json', {}) or {}
HARDWARE_EOL    = _read_json(SCRIPT_DIR / 'hardware_eol.json', {'models': []}) or {'models': []}

# Reuse v1's rich helpers verbatim so we don't reinvent categorization.
sys.path.insert(0, str(SCRIPT_DIR))
try:
    import gen_report as v1
    V1_RULES = v1.RULES
    v1_categorize_svcs  = v1.categorize_svcs
    v1_categorize_apps  = v1.categorize_apps
    v1_detect_security  = v1.detect_security
    v1_find_anomalies   = v1.find_svc_anomalies
    v1_as_list          = v1.as_list
except Exception:
    V1_RULES = {}
    v1_categorize_svcs  = lambda s: {'EDR':[],'PAM':[],'RMM':[],'HyperV':[],'Core':[],'Print':[],'Other':[],'StoppedAuto':[]}
    v1_categorize_apps  = lambda a: {'Security':[],'Management':[],'LOB':[],'Browser':[],'Other':[]}
    v1_detect_security  = lambda a,s: (None,None,None,None,None)
    v1_find_anomalies   = lambda s: []
    v1_as_list          = lambda v: v if isinstance(v, list) else []

# ─── HTML helpers ──────────────────────────────────────────────────────
def h(s):
    return html.escape('' if s is None else str(s))

_PILL_COLOR_MAP = {'red':'crit','yellow':'warn','green':'ok','blue':'info','purple':'info','gray':'neutral',
                   'crit':'crit','warn':'warn','ok':'ok','info':'info','neutral':'neutral'}

def pill(text, color='neutral', size=None):
    cls = _PILL_COLOR_MAP.get(color, 'neutral')
    return f'<span class="pill {cls}"><span class="pill-dot"></span>{h(text)}</span>'

def bar(pct, color=None):
    pct = max(0.0, min(100.0, float(pct or 0)))
    if color is None:
        color = 'var(--crit)' if pct >= 85 else ('var(--warn)' if pct >= 70 else 'var(--success)')
    return f'<div class="pbar"><div style="background:{color};width:{pct:.0f}%"></div></div>'

def kv_table(rows, col_width='180px'):
    body = ''
    for k, v in rows:
        if v in (None, '', '—'): continue
        body += f'<tr><td class="k">{h(k)}</td><td class="v">{v}</td></tr>'
    return f'<table class="kv" style="width:100%;border-collapse:collapse;">{body}</table>' if body else ''

def section_hdr(title, subtitle=''):
    s = f'<div class="card-sub">{h(subtitle)}</div>' if subtitle else ''
    return f'<div class="section-hdr">{h(title)}</div>{s}'

def card(body, extra_class=''):
    return f'<div class="card {extra_class}">{body}</div>'

# ─── Server wrapper ────────────────────────────────────────────────────
class Server:
    def __init__(self, data: dict, path: Path):
        self.d = data
        self.path = path
        self.name = (self.d.get('System',{}).get('Hostname')
                     or self.d.get('System',{}).get('HostName')
                     or path.stem.split('-discovery-')[0])

    def get(self, *keys, default=None):
        cur = self.d
        for k in keys:
            if not isinstance(cur, dict): return default
            cur = cur.get(k)
            if cur is None: return default
        return cur

    # System
    @property
    def os_name(self): return self.get('System','OSName', default='') or ''
    @property
    def os_version(self): return self.get('System','OSVersion', default='')
    @property
    def os_build(self): return self.get('System','OSBuild', default='')
    @property
    def os_eol_date(self): return self.get('System','OSEOLDate', default='')
    @property
    def os_eol_status(self): return self.get('System','OSEOLStatus', default='')
    @property
    def domain(self): return self.get('System','Domain', default='')
    @property
    def hostname(self): return self.get('System','Hostname', default=self.name)
    @property
    def uptime_days(self): return self.get('System','UptimeDays', default=None)
    @property
    def last_boot(self): return self.get('System','LastBoot', default='')
    @property
    def timezone(self): return self.get('System','Timezone', default='')
    @property
    def os_install_date(self): return self.get('System','OSInstallDate', default='')
    @property
    def ps_version(self): return self.get('System','PSVersion', default='')
    @property
    def run_as(self): return self.get('System','RunAsUser', default='')
    # Hardware
    @property
    def manufacturer(self): return self.get('Hardware','Manufacturer', default='') or ''
    @property
    def model(self): return self.get('Hardware','Model', default='') or ''
    @property
    def cpu_name(self): return self.get('Hardware','CPUName', default='') or ''
    @property
    def cpu_cores(self): return self.get('Hardware','CPUCores', default=None)
    @property
    def ram_total(self): return self.get('Hardware','RAMTotalGB', default=None)
    @property
    def ram_avail(self): return self.get('Hardware','RAMAvailGB', default=None)
    @property
    def serial(self): return self.get('Hardware','SerialNumber', default='')
    @property
    def bios_ver(self): return self.get('Hardware','BIOSVersion', default='')
    @property
    def bios_date(self): return self.get('Hardware','BIOSDate', default='')
    @property
    def is_vm(self): return bool(self.get('Hardware','IsVM', default=False))
    @property
    def vm_platform(self): return self.get('Hardware','VMPlatform', default='')
    # Roles
    @property
    def roles(self):
        r = self.get('Roles','InstalledRoles', default=[]) or []
        return [x for x in r if isinstance(x, dict)]
    @property
    def role_names(self):
        return [r.get('Name','') for r in self.roles]
    @property
    def role_displays(self):
        return [r.get('DisplayName', r.get('Name','')) for r in self.roles]
    @property
    def features(self):
        f = self.get('Roles','InstalledFeatures', default=[]) or []
        return [x for x in f if isinstance(x, dict)]
    @property
    def feature_names(self):
        return [f.get('Name','') for f in self.features]
    # Disks
    @property
    def disks(self):
        d = self.d.get('Disks', [])
        if isinstance(d, list):
            return [x for x in d if isinstance(x, dict)]
        # some discoveries use Disks.Volumes
        v = self.get('Disks','Volumes', default=[])
        return [x for x in (v or []) if isinstance(x, dict)]
    # SQL
    @property
    def sql_instances(self):
        s = self.get('SQL','Instances', default=[]) or []
        return [x for x in s if isinstance(x, dict)]
    # Exchange
    @property
    def exchange(self):
        e = self.d.get('Exchange')
        if not isinstance(e, dict): return {}
        return e
    @property
    def has_exchange(self): return bool(self.exchange.get('Installed'))
    # AD
    @property
    def ad(self):
        a = self.d.get('AD')
        # AD may be a list [script_str, dict]; normalize already unwraps
        if isinstance(a, dict): return a
        if isinstance(a, list):
            for x in a:
                if isinstance(x, dict): return x
        return {}
    @property
    def is_dc(self): return 'ADDS' in self.role_names or bool(self.ad.get('DCCount'))
    # DNS / DHCP / NPS
    @property
    def dns(self):
        d = self.d.get('DNS')
        return d if isinstance(d, dict) else {}
    @property
    def dhcp(self):
        d = self.d.get('DHCP')
        return d if isinstance(d, dict) else {}
    @property
    def nps(self):
        n = self.d.get('NPS')
        return n if isinstance(n, dict) else {}
    # FileShares
    @property
    def file_shares(self):
        fs = self.get('FileShares','Shares', default=[]) or []
        return [x for x in fs if isinstance(x, dict)]
    @property
    def file_share_sessions(self):
        return self.get('FileShares','OpenSessions', default=None)
    # Services / Tasks / Network / Printers / EventLog
    @property
    def services(self):
        s = self.d.get('Services', [])
        return [x for x in (s or []) if isinstance(x, dict)]
    @property
    def tasks(self):
        t = self.d.get('Tasks', [])
        return [x for x in (t or []) if isinstance(x, dict)]
    @property
    def network(self):
        n = self.d.get('Network')
        return n if isinstance(n, dict) else {}
    @property
    def printers(self):
        p = self.d.get('Printers', [])
        return [x for x in (p or []) if isinstance(x, dict)]
    @property
    def event_log(self):
        e = self.d.get('EventLog')
        return e if isinstance(e, dict) else {}
    @property
    def errors(self):
        e = self.d.get('Errors', '')
        return e if isinstance(e, str) else ''
    # IIS
    @property
    def iis(self):
        i = self.d.get('IIS')
        return i if isinstance(i, dict) else {}

def load_session(session_dir: Path):
    servers, hypervisors = [], []
    for jf in sorted(session_dir.glob('*.json')):
        raw = _read_json(jf)
        if not raw: continue
        data = normalize(raw)
        if isinstance(data, dict) and data.get('_type') in ('vSphereInventory','HyperVInventory'):
            hypervisors.append(data)
        elif isinstance(data, dict) and 'System' in data and 'Hardware' in data:
            servers.append(Server(data, jf))
    return servers, hypervisors

# ─── EOL lookups ───────────────────────────────────────────────────────
def _eol_color(date_str):
    if not date_str: return 'gray', ''
    try:
        d = datetime.date.fromisoformat(date_str[:10])
    except Exception:
        return 'gray', date_str
    days = (d - TODAY).days
    if days < 0:    return 'red',    f'EOL {date_str} ({-days}d ago)'
    if days < 365: return 'yellow', f'EOL {date_str} ({days}d)'
    return 'green', f'EOL {date_str}'

def os_eol(s: 'Server'):
    # Prefer collector-provided
    if s.os_eol_date:
        col, txt = _eol_color(s.os_eol_date)
        return s.os_eol_date, col, txt
    # Fallback
    hard = {'2008':'2020-01-14','2008 r2':'2020-01-14','2012':'2023-10-10','2012 r2':'2023-10-10',
            '2016':'2027-01-12','2019':'2029-01-09','2022':'2031-10-14','2025':'2034-10-10'}
    n = s.os_name.lower()
    for k, date in hard.items():
        if k in n:
            col, txt = _eol_color(date)
            return date, col, txt
    return None, 'gray', 'EOL unknown'

def hw_eol(mfr, model):
    if not mfr and not model: return None, 'gray', '—'
    key = f'{mfr} {model}'.lower()
    for entry in HARDWARE_EOL.get('models', []):
        if entry.get('match','').lower() in key:
            date = entry.get('eol_date')
            if not date: return None, 'gray', entry.get('note','virtual / n/a')
            col, txt = _eol_color(date)
            return date, col, txt
    return None, 'gray', 'unknown model'

# ─── Role guess ────────────────────────────────────────────────────────
def guess_role(s: Server) -> str:
    roles = set(s.role_names)
    feats = set(s.feature_names)
    tags = []
    if s.is_dc: tags.append('Domain Controller')
    if s.has_exchange: tags.append('Exchange')
    if 'DHCP' in roles: tags.append('DHCP')
    if 'DNS' in roles and not s.is_dc: tags.append('DNS')
    if 'FileAndStorage-Services' in roles and s.file_shares: tags.append('File Server')
    if s.sql_instances:
        n = len(s.sql_instances)
        tags.append(f'SQL × {n}' if n > 1 else 'SQL')
    if 'Remote-Desktop-Services' in roles: tags.append('RDS')
    if 'NPAS' in roles or s.nps.get('Installed'): tags.append('NPS/RADIUS')
    if 'Web-Server' in roles: tags.append('IIS')
    if 'Hyper-V' in roles: tags.append('Hyper-V Host')
    if 'Print-Services' in roles or s.printers: tags.append('Print')
    if not tags: tags.append('Application / Utility Server')
    return ' · '.join(tags)

# ─── Risk engine ───────────────────────────────────────────────────────
def collect_risks(servers, hypervisors):
    risks = []
    def add(sev, title, detail, server=''):
        risks.append({'severity':sev,'title':title,'detail':detail,'server':server})

    for hv in hypervisors:
        if hv.get('_type') == 'vSphereInventory':
            ver = hv.get('Version') or hv.get('APIVersion') or ''
            major = str(ver).strip('v').split('.')[0]
            if major in ('6','7'):
                add('critical', f'ESXi/vCenter {ver} — end-of-support',
                    'Past Broadcom general support. Active zero-days unpatched (CVE-2025-22224/22225/22226). Upgrade to ESXi 8.',
                    hv.get('Server',''))

    for s in servers:
        # OS EOL
        _, c, t = os_eol(s)
        if c == 'red':
            add('critical', f'{s.name} — {s.os_name}', f'OS past end-of-support. {t}', s.name)
        elif c == 'yellow':
            add('warning', f'{s.name} — {s.os_name}', f'OS end-of-support approaching. {t}', s.name)
        # Exchange
        if s.has_exchange:
            v = (s.exchange.get('VersionName') or s.exchange.get('Version') or '').lower()
            es = (s.exchange.get('EOLStatus') or '').lower()
            ed = s.exchange.get('EOLDate') or ''
            if 'eol' in es or 'past' in es or ('2019' in v and ed and ed < TODAY_STR):
                add('critical', f'{s.name} — Exchange past EOL',
                    f'Exchange {s.exchange.get("VersionName","")} ended support {ed}. Migrate to M365 or Exchange SE.', s.name)
            elif 'approach' in es:
                add('warning', f'{s.name} — Exchange EOL approaching', f'EOL {ed}', s.name)
        # SQL
        for inst in s.sql_instances:
            iv = (inst.get('Version') or '').lower()
            ed = (inst.get('EOLStatus') or '').lower()
            name = inst.get('InstanceName','')
            if 'past' in ed or 'eol' in ed:
                add('critical', f'{s.name}\\{name} — {inst.get("Edition","SQL")} past EOL',
                    f'SQL end-of-support {inst.get("EOLDate","")}. Upgrade required.', s.name)
            elif 'approach' in ed:
                add('warning', f'{s.name}\\{name} — SQL EOL approaching', f'EOL {inst.get("EOLDate","")}', s.name)
            # Compat level
            dbs = inst.get('Databases')
            if isinstance(dbs, dict): dbs = [dbs]
            dbs = dbs or []
            for db in dbs:
                if not isinstance(db, dict): continue
                cl = db.get('CompatLevel')
                if cl and cl <= 100:
                    age = (f' created {db.get("CreateDate","")}') if db.get('CreateDate') else ''
                    size = f' — {int(db.get("DataSizeMB") or 0):,} MB' if db.get('DataSizeMB') else ''
                    add('warning', f'{s.name}\\{name} — {db.get("Name","")} compat level {cl}',
                        f'Database at SQL 2008 compat on modern engine{age}{size}. Verify app vendor support before OS upgrade.', s.name)
        # Disk pressure
        for d in s.disks:
            try: pct = float(d.get('UsedPct') or 0)
            except: pct = 0
            if pct >= 85:
                add('warning', f'{s.name} — {d.get("Drive","")}: {pct:.0f}% full',
                    f'{d.get("Label","") or "drive"} {float(d.get("TotalGB",0)):.0f} GB total, {float(d.get("FreeGB",0)):.0f} GB free.', s.name)
        # SMB1
        if 'FS-SMB1' in s.feature_names:
            add('warning', f'{s.name} — SMB 1.0 enabled',
                'Legacy SMB1 protocol feature installed. Disable unless specific legacy dependency.', s.name)
        # DHCP on DC
        if s.is_dc and 'DHCP' in s.role_names:
            add('warning', f'{s.name} — DHCP collocated on DC',
                'DHCP runs on the domain controller. Must migrate scope during any DC replacement. Consider moving off.', s.name)
        # Stale users/computers
        ad = s.ad
        stale_u = ad.get('StaleUsers') or []
        if isinstance(stale_u, list) and len(stale_u) >= 5:
            add('warning', f'{s.name} — {len(stale_u)} stale AD users',
                'Users with no logon in 90+ days. Review for cleanup/deprovisioning.', s.name)
        stale_c = ad.get('StaleComputers') or []
        if isinstance(stale_c, list) and len(stale_c) >= 5:
            add('warning', f'{s.name} — {len(stale_c)} stale AD computers',
                'Computers with no logon in 90+ days. Review for cleanup.', s.name)
        # DHCP scope saturation
        for sc in (s.dhcp.get('Scopes') or []):
            if not isinstance(sc, dict): continue
            in_use = sc.get('InUse', 0) or 0
            avail = sc.get('Available', 0) or 0
            total = in_use + avail
            if total and (in_use / total) >= 0.85:
                add('warning', f'{s.name} — DHCP scope {sc.get("Name","")} {int(in_use/total*100)}% used',
                    f'{in_use} in use / {avail} free on {sc.get("ScopeId","")}', s.name)
        # Event log errors
        ec = s.event_log.get('ErrorCount', 0) or 0
        cc = s.event_log.get('CriticalCount', 0) or 0
        if cc >= 10 or ec >= 100:
            add('warning', f'{s.name} — {cc} critical / {ec} error events',
                'High event log error volume. Review before scoping changes.', s.name)
        # IIS Default Web Site on production
        if s.iis.get('Installed'):
            raw = (s.iis.get('Sites') or {}).get('Raw','') if isinstance(s.iis.get('Sites'), dict) else ''
            if isinstance(raw, str) and 'Default Web Site' in raw and 'Started' in raw:
                pass  # informational; don't alert

    # dedupe + sort
    out, seen = [], set()
    sev_rank = {'critical':0,'warning':1,'info':2}
    for r in sorted(risks, key=lambda x: (sev_rank.get(x['severity'],9), x['title'])):
        key = (r['severity'], r['title'])
        if key in seen: continue
        seen.add(key); out.append(r)
    return out

# ─── Per-server deep card ──────────────────────────────────────────────
def render_server_card(s: Server, risks_for_srv: list[dict]) -> str:
    # Header row
    _, osc, os_txt = os_eol(s)
    _, hwc, hw_txt = hw_eol(s.manufacturer, s.model)
    hw_pill = pill('virtual', 'gray') if s.is_vm else pill(hw_txt, hwc)
    role = guess_role(s)

    header = f'''
<div class="srv-head">
  <div>
    <div class="srv-title mono">{h(s.name)}</div>
    <div class="srv-sub">{h(s.domain)} · {h(role)}</div>
  </div>
  <div class="srv-meta">
    <div style="display:flex;gap:8px;justify-content:flex-end;">{pill(os_txt, osc)} {hw_pill}</div>
    <div style="margin-top:8px;">Uptime {f"{s.uptime_days:.1f}d" if isinstance(s.uptime_days,(int,float)) else "?"} · boot {h(s.last_boot)}</div>
    <div style="color:var(--dim);">Collected {h(s.get("Meta","CollectedAt",default=""))}</div>
  </div>
</div>'''

    # Identity / System
    identity = kv_table([
        ('OS', f'{h(s.os_name)} <span style="color:var(--muted);font-size:8.5pt;">build {h(s.os_build)}, v{h(s.os_version)}</span>'),
        ('Install date', h(s.os_install_date)),
        ('Domain', h(s.domain)),
        ('Timezone', h(s.timezone)),
        ('PowerShell', h(s.ps_version)),
        ('Run as', h(s.run_as)),
    ])

    # Hardware
    hardware = kv_table([
        ('Manufacturer', h(s.manufacturer)),
        ('Model', h(s.model)),
        ('CPU', f'{h(s.cpu_name)} <span style="color:var(--muted);">· {s.cpu_cores or "?"} cores</span>'),
        ('RAM', f'{float(s.ram_total or 0):.0f} GB total <span style="color:var(--muted);">· {float(s.ram_avail or 0):.1f} GB available</span>'),
        ('Serial', h(s.serial) if s.serial and s.serial != 'None' else '—'),
        ('BIOS', f'{h(s.bios_ver)} <span style="color:var(--muted);">· {h(s.bios_date)}</span>'),
        ('VM', f'{pill("yes","blue")} on {h(s.vm_platform)}' if s.is_vm else pill('physical','gray')),
    ])

    # Network
    n = s.network
    adapters = n.get('Adapters') or {}
    if isinstance(adapters, dict): adapters = [adapters]
    adapters = [a for a in adapters if isinstance(a, dict)]
    net_rows = []
    for a in adapters:
        net_rows.append((h(a.get('Description','')), f'{h(a.get("IPAddresses",""))} <span style="color:var(--muted);font-size:8.5pt">MAC {h(a.get("MAC",""))} · GW {h(a.get("Gateway",""))} · DNS {h(a.get("DNS",""))} · DHCP {"on" if a.get("DHCPEnabled") else "off"}</span>'))
    lp_count = len(n.get('ListeningPorts') or [])
    ec_count = len(n.get('EstablishedConns') or [])
    net_rows.append(('Ports', f'{lp_count} listening · {ec_count} established connections'))
    network = kv_table(net_rows)

    # Roles / Features
    role_badges = ''.join(f'<span class="role-chip">{h(r)}</span>' for r in s.role_displays) or '<em style="color:var(--muted);">No roles detected</em>'
    key_feats = [f for f in s.feature_names if f in {'FS-SMB1','Telnet-Client','SNMP-Service','Windows-Defender','Windows-Internal-Database','PowerShell-V2','NET-Framework-Features','RPC-over-HTTP-Proxy'}]
    feat_html = ', '.join(h(x) for x in key_feats) if key_feats else f'{len(s.features)} features installed'
    roles_html = f'<div>{role_badges}</div><div style="font-size:12px;color:var(--muted);margin-top:8px;"><strong style="color:var(--text);">Notable features:</strong> {feat_html}</div>'

    # Disks
    disk_rows = ''
    if s.disks:
        disk_rows = '<table style="width:100%;border-collapse:collapse;font-size:9pt;"><thead><tr style="background:var(--elevated);"><th style="text-align:left;padding:5px 10px;font-size:8pt;color:var(--muted);">Drive</th><th style="text-align:left;padding:5px 10px;font-size:8pt;color:var(--muted);">Label</th><th style="text-align:right;padding:5px 10px;font-size:8pt;color:var(--muted);">Total</th><th style="text-align:right;padding:5px 10px;font-size:8pt;color:var(--muted);">Used</th><th style="text-align:right;padding:5px 10px;font-size:8pt;color:var(--muted);">Free</th><th style="padding:5px 10px;font-size:8pt;color:var(--muted);width:160px;">Usage</th></tr></thead><tbody>'
        for d in s.disks:
            pct = float(d.get('UsedPct') or 0)
            tot = float(d.get('TotalGB') or 0)
            fre = float(d.get('FreeGB') or 0)
            used = tot - fre
            pc = 'red' if pct >= 85 else ('yellow' if pct >= 70 else 'green')
            disk_rows += f'<tr><td style="padding:5px 10px;font-family:monospace;font-weight:600;">{h(d.get("Drive",""))}</td><td style="padding:5px 10px;font-size:9pt;color:var(--muted);">{h(d.get("Label",""))} <span style="font-size:8pt">{h(d.get("Filesystem",""))}</span></td><td style="padding:5px 10px;text-align:right;">{tot:.0f} GB</td><td style="padding:5px 10px;text-align:right;">{used:.0f} GB</td><td style="padding:5px 10px;text-align:right;">{fre:.0f} GB</td><td style="padding:5px 10px;">{pill(f"{pct:.0f}%", pc)}{bar(pct)}</td></tr>'
        disk_rows += '</tbody></table>'

    # SQL block
    sql_html = ''
    if s.sql_instances:
        sql_html = ''
        for inst in s.sql_instances:
            eol_col = 'red' if 'past' in (inst.get('EOLStatus','').lower()) else ('yellow' if 'approach' in (inst.get('EOLStatus','').lower()) else 'green')
            header_sql = f'<div style="margin-top:10px;font-weight:700;">{h(inst.get("InstanceName",""))} <span style="color:var(--muted);font-weight:400;font-size:9pt;">· {h(inst.get("Edition",""))} · {h(inst.get("Version",""))}</span> {pill(inst.get("EOLStatus","")+" "+(inst.get("EOLDate","") or ""), eol_col)}</div><div style="font-size:8.5pt;color:var(--muted);margin-bottom:4px;">Service account: {h(inst.get("ServiceAccount",""))}</div>'
            dbs = inst.get('Databases')
            if isinstance(dbs, dict): dbs = [dbs]
            dbs = dbs or []
            db_rows = ''
            for db in dbs:
                if not isinstance(db, dict): continue
                cl = db.get('CompatLevel')
                cl_col = 'red' if cl and cl <= 100 else ('yellow' if cl and cl <= 110 else 'green')
                db_rows += f'<tr><td style="padding:4px 10px;font-weight:600;">{h(db.get("Name",""))}</td><td style="padding:4px 10px;">{pill(f"compat {cl}", cl_col)}</td><td style="padding:4px 10px;text-align:right;">{int(db.get("DataSizeMB") or 0):,} MB</td><td style="padding:4px 10px;text-align:right;">{int(db.get("LogSizeMB") or 0):,} MB log</td><td style="padding:4px 10px;">{h(db.get("RecoveryModel",""))}</td><td style="padding:4px 10px;font-size:8.5pt;color:var(--muted);">created {h(db.get("CreateDate",""))} · last backup {h(db.get("LastFullBackup",""))}</td></tr>'
            table = f'<table style="width:100%;border-collapse:collapse;font-size:9pt;"><thead><tr style="background:var(--elevated);"><th style="text-align:left;padding:4px 10px;font-size:8pt;color:var(--muted);">Database</th><th style="padding:4px 10px;font-size:8pt;color:var(--muted);">Compat</th><th style="padding:4px 10px;font-size:8pt;color:var(--muted);">Size</th><th style="padding:4px 10px;font-size:8pt;color:var(--muted);">Log</th><th style="padding:4px 10px;font-size:8pt;color:var(--muted);">Recovery</th><th style="padding:4px 10px;font-size:8pt;color:var(--muted);">Detail</th></tr></thead><tbody>{db_rows}</tbody></table>' if db_rows else '<em style="color:var(--muted);font-size:9pt;">No databases reported.</em>'
            sql_html += header_sql + table

    # Exchange
    exch_html = ''
    if s.has_exchange:
        e = s.exchange
        ec = 'red' if 'past' in (e.get('EOLStatus','').lower()) else ('yellow' if 'approach' in (e.get('EOLStatus','').lower()) else 'green')
        exch_html = kv_table([
            ('Version', f'{h(e.get("VersionName",""))} <span style="color:var(--muted);">({h(e.get("Version",""))})</span>'),
            ('EOL', pill(f'{e.get("EOLStatus","")} {e.get("EOLDate","")}'.strip(), ec)),
            ('Mailboxes', str(e.get('MailboxCount','?'))),
            ('Transport', 'running' if e.get('TransportServiceRunning') else 'stopped'),
        ])

    # IIS
    iis_html = ''
    if s.iis.get('Installed'):
        raw = (s.iis.get('Sites') or {}).get('Raw','') if isinstance(s.iis.get('Sites'), dict) else ''
        if isinstance(raw, str):
            raw_short = '<pre style="white-space:pre-wrap;font-size:8.5pt;background:var(--elevated);padding:8px;border-radius:4px;color:var(--text);overflow:auto;max-height:200px;">'+h(raw[:2000]) + ('\n\n… (truncated)' if len(raw) > 2000 else '') +'</pre>'
            iis_html = raw_short

    # AD (only if DC)
    ad_html = ''
    if s.is_dc and s.ad:
        ad = s.ad
        stale_u = ad.get('StaleUsers') or []
        stale_c = ad.get('StaleComputers') or []
        rows = [
            ('Forest', h(ad.get('ForestName', ad.get('Forest','')))),
            ('Domain', h(ad.get('DomainName', ad.get('Domain','')))),
            ('Functional levels', f'Forest {h(ad.get("ForestFL","?"))} · Domain {h(ad.get("DomainFL","?"))}'),
            ('DCs', str(ad.get('DCCount','?'))),
            ('Users / Computers / OUs', f'{ad.get("UserCount","?")} users · {ad.get("ComputerCount","?")} computers · {ad.get("OUCount","?")} OUs'),
            ('PDC Emulator', h(ad.get('PDCEmulator',''))),
            ('RID Master', h(ad.get('RIDMaster',''))),
            ('Stale users (90+ days)', f'{len(stale_u)}' + (' — see detail below' if stale_u else '')),
            ('Stale computers (90+ days)', f'{len(stale_c)}' + (' — see detail below' if stale_c else '')),
        ]
        ad_html = kv_table(rows)
        if stale_u:
            rows_html = ''
            for u in stale_u[:15]:
                if not isinstance(u, dict): continue
                rows_html += f'<tr><td style="padding:3px 10px;">{h(u.get("Name",""))}</td><td style="padding:3px 10px;font-family:monospace;font-size:8.5pt;">{h(u.get("SamAccountName",""))}</td><td style="padding:3px 10px;font-size:8.5pt;color:var(--muted);">{h(u.get("LastLogon",""))}</td></tr>'
            ad_html += f'<div style="margin-top:10px;font-size:9pt;color:var(--accent);font-weight:600;">Stale users (first 15 of {len(stale_u)}):</div><table style="width:100%;border-collapse:collapse;font-size:9pt;"><thead><tr style="background:var(--elevated);"><th style="text-align:left;padding:3px 10px;font-size:8pt;color:var(--muted);">Display name</th><th style="text-align:left;padding:3px 10px;font-size:8pt;color:var(--muted);">SAM</th><th style="text-align:left;padding:3px 10px;font-size:8pt;color:var(--muted);">Last logon</th></tr></thead><tbody>{rows_html}</tbody></table>'

    # DNS
    dns_html = ''
    if s.dns.get('Installed'):
        fwds = s.dns.get('Forwarders') or []
        fwd_ips = []
        for f in fwds:
            if isinstance(f, dict):
                fwd_ips.append(f.get('IPAddressToString') or f.get('Address',''))
            else:
                fwd_ips.append(str(f))
        zones = s.dns.get('Zones') or []
        zone_rows = ''
        for z in zones:
            if not isinstance(z, dict): continue
            zone_rows += f'<tr><td style="padding:3px 10px;font-family:monospace;">{h(z.get("ZoneName",""))}</td><td style="padding:3px 10px;">{h(z.get("ZoneType",""))}</td><td style="padding:3px 10px;">{"AD-integrated" if z.get("IsDsIntegrated") else ""}</td><td style="padding:3px 10px;">{"reverse" if z.get("IsReverseLookupZone") else "forward"}</td></tr>'
        dns_html = kv_table([('Forwarders', ', '.join(h(x) for x in fwd_ips) or '—')])
        if zone_rows:
            dns_html += f'<div style="margin-top:8px;font-size:9pt;color:var(--accent);font-weight:600;">Zones ({len(zones)}):</div><table style="width:100%;border-collapse:collapse;font-size:9pt;"><thead><tr style="background:var(--elevated);"><th style="text-align:left;padding:3px 10px;font-size:8pt;color:var(--muted);">Zone</th><th style="text-align:left;padding:3px 10px;font-size:8pt;color:var(--muted);">Type</th><th style="text-align:left;padding:3px 10px;font-size:8pt;color:var(--muted);">Integration</th><th style="text-align:left;padding:3px 10px;font-size:8pt;color:var(--muted);">Direction</th></tr></thead><tbody>{zone_rows}</tbody></table>'

    # DHCP
    dhcp_html = ''
    if s.dhcp.get('Installed'):
        scopes = s.dhcp.get('Scopes') or []
        rows = ''
        for sc in scopes:
            if not isinstance(sc, dict): continue
            in_use = sc.get('InUse', 0) or 0
            avail = sc.get('Available', 0) or 0
            total = in_use + avail
            pct = (in_use/total*100) if total else 0
            rows += f'<tr><td style="padding:4px 10px;font-weight:600;">{h(sc.get("Name",""))}</td><td style="padding:4px 10px;font-family:monospace;font-size:8.5pt;">{h(sc.get("ScopeId",""))}</td><td style="padding:4px 10px;font-size:8.5pt;color:var(--muted);">{h(sc.get("StartRange",""))} – {h(sc.get("EndRange",""))}</td><td style="padding:4px 10px;">{pill(sc.get("State",""), "green" if sc.get("State")=="Active" else "gray")}</td><td style="padding:4px 10px;text-align:right;">{in_use} / {avail}</td><td style="padding:4px 10px;width:140px;">{pill(f"{pct:.0f}%", "red" if pct>=85 else "yellow" if pct>=70 else "green")}{bar(pct)}</td></tr>'
        dhcp_html = f'<table style="width:100%;border-collapse:collapse;font-size:9pt;"><thead><tr style="background:var(--elevated);"><th style="text-align:left;padding:4px 10px;font-size:8pt;color:var(--muted);">Name</th><th style="text-align:left;padding:4px 10px;font-size:8pt;color:var(--muted);">Subnet</th><th style="text-align:left;padding:4px 10px;font-size:8pt;color:var(--muted);">Range</th><th style="padding:4px 10px;font-size:8pt;color:var(--muted);">State</th><th style="padding:4px 10px;font-size:8pt;color:var(--muted);">In use / Free</th><th style="padding:4px 10px;font-size:8pt;color:var(--muted);">Utilization</th></tr></thead><tbody>{rows}</tbody></table>' if rows else '<em style="color:var(--muted);">DHCP installed, no scopes.</em>'

    # NPS
    nps_html = ''
    if s.nps.get('Installed'):
        clients = s.nps.get('Clients') or []
        policies = s.nps.get('Policies') or []
        nps_html = kv_table([
            ('Clients (RADIUS devices)', f'{len(clients)}' if isinstance(clients, list) else '?'),
            ('Network policies', f'{len(policies)}' if isinstance(policies, list) else '?'),
        ])
        if isinstance(clients, list) and clients:
            cli_rows = ''
            for c in clients[:10]:
                if isinstance(c, dict):
                    cli_rows += f'<tr><td style="padding:3px 10px;">{h(c.get("Name",""))}</td><td style="padding:3px 10px;font-family:monospace;">{h(c.get("Address",""))}</td><td style="padding:3px 10px;">{h(c.get("VendorName",""))}</td></tr>'
            if cli_rows:
                nps_html += f'<div style="margin-top:8px;font-size:9pt;color:var(--accent);font-weight:600;">RADIUS clients:</div><table style="width:100%;border-collapse:collapse;font-size:9pt;"><tbody>{cli_rows}</tbody></table>'

    # File Shares
    fs_html = ''
    if s.file_shares:
        rows = ''
        for fs in s.file_shares[:30]:
            perms = fs.get('Permissions') or []
            pc = len(perms) if isinstance(perms, list) else '?'
            rows += f'<tr><td style="padding:4px 10px;font-weight:600;">{h(fs.get("Name",""))}</td><td style="padding:4px 10px;font-family:monospace;font-size:8.5pt;">{h(fs.get("Path",""))}</td><td style="padding:4px 10px;font-size:8.5pt;color:var(--muted);">{h(fs.get("Description","") or "—")}</td><td style="padding:4px 10px;text-align:right;">{pc} ACLs</td></tr>'
        fs_html = f'<div style="font-size:8.5pt;color:var(--muted);margin-bottom:4px;">Open sessions: {s.file_share_sessions or 0}</div><table style="width:100%;border-collapse:collapse;font-size:9pt;"><thead><tr style="background:var(--elevated);"><th style="text-align:left;padding:4px 10px;font-size:8pt;color:var(--muted);">Share</th><th style="text-align:left;padding:4px 10px;font-size:8pt;color:var(--muted);">Path</th><th style="text-align:left;padding:4px 10px;font-size:8pt;color:var(--muted);">Description</th><th style="padding:4px 10px;font-size:8pt;color:var(--muted);">Permissions</th></tr></thead><tbody>{rows}</tbody></table>'
        if len(s.file_shares) > 30:
            fs_html += f'<div style="font-size:8.5pt;color:var(--muted);margin-top:4px;">… {len(s.file_shares)-30} more</div>'

    # Services summary
    svc_html = ''
    if s.services:
        # Top interesting: not MS, and not stopped
        non_ms_running = [x for x in s.services if not x.get('IsMS', True) and x.get('State') in ('Running','Auto')]
        stopped_auto = [x for x in s.services if x.get('StartMode') in ('Auto','Automatic') and x.get('State') == 'Stopped']
        rows = ''
        for x in non_ms_running[:20]:
            rows += f'<tr><td style="padding:3px 10px;font-weight:600;">{h(x.get("DisplayName",""))}</td><td style="padding:3px 10px;font-size:8.5pt;color:var(--muted);font-family:monospace;">{h(x.get("Name",""))}</td><td style="padding:3px 10px;">{pill(x.get("State",""), "green" if x.get("State") in ("Running","Auto") else "gray")}</td><td style="padding:3px 10px;font-size:8.5pt;">{h(x.get("StartName","") or "LocalSystem")}</td></tr>'
        svc_html = f'<div style="font-size:9pt;color:var(--muted);margin-bottom:4px;">{len(s.services)} services total · {len(non_ms_running)} non-Microsoft running · {len(stopped_auto)} auto-start stopped</div>'
        if rows:
            svc_html += f'<table style="width:100%;border-collapse:collapse;font-size:9pt;"><thead><tr style="background:var(--elevated);"><th style="text-align:left;padding:3px 10px;font-size:8pt;color:var(--muted);">Service</th><th style="text-align:left;padding:3px 10px;font-size:8pt;color:var(--muted);">Name</th><th style="padding:3px 10px;font-size:8pt;color:var(--muted);">State</th><th style="text-align:left;padding:3px 10px;font-size:8pt;color:var(--muted);">Account</th></tr></thead><tbody>{rows}</tbody></table>'
        if stopped_auto:
            svc_html += f'<div style="margin-top:8px;font-size:9pt;color:#fcd34d;"><strong>Auto-start services currently stopped:</strong> ' + ', '.join(h(x.get('Name','')) for x in stopped_auto[:10]) + '</div>'

    # Scheduled tasks
    task_html = ''
    if s.tasks:
        non_ms = [t for t in s.tasks if not (t.get('Path','').startswith('\\Microsoft'))]
        rows = ''
        for t in non_ms[:15]:
            rows += f'<tr><td style="padding:3px 10px;font-weight:600;">{h(t.get("Name",""))}</td><td style="padding:3px 10px;font-family:monospace;font-size:8.5pt;">{h(t.get("Path","")) }</td><td style="padding:3px 10px;">{pill(t.get("State",""), "green" if t.get("State")=="Ready" else "gray")}</td><td style="padding:3px 10px;font-size:8.5pt;color:var(--muted);">last {h(t.get("LastRun",""))} · next {h(t.get("NextRun",""))}</td></tr>'
        task_html = f'<div style="font-size:9pt;color:var(--muted);margin-bottom:4px;">{len(s.tasks)} tasks total · {len(non_ms)} non-Microsoft</div>'
        if rows:
            task_html += f'<table style="width:100%;border-collapse:collapse;font-size:9pt;"><thead><tr style="background:var(--elevated);"><th style="text-align:left;padding:3px 10px;font-size:8pt;color:var(--muted);">Task</th><th style="text-align:left;padding:3px 10px;font-size:8pt;color:var(--muted);">Path</th><th style="padding:3px 10px;font-size:8pt;color:var(--muted);">State</th><th style="text-align:left;padding:3px 10px;font-size:8pt;color:var(--muted);">Runs</th></tr></thead><tbody>{rows}</tbody></table>'

    # Event log
    el = s.event_log
    ev_html = ''
    if el:
        ev_html = kv_table([
            ('Critical events', str(el.get('CriticalCount', 0))),
            ('Error events', str(el.get('ErrorCount', 0))),
        ])
        rc = el.get('RecentCritical') or []
        if isinstance(rc, list) and rc:
            rows = ''
            for ev in rc[:5]:
                if isinstance(ev, dict):
                    rows += f'<tr><td style="padding:3px 10px;font-size:8.5pt;">{h(ev.get("TimeCreated",""))}</td><td style="padding:3px 10px;">{h(ev.get("ProviderName",""))}</td><td style="padding:3px 10px;font-size:8.5pt;color:var(--muted);">{h((ev.get("Message","") or "")[:200])}</td></tr>'
            if rows:
                ev_html += f'<table style="width:100%;border-collapse:collapse;font-size:9pt;margin-top:8px;"><tbody>{rows}</tbody></table>'

    # Risks for this server
    risk_html = ''
    if risks_for_srv:
        items = ''
        for r in risks_for_srv:
            icon = '🔴' if r['severity']=='critical' else '🟡'
            items += f'<li style="margin:5px 0;"><strong>{icon} {h(r["title"])}</strong> — <span style="color:var(--muted);">{h(r["detail"])}</span></li>'
        risk_html = f'<ul style="font-size:9.5pt;padding-left:20px;margin:0;">{items}</ul>'
    else:
        risk_html = '<div style="font-size:9pt;color:var(--success);">✓ No risks detected for this server.</div>'

    # Collection errors
    err_html = ''
    if s.errors:
        err_html = f'<div style="background:rgba(245,158,11,0.12);border-left:4px solid #d97706;padding:10px 14px;margin-top:12px;font-size:9pt;color:#fcd34d;"><strong>Collection errors:</strong><pre style="white-space:pre-wrap;font-size:8.5pt;margin:6px 0 0;">{h(s.errors[:800])}</pre></div>'

    # Assemble — two-column layout
    def sec(title, body):
        if not body: return ''
        return section_hdr(title) + body

    left_col = sec('Identity', identity) + sec('Hardware', hardware) + sec('Network', network) + sec('Roles', roles_html)
    right_col = sec('Findings', risk_html) + sec('Storage', disk_rows) + sec('Event Log', ev_html)

    # Full-width sections (role-specific data)
    full = ''
    full += sec('Active Directory', ad_html)
    full += sec('DNS', dns_html)
    full += sec('DHCP', dhcp_html)
    full += sec('NPS (RADIUS)', nps_html)
    full += sec('SQL Server', sql_html)
    full += sec('Exchange', exch_html)
    full += sec('IIS', iis_html)
    full += sec('File Shares', fs_html)
    full += sec('Services', svc_html)
    full += sec('Scheduled Tasks', task_html)

    body = (header +
            f'<div style="display:grid;grid-template-columns:1fr 1fr;gap:28px;margin-top:14px;"><div>{left_col}</div><div>{right_col}</div></div>' +
            full + err_html)

    return f'<div id="srv-{h(s.name)}">{card(body)}</div>'

# ─── Summary tab ───────────────────────────────────────────────────────
def render_summary(client, servers, hypervisors, risks):
    total_vcpu = sum(int(s.cpu_cores or 0) for s in servers)
    total_ram  = sum(float(s.ram_total or 0) for s in servers)
    total_disk = sum(sum(float(d.get('TotalGB',0) or 0) for d in s.disks) for s in servers)
    total_used = sum(sum((float(d.get('TotalGB',0) or 0) - float(d.get('FreeGB',0) or 0)) for d in s.disks) for s in servers)
    crit = sum(1 for r in risks if r['severity']=='critical')
    warn = sum(1 for r in risks if r['severity']=='warning')
    hv_count = len(hypervisors)
    vm_count = sum(len([v for v in (h.get('VMs',[]) or []) if isinstance(v, dict)]) for h in hypervisors)

    disk_pct = (total_used/total_disk*100) if total_disk else 0
    hero = f'''
<div class="bento">
  <div class="hero">
    <div class="hero-eyebrow">Server Discovery Report</div>
    <div class="hero-title">{h(client)}</div>
    <div class="hero-meta">
      {len(servers)} server{"s" if len(servers)!=1 else ""} · {hv_count} hypervisor{"s" if hv_count!=1 else ""} · {vm_count} VMs discovered · collected {TODAY_STR}
    </div>
    <div class="hero-badges">
      {pill(f'{crit} critical', 'crit')}
      {pill(f'{warn} warnings', 'warn')}
      {pill('Live data', 'info')}
    </div>
  </div>
  <div class="kpi"><div class="kpi-label">Servers</div><div class="kpi-num">{len(servers)}</div><div class="kpi-detail">{sum(1 for s in servers if s.is_vm)} virtual · {sum(1 for s in servers if not s.is_vm)} physical</div></div>
  <div class="kpi"><div class="kpi-label">Compute</div><div class="kpi-num">{total_vcpu}<span class="kpi-unit">vCPU</span></div><div class="kpi-detail">{total_ram:.0f} GB RAM allocated</div></div>
  <div class="kpi"><div class="kpi-label">Storage</div><div class="kpi-num">{total_used:.0f}<span class="kpi-unit">GB</span></div><div class="kpi-detail">of {total_disk:.0f} GB ({disk_pct:.0f}% used)</div></div>
  <div class="kpi {'crit' if crit else 'warn' if warn else ''}"><div class="kpi-label">Findings</div><div class="kpi-num">{crit + warn}</div><div class="kpi-detail">{crit} critical · {warn} warnings</div></div>
</div>'''

    # Inventory table
    rows = []
    for s in servers:
        _, osc, os_txt = os_eol(s)
        _, hwc, hw_txt = hw_eol(s.manufacturer, s.model)
        sc = sum(1 for r in risks if r['server']==s.name and r['severity']=='critical')
        sw = sum(1 for r in risks if r['server']==s.name and r['severity']=='warning')
        flags = []
        if sc: flags.append(pill(f'{sc} crit','crit'))
        if sw: flags.append(pill(f'{sw} warn','warn'))
        if not flags: flags.append(pill('clean','ok'))
        os_short = re.sub(r'Microsoft Windows Server\s+','WS',s.os_name)
        os_short = re.sub(r'\s+Standard|\s+Datacenter|\s+Essentials','',os_short)
        disk_tot = sum(float(d.get('TotalGB',0) or 0) for d in s.disks)
        rows.append(f'''<tr>
<td><a class="mono" href="#srv-{h(s.name)}" onclick="setTab(\'servers\')">{h(s.name)}</a></td>
<td style="color:var(--muted);">{h(guess_role(s))}</td>
<td>{h(os_short)}</td>
<td>{pill(os_txt, osc)}</td>
<td class="mono" style="color:var(--muted);">{s.cpu_cores or "?"}c · {float(s.ram_total or 0):.0f}G · {disk_tot:.0f}G</td>
<td>{pill("virtual","neutral") if s.is_vm else pill(hw_txt, hwc)}</td>
<td style="text-align:right;">{" ".join(flags)}</td>
</tr>''')
    inv_table = card(f'''<div class="card-title">Server Inventory</div>
<div class="card-sub" style="margin-bottom:14px;">Roles auto-classified from installed Windows roles + SQL/Exchange detection. Click a hostname to jump to full detail.</div>
<table class="inv">
<thead><tr>
<th>Host</th><th>Role</th><th>OS</th><th>OS EOL</th><th>Sizing</th><th>HW EOL</th><th style="text-align:right;">Status</th>
</tr></thead><tbody>{"".join(rows)}</tbody></table>''')

    # Hypervisor cards
    hv_cards = ''
    for hv in hypervisors:
        if hv.get('_type') != 'vSphereInventory': continue
        esx_hosts = hv.get('ESXHosts') or []
        vms = [v for v in (hv.get('VMs') or []) if isinstance(v, dict)]
        ver = hv.get('Version') or hv.get('APIVersion','')
        running = sum(1 for v in vms if v.get('PowerState') in ('POWERED_ON','Running'))
        vcpu = sum(int(v.get('vCPU') or 0) for v in vms)
        ram = sum(float(v.get('RAMgb') or 0) for v in vms)
        disk = sum(sum(float(d.get('CapacityGB') or 0) for d in (v.get('Disks') or []) if isinstance(d, dict)) for v in vms)
        major = str(ver).strip('v').split('.')[0]
        ver_pill = pill(f'ESXi {ver} past EOL','crit') if major in ('6','7') else pill(f'ESXi {ver}','ok')

        ds_rows = ''
        for d in (hv.get('Datastores') or []):
            if not isinstance(d, dict): continue
            cap = float(d.get('CapacityGB') or 0)
            free = float(d.get('FreeGB') or 0)
            used = cap - free
            pct = (used/cap*100) if cap else 0
            pc = 'crit' if pct >= 85 else ('warn' if pct >= 70 else 'ok')
            ds_rows += f'<tr><td style="font-weight:600;">{h(d.get("Name",""))}</td><td style="color:var(--muted);font-size:11px;">{h(d.get("Type",""))}</td><td style="text-align:right;" class="mono">{cap:.0f} GB</td><td style="text-align:right;" class="mono">{used:.0f} GB</td><td style="text-align:right;" class="mono">{free:.0f} GB</td><td style="width:180px;">{pill(f"{pct:.0f}%",pc)}<div style="margin-top:4px;">{bar(pct)}</div></td></tr>'

        host_name = ''
        if isinstance(esx_hosts, list) and esx_hosts and isinstance(esx_hosts[0], dict):
            host_name = esx_hosts[0].get('Name','') or ''
        if not host_name:
            host_name = hv.get('Server','')
        hv_cards += card(f'''
<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:20px;">
  <div>
    <div class="card-title">Hypervisor</div>
    <div class="mono" style="font-size:16px;color:var(--text);margin-top:6px;font-weight:600;">{h(host_name)}</div>
    <div style="margin-top:8px;display:flex;gap:8px;align-items:center;">
      {ver_pill} <span class="sub" style="font-size:11px;">vCenter at <span class="mono">{h(hv.get("Server",""))}</span></span>
    </div>
  </div>
  <div style="text-align:right;">
    <div class="kpi-label">Virtual Machines</div>
    <div class="kpi-num" style="font-size:32px;">{len(vms)} <span class="kpi-unit">· {running} running</span></div>
    <div class="kpi-detail">{vcpu} vCPU · {ram:.0f} GB RAM · {disk:.0f} GB allocated</div>
  </div>
</div>
<div class="callout"><strong>Note:</strong> vSphere REST API doesn\'t expose ESXi host hardware (physical CPU, RAM, local storage). Run SOAP collector for full host inventory.</div>
<div class="section-hdr">Datastores</div>
<table class="dt"><thead><tr><th>Name</th><th>Type</th><th style="text-align:right;">Capacity</th><th style="text-align:right;">Used</th><th style="text-align:right;">Free</th><th>Utilization</th></tr></thead><tbody>{ds_rows}</tbody></table>''')

    # Top risks
    top = ''
    for i, r in enumerate(risks[:10]):
        sev_cls = 'crit' if r['severity']=='critical' else 'warn'
        top += f'''<div class="risk-item">
<div style="display:flex;justify-content:space-between;gap:10px;align-items:flex-start;">
<div><span class="pill {sev_cls}" style="margin-right:8px;"><span class="pill-dot"></span>{r["severity"]}</span><span class="risk-title">{h(r["title"])}</span></div>
<span class="risk-num">#{i+1:02d}</span>
</div>
<div class="risk-detail">{h(r["detail"])}</div>
</div>'''
    risks_card = card(f'''<div style="display:flex;justify-content:space-between;align-items:flex-start;">
<div><div class="card-title">Top Risks</div><div class="card-sub">Auto-prioritized. Full list on <a href="#" onclick="setTab(\'risks\');return false;">Risks tab →</a></div></div>
<span class="pill neutral"><span class="pill-dot"></span>{len(risks)} total</span>
</div>
<div class="risk-strip" style="margin-top:12px;">{top or "<div style='padding:14px;color:var(--success);'>✓ No risks detected.</div>"}</div>''')

    return hero + inv_table + hv_cards + risks_card

# ─── Servers tab ───────────────────────────────────────────────────────
def render_servers(servers, risks):
    parts = []
    for s in servers:
        sr = [r for r in risks if r['server']==s.name]
        parts.append(render_server_card(s, sr))
    return ''.join(parts) or '<div style="padding:40px;text-align:center;color:var(--muted);">No server discovery data.</div>'

# ─── Environment tab ───────────────────────────────────────────────────
def render_environment(servers, hypervisors):
    blocks = []

    # AD consolidated (first DC with data wins)
    for s in servers:
        if s.is_dc and s.ad.get('DomainName'):
            ad = s.ad
            stale_u = ad.get('StaleUsers') or []
            blocks.append(card(f'''<div style="font-size:13pt;font-weight:700;color:var(--text);margin-bottom:10px;">Active Directory</div>
{kv_table([
    ('Forest', h(ad.get("ForestName","") or ad.get("Forest",""))),
    ('Domain', h(ad.get("DomainName","") or ad.get("Domain",""))),
    ('Functional levels', f'Forest {h(ad.get("ForestFL","?"))} · Domain {h(ad.get("DomainFL","?"))}'),
    ('DCs', str(ad.get("DCCount","?"))),
    ('Users', str(ad.get("UserCount","?"))),
    ('Computers', str(ad.get("ComputerCount","?"))),
    ('OUs', str(ad.get("OUCount","?"))),
    ('PDC Emulator', h(ad.get("PDCEmulator",""))),
    ('RID Master', h(ad.get("RIDMaster",""))),
    ('Stale users (90+d)', f'{len(stale_u)}'),
])}'''))
            break

    # DHCP consolidated
    all_scopes = []
    for s in servers:
        for sc in (s.dhcp.get('Scopes') or []):
            if isinstance(sc, dict):
                all_scopes.append((s.name, sc))
    if all_scopes:
        rows = ''
        for host, sc in all_scopes:
            in_use = sc.get('InUse',0) or 0
            avail = sc.get('Available',0) or 0
            total = in_use + avail
            pct = (in_use/total*100) if total else 0
            rows += f'<tr><td style="padding:5px 10px;">{h(host)}</td><td style="padding:5px 10px;font-weight:600;">{h(sc.get("Name",""))}</td><td style="padding:5px 10px;font-family:monospace;font-size:8.5pt;">{h(sc.get("ScopeId",""))}</td><td style="padding:5px 10px;font-size:8.5pt;">{h(sc.get("StartRange",""))} – {h(sc.get("EndRange",""))}</td><td style="padding:5px 10px;text-align:right;">{in_use} / {avail}</td><td style="padding:5px 10px;width:140px;">{pill(f"{pct:.0f}%","red" if pct>=85 else "yellow" if pct>=70 else "green")}{bar(pct)}</td></tr>'
        blocks.append(card(f'''<div style="font-size:13pt;font-weight:700;color:var(--text);margin-bottom:10px;">DHCP Scopes (all DCs)</div>
<table style="width:100%;border-collapse:collapse;font-size:9pt;">
<thead><tr style="background:var(--elevated);">
<th style="text-align:left;padding:5px 10px;font-size:8pt;color:var(--muted);">Host</th>
<th style="text-align:left;padding:5px 10px;font-size:8pt;color:var(--muted);">Name</th>
<th style="text-align:left;padding:5px 10px;font-size:8pt;color:var(--muted);">Subnet</th>
<th style="text-align:left;padding:5px 10px;font-size:8pt;color:var(--muted);">Range</th>
<th style="padding:5px 10px;font-size:8pt;color:var(--muted);">In use / Free</th>
<th style="padding:5px 10px;font-size:8pt;color:var(--muted);">Utilization</th>
</tr></thead><tbody>{rows}</tbody></table>'''))

    # DNS zones consolidated
    all_zones = []
    all_fwds = []
    for s in servers:
        if s.dns.get('Installed'):
            for z in (s.dns.get('Zones') or []):
                if isinstance(z, dict):
                    all_zones.append((s.name, z))
            for f in (s.dns.get('Forwarders') or []):
                if isinstance(f, dict):
                    all_fwds.append(f.get('IPAddressToString') or f.get('Address',''))
    if all_zones:
        rows = ''
        for host, z in all_zones:
            rows += f'<tr><td style="padding:4px 10px;">{h(host)}</td><td style="padding:4px 10px;font-family:monospace;">{h(z.get("ZoneName",""))}</td><td style="padding:4px 10px;">{h(z.get("ZoneType",""))}</td><td style="padding:4px 10px;">{"AD-integrated" if z.get("IsDsIntegrated") else "file"}</td><td style="padding:4px 10px;">{"reverse" if z.get("IsReverseLookupZone") else "forward"}</td></tr>'
        blocks.append(card(f'''<div style="font-size:13pt;font-weight:700;color:var(--text);margin-bottom:6px;">DNS</div>
<div style="font-size:9pt;color:var(--muted);margin-bottom:10px;"><strong>Forwarders:</strong> {", ".join(h(x) for x in sorted(set(all_fwds))) or "—"}</div>
<table style="width:100%;border-collapse:collapse;font-size:9pt;">
<thead><tr style="background:var(--elevated);">
<th style="text-align:left;padding:4px 10px;font-size:8pt;color:var(--muted);">Host</th>
<th style="text-align:left;padding:4px 10px;font-size:8pt;color:var(--muted);">Zone</th>
<th style="text-align:left;padding:4px 10px;font-size:8pt;color:var(--muted);">Type</th>
<th style="text-align:left;padding:4px 10px;font-size:8pt;color:var(--muted);">Integration</th>
<th style="text-align:left;padding:4px 10px;font-size:8pt;color:var(--muted);">Direction</th>
</tr></thead><tbody>{rows}</tbody></table>'''))

    # Hypervisor VM inventory (full)
    for hv in hypervisors:
        if hv.get('_type') != 'vSphereInventory': continue
        vms = [v for v in (hv.get('VMs') or []) if isinstance(v, dict)]
        rows = ''
        for v in vms:
            disks = [d for d in (v.get('Disks') or []) if isinstance(d, dict)]
            disk_tot = sum(float(d.get('CapacityGB') or 0) for d in disks)
            state = v.get('PowerState','')
            rows += f'''<tr>
<td style="padding:5px 10px;font-weight:600;">{h(v.get("Name",""))}</td>
<td style="padding:5px 10px;">{pill(state,"green" if state=="POWERED_ON" else "gray")}</td>
<td style="padding:5px 10px;text-align:center;">{v.get("vCPU","?")}</td>
<td style="padding:5px 10px;text-align:right;">{float(v.get("RAMgb",0) or 0):.0f} GB</td>
<td style="padding:5px 10px;text-align:right;">{disk_tot:.0f} GB</td>
<td style="padding:5px 10px;font-family:monospace;font-size:8.5pt;">{h(v.get("IPs","") or "—")}</td>
<td style="padding:5px 10px;font-size:8.5pt;color:var(--muted);">{h(v.get("GuestOS",""))}</td>
</tr>'''
        blocks.append(card(f'''<div style="font-size:13pt;font-weight:700;color:var(--text);margin-bottom:6px;">Virtual Machines — {h(hv.get("Server",""))}</div>
<div style="font-size:9pt;color:var(--muted);margin-bottom:10px;">{len(vms)} VMs on ESXi {h(hv.get("Version","") or hv.get("APIVersion",""))}</div>
<table style="width:100%;border-collapse:collapse;font-size:9pt;">
<thead><tr style="background:var(--text);color:white;">
<th style="text-align:left;padding:7px 10px;">VM</th>
<th style="padding:7px 10px;">State</th>
<th style="padding:7px 10px;">vCPU</th>
<th style="padding:7px 10px;">RAM</th>
<th style="padding:7px 10px;">Disk</th>
<th style="padding:7px 10px;">IP</th>
<th style="padding:7px 10px;">Guest OS</th>
</tr></thead><tbody>{rows}</tbody></table>'''))

    return ''.join(blocks) or '<div style="padding:40px;text-align:center;color:var(--muted);">No environment-level data to consolidate.</div>'

# ─── Risks tab ─────────────────────────────────────────────────────────
def render_risks(risks):
    groups = {'critical':[], 'warning':[], 'info':[]}
    for r in risks: groups.setdefault(r['severity'], []).append(r)
    out = ''
    labels = {'critical':('🔴 Critical','red'),'warning':('🟡 Warning','yellow'),'info':('🔵 Info','blue')}
    for sev in ('critical','warning','info'):
        items = groups.get(sev) or []
        if not items: continue
        lbl,_ = labels[sev]
        rows = ''
        for r in items:
            rows += f'<div style="padding:12px 14px;border-bottom:1px solid var(--border);"><div style="font-size:10pt;font-weight:700;color:var(--text);">{h(r["title"])}</div><div style="font-size:9pt;color:var(--muted);margin-top:3px;">{h(r["detail"])}</div><div style="font-size:8.5pt;color:var(--muted);margin-top:4px;">on: <strong>{h(r["server"])}</strong></div></div>'
        out += card(f'<div style="padding:14px 18px;border-bottom:1px solid var(--border);font-size:12pt;font-weight:700;color:var(--text);">{lbl} <span style="font-size:9pt;font-weight:400;color:var(--muted);">· {len(items)}</span></div><div>{rows}</div>')
    if not out:
        return '<div style="padding:40px;text-align:center;color:var(--success);font-size:12pt;font-weight:600;">✓ No risks detected.</div>'
    return out

# ─── Shell ─────────────────────────────────────────────────────────────
HTML_SHELL = '''<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><title>{title}</title>
<style>
:root {{
  --bg: #0B1220;
  --surface: #131A2B;
  --elevated: #1A2340;
  --elevated-2: #222C4A;
  --border: rgba(255,255,255,0.07);
  --border-2: rgba(255,255,255,0.11);
  --text: #E6EAF2;
  --muted: #8B95A8;
  --dim: #5A6478;
  --accent: #4F8CFF;
  --accent-2: #8B5CF6;
  --success: #22C55E;
  --warn: #F59E0B;
  --crit: #EF4444;
  --info: #38BDF8;
  --mono: "Cascadia Code","Consolas","JetBrains Mono",ui-monospace,monospace;
  --sans: "Segoe UI Variable Display","Segoe UI",system-ui,-apple-system,"Roboto",sans-serif;
}}
* {{ box-sizing: border-box; }}
html,body {{ margin:0; padding:0; }}
body {{
  background: var(--bg);
  color: var(--text);
  font-family: var(--sans);
  font-size: 14px;
  line-height: 1.5;
  -webkit-font-smoothing: antialiased;
  background-image: radial-gradient(1200px 600px at 80% -200px, rgba(79,140,255,0.12), transparent 60%), radial-gradient(900px 400px at -200px 200px, rgba(139,92,246,0.08), transparent 50%);
  background-attachment: fixed;
}}
.wrap {{ max-width: 1400px; margin: 0 auto; padding: 0 28px 60px; }}

/* Header */
.hdr {{
  background: rgba(11,18,32,0.75);
  backdrop-filter: blur(18px);
  -webkit-backdrop-filter: blur(18px);
  border-bottom: 1px solid var(--border);
  padding: 14px 28px;
  display: flex; justify-content: space-between; align-items: center;
  position: sticky; top: 0; z-index: 100;
}}
.brand {{
  font-size: 16px; font-weight: 700; letter-spacing: .5px;
  background: linear-gradient(92deg, #fff 0%, #a7b3ca 100%);
  -webkit-background-clip: text; background-clip: text;
  -webkit-text-fill-color: transparent;
}}
.brand-divider {{ color: var(--dim); margin: 0 14px; }}
.sub {{ color: var(--muted); font-size: 12px; font-weight: 500; letter-spacing: .1px; }}
.ver-chip {{ background: var(--elevated); border: 1px solid var(--border); color: var(--muted); font-size: 11px; font-weight: 600; padding: 4px 10px; border-radius: 99px; letter-spacing: .3px; }}

/* Tab nav */
.tab-nav {{
  display: flex; gap: 4px;
  margin: 18px 0 22px;
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 14px; padding: 6px;
  width: fit-content;
}}
.tab-btn {{
  background: transparent; border: none; color: var(--muted);
  padding: 10px 22px; font-size: 13px; font-weight: 600;
  border-radius: 10px; cursor: pointer; letter-spacing: .2px;
  font-family: var(--sans); transition: all .15s ease;
}}
.tab-btn:hover {{ color: var(--text); background: var(--elevated); }}
.tab-btn.active {{
  background: linear-gradient(135deg, #1c2540 0%, #222c4f 100%);
  color: var(--text);
  box-shadow: 0 0 0 1px var(--border-2), inset 0 1px 0 rgba(255,255,255,0.04);
}}
.tab-pane {{ display: none; animation: fadeIn .25s ease; }}
.tab-pane.active {{ display: block; }}
@keyframes fadeIn {{ from {{ opacity: 0; transform: translateY(4px); }} to {{ opacity: 1; transform: translateY(0); }} }}

a {{ color: var(--accent); text-decoration: none; }}
a:hover {{ text-decoration: underline; }}

/* Cards */
.card {{
  background: linear-gradient(180deg, var(--surface) 0%, rgba(19,26,43,0.94) 100%);
  border: 1px solid var(--border);
  border-radius: 14px; padding: 20px 22px; margin-bottom: 18px;
  position: relative;
}}
.card-title {{ font-size: 14px; font-weight: 700; color: var(--text); letter-spacing: -.01em; }}
.card-sub {{ font-size: 12px; color: var(--muted); margin-top: 2px; }}
.section-hdr {{ font-size: 11px; font-weight: 700; color: var(--accent); text-transform: uppercase; letter-spacing: .8px; margin: 20px 0 10px; }}

/* Bento KPI grid */
.bento {{
  display: grid;
  grid-template-columns: repeat(12, 1fr);
  gap: 16px; margin-bottom: 22px;
}}
.kpi {{
  grid-column: span 3;
  background: linear-gradient(160deg, var(--surface) 0%, var(--elevated) 100%);
  border: 1px solid var(--border);
  border-radius: 14px;
  padding: 18px 20px;
  position: relative; overflow: hidden;
}}
.kpi::before {{
  content: ''; position: absolute; top: -40%; right: -30%;
  width: 200px; height: 200px; border-radius: 50%;
  background: radial-gradient(circle, rgba(79,140,255,0.14), transparent 70%);
}}
.kpi-label {{ font-size: 10px; font-weight: 700; color: var(--muted); text-transform: uppercase; letter-spacing: .9px; }}
.kpi-num {{ font-size: 34px; font-weight: 700; letter-spacing: -.03em; margin-top: 6px; line-height: 1; background: linear-gradient(180deg, #fff 0%, #aab5cc 140%); -webkit-background-clip: text; background-clip: text; -webkit-text-fill-color: transparent; }}
.kpi-unit {{ font-size: 14px; font-weight: 500; color: var(--muted); margin-left: 3px; }}
.kpi-detail {{ font-size: 12px; color: var(--muted); margin-top: 6px; }}
.kpi.crit {{ border-color: rgba(239,68,68,0.25); }}
.kpi.crit::before {{ background: radial-gradient(circle, rgba(239,68,68,0.20), transparent 70%); }}
.kpi.warn {{ border-color: rgba(245,158,11,0.22); }}
.kpi.warn::before {{ background: radial-gradient(circle, rgba(245,158,11,0.18), transparent 70%); }}
.kpi.hero {{ grid-column: span 6; }}

/* Hero banner */
.hero {{
  grid-column: span 12;
  background: linear-gradient(125deg, rgba(79,140,255,0.18) 0%, rgba(139,92,246,0.14) 50%, rgba(34,197,94,0.08) 100%),
              linear-gradient(180deg, var(--elevated) 0%, var(--surface) 100%);
  border: 1px solid var(--border-2);
  border-radius: 18px;
  padding: 28px 32px;
  position: relative; overflow: hidden;
}}
.hero::after {{
  content: ''; position: absolute; inset: 0;
  background: radial-gradient(circle at 90% 10%, rgba(79,140,255,0.18), transparent 45%);
  pointer-events: none;
}}
.hero-eyebrow {{ font-size: 11px; font-weight: 700; color: var(--accent); text-transform: uppercase; letter-spacing: 1.2px; position: relative; z-index: 1; }}
.hero-title {{ font-size: 34px; font-weight: 700; letter-spacing: -.03em; margin: 6px 0 10px; color: var(--text); position: relative; z-index: 1; }}
.hero-meta {{ font-size: 13px; color: var(--muted); position: relative; z-index: 1; }}
.hero-badges {{ margin-top: 16px; display: flex; gap: 8px; flex-wrap: wrap; position: relative; z-index: 1; }}

/* Pills */
.pill {{ display: inline-flex; align-items: center; gap: 6px; padding: 3px 10px; border-radius: 99px; font-size: 11px; font-weight: 600; letter-spacing: .1px; white-space: nowrap; }}
.pill-dot {{ width: 6px; height: 6px; border-radius: 50%; }}
.pill.crit {{ background: rgba(239,68,68,0.14); color: #fca5a5; border: 1px solid rgba(239,68,68,0.25); }}
.pill.crit .pill-dot {{ background: #ef4444; box-shadow: 0 0 8px rgba(239,68,68,0.7); }}
.pill.warn {{ background: rgba(245,158,11,0.14); color: #fcd34d; border: 1px solid rgba(245,158,11,0.25); }}
.pill.warn .pill-dot {{ background: #f59e0b; }}
.pill.ok   {{ background: rgba(34,197,94,0.12); color: #86efac; border: 1px solid rgba(34,197,94,0.22); }}
.pill.ok .pill-dot {{ background: #22c55e; }}
.pill.info {{ background: rgba(56,189,248,0.12); color: #7dd3fc; border: 1px solid rgba(56,189,248,0.22); }}
.pill.info .pill-dot {{ background: #38bdf8; }}
.pill.neutral {{ background: rgba(139,149,168,0.10); color: var(--muted); border: 1px solid var(--border-2); }}
.pill.neutral .pill-dot {{ background: var(--muted); }}

/* Inventory table */
.inv {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
.inv th {{ text-align: left; padding: 10px 14px; font-size: 11px; font-weight: 600; color: var(--muted); text-transform: uppercase; letter-spacing: .6px; border-bottom: 1px solid var(--border); position: sticky; top: 0; background: var(--surface); }}
.inv td {{ padding: 11px 14px; border-bottom: 1px solid var(--border); vertical-align: middle; }}
.inv tr:last-child td {{ border-bottom: none; }}
.inv tr:hover td {{ background: rgba(255,255,255,0.02); }}
.inv a {{ font-weight: 600; }}
.mono {{ font-family: var(--mono); font-size: 12px; }}

/* Generic data tables */
table.dt {{ width: 100%; border-collapse: collapse; font-size: 12.5px; }}
table.dt th {{ text-align: left; padding: 8px 12px; font-size: 10px; font-weight: 700; color: var(--muted); text-transform: uppercase; letter-spacing: .6px; background: var(--elevated); border-bottom: 1px solid var(--border); }}
table.dt td {{ padding: 8px 12px; border-bottom: 1px solid var(--border); }}
table.dt tr:last-child td {{ border-bottom: none; }}
table.dt.dark th {{ background: var(--elevated-2); }}

/* Progress bars */
.pbar {{ background: rgba(255,255,255,0.06); border-radius: 3px; height: 6px; overflow: hidden; }}
.pbar > div {{ height: 100%; border-radius: 3px; }}

/* KV rows */
.kv {{ font-size: 12.5px; }}
.kv td {{ padding: 5px 0; vertical-align: top; }}
.kv .k {{ color: var(--muted); font-size: 11px; text-transform: uppercase; letter-spacing: .4px; width: 180px; font-weight: 600; }}
.kv .v {{ color: var(--text); }}

/* Server header */
.srv-head {{ display: flex; justify-content: space-between; align-items: flex-start; gap: 22px; padding-bottom: 16px; border-bottom: 1px solid var(--border); }}
.srv-title {{ font-size: 20px; font-weight: 700; letter-spacing: -.02em; color: var(--text); }}
.srv-sub {{ font-size: 12px; color: var(--muted); margin-top: 4px; }}
.srv-meta {{ text-align: right; font-size: 12px; color: var(--muted); }}

/* Role badges */
.role-chip {{ display: inline-block; background: rgba(79,140,255,0.10); color: #93b9ff; border: 1px solid rgba(79,140,255,0.22); font-size: 11px; font-weight: 600; padding: 3px 10px; border-radius: 6px; margin: 0 4px 4px 0; }}

/* Two-column content block */
.split {{ display: grid; grid-template-columns: 1fr 1fr; gap: 28px; margin-top: 16px; }}

/* Details (expandable) */
details {{ margin: 8px 0; }}
details summary {{ cursor: pointer; color: var(--muted); font-size: 12px; font-weight: 600; padding: 8px 10px; border-radius: 8px; background: var(--elevated); border: 1px solid var(--border); list-style: none; user-select: none; }}
details summary::-webkit-details-marker {{ display: none; }}
details summary::before {{ content: "▸"; display: inline-block; margin-right: 8px; transition: transform .18s ease; color: var(--dim); }}
details[open] summary::before {{ transform: rotate(90deg); }}
details summary:hover {{ color: var(--text); background: var(--elevated-2); }}
details > div {{ padding: 12px 4px 4px; }}

/* Risk strip in summary */
.risk-strip .risk-item {{ padding: 12px 0; border-bottom: 1px solid var(--border); }}
.risk-strip .risk-item:last-child {{ border-bottom: none; }}
.risk-title {{ font-size: 13px; font-weight: 600; color: var(--text); }}
.risk-detail {{ font-size: 12px; color: var(--muted); margin-top: 3px; }}
.risk-num {{ font-size: 11px; font-family: var(--mono); color: var(--dim); }}

/* Severity grouped cards (risks tab) */
.sev-card .sev-head {{ display: flex; align-items: center; gap: 10px; padding: 14px 18px; border-bottom: 1px solid var(--border); font-size: 14px; font-weight: 700; }}
.sev-count {{ font-size: 11px; font-weight: 500; color: var(--muted); margin-left: auto; font-family: var(--mono); }}

/* Warning callout */
.callout {{ background: rgba(245,158,11,0.08); border-left: 3px solid var(--warn); padding: 10px 14px; border-radius: 6px; font-size: 12px; color: #fcd34d; margin: 12px 0; }}

/* Scrollbar */
::-webkit-scrollbar {{ width: 10px; height: 10px; }}
::-webkit-scrollbar-track {{ background: var(--bg); }}
::-webkit-scrollbar-thumb {{ background: var(--elevated-2); border-radius: 5px; }}
::-webkit-scrollbar-thumb:hover {{ background: #2a355a; }}

/* Footer */
.footer {{ text-align: center; padding: 32px 0 10px; color: var(--dim); font-size: 11px; letter-spacing: .4px; }}

/* Responsive */
@media (max-width: 900px) {{
  .bento .kpi {{ grid-column: span 6; }}
  .bento .kpi.hero {{ grid-column: span 12; }}
  .split {{ grid-template-columns: 1fr; }}
}}
</style></head><body>
<div class="hdr">
  <div style="display:flex;align-items:center;">
    <div class="brand">MAGNA5</div>
    <span class="brand-divider">/</span>
    <div class="sub">{client} · Server Discovery · {date}</div>
  </div>
  <span class="ver-chip">SDT {version}</span>
</div>
<div class="wrap">
<div class="tab-nav">
  <button class="tab-btn active" data-tab="summary" onclick="setTab(\'summary\')">Summary</button>
  <button class="tab-btn" data-tab="servers" onclick="setTab(\'servers\')">Servers</button>
  <button class="tab-btn" data-tab="environment" onclick="setTab(\'environment\')">Environment</button>
  <button class="tab-btn" data-tab="risks" onclick="setTab(\'risks\')">Risks</button>
</div>
<div id="tab-summary" class="tab-pane active">{summary}</div>
<div id="tab-servers" class="tab-pane">{servers}</div>
<div id="tab-environment" class="tab-pane">{environment}</div>
<div id="tab-risks" class="tab-pane">{risks}</div>
<div class="footer">
Generated {date} · Magna5 Solutions Engineering · SDT {version}
</div>
</div>
<script>
function setTab(t){{
  document.querySelectorAll(".tab-btn").forEach(b=>b.classList.toggle("active",b.dataset.tab===t));
  document.querySelectorAll(".tab-pane").forEach(p=>p.classList.toggle("active",p.id==="tab-"+t));
  window.scrollTo({{top:0,behavior:\'smooth\'}});
}}
window.addEventListener("hashchange",()=>{{if(location.hash.startsWith("#srv-"))setTab("servers");}});
if(location.hash.startsWith("#srv-"))setTab("servers");
</script>
</body></html>'''

def build(client, servers, hypervisors, out_path):
    risks = collect_risks(servers, hypervisors)
    html_out = HTML_SHELL.format(
        title=f'{client} — Server Discovery Report',
        client=h(client), date=TODAY_STR, version=VERSION,
        summary=render_summary(client, servers, hypervisors, risks),
        servers=render_servers(servers, risks),
        environment=render_environment(servers, hypervisors),
        risks=render_risks(risks),
    )
    out_path.write_text(html_out, encoding='utf-8')
    return len(risks)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('session_dir')
    ap.add_argument('--client', default=None)
    ap.add_argument('--out', default=None)
    args = ap.parse_args()
    session = Path(args.session_dir).resolve()
    if not session.is_dir():
        print(f'[error] not a directory: {session}', file=sys.stderr); sys.exit(2)
    servers, hypervisors = load_session(session)
    if not servers and not hypervisors:
        print(f'[error] no discovery JSONs in {session}', file=sys.stderr); sys.exit(3)
    client = args.client or session.name
    out = Path(args.out) if args.out else session / f'{re.sub(r"[^A-Za-z0-9_-]+","_",client)}-DiscoveryReport-v2-{TODAY_STR}.html'
    n = build(client, servers, hypervisors, out)
    print(f'[ok] {len(servers)} servers, {len(hypervisors)} hypervisor(s), {n} risks')
    print(f'[ok] wrote {out}')

if __name__ == '__main__':
    main()
