"""
gen_report.py — Magna5 Server Discovery Report Generator
Usage: python gen_report.py <manifest.json>
Produces a single HTML file matching the MEKE gold-standard design.
"""
import json, html as htmlmod, sys, io, os, re, datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# ── DETECTION RULES ───────────────────────────────────────────────────────────
# Loaded from detection_rules.json alongside this script.
# Per-session override: place a detection_rules.json in the session folder —
# its entries are appended to the global rules (additive, not replacing).

def _load_rules():
    def _read(path):
        try:
            with open(path, encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}

    def _merge(base, over):
        result = {}
        for k in set(list(base.keys()) + list(over.keys())):
            bv, ov = base.get(k, []), over.get(k, [])
            if isinstance(bv, list) and isinstance(ov, list):
                result[k] = bv + [x for x in ov if x not in bv]
            elif isinstance(bv, dict) and isinstance(ov, dict):
                result[k] = _merge(bv, ov)
            else:
                result[k] = ov if ov else bv
        return result

    script_dir  = os.path.dirname(os.path.abspath(__file__))
    global_path = os.path.join(script_dir, 'detection_rules.json')
    rules = _read(global_path)

    # Per-session override (session folder = manifest directory)
    if 'manifest_dir' in globals():
        session_path = os.path.join(manifest_dir, 'detection_rules.json')
        if os.path.exists(session_path):
            rules = _merge(rules, _read(session_path))

    return rules

RULES = _load_rules()

if len(sys.argv) < 2:
    print("Usage: python gen_report.py <manifest.json>")
    sys.exit(1)

manifest_path = os.path.abspath(sys.argv[1])
manifest_dir  = os.path.dirname(manifest_path)

with open(manifest_path, encoding='utf-8-sig') as f:
    CFG = json.load(f)

def resolve(p):
    if not p: return p
    if os.path.isabs(p): return p
    return os.path.normpath(os.path.join(manifest_dir, p))

CLIENT       = CFG['client']
CLIENT_FULL  = CFG.get('client_full', CLIENT)
DATE         = CFG['date']
SESSION_DIR  = resolve(CFG['session_dir'])
OUTPUT_DIR   = resolve(CFG.get('output_dir', CFG['session_dir']))
OUTPUT       = os.path.join(OUTPUT_DIR, f"{CLIENT}-DiscoveryReport-{DATE}.html")
LOGO_FILE    = resolve(CFG.get('logo_file', ''))

# ── LOAD DATA ─────────────────────────────────────────────────────────────────
def jload(path):
    with open(path, encoding='utf-8-sig') as f:
        return json.load(f)

inv = jload(os.path.join(SESSION_DIR, CFG['inventory_file']))

# Auto-load all non-empty HV inventory files from the session directory
hv_inventories = []
for _fname in sorted(os.listdir(SESSION_DIR)):
    if 'inventory' not in _fname.lower() or not _fname.endswith('.json'): continue
    _fpath = os.path.join(SESSION_DIR, _fname)
    if os.path.getsize(_fpath) < 100: continue
    try:
        _hinv = jload(_fpath)
        if _hinv.get('_type') in ('HyperVInventory', 'vSphereInventory'):
            hv_inventories.append(_hinv)
    except Exception:
        pass

servers = []
for s in CFG['servers']:
    if s.get('os_type') == 'linux':
        # Linux entry — may have SSH JSON or be a placeholder
        data = jload(os.path.join(SESSION_DIR, s['file'])) if s.get('file') else {}
        servers.append({**s, 'data': data})
    else:
        path = os.path.join(SESSION_DIR, s['file'])
        data = jload(path)
        servers.append({**s, 'data': data})

LOGO_B64 = ''
if LOGO_FILE and os.path.exists(LOGO_FILE):
    with open(LOGO_FILE) as f:
        LOGO_B64 = f.read().strip()

# ── HELPERS ───────────────────────────────────────────────────────────────────
def h(s):
    return htmlmod.escape(str(s)) if s is not None else ''

def as_list(v):
    if isinstance(v, dict): return [v]
    if isinstance(v, list):
        result = []
        for item in v:
            if isinstance(item, dict):  result.append(item)
            elif isinstance(item, list): result.extend(x for x in item if isinstance(x, dict))
        return result
    return []

def pill(text, color):
    return f'<span class="pill pill-{color}">{h(str(text))}</span>'

def flag_div(sev, title, detail):
    return (f'<div class="flag-{sev}">'
            f'<div class="flag-label">{h(title)}</div>'
            f'<div class="flag-detail">{h(detail)}</div></div>\n')

def disk_bar(pct):
    c = '#d63638' if pct >= 85 else ('#f5a623' if pct >= 70 else '#20c800')
    return (f'<div class="disk-bar-bg"><div class="disk-bar-fill" '
            f'style="width:{min(pct,100)}%;background:{c}"></div></div>')

def card(cid, title, body_html, collapsed=False):
    btn = '&#9660; Expand' if collapsed else '&#9650; Collapse'
    bc  = 'card-body collapsed' if collapsed else 'card-body'
    return (f'<div class="card hide-sbr" id="{cid}">\n'
            f'<div class="card-title"><span>{title}</span>'
            f'<button class="collapse-btn" onclick="toggleCard(this)">{btn}</button></div>\n'
            f'<div class="{bc}">\n{body_html}\n</div>\n</div>\n')

def sub(t, style=''):
    s = f' style="{style}"' if style else ''
    return f'<div class="sub-title"{s}>{t}</div>\n'

def top_link(tid):
    return (f'<div style="text-align:right;margin-top:10px;">'
            f'<a href="#top-{tid}" style="color:#5b1fa4;font-size:8.5pt;'
            f'text-decoration:none;font-weight:600;">&#8593; Top</a></div>\n')

def mini_box(title, content, last=False):
    mb = '0' if last else '14px'
    return (f'<div style="background:#f5f4f8;border-radius:8px;padding:14px 16px;margin-bottom:{mb};">'
            f'<div style="font-size:8.5pt;font-weight:700;text-transform:uppercase;letter-spacing:.5px;'
            f'color:#5b1fa4;border-bottom:1.5px solid #ede9fe;padding-bottom:6px;margin-bottom:12px;">'
            f'{title}</div>{content}</div>\n')

def stor_row(drv, lbl, total, free, pct):
    c  = '#d63638' if pct >= 85 else ('#f5a623' if pct >= 70 else '#20c800')
    pc = 'red' if pct >= 85 else ('yellow' if pct >= 70 else 'green')
    lbl_part = f' &mdash; {h(lbl)}' if lbl else ''
    return (f'<tr style="padding:6px 0">'
            f'<td style="font-weight:700;white-space:nowrap">{h(drv)}{lbl_part}</td>'
            f'<td style="font-size:8.5pt;color:#6b6080">{total:.0f} GB total</td>'
            f'<td><span class="pill pill-{pc}">{pct}%</span></td>'
            f'<td style="min-width:100px"><div class="disk-bar-bg"><div class="disk-bar-fill" '
            f'style="width:{min(pct,100)}%;background:{c}"></div></div></td>'
            f'<td style="font-size:8.5pt;color:#6b6080">{free:.0f} GB free</td></tr>\n')

def sbr_grad(crit, warn):
    if crit: return 'linear-gradient(135deg,#d63638,#b92b2e)'
    if warn: return 'linear-gradient(135deg,#f5a623,#e0901a)'
    return 'linear-gradient(135deg,#20c800,#158f00)'

def sbr_badge(crit, warn):
    if crit: return '&#128308; CRITICAL'
    if warn: return '&#9888;&#65039; ATTENTION'
    return '&#9989; HEALTHY'

def tab_cls(crit, warn):
    if crit: return ' has-critical'
    if warn: return ' has-warning'
    return ''

def nav_link(anchor, label):
    return (f'<a href="#{anchor}" style="color:#5b1fa4;font-size:9pt;text-decoration:none;'
            f'font-weight:600;padding:4px 12px;border-radius:4px;background:#ede9fe;'
            f'border:1px solid #c4b5fd;">{label}</a>\n')

def unwrap_ad(raw):
    if isinstance(raw, dict): return raw
    if isinstance(raw, list):
        return next((x for x in raw if isinstance(x, dict) and 'Installed' in x), {})
    return {}

def extract_forwarders(raw):
    if isinstance(raw, list):
        return ', '.join(x.get('IPAddressToString', '') for x in raw
                         if isinstance(x, dict) and x.get('IPAddressToString'))
    if isinstance(raw, str) and raw.strip(): return raw
    return ''

# ── FSMO CROSS-REFERENCE ──────────────────────────────────────────────────────
def _norm_host(fqdn):
    return fqdn.split('.')[0].upper() if fqdn else ''

FSMO_FIELDS = ('PDCEmulator', 'RIDMaster', 'InfrastructureMaster', 'SchemaMaster', 'DomainNamingMaster')
FSMO_HOLDERS = set()
for _s in servers:
    _ad = unwrap_ad(_s['data'].get('AD', {}))
    if isinstance(_ad, dict):
        for _f in FSMO_FIELDS:
            _v = _ad.get(_f, '')
            if _v: FSMO_HOLDERS.add(_norm_host(_v))

# ── ROLE LABEL DERIVATION ─────────────────────────────────────────────────────
def derive_role_label(srv):
    data  = srv['data']
    name  = srv['name'].upper()
    roles = data.get('Roles', {})
    ad    = unwrap_ad(data.get('AD', {}))
    exch  = data.get('Exchange', {})
    sql   = data.get('SQL', {})
    sw    = data.get('FileShares', {})

    if isinstance(exch, list): exch = next((x for x in exch if isinstance(x, dict) and x.get('Installed')), {})
    role_list  = as_list(roles.get('InstalledRoles', []))
    role_names = {r.get('Name', '') for r in role_list}
    has_adds  = 'AD-Domain-Services' in role_names or (isinstance(ad, dict) and ad.get('Installed'))
    has_dhcp  = 'DHCP' in role_names
    has_nps   = 'NPAS' in role_names
    has_print = 'Print-Services' in role_names
    has_iis   = 'Web-Server' in role_names
    has_hv    = 'Hyper-V' in role_names
    sl = as_list(sw.get('Shares', [])) if isinstance(sw, dict) else []
    real = [s for s in sl if isinstance(s, dict)
            and not s.get('Name','').startswith('$')
            and s.get('Name','') not in ('ADMIN$','IPC$','C$','print$','NETLOGON','SYSVOL','address')
            and not (s.get('Path','') or '').endswith('LocalsplOnly')]
    exch_ok = isinstance(exch, dict) and exch.get('Installed')
    sql_inst = {}
    if isinstance(sql, dict): sql_inst = sql.get('Instances', {})
    elif isinstance(sql, list):
        _sq = next((x for x in sql if isinstance(x, dict)), {})
        sql_inst = _sq.get('Instances', {}) if _sq else {}
    has_sql = isinstance(sql_inst, dict) and bool(sql_inst.get('Edition', ''))

    if exch_ok: return 'Exchange Server'
    if has_adds:
        mods = []
        if name in FSMO_HOLDERS: mods.append('FSMO')
        if has_dhcp:             mods.append('DHCP')
        if has_nps:              mods.append('NPS')
        return 'Domain Controller' + (f' ({", ".join(mods)})' if mods else '')
    parts = []
    if real and has_print: parts.append('File & Print Server')
    elif real:             parts.append('File Server')
    elif has_print:        parts.append('Print Server')
    if has_sql:            parts.append('SQL Server')
    if has_iis and not parts: parts.append('Web Server')
    if has_hv:             parts.append('Hyper-V Host')
    return ' · '.join(parts) if parts else 'Windows Server'

# ── FLAG DERIVATION ───────────────────────────────────────────────────────────
def build_flags(srv):
    data  = srv['data']
    name  = srv['name']
    flags = []
    sys_  = data.get('System', {})
    hw    = data.get('Hardware', {})
    roles = data.get('Roles', {})
    ad    = unwrap_ad(data.get('AD', {}))
    exch  = data.get('Exchange', {})
    sql   = data.get('SQL', {})
    if isinstance(exch, list): exch = next((x for x in exch if isinstance(x, dict) and x.get('Installed')), {})
    if isinstance(sql, list):  sql  = next((x for x in sql if isinstance(x, dict)), {})

    feat_names = [f.get('Name','') for f in as_list(roles.get('InstalledFeatures', []))]
    has_smb1   = any('FS-SMB1' in n for n in feat_names)
    eol_status = sys_.get('OSEOLStatus', sys_.get('EOLStatus', ''))
    eol_date   = sys_.get('OSEOLDate',   sys_.get('EOLDate', ''))
    os_name    = sys_.get('OSName', '')

    if eol_status in ('EOL', 'Near EOL') or 'Server 2016' in os_name or 'Server 2008' in os_name:
        sev = 'critical' if eol_status == 'EOL' or 'Server 2008' in os_name else 'warning'
        flags.append((sev, 'Windows Server EOL — Upgrade Required',
            f'{os_name} — support ends {eol_date or "Oct 12, 2027 (WS2016)"}. '
            f'No security patches after end-of-support. Upgrade to Windows Server 2022.'))

    for d in as_list(data.get('Disks', [])):
        pct = d.get('UsedPct', 0); drv = d.get('Drive', '?')
        free = d.get('FreeGB', 0); total = d.get('TotalGB', 0)
        if pct >= 85:
            flags.append(('critical', f'Disk {drv} Near Capacity',
                f'{name} {drv}: {pct}% used — only {free:.1f} GB free of {total:.1f} GB. Risk of service disruption.'))
        elif pct >= 70:
            flags.append(('warning', f'Disk {drv} Space Moderate',
                f'{name} {drv}: {pct}% used ({free:.1f} GB free of {total:.1f} GB). Monitor closely.'))

    if has_smb1:
        flags.append(('critical', 'SMB 1.0/CIFS Enabled — Critical Security Risk',
            f'{name}: SMB 1.0 is the attack vector for WannaCry/NotPetya/EternalBlue ransomware. '
            f'Disable: Remove-WindowsFeature FS-SMB1'))

    ram_t = hw.get('RAMTotalGB', 0); ram_f = hw.get('RAMAvailGB', 0)
    if ram_t > 0 and (1 - ram_f / ram_t) > 0.80:
        pct = int((1 - ram_f / ram_t) * 100)
        flags.append(('warning', f'High Memory Utilization ({pct}%)',
            f'{name}: {pct}% RAM used ({ram_f:.1f} GB free of {ram_t:.1f} GB).'))

    up = float(sys_.get('UptimeDays', 0) or 0)
    if up > 90:
        flags.append(('warning', f'Extended Uptime — {up:.0f} Days',
            f'{name} has not been rebooted in {up:.0f} days. Pending updates may not be applied.'))

    if isinstance(exch, dict) and exch.get('Installed'):
        exch_eol = exch.get('EOLStatus', '')
        if exch_eol in ('EOL', 'Near EOL'):
            sev = 'critical' if exch_eol == 'EOL' else 'warning'
            flags.append((sev, f'Exchange {exch.get("VersionName","")} — {exch_eol}',
                f'{name}: reached end of support on {exch.get("EOLDate","")}. No security patches available.'))

    sql_inst = sql.get('Instances', {}) if isinstance(sql, dict) else {}
    if isinstance(sql_inst, dict) and sql_inst.get('Edition') and sql_inst.get('EOLStatus') == 'EOL':
        flags.append(('critical', f'{sql_inst.get("Edition","")} — EOL and WS2022 Incompatible',
            f'{name}: {sql_inst.get("Edition","")} (instance: {sql_inst.get("InstanceName","")}) EOL {sql_inst.get("EOLDate","")}. '
            f'NOT supported on Windows Server 2022. Must upgrade SQL to 2017+ before OS upgrade.'))

    return flags

# ── SERVICE CATEGORIZATION ────────────────────────────────────────────────────
_SK       = RULES.get('service_keywords', {})
_EDR_SVC  = tuple(_SK.get('edr',    []))
_PAM_SVC  = tuple(_SK.get('pam',    []))
_RMM_SVC  = tuple(_SK.get('rmm',    []))
_HV_SVC   = tuple(_SK.get('hyperv', []))
_PRINT_SVC = tuple(_SK.get('print', []))
_CORE_SVC = tuple(_SK.get('core',   []))

def categorize_svcs(svc_l):
    cats = {'EDR': [], 'PAM': [], 'RMM': [], 'HyperV': [], 'Core': [], 'Print': [],
            'Other': [], 'StoppedAuto': []}
    for svc in svc_l:
        if not isinstance(svc, dict): continue
        n = (svc.get('Name', '') + ' ' + svc.get('DisplayName', '')).lower()
        state = svc.get('State', '')
        mode  = svc.get('StartMode', '')
        if state in ('Stopped', 'Stop') and mode in ('Auto', 'Automatic'):
            cats['StoppedAuto'].append(svc)
        if state != 'Running': continue
        if   any(k in n for k in _EDR_SVC):   cats['EDR'].append(svc)
        elif any(k in n for k in _PAM_SVC):   cats['PAM'].append(svc)
        elif any(k in n for k in _RMM_SVC):   cats['RMM'].append(svc)
        elif any(k in n for k in _HV_SVC):    cats['HyperV'].append(svc)
        elif any(k in n for k in _PRINT_SVC): cats['Print'].append(svc)
        elif any(k in n for k in _CORE_SVC):  cats['Core'].append(svc)
        else:                                  cats['Other'].append(svc)
    return cats

def categorize_apps(app_l):
    cats = {'Security': [], 'Management': [], 'Browser': [], 'Other': []}
    skip_nm = ('microsoft visual c++', 'microsoft .net', 'windows sdk',
               'microsoft update health', 'update for windows', 'security update for')
    for app in app_l:
        if not isinstance(app, dict): continue
        nm  = (app.get('Name') or '').lower()
        pub = (app.get('Publisher') or '').lower()
        if any(s in nm for s in skip_nm): continue
        if pub == 'microsoft corporation' and not any(k in nm for k in ('edge', 'defender')): continue
        combined = nm + ' ' + pub
        _ak = RULES.get('app_keywords', {})
        if any(k in combined for k in _ak.get('security', [])):
            cats['Security'].append(app)
        elif any(k in combined for k in _ak.get('management', [])):
            cats['Management'].append(app)
        elif any(k in combined for k in _ak.get('browser', [])):
            cats['Browser'].append(app)
        else:
            cats['Other'].append(app)
    return cats

def detect_security(app_l, svc_l):
    all_n = [((a.get('Name') or '') + ' ' + (a.get('Publisher') or '')).lower() for a in app_l]
    all_s = [((s.get('DisplayName') or '') + ' ' + (s.get('Name') or '')).lower() for s in svc_l]
    combined = all_n + all_s
    edr = rmm = pam = bdr = rmt = None

    _sp = RULES.get('security_products', {})

    def _first_match(category):
        for entry in _sp.get(category, []):
            k, v = entry[0], entry[1]
            if any(k in n for n in combined):
                return v
        return None

    edr = _first_match('edr')
    rmm = _first_match('rmm')
    pam = _first_match('pam')
    rmt = _first_match('remote_access')
    bdr = _first_match('backup')

    return edr, rmm, pam, bdr, rmt

_SAFE_PATH = tuple(RULES.get('safe_paths', [
    'c:\\windows\\', '"c:\\windows\\',
    'c:\\program files\\common files\\microsoft',
    'c:\\program files\\windows defender',
    'c:\\program files\\microsoft',
    '"c:\\program files\\microsoft', '"c:\\program files\\windows'
]))

def find_svc_anomalies(svc_l):
    out = []
    for svc in svc_l:
        if not isinstance(svc, dict): continue
        if svc.get('State', '') != 'Running': continue
        path = (svc.get('Path', '') or '').strip().lower()
        if not path: continue
        if any(path.startswith(p) or path.lstrip('"').startswith(p.lstrip('"')) for p in _SAFE_PATH): continue
        if 'c:\\windows\\' in path: continue
        out.append({**svc, '_reason': f'Binary outside standard install dirs: {svc.get("Path","")[:120]}'})
    return out

# ── BUILD PER-SERVER TAB ──────────────────────────────────────────────────────
def build_linux_tab(srv):
    sid    = srv['id']
    name   = srv['name']
    ip     = srv.get('ip', '')
    guest  = srv.get('guest_os', 'Linux')
    in_sc  = srv.get('in_scope', True)
    data   = srv['data']
    has_data = bool(data and data.get('_type') == 'LinuxDiscovery')

    # Teal/dark-green colour palette for Linux tabs
    HDR   = '#0f4c5c'
    HDR2  = '#1a7a8a'
    TEAL  = '#0d9488'
    LTEAL = '#ccfbf1'
    DTEAL = '#134e4a'

    def lcard(title, body, collapsed=False):
        cid = f'{sid}-{re.sub(r"[^a-z0-9]","",title.lower())}'
        btn = f'<button class="collapse-btn" onclick="toggleCard(\'{cid}\')">' + ('▲ Collapse' if not collapsed else '▼ Expand') + '</button>'
        state = ' collapsed' if collapsed else ''
        return (f'<div class="card" style="border-color:#a7f3d0;">\n'
                f'<div class="card-title" style="color:{DTEAL};border-bottom-color:{TEAL};">{h(title)}{btn}</div>\n'
                f'<div class="card-body{state}" id="{cid}">{body}</div></div>\n')

    def disk_bar(pct):
        col = '#d63638' if pct >= 85 else '#f5a623' if pct >= 70 else TEAL
        return (f'<div class="disk-bar-bg"><div class="disk-bar-fill" '
                f'style="width:{min(pct,100)}%;background:{col};"></div></div>')

    scope_pill = f'<span class="pill pill-green">IN SCOPE</span>' if in_sc else f'<span class="pill pill-gray">OUT OF SCOPE</span>'

    # ── HEADER BANNER ─────────────────────────────────────────────────────────
    if has_data:
        os_pretty = data.get('OS', {}).get('PrettyName', guest)
        kernel    = data.get('OS', {}).get('Kernel', '')
        arch      = data.get('OS', {}).get('Architecture', '')
        cores     = data.get('CPU', {}).get('Cores', '?')
        cpu_model = data.get('CPU', {}).get('ModelName', '')
        mem_total = data.get('Memory', {}).get('TotalMB', 0)
        mem_used  = data.get('Memory', {}).get('UsedMB', 0)
        mem_free  = data.get('Memory', {}).get('FreeMB', 0)
        mem_pct   = int(mem_used / mem_total * 100) if mem_total else 0
        hostname  = data.get('Hostname', name)
        collected = data.get('CollectedAt', '')
        disks     = data.get('Disks', []) or []
        network   = data.get('Network', []) or []
        services  = data.get('Services', []) or []
        os_line   = h(os_pretty)
        meta_line = ' · '.join(filter(None, [h(kernel), h(arch)]))
    else:
        os_pretty = guest; hostname = name; cores = '?'
        mem_total = mem_used = mem_free = mem_pct = 0
        cpu_model = kernel = arch = collected = ''
        disks = network = services = []
        os_line = h(guest); meta_line = 'SSH discovery not collected'

    header = (f'<div style="background:linear-gradient(135deg,{HDR},{HDR2});border-radius:10px 10px 0 0;'
              f'padding:16px 24px;margin-bottom:0;">'
              f'<div style="display:flex;justify-content:space-between;align-items:center;">'
              f'<div>'
              f'<div style="font-size:18px;font-weight:700;color:#fff;">{h(hostname)} '
              f'<span style="font-size:10pt;font-weight:400;color:rgba(255,255,255,.65);">({h(ip)})</span></div>'
              f'<div style="font-size:9pt;color:rgba(255,255,255,.8);margin-top:4px;">{os_line}</div>'
              f'<div style="font-size:8pt;color:rgba(255,255,255,.55);margin-top:2px;">{meta_line}</div>'
              f'</div>'
              f'<div style="display:flex;flex-direction:column;align-items:flex-end;gap:6px;">'
              f'{scope_pill}'
              f'<span style="background:rgba(255,255,255,.15);color:#fff;font-size:8pt;padding:3px 10px;'
              f'border-radius:12px;font-weight:600;">🐧 Linux / Non-Windows</span>'
              f'</div></div></div>\n')

    # ── NO DATA PLACEHOLDER ────────────────────────────────────────────────────
    if not has_data:
        placeholder = (f'<div style="text-align:center;padding:48px 24px;color:#6b6080;">'
                       f'<div style="font-size:28px;margin-bottom:12px;">🐧</div>'
                       f'<div style="font-size:13pt;font-weight:700;color:{DTEAL};margin-bottom:8px;">'
                       f'SSH Discovery Not Collected</div>'
                       f'<div style="font-size:9.5pt;">This box was identified as Linux/non-Windows.<br>'
                       f'Re-run discovery and choose <strong>Y</strong> at the Linux SSH prompt to collect data.</div>'
                       f'<div style="margin-top:16px;font-size:8.5pt;color:#9ca3af;">'
                       f'Guest OS reported by hypervisor: {h(guest)}</div>'
                       f'</div>')
        tab_html = f'<div id="top-{sid}">\n{header}{placeholder}</div>\n'
        return {'id': sid, 'name': name, 'crit': 0, 'warn': 0, 'in_scope': in_sc, 'tab_html': tab_html}

    # ── SYSTEM OVERVIEW ────────────────────────────────────────────────────────
    ram_gb   = round(mem_total / 1024, 1) if mem_total else '?'
    ram_used = round(mem_used  / 1024, 1) if mem_used  else '?'
    stats = (f'<div class="stat-grid" style="grid-template-columns:repeat(3,1fr);">'
             f'<div class="stat-box"><div class="stat-num" style="color:{TEAL};">{cores}</div>'
             f'<div class="stat-lbl">CPU Cores</div></div>'
             f'<div class="stat-box"><div class="stat-num" style="color:{TEAL};">{ram_gb}</div>'
             f'<div class="stat-lbl">RAM (GB)</div></div>'
             f'<div class="stat-box"><div class="stat-num" style="color:{TEAL};">{mem_pct}%</div>'
             f'<div class="stat-lbl">RAM Used</div></div></div>')
    meta_rows = [('Hostname', hostname), ('IP', ip), ('OS', os_pretty),
                 ('Kernel', kernel), ('Architecture', arch),
                 ('CPU Model', cpu_model), ('Collected', collected)]
    meta_tbl = '<table>' + ''.join(
        f'<tr><td style="font-weight:600;width:130px;color:{DTEAL}">{h(k)}</td>'
        f'<td>{h(v)}</td></tr>'
        for k, v in meta_rows if v) + '</table>'
    overview_body = stats + meta_tbl
    overview_card = lcard('System Overview', overview_body)

    # ── DISKS ──────────────────────────────────────────────────────────────────
    disks_body = ''
    if disks:
        disks_body = ('<table><tr><th>Mount</th><th>Source</th><th>Size</th>'
                      '<th>Used</th><th>Free</th><th style="min-width:120px">Usage</th></tr>\n')
        for i, d in enumerate(disks):
            if not isinstance(d, dict): continue
            pct = d.get('UsePct', 0)
            bg  = 'background:#f5f4f8;' if i % 2 else ''
            disks_body += (f'<tr style="{bg}"><td style="font-family:monospace">{h(d.get("Mount",""))}</td>'
                           f'<td style="font-size:8.5pt;color:#6b6080">{h(d.get("Source",""))}</td>'
                           f'<td>{h(d.get("Size",""))}</td><td>{h(d.get("Used",""))}</td>'
                           f'<td>{h(d.get("Free",""))}</td>'
                           f'<td>{disk_bar(pct)}<span style="font-size:8pt;color:#6b6080">{pct}%</span></td></tr>\n')
        disks_body += '</table>'
    else:
        disks_body = '<div style="color:#9ca3af;font-style:italic">No disk data collected.</div>'
    disks_card = lcard('Disk Storage', disks_body)

    # ── NETWORK ────────────────────────────────────────────────────────────────
    net_body = ''
    if network:
        net_body = '<table><tr><th>Interface</th><th>Addresses</th></tr>\n'
        for i, a in enumerate(network):
            if not isinstance(a, dict): continue
            bg  = 'background:#f5f4f8;' if i % 2 else ''
            addrs = ', '.join(a.get('Addresses') or [])
            net_body += (f'<tr style="{bg}"><td style="font-family:monospace;font-weight:600">'
                         f'{h(a.get("Interface",""))}</td><td>{h(addrs)}</td></tr>\n')
        net_body += '</table>'
    else:
        net_body = '<div style="color:#9ca3af;font-style:italic">No network data collected.</div>'
    net_card = lcard('Network', net_body)

    # ── SERVICES ──────────────────────────────────────────────────────────────
    svc_body = ''
    if services:
        svc_body = ('<div style="display:flex;flex-wrap:wrap;gap:6px;">' +
                    ''.join(f'<span style="background:{LTEAL};color:{DTEAL};border-radius:4px;'
                            f'padding:3px 10px;font-size:8.5pt;font-weight:600;">{h(s.get("Name","") if isinstance(s,dict) else s)}</span>'
                            for s in services[:60]) + '</div>')
        if len(services) > 60:
            svc_body += f'<div style="font-size:8pt;color:#9ca3af;margin-top:8px;">+ {len(services)-60} more</div>'
    else:
        svc_body = '<div style="color:#9ca3af;font-style:italic">No service data collected.</div>'
    svc_card = lcard('Running Services', svc_body, collapsed=len(services) > 20)

    body = f'<div style="padding:0 16px 16px;">\n{overview_card}{disks_card}{net_card}{svc_card}</div>\n'
    tab_html = f'<div id="top-{sid}">\n{header}{body}</div>\n'
    return {'id': sid, 'name': name, 'crit': 0, 'warn': 0, 'in_scope': in_sc, 'tab_html': tab_html}


def build_server_tab(srv):
    data    = srv['data']
    sid     = srv['id']
    name    = srv['name']
    _rl     = srv.get('role_label', '')
    rlabel  = _rl if (_rl and _rl.lower() != name.lower()) else derive_role_label(srv)
    in_sc   = srv.get('in_scope', True)
    flags   = build_flags(srv)
    crit    = sum(1 for f in flags if f[0] == 'critical')
    warn    = sum(1 for f in flags if f[0] == 'warning')

    sys_  = data.get('System', {})
    hw    = data.get('Hardware', {})
    net   = data.get('Network', {})
    roles_d = data.get('Roles', {})
    ad    = unwrap_ad(data.get('AD', {}))
    dns   = data.get('DNS', {})
    dhcp  = data.get('DHCP', {})
    nps   = data.get('NPS', {})
    exch  = data.get('Exchange', {})
    sql   = data.get('SQL', {})
    sw    = data.get('FileShares', {})
    apps  = data.get('Apps', [])
    svcs  = data.get('Services', [])
    meta  = data.get('Meta', {})

    if isinstance(exch, list): exch = next((x for x in exch if isinstance(x, dict) and x.get('Installed')), {})
    if isinstance(sql, list):  sql  = next((x for x in sql if isinstance(x, dict)), {})
    sql_inst = sql.get('Instances', {}) if isinstance(sql, dict) else {}

    # System
    os_name    = sys_.get('OSName', '').replace('Microsoft ', '')
    os_build   = str(sys_.get('OSBuild', ''))
    last_boot  = sys_.get('LastBoot', sys_.get('LastBootTime', ''))
    eol_date   = sys_.get('OSEOLDate', sys_.get('EOLDate', ''))
    eol_status = sys_.get('OSEOLStatus', sys_.get('EOLStatus', ''))
    domain     = (sys_.get('Domain', '') or
                  (ad.get('DomainName', '') if isinstance(ad, dict) else '') or
                  (ad.get('ForestName', '') if isinstance(ad, dict) else ''))
    hostname   = sys_.get('Hostname', name)
    install_dt = sys_.get('OSInstallDate', '')
    ps_ver     = sys_.get('PSVersion', '')
    run_as     = sys_.get('RunAsUser', '')
    up         = float(sys_.get('UptimeDays', 0) or 0)
    collected  = meta.get('CollectedAt', '')

    # Hardware
    ram_t   = float(hw.get('RAMTotalGB', 0) or 0)
    ram_f   = float(hw.get('RAMAvailGB', 0) or 0)
    ram_pct = int((1 - ram_f / ram_t) * 100) if ram_t else 0
    cpu_c   = hw.get('CPUCores', '?')
    cpu_n   = hw.get('CPUName', '')
    vm_plat = hw.get('VMPlatform', '')
    is_vm   = hw.get('IsVM', True)
    mfr     = hw.get('Manufacturer', '')
    model   = hw.get('Model', '')
    serial  = hw.get('SerialNumber', '')
    bios_v  = hw.get('BIOSVersion', '')
    bios_d  = hw.get('BIOSDate', '')
    board   = hw.get('BoardProduct', '')

    # Network
    adapters   = net.get('Adapters', {})
    est_conns  = as_list(net.get('EstablishedConns', []))
    listen_raw = as_list(net.get('ListeningPorts', []))
    listen_ports = [x for x in listen_raw if isinstance(x, dict)]
    if not listen_ports:
        listen_ports = [x for x in est_conns if isinstance(x, dict) and x.get('State') == 'LISTENING']

    # Primary IP from adapters
    if isinstance(adapters, dict):
        ip_str = adapters.get('IPAddresses', adapters.get('IP', ''))
        ip     = ip_str.split(',')[0].strip() if ip_str else ''
        gw     = adapters.get('Gateway', '')
        dns_ip = adapters.get('DNS', '')
        mac    = adapters.get('MAC', '')
    else:
        ip = gw = dns_ip = mac = ''
        for a in as_list(adapters):
            if not isinstance(a, dict): continue
            s = a.get('IPAddresses', a.get('IP', ''))
            if s: ip = s.split(',')[0].strip()
            if a.get('Gateway'): gw = a['Gateway']
            if a.get('DNS'):     dns_ip = a['DNS']
            if a.get('MAC'):     mac = a['MAC']
            break

    # Roles
    role_list    = as_list(roles_d.get('InstalledRoles', []))
    feature_list = as_list(roles_d.get('InstalledFeatures', []))
    role_names_raw = {r.get('Name', '') for r in role_list}
    role_display   = {r.get('Name', ''): r.get('DisplayName', r.get('Name', '')) for r in role_list}
    feat_names     = [f.get('Name', '') for f in feature_list]
    has_smb1  = any('FS-SMB1' in n for n in feat_names)
    has_adds  = 'AD-Domain-Services' in role_names_raw or (isinstance(ad, dict) and ad.get('Installed'))
    has_dhcp  = 'DHCP' in role_names_raw or (isinstance(dhcp, dict) and dhcp.get('Installed'))
    has_dns   = 'DNS' in role_names_raw or (isinstance(dns, dict) and dns.get('Installed'))
    has_nps   = 'NPAS' in role_names_raw or (isinstance(nps, dict) and nps.get('Installed'))
    has_print = 'Print-Services' in role_names_raw
    has_hv    = 'Hyper-V' in role_names_raw
    has_files = 'FileAndStorage-Services' in role_names_raw

    # Shares
    sl = as_list(sw.get('Shares', [])) if isinstance(sw, dict) else []
    real_shares = [s for s in sl if isinstance(s, dict)
                   and not s.get('Name','').startswith('$')
                   and s.get('Name','') not in ('ADMIN$','IPC$','C$','print$','NETLOGON','SYSVOL','address')
                   and not (s.get('Path','') or '').endswith('LocalsplOnly')]

    disks    = as_list(data.get('Disks', []))
    svc_list = as_list(svcs)
    app_list = as_list(apps)

    svc_cats      = categorize_svcs(svc_list)
    app_cats      = categorize_apps(app_list)
    svc_anomalies = find_svc_anomalies(svc_list)
    edr, rmm, pam, bdr, rmt = detect_security(app_list, svc_list)

    running_count     = sum(len(v) for k, v in svc_cats.items() if k != 'StoppedAuto')
    stopped_auto_cnt  = len(svc_cats['StoppedAuto'])
    scope_pill        = pill("IN SCOPE", "green") if in_sc else pill("OUT OF SCOPE", "gray")

    # OS short labels
    os_yr = (re.search(r'20\d\d', os_name) or type('', (), {'group': lambda s,n: ''})()).group(0)
    os_short     = f'WS{os_yr[2:]}' if os_yr else os_name[:8]
    os_eol_str   = ('Supported to 2031' if '2022' in os_name else
                    'Supported to 2029' if '2019' in os_name else
                    'EOL Oct 2027'      if '2016' in os_name else eol_status or 'Check EOL')
    os_eol_color = 'green' if '2022' in os_name or '2019' in os_name else 'yellow'

    # Last boot age
    boot_pill = ''
    if last_boot:
        try:
            bd = datetime.datetime.strptime(last_boot[:10], '%Y-%m-%d')
            da = (datetime.datetime.now() - bd).days
            bclr = 'yellow' if da > 90 else 'green'
            note = ' &mdash; reboot recommended' if da > 90 else ''
            boot_pill = f' &nbsp;<span class="pill pill-{bclr}">{da} days ago{note}</span>'
        except: pass

    # ── SMB1 BANNER ──────────────────────────────────────────────────────────
    smb1_banner = ''
    if has_smb1:
        smb1_banner = (
            '<div style="background:#fff0f0;border:1.5px solid #d63638;border-radius:8px;'
            'padding:12px 16px;margin-bottom:16px;display:flex;align-items:flex-start;gap:12px;">'
            '<span style="font-size:18px;flex-shrink:0">&#9940;</span>'
            '<div><div style="font-size:9.5pt;font-weight:700;color:#d63638;">SMB 1.0/CIFS ENABLED — Critical Security Risk</div>'
            '<div style="font-size:9pt;color:#7f2424;margin-top:3px;">SMB 1.0 is the attack vector for WannaCry, NotPetya, and EternalBlue ransomware. '
            'Disable: <code>Remove-WindowsFeature FS-SMB1</code></div></div></div>\n')

    # ── SECURITY MINI-BOX (SBR) ──────────────────────────────────────────────
    def sec_row(icon, lbl, detail, status, sclr):
        bmap = {'green':'#20c800','yellow':'#f5a623','red':'#d63638','purple':'#5b1fa4','gray':'#9ca3af'}
        bgmap = {'green':'#f0fdf0','yellow':'#fff8e1','red':'#fff0f0','purple':'#f0f4ff','gray':'#f3f4f6'}
        fgmap = {'green':'#065f46','yellow':'#92400e','red':'#991b1b','purple':'#5b1fa4','gray':'#374151'}
        return (f'<div style="display:flex;align-items:center;gap:10px;padding:7px 0;border-bottom:1px solid #f0edf8;">'
                f'<span style="font-size:14px;flex-shrink:0">{icon}</span>'
                f'<div style="flex:1"><div style="font-size:9pt;font-weight:600;color:#271e41">{lbl}</div>'
                f'<div style="font-size:8pt;color:#6b6080">{detail}</div></div>'
                f'<span style="background:{bgmap.get(sclr,"#f3f4f6")};color:{fgmap.get(sclr,"#374151")};'
                f'font-size:7.5pt;font-weight:700;padding:2px 8px;border-radius:10px;'
                f'border:1px solid {bmap.get(sclr,"#9ca3af")};white-space:nowrap">{status}</span></div>\n')

    sec_rows  = sec_row('&#128737;', 'EDR / Endpoint',
                        edr or 'No data collected — verify agent present',
                        'PROTECTED' if edr else 'NOT DETECTED', 'green' if edr else 'yellow')
    sec_rows += sec_row('&#128295;', 'RMM / Management',
                        rmm or 'No data collected — verify agent present',
                        'DETECTED' if rmm else 'NOT DETECTED', 'green' if rmm else 'yellow')
    if pam:
        sec_rows += sec_row('&#128273;', 'Privileged Access (PAM)', pam, 'ACTIVE', 'purple')
    sec_rows += sec_row('&#128190;', 'Backup / BDR',
                        bdr or 'Not detected — verify with client',
                        'DETECTED' if bdr else 'NOT DETECTED', 'green' if bdr else 'yellow')
    if has_smb1:
        sec_rows += sec_row('&#9940;', 'SMB 1.0 Protocol',
                            'ENABLED — ransomware attack vector (disable immediately)',
                            'CRITICAL', 'red')
    security_mini = mini_box('Security &amp; Protection', sec_rows)

    # ── OS & SYSTEM MINI-BOX (SBR left) ──────────────────────────────────────
    plat_disp = vm_plat if vm_plat else ('Physical' if not is_vm else 'VM')
    os_sys_rows = (
        f'<tr><td style="color:#6b6080;width:110px;padding:3px 0">Platform</td>'
        f'<td>{h(plat_disp)}{(" VM on " + h(vm_plat)) if is_vm and vm_plat else ""}</td></tr>'
        f'<tr><td style="color:#6b6080;padding:3px 0">OS</td>'
        f'<td>{h(os_name)} <span class="pill pill-{os_eol_color}">{h(os_eol_str)}</span></td></tr>'
        f'<tr><td style="color:#6b6080;padding:3px 0">Last Reboot</td>'
        f'<td>{h(last_boot[:10] if last_boot else "")}{boot_pill}</td></tr>'
        f'<tr><td style="color:#6b6080;padding:3px 0">RAM</td>'
        f'<td>{ram_t:.1f} GB total &nbsp;'
        f'<span class="pill pill-{"yellow" if ram_pct > 75 else "green"}">'
        f'{ram_f:.1f} GB free ({100-ram_pct}%)</span></td></tr>'
        f'<tr><td style="color:#6b6080;padding:3px 0">Deployed</td>'
        f'<td>{h(install_dt[:10] if install_dt else "")}</td></tr>'
    )
    os_sys_mini = mini_box('OS &amp; System',
        f'<table style="width:100%;font-size:9pt;border-collapse:collapse;">{os_sys_rows}</table>')

    # ── AD MINI-BOX (SBR left, DCs only) ─────────────────────────────────────
    ad_sbr_mini = ''
    if has_adds and isinstance(ad, dict) and ad.get('Installed'):
        fl_raw = str(ad.get('DomainFL', ad.get('ForestFL', '')))
        fl_yr  = (re.search(r'20\d\d', fl_raw) or type('', (), {'group': lambda s,n: fl_raw})()).group(0) if fl_raw else ''
        fl_lbl = f'Windows Server {fl_yr}' if fl_yr else fl_raw
        _stale_raw = ad.get('StaleUsers', '')
        stale_u = len(_stale_raw) if isinstance(_stale_raw, list) else (int(_stale_raw) if str(_stale_raw).isdigit() else 0)
        # FSMORoles from list (actual data) — count as DC indicator
        fsmo_list_sbr = ad.get('FSMORoles', [])
        if not isinstance(fsmo_list_sbr, list): fsmo_list_sbr = []
        fsmo_count = len(fsmo_list_sbr)
        ad_rows = (
            f'<tr><td style="color:#6b6080;width:130px;padding:3px 0">Domain</td><td>{h(domain)}</td></tr>'
            + (f'<tr><td style="color:#6b6080;padding:3px 0">FSMO Roles</td>'
               f'<td><span class="pill pill-purple">{fsmo_count} role(s) on this DC</span></td></tr>'
               if fsmo_count else '')
            + f'<tr><td style="color:#6b6080;padding:3px 0">Functional Level</td>'
            f'<td>{h(fl_lbl)}'
            + ('&nbsp;<span class="pill pill-yellow">upgrade recommended</span>' if fl_yr and fl_yr < '2019' else '')
            + '</td></tr>'
        )
        if stale_u:
            ad_rows += (f'<tr><td style="color:#6b6080;padding:3px 0">Stale Users</td>'
                       f'<td><span class="pill pill-yellow">~{stale_u} accounts (90+ days inactive)</span></td></tr>')
        ad_sbr_mini = mini_box('Active Directory',
            f'<table style="width:100%;font-size:9pt;border-collapse:collapse;">{ad_rows}</table>')

    # ── STORAGE MINI-BOX (SBR left) ───────────────────────────────────────────
    disk_rows_s = ''.join(
        stor_row(d.get('Drive','?'), d.get('Label','') or '', d.get('TotalGB',0),
                 d.get('FreeGB',0), d.get('UsedPct',0))
        for d in disks)
    storage_mini = mini_box('Storage',
        '<table style="width:100%;border-collapse:collapse;">'
        '<tr><th style="text-align:left;font-size:8pt;padding:4px 0;color:#6b6080;font-weight:600">Drive</th>'
        '<th style="font-size:8pt;padding:4px 0;color:#6b6080;font-weight:600"></th>'
        '<th style="font-size:8pt;padding:4px 0;color:#6b6080;font-weight:600">Used</th>'
        '<th style="min-width:100px"></th>'
        '<th style="font-size:8pt;padding:4px 0;color:#6b6080;font-weight:600">Free</th></tr>'
        + disk_rows_s + '</table>')

    # ── SECURITY MINI-BOX (SBR left) ─────────────────────────────────────────
    def _sec_pill(val, missing_warn=True):
        if val:
            return f'<span class="pill pill-green">{h(val)}</span>'
        return f'<span class="pill pill-yellow">&#9888; None detected</span>' if missing_warn else '<span class="pill pill-gray">&mdash;</span>'

    sec_left_rows = (
        f'<tr><td style="color:#6b6080;width:90px;padding:4px 0;font-size:8.5pt;">&#128737;&#65039; EDR</td>'
        f'<td style="padding:4px 0">{_sec_pill(edr, missing_warn=True)}</td></tr>'
        f'<tr><td style="color:#6b6080;padding:4px 0;font-size:8.5pt;">&#9881;&#65039; RMM</td>'
        f'<td style="padding:4px 0">{_sec_pill(rmm, missing_warn=True)}</td></tr>'
        f'<tr><td style="color:#6b6080;padding:4px 0;font-size:8.5pt;">&#128279; Remote</td>'
        f'<td style="padding:4px 0">{_sec_pill(rmt, missing_warn=False)}</td></tr>'
        f'<tr><td style="color:#6b6080;padding:4px 0;font-size:8.5pt;">&#128190; Backup</td>'
        f'<td style="padding:4px 0">{_sec_pill(bdr, missing_warn=False)}</td></tr>'
        f'<tr><td style="color:#6b6080;padding:4px 0;font-size:8.5pt;">&#128273; PAM</td>'
        f'<td style="padding:4px 0">{_sec_pill(pam, missing_warn=False)}</td></tr>'
    )
    sec_left_mini = mini_box('Security &amp; Protection',
        f'<table style="width:100%;font-size:9pt;border-collapse:collapse;">{sec_left_rows}</table>',
        last=True)

    # ── ROLES MINI-BOX (SBR right) ────────────────────────────────────────────
    role_bdg = ''.join(f'<span class="role-badge">{h(role_display.get(r,r))}</span>'
                       for r in sorted(role_names_raw) if r)
    single_point_warn = ''
    if has_adds and has_dhcp and has_dns:
        single_point_warn = (f'<div style="margin-top:10px;font-size:8.5pt;color:#92400e;'
                            f'background:#fff8e1;border-radius:6px;padding:8px 12px;">&#9888;&#65039; '
                            f'All critical network services run on a single VM. If {h(name)} fails: '
                            f'users cannot log in, DNS may not resolve, or DHCP may not assign addresses.</div>')
    roles_mini = mini_box('Roles Running on This Server',
        f'<div style="display:flex;flex-wrap:wrap;gap:6px;">{role_bdg}</div>{single_point_warn}')

    # ── SHARES MINI-BOX (SBR right) ───────────────────────────────────────────
    shares_sbr_mini = ''
    if real_shares:
        shr = ''.join(
            f'<tr><td style="padding:5px 8px;font-weight:600">{h(s.get("Name",""))}</td>'
            f'<td style="padding:5px 8px;font-family:monospace;font-size:8pt">{h(s.get("Path",""))}</td></tr>'
            for s in real_shares[:6])
        shares_sbr_mini = mini_box(f'Network Shares ({len(real_shares)})',
            '<table style="width:100%;font-size:9pt;border-collapse:collapse;">'
            '<tr style="background:#ede9fe"><th style="padding:5px 8px;text-align:left;font-size:8pt">Share</th>'
            '<th style="padding:5px 8px;text-align:left;font-size:8pt">Path</th></tr>'
            + shr + '</table>', last=True)
    else:
        roles_mini = mini_box('Roles Running on This Server',
            f'<div style="display:flex;flex-wrap:wrap;gap:6px;">{role_bdg}</div>{single_point_warn}',
            last=True)

    # ── SBR BLOCK ─────────────────────────────────────────────────────────────
    sbr_html = f'''<div class="sbr-only">
<div style="background:{sbr_grad(crit,warn)};border-radius:10px 10px 0 0;padding:16px 24px;display:flex;justify-content:space-between;align-items:center;">
  <div>
    <div style="font-size:18px;font-weight:700;color:#fff;letter-spacing:.3px;">{h(name)}</div>
    <div style="font-size:9pt;color:rgba(255,255,255,.85);margin-top:3px;">{h(rlabel)} &middot; {h(ip)} &middot; {h(os_name)} &middot; {scope_pill}</div>
  </div>
  <div style="display:flex;align-items:center;gap:14px;">
    <div style="text-align:center;background:rgba(255,255,255,.{25 if crit else 15});border-radius:8px;padding:8px 14px;">
      <div style="font-size:22px;font-weight:700;color:#fff;">{crit}</div>
      <div style="font-size:7pt;color:rgba(255,255,255,.8);text-transform:uppercase;font-weight:600;">Critical</div>
    </div>
    <div style="text-align:center;background:rgba(255,255,255,.{25 if warn else 15});border-radius:8px;padding:8px 14px;">
      <div style="font-size:22px;font-weight:700;color:#fff;">{warn}</div>
      <div style="font-size:7pt;color:rgba(255,255,255,.8);text-transform:uppercase;font-weight:600;">Warning</div>
    </div>
    <span style="background:rgba(255,255,255,.22);color:#fff;font-size:10pt;font-weight:700;padding:6px 18px;border-radius:20px;border:1.5px solid rgba(255,255,255,.5);">{sbr_badge(crit,warn)}</span>
  </div>
</div>
<div style="background:white;border-radius:0 0 10px 10px;border:1px solid #e8e4f0;border-top:none;box-shadow:0 4px 14px rgba(0,0,0,.07);padding:20px 24px;margin-bottom:16px;">
{smb1_banner}<div style="display:grid;grid-template-columns:1fr 1fr;gap:18px;">
<div>{os_sys_mini}{ad_sbr_mini}{storage_mini}{sec_left_mini}</div>
<div>{security_mini}{roles_mini}{shares_sbr_mini}</div>
</div></div></div>
'''

    # ── NAV BAR ───────────────────────────────────────────────────────────────
    sect_links = [('Alerts', f'{sid}-alerts'), ('Overview', f'{sid}-overview'),
                  ('Hardware', f'{sid}-hardware'), ('Applications', f'{sid}-apps'),
                  ('Roles', f'{sid}-roles'), ('Role Config', f'{sid}-roleconfig'),
                  ('Disks', f'{sid}-disks'), ('Network', f'{sid}-network'),
                  ('Listening Ports', f'{sid}-lports'), ('Services', f'{sid}-services')]
    if svc_anomalies: sect_links.append(('Errors', f'{sid}-svc-anomalies'))
    if real_shares:   sect_links.append(('File Shares', f'{sid}-shares'))

    role_conf_links = {}
    if has_adds:  role_conf_links['Active Directory'] = f'{sid}-roleconf-ad'
    if has_dns:   role_conf_links['DNS Server'] = f'{sid}-roleconf-dns'
    if has_dhcp:  role_conf_links['DHCP Server'] = f'{sid}-roleconf-dhcp'
    if has_nps:   role_conf_links['NPS / RADIUS'] = f'{sid}-roleconf-nps'
    if has_hv:    role_conf_links['Hyper-V'] = f'{sid}-roleconf-hyperv'
    if has_files or real_shares: role_conf_links['File and Storage Services'] = f'{sid}-roleconf-files'

    nav_html = (
        f'<div id="top-{sid}" class="hide-sbr" style="background:white;border-radius:8px;'
        f'border:1px solid #e8e4f0;padding:12px 16px;margin-bottom:16px;box-shadow:0 2px 8px rgba(0,0,0,0.04);">'
        f'<div style="display:flex;flex-wrap:wrap;gap:6px;align-items:center;'
        f'margin-bottom:{"8px" if role_conf_links else "0"};">'
        f'<span style="font-size:8pt;color:#6b6080;text-transform:uppercase;letter-spacing:0.5px;margin-right:4px;font-weight:700;">SECTIONS</span>'
        + ''.join(nav_link(a, l) for l, a in sect_links)
        + '</div>\n'
        + (f'<div style="display:flex;flex-wrap:wrap;gap:6px;align-items:center;">'
           f'<span style="font-size:8pt;color:#6b6080;text-transform:uppercase;letter-spacing:0.5px;margin-right:4px;font-weight:700;">ROLES</span>'
           + ''.join(f'<a href="#{anch}" style="color:white;font-size:9pt;text-decoration:none;font-weight:600;padding:4px 12px;border-radius:4px;background:#5b1fa4;">{h(rl)} &#8595;</a>\n'
                     for rl, anch in role_conf_links.items())
           + '</div>\n' if role_conf_links else '')
        + '</div>\n'
    )

    # ── ALERTS CARD ───────────────────────────────────────────────────────────
    alerts_body = ''.join(flag_div(*f) for f in flags) or '<span class="pill pill-green">No critical alerts</span>\n'
    alerts_body += top_link(sid)

    # ── SYSTEM OVERVIEW CARD ──────────────────────────────────────────────────
    dom_short = domain.split('.')[0] if domain else 'WORKGROUP'
    os_disp   = os_name.replace('Windows Server ', 'WS').replace(' Standard','').replace(' Datacenter','').strip() or os_short
    def _ov_sec_row(icon, label, val, warn=True):
        if val:
            badge = f'<span class="pill pill-green">{h(val)}</span>'
        elif warn:
            badge = '<span class="pill pill-yellow">None detected</span>'
        else:
            badge = '<span style="color:#c4b5fd">&mdash;</span>'
        return (f'<div style="display:flex;align-items:center;justify-content:space-between;padding:6px 0;">'
                f'<span style="font-size:9pt;color:#6b6080;">{icon}&nbsp;{label}</span>'
                f'{badge}</div>\n')
    sec_col = (
        f'<div style="font-size:8pt;font-weight:700;text-transform:uppercase;letter-spacing:.8px;'
        f'color:#5b1fa4;margin-bottom:10px;">Security &amp; Protection</div>'
        + _ov_sec_row('🛡️', 'EDR / XDR',    edr, warn=True)
        + _ov_sec_row('⚙️', 'RMM',           rmm, warn=True)
        + _ov_sec_row('🔗', 'Remote Access', rmt, warn=False)
        + _ov_sec_row('💾', 'Backup',        bdr, warn=False)
        + _ov_sec_row('🔑', 'PAM',           pam, warn=False)
    )
    _eol_color = 'red' if eol_status == 'EOL' else ('yellow' if eol_status == 'Near EOL' else 'green')
    _eol_label = 'EOL' if eol_status == 'EOL' else 'Supported'
    _vm_badge  = f'&nbsp;<span class="pill pill-gray">VM · {h(vm_plat)}</span>' if is_vm and vm_plat else ''
    sys_col = (
        f'<div class="stat-grid" style="grid-template-columns:repeat(2,1fr);">'
        f'<div class="stat-box"><div class="stat-num" style="font-size:14px">{h(os_disp)}</div><div class="stat-lbl">OS</div></div>'
        f'<div class="stat-box"><div class="stat-num">{up:.1f}d</div><div class="stat-lbl">Uptime</div></div>'
        f'<div class="stat-box"><div class="stat-num" style="font-size:13px">{h(last_boot[:10] if last_boot else "—")}</div><div class="stat-lbl">Last Boot</div></div>'
        f'<div class="stat-box"><div class="stat-num" style="font-size:{"12" if len(dom_short) > 9 else "14"}px">{h(dom_short)}</div><div class="stat-lbl">Domain</div></div>'
        f'</div>'
        f'<div class="meta-line"><strong>OS:</strong> {h(sys_.get("OSName",""))} (Build {h(os_build)}){_vm_badge}</div>'
        f'<div class="meta-line"><strong>EOL Date:</strong> {h(eol_date)} <span class="pill pill-{_eol_color}">{_eol_label}</span></div>'
        f'<div class="meta-line"><strong>Installed:</strong> {h(install_dt[:10] if install_dt else "")}</div>'
        f'<div class="meta-line"><strong>PowerShell:</strong> {h(ps_ver)}</div>'
        f'<div class="meta-line"><strong>Run As:</strong> {h(run_as)}</div>'
        f'<div class="meta-line"><strong>Collected:</strong> {h(collected)}</div>'
    )
    overview_body = (
        f'<div style="display:grid;grid-template-columns:3fr 2fr;gap:32px;align-items:start;">'
        f'<div>{sys_col}</div>'
        f'<div style="border-left:2px solid #ede9fe;padding-left:24px;">{sec_col}</div>'
        f'</div>'
    ) + top_link(sid)

    # ── HARDWARE CARD ─────────────────────────────────────────────────────────
    plat_pill = pill(vm_plat or 'Physical', 'gray' if not vm_plat else 'purple')
    hw_body = f'''<div class="stat-grid">
<div class="stat-box"><div class="stat-num">{cpu_c}</div><div class="stat-lbl">CPU Cores</div></div>
<div class="stat-box"><div class="stat-num" style="font-size:16px">{ram_t:.1f} GB</div><div class="stat-lbl">RAM Total</div></div>
<div class="stat-box"><div class="stat-num" style="font-size:16px">{ram_f:.1f} GB</div><div class="stat-lbl">RAM Available</div></div>
<div class="stat-box"><div class="stat-num" style="font-size:14px">{plat_pill}</div><div class="stat-lbl">Platform</div></div>
</div>
<div class="meta-line"><strong>CPU:</strong> {h(cpu_n)}</div>
'''
    if mfr or model or serial:
        hw_body += sub('Hardware Identity', 'margin-top:16px')
        hw_body += '<table><tr><th style="width:160px">Property</th><th>Value</th></tr>\n'
        rows = [('Manufacturer', mfr), ('Model', model)]
        if serial:
            dell_lnk = (f' &nbsp;<a href="https://www.dell.com/support/home/en-us/product-support/servicetag/{h(serial)}"'
                       f' target="_blank" style="color:#5b1fa4;font-size:8.5pt">Look up warranty &#8599;</a>'
                       if 'dell' in mfr.lower() else '')
            rows.append(('Serial Number', f'{h(serial)}{dell_lnk}'))
        rows += [('BIOS Version', bios_v), ('BIOS Date', bios_d), ('Board Product', board)]
        for i, (k, v) in enumerate(rows):
            if not v: continue
            bg = ' style="background:#f5f4f8"' if i % 2 == 1 else ''
            safe_v = v if k == 'Serial Number' else h(v)
            hw_body += f'<tr{bg}><td><strong>{k}</strong></td><td>{safe_v}</td></tr>\n'
        hw_body += '</table>\n'
    hw_body += top_link(sid)

    # ── INSTALLED APPS CARD ───────────────────────────────────────────────────
    apps_body = ''
    cat_order = [('Security', 'Security &amp; EDR'), ('Management', 'Management &amp; RMM'),
                 ('Browser', 'Browser'), ('Other', 'Other')]
    for cat_key, cat_lbl in cat_order:
        ca = app_cats.get(cat_key, [])
        if not ca: continue
        apps_body += sub(cat_lbl)
        apps_body += '<table><tr><th>Application</th><th>Version</th><th>Publisher</th><th>Install Date</th></tr>\n'
        for i, ap in enumerate(ca):
            bg = ' style="background:#f5f4f8"' if i % 2 == 1 else ''
            apps_body += (f'<tr{bg}><td>{h(ap.get("Name",""))}</td>'
                         f'<td style="font-family:monospace;font-size:8.5pt">{h(ap.get("Version",""))}</td>'
                         f'<td>{h(ap.get("Publisher",""))}</td>'
                         f'<td>{h(ap.get("InstallDate","") or "&mdash;")}</td></tr>\n')
        apps_body += '</table>\n'
    if not apps_body:
        apps_body = ('<div class="flag-info"><div class="flag-label">No Data Collected</div>'
                    '<div class="flag-detail">Application inventory returned no data. '
                    'Collection may have failed on this server — re-run discovery to retry.</div></div>\n')
    apps_body += top_link(sid)

    # ── ROLES & FEATURES CARD ─────────────────────────────────────────────────
    _role_anchor = {'AD-Domain-Services': f'{sid}-roleconf-ad', 'DHCP': f'{sid}-roleconf-dhcp',
                    'DNS': f'{sid}-roleconf-dns', 'Hyper-V': f'{sid}-roleconf-hyperv',
                    'FileAndStorage-Services': f'{sid}-roleconf-files', 'NPAS': f'{sid}-roleconf-nps'}
    role_bdgs = ''
    for rl in role_list:
        rn = rl.get('Name', ''); rd = rl.get('DisplayName', rn)
        anch = _role_anchor.get(rn)
        if anch:
            role_bdgs += (f'<a href="#{anch}" style="text-decoration:none;">'
                         f'<span class="role-badge" style="font-size:11pt;padding:6px 16px;cursor:pointer;" '
                         f'title="Jump to Role Configuration">{h(rd)} &#8595;</span></a>\n')
        else:
            role_bdgs += f'<span class="role-badge" style="font-size:11pt;padding:6px 16px;">{h(rd)}</span>\n'
    roles_body = sub('INSTALLED ROLES')
    roles_body += f'<div style="background:#f0ebff;border-radius:6px;padding:12px;margin-bottom:16px;"><div class="role-grid">{role_bdgs}</div></div>\n'
    if feature_list:
        roles_body += sub(f'Installed Features ({len(feature_list)})', 'margin-top:14px')
        roles_body += '<table><tr><th>Name</th><th>Display Name</th></tr>\n'
        for i, ft in enumerate(feature_list):
            bg = ' style="background:#f5f4f8"' if i % 2 == 1 else ''
            roles_body += (f'<tr{bg}><td style="font-family:monospace;font-size:8.5pt">{h(ft.get("Name",""))}</td>'
                          f'<td>{h(ft.get("DisplayName",""))}</td></tr>\n')
        roles_body += '</table>\n'
    roles_body += top_link(sid)

    # ── ROLE CONFIGURATION CARD ───────────────────────────────────────────────
    rc = ''

    # Active Directory
    if has_adds and isinstance(ad, dict) and ad.get('Installed'):
        fl_d = str(ad.get('DomainFL', '')); fl_f = str(ad.get('ForestFL', ''))
        fy_d = (re.search(r'20\d\d', fl_d) or type('', (), {'group': lambda s,n: ''})()).group(0)
        fy_f = (re.search(r'20\d\d', fl_f) or type('', (), {'group': lambda s,n: ''})()).group(0)
        uc = str(ad.get('UserCount', '') or '—')
        cc = str(ad.get('ComputerCount', '') or '—')
        ou = str(ad.get('OUCount', '') or '—')
        pdce = ad.get('PDCEmulator', '')
        ridf = ad.get('RIDMaster', '')
        schema_m = ad.get('SchemaMaster', '')
        fsmo_list = ad.get('FSMORoles', [])
        if not isinstance(fsmo_list, list): fsmo_list = []
        _stale_raw2 = ad.get('StaleUsers', '')
        stale_u = len(_stale_raw2) if isinstance(_stale_raw2, list) else (int(_stale_raw2) if str(_stale_raw2).isdigit() else 0)
        def _dn_to_fqdn(v):
            if v and str(v).upper().startswith('DC='):
                return '.'.join(p.split('=')[1] for p in str(v).split(',') if '=' in p)
            return v or ''
        _raw_domain = ad.get('DomainName', '') or ad.get('ForestName', '') or domain
        domain_display = _dn_to_fqdn(_raw_domain)

        rc += f'<div id="{sid}-roleconf-ad"></div>\n'
        rc += sub('Active Directory', 'margin-top:20px')
        rc += f'''<div class="stat-grid">
<div class="stat-box"><div class="stat-num">{h(uc)}</div><div class="stat-lbl">Users</div></div>
<div class="stat-box"><div class="stat-num">{h(cc)}</div><div class="stat-lbl">Computers</div></div>
<div class="stat-box"><div class="stat-num">{h(ou)}</div><div class="stat-lbl">OUs</div></div>
<div class="stat-box"><div class="stat-num">{len(fsmo_list) if fsmo_list else "&mdash;"}</div><div class="stat-lbl">FSMO Roles Here</div></div>
</div>\n'''
        # FSMO roles held by this server — use FSMORoles list directly
        if fsmo_list:
            rc += sub('FSMO Roles Held By This DC', 'margin-top:12px')
            rc += '<div style="display:flex;flex-wrap:wrap;gap:10px;margin:10px 0;">\n'
            rc += ''.join(f'<div style="background:#271e41;color:white;border-radius:6px;padding:8px 18px;font-weight:700;font-size:10.5pt;">{h(fm)}</div>\n' for fm in fsmo_list)
            rc += '</div>\n'
        elif name.upper() in FSMO_HOLDERS:
            # Fallback: cross-ref from PDCEmulator/RIDMaster fields
            my_fsmo_fb = []
            if pdce   and _norm_host(pdce)    == name.upper(): my_fsmo_fb.append('PDC Emulator')
            if ridf   and _norm_host(ridf)    == name.upper(): my_fsmo_fb.append('RID Master')
            if schema_m and _norm_host(schema_m) == name.upper(): my_fsmo_fb.append('Schema Master')
            if my_fsmo_fb:
                rc += sub('FSMO Roles Held By This DC', 'margin-top:12px')
                rc += '<div style="display:flex;flex-wrap:wrap;gap:10px;margin:10px 0;">\n'
                rc += ''.join(f'<div style="background:#271e41;color:white;border-radius:6px;padding:8px 18px;font-weight:700;font-size:10.5pt;">{h(fm)}</div>\n' for fm in my_fsmo_fb)
                rc += '</div>\n'
            else:
                rc += sub('FSMO Roles', 'margin-top:12px')
                rc += '<div style="display:flex;flex-wrap:wrap;gap:10px;margin:10px 0;"><div style="background:#271e41;color:white;border-radius:6px;padding:8px 18px;font-weight:700;font-size:10.5pt;">FSMO Role Holder</div></div>\n'
        # FL warnings
        for yr, raw in ((fy_d, fl_d), (fy_f, fl_f)):
            if yr and yr < '2016':
                level_word = 'Domain' if raw == fl_d else 'Forest'
                rc += (f'<div class="flag-warning" style="margin-bottom:12px;">'
                      f'<div class="flag-label">&#9888; {level_word.upper()} FUNCTIONAL LEVEL UPGRADE RECOMMENDED</div>'
                      f'<div class="flag-detail">{level_word} functional level is <strong>{h(raw)}</strong>. '
                      f'Upgrading to Windows 2016+ enables PAM, improved Kerberos, and gMSA.</div></div>\n')
        # Domain table
        rc += sub('Domain &amp; Forest Details', 'margin-top:12px')
        rc += '<table><tr><th style="width:220px">Property</th><th>Value</th></tr>\n'
        dt_rows = [('Forest Root Domain', h(domain_display), ''),
                   ('Domain Functional Level', f'<span class="pill pill-{"yellow" if fy_d and fy_d<"2016" else "green"}">{h(fl_d)}</span>' if fl_d else '', ' style="background:#f5f4f8"'),
                   ('Forest Functional Level', f'<span class="pill pill-{"yellow" if fy_f and fy_f<"2016" else "green"}">{h(fl_f)}</span>' if fl_f else '', ''),
                   ('PDC Emulator', h(pdce), ' style="background:#f5f4f8"'),
                   ('RID Master', h(ridf), ''),
                   ('Schema Master', h(schema_m), ' style="background:#f5f4f8"')]
        for prop, val, bg in dt_rows:
            if not val: continue
            rc += f'<tr{bg}><td><strong>{prop}</strong></td><td>{val}</td></tr>\n'
        if stale_u:
            rc += f'<tr><td><strong>Stale User Accounts</strong></td><td><span class="pill pill-yellow">~{stale_u} accounts (90+ days inactive)</span></td></tr>\n'
        rc += '</table>\n'

    # DNS
    if isinstance(dns, dict) and dns.get('Installed'):
        zones = as_list(dns.get('Zones', []))
        fwd   = extract_forwarders(dns.get('Forwarders', ''))
        rc += f'<div id="{sid}-roleconf-dns"></div>\n'
        rc += sub('DNS Server', 'margin-top:24px')
        rc += '<table><tr><th style="width:160px">Property</th><th>Value</th></tr>\n'
        rc += f'<tr><td><strong>DNS Zones</strong></td><td>{len(zones)}</td></tr>\n'
        if fwd: rc += f'<tr style="background:#f5f4f8"><td><strong>Forwarders</strong></td><td><code>{h(fwd)}</code></td></tr>\n'
        rc += '</table>\n'
        if zones:
            rc += sub(f'DNS Zones ({len(zones)})', 'margin-top:10px;font-size:10px')
            rc += '<table><tr><th>Zone Name</th><th>Type</th><th>Dynamic Update</th></tr>\n'
            for i, z in enumerate(zones[:20]):
                if not isinstance(z, dict): continue
                bg = ' style="background:#f5f4f8"' if i % 2 == 1 else ''
                rc += (f'<tr{bg}><td style="font-family:monospace;font-size:8.5pt">{h(z.get("ZoneName",""))}</td>'
                      f'<td>{h(z.get("ZoneType",""))}</td><td>{h(str(z.get("DynamicUpdate","")))}</td></tr>\n')
            rc += '</table>\n'

    # DHCP
    if isinstance(dhcp, dict) and dhcp.get('Installed'):
        scopes = as_list(dhcp.get('Scopes', []))
        rc += f'<div id="{sid}-roleconf-dhcp"></div>\n'
        rc += sub('DHCP Server', 'margin-top:24px')
        rc += '<table><tr><th>Scope</th><th>Name</th><th>State</th><th>In Use</th><th>Free</th><th>Start</th><th>End</th></tr>\n'
        for i, sc in enumerate(scopes):
            if not isinstance(sc, dict): continue
            bg = ' style="background:#f5f4f8"' if i % 2 == 1 else ''
            st = sc.get('State', '')
            rc += (f'<tr{bg}><td style="font-family:monospace;font-size:8.5pt">{h(sc.get("ScopeId",""))}</td>'
                  f'<td>{h(sc.get("Name",""))}</td>'
                  f'<td>{pill(st, "green" if st=="Active" else "gray")}</td>'
                  f'<td>{h(str(sc.get("InUse", sc.get("AddressesInUse","?"))))}</td>'
                  f'<td>{h(str(sc.get("Available", sc.get("AddressesFree","?"))))}</td>'
                  f'<td style="font-family:monospace;font-size:8.5pt">{h(sc.get("StartRange",""))}</td>'
                  f'<td style="font-family:monospace;font-size:8.5pt">{h(sc.get("EndRange",""))}</td></tr>\n')
        if not scopes:
            rc += '<tr><td colspan="7" style="color:#6b6080;font-style:italic">DHCP installed — no scope data collected</td></tr>\n'
        rc += '</table>\n'

    # NPS
    if has_nps:
        rc += f'<div id="{sid}-roleconf-nps"></div>\n'
        rc += sub('NPS / RADIUS', 'margin-top:24px')
        if isinstance(nps, dict) and (nps.get('NetworkPolicies') or nps.get('RadiusClients')):
            np_c = len(as_list(nps.get('NetworkPolicies', [])))
            rc_c = len(as_list(nps.get('RadiusClients', [])))
            rc += f'<div style="font-size:9pt;margin-bottom:8px">{np_c} network policies, {rc_c} RADIUS clients configured.</div>\n'
        else:
            rc += ('<div class="flag-info"><div class="flag-label">NPS Installed</div>'
                  '<div class="flag-detail">Network Policy Server (RADIUS) is installed. '
                  'Policies and RADIUS clients must be exported and migrated during DC replacement.</div></div>\n')

    # Hyper-V
    if has_hv:
        hv_data = data.get('HyperV', {})
        rc += f'<div id="{sid}-roleconf-hyperv"></div>\n'
        rc += sub('Hyper-V Host', 'margin-top:24px')
        hv_vms = as_list(hv_data.get('VMs', [])) if isinstance(hv_data, dict) else []
        if hv_vms:
            running_vms = [v for v in hv_vms if v.get('State', '') == 'Running']
            rc += f'''<div class="stat-grid">
<div class="stat-box"><div class="stat-num">{len(hv_vms)}</div><div class="stat-lbl">Total VMs</div></div>
<div class="stat-box"><div class="stat-num" style="color:#20c800">{len(running_vms)}</div><div class="stat-lbl">Running</div></div>
</div>\n'''
            rc += sub('Hosted VMs', 'margin-top:12px')
            rc += '<table><tr><th>VM Name</th><th>State</th><th>vCPU</th><th>RAM (GB)</th><th>Generation</th><th>Uptime</th></tr>\n'
            for i, v in enumerate(hv_vms):
                bg = ' style="background:#f5f4f8"' if i % 2 == 1 else ''
                rc += (f'<tr{bg}><td>{h(v.get("Name",""))}</td>'
                      f'<td>{pill(v.get("State",""), "green" if v.get("State")=="Running" else "gray")}</td>'
                      f'<td>{h(str(v.get("vCPU","")))}</td><td>{h(str(v.get("RAM","")))}</td>'
                      f'<td>{h(str(v.get("Generation","")))}</td><td style="font-family:monospace;font-size:8.5pt">{h(str(v.get("Uptime","")))}</td></tr>\n')
            rc += '</table>\n'
        else:
            rc += ('<div class="flag-info"><div class="flag-label">Hyper-V Role Installed</div>'
                  '<div class="flag-detail">Hyper-V is installed on this server. VM inventory not collected in this run.</div></div>\n')

    # File & Storage
    if has_files or real_shares:
        rc += f'<div id="{sid}-roleconf-files"></div>\n'
        rc += sub('File and Storage Services', 'margin-top:24px')
        if real_shares:
            rc += (f'<div class="flag-info"><div class="flag-label">Shares</div>'
                  f'<div class="flag-detail">{len(real_shares)} business shares detected. '
                  f'See the <a href="#{sid}-shares" style="color:#5b1fa4;font-weight:700">File Shares section</a> below.</div></div>\n')
        else:
            rc += '<div class="flag-info"><div class="flag-label">File and Storage Services</div><div class="flag-detail">File and Storage Services is installed. No non-admin shares detected.</div></div>\n'

    # Exchange
    if isinstance(exch, dict) and exch.get('Installed'):
        eol_c = 'red' if exch.get('EOLStatus') == 'EOL' else ('yellow' if exch.get('EOLStatus') == 'Near EOL' else 'green')
        rc += sub('Exchange Server', 'margin-top:24px')
        rc += '<table><tr><th style="width:160px">Property</th><th>Value</th></tr>\n'
        rc += f'<tr><td><strong>Version</strong></td><td>{h(exch.get("VersionName",""))}</td></tr>\n'
        rc += f'<tr style="background:#f5f4f8"><td><strong>EOL Status</strong></td><td>{pill(exch.get("EOLStatus","?"), eol_c)} &nbsp;{h(exch.get("EOLDate",""))}</td></tr>\n'
        rc += f'<tr><td><strong>Transport</strong></td><td>{pill("Running","green") if exch.get("TransportServiceRunning") else pill("Stopped","red")}</td></tr>\n'
        rc += '</table>\n'

    # SQL
    if isinstance(sql_inst, dict) and sql_inst.get('Edition'):
        sql_eol_status = sql_inst.get('EOLStatus', '')
        sql_eol_c = 'red' if sql_eol_status == 'EOL' else ('yellow' if sql_eol_status == 'Near EOL' else 'green')
        sql_edition = sql_inst.get('Edition', '')
        sql_ver = sql_inst.get('Version', '—')
        sql_eol_date = sql_inst.get('EOLDate', '—')
        sql_svc_acct = sql_inst.get('ServiceAccount', '—')
        sql_inst_name = sql_inst.get('InstanceName', 'Default')
        # Flag old editions
        old_sql_flag = any(yr in sql_edition for yr in ('2016', '2014', '2012'))
        rc += sub('SQL Server', 'margin-top:24px')
        if old_sql_flag:
            rc += (f'<div class="flag-critical" style="margin-bottom:12px;">'
                   f'<div class="flag-label">&#9888; SQL Server EOL / Near-EOL — Upgrade Required</div>'
                   f'<div class="flag-detail">{h(sql_edition)} reaches end of support {h(sql_eol_date)}. '
                   f'<strong>Not supported on Windows Server 2022.</strong> Must upgrade SQL to 2017+ before OS upgrade.</div></div>\n')
        rc += '<table><tr><th style="width:200px">Property</th><th>Value</th></tr>\n'
        rc += f'<tr><td><strong>Edition</strong></td><td>{h(sql_edition)}</td></tr>\n'
        rc += f'<tr style="background:#f5f4f8"><td><strong>Version</strong></td><td><code>{h(sql_ver)}</code></td></tr>\n'
        rc += f'<tr><td><strong>Instance Name</strong></td><td><code>{h(sql_inst_name)}</code></td></tr>\n'
        rc += f'<tr style="background:#f5f4f8"><td><strong>EOL Date</strong></td><td>{h(sql_eol_date)} &nbsp;{pill(sql_eol_status or "Unknown", sql_eol_c)}</td></tr>\n'
        rc += f'<tr><td><strong>Service Account</strong></td><td><code>{h(sql_svc_acct)}</code></td></tr>\n'
        rc += f'<tr style="background:#f5f4f8"><td><strong>WS2022 Compatibility</strong></td><td>{pill("NOT SUPPORTED","red") if sql_eol_status=="EOL" else pill("Check Version","yellow")}</td></tr>\n'
        rc += f'<tr><td><strong>Database List</strong></td><td style="color:#6b6080;font-style:italic;">Database list unavailable — pull from SQL Management Studio</td></tr>\n'
        rc += '</table>\n'

    if not rc.strip():
        rc = '<div class="flag-info"><div class="flag-label">No Role Configuration Data</div><div class="flag-detail">No configurable server roles were detected on this server.</div></div>\n'
    rc += top_link(sid)

    # ── DISK STORAGE CARD ─────────────────────────────────────────────────────
    disk_body = '<table><tr><th>Drive</th><th>Label</th><th>Filesystem</th><th>Total (GB)</th><th>Used (GB)</th><th>Free (GB)</th><th style="min-width:140px">Utilization</th></tr>\n'
    for d in disks:
        pct      = d.get('UsedPct', 0)
        total_gb = d.get('TotalGB', 0)
        free_gb  = d.get('FreeGB', 0)
        used_gb  = round(total_gb - free_gb, 2)
        pc  = 'red' if pct >= 85 else ('yellow' if pct >= 70 else 'green')
        c   = '#d63638' if pct >= 85 else ('#f5a623' if pct >= 70 else '#20c800')
        disk_body += (f'<tr><td><strong>{h(d.get("Drive","?"))}</strong></td>'
                     f'<td>{h(d.get("Label","") or "")}</td>'
                     f'<td>{h(d.get("Filesystem","NTFS"))}</td>'
                     f'<td>{total_gb:.2f}</td><td>{used_gb:.2f}</td><td>{free_gb:.2f}</td>'
                     f'<td><div class="disk-bar-bg"><div class="disk-bar-fill" '
                     f'style="width:{min(pct,100)}%;background:{c}"></div></div>'
                     f'<span style="font-size:8pt;color:#6b6080;">{pct}%</span></td></tr>\n')
    disk_body += '</table>\n' + top_link(sid)

    # ── NETWORK CARD ──────────────────────────────────────────────────────────
    net_body = sub('Network Adapters')
    net_body += '<table><tr><th>Description</th><th>IP Address(es)</th><th>Gateway</th><th>DNS</th><th>MAC</th><th>DHCP</th></tr>\n'
    if isinstance(adapters, dict):
        ip_s = adapters.get('IPAddresses', adapters.get('IP', ''))
        net_body += (f'<tr><td>{h(adapters.get("Description","Network Adapter"))}</td>'
                    f'<td>{h(ip_s)}</td><td>{h(adapters.get("Gateway",""))}</td>'
                    f'<td>{h(adapters.get("DNS",""))}</td>'
                    f'<td style="font-family:monospace;font-size:8.5pt">{h(adapters.get("MAC",""))}</td>'
                    f'<td>{pill("No","green")}</td></tr>\n')
    else:
        for i, a in enumerate(as_list(adapters)):
            if not isinstance(a, dict): continue
            bg  = ' style="background:#f5f4f8"' if i % 2 == 1 else ''
            ip_s = a.get('IPAddresses', a.get('IP', ''))
            net_body += (f'<tr{bg}><td>{h(a.get("Description",""))}</td>'
                        f'<td>{h(ip_s)}</td><td>{h(a.get("Gateway",""))}</td>'
                        f'<td>{h(a.get("DNS",""))}</td>'
                        f'<td style="font-family:monospace;font-size:8.5pt">{h(a.get("MAC",""))}</td>'
                        f'<td>{pill("Yes","green") if a.get("DHCPEnabled") else pill("No","green")}</td></tr>\n')
    net_body += '</table>\n'
    ec = [c for c in est_conns if isinstance(c, dict) and c.get('State') == 'ESTABLISHED']
    if ec:
        show = ec[:15]
        net_body += sub(f'Established Connections ({len(ec)} total{", showing top 15" if len(ec) > 15 else ""})',
                        'margin-top:14px')
        net_body += '<table><tr><th>Process</th><th>Local Port</th><th>Remote</th><th>Protocol</th><th>State</th><th>PID</th></tr>\n'
        for i, c in enumerate(show):
            bg = ' style="background:#f5f4f8"' if i % 2 == 1 else ''
            rem = c.get('Remote', '')
            port = str(c.get('Port', ''))
            if port and ':' not in rem: rem = f'{rem}:{port}'
            net_body += (f'<tr{bg}><td>{h(c.get("Process",""))}</td>'
                        f'<td>{h(str(c.get("LocalPort","")))}</td>'
                        f'<td>{h(rem)}</td><td>{h(c.get("Proto","TCP"))}</td>'
                        f'<td>{h(c.get("State",""))}</td><td>{h(str(c.get("PID","")))}</td></tr>\n')
        net_body += '</table>\n'
    net_body += top_link(sid)

    # ── LISTENING PORTS CARD (collapsed) ──────────────────────────────────────
    lp_body = sub(f'Listening Ports ({len(listen_ports)} unique)')
    lp_body += '<table><tr><th>Port</th><th>Process</th><th>Protocol</th><th>Local IP</th><th>PID</th></tr>\n'
    for i, lp in enumerate(listen_ports[:60]):
        if not isinstance(lp, dict): continue
        bg = ' style="background:#f5f4f8"' if i % 2 == 1 else ''
        lp_body += (f'<tr{bg}><td>{h(str(lp.get("Port", lp.get("LocalPort",""))))}</td>'
                   f'<td>{h(lp.get("Process",""))}</td><td>{h(lp.get("Proto","TCP"))}</td>'
                   f'<td>{h(str(lp.get("LocalIP","")))}</td><td>{h(str(lp.get("PID","")))}</td></tr>\n')
    if not listen_ports:
        lp_body += '<tr><td colspan="5" style="color:#6b6080;font-style:italic">No listening port data collected</td></tr>\n'
    lp_body += '</table>\n' + top_link(sid)

    # ── SERVICES CARD ─────────────────────────────────────────────────────────
    svc_body = f'''<div class="stat-grid" style="grid-template-columns: repeat(2,1fr)">
<div class="stat-box"><div class="stat-num" style="color:#20c800">{running_count}</div><div class="stat-lbl">Running</div></div>
<div class="stat-box"><div class="stat-num" style="color:{"#d63638" if stopped_auto_cnt else "#6b6080"}">{stopped_auto_cnt}</div><div class="stat-lbl">Stopped (Auto-start)</div></div>
</div>\n'''
    if stopped_auto_cnt:
        svc_body += ('<div class="flag-warning" style="margin:8px 0 12px;"><div class="flag-label">&#9888; Stopped Auto-Start Services</div>'
                    '<div class="flag-detail">These services are set to auto-start but are currently stopped:</div></div>\n'
                    '<table><tr><th>Display Name</th><th>Name</th><th>Account</th></tr>\n')
        for i, s in enumerate(svc_cats['StoppedAuto'][:10]):
            bg = ' style="background:#fff8e1"'
            svc_body += (f'<tr{bg}><td>{h(s.get("DisplayName",""))}</td>'
                        f'<td style="font-family:monospace;font-size:8.5pt">{h(s.get("Name",""))}</td>'
                        f'<td style="font-size:8.5pt">{h(s.get("StartName",""))}</td></tr>\n')
        svc_body += '</table>\n'
    svc_body += f'<div class="sub-title" style="margin-top:10px">Running Services ({running_count})</div>\n'
    _edr_lbl = f'EDR / Endpoint Protection — {edr}' if edr else 'EDR / Endpoint Protection'
    cat_display = [('EDR', _edr_lbl), ('PAM', 'Privileged Access Management'),
                   ('RMM', 'RMM / Managed Services'), ('HyperV', 'Virtualization (Hyper-V)'),
                   ('Core', 'Windows Core Services'), ('Print', 'Print Services'), ('Other', 'Other')]
    for cat_k, cat_lbl in cat_display:
        svcs_c = svc_cats.get(cat_k, [])
        if not svcs_c: continue
        svc_body += f'<div class="sub-title" style="margin-top:12px;font-size:9.5px;color:#6b6080">{cat_lbl}</div>\n'
        if cat_k == 'Other' and len(svcs_c) > 10:
            svc_body += ('<div class="other-svc-wrapper" style="margin-bottom:8px">'
                        '<div style="display:flex;justify-content:space-between;align-items:center;">'
                        f'<span style="font-size:8.5pt;color:#6b6080">{len(svcs_c)} services</span>'
                        '<button class="collapse-btn" onclick="var t=this.closest(\'.other-svc-wrapper\').querySelector(\'.other-svc-body\');'
                        'if(t.style.display===\'none\'){t.style.display=\'\';this.textContent=\'&#9650; Collapse\';}else{t.style.display=\'none\';this.textContent=\'&#9660; Expand\';}">'
                        '&#9660; Expand</button></div>'
                        '<div class="other-svc-body" style="display:none">\n'
                        '<table><tr><th>Display Name</th><th>Name</th><th>Start Mode</th><th>Account</th></tr>\n')
            for i, s in enumerate(svcs_c):
                bg = ' style="background:#f5f4f8"' if i % 2 == 1 else ''
                svc_body += (f'<tr{bg}><td>{h(s.get("DisplayName",""))}</td>'
                            f'<td style="font-family:monospace;font-size:8.5pt">{h(s.get("Name",""))}</td>'
                            f'<td>{h(s.get("StartMode",""))}</td>'
                            f'<td style="font-size:8.5pt">{h(s.get("StartName",""))}</td></tr>\n')
            svc_body += '</table>\n</div></div>\n'
        else:
            svc_body += '<table><tr><th>Display Name</th><th>Name</th><th>Start Mode</th><th>Account</th></tr>\n'
            for i, s in enumerate(svcs_c):
                bg = ' style="background:#f5f4f8"' if i % 2 == 1 else ''
                svc_body += (f'<tr{bg}><td>{h(s.get("DisplayName",""))}</td>'
                            f'<td style="font-family:monospace;font-size:8.5pt">{h(s.get("Name",""))}</td>'
                            f'<td>{h(s.get("StartMode",""))}</td>'
                            f'<td style="font-size:8.5pt">{h(s.get("StartName",""))}</td></tr>\n')
            svc_body += '</table>\n'
    svc_body += top_link(sid)

    # ── SERVICE ANOMALIES CARD ────────────────────────────────────────────────
    anom_body = ''
    if svc_anomalies:
        anom_body  = ('<div class="flag-info" style="margin-bottom:14px"><div class="flag-label">Manual Review Required</div>'
                     '<div class="flag-detail">These services warrant manual review. Not all are malicious — confirm with client before taking action.</div></div>\n')
        anom_body += ('<table style="width:100%;table-layout:fixed;border-collapse:collapse;font-size:8.5pt">'
                     '<colgroup><col style="width:18%"><col style="width:22%"><col style="width:14%"><col style="width:46%"></colgroup>'
                     '<tr><th style="padding:6px 8px;text-align:left">Name</th>'
                     '<th style="padding:6px 8px;text-align:left">Display Name</th>'
                     '<th style="padding:6px 8px;text-align:left">Account</th>'
                     '<th style="padding:6px 8px;text-align:left">Why Flagged</th></tr>\n')
        for i, sv in enumerate(svc_anomalies):
            anom_body += (f'<tr style="background:#fff8e1" title="{h(sv.get("Path",""))}">'
                         f'<td style="padding:5px 8px;font-family:monospace;word-break:break-all">{h(sv.get("Name",""))}</td>'
                         f'<td style="padding:5px 8px;word-wrap:break-word">{h(sv.get("DisplayName",""))}</td>'
                         f'<td style="padding:5px 8px;word-break:break-all">{h(sv.get("StartName",""))}</td>'
                         f'<td style="padding:5px 8px;word-break:break-all;color:#555">{h(sv.get("_reason",""))}</td></tr>\n')
        anom_body += '</table>\n' + top_link(sid)

    # ── FILE SHARES CARD ──────────────────────────────────────────────────────
    shares_body = ''
    if sl:
        shares_body = '<table style="font-size:9pt;border-collapse:collapse;width:100%">\n'
        shares_body += ('<tr style="background:#271e41">'
                       '<th style="padding:7px 12px;color:#fff;text-align:left">Share</th>'
                       '<th style="padding:7px 12px;color:#fff;text-align:left">Path</th>'
                       '<th style="padding:7px 12px;color:#fff;text-align:left">Flag</th></tr>\n')
        for i, s in enumerate(sl):
            if not isinstance(s, dict): continue
            perms_raw = s.get('Permissions', [])
            # Handle both list and .NET-serialized dict (useless)
            perms = as_list(perms_raw) if not isinstance(perms_raw, dict) else []
            everyone_full = False
            for p in perms:
                if not isinstance(p, dict): continue
                acct = (p.get('AccountName') or '').lower()
                ar = p.get('AccessRight', '')
                ar_val = ar.get('Value', '') if isinstance(ar, dict) else str(ar)
                if acct == 'everyone' and ar_val.lower() in ('full', 'fullcontrol'):
                    everyone_full = True
                    break
            flag_cell = ('<span style="background:#fee2e2;color:#991b1b;font-size:8pt;font-weight:700;'
                         'padding:3px 10px;border-radius:10px;border:1px solid #fca5a5;">&#9888; Everyone: Full</span>'
                         if everyone_full else
                         '<span style="color:#065f46;font-size:9pt;">&#10003;</span>')
            bg = ' style="background:#f5f4f8"' if i % 2 == 1 else ''
            shares_body += (f'<tr{bg}><td style="padding:5px 8px;font-weight:600">{h(s.get("Name",""))}</td>'
                           f'<td style="padding:5px 8px;font-family:monospace;font-size:8.5pt">{h(s.get("Path",""))}</td>'
                           f'<td style="padding:5px 8px">{flag_cell}</td></tr>\n')
        shares_body += '</table>\n' + top_link(sid)

    # ── ASSEMBLE ──────────────────────────────────────────────────────────────
    tab_html  = sbr_html + nav_html
    tab_html += card(f'{sid}-alerts',     'Alerts',                  alerts_body)
    tab_html += card(f'{sid}-overview',   'System Overview',         overview_body)
    tab_html += card(f'{sid}-hardware',   'Hardware',                hw_body)
    tab_html += card(f'{sid}-apps',       'Installed Applications',  apps_body)
    tab_html += card(f'{sid}-roles',      'Roles &amp; Features',    roles_body, collapsed=True)
    tab_html += card(f'{sid}-roleconfig', 'Role Configuration',      rc)
    tab_html += card(f'{sid}-disks',      'Disk Storage',            disk_body)
    tab_html += card(f'{sid}-network',    'Network',                 net_body)
    tab_html += card(f'{sid}-lports',     'Listening Ports',         lp_body, collapsed=True)
    tab_html += card(f'{sid}-services',   'Services',                svc_body)
    if anom_body:
        tab_html += card(f'{sid}-svc-anomalies', 'Service Anomalies', anom_body)
    if shares_body:
        tab_html += card(f'{sid}-shares', f'File Shares ({len(sl)})', shares_body)

    return {'id': sid, 'name': name, 'role_label': rlabel,
            'in_scope': in_sc, 'crit': crit, 'warn': warn, 'tab_html': tab_html}

# ── HYPER-V HOSTS TAB ─────────────────────────────────────────────────────────
def build_hv_tab():
    if not hv_inventories:
        return '<div style="padding:24px;color:#6b6080;font-style:italic;">No Hyper-V inventory data collected.</div>'

    total_vcpu = 0; total_ram = 0; total_vms = 0; total_running = 0
    all_html = ''

    for hvi in hv_inventories:
        hs         = hvi.get('HostSummary', {})
        is_vsphere = hvi.get('_type') == 'vSphereInventory'
        hv_name    = hvi.get('HVHost', 'Unknown Host')
        host_ip    = hvi.get('HostIP', '')
        esxi_ver   = hvi.get('ESXiVersion', '')
        host_type  = hvi.get('HostType', 'Hyper-V')
        model      = hs.get('Model', '—')
        mfg        = hs.get('Manufacturer', '')
        cpu_model  = hs.get('CPUModel', '—')
        cpu_cores  = hs.get('CPUCores', '?')
        cpu_log    = hs.get('CPULogical', cpu_cores)
        ram_gb     = float(hs.get('TotalRAMgb', 0) or 0)
        vols       = hs.get('Volumes', [])
        datastores = hvi.get('Datastores', [])
        vms        = hvi.get('VMs', [])

        # Aggregate totals
        host_vcpu = sum(v.get('vCPU', 0) or 0 for v in vms)
        host_ram  = sum(float(v.get('RAMgb', 0) or 0) for v in vms)
        running   = sum(1 for v in vms if v.get('State', '') in ('Running', 'POWERED_ON'))
        total_vcpu += host_vcpu; total_ram += host_ram
        total_vms += len(vms); total_running += running

        # VM table
        vm_rows = ''
        for i, vm in enumerate(vms):
            disks     = [d for d in (vm.get('Disks', []) or []) if isinstance(d, dict)]
            disk_gb   = sum(float(d.get('SizeGB', 0) or 0) for d in disks)
            disk_used = sum(float(d.get('UsedGB', 0) or 0) for d in disks)
            disk_str  = f'{disk_gb:.0f} GB' + (f' / {disk_used:.0f} GB used' if disk_used else '')
            state     = vm.get('State', '?')
            sc        = 'green' if state in ('Running', 'POWERED_ON') else 'gray'
            snaps     = vm.get('Snapshots', 0) or 0
            ip_val    = vm.get('IP', '') or vm.get('IPs', '') or '—'
            guest_os  = vm.get('GuestOS', '')
            bg        = ' style="background:#f5f4f8"' if i % 2 == 1 else ''
            last_col  = (f'<td style="padding:6px 10px;font-size:8pt;color:#6b6080">{h(guest_os)}</td>'
                         if is_vsphere else
                         f'<td style="padding:6px 10px;text-align:center">{pill(str(snaps), "yellow" if snaps else "green")}</td>')
            vm_rows += (f'<tr{bg}>'
                       f'<td style="font-weight:600;padding:6px 10px">{h(vm.get("Name",""))}</td>'
                       f'<td style="padding:6px 10px">{pill(state, sc)}</td>'
                       f'<td style="padding:6px 10px;text-align:center">{vm.get("vCPU","?")}</td>'
                       f'<td style="padding:6px 10px;text-align:center">{float(vm.get("RAMgb",0) or 0):.0f} GB</td>'
                       f'<td style="padding:6px 10px;font-family:monospace;font-size:8.5pt">{h(str(ip_val))}</td>'
                       f'<td style="padding:6px 10px;text-align:center">{disk_str}</td>'
                       f'{last_col}'
                       f'</tr>\n')

        last_col_hdr = '<th style="padding:7px 10px;color:#fff">Guest OS</th>' if is_vsphere else '<th style="padding:7px 10px;color:#fff">Snaps</th>'

        # Storage volumes (HyperV) or Datastores (vSphere)
        vol_rows = ''
        if is_vsphere:
            for ds in datastores:
                cap  = float(ds.get('CapacityGB', 0) or 0)
                free = float(ds.get('FreeGB', 0) or 0)
                prov = float(ds.get('ProvisionedGB', 0) or 0)
                pct  = round((cap - free) / cap * 100, 1) if cap else 0
                pc   = 'red' if pct >= 85 else ('yellow' if pct >= 70 else 'green')
                vol_rows += (f'<tr><td style="padding:5px 10px;font-weight:600">{h(ds.get("Name",""))}</td>'
                            f'<td style="padding:5px 10px;font-size:8.5pt;color:#6b6080">{h(ds.get("Type",""))} &middot; {h(ds.get("DriveType",""))}</td>'
                            f'<td style="padding:5px 10px">{cap:.0f} GB</td>'
                            f'<td style="padding:5px 10px">{prov:.0f} GB</td>'
                            f'<td style="padding:5px 10px">{free:.0f} GB free</td>'
                            f'<td style="padding:5px 10px">{pill(f"{pct:.0f}%", pc)}{disk_bar(pct)}</td></tr>\n')
        else:
            for vol in vols:
                pct = float(vol.get('UsedPct', 0) or 0)
                pc  = 'red' if pct >= 85 else ('yellow' if pct >= 70 else 'green')
                vol_rows += (f'<tr><td style="padding:5px 10px;font-family:monospace">{h(vol.get("Drive",""))}</td>'
                            f'<td style="padding:5px 10px;font-size:8.5pt;color:#6b6080">{h(vol.get("Label",""))}</td>'
                            f'<td style="padding:5px 10px">{float(vol.get("TotalGB",0)):.0f} GB</td>'
                            f'<td style="padding:5px 10px">{float(vol.get("FreeGB",0)):.0f} GB free</td>'
                            f'<td style="padding:5px 10px">{pill(f"{pct:.0f}%", pc)}{disk_bar(pct)}</td></tr>\n')

        hv_anchor   = hv_name.lower().replace('.', '').replace('-', '')
        host_badge  = f'ESXi {esxi_ver}' if is_vsphere else 'Hyper-V'
        host_sub    = (f'{h(mfg)} {h(model)} &middot; {h(cpu_model)} &middot; {cpu_cores} cores &middot; {h(host_ip)}'
                       if is_vsphere else
                       f'{h(mfg)} {h(model)} &middot; {h(cpu_model)} &middot; {cpu_cores} cores / {cpu_log} logical &middot; {ram_gb:.0f} GB RAM')
        storage_label = 'Datastores' if is_vsphere else 'Host Storage Volumes'
        storage_hdr   = ('''<tr style="background:#ede9fe"><th style="padding:5px 10px;text-align:left;font-size:8pt">Datastore</th><th style="padding:5px 10px;font-size:8pt">Type</th><th style="padding:5px 10px;font-size:8pt">Capacity</th><th style="padding:5px 10px;font-size:8pt">Provisioned</th><th style="padding:5px 10px;font-size:8pt">Free</th><th style="padding:5px 10px;font-size:8pt">Usage</th></tr>'''
                         if is_vsphere else
                         '''<tr style="background:#ede9fe"><th style="padding:5px 10px;text-align:left;font-size:8pt">Drive</th><th style="padding:5px 10px;font-size:8pt">Label</th><th style="padding:5px 10px;font-size:8pt">Total</th><th style="padding:5px 10px;font-size:8pt">Free</th><th style="padding:5px 10px;font-size:8pt">Usage</th></tr>''')
        host_html = f'''
<div id="hv-host-{hv_anchor}" style="background:#f5f4f8;border-radius:8px;padding:16px 20px;margin-bottom:20px;border:1px solid #e0daf0;">
<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;">
  <div>
    <div style="font-size:13pt;font-weight:700;color:#271e41;">{h(hv_name)} <span style="font-size:8.5pt;font-weight:400;background:#ede9fe;color:#5b1fa4;padding:2px 8px;border-radius:10px;vertical-align:middle">{host_badge}</span></div>
    <div style="font-size:9pt;color:#6b6080;margin-top:2px;">{host_sub}</div>
  </div>
  <div style="text-align:right;">
    <div style="font-size:9pt;font-weight:700;color:#5b1fa4;">{len(vms)} VMs &middot; {running} running</div>
    <div style="font-size:8.5pt;color:#6b6080">{host_vcpu} vCPU allocated &middot; {host_ram:.0f} GB RAM allocated</div>
  </div>
</div>
<table style="width:100%;font-size:9pt;border-collapse:collapse;margin-bottom:14px;">
<tr style="background:#271e41">
  <th style="padding:7px 10px;color:#fff;text-align:left">VM Name</th>
  <th style="padding:7px 10px;color:#fff">State</th>
  <th style="padding:7px 10px;color:#fff">vCPU</th>
  <th style="padding:7px 10px;color:#fff">RAM</th>
  <th style="padding:7px 10px;color:#fff">IP</th>
  <th style="padding:7px 10px;color:#fff">Disk (Size)</th>
  {last_col_hdr}
</tr>
{vm_rows if vm_rows else '<tr><td colspan="7" style="padding:10px;color:#6b6080;font-style:italic;">No VMs found on this host</td></tr>'}
</table>
''' + (f'''<div style="font-size:8.5pt;font-weight:700;color:#5b1fa4;margin-bottom:6px;text-transform:uppercase;letter-spacing:.4px;">{storage_label}</div>
<table style="width:100%;font-size:9pt;border-collapse:collapse;">
{storage_hdr}
{vol_rows}
</table>''' if vol_rows else '') + \
'<div style="text-align:right;margin-top:10px;padding-top:8px;border-top:1px solid #e0daf0;">' \
'<a href="#top-virt" style="color:#5b1fa4;font-size:8.5pt;text-decoration:none;font-weight:600;">&#8593; Top</a>' \
'</div>\n</div>\n'

        all_html += host_html

    # Summary stats card
    summary_card = f'''<div class="stat-grid" style="margin-bottom:18px;">
<div class="stat-box"><div class="stat-num">{len(hv_inventories)}</div><div class="stat-lbl">Virtualization Hosts</div></div>
<div class="stat-box"><div class="stat-num">{total_vms}</div><div class="stat-lbl">Total VMs</div></div>
<div class="stat-box"><div class="stat-num">{total_vcpu}</div><div class="stat-lbl">vCPUs Allocated</div></div>
<div class="stat-box"><div class="stat-num">{total_ram:.0f} GB</div><div class="stat-lbl">RAM Allocated</div></div>
</div>
<div class="flag-info" style="margin-bottom:18px;">
  <div class="flag-label">Cloud Sizing Inputs</div>
  <div class="flag-detail">Total across all Hyper-V hosts: <strong>{total_vcpu} vCPU</strong> allocated &middot; <strong>{total_ram:.0f} GB RAM</strong> allocated &middot; <strong>{total_running} of {total_vms} VMs</strong> running. These are allocated figures — right-size based on actual utilization before quoting.</div>
</div>'''

    sbr_html = f'''<div class="sbr-only">
<div style="background:linear-gradient(135deg,#5b1fa4,#3d1270);border-radius:10px 10px 0 0;padding:16px 24px;display:flex;justify-content:space-between;align-items:center;margin-bottom:0;">
  <div>
    <div style="font-size:18px;font-weight:700;color:#fff;letter-spacing:.3px;">{('Hyper-V / ESX' if any(x.get('_type')=='HyperVInventory' for x in hv_inventories) and any(x.get('_type')=='vSphereInventory' for x in hv_inventories) else 'ESX Host Inventory' if any(x.get('_type')=='vSphereInventory' for x in hv_inventories) else 'Hyper-V Host Inventory')} </div>
    <div style="font-size:9pt;color:rgba(255,255,255,.85);margin-top:3px;">{len(hv_inventories)} hosts &middot; {total_vms} VMs &middot; {total_vcpu} vCPU &middot; {total_ram:.0f} GB RAM allocated</div>
  </div>
  <span style="background:rgba(255,255,255,.22);color:#fff;font-size:10pt;font-weight:700;padding:6px 18px;border-radius:20px;border:1.5px solid rgba(255,255,255,.5);">{len(hv_inventories)} Host(s)</span>
</div>
<div style="background:white;border-radius:0 0 10px 10px;border:1px solid #e8e4f0;border-top:none;box-shadow:0 4px 14px rgba(0,0,0,.07);padding:20px 24px;margin-bottom:16px;">
{summary_card}
</div></div>'''

    # Build host anchor nav pills
    hv_nav = '<div id="hv-top-nav" style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:18px;">\n'
    for hvi2 in hv_inventories:
        hv2_name = hvi2.get('HVHost', 'Unknown Host')
        hv2_anchor = hv2_name.lower().replace('.', '').replace('-', '')
        hv_nav += (f'<a href="#hv-host-{hv2_anchor}" onclick="document.getElementById(\'hv-host-{hv2_anchor}\').scrollIntoView({{behavior:\'smooth\',block:\'start\'}});return false;" style="color:#5b1fa4;font-size:9pt;text-decoration:none;'
                   f'font-weight:700;padding:5px 16px;border-radius:20px;background:#ede9fe;'
                   f'border:1.5px solid #c4b5fd;">{h(hv2_name)}</a>\n')
    hv_nav += '</div>\n'

    top_anchor = '<div id="top-virt"></div>\n'
    _has_hv2 = any(x.get('_type') == 'HyperVInventory'  for x in hv_inventories)
    _has_vs2 = any(x.get('_type') == 'vSphereInventory' for x in hv_inventories)
    _card_title = ('Hyper-V / ESX Summary' if _has_hv2 and _has_vs2
                   else 'ESX Host Summary'    if _has_vs2
                   else 'Hyper-V Host Summary')
    return top_anchor + sbr_html + card('hv-summary', f'{_card_title} ({len(hv_inventories)} Hosts)', hv_nav + summary_card + all_html)

# ── SQL TAB ───────────────────────────────────────────────────────────────────
_DB_VENDORS = [
    ('solarwinds',   'SolarWinds'),
    ('wsus',         'Windows Server Update Services (WSUS)'),
    ('sharepoint',   'Microsoft SharePoint'),
    ('reportserver', 'SQL Reporting Services (SSRS)'),
    ('kiwi',         'Kiwiplan ERP'),
    ('amtech',       'AmTech ERP'),
    ('advantage',    'Advantage Software'),
    ('netsuite',     'NetSuite'),
    ('quickbooks',   'QuickBooks'),
    ('sage',         'Sage'),
    ('dynamics',     'Microsoft Dynamics'),
    ('navision',     'Dynamics NAV'),
    ('greatplains',  'Dynamics GP'),
    ('connectwise',  'ConnectWise'),
    ('labtech',      'ConnectWise Automate'),
    ('autotask',     'Autotask / Datto'),
    ('veeam',        'Veeam Backup'),
    ('kaseya',       'Kaseya'),
    ('halo',         'HaloPSA'),
    ('servicenow',   'ServiceNow'),
    ('adlumin',      'Adlumin MDR'),
    ('huntress',     'Huntress'),
    ('sccm',         'Microsoft SCCM'),
    ('configmgr',    'Microsoft SCCM'),
    ('wordpress',    'WordPress'),
    ('gitlab',       'GitLab'),
]

def _db_vendor(name):
    n = name.lower()
    for k, v in _DB_VENDORS:
        if k in n: return v
    return ''

def build_sql_tab():
    # Collect SQL data from all servers
    sql_servers = []
    for srv in servers:
        data = srv['data']
        sql_raw = data.get('SQL', {})
        if isinstance(sql_raw, list): sql_raw = sql_raw[0] if sql_raw else {}
        if not isinstance(sql_raw, dict): continue
        inst = sql_raw.get('Instances', {})
        if not isinstance(inst, dict) or not inst.get('InstanceName'): continue
        sql_servers.append({'server': srv['name'], 'ip': srv.get('ip',''), 'sql': sql_raw, 'inst': inst})

    if not sql_servers:
        return None

    total_dbs   = sum(len(s['inst'].get('Databases', []) or []) for s in sql_servers if isinstance(s['inst'].get('Databases'), list))
    total_data  = sum(
        sum(float(d.get('DataSizeMB', 0) or 0) for d in (s['inst'].get('Databases') or []) if isinstance(d, dict))
        for s in sql_servers
    )

    summary = (f'<div class="stat-grid" style="margin-bottom:20px;">'
               f'<div class="stat-box"><div class="stat-num">{len(sql_servers)}</div><div class="stat-lbl">SQL Servers</div></div>'
               f'<div class="stat-box"><div class="stat-num">{total_dbs}</div><div class="stat-lbl">Databases</div></div>'
               f'<div class="stat-box"><div class="stat-num">{total_data/1024:.1f} GB</div><div class="stat-lbl">Total Data Size</div></div>'
               f'</div>')

    body = '<div id="top-sql"></div>\n'

    # Nav pills
    nav = '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:18px;">\n'
    for s in sql_servers:
        anc = s['server'].lower().replace('-','').replace('.','')
        nav += (f'<a href="#sql-{anc}" onclick="document.getElementById(\'sql-{anc}\').scrollIntoView({{behavior:\'smooth\',block:\'start\'}});return false;" '
                f'style="color:#5b1fa4;font-size:9pt;text-decoration:none;font-weight:700;padding:5px 16px;'
                f'border-radius:20px;background:#ede9fe;border:1.5px solid #c4b5fd;">{h(s["server"])}</a>\n')
    nav += '</div>\n'

    for s in sql_servers:
        inst = s['inst']
        anc  = s['server'].lower().replace('-','').replace('.','')
        ver  = inst.get('Version','—'); ed = inst.get('Edition','—')
        eol  = inst.get('EOLDate','—'); eol_s = inst.get('EOLStatus','')
        svc  = inst.get('ServiceAccount','—')
        conn = inst.get('Connected', False)
        eol_color = 'red' if eol_s == 'EOL' else ('yellow' if eol_s and 'near' in eol_s.lower() else 'green')

        dbs = [d for d in (inst.get('Databases') or []) if isinstance(d, dict)]

        db_rows = ''
        for i, db in enumerate(dbs):
            name    = db.get('Name','')
            vendor  = _db_vendor(name)
            data_mb = float(db.get('DataSizeMB', 0) or 0)
            log_mb  = float(db.get('LogSizeMB', 0) or 0)
            state   = db.get('State','—')
            recov   = db.get('RecoveryModel','—')
            compat  = db.get('CompatLevel','—')
            backup  = db.get('LastFullBackup','—') or '—'
            sc      = 'green' if state == 'ONLINE' else 'yellow'
            bg      = ' style="background:#f5f4f8"' if i % 2 == 1 else ''
            data_disp = f'{data_mb/1024:.2f} GB' if data_mb >= 1024 else f'{data_mb:.0f} MB'
            log_disp  = f'{log_mb/1024:.2f} GB'  if log_mb  >= 1024 else f'{log_mb:.0f} MB'
            db_rows += (f'<tr{bg}>'
                        f'<td style="padding:6px 10px;font-weight:600;font-family:monospace;font-size:9pt">{h(name)}</td>'
                        f'<td style="padding:6px 10px;font-size:8.5pt;color:#5b1fa4">{h(vendor) if vendor else "<span style=\'color:#c4b5fd\'>—</span>"}</td>'
                        f'<td style="padding:6px 10px;text-align:right">{data_disp}</td>'
                        f'<td style="padding:6px 10px;text-align:right;color:#6b6080">{log_disp}</td>'
                        f'<td style="padding:6px 10px">{pill(state, sc)}</td>'
                        f'<td style="padding:6px 10px;font-size:8.5pt">{h(recov)}</td>'
                        f'<td style="padding:6px 10px;text-align:center;font-size:8.5pt">{h(str(compat))}</td>'
                        f'<td style="padding:6px 10px;font-size:8.5pt;color:#6b6080">{h(backup)}</td>'
                        f'</tr>\n')

        db_section = ''
        if not conn:
            db_section = '<div class="flag-warning"><div class="flag-label">Deep connect failed</div><div class="flag-detail">Could not connect to SQL instance — database list unavailable. Verify the discovery account has SQL login permissions.</div></div>\n'
        elif dbs:
            db_section = (f'<table style="width:100%;font-size:9pt;border-collapse:collapse;margin-top:14px;">'
                          f'<tr style="background:#271e41">'
                          f'<th style="padding:7px 10px;color:#fff;text-align:left">Database</th>'
                          f'<th style="padding:7px 10px;color:#fff;text-align:left">Purpose / Vendor</th>'
                          f'<th style="padding:7px 10px;color:#fff;text-align:right">Data</th>'
                          f'<th style="padding:7px 10px;color:#fff;text-align:right">Log</th>'
                          f'<th style="padding:7px 10px;color:#fff">State</th>'
                          f'<th style="padding:7px 10px;color:#fff">Recovery</th>'
                          f'<th style="padding:7px 10px;color:#fff;text-align:center">Compat</th>'
                          f'<th style="padding:7px 10px;color:#fff">Last Full Backup</th>'
                          f'</tr>\n{db_rows}</table>\n')
        else:
            db_section = '<div style="color:#6b6080;font-style:italic;margin-top:10px;">No user databases found.</div>\n'

        top_lnk = '<div style="text-align:right;margin-top:10px;"><a href="#top-sql" style="color:#5b1fa4;font-size:8.5pt;text-decoration:none;font-weight:600;">&#8593; Top</a></div>\n'

        body += (f'<div id="sql-{anc}" style="background:#f5f4f8;border-radius:8px;padding:16px 20px;'
                 f'margin-bottom:20px;border:1px solid #e0daf0;">\n'
                 f'<div style="display:flex;justify-content:space-between;align-items:start;margin-bottom:12px;">\n'
                 f'<div><div style="font-size:13pt;font-weight:700;color:#271e41;">{h(s["server"])}</div>'
                 f'<div style="font-size:9pt;color:#6b6080;margin-top:2px;">{h(s["ip"])}</div></div>\n'
                 f'<div style="text-align:right;">'
                 f'<div style="font-size:9pt;font-weight:700;color:#5b1fa4;">{h(ed)}</div>'
                 f'<div style="font-size:8.5pt;color:#6b6080;">v{h(ver)} &middot; EOL {h(eol)} '
                 f'<span class="pill pill-{eol_color}" style="font-size:7.5pt;">{h(eol_s) or "—"}</span></div>'
                 f'<div style="font-size:8.5pt;color:#6b6080;margin-top:4px;">Service: {h(svc)}</div>'
                 f'</div></div>\n'
                 f'{db_section}{top_lnk}</div>\n')

    return card('sql-all', f'SQL Server Inventory ({len(sql_servers)} Instance{"s" if len(sql_servers)>1 else ""})',
                nav + summary + body)

# ── BUILD ALL TABS ─────────────────────────────────────────────────────────────
tabs = [(build_linux_tab(s) if s.get('os_type') == 'linux' else build_server_tab(s)) for s in servers]
virt_tab_html = build_hv_tab()
sql_tab_html  = build_sql_tab()

# ── LOGO ──────────────────────────────────────────────────────────────────────
if LOGO_B64:
    src = LOGO_B64 if LOGO_B64.startswith('data:') else f'data:image/png;base64,{LOGO_B64}'
    logo_html = f'<img src="{src}" style="height:32px;" alt="Magna5" class="logo-img">'
else:
    logo_html = '<span style="font-size:14pt;font-weight:700;color:white;letter-spacing:1px;">MAGNA5</span>'

# ── TABS HTML ─────────────────────────────────────────────────────────────────
tab_buttons = ''
for t in tabs:
    is_linux = servers[[s['id'] for s in servers].index(t['id'])].get('os_type') == 'linux' if t['id'] in [s['id'] for s in servers] else False
    cls = 'tab-btn' + (' is-linux' if is_linux else tab_cls(t['crit'], t['warn']))
    scope_ind = '' if t['in_scope'] else ' &#9702;'
    lx_icon = ' 🐧' if is_linux else ''
    tab_buttons += f'<button class="{cls}" data-tab="tab-{t["id"]}" onclick="showTab(\'{t["id"]}\')">{h(t["name"])}{lx_icon}{scope_ind}</button>\n'
_has_hyperv  = any(h.get('_type') == 'HyperVInventory'   for h in hv_inventories)
_has_vsphere = any(h.get('_type') == 'vSphereInventory'  for h in hv_inventories)
_virt_label  = ('Hyper-V / ESX' if _has_hyperv and _has_vsphere
                else 'ESX Hosts'    if _has_vsphere
                else 'Hyper-V Hosts')
tab_buttons += f'<button class="tab-btn" data-tab="tab-virt" onclick="showTab(\'virt\')">{_virt_label}</button>\n'
if sql_tab_html:
    tab_buttons += '<button class="tab-btn" data-tab="tab-sql" onclick="showTab(\'sql\')">SQL</button>\n'

tab_contents = ''
for i, t in enumerate(tabs):
    active = ' active' if i == 0 else ''
    tab_contents += f'<div id="tab-{t["id"]}" class="tab-content{active}">\n{t["tab_html"]}\n</div>\n'
tab_contents += f'<div id="tab-virt" class="tab-content">\n{virt_tab_html}\n</div>\n'
if sql_tab_html:
    tab_contents += f'<div id="tab-sql" class="tab-content">\n{sql_tab_html}\n</div>\n'

# ── FULL HTML ─────────────────────────────────────────────────────────────────
html_out = f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{h(CLIENT_FULL)} Server Discovery Report &mdash; {DATE}</title>
<style>
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
html {{ scroll-padding-top: 100px; }}
body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #f5f4f8; color: #271e41; font-size: 10pt; }}
.wrap {{ max-width: 1040px; margin: 0 auto; padding: 20px; }}
.tab-nav {{ display: flex; gap: 4px; margin-bottom: -1px; flex-wrap: wrap; }}
.tab-btn {{ padding: 8px 18px; background: #ddd9ee; border: 1px solid #c0b8d8; border-bottom: none; border-radius: 6px 6px 0 0; cursor: pointer; font-size: 9.5pt; color: #271e41; font-weight: 600; }}
.tab-btn.active {{ background: white; border-bottom: 1px solid white; color: #5b1fa4; }}
.tab-btn.has-critical {{ border-top: 3px solid #d63638; }}
.tab-btn.has-warning  {{ border-top: 3px solid #f5a623; }}
.tab-btn.is-linux     {{ border-top: 3px solid #0d9488; color: #0f4c5c; }}
.tab-content {{ display: none; }}
.tab-content.active {{ display: block; }}
.card {{ background: white; border-radius: 0 8px 8px 8px; padding: 24px; margin-bottom: 18px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); border: 1px solid #e8e4f0; }}
.card-title {{ font-size: 15px; font-weight: 700; color: #271e41; text-transform: uppercase; letter-spacing: 1px; border-bottom: 2px solid #5b1fa4; padding-bottom: 8px; margin-bottom: 16px; display: flex; justify-content: space-between; align-items: center; }}
.collapse-btn {{ background: none; border: 1px solid #c4b5fd; border-radius: 4px; color: #5b1fa4; font-size: 8pt; padding: 2px 8px; cursor: pointer; font-weight: 600; flex-shrink: 0; }}
.card-body {{ }}
.card-body.collapsed {{ display: none; }}
table {{ width: 100%; border-collapse: collapse; font-size: 9.5pt; }}
th {{ background: #271e41; color: #fff; font-weight: 600; padding: 7px 12px; text-align: left; }}
td {{ padding: 6px 12px; border: 1px solid #d0cce0; vertical-align: top; }}
tr:nth-child(even) td {{ background: #f5f4f8; }}
.flag-critical {{ background: #fff0f0; border-left: 4px solid #d63638; border-radius: 0 6px 6px 0; padding: 12px 16px; margin-bottom: 8px; }}
.flag-warning  {{ background: #fff8e1; border-left: 4px solid #f5a623; border-radius: 0 6px 6px 0; padding: 12px 16px; margin-bottom: 8px; }}
.flag-info     {{ background: #f0f4ff; border-left: 4px solid #5b1fa4; border-radius: 0 6px 6px 0; padding: 12px 16px; margin-bottom: 8px; }}
.flag-ok       {{ background: #f0fdf0; border-left: 4px solid #20c800; border-radius: 0 6px 6px 0; padding: 12px 16px; margin-bottom: 8px; }}
.flag-label    {{ font-size: 8.5pt; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }}
.flag-critical .flag-label {{ color: #d63638; }}
.flag-warning  .flag-label {{ color: #b07a00; }}
.flag-info     .flag-label {{ color: #5b1fa4; }}
.flag-ok       .flag-label {{ color: #20c800; }}
.flag-detail   {{ font-size: 9.5pt; color: #271e41; margin-top: 4px; }}
.pill {{ display: inline-block; padding: 2px 9px; border-radius: 12px; font-size: 8pt; font-weight: 700; }}
.pill-red    {{ background: #fee2e2; color: #991b1b; }}
.pill-yellow {{ background: #fef3c7; color: #92400e; }}
.pill-green  {{ background: #d1fae5; color: #065f46; }}
.pill-gray   {{ background: #f3f4f6; color: #374151; }}
.pill-purple {{ background: #ede9fe; color: #5b1fa4; }}
.stat-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 16px; }}
.stat-box {{ background: #f5f4f8; border-radius: 6px; padding: 12px; text-align: center; }}
.stat-num  {{ font-size: 22px; font-weight: 700; color: #5b1fa4; }}
.stat-lbl  {{ font-size: 8pt; color: #6b6080; text-transform: uppercase; letter-spacing: 0.5px; margin-top: 2px; }}
.role-grid {{ display: flex; flex-wrap: wrap; gap: 8px; }}
.role-badge {{ background: #ede9fe; color: #5b1fa4; border-radius: 4px; padding: 4px 12px; font-size: 9pt; font-weight: 600; border: 1px solid #c4b5fd; }}
.disk-bar-bg  {{ background: #e9e4f5; border-radius: 4px; height: 10px; width: 100%; margin-top: 4px; }}
.disk-bar-fill {{ height: 10px; border-radius: 4px; }}
.sub-title {{ font-size: 13px; font-weight: 700; color: #5b1fa4; text-transform: uppercase; letter-spacing: 0.8px; margin: 16px 0 8px; }}
.meta-line {{ font-size: 8.5pt; color: #6b6080; margin-bottom: 2px; }}
.sbr-only  {{ display: none !important; }}
body.view-sbr .sbr-only {{ display: block !important; }}
body.view-sbr .hide-sbr {{ display: none !important; }}
.view-bar  {{ background: #1a1432; padding: 8px 28px; display: flex; align-items: center; gap: 10px; position: sticky; top: 52px; z-index: 199; border-radius: 0 0 6px 6px; border-top: 1px solid rgba(255,255,255,.08); }}
.view-lbl  {{ font-size: 8pt; color: #a89bc8; text-transform: uppercase; letter-spacing: .5px; font-weight: 700; }}
.view-btn  {{ background: transparent; border: 1px solid #4a3a6a; border-radius: 4px; color: #c4b5fd; font-size: 9pt; padding: 4px 14px; cursor: pointer; font-weight: 600; transition: all .15s; }}
.view-btn:hover {{ background: #2d2060; }}
.view-btn.v-active {{ background: #5b1fa4; color: white; border-color: #5b1fa4; }}
.view-desc {{ font-size: 8.5pt; color: #7c6b9e; margin-left: 6px; }}
details summary {{ cursor: pointer; font-weight: 600; color: #5b1fa4; padding: 6px 0; list-style: none; }}
details summary::before {{ content: '\\25B6  '; font-size: 9pt; }}
details[open] summary::before {{ content: '\\25BC  '; }}
</style>
</head>
<body class="view-adv">
<div style="background:#1a1432;padding:12px 28px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:200;">
  <div style="display:flex;align-items:center;gap:16px;">
    {logo_html}
    <div style="color:rgba(255,255,255,.5);font-size:10pt;">|</div>
    <div style="color:white;font-size:10pt;font-weight:600;">{h(CLIENT_FULL)}</div>
    <div style="color:rgba(255,255,255,.5);font-size:10pt;">|</div>
    <div style="color:rgba(255,255,255,.7);font-size:9pt;">Server Discovery Report &middot; {DATE}</div>
  </div>
  <div style="color:rgba(255,255,255,.5);font-size:8.5pt;">Collected by Magna5 SE Team</div>
</div>
<div class="view-bar">
  <span class="view-lbl">View:</span>
  <button class="view-btn v-active" id="vbtn-adv" onclick="setView('adv')">Advanced</button>
  <button class="view-btn" id="vbtn-sbr" onclick="setView('sbr')">SBR</button>
  <span class="view-desc" id="view-desc">Full technical detail &mdash; SE view</span>
</div>
<div class="wrap">
<div class="tab-nav">
{tab_buttons}
</div>
<div style="background:white;border-radius:0 8px 8px 8px;padding:0;border:1px solid #e8e4f0;margin-bottom:4px;">
{tab_contents}
</div>
</div>
<div style="text-align:center;padding:24px 0 32px;color:#a89bc8;font-size:8pt;letter-spacing:.3px;">
  Generated {DATE} ET &middot; Magna5 Solutions Engineering &middot; SDT v2.9
</div>
<script>
var VIEW_DESCS = {{
  basic: "Simplified view &mdash; SE &amp; CSM",
  adv:   "Full technical detail &mdash; SE view",
  sbr:   "Executive health dashboard &mdash; client &amp; leadership"
}};
function setView(v) {{
  document.body.classList.remove("view-adv","view-sbr");
  document.body.classList.add("view-" + v);
  document.querySelectorAll(".view-btn").forEach(function(b){{ b.classList.remove("v-active"); }});
  var btn = document.getElementById("vbtn-" + v);
  if (btn) btn.classList.add("v-active");
  var desc = document.getElementById("view-desc");
  if (desc) desc.innerHTML = VIEW_DESCS[v] || "";
  try {{ localStorage.setItem("sdView", v); }} catch(e) {{}}
}}
function showTab(id) {{
  document.querySelectorAll('.tab-content').forEach(function(el){{ el.classList.remove('active'); }});
  document.querySelectorAll('.tab-btn').forEach(function(el){{ el.classList.remove('active'); }});
  var tc = document.getElementById('tab-' + id);
  if (tc) tc.classList.add('active');
  var btn = document.querySelector('[data-tab="tab-' + id + '"]');
  if (btn) btn.classList.add('active');
}}
function toggleCard(btn) {{
  var body = btn.closest('.card').querySelector('.card-body');
  if (!body) return;
  body.classList.toggle('collapsed');
  btn.textContent = body.classList.contains('collapsed') ? '\u25bc Expand' : '\u25b2 Collapse';
}}
window.addEventListener('DOMContentLoaded', function() {{
  var saved = 'adv';
  try {{ saved = localStorage.getItem('sdView') || 'adv'; }} catch(e) {{}}
  setView(saved);
  var firstBtn = document.querySelector('.tab-btn');
  if (firstBtn) firstBtn.click();
}});
</script>
</body>
</html>
'''

with open(OUTPUT, 'w', encoding='utf-8') as f:
    f.write(html_out)

print(f"Report written to: {OUTPUT}")
total_crit = sum(t['crit'] for t in tabs)
total_warn = sum(t['warn'] for t in tabs)
print(f"Summary: {total_crit} critical flags, {total_warn} warnings across {len(tabs)} servers")
