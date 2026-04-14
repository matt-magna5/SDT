"""
gen_report_v2.py — Discovery Report Generator (v2 layout)
Reads JSON session files, produces HTML matching MEKE-DiscoveryReport-2026-04-09-v2.html layout exactly.
Usage: python gen_report_v2.py  (edit CONFIG section below for each client)
"""
import json, html as htmlmod, re, sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# ── CONFIG ────────────────────────────────────────────────────────────────────
BASE        = r'C:/Users/matt/OneDrive - Magna5/M5 Obsidian Vault/M5 Ops/Root/zzzTest/Server Discovery Tool'
SESSION_DIR = BASE + '/_output/Discovery-Session-2026-04-10-1924'
DATE        = '2026-04-10'
CLIENT      = 'MEKE'
OUTPUT      = SESSION_DIR + f'/{CLIENT}-DiscoveryReport-{DATE}.html'
LOGO_FILE   = r'C:/Users/matt/AppData/Local/Temp/m5_logo_b64.txt'

# ── LOAD DATA ─────────────────────────────────────────────────────────────────
def jload(path):
    with open(path, encoding='utf-8-sig') as f: return json.load(f)

hv  = jload(f'{SESSION_DIR}/MEKEHV02-discovery-{DATE}.json')
dc  = jload(f'{SESSION_DIR}/MEKEDC02-discovery-{DATE}.json')
inv = jload(f'{SESSION_DIR}/localhost-inventory-{DATE}.json')

with open(LOGO_FILE) as f: LOGO_B64 = f.read().strip()

# ── HELPERS ───────────────────────────────────────────────────────────────────
h = htmlmod.escape

def as_list(v):
    """Flatten potentially-nested PS serialization artefacts into a flat list of dicts."""
    if isinstance(v, dict): return [v]
    if isinstance(v, list):
        result = []
        for item in v:
            if isinstance(item, dict):
                result.append(item)
            elif isinstance(item, list):
                result.extend(x for x in item if isinstance(x, dict))
        return result
    return []

def pill(text, color):
    return f'<span class="pill pill-{color}">{h(str(text))}</span>'

def flag_div(sev, title, detail):
    return (f'<div class="flag-{sev}">'
            f'<div class="flag-label">{h(title)}</div>'
            f'<div class="flag-detail">{h(detail)}</div></div>\n')

def disk_bar(pct):
    c = '#d63638' if pct>=85 else ('#f5a623' if pct>=70 else '#20c800')
    return (f'<div class="disk-bar-bg"><div class="disk-bar-fill" '
            f'style="width:{min(pct,100)}%;background:{c}"></div></div>')

def card(cid, title, body_html, extra_class='hide-sbr', collapsed=False):
    btn = '&#9660; Expand' if collapsed else '&#9650; Collapse'
    bc  = 'card-body collapsed' if collapsed else 'card-body'
    ec  = f' {extra_class}' if extra_class else ''
    return (f'<div class="card{ec}" id="{cid}">\n'
            f'<div class="card-title"><span>{title}</span>'
            f'<button class="collapse-btn" onclick="toggleCard(this)">{btn}</button></div>\n'
            f'<div class="{bc}">\n{body_html}\n</div>\n</div>\n')

def top_link(tid):
    return (f'<div style="text-align:right;margin-top:10px;">'
            f'<a href="#top-{tid}" style="color:#5b1fa4;font-size:8.5pt;'
            f'text-decoration:none;font-weight:600;">&#8593; Top</a></div>\n')

def nav_link(anchor, label):
    return (f'<a href="#{anchor}" style="color:#5b1fa4;font-size:9pt;text-decoration:none;'
            f'font-weight:600;padding:4px 12px;border-radius:4px;background:#ede9fe;'
            f'border:1px solid #c4b5fd;">{label}</a>\n')

def nav_link_dark(anchor, label):
    return (f'<a href="#{anchor}" style="color:white;font-size:9pt;text-decoration:none;'
            f'font-weight:600;padding:4px 12px;border-radius:4px;background:#5b1fa4;">'
            f'{label} &#8595;</a>\n')

def sub(t): return f'<div class="sub-title">{t}</div>\n'

def mini_box(title, content, last=False):
    mb = '0' if last else '14px'
    return (f'<div style="background:#f5f4f8;border-radius:8px;padding:14px 16px;margin-bottom:{mb};">'
            f'<div style="font-size:8.5pt;font-weight:700;text-transform:uppercase;letter-spacing:.5px;'
            f'color:#5b1fa4;border-bottom:1.5px solid #ede9fe;padding-bottom:6px;margin-bottom:12px;">'
            f'{title}</div>{content}</div>\n')

# ── FLAG DERIVATION ───────────────────────────────────────────────────────────
hv_sys = hv.get('System', {})
hv_hw  = hv.get('Hardware', {})
dc_sys = dc.get('System', {})
dc_hw  = dc.get('Hardware', {})

hv_disks_raw = hv.get('Disks', [])
hv_disks = as_list(hv_disks_raw)

dc_ad_raw = dc.get('AD', [])
dc_ad = next((x for x in as_list(dc_ad_raw) if 'ForestName' in x), {})

dc_roles_data = dc.get('Roles', {})
feat_names_dc = [f.get('Name','') for f in as_list(dc_roles_data.get('InstalledFeatures',[]))]
has_smb1 = any('FS-SMB1' in n for n in feat_names_dc)

stale_users = as_list(dc_ad.get('StaleUsers', []))
stale_comps = as_list(dc_ad.get('StaleComputers', []))

# Inventory — needed by flag functions
inv_vms  = as_list(inv.get('VMs', {})) or ([inv['VMs']] if isinstance(inv.get('VMs'), dict) else [])

# HV02 flags
def build_hv_flags():
    flags = []
    if hv_sys.get('Domain','').upper() == 'WORKGROUP':
        flags.append(('warning','Not Domain-Joined',
            'MEKEHV02 is in WORKGROUP — not joined to MCA.mekeel.org. GPO, domain security policies, '
            'and centralized auditing do not apply to this host.'))
    up = hv_sys.get('UptimeDays', 0)
    if up > 300:
        flags.append(('warning','Extended Uptime — Reboot Required',
            f'{up:.0f} days since last reboot. Pending Windows updates and driver patches will not '
            f'apply until the host is rebooted. Schedule a maintenance window.'))
    for d in hv_disks:
        pct = d.get('UsedPct', 0)
        drv, free, total = d.get('Drive','?'), d.get('FreeGB',0), d.get('TotalGB',0)
        if pct >= 85:
            flags.append(('critical',f'Disk {drv} Near Capacity',
                f'{drv}: {pct}% used — only {free:.1f} GB free of {total:.1f} GB. '
                f'Risk of service disruption.'))
        elif pct >= 70:
            flags.append(('warning',f'Disk {drv} Space Moderate',
                f'{drv}: {pct}% used ({free:.1f} GB free of {total:.1f} GB). Monitor closely.'))
    return flags

# DC02 flags
def build_dc_flags():
    flags = []
    if has_smb1:
        flags.append(('critical','SMB 1.0/CIFS Enabled — Critical Security Risk',
            'SMB 1.0 is the attack vector used by WannaCry, NotPetya, and EternalBlue ransomware. '
            'Disable immediately: Remove-WindowsFeature FS-SMB1'))
    dc_c = dc_ad.get('DCCount', {})
    if isinstance(dc_c, dict) and dc_c.get('Count', 0) == 0:
        flags.append(('critical','Single Domain Controller — No Redundancy',
            'Only one DC detected (MEKEDC02 on Hyper-V). AD, DNS, DHCP, and NPS are all single '
            'points of failure. Any outage takes down the entire domain.'))
    ram_free  = dc_hw.get('RAMAvailGB', 999)
    ram_total = dc_hw.get('RAMTotalGB', 1)
    if ram_total > 0 and (1 - ram_free/ram_total) > 0.75:
        pct = int((1-ram_free/ram_total)*100)
        flags.append(('warning','High Memory Utilization',
            f'RAM at {pct}% used ({ram_free:.1f} GB free of {ram_total:.1f} GB). '
            f'DC running AD, DNS, DHCP, NPS, and Print has minimal memory headroom.'))
    if len(stale_users) > 20:
        flags.append(('warning',f'Stale AD Accounts — {len(stale_users)} Inactive Users',
            f'{len(stale_users)} accounts inactive 90+ days. Attack surface risk. '
            f'Audit and disable/remove.'))
    domain_fl = dc_ad.get('DomainFL','')
    if domain_fl and str(domain_fl) < '2019':
        flags.append(('warning',f'AD Functional Level {domain_fl} — Upgrade Recommended',
            f'Domain functional level is {domain_fl}. Upgrading to 2019+ enables enhanced LAPS, '
            f'additional security features, and better Kerberos options.'))
    if len(stale_comps) > 20:
        flags.append(('warning',f'Stale Computer Accounts — {len(stale_comps)} Devices',
            f'{len(stale_comps)} computer accounts inactive 90+ days. Clean up to reduce '
            f'attack surface.'))
    # Check VM data disk usage from inventory
    for vm in inv_vms:
        for vd in as_list(vm.get('Disks', [])):
            size = vd.get('SizeGB', 0)
            used = vd.get('UsedGB', 0)
            pct  = int(used/size*100) if size else 0
            vname = vm.get('Name','?')
            fname = vd.get('Path','?').split('\\')[-1]
            if pct >= 95:
                flags.append(('critical',f'VM Disk at Capacity: {vname} / {fname}',
                    f'{fname} ({vd.get("VHDType","?")} VHD) is {pct}% full '
                    f'({used:.0f} GB / {size:.0f} GB). Disk writes will fail when full — '
                    f'expand or add storage before VM crashes.'))
    return flags

hv_flags = build_hv_flags()
dc_flags = build_dc_flags()
hv_crit = sum(1 for f in hv_flags if f[0]=='critical')
hv_warn = sum(1 for f in hv_flags if f[0]=='warning')
dc_crit = sum(1 for f in dc_flags if f[0]=='critical')
dc_warn = sum(1 for f in dc_flags if f[0]=='warning')

def tab_cls(crit, warn):
    if crit: return ' has-critical'
    if warn: return ' has-warning'
    return ''

def sbr_badge(crit, warn):
    if crit: return '🔴 CRITICAL'
    if warn: return '⚠️ ATTENTION'
    return '✅ HEALTHY'

def sbr_grad(crit, warn):
    if crit: return 'linear-gradient(135deg,#d63638,#b92b2e)'
    if warn: return 'linear-gradient(135deg,#f5a623,#e0901a)'
    return 'linear-gradient(135deg,#20c800,#158f00)'

# ── INVENTORY HELPERS ─────────────────────────────────────────────────────────
inv_vsw = as_list(inv.get('VirtualSwitches', {})) or ([inv['VirtualSwitches']] if isinstance(inv.get('VirtualSwitches'), dict) else [])
inv_host = inv.get('HostSummary', {})
inv_vols = as_list(inv_host.get('Volumes', []))

hv_apps = as_list(hv.get('Apps', []))
hv_edr = next((a['Name'] for a in hv_apps if 'sentinel' in a.get('Name','').lower()), None)
hv_edge = next((a['Version'] for a in hv_apps if 'Edge' in a.get('Name','')), None)

# DC apps — handle nested structure [garbage_str, [actual_apps]]
_dc_apps_raw = dc.get('Apps', [])
dc_apps = []
for item in _dc_apps_raw:
    if isinstance(item, list): dc_apps.extend([x for x in item if isinstance(x,dict)])
    elif isinstance(item, dict): dc_apps.append(item)
dc_edr = next((a['Name'] for a in dc_apps if 'sentinel' in a.get('Name','').lower()), None)
dc_rmm = next((a['Name'] for a in dc_apps if 'n-able' in a.get('Publisher','').lower() and 'Windows Agent' in a.get('Name','')), None)

dc_shares_raw = as_list(dc.get('FileShares',{}).get('Shares',[]))
dc_shares_real = [s for s in dc_shares_raw
                  if not s.get('Name','').startswith('$')
                  and s.get('Name','') not in ('ADMIN$','IPC$','C$','print$')]
dc_printers = [p for p in dc.get('Printers',[]) if isinstance(p,dict)]

dc_roles_list = as_list(dc_roles_data.get('InstalledRoles',[]))
dc_role_names = [r.get('DisplayName',r.get('Name','')) for r in dc_roles_list]
dc_features_list = as_list(dc_roles_data.get('InstalledFeatures',[]))

domain_name = dc_ad.get('ForestName', dc_sys.get('Domain','?'))

# Storage bar row
def stor_row(drv, total, pct):
    c   = '#d63638' if pct>=85 else ('#f5a623' if pct>=70 else '#5b1fa4')
    pc  = 'red' if pct>=85 else ('yellow' if pct>=70 else 'green')
    return (f'<tr style="padding:6px 0">'
            f'<td style="font-weight:700;white-space:nowrap">{h(drv)}</td>'
            f'<td style="font-size:8.5pt;color:#6b6080">{total:.1f} GB total</td>'
            f'<td>{pill(f"{pct}% used", pc)}'
            f'{disk_bar(pct)}</td></tr>\n')

# ── SBR: HV02 ─────────────────────────────────────────────────────────────────
up = hv_sys.get('UptimeDays', 0)
hv_ram_t = hv_hw.get('RAMTotalGB', 0)
hv_ram_f = hv_hw.get('RAMAvailGB', 0)
hv_ram_pct = int((1 - hv_ram_f/hv_ram_t)*100) if hv_ram_t else 0

hv_vol_rows = ''.join(stor_row(v['Drive'], v['TotalGB'], v['UsedPct']) for v in inv_vols)
if not hv_vol_rows:
    hv_vol_rows = ''.join(stor_row(d.get('Drive','?'), d.get('TotalGB',0), d.get('UsedPct',0)) for d in hv_disks)

vm_rows = ''
for vm in inv_vms:
    state = vm.get('State','?')
    sc = 'green' if state=='Running' else 'yellow'
    vcpu = vm.get('vCPU', vm.get('CPUCount','?'))
    ram  = vm.get('RAMgb', vm.get('MemoryGB', 0))
    vm_rows += (f'<tr><td style="padding:5px 8px">{h(vm.get("Name","?"))}</td>'
                f'<td style="padding:5px 8px">{pill(state, sc)}</td>'
                f'<td style="padding:5px 8px">{vcpu} vCPU &middot; {ram:.1f} GB RAM</td></tr>\n')

hv_agents = ''
if hv_edr:
    hv_agents += (f'<div style="display:flex;align-items:center;gap:10px;padding:7px 0;'
                  f'border-bottom:1px solid #f0edf8;">'
                  f'<span style="font-size:14px;flex-shrink:0">🛡</span>'
                  f'<div><div style="font-size:9pt;font-weight:600;color:#271e41;">EDR</div>'
                  f'<div style="font-size:8.5pt;color:#6b6080;">{h(hv_edr)}</div></div></div>\n')
nables = [a['Name'] for a in hv_apps if 'n-able' in a.get('Publisher','').lower()]
if nables:
    hv_agents += (f'<div style="display:flex;align-items:center;gap:10px;padding:7px 0;">'
                  f'<span style="font-size:14px;flex-shrink:0">🔧</span>'
                  f'<div><div style="font-size:9pt;font-weight:600;color:#271e41;">RMM</div>'
                  f'<div style="font-size:8.5pt;color:#6b6080;">{h(", ".join(nables[:2]))}</div></div></div>\n')

hv_sbr = f'''<div class="sbr-only">
<div style="background:{sbr_grad(hv_crit,hv_warn)};border-radius:10px 10px 0 0;padding:16px 24px;display:flex;justify-content:space-between;align-items:center;margin-bottom:0;">
  <div>
    <div style="font-size:18px;font-weight:700;color:#fff;letter-spacing:.3px;">MEKEHV02</div>
    <div style="font-size:9pt;color:rgba(255,255,255,.85);margin-top:3px;">Physical Hyper-V Host &middot; {h(hv_hw.get("Model",""))} &middot; Server 2022</div>
  </div>
  <div style="display:flex;align-items:center;gap:14px;">
    <div style="text-align:center;background:rgba(255,255,255,.{25 if hv_crit else 15});border-radius:8px;padding:8px 14px;">
      <div style="font-size:22px;font-weight:700;color:#fff;">{hv_crit}</div>
      <div style="font-size:7pt;color:rgba(255,255,255,.8);text-transform:uppercase;font-weight:600;">Critical</div>
    </div>
    <div style="text-align:center;background:rgba(255,255,255,.{25 if hv_warn else 15});border-radius:8px;padding:8px 14px;">
      <div style="font-size:22px;font-weight:700;color:#fff;">{hv_warn}</div>
      <div style="font-size:7pt;color:rgba(255,255,255,.8);text-transform:uppercase;font-weight:600;">Warning</div>
    </div>
    <span style="background:rgba(255,255,255,.22);color:#fff;font-size:10pt;font-weight:700;padding:6px 18px;border-radius:20px;border:1.5px solid rgba(255,255,255,.5);">{sbr_badge(hv_crit,hv_warn)}</span>
  </div>
</div>
<div style="background:white;border-radius:0 0 10px 10px;border:1px solid #e8e4f0;border-top:none;box-shadow:0 4px 14px rgba(0,0,0,.07);padding:20px 24px;margin-bottom:16px;">
<div style="display:grid;grid-template-columns:1fr 1fr;gap:18px;">
<div>
{mini_box("Hardware", f'''<div style="font-size:9.5pt;font-weight:700;color:#271e41;">{h(hv_hw.get("Model",""))}</div>
<div style="font-size:8.5pt;color:#6b6080;margin-top:3px;">{h(hv_hw.get("CPUName",""))} &middot; {h(str(hv_hw.get("CPUCores","")))} cores</div>
<div style="font-size:8.5pt;color:#6b6080;margin-top:2px;">RAM: {hv_ram_t:.1f} GB total &middot; {hv_ram_f:.1f} GB free ({100-hv_ram_pct}%)</div>
<div style="margin-top:8px;font-size:8.5pt;">S/N: <span style="font-family:monospace">{h(hv_hw.get("SerialNumber",""))}</span></div>''')}
{mini_box("OS &amp; System", f'''<table style="width:100%;font-size:9pt;border-collapse:collapse;">
<tr><td style="color:#6b6080;width:110px;padding:3px 0">OS</td><td>{h(hv_sys.get("OSName","").replace("Microsoft ",""))} {pill("Supported to 2031","green")}</td></tr>
<tr><td style="color:#6b6080;padding:3px 0">Last Reboot</td><td>{h(hv_sys.get("LastBoot","?").split(" ")[0])} &nbsp;{pill(f"{up:.0f} days ago — reboot recommended","yellow") if up>100 else pill(f"{up:.0f} days","green")}</td></tr>
<tr><td style="color:#6b6080;padding:3px 0">Deployed</td><td>{h(hv_sys.get("OSInstallDate",""))}</td></tr>
<tr><td style="color:#6b6080;padding:3px 0">Domain</td><td>{pill("WORKGROUP — not domain-joined","yellow") if hv_sys.get("Domain","").upper()=="WORKGROUP" else pill(hv_sys.get("Domain",""),"green")}</td></tr>
</table>''')}
{mini_box(f"Hosted VMs ({len(inv_vms)})", f'''<table style="width:100%;font-size:9pt;border-collapse:collapse;">
<tr style="background:#ede9fe"><th style="padding:5px 8px;text-align:left;font-size:8pt">VM</th><th style="padding:5px 8px;text-align:left;font-size:8pt">State</th><th style="padding:5px 8px;text-align:left;font-size:8pt">Resources</th></tr>
{vm_rows}</table>''', last=True)}
</div>
<div>
{mini_box("Installed Agents", hv_agents or '<span style="color:#9b8fb0;font-style:italic">None detected</span>')}
{mini_box("Applications", f'''<table style="width:100%;font-size:9pt;border-collapse:collapse;">
<tr><td style="color:#6b6080;width:110px;padding:4px 0;vertical-align:top">Security</td><td style="padding:4px 0">{h(hv_edr) if hv_edr else "None detected"}</td></tr>
<tr style="background:none"><td style="color:#6b6080;padding:4px 0;vertical-align:top">Management</td><td style="padding:4px 0">{h(", ".join(list(dict.fromkeys(a.get("Name","") for a in hv_apps if "n-able" in a.get("Publisher","").lower()))[:3]))}</td></tr>
<tr><td style="color:#6b6080;padding:4px 0;vertical-align:top">Browser</td><td style="padding:4px 0">{"Microsoft Edge " + h(hv_edge) if hv_edge else "None"}</td></tr>
<tr style="background:none"><td style="color:#6b6080;padding:4px 0;vertical-align:top">LOB Apps</td><td style="padding:4px 0;color:#9b8fb0;font-style:italic">None detected on host</td></tr>
<tr><td style="color:#6b6080;padding:4px 0;vertical-align:top">Risky Apps</td><td style="padding:4px 0">{pill("None flagged","green")}</td></tr>
</table>''')}
{mini_box("Storage", f'''<table style="width:100%;border-collapse:collapse;">
<tr><th style="text-align:left;font-size:8pt;padding:4px 0;color:#6b6080;font-weight:600">Drive</th><th style="font-size:8pt;color:#6b6080;font-weight:600">Size</th><th style="font-size:8pt;color:#6b6080;font-weight:600">Usage</th></tr>
{hv_vol_rows}</table>''', last=True)}
</div>
</div>
</div>
</div>
'''

# ── SBR: DC02 ─────────────────────────────────────────────────────────────────
dc_up = dc_sys.get('UptimeDays', 0)
dc_ram_t = dc_hw.get('RAMTotalGB', 0)
dc_ram_f = dc_hw.get('RAMAvailGB', 0)
dc_ram_pct_used = int((1-dc_ram_f/dc_ram_t)*100) if dc_ram_t else 0

dc_disks_list = as_list(dc.get('Disks', []))
dc_vol_rows = ''.join(stor_row(d.get('Drive','?'), d.get('TotalGB',0), d.get('UsedPct',0)) for d in dc_disks_list)

role_badges = ''.join(f'<span class="role-badge">{h(r)}</span>' for r in dc_role_names)
all_in_one_warn = ''
if len(dc_role_names) >= 4:
    all_in_one_warn = ('<div style="margin-top:10px;font-size:8.5pt;color:#92400e;background:#fff8e1;'
                       'border-radius:6px;padding:8px 12px;">⚠️ All critical network services run on a '
                       'single VM — any outage takes down AD, DNS, DHCP, and authentication simultaneously.</div>\n')

share_rows = ''
for s in dc_shares_real[:6]:
    sname = s.get('Name','')
    spath = s.get('Path','')
    share_rows += (f'<tr><td style="padding:5px 8px;font-weight:600">{h(sname)}</td>'
                   f'<td style="padding:5px 8px;font-family:monospace;font-size:8pt">{h(spath)}</td>'
                   f'<td style="padding:5px 8px"></td></tr>\n')

printer_rows = ''.join(
    f'<tr><td style="padding:3px 0">🖨 {h(p.get("Name",p.get("PrinterName","?")))}</td></tr>\n'
    for p in dc_printers[:6]
)

smb1_banner = ''
if has_smb1:
    smb1_banner = '''<div style="background:#fff0f0;border:1.5px solid #d63638;border-radius:8px;padding:12px 16px;margin-bottom:16px;display:flex;align-items:flex-start;gap:12px;">
<span style="font-size:18px;flex-shrink:0">⛔</span>
<div><div style="font-size:9.5pt;font-weight:700;color:#d63638;">SMB 1.0/CIFS ENABLED — Critical Security Risk</div>
<div style="font-size:9pt;color:#7f2424;margin-top:3px;">SMB 1.0 is the attack vector for WannaCry, NotPetya, and EternalBlue ransomware. Disable: <code>Remove-WindowsFeature FS-SMB1</code></div>
</div></div>\n'''

dc_agents = ''
if dc_edr:
    dc_agents += (f'<div style="display:flex;align-items:center;gap:10px;padding:7px 0;'
                  f'border-bottom:1px solid #f0edf8;">'
                  f'<span style="font-size:14px;flex-shrink:0">🛡</span>'
                  f'<div><div style="font-size:9pt;font-weight:600;color:#271e41;">EDR</div>'
                  f'<div style="font-size:8.5pt;color:#6b6080;">{h(dc_edr)}</div></div></div>\n')
if dc_rmm:
    dc_agents += (f'<div style="display:flex;align-items:center;gap:10px;padding:7px 0;">'
                  f'<span style="font-size:14px;flex-shrink:0">🔧</span>'
                  f'<div><div style="font-size:9pt;font-weight:600;color:#271e41;">RMM</div>'
                  f'<div style="font-size:8.5pt;color:#6b6080;">{h(dc_rmm)}</div></div></div>\n')

fl_lbl = str(dc_ad.get('DomainFL',''))
dc_sbr = f'''<div class="sbr-only">
<div style="background:{sbr_grad(dc_crit,dc_warn)};border-radius:10px 10px 0 0;padding:16px 24px;display:flex;justify-content:space-between;align-items:center;">
  <div>
    <div style="font-size:18px;font-weight:700;color:#fff;letter-spacing:.3px;">MEKEDC02</div>
    <div style="font-size:9pt;color:rgba(255,255,255,.85);margin-top:3px;">Domain Controller (VM) &middot; {h(domain_name)} &middot; Server 2022</div>
  </div>
  <div style="display:flex;align-items:center;gap:14px;">
    <div style="text-align:center;background:rgba(255,255,255,.{25 if dc_crit else 15});border-radius:8px;padding:8px 14px;">
      <div style="font-size:22px;font-weight:700;color:#fff;">{dc_crit}</div>
      <div style="font-size:7pt;color:rgba(255,255,255,.8);text-transform:uppercase;font-weight:600;">Critical</div>
    </div>
    <div style="text-align:center;background:rgba(255,255,255,.{25 if dc_warn else 15});border-radius:8px;padding:8px 14px;">
      <div style="font-size:22px;font-weight:700;color:#fff;">{dc_warn}</div>
      <div style="font-size:7pt;color:rgba(255,255,255,.8);text-transform:uppercase;font-weight:600;">Warning</div>
    </div>
    <span style="background:rgba(255,255,255,.22);color:#fff;font-size:10pt;font-weight:700;padding:6px 18px;border-radius:20px;border:1.5px solid rgba(255,255,255,.5);">{sbr_badge(dc_crit,dc_warn)}</span>
  </div>
</div>
<div style="background:white;border-radius:0 0 10px 10px;border:1px solid #e8e4f0;border-top:none;box-shadow:0 4px 14px rgba(0,0,0,.07);padding:20px 24px;margin-bottom:16px;">
{smb1_banner}
<div style="display:grid;grid-template-columns:1fr 1fr;gap:18px;">
<div>
{mini_box("OS &amp; System", f'''<table style="width:100%;font-size:9pt;border-collapse:collapse;">
<tr><td style="color:#6b6080;width:110px;padding:3px 0">Platform</td><td>Hyper-V VM on MEKEHV02</td></tr>
<tr><td style="color:#6b6080;padding:3px 0">OS</td><td>Server 2022 Standard {pill("Supported to 2031","green")}</td></tr>
<tr><td style="color:#6b6080;padding:3px 0">Last Reboot</td><td>{h(dc_sys.get("LastBoot","?").split(" ")[0])} &nbsp;{pill(f"{dc_up:.0f} days ago","green" if dc_up<60 else "yellow")}</td></tr>
<tr><td style="color:#6b6080;padding:3px 0">RAM</td><td>{dc_ram_t:.1f} GB total &nbsp;{pill(f"{dc_ram_f:.1f} GB free ({100-dc_ram_pct_used}%)","yellow" if dc_ram_pct_used>75 else "green")}</td></tr>
<tr><td style="color:#6b6080;padding:3px 0">Deployed</td><td>{h(dc_sys.get("OSInstallDate",""))}</td></tr>
</table>''')}
{mini_box("Active Directory", f'''<table style="width:100%;font-size:9pt;border-collapse:collapse;">
<tr><td style="color:#6b6080;width:130px;padding:3px 0">Domain</td><td>{h(domain_name)}</td></tr>
<tr><td style="color:#6b6080;padding:3px 0">Domain Controllers</td><td>{pill("1 DC — no redundancy","yellow")}</td></tr>
<tr><td style="color:#6b6080;padding:3px 0">Functional Level</td><td>{pill(fl_lbl,"yellow") if fl_lbl < "2019" else pill(fl_lbl,"purple")} &nbsp;{pill("upgrade recommended","yellow") if fl_lbl < "2019" else ""}</td></tr>
<tr><td style="color:#6b6080;padding:3px 0">Users</td><td>{dc_ad.get("UserCount","?")} total &nbsp;{pill(f"{len(stale_users)} stale (90+ days)","yellow")}</td></tr>
<tr><td style="color:#6b6080;padding:3px 0">Stale Computers</td><td>{pill(f"{len(stale_comps)}+ devices (90+ days)","yellow")}</td></tr>
</table>''')}
{mini_box("Storage", f'''<table style="width:100%;border-collapse:collapse;">
<tr><th style="text-align:left;font-size:8pt;padding:4px 0;color:#6b6080;font-weight:600">Drive</th><th style="font-size:8pt;color:#6b6080;font-weight:600">Size</th><th style="font-size:8pt;color:#6b6080;font-weight:600">Usage</th></tr>
{dc_vol_rows}</table>''', last=True)}
</div>
<div>
{mini_box("Installed Agents", dc_agents or '<span style="color:#9b8fb0;font-style:italic">None detected</span>')}
{mini_box("Server Roles", f'<div style="display:flex;flex-wrap:wrap;gap:6px;">{role_badges}</div>{all_in_one_warn}')}
{mini_box("Network Shares", f'''<table style="width:100%;font-size:9pt;border-collapse:collapse;">
<tr style="background:#ede9fe"><th style="padding:5px 8px;text-align:left;font-size:8pt">Share</th><th style="padding:5px 8px;text-align:left;font-size:8pt">Path</th><th style="padding:5px 8px;text-align:left;font-size:8pt">Notes</th></tr>
{share_rows}</table>''' if share_rows else '<span style="color:#9b8fb0;font-style:italic">None</span>')}
{mini_box("Printers" + (f" ({len(dc_printers)})" if dc_printers else ""),
  f'<table style="width:100%;font-size:8.5pt;border-collapse:collapse;">{printer_rows}</table>' if printer_rows else '<span style="color:#9b8fb0;font-style:italic">No printers detected</span>',
  last=True)}
</div>
</div>
</div>
</div>
'''

# ── SBR: VIRT TAB ─────────────────────────────────────────────────────────────
virt_sbr = f'''<div class="sbr-only">
<div style="background:linear-gradient(135deg,#5b1fa4,#3d1270);border-radius:10px 10px 0 0;padding:16px 24px;display:flex;justify-content:space-between;align-items:center;margin-bottom:0;">
  <div>
    <div style="font-size:18px;font-weight:700;color:#fff;letter-spacing:.3px;">Virtualization Summary</div>
    <div style="font-size:9pt;color:rgba(255,255,255,.85);margin-top:3px;">MEKEHV02 &middot; Hyper-V &middot; {len(inv_vms)} VM(s)</div>
  </div>
  <span style="background:rgba(255,255,255,.22);color:#fff;font-size:10pt;font-weight:700;padding:6px 18px;border-radius:20px;border:1.5px solid rgba(255,255,255,.5);">{len(inv_vms)} VM{"s" if len(inv_vms)!=1 else ""}</span>
</div>
<div style="background:white;border-radius:0 0 10px 10px;border:1px solid #e8e4f0;border-top:none;box-shadow:0 4px 14px rgba(0,0,0,.07);padding:20px 24px;margin-bottom:16px;">
<table style="width:100%;font-size:9pt;border-collapse:collapse;">
<tr style="background:#271e41"><th style="padding:8px 12px;color:#fff;text-align:left">VM Name</th><th style="padding:8px 12px;color:#fff">State</th><th style="padding:8px 12px;color:#fff">vCPU</th><th style="padding:8px 12px;color:#fff">RAM</th><th style="padding:8px 12px;color:#fff">IP</th></tr>
{''.join(f"""<tr><td style="padding:8px 12px;font-weight:600">{h(vm.get("Name","?"))}</td>
<td style="padding:8px 12px;text-align:center">{pill(vm.get("State","?"),"green" if vm.get("State")=="Running" else "yellow")}</td>
<td style="padding:8px 12px;text-align:center">{vm.get("vCPU",vm.get("CPUCount","?"))}</td>
<td style="padding:8px 12px;text-align:center">{vm.get("RAMgb",vm.get("MemoryGB",0)):.1f} GB</td>
<td style="padding:8px 12px;font-family:monospace;font-size:8.5pt">{h(str(vm.get("IPs","?")))}</td></tr>\n""" for vm in inv_vms)}
</table>
</div>
</div>
'''

# ── QUICKJUMP NAVBARS ─────────────────────────────────────────────────────────
hv_nav = f'''<div id="top-hv02" class="hide-sbr" style="background:white;border-radius:8px;border:1px solid #e8e4f0;padding:12px 16px;margin-bottom:16px;box-shadow:0 2px 8px rgba(0,0,0,0.04);">
<div style="display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin-bottom:8px;">
<span style="font-size:8pt;color:#6b6080;text-transform:uppercase;letter-spacing:.5px;font-weight:700;margin-right:4px;">Jump to:</span>
{nav_link("hv02-alerts","Alerts")}{nav_link("hv02-overview","Overview")}{nav_link("hv02-hardware","Hardware")}{nav_link("hv02-apps","Applications")}{nav_link("hv02-roles","Roles")}{nav_link("hv02-roleconfig","Role Config")}{nav_link("hv02-disks","Disks")}{nav_link("hv02-network","Network")}{nav_link("hv02-lports","Listening Ports")}{nav_link("hv02-services","Services")}</div>
</div>
'''

dc_nav = f'''<div id="top-dc02" class="hide-sbr" style="background:white;border-radius:8px;border:1px solid #e8e4f0;padding:12px 16px;margin-bottom:16px;box-shadow:0 2px 8px rgba(0,0,0,0.04);">
<div style="display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin-bottom:8px;">
<span style="font-size:8pt;color:#6b6080;text-transform:uppercase;letter-spacing:.5px;font-weight:700;margin-right:4px;">Jump to:</span>
{nav_link("dc02-alerts","Alerts")}{nav_link("dc02-overview","Overview")}{nav_link("dc02-hardware","Hardware")}{nav_link("dc02-apps","Applications")}{nav_link("dc02-roles","Roles")}{nav_link("dc02-roleconfig","Role Config")}{nav_link("dc02-disks","Disks")}{nav_link("dc02-network","Network")}{nav_link("dc02-lports","Listening Ports")}{nav_link("dc02-services","Services")}{nav_link("dc02-shares","File Shares")}</div>
<div style="display:flex;flex-wrap:wrap;gap:6px;align-items:center;">
<span style="font-size:8pt;color:#6b6080;text-transform:uppercase;letter-spacing:.5px;font-weight:700;margin-right:4px;">Role Sections:</span>
{nav_link_dark("dc02-roleconf-ad","AD Domain Services")}{nav_link_dark("dc02-roleconf-dhcp","DHCP Server")}{nav_link_dark("dc02-roleconf-dns","DNS Server")}{nav_link_dark("dc02-roleconf-files","File and Storage Services")}</div>
</div>
'''

virt_nav = f'''<div id="top-virt" class="hide-sbr" style="background:white;border-radius:8px;border:1px solid #e8e4f0;padding:12px 16px;margin-bottom:16px;box-shadow:0 2px 8px rgba(0,0,0,0.04);">
<div style="display:flex;flex-wrap:wrap;gap:6px;align-items:center;">
<span style="font-size:8pt;color:#6b6080;text-transform:uppercase;letter-spacing:.5px;font-weight:700;margin-right:4px;">Jump to:</span>
{nav_link("virt-summary","Host Summary")}{nav_link("virt-vms","Virtual Machines")}{nav_link("virt-vswitches","Virtual Switches")}{nav_link("virt-hostcfg","Host Config")}</div>
</div>
'''

# ── CARD BODIES ───────────────────────────────────────────────────────────────

## HV02 ALERTS
hv_alerts_html = ''.join(flag_div(*f) for f in hv_flags) or '<span class="pill pill-green">No critical alerts</span>\n'
hv_alerts_html += top_link('hv02')

## HV02 OVERVIEW
hv_net = hv.get('Network', {})
hv_adapt = as_list(hv_net.get('Adapters', {})) or ([hv_net['Adapters']] if isinstance(hv_net.get('Adapters'),dict) else [])
hv_ip = hv_adapt[0].get('IPAddresses','').split(',')[0].strip() if hv_adapt else '?'
hv_gw = hv_adapt[0].get('Gateway','') if hv_adapt else ''
hv_dns_ip = hv_adapt[0].get('DNS','') if hv_adapt else ''

hv_overview_html = f'''<div class="stat-grid">
<div class="stat-box"><div class="stat-num">{hv_hw.get("CPUCores","?")}</div><div class="stat-lbl">CPU Cores</div></div>
<div class="stat-box"><div class="stat-num">{hv_hw.get("RAMTotalGB",0):.0f}</div><div class="stat-lbl">GB RAM</div></div>
<div class="stat-box"><div class="stat-num">{len(inv_vms)}</div><div class="stat-lbl">VMs</div></div>
<div class="stat-box"><div class="stat-num">{hv_sys.get("UptimeDays",0):.0f}</div><div class="stat-lbl">Days Up</div></div>
</div>
<table>
<tr><th>Attribute</th><th>Value</th></tr>
<tr><td>Hostname</td><td>{h(hv_sys.get("Hostname",""))}</td></tr>
<tr><td>OS</td><td>{h(hv_sys.get("OSName",""))} (Build {h(str(hv_sys.get("OSBuild","")))})</td></tr>
<tr><td>Domain</td><td>{pill("WORKGROUP — not domain-joined","yellow") if hv_sys.get("Domain","").upper()=="WORKGROUP" else h(hv_sys.get("Domain",""))}</td></tr>
<tr><td>Last Boot</td><td>{h(hv_sys.get("LastBoot",""))}</td></tr>
<tr><td>Timezone</td><td>{h(hv_sys.get("Timezone",""))}</td></tr>
<tr><td>Run As</td><td><code>{h(hv_sys.get("RunAsUser",""))}</code></td></tr>
<tr><td>CPU</td><td>{h(hv_hw.get("CPUName",""))} ({hv_hw.get("CPUCores","?")} cores / {hv_hw.get("CPULogical", hv_hw.get("CPUCores","?"))} logical)</td></tr>
<tr><td>RAM</td><td>{hv_hw.get("RAMTotalGB",0):.2f} GB total / {hv_hw.get("RAMAvailGB",0):.2f} GB available</td></tr>
<tr><td>IP Address</td><td>{h(hv_ip)}</td></tr>
<tr><td>Gateway</td><td>{h(hv_gw)}</td></tr>
<tr><td>DNS</td><td>{h(hv_dns_ip)}</td></tr>
<tr><td>Model</td><td>{h(hv_hw.get("Manufacturer",""))} {h(hv_hw.get("Model",""))}</td></tr>
<tr><td>Serial Number</td><td><code>{h(hv_hw.get("SerialNumber",""))}</code></td></tr>
<tr><td>BIOS</td><td>{h(hv_hw.get("BIOSVersion",""))} ({h(hv_hw.get("BIOSDate",""))})</td></tr>
<tr><td>OS Install Date</td><td>{h(hv_sys.get("OSInstallDate",""))}</td></tr>
<tr><td>EOL Date</td><td>{h(hv_sys.get("OSEOLDate",""))} {pill(hv_sys.get("OSEOLStatus",""),"green")}</td></tr>
</table>
''' + top_link('hv02')

## HV02 HARDWARE
hv_hw_html = f'''<table>
<tr><th>Component</th><th>Value</th></tr>
<tr><td>Manufacturer</td><td>{h(hv_hw.get("Manufacturer",""))}</td></tr>
<tr><td>Model</td><td>{h(hv_hw.get("Model",""))}</td></tr>
<tr><td>Serial Number</td><td><code>{h(hv_hw.get("SerialNumber",""))}</code></td></tr>
<tr><td>Board Product</td><td>{h(hv_hw.get("BoardProduct",""))}</td></tr>
<tr><td>Board Serial</td><td><code>{h(hv_hw.get("BoardSerial",""))}</code></td></tr>
<tr><td>CPU</td><td>{h(hv_hw.get("CPUName",""))} &middot; {hv_hw.get("CPUCores","?")} cores / {hv_hw.get("CPULogical",hv_hw.get("CPUCores","?"))} threads</td></tr>
<tr><td>RAM</td><td>{hv_hw.get("RAMTotalGB",0):.2f} GB total / {hv_hw.get("RAMAvailGB",0):.2f} GB available</td></tr>
<tr><td>BIOS Version</td><td>{h(hv_hw.get("BIOSVersion",""))}</td></tr>
<tr><td>BIOS Date</td><td>{h(hv_hw.get("BIOSDate",""))}</td></tr>
<tr><td>Platform</td><td>{pill(hv_hw.get("VMPlatform","Physical"),"purple")}</td></tr>
</table>
''' + top_link('hv02')

## HV02 APPS
def apps_table(apps_list, tid):
    rows = ''
    for a in sorted(apps_list, key=lambda x: x.get('Name','')):
        sev = a.get('FlagSeverity','none')
        sc  = 'green' if sev=='none' else ('red' if sev=='critical' else 'yellow')
        idate = str(a.get('InstallDate','') or '').strip()
        rows += f'<tr><td>{h(a.get("Name","?"))}</td><td>{h(a.get("Version","") or "")}</td><td>{h(a.get("Publisher","") or "")}</td><td>{h(idate)}</td><td>{pill(sev,sc)}</td></tr>\n'
    return (f'<table><tr><th>Name</th><th>Version</th><th>Publisher</th><th>Install Date</th><th>Flag</th></tr>\n'
            + rows + f'</table>\n') + top_link(tid)

hv_apps_html = apps_table(hv_apps, 'hv02')

## HV02 ROLES & FEATURES
hv_roles_data = hv.get('Roles', {})
hv_roles_list = as_list(hv_roles_data.get('InstalledRoles', []))
hv_features_list = as_list(hv_roles_data.get('InstalledFeatures', []))

hv_roles_html = sub('Installed Roles')
hv_roles_html += '<div class="role-grid">' + ''.join(f'<span class="role-badge">{h(r.get("DisplayName",r.get("Name","")))}</span>' for r in hv_roles_list) + '</div>\n'
hv_roles_html += sub('Installed Features')
hv_roles_html += '<div class="role-grid">' + ''.join(f'<span class="role-badge" style="background:#f0f4ff;color:#5b1fa4;border-color:#a5b4fc;">{h(f.get("DisplayName",f.get("Name","")))}</span>' for f in hv_features_list) + '</div>\n'
hv_roles_html += top_link('hv02')

## HV02 ROLE CONFIG (Hyper-V)
hv_hv = hv.get('HyperV', {})
hv_vms_hv = as_list(hv_hv.get('VMs', {})) or ([hv_hv['VMs']] if isinstance(hv_hv.get('VMs'),dict) else [])
hv_vsw_hv = as_list(hv_hv.get('VirtualSwitches', {})) or ([hv_hv['VirtualSwitches']] if isinstance(hv_hv.get('VirtualSwitches'),dict) else [])

hv_rc_html = sub('Hyper-V Host Configuration')
hv_rc_html += f'''<table>
<tr><th>Setting</th><th>Value</th></tr>
<tr><td>Live Migration</td><td>{pill("Enabled","green") if hv_hv.get("LiveMigrationEnabled") else pill("Disabled","gray")}</td></tr>
<tr><td>NUMA Spanning</td><td>{pill("Enabled","green") if hv_hv.get("NumaSpanningEnabled") else pill("Disabled","gray")}</td></tr>
<tr><td>Default VM Path</td><td><code>{h(hv_hv.get("DefaultVMPath",""))}</code></td></tr>
<tr><td>Default VHD Path</td><td><code>{h(hv_hv.get("DefaultVHDPath",""))}</code></td></tr>
</table>
'''
if hv_vms_hv:
    hv_rc_html += sub('Virtual Machines (Hyper-V Report)')
    vm_rows2 = ''
    for vm in hv_vms_hv:
        state = vm.get('State','?')
        sc = 'green' if state=='Running' else 'yellow'
        cpu = vm.get('CPUCount','?')
        mem = vm.get('MemoryGB', 0)
        up  = vm.get('Uptime','')
        snaps = vm.get('Checkpoints', vm.get('Snapshots', 0))
        vm_rows2 += (f'<tr><td>{h(vm.get("Name","?"))}</td><td>{pill(state,sc)}</td>'
                     f'<td>{cpu}</td><td>{mem:.1f} GB</td><td>{h(str(up)[:12])}</td>'
                     f'<td>{pill("⚠️ "+str(snaps)+" snapshot(s)","yellow") if snaps else pill("None","green")}</td></tr>\n')
    hv_rc_html += (f'<table><tr><th>Name</th><th>State</th><th>vCPU</th><th>RAM</th>'
                   f'<th>Uptime</th><th>Snapshots</th></tr>\n' + vm_rows2 + '</table>\n')
hv_rc_html += top_link('hv02')

## HV02 DISKS
hv_disk_html = sub('Physical Volumes')
hv_disk_html += '<table><tr><th>Drive</th><th>Label</th><th>FS</th><th>Total GB</th><th>Free GB</th><th>Used %</th><th>Usage Bar</th></tr>\n'
for d in hv_disks:
    pct = d.get('UsedPct', 0)
    rc  = 'style="background:#fff0f0"' if pct>=85 else ('style="background:#fff8e1"' if pct>=70 else '')
    hv_disk_html += (f'<tr {rc}><td><b>{h(d.get("Drive","?"))}</b></td><td>{h(d.get("Label","") or "")}</td>'
                     f'<td>{h(d.get("Filesystem","NTFS"))}</td><td>{d.get("TotalGB",0):.2f}</td>'
                     f'<td>{d.get("FreeGB",0):.2f}</td><td>{pct}%</td>'
                     f'<td style="min-width:100px">{disk_bar(pct)}</td></tr>\n')
hv_disk_html += '</table>\n'

# VM Disk summary
hv_disk_html += sub('VM Disk Files (from Hyper-V inventory)')
vm_disk_rows = ''
for vm in inv_vms:
    for vd in as_list(vm.get('Disks', [])):
        used = vd.get('UsedGB', 0)
        size = vd.get('SizeGB', 0)
        pct  = int(used/size*100) if size else 0
        pname = vd.get('Path', '?').split('\\')[-1]
        pc = 'red' if pct>=95 else ('yellow' if pct>=80 else 'green')
        vm_disk_rows += (f'<tr><td>{h(vm.get("Name","?"))}</td><td style="font-family:monospace;font-size:8.5pt">{h(pname)}</td>'
                         f'<td>{h(vd.get("VHDType","?"))}</td><td>{size:.1f} GB</td><td>{used:.1f} GB</td>'
                         f'<td>{pill(f"{pct}%",pc)}{disk_bar(pct)}</td></tr>\n')
if vm_disk_rows:
    hv_disk_html += (f'<table><tr><th>VM</th><th>Filename</th><th>Type</th><th>Allocated</th>'
                     f'<th>Used</th><th>Usage</th></tr>\n' + vm_disk_rows + '</table>\n')
hv_disk_html += top_link('hv02')

## HV02 NETWORK
hv_net_html = sub('Network Adapters')
adapt_list = as_list(hv_net.get('Adapters',{})) or ([hv_net['Adapters']] if isinstance(hv_net.get('Adapters'),dict) else [])
for ad in adapt_list:
    hv_net_html += f'''<table style="margin-bottom:16px;">
<tr><th>Attribute</th><th>Value</th></tr>
<tr><td>Description</td><td>{h(ad.get("Description",""))}</td></tr>
<tr><td>IP Addresses</td><td>{h(ad.get("IPAddresses",""))}</td></tr>
<tr><td>Gateway</td><td>{h(ad.get("Gateway",""))}</td></tr>
<tr><td>DNS</td><td>{h(ad.get("DNS",""))}</td></tr>
<tr><td>MAC</td><td><code>{h(ad.get("MAC",""))}</code></td></tr>
<tr><td>DHCP</td><td>{pill("Enabled","yellow") if ad.get("DHCPEnabled") else pill("Static","green")}</td></tr>
<tr><td>Subnet Mask</td><td>{h(ad.get("SubnetMasks",""))}</td></tr>
</table>
'''
hv_net_html += top_link('hv02')

## HV02 LISTENING PORTS
hv_ports = as_list(hv_net.get('ListeningPorts', []))
hv_lports_html = f'<table><tr><th>Port</th><th>Protocol</th><th>Process</th><th>PID</th></tr>\n'
for p in sorted(hv_ports, key=lambda x: int(str(x.get('LocalPort',9999)))):
    hv_lports_html += (f'<tr><td>{p.get("LocalPort","?")}</td><td>{h(p.get("Protocol",""))}</td>'
                       f'<td>{h(p.get("ProcessName",""))}</td><td>{p.get("PID","")}</td></tr>\n')
hv_lports_html += '</table>\n' + top_link('hv02')

## HV02 SERVICES
hv_services = as_list(hv.get('Services', []))
hv_svc_html = f'<table><tr><th>Name</th><th>Display Name</th><th>Status</th><th>Start Type</th></tr>\n'
for svc in sorted(hv_services, key=lambda x: x.get('Name','')):
    status = svc.get('Status','')
    sc = 'green' if status=='Running' else ('yellow' if status=='Stopped' else 'gray')
    hv_svc_html += (f'<tr><td><code>{h(svc.get("Name",""))}</code></td><td>{h(svc.get("DisplayName",""))}</td>'
                    f'<td>{pill(status,sc)}</td><td>{h(svc.get("StartType",""))}</td></tr>\n')
hv_svc_html += '</table>\n' + top_link('hv02')

## HV02 SERVICE ANOMALIES
hv_anomaly_svcs = [s for s in hv_services
                   if s.get('StartType') in ('Automatic','Auto') and s.get('Status')=='Stopped'
                   and s.get('Name','') not in ('RemoteRegistry','AppMgmt')]
hv_anom_html = ''
if hv_anomaly_svcs:
    hv_anom_html = '<div class="flag-warning"><div class="flag-label">Auto-Start Services That Are Stopped</div><div class="flag-detail">These services are configured to start automatically but are currently stopped:</div></div>\n'
    hv_anom_html += '<table><tr><th>Name</th><th>Display Name</th><th>Start Type</th></tr>\n'
    for s in hv_anomaly_svcs:
        hv_anom_html += f'<tr><td><code>{h(s.get("Name",""))}</code></td><td>{h(s.get("DisplayName",""))}</td><td>{h(s.get("StartType",""))}</td></tr>\n'
    hv_anom_html += '</table>\n'
else:
    hv_anom_html = '<span class="pill pill-green">No service anomalies detected</span>\n'
hv_anom_html += top_link('hv02')

# ── DC02 CARD BODIES ──────────────────────────────────────────────────────────

## DC02 ALERTS
dc_alerts_html = ''.join(flag_div(*f) for f in dc_flags) or '<span class="pill pill-green">No critical alerts</span>\n'
dc_alerts_html += top_link('dc02')

## DC02 OVERVIEW
dc_net = dc.get('Network', {})
dc_adapt = as_list(dc_net.get('Adapters', {})) or ([dc_net['Adapters']] if isinstance(dc_net.get('Adapters'),dict) else [])
dc_ip = dc_adapt[0].get('IPAddresses','').split(',')[0].strip() if dc_adapt else '?'
dc_gw = dc_adapt[0].get('Gateway','') if dc_adapt else ''
dc_dns_ip = dc_adapt[0].get('DNS','') if dc_adapt else ''
dc_ram_t = dc_hw.get('RAMTotalGB', 0)
dc_ram_f = dc_hw.get('RAMAvailGB', 0)

dc_overview_html = f'''<div class="stat-grid">
<div class="stat-box"><div class="stat-num">{dc_hw.get("CPUCores","?")}</div><div class="stat-lbl">vCPUs</div></div>
<div class="stat-box"><div class="stat-num">{dc_ram_t:.0f}</div><div class="stat-lbl">GB RAM</div></div>
<div class="stat-box"><div class="stat-num">{dc_ad.get("UserCount","?")}</div><div class="stat-lbl">AD Users</div></div>
<div class="stat-box"><div class="stat-num">{dc_sys.get("UptimeDays",0):.0f}</div><div class="stat-lbl">Days Up</div></div>
</div>
<table>
<tr><th>Attribute</th><th>Value</th></tr>
<tr><td>Hostname</td><td>{h(dc_sys.get("Hostname",""))}</td></tr>
<tr><td>OS</td><td>{h(dc_sys.get("OSName",""))} (Build {h(str(dc_sys.get("OSBuild","")))})</td></tr>
<tr><td>Domain</td><td>{h(dc_sys.get("Domain",""))}</td></tr>
<tr><td>Last Boot</td><td>{h(dc_sys.get("LastBoot",""))}</td></tr>
<tr><td>Timezone</td><td>{h(dc_sys.get("Timezone",""))}</td></tr>
<tr><td>Run As</td><td><code>{h(dc_sys.get("RunAsUser",""))}</code></td></tr>
<tr><td>Platform</td><td>{h(dc_hw.get("VMPlatform","Hyper-V"))} on MEKEHV02</td></tr>
<tr><td>CPU</td><td>{h(dc_hw.get("CPUName",""))} &middot; {dc_hw.get("CPUCores","?")} vCPUs</td></tr>
<tr><td>RAM</td><td>{dc_ram_t:.2f} GB total / {dc_ram_f:.2f} GB available</td></tr>
<tr><td>IP Address</td><td>{h(dc_ip)}</td></tr>
<tr><td>DNS</td><td>{h(dc_dns_ip)}</td></tr>
<tr><td>BIOS</td><td>{h(dc_hw.get("BIOSVersion",""))}</td></tr>
<tr><td>OS Install Date</td><td>{h(dc_sys.get("OSInstallDate",""))}</td></tr>
<tr><td>EOL Date</td><td>{h(dc_sys.get("OSEOLDate",""))} {pill(dc_sys.get("OSEOLStatus",""),"green")}</td></tr>
</table>
''' + top_link('dc02')

## DC02 HARDWARE
dc_hw_html = f'''<table>
<tr><th>Component</th><th>Value</th></tr>
<tr><td>Platform</td><td>{pill(dc_hw.get("VMPlatform","Hyper-V"),"purple")} hosted on MEKEHV02</td></tr>
<tr><td>Model</td><td>{h(dc_hw.get("Model",""))}</td></tr>
<tr><td>CPU</td><td>{h(dc_hw.get("CPUName",""))} &middot; {dc_hw.get("CPUCores","?")} vCPUs</td></tr>
<tr><td>RAM</td><td>{dc_ram_t:.2f} GB total / {dc_ram_f:.2f} GB available</td></tr>
<tr><td>BIOS</td><td>{h(dc_hw.get("BIOSVersion",""))} ({h(dc_hw.get("BIOSDate",""))})</td></tr>
<tr><td>Serial Number</td><td><code>{h(dc_hw.get("SerialNumber",""))}</code></td></tr>
</table>
''' + top_link('dc02')

## DC02 APPS
dc_apps_html = apps_table(dc_apps, 'dc02')

## DC02 ROLES
dc_roles_html = sub('Installed Roles')
dc_roles_html += '<div class="role-grid">' + ''.join(f'<span class="role-badge">{h(r.get("DisplayName",r.get("Name","")))}</span>' for r in dc_roles_list) + '</div>\n'
dc_roles_html += sub('Installed Features')
dc_roles_html += '<div class="role-grid">' + ''.join(f'<span class="role-badge" style="background:#f0f4ff;color:#5b1fa4;border-color:#a5b4fc;">{h(f.get("DisplayName",f.get("Name","")))}</span>' for f in dc_features_list) + '</div>\n'
dc_roles_html += top_link('dc02')

## DC02 ROLE CONFIG
dhcp = dc.get('DHCP', {})
dns  = dc.get('DNS', {})
nps  = dc.get('NPS', {})

dc_rc_html = f'<div id="dc02-roleconf-ad"></div>' + sub('Active Directory')
dc_rc_html += f'''<table>
<tr><th>Attribute</th><th>Value</th></tr>
<tr><td>Forest</td><td>{h(dc_ad.get("ForestName",""))}</td></tr>
<tr><td>PDC Emulator</td><td>{h(dc_ad.get("PDCEmulator",""))}</td></tr>
<tr><td>RID Master</td><td>{h(dc_ad.get("RIDMaster",""))}</td></tr>
<tr><td>Domain FL</td><td>{pill(str(dc_ad.get("DomainFL","")),"yellow" if str(dc_ad.get("DomainFL","")) < "2019" else "green")}</td></tr>
<tr><td>OU Count</td><td>{dc_ad.get("OUCount","?")}</td></tr>
<tr><td>User Count</td><td>{dc_ad.get("UserCount","?")}</td></tr>
<tr><td>Stale Users</td><td>{len(stale_users)}</td></tr>
<tr><td>Stale Computers</td><td>{len(stale_comps)}</td></tr>
</table>
'''

if stale_users:
    dc_rc_html += sub(f'Stale User Accounts ({len(stale_users)} inactive 90+ days)')
    dc_rc_html += '<table><tr><th>Name</th><th>SAM Account</th><th>Last Logon</th></tr>\n'
    for u in stale_users[:30]:
        dc_rc_html += f'<tr><td>{h(u.get("Name",""))}</td><td><code>{h(u.get("SamAccountName",""))}</code></td><td>{h(str(u.get("LastLogon","") or "Never"))}</td></tr>\n'
    if len(stale_users) > 30:
        dc_rc_html += f'<tr><td colspan="3" style="color:#6b6080;font-style:italic">... and {len(stale_users)-30} more</td></tr>\n'
    dc_rc_html += '</table>\n'

dc_rc_html += f'<div id="dc02-roleconf-dhcp"></div>' + sub('DHCP Server')
dhcp_scopes = as_list(dhcp.get('Scopes', []))
if dhcp_scopes:
    dc_rc_html += '<table><tr><th>Scope</th><th>Name</th><th>State</th><th>Start IP</th><th>End IP</th><th>Leases</th></tr>\n'
    for sc in dhcp_scopes:
        state = sc.get('State','?')
        scc = 'green' if state=='Active' else 'yellow'
        dc_rc_html += (f'<tr><td><code>{h(sc.get("ScopeID",""))}</code></td><td>{h(sc.get("Name",""))}</td>'
                       f'<td>{pill(state,scc)}</td><td>{h(sc.get("StartRange",""))}</td>'
                       f'<td>{h(sc.get("EndRange",""))}</td><td>{sc.get("ActiveLeases",sc.get("Leases","?"))}</td></tr>\n')
    dc_rc_html += '</table>\n'
else:
    dc_rc_html += '<span class="meta-line">No DHCP scopes collected</span>\n'

dc_rc_html += f'<div id="dc02-roleconf-dns"></div>' + sub('DNS Server')
dns_zones = as_list(dns.get('Zones', []))
dns_fwd = dns.get('Forwarders', '')
dc_rc_html += f'<div class="meta-line">Forwarders: <code>{h(str(dns_fwd))}</code></div>\n'
if dns_zones:
    dc_rc_html += '<table><tr><th>Zone</th><th>Type</th><th>Records</th><th>Dynamic Update</th></tr>\n'
    for z in dns_zones:
        dc_rc_html += (f'<tr><td>{h(z.get("Name",""))}</td><td>{h(z.get("ZoneType",""))}</td>'
                       f'<td>{z.get("RecordCount","?")}</td><td>{h(z.get("DynamicUpdate",""))}</td></tr>\n')
    dc_rc_html += '</table>\n'

if nps.get('Installed'):
    dc_rc_html += sub('NPS / RADIUS')
    nps_policies = as_list(nps.get('Policies', []))
    nps_clients  = as_list(nps.get('Clients', []))
    if nps_clients:
        dc_rc_html += '<table><tr><th>Client Name</th><th>Address</th><th>Enabled</th></tr>\n'
        for nc in nps_clients:
            dc_rc_html += (f'<tr><td>{h(nc.get("Name",""))}</td><td><code>{h(nc.get("Address",""))}</code></td>'
                           f'<td>{pill("Yes","green") if nc.get("Enabled") else pill("No","yellow")}</td></tr>\n')
        dc_rc_html += '</table>\n'
    else:
        dc_rc_html += '<span class="meta-line">NPS installed but no client data collected</span>\n'

dc_rc_html += f'<div id="dc02-roleconf-files"></div>' + sub('File and Storage Services')
if dc_shares_real:
    dc_rc_html += f'<table><tr><th>Share</th><th>Path</th><th>Description</th></tr>\n'
    for s in dc_shares_real:
        dc_rc_html += (f'<tr><td><b>{h(s.get("Name",""))}</b></td>'
                       f'<td style="font-family:monospace;font-size:8.5pt">{h(s.get("Path",""))}</td>'
                       f'<td style="color:#6b6080">{h(s.get("Description","") or "")}</td></tr>\n')
    dc_rc_html += '</table>\n'
dc_rc_html += top_link('dc02')

## DC02 DISKS
dc_disk_html = sub('Physical Volumes')
dc_disk_html += '<table><tr><th>Drive</th><th>Label</th><th>FS</th><th>Total GB</th><th>Free GB</th><th>Used %</th><th>Usage Bar</th></tr>\n'
for d in dc_disks_list:
    pct = d.get('UsedPct', 0)
    rc  = 'style="background:#fff0f0"' if pct>=85 else ('style="background:#fff8e1"' if pct>=70 else '')
    dc_disk_html += (f'<tr {rc}><td><b>{h(d.get("Drive","?"))}</b></td><td>{h(d.get("Label","") or "")}</td>'
                     f'<td>{h(d.get("Filesystem","NTFS"))}</td><td>{d.get("TotalGB",0):.2f}</td>'
                     f'<td>{d.get("FreeGB",0):.2f}</td><td>{pct}%</td>'
                     f'<td style="min-width:100px">{disk_bar(pct)}</td></tr>\n')
dc_disk_html += '</table>\n' + top_link('dc02')

## DC02 NETWORK
dc_net_html = sub('Network Adapters')
for ad in dc_adapt:
    dc_net_html += f'''<table style="margin-bottom:16px;">
<tr><th>Attribute</th><th>Value</th></tr>
<tr><td>Description</td><td>{h(ad.get("Description",""))}</td></tr>
<tr><td>IP Addresses</td><td>{h(ad.get("IPAddresses",""))}</td></tr>
<tr><td>Gateway</td><td>{h(ad.get("Gateway",""))}</td></tr>
<tr><td>DNS</td><td>{h(ad.get("DNS",""))}</td></tr>
<tr><td>MAC</td><td><code>{h(ad.get("MAC",""))}</code></td></tr>
<tr><td>DHCP</td><td>{pill("Enabled","yellow") if ad.get("DHCPEnabled") else pill("Static","green")}</td></tr>
</table>
'''
dc_net_html += top_link('dc02')

## DC02 LISTENING PORTS
dc_ports = as_list(dc_net.get('ListeningPorts', []))
dc_lports_html = f'<table><tr><th>Port</th><th>Protocol</th><th>Process</th><th>PID</th></tr>\n'
for p in sorted(dc_ports, key=lambda x: int(str(x.get('LocalPort',9999)))):
    dc_lports_html += (f'<tr><td>{p.get("LocalPort","?")}</td><td>{h(p.get("Protocol",""))}</td>'
                       f'<td>{h(p.get("ProcessName",""))}</td><td>{p.get("PID","")}</td></tr>\n')
dc_lports_html += '</table>\n' + top_link('dc02')

## DC02 SERVICES
dc_services = as_list(dc.get('Services', []))
dc_svc_html = f'<table><tr><th>Name</th><th>Display Name</th><th>Status</th><th>Start Type</th></tr>\n'
for svc in sorted(dc_services, key=lambda x: x.get('Name','')):
    status = svc.get('Status','')
    sc = 'green' if status=='Running' else ('yellow' if status=='Stopped' else 'gray')
    dc_svc_html += (f'<tr><td><code>{h(svc.get("Name",""))}</code></td><td>{h(svc.get("DisplayName",""))}</td>'
                    f'<td>{pill(status,sc)}</td><td>{h(svc.get("StartType",""))}</td></tr>\n')
dc_svc_html += '</table>\n' + top_link('dc02')

## DC02 SERVICE ANOMALIES
dc_anomaly_svcs = [s for s in dc_services
                   if s.get('StartType') in ('Automatic','Auto') and s.get('Status')=='Stopped'
                   and s.get('Name','') not in ('RemoteRegistry','AppMgmt')]
dc_anom_html = ''
if dc_anomaly_svcs:
    dc_anom_html = '<div class="flag-warning"><div class="flag-label">Auto-Start Services That Are Stopped</div></div>\n'
    dc_anom_html += '<table><tr><th>Name</th><th>Display Name</th><th>Start Type</th></tr>\n'
    for s in dc_anomaly_svcs:
        dc_anom_html += f'<tr><td><code>{h(s.get("Name",""))}</code></td><td>{h(s.get("DisplayName",""))}</td><td>{h(s.get("StartType",""))}</td></tr>\n'
    dc_anom_html += '</table>\n'
else:
    dc_anom_html = '<span class="pill pill-green">No service anomalies detected</span>\n'
dc_anom_html += top_link('dc02')

## DC02 FILE SHARES
dc_shares_html = f'<table><tr><th>Share</th><th>Path</th><th>Open Sessions</th></tr>\n'
for s in dc_shares_real:
    dc_shares_html += (f'<tr><td><b>{h(s.get("Name",""))}</b></td>'
                       f'<td style="font-family:monospace;font-size:8.5pt">{h(s.get("Path",""))}</td>'
                       f'<td>{s.get("OpenSessions",0)}</td></tr>\n')
dc_shares_html += '</table>\n'
if dc_printers:
    dc_shares_html += sub(f'Printers ({len(dc_printers)})')
    dc_shares_html += '<table><tr><th>Printer Name</th><th>Driver</th><th>Port</th><th>Shared</th></tr>\n'
    for p in dc_printers:
        dc_shares_html += (f'<tr><td>{h(p.get("Name",p.get("PrinterName","?")))}</td>'
                           f'<td>{h(p.get("DriverName","") or "")}</td><td>{h(p.get("PortName","") or "")}</td>'
                           f'<td>{pill("Yes","green") if p.get("Shared") else pill("No","gray")}</td></tr>\n')
    dc_shares_html += '</table>\n'
dc_shares_html += top_link('dc02')

# ── VIRT CARD BODIES ──────────────────────────────────────────────────────────

## VIRT HOST SUMMARY
virt_host_html = f'''<table>
<tr><th>Attribute</th><th>Value</th></tr>
<tr><td>Hypervisor Host</td><td>MEKEHV02</td></tr>
<tr><td>Model</td><td>{h(inv_host.get("Manufacturer","Dell Inc."))} {h(inv_host.get("Model",hv_hw.get("Model","")))}</td></tr>
<tr><td>CPU</td><td>{h(inv_host.get("CPUModel",hv_hw.get("CPUName","")))} &middot; {inv_host.get("CPUCores",hv_hw.get("CPUCores","?"))} cores / {inv_host.get("CPULogical",hv_hw.get("CPUCores","?"))} threads</td></tr>
<tr><td>Total RAM</td><td>{inv_host.get("TotalRAMgb",hv_hw.get("RAMTotalGB",0)):.1f} GB</td></tr>
<tr><td>VMs Running</td><td>{len([v for v in inv_vms if v.get("State")=="Running"])}</td></tr>
<tr><td>Total VMs</td><td>{len(inv_vms)}</td></tr>
</table>
''' + top_link('virt')

## VIRT VMs
virt_vms_html = '<table><tr><th>Name</th><th>State</th><th>Gen</th><th>vCPU</th><th>RAM (GB)</th><th>Dynamic</th><th>Uptime (hrs)</th><th>Snapshots</th><th>IPs</th></tr>\n'
for vm in inv_vms:
    state = vm.get('State','?')
    sc = 'green' if state=='Running' else 'yellow'
    snaps = vm.get('Snapshots', 0)
    snap_c = 'yellow' if snaps else 'green'
    virt_vms_html += (f'<tr><td><b>{h(vm.get("Name","?"))}</b></td><td>{pill(state,sc)}</td>'
                      f'<td>{vm.get("Generation","?")}</td><td>{vm.get("vCPU",vm.get("CPUCount","?"))}</td>'
                      f'<td>{vm.get("RAMgb",vm.get("MemoryGB",0)):.2f}</td>'
                      f'<td>{pill("Yes","green") if vm.get("DynamicMemory") else pill("No","gray")}</td>'
                      f'<td>{vm.get("UptimeHours",0):.1f}</td>'
                      f'<td>{pill(str(snaps),snap_c)}</td>'
                      f'<td style="font-family:monospace;font-size:8.5pt">{h(str(vm.get("IPs","?")))}</td></tr>\n')
virt_vms_html += '</table>\n'

# VM disks detail
for vm in inv_vms:
    vdisks = as_list(vm.get('Disks', []))
    if vdisks:
        virt_vms_html += sub(f'Disks: {h(vm.get("Name","?"))}')
        virt_vms_html += '<table><tr><th>Controller</th><th>Path</th><th>Type</th><th>Allocated GB</th><th>Used GB</th><th>Usage</th></tr>\n'
        for vd in vdisks:
            size = vd.get('SizeGB', 0)
            used = vd.get('UsedGB', 0)
            pct  = int(used/size*100) if size else 0
            pc   = 'red' if pct>=95 else ('yellow' if pct>=80 else 'green')
            virt_vms_html += (f'<tr><td>{h(vd.get("ControllerType",""))}</td>'
                              f'<td style="font-family:monospace;font-size:8.5pt">{h(vd.get("Path",""))}</td>'
                              f'<td>{pill(vd.get("VHDType","?"),"purple")}</td>'
                              f'<td>{size:.1f}</td><td>{used:.1f}</td>'
                              f'<td>{pill(f"{pct}%",pc)}{disk_bar(pct)}</td></tr>\n')
        virt_vms_html += '</table>\n'
virt_vms_html += top_link('virt')

## VIRT VSWITCHES
vsw_list = inv_vsw or as_list(hv_hv.get('VirtualSwitches',{})) or ([hv_hv['VirtualSwitches']] if isinstance(hv_hv.get('VirtualSwitches'),dict) else [])
virt_vsw_html = '<table><tr><th>Name</th><th>Type</th><th>Physical Adapter</th></tr>\n'
for vsw in vsw_list:
    virt_vsw_html += (f'<tr><td>{h(vsw.get("Name",""))}</td>'
                      f'<td>{h(vsw.get("Type",vsw.get("SwitchType","")))}</td>'
                      f'<td>{h(vsw.get("NetAdapter",vsw.get("NetAdapterName","")))}</td></tr>\n')
virt_vsw_html += '</table>\n' + top_link('virt')

## VIRT HOST CONFIG
virt_hc_html = f'''<table>
<tr><th>Setting</th><th>Value</th></tr>
<tr><td>Live Migration</td><td>{pill("Enabled","green") if hv_hv.get("LiveMigrationEnabled") else pill("Disabled","gray")}</td></tr>
<tr><td>NUMA Spanning</td><td>{pill("Enabled","green") if hv_hv.get("NumaSpanningEnabled") else pill("Disabled","gray")}</td></tr>
<tr><td>Default VM Path</td><td><code>{h(hv_hv.get("DefaultVMPath",""))}</code></td></tr>
<tr><td>Default VHD Path</td><td><code>{h(hv_hv.get("DefaultVHDPath",""))}</code></td></tr>
<tr><td>Virtual Switches</td><td>{len(vsw_list)}</td></tr>
</table>
''' + top_link('virt')

# ── ASSEMBLE HTML ─────────────────────────────────────────────────────────────
CSS = '''
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Segoe UI', Arial, sans-serif; background: #f5f4f8; color: #271e41; font-size: 10pt; }
.wrap { max-width: 1040px; margin: 0 auto; padding: 20px; }
.tab-nav { display: flex; gap: 4px; margin-bottom: -1px; flex-wrap: wrap; }
.tab-btn { padding: 8px 18px; background: #ddd9ee; border: 1px solid #c0b8d8; border-bottom: none; border-radius: 6px 6px 0 0; cursor: pointer; font-size: 9.5pt; color: #271e41; font-weight: 600; }
.tab-btn.active { background: white; border-bottom: 1px solid white; color: #5b1fa4; }
.tab-btn.has-critical { border-top: 3px solid #d63638; }
.tab-btn.has-warning  { border-top: 3px solid #f5a623; }
.tab-content { display: none; }
.tab-content.active { display: block; }
.card { background: white; border-radius: 0 8px 8px 8px; padding: 24px; margin-bottom: 18px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); border: 1px solid #e8e4f0; }
.card-title { font-size: 15px; font-weight: 700; color: #271e41; text-transform: uppercase; letter-spacing: 1px; border-bottom: 2px solid #5b1fa4; padding-bottom: 8px; margin-bottom: 16px; display: flex; justify-content: space-between; align-items: center; }
.collapse-btn { background: none; border: 1px solid #c4b5fd; border-radius: 4px; color: #5b1fa4; font-size: 8pt; padding: 2px 8px; cursor: pointer; font-weight: 600; flex-shrink: 0; }
.card-body.collapsed { display: none; }
table { width: 100%; border-collapse: collapse; font-size: 9.5pt; }
th { background: #271e41; color: #fff; font-weight: 600; padding: 7px 12px; text-align: left; }
td { padding: 6px 12px; border: 1px solid #d0cce0; vertical-align: top; }
tr:nth-child(even) td { background: #f5f4f8; }
.flag-critical { background: #fff0f0; border-left: 4px solid #d63638; border-radius: 0 6px 6px 0; padding: 12px 16px; margin-bottom: 8px; }
.flag-warning  { background: #fff8e1; border-left: 4px solid #f5a623; border-radius: 0 6px 6px 0; padding: 12px 16px; margin-bottom: 8px; }
.flag-info     { background: #f0f4ff; border-left: 4px solid #5b1fa4; border-radius: 0 6px 6px 0; padding: 12px 16px; margin-bottom: 8px; }
.flag-ok       { background: #f0fdf0; border-left: 4px solid #20c800; border-radius: 0 6px 6px 0; padding: 12px 16px; margin-bottom: 8px; }
.flag-label    { font-size: 8.5pt; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }
.flag-critical .flag-label { color: #d63638; }
.flag-warning  .flag-label { color: #b07a00; }
.flag-info     .flag-label { color: #5b1fa4; }
.flag-ok       .flag-label { color: #20c800; }
.flag-detail   { font-size: 9.5pt; color: #271e41; margin-top: 4px; }
.pill { display: inline-block; padding: 2px 9px; border-radius: 12px; font-size: 8pt; font-weight: 700; }
.pill-red    { background: #fee2e2; color: #991b1b; }
.pill-yellow { background: #fef3c7; color: #92400e; }
.pill-green  { background: #d1fae5; color: #065f46; }
.pill-gray   { background: #f3f4f6; color: #374151; }
.pill-purple { background: #ede9fe; color: #5b1fa4; }
.stat-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 16px; }
.stat-box { background: #f5f4f8; border-radius: 6px; padding: 12px; text-align: center; }
.stat-num  { font-size: 22px; font-weight: 700; color: #5b1fa4; }
.stat-lbl  { font-size: 8pt; color: #6b6080; text-transform: uppercase; letter-spacing: 0.5px; margin-top: 2px; }
.role-grid { display: flex; flex-wrap: wrap; gap: 8px; }
.role-badge { background: #ede9fe; color: #5b1fa4; border-radius: 4px; padding: 4px 12px; font-size: 9pt; font-weight: 600; border: 1px solid #c4b5fd; }
.disk-bar-bg  { background: #e9e4f5; border-radius: 4px; height: 10px; width: 100%; margin-top: 4px; }
.disk-bar-fill { height: 10px; border-radius: 4px; }
details summary { cursor: pointer; font-weight: 600; color: #5b1fa4; padding: 6px 0; list-style: none; }
details summary::before { content: '\\25B6  '; font-size: 9pt; }
details[open] summary::before { content: '\\25BC  '; }
.sub-title { font-size: 13px; font-weight: 700; color: #5b1fa4; text-transform: uppercase; letter-spacing: 0.8px; margin: 16px 0 8px; }
.meta-line { font-size: 8.5pt; color: #6b6080; margin-bottom: 2px; }
.view-bar{background:#1a1432;padding:8px 28px;display:flex;align-items:center;gap:10px;position:sticky;top:0;z-index:200;border-radius:0 0 6px 6px;}
.view-lbl{font-size:8pt;color:#a89bc8;text-transform:uppercase;letter-spacing:.5px;font-weight:700;}
.view-btn{background:transparent;border:1px solid #4a3a6a;border-radius:4px;color:#c4b5fd;font-size:9pt;padding:4px 14px;cursor:pointer;font-weight:600;transition:all .15s;}
.view-btn:hover{background:#2d2060;}
.view-btn.v-active{background:#5b1fa4;color:white;border-color:#5b1fa4;}
.view-desc{font-size:8.5pt;color:#7c6b9e;margin-left:6px;}
.sbr-only{display:none !important;}
body.view-sbr .sbr-only{display:block !important;}
body.view-sbr .hide-sbr{display:none !important;}
'''

JS = '''
function showTab(id) {
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  document.querySelector('[data-tab="'+id+'"]').classList.add('active');
}
function toggleCard(btn) {
  var body = btn.closest('.card').querySelector('.card-body');
  if (body.classList.contains('collapsed')) {
    body.classList.remove('collapsed');
    btn.textContent = '\u25b2 Collapse';
  } else {
    body.classList.add('collapsed');
    btn.textContent = '\u25bc Expand';
  }
}
var VIEW_DESCS = {
  basic: "Full detail \u2014 SE &amp; CSM view",
  adv:   "Full technical detail \u2014 SE view",
  sbr:   "Executive health dashboard \u2014 client &amp; leadership"
};
function setView(v) {
  document.body.classList.remove("view-basic","view-adv","view-sbr");
  document.body.classList.add("view-" + v);
  document.querySelectorAll(".view-btn").forEach(function(b){ b.classList.remove("v-active"); });
  var btn = document.getElementById("vbtn-" + v);
  if (btn) btn.classList.add("v-active");
  var desc = document.getElementById("view-desc");
  if (desc) desc.innerHTML = VIEW_DESCS[v];
  try { localStorage.setItem("sdView", v); } catch(e) {}
}
window.addEventListener("DOMContentLoaded", function() {
  var saved = "adv";
  try { saved = localStorage.getItem("sdView") || "adv"; } catch(e) {}
  setView(saved);
});
'''

out = f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{CLIENT} Server Discovery Report &mdash; {DATE}</title>
<style>{CSS}</style>
</head>
<body class="view-adv">
<div class="view-bar">
  <span class="view-lbl">View:</span>
  <button class="view-btn" id="vbtn-basic" onclick="setView('basic')">Basic</button>
  <button class="view-btn v-active" id="vbtn-adv" onclick="setView('adv')">Advanced</button>
  <button class="view-btn" id="vbtn-sbr" onclick="setView('sbr')">SBR</button>
  <span class="view-desc" id="view-desc">Full technical detail &mdash; SE view</span>
</div>
<div class="wrap">

<div style="background:#271e41;padding:16px 28px;display:flex;justify-content:space-between;align-items:center;border-radius:0 0 4px 4px;margin-bottom:20px">
  <div style="display:flex;align-items:center">
    <img src="{LOGO_B64}" alt="Magna5" style="height:40px;margin-right:20px">
    <div>
      <div style="color:white;font-size:16px;font-weight:700;letter-spacing:1px">SERVER DISCOVERY REPORT</div>
      <div style="color:#a89bc8;font-size:9pt;margin-top:2px">{CLIENT} Environment &mdash; {DATE}</div>
    </div>
  </div>
  <div style="text-align:right">
    <div style="color:#c4b5fd;font-size:8.5pt">Collected: {DATE}</div>
    <div style="color:#a89bc8;font-size:8pt;margin-top:2px">Magna5 Solutions Engineering</div>
  </div>
</div>

<div class="tab-nav">
<button class="tab-btn active{tab_cls(hv_crit,hv_warn)}" data-tab="tab-hv02" onclick="showTab('tab-hv02')">MEKEHV02 &middot; Server 2022</button>
<button class="tab-btn {tab_cls(dc_crit,dc_warn)}" data-tab="tab-dc02" onclick="showTab('tab-dc02')">MEKEDC02 &middot; Server 2022 (VM)</button>
<button class="tab-btn " data-tab="tab-virt" onclick="showTab('tab-virt')">VIRTUALIZATION &middot; MEKEHV02</button>
</div>

<div id="tab-hv02" class="tab-content active">
{hv_sbr}
{hv_nav}
{card("hv02-alerts", "Alerts", hv_alerts_html)}
{card("hv02-overview", "System Overview", hv_overview_html)}
{card("hv02-hardware", "Hardware", hv_hw_html)}
{card("hv02-apps", "Installed Applications", hv_apps_html, collapsed=True)}
{card("hv02-roles", "Roles &amp; Features", hv_roles_html, collapsed=True)}
{card("hv02-roleconfig", "Role Configuration", hv_rc_html, collapsed=True)}
{card("hv02-disks", "Disk Storage", hv_disk_html)}
{card("hv02-network", "Network", hv_net_html)}
{card("hv02-lports", "Listening Ports", hv_lports_html, collapsed=True)}
{card("hv02-services", "Services", hv_svc_html, collapsed=True)}
{card("hv02-svc-anomalies", "Service Anomalies", hv_anom_html)}
</div>

<div id="tab-dc02" class="tab-content">
{dc_sbr}
{dc_nav}
{card("dc02-alerts", "Alerts", dc_alerts_html)}
{card("dc02-overview", "System Overview", dc_overview_html)}
{card("dc02-hardware", "Hardware", dc_hw_html)}
{card("dc02-apps", "Installed Applications", dc_apps_html, collapsed=True)}
{card("dc02-roles", "Roles &amp; Features", dc_roles_html, collapsed=True)}
{card("dc02-roleconfig", "Role Configuration", dc_rc_html, collapsed=True)}
{card("dc02-disks", "Disk Storage", dc_disk_html)}
{card("dc02-network", "Network", dc_net_html)}
{card("dc02-lports", "Listening Ports", dc_lports_html, collapsed=True)}
{card("dc02-services", "Services", dc_svc_html, collapsed=True)}
{card("dc02-svc-anomalies", "Service Anomalies", dc_anom_html)}
{card("dc02-shares", "File Shares", dc_shares_html)}
</div>

<div id="tab-virt" class="tab-content">
{virt_sbr}
{virt_nav}
{card("virt-summary", "Host Summary", virt_host_html)}
{card("virt-vms", "Virtual Machines", virt_vms_html)}
{card("virt-vswitches", "Virtual Switches", virt_vsw_html)}
{card("virt-hostcfg", "Host Configuration", virt_hc_html)}
</div>

</div>
<script>{JS}</script>
</body>
</html>'''

with open(OUTPUT, 'w', encoding='utf-8') as f:
    f.write(out)

size = os.path.getsize(OUTPUT)
print(f"Written: {OUTPUT}")
print(f"Size: {size:,} bytes ({size//1024} KB)")
print(f"HV02: {hv_crit} critical, {hv_warn} warning")
print(f"DC02: {dc_crit} critical, {dc_warn} warning")
