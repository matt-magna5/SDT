"""
gen_qes_report.py — QES Discovery Report
Reads 3 JSON files from Discovery-Session-2026-04-10-2328, produces HTML.
Tabs: QES-RDS-01 | QES-DATA-DC | VIRTUALIZATION (vSphere)
Stub tabs: QES-PENT-01 (pending) | QES-HNY-01 (Linux) | QES-ADL-01 (Linux)
           QES-DATACENTER-DC-02 (off) | VMware vCenter (appliance)
"""
import json, html as htmlmod, re, sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# ── CONFIG ────────────────────────────────────────────────────────────────────
BASE        = r'C:/Users/matt/OneDrive - Magna5/M5 Obsidian Vault/M5 Ops/Root/zzzTest/Server Discovery Tool'
SESSION_DIR = BASE + '/Discovery-Session-2026-04-10-2328'
DATE        = '2026-04-10'
CLIENT      = 'QES'
OUTPUT      = SESSION_DIR + f'/QES-DiscoveryReport-{DATE}.html'
LOGO_FILE   = r'C:/Users/matt/AppData/Local/Temp/m5_logo_b64.txt'

# ── LOAD DATA ─────────────────────────────────────────────────────────────────
def jload(path):
    with open(path, encoding='utf-8-sig') as f: return json.load(f)

rds = jload(f'{SESSION_DIR}/QES-RDS-01-discovery-{DATE}.json')
dc  = jload(f'{SESSION_DIR}/QES-DATA-DC-discovery-{DATE}.json')
inv = jload(f'{SESSION_DIR}/10.200.1.12-inventory-{DATE}.json')

# Load vsphere-perf JSON for datastore mapping (from parse_ntnx_collector.py output)
import glob as _glob
_perf_files = _glob.glob(SESSION_DIR + '/vsphere-perf*.json')
perf = jload(_perf_files[0]) if _perf_files else {}
_vm_ds_map = {v['Name']: v.get('Datastore', '—') for v in perf.get('VMs', [])}

with open(LOGO_FILE) as f: LOGO_B64 = f.read().strip()

# ── HELPERS ───────────────────────────────────────────────────────────────────
h = htmlmod.escape

def as_list(v):
    """Flatten potentially-nested PS serialization artefacts into flat list of dicts."""
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

def stor_row(drv, total, pct):
    c  = '#d63638' if pct>=85 else ('#f5a623' if pct>=70 else '#5b1fa4')
    pc = 'red' if pct>=85 else ('yellow' if pct>=70 else 'green')
    return (f'<tr style="padding:6px 0">'
            f'<td style="font-weight:700;white-space:nowrap">{h(drv)}</td>'
            f'<td style="font-size:8.5pt;color:#6b6080">{total:.1f} GB total</td>'
            f'<td>{pill(f"{pct}% used", pc)}'
            f'{disk_bar(pct)}</td></tr>\n')

def apps_table(apps_list, tid):
    rows = ''
    for a in sorted(apps_list, key=lambda x: x.get('Name','')):
        sev = a.get('FlagSeverity','none')
        sc  = 'green' if sev=='none' else ('red' if sev=='critical' else 'yellow')
        idate = str(a.get('InstallDate','') or '').strip()
        rows += f'<tr><td>{h(a.get("Name","?"))}</td><td>{h(a.get("Version","") or "")}</td><td>{h(a.get("Publisher","") or "")}</td><td>{h(idate)}</td><td>{pill(sev,sc)}</td></tr>\n'
    return (f'<table><tr><th>Name</th><th>Version</th><th>Publisher</th><th>Install Date</th><th>Flag</th></tr>\n'
            + rows + '</table>\n') + top_link(tid)

# ── PARSE JSON DATA ───────────────────────────────────────────────────────────
rds_sys = rds.get('System', {})
rds_hw  = rds.get('Hardware', {})
dc_sys  = dc.get('System', {})
dc_hw   = dc.get('Hardware', {})

# RDS apps — nested structure [str, [app_dicts]]
_rds_apps_raw = rds.get('Apps', [])
rds_apps = []
for item in _rds_apps_raw:
    if isinstance(item, list):  rds_apps.extend([x for x in item if isinstance(x, dict)])
    elif isinstance(item, dict): rds_apps.append(item)

# DC apps — same structure
_dc_apps_raw = dc.get('Apps', [])
dc_apps = []
for item in _dc_apps_raw:
    if isinstance(item, list):  dc_apps.extend([x for x in item if isinstance(x, dict)])
    elif isinstance(item, dict): dc_apps.append(item)

rds_disks = as_list(rds.get('Disks', []))
dc_disks  = as_list(dc.get('Disks', []))

# AD data — nested [ps_garbage_str, {actual_ad_dict}]
dc_ad_raw = dc.get('AD', [])
dc_ad = next((x for x in as_list(dc_ad_raw) if 'ForestName' in x), {})

dc_roles_data     = dc.get('Roles', {})
dc_roles_list     = as_list(dc_roles_data.get('InstalledRoles', []))
dc_features_list  = as_list(dc_roles_data.get('InstalledFeatures', []))
dc_role_names     = [r.get('DisplayName', r.get('Name','')) for r in dc_roles_list]
feat_names_dc     = [f.get('Name','') for f in dc_features_list]
has_smb1          = any('FS-SMB1' in n for n in feat_names_dc)

rds_roles_data    = rds.get('Roles', {})
rds_roles_list    = as_list(rds_roles_data.get('InstalledRoles', []))
rds_features_list = as_list(rds_roles_data.get('InstalledFeatures', []))
rds_role_names    = [r.get('DisplayName', r.get('Name','')) for r in rds_roles_list]

stale_users = as_list(dc_ad.get('StaleUsers', []))
stale_comps = as_list(dc_ad.get('StaleComputers', []))

# vSphere inventory
inv_vms        = inv.get('VMs', [])
inv_datastores = inv.get('Datastores', [])
inv_esx_hosts  = inv.get('ESXHosts', [])

# Agent detection helpers
def find_app(apps, *keywords):
    kw = [k.lower() for k in keywords]
    return next((a['Name'] for a in apps if any(k in a.get('Name','').lower() for k in kw)), None)

def find_publisher(apps, *keywords):
    kw = [k.lower() for k in keywords]
    return next((a['Name'] for a in apps if any(k in a.get('Publisher','').lower() for k in kw)), None)

def find_nable_label(apps):
    """Return clean N-able RMM label (e.g. 'N-able RMM v2025.4.1016'), or None."""
    nable_apps = [a for a in apps if 'n-able' in (a.get('Publisher','') or '').lower()]
    if not nable_apps:
        return None
    win_agent = next((a for a in nable_apps if a.get('Name','').lower() in ('windows agent', 'msp core agent')), None)
    ver = f" v{win_agent['Version']}" if win_agent and win_agent.get('Version') else ''
    return f"N-able RMM{ver}"

rds_edr  = find_app(rds_apps, 'sentinel')
rds_hunt = find_app(rds_apps, 'huntress agent')
rds_adl  = find_app(rds_apps, 'adlumin')
rds_rmm  = find_nable_label(rds_apps)
rds_comm = find_app(rds_apps, 'commvault')

dc_edr   = find_app(dc_apps, 'sentinel')
dc_hunt  = find_app(dc_apps, 'huntress')
dc_adl   = find_app(dc_apps, 'adlumin')
dc_rmm   = find_nable_label(dc_apps)
dc_entra = find_app(dc_apps, 'entra connect sync', 'azure ad connect')

dc_shares_raw  = as_list(dc.get('FileShares',{}).get('Shares',[]))
dc_shares_real = [s for s in dc_shares_raw
                  if not s.get('Name','').startswith('$')
                  and s.get('Name','') not in ('ADMIN$','IPC$','C$','print$')]

domain_name = dc_ad.get('ForestName', dc_sys.get('Domain','QES.CORP'))
fl_lbl      = str(dc_ad.get('DomainFL',''))

# ── FLAG DERIVATION ───────────────────────────────────────────────────────────
def build_rds_flags():
    flags = []
    # Disk usage
    for d in rds_disks:
        pct = d.get('UsedPct', 0)
        drv, free, total = d.get('Drive','?'), d.get('FreeGB',0), d.get('TotalGB',0)
        if pct >= 85:
            flags.append(('critical', f'Disk {drv} Near Capacity',
                f'{drv}: {pct}% used — only {free:.1f} GB free of {total:.1f} GB. Risk of service disruption.'))
        elif pct >= 70:
            flags.append(('warning', f'Disk {drv} Space Moderate',
                f'{drv}: {pct}% used ({free:.1f} GB free of {total:.1f} GB). Monitor closely.'))
    # RAM
    ram_t = rds_hw.get('RAMTotalGB', 0)
    ram_f = rds_hw.get('RAMAvailGB', 0)
    if ram_t > 0 and (1 - ram_f/ram_t) > 0.80:
        pct = int((1 - ram_f/ram_t)*100)
        flags.append(('warning', f'High Memory Utilization ({pct}%)',
            f'RAM at {pct}% used ({ram_f:.1f} GB free of {ram_t:.1f} GB) on the RDS server. '
            f'With {int(ram_t)}GB total, this is likely expected for an analytics RDS server, '
            f'but monitor during peak user load.'))
    return flags

def build_dc_flags():
    flags = []
    # Critical: RAM
    ram_t = dc_hw.get('RAMTotalGB', 0)
    ram_f = dc_hw.get('RAMAvailGB', 0)
    if ram_t > 0:
        pct = int((1 - ram_f/ram_t)*100)
        if pct >= 85:
            flags.append(('critical', f'Domain Controller RAM Critical ({pct}% used)',
                f'QES-DATA-DC has only {ram_f:.2f} GB free of {ram_t:.0f} GB RAM ({pct}% used). '
                f'A DC this memory-constrained is at risk of AD authentication failures, '
                f'DNS query drops, and DHCP service instability under load. '
                f'Immediate RAM allocation increase required.'))
        elif pct >= 75:
            flags.append(('warning', f'High Memory Utilization ({pct}%)',
                f'RAM at {pct}% used ({ram_f:.1f} GB free of {ram_t:.0f} GB).'))
    # SMB1
    if has_smb1:
        flags.append(('critical','SMB 1.0/CIFS Enabled — Critical Security Risk',
            'SMB 1.0 is the attack vector used by WannaCry, NotPetya, and EternalBlue ransomware. '
            'Disable immediately: Remove-WindowsFeature FS-SMB1'))
    # AD Functional Level
    if fl_lbl and 'Windows2008R2' in fl_lbl:
        flags.append(('critical', 'AD Functional Level — Windows 2008 R2 (Extremely Outdated)',
            f'Domain functional level is Windows 2008 R2 — nearly 17 years old. This blocks modern '
            f'LAPS, Protected Users group features, Kerberos AES-only enforcement, and enhanced '
            f'security controls. Upgrade path: raise to 2016 or 2019 FL. '
            f'This is a hard blocker for many modern security hardening measures.'))
    elif fl_lbl and str(fl_lbl) < '2016':
        flags.append(('warning', f'AD Functional Level {fl_lbl} — Upgrade Recommended',
            f'Domain functional level is {fl_lbl}. Upgrading to 2016+ enables modern LAPS, '
            f'additional security features, and better Kerberos options.'))
    # Stale users
    if len(stale_users) > 20:
        flags.append(('warning', f'Stale AD Accounts — {len(stale_users)} Inactive Users',
            f'{len(stale_users)} accounts inactive 90+ days. Attack surface risk. '
            f'Audit and disable/remove.'))
    # Stale comps
    if len(stale_comps) > 15:
        flags.append(('warning', f'Stale Computer Accounts — {len(stale_comps)} Devices',
            f'{len(stale_comps)} computer accounts inactive 90+ days. Clean up to reduce attack surface.'))
    return flags

rds_flags = build_rds_flags()
dc_flags  = build_dc_flags()
rds_crit  = sum(1 for f in rds_flags if f[0]=='critical')
rds_warn  = sum(1 for f in rds_flags if f[0]=='warning')
dc_crit   = sum(1 for f in dc_flags  if f[0]=='critical')
dc_warn   = sum(1 for f in dc_flags  if f[0]=='warning')

def tab_cls(crit, warn):
    if crit: return ' has-critical'
    if warn: return ' has-warning'
    return ''

def sbr_badge(crit, warn):
    if crit: return '&#128308; CRITICAL'
    if warn: return '&#9888;&#65039; ATTENTION'
    return '&#9989; HEALTHY'

def sbr_grad(crit, warn):
    if crit: return 'linear-gradient(135deg,#d63638,#b92b2e)'
    if warn: return 'linear-gradient(135deg,#f5a623,#e0901a)'
    return 'linear-gradient(135deg,#20c800,#158f00)'

# ── RDS-01 SBR PANEL ─────────────────────────────────────────────────────────
rds_up    = rds_sys.get('UptimeDays', 0)
rds_ram_t = rds_hw.get('RAMTotalGB', 0)
rds_ram_f = rds_hw.get('RAMAvailGB', 0)
rds_ram_pct = int((1 - rds_ram_f/rds_ram_t)*100) if rds_ram_t else 0

rds_disk_rows = ''.join(
    stor_row(d.get('Drive','?'), d.get('TotalGB',0), d.get('UsedPct',0)) for d in rds_disks)

def agent_row(icon, label, name, last=False):
    sep = '' if last else 'border-bottom:1px solid #f0edf8;'
    return (f'<div style="display:flex;align-items:center;gap:12px;padding:8px 0;{sep}">'
            f'<span style="font-size:18px;flex-shrink:0;width:26px;text-align:center">{icon}</span>'
            f'<div style="flex:1">'
            f'<div style="display:flex;align-items:center;gap:8px;">'
            f'<span style="font-size:9pt;font-weight:700;color:#271e41;">{label}</span>'
            f'<span class="pill pill-green" style="font-size:7.5pt;">ENABLED</span></div>'
            f'<div style="font-size:8.5pt;color:#6b6080;margin-top:1px;">{h(str(name))}</div>'
            f'</div></div>\n')

rds_agents = ''
if rds_edr:  rds_agents += agent_row('&#128737;', 'EDR',       rds_edr)
if rds_hunt: rds_agents += agent_row('&#128269;', 'MDR',       rds_hunt)
if rds_rmm:  rds_agents += agent_row('&#128421;', 'RMM',       rds_rmm)
if rds_comm: rds_agents += agent_row('&#128452;', 'BDR Agent', rds_comm, last=True)

# LOB apps summary
lob_apps = [a['Name'] for a in rds_apps if any(k in a.get('Name','').lower() for k in ['sas 9','stata18','stata 18','stata 15'])]
lob_summary = ', '.join(list(dict.fromkeys(lob_apps))[:4]) or 'None detected'

# Pre-compute mini boxes that have tricky nested quotes
_rds_os_box = mini_box("OS &amp; System",
    f'<table style="width:100%;font-size:9pt;border-collapse:collapse;">'
    f'<tr><td style="color:#6b6080;width:110px;padding:3px 0">Platform</td><td>VMware VM on qes-vh01.qes.corp</td></tr>'
    f'<tr><td style="color:#6b6080;padding:3px 0">OS</td><td>Server 2019 Standard {pill("EOL Jan 2029","green")}</td></tr>'
    f'<tr><td style="color:#6b6080;padding:3px 0">Last Reboot</td><td>{h(rds_sys.get("LastBoot","?").split(" ")[0])}'
    f' &nbsp;{pill(f"{rds_up:.0f} days ago","green" if rds_up<30 else "yellow")}</td></tr>'
    f'<tr><td style="color:#6b6080;padding:3px 0">RAM</td><td>{rds_ram_t:.0f} GB total &nbsp;'
    f'{pill(f"{rds_ram_f:.0f} GB free ({100-rds_ram_pct}%)","green" if rds_ram_pct<80 else "yellow")}</td></tr>'
    f'<tr><td style="color:#6b6080;padding:3px 0">Domain</td><td>{pill(rds_sys.get("Domain",""),"green")}</td></tr>'
    f'</table>')
_rds_roles_box = mini_box("Server Roles",
    '<div style="display:flex;flex-wrap:wrap;gap:6px;">'
    + ''.join(f'<span class="role-badge">{h(r)}</span>' for r in rds_role_names)
    + '</div>')
_rds_lob_box = mini_box("LOB Applications",
    f'<div style="font-size:9pt;color:#271e41;">{h(lob_summary)}</div>', last=True)
_rds_agent_box = mini_box("Installed Agents",
    rds_agents or '<span style="color:#9b8fb0;font-style:italic">None detected</span>')
_rds_stor_box = mini_box("Storage",
    f'<table style="width:100%;border-collapse:collapse;">'
    f'<tr><th style="text-align:left;font-size:8pt;padding:4px 0;color:#6b6080;font-weight:600">Drive</th>'
    f'<th style="font-size:8pt;color:#6b6080;font-weight:600">Size</th>'
    f'<th style="font-size:8pt;color:#6b6080;font-weight:600">Usage</th></tr>'
    f'{rds_disk_rows}</table>', last=True)

rds_sbr = f'''<div class="sbr-only">
<div style="background:{sbr_grad(rds_crit,rds_warn)};border-radius:10px 10px 0 0;padding:16px 24px;display:flex;justify-content:space-between;align-items:center;margin-bottom:0;">
  <div>
    <div style="font-size:18px;font-weight:700;color:#fff;letter-spacing:.3px;">QES-RDS-01</div>
    <div style="font-size:9pt;color:rgba(255,255,255,.85);margin-top:3px;">Analytics RDS Server &middot; VMware VM &middot; Server 2019</div>
  </div>
  <div style="display:flex;align-items:center;gap:14px;">
    <div style="text-align:center;background:rgba(255,255,255,.{25 if rds_crit else 15});border-radius:8px;padding:8px 14px;">
      <div style="font-size:22px;font-weight:700;color:#fff;">{rds_crit}</div>
      <div style="font-size:7pt;color:rgba(255,255,255,.8);text-transform:uppercase;font-weight:600;">Critical</div>
    </div>
    <div style="text-align:center;background:rgba(255,255,255,.{25 if rds_warn else 15});border-radius:8px;padding:8px 14px;">
      <div style="font-size:22px;font-weight:700;color:#fff;">{rds_warn}</div>
      <div style="font-size:7pt;color:rgba(255,255,255,.8);text-transform:uppercase;font-weight:600;">Warning</div>
    </div>
    <span style="background:rgba(255,255,255,.22);color:#fff;font-size:10pt;font-weight:700;padding:6px 18px;border-radius:20px;border:1.5px solid rgba(255,255,255,.5);">{sbr_badge(rds_crit,rds_warn)}</span>
  </div>
</div>
<div style="background:white;border-radius:0 0 10px 10px;border:1px solid #e8e4f0;border-top:none;box-shadow:0 4px 14px rgba(0,0,0,.07);padding:20px 24px;margin-bottom:16px;">
<div style="display:grid;grid-template-columns:1fr 1fr;gap:18px;">
<div>
{_rds_os_box}
{_rds_roles_box}
{_rds_lob_box}
</div>
<div>
{_rds_agent_box}
{_rds_stor_box}
</div>
</div>
</div>
</div>
'''

# ── DATA-DC SBR PANEL ─────────────────────────────────────────────────────────
dc_up     = dc_sys.get('UptimeDays', 0)
dc_ram_t  = dc_hw.get('RAMTotalGB', 0)
dc_ram_f  = dc_hw.get('RAMAvailGB', 0)
dc_ram_pct_used = int((1 - dc_ram_f/dc_ram_t)*100) if dc_ram_t else 0

dc_disk_rows = ''.join(
    stor_row(d.get('Drive','?'), d.get('TotalGB',0), d.get('UsedPct',0)) for d in dc_disks)

dc_agents = ''
if dc_edr:   dc_agents += agent_row('&#128737;', 'EDR',        dc_edr)
if dc_hunt:  dc_agents += agent_row('&#128269;', 'MDR',        dc_hunt)
if dc_rmm:   dc_agents += agent_row('&#128421;', 'RMM',        dc_rmm)
if dc_entra: dc_agents += agent_row('&#9729;&#65039;', 'Entra Sync', 'Microsoft Entra Connect Sync', last=True)

role_badges = ''.join(f'<span class="role-badge">{h(r)}</span>' for r in dc_role_names)

# Build DC SBR ram pill
_dc_ram_str = (f'{dc_ram_f:.2f} GB free — CRITICAL' if dc_ram_pct_used>=85 else f'{dc_ram_f:.1f} GB free ({100-dc_ram_pct_used}%)')
_dc_ram_clr = 'red' if dc_ram_pct_used>=85 else 'green'
_dc_fl_pill = (pill(fl_lbl,"red") if "2008" in fl_lbl else (pill(fl_lbl,"yellow") if fl_lbl < "2019" else pill(fl_lbl,"green")))
_dc_fl_upg  = pill("upgrade required","red") if "2008" in fl_lbl else ""
_dc_count_str = str(dc_ad.get('DCCount','?')) + ' DCs'
_dc_count_clr = 'green' if (dc_ad.get('DCCount',0) or 0) >= 2 else 'yellow'
_dc_stale_pill = pill(f"{len(stale_users)} stale (90+ days)","yellow") if stale_users else ""

_dc_os_box = mini_box("OS &amp; System",
    f'<table style="width:100%;font-size:9pt;border-collapse:collapse;">'
    f'<tr><td style="color:#6b6080;width:110px;padding:3px 0">Platform</td><td>VMware VM on qes-vh01.qes.corp</td></tr>'
    f'<tr><td style="color:#6b6080;padding:3px 0">OS</td><td>Server 2019 Standard {pill("EOL Jan 2029","green")}</td></tr>'
    f'<tr><td style="color:#6b6080;padding:3px 0">Last Reboot</td><td>{h(dc_sys.get("LastBoot","?").split(" ")[0])}'
    f' &nbsp;{pill(f"{dc_up:.0f} days ago","green" if dc_up<30 else "yellow")}</td></tr>'
    f'<tr><td style="color:#6b6080;padding:3px 0">RAM</td><td>{dc_ram_t:.0f} GB total &nbsp;{pill(_dc_ram_str,_dc_ram_clr)}</td></tr>'
    f'<tr><td style="color:#6b6080;padding:3px 0">Domain</td><td>{pill(domain_name,"green")}</td></tr>'
    f'</table>')
_dc_ad_box = mini_box("Active Directory",
    f'<table style="width:100%;font-size:9pt;border-collapse:collapse;">'
    f'<tr><td style="color:#6b6080;width:130px;padding:3px 0">Forest</td><td>{h(domain_name)}</td></tr>'
    f'<tr><td style="color:#6b6080;padding:3px 0">DC Count</td><td>{pill(_dc_count_str,_dc_count_clr)}</td></tr>'
    f'<tr><td style="color:#6b6080;padding:3px 0">Functional Level</td><td>{_dc_fl_pill} &nbsp;{_dc_fl_upg}</td></tr>'
    f'<tr><td style="color:#6b6080;padding:3px 0">Users</td><td>{dc_ad.get("UserCount","?")} total &nbsp;{_dc_stale_pill}</td></tr>'
    f'</table>', last=True)
_dc_agent_box = mini_box("Installed Agents", dc_agents or '<span style="color:#9b8fb0;font-style:italic">None detected</span>')
_dc_roles_box = mini_box("Server Roles", f'<div style="display:flex;flex-wrap:wrap;gap:6px;">{role_badges}</div>')
_dc_stor_box  = mini_box("Storage",
    f'<table style="width:100%;border-collapse:collapse;">'
    f'<tr><th style="text-align:left;font-size:8pt;padding:4px 0;color:#6b6080;font-weight:600">Drive</th>'
    f'<th style="font-size:8pt;color:#6b6080;font-weight:600">Size</th>'
    f'<th style="font-size:8pt;color:#6b6080;font-weight:600">Usage</th></tr>'
    f'{dc_disk_rows}</table>', last=True)

dc_sbr = (
f'<div class="sbr-only">\n'
f'<div style="background:{sbr_grad(dc_crit,dc_warn)};border-radius:10px 10px 0 0;padding:16px 24px;display:flex;justify-content:space-between;align-items:center;">\n'
f'  <div>\n'
f'    <div style="font-size:18px;font-weight:700;color:#fff;letter-spacing:.3px;">QES-DATA-DC</div>\n'
f'    <div style="font-size:9pt;color:rgba(255,255,255,.85);margin-top:3px;">Domain Controller (VM) &middot; {h(domain_name)} &middot; Server 2019</div>\n'
f'  </div>\n'
f'  <div style="display:flex;align-items:center;gap:14px;">\n'
f'    <div style="text-align:center;background:rgba(255,255,255,.{25 if dc_crit else 15});border-radius:8px;padding:8px 14px;">\n'
f'      <div style="font-size:22px;font-weight:700;color:#fff;">{dc_crit}</div>\n'
f'      <div style="font-size:7pt;color:rgba(255,255,255,.8);text-transform:uppercase;font-weight:600;">Critical</div>\n'
f'    </div>\n'
f'    <div style="text-align:center;background:rgba(255,255,255,.{25 if dc_warn else 15});border-radius:8px;padding:8px 14px;">\n'
f'      <div style="font-size:22px;font-weight:700;color:#fff;">{dc_warn}</div>\n'
f'      <div style="font-size:7pt;color:rgba(255,255,255,.8);text-transform:uppercase;font-weight:600;">Warning</div>\n'
f'    </div>\n'
f'    <span style="background:rgba(255,255,255,.22);color:#fff;font-size:10pt;font-weight:700;padding:6px 18px;border-radius:20px;border:1.5px solid rgba(255,255,255,.5);">{sbr_badge(dc_crit,dc_warn)}</span>\n'
f'  </div>\n'
f'</div>\n'
f'<div style="background:white;border-radius:0 0 10px 10px;border:1px solid #e8e4f0;border-top:none;box-shadow:0 4px 14px rgba(0,0,0,.07);padding:20px 24px;margin-bottom:16px;">\n'
f'<div style="display:grid;grid-template-columns:1fr 1fr;gap:18px;">\n'
f'<div>\n'
+ _dc_os_box + _dc_ad_box +
f'</div>\n'
f'<div>\n'
+ _dc_agent_box + _dc_roles_box + _dc_stor_box +
f'</div>\n'
f'</div>\n'
f'</div>\n'
f'</div>\n'
)

# ── VIRT (vSphere) SBR PANEL ─────────────────────────────────────────────────
running_vms = [v for v in inv_vms if v.get('PowerState') == 'POWERED_ON']
esx_host    = inv_esx_hosts[0] if inv_esx_hosts else {}

total_ds_tb = sum(d.get('CapacityGB',0) for d in inv_datastores) / 1024
used_ds_tb  = sum((d.get('CapacityGB',0) - d.get('FreeGB',0)) for d in inv_datastores) / 1024

virt_sbr = f'''<div class="sbr-only">
<div style="background:linear-gradient(135deg,#5b1fa4,#3d1270);border-radius:10px 10px 0 0;padding:16px 24px;display:flex;justify-content:space-between;align-items:center;margin-bottom:0;">
  <div>
    <div style="font-size:18px;font-weight:700;color:#fff;letter-spacing:.3px;">vSphere Virtualization</div>
    <div style="font-size:9pt;color:rgba(255,255,255,.85);margin-top:3px;">{h(esx_host.get("Name","qes-vh01.qes.corp"))} &middot; VMware vSphere &middot; {len(inv_vms)} VM(s) &middot; {len(inv_datastores)} Datastores</div>
  </div>
  <div style="display:flex;align-items:center;gap:14px;">
    <div style="text-align:center;background:rgba(255,255,255,.15);border-radius:8px;padding:8px 14px;">
      <div style="font-size:22px;font-weight:700;color:#fff;">{len(running_vms)}</div>
      <div style="font-size:7pt;color:rgba(255,255,255,.8);text-transform:uppercase;font-weight:600;">VMs On</div>
    </div>
    <div style="text-align:center;background:rgba(255,255,255,.15);border-radius:8px;padding:8px 14px;">
      <div style="font-size:22px;font-weight:700;color:#fff;">{total_ds_tb:.0f} TB</div>
      <div style="font-size:7pt;color:rgba(255,255,255,.8);text-transform:uppercase;font-weight:600;">Total Storage</div>
    </div>
    <span style="background:rgba(255,255,255,.22);color:#fff;font-size:10pt;font-weight:700;padding:6px 18px;border-radius:20px;border:1.5px solid rgba(255,255,255,.5);">{len(inv_vms)} VM{"s" if len(inv_vms)!=1 else ""}</span>
  </div>
</div>
<div style="background:white;border-radius:0 0 10px 10px;border:1px solid #e8e4f0;border-top:none;box-shadow:0 4px 14px rgba(0,0,0,.07);padding:20px 24px;margin-bottom:16px;">
<table style="width:100%;font-size:9pt;border-collapse:collapse;">
<tr style="background:#271e41"><th style="padding:8px 12px;color:#fff;text-align:left">VM Name</th><th style="padding:8px 12px;color:#fff">State</th><th style="padding:8px 12px;color:#fff">vCPU</th><th style="padding:8px 12px;color:#fff">RAM</th><th style="padding:8px 12px;color:#fff">IP</th></tr>
{''.join(f"""<tr><td style="padding:8px 12px;font-weight:600">{h(vm.get("Name","?"))}</td>
<td style="padding:8px 12px;text-align:center">{pill("ON","green") if vm.get("PowerState")=="POWERED_ON" else pill("OFF","gray")}</td>
<td style="padding:8px 12px;text-align:center">{vm.get("vCPU","?")}</td>
<td style="padding:8px 12px;text-align:center">{vm.get("RAMgb",0):.0f} GB</td>
<td style="padding:8px 12px;font-family:monospace;font-size:8.5pt">{h(str(vm.get("IPs","—")))}</td></tr>\n""" for vm in inv_vms)}
</table>
</div>
</div>
'''

# ── QUICKJUMP NAVBARS ─────────────────────────────────────────────────────────
rds_nav = f'''<div id="top-rds" class="hide-sbr" style="background:white;border-radius:8px;border:1px solid #e8e4f0;padding:12px 16px;margin-bottom:16px;box-shadow:0 2px 8px rgba(0,0,0,0.04);">
<div style="display:flex;flex-wrap:wrap;gap:6px;align-items:center;">
<span style="font-size:8pt;color:#6b6080;text-transform:uppercase;letter-spacing:.5px;font-weight:700;margin-right:4px;">Jump to:</span>
{nav_link("rds-alerts","&#128680; Alerts")}{nav_link("rds-overview","Overview")}{nav_link("rds-disks","Disks")}{nav_link("rds-svc-anomalies","Anomalies")}{nav_link("rds-agents-panel","Agents")}{nav_link("rds-roles","Roles")}{nav_link("rds-apps","Applications")}{nav_link("rds-hardware","Hardware")}{nav_link("rds-services","Services")}{nav_link("rds-lports","Ports")}{nav_link("rds-network","Network")}</div>
</div>
'''

dc_nav = f'''<div id="top-dc" class="hide-sbr" style="background:white;border-radius:8px;border:1px solid #e8e4f0;padding:12px 16px;margin-bottom:16px;box-shadow:0 2px 8px rgba(0,0,0,0.04);">
<div style="display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin-bottom:8px;">
<span style="font-size:8pt;color:#6b6080;text-transform:uppercase;letter-spacing:.5px;font-weight:700;margin-right:4px;">Jump to:</span>
{nav_link("dc-alerts","&#128680; Alerts")}{nav_link("dc-overview","Overview")}{nav_link("dc-disks","Disks")}{nav_link("dc-shares","File Shares")}{nav_link("dc-svc-anomalies","Anomalies")}{nav_link("dc-agents-panel","Agents")}{nav_link("dc-roles","Roles")}{nav_link("dc-apps","Applications")}{nav_link("dc-roleconfig","Role Config")}{nav_link("dc-hardware","Hardware")}{nav_link("dc-services","Services")}{nav_link("dc-lports","Ports")}{nav_link("dc-network","Network")}</div>
<div style="display:flex;flex-wrap:wrap;gap:6px;align-items:center;">
<span style="font-size:8pt;color:#6b6080;text-transform:uppercase;letter-spacing:.5px;font-weight:700;margin-right:4px;">Role Config:</span>
{nav_link_dark("dc-roleconf-ad","AD Domain Services")}{nav_link_dark("dc-roleconf-dhcp","DHCP Server")}{nav_link_dark("dc-roleconf-dns","DNS Server")}</div>
</div>
'''

virt_nav = f'''<div id="top-virt" class="hide-sbr" style="background:white;border-radius:8px;border:1px solid #e8e4f0;padding:12px 16px;margin-bottom:16px;box-shadow:0 2px 8px rgba(0,0,0,0.04);">
<div style="display:flex;flex-wrap:wrap;gap:6px;align-items:center;">
<span style="font-size:8pt;color:#6b6080;text-transform:uppercase;letter-spacing:.5px;font-weight:700;margin-right:4px;">Jump to:</span>
{nav_link("virt-host","ESX Host")}{nav_link("virt-vms","Virtual Machines")}{nav_link("virt-datastores","Datastores")}</div>
</div>
'''

# ── STUB TAB CONTENT ──────────────────────────────────────────────────────────
def stub_banner(color_hex, icon, headline, subhead, detail):
    """Generate a large-format stub banner for skipped/inaccessible VMs."""
    return f'''<div style="background:{color_hex};border-radius:10px;padding:32px 40px;margin-bottom:24px;text-align:center;border:2px solid {color_hex};">
<div style="font-size:48px;margin-bottom:12px;">{icon}</div>
<div style="font-size:28px;font-weight:800;color:white;letter-spacing:2px;text-transform:uppercase;margin-bottom:8px;">{headline}</div>
<div style="font-size:14px;font-weight:600;color:rgba(255,255,255,.85);margin-bottom:16px;">{subhead}</div>
<div style="background:rgba(0,0,0,.2);border-radius:8px;padding:12px 20px;display:inline-block;max-width:560px;">
<div style="font-size:10pt;color:rgba(255,255,255,.9);line-height:1.6;">{detail}</div>
</div>
</div>
'''

def vm_info_table(vm_name, power, ip, vcpu, ram, notes=''):
    ip_str = ip or '—'
    return f'''<div style="background:white;border-radius:8px;border:1px solid #e8e4f0;padding:20px 24px;margin-bottom:16px;">
<table style="width:100%;font-size:9.5pt;border-collapse:collapse;">
<tr><th>Attribute</th><th>Value</th></tr>
<tr><td>VM Name</td><td><b>{h(vm_name)}</b></td></tr>
<tr><td>Power State</td><td>{pill(power,"green" if power=="POWERED_ON" else "gray")}</td></tr>
<tr><td>IP Address</td><td><code>{h(ip_str)}</code></td></tr>
<tr><td>vCPU</td><td>{vcpu}</td></tr>
<tr><td>RAM</td><td>{ram} GB</td></tr>
{f"<tr><td>Notes</td><td>{h(notes)}</td></tr>" if notes else ""}
</table>
</div>
'''

# PENT-01 stub — no data collected yet
pent_content = (
    stub_banner('#7c5cbf', '&#9203;', 'PENDING', 'QES-PENT-01 &mdash; Data Not Yet Collected',
        'This VM was discovered but WinRM/WMI collection has not completed yet. '
        'Manual collection is required. Run Invoke-ServerDiscovery against 10.200.1.50 '
        'and re-generate this report to populate full detail.')
    + vm_info_table('QES-PENT-01','POWERED_ON','10.200.1.50',2,8,'Awaiting manual collection')
)

# HNY-01 stub — Linux
hny_content = (
    stub_banner('#1565c0', '&#128039;', 'LINUX VM &mdash; SKIPPED',
        'QES-HNY-01 &mdash; Linux Guest OS Detected',
        'This VM is running a Linux-based guest OS. WinRM and WMI collection do not apply to Linux. '
        'No Windows discovery data will be collected. The VM is noted here for inventory completeness only.')
    + vm_info_table('QES-HNY-01','POWERED_ON','10.200.1.242, 10.250.250.16',2,4,'Linux guest OS — no WinRM/WMI collection')
)

# ADL-01 stub — Linux
adl_content = (
    stub_banner('#1565c0', '&#128039;', 'LINUX VM &mdash; SKIPPED',
        'QES-ADL-01 &mdash; Linux Guest OS Detected',
        'This VM is running a Linux-based guest OS. WinRM and WMI collection do not apply to Linux. '
        'No Windows discovery data will be collected. The VM is noted here for inventory completeness only.')
    + vm_info_table('QES-ADL-01','POWERED_ON','10.200.1.241',2,6,'Linux guest OS — no WinRM/WMI collection')
)

# DATACENTER-DC-02 stub — powered off
dc2_content = (
    stub_banner('#455a64', '&#128274;', 'POWERED OFF &mdash; SKIPPED',
        'QES-DATACENTER-DC-02 &mdash; VM Is Offline',
        'This VM was powered off at the time of discovery. No data collection was possible. '
        'Power on the VM and re-run discovery to collect data.')
    + vm_info_table('QES-DATACENTER-DC-02','POWERED_OFF','',2,4,'Powered off at time of discovery')
)

# vCenter stub — Linux appliance
vcenter_content = (
    stub_banner('#6a1b9a', '&#128187;', 'VMWARE APPLIANCE &mdash; SKIPPED',
        'VMware vCenter Server 7 &mdash; Linux-Based Appliance',
        'The vCenter Server Appliance (VCSA) is a Linux-based appliance — WinRM collection does not apply. '
        'vCenter inventory data was collected via the vSphere REST API and is shown in the VIRTUALIZATION tab.')
    + vm_info_table('VMware vCenter Server 7','POWERED_ON','10.200.1.12',2,12,'Linux appliance — vSphere REST API used for inventory')
)

# ── RDS-01 CARD BODIES ────────────────────────────────────────────────────────
rds_alerts_html = ''.join(flag_div(*f) for f in rds_flags) or '<span class="pill pill-green">No critical alerts</span>\n'
rds_alerts_html += top_link('rds')

rds_net = rds.get('Network', {})
rds_adapt = as_list(rds_net.get('Adapters', {})) or ([rds_net['Adapters']] if isinstance(rds_net.get('Adapters'), dict) else [])
rds_ip  = rds_adapt[0].get('IPAddresses','').split(',')[0].strip() if rds_adapt else '?'
rds_gw  = rds_adapt[0].get('Gateway','') if rds_adapt else ''
rds_dns_ip = rds_adapt[0].get('DNS','') if rds_adapt else ''

rds_overview_html = f'''<div class="stat-grid">
<div class="stat-box"><div class="stat-num">{rds_hw.get("CPUCores","?")}</div><div class="stat-lbl">vCPUs</div></div>
<div class="stat-box"><div class="stat-num">{rds_ram_t:.0f}</div><div class="stat-lbl">GB RAM</div></div>
<div class="stat-box"><div class="stat-num">{len(rds_apps)}</div><div class="stat-lbl">Apps</div></div>
<div class="stat-box"><div class="stat-num">{rds_up:.0f}</div><div class="stat-lbl">Days Up</div></div>
</div>
<table>
<tr><th>Attribute</th><th>Value</th></tr>
<tr><td>Hostname</td><td>{h(rds_sys.get("Hostname",""))}</td></tr>
<tr><td>OS</td><td>{h(rds_sys.get("OSName",""))} (Build {h(str(rds_sys.get("OSBuild","")))})</td></tr>
<tr><td>Domain</td><td>{h(rds_sys.get("Domain",""))}</td></tr>
<tr><td>Last Boot</td><td>{h(rds_sys.get("LastBoot",""))}</td></tr>
<tr><td>Timezone</td><td>{h(rds_sys.get("Timezone",""))}</td></tr>
<tr><td>Run As</td><td><code>{h(rds_sys.get("RunAsUser",""))}</code></td></tr>
<tr><td>Platform</td><td>VMware VM on qes-vh01.qes.corp</td></tr>
<tr><td>CPU</td><td>{h(rds_hw.get("CPUName",""))} &middot; {rds_hw.get("CPUCores","?")} vCPUs</td></tr>
<tr><td>RAM</td><td>{rds_ram_t:.2f} GB total / {rds_ram_f:.2f} GB available</td></tr>
<tr><td>IP Address</td><td>{h(rds_ip)}</td></tr>
<tr><td>Gateway</td><td>{h(rds_gw)}</td></tr>
<tr><td>DNS</td><td>{h(rds_dns_ip)}</td></tr>
<tr><td>OS Install Date</td><td>{h(rds_sys.get("OSInstallDate",""))}</td></tr>
<tr><td>EOL Date</td><td>{h(rds_sys.get("OSEOLDate",""))} {pill(rds_sys.get("OSEOLStatus",""),"green")}</td></tr>
</table>
''' + top_link('rds')

rds_hw_html = f'''<table>
<tr><th>Component</th><th>Value</th></tr>
<tr><td>Platform</td><td>{pill(rds_hw.get("VMPlatform","VMware"),"purple")} hosted on qes-vh01.qes.corp</td></tr>
<tr><td>Model</td><td>{h(rds_hw.get("Model",""))}</td></tr>
<tr><td>CPU</td><td>{h(rds_hw.get("CPUName",""))} &middot; {rds_hw.get("CPUCores","?")} vCPUs</td></tr>
<tr><td>RAM</td><td>{rds_ram_t:.2f} GB total / {rds_ram_f:.2f} GB available</td></tr>
<tr><td>BIOS</td><td>{h(rds_hw.get("BIOSVersion",""))} ({h(rds_hw.get("BIOSDate",""))})</td></tr>
<tr><td>Serial Number</td><td><code>{h(rds_hw.get("SerialNumber",""))}</code></td></tr>
</table>
''' + top_link('rds')

rds_apps_html = apps_table(rds_apps, 'rds')

rds_roles_html  = sub('Installed Roles')
rds_roles_html += '<div class="role-grid">' + ''.join(f'<span class="role-badge">{h(r.get("DisplayName",r.get("Name","")))}</span>' for r in rds_roles_list) + '</div>\n'
rds_roles_html += sub('Installed Features')
rds_roles_html += '<div class="role-grid">' + ''.join(f'<span class="role-badge" style="background:#f0f4ff;color:#5b1fa4;border-color:#a5b4fc;">{h(f.get("DisplayName",f.get("Name","")))}</span>' for f in rds_features_list) + '</div>\n'
rds_roles_html += top_link('rds')

rds_disk_html  = sub('Disk Volumes')
rds_disk_html += '<table><tr><th>Drive</th><th>Label</th><th>FS</th><th>Total GB</th><th>Free GB</th><th>Used %</th><th>Bar</th></tr>\n'
for d in rds_disks:
    pct = d.get('UsedPct', 0)
    rc  = 'style="background:#fff0f0"' if pct>=85 else ('style="background:#fff8e1"' if pct>=70 else '')
    rds_disk_html += (f'<tr {rc}><td><b>{h(d.get("Drive","?"))}</b></td><td>{h(d.get("Label","") or "")}</td>'
                      f'<td>{h(d.get("Filesystem","NTFS"))}</td><td>{d.get("TotalGB",0):.2f}</td>'
                      f'<td>{d.get("FreeGB",0):.2f}</td><td>{pct}%</td>'
                      f'<td style="min-width:100px">{disk_bar(pct)}</td></tr>\n')
rds_disk_html += '</table>\n' + top_link('rds')

rds_net_html  = sub('Network Adapters')
for ad in rds_adapt:
    rds_net_html += f'''<table style="margin-bottom:16px;">
<tr><th>Attribute</th><th>Value</th></tr>
<tr><td>Description</td><td>{h(ad.get("Description",""))}</td></tr>
<tr><td>IP Addresses</td><td>{h(ad.get("IPAddresses",""))}</td></tr>
<tr><td>Gateway</td><td>{h(ad.get("Gateway",""))}</td></tr>
<tr><td>DNS</td><td>{h(ad.get("DNS",""))}</td></tr>
<tr><td>MAC</td><td><code>{h(ad.get("MAC",""))}</code></td></tr>
<tr><td>DHCP</td><td>{pill("Enabled","yellow") if ad.get("DHCPEnabled") else pill("Static","green")}</td></tr>
</table>
'''
rds_net_html += top_link('rds')

rds_ports = as_list(rds_net.get('ListeningPorts', []))
rds_lports_html  = '<table><tr><th>Port</th><th>Protocol</th><th>Process</th><th>State</th><th>PID</th></tr>\n'
for p in sorted(rds_ports, key=lambda x: int(str(x.get('Port', x.get('LocalPort', 9999))))):
    port  = p.get('Port', p.get('LocalPort', '?'))
    proto = p.get('Proto', p.get('Protocol', ''))
    proc  = p.get('Process', p.get('ProcessName', ''))
    state = p.get('State', '')
    rds_lports_html += (f'<tr><td><b>{port}</b></td><td>{h(proto)}</td>'
                        f'<td><code>{h(proc)}</code></td><td>{h(state)}</td><td>{p.get("PID","")}</td></tr>\n')
rds_lports_html += '</table>\n' + top_link('rds')

rds_services = as_list(rds.get('Services', []))
rds_svc_html  = '<table><tr><th>Name</th><th>Display Name</th><th>Status</th><th>Start Type</th></tr>\n'
for svc in sorted(rds_services, key=lambda x: x.get('Name','')):
    status = svc.get('State', svc.get('Status',''))
    sc = 'green' if status=='Running' else ('yellow' if status in ('Stopped','Stop') else 'gray')
    rds_svc_html += (f'<tr><td><code>{h(svc.get("Name",""))}</code></td><td>{h(svc.get("DisplayName",""))}</td>'
                     f'<td>{pill(status,sc)}</td><td>{h(svc.get("StartMode",svc.get("StartType","")))}</td></tr>\n')
rds_svc_html += '</table>\n' + top_link('rds')

rds_anom_svcs = [s for s in rds_services
                 if s.get('StartMode','') in ('Auto','Automatic') and s.get('State','') in ('Stopped',)
                 and s.get('Name','') not in ('RemoteRegistry','AppMgmt')]
rds_anom_html = ''
if rds_anom_svcs:
    rds_anom_html = '<div class="flag-warning"><div class="flag-label">Auto-Start Services That Are Stopped</div></div>\n'
    rds_anom_html += '<table><tr><th>Name</th><th>Display Name</th></tr>\n'
    for s in rds_anom_svcs:
        rds_anom_html += f'<tr><td><code>{h(s.get("Name",""))}</code></td><td>{h(s.get("DisplayName",""))}</td></tr>\n'
    rds_anom_html += '</table>\n'
else:
    rds_anom_html = '<span class="pill pill-green">No service anomalies detected</span>\n'
rds_anom_html += top_link('rds')

# ── DATA-DC CARD BODIES ───────────────────────────────────────────────────────
dc_alerts_html = ''.join(flag_div(*f) for f in dc_flags) or '<span class="pill pill-green">No critical alerts</span>\n'
dc_alerts_html += top_link('dc')

dc_net = dc.get('Network', {})
dc_adapt = as_list(dc_net.get('Adapters', {})) or ([dc_net['Adapters']] if isinstance(dc_net.get('Adapters'), dict) else [])
dc_ip  = dc_adapt[0].get('IPAddresses','').split(',')[0].strip() if dc_adapt else '?'
dc_gw  = dc_adapt[0].get('Gateway','') if dc_adapt else ''
dc_dns_ip = dc_adapt[0].get('DNS','') if dc_adapt else ''

dc_overview_html = f'''<div class="stat-grid">
<div class="stat-box"><div class="stat-num">{dc_hw.get("CPUCores","?")}</div><div class="stat-lbl">vCPUs</div></div>
<div class="stat-box"><div class="stat-num">{dc_ram_t:.0f}</div><div class="stat-lbl">GB RAM</div></div>
<div class="stat-box"><div class="stat-num">{dc_ad.get("UserCount","?")}</div><div class="stat-lbl">AD Users</div></div>
<div class="stat-box"><div class="stat-num">{dc_up:.0f}</div><div class="stat-lbl">Days Up</div></div>
</div>
<table>
<tr><th>Attribute</th><th>Value</th></tr>
<tr><td>Hostname</td><td>{h(dc_sys.get("Hostname",""))}</td></tr>
<tr><td>OS</td><td>{h(dc_sys.get("OSName",""))} (Build {h(str(dc_sys.get("OSBuild","")))})</td></tr>
<tr><td>Domain</td><td>{h(dc_sys.get("Domain",""))}</td></tr>
<tr><td>Last Boot</td><td>{h(dc_sys.get("LastBoot",""))}</td></tr>
<tr><td>Timezone</td><td>{h(dc_sys.get("Timezone",""))}</td></tr>
<tr><td>Run As</td><td><code>{h(dc_sys.get("RunAsUser",""))}</code></td></tr>
<tr><td>Platform</td><td>VMware VM on qes-vh01.qes.corp</td></tr>
<tr><td>CPU</td><td>{h(dc_hw.get("CPUName",""))} &middot; {dc_hw.get("CPUCores","?")} vCPUs</td></tr>
<tr><td>RAM</td><td>{dc_ram_t:.2f} GB total / {dc_ram_f:.2f} GB available {pill("CRITICAL — only "+str(dc_ram_f)+" GB free","red") if dc_ram_pct_used>=85 else ""}</td></tr>
<tr><td>IP Address</td><td>{h(dc_ip)}</td></tr>
<tr><td>DNS</td><td>{h(dc_dns_ip)}</td></tr>
<tr><td>OS Install Date</td><td>{h(dc_sys.get("OSInstallDate",""))}</td></tr>
<tr><td>EOL Date</td><td>{h(dc_sys.get("OSEOLDate",""))} {pill(dc_sys.get("OSEOLStatus",""),"green")}</td></tr>
</table>
''' + top_link('dc')

dc_hw_html = f'''<table>
<tr><th>Component</th><th>Value</th></tr>
<tr><td>Platform</td><td>{pill(dc_hw.get("VMPlatform","VMware"),"purple")} hosted on qes-vh01.qes.corp</td></tr>
<tr><td>Model</td><td>{h(dc_hw.get("Model",""))}</td></tr>
<tr><td>CPU</td><td>{h(dc_hw.get("CPUName",""))} &middot; {dc_hw.get("CPUCores","?")} vCPUs</td></tr>
<tr><td>RAM</td><td>{dc_ram_t:.2f} GB total / {dc_ram_f:.2f} GB available</td></tr>
<tr><td>BIOS</td><td>{h(dc_hw.get("BIOSVersion",""))} ({h(dc_hw.get("BIOSDate",""))})</td></tr>
<tr><td>Serial Number</td><td><code>{h(dc_hw.get("SerialNumber",""))}</code></td></tr>
</table>
''' + top_link('dc')

dc_apps_html = apps_table(dc_apps, 'dc')

dc_roles_html  = sub('Installed Roles')
dc_roles_html += '<div class="role-grid">' + ''.join(f'<span class="role-badge">{h(r.get("DisplayName",r.get("Name","")))}</span>' for r in dc_roles_list) + '</div>\n'
dc_roles_html += sub('Installed Features')
dc_roles_html += '<div class="role-grid">' + ''.join(f'<span class="role-badge" style="background:#f0f4ff;color:#5b1fa4;border-color:#a5b4fc;">{h(f.get("DisplayName",f.get("Name","")))}</span>' for f in dc_features_list) + '</div>\n'
dc_roles_html += top_link('dc')

## DC Role Config
dhcp = dc.get('DHCP', {})
dns  = dc.get('DNS', {})

dc_rc_html  = f'<div id="dc-roleconf-ad"></div>' + sub('Active Directory')
dc_rc_html += f'''<table>
<tr><th>Attribute</th><th>Value</th></tr>
<tr><td>Forest</td><td>{h(dc_ad.get("ForestName",""))}</td></tr>
<tr><td>PDC Emulator</td><td>{h(dc_ad.get("PDCEmulator",""))}</td></tr>
<tr><td>RID Master</td><td>{h(dc_ad.get("RIDMaster",""))}</td></tr>
<tr><td>Schema Master</td><td>{h(dc_ad.get("SchemaMaster",""))}</td></tr>
<tr><td>Domain FL</td><td>{pill(fl_lbl,"red") if "2008" in fl_lbl else (pill(fl_lbl,"yellow") if fl_lbl < "2019" else pill(fl_lbl,"green"))}</td></tr>
<tr><td>OU Count</td><td>{dc_ad.get("OUCount","?")}</td></tr>
<tr><td>DC Count</td><td>{dc_ad.get("DCCount","?")}</td></tr>
<tr><td>User Count</td><td>{dc_ad.get("UserCount","?")}</td></tr>
<tr><td>Computer Count</td><td>{dc_ad.get("ComputerCount","?")}</td></tr>
<tr><td>Stale Users</td><td>{len(stale_users)} (90+ days inactive)</td></tr>
<tr><td>Stale Computers</td><td>{len(stale_comps)}</td></tr>
<tr><td>Entra Sync</td><td>{pill("Configured — Entra Connect","green") if dc_entra else pill("Not detected","gray")}</td></tr>
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

dc_rc_html += f'<div id="dc-roleconf-dhcp"></div>' + sub('DHCP Server')
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

dc_rc_html += f'<div id="dc-roleconf-dns"></div>' + sub('DNS Server')
dns_zones = as_list(dns.get('Zones', []))
dns_fwd = dns.get('Forwarders', '')
if isinstance(dns_fwd, list):
    fwd_str = ', '.join(item.get('IPAddressToString', str(item)) for item in dns_fwd if isinstance(item, dict)) or '—'
else:
    fwd_str = str(dns_fwd) if dns_fwd else '—'
dc_rc_html += f'<div class="meta-line">Forwarders: <code>{h(fwd_str)}</code></div>\n'
if dns_zones:
    dc_rc_html += '<table><tr><th>Zone</th><th>Type</th><th>DS Integrated</th><th>Reverse Lookup</th></tr>\n'
    for z in dns_zones:
        ds_int = pill("Yes","green") if z.get("IsDsIntegrated") else pill("No","gray")
        rev    = pill("Yes","purple") if z.get("IsReverseLookupZone") else ''
        dc_rc_html += (f'<tr><td>{h(z.get("ZoneName", z.get("Name","?")))}</td>'
                       f'<td>{h(z.get("ZoneType",""))}</td>'
                       f'<td>{ds_int}</td><td>{rev}</td></tr>\n')
    dc_rc_html += '</table>\n'
dc_rc_html += top_link('dc')

dc_disk_html  = sub('Disk Volumes')
dc_disk_html += '<table><tr><th>Drive</th><th>Label</th><th>FS</th><th>Total GB</th><th>Free GB</th><th>Used %</th><th>Bar</th></tr>\n'
for d in dc_disks:
    pct = d.get('UsedPct', 0)
    rc  = 'style="background:#fff0f0"' if pct>=85 else ('style="background:#fff8e1"' if pct>=70 else '')
    dc_disk_html += (f'<tr {rc}><td><b>{h(d.get("Drive","?"))}</b></td><td>{h(d.get("Label","") or "")}</td>'
                     f'<td>{h(d.get("Filesystem","NTFS"))}</td><td>{d.get("TotalGB",0):.2f}</td>'
                     f'<td>{d.get("FreeGB",0):.2f}</td><td>{pct}%</td>'
                     f'<td style="min-width:100px">{disk_bar(pct)}</td></tr>\n')
dc_disk_html += '</table>\n' + top_link('dc')

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
dc_net_html += top_link('dc')

dc_ports = as_list(dc_net.get('ListeningPorts', []))
dc_lports_html  = '<table><tr><th>Port</th><th>Protocol</th><th>Process</th><th>State</th><th>PID</th></tr>\n'
for p in sorted(dc_ports, key=lambda x: int(str(x.get('Port', x.get('LocalPort', 9999))))):
    port  = p.get('Port', p.get('LocalPort', '?'))
    proto = p.get('Proto', p.get('Protocol', ''))
    proc  = p.get('Process', p.get('ProcessName', ''))
    state = p.get('State', '')
    dc_lports_html += (f'<tr><td><b>{port}</b></td><td>{h(proto)}</td>'
                       f'<td><code>{h(proc)}</code></td><td>{h(state)}</td><td>{p.get("PID","")}</td></tr>\n')
dc_lports_html += '</table>\n' + top_link('dc')

dc_services = as_list(dc.get('Services', []))
dc_svc_html  = '<table><tr><th>Name</th><th>Display Name</th><th>Status</th><th>Start Type</th></tr>\n'
for svc in sorted(dc_services, key=lambda x: x.get('Name','')):
    status = svc.get('State', svc.get('Status',''))
    sc = 'green' if status=='Running' else ('yellow' if status in ('Stopped','Stop') else 'gray')
    dc_svc_html += (f'<tr><td><code>{h(svc.get("Name",""))}</code></td><td>{h(svc.get("DisplayName",""))}</td>'
                    f'<td>{pill(status,sc)}</td><td>{h(svc.get("StartMode",svc.get("StartType","")))}</td></tr>\n')
dc_svc_html += '</table>\n' + top_link('dc')

dc_anom_svcs = [s for s in dc_services
                if s.get('StartMode','') in ('Auto','Automatic') and s.get('State','') in ('Stopped',)
                and s.get('Name','') not in ('RemoteRegistry','AppMgmt')]
dc_anom_html = ''
if dc_anom_svcs:
    dc_anom_html = '<div class="flag-warning"><div class="flag-label">Auto-Start Services That Are Stopped</div></div>\n'
    dc_anom_html += '<table><tr><th>Name</th><th>Display Name</th></tr>\n'
    for s in dc_anom_svcs:
        dc_anom_html += f'<tr><td><code>{h(s.get("Name",""))}</code></td><td>{h(s.get("DisplayName",""))}</td></tr>\n'
    dc_anom_html += '</table>\n'
else:
    dc_anom_html = '<span class="pill pill-green">No service anomalies detected</span>\n'
dc_anom_html += top_link('dc')

dc_shares_html = ''
if dc_shares_real:
    dc_shares_html += '<table><tr><th>Share</th><th>Path</th><th>Sessions</th></tr>\n'
    for s in dc_shares_real:
        dc_shares_html += (f'<tr><td><b>{h(s.get("Name",""))}</b></td>'
                           f'<td style="font-family:monospace;font-size:8.5pt">{h(s.get("Path",""))}</td>'
                           f'<td>{s.get("OpenSessions",0)}</td></tr>\n')
    dc_shares_html += '</table>\n'
else:
    dc_shares_html = '<span class="pill pill-gray">No non-admin shares detected</span>\n'
dc_shares_html += top_link('dc')

# ── VIRT CARD BODIES ──────────────────────────────────────────────────────────
virt_host_html = f'''<table>
<tr><th>Attribute</th><th>Value</th></tr>
<tr><td>ESX Host</td><td><b>{h(esx_host.get("Name","qes-vh01.qes.corp"))}</b></td></tr>
<tr><td>State</td><td>{pill(esx_host.get("State","CONNECTED"),"green")}</td></tr>
<tr><td>Power</td><td>{pill(esx_host.get("PowerState","POWERED_ON"),"green")}</td></tr>
<tr><td>vCenter</td><td>VMware vCenter Server 7 ({h(inv.get("Server","10.200.1.12"))})</td></tr>
<tr><td>API Version</td><td>{h(inv.get("APIVersion",""))}</td></tr>
<tr><td>VMs Hosted</td><td>{len(inv_vms)} total ({len(running_vms)} powered on)</td></tr>
<tr><td>Datastores</td><td>{len(inv_datastores)}</td></tr>
<tr><td>Collected</td><td>{h(inv.get("CollectedAt",""))}</td></tr>
</table>
''' + top_link('virt')

virt_vms_html = '<table><tr><th>VM Name</th><th>State</th><th>vCPU</th><th>RAM (GB)</th><th>IP Address(es)</th><th>Datastore</th><th>NIC Type</th></tr>\n'
for vm in inv_vms:
    state = vm.get('PowerState','?')
    sc = 'green' if state=='POWERED_ON' else 'gray'
    nics = vm.get('NICs', [])
    nic_type = ', '.join(n.get('Type','') for n in nics) if nics else '—'
    ds_name = _vm_ds_map.get(vm.get('Name',''), '—')
    virt_vms_html += (f'<tr><td><b>{h(vm.get("Name","?"))}</b></td>'
                      f'<td>{pill(state,sc)}</td>'
                      f'<td>{vm.get("vCPU","?")}</td>'
                      f'<td>{vm.get("RAMgb",0):.0f}</td>'
                      f'<td style="font-family:monospace;font-size:8.5pt">{h(str(vm.get("IPs","—")) or "—")}</td>'
                      f'<td style="font-size:8.5pt">{h(str(ds_name))}</td>'
                      f'<td>{h(nic_type)}</td></tr>\n')
virt_vms_html += '</table>\n'

# VM disk detail
for vm in inv_vms:
    vdisks = vm.get('Disks', [])
    if vdisks:
        virt_vms_html += sub(f'Disks: {h(vm.get("Name","?"))}')
        virt_vms_html += '<table><tr><th>Label</th><th>Capacity (GB)</th><th>Backing Type</th></tr>\n'
        for vd in vdisks:
            virt_vms_html += (f'<tr><td>{h(vd.get("Label","?"))}</td>'
                              f'<td>{vd.get("CapacityGB",0):.0f}</td>'
                              f'<td>{pill(vd.get("BackingType","?"),"purple")}</td></tr>\n')
        virt_vms_html += '</table>\n'
virt_vms_html += top_link('virt')

virt_ds_html = '<table><tr><th>Datastore</th><th>Type</th><th>Capacity (GB)</th><th>Free (GB)</th><th>Used %</th><th>Bar</th></tr>\n'
for ds in inv_datastores:
    cap  = ds.get('CapacityGB', 0)
    free = ds.get('FreeGB', 0)
    used = cap - free
    pct  = int(used/cap*100) if cap else 0
    rc   = 'style="background:#fff0f0"' if pct>=85 else ('style="background:#fff8e1"' if pct>=70 else '')
    virt_ds_html += (f'<tr {rc}><td><b>{h(ds.get("Name","?"))}</b></td>'
                     f'<td>{h(ds.get("Type","VMFS"))}</td>'
                     f'<td>{cap:.1f}</td><td>{free:.1f}</td><td>{pct}%</td>'
                     f'<td style="min-width:120px">{disk_bar(pct)}</td></tr>\n')
virt_ds_html += f'</table>\n<div class="meta-line" style="margin-top:8px;">Total raw capacity: {total_ds_tb:.1f} TB &middot; Used: {used_ds_tb:.1f} TB ({int(used_ds_tb/total_ds_tb*100) if total_ds_tb else 0}%)</div>\n'
virt_ds_html += top_link('virt')

# ── CSS + JS ──────────────────────────────────────────────────────────────────
CSS = '''
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Segoe UI', Arial, sans-serif; background: #f5f4f8; color: #271e41; font-size: 10pt; }
.wrap { max-width: 1040px; margin: 0 auto; padding: 20px; }
.tab-nav { display: flex; gap: 4px; margin-bottom: -1px; flex-wrap: wrap; }
.tab-btn { padding: 8px 18px; background: #ddd9ee; border: 1px solid #c0b8d8; border-bottom: none; border-radius: 6px 6px 0 0; cursor: pointer; font-size: 9.5pt; color: #271e41; font-weight: 600; }
.tab-btn.active { background: white; border-bottom: 1px solid white; color: #5b1fa4; }
.tab-btn.has-critical { border-top: 3px solid #d63638; }
.tab-btn.has-warning  { border-top: 3px solid #f5a623; }
.tab-btn.stub-linux   { border-top: 3px solid #1565c0; opacity: .8; }
.tab-btn.stub-off     { border-top: 3px solid #607d8b; opacity: .7; }
.tab-btn.stub-pending { border-top: 3px solid #7c5cbf; }
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
.adv-only{display:none !important;}
body.view-adv .adv-only{display:block !important;}
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
  basic: "Full detail \u2014 SE & CSM view",
  adv:   "Full technical detail \u2014 SE view",
  sbr:   "Executive health dashboard \u2014 client & leadership"
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

# ── ASSEMBLE HTML ─────────────────────────────────────────────────────────────
out = f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>QES Server Discovery Report &mdash; {DATE}</title>
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
      <div style="color:#a89bc8;font-size:9pt;margin-top:2px">QES Environment (Quantitative Economic Solutions) &mdash; {DATE}</div>
    </div>
  </div>
  <div style="text-align:right">
    <div style="color:#c4b5fd;font-size:8.5pt">Collected: {DATE}</div>
    <div style="color:#a89bc8;font-size:8pt;margin-top:2px">Magna5 Solutions Engineering</div>
  </div>
</div>

<div class="tab-nav">
<button class="tab-btn active{tab_cls(rds_crit,rds_warn)}" data-tab="tab-rds" onclick="showTab('tab-rds')">QES-RDS-01 &middot; Server 2019 (RDS)</button>
<button class="tab-btn{tab_cls(dc_crit,dc_warn)}" data-tab="tab-dc" onclick="showTab('tab-dc')">QES-DATA-DC &middot; Server 2019 (DC)</button>
<button class="tab-btn" data-tab="tab-virt" onclick="showTab('tab-virt')">VIRTUALIZATION &middot; vSphere</button>
<button class="tab-btn stub-pending" data-tab="tab-pent" onclick="showTab('tab-pent')">QES-PENT-01 &middot; Pending</button>
<button class="tab-btn stub-linux" data-tab="tab-hny" onclick="showTab('tab-hny')">QES-HNY-01 &middot; Linux</button>
<button class="tab-btn stub-linux" data-tab="tab-adl" onclick="showTab('tab-adl')">QES-ADL-01 &middot; Linux</button>
<button class="tab-btn stub-off" data-tab="tab-dc2" onclick="showTab('tab-dc2')">QES-DATACENTER-DC-02 &middot; Off</button>
<button class="tab-btn stub-off" data-tab="tab-vcenter" onclick="showTab('tab-vcenter')">VMware vCenter &middot; Appliance</button>
</div>

<div id="tab-rds" class="tab-content active">
{rds_sbr}
{rds_nav}
{card("rds-alerts",       "&#128680; Alerts",                             rds_alerts_html,                              extra_class='hide-sbr')}
{card("rds-overview",     "System Overview",                               rds_overview_html,                            extra_class='hide-sbr')}
{card("rds-disks",        "Disk Storage",                                  rds_disk_html,                                extra_class='hide-sbr')}
{card("rds-svc-anomalies","Service Anomalies",                             rds_anom_html,                                extra_class='hide-sbr')}
{card("rds-agents-panel", "&#128737; Installed Agents",                    rds_agents or '<span style="color:#9b8fb0;font-style:italic">No agents detected</span>', extra_class='adv-only')}
{card("rds-roles",        "Roles &amp; Features",                          rds_roles_html,                               extra_class='adv-only', collapsed=True)}
{card("rds-apps",         "Installed Applications ({} apps)".format(len(rds_apps)), rds_apps_html,            extra_class='adv-only', collapsed=True)}
{card("rds-hardware",     "Hardware",                                      rds_hw_html,                                  extra_class='adv-only')}
{card("rds-services",     "Services",                                      rds_svc_html,                                 extra_class='adv-only', collapsed=True)}
{card("rds-lports",       "Listening Ports ({} open)".format(len(rds_ports)), rds_lports_html,                 extra_class='adv-only', collapsed=True)}
{card("rds-network",      "Network",                                       rds_net_html,                                 extra_class='adv-only')}
</div>

<div id="tab-dc" class="tab-content">
{dc_sbr}
{dc_nav}
{card("dc-alerts",        "&#128680; Alerts",                             dc_alerts_html,                               extra_class='hide-sbr')}
{card("dc-overview",      "System Overview",                               dc_overview_html,                             extra_class='hide-sbr')}
{card("dc-disks",         "Disk Storage",                                  dc_disk_html,                                 extra_class='hide-sbr')}
{card("dc-shares",        "File Shares",                                   dc_shares_html,                               extra_class='hide-sbr')}
{card("dc-svc-anomalies", "Service Anomalies",                             dc_anom_html,                                 extra_class='hide-sbr')}
{card("dc-agents-panel",  "&#128737; Installed Agents",                    dc_agents or '<span style="color:#9b8fb0;font-style:italic">No agents detected</span>', extra_class='adv-only')}
{card("dc-roles",         "Roles &amp; Features",                          dc_roles_html,                                extra_class='adv-only', collapsed=True)}
{card("dc-apps",          "Installed Applications ({} apps)".format(len(dc_apps)), dc_apps_html,             extra_class='adv-only', collapsed=True)}
{card("dc-roleconfig",    "Role Configuration",                            dc_rc_html,                                   extra_class='adv-only', collapsed=True)}
{card("dc-hardware",      "Hardware",                                      dc_hw_html,                                   extra_class='adv-only')}
{card("dc-services",      "Services",                                      dc_svc_html,                                  extra_class='adv-only', collapsed=True)}
{card("dc-lports",        "Listening Ports ({} open)".format(len(dc_ports)), dc_lports_html,                  extra_class='adv-only', collapsed=True)}
{card("dc-network",       "Network",                                       dc_net_html,                                  extra_class='adv-only')}
</div>

<div id="tab-virt" class="tab-content">
{virt_sbr}
{virt_nav}
{card("virt-host", "ESX Host", virt_host_html)}
{card("virt-vms", "Virtual Machines ({} VMs)".format(len(inv_vms)), virt_vms_html)}
{card("virt-datastores", "Datastores ({} volumes)".format(len(inv_datastores)), virt_ds_html)}
</div>

<div id="tab-pent" class="tab-content">
{pent_content}
</div>

<div id="tab-hny" class="tab-content">
{hny_content}
</div>

<div id="tab-adl" class="tab-content">
{adl_content}
</div>

<div id="tab-dc2" class="tab-content">
{dc2_content}
</div>

<div id="tab-vcenter" class="tab-content">
{vcenter_content}
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
print(f"RDS-01 flags: {rds_crit} critical / {rds_warn} warning")
print(f"DATA-DC flags: {dc_crit} critical / {dc_warn} warning")
