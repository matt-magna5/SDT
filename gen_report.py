"""SDT HTML report generator (v4.2 design).

Top nav: Overview / Servers / Hypervisor / Active Directory / SQL /
         Private Cloud Sizing / Commvault Sizing / EOL Risks

Servers view uses a left sidebar nav (searchable) with the full per-server
detail rendered in the right pane.

Usage:
    python gen_report.py <manifest.json>
    python gen_report.py <session_dir>           # also accepted

The output HTML is written to the session directory as
`<Client>-DiscoveryReport-<date>.html` and the path is printed to stdout.
"""
import sys, json
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
from _report_lib import load_session, detect_security, env_rollup, filtered_flags, h

def render(session):
    servers = session['servers']
    env = env_rollup(servers)
    client = session['client']
    vc = session.get('vcenter') or {}

    # ----- vCenter / hypervisor stats -----
    vc_html = ''
    if vc:
        cluster = vc.get('Cluster') or {}
        hosts = vc.get('ESXHosts') or []
        if not isinstance(hosts, list): hosts = []
        vms = vc.get('VMs') or []
        if not isinstance(vms, list): vms = []
        ds_list = vc.get('Datastores') or []
        if not isinstance(ds_list, list): ds_list = []
        vms_on = [v for v in vms if isinstance(v, dict) and str(v.get('PowerState','')).lower() in ('poweredon','on')]
        total_vcpu_on = sum(int(v.get('vCPUs') or 0) for v in vms_on)
        total_ram_on  = sum(float(v.get('RAMgb') or 0) for v in vms_on)
        total_disk_consumed = sum(float(v.get('DiskConsumedGB') or 0) for v in vms_on)
        total_disk_prov = sum(float(v.get('DiskCapGB') or 0) for v in vms_on)
        host_cores = sum(int(h0.get('Cores') or 0) for h0 in hosts if isinstance(h0,dict))
        host_ram   = sum(float(h0.get('RAMgb') or 0) for h0 in hosts if isinstance(h0,dict))
        host_rows = ''.join(
            f'<tr><td>{h(ho.get("Name",""))}</td><td>{h(ho.get("Vendor",""))} {h(ho.get("Model",""))}</td>'
            f'<td>{h(ho.get("ServiceTag",""))}</td><td>{ho.get("Cores","?")}c</td>'
            f'<td>{ho.get("RAMgb","?")} GB</td><td>{ho.get("CPUUsagePct","?")}%</td>'
            f'<td>{ho.get("MemUsagePct","?")}%</td></tr>'
            for ho in hosts if isinstance(ho, dict)
        )
        ds_rows = ''.join(
            f'<tr><td>{h(ds.get("Name",""))}</td><td>{h(ds.get("Type",""))}</td>'
            f'<td>{ds.get("CapacityGB","?"):,} GB</td><td>{ds.get("FreeGB","?"):,} GB</td>'
            f'<td><div class="bar"><div style="width:{min(100,float(ds.get("UsedPct") or 0))}%;'
            f'background:{"#ef4444" if (ds.get("UsedPct") or 0)>=85 else "#22c55e"}"></div></div> {ds.get("UsedPct","?")}%</td></tr>'
            for ds in ds_list if isinstance(ds, dict)
        )
        cluster_cap_gib = (cluster.get('CapacityMiB') or 0) / 1024
        cluster_used_gib = (cluster.get('ConsumedMiB') or 0) / 1024
        cluster_pct = (cluster_used_gib / cluster_cap_gib * 100) if cluster_cap_gib else 0
        vc_html = f'''
<h3>vCenter / Hypervisor <a class="jump" onclick="setViewByName('servers')">View VMs &rsaquo;</a></h3>
<div class="hgrid">
  <div class="hcell">
    <h4>vCenter</h4>
    <div class="stat-row"><b>Server:</b> <code>{h(vc.get("Server",""))}</code></div>
    <div class="stat-row"><b>Version:</b> {h(vc.get("Version",""))}</div>
    <div class="stat-row"><b>Cluster:</b> {h(cluster.get("Name",""))} ({cluster.get("NumHosts","?")} hosts)</div>
    <div class="stat-row"><b>Perf window:</b> {vc.get("DurationDays","?")} days</div>
  </div>
  <div class="hcell">
    <h4>Cluster Capacity</h4>
    <div class="stat"><span class="snum">{cluster_cap_gib/1024:.1f}</span><span class="slbl">TiB total</span></div>
    <div class="stat"><span class="snum">{cluster_used_gib/1024:.1f}</span><span class="slbl">TiB used</span></div>
    <div class="stat"><span class="snum">{cluster_pct:.0f}%</span><span class="slbl">utilization</span></div>
  </div>
  <div class="hcell">
    <h4>VMs (from vCenter)</h4>
    <div class="stat"><span class="snum">{len(vms_on)}</span><span class="slbl">powered on (of {len(vms)})</span></div>
    <div class="stat"><span class="snum">{total_vcpu_on}</span><span class="slbl">vCPU allocated</span></div>
    <div class="stat"><span class="snum">{total_ram_on:.0f}</span><span class="slbl">GB RAM allocated</span></div>
    <div class="stat"><span class="snum">{total_disk_consumed:,.0f} / {total_disk_prov:,.0f}</span><span class="slbl">GB disk used / provisioned</span></div>
  </div>
</div>
<h4>ESX Hosts ({len(hosts)}{' of ' + str(cluster.get('NumHosts')) + ' in cluster' if cluster.get('NumHosts') and len(hosts) < int(cluster.get('NumHosts') or 0) else ''})</h4>
<table class="t"><thead><tr><th>Host</th><th>Model</th><th>Service Tag</th><th>CPU</th><th>RAM</th><th>CPU Use</th><th>Mem Use</th></tr></thead><tbody>{host_rows or "<tr><td colspan=7 class=dim>No host detail.</td></tr>"}</tbody></table>
<h4>Datastores ({len(ds_list)})</h4>
<table class="t"><thead><tr><th>Name</th><th>Type</th><th>Capacity</th><th>Free</th><th>Used</th></tr></thead><tbody>{ds_rows}</tbody></table>
'''

    # ----- WinRM sizing rollup (guest-side data) -----
    total_cores_guest = sum(int(s.get('cores') or 0) for s in servers)
    total_ram_guest = sum(float(s.get('ram_gb') or 0) for s in servers)
    total_disk_used = 0.0
    total_disk_total = 0.0
    for s in servers:
        for d in s.get('disks',[]):
            if isinstance(d, dict):
                total_disk_used += float(d.get('TotalGB') or 0) - float(d.get('FreeGB') or 0)
                total_disk_total += float(d.get('TotalGB') or 0)

    # ----- Security posture % + Detected tools breakdown -----
    have_edr = have_rmm = have_backup = have_remote = 0
    edr_products = {}   # product -> [server names]
    rmm_products = {}
    backup_products = {}
    remote_products = {}
    for s in servers:
        sec = detect_security(s)
        if sec.get('edr'): have_edr += 1
        if sec.get('rmm'): have_rmm += 1
        if sec.get('backup'): have_backup += 1
        if sec.get('remote'): have_remote += 1
        for p in sec.get('edr', []):    edr_products.setdefault(p, []).append(s['name'])
        for p in sec.get('rmm', []):    rmm_products.setdefault(p, []).append(s['name'])
        for p in sec.get('backup', []): backup_products.setdefault(p, []).append(s['name'])
        for p in sec.get('remote', []): remote_products.setdefault(p, []).append(s['name'])
    pct = lambda n: int(round(n*100/env['total'])) if env['total'] else 0

    def _tool_rows(products):
        if not products: return '<tr><td colspan=2 class=dim>None detected</td></tr>'
        return ''.join(
            f'<tr><td><b>{h(p)}</b></td><td>{len(servers_with)} - {h(", ".join(servers_with[:4]))}{"..." if len(servers_with)>4 else ""}</td></tr>'
            for p, servers_with in sorted(products.items(), key=lambda x:-len(x[1]))
        )
    tools_html = f'''
<div class="tools-grid">
  <div><h5>EDR / Endpoint Protection</h5><table class="t"><tbody>{_tool_rows(edr_products)}</tbody></table></div>
  <div><h5>RMM</h5><table class="t"><tbody>{_tool_rows(rmm_products)}</tbody></table></div>
  <div><h5>Backup Agent</h5><table class="t"><tbody>{_tool_rows(backup_products)}</tbody></table></div>
  <div><h5>Remote Access</h5><table class="t"><tbody>{_tool_rows(remote_products)}</tbody></table></div>
</div>'''

    # ----- OS / platform tables -----
    os_rows = ''.join(f'<tr><td>{h(k) or "<i>unknown</i>"}</td><td>{v}</td></tr>'
                      for k,v in sorted(env['os_counts'].items(), key=lambda x:-x[1]))
    plat_rows = ''.join(f'<tr><td>{h(k)}</td><td>{v}</td></tr>' for k,v in env['by_platform'].items())

    # ----- AD rollup (dedupe by domain so we don't double-count multiple DCs) -----
    ad_by_domain = {}  # domain -> {users, comps, ous, dcs, server_with_data}
    for s in servers:
        a = s.get('ad')
        if not isinstance(a, dict) or not a.get('Installed'): continue
        dom = a.get('DomainName')
        if not dom: continue
        if dom not in ad_by_domain or (s['name'] not in ad_by_domain[dom]['servers']):
            ad_by_domain.setdefault(dom, {
                'users': a.get('UserCount') or 0,
                'computers': a.get('ComputerCount') or 0,
                'ous': a.get('OUCount') or 0,
                'dcs': a.get('DCCount') or 0,
                'forest': a.get('ForestName'),
                'domain_fl': a.get('DomainFL'),
                'forest_fl': a.get('ForestFL'),
                'pdc': a.get('PDCEmulator'),
                'rid': a.get('RIDMaster'),
                'schema': a.get('SchemaMaster'),
                'stale_users': len(a.get('StaleUsers') or []) if isinstance(a.get('StaleUsers'), list) else (a.get('StaleUsers') or 0),
                'stale_comps': len(a.get('StaleComputers') or []) if isinstance(a.get('StaleComputers'), list) else (a.get('StaleComputers') or 0),
                'servers': [s['name']],
                'raw': a,  # keep the richest one for AD tab detail
            })
        else:
            ad_by_domain[dom]['servers'].append(s['name'])
    ad_domains = set(ad_by_domain.keys())
    ad_total_users = sum(v['users'] for v in ad_by_domain.values())
    ad_total_comps = sum(v['computers'] for v in ad_by_domain.values())
    ad_total_ous = sum(v['ous'] for v in ad_by_domain.values())

    # ----- Top servers by criticality (top 8) -----
    crit_list = []
    for s in servers:
        ff = filtered_flags(s)
        c = sum(1 for f in ff if str(f.get('Severity','')).lower() in ('red','critical'))
        w = sum(1 for f in ff if str(f.get('Severity','')).lower() in ('warn','warning'))
        crit_list.append((c*10+w, c, w, s))
    crit_list.sort(key=lambda x: x[0], reverse=True)
    top_rows = ''.join(
        f'<tr onclick="jumpToServer(\'{h(s["id"])}\')"><td><b>{h(s["name"])}</b></td>'
        f'<td>{h(s["os"].replace("Microsoft Windows ",""))}</td>'
        f'<td>{f"<span class=cnt crit>{c}</span>" if c else "0"}</td>'
        f'<td>{f"<span class=cnt warn>{w}</span>" if w else "0"}</td></tr>'
        for _,c,w,s in crit_list[:8]
    )

    # ----- Top SQL DBs by size -----
    all_dbs = []
    for s in servers:
        inst = s['sql'].get('Instances') if isinstance(s['sql'], dict) else None
        if not (isinstance(inst, dict) and inst.get('InstanceName')): continue
        dbs = inst.get('Databases', []) or []
        if not isinstance(dbs, list): dbs = []
        for db in dbs:
            if isinstance(db, dict):
                all_dbs.append((db.get('DataSizeMB') or 0, s['name'], db.get('Name',''), db.get('RecoveryModel','')))
    all_dbs.sort(reverse=True)
    sql_top_rows = ''.join(
        f'<tr><td>{h(srv)}</td><td>{h(db)}</td><td>{sz:,} MB</td><td>{h(rec)}</td></tr>'
        for sz,srv,db,rec in all_dbs[:8]
    )

    # ----- Findings (env-wide) -----
    findings_blocks = []
    if env.get('eol_at_risk'):
        findings_blocks.append(f'<div class="f-row crit" onclick="setViewByName(\'eol\')"><b>OS EOL/at-risk &rsaquo;</b><div>{", ".join(env["eol_at_risk"])}</div></div>')
    if env.get('no_edr'):
        p = pct(len(env["no_edr"]))
        findings_blocks.append(f'<div class="f-row warn"><b>No 3rd-party EDR ({p}%)</b><div>{len(env["no_edr"])}/{env["total"]} on Defender baseline only</div></div>')
    if env.get('high_disk'):
        findings_blocks.append(f'<div class="f-row warn"><b>Disk &gt;85%</b><div>{", ".join(env["high_disk"])}</div></div>')
    if env.get('sql_servers'):
        findings_blocks.append(f'<div class="f-row info" onclick="setViewByName(\'sql\')"><b>SQL Server present ({len(env["sql_servers"])}) &rsaquo;</b><div>{", ".join(env["sql_servers"])}</div></div>')
    if env.get('has_smb1'):
        findings_blocks.append(f'<div class="f-row info"><b>SMB 1.0 enabled (env-wide)</b><div>{len(env["has_smb1"])}/{env["total"]} — rolled up, not per-server spam</div></div>')
    if env.get('has_veeam'):
        findings_blocks.append(f'<div class="f-row info"><b>Veeam detected</b><div>{", ".join(env["has_veeam"])}</div></div>')

    overview_html = f'''
{vc_html}

<h3>Guest-Side Sizing Rollup (from WinRM) <a class="jump" onclick="setViewByName('servers')">View all servers &rsaquo;</a></h3>
<div class="hgrid">
  <div class="hcell"><h4>Compute</h4>
    <div class="stat"><span class="snum">{env['total']}</span><span class="slbl">discovered servers</span></div>
    <div class="stat"><span class="snum">{total_cores_guest}</span><span class="slbl">total cores</span></div>
    <div class="stat"><span class="snum">{total_ram_guest:.0f}</span><span class="slbl">GB RAM total</span></div>
  </div>
  <div class="hcell"><h4>Storage (guest disks)</h4>
    <div class="stat"><span class="snum">{total_disk_total:,.0f}</span><span class="slbl">GB provisioned</span></div>
    <div class="stat"><span class="snum">{total_disk_used:,.0f}</span><span class="slbl">GB used</span></div>
    <div class="stat"><span class="snum">{int(total_disk_used*100/total_disk_total) if total_disk_total else 0}%</span><span class="slbl">utilization</span></div>
  </div>
  <div class="hcell"><h4>Active Directory</h4>
    <div class="stat"><span class="snum">{len(ad_domains)}</span><span class="slbl">domain(s)</span></div>
    <div class="stat"><span class="snum">{ad_total_users:,}</span><span class="slbl">users (sum)</span></div>
    <div class="stat"><span class="snum">{ad_total_comps:,}</span><span class="slbl">computers (sum)</span></div>
    <div class="stat"><span class="snum">{ad_total_ous}</span><span class="slbl">OUs (sum)</span></div>
  </div>
</div>

<h3>Security & Management Coverage</h3>
<div class="hgrid">
  <div class="hcell">
    <h4>Coverage %</h4>
    <div class="stat-row"><b>3rd-party EDR:</b> {have_edr}/{env['total']} ({pct(have_edr)}%)</div>
    <div class="stat-row"><b>RMM:</b> {have_rmm}/{env['total']} ({pct(have_rmm)}%)</div>
    <div class="stat-row"><b>Backup agent:</b> {have_backup}/{env['total']} ({pct(have_backup)}%)</div>
    <div class="stat-row"><b>Remote access:</b> {have_remote}/{env['total']} ({pct(have_remote)}%)</div>
  </div>
  <div class="hcell" style="grid-column:span 2;">
    <h4>OS Breakdown</h4>
    <table class="t"><tbody>{os_rows}</tbody></table>
  </div>
</div>

<h3>Detected Tools in Environment</h3>
{tools_html}

<h3>Top Servers by Severity <a class="jump" onclick="setViewByName('servers')">All servers &rsaquo;</a></h3>
<table class="t clickable"><thead><tr><th>Server</th><th>OS</th><th>Critical</th><th>Warning</th></tr></thead><tbody>{top_rows}</tbody></table>

<h3>Top SQL Databases by Size <a class="jump" onclick="setViewByName('sql')">All SQL &rsaquo;</a></h3>
<table class="t"><thead><tr><th>Server</th><th>Database</th><th>Size</th><th>Recovery</th></tr></thead><tbody>{sql_top_rows or "<tr><td colspan=4 class=dim>No SQL databases.</td></tr>"}</tbody></table>

<h3>Environment-Wide Findings</h3>
<div class="findings-list">{''.join(findings_blocks) or '<p class="dim">No findings rolled up.</p>'}</div>
'''

    # =============== SERVERS TAB ==============
    rail_items = []
    for s in servers:
        has_crit = any(str(f.get('Severity','')).lower() in ('red','critical') for f in filtered_flags(s))
        badge = '<span class="dot crit"></span>' if has_crit else ''
        rail_items.append(
            f'<li data-name="{h(s["name"]).lower()}" data-os="{h(s["os"]).lower()}" '
            f'onclick="showServer(this, \'{h(s["id"])}\')">'
            f'<div class="rail-name">{badge}{h(s["name"])}</div>'
            f'<div class="rail-meta">{h(s["os"].replace("Microsoft Windows ",""))}</div></li>'
        )
    rail_html = '\n'.join(rail_items)

    panels = []
    for s in servers:
        sec = detect_security(s)
        ff = filtered_flags(s)
        crit = sum(1 for f in ff if str(f.get('Severity','')).lower() in ('red','critical'))
        warn = sum(1 for f in ff if str(f.get('Severity','')).lower() in ('warn','warning'))

        # Findings block
        flag_html = ''
        if ff:
            rows = ''.join(
                f'<tr><td><span class="sev {str(f.get("Severity","info")).lower()}">{h(str(f.get("Severity","info")).upper())}</span></td>'
                f'<td><b>{h(f.get("Title",""))}</b></td>'
                f'<td>{h(f.get("Detail",""))}</td></tr>'
                for f in ff
            )
            flag_html = f'<h3>Findings ({len(ff)})</h3><table class="t">{rows}</table>'

        # Disks
        disk_html = ''.join(
            f'<tr><td>{h(d.get("Drive",""))}</td><td>{h(d.get("Label",""))}</td>'
            f'<td>{d.get("TotalGB","")} GB</td><td>{d.get("FreeGB","")} GB</td>'
            f'<td><div class="bar"><div style="width:{min(100,float(d.get("UsedPct") or 0))}%;'
            f'background:{"#ef4444" if (d.get("UsedPct") or 0)>=85 else "#22c55e"}"></div></div> {d.get("UsedPct","")}%</td></tr>'
            for d in s['disks']
        )

        # Security & Mgmt
        sec_html = ''
        for cat, label in [('edr','EDR'),('rmm','RMM'),('backup','Backup'),('remote','Remote Access')]:
            val = ', '.join(sec.get(cat, [])) or '<i class="dim">none detected</i>'
            sec_html += f'<div><b>{label}:</b> {val}</div>'

        # Roles & features
        roles = s.get('roles', [])
        if not isinstance(roles, list): roles = []
        role_names = []
        for r in roles:
            if isinstance(r, dict):
                if r.get('Installed') or r.get('InstallState') == 'Installed':
                    nm = r.get('DisplayName') or r.get('Name') or ''
                    if nm: role_names.append(nm)
            elif isinstance(r, str):
                role_names.append(r)
        roles_html = ''
        if role_names:
            roles_html = f'<h3>Installed Roles & Features ({len(role_names)})</h3>' + \
                ''.join(f'<span class="chip">{h(rn)}</span>' for rn in role_names[:60])

        # AD (per-server)
        ad_html = ''
        adobj = s.get('ad')
        if isinstance(adobj, dict) and adobj.get('Installed'):
            su = adobj.get('StaleUsers')
            sc = adobj.get('StaleComputers')
            ad_html = f'''<h3>Active Directory: {h(adobj.get("DomainName",""))}</h3>
<table class="kv">
  <tr><td>Forest</td><td>{h(adobj.get("ForestName",""))}</td></tr>
  <tr><td>Domain / Forest FL</td><td>{h(adobj.get("DomainFL",""))} / {h(adobj.get("ForestFL",""))}</td></tr>
  <tr><td>PDC Emulator</td><td>{h(adobj.get("PDCEmulator",""))}</td></tr>
  <tr><td>RID Master</td><td>{h(adobj.get("RIDMaster",""))}</td></tr>
  <tr><td>Schema Master</td><td>{h(adobj.get("SchemaMaster",""))}</td></tr>
  <tr><td>Users / Computers / OUs</td><td>{adobj.get("UserCount","?")} / {adobj.get("ComputerCount","?")} / {adobj.get("OUCount","?")}</td></tr>
  <tr><td>Stale users (90+ days)</td><td>{len(su) if isinstance(su, list) else su or 0}</td></tr>
  <tr><td>Stale computers</td><td>{len(sc) if isinstance(sc, list) else sc or 0}</td></tr>
</table>
<p class="dim" style="margin-top:6px;font-size:12px;">Full AD breakdown on the Active Directory tab.</p>'''

        # SQL
        sql_html = ''
        sql_inst = s['sql'].get('Instances') if isinstance(s['sql'], dict) else None
        if isinstance(sql_inst, dict) and sql_inst.get('InstanceName'):
            dbs = sql_inst.get('Databases', []) or []
            if not isinstance(dbs, list): dbs = []
            db_rows = ''.join(
                f'<tr><td>{h(db.get("Name"))}</td><td>{db.get("DataSizeMB","")} MB</td>'
                f'<td>{db.get("LogSizeMB","")} MB</td>'
                f'<td>{h(db.get("RecoveryModel",""))}</td>'
                f'<td>{db.get("CompatLevel","")}</td>'
                f'<td>{h(db.get("LastFullBackup",""))}</td></tr>'
                for db in dbs
            )
            sql_html = f'''<h3>SQL Server</h3>
<p><b>{h(sql_inst.get("InstanceName"))}</b> — {h(sql_inst.get("Edition",""))} {h(sql_inst.get("Version",""))} — EOL {h(sql_inst.get("EOLDate",""))} ({h(sql_inst.get("EOLStatus",""))})</p>
<p>Service Account: <code>{h(sql_inst.get("ServiceAccount",""))}</code></p>
<table class="t"><thead><tr><th>Database</th><th>Data</th><th>Log</th><th>Recovery</th><th>Compat</th><th>Last Full</th></tr></thead><tbody>{db_rows}</tbody></table>'''

        # IIS
        iis_html = ''
        iis = s.get('iis') or {}
        if isinstance(iis, dict) and iis.get('Installed'):
            sites = iis.get('Sites', []) or []
            if not isinstance(sites, list): sites = []
            pools = iis.get('AppPools', []) or []
            if not isinstance(pools, list): pools = []
            if sites or pools:
                site_rows = ''.join(
                    f'<tr><td>{h(x.get("Name"))}</td><td>{h(x.get("State"))}</td><td>{h(x.get("Bindings",""))}</td><td>{h(x.get("PhysicalPath",""))}</td></tr>'
                    for x in sites if isinstance(x, dict)
                )
                pool_rows = ''.join(
                    f'<tr><td>{h(x.get("Name"))}</td><td>{h(x.get("State"))}</td><td>{h(x.get("ManagedRuntimeVersion",""))}</td></tr>'
                    for x in pools if isinstance(x, dict)
                )
                iis_html = '<h3>IIS</h3>'
                if site_rows:
                    iis_html += f'<table class="t"><thead><tr><th>Site</th><th>State</th><th>Bindings</th><th>Path</th></tr></thead><tbody>{site_rows}</tbody></table>'
                if pool_rows:
                    iis_html += f'<h4>App Pools</h4><table class="t"><thead><tr><th>Pool</th><th>State</th><th>.NET</th></tr></thead><tbody>{pool_rows}</tbody></table>'

        # Hyper-V VMs
        hv_html = ''
        vms = s.get('hyperv_vms', [])
        if isinstance(vms, list) and vms:
            vm_rows = ''.join(
                f'<tr><td>{h(v.get("Name"))}</td><td>{h(v.get("State"))}</td><td>{v.get("vCPU","")}</td>'
                f'<td>{v.get("RAMgb","")} GB</td><td>{v.get("UptimeHours","")}h</td><td>{v.get("Snapshots","")}</td></tr>'
                for v in vms if isinstance(v, dict)
            )
            hv_html = f'<h3>Hyper-V VMs ({len(vms)})</h3><table class="t"><thead><tr><th>Name</th><th>State</th><th>vCPU</th><th>RAM</th><th>Uptime</th><th>Snaps</th></tr></thead><tbody>{vm_rows}</tbody></table>'

        # Shares
        share_html = ''
        if s['shares']:
            rows = ''.join(
                f'<tr><td>{h(sh.get("Name"))}</td><td>{h(sh.get("Path"))}</td></tr>'
                for sh in s['shares'] if isinstance(sh, dict)
            )
            if rows:
                share_html = f'<h3>File Shares ({len(s["shares"])})</h3><table class="t"><thead><tr><th>Share</th><th>Path</th></tr></thead><tbody>{rows}</tbody></table>'

        # Network adapters
        net_html = ''
        adapters = s.get('network_adapters', [])
        if isinstance(adapters, list) and adapters:
            ad_rows = ''.join(
                f'<tr><td>{h(a.get("Description") or a.get("Name",""))}</td>'
                f'<td>{h(a.get("IPv4") or a.get("IPAddress",""))}</td>'
                f'<td>{h(a.get("MACAddress",""))}</td>'
                f'<td>{h(a.get("DefaultGateway",""))}</td>'
                f'<td>{h(a.get("DNSServers",""))}</td></tr>'
                for a in adapters if isinstance(a, dict)
            )
            if ad_rows:
                net_html = f'<h3>Network</h3><table class="t"><thead><tr><th>Adapter</th><th>IP</th><th>MAC</th><th>Gateway</th><th>DNS</th></tr></thead><tbody>{ad_rows}</tbody></table>'

        # Apps (collapsed)
        apps_html = ''
        if s['apps']:
            app_rows = ''.join(
                f'<tr><td>{h(a.get("Name",""))}</td><td>{h(a.get("Publisher",""))}</td><td>{h(a.get("Version",""))}</td></tr>'
                for a in s['apps'] if isinstance(a, dict)
            )
            apps_html = f'<details><summary><h3 style="display:inline;">Installed Applications ({len(s["apps"])})</h3></summary><table class="t"><thead><tr><th>Name</th><th>Publisher</th><th>Version</th></tr></thead><tbody>{app_rows}</tbody></table></details>'

        panels.append(f'''
<div class="panel" id="p-{h(s['id'])}">
  <div class="phead">
    <h2>{h(s['name'])}</h2>
    <div class="meta">{h(s['os'])} — {h(s['platform'])} — {s.get('cores','?')}c / {s.get('ram_gb','?')} GB RAM</div>
    <div class="badges">
      {f'<span class="badge crit">{crit} critical</span>' if crit else ''}
      {f'<span class="badge warn">{warn} warning</span>' if warn else ''}
      <span class="badge">{len(s['apps'])} apps</span>
      <span class="badge">{s['services_count']} services</span>
      <span class="badge">{len(role_names)} roles</span>
    </div>
  </div>
  <div class="grid">
    <div class="cell"><h4>System</h4>
      <div>Hostname: {h(s['name'])}</div>
      <div>Domain: {h(s['domain'])}</div>
      <div>Uptime: {s['uptime_days']} days</div>
      <div>CPU: {h(s['cpu'])}</div>
      <div>Hardware: {h(s['mfr'])} {h(s['model'])}</div>
    </div>
    <div class="cell"><h4>Security & Management</h4>{sec_html}</div>
  </div>
  {flag_html}
  <h3>Disks</h3>
  <table class="t"><thead><tr><th>Drive</th><th>Label</th><th>Total</th><th>Free</th><th>Used</th></tr></thead><tbody>{disk_html}</tbody></table>
  {net_html}
  {roles_html}
  {ad_html}
  {sql_html}
  {iis_html}
  {hv_html}
  {share_html}
  {apps_html}
</div>''')

    # =============== HYPERVISOR TAB ==============
    hv_tab_html = '<p class="dim">No hypervisor inventory collected.</p>'
    if vc:
        cluster = vc.get('Cluster') or {}
        hosts = vc.get('ESXHosts') or []
        if not isinstance(hosts, list): hosts = []
        vms = vc.get('VMs') or []
        if not isinstance(vms, list): vms = []
        ds_list = vc.get('Datastores') or []
        if not isinstance(ds_list, list): ds_list = []
        lic = vc.get('Licenses') or []
        if not isinstance(lic, list): lic = []

        host_rows = ''.join(
            f'<tr><td><b>{h(ho.get("Name",""))}</b></td><td>{h(ho.get("Vendor",""))} {h(ho.get("Model",""))}</td>'
            f'<td>{h(ho.get("ServiceTag",""))}</td>'
            f'<td>{h(ho.get("CPUModel",""))}</td>'
            f'<td>{ho.get("Cores","?")}c</td><td>{ho.get("RAMgb","?")} GB</td>'
            f'<td>{ho.get("NICs","?")}</td>'
            f'<td>{ho.get("CPUUsagePct","?")}%</td><td>{ho.get("MemUsagePct","?")}%</td>'
            f'<td>{h(ho.get("Hypervisor",""))}</td></tr>'
            for ho in hosts if isinstance(ho, dict)
        )

        ds_rows = ''.join(
            f'<tr><td><b>{h(ds.get("Name",""))}</b></td><td>{h(ds.get("Type",""))}</td>'
            f'<td>{(ds.get("CapacityGB") or 0):,.0f} GB</td>'
            f'<td>{(ds.get("FreeGB") or 0):,.0f} GB</td>'
            f'<td>{(ds.get("ConsumedGB") or 0):,.0f} GB</td>'
            f'<td><div class="bar"><div style="width:{min(100,float(ds.get("UsedPct") or 0))}%;'
            f'background:{"#ef4444" if (ds.get("UsedPct") or 0)>=85 else "#22c55e"}"></div></div> {ds.get("UsedPct","?")}%</td></tr>'
            for ds in ds_list if isinstance(ds, dict)
        )

        # VM table sorted by allocated vCPU desc
        vms_sorted = sorted([v for v in vms if isinstance(v,dict)], key=lambda v: -(int(v.get('vCPUs') or 0)))
        vm_rows = ''.join(
            f'<tr><td><b>{h(v.get("Name",""))}</b></td><td>{h(v.get("PowerState",""))}</td>'
            f'<td>{v.get("vCPUs","?")}</td><td>{(v.get("RAMgb") or 0):.0f} GB</td>'
            f'<td>{(v.get("DiskConsumedGB") or 0):,.0f} / {(v.get("DiskCapGB") or 0):,.0f} GB</td>'
            f'<td>{h(v.get("GuestOS",""))}</td>'
            f'<td>{h(v.get("Datastore",""))}</td>'
            f'<td>{h(v.get("ToolStatus",""))}</td></tr>'
            for v in vms_sorted
        )

        lic_rows = ''.join(
            f'<tr><td>{h(l.get("Name",""))}</td><td><code>{h(str(l.get("Key",""))[:24])}</code></td>'
            f'<td>{l.get("Used","?")}</td><td>{l.get("Total","?")}</td><td>{h(l.get("Expiry",""))}</td></tr>'
            for l in lic if isinstance(l, dict)
        )

        cluster_cap_gib = (cluster.get('CapacityMiB') or 0) / 1024
        cluster_used_gib = (cluster.get('ConsumedMiB') or 0) / 1024
        hv_tab_html = f'''
<div class="hgrid">
  <div class="hcell">
    <h4>vCenter</h4>
    <div class="stat-row"><b>Server:</b> <code>{h(vc.get("Server",""))}</code></div>
    <div class="stat-row"><b>Version:</b> {h(vc.get("Version",""))}</div>
    <div class="stat-row"><b>API:</b> {h(vc.get("APIVersion",""))}</div>
    <div class="stat-row"><b>Collected:</b> {h(vc.get("CollectedAt",""))[:10]}</div>
    <div class="stat-row"><b>Perf window:</b> {vc.get("DurationDays","?")} days</div>
  </div>
  <div class="hcell">
    <h4>Cluster: {h(cluster.get("Name",""))}</h4>
    <div class="stat-row"><b>Hosts in cluster:</b> {cluster.get("NumHosts","?")} ({len(hosts)} collected)</div>
    <div class="stat-row"><b>Capacity:</b> {cluster_cap_gib/1024:.1f} TiB</div>
    <div class="stat-row"><b>Consumed:</b> {cluster_used_gib/1024:.1f} TiB</div>
    <div class="stat-row"><b>CPU usage (cluster):</b> {cluster.get("CPUUsagePct","?")}%</div>
    <div class="stat-row"><b>Mem usage (cluster):</b> {cluster.get("MemUsagePct","?")}%</div>
  </div>
  <div class="hcell">
    <h4>Workload Sizing</h4>
    <div class="stat"><span class="snum">{len([v for v in vms if isinstance(v,dict) and str(v.get("PowerState","")).lower() in ("poweredon","on")])}</span><span class="slbl">VMs powered on</span></div>
    <div class="stat"><span class="snum">{sum(int(v.get("vCPUs") or 0) for v in vms if isinstance(v,dict) and str(v.get("PowerState","")).lower() in ("poweredon","on"))}</span><span class="slbl">total vCPU allocated</span></div>
    <div class="stat"><span class="snum">{sum(float(v.get("RAMgb") or 0) for v in vms if isinstance(v,dict) and str(v.get("PowerState","")).lower() in ("poweredon","on")):.0f}</span><span class="slbl">GB RAM allocated</span></div>
  </div>
</div>

<h3>ESX Hosts ({len(hosts)} of {cluster.get("NumHosts","?") if cluster.get("NumHosts") else len(hosts)})</h3>
<table class="t"><thead><tr><th>Host</th><th>Hardware</th><th>Service Tag</th><th>CPU Model</th><th>Cores</th><th>RAM</th><th>NICs</th><th>CPU Use</th><th>Mem Use</th><th>Hypervisor</th></tr></thead><tbody>{host_rows or "<tr><td colspan=10 class=dim>No hosts.</td></tr>"}</tbody></table>

<h3>Datastores ({len(ds_list)})</h3>
<table class="t"><thead><tr><th>Datastore</th><th>Type</th><th>Capacity</th><th>Free</th><th>Consumed</th><th>Used %</th></tr></thead><tbody>{ds_rows}</tbody></table>

<h3>Licenses</h3>
<table class="t"><thead><tr><th>Product</th><th>Key (truncated)</th><th>Used</th><th>Total</th><th>Expiry</th></tr></thead><tbody>{lic_rows or "<tr><td colspan=5 class=dim>No license data.</td></tr>"}</tbody></table>

<h3>All VMs ({len(vms)}) <span class="dim">sorted by vCPU allocated</span></h3>
<table class="t"><thead><tr><th>VM</th><th>Power</th><th>vCPU</th><th>RAM</th><th>Disk (used/prov)</th><th>Guest OS</th><th>Datastore</th><th>VMware Tools</th></tr></thead><tbody>{vm_rows}</tbody></table>
'''

    # =============== ACTIVE DIRECTORY TAB ==============
    ad_tab_html = '<p class="dim">No Active Directory data captured (no domain controllers reachable via WinRM).</p>'
    if ad_by_domain:
        ad_sections = []
        for dom, info in ad_by_domain.items():
            a = info['raw']
            fsmo = a.get('FSMORoles') or {}
            if not isinstance(fsmo, dict): fsmo = {}
            stale_users = a.get('StaleUsers') if isinstance(a.get('StaleUsers'), list) else []
            stale_comps = a.get('StaleComputers') if isinstance(a.get('StaleComputers'), list) else []
            su_rows = ''.join(
                f'<tr><td>{h(u.get("SamAccountName",""))}</td><td>{h(u.get("Name",""))}</td><td>{h(u.get("LastLogon",""))}</td></tr>'
                for u in stale_users if isinstance(u, dict)
            )
            sc_rows = ''.join(
                f'<tr><td>{h(c.get("SamAccountName",""))}</td><td>{h(c.get("Name",""))}</td><td>{h(c.get("LastLogon",""))}</td></tr>'
                for c in stale_comps if isinstance(c, dict)
            )
            fsmo_rows = ''.join(
                f'<tr><td>{h(k)}</td><td>{h(v)}</td></tr>'
                for k,v in fsmo.items() if isinstance(v,str)
            )
            fl_warn = 'warn' if '2016' in str(info['domain_fl']) or '2012' in str(info['domain_fl']) else ''
            ad_sections.append(f'''
<h3>Domain: {h(dom)}</h3>
<div class="hgrid">
  <div class="hcell">
    <h4>Identity</h4>
    <div class="stat-row"><b>Forest:</b> {h(info['forest'])}</div>
    <div class="stat-row"><b>Domain FL:</b> <span class="badge {fl_warn}">{h(info['domain_fl'])}</span></div>
    <div class="stat-row"><b>Forest FL:</b> {h(info['forest_fl'])}</div>
    <div class="stat-row"><b>Discovered via:</b> {h(", ".join(info['servers']))}</div>
  </div>
  <div class="hcell">
    <h4>Topology</h4>
    <div class="stat"><span class="snum">{info['dcs']}</span><span class="slbl">domain controllers</span></div>
    <div class="stat"><span class="snum">{info['users']:,}</span><span class="slbl">users</span></div>
    <div class="stat"><span class="snum">{info['computers']:,}</span><span class="slbl">computers</span></div>
    <div class="stat"><span class="snum">{info['ous']}</span><span class="slbl">OUs</span></div>
  </div>
  <div class="hcell">
    <h4>Hygiene</h4>
    <div class="stat"><span class="snum">{info['stale_users']}</span><span class="slbl">stale users (90+ days)</span></div>
    <div class="stat"><span class="snum">{info['stale_comps']}</span><span class="slbl">stale computers</span></div>
  </div>
</div>

<h4>FSMO Roles</h4>
<table class="kv">
  <tr><td>PDC Emulator</td><td>{h(info['pdc'])}</td></tr>
  <tr><td>RID Master</td><td>{h(info['rid'])}</td></tr>
  <tr><td>Schema Master</td><td>{h(info['schema'])}</td></tr>
  {fsmo_rows}
</table>

{f'<h4>Stale Users ({len(stale_users)})</h4><details><summary class="dim">Show stale users</summary><table class="t"><thead><tr><th>SAM</th><th>Name</th><th>Last Logon</th></tr></thead><tbody>{su_rows}</tbody></table></details>' if su_rows else ''}

{f'<h4>Stale Computers ({len(stale_comps)})</h4><details><summary class="dim">Show stale computers</summary><table class="t"><thead><tr><th>SAM</th><th>Name</th><th>Last Logon</th></tr></thead><tbody>{sc_rows}</tbody></table></details>' if sc_rows else ''}
''')
        ad_tab_html = ''.join(ad_sections)

    # =============== SQL TAB (env-wide rollup) ==============
    sql_global = []
    for s in servers:
        inst = s['sql'].get('Instances') if isinstance(s['sql'], dict) else None
        if not (isinstance(inst, dict) and inst.get('InstanceName')): continue
        dbs = inst.get('Databases', []) or []
        if not isinstance(dbs, list): dbs = []
        for db in dbs:
            if not isinstance(db, dict): continue
            sql_global.append({
                'server': s['name'],
                'instance': inst.get('InstanceName',''),
                'db': db.get('Name',''),
                'size': db.get('DataSizeMB', 0) or 0,
                'recovery': db.get('RecoveryModel',''),
                'compat': db.get('CompatLevel',''),
                'last_full': db.get('LastFullBackup',''),
            })
    sql_global.sort(key=lambda x: -x['size'])
    sql_rows = ''.join(
        f'<tr><td>{h(r["server"])}</td><td>{h(r["instance"])}</td><td>{h(r["db"])}</td>'
        f'<td>{r["size"]:,} MB</td><td>{h(r["recovery"])}</td><td>{r["compat"]}</td><td>{h(r["last_full"])}</td></tr>'
        for r in sql_global
    )
    sql_tab_html = f'''
<h3>All SQL Databases ({len(sql_global)})</h3>
<p class="dim">Sorted by data size desc. Compat &lt; 130 means legacy SQL 2016- behavior.</p>
<table class="t"><thead><tr><th>Server</th><th>Instance</th><th>Database</th><th>Size</th><th>Recovery</th><th>Compat</th><th>Last Full</th></tr></thead><tbody>{sql_rows or "<tr><td colspan=7 class=dim>No SQL instances detected.</td></tr>"}</tbody></table>
'''

    # =============== PRIVATE CLOUD SIZING TAB ==============
    # Uses 95th-percentile CPU/RAM from vsphere-perf JSON (SDT's own collector
    # OR Nutanix Collector XLSX parsed via parse_ntnx_collector.py - same
    # schema). Falls back to a banner if P95 data is null.
    BUFFER = 1.20  # 20% growth buffer (M5 standard)
    pc_tab_html = ''
    sized_vms = []
    have_p95 = False
    if vc:
        vc_vms = vc.get('VMs') or []
        if not isinstance(vc_vms, list): vc_vms = []
        for v in vc_vms:
            if not isinstance(v, dict): continue
            if str(v.get('PowerState','')).lower() not in ('poweredon','on'): continue
            cpu_obj = v.get('CPU') or {}
            mem_obj = v.get('Memory') or {}
            cpu_p95 = cpu_obj.get('P95') if isinstance(cpu_obj, dict) else None
            mem_p95 = mem_obj.get('P95') if isinstance(mem_obj, dict) else None
            vcpus = int(v.get('vCPUs') or 0)
            ram_gb = float(v.get('RAMgb') or 0)
            disk_used = float(v.get('DiskConsumedGB') or 0)
            disk_prov = float(v.get('DiskCapGB') or 0)
            if cpu_p95 is not None and mem_p95 is not None:
                have_p95 = True
            # Sized = (utilization% / 100) * allocated * buffer
            sized_cpu = (float(cpu_p95) / 100.0) * vcpus * BUFFER if cpu_p95 is not None else None
            sized_ram = (float(mem_p95) / 100.0) * ram_gb * BUFFER if mem_p95 is not None else None
            sized_disk = disk_used * BUFFER if disk_used else disk_prov * BUFFER
            sized_vms.append({
                'name': v.get('Name',''), 'os': v.get('GuestOS',''),
                'vcpus': vcpus, 'ram_gb': ram_gb,
                'cpu_p95': cpu_p95, 'mem_p95': mem_p95,
                'sized_cpu': sized_cpu, 'sized_ram': sized_ram,
                'disk_used': disk_used, 'disk_prov': disk_prov, 'sized_disk': sized_disk,
                'datastore': v.get('Datastore',''),
            })

    if not vc:
        pc_tab_html = '<p class="dim">No vCenter inventory collected — sizing requires hypervisor data.</p>'
    elif not have_p95:
        pc_tab_html = '''
<div class="finding warn" style="background:#78350f33;border-left:4px solid #f59e0b;padding:14px 18px;border-radius:6px;margin-bottom:18px;">
  <b>95th-percentile utilization data not available.</b>
  <p>SDT's vSphere perf collector returned null values OR perf-counter collection wasn't enabled at the cluster.
  To get accurate sizing, run the Nutanix Collector against this vCenter, then drop the resulting XLSX into
  the session folder and run <code>parse_ntnx_collector.py</code> followed by report regeneration.</p>
  <p>Showing <b>provisioned values only</b> below — these are <i>not</i> sizing inputs and will overestimate target cluster requirements.</p>
</div>'''
    if sized_vms:
        # Totals
        sum_alloc_cpu = sum(v['vcpus'] for v in sized_vms)
        sum_alloc_ram = sum(v['ram_gb'] for v in sized_vms)
        sum_sized_cpu = sum((v['sized_cpu'] or 0) for v in sized_vms)
        sum_sized_ram = sum((v['sized_ram'] or 0) for v in sized_vms)
        sum_disk_used = sum(v['disk_used'] for v in sized_vms)
        sum_sized_disk = sum(v['sized_disk'] for v in sized_vms)

        sized_rows = ''.join(
            f'<tr><td><b>{h(v["name"])}</b></td>'
            f'<td>{h((v["os"] or "")[:40])}</td>'
            f'<td>{v["vcpus"]}</td>'
            f'<td>{("{:.1f}%".format(v["cpu_p95"]) if v["cpu_p95"] is not None else "—")}</td>'
            f'<td>{("{:.1f}".format(v["sized_cpu"]) if v["sized_cpu"] is not None else "—")}</td>'
            f'<td>{v["ram_gb"]:.0f}</td>'
            f'<td>{("{:.1f}%".format(v["mem_p95"]) if v["mem_p95"] is not None else "—")}</td>'
            f'<td>{("{:.1f}".format(v["sized_ram"]) if v["sized_ram"] is not None else "—")}</td>'
            f'<td>{v["disk_used"]:,.0f}</td>'
            f'<td>{v["sized_disk"]:,.0f}</td></tr>'
            for v in sorted(sized_vms, key=lambda x:-(x['sized_cpu'] or 0))
        )
        pc_tab_html += f'''
<h3>Sizing Methodology</h3>
<p class="dim">Sized resource = (95th-percentile utilization &times; allocated) &times; 1.20 (20% growth buffer). Powered-off VMs excluded.</p>

<div class="hgrid">
  <div class="hcell">
    <h4>Allocated (current)</h4>
    <div class="stat"><span class="snum">{sum_alloc_cpu}</span><span class="slbl">vCPU allocated</span></div>
    <div class="stat"><span class="snum">{sum_alloc_ram:.0f}</span><span class="slbl">GB RAM allocated</span></div>
    <div class="stat"><span class="snum">{sum_disk_used:,.0f}</span><span class="slbl">GB disk consumed</span></div>
  </div>
  <div class="hcell">
    <h4>Right-Sized (P95 + 20%)</h4>
    <div class="stat"><span class="snum">{sum_sized_cpu:.0f}</span><span class="slbl">vCPU needed</span></div>
    <div class="stat"><span class="snum">{sum_sized_ram:.0f}</span><span class="slbl">GB RAM needed</span></div>
    <div class="stat"><span class="snum">{sum_sized_disk:,.0f}</span><span class="slbl">GB disk needed</span></div>
  </div>
  <div class="hcell">
    <h4>Savings vs Provisioned</h4>
    <div class="stat"><span class="snum">{int((sum_alloc_cpu - sum_sized_cpu)*100/sum_alloc_cpu) if sum_alloc_cpu else 0}%</span><span class="slbl">vCPU reduction</span></div>
    <div class="stat"><span class="snum">{int((sum_alloc_ram - sum_sized_ram)*100/sum_alloc_ram) if sum_alloc_ram else 0}%</span><span class="slbl">RAM reduction</span></div>
  </div>
</div>

<h3>Per-VM Sizing ({len(sized_vms)} VMs)</h3>
<table class="t"><thead><tr>
  <th rowspan=2>VM</th><th rowspan=2>Guest OS</th>
  <th colspan=3 style="text-align:center;">CPU</th>
  <th colspan=3 style="text-align:center;">RAM (GB)</th>
  <th colspan=2 style="text-align:center;">Disk (GB)</th>
</tr><tr>
  <th>Allocated</th><th>P95</th><th>Sized</th>
  <th>Allocated</th><th>P95</th><th>Sized</th>
  <th>Used</th><th>Sized</th>
</tr></thead><tbody>{sized_rows}</tbody></table>
'''

    # =============== COMMVAULT SIZING TAB ==============
    # Front-end backup target sizing: 1:1 consumed match + 20% growth (M5 BDR rule).
    # Excludes AS/400, Linux appliances, vCenter/VCSA management VMs by default.
    EXCLUDE_NAMES = ('vcsa', 'vcenter', 'vrops', 'vrli', 'vmcl', 'qeagle', 'haeagle',
                     'image-hpe-bnservices', 'esxi', 'vxrail-manager')
    cv_vms = []
    cv_excluded = []
    cv_warnings = []
    if vc:
        vc_vms = vc.get('VMs') or []
        if not isinstance(vc_vms, list): vc_vms = []
        for v in vc_vms:
            if not isinstance(v, dict): continue
            nm = (v.get('Name','') or '').lower()
            os_str = (v.get('GuestOS','') or '').lower()
            disk_used = float(v.get('DiskConsumedGB') or 0)
            if any(x in nm for x in EXCLUDE_NAMES) or 'photon' in os_str:
                cv_excluded.append({'name': v.get('Name',''), 'reason': 'mgmt/appliance', 'disk': disk_used})
                continue
            if v.get('IsLinux'):
                # Linux is in scope but flag — file-system backup may need agent
                pass
            if disk_used <= 0:
                cv_warnings.append(v.get('Name',''))
            cv_vms.append({'name': v.get('Name',''), 'os': v.get('GuestOS',''),
                          'disk_used': disk_used, 'disk_prov': float(v.get('DiskCapGB') or 0),
                          'is_linux': v.get('IsLinux', False)})

    cv_total_used = sum(v['disk_used'] for v in cv_vms)
    cv_sized = cv_total_used * BUFFER

    cv_rows = ''.join(
        f'<tr><td><b>{h(v["name"])}</b></td><td>{h((v["os"] or "")[:40])}</td>'
        f'<td>{"Linux" if v["is_linux"] else "Windows"}</td>'
        f'<td>{v["disk_used"]:,.0f}</td><td>{v["disk_prov"]:,.0f}</td>'
        f'<td>{v["disk_used"]*BUFFER:,.0f}</td></tr>'
        for v in sorted(cv_vms, key=lambda x:-x['disk_used'])
    )
    cv_excl_rows = ''.join(
        f'<tr><td>{h(v["name"])}</td><td>{h(v["reason"])}</td><td>{v["disk"]:,.0f} GB (excluded)</td></tr>'
        for v in cv_excluded
    )
    cv_tab_html = f'''
<h3>Commvault Sizing Methodology</h3>
<p class="dim">Front-end backup target = 1:1 match of consumed disk on in-scope VMs &times; 1.20 (20% growth buffer).
Excludes hypervisor management appliances (vCenter/VCSA, vROps, Photon-based VMs) and AS/400 systems by default.</p>

<div class="hgrid">
  <div class="hcell">
    <h4>In Scope</h4>
    <div class="stat"><span class="snum">{len(cv_vms)}</span><span class="slbl">VMs in scope</span></div>
    <div class="stat"><span class="snum">{cv_total_used:,.0f}</span><span class="slbl">GB consumed (1:1)</span></div>
  </div>
  <div class="hcell">
    <h4>Sized Target</h4>
    <div class="stat"><span class="snum">{cv_sized:,.0f}</span><span class="slbl">GB front-end (P95+20%)</span></div>
    <div class="stat"><span class="snum">{cv_sized/1024:.1f}</span><span class="slbl">TiB front-end</span></div>
  </div>
  <div class="hcell">
    <h4>Excluded</h4>
    <div class="stat"><span class="snum">{len(cv_excluded)}</span><span class="slbl">excluded VMs</span></div>
    <div class="stat"><span class="snum">{sum(v["disk"] for v in cv_excluded):,.0f}</span><span class="slbl">GB excluded</span></div>
  </div>
</div>

<h3>In-Scope VMs ({len(cv_vms)})</h3>
<table class="t"><thead><tr><th>VM</th><th>Guest OS</th><th>Type</th><th>Consumed (GB)</th><th>Provisioned (GB)</th><th>Sized (GB)</th></tr></thead><tbody>{cv_rows or "<tr><td colspan=6 class=dim>No in-scope VMs.</td></tr>"}</tbody></table>

{f'<h3>Excluded ({len(cv_excluded)})</h3><table class="t"><thead><tr><th>VM</th><th>Reason</th><th>Size</th></tr></thead><tbody>{cv_excl_rows}</tbody></table>' if cv_excluded else ''}

{f'<p class="dim" style="margin-top:14px;">Warning: {len(cv_warnings)} VM(s) had 0 GB consumed (Tools missing or powered off). Review before quoting.</p>' if cv_warnings else ''}
'''

    # =============== EOL TAB ==============
    def _eol_class(status_text):
        t = str(status_text or '').lower()
        if 'eol' in t or 'unsupported' in t or 'past' in t: return 'eol'
        if 'near' in t or 'extended' in t or 'approaching' in t: return 'near'
        if 'supported' in t: return 'ok'
        # Heuristic on OS names if status is blank
        if '2008' in t or '2012' in t: return 'eol'
        if '2016' in t: return 'near'
        return ''
    eol_rows = []
    for s in servers:
        os_eol = s.get('os_eol','')
        cls = _eol_class(os_eol) or _eol_class(s['os'])
        eol_rows.append(f'<tr class="eol-row {cls}"><td>{h(s["name"])}</td><td>OS</td><td>{h(s["os"])}</td><td>{h(os_eol)}</td></tr>')
        inst = s['sql'].get('Instances') if isinstance(s['sql'], dict) else None
        if isinstance(inst, dict) and inst.get('InstanceName'):
            sql_status = inst.get('EOLStatus','')
            sql_cls = _eol_class(sql_status)
            # SQL 2016 EOL = 2026-07-14 - flag as near-EOL
            if '2016' in str(inst.get('Edition','')) and sql_cls == 'ok': sql_cls = 'near'
            eol_rows.append(f'<tr class="eol-row {sql_cls}"><td>{h(s["name"])}</td><td>SQL</td><td>{h(inst.get("Edition",""))} ({h(inst.get("Version",""))})</td><td>{h(sql_status)} — {h(inst.get("EOLDate",""))}</td></tr>')
    eol_html = f'''
<h3>End-of-Life / Lifecycle Status ({len(eol_rows)} rows)</h3>
<p class="dim">Row color: <span style="background:rgba(127,29,29,0.4);padding:1px 6px;">EOL/unsupported</span> · <span style="background:rgba(120,53,15,0.4);padding:1px 6px;">near-EOL</span> · <span style="background:rgba(20,83,45,0.3);padding:1px 6px;">supported</span></p>
<table class="t"><thead><tr><th>Server</th><th>Component</th><th>Detail</th><th>EOL Status</th></tr></thead><tbody>{''.join(eol_rows)}</tbody></table>
'''

    os_breakdown = ', '.join(f'{k or "?"}: {v}' for k,v in env.get('os_counts',{}).items())

    return f'''<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>{h(client)} - Server Discovery</title>
<style>
*{{box-sizing:border-box;}}body{{margin:0;font-family:'Segoe UI',system-ui,sans-serif;background:#0f172a;color:#e2e8f0;font-size:13.5px;}}
.hdr{{padding:14px 24px;background:#1e293b;border-bottom:1px solid #334155;display:flex;justify-content:space-between;align-items:center;}}
.hdr h1{{margin:0;font-size:18px;}}
.hdr .meta{{color:#94a3b8;font-size:12.5px;}}
.topnav{{background:#1e293b;border-bottom:1px solid #334155;padding:0 24px;display:flex;gap:4px;}}
.tab-btn{{padding:12px 22px;background:transparent;border:none;color:#94a3b8;font-weight:600;font-size:13px;cursor:pointer;border-bottom:2px solid transparent;font-family:inherit;}}
.tab-btn:hover{{color:#e2e8f0;}}
.tab-btn.active{{color:#60a5fa;border-bottom-color:#60a5fa;}}
.view{{display:none;}}
.view.active{{display:block;}}
.pad{{padding:24px;}}
.container{{display:flex;height:calc(100vh - 105px);}}
.rail{{width:260px;background:#1e293b;border-right:1px solid #334155;overflow-y:auto;flex-shrink:0;}}
.rail input{{width:calc(100% - 24px);margin:12px;padding:8px 12px;background:#0f172a;border:1px solid #334155;border-radius:6px;color:#e2e8f0;font-size:12.5px;}}
.rail ul{{margin:0;padding:0;list-style:none;}}
.rail li{{padding:9px 16px;cursor:pointer;border-bottom:1px solid #1e293b;}}
.rail li:hover{{background:#334155;}}
.rail li.active{{background:#1e40af;}}
.rail-name{{font-weight:600;font-size:12.5px;display:flex;align-items:center;gap:8px;}}
.rail-meta{{font-size:11px;color:#94a3b8;margin-top:2px;}}
.dot{{width:8px;height:8px;border-radius:50%;display:inline-block;}}
.dot.crit{{background:#ef4444;}}
.main{{flex:1;overflow-y:auto;padding:20px 28px;}}
.panel{{display:none;}}
.panel.active{{display:block;}}
.phead h2{{margin:0;font-size:22px;}}
.phead .meta{{color:#94a3b8;font-size:12.5px;margin:4px 0 10px;}}
.badges{{display:flex;gap:6px;margin-bottom:18px;flex-wrap:wrap;}}
.badge{{padding:3px 10px;border-radius:99px;font-size:11px;background:#334155;color:#cbd5e1;font-weight:600;}}
.badge.crit{{background:#7f1d1d;color:#fecaca;}}
.badge.warn{{background:#78350f;color:#fed7aa;}}
.grid{{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px;}}
.cell{{background:#1e293b;border:1px solid #334155;padding:14px 18px;border-radius:8px;}}
.cell h4{{margin:0 0 8px;font-size:11px;color:#93c5fd;text-transform:uppercase;letter-spacing:.5px;}}
.cell div{{font-size:12.5px;margin:3px 0;}}
h3{{margin:20px 0 8px;font-size:13px;color:#93c5fd;text-transform:uppercase;letter-spacing:.5px;}}
h4{{margin:14px 0 6px;font-size:12px;color:#93c5fd;text-transform:uppercase;letter-spacing:.5px;}}
table.t, table.kv{{width:100%;border-collapse:collapse;font-size:12.5px;margin-bottom:8px;}}
table.t th, table.kv th{{text-align:left;padding:6px 10px;background:#1e293b;color:#94a3b8;font-size:11px;text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid #334155;}}
table.t td, table.kv td{{padding:6px 10px;border-bottom:1px solid #1e293b;vertical-align:top;}}
table.kv td:first-child{{color:#94a3b8;width:200px;}}
.sev{{display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700;}}
.sev.red,.sev.critical{{background:#7f1d1d;color:#fecaca;}}
.sev.warn,.sev.warning{{background:#78350f;color:#fed7aa;}}
.sev.info{{background:#1e3a8a;color:#bfdbfe;}}
.bar{{display:inline-block;width:140px;height:8px;background:#334155;border-radius:3px;overflow:hidden;vertical-align:middle;}}
.bar div{{height:100%;}}
.chip{{display:inline-block;padding:3px 9px;border-radius:99px;font-size:11px;background:#334155;color:#cbd5e1;margin:2px;}}
.crit{{color:#fecaca;}} .warn{{color:#fed7aa;}} .info{{color:#bfdbfe;}}
.dim{{color:#64748b;font-style:italic;}}
.hgrid{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:18px;margin-bottom:18px;}}
.hcell{{background:#1e293b;border:1px solid #334155;border-radius:8px;padding:16px 20px;}}
.hcell h3{{margin:0 0 10px;}}
.stat{{display:flex;align-items:baseline;gap:10px;margin:6px 0;}}
.snum{{font-size:24px;font-weight:700;color:#60a5fa;}}
.slbl{{font-size:12px;color:#94a3b8;}}
.findings-list .f-row{{padding:10px 14px;border-radius:6px;background:#1e293b;border-left:4px solid #38bdf8;margin:6px 0;}}
.findings-list .f-row.crit{{border-color:#ef4444;}}
.findings-list .f-row.warn{{border-color:#f59e0b;}}
.findings-list .f-row b{{display:block;color:#e2e8f0;margin-bottom:3px;}}
.findings-list .f-row div{{color:#94a3b8;font-size:12.5px;}}
.findings-list .f-row[onclick]{{cursor:pointer;}}
.findings-list .f-row[onclick]:hover{{background:#334155;}}
.jump{{font-size:12px;color:#60a5fa;cursor:pointer;text-decoration:none;font-weight:600;margin-left:8px;}}
.jump:hover{{text-decoration:underline;}}
.stat-row{{font-size:12.5px;margin:4px 0;color:#cbd5e1;}}
.stat-row b{{color:#94a3b8;font-weight:600;}}
.clickable tbody tr{{cursor:pointer;}}
.clickable tbody tr:hover{{background:#334155;}}
.cnt{{display:inline-block;padding:1px 8px;border-radius:99px;font-size:11px;font-weight:700;}}
.cnt.crit{{background:#7f1d1d;color:#fecaca;}}
.cnt.warn{{background:#78350f;color:#fed7aa;}}
.tools-grid{{display:grid;grid-template-columns:1fr 1fr;gap:14px;}}
.tools-grid > div{{background:#1e293b;border:1px solid #334155;border-radius:8px;padding:14px 18px;}}
.tools-grid h5{{margin:0 0 8px;font-size:11px;color:#93c5fd;text-transform:uppercase;letter-spacing:.5px;}}
.eol-row.eol{{background:rgba(127,29,29,0.18);}}
.eol-row.near{{background:rgba(120,53,15,0.18);}}
.eol-row.ok{{background:rgba(20,83,45,0.12);}}
details summary{{cursor:pointer;list-style:none;}}
details summary::-webkit-details-marker{{display:none;}}
details summary::before{{content:'▶ ';color:#94a3b8;font-size:10px;}}
details[open] summary::before{{content:'▼ ';}}
code{{background:#0f172a;padding:1px 6px;border-radius:3px;font-family:'Cascadia Mono',Consolas,monospace;font-size:12px;}}
</style></head><body>
<div class="hdr">
  <h1>{h(client)} — Server Discovery</h1>
  <div class="meta">{session['date']} · {env['total']} servers · {os_breakdown}</div>
</div>
<div class="topnav">
  <button class="tab-btn active" onclick="setView(this,'overview')">Overview</button>
  <button class="tab-btn" onclick="setView(this,'servers')">Servers ({env['total']})</button>
  <button class="tab-btn" onclick="setView(this,'hypervisor')">Hypervisor</button>
  <button class="tab-btn" onclick="setView(this,'ad')">Active Directory</button>
  <button class="tab-btn" onclick="setView(this,'sql')">SQL</button>
  <button class="tab-btn" onclick="setView(this,'pcsizing')">Private Cloud Sizing</button>
  <button class="tab-btn" onclick="setView(this,'cvsizing')">Commvault Sizing</button>
  <button class="tab-btn" onclick="setView(this,'eol')">EOL Risks</button>
</div>

<div class="view active" id="v-overview"><div class="pad">{overview_html}</div></div>

<div class="view" id="v-servers">
  <div class="container">
    <div class="rail">
      <input type="text" placeholder="Filter servers..." oninput="filterRail(this.value)">
      <ul id="rail-list">{rail_html}</ul>
    </div>
    <div class="main" id="main">{''.join(panels)}</div>
  </div>
</div>

<div class="view" id="v-hypervisor"><div class="pad">{hv_tab_html}</div></div>
<div class="view" id="v-ad"><div class="pad">{ad_tab_html}</div></div>
<div class="view" id="v-sql"><div class="pad">{sql_tab_html}</div></div>
<div class="view" id="v-pcsizing"><div class="pad">{pc_tab_html}</div></div>
<div class="view" id="v-cvsizing"><div class="pad">{cv_tab_html}</div></div>
<div class="view" id="v-eol"><div class="pad">{eol_html}</div></div>

<script>
function setView(btn, name) {{
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.getElementById('v-'+name).classList.add('active');
  if (name === 'servers' && !document.querySelector('.panel.active')) {{
    const first = document.querySelector('.rail li');
    if (first) first.click();
  }}
}}
function showServer(li, id) {{
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  document.getElementById('p-'+id).classList.add('active');
  document.querySelectorAll('.rail li').forEach(x => x.classList.remove('active'));
  li.classList.add('active');
  document.getElementById('main').scrollTop = 0;
}}
function filterRail(q) {{
  q = q.toLowerCase();
  document.querySelectorAll('.rail li').forEach(li => {{
    li.style.display = (li.dataset.name.includes(q) || li.dataset.os.includes(q)) ? '' : 'none';
  }});
}}
function setViewByName(name) {{
  const btn = document.querySelector(`.tab-btn[onclick*="'${{name}}'"]`);
  if (btn) btn.click();
}}
function jumpToServer(id) {{
  setViewByName('servers');
  const li = document.querySelector(`.rail li[onclick*="'${{id}}'"]`);
  if (li) li.click();
}}
</script>
</body></html>'''

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('usage: python gen_report.py <manifest.json | session_dir> [out_html]'); sys.exit(1)
    arg = Path(sys.argv[1])
    if arg.is_dir():
        session_dir = arg
    elif arg.is_file() and arg.name.endswith('.json'):
        session_dir = arg.parent
    else:
        print(f'error: not a manifest or directory: {arg}'); sys.exit(2)
    session = load_session(session_dir)
    out = render(session)
    if len(sys.argv) >= 3:
        out_path = Path(sys.argv[2])
    else:
        client_slug = (session['client'] or 'Client').replace(' ','_').replace('/','_')
        date = session['date'] or 'undated'
        out_path = session_dir / f'{client_slug}-DiscoveryReport-{date}.html'
    out_path.write_text(out, encoding='utf-8')
    print(f'Wrote {out_path} ({len(out):,} bytes, {len(session["servers"])} servers)')
