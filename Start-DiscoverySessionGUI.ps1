<#
.SYNOPSIS
    Start-DiscoverySessionGUI.ps1 - Browser-based UI for SDT discovery sessions.

.DESCRIPTION
    Spins up a local HTTP listener on http://localhost:8080, serves a
    self-contained HTML wizard, opens the user's default browser.

    The wizard collects client info, hypervisor target, server list, and
    credentials, then runs discovery with live progress updates.

    Reuses existing Invoke-ServerDiscovery.ps1 per target and gen_report.py
    for the final HTML report. All credentials live in PowerShell memory
    only - never written to disk.

    Console-mode Start-DiscoverySession.ps1 is untouched. This script is
    additive.

.NOTES
    v4.0-alpha  |  2026-04-21
    Requires: PowerShell 5.1+
#>
param(
    [int]    $Port = 8080,
    [switch] $NoOpenBrowser
)

$ErrorActionPreference = 'Stop'
$script:Version   = '4.0.9'
$script:ScriptDir = $PSScriptRoot
$script:BaseUrl   = "http://localhost:$Port"

# Session state shared between HTTP handlers and discovery worker jobs.
$script:Session = [hashtable]::Synchronized(@{
    Status     = 'idle'                       # idle | running | complete | error
    Client     = ''
    OutputDir  = ''
    SessionDir = ''
    Targets    = @()                          # list of { Name; Address; State; Phase; Buddy; Started; Finished }
    LogTail    = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]@())
    ReportPath = ''
    ReportZipPath = ''
    MissingTargets = @()
    StartedAt  = $null
    FinishedAt = $null
})

$script:BuddyFrames = @('(^_^) ','(^_^)>','(o_o) ','(o_o)>','(-_-) ','(>_<) ','(*_*) ','(^_-) ','(._.) ','(T_T) ','(^o^) ','(x_x) ')

function Add-Log([string]$msg) {
    $stamp = (Get-Date).ToString('HH:mm:ss')
    [void]$script:Session.LogTail.Add("[$stamp] $msg")
    # Keep only last 500 lines
    while ($script:Session.LogTail.Count -gt 500) {
        $script:Session.LogTail.RemoveAt(0)
    }
}

# Runs collect_vsphere_perf.py with auto-retry across username formats.
# Returns hashtable: @{ ok=$true/$false; user=<succeeded-format>; log=<combined-log>; file=<json-path>; error=<short>; }
function Invoke-VsphereCollect {
    param(
        [string] $PyExe,
        [string] $ScriptPath,
        [string] $VCenterHost,
        [string] $UserRaw,
        [string] $PassRaw,
        [string] $OutputDir
    )

    # Build username variants to try in order
    $variants = [System.Collections.ArrayList]@()
    [void]$variants.Add($UserRaw)

    if ($UserRaw -notmatch '[\\@]') {
        [void]$variants.Add("$UserRaw@vsphere.local")
        [void]$variants.Add("VSPHERE.LOCAL\$UserRaw")
    }
    if ($UserRaw -match '@') {
        $parts = $UserRaw -split '@', 2
        [void]$variants.Add("$($parts[1].ToUpper())\$($parts[0])")
    }
    if ($UserRaw -match '\\') {
        $parts = $UserRaw -split '\\', 2
        [void]$variants.Add("$($parts[1])@$($parts[0].ToLower())")
    }
    $uniqueVariants = @($variants | Select-Object -Unique)

    $env:SDT_HV_PASS = $PassRaw
    $combined = ''
    $succeeded = $null
    $successFile = $null
    $nonAuthFail = $false
    try {
        foreach ($u in $uniqueVariants) {
            # Track files in output dir BEFORE running so we know which one is new
            $beforeSet = @()
            try { $beforeSet = (Get-ChildItem $OutputDir -Filter '*inventory*.json' -EA 0 | ForEach-Object { $_.FullName }) } catch { }

            $pyArgs = @(
                $ScriptPath,
                '--vcenter',  $VCenterHost,
                '--user',     $u,
                '--pass-env', 'SDT_HV_PASS',
                '--output',   $OutputDir
            )
            $out = & $PyExe $pyArgs 2>&1 | Out-String
            $combined += "`n=== attempt with --user '$u' ===`n$out`n"

            # Find any NEW inventory file written during this attempt
            $afterSet = @()
            try { $afterSet = (Get-ChildItem $OutputDir -Filter '*inventory*.json' -EA 0 | ForEach-Object { $_.FullName }) } catch { }
            $new = $afterSet | Where-Object { $_ -notin $beforeSet } | Select-Object -First 1
            if (-not $new) {
                # Also accept the newest vsphere-perf-*.json if present
                $new = Get-ChildItem $OutputDir -Filter 'vsphere-perf-*.json' -EA 0 | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | ForEach-Object { $_.FullName }
            }
            if ($new -and (Test-Path $new) -and (Get-Item $new).Length -gt 1000) {
                $succeeded = $u
                $successFile = $new
                break
            }

            # If the error isn't auth-related, stop - retrying won't help
            if ($out -notmatch '(?i)(not\s+authenticated|InvalidLogin|Login failed|Incorrect user|password|credentials|SAML|401|403|unauthorized)') {
                $nonAuthFail = $true
                break
            }
        }
    } finally {
        $env:SDT_HV_PASS = $null
    }

    if ($succeeded) {
        return @{ ok=$true; user=$succeeded; log=$combined; file=$successFile }
    }
    $errMsg = if ($nonAuthFail) { 'non-auth failure - check connectivity / SSL / script output' } else { "auth failed for all $($uniqueVariants.Count) username format(s) tried" }
    return @{ ok=$false; log=$combined; error=$errMsg; triedUsers=$uniqueVariants }
}

# -----------------------------------------------------------------------------
# HTML UI - self-contained single page, served as a here-string
# -----------------------------------------------------------------------------
$script:HtmlUI = @'
<!DOCTYPE html><html lang="en"><head>
<meta charset="utf-8"><title>Magna5 SDT - Discovery Session</title>
<style>
:root {
  --bg:#0B1220; --surface:#131A2B; --elevated:#1A2340; --elevated-2:#222C4A;
  --border:rgba(255,255,255,0.07); --border-2:rgba(255,255,255,0.11);
  --text:#E6EAF2; --muted:#8B95A8; --dim:#5A6478;
  --accent:#4F8CFF; --accent-2:#8B5CF6;
  --ok:#22C55E; --warn:#F59E0B; --crit:#EF4444; --info:#38BDF8;
  --mono:"Cascadia Code","Consolas",ui-monospace,monospace;
  --sans:"Segoe UI Variable Display","Segoe UI",system-ui,-apple-system,sans-serif;
}
*{box-sizing:border-box;}html,body{margin:0;padding:0;}
body{background:var(--bg);color:var(--text);font-family:var(--sans);font-size:14px;line-height:1.5;
  background-image:radial-gradient(1200px 600px at 80% -200px,rgba(79,140,255,0.12),transparent 60%),
                   radial-gradient(900px 400px at -200px 200px,rgba(139,92,246,0.08),transparent 50%);
  background-attachment:fixed;min-height:100vh;}
.wrap{max-width:1280px;margin:0 auto;padding:0 28px 60px;}
.hdr{background:rgba(11,18,32,0.75);backdrop-filter:blur(18px);
  border-bottom:1px solid var(--border);padding:14px 28px;display:flex;justify-content:space-between;align-items:center;
  position:sticky;top:0;z-index:100;}
.brand{font-size:16px;font-weight:700;letter-spacing:.5px;
  background:linear-gradient(92deg,#fff 0%,#a7b3ca 100%);
  -webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent;}
.sub{color:var(--muted);font-size:12px;}
.ver-chip{background:var(--elevated);border:1px solid var(--border);color:var(--muted);
  font-size:11px;font-weight:600;padding:4px 10px;border-radius:99px;}
.tab-nav{display:flex;gap:4px;margin:18px 0 22px;background:var(--surface);border:1px solid var(--border);
  border-radius:14px;padding:6px;width:fit-content;}
.tab-btn{background:transparent;border:none;color:var(--muted);padding:10px 22px;font-size:13px;font-weight:600;
  border-radius:10px;cursor:pointer;font-family:var(--sans);}
.tab-btn:hover{color:var(--text);background:var(--elevated);}
.tab-btn.active{background:linear-gradient(135deg,#1c2540 0%,#222c4f 100%);color:var(--text);
  box-shadow:0 0 0 1px var(--border-2),inset 0 1px 0 rgba(255,255,255,0.04);}
.tab-btn:disabled{opacity:.4;cursor:not-allowed;}
.tab-pane{display:none;}
.tab-pane.active{display:block;animation:fadeIn .2s ease;}
@keyframes fadeIn{from{opacity:0;transform:translateY(4px);}to{opacity:1;transform:translateY(0);}}
.card{background:linear-gradient(180deg,var(--surface) 0%,rgba(19,26,43,0.94) 100%);
  border:1px solid var(--border);border-radius:14px;padding:22px 24px;margin-bottom:18px;}
.card-title{font-size:14px;font-weight:700;letter-spacing:-.01em;margin-bottom:2px;}
.card-sub{font-size:12px;color:var(--muted);margin-bottom:14px;}
.section-hdr{font-size:11px;font-weight:700;color:var(--accent);text-transform:uppercase;letter-spacing:.8px;
  margin:18px 0 10px;}
label{display:block;font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;
  letter-spacing:.6px;margin-bottom:6px;}
input,textarea,select{width:100%;background:var(--elevated);border:1px solid var(--border);
  border-radius:8px;padding:10px 14px;color:var(--text);font-family:var(--sans);font-size:13px;
  outline:none;transition:border-color .12s;}
input:focus,textarea:focus,select:focus{border-color:var(--accent);}
textarea{font-family:var(--mono);min-height:120px;resize:vertical;}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px;}
.grid3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px;}
.field{margin-bottom:14px;}
.help{font-size:11px;color:var(--dim);margin-top:4px;}
.btn{background:linear-gradient(135deg,var(--accent) 0%,var(--accent-2) 100%);border:none;color:#fff;
  font-weight:700;font-size:13px;padding:12px 28px;border-radius:10px;cursor:pointer;letter-spacing:.3px;}
.btn:hover{filter:brightness(1.12);}
.btn:disabled{opacity:.4;cursor:not-allowed;filter:none;}
.btn-secondary{background:var(--elevated);color:var(--text);border:1px solid var(--border-2);}
.pill{display:inline-flex;align-items:center;gap:6px;padding:3px 10px;border-radius:99px;
  font-size:11px;font-weight:600;white-space:nowrap;}
.pill .dot{width:6px;height:6px;border-radius:50%;}
.pill.idle{background:rgba(139,149,168,.10);color:var(--muted);border:1px solid var(--border-2);}
.pill.running{background:rgba(79,140,255,.14);color:#93b9ff;border:1px solid rgba(79,140,255,.25);}
.pill.running .dot{background:#4f8cff;box-shadow:0 0 8px rgba(79,140,255,.7);animation:pulse 1s infinite;}
.pill.ok{background:rgba(34,197,94,.12);color:#86efac;border:1px solid rgba(34,197,94,.22);}
.pill.err{background:rgba(239,68,68,.14);color:#fca5a5;border:1px solid rgba(239,68,68,.25);}
@keyframes pulse{0%,100%{opacity:1;}50%{opacity:.4;}}
.pbar{background:rgba(255,255,255,.06);border-radius:4px;height:8px;overflow:hidden;margin:10px 0;}
.pbar > div{height:100%;background:linear-gradient(90deg,var(--accent),var(--accent-2));transition:width .3s;}
.target-row{display:grid;grid-template-columns:28px 1fr auto auto;gap:14px;align-items:center;
  padding:12px 14px;border-bottom:1px solid var(--border);font-size:13px;}
.target-row:last-child{border-bottom:none;}
.buddy{font-family:var(--mono);font-size:12px;color:var(--info);width:48px;}
.target-name{font-family:var(--mono);font-weight:600;color:var(--text);}
.target-phase{color:var(--muted);font-size:12px;}
.logbox{background:#07101f;border:1px solid var(--border);border-radius:10px;padding:14px 16px;
  font-family:var(--mono);font-size:11.5px;color:#c7d1df;height:300px;overflow-y:auto;white-space:pre-wrap;
  line-height:1.55;}
.footer{text-align:center;padding:30px 0 10px;color:var(--dim);font-size:11px;}
.callout{background:rgba(245,158,11,0.08);border-left:3px solid var(--warn);padding:10px 14px;
  border-radius:6px;font-size:12px;color:#fcd34d;margin:12px 0;}
.pw-wrap{position:relative;}
.pw-wrap input{padding-right:40px;}
.pw-toggle{position:absolute;right:8px;top:50%;transform:translateY(-50%);background:transparent;border:none;
  color:var(--muted);cursor:pointer;padding:4px 6px;border-radius:4px;display:flex;align-items:center;justify-content:center;}
.pw-toggle:hover{color:var(--text);background:var(--elevated-2);}
.pw-toggle svg{width:18px;height:18px;}
.hint{display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;border-radius:50%;
  background:var(--elevated-2);color:var(--muted);font-size:10px;font-weight:700;margin-left:8px;cursor:help;
  border:1px solid var(--border-2);vertical-align:middle;user-select:none;position:relative;}
.hint:hover{background:var(--accent);color:#fff;border-color:var(--accent);}
.hint::after{content:attr(data-tip);position:absolute;top:calc(100% + 8px);left:50%;transform:translateX(-50%);
  background:#0b1220;color:var(--text);border:1px solid var(--border-2);border-radius:8px;padding:10px 14px;
  font-size:11.5px;font-weight:500;white-space:normal;width:280px;text-align:left;line-height:1.5;letter-spacing:.1px;
  box-shadow:0 8px 24px rgba(0,0,0,.35);opacity:0;pointer-events:none;transition:opacity .15s;z-index:150;}
.hint:hover::after{opacity:1;}
.hint.right::after{left:auto;right:-8px;transform:none;}
.banner-run{background:linear-gradient(125deg,rgba(79,140,255,0.18) 0%,rgba(139,92,246,0.14) 100%);
  border:1px solid var(--border-2);border-radius:14px;padding:20px 24px;margin-bottom:18px;
  display:flex;justify-content:space-between;align-items:center;}
.banner-run .title{font-size:17px;font-weight:700;}
.banner-run .meta{color:var(--muted);font-size:12px;margin-top:3px;}
</style></head><body>
<div class="hdr">
  <div style="display:flex;align-items:center;gap:14px;">
    <!-- M5 logo mark -->
    <svg width="34" height="34" viewBox="0 0 34 34" xmlns="http://www.w3.org/2000/svg" style="flex-shrink:0;">
      <defs><linearGradient id="lg" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stop-color="#4F8CFF"/>
        <stop offset="100%" stop-color="#8B5CF6"/>
      </linearGradient></defs>
      <rect x="1" y="1" width="32" height="32" rx="8" fill="url(#lg)"/>
      <text x="17" y="22" text-anchor="middle" font-family="Segoe UI Variable Display, Segoe UI, system-ui" font-size="13" font-weight="700" fill="#fff" letter-spacing="-0.5">M5</text>
    </svg>
    <div class="brand">MAGNA5</div>
    <div style="color:var(--dim);">/</div>
    <div class="sub">SDT - Discovery Session</div>
  </div>
  <span class="ver-chip">SDT GUI v__VERSION__</span>
</div>
<div class="wrap">
<div class="tab-nav">
  <button class="tab-btn active" id="tb-setup" onclick="setTab('setup')">Setup</button>
  <button class="tab-btn" id="tb-run" onclick="setTab('run')">Run</button>
  <button class="tab-btn" id="tb-report" onclick="setTab('report')" disabled>Report</button>
</div>

<!-- SETUP TAB -->
<div id="tab-setup" class="tab-pane active">
<form id="setupForm" onsubmit="submitSetup(event)">

<div class="card">
<div class="card-title">Session <span class="hint" data-tip="Client name appears in the final HTML report header. Output folder is where every JSON, log, and final report for this discovery run gets written.">i</span></div>
<div class="card-sub">Client name and output folder. Credentials live in memory only - never saved to disk.</div>
<div class="grid2">
<div class="field"><label>Client name <span class="hint" data-tip="Appears in the report header. Use the sales-facing name (e.g. Acme Corporation).">i</span></label>
<input name="client" required placeholder="Acme Corporation"></div>
<div class="field"><label>Output folder <span class="hint" data-tip="Discovery JSONs + HTML report land here. Default is fine; a per-client subfolder is created automatically.">i</span></label>
<input name="outputDir" value="C:\Temp\sdt\sessions" required></div>
</div>
</div>

<div class="card">
<div class="card-title">Admin credentials <span class="hint" data-tip="Used to remotely sign in to every ticked VM + every manual target. Needs WinRM remoting permissions (domain admin works; local admin works for workgroup hosts).">i</span></div>
<div class="card-sub">Domain admin or local admin that can log into the Windows targets. Held in memory only - never written to disk.</div>
<div class="grid2">
<div class="field"><label>Domain / Local admin <span class="hint" data-tip="Format: DOMAIN\\username for a domain admin, user@domain.local for UPN, or .\\username for a local admin on workgroup hosts.">i</span></label>
<input name="winrmUser" placeholder="DOMAIN\administrator or admin@contoso.local"></div>
<div class="field"><label>Password <span class="hint" data-tip="Password for the admin account. Held in memory only for this session - no file, no registry, no persistence.">i</span></label>
<div class="pw-wrap">
<input name="winrmPass" type="password" autocomplete="off">
<button type="button" class="pw-toggle" onclick="togglePw(this)" title="Show/hide password" aria-label="Toggle password visibility">
<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>
</button>
</div>
</div>
</div>
<div style="display:flex;align-items:center;gap:12px;margin-top:10px;">
<button type="button" class="btn btn-secondary" id="testCredsBtn" onclick="testCreds()">Test creds</button>
<span id="credStatus" style="font-size:12px;color:var(--muted);">Validates against the domain (DOMAIN\user or UPN). Local accounts (.\user) require a manual target.</span>
</div>
</div>

<div class="card">
<div class="card-title">Hypervisor (optional) <span class="hint" data-tip="Hit 'Scan Hypervisor' to connect, list every VM, and tick which ones to collect Windows data from. Skip this whole section if you're running against bare-metal servers only.">i</span></div>
<div class="card-sub">Connect to vCenter or an ESXi host. Hit <strong>Scan Hypervisor</strong> to discover VMs, then tick which ones to collect Windows data from. Or leave "None" for bare-metal-only runs.</div>
<div class="grid3">
<div class="field"><label>Type <span class="hint" data-tip="vCenter/ESXi uses the vSphere SOAP API (works against both). Hyper-V is stubbed for now. 'None' skips hypervisor inventory entirely.">i</span></label>
<select name="hvType"><option value="none">None</option><option value="vsphere">vCenter / ESXi</option><option value="hyperv">Hyper-V Host</option></select></div>
<div class="field"><label>IP / FQDN <span class="hint" data-tip="Hostname or IP of your vCenter server (or single ESXi host). Port 443 must be reachable.">i</span></label>
<input name="hvHost" placeholder="192.168.10.75"></div>
<div class="field"><label>User <span class="hint" data-tip="vCenter SSO account (e.g. administrator@vsphere.local) or local ESXi root. Needs read access to inventory + perf counters.">i</span></label>
<input name="hvUser" placeholder="administrator@vsphere.local"></div>
</div>
<div class="grid2">
<div class="field"><label>Password <span class="hint" data-tip="Password for the hypervisor account. Held in memory only for this session.">i</span></label>
<div class="pw-wrap">
<input name="hvPass" type="password" autocomplete="off">
<button type="button" class="pw-toggle" onclick="togglePw(this)" title="Show/hide password" aria-label="Toggle password visibility">
<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>
</button>
</div>
</div>
<div class="field" style="display:flex;align-items:flex-end;">
<button type="button" class="btn btn-secondary" id="scanHvBtn" onclick="scanHv()">Scan Hypervisor</button>
</div>
</div>
<div id="scanStatus" style="margin-top:10px;font-size:12px;color:var(--muted);"></div>
</div>

<!-- Discovered VMs (appears after hypervisor scan) -->
<div class="card" id="discoveredCard" style="display:none;">
<div class="card-title">Discovered VMs <span class="hint" data-tip="Every VM found on the hypervisor. Ticked rows get per-server WinRM/Windows discovery in addition to the HV inventory. Use the Filter input to narrow by name, IP, or OS.">i</span></div>
<div class="card-sub">Tick rows to include in per-server Windows discovery. Linux / appliance / vCenter boxes are auto-unchecked.</div>
<div style="display:flex;gap:10px;align-items:center;margin:10px 0;">
<input id="vmFilter" placeholder="Filter by name, IP, OS..." oninput="renderVmTable()" style="flex:1;">
<button type="button" class="btn btn-secondary" onclick="toggleAllVms(true)">Select all</button>
<button type="button" class="btn btn-secondary" onclick="toggleAllVms(false)">Clear</button>
</div>
<div id="vmTableWrap" style="max-height:360px;overflow-y:auto;border:1px solid var(--border);border-radius:10px;">
<table class="dt" id="vmTable" style="width:100%;font-size:12.5px;"><thead><tr>
<th style="width:40px;text-align:center;padding:8px 10px;"><input type="checkbox" id="vmAllCbx" onclick="toggleAllVms(this.checked)"></th>
<th style="padding:8px 10px;">Name</th>
<th style="padding:8px 10px;">IP</th>
<th style="padding:8px 10px;">Guest OS</th>
<th style="padding:8px 10px;">Power</th>
</tr></thead><tbody id="vmTableBody"></tbody></table>
</div>
<div id="vmCount" style="font-size:11px;color:var(--muted);margin-top:8px;">0 VMs discovered</div>
</div>

<div class="card">
<div class="card-title">Windows Targets (manual additions) <span class="hint" data-tip="Only needed for hosts that aren't in the hypervisor scan above - e.g. physical boxes, DMZ VMs, or anything the hypervisor can't see. Leave blank if everything's in the HV.">i</span></div>
<div class="card-sub">Hosts not in the hypervisor above. One per line. IP or hostname. Leave blank if HV scan covers everything.
<br>Example: <code style="font-family:var(--mono);color:var(--info);">192.168.10.4</code> or <code style="font-family:var(--mono);color:var(--info);">QES-OFFICE-DC</code></div>
<div class="field"><label>Manual targets <span class="hint" data-tip="One host per line. IP or hostname. The script will try remote WinRM first; if that fails, you'll see an error row in the Run tab.">i</span></label>
<textarea name="targets" placeholder="Optional - only if a host isn't in the HV"></textarea></div>
<div class="callout">
<strong>Coming soon:</strong> subnet auto-discovery (scan button). For now, list targets manually or use AD/DNS export.
</div>
</div>

<div style="display:flex;justify-content:flex-end;gap:10px;">
<button type="submit" class="btn">Run Discovery</button>
</div>
</form>
</div>

<!-- RUN TAB -->
<div id="tab-run" class="tab-pane">
<div class="banner-run">
  <div>
    <div class="title" id="runTitle">Waiting to start...</div>
    <div class="meta" id="runMeta">Submit the setup form to begin.</div>
  </div>
  <span id="runPill" class="pill idle"><span class="dot"></span><span id="runPillText">Idle</span></span>
</div>

<div class="card">
<div class="card-title">Progress</div>
<div class="pbar"><div id="overallBar" style="width:0%"></div></div>
<div id="progressText" style="font-size:12px;color:var(--muted);">0 / 0 targets complete</div>
<div class="section-hdr">Targets</div>
<div id="targetList">
<div style="color:var(--muted);font-style:italic;padding:20px;text-align:center;">No run in progress.</div>
</div>
</div>

<div class="card">
<div class="card-title">Log</div>
<div class="logbox" id="logbox">Waiting for discovery to start...</div>
</div>
</div>

<!-- REPORT TAB -->
<div id="tab-report" class="tab-pane">
<div class="card">
<div class="card-title">Report</div>
<div class="card-sub" id="reportSub">The HTML report will appear here when the discovery session completes.</div>
<div id="reportContent">
<div style="color:var(--muted);font-style:italic;padding:20px;text-align:center;">Run a discovery session first.</div>
</div>
</div>
</div>

<div class="footer" style="line-height:1.8;">
  <div style="font-weight:600;color:var(--muted);">Intellectual property of Magna5, Inc. &middot; All rights reserved.</div>
  <div>Written by <strong style="color:var(--text);">Matthew Kelly</strong> &middot; Magna5 Solutions Engineering</div>
  <div>Questions: <a href="mailto:matthew.kelly@magna5.com" style="color:var(--accent);">matthew.kelly@magna5.com</a></div>
  <div style="margin-top:6px;color:var(--dim);">SDT GUI v__VERSION__</div>
</div>
</div>

<script>
function setTab(t){
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.toggle('active', b.id==='tb-'+t));
  document.querySelectorAll('.tab-pane').forEach(p=>p.classList.toggle('active', p.id==='tab-'+t));
}

// Holds the last hypervisor scan result so submit can pass checked VMs.
window._discoveredVMs = [];

function getHvFields(){
  const form = document.getElementById('setupForm');
  const fd = new FormData(form);
  return {
    hvType: fd.get('hvType') || 'none',
    hvHost: (fd.get('hvHost')||'').trim(),
    hvUser: (fd.get('hvUser')||'').trim(),
    hvPass: fd.get('hvPass') || ''
  };
}

async function scanHv(){
  const btn = document.getElementById('scanHvBtn');
  const status = document.getElementById('scanStatus');
  const hv = getHvFields();
  if (hv.hvType === 'none') { alert('Pick a hypervisor type first.'); return; }
  if (!hv.hvHost || !hv.hvUser || !hv.hvPass) { alert('Hypervisor host, user, and password required.'); return; }

  btn.disabled = true; btn.textContent = 'Scanning...';
  status.style.color = 'var(--muted)';
  status.innerHTML = 'Connecting to ' + escapeHtml(hv.hvHost) + '. This can take 30-60 seconds for a mid-size vCenter...';

  try {
    const resp = await fetch('/api/hv-scan', {
      method: 'POST', headers: {'Content-Type':'application/json'},
      body: JSON.stringify(hv)
    });
    const data = await resp.json();
    if (!resp.ok || !data.ok) {
      const err = data.error || 'scan failed';
      // Also surface the collector log tail if present so user can self-diagnose
      if (data.log) {
        status.style.color = 'var(--crit)';
        status.innerHTML = '<strong>Scan failed:</strong> ' + escapeHtml(err) +
          '<details style="margin-top:10px;"><summary style="cursor:pointer;color:var(--muted);">Show Python/collector output</summary>' +
          '<pre style="background:#07101f;color:#c7d1df;padding:10px;border-radius:6px;font-family:var(--mono);font-size:11px;max-height:300px;overflow:auto;white-space:pre-wrap;">' +
          escapeHtml(data.log) + '</pre></details>';
        return;
      }
      throw new Error(err);
    }

    window._discoveredVMs = (data.vms || []).map(v => ({...v, _checked: isWinVM(v)}));
    document.getElementById('discoveredCard').style.display = '';
    renderVmTable();
    status.style.color = 'var(--ok)';
    status.textContent = 'Scanned ' + window._discoveredVMs.length + ' VMs. Tick the ones to include in Windows discovery below.';
  } catch(err) {
    status.style.color = 'var(--crit)';
    status.textContent = 'Scan failed: ' + err.message;
  } finally {
    btn.disabled = false; btn.textContent = 'Scan Hypervisor';
  }
}

function isWinVM(v){
  const os = (v.GuestOS || v.guestOs || '').toLowerCase();
  const nm = (v.Name || v.name || '').toLowerCase();
  // Uncheck Linux, Photon, vCenter appliances, and anything with 'vcsa'/'vcenter' in the name
  if (/linux|photon|ubuntu|debian|centos|redhat|bsd|coreos/.test(os)) return false;
  if (/vcsa|vcenter|esxi\b/.test(nm)) return false;
  return true;
}

function renderVmTable(){
  const tbody = document.getElementById('vmTableBody');
  const count = document.getElementById('vmCount');
  const filter = (document.getElementById('vmFilter').value || '').toLowerCase();
  const vms = window._discoveredVMs || [];
  const filtered = vms.filter(v => {
    if (!filter) return true;
    const hay = [v.Name, v.name, v.IPs, v.ips, v.GuestOS, v.guestOs, v.PowerState, v.powerState].filter(Boolean).join(' ').toLowerCase();
    return hay.includes(filter);
  });
  tbody.innerHTML = filtered.map((v, i) => {
    const nm  = v.Name || v.name || '';
    const ip  = v.IPs || v.ips || '';
    const os  = v.GuestOS || v.guestOs || '';
    const ps  = v.PowerState || v.powerState || '';
    const idx = vms.indexOf(v);
    const powCls = /on|POWERED_ON/i.test(ps) ? 'ok' : 'neutral';
    return `<tr>
      <td style="text-align:center;padding:6px 10px;"><input type="checkbox" data-vm-idx="${idx}" ${v._checked ? 'checked':''} onchange="window._discoveredVMs[${idx}]._checked=this.checked;updateVmCount();"></td>
      <td style="padding:6px 10px;font-family:var(--mono);font-weight:600;">${escapeHtml(nm)}</td>
      <td style="padding:6px 10px;font-family:var(--mono);font-size:11px;">${escapeHtml(ip)}</td>
      <td style="padding:6px 10px;font-size:11.5px;color:var(--muted);">${escapeHtml(os)}</td>
      <td style="padding:6px 10px;"><span class="pill ${powCls}"><span class="dot"></span>${escapeHtml(ps)}</span></td>
    </tr>`;
  }).join('') || '<tr><td colspan="5" style="color:var(--muted);padding:14px;text-align:center;">No matches.</td></tr>';
  updateVmCount();
}

function updateVmCount(){
  const vms = window._discoveredVMs || [];
  const checked = vms.filter(v => v._checked).length;
  document.getElementById('vmCount').textContent = checked + ' of ' + vms.length + ' selected';
}

function toggleAllVms(on){
  (window._discoveredVMs || []).forEach(v => v._checked = !!on);
  renderVmTable();
}

async function submitSetup(e){
  e.preventDefault();
  const form = document.getElementById('setupForm');
  const data = Object.fromEntries(new FormData(form).entries());
  // Start with manual targets
  let targets = (data.targets||'').split(/\r?\n/).map(s=>s.trim()).filter(Boolean);
  // Add checked VMs (prefer IP, fall back to hostname)
  const checked = (window._discoveredVMs || []).filter(v => v._checked);
  const vmTargets = checked.map(v => (v.IPs || v.ips || v.Name || v.name || '').toString().split(',')[0].trim()).filter(Boolean);
  targets = Array.from(new Set([...targets, ...vmTargets]));
  data.targets = targets;

  const btn = form.querySelector('button[type="submit"]');
  btn.disabled = true; btn.textContent = 'Starting...';

  try {
    const resp = await fetch('/api/start', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify(data)
    });
    if (!resp.ok) { throw new Error(await resp.text()); }
    setTab('run');
    document.getElementById('tb-setup').disabled = true;
    startPolling();
  } catch(err) {
    alert('Failed to start: ' + err.message);
    btn.disabled = false; btn.textContent = 'Run Discovery';
  }
}

let pollTimer = null;
function startPolling(){
  if (pollTimer) return;
  pollTimer = setInterval(async () => {
    try {
      const resp = await fetch('/api/status');
      const s = await resp.json();
      renderStatus(s);
      if (s.Status === 'complete' || s.Status === 'error') {
        clearInterval(pollTimer); pollTimer = null;
        if (s.Status === 'complete') {
          document.getElementById('tb-report').disabled = false;
          renderReport(s);
        }
      }
    } catch(e) { /* transient network error; keep polling */ }
  }, 1000);
}

function renderStatus(s){
  const pill = document.getElementById('runPill');
  const pillText = document.getElementById('runPillText');
  pill.className = 'pill ' + (s.Status==='running'?'running':s.Status==='complete'?'ok':s.Status==='error'?'err':'idle');
  pillText.textContent = s.Status.charAt(0).toUpperCase()+s.Status.slice(1);
  document.getElementById('runTitle').textContent = s.Client ? 'Discovering ' + s.Client : 'Discovery session';
  document.getElementById('runMeta').textContent = (s.SessionDir || '-');
  const tot = s.Targets.length;
  const done = s.Targets.filter(t=>t.State==='done'||t.State==='error').length;
  const pct = tot ? Math.round(done/tot*100) : 0;
  document.getElementById('overallBar').style.width = pct + '%';
  document.getElementById('progressText').textContent = done + ' / ' + tot + ' targets complete (' + pct + '%)';
  const list = document.getElementById('targetList');
  if (tot === 0) { list.innerHTML = '<div style="color:var(--muted);padding:14px;">No targets.</div>'; }
  else {
    list.innerHTML = s.Targets.map(t => {
      const stateCls = t.State==='done'?'ok':t.State==='error'?'err':t.State==='running'?'running':'idle';
      const stateLabel = t.State.charAt(0).toUpperCase()+t.State.slice(1);
      return `<div class="target-row">
        <span class="buddy">${escapeHtml(t.Buddy || '')}</span>
        <div><div class="target-name">${escapeHtml(t.Name)}</div>
        <div class="target-phase">${escapeHtml(t.Phase || '')}</div></div>
        <span class="pill ${stateCls}"><span class="dot"></span>${stateLabel}</span>
      </div>`;
    }).join('');
  }
  const lb = document.getElementById('logbox');
  const atBottom = lb.scrollTop + lb.clientHeight >= lb.scrollHeight - 20;
  lb.textContent = (s.LogTail || []).join('\n');
  if (atBottom) lb.scrollTop = lb.scrollHeight;
}

function renderReport(s){
  const el = document.getElementById('reportContent');
  const sessionDir = s.SessionDir || '';
  const openFolderBtn = sessionDir
    ? `<button class="btn btn-secondary" onclick="openSessionFolder()" style="margin-top:10px;margin-left:8px;">Open session folder</button>`
    : '';
  const viewLogBtn = sessionDir
    ? `<button class="btn btn-secondary" onclick="viewGenLog()" style="margin-top:10px;margin-left:8px;">View gen_report log</button>`
    : '';
  const zipBtn = s.ReportZipPath
    ? `<a href="/api/download-zip" class="btn" style="margin-top:10px;margin-left:8px;display:inline-block;text-decoration:none;">Download zip</a>`
    : '';
  const copyLogsBtn = `<button class="btn btn-secondary" onclick="copyLogs(this)" style="margin-top:10px;margin-left:8px;">Copy logs</button>`;

  // Missing-JSON audit block (only if anything is missing)
  let missingHtml = '';
  if (Array.isArray(s.MissingTargets) && s.MissingTargets.length) {
    const rows = s.MissingTargets.map(m =>
      `<li><code>${escapeHtml(m.Address)}</code> <span style="color:var(--muted);">(${escapeHtml(m.Kind||'')})</span> - <span style="color:var(--warn);">${escapeHtml(m.Reason||'')}</span></li>`
    ).join('');
    missingHtml = `<div style="margin-top:14px;padding:10px 14px;background:#2a1b0f;border-left:3px solid var(--warn);border-radius:6px;">
      <div style="font-size:11px;color:var(--warn);text-transform:uppercase;letter-spacing:.6px;margin-bottom:6px;">Missing JSON (${s.MissingTargets.length})</div>
      <ul style="margin:0;padding-left:18px;font-size:12.5px;line-height:1.55;">${rows}</ul>
    </div>`;
  }

  if (s.ReportPath) {
    el.innerHTML = `<p style="font-size:13px;">Report saved:</p>
      <p><code style="font-family:var(--mono);color:var(--info);font-size:12px;word-break:break-all;">${escapeHtml(s.ReportPath)}</code></p>
      <p>
        <a href="file:///${escapeHtml(s.ReportPath.replace(/\\/g,'/'))}" target="_blank" class="btn" style="display:inline-block;text-decoration:none;">Open report</a>
        ${zipBtn}
        ${openFolderBtn}
        ${viewLogBtn}
        ${copyLogsBtn}
      </p>
      ${missingHtml}
      <div id="genLogBox"></div>`;
    document.getElementById('reportSub').textContent = 'Session complete. Click below to open the HTML report.';
  } else {
    el.innerHTML = `<p style="color:var(--warn);">Report not generated - check log for errors.</p>
      <p style="color:var(--muted);font-size:12px;">Session dir: <code style="font-family:var(--mono);">${escapeHtml(sessionDir)}</code></p>
      <p>${openFolderBtn}${viewLogBtn}${copyLogsBtn}</p>
      ${missingHtml}
      <div id="genLogBox"></div>`;
  }
}

async function copyLogs(btn){
  const orig = btn.textContent;
  btn.textContent = 'Copying...';
  btn.disabled = true;
  try {
    const resp = await fetch('/api/combined-logs');
    const data = await resp.json();
    const text = data.content || '';
    try {
      await navigator.clipboard.writeText(text);
      btn.textContent = 'Copied!';
    } catch(e) {
      // Fallback for non-HTTPS/localhost contexts
      const ta = document.createElement('textarea');
      ta.value = text; document.body.appendChild(ta); ta.select();
      document.execCommand('copy'); document.body.removeChild(ta);
      btn.textContent = 'Copied!';
    }
    setTimeout(()=>{ btn.textContent = orig; btn.disabled = false; }, 1500);
  } catch(e) {
    btn.textContent = 'Failed';
    setTimeout(()=>{ btn.textContent = orig; btn.disabled = false; }, 1500);
  }
}

async function openSessionFolder(){
  try {
    const resp = await fetch('/api/open-folder', { method: 'POST' });
    if (!resp.ok) { alert('Failed to open folder: ' + await resp.text()); }
  } catch(e) { alert('Error: ' + e.message); }
}

async function viewGenLog(){
  const box = document.getElementById('genLogBox');
  box.innerHTML = '<p style="color:var(--muted);font-size:12px;margin-top:14px;">Loading log...</p>';
  try {
    const resp = await fetch('/api/gen-log');
    const data = await resp.json();
    if (data.content) {
      box.innerHTML = `<div style="margin-top:14px;"><div style="font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.6px;margin-bottom:6px;">gen_report.log</div>
        <pre style="background:#07101f;color:#c7d1df;padding:14px;border-radius:8px;font-family:var(--mono);font-size:11.5px;max-height:420px;overflow:auto;white-space:pre-wrap;">${escapeHtml(data.content)}</pre></div>`;
    } else {
      box.innerHTML = `<p style="color:var(--muted);font-size:12px;margin-top:10px;">No log file: ${escapeHtml(data.error || 'unknown')}</p>`;
    }
  } catch(e) {
    box.innerHTML = `<p style="color:var(--crit);font-size:12px;margin-top:10px;">Error loading log: ${escapeHtml(e.message)}</p>`;
  }
}

function escapeHtml(s){
  return String(s||'').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function togglePw(btn){
  const input = btn.parentElement.querySelector('input');
  if (!input) return;
  input.type = (input.type === 'password') ? 'text' : 'password';
}

async function testCreds(){
  const form = document.getElementById('setupForm');
  const user = (form.winrmUser.value || '').trim();
  const pass = form.winrmPass.value || '';
  const btn  = document.getElementById('testCredsBtn');
  const st   = document.getElementById('credStatus');
  if (!user || !pass) {
    st.style.color = 'var(--warn)';
    st.textContent = 'Enter a username and password first.';
    return;
  }
  btn.disabled = true;
  const orig = btn.textContent;
  btn.textContent = 'Checking...';
  st.style.color = 'var(--muted)';
  st.textContent = 'Contacting domain controller...';
  try {
    const resp = await fetch('/api/test-creds', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ winrmUser: user, winrmPass: pass })
    });
    const data = await resp.json();
    if (data.ok) {
      st.style.color = 'var(--ok)';
      st.textContent = 'OK - ' + (data.message || 'credentials valid');
    } else {
      st.style.color = 'var(--crit)';
      st.textContent = data.error || 'validation failed';
    }
  } catch(e) {
    st.style.color = 'var(--crit)';
    st.textContent = 'Error: ' + e.message;
  } finally {
    btn.textContent = orig;
    btn.disabled = false;
  }
}
</script>
</body></html>
'@

# Inject version into template
$script:HtmlUI = $script:HtmlUI -replace '__VERSION__', $script:Version

# -----------------------------------------------------------------------------
# HTTP LISTENER
# -----------------------------------------------------------------------------
function Start-HttpListener {
    # Defensive: if caller botched the port, snap back to default 8080
    if ($Port -le 0 -or $Port -gt 65535) {
        Write-Host "  [warn] Invalid port $Port - falling back to 8080" -ForegroundColor DarkYellow
        $script:Port = 8080
    }
    # Try requested port, then scan a few adjacent ports if it's busy
    $triedPorts = @()
    foreach ($p in @($Port, 8080, 8081, 8082, 8888, 9090)) {
        if ($triedPorts -contains $p) { continue }
        $triedPorts += $p
        $listener = New-Object System.Net.HttpListener
        $prefix   = "http://localhost:$p/"
        $listener.Prefixes.Add($prefix)
        try {
            $listener.Start()
            $script:Port    = $p
            $script:BaseUrl = "http://localhost:$p"
            Write-Host ""
            Write-Host "  SDT GUI running at: $prefix" -ForegroundColor Green
            Write-Host "  (Press Ctrl+C to stop)" -ForegroundColor DarkGray
            Write-Host ""
            return $listener
        } catch {
            try { $listener.Close() } catch { }
            Write-Host "  [skip] port $p busy ($($_.Exception.Message.Split('.')[0]))" -ForegroundColor DarkGray
        }
    }
    throw "Could not bind any of: $($triedPorts -join ', '). Free one up or use -Port <n>."
}

function Send-Response {
    param(
        [System.Net.HttpListenerResponse] $Response,
        [string] $Body,
        [string] $ContentType = 'text/html; charset=utf-8',
        [int]    $StatusCode  = 200
    )
    $Response.StatusCode  = $StatusCode
    $Response.ContentType = $ContentType
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Body)
    $Response.ContentLength64 = $bytes.Length
    $Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $Response.OutputStream.Close()
}

function Send-Json {
    param([System.Net.HttpListenerResponse] $Response, $Data, [int] $StatusCode = 200)
    $json = $Data | ConvertTo-Json -Depth 10 -Compress
    Send-Response -Response $Response -Body $json -ContentType 'application/json; charset=utf-8' -StatusCode $StatusCode
}

function Read-RequestBody {
    param([System.Net.HttpListenerRequest] $Request)
    if (-not $Request.HasEntityBody) { return '' }
    $reader = New-Object System.IO.StreamReader($Request.InputStream, [System.Text.Encoding]::UTF8)
    try { return $reader.ReadToEnd() } finally { $reader.Close() }
}

# -----------------------------------------------------------------------------
# DISCOVERY WORKER - spawns Invoke-ServerDiscovery.ps1 per target, updates state
# -----------------------------------------------------------------------------
function Start-DiscoveryRun {
    param($Payload)

    $script:Session.Status     = 'running'
    $script:Session.StartedAt  = Get-Date
    $script:Session.Client     = $Payload.client
    $client                    = if ($Payload.client) { ($Payload.client -replace '[^A-Za-z0-9_-]+','_') } else { 'CLIENT' }
    $stamp                     = (Get-Date).ToString('yyyy-MM-dd-HHmm')
    $outRoot                   = if ($Payload.outputDir) { $Payload.outputDir } else { 'C:\Temp\sdt\sessions' }
    $sessionDir                = Join-Path $outRoot ("$client-$stamp")
    New-Item -ItemType Directory -Path $sessionDir -Force | Out-Null
    $script:Session.SessionDir = $sessionDir
    $script:Session.Targets    = @()
    $script:Session.LogTail.Clear()
    $script:Session.ReportPath = ''
    $script:Session.ReportZipPath = ''
    $script:Session.MissingTargets = @()
    Add-Log "Session started. Output: $sessionDir"

    # Initialize target rows
    foreach ($t in $Payload.targets) {
        $script:Session.Targets += [ordered]@{
            Name=$t; Address=$t; State='pending'; Phase=''; Buddy=''; Started=$null; Finished=$null; Kind='server'
        }
    }
    # If hypervisor provided, add a synthetic "Hypervisor" row at the top.
    $hvType = "$($Payload.hvType)"
    $hvHost = "$($Payload.hvHost)"
    if ($hvType -and $hvType -ne 'none' -and $hvHost) {
        $alreadyScanned = $script:Session.HvStagingFile -and (Test-Path $script:Session.HvStagingFile)
        $hvRow = [ordered]@{
            Name="$hvType`: $hvHost"
            Address=$hvHost
            State='pending'
            Phase=$(if ($alreadyScanned) { 'inventory already scanned' } else { 'queued' })
            Buddy=''
            Started=$null
            Finished=$null
            Kind='hypervisor'
            HvType=$hvType
            AlreadyScanned=$alreadyScanned
        }
        # Prepend
        $script:Session.Targets = @($hvRow) + @($script:Session.Targets)
        # If we already scanned, move the staging file into the session dir NOW
        if ($alreadyScanned) {
            try {
                $destName = Split-Path -Leaf $script:Session.HvStagingFile
                $dest = Join-Path $sessionDir $destName
                Move-Item -Path $script:Session.HvStagingFile -Destination $dest -Force
                if ($script:Session.HvStagingDir -and (Test-Path $script:Session.HvStagingDir)) {
                    Remove-Item $script:Session.HvStagingDir -Recurse -Force -EA SilentlyContinue
                }
                $script:Session.HvStagingFile = ''
                $script:Session.HvStagingDir = ''
                Add-Log "Hypervisor inventory moved into session from prior scan: $destName"
            } catch {
                Add-Log "Failed to move staged HV inventory: $($_.Exception.Message)"
            }
        }
    }

    # Kick off the worker scriptblock in a runspace so the HTTP listener stays responsive.
    $job = Start-ThreadJob -ScriptBlock {
        param($Session, $Payload, $ScriptDir, $BuddyFrames)

        $invoke = Join-Path $ScriptDir 'Invoke-ServerDiscovery.ps1'
        if (-not (Test-Path $invoke)) {
            $Session.Status = 'error'
            return
        }

        $winrmUser = $Payload.winrmUser
        $winrmPass = $Payload.winrmPass
        $cred = $null
        if ($winrmUser -and $winrmPass) {
            $sec  = ConvertTo-SecureString $winrmPass -AsPlainText -Force
            $cred = New-Object System.Management.Automation.PSCredential($winrmUser, $sec)
        }

        # Pick Python. Prefer portable Python the SDT ships; auto-fetch if missing
        # (only when a hypervisor target is present, since that's the only thing
        # that needs Python right now).
        $pyExe = Join-Path $ScriptDir 'python\python.exe'
        $needPy = $Session.Targets | Where-Object { $_.Kind -eq 'hypervisor' } | Select-Object -First 1
        if ($needPy -and -not (Test-Path $pyExe)) {
            $getPy = Join-Path $ScriptDir 'Get-PortablePython.ps1'
            if (Test-Path $getPy) {
                try { & $getPy 2>&1 | Out-Null } catch { }
            }
        }
        if (-not (Test-Path $pyExe)) { $pyExe = 'python' }

        for ($i = 0; $i -lt $Session.Targets.Count; $i++) {
            $t = $Session.Targets[$i]
            $t.State   = 'running'
            $t.Started = (Get-Date).ToString('HH:mm:ss')
            $t.Buddy   = $BuddyFrames[(Get-Random -Maximum $BuddyFrames.Count)]
            $t.Phase   = 'connecting...'
            $Session.Targets[$i] = $t

            try {
                if ($t.Kind -eq 'hypervisor') {
                    # If already scanned during Setup, we already moved the file - mark done and skip
                    if ($t.AlreadyScanned) {
                        $t.State = 'done'
                        $t.Phase = 'reused scan from setup'
                        $t.Finished = (Get-Date).ToString('HH:mm:ss')
                        $Session.Targets[$i] = $t
                        continue
                    }
                    # Otherwise do a fresh hypervisor discovery via collect_vsphere_perf.py
                    $hvScript = Join-Path $ScriptDir 'collect_vsphere_perf.py'
                    if (-not (Test-Path $hvScript)) { throw "vSphere collector not found at $hvScript" }
                    $t.Phase = 'connecting to vCenter/ESXi'
                    $Session.Targets[$i] = $t
                    # Auto-retry with username format variants
                    $hvResult = Invoke-VsphereCollect -PyExe $pyExe -ScriptPath $hvScript `
                        -VCenterHost $t.Address -UserRaw $Payload.hvUser -PassRaw $Payload.hvPass -OutputDir $Session.SessionDir
                    if ($hvResult.ok -and $hvResult.user -ne $Payload.hvUser) {
                        $t.Phase = "auth OK as $($hvResult.user)"
                        $Session.Targets[$i] = $t
                    } elseif (-not $hvResult.ok) {
                        $t.State='error'
                        $t.Phase = $hvResult.error
                        $Session.Targets[$i] = $t
                    }
                    # Look for any *-inventory-*.json written into the session dir
                    $written = Get-ChildItem $Session.SessionDir -Filter '*inventory*.json' -EA 0 | Select-Object -First 1
                    if ($written) {
                        $t.State='done'; $t.Phase=("inventory saved: {0}" -f $written.Name)
                    } else {
                        $t.State='error'; $t.Phase='inventory not written (check hypervisor creds / reachability)'
                    }
                } else {
                    # Per-server Windows discovery
                    $safeName = ($t.Address -replace '[^A-Za-z0-9_.-]+','_')
                    $perLog = Join-Path $Session.SessionDir ("server-{0}.log" -f $safeName)
                    $t.LogPath = $perLog
                    $Session.Targets[$i] = $t
                    $lines = New-Object System.Collections.ArrayList

                    # ---- Preflight reachability check ---------------------
                    # Tells us WHY a scan is going to fail (off / linux / firewalled)
                    # before we spend a minute trying WinRM.
                    $t.Phase = 'preflight...'
                    $Session.Targets[$i] = $t
                    $pingOk  = $false
                    $winrmOk = $false
                    $sshOk   = $false
                    try { $pingOk = Test-Connection -ComputerName $t.Address -Count 1 -Quiet -EA SilentlyContinue } catch {}
                    try {
                        $tcp = New-Object System.Net.Sockets.TcpClient
                        $ia  = $tcp.BeginConnect($t.Address, 5985, $null, $null)
                        if ($ia.AsyncWaitHandle.WaitOne(1500)) { $winrmOk = $tcp.Connected; $tcp.EndConnect($ia) }
                        $tcp.Close()
                    } catch {}
                    if (-not $winrmOk) {
                        try {
                            $tcp = New-Object System.Net.Sockets.TcpClient
                            $ia  = $tcp.BeginConnect($t.Address, 5986, $null, $null)
                            if ($ia.AsyncWaitHandle.WaitOne(1500)) { $winrmOk = $tcp.Connected; $tcp.EndConnect($ia) }
                            $tcp.Close()
                        } catch {}
                    }
                    try {
                        $tcp = New-Object System.Net.Sockets.TcpClient
                        $ia  = $tcp.BeginConnect($t.Address, 22, $null, $null)
                        if ($ia.AsyncWaitHandle.WaitOne(1500)) { $sshOk = $tcp.Connected; $tcp.EndConnect($ia) }
                        $tcp.Close()
                    } catch {}
                    [void]$lines.Add("[preflight] ping=$pingOk winrm=$winrmOk ssh=$sshOk")

                    if (-not $winrmOk) {
                        if (-not $pingOk -and -not $sshOk) {
                            $t.State='error'; $t.Phase='unreachable (powered off / wrong IP / firewalled)'
                        } elseif ($sshOk -and -not $winrmOk) {
                            $t.State='error'; $t.Phase='looks like Linux (SSH only) - use Linux workflow'
                        } elseif ($pingOk -and -not $winrmOk -and -not $sshOk) {
                            $t.State='error'; $t.Phase='up but WinRM/SSH both closed (Enable-PSRemoting or firewall?)'
                        } else {
                            $t.State='error'; $t.Phase='not Windows or WinRM not enabled'
                        }
                        $t.Finished = (Get-Date).ToString('HH:mm:ss')
                        $Session.Targets[$i] = $t
                        try { [System.IO.File]::WriteAllText($perLog, ($lines -join "`r`n"), [System.Text.Encoding]::UTF8) } catch {}
                        $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] $($t.Address): SKIPPED - $($t.Phase)") | Out-Null
                        continue
                    }

                    # ---- Actual WinRM discovery ---------------------------
                    $invokeArgs = @{ ComputerName = $t.Address; OutputPath = $Session.SessionDir }
                    if ($cred) { $invokeArgs.Credential = $cred }
                    # Snapshot existing discovery JSONs so we can detect what was newly created
                    $preJson = @(Get-ChildItem $Session.SessionDir -Filter '*-discovery-*.json' -EA 0 | Select-Object -ExpandProperty FullName)
                    & $invoke @invokeArgs *>&1 | ForEach-Object {
                        $line = "$_"
                        [void]$lines.Add($line)
                        if ($line -match '^\s*\[(\w+)\]') {
                            $t.Phase = $matches[1]
                            $t.Buddy = $BuddyFrames[(Get-Random -Maximum $BuddyFrames.Count)]
                            $Session.Targets[$i] = $t
                        }
                    }
                    try { [System.IO.File]::WriteAllText($perLog, ($lines -join "`r`n"), [System.Text.Encoding]::UTF8) } catch {}
                    # Verify: did Invoke-ServerDiscovery produce a new discovery JSON?
                    $postJson = @(Get-ChildItem $Session.SessionDir -Filter '*-discovery-*.json' -EA 0 | Select-Object -ExpandProperty FullName)
                    $newJson  = $postJson | Where-Object { $preJson -notcontains $_ }
                    if ($newJson) {
                        $jsonFile = $newJson | Select-Object -First 1
                        $t.JsonPath = $jsonFile
                        $t.State    = 'done'
                        $t.Phase    = "json: $(Split-Path -Leaf $jsonFile)"
                        $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] $($t.Address): JSON written - $(Split-Path -Leaf $jsonFile)") | Out-Null
                    } else {
                        # Sniff the per-server log for a probable cause
                        $tail = ($lines | Select-Object -Last 20) -join ' | '
                        $reason = 'no discovery JSON written'
                        if     ($tail -match '(?i)Access is denied')           { $reason = 'WinRM auth denied (check creds/domain)' }
                        elseif ($tail -match '(?i)cannot find')                { $reason = 'host unreachable or name not resolvable' }
                        elseif ($tail -match '(?i)WinRM|5985|5986')            { $reason = 'WinRM not reachable (port 5985/5986)' }
                        elseif ($tail -match '(?i)timed? ?out')                { $reason = 'connection timed out' }
                        elseif ($tail -match '(?i)ConvertTo-Json')             { $reason = 'data collected but JSON serialize failed - see per-server log' }
                        $t.State    = 'error'
                        $t.Phase    = $reason
                        $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] $($t.Address): NO JSON - $reason (see $(Split-Path -Leaf $perLog))") | Out-Null
                    }
                }
                $t.Finished = (Get-Date).ToString('HH:mm:ss')
            } catch {
                $t.State    = 'error'
                $t.Phase    = $_.Exception.Message
                $t.Finished = (Get-Date).ToString('HH:mm:ss')
            }
            $Session.Targets[$i] = $t
        }

        # ---------------------------------------------------------------
        # Post-session verification: confirm each target produced an
        # expected JSON. Log a compact audit so the SE sees what landed
        # and what didn't without having to dig through the folder.
        # ---------------------------------------------------------------
        $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] --- Post-session JSON audit ---") | Out-Null
        $missing = @()
        for ($i = 0; $i -lt $Session.Targets.Count; $i++) {
            $t = $Session.Targets[$i]
            if ($t.Kind -eq 'hypervisor') {
                $inv = Get-ChildItem $Session.SessionDir -Filter '*inventory*.json' -EA 0 | Select-Object -First 1
                if ($inv) {
                    $Session.LogTail.Add("   [OK] $($t.Address) (hypervisor) -> $($inv.Name)") | Out-Null
                } else {
                    $Session.LogTail.Add("   [MISS] $($t.Address) (hypervisor): $($t.Phase)") | Out-Null
                    $missing += $t
                }
            } else {
                if ($t.JsonPath -and (Test-Path $t.JsonPath)) {
                    $Session.LogTail.Add("   [OK] $($t.Address) -> $(Split-Path -Leaf $t.JsonPath)") | Out-Null
                } else {
                    $Session.LogTail.Add("   [MISS] $($t.Address): $($t.Phase)") | Out-Null
                    $missing += $t
                }
            }
        }
        if ($missing.Count -eq 0) {
            $Session.LogTail.Add("   All $($Session.Targets.Count) target(s) produced JSON. Proceeding to report.") | Out-Null
        } else {
            $Session.LogTail.Add("   $($missing.Count) of $($Session.Targets.Count) target(s) missing JSON. Report will include only what was collected.") | Out-Null
            $Session.MissingTargets = @($missing | ForEach-Object { @{ Address=$_.Address; Reason=$_.Phase; Kind=$_.Kind } })
        }

        # Attempt to generate the HTML report. Capture stdout+stderr to gen_report.log
        # in the session dir so failures are diagnosable without restarting.
        $gen = Join-Path $ScriptDir 'gen_report.py'
        $py  = Join-Path $ScriptDir 'python\python.exe'
        if (-not (Test-Path $py)) { $py = 'python' }
        $reportLog = Join-Path $Session.SessionDir 'gen_report.log'
        try {
            $mf = Join-Path $Session.SessionDir 'manifest.json'
            $manifest = @{
                client      = $Payload.client
                client_full = $Payload.client
                date        = (Get-Date).ToString('yyyy-MM-dd')
                session_dir = '.'
                output_dir  = '.'
                inventory_file = ''
                logo_file   = ''
            } | ConvertTo-Json -Depth 6
            [System.IO.File]::WriteAllText($mf, $manifest, [System.Text.Encoding]::UTF8)
            # Run + capture everything
            $logContent = & $py $gen $mf 2>&1 | Out-String
            [System.IO.File]::WriteAllText($reportLog, $logContent, [System.Text.Encoding]::UTF8)
            # Find the resulting HTML
            $html = Get-ChildItem $Session.SessionDir -Filter '*DiscoveryReport*.html' -EA 0 | Select-Object -First 1
            if ($html) {
                $Session.ReportPath = $html.FullName
            } else {
                # Leave a breadcrumb in the session log so the UI shows the issue
                $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] gen_report.py produced no HTML - see gen_report.log in session dir") | Out-Null
            }
        } catch {
            $errMsg = $_.Exception.Message
            [System.IO.File]::WriteAllText($reportLog, "Exception invoking gen_report.py: $errMsg", [System.Text.Encoding]::UTF8)
            $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] gen_report.py exception: $errMsg") | Out-Null
        }

        # If the report generated successfully, bundle the whole session into
        # a zip for easy handoff. Sits next to the session folder.
        if ($Session.ReportPath) {
            try {
                $clientSlug = ($Payload.client -replace '[^A-Za-z0-9_-]+','_')
                $stampNow = (Get-Date).ToString('yyyy-MM-dd-HHmm')
                $zipName = "$clientSlug-sdt-$stampNow.zip"
                $zipPath = Join-Path (Split-Path $Session.SessionDir -Parent) $zipName
                # Compress everything in the session dir
                Compress-Archive -Path (Join-Path $Session.SessionDir '*') -DestinationPath $zipPath -Force
                if (Test-Path $zipPath) {
                    $Session.ReportZipPath = $zipPath
                    $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] Session bundled: $zipPath") | Out-Null
                }
            } catch {
                $Session.LogTail.Add("[$(Get-Date -f 'HH:mm:ss')] Zip bundle failed: $($_.Exception.Message)") | Out-Null
            }
        }

        $Session.Status     = 'complete'
        $Session.FinishedAt = Get-Date

    } -ArgumentList $script:Session, $Payload, $script:ScriptDir, $script:BuddyFrames

    return $job
}

# -----------------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------------

# Ensure ThreadJob is available (ships with PS 5.1 since WMF 5.1; may need install)
if (-not (Get-Command Start-ThreadJob -EA 0)) {
    try {
        Import-Module ThreadJob -EA Stop
    } catch {
        Write-Host "  [warn] ThreadJob module not available. Installing..." -ForegroundColor Yellow
        Install-Module ThreadJob -Scope CurrentUser -Force -AllowClobber -EA Stop
        Import-Module ThreadJob -EA Stop
    }
}

# Auto-minimize the host PowerShell window so the GUI takes focus.
try {
    Add-Type -Name _SdtWin -Namespace _Sdt -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool ShowWindow(System.IntPtr h, int s);
[System.Runtime.InteropServices.DllImport("kernel32.dll")]
public static extern System.IntPtr GetConsoleWindow();
'@ -ErrorAction Stop
    $h = [_Sdt._SdtWin]::GetConsoleWindow()
    if ($h -ne [System.IntPtr]::Zero) {
        [void][_Sdt._SdtWin]::ShowWindow($h, 6)  # SW_MINIMIZE
    }
} catch { }

$listener = Start-HttpListener

# Open browser - prefer Edge, then Chrome, fall back to default.
# Use --app mode for a clean window without browser chrome.
function Find-ModernBrowser {
    $candidates = @(
        # Edge (Chromium) - standard install locations
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe",
        # Chrome - standard install locations
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "$env:ProgramFiles(x86)\Google\Chrome\Application\chrome.exe",
        "$env:LocalAppData\Google\Chrome\Application\chrome.exe"
    )
    foreach ($p in $candidates) {
        if ($p -and (Test-Path $p)) { return $p }
    }
    # Registry fallback for App Paths
    foreach ($reg in @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe'
    )) {
        try {
            $v = (Get-ItemProperty -Path $reg -Name '(Default)' -EA Stop).'(Default)'
            if ($v -and (Test-Path $v)) { return $v }
        } catch { }
    }
    return $null
}

if (-not $NoOpenBrowser) {
    $browser = Find-ModernBrowser
    try {
        if ($browser) {
            Write-Host "  Launching: $browser" -ForegroundColor DarkGray
            # --app mode strips tab bar/URL bar for a native-feel window
            Start-Process -FilePath $browser -ArgumentList "--app=$script:BaseUrl" | Out-Null
        } else {
            Write-Host "  No Edge/Chrome found; using system default browser." -ForegroundColor DarkYellow
            Start-Process $script:BaseUrl | Out-Null
        }
    } catch {
        # Last-resort fallback
        try { Start-Process $script:BaseUrl | Out-Null } catch { }
    }
}

# Request loop
try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $req     = $context.Request
        $resp    = $context.Response
        $path    = $req.Url.AbsolutePath
        $method  = $req.HttpMethod

        try {
            switch -Regex ("$method $path") {
                '^GET /api/status$' {
                    Send-Json -Response $resp -Data @{
                        Status     = $script:Session.Status
                        Client     = $script:Session.Client
                        SessionDir = $script:Session.SessionDir
                        Targets    = @($script:Session.Targets)
                        LogTail    = @($script:Session.LogTail)
                        ReportPath = $script:Session.ReportPath
                        ReportZipPath = $script:Session.ReportZipPath
                        MissingTargets = @($script:Session.MissingTargets)
                    }
                    break
                }
                '^POST /api/hv-scan$' {
                    $body = Read-RequestBody -Request $req
                    $hv = $null
                    try { $hv = $body | ConvertFrom-Json -ErrorAction Stop } catch {
                        Send-Json -Response $resp -Data @{ ok=$false; error="JSON parse failed" } -StatusCode 400
                        break
                    }
                    if (-not $hv.hvHost -or -not $hv.hvUser -or -not $hv.hvPass) {
                        Send-Json -Response $resp -Data @{ ok=$false; error="hvHost, hvUser, hvPass required" } -StatusCode 400
                        break
                    }
                    # Scan into a temp staging dir so we don't commit to a session folder yet
                    $stageDir = Join-Path $env:TEMP ("sdt-hvscan-" + [guid]::NewGuid().ToString('N').Substring(0,8))
                    New-Item -ItemType Directory -Path $stageDir -Force | Out-Null
                    try {
                        $hvScript = Join-Path $script:ScriptDir 'collect_vsphere_perf.py'
                        if (-not (Test-Path $hvScript)) { throw "collect_vsphere_perf.py not found" }
                        $pyExe = Join-Path $script:ScriptDir 'python\python.exe'
                        if (-not (Test-Path $pyExe)) {
                            $getPy = Join-Path $script:ScriptDir 'Get-PortablePython.ps1'
                            if (Test-Path $getPy) { try { & $getPy 2>&1 | Out-Null } catch {} }
                        }
                        if (-not (Test-Path $pyExe)) { $pyExe = 'python' }
                        # Invoke collector with auto-retry across username formats
                        $collectResult = Invoke-VsphereCollect -PyExe $pyExe -ScriptPath $hvScript `
                            -VCenterHost $hv.hvHost -UserRaw $hv.hvUser -PassRaw $hv.hvPass -OutputDir $stageDir

                        if (-not $collectResult.ok) {
                            $errMsg = "Collector failed: $($collectResult.error)"
                            if ($collectResult.triedUsers) {
                                $errMsg += ". Tried usernames: $($collectResult.triedUsers -join ', ')"
                            }
                            Send-Json -Response $resp -Data @{ ok=$false; error=$errMsg; log=$collectResult.log } -StatusCode 500
                            break
                        }
                        $outFile = Get-Item $collectResult.file
                        if ($collectResult.user -ne $hv.hvUser) {
                            Add-Log ("Hypervisor auth succeeded with adjusted username: {0}" -f $collectResult.user)
                        }
                        $raw = [System.IO.File]::ReadAllText($outFile.FullName, [System.Text.Encoding]::UTF8)
                        $doc = $raw | ConvertFrom-Json
                        # Remember staging so /api/start can move it into the final session dir
                        $script:Session.HvStagingFile = $outFile.FullName
                        $script:Session.HvStagingDir  = $stageDir
                        # Shape VMs for UI
                        $vms = @()
                        foreach ($v in $doc.VMs) {
                            $vms += @{
                                Name       = $v.Name
                                IPs        = $v.IPs
                                GuestOS    = $v.GuestOS
                                PowerState = $v.PowerState
                            }
                        }
                        Send-Json -Response $resp -Data @{ ok=$true; vms=$vms; count=$vms.Count; stagingFile=$outFile.FullName }
                    } catch {
                        Send-Json -Response $resp -Data @{ ok=$false; error=$_.Exception.Message } -StatusCode 500
                    }
                    break
                }
                '^POST /api/start$' {
                    if ($script:Session.Status -eq 'running') {
                        Send-Json -Response $resp -Data @{ error = 'A discovery session is already running.' } -StatusCode 409
                        break
                    }
                    $body = Read-RequestBody -Request $req
                    $payload = $null
                    try { $payload = $body | ConvertFrom-Json -ErrorAction Stop } catch {
                        Send-Json -Response $resp -Data @{ error = "JSON parse failed: $($_.Exception.Message)"; bodyLen = $body.Length } -StatusCode 400
                        break
                    }
                    # Normalize targets: PS ConvertFrom-Json may return $null, string,
                    # or Object[]. Force into an array of non-empty strings.
                    $raw = $null
                    if ($payload) {
                        try { $raw = $payload.targets } catch { $raw = $null }
                    }
                    $targets = @()
                    if ($raw -is [string] -and $raw.Trim()) {
                        $targets = @($raw -split '\r?\n' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                    } elseif ($raw) {
                        $targets = @(@($raw) | ForEach-Object { "$_".Trim() } | Where-Object { $_ })
                    }
                    # Blank targets OK if a hypervisor is set - we'll discover VMs from it.
                    $hvType = ''
                    try { $hvType = "$($payload.hvType)" } catch { $hvType = '' }
                    $hvHost = ''
                    try { $hvHost = "$($payload.hvHost)" } catch { $hvHost = '' }
                    if ($targets.Count -eq 0 -and ($hvType -eq '' -or $hvType -eq 'none' -or -not $hvHost)) {
                        Send-Json -Response $resp -Data @{
                            error = 'No targets and no hypervisor. Provide at least one target host OR hypervisor details.'
                        } -StatusCode 400
                        break
                    }
                    # Replace parsed targets back into payload for worker
                    $payload | Add-Member -NotePropertyName targets -NotePropertyValue $targets -Force
                    $null = Start-DiscoveryRun -Payload $payload
                    Send-Json -Response $resp -Data @{ ok = $true; sessionDir = $script:Session.SessionDir; targetCount = $targets.Count }
                    break
                }
                '^POST /api/test-creds$' {
                    $body = Read-RequestBody -Request $req
                    $c = $null
                    try { $c = $body | ConvertFrom-Json -ErrorAction Stop } catch {
                        Send-Json -Response $resp -Data @{ ok=$false; error='JSON parse failed' } -StatusCode 400; break
                    }
                    $u = "$($c.winrmUser)".Trim()
                    $p = "$($c.winrmPass)"
                    if (-not $u -or -not $p) {
                        Send-Json -Response $resp -Data @{ ok=$false; error='username + password required' } -StatusCode 400; break
                    }
                    $dom = $null; $justUser = $null
                    if ($u -match '^\.\\(.+)$') {
                        Send-Json -Response $resp -Data @{ ok=$false; error='Local account (.\user) cannot be validated without a target. Will be tested during discovery.' }; break
                    } elseif ($u -match '^([^\\]+)\\(.+)$') {
                        $dom = $matches[1]; $justUser = $matches[2]
                    } elseif ($u -like '*@*') {
                        $dom = ($u -split '@',2)[1]; $justUser = $u
                    } else {
                        Send-Json -Response $resp -Data @{ ok=$false; error='Username needs a domain: DOMAIN\user or user@domain.local' }; break
                    }
                    try {
                        Add-Type -AssemblyName System.DirectoryServices.AccountManagement -ErrorAction Stop
                        $ctx = New-Object System.DirectoryServices.AccountManagement.PrincipalContext(
                            [System.DirectoryServices.AccountManagement.ContextType]::Domain, $dom)
                        $valid = $ctx.ValidateCredentials($justUser, $p)
                        $ctx.Dispose()
                        if ($valid) {
                            Send-Json -Response $resp -Data @{ ok=$true; message="Valid on domain $dom" }
                        } else {
                            Send-Json -Response $resp -Data @{ ok=$false; error="Rejected by $dom (wrong password, expired, or locked)" }
                        }
                    } catch {
                        $msg = $_.Exception.Message
                        if ($msg -match '(?i)server.+not.+found|could not contact|server is not operational') {
                            Send-Json -Response $resp -Data @{ ok=$false; error="Cannot reach domain '$dom' - DC unreachable from this host (VPN, DNS, or workgroup machine?)" }
                        } else {
                            Send-Json -Response $resp -Data @{ ok=$false; error="Validation error: $msg" }
                        }
                    }
                    break
                }
                '^POST /api/open-folder$' {
                    $dir = $script:Session.SessionDir
                    if (-not $dir -or -not (Test-Path $dir)) {
                        Send-Json -Response $resp -Data @{ error = "Session folder does not exist: $dir" } -StatusCode 404
                        break
                    }
                    try {
                        Start-Process explorer.exe $dir | Out-Null
                        Send-Json -Response $resp -Data @{ ok = $true; path = $dir }
                    } catch {
                        Send-Json -Response $resp -Data @{ error = $_.Exception.Message } -StatusCode 500
                    }
                    break
                }
                '^GET /api/gen-log$' {
                    $dir = $script:Session.SessionDir
                    if (-not $dir) {
                        Send-Json -Response $resp -Data @{ error = 'no session dir' } -StatusCode 404
                        break
                    }
                    $log = Join-Path $dir 'gen_report.log'
                    if (-not (Test-Path $log)) {
                        Send-Json -Response $resp -Data @{ error = "gen_report.log not found in $dir" } -StatusCode 404
                        break
                    }
                    try {
                        $content = [System.IO.File]::ReadAllText($log, [System.Text.Encoding]::UTF8)
                        Send-Json -Response $resp -Data @{ ok = $true; content = $content; path = $log }
                    } catch {
                        Send-Json -Response $resp -Data @{ error = $_.Exception.Message } -StatusCode 500
                    }
                    break
                }
                '^GET /api/combined-logs$' {
                    $dir = $script:Session.SessionDir
                    $sb  = New-Object System.Text.StringBuilder
                    [void]$sb.AppendLine("===== SDT Combined Logs =====")
                    [void]$sb.AppendLine("Client      : $($script:Session.Client)")
                    [void]$sb.AppendLine("SessionDir  : $dir")
                    [void]$sb.AppendLine("Status      : $($script:Session.Status)")
                    [void]$sb.AppendLine("ReportPath  : $($script:Session.ReportPath)")
                    [void]$sb.AppendLine("ReportZip   : $($script:Session.ReportZipPath)")
                    [void]$sb.AppendLine("Targets     :")
                    foreach ($t in $script:Session.Targets) {
                        $js = if ($t.JsonPath) { Split-Path -Leaf $t.JsonPath } else { '<none>' }
                        [void]$sb.AppendLine(("   {0,-30} state={1,-7} phase={2} json={3}" -f $t.Address, $t.State, $t.Phase, $js))
                    }
                    if ($script:Session.MissingTargets -and $script:Session.MissingTargets.Count -gt 0) {
                        [void]$sb.AppendLine("Missing JSONs:")
                        foreach ($m in $script:Session.MissingTargets) {
                            [void]$sb.AppendLine("   $($m.Address) ($($m.Kind)) - $($m.Reason)")
                        }
                    }
                    [void]$sb.AppendLine("")
                    [void]$sb.AppendLine("----- Session Log -----")
                    foreach ($ln in $script:Session.LogTail) { [void]$sb.AppendLine($ln) }
                    if ($dir -and (Test-Path $dir)) {
                        $gen = Join-Path $dir 'gen_report.log'
                        if (Test-Path $gen) {
                            [void]$sb.AppendLine("")
                            [void]$sb.AppendLine("----- gen_report.log -----")
                            [void]$sb.AppendLine([System.IO.File]::ReadAllText($gen, [System.Text.Encoding]::UTF8))
                        }
                        Get-ChildItem $dir -Filter 'server-*.log' -EA 0 | Sort-Object Name | ForEach-Object {
                            [void]$sb.AppendLine("")
                            [void]$sb.AppendLine("----- $($_.Name) -----")
                            try { [void]$sb.AppendLine([System.IO.File]::ReadAllText($_.FullName, [System.Text.Encoding]::UTF8)) } catch {}
                        }
                    }
                    Send-Json -Response $resp -Data @{ ok = $true; content = $sb.ToString() }
                    break
                }
                '^GET /api/download-zip$' {
                    $zip = $script:Session.ReportZipPath
                    if (-not $zip -or -not (Test-Path $zip)) {
                        Send-Response -Response $resp -Body 'Zip not available yet.' -StatusCode 404 -ContentType 'text/plain'
                        break
                    }
                    try {
                        $bytes = [System.IO.File]::ReadAllBytes($zip)
                        $fname = [System.IO.Path]::GetFileName($zip)
                        $resp.StatusCode = 200
                        $resp.ContentType = 'application/zip'
                        $resp.Headers.Add('Content-Disposition', "attachment; filename=""$fname""")
                        $resp.ContentLength64 = $bytes.Length
                        $resp.OutputStream.Write($bytes, 0, $bytes.Length)
                        $resp.OutputStream.Close()
                    } catch {
                        try { Send-Response -Response $resp -Body "Zip stream error: $($_.Exception.Message)" -StatusCode 500 -ContentType 'text/plain' } catch {}
                    }
                    break
                }
                '^GET /$|^GET /index\.html$' {
                    Send-Response -Response $resp -Body $script:HtmlUI
                    break
                }
                default {
                    Send-Response -Response $resp -Body '404' -StatusCode 404 -ContentType 'text/plain'
                }
            }
        } catch {
            try { Send-Response -Response $resp -Body "Server error: $($_.Exception.Message)" -StatusCode 500 -ContentType 'text/plain' } catch { }
        }
    }
} finally {
    if ($listener) { $listener.Stop(); $listener.Close() }
}
