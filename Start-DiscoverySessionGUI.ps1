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
$script:Version   = '4.0-alpha'
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
.banner-run{background:linear-gradient(125deg,rgba(79,140,255,0.18) 0%,rgba(139,92,246,0.14) 100%);
  border:1px solid var(--border-2);border-radius:14px;padding:20px 24px;margin-bottom:18px;
  display:flex;justify-content:space-between;align-items:center;}
.banner-run .title{font-size:17px;font-weight:700;}
.banner-run .meta{color:var(--muted);font-size:12px;margin-top:3px;}
</style></head><body>
<div class="hdr">
  <div style="display:flex;align-items:center;gap:14px;">
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
<div class="card-title">Session</div>
<div class="card-sub">Client name and output folder. Credentials live in memory only - never saved to disk.</div>
<div class="grid2">
<div class="field"><label>Client name</label>
<input name="client" required placeholder="Acme Corporation"></div>
<div class="field"><label>Output folder</label>
<input name="outputDir" value="C:\Temp\sdt\sessions" required></div>
</div>
</div>

<div class="card">
<div class="card-title">Hypervisor (optional)</div>
<div class="card-sub">Connect to vCenter or an ESXi host to pull VM inventory. Leave "None" if targeting bare metal only.</div>
<div class="grid3">
<div class="field"><label>Type</label>
<select name="hvType"><option value="none">None</option><option value="vsphere">vCenter / ESXi</option><option value="hyperv">Hyper-V Host</option></select></div>
<div class="field"><label>IP / FQDN</label>
<input name="hvHost" placeholder="192.168.10.75"></div>
<div class="field"><label>User</label>
<input name="hvUser" placeholder="administrator@vsphere.local"></div>
</div>
<div class="field"><label>Password</label>
<input name="hvPass" type="password" autocomplete="off"></div>
</div>

<div class="card">
<div class="card-title">Windows Targets</div>
<div class="card-sub">One host per line. IP or hostname. Script will try WinRM remoting; fall back to manual local run if unreachable.
<br>Example: <code style="font-family:var(--mono);color:var(--info);">192.168.10.4</code> or <code style="font-family:var(--mono);color:var(--info);">QES-OFFICE-DC</code></div>
<div class="field"><label>Targets</label>
<textarea name="targets" placeholder="192.168.10.4&#10;192.168.10.5&#10;QES-OFFICE-MGMT"></textarea></div>
<div class="grid2">
<div class="field"><label>WinRM user</label>
<input name="winrmUser" placeholder="DOMAIN\administrator"></div>
<div class="field"><label>WinRM password</label>
<input name="winrmPass" type="password" autocomplete="off"></div>
</div>
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

<div class="footer">Magna5 Solutions Engineering - SDT GUI v__VERSION__</div>
</div>

<script>
function setTab(t){
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.toggle('active', b.id==='tb-'+t));
  document.querySelectorAll('.tab-pane').forEach(p=>p.classList.toggle('active', p.id==='tab-'+t));
}

async function submitSetup(e){
  e.preventDefault();
  const form = document.getElementById('setupForm');
  const data = Object.fromEntries(new FormData(form).entries());
  data.targets = (data.targets||'').split(/\r?\n/).map(s=>s.trim()).filter(Boolean);

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
  if (s.ReportPath) {
    el.innerHTML = `<p style="font-size:13px;">Report saved:</p>
      <p><code style="font-family:var(--mono);color:var(--info);font-size:12px;word-break:break-all;">${escapeHtml(s.ReportPath)}</code></p>
      <p><a href="file:///${escapeHtml(s.ReportPath.replace(/\\/g,'/'))}" target="_blank" class="btn" style="display:inline-block;text-decoration:none;margin-top:10px;">Open report</a></p>`;
    document.getElementById('reportSub').textContent = 'Session complete. Click below to open the HTML report.';
  } else {
    el.innerHTML = '<p style="color:var(--muted);">Report not generated - check log for errors.</p>';
  }
}

function escapeHtml(s){
  return String(s||'').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
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
    $listener = New-Object System.Net.HttpListener
    $prefix   = "http://localhost:$Port/"
    $listener.Prefixes.Add($prefix)
    try { $listener.Start() } catch {
        Write-Host "  [error] Failed to bind $prefix - $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Try a different port with -Port <n>." -ForegroundColor DarkGray
        throw
    }
    Write-Host ""
    Write-Host "  SDT GUI running at: $prefix" -ForegroundColor Green
    Write-Host "  (Press Ctrl+C to stop)" -ForegroundColor DarkGray
    Write-Host ""
    return $listener
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
    Add-Log "Session started. Output: $sessionDir"

    # Initialize target rows
    foreach ($t in $Payload.targets) {
        $script:Session.Targets += [ordered]@{
            Name=$t; Address=$t; State='pending'; Phase=''; Buddy=''; Started=$null; Finished=$null
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

        for ($i = 0; $i -lt $Session.Targets.Count; $i++) {
            $t = $Session.Targets[$i]
            $t.State   = 'running'
            $t.Started = (Get-Date).ToString('HH:mm:ss')
            $t.Buddy   = $BuddyFrames[(Get-Random -Maximum $BuddyFrames.Count)]
            $t.Phase   = 'connecting...'
            $Session.Targets[$i] = $t

            try {
                $args = @{ ComputerName = $t.Address; OutputPath = $Session.SessionDir }
                if ($cred) { $args.Credential = $cred }
                # Run the script; its output + errors land in the job's log.
                & $invoke @args *>&1 | ForEach-Object {
                    $line = "$_"
                    # Cheap phase extractor: look for "[Phase] ..." pattern.
                    if ($line -match '^\s*\[(\w+)\]') {
                        $t.Phase = $matches[1]
                        $t.Buddy = $BuddyFrames[(Get-Random -Maximum $BuddyFrames.Count)]
                        $Session.Targets[$i] = $t
                    }
                }
                $t.State    = 'done'
                $t.Phase    = 'complete'
                $t.Finished = (Get-Date).ToString('HH:mm:ss')
            } catch {
                $t.State    = 'error'
                $t.Phase    = $_.Exception.Message
                $t.Finished = (Get-Date).ToString('HH:mm:ss')
            }
            $Session.Targets[$i] = $t
        }

        # Attempt to generate the HTML report
        $gen = Join-Path $ScriptDir 'gen_report.py'
        $py  = Join-Path $ScriptDir 'python\python.exe'
        if (-not (Test-Path $py)) { $py = 'python' }
        try {
            # Build a minimal manifest - gen_report.py expects CFG with client/date/session_dir etc.
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
            & $py $gen $mf 2>&1 | Out-Null
            # Find the resulting HTML
            $html = Get-ChildItem $Session.SessionDir -Filter '*DiscoveryReport*.html' -EA 0 | Select-Object -First 1
            if ($html) { $Session.ReportPath = $html.FullName }
        } catch { }

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

$listener = Start-HttpListener

# Open browser
if (-not $NoOpenBrowser) {
    try { Start-Process $script:BaseUrl | Out-Null } catch { }
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
                    }
                    break
                }
                '^POST /api/start$' {
                    if ($script:Session.Status -eq 'running') {
                        Send-Json -Response $resp -Data @{ error = 'A discovery session is already running.' } -StatusCode 409
                        break
                    }
                    $body = Read-RequestBody -Request $req
                    $payload = $body | ConvertFrom-Json
                    if (-not $payload.targets -or $payload.targets.Count -eq 0) {
                        Send-Json -Response $resp -Data @{ error = 'No targets provided.' } -StatusCode 400
                        break
                    }
                    $null = Start-DiscoveryRun -Payload $payload
                    Send-Json -Response $resp -Data @{ ok = $true; sessionDir = $script:Session.SessionDir }
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
