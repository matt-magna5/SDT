# SDT v4.2 Backlog

Captured during the 2026-05-19/20 FBA assessment + HTML report redesign discussion.

## GUI flow (Start-DiscoverySessionGUI.ps1)
- [ ] Setup tab credentials should NOT remain visible at top of Run tab — hide
      Setup pane fully once the run starts
- [ ] On scan complete, auto-jump to **Results tab** (currently the Report tab
      stays cluttered with logs and missing-JSON diagnostics)
- [ ] Results tab content should be empty except:
      - "Scan complete" heading
      - Big button: Open HTML Report (already wired to /api/report-html)
      - Big button: Open session folder (already wired to /api/open-folder)
      - Hide all diagnostic / log / copy-logs / zip controls behind a
        small "Details" disclosure
- [ ] Setup + Run tabs remain navigable from the top nav after completion
      (user expectation: tabs preserve last state, only the default landing
      changes to Results)

## HTML report redesign (gen_report.py)
- [ ] Replace per-server top tabs with **left sidebar nav** (Mockup A style
      blue rail with search box). Right pane scrolls all sections inline.
- [ ] Top horizontal nav tabs become the env-wide views:
      **Overview · Servers · Hypervisor · Active Directory · SQL · Private
      Cloud Sizing · Commvault Sizing · EOL Risks**
- [ ] Overview cards: pull vCenter stats (version, cluster capacity,
      datastore totals), guest-side sizing rollup, AD topology, detected
      tools (EDR/RMM/Backup/Remote with which servers have what), top
      servers by severity, top SQL DBs by size, env-wide findings
- [ ] AD tab: full per-domain breakdown (FSMO roles, DC count,
      users/computers/OUs, stale users + stale computers collapsible)
- [ ] Hypervisor tab: deep ESX/vCenter — all hosts (not just first),
      datastores, full VM list sorted by vCPU, licenses
- [ ] EOL tab: color-code rows yellow/red by severity
- [ ] Strip SMB 1.0 noise from per-server findings (roll up to env-wide
      single line)
- [ ] Group near-universal findings (Defender baseline, no third-party EDR
      everywhere, etc.) as ENV-wide rollup not per-server spam

## Private Cloud Sizing tab
- [ ] Read 95th-percentile CPU/RAM from any `vsphere-perf-*.json` in
      session dir (either SDT's collect_vsphere_perf.py output or Nutanix
      Collector XLSX parsed via parse_ntnx_collector.py — same schema)
- [ ] Calculate sizing: `P95 utilization × allocated × 1.20 buffer`
- [ ] Per-VM and total cluster sizing
- [ ] If P95 values are null/missing, show banner: "drop the Nutanix
      Collector XLSX into the session folder and refresh"

## Commvault Sizing tab
- [ ] Use the 1:1 front-end match + 20% growth rule (per M5 BDR team)
- [ ] Pull source disk consumed from SDT discovery JSONs + vSphere
      DiskConsumedGB
- [ ] Exclude AS/400, Linux appliances, vCenter mgmt infra by default
      (with override)

## Underlying bugs
- [ ] **collect_vsphere_perf.py returning nulls for CPU.P95 / Memory.P95**
      — observed in FBA session (2026-05-19). SDT's own perf collector
      should pull 95th-pct from vSphere SOAP API directly so the Nutanix
      Collector XLSX becomes optional, not required. Investigate why the
      perf stats endpoint returned all-None values. Likely causes: stats
      collection not enabled at cluster level, perf query interval
      mismatch, or stats reset window too short.
- [ ] **Only 1 of N ESX hosts in cluster gets collected** — FBA cluster
      reports `NumHosts=4` but `ESXHosts` list has only 1. The collector
      should enumerate all hosts in the cluster, not just the connected
      one.
- [ ] **Multi-domain WinRM auth path** — SDT currently takes one cred
      set. When the environment spans multiple AD domains with no trust
      (FBA = fnbt.com + fb-al.com), 5+ servers came back "WinRM auth
      denied". Need either: (a) per-host cred override UI, or (b) accept
      a second domain\user/password pair on Setup, or (c) auto-detect
      target's domain via DNS and prompt for matching creds.

## GPO deep-dive before Entra migration (flagged 2026-07-17, needs more review before speccing)
- [ ] Matt wants GPO content reviewed in more depth pre-Entra-migration, not just a
      count/inventory — the concern is settings with no Entra ID/Intune equivalent
      getting silently dropped when on-prem DCs are decommissioned.
- [ ] Candidate approach: `Invoke-ServerDiscovery.ps1` already touches AD (FSMO,
      user/computer/OU counts) — extend to pull full GPO report (`Get-GPOReport`
      -ReportType XML per GPO, or `Get-GPResultantSetOfPolicy`) and surface
      settings by category (drive maps, printer deployment, folder redirection,
      software install, security/password policy, logon scripts) so each can be
      matched to an Entra/Intune equivalent or flagged as a gap.
- [ ] Matt said he needs to review this more before it's fully spec'd — don't
      build ahead of him; this entry is just to not lose the ask.
- [ ] Related: also flag Windows 10/11 **Home** edition workstations during
      discovery (Home can't Entra-join/Intune-enroll at all) — SDT already
      collects `Win32_OperatingSystem.Caption` per server/workstation
      (Invoke-ServerDiscovery.ps1 ~line 596), just needs a `cb-Flag 'warning'`
      when Caption contains "Home".

## Skipped / decided against
- LoB Apps tab — too unreliable (substring matching on app names has
  high false-positive rate). Better signal: extract LoB indicators from
  SQL DB names (Verafin, Finastra/LaserPro, BNControl, etc.) shown as a
  card on the SQL tab.
