# SDT v4.1 Backlog

## UX polish
- [ ] Hypervisor card: clean up spacing/alignment. Password field + Scan button row
      is awkwardly sized (see 4/23 screenshot). Button should probably be right-
      aligned and match input field height exactly. Consider moving Scan button
      below the row as a full-width or right-floated action.
- [ ] Grid-column alignment across all cards (currently Type/IP/User grid and
      Password/Scan grid don't line up — inputs wider on row 1 than row 2).
- [ ] Form field label spacing - tighten vertical rhythm inside grid cells.

## Features
- [ ] Subnet auto-discovery (AD+DNS+ping sweep) - still stubbed
- [ ] Hyper-V scan path (currently only vSphere via collect_vsphere_perf.py)
- [ ] Client logo upload on Setup tab -> embeds in final report
- [ ] Export session as .zip (report + JSONs) at end of run
- [ ] M365 Graph integration
- [ ] Network switch / LLDP discovery
- [ ] Auto network diagram
