# Troubleshooting

---

## WinRM Issues

### "Access Denied" connecting to a target
- Verify credentials are correct: domain admin or local admin on the target
- Workgroup hosts require local creds: `.\Administrator` format
- Check if the account is in the local Administrators group on the target
- Try: `Test-WSMan -ComputerName HOSTNAME -Credential $cred`

### WinRM not responding — falls back to WMI
The script automatically tries WMI/DCOM if WinRM fails. This is expected in some environments. WMI fallback gives a partial VM list (no IPs) but still gets the host inventory.

### "The WinRM client cannot process the request" — HTTPS/cert error
On PS 5.1, the script installs a cert bypass (`M5TrustAll`) for self-signed certs. If this still fails:
```powershell
winrm set winrm/config/client '@{AllowUnencrypted="true"}'
winrm set winrm/config/client/auth '@{Basic="true"}'
```

### WinRM left enabled after Ctrl+C
This should not happen — the cleanup handler fires on exit. If it does:
```powershell
# On the affected host:
Stop-Service WinRM -Force
Set-Service WinRM -StartupType Disabled
```

---

## Hypervisor Connection Issues

### Hyper-V — "No VMs returned"
- Confirm WinRM or WMI access to the HV host
- If using local admin: ensure you're using `.\Administrator` format
- Check that Hyper-V role is installed: `Get-WindowsFeature Hyper-V`

### ESXi/vCenter — API connection failed
- Confirm hostname resolves correctly
- Verify credentials (root or admin-level vSphere account)
- Self-signed cert errors are handled automatically on PS 5.1
- Try PS 7+ if cert bypass fails on PS 5.1

### vSphere REST API — "Could not connect"
The script tries the v7 API path first (`/api/session`), then falls back to v6 (`/rest/com/vmware/cis/session`). If both fail:
- Verify the host is actually vCenter/ESXi (not a different appliance)
- Check that the vSphere API is enabled and not blocked by firewall

---

## AD / DNS Scan Issues

### Suggested Servers scan returns nothing
- If the `ActiveDirectory` module isn't installed, the script falls back to ADSI/LDAP — confirm the machine is domain-joined
- If ADSI also fails (non-domain or access denied), the DNS sweep runs — this only works if common HV naming patterns are used
- Manually add targets at the "Additional Targets" prompt

### AD module errors
The script catches all exceptions from the AD scan and falls through gracefully. If you see errors, check:
```powershell
Import-Module ActiveDirectory
Get-ADComputer -Filter * -Properties Name | Select -First 5
```

---

## Report Generation Issues

### "Expecting value: line 1 column 1 (char 0)" — empty JSON
One or more discovery JSON files is 0 bytes. This happens when:
- WinRM connected but discovery script failed silently
- The target disconnected mid-run
- Check `session-log.txt` for the specific failure

Solution: remove the empty JSON entry from the manifest `servers` array, or re-run discovery against that target.

### "AttributeError: 'NoneType' object has no attribute 'lower'"
A field in an installed app has a `None` value where a string is expected. Fixed in v1.20 with `or ''` guards. If you see this in an older report, update `gen_report.py`.

### "JSONDecodeError" on manifest
The manifest file has invalid JSON. Common cause: interrupted write during session. Check `session-log.txt` — if the session completed, the manifest is usually intact. Try opening in a JSON validator.

### Report is blank / all tabs empty
- Confirm the `session_dir` in the manifest points to the right folder
- Confirm the server JSON files listed in `servers` actually exist in that folder

---

## Hyper-V Inventory Issues

### Only one HV host shows in the Hyper-V tab
The report auto-loads all `*-inventory-*.json` files in the session directory with `"_type": "HyperVInventory"`. If hosts are missing:
- Check if their inventory files are 0 bytes (connection failed)
- Verify the HV host was reachable during the session
- DUSOVH1/DUSOVH2-style issues: TTL=64 suggests Linux/ESXi — those can't be inventoried as Hyper-V

### VM RAM / Disk shows 0 in HV tab
The inventory uses `RAMgb` and `Disks[].SizeGB` fields. If these are null, VMware Tools or Hyper-V integration services may not be running on the VM.

---

## Python Issues

### "No module named openpyxl"
```bash
pip install openpyxl
```

### Python not found
The session launcher looks for `python` in PATH. If not found, it skips auto-report generation. Run `gen_report.py` manually:
```bash
python3 gen_report.py manifest.json
```

### Report generation succeeds but HTML is blank in browser
The HTML is self-contained and uses JavaScript for tab switching. If tabs don't work:
- Open in Chrome or Edge (not IE)
- If viewing from a network share, try copying to local disk first (browser security blocks some JS from UNC paths)
