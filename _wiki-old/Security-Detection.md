# Security Detection

The report generator automatically detects installed security and management products from two sources:
- **Installed Applications** — registry-based app list (`Win32_Product` / Uninstall registry keys)
- **Running Services** — Windows service display names and service names

Detection is keyword-based — if any keyword matches any app name, publisher, service display name, or service name (case-insensitive), the product is identified.

---

## Detection Categories

### 🛡️ EDR / XDR

Endpoint Detection & Response / Extended Detection & Response

| Product | Detection Keywords |
|---|---|
| SentinelOne | `sentinelone`, `sentinel agent` |
| CrowdStrike Falcon | `crowdstrike`, `csfalcon`, `csagent` |
| Huntress MDR | `huntress` |
| BlackBerry Cylance | `cylance` |
| VMware Carbon Black | `carbon black`, `cb defense` |
| Palo Alto Cortex XDR | `cortex xdr`, `cyserver`, `traps` |
| Cybereason | `cybereason`, `cramtray` |
| Deep Instinct | `deep instinct` |
| Darktrace | `darktrace` |
| Elastic Security | `elastic security`, `elastic agent`, `elastic endpoint` |
| Trellix (McAfee) | `trellix`, `mcafee` |
| ESET | `eset endpoint`, `ekrn` |
| Sophos | `sophos`, `savservice` |
| Symantec / Norton | `symantec endpoint`, `norton` |
| Kaspersky | `kaspersky`, `avp service` |
| Webroot | `webroot`, `wrsa` |
| Bitdefender | `bitdefender` |
| Malwarebytes | `malwarebytes`, `mbendpointagent` |
| Cisco Secure Endpoint | `cisco secure endpoint`, `cisco amp` |
| F-Secure / WithSecure | `f-secure`, `withsecure` |
| Avast | `avast` |
| AVG | `avg ` |
| Adlumin MDR | `adlumin` |
| Arctic Wolf MDR | `arctic wolf` |
| Blackpoint Cyber | `blackpoint` |
| Ontinue MDR | `ontinue` |
| Netsurion MDR | `netsurion` |
| Microsoft Defender for Endpoint | `microsoft defender for endpoint`, `sense` |

> **Note:** Windows Defender (`windefend`) is included in the EDR service bucket for the Services card, but `detect_security` will identify a third-party EDR if present. If only Defender is found, the Security & Protection panel shows "None detected" for EDR (Defender is considered baseline, not a dedicated EDR product).

---

### ⚙️ RMM

Remote Monitoring & Management

| Product | Detection Keywords |
|---|---|
| N-able N-sight / N-central | `n-able`, `n_able`, `solarwinds`, `advanced monitoring agent`, `windows agent service` |
| NinjaOne RMM | `ninjarmm`, `ninjaone` |
| Kaseya VSA | `kaseya` |
| ConnectWise Automate | `connectwise automate`, `labtech`, `ltsvc` |
| ConnectWise RMM | `connectwise rmm` |
| Datto RMM | `datto rmm`, `autotask endpoint`, `cagservice` |
| Syncro RMM | `syncro`, `kabuto` |
| Atera RMM | `atera` |
| Pulseway | `pulseway`, `pc monitor` |
| Splashtop RMM | `splashtop`, `srservice` |
| ManageEngine Endpoint Central | `manageengine`, `uems agent` |
| Naverisk | `naverisk` |
| Barracuda RMM | `barracuda rmm` |
| GoTo Resolve | `goto resolve`, `logmein` |
| TeamViewer RMM | `teamviewer` |
| ITarian | `itarian` |

---

### 🔗 Remote Access

Unattended remote control and remote support tools

| Product | Detection Keywords |
|---|---|
| ScreenConnect | `screenconnect`, `connectwise control` |
| N-able Take Control | `basupportexpress`, `msp anywhere`, `take control` |
| TeamViewer | `teamviewer` |
| AnyDesk | `anydesk` |
| Splashtop | `splashtop`, `srservice` |
| LogMeIn / GoTo | `logmein`, `goto resolve` |
| BeyondTrust Remote Support | `bomgar`, `beyondtrust remote`, `sra-pin` |
| Dameware | `dameware`, `dwrcs` |
| TightVNC | `tightvnc` |
| UltraVNC | `ultravnc`, `uvnc` |
| RealVNC | `realvnc` |
| TigerVNC | `tigervnc` |
| VNC (generic) | `winvnc` |
| Zoho Assist | `zoho assist`, `zohourmservice` |
| ISL Online | `isl alwayson`, `islalwayson` |
| Supremo | `supremo` |
| Radmin | `rserver3`, `radmin` |
| Netop Remote | `nhostsvc` |
| RustDesk | `rustdesk` |
| Chrome Remote Desktop | `chromoting`, `remoting_host` |
| Parsec | `parsec` |
| NoMachine | `nomachine` |
| Iperius Remote | `iperius remote`, `iperiusremote` |
| Cisco Webex Remote Support | `webex support`, `atashost` |

---

### 💾 Backup

Backup agents and data protection software

| Product | Detection Keywords |
|---|---|
| Veeam | `veeam` |
| Acronis | `acronis` |
| Datto BCDR | `datto bcdr`, `datto backup` |
| Commvault | `commvault`, `cvreplication` |
| Veritas Backup Exec | `backup exec`, `bengine` |
| StorageCraft ShadowProtect | `shadowprotect`, `shadow protect` |
| Arcserve | `arcserve` |
| MSP360 / CloudBerry | `msp360`, `cloudberry` |
| Azure Backup (MARS) | `azure backup`, `cbengine`, `microsoft azure recovery` |
| Druva inSync | `druva`, `insync` |
| Cohesity | `cohesity` |
| Rubrik | `rubrik` |
| Zerto | `zerto` |
| Unitrends | `unitrends` |
| Barracuda Backup | `barracuda backup` |
| Axcient | `axcient` |
| IDrive | `idrive` |
| Carbonite | `carbonite` |
| Backblaze | `backblaze` |
| N2WS Backup | `n2ws` |
| Windows Server Backup | `windows server backup`, `wbengine` |

---

### 🔑 PAM

Privileged Access Management

| Product | Detection Keywords |
|---|---|
| CyberArk EPM | `cyberark`, `vf_agent` |
| BeyondTrust | `beyondtrust`, `avecto`, `pgdriver` |
| Delinea / Thycotic | `delinea`, `thycotic`, `arellia` |
| One Identity Safeguard | `one identity safeguard` |
| WALLIX Bastion | `wallix` |
| HashiCorp Vault | `hashicorp vault` |
| Senhasegura | `senhasegura` |
| ARCON PAM | `arcon pam` |
| Quickpass | `quickpass` |
| Saviynt | `saviynt` |

---

## Adding New Products

To add detection for a new product, edit `detect_security()` in `gen_report.py`:

```python
# In the appropriate _edr / _rmm / _rmt / _bdr / _pam list:
('keyword_to_match', 'Display Name in Report'),
```

Keywords are matched case-insensitively against concatenated app name + publisher + service display name + service name strings.

---

## Services Card Integration

The Services card in the Advanced view groups running services by type. The EDR bucket label dynamically shows the detected product:

- **With detection:** `EDR / Endpoint Protection — Sophos`
- **Without detection:** `EDR / Endpoint Protection`

EDR services recognized in the bucket include service names for all major products listed above.
