# Manifest Format

The session manifest is the index file that `gen_report.py` reads to locate all session data. It is written automatically by `Start-DiscoverySession` at the end of a session.

---

## Schema

```json
{
  "client": "Short client name (used in filenames)",
  "client_full": "Full legal client name (used in report header)",
  "date": "YYYY-MM-DD",
  "session_dir": "Path to session folder (. = same folder as manifest)",
  "output_dir": "Path for report output (. = same as session_dir)",
  "inventory_file": "Filename of primary HV inventory JSON",
  "logo_file": "Path to base64-encoded logo file (optional, leave empty)",
  "servers": [
    {
      "id": "unique-lowercase-id",
      "file": "SERVERNAME-discovery-DATE.json",
      "name": "SERVERNAME",
      "ip": "192.168.1.10",
      "in_scope": true
    }
  ]
}
```

---

## Field Reference

| Field | Required | Notes |
|---|---|---|
| `client` | Yes | Used in output filename: `ClientName-DiscoveryReport-DATE.html` |
| `client_full` | No | Falls back to `client` if omitted |
| `date` | Yes | Format: `YYYY-MM-DD` |
| `session_dir` | Yes | `.` means the same directory as the manifest file |
| `output_dir` | No | Falls back to `session_dir` if omitted |
| `inventory_file` | Yes | Primary HV inventory; additional inventories auto-loaded |
| `logo_file` | No | Path to a `.b64` file containing a base64-encoded PNG logo |
| `servers` | Yes | Array of server entries |

### Server Entry Fields

| Field | Required | Notes |
|---|---|---|
| `id` | Yes | Lowercase, no spaces — used as HTML element ID |
| `file` | Yes | Filename only (not full path) — looked up in `session_dir` |
| `name` | Yes | Display name in tab and report |
| `ip` | No | Shown in tab tooltip and server details |
| `in_scope` | Yes | `true` = full tab; `false` = tab shows with `◦` out-of-scope indicator |

---

## Manual Editing

You can edit the manifest by hand to:
- **Remove a server** — delete its entry from `servers` array
- **Mark out-of-scope** — set `"in_scope": false`
- **Add a server** — add an entry pointing to an existing discovery JSON
- **Change client name** — update `client` and `client_full`
- **Change the primary inventory** — update `inventory_file`

After editing, re-run `gen_report.py` to regenerate the report.

---

## Auto-Loaded Files

The report generator automatically loads additional data beyond the manifest:

| Pattern | What it loads |
|---|---|
| `*-inventory-*.json` with `"_type": "HyperVInventory"` | Additional HV host tabs |
| `*-inventory-*.json` with `"_type": "vSphereInventory"` | ESXi/vCenter data (if collected) |

These files must be in the same `session_dir`. Files with 0 bytes are skipped silently.

---

## Example (Minimal)

```json
{
  "client": "Acme Corp",
  "date": "2026-04-13",
  "session_dir": ".",
  "inventory_file": "ACHV01-inventory-2026-04-13.json",
  "servers": [
    { "id": "acdc01", "file": "ACDC01-discovery-2026-04-13.json", "name": "ACDC01", "ip": "10.0.0.10", "in_scope": true },
    { "id": "acfs01", "file": "ACFS01-discovery-2026-04-13.json", "name": "ACFS01", "ip": "10.0.0.11", "in_scope": true }
  ]
}
```
