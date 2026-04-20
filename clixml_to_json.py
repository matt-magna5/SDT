#!/usr/bin/env python3
"""
clixml_to_json.py - Convert PowerShell Export-Clixml output into the standard
SDT discovery JSON schema.

Used as the fallback path when ConvertTo-Json fails on DC/CIM-heavy boxes.

Usage:
    python clixml_to_json.py <path-to-clixml>  [--out <output.json>]

Emits a .json file next to the input by default. The output schema matches
what gen_report.py / gen_report_v2.py expect, so you can feed it straight
into the report generator.
"""
from __future__ import annotations
import argparse, json, re, sys, xml.etree.ElementTree as ET
from pathlib import Path

NS = '{http://schemas.microsoft.com/powershell/2004/04}'


def _text(elem):
    return elem.text if elem is not None and elem.text is not None else ''


def _parse_value(node):
    """Convert a single CLI-XML node to a Python value."""
    tag = node.tag.replace(NS, '')
    # Primitives
    if tag in ('S', 'URI'):         # string
        return _text(node)
    if tag in ('B',):               # bool
        return _text(node).strip().lower() == 'true'
    if tag in ('I32', 'I64', 'U32', 'U64', 'By', 'SB'):
        try: return int(_text(node))
        except: return _text(node)
    if tag in ('Db', 'Sg'):         # double / single
        try: return float(_text(node))
        except: return _text(node)
    if tag == 'DT':                 # DateTime
        return _text(node)
    if tag == 'Nil':
        return None
    if tag == 'C':                  # char
        try: return chr(int(_text(node)))
        except: return _text(node)
    if tag == 'G':                  # Guid
        return _text(node)
    if tag == 'Version':
        return _text(node)
    if tag == 'TS':                 # TimeSpan
        return _text(node)
    if tag == 'BA':                 # byte array (base64)
        return _text(node)
    if tag == 'Obj':
        return _parse_obj(node)
    if tag in ('LST', 'IE', 'QUE', 'STK'):  # lists/collections
        return [_parse_value(child) for child in node]
    if tag == 'DCT' or tag == 'Dictionary':
        return _parse_dict(node)
    if tag == 'Ref':
        # Reference to previously-seen object by id - collapse to string marker
        return f'<ref:{node.get("RefId", "?")}>'
    # Unknown element - recurse or stringify
    children = list(node)
    if children:
        return [_parse_value(c) for c in children]
    return _text(node)


def _parse_dict(node):
    """DCT/Dictionary element - pairs of En elements with S child keys."""
    out = {}
    for en in node.findall(f'{NS}En'):
        key_node = None
        val_node = None
        for child in en:
            n_attr = child.get('N', '')
            if n_attr == 'Key':   key_node = child
            elif n_attr == 'Value': val_node = child
        if key_node is not None and val_node is not None:
            k = _parse_value(key_node)
            out[str(k)] = _parse_value(val_node)
    return out


def _parse_obj(node):
    """Obj element - either a properties bag (Props/MS/LST) or typed wrapper."""
    # Check for TN/TNRef - type hint
    tn = node.find(f'{NS}TN')
    type_names = []
    if tn is not None:
        type_names = [t.text or '' for t in tn.findall(f'{NS}T')]

    # If it's a dictionary-shaped object, collect Props/MS children
    props = node.find(f'{NS}Props')
    ms = node.find(f'{NS}MS')
    lst = None
    for tag in ('LST', 'IE', 'QUE', 'STK'):
        lst = node.find(f'{NS}{tag}')
        if lst is not None: break
    dct = node.find(f'{NS}DCT')

    # Dict-like object (Hashtable, OrderedDictionary)
    if dct is not None:
        return _parse_dict(dct)

    # List-like object
    if lst is not None:
        items = [_parse_value(c) for c in lst]
        # If there were no Props/MS, return the pure list
        if props is None and ms is None:
            return items
        # Otherwise also collect properties (rare)
        out = {'_items': items}
        for bag in (props, ms):
            if bag is None: continue
            for child in bag:
                name = child.get('N') or child.tag.replace(NS, '')
                out[name] = _parse_value(child)
        return out

    # Plain property bag
    out = {}
    for bag in (props, ms):
        if bag is None: continue
        for child in bag:
            name = child.get('N') or child.tag.replace(NS, '')
            out[name] = _parse_value(child)
    # If the object itself has children outside Props/MS/LST/DCT, collect them too
    if not out:
        for child in node:
            tag = child.tag.replace(NS, '')
            if tag in ('TN', 'TNRef', 'ToString', 'Props', 'MS', 'LST', 'IE', 'QUE', 'STK', 'DCT'):
                continue
            name = child.get('N') or tag
            out[name] = _parse_value(child)
    # If still empty, fall back to ToString rendering
    if not out:
        tostring = node.find(f'{NS}ToString')
        if tostring is not None:
            return _text(tostring)
    return out


def convert(clixml_path: Path, out_path: Path):
    tree = ET.parse(str(clixml_path))
    root = tree.getroot()
    # The root is <Objs>; the first child Obj is our $discoveryResult
    first = None
    for child in root:
        if child.tag.endswith('Obj'):
            first = child
            break
    if first is None:
        raise SystemExit('No <Obj> found in CLI-XML root')
    data = _parse_value(first)
    out_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding='utf-8')
    return data


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('clixml', help='Path to .clixml from Invoke-ServerDiscovery emergency fallback')
    ap.add_argument('--out', default=None, help='Output .json path (default: alongside input)')
    args = ap.parse_args()
    src = Path(args.clixml).resolve()
    if not src.is_file():
        print(f'[error] not a file: {src}', file=sys.stderr); sys.exit(2)
    out = Path(args.out) if args.out else src.with_suffix('').with_suffix('.json')
    # For "hostname-discovery-date.clixml" we want "hostname-discovery-date.json"
    if src.suffix.lower() == '.clixml':
        out = src.with_suffix('.json')
    if args.out:
        out = Path(args.out)
    data = convert(src, out)
    n_keys = len(data) if isinstance(data, dict) else 0
    print(f'[ok] parsed {src.name} -> {out.name} ({n_keys} top-level keys)')


if __name__ == '__main__':
    main()
