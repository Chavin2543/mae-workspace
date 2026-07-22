#!/usr/bin/env python3
"""Restore cached formula values that openpyxl drops on save.

openpyxl writes formulas WITHOUT their cached results, so any plain
openpyxl save makes every formula cell read as None under data_only=True
(Excel still recalculates fine on open, but our extraction scripts and
any pandas/openpyxl consumer break). This tool copies the cached <v>
elements from a donor file (the last good version, e.g. from git) into
the freshly-saved file for every formula cell that lost its cache.

Cells whose formula results legitimately changed will carry a stale cache
until Excel next recalculates — same behaviour as a surgical XML patch.

Usage:
  python3 scripts/restore_formula_caches.py <donor.xlsx> <target.xlsx>
"""
import shutil
import sys
import tempfile
import zipfile

from lxml import etree

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}


def sheet_files(zf):
    """Map sheet name -> xml path inside the xlsx."""
    wbx = etree.fromstring(zf.read("xl/workbook.xml"))
    rels = etree.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rid_to_target = {rel.get("Id"): rel.get("Target")
                     for rel in rels.findall("{*}Relationship")}
    out = {}
    for sh in wbx.findall("m:sheets/m:sheet", NS):
        target = rid_to_target[sh.get("{%s}id" % NS["r"])]
        if not target.startswith("/"):
            target = "xl/" + target
        out[sh.get("name")] = target.lstrip("/")
    return out


def cached_values(tree):
    """cellref -> (t_attr, v_text) for formula cells WITH a cached value."""
    out = {}
    for c in tree.iter("{%s}c" % NS["m"]):
        f = c.find("m:f", NS)
        v = c.find("m:v", NS)
        if f is not None and v is not None:
            out[c.get("r")] = (c.get("t"), v.text)
    return out


def main(donor_path, target_path):
    with zipfile.ZipFile(donor_path) as zd:
        donor_sheets = sheet_files(zd)
        donor_cache = {name: cached_values(etree.fromstring(zd.read(path)))
                       for name, path in donor_sheets.items()}

    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False).name
    restored = 0
    with zipfile.ZipFile(target_path) as zt:
        target_sheets = sheet_files(zt)
        path_to_name = {v: k for k, v in target_sheets.items()}
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
            for item in zt.infolist():
                data = zt.read(item.filename)
                name = path_to_name.get(item.filename)
                if name and name in donor_cache and donor_cache[name]:
                    tree = etree.fromstring(data)
                    cache = donor_cache[name]
                    for c in tree.iter("{%s}c" % NS["m"]):
                        ref = c.get("r")
                        f = c.find("m:f", NS)
                        v = c.find("m:v", NS)
                        # openpyxl leaves either no <v> or an EMPTY <v/> for
                        # external-link and shared formulas — repair both
                        lost = v is None or not (v.text or "").strip()
                        if f is not None and lost and ref in cache:
                            t_attr, v_text = cache[ref]
                            if t_attr is not None:
                                c.set("t", t_attr)
                            if v is None:
                                v = etree.SubElement(c, "{%s}v" % NS["m"])
                            v.text = v_text
                            restored += 1
                    data = etree.tostring(tree, xml_declaration=True,
                                          encoding="UTF-8", standalone=True)
                zo.writestr(item, data)
    shutil.move(tmp, target_path)
    print(f"restored {restored} cached formula values into {target_path}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
