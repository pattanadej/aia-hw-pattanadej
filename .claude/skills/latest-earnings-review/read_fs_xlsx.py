#!/usr/bin/env python3
"""อ่านไฟล์ FINANCIAL_STATEMENTS.XLSX (แบบที่ ก.ล.ต./SET ให้โหลด) ด้วย Python มาตรฐานล้วน ไม่ต้องติดตั้ง openpyxl

ใช้:
  python -I read_fs_xlsx.py <ไฟล์.XLSX> --sheets                  # ลิสต์ชีตและขนาด
  python -I read_fs_xlsx.py <ไฟล์.XLSX> --dump [ชีต ...]          # พิมพ์ทุกแถวที่มีข้อมูล (ทุกชีตถ้าไม่ระบุ)
  python -I read_fs_xlsx.py <ไฟล์.XLSX> --find "ลูกหนี้" "Revenue"  # ค้นแถวที่ป้ายมีคำเหล่านี้ (ไม่สนตัวพิมพ์)
  เพิ่ม --json เพื่อให้ผลเป็น JSON (ให้สคริปต์คำนวณอื่นเรียกต่อ)

หมายเหตุ: อ่านเฉพาะค่าที่เก็บไว้ในไฟล์ (cached value) ไม่คำนวณสูตรใหม่ อ่านอย่างเดียว ไม่แก้ไฟล์
หน่วย/คอลัมน์ (งบรวม/เฉพาะกิจการ, งวดนี้/งวดก่อน/สะสม) ต้องดูหัวตารางของแต่ละชีตเอง — ห้ามสมมติ
"""
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"


def col_index(ref):
    letters = re.match(r"[A-Z]+", ref).group(0)
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n - 1


def load(path):
    z = zipfile.ZipFile(path)
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root.findall("m:si", NS):
            shared.append("".join(t.text or "" for t in si.iter("{%s}t" % NS["m"])))
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    target = {r.get("Id"): r.get("Target") for r in rels.findall("{%s}Relationship" % PKG)}
    sheets = {}
    for s in wb.find("m:sheets", NS):
        tgt = target[s.get("{%s}id" % RNS)].lstrip("/")
        tgt = tgt if tgt.startswith("xl/") else "xl/" + tgt
        rows = []
        root = ET.fromstring(z.read(tgt))
        for row in root.iter("{%s}row" % NS["m"]):
            cells = {}
            for c in row.findall("m:c", NS):
                v = c.find("m:v", NS)
                t = c.get("t")
                if t == "inlineStr":
                    val = "".join(x.text or "" for x in c.iter("{%s}t" % NS["m"]))
                elif v is None or v.text is None:
                    continue
                elif t == "s":
                    val = shared[int(v.text)]
                elif t in ("str", "e", "b"):
                    val = v.text
                else:
                    try:
                        f = float(v.text)
                        val = int(f) if f.is_integer() else f
                    except ValueError:
                        val = v.text
                if val != "":
                    cells[col_index(c.get("r"))] = val
            if cells:
                rows.append((int(row.get("r")), [cells.get(i) for i in range(max(cells) + 1)]))
        sheets[s.get("name")] = rows
    return sheets


def main(argv):
    if len(argv) < 3:
        print(__doc__)
        return 1
    path, args = argv[1], argv[2:]
    as_json = "--json" in args
    args = [a for a in args if a != "--json"]
    sheets = load(path)
    out = []
    if args[0] == "--sheets":
        for name, rows in sheets.items():
            out.append({"sheet": name, "rows": len(rows)})
    elif args[0] == "--dump":
        pick = args[1:] or list(sheets)
        for name in pick:
            for r, vals in sheets[name]:
                out.append({"sheet": name, "row": r, "cells": vals})
    elif args[0] == "--find":
        keys = [k.lower() for k in args[1:]]
        for name, rows in sheets.items():
            for r, vals in rows:
                text = " ".join(str(v) for v in vals if isinstance(v, str)).lower()
                if any(k in text for k in keys):
                    out.append({"sheet": name, "row": r, "cells": vals})
    else:
        print(__doc__)
        return 1
    if as_json:
        print(json.dumps(out, ensure_ascii=False))
    else:
        for o in out:
            if "cells" in o:
                print(f"[{o['sheet']}] r{o['row']}: " + " | ".join("" if v is None else str(v) for v in o["cells"]))
            else:
                print(f"{o['sheet']}\t{o['rows']} แถว")
    return 0


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    sys.exit(main(sys.argv))
