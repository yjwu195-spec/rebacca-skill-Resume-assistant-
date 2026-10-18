"""Read or append a confirmed application while preserving WPS ZIP members.

inspect: python tracker_ooxml.py inspect
append:  python tracker_ooxml.py append confirmed.json [--commit]
Without --commit append only validates and reports the proposed change.
"""
import argparse
import copy
import datetime as dt
import hashlib
import io
import json
import os
from pathlib import Path
import re
import tempfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape
import zipfile

def load_settings():
    config = Path(os.environ.get("REBACCA_CONFIG", str(Path(__file__).resolve().parents[1] / "local-settings.json"))).expanduser()
    if not config.is_file():
        raise ValueError("Missing local-settings.json: run install.ps1 with your actual project, resume, and tracker paths")
    settings = json.loads(config.read_text(encoding="utf-8-sig"))
    for key in ("project_root", "resume_root", "tracker_path"):
        if not isinstance(settings.get(key), str) or not settings[key].strip():
            raise ValueError("Missing local path: " + key)
        if not Path(settings[key]).expanduser().is_absolute():
            raise ValueError("Local path must be absolute: " + key)
    return settings


TRACKER = Path(load_settings()["tracker_path"]).expanduser()
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
HEADERS = ["公司", "岗位", "投递时间", "base", "jd", "链接或渠道"]


def digest(data):
    return hashlib.sha256(data).hexdigest()


def package(data):
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        if z.testzip():
            raise ValueError("Invalid ZIP CRC")
        if len(z.namelist()) != len(set(z.namelist())):
            raise ValueError("Duplicate ZIP names")
        return {n: z.read(n) for n in z.namelist()}


def read_book(parts):
    shared = ["".join(x.itertext()) for x in ET.fromstring(parts.get(
        "xl/sharedStrings.xml", b'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'))]
    rels = {x.get("Id"): x.get("Target") for x in ET.fromstring(parts["xl/_rels/workbook.xml.rels"])}
    result = []
    for sheet in ET.fromstring(parts["xl/workbook.xml"]).find("m:sheets", NS):
        target = rels[sheet.get(REL)]
        path = target.lstrip("/") if target.startswith("/") else "xl/" + target
        root = ET.fromstring(parts[path])
        rows = []
        for row in root.findall("m:sheetData/m:row", NS):
            cells = {}
            for c in row:
                v = c.find("m:v", NS)
                text = v.text if v is not None else ""
                if c.get("t") == "s":
                    text = shared[int(text)]
                elif c.get("t") == "inlineStr":
                    text = "".join(c.find("m:is", NS).itertext())
                cells[re.sub(r"\d", "", c.get("r"))] = text
            rows.append({"row": int(row.get("r")), "cells": cells})
        result.append({"name": sheet.get("name"), "path": path, "rows": rows})
    return result


def summary(parts):
    sheets = read_book(parts)
    return {"path": str(TRACKER), "sheets": sheets,
            "media": {n: digest(v) for n, v in parts.items() if n.startswith("xl/media/")},
            "dispimg_count": sum(v.count(b"<f>_xlfn.DISPIMG") for n, v in parts.items()
                                 if n.startswith("xl/worksheets/") and n.endswith(".xml"))}


def plan(data, record):
    if record.get("confirmed_applied") is not True or record.get("confirmed_sheet") is not True:
        raise ValueError("Explicit applied-status and sheet confirmations are required")
    date = dt.date.fromisoformat(record["date"])
    for key in ("company", "role", "sheet", "location", "jd", "channel"):
        if not isinstance(record.get(key), str) or not record[key].strip():
            raise ValueError("Missing field: " + key)
    # A literal value is required here because this workbook uses numeric month.day.
    shown_date = f"{date.month}.{date.day:02d}"
    before = package(data)
    books = read_book(before)
    sheet = next((s for s in books if s["name"] == record["sheet"]), None)
    if sheet is None:
        raise ValueError("Confirmed sheet does not exist")
    header = sheet["rows"][0]["cells"]
    if [header.get(c) for c in "BCDEFG"] != HEADERS:
        raise ValueError("Unsupported layout: inspect and adapt; do not rebuild workbook")
    matches = [{"sheet": s["name"], **r} for s in books for r in s["rows"]
               if r["cells"].get("B", "").strip() == record["company"].strip()
               and r["cells"].get("C", "").strip() == record["role"].strip()]
    # The existing short date loses the year. Treat same month/day conservatively
    # as a possible duplicate rather than silently adding a second application.
    duplicate = [r for r in matches if r["cells"].get("D") and
                 r["cells"]["D"].strip() in (shown_date, record["date"])]
    if duplicate:
        return None, {"status": "already_present_or_same_short_date", "matches": duplicate}
    if any(not r["cells"].get("D") for r in matches):
        raise ValueError("Similar undated record needs review")
    n = max(r["row"] for r in sheet["rows"]) + 1
    path = sheet["path"]
    xml = before[path].decode("utf-8")
    root = ET.fromstring(before[path])
    if root.find("m:tableParts", NS) is not None or root.find("m:autoFilter", NS) is not None:
        raise ValueError("Table/filter requires a targeted extension not supported by this helper")
    if root.find("m:sheetProtection", NS) is not None:
        raise ValueError("Sheet is protected")
    for merge in root.findall("m:mergeCells/m:mergeCell", NS):
        if max(map(int, re.findall(r"\d+", merge.get("ref")))) >= n:
            raise ValueError("Merged area intersects append region")
    template = root.findall("m:sheetData/m:row", NS)[-1]
    styles = {re.sub(r"\d", "", c.get("r")): c.get("s") for c in template}
    values = dict(zip("BCDEFG", [record["company"], record["role"], shown_date,
                                  record["location"], record["jd"], record["channel"]]))
    attrs = {k: v for k, v in template.attrib.items() if k in ("ht", "customHeight", "spans")}
    attrs["r"] = str(n)
    row = "<row " + " ".join(f'{k}="{v}"' for k, v in attrs.items()) + ">"
    for col, value in values.items():
        style = f' s="{styles[col]}"' if styles.get(col) is not None else ""
        if col == "D":
            row += f'<c r="{col}{n}"{style}><v>{value}</v></c>'
        else:
            row += f'<c r="{col}{n}"{style} t="inlineStr"><is><t xml:space="preserve">{escape(value)}</t></is></c>'
    row += "</row>"
    if xml.count("</sheetData>") != 1:
        raise ValueError("Unexpected sheetData serialization")
    changed = xml.replace("</sheetData>", row + "</sheetData>")
    dim = root.find("m:dimension", NS)
    if dim is None or dim.get("ref") != f"B1:G{n-1}":
        raise ValueError("Unexpected dimension: inspect append area")
    changed = changed.replace(f'<dimension ref="B1:G{n-1}"', f'<dimension ref="B1:G{n}"', 1)
    changed_bytes = changed.encode("utf-8")
    # Removing exactly the new row and reverting dimension must recover every
    # original byte in the changed worksheet, including hyperlinks/formulas.
    restored = changed.replace(row, "", 1).replace(f'<dimension ref="B1:G{n}"', f'<dimension ref="B1:G{n-1}"', 1)
    if restored.encode("utf-8") != before[path]:
        raise ValueError("Existing worksheet bytes changed")
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as src, zipfile.ZipFile(out, "w") as dst:
        dst.comment = src.comment
        for info in src.infolist():
            dst.writestr(copy.copy(info), changed_bytes if info.filename == path else before[info.filename])
    result = out.getvalue()
    after = package(result)
    differences = [name for name in before if before[name] != after[name]]
    if set(before) != set(after) or differences != [path]:
        raise ValueError("Unrelated package members changed")
    actual = next(s for s in read_book(after) if s["name"] == record["sheet"])["rows"][-1]
    if actual != {"row": n, "cells": values}:
        raise ValueError("Appended values failed round-trip verification")
    report = {"status": "validated_plan", "sheet": record["sheet"], "range": f"B{n}:G{n}",
              "record": actual, "historical_matches_preserved": matches,
              "changed_members": differences, "unchanged_member_count": len(before)-1,
              "old_worksheet_bytes_preserved": True, "before_sha256": digest(data),
              "after_sha256": digest(result), "original_directory_writable": None,
              "wps_visual_check": "not_performed", "media_and_dispimg_before": {
                  k: v for k, v in summary(before).items() if k in ("media", "dispimg_count")}}
    return result, report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("operation", choices=["inspect", "append"])
    parser.add_argument("record", nargs="?")
    parser.add_argument("--commit", action="store_true")
    parser.add_argument("--report")
    args = parser.parse_args()
    data = TRACKER.read_bytes()
    if args.operation == "inspect":
        report = summary(package(data))
    else:
        if not args.record:
            parser.error("append requires a confirmed JSON record")
        result, report = plan(data, json.loads(Path(args.record).read_text(encoding="utf-8-sig")))
        if args.commit and result is not None:
            # No fallback path: creation must succeed in the ORIGINAL directory.
            temp = None
            try:
                fd, temp = tempfile.mkstemp(prefix=".rebacca-", suffix=".tmp", dir=TRACKER.parent)
                with os.fdopen(fd, "wb") as f:
                    f.write(result)
                    f.flush()
                    os.fsync(f.fileno())
                if digest(TRACKER.read_bytes()) != digest(data):
                    raise RuntimeError("Workbook changed concurrently; rerun inspection")
                if Path(temp).read_bytes() != result:
                    raise RuntimeError("Temporary write verification failed")
                os.replace(temp, TRACKER)
                temp = None
                if TRACKER.read_bytes() != result:
                    raise RuntimeError("Original-path readback failed")
                report["status"] = "written_and_verified"
                report["original_directory_writable"] = True
                report["path"] = str(TRACKER)
                report["media_and_dispimg_after"] = {k: v for k, v in summary(package(TRACKER.read_bytes())).items()
                                                      if k in ("media", "dispimg_count")}
            finally:
                if temp is not None:
                    Path(temp).unlink(missing_ok=True)
    if args.report:
        Path(args.report).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
