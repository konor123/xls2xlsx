#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Recovery-oriented XLS -> XLSX converter.
No Microsoft Excel automation is used.

Input handling:
- Legacy OLE/BIFF .xls: xlrd, corruption-tolerant mode
- Excel 2003 XML disguised as .xls
- HTML table disguised as .xls
- CSV/TSV/text disguised as .xls
- Misnamed .xlsx: copied after ZIP validation

The internal BIFF recovery path preserves values, merged cells, many cell formats,
column widths and row heights where xlrd can read them. Formula expressions in old
BIFF files may only be available as cached values through xlrd.
"""
from __future__ import annotations

import argparse
import csv
import datetime as _dt
import html
from html.parser import HTMLParser
import io
import json
import os
from pathlib import Path
import re
import shutil
import sys
import traceback
import zipfile
import xml.etree.ElementTree as ET

HERE = Path(__file__).resolve().parent
RUNTIME = HERE / "runtime_packages"
if str(RUNTIME) not in sys.path:
    sys.path.insert(0, str(RUNTIME))

try:
    import xlsxwriter
except Exception as e:
    print(json.dumps({"ok": False, "code": "XLSXWRITER_MISSING", "message": str(e)}, ensure_ascii=False))
    sys.exit(21)

OLE_SIG = bytes.fromhex("D0CF11E0A1B11AE1")
ZIP_SIG = b"PK\x03\x04"


def emit(event: str, **kw):
    payload = {"event": event, **kw}
    print(json.dumps(payload, ensure_ascii=False), flush=True)


def sanitize_sheet_name(name: str, used: set[str]) -> str:
    name = re.sub(r"[\[\]:*?/\\]", "_", (name or "Sheet").strip()) or "Sheet"
    name = name[:31]
    base = name
    i = 2
    while name.lower() in used:
        suffix = f"_{i}"
        name = (base[: 31 - len(suffix)] + suffix)
        i += 1
    used.add(name.lower())
    return name


def rgb_hex(book, idx):
    try:
        rgb = book.colour_map.get(idx)
        if rgb and len(rgb) >= 3:
            return "#{:02X}{:02X}{:02X}".format(*rgb[:3])
    except Exception:
        pass
    return None


def build_xlrd_format(book, wb_out, xf_index, cache):
    if xf_index in cache:
        return cache[xf_index]

    props = {}
    try:
        xf = book.xf_list[xf_index]
    except Exception:
        cache[xf_index] = None
        return None

    try:
        font = book.font_list[xf.font_index]
        if getattr(font, "name", None):
            props["font_name"] = font.name
        h = getattr(font, "height", 0)
        if h:
            props["font_size"] = max(1, h / 20.0)
        if getattr(font, "bold", 0):
            props["bold"] = True
        if getattr(font, "italic", 0):
            props["italic"] = True
        if getattr(font, "underlined", 0) or getattr(font, "underline_type", 0):
            props["underline"] = 1
        if getattr(font, "struck_out", 0):
            props["font_strikeout"] = True
        c = rgb_hex(book, getattr(font, "colour_index", 0))
        if c:
            props["font_color"] = c
    except Exception:
        pass

    try:
        bg = xf.background
        if getattr(bg, "fill_pattern", 0):
            c = rgb_hex(book, getattr(bg, "pattern_colour_index", 0))
            if c:
                props["bg_color"] = c
                props["pattern"] = 1
    except Exception:
        pass

    try:
        b = xf.border
        side_map = [
            ("left", "left_line_style", "left_colour_index"),
            ("right", "right_line_style", "right_colour_index"),
            ("top", "top_line_style", "top_colour_index"),
            ("bottom", "bottom_line_style", "bottom_colour_index"),
        ]
        for out_name, style_attr, color_attr in side_map:
            st = int(getattr(b, style_attr, 0) or 0)
            if st:
                props[out_name] = min(st, 13)
                c = rgb_hex(book, getattr(b, color_attr, 0))
                if c:
                    props[out_name + "_color"] = c
    except Exception:
        pass

    try:
        al = xf.alignment
        hmap = {1: "left", 2: "center", 3: "right", 4: "fill", 5: "justify", 6: "center_across", 7: "distributed"}
        vmap = {0: "top", 1: "vcenter", 2: "bottom", 3: "vjustify", 4: "vdistributed"}
        if getattr(al, "hor_align", 0) in hmap:
            props["align"] = hmap[al.hor_align]
        if getattr(al, "vert_align", 2) in vmap:
            props["valign"] = vmap[al.vert_align]
        if getattr(al, "text_wrapped", 0):
            props["text_wrap"] = True
        rot = int(getattr(al, "rotation", 0) or 0)
        if rot:
            if rot == 255:
                pass
            elif 0 <= rot <= 90:
                props["rotation"] = rot
            elif 91 <= rot <= 180:
                props["rotation"] = 90 - rot
    except Exception:
        pass

    try:
        fmt_obj = book.format_map.get(xf.format_key)
        if fmt_obj and getattr(fmt_obj, "format_str", None):
            props["num_format"] = fmt_obj.format_str
    except Exception:
        pass

    try:
        fmt = wb_out.add_format(props)
    except Exception:
        fmt = None
    cache[xf_index] = fmt
    return fmt


def is_date_format_string(s: str) -> bool:
    if not s:
        return False
    # Heuristic only for XML/HTML helpers.
    s2 = re.sub(r'"[^"]*"', "", s.lower())
    return any(tok in s2 for tok in ("yy", "mm", "dd", "hh", "ss"))


def convert_biff_xls(src: Path, dst: Path):
    try:
        import xlrd
    except Exception as e:
        raise RuntimeError("XLDR_MISSING:" + str(e))

    emit("engine", name="internal-biff")

    book = None
    errs = []
    attempts = [
        dict(formatting_info=True, on_demand=False, ignore_workbook_corruption=True),
        dict(formatting_info=False, on_demand=False, ignore_workbook_corruption=True),
        dict(formatting_info=False, on_demand=False),
    ]
    for opts in attempts:
        try:
            book = xlrd.open_workbook(str(src), **opts)
            break
        except TypeError:
            opts.pop("ignore_workbook_corruption", None)
            try:
                book = xlrd.open_workbook(str(src), **opts)
                break
            except Exception as e:
                errs.append(repr(e))
        except Exception as e:
            errs.append(repr(e))

    if book is None:
        raise RuntimeError("Unable to read the BIFF workbook even in recovery mode: " + " | ".join(errs[-3:]))

    temp = dst.with_suffix(dst.suffix + ".partial")
    if temp.exists():
        temp.unlink()

    wb = xlsxwriter.Workbook(str(temp), {"strings_to_formulas": False, "strings_to_urls": False, "nan_inf_to_errors": True})
    used = set()
    fmt_cache = {}
    has_formats = bool(getattr(book, "formatting_info", False))

    try:
        for sidx, sh in enumerate(book.sheets()):
            ws = wb.add_worksheet(sanitize_sheet_name(sh.name or f"Sheet{sidx+1}", used))
            emit("sheet", index=sidx + 1, total=book.nsheets, name=sh.name)

            if has_formats:
                try:
                    for cidx, ci in sh.colinfo_map.items():
                        width = max(0.1, min(255.0, ci.width / 256.0))
                        ws.set_column(cidx, cidx, width, None, {"hidden": bool(ci.hidden)})
                except Exception:
                    pass
                try:
                    for ridx, ri in sh.rowinfo_map.items():
                        height = None
                        if getattr(ri, "height", 0):
                            height = ri.height / 20.0
                        ws.set_row(ridx, height, None, {"hidden": bool(getattr(ri, "hidden", 0))})
                except Exception:
                    pass

            merged = []
            covered = set()
            try:
                merged = list(sh.merged_cells)
                for rlo, rhi, clo, chi in merged:
                    for rr in range(rlo, rhi):
                        for cc in range(clo, chi):
                            covered.add((rr, cc))
            except Exception:
                merged = []

            # Merged ranges first.
            for rlo, rhi, clo, chi in merged:
                cell = sh.cell(rlo, clo)
                fmt = None
                try:
                    if has_formats:
                        fmt = build_xlrd_format(book, wb, cell.xf_index, fmt_cache)
                except Exception:
                    pass
                val = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        val = xlrd.xldate.xldate_as_datetime(val, book.datemode)
                    except Exception:
                        pass
                elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                    val = bool(val)
                elif cell.ctype == xlrd.XL_CELL_ERROR:
                    val = xlrd.biffh.error_text_from_code.get(val, f"#ERR{val}")
                ws.merge_range(rlo, clo, rhi - 1, chi - 1, val, fmt)

            for r in range(sh.nrows):
                for c in range(sh.ncols):
                    if (r, c) in covered:
                        continue
                    cell = sh.cell(r, c)
                    if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                        # Keep blank formatting when available.
                        if cell.ctype == xlrd.XL_CELL_BLANK and has_formats:
                            fmt = build_xlrd_format(book, wb, cell.xf_index, fmt_cache)
                            if fmt:
                                ws.write_blank(r, c, None, fmt)
                        continue

                    fmt = None
                    if has_formats:
                        try:
                            fmt = build_xlrd_format(book, wb, cell.xf_index, fmt_cache)
                        except Exception:
                            fmt = None

                    if cell.ctype == xlrd.XL_CELL_TEXT:
                        ws.write_string(r, c, str(cell.value), fmt)
                    elif cell.ctype == xlrd.XL_CELL_NUMBER:
                        ws.write_number(r, c, float(cell.value), fmt)
                    elif cell.ctype == xlrd.XL_CELL_DATE:
                        try:
                            dt = xlrd.xldate.xldate_as_datetime(cell.value, book.datemode)
                            ws.write_datetime(r, c, dt, fmt)
                        except Exception:
                            ws.write_number(r, c, float(cell.value), fmt)
                    elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                        ws.write_boolean(r, c, bool(cell.value), fmt)
                    elif cell.ctype == xlrd.XL_CELL_ERROR:
                        txt = xlrd.biffh.error_text_from_code.get(cell.value, f"#ERR{cell.value}")
                        ws.write_string(r, c, txt, fmt)
                    else:
                        ws.write(r, c, cell.value, fmt)

            # Best-effort freeze panes from parsed sheet attributes.
            try:
                hrz = int(getattr(sh, "horz_split_pos", 0) or 0)
                vrt = int(getattr(sh, "vert_split_pos", 0) or 0)
                if hrz or vrt:
                    ws.freeze_panes(hrz, vrt)
            except Exception:
                pass

        wb.close()
        os.replace(temp, dst)
    except Exception:
        try:
            wb.close()
        except Exception:
            pass
        if temp.exists():
            temp.unlink()
        raise


class TableHTMLParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables = []
        self.current = None
        self.row = None
        self.cell = None
        self.cell_attrs = None
        self.in_table = False

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        attrs = dict(attrs)
        if tag == "table":
            self.in_table = True
            self.current = []
        elif self.in_table and tag == "tr":
            self.row = []
        elif self.in_table and tag in ("td", "th"):
            self.cell = []
            self.cell_attrs = attrs
        elif self.cell is not None and tag == "br":
            self.cell.append("\n")

    def handle_data(self, data):
        if self.cell is not None:
            self.cell.append(data)

    def handle_endtag(self, tag):
        tag = tag.lower()
        if self.in_table and tag in ("td", "th") and self.cell is not None:
            text = "".join(self.cell).strip()
            self.row.append((text, self.cell_attrs or {}))
            self.cell = None
            self.cell_attrs = None
        elif self.in_table and tag == "tr":
            if self.row is not None and self.current is not None:
                self.current.append(self.row)
            self.row = None
        elif tag == "table" and self.in_table:
            if self.current is not None:
                self.tables.append(self.current)
            self.current = None
            self.in_table = False


def coerce_text_value(s):
    t = s.strip()
    if t == "":
        return ""
    if re.fullmatch(r"[-+]?\d+", t):
        try:
            return int(t)
        except Exception:
            return t
    if re.fullmatch(r"[-+]?(?:\d+\.\d*|\.\d+)", t):
        try:
            return float(t)
        except Exception:
            return t
    return t


def convert_html(src: Path, dst: Path, decoded: str):
    emit("engine", name="internal-html")
    p = TableHTMLParser()
    p.feed(decoded)
    if not p.tables:
        raise RuntimeError("The file appears to be HTML, but no table was found.")

    temp = dst.with_suffix(dst.suffix + ".partial")
    wb = xlsxwriter.Workbook(str(temp))
    used = set()
    try:
        for ti, table in enumerate(p.tables, 1):
            ws = wb.add_worksheet(sanitize_sheet_name(f"Table{ti}", used))
            occupied = set()
            for r, row in enumerate(table):
                c = 0
                for text, attrs in row:
                    while (r, c) in occupied:
                        c += 1
                    rs = max(1, int(attrs.get("rowspan", "1") or "1"))
                    cs = max(1, int(attrs.get("colspan", "1") or "1"))
                    val = coerce_text_value(html.unescape(text))
                    if rs > 1 or cs > 1:
                        ws.merge_range(r, c, r + rs - 1, c + cs - 1, val)
                        for rr in range(r, r + rs):
                            for cc in range(c, c + cs):
                                occupied.add((rr, cc))
                    else:
                        ws.write(r, c, val)
                    c += cs
        wb.close()
        os.replace(temp, dst)
    except Exception:
        try: wb.close()
        except Exception: pass
        if temp.exists(): temp.unlink()
        raise


def xml_attr(el, local):
    for k, v in el.attrib.items():
        if k.endswith("}" + local) or k == local or k.endswith(":" + local):
            return v
    return None


def convert_xml_spreadsheet(src: Path, dst: Path, raw: bytes):
    emit("engine", name="internal-xml2003")
    root = ET.fromstring(raw)
    worksheets = [e for e in root.iter() if e.tag.endswith("Worksheet")]
    if not worksheets:
        raise RuntimeError("No Worksheet element was found in the Excel XML Spreadsheet.")

    temp = dst.with_suffix(dst.suffix + ".partial")
    wb = xlsxwriter.Workbook(str(temp))
    used = set()
    try:
        for wi, wse in enumerate(worksheets, 1):
            name = xml_attr(wse, "Name") or f"Sheet{wi}"
            ws = wb.add_worksheet(sanitize_sheet_name(name, used))
            table = next((e for e in wse.iter() if e.tag.endswith("Table")), None)
            if table is None:
                continue
            r_out = 0
            for row in [e for e in list(table) if e.tag.endswith("Row")]:
                idx = xml_attr(row, "Index")
                if idx:
                    r_out = int(idx) - 1
                c_out = 0
                for cell in [e for e in list(row) if e.tag.endswith("Cell")]:
                    cidx = xml_attr(cell, "Index")
                    if cidx:
                        c_out = int(cidx) - 1
                    data = next((e for e in list(cell) if e.tag.endswith("Data")), None)
                    txt = "" if data is None else "".join(data.itertext())
                    typ = xml_attr(data, "Type") if data is not None else None
                    val = txt
                    if typ == "Number":
                        try: val = float(txt)
                        except Exception: pass
                    elif typ == "Boolean":
                        val = txt.strip() not in ("0", "false", "False", "")
                    elif typ == "DateTime":
                        try:
                            val = _dt.datetime.fromisoformat(txt.replace("Z", "+00:00")).replace(tzinfo=None)
                        except Exception:
                            pass

                    ma = int(xml_attr(cell, "MergeAcross") or 0)
                    md = int(xml_attr(cell, "MergeDown") or 0)
                    if ma or md:
                        ws.merge_range(r_out, c_out, r_out + md, c_out + ma, val)
                    else:
                        if isinstance(val, _dt.datetime):
                            ws.write_datetime(r_out, c_out, val)
                        else:
                            ws.write(r_out, c_out, val)
                    c_out += 1 + ma
                r_out += 1
        wb.close()
        os.replace(temp, dst)
    except Exception:
        try: wb.close()
        except Exception: pass
        if temp.exists(): temp.unlink()
        raise


def decode_text(raw: bytes):
    # Common Korean/Windows encodings first.
    for enc in ("utf-8-sig", "utf-16", "cp949", "euc-kr", "cp1252", "latin1"):
        try:
            text = raw.decode(enc)
            bad = text.count("\ufffd")
            if bad == 0:
                return text, enc
        except Exception:
            pass
    return raw.decode("latin1", errors="replace"), "latin1"


def convert_delimited(src: Path, dst: Path, decoded: str):
    emit("engine", name="internal-text")
    sample = decoded[:20000]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",\t;|")
    except Exception:
        class D(csv.excel):
            delimiter = "\t" if "\t" in sample else ","
        dialect = D

    rows = list(csv.reader(io.StringIO(decoded), dialect))
    if not rows:
        raise RuntimeError("No text/CSV data was found.")
    temp = dst.with_suffix(dst.suffix + ".partial")
    wb = xlsxwriter.Workbook(str(temp))
    try:
        ws = wb.add_worksheet("Sheet1")
        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                ws.write(r, c, coerce_text_value(value))
        wb.close()
        os.replace(temp, dst)
    except Exception:
        try: wb.close()
        except Exception: pass
        if temp.exists(): temp.unlink()
        raise


def convert_misnamed_xlsx(src: Path, dst: Path):
    with zipfile.ZipFile(src, "r") as z:
        names = set(z.namelist())
        if "xl/workbook.xml" not in names:
            raise RuntimeError("The file is a ZIP container, but it is not a valid XLSX structure.")
    shutil.copyfile(src, dst)
    emit("engine", name="copy-misnamed-xlsx")


def classify(raw: bytes):
    head = raw[:4096]
    if raw.startswith(OLE_SIG):
        return "biff"
    if raw.startswith(ZIP_SIG):
        return "zip"
    text, enc = decode_text(head)
    s = text.lstrip().lower()
    if "<html" in s[:1000] or "<table" in s[:2000]:
        return "html"
    if s.startswith("<?xml") or "<workbook" in s[:2000]:
        return "xml"
    return "text"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("dst")
    args = ap.parse_args()

    src = Path(args.src).resolve()
    dst = Path(args.dst).resolve()
    dst.parent.mkdir(parents=True, exist_ok=True)
    if dst.exists():
        dst.unlink()

    if not src.is_file():
        raise FileNotFoundError(src)

    raw = src.read_bytes()
    kind = classify(raw)
    emit("detected", kind=kind, size=len(raw))

    if kind == "biff":
        convert_biff_xls(src, dst)
    elif kind == "zip":
        convert_misnamed_xlsx(src, dst)
    else:
        decoded, enc = decode_text(raw)
        emit("encoding", name=enc)
        if kind == "html":
            convert_html(src, dst, decoded)
        elif kind == "xml":
            try:
                convert_xml_spreadsheet(src, dst, raw)
            except Exception:
                # Some "XML-like" .xls exports are actually HTML-ish.
                if "<table" in decoded.lower():
                    convert_html(src, dst, decoded)
                else:
                    raise
        else:
            convert_delimited(src, dst, decoded)

    if not dst.exists() or dst.stat().st_size < 100:
        raise RuntimeError("The output file was not created correctly.")

    # Validate generated XLSX ZIP container.
    with zipfile.ZipFile(dst, "r") as z:
        if "xl/workbook.xml" not in set(z.namelist()):
            raise RuntimeError("The generated output is not a valid XLSX structure.")

    emit("done", output=str(dst), bytes=dst.stat().st_size)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        msg = str(e)
        code = "XLDR_MISSING" if msg.startswith("XLDR_MISSING:") else "CONVERT_FAILED"
        print(json.dumps({"ok": False, "code": code, "message": msg, "trace": traceback.format_exc(limit=5)}, ensure_ascii=False), flush=True)
        sys.exit(12 if code == "XLDR_MISSING" else 10)
