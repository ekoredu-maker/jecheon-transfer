# -*- coding: utf-8 -*-
"""제천 내신전보 - 공식 발령통지서 엑셀 템플릿 출력기.

표준 라이브러리(zipfile + ElementTree)만 사용하여 원본 XLSX 템플릿의
서식/도형/인쇄설정을 보존한 채 발령 데이터를 채웁니다.
"""
from __future__ import annotations

import copy
import io
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
XML_NS = "http://www.w3.org/XML/1998/namespace"

ET.register_namespace("x", NS)
ET.register_namespace("r", NS_R)
ET.register_namespace("a", NS_A)
ET.register_namespace("xdr", NS_XDR)

SHEET_PATHS = {
    "초등": "xl/worksheets/sheet2.xml",
    "유치원": "xl/worksheets/sheet3.xml",
    "전문상담": "xl/worksheets/sheet4.xml",
    "보건": "xl/worksheets/sheet5.xml",
}

GROUP_LABELS = {
    "초등": ("초등학교교사", "초등학교 교사"),
    "유치원": ("유치원교사", "유치원교사"),
    "전문상담": ("전문상담교사", "전문상담교사"),
    "보건": ("보건교사", "보건교사"),
}


def group_for_job(job: str) -> str | None:
    j = str(job or "").replace(" ", "")
    if "유치원" in j:
        return "유치원"
    if "전문상담" in j or ("상담" in j and "교사" in j):
        return "전문상담"
    if "보건" in j:
        return "보건"
    if "초등" in j:
        return "초등"
    return None


def official_position(job: str, group: str) -> str:
    if group in GROUP_LABELS:
        return GROUP_LABELS[group][0]
    return str(job or "교사")


def format_date(value: str, spaced: bool = False) -> str:
    raw = str(value or "").strip()
    m = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", raw)
    if not m:
        return raw
    y, mo, d = (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return f"{y}. {mo}. {d}." if spaced else f"{y}.{mo}.{d}."


def _row_number(cell_ref: str) -> int:
    m = re.search(r"(\d+)$", cell_ref or "")
    return int(m.group(1)) if m else 0


def _col_letters(cell_ref: str) -> str:
    m = re.match(r"([A-Z]+)", cell_ref or "")
    return m.group(1) if m else "A"


def _set_cell(cell: ET.Element, value):
    # preserve style and address, replace only value/formula representation
    for child in list(cell):
        if child.tag in (f"{{{NS}}}v", f"{{{NS}}}f", f"{{{NS}}}is"):
            cell.remove(child)
    if value is None or value == "":
        cell.attrib.pop("t", None)
        return
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        cell.attrib.pop("t", None)
        v = ET.SubElement(cell, f"{{{NS}}}v")
        v.text = str(value)
        return
    # artifact_tool이 생성한 원본 템플릿과 동일하게 t="str" + <v>를 사용합니다.
    # 이 방식은 원본 서식/문자열을 가장 안정적으로 왕복 보존합니다.
    cell.set("t", "str")
    v = ET.SubElement(cell, f"{{{NS}}}v")
    v.text = str(value)


def _get_or_make_cell(row: ET.Element, col: str, row_no: int, style_source: ET.Element | None = None) -> ET.Element:
    ref = f"{col}{row_no}"
    for c in row.findall(f"{{{NS}}}c"):
        if c.get("r") == ref:
            return c
    c = ET.Element(f"{{{NS}}}c", {"r": ref})
    if style_source is not None and style_source.get("s") is not None:
        c.set("s", style_source.get("s"))
    row.append(c)
    return c


def _template_cell(row: ET.Element, col: str) -> ET.Element | None:
    for c in row.findall(f"{{{NS}}}c"):
        if _col_letters(c.get("r", "")) == col:
            return c
    return None


def _clone_data_row(template_row: ET.Element, row_no: int, values: list) -> ET.Element:
    row = copy.deepcopy(template_row)
    row.set("r", str(row_no))
    # remove stray cells beyond H, and rewrite addresses
    for c in list(row.findall(f"{{{NS}}}c")):
        col = _col_letters(c.get("r", ""))
        if col > "H":
            row.remove(c)
            continue
        c.set("r", f"{col}{row_no}")
        _set_cell(c, None)
    for idx, col in enumerate("ABCDEFGH"):
        src = _template_cell(template_row, col)
        c = _get_or_make_cell(row, col, row_no, src)
        _set_cell(c, values[idx] if idx < len(values) else None)
    return row


def _set_named_cell(root: ET.Element, ref: str, value):
    for c in root.findall(f".//{{{NS}}}c"):
        if c.get("r") == ref:
            _set_cell(c, value)
            return
    # not expected for this template


def _set_general_column_widths(root: ET.Element):
    cols = root.find(f"{{{NS}}}cols")
    if cols is None:
        sheet_data = root.find(f"{{{NS}}}sheetData")
        cols = ET.Element(f"{{{NS}}}cols")
        if sheet_data is not None:
            idx = list(root).index(sheet_data)
            root.insert(idx, cols)
        else:
            root.append(cols)
    for child in list(cols):
        cols.remove(child)
    widths = [7.0, 11.0, 13.0, 18.0, 13.0, 18.0, 10.0, 8.0]
    for i, width in enumerate(widths, start=1):
        ET.SubElement(cols, f"{{{NS}}}col", {"min": str(i), "max": str(i), "width": str(width), "customWidth": "1"})


def _patch_sheet(xml_bytes: bytes, records: list[dict], group: str, appoint_date: str, all_new: bool) -> bytes:
    root = ET.fromstring(xml_bytes)
    if not all_new:
        _set_general_column_widths(root)
    sheet_data = root.find(f"{{{NS}}}sheetData")
    if sheet_data is None:
        raise ValueError(f"{group} 시트에 sheetData가 없습니다.")

    rows = list(sheet_data.findall(f"{{{NS}}}row"))
    data_rows = [r for r in rows if int(r.get("r", "0")) >= 7]
    if not data_rows:
        raise ValueError(f"{group} 시트에 데이터 행 서식이 없습니다.")
    first_tpl = data_rows[0]
    mid_tpl = data_rows[1] if len(data_rows) > 1 else data_rows[0]
    last_tpl = data_rows[-1]

    # Existing sample/blank data rows removed, headers retained.
    for r in data_rows:
        sheet_data.remove(r)

    # Header text
    _set_named_cell(root, "A1", "유초등교육공무원 신규교사 인사발령" if all_new else "유초등교육공무원 인사발령")
    section_name = GROUP_LABELS[group][1]
    section = f" 1. {section_name}(신규) 임지지정" if all_new else f" 1. {section_name} 인사발령"
    _set_named_cell(root, "A3", section)

    name_counts = Counter(str(r.get("name") or "") for r in records)
    date_cell = format_date(appoint_date, spaced=False)
    n = len(records)
    for i, rec in enumerate(records, start=1):
        if n == 1:
            tpl = last_tpl
        elif i == 1:
            tpl = first_tpl
        elif i == n:
            tpl = last_tpl
        else:
            tpl = mid_tpl

        job = str(rec.get("job") or "")
        pos = official_position(job, group)
        reason = str(rec.get("reason") or "").strip()
        from_school = str(rec.get("fromSchool") or "").strip()
        to_school = str(rec.get("toSchool") or "").strip()
        is_new = "신규" in reason or (not from_school and "신규" in str(rec.get("appointType") or ""))
        appoint_type = str(rec.get("appointType") or "").strip()

        if is_new:
            current_pos, current_dept = "신규 발령", ""
        elif appoint_type in ("복직", "복귀"):
            current_pos, current_dept = appoint_type, from_school
        elif from_school:
            current_pos, current_dept = pos, from_school
        else:
            current_pos, current_dept = reason or "전입", ""

        birth = str(rec.get("birth") or "").strip()
        note = birth if name_counts[str(rec.get("name") or "")] > 1 and birth else ""
        values = [i, rec.get("name") or "", pos, to_school, current_pos, current_dept, date_cell, note]
        sheet_data.append(_clone_data_row(tpl, 6 + i, values))

    last_row = max(6, 6 + n)
    dim = root.find(f"{{{NS}}}dimension")
    if dim is not None:
        dim.set("ref", f"A1:H{last_row}")

    # 신규 전용이면 원본처럼 E:F를 병합하고, 일반 인사발령이면 현직 직위/부서를 분리합니다.
    merges = root.find(f"{{{NS}}}mergeCells")
    if merges is not None:
        for mc in list(merges.findall(f"{{{NS}}}mergeCell")):
            ref = mc.get("ref", "")
            nums = [int(x) for x in re.findall(r"\d+", ref)]
            if nums and max(nums) >= 7:
                merges.remove(mc)
        if all_new:
            for rr in range(7, 7 + n):
                ET.SubElement(merges, f"{{{NS}}}mergeCell", {"ref": f"E{rr}:F{rr}"})
        merges.set("count", str(len(merges.findall(f"{{{NS}}}mergeCell"))))

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _patch_cover(xml_bytes: bytes, counts: dict[str, int], all_new: bool) -> bytes:
    root = ET.fromstring(xml_bytes)
    labels = {
        "초등": "초등학교교사 신규임용" if all_new else "초등학교교사 인사발령",
        "유치원": "유치원교사 신규임용" if all_new else "유치원교사 인사발령",
        "전문상담": "전문상담교사 신규임용" if all_new else "전문상담교사 인사발령",
        "보건": "보건교사 신규임용" if all_new else "보건교사 인사발령",
    }
    for i, group in enumerate(("초등", "유치원", "전문상담", "보건"), start=16):
        _set_named_cell(root, f"C{i}", labels[group])
        _set_named_cell(root, f"D{i}", int(counts.get(group, 0)))
    _set_named_cell(root, "D20", sum(int(counts.get(g, 0)) for g in counts))
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _patch_drawing(xml_bytes: bytes, appoint_date: str, all_new: bool) -> bytes:
    root = ET.fromstring(xml_bytes)
    paras = root.findall(f".//{{{NS_A}}}p")
    if paras:
        texts = paras[0].findall(f".//{{{NS_A}}}t")
        if texts:
            texts[0].text = format_date(appoint_date, spaced=True) + " "
            if len(texts) >= 2:
                texts[1].text = "자 신규교사 인사발령" if all_new else "자 유초등교육공무원 인사발령"
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _remove_calcchain(files: dict[str, bytes]):
    if "xl/calcChain.xml" not in files:
        return
    files.pop("xl/calcChain.xml", None)
    rel_path = "xl/_rels/workbook.xml.rels"
    if rel_path in files:
        root = ET.fromstring(files[rel_path])
        for rel in list(root):
            if rel.get("Type", "").endswith("/calcChain"):
                root.remove(rel)
        files[rel_path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    ct_path = "[Content_Types].xml"
    if ct_path in files:
        root = ET.fromstring(files[ct_path])
        for node in list(root):
            if node.get("PartName") == "/xl/calcChain.xml":
                root.remove(node)
        files[ct_path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)


def build_notice_xlsx(template_path: str, records: list[dict], appoint_date: str) -> tuple[bytes, dict]:
    if not records:
        raise ValueError("통지서에 출력할 발령 대상자가 없습니다.")

    grouped = {g: [] for g in SHEET_PATHS}
    unsupported = []
    for rec in records:
        group = group_for_job(rec.get("job", ""))
        if not group:
            unsupported.append(str(rec.get("job") or "미지정"))
        else:
            grouped[group].append(rec)
    if unsupported:
        kinds = ", ".join(sorted(set(unsupported)))
        raise ValueError(f"현재 공식 통지서 서식에 없는 직종이 있습니다: {kinds}")

    # Stable output ordering: destination school -> priority -> name
    def sk(r):
        try:
            rank = int(str(r.get("rank") or "999999"))
        except Exception:
            rank = 999999
        return (str(r.get("toSchool") or ""), rank, str(r.get("name") or ""))
    for g in grouped:
        grouped[g].sort(key=sk)

    all_new = all("신규" in str(r.get("reason") or "") for r in records)
    counts = {g: len(grouped[g]) for g in grouped}

    with zipfile.ZipFile(template_path, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    files["xl/worksheets/sheet1.xml"] = _patch_cover(files["xl/worksheets/sheet1.xml"], counts, all_new)
    for group, path in SHEET_PATHS.items():
        group_all_new = bool(grouped[group]) and all("신규" in str(r.get("reason") or "") for r in grouped[group])
        files[path] = _patch_sheet(files[path], grouped[group], group, appoint_date, group_all_new)
    if "xl/drawings/drawing1.xml" in files:
        files["xl/drawings/drawing1.xml"] = _patch_drawing(files["xl/drawings/drawing1.xml"], appoint_date, all_new)
    _remove_calcchain(files)

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    meta = {"counts": counts, "total": len(records), "all_new": all_new}
    return out.getvalue(), meta
