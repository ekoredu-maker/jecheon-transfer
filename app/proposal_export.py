# -*- coding: utf-8 -*-
"""제천 내신전보 - 공식 초등교사 전의안 출력기 v3.1.5.

핵심 원칙
1) 학교기본정보(schools)가 학교명/학교급/급지/정렬의 유일한 마스터다.
2) 공식 XLSX는 '서식'만 제공한다. 템플릿 내부 학교명은 출력 기준으로 사용하지 않는다.
3) 원본 OOXML의 네임스페이스/스타일/병합/인쇄설정을 보존하며 셀 값만 교체한다.
4) stale calcChain을 제거하고, 새 문자열은 inlineStr로 기록하여 Excel 복구 경고를 방지한다.
"""
from __future__ import annotations

import io
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_X14AC = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
NS_XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
NS_XR2 = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
NS_XR3 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"
XML_NS = "http://www.w3.org/XML/1998/namespace"

# ElementTree 재직렬화 시 Excel 원본 prefix가 사라지지 않게 고정.
ET.register_namespace("", NS)
ET.register_namespace("r", NS_R)
ET.register_namespace("mc", NS_MC)
ET.register_namespace("x14ac", NS_X14AC)
ET.register_namespace("xr", NS_XR)
ET.register_namespace("xr2", NS_XR2)
ET.register_namespace("xr3", NS_XR3)

DEP_KEYS = [
    "승진전직", "정퇴명퇴면직", "시도간파견출", "국립전출", "석사행정파견",
    "학습연구년", "일방전출", "타시군전출", "관내전출", "휴직",
]
ARR_KEYS = [
    "복직", "관내전입", "타시군전입", "시도간파견입",
    "행정석사파견", "국립전입", "비정기수석", "신규",
]
DEP_COLS = dict(zip(DEP_KEYS, ["I", "K", "M", "O", "Q", "S", "U", "W", "Y", "AA"]))
DEP_CNT_COLS = dict(zip(DEP_KEYS, ["J", "L", "N", "P", "R", "T", "V", "X", "Z", "AB"]))
ARR_COLS = dict(zip(ARR_KEYS, ["AK", "AM", "AO", "AQ", "AS", "AU", "AW", "AY"]))
ARR_CNT_COLS = dict(zip(ARR_KEYS, ["AL", "AN", "AP", "AR", "AT", "AV", "AX", "AZ"]))

DATA_START_ROW = 7
DATA_END_ROW = 29  # 사용자 제공 제천 공식서식: 23개 학교 행


def _plain(v) -> str:
    return re.sub(r"\s+", "", str(v or "").strip())


def _school_token(v: str) -> str:
    """매칭 전용 토큰. 출력명 자체를 바꾸는 용도로 사용하지 않는다."""
    s = _plain(v)
    s = s.replace("초등학교병설유치원", "초병설유치원")
    s = s.replace("초등학교", "초")
    s = s.replace("초중학교", "초")
    return s


def _school_aliases(v: str) -> set[str]:
    raw = _plain(v)
    if not raw:
        return set()
    vals = {raw, _school_token(raw)}
    # 제천 지역 내에서만 쓰는 전의안이므로 '제천' 접두어 차이는 안전한 별칭 후보로 등록.
    for x in list(vals):
        if x.startswith("제천") and len(x) > 2:
            vals.add(x[2:])
    # 표기 흔들림 보정
    for x in list(vals):
        vals.add(x.replace("초등학교", "초"))
        vals.add(x.replace("초등학교", ""))
    return {x for x in vals if x}


def _num(v):
    try:
        if v is None or v == "":
            return 0.0
        return float(str(v).replace(",", ""))
    except Exception:
        return 0.0


def _intish(v):
    n = _num(v)
    return int(n) if float(n).is_integer() else n


def _date_parts(raw: str):
    m = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", str(raw or ""))
    if not m:
        return None
    return tuple(map(int, m.groups()))


def _col_num_from_letters(s: str) -> int:
    n = 0
    for ch in s:
        if "A" <= ch <= "Z":
            n = n * 26 + ord(ch) - 64
    return n


def _col_num(ref: str) -> int:
    m = re.match(r"([A-Z]+)", ref or "")
    return _col_num_from_letters(m.group(1)) if m else 0


def _set_cell(cell: ET.Element, value):
    """셀의 스타일(s)은 유지하고 값/수식만 교체."""
    for ch in list(cell):
        if ch.tag in (f"{{{NS}}}v", f"{{{NS}}}f", f"{{{NS}}}is"):
            cell.remove(ch)
    if value is None or value == "":
        cell.attrib.pop("t", None)
        return
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        cell.attrib.pop("t", None)
        v = ET.SubElement(cell, f"{{{NS}}}v")
        v.text = str(value)
        return
    # 'str' + <v>는 Excel에서 복구 대상이 될 수 있어 inlineStr 사용.
    cell.set("t", "inlineStr")
    is_el = ET.SubElement(cell, f"{{{NS}}}is")
    t_el = ET.SubElement(is_el, f"{{{NS}}}t")
    text = str(value)
    if text[:1].isspace() or text[-1:].isspace() or "\n" in text:
        t_el.set(f"{{{XML_NS}}}space", "preserve")
    t_el.text = text


def _find_cell(root: ET.Element, ref: str) -> ET.Element | None:
    return root.find(f".//{{{NS}}}c[@r='{ref}']")


def _ensure_cell(root: ET.Element, ref: str) -> ET.Element:
    c = _find_cell(root, ref)
    if c is not None:
        return c
    row_no = int(re.search(r"\d+$", ref).group())
    sheet_data = root.find(f"{{{NS}}}sheetData")
    if sheet_data is None:
        raise ValueError("전의안 템플릿 sheetData가 없습니다.")
    row = root.find(f".//{{{NS}}}row[@r='{row_no}']")
    if row is None:
        row = ET.Element(f"{{{NS}}}row", {"r": str(row_no)})
        inserted = False
        for idx, existing in enumerate(list(sheet_data)):
            if int(existing.get("r") or 0) > row_no:
                sheet_data.insert(idx, row)
                inserted = True
                break
        if not inserted:
            sheet_data.append(row)
    c = ET.Element(f"{{{NS}}}c", {"r": ref})
    target_col = _col_num(ref)
    inserted = False
    for idx, existing in enumerate(row.findall(f"{{{NS}}}c")):
        if _col_num(existing.get("r") or "") > target_col:
            row.insert(idx, c)
            inserted = True
            break
    if not inserted:
        row.append(c)
    return c


def _set_ref(root: ET.Element, ref: str, value):
    _set_cell(_ensure_cell(root, ref), value)


def _clear_existing_row_values(root: ET.Element, row_no: int):
    row = root.find(f".//{{{NS}}}row[@r='{row_no}']")
    if row is None:
        return
    for c in row.findall(f"{{{NS}}}c"):
        _set_cell(c, None)


def _shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    out = []
    for si in root.findall(f"{{{NS}}}si"):
        out.append("".join((t.text or "") for t in si.findall(f".//{{{NS}}}t")))
    return out


def _cell_value(cell: ET.Element | None, shared: list[str]):
    if cell is None:
        return None
    if cell.get("t") == "inlineStr":
        return "".join((t.text or "") for t in cell.findall(f".//{{{NS}}}t"))
    v = cell.find(f"{{{NS}}}v")
    if v is None or v.text is None:
        return None
    if cell.get("t") == "s":
        try:
            return shared[int(v.text)]
        except Exception:
            return v.text
    if cell.get("t") == "str":
        return v.text
    try:
        n = float(v.text)
        return int(n) if n.is_integer() else n
    except Exception:
        return v.text


def _join_names(items: list[dict]) -> str:
    vals = []
    for x in items:
        s = str(x.get("display") or x.get("name") or "").strip()
        if s:
            vals.append(s)
    return ", ".join(vals)


def _detail_vals(row: dict, prefix: str):
    keys = ["class", "subject", "other", "senior"] if prefix == "prev" else ["class", "subject", "other", "over", "senior"]
    vals = [row.get(f"{prefix}_{k}") for k in keys]
    has = any(v is not None and v != "" for v in vals)
    return has, vals


def _is_elementary_master(s: dict) -> bool:
    level = _plain(s.get("schoolLevel") or s.get("학교급"))
    name = _plain(s.get("schoolName") or s.get("학교명"))
    if level:
        return "초" in level and "유치" not in level
    return "초" in name and "유치" not in name


def _build_master(payload: dict):
    raw = payload.get("schools") or []
    if not isinstance(raw, list) or not raw:
        raise ValueError("학교기본정보가 없습니다. 전의안은 학교기본정보를 기준으로 생성합니다.")

    masters = []
    seen_name = set()
    seen_code = set()
    for i, s in enumerate(raw):
        if not isinstance(s, dict) or not _is_elementary_master(s):
            continue
        name = str(s.get("schoolName") or s.get("학교명") or "").strip()
        if not name:
            continue
        code = str(s.get("schoolCode") or s.get("학교코드") or "").strip()
        zone = str(s.get("zone") or s.get("급지") or "").strip()
        display = str(s.get("proposalName") or s.get("전의안표시명") or name).strip() or name
        order_raw = s.get("sortOrder") if s.get("sortOrder") not in (None, "") else s.get("정렬순서")
        try:
            order = float(order_raw)
        except Exception:
            order = i + 1
        nk = _plain(name)
        if nk in seen_name:
            raise ValueError(f"학교기본정보 학교명이 중복됩니다: {name}")
        seen_name.add(nk)
        if code:
            ck = _plain(code)
            if ck in seen_code:
                raise ValueError(f"학교기본정보 학교코드가 중복됩니다: {code}")
            seen_code.add(ck)
        masters.append({
            "key": code or f"NAME:{nk}", "code": code, "name": name, "display": display,
            "zone": zone, "order": order, "sourceOrder": i,
        })

    masters.sort(key=lambda x: (x["order"], x["sourceOrder"]))
    if not masters:
        raise ValueError("학교기본정보에서 초등학교를 찾지 못했습니다.")
    capacity = DATA_END_ROW - DATA_START_ROW + 1
    if len(masters) > capacity:
        raise ValueError(f"공식 전의안 학교행은 {capacity}교까지입니다. 학교기본정보에 초등학교가 {len(masters)}교 있습니다.")

    alias_to_key: dict[str, str | None] = {}
    by_key = {m["key"]: m for m in masters}

    def add(alias: str, key: str):
        a = _school_token(alias)
        if not a:
            return
        if a in alias_to_key and alias_to_key[a] != key:
            alias_to_key[a] = None
        else:
            alias_to_key[a] = key

    for m in masters:
        for a in _school_aliases(m["name"]) | _school_aliases(m["display"]):
            add(a, m["key"])
        if m["code"]:
            add(m["code"], m["key"])

    def resolve(v: str, *, allow_blank: bool = False) -> str | None:
        if not _plain(v):
            return None if allow_blank else ""
        candidates = list(_school_aliases(v))
        hits = {alias_to_key.get(_school_token(a)) for a in candidates}
        hits.discard(None)
        if len(hits) == 1:
            return next(iter(hits))
        return ""

    return masters, by_key, resolve


def _patch_package_metadata(name: str, data: bytes) -> bytes | None:
    """stale calcChain 제거.

    Content_Types/relationships는 ElementTree로 다시 직렬화하면 기본 namespace가 ns0 prefix로
    바뀌어 일부 XLSX 파서에서 열기 오류가 발생할 수 있다. 따라서 원본 XML 문자열을 유지한 채
    calcChain 항목만 제거한다.
    """
    if name == "xl/calcChain.xml":
        return None
    if name == "xl/_rels/workbook.xml.rels":
        text = data.decode("utf-8")
        text = re.sub(r'<Relationship\b(?=[^>]*(?:Type="[^"]*/calcChain"|Target="calcChain\.xml"))[^>]*/>', '', text)
        return text.encode("utf-8")
    if name == "[Content_Types].xml":
        text = data.decode("utf-8")
        text = re.sub(r'<Override\b(?=[^>]*PartName="/xl/calcChain\.xml")[^>]*/>', '', text)
        return text.encode("utf-8")
    return data


def _restore_required_namespace_decls(xml_bytes: bytes) -> bytes:
    """ElementTree가 mc:Ignorable에만 등장하는 xr2/xr3 xmlns를 버리는 문제 보정."""
    s = xml_bytes.decode("utf-8")
    decls = {
        "mc": NS_MC,
        "x14ac": NS_X14AC,
        "xr": NS_XR,
        "xr2": NS_XR2,
        "xr3": NS_XR3,
    }
    pos = s.find("<worksheet")
    end = s.find(">", pos)
    if pos < 0 or end < 0:
        return xml_bytes
    head = s[pos:end]
    extra = ""
    for prefix, uri in decls.items():
        if f"xmlns:{prefix}=" not in head:
            extra += f' xmlns:{prefix}="{uri}"'
    if extra:
        s = s[:end] + extra + s[end:]
    return s.encode("utf-8")


def _validate_sheet_xml(sheet_bytes: bytes):
    root = ET.fromstring(sheet_bytes)
    # 행별 셀은 열 순서대로 정렬되어야 함.
    for row in root.findall(f".//{{{NS}}}row"):
        refs = [c.get("r") or "" for c in row.findall(f"{{{NS}}}c")]
        nums = [_col_num(r) for r in refs]
        if nums != sorted(nums):
            raise ValueError(f"전의안 XLSX 내부 셀 순서 오류: {row.get('r')}행")
    # 수식 없는 t=str 셀 금지(문자열은 inlineStr 사용).
    for c in root.findall(f".//{{{NS}}}c[@t='str']"):
        if c.find(f"{{{NS}}}f") is None:
            raise ValueError(f"전의안 XLSX 문자열 셀 형식 오류: {c.get('r')}")


def build_proposal_xlsx(template_path: str, payload: dict) -> tuple[bytes, dict]:
    appoint_date = str(payload.get("appointDate") or "").strip()
    quota = payload.get("quota") or []
    departures = payload.get("departures") or []
    arrivals = payload.get("arrivals") or []
    separate = payload.get("separate") or []
    if not isinstance(quota, list):
        raise ValueError("quota 형식 오류")

    masters, by_key, resolve = _build_master(payload)

    with zipfile.ZipFile(template_path, "r") as zin:
        shared = _shared_strings(zin)
        root = ET.fromstring(zin.read("xl/worksheets/sheet1.xml"))

        # 템플릿 내부 학교명은 '기준'으로 쓰지 않음. 단, 정원 세부항목 fallback을 위한 참고값으로만 안전하게 매칭.
        template_detail_by_key = {}
        for rr in range(DATA_START_ROW, DATA_END_ROW + 1):
            tname = _cell_value(_find_cell(root, f"B{rr}"), shared)
            key = resolve(str(tname or ""), allow_blank=True)
            if not key or key in template_detail_by_key:
                continue
            template_detail_by_key[key] = {
                "prev": [_cell_value(_find_cell(root, f"{c}{rr}"), shared) for c in ["C", "D", "E", "F"]],
                "q228": _cell_value(_find_cell(root, f"G{rr}"), shared),
                "curq": [_cell_value(_find_cell(root, f"{c}{rr}"), shared) for c in ["AD", "AE", "AF", "AG", "AH"]],
                "q301": _cell_value(_find_cell(root, f"AI{rr}"), shared),
            }

        # qmap / 인사 이동을 모두 학교마스터 key로 변환. 미매칭을 조용히 버리지 않음.
        qmap = defaultdict(lambda: {
            "q228": 0, "current228": 0, "q301": 0,
            "prev_class": None, "prev_subject": None, "prev_other": None, "prev_senior": None,
            "cur_class": None, "cur_subject": None, "cur_other": None, "cur_over": None, "cur_senior": None,
            "temporary": 0,
        })
        unmatched = []
        missing_current = []
        for q in quota:
            raw_school = q.get("school") or ""
            key = resolve(raw_school)
            if not key:
                unmatched.append(f"정원:{raw_school or '(빈 학교명)'}")
                continue
            x = qmap[key]
            x["q228"] += _num(q.get("q228"))
            if q.get("current228") in (None, ""):
                missing_current.append(by_key[key]["name"])
            else:
                x["current228"] += _num(q.get("current228"))
            x["q301"] += _num(q.get("q301"))
            x["temporary"] += _num(q.get("temporary"))
            for field in ["prev_class", "prev_subject", "prev_other", "prev_senior", "cur_class", "cur_subject", "cur_other", "cur_over", "cur_senior"]:
                v = q.get(field)
                if v is not None and v != "":
                    if x[field] is None:
                        x[field] = 0
                    x[field] += _num(v)

        if missing_current:
            raise ValueError("2.28 실제 현원이 없는 학교가 있습니다: " + ", ".join(sorted(set(missing_current))))

        dmap = defaultdict(lambda: defaultdict(list))
        amap = defaultdict(lambda: defaultdict(list))
        smap = defaultdict(list)
        bad_categories = []
        for d in departures:
            key = resolve(d.get("school") or "")
            cat = str(d.get("category") or "")
            if not key:
                unmatched.append(f"결원:{d.get('name','')} / {d.get('school','')}")
                continue
            if cat not in DEP_COLS:
                bad_categories.append(f"결원:{d.get('name','')}({cat or '미분류'})")
                continue
            dmap[key][cat].append(d)
        for a in arrivals:
            key = resolve(a.get("school") or "")
            cat = str(a.get("category") or "")
            if not key:
                unmatched.append(f"충원:{a.get('name','')} / {a.get('school','')}")
                continue
            if cat not in ARR_COLS:
                bad_categories.append(f"충원:{a.get('name','')}({cat or '미분류'})")
                continue
            amap[key][cat].append(a)
        for a in separate:
            key = resolve(a.get("school") or "")
            if not key:
                unmatched.append(f"별도정원:{a.get('name','')} / {a.get('school','')}")
                continue
            smap[key].append(a)
        if unmatched:
            raise ValueError("학교기본정보와 매칭되지 않는 학교가 있습니다: " + "; ".join(unmatched[:12]))
        if bad_categories:
            raise ValueError("전의안 사유 미분류: " + ", ".join(bad_categories[:10]))

        # 템플릿의 23개 학교 데이터행은 모두 비우고, 학교기본정보 순서로 재작성.
        for rr in range(DATA_START_ROW, DATA_END_ROW + 1):
            _clear_existing_row_values(root, rr)

        datep = _date_parts(appoint_date)
        if datep:
            y, m, d = datep
            _set_ref(root, "A1", f"{y:04d}.{m:02d}.{d:02d}.자 초등교사 전의안 ")
            _set_ref(root, "C3", f"{str(y)[-2:]}.02.28.자 정원")
            _set_ref(root, "AD3", f"{str(y)[-2:]}.{m:02d}.{d:02d}자 정원")
        _set_ref(root, "I2", "※ 학교명·급지·정렬은 학교기본정보 기준 / 결원·충원·현원계·정원대비 자동 산출")

        missing_prev = []
        missing_cur = []
        fallback_prev = []
        fallback_cur = []
        school_results = []

        for idx, master in enumerate(masters):
            rr = DATA_START_ROW + idx
            key = master["key"]
            q = qmap[key]
            q228 = _intish(q["q228"])
            cur228 = _intish(q["current228"])
            q301 = _intish(q["q301"])

            _set_ref(root, f"A{rr}", idx + 1)
            _set_ref(root, f"B{rr}", master["display"])

            original = template_detail_by_key.get(key) or {"prev": [None]*4, "q228": None, "curq": [None]*5, "q301": None}
            has_prev, prev_vals = _detail_vals(q, "prev")
            if not has_prev:
                if _num(original["q228"]) == _num(q228) and _num(q228) != 0:
                    prev_vals = original["prev"]
                    fallback_prev.append(master["name"])
                elif _num(q228) != 0:
                    prev_vals = [None, None, None, None]
                    missing_prev.append(master["name"])
            for c, v in zip(["C", "D", "E", "F"], prev_vals):
                _set_ref(root, f"{c}{rr}", _intish(v) if v not in (None, "") else None)
            _set_ref(root, f"G{rr}", q228)
            _set_ref(root, f"H{rr}", cur228)

            dep_total = 0
            for k in DEP_KEYS:
                items = dmap[key][k]
                _set_ref(root, f"{DEP_COLS[k]}{rr}", _join_names(items))
                _set_ref(root, f"{DEP_CNT_COLS[k]}{rr}", len(items))
                dep_total += len(items)
            _set_ref(root, f"AC{rr}", dep_total)

            has_cur, cur_vals = _detail_vals(q, "cur")
            if not has_cur:
                if _num(original["q301"]) == _num(q301) and _num(q301) != 0:
                    cur_vals = original["curq"]
                    fallback_cur.append(master["name"])
                elif _num(q301) != 0:
                    cur_vals = [None, None, None, None, None]
                    missing_cur.append(master["name"])
            for c, v in zip(["AD", "AE", "AF", "AG", "AH"], cur_vals):
                _set_ref(root, f"{c}{rr}", _intish(v) if v not in (None, "") else None)
            _set_ref(root, f"AI{rr}", q301)

            surplus_shortage = _num(cur228) - dep_total - _num(q301)
            _set_ref(root, f"AJ{rr}", _intish(surplus_shortage))

            arr_total = 0
            for k in ARR_KEYS:
                items = amap[key][k]
                _set_ref(root, f"{ARR_COLS[k]}{rr}", _join_names(items))
                _set_ref(root, f"{ARR_CNT_COLS[k]}{rr}", len(items))
                arr_total += len(items)
            _set_ref(root, f"BA{rr}", arr_total)
            final_current = _num(cur228) - dep_total + arr_total
            _set_ref(root, f"BB{rr}", _intish(final_current))
            tmp = _intish(q.get("temporary") or 0)
            _set_ref(root, f"BC{rr}", tmp if tmp else None)
            _set_ref(root, f"BD{rr}", _intish(final_current + _num(tmp) - _num(q301)))
            _set_ref(root, f"BE{rr}", _join_names(smap[key]))

            school_results.append({
                "key": key, "school": master["name"], "display": master["display"], "zone": master["zone"],
                "departures": dep_total, "arrivals": arr_total, "final": final_current,
                "prev_vals": prev_vals, "cur_vals": cur_vals,
            })

        # 합계행 6: 학교기본정보에 포함된 학교만 합산.
        _set_ref(root, "A6", "계")
        _set_ref(root, "B6", None)
        for idx, col in enumerate(["C", "D", "E", "F"]):
            vals = [r["prev_vals"][idx] for r in school_results]
            _set_ref(root, f"{col}6", _intish(sum(_num(v) for v in vals)) if any(v not in (None, "") for v in vals) else None)
        _set_ref(root, "G6", _intish(sum(_num(qmap[r["key"]]["q228"]) for r in school_results)))
        _set_ref(root, "H6", _intish(sum(_num(qmap[r["key"]]["current228"]) for r in school_results)))
        for k in DEP_KEYS:
            _set_ref(root, f"{DEP_COLS[k]}6", None)
            _set_ref(root, f"{DEP_CNT_COLS[k]}6", sum(len(dmap[r["key"]][k]) for r in school_results))
        total_dep = sum(r["departures"] for r in school_results)
        _set_ref(root, "AC6", total_dep)
        for idx, col in enumerate(["AD", "AE", "AF", "AG", "AH"]):
            vals = [r["cur_vals"][idx] for r in school_results]
            _set_ref(root, f"{col}6", _intish(sum(_num(v) for v in vals)) if any(v not in (None, "") for v in vals) else None)
        total_q301 = sum(_num(qmap[r["key"]]["q301"]) for r in school_results)
        _set_ref(root, "AI6", _intish(total_q301))
        _set_ref(root, "AJ6", _intish(sum(_num(qmap[r["key"]]["current228"]) - r["departures"] - _num(qmap[r["key"]]["q301"]) for r in school_results)))
        for k in ARR_KEYS:
            _set_ref(root, f"{ARR_COLS[k]}6", None)
            _set_ref(root, f"{ARR_CNT_COLS[k]}6", sum(len(amap[r["key"]][k]) for r in school_results))
        total_arr = sum(r["arrivals"] for r in school_results)
        _set_ref(root, "BA6", total_arr)
        total_final = sum(r["final"] for r in school_results)
        _set_ref(root, "BB6", _intish(total_final))
        total_tmp = sum(_num(qmap[r["key"]]["temporary"]) for r in school_results)
        _set_ref(root, "BC6", _intish(total_tmp) if total_tmp else None)
        _set_ref(root, "BD6", _intish(total_final + total_tmp - total_q301))
        _set_ref(root, "BE6", None)

        # 원본의 dimension/print/병합은 그대로 두고 autoFilter만 데이터행 범위에 맞춤.
        auto = root.find(f"{{{NS}}}autoFilter")
        if auto is not None:
            auto.set("ref", f"A1:BE{DATA_END_ROW}")

        patched_sheet = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        patched_sheet = _restore_required_namespace_decls(patched_sheet)
        _validate_sheet_xml(patched_sheet)

        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                name = item.filename
                if name == "xl/worksheets/sheet1.xml":
                    data = patched_sheet
                else:
                    data = zin.read(name)
                data2 = _patch_package_metadata(name, data)
                if data2 is not None:
                    zout.writestr(item, data2)

    # 패키지 자체 재검증
    result = out.getvalue()
    with zipfile.ZipFile(io.BytesIO(result), "r") as zchk:
        for xml_name in ["xl/worksheets/sheet1.xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml"]:
            ET.fromstring(zchk.read(xml_name))
        _validate_sheet_xml(zchk.read("xl/worksheets/sheet1.xml"))
        if "xl/calcChain.xml" in zchk.namelist():
            raise ValueError("전의안 XLSX calcChain 제거 실패")

    meta = {
        "schools": len(masters),
        "schoolNames": [m["name"] for m in masters],
        "schoolZones": {m["name"]: m["zone"] for m in masters},
        "departures": sum(len(v) for s in dmap.values() for v in s.values()),
        "arrivals": sum(len(v) for s in amap.values() for v in s.values()),
        "separate": sum(len(v) for v in smap.values()),
        "missingPrevDetail": missing_prev,
        "missingCurDetail": missing_cur,
        "fallbackPrevDetail": fallback_prev,
        "fallbackCurDetail": fallback_cur,
    }
    return result, meta
