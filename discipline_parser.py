# -*- coding: utf-8 -*-
"""
discipline_parser.py — 기계·건축·전기 분야 내역서 파서 ('공종별내역서형')

적산 프로그램에서 내려받은 분야 내역서는 대부분 같은 형식이다(샘플 14개 중 13개):
  - '공종별내역서' 시트: 품명·규격·단위·수량·금액 + 숨은 보조열
      품목코드(연결된 일위대가 블록 코드, 합계행은 'TOTAL'),
      공종코드(2자리씩 계층: 010101 → 01·01·01), 일위·단산·자재 플래그(T/F)
  - '일위대가' 시트: 블록 머리행(열 '일위대가'에 블록 코드, 수량 없음) + 구성 행
      직종별 인수는 단위가 '인'인 행이다('건축목공 | 0.15 | 인' 또는 '노무비 | 용접공 | 인').
  - '공종별집계표' 시트: 공종코드 계층별 이름(건물 동 / 설비 단위)
열 순서(단위·수량, 재료비·노무비)는 파일마다 달라 머리행 글자로 찾는다.

1조 일수(항목을 1개 조가 끝내는 데 걸리는 일수) 산정:
  - 연결된 일위대가 블록에 직종별 인수가 있으면: 1조 = 직종별 1인씩.
      1조 일수 = 수량 × max(직종별 인/단위). 다른 블록을 참조하는 행은 수량을 곱해 따라 들어간다.
  - 인수가 없고 노무비만 있으면: 노무비 ÷ 기준 노임(파일에서 관측한 노임의 중앙값) — 노무비 역산(1인 기준)
  - 노무비가 없으면(자재·기자재 구매) 공기에 반영하지 않는다.
장비 사용료 블록 안의 운전사 행은 시간당 단가('1인'이지만 1시간분)라 인수로 세지 않는다
(노임 단가가 MIN_DAILY_WAGE 미만인 행 제외).
"""

from __future__ import annotations

import math
import re
from collections import defaultdict
from statistics import median
from typing import Any, Dict, List, Optional

MIN_DAILY_WAGE = 100_000        # 일 노임으로 볼 최소 단가(원). 이보다 작으면 시간당 운전사 행 등으로 본다
DEFAULT_WAGE = 200_000          # 노임을 하나도 못 찾았을 때 쓰는 기준 노임

# 상위 공종명이 이런 '비용 구분'이면(기자재설치비·배관자재비 등) 설비 단위가 아니라
# 하위 공종이 설비 단위다(화성 기계: '1.1.2 기자재설치비' 아래 '침사지 및 유량조정').
_COST_TYPE_RE = re.compile(r"기자재|설치비|배관|자재비|철거비|지지대|잡철물|공사비|도급|관급")
_LEAD_NUM_RE = re.compile(r"^(\d+(?:\.\d+)*)\s*[.)]?\s*")

_HDR = {
    "name": ("품명", "공종명", "명칭", "공종"),
    "spec": ("규격",),
    "unit": ("단위",),
    "qty": ("수량",),
    "labor": ("노무비",),
    "code": ("품목코드",),
    "gcode": ("공종코드",),
    "f_il": ("일위",),
    "f_mat": ("자재",),
    "blk": ("일위대가",),
}


def _s(v) -> str:
    return "" if v is None else str(v).strip()


def _n(v) -> str:
    return re.sub(r"\s+", "", _s(v))


def _f(v) -> float:
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", ""))
    except (TypeError, ValueError):
        return 0.0


def _cell(row, col):
    return row[col] if col is not None and col < len(row) else None


def _clean_name(nm: str) -> str:
    """'010101  가  설  공  사' → '가설공사', '1.  유입 및  침사 설비' → '1. 유입 및 침사 설비'."""
    nm = re.sub(r"^\d{4,}\s+", "", _s(nm))
    toks = nm.split()
    if len(toks) > 1 and all(len(t) == 1 for t in toks):
        return "".join(toks)
    return " ".join(toks)


def _header_map(ws, max_row: int = 6) -> Optional[Dict[str, int]]:
    """머리행 글자로 열 위치를 찾는다. 노무비는 (단가, 금액) 두 칸이라 금액 열을 쓴다."""
    for row in ws.iter_rows(max_row=max_row, values_only=True):
        m: Dict[str, int] = {}
        for i, v in enumerate(row):
            t = _n(v)
            for k, words in _HDR.items():
                if k not in m and t in words:
                    m[k] = i + 1 if k == "labor" else i
        if "name" in m and "qty" in m:
            return m
    return None


def detect_gjb(wb) -> bool:
    """'공종별내역서형' 분야 내역서인지 (시트명 기준)."""
    names = set(wb.sheetnames)
    return "공종별내역서" in names and "일위대가" in names


def _group_names(wb) -> (Dict[str, str], Dict[str, str]):
    """공종코드 → 이름, 앞 번호('1.2') → 이름. 공종별집계표를 우선하고 공종별내역서 머리행으로 보충."""
    by_code: Dict[str, str] = {}
    by_num: Dict[str, str] = {}
    sheets = [s for s in wb.sheetnames if _n(s).startswith("공종별집계표")] + ["공종별내역서"]
    for sn in sheets:
        ws = wb[sn]
        h = _header_map(ws) or {"name": 0}
        for row in ws.iter_rows(values_only=True):
            raw = _s(_cell(row, h["name"]))
            if not raw or raw.startswith("["):
                continue
            code = _s(_cell(row, h.get("gcode")))
            m = re.match(r"^(\d{4,})\s+", raw)
            if not code and m:
                code = m.group(1)
            name = _clean_name(raw)
            if code.isdigit():
                by_code.setdefault(code, name)
            mn = _LEAD_NUM_RE.match(name)
            if mn and "." in name[: mn.end()] + ".":
                by_num.setdefault(mn.group(1), name)
    return by_code, by_num


def _unit_name(gcode: str, by_code: Dict[str, str], by_num: Dict[str, str]) -> (str, str):
    """항목의 (공종명, 설비·동 단위명). 상위 공종이 비용 구분이면 하위 공종을 단위로 쓴다."""
    leaf = by_code.get(gcode, gcode or "(미분류)")
    parent = by_code.get(gcode[:-2]) if len(gcode) > 2 else None
    if parent is None:
        mn = _LEAD_NUM_RE.match(leaf)
        if mn and "." in mn.group(1):
            parent = by_num.get(mn.group(1).rsplit(".", 1)[0])
    if not parent or _COST_TYPE_RE.search(parent):
        return leaf, leaf
    return leaf, parent


def _parse_blocks(ws) -> (Dict[str, Dict[str, Any]], List[float]):
    """일위대가 블록: {블록코드: {"name", "rows": [...]}}, 관측한 일 노임 목록."""
    h = _header_map(ws)
    blocks: Dict[str, Dict[str, Any]] = {}
    wages: List[float] = []
    if not h or "blk" not in h:
        return blocks, wages
    cur = None
    for row in ws.iter_rows(values_only=True):
        nm = _s(_cell(row, h["name"]))
        bc = _s(_cell(row, h["blk"]))
        q = _cell(row, h["qty"])
        if not nm and not bc:
            continue
        if bc == "TOTAL" or nm.startswith("["):
            cur = None
            continue
        if bc and q in (None, ""):
            cur = blocks.setdefault(bc, {"name": _clean_name(nm), "rows": []})
            continue
        if cur is None or not nm:
            continue
        qty = _f(q)
        unit = _s(_cell(row, h.get("unit")))
        lab = _f(_cell(row, h.get("labor")))
        cur["rows"].append({
            "name": nm, "spec": _s(_cell(row, h.get("spec"))), "unit": unit, "qty": qty,
            "ref": _s(_cell(row, h.get("code"))), "il": _s(_cell(row, h.get("f_il"))), "labor": lab,
        })
        if unit == "인" and qty > 0 and lab / qty >= MIN_DAILY_WAGE:
            wages.append(lab / qty)
    return blocks, wages


def _block_trades(code: str, blocks, memo, stack=()) -> Dict[str, float]:
    """블록 1단위당 직종별 인수. 다른 블록을 참조하는 행은 수량을 곱해 합산한다."""
    if code in memo:
        return memo[code]
    b = blocks.get(code)
    if not b or code in stack or len(stack) > 6:
        return {}
    trades: Dict[str, float] = defaultdict(float)
    for r in b["rows"]:
        if r["unit"] == "인":
            if r["qty"] > 0 and r["labor"] / r["qty"] >= MIN_DAILY_WAGE:
                # '노무비 | 용접공'처럼 품명이 '노무비'면 규격 칸이 직종이다
                # (화성 기계는 '노무비(22년상반기)'라 startswith로 판정)
                trade = r["spec"] if _n(r["name"]).startswith("노무비") else r["name"]
                trades[_n(trade)] += r["qty"]
        elif r["il"] == "T" and r["ref"] and r["ref"] != code and r["ref"] in blocks:
            for t, v in _block_trades(r["ref"], blocks, memo, stack + (code,)).items():
                trades[t] += v * r["qty"]
    memo[code] = dict(trades)
    return memo[code]


def parse_gjb(wb, discipline: str = "") -> Dict[str, Any]:
    """분야 내역서 파싱. 반환: {"items": [...], "base_wage": 원, "units": [단위명...]}.

    항목 필드: discipline, unit_name(설비·동), group(공종), name, spec, unit, qty,
              labor(노무비 금액), days1(1조 일수, float), basis(산정 근거), trades(직종별 인/단위)
    """
    ws = wb["공종별내역서"]
    h = _header_map(ws)
    if not h:
        return {"items": [], "base_wage": DEFAULT_WAGE, "units": []}
    by_code, by_num = _group_names(wb)
    blocks, wages = _parse_blocks(wb["일위대가"])
    base_wage = median(wages) if wages else DEFAULT_WAGE
    memo: Dict[str, Dict[str, float]] = {}
    items: List[Dict[str, Any]] = []
    units: List[str] = []
    cur_gcode = ""
    for row in ws.iter_rows(values_only=True):
        nm = _s(_cell(row, h["name"]))
        code = _s(_cell(row, h.get("code")))
        gcode = _s(_cell(row, h.get("gcode")))
        if gcode and not code:
            # 공종 머리행. 항목 행의 공종코드가 비어 있는 파일이 있어(화성 건축기계설비)
            # 직전 머리행의 공종코드를 이어받는다.
            cur_gcode = gcode
            continue
        if not nm or not code or code == "TOTAL":
            continue
        qty = _f(_cell(row, h["qty"]))
        if qty <= 0:
            continue
        gcode = gcode or cur_gcode
        group, unit_name = _unit_name(gcode, by_code, by_num)
        labor = _f(_cell(row, h.get("labor")))
        trades = _block_trades(code, blocks, memo) if _s(_cell(row, h.get("f_il"))) == "T" else {}
        if trades:
            days1 = qty * max(trades.values())
            basis = "일위대가 인수"
        elif labor > 0:
            days1 = labor / base_wage
            basis = "노무비 역산"
        else:
            days1 = 0.0
            basis = "자재(공기 제외)"
        if unit_name not in units:
            units.append(unit_name)
        items.append({
            "discipline": discipline, "unit_name": unit_name, "group": group,
            "name": nm, "spec": _s(_cell(row, h.get("spec"))), "unit": _s(_cell(row, h.get("unit"))),
            "qty": qty, "labor": labor, "days1": days1, "basis": basis, "trades": trades,
        })
    return {"items": items, "base_wage": base_wage, "units": units}


def package_key(discipline: str, unit_name: str, group: str) -> str:
    return f"{discipline}|{unit_name}|{group}"


def build_packages(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """작업 묶음(설비·동 단위 안의 공종)별 직종 작업량(인·일)을 모은다.

    같은 공종 안에서 직종이 다른 항목(형틀목공·철근공 등)은 동시에 진행되므로, 묶음 일수는
    작업량이 가장 많은 직종이 정한다. 모든 항목을 순차로 더하면 건축이 수천 일로 부풀었다
    (실측: 화성 건축 6,916일). 직종 정보가 없는 노무비 역산 항목은 '노무비역산' 한 직종으로 본다.
    """
    pk: Dict[tuple, Dict[str, Any]] = {}
    for it in items:
        if it["days1"] <= 0:
            continue
        p = pk.setdefault((it["unit_name"], it["group"]), {
            "discipline": it["discipline"], "unit_name": it["unit_name"], "group": it["group"],
            "trades": defaultdict(float), "n_items": 0,
        })
        p["n_items"] += 1
        if it["trades"]:
            for t, v in it["trades"].items():
                p["trades"][t] += it["qty"] * v
        else:
            p["trades"]["노무비역산"] += it["days1"]
    return list(pk.values())


def discipline_days(items: List[Dict[str, Any]], mode: str, crews: int = 1,
                    crew_by_pkg: Optional[Dict[str, int]] = None) -> Dict[str, Any]:
    """분야 작업일수.

    작업 묶음(공종) 일수 = ceil(최대 직종 작업량 ÷ 투입조수)   — 1조 = 직종별 1인
    설비·동 단위 일수    = 단위 안 공종 일수의 합(공종 순차)
    mode "seq"          : 단위끼리도 순차(건축 — 동도 순차)
    mode "unit_parallel": 단위끼리 병행(기계·전기 — 설비 단위별 병행)
    """
    crew_by_pkg = crew_by_pkg or {}
    per_unit: Dict[str, int] = defaultdict(int)
    rows = []
    for p in build_packages(items):
        lead, w = max(p["trades"].items(), key=lambda x: x[1])
        c = max(1, int(crew_by_pkg.get(package_key(p["discipline"], p["unit_name"], p["group"]), crews)))
        d = math.ceil(w / c - 1e-9)
        per_unit[p["unit_name"]] += d
        rows.append({"discipline": p["discipline"], "unit_name": p["unit_name"], "group": p["group"],
                     "n_items": p["n_items"], "lead_trade": lead, "workload": round(w, 1),
                     "crews": c, "days": d})
    total = sum(per_unit.values()) if mode == "seq" else max(per_unit.values(), default=0)
    return {"per_unit": dict(per_unit), "packages": rows, "total": int(total)}
