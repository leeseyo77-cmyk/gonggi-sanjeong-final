# -*- coding: utf-8 -*-
"""
boq_adapters.py — 표준형과 같은 정보를 다른 배치로 담은 내역서를 표준형 배치로 바꾸는 변환기

새 내역서 양식이 나오면 파서 전체를 고치지 않고 여기에 변환기 하나를 더한다. 변환한 워크북은
표준형(산근호표) 파이프라인(산근 Q·A·Q 따라가기, 노무비 역산 근거, 굴착기 규격, 상차분)을 그대로 탄다.

단산·일위·표준 참조형(경주 지방상수도 비상공급망 실측, 2026-11-03)
  - 설계내역서: 맨 앞 빈 열 + 공종 | 품명 | 규격 | 수량 | 단위 | 합계(단가·금액) | 노무비 … | 비고
    공종 번호: ◎ 토목공사 → A. 관로공사 → 1. 1-1구간(…) → 1.1 토공 → 1) 포장깨기 → (1) …
    비고: '단산94참조'(단가산출근거) · '일위200참조'(일위대가) · '표준45참조'(표준시장단가)
  - 단가산출근거: 블록 머리 '37.되메우기 및 다짐 (소형장비) [㎥]' (총괄표 No와 같은 번호)
  - 일위대가: '제1호표 | 품명 | 규격 | | 단위 | … 노무비 금액' 머리행 + 구성 행(직종 '인' 행, 'N참조')
  - 표준시장단가: No | 코드 | 품명 | 규격 | 단위 | 합계 | 노무비 …
  - 노임단가: No | 품명 | 규격 | 단위 | 적용단가

비고의 참조 번호는 설계 중 단가산출 번호를 다시 매기며 어긋난 경우가 있다(실측: 단산 참조 226개 중 49개가
다른 품명을 가리킴 — '터파기(토사) 단산96참조'인데 96은 '콘크리트포장깨기'). 그래서 번호가 가리키는 총괄표
품명이 항목과 맞으면 그 번호, 아니면 총괄표에서 품명·규격이 같은 번호를 찾고, 없으면 연결하지 않는다.
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Tuple

import openpyxl

STD_KEY_BASE = 90000          # 표준시장단가 N → 합성 일위대가 키(90000 + N)
DAN_KEY_BASE = 80000          # 단가산출 N → 합성 일위대가 키(80000 + N) — Q값이 없는 단산 블록의 노무비 역산용

_REF_RE = re.compile(r"(단산|일위|표준)\s*(\d+)\s*참조")


def _s(v) -> str:
    return "" if v is None else str(v).strip()


def _k(v) -> str:
    """품명·규격 비교용: 공백·괄호·기호 제거."""
    return re.sub(r"[\s()\[\]{}·,./\-_~∼×*]", "", _s(v)).upper()


def _num(v) -> Optional[float]:
    return float(v) if isinstance(v, (int, float)) else None


# ── 판별 ────────────────────────────────────────────────────────────────
def is_dansan_ref(wb) -> bool:
    names = set(wb.sheetnames)
    if not {"설계내역서", "단가산출근거", "일위대가"} <= names:
        return False
    ws = wb["설계내역서"]
    for n, row in enumerate(ws.iter_rows(values_only=True)):
        if n > 400:
            break
        for v in row:
            if isinstance(v, str) and _REF_RE.search(v):
                return True
    return False


# ── 총괄표 색인 ──────────────────────────────────────────────────────────
class _Index:
    """번호 → (품명, 규격), 그리고 (품명, 규격)·품명 → 번호."""

    def __init__(self):
        self.by_no: Dict[int, Tuple[str, str]] = {}
        self.by_ns: Dict[Tuple[str, str], int] = {}
        self.by_n: Dict[str, List[int]] = {}

    def add(self, no: int, name: str, spec: str):
        if no in self.by_no:
            return
        self.by_no[no] = (name, spec)
        self.by_ns.setdefault((_k(name), _k(spec)), no)
        self.by_n.setdefault(_k(name), []).append(no)

    def resolve(self, ref: Optional[int], name: str, spec: str) -> Optional[int]:
        kn, ks = _k(name), _k(spec)
        if ref is not None and ref in self.by_no:
            rn, rs = self.by_no[ref]
            if _k(rn) == kn or (kn and _k(rn)[:4] == kn[:4] and (not ks or _k(rs) == ks or not _k(rs))):
                return ref
        if (kn, ks) in self.by_ns:
            return self.by_ns[(kn, ks)]
        cands = self.by_n.get(kn) or []
        return cands[0] if len(cands) == 1 else None


def _index_dansan(wb) -> Tuple[_Index, Dict[int, Dict[str, Any]]]:
    """단가산출총괄표: No | 품명 | 규격 | 단위 | 합계 | 노무비 | 재료비 | 경비."""
    ix, rows = _Index(), {}
    if "단가산출총괄표" in wb.sheetnames:
        for r in wb["단가산출총괄표"].iter_rows(values_only=True):
            r = list(r) + [None] * 8
            no = r[1]
            if isinstance(no, (int, float)) or (isinstance(no, str) and no.strip().isdigit()):
                ix.add(int(no), _s(r[2]), _s(r[3]))
                rows[int(no)] = {"name": _s(r[2]), "spec": _s(r[3]), "unit": _s(r[4]), "labor": _num(r[6]) or 0.0}
    return ix, rows


def _index_ilwi(wb) -> _Index:
    ix = _Index()
    for nm in ("일위대가총괄표", "일위대가"):
        if nm not in wb.sheetnames:
            continue
        for r in wb[nm].iter_rows(values_only=True):
            m = re.match(r"^제\s*(\d+)\s*호표", _s(r[1]) if len(r) > 1 else "")
            if m:
                ix.add(int(m.group(1)), _s(r[2]), _s(r[3]))
    return ix


def _std_table(wb) -> Tuple[_Index, Dict[int, Dict[str, Any]]]:
    ix, rows = _Index(), {}
    if "표준시장단가" in wb.sheetnames:
        for r in wb["표준시장단가"].iter_rows(values_only=True):
            no = r[1] if len(r) > 1 else None
            if isinstance(no, (int, float)) or (isinstance(no, str) and no.strip().isdigit()):
                n = int(no)
                ix.add(n, _s(r[3]), _s(r[4]))
                rows[n] = {"name": _s(r[3]), "spec": _s(r[4]), "unit": _s(r[5]), "labor": _num(r[7]) or 0.0}
    return ix, rows


# ── 변환 ────────────────────────────────────────────────────────────────
def adapt_dansan_ref(wb):
    """표준형(산근호표) 배치의 새 워크북. 시트: 설계내역서·단가산출근거·일위대가표·노무비목록표."""
    out = openpyxl.Workbook()
    out.remove(out.active)
    (ix_d, dan_rows), ix_i = _index_dansan(wb), _index_ilwi(wb)
    ix_s, std_rows = _std_table(wb)
    used_std, used_dan = set(), set()

    def link(text: str, name: str, spec: str) -> str:
        """비고 '단산N참조' → '산근 K호표', '일위N참조' → '대가 K호표', '표준N참조' → '대가 9000N호표'."""
        m = _REF_RE.search(text or "")
        if not m:
            return ""
        kind, n = m.group(1), int(m.group(2))
        if kind == "단산":
            k = ix_d.resolve(n, name, spec)
            if k is None:
                return ""
            # 산근 Q가 있으면 그 값(1순위), 없으면 같은 단산의 노무비로 역산(대가 8000N)
            used_dan.add(k)
            return f"산근 {k}호표 대가 {DAN_KEY_BASE + k}호표"
        if kind == "일위":
            k = ix_i.resolve(n, name, spec)
            return f"대가 {k}호표" if k is not None else ""
        k = ix_s.resolve(n, name, spec)
        if k is None:
            return ""
        used_std.add(k)
        return f"대가 {STD_KEY_BASE + k}호표"

    # 설계내역서: 공종 | 품명 | 규격 | 수량 | 단위 | 참조
    ws = out.create_sheet("설계내역서")
    for r in wb["설계내역서"].iter_rows(values_only=True):
        r = list(r) + [None] * 16
        code, name, spec, qty, unit = _s(r[1]), _s(r[2]), _s(r[3]), r[4], _s(r[5])
        if not code and not name:
            continue
        refs = " ".join(_s(v) for v in r[6:16] if isinstance(v, str))
        ws.append([code, name, spec, qty, unit, link(refs, name, spec)])

    # 단가산출근거: 블록 머리 'N.제목' → '제 N호표 제목' (안쪽 ' 1. 재료비'는 앞 공백이 있어 그대로)
    ws = out.create_sheet("단가산출근거")
    for r in wb["단가산출근거"].iter_rows(values_only=True):
        r = list(r) + [None] * 6
        t = r[1]
        if isinstance(t, str):
            m = re.match(r"^(\d+)\.(?=\S)(.*)$", t)
            if m and isinstance(r[2], (int, float)):
                t = f"제 {m.group(1)}호표 {m.group(2)}"
        ws.append([r[0], t] + list(r[2:6]))

    # 일위대가표: '제 N호표' / 제목 행 / 구성 행 / 합계 행 — 노무비 금액은 8번째 열(인덱스 7)
    ws = out.create_sheet("일위대가표")

    def blk(title, spec, unit, comps, labor_total, no):
        ws.append([f"제 {no}호표"])
        ws.append([title, spec, None, unit])
        for c in comps:
            ws.append(c)
        ws.append(["합  계", None, None, None, None, None, None, labor_total])

    cur = None
    comps: List[list] = []
    for r in wb["일위대가"].iter_rows(values_only=True):
        r = list(r) + [None] * 16
        m = re.match(r"^제\s*(\d+)\s*호표", _s(r[1]))
        if m:
            if cur:
                blk(cur[0], cur[1], cur[2], comps, cur[3], cur[4])
            cur = (_s(r[2]), _s(r[3]), _s(r[5]), _num(r[9]) or 0.0, int(m.group(1)))
            comps = []
            continue
        if cur is None or not _s(r[2]):
            continue
        nm, sp = _s(r[2]), _s(r[3])
        refs = " ".join(_s(v) for v in r[6:16] if isinstance(v, str))
        comps.append([nm, sp, r[4], _s(r[5]), None, None, None, _num(r[9]), None, None, None, None,
                      link(refs, nm, sp)])
    if cur:
        blk(cur[0], cur[1], cur[2], comps, cur[3], cur[4])
    # 표준시장단가·단가산출: 직종별 인수 없이 단위당 노무비만 쓴다 → 노무비 블록(품셈 작업조로 환산은 파서가 한다)
    # 블록 안에 노무비 행 하나를 둔다 — 다른 일위대가가 이 품명을 하위로 참조해 들어와도(제목으로 찾는다)
    # 노무비가 빠지지 않게(빈 블록이면 경주 '가물막이'가 톤마대 노무비를 잃어 0.8 → 69m/일로 부풀었다)
    def labor_blk(d, key):
        blk(d["name"], d["spec"], d["unit"],
            [[d["name"], d["spec"], 1, d["unit"], None, None, None, d["labor"]]], d["labor"], key)

    for n in sorted(used_std):
        d = std_rows.get(n)
        if d and d["labor"] > 0:
            labor_blk(d, STD_KEY_BASE + n)
    for n in sorted(used_dan):
        d = dan_rows.get(n)
        if d and d["labor"] > 0:
            labor_blk(d, DAN_KEY_BASE + n)

    # 노무비목록표: 직종 | | | 단가
    ws = out.create_sheet("노무비목록표")
    if "노임단가" in wb.sheetnames:
        for r in wb["노임단가"].iter_rows(values_only=True):
            r = list(r) + [None] * 6
            if isinstance(r[5], (int, float)) and _s(r[2]):
                ws.append([None, _s(r[2]), None, None, float(r[5])])
    return out
