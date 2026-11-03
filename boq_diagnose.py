# -*- coding: utf-8 -*-
"""
boq_diagnose.py — 내역서를 못 읽거나 일부만 읽었을 때 '왜'와 '엑셀에서 무엇을 바꾸면 읽히는지'

가장 좋은 건 변환기(boq_adapters)를 더해 사용자는 올리기만 하는 것이다. 그 전까지는 사용자가 엑셀을
조금 고쳐 지원 양식(표준형)에 맞출 수 있도록 원인과 고칠 방법을 구체적으로 알려 준다.

표준형(산근호표) — 앱이 가장 잘 읽는 배치
  - 시트: 설계내역서(항목) · 단가산출근거(산근 Q) · 일위대가표(직종별 인수) · 노무비목록표(노임)
  - 설계내역서: A 공종 번호 | B 품명 | C 규격 | D 수량 | E 단위 … 비고에 '산근 N호표' / '대가 N호표'
  - 공종 번호: 'Ⅰ, Ⅱ'(지구·구간) → '1.' 또는 '1.1'(대공종) → '1)'(세부) → 빈 칸(항목)
  - 단가산출근거: B열 블록 머리 '제 N호표', 아래 줄에 'Q = … = 24.33 ㎥/hr' 같은 산식
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional

ITEM_SHEETS = ("설계내역서", "내역서", "도급내역서")
UNIT_SHEETS = ("단가산출근거", "일위대가_산근")
BLOCK_SHEETS = ("일위대가표", "일위대가_호표", "일위대가")
WAGE_SHEETS = ("노무비목록표", "노임", "노임단가")

REF_PATTERNS = [
    ("산근 N호표", re.compile(r"산근\s*\d+\s*호표"), True),
    ("대가 N호표", re.compile(r"대가\s*\d+\s*호표"), True),
    ("단산N참조", re.compile(r"단산\s*\d+\s*참조"), True),
    ("일위N참조", re.compile(r"일위\s*\d+\s*참조"), True),
    ("표준N참조", re.compile(r"표준\s*\d+\s*참조"), True),
    ("호표 N / No.N", re.compile(r"(?:^|\s)(?:호표\s*\d+|No\.\s*\d+)"), False),
    ("단가 N / 일위 N(번호만)", re.compile(r"^(?:단가|일위|산근)\s*\d+$"), False),
]
_HEADER_WORDS = {"name": ("품명", "명칭", "공종명", "품목"), "spec": ("규격",), "qty": ("수량",),
                 "unit": ("단위",)}


def _s(v) -> str:
    return "" if v is None else str(v).strip()


def _col(i: int) -> str:
    return chr(ord("A") + i) if i < 26 else f"{i + 1}번째"


def _header(ws):
    """(행 번호, {키: 열}) — 첫 30행에서 품명·수량·단위가 있는 머리행."""
    for n, row in enumerate(ws.iter_rows(min_row=1, max_row=30, values_only=True), 1):
        cols: Dict[str, int] = {}
        for i, v in enumerate(row):
            t = re.sub(r"\s+", "", _s(v))
            for k, words in _HEADER_WORDS.items():
                if k not in cols and t in words:
                    cols[k] = i
        if {"name", "qty", "unit"} <= cols.keys():
            return n, cols
    return None, {}


def diagnose(wb, tmpl: Optional[Dict[str, Any]] = None) -> List[Dict[str, str]]:
    """[{level: 'ok'|'warn'|'error', title, detail, fix}] — 원본 워크북 기준.

    tmpl(인식된 양식)을 주면 그 양식의 규칙으로 본다 — 열 위치를 스스로 찾는 양식(변환기·머리행 인식)은 열 위치를,
    코드 열로 잇는 양식(코드매칭형)은 비고 표기·공종 열 번호를 따지지 않는다. 없으면 표준형 기준.
    """
    out: List[Dict[str, str]] = []
    names = list(wb.sheetnames)
    tmpl = tmpl or {}
    auto_cols = bool(tmpl.get("adapter") or tmpl.get("detect_header"))
    code_link = tmpl.get("link_strategy") == "column_code"
    tree_in_gj = tmpl.get("hierarchy_strategy", "district_roman") in ("district_roman", "letter_tree")

    def add(level, title, detail, fix=""):
        out.append({"level": level, "title": title, "detail": detail, "fix": fix})

    item = next((s for s in ITEM_SHEETS if s in names), None)
    if item is None:
        cand = [s for s in names if "내역" in s]
        add("error", "공사 항목 시트를 찾지 못함",
            f"시트: {', '.join(names[:12])}",
            ("공사 항목(품명·규격·수량·단위)이 있는 시트 이름을 '설계내역서'로 바꾸세요."
             + (f" 후보: '{cand[0]}'" if cand else "")))
        return out
    add("ok", f"공사 항목 시트: '{item}'", "")
    ws = wb[item]

    hr, cols = _header(ws)
    if not hr:
        add("error", "머리행(품명·수량·단위)을 찾지 못함",
            f"'{item}' 시트 첫 30행에 '품명'·'수량'·'단위' 글자가 한 행에 모두 있지 않습니다.",
            "항목 위에 머리행을 두고 칸 제목을 '품명(또는 명칭)', '규격', '수량', '단위'로 적으세요.")
    else:
        pos = " · ".join(f"{k} {_col(v)}열" for k, v in (("품명", cols["name"]), ("규격", cols.get("spec", -1)),
                                                        ("수량", cols["qty"]), ("단위", cols["unit"])) if v >= 0)
        std = auto_cols or code_link or (cols.get("name") == 1 and cols.get("qty") == 3 and cols.get("unit") == 4)
        add("ok" if std else "warn", f"머리행 {hr}행: {pos}",
            "" if std else "표준형 배치(A 공종 · B 품명 · C 규격 · D 수량 · E 단위)와 열 위치가 다릅니다.",
            "" if std else "앞쪽 빈 열을 지우거나 열을 옮겨 A 공종 번호 · B 품명 · C 규격 · D 수량 · E 단위로 맞추세요.")

    # 공종 번호(계층)
    gj_col = max(0, cols.get("name", 1) - 1)
    codes: Dict[str, int] = {}
    refs: Dict[str, int] = {}
    n_items = 0
    for n, row in enumerate(ws.iter_rows(values_only=True)):
        if n > 3000:
            break
        c = _s(row[gj_col]) if len(row) > gj_col else ""
        if c:
            kind = ("로마숫자(Ⅰ)" if re.match(r"^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+\.?$", c) else
                    "영문(A.)" if re.match(r"^[A-Z]\.?$", c) else
                    "숫자(1.)" if re.match(r"^\d+\.?$", c) else
                    "숫자(1.1)" if re.match(r"^\d+(\.\d+)+\.?$", c) else
                    "괄호(1))" if re.match(r"^\(?\d+\)$", c) else
                    "기호(◎·■)" if c in ("◎", "■", "●", "□") else "기타")
            codes[kind] = codes.get(kind, 0) + 1
        q = row[cols["qty"]] if cols and len(row) > cols["qty"] else None
        if isinstance(q, (int, float)) and q > 0 and (not c or not tree_in_gj):
            n_items += 1
        for v in row:
            if isinstance(v, str):
                for lbl, rx, _ok in REF_PATTERNS:
                    if rx.search(v):
                        refs[lbl] = refs.get(lbl, 0) + 1
                        break
    code_txt = ", ".join(f"{k} {v}개" for k, v in sorted(codes.items(), key=lambda x: -x[1])) or "없음"
    if not tree_in_gj:
        add("ok", "공종 계층: 품명 열의 번호로 읽는 양식", f"수량이 있는 항목 행 약 {n_items}개")
    elif not codes:
        add("error", "공종 번호가 없음", f"공종 열({_col(gj_col)}열)에 번호가 없어 대공종·구간을 나눌 수 없습니다.",
            f"{_col(gj_col)}열에 대공종 번호('1.' 또는 '1.1')를, 구간이 있으면 위에 'Ⅰ, Ⅱ'를 적으세요.")
    elif "기타" in codes and len(codes) == 1:
        add("warn", "공종 번호 형식을 알 수 없음", f"공종 열 번호: {code_txt}",
            "대공종 줄에 '1.'·'1.1', 구간 줄에 'Ⅰ'·'A.'처럼 번호를 적으세요.")
    else:
        add("ok", f"공종 번호: {code_txt}", f"수량이 있는 항목 행 약 {n_items}개")

    # 참조 표기(1일 작업량 근거 연결)
    supported = {"산근 N호표", "대가 N호표", "단산N참조", "일위N참조", "표준N참조"}
    ref_txt = ", ".join(f"'{k}' {v}개" for k, v in sorted(refs.items(), key=lambda x: -x[1])) or "없음"
    if code_link:
        cc = tmpl.get("link_code_col")
        n_code = sum(1 for n, row in enumerate(ws.iter_rows(values_only=True))
                     if n < 5000 and cc is not None and len(row) > cc and _s(row[cc]))
        if n_code:
            add("ok", f"근거 참조: 코드 열({_col(cc)}열) {n_code}개", "")
        else:
            add("warn", "근거 코드 열이 비어 있음", f"{_col(cc)}열에 일위대가·산근 코드가 없습니다.",
                "설계 프로그램에서 '코드' 열을 포함해 내려받으세요.")
    elif not refs:
        add("warn", "근거 참조 표기가 없음",
            "항목이 단가산출근거·일위대가 어느 블록을 쓰는지 적힌 칸을 찾지 못했습니다 — 1일 작업량을 품명으로만 "
            "찾게 되어 '매칭 안 됨'이 많아집니다.",
            "비고 칸에 단가산출근거는 '산근 N호표', 일위대가는 '대가 N호표'로 적으세요.")
    elif not (set(refs) & supported):
        add("error", "지원하지 않는 참조 표기", f"찾은 표기: {ref_txt}",
            "비고 칸 표기를 '산근 N호표'(단가산출근거) / '대가 N호표'(일위대가)로 바꾸면 읽습니다. "
            "또는 이 파일을 보내 주시면 변환기를 추가합니다.")
    else:
        add("ok", f"근거 참조: {ref_txt}", "")

    # 근거 시트
    us = next((s for s in UNIT_SHEETS if s in names), None)
    bs = next((s for s in BLOCK_SHEETS if s in names), None)
    wsn = next((s for s in WAGE_SHEETS if s in names), None)
    if not us and not bs:
        add("warn", "단가산출근거·일위대가 시트 없음",
            "Q값·직종별 인수를 쓸 수 없어 1일 작업량은 표준품셈·가이드라인으로만 찾습니다.",
            "설계 내역서 원본에서 '단가산출근거'·'일위대가표' 시트까지 함께 내려받아 올리세요.")
    else:
        add("ok", "근거 시트: " + ", ".join(s for s in (us, bs, wsn) if s), "")
        if us:
            u = wb[us]
            heads = 0
            for n, row in enumerate(u.iter_rows(values_only=True)):
                if n > 20000:
                    break
                for v in row[:3]:
                    if isinstance(v, str) and re.search(r"제\s*\d+\s*호표|^\d+\.\S", v.strip()):
                        heads += 1
                        break
            if heads == 0:
                add("warn", f"'{us}' 블록 머리를 찾지 못함",
                    "블록마다 '제 N호표' 같은 머리줄이 있어야 번호로 Q값을 찾습니다.",
                    "각 산출 블록 첫 줄을 '제 N호표 품명'으로 적으세요.")
    if not wsn and (us or bs):
        add("warn", "노임 시트 없음", "직종별 노임을 몰라 노무비 역산·작업조 환산이 기본 노임으로만 됩니다.",
            "노임 단가 시트를 '노무비목록표'로 두고 B열 직종 · E열 단가로 적으세요.")
    return out


def quality(items: List[Dict[str, Any]], code_daily: Dict, labor_daily: Dict, aq_daily: Dict,
            non_work=lambda c: False) -> Dict[str, Any]:
    """읽은 항목의 연결 상태 — 공사비 외·자재 대공종은 뺀다."""
    work = [i for i in items if not non_work(i.get("category") or "")]
    linked = [i for i in work if i.get("code") is not None or i.get("code_alt") is not None]
    q_hit = [i for i in work if i.get("code") in code_daily]
    # 노무비·A·Q 키: 일위대가 참조(code_alt), 코드매칭형처럼 하나뿐이면 그 코드
    lk = lambda i: i.get("code_alt") if i.get("code_alt") is not None else i.get("code")   # noqa: E731
    l_hit = [i for i in work if lk(i) is not None and (lk(i) in labor_daily or lk(i) in aq_daily)]
    n = len(work) or 1
    found = len({id(i) for i in q_hit + l_hit})          # Q값과 인수가 둘 다 있는 항목은 한 번만
    return {"n_items": len(items), "n_work": len(work), "linked": len(linked), "q": len(q_hit),
            "labor": len(l_hit), "found": found, "linked_pct": round(100 * len(linked) / n),
            "rate_pct": round(100 * found / n)}
