# -*- coding: utf-8 -*-
"""
universal_parser.py — 설정(config) 기반 범용 내역서/단가근거 파서

배경
----
회사·적산 소프트웨어마다 설계내역서 엑셀 양식이 다르다 (시트명, 열 배치,
계층 표시 방식, 단가근거 연결 방식 등). 하지만 시공량 산식(Q=/HR)은 전부
같은 국가 표준품셈 방법론을 쓰기 때문에 **완전히 동일**하다.

그래서 이 모듈은 "매번 다른 부분"만 TEMPLATES 설정으로 분리하고,
"항상 같은 부분"(Q값 추출 정규식)은 하나의 엔진으로 공유한다.
새 회사 양식을 만나면 파이썬 코드를 새로 짜지 않고, TEMPLATES에
설정 딕셔너리 하나를 추가하는 것으로 대응하는 게 목표다.

검증 상태
---------
아래 2개 템플릿은 실제 파일로 검증 완료. 결과가 기존 전용 파서
(hopyo_parser.py, naeyeok_parser.py)와 정확히 일치함을 확인함:
  - "표준형(산근호표)": 3_토목도급-동부지구... 파일
  - "코드매칭형(내역서산근)": (토목공사)화성 관로R18... 파일

새 템플릿 추가 방법
-------------------
1. 새 엑셀의 시트명/열 배치를 실측 확인 (openpyxl로 몇 행 찍어보기)
2. 아래 TEMPLATES 리스트에 dict 하나 추가:
   - item_sheet_names: 항목 시트 이름 후보 리스트
   - unit_price_sheet_names: 단가근거 시트 이름 후보 리스트
   - name_col, spec_col, qty_col, unit_col: 항목시트 열 인덱스
   - leaf_strategy: "column_tag"(특정 열에 값 있으면 리프) 또는
                    "regex_hierarchy"(gong_jong 열이 특정 패턴 아니면 리프)
   - hierarchy_strategy: "major_number_prefix"(들여쓰기0 "N.이름") 또는
                          "district_roman"(로마숫자+"N.N.N"코드)
   - link_strategy: "column_code"(특정 열 값이 코드) 또는
                    "text_regex"(정규식으로 텍스트에서 추출)
   - unit_price_code_col, unit_price_formula_col: 단가근거 시트 열 인덱스
3. python universal_parser.py <새엑셀경로> 로 검증 (자동판별 실패시 강제 지정 가능)
"""

from __future__ import annotations

import re
import sys
from typing import Dict, List, Optional, Tuple, Any


# ---------------------------------------------------------------------------
# 공용 Q값 추출 엔진 (모든 템플릿이 동일하게 재사용 — 국가 표준품셈 방법론이 같으므로)
# ---------------------------------------------------------------------------

# 번호가 붙은 Q 변수(Q1, Q2 …). 'Q = …'(주작업)와 함께 나오면 순차 공정이다.
# 예: 제228호표 '열연강판 항타 및 항발' — 1.항타 Q=3.79, 2.항발 Q1=7.80.
# 한 본에 두 작업을 모두 하므로 소요시간을 더해야 하고,
# 작업량으로는 역수를 합산한다: 1/Q_합 = 1/Q + 1/Q1.
# (실측: 항타만 쓰면 30.3본/일이지만 합산하면 20.4본/일 — 1.49배 과대였음)
def _build_numbered_q_re(unit_char_class: str):
    return re.compile(
        r"\bQ(\d)\s*=.*?([\d,.]+)\s*([" + unit_char_class + r"]+)\s*/\s*hr",
        re.IGNORECASE,
    )


def _build_q_regexes(unit_char_class: str):
    """템플릿별 단위 문자클래스로 정방향/역수 Q정규식을 생성."""
    fwd = re.compile(r"=\s*([\d,.]+)\s*([" + unit_char_class + r"]+)/hr", re.IGNORECASE)
    inv = re.compile(r"=\s*([\d,.]+)\s*hr/([" + unit_char_class + r"]+)", re.IGNORECASE)
    return fwd, inv


# 기본값(하위호환용) — 실제로는 항상 템플릿의 unit_char_class를 통해 생성해서 씀
_Q_FORWARD_RE, _Q_INVERSE_RE = _build_q_regexes(r"가-힣㎡㎥")
_HR_TO_DAY = 8


def _to_float(s: str) -> float:
    return float(s.replace(",", ""))


# 일 기준(Q = 600 ㎡/일) 산식.
# 단가산출근거에는 시간당(/hr)과 일당(/일) 표기가 섞여 있는데,
#   ① 'Q = 300 ㎡/일 /8 Hr = 37.50'  → 최종값이 시간당이므로 ×8 (기존 /hr 로직이 처리)
#   ② 'Q = 600 ㎡/일 * 0.18 = 108.00' → 최종값이 일당이므로 그대로 사용
#   ③ 'Q = 600 ㎡/일'                → 일당 그대로 사용
# ②③은 /hr 표기가 없어 기존 파서가 통째로 놓치고 있었다(실측 9개 호표).
# '/일' 뒤에 시간 환산(/8 Hr 등)이 없는 경우에만 일당으로 인식한다.
# 'Q = 8000 ㎡/일 /8 Hr = 1000.00' 형태.
# 일당 시공량을 8시간으로 나눠 시간당을 구한 산식으로, 최종값 뒤에 단위가 없어
# 기존 /hr 정규식이 놓치고 있었다(실측 13건, 제222·317호표 등).
# 이 경우 '/일' 앞의 값이 곧 일작업량이므로 그대로 쓴다.
_Q_DAY_DIV8_RE = re.compile(
    r"Q\s*\d*\s*=\s*([\d,.]+)\s*([가-힣㎡㎥A-Za-z0-9]+)\s*/\s*일\s*/\s*8\s*Hr",
    re.IGNORECASE,
)

_Q_PER_DAY_RE = re.compile(
    r"Q\s*\d*\s*=\s*([\d,.]+)\s*([가-힣㎡㎥A-Za-z0-9]+)\s*/\s*일"
    r"(?!\s*/)"                      # 바로 뒤에 '/8 Hr' 같은 환산이 오면 제외
    r"(?:\s*\*\s*([\d.]+))?"        # '* 0.18' 같은 계수(선택)
    r"\s*(?:=\s*([\d,.]+))?\s*$",   # '= 108.00' 최종값(선택)
    re.IGNORECASE,
)


# 조립식 간이흙막이 등 '공용일수' 방식 블록.
#   L = 30 m
#   ① 굴착소요시간   D1 = 0.17 일
#   ② 흙막이 박기및뽑기 D2 = 1.38 일
#   ③ 관 접합및부설   D3 = 0.5 일
#   ④ 되메우기 및 기타 D4 = 0.5 일
#   공용일수 D = D1+D2+D3+D4 = 2.55 일
# 여기서 공용일수는 '가설재가 현장에 묶여 있는 기간'(임대료 산출용)이라
# 그대로 쓰면 관부설·되메우기가 별도 항목과 중복 계상된다.
# 흙막이 항목의 작업량은 흙막이 자체 작업분(D2)만으로 잡는다: L ÷ D2 (m/일).
_SHORING_L_RE = re.compile(r"L\s*=\s*([\d.]+)\s*m", re.IGNORECASE)
_SHORING_D2_RE = re.compile(r"②[^:]*:\s*D2\s*=.*?=\s*([\d.]+)\s*일")


def _extract_shoring_rate(rows, start: int, end: int, formula_col: int):
    """공용일수 블록에서 흙막이 작업량(L/D2, m/일)을 구한다. 없으면 None."""
    has_shared = False
    L = None
    D2 = None
    for j in range(start, end):
        cell = rows[j][formula_col] if len(rows[j]) > formula_col else None
        if not isinstance(cell, str):
            continue
        if "공용일수" in cell:
            has_shared = True
        if L is None:
            mL = _SHORING_L_RE.search(cell)
            if mL:
                try:
                    L = float(mL.group(1))
                except ValueError:
                    pass
        mD = _SHORING_D2_RE.search(cell)
        if mD:
            try:
                D2 = float(mD.group(1))
            except ValueError:
                pass
    if has_shared and L and D2 and D2 > 0:
        return round(L / D2, 4), "m"
    return None


def _extract_q_from_block(rows, start: int, end: int, formula_col: int, q_forward_re=None, q_inverse_re=None, unit_char_class=None):
    """[start, end) 범위에서 Q 산식을 찾아 (일작업량, 단위) 반환.

    우선순위: 시간당(/hr) → 일당(/일) → 역수(hr/단위)
    """
    # 공용일수 방식(간이흙막이 등)은 블록 안의 굴착 Q(㎥/hr)를 잡으면
    # 항목 단위(M)와 어긋나므로 먼저 확인한다.
    _shoring = _extract_shoring_rate(rows, start, end, formula_col)
    if _shoring:
        return _shoring

    q_forward_re = q_forward_re or _Q_FORWARD_RE
    q_inverse_re = q_inverse_re or _Q_INVERSE_RE
    numbered_re = _build_numbered_q_re(unit_char_class or r"가-힣㎡㎥")
    found_fwd = None
    found_inv = None
    found_day = None
    numbered = {}
    for j in range(start, end):
        row = rows[j]
        cell = row[formula_col] if len(row) > formula_col else None
        if not isinstance(cell, str):
            continue
        mn = numbered_re.search(cell)
        if mn:
            numbered.setdefault(mn.group(1), (_to_float(mn.group(2)), mn.group(3)))
        if found_fwd is None and not mn:
            m = q_forward_re.search(cell)
            if m:
                found_fwd = (_to_float(m.group(1)), m.group(2))
        if found_day is None:
            md8 = _Q_DAY_DIV8_RE.search(cell.strip())
            if md8:
                v8 = _to_float(md8.group(1))
                if v8 > 0:
                    found_day = (v8, md8.group(2))
        if found_day is None:
            md = _Q_PER_DAY_RE.search(cell.strip())
            if md:
                base = _to_float(md.group(1))
                unit_d = md.group(2)
                # 최종값이 명시돼 있으면 그것을, 아니면 기준값×계수를 쓴다
                if md.group(4):
                    val_d = _to_float(md.group(4))
                elif md.group(3):
                    val_d = base * float(md.group(3))
                else:
                    val_d = base
                if val_d > 0:
                    found_day = (val_d, unit_d)
        if found_fwd is None and found_inv is None:
            mi = q_inverse_re.search(cell)
            if mi:
                found_inv = (_to_float(mi.group(1)), mi.group(2))
        # 번호 Q까지 모아야 하므로 조기 종료하지 않는다
    if found_fwd is not None:
        val, unit = found_fwd
        # 'Q'와 'Q1…'이 함께 있으면 순차 공정 → 소요시간 합산(작업량은 역수 합)
        if numbered and val > 0:
            inv_sum = 1.0 / val + sum(1.0 / v for v, _u in numbered.values() if v > 0)
            if inv_sum > 0:
                val = 1.0 / inv_sum
        return round(val * _HR_TO_DAY, 4), unit
    if numbered:
        # 'Q1'만 있는 경우(항발 전용 호표 등)는 그 값이 곧 작업량이다.
        # 여러 개면 순차로 보고 역수 합산한다.
        vals = [(v, u) for v, u in numbered.values() if v > 0]
        if vals:
            inv_sum = sum(1.0 / v for v, _u in vals)
            if inv_sum > 0:
                return round((1.0 / inv_sum) * _HR_TO_DAY, 4), vals[0][1]
    if found_day is not None:
        val, unit = found_day
        return round(val, 4), unit          # 일당이므로 ×8 하지 않음
    if found_inv is not None:
        val, unit = found_inv
        if val > 0:
            return round(_HR_TO_DAY / val, 4), unit
    return None


# ---------------------------------------------------------------------------
# 템플릿 레지스트리
# ---------------------------------------------------------------------------

TEMPLATES: List[Dict[str, Any]] = [
    {
        "id": "standard_hopyo",
        "name": "표준형(산근호표)",
        "item_sheet_names": ["설계내역서"],
        "unit_price_sheet_names": ["단가산출근거"],
        "name_col": 1,
        "spec_col": 2,
        "qty_col": 3,
        "unit_col": 4,
        "leaf_strategy": "regex_hierarchy",
        "hierarchy_strategy": "district_roman",
        "gong_jong_col": 0,
        "hierarchy_leaf_regex": None,  # 로마숫자/3단계코드가 아니고 name이 있으면 리프 (district_roman 내부 처리)
        "link_strategy": "text_regex",
        "link_regex": r"산근\s*(\d+)\s*호표",
        # 설계내역서에는 '산근 N호표'(단가산출근거)와 '대가 N호표'(일위대가표) 두 종류
        # 참조가 함께 쓰인다. 후자는 노무비 역산 키로 연결된다.
        "link_regex_alt": r"대가\s*(\d+)\s*호표",
        "link_key_type": "int",
        "unit_price_block_start_col": 1,
        "unit_price_block_start_regex": r"제\s*(\d+)\s*호표",
        "unit_price_formula_col": 1,
        "unit_char_class": r"가-힣㎡㎥",
        # 노무비 역산 설정 (Q산식 없는 호표용)
        "labor": {
            "wage_sheet": "노무비목록표", "wage_name_col": 1, "wage_rate_col": 4,
            "block_sheet": "일위대가표",
            "block_start_col": 0, "block_start_regex": r"제\s*(\d+)\s*호표",
            "title_offset": 1, "title_col": 0, "unit_col": 3,
            "sum_row_keyword": "합", "labor_amount_col": 7,
            "base_qty": 1.0, "key_type": "int",
        },
    },
    {
        "id": "code_match_naeyeok",
        "name": "코드매칭형(내역서산근)",
        "item_sheet_names": ["내역서"],
        "unit_price_sheet_names": ["일위대가_산근"],
        "name_col": 0,
        "spec_col": 1,
        "qty_col": 2,
        "unit_col": 3,
        "leaf_strategy": "column_tag",
        "leaf_tag_col": 12,
        "hierarchy_strategy": "major_number_prefix",
        "line_marker_symbol": "■",
        "line_spec_col": 1,
        "link_strategy": "column_code",
        "link_code_col": 17,
        "link_key_type": "str",
        "unit_price_block_start_col": 7,
        "unit_price_block_start_regex": None,  # 코드매칭형은 정규식 없이 값 존재 자체가 블록시작
        "unit_price_formula_col": 0,
        "unit_char_class": r"가-힣㎡㎥A-Za-z0-9",
        # 노무비 역산 설정 (화성 양식: 코드가 열14, 블록 헤더행에 수량/노무비가 함께 있음)
        "labor": {
            "wage_sheet": "노임", "wage_name_col": 0, "wage_rate_col": 3,
            "block_sheet": "일위대가_호표",
            "block_code_col": 14, "title_col": 0, "qty_col": 2, "unit_col": 3,
            "labor_amount_col": 9,
            "key_type": "str",
        },
    },
]


def detect_template(wb) -> Optional[Dict[str, Any]]:
    """워크북 시트명으로 등록된 템플릿 중 매칭되는 것을 찾는다."""
    if wb is None:
        return None
    names = set(wb.sheetnames)
    for tmpl in TEMPLATES:
        if any(s in names for s in tmpl["item_sheet_names"]) and \
           any(s in names for s in tmpl["unit_price_sheet_names"]):
            return tmpl
    return None


def _get_sheet(wb, name_candidates):
    for n in name_candidates:
        if n in wb.sheetnames:
            return wb[n]
    return None


# ---------------------------------------------------------------------------
# 단가근거 시트 파싱 (공용 — 블록 탐지 방식만 템플릿마다 다름)
# ---------------------------------------------------------------------------

def parse_unit_price_generic(wb, tmpl: Dict[str, Any]) -> Dict[Any, Tuple[float, str]]:
    """
    단가근거 시트에서 {키(호표번호 또는 코드): (일작업량, 단위)} 추출.
    블록 시작 판별은 템플릿의 unit_price_block_start_col/regex를 따르고,
    Q값 추출 자체는 모든 템플릿이 공유하는 _extract_q_from_block을 쓴다.
    """
    ws = _get_sheet(wb, tmpl["unit_price_sheet_names"])
    if ws is None:
        return {}

    rows = [r for r in ws.iter_rows(values_only=True)]
    block_col = tmpl["unit_price_block_start_col"]
    block_regex = tmpl.get("unit_price_block_start_regex")
    formula_col = tmpl["unit_price_formula_col"]
    key_type = tmpl.get("link_key_type", "str")
    unit_char_class = tmpl.get("unit_char_class", r"가-힣㎡㎥")
    q_fwd_re, q_inv_re = _build_q_regexes(unit_char_class)

    # 1) 블록 시작행 인덱싱
    starts: List[Tuple[Any, int]] = []
    seen_keys = set()
    if block_regex:
        pat = re.compile(block_regex)
        for i, r in enumerate(rows):
            c = r[block_col] if len(r) > block_col else None
            if isinstance(c, str):
                m = pat.search(c)
                if m:
                    key = int(m.group(1)) if key_type == "int" else m.group(1)
                    if key not in seen_keys:
                        seen_keys.add(key)
                        starts.append((key, i))
    else:
        for i, r in enumerate(rows):
            c = r[block_col] if len(r) > block_col else None
            if isinstance(c, str) and c.strip():
                key = c.strip()
                if key not in seen_keys:
                    seen_keys.add(key)
                    starts.append((key, i))

    if not starts:
        return {}

    # 2) 블록별 Q값 추출
    result: Dict[Any, Tuple[float, str]] = {}
    for idx, (key, s) in enumerate(starts):
        end = starts[idx + 1][1] if idx + 1 < len(starts) else len(rows)
        found = _extract_q_from_block(rows, s, end, formula_col, q_fwd_re, q_inv_re, unit_char_class)
        if found:
            result[key] = found
    return result


# ---------------------------------------------------------------------------
# 항목 시트 파싱 (계층 전략별 분기)
# ---------------------------------------------------------------------------

def _parse_items_major_number_prefix(ws, tmpl: Dict[str, Any]) -> List[Dict]:
    """'N. 이름' 대공종 + 라인/구분 추적. leaf_strategy='column_tag' 전제.

    두 가지 코드매칭형 변형을 모두 지원한다:
      - 화성형: 들여쓰기 0의 '■ 관로'(라인) + 'N. 토공'(대공종)
      - 명륜형: '[주간공사]'/'[야간공사]'(구분) + ' 1. 토     공'(들여쓰기·내부공백 있음)
    대공종 판정에서 들여쓰기를 요구하지 않으며, 이름 중간 공백은 하나로 정규화한다.
    """
    major_re = re.compile(r"^\d+\.\s*(.+)$")
    bracket_re = re.compile(r"^\[(.+)\]$")
    name_col = tmpl["name_col"]
    spec_col = tmpl["spec_col"]
    qty_col = tmpl["qty_col"]
    unit_col = tmpl["unit_col"]
    leaf_tag_col = tmpl["leaf_tag_col"]
    line_symbol = tmpl["line_marker_symbol"]
    line_spec_col = tmpl["line_spec_col"]
    link_code_col = tmpl.get("link_code_col")

    current_major = None
    current_line = ""
    items: List[Dict] = []

    for row in ws.iter_rows(values_only=True):
        raw_name = row[name_col] if len(row) > name_col else None
        if not raw_name:
            continue
        s = str(raw_name)
        stripped = s.lstrip(" ")
        indent = len(s) - len(stripped)

        tag = row[leaf_tag_col] if len(row) > leaf_tag_col else None
        qty = row[qty_col] if len(row) > qty_col else None
        is_leaf = tag not in (None, "") and isinstance(qty, (int, float))

        if not is_leaf:
            # 라인 표시(■ 등)는 들여쓰기 0에서만 인식 (본문 항목명과 혼동 방지)
            if indent == 0 and line_symbol and stripped.startswith(line_symbol):
                spec = str(row[line_spec_col]).strip() if len(row) > line_spec_col and row[line_spec_col] else ""
                current_line = spec
                continue
            # '[주간공사]' 같은 대괄호 구분을 라인으로 사용
            mb = bracket_re.match(stripped)
            if mb:
                current_line = mb.group(1).strip()
                continue
            # 'N. 이름' 대공종 (들여쓰기 무관, 내부 공백 정규화).
            # 단, 진짜 대공종은 소계행이라 단위가 '식'이다. '5.2) 조립식PC맨홀'처럼
            # 단위가 빈 하위 분류를 대공종으로 오인하면 커버리지가 크게 떨어진다(화성 실측).
            _u = row[unit_col] if len(row) > unit_col else None
            _u = str(_u).strip() if _u else ""
            m = major_re.match(stripped)
            if m and _u == "식":
                current_major = re.sub(r"\s+", " ", m.group(1)).strip()
            continue

        if is_leaf:
            spec = str(row[spec_col]).strip() if len(row) > spec_col and row[spec_col] else ""
            unit = str(row[unit_col]).strip() if len(row) > unit_col and row[unit_col] else ""
            code = row[link_code_col] if link_code_col is not None and len(row) > link_code_col else None
            code = code.strip() if isinstance(code, str) else None
            items.append({
                "name": stripped, "spec": spec, "qty": qty, "unit": unit,
                "code": code, "category": current_major, "line": current_line,
            })
    return items


def _parse_items_district_roman(ws, tmpl: Dict[str, Any]) -> List[Dict]:
    """
    로마숫자 지구 + 'N.N.N' 3단계 코드(대분류) + 'N)' 소분류 + '(N) #N...' 세부구분자.

    app.py 레거시 계층파서를 그대로 이식한 상태머신. 단순 flat 추출이 아니라
    아래 병합 규칙까지 정확히 재현해야 한다 (실측으로 차이 확인됨):
      - sub_category(예: '1) 토공')는 같은 level+name+district 조합이면
        재사용(reuse)된다 — 시트 내 여러 곳에 흩어져 나와도 하나로 누적.
      - 항목은 현재 활성 컨테이너(sub_sub > sub > category) 안에서
        (name, spec)이 같으면 수량을 합산한다.
      - 예외: '#N' 형태 세부구분자(sub_sub_category)이고 그 부모 sub_category
        이름에 '추진'이 포함되면, sub_sub 레벨을 건너뛰고 부모 sub_category
        레벨로 합산한다 (여러 #N 추진 구간의 같은 항목을 하나로 합치기 위함).
    내부적으로 중첩 hierarchy를 만들어 정확히 합산한 뒤, 최종적으로
    {name,spec,qty,unit,code,category,line} 평평한 리스트로 변환해 반환한다.
    """
    roman_nums = ['Ⅰ', 'Ⅱ', 'Ⅲ', 'Ⅳ', 'Ⅴ', 'Ⅵ', 'Ⅶ', 'Ⅷ', 'Ⅸ', 'Ⅹ']
    # 대공종 레벨 자동 판별 (1/2/3단계 중 하나):
    #   관로 내역서(3_토목도급): '1.1.1 토공' → 3단계가 대공종
    #   처리장 내역서(본대/의성): '1.1 가시설공' → 2단계가 대공종
    #   종합발주 내역서(풍각):    '1. 토공'    → 1단계가 대공종
    # 각 레벨에서 '공종형' 이름(…공/…공사/…자재대 등)이 붙은 헤더 수를 세어,
    # 가장 많은 레벨을 대공종으로 삼는다. 동수면 얕은(상위) 레벨을 우선한다
    # (상위가 진짜 대공종이고 하위는 그 세부인 경우가 일반적).
    _lv1 = re.compile(r"^\d+$")
    _lv2 = re.compile(r"^\d+\.\d+$")
    _lv3 = re.compile(r"^\d+\.\d+\.\d+$")
    _gj_col = tmpl["gong_jong_col"]
    _nm_col = tmpl["name_col"]
    _JOB_SUFFIX = ("공", "공사", "자재비", "자재대", "운반공", "부대공")
    _score = {1: 0, 2: 0, 3: 0}
    for _row in ws.iter_rows(values_only=True):
        _g = _row[_gj_col] if len(_row) > _gj_col else None
        _n = _row[_nm_col] if len(_row) > _nm_col else None
        if not (_g and _n):
            continue
        _gs = str(_g).strip()
        _ns = str(_n).strip()
        if not _ns or not _ns.endswith(_JOB_SUFFIX):
            continue
        if _lv1.match(_gs):
            _score[1] += 1
        elif _lv2.match(_gs):
            _score[2] += 1
        elif _lv3.match(_gs):
            _score[3] += 1
    # 레벨 선택 규칙 (실측 6개 파일로 검증):
    #   상위 레벨이 대공종이려면 '공종형 헤더가 충분히 많아야' 한다.
    #   - 풍각: 레벨1에 11개(토공/구조물공/…) → 레벨1이 대공종
    #   - 정수장: 레벨1은 2개뿐이고 그마저 '관급자재비'라 실제 대공종이 아님 →
    #             레벨2(가시설공/토공/구조물공사 8개)가 대공종
    #   - 본대: 레벨2 8개 vs 레벨3 18개지만 레벨3은 '2.1.1 토공'처럼 세부 반복 →
    #           하위가 더 많아도 상위(레벨2)를 우선
    #   따라서 '해당 레벨 점수가 4 이상'인 가장 얕은 레벨을 채택하고,
    #   아무 레벨도 그 기준을 못 넘으면 점수가 가장 높은 레벨을 쓴다.
    _MIN_MAJOR = 4
    _best = None
    for lv in (1, 2, 3):
        if _score[lv] >= _MIN_MAJOR:
            _best = lv
            break
    if _best is None:
        _best = max((1, 2, 3), key=lambda lv: (_score[lv], -lv)) if any(_score.values()) else 3
    major_code_re = {1: _lv1, 2: _lv2, 3: _lv3}[_best]
    # 대공종보다 한 단계 위(사업/구역 구분) 정규식.
    # 예: 대공종이 '1.2 토공'(2단계)이면 상위는 '1 신설오수관로'(1단계).
    # 처리장+관로가 한 내역서에 있으면 같은 '토공'이라도 조건·조수가 달라
    # 분리해서 봐야 하므로, 이 상위 구분을 항목에 함께 기록한다.
    parent_code_re = {1: None, 2: _lv1, 3: _lv2}[_best]
    sub_re = re.compile(r"^\d+\)$")
    hash_paren_re = re.compile(r"^\(\d+\)$")
    hash_name_re = re.compile(r"^#\d+")
    hash_gj_re = re.compile(r"^#\d+")

    gong_jong_col = tmpl["gong_jong_col"]
    name_col = tmpl["name_col"]
    spec_col = tmpl["spec_col"]
    qty_col = tmpl["qty_col"]
    unit_col = tmpl["unit_col"]
    link_regex = re.compile(tmpl["link_regex"]) if tmpl.get("link_regex") else None
    link_regex_alt = re.compile(tmpl["link_regex_alt"]) if tmpl.get("link_regex_alt") else None
    key_type = tmpl.get("link_key_type", "str")

    hierarchy: List[Dict] = []
    current_district = None
    current_parent = None      # 상위 사업 구분명 (예: 신설오수관로 / 배수설비)
    current_category = None
    current_sub_category = None
    current_sub_sub_category = None

    def _merge_or_append(container_items: List[Dict], item: Dict):
        existing = next((i for i in container_items
                          if i['name'] == item['name'] and i.get('spec') == item.get('spec')), None)
        if existing:
            existing['qty'] = existing.get('qty', 0) + item.get('qty', 0)
        else:
            container_items.append(item)

    for row in ws.iter_rows(values_only=True):
        gj = row[gong_jong_col] if len(row) > gong_jong_col else None
        gj = str(gj).strip() if gj else ""
        name = row[name_col] if len(row) > name_col else None
        name = str(name).strip() if name else ""

        if gj in roman_nums:
            current_district = gj
            current_sub_category = None
            current_sub_sub_category = None
            continue

        if parent_code_re is not None and parent_code_re.match(gj):
            current_parent = name
            # 상위 구분이 바뀌면 하위 상태 초기화
            if current_category:
                if current_sub_category:
                    current_category['sub_categories'].append(current_sub_category)
                    current_sub_category = None
                if current_category.get('items') or current_category.get('sub_categories'):
                    hierarchy.append(current_category)
                current_category = None
            continue

        if major_code_re.match(gj):
            if current_category:
                if current_sub_category:
                    current_category['sub_categories'].append(current_sub_category)
                    current_sub_category = None
                if current_category.get('items') or current_category.get('sub_categories'):
                    hierarchy.append(current_category)
            current_category = {'level': gj, 'name': name, 'parent': current_parent, 'items': [], 'sub_categories': []}
            current_sub_category = None
            continue

        is_hash_separator = bool(
            (hash_paren_re.match(gj) and name and hash_name_re.match(name)) or hash_gj_re.match(gj)
        )
        if is_hash_separator:
            if current_sub_category:
                current_sub_sub_category = {'level': gj, 'name': name, 'district': current_district, 'items': []}
                current_sub_category.setdefault('sub_categories', []).append(current_sub_sub_category)
            continue

        if sub_re.match(gj):
            if current_category:
                if current_sub_sub_category and current_sub_category:
                    if not any(s is current_sub_sub_category for s in current_sub_category['sub_categories']):
                        current_sub_category.setdefault('sub_categories', []).append(current_sub_sub_category)
                current_sub_sub_category = None

                existing_sub = next((s for s in current_category['sub_categories']
                                      if s['level'] == gj and s['name'] == name and s.get('district') == current_district), None)
                if existing_sub:
                    if current_sub_category and current_sub_category is not existing_sub:
                        if not any(s is current_sub_category for s in current_category['sub_categories']):
                            current_category['sub_categories'].append(current_sub_category)
                    current_sub_category = existing_sub
                else:
                    if current_sub_category:
                        if not any(s is current_sub_category for s in current_category['sub_categories']):
                            current_category['sub_categories'].append(current_sub_category)
                    current_sub_category = {'level': gj, 'name': name, 'items': [], 'sub_categories': [], 'district': current_district}
            continue

        # 항목 행: gong_jong 비어있고 name 있음
        if current_category and not gj and name:
            qty_val = row[qty_col] if len(row) > qty_col else None
            unit_val = row[unit_col] if len(row) > unit_col else None
            try:
                qty = float(qty_val) if qty_val else 0
            except (TypeError, ValueError):
                qty = 0
            if qty <= 0:
                continue
            spec = str(row[spec_col]).strip() if len(row) > spec_col and row[spec_col] else ""
            unit = str(unit_val).strip() if unit_val else ""

            code = None
            code_alt = None
            for v in row:
                if not isinstance(v, str):
                    continue
                if code is None and link_regex:
                    m = link_regex.search(v)
                    if m:
                        code = int(m.group(1)) if key_type == "int" else m.group(1)
                if code_alt is None and link_regex_alt:
                    m2 = link_regex_alt.search(v)
                    if m2:
                        code_alt = int(m2.group(1)) if key_type == "int" else m2.group(1)

            item = {'name': name, 'spec': spec, 'qty': qty, 'unit': unit,
                    'district': current_district, 'code': code, 'code_alt': code_alt}

            if (current_sub_sub_category and current_sub_category and
                    "추진" in current_sub_category.get('name', '') and
                    hash_name_re.match(current_sub_sub_category.get('name', ''))):
                _merge_or_append(current_sub_category['items'], item)
            elif current_sub_sub_category:
                _merge_or_append(current_sub_sub_category['items'], item)
            elif current_sub_category:
                _merge_or_append(current_sub_category['items'], item)
            else:
                _merge_or_append(current_category['items'], item)

    if current_category:
        if current_sub_sub_category and current_sub_category:
            if not any(s is current_sub_sub_category for s in current_sub_category['sub_categories']):
                current_sub_category.setdefault('sub_categories', []).append(current_sub_sub_category)
        if current_sub_category:
            if not any(s is current_sub_category for s in current_category['sub_categories']):
                current_category['sub_categories'].append(current_sub_category)
        if current_category.get('items') or current_category.get('sub_categories'):
            hierarchy.append(current_category)

    # 중첩 hierarchy → 평평한 items 리스트로 변환 (category=최상위 대분류 이름, line=지구)
    def _strip_major_prefix(name: str) -> str:
        m = re.match(r"^\d+\.\s*(.+)$", name)
        return m.group(1).strip() if m else name

    flat_items: List[Dict] = []

    def _walk(container, top_category_name, district, parent_name):
        for it in container.get('items', []):
            flat_items.append({
                'name': it['name'], 'spec': it.get('spec', ''), 'qty': it.get('qty', 0),
                'unit': it.get('unit', ''), 'code': it.get('code'), 'code_alt': it.get('code_alt'),
                'category': top_category_name, 'line': it.get('district', district),
                'parent': parent_name,
            })
        for sub in container.get('sub_categories', []):
            _walk(sub, top_category_name, sub.get('district', district), parent_name)

    for cat in hierarchy:
        top_name = _strip_major_prefix(cat['name'])
        _parent = cat.get('parent')
        _parent = _strip_major_prefix(_parent) if _parent else None
        _walk(cat, top_name, None, _parent)

    return flat_items


def parse_items_generic(wb, tmpl: Dict[str, Any]) -> List[Dict]:
    """템플릿의 hierarchy_strategy에 따라 적절한 파서로 위임."""
    ws = _get_sheet(wb, tmpl["item_sheet_names"])
    if ws is None:
        return []
    strategy = tmpl["hierarchy_strategy"]
    if strategy == "major_number_prefix":
        return _parse_items_major_number_prefix(ws, tmpl)
    elif strategy == "district_roman":
        return _parse_items_district_roman(ws, tmpl)
    else:
        raise ValueError(f"알 수 없는 hierarchy_strategy: {strategy}")


def parse_with_template(wb, tmpl: Dict[str, Any]):
    """편의 함수: 항목 + 단가근거를 한번에 파싱."""
    items = parse_items_generic(wb, tmpl)
    unit_prices = parse_unit_price_generic(wb, tmpl)
    return items, unit_prices


def parse_auto(wb):
    """템플릿 자동판별 후 파싱. 매칭 실패시 (None, [], {}) 반환."""
    tmpl = detect_template(wb)
    if tmpl is None:
        return None, [], {}
    items, unit_prices = parse_with_template(wb, tmpl)
    return tmpl, items, unit_prices


# ---------------------------------------------------------------------------
# CLI 검증
# ---------------------------------------------------------------------------

def _main(argv):
    import warnings
    warnings.filterwarnings("ignore")
    from openpyxl import load_workbook

    if len(argv) < 2:
        print("사용법: python universal_parser.py <엑셀경로> [템플릿id 강제지정]", file=sys.stderr)
        return 2
    path = argv[1]
    force_id = argv[2] if len(argv) > 2 else None

    print(f"[로드] {path}")
    wb = load_workbook(path, read_only=True, data_only=True)

    if force_id:
        tmpl = next((t for t in TEMPLATES if t["id"] == force_id), None)
        if tmpl is None:
            print(f"알 수 없는 템플릿id: {force_id}")
            return 2
    else:
        tmpl = detect_template(wb)

    if tmpl is None:
        print("❌ 매칭되는 템플릿이 없습니다. (신규 양식 — TEMPLATES에 추가 필요)")
        print(f"   시트 목록: {wb.sheetnames}")
        return 1

    print(f"[템플릿] {tmpl['name']} ({tmpl['id']})")
    items, unit_prices = parse_with_template(wb, tmpl)
    print(f"[파싱] 항목 {len(items)}개, 단가근거 {len(unit_prices)}개")

    matched = [it for it in items if it.get("code") in unit_prices]
    rate = (len(matched) / len(items) * 100) if items else 0.0
    print(f"[매칭] {len(matched)}개 ({rate:.1f}%)")

    from collections import Counter
    cat_counts = Counter(it.get("category") for it in items if it.get("category"))
    print("[대공종별 항목수]")
    for cat, cnt in cat_counts.most_common(15):
        print(f"    {cat:15} {cnt}개")

    return 0



# ---------------------------------------------------------------------------
# 노무비 역산 기반 일작업량 추정 (Q=/HR 산식이 없는 블록용)
# ---------------------------------------------------------------------------
#
# 배경: 적산 프로그램이 출력한 엑셀은 일위대가 블록에 원가(재료비/노무비/경비)만
# 담고, 표준품셈 인수(공수) 자체는 프로그램 내부 DB에 있어 엑셀에 안 나온다.
# 실측 결과 화성 R18 파일 기준 일위대가_호표 236개 블록 중 222개가 이 경우라
# 관로공·구조물공·가시설공·추진공·배수설비공이 통째로 매칭 0%가 된다.
#
# 대안: 블록의 노무비를 해당 직종의 노임단가로 나누면 소요 인수(인·일)가 나오고,
# 그 역수가 1인 기준 일작업량이 된다.
#     일작업량 = (기준수량 × 노임단가) / 노무비
#
# 직종 선택 규칙 (사용자 확정):
#   1순위 항목명 키워드 → 2순위 대공종 → 3순위 보통인부(일반노무)
# 같은 직종명이 여러 등급으로 중복될 경우 보수적으로 '낮은 단가'를 채택한다
# (단가가 낮을수록 역산 인수가 커져 일작업량이 작게 = 공기가 길게 잡힘).

# 항목명 키워드 → 직종 (1순위)
_LABOR_BY_ITEM = [
    (("주철관", "PVC", "PE다중벽", "GRP", "복합관", "접합", "부설", "배수설비", "관로", "수밀"), "배관공"),
    (("맨홀", "콘크리트", "구조물", "PC", "타설", "거푸집"), "콘크리트공"),
    (("아스팔트", "포장", "기층", "표층", "코팅"), "포장공"),
    (("추진", "착암", "천공", "갱구"), "착암공"),
    (("비계", "가시설", "흙막이", "동바리"), "비계공"),
    (("전기", "전선", "케이블", "내선"), "내선전공"),
    (("통신",), "통신외선공"),
    (("석축", "석공", "돌쌓기"), "석공"),
    (("목공", "목재"), "건축목공"),
]

# 대공종 → 직종 (2순위)
_LABOR_BY_CATEGORY = {
    "관로공": "배관공",
    "관접합 및 부설": "배관공",
    "배수설비공": "배관공",
    "구조물공": "콘크리트공",
    "포장공": "포장공",
    "추진공": "착암공",
    "관 추진공": "착암공",
    "가시설공": "비계공",
}

_DEFAULT_LABOR = "보통인부"

# 역산 제외 단위: 시공 물량이 아니라 인력·비용·기간성 항목
# (예: '교통정리신호수 2,733일' = 상주 인력 투입량, '품질관리비 1식' = 비용)
# 이런 항목에 역산을 적용하면 수량이 그대로 일수로 잡혀 공기가 폭증한다.
_NON_WORK_UNITS = {"일", "식", "인", "월", "년", "회", "대", "человек"}


def parse_wage_table(wb) -> Dict[str, float]:
    """'노임' 시트에서 {직종명(공백제거): 단가} 반환. 동일 직종 중복 시 낮은 단가 채택."""
    if wb is None or "노임" not in wb.sheetnames:
        return {}
    ws = wb["노임"]
    wages: Dict[str, float] = {}
    for r in ws.iter_rows(values_only=True):
        name = r[0] if len(r) > 0 else None
        rate = r[3] if len(r) > 3 else None
        if isinstance(name, str) and isinstance(rate, (int, float)) and rate > 0:
            key = name.replace(" ", "").strip()
            if key not in wages or rate < wages[key]:
                wages[key] = float(rate)
    return wages


def pick_labor_type(item_name: str, category: str) -> str:
    """항목명 → 대공종 → 기본값 순으로 직종 결정."""
    nm = item_name or ""
    for keywords, job in _LABOR_BY_ITEM:
        if any(k in nm for k in keywords):
            return job
    cat = (category or "").strip()
    if cat in _LABOR_BY_CATEGORY:
        return _LABOR_BY_CATEGORY[cat]
    return _DEFAULT_LABOR


def parse_labor_derived_rates(wb, sheet_name: str = "일위대가_호표",
                              code_col: int = 14, name_col: int = 0,
                              qty_col: int = 2, unit_col: int = 3,
                              labor_col: int = 9) -> Dict[str, Tuple[float, str, str]]:
    """
    일위대가 블록의 노무비를 노임단가로 나눠 일작업량을 역산.

    Returns
    -------
    dict[str, tuple[float, str, str]]
        {코드: (일작업량, 단위, 사용직종)}
    """
    if wb is None or sheet_name not in wb.sheetnames:
        return {}
    wages = parse_wage_table(wb)
    if not wages:
        return {}

    ws = wb[sheet_name]
    rows = [r for r in ws.iter_rows(values_only=True)]
    result: Dict[str, Tuple[float, str, str]] = {}

    for r in rows:
        code = r[code_col] if len(r) > code_col else None
        if not isinstance(code, str) or not code.strip():
            continue
        code = code.strip()
        if code in result:
            continue

        name = r[name_col] if len(r) > name_col else ""
        name = str(name) if name else ""
        base_qty = r[qty_col] if len(r) > qty_col else None
        unit = r[unit_col] if len(r) > unit_col else ""
        labor_cost = r[labor_col] if len(r) > labor_col else None

        if not isinstance(base_qty, (int, float)) or base_qty <= 0:
            continue
        if not isinstance(labor_cost, (int, float)) or labor_cost <= 0:
            continue

        unit_s = str(unit).strip() if unit else ""
        # 시공 물량이 아닌 단위(일/식/인 등)는 역산 대상에서 제외
        if unit_s in _NON_WORK_UNITS or not unit_s:
            continue

        job = pick_labor_type(name, "")
        wage = wages.get(job.replace(" ", "")) or wages.get(_DEFAULT_LABOR.replace(" ", ""))
        if not wage:
            continue

        manday_per_unit = labor_cost / wage / base_qty
        if manday_per_unit <= 0:
            continue
        daily = round(1.0 / manday_per_unit, 4)
        result[code] = (daily, unit_s, job)

    return result




def parse_labor_derived_by_template(wb, tmpl: Dict[str, Any]) -> Dict[Any, Tuple[float, str, str]]:
    """
    템플릿 설정(tmpl["labor"])에 따라 노무비 역산으로 일작업량을 구한다.
    두 가지 블록 구조를 모두 지원:
      (a) block_code_col 지정  → 헤더행에 코드·수량·노무비가 함께 있는 형태 (화성 양식)
      (b) block_start_regex 지정 → '제 N호표' 헤더 + 별도 합계행에 노무비 (표준형 양식)
    반환: {키: (일작업량, 단위, 사용직종)}
    """
    cfg = tmpl.get("labor")
    if not wb or not cfg:
        return {}
    if cfg["wage_sheet"] not in wb.sheetnames or cfg["block_sheet"] not in wb.sheetnames:
        return {}

    # 1) 노임표 (동일 직종 중복 시 낮은 단가 = 보수적)
    wages: Dict[str, float] = {}
    for r in wb[cfg["wage_sheet"]].iter_rows(values_only=True):
        nc, rc = cfg["wage_name_col"], cfg["wage_rate_col"]
        nm = r[nc] if len(r) > nc else None
        rt = r[rc] if len(r) > rc else None
        if isinstance(nm, str) and isinstance(rt, (int, float)) and rt > 0:
            k = nm.replace(" ", "").strip()
            if k and (k not in wages or rt < wages[k]):
                wages[k] = float(rt)
    if not wages:
        return {}

    rows = [r for r in wb[cfg["block_sheet"]].iter_rows(values_only=True)]
    result: Dict[Any, Tuple[float, str, str]] = {}

    _BASE_RE = re.compile(r"^([\d,\.]+)\s*([A-Za-z가-힣㎡㎥]+)$")

    def _explicit_base(s, e):
        """블록에 '합계 | 520M' + '적용 | M당' 구조가 있으면 그 실제 기준수량을 반환.

        일부 일위대가(예: 하수관CCTV조사)는 '1일에 3인이 520M'을 산출한 뒤
        M당 단가로 환산해 표기한다. 이때 헤더 수량 '1M'은 단가 기준일 뿐이고,
        인수 3인의 실제 기준은 520M이다. 이를 구분하지 않으면
        1M÷3인 = 0.33M/인일 같은 비현실적 값이 나온다(실측 확인).
        """
        base = None
        has_apply = False
        for j in range(s + 1, e):
            c0 = str(rows[j][0] or "").strip()
            c1 = rows[j][1] if len(rows[j]) > 1 else None
            if c0.startswith("합") and isinstance(c1, str):
                m = _BASE_RE.match(c1.strip())
                if m:
                    try:
                        base = float(m.group(1).replace(",", ""))
                    except ValueError:
                        base = None
            if c0.startswith("적용"):
                has_apply = True

            # 변형 표기: '1개소당 | 100 | 식' 처럼 '1<단위>당' 행의 수량이 기준수량인 경우.
            # 실측(풍각): 플랫타이 구멍메우기는 보통인부 0.62인이 '100개소' 기준인데
            # 헤더 수량이 비어 있어 1개소로 잘못 계산되면 1.6개소/일(공식 161개소/일의 1/100)이 된다.
            if base is None and re.match(r"^1\s*[가-힣㎡㎥A-Za-z]+\s*당$", c0):
                q_ = rows[j][2] if len(rows[j]) > 2 else None
                if isinstance(q_, (int, float)) and q_ > 1:
                    base = float(q_)
                    has_apply = True

        return base if (has_apply and base and base > 0) else None

    _SURCHARGE_KW = ("할증", "야간", "요율", "%")

    def _labor_excl_surcharge(s, e, lac):
        """블록 하위 행의 노무비 합에서 할증·요율 성격 행을 제외한 실노무비.

        동일 규격인데 야간 블록에만 '야간할증 | 노무비의 87.5%' 행이 붙어
        노무비가 1.87배가 되고(명륜 실측: 흄관 D600 135,265 → 253,621),
        이를 그대로 역산하면 야간 작업량이 주간의 절반으로 잘못 계산된다.
        야간은 단가가 비싼 것이지 시공 속도가 절반인 게 아니므로 할증을 뺀다.
        """
        tot = 0.0
        for j in range(s + 1, e):
            r = rows[j]
            nm = str(r[0] or "")
            u = str(r[3] or "") if len(r) > 3 else ""
            if any(k in nm for k in _SURCHARGE_KW) or u.strip() == "%":
                continue
            v = r[lac] if len(r) > lac else None
            if isinstance(v, (int, float)) and v > 0:
                tot += float(v)
        return tot if tot > 0 else None

    # (제목, 규격) → 블록 범위. 하위 참조 블록의 인수를 따라가기 위한 색인.
    _block_by_title = {}

    def _index_blocks(starts_list):
        for _n, (_i, _e) in enumerate(starts_list):
            _t = str(rows[_i + 1][0] or "").strip() if _i + 1 < len(rows) else ""
            _sp = str(rows[_i + 1][1] or "").strip() if _i + 1 < len(rows) else ""
            if _t and (_t, _sp) not in _block_by_title:
                _block_by_title[(_t, _sp)] = (_i, _e)

    def _manhours(s, e, depth=0):
        """블록 [s+1, e) 안에서 단위가 '인'인 행의 수량(=소요 인수) 합.

        노무비 합계에는 야간할증·요율(설치간격)·공구손료·공장관리비 같은
        '인건비가 아닌' 계수가 섞여 있어(실측: 가시설 안전난간 노무비 19,130 중
        요율 19,057이 계수), 노무비로 역산하면 작업량이 과소평가되고 특히
        -야간 항목은 할증 때문에 주간의 절반으로 잡히는 구조적 오류가 생긴다.
        직접 인수만 합산하면 할증·요율이 제외되어 야간도 주간과 같은 값이 나온다.

        또한 '철근 공장가공 및 조립'처럼 하위가 다른 일위대가를 참조해 직접 인수가
        없는 블록은, 그 하위 블록의 인수를 재귀적으로 끌어와 합산한다
        (실측: 재귀 미적용 시 0.23TON/인일 → 적용 시 2.94TON/인일, 공식 4TON).
        """
        tot = 0.0
        for j in range(s + 1, e):
            r = rows[j]
            u = str(r[3]).strip() if len(r) > 3 and r[3] else ""
            q = r[2] if len(r) > 2 else None
            if not isinstance(q, (int, float)) or q <= 0:
                continue
            if u == "인":
                tot += float(q)
            elif depth < 2 and u not in ("", "%"):
                nm = str(r[0] or "").strip()
                sp = str(r[1] or "").strip() if len(r) > 1 and r[1] else ""
                sub = _block_by_title.get((nm, sp))
                if sub and sub[0] != s:
                    tot += float(q) * _manhours(sub[0], sub[1], depth + 1)
        return tot

    def _emit(key, title, unit, base_qty, labor_cost, manhours=0.0, explicit_base=None):
        unit_s = str(unit).strip() if unit else ""
        if not unit_s or unit_s in _NON_WORK_UNITS:
            return
        if not (isinstance(base_qty, (int, float)) and base_qty > 0):
            return
        job = pick_labor_type(str(title or ""), "")

        # 1순위: 직접 인수 기반 (할증·요율 배제, 야간=주간 동일값)
        if manhours and manhours > 0:
            _bq = explicit_base if (explicit_base and explicit_base > 0) else base_qty
            daily = _bq / manhours
            if daily > 0 and key not in result:
                result[key] = (round(daily, 4), unit_s, job)
            return

        # 2순위: 노무비 ÷ 노임단가 (하위가 다른 일위대가를 참조해 직접 인수가 없는 블록)
        if not (isinstance(labor_cost, (int, float)) and labor_cost > 0):
            return
        w = wages.get(job.replace(" ", "")) or wages.get(_DEFAULT_LABOR.replace(" ", ""))
        if not w:
            return
        daily = (base_qty * w) / labor_cost
        if daily > 0 and key not in result:
            result[key] = (round(daily, 4), unit_s, job)

    if "block_code_col" in cfg:
        # (a) 헤더행 일체형 — 코드행 = 블록 시작, 다음 코드행 직전까지가 그 블록
        cc = cfg["block_code_col"]
        code_idx = [(i, str(r[cc]).strip()) for i, r in enumerate(rows)
                    if len(r) > cc and isinstance(r[cc], str) and str(r[cc]).strip()]
        # 하위 참조 인수 재귀 합산용 색인 (헤더행이 곧 제목)
        _tc = cfg["title_col"]
        for _n, (_i, _c) in enumerate(code_idx):
            _e = code_idx[_n + 1][0] if _n + 1 < len(code_idx) else len(rows)
            _t = str(rows[_i][_tc] or "").strip() if len(rows[_i]) > _tc else ""
            _sp = str(rows[_i][1] or "").strip() if len(rows[_i]) > 1 and rows[_i][1] else ""
            if _t and (_t, _sp) not in _block_by_title:
                _block_by_title[(_t, _sp)] = (_i, _e)
        for n, (i, code) in enumerate(code_idx):
            e = code_idx[n + 1][0] if n + 1 < len(code_idx) else len(rows)
            r = rows[i]
            _lac = cfg["labor_amount_col"]
            _lab_raw = r[_lac] if len(r) > _lac else None
            _lab_net = _labor_excl_surcharge(i, e, _lac)
            _emit(code,
                  r[cfg["title_col"]] if len(r) > cfg["title_col"] else "",
                  r[cfg["unit_col"]] if len(r) > cfg["unit_col"] else "",
                  r[cfg["qty_col"]] if len(r) > cfg["qty_col"] else None,
                  _lab_net if _lab_net else _lab_raw,
                  _manhours(i, e), _explicit_base(i, e))
    else:
        # (b) '제 N호표' 블록형 — 합계행의 노무비 사용
        pat = re.compile(cfg["block_start_regex"])
        sc = cfg["block_start_col"]
        starts = []
        for i, r in enumerate(rows):
            c = r[sc] if len(r) > sc else None
            if isinstance(c, str):
                m = pat.search(c)
                if m:
                    starts.append((i, m.group(1)))
        # 하위 참조 인수 재귀 합산용 색인 (제목은 블록 시작 다음 행)
        _index_blocks([(st_i, (starts[k + 1][0] if k + 1 < len(starts) else len(rows)))
                       for k, (st_i, _) in enumerate(starts)])
        for idx, (s, no) in enumerate(starts):
            e = starts[idx + 1][0] if idx + 1 < len(starts) else len(rows)
            ti = s + cfg.get("title_offset", 1)
            if ti >= len(rows):
                continue
            title = rows[ti][cfg["title_col"]] if len(rows[ti]) > cfg["title_col"] else ""
            unit = rows[ti][cfg["unit_col"]] if len(rows[ti]) > cfg["unit_col"] else ""
            labor_cost = None
            kw = cfg.get("sum_row_keyword", "합")
            lac = cfg["labor_amount_col"]
            for j in range(s, e):
                c0 = rows[j][0] if len(rows[j]) > 0 else None
                if isinstance(c0, str) and kw in c0:
                    v = rows[j][lac] if len(rows[j]) > lac else None
                    if isinstance(v, (int, float)) and v > 0:
                        labor_cost = v
            key = int(no) if cfg.get("key_type") == "int" else no
            _emit(key, title, unit, cfg.get("base_qty", 1.0), labor_cost, _manhours(ti, e), _explicit_base(ti, e))

    return result


# ---------------------------------------------------------------------------
# 동일 계열 Q값 승계 (기계 위주 항목의 노무비 역산 과소평가 보정)
# ---------------------------------------------------------------------------
#
# 배경: '터파기(B=6.0m이상) 토사,육상,기계90%+인력10%'처럼 기계 비중이 높은 항목이
# 산근에 Q산식 없이 호표에만 있으면 노무비 역산으로 떨어지는데, 기계비는 경비로
# 빠지고 노무비엔 인력 10%만 잡혀 실제보다 크게 과소평가된다.
#   실측 예: B=4~6m는 Q산식 138.2㎥/일인데, 조건이 동일한 B=6m이상은 역산 32.6㎥/일.
#           운반-토사는 Q산식 422.4 vs 역산 39.7 (10배 차이).
#
# 보정: 같은 (기본항목명, 규격) 그룹에 Q산식 값이 존재하면, 역산으로 빠진 항목은
# 그 그룹의 Q값 최댓값을 승계한다. 장비를 바꾸지 않는 한 굴착폭이 넓을수록
# 작업량은 크거나 같다는 실무 사실과도 부합한다(사용자 확인).

def _series_key(name: str, spec: str):
    """괄호 앞 기본명 + 규격을 계열 키로. '터파기(B=6.0m이상)' → '터파기'."""
    base = (name or "").split("(")[0].strip()
    return (base, (spec or "").strip())


# 시공 조건이 같은데 이름·규격 표기만 다른 계열을 잇기 위한 정규화.
# 실측(화성): 아스팔트 기층과 표층은 시공폭 구분이 1:1로 대응하는데
#   기층 'T=250mm, 3.0m≤B'  /  표층 '3.0m≤B'
# 처럼 기층에만 두께 접두사가 붙어 있어 계열 매칭이 되지 않았고,
# 그 결과 기층은 노무비 역산(123㎡/일)으로 떨어져 표층(4800㎡/일)의 1/39이 되었다.
# 같은 장비·인원으로 하는 포설 작업이므로 시공폭이 같으면 같은 계열로 본다.
_LAYER_WORDS = ("기층", "표층", "중간층")
_SPEC_DROP_RE = re.compile(r"^\s*T\s*=\s*[\d.]+\s*(mm|cm|㎜|㎝)?\s*,\s*", re.IGNORECASE)


def _series_key_loose(name: str, spec: str):
    """느슨한 계열 키: 층 구분어(기층/표층)와 규격의 두께 접두사를 제거.

    '아스팔트기층' + 'T=250mm, 3.0m≤B'  →  ('아스팔트', '3.0m≤B')
    '아스팔트표층' + '3.0m≤B'           →  ('아스팔트', '3.0m≤B')
    """
    base = (name or "").split("(")[0].strip()
    for w in _LAYER_WORDS:
        base = base.replace(w, "")
    base = re.sub(r"\s+", "", base)
    sp = _SPEC_DROP_RE.sub("", (spec or "").strip())
    return (base, sp.strip())


def apply_series_fallback(items, unit_prices, labor_rates):
    """
    items: parse_items_generic 결과
    unit_prices: {코드: (일작업량, 단위)}
    labor_rates: {키: (일작업량, 단위, 직종)}
    반환: {승계키: (일작업량, 단위)} — 역산 대신 쓸 Q승계 값
    """
    from collections import defaultdict
    series_q = defaultdict(list)
    series_q_loose = defaultdict(list)
    fallback = {}

    # 1) 계열별 Q산식 값 수집 (엄격 키 + 느슨한 키 둘 다)
    for it in items:
        c = it.get("code")
        if c in unit_prices:
            d, u = unit_prices[c]
            if d and d > 0:
                series_q[_series_key(it.get("name"), it.get("spec"))].append((d, u))
                series_q_loose[_series_key_loose(it.get("name"), it.get("spec"))].append((d, u))

    # 2) 역산으로 빠진 항목에 같은 계열 Q 최댓값 승계
    #    엄격 키를 먼저 보고, 없으면 느슨한 키(층 구분어·두께 접두사 제거)로 재시도한다.
    for it in items:
        c = it.get("code")
        if c in unit_prices:
            continue
        # 산근 번호와 일위대가표 번호는 별개 체계이므로 폴백하지 않는다
        k = it.get("code_alt") if it.get("code_alt") is not None else c
        if k not in labor_rates:
            continue
        cand = series_q.get(_series_key(it.get("name"), it.get("spec")))
        if not cand:
            cand = series_q_loose.get(_series_key_loose(it.get("name"), it.get("spec")))
        if not cand:
            continue
        best_d, best_u = max(cand, key=lambda x: x[0])
        lab_d = labor_rates[k][0]
        # 역산값이 계열 Q 최댓값보다 작을 때만 승계 (과소평가 보정 목적)
        if best_d > lab_d:
            if c is not None:
                fallback[c] = (best_d, best_u)
            if k is not None:
                fallback[k] = (best_d, best_u)
    return fallback


if __name__ == "__main__":
    sys.exit(_main(sys.argv))