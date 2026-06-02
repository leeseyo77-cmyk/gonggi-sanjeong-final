# -*- coding: utf-8 -*-
"""
hopyo_parser.py — 단가산출근거 호표별 일작업량 파서 (형식1: Q=/HR 또는 Q1=/HR)

배경
----
설계자가 단가산출근거에 직접 산정한 시공량(시간당 작업량 Q)을 호표번호로
1:1 매칭해서, 가이드라인보다 우선 적용하기 위한 데이터 소스를 만든다.

처리 대상 형식
--------------
형식1만 처리 (1차 구현):
  "Q = 60/Tc = X.XX 단위/HR"   (천공후항타)
  "Q1 = 60/Tc1 = X.XX 단위/HR" (항발)
변수명이 Q/Q1 어느 쪽이든 동일하게 매칭 (정규식이 `=`만 기준).
시간당 → 일당 환산은 ×8.

전제 조건
---------
- 워크북은 반드시 `data_only=True` 로 로드되어야 함.
  단가산출근거 시트의 Q값 텍스트는 Excel 수식 결과로 만들어지므로,
  data_only=False 로 열면 수식 원문이 들어와 정규식이 매칭되지 않는다.
- 호표 시작 라인 '제 N호표'와 Q 산식 라인 둘 다 열 인덱스 1에 있음.

반환 형식
---------
{호표번호(int): (일작업량(float, 본/일 등), 단위(str))}

호표 228 단일 예외
------------------
'열연강판 항타 및 항발' 호표는 한 블록에 Q=(항타)와 Q1=(항발) 두 줄이 있다.
현재 구현은 문서 순서상 첫 줄(항타)을 사용한다 → 보수적(작은 값) 선택.

표준 사용
---------
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    hopyo_daily = parse_hopyo_daily_amounts(wb)
    # 예) hopyo_daily[74] == (34.48, '본')

CLI 검증
--------
    python hopyo_parser.py <엑셀경로>
"""

from __future__ import annotations

import re
import sys
from typing import Dict, Tuple

# 열 인덱스 1의 '제 N호표' 패턴 (블록 시작 표시)
_HOPYO_BLOCK_RE = re.compile(r"제\s*(\d+)\s*호표")

# 'Q = ... = X 단위/HR' 또는 'Q1 = ... = X 단위/HR'에서 X와 단위 추출.
# 둘 다 마지막 `= 숫자 단위/HR` 구조라 동일한 정규식으로 잡힌다.
_HOPYO_Q_RE = re.compile(r"=\s*([\d.]+)\s*([가-힣㎡㎥]+)/HR")

# 시간당 → 일당 환산 계수 (8시간 작업)
_HR_TO_DAY = 8


def parse_hopyo_daily_amounts(wb) -> Dict[int, Tuple[float, str]]:
    """
    단가산출근거 시트를 1회 스캔해서 호표번호 → (일작업량, 단위) dict 반환.

    Parameters
    ----------
    wb : openpyxl.Workbook
        반드시 ``data_only=True`` 로 로드된 워크북.

    Returns
    -------
    dict[int, tuple[float, str]]
        {호표번호: (일작업량, 단위)}. 형식1을 못 찾은 호표는 포함되지 않음.
    """
    if wb is None or "단가산출근거" not in wb.sheetnames:
        return {}

    ws = wb["단가산출근거"]
    rows = [r for r in ws.iter_rows(values_only=True)]

    # 1) 각 호표 시작 행 인덱싱 (첫 출현만 채택)
    starts: Dict[int, int] = {}
    for i, r in enumerate(rows):
        c1 = r[1] if len(r) > 1 else None
        if isinstance(c1, str):
            m = _HOPYO_BLOCK_RE.search(c1)
            if m:
                n = int(m.group(1))
                if n not in starts:
                    starts[n] = i
    if not starts:
        return {}

    # 2) 각 호표 블록(다음 호표 시작 직전까지) 안에서 첫 Q=/HR 라인 추출
    ordered = sorted(starts.items(), key=lambda kv: kv[1])
    result: Dict[int, Tuple[float, str]] = {}
    for idx, (n, s) in enumerate(ordered):
        end = ordered[idx + 1][1] if idx + 1 < len(ordered) else len(rows)
        for j in range(s, end):
            c1 = rows[j][1] if len(rows[j]) > 1 else None
            if isinstance(c1, str):
                m = _HOPYO_Q_RE.search(c1)
                if m:
                    daily = round(float(m.group(1)) * _HR_TO_DAY, 4)
                    result[n] = (daily, m.group(2))
                    break
    return result


def parse_hopyo_daily_amounts_from_path(path: str) -> Dict[int, Tuple[float, str]]:
    """파일 경로에서 직접 로드해서 파싱. 통합 외 단독 실행/테스트용 헬퍼."""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        return parse_hopyo_daily_amounts(wb)
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# CLI 검증: python hopyo_parser.py <엑셀경로>
# ---------------------------------------------------------------------------

# (호표번호, 기대 일작업량, 허용오차, 설명)
_VERIFY_POINTS = [
    (1,   32.56, 0.05, "H-PILE 천공후항타 H=3.55m"),
    (40,  58.48, 0.05, "SHEET PILE 천공후항타 H=3.8m  (검증치 58.5)"),
    (46,  54.00, 0.05, "SHEET PILE 천공후항타 H=4.48m (검증치 54.0)"),
    (74,  34.48, 0.05, "SHEET PILE 천공후항타 H=9.5m  (검증치 34.5)"),
    (15,  60.00, 0.05, "H-PILE 항발 H=3.55m"),
    (25,  51.92, 0.05, "H-PILE 항발 H=6.21m"),
    (110, 39.60, 0.05, "SHEET PILE 항발 H=9.5m"),
    (228, 30.32, 0.05, "열연강판 항타 및 항발 (첫 Q라인=항타)"),
]


def _main(argv):
    import warnings
    warnings.filterwarnings("ignore")  # openpyxl Print area 경고 억제

    if len(argv) < 2:
        print("사용법: python hopyo_parser.py <엑셀경로>", file=sys.stderr)
        return 2
    path = argv[1]

    print(f"[로드] {path}")
    result = parse_hopyo_daily_amounts_from_path(path)
    print(f"[파싱] 형식1(Q=/HR | Q1=/HR) 보유 호표: {len(result)}개")

    # 단위 분포 (대부분 '본'이어야 함)
    from collections import Counter
    units = Counter(u for _, u in result.values())
    print(f"[단위] {dict(units)}")

    # 검증 포인트 평가
    print()
    print("=== 검증 포인트 ===")
    ok = 0
    fail = 0
    for n, expected, tol, desc in _VERIFY_POINTS:
        if n not in result:
            print(f"  ✗ 호표 {n:>3}  {desc}  → 매칭 안 됨 (기대 {expected})")
            fail += 1
            continue
        daily, unit = result[n]
        diff = abs(daily - expected)
        status = "✓" if diff <= tol else "✗"
        print(f"  {status} 호표 {n:>3}  {desc}")
        print(f"        실제 {daily}{unit}, 기대 {expected}, 오차 {diff:.4f} (허용 {tol})")
        if diff <= tol:
            ok += 1
        else:
            fail += 1

    print()
    print(f"=== 결과: {ok}/{len(_VERIFY_POINTS)} 통과, {fail} 실패 ===")
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    sys.exit(_main(sys.argv))