# -*- coding: utf-8 -*-
"""
guideline_work.py — 「2026년 적정 공사기간 확보를 위한 가이드라인」 부록4 1일 작업량(건축분야 등)

부록4는 "표준품셈의 생산성을 기준(1개 작업조를 투입한 결과)으로 산출된" 1일 작업량이다(p95).
앱의 투입조수 1 = 그 1개 작업조다.

적용 순위(사용자 기준 2026-11-02): 내역서 Q값(산근·호표·A·Q) → 가이드라인 → 품셈·노무비 역산.
가이드라인 조건(난이도·높이·규격)이 내역서 규격과 맞을 때만 쓴다. 조건이 다르면(유로폼 '복잡',
높이 초과 등) 가이드라인에 값이 없는 것으로 보고 다음 순위로 넘긴다 — 엉뚱한 조건의 값을 쓰지 않게.

설치 및 해체처럼 한 항목에 여러 작업이 묶이면 같은 조가 순서대로 하는 것으로 보고 일수를 더한다
(1일 작업량 = 1 ÷ Σ 1/작업량_i — 산근 A·Q 따라가기와 같은 원칙).

쪽 번호는 가이드라인 PDF 인쇄 쪽(부록4 2. 건축분야 p105~p116).
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Tuple

PAGE = {"가설": 105, "토공": 106, "기초": "107~108", "철근콘크리트": "109~111", "철골": 112,
        "조적": 112, "미장": 112, "방수": 113, "목공": 113, "창호": 114, "타일": 114, "수장": "114~115",
        "조경": "115~116"}

_FORM_WORK = {"형틀목공": 4, "보통인부": 1}


def _norm(s: str) -> str:
    s = re.sub(r"\s+", "", str(s or "")).upper()
    return (s.replace("∼", "~").replace("～", "~").replace("〜", "~").replace("III", "Ⅲ")
            .replace("II", "Ⅱ").replace("Ｍ", "M").replace("㎜", "MM"))


def unit_key(u: str) -> str:
    """단위 표기 통일(m/M, ㎡/M2, 공㎥/공M3, ton/TON/톤, 개소/개/EA)."""
    u = _norm(u).split("/")[0]
    u = u.replace("㎡", "M2").replace("㎥", "M3").replace("M²", "M2").replace("M³", "M3")
    if u in ("TON", "톤", "T"):
        return "TON"
    if u in ("개소", "개", "EA", "틀", "SET"):
        return "EA"
    return u


def unit_ok(rate_unit: str, item_unit: str) -> bool:
    a, b = unit_key(rate_unit), unit_key(item_unit)
    return not a or not b or a == b


def _ops(n: str, default_both: bool = True) -> Tuple[str, ...]:
    """항목명이 말하는 작업: 설치·해체 둘 다 / 하나. 둘 다 안 쓰였으면 default_both."""
    i = any(k in n for k in ("설치", "조립", "붙임", "부착"))
    r = any(k in n for k in ("해체", "철거"))
    if i and r:
        return ("설치", "해체")
    if r:
        return ("해체",)
    if i:
        return ("설치",)
    return ("설치", "해체") if default_both else ("설치",)


def _over(t: str, lim: float) -> bool:
    """규격의 높이가 lim(m)를 넘는가. 'Xm초과'(X ≥ lim)나, 초과 표기 없이 'Ym이하'(Y > lim)."""
    for m in re.finditer(r"(\d+(?:\.\d+)?)M초과", t):
        if float(m.group(1)) >= lim - 1e-9:
            return True
    if "초과" not in t:
        for m in re.finditer(r"(\d+(?:\.\d+)?)M(?:까지|이하)", t):
            if float(m.group(1)) > lim + 1e-9:
                return True
    return False


def _difficulty(t: str) -> Optional[str]:
    for d in ("복잡", "보통", "간단"):
        if d in t:
            return d
    return None


def _mm(t: str) -> Optional[float]:
    m = re.search(r"(\d+(?:\.\d+)?)(?:MM|T\b|T\*|T,)", t)
    return float(m.group(1)) if m else None


Step = Tuple[str, float, Dict[str, int]]        # (작업, 1일 작업량, 작업조)


def _steps(n: str, t: str, unit: str, ctx: str) -> Optional[Tuple[str, Any, str, List[Step]]]:
    """(공종 표기, 쪽, 단위, [작업 단계]) 또는 None. n=항목명, t=항목명|규격 (정규화)."""
    bld = ctx != "토목"

    # ── 가설 ─────────────────────────────────────────────────────────
    if "동바리" in n:
        both = _ops(n) == ("설치", "해체")
        if "시스템" in n and both and not _over(t, 10):
            return "시스템동바리 설치 및 해체(10m 이하)", PAGE["가설"], "공M3", [
                ("설치 및 해체", 69, {"형틀목공": 4, "보통인부": 1})]
        if "강관" in n and both and not _over(t, 3.5):
            if bld:
                return "강관동바리 설치 및 해체(건축, 3.5m 이하)", PAGE["가설"], "M2", [
                    ("설치 및 해체", 100, {"형틀목공": 5, "보통인부": 1})]
            return "강관동바리 설치 및 해체(토목, 2.5~3.5m)", PAGE["가설"], "공M3", [
                ("설치 및 해체", 51, {"형틀목공": 3, "보통인부": 1})]
        return None
    if "비계" in n and not any(k in n for k in ("브라켓", "외부", "달비계", "말비계", "수평")):
        if "시스템" in n and _ops(n) == ("설치", "해체") and not _over(t, 10):
            return "시스템비계 설치 및 해체(10m 이하)", PAGE["가설"], "M2", [
                ("설치 및 해체", 100, {"비계공": 4, "보통인부": 1})]
        if "강관" in n and _ops(n, False) == ("설치",) and not _over(t, 10):
            return "강관비계 설치(쌍줄, 10m 이하)", PAGE["가설"], "M2", [
                ("설치", 60, {"비계공": 3, "보통인부": 1})]
        return None
    if "가설울타리" in n and ("EGI" in t or "휀스" in t or "펜스" in t) and _ops(n, False) == ("설치",):
        return "가설울타리(EGI휀스) 설치", PAGE["가설"], "M", [("설치", 115, {"비계공": 3, "보통인부": 1})]

    # ── 토공·기초(건축) ───────────────────────────────────────────────
    if bld and ("터파기" in n or "굴착" in n) and not any(k in n for k in ("운반", "천공", "암", "인력")):
        b = re.search(r"(0\.6|1\.0)M3", t)
        if b:
            v = 560 if b.group(1) == "0.6" else 514
            return f"토사굴착(굴삭기 {b.group(1)}㎥)", PAGE["토공"], "M3", [("굴착", v, {"굴삭기": 1})]
    if bld and "지정" in n and "모래" in t:
        return "기초지정(모래지정)", PAGE["토공"], "M3", [("지정", 65, {"보통인부": 1})]
    if ("띠장" in n or "버팀보" in n or "수평지보" in n) and not any(
            k in n for k in ("제작", "연결", "잭", "JACK", "보걸이", "브라켓", "받침")):
        # 연결재·잭은 가이드라인 설치 작업에 포함된다('가공·연결재·보강재 설치작업 포함') — 따로 잡지 않는다
        ops = _ops(n)
        val = {"설치": 10, "해체": 15}
        return "수평지보공(H-Beam) " + "·".join(ops), PAGE["기초"], "본", [
            (o, val[o], {"철골공": 2, "용접공": 4, "보통인부": 2}) for o in ops]
    if "두부정리" in n:
        if "강관" in t:
            return "강관말뚝 두부정리", PAGE["기초"], "본", [("두부정리", 21, {"용접공": 1, "보통인부": 1})]
        if "콘크리트" in t or "PHC" in t:
            return "콘크리트말뚝 두부정리", PAGE["기초"], "본", [("두부정리", 18, {"할석공": 1, "보통인부": 1})]
        return None
    if bld and "말뚝" in n and ("PHC" in t or "기성" in n) and not any(
            k in n for k in ("이음", "운반", "절단", "보강", "재하", "시험")):
        return "기성말뚝 기초(천공·근입)", PAGE["기초"], "본", [
            ("시공", 15, {"보링공": 1, "기계설비공": 1, "특별인부": 2, "보통인부": 1, "용접공": 1})]
    if bld and "천공" in n and ("H-PILE" in t or "H파일" in t or "말뚝박기" in t):
        return "말뚝박기용 천공(H-Pile)", PAGE["기초"], "본", [
            ("천공", 27, {"보링공": 1, "특별인부": 1, "보통인부": 1, "용접공": 1})]

    # ── 철근콘크리트 ─────────────────────────────────────────────────
    if "철근" in n and ("가공" in n or "조립" in n) and "콘크리트" not in n:
        from crew_space import _rebar_type
        ty = _rebar_type(re.sub(r"TYPE-?([123])", lambda m: "TYPE-" + "ⅠⅡⅢ"[int(m.group(1)) - 1], t))
        steps: List[Step] = []
        if "가공" in n and "공장가공" not in t:
            steps.append(("현장가공", 4.0, {"철근공": 3, "보통인부": 1}))
        if "조립" in n:
            if bld:
                if ty.startswith("Ⅰ"):
                    v = 3.4
                elif "기계적" in t or "커플러" in t:
                    v = 3.5
                elif ty.startswith("Ⅱ"):
                    v = 3.0
                else:
                    return None
                steps.append((f"현장조립(건축 TYPE-{ty.split('-')[0]})", v, {"철근공": 6, "보통인부": 2}))
            elif ty == "Ⅱ-1":
                steps.append(("현장조립(토목 TYPE-Ⅱ-1)", 2.5, {"철근공": 5, "보통인부": 2}))
            else:
                return None
        return ("철근 " + " + ".join(s[0] for s in steps), PAGE["철근콘크리트"], "TON", steps) if steps else None
    if "유로폼" in n and "지수" not in n:
        d = _difficulty(t)
        if d not in ("간단", "보통") or _over(t, 7):
            return None
        val = {"설치": 65 if d == "간단" else 50, "해체": 115}
        ops = _ops(n)
        return f"유로폼 {'·'.join(ops)}({d}, 7m 이하)", PAGE["철근콘크리트"], "M2", [
            (o, val[o], dict(_FORM_WORK)) for o in ops]
    if "합판" in n and "거푸집" in n and not any(k in n for k in ("원형", "문양", "곡면")):
        d = _difficulty(t)
        if d is None:
            return None
        val = {"설치": {"간단": 70, "보통": 60, "복잡": 40}[d], "해체": 100}
        ops = _ops(n)
        return f"합판거푸집 {'·'.join(ops)}({d})", PAGE["철근콘크리트"], "M2", [
            (o, val[o], {"형틀목공": 5, "보통인부": 2}) for o in ops]
    if ("알루미늄폼" in n or "갱폼" in n) and _ops(n) == ("설치", "해체"):
        tbl = {"알루미늄폼": {"셋팅": 30, "마감": 40, "일반": 70}, "갱폼": {"셋팅": 40, "마감": 50, "일반": 90}}
        kind = "알루미늄폼" if "알루미늄폼" in n else "갱폼"
        for k, v in tbl[kind].items():
            if k + "층" in t or (k == "셋팅" and "세팅" in t):
                return f"{kind} 설치 및 해체({k}층)", PAGE["철근콘크리트"], "M2", [
                    ("설치 및 해체", v, dict(_FORM_WORK))]
        return None
    if "타설" in n and "펌프" in t and "콘크리트" in t:
        crew = {"콘크리트공": 4, "특별인부": 2, "보통인부": 1}
        if "무근" in t:
            return "콘크리트 펌프카 타설(무근)", PAGE["철근콘크리트"], "M3", [("타설", 189, crew)]
        if re.search(r"현장조건[-(]?(1|Ⅰ)", t) or re.search(r"현장조건\(F2\)TYPE-?(1|Ⅰ)", t):
            return "콘크리트 펌프카 타설(현장조건 Type-Ⅰ)", PAGE["철근콘크리트"], "M3", [("타설", 182, crew)]
        return "콘크리트 펌프카 타설(철근, Type-Ⅱ)", PAGE["철근콘크리트"], "M3", [("타설", 130, crew)]

    # ── 철골 ────────────────────────────────────────────────────────
    if "데크" in n and "플레이트" in n:
        return "데크플레이트 설치", PAGE["철골"], "M2", [("설치", 85, {"철골공": 3, "용접공": 1, "특별인부": 1})]
    if "철골" in n and ("세우기" in n or "설치" in n) and not any(
            k in n for k in ("도장", "내화", "운반", "가공", "제작", "앵커", "경량", "천장", "천정")):
        if "탑다운" in t:
            return "탑다운공법 철골 지하 현장세우기", PAGE["철골"], "TON", [
                ("세우기", 4.5, {"철골공": 4, "비계공": 2, "특별인부": 1})]
        return "철골 현장세우기(20층 미만)", PAGE["철골"], "TON", [
            ("세우기", 11.3, {"철골공": 5, "비계공": 2, "특별인부": 1})]

    # ── 마감: 조적·미장 ─────────────────────────────────────────────
    if "벽돌" in n and "쌓기" in n and "0.5B" in t and not _over(t, 3.6):
        return "벽돌쌓기(0.5B, 3.6m 이하)", PAGE["조적"], "M2", [("쌓기", 27, {"조적공": 3, "보통인부": 1})]
    if "블록" in n and "쌓기" in n and "보강" in t and not _over(t, 3.6):
        return "블록 보강쌓기(3.6m 이하)", PAGE["조적"], "M2", [("쌓기", 14, {"조적공": 2, "보통인부": 1})]
    if "경량기포" in n and "타설" in n:
        return "경량기포콘크리트 타설", PAGE["미장"], "M3", [
            ("타설", 113, {"일반기계운전사": 1, "미장공": 5, "보통인부": 3})]
    if "모르타르" in n and "타설" in n:
        return "모르타르 타설", PAGE["미장"], "M3", [("타설", 85, {"일반기계운전사": 1, "미장공": 5, "보통인부": 3})]
    if "모르타르" in n and "바름" in n and not any(k in t for k in ("방수", "타일", "바탕고르기")) \
            and not _over(t, 3.6):
        mm = _mm(t)
        if mm is None or mm <= 24:
            return "시멘트 모르타르 바름(24mm 이하, 3.6m 이하)", PAGE["미장"], "M2", [
                ("바름", 42, {"미장공": 3, "보통인부": 1})]
        return None
    if "콘크리트면" in n and ("마무리" in n or "마감" in n) and "부분" not in t and not _over(t, 3.6):
        return "콘크리트면 마무리(전면, 3.6m 이하)", PAGE["미장"], "M2", [
            ("마무리", 17, {"미장공": 3, "견출공": 1, "보통인부": 1})]

    # ── 방수 ────────────────────────────────────────────────────────
    # 가이드라인 방수는 모두 '바닥' 기준 — 벽·수직부·천장은 값이 없다. 보호재·보호층은 방수가 아니다.
    wp = {"방수공": 3, "보통인부": 1}
    floor = not any(k in t for k in ("벽", "천장", "천정", "수직", "보호"))
    if "바탕처리" in n and floor:
        return "방수 바탕처리(바닥)", PAGE["방수"], "M2", [("바탕처리", 100, wp)]
    if "프라이머" in n:
        return "방수 프라이머 바름", PAGE["방수"], "M2", [("바름", 272, wp)]
    if "도막" in n and "방수" in n and floor:
        # 가이드라인은 도막 1층(회) — '2회 바름'이면 같은 조가 두 번 바른다
        k = re.search(r"([2-5])(?:회|층)", t)
        layers = int(k.group(1)) if k else 1
        return f"도막방수(바닥, {layers}층)", PAGE["방수"], "M2", [(f"도막 {i + 1}층", 200, wp)
                                                                for i in range(layers)]
    if "시트" in n and "방수" in n and floor:
        return "시트방수(접착식 1겹, 바닥)", PAGE["방수"], "M2", [("붙임", 88, wp)]
    if "액체방수" in n:
        return "시멘트 액체방수", PAGE["방수"], "M2", [("방수", 40, wp)]

    # ── 목공·창호 ───────────────────────────────────────────────────
    if "데크" in n and ("목재" in t or "합성목" in t):
        if "틀" in n:
            return "목재데크틀 설치", PAGE["목공"], "TON", [("설치", 0.9, {"철공": 8, "용접공": 1, "보통인부": 4})]
        return "목재데크 설치", PAGE["목공"], "M2", [("설치", 17, {"건축목공": 3, "보통인부": 1})]
    if "칸막이" in n and "틀" in n and "경량" not in n:
        return "칸막이벽틀", PAGE["목공"], "M2", [("설치", 20, {"건축목공": 2, "보통인부": 1})]
    if "합판" in n and "거푸집" not in n:
        if "수장" in n:
            return "수장합판", PAGE["목공"], "M2", [("설치", 37, {"건축목공": 2, "보통인부": 1})]
        if "벽" in n:
            return "벽체합판", PAGE["목공"], "M2", [("설치", 40, {"건축목공": 2, "보통인부": 1})]
    if "합성수지" in n and "창" in n:
        return "합성수지 창호 설치", PAGE["창호"], "EA", [("설치", 8, {"창호공": 4, "보통인부": 1})]
    if "유리" in n and "복층" in t and not any(k in n for k in ("운반", "코킹", "필름")):
        return "창호유리 설치(복층)", PAGE["창호"], "M2", [("설치", 42, {"유리공": 5, "보통인부": 1})]

    # ── 타일 ────────────────────────────────────────────────────────
    tile = {"타일공": 3, "보통인부": 1, "줄눈공": 1}
    if "타일" in n and "붙이기" in n:
        if "떠붙이기" in n:
            return "타일 떠붙이기(벽)", PAGE["타일"], "M2", [("붙이기", 21, tile)]
        if "압착" in n:
            return "타일 압착붙이기(바닥)", PAGE["타일"], "M2", [("붙이기", 27, tile)]
        if "접착" in n:
            return "타일 접착붙이기(벽)", PAGE["타일"], "M2", [("붙이기", 39, tile)]
    if "바탕고르기" in n or ("바탕" in n and "모르타르" in n and "타일" in t):
        return "타일 바탕고르기(벽)", PAGE["타일"], "M2", [("바름", 63, {"미장공": 3, "보통인부": 1})]

    # ── 수장·금속 ───────────────────────────────────────────────────
    if "석고" in n and ("판" in n or "보드" in n) and "운반" not in n:
        if any(k in t for k in ("천장", "천정")):
            return "석고판 설치(천장, 1겹)", PAGE["수장"], "M2", [("설치", 45, {"내장공": 2, "보통인부": 1})]
        return "석고판 설치(벽, 2겹)", PAGE["수장"], "M2", [("설치", 45, {"내장공": 2, "보통인부": 1})]
    if any(k in n for k in ("단열재", "압출발포", "비드법", "발포폴리스티렌", "차음재")):
        if "깔기" in n:
            return "단열재(차음재) 슬래브위 깔기", PAGE["수장"], "M2", [("깔기", 260, {"내장공": 2, "보통인부": 1})]
        mm = _mm(t)
        if "붙이기" in n and "벽" in t and (mm is None or mm <= 100):
            return "단열재 설치(벽, 접착 1겹, 100mm 이하)", PAGE["수장"], "M2", [
                ("붙이기", 70, {"내장공": 4, "보통인부": 1})]
        return None
    if "경량" in n and ("벽체틀" in n or "스터드" in t) and "천장" not in n:
        return "경량벽체틀 설치", PAGE["수장"], "M2", [("설치", 65, {"내장공": 2, "보통인부": 1})]
    if "도배" in n or "벽지" in n or "천장지" in n:
        if "천장" in t or "천정" in t:
            return "도배(천장지)", PAGE["수장"], "M2", [("바름", 65, {"도배공": 2, "보통인부": 1})]
        return "도배(벽지)", PAGE["수장"], "M2", [("바름", 85, {"도배공": 2, "보통인부": 1})]
    if "바닥재" in n or "플로어링" in n:
        if "플로어링" in t or "마루" in t:
            return "플로어링 마루 설치", PAGE["수장"], "M2", [("깔기", 50, {"내장공": 2, "보통인부": 1})]
        if "PVC" in t or "비닐" in t:
            if "시트" in t:
                return "PVC계 바닥재(시트형)", PAGE["수장"], "M2", [("깔기", 100, {"내장공": 2, "보통인부": 1})]
            if "타일" in t:
                return "PVC계 바닥재(타일형)", PAGE["수장"], "M2", [("깔기", 40, {"내장공": 2, "보통인부": 1})]

    # ── 조경 ────────────────────────────────────────────────────────
    if bld and ("줄떼" in n or "평떼" in n):
        v = 170 if "줄떼" in n else 150
        return ("잔디 " + ("줄떼" if v == 170 else "평떼")), PAGE["조경"], "M2", [
            ("붙임", v, {"조경공": 1, "보통인부": 4})]
    if unit_key(unit) == "주":
        b = re.search(r"(?:^|[^A-Z]|X)B=?(\d+)", t)          # 'H3.0XW1.5XB10' — 규격 구분자 X 뒤도 본다
        r = re.search(r"(?:^|[^A-Z]|X)R=?(\d+)", t)
        tbl_b = ((8, 9, 17), (10, 17, 12), (18, 24, 9), (25, 34, 7))
        tbl_r = ((9, 11, 17), (12, 20, 12), (21, 29, 9), (30, 41, 7))
        for m, tbl, lbl in ((b, tbl_b, "흉고직경"), (r, tbl_r, "근원직경")):
            if m:
                x = int(m.group(1))
                for lo, hi, v in tbl:
                    if lo <= x <= hi:
                        return f"교목 식재(기계, {lbl} {lo}~{hi})", PAGE["조경"], "주", [
                            ("식재", v, {"조경공": 3, "보통인부": 1})]
                return None
    return None


def match(name: str, spec: str = "", unit: str = "", ctx: str = "토목") -> Optional[Dict[str, Any]]:
    """가이드라인 부록4 1일 작업량(1개 작업조). 조건이 맞지 않거나 단위가 다르면 None.

    ctx: "토목" 또는 "건축"(건축·조경 분야). 철근 현장조립·강관동바리처럼 토목/건축 값이 다른 공종을 가른다.
    반환: {"daily", "unit", "label", "page", "steps": [(작업, 작업량, 작업조)], "crew": 병목 작업조}
    """
    n = _norm(name)
    if not n or any(k in n for k in ("운반", "소운반", "자재대")):
        return None
    t = n + "|" + _norm(spec)
    got = _steps(n, t, unit, ctx)
    if not got:
        return None
    label, page, g_unit, steps = got
    # 해체·철거 항목에 '설치' 값만 있는 공종(석고판·도배 등)을 씌우지 않는다
    if any(k in n for k in ("해체", "철거")) and not any("해체" in op for op, _, _ in steps):
        return None
    if not unit_ok(g_unit, unit):
        return None
    inv = sum(1.0 / v for _, v, _ in steps if v > 0)
    if inv <= 0:
        return None
    crew: Dict[str, int] = {}
    for _, _, c in steps:
        for k, v in c.items():
            crew[k] = max(crew.get(k, 0), v)
    return {"daily": 1.0 / inv, "unit": show_unit(g_unit), "label": label, "page": page, "steps": steps, "crew": crew}


_SHOW = {"M2": "㎡", "M3": "㎥", "공M3": "공㎥", "TON": "ton", "EA": "개소", "M": "m"}


def show_unit(u: str) -> str:
    return _SHOW.get(u, u)


_EQUIP = ("굴삭기", "크레인", "펌프", "롤러", "장비", "트럭")


def unit_labor(m: Dict[str, Any], wages: Dict[str, float], default_wage: float = 0.0) -> Optional[float]:
    """가이드라인 값의 단위 노무비(원/단위) = Σ 단계별(작업조 하루 노임 ÷ 1일 작업량).

    wages는 내역서 노임표({직종: 일 노임}, 공백 없는 직종명). 장비는 노무비가 아니라 뺀다.
    노임을 모르는 직종이 있으면 default_wage, 그것도 없으면 None(비교 불가).
    """
    total = 0.0
    for _, v, crew in m["steps"]:
        day = 0.0
        for t, n in crew.items():
            if any(k in t for k in _EQUIP):
                continue
            w = wages.get(t.replace(" ", "")) or default_wage
            if not w:
                return None
            day += w * n
        total += day / v
    return total if total > 0 else None


def labor_similar(item_ul: Optional[float], guide_ul: Optional[float], tol: float = 1.5) -> Optional[bool]:
    """내역서 단위 노무비와 가이드라인 단위 노무비가 비슷한가(비율 1/tol ~ tol). 비교할 값이 없으면 None."""
    if not item_ul or not guide_ul or item_ul <= 0 or guide_ul <= 0:
        return None
    r = item_ul / guide_ul
    return 1 / tol <= r <= tol


def crew_text(crew: Dict[str, int]) -> str:
    return " + ".join(f"{k} {v}" for k, v in crew.items())


def describe(m: Dict[str, Any]) -> str:
    """'유로폼 설치·해체(보통) = 1 ÷ (1/50 + 1/115) = 34.8㎡/조 (가이드라인 부록4 p109, 형틀목공 4 + 보통인부 1)'"""
    st = m["steps"]
    if len(st) > 1:
        calc = " = 1 ÷ (" + " + ".join(f"1/{v:g}" for _, v, _ in st) + f") = {m['daily']:.2f}"
    else:
        calc = f" {m['daily']:g}"
    return f"{m['label']}{calc}{m['unit']}/조 (가이드라인 부록4 p{m['page']}, {crew_text(m['crew'])})"
