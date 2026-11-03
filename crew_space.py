# -*- coding: utf-8 -*-
"""
crew_space.py — 면적 기준 동시 투입 상한(터파기·철근·거푸집)

개략 검토에서는 '항목당 최대 투입조수' 하나로 충분하지만, 상세 검토에서는 구조물 면적에서
동시에 붙일 수 있는 조수의 상한을 근거 있게 구한다(사용자가 켰을 때만).

    동시 투입 상한(조) = 작업면적 ÷ 1조 소요면적
    1조 소요면적       = 1조 인원 × 1인 작업공간

근거
  - 1인 작업공간: Thomas & Smith (1990), "Loss of Construction Labor Productivity Due to
    Inefficiencies and Disruptions: The Weight of Expert Opinion", PTI Report 9019,
    Pennsylvania State Univ. — 1인당 최소 19㎡, 최적 생산성 30.2㎡ (Francis 2020,
    Front. Built Environ. 6:67에서 재인용). 국내 법령·연구에는 1인당 면적 기준이 없다.
  - 1조 인원: 2026 표준품셈 작업조 — 철근 현장조립(6-2-3) 토목 Ⅰ-1 8인(철근공 6+보통인부 2),
    Ⅱ-1 7인(5+2), 유로폼(6-3-3) 5인(형틀목공 4+보통인부 1), 합판거푸집(6-3-1) 7인(5+2).
    앱의 1일 작업량이 노무비 역산(직종별 1인 = 1조)에서 왔으면 1조 인원은 직종 수다.
  - 굴착기: 국내 규칙은 접촉 위험 장소 출입 금지(산업안전보건기준에 관한 규칙 제200조)만 두고
    숫자 기준이 없다 → 작업반경 R로 1대 영역을 (2R)²로 본다(반경이 겹치지 않게).
    R은 내역서(산근·일위대가)의 굴착기 버킷 용량에서 제원표의 최대 굴착반경으로 정한다(REACH_ANCHORS).
  - 1인 작업공간 기본값은 최적 생산성 30.2㎡(사용자 결정 — 최소 19㎡가 아니라 최적 기준).

작업면적
  - 철근: 바닥슬라브 + 벽체(한 면) — 벽은 비계 위에서 면을 따라 배근한다고 보고 면적을 쓴다
  - 거푸집: 벽체 양면 + 바닥슬라브(하부)
  - 터파기: 구조물 바닥슬라브 합(굴착 평면), 부지면적을 넘지 않게
"""

from __future__ import annotations

import math
import re
from typing import Any, Dict, Optional

PERSON_AREA_MIN = 19.0      # ㎡/인 — Thomas & Smith(1990) 최소
PERSON_AREA_OPT = 30.2      # ㎡/인 — 최적 생산성(기본값)

# 버킷 용량(㎥) → 최대 굴착반경(m). HD건설기계 제원(장비왕 포크레인 제원표):
#   030~035급 0.11~0.13㎥ 4.9~5.3m, 02급 DX65-7K 0.17㎥ 6.305m, 03급 DX85R-7 0.28㎥ 7.15m,
#   06급 DX150LC-7 0.59㎥ 8.5m, 08급 DX240 0.92㎥ 9.945m (HX220A 0.92㎥ 9.98m).
# 사이 규격(0.4·0.7㎥)은 선형 보간, 범위 밖은 끝값.
REACH_ANCHORS = [(0.12, 5.0), (0.17, 6.3), (0.28, 7.15), (0.59, 8.5), (0.92, 9.95)]

# 품셈 작업조 인원
PUMSEM_CREW = {
    "철근": {"Ⅰ": 8, "Ⅱ": 7, "Ⅲ": 7},
    "유로폼": 5,
    "합판": 7,
}


def excavator_reach(bucket_m3: Optional[float]) -> Optional[float]:
    """버킷 용량(㎥) → 최대 굴착반경(m). 모르면 None."""
    if not bucket_m3 or bucket_m3 <= 0:
        return None
    pts = REACH_ANCHORS
    if bucket_m3 <= pts[0][0]:
        return pts[0][1]
    for (q1, r1), (q2, r2) in zip(pts, pts[1:]):
        if bucket_m3 <= q2:
            return round(r1 + (r2 - r1) * (bucket_m3 - q1) / (q2 - q1), 2)
    return pts[-1][1]


def trade_class(name: str) -> Optional[str]:
    """세부공종명 → '터파기' / '철근' / '거푸집' / None."""
    n = name or ""
    if any(k in n for k in ("운반", "상차", "깨기", "철거")):
        return None
    if "터파기" in n or ("굴착" in n and "천공" not in n):
        return "터파기"
    if "철근" in n and any(k in n for k in ("조립", "가공", "배근")):
        return "철근"
    if any(k in n for k in ("거푸집", "유로폼", "합판")):
        return "거푸집"
    return None


def pumsem_persons(cls: str, name: str = "", spec: str = "") -> int:
    """품셈 작업조 인원(1조)."""
    t = f"{name} {spec}"
    if cls == "철근":
        m = re.search(r"TYPE\s*-?\s*(Ⅲ|Ⅱ|Ⅰ)", t.upper().replace("III", "Ⅲ").replace("II", "Ⅱ"))
        return PUMSEM_CREW["철근"].get(m.group(1) if m else "Ⅱ", 7)
    if cls == "거푸집":
        return PUMSEM_CREW["유로폼"] if "유로폼" in t else PUMSEM_CREW["합판"]
    return 1


def crew_persons(cls: str, name: str, spec: str, method: str, persons: Optional[int] = None) -> int:
    """앱에서 이 항목의 '1조'가 몇 명인지.

    1일 작업량이 노무비 역산(직종별 인수 ÷ 품셈 구성원)에서 왔으면 그 구성원 인원(persons),
    그 밖(산근 A·Q, 품셈 작업조 등)은 품셈 작업조 인원.
    """
    if cls == "터파기":
        return 1
    if str(method).startswith("노무비역산") and persons:
        return max(1, int(persons))
    return pumsem_persons(cls, name, spec)


def work_area(cls: str, slab: float, wall: float, excav: float = 0.0) -> float:
    """공종별 작업면적(㎡). wall은 벽체 한 면 면적."""
    if cls == "철근":
        return max(0.0, slab) + max(0.0, wall)
    if cls == "거푸집":
        return 2 * max(0.0, wall) + max(0.0, slab)
    if cls == "터파기":
        return max(0.0, excav)
    return 0.0


def area_cap(cls: str, area: float, persons: int, person_area: float = PERSON_AREA_OPT,
             excavator_radius: float = 6.0) -> Optional[int]:
    """동시 투입 상한(조). 면적이 없으면 None(상한 미적용)."""
    if area <= 0:
        return None
    if cls == "터파기":
        unit = (2 * excavator_radius) ** 2
    else:
        unit = max(1, persons) * person_area
    if unit <= 0:
        return None
    return max(1, int(math.floor(area / unit)))


def explain(cls: str, area: float, persons: int, person_area: float, excavator_radius: float,
            cap: Optional[int]) -> str:
    """상한 근거 한 줄."""
    if cap is None:
        return "면적 미입력 — 상한 미적용"
    if cls == "터파기":
        return (f"굴착 평면 {area:,.0f}㎡ ÷ 굴착기 1대 영역 (2×{excavator_radius:g}m)² "
                f"= {area / (2 * excavator_radius) ** 2:.1f} → {cap}대")
    return (f"작업면적 {area:,.0f}㎡ ÷ (1조 {persons}인 × {person_area:g}㎡) "
            f"= {area / (persons * person_area):.1f} → {cap}조")


# ── 내역서로 어림한 구조물 면적(개략 검토 기본값) ─────────────────────────────
BLINDING_T = 0.1    # 버림콘크리트 두께(m) — 풍각 일반도 단면 T=100, 일반적인 값


def boq_quantities(items) -> Dict[str, float]:
    """구조물 한 곳의 내역서 항목 → {'버림': ㎥, '철근콘크리트': ㎥, '거푸집': ㎡}.

    버림은 이름에 '버림'·'무근'이 든 콘크리트(깨기 제외), 철근콘크리트는 이름이나 규격에 '철근'이 든
    타설 항목, 거푸집은 거푸집·유로폼·합판 면적 합.
    """
    q = {"버림": 0.0, "철근콘크리트": 0.0, "거푸집": 0.0}
    for it in items:
        n, sp_ = str(it.get("name", "")), str(it.get("spec", ""))
        u = str(it.get("unit", "")).replace("M3", "㎥").replace("m3", "㎥").replace("M2", "㎡").replace("m2", "㎡")
        v = float(it.get("qty") or 0)
        if v <= 0 or any(k in n for k in ("깨기", "철거", "운반")):
            continue
        if u == "㎥" and any(k in n for k in ("버림", "무근")):
            q["버림"] += v
        elif u == "㎥" and ("타설" in n or "레미콘" in n) and ("철근" in n or "철근" in sp_):
            q["철근콘크리트"] += v
        elif u == "㎡" and any(k in n for k in ("거푸집", "유로폼", "합판")):
            q["거푸집"] += v
    return q


def estimate_areas(items, blinding_t: float = BLINDING_T):
    """내역서로 어림한 (바닥슬라브 ㎡, 벽체 한 면 ㎡, 근거). 어림할 수 없으면 0.

    바닥 ≈ 버림 ㎥ ÷ 두께(풍각: 45.4㎥ → 454㎡, 일반도 459㎡)
    벽(한 면) ≈ (거푸집 − 바닥) ÷ 2 — 거푸집은 벽 양면 + 슬라브 하부
    """
    q = boq_quantities(items)
    slab = q["버림"] / blinding_t if q["버림"] > 0 else 0.0
    wall = max(0.0, (q["거푸집"] - slab) / 2) if q["거푸집"] > 0 else 0.0
    parts = []
    if slab:
        parts.append(f"바닥 = 버림 {q['버림']:,.1f}㎥ ÷ {blinding_t:g}m")
    if wall:
        parts.append(f"벽 = (거푸집 {q['거푸집']:,.0f}㎡ − 바닥 {slab:,.0f}㎡) ÷ 2")
    return round(slab, 1), round(wall, 1), " · ".join(parts) or "내역서로 어림 불가(버림·거푸집 항목 없음)"


# ── 품셈 작업조(노무비만 있는 블록을 1조 기준으로 바꿀 때) ─────────────────────
# 표준시장단가처럼 직종별 인수 없이 ㎡당 노무비만 있는 블록은 그동안 '노무비 ÷ 보통인부 노임'(1인 기준)으로
# 계산돼 품셈 1조보다 수 배 느리게 나왔다(공주정수장 유로폼 4㎡/일 vs 품셈 25~40㎡/조). 공종이 품셈 작업조와
# 맞으면 작업조의 하루 노임 합으로 나눠 1조 기준으로 바꾼다.
#   단계가 하나:  1일 작업량 = 작업조 하루 노임 합 ÷ 단위당 노무비
#   단계가 여럿(설치 → 해체 등, 순서대로): 품셈 단위당 노무비 P = Σ(작업조 노임_i ÷ 품셈 일당시공량_i),
#     실제 단위당 노무비 L로 보정 k = L ÷ P, 1일 작업량 = 1 ÷ (k × Σ 1/일당시공량_i)
# 2026 표준품셈 원문 대조(쪽): 유로폼 6-3-3(p195), 합판 6-3-1(p193), 철근 6-2-2·6-2-3(p189~190),
# 강관 동바리 2-6-1(p101), 시스템 동바리 2-6-3(p102), 강관 비계 2-7-1·시스템 비계 2-7-2(p103),
# 콘크리트 펌프차 타설 6-1-4(p186).
_REBAR_ASSY = {"Ⅰ-1": ({"철근공": 6, "보통인부": 2}, 3.4), "Ⅰ-2": ({"철근공": 4, "보통인부": 1}, 2.2),
               "Ⅱ-1": ({"철근공": 5, "보통인부": 2}, 2.6), "Ⅱ-2": ({"철근공": 2, "보통인부": 1}, 1.1),
               "Ⅲ": ({"철근공": 5, "보통인부": 2}, 2.4)}
_REBAR_FAB = {"Ⅰ": 4.5, "Ⅱ": 4.0, "Ⅲ": 3.5}


def _rebar_type(t: str) -> str:
    u = t.upper().replace("III", "Ⅲ").replace("II", "Ⅱ").replace("TYPE-I", "TYPE-Ⅰ")
    m = re.search(r"TYPE\s*-?\s*(Ⅲ|Ⅱ|Ⅰ)(?:\s*-\s*([12]))?", u)
    if not m:
        return "Ⅱ-1"
    return m.group(1) if m.group(1) == "Ⅲ" else f"{m.group(1)}-{m.group(2) or '1'}"


# 노무비만 있는 항목의 구성원을 정할 때 쓰는 품셈 단위당 인수(2026 표준품셈) 또는 가이드라인 작업조.
# (키워드 묶음 — 묶음마다 하나 이상 포함, 직종별 인수, 출처)
_PUMSEM_TRADES = [
    ((("수성페인트", "유성페인트", "페인트"), ("롤러", "칠")), {"도장공": 0.012, "보통인부": 0.002}, "품셈 11-2-2"),
    ((("에폭시",), ("라이닝",)), {"도장공": 0.044, "보통인부": 0.023}, "품셈 11-2-8"),
    ((("에폭시",), ("코팅", "페인트")), {"도장공": 0.039, "보통인부": 0.008}, "품셈 11-2-8"),
    ((("현장정리", "준공청소"),), {"보통인부": 0.13}, "품셈 2-11-2"),
    ((("방수",),), {"방수공": 0.03, "보통인부": 0.01}, "가이드라인 부록4 방수공사(방수공 3 + 보통인부 1)"),
    ((("콘크리트면",), ("마감", "마무리")), {"미장공": 0.03, "견출공": 0.01, "보통인부": 0.01},
     "가이드라인 부록4 콘크리트면 마무리(미장공 3 + 견출공 1 + 보통인부 1)"),
    ((("모르타르", "미장"), ("바름",)), {"미장공": 0.03, "보통인부": 0.01},
     "가이드라인 부록4 시멘트 모르타르 바름(미장공 3 + 보통인부 1)"),
]


def crew_steps(title: str, spec: str = "", fallback: bool = True):
    """공종 → [(작업조 {직종: 인원}, 품셈 일당시공량)]. 품셈 작업조를 모르면 None.

    fallback=True면 품셈 작업조 표에 없는 공종도 품셈 단위당 인수(_PUMSEM_TRADES)·가이드라인 작업조로
    구성원을 정한다 — 직종별 인수가 없는(노무비만 있는) 항목용. 인수가 있는 항목은 그 인수가 곧 품셈
    구성이라 fallback을 쓰지 않는다(pumsem_crew)."""
    t = f"{title} {spec}"
    n = re.sub(r"\s+", "", str(title or ""))
    if any(k in n for k in ("깨기", "운반", "절단")):
        return None
    both = ("해체" in n or "철거" in n) or not ("설치" in n)
    if "동바리" in n:
        if "시스템" in n:
            steps = [({"형틀목공": 4, "보통인부": 1}, 130)]
            return steps + [({"형틀목공": 2, "보통인부": 2}, 150)] if both else steps
        if "강관" in n:
            return [({"형틀목공": 3, "보통인부": 1}, 45)]
    if "비계" in n and "브라켓" not in n:
        if "시스템" in n:
            steps = [({"비계공": 4, "보통인부": 1}, 150)]
            return steps + [({"비계공": 2, "보통인부": 2}, 150)] if both else steps
        return [({"비계공": 3, "보통인부": 1}, 55)]
    if "유로폼" in n and "지수" not in n:
        return [({"형틀목공": 4, "보통인부": 1}, 35)]
    if "거푸집" in n and not any(k in n for k in ("원형", "강재", "문양", "유로폼")):
        return [({"형틀목공": 5, "보통인부": 2}, 30)]
    if "철근" in n and ("조립" in n or "가공" in n) and "콘크리트" not in n:
        ty = _rebar_type(t)
        assy = _REBAR_ASSY.get(ty, _REBAR_ASSY["Ⅱ-1"])
        fab = ({"철근공": 3, "보통인부": 1}, _REBAR_FAB.get(ty.split("-")[0], 4.0))
        if "공장가공" in t.replace(" ", ""):
            return [assy]                         # 가공은 공장 — 현장은 조립만
        if "조립" in n and "가공" in n:
            return [fab, assy]
        return [assy] if "조립" in n else [fab]
    # 천장재 붙임 — 품셈 5-2-1 흡음텍스 설치·5-2-2 열경화성수지천장판 설치(내장공 2 + 보통인부 1, 45㎡).
    # 압출성형 시멘트판(천장 텍스류)은 품셈에 없어 같은 천장재 구성원으로 본다(사용자 확인 2026-11-03).
    if any(k in n for k in ("텍스", "천장판", "천정판", "압출성형", "시멘트판")) and not any(
            k in n for k in ("해체", "철거", "운반")):
        return [({"내장공": 2, "보통인부": 1}, 45)]
    if "타설" in n and ("펌프" in n or "펌프" in t) and "콘크리트" in t:
        rc = "철근" in t
        return [({"콘크리트공": 4 if rc else 3, "특별인부": 1, "형틀목공": 1, "보통인부": 1}, 100)]
    if not fallback:
        return None
    # 품셈 단위당 인수로만 나오는 공종 — 구성원은 인수 비율 그대로(가장 적게 드는 직종 1인)
    for keys, trades, _src in _PUMSEM_TRADES:
        if all(any(k in n for k in alt) for alt in keys) and "해체" not in n and "보호" not in n:
            crew = ratio_crew(trades)
            return [(crew, crew_rate(trades, crew)[0])]
    # 가이드라인 부록4의 작업조(품셈 생산성 기준) — 1일 작업량 값은 노무비가 맞을 때만 쓰지만(guideline_work),
    # 구성원은 같은 공종이면 그대로 쓴다. 예) 우레탄도막방수 3mm: 노무비가 1층 값의 4배라 200㎡/일은 안 쓰고,
    # 구성원 방수공 3 + 보통인부 1로 노무비 ÷ 작업조 노임 → 그 구성원의 1일 작업량
    try:
        import guideline_work as gw
        m = gw.match(title, spec, "", "토목")
    except Exception:            # noqa: BLE001
        m = None
    if m:
        steps = [({k: n for k, n in c.items() if not any(e in k for e in gw._EQUIP)}, v) for _, v, c in m["steps"]]
        if all(c for c, _ in steps):          # 장비만 있는 작업조(토사굴착 등)는 구성원이 아니다
            return steps
    return None


def crew_text(steps) -> str:
    return " → ".join(" + ".join(f"{k} {v}" for k, v in c.items()) for c, _ in steps)


# ── 1조 = 표준품셈 구성원(사용자 기준 2026-11-03) ───────────────────────────
# 공기 산정의 조수는 표준품셈의 구성원(직종·인원) 단위로 투입하고, 1일 작업량은 그 구성원이 하루에
# 하는 양이다. 일위대가의 직종별 인수(인/단위)는 품셈 구성원이 단위 물량에 드는 품이므로
#   1조 1일 작업량 = min(구성원 인원_직종 ÷ 인수_직종)
# 구성원은 품셈 작업조를 아는 공종(crew_steps)이면 그 작업조, 모르면 인수 비율 그대로(가장 적게 드는
# 주요 직종 = 1인). 예) 방수공 0.03 · 보통인부 0.01 → 방수공 3 + 보통인부 1 → 100㎡/조·일.
# 주 직종의 1/10도 안 드는 직종은 1인(그 직종이 병목이 되지 않는다).
# 직종별 1인(병목 직종)으로 보던 예전 방식은 품셈 구성원이 아니어서 작업일수가 구성원 수만큼 길었다.

def ratio_crew(trades: Dict[str, float]) -> Dict[str, int]:
    """직종별 인수 → 품셈 구성원(인수 비율, 가장 적게 드는 주요 직종 1인)."""
    tr = {t: float(v) for t, v in (trades or {}).items() if v and v > 0}
    if not tr:
        return {}
    lv = max(tr.values())
    sig = {t: v for t, v in tr.items() if v >= lv * 0.1}
    base = min(sig.values())
    return {t: (max(1, math.ceil(v / base - 1e-6)) if t in sig else 1) for t, v in tr.items()}


def pumsem_crew(title: str, spec: str = "", trades: Optional[Dict[str, float]] = None):
    """항목의 1조 구성원과 출처. (구성원 {직종: 인원}, '품셈 작업조' | '품셈 인수 구성' | '')."""
    trades = {t: float(v) for t, v in (trades or {}).items() if v and v > 0}
    # 인수가 있으면 그 인수가 품셈 구성 — 품셈 작업조 표에 있는 공종만 작업조로 바꾸고 나머지는 인수 비율
    steps = crew_steps(title, spec, fallback=not trades)
    if steps:
        crew: Dict[str, int] = {}
        for c, _ in steps:
            for k, v in c.items():
                crew[k] = max(crew.get(k, 0), v)
        if trades:
            # 품셈 작업조에 없는 직종(내역서에 더 붙은 직종)은 주 직종과의 인수 비율로 붙인다
            lead_t = max(crew, key=crew.get)
            ref = trades.get(lead_t) or max(trades.values())
            for t, v in trades.items():
                if t not in crew:
                    crew[t] = max(1, math.ceil(v / ref * crew[lead_t] - 1e-6))
        return crew, "품셈 작업조"
    if trades:
        return ratio_crew(trades), "품셈 인수 구성"
    return {}, ""


def crew_rate(trades: Dict[str, float], crew: Dict[str, int]):
    """1조(구성원)가 하루에 하는 양(인수의 기준수량 단위)과 병목 직종. (작업량, 직종)."""
    best = None
    for t, v in (trades or {}).items():
        if v and v > 0:
            r = max(1, int(crew.get(t, 1))) / float(v)
            if best is None or r < best[0]:
                best = (r, t)
    return best if best else (0.0, "")


def crew_label(crew: Dict[str, int]) -> str:
    return " + ".join(f"{k} {v}" for k, v in sorted(crew.items(), key=lambda x: -x[1]))
