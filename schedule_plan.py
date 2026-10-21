# -*- coding: utf-8 -*-
"""
schedule_plan.py — 예정공정표의 토목 공정 배치(작업일수 축)

발주처 제출용 공사예정표 샘플(풍각·영해 증설)처럼 토목을 고정된 구분(토공·가시설·구조물공·
관로공·우수공·포장공·부대공)으로 묶고, 그 안을 작업 단위 막대(표토제거·터파기·되메우기,
SHEET-PILE 설치·철거, 무근콘크리트·철근·거푸집·방수 …)로 나눈다.

배치는 앱의 공기 모델을 그대로 따른다.
  - 최장(병행): 대공종끼리 병행 → 토목 공기 C = 가장 긴 대공종. 구분은 공사 순서대로
    (가시설 설치·터파기 → 구조물 → 가시설 철거·되메우기·관로·우수 → 포장·부대) 놓되
    모두 [0, C] 안에 들어간다.
  - 합산(순차): 대공종을 공사 순서대로 이어 붙인다 → 합계가 C.
막대 길이는 해당 작업의 작업일수(항목 일수 합)다. 구조물공 안의 작업은 일수 비율대로
구조물공 기간을 나눠 차례로 잇는다(항목이 구간 안에서 순차라는 앱의 모델과 같다).

건축은 구조물공이 끝나면 착수한다(되메우기 전부터 건축을 이어갈 수 있다 — 사용자 기준).
이 모듈의 struct_end를 사업 전체 공기 탭과 예정공정표가 함께 쓴다.

반환 단위는 '본공사 시작부터 센 작업일수'이고, 달력(개월) 환산은 호출부가 한다.
"""

from __future__ import annotations

from typing import Any, Dict, List, Sequence, Tuple

CATS = ["가시설", "토공", "구조물공", "관로공", "우수공", "포장공", "부대공"]


def civil_cat(name: str) -> str:
    """대공종 이름 → 공정표 구분."""
    n = name or ""
    if any(k in n for k in ("가시설", "흙막이")):
        return "가시설"
    if "토공" in n:
        return "토공"
    if "구조물" in n:
        return "구조물공"
    if any(k in n for k in ("관로", "배관", "관부설", "추진")):
        return "관로공"
    if any(k in n for k in ("우수", "배수")):
        return "우수공"
    if "포장" in n:
        return "포장공"
    return "부대공"


def _has(n: str, kws) -> bool:
    return any(k in n for k in kws)


def _pile_base(items) -> str:
    up = " ".join(str(i.get("name", "")).upper() for i in items)
    if "SHEET" in up or "시트파일" in up:
        return "SHEET-PILE"
    if "H-PILE" in up or "H PILE" in up or "H-빔" in up:
        return "H-PILE"
    return "가시설"


# 구조물공 작업 — (이름, 키워드). 위에서부터 공사 순서. 어디에도 안 맞는 항목은 '구조물 공사(거푸집)'.
_STRUCT_ACTS = [
    ("무근콘크리트타설", ("무근", "버림")),
    ("방수공사", ("방수",)),
    ("철근가공 및 조립", ("철근",)),
    ("가설공사(비계 및 동바리)", ("동바리", "비계")),
    ("구조물 공사(거푸집)", ("거푸집", "유로폼", "합판", "폼")),
    ("콘크리트타설", ("타설", "레미콘", "펌프카")),
    ("구체배관", ("구체", "슬리브", "배관")),
]
_STRUCT_ORDER = ["무근콘크리트타설", "가설공사(비계 및 동바리) 설치", "철근가공 및 조립",
                 "구조물 공사(거푸집)", "콘크리트타설", "구체배관",
                 "가설공사(비계 및 동바리) 철거", "방수공사"]


def _struct_acts(items) -> Dict[str, float]:
    acc: Dict[str, float] = {}
    for it in items:
        d = float(it.get("days") or 0)
        if d <= 0:
            continue
        n = str(it.get("name", ""))
        lab = next((a for a, kws in _STRUCT_ACTS if _has(n, kws)), "구조물 공사(거푸집)")
        if lab == "가설공사(비계 및 동바리)":
            acc[lab + " 설치"] = acc.get(lab + " 설치", 0) + d / 2
            acc[lab + " 철거"] = acc.get(lab + " 철거", 0) + d / 2
        else:
            acc[lab] = acc.get(lab, 0) + d
    return acc


def _earth_acts(items) -> Tuple[List[Tuple[str, float]], List[Tuple[str, float]]]:
    """토공 → (구조물 전 작업, 구조물 후 작업). 운반 등 나머지는 터파기에 붙인다."""
    early = {"표토제거": 0.0, "터파기": 0.0}
    late = {"되메우기": 0.0, "성토": 0.0}
    for it in items:
        d = float(it.get("days") or 0)
        if d <= 0:
            continue
        n = str(it.get("name", ""))
        if "표토" in n:
            early["표토제거"] += d
        elif _has(n, ("되메우기", "뒤채움", "다짐")):
            late["되메우기"] += d
        elif _has(n, ("성토", "흙쌓기")):
            late["성토"] += d
        else:
            early["터파기"] += d
    return ([(k, v) for k, v in early.items() if v > 0], [(k, v) for k, v in late.items() if v > 0])


def _pile_acts(items) -> Tuple[float, float]:
    """가시설 → (설치 일수, 철거 일수). '설치 및 철거' 항목은 반씩 나눈다."""
    ins = rem = 0.0
    for it in items:
        d = float(it.get("days") or 0)
        if d <= 0:
            continue
        n = str(it.get("name", ""))
        both = _has(n, ("설치 및 철거", "설치및철거", "설치 및 해체", "설치및해체", "및 철거", "및 해체"))
        if both:
            ins += d / 2
            rem += d / 2
        elif _has(n, ("인발", "철거", "해체", "뽑기")):
            rem += d
        else:
            ins += d
    return ins, rem


def _line_acts(majors) -> List[Tuple[str, float]]:
    """관로·우수·포장·부대: 구간(라인)마다 막대 하나. 구간은 서로 병행이라 같은 시점에 시작한다."""
    out: Dict[str, float] = {}
    for mj in majors:
        lines: Dict[str, float] = {}
        for it in mj["items"]:
            d = float(it.get("days") or 0)
            if d > 0:
                ln = it.get("line") or "(공통)"
                lines[ln] = lines.get(ln, 0) + d
        if any(k != "(공통)" for k in lines):
            for k, v in lines.items():
                lab = mj["name"] if k == "(공통)" else k
                out[lab] = max(out.get(lab, 0), v)
        elif mj["days"] > 0:
            out[mj["name"]] = max(out.get(mj["name"], 0), float(mj["days"]))
    return sorted(out.items(), key=lambda x: -x[1])[:6]


def plan_civil(majors: Sequence[Dict[str, Any]], combine_sum: bool = False) -> Dict[str, Any]:
    """토목 막대 배치.

    majors: [{"name": 대공종명, "days": 작업일수, "items": [{"name", "line", "days"}]}]
            (주공정으로 고른 대공종만, 자재·운반류 제외)
    반환: {"bars": [{"구분", "공종", "s", "e", "days"}], "civil_days": C,
           "struct_end": 구조물공 완료 시점, "struct_span": (시작, 끝) 또는 None}
    """
    majors = [m for m in majors if int(m.get("days") or 0) > 0]
    if not majors:
        return {"bars": [], "civil_days": 0, "struct_end": 0, "struct_span": None}
    by_cat: Dict[str, List[Dict[str, Any]]] = {}
    for m in majors:
        by_cat.setdefault(civil_cat(m["name"]), []).append(m)
    agg = (lambda xs: sum(xs)) if combine_sum else (lambda xs: max(xs, default=0))
    cat_days = {c: agg([float(m["days"]) for m in ms]) for c, ms in by_cat.items()}
    C = sum(cat_days.values()) if combine_sum else max(cat_days.values())
    items_of = {c: [it for m in ms for it in m["items"]] for c, ms in by_cat.items()}

    # 구분별 작업 비율
    pile_in, pile_out = _pile_acts(items_of.get("가시설", []))
    pile_tot = (pile_in + pile_out) or 1.0
    d_pile = cat_days.get("가시설", 0.0)
    d_pile_in, d_pile_out = d_pile * pile_in / pile_tot, d_pile * pile_out / pile_tot
    if d_pile and not (pile_in or pile_out):
        d_pile_in, d_pile_out = d_pile / 2, d_pile / 2
    e_early, e_late = _earth_acts(items_of.get("토공", []))
    e_tot = sum(v for _, v in e_early + e_late) or 1.0
    d_earth = cat_days.get("토공", 0.0)
    early = [(k, d_earth * v / e_tot) for k, v in e_early]
    late = [(k, d_earth * v / e_tot) for k, v in e_late]
    if d_earth and not (early or late):
        early = [("토공", d_earth)]
    d_struct = cat_days.get("구조물공", 0.0)

    bars: List[Dict[str, Any]] = []

    def put(cat, name, s, d, days=None):
        if d <= 0:
            return
        bars.append({"구분": cat, "공종": name, "s": round(s, 3), "e": round(s + d, 3),
                     "days": int(round(days if days is not None else d))})

    def chain(cat, acts, s0):
        t = s0
        for nm, d in acts:
            put(cat, nm, t, d)
            t += d
        return t

    def fit(s, d):
        """[s, s+d)가 토목 기간 [0, C]를 넘지 않게 앞으로 당긴다."""
        return max(0.0, min(s, C - d))

    pile_base = _pile_base(items_of.get("가시설", []))
    if combine_sum:
        # 공사 순서대로 이어 붙인다(합계 = C)
        t = 0.0
        if d_pile_in:
            put("가시설", f"{pile_base} 설치", t, d_pile_in)
            t += d_pile_in
        t = chain("토공", early, t)
        s_struct = t
        t += d_struct
        struct_end = t
        if d_pile_out:
            put("가시설", f"{pile_base} 철거", t, d_pile_out)
            t += d_pile_out
        t = chain("토공", late, t)
        starts = {}
        for c in ("관로공", "우수공", "포장공", "부대공"):
            starts[c] = t
            t += cat_days.get(c, 0.0)
    else:
        pre = max(d_pile_in, sum(d for _, d in early))
        s_struct = fit(pre, d_struct) if d_struct else 0.0
        struct_end = s_struct + d_struct if d_struct else 0.0
        if d_pile_in:
            put("가시설", f"{pile_base} 설치", 0.0, d_pile_in)
        if d_pile_out:
            put("가시설", f"{pile_base} 철거", fit(struct_end, d_pile_out), d_pile_out)
        chain("토공", early, 0.0)
        late_len = sum(d for _, d in late)
        chain("토공", late, fit(struct_end, late_len))
        starts = {}
        for c in ("관로공", "우수공"):
            starts[c] = fit(struct_end, cat_days.get(c, 0.0))
        for c in ("포장공", "부대공"):
            starts[c] = max(0.0, C - cat_days.get(c, 0.0))

    # 구조물공: 작업 비율대로 구조물 기간을 나눠 차례로
    if d_struct:
        acts = _struct_acts(items_of.get("구조물공", []))
        tot = sum(acts.values()) or 1.0
        seq = [(a, d_struct * acts[a] / tot) for a in _STRUCT_ORDER if acts.get(a)]
        chain("구조물공", seq or [("구조물 공사", d_struct)], s_struct)

    # 관로·우수·포장·부대: 구간별 막대(병행 → 같은 시작)
    for c in ("관로공", "우수공", "포장공", "부대공"):
        if c not in by_cat:
            continue
        s0 = starts.get(c, 0.0)
        span = cat_days.get(c, 0.0)
        for nm, d in _line_acts(by_cat[c]):
            put(c, nm, s0, min(d, span) if span else d, d)

    order = {c: i for i, c in enumerate(CATS)}
    bars.sort(key=lambda b: (order.get(b["구분"], 99), b["s"]))
    return {"bars": bars, "civil_days": float(C), "struct_end": float(struct_end if d_struct else C),
            "struct_span": (s_struct, struct_end) if d_struct else None}


def project_days(civil_days: float, struct_end: float, arch: float = 0, land: float = 0) -> int:
    """사업 전체 순작업일수 — 건축은 구조물공 완료 후, 조경은 토목 완료 후 착수."""
    return int(round(max(civil_days + land, (struct_end + arch) if arch else 0, civil_days)))
