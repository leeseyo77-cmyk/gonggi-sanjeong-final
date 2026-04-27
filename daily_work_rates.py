# daily_work_rates.py
# 출처: 적정 공사기간 확보 가이드라인 부록1,2 (국토교통부, 2025.01.)
# 구조: {공종키: {"daily": 1일작업량, "unit": 단위, "crews": 기본조수, "hours": 작업시간, "condition": 작업조건}}

# ══════════════════════════════════════════════════════════════
# 하수관로 주요공종 1일 작업량
# ══════════════════════════════════════════════════════════════
DAILY_WORK = {

    # ── 1. 포장 깨기/절단 ────────────────────────────────────
    "아스팔트포장절단":    {"daily":600,  "unit":"m",  "crews":5, "hours":8, "condition":"아스팔트포장"},
    "아스팔트포장깨기":    {"daily":54,   "unit":"㎥", "crews":5, "hours":8, "condition":"B.H0.7㎥+대형브레이카"},
    "콘크리트포장절단":    {"daily":600,  "unit":"m",  "crews":5, "hours":8, "condition":"콘크리트포장"},
    "콘크리트포장깨기":    {"daily":54,   "unit":"㎥", "crews":5, "hours":8, "condition":"B.H0.7㎥+대형브레이카"},

    # ── 2. 터파기 ────────────────────────────────────────────
    "터파기_토사_육상":    {"daily":560,  "unit":"㎥", "crews":1, "hours":8, "condition":"B.H 0.7㎥ 기준"},
    "터파기_토사_용수":    {"daily":420,  "unit":"㎥", "crews":1, "hours":8, "condition":"용수발생 25%감"},
    "터파기_풍화암":       {"daily":280,  "unit":"㎥", "crews":1, "hours":8, "condition":"브레이커 병용"},
    "터파기_연암":         {"daily":240,  "unit":"㎥", "crews":1, "hours":8, "condition":"브레이커 병용"},
    "터파기_보통암":       {"daily":176,  "unit":"㎥", "crews":1, "hours":8, "condition":"발파"},
    "터파기_경암":         {"daily":128,  "unit":"㎥", "crews":1, "hours":8, "condition":"발파"},

    # ── 3. 되메우기·모래기초 ─────────────────────────────────
    "되메우기_관상단_토사":{"daily":79,   "unit":"㎥", "crews":5, "hours":8, "condition":"B/H+진동롤러"},
    "되메우기_관주위_토사":{"daily":42,   "unit":"㎥", "crews":5, "hours":8, "condition":"B/H+램머"},
    "모래부설다짐":        {"daily":79,   "unit":"㎥", "crews":5, "hours":8, "condition":"물다짐"},

    # ── 4. 관 부설 (PE다중벽관/고강성PVC) ────────────────────
    "PE관_D150":  {"daily":7,  "unit":"개소", "crews":3, "hours":8, "condition":"D150mm×6.0m"},
    "PE관_D200":  {"daily":5,  "unit":"개소", "crews":3, "hours":8, "condition":"D200mm×6.0m"},
    "PE관_D250":  {"daily":4,  "unit":"개소", "crews":3, "hours":8, "condition":"D250mm×6.0m"},
    "PE관_D300":  {"daily":4,  "unit":"개소", "crews":3, "hours":8, "condition":"D300mm×6.0m"},
    "PE관_D350":  {"daily":3,  "unit":"개소", "crews":3, "hours":8, "condition":"D350mm×6.0m"},
    "PE관_D400":  {"daily":3,  "unit":"개소", "crews":3, "hours":8, "condition":"D400mm×6.0m"},
    "PE관_D450":  {"daily":2,  "unit":"개소", "crews":3, "hours":8, "condition":"D450mm×6.0m"},
    "PE관_D500":  {"daily":2,  "unit":"개소", "crews":3, "hours":8, "condition":"D500mm×6.0m"},

    # ── 5. 흄관/RC관 부설 ────────────────────────────────────
    "흄관_D250":  {"daily":43, "unit":"m",   "crews":3, "hours":8, "condition":"소켓접합 6m기준"},
    "흄관_D300":  {"daily":43, "unit":"m",   "crews":3, "hours":8, "condition":"소켓접합 6m기준"},
    "흄관_D350":  {"daily":43, "unit":"m",   "crews":3, "hours":8, "condition":"소켓접합 6m기준"},
    "흄관_D400":  {"daily":43, "unit":"m",   "crews":3, "hours":8, "condition":"소켓접합 6m기준"},
    "흄관_D450":  {"daily":43, "unit":"m",   "crews":3, "hours":8, "condition":"소켓접합 6m기준"},
    "흄관_D500":  {"daily":43, "unit":"m",   "crews":3, "hours":8, "condition":"소켓접합 6m기준"},
    "흄관_D600":  {"daily":32, "unit":"m",   "crews":3, "hours":8, "condition":"소켓접합 6m기준"},
    "흄관_D700":  {"daily":26, "unit":"m",   "crews":3, "hours":8, "condition":"소켓접합 6m기준"},
    "흄관_D800":  {"daily":20, "unit":"m",   "crews":3, "hours":8, "condition":"소켓접합 6m기준"},
    "흄관_D900":  {"daily":16, "unit":"m",   "crews":3, "hours":8, "condition":"소켓접합 6m기준"},
    "흄관_D1000": {"daily":14, "unit":"m",   "crews":3, "hours":8, "condition":"소켓접합 6m기준"},
    "흄관_D1200": {"daily":10, "unit":"m",   "crews":3, "hours":8, "condition":"소켓접합 6m기준"},

    # ── 6. 맨홀 설치 ─────────────────────────────────────────
    "소형맨홀_D600":  {"daily":8.16, "unit":"개소", "crews":5, "hours":8, "condition":"하부구체+상판"},
    "PC맨홀1호_D900": {"daily":8.16, "unit":"개소", "crews":3, "hours":8, "condition":"하부구체+상판"},
    "PC맨홀2호_D1200":{"daily":5.0,  "unit":"개소", "crews":3, "hours":8, "condition":"하부구체+상판"},
    "PC맨홀3호_D1500":{"daily":4.0,  "unit":"개소", "crews":3, "hours":8, "condition":"하부구체+상판"},
    "맨홀뚜껑설치":   {"daily":1.43, "unit":"조",   "crews":3, "hours":8, "condition":"주철재 φ648"},

    # ── 7. 배수설비 ──────────────────────────────────────────
    "오수받이설치":   {"daily":4,   "unit":"개소", "crews":3, "hours":8, "condition":"소형구조물"},
    "배수설비":       {"daily":4,   "unit":"개소", "crews":3, "hours":8, "condition":"연결관 포함"},

    # ── 8. 포장복구 ──────────────────────────────────────────
    "보조기층포설":   {"daily":800, "unit":"㎡",  "crews":5, "hours":8, "condition":"기계포설"},
    "아스콘포장":     {"daily":600, "unit":"㎡",  "crews":5, "hours":8, "condition":"기계시공 본선"},
    "콘크리트포장":   {"daily":400, "unit":"㎡",  "crews":5, "hours":8, "condition":"기계시공"},

    # ── 9. 가시설 ────────────────────────────────────────────
    "가시설흙막이":   {"daily":28,  "unit":"m",  "crews":5, "hours":8, "condition":"조립식 간이흙막이 L=28m"},
    "강관말뚝박기":   {"daily":15,  "unit":"본",  "crews":3, "hours":8, "condition":"H파일 20m"},

    # ── 10. 추진공 ───────────────────────────────────────────
    "강관압입추진_D450": {"daily":8, "unit":"m", "crews":1, "hours":8, "condition":"D450mm 토사"},
    "강관압입추진_D600": {"daily":6, "unit":"m", "crews":1, "hours":8, "condition":"D600mm 토사"},
    "강관압입추진_D800": {"daily":5, "unit":"m", "crews":1, "hours":8, "condition":"D800mm 토사"},
}

# ── 공종명 키워드 → DAILY_WORK 키 매핑 ───────────────────────
WORK_KEY_MAP = [
    # (공종명 키워드 리스트, 규격 키워드 리스트, DAILY_WORK 키)
    (["아스팔트포장절단","포장절단"],          [],           "아스팔트포장절단"),
    (["아스팔트포장깨기","포장깨기"],          [],           "아스팔트포장깨기"),
    (["콘크리트포장절단"],                    [],           "콘크리트포장절단"),
    (["콘크리트포장깨기"],                    [],           "콘크리트포장깨기"),

    # 터파기 - 토질별
    (["터파기","굴착"],  ["경암"],             "터파기_경암"),
    (["터파기","굴착"],  ["보통암"],           "터파기_보통암"),
    (["터파기","굴착"],  ["연암"],             "터파기_연암"),
    (["터파기","굴착"],  ["풍화암"],           "터파기_풍화암"),
    (["터파기","굴착"],  ["용수"],             "터파기_토사_용수"),
    (["터파기","굴착"],  [],                   "터파기_토사_육상"),

    # 되메우기
    (["되메우기"],       ["관주위","관 주위"], "되메우기_관주위_토사"),
    (["모래부설","모래기초","모래,관기초"],[], "모래부설다짐"),
    (["되메우기"],       [],                   "되메우기_관상단_토사"),

    # PE관/고강성PVC/이중벽관 - 관경별
    (["PE다중벽","이중벽","고강성PVC","PE관"],["D150","150mm","Φ150"], "PE관_D150"),
    (["PE다중벽","이중벽","고강성PVC","PE관"],["D200","200mm","Φ200"], "PE관_D200"),
    (["PE다중벽","이중벽","고강성PVC","PE관"],["D250","250mm","Φ250"], "PE관_D250"),
    (["PE다중벽","이중벽","고강성PVC","PE관"],["D300","300mm","Φ300"], "PE관_D300"),
    (["PE다중벽","이중벽","고강성PVC","PE관"],["D350","350mm","Φ350"], "PE관_D350"),
    (["PE다중벽","이중벽","고강성PVC","PE관"],["D400","400mm","Φ400"], "PE관_D400"),
    (["PE다중벽","이중벽","고강성PVC","PE관"],["D450","450mm","Φ450"], "PE관_D450"),
    (["PE다중벽","이중벽","고강성PVC","PE관"],["D500","500mm","Φ500"], "PE관_D500"),
    (["PE다중벽","이중벽","고강성PVC","PE관"],[],                      "PE관_D200"),  # 기본값

    # 흄관/RC관 - 관경별
    (["흄관","RC관","원심력"],["D1200","1200mm"], "흄관_D1200"),
    (["흄관","RC관","원심력"],["D1000","1000mm"], "흄관_D1000"),
    (["흄관","RC관","원심력"],["D900","900mm"],   "흄관_D900"),
    (["흄관","RC관","원심력"],["D800","800mm"],   "흄관_D800"),
    (["흄관","RC관","원심력"],["D700","700mm"],   "흄관_D700"),
    (["흄관","RC관","원심력"],["D600","600mm"],   "흄관_D600"),
    (["흄관","RC관","원심력"],["D500","500mm"],   "흄관_D500"),
    (["흄관","RC관","원심력"],["D450","450mm"],   "흄관_D450"),
    (["흄관","RC관","원심력"],["D400","400mm"],   "흄관_D400"),
    (["흄관","RC관","원심력"],["D350","350mm"],   "흄관_D350"),
    (["흄관","RC관","원심력"],["D300","300mm"],   "흄관_D300"),
    (["흄관","RC관","원심력"],["D250","250mm"],   "흄관_D250"),
    (["흄관","RC관","원심력"],[],                 "흄관_D300"),  # 기본값

    # 맨홀
    (["소형맨홀","D600맨홀"], [],                     "소형맨홀_D600"),
    (["맨홀뚜껑"],            [],                     "맨홀뚜껑설치"),
    (["맨홀"],                ["D1500","1500mm","3호"], "PC맨홀3호_D1500"),
    (["맨홀"],                ["D1200","1200mm","2호"], "PC맨홀2호_D1200"),
    (["맨홀"],                [],                     "PC맨홀1호_D900"),

    # 배수설비
    (["오수받이"],    [], "오수받이설치"),
    (["배수설비"],    [], "배수설비"),

    # 포장복구
    (["보조기층"],    [], "보조기층포설"),
    (["아스콘포장","아스팔트포장","아스팔트표층","아스팔트기층"], [], "아스콘포장"),
    (["콘크리트포장","콘크리트표층"], [], "콘크리트포장"),

    # 가시설
    (["가시설","흙막이","안전난간"], [], "가시설흙막이"),

    # 추진공
    (["추진","강관압입"], ["D800","800mm"], "강관압입추진_D800"),
    (["추진","강관압입"], ["D600","600mm"], "강관압입추진_D600"),
    (["추진","강관압입"], [],               "강관압입추진_D450"),
]

def get_work_key(name: str, spec: str) -> str | None:
    """공종명+규격 → DAILY_WORK 키 반환"""
    for name_kws, spec_kws, key in WORK_KEY_MAP:
        if not any(kw in name for kw in name_kws):
            continue
        if spec_kws and not any(kw in name+spec for kw in spec_kws):
            continue
        return key
    return None

def calc_work_days(name: str, spec: str, qty: float,
                   crews: int = None, hours: int = 8) -> dict:
    """
    작업일수 계산
    반환: {
        "key": 매핑키,
        "daily": 1일작업량,
        "unit": 단위,
        "crews": 조수,
        "hours": 작업시간,
        "work_days": 작업일수,
        "condition": 작업조건,
    }
    """
    key = get_work_key(name, spec)
    if not key or not qty:
        return None

    info = DAILY_WORK[key].copy()
    if crews:
        info["crews"] = crews
    if hours:
        info["hours"] = hours

    daily_total = info["daily"] * info["crews"]
    work_days   = qty / daily_total if daily_total > 0 else 0

    info["key"]       = key
    info["qty"]       = qty
    info["work_days"] = round(work_days, 2)
    info["work_days_ceil"] = math.ceil(work_days)
    return info

import math

if __name__ == "__main__":
    print("=== 1일 작업량 기반 작업일수 테스트 ===")
    tests = [
        ("터파기(B=6.0m이상)", "토사,육상", 53227),
        ("터파기(B=6.0m이상)", "토사,용수", 6974),
        ("터파기(B=6.0m이상)", "연암,용수", 1301),
        ("고강성PVC 이중벽관 접합및부설", "D=200mm", 11857),
        ("조립식PC맨홀설치", "D900,하부구체", 502),
        ("되메우기 및 다짐", "관상단,토사", 17563),
        ("아스팔트포장절단", "", 25268),
        ("보조기층포설", "", 5000),
        ("아스콘포장", "", 5000),
    ]
    for name, spec, qty in tests:
        r = calc_work_days(name, spec, qty, crews=5)
        if r:
            print(f"  {name[:25]:25s} {spec[:15]:15s} {qty:>8,} → {r['work_days_ceil']:>4}일 (1일:{r['daily']}{r['unit']}×{r['crews']}조)")
        else:
            print(f"  {name[:25]:25s} → 매핑 없음")