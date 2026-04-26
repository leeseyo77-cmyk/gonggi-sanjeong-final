# labor_rates_2025.py
# 출처: 2025년 건설공사 표준품셈 (국토교통부)
# 단위: 인/단위 (인력합계 기준, 장비제외)
# 주의: 본 딕셔너리는 인력(노무) 품만 포함. 장비경비는 별도계상.

# ══════════════════════════════════════════════════════════════
# 1. 터파기 (기계) - 3-2-4 터파기(기계) 2025년 신설
# 구조: {토질: {Type: 일당시공량(m3)}} → 인/m3 = 1/시공량
# TypeⅠ: 일반, TypeⅡ: 지장물·가시설 방해, TypeⅢ: 협소(상하수도 도심지)
# ══════════════════════════════════════════════════════════════

# 일당 시공량 (굴착기 1대 기준)
EXCAVATION_DAILY_PROD = {
    "보통토사": {"TypeⅠ": 560, "TypeⅡ": 420, "TypeⅢ": 190},
    "혼합토사": {"TypeⅠ": 390, "TypeⅡ": 300, "TypeⅢ": 150},
    "풍화암":   {"TypeⅠ": 38,  "TypeⅡ": 35,  "TypeⅢ": None},
    "연암":     {"TypeⅠ": 30,  "TypeⅡ": 28,  "TypeⅢ": None},
    "보통암":   {"TypeⅠ": 22,  "TypeⅡ": 19,  "TypeⅢ": None},
    "경암":     {"TypeⅠ": 16,  "TypeⅡ": 14,  "TypeⅢ": None},
}

# 보정계수
EXCAVATION_CORRECTION = {
    "용수발생": 0.75,    # 시공량 25% 감 → 1/0.75배 품 증가
    "심도5m초과": 0.91,  # 시공량 9% 감
}

def get_excavation_labor(soil_type="보통토사", work_type="TypeⅡ", corrections=None):
    """
    터파기 인/m3 계산
    - soil_type: 보통토사, 혼합토사, 풍화암, 연암, 보통암, 경암
    - work_type: TypeⅠ, TypeⅡ, TypeⅢ (상하수도 도심지 = TypeⅡ~Ⅲ)
    - corrections: ["용수발생", "심도5m초과"] 등
    반환: 인/m3 (소수점 4자리)
    """
    prod = EXCAVATION_DAILY_PROD.get(soil_type, {}).get(work_type)
    if not prod:
        return None
    
    corr = 1.0
    if corrections:
        for c in corrections:
            if c in EXCAVATION_CORRECTION:
                corr *= EXCAVATION_CORRECTION[c]
    
    # 보정 후 실제 시공량
    actual_prod = prod * corr
    # 인/m3 = 1인·일 / 시공량
    return round(1.0 / actual_prod, 5)

# ══════════════════════════════════════════════════════════════
# 2. 관 부설 및 접합 - 제6장 관부설 및 접합공사 (2025년 보완)
# 구조: {관종: {관경(mm): {직종: 인/본}}}
# 단위: 본당 (6m 직관 기준)
# 시공조건: A=병행, B=분리(75%품=133%시공량), C=단독(50%품=200%시공량)
# ══════════════════════════════════════════════════════════════

# 주철관 타이튼 접합 - 6-2-1 (일당 시공량 → 인/본 변환)
# 배관공(수도) 2인 + 보통인부 1인 = 3인/일당
DUCTILE_IRON_TYTON = {
    125:  {"배관공": round(2/18, 4), "보통인부": round(1/18, 4)},
    150:  {"배관공": round(2/16, 4), "보통인부": round(1/16, 4)},
    200:  {"배관공": round(2/12, 4), "보통인부": round(1/12, 4)},
    250:  {"배관공": round(2/10, 4), "보통인부": round(1/10, 4)},
    300:  {"배관공": round(3/9,  4), "보통인부": round(1/9,  4)},
    350:  {"배관공": round(3/9,  4), "보통인부": round(1/9,  4)},
    400:  {"배관공": round(3/8,  4), "보통인부": round(1/8,  4)},
    450:  {"배관공": round(3/7,  4), "보통인부": round(1/7,  4)},
    500:  {"배관공": round(3/6,  4), "보통인부": round(1/6,  4)},
}

# 원심력 철근콘크리트관 소켓 - 6-6-1 (일당 시공량)
RCPIPE_SOCKET = {
    250:  {"배관공": round(2/20, 4), "보통인부": round(1/20, 4)},
    300:  {"배관공": round(2/15, 4), "보통인부": round(1/15, 4)},
    350:  {"배관공": round(2/13, 4), "보통인부": round(1/13, 4)},
    400:  {"배관공": round(2/11, 4), "보통인부": round(1/11, 4)},
    450:  {"배관공": round(2/9,  4), "보통인부": round(1/9,  4)},
    500:  {"배관공": round(2/8,  4), "보통인부": round(1/8,  4)},
    600:  {"배관공": round(2/6.5,4), "보통인부": round(1/6.5,4)},
    700:  {"배관공": round(2/5.5,4), "보통인부": round(1/5.5,4)},
    800:  {"배관공": round(2/5,  4), "보통인부": round(1/5,  4)},
    900:  {"배관공": round(2/5,  4), "보통인부": round(1/5,  4)},
    1000: {"배관공": round(3/4.5,4), "보통인부": round(1/4.5,4)},
    1100: {"배관공": round(3/4.0,4), "보통인부": round(1/4.0,4)},
    1200: {"배관공": round(3/3.5,4), "보통인부": round(1/3.5,4)},
    1350: {"배관공": round(3/3.5,4), "보통인부": round(1/3.5,4)},
    1500: {"배관공": round(4/3.5,4), "보통인부": round(1/3.5,4)},
}

# 고강성PVC 이중벽관 / 내충격PVC하수관 - 6-4-2 고무링접합 준용 (일당)
# 하수관로에서 가장 많이 쓰이는 관종
PVC_DOUBLEWWALL = {
    100:  {"배관공": round(2/25, 4), "보통인부": round(1/25, 4)},
    150:  {"배관공": round(2/19, 4), "보통인부": round(1/19, 4)},
    200:  {"배관공": round(2/15, 4), "보통인부": round(1/15, 4)},
    250:  {"배관공": round(2/10, 4), "보통인부": round(1/10, 4)},
    300:  {"배관공": round(2/9,  4), "보통인부": round(1/9,  4)},
}

# 내충격PVC수도관 - 6-7-4 (본당)
HI_PVC = {
    50:  {"배관공": 0.07, "보통인부": 0.04},
    75:  {"배관공": 0.09, "보통인부": 0.05},
    100: {"배관공": 0.11, "보통인부": 0.06},
    150: {"배관공": 0.15, "보통인부": 0.08},
    200: {"배관공": 0.19, "보통인부": 0.10},
    250: {"배관공": 0.23, "보통인부": 0.12},
    300: {"배관공": 0.27, "보통인부": 0.14},
}

# 유리섬유복합관(GRP) - 6-7-3 비압력관 (본당)
GRP_PIPE = {
    150:  {"배관공": 0.24, "보통인부": 0.09},
    200:  {"배관공": 0.30, "보통인부": 0.12},
    250:  {"배관공": 0.14, "보통인부": 0.06},
    300:  {"배관공": 0.16, "보통인부": 0.06},
    350:  {"배관공": 0.18, "보통인부": 0.07},
    400:  {"배관공": 0.22, "보통인부": 0.09},
    450:  {"배관공": 0.26, "보통인부": 0.10},
    500:  {"배관공": 0.31, "보통인부": 0.12},
    600:  {"배관공": 0.40, "보통인부": 0.16},
    700:  {"배관공": 0.49, "보통인부": 0.19},
    800:  {"배관공": 0.58, "보통인부": 0.23},
    900:  {"배관공": 0.66, "보통인부": 0.27},
    1000: {"배관공": 0.75, "보통인부": 0.30},
    1200: {"배관공": 0.93, "보통인부": 0.37},
    1350: {"배관공": 1.06, "보통인부": 0.42},
    1500: {"배관공": 1.20, "보통인부": 0.48},
}

# 파형강관 - 6-7-2 (본당)
CORRUGATED_STEEL = {
    250:  {"배관공": 0.04, "보통인부": 0.02},
    300:  {"배관공": 0.06, "보통인부": 0.03},
    400:  {"배관공": 0.10, "보통인부": 0.05},
    450:  {"배관공": 0.12, "보통인부": 0.06},
    500:  {"배관공": 0.13, "보통인부": 0.07},
    600:  {"배관공": 0.17, "보통인부": 0.08},
    700:  {"배관공": 0.21, "보통인부": 0.10},
    800:  {"배관공": 0.24, "보통인부": 0.12},
    1000: {"배관공": 0.32, "보통인부": 0.16},
    1200: {"배관공": 0.39, "보통인부": 0.19},
    1500: {"배관공": 0.50, "보통인부": 0.25},
}

# ══════════════════════════════════════════════════════════════
# 3. 관 부설 시공조건 보정 (6-1-1 표)
# ══════════════════════════════════════════════════════════════
PIPE_INSTALL_CONDITION = {
    "A": {"품요율": 1.00, "시공량배율": 1.00, "설명": "굴착·복구 병행, 연속굴착 불가"},
    "B": {"품요율": 0.75, "시공량배율": 1.33, "설명": "굴착 선행, 부설 연속시공"},
    "C": {"품요율": 0.50, "시공량배율": 2.00, "설명": "굴착 완료 후 단독시공"},
}

# ══════════════════════════════════════════════════════════════
# 4. 통합 조회 함수
# ══════════════════════════════════════════════════════════════

# 관종 매핑 (내역서 공종명 → 딕셔너리 키)
PIPE_TYPE_MAP = {
    "주철관": DUCTILE_IRON_TYTON,
    "타이튼": DUCTILE_IRON_TYTON,
    "원심력철근콘크리트": RCPIPE_SOCKET,
    "RC관": RCPIPE_SOCKET,
    "흄관": RCPIPE_SOCKET,
    "고강성PVC": PVC_DOUBLEWWALL,
    "이중벽관": PVC_DOUBLEWWALL,
    "PE다중벽": PVC_DOUBLEWWALL,
    "내충격PVC": HI_PVC,
    "유리섬유복합관": GRP_PIPE,
    "GRP관": GRP_PIPE,
    "파형강관": CORRUGATED_STEEL,
}

def get_pipe_labor(pipe_name: str, diameter_mm: int, condition="A") -> dict:
    """
    관 부설 인/본 계산
    - pipe_name: 내역서 공종명 (예: "고강성PVC 이중벽관 접합및부설")
    - diameter_mm: 관경 (mm 정수)
    - condition: A, B, C (시공조건)
    반환: {"배관공": X, "보통인부": Y, "합계": Z, "단위": "인/본"}
    """
    # 관종 딕셔너리 탐색
    pipe_dict = None
    for key, d in PIPE_TYPE_MAP.items():
        if key in pipe_name:
            pipe_dict = d
            break
    
    if pipe_dict is None:
        pipe_dict = PVC_DOUBLEWWALL  # 기본값
    
    # 가장 가까운 관경 찾기
    available = sorted(pipe_dict.keys())
    closest = min(available, key=lambda x: abs(x - diameter_mm))
    rates = pipe_dict[closest]
    
    # 시공조건 보정
    cond = PIPE_INSTALL_CONDITION.get(condition, PIPE_INSTALL_CONDITION["A"])
    factor = cond["품요율"]
    
    result = {}
    total = 0
    for labor_type, rate in rates.items():
        adjusted = round(rate * factor, 4)
        result[labor_type] = adjusted
        total += adjusted
    
    result["합계"] = round(total, 4)
    result["단위"] = "인/본"
    result["관경"] = closest
    result["시공조건"] = condition
    return result

def get_excavation_labor_detail(spec_str: str) -> dict:
    """
    내역서 규격 문자열에서 토질·조건 자동 추출 후 인/m3 계산
    - spec_str: 예) "토사,육상", "토사,용수", "연암", "B=4.0m이상,토사"
    """
    soil_type = "보통토사"
    work_type = "TypeⅡ"  # 상하수도 도심지 기본값
    corrections = []
    
    # 토질 판단
    if "경암" in spec_str:
        soil_type = "경암"
    elif "보통암" in spec_str:
        soil_type = "보통암"
    elif "연암" in spec_str:
        soil_type = "연암"
    elif "풍화암" in spec_str:
        soil_type = "풍화암"
    elif "혼합" in spec_str or "자갈" in spec_str or "호박돌" in spec_str:
        soil_type = "혼합토사"
    else:
        soil_type = "보통토사"
    
    # 용수 여부
    if "용수" in spec_str:
        corrections.append("용수발생")
    
    # 심도
    if "5m초과" in spec_str or "깊이5m" in spec_str:
        corrections.append("심도5m초과")
    
    labor = get_excavation_labor(soil_type, work_type, corrections)
    
    return {
        "토질": soil_type,
        "작업유형": work_type,
        "보정조건": corrections,
        "인/m3": labor,
        "단위": "인/m3"
    }


if __name__ == "__main__":
    print("=== 터파기 품셈 테스트 ===")
    cases = [
        ("보통토사", "TypeⅡ", []),
        ("보통토사", "TypeⅡ", ["용수발생"]),
        ("연암",     "TypeⅠ", []),
        ("경암",     "TypeⅠ", []),
    ]
    for soil, wt, corr in cases:
        val = get_excavation_labor(soil, wt, corr)
        print(f"  {soil} {wt} {corr}: {val} 인/m3")

    print("\n=== 관 부설 품셈 테스트 ===")
    cases2 = [
        ("고강성PVC 이중벽관 접합및부설", 200, "A"),
        ("고강성PVC 이중벽관 접합및부설", 300, "A"),
        ("원심력철근콘크리트관 부설",      600, "A"),
        ("주철관 타이튼 부설",             300, "B"),
    ]
    for name, dia, cond in cases2:
        result = get_pipe_labor(name, dia, cond)
        print(f"  {name} D={dia}mm [{cond}]: {result['합계']} 인/본")

    print("\n=== 내역서 규격 자동 파싱 테스트 ===")
    specs = ["토사,육상", "토사,용수", "연암", "경암"]
    for s in specs:
        r = get_excavation_labor_detail(s)
        print(f"  '{s}' → {r['토질']} {r['보정조건']}: {r['인/m3']} 인/m3")
