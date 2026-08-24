# -*- coding: utf-8 -*-
"""
pumsem_crew.py — 2026년 건설공사 표준품셈 발췌 (작업조 구성 + 일당 시공량)

배경
----
품셈 1-2-8(작업조 구성 및 적용)은 이렇게 정의한다.
    "작업조는 일당시공량을 시공하기 위한 필수자원(인력, 장비)의 조합으로 제시"
    "인력품 산정(인) : 인력(인) ÷ 시공량(일당)"

즉 품셈의 시공량은 **1개 작업조** 기준이고, 앱의 '투입조수 1'도 이 작업조 1세트로
보는 것이 실무 감각(세트로 움직임)과 품셈 체계에 모두 맞는다.

반면 일위대가 노무비를 노임단가로 나누는 역산은 **1인 기준** 값이 나오므로,
같은 화면에서 두 값을 섞으면 조수의 의미가 달라져 혼란이 생긴다.
이 모듈은 상하수도 공사에서 자주 쓰는 공종에 대해 품셈 원문의
작업조 인원과 일당 시공량을 담아, 역산 대신 품셈 값을 쓰도록 한다.

수록 범위
--------
7개 실제 내역서(관로·하수처리장·정수장)에서 사용 빈도가 높은 공종 위주.
품셈 전체(2,237개 항목)를 자동 추출하면 표 구조가 제각각이라 오탐이 생기므로,
사용 빈도가 높은 항목만 원문을 확인해 수록했다.

각 항목 구조
-----------
    "키": {
        "code":    품셈 항목번호(추적용),
        "name":    품셈 항목명,
        "crew":    {직종: 인원} — 1개 작업조 구성(인력만)
        "persons": 작업조 인력 합계(인)
        "equip":   {장비: 대수} — 참고용
        "daily":   {규격: 시공량} 또는 단일 값
        "unit":    시공량 단위
        "note":    적용 조건
    }
"""

# ── 관 부설 및 접합 ──────────────────────────────────────────
# 6-6-2 수밀밴드 접합 및 부설: 철근콘크리트관(흄관) 부설·접합
HUME_PIPE_LAYING = {
    "code": "6-6-2",
    "name": "수밀밴드 접합 및 부설(철근콘크리트관)",
    "unit": "본",
    "note": "관부설, 수밀밴드 접합, 위치·구배 확인, 관로표시테이프 부설 포함",
    # 관경 구간별로 작업조 구성이 달라진다
    "ranges": [
        {"max_dia": 800,  "crew": {"배관공(수도)": 2, "보통인부": 1}, "persons": 3,
         "equip": {"양중장비": 1},
         "daily": {250: 21, 300: 16, 350: 13, 400: 11, 450: 10, 500: 9,
                   600: 7, 700: 6, 800: 5}},
        {"max_dia": 1350, "crew": {"배관공(수도)": 3, "보통인부": 1}, "persons": 4,
         "equip": {"양중장비": 1},
         "daily": {900: 5, 1000: 5, 1100: 4, 1200: 4, 1350: 3.5}},
        {"max_dia": 2000, "crew": {"배관공(수도)": 4, "보통인부": 1}, "persons": 5,
         "equip": {"양중장비": 1},
         "daily": {1500: 3.5, 1650: 3, 1800: 3, 2000: 2.5}},
    ],
}

# 2-4-14 원심력철근콘크리트관 철거
HUME_PIPE_REMOVAL = {
    "code": "2-4-14",
    "name": "원심력철근콘크리트관 철거",
    "unit": "본",
    "crew": {"배관공(수도)": 2, "보통인부": 1},
    "persons": 3,
    "daily": {250: 43, 300: 39, 350: 35, 400: 31, 450: 28, 500: 26,
              600: 22, 700: 18, 800: 16, 900: 13, 1000: 11},
    "note": "관 철거 기준(부설 품셈과 별개)",
}

# ── 토공 ────────────────────────────────────────────────────
# 3-4-6 되메우기 및 다짐(소형장비)
BACKFILL_SMALL = {
    "code": "3-4-6",
    "name": "되메우기 및 다짐(소형장비)",
    "unit": "㎥",
    "crew": {"특별인부": 1, "보통인부": 1},
    "persons": 2,
    "equip": {"굴착기0.2㎥": 1, "진동롤러(핸드가이드식)0.7ton": 1, "살수차": 0.5},
    "daily": 130,
    "note": "포설 및 고르기, 다짐 작업 포함",
}

# ── 콘크리트 ────────────────────────────────────────────────
# 6-1-1 레디믹스트콘크리트 타설
CONCRETE_PLACING = {
    "code": "6-1-1",
    "name": "레디믹스트콘크리트 타설",
    "unit": "㎥",
    "types": {
        "인력운반": {"crew": {"콘크리트공": 3, "보통인부": 3}, "persons": 6,
                  "daily": {"무근구조물": 23, "철근구조물": 20}},
        "장비사용": {"crew": {"콘크리트공": 3, "보통인부": 1}, "persons": 4,
                  "equip": {"굴착기(0.6~0.8㎥)": 1},
                  "daily": {"무근구조물": 63, "철근구조물": 55}},
    },
    "note": "개소별 소량(12㎥ 이하)이 산재하면 시공량 50%까지 감할 수 있음",
}

# ── 가설 ────────────────────────────────────────────────────
# 2-6-3 시스템 동바리 설치 및 해체 (설치/해체 각각 별도 조)
SYSTEM_SUPPORT = {
    "code": "2-6-3",
    "name": "시스템 동바리 설치 및 해체",
    "unit": "공㎥",
    "phases": {
        "설치": {
            "crew_by_height": {
                "5m이하":        {"crew": {"형틀목공": 4, "보통인부": 1}, "persons": 5, "daily": 130},
                "5m초과~10m이하":  {"crew": {"형틀목공": 4, "보통인부": 1}, "persons": 5,
                                 "equip": {"크레인": 0.5}, "daily": 120},
                "10m초과~20m이하": {"crew": {"형틀목공": 4, "보통인부": 1}, "persons": 5,
                                 "equip": {"크레인": 0.5}, "daily": 105},
                "20m초과~30m이하": {"crew": {"형틀목공": 4, "보통인부": 1}, "persons": 5,
                                 "equip": {"크레인": 0.5}, "daily": 85},
            }
        },
        "해체": {
            "crew_by_height": {
                "5m이하":        {"crew": {"형틀목공": 2, "보통인부": 2}, "persons": 4, "daily": 150},
                "5m초과~10m이하":  {"crew": {"형틀목공": 2, "보통인부": 2}, "persons": 4,
                                 "equip": {"크레인": 0.5}, "daily": 150},
                "10m초과~20m이하": {"crew": {"형틀목공": 2, "보통인부": 2}, "persons": 4,
                                 "equip": {"크레인": 0.5}, "daily": 140},
                "20m초과~30m이하": {"crew": {"형틀목공": 2, "보통인부": 2}, "persons": 4,
                                 "equip": {"크레인": 0.5}, "daily": 115},
            }
        },
    },
    # 설치간격에 따른 요율(멍에간격 기준)
    "spacing_factor": {"0.6m이하": -0.17, "0.6m초과~0.8m이하": 0.0, "0.8m초과": +0.11},
    "note": "설치·해체를 함께 하는 항목은 두 작업 소요시간을 합산(1/설치 + 1/해체)",
}

# 2-3-1 철제조립식 가설건축물 설치 및 해체 (바닥면적 ㎡당 인수 방식)
TEMP_BUILDING = {
    "code": "2-3-1",
    "name": "철제조립식 가설건축물 설치 및 해체",
    "unit": "㎡",
    "per_unit_labor": {   # ㎡당 소요 인수
        "사무실": {"건축목공": 0.26, "보통인부": 0.11, "persons_per_unit": 0.37},
        "창고":   {"건축목공": 0.20, "보통인부": 0.09, "persons_per_unit": 0.29},
    },
    "equip": {"크레인10ton": {"사무실": 0.19, "창고": 0.15}},  # hr
    "note": "샌드위치판넬 조립식 가설건축물. 착공 전 준비기간에 포함되는 것이 일반적",
}

# 2-4-1 강관 지주 설치 및 해체 (가설울타리)
TEMP_FENCE_POST = {
    "code": "2-4-1",
    "name": "강관 지주 설치 및 해체",
    "unit": "m",
    "crew": {"비계공": 3, "보통인부": 1},
    "persons": 4,
    "equip": {"굴착기0.2㎥": 0.5},
    "daily": {"지주높이4m이하_설치": 100, "지주높이4m이하_해체": 250,
              "지주높이7m이하_설치": 70,  "지주높이7m이하_해체": 180},
    "note": "지주간격 2.0m 기준. 지반평탄, 강관매입, 보조기둥 포함",
}

# ── 포장 ────────────────────────────────────────────────────
# 1-6-6 포장줄눈 절단
PAVEMENT_CUT = {
    "code": "1-6-6",
    "name": "포장줄눈 절단",
    "unit": "m",
    "crew": {"특별인부": 1, "보통인부": 1},
    "persons": 2,
    "equip": {"커터320~400㎜": 1},
    "daily": 600,
    "note": "콘크리트포장 표층면 절단(절단깊이 10㎝ 이하), 절단면 물청소 포함",
}


# ── 조회 헬퍼 ────────────────────────────────────────────────

ALL_ITEMS = {
    "흄관부설접합": HUME_PIPE_LAYING,
    "흄관철거": HUME_PIPE_REMOVAL,
    "되메우기소형": BACKFILL_SMALL,
    "콘크리트타설": CONCRETE_PLACING,
    "시스템동바리": SYSTEM_SUPPORT,
    "가설건축물": TEMP_BUILDING,
    "가설울타리지주": TEMP_FENCE_POST,
    "포장줄눈절단": PAVEMENT_CUT,
}


def hume_pipe_rate(diameter_mm, removal=False):
    """흄관 관경별 (시공량, 작업조 인원) 반환. 없으면 None.

    removal=True면 철거 품셈(2-4-14), 아니면 부설·접합(6-6-2).
    """
    if removal:
        d = HUME_PIPE_REMOVAL["daily"]
        if diameter_mm in d:
            return d[diameter_mm], HUME_PIPE_REMOVAL["persons"]
        # 가장 가까운 관경으로 근사
        keys = sorted(d)
        for k in keys:
            if diameter_mm <= k:
                return d[k], HUME_PIPE_REMOVAL["persons"]
        return d[keys[-1]], HUME_PIPE_REMOVAL["persons"]

    for rng in HUME_PIPE_LAYING["ranges"]:
        if diameter_mm <= rng["max_dia"]:
            d = rng["daily"]
            if diameter_mm in d:
                return d[diameter_mm], rng["persons"]
            for k in sorted(d):
                if diameter_mm <= k:
                    return d[k], rng["persons"]
            return d[sorted(d)[-1]], rng["persons"]
    return None


def system_support_rate(height_label, spacing_label=None):
    """시스템 동바리 설치+해체 합산 시공량(공㎥/조·일)과 조 인원 반환.

    설치와 해체는 별도 작업이므로 소요시간을 합산한다.
    조 인원은 설치조(5인) 기준으로 본다(해체조 4인보다 크므로 보수적).
    """
    inst = SYSTEM_SUPPORT["phases"]["설치"]["crew_by_height"].get(height_label)
    demo = SYSTEM_SUPPORT["phases"]["해체"]["crew_by_height"].get(height_label)
    if not inst or not demo:
        return None
    combined = 1.0 / (1.0 / inst["daily"] + 1.0 / demo["daily"])
    if spacing_label:
        f = SYSTEM_SUPPORT["spacing_factor"].get(spacing_label, 0.0)
        combined *= (1.0 + f)
    return round(combined, 2), inst["persons"]


def temp_building_rate(kind="사무실"):
    """가설건축물 ㎡당 인수 → 1인 기준 시공량(㎡/인일) 반환."""
    info = TEMP_BUILDING["per_unit_labor"].get(kind)
    if not info:
        return None
    per = info["persons_per_unit"]
    return round(1.0 / per, 2), per