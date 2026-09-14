import streamlit as st
import pandas as pd
import openpyxl
import zipfile
from openpyxl.utils.exceptions import InvalidFileException
from universal_parser import detect_template, parse_items_generic, parse_unit_price_generic, parse_labor_derived_by_template, apply_series_fallback
import math
import re
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

# 가이드라인 데이터 및 기상 데이터 import
try:
    from pumsem_crew import hume_pipe_rate, system_support_rate
    HAS_PUMSEM_CREW = True
except Exception:
    HAS_PUMSEM_CREW = False

try:
    from daily_work_rates import DAILY_WORK, WORK_KEY_MAP
    HAS_DAILY_WORK = True
except Exception:
    HAS_DAILY_WORK = False
    DAILY_WORK, WORK_KEY_MAP = {}, []

try:
    from guideline_weather_data import (
        WEATHER_NON_WORK, STATIONS, CONDITION_LABELS, DEFAULT_CONDITIONS,
        get_weather_non_work_days,
    )
    HAS_GUIDELINE_WEATHER = True
except Exception:
    HAS_GUIDELINE_WEATHER = False
    STATIONS, CONDITION_LABELS, DEFAULT_CONDITIONS = [], {}, ()

try:
    from guideline_data import GUIDELINE_APPENDIX_FULL, GUIDELINE_APPENDIX
    from weather_data import (
        RAIN_DAYS, COLD_DAYS, HOT_DAYS, REGION_MAPPING,
        get_total_non_work_days, get_monthly_breakdown
    )
    from holiday_data import (
        LEGAL_HOLIDAYS, get_legal_holidays, get_total_holidays,
        calc_overlap_days, get_total_non_work_days_with_holidays,
        get_holiday_breakdown_monthly
    )
    REGIONS = list(REGION_MAPPING.keys())
except ImportError as e:
    print(f"Import 실패: {e}")
    # 파일이 없을 경우 기본값 사용
    GUIDELINE_APPENDIX_FULL = {}
    GUIDELINE_APPENDIX = {}
    RAIN_DAYS = {}
    COLD_DAYS = {}
    HOT_DAYS = {}
    REGION_MAPPING = {"서울": "서울"}
    REGIONS = ["서울"]
    LEGAL_HOLIDAYS = {}
    def get_total_non_work_days(region, start_date, end_date, check_rain=True, check_cold=True, check_hot=True):
        return 0
    def get_monthly_breakdown(region, start_date, end_date, check_rain=True, check_cold=True, check_hot=True):
        return []
    def get_legal_holidays(year, month):
        return 0
    def get_total_holidays(start_date, end_date):
        return 0
    def calc_overlap_days(a, b, c):
        return 0
    def get_total_non_work_days_with_holidays(*args, **kwargs):
        return {"total": 0, "weather": 0, "holidays": 0, "overlap": 0, "formula": ""}
    def get_holiday_breakdown_monthly(start_date, end_date):
        return []

st.set_page_config(page_title="상하수도 공기산정", layout="wide", initial_sidebar_state="expanded")

# ══════════════════════════════════════════════════════════════
# 공종 키워드 매핑
# ══════════════════════════════════════════════════════════════
KEYWORD_MAP_DETAIL = {
    "포장복구": ["포장복구","아스팔트포장","아스팔트+콘크리트포장","콘크리트포장","보도포장","인도포장","보조기층","택코팅","프라임코팅","기층","표층","차선도색","노면절삭","아스팔트 노면 절삭","절삭후","보도블럭","과속방지턱","줄눈설치","줄눈"],
    "굴착공": ["터파기","굴착","줄파기","착공","시굴","포장깨기","포장절단"],
    "관부설공": ["관 부설","관부설","PE다중벽관","고강성PVC","주철관","GRP관",
                 "유리섬유복합관","흄관","이중벽관","강관부설","콘크리트관"],
    "되메우기": ["되메우기","뒤채움","복토","성토"],
    "맨홀공": ["맨홀","우수받이","집수정","토실","슬라이딩"],
    "배수설비": ["배수설비","빗물받이"],
    "추진공": ["추진공","추진관","추진"],
}

SKIP_NAMES = [
    "남천지구","동부지구","신설오수관로","간선관로","지선관로",
    "순공사비","배수설비공사","토공","관로공","구조물공","포장공",
    "추진공","부대공","안전관리비","환경보전비","소계","합계","계",
]

PIPE_EXCLUDE = ["절단","이형관","하차비","단관","마감캡","추진관"]
MACHINE_BASED = ["터파기","굴착","되메우기","모래기초","모래부설","모래,관기초"]

# ══════════════════════════════════════════════════════════════
# 가이드라인 부록 데이터 (대폭 확장)
# ══════════════════════════════════════════════════════════════
GUIDELINE_APPENDIX = {
    # 포장공
    "아스팔트포장 절단": {"daily": 1000, "unit": "m"},
    "아스팔트포장깨기 (B.H0.4㎥)": {"daily": 515, "unit": "㎡"},
    "아스팔트포장깨기 (B.H0.7㎥)": {"daily": 1047, "unit": "㎡"},
    "콘크리트포장 절단": {"daily": 500, "unit": "m"},
    "콘크리트포장 깨기": {"daily": 300, "unit": "㎡"},
    "아스팔트포장(기층)": {"daily": 800, "unit": "㎡"},
    "아스팔트포장(택코팅)": {"daily": 2000, "unit": "㎡"},
    "아스팔트포장(프라임코팅)": {"daily": 2000, "unit": "㎡"},
    "아스팔트포장(표층)": {"daily": 1000, "unit": "㎡"},
    "택코팅": {"daily": 2000, "unit": "㎡"},
    "프라임코팅": {"daily": 2000, "unit": "㎡"},
    "기층": {"daily": 800, "unit": "㎡"},
    "표층": {"daily": 1000, "unit": "㎡"},
    
    # 터파기
    "터파기(토사:육상) B/H 0.4㎥": {"daily": 260, "unit": "㎥"},
    "터파기(토사:육상) B/H 0.7㎥": {"daily": 530, "unit": "㎥"},
    "터파기(암:육상) B/H 0.4㎥": {"daily": 130, "unit": "㎥"},
    "터파기(암:육상) B/H 0.7㎥": {"daily": 265, "unit": "㎥"},
    
    # 되메우기 (폭별 추가)
    "되메우기(진동롤러) 2.5ton": {"daily": 600, "unit": "㎥"},
    "되메우기(진동롤러) 4.0ton": {"daily": 950, "unit": "㎥"},
    "되메우기(진동콤팩터)": {"daily": 400, "unit": "㎥"},
    "되메우기(B=1.5~2.5m)": {"daily": 450, "unit": "㎥"},
    "되메우기(B=2.5~4.0m)": {"daily": 650, "unit": "㎥"},
    "되메우기(B=4.0m이상)": {"daily": 850, "unit": "㎥"},
    
    # 관부설
    "관부설(D200)": {"daily": 5, "unit": "본/일"},
    "관부설(D300)": {"daily": 4, "unit": "본/일"},
    "관부설(D450)": {"daily": 3, "unit": "본/일"},
    "관부설(D600)": {"daily": 2.5, "unit": "본/일"},
    "관부설(D800)": {"daily": 2, "unit": "본/일"},
    "관부설(D1000)": {"daily": 1.5, "unit": "본/일"},
    "관부설(D1200)": {"daily": 1.2, "unit": "본/일"},
    
    # 맨홀공 (대폭 확장)
    "원형맨홀 Φ1200": {"daily": 2.5, "unit": "개소/일"},
    "원형맨홀 Φ1500": {"daily": 2.5, "unit": "개소/일"},
    "각형맨홀 1800×2400": {"daily": 1.5, "unit": "개소/일"},
    "우수받이": {"daily": 5, "unit": "개소/일"},
    "조립식 PC맨홀": {"daily": 3, "unit": "개소/일"},
    "GRP맨홀": {"daily": 2, "unit": "개소/일"},
    "조립식맨홀설치(소형)": {"daily": 8, "unit": "개소/일"},
    "조립식맨홀 상부구체": {"daily": 10, "unit": "개소/일"},
    "조립식맨홀 연직구체": {"daily": 12, "unit": "개소/일"},
    "조립식맨홀 하부구체": {"daily": 8, "unit": "개소/일"},
    "맨홀뚜껑설치": {"daily": 20, "unit": "개소/일"},
    "맨홀뚜껑": {"daily": 20, "unit": "개소/일"},
    
    # 배수설비
    "빗물받이": {"daily": 5, "unit": "개소/일"},
    "집수받이": {"daily": 4, "unit": "개소/일"},
    "우수토실": {"daily": 3, "unit": "개소/일"},
    "배수설비": {"daily": 4, "unit": "개소/일"},
    
    # 추진공
    "추진설비공": {"daily": 1, "unit": "개소/일"},
    "사토": {"daily": 100, "unit": "㎥/일"},
    "추진마감벽설치": {"daily": 3, "unit": "개소/일"},
    "추진마감벽": {"daily": 3, "unit": "개소/일"},
    "천공홀 되메우기": {"daily": 500, "unit": "m"},
    
    # 가시설공
    "조립식 간이 흙막이": {"daily": 50, "unit": "㎡/일"},
    "H-PILE 항타": {"daily": 8, "unit": "본/일"},
    
    # 기타
    "보조기층": {"daily": 500, "unit": "㎡/일"},
    "모래기초": {"daily": 400, "unit": "㎥/일"},
}

# ══════════════════════════════════════════════════════════════
# 표준품셈 노무량
# ══════════════════════════════════════════════════════════════
def get_excavation_labor(spec_str):
    labor_table = {
        0.4: {"인/m3": 0.130},
        0.7: {"인/m3": 0.085},
        1.0: {"인/m3": 0.070},
    }
    if "0.4" in spec_str or "B.H0.4" in spec_str or "B/H 0.4" in spec_str:
        return labor_table[0.4]
    elif "0.7" in spec_str or "B.H0.7" in spec_str or "B/H 0.7" in spec_str:
        return labor_table[0.7]
    elif "1.0" in spec_str or "B.H1.0" in spec_str or "B/H 1.0" in spec_str:
        return labor_table[1.0]
    return {"인/m3": 0.085}

def get_pipe_labor(diameter):
    pipe_labor = {
        200: {"합계": 0.396},
        300: {"합계": 0.494},
        450: {"합계": 0.653},
        600: {"합계": 0.792},
        800: {"합계": 0.990},
        1000: {"합계": 1.188},
        1200: {"합계": 1.386},
    }
    closest = min(pipe_labor.keys(), key=lambda x: abs(x - diameter))
    return pipe_labor.get(closest, {"합계": 0.5})

def is_machine_based(name):
    return any(kw in name for kw in MACHINE_BASED)

def extract_diameter(spec_str):
    patterns = [r'D\s*[=＝]?\s*(\d+)',r'Φ\s*(\d+)',r'φ\s*(\d+)',
                r'(\d{2,4})\s*(?:mm|㎜)',r'[D=]?(\d{2,4})']
    for pat in patterns:
        m = re.search(pat, spec_str)
        if m:
            val = int(m.group(1))
            if 50 <= val <= 3000:
                return val
    return None

# 공기 산정에서 제외할 대공종 키워드.
# 자재·운반은 병행 작업이라 공기를 만들지 않고, 나머지는 시공이 아닌 비용·서류 항목이다.
# (풍각 공식 산정 대조: 시공 8개 공종만 작업일수에 반영하고
#  품질관리비·시공상세도·안전관리비·부지임대료·지반조사 등은 제외했음)
NON_WORK_CATEGORY_KEYWORDS = (
    "자재", "운반",
    "관리비", "작성비", "임대료", "지반조사", "시운전비", "관리 활동비",
)


# 기본 투입조수. 1조 기준으로 산정하고, 필요한 공종만 조수를 올리는 방식이
# 실무 감각에 맞고 과소 산정을 막는다.
DEFAULT_CREW = 1


def days_to_months_text(days) -> str:
    """일수를 '≒ N개월' 표기로. 30.4일/월(연 365일÷12) 기준."""
    try:
        d = float(days)
    except (TypeError, ValueError):
        return ""
    if d <= 0:
        return ""
    # 표시 자릿수(소수 1자리)로 먼저 반올림한다. 안 그러면 47.96개월이
    # '3년 12.0개월'로 표시된다(실측).
    m = round(d / 30.4, 1)
    # 12개월 이상이면 'N년 M개월'도 함께
    if m >= 12:
        y = int(m // 12)
        rm = round(m - y * 12, 1)
        return f"≒ {m:.1f}개월 ({y}년 {rm:.1f}개월)"
    return f"≒ {m:.1f}개월"


def applied_condition_lines(result) -> list:
    """비작업일수 계산에 실제로 적용한 기상조건을 '✅ 조건명' 목록으로 반환(표시용).

    가이드라인 13개 조건 방식이면 계산 시 저장한 조건 목록(conditions)을 쓴다.
    예전처럼 강우/한랭/폭염 플래그를 읽으면, 가이드라인 방식에서는 호환용으로
    항상 True라서 실제 선택과 무관하게 전부 ✅로 표시되었다.
    """
    conds = result.get("conditions")
    if conds is not None:
        if not conds:
            return ["❌ 기상조건 미적용"]
        return [f"✅ {CONDITION_LABELS.get(c, c)}" for c in conds]
    return [
        f"{'✅' if result.get('include_rain') else '❌'} 강우일",
        f"{'✅' if result.get('include_cold') else '❌'} 한랭일 (-10°C 이하)",
        f"{'✅' if result.get('include_hot') else '❌'} 폭염일 (33°C 이상)",
    ]


# 출처별 신뢰도 아이콘.
#   ✅ 설계자 직접 산정값(Q산식)·수동입력·가이드라인 → 그대로 사용 가능
#   ⚠️ 승계·역산 추정값 → 실제와 다를 수 있어 검토 권장
#   ❌ 매칭 안 됨 → 값이 없어 공기에 반영되지 않음(수동입력 필요)
#   ⏸ 병행(제외) → 의도적으로 0일 처리
def source_badge(method: str) -> str:
    m = str(method or "")
    if m.startswith("노무비역산") or m.startswith("동일계열"):
        return "⚠️ " + m
    if m == "매칭 안 됨":
        return "❌ " + m
    if m.startswith("병행") or m in ("사용자 제외", "준비기간 포함"):
        return "⏸ " + m
    if m in ("-", ""):
        return m
    return "✅ " + m


def _rate_to_float(rate_text):
    """'138.2M3' / '5본/일' 같은 표기에서 숫자만 추출."""
    m = re.search(r"[\d,]+(?:\.\d+)?", str(rate_text or ""))
    if not m:
        return None
    try:
        return float(m.group(0).replace(",", ""))
    except ValueError:
        return None


@st.fragment
def render_detail_table(detail_items, source_items, base_crew, key):
    """상세공종 표를 '1일작업량 수정'과 '투입조수' 열만 편집 가능하게 렌더링.

    - 투입조수: 항목별 조수(crew_by_item). 같은 공종이라도 장비·인원이 다른 항목에 사용.
    - 1일작업량 수정: 값을 넣으면 수동입력(manual_rates)으로 저장되어 0순위로 적용된다.
      추정값(노무비역산·승계)이 실제와 다를 때 표에서 바로 고칠 수 있게 하려는 것으로,
      기존에는 '매칭 안 됨' 항목만 수동입력이 가능해 이미 값이 있는 항목은
      화면에서 고칠 방법이 없었다.
    변경값은 calc_days_priority가 최우선으로 참조하므로 총계·소계·상세에 즉시 반영된다.
    """
    if not detail_items:
        return
    # 내부용 키(_로 시작)는 화면에 노출하지 않는다
    _view = []
    for i, d in enumerate(detail_items):
        row = {k: v for k, v in d.items() if not str(k).startswith("_")}
        it = source_items[i] if i < len(source_items) else {}
        ik = f"{it.get('name', '')}|{it.get('spec', '')}"
        # 아직 반영 전인 편집값이 있으면 그것을 보여준다(입력이 사라지지 않도록)
        _pend_r = st.session_state.get("pending_rate", {}).get(ik)
        if _pend_r:
            row["1일작업량 수정"] = float(_pend_r[0])
        else:
            mr = st.session_state.get("manual_rates", {}).get(ik, {})
            row["1일작업량 수정"] = float(mr.get("daily", 0.0) or 0.0)
        _pend_c = st.session_state.get("pending_crew", {}).get(ik)
        if _pend_c:
            row["투입조수"] = int(_pend_c[0])
        # 제외 여부 (사용자가 공기에서 빼기로 판단한 항목)
        _pend_x = st.session_state.get("pending_exclude", {}).get(ik)
        if _pend_x is not None:
            row["제외"] = bool(_pend_x)
        else:
            row["제외"] = ik in st.session_state.get("excluded_items", set())
        _view.append(row)

    _editable = {"투입조수", "1일작업량 수정", "제외"}
    edited = st.data_editor(
        pd.DataFrame(_view),
        hide_index=True,
        width="stretch",
        key=key,
        disabled=[c for c in _view[0].keys() if c not in _editable],
        column_config={
            "투입조수": st.column_config.NumberColumn(
                "투입조수", min_value=1, max_value=200, step=1,
                help="이 항목만의 투입조수. 값을 바꾸면 작업일수가 즉시 다시 계산됩니다.",
            ),
            "1일작업량 수정": st.column_config.NumberColumn(
                "1일작업량 수정", min_value=0.0, step=0.1, format="%.2f",
                help="추정값이 실제와 다르면 여기에 직접 입력하세요(0이면 기존 값 사용). "
                     "입력하면 수동입력으로 저장되어 가장 먼저 적용됩니다.",
            ),
            "제외": st.column_config.CheckboxColumn(
                "제외", help="체크하면 이 항목을 공기 계산에서 뺍니다(0일 처리). "
                            "본공정과 병행되거나 공기에 영향이 없다고 판단한 항목에 사용하세요.",
            ),
        },
    )
    st.session_state.setdefault("crew_by_item", {})
    st.session_state.setdefault("manual_rates", {})
    # 편집값은 곧바로 반영하지 않고 '대기 목록'에 모은다.
    # 셀을 하나 고칠 때마다 전체 재계산(st.rerun)이 돌면 매번 수 초씩 걸려
    # 여러 항목을 연달아 수정하기 어렵기 때문이다.
    st.session_state.setdefault("pending_crew", {})
    st.session_state.setdefault("pending_rate", {})
    st.session_state.setdefault("pending_exclude", {})
    st.session_state.setdefault("excluded_items", set())

    pending_here = 0
    for i, item in enumerate(source_items):
        if i >= len(edited):
            break
        ik = f"{item['name']}|{item.get('spec', '')}"

        # 투입조수
        try:
            new_c = int(edited.iloc[i]["투입조수"])
            cur_c = int(st.session_state["crew_by_item"].get(ik, base_crew))
            if new_c != cur_c:
                st.session_state["pending_crew"][ik] = (new_c, int(base_crew))
                pending_here += 1
            else:
                st.session_state["pending_crew"].pop(ik, None)
        except Exception:
            pass

        # 제외 체크
        try:
            new_x = bool(edited.iloc[i]["제외"])
            cur_x = ik in st.session_state["excluded_items"]
            if new_x != cur_x:
                st.session_state["pending_exclude"][ik] = new_x
            else:
                st.session_state["pending_exclude"].pop(ik, None)
        except Exception:
            pass

        # 1일작업량 수정
        try:
            new_r = float(edited.iloc[i]["1일작업량 수정"] or 0.0)
            cur_r = float(st.session_state["manual_rates"].get(ik, {}).get("daily", 0.0) or 0.0)
            if abs(new_r - cur_r) > 1e-9:
                st.session_state["pending_rate"][ik] = (new_r, item.get("unit", ""))
                pending_here += 1
            else:
                st.session_state["pending_rate"].pop(ik, None)
        except Exception:
            pass

    # 이 표에서 수정한 항목만 세어 바로 아래에 반영 버튼을 둔다.
    # @st.fragment 덕분에 셀 편집 시 이 표만 다시 그려지고 전체 재계산은 일어나지 않는다.
    # 버튼을 누를 때만 st.rerun()으로 전체를 갱신한다.
    _mine = {f"{it['name']}|{it.get('spec','')}" for it in source_items}
    _pc_mine = {k: v for k, v in st.session_state["pending_crew"].items() if k in _mine}
    _pr_mine = {k: v for k, v in st.session_state["pending_rate"].items() if k in _mine}
    _px_mine = {k: v for k, v in st.session_state["pending_exclude"].items() if k in _mine}
    _n_mine = len(_pc_mine) + len(_pr_mine) + len(_px_mine)
    if _n_mine:
        _b1, _b2 = st.columns([3, 1])
        with _b1:
            st.caption(f"✏️ 이 표에서 수정한 {_n_mine}건이 아직 반영되지 않았습니다. 다 고친 뒤 오른쪽 버튼을 누르세요.")
        with _b2:
            if st.button(f"🔄 {_n_mine}건 반영", key=f"apply_{key}", type="primary", width="stretch"):
                st.session_state.setdefault("crew_by_item", {})
                st.session_state.setdefault("manual_rates", {})
                for _ik, (_v, _base) in _pc_mine.items():
                    if _v == _base:
                        st.session_state["crew_by_item"].pop(_ik, None)
                    else:
                        st.session_state["crew_by_item"][_ik] = _v
                    st.session_state["pending_crew"].pop(_ik, None)
                for _ik, (_v, _u) in _pr_mine.items():
                    if _v > 0:
                        st.session_state["manual_rates"][_ik] = {"daily": _v, "unit": _u}
                    else:
                        st.session_state["manual_rates"].pop(_ik, None)
                    st.session_state["pending_rate"].pop(_ik, None)
                for _ik, _v in _px_mine.items():
                    if _v:
                        st.session_state["excluded_items"].add(_ik)
                    else:
                        st.session_state["excluded_items"].discard(_ik)
                    st.session_state["pending_exclude"].pop(_ik, None)
                st.rerun()


def _render_apply_rates(scope_key: str):
    """대기 중인 1일작업량 수정을 한 번에 반영하는 버튼.

    수동입력 관리 탭의 검토 섹션용. 값을 입력할 때마다 즉시 반영하면
    입력 하나마다 전체 재계산이 돌아 느리고, 반대로 rerun을 안 하면
    결과표·총공사기간이 갱신되지 않아 "저장해도 안 바뀐다"처럼 보인다.
    그래서 대기 목록에 모았다가 버튼을 눌렀을 때만 반영·재계산한다.
    """
    _pending = st.session_state.get("pending_rate", {})
    if not _pending:
        return
    _n = len(_pending)
    _a1, _a2 = st.columns([3, 1])
    with _a1:
        st.warning(f"✏️ 수정한 {_n}건이 아직 반영되지 않았습니다. 다 입력한 뒤 오른쪽 버튼을 누르세요.")
    with _a2:
        if st.button(f"🔄 {_n}건 반영", key=f"apply_rates_{scope_key}",
                     type="primary", width="stretch"):
            st.session_state.setdefault("manual_rates", {})
            for _ik, (_v, _u) in _pending.items():
                if _v > 0:
                    st.session_state["manual_rates"][_ik] = {"daily": _v, "unit": _u}
                else:
                    st.session_state["manual_rates"].pop(_ik, None)
            st.session_state["pending_rate"] = {}
            st.rerun()


def is_non_work_category(cat_name: str) -> bool:
    """대공종명이 시공이 아닌(비용·서류·자재) 성격이면 True."""
    return any(kw in (cat_name or "") for kw in NON_WORK_CATEGORY_KEYWORDS)


def naeyeok_to_hierarchy(hier_items):
    """
    naeyeok_parser.parse_naeyeok_hierarchy_items()의 평평한(flat) 리스트를
    기존 app.py가 기대하는 hierarchy 딕셔너리 구조로 변환하는 어댑터.

    기존 구조({'level','name','items','sub_categories'})와 동일한 모양으로
    맞춰서, 결과표/주공정선택/수동입력 등 하위 로직을 신양식·구양식 모두
    무수정으로 재사용할 수 있게 한다.

    level은 "90.N.1" 형식의 합성 코드를 씀 (구양식의 "1.1"/"1.2"/"2.1" 등과
    절대 겹치지 않도록 90번대를 예약 — 크루설정 탭의 major_names 하드코딩
    라벨과 우연히 매칭되어 엉뚱한 이름이 붙는 걸 방지).
    같은 대공종(예: 토공)은 라인(A-LINE 등)이 달라도 같은 major_key로 묶여
    최종적으로 "같은 공종명끼리 합산" 단계에서 하나로 합쳐진다 — 구양식이
    지구(Ⅰ/Ⅱ/Ⅲ)별로 반복되는 같은 공종을 합산하던 것과 동일한 동작.
    """
    major_order = {}
    cats = {}
    _split = bool(st.session_state.get("split_by_parent", False))
    for it in hier_items:
        cat_name = it.get("category") or "기타"
        # 상위 사업 구분 분리 옵션:
        # 한 내역서에 처리장과 관로가 같이 있으면 같은 '토공'이라도 굴착 조건·투입조수가
        # 달라 따로 산정해야 한다(실측: 3_토목도급의 토공이 간선관로/지선관로/옥외배수관/
        # 옥내배수관 4개로 나뉘어 있는데 하나로 병합됨). 켜면 '간선관로 > 토공'처럼 분리한다.
        if _split and it.get("parent"):
            cat_name = f"{it['parent']} > {cat_name}"
        if cat_name not in major_order:
            major_order[cat_name] = len(major_order) + 1
        idx = major_order[cat_name]
        if cat_name not in cats:
            cats[cat_name] = {
                "level": f"90.{idx}.1",
                "name": cat_name,
                "items": [],
                "sub_categories": [
                    # 항목을 합성 sub_category 하나로 감싼다.
                    # 이유: 결과표의 세부항목 렌더링과 '매칭 안 됨' 수집(unmatched_all)이
                    # sub_categories 경로에만 구현되어 있어서, items 직속으로 두면
                    # 수동입력(TAB6)에 항목이 아예 안 뜬다.
                    {
                        "level": "1)",
                        "name": cat_name,
                        "items": [],
                        "sub_categories": [],
                        "district": None,
                    }
                ],
            }
        cats[cat_name]["sub_categories"][0]["items"].append({
            "name": it.get("name", ""),
            "spec": it.get("spec", ""),
            "qty": it.get("qty", 0),
            "unit": it.get("unit", ""),
            "district": it.get("line", ""),
        })
    return list(cats.values())


def calc_days_priority(name, spec, qty, crews=DEFAULT_CREW, item_unit=""):
    """
    우선순위:
    0. 수동입력
    1. 단가산출근거 호표 Q값 (호표번호 1:1 매칭) / 일위대가_산근 코드 Q값 (신양식, 코드 1:1 매칭)
    2. 가이드라인 부록
    3. 표준품셈 Man-day
    4. 단가산출근거 Q값 (항목명 기반)
    
    item_unit: 실제 항목의 단위 (M, 개소, 본 등)
    crews: 대공종 단위 투입조수. 단, 항목별로 따로 지정된 조수가 있으면 그 값이 우선한다.
    """
    if not qty or qty <= 0:
        return 0, "-", "-"

    def _unit_ok(rate_unit, it_unit):
        """1일작업량 단위와 항목 수량 단위가 호환되는지 확인.

        서로 다른 단위끼리 나누면(예: 82m ÷ 1.8개소) 의미 없는 일수가 나온다.
        표기 차이(m/M/㎥/M3 등)는 정규화해서 같은 것으로 본다.
        한쪽이 비어 있으면 판정하지 않고 통과시킨다.
        """
        def norm(u):
            u = str(u or "").strip().upper()
            u = u.replace("㎡", "M2").replace("㎥", "M3").replace("M³", "M3").replace("M²", "M2")
            return u
        a, b = norm(rate_unit), norm(it_unit)
        if not a or not b:
            return True
        return a == b

    # 가설건축물(현장사무소·창고·숙소 등)은 착공 전 준비기간에 설치하므로
    # 본공정 작업일수에 넣지 않는다. 넣으면 230㎡ 사무소가 87일씩 잡혀
    # 부대공 공기를 크게 부풀린다(실측 확인).
    # 준비기간은 비작업일수 탭의 '준비기간(개월)'로 별도 계상한다.
    try:
        _nm_t = name or ""
        if any(k in _nm_t for k in ("가설사무소", "가설건축물", "현장사무소", "가설사무실", "현장사무실")):
            return 0, "준비기간", "준비기간 포함"
    except Exception:
        pass

    # 사용자가 공기에서 빼기로 체크한 항목은 0일 처리.
    # 병행 작업이거나 공기에 영향이 없다고 설계자가 판단한 경우에 쓴다.
    try:
        if f"{name}|{spec or ''}" in st.session_state.get("excluded_items", set()):
            return 0, "제외", "사용자 제외"
    except Exception:
        pass

    # 항목별 조수 우선 적용.
    # 같은 대공종 안에서도 장비·인원이 다른 항목이 있어(예: 대구경 추진 vs 소구경 부설)
    # 상세표에서 항목마다 조수를 조정할 수 있게 한다. 여기 한 곳에서 처리하면
    # 총계·소계·상세 등 모든 호출부에 동일하게 반영된다.
    try:
        _cm = st.session_state.get("crew_by_item", {})
        _cv = _cm.get(f"{name}|{spec or ''}")
        if _cv:
            crews = int(_cv)
    except Exception:
        pass

    # 운반류는 본공정(토공·관부설 등)과 병행되므로 공기에서 제외 (토글 ON일 때).
    # 한 곳에서 처리 → 총계·sub합·상세 모든 호출부가 자동으로 동일하게 0일/병행 처리됨.
    #
    # 단, '상차'는 제외하지 않는다: 백호가 굴착을 멈추고 직접 싣는 작업이라
    # 그 시간만큼 본공정이 지연되므로 공기에 반영해야 한다(실무 확인).
    # 반면 덤프가 현장을 떠나 이동하는 운반·소운반·사토는 그동안 백호가 계속
    # 굴착할 수 있어 병행으로 본다.
    # '폐기물상차(적치장)'처럼 상차와 운반이 한 이름에 섞인 경우도 상차 작업이
    # 포함되어 있으므로 '상차'가 들어가면 제외하지 않는다(상차 우선 판정).
    try:
        _nm_h = name or ""
        _sp_h = spec or ""
        # 상차 판정: 항목명뿐 아니라 규격도 본다.
        # 실측(풍각): '흙운반(적치:토사) | 굴착기 직상차+덤프 운반:L=2.0km'처럼
        # 상차 작업이 규격에만 표기된 경우가 있어 이름만 보면 통째로 제외돼 버린다.
        # 단, '하치장 상차도(VAT포함)'는 자재 인도조건(상차도)이지 상차 작업이 아니므로 제외한다.
        _has_loading = ("상차" in _nm_h) or ("상차" in _sp_h and "상차도" not in _sp_h)
        if st.session_state.get("exclude_haul", False) and not _has_loading and any(
            k in _nm_h for k in ("운반", "사토", "잔토")
        ):
            return 0, "병행", "병행(제외)"
    except Exception:
        pass

    # 0순위: 수동입력 (최우선)
    try:
        if "manual_rates" in st.session_state:
            manual_key = f"{name}|{spec}"  # 이름+규격으로 키 생성
            if manual_key in st.session_state["manual_rates"]:
                manual_data = st.session_state["manual_rates"][manual_key]
                daily_val = manual_data.get("daily", 0)
                unit = manual_data.get("unit", "")
                if daily_val > 0:
                    days = math.ceil(qty / (daily_val * crews))
                    return days, f"{daily_val:.1f}{unit}", "수동입력"
    except Exception:
        pass

    # 0.5순위: 실무 기준값 override
    # 배수설비(가옥·공장 인입)는 일위대가 1개소 안에 관로+토공+포장이 모두 묶여 있어
    # 노무비 역산 시 0.03~0.08개소/일(1개소당 13~14인·일)로 잡히지만,
    # 실제 현장은 '1개조가 하루에 1가구(1개소)'를 완료한다. 실측 기준을 우선 적용한다.
    try:
        _nm_s = name or ""
        # 단위 인자가 항상 전달되지는 않으므로(집계 호출부는 생략) 항목명으로만 판정한다.
        # '연막시험' 등 검사 성격 항목은 제외.
        if "배수설비" in _nm_s and not any(k in _nm_s for k in ("시험", "검사", "조사")):
            days = math.ceil(qty / (1.0 * crews))
            return days, "1.0개소", "실무기준(1조=1개소/일)"
    except Exception:
        pass

    # 1순위: 단가산출근거 호표 Q값 (호표번호로 1:1 매칭; 형식1=Q=/HR 기반)
    try:
        hopyo_map = st.session_state.get("hopyo_by_item", {})
        hopyo_daily = st.session_state.get("hopyo_daily", {})
        hopyo_no = hopyo_map.get((name, spec))
        if hopyo_no is not None and hopyo_no in hopyo_daily:
            daily_val, unit = hopyo_daily[hopyo_no]
            if daily_val and daily_val > 0:
                days = math.ceil(qty / (daily_val * crews))
                return days, f"{daily_val:.1f}{unit}", "단가산출근거(호표)"
    except Exception:
        pass

    # 1순위: 일위대가_산근 코드 Q값 (신양식 — 코드로 1:1 매칭; 호표 방식과 동일 우선순위)
    try:
        naeyeok_map = st.session_state.get("naeyeok_code_by_item", {})
        naeyeok_daily = st.session_state.get("naeyeok_code_daily", {})
        naeyeok_code = naeyeok_map.get((name, spec))
        if naeyeok_code is not None and naeyeok_code in naeyeok_daily:
            daily_val, unit = naeyeok_daily[naeyeok_code]
            if daily_val and daily_val > 0:
                days = math.ceil(qty / (daily_val * crews))
                return days, f"{daily_val:.1f}{unit}", "일위대가_산근(코드)"
    except Exception:
        pass

    # 1.4순위: 동일 계열 Q값 승계 (기계 위주 항목의 역산 과소평가 보정)
    try:
        _fb = st.session_state.get("naeyeok_series_fallback", {})
        _km = st.session_state.get("naeyeok_labor_key_by_item", {})
        _cm = st.session_state.get("naeyeok_code_by_item", {})
        _strict = st.session_state.get("labor_key_strict", False)
        _key = _km.get((name, spec)) if _strict else _km.get((name, spec), _cm.get((name, spec)))
        if _key is not None and _key in _fb:
            daily_val, unit = _fb[_key]
            if daily_val and daily_val > 0 and _unit_ok(unit, item_unit):
                days = math.ceil(qty / (daily_val * crews))
                return days, f"{daily_val:.1f}{unit}", "동일계열 Q승계"
    except Exception:
        pass

    # 1.44순위: 표준품셈 작업조 기준 (품셈 원문에서 확인한 공종)
    # 품셈 1-2-8: "작업조는 일당시공량을 시공하기 위한 필수자원의 조합"
    # → 여기 값은 '1개 작업조'가 하루에 하는 양이고, 앱의 투입조수 1이 그 1세트다.
    # 노무비 역산(1인 기준)보다 근거가 확실하므로 우선 적용한다.
    try:
        if HAS_PUMSEM_CREW:
            _nm_p = name or ""
            _sp_p = spec or ""
            _txt_p = f"{_nm_p} {_sp_p}"

            # 흄관(철근콘크리트관) 부설·접합 / 철거 — 관경별
            if any(k in _nm_p for k in ("흄관", "원심력", "철근콘크리트관")):
                _md = re.search(r"[DØΦ]?\s*(\d{3,4})\s*(?:㎜|mm)?", _sp_p)
                if _md:
                    _dia = int(_md.group(1))
                    _is_rm = any(k in _nm_p for k in ("철거", "제거"))
                    _r = hume_pipe_rate(_dia, removal=_is_rm)
                    if _r and _unit_ok("본", item_unit):
                        _dv, _pers = _r
                        days = math.ceil(qty / (_dv * crews))
                        return days, f"{_dv:g}본", f"표준품셈 조기준({_pers}인/조)"

            # 시스템 동바리 설치 및 해체 — 높이·설치간격별
            if "시스템" in _nm_p and "동바리" in _nm_p:
                # 내역서 표기는 '5m 초과∼10m 이하'처럼 공백과 물결표(∼)가 섞여 있어
                # 공백 제거 + 물결표 통일 후 비교한다.
                _norm_p = re.sub(r"\s+", "", _txt_p).replace("∼", "~")
                _h = None
                for _lbl in ("10m초과~20m이하", "20m초과~30m이하", "5m초과~10m이하", "5m이하"):
                    if _lbl in _norm_p:
                        _h = _lbl
                        break
                _sp_lbl = None
                for _lbl in ("0.6m초과~0.8m이하", "0.6m이하", "0.8m초과"):
                    if _lbl in _norm_p:
                        _sp_lbl = _lbl
                        break
                if _h:
                    _r = system_support_rate(_h, _sp_lbl)
                    if _r and _unit_ok("공㎥", item_unit):
                        _dv, _pers = _r
                        days = math.ceil(qty / (_dv * crews))
                        return days, f"{_dv:g}공㎥", f"표준품셈 조기준({_pers}인/조)"
    except Exception:
        pass

    # 1.45순위: 표준품셈 테이블 (관부설·되메우기·포장·맨홀·추진 등 주요 공종)
    # 흄관 접합·부설처럼 표준품셈에 관경별 1일 시공량이 표로 있는 항목은
    # 노무비 역산보다 이 값을 우선한다. 역산은 노무비를 1인 단가로 나눠
    # '1인 기준'이 나오는데, 품셈 표는 '1개 조' 기준이라 4배가량 차이가 났다
    # (실측: 흄관 D300 역산 4.06본 vs 품셈 16본/조).
    # 앱의 투입조수 1 = 품셈 표의 1개 조(세트)로 본다.
    try:
        if HAS_DAILY_WORK:
            _t = f"{name} {spec or ''}"
            # 적용 제외: 철거·하차비 등은 부설 품셈과 성격이 달라 제외한다
            # (실측: '흄관하차비', '콘크리트관철거(흄관)'가 부설 품셈을 가져가는 오탐 발생).
            # 부설 품셈이 엉뚱한 작업에 붙는 것을 막는다.
            # '절단'은 가이드라인에 별도 값(콘크리트포장절단 500m 등)이 있어
            # 부설 품셈을 쓰면 크게 과소평가된다(실측: 2.3개소/일).
            _skip = any(k in name for k in
                        ("철거", "하차", "제거", "폐기", "운반", "상차", "절단", "청소", "시험"))
            for _kws, _dias, _key in ([] if _skip else WORK_KEY_MAP):
                if not any(k in _t for k in _kws):
                    continue
                if _dias and not any(d in _t for d in _dias):
                    continue
                _info = DAILY_WORK.get(_key)
                if not _info:
                    break
                _dv = _info.get("daily")
                _du = _info.get("unit", "")
                # DAILY_WORK의 daily 값은 '1개 조'의 하루 생산량이다.
                # crews 필드는 그 조를 이루는 인원수(예: 포장 5인, 배관 3인)이지
                # '몇 개 조'가 아니므로 값을 나누면 안 된다.
                # (검증: 아스팔트포장 600㎡를 5로 나누면 120㎡가 되어
                #  실측 Q산식 4,800㎡의 1/40이 되어버림)
                if _dv and _dv > 0 and _unit_ok(_du, item_unit):
                    days = math.ceil(qty / (_dv * crews))
                    return days, f"{_dv:g}{_du}", "표준품셈"
                break
    except Exception:
        pass

    # 1.5순위: 일위대가 노무비 역산 (Q산식이 없는 블록 — 설계자 직접산정값보다는 신뢰도 낮음)
    try:
        labor_key_map = st.session_state.get("naeyeok_labor_key_by_item", {})
        naeyeok_map = st.session_state.get("naeyeok_code_by_item", {})
        labor_daily = st.session_state.get("naeyeok_labor_daily", {})
        _strict2 = st.session_state.get("labor_key_strict", False)
        naeyeok_code = (labor_key_map.get((name, spec)) if _strict2
                        else labor_key_map.get((name, spec), naeyeok_map.get((name, spec))))
        if naeyeok_code is not None and naeyeok_code in labor_daily:
            daily_val, unit, job = labor_daily[naeyeok_code]
            if daily_val and daily_val > 0 and _unit_ok(unit, item_unit):
                days = math.ceil(qty / (daily_val * crews))
                return days, f"{daily_val:.1f}{unit}", f"노무비역산({job})"
    except Exception:
        pass

    # 1순위: 가이드라인
    try:
        # 정확한 매칭 시도
        full_name = f"{name} {spec}".strip()
        
        # GUIDELINE_APPENDIX_FULL 우선 사용 (확장판)
        guideline_data = GUIDELINE_APPENDIX_FULL if GUIDELINE_APPENDIX_FULL else GUIDELINE_APPENDIX
        
        # 띄어쓰기 제거 및 괄호 제거한 버전 준비
        full_name_no_space = full_name.replace(" ", "").replace("(", "").replace(")", "")
        name_no_space = name.replace(" ", "").replace("(", "").replace(")", "")
        
        for key, val in guideline_data.items():
            matched = False
            key_no_space = key.replace(" ", "").replace("(", "").replace(")", "")
            
            # 매칭 조건 (우선순위)
            # 1. 정확한 전체 매칭 (띄어쓰기/괄호 무시)
            if key_no_space == full_name_no_space or key_no_space == name_no_space:
                matched = True
            
            # 2. 가이드라인 키가 항목명에 포함 (띄어쓰기/괄호 무시)
            elif key_no_space in full_name_no_space or key_no_space in name_no_space:
                matched = True
            
            # 3. 항목명이 가이드라인 키에 포함 (띄어쓰기/괄호 무시)
            elif name_no_space in key_no_space:
                matched = True
            
            # 4. 원본 문자열 매칭 (띄어쓰기 있는 버전)
            elif key == full_name or key == name or key in full_name or key in name:
                matched = True
            
            # 5. 핵심 키워드 매칭 (특수 케이스)
            # "조립식맨홀설치" → "조립식맨홀" 매칭
            elif "조립식맨홀" in key and "조립식맨홀" in name:
                matched = True
            elif "맨홀뚜껑" in key and "맨홀뚜껑" in name:
                matched = True
            # 🔥 수정: 추진공은 "강관" + "추진" 조합만 매칭 (오매칭 방지)
            # 원래 코드: elif "추진" in key and "추진" in name and len(key) > 2:
            # 문제: "추진" 키워드만으로 다른 항목과 오매칭되어 1850일 폭발
            elif "추진" in key and "추진" in name and "강관" in key and "강관" in (name + spec):
                matched = True
            
            if matched:
                base_daily = val.get("daily", 0)
                unit = val.get("unit", "")
                
                # ⚠️ 의심스러운 가이드라인 스킵 (이상 값 방지)
                # 1. unit="일" 같이 명확하지 않은 단위
                # 2. daily가 1 미만으로 너무 작은 값 (1m/일 미만은 비정상)
                unit_clean = unit.strip().lower()
                if unit_clean in ["일", "day", "days", ""]:
                    # 단위가 "일"이면 의미 불명확 → 매칭 안 됨으로 처리
                    continue
                if base_daily > 0 and base_daily < 1:
                    # 1일에 1단위 미만은 너무 작음 → 매칭 안 됨으로 처리
                    continue
                
                # ⚠️ 단위 불일치 체크
                if item_unit and unit and base_daily > 0:
                    # 단위 정규화
                    item_unit_clean = item_unit.strip().lower().replace(" ", "")
                    guideline_unit_clean = unit.split("/")[0].strip().lower().replace(" ", "")
                    
                    # M/일 vs 개소, 본/일 vs M 같은 불일치 감지
                    unit_mismatch = False
                    if "m" in guideline_unit_clean or "ｍ" in guideline_unit_clean:
                        if item_unit_clean not in ["m", "ｍ", "m3", "㎥"]:
                            unit_mismatch = True
                    elif "본" in guideline_unit_clean:
                        if item_unit_clean not in ["본", "ea", "개"]:
                            unit_mismatch = True
                    elif "개소" in guideline_unit_clean or "ea" in guideline_unit_clean:
                        if item_unit_clean not in ["개소", "ea", "개", "본"]:
                            unit_mismatch = True
                    
                    if unit_mismatch:
                        # 단위 불일치 → 이 가이드라인은 스킵
                        continue
                
                if base_daily > 0:
                    if is_machine_based(name):
                        days = math.ceil(qty / (base_daily * crews))
                        label = f"{base_daily}{unit}"  # 조수 제거
                    else:
                        days = math.ceil(qty / (base_daily * crews))
                        label = f"{base_daily}{unit}"  # 조수 제거
                    return days, label, "가이드라인"
        
        # 관부설 직경별 매칭
        if any(kw in name for kw in ["관 부설","관부설","고강성PVC","PE다중벽","이중벽관","주철관","GRP관"]):
            dia = extract_diameter(spec)
            if dia:
                pipe_rates = {200:5, 300:4, 450:3, 600:2.5, 800:2, 1000:1.5, 1200:1.2}
                closest = min(pipe_rates.keys(), key=lambda x: abs(x - dia))
                daily = pipe_rates[closest]
                days = math.ceil(qty / (daily * crews))
                return days, f"{daily}본/일", "가이드라인"  # 조수 제거
    except Exception:
        pass

    # 2순위: 표준품셈
    try:
        manday = 0
        if any(kw in name for kw in ["터파기","굴착","줄파기"]) and "운반" not in name:
            info = get_excavation_labor(spec)
            rate = info.get("인/m3")
            if rate:
                manday = rate * qty

        pipe_kws = ["관 부설","관부설","이중벽관","주철관","흄관","콘크리트관",
                    "GRP관","유리섬유복합관","파형강관","PE다중벽","고강성PVC","강관부설"]
        if any(kw in name for kw in pipe_kws) and not manday:
            dia = extract_diameter(spec)
            if dia:
                info = get_pipe_labor(dia)
                rate = info.get("합계")
                if rate:
                    manday = rate * qty

        if manday > 0:
            days = math.ceil(manday / (8 * crews))
            return days, f"{round(manday/qty,3)}인/단위", "표준품셈"  # 조수 제거
    except Exception:
        pass
    
    # 3순위: 단가산출근거
    try:
        if "dangagun_cache" in st.session_state:
            cache = st.session_state["dangagun_cache"]
            
            # 항목명 + 규격으로 매칭 시도
            full_name = f"{name} {spec}".strip()
            
            for cached_name, info in cache.items():
                # 정확한 매칭 우선
                if cached_name == full_name or cached_name in full_name or full_name in cached_name:
                    # hourly 값 (시간당)
                    if "hourly" in info:
                        hourly_val = info.get("hourly", 0)
                        unit = info.get("unit", "")
                        if hourly_val > 0:
                            daily_val = hourly_val * 8
                            days = math.ceil(qty / (daily_val * crews))
                            return days, f"{daily_val:.1f}{unit.replace('/Hr','/일')}", "단가산출근거"  # 조수 제거
                    
                    # daily 값 (1일 작업량)
                    elif "daily" in info:
                        daily_val = info.get("daily", 0)
                        unit = info.get("unit", "")
                        if daily_val > 0:
                            days = math.ceil(qty / (daily_val * crews))
                            return days, f"{daily_val:.1f}{unit}", "단가산출근거"  # 조수 제거
                
                # 항목명만으로도 매칭 시도
                if name in cached_name or cached_name in name:
                    if "hourly" in info:
                        hourly_val = info.get("hourly", 0)
                        unit = info.get("unit", "")
                        if hourly_val > 0:
                            daily_val = hourly_val * 8
                            days = math.ceil(qty / (daily_val * crews))
                            return days, f"{daily_val:.1f}{unit.replace('/Hr','/일')}", "단가산출근거"  # 조수 제거
                    
                    elif "daily" in info:
                        daily_val = info.get("daily", 0)
                        unit = info.get("unit", "")
                        if daily_val > 0:
                            days = math.ceil(qty / (daily_val * crews))
                            return days, f"{daily_val:.1f}{unit}", "단가산출근거"  # 조수 제거
    except Exception:
        pass

    # 4순위: 매칭 안 됨 → 수동입력 필요
    return 0, "⚠️ 수동입력", "매칭 안 됨"

# ══════════════════════════════════════════════════════════════
# 비작업일수
# ══════════════════════════════════════════════════════════════
HOLIDAYS_DB = {
    2025:{1:8,2:4,3:7,4:4,5:6,6:6,7:4,8:6,9:4,10:9,11:5,12:5},
    2026:{1:5,2:7,3:6,4:4,5:7,6:5,7:4,8:7,9:7,10:7,11:5,12:5},
}
RAIN = {1:0,2:1,3:2,4:3,5:4,6:6,7:8,8:7,9:5,10:3,11:2,12:1}

def get_kr_holidays(year):
    m = HOLIDAYS_DB.get(year, {})
    holidays = set()
    for month, count in m.items():
        for day in range(1, count + 1):
            try:
                holidays.add(datetime(year, month, day).date())
            except:
                pass
    return holidays

def calc_completion_date(start, work_days):
    current, worked = start, 0
    kr_holidays = get_kr_holidays(start.year) | get_kr_holidays(start.year + 1)
    while worked < work_days:
        if current.weekday() == 6 or current in kr_holidays or current.day % 30 < RAIN[current.month]:
            current += timedelta(days=1)
            continue
        worked += 1
        current += timedelta(days=1)
    return current - timedelta(days=1)

# ══════════════════════════════════════════════════════════════
# 엑셀 파서
# ══════════════════════════════════════════════════════════════
def parse_by_keyword(file):
    # 🔥 디버그: 파일 로그
    import datetime
    log_file = "debug_log.txt"
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(f"\n{'='*60}\n")
        f.write(f"파싱 시작: {datetime.datetime.now()}\n")
        f.write(f"{'='*60}\n")
    
    # 🔥 디버그: 파싱 시작
    print(f"\n{'🔥'*30}")
    print(f"📂 parse_by_keyword 시작")
    print(f"{'🔥'*30}\n")
    
    wb = openpyxl.load_workbook(file, data_only=True)  # read_only=False로 변경
    skip_sheets = ["목차","안내","INITIAL","초기","index"]
    priority = ["설계내역서","내역서","공사비내역서"]
    target_sheet = None
    
    for p in priority:
        if p in wb.sheetnames:
            target_sheet = p
            break
    if not target_sheet:
        for sname in wb.sheetnames:
            if any(sk in sname for sk in skip_sheets):
                continue
            if "내역" in sname:
                target_sheet = sname
                break
    if not target_sheet:
        for sname in wb.sheetnames:
            if not any(sk in sname for sk in skip_sheets):
                target_sheet = sname
                break
    if not target_sheet:
        target_sheet = wb.sheetnames[0]

    ws = wb[target_sheet]
    col_info = {"시트명": target_sheet}
    header_row = None
    
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=30, values_only=True), 1):
        row_str = " ".join([str(c) for c in row if c])
        if any(k in row_str for k in ["공종","품명","세부품명","명칭","내역"]):
            header_row = row_idx
            break
    
    if not header_row:
        header_row = 1
    
    col_info["헤더행"] = header_row
    
    # ══════════════════════════════════════════════════════════════
    # 지구 경계 찾기 (로마숫자)
    # ══════════════════════════════════════════════════════════════
    districts = {}
    roman_nums = ['Ⅰ', 'Ⅱ', 'Ⅲ', 'Ⅳ', 'Ⅴ', 'Ⅵ', 'Ⅶ', 'Ⅷ', 'Ⅸ', 'Ⅹ']
    
    all_rows_raw = list(ws.iter_rows(min_row=1, values_only=False))
    
    for row_idx, row in enumerate(all_rows_raw):
        a_val = str(row[0].value or "").strip()
        b_val = str(row[1].value or "").strip() if len(row) > 1 else ""
        
        if a_val in roman_nums:
            districts[a_val] = {
                'name': b_val,
                'start_row': row_idx,
                'end_row': None
            }
    
    # 지구별 end_row 설정
    district_keys = sorted(districts.keys(), key=lambda x: roman_nums.index(x))
    for i, key in enumerate(district_keys):
        if i + 1 < len(district_keys):
            next_key = district_keys[i + 1]
            districts[key]['end_row'] = districts[next_key]['start_row'] - 1
        else:
            districts[key]['end_row'] = len(all_rows_raw) - 1
    
    col_info["districts"] = districts
    
    # ══════════════════════════════════════════════════════════════
    # 지구별 데이터 파싱
    # ══════════════════════════════════════════════════════════════
    results = []
    
    for district, info in districts.items():
        district_rows = all_rows_raw[info['start_row']:info['end_row']+1]
        
        for local_idx, row in enumerate(district_rows):
            row_idx = info['start_row'] + local_idx
            
            # values_only=False이므로 .value 접근
            gong_jong_val = row[0].value
            name_val = row[1].value if len(row) > 1 else None
            spec_val = row[2].value if len(row) > 2 else None
            qty_val = row[3].value if len(row) > 3 else None
            unit_val = row[4].value if len(row) > 4 else None
            
            gong_jong = str(gong_jong_val).strip() if gong_jong_val else ""
            name = str(name_val).strip() if name_val else ""
            spec = str(spec_val).strip() if spec_val else ""
            unit = str(unit_val).strip() if unit_val else ""
            
            if not name or any(skip in name for skip in SKIP_NAMES):
                continue
            
            try:
                qty = float(qty_val) if qty_val else 0
            except:
                qty = 0
            
            if qty <= 0:
                continue
            
            group = "기타"
            for grp, keywords in KEYWORD_MAP_DETAIL.items():
                if any(kw in name for kw in keywords):
                    group = grp
                    break
            
            if group == "관부설공" and any(ex in name for ex in PIPE_EXCLUDE):
                group = "기타"
            
            detail_spec = spec
            if not detail_spec and name:
                spec_match = re.search(r'\([^)]+\)', name)
                if spec_match:
                    detail_spec = spec_match.group(0)
            
            # 호표 참조 추출 (행 전체 스캔; 통상 비고열에 '산근 N호표' 텍스트가 있음)
            hopyo_num = None
            for cell in row:
                v = cell.value
                if isinstance(v, str):
                    m_ref = re.search(r'산근\s*(\d+)\s*호표', v)
                    if m_ref:
                        hopyo_num = int(m_ref.group(1))
                        break

            results.append({
                "row_idx": row_idx,
                "district": district,
                "district_name": info['name'],
                "gong_jong": gong_jong,
                "group": group,
                "name": name,
                "spec": detail_spec,
                "qty": qty,
                "unit": unit,
                "amount": row[5].value if len(row) > 5 else 0,
                "labor": row[6].value if len(row) > 6 else 0,
                "hopyo": hopyo_num,
            })
    
    # 원본 순서로 정렬 (row_idx 기준)
    results.sort(key=lambda x: x["row_idx"])
    
    # 중복 제거 (전체 데이터 기준)
    merged = {}
    for r in results:
        key = (r["name"], r["spec"])
        if key not in merged:
            merged[key] = dict(r)
        else:
            merged[key]["qty"] = (merged[key].get("qty") or 0) + (r.get("qty") or 0)
            merged[key]["amount"] = (merged[key].get("amount") or 0) + (r.get("amount") or 0)
            merged[key]["labor"] = (merged[key].get("labor") or 0) + (r.get("labor") or 0)
            # 호표는 같은 (name, spec)이면 동일하다고 가정. 만일 첫 행에 비어있고 뒤에 채워졌다면 채워준다.
            if merged[key].get("hopyo") is None and r.get("hopyo") is not None:
                merged[key]["hopyo"] = r["hopyo"]
    
    return list(merged.values()), col_info


def _extract_dangagun(wb):
    """단가산출근거 시트에서 항목별 Q값(시간당·1일 작업량)을 추출한다."""
    dangagun_cache = {}
    if '단가산출근거' not in wb.sheetnames:
        return dangagun_cache
    ws_danga = wb['단가산출근거']
    current_item = None

    for row in ws_danga.iter_rows(min_row=1, values_only=True):
        row_text = " ".join([str(c) for c in row if c])

        # 항목명 추출 (규격 포함)
        if row[1] and "/" in str(row[1]):
            item_text = str(row[1]).strip()
            if "/" in item_text:
                current_item = item_text.split("/")[0].strip()

        # Q 값 추출 (다양한 패턴)
        if current_item and "Q =" in row_text:
            # 패턴 1: Q = 숫자 단위/HR
            match1 = re.search(r'Q\s*=\s*([\d.]+)\s*([^\s]+/HR)', row_text, re.IGNORECASE)
            if match1:
                hourly_val = float(match1.group(1))
                unit = match1.group(2).replace("HR", "Hr").replace("hr", "Hr")
                dangagun_cache[current_item] = {"hourly": hourly_val, "unit": unit}
                continue

            # 패턴 2: Q = 숫자/일 /8 Hr = 숫자 단위/Hr
            match2 = re.search(r'=\s*([\d.]+)\s*([^\s/]+)/Hr', row_text, re.IGNORECASE)
            if match2:
                hourly_val = float(match2.group(1))
                unit = match2.group(2) + "/Hr"
                dangagun_cache[current_item] = {"hourly": hourly_val, "unit": unit}
                continue

        # 1세트 = N일 패턴
        if current_item and "세트" in row_text and "일" in row_text:
            match3 = re.search(r'(\d+)\s*세트\s*=\s*([\d.]+)\s*일', row_text)
            if match3:
                sets = float(match3.group(1))
                days = float(match3.group(2))
                # 1일 = sets/days 세트
                daily_val = sets / days
                dangagun_cache[current_item] = {"daily": daily_val, "unit": "세트/일"}
                continue
    return dangagun_cache


@st.cache_data(show_spinner=False, max_entries=3)
def parse_workbook_cached(file_bytes: bytes):
    """업로드한 엑셀을 한 번만 파싱하고, 결과를 파일 내용 기준으로 캐시한다.

    Streamlit은 위젯을 건드릴 때마다 스크립트 전체를 다시 실행한다. 예전에는
    그때마다 엑셀을 3번(양식 판별·키워드 파싱·계층 파싱) 새로 읽어서 재실행 1회에
    수 초씩 걸렸다. 내용이 같은 파일이면 캐시된 결과를 돌려준다.
    st.cache_data는 호출마다 복사본을 주므로 화면 쪽에서 값을 고쳐도 캐시는 그대로다.

    반환: (all_rows, col_info, parsed)
      parsed["dangagun"]: 단가산출근거 Q값
      parsed["template"]: 인식된 템플릿의 파싱 결과 (인식 실패 시 None)
    """
    all_rows, col_info = parse_by_keyword(BytesIO(file_bytes))
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    parsed = {"dangagun": _extract_dangagun(wb), "template": None,
              "sheets": list(wb.sheetnames)}
    tmpl = detect_template(wb)
    if tmpl is not None:
        items = parse_items_generic(wb, tmpl)
        code_daily = parse_unit_price_generic(wb, tmpl)
        labor_daily = parse_labor_derived_by_template(wb, tmpl)
        parsed["template"] = {
            "name": tmpl["name"],
            "has_alt": bool(tmpl.get("link_regex_alt")),
            "items": items,
            "code_daily": code_daily,
            "labor_daily": labor_daily,
            # 기계 위주 항목의 역산 과소평가 보정: 같은 계열 Q산식 값 승계
            "series_fallback": apply_series_fallback(items, code_daily, labor_daily),
        }
    return all_rows, col_info, parsed

# ══════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════
st.sidebar.header("⚙️ 기본 설정")

# 상위 사업 구분으로 대공종 분리 (처리장+관로 혼합 발주 대응)
st.sidebar.checkbox(
    "구조물+관로사업",
    value=False,
    key="split_by_parent",
    help=(
        "하나의 내역서에 처리장·정수장 같은 구조물 사업과 관로 사업이 함께 들어있을 때 켜세요.\n\n"
        "같은 '토공'이라도 구조물은 넓은 부지를 대규모로 굴착하고 관로는 좁고 긴 선형으로 "
        "굴착해서, 조건과 투입조수가 다릅니다. 켜면 '간선관로 > 토공', '옥내배수관 > 토공'처럼 "
        "사업별로 나눠서 조수와 작업일수를 각각 설정할 수 있습니다.\n\n"
        "단일 사업이면 꺼두세요(화면이 간단해집니다). 변경 후에는 엑셀을 다시 업로드하세요."
    ),
)

# 공사 유형 selectbox 제거:
# 양식·계층 구조는 universal_parser가 자동 판별하고, 실제로 구분이 필요한 부분
# (관로 vs 처리장의 조수 차이)은 대공종별 투입조수 설정에서 이미 개별 지정한다.
# 처리장+관로 혼합 발주도 대공종이 분리되어 나오므로 유형 선택이 불필요하다.

# 📋 워크플로우 가이드
st.sidebar.markdown("""
<div style='background: linear-gradient(135deg, #1e3a8a 0%, #1e40af 100%); 
            padding: 16px; border-radius: 10px; margin-bottom: 20px;'>
    <h3 style='color: white; margin: 0 0 12px 0; font-size: 16px;'>📋 작업 순서</h3>
    <ol style='color: #dbeafe; margin: 0; padding-left: 20px; font-size: 13px; line-height: 1.8;'>
        <li>📂 <strong>엑셀 인식</strong><br><span style='font-size: 11px; color: #93c5fd;'>설계내역서 업로드</span></li>
        <li>📝 <strong>수동입력</strong><br><span style='font-size: 11px; color: #93c5fd;'>매칭 안 된 항목 입력</span></li>
        <li>🌧 <strong>비작업일수</strong><br><span style='font-size: 11px; color: #93c5fd;'>기상·휴일 반영</span></li>
        <li>📋 <strong>공기산정</strong><br><span style='font-size: 11px; color: #93c5fd;'>투입조수·작업일수</span></li>
        <li>🔍 <strong>CP 분석</strong><br><span style='font-size: 11px; color: #93c5fd;'>주요공종 식별</span></li>
        <li>📅 <strong>예정공정표</strong><br><span style='font-size: 11px; color: #93c5fd;'>월 단위 공정표 생성</span></li>
    </ol>
</div>
""", unsafe_allow_html=True)

st.sidebar.info("📅 **공사 시작일**은 TAB '비작업일수 계산기'에서 설정")
st.title("상하수도 공사기간 산정 시스템")
st.markdown("---")

tab2, tab6, tab4, tab1, tab3, tab5 = st.tabs([
    "📂 엑셀 내역서 인식",
    "📝 수동입력 관리",
    "🌧 비작업일수 계산기",
    "📋 공기산정",
    "🔍 주요공종 CP 분석",
    "📅 예정공정표"
])

# ══════════════════════════════════════════════════════════════
# TAB 2
# ══════════════════════════════════════════════════════════════
with tab2:
    st.subheader("📂 엑셀 내역서 자동 인식")
    st.caption("도급 설계내역서 업로드 → 계층 구조 자동 파싱")

    uploaded = st.file_uploader(
        "설계내역서 엑셀 (.xlsx 형식만 지원)",
        type=["xlsx"],
        help="구버전 .xls 파일은 지원하지 않습니다. 엑셀에서 열어 '다른 이름으로 저장' → 'Excel 통합 문서(.xlsx)'로 저장한 뒤 업로드해주세요.",
    )

    if uploaded:
        # 프로그레스 바
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            status_text.text("📂 엑셀 파일 로드 중...")
            progress_bar.progress(20)
            
            # 파싱 결과는 파일 내용 기준으로 캐시된다(parse_workbook_cached 참고)
            with st.spinner("파싱 중..."):
                all_rows, col_info, _parsed = parse_workbook_cached(uploaded.getvalue())
            _tmpl_data = _parsed["template"]
            # 템플릿이 인식되면 구양식 전용 게이트(matched)를 우회한다
            _early_is_new = _tmpl_data is not None
            
            progress_bar.progress(40)
            status_text.text("✅ 파싱 완료!")
            
            matched = [r for r in all_rows if r["group"] != "기타"]
            
            progress_bar.progress(60)
            status_text.text("🔍 계층 구조 분석 중...")
            
            if matched or _early_is_new:
                status_text.text("📊 UI 생성 중...")
                progress_bar.progress(80)
                
                st.markdown("---")
                
                # 단가산출근거 Q값 (parse_workbook_cached에서 추출)
                dangagun_cache = _parsed["dangagun"]
                st.session_state["dangagun_cache"] = dangagun_cache
                
                if dangagun_cache:
                    st.info(f"✅ 단가산출근거에서 {len(dangagun_cache)}개 항목 Q값 추출")
                
                # 계층 구조 파싱 (설정 기반 범용 엔진 — universal_parser.py)
                hierarchy = []
                if _tmpl_data is not None:
                    _uni_items = _tmpl_data["items"]
                    hierarchy = naeyeok_to_hierarchy(_uni_items)
                    st.session_state["naeyeok_code_daily"] = _tmpl_data["code_daily"]
                    st.session_state["naeyeok_labor_daily"] = _tmpl_data["labor_daily"]
                    # 기계 위주 항목의 역산 과소평가 보정: 같은 계열 Q산식 값 승계
                    st.session_state["naeyeok_series_fallback"] = _tmpl_data["series_fallback"]
                    st.session_state["naeyeok_code_by_item"] = {
                        (it["name"], it.get("spec", "")): it["code"]
                        for it in _uni_items if it.get("code")
                    }
                    # 노무비 역산용 키.
                    # 주의: 표준형에서 '산근 N호표'(단가산출근거)와 '대가 N호표'(일위대가표)는
                    # 서로 다른 번호 체계다. 산근 번호를 일위대가표 키로 쓰면 엉뚱한 호표에
                    # 매칭된다(실측: 아스팔트포장이 '석축헐기 및 복구' 값을 가져와 0.16㎡/일).
                    # 따라서 link_regex_alt가 있는 템플릿(표준형)에서는 code_alt만 사용하고,
                    # alt 개념이 없는 템플릿(코드매칭형)에서만 code로 폴백한다.
                    _has_alt = _tmpl_data["has_alt"]
                    # 표준형처럼 '산근 N호표'(단가산출근거)와 '대가 N호표'(일위대가표)가
                    # 서로 다른 번호 체계인 양식에서는, 산근 번호를 일위대가 키로 폴백하면
                    # 엉뚱한 호표에 매칭된다(실측: '아스팔트포장 절단 82m'이 산근56을
                    # 일위대가표56 '그레이팅 뚜껑/개소'로 읽어 1.8개소/일이 됨).
                    st.session_state["labor_key_strict"] = _has_alt
                    if _has_alt:
                        st.session_state["naeyeok_labor_key_by_item"] = {
                            (it["name"], it.get("spec", "")): it["code_alt"]
                            for it in _uni_items if it.get("code_alt") is not None
                        }
                    else:
                        st.session_state["naeyeok_labor_key_by_item"] = {
                            (it["name"], it.get("spec", "")): it["code"]
                            for it in _uni_items if it.get("code") is not None
                        }
                    st.session_state["hopyo_daily"] = {}
                    st.session_state["hopyo_by_item"] = {}
                    # 수집 목록은 매 실행마다 다시 채워지므로 항상 비운다.
                    st.session_state["unmatched_all"] = {}
                    st.session_state["series_review_all"] = {}
                    st.session_state["labor_review_all"] = {}
                    st.session_state["crew_targets"] = {}

                    # 사용자가 입력한 값(제외·조수·작업량)은 '파일이 실제로 바뀐 경우'에만 초기화한다.
                    # Streamlit은 버튼을 누를 때마다 스크립트를 처음부터 다시 실행하는데,
                    # 업로더에 파일이 남아 있어 이 블록도 매번 재실행된다.
                    # 무조건 초기화하면 반영 버튼을 누르는 순간 제외·수정이 지워져
                    # "체크해도 안 빠진다"처럼 보인다(실측 확인).
                    _file_sig = f"{getattr(uploaded, 'name', '')}|{len(_uni_items)}"
                    if st.session_state.get("_loaded_file_sig") != _file_sig:
                        st.session_state["_loaded_file_sig"] = _file_sig
                        st.session_state["pending_crew"] = {}
                        st.session_state["pending_rate"] = {}
                        st.session_state["pending_exclude"] = {}
                        st.session_state["excluded_items"] = set()
                        st.session_state["crew_by_item"] = {}
                        st.session_state["manual_rates"] = {}
                    if _uni_items:
                        st.info(f"✅ 템플릿 인식: {_tmpl_data['name']} — {len(_uni_items)}개 항목, {len(hierarchy)}개 대공종")
                        if not _tmpl_data["code_daily"] and not _tmpl_data["labor_daily"]:
                            st.caption(
                                "ℹ️ 이 파일에는 단가산출근거·일위대가 시트가 없어 Q값·노무비 역산을 쓸 수 없습니다. "
                                "1일 작업량은 표준품셈·가이드라인으로만 매칭되므로 '매칭 안 됨' 항목이 많을 수 있습니다."
                            )
                else:
                    st.warning("⚠️ 인식할 수 없는 엑셀 양식입니다. 지원 양식: 표준형(산근호표), 코드매칭형(내역서산근), 내역서 단독형(설계내역서·도급내역서)")
                
                if hierarchy:
                    # 중간 번호 기준 그룹핑 (1.1.X, 1.2.X, 2.1.X... 구분)
                    major_groups = {}
                    
                    for cat in hierarchy:
                        level = cat['level']
                        name = cat['name']
                        
                        # 중간 번호 추출 (1.1.X → "1.1")
                        parts = level.split('.')
                        if len(parts) >= 2:
                            major_key = f"{parts[0]}.{parts[1]}"
                        else:
                            major_key = parts[0]
                        
                        if major_key not in major_groups:
                            major_groups[major_key] = []
                        major_groups[major_key].append(cat)
                    
                    # 각 그룹 내에서 번호 순서 정렬
                    for major_key in major_groups:
                        major_groups[major_key].sort(key=lambda x: tuple(int(p) for p in x['level'].split('.')))
                    
                    # 프로그레스 완료
                    progress_bar.progress(100)
                    status_text.empty()
                    progress_bar.empty()
                    
                    st.success(f"✅ 파싱 완료! {len(major_groups)}개 공종 그룹, {sum(len(v) for v in major_groups.values())}개 주공종 인식")
                    
                    # 그룹명 정의
                    major_names = {
                        "1.1": "🏗️ 하수관로공사",
                        "1.2": "🔧 관로 부대공사",
                        "2.1": "💧 배수설비공사",
                        "2.2": "⚙️ 기계설비",
                        "3.1": "⚡ 전기공사",
                    }
                    
                    # 탭 생성
                    sorted_keys = sorted(major_groups.keys(), key=lambda x: tuple(int(p) for p in x.split('.')))
                    def _major_tab_label(key):
                        if key.startswith("90."):
                            # 신양식 합성코드: 그룹 내 첫 카테고리 이름으로 표시 (그룹당 대공종 1개)
                            cats = major_groups.get(key, [])
                            return f"📁 {cats[0]['name']}" if cats else f"📁 {key}"
                        return major_names.get(key, f"📁 {key}")
                    tab_labels = [_major_tab_label(key) for key in sorted_keys]
                    
                    major_tabs = st.tabs(tab_labels)
                    
                    # 🔧 session_state 초기화
                    if 'crew_by_main' not in st.session_state:
                        st.session_state['crew_by_main'] = {}
                    
                    # 🔧 모든 카테고리 crew 기본값 미리 설정 (TAB 열기 전에!)
                    all_crew_settings = {}
                    for major_key in sorted_keys:
                        for cat in major_groups[major_key]:
                            cat_name = cat['name']
                            cat_level = cat['level']
                            cat_full = f"{cat_level} {cat_name}"
                            # 기본값 3조
                            all_crew_settings[cat_name] = st.session_state['crew_by_main'].get(cat_full, DEFAULT_CREW)
                    
                    for tab_idx, (major_key, major_tab) in enumerate(zip(sorted_keys, major_tabs)):
                        with major_tab:
                            cats_in_major = major_groups[major_key]
                            
                            st.markdown(f"### 🔧 투입조수 설정")
                            
                            if 'crew_by_main' not in st.session_state:
                                st.session_state['crew_by_main'] = {}
                            
                            cols = st.columns(min(len(cats_in_major), 4))
                            
                            for idx, cat in enumerate(cats_in_major):
                                cat_level = cat['level']
                                cat_name = cat['name']
                                cat_full = f"{cat_level} {cat_name}"
                                
                                default_crew = st.session_state['crew_by_main'].get(cat_full, DEFAULT_CREW)
                                
                                with cols[idx % len(cols)]:
                                    crew_val = st.number_input(
                                        f"{cat_full}(조)",
                                        min_value=1,
                                        max_value=200,
                                        value=default_crew,
                                        key=f"crew_{major_key.replace('.', '_')}_{idx}",
                                        help="가옥별 배수설비처럼 여러 팀이 동시 투입되는 공종은 수십 조가 될 수 있습니다."
                                    )
                                    all_crew_settings[cat_name] = crew_val
                                    st.session_state['crew_by_main'][cat_full] = crew_val
                    
                    crew_settings = all_crew_settings
                    
                    # ── 🎯 주공정 선택 + 운반류 자동 제외 ──
                    _major_seen = []
                    for _cat in hierarchy:
                        if _cat['name'] not in _major_seen:
                            _major_seen.append(_cat['name'])
                    _default_major = [c for c in _major_seen if not is_non_work_category(c)]
                    st.markdown("---")
                    st.markdown("### 🎯 주공정 선택")
                    st.caption("공기를 지배하는 주공정만 선택하세요. 선택된 공종을 어떻게 합칠지는 아래에서 정합니다.")
                    selected_major = st.multiselect(
                        "공기에 반영할 대공종",
                        options=_major_seen,
                        default=_default_major,
                        key="selected_major_widget",
                    )
                    st.session_state["selected_major"] = selected_major
                    
                    _combine_mode = st.radio(
                        "주공정 합산 방식",
                        options=["최장(병행)", "합산(순차)"],
                        index=0,
                        horizontal=True,
                        key="combine_mode_widget",
                        help=(
                            "최장(병행): 선택 공종이 동시에 진행된다고 보고 가장 긴 공종을 공기로 사용합니다. "
                            "당일 굴착·부설·복구처럼 함께 나가는 공사에 적합합니다.\n\n"
                            "합산(순차): 선택 공종이 앞뒤로 이어진다고 보고 작업일수를 모두 더합니다. "
                            "관로공 완료 후 배수설비공을 시행하는 것처럼 선후 관계가 뚜렷할 때 적합합니다."
                        ),
                    )
                    st.session_state["combine_mode"] = _combine_mode
                    exclude_haul = st.checkbox(
                        "운반·사토류 항목 자동 제외 (덤프 이동 중 굴착이 계속되므로 병행 처리, 상차는 공기에 반영)",
                        value=True,
                        key="exclude_haul_widget",
                    )
                    st.session_state["exclude_haul"] = exclude_haul
                    
                    st.markdown("---")
                    st.markdown("### 📊 공종별 작업일수 계산 결과")

                    # 출처 범례 — 값이 어디서 왔는지에 따라 신뢰도가 다르므로
                    # 사용자가 필요할 때 펼쳐볼 수 있게 접힌 상태로 제공한다.
                    with st.expander("ℹ️ '출처' 표기 설명 — 1일 작업량을 어디서 가져왔는지", expanded=False):
                        st.markdown("""
**적용 우선순위** (위에 있을수록 먼저 적용됩니다)

| 출처 | 의미 | 신뢰도 |
|---|---|---|
| **수동입력** | 사용자가 직접 입력한 값. 다른 모든 값보다 우선합니다. | 사용자 판단 |
| **실무기준(1조=1개소/일)** | 배수설비처럼 일위대가 1개소에 여러 공종이 묶여 있어 역산이 맞지 않는 항목에 적용한 현장 기준값. | 높음 |
| **단가산출근거(호표)**<br>**일위대가_산근(코드)** | 설계자가 단가산출근거·일위대가 산근에 직접 적어둔 시간당 작업량(`Q = … 단위/HR`)을 읽어 ×8시간으로 환산한 값. | **가장 높음** |
| **동일계열 Q승계** | 같은 계열(동일 항목명·규격)에 Q산식이 있어 그 값을 물려받은 항목. 아스팔트 기층↔표층처럼 시공 조건이 같을 때 적용됩니다. | 중간 (검토 권장) |
| **노무비역산(직종)** | 일위대가에 인수(단위 '인')나 Q산식이 없어, 노무비를 해당 직종 노임단가로 나눠 되짚은 추정값. 할증·요율은 제외했습니다. | 낮음 (검토 권장) |
| **가이드라인** | 국토교통부 「적정 공사기간 확보를 위한 가이드라인」 부록의 공종별 1일 작업량. | 높음 |
| **단가산출근거** | 항목명 기준으로 단가산출근거에서 찾은 Q값. | 중간 |
| **병행(제외)** | 운반·상차·사토처럼 본공정과 함께 진행되어 공기를 따로 만들지 않는 항목. 0일로 처리됩니다. | — |
| **매칭 안 됨** | 어디에서도 값을 찾지 못한 항목. **수동입력 관리** 탭에서 직접 입력해야 공기에 반영됩니다. | 입력 필요 |

💡 **동일계열 Q승계**·**노무비역산**은 추정값이라 실제와 차이가 날 수 있습니다.
'수동입력 관리 → 추정값 검토' 탭에서 작업일수가 큰 항목부터 확인해 보세요.
                        """)
                    
                    result_rows = []
                    
                    for cat in hierarchy:
                        cat_name = cat['name']
                        cat_level = cat['level']
                        cat_crew = crew_settings.get(cat_name, DEFAULT_CREW)
                        
                        all_cat_items = list(cat.get('items', []))
                        for sub in cat.get('sub_categories', []):
                            all_cat_items.extend(sub.get('items', []))
                            # sub_sub_categories도 포함 (3단계 계층)
                            for sub_sub in sub.get('sub_categories', []):
                                all_cat_items.extend(sub_sub.get('items', []))
                        
                        # 라인(A-LINE 등)·지구(Ⅰ/Ⅱ 등)는 지리적으로 분리된 구간이라 동시 시공한다.
                        # 따라서 같은 공종 안에서 라인별로 일수를 나눠 합산한 뒤, 최장 라인을 그
                        # 공종의 작업일수로 삼는다 (라인 내부는 순차 → 합산, 라인 간은 병행 → 최대).
                        _line_days = {}
                        for item in all_cat_items:
                            _d = calc_days_priority(item['name'], item.get('spec', ''), item.get('qty', 0), cat_crew, item.get('unit', ''))[0]
                            _ln = item.get('district') or '(공통)'
                            _line_days[_ln] = _line_days.get(_ln, 0) + _d
                        cat_total_days = max(_line_days.values()) if _line_days else 0
                        
                        if all_cat_items:
                            result_rows.append({
                                "level": cat_level,
                                "공종": f"{cat_level} {cat_name}",
                                "물량": f"{len(all_cat_items)}개 항목",
                                "투입조수": f"{cat_crew}조",
                                "작업일수(일)": int(cat_total_days),
                                "라인별일수": {k: int(v) for k, v in _line_days.items()},
                                "세부항목": all_cat_items,
                                "하위카테고리": cat.get('sub_categories', []),
                                "crew": cat_crew,
                                "major_key": '.'.join(cat_level.split('.')[:2])  # "1.1", "2.1" 등
                            })
                    
                    # 정렬
                    def sort_key(row):
                        level = row['level']
                        parts = level.split('.')
                        return tuple(int(p) for p in parts)
                    
                    result_rows_sorted = sorted(result_rows, key=sort_key)
                    
                    # ═══════════════════════════════════════════════════════
                    # 같은 공종명끼리 합산 (토공 + 토공 → 토공)
                    # ═══════════════════════════════════════════════════════
                    merged_rows = {}
                    for row in result_rows_sorted:
                        # 공종명 추출 (번호 제거: "1.1.1 토공" → "토공", "1.2.1 관로 부대공사" → "관로 부대공사")
                        parts = row['공종'].split(maxsplit=1)  # 첫 번째 공백으로만 분리
                        if len(parts) > 1 and parts[0][0].isdigit():
                            cat_name = parts[1]  # 번호 이후 전체를 공종명으로
                        else:
                            cat_name = row['공종']
                        
                        # 키를 major_key + 공종명으로 설정 (같은 이름이라도 다른 그룹은 구분)
                        merge_key = f"{row['major_key']}_{cat_name}"
                        
                        if merge_key not in merged_rows:
                            merged_rows[merge_key] = {
                                "level": row['level'],
                                "공종": cat_name,
                                "공종명_pure": cat_name,
                                "물량": 0,
                                "투입조수": row['투입조수'],
                                "작업일수(일)": 0,
                                "세부항목": [],
                                "하위카테고리": [],
                                "crew": row['crew'],
                                "major_key": row['major_key'],
                                "원본_공종들": [],
                                "_라인별누적": {}
                            }
                        
                        # 합산 (단, 작업일수는 라인별로 누적한 뒤 최장 라인을 취한다 — 라인 간 병행)
                        for _ln, _d in (row.get("라인별일수") or {}).items():
                            merged_rows[merge_key]["_라인별누적"][_ln] = \
                                merged_rows[merge_key]["_라인별누적"].get(_ln, 0) + _d
                        if not row.get("라인별일수"):
                            # 라인 정보가 없는 구조면 기존 방식대로 합산
                            merged_rows[merge_key]["_라인별누적"]["(공통)"] = \
                                merged_rows[merge_key]["_라인별누적"].get("(공통)", 0) + row["작업일수(일)"]
                        merged_rows[merge_key]["작업일수(일)"] = max(merged_rows[merge_key]["_라인별누적"].values())
                        merged_rows[merge_key]["세부항목"].extend(row["세부항목"])
                        merged_rows[merge_key]["하위카테고리"].extend(row["하위카테고리"])
                        merged_rows[merge_key]["원본_공종들"].append(row['공종'])
                    
                    # 물량 정보 업데이트
                    for merge_key, row in merged_rows.items():
                        total_items = len(row["세부항목"])
                        row["물량"] = f"{total_items}개 항목"
                        
                        # 공종명 정리: "1.1_관로 부대공사" → "관로 부대공사"
                        pure_name = row["공종"]  # 이미 번호 제거된 순수 공종명
                        
                        if len(row["원본_공종들"]) > 1:
                            row["공종"] = f"{pure_name} (통합)"
                        else:
                            row["공종"] = pure_name
                        
                        # major_key가 없으면 첫 번째 원본 공종에서 추출
                        if not row.get("major_key") and row.get("원본_공종들"):
                            first_gong_jong = row["원본_공종들"][0]
                            # "1.2.1 관로 부대공사" → "1.2"
                            parts = first_gong_jong.split(maxsplit=1)
                            if parts and parts[0][0].isdigit():
                                level_parts = parts[0].split('.')
                                if len(level_parts) >= 2:
                                    row["major_key"] = f"{level_parts[0]}.{level_parts[1]}"
                    
                    result_rows_merged = list(merged_rows.values())
                    # 주공정으로 선택된 대공종만 공기에 반영. 합산 방식은 사용자 선택(최장/합산)을 따른다.
                    _sel_major = set(st.session_state.get("selected_major", []))
                    _major_rows = [r for r in result_rows_merged if r.get("공종명_pure") in _sel_major] or result_rows_merged
                    if st.session_state.get("combine_mode", "최장(병행)").startswith("합산"):
                        max_days = sum(r["작업일수(일)"] for r in _major_rows)
                    else:
                        max_days = max((r["작업일수(일)"] for r in _major_rows), default=0)
                    
                    # 그룹별로 표시
                    grouped_results = {}
                    for row in result_rows_merged:
                        major_key = row['major_key']
                        if major_key not in grouped_results:
                            grouped_results[major_key] = []
                        grouped_results[major_key].append(row)
                    
                    # 🔧 session_state에 저장 (TAB 2에서 사용)
                    st.session_state['grouped_results'] = grouped_results
                    st.session_state['group_names'] = {
                        "1.1": "🏗️ 하수관로공사",
                        "1.2": "🔧 관로 부대공사",
                        "2.1": "💧 배수설비공사",
                        "2.2": "⚙️ 기계설비",
                    }
                    st.session_state['max_days'] = max_days
                    
                    # 그룹명
                    group_names = {
                        "1.1": "🏗️ 하수관로공사",
                        "1.2": "🔧 관로 부대공사",
                        "2.1": "💧 배수설비공사",
                        "2.2": "⚙️ 기계설비",
                    }
                    
                    # 그룹별 expander
                    for major_key in sorted(grouped_results.keys(), key=lambda x: tuple(int(p) for p in x.split('.'))):
                        rows_in_group = grouped_results[major_key]
                        if major_key.startswith("90."):
                            # 신양식 합성코드(90.x)는 코드 대신 해당 그룹의 대공종명으로 표시
                            _first_name = rows_in_group[0].get("공종명_pure") or rows_in_group[0].get("공종", major_key)
                            group_name = f"📁 {_first_name}"
                        else:
                            group_name = group_names.get(major_key, f"📁 {major_key}")
                        
                        # 접힌 상태에서도 공기를 알 수 있도록 헤더에 작업일수/배지 표시
                        _grp_days = max((r["작업일수(일)"] for r in rows_in_group), default=0)
                        _sel_grp = set(st.session_state.get("selected_major", []))
                        _grp_selected = (not _sel_grp) or any(r.get("공종명_pure") in _sel_grp for r in rows_in_group)
                        _grp_badge = "🎯" if _grp_selected else "⚪"
                        _grp_peak = max((r["작업일수(일)"] for r in _major_rows), default=0)
                        _grp_mark = " 🔴" if (_grp_days == _grp_peak and _grp_peak > 0 and _grp_selected) else ""
                        _grp_suffix = "" if _grp_selected else " · 병행"
                        
                        # outer expander를 닫힌 상태로 시작 (성능 향상)
                        with st.expander(
                            f"{_grp_badge} **{group_name.replace('📁 ', '')}** — {_grp_days}일{_grp_mark}"
                            f" ({len(rows_in_group)}개 공종{_grp_suffix})",
                            expanded=False
                        ):
                            for idx, row in enumerate(rows_in_group):
                                _peak = max((r["작업일수(일)"] for r in _major_rows), default=0)
                                is_max = (row["작업일수(일)"] == _peak and _peak > 0)
                                _sel_major_disp = set(st.session_state.get("selected_major", []))
                                _is_selected_major = (not _sel_major_disp) or (row.get("공종명_pure") in _sel_major_disp)
                                _major_badge = "🎯" if _is_selected_major else "⚪"
                                
                                with st.expander(
                                    f"{'🔴' if is_max else '▶'} {_major_badge} **{row['공종']}** - {row['작업일수(일)']}일"
                                    + ("" if _is_selected_major else " _(병행·공기 미반영)_"),
                                    expanded=False
                                ):
                                    # 라인(구간)별 내역 — 라인 간 병행이므로 최장 라인이 이 공종의 작업일수
                                    _lb = row.get("_라인별누적") or {}
                                    if len(_lb) > 1:
                                        _lb_txt = " · ".join(
                                            f"**{k} {v}일**" if v == max(_lb.values()) else f"{k} {v}일"
                                            for k, v in sorted(_lb.items(), key=lambda x: -x[1])
                                        )
                                        st.caption(f"🔀 구간별 병행: {_lb_txt} → 최장 구간 기준 {row['작업일수(일)']}일")
                                    
                                    # 최장 공종 대비 이 공종이 과도하게 길면 필요 조수 역산 안내
                                    try:
                                        _cur_days = int(row['작업일수(일)'])
                                        _cur_crew = int(row.get('crew', DEFAULT_CREW) or DEFAULT_CREW)
                                        if _peak > 0 and _cur_days > _peak and _cur_crew >= 1:
                                            _need = -(-(_cur_days * _cur_crew) // _peak)
                                            st.caption(
                                                f"⚠️ 이 공종이 최장 공종({_peak}일)보다 깁니다 — "
                                                f"현재 {_cur_crew}조 → **{_need}조** 투입 시 {_peak}일 내 완료 가능"
                                            )
                                    except Exception:
                                        pass
                                    
                                    # 하위 카테고리별 표시
                                    if row['하위카테고리']:
                                        # 같은 name의 sub_category를 하나로 합치기
                                        merged_by_name = {}
                                        for sub in row['하위카테고리']:
                                            sub_name = sub['name']
                                            if sub_name not in merged_by_name:
                                                merged_by_name[sub_name] = {
                                                    'level': sub.get('level', ''),  # 첫 번째 level 사용
                                                    'name': sub_name,
                                                    'items': [],
                                                    'sub_categories': []  # 3단계 계층
                                                }
                                            # 항목 합치기 (name+spec 기준으로 중복 제거하면서)
                                            for item in sub.get('items', []):
                                                existing = next((i for i in merged_by_name[sub_name]['items']
                                                               if i['name'] == item['name'] and i.get('spec') == item.get('spec')), None)
                                                if existing:
                                                    existing['qty'] = existing.get('qty', 0) + item.get('qty', 0)
                                                else:
                                                    merged_by_name[sub_name]['items'].append(item)
                                            # sub_categories도 합치기 (name 기준으로 중복 제거)
                                            for sub_sub in sub.get('sub_categories', []):
                                                sub_sub_name = sub_sub['name']
                                                sub_sub_level = sub_sub.get('level', '')
                                                
                                                # 같은 name+level의 sub_sub 찾기
                                                existing_sub_sub = next((s for s in merged_by_name[sub_name]['sub_categories']
                                                                       if s['name'] == sub_sub_name and s.get('level') == sub_sub_level), None)
                                                
                                                if existing_sub_sub:
                                                    # 기존 sub_sub에 항목 합치기
                                                    for item in sub_sub.get('items', []):
                                                        existing_item = next((i for i in existing_sub_sub['items']
                                                                           if i['name'] == item['name'] and i.get('spec') == item.get('spec')), None)
                                                        if existing_item:
                                                            existing_item['qty'] = existing_item.get('qty', 0) + item.get('qty', 0)
                                                        else:
                                                            existing_sub_sub['items'].append(item)
                                                else:
                                                    # 새로운 sub_sub 추가
                                                    merged_by_name[sub_name]['sub_categories'].append(sub_sub)
                                        
                                        # 합쳐진 sub_category 표시
                                        for sub_name, sub_data in merged_by_name.items():
                                            sub_level = sub_data.get('level', '')
                                            sub_items = sub_data.get('items', [])
                                            
                                            # 항목이 없으면 건너뛰기
                                            if not sub_items and not sub_data.get('sub_categories'):
                                                continue
                                            
                                            # sub_days 계산 (sub_items + sub_sub_categories)
                                            sub_days = 0
                                            for item in sub_items:
                                                d, _, _ = calc_days_priority(item['name'], item.get('spec', ''), item.get('qty', 0), row['crew'], item.get('unit', ''))
                                                sub_days += d
                                                
                                                # 추진공 디버깅 (큰 작업일수만)
                                                if "추진" in sub_name and d > 100:
                                                    print(f"⚠️ 큰 작업일수: {sub_name} - {item['name']} ({item.get('spec', '')}) qty={item.get('qty', 0)} unit={item.get('unit', '')} → {d}일")
                                            
                                            # sub_sub_categories가 있으면 그것도 포함
                                            for sub_sub in sub_data.get('sub_categories', []):
                                                for item in sub_sub.get('items', []):
                                                    d, _, _ = calc_days_priority(item['name'], item.get('spec', ''), item.get('qty', 0), row['crew'], item.get('unit', ''))
                                                    sub_days += d
                                            
                                            # 추진공인 경우 총 연장 계산 (M 단위 항목만)
                                            total_length = 0
                                            if "추진공" in sub_name or "추진 가시설공" in sub_name:
                                                for item in sub_items:
                                                    if item.get('unit') in ['M', 'm', 'M', 'ｍ']:
                                                        total_length += item.get('qty', 0)
                                            
                                            # 헤더 표시
                                            if total_length > 0:
                                                st.markdown(f"#### {sub_level} {sub_name} (총 연장: {total_length:,.1f} M, 작업일수: {sub_days}일)")
                                            else:
                                                st.markdown(f"#### {sub_level} {sub_name} ({sub_days}일)")
                                            
                                            # 🔥 sub_items 먼저 표시 (추진공의 경우 여기에 합산됨!)
                                            if sub_items:
                                                detail_items = []
                                                for item in sub_items:
                                                    d, label, method = calc_days_priority(
                                                        item['name'],
                                                        item.get('spec', ''),
                                                        item.get('qty', 0),
                                                        row['crew'],
                                                        item.get('unit', '')
                                                    )
                                                    detail_items.append({
                                                        "세부공종": item['name'],
                                                        "규격": item.get('spec', ''),
                                                        "수량": f"{item.get('qty', 0):,.1f}",
                                                        "단위": item.get('unit', ''),
                                                        "1일작업량": label,
                                                        "투입조수": int(st.session_state.get("crew_by_item", {}).get(f"{item['name']}|{item.get('spec','')}", row['crew'])),
                                                        "작업일수": int(d),
                                                        "출처": source_badge(method), "_출처원본": method
                                                    })
                                                
                                                if detail_items:
                                                    render_detail_table(
                                                        detail_items, sub_items, row['crew'],
                                                        key=f"dt1_{row['공종']}_{sub_name}",
                                                    )
                                                    
                                                    # 🔧 수동입력 UI: 매칭 안 된 항목만
                                                    unmatched_items = [
                                                        (idx, item, detail_items[idx]) 
                                                        for idx, item in enumerate(sub_items)
                                                        if detail_items[idx].get("_출처원본", detail_items[idx]["출처"]) == "매칭 안 됨"
                                                    ]
                                                    
                                                    # 자재·운반류 대공종은 시공 행위가 아니라 1일 작업량 개념이 없으므로
                                                    # 수동입력 수집에서 제외 (공기 계산에서도 기본 제외되는 카테고리)
                                                    _cat_pure = row.get('공종명_pure') or row.get('공종', '')
                                                    _is_material_cat = is_non_work_category(_cat_pure)

                                                    # 🔧 항목별 조수 조정 대상 수집 (작업일수 3일 이상)
                                                    # 같은 공종이라도 항목마다 장비·인원이 달라 조수를 따로 잡아야 하는
                                                    # 경우가 있다(예: 대구경 추진 vs 소구경 부설).
                                                    if not _is_material_cat:
                                                        st.session_state.setdefault("crew_targets", {})
                                                        for _ci, _cit in enumerate(sub_items):
                                                            _cd = detail_items[_ci]
                                                            if int(_cd.get("작업일수", 0) or 0) < 3:
                                                                continue
                                                            _ckey = f"{_cit['name']}|{_cit.get('spec','')}"
                                                            st.session_state["crew_targets"][_ckey] = {
                                                                "name": _cit['name'],
                                                                "spec": _cit.get('spec', ''),
                                                                "qty": _cit.get('qty', 0),
                                                                "unit": _cit.get('unit', ''),
                                                                "category": row.get('공종명_pure', row['공종']),
                                                                "base_crew": int(row['crew']),
                                                                "days": int(_cd.get("작업일수", 0) or 0),
                                                                "rate": _cd.get("1일작업량", ""),
                                                            }

                                                    # 🔎 동일계열 Q승계 항목은 '추정값'이므로 설계자 검토 대상으로 별도 수집
                                                    _series_items = [
                                                        (idx, item, detail_items[idx])
                                                        for idx, item in enumerate(sub_items)
                                                        if detail_items[idx].get("_출처원본", detail_items[idx]["출처"]) == "동일계열 Q승계"
                                                    ]
                                                    if _series_items and not _is_material_cat:
                                                        if "series_review_all" not in st.session_state:
                                                            st.session_state["series_review_all"] = {}
                                                        _rk = f"{major_key}_{row['공종']}"
                                                        if _rk not in st.session_state["series_review_all"]:
                                                            st.session_state["series_review_all"][_rk] = {
                                                                "major_key": major_key,
                                                                "group_name": group_name,
                                                                "category": row.get('공종명_pure', row['공종']),
                                                                "items": []
                                                            }
                                                        for _i2, _it2, _di2 in _series_items:
                                                            _mk2 = f"{_it2['name']}|{_it2.get('spec', '')}"
                                                            if not [x for x in st.session_state["series_review_all"][_rk]["items"] if x["manual_key"] == _mk2]:
                                                                st.session_state["series_review_all"][_rk]["items"].append({
                                                                    "name": _it2['name'],
                                                                    "spec": _it2.get('spec', ''),
                                                                    "qty": _it2.get('qty', 0),
                                                                    "unit": _it2.get('unit', ''),
                                                                    "manual_key": _mk2,
                                                                    "sub_name": sub_name,
                                                                    "승계값": _di2.get("1일작업량", ""),
                                                                })
                                                    
                                                    # 🔎 노무비역산 항목도 '추정값'이므로 검토 대상으로 수집.
                                                    # 일위대가에 인수가 없어 노무비로 되짚은 값이라,
                                                    # 설계자가 다른 품셈을 적용한 경우와 어긋날 수 있다
                                                    # (실측: 세라믹방수 8.3㎡ vs 공식 129.1㎡ — 다른 품셈 대체 적용).
                                                    # 항목이 많으므로 작업일수가 큰 것만 모아 공기 영향 순으로 검토한다.
                                                    _labor_items = [
                                                        (idx, item, detail_items[idx])
                                                        for idx, item in enumerate(sub_items)
                                                        if str(detail_items[idx].get("_출처원본", detail_items[idx].get("출처", ""))).startswith("노무비역산")
                                                        and int(detail_items[idx].get("작업일수", 0) or 0) >= 5
                                                    ]
                                                    if _labor_items and not _is_material_cat:
                                                        if "labor_review_all" not in st.session_state:
                                                            st.session_state["labor_review_all"] = {}
                                                        _rk3 = f"{major_key}_{row['공종']}"
                                                        if _rk3 not in st.session_state["labor_review_all"]:
                                                            st.session_state["labor_review_all"][_rk3] = {
                                                                "major_key": major_key,
                                                                "group_name": group_name,
                                                                "category": row.get('공종명_pure', row['공종']),
                                                                "items": []
                                                            }
                                                        for _i3, _it3, _di3 in _labor_items:
                                                            _mk3 = f"{_it3['name']}|{_it3.get('spec', '')}"
                                                            if not [x for x in st.session_state["labor_review_all"][_rk3]["items"] if x["manual_key"] == _mk3]:
                                                                st.session_state["labor_review_all"][_rk3]["items"].append({
                                                                    "name": _it3['name'],
                                                                    "spec": _it3.get('spec', ''),
                                                                    "qty": _it3.get('qty', 0),
                                                                    "unit": _it3.get('unit', ''),
                                                                    "manual_key": _mk3,
                                                                    "sub_name": sub_name,
                                                                    "추정값": _di3.get("1일작업량", ""),
                                                                    "작업일수": int(_di3.get("작업일수", 0) or 0),
                                                                })

                                                    if unmatched_items and not _is_material_cat:
                                                        st.markdown("---")
                                                        st.info(f"📝 매칭 안 된 항목: {len(unmatched_items)}개 - **TAB 6 '수동입력 관리'**에서 입력하세요!")
                                                        
                                                        # 🔧 session_state에 매칭 안 된 항목 수집
                                                        if "unmatched_all" not in st.session_state:
                                                            st.session_state["unmatched_all"] = {}
                                                        
                                                        # 카테고리별로 저장
                                                        cat_key = f"{major_key}_{row['공종']}"
                                                        if cat_key not in st.session_state["unmatched_all"]:
                                                            st.session_state["unmatched_all"][cat_key] = {
                                                                "major_key": major_key,
                                                                "group_name": group_name,
                                                                "category": row.get('공종명_pure', row['공종']),
                                                                "items": []
                                                            }
                                                        
                                                        for idx_um, item_um, _ in unmatched_items:
                                                            manual_key_um = f"{item_um['name']}|{item_um.get('spec', '')}"
                                                            # 중복 방지
                                                            existing = [i for i in st.session_state["unmatched_all"][cat_key]["items"] if i["manual_key"] == manual_key_um]
                                                            if not existing:
                                                                st.session_state["unmatched_all"][cat_key]["items"].append({
                                                                    "name": item_um['name'],
                                                                    "spec": item_um.get('spec', ''),
                                                                    "qty": item_um.get('qty', 0),
                                                                    "unit": item_um.get('unit', ''),
                                                                    "manual_key": manual_key_um,
                                                                    "sub_name": sub_name
                                                                })
                                                        
                                                        # 🔧 수동입력 UI 일시 비활성화 (성능 문제로)
                                                        # TODO: 추후 별도 페이지로 분리
                                                        if False:  # 비활성화
                                                            # session_state 초기화
                                                            if "manual_rates" not in st.session_state:
                                                                st.session_state["manual_rates"] = {}
                                                            
                                                            for idx, item, detail_row in unmatched_items:
                                                                manual_key = f"{item['name']}|{item.get('spec', '')}"
                                                                
                                                                # 고유 키 생성 (카테고리명 포함으로 완전히 고유하게!)
                                                            cat_name = row['공종']
                                                            unique_key = f"{major_key}_{cat_name}_{sub_level}_{sub_name}_{manual_key}_{idx}".replace(" ", "_").replace("(", "").replace(")", "").replace("/", "_").replace("|", "_")
                                                            
                                                            col1, col2, col3, col4 = st.columns([3, 2, 1, 1])
                                                            
                                                            with col1:
                                                                st.text(f"{item['name']} ({item.get('spec', '')})")
                                                            
                                                            with col2:
                                                                # 기존 값 불러오기
                                                                existing_val = 0
                                                                existing_unit = item.get('unit', '') + "/일"
                                                                if manual_key in st.session_state["manual_rates"]:
                                                                    existing_val = st.session_state["manual_rates"][manual_key].get("daily", 0)
                                                                    existing_unit = st.session_state["manual_rates"][manual_key].get("unit", existing_unit)
                                                                
                                                                daily_rate = st.number_input(
                                                                    "1일 작업량",
                                                                    min_value=0.0,
                                                                    value=float(existing_val),
                                                                    step=0.1,
                                                                    key=f"manual_input_{unique_key}",
                                                                    label_visibility="collapsed"
                                                                )
                                                            
                                                            with col3:
                                                                unit_input = st.text_input(
                                                                    "단위",
                                                                    value=existing_unit,
                                                                    key=f"manual_unit_{unique_key}",
                                                                    label_visibility="collapsed"
                                                                )
                                                            
                                                            with col4:
                                                                if st.button("저장", key=f"manual_save_{unique_key}"):
                                                                    if daily_rate > 0:
                                                                        st.session_state["manual_rates"][manual_key] = {
                                                                            "daily": daily_rate,
                                                                            "unit": unit_input
                                                                        }
                                                                        st.success("✅ 저장됨!")
                                                                        st.rerun()
                                                            
                                                            # 계산 결과 미리보기
                                                            if daily_rate > 0:
                                                                calc_days = math.ceil(item.get('qty', 0) / (daily_rate * row['crew']))
                                                                st.caption(f"→ 예상 작업일수: **{calc_days}일** (조수: {row['crew']})")
                                            
                                            # sub_sub_categories가 있으면 추가로 표시 (3단계 계층)
                                            if sub_data.get('sub_categories'):
                                                for sub_sub in sub_data['sub_categories']:
                                                    sub_sub_level = sub_sub.get('level', '')
                                                    sub_sub_name = sub_sub['name']
                                                    sub_sub_items = sub_sub.get('items', [])
                                                    
                                                    if not sub_sub_items:
                                                        continue
                                                    
                                                    sub_sub_days = sum(
                                                        calc_days_priority(item['name'], item.get('spec', ''), item.get('qty', 0), row['crew'], item.get('unit', ''))[0]
                                                        for item in sub_sub_items
                                                    )
                                                    
                                                    # sub_sub 헤더 (들여쓰기로 계층 표현)
                                                    st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;**{sub_sub_level} {sub_sub_name}** ({sub_sub_days}일)")
                                                    
                                                    detail_items = []
                                                    for item in sub_sub_items:
                                                        d, label, method = calc_days_priority(
                                                            item['name'],
                                                            item.get('spec', ''),
                                                            item.get('qty', 0),
                                                            row['crew'],
                                                            item.get('unit', '')
                                                        )
                                                        detail_items.append({
                                                            "세부공종": item['name'],
                                                            "규격": item.get('spec', ''),
                                                            "수량": f"{item.get('qty', 0):,.1f}",
                                                            "단위": item.get('unit', ''),
                                                            "1일작업량": label,
                                                            "투입조수": int(st.session_state.get("crew_by_item", {}).get(f"{item['name']}|{item.get('spec','')}", row['crew'])),
                                                            "작업일수": int(d),
                                                            "출처": source_badge(method), "_출처원본": method
                                                        })
                                                    
                                                    if detail_items:
                                                        render_detail_table(
                                                            detail_items, sub_sub_items, row['crew'],
                                                            key=f"dt2_{row['공종']}_{sub_name}_{sub_sub_name}",
                                                        )
                                            
                                            # (아래 elif 블록은 제거됨 - 위에서 이미 처리)
                                            if False:
                                                detail_items = []
                                                for item in sub_items:
                                                    d, label, method = calc_days_priority(
                                                        item['name'],
                                                        item.get('spec', ''),
                                                        item.get('qty', 0),
                                                        row['crew'],
                                                        item.get('unit', '')
                                                    )
                                                    detail_items.append({
                                                        "세부공종": item['name'],
                                                        "규격": item.get('spec', ''),
                                                        "수량": f"{item.get('qty', 0):,.1f}",
                                                        "단위": item.get('unit', ''),
                                                        "1일작업량": label,
                                                        "투입조수": int(st.session_state.get("crew_by_item", {}).get(f"{item['name']}|{item.get('spec','')}", row['crew'])),
                                                        "작업일수": int(d),
                                                        "출처": source_badge(method), "_출처원본": method
                                                    })
                                                
                                                if detail_items:
                                                    render_detail_table(
                                                        detail_items, sub_items, row['crew'],
                                                        key=f"dt3_{row['공종']}_{sub_name}",
                                                    )
                                    
                                    # 직접 항목도 있으면 표시
                                    # 하위카테고리에 속하지 않은 항목만. 예전에는 항목마다 전체 목록과
                                    # 딕셔너리 동등 비교를 해서 항목 수의 제곱에 비례해 느려졌다.
                                    _in_sub = {id(i) for sub in row['하위카테고리'] for i in sub.get('items', [])}
                                    direct_items = [item for item in row['세부항목'] if id(item) not in _in_sub]
                                    if direct_items:
                                        detail_items = []
                                        for item in direct_items:
                                            d, label, method = calc_days_priority(
                                                item['name'],
                                                item.get('spec', ''),
                                                item.get('qty', 0),
                                                row['crew'],
                                                item.get('unit', '')
                                            )
                                            detail_items.append({
                                                "세부공종": item['name'],
                                                "규격": item.get('spec', ''),
                                                "수량": f"{item.get('qty', 0):,.1f}",
                                                "단위": item.get('unit', ''),
                                                "1일작업량": label,
                                                "투입조수": int(st.session_state.get("crew_by_item", {}).get(f"{item['name']}|{item.get('spec','')}", row['crew'])),
                                                "작업일수": int(d),
                                                "출처": source_badge(method), "_출처원본": method
                                            })
                                        
                                        if detail_items:
                                            render_detail_table(
                                                detail_items, direct_items, row['crew'],
                                                key=f"dt4_{row['공종']}",
                                            )
                    
                    ca, cb = st.columns(2)
                    _cm_lbl = "합산(순차)" if st.session_state.get("combine_mode","최장(병행)").startswith("합산") else "최장(병행)"
                    ca.metric(f"🔴 주공정 ({_cm_lbl})", f"{max_days}일")
                    cb.metric("총 공종", f"{len(result_rows_merged)}개")
                    
                    # session_state 저장
                    st.session_state["work_result"] = {
                        "rows": result_rows_merged,
                        "hierarchy": hierarchy,
                        "crew_settings": crew_settings,
                    }
                    st.session_state["total_work_days"] = int(max_days)
                
                # ══════════════════════════════════════════════════════════════
                # 지구별 상세 섹션
                # ══════════════════════════════════════════════════════════════
                if col_info.get("districts"):
                    st.markdown("---")
                    st.markdown("## 📍 지구별 상세")
                    
                    # 지구별 데이터 그룹핑
                    district_data = {}
                    district_names = col_info["districts"]
                    
                    for item in matched:
                        district = item.get("district", "전체")
                        if district not in district_data:
                            district_data[district] = []
                        district_data[district].append(item)
                    
                    # 지구 선택
                    selected_district = st.selectbox(
                        "🏗️ 지구 선택",
                        options=list(district_data.keys()),
                        format_func=lambda x: f"{x}. {district_names.get(x, {}).get('name', x)}",
                        key="district_selector"
                    )
                    
                    selected_data = district_data[selected_district]
                    
                    st.markdown(f"### {selected_district}. {district_names.get(selected_district, {}).get('name', selected_district)}")
                    st.caption(f"총 {len(selected_data)}개 항목")
                    
                    # 공종별 그룹핑
                    group_names = {
                        "포장복구": "🛣️ 포장공",
                        "굴착공": "⛏️ 토공",
                        "관부설공": "🔧 관부설공",
                        "되메우기": "📦 되메우기",
                        "맨홀공": "🕳️ 구조물공",
                        "배수설비": "💧 배수설비",
                        "추진공": "🚇 추진공",
                        "기타": "📋 기타"
                    }
                    
                    grouped_by_type = {}
                    for item in selected_data:
                        group = item.get("group", "기타")
                        if group not in grouped_by_type:
                            grouped_by_type[group] = []
                        grouped_by_type[group].append(item)
                    
                    # 공종별 투입조수 설정
                    st.markdown("#### 🔧 투입조수 설정")
                    
                    # 해당 지구 총 공사기간 계산 및 표시 (투입조수 설정 전에)
                    # 임시로 기본 조수 3으로 계산
                    temp_total_days = 0
                    for group, items in grouped_by_type.items():
                        for item in items:
                            days, _, _ = calc_days_priority(
                                item["name"],
                                item.get("spec", ""),
                                item.get("qty", 0),
                                DEFAULT_CREW,
                                item.get("unit", "")
                            )
                            temp_total_days = max(temp_total_days, days)
                    
                    st.info(f"📊 **{selected_district} 지구 예상 공사기간:** {int(temp_total_days)}일 (기본 3조 기준)")
                    
                    group_crews = {}
                    
                    cols = st.columns(4)
                    for idx, (group, items) in enumerate(grouped_by_type.items()):
                        with cols[idx % 4]:
                            crew = st.number_input(
                                f"{group_names.get(group, group)}",
                                min_value=1,
                                max_value=30,
                                value=3,
                                key=f"crew_{selected_district}_{group}"
                            )
                            group_crews[group] = crew
                    
                    st.markdown("---")
                    
                    # 공종별 표 표시
                    for group in ["포장복구", "맨홀공", "굴착공", "관부설공", "되메우기", "배수설비", "추진공", "기타"]:
                        if group not in grouped_by_type:
                            continue
                        
                        items = grouped_by_type[group]
                        crew = group_crews[group]
                        
                        with st.expander(f"{group_names.get(group, group)} ({len(items)}개 항목)", expanded=(group in ["포장복구", "맨홀공", "관부설공"])):
                            display_items = []
                            for item in items:
                                days, label, method = calc_days_priority(
                                    item["name"],
                                    item.get("spec", ""),
                                    item.get("qty", 0),
                                    crew,
                                    item.get("unit", "")
                                )
                                display_items.append({
                                    "공종": item["name"],
                                    "규격": item.get("spec", ""),
                                    "물량": item.get("qty", 0),
                                    "단위": item.get("unit", ""),
                                    "1일작업량": label,  # "5본/일x3조" → "5본/일"로 변경 필요
                                    "조": crew,
                                    "일수": int(days),
                                    "출처": source_badge(method), "_출처원본": method
                                })
                            
                            if display_items:
                                display_df = pd.DataFrame(
                                    [{k: v for k, v in d.items() if not str(k).startswith("_")}
                                     for d in display_items]
                                )
                                
                                st.dataframe(
                                    display_df,
                                    width="stretch",
                                    height=400,
                                    column_config={
                                        "공종": st.column_config.TextColumn("공종", width="large"),
                                        "규격": st.column_config.TextColumn("규격", width="large"),
                                        "물량": st.column_config.NumberColumn("물량", width="medium", format="%.1f"),
                                        "단위": st.column_config.TextColumn("단위", width="small"),
                                        "1일작업량": st.column_config.TextColumn("1일작업량", width="medium"),
                                        "조": st.column_config.NumberColumn("조", width="small"),
                                        "일수": st.column_config.NumberColumn("일수", width="small"),
                                        "출처": st.column_config.TextColumn("출처", width="medium"),
                                    }
                                )
                                
                                total_days = sum(item["일수"] for item in display_items)
                                st.metric(f"{group_names.get(group, group)} 총 작업일수", f"{total_days}일")
                    
            else:
                # 어떤 양식으로도 읽지 못한 파일. 예전에는 이 분기가 없어서 진행 막대가
                # 60%에서 멈춘 채 아무 안내도 없었다(실측: 동부+남천 준공내역서).
                progress_bar.empty()
                status_text.empty()
                _sheets = ", ".join(_parsed.get("sheets") or []) or "(없음)"
                st.warning(
                    "⚠️ 내역서 양식을 인식하지 못했습니다.\n\n"
                    f"**이 파일의 시트:** {_sheets}\n\n"
                    "**지원 양식** (시트 이름 기준)\n"
                    "- 표준형: `설계내역서` + `단가산출근거`\n"
                    "- 코드매칭형: `내역서` + `일위대가_산근`\n"
                    "- 내역서 단독형: `설계내역서` 또는 `도급내역서` "
                    "(머리행에 공종·명칭·규격·수량·단위)\n\n"
                    "내역서 시트 이름이 다르면 위 이름 중 하나로 바꿔 저장한 뒤 다시 올려주세요."
                )
        except (zipfile.BadZipFile, InvalidFileException):
            st.error(
                "❌ 이 파일은 .xlsx 형식이 아닙니다 (구버전 .xls로 추정).\n\n"
                "엑셀에서 파일을 연 뒤 **[다른 이름으로 저장] → [Excel 통합 문서(.xlsx)]** 로 "
                "다시 저장하고 업로드해주세요."
            )
        except Exception as e:
            st.error(f"파싱 실패: {e}")
            import traceback
            st.code(traceback.format_exc())
    else:
        st.info("도급 설계내역서 엑셀을 업로드해주세요.")

# ══════════════════════════════════════════════════════════════
# TAB 1
# ══════════════════════════════════════════════════════════════
with tab1:
    st.subheader("📋 공기산정 요약")
    
    # session_state에서 지역 가져오기 (비작업일수 탭에서 설정)
    selected_region = st.session_state.get("selected_region", "서울")
    
    # 지역이 설정되지 않은 경우 안내
    if "selected_region" not in st.session_state:
        st.warning("⚠️ 먼저 **'비작업일수 계산기'** 탭에서 공사 지역을 선택해주세요!")
    
    # 탭 제목('공기산정 요약')과 중복되므로 별도 섹션 제목 없이 지표만 표시
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("📍 공사 지역", selected_region)
    
    with col2:
        total_work_days = st.session_state.get("total_work_days", 0)
        st.metric("💼 총 순작업일수", f"{total_work_days}일")
    
    with col3:
        # 비작업일수 결과 표시 (있는 경우)
        if "weather_result" in st.session_state:
            non_work = st.session_state["weather_result"].get("non_work_days", 0)
            st.metric("🚫 비작업일수", f"{non_work}일")
        else:
            st.metric("🚫 비작업일수", "미계산")
    
    st.markdown("---")
    
    if "work_result" in st.session_state:
        st.success("✅ 엑셀 인식 탭에서 계산 완료!")
        
        # 최종 결과 (비작업일수 계산 완료된 경우)
        if "weather_result" in st.session_state:
            result = st.session_state["weather_result"]
            
            st.markdown("### 🎯 최종 공기산정 결과")
            
            col_a, col_b, col_c, col_d = st.columns(4)
            # 총 공사기간에는 준비·시운전·정리기간이 포함되므로
            # '작업+비작업'과 값이 다르다. 라벨에 그 사실을 명시해 혼동을 막는다.
            _extra_a = (result.get("prep_days", 0) + result.get("wrapup_days", 0)
                        + result.get("commission_days", 0))
            _lbl_a = "📅 총 공사기간" + (" (준비·시운전·정리 포함)" if _extra_a else "")
            col_a.metric(_lbl_a, f"{result['total_days']}일",
                         delta=days_to_months_text(result['total_days']), delta_color="off")
            col_b.metric("💼 순작업일수", f"{result['work_days']}일")
            col_c.metric("🚫 비작업일수", f"{result['non_work_days']}일")
            col_d.metric("📍 적용 지역", result['region'])

            # 공사기간 구성 (가이드라인: 준비 + 작업 + 비작업 + 시운전 + 정리)
            # 상단 메트릭만 보면 '작업+비작업'과 총계가 안 맞아 보이므로,
            # 구성요소를 같은 크기의 지표 카드로 나란히 보여준다.
            if _extra_a:
                _pp = result.get("prep_days", 0)
                _ww = result.get("wrapup_days", 0)
                _cc2 = result.get("commission_days", 0)
                _items = [
                    ("🏁 준비기간", _pp),
                    ("💼 작업일수", result["work_days"]),
                    ("🚫 비작업일수", result["non_work_days"]),
                    ("⚙️ 시운전", _cc2),
                    ("🧹 정리기간", _ww),
                ]
                _items = [(lb, v) for lb, v in _items if v]
                st.markdown("#### 🗓️ 공사기간 구성")
                _bcols = st.columns(len(_items) + 1)
                for _bi, (_lb, _v) in enumerate(_items):
                    _sign = "" if _bi == 0 else "＋"
                    _bcols[_bi].metric(f"{_sign}{_lb}", f"{_v}일")
                _bcols[-1].metric("＝ 총 공사기간", f"{result['total_days']}일",
                                  delta=days_to_months_text(result["total_days"]),
                                  delta_color="off")
            
            _cond_md = "\n            ".join(f"- {l}" for l in applied_condition_lines(result))
            st.info(f"""
            **📍 {result.get('station') or result['region']} 공기산정 결과**
            - 착공일: {result['start_date'].strftime('%Y년 %m월 %d일')}
            - 준공일: {result['end_date'].strftime('%Y년 %m월 %d일')}
            - 총 공사기간: **{result['total_days']}일** {days_to_months_text(result['total_days'])}
            
            **적용된 비작업일 조건:**
            {_cond_md}
            """)
        else:
            st.info("👉 **'비작업일수 계산기'** 탭에서 비작업일수를 계산하면 최종 공기산정 결과가 표시됩니다!")
    else:
        st.warning("⚠️ **'엑셀 내역서 인식'** 탭에서 엑셀을 먼저 업로드해주세요.")

# ══════════════════════════════════════════════════════════════
# TAB 3
# ══════════════════════════════════════════════════════════════
with tab3:
    st.subheader("주요공종 CP 분석")
    work_result = st.session_state.get("work_result")
    if work_result:
        result_rows = work_result["rows"]
        
        display_data = []
        for r in result_rows:
            # 자재·운반류는 시공 공정이 아니므로 CP(주공정) 분석 대상에서 제외
            _cp_cat = r.get("공종명_pure") or r.get("공종", "")
            if is_non_work_category(_cp_cat):
                continue
            display_data.append({
                "공종": r["공종"],
                "물량": r["물량"],
                "투입조수": r["투입조수"],
                "작업일수(일)": r["작업일수(일)"]
            })
        
        df_cp = pd.DataFrame(display_data)
        df_cp = df_cp[df_cp["작업일수(일)"] > 0].copy()
        max_days = df_cp["작업일수(일)"].max() if len(df_cp) > 0 else 0

        def hl_cp(row):
            if row["작업일수(일)"] == max_days:
                return ["background-color:#3d0000;color:#ff6b6b"] * len(row)
            return [""] * len(row)

        st.dataframe(df_cp.style.apply(hl_cp, axis=1), hide_index=True, width="stretch")
        
        fig_bar = px.bar(df_cp, x="작업일수(일)", y="공종", orientation="h", text="작업일수(일)",
                         color="작업일수(일)", color_continuous_scale=["#27AE60","#F39C12","#E74C3C"])
        fig_bar.update_layout(height=350, showlegend=False, yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig_bar, width="stretch")
    else:
        st.warning("TAB 2에서 엑셀을 먼저 업로드해주세요.")

# ══════════════════════════════════════════════════════════════
# TAB 4
# ══════════════════════════════════════════════════════════════
with tab4:
    st.subheader("🌧 비작업일수 계산기")
    st.caption("지역별 기상 데이터를 기반으로 비작업일수를 계산합니다")
    
    # ──────────────────────────────────
    # 1. 지역 선택
    # ──────────────────────────────────
    st.markdown("### 🌍 공사 지역 선택")
    
    if HAS_GUIDELINE_WEATHER:
        # 시·도 → 관측지점 2단계 선택.
        # @st.fragment로 감싸면 시·도를 바꿔도 이 영역만 다시 그려지고
        # 앱 전체(파싱·계산)는 재실행되지 않는다. '지역 적용'을 눌렀을 때만
        # 전체를 갱신하므로, 두 단계를 모두 고른 뒤 한 번만 리뉴얼된다.
        _sido_map = {}
        for _s in STATIONS:
            _sido_map.setdefault(_s.split(" ", 1)[0], []).append(_s)
        _sido_list = sorted(_sido_map.keys())

        if "selected_station" not in st.session_state:
            st.session_state["selected_station"] = STATIONS[0] if STATIONS else ""
            st.session_state["selected_region"] = st.session_state["selected_station"].split(" ")[0]

        @st.fragment
        def _region_picker():
            _cur = st.session_state.get("selected_station", "")
            _cur_sido = _cur.split(" ")[0] if _cur else _sido_list[0]
            _c0, _c1, _c2 = st.columns([1.2, 1.2, 1])
            with _c0:
                _sido = st.selectbox(
                    "시·도", options=_sido_list,
                    index=_sido_list.index(_cur_sido) if _cur_sido in _sido_list else 0,
                    key="region_sido",
                    help="공사 현장이 속한 광역시·도를 먼저 선택하세요.",
                )
            _stns = _sido_map[_sido]
            _names = [x.split(" ", 1)[1] if " " in x else x for x in _stns]
            with _c1:
                _name = st.selectbox(
                    "관측지점", options=_names,
                    index=_stns.index(_cur) if _cur in _stns else 0,
                    key="region_station",
                    help="현장에서 가장 가까운 기상관측지점을 고르세요.",
                )
            _pick = _stns[_names.index(_name)]
            with _c2:
                st.markdown("&nbsp;", unsafe_allow_html=True)
                if st.button("📍 지역 적용", type="primary", width="stretch",
                             disabled=(_pick == _cur)):
                    st.session_state["selected_station"] = _pick
                    st.session_state["selected_region"] = _pick.split(" ")[0]
                    st.rerun()
            if _pick != _cur:
                st.caption(f"선택: {_pick} → **지역 적용**을 누르면 반영됩니다.")

        _region_picker()
        selected_station = st.session_state.get("selected_station", "")
        selected_region = st.session_state.get("selected_region", "")
        st.info(f"📍 현재 적용 지점: **{selected_station}**")
    else:
        default_region = st.session_state.get("selected_region", "서울")
        if default_region not in REGIONS:
            default_region = "서울"
        selected_region = st.selectbox(
            "공사 지역", options=REGIONS, index=REGIONS.index(default_region),
            help="지역별 기상 데이터 적용",
        )
        st.session_state["selected_region"] = selected_region
        selected_station = None
        st.info(f"📍 현재 적용 지역: **{selected_region}**")
    
    st.markdown("---")
    
    # ──────────────────────────────────
    # 2. 비작업일 조건
    # ──────────────────────────────────
    st.markdown("### ⚙️ 비작업일 조건 설정")
    st.caption("📖 가이드라인 공식: 비작업일수 = 기상조건(A) + 법정공휴일(B) - 중복일수(C)")
    
    st.markdown("**🌤️ 기상조건 (A)**")
    if HAS_GUIDELINE_WEATHER:
        st.caption(
            "가이드라인 부록3의 13개 조건 중 이 공사에 적용할 항목을 고르세요. "
            "선택한 조건의 비작업일수를 합산합니다. "
            "기본값은 실제 공기산정 보고서에서 통용되는 조합입니다."
        )
        _sel_conds = []
        _cond_keys = list(CONDITION_LABELS.keys())
        _cc = st.columns(2)
        for _ci, _ck in enumerate(_cond_keys):
            with _cc[_ci % 2]:
                if st.checkbox(
                    CONDITION_LABELS[_ck],
                    value=(_ck in DEFAULT_CONDITIONS),
                    key=f"wcond_{_ck}",
                ):
                    _sel_conds.append(_ck)
        st.session_state["weather_conditions"] = _sel_conds
        include_rain = include_cold = include_hot = True  # 구 로직 호환용
    else:
        col1, col2, col3 = st.columns(3)
        with col1:
            include_rain = st.checkbox("🌧️ 강우일 포함", value=True, help="월별 평균 강우일수")
        with col2:
            include_cold = st.checkbox("❄️ 한랭일 포함", value=True, help="일 최저기온 -10°C 이하")
        with col3:
            include_hot = st.checkbox("🔥 폭염일 포함", value=True, help="일 최고기온 33°C 이상")
        _sel_conds = None
    
    st.markdown("**📅 법정공휴일 (B)**")
    col_h1, col_h2 = st.columns(2)
    with col_h1:
        include_holidays = st.checkbox(
            "📅 법정공휴일 포함",
            value=True,
            help="가이드라인 부록1: 일요일(52일) + 명절 + 국경일 + 기타 + 대체공휴일"
        )
    with col_h2:
        min_weekly_rest = st.checkbox(
            "⚖️ 주 40시간 근무제 보장",
            value=True,
            help="월별 비작업일수가 주 40시간 근무제 일수보다 작으면 보정"
        )

    # ──────────────────────────────────
    # 준비기간 · 정리기간 · 시운전
    # 가이드라인상 공사기간 = 준비기간 + 작업일수 + 비작업일수 + 정리기간
    # 상하수도 공사는 통상 준비 2개월, 정리 1개월을 잡고,
    # 하수처리시설·정수장 등 시설공사는 준공 전 시운전 기간(약 4개월)이 추가된다.
    # 현장사무실 설치 같은 항목은 이 준비기간에 포함되므로
    # 공종별 작업일수에서는 빼는 것이 맞다(상세표의 '제외' 사용).
    # ──────────────────────────────────
    st.markdown("**🗓️ 준비 · 정리 · 시운전 기간**")
    st.caption(
        "가이드라인 공식: 공사기간 = **준비기간 + 작업일수 + 비작업일수 + 정리기간**. "
        "현장사무실·가설건축물 설치는 준비기간에 포함되므로 공종별 작업일수에서는 제외하세요."
    )
    _pc1, _pc2, _pc3 = st.columns(3)
    with _pc1:
        prep_months = st.number_input(
            "준비기간(개월)", min_value=0.0, max_value=12.0, value=2.0, step=0.5,
            help="인허가·용지보상·현장사무실 설치 등. 상하수도 공사는 통상 2개월.",
        )
    with _pc2:
        wrapup_months = st.number_input(
            "정리기간(개월)", min_value=0.0, max_value=12.0, value=1.0, step=0.5,
            help="현장 정리·준공 서류 등. 상하수도 공사는 통상 1개월.",
        )
    with _pc3:
        commission_months = st.number_input(
            "시운전(개월)", min_value=0.0, max_value=12.0, value=0.0, step=0.5,
            help="하수처리시설·정수장 등 시설공사는 준공 전 시운전이 필요합니다(통상 4개월). "
                 "관로 공사만 있으면 0으로 두세요.",
        )
    
    st.markdown("---")
    
    # ──────────────────────────────────
    # 3. 공사 기간 설정
    # ──────────────────────────────────
    st.markdown("### 📅 공사 기간 설정")
    
    col_a, col_b = st.columns(2)
    with col_a:
        start_date = st.date_input(
            "착공일",
            value=st.session_state.get("start_date", datetime.now().date()),
            key="weather_start_date"
        )
        st.session_state["start_date"] = start_date
    
    with col_b:
        # TAB 2에서 계산된 순작업일수 자동 입력.
        # 순공기가 '새로 계산됐을 때만' 위젯 상태를 직접 동기화한다.
        # (사용자가 수동으로 고친 값은 다음 재계산 전까지 유지됨)
        default_work_days = int(st.session_state.get("total_work_days", 100) or 0)
        if default_work_days >= 1 and st.session_state.get("_synced_work_days") != default_work_days:
            st.session_state["weather_work_days"] = default_work_days
            st.session_state["_synced_work_days"] = default_work_days
        # value= 인자는 넘기지 않고 세션 상태로만 값을 준다.
        # 산정 결과가 0일(전부 '매칭 안 됨'이거나 전부 제외)일 때 value=0을 넘기면
        # min_value=1 위반으로 예외가 나서, 뒤에 그려지는 탭(수동입력 관리 포함)이
        # 통째로 사라졌다. 정작 수동입력이 필요한 상황에서 입력 화면이 막히는 셈이었다.
        st.session_state.setdefault("weather_work_days", max(1, default_work_days))
        # max_value를 동적으로 설정 (값이 크면 max도 자동으로 늘림)
        max_val = max(10000, default_work_days + 1000,
                      int(st.session_state["weather_work_days"]))
        work_days = st.number_input(
            "순작업일수",
            min_value=1,
            max_value=max_val,
            key="weather_work_days",
            help="TAB '엑셀 내역서 인식'에서 자동 계산된 값 (재계산 시 자동 갱신)"
        )
        st.session_state["work_days_input"] = work_days
        if "work_result" in st.session_state and default_work_days <= 0:
            st.caption("⚠️ 산정된 순작업일수가 0일입니다. '수동입력 관리' 탭에서 "
                       "매칭 안 된 항목의 1일 작업량을 입력하세요.")

    # 순작업일수가 바뀌었는데 이전 계산 결과가 남아 있으면 안내.
    # 수동입력·조수 변경으로 작업일수가 달라져도 '비작업일수 계산' 버튼을 다시
    # 누르기 전까지는 총 공사기간이 옛 값으로 표시되어 "안 바뀐다"처럼 보인다.
    _prev_res = st.session_state.get("weather_result") or {}
    _prev_wd = _prev_res.get("work_days")
    if _prev_wd is not None and int(_prev_wd) != int(work_days):
        st.warning(
            f"⚠️ 순작업일수가 **{int(_prev_wd)}일 → {int(work_days)}일**로 바뀌었습니다. "
            "아래 **비작업일수 계산** 버튼을 다시 눌러야 총 공사기간에 반영됩니다."
        )
    
    st.markdown("---")
    
    # ──────────────────────────────────
    # 4. 계산 버튼
    # ──────────────────────────────────
    if st.button("📊 비작업일수 계산", type="primary", width="stretch"):
        from datetime import datetime as dt
        
        # datetime 변환
        start_dt = dt.combine(start_date, dt.min.time())
        
        # 준비·정리·시운전은 달력 기준 기간이다(작업일수 산정 대상이 아니라 공사기간 구성요소).
        _prep_days = int(round(prep_months * 30.4))
        _wrap_days = int(round(wrapup_months * 30.4))
        _comm_days = int(round(commission_months * 30.4))

        # 공사기간(착공~준공) = 준비 → 본공사(작업+비작업) → 시운전 → 정리.
        # 비작업일수는 본공사가 실제로 진행되는 구간으로 계산해야 계절이 맞고,
        # 준공일-착공일도 총 공사기간과 일치한다. 예전에는 본공사를 착공일에 바로
        # 시작시키고 총 공사기간에만 준비기간을 더해서, 화면의 착공~준공 날짜 차이가
        # 총 공사기간보다 준비기간만큼 짧게 나왔다.
        work_start = start_dt + timedelta(days=_prep_days)

        # 종료일 추정: 순작업일수 * 1.5
        rough_end_date = work_start + timedelta(days=int(work_days * 1.5))
        
        # 반복 계산: 정확한 종료일 찾기
        for _ in range(5):
            # 1. 기상조건 비작업일수 (A)
            if HAS_GUIDELINE_WEATHER and selected_station:
                # 조건을 하나도 고르지 않았으면 기상 비작업일수는 0이어야 한다.
                # get_weather_non_work_days는 빈 목록을 받으면 기본 조건으로 대체하므로,
                # 없는 키를 넘겨 '조건 없음'(월별 행은 유지, 값은 0)으로 계산시킨다.
                _wres = get_weather_non_work_days(
                    selected_station, work_start, rough_end_date,
                    conditions=st.session_state.get("weather_conditions") or ["__none__"],
                )
                weather_days = _wres["total"]
                st.session_state["weather_detail"] = _wres
            else:
                weather_days = get_total_non_work_days(
                    selected_region,
                    work_start,
                    rough_end_date,
                    check_rain=include_rain,
                    check_cold=include_cold,
                    check_hot=include_hot
                )
            
            if isinstance(weather_days, dict):
                weather_days = weather_days.get("total", 0)
            
            # 2. 가이드라인 공식 적용 (A + B - C)
            result = get_total_non_work_days_with_holidays(
                weather_days,
                work_start,
                rough_end_date,
                include_holidays=include_holidays,
                min_weekly_rest=min_weekly_rest
            )
            
            non_work_days = result["total"]
            
            # 본공사 종료일 = 본공사 시작일 + 순작업일수 + 비작업일수
            calculated_end = work_start + timedelta(days=int(work_days + non_work_days - 1))
            
            # 수렴 체크
            if abs((rough_end_date - calculated_end).days) <= 1:
                break
            rough_end_date = calculated_end
        
        # 본공사 종료 후 시운전·정리기간이 이어진다(준비기간은 위에서 착공일 뒤에 배치).
        completion_date = calculated_end + timedelta(days=_comm_days + _wrap_days)
        total_days = (completion_date - start_dt).days + 1
        
        # 결과 저장
        st.session_state["weather_result"] = {
            "region": selected_region,
            "start_date": start_dt,
            "end_date": completion_date,
            "total_days": total_days,
            "prep_days": _prep_days,
            "wrapup_days": _wrap_days,
            "commission_days": _comm_days,
            "work_start": work_start,
            "work_end": calculated_end,
            "station": selected_station,
            # 가이드라인 방식이면 실제 적용한 조건 목록, 구버전 방식이면 None
            "conditions": (list(st.session_state.get("weather_conditions") or [])
                           if (HAS_GUIDELINE_WEATHER and selected_station) else None),
            "work_days": work_days,
            "non_work_days": non_work_days,
            "weather_days": result["weather"],
            "holiday_days": result["holidays"],
            "overlap_days": result["overlap"],
            "formula": result["formula"],
            "include_rain": include_rain,
            "include_cold": include_cold,
            "include_hot": include_hot,
            "include_holidays": include_holidays,
            "min_weekly_rest": min_weekly_rest,
        }
        
        st.success(f"✅ 준공일: **{completion_date.strftime('%Y년 %m월 %d일')}**")
        # 공기산정 탭(tab1)이 이 탭보다 먼저 실행되므로, 방금 저장한 weather_result를
        # 같은 실행에서 못 읽는다. 즉시 rerun 해서 모든 탭이 새 값으로 다시 렌더되게 한다.
        st.rerun()
    
    # ──────────────────────────────────
    # 5. 결과 표시
    # ──────────────────────────────────
    if "weather_result" in st.session_state:
        result = st.session_state["weather_result"]
        
        st.markdown("### 📊 계산 결과")
        
        # 메인 메트릭
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        col_m1.metric("📍 지역", result["region"])
        _extra_m = (result.get("prep_days", 0) + result.get("wrapup_days", 0)
                    + result.get("commission_days", 0))
        _lbl_m = "📅 총 공사기간" + (" (준비·시운전·정리 포함)" if _extra_m else "")
        col_m2.metric(_lbl_m, f"{result['total_days']}일",
                      delta=days_to_months_text(result['total_days']), delta_color="off")
        col_m3.metric("💼 순작업일수", f"{result['work_days']}일")
        col_m4.metric("🚫 비작업일수", f"{result['non_work_days']}일")
        
        # 비작업일수 세부 (A + B - C)
        st.markdown("#### 🧮 비작업일수 세부 (가이드라인 공식)")
        col_a, col_b, col_c, col_d = st.columns(4)
        col_a.metric("🌤️ 기상조건 (A)", f"{result.get('weather_days', 0)}일")
        col_b.metric("📅 법정공휴일 (B)", f"{result.get('holiday_days', 0)}일")
        col_c.metric("⚠️ 중복 (C)", f"{result.get('overlap_days', 0)}일")
        col_d.metric("📌 공식", f"A+B-C = {result.get('non_work_days', 0)}일")
        
        st.caption(f"📐 계산식: {result.get('formula', '')}")

        # 공사기간 구성 내역 (가이드라인: 준비 + 작업 + 비작업 + 시운전 + 정리)
        _pd = result.get("prep_days", 0)
        _wd_ = result.get("wrapup_days", 0)
        _cd = result.get("commission_days", 0)
        if _pd or _wd_ or _cd:
            st.markdown("**🗓️ 공사기간 구성**")
            _parts = [
                f"준비 {_pd}일" if _pd else None,
                f"작업 {result.get('work_days', 0)}일",
                f"비작업 {result.get('non_work_days', 0)}일",
                f"시운전 {_cd}일" if _cd else None,
                f"정리 {_wd_}일" if _wd_ else None,
            ]
            _txt = " + ".join(p for p in _parts if p)
            st.info(
                f"{_txt} = **{result.get('total_days', 0)}일** "
                f"{days_to_months_text(result.get('total_days', 0))}"
            )
        
        # 적용 조건
        _cond_md4 = "\n        ".join(f"- {l}" for l in applied_condition_lines(result))
        _ws4 = result.get("work_start", result["start_date"])
        _we4 = result.get("work_end", result["end_date"])
        st.info(f"""
        **적용된 조건:**
        {_cond_md4}
        - {'✅' if result.get('include_holidays', False) else '❌'} 법정공휴일
        - {'✅' if result.get('min_weekly_rest', False) else '❌'} 주 40시간 근무제 보장
        - **공사기간(착공~준공)**: {result['start_date'].strftime('%Y-%m-%d')} ~ {result['end_date'].strftime('%Y-%m-%d')}
        - **본공사 기간(비작업일수 산정 구간)**: {_ws4.strftime('%Y-%m-%d')} ~ {_we4.strftime('%Y-%m-%d')}
        """)
        
        # 월별 통합 상세 표
        try:
            from datetime import datetime as dt
            import pandas as pd
            
            from calendar import monthrange

            # 비작업일수를 실제로 산정한 본공사 구간 기준으로 표시한다.
            _ws_m = result.get("work_start", result["start_date"])
            _we_m = result.get("work_end", result["end_date"])
            monthly_holidays = get_holiday_breakdown_monthly(_ws_m, _we_m)
            holiday_dict = {h["월"]: h["법정공휴일"] for h in monthly_holidays} if monthly_holidays else {}

            _wdet = st.session_state.get("weather_detail") or {}
            _by_cond = _wdet.get("by_condition") or {}
            _monthly_w = _wdet.get("monthly") or []
            _guideline_mode = result.get("conditions") is not None or bool(_monthly_w)

            # (월, 대상일수, 달력일수, 기상 A, 추가 열)
            _months = []
            monthly_weather = []
            if _guideline_mode:
                # 가이드라인 방식: 계산 때 쓴 조건별 월 값(weather_detail)을 그대로 쓴다.
                # 예전에는 여기서 구버전 weather_data를 지역명으로 조회했는데, 지점의
                # 시·도명('강원도')이 구버전 키('강원')와 달라 기상일수가 전부 0으로 나왔다.
                for _mw in _monthly_w:
                    _y, _mo = map(int, _mw["월"].split("-"))
                    _dim = monthrange(_y, _mo)[1]
                    _months.append((_mw["월"], int(_mw.get("일수", _dim)), _dim,
                                    float(_mw.get("합계", 0) or 0), {}))
            else:
                monthly_weather = get_monthly_breakdown(
                    result["region"], _ws_m, _we_m,
                    check_rain=result["include_rain"],
                    check_cold=result["include_cold"],
                    check_hot=result["include_hot"],
                )
                for m in monthly_weather:
                    _y, _mo = map(int, m["month"].split("-"))
                    _dim = monthrange(_y, _mo)[1]
                    rain, cold, hot = m.get("rain", 0), m.get("cold", 0), m.get("hot", 0)
                    _months.append((m["month"], _dim, _dim, rain + cold + hot, {
                        "🌧️ 강우": f"{rain:.1f}", "❄️ 한랭": f"{cold:.1f}", "🔥 폭염": f"{hot:.1f}",
                    }))

            if _months:
                st.markdown("### 📅 월별 비작업일수 상세")
                st.caption(
                    "📖 본공사 구간의 월별 기상조건(A) + 공휴일(B) − 중복일수(C). "
                    "첫·마지막 달은 해당 일수만큼 안분했고, 월별 반올림 때문에 합계가 "
                    "전체 공식 결과와 1~2일 다를 수 있습니다."
                )

                monthly_data = []
                for month_str, _days, cal_days, weather_total, _extra_cols in _months:
                    _h_full = holiday_dict.get(month_str, 0) if result.get("include_holidays", False) else 0
                    holidays = round(_h_full * _days / cal_days) if cal_days else 0  # B (부분 월 안분)
                    # 중복일수 (C = A × B ÷ 대상일수)
                    overlap = round(weather_total * holidays / _days) if _days > 0 else 0
                    month_non_work = round(weather_total + holidays - overlap)
                    # 주 40시간 근무제 보장
                    min_rest = round(_days / 7)
                    if result.get("min_weekly_rest", False) and month_non_work < min_rest:
                        month_non_work = min_rest
                    monthly_data.append({
                        "월": month_str,
                        **_extra_cols,
                        "🌤️ 기상(A)": f"{weather_total:.1f}",
                        "📅 공휴일(B)": f"{holidays}",
                        "⚠️ 중복(C)": f"{overlap}",
                        "📊 비작업": f"{month_non_work}",
                        "📆 대상일수": f"{_days}",
                    })

                df_monthly = pd.DataFrame(monthly_data)
                st.dataframe(df_monthly, hide_index=True, width="stretch")
                
                # ──────────────────────────────────
                # ──────────────────────────────────
                # 항목별 분석
                # 가이드라인 방식(13개 조건)으로 계산하면 조건 구성이 프로젝트마다
                # 달라지므로, 예전처럼 강우/한랭/폭염 3개 키를 고정으로 읽으면
                # 모두 0으로 표시된다. weather_detail의 조건별 결과를 그대로 쓴다.
                # ──────────────────────────────────
                st.markdown("### 📈 항목별 비작업일수 분석")

                # 법정공휴일은 공식의 B(부분 월 안분)와 같은 값을 쓴다. 월 전체 공휴일수를
                # 단순 합산하면 위 B 지표와 숫자가 달라 보였다(실측 B 255일 vs 266일).
                total_holiday = result.get("holiday_days", 0)

                if _guideline_mode:
                    _labels, _values = [], []
                    for _ck, _cv in _by_cond.items():
                        _labels.append(CONDITION_LABELS.get(_ck, _ck))
                        _values.append(round(_cv, 1))
                    _labels.append("법정공휴일")
                    _values.append(total_holiday)
                    _labels.append("중복일수(차감)")
                    _values.append(result.get("overlap_days", 0))

                    df_chart = pd.DataFrame({"항목": _labels, "일수": _values})
                    col_chart1, col_chart2 = st.columns([2, 1])
                    with col_chart1:
                        st.bar_chart(df_chart.set_index("항목"))
                    with col_chart2:
                        st.markdown("**📊 합계**")
                        for _ck, _cv in _by_cond.items():
                            st.metric(CONDITION_LABELS.get(_ck, _ck), f"{_cv:.1f}일")
                        st.metric("📅 법정공휴일", f"{total_holiday}일")

                    # 조건별 월별 상세
                    if _monthly_w:
                        with st.expander("🌤️ 기상조건 월별 상세", expanded=False):
                            st.dataframe(pd.DataFrame(_monthly_w), hide_index=True, width="stretch")
                            st.caption(
                                f"📌 {st.session_state.get('selected_station', result.get('region',''))} "
                                "관측지점 기준 (가이드라인 부록3, 2015~2024 월평균)"
                            )
                else:
                    # 구버전 weather_data 경로(가이드라인 데이터 미적용)
                    total_rain = sum(m.get("rain", 0) for m in monthly_weather)
                    total_cold = sum(m.get("cold", 0) for m in monthly_weather)
                    total_hot = sum(m.get("hot", 0) for m in monthly_weather)
                    df_chart = pd.DataFrame({
                        "항목": ["🌧️ 강우일", "❄️ 한랭일", "🔥 폭염일", "📅 법정공휴일", "⚠️ 중복일수"],
                        "일수": [total_rain, total_cold, total_hot, total_holiday,
                                result.get("overlap_days", 0)],
                    })
                    col_chart1, col_chart2 = st.columns([2, 1])
                    with col_chart1:
                        st.bar_chart(df_chart.set_index("항목"))
                    with col_chart2:
                        st.markdown("**📊 합계**")
                        st.metric("🌧️ 강우일", f"{total_rain:.1f}일")
                        st.metric("❄️ 한랭일", f"{total_cold:.1f}일")
                        st.metric("🔥 폭염일", f"{total_hot:.1f}일")
                        st.metric("📅 법정공휴일", f"{total_holiday}일")

                with st.expander("📅 법정공휴일 월별 상세", expanded=False):
                    if monthly_holidays:
                        holiday_detail = [{"월": h["월"], "공휴일수": f"{h['법정공휴일']}일"} for h in monthly_holidays]
                        st.dataframe(pd.DataFrame(holiday_detail), hide_index=True, width="stretch")
                        st.caption("📌 부록1: 일요일(52일) + 명절 + 국경일 + 기타 공휴일 + 대체공휴일")
                    else:
                        st.info("법정공휴일이 포함되지 않았습니다.")
                
                # 가이드라인 공식 설명
                if _guideline_mode:
                    _conds_f = result.get("conditions")
                    if _conds_f is None:
                        _conds_f = list(_by_cond)
                    _a_desc = " + ".join(CONDITION_LABELS.get(c, c) for c in _conds_f) or "적용 조건 없음"
                    _src_desc = (f"가이드라인 부록3 기상청 관측자료 2015~2024 월평균 "
                                 f"({result.get('station') or result['region']} 관측지점)")
                else:
                    _a_desc = "강우 + 한랭 + 폭염"
                    _src_desc = "기상청 평년값 (1991-2020)"
                with st.expander("📐 계산 공식 설명", expanded=False):
                    st.markdown(f"""
                    ### 가이드라인 19페이지 공식
                    
                    **비작업일수 = A + B - C**
                    
                    | 항목 | 내용 | 값 |
                    |------|------|-----|
                    | **A** | 기상조건 비작업일수 ({_a_desc}) | {result.get('weather_days', 0)}일 |
                    | **B** | 법정 공휴일수 (부록1 기준) | {result.get('holiday_days', 0)}일 |
                    | **C** | 중복일수 = A × B ÷ 달력일수 (소수점 반올림) | {result.get('overlap_days', 0)}일 |
                    | **계** | A + B - C | **{result.get('non_work_days', 0)}일** |
                    
                    ### 주 40시간 근무제 보장
                    - 월별 비작업일수가 주 40시간 근무제 일수보다 작을 경우, 보정 적용
                    - 일반적으로 주 1일 휴식 보장 (월 4~5일)
                    
                    ### 데이터 출처
                    - **기상 데이터**: {_src_desc}
                    - **법정공휴일**: 「관공서의 공휴일에 관한 규정」 (부록1, 2026-2035)
                    """)
        except Exception as e:
            st.error(f"월별 상세 정보 표시 오류: {e}")
            import traceback
            st.code(traceback.format_exc())
    
    # ──────────────────────────────────
    # 6. 지역 기상 정보 미리보기
    # ──────────────────────────────────
    _stat_name = selected_station if (HAS_GUIDELINE_WEATHER and selected_station) else f"{selected_region} 지역"
    with st.expander(f"📊 {_stat_name} 연간 기상 통계", expanded=False):
        if HAS_GUIDELINE_WEATHER and selected_station:
            # 가이드라인 방식: 선택한 조건의 관측지점 월평균 비작업일수.
            # (구버전 RAIN_DAYS는 '강원' 같은 짧은 키라 '강원도 ○○' 지점에서는 표가 비었다)
            _conds_s = st.session_state.get("weather_conditions") or []
            if _conds_s:
                _stat = {"월": [f"{m}월" for m in range(1, 13)]}
                for _c in _conds_s:
                    _stat[CONDITION_LABELS.get(_c, _c)] = list(
                        WEATHER_NON_WORK.get(_c, {}).get(selected_station) or [0.0] * 12)
                df_stats = pd.DataFrame(_stat)
                df_stats["합계"] = df_stats.drop(columns=["월"]).sum(axis=1).round(1)
                st.dataframe(df_stats, hide_index=True, width="stretch")
                st.metric("연간 합계 (선택 조건 단순 합산)", f"{df_stats['합계'].sum():.1f}일")
                st.caption("가이드라인 부록3, 2015~2024 월평균. 조건 간 중복은 보정하지 않고 단순 합산합니다.")
            else:
                st.info("선택된 기상조건이 없습니다.")
        elif selected_region in RAIN_DAYS:
            import pandas as pd
            
            months = list(range(1, 13))
            data = {
                "월": [f"{m}월" for m in months],
                "🌧️ 강우일": [RAIN_DAYS[selected_region].get(m, 0) for m in months],
                "❄️ 한랭일": [COLD_DAYS[selected_region].get(m, 0) for m in months],
                "🔥 폭염일": [HOT_DAYS[selected_region].get(m, 0) for m in months],
            }
            df_stats = pd.DataFrame(data)
            df_stats["합계"] = df_stats["🌧️ 강우일"] + df_stats["❄️ 한랭일"] + df_stats["🔥 폭염일"]
            
            st.dataframe(df_stats, hide_index=True, width="stretch")
            
            # 연간 합계
            annual_rain = sum(RAIN_DAYS[selected_region].values())
            annual_cold = sum(COLD_DAYS[selected_region].values())
            annual_hot = sum(HOT_DAYS[selected_region].values())
            
            col_s1, col_s2, col_s3, col_s4 = st.columns(4)
            col_s1.metric("연간 강우일", f"{annual_rain:.1f}일")
            col_s2.metric("연간 한랭일", f"{annual_cold:.1f}일")
            col_s3.metric("연간 폭염일", f"{annual_hot:.1f}일")
            col_s4.metric("연간 총합", f"{annual_rain + annual_cold + annual_hot:.1f}일")

# ══════════════════════════════════════════════════════════════
# TAB 5: 예정공정표
# ══════════════════════════════════════════════════════════════
with tab5:
    st.subheader("📅 예정공정표")
    st.caption("산정된 주공정별 작업일수를 달력 기간으로 환산해 월 단위 공정표를 생성합니다")

    _gr = st.session_state.get('grouped_results')
    _wr = st.session_state.get('weather_result')

    if not _gr:
        st.warning("먼저 '엑셀 내역서 인식' 탭에서 내역서를 업로드하고 계산을 완료해주세요.")
    elif not _wr:
        st.warning("먼저 '비작업일수 계산기' 탭에서 착공일과 비작업일수를 계산해주세요. (착공일·달력환산에 필요)")
    else:
        import math as _math
        from datetime import timedelta as _td, date as _date

        def _add_months(d, n):
            """d로부터 n개월 후의 (연, 월). 일 성분은 무시(월 단위 공정표라 불필요)."""
            t = (d.year * 12 + (d.month - 1)) + n
            return t // 12, t % 12 + 1

        # ── 1) 주공정별 작업일수 수집 (선택된 주공정만) ──
        _sel = set(st.session_state.get('selected_major') or [])
        _rows_all = [r for g in _gr.values() for r in g]
        _sched_rows = [r for r in _rows_all if (not _sel) or (r.get('공종명_pure') in _sel)]
        # 자재·운반류는 시공 공정이 아니므로 공정표에서 제외 (CP 분석과 동일 기준)
        _sched_rows = [r for r in _sched_rows if not is_non_work_category(r.get('공종명_pure') or '')]

        # 달력 환산: 앱의 공기 모델과 일관되게 배율을 잡는다.
        # - 최장(병행) 모드: 최장 공종이 본공사 기간을 채우고 나머지는 비례 배분
        # - 합산(순차) 모드: 선택 공종 작업일수의 '합'이 본공사 기간을 채우도록 배분
        # total_days에는 준비·시운전·정리가 이미 들어 있으므로 공정 막대는 본공사 기간
        # (작업+비작업)에만 배분하고, 준비·시운전·정리는 아래에서 별도 행으로 붙인다.
        # (예전에는 total_days 전체로 배분한 뒤 준비 4·3개월, 시운전 2개월 행을
        #  고정값으로 또 붙여서 공정표가 산정 공기보다 길어졌다)
        _total_days = _wr.get('total_days') or 0
        _work_days = _wr.get('work_days') or 1
        _prep_d = int(_wr.get('prep_days', 0) or 0)
        _comm_d = int(_wr.get('commission_days', 0) or 0)
        _wrap_d = int(_wr.get('wrapup_days', 0) or 0)
        _constr_days = max(1, _total_days - _prep_d - _comm_d - _wrap_d)
        _start_date = _wr.get('start_date')
        _start_d = _start_date.date() if hasattr(_start_date, 'date') else _start_date

        _is_sum_mode = st.session_state.get("combine_mode", "최장(병행)").startswith("합산")
        _max_wd = max((int(r.get('작업일수(일)', 0) or 0) for r in _sched_rows), default=1) or 1
        _sum_wd = sum(int(r.get('작업일수(일)', 0) or 0) for r in _sched_rows) or 1
        _base_wd = _sum_wd if _is_sum_mode else _max_wd
        _scale = _constr_days / _base_wd  # 작업일수 → 달력일수 배율

        st.info(
            f"착공일 {_start_d} · 총공사기간 {_total_days}일 "
            f"(준비 {_prep_d} + 본공사 {_constr_days} + 시운전 {_comm_d} + 정리 {_wrap_d}) · "
            f"본공사를 {'선택 공종 합산' if _is_sum_mode else '최장 주공정'} {_base_wd}일 기준 배분 "
            f"(배율 {_scale:.2f})"
        )

        # ── 2) 표준 시퀀스 순서 + 순차(계단식) 시작월 자동 제안 ──
        # 샘플 예정공정표처럼 토공→관로→구조물→포장이 계단식으로 이어지도록,
        # 각 공정은 '선행 공정이 일정 비율(25%) 진행된 시점'에 시작하는 것으로 제안.
        # 특수 의존성: 배수설비공은 관로류(관로/관접합/부설) '완료 후' 시작 (현장 연결 순서).
        _SEQ = ["가시설", "토공", "추진", "관로", "관접합", "부설", "구조물", "포장", "부대", "배수", "기타"]
        def _seq_key(name):
            for i, kw in enumerate(_SEQ):
                if kw in (name or ""):
                    return i
            return len(_SEQ)

        _sched_rows = sorted(_sched_rows, key=lambda r: _seq_key(r.get('공종명_pure', '')))

        # ── 표시 단위 선택: 대공종 요약 vs 라인/구간별 세부 ──
        _detail_mode = st.radio(
            "공정표 표시 단위",
            options=["대공종 요약", "라인·구간별 세부"],
            index=1,
            horizontal=True,
            key="sched_detail_mode",
            help=(
                "대공종 요약: 토공·관로공처럼 공종당 막대 1개.\n\n"
                "라인·구간별 세부: A-LINE/B-LINE-주간/야간, 지구Ⅰ·Ⅱ처럼 구간마다 막대를 나눠 "
                "실제 예정공정표처럼 표기합니다(구간명·투입조수·작업일수 표시)."
            ),
        )
        _is_detail = _detail_mode.startswith("라인")

        # 배치 시차: 합산(순차) 모드면 겹침 없이 앞 공정 종료 후 착수, 아니면 25% 계단식
        _OVERLAP = 1.0 if _is_sum_mode else 0.25
        _prefill = []
        _prev_start = 1
        _prev_months = 0
        _pipe_end = None  # 관로류 종료 개월차 (배수설비 의존성용)
        _NON_WORK_GROUPS = ("공사준비", "시운전", "준공정리")

        # 공사준비 행 — 비작업일수 탭의 준비기간(착공 직후 인허가·용지보상·현장사무실 등).
        # 본공사는 준비기간이 끝난 다음 달부터 시작한다.
        _prep_m = int(round(_prep_d / 30.4))
        _work_start_m = 1 + _prep_m
        if _prep_d > 0:
            _prefill.append({"구분": "공사준비", "공종": "공사준비(인허가·용지보상·가설시설 등)", "조수": 0,
                             "시작(개월차)": 1, "기간(개월)": max(1, _prep_m), "작업일수": _prep_d})

        for r in _sched_rows:
            _nm = r.get('공종명_pure') or r.get('공종', '')
            _wd_total = int(r.get('작업일수(일)', 0) or 0)
            if _wd_total <= 0:
                continue
            _crew = int(r.get('crew', DEFAULT_CREW) or DEFAULT_CREW)

            _lines = r.get("_라인별누적") or {}
            if _is_detail and len(_lines) > 1:
                _units = [(f"{_nm} ({ln})", int(d)) for ln, d in
                          sorted(_lines.items(), key=lambda x: -x[1]) if int(d) > 0]
            else:
                _units = [(_nm, _wd_total)]

            _grp_start = None
            for _label, _wd in _units:
                _months = max(1, int(_math.ceil((_wd * _scale) / 30.4)))

                if any(kw in _label for kw in ("배수",)) and _pipe_end is not None:
                    _start = _pipe_end + 1
                elif not any(p["구분"] not in _NON_WORK_GROUPS for p in _prefill):
                    _start = _work_start_m
                elif _grp_start is not None:
                    # 같은 공종의 라인들은 서로 병행 → 동일 시작월
                    _start = _grp_start
                else:
                    _start = _prev_start + max(1, int(_math.ceil(_prev_months * _OVERLAP)))

                if _grp_start is None:
                    _grp_start = _start

                _prefill.append({"구분": _nm, "공종": _label, "조수": _crew,
                                 "시작(개월차)": _start, "기간(개월)": _months, "작업일수": _wd})

                if any(kw in _label for kw in ("관로", "관접합", "부설")):
                    _e = _start + _months - 1
                    _pipe_end = _e if _pipe_end is None else max(_pipe_end, _e)

            if _units:
                _longest = max(max(1, int(_math.ceil((w * _scale) / 30.4))) for _, w in _units)
                _prev_start, _prev_months = _grp_start, _longest

        # 시운전·준공정리 행 — 비작업일수 탭에 입력한 기간 그대로 본공사 뒤에 붙인다.
        # (예전에는 공종명에 '구조물·처리·설비' 등이 있으면 시운전 2개월을 고정으로 붙여,
        #  입력한 시운전 기간과 무관하게 공정표가 늘어났다)
        _work_rows = [p for p in _prefill if p["구분"] not in _NON_WORK_GROUPS]
        _end_all = max((p["시작(개월차)"] + p["기간(개월)"] - 1 for p in _work_rows),
                       default=_work_start_m - 1)
        for _grp_nm, _label_nm, _dd in (("시운전", "시운전/시설인계", _comm_d),
                                        ("준공정리", "준공정리(현장정리·준공서류)", _wrap_d)):
            if _dd > 0:
                _mm = max(1, int(round(_dd / 30.4)))
                _prefill.append({"구분": _grp_nm, "공종": _label_nm, "조수": 0,
                                 "시작(개월차)": _end_all + 1, "기간(개월)": _mm, "작업일수": _dd})
                _end_all += _mm

        if not _prefill:
            st.warning("주공정으로 선택된 공종 중 작업일수가 있는 항목이 없습니다.")
        else:
            import pandas as _pd
            st.markdown("#### ✏️ 공정 배치 (자동 제안값, 직접 수정 가능)")
            st.caption("시작(개월차)=착공 후 몇 번째 달부터, 기간(개월)=지속 개월수. 보상·인허가 등 프로젝트 사정에 맞게 조정하세요.")
            _edited = st.data_editor(
                _pd.DataFrame(_prefill),
                hide_index=True,
                width="stretch",
                num_rows="dynamic",
                column_config={
                    "구분": st.column_config.TextColumn(),
                    "공종": st.column_config.TextColumn(required=True),
                    "조수": st.column_config.NumberColumn(min_value=0, step=1),
                    "시작(개월차)": st.column_config.NumberColumn(min_value=1, step=1, required=True),
                    "기간(개월)": st.column_config.NumberColumn(min_value=1, step=1, required=True),
                    "작업일수": st.column_config.NumberColumn(),
                },
                key="schedule_editor",
            )

            _proj_name = st.text_input("공사명", value="상하수도 공사", key="sched_project_name")

            _max_month = int((_edited["시작(개월차)"] + _edited["기간(개월)"] - 1).max())
            _base_months = int(_math.ceil(_total_days / 30.4))
            if _max_month > _base_months:
                st.caption(f"ℹ️ 순차 배치·의존성(배수설비는 관로 완료 후) 반영으로 전체 {_max_month}개월 — 병행 기준 공기({_base_months}개월)보다 깁니다. 표에서 시작월을 앞당겨 조정할 수 있습니다.")

            # ── 3) 화면 미리보기 (plotly 간트) ──
            try:
                import plotly.express as _px
                _gantt = []
                for _, _row in _edited.iterrows():
                    _sy, _sm2 = _add_months(_start_d, int(_row["시작(개월차)"]) - 1)
                    _ey, _em2 = _add_months(_start_d, int(_row["시작(개월차)"]) - 1 + int(_row["기간(개월)"]))
                    _gantt.append({"공종": _row["공종"], "시작": _date(_sy, _sm2, 1), "종료": _date(_ey, _em2, 1)})
                _gdf = _pd.DataFrame(_gantt)
                _fig = _px.timeline(_gdf, x_start="시작", x_end="종료", y="공종", color="공종")
                _fig.update_yaxes(autorange="reversed")
                _fig.update_layout(showlegend=False, height=90 + 40 * len(_gdf), margin=dict(l=10, r=10, t=10, b=10))
                st.plotly_chart(_fig, width="stretch")
            except Exception as _e:
                st.caption(f"미리보기 생략: {_e}")

            # ── 4) 엑셀 예정공정표 생성 (업로드 양식과 동일 구조: 월당 2열, 연/월 2단 헤더) ──
            if st.button("📥 예정공정표 엑셀 생성", type="primary", width="stretch"):
                try:
                    from openpyxl import Workbook
                    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
                    from openpyxl.utils import get_column_letter
                    from io import BytesIO

                    wb_s = Workbook()
                    ws_s = wb_s.active
                    ws_s.title = "예정공정표"

                    _thin = Side(style="thin", color="999999")
                    _border = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
                    _bar_fill = PatternFill("solid", fgColor="4F81BD")
                    _hdr_fill = PatternFill("solid", fgColor="DCE6F1")

                    COL0 = 3          # 공종명 열 (C) — B열은 구분(대공종) 표기용
                    MONTH_W = 2       # 월당 열 수 (양식과 동일)
                    FIRST_MC = 4      # 첫 월 시작 열 (D)

                    # 제목
                    ws_s.cell(row=1, column=1, value=f"{_proj_name} 예정공정표").font = Font(size=14, bold=True)

                    # 연/월 2단 헤더
                    _ycells = {}
                    for m in range(_max_month):
                        _y, _mm = _add_months(_start_d, m)
                        c = FIRST_MC + m * MONTH_W
                        ws_s.merge_cells(start_row=3, start_column=c, end_row=3, end_column=c + MONTH_W - 1)
                        mc = ws_s.cell(row=3, column=c, value=_mm)
                        mc.alignment = Alignment(horizontal="center")
                        mc.fill = _hdr_fill
                        mc.border = _border
                        _ycells.setdefault(_y, []).append(c)
                    for _y, _cols in _ycells.items():
                        c1, c2 = min(_cols), max(_cols) + MONTH_W - 1
                        ws_s.merge_cells(start_row=2, start_column=c1, end_row=2, end_column=c2)
                        yc = ws_s.cell(row=2, column=c1, value=f"{_y}년")
                        yc.alignment = Alignment(horizontal="center")
                        yc.font = Font(bold=True)
                        yc.fill = _hdr_fill
                        yc.border = _border

                    ws_s.merge_cells(start_row=2, start_column=COL0, end_row=3, end_column=COL0)
                    hc = ws_s.cell(row=2, column=COL0, value="공종")
                    hc.alignment = Alignment(horizontal="center", vertical="center")
                    hc.font = Font(bold=True)
                    hc.fill = _hdr_fill
                    hc.border = _border

                    # 공정 막대 (공종당 2행: 막대행 + 여백행 — 실제 예정공정표 양식과 유사)
                    # 막대 안에는 샘플처럼 "공종명_N조"와 작업일수를 함께 표기한다.
                    _prep_fill = PatternFill("solid", fgColor="9BBB59")  # 공사준비 행 구분색
                    _r = 4
                    _last_group = None
                    for _, _row in _edited.iterrows():
                        _nm = str(_row["공종"])
                        _grp = str(_row.get("구분", "") or "")
                        _crew_v = int(_row.get("조수", 0) or 0)
                        _wd_v = int(_row.get("작업일수", 0) or 0)
                        _sm = int(_row["시작(개월차)"])
                        _dm = int(_row["기간(개월)"])

                        # 구분(대공종)이 바뀌면 좌측에 구분명을 한 번 표기
                        _left = _nm if _grp in ("", _nm) else f"  {_nm}"
                        if _grp and _grp != _last_group:
                            ws_s.cell(row=_r, column=COL0 - 1, value=_grp).font = Font(bold=True, size=9)
                            _last_group = _grp

                        nc = ws_s.cell(row=_r, column=COL0, value=_left)
                        nc.alignment = Alignment(vertical="center")
                        nc.border = _border

                        _bar_txt = _nm if _crew_v <= 0 else f"{_nm}_{_crew_v}조"
                        if _wd_v > 0:
                            _bar_txt += f"  {_wd_v}일"

                        c1 = FIRST_MC + (_sm - 1) * MONTH_W
                        c2 = FIRST_MC + (_sm - 1 + _dm) * MONTH_W - 1
                        ws_s.merge_cells(start_row=_r, start_column=c1, end_row=_r, end_column=c2)
                        bar = ws_s.cell(row=_r, column=c1, value=_bar_txt)
                        bar.fill = _prep_fill if _grp in _NON_WORK_GROUPS else _bar_fill
                        bar.font = Font(color="FFFFFF", size=9)
                        bar.alignment = Alignment(horizontal="center", vertical="center")
                        for cc in range(c1, c2 + 1):
                            ws_s.cell(row=_r, column=cc).border = _border
                        _r += 2

                    # 격자(빈 월칸 테두리) + 열폭
                    for rr in range(4, _r):
                        for m in range(_max_month):
                            for k in range(MONTH_W):
                                cell = ws_s.cell(row=rr, column=FIRST_MC + m * MONTH_W + k)
                                if cell.border is None or cell.border.left is None or cell.border.left.style is None:
                                    cell.border = _border
                    ws_s.column_dimensions[get_column_letter(COL0 - 1)].width = 14
                    ws_s.column_dimensions[get_column_letter(COL0)].width = 26
                    for m in range(_max_month * MONTH_W):
                        ws_s.column_dimensions[get_column_letter(FIRST_MC + m)].width = 3.2

                    _buf = BytesIO()
                    wb_s.save(_buf)
                    _buf.seek(0)
                    st.download_button(
                        label="📥 예정공정표 다운로드",
                        data=_buf,
                        file_name=f"예정공정표_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        width="stretch",
                    )
                    st.success("✅ 예정공정표가 생성되었습니다. 다운로드 버튼을 눌러주세요.")
                except Exception as e:
                    st.error(f"공정표 생성 실패: {e}")
                    import traceback
                    st.code(traceback.format_exc())
# ══════════════════════════════════════════════════════════════
# TAB 6: 수동입력 관리
# ══════════════════════════════════════════════════════════════
with tab6:
    st.subheader("📝 수동입력 관리")
    st.caption("매칭 안 된 항목들에 대해 1일 작업량을 직접 입력하세요")
    
    # session_state 초기화
    if "manual_rates" not in st.session_state:
        st.session_state["manual_rates"] = {}
    
    if "unmatched_all" not in st.session_state or not st.session_state["unmatched_all"]:
        st.info("📂 먼저 TAB 2에서 엑셀 파일을 업로드하세요!")
    else:
        unmatched_all = st.session_state["unmatched_all"]
        
        # 그룹별로 정리
        group_data = {
            "1.1": {"name": "🏗️ 하수관로공사", "categories": []},
            "1.2": {"name": "🔧 관로 부대공사", "categories": []},
            "2.1": {"name": "💧 배수설비공사", "categories": []},
            "2.2": {"name": "⚙️ 기계설비", "categories": []},
        }
        
        for cat_key, cat_data in unmatched_all.items():
            mk = cat_data["major_key"]
            if mk not in group_data:
                # 신양식 합성코드(90.x) 등 미등록 그룹은 동적 생성 (대공종명으로 라벨)
                group_data[mk] = {"name": f"📁 {cat_data.get('category', mk)}", "categories": []}
            group_data[mk]["categories"].append(cat_data)
        
        # 전체 통계
        total_unmatched = sum(
            sum(len(cat["items"]) for cat in gd["categories"])
            for gd in group_data.values()
        )
        saved_count = len(st.session_state["manual_rates"])
        
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.metric("전체 매칭 안 된 항목", f"{total_unmatched}개")
        with col_b:
            st.metric("저장된 항목", f"{saved_count}개", delta=f"{total_unmatched - saved_count}개 남음")
        with col_c:
            if st.button("🗑️ 모든 저장 초기화", width="stretch"):
                st.session_state["manual_rates"] = {}
                st.rerun()
        
        st.markdown("---")

        # ══════════════════════════════════════════════════════
        # 상위 2탭: 추정값 검토 / 미매칭 수동입력
        # 두 작업은 성격이 다르고 둘 다 중요하므로(검토를 놓치기 쉬움)
        # 접힌 expander 대신 대등한 탭으로 분리해 노출한다.
        # ══════════════════════════════════════════════════════
        _n_sr = sum(len(_c["items"]) for _c in st.session_state.get("series_review_all", {}).values())
        _n_lr = sum(len(_c["items"]) for _c in st.session_state.get("labor_review_all", {}).values())
        _n_um = sum(len(cat["items"]) for gd in group_data.values() for cat in gd["categories"])
        _main_tabs = st.tabs([
            f"🔎 추정값 검토 ({_n_sr + _n_lr})",
            f"✏️ 미매칭 수동입력 ({_n_um})",
        ])

        with _main_tabs[0]:
            st.caption(
                "일위대가에서 직접 값을 못 구해 **추정**한 항목입니다. "
                "설계자가 다른 품셈을 적용했다면 실제와 차이가 클 수 있으니, "
                "공기 영향이 큰 항목부터 확인하고 필요하면 직접 입력하세요."
            )
            if _n_sr + _n_lr == 0:
                st.success("추정값으로 계산된 항목이 없습니다.")
            # 값을 입력할 때마다 전체 재실행되면 느리므로 @st.fragment로 감싼다.
            # 이 영역만 다시 그려지고, 반영 버튼을 눌렀을 때만 전체가 갱신된다.
            @st.fragment
            def _review_section():
                # ══════════════════════════════════════════════════════
                # 🔎 동일계열 Q승계 항목 검토
                # 산근에 Q산식이 없어 같은 계열(동일 항목명·규격)의 Q값을 물려받은 항목들.
                # 추정값이므로 설계자가 한 번 확인하고, 필요하면 여기서 바로 수정한다.
                # ══════════════════════════════════════════════════════
                _sr = st.session_state.get("series_review_all", {})
                _sr_items = []
                for _k, _c in _sr.items():
                    for _it in _c["items"]:
                        _sr_items.append({**_it, "category": _c.get("category", "")})
                if _sr_items:
                    with st.expander(f"🔎 승계값 검토 필요 ({len(_sr_items)}개) — 같은 계열 Q값을 물려받은 항목", expanded=False):
                        st.caption(
                            "산근에 Q산식이 없어 노무비 역산으로는 과소평가되던 항목에, "
                            "시공조건이 같은 계열(동일 항목명·규격)의 Q값을 승계했습니다. "
                            "장비·조건이 실제로 동일한지 확인하고, 다르면 아래에서 1일 작업량을 직접 입력하세요. "
                            "입력하면 수동입력(0순위)으로 승계값을 덮어씁니다."
                        )
                        for _i, _it in enumerate(_sr_items):
                            _c1, _c2, _c3, _c4 = st.columns([2.6, 2, 1.4, 1.4])
                            with _c1:
                                st.text(f"[{_it.get('category','')}] {_it['name'][:26]}")
                            with _c2:
                                st.text(f"{_it.get('spec','')[:24]}")
                            with _c3:
                                st.text(f"승계 {_it.get('승계값','')}")
                            with _c4:
                                _mk_s = _it["manual_key"]
                                _cur = st.session_state["manual_rates"].get(_mk_s, {})
                                _pend_s = st.session_state.get("pending_rate", {}).get(_mk_s)
                                _val = st.number_input(
                                    "수정", min_value=0.0, step=0.1,
                                    value=float(_pend_s[0]) if _pend_s else float(_cur.get("daily", 0.0)),
                                    key=f"sr_{_i}_{_mk_s}", label_visibility="collapsed",
                                )
                                # 입력 즉시 반영하지 않고 대기 목록에 모은다(입력할 때마다
                                # 전체 재계산이 돌면 느리고, rerun이 없으면 결과표가 안 바뀐다).
                                st.session_state.setdefault("pending_rate", {})
                                _cur_v = float(_cur.get("daily", 0.0) or 0.0)
                                if abs(_val - _cur_v) > 1e-9:
                                    st.session_state["pending_rate"][_mk_s] = (_val, _it.get("unit", ""))
                                else:
                                    st.session_state["pending_rate"].pop(_mk_s, None)
                        st.caption("※ 0으로 두면 승계값을 그대로 사용합니다.")
                        _render_apply_rates("sr")
                    st.markdown("---")

                # ══════════════════════════════════════════════════════
                # 🔎 노무비역산 추정값 검토 (작업일수 큰 항목)
                # 일위대가에 인수가 없어 노무비로 되짚은 값이라 정밀도가 낮다.
                # 설계자가 다른 품셈을 적용했다면 크게 어긋날 수 있으므로
                # 공기 영향이 큰(작업일수 5일 이상) 항목부터 확인한다.
                # ══════════════════════════════════════════════════════
                _lr = st.session_state.get("labor_review_all", {})
                _lr_items = []
                for _k, _c in _lr.items():
                    for _it in _c["items"]:
                        _lr_items.append({**_it, "category": _c.get("category", "")})
                _lr_items.sort(key=lambda x: -int(x.get("작업일수", 0) or 0))
                if _lr_items:
                    _top = _lr_items[:40]
                    with st.expander(
                        f"🔎 역산 추정값 검토 ({len(_lr_items)}개 중 상위 {len(_top)}개) — 노무비로 되짚은 값",
                        expanded=False,
                    ):
                        st.caption(
                            "일위대가에 직접 인수(단위 '인')가 없어 노무비를 노임단가로 나눠 추정한 값입니다. "
                            "설계자가 별도 품셈이나 자체 기준을 적용한 항목은 실제와 차이가 클 수 있습니다. "
                            "작업일수가 큰 순서로 정렬했으니 공기에 영향이 큰 항목부터 확인하고, "
                            "필요하면 1일 작업량을 직접 입력하세요(수동입력이 0순위로 우선 적용됩니다)."
                        )
                        for _i4, _it4 in enumerate(_top):
                            _c1, _c2, _c3, _c4, _c5 = st.columns([2.4, 1.8, 1.0, 1.2, 1.2])
                            with _c1:
                                st.text(f"[{_it4.get('category','')}] {_it4['name'][:24]}")
                            with _c2:
                                st.text(f"{_it4.get('spec','')[:22]}")
                            with _c3:
                                st.text(f"{_it4.get('작업일수',0)}일")
                            with _c4:
                                st.text(f"추정 {_it4.get('추정값','')}")
                            with _c5:
                                _mk4 = _it4["manual_key"]
                                _cur4 = st.session_state["manual_rates"].get(_mk4, {})
                                _pend4 = st.session_state.get("pending_rate", {}).get(_mk4)
                                _val4 = st.number_input(
                                    "수정", min_value=0.0, step=0.1,
                                    value=float(_pend4[0]) if _pend4 else float(_cur4.get("daily", 0.0)),
                                    key=f"lr_{_i4}_{_mk4}", label_visibility="collapsed",
                                )
                                st.session_state.setdefault("pending_rate", {})
                                _cur4_v = float(_cur4.get("daily", 0.0) or 0.0)
                                if abs(_val4 - _cur4_v) > 1e-9:
                                    st.session_state["pending_rate"][_mk4] = (_val4, _it4.get("unit", ""))
                                else:
                                    st.session_state["pending_rate"].pop(_mk4, None)
                        st.caption("※ 0으로 두면 추정값을 그대로 사용합니다.")
                        _render_apply_rates("lr")

            _review_section()


        with _main_tabs[1]:
            st.caption("일위대가·품셈 어디에도 값이 없어 1일 작업량을 직접 입력해야 하는 항목입니다.")
            # 하위 탭 (4개 그룹)
            sub_tab_labels = []
            sub_tab_keys = []
            for mk, gd in group_data.items():
                count = sum(len(cat["items"]) for cat in gd["categories"])
                if count > 0:
                    sub_tab_labels.append(f"{gd['name']} ({count})")
                    sub_tab_keys.append(mk)
        
            if not sub_tab_labels:
                st.success("🎉 모든 항목이 매칭되었습니다!")
            else:
                sub_tabs = st.tabs(sub_tab_labels)
            
                # 🚀 @st.fragment로 페이지 변경 시 전체 재실행 방지
                @st.fragment
                def render_manual_input_page(mk, gd):
                    # 모든 항목을 평탄화
                    all_items = []
                    for cat in gd["categories"]:
                        for item in cat["items"]:
                            all_items.append({
                                **item,
                                "category": cat["category"]
                            })
                
                    if not all_items:
                        st.info("매칭 안 된 항목이 없습니다.")
                        return
                
                    # 🎯 주공정 항목을 위로 정렬 (선택 안 했으면 전체를 주공정으로 간주 → 정렬 안 함)
                    _sel_major_manual = set(st.session_state.get("selected_major", []))
                    if _sel_major_manual:
                        all_items.sort(key=lambda it: 0 if it.get("category") in _sel_major_manual else 1)
                
                    # 페이지네이션
                    items_per_page = 20
                    total_pages = (len(all_items) + items_per_page - 1) // items_per_page
                
                    page_key = f"page_{mk}"
                    if page_key not in st.session_state:
                        st.session_state[page_key] = 1
                
                    # 페이지 선택
                    col_p1, col_p2, col_p3 = st.columns([1, 2, 1])
                    with col_p1:
                        if st.button("◀ 이전", key=f"prev_{mk}", disabled=(st.session_state[page_key] <= 1)):
                            st.session_state[page_key] -= 1
                            st.rerun(scope="fragment")
                    with col_p2:
                        st.markdown(f"<div style='text-align: center; padding: 8px;'>페이지 {st.session_state[page_key]} / {total_pages}</div>", unsafe_allow_html=True)
                    with col_p3:
                        if st.button("다음 ▶", key=f"next_{mk}", disabled=(st.session_state[page_key] >= total_pages)):
                            st.session_state[page_key] += 1
                            st.rerun(scope="fragment")
                
                    # 현재 페이지 항목
                    start_idx = (st.session_state[page_key] - 1) * items_per_page
                    end_idx = min(start_idx + items_per_page, len(all_items))
                    page_items = all_items[start_idx:end_idx]
                
                    st.markdown("---")
                    st.markdown(f"### 📋 {start_idx + 1} ~ {end_idx} 번째 항목")
                
                    # 직전 저장 결과 안내 (전체 재실행 후에도 보이도록 세션에 담아둠)
                    _msg = st.session_state.pop("manual_saved_msg", None)
                    if _msg:
                        st.success(_msg)

                    # 일괄 입력 폼
                    with st.form(key=f"form_{mk}_{st.session_state[page_key]}"):
                        # 헤더
                        col_h0, col_h1, col_h2, col_h3, col_h4, col_h5 = st.columns([1.3, 2, 2, 1, 1.5, 1])
                        with col_h0:
                            st.markdown("**대공종**")
                        with col_h1:
                            st.markdown("**항목명**")
                        with col_h2:
                            st.markdown("**규격**")
                        with col_h3:
                            st.markdown("**수량**")
                        with col_h4:
                            st.markdown("**1일 작업량**")
                        with col_h5:
                            st.markdown("**단위**")
                    
                        st.markdown("---")
                    
                        # 입력 폼
                        form_inputs = {}
                        for i, item in enumerate(page_items):
                            manual_key = item["manual_key"]
                        
                            # 기존 값
                            existing_val = 0.0
                            existing_unit = item.get('unit', '') + "/일"
                            if manual_key in st.session_state["manual_rates"]:
                                existing_val = st.session_state["manual_rates"][manual_key].get("daily", 0)
                                existing_unit = st.session_state["manual_rates"][manual_key].get("unit", existing_unit)
                        
                            col0, col1, col2, col3, col4, col5 = st.columns([1.3, 2, 2, 1, 1.5, 1])
                            with col0:
                                _cat = item.get("category", "")
                                if not _sel_major_manual or _cat in _sel_major_manual:
                                    st.markdown(f"🎯 {_cat}")
                                else:
                                    st.markdown(f"⚪ {_cat}")
                            with col1:
                                st.text(item["name"])
                            with col2:
                                st.text(item.get("spec", ""))
                            with col3:
                                st.text(f"{item.get('qty', 0):,.1f}")
                            with col4:
                                daily = st.number_input(
                                    "daily",
                                    min_value=0.0,
                                    value=float(existing_val),
                                    step=0.1,
                                    key=f"in_{mk}_{start_idx + i}",
                                    label_visibility="collapsed"
                                )
                            with col5:
                                unit_in = st.text_input(
                                    "unit",
                                    value=existing_unit,
                                    key=f"un_{mk}_{start_idx + i}",
                                    label_visibility="collapsed"
                                )
                        
                            form_inputs[manual_key] = {"daily": daily, "unit": unit_in}
                    
                        st.markdown("---")
                    
                        # 일괄 저장 버튼
                        submitted = st.form_submit_button("💾 이 페이지 일괄 저장", width="stretch", type="primary")
                    
                        if submitted:
                            saved = 0
                            for mk_key, vals in form_inputs.items():
                                if vals["daily"] > 0:
                                    st.session_state["manual_rates"][mk_key] = vals
                                    saved += 1
                                else:
                                    # 0으로 지우면 기존 수동입력도 해제
                                    st.session_state["manual_rates"].pop(mk_key, None)
                            st.session_state["manual_saved_msg"] = f"✅ {saved}개 항목 저장 완료! 결과표에 반영되었습니다."
                            # scope="fragment"로 재실행하면 이 폼만 다시 그려져
                            # 본문 결과표(작업일수·공기)가 갱신되지 않아 "반응 없음"처럼 보인다.
                            # 저장은 전체 계산에 영향을 주므로 앱 전체를 재실행한다.
                            st.rerun()
            
                # 각 sub_tab에서 fragment 함수 호출
                for sub_tab, mk in zip(sub_tabs, sub_tab_keys):
                    with sub_tab:
                        render_manual_input_page(mk, group_data[mk])