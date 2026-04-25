import streamlit as st
import pandas as pd
import math
import holidays
from datetime import date, timedelta
import plotly.express as px

st.set_page_config(page_title="상하수도 공기산정", layout="wide")

# ── 비밀번호 로그인 ───────────────────────────────────────────
def check_password():
    if st.session_state.get("authenticated"):
        return True
    st.title("상하수도 관로공사 공기산정 시스템")
    st.markdown("---")
    st.subheader("로그인")
    pw = st.text_input("비밀번호를 입력하세요", type="password")
    if st.button("로그인"):
        if pw == "1234":
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("비밀번호가 틀렸습니다.")
    return False

if not check_password():
    st.stop()

# ── 표준품셈 딕셔너리 ─────────────────────────────────────────
LABOR_RATES = {
    "준비공": {
        "규준틀 설치": {"unit": "개소", "보통인부": 0.5},
    },
    "굴착공": {
        "터파기(기계)": {"unit": "m³", "특수작업원": 0.02, "보통인부": 0.03},
        "버력운반(기계)": {"unit": "m³", "특수작업원": 0.01, "보통인부": 0.02},
    },
    "관부설공": {
        "관 부설·접합": {
            "200mm": {"unit": "m", "배관공": 0.45, "보통인부": 0.35},
            "300mm": {"unit": "m", "배관공": 0.65, "보통인부": 0.50},
        },
        "수압시험": {
            "200mm": {"unit": "m", "배관공": 0.02, "보통인부": 0.02},
            "300mm": {"unit": "m", "배관공": 0.03, "보통인부": 0.02},
        },
    },
    "되메우기공": {
        "모래기초 포설": {"unit": "m³", "보통인부": 0.35},
        "되메우기(기계다짐)": {"unit": "m³", "특수작업원": 0.02, "보통인부": 0.10},
    },
    "포장복구공": {
        "보조기층 포설": {"unit": "m²", "특수작업원": 0.008, "보통인부": 0.020},
        "아스콘포장":   {"unit": "m²", "특수작업원": 0.010, "보통인부": 0.025},
    },
}

# ── 키워드 매핑 ───────────────────────────────────────────────
KEYWORD_MAP = {
    "준비공":     ["규준", "가시설", "토류", "교통"],
    "굴착공":     ["터파기", "굴착"],
    "관부설공":   ["관 부설", "관부설", "배관", "접합"],
    "되메우기공": ["모래", "기초", "되메우기", "다짐"],
    "포장복구공": ["아스콘", "포장", "복구"],
    "맨홀공":     ["맨홀"],
}

# ── 가이드라인 비작업일수 데이터 ──────────────────────────────
HOLIDAYS_2025 = {1:3,2:2,3:1,4:0,5:2,6:1,7:0,8:1,9:3,10:2,11:0,12:1}

RAIN_DAYS_SEOUL = {1:0.0,2:0.0,3:0.0,4:0.0,5:0.2,6:0.6,7:7.1,8:7.7,9:0.0,10:0.0,11:0.0,12:0.0}
COLD_DAYS_SEOUL = {1:15.0,2:12.0,3:3.0,4:0.0,5:0.0,6:0.0,7:0.0,8:0.0,9:0.0,10:0.5,11:5.0,12:12.0}
HEAT_DAYS_SEOUL = {1:0.0,2:0.0,3:0.0,4:0.0,5:0.0,6:0.5,7:2.0,8:3.0,9:0.0,10:0.0,11:0.0,12:0.0}
WIND_DAYS_SEOUL = {1:1.0,2:0.5,3:0.5,4:0.5,5:0.0,6:0.0,7:0.5,8:0.5,9:0.5,10:0.5,11:0.5,12:1.0}

CITY_CORRECTION = {"서울":1.0,"인천":0.95,"수원":1.0,"부산":0.8,"대구":1.5,"광주":1.1,"대전":0.9}

# ── 공통 함수 ─────────────────────────────────────────────────
def calc_manday(rates, quantity):
    total = 0.0
    for k, v in rates.items():
        if k != "unit":
            total += v * quantity
    return round(total, 2)

def to_days(manday, workers):
    if workers <= 0:
        return 0
    return math.ceil(manday / workers)

def get_work_end_date(start, work_days):
    kr_holidays = holidays.KR()
    current = start
    worked = 0
    RAIN = {1:2,2:2,3:3,4:4,5:5,6:7,7:11,8:10,9:6,10:3,11:3,12:2}
    while worked < work_days:
        if current.weekday() == 6:
            current += timedelta(days=1); continue
        if current in kr_holidays:
            current += timedelta(days=1); continue
        if current.day % 30 < RAIN[current.month]:
            current += timedelta(days=1); continue
        worked += 1
        current += timedelta(days=1)
    return current - timedelta(days=1)

# ═══════════════════════════════════════════════════════════════
# 사이드바
# ═══════════════════════════════════════════════════════════════
st.sidebar.header("기본 설정")
pipe_dia   = st.sidebar.selectbox("관경", ["200mm", "300mm"])
start_date = st.sidebar.date_input("착공 예정일", value=date.today())
st.sidebar.markdown("---")
st.sidebar.header("공종별 투입 인원 (명/일)")

if "workers" not in st.session_state:
    st.session_state.workers = {
        "준비공":4,"굴착공":6,"관부설공":4,"되메우기공":4,"포장복구공":4
    }

for 공종 in ["준비공","굴착공","관부설공","되메우기공","포장복구공"]:
    col_a, col_b, col_c = st.sidebar.columns([1,2,1])
    with col_a:
        if st.button("－", key=f"minus_{공종}"):
            st.session_state.workers[공종] = max(1, st.session_state.workers[공종]-1)
    with col_b:
        st.markdown(f"<div style='text-align:center;padding-top:6px'><b>{공종}</b><br>{st.session_state.workers[공종]}명</div>", unsafe_allow_html=True)
    with col_c:
        if st.button("＋", key=f"plus_{공종}"):
            st.session_state.workers[공종] = min(50, st.session_state.workers[공종]+1)

w = st.session_state.workers

# ═══════════════════════════════════════════════════════════════
# 탭 구성
# ═══════════════════════════════════════════════════════════════
st.title("상하수도 관로공사 공기산정 시스템")
st.markdown("---")

tab1, tab2, tab3 = st.tabs(["📋 공기산정", "📂 엑셀 내역서 인식", "🌧 비작업일수 계산기"])

# ═══════════════════════════════════════════════════════════════
# TAB 1: 공기산정
# ═══════════════════════════════════════════════════════════════
with tab1:
    st.subheader("공종별 물량 입력")
    st.caption("엑셀 내역서 보면서 아래 5개 숫자만 입력하세요.")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**준비공**")
        q_준비 = st.number_input("규준틀 설치 (개소)", min_value=0.0, value=st.session_state.get("q_준비",5.0), step=1.0, key="q_준비")
        st.markdown("**굴착공**")
        q_터파기 = st.number_input("터파기 물량 (m³)", min_value=0.0, value=st.session_state.get("q_터파기",350.0), step=10.0, key="q_터파기")
        st.markdown("**관부설공**")
        q_관부설 = st.number_input("관 부설 연장 (m)", min_value=0.0, value=st.session_state.get("q_관부설",120.0), step=10.0, key="q_관부설")
    with col2:
        st.markdown("**되메우기공**")
        q_되메우기 = st.number_input("되메우기 물량 (m³)", min_value=0.0, value=st.session_state.get("q_되메우기",180.0), step=10.0, key="q_되메우기")
        st.markdown("**포장복구공**")
        q_포장 = st.number_input("포장 면적 (m²)", min_value=0.0, value=st.session_state.get("q_포장",60.0), step=5.0, key="q_포장")

    st.markdown("---")

    # Man-day 계산
    md_준비     = calc_manday(LABOR_RATES["준비공"]["규준틀 설치"], q_준비)
    md_굴착     = calc_manday(LABOR_RATES["굴착공"]["터파기(기계)"], q_터파기) + calc_manday(LABOR_RATES["굴착공"]["버력운반(기계)"], q_터파기)
    md_관부설   = calc_manday(LABOR_RATES["관부설공"]["관 부설·접합"][pipe_dia], q_관부설) + calc_manday(LABOR_RATES["관부설공"]["수압시험"][pipe_dia], q_관부설)
    md_되메우기 = calc_manday(LABOR_RATES["되메우기공"]["모래기초 포설"], q_되메우기) + calc_manday(LABOR_RATES["되메우기공"]["되메우기(기계다짐)"], q_되메우기)
    md_포장     = calc_manday(LABOR_RATES["포장복구공"]["보조기층 포설"], q_포장) + calc_manday(LABOR_RATES["포장복구공"]["아스콘포장"], q_포장)

    d_준비     = to_days(md_준비,     w["준비공"])
    d_굴착     = to_days(md_굴착,     w["굴착공"])
    d_관부설   = to_days(md_관부설,   w["관부설공"])
    d_되메우기 = to_days(md_되메우기, w["되메우기공"])
    d_포장     = to_days(md_포장,     w["포장복구공"])
    d_total    = d_준비 + d_굴착 + d_관부설 + d_되메우기 + d_포장

    # 결과 표 (크리티컬패스 강조)
    st.subheader("공기산정 결과")
    st.caption("🔴 전체 공종이 크리티컬 패스입니다 (순차 FS 관계)")

    result_data = {
        "대공종":         ["준비공","굴착공","관부설공","되메우기공","포장복구공"],
        "투입인원 (명)":  [w["준비공"],w["굴착공"],w["관부설공"],w["되메우기공"],w["포장복구공"]],
        "Man-day (인일)": [md_준비,md_굴착,md_관부설,md_되메우기,md_포장],
        "작업일수 (일)":  [d_준비,d_굴착,d_관부설,d_되메우기,d_포장],
        "크리티컬패스":   ["🔴","🔴","🔴","🔴","🔴"],
    }
    result_df = pd.DataFrame(result_data)

    # 크리티컬패스 행 빨간색 강조
    def highlight_critical(row):
        return ['background-color: #3d0000; color: #ff6b6b'] * len(row)

    st.dataframe(
        result_df.style.apply(highlight_critical, axis=1),
        hide_index=True,
        use_container_width=True
    )

    st.markdown("---")
    st.subheader("공기 요약")
    col_a, col_b, col_c, col_d = st.columns(4)
    col_a.metric("순 작업일수", f"{d_total} 일")
    col_b.metric("투입 인원 합계", f"{sum(w.values())} 명")
    col_c.metric("총 Man-day", f"{round(md_준비+md_굴착+md_관부설+md_되메우기+md_포장,1)} 인일")
    col_d.metric("관경", pipe_dia)

    # 시나리오 비교
    st.markdown("---")
    st.subheader("조수 시나리오 비교")
    scenarios = []
    for label, factor in [("절반 인원",0.5),("현재 인원",1.0),("1.5배 인원",1.5),("2배 인원",2.0)]:
        sw = {k: max(1, round(v*factor)) for k,v in w.items()}
        sd = (to_days(md_준비,sw["준비공"]) + to_days(md_굴착,sw["굴착공"]) +
              to_days(md_관부설,sw["관부설공"]) + to_days(md_되메우기,sw["되메우기공"]) +
              to_days(md_포장,sw["포장복구공"]))
        end = get_work_end_date(start_date, sd)
        scenarios.append({
            "시나리오": label,
            "준비공(명)": sw["준비공"], "굴착공(명)": sw["굴착공"],
            "관부설공(명)": sw["관부설공"], "되메우기(명)": sw["되메우기공"],
            "포장복구(명)": sw["포장복구공"],
            "순작업일수(일)": sd,
            "준공예정일": end.strftime("%Y-%m-%d"),
        })
    st.dataframe(pd.DataFrame(scenarios), hide_index=True, use_container_width=True)

    # 간트차트
    st.markdown("---")
    st.subheader("간트차트")

    s1 = start_date
    e1 = get_work_end_date(s1, d_준비)
    s2 = e1 + timedelta(days=1); e2 = get_work_end_date(s2, d_굴착)
    s3 = e2 + timedelta(days=1); e3 = get_work_end_date(s3, d_관부설)
    s4 = e3 + timedelta(days=1); e4 = get_work_end_date(s4, d_되메우기)
    s5 = e4 + timedelta(days=1); e5 = get_work_end_date(s5, d_포장)

    gantt_data = pd.DataFrame([
        dict(Task="준비공",     Start=str(s1), Finish=str(e1), 인원=f"{w['준비공']}명",     작업일=f"{d_준비}일"),
        dict(Task="굴착공",     Start=str(s2), Finish=str(e2), 인원=f"{w['굴착공']}명",     작업일=f"{d_굴착}일"),
        dict(Task="관부설공",   Start=str(s3), Finish=str(e3), 인원=f"{w['관부설공']}명",   작업일=f"{d_관부설}일"),
        dict(Task="되메우기공", Start=str(s4), Finish=str(e4), 인원=f"{w['되메우기공']}명", 작업일=f"{d_되메우기}일"),
        dict(Task="포장복구공", Start=str(s5), Finish=str(e5), 인원=f"{w['포장복구공']}명", 작업일=f"{d_포장}일"),
    ])

    colors = {"준비공":"#5DCAA5","굴착공":"#378ADD","관부설공":"#D85A30","되메우기공":"#EF9F27","포장복구공":"#7F77DD"}
    fig = px.timeline(gantt_data, x_start="Start", x_end="Finish", y="Task", color="Task",
                      color_discrete_map=colors, hover_data={"인원":True,"작업일":True,"Task":False})
    fig.update_yaxes(autorange="reversed")
    fig.update_layout(height=350, showlegend=False, xaxis_title="날짜", yaxis_title="",
                      margin=dict(l=10,r=10,t=30,b=10))
    fig.update_traces(marker_line_color="red", marker_line_width=2)
    st.plotly_chart(fig, use_container_width=True)
    st.caption("빨간 테두리 = 크리티컬 패스 | 모든 공종이 순차 연결되어 전체가 크리티컬 패스입니다.")

    schedule_df = pd.DataFrame({
        "공종":     ["준비공","굴착공","관부설공","되메우기공","포장복구공"],
        "착수일":   [str(s1),str(s2),str(s3),str(s4),str(s5)],
        "완료일":   [str(e1),str(e2),str(e3),str(e4),str(e5)],
        "투입인원": [f"{w['준비공']}명",f"{w['굴착공']}명",f"{w['관부설공']}명",f"{w['되메우기공']}명",f"{w['포장복구공']}명"],
        "작업일수": [f"{d_준비}일",f"{d_굴착}일",f"{d_관부설}일",f"{d_되메우기}일",f"{d_포장}일"],
    })
    st.dataframe(schedule_df, hide_index=True, use_container_width=True)


# ═══════════════════════════════════════════════════════════════
# TAB 2: 엑셀 내역서 인식
# ═══════════════════════════════════════════════════════════════
with tab2:
    st.subheader("엑셀 내역서 자동 인식")
    st.caption("엑셀 파일을 업로드하면 공종을 자동으로 탐지해서 물량을 채워드립니다.")

    uploaded = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx","xls"])

    if uploaded:
        try:
            xl = pd.read_excel(uploaded, header=None)

            # 공종명 컬럼 자동 탐지
            COL_CANDIDATES = ["공종명","품명","공종","작업명","명칭"]
            UNIT_CANDIDATES = ["단위","규격단위"]
            QTY_CANDIDATES  = ["물량","수량","규격수량"]

            header_row = None
            name_col = unit_col = qty_col = None

            for i, row in xl.iterrows():
                for j, cell in enumerate(row):
                    if str(cell).strip() in COL_CANDIDATES:
                        header_row = i
                        name_col = j
                    if str(cell).strip() in UNIT_CANDIDATES and header_row == i:
                        unit_col = j
                    if str(cell).strip() in QTY_CANDIDATES and header_row == i:
                        qty_col = j
                if header_row is not None:
                    break

            if header_row is None:
                st.error("공종명 컬럼을 찾지 못했습니다. 컬럼명이 '공종명', '품명', '공종', '작업명' 중 하나인지 확인해주세요.")
            else:
                data = xl.iloc[header_row+1:].reset_index(drop=True)
                data.columns = range(len(data.columns))

                matched = []
                unmatched = []

                for _, row in data.iterrows():
                    cell_val = str(row[name_col]).strip() if name_col is not None else ""
                    if not cell_val or cell_val == "nan":
                        continue

                    unit_val = str(row[unit_col]).strip() if unit_col is not None and unit_col < len(row) else "-"
                    qty_val  = row[qty_col] if qty_col is not None and qty_col < len(row) else 0
                    try:
                        qty_val = float(qty_val)
                    except:
                        qty_val = 0.0

                    # 키워드 매핑
                    mapped = None
                    for 대공종, keywords in KEYWORD_MAP.items():
                        if any(kw in cell_val for kw in keywords):
                            mapped = 대공종
                            break

                    if mapped:
                        matched.append({"공종명": cell_val, "단위": unit_val, "물량": qty_val, "매핑공종": mapped})
                    else:
                        unmatched.append({"공종명": cell_val, "단위": unit_val, "물량": qty_val})

                # 인식 결과 표시
                st.markdown("---")
                st.markdown(f"**인식 결과: ✅ 매핑 {len(matched)}건 | ⚠️ 미인식 {len(unmatched)}건**")

                if matched:
                    st.markdown("#### ✅ 자동 매핑된 항목")
                    matched_df = pd.DataFrame(matched)
                    st.dataframe(
                        matched_df.style.applymap(lambda _: "background-color: #1a3a2a; color: #4CAF50"),
                        hide_index=True, use_container_width=True
                    )

                if unmatched:
                    st.markdown("#### ⚠️ 미인식 항목 — 공종 수동 선택")
                    공종목록 = ["(선택안함)","준비공","굴착공","관부설공","되메우기공","포장복구공","맨홀공","기타"]
                    for idx, item in enumerate(unmatched):
                        col_a, col_b, col_c, col_d = st.columns([3,1,1,2])
                        col_a.markdown(f"<span style='color:#FFA500'>{item['공종명']}</span>", unsafe_allow_html=True)
                        col_b.write(item['단위'])
                        col_c.write(item['물량'])
                        sel = col_d.selectbox("공종 선택", 공종목록, key=f"unmatched_{idx}")
                        if sel != "(선택안함)":
                            matched.append({"공종명": item["공종명"], "단위": item["단위"], "물량": item["물량"], "매핑공종": sel})

                # 물량 적용 버튼
                st.markdown("---")
                if st.button("✅ 인식된 물량을 공기산정에 적용"):
                    # 대공종별 물량 합산
                    qty_map = {"준비공":0,"굴착공":0,"관부설공":0,"되메우기공":0,"포장복구공":0}
                    for item in matched:
                        if item["매핑공종"] in qty_map:
                            qty_map[item["매핑공종"]] += item["물량"]

                    st.session_state["q_준비"]     = qty_map["준비공"]
                    st.session_state["q_터파기"]   = qty_map["굴착공"]
                    st.session_state["q_관부설"]   = qty_map["관부설공"]
                    st.session_state["q_되메우기"] = qty_map["되메우기공"]
                    st.session_state["q_포장"]     = qty_map["포장복구공"]
                    st.success("공기산정 탭으로 이동하면 물량이 반영되어 있습니다!")

        except Exception as e:
            st.error(f"파일 읽기 오류: {e}")
    else:
        st.info("엑셀 파일을 업로드해주세요. 컬럼명에 '공종명', '품명', '공종', '작업명' 중 하나가 있어야 합니다.")


# ═══════════════════════════════════════════════════════════════
# TAB 3: 비작업일수 계산기 (가이드라인 기준)
# ═══════════════════════════════════════════════════════════════
with tab3:
    st.subheader("비작업일수 계산기")
    st.caption("2024년 적정공사기간 확보 가이드라인 기준으로 비작업일수를 산출합니다.")

    col1, col2 = st.columns(2)
    with col1:
        start_year  = st.selectbox("공사 시작 연도", [2024,2025,2026], index=1)
        start_month = st.selectbox("공사 시작 월", list(range(1,13)), index=2, format_func=lambda x: f"{x}월")
        duration_months = st.number_input("공사 기간 (개월)", min_value=1, max_value=60, value=6)
        city = st.selectbox("공사 지역", list(CITY_CORRECTION.keys()))

    with col2:
        st.markdown("**기상 조건 선택**")
        use_rain = st.checkbox("강우 (일강수량 5mm 이상)", value=True)
        use_cold = st.checkbox("동절기 (최저기온 0℃ 이하)", value=True)
        use_heat = st.checkbox("혹서기 (최고기온 35℃ 이상)", value=False)
        use_wind = st.checkbox("강풍 (최대순간풍속 15m/s 이상)", value=False)
        prep_days    = st.number_input("준비기간 (일)", value=60, min_value=0)
        cleanup_days = st.number_input("정리기간 (일)", value=30, min_value=0)

    st.markdown("---")

    # 월별 비작업일수 계산
    corr = CITY_CORRECTION.get(city, 1.0)
    rows = []
    total_applied = 0

    for i in range(int(duration_months)):
        m = ((start_month - 1 + i) % 12) + 1

        # A: 기상 비작업일수
        A = 0.0
        if use_rain: A += RAIN_DAYS_SEOUL[m] * corr
        if use_cold: A += COLD_DAYS_SEOUL[m] * corr
        if use_heat: A += HEAT_DAYS_SEOUL[m] * corr
        if use_wind: A += WIND_DAYS_SEOUL[m] * corr

        # B: 법정공휴일
        B = HOLIDAYS_2025.get(m, 0)

        # 달력일수 (근사: 30일)
        calendar = 30

        # C: 중복일수
        C = round(A * B / calendar, 0) if calendar > 0 else 0

        # 비작업일수
        non_work = A + B - C

        # 최소 8일 적용
        applied = max(8.0, non_work)
        total_applied += applied

        rows.append({
            "월":           f"{m}월",
            "기상비작업일(A)": round(A, 1),
            "법정공휴일(B)":   B,
            "중복일수(C)":     int(C),
            "비작업일수":      round(non_work, 1),
            "적용일수":        round(applied, 1),
        })

    nw_df = pd.DataFrame(rows)

    # 합계 행 추가
    total_row = pd.DataFrame([{
        "월": "합계",
        "기상비작업일(A)": round(nw_df["기상비작업일(A)"].sum(), 1),
        "법정공휴일(B)":   nw_df["법정공휴일(B)"].sum(),
        "중복일수(C)":     nw_df["중복일수(C)"].sum(),
        "비작업일수":      round(nw_df["비작업일수"].sum(), 1),
        "적용일수":        round(total_applied, 1),
    }])
    nw_df_display = pd.concat([nw_df, total_row], ignore_index=True)

    st.dataframe(nw_df_display, hide_index=True, use_container_width=True)

    # 총 공사기간 계산
    st.markdown("---")
    st.subheader("총 공사기간 산출")

    work_days_total = d_total  # tab1에서 계산된 순작업일수
    total_duration  = work_days_total + int(total_applied) + prep_days + cleanup_days

    col_a, col_b, col_c, col_d, col_e = st.columns(5)
    col_a.metric("순 작업일수",  f"{work_days_total} 일")
    col_b.metric("비작업일수",   f"{int(total_applied)} 일")
    col_c.metric("준비기간",     f"{prep_days} 일")
    col_d.metric("정리기간",     f"{cleanup_days} 일")
    col_e.metric("총 공사기간",  f"{total_duration} 일", delta=f"약 {round(total_duration/30,1)}개월")

    st.info(f"""
    **총 공사기간 산출 기준 (가이드라인)**
    - 순 작업일수: {work_days_total}일 (공기산정 탭 결과)
    - 비작업일수: {int(total_applied)}일 (가이드라인 기준, 월 최소 8일 적용)
    - 준비기간: {prep_days}일 (착공 전 준비)
    - 정리기간: {cleanup_days}일 (주요공종 완료 후)
    - **합계: {total_duration}일 (약 {round(total_duration/30,1)}개월)**
    """)