# VERSION 3.0 - 지구별 분리 + 전체 TAB
import streamlit as st
import pandas as pd
import openpyxl
import re
from datetime import datetime, timedelta

# 페이지 설정
st.set_page_config(page_title="상하수도 공기산정", layout="wide", initial_sidebar_state="expanded")

# ============================================================================
# 임포트
# ============================================================================
try:
    from guideline_data import PAVEMENT
    from weather_data import HEAT_DAYS, REGIONS, get_heat_days_by_region, get_total_non_work_days
except ImportError:
    st.warning("guideline_data.py 또는 weather_data.py 파일이 없습니다.")
    PAVEMENT = {}
    HEAT_DAYS = {}
    REGIONS = ["서울"]
    def get_heat_days_by_region(region, month=None):
        return 0.0
    def get_total_non_work_days(region, start, end):
        return 0.0

# ============================================================================
# 파싱 함수
# ============================================================================

def parse_excel_tree(ws):
    """엑셀을 트리 구조로 파싱"""
    current_path = {
        'roman': None,
        'level1': None,
        'level2': None,
        'level3': None,
        'sub1': None,
    }
    
    tree = {}
    
    for row in ws.iter_rows(min_row=1, values_only=True):
        col_a = str(row[0]).strip() if row[0] else ""
        col_b = str(row[1]).strip() if row[1] else ""
        col_c = str(row[2]).strip() if row[2] else ""
        col_d = row[3] if row[3] else ""
        col_e = str(row[4]).strip() if row[4] else ""
        
        if not col_a and not col_b:
            continue
        
        # 로마숫자
        if col_a.startswith('Ⅰ') or col_a.startswith('Ⅱ') or col_a.startswith('Ⅲ') or col_a.startswith('Ⅳ') or col_a.startswith('Ⅴ'):
            current_path = {k: None for k in current_path}
            current_path['roman'] = col_a
            tree[col_a] = {'name': col_b, 'children': {}}
            continue
        
        # 1., 2.
        if re.match(r'^\d+\.$', col_a):
            current_path['level1'] = col_a.rstrip('.')
            current_path['level2'] = None
            current_path['level3'] = None
            current_path['sub1'] = None
            
            roman = current_path['roman']
            if roman and roman in tree:
                tree[roman]['children'][current_path['level1']] = {'name': col_b, 'children': {}}
            continue
        
        # 1.1, 1.2
        if re.match(r'^\d+\.\d+$', col_a) and col_a.count('.') == 1:
            current_path['level2'] = col_a
            current_path['level3'] = None
            current_path['sub1'] = None
            
            roman = current_path['roman']
            lv1 = current_path['level1']
            if roman and lv1 and roman in tree and lv1 in tree[roman]['children']:
                tree[roman]['children'][lv1]['children'][col_a] = {'name': col_b, 'children': {}}
            continue
        
        # 1.1.1 (공종)
        if re.match(r'^\d+\.\d+\.\d+$', col_a):
            current_path['level3'] = col_a
            current_path['sub1'] = None
            
            roman = current_path['roman']
            lv1 = current_path['level1']
            lv2 = current_path['level2']
            if roman and lv1 and lv2:
                if roman in tree and lv1 in tree[roman]['children'] and lv2 in tree[roman]['children'][lv1]['children']:
                    tree[roman]['children'][lv1]['children'][lv2]['children'][col_a] = {
                        'name': col_b,
                        'items': [],
                        'sub_categories': {}
                    }
            continue
        
        # 1), 2)
        if re.match(r'^\d+\)$', col_a):
            current_path['sub1'] = col_a
            
            roman = current_path['roman']
            lv1 = current_path['level1']
            lv2 = current_path['level2']
            lv3 = current_path['level3']
            if roman and lv1 and lv2 and lv3:
                if (roman in tree and lv1 in tree[roman]['children'] and 
                    lv2 in tree[roman]['children'][lv1]['children'] and
                    lv3 in tree[roman]['children'][lv1]['children'][lv2]['children']):
                    tree[roman]['children'][lv1]['children'][lv2]['children'][lv3]['sub_categories'][col_a] = {
                        'name': col_b,
                        'items': []
                    }
            continue
        
        # (1), (2) - 무시
        if re.match(r'^\(\d+\)$', col_a):
            continue
        
        # 항목
        if not col_a and col_b:
            roman = current_path['roman']
            lv1 = current_path['level1']
            lv2 = current_path['level2']
            lv3 = current_path['level3']
            sub1 = current_path['sub1']
            
            if not (roman and lv1 and lv2 and lv3):
                continue
            
            qty = 0
            try:
                if col_d:
                    qty = float(str(col_d).replace(',', ''))
            except:
                pass
            
            item = {
                'name': col_b,
                'spec': col_c,
                'qty': qty,
                'unit': col_e
            }
            
            if sub1:
                if (roman in tree and lv1 in tree[roman]['children'] and 
                    lv2 in tree[roman]['children'][lv1]['children'] and
                    lv3 in tree[roman]['children'][lv1]['children'][lv2]['children'] and
                    sub1 in tree[roman]['children'][lv1]['children'][lv2]['children'][lv3]['sub_categories']):
                    tree[roman]['children'][lv1]['children'][lv2]['children'][lv3]['sub_categories'][sub1]['items'].append(item)
            else:
                if (roman in tree and lv1 in tree[roman]['children'] and 
                    lv2 in tree[roman]['children'][lv1]['children'] and
                    lv3 in tree[roman]['children'][lv1]['children'][lv2]['children']):
                    tree[roman]['children'][lv1]['children'][lv2]['children'][lv3]['items'].append(item)
    
    return tree


def flatten_categories(tree):
    """트리를 공종 리스트로 변환"""
    categories = []
    
    for roman_key, roman_data in tree.items():
        for lv1_key, lv1_data in roman_data.get('children', {}).items():
            for lv2_key, lv2_data in lv1_data.get('children', {}).items():
                for lv3_key, lv3_data in lv2_data.get('children', {}).items():
                    full_path = f"{roman_key} > {lv1_key}. {lv1_data['name']} > {lv2_key} {lv2_data['name']}"
                    
                    categories.append({
                        'roman': roman_key,
                        'roman_name': roman_data['name'],
                        'level': lv3_key,
                        'name': lv3_data['name'],
                        'full_path': full_path,
                        'items': lv3_data.get('items', []),
                        'sub_categories': lv3_data.get('sub_categories', {})
                    })
    
    return categories


def calc_work_days(item_name, item_spec, qty, unit):
    """작업일수 계산"""
    if not qty or qty <= 0:
        return 0, None
    
    name_no_space = item_name.replace(" ", "")
    
    # 1. 이름 + 스펙
    full_key = f"{item_name} {item_spec}".strip()
    if full_key in PAVEMENT:
        daily = PAVEMENT[full_key].get('daily', 0)
        # 비정상적인 값 무시 (하루 작업량이 1 미만이면 이상함)
        if daily >= 1:
            days = max(1, round(qty / daily))
            return days, f"{daily}{unit}/일"
    
    # 2. 이름만
    if item_name in PAVEMENT:
        daily = PAVEMENT[item_name].get('daily', 0)
        if daily >= 1:
            days = max(1, round(qty / daily))
            return days, f"{daily}{unit}/일"
    
    # 3. 공백 제거
    if name_no_space in PAVEMENT:
        daily = PAVEMENT[name_no_space].get('daily', 0)
        if daily >= 1:
            days = max(1, round(qty / daily))
            return days, f"{daily}{unit}/일"
    
    # 4. 부분 매칭
    for key in PAVEMENT:
        if key in item_name or item_name in key:
            daily = PAVEMENT[key].get('daily', 0)
            if daily >= 1:
                days = max(1, round(qty / daily))
                return days, f"{daily}{unit}/일"
    
    return 0, None

# ============================================================================
# 메인 앱
# ============================================================================

st.title("🏗️ 상하수도 공사 공기산정")

# 사이드바
with st.sidebar:
    st.header("📤 파일 업로드")
    uploaded_file = st.file_uploader("엑셀 파일 업로드", type=['xlsx'])

if not uploaded_file:
    st.info("👈 왼쪽에서 엑셀 파일을 업로드하세요!")
    st.stop()

# 엑셀 읽기
try:
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    ws = wb['설계내역서']
except Exception as e:
    st.error(f"엑셀 파일 읽기 실패: {e}")
    st.stop()

# 파싱
with st.spinner("엑셀 파일 분석 중..."):
    tree = parse_excel_tree(ws)
    categories = flatten_categories(tree)

st.success(f"✅ {len(categories)}개 공종 파싱 완료!")

# ============================================================================
# TAB 구조
# ============================================================================

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📋 공기산정",
    "📂 엑셀 내역서 인식",
    "🔍 주요공종 CP 분석",
    "🌧 비작업일수 계산기",
    "📄 공기산정 보고서"
])

# ============================================================================
# TAB 2: 상세 분석
# ============================================================================
with tab2:
    st.header("📊 상세 분석")
    
    # 투입조수 설정 초기화
    if 'crew_settings' not in st.session_state:
        st.session_state['crew_settings'] = {}
    
    # 지구별로 그룹핑
    for roman_key in sorted(tree.keys()):
        roman_data = tree[roman_key]
        
        st.markdown(f"## 📍 {roman_key} {roman_data['name']}")
        
        # 해당 지구의 공종만 필터링
        district_categories = []
        for cat in categories:
            if cat['full_path'].startswith(roman_key):
                district_categories.append(cat)
        
        # 이 지구의 투입조수 설정
        st.markdown("### 🔧 투입조수 설정")
        
        # 변경 감지를 위한 임시 변수
        changed = False
        
        cols = st.columns(min(len(district_categories), 4))
        
        for idx, cat in enumerate(district_categories):
            work_key = f"{roman_key}_{cat['level']} {cat['name']}"
            default_crew = st.session_state['crew_settings'].get(work_key, 3)
            
            with cols[idx % len(cols)]:
                crew_val = st.number_input(
                    f"{cat['level']} {cat['name']}",
                    min_value=1,
                    max_value=30,
                    value=default_crew,
                    step=1,
                    key=f"crew_{roman_key}_{idx}",
                    help="투입조수 (1~30조)"
                )
                # 값이 변경되었는지 확인
                if st.session_state['crew_settings'].get(work_key, 3) != crew_val:
                    changed = True
                st.session_state['crew_settings'][work_key] = crew_val
        
        st.markdown("---")
        
        # 공종 리스트 표시
        for cat in district_categories:
            # 투입조수 가져오기
            work_key = f"{roman_key}_{cat['level']} {cat['name']}"
            crew = st.session_state['crew_settings'].get(work_key, 3)
            
            # 총 작업일수 계산
            total_days = 0
            item_count = 0
            
            # 직접 항목
            for item in cat['items']:
                days, _ = calc_work_days(item['name'], item.get('spec', ''), item.get('qty', 0), item.get('unit', ''))
                total_days += max(1, round(days / crew))  # 투입조수로 나누기
                if item.get('qty', 0) > 0:
                    item_count += 1
            
            # sub_category 항목
            sub_days_map = {}
            for sub_key, sub_data in cat['sub_categories'].items():
                sub_days = 0
                for item in sub_data.get('items', []):
                    days, _ = calc_work_days(item['name'], item.get('spec', ''), item.get('qty', 0), item.get('unit', ''))
                    sub_days += max(1, round(days / crew))  # 투입조수로 나누기
                    if item.get('qty', 0) > 0:
                        item_count += 1
                sub_days_map[sub_key] = sub_days
                total_days += sub_days
            
            # Expander
            with st.expander(f"▶ {cat['level']} {cat['name']} - {total_days}일 ({crew}조, {item_count}개)", expanded=False):
                
                # Sub-categories
                for sub_key in sorted(cat['sub_categories'].keys()):
                    sub_data = cat['sub_categories'][sub_key]
                    sub_days = sub_days_map.get(sub_key, 0)
                    
                    st.markdown(f"### {sub_key} {sub_data['name']} ({sub_days}일)")
                    
                    # 항목 테이블
                    if sub_data.get('items'):
                        rows = []
                        for item in sub_data['items']:
                            days, rate = calc_work_days(item['name'], item.get('spec', ''), item.get('qty', 0), item.get('unit', ''))
                            adjusted_days = max(1, round(days / crew))
                            rows.append({
                                '세부공종': item['name'],
                                '규격': item.get('spec', ''),
                                '수량': f"{item.get('qty', 0):,.1f}",
                                '단위': item.get('unit', ''),
                                '1일작업량': rate or '-',
                                '작업일수(1조)': days,
                                f'작업일수({crew}조)': adjusted_days,
                                '출처': '가이드라인' if rate else '-'
                            })
                        
                        if rows:
                            df = pd.DataFrame(rows)
                            st.dataframe(df, use_container_width=True, hide_index=True)
                
                # 직접 항목
                if cat['items']:
                    st.markdown("### 직접 항목")
                    rows = []
                    for item in cat['items']:
                        days, rate = calc_work_days(item['name'], item.get('spec', ''), item.get('qty', 0), item.get('unit', ''))
                        adjusted_days = max(1, round(days / crew))
                        rows.append({
                            '세부공종': item['name'],
                            '규격': item.get('spec', ''),
                            '수량': f"{item.get('qty', 0):,.1f}",
                            '단위': item.get('unit', ''),
                            '1일작업량': rate or '-',
                            '작업일수(1조)': days,
                            f'작업일수({crew}조)': adjusted_days,
                            '출처': '가이드라인' if rate else '-'
                        })
                    
                    if rows:
                        df = pd.DataFrame(rows)
                        st.dataframe(df, use_container_width=True, hide_index=True)
        
        st.markdown("---")  # 지구 구분선
    
    # 세션 저장
    st.session_state["tree"] = tree
    st.session_state["categories"] = categories

# ============================================================================
# TAB 1: 공기산정 요약
# ============================================================================
with tab1:
    st.subheader("📋 공기산정 요약")
    
    st.markdown("### 📊 전체 공기 요약")
    
    total_work_days = 0
    
    # 투입조수 설정 확인
    if 'crew_settings' not in st.session_state:
        st.session_state['crew_settings'] = {}
    
    # 전체 작업일수 계산 (투입조수 반영)
    for cat in categories:
        work_key = f"{cat['roman']}_{cat['level']} {cat['name']}"
        crew = st.session_state['crew_settings'].get(work_key, 3)
        
        for item in cat['items']:
            days, _ = calc_work_days(item['name'], item.get('spec', ''), item.get('qty', 0), item.get('unit', ''))
            total_work_days += max(1, round(days / crew))
        
        for sub_data in cat['sub_categories'].values():
            for item in sub_data.get('items', []):
                days, _ = calc_work_days(item['name'], item.get('spec', ''), item.get('qty', 0), item.get('unit', ''))
                total_work_days += max(1, round(days / crew))
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("🔴 총 작업일수", f"{total_work_days}일")
    with col2:
        st.metric("📁 총 공종", f"{len(categories)}개")
    with col3:
        st.metric("🏗️ 지구", f"{len(tree)}개")
    
    st.info("📌 TAB 2에서 상세 내역을 확인하세요!")

# ============================================================================
# TAB 3: CP 분석
# ============================================================================
with tab3:
    st.subheader("🔍 주요공종 CP 분석")
    st.info("🚧 준비중입니다")

# ============================================
# TAB 4: 비작업일수 계산기
# ============================================
with tab4:
    st.header("☂️ 비작업일수 계산기")
    
    st.markdown("""
    공사 기간 중 기후 조건에 따른 비작업일수를 계산합니다.
    - **강우일**: 일 강수량 기준 작업 불가일
    - **한랭일**: 일 최저기온 -10°C 이하
    - **폭염일**: 일 최고기온 33°C 이상
    """)
    
    # 기본 설정
    col1, col2 = st.columns(2)
    
    with col1:
        calc_start_date = st.date_input(
            "공사 시작일",
            value=datetime(2026, 12, 25),
            help="공사가 시작되는 날짜를 선택하세요"
        )
    
    with col2:
        calc_region = st.selectbox(
            "지역 선택",
            options=list(REGION_MAPPING.keys()),
            index=0,
            help="공사 지역을 선택하세요"
        )
    
    work_days_input = st.number_input(
        "순공기(작업일수)",
        min_value=1,
        max_value=10000,
        value=1200,
        step=10,
        help="실제 작업이 필요한 일수를 입력하세요"
    )
    
    # ✅ 기후 조건 체크박스 추가
    st.subheader("🌦️ 기후 조건 선택")
    st.caption("제외할 기후 조건을 선택하세요")
    
    col_a, col_b, col_c = st.columns(3)
    
    with col_a:
        check_rain = st.checkbox(
            "💧 강우일 제외", 
            value=True,
            help="강수량 기준 작업 불가일을 포함합니다"
        )
    
    with col_b:
        check_cold = st.checkbox(
            "❄️ 한랭일 제외", 
            value=True,
            help="일 최저기온 -10°C 이하인 날을 포함합니다"
        )
    
    with col_c:
        check_hot = st.checkbox(
            "🌡️ 폭염일 제외", 
            value=True,
            help="일 최고기온 33°C 이상인 날을 포함합니다"
        )
    
    st.divider()
    
    # 계산 버튼
    if st.button("🔢 비작업일수 계산", type="primary", use_container_width=True):
        try:
            # 종료일 계산
            calc_end_date = calc_start_date + timedelta(days=work_days_input)
            
            # 비작업일수 계산 (체크박스 값 전달)
            non_work_days = get_total_non_work_days(
                calc_region, 
                calc_start_date, 
                calc_end_date,
                check_rain=check_rain,
                check_cold=check_cold,
                check_hot=check_hot
            )
            
            # 실제 총공기
            total_days = work_days_input + non_work_days
            actual_end_date = calc_start_date + timedelta(days=total_days)
            
            # 결과 표시
            st.success(f"✅ 계산 완료!")
            
            # 메트릭 표시
            metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)
            
            with metric_col1:
                st.metric(
                    label="순공기",
                    value=f"{work_days_input:,}일",
                    help="실제 작업일수"
                )
            
            with metric_col2:
                st.metric(
                    label="비작업일수",
                    value=f"{non_work_days:,}일",
                    delta=f"{(non_work_days/work_days_input*100):.1f}%",
                    help="기후 조건으로 인한 작업 불가일"
                )
            
            with metric_col3:
                st.metric(
                    label="총 공기",
                    value=f"{total_days:,}일",
                    help="순공기 + 비작업일수"
                )
            
            with metric_col4:
                st.metric(
                    label="예상 완공일",
                    value=actual_end_date.strftime('%y.%m.%d'),
                    help="공사 종료 예정일"
                )
            
            # 상세 정보
            st.info(f"""
            📅 **공사 기간**: {calc_start_date.strftime('%Y년 %m월 %d일')} ~ {actual_end_date.strftime('%Y년 %m월 %d일')}  
            📍 **지역**: {calc_region}  
            🌦️ **적용 조건**: {'강우일' if check_rain else ''} {'한랭일' if check_cold else ''} {'폭염일' if check_hot else ''}
            """)
            
            # 월별 상세 내역
            st.subheader("📊 월별 비작업일수 상세")
            
            monthly_data = get_monthly_breakdown(
                calc_region,
                calc_start_date,
                calc_end_date,
                check_rain=check_rain,
                check_cold=check_cold,
                check_hot=check_hot
            )
            
            if monthly_data:
                # 데이터프레임 생성
                import pandas as pd
                df_monthly = pd.DataFrame(monthly_data)
                df_monthly.columns = ["월", "강우일", "한랭일", "폭염일", "합계"]
                
                # 표 표시
                st.dataframe(
                    df_monthly,
                    use_container_width=True,
                    hide_index=True
                )
                
                # 차트 표시
                import plotly.graph_objects as go
                
                fig = go.Figure()
                
                if check_rain:
                    fig.add_trace(go.Bar(
                        name='강우일',
                        x=df_monthly['월'],
                        y=df_monthly['강우일'],
                        marker_color='#4A90E2'
                    ))
                
                if check_cold:
                    fig.add_trace(go.Bar(
                        name='한랭일',
                        x=df_monthly['월'],
                        y=df_monthly['한랭일'],
                        marker_color='#5BC0DE'
                    ))
                
                if check_hot:
                    fig.add_trace(go.Bar(
                        name='폭염일',
                        x=df_monthly['월'],
                        y=df_monthly['폭염일'],
                        marker_color='#F0AD4E'
                    ))
                
                fig.update_layout(
                    title='월별 비작업일수',
                    xaxis_title='월',
                    yaxis_title='일수',
                    barmode='stack',
                    height=400,
                    hovermode='x unified'
                )
                
                st.plotly_chart(fig, use_container_width=True)
            
        except Exception as e:
            st.error(f"❌ 계산 중 오류 발생")
            st.error(f"오류 내용: {str(e)}")
            st.info("날짜와 지역을 다시 확인해주세요.")
            
            # 디버그 정보
            with st.expander("🐛 디버그 정보"):
                st.code(f"""
시작일: {calc_start_date}
지역: {calc_region}
순공기: {work_days_input}
강우일 체크: {check_rain}
한랭일 체크: {check_cold}
폭염일 체크: {check_hot}
에러: {str(e)}
                """)

# ============================================================================
# TAB 5: 보고서
# ============================================================================
with tab5:
    st.subheader("📄 공기산정 보고서")
    st.info("🚧 준비중입니다")