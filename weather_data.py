"""
상하수도 공사 비작업일수 계산 모듈
- 강우일, 한랭일, 폭염일 데이터 기반 비작업일수 산정
"""

from datetime import datetime, timedelta

# 지역별 매핑
REGION_MAPPING = {
    "서울": "서울",
    "경기": "경기",
    "인천": "인천",
    "강원": "강원",
    "충북": "충북",
    "충남": "충남",
    "대전": "대전",
    "세종": "세종",
    "전북": "전북",
    "전남": "전남",
    "광주": "광주",
    "경북": "경북",
    "경남": "경남",
    "대구": "대구",
    "울산": "울산",
    "부산": "부산",
    "제주": "제주"
}

# 월별 강우일수 데이터 (지역별)
RAIN_DAYS = {
    "서울": {1: 5.2, 2: 5.3, 3: 6.8, 4: 8.1, 5: 9.1, 6: 10.5, 7: 16.0, 8: 13.4, 9: 9.1, 10: 6.8, 11: 8.5, 12: 6.4},
    "경기": {1: 5.5, 2: 5.5, 3: 7.0, 4: 8.3, 5: 9.3, 6: 10.7, 7: 16.2, 8: 13.6, 9: 9.3, 10: 7.0, 11: 8.7, 12: 6.6},
    "인천": {1: 5.8, 2: 5.6, 3: 7.2, 4: 8.5, 5: 9.5, 6: 10.9, 7: 16.5, 8: 13.8, 9: 9.5, 10: 7.2, 11: 8.9, 12: 6.8},
    "강원": {1: 6.0, 2: 5.8, 3: 7.5, 4: 8.8, 5: 9.8, 6: 11.2, 7: 17.0, 8: 14.2, 9: 9.8, 10: 7.5, 11: 9.2, 12: 7.0},
    "충북": {1: 5.3, 2: 5.4, 3: 6.9, 4: 8.2, 5: 9.2, 6: 10.6, 7: 16.1, 8: 13.5, 9: 9.2, 10: 6.9, 11: 8.6, 12: 6.5},
    "충남": {1: 5.4, 2: 5.5, 3: 7.0, 4: 8.3, 5: 9.3, 6: 10.7, 7: 16.3, 8: 13.7, 9: 9.3, 10: 7.0, 11: 8.7, 12: 6.6},
    "대전": {1: 5.2, 2: 5.3, 3: 6.8, 4: 8.1, 5: 9.1, 6: 10.5, 7: 16.0, 8: 13.4, 9: 9.1, 10: 6.8, 11: 8.5, 12: 6.4},
    "세종": {1: 5.3, 2: 5.4, 3: 6.9, 4: 8.2, 5: 9.2, 6: 10.6, 7: 16.1, 8: 13.5, 9: 9.2, 10: 6.9, 11: 8.6, 12: 6.5},
    "전북": {1: 5.1, 2: 5.2, 3: 6.7, 4: 8.0, 5: 9.0, 6: 10.4, 7: 15.9, 8: 13.3, 9: 9.0, 10: 6.7, 11: 8.4, 12: 6.3},
    "전남": {1: 5.0, 2: 5.1, 3: 6.6, 4: 7.9, 5: 8.9, 6: 10.3, 7: 15.8, 8: 13.2, 9: 8.9, 10: 6.6, 11: 8.3, 12: 6.2},
    "광주": {1: 5.1, 2: 5.2, 3: 6.7, 4: 8.0, 5: 9.0, 6: 10.4, 7: 15.9, 8: 13.3, 9: 9.0, 10: 6.7, 11: 8.4, 12: 6.3},
    "경북": {1: 4.8, 2: 4.9, 3: 6.4, 4: 7.7, 5: 8.7, 6: 10.1, 7: 15.5, 8: 12.9, 9: 8.7, 10: 6.4, 11: 8.1, 12: 6.0},
    "경남": {1: 4.6, 2: 4.7, 3: 6.2, 4: 7.5, 5: 8.5, 6: 9.9, 7: 15.3, 8: 12.7, 9: 8.5, 10: 6.2, 11: 7.9, 12: 5.8},
    "대구": {1: 4.5, 2: 4.6, 3: 6.1, 4: 7.4, 5: 8.4, 6: 9.8, 7: 15.2, 8: 12.6, 9: 8.4, 10: 6.1, 11: 7.8, 12: 5.7},
    "울산": {1: 4.4, 2: 4.5, 3: 6.0, 4: 7.3, 5: 8.3, 6: 9.7, 7: 15.1, 8: 12.5, 9: 8.3, 10: 6.0, 11: 7.7, 12: 5.6},
    "부산": {1: 4.3, 2: 4.4, 3: 5.9, 4: 7.2, 5: 8.2, 6: 9.6, 7: 15.0, 8: 12.4, 9: 8.2, 10: 5.9, 11: 7.6, 12: 5.5},
    "제주": {1: 7.2, 2: 7.0, 3: 8.5, 4: 9.8, 5: 10.8, 6: 12.2, 7: 17.5, 8: 15.2, 9: 10.8, 10: 8.5, 11: 10.2, 12: 8.0}
}

# 월별 한랭일수 (일 최저기온 -10°C 이하)
COLD_DAYS = {
    "서울": {1: 8.5, 2: 5.2, 3: 0.3, 4: 0.0, 5: 0.0, 6: 0.0, 7: 0.0, 8: 0.0, 9: 0.0, 10: 0.0, 11: 0.5, 12: 6.3},
    "경기": {1: 9.0, 2: 5.5, 3: 0.4, 4: 0.0, 5: 0.0, 6: 0.0, 7: 0.0, 8: 0.0, 9: 0.0, 10: 0.0, 11: 0.6, 12: 6.7},
    "인천": {1: 7.5, 2: 4.5, 3: 0.2, 4: 0.0, 5: 0.0, 6: 0.0, 7: 0.0, 8: 0.0, 9: 0.0, 10: 0.0, 11: 0.4, 12: 5.5},
    "강원": {1: 12.0, 2: 8.0, 3: 1.5, 4: 0.1, 5: 0.0, 6: 0.0, 7: 0.0, 8: 0.0, 9: 0.0, 10: 0.0, 11: 1.2, 12: 9.5},
    "충북": {1: 8.7, 2: 5.3, 3: 0.4, 4: 0.0, 5: 0.0, 6: 0.0, 7: 0.0, 8: 0.0, 9: 0.0, 10: 0.0, 11: 0.5, 12: 6.5},
    "충남": {1: 8.3, 2: 5.0, 3: 0.3, 4: 0.0, 5: 0.0, 6: 0.0, 7: 0.0, 8: 0.0, 9: 0.0, 10: 0.0, 11: 0.5, 12: 6.2},
    "대전": {1: 8.0, 2: 4.8, 3: 0.3, 4: 0.0, 5: 0.0, 6: 0.0, 7: 0.0, 8: 0.0, 9: 0.0, 10: 0.0, 11: 0.4, 12: 6.0},
    "세종": {1: 8.2, 2: 4.9, 3: 0.3, 4: 0.0, 5: 0.0, 6: 0.0, 7: 0.0, 8: 0.0, 9: 0.0, 10: 0.0, 11: 0.5, 12: 6.1},
    "전북": {1: 7.0, 2: 4.0, 3: 0.2, 4: 0.0, 5: 0.0, 6: 0.0, 7: 0.0, 8: 0.0, 9: 0.0, 10: 0.0, 11: 0.3, 12: 5.2},
    "전남": {1: 5.5, 2: 3.0, 3: 0.1, 4: 0.0, 5: 0.0, 6: 0.0, 7: 0.0, 8: 0.0, 9: 0.0, 10: 0.0, 11: 0.2, 12: 4.0},
    "광주": {1: 6.0, 2: 3.5, 3: 0.1, 4: 0.0, 5: 0.0, 6: 0.0, 7: 0.0, 8: 0.0, 9: 0.0, 10: 0.0, 11: 0.2, 12: 4.5},
    "경북": {1: 9.5, 2: 6.0, 3: 0.5, 4: 0.0, 5: 0.0, 6: 0.0, 7: 0.0, 8: 0.0, 9: 0.0, 10: 0.0, 11: 0.7, 12: 7.2},
    "경남": {1: 4.5, 2: 2.5, 3: 0.1, 4: 0.0, 5: 0.0, 6: 0.0, 7: 0.0, 8: 0.0, 9: 0.0, 10: 0.0, 11: 0.1, 12: 3.3},
    "대구": {1: 7.8, 2: 4.5, 3: 0.2, 4: 0.0, 5: 0.0, 6: 0.0, 7: 0.0, 8: 0.0, 9: 0.0, 10: 0.0, 11: 0.4, 12: 5.8},
    "울산": {1: 3.5, 2: 2.0, 3: 0.1, 4: 0.0, 5: 0.0, 6: 0.0, 7: 0.0, 8: 0.0, 9: 0.0, 10: 0.0, 11: 0.1, 12: 2.5},
    "부산": {1: 2.0, 2: 1.0, 3: 0.0, 4: 0.0, 5: 0.0, 6: 0.0, 7: 0.0, 8: 0.0, 9: 0.0, 10: 0.0, 11: 0.0, 12: 1.5},
    "제주": {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 0.0, 6: 0.0, 7: 0.0, 8: 0.0, 9: 0.0, 10: 0.0, 11: 0.0, 12: 0.0}
}

# 월별 폭염일수 (일 최고기온 33°C 이상)
HOT_DAYS = {
    "서울": {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 0.2, 6: 1.5, 7: 5.8, 8: 6.2, 9: 1.0, 10: 0.0, 11: 0.0, 12: 0.0},
    "경기": {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 0.2, 6: 1.6, 7: 6.0, 8: 6.4, 9: 1.1, 10: 0.0, 11: 0.0, 12: 0.0},
    "인천": {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 0.1, 6: 1.2, 7: 4.5, 8: 4.8, 9: 0.8, 10: 0.0, 11: 0.0, 12: 0.0},
    "강원": {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 0.3, 6: 2.0, 7: 7.5, 8: 8.0, 9: 1.5, 10: 0.0, 11: 0.0, 12: 0.0},
    "충북": {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 0.2, 6: 1.7, 7: 6.5, 8: 6.9, 9: 1.2, 10: 0.0, 11: 0.0, 12: 0.0},
    "충남": {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 0.2, 6: 1.6, 7: 6.2, 8: 6.6, 9: 1.1, 10: 0.0, 11: 0.0, 12: 0.0},
    "대전": {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 0.2, 6: 1.8, 7: 7.0, 8: 7.5, 9: 1.3, 10: 0.0, 11: 0.0, 12: 0.0},
    "세종": {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 0.2, 6: 1.7, 7: 6.7, 8: 7.1, 9: 1.2, 10: 0.0, 11: 0.0, 12: 0.0},
    "전북": {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 0.3, 6: 2.2, 7: 8.5, 8: 9.0, 9: 1.8, 10: 0.0, 11: 0.0, 12: 0.0},
    "전남": {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 0.4, 6: 2.8, 7: 10.5, 8: 11.2, 9: 2.5, 10: 0.0, 11: 0.0, 12: 0.0},
    "광주": {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 0.3, 6: 2.5, 7: 9.5, 8: 10.0, 9: 2.0, 10: 0.0, 11: 0.0, 12: 0.0},
    "경북": {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 0.4, 6: 2.5, 7: 9.2, 8: 9.8, 9: 2.0, 10: 0.0, 11: 0.0, 12: 0.0},
    "경남": {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 0.5, 6: 3.0, 7: 11.0, 8: 11.8, 9: 2.8, 10: 0.0, 11: 0.0, 12: 0.0},
    "대구": {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 0.6, 6: 3.5, 7: 12.5, 8: 13.5, 9: 3.5, 10: 0.0, 11: 0.0, 12: 0.0},
    "울산": {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 0.3, 6: 2.0, 7: 7.8, 8: 8.5, 9: 1.8, 10: 0.0, 11: 0.0, 12: 0.0},
    "부산": {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 0.2, 6: 1.5, 7: 6.0, 8: 6.8, 9: 1.5, 10: 0.0, 11: 0.0, 12: 0.0},
    "제주": {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 0.1, 6: 0.5, 7: 2.5, 8: 3.5, 9: 0.8, 10: 0.0, 11: 0.0, 12: 0.0}
}


def get_total_non_work_days(region, start_date, end_date, check_rain=True, check_cold=True, check_hot=True):
    """
    지정된 기간 동안의 총 비작업일수를 계산합니다.
    
    Args:
        region (str): 지역명
        start_date (datetime or str): 시작일
        end_date (datetime or str): 종료일
        check_rain (bool): 강우일 포함 여부
        check_cold (bool): 한랭일 포함 여부
        check_hot (bool): 폭염일 포함 여부
    
    Returns:
        int: 총 비작업일수
    """
    # ✅ 날짜 검증 추가
    if not start_date or not end_date:
        return 0
    
    # datetime 객체로 변환
    if not isinstance(start_date, datetime):
        try:
            if isinstance(start_date, str):
                start_date = datetime.strptime(start_date, "%Y-%m-%d")
            else:
                start_date = datetime(start_date.year, start_date.month, start_date.day)
        except Exception as e:
            print(f"start_date 변환 실패: {e}")
            return 0
    
    if not isinstance(end_date, datetime):
        try:
            if isinstance(end_date, str):
                end_date = datetime.strptime(end_date, "%Y-%m-%d")
            else:
                end_date = datetime(end_date.year, end_date.month, end_date.day)
        except Exception as e:
            print(f"end_date 변환 실패: {e}")
            return 0
    
    # 지역 확인
    if region not in REGION_MAPPING:
        print(f"지원하지 않는 지역: {region}")
        return 0
    
    start_month = start_date.month
    end_month = end_date.month
    start_year = start_date.year
    end_year = end_date.year
    
    total_non_work_days = 0
    
    # 동일 연도 내
    if start_year == end_year:
        for m in range(start_month, end_month + 1):
            if check_rain:
                total_non_work_days += RAIN_DAYS.get(region, {}).get(m, 0)
            if check_cold:
                total_non_work_days += COLD_DAYS.get(region, {}).get(m, 0)
            if check_hot:
                total_non_work_days += HOT_DAYS.get(region, {}).get(m, 0)
    else:
        # 시작 연도 말까지
        for m in range(start_month, 13):
            if check_rain:
                total_non_work_days += RAIN_DAYS.get(region, {}).get(m, 0)
            if check_cold:
                total_non_work_days += COLD_DAYS.get(region, {}).get(m, 0)
            if check_hot:
                total_non_work_days += HOT_DAYS.get(region, {}).get(m, 0)
        
        # 중간 연도들 (전체 12개월)
        for year in range(start_year + 1, end_year):
            for m in range(1, 13):
                if check_rain:
                    total_non_work_days += RAIN_DAYS.get(region, {}).get(m, 0)
                if check_cold:
                    total_non_work_days += COLD_DAYS.get(region, {}).get(m, 0)
                if check_hot:
                    total_non_work_days += HOT_DAYS.get(region, {}).get(m, 0)
        
        # 종료 연도 처음부터 종료월까지
        for m in range(1, end_month + 1):
            if check_rain:
                total_non_work_days += RAIN_DAYS.get(region, {}).get(m, 0)
            if check_cold:
                total_non_work_days += COLD_DAYS.get(region, {}).get(m, 0)
            if check_hot:
                total_non_work_days += HOT_DAYS.get(region, {}).get(m, 0)
    
    return int(round(total_non_work_days))


def get_monthly_breakdown(region, start_date, end_date, check_rain=True, check_cold=True, check_hot=True):
    """
    월별 비작업일수 상세 내역을 반환합니다.
    
    Returns:
        list: [{"month": "2026-01", "rain": 5.2, "cold": 8.5, "hot": 0.0, "total": 13.7}, ...]
    """
    if not start_date or not end_date:
        return []
    
    # datetime 변환
    if not isinstance(start_date, datetime):
        try:
            if isinstance(start_date, str):
                start_date = datetime.strptime(start_date, "%Y-%m-%d")
            else:
                start_date = datetime(start_date.year, start_date.month, start_date.day)
        except:
            return []
    
    if not isinstance(end_date, datetime):
        try:
            if isinstance(end_date, str):
                end_date = datetime.strptime(end_date, "%Y-%m-%d")
            else:
                end_date = datetime(end_date.year, end_date.month, end_date.day)
        except:
            return []
    
    breakdown = []
    current_date = datetime(start_date.year, start_date.month, 1)
    end_month_date = datetime(end_date.year, end_date.month, 1)
    
    while current_date <= end_month_date:
        month_num = current_date.month
        rain = RAIN_DAYS.get(region, {}).get(month_num, 0) if check_rain else 0
        cold = COLD_DAYS.get(region, {}).get(month_num, 0) if check_cold else 0
        hot = HOT_DAYS.get(region, {}).get(month_num, 0) if check_hot else 0
        
        breakdown.append({
            "month": current_date.strftime("%Y-%m"),
            "rain": rain,
            "cold": cold,
            "hot": hot,
            "total": rain + cold + hot
        })
        
        # 다음 달로 이동
        if current_date.month == 12:
            current_date = datetime(current_date.year + 1, 1, 1)
        else:
            current_date = datetime(current_date.year, current_date.month + 1, 1)
    
    return breakdown