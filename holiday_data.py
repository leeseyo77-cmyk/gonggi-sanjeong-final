"""
상하수도 공사 법정공휴일 데이터 모듈
- 2026년 적정 공사기간 확보를 위한 가이드라인 부록1 기준
- 일요일 + 명절 + 국경일 + 기타 공휴일 포함
"""

from datetime import datetime, timedelta
from calendar import monthrange

# ══════════════════════════════════════════════════════════════
# 법정 공휴일수 (2026-2035년) - 가이드라인 부록 1
# ══════════════════════════════════════════════════════════════
# 일요일(52일) + 명절(6일) + 국경일(4일) + 기타(5일) + 대체공휴일 포함
# 단, 공직선거법 제34조에 따른 선거일과 정부 수시 지정일은 제외

LEGAL_HOLIDAYS = {
    2026: {1: 5, 2: 7, 3: 6, 4: 4, 5: 7, 6: 5, 7: 4, 8: 7, 9: 7, 10: 7, 11: 5, 12: 5},   # 소계: 69
    2027: {1: 6, 2: 7, 3: 5, 4: 4, 5: 7, 6: 4, 7: 4, 8: 6, 9: 7, 10: 8, 11: 4, 12: 6},   # 소계: 68
    2028: {1: 9, 2: 4, 3: 5, 4: 5, 5: 6, 6: 5, 7: 5, 8: 5, 9: 4, 10: 10, 11: 4, 12: 6},  # 소계: 68
    2029: {1: 5, 2: 7, 3: 5, 4: 5, 5: 7, 6: 5, 7: 5, 8: 5, 9: 8, 10: 6, 11: 4, 12: 6},   # 소계: 68
    2030: {1: 5, 2: 7, 3: 6, 4: 4, 5: 6, 6: 6, 7: 4, 8: 5, 9: 8, 10: 6, 11: 4, 12: 6},   # 소계: 67
    2031: {1: 8, 2: 4, 3: 7, 4: 4, 5: 6, 6: 6, 7: 4, 8: 6, 9: 5, 10: 8, 11: 5, 12: 5},   # 소계: 68
    2032: {1: 5, 2: 8, 3: 5, 4: 4, 5: 7, 6: 4, 7: 4, 8: 6, 9: 7, 10: 8, 11: 4, 12: 6},   # 소계: 68
    2033: {1: 7, 2: 6, 3: 5, 4: 4, 5: 7, 6: 5, 7: 5, 8: 5, 9: 7, 10: 7, 11: 4, 12: 5},   # 소계: 67
    2034: {1: 5, 2: 7, 3: 5, 4: 5, 5: 6, 6: 5, 7: 5, 8: 5, 9: 7, 10: 7, 11: 4, 12: 6},   # 소계: 67
    2035: {1: 5, 2: 7, 3: 5, 4: 5, 5: 7, 6: 5, 7: 5, 8: 5, 9: 8, 10: 6, 11: 4, 12: 6},   # 소계: 68
}


def get_legal_holidays(year, month):
    """
    특정 연/월의 법정 공휴일수를 반환합니다.
    
    Args:
        year (int): 연도
        month (int): 월
    
    Returns:
        int: 해당 월의 법정 공휴일수
    """
    if year in LEGAL_HOLIDAYS:
        return LEGAL_HOLIDAYS[year].get(month, 0)
    # 데이터가 없는 연도는 평균값 사용 (약 5.7일)
    return 6


def get_total_holidays(start_date, end_date):
    """
    지정 기간의 총 법정 공휴일수를 반환합니다.
    
    Args:
        start_date (datetime or str): 시작일
        end_date (datetime or str): 종료일
    
    Returns:
        int: 총 법정 공휴일수
    """
    if not start_date or not end_date:
        return 0
    
    # datetime 변환
    if not isinstance(start_date, datetime):
        if isinstance(start_date, str):
            start_date = datetime.strptime(start_date, "%Y-%m-%d")
        else:
            start_date = datetime(start_date.year, start_date.month, start_date.day)
    
    if not isinstance(end_date, datetime):
        if isinstance(end_date, str):
            end_date = datetime.strptime(end_date, "%Y-%m-%d")
        else:
            end_date = datetime(end_date.year, end_date.month, end_date.day)
    
    total = 0
    
    # 시작일과 종료일이 같은 월
    if start_date.year == end_date.year and start_date.month == end_date.month:
        # 부분 월: 전체 월 공휴일수 × (해당 기간 일수 / 달력 일수)
        month_holidays = get_legal_holidays(start_date.year, start_date.month)
        month_days = monthrange(start_date.year, start_date.month)[1]
        period_days = (end_date - start_date).days + 1
        total = round(month_holidays * period_days / month_days)
        return total
    
    # 시작 월 (부분)
    start_month_holidays = get_legal_holidays(start_date.year, start_date.month)
    start_month_days = monthrange(start_date.year, start_date.month)[1]
    start_remaining_days = start_month_days - start_date.day + 1
    total += round(start_month_holidays * start_remaining_days / start_month_days)
    
    # 중간 월 (전체)
    cur_year = start_date.year
    cur_month = start_date.month + 1
    if cur_month > 12:
        cur_month = 1
        cur_year += 1
    
    while (cur_year, cur_month) < (end_date.year, end_date.month):
        total += get_legal_holidays(cur_year, cur_month)
        cur_month += 1
        if cur_month > 12:
            cur_month = 1
            cur_year += 1
    
    # 종료 월 (부분)
    end_month_holidays = get_legal_holidays(end_date.year, end_date.month)
    end_month_days = monthrange(end_date.year, end_date.month)[1]
    total += round(end_month_holidays * end_date.day / end_month_days)
    
    return total


def calc_overlap_days(weather_days, holiday_days, calendar_days):
    """
    기상조건과 법정공휴일의 중복일수를 계산합니다.
    
    가이드라인 19페이지 공식:
        중복일수(C) = 기상조건 비작업일수(A) × 법정공휴일수(B) ÷ 달력일수
        (소수점 첫째자리에서 반올림)
    
    Args:
        weather_days (float): 기상조건 비작업일수 (A)
        holiday_days (float): 법정 공휴일수 (B)
        calendar_days (int): 달력 일수
    
    Returns:
        int: 중복일수
    """
    if calendar_days <= 0:
        return 0
    overlap = (weather_days * holiday_days) / calendar_days
    return round(overlap)


def get_total_non_work_days_with_holidays(
    weather_non_work_days,
    start_date,
    end_date,
    include_holidays=True,
    min_weekly_rest=True
):
    """
    가이드라인 공식에 따라 총 비작업일수를 계산합니다.
    
    공식: 비작업일수 = A + B - C
        A: 기상조건 비작업일수 (강우/한랭/폭염)
        B: 법정 공휴일수
        C: A × B ÷ 달력일수 (중복일수)
    
    Args:
        weather_non_work_days (int): 기상조건 비작업일수
        start_date (datetime): 시작일
        end_date (datetime): 종료일
        include_holidays (bool): 법정공휴일 포함 여부
        min_weekly_rest (bool): 주 40시간 근무제 보장 여부
    
    Returns:
        dict: {
            "total": 총 비작업일수,
            "weather": 기상조건 비작업일수 (A),
            "holidays": 법정 공휴일수 (B),
            "overlap": 중복일수 (C),
            "formula": "A + B - C"
        }
    """
    A = weather_non_work_days
    B = get_total_holidays(start_date, end_date) if include_holidays else 0
    
    # 달력 일수
    if isinstance(start_date, datetime) and isinstance(end_date, datetime):
        calendar_days = (end_date - start_date).days + 1
    else:
        calendar_days = 30  # 기본값
    
    C = calc_overlap_days(A, B, calendar_days)
    
    total = A + B - C
    
    # 주 40시간 근무제 보장 (한 주에 최소 1일 휴식)
    if min_weekly_rest:
        weeks = calendar_days / 7
        min_rest_days = round(weeks)  # 주당 1일
        if total < min_rest_days:
            total = min_rest_days
    
    return {
        "total": total,
        "weather": A,
        "holidays": B,
        "overlap": C,
        "calendar_days": calendar_days,
        "formula": f"{A} + {B} - {C} = {A + B - C}"
    }


def get_holiday_breakdown_monthly(start_date, end_date):
    """
    월별 법정 공휴일수 상세 내역을 반환합니다.
    
    Returns:
        list: [{"month": "2026-01", "holidays": 5}, ...]
    """
    if not start_date or not end_date:
        return []
    
    if not isinstance(start_date, datetime):
        if isinstance(start_date, str):
            start_date = datetime.strptime(start_date, "%Y-%m-%d")
        else:
            start_date = datetime(start_date.year, start_date.month, start_date.day)
    
    if not isinstance(end_date, datetime):
        if isinstance(end_date, str):
            end_date = datetime.strptime(end_date, "%Y-%m-%d")
        else:
            end_date = datetime(end_date.year, end_date.month, end_date.day)
    
    result = []
    cur = datetime(start_date.year, start_date.month, 1)
    
    while cur <= end_date:
        month_holidays = get_legal_holidays(cur.year, cur.month)
        result.append({
            "월": f"{cur.year}-{cur.month:02d}",
            "법정공휴일": month_holidays
        })
        # 다음 달로
        if cur.month == 12:
            cur = datetime(cur.year + 1, 1, 1)
        else:
            cur = datetime(cur.year, cur.month + 1, 1)
    
    return result