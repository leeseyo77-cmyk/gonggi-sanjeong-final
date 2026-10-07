"""월별 비작업일수 — 공공 건설공사의 공사기간 산정기준 제8조 방식(수 초, 샘플 불필요).

실행: python -m unittest discover -s tests -v
"""
import unittest

import _support  # noqa: F401  (저장소 경로 등록)
import holiday_data as hd


class RoundTest(unittest.TestCase):
    def test_half_up_like_excel(self):
        # 파이썬 round(2.5)=2(짝수 반올림)와 달리 엑셀 ROUND·고시 반올림은 3
        self.assertEqual(hd.round_half_up(2.5), 3)
        self.assertEqual(hd.round_half_up(0.5), 1)
        self.assertEqual(hd.round_half_up(12.49), 12)


class MonthlyNonWorkTest(unittest.TestCase):
    def test_monthly_overlap_and_minimum(self):
        # 2026-10: A 6.2 + B 7 − C round(6.2×7/31=1.4)=1 → 12 (8일 이상)
        # 2026-04: A 1.0 + B 4 − C round(0.13)=0 → 5 → 주 40시간 근무제 8일 적용
        r = hd.monthly_non_work([{"월": "2026-10", "일수": 31, "합계": 6.2},
                                 {"월": "2026-04", "일수": 30, "합계": 1.0}])
        oct_, apr = r["rows"]
        self.assertEqual((oct_["B"], oct_["C"], oct_["비작업"], oct_["적용"]), (7, 1, 12, 12))
        self.assertEqual((apr["B"], apr["비작업"], apr["최소"], apr["적용"]), (4, 5, 8, 8))
        self.assertEqual(r["total"], 20)

    def test_partial_month_prorated(self):
        # 2027-02의 10일만: 공휴일 7×10/28=2.5→3, 최소 8×10/28=2.86→3
        r = hd.monthly_non_work([{"월": "2027-02", "일수": 10, "합계": 2.7}])
        row = r["rows"][0]
        self.assertEqual((row["B"], row["최소"]), (3, 3))

    def test_switches(self):
        r = hd.monthly_non_work([{"월": "2026-04", "일수": 30, "합계": 1.0}],
                                include_holidays=False, min_weekly_rest=False)
        self.assertEqual((r["rows"][0]["B"], r["total"]), (0, 1))


if __name__ == "__main__":
    unittest.main()
