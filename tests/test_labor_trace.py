"""노무비 역산 — 직종별 인수는 병목 직종 기준(1조 = 직종별 1인)으로 환산하는지 확인한다.

품셈의 직종별 인수는 여러 직종이 함께 드는 조합이라, 인수를 모두 더해 나누면(1인 기준)
작업일수가 1.5~3배 길게 나왔다. 풍각 샘플의 호표로 실제 값을 대조한다(샘플 없으면 건너뜀).
실행: python -m unittest discover -s tests -v
"""
import unittest

import openpyxl

from _support import SAMPLES, have

import universal_parser as up

PUNGGAK = SAMPLES / "풍각 공공하수처리장 증설사업260609.xlsx"


@unittest.skipUnless(have(PUNGGAK), "풍각 샘플 없음")
class BottleneckTradeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        wb = openpyxl.load_workbook(PUNGGAK, data_only=True, keep_links=False)
        _, traces = up.parse_labor_derived_by_template(wb, up.detect_template(wb), with_trace=True)
        cls.by_ho = {t.get("호표"): t for t in traces.values()}

    def test_two_trades(self):
        # 제96호표 보강토옹벽 블록: 특별인부 0.21 + 보통인부 0.09 → 병목 특별인부
        t = self.by_ho[96]
        self.assertEqual(t["방식"], "직종별 인수(병목 직종)")
        self.assertEqual(t["병목 직종"], "특별인부")
        self.assertAlmostEqual(t["일작업량"], 1 / 0.21, places=3)
        self.assertAlmostEqual(t["1인 기준 일작업량"], 1 / 0.30, places=3)

    def test_four_trades(self):
        # 제100호표 세라믹 방수: 미장공 0.04 + 도장공 0.03 + 보통인부 0.03 + 특별인부 0.02
        t = self.by_ho[100]
        self.assertEqual(t["병목 직종"], "미장공")
        self.assertAlmostEqual(t["일작업량"], 25.0, places=3)        # 1인 기준이면 8.33
        self.assertAlmostEqual(t["인수합"], 0.12, places=6)


if __name__ == "__main__":
    unittest.main()
