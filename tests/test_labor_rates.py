"""labor_rates_2025 — 표준품셈 터파기·관부설의 1조(작업조) 일당 시공량(수 초, 샘플 불필요).

값은 2026 건설공사 표준품셈 원문과 대조했다: 3-2-1 굴착(인력/토사), 3-2-4 터파기(기계),
6-7-2 파형강관, 6-7-3 유리섬유복합관(비압력관).
실행: python -m unittest discover -s tests -v
"""
import unittest

import _support  # noqa: F401  (저장소 경로 등록)
import labor_rates_2025 as lr


class ExcavationTest(unittest.TestCase):
    def test_manual_crew_is_one_special_worker(self):
        # 3-2-1: 특별인부 1인 일당 보통토사 3.6 / 경질 2.7 / 자갈섞인 2.2 / 호박돌 섞인 1.2㎥
        self.assertEqual(lr.manual_excavation_daily("토사,육상")[0], 3.6)
        self.assertEqual(lr.manual_excavation_daily("경질토사")[0], 2.7)
        self.assertEqual(lr.manual_excavation_daily("자갈섞인토사")[0], 2.2)
        self.assertEqual(lr.manual_excavation_daily("호박돌섞인토사")[0], 1.2)

    def test_machine_crew_is_one_excavator(self):
        # 3-2-4: 굴착기 1대, 보통토사 TypeⅡ 420㎥ · 용수 25% 감 · 연암 TypeⅡ 28㎥
        self.assertEqual(lr.machine_excavation_daily("토사,육상")[0], 420.0)
        self.assertEqual(lr.machine_excavation_daily("토사,용수")[0], 315.0)
        self.assertEqual(lr.machine_excavation_daily("연암")[0], 28.0)


class PipeTest(unittest.TestCase):
    def test_per_pipe_tables_use_bottleneck_trade(self):
        # 본당 인수만 있는 표 → 1조 = 직종별 1인, 병목(배관공)으로 하루 시공량
        daily, crew, dia = lr.pipe_crew_daily("유리섬유복합관(직관)", 200)
        self.assertAlmostEqual(daily, 1 / 0.30, places=2)          # 지금까진 20본으로 잡혔다
        self.assertEqual(crew, {"배관공": 1, "보통인부": 1})
        self.assertAlmostEqual(lr.pipe_crew_daily("파형강관", 300)[0], 1 / 0.06, places=2)

    def test_crew_tables_keep_pumsem_crew(self):
        # 일당 시공량 표 → 품셈 작업조 그대로(주철관 D300: 배관공 3 + 보통인부 1, 9본)
        daily, crew, _ = lr.pipe_crew_daily("주철관 타이튼", 300)
        self.assertEqual(daily, 9.0)
        self.assertEqual(crew, {"배관공": 3, "보통인부": 1})

    def test_unknown_pipe_is_not_guessed(self):
        self.assertIsNone(lr.pipe_crew_daily("강관부설", 300))


if __name__ == "__main__":
    unittest.main()
