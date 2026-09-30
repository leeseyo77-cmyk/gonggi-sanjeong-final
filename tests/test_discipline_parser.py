"""discipline_parser 단위 테스트 — 샘플 파일 없이 합성 항목으로 돈다(수 초).

실행: python -m unittest discover -s tests -v
"""
import unittest

import _support  # noqa: F401  (저장소 경로 등록)
import discipline_parser as dp


def _item(unit, group, days, disc="건축"):
    """직종 정보 없는 항목 = 노무비 역산 1직종. days1이 곧 작업량(인·일)."""
    return {"discipline": disc, "unit_name": unit, "group": group, "name": group,
            "qty": 1.0, "days1": days, "trades": {}}


class CrewTypeTest(unittest.TestCase):
    def test_classification(self):
        cases = [
            (("철근콘크리트공사", "노무비역산"), "면 작업"),
            (("가설공사", "노무비역산"), "면 작업"),
            (("5.3) 배관설치비", "플랜트용접공"), "선 작업"),
            (("2.1.1 전력 간선설비", "저압케이블전공"), "선 작업"),
            (("2.1.3 트레이설비", "내선전공"), "선 작업"),
            (("5.1) 기자재설치비", "기계설비공"), "설비 단위"),
            (("1.2 계측제어설비공사", ""), "설비 단위"),
            (("강관 항타공사", ""), "장비 의존"),
        ]
        for (group, lead), want in cases:
            with self.subTest(group=group):
                self.assertEqual(dp.crew_type(group, lead), want)

    def test_unknown_falls_back_to_middle(self):
        self.assertEqual(dp.crew_type("알 수 없는 공종", ""), "선 작업")


class DisciplineDaysTest(unittest.TestCase):
    ITEMS = [_item("A동", "철근콘크리트공사", 100), _item("B동", "철근콘크리트공사", 40)]

    def test_seq_sums_units(self):
        self.assertEqual(dp.discipline_days(self.ITEMS, "seq")["total"], 140)

    def test_parallel_takes_longest_unit(self):
        self.assertEqual(dp.discipline_days(self.ITEMS, "unit_parallel")["total"], 100)

    def test_uniform_crews_ignore_caps_by_default(self):
        # caps·crew_by_unit을 안 주면 예전처럼 상한 없이 균일 조수(설비 단위도 5조)
        items = [_item("A", "기자재설치비", 100)]
        self.assertEqual(dp.discipline_days(items, "unit_parallel", crews=5)["total"], 20)

    def test_caps_apply_per_package(self):
        items = [_item("A", "기자재설치비", 100)]          # 설비 단위, 상한 1
        res = dp.discipline_days(items, "unit_parallel", crew_by_unit={"A": 5}, caps={})
        self.assertEqual(res["total"], 100)


class PlanCrewsTest(unittest.TestCase):
    def test_budget_goes_to_larger_unit(self):
        items = [_item("큰동", "철근콘크리트공사", 300), _item("작은동", "철근콘크리트공사", 100)]
        r = dp.plan_crews(items, "unit_parallel", total_crews=4)
        self.assertEqual(r["crew_by_unit"], {"큰동": 3, "작은동": 1})
        self.assertEqual(r["total_days"], 100)

    def test_minimum_is_one_crew_per_unit(self):
        items = [_item("A", "철근콘크리트공사", 50), _item("B", "철근콘크리트공사", 50)]
        r = dp.plan_crews(items, "unit_parallel", total_crews=1)
        self.assertEqual(r["total_crews"], 2)

    def test_target_mode_stops_when_met(self):
        items = [_item("A", "철근콘크리트공사", 300)]
        r = dp.plan_crews(items, "unit_parallel", target_days=100)
        self.assertEqual((r["total_crews"], r["total_days"]), (3, 100))

    def test_cap_limits_target(self):
        items = [_item("A", "철근콘크리트공사", 400)]
        r = dp.plan_crews(items, "unit_parallel", caps={"면 작업": 2}, target_days=100)
        self.assertEqual((r["total_crews"], r["total_days"]), (2, 200))

    def test_unused_budget_when_capped(self):
        items = [_item("A", "기자재설치비", 100)]          # 설비 단위 → 더 넣어도 안 줄어듦
        r = dp.plan_crews(items, "unit_parallel", total_crews=10)
        self.assertEqual((r["total_crews"], r["total_days"]), (1, 100))

    def test_seq_mode_spreads_over_all_units(self):
        items = [_item("A", "철근콘크리트공사", 100), _item("B", "철근콘크리트공사", 100)]
        r = dp.plan_crews(items, "seq", total_crews=4)
        self.assertEqual(r["total_days"], 100)            # 50 + 50


if __name__ == "__main__":
    unittest.main()
