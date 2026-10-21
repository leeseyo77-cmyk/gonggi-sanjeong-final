"""schedule_plan — 토목 공정 배치(작업 단위 막대, 구조물공 완료 시점) 단위 테스트(샘플 불필요).

실행: python -m unittest discover -s tests -v
"""
import unittest

import _support  # noqa: F401  (저장소 경로 등록)
import schedule_plan as sp


def _it(name, days, line="(공통)"):
    return {"name": name, "days": days, "line": line}


MAJORS = [
    {"name": "가시설공", "days": 40, "items": [_it("SHEET PILE 항타", 20), _it("SHEET PILE 인발", 10),
                                              _it("띠장 설치 및 철거", 10)]},
    {"name": "토공", "days": 60, "items": [_it("표토제거", 5), _it("터파기", 35), _it("되메우기", 20)]},
    {"name": "구조물공사", "days": 300, "items": [
        _it("버림콘크리트 타설", 10, "A동"), _it("유로폼", 100, "A동"), _it("철근 현장가공 및 조립", 60, "A동"),
        _it("시스템 동바리 설치 및 해체", 40, "A동"), _it("수조방수", 50, "A동"), _it("레미콘 타설", 40, "A동"),
        _it("유로폼", 80, "B동")]},
    {"name": "관로공", "days": 50, "items": [_it("관부설", 50, "유입관로"), _it("관부설", 30, "방류관로")]},
    {"name": "포장공", "days": 10, "items": [_it("아스팔트포장", 10)]},
]


class CategoryTest(unittest.TestCase):
    def test_civil_cat(self):
        self.assertEqual(sp.civil_cat("가시설공(처리시설)"), "가시설")
        self.assertEqual(sp.civil_cat("구조물공사"), "구조물공")
        self.assertEqual(sp.civil_cat("구내배관"), "관로공")
        self.assertEqual(sp.civil_cat("우수공"), "우수공")
        self.assertEqual(sp.civil_cat("조립식 PC 맨홀"), "부대공")


class ParallelPlanTest(unittest.TestCase):
    """최장(병행): 토목 공기 = 가장 긴 대공종(구조물공 300일), 막대는 모두 [0, 300] 안."""

    @classmethod
    def setUpClass(cls):
        cls.p = sp.plan_civil(MAJORS, combine_sum=False)
        cls.by = {(b["구분"], b["공종"]): b for b in cls.p["bars"]}

    def test_civil_days_and_bounds(self):
        self.assertEqual(self.p["civil_days"], 300)
        self.assertTrue(all(0 <= b["s"] < b["e"] <= 300 + 1e-6 for b in self.p["bars"]))

    def test_activity_bars(self):
        names = {k[1] for k in self.by}
        for n in ("SHEET-PILE 설치", "SHEET-PILE 철거", "표토제거", "터파기", "되메우기",
                  "무근콘크리트타설", "구조물 공사(거푸집)", "철근가공 및 조립", "방수공사",
                  "가설공사(비계 및 동바리) 설치", "가설공사(비계 및 동바리) 철거", "유입관로", "방류관로"):
            self.assertIn(n, names)

    def test_structure_activities_fill_structure_period(self):
        st_bars = sorted((b for b in self.p["bars"] if b["구분"] == "구조물공"), key=lambda b: b["s"])
        self.assertAlmostEqual(st_bars[0]["s"], self.p["struct_span"][0], places=2)
        self.assertAlmostEqual(st_bars[-1]["e"], self.p["struct_span"][1], places=2)
        for a, b in zip(st_bars, st_bars[1:]):                       # 차례로 이어진다
            self.assertAlmostEqual(a["e"], b["s"], places=2)
        self.assertEqual(st_bars[0]["공종"], "무근콘크리트타설")
        self.assertEqual(st_bars[-1]["공종"], "방수공사")

    def test_order_logic(self):
        # 터파기는 처음부터, 되메우기·가시설 철거는 구조물공 이후(토목 기간 안에 들어가게 당김)
        self.assertEqual(self.by[("토공", "표토제거")]["s"], 0)
        self.assertLessEqual(self.by[("토공", "되메우기")]["e"], 300 + 1e-6)
        self.assertGreater(self.by[("가시설", "SHEET-PILE 철거")]["s"], self.by[("가시설", "SHEET-PILE 설치")]["s"])
        # 관로 구간은 병행 → 같은 시작
        self.assertEqual(self.by[("관로공", "유입관로")]["s"], self.by[("관로공", "방류관로")]["s"])
        # 포장은 끝에 맞춘다(구간이 없으면 대공종 이름으로 막대 하나)
        self.assertAlmostEqual(self.by[("포장공", "포장공")]["e"], 300, places=2)

    def test_project_days_architecture_after_structure(self):
        # 구조물공 300일이 토목 공기와 같으므로 건축은 300일차부터
        self.assertEqual(self.p["struct_end"], 300)
        self.assertEqual(sp.project_days(300, self.p["struct_end"], arch=200), 500)


class SequentialPlanTest(unittest.TestCase):
    """합산(순차): 대공종을 공사 순서대로 잇는다 — 합계가 토목 공기."""

    def test_chain_sums_to_civil(self):
        p = sp.plan_civil(MAJORS, combine_sum=True)
        self.assertEqual(p["civil_days"], 460)
        self.assertAlmostEqual(max(b["e"] for b in p["bars"]), 460, places=2)
        # 가시설 설치 20일(설치 20 + 반반 5 → 25/40×40) → 토공 앞부분 → 구조물
        self.assertAlmostEqual(p["struct_span"][1] - p["struct_span"][0], 300, places=2)


class ProjectDaysTest(unittest.TestCase):
    def test_structure_ends_before_civil(self):
        # 구조물공이 600일에 끝나고 토목은 900일: 건축 500일 → 1,100일(토목+건축 1,400일이 아님)
        self.assertEqual(sp.project_days(900, 600, arch=500), 1100)
        self.assertEqual(sp.project_days(900, 600, arch=100), 900)       # 건축이 먼저 끝나면 토목
        self.assertEqual(sp.project_days(900, 600, arch=100, land=50), 950)


if __name__ == "__main__":
    unittest.main()
