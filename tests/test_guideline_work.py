"""가이드라인 부록4 1일 작업량(guideline_work)과 분야 1조 = 표준품셈 구성원 환산(샘플 불필요).

실행: python -m unittest discover -s tests -v
"""
import unittest

import _support  # noqa: F401  (저장소 경로 등록)
import discipline_parser as dp
import guideline_work as gw


class MatchTest(unittest.TestCase):
    def daily(self, *a):
        m = gw.match(*a)
        return round(m["daily"], 2) if m else None

    def test_conditions_must_match(self):
        # 유로폼 보통·7m 이하 설치 50 + 해체 115 → 일수 합산
        self.assertEqual(self.daily("유로폼", "보통, 0-7m이하", "㎡", "건축"), round(1 / (1 / 50 + 1 / 115), 2))
        # 가이드라인에 '복잡'·7m 초과 값은 없다 → 다음 순위로
        self.assertIsNone(gw.match("유로폼 설치 및 해체", "6회, 복잡, 수직고 7m 초과 10m까지", "M2", "건축"))
        self.assertIsNone(gw.match("시스템비계 설치해체", "H=10m초과~20m이하", "M2", "건축"))
        self.assertIsNone(gw.match("모르타르 바름", "내벽2회, 18mm, 3.6m 초과", "M2", "건축"))
        self.assertIsNone(gw.match("콘크리트면 부분마감", "3.6m 초과, 벽체", "㎡", "건축"))

    def test_waterproofing_and_finishes(self):
        self.assertEqual(self.daily("우레탄도막방수", "바닥, 비노출 3MM", "㎡", "건축"), 200)
        self.assertEqual(self.daily("바탕처리", "보통, 바닥", "M2", "건축"), 100)
        self.assertEqual(self.daily("시스템동바리 설치 및 해체", "5m 초과~10m 이하", "공M3", "건축"), 69)

    def test_rebar_civil_vs_building(self):
        self.assertEqual(self.daily("철근 현장가공 및 조립", "TYPE-1", "ton", "건축"), round(1 / (1 / 4 + 1 / 3.4), 2))
        self.assertEqual(self.daily("철근가공조립", "철근 공장가공및조립,Type-Ⅱ", "TON", "토목"), 2.5)

    def test_unit_must_match(self):
        self.assertIsNone(gw.match("우레탄도막방수", "바닥", "M", "건축"))
        self.assertIsNone(gw.match("흙운반", "덤프", "㎥", "토목"))


class CrewTest(unittest.TestCase):
    def test_ratio_crew(self):
        # 인수 비율 그대로(가장 적게 드는 직종 1인, 나누어떨어지지 않으면 올림)
        self.assertEqual(dp.item_crew("x", "", {"형틀목공": 0.176, "보통인부": 0.044}),
                         ({"형틀목공": 4, "보통인부": 1}, "품셈 인수 구성"))
        self.assertEqual(dp.item_crew("x", "", {"방수공": 0.03, "보통인부": 0.012})[0], {"방수공": 3, "보통인부": 1})
        self.assertEqual(dp.item_crew("x", "", {"보통인부": 0.13})[0], {"보통인부": 1})

    def test_item_crew_prefers_pumsem(self):
        crew, src = dp.item_crew("유로폼 설치 및 해체", "복잡", {"형틀목공": 0.2, "보통인부": 0.03})
        self.assertEqual((crew, src), ({"형틀목공": 4, "보통인부": 1}, "품셈 작업조"))


if __name__ == "__main__":
    unittest.main()
