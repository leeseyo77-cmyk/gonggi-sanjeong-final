"""crew_space — 면적 기준 동시 투입 상한(터파기·철근·거푸집) 단위 테스트(샘플 불필요).

실행: python -m unittest discover -s tests -v
"""
import unittest

import _support  # noqa: F401  (저장소 경로 등록)
import crew_space as cs
import universal_parser as up


class ClassTest(unittest.TestCase):
    def test_trade_class(self):
        self.assertEqual(cs.trade_class("터파기(육상토사:H=0-2.7m이하)"), "터파기")
        self.assertEqual(cs.trade_class("철근 현장가공 및 조립"), "철근")
        self.assertEqual(cs.trade_class("유로폼"), "거푸집")
        self.assertEqual(cs.trade_class("합판거푸집"), "거푸집")
        self.assertIsNone(cs.trade_class("철근운반(하치장-현장)"))
        self.assertIsNone(cs.trade_class("철근콘크리트 깨기"))
        self.assertIsNone(cs.trade_class("레미콘타설"))


class PersonsTest(unittest.TestCase):
    def test_pumsem_crews(self):
        # 2026 표준품셈: 철근 조립 토목 Ⅰ-1 8인 / Ⅱ-1 7인, 유로폼 5인, 합판 7인
        self.assertEqual(cs.pumsem_persons("철근", "철근 현장가공 및 조립", "TYPE-Ⅰ"), 8)
        self.assertEqual(cs.pumsem_persons("철근", "철근 현장가공 및 조립", "TYPE-Ⅱ"), 7)
        self.assertEqual(cs.pumsem_persons("거푸집", "유로폼", "0-7m:복잡"), 5)
        self.assertEqual(cs.pumsem_persons("거푸집", "합판거푸집", "0-7m:3회"), 7)

    def test_labor_derived_crew_is_pumsem_members(self):
        # 노무비 역산이면 1조 인원은 품셈 구성원 인원(역산 근거의 '작업조 인원')
        self.assertEqual(cs.crew_persons("거푸집", "유로폼", "", "노무비역산(형틀목공)", persons=4), 4)
        self.assertEqual(cs.crew_persons("거푸집", "유로폼", "", "일위대가→산근(A·Q)", persons=4), 5)

    def test_pumsem_crew_and_rate(self):
        # 품셈 작업조를 아는 공종은 그 작업조, 모르면 인수 비율 그대로(가장 적게 드는 직종 1인)
        self.assertEqual(cs.pumsem_crew("유로폼 설치 및 해체", "복잡", {"형틀목공": 0.2, "보통인부": 0.03}),
                         ({"형틀목공": 4, "보통인부": 1}, "품셈 작업조"))
        crew, src = cs.pumsem_crew("시트방수", "", {"방수공": 0.03, "보통인부": 0.01})
        self.assertEqual((crew, src), ({"방수공": 3, "보통인부": 1}, "품셈 인수 구성"))
        self.assertAlmostEqual(cs.crew_rate({"방수공": 0.03, "보통인부": 0.01}, crew)[0], 100.0)
        self.assertEqual(cs.ratio_crew({"보통인부": 0.13}), {"보통인부": 1})
        # 인수 없이 노무비만 있으면 품셈 단위당 인수·가이드라인 작업조로 구성원(수성페인트 11-2-2 → 6 : 1)
        self.assertEqual(cs.pumsem_crew("수성페인트 롤러칠", "내부 2회")[0], {"도장공": 6, "보통인부": 1})
        self.assertEqual(cs.pumsem_crew("무석면압출성형 시멘트판", "50T")[0], {"내장공": 2, "보통인부": 1})
        # 주 직종의 1/10도 안 드는 직종은 1인(구성원을 부풀리지 않는다)
        self.assertEqual(cs.ratio_crew({"철근공": 2.0, "보통인부": 0.5, "특별인부": 0.01}),
                         {"철근공": 4, "보통인부": 1, "특별인부": 1})


class CapTest(unittest.TestCase):
    def test_rebar_cap(self):
        # 기본은 최적 생산성 30.2㎡/인: 2,200㎡ ÷ (7인 × 30.2㎡) = 10.4 → 10조 (최소 19㎡면 16조)
        area = cs.work_area("철근", 1200, 1000)
        self.assertEqual(area, 2200)
        self.assertEqual(cs.area_cap("철근", area, 7), 10)
        self.assertEqual(cs.area_cap("철근", area, 7, person_area=cs.PERSON_AREA_MIN), 16)

    def test_excavator_reach_from_bucket(self):
        # 제원표: 0.17㎥ 6.305m, 0.59㎥ 8.5m, 0.92㎥ 9.945m — 사이는 보간
        self.assertEqual(cs.excavator_reach(0.59), 8.5)
        self.assertEqual(cs.excavator_reach(1.0), 9.95)
        self.assertEqual(cs.excavator_reach(0.12), 5.0)
        self.assertAlmostEqual(cs.excavator_reach(0.7), 8.98, places=2)
        self.assertIsNone(cs.excavator_reach(None))

    def test_form_cap_uses_both_wall_faces(self):
        area = cs.work_area("거푸집", 1200, 1000)          # 2×1,000 + 1,200
        self.assertEqual(area, 3200)
        self.assertEqual(cs.area_cap("거푸집", area, 5, person_area=30.2), 21)   # 3200/151 = 21.2

    def test_excavator_cap(self):
        # 굴착 평면 2,000㎡ ÷ (2×6m)² = 13.9 → 13대
        self.assertEqual(cs.area_cap("터파기", 2000, 1, excavator_radius=6), 13)

    def test_no_area_no_cap_and_minimum_one(self):
        self.assertIsNone(cs.area_cap("철근", 0, 7))
        self.assertEqual(cs.area_cap("철근", 50, 7), 1)    # 면적이 작아도 최소 1조

    def test_explain(self):
        self.assertIn("7인 × 19㎡", cs.explain("철근", 2200, 7, 19, 6, 16))
        self.assertIn("굴착기 1대", cs.explain("터파기", 2000, 1, 19, 6, 13))


class BucketParseTest(unittest.TestCase):
    def test_bucket_formats(self):
        # 산근·일위대가에 적힌 굴착기 규격 표기들
        for text, q in (("1.장비(굴착기 0.7㎥):90%", 0.7), ("굴삭기(무한궤도) / 0.2㎥", 0.2),
                        ("백호우(0.7 ㎥)         :  1  대", 0.7), ("1. 상차, 굴삭기 0.7 M3 <품셈 9-3>", 0.7),
                        ("BQC|굴착기1.0㎥", 1.0), ("굴삭기          0.12㎥                1", 0.12)):
            self.assertEqual(up._first_bucket([(None, text)], 0, 1), q, text)
        self.assertIsNone(up._first_bucket([(None, "덤프트럭 15ton")], 0, 1))


if __name__ == "__main__":
    unittest.main()
