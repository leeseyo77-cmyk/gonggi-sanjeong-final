"""앱 스모크·회귀 테스트 — 광주 제1정수장 샘플로 app.py를 헤드리스 실행한다(수 분).

샘플이 없거나 GONGGI_SKIP_APP=1이면 건너뛴다.
실행: python -m unittest discover -s tests -v

기준값(2,517일·4,078일 등)은 현재 산정 모델의 결과다. 모델을 의도적으로 바꿨다면
결과를 확인한 뒤 여기 숫자를 갱신한다.
"""
import os
import unittest

import openpyxl

from _support import (GWANGJU, OUT, app_test, button, exceptions, have, metric,
                      widget)

SKIP = os.environ.get("GONGGI_SKIP_APP") == "1" or not have(*GWANGJU.values())
WHY = "광주 샘플이 없거나 GONGGI_SKIP_APP=1"


@unittest.skipIf(SKIP, WHY)
class CivilTest(unittest.TestCase):
    def test_civil_work_days_baseline(self):
        at = app_test(civil=GWANGJU["토목"])
        self.assertEqual(exceptions(at), [])
        self.assertEqual(at.session_state["total_work_days"], 2517)


@unittest.skipIf(SKIP, WHY)
class ProjectTabTest(unittest.TestCase):
    def test_defaults_then_include_electrical(self):
        at = app_test(civil=GWANGJU["토목"],
                      discs={"건축": GWANGJU["건축"], "전기": GWANGJU["전기"]})
        self.assertEqual(exceptions(at), [])
        # 기본: 토목 + 건축(동 병행) — 전기는 공기 미반영
        self.assertEqual(metric(at, "사업 전체 순작업일수"), "4078일")

        # 전기를 켜면 건축 마감 후 착수(기본) → 2,517 + 1,561 + 984
        widget(at, "checkbox", "disc_use_전기").set_value(True).run()
        self.assertEqual(exceptions(at), [])
        self.assertEqual(metric(at, "사업 전체 순작업일수"), "5062일")

        # 비작업일수 계산기로 넘기기
        button(at, "📥 전체 순작업일수").click().run()
        self.assertEqual(exceptions(at), [])
        self.assertEqual(at.session_state["weather_work_days"], 5062)

    def test_crew_allocation_goes_to_larger_building(self):
        at = app_test(civil=GWANGJU["토목"], discs={"건축": GWANGJU["건축"]})
        widget(at, "number_input", "disc_tc_건축").set_value(8).run()
        self.assertEqual(exceptions(at), [])
        # 면 작업 상한 4조 → 활성탄흡착지 4조가 한계, 402일
        self.assertEqual(metric(at, "건축 작업일수"), "402일")


@unittest.skipIf(SKIP, WHY)
class CrewReverseAndScheduleTest(unittest.TestCase):
    """비작업일수 계산 후에만 나오는 토목 조수 역산 화면과 예정공정표 엑셀."""

    @classmethod
    def setUpClass(cls):
        cls.at = app_test(civil=GWANGJU["토목"], discs={"건축": GWANGJU["건축"]})
        button(cls.at, "📊 비작업일수 계산").click().run()

    def test_unreachable_target_is_flagged(self):
        at = self.at
        widget(at, "number_input", "target_months_input").set_value(24.0).run()
        self.assertEqual(exceptions(at), [])
        self.assertIsNotNone(metric(at, "역산안 (필요 조수)"))
        self.assertTrue(any("목표 공기가 무리" in e.value for e in at.error))

        widget(at, "number_input", "target_months_input").set_value(240.0).run()
        self.assertFalse(any("목표 공기가 무리" in e.value for e in at.error))

    def test_schedule_excel_states_crew_assumption(self):
        at = self.at
        for f in OUT.glob("예정공정표_*.xlsx"):
            f.unlink()
        button(at, "📥 예정공정표 엑셀 생성").click().run()
        self.assertEqual(exceptions(at), [])
        files = list(OUT.glob("예정공정표_*.xlsx"))
        self.assertEqual(len(files), 1)
        ws = openpyxl.load_workbook(files[0]).active
        notes = [c.value for row in ws.iter_rows() for c in row
                 if isinstance(c.value, str) and c.value.startswith("※")]
        self.assertTrue(any("가정값" in n for n in notes), notes)


if __name__ == "__main__":
    unittest.main()
