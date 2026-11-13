"""앱 스모크·회귀 테스트 — 광주 제1정수장 샘플로 app.py를 헤드리스 실행한다(수 분).

샘플이 없거나 GONGGI_SKIP_APP=1이면 건너뛴다.
실행: python -m unittest discover -s tests -v

기준값(1,084일·2,645일 등)은 현재 산정 모델의 결과다. 모델을 의도적으로 바꿨다면
결과를 확인한 뒤 여기 숫자를 갱신한다.
"""
import json
import os
import unittest
import zipfile

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
        # 노무비 역산은 병목 직종 기준(1조 = 직종별 1인). 1인 합산 기준이던 때는 2,517일.
        # 일위대가 → 산근 A·Q를 따라가기 전에는 2,198일(철근 가공·조립이 역산 0.19t/일로 떨어짐)
        # 가이드라인을 내역서 Q값 다음 순위로 올린 뒤 1,079일(가시설 띠장·시스템비계·시트방수 등),
        # 1조 = 표준품셈 구성원(직종별 1인이 아니라 인수 비율 그대로의 구성원)으로 바꾼 뒤 808일,
        # 구조물을 차례로 시공(순차)하는 것을 기본으로 바꾼 뒤(사용자 기준 2026-11-09) 1,095일
        self.assertEqual(at.session_state["total_work_days"], 1095)
        days = {r["공종명_pure"]: r["작업일수(일)"] for r in at.session_state["work_result"]["rows"]}
        self.assertEqual((days["토공"], days["구조물공사"]), (808, 1095))     # 구조물별 일수 합

        # 구조물공사를 병행으로 고르면 가장 긴 구조물(오존접촉조 683일)이 공종 일수 — 토공 808일이 공기
        widget(at, "multiselect", "parallel_subgroup_widget").set_value(["구조물공사"]).run()
        self.assertEqual(exceptions(at), [])
        self.assertEqual(at.session_state["total_work_days"], 808)
        days = {r["공종명_pure"]: r["작업일수(일)"] for r in at.session_state["work_result"]["rows"]}
        self.assertEqual((days["토공"], days["구조물공사"]), (808, 683))

        # 상세표는 공기를 정하는 구간(🔴)의 항목이 맨 위 — 구조물공사는 오존접촉조 683일
        tables = [d.value for d in at.dataframe if "주공정" in d.value.columns]
        struct = next(df for df in tables if "오존접촉조" in set(df["구간"]))
        crit = struct["주공정"] == "🔴"
        self.assertTrue(crit.iloc[:int(crit.sum())].all())               # 🔴 행이 모두 위쪽에
        self.assertEqual(set(struct.loc[crit, "구간"]), {"오존접촉조"})
        self.assertEqual(int(struct.loc[crit, "작업일수"].sum()), 683)


@unittest.skipIf(SKIP, WHY)
class ProjectTabTest(unittest.TestCase):
    def test_civil_plus_architecture(self):
        at = app_test(civil=GWANGJU["토목"], discs={"건축": GWANGJU["건축"]})
        self.assertEqual(exceptions(at), [])
        # 기계·전기는 앱에서 다루지 않는다(설치만 검토 — 공기를 지배하지 않음)
        self.assertFalse(any(w.key == "disc_use_전기" for w in at.checkbox))
        # 토목 1,095(구조물 순차) + 건축(동 순차, 1조 = 품셈 구성원) 556 — 둘 다 병행이던 때는 808 + 404
        self.assertEqual(metric(at, "사업 전체 순작업일수"), "1651일")

        # 비작업일수 계산기로 넘기기
        button(at, "📥 전체 순작업일수").click().run()
        self.assertEqual(exceptions(at), [])
        self.assertEqual(at.session_state["weather_work_days"], 1651)

        # 건축을 공기에서 빼면 토목만 남는다
        widget(at, "checkbox", "disc_use_건축").set_value(False).run()
        self.assertEqual(exceptions(at), [])
        self.assertIsNone(metric(at, "사업 전체 순작업일수"))

    def test_crew_allocation_goes_to_larger_building(self):
        at = app_test(civil=GWANGJU["토목"], discs={"건축": GWANGJU["건축"]})
        widget(at, "number_input", "disc_tc_건축").set_value(8).run()
        self.assertEqual(exceptions(at), [])
        # 동 순차(기본) 8조 161일 — 동 병행이던 때는 면 작업 상한 4조가 한계라 107일
        self.assertEqual(metric(at, "건축 작업일수"), "161일")

    def test_user_adjustments(self):
        """토목처럼 항목 1일작업량 수정·제외와 공종 조수 지정이 건축 일수에 반영된다."""
        import discipline_parser as dp
        at = app_test(civil=GWANGJU["토목"], discs={"건축": GWANGJU["건축"]})
        widget(at, "number_input", "disc_tc_건축").set_value(8).run()
        items = at.session_state["project_disc_res"]["건축"]["items"]
        top = max(items, key=lambda i: i["days1"])
        # 가장 긴 항목을 빼면 그 항목이 사라지고 공기가 늘지 않는다
        at.session_state["disc_adj_건축"] = {"rates": {}, "excl": {dp.item_key(top)}, "pkg": {}}
        at.run()
        self.assertEqual(exceptions(at), [])
        now = at.session_state["project_disc_res"]["건축"]
        self.assertTrue(all(i["days1"] == 0 for i in now["items"] if dp.item_key(i) == dp.item_key(top)))
        self.assertLessEqual(now["total"], 161)
        # 모든 항목의 1일작업량을 아주 크게 → 공종마다 1일, 공종 조수 지정도 오류 없이 반영
        keys = {dp.item_key(i) for i in items}
        pk = now["res"]["packages"][0]
        at.session_state["disc_adj_건축"] = {
            "rates": {k: 1e9 for k in keys}, "excl": set(),
            "pkg": {dp.package_key(pk["discipline"], pk["unit_name"], pk["group"]): 2}}
        at.run()
        self.assertEqual(exceptions(at), [])
        now = at.session_state["project_disc_res"]["건축"]
        self.assertTrue(all(p["days"] <= 1 for p in now["res"]["packages"]))
        self.assertIn("사용자 조정", "".join(m.proto.delta for m in at.metric))


@unittest.skipIf(SKIP, WHY)
class CrewReverseAndScheduleTest(unittest.TestCase):
    """비작업일수 계산 후에만 나오는 토목 조수 역산 화면과 예정공정표 엑셀."""

    @classmethod
    def setUpClass(cls):
        cls.at = app_test(civil=GWANGJU["토목"], discs={"건축": GWANGJU["건축"]})
        widget(cls.at, "number_input", "disc_tc_건축").set_value(8).run()
        button(cls.at, "📊 비작업일수 계산").click().run()

    def test_project_settings_survive_weather_calc(self):
        # '비작업일수 계산'이 탭 중간에서 바로 재실행하면 사업 전체 탭 위젯이 그려지지 않아
        # Streamlit이 그 상태(조수·업로드)를 지웠다 → 재실행을 실행 끝으로 미뤄 고쳤다
        self.assertEqual(self.at.session_state["disc_tc_건축"], 8)
        self.assertEqual(metric(self.at, "건축 작업일수"), "161일")

    def test_monthly_table_matches_work_days(self):
        # 월별 비작업일수 표(부록 별표1)의 작업가능일수 합 = 순작업일수, 달력일수 합 = 작업 + 비작업
        wr = self.at.session_state["weather_result"]
        rows = wr["monthly_rows"]
        self.assertEqual(sum(r["대상일수"] - r["적용"] for r in rows), wr["work_days"])
        self.assertEqual(sum(r["대상일수"] for r in rows), wr["work_days"] + wr["non_work_days"])
        self.assertEqual((wr["work_end"] - wr["work_start"]).days + 1,
                         wr["work_days"] + wr["non_work_days"])

    def test_appendix_workbook_generated(self):
        at = self.at
        for f in OUT.glob("공사기간 산정 부록_*.xlsx"):
            f.unlink()
        button(at, "📥 부록 엑셀 생성").click().run()
        self.assertEqual(exceptions(at), [])
        # 부록이 다시 묶은 토목 작업일수가 앱 산정값과 같아야 한다(다르면 경고가 뜬다)
        self.assertFalse(any("부록의 토목 작업일수" in w.value for w in at.warning))
        files = list(OUT.glob("공사기간 산정 부록_*.xlsx"))
        self.assertEqual(len(files), 1)
        wb = openpyxl.load_workbook(files[0])
        self.assertEqual(wb.sheetnames[:3], ["총괄", "부록1. 작업일수 산정근거(토목)",
                                             "부록2. 작업일수 산정근거(건축)"])
        self.assertIn("별표1. 비작업일수 산정", wb.sheetnames)

    def test_report_hwpx_generated(self):
        at = self.at
        for f in OUT.glob("공사기간 산정 검토 보고서_*.hwpx"):
            f.unlink()
        button(at, "📄 보고서(hwpx) 생성").click().run()
        self.assertEqual(exceptions(at), [])
        files = list(OUT.glob("공사기간 산정 검토 보고서_*.hwpx"))
        self.assertEqual(len(files), 1)
        with zipfile.ZipFile(files[0]) as z:
            sec = z.read("Contents/section0.xml").decode("utf-8")
        self.assertNotIn("{{", sec)
        self.assertIn("토목공사 작업일수(대공종 중 최장)", sec)

    def test_unreachable_target_is_flagged(self):
        at = self.at
        widget(at, "number_input", "target_months_input").set_value(12.0).run()
        self.assertEqual(exceptions(at), [])
        self.assertIsNotNone(metric(at, "역산안 (필요 조수)"))
        self.assertTrue(any("목표 공기가 무리" in e.value for e in at.error))

        widget(at, "number_input", "target_months_input").set_value(240.0).run()
        self.assertFalse(any("목표 공기가 무리" in e.value for e in at.error))

    def test_area_caps_optional(self):
        at = self.at
        widget(at, "number_input", "target_months_input").set_value(24.0).run()
        # 기본은 꺼짐 — 개략 검토에서는 '항목당 최대 투입조수'만 쓴다
        self.assertFalse(widget(at, "checkbox", "area_caps_on").value)
        at.session_state["struct_areas"] = {"오존접촉조": {"slab": 1500.0, "wall": 2500.0}}
        widget(at, "checkbox", "area_caps_on").set_value(True).run()
        self.assertEqual(exceptions(at), [])
        tables = [d.value for d in at.dataframe if "근거" in d.value.columns]
        self.assertTrue(tables)
        t = tables[0]
        # 유로폼(오존접촉조): 벽체 양면 5,000 + 슬라브 1,500 = 6,500㎡ ÷ (1조 인원 × 30.2㎡)
        euro = t[(t["세부공종"] == "유로폼") & (t["구조물"] == "오존접촉조")]
        self.assertEqual(len(euro), 1)
        self.assertIn("6,500㎡", euro["근거"].iloc[0])
        # 면적을 넣지 않은 구조물은 내역서로 어림한 값이 기본으로 채워진다
        self.assertEqual(at.session_state["struct_areas"]["역T옹벽"]["src"], "내역서 어림")
        widget(at, "checkbox", "area_caps_on").set_value(False).run()

    def test_claude_crew_manual_mode(self):
        at = self.at
        widget(at, "number_input", "target_months_input").set_value(24.0).run()
        # 수동 모드가 기본 — 지시문에 목표·공종 표가 들어간다
        prompt = next(c.value for c in at.code if "목표 순작업일수" in c.value)
        self.assertIn("1조 기준 일수", prompt)
        key = "유로폼|0-7m:복잡"
        self.assertIn(key, prompt)
        ans = json.dumps({"recommendations": [{"key": key, "crew": 99, "reason": "테스트"}],
                          "overall": "검토", "risks": ["빠듯함"]}, ensure_ascii=False)
        widget(at, "text_area", "crew_ai_paste").set_value(ans).run()
        button(at, "📋 붙여 넣은 답 반영").click().run()
        self.assertEqual(exceptions(at), [])
        self.assertIsNotNone(metric(at, "AI 추천안 주공정"))
        # 상한(항목당 3조)을 넘는 추천은 상한으로 잘린다
        ai = next(d.value for d in at.dataframe if "AI 추천" in d.value.columns)
        row = ai[ai["세부공종"] == "유로폼"].iloc[0]
        self.assertEqual(int(row["AI 추천"]), 3)
        self.assertIn("상한 3조", row["조정"])
        button(at, "✅ AI 추천을").click().run()
        self.assertEqual(exceptions(at), [])
        ed = next(d.value for d in at.dataframe if "적용 조수" in d.value.columns and "필요 조수" in d.value.columns)
        self.assertEqual(int(ed[ed["세부공종"] == "유로폼"]["적용 조수"].iloc[0]), 3)
        button(at, "🗑️ 추천 지우기").click().run()

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
        # 사업 전체 공기 탭의 건축이 동 단위로 공정표에 들어간다
        col_a = [ws.cell(r, 1).value for r in range(5, ws.max_row + 1)]
        col_b = [ws.cell(r, 2).value for r in range(5, ws.max_row + 1)]
        self.assertIn("건축공사", col_a)
        self.assertIn("활성탄흡착지", col_b)
        with zipfile.ZipFile(files[0]) as z:
            self.assertIn("xl/drawings/drawing1.xml", z.namelist())


@unittest.skipIf(SKIP, WHY)
class ManualRateTest(unittest.TestCase):
    """수동입력 관리 → 추정값 검토에서 고친 값이 순작업일수·총 공사기간에 바로 반영된다."""

    def test_review_apply_reruns_and_recalculates_duration(self):
        at = app_test(civil=GWANGJU["토목"])
        # 구조물공사를 병행으로 — 토공(808일)이 공기, 구조물 항목은 주공정이 아니다
        widget(at, "multiselect", "parallel_subgroup_widget").set_value(["구조물공사"]).run()
        button(at, "📊 비작업일수 계산").click().run()
        total0 = at.session_state["weather_result"]["total_days"]
        # 추정값 검토(프래그먼트)의 반영 버튼 — 예전에는 앱 전체가 다시 실행되지 않았다
        ni = next(w for w in at.number_input if (w.key or "").startswith("lr_0_"))
        ni.set_value(1000.0).run()
        next(b for b in at.button if "건 반영" in b.label).click().run()
        self.assertEqual(exceptions(at), [])
        self.assertIn(ni.key[len("lr_0_"):], at.session_state["manual_rates"])
        # 주공정이 아닌 항목이면 공기가 그대로인 이유를 알려 준다
        self.assertTrue(any("주공정" in i.value for i in at.info))

        # 주공정 항목(토공 연암 터파기 535일)을 고치면 '비작업일수 계산'을 다시 누르지 않아도
        # 순작업일수와 총 공사기간이 함께 줄어든다
        key = "터파기(육상연암:H=2.7m이상)|굴삭기+대형브레이카 100%+크램쉘"
        at.session_state["manual_rates"] = {key: {"daily": 1e6, "unit": "㎥"}}
        at.run()
        at.run()                                   # 미뤄 둔 재실행(rerun_after_render)
        self.assertEqual(exceptions(at), [])
        wd = at.session_state["total_work_days"]
        self.assertLess(wd, 808)
        self.assertEqual(at.session_state["weather_result"]["work_days"], wd)
        self.assertLess(at.session_state["weather_result"]["total_days"], total0)


if __name__ == "__main__":
    unittest.main()
