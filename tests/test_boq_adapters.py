"""단산·일위·표준 참조형(경주 상수도형) 변환기 — 합성 내역서로 판별·계층·참조 연결·Q값을 확인한다(샘플 불필요).

실행: python -m unittest discover -s tests -v
"""
import unittest

import openpyxl

import _support  # noqa: F401  (저장소 경로 등록)
import universal_parser as up


def _book():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("설계내역서")
    ws.append([None, "내역서"])
    ws.append([None, "공종", "품    명", "규   격", "수 량", "단위"])
    rows = [
        ["◎", "토목공사"], ["A.", "관로공사"], ["1.", "1-1구간(가-나)"], ["1.1", "토공"], ["1)", "터파기"],
        [None, "터파기(토사)", "소형장비", 100, "㎥", None, None, None, None, None, None, None, None, "단산9참조"],
        ["1.2", "가시설공"],
        [None, "말뚝박기용천공(토사)", "Φ500mm이하", 50, "M", None, None, None, None, None, None, None, None,
         "단산2참조"],
        ["8.", "주요자재대"],                       # 설계자 번호 실수 — 구간이 아니라 대공종
        [None, "주철관", "D400", 10, "본"],
        ["2.", "1-2구간(나-다)"], ["2.1", "관로공"],
        [None, "주철관부설", "Φ400mm", 20, "본", None, None, None, None, None, None, None, None, "표준1참조"],
        [None, "띠장 설치", "L=5m", 8, "본", None, None, None, None, None, None, None, None, "일위1참조"],
        ["B.", "가압장공사"], ["1.", "토공"],
        [None, "터파기(토사)", "소형장비", 30, "㎥", None, None, None, None, None, None, None, None, "단산1참조"],
        ["◎", "관급자재대"], ["A.", "토목공사"],
        [None, "레미콘", "25-24-15", 40, "㎥"],
    ]
    for r in rows:
        ws.append([None] + r)

    t = wb.create_sheet("단가산출총괄표")
    t.append([None, "No", "품명", "규격", "단위", "합계", "노무비"])
    t.append([None, 1, "터파기(토사)", "소형장비", "㎥", 5000, 4000])
    t.append([None, 2, "말뚝박기용천공(토사)", "Φ500mm이하", "M", 19062, 10335])
    t.append([None, 9, "콘크리트포장깨기", "소형장비", "㎥", 58631, 44744])
    d = wb.create_sheet("단가산출근거")
    for line, amt in [("1.터파기(토사) (소형장비) [㎥]", 5000), ("        Q = 24.33 ㎥/hr", None),
                      ("2.말뚝박기용천공(토사) (Φ500mm이하) [M]", 19062),
                      ("        TQ = (3 + 9.6 + 2) / 0.8 ÷ 60 = 0.3hr/본", None),
                      ("        T = TQ / L1 = 0.3 / 10 = 0.03hr/M", None),
                      ("9.콘크리트포장깨기(소형장비)[㎥]", 58631), ("        Q = 1.75 ㎥/hr", None)]:
        d.append([None, line, amt])

    i = wb.create_sheet("일위대가")
    i.append([None, "호  표", "품    명", "규   격", "수 량", "단위"])
    i.append([None, "제1호표", "띠장 설치", "L=5m", None, "본", None, 200000, None, 180000])
    i.append([None, None, "비 계 공", None, 0.6, "인", 280000, None, 280000, 168000])
    i.append([None, None, "보통인부", None, 0.2, "인", 170000, None, 170000, 34000])
    wb.create_sheet("일위대가총괄표").append([None, "제1호표", "띠장 설치", "L=5m", "본"])
    s = wb.create_sheet("표준시장단가")
    s.append([None, "No", "코드", "품명", "규격", "단위", "합계", "노무비"])
    s.append([None, 1, "X1", "주철관부설", "Φ400mm", "본", 70000, 61016])
    n = wb.create_sheet("노임단가")
    for k, (nm, w) in enumerate([("보통인부", 169804), ("배관공", 250572), ("비 계 공", 280000)], 1):
        n.append([None, k, nm, None, "인", w])
    return wb


class DansanRefTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        wb = _book()
        cls.tmpl = up.detect_template(wb)
        cls.wb = up.adapt_workbook(wb, cls.tmpl)
        cls.items = up.parse_items_generic(cls.wb, cls.tmpl)
        cls.cd = up.parse_unit_price_generic(cls.wb, cls.tmpl)

    def test_detected(self):
        self.assertEqual(self.tmpl["id"], "dansan_ref")

    def test_tree(self):
        got = [(i["name"], i["category"], i["line"]) for i in self.items]
        self.assertEqual(got, [
            ("터파기(토사)", "토공", "1-1구간(가-나)"),
            ("말뚝박기용천공(토사)", "가시설공", "1-1구간(가-나)"),
            ("주철관", "주요자재대", "1-1구간(가-나)"),          # '8.'이지만 구간 안 대공종
            ("주철관부설", "관로공", "1-2구간(나-다)"),
            ("띠장 설치", "관로공", "1-2구간(나-다)"),
            ("터파기(토사)", "토공", "가압장공사"),             # 구간 없는 영문 단계 = 라인
            ("레미콘", "관급자재대 (공사비 외)", ""),
        ])

    def test_misnumbered_ref_relinked_by_name(self):
        # 비고 '단산9참조'는 콘크리트포장깨기를 가리킨다 → 품명·규격이 같은 단산 1로 바로잡는다
        self.assertEqual(self.items[0]["code"], 1)
        self.assertEqual(self.items[4]["code_alt"], 1)                       # 일위1
        self.assertEqual(self.items[3]["code_alt"], 90001)                   # 표준1 → 합성 블록

    def test_q_uses_header_unit(self):
        self.assertAlmostEqual(self.cd[1][0], 24.33 * 8, places=2)
        # 공당 0.3hr/본이 아니라 머리 단위 [M]과 같은 0.03hr/M → 266.7M/일
        self.assertEqual(self.cd[2][1], "M")
        self.assertAlmostEqual(self.cd[2][0], 8 / 0.03, places=1)

    def test_labor_blocks(self):
        res, tr = up.parse_labor_derived_by_template(self.wb, self.tmpl, with_trace=True, code_daily=self.cd)
        # 일위대가 인수: 비계공 0.6 · 보통인부 0.2 → 구성원 비계공 3 + 보통인부 1 → 5본/조·일
        self.assertAlmostEqual(res[1][0], 5.0, places=3)
        # 표준시장단가(노무비만) → 품셈 관 작업조(배관공 3 + 보통인부 1)로 환산
        self.assertIn(90001, res)
        self.assertEqual(tr[90001]["방식"], "노무비 ÷ 작업조 노임(품셈)")


class DiagnoseTest(unittest.TestCase):
    """읽지 못한 이유와 엑셀에서 고칠 방법(boq_diagnose)."""

    def test_supported_file_is_clean(self):
        import boq_diagnose as bdg
        wb = _book()
        levels = [d["level"] for d in bdg.diagnose(wb, up.detect_template(wb))]
        self.assertNotIn("error", levels)
        self.assertNotIn("warn", levels)

    def test_unknown_reference_explains_fix(self):
        import boq_diagnose as bdg
        wb = _book()
        ws = wb["설계내역서"]
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and "참조" in c.value:
                    c.value = c.value.replace("단산", "단가 ").replace("일위", "일위 ").replace("참조", "")
        self.assertIsNone(up.detect_template(wb) and up.detect_template(wb).get("adapter"))
        diag = bdg.diagnose(wb, None)
        bad = [d for d in diag if d["level"] == "error"]
        self.assertTrue(any("참조" in d["title"] and "산근 N호표" in d["fix"] for d in bad), diag)

    def test_missing_header(self):
        import boq_diagnose as bdg
        wb = openpyxl.Workbook()
        wb.active.title = "설계내역서"
        wb.active.append(["", "터파기", "토사", 10, "㎥"])
        diag = bdg.diagnose(wb, None)
        self.assertEqual(diag[1]["level"], "error")
        self.assertIn("품명", diag[1]["fix"])


if __name__ == "__main__":
    unittest.main()
