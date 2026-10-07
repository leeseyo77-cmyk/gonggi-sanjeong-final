"""report_hwpx — 공사기간 산정 검토 보고서(hwpx) 생성(수 초, 샘플 불필요).

한글 2022에서 열어 PDF로 내보내 확인했다(광주: 19쪽, 긴 표가 쪽마다 머리행을 반복하며 이어짐,
월별 비작업일수 계 1,627일). 여기서는 패키지 구조·자리표시·표 칸 맞춤을 본다.
실행: python -m unittest discover -s tests -v
"""
import copy
import io
import re
import unittest
import zipfile
from xml.etree import ElementTree as ET

import _support  # noqa: F401  (저장소 경로 등록)
import report_hwpx as rh
from test_report_excel import DATA as BASE

DATA = copy.deepcopy(BASE)
DATA["summary"] = {"work": 1500, "non_work": 700, "total": 2291}
for r in DATA["nonwork"]["rows"]:
    r.update({"C": 1, "비작업": 7, "적용": 7,
              "cond": {"혹서기(일최고기온 33℃ 이상)": 0.5, "일강수량 5㎜ 이상": 1.8,
                       "일최대순간풍속 15m/s 이상": 1.1}})


class BuildTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.raw = rh.build_report_hwpx(DATA)
        cls.zip = zipfile.ZipFile(io.BytesIO(cls.raw))
        cls.sec = cls.zip.read("Contents/section0.xml").decode("utf-8")

    def test_package_layout(self):
        infos = self.zip.infolist()
        self.assertEqual(infos[0].filename, "mimetype")                 # 첫 항목, 압축 안 함
        self.assertEqual(infos[0].compress_type, zipfile.ZIP_STORED)
        self.assertEqual(self.zip.read("mimetype"), b"application/hwp+zip")
        for n in self.zip.namelist():
            if n.endswith((".xml", ".hpf", ".rdf")):
                ET.fromstring(self.zip.read(n))                           # 모든 XML이 올바름

    def test_placeholders_filled(self):
        self.assertNotIn("{{", self.sec)
        self.assertIn("테스트 사업", self.sec)
        self.assertIn("2,291일 ≒ 75개월", self.sec)                      # 2291/365×12 = 75.3
        self.assertIn("상수도공사의 준비기간을 적용", self.sec)
        self.assertIn("강원도 강릉", self.sec)
        self.assertIn("기계공사를 완료하는 것으로 계획", self.sec)        # 공기 미반영 분야 문장

    def test_generated_tables_split_across_pages(self):
        # 글자처럼 취급하면 한 쪽을 넘는 표가 잘린다 → 새로 만든 표는 본문 흐름 + 셀 단위 쪽 나눔
        gen = re.findall(r'<hp:tbl [^>]*rowCnt="(\d+)"[^>]*>.*?<hp:pos treatAsChar="(\d)"', self.sec, flags=re.S)
        self.assertTrue(any(t == "0" for _, t in gen))
        self.assertIn('pageBreak="CELL" repeatHeader="1"', self.sec)

    def test_no_sample_project_data(self):
        for name in ("의성", "영덕", "영해", "가평"):
            self.assertNotIn(name, self.sec)


class TableTest(unittest.TestCase):
    def test_grid_validation_catches_gaps(self):
        t = rh.Table([1000, 1000])
        t.add(0, 0, "a")
        with self.assertRaises(AssertionError):
            t.validate()

    def test_civil_table_spans(self):
        t = rh.civil_work_table(DATA["civil"])
        t.validate()
        # 토공 1항목 + 소계, 구조물공 2구간(각 1항목) + 소계, 머리행, 마지막 합계
        self.assertEqual(t.n_rows, 1 + 2 + 3 + 1)

    def test_nonwork_groups(self):
        t = rh.nonwork_table(DATA["nonwork"]["rows"])
        t.validate()
        texts = [c["t"] for c in t.cells]
        self.assertIn("1.8", texts)                                        # 강우 열
        self.assertIn("26년", texts)


if __name__ == "__main__":
    unittest.main()
