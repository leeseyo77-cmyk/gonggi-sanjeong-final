"""drawing_reader — 일반도 PDF 쪽 고르기·Claude 요청 형식·면적 계산·내역서 대조(API 호출 없음).

Claude 호출은 가짜 클라이언트로 요청 형식만 확인한다(비용 없음). 풍각 일반도 샘플이 있으면
실제 PDF에서 후보 쪽을 고르는지도 본다.
실행: python -m unittest discover -s tests -v
"""
import json
import unittest
from types import SimpleNamespace

import pymupdf

from _support import SAMPLES, have

import crew_space as cs
import drawing_reader as dr

PUNGGAK_DWG = SAMPLES / "도면" / "풍각(토목, 건축, 기계, 조경).pdf"


def _pdf(pages):
    doc = pymupdf.open()
    for t in pages:
        pg = doc.new_page()
        pg.insert_text((50, 72), t, fontname="korea")
    return doc.tobytes()


class PagesTest(unittest.TestCase):
    def test_find_plan_pages(self):
        data = _pdf(["도면 목차", "처리시설 일반도 (01) 상부슬래브 평면도", "처리시설 구조물 배근도 (03) 평면도",
                     "처리시설 일반도 (03) A-A 단면도"])
        self.assertEqual([p["page"] for p in dr.find_plan_pages(data)], [2, 4])
        self.assertEqual(dr.page_count(data), 4)
        out = dr.render_pages(data, [2])
        self.assertEqual(out[0][0], 2)
        self.assertTrue(out[0][1].startswith(b"\x89PNG"))
        self.assertIn("일반도", out[0][2])

    @unittest.skipUnless(have(PUNGGAK_DWG), "풍각 일반도 샘플 없음")
    def test_punggak_candidates(self):
        pages = [p["page"] for p in dr.find_plan_pages(PUNGGAK_DWG.read_bytes())]
        # 24~32쪽이 처리시설 일반도(평면도·단면도)
        for p in (24, 25, 26):
            self.assertIn(p, pages)
        self.assertLessEqual(len(pages), 20)


class _FakeStream:
    def __init__(self, msg):
        self.msg = msg

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def get_final_message(self):
        return self.msg


class _FakeClient:
    def __init__(self, payload, stop_reason="end_turn"):
        self.kwargs = None
        msg = SimpleNamespace(stop_reason=stop_reason, model="claude-opus-5-5",
                              content=[SimpleNamespace(type="text", text=json.dumps(payload))],
                              usage=SimpleNamespace(input_tokens=1000, output_tokens=200))
        outer = self

        class _M:
            def stream(self, **kw):
                outer.kwargs = kw
                return _FakeStream(msg)

        self.beta = SimpleNamespace(messages=_M())


STRUCT = {"name": "구조물공사", "slab_area_m2": 458.9, "top_slab_area_m2": 325.2,
          "exterior_wall_length_m": 94.8, "interior_wall_length_m": 100.0, "wall_height_m": 4.9,
          "exterior_wall_thickness_m": 0.6, "interior_wall_thickness_m": 0.3,
          "base_slab_thickness_m": 0.8, "top_slab_thickness_m": 0.3, "pages": [24, 26],
          "calc": "29.3×11.1 + 19.1×7.0", "notes": ""}


class ClaudeRequestTest(unittest.TestCase):
    def test_request_shape(self):
        fake = _FakeClient({"structures": [STRUCT]})
        res = dr.read_drawings("k", [(24, b"\x89PNGfake", "29,300 11,100")], ["구조물공사"], client=fake)
        kw = fake.kwargs
        self.assertEqual(kw["model"], "claude-opus-5-5")
        self.assertEqual(kw["output_config"]["format"]["type"], "json_schema")
        self.assertEqual(kw["fallbacks"], "default")
        self.assertIn("server-side-fallback-2026-07-01", kw["betas"])
        blocks = kw["messages"][0]["content"]
        self.assertEqual([b["type"] for b in blocks], ["text", "image", "text"])
        self.assertIn("구조물공사", blocks[-1]["text"])
        self.assertEqual(res["structures"][0]["name"], "구조물공사")
        self.assertEqual(res["usage"], {"input": 1000, "output": 200})

    def test_refusal_raises(self):
        fake = _FakeClient({"structures": []}, stop_reason="refusal")
        with self.assertRaises(ValueError):
            dr.read_drawings("k", [(1, b"x", "")], [], client=fake)


class AreaCheckTest(unittest.TestCase):
    def test_areas_and_check(self):
        est = dr.areas_from_dims(STRUCT)
        self.assertEqual(est["slab"], 458.9)
        self.assertAlmostEqual(est["wall"], (94.8 + 100) * 4.9, places=1)
        # 풍각 내역서: 버림 45.4㎥, 철근콘크리트 927㎥, 거푸집 2,358㎡
        boq = {"버림": 45.37, "철근콘크리트": 927.46, "거푸집": 2358.0}
        self.assertLess(abs(dr.diff_pct(est["버림"], boq["버림"])), 2)
        self.assertEqual(dr.check(est, boq), [])                  # 30% 안이면 통과
        self.assertTrue(dr.check(dict(est, 철근콘크리트=2000), boq))   # 크게 다르면 표시


class BoqEstimateTest(unittest.TestCase):
    def test_estimate_from_boq(self):
        # 풍각 구조물공사: 버림 45.37㎥ → 바닥 454㎡(일반도 459㎡), 거푸집 2,258㎡ → 벽 한 면 902㎡
        items = [{"name": "무근콘크리트 펌프차 타설", "qty": 45.37, "unit": "㎥"},
                 {"name": "철근콘크리트 펌프차 타설", "qty": 927.46, "unit": "㎥"},
                 {"name": "유로폼(6회 사용)", "qty": 1933.09, "unit": "㎡"},
                 {"name": "합판거푸집", "qty": 324.49, "unit": "㎡"},
                 {"name": "무근콘크리트 깨기", "qty": 86, "unit": "㎥"}]
        slab, wall, basis = cs.estimate_areas(items)
        self.assertEqual(slab, 453.7)
        self.assertEqual(wall, 901.9)
        self.assertIn("버림", basis)
        q = cs.boq_quantities(items)
        self.assertEqual(q["철근콘크리트"], 927.46)
        self.assertEqual(cs.estimate_areas([])[:2], (0.0, 0.0))


class ManualModeTest(unittest.TestCase):
    """수동 모드: 발췌 PDF + 지시문 → claude.ai 답(JSON)을 붙여 넣으면 API 모드와 같은 결과."""

    def test_excerpt_pdf_and_prompt(self):
        data = _pdf(["표지", "처리시설 일반도 (01) 평면도", "처리시설 일반도 (03) 단면도"])
        ex = dr.extract_pages_pdf(data, [3, 2])
        self.assertEqual(dr.page_count(ex), 2)
        prompt = dr.manual_prompt(["구조물공사"], [24, 26])
        self.assertIn("24, 26쪽", prompt)
        self.assertIn("구조물공사", prompt)
        self.assertIn('"slab_area_m2"', prompt)

    def test_parse_answer_with_code_fence_and_chatter(self):
        ans = ("도면을 읽었습니다.\n```json\n" + json.dumps({"structures": [STRUCT]}, ensure_ascii=False)
               + "\n```\n참고하세요.")
        res = dr.parse_manual(ans)
        self.assertEqual(res["model"], "수동(claude.ai)")
        self.assertEqual(res["structures"][0]["slab_area_m2"], 458.9)
        self.assertEqual(dr.areas_from_dims(res["structures"][0])["slab"], 458.9)

    def test_parse_coerces_numbers_and_reports_missing(self):
        s = dict(STRUCT, slab_area_m2="1,234.5")
        self.assertEqual(dr.parse_manual(json.dumps({"structures": [s]}))["structures"][0]["slab_area_m2"], 1234.5)
        bad = {k: v for k, v in STRUCT.items() if k != "wall_height_m"}
        with self.assertRaises(ValueError) as cm:
            dr.parse_manual(json.dumps({"structures": [bad]}))
        self.assertIn("wall_height_m", str(cm.exception))
        with self.assertRaises(ValueError):
            dr.parse_manual("JSON이 아닌 답")


if __name__ == "__main__":
    unittest.main()
