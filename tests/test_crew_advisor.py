"""crew_advisor — Claude 투입조수 추천(지시문·답 해석·상한 적용, API 호출 없음).

실행: python -m unittest discover -s tests -v
"""
import json
import unittest

import _support  # noqa: F401  (저장소 경로 등록)
import crew_advisor as ca
from test_drawing_reader import _FakeClient

ITEMS = [{"key": "유로폼|0-7m:복잡", "대공종": "구조물공사", "세부공종": "유로폼", "규격": "0-7m:복잡",
          "구간 수": 1, "1조 구성": "품셈 작업조 5인", "1조 기준 일수": 193, "현재 조수": 1,
          "역산 필요 조수": 3, "상한(cap)": 4, "상한 근거": "작업면적 6,500㎡ ÷ (1조 5인 × 30.2㎡)"}]
PROJECT = {"target": 600, "months": 36, "combine_sum": False, "now": 1084, "rec": 590, "capacity": 20000}


class ContextTest(unittest.TestCase):
    def test_context_and_prompt(self):
        ctx = ca.build_context(ITEMS, PROJECT)
        self.assertIn("목표 순작업일수: 600일", ctx)
        self.assertIn("유로폼|0-7m:복잡", ctx)
        self.assertIn("처리용량 20,000", ctx)
        p = ca.manual_prompt(ctx)
        self.assertIn('"recommendations"', p)
        self.assertIn("상한(cap)", p)


class ParseApplyTest(unittest.TestCase):
    def test_parse_manual(self):
        ans = "```json\n" + json.dumps({"recommendations": [{"key": "유로폼|0-7m:복잡", "crew": "3조",
                                                             "reason": "193일 ÷ 3 = 65일"}],
                                        "overall": "가능", "risks": []}, ensure_ascii=False) + "\n```"
        res = ca.parse_manual(ans)
        self.assertEqual(res["recommendations"][0]["crew"], 3)
        self.assertEqual(res["model"], "수동(claude.ai)")
        with self.assertRaises(ValueError):
            ca.parse_manual('{"overall": "x"}')

    def test_apply_clamps_to_cap_and_ignores_unknown(self):
        base = {"a": 1, "b": 2}
        crew, notes, unknown = ca.apply_recommendations(
            [{"key": "a", "crew": 9, "reason": ""}, {"key": "b", "crew": 0, "reason": ""},
             {"key": "zz", "crew": 2, "reason": ""}], base, {"a": 4})
        self.assertEqual(crew, {"a": 4, "b": 1})
        self.assertIn("상한 4조", notes["a"])
        self.assertEqual(unknown, ["zz"])


class ApiShapeTest(unittest.TestCase):
    def test_request_shape(self):
        fake = _FakeClient({"recommendations": [{"key": "유로폼|0-7m:복잡", "crew": 3, "reason": "r"}],
                            "overall": "o", "risks": []})
        res = ca.recommend_api("k", ca.build_context(ITEMS, PROJECT), client=fake)
        kw = fake.kwargs
        self.assertEqual(kw["system"], ca.SYSTEM)
        self.assertEqual(kw["output_config"]["format"]["schema"], ca.SCHEMA)
        self.assertEqual(kw["fallbacks"], "default")
        self.assertEqual(res["recommendations"][0]["crew"], 3)


if __name__ == "__main__":
    unittest.main()
