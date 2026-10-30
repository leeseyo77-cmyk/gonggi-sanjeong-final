"""session_store — 중간저장(자동 저장·저장 파일). 단위 테스트 + 앱에서 연결이 끊긴 뒤 이어서 작업.

실행: python -m unittest discover -s tests -v
"""
import datetime as dt
import os
import shutil
import unittest

from _support import GWANGJU, OUT, app_test, exceptions, have

import session_store as ss

AUTO = OUT / "autosave_test"


class EncodeTest(unittest.TestCase):
    def test_round_trip(self):
        src = {"manual_rates": {"유로폼|복잡": {"daily": 30.0, "unit": "㎡"}},
               "excluded_items": {"흙운반|L=2km"}, "start_date": dt.date(2026, 3, 1),
               "weather_result": {"start_date": dt.datetime(2026, 3, 1, 9, 0), "rows": [(1, 2)]},
               "crew_by_item": {"터파기|토사": 3}}
        back = {k: ss.decode(v) for k, v in ss.collect(src).items()}
        self.assertEqual(back, src)

    def test_never_saves_api_key_or_buttons(self):
        src = {"anthropic_api_key": "sk-ant-x", "anthropic_api_key_crew": "sk", "work_save_make": True,
               "crew_ai_run": False, "crew_editor_1": {}, "disc_use_건축": True, "crew_90_1_0": 2}
        self.assertEqual(set(ss.collect(src)), {"disc_use_건축", "crew_90_1_0"})


class FileTest(unittest.TestCase):
    def test_zip_round_trip(self):
        state = ss.collect({"crew_by_item": {"a|b": 2}})
        files = {"civil": [("내역서.xlsx", b"PK\x03\x04data")], "disc_건축": [("건축.xlsx", b"x"), ("건축2.xlsx", b"y")]}
        st2, f2, man = ss.read_zip(ss.build_zip(state, files))
        self.assertEqual(st2, state)
        self.assertEqual(f2, files)
        self.assertEqual(man["version"], ss.VERSION)
        with self.assertRaises(ValueError):
            ss.read_zip(b"not a zip")

    def test_autosave_and_bad_ids(self):
        os.environ["GONGGI_AUTOSAVE_DIR"] = str(AUTO)
        try:
            wid = ss.new_id()
            self.assertTrue(ss.autosave(wid, {"x": 1}, {"civil": [("a.xlsx", b"123")]}))
            got = ss.load_autosave(wid)
            self.assertEqual(got[0], {"x": 1})
            self.assertEqual(got[1], {"civil": [("a.xlsx", b"123")]})
            self.assertIsNone(ss.load_autosave("../../etc"))        # 주소창 값으로 경로를 만들지 않는다
            self.assertFalse(ss.autosave("../x", {}, {}))
        finally:
            os.environ.pop("GONGGI_AUTOSAVE_DIR", None)
            shutil.rmtree(AUTO, ignore_errors=True)


@unittest.skipIf(os.environ.get("GONGGI_SKIP_APP") == "1" or not have(GWANGJU["토목"]), "광주 샘플 없음")
class ResumeAfterDisconnectTest(unittest.TestCase):
    def test_resume_with_same_work_code(self):
        os.environ["GONGGI_AUTOSAVE_DIR"] = str(AUTO)
        try:
            at = app_test(civil=GWANGJU["토목"])
            at.session_state["crew_by_item"] = {"유로폼|0-7m:복잡": 4}      # 사용자가 바꾼 조수
            at.run()
            wid = at.session_state["_work_id"]
            days = at.session_state["total_work_days"]
            self.assertTrue((AUTO / wid / "state.json").exists())

            # 연결이 끊겨 새 세션: 업로드 칸은 비어 있고 주소창의 작업 코드만 같다
            at2 = app_test(civil=None, query={"w": wid})
            self.assertEqual(exceptions(at2), [])
            self.assertEqual(at2.session_state["crew_by_item"], {"유로폼|0-7m:복잡": 4})
            self.assertEqual(at2.session_state["total_work_days"], days)
            self.assertTrue(any("저장된 작업의 내역서" in c.value for c in at2.caption))
        finally:
            os.environ.pop("GONGGI_AUTOSAVE_DIR", None)
            shutil.rmtree(AUTO, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
