"""AppTest용 래퍼: file_uploader를 샘플 파일로 바꿔치기한 뒤 app.py를 실행한다.

환경변수(tests/_support.app_test가 설정):
  GONGGI_APP            실행할 app.py
  GONGGI_SAMPLE         토목 내역서(기본 업로더)
  GONGGI_DISC_<분야>    분야별 업로더(key=disc_up_<분야>)에 넣을 파일, '||'로 여러 개
  GONGGI_DL_DIR         다운로드 버튼에 실린 파일을 저장할 폴더
"""
import io
import os
import pathlib
import sys

import streamlit as st

REPO = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
os.chdir(REPO)

SAMPLE = os.environ.get("GONGGI_SAMPLE", "")
APP = os.environ.get("GONGGI_APP", str(REPO / "app.py"))


class _Up(io.BytesIO):
    def __init__(self, data, name):
        super().__init__(data)
        self.name = name


_data = pathlib.Path(SAMPLE).read_bytes() if SAMPLE else None
_name = pathlib.Path(SAMPLE).name if SAMPLE else ""


def _uploader(*a, **k):
    """key가 disc_up_<분야>면 GONGGI_DISC_<분야> 파일을, 아니면 토목 샘플을 돌려준다."""
    key = str(k.get("key") or "")
    if key.startswith("disc_up_"):
        p = os.environ.get("GONGGI_DISC_" + key[len("disc_up_"):], "")
        if p:
            return [_Up(pathlib.Path(x).read_bytes(), os.path.basename(x))
                    for x in p.split("||") if x]
        return []
    return _Up(_data, _name) if _data else None


st.file_uploader = _uploader

_DL_DIR = os.environ.get("GONGGI_DL_DIR", "")
if _DL_DIR:
    _orig_dl = st.download_button

    def _dl(label, data=None, file_name=None, *a, **k):
        """다운로드 버튼에 실린 파일을 검증용으로 저장한다."""
        try:
            raw = data.getvalue() if hasattr(data, "getvalue") else data
            if isinstance(raw, (bytes, bytearray)) and file_name:
                pathlib.Path(_DL_DIR, file_name).write_bytes(raw)
        except Exception:
            pass
        return _orig_dl(label, data, file_name, *a, **k)

    st.download_button = _dl

_src = pathlib.Path(APP).read_text(encoding="utf-8")
exec(compile(_src, "app.py", "exec"), {"__name__": "__main__"})
