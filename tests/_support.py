"""테스트 공용: 경로, 회귀 기준 샘플, AppTest 실행기.

샘플 내역서는 저장소에 올리지 않는다(.gitignore). 기본 위치는 저장소의 '샘플파일/'이고,
다른 곳에 두었다면 환경변수 GONGGI_SAMPLES로 알려 준다. 샘플이 없으면 앱 테스트는 건너뛴다.
"""
import os
import pathlib
import sys

REPO = pathlib.Path(__file__).resolve().parent.parent
SAMPLES = pathlib.Path(os.environ.get("GONGGI_SAMPLES", REPO / "샘플파일"))
OUT = REPO / "tests" / "_out"

if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))

# 광주 제1정수장 세트 — 앱 회귀 기준값을 잡은 샘플
GWANGJU = {
    "토목": SAMPLES / "제1정수장 고도정수처리시설 설치사업240415.xlsx",
    "기계": SAMPLES / "기계" / "1. 제1정수장 고도정수처리시설 설치사업_기계_240408.xlsx",
    "건축": SAMPLES / "건축 및 건축기계" / "1. 건축 내역서_광주 제1정수장 고도정수처리시설.xlsx",
    "전기": SAMPLES / "전기 및 계측제어" / "광주 제1정수장 고도정수처리시설 설치사업(전기 및 계측제어).xlsx",
}

_DISC_KEYS = ("기계", "건축", "건축기계설비", "전기", "조경")


def have(*paths) -> bool:
    return all(pathlib.Path(p).exists() for p in paths)


def app_test(civil=None, discs=None, timeout=1800):
    """샘플 파일을 업로드한 상태로 app.py를 한 번 실행한 AppTest를 돌려준다.

    file_uploader는 tests/app_wrapper.py가 바꿔치기한다. 다운로드 버튼에 실린 파일은
    tests/_out/에 저장된다(예정공정표 엑셀 확인용).
    """
    from streamlit.testing.v1 import AppTest

    os.environ["GONGGI_APP"] = str(REPO / "app.py")
    os.environ["GONGGI_SAMPLE"] = str(civil) if civil else ""
    for k in _DISC_KEYS:
        os.environ.pop("GONGGI_DISC_" + k, None)
    for k, v in (discs or {}).items():
        os.environ["GONGGI_DISC_" + k] = str(v)
    OUT.mkdir(exist_ok=True)
    os.environ["GONGGI_DL_DIR"] = str(OUT)
    at = AppTest.from_file(str(REPO / "tests" / "app_wrapper.py"), default_timeout=timeout)
    at.run()
    return at


def exceptions(at):
    return [str(e.value)[:300] for e in at.exception]


def metric(at, label):
    """라벨이 정확히 일치하는 metric의 값(없으면 None)."""
    for m in at.metric:
        if m.label == label:
            return m.value
    return None


def button(at, prefix):
    for b in at.button:
        if (b.label or "").startswith(prefix):
            return b
    raise AssertionError(f"버튼 없음: {prefix}")


def widget(at, kind, key):
    for w in getattr(at, kind):
        if w.key == key:
            return w
    raise AssertionError(f"{kind} 없음: {key}")
