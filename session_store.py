# -*- coding: utf-8 -*-
"""
session_store.py — 작업 중간저장(자동 저장·저장 파일)

Streamlit은 연결이 끊기거나 브라우저를 새로 고치면 새 세션이 열려 입력한 값이 모두 사라진다.
  1) 자동 저장: 매 실행이 끝날 때 입력값과 올린 파일을 서버 임시 폴더에 저장하고, 주소창의
     ?w=<작업 코드>로 다시 들어오면(새로 고침·재연결) 그대로 복원한다. 앱이 재배포·재시작되면
     임시 폴더가 지워지므로 오래 보관할 작업은 저장 파일로 받아 둔다.
  2) 저장 파일(.zip): 입력값(state.json)과 올린 내역서를 묶어 내려받고, 나중에 올려 이어서 작업한다.
API 키는 저장하지 않는다. 저장 파일은 pickle이 아닌 JSON이라 올린 파일로 코드가 실행될 위험이 없다.
"""

from __future__ import annotations

import datetime as _dt
import hashlib
import io
import json
import os
import re
import shutil
import tempfile
import time
import uuid
import zipfile
from typing import Any, Dict, List, Optional, Tuple

VERSION = 1
AUTOSAVE_DAYS = 7

# 입력 위젯(키로 복원). API 키·버튼·업로더·편집표 상태는 넣지 않는다.
WIDGET_KEYS = {
    "area_caps_on", "exclude_haul_widget", "split_by_parent", "weather_start_date",
    "parallel_subgroup_widget", "sched_flood_months", "selected_major_widget", "crew_cap_input",
    "excavator_radius", "person_area", "plant_capacity", "proj_target_months", "site_area",
    "site_crew_limit", "target_months_input", "weather_work_days", "combine_mode_widget", "crew_ai_mode",
    "dwg_mode", "crew_ai_model", "dwg_model", "region_sido", "region_station", "report_prep_label",
    "report_title", "sched_project_name", "crew_show_all", "crew_ai_paste", "dwg_paste",
    "weather_region_legacy", "weather_holidays", "weather_min_rest", "commission_months",
    "weather_rain", "weather_cold", "weather_hot",
}
WIDGET_PATTERNS = [re.compile(p) for p in (
    r"^disc_(par|use|tgt|how|tc)_", r"^wcond_", r"^crew_cap_(?!input$)", r"^crew_\d",
    # 지구·공종별 투입조수(crew_<지구>_<공종>) — 버튼·편집표 키와 겹치지 않게 아래 NOT_CREW로 거른다
    r"^crew_(?!(ai_|by_item$|edit_|show_all$|targets$|editor_|reset_|cap_input$))",
)]
# 위젯이 아닌 입력·결과(앱이 파일에서 다시 계산하지 못하는 것)
DATA_KEYS = {
    "manual_rates", "crew_by_item", "excluded_items", "struct_areas", "crew_ai_apply", "crew_ai_result",
    "dwg_result", "combine_mode", "exclude_haul", "selected_major", "parallel_subgroup_cats",
    "weather_result", "weather_conditions", "selected_region", "selected_station", "start_date",
    "_proj_applied_days", "_loaded_file_sig", "_synced_work_days", "work_days_input",
}
NEVER = re.compile(r"api_key", re.IGNORECASE)


def wanted(key: str) -> bool:
    if NEVER.search(key):
        return False
    return key in WIDGET_KEYS or key in DATA_KEYS or any(p.search(key) for p in WIDGET_PATTERNS)


# ── JSON 변환(집합·날짜·튜플을 표시해 되살린다) ────────────────────────────
def encode(o: Any) -> Any:
    if isinstance(o, dict):
        if all(isinstance(k, str) for k in o):
            return {k: encode(v) for k, v in o.items()}
        return {"__dict__": [[encode(k), encode(v)] for k, v in o.items()]}
    if isinstance(o, (set, frozenset)):
        return {"__set__": [encode(x) for x in o]}
    if isinstance(o, tuple):
        return {"__tuple__": [encode(x) for x in o]}
    if isinstance(o, list):
        return [encode(x) for x in o]
    if isinstance(o, _dt.datetime):
        return {"__dt__": o.isoformat()}
    if isinstance(o, _dt.date):
        return {"__date__": o.isoformat()}
    if hasattr(o, "item") and callable(o.item):          # numpy 수
        return o.item()
    if o is None or isinstance(o, (str, int, float, bool)):
        return o
    raise TypeError(type(o).__name__)


def decode(o: Any) -> Any:
    if isinstance(o, list):
        return [decode(x) for x in o]
    if isinstance(o, dict):
        if set(o) == {"__set__"}:
            return {decode(x) for x in o["__set__"]}
        if set(o) == {"__tuple__"}:
            return tuple(decode(x) for x in o["__tuple__"])
        if set(o) == {"__dt__"}:
            return _dt.datetime.fromisoformat(o["__dt__"])
        if set(o) == {"__date__"}:
            return _dt.date.fromisoformat(o["__date__"])
        if set(o) == {"__dict__"}:
            return {decode(k): decode(v) for k, v in o["__dict__"]}
        return {k: decode(v) for k, v in o.items()}
    return o


def collect(session) -> Dict[str, Any]:
    """세션에서 저장할 값만 골라 JSON으로 바꿀 수 있는 형태로. 바꿀 수 없는 값은 건너뛴다."""
    out = {}
    for k in list(session.keys()):
        if not isinstance(k, str) or not wanted(k):
            continue
        try:
            out[k] = encode(session[k])
        except (TypeError, ValueError):
            continue
    return out


def apply(session, state: Dict[str, Any]) -> int:
    """저장된 값을 세션에 넣는다(위젯이 그려지기 전에 불러야 한다). 넣은 개수를 돌려준다."""
    n = 0
    for k, v in state.items():
        if not wanted(k):
            continue
        try:
            session[k] = decode(v)
            n += 1
        except Exception:      # noqa: BLE001 — 한 값이 실패해도 나머지는 복원
            continue
    return n


# ── 올린 파일 ────────────────────────────────────────────────────────────
class MemFile(io.BytesIO):
    """저장해 둔 파일을 업로드 파일처럼(이름·getvalue) 쓰게 한다."""

    def __init__(self, name: str, data: bytes):
        super().__init__(data)
        self.name = name


Files = Dict[str, List[Tuple[str, bytes]]]       # {"civil": [(이름, 바이트)], "disc_건축": [...]}


def _safe(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", name)[:120] or "file"


def build_zip(state: Dict[str, Any], files: Files) -> bytes:
    buf = io.BytesIO()
    manifest = {"version": VERSION, "saved_at": _dt.datetime.now().isoformat(timespec="seconds"),
                "files": {}}
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for slot, lst in files.items():
            manifest["files"][slot] = []
            for i, (nm, data) in enumerate(lst):
                path = f"files/{_safe(slot)}/{i}_{_safe(nm)}"
                z.writestr(path, data)
                manifest["files"][slot].append({"name": nm, "path": path})
        z.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=1))
        z.writestr("state.json", json.dumps(state, ensure_ascii=False))
    return buf.getvalue()


def read_zip(data: bytes) -> Tuple[Dict[str, Any], Files, Dict[str, Any]]:
    """저장 파일 → (상태, 파일, 머리정보). 형식이 아니면 ValueError."""
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
        manifest = json.loads(z.read("manifest.json").decode("utf-8"))
        state = json.loads(z.read("state.json").decode("utf-8"))
    except (zipfile.BadZipFile, KeyError, json.JSONDecodeError) as e:
        raise ValueError("공기산정 작업 저장 파일이 아닙니다.") from e
    files: Files = {}
    for slot, lst in (manifest.get("files") or {}).items():
        files[slot] = [(f["name"], z.read(f["path"])) for f in lst]
    return state, files, manifest


# ── 자동 저장(서버 임시 폴더, 작업 코드별) ─────────────────────────────────
def new_id() -> str:
    return uuid.uuid4().hex[:12]


def _root() -> str:
    return os.environ.get("GONGGI_AUTOSAVE_DIR") or os.path.join(tempfile.gettempdir(), "gonggi_autosave")


def _dir(work_id: str) -> Optional[str]:
    if not re.fullmatch(r"[0-9a-f]{8,32}", work_id or ""):
        return None                       # 주소창 값으로 임의 경로를 만들지 않게
    return os.path.join(_root(), work_id)


def autosave(work_id: str, state: Dict[str, Any], files: Files) -> bool:
    d = _dir(work_id)
    if not d:
        return False
    os.makedirs(d, exist_ok=True)
    meta = {"version": VERSION, "saved_at": time.time(), "files": {}}
    for slot, lst in files.items():
        meta["files"][slot] = []
        for nm, data in lst:
            h = hashlib.sha1(data).hexdigest()[:16]
            path = os.path.join(d, f"{h}.bin")
            if not os.path.exists(path):          # 같은 파일은 다시 쓰지 않는다
                with open(path, "wb") as f:
                    f.write(data)
            meta["files"][slot].append({"name": nm, "hash": h})
    tmp = os.path.join(d, "state.json.tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump({"meta": meta, "state": state}, f, ensure_ascii=False)
    os.replace(tmp, os.path.join(d, "state.json"))
    # 지금 쓰지 않는 파일은 지운다
    keep = {f"{x['hash']}.bin" for lst in meta["files"].values() for x in lst}
    for fn in os.listdir(d):
        if fn.endswith(".bin") and fn not in keep:
            try:
                os.remove(os.path.join(d, fn))
            except OSError:
                pass
    return True


def load_autosave(work_id: str):
    """(상태, 파일, 저장 시각) 또는 None."""
    d = _dir(work_id)
    p = os.path.join(d, "state.json") if d else None
    if not p or not os.path.exists(p):
        return None
    try:
        with open(p, encoding="utf-8") as f:
            raw = json.load(f)
        files: Files = {}
        for slot, lst in (raw["meta"].get("files") or {}).items():
            files[slot] = []
            for x in lst:
                with open(os.path.join(d, f"{x['hash']}.bin"), "rb") as f:
                    files[slot].append((x["name"], f.read()))
        return raw["state"], files, raw["meta"].get("saved_at")
    except (OSError, KeyError, json.JSONDecodeError):
        return None


def cleanup(days: int = AUTOSAVE_DAYS) -> None:
    root = _root()
    if not os.path.isdir(root):
        return
    cut = time.time() - days * 86400
    for wid in os.listdir(root):
        p = os.path.join(root, wid)
        try:
            if os.path.getmtime(p) < cut:
                shutil.rmtree(p, ignore_errors=True)
        except OSError:
            pass
