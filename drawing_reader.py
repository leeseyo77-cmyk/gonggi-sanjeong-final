# -*- coding: utf-8 -*-
"""
drawing_reader.py — 일반도 PDF에서 구조물 면적 읽기(Claude, 사용자 API 키)

면적 기준 동시 투입 상한(crew_space)에 필요한 구조물별 바닥슬라브·벽체 면적을 도면에서 얻는다.
  1) PDF에서 '일반도'이면서 평면도·단면도가 든 쪽을 고른다(배근도 제외).
  2) 그 쪽을 그림(PNG)과 추출 글자(치수 숫자)로 Claude에 넘겨 구조물별 치수를 정해진 형식(JSON)으로 받는다.
  3) 앱이 면적을 계산하고 내역서 물량(버림·철근콘크리트·거푸집)과 대조해 잘못 읽은 값을 걸러낸다.
풍각 처리시설 일반도로 손 계산해 본 결과: 버림 +1%, 콘크리트 +5%, 거푸집 +21%(칸막이벽 어림) — 대조로
충분히 걸러낼 수 있는 수준.

API 키는 화면에서 받아 이 요청에만 쓰고 저장하지 않는다. 비용은 키 주인에게 청구된다.
"""

from __future__ import annotations

import base64
import json
import re
from typing import Any, Dict, List, Optional, Sequence

from claude_call import MODEL, MODELS, call_structured, extract_json
ZOOM = 1.3          # A3 도면 → 약 1,550×1,100px. API가 긴 변 1,568px 넘는 그림을 줄이므로 그 안에 맞춘다
MAX_PAGES = 12

SYSTEM = (
    "당신은 상하수도 처리시설 토목 구조물 도면을 읽는 기술자입니다. 주어진 일반도(평면도·단면도)의 치수만 "
    "근거로 구조물별 치수를 읽어 정해진 형식으로 답합니다.\n"
    "- 도면 치수는 mm입니다. 답은 m 단위로 바꿔 적습니다(29,300 → 29.3).\n"
    "- slab_area_m2: 구조물 바닥(하부슬래브) 외곽 면적. 평면이 여러 직사각형이면 나눠 곱해 더합니다.\n"
    "- top_slab_area_m2: 상부슬래브가 덮는 면적(개방 수조는 빼고), 없으면 0.\n"
    "- exterior_wall_length_m: 외벽 중심선 길이 합. interior_wall_length_m: 내부 칸막이벽 길이 합"
    "(일부 구간만 있는 벽은 그 길이만).\n"
    "- wall_height_m: 벽 높이(하부슬래브 윗면~상부슬래브 아랫면, 구간마다 다르면 길이 가중 평균).\n"
    "- 두께는 단면도·평면도 치수에서 읽습니다. 읽을 수 없는 값은 0으로 두고 notes에 이유를 적습니다.\n"
    "- name은 사용자가 준 내역서 구조물명 중 맞는 것을 그대로 쓰고, 없으면 도면의 구조물명을 씁니다.\n"
    "- calc에는 면적·길이를 어떤 치수로 계산했는지 식을 짧게 적습니다(검토자가 도면과 대조할 수 있게)."
)

_NUM = {"type": "number"}
SCHEMA = {
    "type": "object",
    "properties": {
        "structures": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "name": {"type": "string"},
                    "slab_area_m2": _NUM,
                    "top_slab_area_m2": _NUM,
                    "exterior_wall_length_m": _NUM,
                    "interior_wall_length_m": _NUM,
                    "wall_height_m": _NUM,
                    "exterior_wall_thickness_m": _NUM,
                    "interior_wall_thickness_m": _NUM,
                    "base_slab_thickness_m": _NUM,
                    "top_slab_thickness_m": _NUM,
                    "pages": {"type": "array", "items": {"type": "integer"}},
                    "calc": {"type": "string"},
                    "notes": {"type": "string"},
                },
                "required": ["name", "slab_area_m2", "top_slab_area_m2", "exterior_wall_length_m",
                             "interior_wall_length_m", "wall_height_m", "exterior_wall_thickness_m",
                             "interior_wall_thickness_m", "base_slab_thickness_m", "top_slab_thickness_m",
                             "pages", "calc", "notes"],
                "additionalProperties": False,
            },
        },
    },
    "required": ["structures"],
    "additionalProperties": False,
}


# ── PDF ──────────────────────────────────────────────────────────────────
def find_plan_pages(pdf_bytes: bytes) -> List[Dict[str, Any]]:
    """일반도 중 평면도·단면도가 든 쪽 [{"page": 1부터, "title": 도면명}]. 배근도·상세도는 뺀다."""
    import pymupdf
    out = []
    with pymupdf.open(stream=pdf_bytes, filetype="pdf") as doc:
        for i, pg in enumerate(doc):
            t = pg.get_text()
            flat = re.sub(r"\s+", "", t)
            if "일반도" not in flat or "배근도" in flat:
                continue
            if not any(k in flat for k in ("평면도", "단면도")):
                continue
            m = re.search(r"([가-힣·]+일반도\(?\d*\)?)", flat)
            out.append({"page": i + 1, "title": m.group(1) if m else "일반도"})
    return out


def page_count(pdf_bytes: bytes) -> int:
    import pymupdf
    with pymupdf.open(stream=pdf_bytes, filetype="pdf") as doc:
        return doc.page_count


def render_pages(pdf_bytes: bytes, pages: Sequence[int], zoom: float = ZOOM):
    """[(쪽, PNG 바이트, 추출 글자)]."""
    import pymupdf
    out = []
    with pymupdf.open(stream=pdf_bytes, filetype="pdf") as doc:
        for p in pages:
            pg = doc[p - 1]
            png = pg.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom)).tobytes("png")
            out.append((p, png, " ".join(pg.get_text().split())[:4000]))
    return out


# ── Claude ───────────────────────────────────────────────────────────────
def build_messages(rendered, struct_names: Sequence[str]) -> List[Dict[str, Any]]:
    content: List[Dict[str, Any]] = []
    for p, png, txt in rendered:
        content.append({"type": "text", "text": f"[{p}쪽 도면 — 추출 글자(치수 확인용): {txt}]"})
        content.append({"type": "image", "source": {"type": "base64", "media_type": "image/png",
                                                    "data": base64.standard_b64encode(png).decode("ascii")}})
    names = ", ".join(struct_names) if struct_names else "(내역서 구조물명 없음)"
    content.append({"type": "text", "text": (
        f"내역서의 구조물명: {names}\n"
        "위 도면에서 구조물별 치수를 읽어 형식에 맞게 답하세요. 도면에 없는 구조물은 넣지 마세요.")})
    return [{"role": "user", "content": content}]


def read_drawings(api_key: str, rendered, struct_names: Sequence[str], model: str = MODEL,
                  client=None) -> Dict[str, Any]:
    """도면 그림을 Claude에 보내 구조물별 치수를 받는다.

    반환: {"structures": [...], "model": 응답 모델, "usage": {...}} — 거절·형식 오류면 ValueError.
    client는 테스트용(가짜 클라이언트 주입).
    """
    return call_structured(api_key, SYSTEM, build_messages(rendered, struct_names)[0]["content"], SCHEMA,
                           model=model, client=client)


# ── 면적 계산·내역서 대조 ─────────────────────────────────────────────────
def areas_from_dims(s: Dict[str, Any]) -> Dict[str, float]:
    """치수 → 바닥·벽(한 면) 면적과 검산용 물량(버림·철근콘크리트·거푸집 추정)."""
    ext, inn, h = s["exterior_wall_length_m"], s["interior_wall_length_m"], s["wall_height_m"]
    wall = (ext + inn) * h
    slab, top = s["slab_area_m2"], s["top_slab_area_m2"]
    conc = (slab * s["base_slab_thickness_m"] + top * s["top_slab_thickness_m"]
            + ext * h * s["exterior_wall_thickness_m"] + inn * h * s["interior_wall_thickness_m"])
    return {"slab": round(slab, 1), "wall": round(wall, 1), "버림": round(slab * 0.1, 1),
            "철근콘크리트": round(conc, 1), "거푸집": round(2 * wall + top, 1)}


def diff_pct(est: float, boq: float) -> Optional[float]:
    return round((est - boq) / boq * 100, 1) if boq else None


def check(est: Dict[str, float], boq: Dict[str, float], tol: float = 30.0) -> List[str]:
    """도면 추정 물량과 내역서 물량 차이가 tol%를 넘는 항목 목록(잘못 읽었을 가능성)."""
    bad = []
    for k in ("버림", "철근콘크리트", "거푸집"):
        d = diff_pct(est.get(k, 0), boq.get(k, 0))
        if d is not None and abs(d) > tol:
            bad.append(f"{k} {d:+.0f}%")
    return bad


# ── 수동 모드(claude.ai에 붙여넣기 — API 키 없이) ──────────────────────────
def extract_pages_pdf(pdf_bytes: bytes, pages: Sequence[int]) -> bytes:
    """고른 쪽만 담은 PDF — claude.ai에 올리기 좋게 작게 만든다."""
    import pymupdf
    with pymupdf.open(stream=pdf_bytes, filetype="pdf") as src, pymupdf.open() as out:
        for p in sorted(pages):
            out.insert_pdf(src, from_page=p - 1, to_page=p - 1)
        return out.tobytes(garbage=3, deflate=True)


def manual_prompt(struct_names: Sequence[str], pages: Sequence[int]) -> str:
    """claude.ai 채팅에 PDF와 함께 붙여 넣을 지시문. 답은 API 모드와 같은 JSON 형식."""
    example = {"structures": [{
        "name": "구조물명", "slab_area_m2": 0, "top_slab_area_m2": 0, "exterior_wall_length_m": 0,
        "interior_wall_length_m": 0, "wall_height_m": 0, "exterior_wall_thickness_m": 0,
        "interior_wall_thickness_m": 0, "base_slab_thickness_m": 0, "top_slab_thickness_m": 0,
        "pages": [1], "calc": "면적·길이 계산식", "notes": "읽지 못한 값과 이유"}]}
    names = ", ".join(struct_names) if struct_names else "(없음)"
    return (
        f"{SYSTEM}\n\n"
        f"첨부한 PDF는 원본 도면의 {', '.join(str(p) for p in sorted(pages))}쪽을 발췌한 것입니다"
        "(pages에는 첨부 PDF의 쪽 번호를 적으세요).\n"
        f"내역서의 구조물명: {names}\n\n"
        "도면에서 구조물별 치수를 읽어 아래 JSON 형식으로만 답하세요. 설명 문장 없이 JSON 하나만, "
        "모든 키를 빠짐없이, 숫자는 숫자로 적습니다.\n\n"
        + json.dumps(example, ensure_ascii=False, indent=2)
    )


_REQ = SCHEMA["properties"]["structures"]["items"]["required"]


def parse_manual(text: str) -> Dict[str, Any]:
    """claude.ai 답을 API 모드 결과와 같은 형태로. 코드블록·앞뒤 설명이 섞여 있어도 JSON만 뽑는다."""
    data = extract_json(text)
    if isinstance(data, list):
        data = {"structures": data}
    sts = data.get("structures") if isinstance(data, dict) else None
    if not isinstance(sts, list) or not sts:
        raise ValueError("'structures' 목록이 없습니다. 지시문의 형식대로 다시 요청하세요.")
    out = []
    for i, s in enumerate(sts, 1):
        miss = [k for k in _REQ if k not in s]
        if miss:
            raise ValueError(f"{i}번째 구조물에 빠진 항목: {', '.join(miss)}")
        row = dict(s)
        for k in _REQ:
            if k.endswith(("_m", "_m2")):
                try:
                    row[k] = float(str(s[k]).replace(",", "")) if s[k] not in (None, "") else 0.0
                except ValueError as e:
                    raise ValueError(f"{i}번째 구조물 '{k}' 값이 숫자가 아닙니다: {s[k]}") from e
        row["name"] = str(s["name"]).strip()
        row["pages"] = [int(p) for p in (s.get("pages") or []) if str(p).strip().isdigit()]
        out.append(row)
    return {"structures": out, "model": "수동(claude.ai)", "usage": {}}
