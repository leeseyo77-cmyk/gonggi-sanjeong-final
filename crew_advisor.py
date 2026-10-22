# -*- coding: utf-8 -*-
"""
crew_advisor.py — Claude 투입조수 추천(사례 없이 시작하는 1차 버전)

앱이 계산한 값(크리티컬 공종의 1조 기준 작업일수·1조 구성·동시 투입 상한·목표 순작업일수)을
Claude에 넘기면, Claude는 '공종별 조수 + 이유'만 제안한다. 작업일수와 공기는 앱이 같은
계산식으로 다시 구하고, 상한을 넘는 값은 상한으로 잘라 표시한다 — AI는 계산하지 않는다.

조수는 현장 경험이 정하는 값이다(앱은 그동안 역산만 했다). 과거 사례(감리 현장의 예정공정표·
공종별 투입인원)가 모이면 그 사례를 함께 넘겨 회사 기준에 가까운 추천으로 개선한다.
"""

from __future__ import annotations

import json
from typing import Any, Dict, List, Sequence

from claude_call import MODEL, call_structured, extract_json

MAX_ITEMS = 60      # 넘기는 공종 수 상한(크리티컬·일수 큰 순)

SYSTEM = (
    "당신은 상하수도 처리시설 토목공사의 공정관리 기술자입니다. 앱이 계산한 공종별 '1조 기준 작업일수'와 "
    "'동시 투입 상한'을 보고, 목표 순작업일수 안에 주공정이 끝나도록 현장에서 실제로 운영할 만한 투입조수를 "
    "제안합니다.\n"
    "- 조수(crew)는 1 이상의 정수이고, 공종마다 주어진 상한(cap)을 넘으면 안 됩니다.\n"
    "- 1조의 뜻은 공종마다 다릅니다('1조 구성' 참고: 품셈 작업조이거나 직종별 1인).\n"
    "- 같은 구간(라인) 안 공종은 순차, 구간끼리는 병행입니다(최장 모드). 합산 모드면 대공종 일수의 합이 공기입니다.\n"
    "- 계산상 최소 조수가 아니라 현장 운영을 고려하세요: 연속 작업이 끊기지 않게 같은 직종 공종은 조수를 비슷하게, "
    "장비가 정하는 공종(터파기 등)은 장비 대수, 작업면적이 좁은 구조물은 과밀을 피함, 조수를 바꾸는 공종 수는 "
    "필요한 만큼만.\n"
    "- 목표가 무리라고 판단되면 그 이유를 risks에 적으세요.\n"
    "- reason은 한두 문장으로, 근거 숫자(현재 일수·필요 일수·상한)를 들어 적습니다.\n"
    "- 주어진 공종 키(key)를 그대로 쓰고, 바꿀 필요가 없는 공종은 현재 조수를 그대로 적어도 됩니다."
)

SCHEMA = {
    "type": "object",
    "properties": {
        "recommendations": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {"key": {"type": "string"}, "crew": {"type": "integer"},
                               "reason": {"type": "string"}},
                "required": ["key", "crew", "reason"],
                "additionalProperties": False,
            },
        },
        "overall": {"type": "string"},
        "risks": {"type": "array", "items": {"type": "string"}},
    },
    "required": ["recommendations", "overall", "risks"],
    "additionalProperties": False,
}


def build_context(items: Sequence[Dict[str, Any]], project: Dict[str, Any]) -> str:
    """Claude에 넘길 표(글). items: 공종 dict 목록, project: 목표·모드 등."""
    lines = [
        f"- 목표 순작업일수: {project['target']}일 (발주처 목표 공기 {project['months']:g}개월에서 준비·시운전·정리·"
        f"비작업일수를 뺀 값)",
        f"- 공기 합산 방식: {'합산(순차)' if project.get('combine_sum') else '최장(병행)'}",
        f"- 현재 주공정: {project['now']}일 / 앱 역산안(최소 조수): {project['rec']}일",
    ]
    if project.get("capacity"):
        lines.append(f"- 시설 규모: 처리용량 {project['capacity']:,}㎥/일")
    head = ["key", "대공종", "세부공종", "규격", "구간 수", "1조 구성", "1조 기준 일수", "현재 조수",
            "역산 필요 조수", "상한(cap)", "상한 근거"]
    rows = [" | ".join(head)]
    for it in list(items)[:MAX_ITEMS]:
        rows.append(" | ".join(str(it.get(h, "")) for h in head))
    return ("[사업 조건]\n" + "\n".join(lines) + "\n\n[공종별 현황 — 크리티컬·일수 큰 순]\n" + "\n".join(rows))


def recommend_api(api_key: str, context: str, model: str = MODEL, client=None) -> Dict[str, Any]:
    content = [{"type": "text", "text": context + "\n\n위 조건으로 공종별 투입조수를 제안하세요."}]
    return call_structured(api_key, SYSTEM, content, SCHEMA, model=model, client=client)


def manual_prompt(context: str) -> str:
    example = {"recommendations": [{"key": "공종 key 그대로", "crew": 2, "reason": "근거"}],
               "overall": "전체 판단 요약", "risks": ["목표 달성 위험 요인"]}
    return (f"{SYSTEM}\n\n{context}\n\n위 조건으로 공종별 투입조수를 제안하세요. 설명 문장 없이 아래 JSON 하나만, "
            "모든 키를 빠짐없이 답하세요.\n\n" + json.dumps(example, ensure_ascii=False, indent=2))


def parse_manual(text: str) -> Dict[str, Any]:
    data = extract_json(text)
    if isinstance(data, list):
        data = {"recommendations": data, "overall": "", "risks": []}
    recs = data.get("recommendations") if isinstance(data, dict) else None
    if not isinstance(recs, list) or not recs:
        raise ValueError("'recommendations' 목록이 없습니다. 지시문의 형식대로 다시 요청하세요.")
    out = []
    for i, r in enumerate(recs, 1):
        if "key" not in r or "crew" not in r:
            raise ValueError(f"{i}번째 추천에 key 또는 crew가 없습니다.")
        try:
            crew = int(float(str(r["crew"]).replace("조", "").strip()))
        except ValueError as e:
            raise ValueError(f"{i}번째 추천의 crew가 숫자가 아닙니다: {r['crew']}") from e
        out.append({"key": str(r["key"]).strip(), "crew": crew, "reason": str(r.get("reason", ""))})
    return {"recommendations": out, "overall": str(data.get("overall", "")),
            "risks": [str(x) for x in (data.get("risks") or [])], "model": "수동(claude.ai)", "usage": {}}


def apply_recommendations(recs: Sequence[Dict[str, Any]], base: Dict[str, int], caps: Dict[str, int]):
    """추천을 앱 조수 조합으로. 상한 초과는 상한으로, 1 미만은 1로 자르고 그 사실을 남긴다.

    반환: (조수 dict — 추천 없는 공종은 base 그대로, {키: 메모}, 모르는 키 목록)
    """
    crew = dict(base)
    notes, unknown = {}, []
    for r in recs:
        k = r["key"]
        if k not in base:
            unknown.append(k)
            continue
        c, cap = int(r["crew"]), int(caps.get(k, 10 ** 6))
        if c > cap:
            notes[k] = f"상한 {cap}조로 조정(추천 {c}조)"
            c = cap
        if c < 1:
            notes[k] = f"1조로 조정(추천 {c}조)"
            c = 1
        crew[k] = c
    return crew, notes, unknown
