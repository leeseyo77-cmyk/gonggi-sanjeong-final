# -*- coding: utf-8 -*-
"""
claude_call.py — Claude 구조화 출력 호출과 수동 모드(claude.ai 붙여넣기) 답 해석 공통부

도면 읽기(drawing_reader)와 조수 추천(crew_advisor)이 함께 쓴다.
  - API 모드: 사용자 키로 호출, 정해진 JSON 형식(output_config.format)으로 받는다.
    안전 분류기 거절에 대비해 server-side fallback(권장 모델로 재시도)을 켠다.
  - 수동 모드: 같은 지시문을 claude.ai 채팅에 붙여 넣고, 돌아온 답에서 JSON만 뽑는다.
API 키는 호출에만 쓰고 저장하지 않는다.
"""

from __future__ import annotations

import json
import re
from typing import Any, Dict, List

MODEL = "claude-opus-5-5"
MODELS = {"claude-opus-5-5": "Claude Opus 5.5 (기본·정확)", "claude-sonnet-5-5": "Claude Sonnet 5.5 (저렴)"}
PRICE = {"claude-opus-5-5": (4, 20), "claude-sonnet-5-5": (2, 10)}      # $/백만 토큰 (입력, 출력)


def call_structured(api_key: str, system: str, content: List[Dict[str, Any]], schema: Dict[str, Any],
                    model: str = MODEL, effort: str = "high", client=None) -> Dict[str, Any]:
    """Claude에 한 번 묻고 JSON(dict)을 받는다. 반환에 'model'·'usage'를 붙인다.

    거절·길이 초과·형식 오류는 ValueError. client는 테스트용(가짜 클라이언트 주입).
    """
    if client is None:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key, max_retries=2)
    kwargs = dict(
        model=model, max_tokens=16000, system=system,
        messages=[{"role": "user", "content": content}],
        output_config={"effort": effort, "format": {"type": "json_schema", "schema": schema}},
    )
    if model in ("claude-opus-5-5", "claude-sonnet-5-5"):
        kwargs.update(betas=["server-side-fallback-2026-07-01"], fallbacks="default")
    # 그림·긴 표 + high effort는 오래 걸릴 수 있다 → 스트리밍으로 받아 시간 초과를 피한다
    with client.beta.messages.stream(**kwargs) as stream:
        msg = stream.get_final_message()
    if msg.stop_reason == "refusal":
        raise ValueError("모델이 요청을 거절했습니다(안전 분류기). 내용을 줄이거나 다시 시도하세요.")
    if msg.stop_reason == "max_tokens":
        raise ValueError("응답이 길이 한도에서 잘렸습니다. 넘기는 양을 줄여 다시 시도하세요.")
    text = next((b.text for b in msg.content if getattr(b, "type", "") == "text"), "")
    try:
        data = json.loads(text)
    except json.JSONDecodeError as e:
        raise ValueError(f"응답을 해석하지 못했습니다: {e}") from e
    u = getattr(msg, "usage", None)
    data["model"] = getattr(msg, "model", model)
    data["usage"] = {"input": getattr(u, "input_tokens", 0) or 0, "output": getattr(u, "output_tokens", 0) or 0}
    return data


def extract_json(text: str) -> Any:
    """claude.ai 답에서 JSON만 뽑는다(```json 코드블록·앞뒤 설명 허용)."""
    t = (text or "").strip()
    m = re.search(r"```(?:json)?\s*([\[{].*[\]}])\s*```", t, flags=re.S)
    if m:
        t = m.group(1)
    elif "{" in t:
        t = t[t.index("{"): t.rindex("}") + 1]
    try:
        return json.loads(t)
    except (json.JSONDecodeError, ValueError) as e:
        raise ValueError(f"붙여 넣은 내용에서 JSON을 읽지 못했습니다: {e}") from e


def error_message(e: Exception) -> str:
    """API 예외 → 화면 문구."""
    try:
        import anthropic
    except ImportError:
        return "anthropic 패키지가 설치되지 않았습니다(requirements.txt 확인)."
    if isinstance(e, anthropic.AuthenticationError):
        return "API 키가 올바르지 않습니다."
    if isinstance(e, anthropic.PermissionDeniedError):
        return "이 API 키로는 해당 모델을 쓸 수 없습니다."
    if isinstance(e, anthropic.RateLimitError):
        return "요청 한도를 넘었습니다. 잠시 뒤 다시 시도하세요."
    if isinstance(e, anthropic.APIStatusError):
        return f"API 오류({e.status_code}): {e.message}"
    if isinstance(e, anthropic.APIConnectionError):
        return "Anthropic 서버에 연결하지 못했습니다(네트워크 확인)."
    return str(e)
