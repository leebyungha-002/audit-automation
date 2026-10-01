"""Claude API 호출 래퍼. JSON 스키마 강제 + 결과 캐시 + 사용량 집계."""
import hashlib
import json

import anthropic

from .db import Store, now

FALLBACK_BETA = "server-side-fallback-2026-07-01"


class LLMError(Exception):
    """재시도해도 소용없는 호출 실패(인증, 거절, 잘림 등)."""


class LLM:
    def __init__(self, cfg: dict, store: Store, log):
        self.cfg = cfg["llm"]
        self.store = store
        self.log = log
        self.client = anthropic.Anthropic()  # ANTHROPIC_API_KEY는 환경변수(.env)에서 읽는다
        self.usage = {"calls": 0, "cache_hits": 0, "input_tokens": 0, "output_tokens": 0, "cost_usd": 0.0}

    def _cache_key(self, model: str, system: str, content: list, schema: dict) -> str:
        payload = json.dumps(
            [model, self.cfg["prompt_version"], self.cfg["effort"], system, content, schema],
            ensure_ascii=False, sort_keys=True,
        )
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()

    def _add_usage(self, model: str, input_tokens: int, output_tokens: int) -> None:
        price = self.cfg["pricing_usd_per_mtok"].get(model, {"input": 0, "output": 0})
        self.usage["calls"] += 1
        self.usage["input_tokens"] += input_tokens
        self.usage["output_tokens"] += output_tokens
        self.usage["cost_usd"] += (input_tokens * price["input"] + output_tokens * price["output"]) / 1_000_000

    def call_json(self, purpose: str, model: str, system: str, content: list, schema: dict,
                  use_cache: bool = True) -> dict:
        """content는 Messages API의 user content 블록 리스트. 스키마에 맞는 dict를 돌려준다."""
        key = self._cache_key(model, system, content, schema)
        if use_cache:
            hit = self.store.query("SELECT response_text FROM llm_cache WHERE cache_key=?", (key,))
            if hit:
                self.usage["cache_hits"] += 1
                return json.loads(hit[0]["response_text"])

        kwargs = dict(
            model=model,
            max_tokens=self.cfg["max_tokens"],
            system=system,
            messages=[{"role": "user", "content": content}],
            output_config={"effort": self.cfg["effort"], "format": {"type": "json_schema", "schema": schema}},
        )
        try:
            if self.cfg.get("refusal_fallback"):
                resp = self.client.beta.messages.create(betas=[FALLBACK_BETA], fallbacks="default", **kwargs)
            else:
                resp = self.client.messages.create(**kwargs)
        except anthropic.AuthenticationError as e:
            raise LLMError(f"API 키 인증 실패: {e.message}") from e
        except anthropic.PermissionDeniedError as e:
            raise LLMError(f"API 권한 없음: {e.message}") from e
        except anthropic.BadRequestError as e:
            kind = "크레딧 부족" if "credit" in str(e.message).lower() else "잘못된 요청"
            raise LLMError(f"{kind}: {e.message}") from e
        except anthropic.RateLimitError as e:
            raise LLMError(f"요청 한도 초과(SDK 재시도 후에도 실패): {e.message}") from e
        except anthropic.APIStatusError as e:
            raise LLMError(f"API 오류 {e.status_code}: {e.message}") from e
        except anthropic.APIConnectionError as e:
            raise LLMError(f"네트워크 오류: {e}") from e

        self._add_usage(model, resp.usage.input_tokens, resp.usage.output_tokens)
        if resp.stop_reason == "refusal":
            raise LLMError("모델이 요청을 거절함(stop_reason=refusal)")
        if resp.stop_reason == "max_tokens":
            raise LLMError("출력이 max_tokens에서 잘림 — config의 max_tokens를 늘릴 것")

        text = next((b.text for b in resp.content if b.type == "text"), "")
        data = json.loads(text)  # 파싱 실패(JSONDecodeError)는 호출부에서 재시도
        self.store.upsert("llm_cache", {
            "cache_key": key, "purpose": purpose, "model": resp.model, "response_text": text,
            "input_tokens": resp.usage.input_tokens, "output_tokens": resp.usage.output_tokens,
            "created_at": now(),
        })
        self.store.commit()
        return data

    def usage_summary(self) -> str:
        u = self.usage
        return (f"API 호출 {u['calls']}회, 캐시 적중 {u['cache_hits']}회, "
                f"입력 {u['input_tokens']:,} / 출력 {u['output_tokens']:,} 토큰, 비용 약 ${u['cost_usd']:.3f}")
