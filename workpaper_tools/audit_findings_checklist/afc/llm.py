"""Claude API 호출 래퍼. JSON 스키마 강제 + 결과 캐시 + 사용량 집계 + 배치(Batch API) 지원."""
import hashlib
import json

import anthropic
from anthropic.types.message_create_params import MessageCreateParamsNonStreaming
from anthropic.types.messages.batch_create_params import Request

from .db import Store, now

FALLBACK_BETA = "server-side-fallback-2026-07-01"
BATCH_DISCOUNT = 0.5  # Batch API는 표준 단가의 50%


class LLMError(Exception):
    """재시도해도 소용없는 호출 실패(인증, 거절, 잘림 등)."""


class LLM:
    def __init__(self, cfg: dict, store: Store, log):
        self.cfg = cfg["llm"]
        self.store = store
        self.log = log
        self.client = anthropic.Anthropic()  # ANTHROPIC_API_KEY는 환경변수(.env)에서 읽는다
        self.usage = {"calls": 0, "cache_hits": 0, "input_tokens": 0, "output_tokens": 0, "cost_usd": 0.0}

    # ── 캐시 ──────────────────────────────────────────────
    def cache_key(self, model: str, system: str, content: list, schema: dict) -> str:
        payload = json.dumps(
            [model, self.cfg["prompt_version"], self.cfg["effort"], system, content, schema],
            ensure_ascii=False, sort_keys=True,
        )
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()

    def cache_get(self, key: str) -> dict | None:
        hit = self.store.query("SELECT response_text FROM llm_cache WHERE cache_key=?", (key,))
        if not hit:
            return None
        self.usage["cache_hits"] += 1
        return json.loads(hit[0]["response_text"])

    def cache_put(self, key: str, purpose: str, model: str, text: str, input_tokens: int, output_tokens: int) -> None:
        self.store.upsert("llm_cache", {
            "cache_key": key, "purpose": purpose, "model": model, "response_text": text,
            "input_tokens": input_tokens, "output_tokens": output_tokens, "created_at": now(),
        })
        self.store.commit()

    # ── 요청 구성·응답 해석 ────────────────────────────────
    def params(self, model: str, system: str, content: list, schema: dict) -> dict:
        return dict(
            model=model,
            max_tokens=self.cfg["max_tokens"],
            system=system,
            messages=[{"role": "user", "content": content}],
            output_config={"effort": self.cfg["effort"], "format": {"type": "json_schema", "schema": schema}},
        )

    def add_usage(self, model: str, input_tokens: int, output_tokens: int, discount: float = 1.0) -> None:
        price = self.cfg["pricing_usd_per_mtok"].get(model, {"input": 0, "output": 0})
        self.usage["calls"] += 1
        self.usage["input_tokens"] += input_tokens
        self.usage["output_tokens"] += output_tokens
        self.usage["cost_usd"] += discount * (input_tokens * price["input"] + output_tokens * price["output"]) / 1_000_000

    def parse_message(self, message, key: str, purpose: str, requested_model: str, discount: float = 1.0) -> dict:
        """응답 메시지에서 JSON을 꺼내 캐시에 넣고 돌려준다. 파싱 실패(JSONDecodeError)는 호출부에서 처리."""
        self.add_usage(requested_model, message.usage.input_tokens, message.usage.output_tokens, discount)
        if message.stop_reason == "refusal":
            raise LLMError("모델이 요청을 거절함(stop_reason=refusal)")
        if message.stop_reason == "max_tokens":
            raise LLMError("출력이 max_tokens에서 잘림 — config의 max_tokens를 늘릴 것")
        text = next((b.text for b in message.content if b.type == "text"), "")
        data = json.loads(text)
        self.cache_put(key, purpose, message.model, text, message.usage.input_tokens, message.usage.output_tokens)
        return data

    # ── 즉시 호출 ─────────────────────────────────────────
    def call_json(self, purpose: str, model: str, system: str, content: list, schema: dict,
                  use_cache: bool = True) -> dict:
        """content는 Messages API의 user content 블록 리스트. 스키마에 맞는 dict를 돌려준다."""
        key = self.cache_key(model, system, content, schema)
        if use_cache:
            cached = self.cache_get(key)
            if cached is not None:
                return cached

        kwargs = self.params(model, system, content, schema)
        try:
            if self.cfg.get("refusal_fallback"):
                resp = self.client.beta.messages.create(betas=[FALLBACK_BETA], fallbacks="default", **kwargs)
            else:
                resp = self.client.messages.create(**kwargs)
        except anthropic.APIError as e:
            raise self._translate(e) from e
        return self.parse_message(resp, key, purpose, model)

    @staticmethod
    def _translate(e: anthropic.APIError) -> LLMError:
        if isinstance(e, anthropic.AuthenticationError):
            return LLMError(f"API 키 인증 실패: {e.message}")
        if isinstance(e, anthropic.PermissionDeniedError):
            return LLMError(f"API 권한 없음: {e.message}")
        if isinstance(e, anthropic.BadRequestError):
            kind = "크레딧 부족" if "credit" in str(e.message).lower() else "잘못된 요청"
            return LLMError(f"{kind}: {e.message}")
        if isinstance(e, anthropic.RateLimitError):
            return LLMError(f"요청 한도 초과(SDK 재시도 후에도 실패): {e.message}")
        if isinstance(e, anthropic.APIStatusError):
            return LLMError(f"API 오류 {e.status_code}: {e.message}")
        if isinstance(e, anthropic.APIConnectionError):
            return LLMError(f"네트워크 오류: {e}")
        return LLMError(f"API 오류: {e}")

    # ── 배치 호출 (비동기, 50% 할인. 거절 시 대체 모델 재실행은 배치에서 지원되지 않는다) ──
    def submit_batch(self, model: str, schema: dict, items: list[tuple[str, str, list]]) -> str:
        """items: (custom_id, system, content). 배치 ID를 돌려준다."""
        requests = [Request(custom_id=cid, params=MessageCreateParamsNonStreaming(**self.params(model, system, content, schema)))
                    for cid, system, content in items]
        try:
            return self.client.messages.batches.create(requests=requests).id
        except anthropic.APIError as e:
            raise self._translate(e) from e

    def batch_ended(self, batch_id: str) -> tuple[bool, str]:
        try:
            batch = self.client.messages.batches.retrieve(batch_id)
        except anthropic.APIError as e:
            raise self._translate(e) from e
        c = batch.request_counts
        return batch.processing_status == "ended", f"처리 중 {c.processing}, 성공 {c.succeeded}, 오류 {c.errored}"

    def batch_results(self, batch_id: str):
        try:
            yield from self.client.messages.batches.results(batch_id)
        except anthropic.APIError as e:
            raise self._translate(e) from e

    def usage_summary(self) -> str:
        u = self.usage
        return (f"API 호출 {u['calls']}회, 캐시 적중 {u['cache_hits']}회, "
                f"입력 {u['input_tokens']:,} / 출력 {u['output_tokens']:,} 토큰, 비용 약 ${u['cost_usd']:.3f}")
