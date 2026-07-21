from __future__ import annotations

import json
import time
import urllib.error
import urllib.request
from dataclasses import dataclass
from typing import Callable

from .config import Settings

AI_SCHEMA = {
    "name": "knowledge_preprocessor_result",
    "strict": True,
    "schema": {"type": "object", "additionalProperties": False, "required": ["summary", "keywords", "categories", "machine_models", "cnc_systems", "visibility", "version", "release_date", "knowledge_points"], "properties": {
        "summary": {"type": "string"}, "keywords": {"type": "array", "items": {"type": "string"}},
        "categories": {"type": "array", "items": {"type": "object", "additionalProperties": False, "required": ["name", "confidence", "reason", "evidence"], "properties": {"name": {"type": "string"}, "confidence": {"type": "number"}, "reason": {"type": "string"}, "evidence": {"type": "string"}}}},
        "machine_models": {"type": "array", "items": {"type": "string"}},
        "cnc_systems": {"type": "array", "items": {"type": "object", "additionalProperties": False, "required": ["brand", "model", "version", "evidence"], "properties": {"brand": {"type": "string"}, "model": {"type": "string"}, "version": {"type": "string"}, "evidence": {"type": "string"}}}},
        "visibility": {"type": "string"}, "version": {"type": "string"}, "release_date": {"type": "string"}, "knowledge_points": {"type": "array", "items": {"type": "string"}}
    }}
}


class AIResponseError(ValueError): pass


def validate_ai_result(value: object) -> dict:
    if not isinstance(value, dict) or set(value) != set(AI_SCHEMA["schema"]["required"]):
        raise AIResponseError("AI 返回字段不符合严格 Schema")
    if not isinstance(value["categories"], list) or not isinstance(value["machine_models"], list) or not isinstance(value["cnc_systems"], list):
        raise AIResponseError("AI 返回列表字段不合法")
    for category in value["categories"]:
        if not isinstance(category, dict) or set(category) != {"name", "confidence", "reason", "evidence"} or not isinstance(category["confidence"], (int, float)) or not 0 <= category["confidence"] <= 1:
            raise AIResponseError("分类对象不合法")
    return value


@dataclass
class AIAnalyzer:
    settings: Settings
    transport: Callable[[dict], dict] | None = None

    def analyze(self, context: str) -> dict | None:
        if not self.settings.ai_enabled or not self.settings.api_base_url or not self.settings.text_model:
            return None
        last_error: Exception | None = None
        for attempt in range(self.settings.retries + 1):
            try:
                response = self.transport(self._payload(context)) if self.transport else self._request(self._payload(context))
                content = response["choices"][0]["message"]["content"]
                return validate_ai_result(json.loads(content))
            except (KeyError, TypeError, json.JSONDecodeError, urllib.error.URLError, TimeoutError, AIResponseError) as exc:
                last_error = exc
                if attempt < self.settings.retries: time.sleep(min(2 ** attempt, 3))
        raise AIResponseError(f"AI 分析失败：{type(last_error).__name__}")

    def _payload(self, context: str) -> dict:
        instruction = "只根据提供材料输出 JSON；没有证据的字段填待人工确认，禁止推测。"
        return {"model": self.settings.text_model, "messages": [{"role": "system", "content": instruction}, {"role": "user", "content": context[:120000]}], "response_format": {"type": "json_schema", "json_schema": AI_SCHEMA}, "temperature": 0}

    def _request(self, payload: dict) -> dict:
        url = self.settings.api_base_url.rstrip("/") + "/chat/completions"
        headers = {"Content-Type": "application/json"}
        if self.settings.api_key: headers["Authorization"] = f"Bearer {self.settings.api_key}"
        request = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"), headers=headers, method="POST")
        with urllib.request.urlopen(request, timeout=self.settings.timeout_seconds) as response:
            return json.loads(response.read().decode("utf-8"))

    def embeddings(self, texts: list[str]) -> list[list[float]] | None:
        """Optional OpenAI-compatible semantic vectors; callers fall back safely on failure."""
        if not self.settings.semantic_similarity_enabled or not self.settings.api_base_url or not self.settings.embedding_model:
            return None
        headers = {"Content-Type": "application/json"}
        if self.settings.api_key: headers["Authorization"] = f"Bearer {self.settings.api_key}"
        payload = {"model": self.settings.embedding_model, "input": [item[:120000] for item in texts]}
        request = urllib.request.Request(self.settings.api_base_url.rstrip("/") + "/embeddings", data=json.dumps(payload).encode("utf-8"), headers=headers, method="POST")
        with urllib.request.urlopen(request, timeout=self.settings.timeout_seconds) as response:
            data = json.loads(response.read().decode("utf-8"))["data"]
        return [item["embedding"] for item in sorted(data, key=lambda item: item["index"])]
