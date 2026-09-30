"""HTTP client for a remote Ollama server on the LAN (no paid APIs)."""
from __future__ import annotations

import json
import logging
import re
from typing import Any, Dict, List, Optional

import requests

from .config import normalize_ollama_url

LOG = logging.getLogger("report_extractor.ollama")


class OllamaError(RuntimeError):
    pass


def parse_models(tags_json: Dict[str, Any]) -> List[str]:
    """Extract model names from a ``GET /api/tags`` payload."""
    models: List[str] = []
    for m in (tags_json or {}).get("models", []) or []:
        name = m.get("name") or m.get("model")
        if name and name not in models:
            models.append(name)
    return models


def preferred_model(models: List[str]) -> Optional[str]:
    """Pick the best default: name containing both ``qwen`` and ``4b``; else any qwen; else first."""
    def score(name: str) -> int:
        n = name.lower()
        s = 0
        if "qwen" in n:
            s += 10
        if re.search(r"(^|[^0-9.])4b\b", n) or ":4b" in n or "-4b" in n:
            s += 5
        if "instruct" in n:
            s += 1
        if "embed" in n:
            s -= 50
        return s
    if not models:
        return None
    best = max(models, key=score)
    return best


class OllamaClient:
    def __init__(self, server: str, timeout: int = 180):
        self.base = normalize_ollama_url(server)
        self.timeout = timeout

    # ------------------------------------------------------------------
    def list_models(self, timeout: Optional[float] = None) -> List[str]:
        url = f"{self.base}/api/tags"
        try:
            r = requests.get(url, timeout=timeout or 8)
            r.raise_for_status()
            return parse_models(r.json())
        except requests.RequestException as e:
            raise OllamaError(f"Không kết nối được Ollama tại {self.base}: {e}") from e
        except ValueError as e:
            raise OllamaError(f"Phản hồi không hợp lệ từ {url}: {e}") from e

    def test_connection(self) -> Dict[str, Any]:
        models = self.list_models()
        return {"ok": True, "server": self.base, "models": models,
                "preferred": preferred_model(models)}

    # ------------------------------------------------------------------
    def generate_json(self, model: str, prompt: str, system: str = "",
                      num_ctx: int = 8192, num_predict: int = 1024) -> Dict[str, Any]:
        """Call ``/api/generate`` with ``format=json`` and temperature 0.

        Returns the parsed JSON object produced by the model.
        """
        payload = {
            "model": model,
            "prompt": prompt,
            "stream": False,
            "format": "json",
            "keep_alive": "10m",
            # qwen3 family: disable "thinking" so the response is the JSON object only
            "think": False,
            "options": {
                "temperature": 0,
                "top_p": 1,
                "seed": 0,
                "num_ctx": num_ctx,
                "num_predict": num_predict,
            },
        }
        if system:
            payload["system"] = system
        url = f"{self.base}/api/generate"
        try:
            r = requests.post(url, json=payload, timeout=self.timeout)
            r.raise_for_status()
            data = r.json()
        except requests.RequestException as e:
            raise OllamaError(f"Lỗi gọi Ollama ({url}): {e}") from e
        text = data.get("response", "") if isinstance(data, dict) else ""
        # some servers ignore "think": strip a <think>...</think> preamble if present
        text = re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL)
        return parse_json_response(text)


def parse_json_response(text: str) -> Dict[str, Any]:
    """Robustly parse a JSON object from LLM output (handles code fences / prose)."""
    if not text:
        raise OllamaError("Ollama trả về nội dung rỗng")
    t = text.strip()
    t = re.sub(r"^```(?:json)?", "", t).strip()
    t = re.sub(r"```$", "", t).strip()
    try:
        obj = json.loads(t)
        if isinstance(obj, dict):
            return obj
    except json.JSONDecodeError:
        pass
    # find first {...} balanced object
    start = t.find("{")
    while start != -1:
        depth = 0
        for i in range(start, len(t)):
            if t[i] == "{":
                depth += 1
            elif t[i] == "}":
                depth -= 1
                if depth == 0:
                    try:
                        obj = json.loads(t[start:i + 1])
                        if isinstance(obj, dict):
                            return obj
                    except json.JSONDecodeError:
                        break
        start = t.find("{", start + 1)
    raise OllamaError(f"Không phân tích được JSON từ Ollama: {text[:200]!r}")
