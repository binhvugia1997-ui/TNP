"""HTTP client for a remote Ollama server on the LAN (no paid APIs)."""
from __future__ import annotations

import json
import logging
import re
import time
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
        self.last_call: Dict[str, Any] = {}

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
    def build_generate_payload(self, model: str, prompt: str, system: str = "",
                               num_ctx: int = 4096, num_predict: int = 256) -> Dict[str, Any]:
        """Exact body sent to ``POST /api/generate`` (exposed so tests can assert on it)."""
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
        return payload

    def generate_json(self, model: str, prompt: str, system: str = "",
                      num_ctx: int = 4096, num_predict: int = 256) -> Dict[str, Any]:
        """Call ``/api/generate`` with ``format=json``, ``stream=false``, ``think=false``, temperature 0.

        Returns the parsed JSON object produced by the model.  Timing/size statistics of
        the last call are kept in ``self.last_call`` (no model text is stored there).
        """
        payload = self.build_generate_payload(model, prompt, system, num_ctx, num_predict)
        url = f"{self.base}/api/generate"
        self.last_call = {"model": model, "prompt_chars": len(prompt) + len(system), "status": None,
                          "seconds": None, "response_chars": None}
        LOG.info("Ollama POST %s model=%s prompt=%d chars think=false stream=false format=json timeout=%ss",
                 url, model, len(prompt) + len(system), self.timeout)
        t0 = time.perf_counter()
        try:
            r = requests.post(url, json=payload, timeout=self.timeout)
            self.last_call["status"] = r.status_code
            self.last_call["seconds"] = round(time.perf_counter() - t0, 2)
            r.raise_for_status()
            data = r.json()
        except requests.RequestException as e:
            self.last_call["seconds"] = round(time.perf_counter() - t0, 2)
            LOG.warning("Ollama request failed after %.1fs (HTTP %s): %s", self.last_call["seconds"],
                        self.last_call["status"], e)
            raise OllamaError(f"Lỗi gọi Ollama ({url}) sau {self.last_call['seconds']}s: {e}") from e
        text = data.get("response", "") if isinstance(data, dict) else ""
        self.last_call["response_chars"] = len(text)
        if isinstance(data, dict):
            # server-side timings reported by Ollama (nanoseconds)
            for k in ("total_duration", "load_duration", "prompt_eval_duration", "eval_duration"):
                if isinstance(data.get(k), (int, float)):
                    self.last_call[k.replace("_duration", "_s")] = round(data[k] / 1e9, 2)
            for k in ("prompt_eval_count", "eval_count"):
                if isinstance(data.get(k), int):
                    self.last_call[k] = data[k]
            if data.get("thinking"):
                self.last_call["thinking_chars"] = len(str(data["thinking"]))   # length only, never logged
        LOG.info("Ollama response: HTTP %s in %.1fs, %d chars, load %ss, prompt_eval %s tok, eval %s tok",
                 self.last_call["status"], self.last_call["seconds"], len(text),
                 self.last_call.get("load_s", "?"), self.last_call.get("prompt_eval_count", "?"),
                 self.last_call.get("eval_count", "?"))
        # some servers ignore "think": strip a <think>...</think> preamble if present
        text = re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL)
        return parse_json_response(text)

    def smoke_test(self, model: str) -> Dict[str, Any]:
        """Tiny classification request through the SAME code path as production."""
        prompt = ('Slides of a PowerPoint quality report (2 slides). Each line: S<number> | title | pics | text.\n'
                  'Return ONLY this JSON: {"qpn_slide":int|null,"cause_slides":[int],"confidence":0..1}\n'
                  'S1 | Quality Problem Notice | pics=1 | Model A185 Xước 15ea\n'
                  'S2 | 2. NGUYÊN NHÂN | pics=0 | Khay chứa không có lớp lót')
        t0 = time.perf_counter()
        try:
            obj = self.generate_json(model, prompt, system=SYSTEM_SMOKE, num_ctx=1024, num_predict=64)
            ok = True
            err = ""
        except OllamaError as e:
            obj, ok, err = {}, False, str(e)
        return {"ok": ok, "error": err, "seconds": round(time.perf_counter() - t0, 2),
                "result": obj, "stats": dict(getattr(self, "last_call", {}) or {}),
                "expected": {"qpn_slide": 1, "cause_slides": [2]}}


SYSTEM_SMOKE = "You classify slides. Answer with one compact JSON object of slide numbers only. No reasoning text."


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
