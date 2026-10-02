"""OllamaClient against a fake local HTTP server (same wire protocol as Ollama)."""
import json
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

import pytest

from app.ollama_client import OllamaClient, OllamaError


class Handler(BaseHTTPRequestHandler):
    last_payload = None

    def log_message(self, *a):  # silence
        pass

    def _send(self, obj, code=200):
        body = json.dumps(obj).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path == "/api/tags":
            self._send({"models": [{"name": "qwen3:4b"}, {"name": "llama3:8b"}]})
        else:
            self._send({"error": "nf"}, 404)

    def do_POST(self):
        n = int(self.headers.get("Content-Length", 0))
        Handler.last_payload = json.loads(self.rfile.read(n))
        if Handler.last_payload.get("prompt", "").startswith("Slides of a PowerPoint quality report (2 slides)"):
            resp = {"qpn_slide": 1, "cause_slides": [2], "confidence": 0.9}      # smoke-test prompt
        else:
            resp = {"qpn_slide": 2, "cause_slides": [3]}
        self._send({"model": "qwen3:4b", "response": json.dumps(resp), "done": True,
                    "total_duration": 1_500_000_000, "load_duration": 200_000_000,
                    "prompt_eval_count": 120, "prompt_eval_duration": 300_000_000,
                    "eval_count": 25, "eval_duration": 900_000_000})


@pytest.fixture(scope="module")
def server():
    srv = HTTPServer(("127.0.0.1", 0), Handler)
    t = threading.Thread(target=srv.serve_forever, daemon=True)
    t.start()
    yield f"127.0.0.1:{srv.server_address[1]}"
    srv.shutdown()


def test_test_connection_and_model_detection(server):
    info = OllamaClient(server).test_connection()
    assert info["ok"] and info["server"].startswith("http://")
    assert info["models"] == ["qwen3:4b", "llama3:8b"]
    assert info["preferred"] == "qwen3:4b"


def test_generate_json_uses_temperature_zero_and_json_format(server):
    out = OllamaClient(server).generate_json("qwen3:4b", "hello", system="sys")
    assert out == {"qpn_slide": 2, "cause_slides": [3]}
    p = Handler.last_payload
    assert p["options"]["temperature"] == 0 and p["format"] == "json" and p["stream"] is False
    assert p["model"] == "qwen3:4b" and p["system"] == "sys"


def test_connection_failure_is_clean_error():
    with pytest.raises(OllamaError):
        OllamaClient("127.0.0.1:1").list_models(timeout=1)


def test_every_generate_request_disables_thinking(server, sample_tree):
    """Captures the real HTTP body: think=false, stream=false, format=json for the production prompt."""
    from app.classifier import classify
    from app.pptx_parser import parse_pptx
    client = OllamaClient(server)
    report = parse_pptx(sample_tree["files"][0])
    cls = classify(report, client, "qwen3:4b")
    p = Handler.last_payload
    assert p["think"] is False
    assert p["stream"] is False and p["format"] == "json"
    assert p["options"]["temperature"] == 0
    assert p["options"]["num_predict"] <= 256 and p["options"]["num_ctx"] <= 4096
    assert p["model"] == "qwen3:4b"
    assert cls.source.startswith("qwen")
    # the smoke test goes through the very same code path / payload builder
    res = client.smoke_test("qwen3:4b")
    assert res["ok"] and res["result"]["qpn_slide"] == 1
    assert Handler.last_payload["think"] is False and Handler.last_payload["stream"] is False
    assert Handler.last_payload["format"] == "json"


def test_generate_records_timing_diagnostics(server):
    client = OllamaClient(server)
    client.generate_json("qwen3:4b", "hello", system="sys")
    st = client.last_call
    assert st["status"] == 200 and isinstance(st["seconds"], float)
    assert st["response_chars"] > 0 and st["prompt_chars"] == len("hello") + len("sys")
    assert st["load_s"] == 0.2 and st["prompt_eval_count"] == 120 and st["eval_count"] == 25
    assert "thinking_chars" not in st                         # nothing hidden is stored or logged


def test_timeout_records_duration_and_is_clean_error():
    client = OllamaClient("127.0.0.1:1", timeout=1)
    with pytest.raises(OllamaError) as ei:
        client.generate_json("qwen3:4b", "x")
    assert "Lỗi gọi Ollama" in str(ei.value) and client.last_call["seconds"] is not None
