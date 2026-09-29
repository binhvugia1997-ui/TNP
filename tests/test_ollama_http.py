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
        self._send({"model": "qwen3:4b", "response": json.dumps({"qpn_slide": 2, "cause_slides": [3]}), "done": True})


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
