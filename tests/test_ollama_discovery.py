"""PROMPT-003 §64 – user-initiated Ollama LAN discovery (fully mocked: no real sockets, no real network)."""
import ipaddress
import threading
import time

import pytest

from app import ollama_discovery as od
from app.ollama_discovery import (LocalInterface, OllamaDiscovery, OllamaDiscoveryResult, candidate_hosts,
                                  discovery_summary_vi, is_private_lan_address, model_available, parse_ip_addr,
                                  parse_ipconfig, select_scan_networks, verify_ollama)

IPCONFIG = """
Windows IP Configuration

Ethernet adapter Ethernet:
   Connection-specific DNS Suffix  . :
   IPv4 Address. . . . . . . . . . . : 192.168.1.23
   Subnet Mask . . . . . . . . . . . : 255.255.255.0
   Default Gateway . . . . . . . . . : 192.168.1.1

Wireless LAN adapter Wi-Fi:
   Media State . . . . . . . . . . . : Media disconnected
   Connection-specific DNS Suffix  . :

Ethernet adapter vEthernet (WSL):
   IPv4 Address. . . . . . . . . . . : 172.29.80.1
   Subnet Mask . . . . . . . . . . . : 255.255.240.0

Ethernet adapter Ethernet 2:
   Autoconfiguration IPv4 Address. . : 169.254.10.5
   Subnet Mask . . . . . . . . . . . : 255.255.0.0
"""


# ---------------------------------------------------------------- 1-6: subnet detection & filtering
def test_01_parse_ipconfig_pairs_ipv4_with_mask_and_skips_disconnected():
    ifs = parse_ipconfig(IPCONFIG)
    assert [(i.address, i.prefix) for i in ifs] == [("192.168.1.23", 24), ("172.29.80.1", 20), ("169.254.10.5", 16)]


def test_02_parse_ip_addr_linux():
    txt = "1: lo    inet 127.0.0.1/8 scope host lo\n2: eth0    inet 10.20.30.40/24 brd 10.20.30.255 scope global eth0\n"
    assert [(i.address, i.prefix) for i in parse_ip_addr(txt)] == [("127.0.0.1", 8), ("10.20.30.40", 24)]


@pytest.mark.parametrize("addr,ok", [("192.168.1.5", True), ("10.0.0.7", True), ("172.16.0.1", True),
                                     ("172.31.255.9", True), ("172.32.0.1", False), ("127.0.0.1", False),
                                     ("169.254.1.1", False), ("8.8.8.8", False), ("0.0.0.0", False),
                                     ("224.0.0.1", False), ("abc", False), ("::1", False)])
def test_03_private_lan_filter(addr, ok):
    assert is_private_lan_address(addr) is ok


def test_04_select_networks_drops_loopback_link_local_public_and_disconnected():
    nets = select_scan_networks([LocalInterface("127.0.0.1", 8), LocalInterface("169.254.10.5", 16),
                                 LocalInterface("203.0.113.9", 24), LocalInterface("192.168.1.23", 24)])
    assert nets == [ipaddress.ip_network("192.168.1.0/24")]


def test_05_large_subnets_are_narrowed_to_own_slash24_and_capped():
    nets = select_scan_networks([LocalInterface("10.5.77.12", 16), LocalInterface("172.29.80.1", 20)])
    assert nets == [ipaddress.ip_network("10.5.77.0/24"), ipaddress.ip_network("172.29.80.0/24")]
    assert sum(n.num_addresses - 2 for n in nets) <= od.MAX_HOSTS
    assert select_scan_networks([LocalInterface("192.168.1.9", 24)], max_hosts=100) == []     # budget respected


def test_06_candidates_exclude_own_address_and_dedupe():
    nets = select_scan_networks([LocalInterface("192.168.1.23", 24), LocalInterface("192.168.1.23", 24)])
    hosts = candidate_hosts(nets, ["192.168.1.23"])
    assert len(nets) == 1 and len(hosts) == 253 and "192.168.1.23" not in hosts
    assert hosts[0] == "192.168.1.1" and hosts[-1] == "192.168.1.254"
    assert len(set(hosts)) == len(hosts)


# ---------------------------------------------------------------- 7-12: probing / verification
def _disc(open_hosts, models_by_host=None, ifaces=None, **kw):
    models_by_host = models_by_host or {}
    calls = {"probe": [], "verify": [], "ports": set()}

    def probe(host, port, timeout):
        calls["probe"].append(host)
        calls["ports"].add(port)
        return host in open_hosts

    def verify(host, port, timeout):
        calls["verify"].append(host)
        calls["ports"].add(port)
        if host in models_by_host:
            return OllamaDiscoveryResult(host, port, list(models_by_host[host]), 5)
        return None
    ifaces = ifaces or [LocalInterface("192.168.1.23", 24)]
    d = OllamaDiscovery(probe=probe, verifier=verify, interfaces=lambda: ifaces, **kw)
    return d, calls


def test_07_only_port_11434_is_probed_and_open_port_requires_api_verification():
    d, calls = _disc({"192.168.1.50", "192.168.1.60"}, {"192.168.1.50": ["qwen3:4b"]})
    res = d.run()
    assert calls["ports"] == {11434}
    assert sorted(calls["verify"]) == ["192.168.1.50", "192.168.1.60"]      # only open ports reach /api/tags
    assert [r.endpoint for r in res] == ["192.168.1.50:11434"]              # 60 is open but not Ollama
    assert res[0].models == ["qwen3:4b"] and res[0].base_url == "http://192.168.1.50:11434"


def test_08_verify_uses_project_ollama_client_and_handles_errors():
    class GoodClient:
        def __init__(self, server, timeout=None): self.server = server
        def list_models(self, timeout=None): return ["qwen3:4b", "llama3:8b"]

    class BadClient(GoodClient):
        def list_models(self, timeout=None): raise RuntimeError("not ollama")

    class WeirdClient(GoodClient):
        def list_models(self, timeout=None): return {"unexpected": 1}
    r = verify_ollama("192.168.1.50", client_factory=GoodClient)
    assert r is not None and r.models == ["qwen3:4b", "llama3:8b"] and r.host == "192.168.1.50"
    assert verify_ollama("192.168.1.50", client_factory=BadClient) is None
    assert verify_ollama("192.168.1.50", client_factory=WeirdClient) is None


def test_09_verify_default_factory_is_ollama_client(monkeypatch):
    import app.ollama_client as oc
    seen = {}

    class Fake:
        def __init__(self, server, timeout=None): seen["server"] = server
        def list_models(self, timeout=None): seen["timeout"] = timeout; return ["qwen3:4b"]
    monkeypatch.setattr(oc, "OllamaClient", Fake)
    assert verify_ollama("10.0.0.5", timeout=2.0).models == ["qwen3:4b"]
    assert seen == {"server": "http://10.0.0.5:11434", "timeout": 2.0}


def test_10_per_host_errors_are_swallowed_and_scan_continues():
    def probe(host, port, timeout):
        if host.endswith(".2"):
            raise OSError("weird NIC error")
        return host.endswith(".3")

    def verify(host, port, timeout):
        raise ValueError("bad json")
    d = OllamaDiscovery(probe=probe, verifier=verify, interfaces=lambda: [LocalInterface("192.168.1.1", 29)])
    assert d.run() == [] and d.checked == d.total == 5


def test_11_no_hosts_when_no_private_interface():
    d, calls = _disc(set(), ifaces=[LocalInterface("127.0.0.1", 8), LocalInterface("8.8.8.8", 24)])
    assert d.run() == [] and d.total == 0 and calls["probe"] == []


def test_12_results_are_deduplicated_and_sorted():
    d, _ = _disc({"192.168.1.9", "192.168.1.100"}, {"192.168.1.9": [], "192.168.1.100": ["qwen3:4b"]})
    res = d.run()
    assert [r.host for r in res] == ["192.168.1.9", "192.168.1.100"]
    assert len({r.endpoint for r in res}) == len(res)


# ---------------------------------------------------------------- 13-17: concurrency / timeouts / cancel / progress
def test_13_bounded_concurrency_and_short_timeouts():
    d, calls = _disc(set(), max_workers=8, connect_timeout=0.2)
    seen_timeouts = []
    orig = d._probe

    def probe(host, port, timeout):
        seen_timeouts.append(timeout)
        time.sleep(0.002)
        return orig(host, port, timeout)
    d._probe = probe
    d.run()
    assert d.peak_workers <= 8 and d.max_workers == 8
    assert set(seen_timeouts) == {0.2}
    assert OllamaDiscovery(max_workers=100000).max_workers == 256                 # hard upper bound


def test_14_progress_callback_monotonic_and_reaches_total():
    d, _ = _disc({"192.168.1.50"}, {"192.168.1.50": ["qwen3:4b"]}, ifaces=[LocalInterface("192.168.1.5", 28)])
    seen = []
    d.run(progress=lambda c, t: seen.append((c, t)))
    assert seen[0] == (0, 13) and seen[-1] == (13, 13)
    assert all(a[0] <= b[0] for a, b in zip(seen, seen[1:]))


def test_15_cancel_stops_early_and_reports_partial():
    started, released = threading.Event(), threading.Event()

    def probe(host, port, timeout):                  # deterministic: probes block until the test cancels
        started.set()
        released.wait(5)
        return False
    d = OllamaDiscovery(probe=probe, verifier=lambda *a: None, interfaces=lambda: [LocalInterface("10.1.1.1", 24)],
                        max_workers=2)
    out = {}

    def work():
        out["res"] = d.run()
    t = threading.Thread(target=work)
    t.start()
    started.wait(5)
    d.cancel()
    released.set()
    t.join(20)
    assert not t.is_alive() and out["res"] == [] and d.cancelled
    assert d.checked < d.total                                   # really stopped early


def test_16_cancel_before_run_probes_nothing():
    d, calls = _disc(set())
    d.cancel()
    d.run()
    assert calls["verify"] == [] and d.results == []


def test_17_found_callback_fires_per_result_during_scan():
    d, _ = _disc({"192.168.1.50"}, {"192.168.1.50": ["qwen3:4b"]})
    found = []
    d.run(found=found.append)
    assert [f.endpoint for f in found] == ["192.168.1.50:11434"]


# ---------------------------------------------------------------- 18-24: result model / messages / policy
def test_18_result_dataclass_fields_and_summary():
    r = OllamaDiscoveryResult("192.168.1.50", 11434, ["qwen3:4b", "llama3"], 42)
    assert (r.host, r.port, r.base_url, r.models, r.latency_ms) == ("192.168.1.50", 11434, "http://192.168.1.50:11434",
                                                                      ["qwen3:4b", "llama3"], 42)
    assert "192.168.1.50:11434" in r.summary_vi() and "qwen3:4b" in r.summary_vi()
    assert "(chưa có model)" in OllamaDiscoveryResult("10.0.0.1").summary_vi()


def test_19_model_available_exact_tag_only():
    assert model_available("qwen3:4b", ["qwen3:4b"]) is True
    assert model_available("qwen3:4b", ["qwen3:1.7b", "qwen3:latest"]) is False
    assert model_available("qwen3", ["qwen3:latest"]) is True
    assert model_available("", ["qwen3:4b"]) is False
    assert od.missing_model_message("qwen3:4b") == "Model qwen3:4b chưa có trên server này.\nVui lòng chọn model có sẵn."


def test_20_no_results_message_matches_spec():
    msg = discovery_summary_vi([], 254, 254, False)
    assert "Không tìm thấy Ollama trong mạng LAN." in msg
    for hint in ("nhập địa chỉ server thủ công", "kiểm tra máy Ollama đã chạy chưa", "kiểm tra firewall/network",
                 "cài Ollama trên máy này nếu cần"):
        assert hint in msg
    assert "(đã dừng)" in discovery_summary_vi([], 10, 254, True)
    assert "Tìm thấy 1 server" in discovery_summary_vi([OllamaDiscoveryResult("10.0.0.2")], 254, 254, False)


def test_21_module_never_scans_automatically_on_import(monkeypatch):
    """Importing the module / constructing a scanner must not touch the network."""
    import importlib
    calls = []
    monkeypatch.setattr(od.socket, "create_connection", lambda *a, **k: calls.append(a))
    importlib.reload(od)
    OllamaDiscovery()
    assert calls == []


def test_22_system_interfaces_fallback_without_tools(monkeypatch):
    monkeypatch.setattr(od, "_run", lambda cmd: "")
    monkeypatch.setattr(od.socket, "getaddrinfo", lambda *a, **k: [(None, None, None, None, ("192.168.7.7", 0))])
    monkeypatch.setattr(od, "_default_route_address", lambda: "10.9.9.9")
    ifs = od.system_interfaces()
    assert sorted(i.address for i in ifs) == ["10.9.9.9", "192.168.7.7"] and all(i.prefix == 24 for i in ifs)


def test_23_windows_path_uses_ipconfig(monkeypatch):
    monkeypatch.setattr(od.platform, "system", lambda: "Windows")
    monkeypatch.setattr(od, "_run", lambda cmd: IPCONFIG if cmd == ["ipconfig"] else "")
    nets = select_scan_networks(od.system_interfaces())
    assert nets == [ipaddress.ip_network("192.168.1.0/24"), ipaddress.ip_network("172.29.80.0/24")]


def test_24_tcp_probe_uses_timeout_and_returns_false_on_error(monkeypatch):
    seen = {}

    class Sock:
        def __enter__(self): return self
        def __exit__(self, *a): return False

    def fake_connect(addr, timeout):
        seen["addr"], seen["timeout"] = addr, timeout
        if addr[0] == "192.168.1.9":
            raise OSError("refused")
        return Sock()
    monkeypatch.setattr(od.socket, "create_connection", fake_connect)
    assert od.tcp_port_open("192.168.1.8") is True and seen == {"addr": ("192.168.1.8", 11434), "timeout": od.CONNECT_TIMEOUT}
    assert od.tcp_port_open("192.168.1.9") is False
    assert od.CONNECT_TIMEOUT <= 0.5 and od.VERIFY_TIMEOUT <= 5
