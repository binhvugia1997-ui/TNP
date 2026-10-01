"""User-initiated Ollama LAN discovery (PROMPT-003).

Deliberately narrow: **only** TCP port 11434 on the active private IPv4 subnet(s) of this PC, bounded host
count, short timeouts, bounded thread pool, cooperative cancellation.  A host whose port is open is reported
only after the real Ollama HTTP API (``GET /api/tags`` through :class:`app.ollama_client.OllamaClient`)
answered – so the result also carries the server's installed model list.  Nothing here runs automatically:
the GUI starts a scan only when the user clicks "Tìm Ollama trong mạng LAN", and the GUI never switches
server without the user's explicit confirmation.

No nmap, no Internet / public ranges, no arbitrary ports, no firewall changes, no admin rights, no telemetry.
"""
from __future__ import annotations

import ipaddress
import platform
import re
import socket
import subprocess
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from typing import Callable, Iterable, List, Optional, Sequence, Tuple

OLLAMA_PORT = 11434
MAX_HOSTS = 1024               # never probe more than this many addresses in one scan (/22 worth)
MAX_PREFIX_HOSTS = 254         # subnets larger than /24 are narrowed to the /24 around this PC's own address
CONNECT_TIMEOUT = 0.35         # seconds per TCP probe
VERIFY_TIMEOUT = 3.0           # seconds for GET /api/tags on an open port
MAX_WORKERS = 64

PRIVATE_NETS = (ipaddress.ip_network("10.0.0.0/8"), ipaddress.ip_network("172.16.0.0/12"),
                ipaddress.ip_network("192.168.0.0/16"))
NO_RESULTS_VI = ("Không tìm thấy Ollama trong mạng LAN.\n\nBạn có thể:\n"
                 "• nhập địa chỉ server thủ công;\n• kiểm tra máy Ollama đã chạy chưa;\n"
                 "• kiểm tra firewall/network;\n• cài Ollama trên máy này nếu cần.")


# ----------------------------------------------------------------------------- result model
@dataclass
class OllamaDiscoveryResult:
    host: str
    port: int = OLLAMA_PORT
    models: List[str] = field(default_factory=list)
    latency_ms: int = 0

    @property
    def base_url(self) -> str:
        return f"http://{self.host}:{self.port}"

    @property
    def endpoint(self) -> str:
        return f"{self.host}:{self.port}"

    def summary_vi(self) -> str:
        models = ", ".join(self.models) if self.models else "(chưa có model)"
        return f"{self.endpoint}\nModels: {models}\nTrạng thái: Sẵn sàng ({self.latency_ms} ms)"


@dataclass
class LocalInterface:
    address: str
    prefix: int = 24

    @property
    def ip(self) -> ipaddress.IPv4Address:
        return ipaddress.ip_address(self.address)

    @property
    def network(self) -> ipaddress.IPv4Network:
        return ipaddress.ip_network(f"{self.address}/{self.prefix}", strict=False)


# ----------------------------------------------------------------------------- local subnet detection
def is_private_lan_address(address: str) -> bool:
    """RFC1918 unicast address that is neither loopback, link-local (169.254/16), multicast nor public."""
    try:
        ip = ipaddress.ip_address(address)
    except ValueError:
        return False
    if ip.version != 4 or ip.is_loopback or ip.is_link_local or ip.is_multicast or ip.is_unspecified:
        return False
    return any(ip in net for net in PRIVATE_NETS)


def _mask_to_prefix(mask: str) -> int:
    try:
        return ipaddress.IPv4Network(f"0.0.0.0/{mask}").prefixlen
    except ValueError:
        return 24


def parse_ipconfig(text: str) -> List[LocalInterface]:
    """IPv4 address / subnet mask pairs from Windows ``ipconfig`` output (adapters without an address, i.e.
    disconnected ones, simply have no pair)."""
    out: List[LocalInterface] = []
    addr: Optional[str] = None
    for line in text.splitlines():
        m = re.search(r"IPv4[^:]*:\s*([0-9.]+)", line)
        if m:
            addr = m.group(1)
            continue
        m = re.search(r"(?:Subnet Mask|Mặt nạ mạng con)[^:]*:\s*([0-9.]+)", line)
        if m and addr:
            out.append(LocalInterface(addr, _mask_to_prefix(m.group(1))))
            addr = None
    return out


def parse_ip_addr(text: str) -> List[LocalInterface]:
    """IPv4 interfaces from Linux ``ip -4 -o addr`` output."""
    out: List[LocalInterface] = []
    for m in re.finditer(r"inet\s+([0-9.]+)/(\d+)", text):
        out.append(LocalInterface(m.group(1), int(m.group(2))))
    return out


def _run(cmd: Sequence[str]) -> str:
    try:
        return subprocess.run(list(cmd), capture_output=True, text=True, timeout=5, encoding="utf-8",
                              errors="replace").stdout
    except Exception:  # noqa: BLE001
        return ""


def _default_route_address() -> Optional[str]:
    """Address of the interface that would carry outbound traffic (UDP connect trick – nothing is sent)."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            s.connect(("10.255.255.255", 1))
            return s.getsockname()[0]
        finally:
            s.close()
    except OSError:
        return None


def system_interfaces() -> List[LocalInterface]:
    """Best-effort IPv4 interface list of this PC (ipconfig on Windows, ``ip`` on Linux, getaddrinfo fallback)."""
    ifaces: List[LocalInterface] = []
    if platform.system() == "Windows":
        ifaces = parse_ipconfig(_run(["ipconfig"]))
    else:
        ifaces = parse_ip_addr(_run(["ip", "-4", "-o", "addr"]))
    if not ifaces:
        seen = set()
        try:
            for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
                seen.add(info[4][0])
        except OSError:
            pass
        d = _default_route_address()
        if d:
            seen.add(d)
        ifaces = [LocalInterface(a, 24) for a in sorted(seen)]
    return ifaces


def select_scan_networks(ifaces: Iterable[LocalInterface], max_hosts: int = MAX_HOSTS) -> List[ipaddress.IPv4Network]:
    """Private, active, non-loopback interfaces → bounded networks to probe.

    Public / loopback / link-local / unspecified addresses are dropped.  Networks wider than /24 (large VPN,
    corporate /16, Hyper-V /20 …) are narrowed to the /24 that contains this PC's own address so one scan never
    exceeds a few hundred hosts.  Duplicate networks are merged, the total is capped at ``max_hosts``.
    """
    nets: List[ipaddress.IPv4Network] = []
    for itf in ifaces:
        if not is_private_lan_address(itf.address):
            continue
        prefix = itf.prefix if 8 <= itf.prefix <= 30 else 24
        net = ipaddress.ip_network(f"{itf.address}/{prefix}", strict=False)
        if net.num_addresses - 2 > MAX_PREFIX_HOSTS:
            net = ipaddress.ip_network(f"{itf.address}/24", strict=False)
        if net not in nets:
            nets.append(net)
    out: List[ipaddress.IPv4Network] = []
    budget = max_hosts
    for net in nets:
        hosts = max(0, net.num_addresses - 2)
        if hosts > budget:
            break
        out.append(net)
        budget -= hosts
    return out


def candidate_hosts(networks: Iterable[ipaddress.IPv4Network], own_addresses: Iterable[str] = ()) -> List[str]:
    own = set(own_addresses)
    hosts: List[str] = []
    for net in networks:
        for ip in net.hosts():
            a = str(ip)
            if a not in own and a not in hosts:
                hosts.append(a)
    return hosts


# ----------------------------------------------------------------------------- probing / verification
def tcp_port_open(host: str, port: int = OLLAMA_PORT, timeout: float = CONNECT_TIMEOUT) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False


def verify_ollama(host: str, port: int = OLLAMA_PORT, timeout: float = VERIFY_TIMEOUT,
                  client_factory: Optional[Callable] = None) -> Optional[OllamaDiscoveryResult]:
    """An open port is not enough: the host must answer the real Ollama ``/api/tags`` API (via the project's
    :class:`OllamaClient`).  Returns the result with the server's model list, or ``None``."""
    if client_factory is None:
        from .ollama_client import OllamaClient
        client_factory = OllamaClient
    t0 = time.monotonic()
    try:
        models = client_factory(f"http://{host}:{port}").list_models(timeout=timeout)
    except Exception:  # noqa: BLE001 – OllamaError, HTTP errors, non-JSON answers: not an Ollama server
        return None
    if not isinstance(models, list):
        return None
    return OllamaDiscoveryResult(host, port, [str(m) for m in models], int((time.monotonic() - t0) * 1000))


class OllamaDiscovery:
    """One bounded, cancellable scan.  ``progress(checked, total)`` and ``found(result)`` callbacks are invoked
    from worker threads – the GUI forwards them through its event queue, never touching Tk directly."""

    def __init__(self, port: int = OLLAMA_PORT, max_workers: int = MAX_WORKERS, connect_timeout: float = CONNECT_TIMEOUT,
                 verify_timeout: float = VERIFY_TIMEOUT, probe: Optional[Callable[[str, int, float], bool]] = None,
                 verifier: Optional[Callable[..., Optional[OllamaDiscoveryResult]]] = None,
                 interfaces: Optional[Callable[[], List[LocalInterface]]] = None, max_hosts: int = MAX_HOSTS):
        self.port = port
        self.max_workers = max(1, min(int(max_workers), 256))
        self.connect_timeout = connect_timeout
        self.verify_timeout = verify_timeout
        self._probe = probe or tcp_port_open
        self._verify = verifier or verify_ollama
        self._interfaces = interfaces or system_interfaces
        self.max_hosts = max_hosts
        self.cancel_event = threading.Event()
        self.checked = 0
        self.total = 0
        self.results: List[OllamaDiscoveryResult] = []
        self.networks: List[ipaddress.IPv4Network] = []
        self.peak_workers = 0
        self._active = 0
        self._lock = threading.Lock()

    def cancel(self) -> None:
        self.cancel_event.set()

    @property
    def cancelled(self) -> bool:
        return self.cancel_event.is_set()

    def plan(self) -> List[str]:
        ifaces = self._interfaces()
        self.networks = select_scan_networks(ifaces, self.max_hosts)
        return candidate_hosts(self.networks, (i.address for i in ifaces))

    def _check(self, host: str) -> Optional[OllamaDiscoveryResult]:
        with self._lock:
            self._active += 1
            self.peak_workers = max(self.peak_workers, self._active)
        try:
            if self.cancel_event.is_set():
                return None
            try:
                if not self._probe(host, self.port, self.connect_timeout):
                    return None
            except Exception:  # noqa: BLE001 – one bad host never aborts the scan
                return None
            if self.cancel_event.is_set():
                return None
            try:
                return self._verify(host, self.port, self.verify_timeout)
            except Exception:  # noqa: BLE001
                return None
        finally:
            with self._lock:
                self._active -= 1
                self.checked += 1

    def run(self, progress: Optional[Callable[[int, int], None]] = None,
            found: Optional[Callable[[OllamaDiscoveryResult], None]] = None) -> List[OllamaDiscoveryResult]:
        hosts = self.plan()
        self.total = len(hosts)
        self.checked = 0
        self.results = []
        if progress:
            progress(0, self.total)
        if not hosts:
            return self.results
        seen: set = set()
        with ThreadPoolExecutor(max_workers=self.max_workers, thread_name_prefix="ollama-discovery") as pool:
            futures = {pool.submit(self._check, h): h for h in hosts}
            for fut in as_completed(futures):
                res = fut.result()
                if res is not None and res.endpoint not in seen:
                    seen.add(res.endpoint)
                    self.results.append(res)
                    if found:
                        found(res)
                if progress:
                    progress(self.checked, self.total)
                if self.cancel_event.is_set():
                    for f in futures:
                        f.cancel()
                    break
        self.results.sort(key=lambda r: tuple(int(x) for x in r.host.split(".")))
        return self.results


def model_available(model: str, models: Sequence[str]) -> bool:
    """Exact tag match first, then the bare name (``qwen3:4b`` ↔ ``qwen3:4b`` only; ``qwen3`` ↔ ``qwen3:latest``)."""
    m = (model or "").strip()
    if not m:
        return False
    if m in models:
        return True
    return ":" not in m and f"{m}:latest" in models


def missing_model_message(model: str) -> str:
    return f"Model {model} chưa có trên server này.\nVui lòng chọn model có sẵn."


def discovery_summary_vi(results: Sequence[OllamaDiscoveryResult], checked: int, total: int, cancelled: bool) -> str:
    head = f"Đã kiểm tra {checked} / {total} địa chỉ" + (" (đã dừng)" if cancelled else "")
    if not results:
        return head + "\n\n" + NO_RESULTS_VI
    return head + f"\nTìm thấy {len(results)} server Ollama trong mạng LAN."


def split_endpoint_tuple(endpoint: str) -> Tuple[str, int]:
    host, _, port = endpoint.partition(":")
    return host, int(port or OLLAMA_PORT)
