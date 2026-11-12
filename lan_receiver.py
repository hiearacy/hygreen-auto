"""Opt-in, authenticated local-LAN notification receiver.

Accepts only a narrowly scoped sender name and UTF-8 text from a paired device.
Never opens/controls the KakaoTalk desktop application or an Excel workbook.
"""
from __future__ import annotations

import hmac
import ipaddress
import os
import secrets
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from uuid import uuid4

PORT = 8765
PATH = "/api/v1/notification"
MAX_BODY = 16 * 1024
SENDER_NAMES = {"NOL 알림", "여기어때"}
LAN_RANGES = tuple(ipaddress.ip_network(x) for x in (
    "127.0.0.0/8", "10.0.0.0/8", "172.16.0.0/12", "192.168.0.0/16"))


def permitted_address(host: str) -> bool:
    try:
        ip = ipaddress.ip_address(host)
        return isinstance(ip, ipaddress.IPv4Address) and any(ip in net for net in LAN_RANGES)
    except ValueError:
        return False


def safe_sender(content: str) -> bool:
    """Require the MacroDroid's notification title on the first line."""
    first = content.lstrip("\ufeff\r\n ").splitlines()[0].strip() if content.strip() else ""
    return first in SENDER_NAMES


def get_local_ip() -> str:
    """Guess the Windows PC LAN IP without sending network traffic."""
    import socket
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.connect(("192.0.2.1", 80))  # UDP connect: no datagram sent.
            address = sock.getsockname()[0]
            if permitted_address(address) and not address.startswith("127."):
                return address
    except OSError:
        pass
    return "PC의 IPv4 주소 확인 필요"


class NotificationReceiver:
    """Background HTTP receiver. Writes atomically to a private local inbox."""

    def __init__(self, inbox: Path, token: str, host: str = "0.0.0.0", port: int = PORT):
        if len(token) < 32:
            raise ValueError("수신 인증키가 너무 짧습니다")
        self.inbox = Path(inbox)
        self.token = token
        self.host = host
        self.port = port
        self.server: ThreadingHTTPServer | None = None
        self.thread: threading.Thread | None = None

    def start(self) -> int:
        if self.server:
            return self.server.server_address[1]
        self.inbox.mkdir(parents=True, exist_ok=True)
        inbox = self.inbox
        secret = self.token

        class Handler(BaseHTTPRequestHandler):
            server_version = "HygreenAuto/2"
            sys_version = ""

            def log_message(self, fmt, *args):
                return  # Never log contact names, request bodies or the token.

            def reply(self, code: int, msg: str):
                body = (msg + "\n").encode("utf-8")
                self.send_response(code)
                self.send_header("Content-Type", "text/plain; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def do_POST(self):
                if self.path != PATH:
                    return self.reply(404, "not found")
                if not permitted_address(self.client_address[0]):
                    return self.reply(403, "local-network access only")
                supplied = self.headers.get("X-Hygreen-Token", "")
                if not hmac.compare_digest(supplied, secret):
                    return self.reply(401, "unauthorized")
                content_type = self.headers.get("Content-Type", "").lower()
                if not content_type.startswith("text/plain"):
                    return self.reply(415, "text/plain required")
                try:
                    size = int(self.headers.get("Content-Length", "-1"))
                except ValueError:
                    return self.reply(400, "invalid length")
                if size <= 0 or size > MAX_BODY:
                    return self.reply(413, "invalid content size")
                content = self.rfile.read(size)
                try:
                    message = content.decode("utf-8-sig").strip()
                except UnicodeError:
                    return self.reply(400, "invalid utf-8")
                if not safe_sender(message):
                    return self.reply(422, "only NOL 알림 or 여기어때 allowed")
                # Do not trust HTTP clients to write paths; all names generated locally.
                path = inbox / f"notice_{time.time_ns()}_{uuid4().hex}.txt"
                temp = inbox / (path.name + ".part")
                try:
                    with open(temp, "x", encoding="utf-8") as f:
                        f.write(message)
                        f.write("\n")
                        f.flush()
                        os.fsync(f.fileno())
                    temp.replace(path)
                except OSError:
                    try: temp.unlink(missing_ok=True)
                    except OSError: pass
                    return self.reply(500, "unable to queue notice")
                return self.reply(202, "received; pending validation")

            def do_GET(self):
                # Health endpoint is available only on the local network, no customer data.
                if self.path != "/health" or not permitted_address(self.client_address[0]):
                    return self.reply(404, "not found")
                return self.reply(200, "hygreen-auto receiver ready")

        class Server(ThreadingHTTPServer):
            daemon_threads = True
            allow_reuse_address = True

        self.server = Server((self.host, self.port), Handler)
        self.thread = threading.Thread(target=self.server.serve_forever, name="HygreenLanReceiver", daemon=True)
        self.thread.start()
        return self.server.server_address[1]

    def stop(self) -> None:
        server = self.server
        self.server = None
        if server:
            server.shutdown()
            server.server_close()
        if self.thread:
            self.thread.join(timeout=3)
        self.thread = None


def new_token() -> str:
    return secrets.token_urlsafe(32)