"""
Lightweight smoke test for the Banki TXT konverter HTTP server.

Spins up `python app.py`, hits the main routes, and checks status codes
plus a handful of HTML markers from the redesigned UI. No browser is
needed and the only runtime dependency is stdlib + `openpyxl` (already
in requirements.txt).

Run:
    python -m tests.test_smoke
or  python tests/test_smoke.py
"""
from __future__ import annotations

import os
import socket
import subprocess
import sys
import time
import unittest
import urllib.request
import urllib.error
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
HOST = "127.0.0.1"
PORT = 8765
BASE = f"http://{HOST}:{PORT}"


def _wait_for_port(host: str, port: int, timeout: float = 15.0) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with socket.create_connection((host, port), timeout=1.0):
                return True
        except OSError:
            time.sleep(0.25)
    return False


def _http_get(path: str, timeout: float = 5.0) -> tuple[int, bytes, str]:
    req = urllib.request.Request(BASE + path, headers={"User-Agent": "banki-smoke/1.0"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.status, r.read(), r.headers.get("Content-Type", "")
    except urllib.error.HTTPError as e:
        return e.code, e.read() if e.fp else b"", e.headers.get("Content-Type", "") if e.headers else ""


class SmokeTest(unittest.TestCase):
    proc: subprocess.Popen | None = None

    @classmethod
    def setUpClass(cls) -> None:
        env = os.environ.copy()
        env.setdefault("PYTHONUNBUFFERED", "1")
        cls.proc = subprocess.Popen(
            [sys.executable, "app.py"],
            cwd=str(PROJECT_ROOT),
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
        )
        if not _wait_for_port(HOST, PORT, timeout=20):
            out = b""
            try:
                out = cls.proc.stdout.read(2000) if cls.proc.stdout else b""
            except Exception:
                pass
            cls.tearDownClass()
            raise RuntimeError(f"Server did not start on {BASE}\n{out.decode(errors='replace')}")

    @classmethod
    def tearDownClass(cls) -> None:
        if cls.proc and cls.proc.poll() is None:
            cls.proc.terminate()
            try:
                cls.proc.wait(timeout=5)
            except subprocess.TimeoutExpired:
                cls.proc.kill()

    # ---- Tests ----

    def test_index_loads_with_redesign_markers(self) -> None:
        status, body, ctype = _http_get("/")
        self.assertEqual(status, 200)
        self.assertIn("text/html", ctype)
        html = body.decode("utf-8", errors="replace")
        for marker in [
            "/static/styles-base.css",
            "/static/tokens.css",
            "/static/app.js",
            "/static/js/toast.js",
            "/static/js/validation.js",
            "/static/js/registry-retry.js",
            'id="resultEmptyState"',
            'id="companySelect"',
            'id="convertBtn"',
            "window.__BANKS",
        ]:
            self.assertIn(marker, html, f"missing marker: {marker}")

    def test_static_assets_served(self) -> None:
        for path, ct_prefix in [
            ("/static/styles-base.css", "text/css"),
            ("/static/tokens.css", "text/css"),
            ("/static/app.js", "application/javascript"),
            ("/static/js/toast.js", "application/javascript"),
            ("/static/js/combobox.js", "application/javascript"),
            ("/static/js/registry-retry.js", "application/javascript"),
        ]:
            with self.subTest(path=path):
                status, body, ctype = _http_get(path)
                self.assertEqual(status, 200, f"{path} -> {status}")
                self.assertTrue(ctype.startswith(ct_prefix), f"{path} ctype={ctype}")
                self.assertGreater(len(body), 50, f"{path} body too short")

    def test_static_path_traversal_blocked(self) -> None:
        for path in ["/static/../app.py", "/static/.git/config", "/static//etc/passwd"]:
            with self.subTest(path=path):
                status, _, _ = _http_get(path)
                self.assertIn(status, (400, 404), f"{path} -> {status}")

    def test_api_endpoints_respond(self) -> None:
        for path in ["/api/settings", "/api/companies", "/api/bank-registry"]:
            with self.subTest(path=path):
                status, body, ctype = _http_get(path)
                self.assertEqual(status, 200, f"{path} -> {status}")
                self.assertIn("application/json", ctype)
                self.assertTrue(body.startswith(b"{") or body.startswith(b"["))

    def test_excel_template_downloads(self) -> None:
        status, body, ctype = _http_get("/template.xlsx")
        self.assertEqual(status, 200)
        self.assertIn("spreadsheetml", ctype)
        # XLSX is a zip; first bytes are "PK".
        self.assertEqual(body[:2], b"PK")


if __name__ == "__main__":
    unittest.main(verbosity=2)
