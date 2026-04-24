from __future__ import annotations

import contextlib
import os
import socket
import threading
import time
import uuid

import pytest
from werkzeug.serving import make_server

from app.extensions import db
from app.models import User

try:
    from selenium import webdriver
    from selenium.common.exceptions import WebDriverException
    from selenium.webdriver.chrome.options import Options as ChromeOptions
except ModuleNotFoundError:  # pragma: no cover
    webdriver = None  # type: ignore[assignment]
    WebDriverException = Exception  # type: ignore[assignment, misc]
    ChromeOptions = None  # type: ignore[assignment]


def _get_free_port() -> int:
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


@pytest.fixture
def e2e_base_url(app) -> str:
    """Serves the Flask app in-process (thread) so TestConfig in-memory DB is shared."""
    port = _get_free_port()
    # In-memory SQLite + shared connection (TestConfig / StaticPool) is not
    # safe for concurrent use on one connection; one worker thread is enough
    # for E2E and prevents sqlite3 "bad parameter or other API misuse".
    srv = make_server("127.0.0.1", port, app, threaded=False, processes=1)
    thread = threading.Thread(target=srv.serve_forever, daemon=True)
    thread.start()
    deadline = time.time() + 8.0
    while time.time() < deadline:
        with contextlib.closing(socket.socket(socket.AF_INET, socket.SOCK_STREAM)) as sock:
            try:
                sock.settimeout(0.2)
                sock.connect(("127.0.0.1", port))
                break
            except OSError:
                time.sleep(0.05)
    else:
        srv.shutdown()
        thread.join(timeout=2.0)
        pytest.fail("E2E server did not start listening in time")
    try:
        yield f"http://127.0.0.1:{port}"
    finally:
        try:
            srv.shutdown()
        except (OSError, Exception):
            pass
        thread.join(timeout=5.0)


@pytest.fixture
def e2e_user(app):
    """A known user+password in the in-memory test DB (same app as e2e_base_url)."""
    with app.app_context():
        email = f"e2e_{uuid.uuid4().hex[:10]}@student.uwa.edu.au"
        u = User(
            email=email,
            full_name="E2E User",
            student_id="123",
        )
        u.set_password("E2E_pass_99")
        db.session.add(u)
        db.session.commit()
        return {"email": email, "password": "E2E_pass_99"}


def _env_headless() -> list[str]:
    o = [os.environ.get("SELENIUM_HEADLESS", "1")]
    if o[0] not in ("0", "false", "False", "no"):
        return [
            "--headless=new",
            "--no-sandbox",
            "--disable-gpu",
            "--disable-dev-shm-usage",
        ]
    return []


@pytest.fixture
def e2e_driver(e2e_base_url):
    """Chromium/Chrome; Selenium Manager fetches a matching driver (Selenium 4.6+)."""
    if webdriver is None or ChromeOptions is None:  # pragma: no cover
        pytest.skip("selenium is not installed")

    options = ChromeOptions()
    for arg in _env_headless():
        options.add_argument(arg)
    if os.environ.get("CI", "").lower() in ("1", "true", "yes"):
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
    options.set_capability("pageLoadStrategy", "eager")

    try:
        dr = webdriver.Chrome(options=options)
    except (WebDriverException, OSError) as exc:  # pragma: no cover
        pytest.skip(f"Chrome/ChromeDriver not available: {exc!s}")

    dr.set_page_load_timeout(30)
    with contextlib.suppress(Exception):
        dr.set_window_size(1280, 900)
    try:
        yield dr
    finally:
        with contextlib.suppress(Exception):
            dr.quit()
