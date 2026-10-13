"""Unit tests for deadline/filo_deadline.py (standard library + pytest only).

    python -m pytest deadline/tests
"""

from __future__ import annotations

import os
import re
import sys
import time
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "deadline"))

import filo_deadline as fd  # noqa: E402


class FakeClock:
    def __init__(self, now: float = 1000.0):
        self.now = now

    def __call__(self) -> float:
        return self.now


# ----- settings --------------------------------------------------------------


def test_budget_default_and_overrides():
    assert fd.budget_seconds({}) == 95.0
    assert fd.budget_seconds({"REQUEST_DEADLINE_SECONDS": "60"}) == 60.0
    assert fd.budget_seconds({"REQUEST_DEADLINE_SECONDS": " 12.5 "}) == 12.5
    for bad in ("", "  ", "$(REQUEST_DEADLINE_SECONDS)", "abc", "0", "-5", "nan", "inf"):
        assert fd.budget_seconds({"REQUEST_DEADLINE_SECONDS": bad}) == 95.0, bad


# ----- time spent before the app ---------------------------------------------


def test_elapsed_before_app_from_proxy_headers():
    # The proxy received the request at 100.0, forwarded it at 100.5 (0.5 s in
    # the proxy), the app sees it at 130.5 (30 s in a busy app's listen backlog).
    assert fd.elapsed_before_app("100.500", "0.500", 130.5) == pytest.approx(30.5)
    # Only the elapsed header (no clock): counts what the proxy itself measured.
    assert fd.elapsed_before_app(None, "2.25", 50.0) == pytest.approx(2.25)
    # Missing / nonsense / future values never make the budget larger.
    assert fd.elapsed_before_app(None, None, 50.0) == 0.0
    assert fd.elapsed_before_app("garbage", "x", 50.0) == 0.0
    assert fd.elapsed_before_app("200.0", "0.1", 50.0) == 0.0  # clock step backwards
    assert fd.elapsed_before_app("nan", "0", 50.0) == 0.0


# ----- Deadline --------------------------------------------------------------


def test_deadline_counts_down_and_expires():
    clock = FakeClock()
    deadline = fd.Deadline(95, clock=clock)
    assert deadline.remaining() == pytest.approx(95)
    clock.now += 60
    assert deadline.remaining() == pytest.approx(35)
    assert deadline.elapsed() == pytest.approx(60)
    deadline.check()
    clock.now += 35
    assert deadline.expired()
    with pytest.raises(fd.DeadlineExceeded) as info:
        deadline.check()
    assert info.value.message == "Conversion took too long"


def test_already_elapsed_is_clamped_to_the_budget():
    clock = FakeClock()
    assert fd.Deadline(95, already_elapsed=40, clock=clock).remaining() == pytest.approx(55)
    assert fd.Deadline(95, already_elapsed=500, clock=clock).expired()
    assert fd.Deadline(95, already_elapsed=-10, clock=clock).remaining() == pytest.approx(95)


def test_from_headers_counts_proxy_and_backlog_time():
    clock = FakeClock()
    headers = {"X-Filo-Request-Msec": "1700000000.000", "X-Filo-Request-Time": "0.010"}
    deadline = fd.Deadline.from_headers(headers.get, 95, clock=clock, wall=lambda: 1700000020.0)
    assert deadline.remaining() == pytest.approx(95 - 20.01)
    # No headers (app reached directly): the budget starts now.
    assert fd.Deadline.from_headers({}.get, 95, clock=clock).remaining() == pytest.approx(95)


def test_expire_stops_work_at_once():
    deadline = fd.Deadline(95)
    deadline.expire()
    assert deadline.remaining() == 0.0
    with pytest.raises(fd.DeadlineExceeded):
        deadline.check()


def test_timeout_is_capped_and_never_zero():
    clock = FakeClock()
    deadline = fd.Deadline(95, clock=clock)
    assert deadline.timeout(85) == pytest.approx(85)  # LibreOffice: min(85, left)
    clock.now += 30
    assert deadline.timeout(85) == pytest.approx(65)
    assert deadline.timeout() == pytest.approx(65)
    clock.now += 65
    # pytesseract reads timeout=0 as "no timeout": raise instead.
    with pytest.raises(fd.DeadlineExceeded):
        deadline.timeout(85)


# ----- run_in_child ----------------------------------------------------------

posix_only = pytest.mark.skipif(not hasattr(os, "fork"), reason="needs fork")


def _add(a, b):
    return {"sum": a + b, "pid": os.getpid()}


def _boom():
    raise ValueError("bad input")


def _die():
    os._exit(3)


def _sleep(seconds):
    time.sleep(seconds)
    return "late"


def _spawn_grandchild_then_hang(pid_file):
    import subprocess

    child = subprocess.Popen(["sleep", "60"])
    Path(pid_file).write_text(str(child.pid))
    time.sleep(60)


def _alive(pid: int) -> bool:
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    # A zombie still "exists": check its state where /proc is available.
    stat = Path(f"/proc/{pid}/stat")
    if stat.exists():
        return stat.read_text().split(")")[-1].split()[0] != "Z"
    return True


@posix_only
def test_run_in_child_returns_the_result():
    result = fd.run_in_child(_add, (2, 3), fd.Deadline(10))
    assert result["sum"] == 5
    assert result["pid"] != os.getpid()


@posix_only
def test_run_in_child_reports_failures():
    with pytest.raises(fd.ChildFailed, match="ValueError: bad input"):
        fd.run_in_child(_boom, (), fd.Deadline(10))
    with pytest.raises(fd.ChildFailed, match="code 3"):
        fd.run_in_child(_die, (), fd.Deadline(10))


@posix_only
def test_run_in_child_stops_at_the_deadline():
    deadline = fd.Deadline(1.0)
    started = time.monotonic()
    with pytest.raises(fd.DeadlineExceeded):
        fd.run_in_child(_sleep, (30,), deadline)
    took = time.monotonic() - started
    assert 0.9 <= took < 3.0
    assert deadline.expired()


@posix_only
def test_run_in_child_kills_the_whole_process_group(tmp_path):
    pid_file = tmp_path / "grandchild.pid"
    with pytest.raises(fd.DeadlineExceeded):
        fd.run_in_child(_spawn_grandchild_then_hang, (str(pid_file),), fd.Deadline(1.5))
    grandchild = int(pid_file.read_text())
    for _ in range(50):
        if not _alive(grandchild):
            break
        time.sleep(0.1)
    assert not _alive(grandchild), "a process spawned by the conversion outlived the deadline"


@posix_only
def test_run_in_child_with_no_time_left_does_not_start():
    deadline = fd.Deadline(95)
    deadline.expire()
    with pytest.raises(fd.DeadlineExceeded):
        fd.run_in_child(_add, (1, 1), deadline)


# ----- container config: nothing may cut a request before the app does -------


def _dockerfile_env(name: str) -> str:
    match = re.search(rf"^\s*(?:ENV\s+)?{name}=(\S+)", (REPO / "Dockerfile").read_text(), re.M)
    assert match, name
    return match.group(1)


def test_image_defaults():
    assert _dockerfile_env("MAX_PDF_PAGES") == "50"
    assert _dockerfile_env("REQUEST_DEADLINE_SECONDS") == "95"
    assert _dockerfile_env("LO_TIMEOUT_SECONDS") == "85"
    assert _dockerfile_env("LO_MAX_CONCURRENCY") == "2"
    assert "COPY deadline/ /app/deadline/" in (REPO / "Dockerfile").read_text()


def test_budget_is_under_a_100_s_proxy_limit_and_gunicorn_waits_longer():
    budget = float(_dockerfile_env("REQUEST_DEADLINE_SECONDS"))
    assert budget < 100  # many reverse proxies and CDNs cut a request at 100 s
    gunicorn = (REPO / "pdf_to_doc_converter" / "gunicorn_config.py").read_text()
    match = re.search(r'GUNICORN_TIMEOUT", "(\d+)"', gunicorn)
    assert match and int(match.group(1)) >= budget


def test_deadline_header_names():
    assert fd.NGINX_MSEC_HEADER == "X-Filo-Request-Msec"
    assert fd.NGINX_ELAPSED_HEADER == "X-Filo-Request-Time"
