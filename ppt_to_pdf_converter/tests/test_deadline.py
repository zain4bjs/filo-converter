"""Request-deadline behaviour of ppt_app (ppt_to_pdf_converter/main.py).

Needs the converter's own dependencies (PyMuPDF, FastAPI, tesseract), so run
it in the image:

    docker run --rm -u root -v "$PWD:/src:ro" --entrypoint sh <image> -c \
      'pip install -q pytest && cd /src && python -m pytest -p no:cacheprovider ppt_to_pdf_converter/tests'
"""

from __future__ import annotations

import asyncio
import json
import subprocess
import sys
import time
from pathlib import Path
from types import SimpleNamespace

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import main  # noqa: E402
from filo_deadline import Deadline, DeadlineExceeded  # noqa: E402


class FakeRequest:
    def __init__(self, seconds: float):
        self.state = SimpleNamespace(deadline=Deadline(seconds))
        self.url = SimpleNamespace(path="/convert/test")

    async def is_disconnected(self) -> bool:
        return False


def _run(coro):
    return asyncio.run(coro)


def test_run_bounded_answers_at_the_deadline_and_keeps_the_slot_until_the_work_stops():
    async def scenario():
        slots = asyncio.Semaphore(1)
        request = FakeRequest(0.5)
        stopped = []

        def slow(deadline):
            # Like the page loops: check between units of work.
            for _ in range(40):
                time.sleep(0.05)
                if deadline.expired():
                    stopped.append(time.monotonic())
                    raise DeadlineExceeded()
            return "done"

        started = time.monotonic()
        with pytest.raises(DeadlineExceeded):
            await main._run_bounded(slots, request, slow, request.state.deadline)
        answered = time.monotonic() - started
        assert 0.45 <= answered < 0.8
        # The thread notices at its next check and gives the slot back.
        for _ in range(40):
            if not slots.locked():
                break
            await asyncio.sleep(0.05)
        assert not slots.locked()
        assert stopped

    _run(scenario())


def test_queued_request_gets_504_when_its_budget_runs_out():
    async def scenario():
        slots = asyncio.Semaphore(1)
        await slots.acquire()  # someone else's long conversion
        started = time.monotonic()
        with pytest.raises(DeadlineExceeded):
            await main._run_bounded(slots, FakeRequest(0.4), lambda: "never")
        assert time.monotonic() - started < 0.7

    _run(scenario())


def test_slot_wait_shorter_than_the_budget_is_still_503(monkeypatch):
    monkeypatch.setattr(main, "SLOT_WAIT_SECONDS", 0.3)

    async def scenario():
        slots = asyncio.Semaphore(1)
        await slots.acquire()
        with pytest.raises(main.HTTPException) as info:
            await main._run_bounded(slots, FakeRequest(30), lambda: "never")
        assert info.value.status_code == 503

    _run(scenario())


def test_fast_work_is_unaffected():
    async def scenario():
        slots = asyncio.Semaphore(1)
        assert await main._run_bounded(slots, FakeRequest(5), lambda x: x * 2, 21) == 42
        assert not slots.locked()

    _run(scenario())


def test_libreoffice_timeout_is_the_smaller_of_85_s_and_what_is_left(monkeypatch, tmp_path):
    seen = []

    def fake_process(command, timeout=None, tmp_dir=None):
        seen.append(timeout)
        raise subprocess.TimeoutExpired(command, timeout)

    monkeypatch.setattr(main, "_run_soffice_process", fake_process)
    monkeypatch.setattr(main, "_profile_template", lambda: None)
    monkeypatch.setattr(main, "get_soffice_path", lambda: "soffice")
    out = tmp_path / "out"
    out.mkdir()
    deadline = Deadline(10)
    with pytest.raises(DeadlineExceeded):
        main._run_soffice(tmp_path / "in.pptx", out, "pdf", None, deadline)
    assert 9 < seen[0] <= 10
    assert deadline.expired()

    (tmp_path / "lo_profile").rename(tmp_path / "old_profile")
    with pytest.raises(DeadlineExceeded):
        main._run_soffice(tmp_path / "in.pptx", out, "pdf", None, Deadline(500))
    assert seen[1] == main.CONVERSION_TIMEOUT_SECONDS == 85


def _pdf(path: Path, pages: int) -> Path:
    import fitz

    with fitz.open() as doc:
        for number in range(pages):
            doc.new_page().insert_text((72, 72), f"page {number}")
        doc.save(str(path))
    return path


def test_page_loops_stop_once_the_budget_is_spent(tmp_path):
    pdf = _pdf(tmp_path / "in.pdf", 5)
    spent = Deadline(95)
    spent.expire()
    with pytest.raises(DeadlineExceeded):
        main._build_presentation(pdf, tmp_path, False, False, 150, False, False, spent)
    with pytest.raises(DeadlineExceeded):
        main._build_workbook(pdf, tmp_path, True, False, spent)
    with pytest.raises(DeadlineExceeded):
        main._extract_text_layer(pdf, main.fitz.TEXTFLAGS_TEXT, False, "auto", spent)
    # With time left they still work.
    assert main._build_presentation(pdf, tmp_path, False, False, 150, False, False, Deadline(60)).exists()


def test_tesseract_is_killed_at_the_deadline():
    import random

    from PIL import Image

    random.seed(1)
    noise = Image.frombytes("L", (2500, 3500), bytes(random.getrandbits(8) for _ in range(2500 * 3500)))
    deadline = Deadline(0.5)
    started = time.monotonic()
    with pytest.raises(DeadlineExceeded):
        main._ocr_image(noise.convert("RGB"), "eng", deadline)
    # pytesseract first writes the page to a temp image (seconds for this
    # incompressible noise, well under one for a real page); tesseract itself
    # then gets only what is left of the budget.
    assert time.monotonic() - started < 6.0
    assert deadline.expired()


def test_tesseract_failure_is_still_ocr_failed(monkeypatch):
    # TesseractError subclasses RuntimeError: it must keep its own 500 "OCR failed".
    from PIL import Image

    def boom(*args, **kwargs):
        raise main.pytesseract.TesseractError(1, "bad image")

    monkeypatch.setattr(main.pytesseract, "image_to_string", boom)
    with pytest.raises(main.HTTPException) as info:
        main._ocr_image(Image.new("RGB", (10, 10)), "eng", Deadline(60))
    assert info.value.status_code == 500 and info.value.detail == "OCR failed"


# ----- through the ASGI app: middleware + 504 handler ------------------------


def _asgi_post(path: str, headers: dict[str, str], filename: str, content: bytes):
    boundary = "filodeadlinetest"
    body = (
        f"--{boundary}\r\nContent-Disposition: form-data; name=\"file\"; filename=\"{filename}\"\r\n"
        "Content-Type: application/octet-stream\r\n\r\n"
    ).encode() + content + f"\r\n--{boundary}--\r\n".encode()
    raw_headers = [
        (b"content-type", f"multipart/form-data; boundary={boundary}".encode()),
        (b"content-length", str(len(body)).encode()),
    ] + [(name.lower().encode(), value.encode()) for name, value in headers.items()]
    scope = {
        "type": "http",
        "asgi": {"version": "3.0"},
        "http_version": "1.1",
        "method": "POST",
        "scheme": "http",
        "path": path,
        "raw_path": path.encode(),
        "query_string": b"",
        "root_path": "",
        "headers": raw_headers,
        "client": ("127.0.0.1", 1234),
        "server": ("127.0.0.1", 8081),
    }
    sent = {"body": b""}
    chunks = [{"type": "http.request", "body": body, "more_body": False}]

    async def receive():
        if chunks:
            return chunks.pop(0)
        await asyncio.sleep(3600)

    async def send(message):
        if message["type"] == "http.response.start":
            sent["status"] = message["status"]
        elif message["type"] == "http.response.body":
            sent["body"] += message.get("body", b"")

    asyncio.run(main.app(scope, receive, send))
    return sent["status"], sent["body"]


def test_request_that_spent_its_budget_before_the_app_gets_a_clean_504(tmp_path):
    pdf = _pdf(tmp_path / "in.pdf", 3).read_bytes()
    now = time.time()
    # A proxy in front received it 96 s ago (e.g. queued behind other work).
    headers = {"X-Filo-Request-Msec": f"{now - 95.9:.3f}", "X-Filo-Request-Time": "0.100"}
    for path in ("/convert/pdf-to-ppt", "/convert/pdf-to-xls", "/convert/pdf-to-text", "/convert/pdf-to-doc"):
        status, body = _asgi_post(path, headers, "in.pdf", pdf)
        assert status == 504, path
        assert json.loads(body) == {"error": "Conversion took too long"}, path


def test_fresh_request_converts(tmp_path):
    pdf = _pdf(tmp_path / "in.pdf", 3).read_bytes()
    now = time.time()
    headers = {"X-Filo-Request-Msec": f"{now:.3f}", "X-Filo-Request-Time": "0.001"}
    status, body = _asgi_post("/convert/pdf-to-text", headers, "in.pdf", pdf)
    assert status == 200
    assert b"page 2" in body
