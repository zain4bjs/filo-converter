"""Request-deadline behaviour of pdf_app (pdf_to_doc_converter/api.py).

Needs the converter's dependencies, so run it in the image (see
ppt_to_pdf_converter/tests/test_deadline.py for the command).
"""

from __future__ import annotations

import io
import json
import os
import sys
import time
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import api  # noqa: E402


def _pdf_bytes(pages: int = 2) -> bytes:
    import fitz

    with fitz.open() as doc:
        for number in range(pages):
            doc.new_page().insert_text((72, 72), f"page {number}")
        return doc.tobytes()


@pytest.fixture()
def client():
    api.app.config["TESTING"] = True
    return api.app.test_client()


def _post(client, headers=None):
    return client.post(
        "/convert",
        data={"file": (io.BytesIO(_pdf_bytes()), "in.pdf")},
        headers=headers or {},
        content_type="multipart/form-data",
    )


def test_converts_within_budget(client):
    response = _post(client)
    assert response.status_code == 200
    assert response.headers["Content-Type"].startswith(api.DOCX_MIME)


def test_budget_spent_in_the_queue_answers_504_without_converting(client, monkeypatch):
    calls = []
    monkeypatch.setattr(api, "_run_converter", lambda *a: calls.append(a))
    now = time.time()
    response = _post(
        client, {"X-Filo-Request-Msec": f"{now - 96:.3f}", "X-Filo-Request-Time": "0.050"}
    )
    assert response.status_code == 504
    assert json.loads(response.data) == {"error": "Conversion took too long"}
    assert calls == []


def _slow_converter(pdf_path, docx_path, requested_mode=None):
    Path(pdf_path).with_name("child.pid").write_text(str(os.getpid()))
    time.sleep(60)


def test_slow_conversion_is_killed_and_answered_504_at_the_deadline(client, monkeypatch, tmp_path):
    monkeypatch.setattr(api, "REQUEST_DEADLINE_SECONDS", 2.0)
    monkeypatch.setattr(api, "_run_converter", _slow_converter)
    monkeypatch.setattr(api.tempfile, "mkdtemp", lambda prefix="": str(tmp_path / "work"))
    (tmp_path / "work").mkdir()
    started = time.monotonic()
    response = _post(client)
    took = time.monotonic() - started
    assert response.status_code == 504
    assert json.loads(response.data) == {"error": "Conversion took too long"}
    assert 1.8 <= took < 4.0
    # Work dir removed, and the child really is gone.
    assert not (tmp_path / "work").exists()


def test_converter_crash_is_a_500(client, monkeypatch):
    def crash(*_args):
        raise RuntimeError("boom")

    monkeypatch.setattr(api, "_run_converter", crash)
    response = _post(client)
    assert response.status_code == 500
    assert json.loads(response.data) == {"error": "Conversion failed."}
