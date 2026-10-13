import io
import logging
import os
import shutil
import sys
import tempfile
import time
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from PIL import Image
from werkzeug.exceptions import HTTPException, RequestEntityTooLarge
from werkzeug.utils import secure_filename

try:
    from . import main as converter_main
except ImportError:
    import main as converter_main

# Shared "Made with Filo" watermark (repo root watermark/, /app/watermark in
# the image). Applied only when the request carries X-Filo-Watermark: 1.
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "watermark"))
import filo_watermark  # noqa: E402

# Per-request time budget (repo root deadline/, /app/deadline in the image):
# answers 504 before a typical 100 s upstream proxy limit. See filo_deadline.
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "deadline"))
import filo_deadline  # noqa: E402


DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def _env_int(name, default, minimum=1):
    try:
        return max(minimum, int(os.environ.get(name, str(default))))
    except Exception:
        return default


# Upload cap, shared with the other app (keep any fronting proxy's body
# limit in sync). Flask rejects anything larger with 413 before conversion starts.
MAX_UPLOAD_MB = _env_int("MAX_UPLOAD_MB", 30)

# Page cap: very long PDFs tie up a worker for many minutes. Checked with
# PyMuPDF before converting; over the cap -> 422 (non-retryable in the app).
# A PDF under the cap that still cannot convert within the request budget
# (REQUEST_DEADLINE_SECONDS, below) gets a 504 instead (see README, "Limits").
MAX_PDF_PAGES = _env_int("MAX_PDF_PAGES", 50)

# Total time a conversion request may take, counted from when a reverse proxy
# in front received it, or else from when this app did (filo_deadline): 95 s
# by default, under a typical 100 s upstream proxy limit.
# The conversion runs in a forked child that is killed (with everything it
# spawned) when the budget runs out, and the client gets 504.
REQUEST_DEADLINE_SECONDS = filo_deadline.budget_seconds()

# Rasterising a page allocates width x height pixels before any other check
# can run, and a PDF may declare a page hundreds of inches wide. The pipeline
# renders at up to 300 DPI (OCR and exact mode), so pages whose 300 DPI render
# would exceed this are refused with 422 up front. Same cap as the ppt app;
# 150 MP still fits an A0 page (~139 MP at 300 DPI).
MAX_RENDER_PIXELS = _env_int("MAX_RENDER_PIXELS", 150_000_000)
_RENDER_DPI = 300

# mode=exact renders every page to an image and embeds all of them, which is
# far heavier per page than the other modes, so it gets its own cap; it can
# only be lower than MAX_PDF_PAGES, never higher.
MAX_EXACT_PAGES = min(_env_int("MAX_EXACT_PAGES", 50), MAX_PDF_PAGES)

# Decompression-bomb guard for every Pillow image this process opens (page
# renders, OCR input, embedded-image optimisation). Pillow's own default is
# ~89.5M px, so this pins roughly the same limit explicitly.
Image.MAX_IMAGE_PIXELS = _env_int("MAX_IMAGE_PIXELS", 90_000_000)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024

# Full error details go to the server log only; clients get a short fixed
# message. Under gunicorn, reuse its error log handlers so these land in the
# same stream as gunicorn (stderr).
logger = logging.getLogger("pdf_to_doc_converter.api")
_gunicorn_logger = logging.getLogger("gunicorn.error")
if _gunicorn_logger.handlers:
    logger.handlers = _gunicorn_logger.handlers
    logger.setLevel(_gunicorn_logger.level)
    logger.propagate = False


class _ClientError(Exception):
    """A request problem with a fixed, client-safe message and status."""

    def __init__(self, message, status):
        super().__init__(message)
        self.message = message
        self.status = status


@app.errorhandler(RequestEntityTooLarge)
def _too_large(_exc):
    return jsonify({"error": f"File too large (max {MAX_UPLOAD_MB} MB)."}), 413


@app.errorhandler(HTTPException)
def _http_error(exc):
    # 404/405/etc. as short JSON instead of Werkzeug's HTML pages.
    return jsonify({"error": exc.name}), exc.code


@app.errorhandler(Exception)
def _unhandled_error(_exc):
    logger.exception("Unhandled error in pdf_to_doc_converter")
    return jsonify({"error": "Internal server error."}), 500


class _ContinuousMode:
    def get(self):
        return "Continuous Reconstruction"


def _api_max_profiles():
    try:
        return max(1, int(os.environ.get("PDF2DOCX_MAX_PROFILES", "2")))
    except Exception:
        return 2


def _new_conversion_job(pdf_path, docx_path):
    job = type("ConversionJob", (object,), {})()
    job.pdf_path = pdf_path
    job.docx_path = docx_path
    job.reconstruction_mode = _ContinuousMode()
    job.FORCE_STRONG_CONTINUOUS = True
    job.max_profiles = _api_max_profiles()
    # Bind the PDF2Doc methods the pipeline relies on onto the lightweight job
    # object so they can call each other via `self`.
    for method_name in (
        "_resolve_poppler_path",
        "_resolve_tesseract_cmd",
        "_iter_rendered_pages",
        "_preprocess_ocr_image",
        "_run_pdf2docx",
        "_run_ocr_conversion",
        "_run_exact_layout_conversion",
        "_run_auto_pipeline",
        "_convert_pdf_subset",
        "_convert_mixed_document",
    ):
        setattr(job, method_name, getattr(converter_main.PDF2Doc, method_name).__get__(job, type(job)))
    return job


def _select_mode(pdf_path, requested_mode):
    mode = (requested_mode or os.environ.get("PDF2DOCX_DEFAULT_MODE", "auto")).strip().lower()
    if mode not in {"auto", "editable", "exact"}:
        raise ValueError("Invalid mode. Use auto, editable, or exact.")
    return mode


def _check_pdf(pdf_path, mode):
    """Reject PDFs that cannot or should not be converted, before converting.

    Unreadable, password-protected, over MAX_PDF_PAGES (MAX_EXACT_PAGES in
    exact mode), or with a page too large to rasterise safely.
    """
    fitz = getattr(converter_main, "fitz", None)
    if fitz is None:
        return  # PyMuPDF unavailable: leave it to the converter, as before.
    page_cap = MAX_EXACT_PAGES if mode == "exact" else MAX_PDF_PAGES
    scale = (_RENDER_DPI / 72.0) ** 2
    try:
        with fitz.open(pdf_path) as doc:
            if doc.needs_pass:
                # The converter has no password handling; it would only fail
                # later (and bypass the checks below).
                raise _ClientError("PDF is password protected.", 400)
            if doc.page_count > page_cap:
                raise _ClientError(f"PDF has too many pages (max {page_cap}).", 422)
            for page in doc:
                rect = page.rect
                if abs(rect.width) * abs(rect.height) * scale > MAX_RENDER_PIXELS:
                    raise _ClientError("PDF page is too large.", 422)
    except _ClientError:
        raise
    except Exception:
        logger.warning("PyMuPDF could not open uploaded PDF", exc_info=True)
        raise _ClientError("Could not read the PDF file.", 422)


def _run_converter(pdf_path, docx_path, requested_mode=None):
    mode = _select_mode(pdf_path, requested_mode)
    job = _new_conversion_job(pdf_path, docx_path)
    pdf_name = Path(pdf_path).stem or "document"

    if mode == "exact":
        converter_main.PDF2Doc._run_exact_layout_conversion(job, pdf_name)
        report = {
            "profile_name": "exact_layout",
            "score": 1.0,
            "metrics": {
                "word_compatibility_sanitized": False,
                "mode": mode,
            },
        }
    elif mode == "auto":
        report = job._run_auto_pipeline(pdf_name)
    else:  # editable
        report = job._run_pdf2docx()
    report.setdefault("metrics", {})["mode"] = mode

    # Downscale/recompress oversized embedded images (pdf2docx extracts them at
    # full native resolution, bloating the file and disturbing layout).
    try:
        opt = converter_main.optimize_docx_images(docx_path)
        report["metrics"]["images_optimized"] = opt.get("optimized", 0)
        report["metrics"]["image_bytes_saved"] = opt.get("bytes_saved", 0)
    except Exception:
        pass
    return report


@app.get("/health")
def health():
    return jsonify({"status": "ok"})


def _convert_in_child(work_dir, pdf_path, docx_path, requested_mode):
    """Body of the forked conversion child (filo_deadline.run_in_child).

    Everything the converter writes to a temp location (pdf2docx profile
    variants, OCRmyPDF output and its workers' files) goes under the request's
    own work_dir, so when the child is killed at the deadline the request's
    cleanup removes it too instead of leaving it behind in /tmp. Only the
    forked child changes these settings; the gunicorn worker keeps its own.
    """
    scratch = os.path.join(work_dir, "tmp")
    os.makedirs(scratch, exist_ok=True)
    os.environ["TMPDIR"] = scratch
    tempfile.tempdir = scratch
    return _run_converter(pdf_path, docx_path, requested_mode)


def _deadline_response():
    return jsonify({"error": filo_deadline.MESSAGE}), 504


@app.post("/convert")
@app.post("/api/convert")
def convert_pdf_to_docx():
    # Before touching request.files: reading the upload counts toward the budget.
    deadline = filo_deadline.Deadline.from_headers(request.headers.get, REQUEST_DEADLINE_SECONDS)
    if deadline.expired():
        # Spent the whole budget queued for a worker: do not start.
        logger.warning("Request deadline already spent before conversion on %s", request.path)
        return _deadline_response()
    upload = request.files.get("file")
    if upload is None:
        return jsonify({"error": "Missing file field. Send multipart/form-data with key 'file'."}), 400
    if not upload.filename:
        return jsonify({"error": "No file selected."}), 400
    if not upload.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Only PDF files are supported."}), 400
    requested_mode = request.form.get("mode") or request.args.get("mode")
    try:
        mode = _select_mode(None, requested_mode)
    except ValueError:
        if requested_mode:
            return jsonify({"error": "Invalid mode. Use auto, editable, or exact."}), 400
        # A bad PDF2DOCX_DEFAULT_MODE is a server misconfiguration, not a bad
        # request: answer 500 (retryable), as before this check existed.
        logger.error("Invalid PDF2DOCX_DEFAULT_MODE=%r", os.environ.get("PDF2DOCX_DEFAULT_MODE"))
        return jsonify({"error": "Internal server error."}), 500

    work_dir = tempfile.mkdtemp(prefix="pdf_to_docx_api_")
    try:
        safe_name = secure_filename(upload.filename) or "document.pdf"
        stem = Path(safe_name).stem or "document"
        pdf_path = os.path.join(work_dir, "input.pdf")
        docx_path = os.path.join(work_dir, f"{stem}.docx")

        upload.save(pdf_path)
        _check_pdf(pdf_path, mode)
        started_at = time.monotonic()
        report = filo_deadline.run_in_child(
            _convert_in_child, (work_dir, pdf_path, docx_path, requested_mode), deadline
        )
        if filo_watermark.watermark_requested(request.headers.get(filo_watermark.HEADER)):
            # A failure here fails the request (500): never hand out an
            # unmarked file that should carry the mark.
            filo_watermark.watermark_docx(docx_path)
        elapsed = time.monotonic() - started_at

        with open(docx_path, "rb") as file_obj:
            docx_bytes = file_obj.read()

        response = send_file(
            io.BytesIO(docx_bytes),
            mimetype=DOCX_MIME,
            as_attachment=True,
            download_name=f"{stem}.docx",
        )
        response.headers["X-Conversion-Profile"] = report.get("profile_name", "")
        response.headers["X-Conversion-Score"] = f"{report.get('score', 0.0):.4f}"
        response.headers["X-Conversion-Seconds"] = f"{elapsed:.3f}"
        response.headers["X-Conversion-Mode"] = report.get("metrics", {}).get("mode", "")
        response.headers["X-Word-Compatible"] = str(
            bool(report.get("metrics", {}).get("word_compatibility_sanitized", False))
        ).lower()
        pipeline = report.get("pipeline")
        if pipeline:
            response.headers["X-Conversion-Page-Kinds"] = ",".join(
                f"{k}:{v}" for k, v in pipeline.get("page_kinds", {}).items()
            )
            response.headers["X-Conversion-OCR-Applied"] = str(bool(pipeline.get("ocr_applied"))).lower()
            response.headers["X-Conversion-Rasterized-Pages"] = ",".join(
                str(p) for p in pipeline.get("rasterized_pages", [])
            )
            response.headers["X-Conversion-Image-Fallback"] = str(bool(pipeline.get("image_fallback"))).lower()
        return response
    except _ClientError as exc:
        return jsonify({"error": exc.message}), exc.status
    except filo_deadline.DeadlineExceeded:
        logger.warning(
            "PDF to DOCX stopped at the request deadline (%.0f s) on %s",
            REQUEST_DEADLINE_SECONDS,
            request.path,
        )
        return _deadline_response()
    except filo_deadline.ChildFailed as exc:
        logger.error("PDF to DOCX conversion failed: %s", exc)
        return jsonify({"error": "Conversion failed."}), 500
    except Exception:
        # Never echo exception text/tracebacks to clients; log them instead.
        logger.exception("PDF to DOCX conversion failed")
        return jsonify({"error": "Conversion failed."}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=int(os.environ.get("PORT", "5000")), debug=False)
