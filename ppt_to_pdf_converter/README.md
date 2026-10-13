# Office ↔ PDF Converter

FastAPI service converting office documents to PDF (LibreOffice headless) and
PDF back to Word documents (LibreOffice headless), spreadsheets, decks, and
plain text (PyMuPDF / pdfplumber).

## Run Locally

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python -m uvicorn main:app --host 127.0.0.1 --port 8081 --reload
```

Interactive docs: http://localhost:8081/docs

## Endpoints

| Endpoint | Engine | Input → Output |
| --- | --- | --- |
| `POST /convert/ppt-to-pdf` | LibreOffice | ppt, pptx, pps, ppsx, odp, doc, docx → PDF |
| `POST /convert/xls-to-pdf` | LibreOffice (Calc filter) | xls, xlsx, xlsm, xlsb, ods, csv → PDF |
| `POST /convert/pdf-to-doc` | LibreOffice (`writer_pdf_import`) | PDF → docx (visual copy) |
| `POST /convert/pdf-to-xls` | pdfplumber + openpyxl | PDF tables → xlsx |
| `POST /convert/pdf-to-ppt` | PyMuPDF + python-pptx | PDF pages → pptx (image slides) |
| `POST /convert/pdf-to-text` | PyMuPDF | PDF text layer → txt |
| `GET /health` | — | liveness probe |

All conversion endpoints take a multipart `file` field and stream the result
back as a download.

```bash
curl -F "file=@sample.pptx" http://localhost:8081/convert/ppt-to-pdf --output out.pdf
curl -F "file=@sample.pdf"  http://localhost:8081/convert/pdf-to-text --output out.txt
```

### `pdf-to-text` options

| Field | Default | Effect |
| --- | --- | --- |
| `include_page_markers` | `true` | Insert `----- Page N -----` separators |
| `sort_reading_order` | `false` | Re-order blocks top-left → bottom-right. Fixes scrambled PDFs, but also indents to preserve horizontal layout and can interleave real multi-column text |
| `dehyphenate` | `false` | Rejoin words split across line breaks |
| `ocr` | `auto` | `auto` OCRs only pages with no text layer; `force` OCRs every page (when the embedded text layer is itself garbage); `off` never OCRs |
| `ocr_language` | `eng` | Tesseract language code, e.g. `eng+deu` or `ara`. Installed: `eng`, `deu`, `fra`, `spa`, `ita`, `tur`, `ara` |

OCR renders at 300 DPI and runs in a worker thread. It costs roughly **0.5s
per page**, so `auto` can be far slower than `off` on documents with many
image-only pages — pass `ocr=off` when you know the PDF is digital and want
raw speed. Blank pages are skipped without rasterising.

### `pdf-to-ppt` options

| Field | Default | Effect |
| --- | --- | --- |
| `grayscale` | `false` | Greyscale slides |
| `black_and_white` | `false` | 1-bit output; much smaller files for scanned text |
| `threshold` | `150` | Luminance cut-off (0–255) for `black_and_white` |
| `invert` | `false` | Invert colours |

### `pdf-to-doc` notes

Takes only a `file` field. The output reproduces the page as positioned text
frames and shapes, so it opens looking like the original — but the text does
not reflow, because a PDF carries no paragraph or table structure to recover.
Use it when the document is meant to be read or printed, not edited.

LibreOffice needs three corrections before its output matches the source, all
applied to `word/document.xml` after conversion:

* `--infilter=writer_pdf_import` is mandatory. Without it Draw claims the `.pdf`
  extension, `soffice` exits 0 and writes nothing at all.
* Page margins are zeroed. LibreOffice keeps a 1 cm margin and positions frames
  relative to it, which shifts the whole page down and right.
* Vector fills are exported as legacy VML, which renderers stack above
  DrawingML text frames whatever the document order says — hiding, say, a white
  heading drawn over a coloured banner. A fill is pushed into the behind-text
  layer only when it covers a frame the PDF painted later, so a scan laid over
  a hidden text layer still keeps that text covered.

Each conversion runs against a private LibreOffice profile: a second instance
started against a locked profile exits 0 without writing output, which would
otherwise fail concurrent requests at random.

### `pdf-to-xls` options

| Field | Default | Effect |
| --- | --- | --- |
| `include_images` | `true` | Render pictures and vector charts onto a `Visuals` sheet |

Tables from every page are merged into one continuous `Data` sheet, with
headers repeated on continuation pages written only once. When a PDF contains
no real tables, the workbook opens on `Visuals` instead of `Data`.

## Behaviour Notes

* **PDF → xlsx** only recovers genuine tabular data. Sparse detections (chart
  gridlines) and prose sliced by ruling lines are rejected; such pages fall
  back to plain text so nothing is dropped.
* **PDF → pptx** produces image slides — a faithful visual copy, not editable
  text. That is the reliable way to turn any PDF into a deck.
* **PDF → docx** is a visual copy too. Fonts missing from the host are
  substituted, so italics and glyph widths can drift; page size snaps to
  LibreOffice's nearest standard (A4 lands ~0.35 pt off).
* **Scanned / image-only PDFs** are handled by `pdf-to-text` via OCR. The
  `pdf-to-xls` endpoint has no OCR and still returns `422` for them.
* Corrupt or password-protected PDFs return `400`.
* An OCR request on a host without the tesseract binary returns `503`.
* Uploads are capped at `MAX_UPLOAD_MB` (30 MB); LibreOffice conversions time
  out after `LO_TIMEOUT_SECONDS` (85 s, `504`), and the whole soffice process
  group is killed.
* Every conversion request has a total budget, `REQUEST_DEADLINE_SECONDS`
  (95 s): upload, queueing for a slot and the conversion itself. When it runs
  out the route answers `504 {"error": "Conversion took too long"}`: a
  LibreOffice run gets min(`LO_TIMEOUT_SECONDS`, time left) and is killed,
  Tesseract runs get the time left as their timeout, and the page loops stop at
  the next page. See the top-level README, "Request deadline".

## Limits and Errors

All limits are environment variables. Empty values and unexpanded CI macros
(`$(VAR)`) fall back to the default.

| Variable | Default | Effect |
| --- | --- | --- |
| `MAX_UPLOAD_MB` | `30` | Larger uploads return `413`. Checked on `Content-Length` before the body is parsed, then per file while streaming to disk |
| `MAX_PDF_PAGES` | `50` | Longer PDFs return `422` before any work starts |
| `REQUEST_DEADLINE_SECONDS` | `95` | Total time per conversion request; past it the route returns `504` |
| `LO_TIMEOUT_SECONDS` | `85` | Longest single LibreOffice run (capped by what is left of the request budget) |
| `LO_MAX_CONCURRENCY` | `2` | LibreOffice runs and OCR jobs allowed at once; further requests queue |
| `MAX_IMAGE_PIXELS` | `90000000` | Pillow's decompression-bomb limit |
| `MAX_RENDER_PIXELS` | `150000000` | Largest page or region PyMuPDF will rasterise; above it returns `422`. Fits A0 at 300 DPI |

Every blocking step (soffice, PyMuPDF, pdfplumber, python-pptx, Tesseract)
runs in a worker thread, so a long conversion never stalls `/health` or other
requests. For `pdf-to-text`, only documents that actually need OCR wait for a
concurrency slot; digital PDFs are extracted straight away.

Errors are JSON with a short fixed message, for example
`{"error": "File too large"}`. Exception text, LibreOffice or Tesseract output,
and file paths are never returned; they are logged server-side instead.
Unexpected failures return `500 {"error": "Conversion failed"}`. Request
validation failures keep FastAPI's `422` status but return
`{"error": "Invalid request"}` rather than echoing the input.

| Status | Meaning |
| --- | --- |
| `400` | Unsupported extension, corrupt or password-protected PDF, bad option value |
| `413` | Upload over `MAX_UPLOAD_MB` |
| `422` | Too many pages, page too large to render, no extractable content, invalid form field |
| `500` | Conversion failed |
| `503` | OCR requested but tesseract is not installed |
| `504` | LibreOffice timed out |

## LibreOffice Hardening

Uploaded documents are untrusted, so every LibreOffice profile is seeded with
[`lo_profile/registrymodifications.xcu`](lo_profile/registrymodifications.xcu)
before LibreOffice first starts against it:

* macros disabled outright (`DisableMacrosExecution`, `MacroSecurityLevel=3`);
* links from untrusted documents blocked (`BlockUntrustedRefererLinks`), and
  external links never updated on load in Calc and Writer;
* no formula recalculation on load, so functions such as `WEBSERVICE()` are
  never evaluated (this pins LibreOffice's existing default);
* no JVM (the image ships no Java runtime).

The file cites the officecfg schema for each key. After building the shared
profile template, the service checks the values survived LibreOffice's rewrite
of that file and re-seeds it if not. If the settings file is missing, the
conversion fails instead of running unhardened. A document that relied on
remote linked content now renders without it; local content is unaffected.

All writable paths live under the system temp directory (`TMPDIR`, else
`/tmp`): uploads, outputs, the profile template and per-request profiles.
soffice gets `HOME` set to the temp directory when the inherited `HOME` is not
writable. That lets the container run as a non-root user with a read-only root
filesystem and a tmpfs on `/tmp`.

## Deployment

System packages:

* **LibreOffice** on `PATH` (or the macOS path in `SOFFICE_CANDIDATES`) — for
  the two `*-to-pdf` endpoints and `pdf-to-doc`.
* **tesseract-ocr** plus the language data you need — for `pdf-to-text` OCR.
  Without it the service still runs; only `ocr=auto|force` returns 503.

```dockerfile
RUN apt-get update && apt-get install -y libreoffice tesseract-ocr tesseract-ocr-eng
```

Everything else is a pure Python wheel. `main.py` imports `../deadline/filo_deadline.py`
and `../watermark/`, so keep the repository layout. Keep `lo_profile/` next to `main.py`
in the image; the service will not run LibreOffice without it.
