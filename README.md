# Filo conversion engine

Source of the conversion engine behind Filo, https://filodocs.app (AGPL-3.0). History not included.

Document conversion engine behind Filo (filodocs.app): PDF to Word/PowerPoint/Excel/Text with OCR, Office to PDF via LibreOffice, optional "Made with Filo" watermark.

This repository contains only the conversion code. Access control, accounts,
rate limiting and hosting configuration of the production service are not
part of it; put your own reverse proxy in front if you expose the engine.

## Layout

| Folder | Service | Stack |
|---|---|---|
| `pdf_to_doc_converter/` | PDF to Word, layout-preserving (port 5000) | Flask, gunicorn, pdf2docx, PyMuPDF, Tesseract |
| `ppt_to_pdf_converter/` | Office to PDF, PDF to Word/Excel/PowerPoint/Text (port 8081) | FastAPI, uvicorn, LibreOffice, PyMuPDF, pdfplumber, Tesseract |
| `watermark/` | Optional "Made with Filo" mark, imported by both services | PyMuPDF, python-docx, python-pptx, openpyxl, Pillow |
| `deadline/` | Per-request time budget (`filo_deadline.py`), imported by both services | standard library |
| `docker/start.sh` | Starts both services in the container | bash |

Each service folder has its own README with route details and behaviour notes.

## Build and run

```bash
docker buildx build --load -t filo-converter .
docker run --rm -p 5000:5000 -p 8081:8081 \
  --read-only --tmpfs /tmp:rw,size=2g --cap-drop=ALL \
  filo-converter
```

The image runs as the non-root user `app` (uid 10001) and writes only under
`/tmp` (`HOME=/tmp`), so `--read-only` with a tmpfs on `/tmp` works. The
container stops when either service exits; run it with a restart policy.
Images build for `linux/amd64` and `linux/arm64` (the LibreOffice download is
pinned per architecture).

| Port | Service | Variable |
|---|---|---|
| 5000 | `pdf_to_doc_converter` (`/convert`, `/api/convert`, `/health`) | `PDF_APP_PORT` |
| 8081 | `ppt_to_pdf_converter` (`/convert/<route>`, `/health`) | `OFFICE_APP_PORT` |

Both bind to `BIND_HOST` (default `0.0.0.0` inside the container). To serve
them on one port, route `/convert/` (with a trailing path) to 8081 and
everything else to 5000 in your proxy.

Without Docker, each service runs on its own: see the service READMEs
(LibreOffice and Tesseract must be installed on the host).

## Endpoints

All conversion endpoints take `multipart/form-data` with the upload in field
`file` (the original filename; its extension is checked). Success is HTTP 200
with the file bytes and a `Content-Disposition: attachment` filename. Errors
are short fixed JSON messages, `{"error": "..."}`; exception text is only
logged server-side.

Port 5000 (`pdf_to_doc_converter`):

| Method | Path | Fields | Returns |
|---|---|---|---|
| POST | `/convert`, `/api/convert` | `file` (.pdf); optional `mode` = `auto` (default) \| `editable` \| `exact` (form or query) | `.docx`, plus `X-Conversion-Profile`, `X-Conversion-Score`, `X-Conversion-Seconds`, `X-Conversion-Mode` headers |
| GET | `/health` | | `{"status":"ok"}` |

Port 8081 (`ppt_to_pdf_converter`):

| Method | Path | Fields | Returns |
|---|---|---|---|
| POST | `/convert/ppt-to-pdf` | `file` (.ppt .pptx .pps .ppsx .odp .doc .docx) | PDF |
| POST | `/convert/xls-to-pdf` | `file` (.xls .xlsx .xlsm .xlsb .ods .csv) | PDF |
| POST | `/convert/pdf-to-doc` | `file` (.pdf) | `.docx` (LibreOffice visual copy) |
| POST | `/convert/pdf-to-xls` | `file`; `include_images` (bool, true) | `.xlsx` |
| POST | `/convert/pdf-to-ppt` | `file`; `grayscale`, `black_and_white`, `invert` (bool, false); `threshold` (int, 150) | `.pptx` |
| POST | `/convert/pdf-to-text` | `file`; `include_page_markers` (true), `sort_reading_order` (false), `dehyphenate` (false), `ocr` = `auto` \| `force` \| `off`, `ocr_language` (default `eng`, e.g. `eng+deu`) | `.txt` (UTF-8) |
| GET | `/health` | | `{"status":"ok"}` |

Examples:

```bash
curl -fsS -F "file=@scan.pdf" -o scan.docx http://localhost:5000/convert
curl -fsS -F "file=@deck.pptx" -o deck.pdf http://localhost:8081/convert/ppt-to-pdf
curl -fsS -F "file=@scan.pdf" -F ocr=force -F ocr_language=eng \
  -o scan.txt http://localhost:8081/convert/pdf-to-text
```

OCR languages installed in the image: `eng`, `deu`, `fra`, `spa`, `ita`,
`tur`, `ara` (add `tesseract-ocr-<lang>` packages in the `Dockerfile` for
more). An `ocr_language` that is not installed gets 400.

LibreOffice and OCR work runs off the event loop under a semaphore
(`LO_MAX_CONCURRENCY`). Each LibreOffice run uses a fresh profile seeded from
`ppt_to_pdf_converter/lo_profile/registrymodifications.xcu`: macros disabled,
links from untrusted documents blocked, no external link updates and no
formula recalculation on load.

## Environment variables

Values that are empty or start with `$(` (an unexpanded CI macro) are treated
as unset.

| Variable | Default | Used by | Meaning |
|---|---|---|---|
| `MAX_UPLOAD_MB` | `30` | both | Upload cap; larger requests get 413. |
| `MAX_PDF_PAGES` | `50` | both | Page cap; longer PDFs get 422 before any conversion work. |
| `MAX_EXACT_PAGES` | `50` (capped at `MAX_PDF_PAGES`) | 5000 | Page cap for `mode=exact` (every page embedded as a 300 DPI image). |
| `REQUEST_DEADLINE_SECONDS` | `95` | both | Total time budget of one conversion request; past it the route answers 504 (see "Request deadline"). |
| `LO_TIMEOUT_SECONDS` | `85` | 8081 | A LibreOffice run (Office to PDF, pdf-to-doc) is killed after this, or after what is left of the request budget if that is less (whole process group), and the client gets 504. |
| `LO_MAX_CONCURRENCY` | `2` | 8081 | Concurrent LibreOffice runs and OCR jobs; further requests queue. |
| `PDF_MAX_CONCURRENCY` | `2` | 8081 | Concurrent pdf-to-ppt / pdf-to-xls / pdf-to-text extraction jobs. |
| `SLOT_WAIT_SECONDS` | `120` | 8081 | Longest wait for one of the slots above before a retryable 503. |
| `MAX_UNPACKED_MB` | `1000` | 8081 | OOXML/ODF (zip) uploads that declare a larger unpacked size get 422. |
| `MAX_RENDER_PIXELS` | `150000000` | both | Pages whose render would exceed this many pixels get 422 (fits A0 at 300 DPI). |
| `MAX_IMAGE_PIXELS` | `90000000` | both | Pillow's decompression-bomb limit. |
| `OMP_THREAD_LIMIT` | `1` (image) | tesseract | One OpenMP thread per tesseract process. |
| `PDF2DOCX_DEFAULT_MODE` | `auto` | 5000 | Mode when the request sends none. |
| `PDF2DOCX_MAX_PROFILES` | `1` (image) | 5000 | Reconstruction profiles tried per document in `auto`. |
| `WEB_CONCURRENCY`, `GUNICORN_*` | see `Dockerfile` | 5000 | gunicorn workers, threads, timeouts. |
| `BIND_HOST`, `PDF_APP_PORT`, `OFFICE_APP_PORT` | `0.0.0.0`, `5000`, `8081` | `docker/start.sh` | Listen addresses. |

With the defaults, the slowest conversions (OCR of a 50-page scan, LibreOffice
PDF to Word on long text PDFs) finish in about a minute on 4 cores, under the
100 s request timeout many proxies and CDNs apply.

## Request deadline

Many reverse proxies and CDNs cut a request at 100 s, after which any work
still running is wasted. So every conversion route has one total budget,
`REQUEST_DEADLINE_SECONDS` (default 95 s), covering the upload, queueing for a
worker or slot, and the conversion itself. When it runs out the route stops
the work and answers `504 {"error": "Conversion took too long"}`:

- port 5000 runs each conversion in a forked child process and kills it (and
  everything it started, e.g. OCR workers) at the deadline;
- port 8081 gives each LibreOffice run min(`LO_TIMEOUT_SECONDS`, time left)
  and Tesseract runs the time left, and its page loops stop at the next page.

By default the budget starts when the app sees the request. A reverse proxy in
front can make it count from the moment the proxy received the request by
setting two headers on every request it forwards (replacing any values the
client sent): `X-Filo-Request-Msec`, its clock in seconds since the epoch when
it forwarded the request, and `X-Filo-Request-Time`, how many seconds the
request had already spent with it (nginx: `$msec` and `$request_time`). Keep
the proxy's own read timeout and `GUNICORN_TIMEOUT` above the budget, so the
app's clean 504 arrives first. A PDF under `MAX_PDF_PAGES` that still cannot
finish in time gets the 504.

## Watermark

When a fronting proxy sets `X-Filo-Watermark: 1` the output is marked. The
header value must be exactly `1`; absent or any other value means no mark.
The engine trusts the header as given, so a proxy that decides it must
overwrite whatever the client sent.

| Route | Output | Mark |
|---|---|---|
| `/convert`, `/api/convert`, `/convert/pdf-to-doc` | DOCX | every footer: wordmark image + "Made with Filo · filodocs.app" |
| `/convert/ppt-to-pdf`, `/convert/xls-to-pdf` | PDF | every page, bottom right: wordmark + "Made with Filo" + "filodocs.app" (60 % opacity), about 14 % of the page width (56 to 150 pt) |
| `/convert/pdf-to-xls` | XLSX | every sheet's footer: "Made with Filo · filodocs.app" |
| `/convert/pdf-to-ppt` | PPTX | every slide, bottom right: wordmark picture + two-line text box |
| `/convert/pdf-to-text` | TXT | never |

If marking fails the request fails (500): a file that should carry the mark is
never sent without it. The images are `watermark/assets/watermark-filo-mark.png`
(halo and 60 % opacity baked in) and `watermark-filo.png`, built by
`python watermark/make_assets.py` from `watermark/assets/logo-source.png`.
The Filo name and logo identify the Filo service; if you run a modified
version, do not present its output as Filo's (leave the header unset).

## LibreOffice

The image installs The Document Foundation's own LibreOffice build
(**26.8.1**) instead of Debian bookworm's 7.4 package, because LibreOffice
parses every uploaded Office file and must be a release upstream still
patches. The tarball is pinned by version and SHA-256 per architecture.
Installed parts: Writer, Calc, Impress, Draw with PDF import, Math, the
graphic and XSLT filters, the icon theme and LibreOffice's own fonts. Not
installed: the script providers, Base, extensions, desktop integration and
the updater. There is no Java runtime in the image.

To upgrade:

1. Pick the version from https://download.documentfoundation.org/libreoffice/stable/
   and note its full build number (for example `26.8.1.1`).
2. Download `LibreOffice_<v>_Linux_x86-64_deb.tar.gz` and
   `LibreOffice_<v>_Linux_aarch64_deb.tar.gz` with their `.asc` files and check
   both with `gpg --verify` against the LibreOffice Build Team key
   `C2839ECAD9408FBE9531C3E9F434A1EFAFEEAEA3`.
3. Put the `sha256sum` of each tarball into `LO_SHA256_AMD64` /
   `LO_SHA256_ARM64` in the `Dockerfile`, with `LO_VERSION` and `LO_BUILD`.
4. Build, then check `soffice --version` and that the logs do not say
   "LibreOffice dropped hardened settings from the template" (a renamed or
   removed profile key; fix `registrymodifications.xcu`).

## Tests

```bash
python -m pytest deadline/tests watermark/tests
```

`deadline/tests` needs only pytest; `watermark/tests` needs PyMuPDF,
python-docx, python-pptx, openpyxl and Pillow (the pinned versions in the two
`requirements.txt` files). The service tests need the full runtime
(LibreOffice, Tesseract), so run every suite in the image, one pytest run per
suite (two of them share a test file name); no network:

```bash
docker run --rm -u root -v "$PWD:/src:ro" --entrypoint sh filo-converter -c \
  'pip install -q pytest && cd /src && for s in deadline watermark pdf_to_doc_converter ppt_to_pdf_converter; do \
   python -m pytest -q -p no:cacheprovider "$s/tests" || exit 1; done'
```

## License

This repository is licensed under the **GNU Affero General Public License
v3.0** (`LICENSE`), chiefly because it links PyMuPDF (AGPL-3.0). If you run a
modified version as a network service, you must offer its users the
corresponding source.

`pdf_to_doc_converter/` is derived from "PDF-Converter" by Süleyman Bozkurt
(MIT, kept in `pdf_to_doc_converter/LICENSE`). Third-party components and
their licenses are listed in `NOTICE`.

Contact: support@filodocs.app
