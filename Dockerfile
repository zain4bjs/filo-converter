# syntax=docker/dockerfile:1
# Filo conversion engine: two HTTP services in one image.
#   - pdf_to_doc_converter  (Flask, gunicorn)  port 5000
#       POST /convert, /api/convert  PDF to Word (pdf2docx + OCR)
#   - ppt_to_pdf_converter  (FastAPI, uvicorn) port 8081
#       POST /convert/<route>        Office to PDF, PDF to Word/Excel/PowerPoint/Text
# Both import watermark/ and deadline/ from /app.
# docker/start.sh runs both; the container stops when either exits, so run it
# with a restart policy. Everything writable lives under /tmp (HOME=/tmp), so
# the container can run with --read-only and a tmpfs on /tmp.
#
# Python stays on 3.13 so the pinned wheels (numpy 2.5, PyMuPDF 1.27.x) install.

# ---- LibreOffice, official TDF build (see README, "LibreOffice") ----
# Debian bookworm ships LibreOffice 7.4 (upstream end of life 2023-06); the
# converter opens untrusted Office files with it, so the image installs the
# Document Foundation's own .deb build instead: a release upstream still
# patches, in /opt/libreoffice<series>, with its own copies of the document
# libraries, so nothing else in the image changes. The tarball is pinned by
# version and SHA-256 per architecture (each checked against TDF's GPG
# signature, key C2839ECAD9408FBE9531C3E9F434A1EFAFEEAEA3, when pinned). The
# download tries the current-release path first, then TDF's permanent archive
# (the same file; a release leaves stable/ once newer bugfix releases ship).
# To upgrade: bump LO_VERSION, LO_BUILD and both checksums together.
FROM python:3.13-slim-bookworm AS lo-fetch
ARG TARGETARCH
ARG LO_VERSION=26.8.1
ARG LO_BUILD=26.8.1.1
ARG LO_SHA256_AMD64=30903df3b9f61360d9660cd707de48cd2831469114492a5008ed58a0ac77d044
ARG LO_SHA256_ARM64=1b069c20dd237f6decad3ea02cb02f39fdf458b03d09c2a9facb7b9b71e6a27a
RUN set -eu; \
    case "$TARGETARCH" in \
      amd64) dir=x86_64; arch=x86-64; sum="$LO_SHA256_AMD64" ;; \
      arm64) dir=aarch64; arch=aarch64; sum="$LO_SHA256_ARM64" ;; \
      *) echo "no pinned LibreOffice build for $TARGETARCH" >&2; exit 1 ;; \
    esac; \
    stable="https://download.documentfoundation.org/libreoffice/stable/${LO_VERSION}/deb/${dir}/LibreOffice_${LO_VERSION}_Linux_${arch}_deb.tar.gz"; \
    archive="https://downloadarchive.documentfoundation.org/libreoffice/old/${LO_BUILD}/deb/${dir}/LibreOffice_${LO_BUILD}_Linux_${arch}_deb.tar.gz"; \
    # (socket timeout: a stalled mirror fails the attempt instead of hanging the build)
    fetch() { python -c "import socket, sys, urllib.request; socket.setdefaulttimeout(60); urllib.request.urlretrieve(sys.argv[1], sys.argv[2])" "$1" /tmp/lo.tar.gz \
      && echo "$sum  /tmp/lo.tar.gz" | sha256sum -c -; }; \
    ok=; \
    for attempt in 1 2 3; do \
      for url in "$stable" "$archive"; do \
        if fetch "$url"; then ok=1; break 2; fi; \
        echo "LibreOffice download from $url failed (attempt $attempt)"; rm -f /tmp/lo.tar.gz; \
      done; \
      sleep 15; \
    done; \
    test -n "$ok"; \
    mkdir -p /lo-debs; \
    tar -xzf /tmp/lo.tar.gz -C /lo-debs --strip-components=2 --wildcards '*/DEBS/*.deb'; \
    rm /tmp/lo.tar.gz; \
    ls /lo-debs | grep -q '^libobasis'

FROM python:3.13-slim-bookworm AS runtime

ENV PIP_NO_CACHE_DIR=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

WORKDIR /app

RUN --mount=type=cache,target=/var/cache/apt,sharing=locked \
    --mount=type=cache,target=/var/lib/apt/lists,sharing=locked \
    apt-get update && apt-get install -y --no-install-recommends \
    poppler-utils \
    tesseract-ocr \
    tesseract-ocr-deu \
    tesseract-ocr-fra \
    tesseract-ocr-spa \
    tesseract-ocr-ita \
    tesseract-ocr-tur \
    tesseract-ocr-ara \
    ghostscript \
    libglib2.0-0 \
    libgl1 \
    libfreetype6 \
    fontconfig \
    fonts-liberation \
    fonts-dejavu \
    fonts-dejavu-core \
    fonts-crosextra-carlito \
    fonts-crosextra-caladea \
    fonts-freefont-ttf \
    fonts-noto \
    fonts-noto-cjk \
    fonts-noto-color-emoji \
    # --- LibreOffice (TDF build, installed below) runtime libraries ---
    # What soffice.bin and its core modules link against (checked with ldd);
    # the GTK/Qt/GStreamer/Java plugins it also ships are never loaded headless.
    libcairo2 \
    libcups2 \
    libdbus-1-3 \
    libgssapi-krb5-2 \
    libharfbuzz0b \
    libice6 \
    libnspr4 \
    libnss3 \
    libsm6 \
    libx11-xcb1 \
    libxext6 \
    libxinerama1 \
    libxrandr2 \
    libxrender1 \
    && rm -rf /var/lib/apt/lists/* \
    && fc-cache -f


# LibreOffice from the pinned TDF tarball (lo-fetch stage). Only the parts a
# headless conversion uses: Writer, Calc, Impress, Draw (+ PDF import, used by
# /convert/pdf-to-doc), Math (formula objects in documents), the graphic and
# XSLT filters, plus the icon theme and LibreOffice's own font set, which the
# core packages require (those fonts stay private to /opt/libreoffice*; the
# system fontconfig set above is unchanged). Left out on purpose: the Python,
# JavaScript and BeanShell script providers, Base and its database drivers,
# extensions, desktop integration and the online updater. "soffice" on PATH
# is what ppt_to_pdf_converter looks for.
RUN --mount=type=bind,from=lo-fetch,source=/lo-debs,target=/lo-debs \
    set -eu; cd /lo-debs; \
    dpkg -i \
      libreoffice*-ure_*.deb \
      libreoffice[0-9][0-9].[0-9]_*.deb \
      libobasis*-core_*.deb \
      libobasis*-images_*.deb \
      libobasis*-ooofonts_*.deb \
      libobasis*-en-us_*.deb libreoffice*-en-us_*.deb \
      libobasis*-writer_*.deb libreoffice*-writer_*.deb \
      libobasis*-calc_*.deb libreoffice*-calc_*.deb \
      libobasis*-impress_*.deb libreoffice*-impress_*.deb \
      libobasis*-draw_*.deb libreoffice*-draw_*.deb \
      libobasis*-math_*.deb libreoffice*-math_*.deb \
      libobasis*-graphicfilter_*.deb \
      libobasis*-xsltfilter_*.deb \
      libobasis*-extension-pdf-import_*.deb \
      libobasis*-ooolinguistic_*.deb; \
    lo_dir="$(ls -d /opt/libreoffice*/program)"; \
    ln -s "$lo_dir/soffice" /usr/local/bin/soffice; \
    soffice --version

# Unprivileged runtime user. No login shell, no home on the image: HOME is
# the /tmp tmpfs (see ENV below).
RUN groupadd --system --gid 10001 app \
    && useradd --system --uid 10001 --gid 10001 --home-dir /tmp \
       --no-create-home --shell /usr/sbin/nologin app

# ---- Python dependencies (both apps share one interpreter) ----
COPY pdf_to_doc_converter/requirements.txt /app/pdf_to_doc_converter/requirements.txt
COPY ppt_to_pdf_converter/requirements.txt /app/ppt_to_pdf_converter/requirements.txt
RUN pip install --upgrade pip \
    && pip install --no-cache-dir -r /app/pdf_to_doc_converter/requirements.txt \
    && pip install --no-cache-dir -r /app/ppt_to_pdf_converter/requirements.txt

# ---- App code ----
COPY pdf_to_doc_converter/ /app/pdf_to_doc_converter/
COPY ppt_to_pdf_converter/ /app/ppt_to_pdf_converter/
# "Made with Filo" watermark (module + logo), imported by both apps.
COPY watermark/ /app/watermark/
# Per-request time budget (README, "Request deadline"), imported by both apps.
COPY deadline/ /app/deadline/
COPY --chmod=0755 docker/start.sh /usr/local/bin/start.sh

ENV HOME=/tmp \
    TMPDIR=/tmp \
    XDG_CACHE_HOME=/tmp/.cache \
    XDG_CONFIG_HOME=/tmp/.config \
    NUMBA_CACHE_DIR=/tmp/.cache/numba \
    MPLCONFIGDIR=/tmp/.cache/matplotlib

# Engine limits (see README, "Environment variables" and "Request deadline").
# Every conversion request has a total budget of REQUEST_DEADLINE_SECONDS
# (upload, queueing and the conversion): past it the route answers a clean 504,
# before a typical 100 s upstream proxy limit. LibreOffice runs get
# min(LO_TIMEOUT_SECONDS, what is left of that budget).
ENV MAX_UPLOAD_MB=30 \
    MAX_PDF_PAGES=50 \
    REQUEST_DEADLINE_SECONDS=95 \
    LO_TIMEOUT_SECONDS=85 \
    LO_MAX_CONCURRENCY=2

# gunicorn tuning for the PDF to Word app (pdf_to_doc_converter/gunicorn_config.py).
ENV WEB_CONCURRENCY=2 \
    GUNICORN_THREADS=1 \
    GUNICORN_PRELOAD=true \
    PDF2DOCX_MAX_PROFILES=1

# tesseract is built with OpenMP and otherwise starts one thread per HOST core
# for every page it reads. OCR output does not depend on the thread count.
ENV OMP_THREAD_LIMIT=1

USER app

EXPOSE 5000 8081

CMD ["/usr/local/bin/start.sh"]
