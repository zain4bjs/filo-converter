import os


# Loopback by default; the image's start script passes --bind to publish it.
bind = f"127.0.0.1:{os.environ.get('PORT', '5000')}"
workers = int(os.environ.get("WEB_CONCURRENCY", "1"))
threads = int(os.environ.get("GUNICORN_THREADS", "2"))
# Backstop only: each request stops itself at REQUEST_DEADLINE_SECONDS (95 s,
# filo_deadline), so this must stay above that budget or gunicorn would kill
# the worker first (no clean 504, the client sees a 502).
timeout = int(os.environ.get("GUNICORN_TIMEOUT", "600"))
graceful_timeout = int(os.environ.get("GUNICORN_GRACEFUL_TIMEOUT", "60"))
# The container runs --read-only with a tmpfs on /tmp, so the worker heartbeat
# files must live under /tmp. An override outside /tmp is ignored.
_worker_tmp_dir = os.environ.get("GUNICORN_WORKER_TMP_DIR", "/tmp")
_worker_tmp_abs = os.path.abspath(_worker_tmp_dir)
worker_tmp_dir = _worker_tmp_abs if (_worker_tmp_abs + "/").startswith("/tmp/") else "/tmp"
max_requests = int(os.environ.get("GUNICORN_MAX_REQUESTS", "50"))
max_requests_jitter = int(os.environ.get("GUNICORN_MAX_REQUESTS_JITTER", "10"))
preload_app = os.environ.get("GUNICORN_PRELOAD", "false").lower() == "true"
# Error log to stderr; never to a file on the read-only fs.
# (Access logging stays off; log requests in a fronting proxy if needed.)
errorlog = "-"
