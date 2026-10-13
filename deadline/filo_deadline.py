"""Per-request time budget for the conversion routes.

Many reverse proxies and CDNs cut a request at 100 s. A conversion that is
still running after that only burns CPU for a client that has already been
given up on, with no useful message. So every conversion route gets one total
budget, REQUEST_DEADLINE_SECONDS (default 95), and answers a clean 504
``{"error": "Conversion took too long"}`` when it runs out, before such a
proxy limit.

When a reverse proxy in front passes its own clock and how long the request
had already spent with it in two headers (``X-Filo-Request-Msec`` and
``X-Filo-Request-Time``, e.g. nginx's ``$msec`` and ``$request_time``), the
budget counts from the moment that proxy received the request, so time spent
queued in front of the app is counted too. The proxy must set both headers on
every request (replacing a client's own values). Without them (the app run
directly, tests) the budget starts when the app saw the request.

Shared by the PDF to Word app (Flask) and the Office app (FastAPI); standard
library only.
"""

from __future__ import annotations

import math
import multiprocessing
import os
import signal
import threading
import time
from typing import Any, Callable, Mapping

DEFAULT_SECONDS = 95.0
ENV_NAME = "REQUEST_DEADLINE_SECONDS"

# Set by a reverse proxy in front, when there is one (see the module docstring).
NGINX_MSEC_HEADER = "X-Filo-Request-Msec"  # the proxy's clock when it forwarded the request
NGINX_ELAPSED_HEADER = "X-Filo-Request-Time"  # seconds the request had spent in the proxy

# The client-facing message of the 504.
MESSAGE = "Conversion took too long"


class DeadlineExceeded(Exception):
    """The request's time budget ran out; answer 504 with MESSAGE."""

    def __init__(self, message: str = MESSAGE):
        super().__init__(message)
        self.message = message


def budget_seconds(environ: Mapping[str, str] | None = None) -> float:
    """REQUEST_DEADLINE_SECONDS, or DEFAULT_SECONDS when unset or invalid.

    Empty values and unexpanded CI macros ("$(VAR)") count as unset,
    like the apps' other settings; so does anything that is not a positive
    finite number (a typo must not disable the limit).
    """
    environ = os.environ if environ is None else environ
    raw = (environ.get(ENV_NAME) or "").strip()
    if not raw or raw.startswith("$("):
        return DEFAULT_SECONDS
    try:
        value = float(raw)
    except ValueError:
        return DEFAULT_SECONDS
    if not math.isfinite(value) or value <= 0:
        return DEFAULT_SECONDS
    return value


def _float(value: Any) -> float | None:
    try:
        number = float(str(value).strip())
    except (TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None


def elapsed_before_app(msec: Any, request_time: Any, now_wall: float) -> float:
    """Seconds the request had already been in flight when the app saw it.

    ``msec`` and ``request_time`` are the proxy's clock and elapsed time
    taken at the same moment, so ``msec - request_time`` is when the proxy first
    received the request, on the container's clock (shared with the app).
    Missing or nonsense values count as 0; a negative result (clock step)
    counts as 0.
    """
    sent_at = _float(msec)
    in_nginx = _float(request_time)
    if sent_at is None:
        return max(0.0, in_nginx or 0.0)
    elapsed = now_wall - (sent_at - max(0.0, in_nginx or 0.0))
    return elapsed if elapsed > 0 else 0.0


class Deadline:
    """The time left for one request. Thread-safe to read and to expire."""

    def __init__(
        self,
        seconds: float | None = None,
        already_elapsed: float = 0.0,
        clock: Callable[[], float] = time.monotonic,
    ):
        self.seconds = budget_seconds() if seconds is None else float(seconds)
        self._clock = clock
        spent = min(max(0.0, float(already_elapsed)), self.seconds)
        self._end = clock() + self.seconds - spent
        self._expired = threading.Event()

    @classmethod
    def from_headers(
        cls,
        get_header: Callable[[str], Any],
        seconds: float | None = None,
        clock: Callable[[], float] = time.monotonic,
        wall: Callable[[], float] = time.time,
    ) -> "Deadline":
        """Budget for a request, counting the time it already spent in a proxy."""
        elapsed = elapsed_before_app(get_header(NGINX_MSEC_HEADER), get_header(NGINX_ELAPSED_HEADER), wall())
        return cls(seconds, already_elapsed=elapsed, clock=clock)

    def remaining(self) -> float:
        if self._expired.is_set():
            return 0.0
        return max(0.0, self._end - self._clock())

    def elapsed(self) -> float:
        return self.seconds - self.remaining()

    def expired(self) -> bool:
        return self.remaining() <= 0.0

    def expire(self) -> None:
        """Mark the budget as spent, so work still running stops at its next check."""
        self._expired.set()

    def check(self) -> None:
        """Raise DeadlineExceeded once the budget is spent."""
        if self.expired():
            raise DeadlineExceeded()

    def timeout(self, cap: float | None = None) -> float:
        """Time left (at most ``cap``) for a blocking call; raises when none is left.

        Never returns 0: several libraries read a 0 timeout as "no timeout"
        (pytesseract does).
        """
        left = self.remaining()
        if left <= 0.0:
            raise DeadlineExceeded()
        return left if cap is None else min(cap, left)


# ----- Running work in a killable child process -----------------------------


def _child_main(conn, func, args, kwargs) -> None:
    # Own session and process group, so the parent can kill everything this
    # work spawns (ocrmypdf and its tesseract workers) with one killpg.
    try:
        os.setsid()
    except OSError:
        pass
    try:
        result = func(*args, **kwargs)
    except BaseException as exc:  # reported to the parent, never re-raised here
        try:
            conn.send(("error", f"{type(exc).__name__}: {exc}"))
        except Exception:
            pass
        raise
    try:
        conn.send(("ok", result))
    except Exception as exc:  # unpicklable result
        conn.send(("error", f"result not transferable: {type(exc).__name__}"))
    finally:
        conn.close()


class ChildFailed(Exception):
    """The child raised or died; the message is for the server log only."""


def _kill_group(process) -> None:
    try:
        os.killpg(process.pid, signal.SIGKILL)
    except (ProcessLookupError, PermissionError):
        pass
    try:
        process.kill()
    except Exception:
        pass
    process.join()


def run_in_child(
    func: Callable[..., Any],
    args: tuple = (),
    deadline: Deadline | None = None,
    kwargs: dict | None = None,
) -> Any:
    """Run ``func(*args)`` in a forked child, killing it when the deadline ends.

    Returns what ``func`` returned (it must be picklable). Raises
    DeadlineExceeded when the budget ran out first (the child and its whole
    process group are SIGKILLed and reaped, nothing keeps running), and
    ChildFailed when ``func`` raised or the child died. Files the child was
    writing are left for the caller's cleanup.

    POSIX fork only: the child inherits the already-imported converter, so
    nothing is re-imported or pickled on the way in.
    """
    deadline = deadline or Deadline()
    wait = deadline.timeout()
    context = multiprocessing.get_context("fork")
    parent_conn, child_conn = context.Pipe(duplex=False)
    process = context.Process(target=_child_main, args=(child_conn, func, args, kwargs or {}), daemon=True)
    process.start()
    child_conn.close()
    try:
        # poll() is also True when the child died without sending (EOF).
        if not parent_conn.poll(wait):
            deadline.expire()
            _kill_group(process)
            raise DeadlineExceeded()
        try:
            status, payload = parent_conn.recv()
        except EOFError:
            process.join()
            raise ChildFailed(f"child exited with code {process.exitcode} without a result") from None
        process.join(timeout=5)
        if process.is_alive():
            _kill_group(process)
        if status != "ok":
            raise ChildFailed(payload)
        return payload
    finally:
        parent_conn.close()
        if process.is_alive():
            _kill_group(process)
