"""Central crash logging for every tool in the package.

Call :func:`install` once at program start (each ``main()`` entry point
already does).  From then on, any uncaught exception — in plain Python
code, in a background thread, in a Tkinter callback, or an *unraisable*
one (``__del__`` etc.) — is written with its full traceback to a log file
before the usual handling runs.

Python hooks cannot see a **native** crash (a Tcl/Tk abort or a segfault
in a C extension — the usual way a Tkinter app dies on macOS without a
Python traceback), so :func:`install` additionally enables
``faulthandler`` into ``fatal_faults.log``: on a fatal signal the Python
stacks of all threads are dumped there.  A ``session ended`` marker is
written on clean interpreter exit — a ``session started`` line without a
matching end marker means the process died hard; look in
``fatal_faults.log`` for where.

Log location: ``$DATA_EXTRACTION_LOG_DIR`` if set, otherwise
``~/.data_extraction/logs/`` (``crashes.log`` + ``fatal_faults.log``).
"""

from __future__ import annotations

import atexit
import faulthandler
import logging
import os
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path

LOG_DIR = Path(os.environ.get("DATA_EXTRACTION_LOG_DIR",
                              Path.home() / ".data_extraction" / "logs"))
LOG_FILE = LOG_DIR / "crashes.log"
FAULT_FILE = LOG_DIR / "fatal_faults.log"

_logger: logging.Logger | None = None
_installed = False
_fault_fh = None            # kept referenced: faulthandler writes into it


def get_logger() -> logging.Logger:
    """Return the crash logger, creating the log file on first use."""
    global _logger
    if _logger is None:
        logger = logging.getLogger("data_extraction.crash")
        logger.setLevel(logging.ERROR)
        logger.propagate = False
        try:
            LOG_DIR.mkdir(parents=True, exist_ok=True)
            handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
            handler.setFormatter(logging.Formatter(
                "%(asctime)s | %(levelname)s | %(message)s"))
            logger.addHandler(handler)
        except OSError:
            # Unwritable log dir must never take the app down with it.
            logger.addHandler(logging.StreamHandler(sys.stderr))
        _logger = logger
    return _logger


def log_exception(context: str, exc_type, exc_value, exc_tb) -> None:
    """Write one formatted traceback to the crash log."""
    try:
        text = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        get_logger().error("%s\n%s", context, text)
    except Exception:
        pass  # logging failures are silently dropped


def _sys_hook(exc_type, exc_value, exc_tb):
    if not issubclass(exc_type, (KeyboardInterrupt, SystemExit)):
        log_exception("Uncaught exception (main thread)",
                      exc_type, exc_value, exc_tb)
    sys.__excepthook__(exc_type, exc_value, exc_tb)


def _thread_hook(args):
    if not issubclass(args.exc_type, SystemExit):
        thread = getattr(args, "thread", None)
        name = getattr(thread, "name", "?")
        log_exception(f"Uncaught exception (thread {name!r})",
                      args.exc_type, args.exc_value, args.exc_traceback)
    sys.stderr.write("".join(traceback.format_exception(
        args.exc_type, args.exc_value, args.exc_traceback)))


def _unraisable_hook(args):
    """Exceptions Python cannot raise (``__del__``, weakref callbacks …) —
    they otherwise vanish with a one-line stderr note and no traceback."""
    log_exception(
        f"Unraisable exception ({args.err_msg or 'no message'}; "
        f"object: {args.object!r})",
        args.exc_type, args.exc_value, args.exc_traceback)
    sys.__unraisablehook__(args)


def _tk_callback_hook(self, exc_type, exc_value, exc_tb):
    """Replacement for Tk.report_callback_exception on all roots."""
    log_exception("Uncaught exception (Tk callback)",
                  exc_type, exc_value, exc_tb)
    traceback.print_exception(exc_type, exc_value, exc_tb)
    try:
        from tkinter import messagebox
        if self.winfo_exists():
            messagebox.showerror(
                "Unexpected Error",
                f"{exc_type.__name__}: {exc_value}\n\n"
                f"Details were logged to:\n{LOG_FILE}",
                parent=self)
    except Exception:
        pass  # never let the error dialog cause a second crash


def install() -> None:
    """Install crash hooks for the interpreter, threads, Tkinter,
    unraisable exceptions and native faults."""
    global _installed, _fault_fh
    if _installed:
        return
    _installed = True
    sys.excepthook = _sys_hook
    threading.excepthook = _thread_hook
    sys.unraisablehook = _unraisable_hook
    try:
        import tkinter
        tkinter.Tk.report_callback_exception = _tk_callback_hook
    except Exception:
        pass  # headless environment without Tk support

    started = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    prog = os.path.basename(sys.argv[0]) or "python"
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        with open(LOG_FILE, "a", encoding="utf-8") as fh:
            fh.write(f"--- session started {started} ({prog}) ---\n")
    except OSError:
        pass

    # Native crashes (a Tcl/Tk abort, a segfault in a C extension) kill the
    # process without ever reaching the Python hooks above — the classic
    # "the log only has the session line" failure. faulthandler dumps every
    # thread's Python stack into FAULT_FILE when a fatal signal hits.
    try:
        _fault_fh = open(FAULT_FILE, "a", encoding="utf-8")
        _fault_fh.write(f"--- session started {started} ({prog}) ---\n")
        _fault_fh.flush()
        faulthandler.enable(file=_fault_fh, all_threads=True)
    except Exception:
        _fault_fh = None  # never let diagnostics take the app down

    # A start line without this end marker = the process died hard
    # (native crash or kill) — check fatal_faults.log for the stacks.
    def _session_end():
        ended = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as fh:
                fh.write(f"--- session ended {ended} ({prog}) ---\n")
        except OSError:
            pass

    atexit.register(_session_end)
