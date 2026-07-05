"""Central crash logging for every tool in the package.

Call :func:`install` once at program start (each ``main()`` entry point
already does).  From then on, any uncaught exception — in plain Python
code, in a background thread, or inside a Tkinter callback — is written
with its full traceback to a log file before the usual handling runs.

Log location: ``$DATA_EXTRACTION_LOG_DIR`` if set, otherwise
``~/.data_extraction/logs/crashes.log``.
"""

from __future__ import annotations

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

_logger: logging.Logger | None = None
_installed = False


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
        log_exception(f"Uncaught exception (thread {args.thread.name!r})",
                      args.exc_type, args.exc_value, args.exc_traceback)
    sys.stderr.write("".join(traceback.format_exception(
        args.exc_type, args.exc_value, args.exc_traceback)))


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
    """Install crash hooks for the interpreter, threads and Tkinter."""
    global _installed
    if _installed:
        return
    _installed = True
    sys.excepthook = _sys_hook
    threading.excepthook = _thread_hook
    try:
        import tkinter
        tkinter.Tk.report_callback_exception = _tk_callback_hook
    except Exception:
        pass  # headless environment without Tk support

    started = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        with open(LOG_FILE, "a", encoding="utf-8") as fh:
            fh.write(f"--- session started {started} "
                     f"({os.path.basename(sys.argv[0]) or 'python'}) ---\n")
    except OSError:
        pass
