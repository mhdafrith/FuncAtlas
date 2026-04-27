"""
core/logger.py
──────────────
FuncAtlas application logger  (enhanced with detailed event tracking).

Provides:
  • get_logger(name)          – standard module-level Logger
  • get_log_file_path()       – path to the current rotating log file
  • log_user_action(...)      – one-liner to record a user interaction
  • log_file_upload(...)      – log every file/folder the user selects
  • log_output_file(...)      – log every file written by the application
  • log_function_extraction() – log extraction results per file

Usage
-----
    from core.logger import get_logger, log_user_action, log_file_upload
    log = get_logger(__name__)

    log.info("Something happened")
    log_user_action("click", "Generate Report button", page="report")
    log_file_upload("folder", "/path/to/target", field="Target Base Folder")
    log_output_file("/path/to/output.xlsx", kind="Excel Report")

Log format (each line)
-----------------------
    2025-04-27 15:30:00 | INFO     | funcatlas.main_window | Message text
    2025-04-27 15:30:00 | EVENT    | funcatlas.events      | [CLICK] Generate Report button | page=report
    2025-04-27 15:30:00 | UPLOAD   | funcatlas.events      | [FOLDER] /path/to/target | field=Target Base Folder
    2025-04-27 15:30:00 | OUTPUT   | funcatlas.events      | [Excel Report] /path/to/output.xlsx
"""

import logging
import os
import sys
from logging.handlers import RotatingFileHandler
from datetime import datetime

# ── Configuration ─────────────────────────────────────────────────────────────
LOG_DIR        = os.path.join(os.path.expanduser("~"), ".funcatlas", "logs")
LOG_FILE_MAX   = 10 * 1024 * 1024   # 10 MB per file
LOG_BACKUP_CNT = 5                   # keep 5 rotated files
LOG_LEVEL      = logging.DEBUG

_CONSOLE_LEVEL = logging.INFO
_FILE_LEVEL    = logging.DEBUG

# ── Custom level numbers for event/upload/output entries ─────────────────────
_LEVEL_EVENT  = 25          # between INFO (20) and WARNING (30)
_LEVEL_UPLOAD = 26
_LEVEL_OUTPUT = 27

logging.addLevelName(_LEVEL_EVENT,  "EVENT ")
logging.addLevelName(_LEVEL_UPLOAD, "UPLOAD")
logging.addLevelName(_LEVEL_OUTPUT, "OUTPUT")

_FMT      = "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s"
_DATE_FMT = "%Y-%m-%d %H:%M:%S"

# ── Internal state ─────────────────────────────────────────────────────────────
_initialized     = False
_log_file_path   = ""
_event_logger    = None


# ─────────────────────────────────────────────────────────────────────────────
def _initialize():
    global _initialized, _log_file_path, _event_logger

    if _initialized:
        return

    os.makedirs(LOG_DIR, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    _log_file_path = os.path.join(LOG_DIR, f"funcatlas_{ts}.log")

    formatter = logging.Formatter(_FMT, datefmt=_DATE_FMT)

    root = logging.getLogger("funcatlas")
    root.setLevel(LOG_LEVEL)

    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(_CONSOLE_LEVEL)
    ch.setFormatter(formatter)
    root.addHandler(ch)

    try:
        fh = RotatingFileHandler(
            _log_file_path,
            maxBytes=LOG_FILE_MAX,
            backupCount=LOG_BACKUP_CNT,
            encoding="utf-8",
        )
        fh.setLevel(_FILE_LEVEL)
        fh.setFormatter(formatter)
        root.addHandler(fh)
    except OSError as exc:
        root.warning("Could not open log file %s: %s", _log_file_path, exc)

    _event_logger = logging.getLogger("funcatlas.events")

    for noisy in ("PIL", "openpyxl", "PySide6"):
        logging.getLogger(noisy).setLevel(logging.WARNING)

    root.info("=" * 72)
    root.info("FuncAtlas session started")
    root.info("Log file : %s", _log_file_path)
    root.info("=" * 72)
    _initialized = True


# ─────────────────────────────────────────────────────────────────────────────
def get_logger(name):
    """Return a child logger under the 'funcatlas' hierarchy."""
    _initialize()
    short = name.replace("funcatlas_updated.", "").replace("new.", "")
    return logging.getLogger(f"funcatlas.{short}")


def get_log_file_path():
    """Return the path of the current session log file."""
    _initialize()
    return _log_file_path


# ── Structured event helpers ──────────────────────────────────────────────────

def log_user_action(action, target, page="", extra=""):
    """Log a user-initiated UI action (click, toggle, clear, navigate, …).

    Args:
        action : verb   e.g. "click", "navigate", "toggle", "clear"
        target : noun   e.g. "Generate Report button", "View page"
        page   : which page was active (optional)
        extra  : any extra context string (optional)
    """
    _initialize()
    parts = [f"[{action.upper()}] {target}"]
    if page:
        parts.append(f"page={page}")
    if extra:
        parts.append(extra)
    _event_logger.log(_LEVEL_EVENT, " | ".join(parts))


def log_file_upload(kind, path, field="", count=0):
    """Log a file/folder selected by the user.

    Args:
        kind  : "file" | "folder" | "files" | "excel"
        path  : absolute path (or newline-joined list for multi-select)
        field : label of the input widget, e.g. "Target Base Folder"
        count : number of items selected (multi-select)
    """
    _initialize()
    tag = kind.upper()
    parts = [f"[{tag}] {path}"]
    if field:
        parts.append(f"field={field!r}")
    if count:
        parts.append(f"count={count}")
    if kind in ("file", "excel") and os.path.isfile(path):
        try:
            sz = os.path.getsize(path)
            parts.append(f"size={sz:,} bytes")
        except OSError:
            pass
    _event_logger.log(_LEVEL_UPLOAD, " | ".join(parts))


def log_output_file(path, kind="File"):
    """Log a file the application writes/creates.

    Args:
        path : absolute path of the output file
        kind : human label e.g. "Excel Report", "HTML Report", "Temp TXT"
    """
    _initialize()
    _event_logger.log(_LEVEL_OUTPUT, "[%s] %s", kind, path)


def log_function_extraction(file_path, function_names):
    """Log functions detected inside a single source file.

    Args:
        file_path      : absolute path of the scanned source file
        function_names : list of detected function name strings
    """
    _initialize()
    short = os.path.basename(file_path)
    count = len(function_names)
    if count == 0:
        _event_logger.log(_LEVEL_EVENT,
                          "[EXTRACT] %s | 0 functions found", short)
    else:
        preview = ", ".join(function_names[:8])
        tail    = f", … (+{count - 8} more)" if count > 8 else ""
        _event_logger.log(_LEVEL_EVENT,
                          "[EXTRACT] %s | %d functions: %s%s",
                          short, count, preview, tail)