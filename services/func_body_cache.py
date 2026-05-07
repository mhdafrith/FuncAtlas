"""
services/func_body_cache.py
───────────────────────────
Disk-backed cache for pre-extracted function bodies.

Layout
------
  <CACHE_ROOT>/
    <folder_key>/          ← one sub-dir per source folder
      _index.json          ← {index_key: {display_name, source_file, txt_name}}
      <safe_path>__<fn>.txt

index_key format: "fn_name_lower|/normalized/source/file/path"

Thread-safety
-------------
All mutating operations are protected by _LOCK.  Read operations (is_ready,
load_index, get_body) are lock-free (they only read already-written files).
"""

import hashlib
import json
import os
import re
import shutil
import tempfile
from threading import Lock

from core.logger import get_logger

_log  = get_logger(__name__)
_LOCK = Lock()

# ── public constants ──────────────────────────────────────────────────────────
CACHE_ROOT = os.path.join(tempfile.gettempdir(), "FuncAtlas_BodyCache")
INDEX_FILE = "_index.json"

# In-memory body cache:  (norm_file_path, fn_name_lower) → body_text
# Populated when bodies are written; consulted by extract_function_body.
MEMORY_BODY_CACHE: dict = {}
MEMORY_BODY_LOCK  = Lock()


# ── folder key helpers ────────────────────────────────────────────────────────

def _folder_key(folder_path: str) -> str:
    """Stable, filesystem-safe key that uniquely identifies a source folder."""
    norm = os.path.normpath(folder_path.strip()).lower()
    h    = hashlib.md5(norm.encode("utf-8", "replace")).hexdigest()[:12]
    safe = re.sub(r"[^A-Za-z0-9_.-]", "_", os.path.basename(norm))[:40]
    return f"{safe}__{h}"


def cache_dir_for(folder_path: str) -> str:
    """Return the cache sub-directory for a given source folder."""
    return os.path.join(CACHE_ROOT, _folder_key(folder_path))


# ── status helpers ────────────────────────────────────────────────────────────

def is_ready(folder_path: str) -> bool:
    """True iff the folder has a complete _index.json in the cache."""
    return os.path.isfile(os.path.join(cache_dir_for(folder_path), INDEX_FILE))


def load_index(folder_path: str) -> dict:
    """Load the _index.json for a cached folder (returns {} on failure)."""
    idx_path = os.path.join(cache_dir_for(folder_path), INDEX_FILE)
    try:
        with open(idx_path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


# ── body lookup ───────────────────────────────────────────────────────────────

def get_body(folder_path: str, file_path: str, fn_name: str) -> "str | None":
    """
    Return the pre-extracted body for (file_path, fn_name) from the disk cache,
    or None if it is not available.
    """
    # Fast in-memory path first
    mem_key = (os.path.normpath(file_path), fn_name.lower())
    with MEMORY_BODY_LOCK:
        hit = MEMORY_BODY_CACHE.get(mem_key)
    if hit is not None:
        return hit

    # Disk path
    idx = load_index(folder_path)
    if not idx:
        return None
    norm_fp   = os.path.normpath(file_path)
    index_key = f"{fn_name.lower()}|{norm_fp}"
    entry     = idx.get(index_key)
    if not entry:
        return None
    txt_path = os.path.join(cache_dir_for(folder_path), entry.get("txt_name", ""))
    if not os.path.isfile(txt_path):
        return None
    try:
        with open(txt_path, "r", encoding="utf-8", errors="replace") as fh:
            body = fh.read()
        # Populate memory cache for next lookup
        with MEMORY_BODY_LOCK:
            MEMORY_BODY_CACHE[mem_key] = body
        return body
    except Exception:
        return None


def store_body_in_memory(file_path: str, fn_name: str, body: str) -> None:
    """Store a body in the in-memory cache (called by PreExtractWorker)."""
    key = (os.path.normpath(file_path), fn_name.lower())
    with MEMORY_BODY_LOCK:
        MEMORY_BODY_CACHE[key] = body


# ── cache maintenance ─────────────────────────────────────────────────────────

def clear_all() -> None:
    """Delete the entire cache root and flush the in-memory cache."""
    with _LOCK:
        _flush_memory()
        try:
            if os.path.isdir(CACHE_ROOT):
                shutil.rmtree(CACHE_ROOT, ignore_errors=True)
                _log.info("func_body_cache: cleared entire cache at %s", CACHE_ROOT)
        except Exception as exc:
            _log.warning("func_body_cache: clear_all failed: %s", exc)


def clear_for(folder_paths: list) -> None:
    """Delete cache entries for specific source folders and flush their memory entries."""
    with _LOCK:
        _flush_memory()
        for fp in folder_paths:
            d = cache_dir_for(fp)
            try:
                if os.path.isdir(d):
                    shutil.rmtree(d, ignore_errors=True)
                    _log.info("func_body_cache: cleared cache for %s", fp)
            except Exception as exc:
                _log.warning("func_body_cache: clear_for(%s) failed: %s", fp, exc)


def _flush_memory() -> None:
    """Clear the in-memory body cache (call while holding _LOCK)."""
    with MEMORY_BODY_LOCK:
        MEMORY_BODY_CACHE.clear()
