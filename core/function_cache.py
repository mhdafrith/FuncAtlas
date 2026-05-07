"""
core/function_cache.py
──────────────────────
Disk-backed cache for extracted function bodies.

Directory layout
────────────────
  <system_temp>/FuncAtlas_Cache/
    target/
      <safe_folder_name>/
        _meta.json          ← scan records: {file_path: {display_name, functions}}
        _index.json         ← body index:   {fn|file_path: {display_name, source_file, txt_name}}
        <safe_path>__<fn>.txt
    reference/
      <safe_folder_name>/
        _meta.json
        _index.json
        <safe_path>__<fn>.txt

Public API
──────────
  FunctionCache                  – singleton-style class (instantiate once)
    .extract_and_cache(bases, function_filter, progress_cb, cancel_check)
       → dict  {folder_path: {'meta': {...}, 'index': {...}, 'dir': str}}
    .get_body(folder_path, role, file_path, func_name) -> str
    .get_meta(folder_path, role)  -> dict | None
    .get_index(folder_path, role) -> dict | None
    .is_cached(folder_path, role) -> bool
    .clear()                      – remove everything
    .clear_role(role)             – remove 'target' or 'reference' subtree
"""

from __future__ import annotations

import json
import os
import re
import shutil
import tempfile
from typing import Callable, Dict, List, Optional, Tuple

from core.utils import (
    normalize_path,
    iter_source_files,
    extract_function_body,
    detect_functions_in_file,
)

# ── Cache root ────────────────────────────────────────────────────────────────
_CACHE_ROOT = os.path.join(tempfile.gettempdir(), "FuncAtlas_Cache")
_TARGET_DIR = os.path.join(_CACHE_ROOT, "target")
_REF_DIR    = os.path.join(_CACHE_ROOT, "reference")

_ROLE_DIRS = {
    "target":    _TARGET_DIR,
    "reference": _REF_DIR,
}


# ── Name helpers ──────────────────────────────────────────────────────────────

def _safe_folder_name(folder_path: str) -> str:
    """Derive a filesystem-safe directory name from a full folder path.

    Keeps the last two path components joined with '__' so it is still
    human-readable, e.g.  C:\\work\\proj\\target_v2  →  proj__target_v2
    """
    norm = normalize_path(folder_path)
    parts = norm.replace("\\", "/").rstrip("/").split("/")
    tag   = "__".join(p for p in parts[-2:] if p) or "base"
    # strip illegal characters
    tag   = re.sub(r'[<>:"|?*\s]+', "_", tag).strip("_") or "base"
    # append a short stable hash to disambiguate collisions
    import hashlib
    h = hashlib.md5(norm.encode("utf-8", errors="replace")).hexdigest()[:6]
    return f"{tag}__{h}"


def _safe_txt_name(file_path: str, func_name: str) -> str:
    """Build a unique, flat .txt filename for (file_path, func_name)."""
    norm = normalize_path(file_path)
    safe = re.sub(r'[/\\]', "__", norm).strip("_")
    safe = re.sub(r'[<>:"|?*\s]+', "_", safe)
    fn   = re.sub(r'[^A-Za-z0-9_]', "_", func_name)
    return f"{safe}__{fn}.txt"


# ── FunctionCache ─────────────────────────────────────────────────────────────

class FunctionCache:
    """Manages disk-cached function bodies for target and reference bases."""

    # ── cache-dir helpers ─────────────────────────────────────────────────────

    def _cache_dir(self, folder_path: str, role: str) -> str:
        role_dir = _ROLE_DIRS.get(role, _REF_DIR)
        return os.path.join(role_dir, _safe_folder_name(folder_path))

    def is_cached(self, folder_path: str, role: str) -> bool:
        d = self._cache_dir(folder_path, role)
        return os.path.isfile(os.path.join(d, "_meta.json"))

    def get_meta(self, folder_path: str, role: str) -> Optional[dict]:
        """Return scan records {file_path: {display_name, functions}} or None."""
        p = os.path.join(self._cache_dir(folder_path, role), "_meta.json")
        if not os.path.isfile(p):
            return None
        try:
            with open(p, encoding="utf-8") as fh:
                return json.load(fh)
        except Exception:
            return None

    def get_index(self, folder_path: str, role: str) -> Optional[dict]:
        """Return body index {fn|file_path: {display_name, source_file, txt_name}} or None."""
        p = os.path.join(self._cache_dir(folder_path, role), "_index.json")
        if not os.path.isfile(p):
            return None
        try:
            with open(p, encoding="utf-8") as fh:
                return json.load(fh)
        except Exception:
            return None

    def get_body(self, folder_path: str, role: str,
                 file_path: str, func_name: str) -> Optional[str]:
        """Read a cached .txt body.  Returns None when not found on disk."""
        d = self._cache_dir(folder_path, role)
        txt = _safe_txt_name(file_path, func_name)
        full = os.path.join(d, txt)
        if os.path.isfile(full):
            try:
                with open(full, encoding="utf-8", errors="replace") as fh:
                    return fh.read()
            except Exception:
                pass
        return None

    # ── cache clearing ────────────────────────────────────────────────────────

    def clear(self):
        """Remove the entire cache (both target and reference)."""
        if os.path.isdir(_CACHE_ROOT):
            try:
                shutil.rmtree(_CACHE_ROOT, ignore_errors=True)
            except Exception:
                pass

    def clear_role(self, role: str):
        """Remove only the 'target' or 'reference' subtree."""
        role_dir = _ROLE_DIRS.get(role)
        if role_dir and os.path.isdir(role_dir):
            try:
                shutil.rmtree(role_dir, ignore_errors=True)
            except Exception:
                pass

    def clear_folder(self, folder_path: str, role: str):
        """Remove cached data for one specific folder."""
        d = self._cache_dir(folder_path, role)
        if os.path.isdir(d):
            try:
                shutil.rmtree(d, ignore_errors=True)
            except Exception:
                pass

    # ── extraction ────────────────────────────────────────────────────────────

    def extract_and_cache(
        self,
        bases: List[Dict],
        function_filter: Optional[set] = None,
        progress_cb: Optional[Callable[[str, int, int, str], None]] = None,
        cancel_check: Optional[Callable[[], bool]] = None,
    ) -> Dict[str, dict]:
        """Extract & cache function bodies for every base.

        Parameters
        ----------
        bases : list of dicts
            Each dict: {
              'folder_path': str,   # absolute path to source folder
              'role': str,          # 'target' | 'reference'
              'label': str,         # human-readable name
            }
        function_filter : set[str] | None
            Lowercase normalised function names.  When provided the TARGET base
            is filtered to only these functions; reference bases always extract
            all functions.
        progress_cb : callable(label, current_file, total_files, func_name)
            Called after each file is scanned.
        cancel_check : callable() -> bool
            Returns True when the operation should abort.

        Returns
        -------
        dict  {folder_path: {'meta': {...}, 'index': {...}, 'dir': str}}
        """
        results: Dict[str, dict] = {}

        for base in bases:
            folder_path = normalize_path(base["folder_path"])
            role        = base["role"]           # 'target' | 'reference'
            label       = base.get("label", os.path.basename(folder_path))
            apply_filter = (role == "target") and bool(function_filter)

            if cancel_check and cancel_check():
                break

            # ── prepare cache directory ───────────────────────────────────
            cache_dir = self._cache_dir(folder_path, role)
            # Wipe stale data for this folder so we always start clean
            if os.path.isdir(cache_dir):
                shutil.rmtree(cache_dir, ignore_errors=True)
            os.makedirs(cache_dir, exist_ok=True)

            # ── scan source files ─────────────────────────────────────────
            file_entries = list(iter_source_files(folder_path))
            total = max(1, len(file_entries))

            meta_records: Dict[str, dict] = {}      # file_path -> {display_name, functions}
            index_data:   Dict[str, dict] = {}      # fn|file_path -> {display_name, source_file, txt_name}
            seen_pairs:   set             = set()

            for file_idx, (file_path, file_name) in enumerate(file_entries, 1):
                if cancel_check and cancel_check():
                    break

                try:
                    funcs = detect_functions_in_file(file_path)
                except Exception:
                    funcs = []

                if not isinstance(funcs, list):
                    funcs = []

                # Filter: target + active filter → only listed functions
                if apply_filter:
                    funcs = [fn for fn in funcs
                             if fn.strip().lower() in function_filter]

                meta_records[file_path] = {
                    "display_name": file_name,
                    "functions":    funcs,
                }

                for fn in funcs:
                    pair_key = (fn.lower(), normalize_path(file_path))
                    if pair_key in seen_pairs:
                        continue
                    seen_pairs.add(pair_key)

                    body = extract_function_body(file_path, fn)
                    txt_name = _safe_txt_name(file_path, fn)

                    if progress_cb:
                        progress_cb(label, file_idx, total, fn)

                    try:
                        with open(os.path.join(cache_dir, txt_name),
                                  "w", encoding="utf-8", errors="ignore") as fh:
                            fh.write(body)
                        index_key = f"{fn.lower()}|{normalize_path(file_path)}"
                        index_data[index_key] = {
                            "display_name": fn,
                            "source_file":  file_path,
                            "txt_name":     txt_name,
                        }
                    except Exception:
                        pass

            # ── write metadata ────────────────────────────────────────────
            try:
                with open(os.path.join(cache_dir, "_meta.json"),
                          "w", encoding="utf-8") as fh:
                    json.dump(meta_records, fh, indent=2)
                with open(os.path.join(cache_dir, "_index.json"),
                          "w", encoding="utf-8") as fh:
                    json.dump(index_data, fh, indent=2)
            except Exception:
                pass

            results[folder_path] = {
                "meta":  meta_records,
                "index": index_data,
                "dir":   cache_dir,
            }

        return results


# ── Module-level singleton ────────────────────────────────────────────────────
FUNCTION_CACHE = FunctionCache()
