"""
services/pre_extract_worker.py
──────────────────────────────
PreExtractWorker — QObject worker launched in a QThread immediately after the
user clicks Submit on the Reference Bases page.

What it does
────────────
For EVERY source folder (target + all reference bases):
  1. Scan all C/C++ source files with detect_functions_in_file().
  2. Extract each function body with extract_function_body().
  3. Write the body to  <CACHE_ROOT>/<folder_key>/<safe_path>__<fn>.txt
  4. Write a  _index.json  that maps  "fn_lower|/norm/path" → metadata.
  5. Populate the in-memory MEMORY_BODY_CACHE so future lookups are instant.

After this worker finishes, extract_function_body() in core/utils.py will find
bodies in MEMORY_BODY_CACHE (O(1)), and the BuiltinExtractionWorker used by the
Report page will find the .txt files already on disk and skip re-scanning.
"""

import json
import os
import re
import shutil

from PySide6.QtCore import QObject, Signal

from core.logger import get_logger
from core.utils import (
    normalize_name, normalize_path,
    iter_source_files, detect_functions_in_file, extract_function_body,
)
from services.func_body_cache import (
    CACHE_ROOT, INDEX_FILE, cache_dir_for, store_body_in_memory,
)

_log = get_logger(__name__)


class PreExtractWorker(QObject):
    """
    Signals
    -------
    progress(pct: int, message: str)   — overall 0-100 % + status text
    folder_started(label: str)         — name of the folder now being processed
    folder_done(folder_path, fn_count) — folder finished; fn_count bodies saved
    log(message: str)                  — human-readable log line
    finished()                         — all folders done (or cancelled)
    error(message: str)                — unrecoverable failure
    """

    progress       = Signal(int, str)
    folder_started = Signal(str)
    folder_done    = Signal(str, int)
    log            = Signal(str)
    finished       = Signal()
    error          = Signal(str)

    def __init__(self, folders: list, function_filter=None):
        """
        Parameters
        ----------
        folders : list of dict
            Each entry: {"path": str, "role": "target"|"reference", "label": str}
        function_filter : iterable of str, optional
            Normalized function names to include for the TARGET folder only.
            Reference folders always extract every detected function.
        """
        super().__init__()
        self.folders         = folders
        self.function_filter = (
            {normalize_name(fn) for fn in function_filter if fn}
            if function_filter else set()
        )
        self._cancel = False

    def cancel(self) -> None:
        self._cancel = True

    # ── entry point ───────────────────────────────────────────────────────────

    def run(self) -> None:
        try:
            self._run()
        except Exception as exc:
            import traceback
            self.error.emit(f"{exc}\n{traceback.format_exc()}")

    # ── main logic ────────────────────────────────────────────────────────────

    def _run(self) -> None:
        os.makedirs(CACHE_ROOT, exist_ok=True)
        total = len(self.folders)
        if total == 0:
            self.finished.emit()
            return

        for fi, entry in enumerate(self.folders):
            if self._cancel:
                self.log.emit("Pre-extraction cancelled.")
                self.finished.emit()
                return

            folder_path = normalize_path(entry.get("path", ""))
            role        = entry.get("role", "reference")
            label       = entry.get("label") or os.path.basename(folder_path) or "folder"
            is_target   = (role == "target")

            if not os.path.isdir(folder_path):
                self.log.emit(f"⚠ Folder not found, skipping: {folder_path}")
                continue

            self.folder_started.emit(label)
            self.log.emit(f"[{fi+1}/{total}] Pre-extracting: {label}")

            # Clear old cache entry for this specific folder before writing
            out_dir = cache_dir_for(folder_path)
            if os.path.isdir(out_dir):
                shutil.rmtree(out_dir, ignore_errors=True)
            os.makedirs(out_dir, exist_ok=True)

            file_entries = list(iter_source_files(folder_path))
            total_files  = max(1, len(file_entries))
            index_data: dict = {}
            extracted   = 0
            seen_pairs: set = set()

            for fii, (full_path, file_name) in enumerate(file_entries):
                if self._cancel:
                    self.log.emit("Pre-extraction cancelled.")
                    self.finished.emit()
                    return

                # Compute blended progress:  fi folders done + current file fraction
                folder_share = 100 // total          # pct each folder contributes
                base_pct     = fi * folder_share
                file_pct     = int((fii / total_files) * folder_share)
                self.progress.emit(
                    min(base_pct + file_pct, 99),
                    f"[{fi+1}/{total}] {file_name}",
                )

                functions = detect_functions_in_file(full_path)
                if not functions:
                    continue

                for fn in functions:
                    if self._cancel:
                        self.log.emit("Pre-extraction cancelled.")
                        self.finished.emit()
                        return

                    key = normalize_name(fn)
                    if not key:
                        continue

                    # Deduplicate by (fn_key, file_path)
                    pair = (key, normalize_path(full_path))
                    if pair in seen_pairs:
                        continue

                    # For target only: honour function_filter when provided
                    if is_target and self.function_filter and key not in self.function_filter:
                        continue

                    body = extract_function_body(full_path, fn)

                    # Build a flat, filesystem-safe .txt filename
                    norm_fp   = normalize_path(full_path)
                    safe_path = re.sub(r"[/\\]", "__", norm_fp).strip("_")
                    safe_path = re.sub(r'[<>:"|?*]', "_", safe_path)
                    txt_name  = f"{safe_path}__{fn}.txt"

                    index_key = f"{fn.lower()}|{norm_fp}"

                    txt_path = os.path.join(out_dir, txt_name)
                    try:
                        with open(txt_path, "w", encoding="utf-8", errors="ignore") as fh:
                            fh.write(body)
                        index_data[index_key] = {
                            "display_name": fn,
                            "source_file":  full_path,
                            "txt_name":     txt_name,
                        }
                        seen_pairs.add(pair)
                        extracted += 1
                        # Warm the in-memory cache immediately
                        store_body_in_memory(full_path, fn, body)
                    except Exception as exc:
                        self.log.emit(f"  ⚠ Could not write {txt_name}: {exc}")

            # Write _index.json for this folder
            idx_path = os.path.join(out_dir, INDEX_FILE)
            try:
                with open(idx_path, "w", encoding="utf-8") as fh:
                    json.dump(index_data, fh, ensure_ascii=False, indent=2)
            except Exception as exc:
                self.log.emit(f"  ⚠ Could not write index for {label}: {exc}")

            self.folder_done.emit(folder_path, extracted)
            self.log.emit(
                f"  ✓ {extracted} functions cached for '{label}' "
                f"({total_files} files scanned)"
            )
            _log.info(
                "pre_extract: folder='%s' extracted=%d files=%d",
                label, extracted, total_files,
            )

        self.progress.emit(100, "Pre-extraction complete")
        self.log.emit("✅ All function bodies pre-extracted and cached.")
        self.finished.emit()
