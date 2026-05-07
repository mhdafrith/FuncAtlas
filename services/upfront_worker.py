"""
services/upfront_worker.py
──────────────────────────
UpfrontExtractionWorker  –  Qt worker that runs at Submit time to extract
every function body from target + reference bases and persist them to the
disk cache (core/function_cache.py).

Once this worker finishes, all subsequent navigation (view page source
switching, diff rendering, report generation, complexity analysis) can
read function bodies straight from the .txt cache without re-scanning.

Signals
───────
  started(label)                 – a new base is beginning
  progress(label, current, total, func_name)
  base_done(label, count)        – one base finished, `count` functions written
  finished(results)              – all bases done; results = {folder_path: {...}}
  error(message)                 – something went wrong
  log(message)                   – informational text
"""

from __future__ import annotations

from typing import Dict, List, Optional

from PySide6.QtCore import QObject, Signal

from core.function_cache import FUNCTION_CACHE
from core.logger import get_logger

_log = get_logger(__name__)


class UpfrontExtractionWorker(QObject):
    """Extracts and caches function bodies for all supplied bases.

    Parameters
    ----------
    bases : list[dict]
        Each dict must contain:
          'folder_path' : str  – absolute path to the source folder
          'role'        : str  – 'target' | 'reference'
          'label'       : str  – human-readable name shown in the UI
    function_filter : set[str] | None
        Lowercase-normalised function names.  When provided the TARGET base
        is filtered; reference bases always extract everything.
    """

    # Qt signals
    started   = Signal(str)                   # label
    progress  = Signal(str, int, int, str)    # label, current_file, total_files, func_name
    base_done = Signal(str, int)              # label, extracted_count
    finished  = Signal(dict)                  # {folder_path: {'meta', 'index', 'dir'}}
    error     = Signal(str)
    log       = Signal(str)

    def __init__(
        self,
        bases: List[Dict],
        function_filter: Optional[set] = None,
    ):
        super().__init__()
        self.bases            = bases
        self.function_filter  = function_filter or set()
        self._cancel_requested = False

    # ── cancellation ─────────────────────────────────────────────────────────

    def cancel(self):
        self._cancel_requested = True

    # ── main run ──────────────────────────────────────────────────────────────

    def run(self):
        if not self.bases:
            self.error.emit("No bases provided.")
            return

        _log.info("UpfrontExtractionWorker: starting for %d base(s)", len(self.bases))

        try:
            per_base_counts: Dict[str, int] = {}

            def _progress(label: str, current: int, total: int, func_name: str):
                self.progress.emit(label, current, total, func_name)

            def _cancel_check() -> bool:
                return self._cancel_requested

            # Wrap extract_and_cache so we can emit base_done per base
            all_results: dict = {}
            for base in self.bases:
                if self._cancel_requested:
                    self.error.emit("__CANCELLED__")
                    return

                label       = base.get("label", base.get("folder_path", ""))
                folder_path = base["folder_path"]

                self.started.emit(label)
                self.log.emit(f"Extracting: {label}")

                result = FUNCTION_CACHE.extract_and_cache(
                    bases=[base],
                    function_filter=self.function_filter if self.function_filter else None,
                    progress_cb=_progress,
                    cancel_check=_cancel_check,
                )

                if self._cancel_requested:
                    self.error.emit("__CANCELLED__")
                    return

                all_results.update(result)

                # Count extracted functions for this base
                base_meta  = result.get(folder_path, {}).get("meta", {})
                count      = sum(len(v.get("functions", [])) for v in base_meta.values())
                per_base_counts[label] = count

                self.base_done.emit(label, count)
                self.log.emit(f"Done: {label} → {count} function(s) cached")
                _log.info("UpfrontExtractionWorker: %s → %d functions", label, count)

            _log.info("UpfrontExtractionWorker: all bases complete")
            self.finished.emit(all_results)

        except Exception as exc:
            _log.exception("UpfrontExtractionWorker.run crashed")
            self.error.emit(str(exc))
