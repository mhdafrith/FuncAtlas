"""
main.py
───────
FuncAtlas entry point.
Run with:  python main.py
"""

import sys
import os
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QFont, QIcon
from PySide6.QtCore import Qt

# ── Logger must be initialised before anything else ──────────────────────────
from core.logger import get_logger, get_log_file_path
log = get_logger(__name__)


def main():
    log.info("FuncAtlas starting up")

    # High-DPI support
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)

    app = QApplication(sys.argv)

    # ── App icon ──────────────────────────────────────────────────────────────
    _base_dir = os.path.dirname(os.path.abspath(__file__))
    _icon_path = os.path.join(_base_dir, "app_icon.png")
    if os.path.isfile(_icon_path):
        app.setWindowIcon(QIcon(_icon_path))
        log.debug("App icon loaded from %s", _icon_path)
    else:
        log.warning("App icon not found at %s", _icon_path)

    # Import window AFTER QApplication exists so Qt internals are fully ready
    from main_window import ReuseAnalysisWindow

    app.setFont(QFont("Segoe UI", 11))

    log.info("Main window creating")
    window = ReuseAnalysisWindow()
    window.show()
    log.info("Main window shown — log file: %s", get_log_file_path())

    exit_code = app.exec()
    log.info("Application exiting with code %d", exit_code)

    # closeEvent already ran when the user closed the window, but call
    # close() explicitly here as a safety net in case exec() returned
    # for another reason (e.g. app.quit()).  Then delete before QApplication
    # goes out of scope so no threads are still alive during Qt teardown.
    window.close()
    del window

    sys.exit(exit_code)


if __name__ == "__main__":
    main()