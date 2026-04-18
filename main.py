"""
main.py
───────
FuncAtlas entry point.
Run with:  python main.py
"""

import sys
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QFont
from PySide6.QtCore import Qt


def main():
    # High-DPI support
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)

    app = QApplication(sys.argv)

    # Import window AFTER QApplication exists so Qt internals are fully ready
    from main_window import ReuseAnalysisWindow

    app.setFont(QFont("Segoe UI", 11))

    window = ReuseAnalysisWindow()
    window.show()

    exit_code = app.exec()

    # closeEvent already ran when the user closed the window, but call
    # close() explicitly here as a safety net in case exec() returned
    # for another reason (e.g. app.quit()).  Then delete before QApplication
    # goes out of scope so no threads are still alive during Qt teardown.
    window.close()
    del window

    sys.exit(exit_code)


if __name__ == "__main__":
    main()