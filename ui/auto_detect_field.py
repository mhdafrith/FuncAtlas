"""
ui/auto_detect_field.py
───────────────────────
AutoDetectColumnField widget.

detect_kind="function"     → detect function-name column from Function List excel files
detect_kind="db_function"  → detect function-name column from Consolidated DB excel  (FIX 2/3)
detect_kind="base"         → detect base-file-name column from Consolidated DB excel
"""

from PySide6.QtCore import Qt, QThread, Signal, QObject
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QMessageBox
)
from ui.widgets import IconTextButton
from core.utils import (
    normalize_excel_reference, is_valid_excel_reference,
    extract_excel_column_letters, extract_excel_row_number, excel_col_to_index
)
from core.logger import get_logger, log_user_action

_log = get_logger(__name__)


class _DetectWorker(QObject):
    finished = Signal(dict)
    error    = Signal(str)

    def __init__(self, excel_path, kind):
        super().__init__()
        self.excel_path = excel_path
        # Normalise kind: "db_function" uses the same scoring as "function"
        self.kind = "function" if kind == "db_function" else kind

    def run(self):
        try:
            from services.analysis import detect_best_column_in_workbook
            result = detect_best_column_in_workbook(self.excel_path, self.kind)
            if result:
                self.finished.emit(result)
            else:
                self.error.emit(f"Could not detect the {self.kind} column automatically.")
        except Exception as e:
            self.error.emit(str(e))


class AutoDetectColumnField(QWidget):
    def __init__(self, label_text, placeholder_text, button_icon, detect_kind, owner_window):
        super().__init__()
        self.detect_kind   = detect_kind
        self.owner_window  = owner_window
        self.detected_info = None
        self._thread       = None
        self._worker       = None
        self._threads: list = []

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        self.label = QLabel(label_text)
        self.label.setObjectName("fieldLabel")
        layout.addWidget(self.label)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(10)

        self.input = QLineEdit()
        self.input.setPlaceholderText(placeholder_text)
        self.input.setFixedHeight(38)
        self.input.editingFinished.connect(self.manual_apply)

        self.button = IconTextButton("Auto Detect", button_icon)
        self.button.setObjectName("pickerButton")
        self.button.setFixedSize(130, 38)
        self.button.clicked.connect(self.detect_from_excel)

        row.addWidget(self.input, 1)
        row.addWidget(self.button, 0, Qt.AlignBottom)
        layout.addLayout(row)

        self.preview = QLabel("Preview: not detected")
        self.preview.setObjectName("panelSubtitle")
        self.preview.setWordWrap(True)
        layout.addWidget(self.preview)

    def _get_source_excel(self):
        """
        Resolve the Excel file to scan based on detect_kind:

        "function"     → first file from the Function List xlsx picker
        "db_function"  → the Consolidated DB Excel (FIX 2/3: detects function
                          names inside the consolidated workbook)
        "base"         → the Consolidated DB Excel (same as db_function source,
                          but scorer looks for base-file-name headers)
        """
        if self.detect_kind == "function":
            # Detect from Function List excel (left toggle panel)
            files = self.owner_window.con_function_field.value()
            if not files:
                QMessageBox.warning(self, "Missing File",
                    "Please upload at least one Function List Excel file first.")
                return None
            return files[0]

        elif self.detect_kind in ("db_function", "base"):
            # FIX 2 & 3: both read from the Consolidated DB Excel
            path = self.owner_window.con_db_excel_field.value()
            if not path:
                QMessageBox.warning(self, "Missing File",
                    "Please upload the Consolidated DB Excel file first.")
                return None
            return path

        else:
            # Fallback: try DB excel
            path = getattr(self.owner_window, "con_db_excel_field", None)
            if path:
                v = path.value()
                if v:
                    return v
            QMessageBox.warning(self, "Missing File",
                "Cannot determine source Excel for auto-detection.")
            return None

    def manual_apply(self):
        value = normalize_excel_reference(self.input.text())
        self.input.setText(value)
        if is_valid_excel_reference(value):
            col    = extract_excel_column_letters(value)
            row_no = extract_excel_row_number(value)
            col_no = excel_col_to_index(col)
            self.preview.setText(
                f"Preview: {value} | Column={col} | Column No={col_no} | Row={row_no}")
            log_user_action("manual_entry", f"Column ref={value}",
                            extra=f"kind={self.detect_kind}")
            _log.debug("AutoDetectColumnField manual_apply: kind=%s ref=%s",
                       self.detect_kind, value)
        else:
            self.preview.setText("Preview: invalid reference")

    def detect_from_excel(self):
        excel_path = self._get_source_excel()
        if not excel_path:
            return

        _log.info("AutoDetect starting: kind=%s  excel=%s", self.detect_kind, excel_path)
        log_user_action("click", "Auto Detect button",
                        extra=f"kind={self.detect_kind} | source={excel_path}")

        self.button.setEnabled(False)
        self.button.setText("Detecting...")
        self.preview.setText("Detecting — please wait…")

        self._worker = _DetectWorker(excel_path, self.detect_kind)
        self._thread = QThread(parent=self)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._on_detected)
        self._worker.error.connect(self._on_detect_error)
        self._worker.finished.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._worker.finished.connect(self._worker.deleteLater)
        self._worker.error.connect(self._worker.deleteLater)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.finished.connect(
            lambda t=self._thread: self._threads.remove(t) if t in self._threads else None)
        self._threads.append(self._thread)
        self._thread.start()

    def _on_detected(self, result: dict):
        self.detected_info = result
        self.input.setText(result["ref"])
        self.preview.setText(
            f"Detected: Sheet={result['sheet']} | Ref={result['ref']} | "
            f"Column={result['col_letter']} | Column No={result['col_index']} | Header={result['header']}"
        )
        self.button.setText("Auto Detect")
        self.button.setEnabled(True)
        _log.info("AutoDetect success: kind=%s  sheet=%s  ref=%s  header=%r",
                  self.detect_kind, result.get("sheet"), result.get("ref"), result.get("header"))
        log_user_action("auto_detect_success",
                        f"ref={result.get('ref')} header={result.get('header')!r}",
                        extra=f"kind={self.detect_kind} | sheet={result.get('sheet')}")

    def _on_detect_error(self, msg: str):
        self.preview.setText("Detection failed — enter manually.")
        QMessageBox.critical(self, "Detection Error", msg)
        self.button.setText("Auto Detect")
        self.button.setEnabled(True)
        _log.warning("AutoDetect FAILED: kind=%s  error=%s", self.detect_kind, msg)

    def clear_selection(self):
        self.input.clear()
        self.preview.setText("Preview: not detected")
        self.detected_info = None

    def value(self) -> str:
        return normalize_excel_reference(self.input.text().strip())