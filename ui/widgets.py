"""
ui/widgets.py
─────────────
Reusable Qt widgets: NavButton, FolderField (with native OS dialog),
MultiFileField variants, StatChip, CollapsiblePanel, PremiumCard, etc.

KEY CHANGES vs original:
  • FolderField single-select  → QFileDialog.getExistingDirectory (native OS dialog)
  • FolderField multi-select   → Custom list dialog that internally uses
      QFileDialog.getExistingDirectory for each folder add (native dialog inside)
  • XlsxMultiFileField         → QFileDialog.getOpenFileNames (already native)
"""

import os
from PySide6.QtCore import Qt, QSize, QTimer, QPropertyAnimation, QEasingCurve, Signal
from core.logger import get_logger, log_file_upload, log_user_action

_log = get_logger(__name__)
from PySide6.QtGui import QColor, QIcon, QPainter, QPainterPath
from PySide6.QtWidgets import (
    QWidget, QFrame, QLabel, QPushButton, QProgressBar,
    QVBoxLayout, QHBoxLayout, QTextEdit, QLineEdit, QListWidget,
    QSizePolicy, QGraphicsDropShadowEffect, QFileDialog,
    QDialog, QAbstractItemView, QScrollArea
)
from core.utils import normalize_path, summarize_paths


# ── Qt shadow helper ─────────────────────────────────────────────────────────
def add_shadow(widget: QWidget, blur: int = 18, y_offset: int = 5, alpha: int = 80):
    shadow = QGraphicsDropShadowEffect(widget)
    shadow.setBlurRadius(blur)
    shadow.setOffset(0, y_offset)
    shadow.setColor(QColor(0, 0, 0, alpha))
    widget.setGraphicsEffect(shadow)
    return shadow


# ── ShimmerLabel ─────────────────────────────────────────────────────────────
class ShimmerLabel(QLabel):
    def __init__(self, text="Loading", parent=None):
        super().__init__(text, parent)
        self._dots = 0
        self._base = text
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)

    def start(self, text=None):
        if text:
            self._base = text
        self._dots = 0
        self.setText(self._base)
        self._timer.start(240)

    def stop(self, text=None):
        self._timer.stop()
        self.setText(text or self._base)

    def _tick(self):
        self._dots = (self._dots + 1) % 4
        self.setText(self._base + "." * self._dots)


# ── NavButton ────────────────────────────────────────────────────────────────
class NavButton(QPushButton):
    """Sidebar navigation button with icon + word-wrapped text.
    Icon and text are both placed inside a QHBoxLayout so they never overlap.
    Height auto-adjusts: 34 px single-line, 52 px two-line.
    """
    _FIXED_W = 196   # all nav buttons share this width

    def __init__(self, text: str, icon: QIcon):
        super().__init__()          # NO text/icon on the QPushButton itself
        self.setCheckable(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setFixedWidth(self._FIXED_W)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        # Do NOT call setIcon() — we render the icon ourselves in the layout

        from PySide6.QtWidgets import QHBoxLayout, QLabel as _QL
        from PySide6.QtGui import QPixmap

        self._inner_layout = QHBoxLayout(self)
        self._inner_layout.setContentsMargins(10, 4, 10, 4)
        self._inner_layout.setSpacing(8)

        # Icon label
        self._icon_label = _QL()
        self._icon_label.setFixedSize(18, 18)
        self._icon_label.setAlignment(Qt.AlignCenter)
        self._icon_label.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self._icon_label.setStyleSheet("background: transparent;")
        self._set_icon_pixmap(icon)
        self._inner_layout.addWidget(self._icon_label, 0, Qt.AlignVCenter)

        # Text label — word-wraps for long names
        self._text_label = _QL(text)
        self._text_label.setWordWrap(True)
        self._text_label.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        self._text_label.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self._text_label.setStyleSheet(
            "background: transparent; font-weight: 800; font-size: 13px;"
        )
        self._inner_layout.addWidget(self._text_label, 1)

        h = 60  # Uniform height for all buttons
        self.setFixedHeight(h)

        self._shadow = add_shadow(self, blur=14, y_offset=3, alpha=34)
        self._current_icon = icon

        # Update label color when checked state changes
        self.toggled.connect(self._on_toggled)

    def _on_toggled(self, checked: bool):
        """Keep text label color in sync with checked (active) state."""
        if checked:
            # Walk up to find the main window and read the actual accent color
            import re as _re
            accent_hex = None
            parent = self.parent()
            while parent is not None:
                if hasattr(parent, "accent_color"):
                    accent_hex = parent.accent_color.name()
                    break
                parent = parent.parent()
            # Fallback: parse from current global stylesheet
            if accent_hex is None:
                from PySide6.QtWidgets import QApplication
                ss = QApplication.instance().styleSheet() if QApplication.instance() else ""
                hits = _re.findall(r'#[0-9A-Fa-f]{6}', ss)
                accent_hex = hits[0] if hits else "#3BA8FF"
            from PySide6.QtGui import QColor as _QC
            bg = _QC(accent_hex)
            lum = (0.299 * bg.red() + 0.587 * bg.green() + 0.114 * bg.blue()) / 255
            color = "#1a1a1a" if lum > 0.55 else "#FFFFFF"
        else:
            color = ""   # empty = inherit from stylesheet
        self._text_label.setStyleSheet(
            f"background: transparent; font-weight: 800; font-size: 13px; color: {color};"
        )

    def _set_icon_pixmap(self, icon: QIcon):
        if icon and not icon.isNull():
            pm = icon.pixmap(QSize(16, 16))
            self._icon_label.setPixmap(pm)

    def setIcon(self, icon: QIcon):
        """Override so rebuild_icons() still works."""
        self._current_icon = icon
        self._set_icon_pixmap(icon)

    def icon(self) -> QIcon:
        return self._current_icon

    def setText(self, text: str):
        self._text_label.setText(text)
        h = 60  # Uniform height for all buttons
        self.setFixedHeight(h)

    def text(self) -> str:
        return self._text_label.text()

    def enterEvent(self, event):
        try:
            if self._shadow:
                self._shadow.setBlurRadius(18)
                self._shadow.setOffset(0, 4)
        except RuntimeError:
            self._shadow = None
        super().enterEvent(event)

    def leaveEvent(self, event):
        try:
            if self._shadow:
                self._shadow.setBlurRadius(14)
                self._shadow.setOffset(0, 3)
        except RuntimeError:
            self._shadow = None
        super().leaveEvent(event)




# ── IconTextButton ────────────────────────────────────────────────────────────
class IconTextButton(QPushButton):
    def __init__(self, text: str, icon: QIcon):
        super().__init__(text)
        self.setCursor(Qt.PointingHandCursor)
        self.setIcon(icon)
        self.setIconSize(QSize(16, 16))
        self.setMinimumHeight(40)
        self.setMinimumWidth(150)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)


# ── ProgressButton ───────────────────────────────────────────────────────────
class ProgressButton(QPushButton):
    """A button that shows a left-to-right fill animation while busy."""
    def __init__(self, text: str, icon: QIcon):
        super().__init__(text)
        self.setCursor(Qt.PointingHandCursor)
        self.setIcon(icon)
        self.setIconSize(QSize(16, 16))
        self.setMinimumHeight(40)
        self.setMinimumWidth(150)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self._progress   = 0          # 0-100
        self._fill_color = QColor("#1565C0")   # fill color (blue)
        self._text_color = QColor("#ffffff")
        self._timer      = QTimer(self)
        self._timer.setInterval(30)
        self._timer.timeout.connect(self._tick)
        self._direction  = 1          # 1 = filling, -1 = draining

    # ── public API ───────────────────────────────────────────────────────────
    def start_progress(self, label: str = ""):
        """Start the fill animation and update label."""
        if label:
            self.setText(label)
        self._progress  = 0
        self._direction = 1
        self.setEnabled(False)
        self._timer.start()
        self.update()

    def stop_progress(self, label: str = ""):
        """Stop the animation, restore label."""
        self._timer.stop()
        self._progress = 0
        if label:
            self.setText(label)
        self.setEnabled(True)
        self.update()

    # ── internal ─────────────────────────────────────────────────────────────
    def _tick(self):
        self._progress += self._direction * 2
        if self._progress >= 95:
            # Hold near-full until stop_progress is called
            self._progress = 95
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        r = self.rect()
        radius = 8

        # Read accent color dynamically from the button palette (set by stylesheet)
        # Falls back to blue if no color is set
        btn_bg = self.palette().button().color()
        if not btn_bg.isValid() or btn_bg == QColor(0, 0, 0) or btn_bg.lightness() < 5:
            btn_bg = QColor("#2196F3")

        # Background: use accent when enabled, slightly darker when disabled
        bg = btn_bg if self.isEnabled() else btn_bg.darker(120)
        path = QPainterPath()
        path.addRoundedRect(r.x(), r.y(), r.width(), r.height(), radius, radius)
        painter.fillPath(path, bg)

        # Fill overlay: slightly lighter shade of the same accent
        if self._progress > 0:
            fill_color = btn_bg.lighter(130)
            fill_w = int(r.width() * self._progress / 100)
            fill_path = QPainterPath()
            fill_path.addRoundedRect(r.x(), r.y(), fill_w, r.height(), radius, radius)
            painter.fillPath(fill_path, fill_color)

        # Text — auto-contrast: dark text on light accents, white on dark accents
        r2, g2, b2 = btn_bg.red(), btn_bg.green(), btn_bg.blue()
        lum = (0.299 * r2 + 0.587 * g2 + 0.114 * b2) / 255
        text_color = QColor("#1a1a1a") if lum > 0.55 else QColor("#ffffff")
        painter.setPen(text_color)
        font = self.font()
        painter.setFont(font)
        painter.drawText(r, Qt.AlignCenter, self.text())
        painter.end()


# ── SectionTitle ─────────────────────────────────────────────────────────────
class SectionTitle(QWidget):
    def __init__(self, title: str, subtitle: str = ""):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)
        title_label = QLabel(title)
        title_label.setObjectName("sectionTitle")
        layout.addWidget(title_label)
        if subtitle:
            sub = QLabel(subtitle)
            sub.setObjectName("sectionSubtitle")
            sub.setWordWrap(True)
            layout.addWidget(sub)


# ── StatChip ──────────────────────────────────────────────────────────────────
class StatChip(QFrame):
    def __init__(self, label: str, value: str, tone: str = "neutral"):
        super().__init__()
        self.setObjectName("statChip")
        self.setProperty("tone", tone)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(2)
        self.value_label = QLabel(value)
        self.value_label.setObjectName("statChipValue")
        self.name_label = QLabel(label)
        self.name_label.setObjectName("statChipLabel")
        self.name_label.setWordWrap(True)
        layout.addWidget(self.value_label)
        layout.addWidget(self.name_label)

    def set_value(self, value: str):
        self.value_label.setText(value)


# ── CollapsiblePanel ──────────────────────────────────────────────────────────
class CollapsiblePanel(QFrame):
    def __init__(self, title: str, content_widget: QWidget, expanded: bool = False):
        super().__init__()
        self.setObjectName("collapsiblePanel")
        self._content = content_widget
        self._expanded = expanded
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        self.toggle_btn = QPushButton(title)
        self.toggle_btn.setObjectName("collapseToggle")
        self.toggle_btn.setCheckable(True)
        self.toggle_btn.setChecked(expanded)
        self.toggle_btn.clicked.connect(self.set_expanded)
        layout.addWidget(self.toggle_btn)
        layout.addWidget(self._content)
        self.set_expanded(expanded)

    def set_expanded(self, expanded: bool):
        self._expanded = bool(expanded)
        self.toggle_btn.setChecked(self._expanded)
        self.toggle_btn.setText(("▼ " if self._expanded else "▶ ") + self.toggle_btn.text().lstrip("▼▶ "))
        self._content.setVisible(self._expanded)


# ── PremiumCard ───────────────────────────────────────────────────────────────
class PremiumCard(QFrame):
    def __init__(self, title, subtitle, icon, accent, button_text=None, button_callback=None):
        super().__init__()
        self.setObjectName("premiumCard")
        self._shadow = add_shadow(self, blur=28, y_offset=9, alpha=80)
        self._hover_anim = QPropertyAnimation(self, b"maximumHeight", self)
        self._hover_anim.setDuration(150)
        self._hover_anim.setEasingCurve(QEasingCurve.OutCubic)
        self.setMinimumHeight(200)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(18, 16, 18, 16)
        body_layout.setSpacing(10)
        top = QHBoxLayout()
        top.setSpacing(12)
        self.icon_box = QLabel()
        self.icon_box.setFixedSize(56, 56)
        self.icon_box.setAlignment(Qt.AlignCenter)
        self.icon_box.setPixmap(icon.pixmap(26, 26))
        title_label = QLabel(title)
        title_label.setObjectName("cardTitle")
        title_label.setWordWrap(True)
        top.addWidget(self.icon_box, alignment=Qt.AlignTop)
        top.addWidget(title_label, 1, alignment=Qt.AlignVCenter)
        sub = QLabel(subtitle)
        sub.setObjectName("cardSubtitle")
        sub.setWordWrap(True)
        body_layout.addLayout(top)
        body_layout.addWidget(sub)
        body_layout.addStretch()
        self.action_btn = None
        if button_text and button_callback:
            self.action_btn = QPushButton(button_text)
            self.action_btn.setObjectName("smallPrimaryButton")
            self.action_btn.clicked.connect(button_callback)
            body_layout.addWidget(self.action_btn, alignment=Qt.AlignLeft)
        layout.addWidget(body)
        self.update_accent(accent)

    def update_accent(self, accent: str):
        from PySide6.QtGui import QColor
        self.icon_box.setStyleSheet(f"background:{accent}; border-radius:16px;")
        if self.action_btn:
            c = QColor(accent)
            lum = (0.299 * c.red() + 0.587 * c.green() + 0.114 * c.blue()) / 255
            btn_txt = "#1a1a1a" if lum > 0.55 else "white"
            self.action_btn.setStyleSheet(f"""
                QPushButton {{
                    background: {accent}; color: {btn_txt}; border: 1px solid {accent};
                    border-radius: 12px; min-height: 34px; min-width: 104px;
                    padding: 6px 16px; font-weight: 900; text-align: center;
                }}
                QPushButton:hover {{ background: {accent}; border: 1px solid {accent}; }}
            """)

    def enterEvent(self, event):
        try:
            if self._shadow:
                self._shadow.setBlurRadius(36)
                self._shadow.setOffset(0, 12)
        except RuntimeError:
            self._shadow = None
        super().enterEvent(event)

    def leaveEvent(self, event):
        try:
            if self._shadow:
                self._shadow.setBlurRadius(28)
                self._shadow.setOffset(0, 9)
        except RuntimeError:
            self._shadow = None
        super().leaveEvent(event)


# ── FolderField ───────────────────────────────────────────────────────────────
# CHANGE 1: single-select now uses native OS file dialog directly.
# CHANGE 2: multi-select opens a custom list dialog, but each individual
#           folder is picked via native QFileDialog.getExistingDirectory.
class FolderField(QWidget):
    selectionChanged = Signal()  # emitted whenever the selection changes (point 6)

    def __init__(self, label_text, button_text, button_icon, multi=False,
                 multi_line_display=False, show_clear_btn=False):
        super().__init__()
        self.multi = multi
        self.multi_line_display = multi_line_display
        self.selected_paths: list[str] = []

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        self.label = QLabel(label_text)
        self.label.setObjectName("fieldLabel")
        layout.addWidget(self.label)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(10)

        if self.multi_line_display:
            # Multi-select: tall QTextEdit box
            self.display = QTextEdit()
            self.display.setReadOnly(True)
            self.display.setPlaceholderText("No folder selected")
            self.display.setFixedHeight(120)
            self.display.setLineWrapMode(QTextEdit.NoWrap)
        else:
            # Single-select: slim QLineEdit
            self.display = QLineEdit()
            self.display.setReadOnly(True)
            self.display.setPlaceholderText("No folder selected")
            self.display.setFixedHeight(38)

        self.button = IconTextButton(button_text, button_icon)
        self.button.setObjectName("pickerButton")
        self.button.setFixedSize(136, 38)
        self.button.clicked.connect(self.pick_folder)

        row.addWidget(self.display, 1)

        if self.multi:
            # ── Reference Bases: vertical stack — Add Folder above Clear ──────
            btn_col = QVBoxLayout()
            btn_col.setSpacing(6)
            btn_col.setContentsMargins(0, 0, 0, 0)

            self.clear_btn = QPushButton("✕ Clear")
            self.clear_btn.setObjectName("clearButton")
            self.clear_btn.setFixedSize(136, 38)
            self.clear_btn.clicked.connect(self.clear_selection)

            btn_col.addWidget(self.button)      # Add Folder on top
            btn_col.addWidget(self.clear_btn)   # Clear below
            btn_col.addStretch()
            row.addLayout(btn_col)

        elif show_clear_btn:
            # ── Target Base: Upload Folder + Clear side by side ───────────────
            self.clear_btn = QPushButton("✕ Clear")
            self.clear_btn.setObjectName("clearButton")
            self.clear_btn.setFixedSize(80, 38)
            self.clear_btn.clicked.connect(self.clear_selection)

            row.addWidget(self.button, 0, Qt.AlignVCenter)
            row.addWidget(self.clear_btn, 0, Qt.AlignVCenter)

        layout.addLayout(row)

        if self.multi:
            self.count_label = QLabel("0 folders selected")
            self.count_label.setObjectName("panelSubtitle")
            layout.addWidget(self.count_label)

    # ── single-select: native OS folder dialog ───────────────────────────────
    def _pick_single_folder_native(self, title: str = "Select Folder") -> str:
        """Opens the operating-system's native folder browser directly."""
        folder = QFileDialog.getExistingDirectory(None, title, "",
                    QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks)
        return folder

    # ── multi-select: native OS folder picker accumulator ────────────────────
    def _pick_multi_folders_dialog(self, title: str = "Select Reference Folders"):
        """
        Opens the native folder picker.
        Returns a list of the RAW paths the user just selected,
        or None if the user cancelled.  Merging with existing
        selected_paths is handled entirely by pick_folder().
        """
        chosen = self._windows_shell_multi_folder(title)
        if chosen is None:
            # Fallback: Qt dialog (non-Windows or COM failure)
            chosen = self._qt_multi_folder_fallback(title)
        # None means cancelled — propagate None so pick_folder can bail out
        return chosen

    def _windows_shell_multi_folder(self, title: str):
        """
        Use Windows IFileOpenDialog COM interface directly via ctypes.
        Returns list of selected folder paths, or None on cancel/error.
        """
        import sys
        if sys.platform != "win32":
            return None
        try:
            import ctypes
            import ctypes.wintypes

            # COM GUIDs
            CLSID_FileOpenDialog = ctypes.c_byte * 16
            clsid = CLSID_FileOpenDialog(
                0xDC, 0x1C, 0x5A, 0x9C, 0xE4, 0x79, 0xAC, 0x4F,
                0x94, 0x78, 0x70, 0x6D, 0xC6, 0xA1, 0x1F, 0x3E)
            iid_IFileOpenDialog = CLSID_FileOpenDialog(
                0xD5, 0x7A, 0xBA, 0xD7, 0x12, 0x18, 0x63, 0x43,
                0xA8, 0xD3, 0x3B, 0x1A, 0xF4, 0x7A, 0x07, 0x70)

            ole32 = ctypes.windll.ole32
            ole32.CoInitialize(None)

            dialog = ctypes.c_void_p()
            hr = ole32.CoCreateInstance(
                ctypes.byref((ctypes.c_byte * 16)(*clsid)),
                None, 1,
                ctypes.byref((ctypes.c_byte * 16)(*iid_IFileOpenDialog)),
                ctypes.byref(dialog)
            )
            if hr != 0:
                return None

            # vtable layout for IFileOpenDialog (inherits IFileDialog → IModalWindow)
            vtable = ctypes.cast(dialog, ctypes.POINTER(ctypes.c_void_p))
            vtable_ptr = ctypes.cast(vtable[0], ctypes.POINTER(ctypes.c_void_p))

            # IModalWindow::Show  = vtable[3]
            # IFileDialog::SetOptions = vtable[9]
            # IFileDialog::GetOptions = vtable[8]
            # IFileDialog::SetTitle   = vtable[17]
            # IFileOpenDialog::GetResults = vtable[27]
            # IShellItemArray::GetCount   / GetItemAt

            FOS_PICKFOLDERS      = 0x00000020
            FOS_ALLOWMULTISELECT = 0x00000200
            FOS_FORCEFILESYSTEM  = 0x00000040

            # GetOptions
            GetOptions = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p,
                                            ctypes.POINTER(ctypes.c_uint32))
            get_options = GetOptions(vtable_ptr[8])
            opts = ctypes.c_uint32(0)
            get_options(dialog, ctypes.byref(opts))

            # SetOptions
            SetOptions = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p,
                                            ctypes.c_uint32)
            set_options = SetOptions(vtable_ptr[9])
            set_options(dialog,
                        opts.value | FOS_PICKFOLDERS | FOS_ALLOWMULTISELECT | FOS_FORCEFILESYSTEM)

            # SetTitle
            SetTitle = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p,
                                          ctypes.c_wchar_p)
            SetTitle(vtable_ptr[17])(dialog, title)

            # Show (blocks until user closes)
            hwnd = ctypes.windll.user32.GetForegroundWindow()
            Show = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p,
                                      ctypes.wintypes.HWND)
            hr = Show(vtable_ptr[3])(dialog, hwnd)
            if hr != 0:  # S_OK = 0; HRESULT_FROM_WIN32(ERROR_CANCELLED) = 0x800704C7
                return None  # user cancelled

            # GetResults → IShellItemArray
            iid_IShellItemArray = (ctypes.c_byte * 16)(
                0xB6, 0x3E, 0xA7, 0xB1, 0x26, 0x32, 0xD8, 0x4E,
                0xAA, 0x44, 0x1D, 0x09, 0xAF, 0x9F, 0x49, 0xFC)
            GetResults = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p,
                                            ctypes.POINTER(ctypes.c_void_p))
            item_array = ctypes.c_void_p()
            hr = GetResults(vtable_ptr[27])(dialog, ctypes.byref(item_array))
            if hr != 0 or not item_array:
                return None

            arr_vtable = ctypes.cast(
                ctypes.cast(item_array, ctypes.POINTER(ctypes.c_void_p))[0],
                ctypes.POINTER(ctypes.c_void_p))

            # IShellItemArray::GetCount = vtable[4]
            GetCount = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p,
                                          ctypes.POINTER(ctypes.c_uint32))
            count = ctypes.c_uint32(0)
            GetCount(arr_vtable[4])(item_array, ctypes.byref(count))

            # IShellItemArray::GetItemAt = vtable[5]
            GetItemAt = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p,
                                           ctypes.c_uint32, ctypes.POINTER(ctypes.c_void_p))

            SIGDN_FILESYSPATH = ctypes.c_int(0x80058000)
            results = []
            for i in range(count.value):
                item = ctypes.c_void_p()
                hr = GetItemAt(arr_vtable[5])(item_array, i, ctypes.byref(item))
                if hr != 0 or not item:
                    continue
                item_vtable = ctypes.cast(
                    ctypes.cast(item, ctypes.POINTER(ctypes.c_void_p))[0],
                    ctypes.POINTER(ctypes.c_void_p))
                # IShellItem::GetDisplayName = vtable[4]
                GetDisplayName = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p,
                                                    ctypes.c_int, ctypes.POINTER(ctypes.c_wchar_p))
                path_ptr = ctypes.c_wchar_p()
                hr = GetDisplayName(item_vtable[4])(item, SIGDN_FILESYSPATH, ctypes.byref(path_ptr))
                if hr == 0 and path_ptr.value:
                    results.append(path_ptr.value)
                # Release IShellItem
                Release = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
                Release(item_vtable[2])(item)

            # Release IShellItemArray and dialog
            arr_release_vt = ctypes.cast(
                ctypes.cast(item_array, ctypes.POINTER(ctypes.c_void_p))[0],
                ctypes.POINTER(ctypes.c_void_p))
            Release = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
            Release(arr_release_vt[2])(item_array)
            Release(vtable_ptr[2])(dialog)

            return results if results else None

        except Exception:
            return None

    def _qt_multi_folder_fallback(self, title: str):
        """
        Non-Windows fallback: Qt's own folder dialog with multi-select enabled
        on its internal views.  Returns list of paths or None on cancel.
        """
        from PySide6.QtWidgets import QListView, QTreeView

        dlg = QFileDialog(None, title)
        dlg.setFileMode(QFileDialog.FileMode.Directory)
        dlg.setOption(QFileDialog.Option.DontUseNativeDialog, False)
        dlg.setOption(QFileDialog.Option.ShowDirsOnly, True)

        if self.selected_paths:
            start_dir = os.path.dirname(self.selected_paths[0])
            if os.path.isdir(start_dir):
                dlg.setDirectory(start_dir)

        for view in dlg.findChildren(QListView):
            view.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        for view in dlg.findChildren(QTreeView):
            view.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)

        if dlg.exec() != QFileDialog.DialogCode.Accepted:
            return None
        chosen = [p for p in dlg.selectedFiles() if os.path.isdir(p)]
        return chosen if chosen else None

    def pick_folder(self):
        if self.multi:
            old_paths = list(self.selected_paths)
            old_set   = set(old_paths)

            # Returns only what the user just picked, or None if cancelled
            raw_picks = self._pick_multi_folders_dialog("Select Reference Folders")

            # User cancelled — do nothing
            if raw_picks is None:
                return

            # Normalize all picked paths
            picked = [normalize_path(p) for p in raw_picks]

            # Split into truly-new vs already-in-list (case-insensitive on Windows)
            old_set_lower = {p.lower() for p in old_paths}
            added = [p for p in picked if p.lower() not in old_set_lower]
            dups  = [p for p in picked if p.lower() in old_set_lower]

            if dups and not added:
                # Every pick was already in the list — nothing to add
                from PySide6.QtWidgets import QMessageBox
                names = "\n".join(os.path.basename(p) or p for p in dups)
                QMessageBox.warning(
                    None, "Already Added",
                    f"The following folder(s) are already in the list — no changes made:\n\n{names}"
                )
                return  # don't update
            elif dups:
                # Some new, some duplicate — add new ones, warn about skipped
                from PySide6.QtWidgets import QMessageBox
                names = "\n".join(os.path.basename(p) or p for p in dups)
                QMessageBox.information(
                    None, "Some Folders Already Added",
                    f"The following folder(s) were already in the list and were skipped:\n\n{names}"
                )

            # Merge: keep old order, append only new picks
            self.selected_paths = old_paths + added
            self._update_display()
            self.selectionChanged.emit()
            # ── Log every newly added folder ─────────────────────────────────
            label = getattr(self, "label_text", "FolderField")
            for p in added:
                log_file_upload("folder", p, field=label)
            _log.debug("FolderField(multi): %d new folder(s) added, %d total",
                       len(added), len(self.selected_paths))
        else:
            # CHANGE: single select uses native OS dialog directly (no custom dialog)
            folder = self._pick_single_folder_native("Select Folder")
            if folder:
                self.selected_paths = [normalize_path(folder)]
                self._update_display()
                self.selectionChanged.emit()
                # ── Log the single folder pick ────────────────────────────────
                label = getattr(self, "label_text", "FolderField")
                log_file_upload("folder", self.selected_paths[0], field=label)
                _log.debug("FolderField(single): selected %s", self.selected_paths[0])

    def _update_display(self):
        if self.multi:
            text = summarize_paths(self.selected_paths, "folder")
            if isinstance(self.display, QTextEdit):
                self.display.setPlainText(text)
            else:
                self.display.setText(text.replace("\n", " | "))
            if hasattr(self, "count_label"):
                count = len(self.selected_paths)
                self.count_label.setText(f"{count} folder{'s' if count != 1 else ''} selected")
        else:
            value = self.selected_paths[0] if self.selected_paths else ""
            if isinstance(self.display, QTextEdit):
                self.display.setPlainText(value)
            else:
                self.display.setText(value)

    def clear_selection(self):
        self.selected_paths = []
        if isinstance(self.display, QTextEdit):
            self.display.clear()
        else:
            self.display.setText("")
        if hasattr(self, "count_label"):
            self.count_label.setText("0 folders selected")
        self.selectionChanged.emit()  # point 6

    def value(self) -> list:
        return self.selected_paths


# ── MultiFileField (base) ─────────────────────────────────────────────────────
class MultiFileField(QWidget):
    """Base multi-file picker — slim single-line display with Choose Files + Clear side by side."""
    selectionChanged = Signal()  # emitted when files are picked or cleared (point 6)

    def __init__(self, label_text, button_text, button_icon, filter_text, allowed_exts):
        super().__init__()
        self.selected_paths: list[str] = []
        self.filter_text = filter_text
        self.allowed_exts = tuple(ext.lower() for ext in allowed_exts)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        self.label = QLabel(label_text)
        self.label.setObjectName("fieldLabel")
        layout.addWidget(self.label)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(10)

        # ── Single-line display (mirrors Target Base Folder) ──────────────────
        self.display = QLineEdit()
        self.display.setReadOnly(True)
        self.display.setPlaceholderText("No files selected")
        self.display.setFixedHeight(38)

        # ── Choose Files button ───────────────────────────────────────────────
        self.button = IconTextButton(button_text, button_icon)
        self.button.setObjectName("pickerButton")
        self.button.setFixedSize(136, 38)
        self.button.clicked.connect(self.pick_files)

        # ── Clear button ──────────────────────────────────────────────────────
        self.clear_btn = QPushButton("✕ Clear")
        self.clear_btn.setObjectName("clearButton")
        self.clear_btn.setFixedSize(80, 38)
        self.clear_btn.clicked.connect(self.clear_selection)

        row.addWidget(self.display, 1)
        row.addWidget(self.button, 0, Qt.AlignVCenter)
        row.addWidget(self.clear_btn, 0, Qt.AlignVCenter)
        layout.addLayout(row)

    def pick_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Files", "", self.filter_text)
        if not files:
            return
        cleaned = [normalize_path(p) for p in files if p.lower().endswith(self.allowed_exts)]
        if not cleaned:
            return

        existing_set = set(self.selected_paths)
        already_selected = [p for p in cleaned if p in existing_set]
        new_files        = [p for p in cleaned if p not in existing_set]

        if already_selected:
            names = "\n".join(os.path.basename(p) for p in already_selected)
            if not new_files:
                from PySide6.QtWidgets import QMessageBox
                QMessageBox.warning(
                    self, "Already Selected",
                    f"The following file(s) are already selected — no changes made:\n\n{names}"
                )
                return
            else:
                from PySide6.QtWidgets import QMessageBox
                QMessageBox.information(
                    self, "Some Files Already Selected",
                    f"The following file(s) were already selected and were skipped:\n\n{names}"
                )

        self.selected_paths = self.selected_paths + new_files
        self.display.setText(summarize_paths(self.selected_paths, "file").replace("\n", " | "))
        self.selectionChanged.emit()
        # ── Log every newly selected file ─────────────────────────────────────
        label = getattr(self, "label_text", "MultiFileField")
        for p in new_files:
            log_file_upload("file", p, field=label)
        _log.debug("MultiFileField: %d new file(s) added, %d total",
                   len(new_files), len(self.selected_paths))

    def clear_selection(self):
        self.selected_paths = []
        self.display.clear()
        self.selectionChanged.emit()  # point 6

    def value(self) -> list:
        return self.selected_paths


class TxtMultiFileField(MultiFileField):
    """Function list files — xlsx only. Blue border on display box."""
    def __init__(self, label_text, button_text, button_icon):
        super().__init__(
            label_text, button_text, button_icon,
            "Excel Files (*.xlsx);;All Files (*.*)",
            (".xlsx",)
        )
        # Blue border — visually marks the Function List source box
        self.display.setStyleSheet(
            "QLineEdit { border: 2px solid #1E90FF; border-radius: 8px;"
            " padding: 4px 8px; background: transparent; }"
            "QLineEdit:focus { border-color: #4DB6FF; }"
        )


class XlsxMultiFileField(MultiFileField):
    def __init__(self, label_text, button_text, button_icon):
        super().__init__(
            label_text, button_text, button_icon,
            "Excel Files (*.xlsx);;All Files (*.*)",
            (".xlsx",)
        )


# ── TargetFolderInputField ────────────────────────────────────────────────────
class TargetFolderInputField(QWidget):
    """Single folder picker — slim single-row layout.
    Upload Folder + Clear are placed side-by-side (parallel) matching the
    other input fields on the Consolidated Database page."""

    def __init__(self, label_text, button_icon):
        super().__init__()
        self.selected_path: str = ""

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        self.label = QLabel(label_text)
        self.label.setObjectName("fieldLabel")
        layout.addWidget(self.label)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(10)

        # ── Slim single-line display — same height as the other fields ────────
        self.display = QLineEdit()
        self.display.setReadOnly(True)
        self.display.setPlaceholderText("No folder selected — will scan all sub-folders")
        self.display.setFixedHeight(38)          # reduced from 120 → matches MultiFileField

        # ── Upload Folder + Clear side-by-side (parallel) ─────────────────────
        self.button = IconTextButton("Upload Folder", button_icon)
        self.button.setObjectName("pickerButton")
        self.button.setFixedSize(136, 38)
        self.button.clicked.connect(self.pick_folder)

        self.clear_btn = QPushButton("✕ Clear")
        self.clear_btn.setObjectName("clearButton")
        self.clear_btn.setFixedSize(80, 38)
        self.clear_btn.clicked.connect(self.clear_selection)

        row.addWidget(self.display, 1)
        row.addWidget(self.button, 0, Qt.AlignVCenter)
        row.addWidget(self.clear_btn, 0, Qt.AlignVCenter)
        layout.addLayout(row)

    def pick_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Select Target Folder", "",
            QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks
        )
        if folder:
            self.selected_path = normalize_path(folder)
            self.display.setText(self.selected_path)
            log_file_upload("folder", self.selected_path, field="Target Folder")
            _log.debug("TargetFolderInputField: selected %s", self.selected_path)

    def clear_selection(self):
        self.selected_path = ""
        self.display.clear()

    def value(self) -> str:
        return self.selected_path


# ── ExcelFileField ────────────────────────────────────────────────────────────
class ExcelFileField(QWidget):
    """Single Excel file picker — slim single-row layout matching MultiFileField.
    Green border marks this as the Consolidated DB source.
    Upload Excel + Clear are placed side-by-side (parallel) to the right of the
    display box, exactly like the Function List Files field above it."""
    def __init__(self, label_text, button_text, button_icon):
        super().__init__()
        self.selected_path = ""

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        self.label = QLabel(label_text)
        self.label.setObjectName("fieldLabel")
        layout.addWidget(self.label)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(10)

        # ── Slim single-line display — same height as Function List box ───────
        self.display = QLineEdit()
        self.display.setReadOnly(True)
        self.display.setPlaceholderText("No Excel file selected")
        self.display.setFixedHeight(38)          # reduced from 120 → matches MultiFileField
        # Green border — visually marks the Consolidated DB source box
        self.display.setStyleSheet(
            "QLineEdit { border: 2px solid #27AE60; border-radius: 8px;"
            " padding: 4px 8px; background: transparent; }"
            "QLineEdit:focus { border-color: #2ECC71; }"
        )

        # ── Upload + Clear side-by-side (parallel) ────────────────────────────
        self.button = IconTextButton(button_text, button_icon)
        self.button.setObjectName("pickerButton")
        self.button.setFixedSize(136, 38)
        self.button.clicked.connect(self.pick_file)

        self.clear_btn = QPushButton("✕ Clear")
        self.clear_btn.setObjectName("clearButton")
        self.clear_btn.setFixedSize(80, 38)
        self.clear_btn.clicked.connect(self.clear_selection)

        row.addWidget(self.display, 1)
        row.addWidget(self.button, 0, Qt.AlignVCenter)
        row.addWidget(self.clear_btn, 0, Qt.AlignVCenter)
        layout.addLayout(row)

    def pick_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "",
            "Excel Files (*.xlsx *.xlsm *.xls);;All Files (*.*)"
        )
        if not file_path:
            return
        self.selected_path = normalize_path(file_path)
        self.display.setText(self.selected_path)
        log_file_upload("excel", self.selected_path, field="Excel File")
        _log.debug("ExcelFileField: selected %s", self.selected_path)

    def clear_selection(self):
        self.selected_path = ""
        self.display.clear()

    def value(self) -> str:
        return self.selected_path


# ── OutputLinkField ───────────────────────────────────────────────────────────
class OutputLinkField(QWidget):
    """Output path display. open_btn starts disabled (greyed) until set_output() is called.
    The open_btn is intentionally NOT added to this widget's layout —
    consolidated_page.py places it in the bottom button row instead (point 5)."""
    def __init__(self, label_text, button_icon):
        super().__init__()
        self.output_file_path = ""

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        self.label = QLabel(label_text)
        self.label.setObjectName("fieldLabel")
        layout.addWidget(self.label)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(10)

        self.link_label = QLabel("No output generated yet")
        self.link_label.setObjectName("panelSubtitle")
        self.link_label.setTextInteractionFlags(Qt.TextBrowserInteraction)
        self.link_label.setOpenExternalLinks(False)
        self.link_label.linkActivated.connect(self.open_link)

        # open_btn is created here but NOT added to this layout.
        # consolidated_page adds it to the bottom button row.
        self.open_btn = IconTextButton("Open Output", button_icon)
        self.open_btn.setObjectName("pickerButton")
        self.open_btn.setFixedSize(150, 44)
        self.open_btn.clicked.connect(self.open_output)
        self.open_btn.setEnabled(False)   # greyed until output exists (point 6)

        row.addWidget(self.link_label, 1)
        layout.addLayout(row)

    def set_output(self, path: str):
        from PySide6.QtCore import QUrl
        self.output_file_path = path
        file_url = QUrl.fromLocalFile(path).toString()
        self.link_label.setText(f'<a href="{file_url}">{path}</a>')
        self.open_btn.setEnabled(True)

    def open_link(self, url: str):
        if url:
            from PySide6.QtCore import QUrl
            from PySide6.QtGui import QDesktopServices
            QDesktopServices.openUrl(QUrl(url))

    def open_output(self):
        if self.output_file_path and os.path.exists(self.output_file_path):
            from PySide6.QtCore import QUrl
            from PySide6.QtGui import QDesktopServices
            QDesktopServices.openUrl(QUrl.fromLocalFile(self.output_file_path))

    def clear_selection(self):
        self.output_file_path = ""
        self.link_label.setText("No output generated")
        self.open_btn.setEnabled(False)


# ── StepStatusWidget ──────────────────────────────────────────────────────────
class StepStatusWidget(QWidget):
    STATE_PENDING = "pending"
    STATE_RUNNING = "running"
    STATE_DONE    = "done"
    STATE_ERROR   = "error"

    # Card colors per state  (bg, border)
    _COLORS = {
        STATE_PENDING: ("#2A3347", "#3D5070"),
        STATE_RUNNING: ("#1E3A5F", "#1DA1F2"),
        STATE_DONE:    ("#1A3828", "#2EA043"),
        STATE_ERROR:   ("#3D2020", "#CF2020"),
    }

    def __init__(self, parent=None):
        super().__init__(parent)
        self._layout = QVBoxLayout(self)
        self._layout.setContentsMargins(0, 0, 0, 0)
        self._layout.setSpacing(5)
        self._rows: list[dict] = []

    def clear_steps(self):
        while self._layout.count():
            item = self._layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self._rows.clear()

    def add_step(self, label: str):
        step_num = len(self._rows) + 1

        # Outer card
        card = QFrame()
        card.setObjectName("stepRow")
        bg, border = self._COLORS[self.STATE_PENDING]
        card.setStyleSheet(
            f"QFrame#stepRow {{ background: {bg}; border: 1px solid {border};"
            f" border-radius: 10px; }}"
        )

        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(12, 8, 12, 8)
        card_layout.setSpacing(4)

        # ── Top row: number badge · icon · name · status text ─────────────────
        top_row = QHBoxLayout()
        top_row.setSpacing(8)
        top_row.setContentsMargins(0, 0, 0, 0)

        # Step number badge
        num_lbl = QLabel(str(step_num))
        num_lbl.setFixedSize(22, 22)
        num_lbl.setAlignment(Qt.AlignCenter)
        num_lbl.setStyleSheet(
            "background: #3D5070; color: #A0C4E0; border-radius: 11px;"
            " font-size: 10px; font-weight: 900;"
        )

        # State icon
        icon_lbl = QLabel("⏳")
        icon_lbl.setFixedWidth(20)
        icon_lbl.setAlignment(Qt.AlignCenter)
        icon_lbl.setStyleSheet("font-size: 14px; background: transparent;")

        # Step name
        name_lbl = QLabel(label)
        name_lbl.setStyleSheet(
            "color: #BDD3E4; font-weight: 700; font-size: 12px; background: transparent;"
        )
        name_lbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        # Status detail (right-aligned)
        status_lbl = QLabel("Waiting…")
        status_lbl.setStyleSheet(
            "color: #6A8DA3; font-size: 10px; background: transparent;"
        )
        status_lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        status_lbl.setMinimumWidth(140)
        status_lbl.setMaximumWidth(260)

        top_row.addWidget(num_lbl)
        top_row.addWidget(icon_lbl)
        top_row.addWidget(name_lbl, 1)
        top_row.addWidget(status_lbl)
        card_layout.addLayout(top_row)

        # ── Per-step progress bar ──────────────────────────────────────────────
        pbar = QProgressBar()
        pbar.setRange(0, 100)
        pbar.setValue(0)
        pbar.setFixedHeight(5)
        pbar.setTextVisible(False)
        pbar.setStyleSheet(
            "QProgressBar { background: #3D5070; border: none; border-radius: 2px; }"
            "QProgressBar::chunk { background: #3D5070; border-radius: 2px; }"
        )
        card_layout.addWidget(pbar)

        self._layout.addWidget(card)
        self._rows.append({
            "label":   label,
            "frame":   card,
            "icon":    icon_lbl,
            "status":  status_lbl,
            "pbar":    pbar,
            "num_lbl": num_lbl,
        })

    def _apply_card_style(self, row: dict, state: str):
        bg, border = self._COLORS.get(state, self._COLORS[self.STATE_PENDING])
        row["frame"].setStyleSheet(
            f"QFrame#stepRow {{ background: {bg}; border: 1px solid {border};"
            f" border-radius: 10px; }}"
        )

    def set_state(self, label: str, state: str, detail: str = "", pct: int = -1):
        for row in self._rows:
            if row["label"] == label:
                self._apply_card_style(row, state)

                if state == self.STATE_RUNNING:
                    row["icon"].setText("🔄")
                    row["status"].setText(detail or "Running…")
                    row["status"].setStyleSheet(
                        "color: #1DA1F2; font-size: 10px; background: transparent;"
                    )
                    row["num_lbl"].setStyleSheet(
                        "background: #1A4A70; color: #1DA1F2; border-radius: 11px;"
                        " font-size: 10px; font-weight: 900;"
                    )
                    # update pbar if pct provided
                    if 0 <= pct <= 100:
                        row["pbar"].setValue(pct)
                        row["pbar"].setStyleSheet(
                            "QProgressBar { background: #1A3A5F; border: none; border-radius: 2px; }"
                            "QProgressBar::chunk { background: #1DA1F2; border-radius: 2px; }"
                        )
                    else:
                        cur = row["pbar"].value()
                        if cur < 5:
                            row["pbar"].setValue(5)
                        row["pbar"].setStyleSheet(
                            "QProgressBar { background: #1A3A5F; border: none; border-radius: 2px; }"
                            "QProgressBar::chunk { background: #1DA1F2; border-radius: 2px; }"
                        )

                elif state == self.STATE_DONE:
                    row["icon"].setText("✅")
                    row["status"].setText(detail or "Done")
                    row["status"].setStyleSheet(
                        "color: #2EA043; font-size: 10px; background: transparent;"
                    )
                    row["num_lbl"].setStyleSheet(
                        "background: #1A4A30; color: #2EA043; border-radius: 11px;"
                        " font-size: 10px; font-weight: 900;"
                    )
                    row["pbar"].setValue(100)
                    row["pbar"].setStyleSheet(
                        "QProgressBar { background: #1A3828; border: none; border-radius: 2px; }"
                        "QProgressBar::chunk { background: #2EA043; border-radius: 2px; }"
                    )

                elif state == self.STATE_ERROR:
                    row["icon"].setText("❌")
                    row["status"].setText(detail or "Error")
                    row["status"].setStyleSheet(
                        "color: #CF2020; font-size: 10px; background: transparent;"
                    )
                    row["num_lbl"].setStyleSheet(
                        "background: #4A1A1A; color: #CF2020; border-radius: 11px;"
                        " font-size: 10px; font-weight: 900;"
                    )
                    row["pbar"].setStyleSheet(
                        "QProgressBar { background: #3A1A1A; border: none; border-radius: 2px; }"
                        "QProgressBar::chunk { background: #CF2020; border-radius: 2px; }"
                    )
                break

    def set_step_progress(self, label: str, pct: int):
        """Update just the progress bar of a running step (0-100)."""
        for row in self._rows:
            if row["label"] == label:
                row["pbar"].setValue(max(0, min(100, pct)))
                break