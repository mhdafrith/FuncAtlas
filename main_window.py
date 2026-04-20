"""
main_window.py
──────────────
ReuseAnalysisWindow — the central QMainWindow.
All page-builder logic lives in pages/*.py.
All heavy-lifting lives in services/*.py.
This file only wires navigation, theme, and business logic slots.
"""

import os
import re
import sys
from collections import OrderedDict

from PySide6.QtCore import Qt, QSize, QTimer, QPropertyAnimation, QEasingCurve, QThread
from PySide6.QtGui import QColor, QFont, QPixmap, QDesktopServices, QIcon
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFrame, QLabel, QPushButton,
    QVBoxLayout, QHBoxLayout, QStackedWidget, QScrollArea, QSizePolicy,
    QGraphicsOpacityEffect, QColorDialog, QFontDialog, QMessageBox,
    QTreeWidgetItem, QTableWidget, QTableWidgetItem, QHeaderView, QMenu,
    QDialog, QTextEdit, QFileDialog
)

from core.theme import ThemeManager, VectorIconFactory
from core.utils import (
    normalize_path, normalize_name, clean_text,
    extract_function_body, SCAN_CACHE, SCAN_CACHE_LOCK,
    is_valid_excel_reference
)
from ui.widgets import (
    NavButton, IconTextButton, SectionTitle, PremiumCard,
    StatChip, CollapsiblePanel, add_shadow
)
from ui.dialogs import HelpOverlayDialog, CompletionPopupDialog
from services.analysis import (
    scan_source_for_all_functions, match_target_with_function_list,
    match_target_with_reference_bases, merge_record_sets,
    parse_function_list_files, ConsolidatedWorker,
    extract_functions_from_folder_to_excel,
    detect_best_column_in_workbook
)
from services.report_worker import ReportCompareWorker
from services.analysis import BuiltinExtractionWorker

# Page builders
from pages.home import create_home_page
from pages.input_page import create_input_page
from pages.reference_page import create_reference_page
from pages.consolidated_page import create_consolidated_page
from pages.view_page import create_view_page
from pages.diff_page import create_diff_page
from pages.report_page import create_report_page
from pages.help_page import create_help_page
from pages.settings_page import create_settings_page
from pages.complexity_page import create_complexity_page


def _load_home_hero_pixmap() -> QPixmap:
    candidates = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "image.png"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)),
                     "Gemini_Generated_Image_ps5x2aps5x2aps5x (1).png"),
    ]
    for path in candidates:
        if os.path.exists(path):
            pm = QPixmap(path)
            if not pm.isNull():
                return pm
    return QPixmap()


def _find_function_in_folder(folder: str, function_name: str) -> list:
    matches = []
    for root, _, files in os.walk(folder):
        for file in files:
            if not file.endswith((".c", ".cpp", ".h", ".txt")):
                continue
            full_path = os.path.join(root, file)
            try:
                from core.utils import read_source_file as _read_src
                content = _read_src(full_path)
            except Exception:
                continue
            if re.compile(rf"\b{function_name}\s*\([^;]*\)\s*\{{").search(content):
                matches.append(full_path)
    return matches


# ─────────────────────────────────────────────────────────────────────────────
class ReuseAnalysisWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("FuncAtlas")
        self.resize(1500, 900)
        self.setMinimumSize(1240, 740)

        self.default_accent   = QColor("#3BA8FF")
        self.accent_color     = QColor(self.default_accent)
        self.base_font_family = "Segoe UI"
        self.base_font_size   = 12   # point 7: bigger font size throughout
        self.current_theme    = "light"
        self.theme            = ThemeManager.THEMES[self.current_theme]

        self.icons_white = VectorIconFactory(QColor("#FFFFFF"))
        self.icons       = self.icons_white
        self.nav_buttons: dict  = {}
        self.pages: dict        = {}
        self.accent_cards: list = []
        self.home_stat_chips    = []

        self.function_records           = OrderedDict()
        self.current_function_name      = ""
        self.current_function_file      = ""
        self.current_function_body      = ""
        self.current_function_list_paths = []
        self.available_sources          = []
        self.active_source_root         = ""
        self.current_reference_folders  = []
        self._ref_function_filter       = []
        self.con_thread = None
        self.con_worker = None
        self._active_threads: list[QThread] = []  # strong refs to all running threads

        self._report_ext_thread  = None
        self._report_ext_worker  = None
        self._report_cmp_thread  = None
        self._report_cmp_worker  = None
        self._report_output_file = ""
        self._last_report_excel  = ""
        self._last_complexity_html = ""
        self._report_bases_list  = []
        self._report_total_steps = 0
        self._report_done_steps  = 0

        self._setup_ui()
        self.apply_styles()
        self.rebuild_icons()
        self.show_page("home")

    # ── accent card registry ──────────────────────────────────────────────────
    def register_accent_card(self, card: PremiumCard) -> PremiumCard:
        self.accent_cards.append(card)
        return card

    # ── navigation helpers ────────────────────────────────────────────────────
    def wire_animated_navigation(self, button: QPushButton, destination: str):
        if not button:
            return
        try:
            button.clicked.disconnect()
        except Exception:
            pass
        button.clicked.connect(
            lambda checked=False, btn=button, dest=destination:
                self.animate_button_and_navigate(btn, dest)
        )

    def animate_button_and_navigate(self, button: QPushButton, destination: str):
        self.animate_button_click(button, lambda dest=destination: self.show_page(dest))

    def animate_button_click(self, button: QPushButton, callback=None):
        if callback:
            callback()

    def pulse_button(self, button: QPushButton):
        self.animate_button_click(button, None)

    # ── hero image ────────────────────────────────────────────────────────────
    def refresh_home_hero_image(self):
        if not hasattr(self, "home_hero_image"):
            return
        pm = _load_home_hero_pixmap()
        if pm.isNull():
            self.home_hero_image.clear()
            self.home_hero_image.setText("FuncAtlas")
            return
        sz = self.home_hero_image.size()
        w  = sz.width()  if sz.width()  > 10 else 700
        h  = sz.height() if sz.height() > 10 else 360
        scaled = pm.scaled(w, h, Qt.KeepAspectRatioByExpanding, Qt.SmoothTransformation)
        self.home_hero_image.setPixmap(scaled)
        self.home_hero_image.setAlignment(Qt.AlignCenter)
        self.home_hero_image.setScaledContents(False)

    def animate_card_entrance(self, cards: list):
        for card in cards:
            card.show()
            card.update()

    def lighten_color(self, color: QColor, factor: int = 112) -> str:
        return color.lighter(factor).name()

    # ── styles ────────────────────────────────────────────────────────────────
    def apply_styles(self):
        t           = self.theme
        accent      = self.accent_color.name()
        accent_hover = QColor(accent).lighter(114).name()
        accent_dark  = QColor(accent).darker(108).name()
        if self.current_theme == "dark":
            glossy_primary = (f"qlineargradient(x1:0,y1:0,x2:0,y2:1,"
                              f" stop:0 {QColor(accent).lighter(114).name()},"
                              f" stop:0.48 {accent}, stop:1 {accent_dark})")
            glossy_surface = (f"qlineargradient(x1:0,y1:0,x2:0,y2:1,"
                              f" stop:0 {t['bg_card']}, stop:1 {t['bg_soft']})")
        else:
            glossy_primary = (f"qlineargradient(x1:0,y1:0,x2:0,y2:1,"
                              f" stop:0 {QColor(accent).lighter(128).name()},"
                              f" stop:0.55 {accent}, stop:1 {accent_dark})")
            glossy_surface = "qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #FFFFFF, stop:1 #EEF4FB)"

        self.setStyleSheet(f"""
            QWidget {{
                background: {t['bg_main']};
                color: {t['text_primary']};
                font-family: "{self.base_font_family}";
                font-size: {self.base_font_size + 1}px;
            }}
            QLabel {{ background: transparent; color: {t['text_primary']}; }}
            QFrame#sidebar {{
                background: {t['bg_sidebar']};
                border-right: 1px solid {t['border']};
            }}
            QFrame#topHeader {{
                background: transparent;
                border: none;
                border-radius: 0px;
            }}
            QFrame#pageCard {{
                background: {t['bg_card']};
                border: 1px solid {t['border']};
                border-radius: 16px;
            }}
            QFrame#heroCard {{
                background: {t['bg_card']};
                border: 1px solid {t['border_strong']};
                border-radius: 20px;
            }}
            QFrame#softPanel {{
                background: {t['bg_soft']};
                border: 1px solid {t['border']};
                border-radius: 12px;
            }}
            QFrame#premiumCard {{
                background: {t['bg_card']};
                border: 1px solid {t['border']};
                border-radius: 18px;
            }}
            QFrame#quickActionsFrame {{
                background: {t['bg_card']};
                border: 1px solid {t['border_strong']};
                border-radius: 20px;
            }}
            QLabel#complexityInfoLabel {{
                color: {t['text_muted']}; font-size: {self.base_font_size}px;
            }}
            QLineEdit#complexityReportDisplay {{
                background: {t['bg_input']}; border: 1px solid {t['border']};
                border-radius: 10px; padding: 8px 12px; color: {t['text_primary']};
            }}
            QLabel#brandTitle {{ color: {t['text_primary']}; font-size: {self.base_font_size+11}px; font-weight: 900; }}
            QLabel#brandSubtitle {{ color: {t['text_secondary']}; font-size: {self.base_font_size-1}px; }}
            QLabel#headerTitle {{ color: {t['text_primary']}; font-size: {self.base_font_size+4}px; font-weight: 900; }}
            QLabel#headerSubtitle {{ color: {t['text_secondary']}; font-size: {self.base_font_size-1}px; }}
            QLabel#sectionTitle {{ color: {t['text_primary']}; font-size: {self.base_font_size+2}px; font-weight: 900; }}
            QLabel#sectionSubtitle {{ color: {t['text_muted']}; font-size: {self.base_font_size-1}px; }}
            QLabel#fieldLabel {{ color: {t['text_secondary']}; font-weight: 800; font-size: {self.base_font_size}px; }}
            QLabel#panelTitle {{ color: {t['text_primary']}; font-weight: 900; font-size: {self.base_font_size+1}px; }}
            QLabel#panelSubtitle {{ color: {t['text_muted']}; font-size: {self.base_font_size-1}px; }}
            QLabel#cardTitle {{ color: {t['text_primary']}; font-weight: 900; font-size: {self.base_font_size+1}px; }}
            QLabel#cardSubtitle {{ color: {t['text_secondary']}; font-size: {self.base_font_size-1}px; }}
            QLabel#heroTitle {{ color: {t['text_primary']}; font-size: {self.base_font_size+18}px; font-weight: 900; }}
            QLabel#heroKicker {{ color: {accent}; font-size: {self.base_font_size-1}px; font-weight: 900; letter-spacing: 2px; }}
            QLabel#heroSubtitle {{ color: {t['text_secondary']}; font-size: {self.base_font_size+1}px; }}
            QLabel#heroBadge {{
                background: {accent}; color: white;
                border-radius: 14px; font-size: {self.base_font_size+1}px; font-weight: 900;
            }}
            QLabel#statChipValue {{ color: {accent}; font-size: {self.base_font_size+3}px; font-weight: 900; }}
            QLabel#statChipLabel {{ color: {t['text_muted']}; font-size: {self.base_font_size-2}px; font-weight: 700; }}
            QFrame#statChip {{
                background: {t['bg_soft']}; border: 1px solid {t['border']}; border-radius: 14px;
            }}
            QPushButton {{
                background: {glossy_surface}; color: {t['text_primary']};
                border: 1px solid {t['border']};
                border-radius: 16px; padding: 5px 12px; text-align: left; font-weight: 800;
            }}
            QPushButton:hover {{ border: 1px solid {accent}; }}
            QPushButton:pressed {{ padding-top: 7px; padding-bottom: 3px; }}
            QPushButton:checked {{ background: {glossy_primary}; color: white; border: 1px solid {accent_dark}; }}
            QPushButton#smallPrimaryButton {{
                background: {glossy_primary}; color: white; border: 1px solid {accent_dark};
                border-radius: 14px; min-height: 32px; min-width: 70px;
                padding: 6px 14px; font-size: {self.base_font_size}px; font-weight: 900; text-align: center;
            }}
            QPushButton#smallPrimaryButton:hover {{ border: 1px solid {accent_hover}; }}
            QPushButton#ghostButton {{
                background: {glossy_surface}; color: {t['text_primary']}; border: 1px solid {t['border_strong']};
                border-radius: 14px; min-height: 32px; min-width: 54px;
                padding: 6px 12px; font-weight: 800; font-size: {self.base_font_size}px;
            }}
            QPushButton#ghostButton:hover {{ border: 1px solid {accent}; }}
            QPushButton#pickerButton {{
                background: {glossy_primary}; color: white; border: 1px solid {accent_dark};
                border-radius: 13px; min-height: 34px; min-width: 108px;
                padding: 6px 14px; font-weight: 900; font-size: {self.base_font_size}px; text-align: center;
            }}
            QPushButton#pickerButton:hover {{ border: 1px solid {accent_hover}; background: {glossy_primary}; }}
            QPushButton#pickerButton:disabled {{
                background: {t['bg_soft']}; color: {t['text_muted']};
                border: 1px solid {t['border']};
            }}
            QPushButton#clearButton {{
                background: {glossy_primary}; color: white; border: 1px solid {accent_dark};
                border-radius: 13px; min-height: 34px; min-width: 80px;
                padding: 6px 14px; font-weight: 900; font-size: {self.base_font_size}px; text-align: center;
            }}
            QPushButton#clearButton:hover {{ border: 1px solid {accent_hover}; background: {glossy_primary}; }}
            QPushButton#clearButton:pressed {{ background: {accent_dark}; }}
            QPushButton#clearButtonRect {{
                background: {glossy_primary}; color: white; border: 1px solid {accent_dark};
                border-radius: 13px; min-height: 34px; min-width: 80px;
                padding: 6px 14px; font-weight: 900; font-size: {self.base_font_size}px; text-align: center;
            }}
            QPushButton#clearButtonRect:hover {{ border: 1px solid {accent_hover}; background: {glossy_primary}; }}
            QPushButton#clearButtonRect:pressed {{ background: {accent_dark}; }}
            QPushButton#toggleLeft {{
                background: {t['bg_soft']}; color: {t['text_muted']};
                border: 1px solid {accent_dark};
                border-top-left-radius: 10px; border-bottom-left-radius: 10px;
                border-top-right-radius: 0px; border-bottom-right-radius: 0px;
                min-width: 160px; padding: 4px 14px; font-weight: 800;
                font-size: {self.base_font_size}px;
            }}
            QPushButton#toggleLeft:checked {{
                background: {t['accent']}; color: white;border: 2px solid {accent_hover};
            }}
            QPushButton#toggleRight {{
                background: {t['bg_soft']}; color: {t['text_muted']};
                border: 1px solid {accent_dark};
                border-top-right-radius: 10px; border-bottom-right-radius: 10px;
                border-top-left-radius: 0px; border-bottom-left-radius: 0px;
                min-width: 140px; padding: 4px 14px; font-weight: 800;
                font-size: {self.base_font_size}px;
            }}
            QPushButton#toggleRight:checked {{
                background: {t['accent']}; color: white; border: 2px solid {accent_hover};
            }}
            QWidget#funcColField QLineEdit:disabled {{
                background: {t['bg_soft']}; color: {t['text_muted']};
                border: 1px solid {t['border']};
            }}
            QWidget#funcColField QPushButton:disabled {{
                background: {t['bg_soft']}; color: {t['text_muted']};
                border: 1px solid {t['border']};
            }}
            /* ── 2-colour scheme: Blue = Function inputs, Green = Consolidated DB inputs ── */
            /* Function List Files box */
            QWidget#funcListField QTextEdit {{
                border: 2px solid #2196F3; border-radius: 10px;
            }}
            QWidget#funcListField QLabel#fieldLabel {{ color: #1565C0; font-weight: 900; }}
            /* Function Name Column box */
            QWidget#funcColField QLineEdit {{
                border: 2px solid #2196F3; border-radius: 10px;
            }}
            QWidget#funcColField QLabel#fieldLabel {{ color: #1565C0; font-weight: 900; }}
            /* Consolidated DB Excel File box */
            QWidget#dbExcelField QTextEdit, QWidget#dbExcelField QLineEdit {{
                border: 2px solid #43A047; border-radius: 10px;
            }}
            QWidget#dbExcelField QLabel#fieldLabel {{ color: #2E7D32; font-weight: 900; }}
            /* Base File Name Column box */
            QWidget#baseColField QLineEdit {{
                border: 2px solid #43A047; border-radius: 10px;
            }}
            QWidget#baseColField QLabel#fieldLabel {{ color: #2E7D32; font-weight: 900; }}
            QPushButton#collapseToggle {{
                background: {t['bg_soft']}; color: {t['text_primary']}; border: 1px solid {t['border']};
                border-radius: 12px; padding: 8px 14px; text-align: left; font-weight: 800;
            }}
            QLineEdit, QTextEdit {{
                background: {t['bg_input']}; color: {t['text_primary']};
                border: 1px solid {t['border']}; border-radius: 10px; padding: 6px 10px;
            }}
            QLineEdit:focus, QTextEdit:focus {{ border: 1px solid {accent}; }}
            QComboBox {{
                background: {t['bg_input']}; color: {t['text_primary']};
                border: 1px solid {t['border']}; border-radius: 10px; padding: 5px 10px; font-weight: 700;
            }}
            QComboBox:focus {{ border: 1px solid {accent}; }}
            QTreeWidget {{
                background: {t['bg_input']}; color: {t['text_primary']};
                border: 1px solid {t['border']}; border-radius: 10px;
            }}
            QTreeWidget::item:selected {{ background: {accent}55; color: {t['text_primary']}; }}
            QScrollBar:vertical {{ background: transparent; width: 10px; margin: 2px; }}
            QScrollBar::handle:vertical {{ background: {t['border_strong']}; min-height: 28px; border-radius: 5px; }}
            QScrollBar::handle:vertical:hover {{ background: {accent}; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; background: transparent; border: none; }}
            QProgressBar {{
                background: {t['bg_input']}; border: 1px solid {t['border']}; border-radius: 4px;
                text-align: center; height: 8px; font-weight: 800; color: {t['text_primary']};
            }}
            QProgressBar::chunk {{ background: {glossy_primary}; border-radius: 3px; margin: 1px; }}
            QProgressBar#reportOverallProgress {{
                background: {t['bg_input']}; border: 1px solid {t['border']}; border-radius: 4px; height: 8px;
            }}
            QProgressBar#reportOverallProgress::chunk {{ background: {glossy_primary}; border-radius: 3px; margin: 1px; }}
            QProgressBar#reportStepProgress {{
                background: {t['bg_input']}; border: none; border-radius: 2px; height: 5px;
            }}
            QProgressBar#reportStepProgress::chunk {{ background: {accent}88; border-radius: 2px; }}
            QSplitter::handle {{ background: {t['bg_soft']}; border-radius: 4px; }}
        """)

        if hasattr(self, "header_status"):
            self.header_status.setStyleSheet(
                f"background: transparent; color: {accent}; border: 1px solid {t['border_strong']};"
                f" border-radius: 12px; padding: 6px 12px; font-weight: 900;"
            )
        if hasattr(self, "theme_btn"):
            self.theme_btn.setText("☀ Light" if self.current_theme == "dark" else "🌙 Dark")
        if hasattr(self, "settings_theme_combo"):
            self.settings_theme_combo.blockSignals(True)
            self.settings_theme_combo.setCurrentIndex(0 if self.current_theme == "dark" else 1)
            self.settings_theme_combo.blockSignals(False)
        if hasattr(self, "theme_note_label"):
            self.theme_note_label.setText(f"Current theme: {t['name']} • Accent: {accent}")
        if hasattr(self, "report_pct_label"):
            self.report_pct_label.setStyleSheet(
                f"color: {accent}; font-weight: 900; font-size: {self.base_font_size+1}px; background: transparent;"
            )
        if hasattr(self, "report_phase_label"):
            self.report_phase_label.setStyleSheet(
                f"color: {accent}; font-weight: 900; font-size: {self.base_font_size+1}px; background: transparent;"
            )
        for card in self.accent_cards:
            try:
                card._shadow = add_shadow(card, blur=28, y_offset=9, alpha=80)
            except Exception:
                pass

    def rebuild_icons(self):
        icon_color = QColor("#FFFFFF") if self.current_theme == "dark" else QColor("#102134")
        self.icons = VectorIconFactory(icon_color)
        self.icons_white = self.icons
        icon_map = {
            "home": "home", "input": "input", "view": "view",
            "diff": "diff", "report": "report", "complexity": "settings",
            "help": "help", "settings": "settings",
        }
        for key, btn in self.nav_buttons.items():
            btn.setIcon(self.icons.icon(icon_map[key], 18))

        for field_attr, icon_name in [
            ("ref_target_field",      "folder"),
            ("ref_bases_field",       "folder"),
            ("ref_function_field",    "document"),
            ("con_function_field",    "document"),
            ("con_folder_field",      "folder"),
            ("con_db_excel_field",    "excel"),
            ("con_func_col_field",    "column"),
            ("con_base_col_field",    "column"),
            ("con_output_link_field", "link"),
        ]:
            if hasattr(self, field_attr) and hasattr(getattr(self, field_attr), "button"):
                getattr(self, field_attr).button.setIcon(self.icons.icon(icon_name, 15))

        if hasattr(self, "_report_browse_btn"):
            self._report_browse_btn.setIcon(self.icons.icon("folder", 15))

        for attr, icon_name in [
            ("settings_accent_btn", "palette"),
            ("settings_font_btn",   "font"),
            ("settings_reset_btn",  "reset"),
            ("con_open_output_btn", "link"),
        ]:
            if hasattr(self, attr):
                getattr(self, attr).setIcon(self.icons.icon(icon_name, 15 if attr != "theme_btn" else 18))

        for attr, icon_name in [
            ("ref_submit_btn", "submit"), ("ref_back_btn", "back"),   ("ref_clear_btn", "clear"),
            ("con_submit_btn", "submit"), ("con_back_btn", "back"),   ("con_clear_btn", "clear"),
            ("report_generate_btn", "submit"), ("report_clear_btn", "clear"),
            ("report_open_btn", "link"),
        ]:
            if hasattr(self, attr):
                getattr(self, attr).setIcon(self.icons.icon(icon_name, 15))

    # ── theme actions ─────────────────────────────────────────────────────────
    def set_theme_mode(self, mode: str):
        mode = (mode or "dark").lower()
        if mode not in ThemeManager.THEMES:
            mode = "dark"
        self.current_theme = mode
        self.theme = ThemeManager.THEMES[mode]
        self.apply_styles()
        self.rebuild_icons()

    def toggle_theme(self):
        self.set_theme_mode("light" if self.current_theme == "dark" else "dark")

    def on_theme_combo_changed(self, idx: int):
        self.set_theme_mode("dark" if idx == 0 else "light")

    def pick_accent_color(self):
        color = QColorDialog.getColor(self.accent_color, self, "Choose Accent Color")
        if color.isValid():
            self.accent_color = color
            self.apply_styles()
            self.rebuild_icons()

    def pick_font(self):
        font, ok = QFontDialog.getFont(QFont(self.base_font_family, self.base_font_size), self, "Choose Font")
        if ok:
            self.base_font_family = font.family()
            pt = font.pointSize() if font.pointSize() > 0 else 10
            self.base_font_size = max(9, min(pt, 14))
            self.apply_styles()
            self.rebuild_icons()

    def reset_theme(self):
        self.accent_color     = QColor(self.default_accent)
        self.base_font_family = "Segoe UI"
        self.base_font_size   = 10
        self.set_theme_mode("dark")

    def reset_all(self):
        """Hard reset: clear all cache, progress, and form state across every page."""
        from PySide6.QtWidgets import QMessageBox
        reply = QMessageBox.question(
            self, "Reset Everything",
            "This will clear all loaded data, cache, and form inputs.\nAre you sure?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        # ── Clear scan cache ──────────────────────────────────────────────────
        with SCAN_CACHE_LOCK:
            SCAN_CACHE.clear()

        # ── Reset all core state ──────────────────────────────────────────────
        self.function_records            = OrderedDict()
        self.current_function_name       = ""
        self.current_function_file       = ""
        self.current_function_body       = ""
        self.current_function_list_paths = []
        self.available_sources           = []
        self.active_source_root          = ""
        self.current_reference_folders   = []
        self._ref_function_filter        = []

        # ── Clear reference form ──────────────────────────────────────────────
        if hasattr(self, "ref_target_field"):
            self.ref_target_field.clear_selection()
        if hasattr(self, "ref_bases_field"):
            self.ref_bases_field.clear_selection()
        if hasattr(self, "ref_function_field"):
            self.ref_function_field.clear_selection()

        # ── Clear consolidated form ───────────────────────────────────────────
        if hasattr(self, "con_function_field"):
            self.con_function_field.clear_selection()
        if hasattr(self, "con_db_excel_field"):
            self.con_db_excel_field.clear_selection()
        if hasattr(self, "con_func_col_field"):
            self.con_func_col_field.clear_selection()
        if hasattr(self, "con_base_col_field"):
            self.con_base_col_field.clear_selection()
        if hasattr(self, "con_output_link_field"):
            self.con_output_link_field.clear_selection()
        if hasattr(self, "con_folder_field"):
            self.con_folder_field.clear_selection()

        # ── Reset view page ───────────────────────────────────────────────────
        if hasattr(self, "source_combo"):
            self.source_combo.blockSignals(True)
            self.source_combo.clear()
            self.source_combo.blockSignals(False)
        if hasattr(self, "tree"):
            self.tree.clear()
        if hasattr(self, "view_text"):
            self.view_text.setText("")
        if hasattr(self, "view_title"):
            self.view_title.setText("")
        if hasattr(self, "view_meta"):
            self.view_meta.setText("")
        if hasattr(self, "view_mode_chip"):
            self.view_mode_chip.setVisible(False)

        # ── Reset diff page (including raw data so filters stay empty) ────────
        self._diff_raw_data = []
        self._clear_diff()

        # ── Reset report fields ───────────────────────────────────────────────
        if hasattr(self, "report_output_field"):
            self.report_output_field.clear_selection()
        if hasattr(self, "report_target_field"):
            self.report_target_field.clear_selection()

        QMessageBox.information(self, "Reset Complete", "All data and cache have been cleared.")

    # ── form clear ────────────────────────────────────────────────────────────
    def clear_reference_form(self):
        self.ref_target_field.clear_selection()
        self.ref_bases_field.clear_selection()
        self.ref_function_field.clear_selection()
        self._ref_function_filter = []
        # Point 6: auto-clear diff when inputs are cleared
        self._clear_diff()
        self.available_sources = []
        self.current_function_list_paths = []
        self.current_reference_folders = []
        self.function_records = OrderedDict()
        if hasattr(self, "source_combo"):
            self.source_combo.blockSignals(True)
            self.source_combo.clear()
            self.source_combo.blockSignals(False)
        if hasattr(self, "view_mode_chip"):
            self.view_mode_chip.setVisible(False)

    def clear_consolidated_form(self):
        self.con_function_field.clear_selection()
        self.con_db_excel_field.clear_selection()
        self.con_func_col_field.clear_selection()
        self.con_base_col_field.clear_selection()
        self.con_output_link_field.clear_selection()
        if hasattr(self, "con_folder_field"):
            self.con_folder_field.clear_selection()
        # Reset toggle back to Excel mode
        if hasattr(self, "con_toggle_excel_btn"):
            self.con_toggle_excel_btn.setChecked(True)
            self.con_source_stack.setCurrentIndex(0)
            self.con_func_col_field.setEnabled(True)

    # ── UI scaffold ───────────────────────────────────────────────────────────
    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QHBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # Sidebar
        self.sidebar = QFrame()
        self.sidebar.setObjectName("sidebar")
        self.sidebar.setFixedWidth(230)
        sidebar_layout = QVBoxLayout(self.sidebar)
        sidebar_layout.setContentsMargins(12, 16, 12, 16)
        sidebar_layout.setSpacing(8)

        brand_wrap   = QWidget()
        brand_wrap.setStyleSheet("background: transparent;")
        brand_layout = QVBoxLayout(brand_wrap)
        brand_layout.setContentsMargins(0, 0, 0, 0)
        brand_layout.setSpacing(2)
        brand_title    = QLabel("FuncAtlas")
        brand_title.setObjectName("brandTitle")
        brand_subtitle = QLabel("Enterprise Desktop Suite")
        brand_subtitle.setObjectName("brandSubtitle")
        brand_layout.addWidget(brand_title)
        brand_layout.addWidget(brand_subtitle)
        sidebar_layout.addWidget(brand_wrap)

        nav_items = [
            ("home",       "Home",                    "home"),
            ("input",      "Input",                   "input"),
            ("view",       "View",                    "view"),
            ("diff",       "Diff",                    "diff"),
            ("report",     "Report",                  "report"),
            ("complexity", "Complexity & Compatibility", "settings"),
            ("help",       "Help",                    "help"),
            ("settings",   "Settings",                "settings"),
        ]
        for key, text, icon_name in nav_items:
            btn = NavButton(text, self.icons.icon(icon_name, 18))
            btn.clicked.connect(lambda checked=False, k=key: self.show_page(k))
            sidebar_layout.addWidget(btn, alignment=Qt.AlignLeft)
            self.nav_buttons[key] = btn
        sidebar_layout.addStretch()
        root.addWidget(self.sidebar)

        # Right content
        right_wrap   = QWidget()
        right_layout = QVBoxLayout(right_wrap)
        right_layout.setContentsMargins(12, 12, 12, 12)
        right_layout.setSpacing(10)

        self.header = QFrame()
        self.header.setObjectName("topHeader")
        self.header.setFixedHeight(82)
        header_layout = QHBoxLayout(self.header)
        header_layout.setContentsMargins(18, 12, 18, 12)

        header_left = QWidget()
        hl_layout   = QVBoxLayout(header_left)
        hl_layout.setContentsMargins(0, 0, 0, 0)
        hl_layout.setSpacing(1)
        self.header_title    = QLabel("Dashboard")
        self.header_title.setObjectName("headerTitle")
        self.header_subtitle = QLabel("Professional reusable function analysis workspace")
        self.header_subtitle.setObjectName("headerSubtitle")
        hl_layout.addWidget(self.header_title)
        hl_layout.addWidget(self.header_subtitle)
        header_layout.addWidget(header_left)
        header_layout.addStretch()

        self.theme_btn = IconTextButton("🌙 Dark", QIcon())
        self.theme_btn.setObjectName("pickerButton")
        self.theme_btn.setMinimumWidth(118)
        self.theme_btn.setFixedHeight(38)
        self.theme_btn.clicked.connect(self.toggle_theme)

        self.reset_all_btn = IconTextButton("🔄 Reset", QIcon())
        self.reset_all_btn.setObjectName("pickerButton")
        self.reset_all_btn.setMinimumWidth(118)
        self.reset_all_btn.setFixedHeight(38)
        self.reset_all_btn.clicked.connect(self.reset_all)

        header_layout.addWidget(self.theme_btn)
        header_layout.addSpacing(8)
        header_layout.addWidget(self.reset_all_btn)
        right_layout.addWidget(self.header)

        self.stack = QStackedWidget()
        right_layout.addWidget(self.stack)
        root.addWidget(right_wrap)

        # Build all pages
        create_home_page(self)
        create_input_page(self)
        create_reference_page(self)
        create_consolidated_page(self)
        create_view_page(self)
        create_diff_page(self)
        create_report_page(self)
        create_complexity_page(self)
        create_help_page(self)
        create_settings_page(self)

        # ── Global loading overlay (covers entire window) ─────────────────────
        self._setup_global_overlay()

    def _setup_global_overlay(self):
        """Create a full-window dim + spinner overlay parented to centralWidget."""
        central = self.centralWidget()

        self._global_overlay = QWidget(central)
        self._global_overlay.setObjectName("globalOverlay")
        self._global_overlay.setAttribute(Qt.WA_TransparentForMouseEvents, False)
        self._global_overlay.setVisible(False)

        # Semi-transparent black dim
        self._global_overlay_dim = QLabel(self._global_overlay)
        self._global_overlay_dim.setStyleSheet(
            "background: rgba(0,0,0,150); border-radius: 0px;"
        )

        # Centered spinner card
        spinner_card = QFrame(self._global_overlay)
        spinner_card.setObjectName("globalSpinnerCard")
        spinner_card.setFixedSize(160, 160)
        spinner_card.setStyleSheet(
            "QFrame#globalSpinnerCard {"
            "  background: rgba(255,255,255,235);"
            "  border-radius: 18px;"
            "}"
        )
        sc_layout = QVBoxLayout(spinner_card)
        sc_layout.setContentsMargins(16, 16, 16, 16)
        sc_layout.setSpacing(10)
        sc_layout.setAlignment(Qt.AlignCenter)

        self._global_spinner_lbl = QLabel("⠋")
        self._global_spinner_lbl.setAlignment(Qt.AlignCenter)
        self._global_spinner_lbl.setStyleSheet(
            "font-size: 42px; color: #1565C0; background: transparent;"
        )
        self._global_loading_txt = QLabel("Please wait…")
        self._global_loading_txt.setAlignment(Qt.AlignCenter)
        self._global_loading_txt.setStyleSheet(
            "font-size: 13px; font-weight: 700; color: #333; background: transparent;"
        )
        sc_layout.addWidget(self._global_spinner_lbl)
        sc_layout.addWidget(self._global_loading_txt)
        self._global_spinner_card = spinner_card

        _FRAMES = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
        self._global_spinner_idx = 0
        self._global_spinner_timer = QTimer(self)
        self._global_spinner_timer.setInterval(80)

        def _tick():
            self._global_spinner_idx = (self._global_spinner_idx + 1) % len(_FRAMES)
            self._global_spinner_lbl.setText(_FRAMES[self._global_spinner_idx])

        self._global_spinner_timer.timeout.connect(_tick)

        def _resize_global_overlay():
            w, h = central.width(), central.height()
            self._global_overlay.setGeometry(0, 0, w, h)
            self._global_overlay_dim.setGeometry(0, 0, w, h)
            cx = (w - spinner_card.width()) // 2
            cy = (h - spinner_card.height()) // 2
            spinner_card.move(cx, cy)

        self._global_overlay.resizeEvent = lambda e: _resize_global_overlay()
        _orig_resize = central.resizeEvent if hasattr(central, "resizeEvent") else None

        def _central_resize(e):
            _resize_global_overlay()
            if _orig_resize:
                _orig_resize(e)
            else:
                e.accept()

        central.resizeEvent = _central_resize

    def show_loading(self, message="Please wait…"):
        """Show the global full-window loading overlay."""
        central = self.centralWidget()
        w, h = central.width(), central.height()
        self._global_overlay.setGeometry(0, 0, w, h)
        self._global_overlay_dim.setGeometry(0, 0, w, h)
        cx = (w - self._global_spinner_card.width()) // 2
        cy = (h - self._global_spinner_card.height()) // 2
        self._global_spinner_card.move(cx, cy)
        self._global_loading_txt.setText(message)
        self._global_overlay.raise_()
        self._global_overlay.setVisible(True)
        self._global_spinner_timer.start()
        from PySide6.QtWidgets import QApplication
        QApplication.processEvents()

    def hide_loading(self):
        """Hide the global loading overlay."""
        self._global_spinner_timer.stop()
        self._global_overlay.setVisible(False)

    # ── page navigation ───────────────────────────────────────────────────────
    def make_scroll_page(self, content_widget: QWidget) -> QScrollArea:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setContentsMargins(0, 0, 0, 0)
        scroll.setWidget(content_widget)
        return scroll

    def set_header(self, title: str, subtitle: str):
        self.header_title.setText(title)
        self.header_subtitle.setText(subtitle)

    def show_page(self, page_name: str):
        headers = {
            "home":         ("Dashboard",            "Professional overview of the FuncAtlas workspace"),
            "input":        ("Input Workspace",       "Choose the correct data input flow"),
            "reference":    ("Reference Bases",       "Upload target base folder, reference folders, and function list file"),
            "consolidated": ("Consolidated Database", "Upload function list files, consolidated Excel, detect columns, and auto-generate output"),
            "view":         ("Function Explorer",     "Browse selected source files and click functions to preview body"),
            "diff":         ("Diff Workspace",        "Compare logic and reusable code structures"),
            "report":       ("Report Generator",           "Extracting Functions, Comparing and Generating Reports"),
            "complexity":   ("Complexity & compatibility",  "Review the generated report and manage complexity settings"),
            "help":         ("Help & Guidance",             "Instructions, workflow rules, and user support"),
            "settings":     ("Settings",              "Appearance, font, and theme controls"),
        }
        active_nav = {
            "home": "home", "input": "input", "reference": "input", "consolidated": "input",
            "view": "view", "diff": "diff", "report": "report",
            "complexity": "complexity",
            "help": "help", "settings": "settings",
        }
        for btn in self.nav_buttons.values():
            btn.setChecked(False)
        self.nav_buttons[active_nav[page_name]].setChecked(True)

        widget = self.pages[page_name]
        self.stack.setCurrentWidget(widget)
        self.stack.updateGeometry()
        widget.updateGeometry()
        widget.adjustSize()
        widget.update()
        QApplication.processEvents()

        effect = QGraphicsOpacityEffect(widget)
        widget.setGraphicsEffect(effect)
        fade = QPropertyAnimation(effect, b"opacity", self)
        fade.setDuration(190)
        fade.setStartValue(0.0)
        fade.setEndValue(1.0)
        fade.setEasingCurve(QEasingCurve.OutCubic)
        fade.finished.connect(lambda: widget.setGraphicsEffect(None))
        fade.start()
        self._page_fade = fade

        title, subtitle = headers[page_name]
        self.set_header(title, subtitle)
        if hasattr(self, "header_status"):
            self.header_status.setText(title)

        if page_name == "home" and hasattr(self, "home_cards"):
            QTimer.singleShot(220, lambda: self.animate_card_entrance(self.home_cards))
        elif page_name == "input" and hasattr(self, "input_cards"):
            QTimer.singleShot(220, lambda: self.animate_card_entrance(self.input_cards))

    # ── source helpers ────────────────────────────────────────────────────────
    def get_source_display_name(self, path: str) -> str:
        normalized = normalize_path(path)
        parts = [p for p in normalized.replace("\\", "/").split("/") if p]
        if not parts:
            return normalized
        if parts[-1].lower() == "src" and len(parts) >= 2:
            return parts[-2]
        return parts[-1]

    def build_source_entry(self, source_type: str, path: str) -> dict:
        base_name = self.get_source_display_name(path)
        prefix = {"target": "Target Base", "reference": "Reference Base",
                  "consolidated": "Consolidated DB"}.get(source_type, "Source")
        return {"type": source_type, "path": normalize_path(path), "label": f"{prefix} - {base_name}"}

    def register_sources(self, entries: list, function_list_paths: list, reference_folders: list = None):
        self.available_sources           = entries
        self.current_function_list_paths = function_list_paths or []
        self.current_reference_folders   = reference_folders or []
        self.source_combo.blockSignals(True)
        self.source_combo.clear()
        for entry in self.available_sources:
            self.source_combo.addItem(entry["label"], entry["path"])
        self.source_combo.blockSignals(False)
        if self.available_sources:
            self.source_combo.setCurrentIndex(0)
            self.load_selected_source()
            self.show_page("view")

    def load_selected_source(self):
        if not self.available_sources:
            return
        idx = self.source_combo.currentIndex()
        if idx < 0 or idx >= len(self.available_sources):
            return
        source = self.available_sources[idx]
        self.active_source_root = source["path"]

        # Points 1,2,3: filtering only applies to TARGET base; refs always show all.
        is_target = (source.get("type") == "target")

        try:
            all_records     = scan_source_for_all_functions(self.active_source_root)
            matched_by_list = OrderedDict()
            matched_by_ref  = OrderedDict()
            function_filter = getattr(self, "_ref_function_filter", [])
            has_xlsx        = bool(self.current_function_list_paths)  # point 4

            if is_target:
                if has_xlsx:
                    # Point 2: xlsx uploaded -> only show functions that match the list
                    matched_by_list = match_target_with_function_list(
                        self.active_source_root, self.current_function_list_paths)
                    self.function_records = matched_by_list if matched_by_list else all_records
                else:
                    # Point 1: no xlsx -> show ALL target functions regardless of
                    # whether they exist in reference folders (reference only affects diff)
                    self.function_records = all_records
            else:
                # Point 3: reference bases -> always all functions, no filtering
                self.function_records = all_records

            # Point 4: mode chip visible ONLY when xlsx function list is loaded for target
            if hasattr(self, "view_mode_chip"):
                if has_xlsx and is_target:
                    fn_count = sum(len(v["functions"]) for v in self.function_records.values())
                    self.view_mode_chip.setText(
                        "  Function List Filter  ·  {} functions  ".format(fn_count)
                    )
                    self.view_mode_chip.setVisible(True)
                else:
                    self.view_mode_chip.setVisible(False)

            self.populate_view_tree()
            self.populate_diff_tree()
            self.view_title.setText(source["label"])
            self.view_meta.setText(self.active_source_root)
            self.view_text.setText(
                "Selected source:\n{}\n{}\n\nLoaded Files: {}\n\n"
                "Drag the center divider left or right to resize the panes.\n"
                "Click a function on the left to preview the real body.".format(
                    source["label"], self.active_source_root, len(self.function_records))
            )
            self.current_function_name = ""
            self.current_function_file = ""
            self.current_function_body = ""
        except Exception as e:
            QMessageBox.critical(self, "Load Error", str(e))

    def on_source_changed(self, _index):
        self.show_loading("Loading source…")
        try:
            self.load_selected_source()
        finally:
            self.hide_loading()

    # ── tree population ───────────────────────────────────────────────────────
    def populate_view_tree(self):
        self.tree.clear()
        if not self.function_records:
            self.view_text.setText("No matching files/functions found for the selected source.")
            return
        for file_path, info in self.function_records.items():
            parent_text = f"{info['display_name']} ({len(info['functions'])})"
            parent = QTreeWidgetItem([parent_text])
            parent.setData(0, Qt.UserRole, {"type": "file", "file_path": file_path})
            for func_name in info["functions"]:
                child = QTreeWidgetItem([func_name])
                child.setData(0, Qt.UserRole, {"type": "function", "file_path": file_path, "function_name": func_name})
                parent.addChild(child)
            self.tree.addTopLevelItem(parent)
        self.tree.expandAll()

    def populate_diff_tree(self):
        """Build diff tree:
        - Fix 1: always use the TARGET source on the left; ignore which source
          is currently selected in the View dropdown.
        - Fix 2: support multiple reference folders (one tab per ref)
        - Tag each function added/deleted/modified for filter buttons
        """
        import difflib as _difflib
        self.diff_tree.clear()
        self._diff_raw_data = []   # reset before repopulating
        ref_folders = getattr(self, "current_reference_folders", [])

        # --- Fix 1: resolve target records independently of the current view selection ---
        target_entry = next(
            (s for s in getattr(self, "available_sources", []) if s.get("type") == "target"),
            None,
        )
        if target_entry:
            try:
                target_records = scan_source_for_all_functions(target_entry["path"])
            except Exception:
                target_records = {}
        else:
            # Fallback: if no explicit target, use whatever is in function_records
            target_records = self.function_records

        # Build reference index per FOLDER:
        # {display_name -> {func_name -> [(folder_path, file_path), ...]}}
        ref_index: dict = {}
        for folder in ref_folders:
            for fp, rinfo in scan_source_for_all_functions(folder).items():
                dn = rinfo["display_name"]
                ref_index.setdefault(dn, {})
                for fn in rinfo["functions"]:
                    ref_index[dn].setdefault(fn, []).append((folder, fp))

        for file_path, info in target_records.items():
            dn = info["display_name"]
            ref_fns = ref_index.get(dn, {})  # {func_name -> [(folder, fp)]}

            # If this file does not exist in ANY reference folder,
            # skip it entirely — it should not appear in the diff tree.
            if not ref_fns and ref_folders:
                continue

            tagged_children = []
            for func_name in info["functions"]:
                if func_name in ref_fns:
                    # Check across ALL reference copies — use worst-case tag
                    # (if any ref differs → modified; if all identical → equal)
                    overall_tag = "equal"
                    best_tgt_body = ""
                    best_ref_body = ""
                    for folder, ref_fp in ref_fns[func_name]:
                        tgt_body = extract_function_body(file_path, func_name)
                        ref_body = extract_function_body(ref_fp, func_name)
                        t_lines  = [l for l in tgt_body.splitlines() if l.strip()]
                        r_lines  = [l for l in ref_body.splitlines() if l.strip()]
                        ratio    = _difflib.SequenceMatcher(None, t_lines, r_lines).ratio()
                        if ratio < 1.0:
                            overall_tag = "modified"
                            best_tgt_body = tgt_body
                            best_ref_body = ref_body
                            break
                    tag = overall_tag

                    # Compute line-level diff flags for filter button logic.
                    # These track whether the function contains added lines
                    # (only_target), deleted lines (only_ref), or modified
                    # lines — independent of the function-level tag.
                    has_added    = False
                    has_deleted  = False
                    has_modified = False
                    if tag == "modified":
                        tgt_lines = best_tgt_body.splitlines()
                        ref_lines = best_ref_body.splitlines()
                        matcher   = _difflib.SequenceMatcher(None, tgt_lines, ref_lines)
                        for op, i1, i2, j1, j2 in matcher.get_opcodes():
                            if op == "equal":
                                continue
                            nc = tgt_lines[i1:i2]
                            rc = ref_lines[j1:j2]
                            if op == "replace":
                                if len(nc) == 1 and len(rc) == 1:
                                    sim = _difflib.SequenceMatcher(
                                        None, nc[0].strip(), rc[0].strip()
                                    ).ratio()
                                    if sim >= 0.5:
                                        has_modified = True
                                    else:
                                        if any(l.strip() for l in nc): has_added   = True
                                        if any(l.strip() for l in rc): has_deleted = True
                                else:
                                    if any(l.strip() for l in nc): has_added   = True
                                    if any(l.strip() for l in rc): has_deleted = True
                            elif op == "delete":
                                if any(l.strip() for l in nc): has_added = True
                            elif op == "insert":
                                if any(l.strip() for l in rc): has_deleted = True

                    # Always include functions that exist in both target and
                    # reference — "equal" functions appear without highlight so
                    # users can see every matching function, not just changed ones.
                    tagged_children.append((
                        func_name, tag, file_path,
                        ref_fns.get(func_name, []),  # list of (folder, fp)
                        has_added, has_deleted, has_modified,
                    ))
                else:
                    # Function only in target and NOT present in any reference
                    # file — skip it; there is nothing to diff against.
                    continue

            # Deleted: in reference but NOT in target
            target_fn_set = set(info["functions"])
            for ref_fn, ref_copies in ref_fns.items():
                if ref_fn not in target_fn_set:
                    tagged_children.append((
                        ref_fn, "deleted", file_path, ref_copies,
                        False, True, False,  # has_added=F, has_deleted=T, has_modified=F
                    ))

            if not tagged_children:
                continue

            # Store raw entry so _apply_diff_filter can rebuild the tree cleanly
            self._diff_raw_data.append({
                "display_name": dn,
                "file_path":    file_path,
                "functions":    tagged_children,
            })

        # Delegate all tree-building to _apply_diff_filter so the active
        # filter/search is always respected from a single code path.
        self._apply_diff_filter()

    def filter_tree_items(self, text: str, tree_widget=None):
        query = text.strip().lower()
        tree  = tree_widget if tree_widget is not None else self.tree
        # If this is the diff tree, delegate to _apply_diff_filter so tag
        # filter and search filter are always combined correctly (Bug 3 fix).
        if tree is getattr(self, "diff_tree", None):
            self._apply_diff_filter()
            return
        for i in range(tree.topLevelItemCount()):
            parent = tree.topLevelItem(i)
            parent_visible = False
            for j in range(parent.childCount()):
                child = parent.child(j)
                child_match = (query in child.text(0).lower()) if query else True
                child.setHidden(not child_match)
                if child_match:
                    parent_visible = True
            parent_match = (query in parent.text(0).lower()) if query else True
            parent.setHidden(not (parent_match or parent_visible))

    # ── tree click handlers ───────────────────────────────────────────────────
    def on_tree_item_clicked(self, item, column):
        payload = item.data(0, Qt.UserRole)
        if not payload:
            return
        if payload["type"] == "file":
            fp = payload["file_path"]
            self.current_function_name = ""; self.current_function_file = fp; self.current_function_body = ""
            self.view_title.setText(os.path.basename(fp))
            self.view_meta.setText(fp)
            self.view_text.setText(f"File selected:\n{fp}\n\nFunctions are listed under this file.")
        elif payload["type"] == "function":
            fp   = payload["file_path"]
            name = payload["function_name"]
            body = extract_function_body(fp, name)
            self.current_function_name = name
            self.current_function_file = fp
            self.current_function_body = body
            self.view_title.setText(name)
            self.view_meta.setText(fp)
            self.view_text.setText(body)

    def on_diff_item_clicked(self, item, column):
        """Show diff for clicked function.
        Issue 2: one tab per reference folder.
        Issue 3: colour on BOTH left and right columns correctly.
        """
        import difflib
        payload = item.data(0, Qt.UserRole)
        if not payload or payload["type"] != "function":
            return

        function_name = payload["function_name"]
        target_file   = payload["file_path"]
        ref_copies    = payload.get("ref_copies", [])   # [(folder_path, file_path), ...]
        diff_tag      = payload.get("diff_tag", "")

        # Update fullscreen info bar if active
        if getattr(self, "_diff_fullscreen", False):
            self._diff_update_fs_info()
            self._diff_update_nav_state()

        # Always clear old diff first
        self.diff_tabs.clear()

        # Fix 1: always derive target_label from the target source entry, not the
        # currently-selected source in the view dropdown
        target_entry_for_label = next(
            (s for s in getattr(self, "available_sources", []) if s.get("type") == "target"),
            None,
        )
        if target_entry_for_label:
            base_path    = os.path.normpath(target_entry_for_label["path"])
        else:
            base_path    = os.path.normpath(self.active_source_root)
        parts        = base_path.split(os.sep)
        target_label = (parts[parts.index("src") - 1]
                        if "src" in parts and parts.index("src") > 0
                        else os.path.basename(base_path))

        target_code = (extract_function_body(target_file, function_name)
                       if target_file and os.path.isfile(target_file) else "")

        def _make_ref_label(folder_path):
            norm = os.path.normpath(folder_path).split(os.sep)
            return (norm[norm.index("src") - 1]
                    if "src" in norm and norm.index("src") > 0
                    else os.path.basename(folder_path))

        def _build_table(tgt_code, ref_code, tgt_lbl, ref_lbl, fn_tag):
            """Build a QTableWidget showing the side-by-side diff with correct
            colour on BOTH columns (Issue 3):
              only_target  → green  left,  empty  right
              only_ref     → empty  left,  yellow right
              modified     → blue   left,  blue   right
              equal        → no colour either side
            """
            table = QTableWidget()
            table.setColumnCount(2)
            table.setHorizontalHeaderLabels([tgt_lbl, ref_lbl])
            table.setEditTriggers(QTableWidget.NoEditTriggers)
            table.setSelectionMode(QTableWidget.NoSelection)
            table.verticalHeader().setVisible(False)
            table.setWordWrap(False)
            table.setShowGrid(True)
            fnt = QFont("Consolas")
            fnt.setStyleHint(QFont.Monospace)
            fnt.setFixedPitch(True)
            table.setFont(fnt)
            table.setHorizontalScrollMode(QTableWidget.ScrollPerPixel)
            table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            hdr = table.horizontalHeader()
            hdr.setSectionResizeMode(0, QHeaderView.Stretch)
            hdr.setSectionResizeMode(1, QHeaderView.Stretch)

            new_lines = tgt_code.splitlines() if tgt_code else []
            ref_lines = ref_code.splitlines() if ref_code else []

            rows_diff = []
            if fn_tag == "added":
                # Function only in target — entire body is "added"
                for ln in new_lines:
                    rows_diff.append(("only_target", ln, ""))
            elif fn_tag == "deleted":
                # Function only in reference — entire body is "deleted"
                for ln in ref_lines:
                    rows_diff.append(("only_ref", "", ln))
            else:
                # Modified — line-by-line diff
                matcher = difflib.SequenceMatcher(None, new_lines, ref_lines)
                for op, i1, i2, j1, j2 in matcher.get_opcodes():
                    nc = new_lines[i1:i2]
                    rc = ref_lines[j1:j2]
                    if op == "equal":
                        for n, r in zip(nc, rc):
                            rows_diff.append(("equal", n, r))
                    elif op == "replace":
                        # A replaced block means lines changed between target and ref.
                        # Only pair lines as "modified" (blue) when BOTH sides have
                        # exactly 1 line AND the lines are genuinely similar (ratio >= 0.5).
                        # If they are completely different lines (e.g. difflib just happened
                        # to pair a comment with a variable declaration), emit them
                        # separately: target lines → green (only_target),
                        # ref lines → yellow (only_ref).
                        if len(nc) == 1 and len(rc) == 1:
                            similarity = difflib.SequenceMatcher(
                                None, nc[0].strip(), rc[0].strip()
                            ).ratio()
                            if similarity >= 0.5:
                                rows_diff.append(("modified", nc[0], rc[0]))
                            else:
                                rows_diff.append(("only_target", nc[0], ""))
                                rows_diff.append(("only_ref", "", rc[0]))
                        else:
                            for n in nc:
                                rows_diff.append(("only_target", n, ""))
                            for r in rc:
                                rows_diff.append(("only_ref", "", r))
                    elif op == "delete":
                        # Lines present in target but absent in ref → green (added in target)
                        for n in nc:
                            rows_diff.append(("only_target", n, ""))
                    elif op == "insert":
                        # Lines present in ref but absent in target → yellow (deleted from target)
                        for r in rc:
                            rows_diff.append(("only_ref", "", r))

            table.setRowCount(len(rows_diff))
            for row_idx, (rtag, nl, rl) in enumerate(rows_diff):
                # Fix 1: do NOT replace tabs — preserve original indentation
                ni = QTableWidgetItem(nl)
                ri = QTableWidgetItem(rl)
                ni.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                ri.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)

                # All three colours appear on BOTH cells so the row is clearly visible.
                # Guard: only apply a colour when the triggering side has actual content,
                # so that blank placeholder rows (e.g. empty lines from trailing newlines)
                # don't get accidentally painted.
                #   only_target → GREEN  on BOTH cells  (content added in target; rl is empty placeholder)
                #   only_ref    → YELLOW on BOTH cells  (content deleted from target; nl is empty placeholder)
                #                 *** skip if rl is blank — nothing to highlight ***
                #   modified    → BLUE   on BOTH cells  (content changed on both sides)
                #   equal       → no colour
                if rtag == "only_target" and nl.strip():
                    ni.setBackground(QColor("#2E7D32"))
                    ni.setForeground(QColor("#FFFFFF"))
                    ri.setBackground(QColor("#2E7D32"))
                    ri.setForeground(QColor("#FFFFFF"))
                elif rtag == "only_ref" and rl.strip():
                    # Only colour when the reference side actually has content
                    ni.setBackground(QColor("#F9A825"))
                    ni.setForeground(QColor("#000000"))
                    ri.setBackground(QColor("#F9A825"))
                    ri.setForeground(QColor("#000000"))
                elif rtag == "modified":
                    ni.setBackground(QColor("#1565C0"))
                    ni.setForeground(QColor("#FFFFFF"))
                    ri.setBackground(QColor("#1565C0"))
                    ri.setForeground(QColor("#FFFFFF"))
                # equal or empty placeholder → no colour on either side

                table.setItem(row_idx, 0, ni)
                table.setItem(row_idx, 1, ri)

            table.resizeRowsToContents()
            return table

        if not ref_copies:
            # Added function — no reference copy; show target body only
            table = _build_table(target_code, "", target_label, "Reference", diff_tag)
            self.diff_tabs.addTab(table, "{} vs Reference".format(target_label))
        else:
            # Issue 2: one tab PER reference folder
            for folder_path, ref_fp in ref_copies:
                ref_label = _make_ref_label(folder_path)
                ref_code  = (extract_function_body(ref_fp, function_name)
                             if ref_fp and os.path.isfile(ref_fp) else "")
                table = _build_table(target_code, ref_code, target_label, ref_label, diff_tag)
                tab_title = "{} vs {}".format(target_label, ref_label)
                self.diff_tabs.addTab(table, tab_title)

    # ── diff controls ─────────────────────────────────────────────────────────
    def _toggle_diff_fullscreen(self):
        self._diff_fullscreen = not self._diff_fullscreen
        if self._diff_fullscreen:
            self.sidebar.hide()
            self.header.hide()
            self._diff_left_panel.hide()
            self._diff_fs_btn.setText("Exit Fullscreen")
            # Hide filter buttons and clear button in fullscreen
            if hasattr(self, "_diff_filter_icon_lbl"):
                self._diff_filter_icon_lbl.setVisible(False)
            for btn in getattr(self, "_diff_filter_btns", {}).values():
                btn.setVisible(False)
            if hasattr(self, "_diff_clear_btn"):
                self._diff_clear_btn.setVisible(False)
            # Show info bar (file + func name) and prev/next buttons
            if hasattr(self, "_diff_fs_info_bar"):
                self._diff_fs_info_bar.setVisible(True)
            self._diff_update_fs_info()
            self._diff_connect_nav()
        else:
            self.sidebar.show()
            self.header.show()
            self._diff_left_panel.show()
            self._diff_fs_btn.setText("⛶  Fullscreen")
            # Restore filter buttons and clear button
            if hasattr(self, "_diff_filter_icon_lbl"):
                self._diff_filter_icon_lbl.setVisible(True)
            for btn in getattr(self, "_diff_filter_btns", {}).values():
                btn.setVisible(True)
            if hasattr(self, "_diff_clear_btn"):
                self._diff_clear_btn.setVisible(True)
            # Hide info bar
            if hasattr(self, "_diff_fs_info_bar"):
                self._diff_fs_info_bar.setVisible(False)

    def _diff_update_fs_info(self):
        """Update the fullscreen info bar labels with current file/function."""
        item = self.diff_tree.currentItem()
        if item and item.data(0, Qt.UserRole):
            payload = item.data(0, Qt.UserRole)
            if payload.get("type") == "function":
                fn  = payload.get("function_name", "")
                fp  = payload.get("file_path", "")
                fname = os.path.basename(fp) if fp else ""
                if hasattr(self, "_diff_fs_file_label"):
                    self._diff_fs_file_label.setText(f"📄 {fname}")
                if hasattr(self, "_diff_fs_func_label"):
                    self._diff_fs_func_label.setText(f"ƒ  {fn}")
                return
        if hasattr(self, "_diff_fs_file_label"):
            self._diff_fs_file_label.setText("")
        if hasattr(self, "_diff_fs_func_label"):
            self._diff_fs_func_label.setText("")

    def _diff_connect_nav(self):
        """Wire Prev/Next buttons for keyboard-free navigation in fullscreen."""
        try:
            self._diff_prev_btn.clicked.disconnect()
            self._diff_next_btn.clicked.disconnect()
        except Exception:
            pass
        self._diff_prev_btn.clicked.connect(lambda: self._diff_navigate(-1))
        self._diff_next_btn.clicked.connect(lambda: self._diff_navigate(1))
        self._diff_update_nav_state()

    def _diff_get_all_function_items(self):
        """Return flat list of all visible function-level QTreeWidgetItems."""
        items = []
        root = self.diff_tree.invisibleRootItem()
        for i in range(root.childCount()):
            file_item = root.child(i)
            for j in range(file_item.childCount()):
                fn_item = file_item.child(j)
                payload = fn_item.data(0, Qt.UserRole)
                if payload and payload.get("type") == "function":
                    items.append(fn_item)
        return items

    def _diff_navigate(self, direction: int):
        """Navigate to the next (+1) or previous (-1) function item."""
        items = self._diff_get_all_function_items()
        if not items:
            return
        current = self.diff_tree.currentItem()
        try:
            idx = items.index(current)
        except ValueError:
            idx = -1
        new_idx = idx + direction
        if new_idx < 0 or new_idx >= len(items):
            return
        new_item = items[new_idx]
        self.diff_tree.setCurrentItem(new_item)
        self.diff_tree.scrollToItem(new_item)
        self.on_diff_item_clicked(new_item, 0)
        self._diff_update_fs_info()
        self._diff_update_nav_state()

    def _diff_update_nav_state(self):
        """Enable/disable Prev/Next based on current position."""
        if not hasattr(self, "_diff_prev_btn"):
            return
        items = self._diff_get_all_function_items()
        current = self.diff_tree.currentItem()
        try:
            idx = items.index(current)
        except ValueError:
            idx = -1
        self._diff_prev_btn.setEnabled(idx > 0)
        self._diff_next_btn.setEnabled(0 <= idx < len(items) - 1)

    def _clear_diff(self):
        self.diff_tabs.clear()
        self.diff_tree.clear()
        self.diff_search_box.clear()
        # Point 6: reset filter buttons when diff is cleared
        self._diff_active_filter = None
        if hasattr(self, "_diff_filter_btns") and hasattr(self, "_diff_style_btn"):
            for k, btn in self._diff_filter_btns.items():
                color = btn.property("activeColor")
                fg    = btn.property("activeFg")
                btn.setChecked(False)
                self._diff_style_btn(btn, active=False, greyed=False, color=color, fg=fg)
        with SCAN_CACHE_LOCK:
            SCAN_CACHE.clear()

    def _apply_diff_filter(self):
        """Rebuild the diff tree from raw data, applying the active tag filter
        and search text.  Rebuilding from scratch (instead of hide/show) is the
        only reliable way to guarantee Qt renders exactly what we want.
        Does NOT touch diff_tabs so the open diff view is preserved.

        Filter semantics (matches screenshot behaviour):
          "Added"    → show functions that contain ANY added lines in their diff
                       (only_target lines in a modified function, or purely deleted
                        functions are excluded — they have no added lines)
          "Deleted"  → show functions that contain ANY deleted lines (only_ref lines)
          "Modified" → show functions that contain ANY modified (blue) line pairs
        A function tagged "modified" can simultaneously satisfy all three filters
        depending on which line types its diff contains.
        """
        active      = getattr(self, "_diff_active_filter", None)
        search_text = self.diff_search_box.text().strip().lower()

        TAG_COLORS = {
            "added":    ("#2E7D32", "#FFFFFF"),
            "deleted":  ("#F9A825", "#000000"),
            "modified": ("#1565C0", "#FFFFFF"),
            # "equal" has no color — rendered with default background
        }

        # ── Rebuild tree from stored raw data ────────────────────────────────
        self.diff_tree.clear()

        for entry in getattr(self, "_diff_raw_data", []):
            dn        = entry["display_name"]
            file_path = entry["file_path"]
            all_fns   = entry["functions"]
            # Tuple layout: (func_name, tag, tgt_fp, ref_copies,
            #                has_added, has_deleted, has_modified)

            # Filter functions by active filter AND search text.
            # The filter matches against LINE-LEVEL diff content, not the
            # function-level tag, so that e.g. "Added" shows modified functions
            # that contain added lines.
            visible_fns = []
            for fn_tuple in all_fns:
                func_name, tag, tgt_fp, ref_copies = fn_tuple[0], fn_tuple[1], fn_tuple[2], fn_tuple[3]
                has_added    = fn_tuple[4] if len(fn_tuple) > 4 else (tag == "added")
                has_deleted  = fn_tuple[5] if len(fn_tuple) > 5 else (tag == "deleted")
                has_modified = fn_tuple[6] if len(fn_tuple) > 6 else (tag == "modified")

                if active is None:
                    # No filter active → show everything including equal functions
                    tag_ok = True
                elif tag == "equal":
                    # Equal functions are never shown when a filter button is active
                    tag_ok = False
                elif active == "added":
                    # Show functions that have added lines OR are purely deleted
                    # (deleted functions show as green "only_target" is N/A;
                    #  actually deleted functions have only_ref lines so they
                    #  appear under "Deleted" — keep consistent)
                    tag_ok = has_added
                elif active == "deleted":
                    tag_ok = has_deleted
                elif active == "modified":
                    tag_ok = has_modified
                else:
                    tag_ok = (tag == active)

                search_ok = (not search_text) or (
                    search_text in func_name.lower() or
                    search_text in dn.lower()
                )
                if tag_ok and search_ok:
                    visible_fns.append((func_name, tag, tgt_fp, ref_copies))

            # Skip this file entirely if nothing passes the filter
            if not visible_fns:
                continue

            parent = QTreeWidgetItem([dn])
            parent.setData(0, Qt.UserRole, {
                "type":      "file",
                "file_path": file_path,
            })

            for func_name, tag, tgt_fp, ref_copies in visible_fns:
                child = QTreeWidgetItem([func_name])
                child.setData(0, Qt.UserRole, {
                    "type":          "function",
                    "file_path":     tgt_fp,
                    "ref_copies":    ref_copies,
                    "function_name": func_name,
                    "diff_tag":      tag,
                })
                parent.addChild(child)

            self.diff_tree.addTopLevelItem(parent)

        self.diff_tree.expandAll()

    # ── submit reset helper ───────────────────────────────────────────────────
    def _reset_view_and_diff(self):
        """Clear view page and diff page state fully before a new submit."""
        # Clear scan cache so re-upload always reads fresh files
        with SCAN_CACHE_LOCK:
            SCAN_CACHE.clear()

        # Reset core data
        self.function_records = OrderedDict()
        self.available_sources = []
        self.current_function_list_paths = []
        self.current_reference_folders = []
        self._ref_function_filter = []
        self.current_function_name = ""
        self.current_function_file = ""
        self.current_function_body = ""

        # Reset view page widgets
        if hasattr(self, "source_combo"):
            self.source_combo.blockSignals(True)
            self.source_combo.clear()
            self.source_combo.blockSignals(False)
        if hasattr(self, "tree"):
            self.tree.clear()
        if hasattr(self, "view_text"):
            self.view_text.setText("")
        if hasattr(self, "view_title"):
            self.view_title.setText("")
        if hasattr(self, "view_meta"):
            self.view_meta.setText("")
        if hasattr(self, "view_mode_chip"):
            self.view_mode_chip.setVisible(False)

        # Reset diff page completely
        self._clear_diff()

    # ── submit reference ──────────────────────────────────────────────────────
    def submit_reference(self):
        target_folders = self.ref_target_field.value()
        ref_folders    = self.ref_bases_field.value()
        function_list  = self.ref_function_field.value()

        if not target_folders:
            QMessageBox.warning(self, "Missing Input", "Please select Target Base Folder.")
            return
        if not ref_folders:
            QMessageBox.warning(self, "Missing Input", "Please select at least one Reference Base Folder.")
            return

        # Show submitting state
        self.ref_submit_btn.setEnabled(False)
        self.ref_submit_btn.setText("Submitting...")
        self.show_loading("Scanning source folders…")

        try:
            # Fix 2: full reset before loading new data
            self._reset_view_and_diff()

            target_root = target_folders[0]
            self.current_target_folders = ref_folders

            self._ref_function_filter = parse_function_list_files(function_list) if function_list else []

            entries = [self.build_source_entry("target", target_root)]
            for folder in ref_folders:
                entries.append(self.build_source_entry("reference", folder))
            self.register_sources(entries, function_list, ref_folders)

            ref_text = "\n".join(ref_folders) or "No reference folders selected"
            fn_text  = "\n".join(function_list) or "No function list selected"
            self.hide_loading()
            QMessageBox.information(self, "Success",
                f"Reference data loaded successfully.\n\nTarget Base Folder:\n{target_root}\n\n"
                f"Reference Base Folders:\n{ref_text}\n\nFunction List:\n{fn_text}\n\n"
                f"Dropdown Sources: {len(entries)}")
        finally:
            self.hide_loading()
            self.ref_submit_btn.setEnabled(True)
            self.ref_submit_btn.setText("Submit")

    # ── submit consolidated ───────────────────────────────────────────────────
    def submit_consolidated(self):
        # ── Determine source mode ─────────────────────────────────────────────
        use_folder_mode = hasattr(self, "con_toggle_folder_btn") and \
                          self.con_toggle_folder_btn.isChecked()

        # Fix 2: full reset before loading new consolidated data
        self._reset_view_and_diff()

        if use_folder_mode:
            # Folder mode: extract functions to temp excel, auto-detect func col
            target_folder = self.con_folder_field.value()
            if not target_folder or not os.path.isdir(target_folder):
                QMessageBox.warning(self, "Missing Input",
                    "Please select a Target Folder to extract functions from."); return

            consolidated_excel = self.con_db_excel_field.value()
            if not consolidated_excel:
                QMessageBox.warning(self, "Missing Input",
                    "Please select Consolidated DB Excel file."); return

            base_col_ref = self.con_base_col_field.value()
            if not base_col_ref or not is_valid_excel_reference(base_col_ref):
                QMessageBox.warning(self, "Invalid Input",
                    "Base Col must be detected or entered like B2, C1, G3, etc."); return

            # Extract to temp Excel
            try:
                temp_excel_path, fn_count = extract_functions_from_folder_to_excel(target_folder)
            except Exception as exc:
                QMessageBox.critical(self, "Extraction Error",
                    f"Failed to extract functions from folder:\n{exc}"); return

            if fn_count == 0:
                QMessageBox.warning(self, "No Functions Found",
                    "No functions were detected in the selected folder."); return

            # Auto-detect function name column from the generated temp excel
            # It's always column C ("Function Name"), row 1 → reference = "C1"
            func_col_ref = "C1"
            # Reflect detected value in the (disabled) UI field
            self.con_func_col_field.input.setText(func_col_ref)
            self.con_func_col_field.preview.setText("Preview: auto-detected (folder mode)")

            function_list_files = [temp_excel_path]

        else:
            # Normal excel-list mode
            function_list_files = self.con_function_field.value()
            consolidated_excel  = self.con_db_excel_field.value()
            func_col_ref        = self.con_func_col_field.value()
            base_col_ref        = self.con_base_col_field.value()

            if not function_list_files:
                QMessageBox.warning(self, "Missing Input",
                    "Please select one or more Function List files."); return
            if not consolidated_excel:
                QMessageBox.warning(self, "Missing Input",
                    "Please select Consolidated DB Excel file."); return
            if not func_col_ref or not is_valid_excel_reference(func_col_ref):
                QMessageBox.warning(self, "Invalid Input",
                    "Func Col must be detected or entered like A2, B1, F3, etc."); return
            if not base_col_ref or not is_valid_excel_reference(base_col_ref):
                QMessageBox.warning(self, "Invalid Input",
                    "Base Col must be detected or entered like B2, C1, G3, etc."); return

        preferred_sheet = None
        if getattr(self.con_func_col_field, "detected_info", None):
            preferred_sheet = self.con_func_col_field.detected_info.get("sheet")
        elif getattr(self.con_base_col_field, "detected_info", None):
            preferred_sheet = self.con_base_col_field.detected_info.get("sheet")

        self.con_submit_btn.setEnabled(False)
        self.con_submit_btn.setText("Submitting...")
        self.con_output_link_field.clear_selection()
        self.show_loading("Processing consolidated data…")

        self.con_thread = QThread(parent=self)
        self.con_worker = ConsolidatedWorker(
            function_list_files, consolidated_excel, func_col_ref, base_col_ref,
            preferred_sheet=preferred_sheet)
        self.con_worker.moveToThread(self.con_thread)
        self.con_thread.started.connect(self.con_worker.run)
        self.con_worker.finished.connect(self.on_consolidated_finished)
        self.con_worker.error.connect(self.on_consolidated_error)
        self.con_worker.finished.connect(self.con_thread.quit)
        self.con_worker.error.connect(self.con_thread.quit)
        self.con_worker.finished.connect(self.con_worker.deleteLater)
        self.con_worker.error.connect(self.con_worker.deleteLater)
        self.con_thread.finished.connect(self.con_thread.deleteLater)
        self.con_thread.finished.connect(
            lambda t=self.con_thread: self._active_threads.remove(t)
            if t in self._active_threads else None)
        self._active_threads.append(self.con_thread)
        self.con_thread.start()

    def on_consolidated_finished(self, result: dict):
        self.hide_loading()
        self.con_submit_btn.setEnabled(True)
        self.con_submit_btn.setText("Submit")
        self.con_output_link_field.set_output(result["output_file"])
        QMessageBox.information(self, "Success",
            f"Output Excel generated successfully.\n\n"
            f"Function List Files: {result['function_list_count']}\n"
            f"Functions Read: {result['functions_read']}\n"
            f"Matched Functions In Output: {result['matched_count']}\n"
            f"Unmatched Parsed Functions: {result.get('unmatched_count', 0)}\n"
            f"Output File:\n{result['output_file']}")
        self.con_worker = None; self.con_thread = None

    def on_consolidated_error(self, message: str):
        self.hide_loading()
        self.con_submit_btn.setEnabled(True)
        self.con_submit_btn.setText("Submit")
        QMessageBox.critical(self, "Processing Error", message)
        self.con_worker = None; self.con_thread = None

    # ── report folder picker ──────────────────────────────────────────────────
    def _pick_report_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder", "",
                    QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks)
        if not folder:
            return
        folder = normalize_path(folder)
        self._report_folder_path.clear()
        self._report_folder_path.append(folder)
        self._report_folder_display.setText(folder)

    # ── report slots ──────────────────────────────────────────────────────────
    def _set_report_progress(self, pct: int, status_text: str = ""):
        """Set the OVERALL cumulative progress bar (0-100) directly."""
        pct = max(0, min(100, int(pct)))
        self.report_progress_bar.setValue(pct)
        self.report_pct_label.setText(f"{pct} %")
        if status_text:
            self.report_status_label.setText(status_text)
            if hasattr(self, "report_summary_chips"):
                self.report_summary_chips["phase"].set_value(status_text[:22])

    def _cumulative_progress(self, current_step_pct: int) -> int:
        """
        Compute cumulative overall progress.
        Each completed step contributes its full share; the active step
        contributes a fractional share based on current_step_pct (0-100).
        """
        total = getattr(self, "_report_total_steps", 1) or 1
        done  = getattr(self, "_report_done_steps",  0)
        overall = (done + current_step_pct / 100.0) / total * 100
        return max(0, min(100, int(overall)))

    def _update_report_step_chip(self):
        if hasattr(self, "report_summary_chips") and hasattr(self, "_report_total_steps"):
            self.report_summary_chips["steps"].set_value(str(self._report_total_steps or 0))

    def _on_base_extraction_started(self, label: str):
        from ui.widgets import StepStatusWidget
        self.report_phase_label.setText(f"Extracting — {label}")
        overall = self._cumulative_progress(0)
        self._set_report_progress(overall, f"{label}: starting …")
        self.report_step_widget.set_state(label, StepStatusWidget.STATE_RUNNING,
                                          "Scanning source files …", pct=0)

    def _on_base_extraction_progress(self, label: str, pct: int, detail: str):
        from ui.widgets import StepStatusWidget
        self.report_phase_label.setText(f"Extracting — {label}")
        overall = self._cumulative_progress(pct)
        self._set_report_progress(overall, f"{label}: {detail}")
        self.report_step_widget.set_state(label, StepStatusWidget.STATE_RUNNING,
                                          detail, pct=pct)

    def _on_compare_progress(self, pct: int, detail: str):
        from ui.widgets import StepStatusWidget
        self.report_phase_label.setText("Comparing & Writing Excel Report")
        overall = self._cumulative_progress(pct)
        self._set_report_progress(overall, detail)
        self.report_step_widget.set_state("📊 Comparing & Writing Excel Report",
                                          StepStatusWidget.STATE_RUNNING, detail, pct=pct)

    def _on_report_generate(self):
        from ui.widgets import StepStatusWidget
        output_root = self.report_output_field.value()
        if not output_root:
            QMessageBox.warning(self, "Missing Folder", "Please select an output folder.")
            return
        if not self.available_sources:
            QMessageBox.warning(self, "No Data Loaded",
                "Please submit reference bases first (Input → Reference Bases).")
            return

        target_entry = next((s for s in self.available_sources if s["type"] == "target"), None)
        if not target_entry:
            QMessageBox.warning(self, "No Target", "No target base found in submitted sources.")
            return

        ref_entries = [s for s in self.available_sources if s["type"] == "reference"]
        bases = [{"label": target_entry["label"], "src_path": target_entry["path"], "is_target": True}]
        for ref in ref_entries:
            bases.append({"label": ref["label"], "src_path": ref["path"], "is_target": False})

        self._report_bases_list  = bases
        self._report_total_steps = len(bases) + 1
        self._report_done_steps  = 0
        self._update_report_step_chip()
        self._report_target_label = target_entry["label"]
        self._report_ref_labels   = [r["label"] for r in ref_entries]
        self._report_output_root  = output_root

        self._set_report_progress(0, "Preparing extraction …")
        self.report_log_box.clear()
        self.report_phase_label.setText("Extracting — waiting to start")
        self.report_generate_btn.setEnabled(False)
        self.report_generate_btn.setText("Running …")
        self.report_open_btn.setEnabled(False)
        self._report_output_file = ""
        if hasattr(self, "report_summary_chips"):
            self.report_summary_chips["phase"].set_value("Starting")
            self.report_summary_chips["result"].set_value("Pending")
        self.show_loading("Generating report…")

        import shutil as _shutil
        with SCAN_CACHE_LOCK:
            SCAN_CACHE.clear()
        _extracted_root = os.path.join(output_root, 'FuncAtlas_Extracted')
        if os.path.isdir(_extracted_root):
            try:
                _shutil.rmtree(_extracted_root)
                self.report_log_box.append("🗑 Cleared previous extraction cache.")
            except Exception as _ce:
                self.report_log_box.append(f"Warning: could not clear cache: {_ce}")

        self.report_step_widget.clear_steps()
        for b in bases:
            self.report_step_widget.add_step(b["label"])
        self.report_step_widget.add_step("📊 Comparing & Writing Excel Report")

        self._report_ext_thread = QThread(parent=self)
        self._report_ext_worker = BuiltinExtractionWorker(
            bases, output_root,
            function_filter=self._ref_function_filter if self._ref_function_filter else None
        )
        self._report_ext_worker.moveToThread(self._report_ext_thread)
        self._report_ext_thread.started.connect(self._report_ext_worker.run)
        self._report_ext_worker.base_started.connect(self._on_base_extraction_started)
        self._report_ext_worker.base_progress.connect(self._on_base_extraction_progress)
        self._report_ext_worker.step_done.connect(self._on_step_extracted)
        self._report_ext_worker.log.connect(self._report_append_log)
        self._report_ext_worker.finished.connect(self._on_extraction_done)
        self._report_ext_worker.error.connect(self._on_report_error)
        self._report_ext_worker.finished.connect(self._report_ext_thread.quit)
        self._report_ext_worker.error.connect(self._report_ext_thread.quit)
        self._report_ext_worker.finished.connect(self._report_ext_worker.deleteLater)
        self._report_ext_thread.finished.connect(self._report_ext_thread.deleteLater)
        self._report_ext_thread.finished.connect(
            lambda t=self._report_ext_thread: self._active_threads.remove(t)
            if t in self._active_threads else None)
        self._active_threads.append(self._report_ext_thread)
        self._report_ext_thread.start()

    def _on_step_extracted(self, label: str, func_count: int):
        from ui.widgets import StepStatusWidget
        self._report_done_steps = getattr(self, "_report_done_steps", 0) + 1
        self.report_step_widget.set_state(label, StepStatusWidget.STATE_DONE,
                                          f"{func_count} functions extracted")
        overall = self._cumulative_progress(100)
        self._set_report_progress(overall, f"{label}: completed")

    def _on_extraction_done(self, results: dict):
        from ui.widgets import StepStatusWidget
        self._report_append_log("─── All extractions complete. Starting comparison … ───")
        self.report_phase_label.setText("Comparing & Writing Excel Report")
        overall = self._cumulative_progress(0)
        self._set_report_progress(overall, "Preparing comparison …")
        self.report_step_widget.set_state("📊 Comparing & Writing Excel Report",
                                          StepStatusWidget.STATE_RUNNING, "Preparing comparison …")

        target_folder = results.get(self._report_target_label, "")
        ref_folders   = [results[lbl] for lbl in self._report_ref_labels if lbl in results]

        if not target_folder or not os.path.isdir(target_folder):
            self._on_report_error(f"Target extraction folder not found:\n{target_folder}")
            return

        _bases_list   = getattr(self, '_report_bases_list', [])
        _target_src   = next((b['src_path'] for b in _bases_list if b['label'] == self._report_target_label), '')
        _ref_srcs     = [next((b['src_path'] for b in _bases_list if b['label'] == lbl), '')
                         for lbl in self._report_ref_labels]

        self._report_cmp_thread = QThread(parent=self)
        self._report_cmp_worker = ReportCompareWorker(
            target_label    = self._report_target_label,
            target_folder   = target_folder,
            ref_labels      = self._report_ref_labels,
            ref_folders     = ref_folders,
            output_root     = self._report_output_root,
            target_src_path = _target_src,
            ref_src_paths   = _ref_srcs,
            # Pass user-configured complexity settings from Complexity page
            weights         = getattr(self, '_complexity_weights', None),
            bands           = getattr(self, '_complexity_bands',   None),
        )
        self._report_cmp_worker.moveToThread(self._report_cmp_thread)
        self._report_cmp_thread.started.connect(self._report_cmp_worker.run)
        self._report_cmp_worker.progress.connect(self._on_compare_progress)
        self._report_cmp_worker.log.connect(self._report_append_log)
        self._report_cmp_worker.finished.connect(self._on_compare_done)
        self._report_cmp_worker.error.connect(self._on_report_error)
        self._report_cmp_worker.finished.connect(self._report_cmp_thread.quit)
        self._report_cmp_worker.error.connect(self._report_cmp_thread.quit)
        self._report_cmp_worker.finished.connect(self._report_cmp_worker.deleteLater)
        self._report_cmp_thread.finished.connect(self._report_cmp_thread.deleteLater)
        self._report_cmp_thread.finished.connect(
            lambda t=self._report_cmp_thread: self._active_threads.remove(t)
            if t in self._active_threads else None)
        self._active_threads.append(self._report_cmp_thread)
        self._report_cmp_thread.start()

    # ── HTML report generation ────────────────────────────────────────────────
    def _on_report_generate_html(self):
        """Same pipeline as _on_report_generate but produces an HTML report."""
        from ui.widgets import StepStatusWidget
        output_root = self.report_output_field.value()
        if not output_root:
            QMessageBox.warning(self, "Missing Folder", "Please select an output folder.")
            return
        if not self.available_sources:
            QMessageBox.warning(self, "No Data Loaded",
                "Please submit reference bases first (Input → Reference Bases).")
            return

        target_entry = next((s for s in self.available_sources if s["type"] == "target"), None)
        if not target_entry:
            QMessageBox.warning(self, "No Target", "No target base found in submitted sources.")
            return

        ref_entries = [s for s in self.available_sources if s["type"] == "reference"]
        bases = [{"label": target_entry["label"], "src_path": target_entry["path"], "is_target": True}]
        for ref in ref_entries:
            bases.append({"label": ref["label"], "src_path": ref["path"], "is_target": False})

        self._report_bases_list   = bases
        self._report_total_steps  = len(bases) + 1
        self._update_report_step_chip()
        self._report_target_label = target_entry["label"]
        self._report_ref_labels   = [r["label"] for r in ref_entries]
        self._report_output_root  = output_root
        self._report_html_mode    = True   # flag so _on_extraction_done_html is used

        self._set_report_progress(0, "Preparing extraction …")
        self.report_log_box.clear()
        self.report_phase_label.setText("Extracting — waiting to start")
        self.report_generate_html_btn.setEnabled(False)
        self.report_generate_html_btn.setText("Running …")
        self.report_open_btn.setEnabled(False)
        self._report_output_file = ""
        if hasattr(self, "report_summary_chips"):
            self.report_summary_chips["phase"].set_value("Starting")
            self.report_summary_chips["result"].set_value("Pending")
        self.show_loading("Generating HTML report…")

        import shutil as _shutil
        with SCAN_CACHE_LOCK:
            SCAN_CACHE.clear()
        _extracted_root = os.path.join(output_root, 'FuncAtlas_Extracted')
        if os.path.isdir(_extracted_root):
            try:
                _shutil.rmtree(_extracted_root)
                self.report_log_box.append("🗑 Cleared previous extraction cache.")
            except Exception as _ce:
                self.report_log_box.append(f"Warning: could not clear cache: {_ce}")

        self.report_step_widget.clear_steps()
        for b in bases:
            self.report_step_widget.add_step(b["label"])
        self.report_step_widget.add_step("📊 Comparing & Writing HTML Report")

        self._report_ext_thread = QThread(parent=self)
        self._report_ext_worker = BuiltinExtractionWorker(
            bases, output_root,
            function_filter=self._ref_function_filter if self._ref_function_filter else None
        )
        self._report_ext_worker.moveToThread(self._report_ext_thread)
        self._report_ext_thread.started.connect(self._report_ext_worker.run)
        self._report_ext_worker.base_started.connect(self._on_base_extraction_started)
        self._report_ext_worker.base_progress.connect(self._on_base_extraction_progress)
        self._report_ext_worker.step_done.connect(self._on_step_extracted)
        self._report_ext_worker.log.connect(self._report_append_log)
        self._report_ext_worker.finished.connect(self._on_extraction_done_html)
        self._report_ext_worker.error.connect(self._on_report_error_html)
        self._report_ext_worker.finished.connect(self._report_ext_thread.quit)
        self._report_ext_worker.error.connect(self._report_ext_thread.quit)
        self._report_ext_worker.finished.connect(self._report_ext_worker.deleteLater)
        self._report_ext_thread.finished.connect(self._report_ext_thread.deleteLater)
        self._report_ext_thread.finished.connect(
            lambda t=self._report_ext_thread: self._active_threads.remove(t)
            if t in self._active_threads else None)
        self._active_threads.append(self._report_ext_thread)
        self._report_ext_thread.start()

    def _on_extraction_done_html(self, results: dict):
        from ui.widgets import StepStatusWidget
        self._report_append_log("─── All extractions complete. Starting HTML comparison … ───")
        self.report_phase_label.setText("Comparing & Writing HTML Report")
        self._set_report_progress(0, "Preparing comparison …")
        self.report_step_widget.set_state("📊 Comparing & Writing HTML Report",
                                          StepStatusWidget.STATE_RUNNING, "Preparing comparison …")

        target_folder = results.get(self._report_target_label, "")
        ref_folders   = [results[lbl] for lbl in self._report_ref_labels if lbl in results]

        if not target_folder or not os.path.isdir(target_folder):
            self._on_report_error_html(f"Target extraction folder not found:\n{target_folder}")
            return

        _bases_list  = getattr(self, '_report_bases_list', [])
        _target_src  = next((b['src_path'] for b in _bases_list if b['label'] == self._report_target_label), '')
        _ref_srcs    = [next((b['src_path'] for b in _bases_list if b['label'] == lbl), '')
                        for lbl in self._report_ref_labels]

        self._report_cmp_thread = QThread(parent=self)
        self._report_cmp_worker = ReportCompareWorker(
            target_label    = self._report_target_label,
            target_folder   = target_folder,
            ref_labels      = self._report_ref_labels,
            ref_folders     = ref_folders,
            output_root     = self._report_output_root,
            target_src_path = _target_src,
            ref_src_paths   = _ref_srcs,
            weights         = getattr(self, '_complexity_weights', None),
            bands           = getattr(self, '_complexity_bands',   None),
        )
        self._report_cmp_worker.moveToThread(self._report_cmp_thread)
        self._report_cmp_thread.started.connect(self._report_cmp_worker.run)
        self._report_cmp_worker.progress.connect(self._on_compare_progress)
        self._report_cmp_worker.log.connect(self._report_append_log)
        self._report_cmp_worker.finished.connect(self._on_compare_done_html)
        self._report_cmp_worker.error.connect(self._on_report_error_html)
        self._report_cmp_worker.finished.connect(self._report_cmp_thread.quit)
        self._report_cmp_worker.error.connect(self._report_cmp_thread.quit)
        self._report_cmp_worker.finished.connect(self._report_cmp_worker.deleteLater)
        self._report_cmp_thread.finished.connect(self._report_cmp_thread.deleteLater)
        self._report_cmp_thread.finished.connect(
            lambda t=self._report_cmp_thread: self._active_threads.remove(t)
            if t in self._active_threads else None)
        self._active_threads.append(self._report_cmp_thread)
        self._report_cmp_thread.start()

    def _on_compare_done_html(self, out_excel: str):
        """Convert the generated Excel to HTML — ask user where to save."""
        from ui.widgets import StepStatusWidget
        default_name = os.path.splitext(os.path.basename(out_excel))[0] + "_FuncAtlas_Report.html"
        default_dir  = os.path.dirname(out_excel)
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Save HTML Report As", os.path.join(default_dir, default_name),
            "HTML Files (*.html);;All Files (*.*)"
        )
        if not save_path:
            self.report_step_widget.set_state("📊 Comparing & Writing HTML Report",
                                              StepStatusWidget.STATE_DONE, "Cancelled by user")
            self._set_report_progress(100, "HTML save cancelled")
            self.report_generate_html_btn.setEnabled(True)
            self.report_generate_html_btn.setText("Generate HTML Report")
            self.hide_loading()
            return
        try:
            html_path = self._write_html_from_excel(out_excel, save_path=save_path)
        except Exception as exc:
            self._on_report_error_html(f"HTML conversion failed: {exc}")
            return

        self._report_output_file = html_path
        self._last_report_excel  = out_excel
        self.report_step_widget.set_state("📊 Comparing & Writing HTML Report",
                                          StepStatusWidget.STATE_DONE, "HTML report saved")
        self._set_report_progress(100, "HTML report ready")
        self.report_phase_label.setText("✅ Complete")
        self.report_generate_html_btn.setEnabled(True)
        self.report_generate_html_btn.setText("Generate HTML Report")
        self.report_open_btn.setEnabled(True)
        if hasattr(self, "report_summary_chips"):
            self.report_summary_chips["phase"].set_value("Completed")
            self.report_summary_chips["result"].set_value("HTML ready")
        self._report_append_log(f"✓ HTML Report: {html_path}")
        self.hide_loading()
        # Auto-open the HTML report in the default browser
        from PySide6.QtCore import QUrl
        QDesktopServices.openUrl(QUrl.fromLocalFile(html_path))

    def _on_complexity_generate_html(self):
        """Generate HTML report from the Excel on the Complexity page — ask where to save."""
        report_path = getattr(self, 'complexity_report_display', None)
        if report_path is None:
            QMessageBox.warning(self, "Not Available", "Complexity report display not found.")
            return
        path = report_path.text().strip()
        if not path or not os.path.isfile(path):
            QMessageBox.warning(self, "No Report",
                "Please select or generate an Excel report first\n"
                "(use Browse Report or Generate Report on this page).")
            return
        default_name = os.path.splitext(os.path.basename(path))[0] + "_FuncAtlas_Report.html"
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Save HTML Report As",
            os.path.join(os.path.dirname(path), default_name),
            "HTML Files (*.html);;All Files (*.*)"
        )
        if not save_path:
            return
        self.show_loading("Generating HTML report…")
        try:
            html_path = self._write_html_from_excel(path, save_path=save_path)
        except Exception as exc:
            self.hide_loading()
            QMessageBox.critical(self, "HTML Report Error", f"Failed to generate HTML:\n{exc}")
            return
        self.hide_loading()
        self._last_complexity_html = html_path
        QMessageBox.information(self, "HTML Report Ready",
            f"HTML report saved successfully.\n\nFile:\n{html_path}")
        # Auto-open the HTML report after generation
        from PySide6.QtCore import QUrl
        QDesktopServices.openUrl(QUrl.fromLocalFile(html_path))

    def _write_html_from_excel(self, excel_path: str, save_path: str = None) -> str:
        """
        Read all four sheets of FuncAtlas_Report.xlsx and produce a
        styled standalone HTML report:
          • Summary block at top (from Sheet 2 'Summary')
          • Detail table: S.No | File Name | Function Name | Reuse/New |
                          Which Base | Complexity Level | Compatibility %
            (joined from Sheet1, Sheet3, Sheet4)
        """
        import openpyxl

        wb = openpyxl.load_workbook(excel_path, data_only=True)

        # ── Sheet 2: Summary ──────────────────────────────────────────────────
        summary_rows = []
        if "Summary" in wb.sheetnames:
            ws2 = wb["Summary"]
            for row in ws2.iter_rows(min_row=2, values_only=True):
                if row and row[0] is not None:
                    summary_rows.append((str(row[0]), str(row[1]) if row[1] is not None else ""))

        # ── Sheet 1: Function_Match_Report ────────────────────────────────────
        # Columns: File Name | Function Name | Target File Path | Ref Match%... | Reuse/New | Which Base | Ref file path
        sheet1_data = {}   # key = (file_name, func_name) → {reuse_status, which_base}
        if "Function_Match_Report" in wb.sheetnames:
            ws1 = wb["Function_Match_Report"]
            hdrs1 = [str(c).strip() if c else "" for c in next(ws1.iter_rows(min_row=1, max_row=1, values_only=True))]
            # locate key columns by header name
            try:
                col_file   = hdrs1.index("File Name")
                col_func   = hdrs1.index("Function Name")
                col_status = hdrs1.index("Reuse/New")
                col_which  = hdrs1.index("Suggested Reference Base")
            except ValueError:
                col_file, col_func, col_status, col_which = 0, 1, len(hdrs1)-3, len(hdrs1)-2
            for row in ws1.iter_rows(min_row=2, values_only=True):
                if not row or row[col_func] is None:
                    continue
                key = (str(row[col_file] or ""), str(row[col_func] or ""))
                sheet1_data[key] = {
                    "reuse_status": str(row[col_status] or ""),
                    "which_base":   str(row[col_which]  or ""),
                }

        # ── Sheet 3: Complexity_Compatibility ─────────────────────────────────
        # Columns: File Name | Function Name | Target File Path | ...constructs... | Complexity Score | Complexity Level
        sheet3_data = {}   # key = (file_name, func_name) → complexity_level
        if "Complexity_Compatibility" in wb.sheetnames:
            ws3 = wb["Complexity_Compatibility"]
            hdrs3 = [str(c).strip() if c else "" for c in next(ws3.iter_rows(min_row=1, max_row=1, values_only=True))]
            try:
                c3_file  = hdrs3.index("File Name")
                c3_func  = hdrs3.index("Function Name")
                c3_level = hdrs3.index("Complexity Level")
            except ValueError:
                c3_file, c3_func, c3_level = 0, 1, len(hdrs3) - 1
            for row in ws3.iter_rows(min_row=2, values_only=True):
                if not row or row[c3_func] is None:
                    continue
                key = (str(row[c3_file] or ""), str(row[c3_func] or ""))
                sheet3_data[key] = str(row[c3_level] or "")

        # ── Sheet 4: Compatibility_Score ──────────────────────────────────────
        # Columns: File Name | Function Name | File Path | Available Scenarios | Handled | Unhandled | Compatibility %
        sheet4_data = {}   # key = (file_name, func_name) → compat_pct string
        if "Compatibility_Score" in wb.sheetnames:
            ws4 = wb["Compatibility_Score"]
            hdrs4 = [str(c).strip() if c else "" for c in next(ws4.iter_rows(min_row=1, max_row=1, values_only=True))]
            try:
                c4_file  = hdrs4.index("File Name")
                c4_func  = hdrs4.index("Function Name")
                c4_compat= hdrs4.index("Compatibility %")
            except ValueError:
                c4_file, c4_func, c4_compat = 0, 1, len(hdrs4) - 1
            for row in ws4.iter_rows(min_row=2, values_only=True):
                if not row or row[c4_func] is None:
                    continue
                key = (str(row[c4_file] or ""), str(row[c4_func] or ""))
                sheet4_data[key] = str(row[c4_compat] or "")

        wb.close()

        # ── Build merged rows ─────────────────────────────────────────────────
        # Use Sheet 1 as the master list; enrich with sheets 3 & 4
        merged = []
        for (file_name, func_name), s1 in sheet1_data.items():
            key = (file_name, func_name)
            merged.append({
                "file_name":    file_name,
                "func_name":    func_name,
                "reuse_status": s1["reuse_status"],
                "which_base":   s1["which_base"],
                "complexity":   sheet3_data.get(key, "—"),
                "compat_pct":   sheet4_data.get(key, "—"),
            })

        # ── HTML helpers ──────────────────────────────────────────────────────
        def _status_class(v):
            if v == "Reuse":           return "reuse"
            if v == "Reuse (Modified)": return "modified"
            if v == "New":             return "new-func"
            return ""

        def _complexity_class(v):
            mapping = {"Low": "cx-low", "Medium": "cx-medium",
                       "High": "cx-high", "Very High": "cx-veryhigh", "Complex": "cx-complex"}
            return mapping.get(v, "")

        def _compat_class(v):
            try:
                pct = float(str(v).replace("%", "").strip())
                if pct >= 75: return "cp-high"
                if pct >= 40: return "cp-mid"
                return "cp-low"
            except Exception:
                return ""

        # ── Summary HTML block ────────────────────────────────────────────────
        summary_html = ""
        if summary_rows:
            rows_html = "".join(
                f'<tr><td class="sum-key">{k}</td><td class="sum-val">{v}</td></tr>'
                for k, v in summary_rows
            )
            summary_html = f"""
<section class="summary-block">
  <h2>Summary</h2>
  <table class="summary-table">
    <tbody>{rows_html}</tbody>
  </table>
</section>"""

        # ── Detail table rows ─────────────────────────────────────────────────
        detail_rows_html = ""
        html_idx = 1
        for row in merged:
            if row["reuse_status"] == "Reuse":
                continue
            sc  = _status_class(row["reuse_status"])
            cxc = _complexity_class(row["complexity"])
            cpc = _compat_class(row["compat_pct"])
            detail_rows_html += (
                f'<tr>'
                f'<td class="center">{html_idx}</td>'
                f'<td>{row["file_name"]}</td>'
                f'<td>{row["func_name"]}</td>'
                f'<td class="center {sc}">{row["reuse_status"]}</td>'
                f'<td>{row["which_base"]}</td>'
                f'<td class="center {cxc}">{row["complexity"]}</td>'
                f'<td class="center {cpc}">{row["compat_pct"]}</td>'
                f'</tr>\n'
            )
            html_idx += 1

        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>FuncAtlas Report</title>
<style>
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: Arial, sans-serif; font-size: 13px;
    background: #f0f4f9; color: #1a2a3a; padding: 28px 32px;
  }}
  h1 {{ color: #1f4e78; font-size: 22px; margin-bottom: 6px; }}
  .subtitle {{ color: #5a7a9a; font-size: 12px; margin-bottom: 24px; }}
  h2 {{ color: #1f4e78; font-size: 15px; margin-bottom: 10px; }}

  /* ── Summary block ── */
  .summary-block {{ margin-bottom: 28px; }}
  .summary-table {{ border-collapse: collapse; min-width: 480px; background: #fff;
                    border-radius: 8px; overflow: hidden;
                    box-shadow: 0 2px 8px rgba(0,0,0,.08); }}
  .summary-table td {{ padding: 7px 14px; border-bottom: 1px solid #dce8f4; }}
  .summary-table tr:last-child td {{ border-bottom: none; }}
  td.sum-key {{ background: #ebf3fb; font-weight: bold; color: #1a3a5c;
                width: 200px; white-space: nowrap; }}
  td.sum-val {{ color: #2a4a6a; }}

  /* ── Legend ── */
  .legend {{ display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 14px; }}
  .legend span {{ display: inline-block; padding: 3px 14px; border-radius: 12px;
                  font-size: 11px; font-weight: bold; }}

  /* ── Detail table ── */
  .detail-table {{ border-collapse: collapse; width: 100%; background: #fff;
                   box-shadow: 0 2px 10px rgba(0,0,0,.10); border-radius: 8px;
                   overflow: hidden; }}
  .detail-table th {{
    background: #1f4e78; color: #fff; padding: 10px 12px;
    text-align: center; font-size: 12px; white-space: nowrap;
  }}
  .detail-table td {{ padding: 7px 12px; border: 1px solid #d0e4f0; vertical-align: middle; }}
  .detail-table tr:nth-child(even) td {{ background: #f5f9fd; }}
  .detail-table tr:hover td {{ background: #ddeeff; }}
  .center {{ text-align: center; }}

  /* Reuse/New */
  .reuse     {{ background: #c6efce !important; color: #276221; font-weight: bold; }}
  .modified  {{ background: #ffeb9c !important; color: #9c5700; font-weight: bold; }}
  .new-func  {{ background: #ffc7ce !important; color: #9c0006; font-weight: bold; }}

  /* Complexity Level */
  .cx-low      {{ background: #d6e4bc !important; color: #3a5a10; font-weight: bold; }}
  .cx-medium   {{ background: #ffe699 !important; color: #7a5500; font-weight: bold; }}
  .cx-high     {{ background: #f4b183 !important; color: #7a3010; font-weight: bold; }}
  .cx-veryhigh {{ background: #ff7070 !important; color: #5a0000; font-weight: bold; }}
  .cx-complex  {{ background: #cc0000 !important; color: #fff;    font-weight: bold; }}

  /* Compatibility % */
  .cp-high {{ background: #c6efce !important; color: #276221; font-weight: bold; }}
  .cp-mid  {{ background: #ffeb9c !important; color: #9c5700; font-weight: bold; }}
  .cp-low  {{ background: #ffc7ce !important; color: #9c0006; font-weight: bold; }}
</style>
</head>
<body>

<h1>FuncAtlas — Function Analysis Report</h1>
<p class="subtitle">Generated from: {os.path.basename(excel_path)}</p>

{summary_html}

<section>
  <h2>Function Detail</h2>
  <div class="legend">
    <span class="modified">Reuse (Modified)</span>
    <span class="new-func">New</span>
    <span style="width:12px;"></span>
    <span class="cx-low">Low</span>
    <span class="cx-medium">Medium</span>
    <span class="cx-high">High</span>
    <span class="cx-veryhigh">Very High</span>
    <span class="cx-complex">Complex</span>
    <span style="width:12px;"></span>
    <span class="cp-high">Compat ≥75%</span>
    <span class="cp-mid">Compat ≥40%</span>
    <span class="cp-low">Compat &lt;40%</span>
  </div>
  <table class="detail-table">
    <thead>
      <tr>
        <th>S.No</th>
        <th>File Name</th>
        <th>Function Name</th>
        <th>Status</th>
        <th>Which Base</th>
        <th>Complexity Level</th>
        <th>Compatibility %</th>
      </tr>
    </thead>
    <tbody>
{detail_rows_html}    </tbody>
  </table>
</section>

</body>
</html>"""

        html_path = save_path if save_path else (os.path.splitext(excel_path)[0] + "_FuncAtlas_Report.html")
        with open(html_path, "w", encoding="utf-8") as fh:
            fh.write(html)
        return html_path

    def _on_report_error_html(self, msg: str):
        self.hide_loading()
        self._report_append_log(f"ERROR: {msg}")
        self.report_phase_label.setText("⚠ Error")
        self.report_generate_html_btn.setEnabled(True)
        self.report_generate_html_btn.setText("Generate HTML Report")
        if hasattr(self, "report_summary_chips"):
            self.report_summary_chips["phase"].set_value("Error")
            self.report_summary_chips["result"].set_value("Failed")
        QMessageBox.critical(self, "HTML Report Error", msg)

    # ── Excel report generation ───────────────────────────────────────────────
    def _on_compare_done(self, out_file: str):
        from ui.widgets import StepStatusWidget
        self._report_output_file = out_file
        self._last_report_excel  = out_file
        # Auto-populate Complexity page with the just-generated report
        if hasattr(self, 'complexity_report_display'):
            self.complexity_report_display.setText(out_file)
        self.report_step_widget.set_state("📊 Comparing & Writing Excel Report",
                                          StepStatusWidget.STATE_DONE, "Excel report written")
        self._set_report_progress(100, "Report ready")
        self.report_phase_label.setText("✅ Complete")
        self.report_generate_btn.setEnabled(True)
        self.report_generate_btn.setText("Generate Report")
        self.report_open_btn.setEnabled(True)
        if hasattr(self, "report_summary_chips"):
            self.report_summary_chips["phase"].set_value("Completed")
            self.report_summary_chips["result"].set_value("Excel ready")
        self._report_append_log(f"✓ Report: {out_file}")
        self.hide_loading()
        QMessageBox.information(self, "Report Ready",
            f"FuncAtlas report generated successfully.\n\nFile:\n{out_file}")

    def _on_report_error(self, msg: str):
        self.hide_loading()
        self._report_append_log(f"ERROR: {msg}")
        self.report_phase_label.setText("⚠ Error")
        self.report_status_label.setText("Error — see log below.")
        self.report_generate_btn.setEnabled(True)
        self.report_generate_btn.setText("Generate Report")
        if hasattr(self, "report_summary_chips"):
            self.report_summary_chips["phase"].set_value("Error")
            self.report_summary_chips["result"].set_value("Failed")
        QMessageBox.critical(self, "Report Error", msg)

    def _report_append_log(self, line: str):
        self.report_log_box.append(line)
        if hasattr(self, "report_log_panel"):
            self.report_log_panel.set_expanded(True)
        sb = self.report_log_box.verticalScrollBar()
        sb.setValue(sb.maximum())

    def _on_report_clear(self):
        self.report_output_field.clear_selection()
        self._set_report_progress(0, "Idle — ready to generate.")
        self.report_log_box.clear()
        self.report_phase_label.setText("Waiting to start")
        self.report_open_btn.setEnabled(False)
        self._report_output_file = ""
        self.report_step_widget.clear_steps()
        self._report_total_steps = 0
        self._update_report_step_chip()
        if hasattr(self, "report_summary_chips"):
            self.report_summary_chips["phase"].set_value("Idle")
            self.report_summary_chips["result"].set_value("Not generated")
        if hasattr(self, "report_log_panel"):
            self.report_log_panel.set_expanded(False)

    def _on_report_open(self):
        """Open the Excel report generated on the Report page directly — no format choice dialog."""
        from PySide6.QtCore import QUrl

        # Report page only produces an Excel report; open it directly.
        excel_path = getattr(self, "_last_report_excel", "") or ""

        if excel_path and os.path.exists(excel_path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(excel_path))
            return

        # Fallback: let the user browse for the file if nothing was generated yet.
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Report", "",
            "Excel Files (*.xlsx *.xls);;All Files (*.*)"
        )
        if path:
            QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    # ── report help menus ─────────────────────────────────────────────────────
    def _show_report_help_menu(self):
        menu = QMenu(self)
        menu.addAction('📖  How to Use',     self._show_report_how_to_use)
        menu.addAction('✅  Prerequisites',  self._show_report_prerequisites)
        menu.addAction('📋  View Log File',  self._show_report_log_dialog)
        menu.exec(self.report_help_btn.mapToGlobal(self.report_help_btn.rect().bottomLeft()))

    def _show_report_how_to_use(self):
        sections = [
            {"title": "Select Output Folder",        "body": "Choose the folder where extracted functions and the final Excel report will be saved."},
            {"title": "Load Reference Data First",    "body": "Go to Input → Reference Bases and submit the target base, reference bases, and function list before running the report."},
            {"title": "Click Generate Report",        "body": "The tool first extracts functions from the target and every reference base."},
            {"title": "Wait for Step Completion",     "body": "Each base runs one by one. The step list shows Waiting, Running, and Complete states clearly."},
            {"title": "Comparison Starts Automatically", "body": "After extraction, FuncAtlas compares target functions against the extracted reference functions and writes the Excel report."},
            {"title": "Open Report",                  "body": "Once processing finishes, use Open Report to open the generated Excel file directly."},
            {"title": "Check Detailed Log",           "body": "Use View Log File from the Help menu if you need to inspect progress details or troubleshoot failures."},
        ]
        HelpOverlayDialog(self, "Help — How to Use", sections,
            footer_text="⚠ Prerequisites: Windows OS · Microsoft Excel installed · Input data must be submitted before generating the report.",
            tip_text="💡 Tip: Use a local drive path for output. Network folders and locked Excel files slow this workflow down or break it."
        ).exec()

    def _show_report_prerequisites(self):
        sections = [
            {"title": "Reference data must be submitted", "body": "The report page depends on the target base and reference bases already loaded from the Input → Reference Bases page."},
            {"title": "Output folder must be writable",   "body": "Pick a folder where Excel files and extracted text files can be created. Read-only folders will fail."},
            {"title": "Close locked Excel files",         "body": "If the report file is already open in Excel, writing can fail or force a renamed output."},
            {"title": "Large inputs take real time",      "body": "This build is threaded, so the UI stays responsive, but extraction and comparison time still depends on file count and file size."},
        ]
        HelpOverlayDialog(self, "Help — Prerequisites", sections,
            footer_text="⚠ Do not run the report before loading target and reference sources. That is the main reason users get empty or failed output.",
            tip_text="💡 Tip: Keep output on SSD/local disk and avoid opening the same report while generation is still running."
        ).exec()

    def _show_report_log_dialog(self):
        dlg = QDialog(self)
        dlg.setWindowTitle('Report Log')
        dlg.setModal(True)
        dlg.resize(860, 520)
        outer = QVBoxLayout(dlg)
        outer.setContentsMargins(0, 0, 0, 0)
        shell = QFrame()
        shell.setObjectName('aboutDialogShell')
        shell_layout = QVBoxLayout(shell)
        shell_layout.setContentsMargins(24, 24, 24, 24)
        head = QHBoxLayout()
        title_lbl = QLabel('Report Log File')
        title_lbl.setObjectName('aboutDialogTitle')
        close_btn = QPushButton('×')
        close_btn.setObjectName('aboutDialogCloseButton')
        close_btn.setFixedSize(42, 42)
        close_btn.clicked.connect(dlg.accept)
        head.addWidget(title_lbl); head.addStretch(); head.addWidget(close_btn)
        shell_layout.addLayout(head)
        log_box = QTextEdit()
        log_box.setReadOnly(True)
        log_box.setPlainText(self.report_log_box.toPlainText() or 'No log entries yet.')
        shell_layout.addWidget(log_box, 1)
        outer.addWidget(shell)
        dlg.exec()

    # ── About dialog ──────────────────────────────────────────────────────────
    def show_about_dialog(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("About FuncAtlas")
        dlg.setModal(True)
        dlg.setFixedSize(790, 620)
        outer = QVBoxLayout(dlg)
        outer.setContentsMargins(0, 0, 0, 0)
        shell = QFrame(); shell.setObjectName("aboutDialogShell")
        shell_layout = QVBoxLayout(shell)
        shell_layout.setContentsMargins(34, 26, 34, 28)
        shell_layout.setSpacing(20)

        head = QHBoxLayout()
        title_lbl = QLabel("About — FuncAtlas"); title_lbl.setObjectName("aboutDialogTitle")
        close_btn = QPushButton("×"); close_btn.setObjectName("aboutDialogCloseButton")
        close_btn.setFixedSize(42, 42); close_btn.clicked.connect(dlg.accept)
        head.addWidget(title_lbl); head.addStretch(); head.addWidget(close_btn)
        shell_layout.addLayout(head)

        hero = QFrame(); hero.setObjectName("aboutDialogHero")
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(22, 20, 22, 18); hero_layout.setSpacing(6)
        h1 = QLabel("FuncAtlas — Automated Function Analysis & Reporting Engine")
        h1.setObjectName("aboutDialogHeroTitle")
        h2 = QLabel("Version 1.0.0 · Reuse Analysis Workspace")
        h2.setObjectName("aboutDialogHeroSubtitle")
        hero_layout.addWidget(h1); hero_layout.addWidget(h2)
        shell_layout.addWidget(hero)

        bullets_frame = QFrame(); bullets_frame.setObjectName("aboutDialogBody")
        bullets_layout = QVBoxLayout(bullets_frame)
        bullets_layout.setContentsMargins(2, 8, 2, 2); bullets_layout.setSpacing(18)
        items = [
            ("🔎", "Compare target and reference source files from selected folders"),
            ("📄", "Read function lists from TXT and XLSX files without changing the source data"),
            ("📊", "Generate Excel summary reports for matched functions and reuse status"),
            ("🧭", "Preview detected files and function bodies inside the explorer workspace"),
            ("🔐", "Background processing with progress tracking and safer output handling"),
            ("🧹", "Avoid locked-file overwrite failures by writing a new report name when needed"),
            ("📝", "Detailed runtime logging in the report panel for easier troubleshooting"),
        ]
        for icon_text, line in items:
            row = QHBoxLayout(); row.setSpacing(14)
            icon = QLabel(icon_text); icon.setObjectName("aboutDialogBulletIcon")
            icon.setFixedWidth(28); icon.setAlignment(Qt.AlignTop | Qt.AlignHCenter)
            txt = QLabel(line); txt.setObjectName("aboutDialogBulletText"); txt.setWordWrap(True)
            row.addWidget(icon); row.addWidget(txt, 1)
            bullets_layout.addLayout(row)
        shell_layout.addWidget(bullets_frame)
        outer.addWidget(shell)

        dlg.setStyleSheet(f"""
            QDialog {{ background: transparent; }}
            QFrame#aboutDialogShell {{ background: {self.theme['bg_header']}; border: 1px solid {self.theme['border_strong']}; border-radius: 26px; }}
            QLabel#aboutDialogTitle {{ color: {self.accent_color.name()}; font-size: 17px; font-weight: 900; background: transparent; }}
            QFrame#aboutDialogHero {{ background: {self.theme['bg_card']}; border: 1px solid {self.theme['border']}; border-radius: 18px; }}
            QLabel#aboutDialogHeroTitle {{ color: {self.theme['text_primary']}; font-size: 15px; font-weight: 900; background: transparent; }}
            QLabel#aboutDialogHeroSubtitle {{ color: {self.theme['text_muted']}; font-size: 11px; background: transparent; }}
            QLabel#aboutDialogBulletIcon {{ color: {self.accent_color.name()}; font-size: 17px; background: transparent; }}
            QLabel#aboutDialogBulletText {{ color: {self.theme['text_secondary']}; font-size: 12px; font-weight: 600; background: transparent; }}
            QPushButton#aboutDialogCloseButton {{ background: transparent; color: {self.theme['text_muted']}; border: none; font-size: 26px; font-weight: 500; }}
            QPushButton#aboutDialogCloseButton:hover {{ color: {self.theme['text_primary']}; }}
        """)
        dlg.exec()

    # ── resize handler ────────────────────────────────────────────────────────
    def resizeEvent(self, event):
        super().resizeEvent(event)
        try:
            self.refresh_home_hero_image()
        except Exception:
            pass

    def closeEvent(self, event):
        """Stop and wait for every background thread before closing."""
        # 1. Threads tracked in _active_threads (con + report workers)
        for thread in list(self._active_threads):
            if thread and thread.isRunning():
                thread.quit()
                thread.wait(3000)

        # 2. AutoDetectColumnField threads — field keeps a _threads LIST
        for field_attr in ("con_func_col_field", "con_base_col_field"):
            field = getattr(self, field_attr, None)
            if field is None:
                continue
            for thread in list(getattr(field, "_threads", [])):
                if thread and thread.isRunning():
                    thread.quit()
                    thread.wait(2000)

        # 3. Report threads may have been created but not yet added (race guard)
        for t_attr in ("_report_ext_thread", "_report_cmp_thread", "con_thread"):
            t = getattr(self, t_attr, None)
            if t and t.isRunning():
                t.quit()
                t.wait(3000)

        super().closeEvent(event)