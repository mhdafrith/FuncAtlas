# pages/complexity_page.py — Complexity & Compatibility page (UPDATED)

import os
import json

from PySide6.QtCore import Qt, QThread
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFrame, QLabel,
    QPushButton, QLineEdit, QDialog, QScrollArea,
    QFileDialog, QSpinBox, QAbstractSpinBox, QMessageBox,
    QTextEdit, QProgressBar, QSizePolicy
)
from ui.widgets import SectionTitle, IconTextButton
from services.complexity_worker import ComplexityAppendWorker

# ── JSON persistence ──────────────────────────────────────────────────────────
_SETTINGS_DIR = os.path.join(os.path.expanduser("~"), ".funcatlas")
_HANDLED_JSON = os.path.join(_SETTINGS_DIR, "handled_scenarios.json")

ALL_SCENARIOS = [
    "If...Else",
    "If...Else if...Else",
    "Nested If",
    "Switch",
    "For",
    "While",
    "Do...While",
    "Return",
    "Function Call",
    "Pointers",
    "Struct",
    "Assign",
]


def _load_handled_scenarios():
    os.makedirs(_SETTINGS_DIR, exist_ok=True)
    if os.path.isfile(_HANDLED_JSON):
        try:
            with open(_HANDLED_JSON, "r", encoding="utf-8") as f:
                data = json.load(f)
                return set(data.get("handled_scenarios", []))
        except Exception:
            pass
    return set()


def _save_handled_scenarios(handled):
    os.makedirs(_SETTINGS_DIR, exist_ok=True)
    with open(_HANDLED_JSON, "w", encoding="utf-8") as f:
        json.dump({"handled_scenarios": sorted(handled)}, f, indent=2)


# ── Combined Complexity + Compatibility Settings Dialog ───────────────────────
class ComplexitySettingsDialog(QDialog):
    DEFAULT_WEIGHTS = [
        ("If...Else",           1),
        ("If...Else if...Else", 2),
        ("Nested If",           4),
        ("Switch",              2),
        ("For",                 2),
        ("While",               2),
        ("Do...While",          2),
        ("Return",              1),
        ("Function Call",       4),
        ("Pointers",            7),
        ("Struct",              3),
        ("Assign",              1),
    ]

    DEFAULT_BANDS = [
        ("Low",       1,   5),
        ("Medium",    6,  12),
        ("High",     13,  25),
        ("Very High", 26,  40),
        ("Complex",  41, 999),
    ]

    def __init__(self, parent, weights=None, bands=None, handled_scenarios=None):
        super().__init__(parent)
        self.setWindowTitle("Complexity & Compatibility Settings")
        self.setModal(True)
        self.setMinimumSize(1200, 820)
        self.resize(1440, 920)
        self.setMinimumHeight(820)

        self._weights = weights if weights is not None else [list(r) for r in self.DEFAULT_WEIGHTS]
        self._bands   = bands   if bands   is not None else [list(r) for r in self.DEFAULT_BANDS]
        self._handled = set(handled_scenarios) if handled_scenarios is not None else _load_handled_scenarios()

        self._show_complexity   = (weights is not None or bands is not None)
        self._show_compatibility = (handled_scenarios is not None)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)

        shell = QFrame()
        shell.setObjectName("complexityDialogShell")
        shell_layout = QVBoxLayout(shell)
        shell_layout.setContentsMargins(32, 28, 32, 28)
        shell_layout.setSpacing(20)

        # ── Header row ────────────────────────────────────────────────────────
        head_row = QHBoxLayout()
        title_lbl = QLabel("Complexity & Compatibility Settings")
        title_lbl.setObjectName("complexityDialogTitle")

        # FIX 4: show Save & Close only in complexity mode; compat uses its own Save Changes btn
        # FIX 5: show Default button (reset to defaults) only in complexity mode
        head_row.addWidget(title_lbl)
        head_row.addStretch()

        if self._show_complexity:
            # Default button — resets weights and bands to built-in defaults (FIX 5)
            default_btn = QPushButton("  \u21BA  Default")
            default_btn.setObjectName("complexityDefaultBtn")
            default_btn.setFixedSize(130, 44)
            default_btn.setToolTip("Reset all weights and bands to their default values")
            default_btn.clicked.connect(self._on_reset_defaults)
            head_row.addWidget(default_btn)
            head_row.addSpacing(8)

            # Save & Close only for complexity settings (FIX 4: removed from compat dialog)
            save_btn = QPushButton("  Save \u0026 Close")
            save_btn.setObjectName("complexitySaveBtn")
            save_btn.setFixedSize(150, 44)
            save_btn.clicked.connect(self._on_save)
            head_row.addWidget(save_btn)
            head_row.addSpacing(8)

        # Close button always present
        close_btn = QPushButton("  Close")
        close_btn.setObjectName("complexityCloseBtn")
        close_btn.setFixedSize(120, 44)
        close_btn.clicked.connect(self.reject)
        head_row.addWidget(close_btn)
        shell_layout.addLayout(head_row)

        # ── Body ─────────────────────────────────────────────────────────────
        body_row = QHBoxLayout()
        body_row.setSpacing(24)

        if self._show_complexity:
            # Col 1: Weightage
            left_frame = QFrame()
            ll = QVBoxLayout(left_frame)
            ll.setContentsMargins(0, 0, 0, 0)
            ll.setSpacing(8)
            ll.addWidget(self._section_hdr(
                "Weightage",
                "One editable weight value for each C-language scenario"
            ))

            w_scroll = QScrollArea()
            w_scroll.setWidgetResizable(True)
            w_scroll.setFrameShape(QFrame.NoFrame)
            w_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            w_scroll.setMinimumHeight(500)
            w_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

            w_inner = QFrame()
            w_inner.setObjectName("complexityTableInner")
            wil = QVBoxLayout(w_inner)
            wil.setContentsMargins(0, 0, 0, 0)
            wil.setSpacing(0)
            wil.addWidget(self._tbl_hdr(["Scenario", "Weightage"]))

            self._weight_spins = []
            for i, (sc, val) in enumerate(self._weights):
                rw, sp = self._stepper_row(sc, val, i)
                wil.addWidget(rw)
                self._weight_spins.append((sc, sp))

            wil.addStretch()
            w_scroll.setWidget(w_inner)
            ll.addWidget(w_scroll, 1)
            body_row.addWidget(left_frame, 1)

            # Col 2: Complexity Bands
            mid_frame = QFrame()
            ml = QVBoxLayout(mid_frame)
            ml.setContentsMargins(0, 0, 0, 0)
            ml.setSpacing(8)
            ml.addWidget(self._section_hdr(
                "Complexity Value",
                "Five editable bands with start and end values"
            ))

            band_frame = QFrame()
            band_frame.setObjectName("complexityTableInner")
            bfl = QVBoxLayout(band_frame)
            bfl.setContentsMargins(0, 0, 0, 0)
            bfl.setSpacing(0)
            bfl.addWidget(self._tbl_hdr(["Level", "Start", "", "End", ""]))

            self._band_spins = []
            for i, (lv, s, e) in enumerate(self._bands):
                rw, ss, es = self._band_row(lv, s, e, i)
                bfl.addWidget(rw)
                self._band_spins.append((lv, ss, es))

            bfl.addStretch()
            ml.addWidget(band_frame)

            note = QLabel(
                "Each level has two editable boxes: start value and end value.\n"
                "Example: Low=1 to 5, Medium=6 to 12, High=13 to 25, "
                "Very High=26 to 40, Complex=41 and above."
            )
            note.setObjectName("complexityNote")
            note.setWordWrap(True)
            ml.addWidget(note)
            ml.addStretch()
            body_row.addWidget(mid_frame, 1)

        # ── Compatibility panel — redesigned as a single table ────────────────
        if self._show_compatibility:
            compat_frame = QFrame()
            compat_frame.setObjectName("compatFrame")
            cl = QVBoxLayout(compat_frame)
            cl.setContentsMargins(0, 0, 0, 0)
            cl.setSpacing(10)

            # Section header + subtitle
            cl.addWidget(self._section_hdr(
                "Compatibility Scenarios",
                "Click to mark which C scenarios your system handles"
            ))

            # ── Handled Scenarios Summary Box ────────────────────────────────
            self._summary_frame = QFrame()
            self._summary_frame.setObjectName("handledSummaryBox")
            summary_layout = QVBoxLayout(self._summary_frame)
            summary_layout.setContentsMargins(14, 10, 14, 10)
            summary_layout.setSpacing(4)

            summary_title_row = QHBoxLayout()
            summary_title_lbl = QLabel("✓  Already Handled Scenarios")
            summary_title_lbl.setObjectName("handledSummaryTitle")
            summary_title_row.addWidget(summary_title_lbl)
            summary_title_row.addStretch()
            self._summary_count_lbl = QLabel("")
            self._summary_count_lbl.setObjectName("handledSummaryCount")
            summary_title_row.addWidget(self._summary_count_lbl)
            summary_layout.addLayout(summary_title_row)

            self._summary_text_lbl = QLabel("")
            self._summary_text_lbl.setObjectName("handledSummaryText")
            self._summary_text_lbl.setWordWrap(True)
            summary_layout.addWidget(self._summary_text_lbl)

            cl.addWidget(self._summary_frame)

            # ── Modify / Save / Cancel button row ────────────────────────────
            self._edit_mode = False

            btn_row = QHBoxLayout()
            btn_row.addStretch()

            self.modify_btn = QPushButton("  ✎  Modify")
            self.modify_btn.setObjectName("compatModifyBtn")
            self.modify_btn.setFixedSize(130, 36)
            self.modify_btn.clicked.connect(self._enable_edit_mode)

            self.save_edit_btn = QPushButton("✔  Save Changes")
            self.save_edit_btn.setObjectName("compatSaveEditBtn")
            self.save_edit_btn.setFixedSize(150, 36)
            self.save_edit_btn.clicked.connect(self._save_changes)

            self.cancel_edit_btn = QPushButton("✖  Cancel")
            self.cancel_edit_btn.setObjectName("compatCancelEditBtn")
            self.cancel_edit_btn.setFixedSize(110, 36)
            self.cancel_edit_btn.clicked.connect(self._cancel_edit)

            btn_row.addWidget(self.modify_btn)
            btn_row.addSpacing(8)
            btn_row.addWidget(self.save_edit_btn)
            btn_row.addSpacing(8)
            btn_row.addWidget(self.cancel_edit_btn)
            cl.addLayout(btn_row)

            # ── Single-table scroll area ──────────────────────────────────────
            tbl_scroll = QScrollArea()
            tbl_scroll.setWidgetResizable(True)
            tbl_scroll.setFrameShape(QFrame.NoFrame)
            tbl_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

            tbl_inner = QFrame()
            tbl_inner.setObjectName("complexityTableInner")
            til = QVBoxLayout(tbl_inner)
            til.setContentsMargins(0, 0, 0, 0)
            til.setSpacing(0)

            # Header row
            til.addWidget(self._tbl_hdr(["C Scenario", "Status"]))

            # One row per scenario: label + status badge button
            self._status_btns = {}
            for i, sc in enumerate(ALL_SCENARIOS):
                row, badge = self._scenario_status_row(sc, sc in self._handled, i)
                til.addWidget(row)
                self._status_btns[sc] = badge

            til.addStretch()
            tbl_scroll.setWidget(tbl_inner)
            cl.addWidget(tbl_scroll, 1)

            # Start in view-only mode
            self._set_edit_enabled(False)
            self._refresh_summary()  # populate summary on open

            body_row.addWidget(compat_frame, 1)

        if not self._show_complexity and not self._show_compatibility:
            fallback = QLabel("No settings section selected.")
            fallback.setObjectName("complexityNote")
            body_row.addWidget(fallback)

        shell_layout.addLayout(body_row, 1)
        outer.addWidget(shell)
        self._apply_styles()

    # ── helpers ───────────────────────────────────────────────────────────────
    def _section_hdr(self, title, sub):
        w = QWidget()
        l = QVBoxLayout(w)
        l.setContentsMargins(0, 0, 0, 0)
        l.setSpacing(2)
        t = QLabel(title)
        t.setObjectName("complexitySectionTitle")
        s = QLabel(sub)
        s.setObjectName("complexitySectionSub")
        l.addWidget(t)
        l.addWidget(s)
        return w

    def _tbl_hdr(self, cols):
        h = QFrame()
        h.setObjectName("complexityTableHeader")
        hl = QHBoxLayout(h)
        hl.setContentsMargins(16, 10, 16, 10)
        for i, txt in enumerate(cols):
            lb = QLabel(txt)
            lb.setObjectName("complexityColHeader")
            lb.setAlignment(Qt.AlignCenter)
            hl.addWidget(lb, 1 if i == 0 else 0)
        return h

    def _spin(self, val, w=80):
        sp = QSpinBox()
        sp.setObjectName("complexitySpinBox")
        sp.setRange(0, 9999)
        sp.setValue(val)
        sp.setFixedWidth(w)
        sp.setAlignment(Qt.AlignCenter)
        sp.setButtonSymbols(QAbstractSpinBox.NoButtons)
        return sp

    def _sbtn(self, sym):
        b = QPushButton(sym)
        if sym == "+":
            b.setObjectName("complexityStepPlusBtn")
            b.setFixedSize(36, 36)
        else:
            b.setObjectName("complexityStepBtn")
            b.setFixedSize(32, 32)
        return b

    def _stepper_row(self, label, val, idx):
        row = QFrame()
        row.setObjectName("complexityRowEven" if idx % 2 == 0 else "complexityRowOdd")
        rl = QHBoxLayout(row)
        rl.setContentsMargins(16, 8, 16, 8)
        rl.setSpacing(8)
        lb = QLabel(label)
        lb.setObjectName("complexityScenarioLabel")
        m = self._sbtn("\u2212")
        sp = self._spin(val)
        p = self._sbtn("+")
        m.clicked.connect(lambda _, s=sp: s.setValue(max(0, s.value() - 1)))
        p.clicked.connect(lambda _, s=sp: s.setValue(s.value() + 1))
        rl.addWidget(lb, 1)
        rl.addWidget(m)
        rl.addWidget(sp)
        rl.addWidget(p)
        return row, sp

    def _band_row(self, level, start, end, idx):
        row = QFrame()
        row.setObjectName("complexityRowEven" if idx % 2 == 0 else "complexityRowOdd")
        rl = QHBoxLayout(row)
        rl.setContentsMargins(16, 8, 16, 8)
        rl.setSpacing(8)
        lb = QLabel(level)
        lb.setObjectName("complexityScenarioLabel")
        lb.setMinimumWidth(80)
        ss = self._spin(start, 65)
        es = self._spin(end, 65)
        sm = self._sbtn("\u2212")
        sp2 = self._sbtn("+")
        em = self._sbtn("\u2212")
        ep = self._sbtn("+")
        sm.clicked.connect(lambda _, s=ss: s.setValue(max(0, s.value() - 1)))
        sp2.clicked.connect(lambda _, s=ss: s.setValue(s.value() + 1))
        em.clicked.connect(lambda _, s=es: s.setValue(max(0, es.value() - 1)))
        ep.clicked.connect(lambda _, s=es: s.setValue(es.value() + 1))
        rl.addWidget(lb, 1)
        rl.addWidget(sm)
        rl.addWidget(ss)
        rl.addWidget(sp2)
        rl.addSpacing(4)
        rl.addWidget(em)
        rl.addWidget(es)
        rl.addWidget(ep)
        return row, ss, es

    def _scenario_status_row(self, scenario: str, is_handled: bool, idx: int):
        """Single-row table entry: scenario name + toggle badge."""
        row = QFrame()
        row.setObjectName("complexityRowEven" if idx % 2 == 0 else "complexityRowOdd")
        rl = QHBoxLayout(row)
        rl.setContentsMargins(16, 8, 16, 8)
        rl.setSpacing(8)

        lb = QLabel(scenario)
        lb.setObjectName("complexityScenarioLabel")

        badge = QPushButton()
        badge.setObjectName("statusBadgeHandled" if is_handled else "statusBadgeNotHandled")
        badge.setFixedSize(140, 30)
        badge.setProperty("handled", is_handled)
        self._update_badge_text(badge, is_handled)
        badge.clicked.connect(lambda _, b=badge, s=scenario: self._toggle_status(b, s))
        badge.setEnabled(False)  # disabled until edit mode

        rl.addWidget(lb, 1)
        rl.addWidget(badge)
        return row, badge

    def _update_badge_text(self, badge: QPushButton, is_handled: bool):
        if is_handled:
            badge.setText("✓  Handled")
        else:
            badge.setText("○  Not Handled")

    def _toggle_status(self, badge: QPushButton, scenario: str):
        """Toggle handled/not-handled when in edit mode."""
        currently_handled = badge.property("handled")
        new_state = not currently_handled
        badge.setProperty("handled", new_state)
        self._update_badge_text(badge, new_state)
        # Re-apply object name to trigger stylesheet update
        badge.setObjectName("statusBadgeHandled" if new_state else "statusBadgeNotHandled")
        badge.setStyle(badge.style())  # force style refresh
        self._refresh_summary()  # keep summary in sync

    def _refresh_summary(self):
        """Update the handled-scenarios summary box."""
        if not hasattr(self, "_status_btns") or not hasattr(self, "_summary_text_lbl"):
            return
        handled_list = [sc for sc, badge in self._status_btns.items() if badge.property("handled")]
        count = len(handled_list)
        total = len(self._status_btns)
        self._summary_count_lbl.setText(f"{count} / {total}")
        if handled_list:
            self._summary_text_lbl.setText("  •  ".join(handled_list))
            self._summary_text_lbl.setVisible(True)
        else:
            self._summary_text_lbl.setText("None marked as handled yet.")
            self._summary_text_lbl.setVisible(True)

    # ── Compatibility edit actions ─────────────────────────────────────────────
    def _enable_edit_mode(self):
        """Show warning dialog before entering edit mode."""
        msg = QMessageBox(self)
        msg.setWindowTitle("Modify Scenarios")
        msg.setIcon(QMessageBox.Warning)
        msg.setText("<b>You are about to modify handled scenarios.</b>")
        msg.setInformativeText(
            "This will change which C scenarios are marked as handled by your system.\n\n"
            "Click each scenario's status badge to toggle between Handled / Not Handled.\n"
            "Use Save Changesto apply or Cancel to discard."
        )
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        msg.setDefaultButton(QMessageBox.Ok)
        msg.button(QMessageBox.Ok).setText("  Proceed  ")
        msg.button(QMessageBox.Cancel).setText("  Cancel  ")

        if msg.exec() == QMessageBox.Ok:
            # Snapshot current state for cancel
            self._handled_snapshot = {
                sc for sc, badge in self._status_btns.items()
                if badge.property("handled")
            }
            self._edit_mode = True
            self._set_edit_enabled(True)

    def _set_edit_enabled(self, enabled: bool):
        """Toggle between view-only and edit mode."""
        for badge in self._status_btns.values():
            badge.setEnabled(enabled)

        self.save_edit_btn.setVisible(enabled)
        self.cancel_edit_btn.setVisible(enabled)
        self.modify_btn.setVisible(not enabled)

        # Visual cue: highlight table border in edit mode
        if enabled:
            for badge in self._status_btns.values():
                badge.setStyle(badge.style())

    def _save_changes(self):
        """Persist changes to JSON and exit edit mode."""
        self._handled = {
            sc for sc, badge in self._status_btns.items()
            if badge.property("handled")
        }
        _save_handled_scenarios(self._handled)
        QMessageBox.information(self, "Saved", "Handled scenarios saved successfully.")
        self._edit_mode = False
        self._set_edit_enabled(False)

    def _cancel_edit(self):
        """Discard unsaved edits and restore the snapshot."""
        for sc, badge in self._status_btns.items():
            was_handled = sc in self._handled_snapshot
            badge.setProperty("handled", was_handled)
            self._update_badge_text(badge, was_handled)
            badge.setObjectName("statusBadgeHandled" if was_handled else "statusBadgeNotHandled")
            badge.setStyle(badge.style())
        self._edit_mode = False
        self._set_edit_enabled(False)
        self._refresh_summary()

    def _on_save(self):
        """Save & Close — used only for complexity settings (Fix 4)."""
        self.accept()

    def _on_reset_defaults(self):
        """FIX 5: Reset all weight spinboxes and band spinboxes to built-in defaults."""
        reply = QMessageBox.question(
            self, "Reset to Defaults",
            "Reset all weights and complexity bands to their default values?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        # Reset weight spinboxes
        default_w = {sc: val for sc, val in self.DEFAULT_WEIGHTS}
        for sc, sp in self._weight_spins:
            if sc in default_w:
                sp.setValue(default_w[sc])

        # Reset band spinboxes
        default_b = {lv: (s, e) for lv, s, e in self.DEFAULT_BANDS}
        for lv, ss, es in self._band_spins:
            if lv in default_b:
                ss.setValue(default_b[lv][0])
                es.setValue(default_b[lv][1])

    def _apply_styles(self):
        self.setStyleSheet("""
            QDialog, QFrame#complexityDialogShell { background: #0D1B2A; }
            QWidget { background: #0D1B2A; color: #C8D8E8; }
            QLabel#complexityDialogTitle  { color: white; font-size: 21px; font-weight: 900; }
            QPushButton#complexitySaveBtn {
                background: #27AE60; color: white; border: none;
                border-radius: 14px; font-size: 14px; font-weight: 800;
            }
            QPushButton#complexitySaveBtn:hover { background: #2ECC71; }
            QPushButton#complexityDefaultBtn {
                background: #1A3050; color: #8AAAC8;
                border: 2px solid #1D3347;
                border-radius: 14px; font-size: 14px; font-weight: 800;
            }
            QPushButton#complexityDefaultBtn:hover {
                background: #E67E22; color: white; border-color: #F39C12;
            }
            QPushButton#complexityCloseBtn {
                background: #C0392B; color: white; border: none;
                border-radius: 14px; font-size: 14px; font-weight: 800;
            }
            QPushButton#complexityCloseBtn:hover { background: #E74C3C; }
            QLabel#complexitySectionTitle { color: white; font-size: 15px; font-weight: 900; }
            QLabel#complexitySectionSub   { color: #8899AA; font-size: 11px; }
            QFrame#complexityTableInner   {
                background: #101E2E; border: 1px solid #1D3347; border-radius: 10px;
            }
            QFrame#complexityTableHeader  {
                background: #162840; border-bottom: 1px solid #1D3347;
                border-top-left-radius: 10px; border-top-right-radius: 10px;
            }
            QLabel#complexityColHeader    { color: white; font-weight: 800; font-size: 12px; }
            QFrame#complexityRowEven      { background: #101E2E; }
            QFrame#complexityRowOdd       { background: #0D1B2A; }
            QLabel#complexityScenarioLabel { color: #C8D8E8; font-size: 12px; }
            QPushButton#complexityStepBtn {
                background: #1A3050; color: #1E90FF; border: 1px solid #1D3347;
                border-radius: 8px; font-size: 15px; font-weight: 900;
            }
            QPushButton#complexityStepBtn:hover { background: #1E90FF; color: white; }
            QPushButton#complexityStepPlusBtn {
                background: #2ECC71; color: white;
                border: 2px solid #27AE60;
                border-radius: 8px; font-size: 18px; font-weight: 900;
            }
            QPushButton#complexityStepPlusBtn:hover { background: #58D68D; color: white; border-color: #2ECC71; }
            QSpinBox#complexitySpinBox {
                background: #101E2E; color: white; border: 1px solid #1D3347;
                border-radius: 8px; font-size: 13px; font-weight: 700; padding: 4px;
            }
            QLabel#complexityNote { color: #7A8FA0; font-size: 11px; }
            QScrollBar:vertical { background: #0D1B2A; width: 8px; border-radius: 4px; }
            QScrollBar::handle:vertical { background: #1D3347; border-radius: 4px; min-height: 20px; }

            /* ── Handled Scenarios Summary Box ── */
            QFrame#handledSummaryBox {
                background: #0A1E10; border: 1px solid #27AE60;
                border-radius: 10px;
            }
            QLabel#handledSummaryTitle {
                color: #2ECC71; font-size: 12px; font-weight: 900;
            }
            QLabel#handledSummaryCount {
                color: #27AE60; font-size: 12px; font-weight: 800;
                background: #1A5C35; border-radius: 6px; padding: 2px 8px;
            }
            QLabel#handledSummaryText {
                color: #8ECFA8; font-size: 11px; font-weight: 600;
                line-height: 1.5;
            }

            /* ── Compatibility status badges ── */
            QPushButton#statusBadgeHandled {
                background: #1A5C35; color: #2ECC71;
                border: 1px solid #27AE60; border-radius: 8px;
                font-size: 12px; font-weight: 800;
            }
            QPushButton#statusBadgeHandled:hover:enabled {
                background: #27AE60; color: white; cursor: pointer;
            }
            QPushButton#statusBadgeHandled:disabled {
                background: #1A5C35; color: #2ECC71; border-color: #27AE60;
            }
            QPushButton#statusBadgeNotHandled {
                background: #1E2A38; color: #7A8FA0;
                border: 1px solid #2A3F55; border-radius: 8px;
                font-size: 12px; font-weight: 800;
            }
            QPushButton#statusBadgeNotHandled:hover:enabled {
                background: #1A3050; color: #C8D8E8; border-color: #1E90FF;
            }
            QPushButton#statusBadgeNotHandled:disabled {
                background: #1E2A38; color: #556677; border-color: #1D3347;
            }

            /* ── Compat action buttons ── */
            QPushButton#compatModifyBtn {
                background: #162840; color: #8AAAC8;
                border: 2px solid #1D3347; border-radius: 10px;
                font-size: 12px; font-weight: 800;
            }
            QPushButton#compatModifyBtn:hover { background: #1E4070; color: white; border-color: #1E90FF; }
            QPushButton#compatSaveEditBtn {
                background: #1A5C35; color: #2ECC71;
                border: 1px solid #27AE60; border-radius: 10px;
                font-size: 12px; font-weight: 800;
            }
            QPushButton#compatSaveEditBtn:hover { background: #27AE60; color: white; }
            QPushButton#compatCancelEditBtn {
                background: #2C1A1A; color: #E74C3C;
                border: 1px solid #C0392B; border-radius: 10px;
                font-size: 12px; font-weight: 800;
            }
            QPushButton#compatCancelEditBtn:hover { background: #C0392B; color: white; }
        """)

    def get_weights(self):
        if not hasattr(self, "_weight_spins"):
            return []
        return [(s, sp.value()) for s, sp in self._weight_spins]

    def get_bands(self):
        if not hasattr(self, "_band_spins"):
            return []
        return [(lv, s.value(), e.value()) for lv, s, e in self._band_spins]

    def get_handled_scenarios(self):
        return set(self._handled)


# ── Page builder ──────────────────────────────────────────────────────────────
def create_complexity_page(win):
    page = QWidget()
    layout = QVBoxLayout(page)
    layout.setContentsMargins(4, 4, 4, 4)
    layout.setSpacing(14)

    # ── Top Row: title on left, settings buttons on top-right ─────────────────
    header_row = QHBoxLayout()
    title = QLabel("Complexity & Compatibility")
    title.setObjectName("sectionTitle")
    header_row.addWidget(title)
    header_row.addStretch()

    win.complexity_settings_btn = IconTextButton(
        "Complexity Settings", win.icons_white.icon("settings", 15)
    )
    win.complexity_settings_btn.setObjectName("pickerButton")
    win.complexity_settings_btn.setFixedSize(190, 36)

    win.compatibility_settings_btn = IconTextButton(
        "Compatibility Settings", win.icons_white.icon("settings", 15)
    )
    win.compatibility_settings_btn.setObjectName("pickerButton")
    win.compatibility_settings_btn.setFixedSize(205, 36)

    header_row.addWidget(win.complexity_settings_btn)
    header_row.addWidget(win.compatibility_settings_btn)
    layout.addLayout(header_row)

    # ── Card ─────────────────────────────────────────────
    report_card = QFrame()
    report_card.setObjectName("pageCard")
    rc_layout = QVBoxLayout(report_card)
    rc_layout.setContentsMargins(22, 18, 22, 18)
    rc_layout.setSpacing(12)

    rc_layout.addWidget(SectionTitle("Report", ""))

    # ── Input + Browse/Clear in the same row ─────────────────────────────────
    file_row = QHBoxLayout()
    file_row.setSpacing(10)

    win.complexity_report_display = QLineEdit()
    win.complexity_report_display.setReadOnly(True)
    win.complexity_report_display.setPlaceholderText("No report selected")
    win.complexity_report_display.setFixedHeight(52)
    win.complexity_report_display.setObjectName("complexityReportDisplay")

    win.complexity_browse_btn = IconTextButton(
        "Browse Report", win.icons_white.icon("excel", 15)
    )
    win.complexity_browse_btn.setObjectName("pickerButton")
    win.complexity_browse_btn.setFixedSize(160, 36)

    win.complexity_clear_btn = QPushButton("✕ Clear")
    win.complexity_clear_btn.setObjectName("clearButton")
    win.complexity_clear_btn.setFixedSize(120, 36)

    file_row.addWidget(win.complexity_report_display, 1)
    file_row.addWidget(win.complexity_browse_btn)
    file_row.addWidget(win.complexity_clear_btn)
    rc_layout.addLayout(file_row)

    # ── Generate + Open Report Buttons centered below the card content ─────────
    gen_row = QHBoxLayout()
    gen_row.addStretch()

    win.complexity_generate_btn = IconTextButton(
        "Generate Report", win.icons_white.icon("submit", 15)
    )
    win.complexity_generate_btn.setObjectName("smallPrimaryButton")
    win.complexity_generate_btn.setFixedSize(180, 40)

    win.complexity_generate_html_btn = IconTextButton(
        "Generate HTML Report", win.icons_white.icon("submit", 15)
    )
    win.complexity_generate_html_btn.setObjectName("smallPrimaryButton")
    win.complexity_generate_html_btn.setFixedSize(200, 40)

    win.complexity_open_report_btn = IconTextButton(
        "Open Report", win.icons_white.icon("excel", 15)
    )
    win.complexity_open_report_btn.setObjectName("pickerButton")
    win.complexity_open_report_btn.setFixedSize(160, 40)
    win.complexity_open_report_btn.setEnabled(False)  # enabled once a report is selected

    gen_row.addWidget(win.complexity_generate_btn)
    gen_row.addSpacing(12)
    gen_row.addWidget(win.complexity_generate_html_btn)
    gen_row.addSpacing(12)
    gen_row.addWidget(win.complexity_open_report_btn)
    gen_row.addStretch()
    rc_layout.addLayout(gen_row)

    info_lbl = QLabel(
        "Use Browse Report to select an Excel report file.\n"
        "Complexity Settings contains only weights and bands.\n"
        "Compatibility Settings contains only handled scenarios."
    )
    info_lbl.setObjectName("complexityInfoLabel")
    info_lbl.setWordWrap(True)
    rc_layout.addWidget(info_lbl)

    # ── Progress + Logs ──────────────────────────────────────────────────────
    win.complexity_progress_bar = QProgressBar()
    win.complexity_progress_bar.setRange(0, 100)
    win.complexity_progress_bar.setValue(0)
    win.complexity_progress_bar.setFixedHeight(20)
    win.complexity_progress_bar.setTextVisible(False)
    win.complexity_progress_bar.setVisible(False)

    win.complexity_status_lbl = QLabel("")
    win.complexity_status_lbl.setObjectName("panelSubtitle")
    win.complexity_status_lbl.setVisible(False)

    win.complexity_log = QTextEdit()
    win.complexity_log.setReadOnly(True)
    win.complexity_log.setFixedHeight(110)
    win.complexity_log.setPlaceholderText("Log will appear here …")
    win.complexity_log.setVisible(False)

    rc_layout.addWidget(win.complexity_progress_bar)
    rc_layout.addWidget(win.complexity_status_lbl)
    rc_layout.addWidget(win.complexity_log)

    layout.addWidget(report_card)
    layout.addStretch()

    # ── State ────────────────────────────────────────────────────────────────
    win._complexity_weights = None
    win._complexity_bands = None
    win._handled_scenarios = _load_handled_scenarios()
    win._cx_thread = None
    win._cx_worker = None

    # ── Actions ──────────────────────────────────────────────────────────────
    def browse_report():
        path, _ = QFileDialog.getOpenFileName(
            win, "Select Report Excel", "",
            "Excel Files (*.xlsx *.xls *.xlsm);;All Files (*.*)"
        )
        if path:
            win.complexity_report_display.setText(path)
            win.complexity_open_report_btn.setEnabled(True)

    def clear_report():
        win.complexity_report_display.clear()
        win.complexity_open_report_btn.setEnabled(False)

    def open_report():
        import subprocess, sys
        from PySide6.QtCore import QUrl
        from PySide6.QtGui import QDesktopServices
        from PySide6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel as _QLabel

        excel_path = win.complexity_report_display.text().strip()
        html_path  = getattr(win, "_last_complexity_html", "") or ""

        def _open(path):
            if sys.platform.startswith("win"):
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])

        # Build list of available files
        choices = []
        if excel_path and os.path.isfile(excel_path):
            choices.append(("📊  Open Excel Report", excel_path))
        if html_path and os.path.isfile(html_path):
            choices.append(("🌐  Open HTML Report", html_path))

        if not choices:
            QMessageBox.warning(win, "No Report", "No valid report file is selected.")
            return

        if len(choices) == 1:
            _open(choices[0][1])
            return

        # Both Excel and HTML available — show choice dialog
        dlg = QDialog(win)
        dlg.setWindowTitle("Open Report")
        dlg.setFixedSize(380, 120)
        vl = QVBoxLayout(dlg)
        vl.setSpacing(12)
        vl.addWidget(_QLabel("Which report would you like to open?"))
        btn_row = QHBoxLayout()
        for label, fpath in choices:
            btn = QPushButton(label)
            btn.setFixedHeight(38)
            btn.clicked.connect(lambda _, p=fpath, d=dlg: (_open(p), d.accept()))
            btn_row.addWidget(btn)
        vl.addLayout(btn_row)
        dlg.exec()

    # Complexity Settings = weights + bands only
    def open_complexity_settings():
        weights = win._complexity_weights
        bands = win._complexity_bands
        if weights is None:
            weights = [list(r) for r in ComplexitySettingsDialog.DEFAULT_WEIGHTS]
        if bands is None:
            bands = [list(r) for r in ComplexitySettingsDialog.DEFAULT_BANDS]

        dlg = ComplexitySettingsDialog(
            win,
            weights=weights,
            bands=bands,
            handled_scenarios=None
        )
        if dlg.exec():
            win._complexity_weights = dlg.get_weights()
            win._complexity_bands = dlg.get_bands()

    # Compatibility Settings = handled scenarios only
    def open_compatibility_settings():
        dlg = ComplexitySettingsDialog(
            win,
            weights=None,
            bands=None,
            handled_scenarios=win._handled_scenarios
        )
        if dlg.exec():
            win._handled_scenarios = dlg.get_handled_scenarios()
            _save_handled_scenarios(win._handled_scenarios)

    def _set_running_ui(running: bool):
        win.complexity_generate_btn.setEnabled(not running)
        win.complexity_generate_btn.setText("Running …" if running else "Generate Report")
        win.complexity_progress_bar.setVisible(running)
        win.complexity_status_lbl.setVisible(running)
        win.complexity_log.setVisible(running)

    def _on_progress(pct: int, msg: str):
        win.complexity_progress_bar.setValue(max(0, min(100, int(pct))))
        win.complexity_status_lbl.setText(msg)

    def _on_log(msg: str):
        win.complexity_log.append(msg)

    def _on_generate_done(path: str):
        _set_running_ui(False)
        win.complexity_progress_bar.setValue(100)
        win.complexity_status_lbl.setText("Done")
        win.complexity_open_report_btn.setEnabled(True)
        QMessageBox.information(
            win, "Success",
            f"Compatibility sheet appended successfully.\n\nUpdated report:\n{path}"
        )
        win._cx_thread = None
        win._cx_worker = None

    def _on_generate_error(message: str):
        _set_running_ui(False)
        QMessageBox.critical(win, "Error", message)
        win._cx_thread = None
        win._cx_worker = None

    def generate():
        report_path = win.complexity_report_display.text().strip()
        if not report_path or not os.path.isfile(report_path):
            QMessageBox.warning(
                win, "No Report Loaded",
                "No report is loaded.\n\nSelect a report file first."
            )
            return

        target_entry = next(
            (s for s in getattr(win, "available_sources", []) if s.get("type") == "target"),
            None
        )
        if not target_entry:
            QMessageBox.warning(
                win, "No Source Loaded",
                "No target source is loaded.\n\nPlease go to the Input page and load the target source."
            )
            return

        src = target_entry["path"]
        if not src or not os.path.isdir(src):
            QMessageBox.warning(win, "Source Not Found", f"Target folder not found:\n{src}")
            return

        weights = win._complexity_weights if win._complexity_weights else None
        bands = win._complexity_bands if win._complexity_bands else None

        _set_running_ui(True)
        win.complexity_progress_bar.setValue(0)
        win.complexity_status_lbl.setText("Starting …")
        win.complexity_log.clear()

        class _CxThread(QThread):
            def __init__(self, worker):
                super().__init__()
                self._worker = worker

            def run(self):
                self._worker.run()

        worker = ComplexityAppendWorker(
            report_path=report_path,
            source_folder=src,
            weights=weights,
            bands=bands,
            handled_scenarios=win._handled_scenarios,
        )
        thread = _CxThread(worker)

        worker.progress.connect(_on_progress)
        worker.log.connect(_on_log)
        worker.finished.connect(_on_generate_done)
        worker.error.connect(_on_generate_error)

        worker.finished.connect(thread.quit)
        worker.error.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        worker.error.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)

        win._cx_thread = thread
        win._cx_worker = worker

        thread.start()

    # ── Connect ──────────────────────────────────────────────────────────────
    win.complexity_browse_btn.clicked.connect(browse_report)
    win.complexity_clear_btn.clicked.connect(clear_report)
    win.complexity_open_report_btn.clicked.connect(open_report)
    win.complexity_settings_btn.clicked.connect(open_complexity_settings)
    win.compatibility_settings_btn.clicked.connect(open_compatibility_settings)
    win.complexity_generate_btn.clicked.connect(generate)
    win.complexity_generate_html_btn.clicked.connect(lambda: win._on_complexity_generate_html())

    scroll = win.make_scroll_page(page)
    win.stack.addWidget(scroll)
    win.pages["complexity"] = scroll