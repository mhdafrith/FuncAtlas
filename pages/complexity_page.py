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

SCENARIO_GROUPS = [
    ("Control Flow", [
        "If Statement", "Else Branch", "Else-if Chain", "Switch Statement",
        "Case Label", "Default Label", "Break Statement", "For Loop",
        "Return Statement", "Continue Statement", "While Loop", "Do...While",
    ]),
    ("Language Constructs", [
        "Function Call", "Function Definition", "Function Declaration",
        "Address-of Operator", "Pointer Declaration", "Pointer Dereference",
        "Arrow Operator", "Void Pointer", "Pointer Arithmetic",
        "Array Access", "Array Declaration", "String Literal", "Char Array",
        "Cast Operation", "Typedef", "Enum Definition", "Sizeof Operator",
        "Const Declaration", "Signed/Unsigned", "Short/Long", "Static Keyword",
        "Compound Assignment", "Increment", "Decrement",
        "Bitwise AND", "Bitwise OR", "Bitwise XOR", "Bitwise NOT",
        "Left Shift", "Right Shift", "Logical AND", "Logical OR", "Logical NOT",
        "Designated Initializer", "Compound Literal", "Bit Field",
    ]),
    ("Preprocessor / Safety / Memory", [
        "Macro Definition", "Ifdef Directive", "If Directive", "Elif Directive",
        "Else Directive", "Endif Directive", "Pragma Directive",
        "NULL Check", "NULL Assignment",
        "memset", "memcpy", "fabs",
    ]),
]

ALL_SCENARIOS = [sc for _, scs in SCENARIO_GROUPS for sc in scs]


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
        ("If Statement",           1), ("Else Branch",            1),
        ("Else-if Chain",          2), ("Switch Statement",       2),
        ("Case Label",             1), ("Default Label",          1),
        ("Break Statement",        1), ("For Loop",               2),
        ("Return Statement",       1), ("Continue Statement",     1),
        ("While Loop",             2), ("Do...While",             2),
        ("Function Call",          4), ("Function Definition",    3),
        ("Function Declaration",   2), ("Address-of Operator",    3),
        ("Pointer Declaration",    4), ("Pointer Dereference",    4),
        ("Arrow Operator",         3), ("Void Pointer",           5),
        ("Pointer Arithmetic",     5), ("Array Access",           2),
        ("Array Declaration",      2), ("String Literal",         1),
        ("Char Array",             2), ("Cast Operation",         3),
        ("Typedef",                2), ("Enum Definition",        2),
        ("Sizeof Operator",        2), ("Const Declaration",      1),
        ("Signed/Unsigned",        1), ("Short/Long",             1),
        ("Static Keyword",         2), ("Compound Assignment",    1),
        ("Increment",              1), ("Decrement",              1),
        ("Bitwise AND",            2), ("Bitwise OR",             2),
        ("Bitwise XOR",            2), ("Bitwise NOT",            2),
        ("Left Shift",             2), ("Right Shift",            2),
        ("Logical AND",            1), ("Logical OR",             1),
        ("Logical NOT",            1), ("Designated Initializer", 2),
        ("Compound Literal",       2), ("Bit Field",              3),
        ("Macro Definition",       2), ("Ifdef Directive",        1),
        ("If Directive",           1), ("Elif Directive",         1),
        ("Else Directive",         1), ("Endif Directive",        1),
        ("Pragma Directive",       1), ("NULL Check",             3),
        ("NULL Assignment",        2), ("memset",                 3),
        ("memcpy",                 3), ("fabs",                   2),
    ]

    DEFAULT_BANDS = [
        ("Low",       1,   5),
        ("Medium",    6,  12),
        ("High",     13,  25),
        ("Very High", 26,  40),
        ("Complex",  41, 999),
    ]

    # Colour palette for the 5 bands (Low → Complex)
    BAND_COLORS = ["#27AE60", "#2980B9", "#F39C12", "#E67E22", "#C0392B"]

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

        self._show_complexity    = (weights is not None or bands is not None)
        self._show_compatibility = (handled_scenarios is not None)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)

        shell = QFrame()
        shell.setObjectName("complexityDialogShell")
        shell_layout = QVBoxLayout(shell)
        shell_layout.setContentsMargins(32, 14, 32, 14)   # tighter padding — reclaimed space
        shell_layout.setSpacing(10)

        # ── Header row: action buttons only (title removed — reclaim vertical space) ──
        head_row = QHBoxLayout()
        head_row.addStretch()

        if self._show_complexity:
            default_btn = QPushButton("  \u21BA  Default")
            default_btn.setObjectName("complexityDefaultBtn")
            default_btn.setFixedSize(130, 40)
            default_btn.setToolTip("Reset all weights and bands to their default values")
            default_btn.clicked.connect(self._on_reset_defaults)
            head_row.addWidget(default_btn)
            head_row.addSpacing(8)

            save_btn = QPushButton("  Save \u0026 Close")
            save_btn.setObjectName("complexitySaveBtn")
            save_btn.setFixedSize(150, 40)
            save_btn.clicked.connect(self._on_save)
            head_row.addWidget(save_btn)
            head_row.addSpacing(8)

        close_btn = QPushButton("  Close")
        close_btn.setObjectName("complexityCloseBtn")
        close_btn.setFixedSize(120, 40)
        close_btn.clicked.connect(self.reject)
        head_row.addWidget(close_btn)
        shell_layout.addLayout(head_row)

        # ── Body ─────────────────────────────────────────────────────────────
        body_row = QHBoxLayout()
        body_row.setSpacing(24)

        if self._show_complexity:
            # ── Col 1: Weightage (scrollable) ─────────────────────────────────
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

            # ── Col 2: Complexity Bands + Band Scale Preview ───────────────────
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
                "Example: Low=1–5 · Medium=6–12 · High=13–25 · "
                "Very High=26–40 · Complex=41+"
            )
            note.setObjectName("complexityNote")
            note.setWordWrap(True)
            ml.addWidget(note)

            # ── Band Scale Preview — fills the empty gap below the bands ──────
            ml.addWidget(self._build_band_preview())
            ml.addStretch()

            body_row.addWidget(mid_frame, 1)

        # ── Compatibility panel ───────────────────────────────────────────────
        if self._show_compatibility:
            compat_frame = QFrame()
            compat_frame.setObjectName("compatFrame")
            cl = QVBoxLayout(compat_frame)
            cl.setContentsMargins(0, 0, 0, 0)
            cl.setSpacing(6)          # tight — eliminates dead vertical space

            cl.addWidget(self._section_hdr(
                "Compatibility Scenarios",
                "Click to mark which C scenarios your system handles"
            ))

            # ── Handled Scenarios Summary Box ────────────────────────────────
            self._summary_frame = QFrame()
            self._summary_frame.setObjectName("handledSummaryBox")
            summary_layout = QVBoxLayout(self._summary_frame)
            summary_layout.setContentsMargins(12, 6, 12, 6)
            summary_layout.setSpacing(2)

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

            # ── Modify / Save / Cancel button row (directly below summary) ────
            self._edit_mode = False
            btn_row = QHBoxLayout()
            btn_row.setContentsMargins(0, 2, 0, 2)
            btn_row.addStretch()

            self.modify_btn = QPushButton("  ✎  Modify")
            self.modify_btn.setObjectName("compatModifyBtn")
            self.modify_btn.setFixedSize(130, 30)
            self.modify_btn.clicked.connect(self._enable_edit_mode)

            self.save_edit_btn = QPushButton("✔  Save Changes")
            self.save_edit_btn.setObjectName("compatSaveEditBtn")
            self.save_edit_btn.setFixedSize(150, 30)
            self.save_edit_btn.clicked.connect(self._save_changes)

            self.cancel_edit_btn = QPushButton("✖  Cancel")
            self.cancel_edit_btn.setObjectName("compatCancelEditBtn")
            self.cancel_edit_btn.setFixedSize(110, 30)
            self.cancel_edit_btn.clicked.connect(self._cancel_edit)

            btn_row.addWidget(self.modify_btn)
            btn_row.addSpacing(8)
            btn_row.addWidget(self.save_edit_btn)
            btn_row.addSpacing(8)
            btn_row.addWidget(self.cancel_edit_btn)
            cl.addLayout(btn_row)

            # ── 4-column grid (all 60 scenarios visible at once) ──────────────
            from PySide6.QtWidgets import QSizePolicy

            grid_container = QFrame()
            grid_container.setObjectName("compatGridContainer")
            grid_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            grid_outer = QVBoxLayout(grid_container)
            grid_outer.setContentsMargins(0, 0, 0, 0)
            grid_outer.setSpacing(4)

            self._status_btns = {}

            # Distribute scenarios evenly across 4 columns
            all_flat = [(sc, grp) for grp, scs in SCENARIO_GROUPS for sc in scs]
            total = len(all_flat)
            NUM_COLS = 4
            base_per_col = total // NUM_COLS
            remainder    = total % NUM_COLS
            col_sizes = [base_per_col + (1 if i < remainder else 0) for i in range(NUM_COLS)]

            col_frames_layout = QHBoxLayout()
            col_frames_layout.setSpacing(6)
            col_frames_layout.setContentsMargins(0, 0, 0, 0)

            sc_iter = iter(all_flat)
            for col_i, col_size in enumerate(col_sizes):
                col_frame = QFrame()
                col_frame.setObjectName("compatColFrame")
                col_frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
                col_vl = QVBoxLayout(col_frame)
                col_vl.setContentsMargins(0, 0, 0, 0)
                col_vl.setSpacing(0)

                cur_group = None
                local_row = 0
                for _ in range(col_size):
                    item = next(sc_iter, None)
                    if item is None:
                        break
                    sc, grp = item

                    if grp != cur_group:
                        cur_group = grp
                        sec_hdr = QFrame()
                        sec_hdr.setObjectName("compatSectionHdr")
                        sec_hdr.setFixedHeight(18)
                        sec_hl = QHBoxLayout(sec_hdr)
                        sec_hl.setContentsMargins(6, 1, 6, 1)
                        lbl = QLabel(grp)
                        lbl.setObjectName("compatSectionHdrLabel")
                        sec_hl.addWidget(lbl)
                        col_vl.addWidget(sec_hdr)

                    row_w, badge = self._scenario_status_row(sc, sc in self._handled, local_row)
                    row_w.setFixedHeight(26)
                    col_vl.addWidget(row_w)
                    self._status_btns[sc] = badge
                    local_row += 1

                col_vl.addStretch()          # push content to top; aligns rows across columns
                col_frames_layout.addWidget(col_frame, 1)

            grid_outer.addLayout(col_frames_layout, 1)
            cl.addWidget(grid_container, 1)

            self._set_edit_enabled(False)
            self._refresh_summary()

            body_row.addWidget(compat_frame, 1)

        if not self._show_complexity and not self._show_compatibility:
            fallback = QLabel("No settings section selected.")
            fallback.setObjectName("complexityNote")
            body_row.addWidget(fallback)

        shell_layout.addLayout(body_row, 1)
        outer.addWidget(shell)
        self._apply_styles()

        # Compat-only: auto-fit to 92 % screen height
        if self._show_compatibility and not self._show_complexity:
            from PySide6.QtGui import QGuiApplication
            screen = QGuiApplication.primaryScreen()
            if screen:
                sg = screen.availableGeometry()
                h = int(sg.height() * 0.92)
                w = min(int(sg.width() * 0.90), 1400)
                self.resize(w, h)
                self.setMinimumHeight(h)

    # ── Band Scale Preview ────────────────────────────────────────────────────
    def _build_band_preview(self) -> QFrame:
        """
        Visual band-scale preview that fills the empty gap below the bands
        table.  Each band renders as a proportional coloured horizontal bar.
        """
        preview = QFrame()
        preview.setObjectName("bandPreviewFrame")
        pl = QVBoxLayout(preview)
        pl.setContentsMargins(16, 14, 16, 14)
        pl.setSpacing(10)

        hdr = QLabel("Band Scale Preview")
        hdr.setObjectName("complexitySectionTitle")
        pl.addWidget(hdr)

        sub = QLabel("Relative bar width represents the score range of each level")
        sub.setObjectName("complexitySectionSub")
        pl.addWidget(sub)

        # Compute max display range (cap ∞ at 80 for visual)
        max_range = max((min(e, 80) - s + 1) for _, s, e in self._bands) or 1

        for i, (lv, s, e) in enumerate(self._bands):
            color = self.BAND_COLORS[i]
            rng   = min(e, 80) - s + 1
            ratio = max(rng / max_range, 0.08)

            row_layout = QHBoxLayout()
            row_layout.setSpacing(10)

            lv_lbl = QLabel(lv)
            lv_lbl.setObjectName("bandPreviewLevelLabel")
            lv_lbl.setFixedWidth(72)
            lv_lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            row_layout.addWidget(lv_lbl)

            # Proportional coloured bar using stretch factors
            bar_wrap = QHBoxLayout()
            bar_wrap.setContentsMargins(0, 0, 0, 0)
            bar_wrap.setSpacing(0)

            bar = QFrame()
            bar.setObjectName(f"bandBar_{i}")
            bar.setFixedHeight(22)
            bar.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            bar.setStyleSheet(
                f"QFrame#bandBar_{i} {{"
                f"  background: qlineargradient(x1:0,y1:0,x2:1,y2:0,"
                f"    stop:0 {color}, stop:1 {color}88);"
                f"  border-radius: 5px;"
                f"  min-width: {max(int(ratio * 220), 18)}px;"
                f"}}"
            )
            bar_wrap.addWidget(bar, int(ratio * 10))
            bar_wrap.addStretch(max(int((1 - ratio) * 10), 1))

            bar_container = QWidget()
            bar_container.setLayout(bar_wrap)
            row_layout.addWidget(bar_container, 1)

            range_str = f"{s} – ∞" if e >= 999 else f"{s} – {e}"
            range_lbl = QLabel(range_str)
            range_lbl.setObjectName("bandPreviewRangeLabel")
            range_lbl.setFixedWidth(58)
            range_lbl.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            row_layout.addWidget(range_lbl)

            pl.addLayout(row_layout)

        return preview

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
        rl.setContentsMargins(6, 2, 6, 2)
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

        # Coloured left-edge indicator for the band level
        indicator = QFrame()
        indicator.setFixedSize(4, 22)
        indicator.setStyleSheet(
            f"background: {self.BAND_COLORS[idx]}; border-radius: 2px;"
        )
        rl.addWidget(indicator)

        lb = QLabel(level)
        lb.setObjectName("complexityScenarioLabel")
        lb.setMinimumWidth(72)
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
        row = QFrame()
        row.setObjectName("compatRowEven" if idx % 2 == 0 else "compatRowOdd")
        rl = QHBoxLayout(row)
        rl.setContentsMargins(10, 3, 10, 3)
        rl.setSpacing(6)

        lb = QLabel(scenario)
        lb.setObjectName("complexityScenarioLabel")

        badge = QPushButton()
        badge.setObjectName("statusBadgeHandled" if is_handled else "statusBadgeNotHandled")
        badge.setFixedSize(80, 18)
        badge.setProperty("handled", is_handled)
        self._update_badge_text(badge, is_handled)
        badge.clicked.connect(lambda _, b=badge, s=scenario: self._toggle_status(b, s))
        badge.setEnabled(False)

        rl.addWidget(lb, 1)
        rl.addWidget(badge)
        return row, badge

    def _update_badge_text(self, badge: QPushButton, is_handled: bool):
        badge.setText("✓ Handled" if is_handled else "○")

    def _toggle_status(self, badge: QPushButton, scenario: str):
        new_state = not badge.property("handled")
        badge.setProperty("handled", new_state)
        self._update_badge_text(badge, new_state)
        badge.setObjectName("statusBadgeHandled" if new_state else "statusBadgeNotHandled")
        badge.setStyle(badge.style())
        self._refresh_summary()

    def _refresh_summary(self):
        if not hasattr(self, "_status_btns") or not hasattr(self, "_summary_text_lbl"):
            return
        handled_list = [sc for sc, badge in self._status_btns.items() if badge.property("handled")]
        count = len(handled_list)
        total = len(self._status_btns)
        self._summary_count_lbl.setText(f"{count} / {total}")
        text = "  •  ".join(handled_list) if handled_list else "None marked as handled yet."
        self._summary_text_lbl.setText(text)
        self._summary_text_lbl.setVisible(True)

    # ── Compatibility edit actions ─────────────────────────────────────────────
    def _enable_edit_mode(self):
        msg = QMessageBox(self)
        msg.setWindowTitle("Modify Scenarios")
        msg.setIcon(QMessageBox.Warning)
        msg.setText("<b>You are about to modify handled scenarios.</b>")
        msg.setInformativeText(
            "This will change which C scenarios are marked as handled by your system.\n\n"
            "Click each scenario's status badge to toggle between Handled / Not Handled.\n"
            "Use Save Changes to apply or Cancel to discard."
        )
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        msg.setDefaultButton(QMessageBox.Ok)
        msg.button(QMessageBox.Ok).setText("  Proceed  ")
        msg.button(QMessageBox.Cancel).setText("  Cancel  ")
        if msg.exec() == QMessageBox.Ok:
            self._handled_snapshot = {
                sc for sc, badge in self._status_btns.items() if badge.property("handled")
            }
            self._edit_mode = True
            self._set_edit_enabled(True)

    def _set_edit_enabled(self, enabled: bool):
        for badge in self._status_btns.values():
            badge.setEnabled(enabled)
        self.save_edit_btn.setVisible(enabled)
        self.cancel_edit_btn.setVisible(enabled)
        self.modify_btn.setVisible(not enabled)
        if enabled:
            for badge in self._status_btns.values():
                badge.setStyle(badge.style())

    def _save_changes(self):
        self._handled = {
            sc for sc, badge in self._status_btns.items() if badge.property("handled")
        }
        _save_handled_scenarios(self._handled)
        QMessageBox.information(self, "Saved", "Handled scenarios saved successfully.")
        self._edit_mode = False
        self._set_edit_enabled(False)

    def _cancel_edit(self):
        for sc, badge in self._status_btns.items():
            was = sc in self._handled_snapshot
            badge.setProperty("handled", was)
            self._update_badge_text(badge, was)
            badge.setObjectName("statusBadgeHandled" if was else "statusBadgeNotHandled")
            badge.setStyle(badge.style())
        self._edit_mode = False
        self._set_edit_enabled(False)
        self._refresh_summary()

    def _on_save(self):
        self.accept()

    def _on_reset_defaults(self):
        reply = QMessageBox.question(
            self, "Reset to Defaults",
            "Reset all weights and complexity bands to their default values?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return
        default_w = {sc: val for sc, val in self.DEFAULT_WEIGHTS}
        for sc, sp in self._weight_spins:
            if sc in default_w:
                sp.setValue(default_w[sc])
        default_b = {lv: (s, e) for lv, s, e in self.DEFAULT_BANDS}
        for lv, ss, es in self._band_spins:
            if lv in default_b:
                ss.setValue(default_b[lv][0])
                es.setValue(default_b[lv][1])

    def _apply_styles(self):
        self.setStyleSheet("""
            QDialog, QFrame#complexityDialogShell { background: #0D1B2A; }
            QWidget { background: #0D1B2A; color: #C8D8E8; }

            /* ── Buttons ── */
            QPushButton#complexitySaveBtn {
                background: #27AE60; color: white; border: none;
                border-radius: 12px; font-size: 14px; font-weight: 800;
            }
            QPushButton#complexitySaveBtn:hover { background: #2ECC71; }
            QPushButton#complexityDefaultBtn {
                background: #1A3050; color: #8AAAC8;
                border: 2px solid #1D3347;
                border-radius: 12px; font-size: 14px; font-weight: 800;
            }
            QPushButton#complexityDefaultBtn:hover {
                background: #E67E22; color: white; border-color: #F39C12;
            }
            QPushButton#complexityCloseBtn {
                background: #C0392B; color: white; border: none;
                border-radius: 12px; font-size: 14px; font-weight: 800;
            }
            QPushButton#complexityCloseBtn:hover { background: #E74C3C; }

            /* ── Section typography ── */
            QLabel#complexitySectionTitle { color: white; font-size: 15px; font-weight: 900; }
            QLabel#complexitySectionSub   { color: #8899AA; font-size: 11px; }

            /* ── Weightage / band tables ── */
            QFrame#complexityTableInner {
                background: #101E2E; border: 1px solid #1D3347; border-radius: 10px;
            }
            QFrame#complexityTableHeader {
                background: #162840; border-bottom: 1px solid #1D3347;
                border-top-left-radius: 10px; border-top-right-radius: 10px;
            }
            QLabel#complexityColHeader    { color: white; font-weight: 800; font-size: 12px; }
            QFrame#complexityRowEven      { background: #101E2E; max-height: 26px; }
            QFrame#complexityRowOdd       { background: #0D1B2A; max-height: 26px; }
            QLabel#complexityScenarioLabel { color: #C8D8E8; font-size: 9px; }
            QPushButton#complexityStepBtn {
                background: #1A3050; color: #1E90FF; border: 1px solid #1D3347;
                border-radius: 8px; font-size: 15px; font-weight: 900;
            }
            QPushButton#complexityStepBtn:hover { background: #1E90FF; color: white; }
            QPushButton#complexityStepPlusBtn {
                background: #2ECC71; color: white; border: 2px solid #27AE60;
                border-radius: 8px; font-size: 18px; font-weight: 900;
            }
            QPushButton#complexityStepPlusBtn:hover {
                background: #58D68D; color: white; border-color: #2ECC71;
            }
            QSpinBox#complexitySpinBox {
                background: #101E2E; color: white; border: 1px solid #1D3347;
                border-radius: 8px; font-size: 13px; font-weight: 700; padding: 4px;
            }
            QLabel#complexityNote { color: #7A8FA0; font-size: 11px; }
            QScrollBar:vertical { background: #0D1B2A; width: 8px; border-radius: 4px; }
            QScrollBar::handle:vertical { background: #1D3347; border-radius: 4px; min-height: 20px; }

            /* ── Band Scale Preview ── */
            QFrame#bandPreviewFrame {
                background: #0A1628; border: 1px solid #1D3347; border-radius: 10px;
            }
            QLabel#bandPreviewLevelLabel {
                color: #C8D8E8; font-size: 11px; font-weight: 700;
            }
            QLabel#bandPreviewRangeLabel {
                color: #8899AA; font-size: 11px; font-weight: 600;
            }

            /* ── Handled Scenarios Summary Box ── */
            QFrame#handledSummaryBox {
                background: #0A1E10; border: 1px solid #27AE60; border-radius: 10px;
            }
            QLabel#handledSummaryTitle  { color: #2ECC71; font-size: 12px; font-weight: 900; }
            QLabel#handledSummaryCount  {
                color: #27AE60; font-size: 12px; font-weight: 800;
                background: #1A5C35; border-radius: 6px; padding: 2px 8px;
            }
            QLabel#handledSummaryText   {
                color: #8ECFA8; font-size: 11px; font-weight: 600; line-height: 1.5;
            }

            /* ── Compat 4-column grid ── */
            QFrame#compatGridContainer  { background: transparent; }
            QFrame#compatColFrame {
                background: #101E2E; border: 1px solid #1D3347; border-radius: 8px;
            }
            QFrame#compatRowEven        { background: #101E2E; max-height: 26px; }
            QFrame#compatRowOdd         { background: #0D1B2A; max-height: 26px; }

            /* ── Compat group headers ── */
            QFrame#compatSectionHdr {
                background: #162840; border-radius: 3px;
                border-left: 3px solid #1E90FF;
                margin: 1px 2px 1px 2px; max-height: 18px;
            }
            QLabel#compatSectionHdrLabel {
                color: #7EB8E8; font-size: 8px; font-weight: 900;
                letter-spacing: 0.3px; padding: 0px;
            }

            /* ── Compatibility status badges ── */
            QPushButton#statusBadgeHandled {
                background: #1A5C35; color: #2ECC71;
                border: 1px solid #27AE60; border-radius: 4px;
                font-size: 9px; font-weight: 800; padding: 0 2px;
            }
            QPushButton#statusBadgeHandled:hover:enabled  { background: #27AE60; color: white; }
            QPushButton#statusBadgeHandled:disabled       {
                background: #1A5C35; color: #2ECC71; border-color: #27AE60;
            }
            QPushButton#statusBadgeNotHandled {
                background: #1E2A38; color: #556677;
                border: 1px solid #2A3F55; border-radius: 4px;
                font-size: 11px; font-weight: 800; padding: 0 2px;
            }
            QPushButton#statusBadgeNotHandled:hover:enabled {
                background: #1A3050; color: #C8D8E8; border-color: #1E90FF;
            }
            QPushButton#statusBadgeNotHandled:disabled    {
                background: #1E2A38; color: #445566; border-color: #1D3347;
            }

            /* ── Compat action buttons ── */
            QPushButton#compatModifyBtn {
                background: #162840; color: #8AAAC8; border: 2px solid #1D3347;
                border-radius: 10px; font-size: 12px; font-weight: 800;
            }
            QPushButton#compatModifyBtn:hover   { background: #1E4070; color: white; border-color: #1E90FF; }
            QPushButton#compatSaveEditBtn {
                background: #1A5C35; color: #2ECC71;
                border: 1px solid #27AE60; border-radius: 10px;
                font-size: 12px; font-weight: 800;
            }
            QPushButton#compatSaveEditBtn:hover  { background: #27AE60; color: white; }
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

    report_card = QFrame()
    report_card.setObjectName("pageCard")
    rc_layout = QVBoxLayout(report_card)
    rc_layout.setContentsMargins(22, 18, 22, 18)
    rc_layout.setSpacing(12)

    rc_layout.addWidget(SectionTitle("Report", ""))

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
    win.complexity_open_report_btn.setEnabled(False)

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

    win._complexity_weights = None
    win._complexity_bands = None
    win._handled_scenarios = _load_handled_scenarios()
    win._cx_thread = None
    win._cx_worker = None

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

        dlg = QDialog(win)
        dlg.setWindowTitle("Open Report")
        dlg.setFixedSize(380, 120)
        vl = QVBoxLayout(dlg)
        vl.setSpacing(12)
        vl.addWidget(_QLabel("Which report would you like to open?"))
        btn_row2 = QHBoxLayout()
        for label, fpath in choices:
            btn = QPushButton(label)
            btn.setFixedHeight(38)
            btn.clicked.connect(lambda _, p=fpath, d=dlg: (_open(p), d.accept()))
            btn_row2.addWidget(btn)
        vl.addLayout(btn_row2)
        dlg.exec()

    def open_complexity_settings():
        weights = win._complexity_weights or [list(r) for r in ComplexitySettingsDialog.DEFAULT_WEIGHTS]
        bands   = win._complexity_bands   or [list(r) for r in ComplexitySettingsDialog.DEFAULT_BANDS]
        dlg = ComplexitySettingsDialog(win, weights=weights, bands=bands, handled_scenarios=None)
        if dlg.exec():
            win._complexity_weights = dlg.get_weights()
            win._complexity_bands   = dlg.get_bands()

    def open_compatibility_settings():
        dlg = ComplexitySettingsDialog(win, weights=None, bands=None,
                                       handled_scenarios=win._handled_scenarios)
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
            QMessageBox.warning(win, "No Report Loaded",
                                "No report is loaded.\n\nSelect a report file first.")
            return

        target_entry = next(
            (s for s in getattr(win, "available_sources", []) if s.get("type") == "target"), None
        )
        if not target_entry:
            QMessageBox.warning(win, "No Source Loaded",
                                "No target source is loaded.\n\n"
                                "Please go to the Input page and load the target source.")
            return

        src = target_entry["path"]
        if not src or not os.path.isdir(src):
            QMessageBox.warning(win, "Source Not Found", f"Target folder not found:\n{src}")
            return

        weights = win._complexity_weights or None
        bands   = win._complexity_bands   or None

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
            report_path=report_path, source_folder=src,
            weights=weights, bands=bands,
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