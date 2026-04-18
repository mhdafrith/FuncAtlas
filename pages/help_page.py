"""pages/help_page.py — Help & Guidance page."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFrame, QLabel,
    QPushButton, QGridLayout
)
from ui.widgets import SectionTitle


def create_help_page(win):
    page = QWidget()
    layout = QVBoxLayout(page)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(14)

    # ── top card ──────────────────────────────────────────────────────────────
    top_card = QFrame()
    top_card.setObjectName("pageCard")
    top_layout = QVBoxLayout(top_card)
    top_layout.setContentsMargins(22, 20, 22, 20)
    top_layout.setSpacing(14)
    top_layout.addWidget(SectionTitle("Help & Guidance", "Instructions, workflow rules, and user support."))

    intro = QLabel(
        "This page is the built-in guide for FuncAtlas. Use it when you need "
        "the actual workflow, required setup, or the report log viewer."
    )
    intro.setObjectName("panelSubtitle")
    intro.setWordWrap(True)
    top_layout.addWidget(intro)

    btn_row = QHBoxLayout()
    btn_row.setSpacing(12)
    btn_how = QPushButton("📖  How to Use")
    btn_pre = QPushButton("✅  Prerequisites")
    btn_log = QPushButton("📋  View Log File")
    for btn in (btn_how, btn_pre, btn_log):
        btn.setObjectName("smallPrimaryButton")
        btn.setCursor(Qt.PointingHandCursor)
        btn.setMinimumHeight(40)
        btn_row.addWidget(btn)
    btn_row.addStretch()
    btn_how.clicked.connect(win._show_report_how_to_use)
    btn_pre.clicked.connect(win._show_report_prerequisites)
    btn_log.clicked.connect(win._show_report_log_dialog)
    top_layout.addLayout(btn_row)
    layout.addWidget(top_card)

    # ── step boxes ────────────────────────────────────────────────────────────
    grid_wrap = QWidget()
    grid = QGridLayout(grid_wrap)
    grid.setContentsMargins(0, 0, 0, 0)
    grid.setHorizontalSpacing(14)
    grid.setVerticalSpacing(14)

    def make_help_box(number: str, title: str, body: str):
        box = QFrame()
        box.setObjectName("pageCard")
        bl = QHBoxLayout(box)
        bl.setContentsMargins(18, 16, 18, 16)
        bl.setSpacing(14)
        badge = QLabel(number)
        badge.setAlignment(Qt.AlignCenter)
        badge.setFixedSize(38, 38)
        badge.setStyleSheet(
            f"background: {win.accent_color.name()}; color: white; "
            f"border-radius: 10px; font-size: 16px; font-weight: 900;"
        )
        txt_wrap = QVBoxLayout()
        txt_wrap.setContentsMargins(0, 0, 0, 0)
        txt_wrap.setSpacing(4)
        t = QLabel(title)
        t.setStyleSheet(
            f"color: {win.theme['text_primary']}; font-size: 15px; font-weight: 900;"
        )
        t.setWordWrap(True)
        b = QLabel(body)
        b.setObjectName("panelSubtitle")
        b.setWordWrap(True)
        txt_wrap.addWidget(t)
        txt_wrap.addWidget(b)
        bl.addWidget(badge, 0, Qt.AlignTop)
        bl.addLayout(txt_wrap, 1)
        return box

    step_data = [
        ("1", "Select Base Folder",
         "Choose the original/reference source folder. Use the Input page to load the correct base before running anything."),
        ("2", "Select Target Folder",
         "Choose the modified or new source folder that must be compared against the reference inputs."),
        ("3", "Select Output Folder",
         "Pick a writable folder. FuncAtlas saves extracted functions and the Excel report there."),
        ("4", "Load Function List",
         "If your workflow depends on a function list, submit the TXT or XLSX files first from the Input area."),
        ("5", "Run Report Generator",
         "The report flow extracts functions first, then compares them, then writes the Excel report."),
        ("6", "Watch Progress",
         "Each base is processed in order. Waiting, Running, and Complete states tell you exactly where the pipeline is."),
        ("7", "Open Result",
         "After completion, use Open Report. If something failed, inspect the detailed log instead of guessing."),
    ]
    for i, (num, title, body) in enumerate(step_data):
        grid.addWidget(make_help_box(num, title, body), i, 0)
    layout.addWidget(grid_wrap)

    # ── footer warnings ───────────────────────────────────────────────────────
    foot1 = QFrame()
    foot1.setObjectName("pageCard")
    foot1.setStyleSheet("QFrame#pageCard { background: #3A2200; border: 1px solid #E5A100; border-radius: 16px; }")
    f1 = QHBoxLayout(foot1)
    f1.setContentsMargins(18, 14, 18, 14)
    warn = QLabel("⚠ Prerequisites: Windows OS · Microsoft Excel installed · Target/reference inputs must be submitted before running reports.")
    warn.setStyleSheet("color: #FFD37A; font-size: 12px; font-weight: 700; background: transparent;")
    warn.setWordWrap(True)
    f1.addWidget(warn)
    layout.addWidget(foot1)

    foot2 = QFrame()
    foot2.setObjectName("pageCard")
    foot2.setStyleSheet("QFrame#pageCard { background: #042B24; border: 1px solid #00D39A; border-radius: 16px; }")
    f2 = QHBoxLayout(foot2)
    f2.setContentsMargins(18, 14, 18, 14)
    tip = QLabel("💡 Tip: Keep output on a local disk and close any open Excel report file before generating again. Locked files are a common failure point.")
    tip.setStyleSheet("color: #8AF5D0; font-size: 12px; font-weight: 700; background: transparent;")
    tip.setWordWrap(True)
    f2.addWidget(tip)
    layout.addWidget(foot2)
    layout.addStretch()

    scroll = win.make_scroll_page(page)
    win.stack.addWidget(scroll)
    win.pages["help"] = scroll