"""pages/report_page.py — Report Generator page."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFrame, QLabel,
    QLineEdit, QProgressBar, QTextEdit, QScrollArea, QFileDialog, QMenu
)
from ui.widgets import SectionTitle, IconTextButton, StatChip, StepStatusWidget, CollapsiblePanel
from core.utils import normalize_path


def create_report_page(win):
    page = QWidget()
    layout = QVBoxLayout(page)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(12)

    # ── Input card ────────────────────────────────────────────────────────────
    input_card = QFrame()
    input_card.setObjectName("pageCard")
    input_layout = QVBoxLayout(input_card)
    input_layout.setContentsMargins(22, 18, 22, 18)
    input_layout.setSpacing(12)

    folder_section = QVBoxLayout()
    folder_section.setSpacing(6)
    folder_lbl = QLabel("Output Folder *")
    folder_lbl.setObjectName("fieldLabel")
    folder_section.addWidget(folder_lbl)

    folder_row = QHBoxLayout()
    folder_row.setSpacing(8)

    win._report_folder_display = QLineEdit()
    win._report_folder_display.setReadOnly(True)
    win._report_folder_display.setPlaceholderText("No folder selected")
    win._report_folder_display.setMinimumHeight(38)

    win._report_browse_btn = IconTextButton("Browse Folder", win.icons_white.icon("folder", 15))
    win._report_browse_btn.setObjectName("pickerButton")
    win._report_browse_btn.setFixedSize(142, 40)
    win._report_browse_btn.clicked.connect(win._pick_report_folder)

    win.report_clear_btn = IconTextButton("Clear", win.icons_white.icon("clear", 15))
    win.report_clear_btn.setObjectName("pickerButton")
    win.report_clear_btn.setFixedSize(110, 40)
    win.report_clear_btn.clicked.connect(win._on_report_clear)

    folder_row.addWidget(win._report_folder_display, 1)
    folder_row.addWidget(win._report_browse_btn)
    folder_row.addWidget(win.report_clear_btn)
    folder_section.addLayout(folder_row)
    input_layout.addLayout(folder_section)

    # Simple proxy so the rest of the main-window code can call .value() / .clear_selection()
    _display   = win._report_folder_display
    _path_list = []
    win._report_folder_path = _path_list

    class _FolderProxy:
        def value(self_):
            return _path_list[0] if _path_list else ""
        def clear_selection(self_):
            _path_list.clear()
            _display.clear()

    win.report_output_field = _FolderProxy()

    # ── action buttons ────────────────────────────────────────────────────────
    btn_row = QHBoxLayout()
    btn_row.setSpacing(8)

    win.report_generate_btn = IconTextButton("Generate Report", win.icons_white.icon("submit", 15))
    win.report_generate_btn.setObjectName("smallPrimaryButton")
    win.report_generate_btn.setFixedSize(172, 40)
    win.report_generate_btn.clicked.connect(win._on_report_generate)

    win.report_cancel_btn = IconTextButton("Cancel", win.icons_white.icon("clear", 15))
    win.report_cancel_btn.setObjectName("pickerButton")
    win.report_cancel_btn.setFixedSize(110, 40)
    win.report_cancel_btn.setVisible(False)
    win.report_cancel_btn.clicked.connect(win._on_report_cancel)

    win.report_open_btn = IconTextButton("Open Report", win.icons_white.icon("link", 15))
    win.report_open_btn.setObjectName("pickerButton")
    win.report_open_btn.setFixedSize(136, 40)
    win.report_open_btn.setEnabled(False)
    win.report_open_btn.clicked.connect(win._on_report_open)

    btn_row.addWidget(win.report_generate_btn)
    btn_row.addWidget(win.report_cancel_btn)
    btn_row.addWidget(win.report_open_btn)
    btn_row.addStretch()
    input_layout.addLayout(btn_row)
    layout.addWidget(input_card)

    # ── Progress card ─────────────────────────────────────────────────────────
    progress_card = QFrame()
    progress_card.setObjectName("pageCard")
    prog_layout = QVBoxLayout(progress_card)
    prog_layout.setContentsMargins(22, 18, 22, 18)
    prog_layout.setSpacing(12)
    prog_layout.addWidget(SectionTitle(
        "Progress",
        "Step-by-step extraction and comparison status. Each base is processed in order."
    ))

    win.report_summary_chips = {
        "steps":  StatChip("Pipeline steps", "0", "accent"),
        "phase":  StatChip("Current phase", "Idle"),
        "result": StatChip("Output", "Pending"),
    }
    win.report_summary_row = QWidget()
    chip_row = QHBoxLayout(win.report_summary_row)
    chip_row.setContentsMargins(0, 0, 0, 0)
    chip_row.setSpacing(12)
    for chip in win.report_summary_chips.values():
        chip.setMinimumHeight(76)
        chip_row.addWidget(chip)
    win.report_summary_row.setVisible(True)
    prog_layout.addWidget(win.report_summary_row)

    # Overall progress bar
    overall_lbl = QLabel("Overall Progress")
    overall_lbl.setStyleSheet("font-size: 11px; font-weight: 700; color: #888; margin-top: 4px;")
    prog_layout.addWidget(overall_lbl)

    win.report_progress_bar = QProgressBar()
    win.report_progress_bar.setRange(0, 100)
    win.report_progress_bar.setValue(0)
    win.report_progress_bar.setFixedHeight(8)
    win.report_progress_bar.setTextVisible(False)
    win.report_progress_bar.setObjectName("reportOverallProgress")
    prog_layout.addWidget(win.report_progress_bar)

    # Hidden step-level progress bar kept for code compatibility (not shown)
    win.report_step_progress_bar = QProgressBar()
    win.report_step_progress_bar.setRange(0, 100)
    win.report_step_progress_bar.setValue(0)
    win.report_step_progress_bar.setVisible(False)
    win.report_step_progress_bar.setObjectName("reportStepProgress")

    info_row = QHBoxLayout()
    info_row.setSpacing(10)

    win.report_phase_label = QLabel("Waiting to start")
    win.report_phase_label.setObjectName("reportPhaseLabel")
    win.report_phase_label.setStyleSheet("font-weight: 900; color: #555;")

    win.report_pct_label = QLabel("0 %")
    win.report_pct_label.setFixedWidth(72)
    win.report_pct_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
    win.report_pct_label.setStyleSheet(
        "color: #1DA1F2; font-weight: 900; font-size: 12px; background: transparent;"
    )

    win.report_status_label = QLabel("Ready to generate.")
    win.report_status_label.setObjectName("panelSubtitle")
    win.report_status_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

    info_row.addWidget(win.report_phase_label, 1)
    info_row.addWidget(win.report_pct_label, 0, Qt.AlignLeft)
    info_row.addWidget(win.report_status_label, 0, Qt.AlignRight)
    prog_layout.addLayout(info_row)

    # Step list scroll area
    step_scroll = QScrollArea()
    step_scroll.setWidgetResizable(True)
    step_scroll.setFrameShape(QFrame.NoFrame)
    step_scroll.setMinimumHeight(280)
    step_inner = QWidget()
    step_inner_layout = QVBoxLayout(step_inner)
    step_inner_layout.setContentsMargins(0, 0, 0, 0)
    step_inner_layout.setSpacing(0)
    win.report_step_widget = StepStatusWidget()
    step_inner_layout.addWidget(win.report_step_widget)
    step_inner_layout.addStretch()
    step_scroll.setWidget(step_inner)
    prog_layout.addWidget(step_scroll, 1)

    win.report_log_box = QTextEdit()
    win.report_log_box.setReadOnly(True)
    win.report_log_box.setFixedHeight(130)
    win.report_log_box.setPlaceholderText("Extraction and comparison log will appear here …")
    win.report_log_box.setLineWrapMode(QTextEdit.NoWrap)
    win.report_log_panel = CollapsiblePanel("Detailed Log", win.report_log_box, expanded=False)
    prog_layout.addWidget(win.report_log_panel)

    layout.addWidget(progress_card, 1)

    scroll = win.make_scroll_page(page)
    win.stack.addWidget(scroll)
    win.pages["report"] = scroll