"""pages/reference_page.py — Reference Bases input form."""
from PySide6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QFrame
from ui.widgets import (
    SectionTitle, IconTextButton,
    FolderField, XlsxMultiFileField
)


def create_reference_page(win):
    page = QWidget()
    layout = QVBoxLayout(page)
    layout.setContentsMargins(0, 0, 0, 0)

    card = QFrame()
    card.setObjectName("pageCard")
    card_layout = QVBoxLayout(card)
    card_layout.setContentsMargins(22, 18, 22, 18)
    card_layout.setSpacing(12)
    card_layout.addWidget(SectionTitle(
        "Reference Bases Input",
        "Target Base and Reference Bases are mandatory. Function List supports *.xlsx only."
    ))

    # ── Target Base: single folder — native OS dialog ────────────────────────
    win.ref_target_field = FolderField(
        "Target Base Folder *", "Upload Folder",
        win.icons_white.icon("folder", 15),
        multi=False, show_clear_btn=True
    )

    # ── Reference Bases: multi-folder — custom list + native add ────────────
    win.ref_bases_field = FolderField(
        "Reference Bases Folders *", "Add Folder",
        win.icons_white.icon("folder", 15),
        multi=True, multi_line_display=True
    )

    # ── Function list: native multi-file picker ──────────────────────────────
    win.ref_function_field = XlsxMultiFileField(
        "Function List Files *.xlsx", "Choose Files",
        win.icons_white.icon("document", 15)
    )

    card_layout.addWidget(win.ref_target_field)
    card_layout.addWidget(win.ref_bases_field)
    card_layout.addWidget(win.ref_function_field)

    # Point 6: when inputs change (re-upload or clear), auto-clear the diff
    def _on_ref_input_changed():
        """Called whenever target/reference/function inputs change — clears diff."""
        if hasattr(win, "_clear_diff"):
            win._clear_diff()

    win.ref_target_field.selectionChanged.connect(_on_ref_input_changed)
    win.ref_bases_field.selectionChanged.connect(_on_ref_input_changed)
    win.ref_function_field.selectionChanged.connect(_on_ref_input_changed)

    btn_row = QHBoxLayout()
    btn_row.setSpacing(8)

    win.ref_submit_btn = IconTextButton("Submit", win.icons_white.icon("submit", 15))
    win.ref_submit_btn.setObjectName("smallPrimaryButton")
    win.ref_submit_btn.setFixedSize(124, 40)

    win.ref_back_btn = IconTextButton("Back", win.icons_white.icon("back", 15))
    win.ref_back_btn.setObjectName("pickerButton")
    win.ref_back_btn.setFixedSize(124, 40)

    win.ref_clear_btn = IconTextButton("Clear", win.icons_white.icon("clear", 15))
    win.ref_clear_btn.setObjectName("pickerButton")
    win.ref_clear_btn.setFixedSize(124, 40)

    win.ref_submit_btn.clicked.connect(lambda: win.animate_button_click(win.ref_submit_btn, win.submit_reference))
    win.ref_back_btn.clicked.connect(lambda: win.animate_button_and_navigate(win.ref_back_btn, "input"))
    win.ref_clear_btn.clicked.connect(lambda: win.animate_button_click(win.ref_clear_btn, win.clear_reference_form))

    btn_row.addWidget(win.ref_submit_btn)
    btn_row.addWidget(win.ref_back_btn)
    btn_row.addWidget(win.ref_clear_btn)
    btn_row.addStretch()
    card_layout.addLayout(btn_row)

    layout.addWidget(card)
    layout.addStretch()

    scroll = win.make_scroll_page(page)
    win.stack.addWidget(scroll)
    win.pages["reference"] = scroll