"""pages/consolidated_page.py — Consolidated DB input form."""
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFrame, QLabel,
    QPushButton, QButtonGroup, QStackedWidget
)
from PySide6.QtCore import Qt
from ui.widgets import (
    SectionTitle, IconTextButton,
    TxtMultiFileField, ExcelFileField, OutputLinkField,
    TargetFolderInputField
)
from ui.auto_detect_field import AutoDetectColumnField


def create_consolidated_page(win):
    page = QWidget()
    layout = QVBoxLayout(page)
    layout.setContentsMargins(0, 0, 0, 0)

    card = QFrame()
    card.setObjectName("pageCard")
    card_layout = QVBoxLayout(card)
    card_layout.setContentsMargins(22, 18, 22, 18)
    card_layout.setSpacing(14)
    card_layout.addWidget(SectionTitle(
        "Consolidated DB Input",
        "Upload multiple function list files, one Excel with multiple sheets, "
        "auto-detect Func/Base columns, and auto-generate output."
    ))

    # ── Source toggle: Function List Files <-> Target Folder ──────────────────
    toggle_row = QHBoxLayout()
    toggle_row.setSpacing(0)
    toggle_row.setContentsMargins(0, 0, 0, 0)

    win.con_toggle_excel_btn = QPushButton("Function List Files")
    win.con_toggle_excel_btn.setObjectName("toggleLeft")
    win.con_toggle_excel_btn.setCheckable(True)
    win.con_toggle_excel_btn.setChecked(True)
    win.con_toggle_excel_btn.setFixedHeight(34)

    win.con_toggle_folder_btn = QPushButton("Target Folder")
    win.con_toggle_folder_btn.setObjectName("toggleRight")
    win.con_toggle_folder_btn.setCheckable(True)
    win.con_toggle_folder_btn.setFixedHeight(34)

    win._con_toggle_group = QButtonGroup(page)
    win._con_toggle_group.setExclusive(True)
    win._con_toggle_group.addButton(win.con_toggle_excel_btn)
    win._con_toggle_group.addButton(win.con_toggle_folder_btn)

    toggle_row.addWidget(win.con_toggle_excel_btn)
    toggle_row.addWidget(win.con_toggle_folder_btn)
    toggle_row.addStretch()
    card_layout.addLayout(toggle_row)

    # ── Context indicator banner ──────────────────────────────────────────────
    win._con_mode_banner = QLabel(
        "  Mode: Function List Files  —  Upload Excel files containing function names")
    win._con_mode_banner.setStyleSheet(
        "background: #E3F2FD; color: #0D47A1; border: 1px solid #90CAF9;"
        " border-radius: 8px; padding: 6px 12px; font-weight: 700; font-size: 12px;"
    )
    win._con_mode_banner.setWordWrap(True)
    card_layout.addWidget(win._con_mode_banner)

    # ── FIX 1: Source stack placed FIRST (above Consolidated DB field) ────────
    win.con_source_stack = QStackedWidget()

    # Sheet 0 — xlsx multi-file picker
    sheet0 = QWidget()
    s0_layout = QVBoxLayout(sheet0)
    s0_layout.setContentsMargins(0, 0, 0, 0)
    win.con_function_field = TxtMultiFileField(
        "Function List Files *.xlsx", "Upload Excel",
        win.icons_white.icon("excel", 15)
    )
    win.con_function_field.setObjectName("funcListField")
    s0_layout.addWidget(win.con_function_field)
    win.con_source_stack.addWidget(sheet0)

    # Sheet 1 — target folder picker
    sheet1 = QWidget()
    s1_layout = QVBoxLayout(sheet1)
    s1_layout.setContentsMargins(0, 0, 0, 0)
    win.con_folder_field = TargetFolderInputField(
        "Target Folder  (scans all sub-folders automatically)",
        win.icons_white.icon("folder", 15)
    )
    win.con_folder_field.setObjectName("funcListField")
    s1_layout.addWidget(win.con_folder_field)
    win.con_source_stack.addWidget(sheet1)

    # Source stack ABOVE consolidated DB (Fix 1)
    card_layout.addWidget(win.con_source_stack)

    # ── FIX 1: Consolidated DB Excel field — same style as function list ───────
    win.con_db_excel_field = ExcelFileField(
        "Consolidated DB Excel File *", "Upload Excel",
        win.icons_white.icon("excel", 15)
    )
    # Use same objectName so it gets the same green-border styling as funcListField
    win.con_db_excel_field.setObjectName("funcListField")
    card_layout.addWidget(win.con_db_excel_field)

    # ── FIX 2 & 3: Column auto-detect fields ──────────────────────────────────
    # Left  (detect_kind="function")    → reads from Function List excel files
    # Right (detect_kind="db_function") → reads from Consolidated DB excel
    columns_row = QHBoxLayout()
    columns_row.setSpacing(14)

    win.con_func_col_field = AutoDetectColumnField(
        "Function Name Cell - Function list Excel",
        "e.g. D11 — column with function names",
        win.icons_white.icon("column", 15), "function", win
    )
    win.con_func_col_field.setObjectName("funcColField")

    # FIX 2: new kind "db_function" — auto-detect reads from con_db_excel_field
    win.con_base_col_field = AutoDetectColumnField(
        "Function Name Cell - Consolidated DB",
        "e.g. E11 — column with function names in DB",
        win.icons_white.icon("column", 15), "db_function", win
    )
    win.con_base_col_field.setObjectName("baseColField")

    columns_row.addWidget(win.con_func_col_field, 1)
    columns_row.addWidget(win.con_base_col_field, 1)
    card_layout.addLayout(columns_row)

    # ── Wire toggle buttons ───────────────────────────────────────────────────
    def _on_source_toggle():
        use_folder = win.con_toggle_folder_btn.isChecked()
        win.con_source_stack.setCurrentIndex(1 if use_folder else 0)

        if use_folder:
            win._con_mode_banner.setText(
                "  Mode: Target Folder  —  Functions are extracted directly from source files in the folder"
            )
            win._con_mode_banner.setStyleSheet(
                "background: #E8F5E9; color: #1B5E20; border: 1px solid #A5D6A7;"
                " border-radius: 8px; padding: 6px 12px; font-weight: 700; font-size: 12px;"
            )
            win.con_function_field.clear_selection()
        else:
            win._con_mode_banner.setText(
                "  Mode: Function List Files  —  Upload Excel files containing function names"
            )
            win._con_mode_banner.setStyleSheet(
                "background: #E3F2FD; color: #0D47A1; border: 1px solid #90CAF9;"
                " border-radius: 8px; padding: 6px 12px; font-weight: 700; font-size: 12px;"
            )
            win.con_folder_field.clear_selection()

        # In folder mode the function-list column is not needed
        win.con_func_col_field.setEnabled(not use_folder)
        win.con_func_col_field.setToolTip(
            "Auto-detected from extracted folder data" if use_folder else ""
        )

    win.con_toggle_excel_btn.toggled.connect(lambda _: _on_source_toggle())
    win.con_toggle_folder_btn.toggled.connect(lambda _: _on_source_toggle())

    # ── Output link ───────────────────────────────────────────────────────────
    win.con_output_link_field = OutputLinkField(
        "Output Excel Link", win.icons_white.icon("link", 15)
    )
    win.con_open_output_btn = win.con_output_link_field.open_btn
    card_layout.addWidget(win.con_output_link_field)

    # ── Bottom button row ─────────────────────────────────────────────────────
    btn_row = QHBoxLayout()
    btn_row.setSpacing(10)

    win.con_submit_btn = IconTextButton("Submit", win.icons_white.icon("submit", 15))
    win.con_submit_btn.setObjectName("smallPrimaryButton")
    win.con_submit_btn.setFixedSize(130, 44)

    win.con_back_btn = IconTextButton("Back", win.icons_white.icon("back", 15))
    win.con_back_btn.setObjectName("pickerButton")
    win.con_back_btn.setFixedSize(130, 44)

    win.con_clear_btn = IconTextButton("Clear", win.icons_white.icon("clear", 15))
    win.con_clear_btn.setObjectName("clearButtonRect")
    win.con_clear_btn.setFixedSize(130, 44)

    win.con_open_output_btn.setFixedSize(150, 44)

    win.con_submit_btn.clicked.connect(
        lambda: win.animate_button_click(win.con_submit_btn, win.submit_consolidated))
    win.con_back_btn.clicked.connect(
        lambda: win.animate_button_and_navigate(win.con_back_btn, "input"))
    win.con_clear_btn.clicked.connect(
        lambda: win.animate_button_click(win.con_clear_btn, win.clear_consolidated_form))

    btn_row.addWidget(win.con_submit_btn)
    btn_row.addWidget(win.con_back_btn)
    btn_row.addWidget(win.con_clear_btn)
    btn_row.addWidget(win.con_open_output_btn)
    btn_row.addStretch()
    card_layout.addLayout(btn_row)

    layout.addWidget(card)
    layout.addStretch()

    scroll = win.make_scroll_page(page)
    win.stack.addWidget(scroll)
    win.pages["consolidated"] = scroll
