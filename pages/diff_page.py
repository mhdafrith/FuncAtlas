"""pages/diff_page.py — Diff Workspace page."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFrame, QLabel,
    QSplitter, QLineEdit, QTreeWidget, QPushButton, QTabWidget
)


def create_diff_page(win):
    page = QWidget()
    layout = QVBoxLayout(page)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(0)

    # ── toolbar ───────────────────────────────────────────────────────────────
    toolbar = QWidget()
    toolbar.setFixedHeight(52)
    tb_layout = QHBoxLayout(toolbar)
    tb_layout.setContentsMargins(8, 8, 8, 8)
    tb_layout.setSpacing(8)

    filter_icon_lbl = QLabel("🔍 Filter:")
    filter_icon_lbl.setStyleSheet(
        "font-size: 12px; font-weight: 700; color: #555; padding: 0 6px;"
    )
    tb_layout.addWidget(filter_icon_lbl)
    win._diff_filter_icon_lbl = filter_icon_lbl

    # track current active filter; None = all shown
    win._diff_active_filter = None

    FILTER_DEFS = [
        ("added",    "#2E7D32", "#FFFFFF", "  ✚  Added  "),
        ("deleted",  "#F9A825", "#000000", "  ✖  Deleted  "),
        ("modified", "#1565C0", "#FFFFFF", "  ✎  Modified  "),
    ]
    win._diff_filter_btns = {}

    def _style_btn(btn, active, greyed, color, fg):
        """Apply explicit inline style that overrides global QPushButton stylesheet."""
        obj = btn.objectName()
        if greyed:
            btn.setStyleSheet(
                f"QPushButton#{obj} {{"
                f"  background: #cccccc !important; color: #888888 !important;"
                f"  border: 2px solid #bbbbbb !important; border-radius: 8px !important;"
                f"  font-size: 12px; font-weight: 700; padding: 2px 14px;"
                f"  min-height: 30px; min-width: 100px;"
                f"}}"
                f"QPushButton#{obj}:hover {{ background: #bbbbbb !important; }}"
            )
        elif active:
            btn.setStyleSheet(
                f"QPushButton#{obj} {{"
                f"  background: {color} !important; color: {fg} !important;"
                f"  border: 2px solid {color} !important; border-radius: 8px !important;"
                f"  font-size: 12px; font-weight: 700; padding: 2px 14px;"
                f"  min-height: 30px; min-width: 100px;"
                f"}}"
                f"QPushButton#{obj}:hover {{ opacity: 0.9; }}"
            )
        else:
            btn.setStyleSheet(
                f"QPushButton#{obj} {{"
                f"  background: transparent !important; color: {color} !important;"
                f"  border: 2px solid {color} !important; border-radius: 8px !important;"
                f"  font-size: 12px; font-weight: 700; padding: 2px 14px;"
                f"  min-height: 30px; min-width: 100px;"
                f"}}"
                f"QPushButton#{obj}:hover {{ background: {color}30 !important; }}"
            )

    def _on_filter_clicked(clicked_key):
        # Clear the diff tabs so view refreshes when switching filters
        win.diff_tabs.clear()

        if win._diff_active_filter == clicked_key:
            # Toggle off → show all
            win._diff_active_filter = None
            for k, btn in win._diff_filter_btns.items():
                c  = btn.property("activeColor")
                fg = btn.property("activeFg")
                btn.setChecked(False)
                _style_btn(btn, active=False, greyed=False, color=c, fg=fg)
        else:
            win._diff_active_filter = clicked_key
            for k, btn in win._diff_filter_btns.items():
                c  = btn.property("activeColor")
                fg = btn.property("activeFg")
                if k == clicked_key:
                    btn.setChecked(True)
                    _style_btn(btn, active=True, greyed=False, color=c, fg=fg)
                else:
                    btn.setChecked(False)
                    _style_btn(btn, active=False, greyed=True, color=c, fg=fg)
        win._apply_diff_filter()

    for key, color, fg, label in FILTER_DEFS:
        btn = QPushButton(label)
        btn.setObjectName(f"diffFilterBtn_{key}")
        btn.setCheckable(True)
        btn.setFixedHeight(32)
        btn.setMinimumWidth(110)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setProperty("filterKey",   key)
        btn.setProperty("activeColor", color)
        btn.setProperty("activeFg",    fg)
        _style_btn(btn, active=False, greyed=False, color=color, fg=fg)
        btn.clicked.connect(lambda checked=False, k=key: _on_filter_clicked(k))
        win._diff_filter_btns[key] = btn
        tb_layout.addWidget(btn)

    # Store _style_btn so _clear_diff can use it to reset button styles
    win._diff_style_btn = _style_btn

    tb_layout.addStretch()

    # ── Fullscreen info bar: filename + funcname + prev/next (hidden normally) ─
    win._diff_fs_info_bar = QWidget()
    fs_info_layout = QHBoxLayout(win._diff_fs_info_bar)
    fs_info_layout.setContentsMargins(0, 0, 0, 0)
    fs_info_layout.setSpacing(6)

    win._diff_fs_file_label = QLabel("")
    win._diff_fs_file_label.setStyleSheet(
        "font-size: 12px; font-weight: 700; color: #1565C0; padding: 0 4px;"
    )
    win._diff_fs_func_label = QLabel("")
    win._diff_fs_func_label.setStyleSheet(
        "font-size: 12px; font-weight: 700; color: #2E7D32; padding: 0 4px;"
    )
    _sep = QLabel("›")
    _sep.setStyleSheet("color: #888; font-size: 13px;")

    win._diff_prev_btn = QPushButton("◀  Prev")
    win._diff_prev_btn.setObjectName("diffPrevBtn")
    win._diff_prev_btn.setFixedHeight(32)
    win._diff_prev_btn.setMinimumWidth(80)
    win._diff_prev_btn.setCursor(Qt.PointingHandCursor)
    win._diff_prev_btn.setStyleSheet(
        "QPushButton#diffPrevBtn { background: #455A64; color: #FFFFFF; border: 2px solid #263238;"
        " border-radius: 8px; font-size: 12px; font-weight: 800; padding: 0px 10px; }"
        "QPushButton#diffPrevBtn:hover { background: #546E7A; }"
        "QPushButton#diffPrevBtn:pressed { background: #263238; }"
        "QPushButton#diffPrevBtn:disabled { background: #B0BEC5; color: #ECEFF1; border-color: #90A4AE; }"
    )

    win._diff_next_btn = QPushButton("Next  ▶")
    win._diff_next_btn.setObjectName("diffNextBtn")
    win._diff_next_btn.setFixedHeight(32)
    win._diff_next_btn.setMinimumWidth(80)
    win._diff_next_btn.setCursor(Qt.PointingHandCursor)
    win._diff_next_btn.setStyleSheet(
        "QPushButton#diffNextBtn { background: #455A64; color: #FFFFFF; border: 2px solid #263238;"
        " border-radius: 8px; font-size: 12px; font-weight: 800; padding: 0px 10px; }"
        "QPushButton#diffNextBtn:hover { background: #546E7A; }"
        "QPushButton#diffNextBtn:pressed { background: #263238; }"
        "QPushButton#diffNextBtn:disabled { background: #B0BEC5; color: #ECEFF1; border-color: #90A4AE; }"
    )

    fs_info_layout.addWidget(win._diff_fs_file_label)
    fs_info_layout.addWidget(_sep)
    fs_info_layout.addWidget(win._diff_fs_func_label)
    fs_info_layout.addStretch()
    fs_info_layout.addWidget(win._diff_prev_btn)
    fs_info_layout.addWidget(win._diff_next_btn)

    win._diff_fs_info_bar.setVisible(False)
    tb_layout.addWidget(win._diff_fs_info_bar)

    # ── Fullscreen button ─────────────────────────────────────────────────────
    win._diff_fullscreen = False
    win._diff_fs_btn = QPushButton("⛶  Fullscreen")
    win._diff_fs_btn.setObjectName("diffFsBtn")
    win._diff_fs_btn.setFixedSize(148, 34)
    win._diff_fs_btn.setCursor(Qt.PointingHandCursor)
    win._diff_fs_btn.setStyleSheet(
        "QPushButton#diffFsBtn { background: #1565C0; color: #FFFFFF; border: 2px solid #0D47A1;"
        " border-radius: 8px; font-size: 12px; font-weight: 800; padding: 0px 10px; }"
        "QPushButton#diffFsBtn:hover { background: #1976D2; }"
        "QPushButton#diffFsBtn:pressed { background: #0D47A1; }"
    )
    win._diff_fs_btn.clicked.connect(win._toggle_diff_fullscreen)
    tb_layout.addWidget(win._diff_fs_btn)

    # ── Clear button (hidden in fullscreen) ───────────────────────────────────
    win._diff_clear_btn = QPushButton("🗑  Clear")
    win._diff_clear_btn.setObjectName("diffClearBtn")
    win._diff_clear_btn.setFixedSize(110, 34)
    win._diff_clear_btn.setCursor(Qt.PointingHandCursor)
    win._diff_clear_btn.setStyleSheet(
        "QPushButton#diffClearBtn { background: #B71C1C; color: #FFFFFF; border: 2px solid #7F0000;"
        " border-radius: 8px; font-size: 12px; font-weight: 800; padding: 0px 10px; }"
        "QPushButton#diffClearBtn:hover { background: #C62828; }"
        "QPushButton#diffClearBtn:pressed { background: #7F0000; }"
    )
    win._diff_clear_btn.clicked.connect(win._clear_diff)
    tb_layout.addSpacing(6)
    tb_layout.addWidget(win._diff_clear_btn)
    layout.addWidget(toolbar)

    # ── splitter: left tree | right tabs ─────────────────────────────────────
    win._diff_splitter = QSplitter(Qt.Horizontal)

    win._diff_left_panel = QFrame()
    win._diff_left_panel.setObjectName("pageCard")
    left_layout = QVBoxLayout(win._diff_left_panel)
    left_layout.setContentsMargins(8, 8, 8, 8)
    left_layout.setSpacing(6)

    win.diff_search_box = QLineEdit()
    win.diff_search_box.setPlaceholderText("Search function name...")
    win.diff_search_box.setMinimumHeight(36)
    win.diff_search_box.textChanged.connect(
        lambda text: win.filter_tree_items(text, win.diff_tree)
    )

    win.diff_tree = QTreeWidget()
    win.diff_tree.setHeaderLabel("Files and Functions")
    win.diff_tree.itemClicked.connect(win.on_diff_item_clicked)

    left_layout.addWidget(win.diff_search_box)
    left_layout.addWidget(win.diff_tree)

    win.diff_tabs = QTabWidget()
    win._diff_splitter.addWidget(win._diff_left_panel)
    win._diff_splitter.addWidget(win.diff_tabs)
    win._diff_splitter.setSizes([300, 900])
    layout.addWidget(win._diff_splitter, 1)

    win.stack.addWidget(page)
    win.pages["diff"] = page
