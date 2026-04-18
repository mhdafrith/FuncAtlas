"""pages/view_page.py — Function Explorer page."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QFrame, QLabel, QHBoxLayout,
    QComboBox, QSplitter, QLineEdit, QTreeWidget, QTextEdit, QSizePolicy
)
from ui.widgets import SectionTitle


def create_view_page(win):
    page = QWidget()
    page_layout = QVBoxLayout(page)
    page_layout.setContentsMargins(0, 0, 0, 0)
    page_layout.setSpacing(10)

    # ── source selection bar ─────────────────────────────────────────────────
    source_bar = QFrame()
    source_bar.setObjectName("pageCard")
    source_bar.setFixedHeight(68)
    sb_layout = QHBoxLayout(source_bar)
    sb_layout.setContentsMargins(14, 10, 14, 10)
    sb_layout.setSpacing(10)

    source_label = QLabel("Source Selection")
    source_label.setObjectName("fieldLabel")
    source_label.setFixedWidth(120)

    win.source_combo = QComboBox()
    win.source_combo.setMinimumHeight(36)
    win.source_combo.currentIndexChanged.connect(win.on_source_changed)
    win.source_combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

    # Mode chip — only visible when Function List (.xlsx) is loaded (point 4)
    win.view_mode_chip = QLabel()
    win.view_mode_chip.setObjectName("modeBadge")
    win.view_mode_chip.setFixedHeight(30)
    win.view_mode_chip.setVisible(False)
    win.view_mode_chip.setStyleSheet(
        "QLabel#modeBadge {"
        "  background: #1565C0; color: #fff;"
        "  border-radius: 6px; font-size: 11px; font-weight: 700;"
        "  padding: 2px 12px;"
        "}"
    )

    sb_layout.addWidget(source_label)
    sb_layout.addWidget(win.source_combo, 1)
    sb_layout.addWidget(win.view_mode_chip)
    page_layout.addWidget(source_bar, 0, Qt.AlignTop)

    # ── splitter: left tree | right preview ──────────────────────────────────
    win.view_splitter = QSplitter(Qt.Horizontal)
    win.view_splitter.setChildrenCollapsible(False)
    win.view_splitter.setHandleWidth(12)

    # Left panel
    left = QFrame()
    left.setObjectName("pageCard")
    left.setMinimumWidth(360)
    left_layout = QVBoxLayout(left)
    left_layout.setContentsMargins(14, 14, 14, 14)
    left_layout.setSpacing(8)
    left_layout.addWidget(SectionTitle("Loaded Files & Functions",
                                       "Drag the center divider left or right to adjust panel width."))
    win.search_box = QLineEdit()
    win.search_box.setPlaceholderText("Search function name...")
    win.search_box.setMinimumHeight(36)
    win.search_box.textChanged.connect(win.filter_tree_items)

    win.tree = QTreeWidget()
    win.tree.setHeaderLabel("Files and Functions")
    win.tree.itemClicked.connect(win.on_tree_item_clicked)
    win.tree.setMinimumHeight(430)
    win.tree.setIndentation(16)
    win.tree.header().setStretchLastSection(True)
    from PySide6.QtCore import Qt as _Qt
    win.tree.setHorizontalScrollBarPolicy(_Qt.ScrollBarAsNeeded)

    left_layout.addWidget(win.search_box)
    left_layout.addWidget(win.tree, 1)

    # Right panel
    right = QFrame()
    right.setObjectName("pageCard")
    right.setMinimumWidth(360)
    right_layout = QVBoxLayout(right)
    right_layout.setContentsMargins(10, 10, 10, 10)
    right_layout.setSpacing(8)

    top = QFrame()
    top.setObjectName("softPanel")
    top.setFixedHeight(78)
    top_layout = QVBoxLayout(top)
    top_layout.setContentsMargins(14, 10, 14, 10)
    top_layout.setSpacing(2)

    win.view_title = QLabel("Select a function")
    win.view_title.setObjectName("panelTitle")
    win.view_meta = QLabel("Function preview panel")
    win.view_meta.setObjectName("panelSubtitle")
    win.view_meta.setWordWrap(True)
    top_layout.addWidget(win.view_title)
    top_layout.addWidget(win.view_meta)

    win.view_text = QTextEdit()
    win.view_text.setReadOnly(True)
    win.view_text.setLineWrapMode(QTextEdit.NoWrap)
    win.view_text.setHorizontalScrollBarPolicy(_Qt.ScrollBarAsNeeded)
    win.view_text.setText(
        "Load target/reference inputs first.\n\n"
        "Then:\n1. choose source root from the dropdown\n"
        "2. files and functions will appear on the left\n"
        "3. drag the center divider left or right if you want more space\n"
        "4. click a function name to preview the real body here"
    )
    right_layout.addWidget(top)
    right_layout.addWidget(win.view_text, 1)

    win.view_splitter.addWidget(left)
    win.view_splitter.addWidget(right)
    win.view_splitter.setStretchFactor(0, 1)
    win.view_splitter.setStretchFactor(1, 1)
    win.view_splitter.setSizes([560, 780])
    page_layout.addWidget(win.view_splitter, 1)

    scroll = win.make_scroll_page(page)
    win.stack.addWidget(scroll)
    win.pages["view"] = scroll