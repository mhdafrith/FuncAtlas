"""pages/settings_page.py — Appearance settings page."""
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFrame, QLabel, QComboBox
)
from ui.widgets import SectionTitle, IconTextButton


def create_settings_page(win):
    page = QWidget()
    layout = QVBoxLayout(page)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(12)

    card = QFrame()
    card.setObjectName("pageCard")
    card_layout = QVBoxLayout(card)
    card_layout.setContentsMargins(20, 18, 20, 18)
    card_layout.setSpacing(14)
    card_layout.addWidget(SectionTitle(
        "Appearance Settings",
        "Control dark/light theme, accent, font, and theme reset from one place."
    ))

    # Theme selector
    theme_row = QHBoxLayout()
    theme_row.setSpacing(10)
    theme_label = QLabel("Theme Mode")
    theme_label.setObjectName("fieldLabel")
    win.settings_theme_combo = QComboBox()
    win.settings_theme_combo.addItems(["Dark", "Light"])
    win.settings_theme_combo.setMinimumHeight(36)
    win.settings_theme_combo.currentIndexChanged.connect(win.on_theme_combo_changed)
    theme_row.addWidget(theme_label)
    theme_row.addWidget(win.settings_theme_combo)
    theme_row.addStretch()

    # Action buttons
    btn_row = QHBoxLayout()
    btn_row.setSpacing(10)
    win.settings_accent_btn = IconTextButton("Change Accent Color", win.icons.icon("palette", 15))
    win.settings_accent_btn.setObjectName("ghostButton")
    win.settings_accent_btn.clicked.connect(win.pick_accent_color)

    win.settings_font_btn = IconTextButton("Change Font", win.icons.icon("font", 15))
    win.settings_font_btn.setObjectName("ghostButton")
    win.settings_font_btn.clicked.connect(win.pick_font)

    win.settings_reset_btn = IconTextButton("Reset Theme", win.icons.icon("reset", 15))
    win.settings_reset_btn.setObjectName("ghostButton")
    win.settings_reset_btn.clicked.connect(win.reset_theme)

    btn_row.addWidget(win.settings_accent_btn)
    btn_row.addWidget(win.settings_font_btn)
    btn_row.addWidget(win.settings_reset_btn)
    btn_row.addStretch()

    win.theme_note_label = QLabel("")
    win.theme_note_label.setObjectName("panelSubtitle")
    win.theme_note_label.setWordWrap(True)

    card_layout.addLayout(theme_row)
    card_layout.addLayout(btn_row)
    card_layout.addWidget(win.theme_note_label)
    card_layout.addStretch()

    layout.addWidget(card)
    layout.addStretch()

    scroll = win.make_scroll_page(page)
    win.stack.addWidget(scroll)
    win.pages["settings"] = scroll