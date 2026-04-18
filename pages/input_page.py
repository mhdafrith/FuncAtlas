"""pages/input_page.py — Input selection page (Reference Bases vs Consolidated DB)."""
from PySide6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QFrame
from ui.widgets import SectionTitle, PremiumCard


def create_input_page(win):
    page = QWidget()
    layout = QVBoxLayout(page)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(12)

    right = QFrame()
    right.setObjectName("pageCard")
    right_layout = QVBoxLayout(right)
    right_layout.setContentsMargins(20, 20, 20, 20)
    right_layout.setSpacing(12)
    right_layout.addWidget(SectionTitle("Input Selection Overview", "Pick the flow that matches your data source."))

    row = QHBoxLayout()
    row.setSpacing(14)
    accent = win.accent_color.name()

    win.input_reference_card = win.register_accent_card(PremiumCard(
        "Reference Bases Mode",
        "Use this when you have one target base folder and optional reference base folders.",
        win.icons_white.icon("folder", 34), accent,
        "Go to Reference Form", lambda: win.show_page("reference")
    ))
    win.input_consolidated_card = win.register_accent_card(PremiumCard(
        "Consolidated DB Mode",
        "Use this when your source comes from one consolidated database Excel workflow.",
        win.icons_white.icon("database", 34), accent,
        "Go to Consolidated Form", lambda: win.show_page("consolidated")
    ))
    row.addWidget(win.input_reference_card)
    row.addWidget(win.input_consolidated_card)
    win.input_cards = [win.input_reference_card, win.input_consolidated_card]

    for btn, dest in [
        (win.input_reference_card.action_btn,    "reference"),
        (win.input_consolidated_card.action_btn, "consolidated"),
    ]:
        win.wire_animated_navigation(btn, dest)

    right_layout.addLayout(row)
    right_layout.addStretch()
    layout.addWidget(right)
    layout.addStretch()

    scroll = win.make_scroll_page(page)
    win.stack.addWidget(scroll)
    win.pages["input"] = scroll