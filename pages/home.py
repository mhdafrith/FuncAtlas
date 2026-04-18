"""pages/home.py — Home / Dashboard page builder."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFrame, QLabel, QSizePolicy
)
from ui.widgets import SectionTitle, PremiumCard


def create_home_page(win):
    page = QWidget()
    layout = QVBoxLayout(page)
    layout.setContentsMargins(4, 4, 4, 4)
    layout.setSpacing(12)

    # ── Hero card — centered, no image, no badge ─────────────────────────────
    hero = QFrame()
    hero.setObjectName("heroCard")
    hero.setMinimumHeight(220)
    hero.setMaximumHeight(260)
    hero.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
    hero_layout = QVBoxLayout(hero)
    hero_layout.setContentsMargins(32, 28, 32, 28)
    hero_layout.setSpacing(10)
    hero_layout.setAlignment(Qt.AlignCenter)

    hero_kicker = QLabel("WELCOME TO")
    hero_kicker.setObjectName("heroKicker")
    hero_kicker.setAlignment(Qt.AlignCenter)
    hero_kicker.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

    hero_title = QLabel("FuncAtlas")
    hero_title.setObjectName("heroTitle")
    hero_title.setAlignment(Qt.AlignCenter)
    hero_title.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

    hero_subtitle = QLabel(
        "Explore your function network across source folders, reference "
        "bases, consolidated DB flows, and report generation in one "
        "premium workspace."
    )
    hero_subtitle.setObjectName("heroSubtitle")
    hero_subtitle.setWordWrap(True)
    hero_subtitle.setAlignment(Qt.AlignCenter)
    hero_subtitle.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

    hero_layout.addWidget(hero_kicker)
    hero_layout.addWidget(hero_title)
    hero_layout.addWidget(hero_subtitle)

    layout.addWidget(hero)

    # ── Quick Actions ─────────────────────────────────────────────────────────
    layout.addWidget(SectionTitle("Quick Actions", "Move directly into the core workflow areas."))

    cards_frame = QFrame()
    cards_frame.setObjectName("quickActionsFrame")
    cards_frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
    cards_layout = QHBoxLayout(cards_frame)
    cards_layout.setContentsMargins(20, 20, 20, 20)  # Padding for shadow
    cards_layout.setSpacing(18)  # Space between cards for shadow

    accent = win.accent_color.name()

    win.home_ref_card = win.register_accent_card(PremiumCard(
        "Reference Base Mode",
        "Load target base folder, reference folders, and function list files.",
        win.icons_white.icon("folder", 34), accent, "Input", lambda: win.show_page("reference")
    ))
    win.home_con_card = win.register_accent_card(PremiumCard(
        "Consolidated DB Mode",
        "Upload function list files, consolidated Excel, and auto-generate output.",
        win.icons_white.icon("excel", 34), accent, "Input", lambda: win.show_page("consolidated")
    ))

    for card in (win.home_ref_card, win.home_con_card):
        card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        card.setMinimumHeight(200)
        card.setMaximumHeight(16777215)  # remove height cap so card fills frame

    cards_layout.addWidget(win.home_ref_card, 1)
    cards_layout.addWidget(win.home_con_card, 1)

    win.home_cards = [win.home_ref_card, win.home_con_card]

    for btn, dest in [
        (win.home_ref_card.action_btn, "reference"),
        (win.home_con_card.action_btn, "consolidated"),
    ]:
        win.wire_animated_navigation(btn, dest)

    layout.addWidget(cards_frame)
    layout.addStretch()
    scroll = win.make_scroll_page(page)
    win.stack.addWidget(scroll)
    win.pages["home"] = scroll