"""
ui/dialogs.py
─────────────
Standalone modal dialogs used throughout FuncAtlas.

Exports:
  HelpOverlayDialog     – numbered step-by-step help overlay
  CompletionPopupDialog – success summary popup with stat cards
"""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QFrame, QLabel, QPushButton,
    QVBoxLayout, QHBoxLayout, QScrollArea, QWidget
)


class HelpOverlayDialog(QDialog):
    """Numbered step-by-step help modal."""

    def __init__(self, parent, title: str, sections: list,
                 footer_text: str = "", tip_text: str = ""):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.setFixedSize(760, 620)
        self.setWindowFlag(Qt.WindowContextHelpButtonHint, False)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        shell = QFrame()
        shell.setObjectName('helpDialogShell')
        shell_layout = QVBoxLayout(shell)
        shell_layout.setContentsMargins(32, 28, 32, 28)
        shell_layout.setSpacing(18)

        # ── Header ────────────────────────────────────────────────────────────
        head = QHBoxLayout()
        head.setContentsMargins(0, 0, 0, 0)
        head.setSpacing(10)

        icon_lbl = QLabel("✅")
        icon_lbl.setObjectName("helpDialogIcon")
        icon_lbl.setFixedSize(36, 36)
        icon_lbl.setAlignment(Qt.AlignCenter)

        title_lbl = QLabel(title)
        title_lbl.setObjectName('helpDialogTitle')
        head.addWidget(icon_lbl)
        head.addWidget(title_lbl)
        head.addStretch()
        shell_layout.addLayout(head)

        # Thin divider
        divider = QFrame()
        divider.setFixedHeight(1)
        divider.setObjectName("helpDialogDivider")
        shell_layout.addWidget(divider)

        # ── Scrollable body ───────────────────────────────────────────────────
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(0, 4, 0, 4)
        body_layout.setSpacing(10)

        for index, section in enumerate(sections, start=1):
            card = QFrame()
            card.setObjectName('helpStepCard')
            card_layout = QHBoxLayout(card)
            card_layout.setContentsMargins(18, 16, 18, 16)
            card_layout.setSpacing(14)

            badge = QLabel(str(index))
            badge.setObjectName('helpStepBadge')
            badge.setAlignment(Qt.AlignCenter)
            badge.setFixedSize(34, 34)

            text_wrap = QVBoxLayout()
            text_wrap.setContentsMargins(0, 0, 0, 0)
            text_wrap.setSpacing(4)
            heading = QLabel(section.get('title', ''))
            heading.setObjectName('helpStepTitle')
            heading.setWordWrap(True)
            body_lbl = QLabel(section.get('body', ''))
            body_lbl.setObjectName('helpStepBody')
            body_lbl.setWordWrap(True)
            text_wrap.addWidget(heading)
            text_wrap.addWidget(body_lbl)
            card_layout.addWidget(badge, 0, Qt.AlignTop)
            card_layout.addLayout(text_wrap, 1)
            body_layout.addWidget(card)

        if footer_text:
            footer = QFrame()
            footer.setObjectName('helpFooterWarning')
            fl = QHBoxLayout(footer)
            fl.setContentsMargins(16, 14, 16, 14)
            fl.setSpacing(10)
            footer_lbl = QLabel(footer_text)
            footer_lbl.setObjectName('helpFooterText')
            footer_lbl.setWordWrap(True)
            fl.addWidget(footer_lbl)
            body_layout.addWidget(footer)

        if tip_text:
            tip = QFrame()
            tip.setObjectName('helpFooterTip')
            tl = QHBoxLayout(tip)
            tl.setContentsMargins(16, 14, 16, 14)
            tl.setSpacing(10)
            tip_lbl = QLabel(tip_text)
            tip_lbl.setObjectName('helpFooterTipText')
            tip_lbl.setWordWrap(True)
            tl.addWidget(tip_lbl)
            body_layout.addWidget(tip)

        body_layout.addStretch()
        scroll.setWidget(body)
        shell_layout.addWidget(scroll, 1)

        # ── Footer divider + Close button ────────────────────────────────────
        foot_divider = QFrame()
        foot_divider.setFixedHeight(1)
        foot_divider.setObjectName("helpDialogDivider")
        shell_layout.addWidget(foot_divider)

        close_btn = QPushButton("Close")
        close_btn.setObjectName("helpDialogCloseBtn")
        close_btn.setFixedSize(120, 42)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.clicked.connect(self.accept)
        shell_layout.addWidget(close_btn, 0, Qt.AlignRight)

        outer.addWidget(shell)

        self.setStyleSheet("""
            QDialog { background: rgba(2, 10, 18, 0.94); }
            QFrame#helpDialogShell {
                background: #0A1630;
                border: 1px solid #1C8DFF;
                border-radius: 28px;
            }
            QFrame#helpDialogDivider {
                background: #1B3560;
                border: none;
            }
            QLabel#helpDialogIcon {
                font-size: 20px;
                background: transparent;
            }
            QLabel#helpDialogTitle {
                color: #55A6FF;
                font-size: 20px;
                font-weight: 900;
                background: transparent;
            }
            QFrame#helpStepCard {
                background: #0D1C38;
                border: 1px solid #1B4D85;
                border-radius: 14px;
            }
            QLabel#helpStepBadge {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #78B8FF, stop:1 #3E87F5);
                color: white;
                border-radius: 10px;
                font-size: 16px;
                font-weight: 900;
            }
            QLabel#helpStepTitle {
                color: #F4F8FF;
                font-size: 15px;
                font-weight: 800;
            }
            QLabel#helpStepBody {
                color: #90A4BE;
                font-size: 13px;
                line-height: 1.5;
            }
            QFrame#helpFooterWarning {
                background: #372000;
                border-left: 4px solid #FFB21E;
                border-radius: 10px;
            }
            QFrame#helpFooterTip {
                background: #032C23;
                border-left: 4px solid #00E18D;
                border-radius: 10px;
            }
            QLabel#helpFooterText {
                color: #D8A24A;
                font-size: 13px;
                font-weight: 700;
            }
            QLabel#helpFooterTipText {
                color: #76E2BC;
                font-size: 13px;
                font-weight: 700;
            }
            QPushButton#helpDialogCloseBtn {
                background: #1A3A6E;
                color: #A8C8FF;
                border: 1px solid #2A5EA8;
                border-radius: 10px;
                font-size: 14px;
                font-weight: 700;
            }
            QPushButton#helpDialogCloseBtn:hover {
                background: #1C8DFF;
                color: #FFFFFF;
                border-color: #1C8DFF;
            }
        """)


class CompletionPopupDialog(QDialog):
    """Green completion summary popup with stat cards."""

    def __init__(self, parent, title: str, subtitle: str,
                 total=None, diff=None, no_diff=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.setFixedSize(660, 540)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        shell = QFrame()
        shell.setObjectName('doneDialogShell')
        shell_layout = QVBoxLayout(shell)
        shell_layout.setContentsMargins(34, 24, 34, 26)
        shell_layout.setSpacing(18)

        # Close button row
        top = QHBoxLayout()
        top.setContentsMargins(0, 0, 0, 0)
        top.addStretch()
        close_btn = QPushButton('×')
        close_btn.setObjectName('aboutDialogCloseButton')
        close_btn.setFixedSize(40, 40)
        close_btn.clicked.connect(self.accept)
        top.addWidget(close_btn)
        shell_layout.addLayout(top)

        # Check badge
        badge = QLabel('✓')
        badge.setObjectName('doneDialogBadge')
        badge.setAlignment(Qt.AlignCenter)
        badge.setFixedSize(84, 84)
        shell_layout.addWidget(badge, 0, Qt.AlignHCenter)

        # Title / subtitle
        title_lbl = QLabel(title)
        title_lbl.setObjectName('doneDialogTitle')
        title_lbl.setAlignment(Qt.AlignCenter)
        shell_layout.addWidget(title_lbl)

        subtitle_lbl = QLabel(subtitle)
        subtitle_lbl.setObjectName('doneDialogSubtitle')
        subtitle_lbl.setWordWrap(True)
        subtitle_lbl.setAlignment(Qt.AlignCenter)
        shell_layout.addWidget(subtitle_lbl)

        # Stats row
        stats_row = QHBoxLayout()
        stats_row.setSpacing(18)
        for value, caption, tone in [
            (total,   'TOTAL',   'neutral'),
            (diff,    'DIFF',    'warning'),
            (no_diff, 'NO DIFF', 'success'),
        ]:
            card = QFrame()
            card.setProperty('tone', tone)
            card.setObjectName('doneStatCard')
            cl = QVBoxLayout(card)
            cl.setContentsMargins(16, 16, 16, 14)
            cl.setSpacing(2)
            value_lbl = QLabel('—' if value is None else str(value))
            value_lbl.setObjectName('doneStatValue')
            value_lbl.setAlignment(Qt.AlignCenter)
            caption_lbl = QLabel(caption)
            caption_lbl.setObjectName('doneStatLabel')
            caption_lbl.setAlignment(Qt.AlignCenter)
            cl.addWidget(value_lbl)
            cl.addWidget(caption_lbl)
            stats_row.addWidget(card)
        shell_layout.addLayout(stats_row)

        # Done button
        done_btn = QPushButton('Done')
        done_btn.setObjectName('doneDialogButton')
        done_btn.setFixedSize(180, 56)
        done_btn.clicked.connect(self.accept)
        shell_layout.addWidget(done_btn, 0, Qt.AlignHCenter)
        outer.addWidget(shell)

        self.setStyleSheet("""
            QDialog { background: rgba(2, 10, 18, 0.94); }
            QFrame#doneDialogShell {
                background: #0A1630;
                border: 2px solid #00C86A;
                border-radius: 30px;
            }
            QLabel#doneDialogBadge {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #7BE39B, stop:1 #43C970);
                color: #FFF3FF;
                border-radius: 18px;
                font-size: 52px;
                font-weight: 1000;
            }
            QLabel#doneDialogTitle {
                color: #12C45D;
                font-size: 30px;
                font-weight: 1000;
            }
            QLabel#doneDialogSubtitle {
                color: #B4C1D4;
                font-size: 15px;
            }
            QFrame#doneStatCard {
                background: #0E1C39;
                border: 1px solid #1E4D88;
                border-radius: 18px;
                min-height: 112px;
            }
            QFrame#doneStatCard[tone="warning"] QLabel#doneStatValue { color: #FFBE30; }
            QFrame#doneStatCard[tone="success"] QLabel#doneStatValue { color: #00E3A0; }
            QLabel#doneStatValue {
                color: #D8E1F0;
                font-size: 28px;
                font-weight: 1000;
            }
            QLabel#doneStatLabel {
                color: #7E8EAA;
                font-size: 12px;
                font-weight: 900;
                letter-spacing: 1px;
            }
            QPushButton#doneDialogButton {
                background: #16A34A;
                color: white;
                border: 1px solid #16A34A;
                border-radius: 16px;
                font-size: 18px;
                font-weight: 900;
                padding: 8px 18px;
                text-align: center;
            }
            QPushButton#doneDialogButton:hover { background: #1AB154; }
            QPushButton#aboutDialogCloseButton {
                background: transparent;
                color: #8899AA;
                border: none;
                font-size: 26px;
                font-weight: 500;
            }
            QPushButton#aboutDialogCloseButton:hover { color: #FFFFFF; }
        """)