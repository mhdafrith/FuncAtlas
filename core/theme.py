"""
core/theme.py
─────────────
Theme colour palettes and the VectorIconFactory.
No page/widget logic here.
"""

from PySide6.QtCore import Qt, QRectF, QPointF
from PySide6.QtGui import QColor, QPainter, QPixmap, QIcon, QPen, QPainterPath


# ── Theme Manager ─────────────────────────────────────────────────────────────
class ThemeManager:
    THEMES = {
        "dark": {
            "name": "Dark",
            "bg_main": "#09111B",
            "bg_sidebar": "#081019",
            "bg_header": "#0E1823",
            "bg_card": "#111B27",
            "bg_soft": "#132131",
            "bg_input": "#0A1520",
            "text_primary": "#F7FAFC",
            "text_secondary": "#C2CEDA",
            "text_muted": "#8AA0B5",
            "border": "#223447",
            "border_strong": "#34506A",
            "accent": "#3BA8FF",
            "accent_hover": "#67BCFF",
            "success": "#22C55E",
            "warning": "#F59E0B",
            "danger": "#EF4444",
        },
        "light": {
            "name": "Light",
            "bg_main": "#F4F7FB",
            "bg_sidebar": "#EAF0F8",
            "bg_header": "#FFFFFF",
            "bg_card": "#FFFFFF",
            "bg_soft": "#F7FAFD",
            "bg_input": "#FFFFFF",
            "text_primary": "#0F172A",
            "text_secondary": "#475569",
            "text_muted": "#64748B",
            "border": "#D7E0EA",
            "border_strong": "#B8C5D6",
            "accent": "#2563EB",
            "accent_hover": "#3B82F6",
            "success": "#16A34A",
            "warning": "#D97706",
            "danger": "#DC2626",
        },
    }


# ── VectorIconFactory ────────────────────────────────────────────────────────
class VectorIconFactory:
    def __init__(self, color: QColor):
        self.color = color

    def icon(self, name: str, size: int = 24) -> QIcon:
        pixmap = QPixmap(size, size)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        pen = QPen(self.color, max(1.4, size / 14), Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)
        painter.setPen(pen)
        painter.setBrush(Qt.NoBrush)

        draw_fn = {
            "home": self.draw_home,
            "input": self.draw_upload,
            "view": self.draw_eye,
            "diff": self.draw_diff,
            "report": self.draw_report,
            "help": self.draw_help,
            "settings": self.draw_settings,
            "document": self.draw_document,
            "database": self.draw_database,
            "palette": self.draw_palette,
            "folder": self.draw_folder,
            "clear": self.draw_clear,
            "back": self.draw_back,
            "submit": self.draw_submit,
            "font": self.draw_font,
            "reset": self.draw_reset,
            "excel": self.draw_excel,
            "column": self.draw_column,
            "link": self.draw_link,
        }.get(name)
        if draw_fn:
            draw_fn(painter, size)
        else:
            painter.drawEllipse(QRectF(size * 0.2, size * 0.2, size * 0.6, size * 0.6))
        painter.end()
        return QIcon(pixmap)

    def draw_home(self, p, s):
        roof = QPainterPath()
        roof.moveTo(s * 0.18, s * 0.5)
        roof.lineTo(s * 0.5, s * 0.22)
        roof.lineTo(s * 0.82, s * 0.5)
        p.drawPath(roof)
        p.drawRoundedRect(QRectF(s * 0.28, s * 0.5, s * 0.44, s * 0.27), 2, 2)

    def draw_upload(self, p, s):
        p.drawLine(QPointF(s * 0.5, s * 0.18), QPointF(s * 0.5, s * 0.67))
        arrow = QPainterPath()
        arrow.moveTo(s * 0.34, s * 0.35)
        arrow.lineTo(s * 0.5, s * 0.18)
        arrow.lineTo(s * 0.66, s * 0.35)
        p.drawPath(arrow)
        p.drawLine(QPointF(s * 0.25, s * 0.78), QPointF(s * 0.75, s * 0.78))

    def draw_eye(self, p, s):
        path = QPainterPath()
        path.moveTo(s * 0.15, s * 0.5)
        path.cubicTo(s * 0.28, s * 0.25, s * 0.72, s * 0.25, s * 0.85, s * 0.5)
        path.cubicTo(s * 0.72, s * 0.75, s * 0.28, s * 0.75, s * 0.15, s * 0.5)
        p.drawPath(path)
        p.drawEllipse(QRectF(s * 0.4, s * 0.4, s * 0.2, s * 0.2))

    def draw_diff(self, p, s):
        p.drawLine(QPointF(s * 0.2, s * 0.35), QPointF(s * 0.8, s * 0.35))
        p.drawLine(QPointF(s * 0.2, s * 0.65), QPointF(s * 0.8, s * 0.65))
        left = QPainterPath()
        left.moveTo(s * 0.22, s * 0.35); left.lineTo(s * 0.34, s * 0.24)
        left.moveTo(s * 0.22, s * 0.35); left.lineTo(s * 0.34, s * 0.46)
        right = QPainterPath()
        right.moveTo(s * 0.78, s * 0.65); right.lineTo(s * 0.66, s * 0.54)
        right.moveTo(s * 0.78, s * 0.65); right.lineTo(s * 0.66, s * 0.76)
        p.drawPath(left); p.drawPath(right)

    def draw_report(self, p, s):
        p.drawRoundedRect(QRectF(s * 0.24, s * 0.18, s * 0.52, s * 0.64), 2, 2)
        p.drawLine(QPointF(s * 0.34, s * 0.62), QPointF(s * 0.34, s * 0.42))
        p.drawLine(QPointF(s * 0.5,  s * 0.62), QPointF(s * 0.5,  s * 0.32))
        p.drawLine(QPointF(s * 0.66, s * 0.62), QPointF(s * 0.66, s * 0.5))

    def draw_help(self, p, s):
        path = QPainterPath()
        path.moveTo(s * 0.36, s * 0.36)
        path.cubicTo(s * 0.36, s * 0.22, s * 0.48, s * 0.16, s * 0.6, s * 0.18)
        path.cubicTo(s * 0.72, s * 0.2, s * 0.8, s * 0.3, s * 0.78, s * 0.42)
        path.cubicTo(s * 0.76, s * 0.54, s * 0.66, s * 0.58, s * 0.58, s * 0.64)
        path.lineTo(s * 0.58, s * 0.72)
        p.drawPath(path)
        p.drawPoint(QPointF(s * 0.58, s * 0.84))

    def draw_settings(self, p, s):
        p.drawEllipse(QRectF(s * 0.28, s * 0.28, s * 0.44, s * 0.44))
        p.drawLine(QPointF(s * 0.5, s * 0.08), QPointF(s * 0.5, s * 0.24))
        p.drawLine(QPointF(s * 0.5, s * 0.76), QPointF(s * 0.5, s * 0.92))
        p.drawLine(QPointF(s * 0.08, s * 0.5), QPointF(s * 0.24, s * 0.5))
        p.drawLine(QPointF(s * 0.76, s * 0.5), QPointF(s * 0.92, s * 0.5))

    def draw_document(self, p, s):
        p.drawRoundedRect(QRectF(s * 0.24, s * 0.16, s * 0.52, s * 0.68), 2, 2)
        p.drawLine(QPointF(s * 0.34, s * 0.38), QPointF(s * 0.66, s * 0.38))
        p.drawLine(QPointF(s * 0.34, s * 0.52), QPointF(s * 0.66, s * 0.52))
        p.drawLine(QPointF(s * 0.34, s * 0.66), QPointF(s * 0.56, s * 0.66))

    def draw_database(self, p, s):
        p.drawEllipse(QRectF(s * 0.24, s * 0.14, s * 0.52, s * 0.18))
        p.drawLine(QPointF(s * 0.24, s * 0.23), QPointF(s * 0.24, s * 0.7))
        p.drawLine(QPointF(s * 0.76, s * 0.23), QPointF(s * 0.76, s * 0.7))
        p.drawArc(QRectF(s * 0.24, s * 0.52, s * 0.52, s * 0.18), 0, -180 * 16)
        p.drawArc(QRectF(s * 0.24, s * 0.34, s * 0.52, s * 0.18), 0, -180 * 16)
        p.drawArc(QRectF(s * 0.24, s * 0.61, s * 0.52, s * 0.18), 180 * 16, 180 * 16)

    def draw_palette(self, p, s):
        p.drawEllipse(QRectF(s * 0.16, s * 0.16, s * 0.68, s * 0.68))
        p.drawEllipse(QRectF(s * 0.28, s * 0.3, s * 0.07, s * 0.07))
        p.drawEllipse(QRectF(s * 0.4, s * 0.24, s * 0.07, s * 0.07))
        p.drawEllipse(QRectF(s * 0.53, s * 0.28, s * 0.07, s * 0.07))

    def draw_folder(self, p, s):
        path = QPainterPath()
        path.moveTo(s * 0.16, s * 0.34); path.lineTo(s * 0.38, s * 0.34)
        path.lineTo(s * 0.45, s * 0.24); path.lineTo(s * 0.84, s * 0.24)
        path.lineTo(s * 0.84, s * 0.74); path.lineTo(s * 0.16, s * 0.74)
        path.closeSubpath()
        p.drawPath(path)

    def draw_clear(self, p, s):
        p.drawLine(QPointF(s * 0.28, s * 0.28), QPointF(s * 0.72, s * 0.72))
        p.drawLine(QPointF(s * 0.72, s * 0.28), QPointF(s * 0.28, s * 0.72))

    def draw_back(self, p, s):
        p.drawLine(QPointF(s * 0.74, s * 0.5), QPointF(s * 0.26, s * 0.5))
        p.drawLine(QPointF(s * 0.26, s * 0.5), QPointF(s * 0.42, s * 0.34))
        p.drawLine(QPointF(s * 0.26, s * 0.5), QPointF(s * 0.42, s * 0.66))

    def draw_submit(self, p, s):
        path = QPainterPath()
        path.moveTo(s * 0.34, s * 0.24)
        path.lineTo(s * 0.34, s * 0.76)
        path.lineTo(s * 0.76, s * 0.5)
        path.closeSubpath()
        p.drawPath(path)

    def draw_font(self, p, s):
        p.drawLine(QPointF(s * 0.25, s * 0.22), QPointF(s * 0.75, s * 0.22))
        p.drawLine(QPointF(s * 0.5, s * 0.22), QPointF(s * 0.5, s * 0.78))
        p.drawLine(QPointF(s * 0.34, s * 0.52), QPointF(s * 0.66, s * 0.52))

    def draw_reset(self, p, s):
        path = QPainterPath()
        path.moveTo(s * 0.3, s * 0.3)
        path.arcTo(QRectF(s * 0.22, s * 0.22, s * 0.56, s * 0.56), 45, 260)
        p.drawPath(path)
        p.drawLine(QPointF(s * 0.3, s * 0.3), QPointF(s * 0.18, s * 0.3))
        p.drawLine(QPointF(s * 0.3, s * 0.3), QPointF(s * 0.3, s * 0.18))

    def draw_excel(self, p, s):
        p.drawRoundedRect(QRectF(s * 0.22, s * 0.16, s * 0.56, s * 0.68), 2, 2)
        p.drawLine(QPointF(s * 0.34, s * 0.3), QPointF(s * 0.66, s * 0.3))
        p.drawLine(QPointF(s * 0.34, s * 0.46), QPointF(s * 0.66, s * 0.46))
        p.drawLine(QPointF(s * 0.34, s * 0.62), QPointF(s * 0.66, s * 0.62))
        p.drawLine(QPointF(s * 0.34, s * 0.22), QPointF(s * 0.34, s * 0.78))

    def draw_column(self, p, s):
        p.drawRoundedRect(QRectF(s * 0.22, s * 0.18, s * 0.56, s * 0.64), 2, 2)
        p.drawLine(QPointF(s * 0.40, s * 0.18), QPointF(s * 0.40, s * 0.82))
        p.drawLine(QPointF(s * 0.58, s * 0.18), QPointF(s * 0.58, s * 0.82))
        p.drawLine(QPointF(s * 0.22, s * 0.36), QPointF(s * 0.78, s * 0.36))

    def draw_link(self, p, s):
        p.drawEllipse(QRectF(s * 0.18, s * 0.34, s * 0.26, s * 0.20))
        p.drawEllipse(QRectF(s * 0.56, s * 0.34, s * 0.26, s * 0.20))
        p.drawLine(QPointF(s * 0.37, s * 0.44), QPointF(s * 0.63, s * 0.44))