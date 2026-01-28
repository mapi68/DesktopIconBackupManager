"""Icon Preview Widget for visualizing icon layouts"""

from typing import Dict, Tuple

from PyQt6.QtWidgets import QWidget, QToolTip
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPainter, QColor

from config import Config


class IconPreviewWidget(QWidget):
    """Widget to display a preview of icon positions"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(Config.PREVIEW_WIDTH, Config.PREVIEW_HEIGHT)
        self.icons = {}
        self.screen_res = (1920, 1080)
        self.setMouseTracking(True)
        self.setStyleSheet(f"""
            QWidget {{
                background-color: {Config.COLOR_BACKGROUND};
                border: 2px solid {Config.COLOR_BORDER};
                border-radius: 4px;
            }}
            QToolTip {{
                color: {Config.COLOR_TOOLTIP_TEXT};
                background-color: {Config.COLOR_TOOLTIP_BG};
                border: 1px solid {Config.COLOR_TOOLTIP_TEXT};
                font-family: 'Segoe UI';
                font-size: 12px;
            }}
        """)

    def update_preview(self, icons: Dict, res_tuple: Tuple[int, int]):
        """Update the preview with new icon positions"""
        self.icons = icons
        self.screen_res = res_tuple if res_tuple else (1920, 1080)
        self.update()

    def paintEvent(self, event):
        """Paint the icon positions as dots"""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        if not self.icons:
            painter.setPen(QColor(Config.COLOR_TEXT_DIM))
            painter.drawText(
                self.rect(),
                Qt.AlignmentFlag.AlignCenter,
                self.tr("No Preview Available"),
            )
            return
        scale_x = self.width() / self.screen_res[0]
        scale_y = self.height() / self.screen_res[1]
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor(Config.COLOR_ICON_DOT))
        for pos in self.icons.values():
            px = int(pos[0] * scale_x)
            py = int(pos[1] * scale_y)
            px = max(
                Config.ICON_DOT_MARGIN, min(px, self.width() - Config.ICON_DOT_MARGIN)
            )
            py = max(
                Config.ICON_DOT_MARGIN, min(py, self.height() - Config.ICON_DOT_MARGIN)
            )
            radius = Config.ICON_DOT_SIZE // 2
            painter.drawEllipse(
                px - radius, py - radius, Config.ICON_DOT_SIZE, Config.ICON_DOT_SIZE
            )

    def mouseMoveEvent(self, event):
        """Show tooltip with icon name when hovering over dots"""
        if not self.icons:
            return
        scale_x = self.width() / self.screen_res[0]
        scale_y = self.height() / self.screen_res[1]
        found_icon = None
        for name, pos in self.icons.items():
            ix = int(pos[0] * scale_x)
            iy = int(pos[1] * scale_y)
            dx = event.position().x() - ix
            dy = event.position().y() - iy
            if (dx * dx + dy * dy) < 144:  # 12 pixels radius
                found_icon = name
                break
        if found_icon:
            QToolTip.showText(event.globalPosition().toPoint(), found_icon, self)
        else:
            QToolTip.hideText()
