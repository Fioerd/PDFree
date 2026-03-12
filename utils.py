"""Shared Qt utilities for PDFree tool modules."""

from PySide6.QtCore import QEvent, QObject, QSize
from PySide6.QtGui import QImage, QPixmap
from PySide6.QtWidgets import QPushButton, QScrollArea

from colors import G100, G700
from icons import svg_icon


def _fitz_pix_to_qpixmap(pix) -> QPixmap:
    """Convert a fitz.Pixmap (RGB) to a QPixmap."""
    try:
        data = pix.samples_mv
    except AttributeError:
        data = pix.samples
    img = QImage(data, pix.width, pix.height, pix.stride,
                 QImage.Format.Format_RGB888)
    return QPixmap.fromImage(img.copy())


def _make_back_button(text: str, callback, color: str = G700) -> QPushButton:
    """Return a styled back button with an arrow-left icon."""
    btn = QPushButton(f"  {text}")
    btn.setIcon(svg_icon("arrow-left", color, 14))
    btn.setIconSize(QSize(14, 14))
    btn.setFixedHeight(36)
    btn.setStyleSheet(f"""
        QPushButton {{
            background: transparent; color: {color}; border: none;
            border-radius: 6px; font: 13px 'Segoe UI';
            text-align: left; padding: 0 8px;
        }}
        QPushButton:hover {{ background: {G100}; }}
    """)
    btn.clicked.connect(callback)
    return btn


class _WheelToHScroll(QObject):
    """Route vertical wheel events to horizontal scroll on a QScrollArea."""

    def __init__(self, sa: QScrollArea):
        super().__init__(sa)
        self._sa = sa
        sa.viewport().installEventFilter(self)

    def eventFilter(self, obj, event):
        if obj is self._sa.viewport() and event.type() == QEvent.Type.Wheel:
            delta = event.angleDelta().y()
            sb = self._sa.horizontalScrollBar()
            sb.setValue(sb.value() - delta // 4)
            return True
        return super().eventFilter(obj, event)
