"""View Tool – Full-featured PDF viewer with annotations, search, forms.

PySide6 port of the original tkinter/customtkinter implementation.
All original features preserved:
  - Open / Save / Print / Add PDF
  - Undo / Redo (Ctrl+Z / Ctrl+Y)
  - Zoom (Fit / Width / +/-, Ctrl+scroll)
  - Rotation (90° increments)
  - Page navigation (first/prev/next/last, page entry field)
  - Thumbnail sidebar with lazy rendering
  - TOC sidebar
  - Tools sidebar (13 tools + color picker + width slider)
  - Text selection (flow-based) with vibrant blue overlay
  - Search (Ctrl+F, highlight all results)
  - Annotations: highlight, underline, strikethrough, freehand, text box,
    sticky note, rect, circle, line, arrow
  - Signature dialog (draw + insert as image)
  - Form filling (text fields, checkboxes, comboboxes as overlay widgets)
  - Double-click to edit existing annotations
  - Keyboard shortcuts (V/S/H/U/K/X/N/D/R/O/L/A for tools)
  - Cleanup method
"""

import io
import math
import os
import subprocess
import tempfile
from enum import Enum
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QLineEdit, QScrollArea, QSlider, QSizePolicy, QFrame,
    QTabWidget, QDialog, QDialogButtonBox, QInputDialog,
    QFileDialog, QMessageBox, QApplication, QCheckBox, QComboBox,
    QTextEdit, QSplitter, QStackedWidget, QGridLayout, QColorDialog,
)
from PySide6.QtCore import (
    Qt, QTimer, QPoint, QPointF, QRectF, QSize, QEvent, QObject,
    QByteArray, QBuffer,
)
from PySide6.QtGui import (
    QPainter, QColor, QPen, QBrush, QPainterPath, QImage, QPixmap,
    QKeySequence, QShortcut, QFont, QFontMetrics, QCursor, QIcon,
)
from icons import svg_pixmap, svg_icon
from colors import (
    BLUE, BLUE_HOVER, GREEN, ORANGE, YELLOW_HL, RED,
    G50, G100, G200, G300, G400, G500, G600, G700, G800, G900,
    WHITE, THUMB_BG,
)
from utils import _fitz_pix_to_qpixmap, _make_back_button

try:
    import fitz  # pymupdf
except ImportError:
    fitz = None

# ---------------------------------------------------------------------------
# Colors (view_tool-specific overrides / additions)
# ---------------------------------------------------------------------------
TOOL_ACTIVE    = "#DBEAFE"
TOOL_BORDER    = "#93C5FD"
SIDEBAR_BG     = "#E2E6EC"
SIDEBAR_BORDER = "#C8CDD5"
TAB_BG         = "#EDF0F4"

# Selection highlight (Windows-style soft blue: #0078D4 @ ~40% opacity)
SEL_BLUE_R, SEL_BLUE_G, SEL_BLUE_B = 0, 120, 212
SEL_ALPHA = 102

# Annotation preset colors  (display_name, hex, fitz_rgb)
ANNOT_COLORS = [
    ("Yellow",  "#FBBF24", (1.0, 0.75, 0.14)),
    ("Red",     "#EF4444", (0.94, 0.27, 0.27)),
    ("Blue",    "#3B82F6", (0.23, 0.51, 0.96)),
    ("Green",   "#22C55E", (0.13, 0.77, 0.37)),
    ("Orange",  "#F97316", (0.98, 0.45, 0.09)),
    ("Black",   "#111827", (0.07, 0.09, 0.15)),
]


# ---------------------------------------------------------------------------
# Tool Enum
# ---------------------------------------------------------------------------
class Tool(Enum):
    VIEW          = "view"
    SELECT        = "select"
    HIGHLIGHT     = "highlight"
    UNDERLINE     = "underline"
    STRIKETHROUGH = "strikethrough"
    FREEHAND      = "freehand"
    TEXT_BOX      = "textbox"
    STICKY_NOTE   = "note"
    RECT          = "rect"
    CIRCLE        = "circle"
    LINE          = "line"
    ARROW         = "arrow"
    SIGN          = "sign"
    EXCERTER      = "excerter"


# (tool, icon, label, shortcut_key)
TOOL_DEFS = [
    (Tool.SELECT,        "\u270f",     "Select",    "S"),
    (Tool.HIGHLIGHT,     "\U0001f58d", "Highlight", "H"),
    (Tool.UNDERLINE,     "_",          "Underline", "U"),
    (Tool.STRIKETHROUGH, "\u2014",     "Strikeout", "K"),
    (Tool.TEXT_BOX,      "T",          "Text Box",  "X"),
    (Tool.STICKY_NOTE,   "\U0001f4ac", "Note",      "N"),
    (Tool.FREEHAND,      "\u270d",     "Freehand",  "D"),
    (Tool.RECT,          "\u25a1",     "Rectangle", "R"),
    (Tool.CIRCLE,        "\u25cb",     "Circle",    "O"),
    (Tool.LINE,          "/",          "Line",      "L"),
    (Tool.ARROW,         "\u2192",     "Arrow",     "A"),
    (Tool.SIGN,          "\u2712",     "Sign",      ""),
    (Tool.EXCERTER,      "\u2702",     "Excerter",  "E"),
]

FIT_PAGE  = -1.0
FIT_WIDTH = -2.0
MAX_UNDO  = 30


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _catmull_rom_segment(p0, p1, p2, p3, num_pts=6):
    """Catmull-Rom spline interpolation between p1 and p2."""
    result = []
    for i in range(num_pts):
        t = i / num_pts
        t2 = t * t
        t3 = t2 * t
        x = 0.5 * ((2 * p1[0]) +
            (-p0[0] + p2[0]) * t +
            (2 * p0[0] - 5 * p1[0] + 4 * p2[0] - p3[0]) * t2 +
            (-p0[0] + 3 * p1[0] - 3 * p2[0] + p3[0]) * t3)
        y = 0.5 * ((2 * p1[1]) +
            (-p0[1] + p2[1]) * t +
            (2 * p0[1] - 5 * p1[1] + 4 * p2[1] - p3[1]) * t2 +
            (-p0[1] + 3 * p1[1] - 3 * p2[1] + p3[1]) * t3)
        result.append((x, y))
    return result


def _smooth_stroke(points, num_interp=6):
    """Smooth a list of (x, y) points using Catmull-Rom spline."""
    if len(points) < 3:
        return list(points)
    result = []
    pts = list(points)
    for i in range(len(pts) - 1):
        p0 = pts[max(i - 1, 0)]
        p1 = pts[i]
        p2 = pts[min(i + 1, len(pts) - 1)]
        p3 = pts[min(i + 2, len(pts) - 1)]
        result.extend(_catmull_rom_segment(p0, p1, p2, p3, num_interp))
    result.append(pts[-1])
    return result


# ---------------------------------------------------------------------------
# Signature Dialog
# ---------------------------------------------------------------------------

class _SigCanvas(QWidget):
    """Inner drawing canvas for the signature dialog."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(400, 150)
        self.setStyleSheet(f"background: {WHITE}; border: 1px solid {G300};")
        self._strokes: list[list[QPointF]] = []
        self._current: list[QPointF] = []

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._current = [event.position()]
            self.update()

    def mouseMoveEvent(self, event):
        if event.buttons() & Qt.MouseButton.LeftButton and self._current is not None:
            self._current.append(event.position())
            self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton and self._current:
            self._strokes.append(list(self._current))
            self._current = []
            self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.fillRect(self.rect(), QColor(WHITE))
        pen = QPen(QColor(G900), 2, Qt.PenStyle.SolidLine,
                   Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
        p.setPen(pen)
        for stroke in self._strokes:
            if len(stroke) >= 2:
                path = QPainterPath()
                path.moveTo(stroke[0])
                for pt in stroke[1:]:
                    path.lineTo(pt)
                p.drawPath(path)
        if len(self._current) >= 2:
            path = QPainterPath()
            path.moveTo(self._current[0])
            for pt in self._current[1:]:
                path.lineTo(pt)
            p.drawPath(path)

    def clear(self):
        self._strokes.clear()
        self._current = []
        self.update()

    def get_png_bytes(self):
        """Render strokes to a PNG with transparent background, return bytes."""
        img = QImage(400, 150, QImage.Format.Format_ARGB32)
        img.fill(Qt.GlobalColor.transparent)
        p = QPainter(img)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        pen = QPen(QColor(G900), 2, Qt.PenStyle.SolidLine,
                   Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
        p.setPen(pen)
        for stroke in self._strokes:
            if len(stroke) >= 2:
                # Apply Catmull-Rom smoothing
                raw = [(pt.x(), pt.y()) for pt in stroke]
                smoothed = _smooth_stroke(raw, num_interp=6)
                path = QPainterPath()
                path.moveTo(smoothed[0][0], smoothed[0][1])
                for sx, sy in smoothed[1:]:
                    path.lineTo(sx, sy)
                p.drawPath(path)
        p.end()
        buf = QByteArray()
        tmp_buf = QBuffer(buf)
        tmp_buf.open(QBuffer.OpenModeFlag.WriteOnly)
        img.save(tmp_buf, "PNG")
        tmp_buf.close()
        return bytes(buf.data())


class SignatureDialog(QDialog):
    """Dialog for drawing a signature to be inserted into the PDF."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Draw Signature")
        self.setFixedSize(440, 280)
        self.setModal(True)
        self._png_bytes = None

        lay = QVBoxLayout(self)
        lay.setSpacing(8)

        lbl = QLabel("Draw your signature below:")
        lbl.setStyleSheet(f"font: 13px 'Segoe UI'; color: {G700};")
        lay.addWidget(lbl)

        self._sig_canvas = _SigCanvas(self)
        lay.addWidget(self._sig_canvas, alignment=Qt.AlignmentFlag.AlignHCenter)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)

        clear_btn = QPushButton("Clear")
        clear_btn.setFixedHeight(34)
        clear_btn.setStyleSheet(
            f"QPushButton {{ background: {G400}; color: {WHITE}; border-radius: 6px; "
            f"font: 12px 'Segoe UI'; padding: 0 16px; }}"
            f"QPushButton:hover {{ background: {G500}; }}")
        clear_btn.clicked.connect(self._sig_canvas.clear)

        apply_btn = QPushButton("Apply")
        apply_btn.setFixedHeight(34)
        apply_btn.setStyleSheet(
            f"QPushButton {{ background: {BLUE}; color: {WHITE}; border-radius: 6px; "
            f"font: 12px 'Segoe UI'; font-weight: bold; padding: 0 16px; }}"
            f"QPushButton:hover {{ background: {BLUE_HOVER}; }}")
        apply_btn.clicked.connect(self._on_apply)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setFixedHeight(34)
        cancel_btn.setStyleSheet(
            f"QPushButton {{ background: {WHITE}; color: {G700}; border: 1px solid {G300}; "
            f"border-radius: 6px; font: 12px 'Segoe UI'; padding: 0 16px; }}"
            f"QPushButton:hover {{ background: {G100}; }}")
        cancel_btn.clicked.connect(self.reject)

        btn_row.addStretch()
        btn_row.addWidget(clear_btn)
        btn_row.addWidget(apply_btn)
        btn_row.addWidget(cancel_btn)
        btn_row.addStretch()
        lay.addLayout(btn_row)

    def _on_apply(self):
        total_pts = sum(len(s) for s in self._sig_canvas._strokes)
        if total_pts < 3:
            self.reject()
            return
        self._png_bytes = self._sig_canvas.get_png_bytes()
        self.accept()

    def get_png_bytes(self):
        return self._png_bytes


# ---------------------------------------------------------------------------
# PDF Canvas Widget
# ---------------------------------------------------------------------------

class PDFCanvas(QWidget):
    """Inner widget inside QScrollArea that renders the PDF page + overlays."""

    def __init__(self, view_tool):
        super().__init__()
        self._vt = view_tool
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setAttribute(Qt.WidgetAttribute.WA_OpaquePaintEvent, True)
        self._pixmap: QPixmap | None = None

    def set_pixmap(self, pm: QPixmap):
        self._pixmap = pm
        self.update()

    # ------------------------------------------------------------------
    # Paint
    # ------------------------------------------------------------------

    def paintEvent(self, event):
        vt = self._vt
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)

        # Background
        p.fillRect(self.rect(), QColor(G50))

        if self._pixmap is None:
            p.setPen(QColor(G400))
            p.setFont(QFont("Segoe UI", 16))
            p.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter,
                       "Open a PDF to start viewing")
            return

        ox = int(vt._page_ox)
        oy = int(vt._page_oy)

        # Page image
        p.drawPixmap(ox, oy, self._pixmap)

        # Page border
        pen = QPen(QColor(G300), 1)
        p.setPen(pen)
        p.drawRect(ox - 1, oy - 1,
                   self._pixmap.width() + 1, self._pixmap.height() + 1)

        # --- Search highlights ---
        if vt._search_flat and vt.doc:
            for idx, (pg, rect) in enumerate(vt._search_flat):
                if pg != vt.current_page:
                    continue
                x0, y0, x1, y1 = vt._pdf_rect_to_canvas(rect)
                is_current = (idx == vt._search_idx)
                if is_current:
                    fill = QColor(249, 115, 22, 180)   # orange
                else:
                    fill = QColor(253, 224, 71, 140)   # yellow
                p.fillRect(QRectF(x0, y0, x1 - x0, y1 - y0), fill)

        # --- Text selection ---
        if vt._selected_words and fitz:
            # Merge all selected chars into one bounding rect per line so the
            # highlight is a solid, gap-free band regardless of inter-character
            # spacing or kerning gaps between individual glyph boxes.
            line_rects: dict = {}
            for w in vt._selected_words:
                key = (w[5], w[6])  # (block_no, line_no)
                if key not in line_rects:
                    line_rects[key] = [w[0], w[1], w[2], w[3]]
                else:
                    e = line_rects[key]
                    e[0] = min(e[0], w[0]); e[1] = min(e[1], w[1])
                    e[2] = max(e[2], w[2]); e[3] = max(e[3], w[3])
            sel_path = QPainterPath()
            for bbox in line_rects.values():
                cx0, cy0, cx1, cy1 = vt._pdf_rect_to_canvas(fitz.Rect(*bbox))
                sel_path.addRect(QRectF(cx0, cy0, cx1 - cx0, cy1 - cy0))
            sel_color = QColor(SEL_BLUE_R, SEL_BLUE_G, SEL_BLUE_B, SEL_ALPHA)
            p.fillPath(sel_path.simplified(), sel_color)

        # --- Excerter rubber-band + flash (independent of _drag_start) ---
        if vt._rb_start is not None and vt._rb_current is not None:
            sx, sy = vt._rb_start
            cx2, cy2 = vt._rb_current
            rx = int(min(sx, cx2));  ry = int(min(sy, cy2))
            rw = int(abs(cx2 - sx)); rh = int(abs(cy2 - sy))
            p.fillRect(QRectF(rx, ry, rw, rh), QColor(59, 130, 246, 20))
            p.setPen(QPen(QColor("#3B82F6"), 2, Qt.PenStyle.DashLine))
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawRect(QRectF(rx, ry, rw, rh))
            p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
            p.setPen(QPen(QColor("#2563EB"), 2))
            p.setBrush(QBrush(QColor("#FFFFFF")))
            x0c, y0c = rx, ry
            x1c, y1c = rx + rw, ry + rh
            for hx, hy in [(x0c-6, y0c-6), (x1c-6, y0c-6),
                           (x0c-6, y1c-6), (x1c-6, y1c-6)]:
                p.drawEllipse(QRectF(hx, hy, 12, 12))
            p.setRenderHint(QPainter.RenderHint.Antialiasing, False)

        if vt._excerpt_flash_rect is not None:
            fx, fy, fw, fh = vt._excerpt_flash_rect
            p.fillRect(QRectF(fx, fy, fw, fh), QColor(59, 130, 246, 50))
            p.setPen(QPen(QColor("#3B82F6"), 2))
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawRect(QRectF(fx, fy, fw, fh))

        # --- Live annotation preview ---
        if vt._drag_start is not None and vt.doc:
            self._paint_live_preview(p, vt)

    def _paint_live_preview(self, p, vt):
        """Draw in-progress annotation shape during drag."""
        if vt._tool == Tool.FREEHAND and len(vt._freehand_pts) >= 2:
            _, hex_c, _ = vt._annot_color
            pen = QPen(QColor(hex_c), vt._stroke_width,
                       Qt.PenStyle.SolidLine,
                       Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
            p.setPen(pen)
            pts = vt._freehand_pts
            for i in range(len(pts) - 1):
                c0 = vt._pdf_to_canvas(pts[i][0], pts[i][1])
                c1 = vt._pdf_to_canvas(pts[i+1][0], pts[i+1][1])
                p.drawLine(QPointF(c0[0], c0[1]), QPointF(c1[0], c1[1]))
            return

        if vt._tool in (Tool.RECT, Tool.CIRCLE, Tool.LINE, Tool.ARROW):
            if vt._drag_start is None or vt._drag_current is None:
                return
            sx, sy = vt._drag_start
            cx, cy = vt._drag_current
            dx = cx - sx
            dy = cy - sy

            mods = QApplication.keyboardModifiers()
            shift = bool(mods & Qt.KeyboardModifier.ShiftModifier)
            if shift:
                if vt._tool in (Tool.RECT, Tool.CIRCLE):
                    side = max(abs(dx), abs(dy))
                    dx = side if dx >= 0 else -side
                    dy = side if dy >= 0 else -side
                elif vt._tool in (Tool.LINE, Tool.ARROW):
                    angle = math.atan2(dy, dx)
                    snap = round(angle / (math.pi / 4)) * (math.pi / 4)
                    length = math.hypot(dx, dy)
                    dx = length * math.cos(snap)
                    dy = length * math.sin(snap)

            ex = sx + dx
            ey = sy + dy

            _, hex_c, _ = vt._annot_color
            pen = QPen(QColor(hex_c), vt._stroke_width,
                       Qt.PenStyle.SolidLine,
                       Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
            p.setPen(pen)
            p.setBrush(Qt.BrushStyle.NoBrush)

            if vt._tool == Tool.RECT:
                p.drawRect(QRectF(min(sx, ex), min(sy, ey),
                                  abs(ex - sx), abs(ey - sy)))
            elif vt._tool == Tool.CIRCLE:
                p.drawEllipse(QRectF(min(sx, ex), min(sy, ey),
                                     abs(ex - sx), abs(ey - sy)))
            elif vt._tool in (Tool.LINE, Tool.ARROW):
                p.drawLine(QPointF(sx, sy), QPointF(ex, ey))
                if vt._tool == Tool.ARROW:
                    self._paint_arrowhead(p, sx, sy, ex, ey,
                                          size=max(8, vt._stroke_width * 4))

    def _paint_arrowhead(self, p, x1, y1, x2, y2, size=10):
        angle = math.atan2(y2 - y1, x2 - x1)
        ap1 = QPointF(x2 - size * math.cos(angle - 0.4),
                      y2 - size * math.sin(angle - 0.4))
        ap2 = QPointF(x2 - size * math.cos(angle + 0.4),
                      y2 - size * math.sin(angle + 0.4))
        path = QPainterPath()
        path.moveTo(x2, y2)
        path.lineTo(ap1)
        path.lineTo(ap2)
        path.closeSubpath()
        p.fillPath(path, p.pen().color())

    # ------------------------------------------------------------------
    # Mouse events
    # ------------------------------------------------------------------

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            pos = event.position()
            self._vt._on_mouse_down(pos.x(), pos.y())

    def mouseMoveEvent(self, event):
        if event.buttons() & Qt.MouseButton.LeftButton:
            pos = event.position()
            self._vt._on_mouse_move(pos.x(), pos.y())

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            pos = event.position()
            self._vt._on_mouse_up(pos.x(), pos.y())

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            pos = event.position()
            self._vt._on_double_click(pos.x(), pos.y())

    def wheelEvent(self, event):
        mods = QApplication.keyboardModifiers()
        delta = event.angleDelta().y()
        if mods & Qt.KeyboardModifier.ControlModifier:
            if delta > 0:
                self._vt._zoom_in()
            else:
                self._vt._zoom_out()
        else:
            vt = self._vt
            if vt.zoom == FIT_PAGE:
                if delta > 0:
                    vt._prev_page()
                else:
                    vt._next_page()
            else:
                event.ignore()  # let scroll area handle it

    # ------------------------------------------------------------------
    # Key events (single-key tool shortcuts)
    # ------------------------------------------------------------------

    def keyPressEvent(self, event):
        focused = QApplication.focusWidget()
        if isinstance(focused, (QLineEdit, QTextEdit)):
            super().keyPressEvent(event)
            return
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            super().keyPressEvent(event)
            return

        shortcut_map = {
            Qt.Key.Key_V: Tool.VIEW,
            Qt.Key.Key_S: Tool.SELECT,
            Qt.Key.Key_H: Tool.HIGHLIGHT,
            Qt.Key.Key_U: Tool.UNDERLINE,
            Qt.Key.Key_K: Tool.STRIKETHROUGH,
            Qt.Key.Key_X: Tool.TEXT_BOX,
            Qt.Key.Key_N: Tool.STICKY_NOTE,
            Qt.Key.Key_D: Tool.FREEHAND,
            Qt.Key.Key_R: Tool.RECT,
            Qt.Key.Key_O: Tool.CIRCLE,
            Qt.Key.Key_L: Tool.LINE,
            Qt.Key.Key_A: Tool.ARROW,
            Qt.Key.Key_E: Tool.EXCERTER,
        }
        key = event.key()
        if key in shortcut_map:
            self._vt._set_tool(shortcut_map[key])
        else:
            super().keyPressEvent(event)


# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# Recent-color swatch – shows delete button on hover
# ---------------------------------------------------------------------------

class _RecentSwatch(QWidget):
    """A 26×26 colored swatch that reveals a ✕ delete button on hover."""

    def __init__(self, hex_color: str, on_select, on_delete, parent=None):
        super().__init__(parent)
        self.setFixedSize(26, 26)

        self._btn = QPushButton(self)
        self._btn.setGeometry(0, 0, 26, 26)
        self._btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self._btn.setToolTip(hex_color)
        self._btn.setStyleSheet(
            f"QPushButton {{ background: {hex_color}; border-radius: 5px; "
            f"border: 1px solid rgba(0,0,0,0.12); }}"
            f"QPushButton:hover {{ border: 2px solid white; }}"
        )
        self._btn.clicked.connect(on_select)

        self._del = QPushButton("×", self)
        self._del.setGeometry(14, -3, 14, 14)
        self._del.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self._del.setStyleSheet(
            "QPushButton { background: #EF4444; color: white; border-radius: 7px; "
            "font: bold 9px 'Segoe UI'; border: none; }"
            "QPushButton:hover { background: #DC2626; }"
        )
        self._del.clicked.connect(on_delete)
        self._del.setVisible(False)

    def enterEvent(self, event):
        self._del.setVisible(True)
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._del.setVisible(False)
        super().leaveEvent(event)


# ---------------------------------------------------------------------------
# Viewport event filter – detects resize for FIT_PAGE / FIT_WIDTH re-render
# ---------------------------------------------------------------------------

class _ViewportFilter(QObject):
    def __init__(self, view_tool):
        super().__init__()
        self._vt = view_tool

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Type.Resize:
            vt = self._vt
            if vt.doc and vt.zoom in (FIT_PAGE, FIT_WIDTH):
                QTimer.singleShot(0, vt._render_page)
        return False


# ---------------------------------------------------------------------------
# Main ViewTool widget
# ---------------------------------------------------------------------------

class ViewTool(QWidget):
    THUMB_W   = 80
    ZOOM_STEP = 0.25
    ZOOM_MIN  = 0.25
    ZOOM_MAX  = 5.0
    SIDEBAR_W = 256

    def __init__(self, parent=None, initial_path: str = "", back_callback=None):
        super().__init__(parent)
        self._back_callback = back_callback

        if fitz is None:
            lay = QVBoxLayout(self)
            lbl = QLabel(
                "\u26a0  Missing dependencies.\n\npip install pymupdf")
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl.setStyleSheet(f"color: {G500}; font: 16px 'Segoe UI';")
            lay.addWidget(lbl)
            return

        # -- Document state --
        self.pdf_path     = ""
        self.doc          = None
        self.total_pages  = 0
        self.current_page = 0
        self._modified    = False

        # -- View state --
        self.zoom         = FIT_PAGE
        self._rotation    = 0
        self._thumb_imgs: list = []
        self._highlighted_thumb_frame = None
        self._thumb_render_next = 0
        self._thumb_timer = None

        # -- Coordinate mapping (set during render) --
        self._page_ox    = 0.0
        self._page_oy    = 0.0
        self._render_mat = fitz.Matrix(1, 1) if fitz else None
        self._inv_mat    = fitz.Matrix(1, 1) if fitz else None

        # -- Tool state --
        self._tool           = Tool.VIEW
        self._annot_color_idx = 0
        self._stroke_width   = 2
        self._drag_start     = None   # (canvas_x, canvas_y)
        self._drag_current   = None   # (canvas_x, canvas_y)
        self._freehand_pts: list[tuple] = []
        self._selected_words: list = []
        self._selection_text = ""
        self._shift_held     = False
        self._custom_color   = None   # (name, hex, fitz_rgb) or None
        self._recent_colors: list[str] = []  # hex strings, newest first

        # -- Excerter state --
        self._excerpts: list = []          # list of {path, page, rect, thumb, label}
        self._excerpt_out = fitz.Document() if fitz else None  # accumulating output
        self._excerpt_y_cursor  = 0.0
        self._excerpt_has_page  = False
        self._excerpt_flash_rect: tuple | None = None  # (x, y, w, h) canvas coords
        self._rb_start:   tuple | None = None   # rubber-band start (canvas)
        self._rb_current: tuple | None = None   # rubber-band current (canvas)
        self._page_iw: float = 0.0              # rendered page width  (canvas px)
        self._page_ih: float = 0.0              # rendered page height (canvas px)

        # -- Search state --
        self._search_results: list[tuple] = []
        self._search_flat: list[tuple]    = []
        self._search_idx   = -1
        self._search_visible = False

        # -- Form widget overlays --
        self._form_widgets: list = []

        # -- Undo / Redo --
        self._undo_stack: list[tuple] = []
        self._redo_stack: list[tuple] = []

        self._build_ui()
        self._install_shortcuts()

        if initial_path:
            self.pdf_path = initial_path
            self._load_pdf()

    # ==================================================================
    # BUILD UI
    # ==================================================================

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # -- Header (unified, matches Figma design) --------------------
        toolbar = QFrame()
        toolbar.setFixedHeight(64)
        toolbar.setStyleSheet(
            f"QFrame {{ background: {WHITE}; border-bottom: 1px solid {G200}; }}"
        )
        tb_lay = QHBoxLayout(toolbar)
        tb_lay.setContentsMargins(24, 0, 24, 0)
        tb_lay.setSpacing(0)

        # Left group: Back link + action buttons
        left_grp = QWidget()
        left_grp.setStyleSheet("background: transparent;")
        left_lay = QHBoxLayout(left_grp)
        left_lay.setContentsMargins(0, 0, 0, 0)
        left_lay.setSpacing(8)

        if self._back_callback:
            back_btn = _make_back_button("Back to Home", self._back_callback, color=G500)
            left_lay.addWidget(back_btn)
            bk_div = QFrame()
            bk_div.setFixedSize(1, 24)
            bk_div.setStyleSheet(f"background: {G200}; border: none;")
            left_lay.addWidget(bk_div)

        def _hbtn(text, bg=WHITE, hover=G100, fg=G700, border=G300, bold=False):
            b = QPushButton(text)
            b.setFixedHeight(36)
            b.setStyleSheet(
                f"QPushButton {{ background: {bg}; color: {fg}; "
                f"border: 1px solid {border}; border-radius: 8px; "
                f"font: {'bold ' if bold else ''}13px 'Segoe UI'; padding: 0 16px; }}"
                f"QPushButton:hover {{ background: {hover}; }}"
            )
            return b

        open_btn = _hbtn("Open")
        open_btn.clicked.connect(self._pick_pdf)
        left_lay.addWidget(open_btn)

        save_btn = _hbtn("Save")
        save_btn.clicked.connect(self._save_pdf)
        left_lay.addWidget(save_btn)

        print_btn = _hbtn("Print")
        print_btn.clicked.connect(self._print_pdf)
        left_lay.addWidget(print_btn)

        left_lay.addSpacing(8)

        add_btn = _hbtn("+ Add PDF", bg=GREEN, hover="#15803D", fg=WHITE, border=GREEN, bold=True)
        add_btn.clicked.connect(self._add_pdf)
        left_lay.addWidget(add_btn)

        # Hidden labels for compatibility (file name is no longer shown in toolbar)
        self._file_lbl = QLabel()
        self._file_lbl.hide()

        tb_lay.addWidget(left_grp)
        tb_lay.addStretch()

        # Right group: Pages/TOC/Tools pill tabs + divider + zoom controls
        right_grp = QWidget()
        right_grp.setStyleSheet("background: transparent;")
        right_lay = QHBoxLayout(right_grp)
        right_lay.setContentsMargins(0, 0, 0, 0)
        right_lay.setSpacing(16)

        # Pill tab group
        tab_pill = QFrame()
        tab_pill.setStyleSheet(
            f"QFrame {{ background: {G100}; border-radius: 8px; border: none; }}"
        )
        pill_lay = QHBoxLayout(tab_pill)
        pill_lay.setContentsMargins(4, 4, 4, 4)
        pill_lay.setSpacing(0)
        self._tab_btns: dict[int, QPushButton] = {}
        for label, idx in [("Pages", 0), ("TOC", 1), ("Tools", 2)]:
            b = QPushButton(label)
            b.setFixedHeight(24)
            b.setStyleSheet(
                f"QPushButton {{ background: {'#FFFFFF' if idx == 2 else 'transparent'}; "
                f"color: {G700}; border: none; border-radius: 4px; "
                f"font: 12px 'Segoe UI'; padding: 0 12px; }}"
                f"QPushButton:hover {{ background: {G200}; }}"
            )
            b.clicked.connect(lambda checked=False, i=idx: self._switch_sidebar_tab(i))
            pill_lay.addWidget(b)
            self._tab_btns[idx] = b
        right_lay.addWidget(tab_pill)

        r_div = QFrame()
        r_div.setFixedSize(1, 24)
        r_div.setStyleSheet(f"background: {G200}; border: none;")
        right_lay.addWidget(r_div)

        # Zoom controls
        zoom_grp = QWidget()
        zoom_grp.setStyleSheet("background: transparent;")
        zoom_lay = QHBoxLayout(zoom_grp)
        zoom_lay.setContentsMargins(0, 0, 0, 0)
        zoom_lay.setSpacing(4)

        rotate_btn = QPushButton("↺")
        rotate_btn.setFixedSize(32, 32)
        rotate_btn.setFont(QFont("Segoe UI", 14))
        rotate_btn.setStyleSheet(
            f"QPushButton {{ background: transparent; color: {G700}; border: none; border-radius: 6px; }}"
            f"QPushButton:hover {{ background: {G100}; }}"
        )
        rotate_btn.clicked.connect(self._rotate_view)
        zoom_lay.addWidget(rotate_btn)

        self._btn_fit = QPushButton("Fit")
        self._btn_fit.setFixedHeight(32)
        self._btn_fit.setStyleSheet(
            f"QPushButton {{ background: {WHITE}; color: {G700}; border: 1px solid {G300}; "
            f"border-radius: 8px; font: 13px 'Segoe UI'; padding: 0 13px; }}"
            f"QPushButton:hover {{ background: {G100}; }}"
        )
        self._btn_fit.clicked.connect(self._zoom_fit)
        zoom_lay.addWidget(self._btn_fit)

        self._btn_fitw = QPushButton("W")
        self._btn_fitw.setFixedHeight(32)
        self._btn_fitw.setStyleSheet(
            f"QPushButton {{ background: {WHITE}; color: {G700}; border: 1px solid {G300}; "
            f"border-radius: 8px; font: bold 13px 'Segoe UI'; padding: 0 13px; }}"
            f"QPushButton:hover {{ background: {G100}; }}"
        )
        self._btn_fitw.clicked.connect(self._zoom_fit_width)
        zoom_lay.addWidget(self._btn_fitw)

        zoom_out_btn = QPushButton("−")
        zoom_out_btn.setFixedSize(32, 32)
        zoom_out_btn.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        zoom_out_btn.setStyleSheet(
            f"QPushButton {{ background: transparent; color: {G700}; border: none; border-radius: 6px; }}"
            f"QPushButton:hover {{ background: {G100}; }}"
        )
        zoom_out_btn.clicked.connect(self._zoom_out)
        zoom_lay.addWidget(zoom_out_btn)

        self._zoom_lbl = QLabel("Fit")
        self._zoom_lbl.setFixedWidth(45)
        self._zoom_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._zoom_lbl.setStyleSheet(
            f"color: {G700}; font: bold 12px 'Segoe UI'; background: transparent; border: none;"
        )
        zoom_lay.addWidget(self._zoom_lbl)

        zoom_in_btn = QPushButton("+")
        zoom_in_btn.setFixedSize(32, 32)
        zoom_in_btn.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        zoom_in_btn.setStyleSheet(
            f"QPushButton {{ background: transparent; color: {G700}; border: none; border-radius: 6px; }}"
            f"QPushButton:hover {{ background: {G100}; }}"
        )
        zoom_in_btn.clicked.connect(self._zoom_in)
        zoom_lay.addWidget(zoom_in_btn)

        right_lay.addWidget(zoom_grp)

        r_div2 = QFrame()
        r_div2.setFixedSize(1, 24)
        r_div2.setStyleSheet(f"background: {G200}; border: none;")
        right_lay.addWidget(r_div2)

        self._sidebar_toggle_btn = QPushButton()
        self._sidebar_toggle_btn.setIcon(svg_icon("chevron-left", G700, 16))
        self._sidebar_toggle_btn.setIconSize(QSize(16, 16))
        self._sidebar_toggle_btn.setFixedSize(32, 32)
        self._sidebar_toggle_btn.setToolTip("Hide sidebar")
        self._sidebar_toggle_btn.setStyleSheet(
            f"QPushButton {{ background: transparent; border: none; border-radius: 6px; }}"
            f"QPushButton:hover {{ background: {G100}; }}"
        )
        self._sidebar_toggle_btn.clicked.connect(self._toggle_sidebar)
        right_lay.addWidget(self._sidebar_toggle_btn)

        tb_lay.addWidget(right_grp)
        root.addWidget(toolbar)

        # -- Body: splitter (sidebar | canvas) -------------------------
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setHandleWidth(1)
        splitter.setStyleSheet("QSplitter::handle { background: " + SIDEBAR_BORDER + "; }")

        # -- Sidebar (Figma: 256px, white, border-right) ---------------
        sidebar = QWidget()
        sidebar.setFixedWidth(self.SIDEBAR_W)
        sidebar.setStyleSheet(
            f"background: {WHITE}; border-right: 1px solid {G200};"
        )
        self._sidebar = sidebar
        sb_lay = QVBoxLayout(sidebar)
        sb_lay.setContentsMargins(0, 0, 0, 0)
        sb_lay.setSpacing(0)

        # QStackedWidget: index 0=Pages, 1=TOC, 2=Tools
        self._sidebar_stack = QStackedWidget()
        self._sidebar_stack.setStyleSheet("background: transparent;")

        # -- Pages panel -----------------------------------------------
        pages_w = QWidget()
        pages_lay = QVBoxLayout(pages_w)
        pages_lay.setContentsMargins(0, 0, 0, 0)
        self._thumb_scroll = QScrollArea()
        self._thumb_scroll.setWidgetResizable(True)
        self._thumb_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._thumb_scroll.setStyleSheet(
            f"QScrollArea {{ background: {WHITE}; border: none; }}")
        self._thumb_container = QWidget()
        self._thumb_container.setStyleSheet(f"background: {WHITE};")
        self._thumb_layout = QVBoxLayout(self._thumb_container)
        self._thumb_layout.setContentsMargins(4, 4, 4, 4)
        self._thumb_layout.setSpacing(4)
        self._thumb_layout.addStretch()
        self._thumb_scroll.setWidget(self._thumb_container)
        pages_lay.addWidget(self._thumb_scroll)
        self._sidebar_stack.addWidget(pages_w)   # index 0

        # -- TOC panel (contains both PDF TOC and Excerpts section) -------
        toc_w = QWidget()
        toc_lay = QVBoxLayout(toc_w)
        toc_lay.setContentsMargins(0, 0, 0, 0)
        toc_lay.setSpacing(0)

        # PDF Table of Contents (top half)
        self._toc_scroll = QScrollArea()
        self._toc_scroll.setWidgetResizable(True)
        self._toc_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._toc_scroll.setStyleSheet(
            f"QScrollArea {{ background: {WHITE}; border: none; }}")
        self._toc_container = QWidget()
        self._toc_container.setStyleSheet(f"background: {WHITE};")
        self._toc_layout = QVBoxLayout(self._toc_container)
        self._toc_layout.setContentsMargins(4, 4, 4, 4)
        self._toc_layout.setSpacing(1)
        self._toc_layout.addStretch()
        init_toc_lbl = QLabel("Open a PDF to\nsee table of contents")
        init_toc_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        init_toc_lbl.setStyleSheet(f"color: {G400}; font: 11px 'Segoe UI';")
        self._toc_layout.insertWidget(0, init_toc_lbl)
        self._toc_scroll.setWidget(self._toc_container)
        toc_lay.addWidget(self._toc_scroll, 1)

        # Divider between TOC and Excerpts
        div = QFrame()
        div.setFrameShape(QFrame.Shape.HLine)
        div.setFixedHeight(1)
        div.setStyleSheet(f"background: {G200}; border: none;")
        toc_lay.addWidget(div)

        # Excerpts section (bottom half)
        self._build_excerpts_tab(toc_lay)
        self._sidebar_stack.addWidget(toc_w)     # index 1

        # -- Tools panel -----------------------------------------------
        tools_w = QWidget()
        tools_lay_outer = QVBoxLayout(tools_w)
        tools_lay_outer.setContentsMargins(0, 0, 0, 0)
        tools_lay_outer.setSpacing(0)
        self._build_tools_tab(tools_lay_outer)
        self._sidebar_stack.addWidget(tools_w)   # index 2

        sb_lay.addWidget(self._sidebar_stack, 1)
        # Start on Tools tab
        self._active_tab_idx = 2
        self._sidebar_stack.setCurrentIndex(2)

        splitter.addWidget(sidebar)

        # -- Canvas area -----------------------------------------------
        canvas_container = QWidget()
        canvas_container.setStyleSheet(f"background: {G50};")
        cc_lay = QVBoxLayout(canvas_container)
        cc_lay.setContentsMargins(0, 0, 0, 0)
        cc_lay.setSpacing(0)

        self._scroll_area = QScrollArea()
        self._scroll_area.setWidgetResizable(False)
        self._scroll_area.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._scroll_area.setStyleSheet(
            f"QScrollArea {{ background: {G50}; border: none; }}")

        self._canvas = PDFCanvas(self)
        self._scroll_area.setWidget(self._canvas)

        cc_lay.addWidget(self._scroll_area)

        # Search bar (hidden initially)
        self._search_bar = QWidget()
        self._search_bar.setFixedHeight(44)
        self._search_bar.setStyleSheet(
            f"background: {G100}; border-top: 1px solid {G200};")
        sb2 = QHBoxLayout(self._search_bar)
        sb2.setContentsMargins(12, 6, 12, 6)
        sb2.setSpacing(4)

        sb2.addStretch()

        srch_icon = QLabel("\U0001f50d")
        srch_icon.setStyleSheet(f"color: {G500}; background: transparent; border: none;")
        sb2.addWidget(srch_icon)

        self._search_entry = QLineEdit()
        self._search_entry.setFixedSize(250, 30)
        self._search_entry.setPlaceholderText("Search\u2026")
        self._search_entry.setStyleSheet(
            f"QLineEdit {{ background: {WHITE}; border: 1px solid {G300}; "
            f"border-radius: 5px; color: {G900}; font: 12px 'Segoe UI'; padding: 0 6px; }}")
        self._search_entry.returnPressed.connect(self._do_search)
        sb2.addWidget(self._search_entry)

        sprev_btn = self._small_btn("\u25c0")
        sprev_btn.clicked.connect(self._search_prev)
        sb2.addWidget(sprev_btn)

        snext_btn = self._small_btn("\u25b6")
        snext_btn.clicked.connect(self._search_next)
        sb2.addWidget(snext_btn)

        self._search_count_lbl = QLabel("")
        self._search_count_lbl.setFixedWidth(80)
        self._search_count_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._search_count_lbl.setStyleSheet(
            f"color: {G500}; font: 11px 'Segoe UI'; background: transparent; border: none;")
        sb2.addWidget(self._search_count_lbl)

        close_s_btn = self._small_btn("\u2715")
        close_s_btn.clicked.connect(self._hide_search)
        sb2.addWidget(close_s_btn)

        sb2.addStretch()
        self._search_bar.setVisible(False)
        cc_lay.addWidget(self._search_bar)

        # Nav bar (Figma: 48px, white, Prev | Page pill | Next | divider | Quick Jump)
        nav_bar = QFrame()
        nav_bar.setFixedHeight(48)
        nav_bar.setStyleSheet(
            f"QFrame {{ background: {WHITE}; border-top: 1px solid {G200}; }}"
        )
        nav_lay = QHBoxLayout(nav_bar)
        nav_lay.setContentsMargins(24, 0, 24, 0)
        nav_lay.setSpacing(16)
        nav_lay.addStretch()

        # Hidden first/last for compatibility
        self._btn_first = QPushButton()
        self._btn_first.hide()
        self._btn_first.clicked.connect(self._first_page)
        self._btn_last = QPushButton()
        self._btn_last.hide()
        self._btn_last.clicked.connect(self._last_page)

        self._btn_prev = QPushButton("Prev")
        self._btn_prev.setIcon(svg_icon("chevron-left", G700, 14))
        self._btn_prev.setIconSize(QSize(14, 14))
        self._btn_prev.setFixedHeight(32)
        self._btn_prev.setStyleSheet(
            f"QPushButton {{ background: {WHITE}; color: {G700}; "
            f"border: 1px solid {G300}; border-radius: 8px; "
            f"font: 13px 'Segoe UI'; padding: 0 12px; }}"
            f"QPushButton:hover {{ background: {G100}; }}"
            f"QPushButton:disabled {{ color: {G300}; border-color: {G200}; }}"
        )
        self._btn_prev.clicked.connect(self._prev_page)
        self._btn_prev.setEnabled(False)
        nav_lay.addWidget(self._btn_prev)

        # "Page X / Y" pill
        page_pill = QFrame()
        page_pill.setStyleSheet(
            f"QFrame {{ background: {G100}; border-radius: 8px; border: none; }}"
        )
        pill_lay2 = QHBoxLayout(page_pill)
        pill_lay2.setContentsMargins(12, 4, 12, 4)
        pill_lay2.setSpacing(4)

        page_word = QLabel("Page")
        page_word.setStyleSheet(
            f"color: {G700}; font: bold 12px 'Segoe UI'; background: transparent;"
        )
        pill_lay2.addWidget(page_word)

        self._page_entry = QLineEdit()
        self._page_entry.setFixedSize(28, 20)
        self._page_entry.setText("–")
        self._page_entry.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._page_entry.setStyleSheet(
            f"QLineEdit {{ background: transparent; border: none; "
            f"color: {G700}; font: bold 12px 'Segoe UI'; }}"
        )
        self._page_entry.returnPressed.connect(self._goto_page)
        pill_lay2.addWidget(self._page_entry)

        sep_lbl = QLabel("/")
        sep_lbl.setStyleSheet(
            f"color: {G500}; font: 12px 'Segoe UI'; background: transparent;"
        )
        pill_lay2.addWidget(sep_lbl)

        self._total_lbl = QLabel("–")
        self._total_lbl.setStyleSheet(
            f"color: {G700}; font: bold 12px 'Segoe UI'; background: transparent;"
        )
        pill_lay2.addWidget(self._total_lbl)
        nav_lay.addWidget(page_pill)

        self._btn_next = QPushButton("Next")
        self._btn_next.setIcon(svg_icon("chevron-right", G700, 14))
        self._btn_next.setIconSize(QSize(14, 14))
        self._btn_next.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self._btn_next.setFixedHeight(32)
        self._btn_next.setStyleSheet(
            f"QPushButton {{ background: {WHITE}; color: {G700}; "
            f"border: 1px solid {G300}; border-radius: 8px; "
            f"font: 13px 'Segoe UI'; padding: 0 12px; }}"
            f"QPushButton:hover {{ background: {G100}; }}"
            f"QPushButton:disabled {{ color: {G300}; border-color: {G200}; }}"
        )
        self._btn_next.clicked.connect(self._next_page)
        self._btn_next.setEnabled(False)
        nav_lay.addWidget(self._btn_next)

        nav_div = QFrame()
        nav_div.setFixedSize(1, 16)
        nav_div.setStyleSheet(f"background: {G200}; border: none;")
        nav_lay.addWidget(nav_div)

        qj_lbl = QLabel("QUICK JUMP")
        qj_lbl.setStyleSheet(
            f"color: {G500}; font: bold 10px 'Segoe UI'; "
            f"letter-spacing: 0.5px; background: transparent;"
        )
        nav_lay.addWidget(qj_lbl)

        qj_entry = QLineEdit()
        qj_entry.setFixedSize(48, 28)
        qj_entry.setAlignment(Qt.AlignmentFlag.AlignCenter)
        qj_entry.setStyleSheet(
            f"QLineEdit {{ background: {WHITE}; border: 1px solid {G200}; "
            f"border-radius: 4px; color: {G700}; font: bold 12px 'Segoe UI'; }}"
        )
        qj_entry.returnPressed.connect(
            lambda: (self._page_entry.setText(qj_entry.text()), self._goto_page())
        )
        nav_lay.addWidget(qj_entry)

        nav_lay.addStretch()
        cc_lay.addWidget(nav_bar)

        splitter.addWidget(canvas_container)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)

        root.addWidget(splitter, 1)

        # Install viewport resize filter
        self._vp_filter = _ViewportFilter(self)
        self._scroll_area.viewport().installEventFilter(self._vp_filter)

    # -- Reusable button factories ------------------------------------

    def _small_btn(self, text):
        b = QPushButton(text)
        b.setFixedSize(30, 28)
        b.setStyleSheet(
            f"QPushButton {{ background: {WHITE}; color: {G700}; "
            f"border: 1px solid {G300}; border-radius: 5px; "
            f"font: 12px 'Segoe UI'; }}"
            f"QPushButton:hover {{ background: {G100}; }}")
        return b

    def _nav_btn(self, text):
        b = QPushButton(text)
        b.setFixedSize(34, 30)
        b.setStyleSheet(
            f"QPushButton {{ background: {WHITE}; color: {G700}; "
            f"border: 1px solid {G300}; border-radius: 5px; "
            f"font: 13px 'Segoe UI'; }}"
            f"QPushButton:hover {{ background: {G100}; }}"
            f"QPushButton:disabled {{ color: {G300}; border-color: {G200}; }}")
        return b

    # ==================================================================
    # TOOLS TAB
    # ==================================================================

    def _build_tools_tab(self, outer_layout):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setStyleSheet(
            f"QScrollArea {{ background: {WHITE}; border: none; }}"
        )

        inner = QWidget()
        inner.setStyleSheet(f"background: {WHITE};")
        lay = QVBoxLayout(inner)
        lay.setContentsMargins(10, 0, 10, 0)
        lay.setSpacing(0)

        def _section_hdr(text):
            lbl = QLabel(text)
            lbl.setFixedHeight(32)
            lbl.setStyleSheet(
                f"color: {G500}; font: bold 10px 'Segoe UI'; "
                f"letter-spacing: 1px; background: transparent; padding: 0 16px;"
            )
            return lbl

        # ---- TOOLS section -------------------------------------------
        tools_sec = QWidget()
        tools_sec.setStyleSheet("background: transparent;")
        ts_lay = QVBoxLayout(tools_sec)
        ts_lay.setContentsMargins(0, 16, 0, 8)
        ts_lay.setSpacing(2)
        ts_lay.addWidget(_section_hdr("TOOLS"))

        self._tool_buttons: dict[Tool, QPushButton] = {}

        def _tool_btn(icon_text, label, shortcut, tool):
            b = QPushButton()
            b.setFixedHeight(34)
            b.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
            hint = f" ({shortcut})" if shortcut else ""
            b.setText(f"  {label}{hint}")
            b.setStyleSheet(
                f"QPushButton {{ background: transparent; color: {G700}; border: none; "
                f"border-radius: 6px; font: 12px 'Segoe UI'; text-align: left; "
                f"padding: 0 12px 0 16px; }}"
                f"QPushButton:hover {{ background: {G100}; }}"
            )
            b.clicked.connect(lambda checked=False, t=tool: self._set_tool(t))
            ts_lay.addWidget(b)
            self._tool_buttons[tool] = b

        _tool_btn("👆", "View", "V", Tool.VIEW)
        for tool, icon, label, shortcut in TOOL_DEFS:
            _tool_btn(icon, label, shortcut, tool)

        self._style_tool_btn(self._tool_buttons[Tool.VIEW], True)
        lay.addWidget(tools_sec)

        # ---- Separator -----------------------------------------------
        sep_line = QFrame()
        sep_line.setFrameShape(QFrame.Shape.HLine)
        sep_line.setStyleSheet(f"color: {G200};")
        sep_line.setFixedHeight(1)
        lay.addWidget(sep_line)

        # ---- PROPERTIES section --------------------------------------
        props_sec = QWidget()
        props_sec.setStyleSheet("background: transparent;")
        ps_lay = QVBoxLayout(props_sec)
        ps_lay.setContentsMargins(0, 16, 0, 16)
        ps_lay.setSpacing(16)
        ps_lay.addWidget(_section_hdr("PROPERTIES"))

        # COLOR card
        color_card = QFrame()
        color_card.setStyleSheet(
            f"QFrame {{ background: {WHITE}; border: 1px solid {G200}; "
            f"border-radius: 12px; }}"
        )
        cc_inner = QVBoxLayout(color_card)
        cc_inner.setContentsMargins(13, 13, 13, 13)
        cc_inner.setSpacing(10)

        # COLOR header row
        col_hdr_row = QWidget()
        col_hdr_row.setStyleSheet("background: transparent;")
        chr_lay = QHBoxLayout(col_hdr_row)
        chr_lay.setContentsMargins(0, 0, 0, 0)
        chr_lay.setSpacing(8)
        dot = QLabel()
        dot.setFixedSize(6, 6)
        dot.setStyleSheet(f"background: {BLUE}; border-radius: 3px;")
        chr_lay.addWidget(dot)
        col_hdr_lbl = QLabel("COLOR")
        col_hdr_lbl.setStyleSheet(
            f"color: {G500}; font: bold 11px 'Segoe UI'; "
            f"letter-spacing: 0.55px; background: transparent;"
        )
        chr_lay.addWidget(col_hdr_lbl)
        chr_lay.addStretch()
        cc_inner.addWidget(col_hdr_row)

        # Picker row: clickable preview swatch + hex input + apply
        hex_row_w = QWidget()
        hex_row_w.setStyleSheet("background: transparent;")
        hex_rlay = QHBoxLayout(hex_row_w)
        hex_rlay.setContentsMargins(0, 0, 0, 0)
        hex_rlay.setSpacing(6)

        # Clickable color preview — opens QColorDialog
        self._hex_preview = QPushButton()
        self._hex_preview.setFixedSize(32, 32)
        self._hex_preview.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self._hex_preview.setToolTip("Click to open color picker")
        self._hex_preview.setStyleSheet(
            f"QPushButton {{ background: {BLUE}; border-radius: 6px; "
            f"border: 1px solid {G200}; }}"
            f"QPushButton:hover {{ border: 2px solid {G300}; }}"
        )
        self._hex_preview.clicked.connect(self._open_color_picker)
        hex_rlay.addWidget(self._hex_preview)

        self._hex_entry = QLineEdit()
        self._hex_entry.setFixedHeight(32)
        self._hex_entry.setPlaceholderText("#3B82F6")
        self._hex_entry.setStyleSheet(
            f"QLineEdit {{ background: {G50}; border: 1px solid {G200}; "
            f"border-radius: 6px; color: {G700}; font: 11px 'Consolas'; padding: 0 9px; }}"
        )
        self._hex_entry.returnPressed.connect(self._apply_hex_color)
        hex_rlay.addWidget(self._hex_entry, 1)

        apply_btn = QPushButton("Apply")
        apply_btn.setFixedHeight(32)
        apply_btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        apply_btn.setStyleSheet(
            f"QPushButton {{ background: {G100}; border: 1px solid {G200}; "
            f"border-radius: 6px; color: {G700}; font: 11px 'Segoe UI'; padding: 0 8px; }}"
            f"QPushButton:hover {{ background: {G200}; }}"
        )
        apply_btn.clicked.connect(self._apply_hex_color)
        hex_rlay.addWidget(apply_btn)
        cc_inner.addWidget(hex_row_w)

        # Preset swatches
        pre_lbl = QLabel("PRESETS")
        pre_lbl.setStyleSheet(
            f"color: {G400}; font: bold 9px 'Segoe UI'; "
            f"letter-spacing: 1px; background: transparent;")
        cc_inner.addWidget(pre_lbl)

        swatches_w = QWidget()
        swatches_w.setStyleSheet("background: transparent;")
        sw_lay = QHBoxLayout(swatches_w)
        sw_lay.setContentsMargins(0, 0, 0, 0)
        sw_lay.setSpacing(5)
        self._color_btns: list[QPushButton] = []
        for i, (name, hex_c, _) in enumerate(ANNOT_COLORS):
            btn = QPushButton()
            btn.setFixedSize(26, 26)
            btn.setToolTip(name)
            btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
            is_sel = (i == 0)
            btn.setStyleSheet(
                f"QPushButton {{ background: {hex_c}; border-radius: 5px; "
                f"border: {'2px solid white' if is_sel else '1px solid rgba(0,0,0,0.1)'}; }}"
                f"QPushButton:hover {{ border: 2px solid white; }}"
            )
            btn.clicked.connect(lambda checked=False, idx=i: self._set_color(idx))
            sw_lay.addWidget(btn)
            self._color_btns.append(btn)
        sw_lay.addStretch()
        cc_inner.addWidget(swatches_w)

        # Recent colors section (hidden when empty)
        self._recent_section = QWidget()
        self._recent_section.setStyleSheet("background: transparent;")
        recent_lay = QVBoxLayout(self._recent_section)
        recent_lay.setContentsMargins(0, 0, 0, 0)
        recent_lay.setSpacing(5)

        recent_hdr = QLabel("RECENT")
        recent_hdr.setStyleSheet(
            f"color: {G400}; font: bold 9px 'Segoe UI'; letter-spacing: 1px;")
        recent_lay.addWidget(recent_hdr)

        self._recent_grid_w = QWidget()
        self._recent_grid_w.setStyleSheet("background: transparent;")
        self._recent_grid = QGridLayout(self._recent_grid_w)
        self._recent_grid.setContentsMargins(0, 0, 0, 0)
        self._recent_grid.setSpacing(4)
        recent_lay.addWidget(self._recent_grid_w)

        self._recent_section.setVisible(False)
        cc_inner.addWidget(self._recent_section)

        ps_lay.addWidget(color_card)

        # LINE WIDTH card
        width_card = QFrame()
        width_card.setStyleSheet("background: transparent;")
        wc_lay = QVBoxLayout(width_card)
        wc_lay.setContentsMargins(0, 8, 0, 0)
        wc_lay.setSpacing(12)

        width_hdr_row = QWidget()
        width_hdr_row.setStyleSheet("background: transparent;")
        whr_lay = QHBoxLayout(width_hdr_row)
        whr_lay.setContentsMargins(0, 0, 0, 0)
        whr_lay.setSpacing(0)
        width_lbl_hdr = QLabel("LINE WIDTH")
        width_lbl_hdr.setStyleSheet(
            f"color: {G500}; font: bold 11px 'Segoe UI'; "
            f"letter-spacing: 0.55px; background: transparent;"
        )
        whr_lay.addWidget(width_lbl_hdr)
        whr_lay.addStretch()
        self._width_lbl = QLabel("2.0px")
        self._width_lbl.setStyleSheet(
            f"color: {BLUE}; font: bold 12px 'Segoe UI'; background: transparent;"
        )
        whr_lay.addWidget(self._width_lbl)
        wc_lay.addWidget(width_hdr_row)

        self._width_slider = QSlider(Qt.Orientation.Horizontal)
        self._width_slider.setRange(1, 10)
        self._width_slider.setValue(2)
        self._width_slider.setStyleSheet(
            f"QSlider::groove:horizontal {{ background: {G200}; height: 6px; "
            f"border-radius: 999px; }}"
            f"QSlider::handle:horizontal {{ background: {WHITE}; width: 16px; "
            f"height: 16px; margin: -5px 0; border-radius: 8px; "
            f"border: 2px solid {BLUE}; }}"
            f"QSlider::sub-page:horizontal {{ background: {BLUE}; border-radius: 999px; }}"
        )
        self._width_slider.valueChanged.connect(self._on_width_change)
        wc_lay.addWidget(self._width_slider)

        ps_lay.addWidget(width_card)
        lay.addWidget(props_sec)

        lay.addStretch()
        scroll.setWidget(inner)
        outer_layout.addWidget(scroll)

    # ==================================================================
    # EXCERPTS TAB
    # ==================================================================

    def _build_excerpts_tab(self, outer_layout):
        # Section header row: "EXCERPTS" label + count + Save button
        hdr = QWidget()
        hdr.setFixedHeight(36)
        hdr.setStyleSheet(f"background: {WHITE};")
        hdr_lay = QHBoxLayout(hdr)
        hdr_lay.setContentsMargins(12, 0, 12, 0)
        hdr_lay.setSpacing(6)
        sec_lbl = QLabel("EXCERPTS")
        sec_lbl.setStyleSheet(
            f"color: {G500}; font: bold 10px 'Segoe UI'; letter-spacing: 1px;")
        hdr_lay.addWidget(sec_lbl)
        self._excerpt_count_lbl = QLabel("0")
        self._excerpt_count_lbl.setStyleSheet(
            f"color: {G400}; font: 10px 'Segoe UI';")
        hdr_lay.addWidget(self._excerpt_count_lbl)
        hdr_lay.addStretch()
        save_btn = QPushButton("Save PDF")
        save_btn.setFixedHeight(24)
        save_btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        save_btn.setStyleSheet(
            f"QPushButton {{ background: #16A34A; color: white; border: none; "
            f"border-radius: 5px; font: bold 10px 'Segoe UI'; padding: 0 10px; }}"
            f"QPushButton:hover {{ background: #15803D; }}"
        )
        save_btn.clicked.connect(self._save_excerpt_pdf)
        hdr_lay.addWidget(save_btn)
        outer_layout.addWidget(hdr)

        # Scroll area for excerpt cards
        self._excerpt_scroll = QScrollArea()
        self._excerpt_scroll.setWidgetResizable(True)
        self._excerpt_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._excerpt_scroll.setStyleSheet(
            f"QScrollArea {{ background: {WHITE}; border: none; }}"
        )
        self._excerpt_container = QWidget()
        self._excerpt_container.setStyleSheet(f"background: {WHITE};")
        self._excerpt_list_lay = QVBoxLayout(self._excerpt_container)
        self._excerpt_list_lay.setContentsMargins(8, 8, 8, 8)
        self._excerpt_list_lay.setSpacing(6)
        self._excerpt_list_lay.addStretch()

        empty_lbl = QLabel("Draw a region with the\nExcerter tool (E) to\ncapture excerpts")
        empty_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        empty_lbl.setStyleSheet(f"color: {G400}; font: 11px 'Segoe UI';")
        self._excerpt_list_lay.insertWidget(0, empty_lbl)
        self._excerpt_empty_lbl = empty_lbl

        self._excerpt_scroll.setWidget(self._excerpt_container)
        outer_layout.addWidget(self._excerpt_scroll, 1)

    def _do_excerpt(self, crop_rect, flash_rect=None):
        """Capture a region from the current page and add it to the Excerpts panel."""
        if not self.doc or not fitz or self._excerpt_out is None:
            return
        page_idx = self.current_page
        src_path = self.pdf_path

        # Build thumbnail (fit into 64×64)
        try:
            page = self.doc[page_idx]
            clip_w = max(1.0, crop_rect.width)
            clip_h = max(1.0, crop_rect.height)
            scale = min(64 / clip_w, 64 / clip_h) * 72 / 72
            mat = fitz.Matrix(scale, scale)
            pix = page.get_pixmap(matrix=mat, clip=crop_rect, alpha=False)
            thumb = _fitz_pix_to_qpixmap(pix)
        except Exception:
            thumb = None

        # Pack snippet onto A4 pages vertically
        _A4_W, _A4_H = 595.0, 842.0
        try:
            src_doc = fitz.open(src_path)
            snip_w = crop_rect.width
            snip_h = crop_rect.height
            if snip_w > _A4_W:
                scale  = _A4_W / snip_w
                snip_w = _A4_W
                snip_h = snip_h * scale
                dest_x = 0.0
            else:
                dest_x = crop_rect.x0
                if dest_x + snip_w > _A4_W:
                    dest_x = _A4_W - snip_w
            if not self._excerpt_has_page or (self._excerpt_y_cursor + snip_h > _A4_H):
                self._excerpt_out.new_page(width=_A4_W, height=_A4_H)
                self._excerpt_y_cursor = 0.0
                self._excerpt_has_page = True
            pg = self._excerpt_out[-1]
            dest_rect = fitz.Rect(dest_x, self._excerpt_y_cursor,
                                  dest_x + snip_w, self._excerpt_y_cursor + snip_h)
            pg.show_pdf_page(dest_rect, src_doc, page_idx, clip=crop_rect)
            self._excerpt_y_cursor += snip_h
            src_doc.close()
        except Exception:
            pass

        label = f"Page {page_idx + 1}  –  {int(crop_rect.width)}×{int(crop_rect.height)} pt"
        self._excerpts.append({
            "path": src_path, "page": page_idx,
            "rect": crop_rect, "thumb": thumb, "label": label,
        })
        self._rebuild_excerpt_list()
        # Flash feedback on the canvas
        if flash_rect:
            self._excerpt_flash_rect = flash_rect
            self._canvas.update()
            QTimer.singleShot(240, self._clear_excerpt_flash)
        # Switch sidebar to TOC tab (which now contains the Excerpts section)
        self._switch_sidebar_tab(1)

    def _rebuild_excerpt_list(self):
        """Rebuild the excerpt card list in the sidebar."""
        lay = self._excerpt_list_lay
        # Remove all items from the layout. Skip deleting _excerpt_empty_lbl
        # (it's a persistent widget — just taken out of the layout, not destroyed).
        while lay.count() > 1:
            item = lay.takeAt(0)
            w = item.widget() if item else None
            if w is not None and w is not self._excerpt_empty_lbl:
                w.deleteLater()

        n = len(self._excerpts)
        self._excerpt_count_lbl.setText(f"({n})")
        # Re-insert the empty label at position 0 when there are no excerpts
        if n == 0:
            lay.insertWidget(0, self._excerpt_empty_lbl)
            self._excerpt_empty_lbl.setVisible(True)
        else:
            self._excerpt_empty_lbl.setVisible(False)

        for i, exc in enumerate(self._excerpts):
            card = QFrame()
            card.setStyleSheet(
                f"QFrame {{ background: {WHITE}; border: 1px solid {G200}; "
                f"border-radius: 8px; }}"
            )
            card_lay = QHBoxLayout(card)
            card_lay.setContentsMargins(8, 8, 8, 8)
            card_lay.setSpacing(8)

            # Thumbnail
            thumb_lbl = QLabel()
            thumb_lbl.setFixedSize(56, 56)
            thumb_lbl.setStyleSheet(
                f"border: 1px solid {G200}; border-radius: 4px; background: {G50};")
            thumb_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            if exc.get("thumb"):
                thumb_lbl.setPixmap(
                    exc["thumb"].scaled(54, 54,
                                        Qt.AspectRatioMode.KeepAspectRatio,
                                        Qt.TransformationMode.SmoothTransformation))
            card_lay.addWidget(thumb_lbl)

            # Label + index
            info_col = QVBoxLayout()
            info_col.setSpacing(2)
            idx_lbl = QLabel(f"#{i + 1}")
            idx_lbl.setStyleSheet(
                f"color: {BLUE}; font: bold 11px 'Segoe UI'; border: none;")
            info_col.addWidget(idx_lbl)
            lbl = QLabel(exc["label"])
            lbl.setWordWrap(True)
            lbl.setStyleSheet(f"color: {G700}; font: 10px 'Segoe UI'; border: none;")
            info_col.addWidget(lbl)
            info_col.addStretch()
            card_lay.addLayout(info_col, 1)

            # Delete button
            del_btn = QPushButton()
            del_btn.setIcon(QIcon(svg_pixmap("x", G400, 12)))
            del_btn.setIconSize(QSize(12, 12))
            del_btn.setFixedSize(22, 22)
            del_btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
            del_btn.setStyleSheet(
                f"QPushButton {{ background: transparent; "
                f"border: none; border-radius: 4px; }}"
                f"QPushButton:hover {{ background: #FEE2E2; }}"
            )
            del_btn.clicked.connect(lambda checked=False, idx=i: self._delete_excerpt(idx))
            card_lay.addWidget(del_btn, 0, Qt.AlignmentFlag.AlignTop)

            lay.insertWidget(lay.count() - 1, card)

    def _delete_excerpt(self, idx: int):
        if idx < 0 or idx >= len(self._excerpts):
            return
        self._excerpts.pop(idx)
        # Rebuild output doc without the deleted page
        if self._excerpt_out and fitz:
            self._excerpt_out.close()
            self._excerpt_out = fitz.Document()
            self._excerpt_y_cursor = 0.0
            self._excerpt_has_page = False
            _A4_W, _A4_H = 595.0, 842.0
            for exc in self._excerpts:
                try:
                    src    = fitz.open(exc["path"])
                    snip_w = exc["rect"].width
                    snip_h = exc["rect"].height
                    if snip_w > _A4_W:
                        scale  = _A4_W / snip_w
                        snip_w = _A4_W
                        snip_h = snip_h * scale
                        dest_x = 0.0
                    else:
                        dest_x = exc["rect"].x0
                        if dest_x + snip_w > _A4_W:
                            dest_x = _A4_W - snip_w
                    if not self._excerpt_has_page or (self._excerpt_y_cursor + snip_h > _A4_H):
                        self._excerpt_out.new_page(width=_A4_W, height=_A4_H)
                        self._excerpt_y_cursor = 0.0
                        self._excerpt_has_page = True
                    pg = self._excerpt_out[-1]
                    dr = fitz.Rect(dest_x, self._excerpt_y_cursor,
                                   dest_x + snip_w, self._excerpt_y_cursor + snip_h)
                    pg.show_pdf_page(dr, src, exc["page"], clip=exc["rect"])
                    self._excerpt_y_cursor += snip_h
                    src.close()
                except Exception:
                    pass
        self._rebuild_excerpt_list()

    def _clear_excerpt_flash(self):
        self._excerpt_flash_rect = None
        self._canvas.update()

    def _save_excerpt_pdf(self):
        if not self._excerpts:
            QMessageBox.information(self, "Excerpts", "No excerpts captured yet.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Excerpt PDF", "", "PDF Files (*.pdf)")
        if not path:
            return
        if not path.lower().endswith(".pdf"):
            path += ".pdf"
        try:
            self._excerpt_out.save(path)
            QMessageBox.information(self, "Saved", f"Excerpt PDF saved to:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not save:\n{e}")

    def _toggle_sidebar(self):
        if self._sidebar.isVisible():
            self._sidebar.hide()
            self._sidebar_toggle_btn.setIcon(svg_icon("chevron-right", G700, 16))
            self._sidebar_toggle_btn.setToolTip("Show sidebar")
        else:
            self._sidebar.show()
            self._sidebar_toggle_btn.setIcon(svg_icon("chevron-left", G700, 16))
            self._sidebar_toggle_btn.setToolTip("Hide sidebar")

    def _switch_sidebar_tab(self, idx: int):
        self._active_tab_idx = idx
        self._sidebar_stack.setCurrentIndex(idx)
        self._update_tab_btn_styles()

    def _update_tab_btn_styles(self):
        for i, btn in self._tab_btns.items():
            active = (i == self._active_tab_idx)
            btn.setStyleSheet(
                f"QPushButton {{ background: {'#FFFFFF' if active else 'transparent'}; "
                f"color: {G700}; border: none; border-radius: 4px; "
                f"font: 12px 'Segoe UI'; padding: 0 12px; }}"
                f"QPushButton:hover {{ background: {G200}; }}"
            )

    def _style_tool_btn(self, btn: QPushButton, active: bool):
        if active:
            btn.setStyleSheet(
                f"QPushButton {{ background: {TOOL_ACTIVE}; color: {BLUE}; "
                f"border-left: 3px solid {BLUE}; border-radius: 6px; "
                f"font: 12px 'Segoe UI'; text-align: left; padding: 0 12px 0 13px; }}"
            )
        else:
            btn.setStyleSheet(
                f"QPushButton {{ background: transparent; color: {G700}; "
                f"border: none; border-radius: 6px; "
                f"font: 12px 'Segoe UI'; text-align: left; padding: 0 12px 0 16px; }}"
                f"QPushButton:hover {{ background: {G100}; }}"
            )

    def _set_tool(self, tool: Tool):
        self._tool = tool
        for t, btn in self._tool_buttons.items():
            self._style_tool_btn(btn, t == tool)
        cursors = {
            Tool.VIEW:          Qt.CursorShape.ArrowCursor,
            Tool.SELECT:        Qt.CursorShape.IBeamCursor,
            Tool.HIGHLIGHT:     Qt.CursorShape.IBeamCursor,
            Tool.UNDERLINE:     Qt.CursorShape.IBeamCursor,
            Tool.STRIKETHROUGH: Qt.CursorShape.IBeamCursor,
            Tool.FREEHAND:      Qt.CursorShape.CrossCursor,
            Tool.TEXT_BOX:      Qt.CursorShape.CrossCursor,
            Tool.STICKY_NOTE:   Qt.CursorShape.CrossCursor,
            Tool.RECT:          Qt.CursorShape.CrossCursor,
            Tool.CIRCLE:        Qt.CursorShape.CrossCursor,
            Tool.LINE:          Qt.CursorShape.CrossCursor,
            Tool.ARROW:         Qt.CursorShape.CrossCursor,
            Tool.SIGN:          Qt.CursorShape.CrossCursor,
            Tool.EXCERTER:      Qt.CursorShape.CrossCursor,
        }
        self._canvas.setCursor(QCursor(cursors.get(tool, Qt.CursorShape.ArrowCursor)))

    def _set_color(self, idx: int):
        self._annot_color_idx = idx
        self._custom_color = None
        _, sel_hex, _ = ANNOT_COLORS[idx]
        for i, btn in enumerate(self._color_btns):
            _, hex_c, _ = ANNOT_COLORS[i]
            is_sel = (i == idx)
            btn.setStyleSheet(
                f"QPushButton {{ background: {hex_c}; border-radius: 5px; "
                f"border: {'2px solid white' if is_sel else '1px solid rgba(0,0,0,0.1)'}; }}"
                f"QPushButton:hover {{ border: 2px solid {WHITE}; }}"
            )
        self._hex_preview.setStyleSheet(
            f"QPushButton {{ background: {sel_hex}; border-radius: 6px; "
            f"border: 1px solid {G200}; }}"
            f"QPushButton:hover {{ border: 2px solid {G300}; }}"
        )
        self._hex_entry.setText(sel_hex)

    def _apply_custom_color(self, r: int, g: int, b: int):
        """Apply an RGB color as custom, update UI, and add to recents."""
        hex_str = f"#{r:02X}{g:02X}{b:02X}"
        fitz_rgb = (r / 255.0, g / 255.0, b / 255.0)
        self._custom_color = ("Custom", hex_str, fitz_rgb)
        self._annot_color_idx = -1
        for i, btn in enumerate(self._color_btns):
            _, hex_c, _ = ANNOT_COLORS[i]
            btn.setStyleSheet(
                f"QPushButton {{ background: {hex_c}; border-radius: 5px; "
                f"border: 1px solid rgba(0,0,0,0.1); }}"
                f"QPushButton:hover {{ border: 2px solid {WHITE}; }}"
            )
        self._hex_preview.setStyleSheet(
            f"QPushButton {{ background: {hex_str}; border-radius: 6px; "
            f"border: 1px solid {G200}; }}"
            f"QPushButton:hover {{ border: 2px solid {G300}; }}"
        )
        self._hex_entry.setText(hex_str)
        self._add_to_recent(hex_str)

    def _apply_hex_color(self):
        raw = self._hex_entry.text().strip().lstrip('#')
        if len(raw) == 3:
            raw = raw[0]*2 + raw[1]*2 + raw[2]*2
        if len(raw) != 6:
            return
        try:
            r = int(raw[0:2], 16)
            g = int(raw[2:4], 16)
            b = int(raw[4:6], 16)
        except ValueError:
            return
        self._apply_custom_color(r, g, b)

    def _open_color_picker(self):
        _, current_hex, _ = self._annot_color
        initial = QColor(current_hex)
        color = QColorDialog.getColor(
            initial, self, "Pick Color",
            QColorDialog.ColorDialogOption.ShowAlphaChannel)
        if color.isValid():
            self._apply_custom_color(color.red(), color.green(), color.blue())

    def _add_to_recent(self, hex_str: str):
        hex_str = hex_str.upper()
        if hex_str in self._recent_colors:
            self._recent_colors.remove(hex_str)
        self._recent_colors.insert(0, hex_str)
        if len(self._recent_colors) > 16:
            self._recent_colors.pop()
        self._rebuild_recent_swatches()

    def _rebuild_recent_swatches(self):
        # Clear grid
        while self._recent_grid.count():
            item = self._recent_grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        n = len(self._recent_colors)
        self._recent_section.setVisible(n > 0)
        cols = 7
        for i, hex_c in enumerate(self._recent_colors):
            row, col = divmod(i, cols)
            swatch = _RecentSwatch(
                hex_c,
                on_select=lambda checked=False, h=hex_c: self._apply_recent_color(h),
                on_delete=lambda checked=False, h=hex_c: self._delete_recent_color(h),
            )
            self._recent_grid.addWidget(swatch, row, col)

    def _apply_recent_color(self, hex_str: str):
        raw = hex_str.lstrip('#')
        r = int(raw[0:2], 16)
        g = int(raw[2:4], 16)
        b = int(raw[4:6], 16)
        self._apply_custom_color(r, g, b)

    def _delete_recent_color(self, hex_str: str):
        if hex_str in self._recent_colors:
            self._recent_colors.remove(hex_str)
        self._rebuild_recent_swatches()

    def _on_width_change(self, val: int):
        self._stroke_width = val
        self._width_lbl.setText(f"{val}.0px")

    @property
    def _annot_color(self):
        if self._custom_color:
            return self._custom_color
        return ANNOT_COLORS[max(0, self._annot_color_idx)]

    # ==================================================================
    # SHORTCUTS
    # ==================================================================

    def _install_shortcuts(self):
        shortcuts = [
            ("Ctrl+F",  self._toggle_search),
            ("Ctrl+S",  self._save_pdf),
            ("Ctrl+C",  self._copy_selection),
            ("Ctrl+Z",  self._undo),
            ("Ctrl+Y",  self._redo),
            ("Escape",  self._escape),
            ("Left",    self._prev_page),
            ("Right",   self._next_page),
            ("Home",    self._first_page),
            ("End",     self._last_page),
            ("+",       self._zoom_in),
            ("-",       self._zoom_out),
        ]
        for key, slot in shortcuts:
            sc = QShortcut(QKeySequence(key), self)
            sc.activated.connect(slot)

    # ==================================================================
    # UNDO / REDO
    # ==================================================================

    def _push_undo(self):
        if not self.doc:
            return
        try:
            buf = self.doc.tobytes()
            self._undo_stack.append((buf, self.current_page))
            if len(self._undo_stack) > MAX_UNDO:
                self._undo_stack.pop(0)
            self._redo_stack.clear()
        except Exception:
            pass

    def _undo(self):
        if not self._undo_stack or not self.doc:
            return
        try:
            cur_buf = self.doc.tobytes()
            self._redo_stack.append((cur_buf, self.current_page))
            if len(self._redo_stack) > MAX_UNDO:
                self._redo_stack.pop(0)
            buf, page_idx = self._undo_stack.pop()
            self.doc.close()
            self.doc = fitz.open(stream=buf, filetype="pdf")
            self.total_pages = len(self.doc)
            self._modified = bool(self._undo_stack)
            self._show_page(min(page_idx, self.total_pages - 1))
        except Exception as e:
            QMessageBox.critical(self, "Undo Error", str(e))

    def _redo(self):
        if not self._redo_stack or not self.doc:
            return
        try:
            cur_buf = self.doc.tobytes()
            self._undo_stack.append((cur_buf, self.current_page))
            buf, page_idx = self._redo_stack.pop()
            self.doc.close()
            self.doc = fitz.open(stream=buf, filetype="pdf")
            self.total_pages = len(self.doc)
            self._modified = True
            self._show_page(min(page_idx, self.total_pages - 1))
        except Exception as e:
            QMessageBox.critical(self, "Redo Error", str(e))

    # ==================================================================
    # FILE LOADING
    # ==================================================================

    def _pick_pdf(self):
        p, _ = QFileDialog.getOpenFileName(self, "Open PDF", "",
                                           "PDF Files (*.pdf)")
        if not p:
            return
        self.pdf_path = p
        self._load_pdf()

    def _load_pdf(self):
        try:
            if self.doc:
                self.doc.close()
            self.doc = fitz.open(self.pdf_path)
            self.total_pages = len(self.doc)
            self.current_page = 0
            self.zoom = FIT_PAGE
            self._rotation = 0
            self._modified = False
            self._search_results.clear()
            self._search_flat.clear()
            self._search_idx = -1
            self._undo_stack.clear()
            self._redo_stack.clear()

            name = Path(self.pdf_path).name
            self._file_lbl.setText(name)
            self._total_lbl.setText(str(self.total_pages))
            self._update_zoom_label()

            self._render_thumbnails()
            self._build_toc()
            self._show_page(0)
        except Exception as e:
            self.total_pages = 0
            QMessageBox.critical(self, "Error", f"Could not load PDF:\n{e}")

    def _add_pdf(self):
        if not self.doc:
            self._pick_pdf()
            return
        p, _ = QFileDialog.getOpenFileName(self, "Add PDF", "",
                                           "PDF Files (*.pdf)")
        if not p:
            return
        try:
            self._push_undo()
            src = fitz.open(p)
            self.doc.insert_pdf(src)
            src.close()
            self.total_pages = len(self.doc)
            self._modified = True

            name = Path(self.pdf_path).name
            self._file_lbl.setText(f"{name} (+{Path(p).name})")
            self._total_lbl.setText(str(self.total_pages))

            self._render_thumbnails()
            self._build_toc()
            self._show_page(self.current_page)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not add PDF:\n{e}")

    # ==================================================================
    # THUMBNAILS
    # ==================================================================

    def _render_thumbnails(self):
        # Clear existing thumbnails
        while self._thumb_layout.count() > 1:
            item = self._thumb_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self._thumb_imgs.clear()
        self._highlighted_thumb_frame = None
        if not self.doc:
            return

        for i in range(self.total_pages):
            frame = QWidget()
            frame.setStyleSheet(
                f"background: {WHITE}; border-radius: 4px;")
            frame.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
            frame_lay = QVBoxLayout(frame)
            frame_lay.setContentsMargins(4, 4, 4, 2)
            frame_lay.setSpacing(2)
            frame_lay.setAlignment(Qt.AlignmentFlag.AlignHCenter)

            ph = int(self.THUMB_W * 1.4)
            img_lbl = QLabel()
            img_lbl.setFixedSize(self.THUMB_W, ph)
            img_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            img_lbl.setStyleSheet(
                f"background: {G200}; border: 2px solid {G300}; border-radius: 2px;")
            frame_lay.addWidget(img_lbl, alignment=Qt.AlignmentFlag.AlignHCenter)

            num_lbl = QLabel(str(i + 1))
            num_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            num_lbl.setStyleSheet(
                f"color: {G500}; font: 10px 'Segoe UI'; background: transparent;")
            frame_lay.addWidget(num_lbl)

            frame.mousePressEvent = lambda e, idx=i: self._show_page(idx)
            self._thumb_layout.insertWidget(
                self._thumb_layout.count() - 1, frame)
            self._thumb_imgs.append((None, frame, img_lbl))

        self._thumb_render_next = 0
        QTimer.singleShot(0, self._render_thumb_batch)

    def _render_thumb_batch(self, batch=6):
        if not self.doc:
            return
        start = self._thumb_render_next
        end   = min(start + batch, self.total_pages)

        for i in range(start, end):
            pm, frame, img_lbl = self._thumb_imgs[i]
            if pm is not None:
                continue
            page = self.doc[i]
            s = self.THUMB_W / page.rect.width
            pix = page.get_pixmap(matrix=fitz.Matrix(s, s), alpha=False)
            new_pm = _fitz_pix_to_qpixmap(pix)
            self._thumb_imgs[i] = (new_pm, frame, img_lbl)
            img_lbl.setFixedSize(new_pm.width(), new_pm.height())
            img_lbl.setPixmap(new_pm)
            img_lbl.setStyleSheet(
                f"background: {WHITE}; border: 2px solid {G300}; border-radius: 2px;")

        self._thumb_render_next = end
        if end < self.total_pages:
            QTimer.singleShot(0, self._render_thumb_batch)

    def _highlight_thumb(self, idx: int):
        # Clear previous
        if self._highlighted_thumb_frame is not None:
            try:
                _, frame, img_lbl = self._highlighted_thumb_frame
                img_lbl.setStyleSheet(
                    f"background: {WHITE}; border: 2px solid {G300}; border-radius: 2px;")
            except Exception:
                pass
        if 0 <= idx < len(self._thumb_imgs):
            _, frame, img_lbl = self._thumb_imgs[idx]
            img_lbl.setStyleSheet(
                f"background: {WHITE}; border: 3px solid {BLUE}; border-radius: 2px;")
            self._highlighted_thumb_frame = self._thumb_imgs[idx]
            # Scroll to thumbnail
            QTimer.singleShot(50, lambda: self._scroll_to_thumb(idx))

    def _scroll_to_thumb(self, idx: int):
        if 0 <= idx < len(self._thumb_imgs):
            _, frame, _ = self._thumb_imgs[idx]
            self._thumb_scroll.ensureWidgetVisible(frame)

    # ==================================================================
    # TOC
    # ==================================================================

    def _build_toc(self):
        while self._toc_layout.count() > 1:
            item = self._toc_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if not self.doc:
            return

        toc = self.doc.get_toc(simple=True)
        if not toc:
            lbl = QLabel("No table of contents")
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl.setStyleSheet(f"color: {G400}; font: 11px 'Segoe UI';")
            self._toc_layout.insertWidget(0, lbl)
            return

        for i, (level, title, page_num) in enumerate(toc):
            indent = (level - 1) * 12
            bold = "bold" if level == 1 else "normal"
            btn = QPushButton(title)
            btn.setFixedHeight(26)
            btn.setStyleSheet(
                f"QPushButton {{ background: transparent; color: {G700}; "
                f"border: none; text-align: left; "
                f"font: {bold} 11px 'Segoe UI'; "
                f"padding-left: {indent + 8}px; }}"
                f"QPushButton:hover {{ background: {G200}; border-radius: 4px; }}")
            btn.clicked.connect(
                lambda checked=False, p=page_num: self._show_page(p - 1))
            self._toc_layout.insertWidget(i, btn)

    # ==================================================================
    # COORDINATE MAPPING
    # ==================================================================

    def _canvas_to_pdf(self, cx: float, cy: float):
        """Convert canvas widget coordinates to PDF page coordinates."""
        pt = fitz.Point(cx - self._page_ox, cy - self._page_oy) * self._inv_mat
        return pt.x, pt.y

    def _pdf_to_canvas(self, px: float, py: float):
        """Convert PDF page coordinates to canvas widget coordinates."""
        pt = fitz.Point(px, py) * self._render_mat
        return pt.x + self._page_ox, pt.y + self._page_oy

    def _pdf_rect_to_canvas(self, r):
        """Convert fitz.Rect to canvas widget coords (x0, y0, x1, y1)."""
        x0, y0 = self._pdf_to_canvas(r.x0, r.y0)
        x1, y1 = self._pdf_to_canvas(r.x1, r.y1)
        return min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1)

    # ==================================================================
    # PAGE RENDERING
    # ==================================================================

    def _show_page(self, idx: int):
        if not self.doc or idx < 0 or idx >= self.total_pages:
            return
        self.current_page = idx
        self._clear_form_widgets()
        self._render_page()

        self._page_entry.setText(str(idx + 1))
        has_prev = idx > 0
        has_next = idx < self.total_pages - 1
        self._btn_first.setEnabled(has_prev)
        self._btn_prev.setEnabled(has_prev)
        self._btn_next.setEnabled(has_next)
        self._btn_last.setEnabled(has_next)
        self._highlight_thumb(idx)

    def _render_page(self):
        if not self.doc:
            return

        vp = self._scroll_area.viewport()
        cw = max(vp.width(), 300)
        ch = max(vp.height(), 300)
        page = self.doc[self.current_page]
        pw, ph = page.rect.width, page.rect.height

        if self._rotation in (90, 270):
            pw, ph = ph, pw

        if self.zoom == FIT_PAGE:
            fw, fh = cw - 40, ch - 40
            scale = min(fw / pw, fh / ph)
            scale = max(scale, 0.05)
        elif self.zoom == FIT_WIDTH:
            scale = (cw - 40) / pw
            scale = max(scale, 0.05)
        else:
            scale = self.zoom

        mat = fitz.Matrix(scale, scale).prerotate(self._rotation)
        self._render_mat = mat
        self._inv_mat = ~mat

        pix = page.get_pixmap(matrix=mat, alpha=False)
        pm = _fitz_pix_to_qpixmap(pix)
        iw, ih = pm.width(), pm.height()

        if self.zoom == FIT_PAGE:
            canvas_w = max(cw, iw)
            canvas_h = max(ch, ih)
            ox = (canvas_w - iw) / 2
            oy = (canvas_h - ih) / 2
        else:
            pad = 20
            canvas_w = max(iw + pad * 2, cw)
            canvas_h = max(ih + pad * 2, ch)
            ox = (canvas_w - iw) / 2
            oy = pad

        self._page_ox = ox
        self._page_oy = oy
        self._page_iw = float(iw)
        self._page_ih = float(ih)

        self._canvas.setFixedSize(int(canvas_w), int(canvas_h))
        self._canvas.set_pixmap(pm)

        # Overlay form widgets
        self._draw_form_widgets()

    def _escape(self):
        if self._search_visible:
            self._hide_search()
        else:
            self._set_tool(Tool.VIEW)
            self._selected_words.clear()
            self._selection_text = ""
            self._canvas.update()

    # ==================================================================
    # NAVIGATION
    # ==================================================================

    def _first_page(self):
        self._show_page(0)

    def _prev_page(self):
        if self.current_page > 0:
            self._show_page(self.current_page - 1)

    def _next_page(self):
        if self.current_page < self.total_pages - 1:
            self._show_page(self.current_page + 1)

    def _last_page(self):
        if self.total_pages > 0:
            self._show_page(self.total_pages - 1)

    def _goto_page(self):
        try:
            num = int(self._page_entry.text())
            if 1 <= num <= self.total_pages:
                self._show_page(num - 1)
        except ValueError:
            pass

    # ==================================================================
    # ZOOM & ROTATION
    # ==================================================================

    def _zoom_fit(self):
        self.zoom = FIT_PAGE
        self._update_zoom_label()
        if self.doc:
            self._show_page(self.current_page)

    def _zoom_fit_width(self):
        self.zoom = FIT_WIDTH
        self._update_zoom_label()
        if self.doc:
            self._show_page(self.current_page)

    def _zoom_in(self):
        if self.zoom in (FIT_PAGE, FIT_WIDTH):
            self.zoom = self._effective_zoom()
        self.zoom = min(self.zoom + self.ZOOM_STEP, self.ZOOM_MAX)
        self._update_zoom_label()
        if self.doc:
            self._show_page(self.current_page)

    def _zoom_out(self):
        if self.zoom in (FIT_PAGE, FIT_WIDTH):
            self.zoom = self._effective_zoom()
        self.zoom = max(self.zoom - self.ZOOM_STEP, self.ZOOM_MIN)
        self._update_zoom_label()
        if self.doc:
            self._show_page(self.current_page)

    def _effective_zoom(self):
        if not self.doc:
            return 1.0
        vp = self._scroll_area.viewport()
        cw = max(vp.width(), 300)
        ch = max(vp.height(), 300)
        page = self.doc[self.current_page]
        pw, ph = page.rect.width, page.rect.height
        if self._rotation in (90, 270):
            pw, ph = ph, pw
        if self.zoom == FIT_WIDTH:
            return (cw - 40) / pw
        fw, fh = cw - 40, ch - 40
        return min(fw / pw, fh / ph)

    def _update_zoom_label(self):
        if self.zoom == FIT_PAGE:
            self._zoom_lbl.setText("Fit")
            self._btn_fit.setStyleSheet(
                f"QPushButton {{ background: {BLUE}; color: {WHITE}; "
                f"border-radius: 5px; font: 11px 'Segoe UI'; }}"
                f"QPushButton:hover {{ background: {BLUE_HOVER}; }}")
            self._btn_fitw.setStyleSheet(
                f"QPushButton {{ background: {WHITE}; color: {G700}; "
                f"border: 1px solid {G300}; border-radius: 5px; "
                f"font: bold 11px 'Segoe UI'; }}"
                f"QPushButton:hover {{ background: {G100}; }}")
        elif self.zoom == FIT_WIDTH:
            self._zoom_lbl.setText("Width")
            self._btn_fit.setStyleSheet(
                f"QPushButton {{ background: {WHITE}; color: {G700}; "
                f"border: 1px solid {G300}; border-radius: 5px; font: 11px 'Segoe UI'; }}"
                f"QPushButton:hover {{ background: {G100}; }}")
            self._btn_fitw.setStyleSheet(
                f"QPushButton {{ background: {BLUE}; color: {WHITE}; "
                f"border-radius: 5px; font: bold 11px 'Segoe UI'; }}"
                f"QPushButton:hover {{ background: {BLUE_HOVER}; }}")
        else:
            self._zoom_lbl.setText(f"{int(self.zoom * 100)}%")
            self._btn_fit.setStyleSheet(
                f"QPushButton {{ background: {WHITE}; color: {G700}; "
                f"border: 1px solid {G300}; border-radius: 5px; font: 11px 'Segoe UI'; }}"
                f"QPushButton:hover {{ background: {G100}; }}")
            self._btn_fitw.setStyleSheet(
                f"QPushButton {{ background: {WHITE}; color: {G700}; "
                f"border: 1px solid {G300}; border-radius: 5px; "
                f"font: bold 11px 'Segoe UI'; }}"
                f"QPushButton:hover {{ background: {G100}; }}")

    def _rotate_view(self):
        self._rotation = (self._rotation + 90) % 360
        if self.doc:
            self._show_page(self.current_page)

    # ==================================================================
    # SEARCH  (Ctrl+F)
    # ==================================================================

    def _toggle_search(self):
        if self._search_visible:
            self._hide_search()
        else:
            self._show_search()

    def _show_search(self):
        self._search_bar.setVisible(True)
        self._search_visible = True
        self._search_entry.setFocus()

    def _hide_search(self):
        self._search_bar.setVisible(False)
        self._search_visible = False
        self._search_results.clear()
        self._search_flat.clear()
        self._search_idx = -1
        self._search_count_lbl.setText("")
        if self.doc:
            self._canvas.update()

    def _do_search(self):
        query = self._search_entry.text().strip()
        if not query or not self.doc:
            return
        self._search_results.clear()
        self._search_flat.clear()
        for i in range(self.total_pages):
            page = self.doc[i]
            rects = page.search_for(query)
            if rects:
                self._search_results.append((i, rects))
                for r in rects:
                    self._search_flat.append((i, r))
        total = len(self._search_flat)
        if total == 0:
            self._search_count_lbl.setText("0 results")
            self._search_idx = -1
        else:
            self._search_idx = 0
            for j, (pg, _) in enumerate(self._search_flat):
                if pg >= self.current_page:
                    self._search_idx = j
                    break
            self._goto_search_result()

    def _search_next(self):
        if not self._search_flat:
            return
        self._search_idx = (self._search_idx + 1) % len(self._search_flat)
        self._goto_search_result()

    def _search_prev(self):
        if not self._search_flat:
            return
        self._search_idx = (self._search_idx - 1) % len(self._search_flat)
        self._goto_search_result()

    def _goto_search_result(self):
        if self._search_idx < 0 or not self._search_flat:
            return
        pg, rect = self._search_flat[self._search_idx]
        total = len(self._search_flat)
        self._search_count_lbl.setText(f"{self._search_idx + 1}/{total}")
        if pg != self.current_page:
            self._show_page(pg)
        else:
            self._canvas.update()

    # ==================================================================
    # TEXT SELECTION & COPY
    # ==================================================================

    def _get_words(self):
        if not self.doc:
            return []
        page = self.doc[self.current_page]
        return page.get_text("words")

    def _get_chars(self):
        """Return char tuples (x0,y0,x1,y1, char, block_no, line_no, char_no) sorted in reading order."""
        if not self.doc:
            return []
        page = self.doc[self.current_page]
        chars = []
        for block in page.get_text("rawdict")["blocks"]:
            if block.get("type") != 0:
                continue
            bno = block["number"]
            for lno, line in enumerate(block["lines"]):
                for span in line["spans"]:
                    for cno, ch in enumerate(span["chars"]):
                        x0, y0, x1, y1 = ch["bbox"]
                        if x1 > x0:  # skip zero-width glyphs
                            chars.append((x0, y0, x1, y1, ch["c"], bno, lno, cno))
        chars.sort(key=lambda c: (c[5], c[6], c[0]))
        return chars

    def _select_chars_flow(self, sx, sy, ex, ey):
        """Select individual characters in reading-flow order between two PDF points."""
        chars = self._get_chars()
        if not chars:
            return []
        si = self._nearest_char_index(chars, sx, sy)
        ei = self._nearest_char_index(chars, ex, ey)
        if si is None or ei is None:
            return []
        lo, hi = min(si, ei), max(si, ei)
        return chars[lo:hi + 1]

    @staticmethod
    def _nearest_char_index(chars, px, py):
        # Direct hit: cursor is inside the char's bounding box
        for i, c in enumerate(chars):
            if c[0] <= px <= c[2] and c[1] <= py <= c[3]:
                return i
        # Fallback: nearest by distance, strongly preferring same-line characters
        best_idx = None
        best_dist = float('inf')
        for i, c in enumerate(chars):
            cx = (c[0] + c[2]) * 0.5
            cy = (c[1] + c[3]) * 0.5
            char_h = max(c[3] - c[1], 1.0)
            dy = abs(py - cy)
            dx = abs(px - cx)
            # Within the same text line use only horizontal distance; otherwise
            # heavily penalise vertical offset so off-line chars are avoided.
            d = dx if dy < char_h * 0.6 else dx * dx + (dy * 4.0) ** 2
            if d < best_dist:
                best_dist = d
                best_idx = i
        return best_idx

    def _copy_selection(self):
        if self._selection_text:
            QApplication.clipboard().setText(self._selection_text)

    # ==================================================================
    # MOUSE EVENTS
    # ==================================================================

    def _on_mouse_down(self, cx: float, cy: float):
        if not self.doc:
            return

        # EXCERTER uses its own rubber-band state, not _drag_start
        if self._tool == Tool.EXCERTER:
            if (self._page_ox <= cx <= self._page_ox + self._page_iw and
                    self._page_oy <= cy <= self._page_oy + self._page_ih):
                self._rb_start = (cx, cy)
                self._rb_current = (cx, cy)
                self._canvas.update()
            return

        self._drag_start = (cx, cy)
        self._drag_current = (cx, cy)
        px, py = self._canvas_to_pdf(cx, cy)

        if self._tool == Tool.FREEHAND:
            self._freehand_pts = [(px, py)]
        elif self._tool == Tool.SIGN:
            self._open_sign_dialog(cx, cy)
        elif self._tool == Tool.TEXT_BOX:
            self._open_textbox_dialog(px, py)
        elif self._tool == Tool.STICKY_NOTE:
            self._open_sticky_dialog(px, py)

    def _on_mouse_move(self, cx: float, cy: float):
        if not self.doc:
            return

        # EXCERTER rubber-band — independent update
        if self._tool == Tool.EXCERTER:
            if self._rb_start is not None:
                self._rb_current = (cx, cy)
                self._canvas.update()
            return

        if self._drag_start is None:
            return
        self._drag_current = (cx, cy)

        mods = QApplication.keyboardModifiers()
        self._shift_held = bool(mods & Qt.KeyboardModifier.ShiftModifier)

        sx, sy = self._drag_start
        px, py = self._canvas_to_pdf(cx, cy)

        if self._tool == Tool.VIEW:
            return

        elif self._tool in (Tool.SELECT, Tool.HIGHLIGHT,
                             Tool.UNDERLINE, Tool.STRIKETHROUGH):
            start_px, start_py = self._canvas_to_pdf(sx, sy)
            self._selected_words = self._select_chars_flow(
                start_px, start_py, px, py)
            self._selection_text = "".join(w[4] for w in self._selected_words)
            self._canvas.update()

        elif self._tool == Tool.FREEHAND:
            self._freehand_pts.append((px, py))
            self._canvas.update()

        elif self._tool in (Tool.RECT, Tool.CIRCLE, Tool.LINE, Tool.ARROW):
            self._canvas.update()

    def _on_mouse_up(self, cx: float, cy: float):
        # EXCERTER uses its own rubber-band state — handle before the _drag_start guard
        if self._tool == Tool.EXCERTER:
            if self._rb_start is not None and self.doc and fitz:
                sx, sy = self._rb_start
                x0c, y0c = min(sx, cx), min(sy, cy)
                x1c, y1c = max(sx, cx), max(sy, cy)
                self._rb_start = None
                self._rb_current = None
                self._canvas.update()
                if x1c - x0c > 4 and y1c - y0c > 4:
                    px0, py0 = self._canvas_to_pdf(x0c, y0c)
                    px1, py1 = self._canvas_to_pdf(x1c, y1c)
                    crop_rect = fitz.Rect(px0, py0, px1, py1)
                    crop_rect.normalize()
                    flash = (x0c, y0c, x1c - x0c, y1c - y0c)
                    self._do_excerpt(crop_rect, flash_rect=flash)
            return

        if not self.doc or self._drag_start is None:
            return

        mods = QApplication.keyboardModifiers()
        self._shift_held = bool(mods & Qt.KeyboardModifier.ShiftModifier)

        sx, sy = self._drag_start
        px, py = self._canvas_to_pdf(cx, cy)
        start_px, start_py = self._canvas_to_pdf(sx, sy)
        self._drag_start = None
        self._drag_current = None

        page = self.doc[self.current_page]
        _, _, fitz_rgb = self._annot_color

        if self._tool == Tool.SELECT:
            pass  # selection done in move

        elif self._tool in (Tool.HIGHLIGHT, Tool.UNDERLINE, Tool.STRIKETHROUGH):
            if self._selected_words:
                self._push_undo()
                # Merge chars into one rect per line (PDF coords) so
                # add_highlight_annot gets one quad per line, not one per
                # character — prevents alpha stacking / dark seams.
                line_rects: dict = {}
                for w in self._selected_words:
                    key = (w[5], w[6])
                    if key not in line_rects:
                        line_rects[key] = [w[0], w[1], w[2], w[3]]
                    else:
                        e = line_rects[key]
                        e[0] = min(e[0], w[0])
                        e[1] = min(e[1], w[1])
                        e[2] = max(e[2], w[2])
                        e[3] = max(e[3], w[3])
                quads = [fitz.Rect(*r).quad for r in line_rects.values()]
                if quads:
                    if self._tool == Tool.HIGHLIGHT:
                        annot = page.add_highlight_annot(quads)
                    elif self._tool == Tool.UNDERLINE:
                        annot = page.add_underline_annot(quads)
                    else:
                        annot = page.add_strikeout_annot(quads)
                    annot.set_colors(stroke=fitz_rgb)
                    annot.update()
                    self._modified = True
                    self._selected_words.clear()
                    self._selection_text = ""
                    self._show_page(self.current_page)

        elif self._tool == Tool.FREEHAND:
            if len(self._freehand_pts) >= 2:
                self._push_undo()
                points = [(float(x), float(y)) for x, y in self._freehand_pts]
                annot = page.add_ink_annot([points])
                annot.set_colors(stroke=fitz_rgb)
                annot.set_border(width=self._stroke_width)
                annot.update()
                self._modified = True
                self._freehand_pts.clear()
                self._show_page(self.current_page)

        elif self._tool == Tool.RECT:
            dx = cx - sx
            dy = cy - sy
            if self._shift_held:
                side = max(abs(dx), abs(dy))
                dx = side if dx >= 0 else -side
                dy = side if dy >= 0 else -side
            end_px, end_py = self._canvas_to_pdf(sx + dx, sy + dy)
            r = fitz.Rect(fitz.Point(start_px, start_py),
                          fitz.Point(end_px, end_py))
            r.normalize()
            if r.width > 2 and r.height > 2:
                self._push_undo()
                annot = page.add_rect_annot(r)
                annot.set_colors(stroke=fitz_rgb)
                annot.set_border(width=self._stroke_width)
                annot.update()
                self._modified = True
                self._show_page(self.current_page)

        elif self._tool == Tool.CIRCLE:
            dx = cx - sx
            dy = cy - sy
            if self._shift_held:
                side = max(abs(dx), abs(dy))
                dx = side if dx >= 0 else -side
                dy = side if dy >= 0 else -side
            end_px, end_py = self._canvas_to_pdf(sx + dx, sy + dy)
            r = fitz.Rect(fitz.Point(start_px, start_py),
                          fitz.Point(end_px, end_py))
            r.normalize()
            if r.width > 2 and r.height > 2:
                self._push_undo()
                annot = page.add_circle_annot(r)
                annot.set_colors(stroke=fitz_rgb)
                annot.set_border(width=self._stroke_width)
                annot.update()
                self._modified = True
                self._show_page(self.current_page)

        elif self._tool in (Tool.LINE, Tool.ARROW):
            dx = cx - sx
            dy = cy - sy
            if self._shift_held:
                angle = math.atan2(dy, dx)
                snap = round(angle / (math.pi / 4)) * (math.pi / 4)
                length = math.hypot(dx, dy)
                dx = length * math.cos(snap)
                dy = length * math.sin(snap)
            end_px, end_py = self._canvas_to_pdf(sx + dx, sy + dy)
            p1 = fitz.Point(start_px, start_py)
            p2 = fitz.Point(end_px, end_py)
            if abs(p1.x - p2.x) > 2 or abs(p1.y - p2.y) > 2:
                self._push_undo()
                annot = page.add_line_annot(p1, p2)
                annot.set_colors(stroke=fitz_rgb)
                annot.set_border(width=self._stroke_width)
                if self._tool == Tool.ARROW:
                    annot.set_line_ends(
                        fitz.PDF_ANNOT_LE_NONE,
                        fitz.PDF_ANNOT_LE_CLOSED_ARROW)
                annot.update()
                self._modified = True
                self._show_page(self.current_page)

    # ==================================================================
    # DOUBLE-CLICK TO EDIT EXISTING TEXT ANNOTATIONS
    # ==================================================================

    def _on_double_click(self, cx: float, cy: float):
        if not self.doc:
            return
        px, py = self._canvas_to_pdf(cx, cy)
        click_pt = fitz.Point(px, py)
        page = self.doc[self.current_page]

        # Word-snap selection for text tools
        if self._tool in (Tool.SELECT, Tool.HIGHLIGHT,
                          Tool.UNDERLINE, Tool.STRIKETHROUGH):
            for w in self._get_words():
                if w[0] <= px <= w[2] and w[1] <= py <= w[3]:
                    self._selected_words = [w]
                    self._selection_text = w[4]
                    self._canvas.update()
                    return

        for annot in page.annots():
            if annot.rect.contains(click_pt):
                atype = annot.type[0]
                if atype == fitz.PDF_ANNOT_FREE_TEXT:
                    self._open_textbox_dialog(px, py, existing_annot=annot)
                    return
                elif atype == fitz.PDF_ANNOT_TEXT:
                    self._edit_sticky(annot)
                    return

    def _edit_sticky(self, annot):
        old_text = annot.info.get("content", "")
        text, ok = QInputDialog.getText(
            self, "Edit Sticky Note", "Edit note text:", text=old_text)
        if not ok:
            return
        self._push_undo()
        page = self.doc[self.current_page]
        if text:
            annot.set_info(content=text)
            annot.update()
        else:
            page.delete_annot(annot)
        self._modified = True
        self._show_page(self.current_page)

    # ==================================================================
    # TEXT BOX & STICKY NOTE
    # ==================================================================

    def _open_textbox_dialog(self, pdf_x, pdf_y, existing_annot=None):
        old_text = ""
        if existing_annot:
            old_text = existing_annot.info.get("content", "")
        title  = "Edit Text Box" if existing_annot else "Add Text Box"
        prompt = "Edit text box:" if existing_annot else "Enter text for the text box:"
        text, ok = QInputDialog.getText(self, title, prompt, text=old_text)
        if not ok:
            return
        if text == "" and not existing_annot:
            return

        self._push_undo()
        page = self.doc[self.current_page]
        _, hex_c, fitz_rgb = self._annot_color
        fontsize = max(8, self._stroke_width * 3)

        if existing_annot:
            old_rect = existing_annot.rect
            page.delete_annot(existing_annot)
            if text:
                width = max(old_rect.width, len(text) * fontsize * 0.6)
                height = max(old_rect.height, fontsize * 2)
                rect = fitz.Rect(old_rect.x0, old_rect.y0,
                                 old_rect.x0 + width, old_rect.y0 + height)
                annot = page.add_freetext_annot(
                    rect, text, fontsize=fontsize,
                    text_color=fitz_rgb, fontname="helv",
                    fill_color=None)
                annot.update()
        else:
            width = max(100, len(text) * fontsize * 0.6)
            height = fontsize * 2.5
            rect = fitz.Rect(pdf_x, pdf_y, pdf_x + width, pdf_y + height)
            annot = page.add_freetext_annot(
                rect, text, fontsize=fontsize,
                text_color=fitz_rgb, fontname="helv",
                fill_color=None)
            annot.update()

        self._modified = True
        self._show_page(self.current_page)

    def _open_sticky_dialog(self, pdf_x, pdf_y):
        text, ok = QInputDialog.getText(self, "Add Sticky Note", "Enter note text:")
        if ok and text:
            self._push_undo()
            page = self.doc[self.current_page]
            point = fitz.Point(pdf_x, pdf_y)
            annot = page.add_text_annot(point, text, icon="Comment")
            _, _, fitz_rgb = self._annot_color
            annot.set_colors(stroke=fitz_rgb)
            annot.update()
            self._modified = True
            self._show_page(self.current_page)

    # ==================================================================
    # SIGNATURE
    # ==================================================================

    def _open_sign_dialog(self, cx: float, cy: float):
        pdf_x, pdf_y = self._canvas_to_pdf(cx, cy)
        dlg = SignatureDialog(self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        png_bytes = dlg.get_png_bytes()
        if not png_bytes:
            return

        self._push_undo()
        # Determine size from the QImage we wrote
        img = QImage.fromData(png_bytes, "PNG")
        sig_w = img.width() * 0.5
        sig_h = img.height() * 0.5
        if sig_w < 1 or sig_h < 1:
            sig_w, sig_h = 100, 50

        page = self.doc[self.current_page]
        rect = fitz.Rect(pdf_x, pdf_y, pdf_x + sig_w, pdf_y + sig_h)
        page.insert_image(rect, stream=png_bytes)
        self._modified = True
        self._show_page(self.current_page)

    # ==================================================================
    # FORM FILLING
    # ==================================================================

    def _clear_form_widgets(self):
        for w in self._form_widgets:
            try:
                w.deleteLater()
            except Exception:
                pass
        self._form_widgets.clear()

    def _draw_form_widgets(self):
        if not self.doc:
            return
        page = self.doc[self.current_page]
        widget_iter = page.widgets()
        if widget_iter is None:
            return

        for widget in widget_iter:
            rect = widget.rect
            x0, y0, x1, y1 = self._pdf_rect_to_canvas(rect)
            w = max(int(x1 - x0), 20)
            h = max(int(y1 - y0), 18)

            if widget.field_type == fitz.PDF_WIDGET_TYPE_TEXT:
                entry = QLineEdit(self._canvas)
                entry.setGeometry(int(x0), int(y0), w, h)
                entry.setStyleSheet(
                    f"QLineEdit {{ background: {WHITE}; border: 1px solid {BLUE}; "
                    f"color: {G900}; font: {max(9, h - 8)}px 'Segoe UI'; }}")
                if widget.field_value:
                    entry.setText(str(widget.field_value))
                entry._pdf_widget = widget
                entry.editingFinished.connect(
                    lambda ent=entry: self._update_form_field(ent))
                entry.show()
                self._form_widgets.append(entry)

            elif widget.field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
                cb = QCheckBox(self._canvas)
                cb.setGeometry(int(x0), int(y0), h, h)
                cb.setChecked(bool(widget.field_value))
                cb._pdf_widget = widget
                cb.stateChanged.connect(
                    lambda state, cbox=cb: self._update_checkbox(cbox))
                cb.show()
                self._form_widgets.append(cb)

            elif widget.field_type == fitz.PDF_WIDGET_TYPE_COMBOBOX:
                choices = widget.choice_values or []
                if choices:
                    combo = QComboBox(self._canvas)
                    combo.setGeometry(int(x0), int(y0), w, h)
                    combo.addItems(choices)
                    combo.setStyleSheet(
                        f"QComboBox {{ background: {WHITE}; border: 1px solid {BLUE}; "
                        f"color: {G900}; }}")
                    if widget.field_value and widget.field_value in choices:
                        combo.setCurrentText(widget.field_value)
                    combo._pdf_widget = widget
                    combo.currentTextChanged.connect(
                        lambda val, c=combo: self._update_combo(c, val))
                    combo.show()
                    self._form_widgets.append(combo)

    def _update_form_field(self, entry):
        self._push_undo()
        widget = entry._pdf_widget
        widget.field_value = entry.text()
        widget.update()
        self._modified = True

    def _update_checkbox(self, cb):
        self._push_undo()
        widget = cb._pdf_widget
        widget.field_value = cb.isChecked()
        widget.update()
        self._modified = True

    def _update_combo(self, combo, val):
        self._push_undo()
        widget = combo._pdf_widget
        widget.field_value = val
        widget.update()
        self._modified = True

    # ==================================================================
    # SAVE & PRINT
    # ==================================================================

    def _save_pdf(self):
        if not self.doc:
            return
        initial = Path(self.pdf_path).name if self.pdf_path else "output.pdf"
        path, _ = QFileDialog.getSaveFileName(
            self, "Save PDF", initial, "PDF Files (*.pdf)")
        if not path:
            return
        try:
            if path == self.pdf_path:
                self.doc.save(path, incremental=True,
                              encryption=fitz.PDF_ENCRYPT_KEEP)
            else:
                self.doc.save(path)
            self._modified = False
            QMessageBox.information(self, "Saved", f"PDF saved to:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not save:\n{e}")

    def _print_pdf(self):
        if not self.doc:
            return
        try:
            tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            tmp_path = tmp.name
            tmp.close()
            self.doc.save(tmp_path)
            if os.name == "nt":
                os.startfile(tmp_path, "print")
            elif os.name == "posix":
                subprocess.run(["lpr", tmp_path], check=True)
            else:
                QMessageBox.information(
                    self, "Print",
                    f"PDF saved to {tmp_path}\nPlease print it manually.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Print failed:\n{e}")

    # ==================================================================
    # CLEANUP
    # ==================================================================

    def cleanup(self):
        if self._thumb_timer is not None:
            self._thumb_timer = None
        self._clear_form_widgets()
        if self.doc:
            self.doc.close()
            self.doc = None

    def closeEvent(self, event):
        self.cleanup()
        super().closeEvent(event)
