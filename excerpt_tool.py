"""Excerpt Tool – Multi-document rubber-band region capture to a new PDF.

Load multiple PDFs, drag to select rectangular regions on any page,
and collect them into a growing excerpt document. Save at any time.
Uses native PDF crop (show_pdf_page with clip) to preserve text and links.

PySide6 port – no tkinter/customtkinter dependencies.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, cast

from PySide6.QtWidgets import (
    QWidget, QFrame, QLabel, QPushButton, QLineEdit,
    QScrollArea, QHBoxLayout, QVBoxLayout, QFileDialog, QMessageBox,
    QSizePolicy, QApplication,
)
from PySide6.QtCore import Qt, QTimer, QEvent, QObject, QRect, QPoint, QSize
from PySide6.QtGui import (
    QPainter, QColor, QPixmap, QPen, QPainterPath,
    QFont, QCursor, QBrush, QIcon,
)
from icons import svg_pixmap, svg_icon
from colors import (
    BG, WHITE, G100, G200, G300, G400, G500, G700, G800, G900,
    BLUE, BLUE_DIM, BLUE_DARK, BLUE_HOVER, BLUE_MED,
    GREEN, GREEN_HOVER, GREEN_TXT, RED, SEL_BLUE,
    SIDEBAR_BG, CARD_BG,
)
from utils import _fitz_pix_to_qpixmap, _WheelToHScroll

try:
    import fitz  # pymupdf
except ImportError:
    fitz = None


# ---------------------------------------------------------------------------
# Snippet data structure
# ---------------------------------------------------------------------------

@dataclass
class Snippet:
    source_path: str          # absolute path to source PDF
    page_index:  int          # 0-based page index within that PDF
    crop_rect:   Any          # fitz.Rect in PDF coordinate space
    label:       str  = ""
    thumbnail:   QPixmap | None = field(default=None, repr=False)  # QPixmap ref


# ---------------------------------------------------------------------------
# ExcerptCanvas – custom page preview widget with rubber-band selection
# ---------------------------------------------------------------------------

class ExcerptCanvas(QWidget):
    """Renders a single PDF page pixmap with drop-shadow; supports
    rubber-band (drag) region selection."""

    def __init__(self, et: "ExcerptTool", parent=None):
        super().__init__(parent)
        self._et = et
        self.setMouseTracking(True)
        self.setCursor(QCursor(Qt.CursorShape.CrossCursor))
        self.setMinimumSize(300, 300)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)

    # ------------------------------------------------------------------ paint
    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing, False)

        et = self._et

        # Background
        p.fillRect(self.rect(), QColor(BG))

        if et._preview_pixmap is None:
            # Placeholder text
            p.setPen(QColor(G400))
            f = QFont("Segoe UI", 14)
            p.setFont(f)
            p.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter,
                       "Add a PDF, then drag to select a region")
            return

        ox = int(et._page_ox)
        oy = int(et._page_oy)
        iw = et._preview_pixmap.width()
        ih = et._preview_pixmap.height()

        # Soft drop shadow (subtle, offset 3-4px)
        p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        shadow_color = QColor(0, 0, 0, 20)
        p.fillRect(ox + 4, oy + 4, iw, ih, shadow_color)
        p.fillRect(ox + 3, oy + 3, iw, ih, QColor(0, 0, 0, 12))
        p.setRenderHint(QPainter.RenderHint.Antialiasing, False)

        # Page image
        p.drawPixmap(ox, oy, et._preview_pixmap)

        # Page border
        p.setPen(QPen(QColor(G200), 1))
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawRect(ox - 1, oy - 1, iw + 1, ih + 1)

        # Rubber-band overlay
        if et._rb_start is not None and et._rb_current is not None:
            sx, sy = et._rb_start
            cx, cy = et._rb_current
            rx = int(min(sx, cx))
            ry = int(min(sy, cy))
            rw = int(abs(cx - sx))
            rh = int(abs(cy - sy))

            # Fill with translucent blue
            p.fillRect(rx, ry, rw, rh, QColor(59, 130, 246, 20))

            # Dashed blue border
            pen = QPen(QColor(SEL_BLUE), 2, Qt.PenStyle.DashLine)
            p.setPen(pen)
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawRect(rx, ry, rw, rh)

            # Corner handles (12x12 white circles with blue border)
            p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
            handle_pen = QPen(QColor(BLUE_DARK), 2)
            p.setPen(handle_pen)
            p.setBrush(QBrush(QColor(WHITE)))
            x0, y0, x1, y1 = rx, ry, rx + rw, ry + rh
            for hx, hy in [(x0 - 6, y0 - 6), (x1 - 6, y0 - 6),
                           (x0 - 6, y1 - 6), (x1 - 6, y1 - 6)]:
                p.drawEllipse(hx, hy, 12, 12)
            p.setRenderHint(QPainter.RenderHint.Antialiasing, False)

        # Flash feedback (blue tint overlay)
        if et._flash_rect is not None:
            fx, fy, fw, fh = et._flash_rect
            p.fillRect(QRect(fx, fy, fw, fh), QColor(59, 130, 246, 50))
            p.setPen(QPen(QColor(BLUE), 2))
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawRect(QRect(fx, fy, fw, fh))

    # -------------------------------------------------------------- mouse events
    def mousePressEvent(self, event):
        if event.button() != Qt.MouseButton.LeftButton:
            return
        et = self._et
        if et._active_doc is None or et._preview_pixmap is None:
            return
        cx = float(event.position().x())
        cy = float(event.position().y())
        if not et._point_on_page(cx, cy):
            return
        et._rb_start = (cx, cy)
        et._rb_current = (cx, cy)
        self.update()

    def mouseMoveEvent(self, event):
        et = self._et
        if et._rb_start is None:
            return
        cx = float(event.position().x())
        cy = float(event.position().y())
        et._rb_current = (cx, cy)
        self.update()

    def mouseReleaseEvent(self, event):
        if event.button() != Qt.MouseButton.LeftButton:
            return
        et = self._et
        if et._rb_start is None:
            return
        cx = float(event.position().x())
        cy = float(event.position().y())
        sx, sy = et._rb_start
        et._rb_start = None
        et._rb_current = None
        self.update()

        # Reject tiny drags
        if abs(cx - sx) < 8 or abs(cy - sy) < 8:
            return

        # Clamp to page bounds
        x0 = max(min(sx, cx), et._page_ox)
        y0 = max(min(sy, cy), et._page_oy)
        x1 = min(max(sx, cx), et._page_ox + et._page_iw)
        y1 = min(max(sy, cy), et._page_oy + et._page_ih)
        if x1 <= x0 or y1 <= y0:
            return

        crop_rect = et._canvas_to_pdf_rect(x0, y0, x1, y1)
        if crop_rect.is_empty or crop_rect.width < 2 or crop_rect.height < 2:
            return

        et._do_capture(crop_rect)
        et._flash_feedback(x0, y0, x1, y1)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        et = self._et
        if et._active_doc is not None and et._preview_pixmap is not None:
            et._render_page()

    def wheelEvent(self, event):
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            delta = event.angleDelta().y()
            if delta > 0:
                self._et._zoom_in()
            elif delta < 0:
                self._et._zoom_out()
            event.accept()
        else:
            super().wheelEvent(event)

    def keyPressEvent(self, event):
        self._et.keyPressEvent(event)


# ===========================================================================
# ExcerptTool – main widget
# ===========================================================================

class ExcerptTool(QWidget):
    THUMB_W = 80    # page thumbnail width in bottom strip
    SNIP_W  = 60    # snippet thumbnail width in left panel list
    LEFT_W  = 320   # fixed left panel width in pixels

    def __init__(self, parent=None):
        super().__init__(parent)

        if fitz is None:
            lay = QVBoxLayout(self)
            lbl = QLabel("Missing dependencies.\n\npip install pymupdf")
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl.setStyleSheet(f"color: {G500}; font: 16px 'Segoe UI';")
            lay.addWidget(lbl)
            return

        # ---- Multi-document state ----
        # Each entry: {"path": str, "doc": fitz.Document, "name": str}
        self._pdf_list: list  = []
        self._active_idx: int = -1

        # ---- Per-active-doc view state ----
        self._current_page: int   = 0
        self._page_ox: float      = 0.0
        self._page_oy: float      = 0.0
        self._page_iw: float      = 0.0
        self._page_ih: float      = 0.0
        self._render_mat          = fitz.Matrix(1, 1)
        self._inv_mat             = fitz.Matrix(1, 1)
        self._preview_pixmap: QPixmap | None = None
        self._zoom_factor: float = 1.0
        self._zoom_lbl = None

        # Thumb strip state
        self._thumb_pixmaps: list  = []   # list of QPixmap | None
        self._thumb_frames: list   = []   # list of QFrame widgets
        self._thumb_render_next: int = 0
        self._highlighted_thumb: int = -1
        self._thumb_timer: QTimer | None = None

        # ---- Rubber-band state ----
        self._rb_start: tuple | None   = None
        self._rb_current: tuple | None = None

        # Flash feedback state
        self._flash_rect: tuple | None = None

        # ---- Snippets & output doc ----
        self._snippets: list = []
        self._out_doc        = fitz.open()   # empty in-memory PDF
        self._out_y_cursor   = 0.0           # current y position on active A4 page
        self._out_has_page   = False         # whether an A4 page exists yet

        # Widget refs set during _build_ui; initialised here so methods that
        # reference them don't need hasattr() guards.
        self._active_file_lbl: QLabel | None = None
        self._page_nav_lbl: QLabel | None    = None

        self._build_ui()

    # ==========================================================================
    # BUILD UI
    # ==========================================================================

    def _build_ui(self):
        root_lay = QVBoxLayout(self)
        root_lay.setContentsMargins(0, 0, 0, 0)
        root_lay.setSpacing(0)

        # ---- Top area: left panel + right preview ----
        top = QWidget()
        top_lay = QHBoxLayout(top)
        top_lay.setContentsMargins(0, 0, 0, 0)
        top_lay.setSpacing(0)

        # Left panel (fixed width)
        left_frame = QFrame()
        left_frame.setFixedWidth(self.LEFT_W)
        left_frame.setStyleSheet(
            f"QFrame {{ background: {SIDEBAR_BG}; border-right: 1px solid {G200}; }}"
        )
        left_frame.setSizePolicy(QSizePolicy.Policy.Fixed,
                                  QSizePolicy.Policy.Expanding)
        self._left_frame = left_frame
        self._build_left_panel(left_frame)
        top_lay.addWidget(left_frame)

        # Right panel (expandable)
        right_frame = QFrame()
        right_frame.setObjectName("RightFrame")
        right_frame.setStyleSheet(f"QFrame#RightFrame {{ background: {BG}; border: none; }}")
        right_frame.setSizePolicy(QSizePolicy.Policy.Expanding,
                                   QSizePolicy.Policy.Expanding)
        self._right_frame = right_frame
        self._build_right_panel(right_frame)
        top_lay.addWidget(right_frame, 1)

        root_lay.addWidget(top, 1)

        # ---- Bottom thumbnail strip ----
        self._build_thumb_strip(root_lay)

    # --------------------------------------------------------------------------
    # Left panel
    # --------------------------------------------------------------------------

    def _build_left_panel(self, left: QFrame):
        lay = QVBoxLayout(left)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        inner = QWidget()
        inner.setStyleSheet("background: transparent;")
        inner_lay = QVBoxLayout(inner)
        inner_lay.setContentsMargins(16, 16, 16, 16)
        inner_lay.setSpacing(0)

        # ---- Active file card ----
        self._file_card = QFrame()
        self._file_card.setStyleSheet(
            f"QFrame {{ background: {WHITE}; border: 1px solid {G200}; "
            f"border-radius: 12px; }}"
        )
        file_card_lay = QHBoxLayout(self._file_card)
        file_card_lay.setContentsMargins(12, 12, 12, 12)
        file_card_lay.setSpacing(10)

        # Blue icon bg
        icon_bg = QFrame()
        icon_bg.setFixedSize(40, 40)
        icon_bg.setStyleSheet(
            f"QFrame {{ background: {BLUE_DIM}; border-radius: 8px; border: none; }}"
        )
        icon_bg_lay = QVBoxLayout(icon_bg)
        icon_bg_lay.setContentsMargins(0, 0, 0, 0)
        icon_lbl = QLabel()
        icon_lbl.setPixmap(svg_pixmap("file-text", BLUE_DARK, 20))
        icon_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_lbl.setStyleSheet("background: transparent; border: none;")
        icon_bg_lay.addWidget(icon_lbl, 0, Qt.AlignmentFlag.AlignCenter)
        file_card_lay.addWidget(icon_bg)

        # File name label
        self._active_file_lbl = QLabel("No file loaded")
        self._active_file_lbl.setStyleSheet(
            f"color: {G800}; font: bold 13px 'Segoe UI'; "
            f"background: transparent; border: none;"
        )
        self._active_file_lbl.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        file_card_lay.addWidget(self._active_file_lbl, 1)

        # Browse button
        self._browse_lbl = QPushButton("Browse")
        self._browse_lbl.setStyleSheet(
            f"QPushButton {{ background: transparent; color: {BLUE_DARK}; "
            f"font: 12px 'Segoe UI'; border: none; padding: 0; }}"
            f"QPushButton:hover {{ color: {BLUE}; }}"
        )
        self._browse_lbl.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self._browse_lbl.clicked.connect(self._add_pdf)
        file_card_lay.addWidget(self._browse_lbl)

        inner_lay.addWidget(self._file_card)
        inner_lay.addSpacing(16)

        # ---- LOAD SOURCE section ----
        load_hdr = QLabel("LOAD SOURCE")
        load_hdr.setStyleSheet(
            f"color: {G400}; font: bold 10px 'Segoe UI'; "
            f"letter-spacing: 1px; background: transparent;"
        )
        inner_lay.addWidget(load_hdr)
        inner_lay.addSpacing(8)

        self._add_btn = QPushButton("+ Add PDF")
        self._add_btn.setFixedHeight(40)
        self._add_btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self._add_btn.setStyleSheet(
            f"QPushButton {{ background: {BLUE_DARK}; color: {WHITE}; "
            f"font: bold 13px 'Segoe UI'; border: none; border-radius: 8px; }}"
            f"QPushButton:hover {{ background: #1D4ED8; }}"
        )
        self._add_btn.clicked.connect(self._add_pdf)
        inner_lay.addWidget(self._add_btn)
        inner_lay.addSpacing(6)

        size_lbl = QLabel("Max file size: 50MB")
        size_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        size_lbl.setStyleSheet(
            f"color: {G400}; font: 11px 'Segoe UI'; background: transparent;"
        )
        inner_lay.addWidget(size_lbl)
        inner_lay.addSpacing(16)

        # ---- LOADED PDFS section ----
        pdfs_hdr_row = QWidget()
        pdfs_hdr_row.setStyleSheet("background: transparent;")
        pdfs_hdr_lay = QHBoxLayout(pdfs_hdr_row)
        pdfs_hdr_lay.setContentsMargins(0, 0, 0, 0)
        pdfs_hdr_lay.setSpacing(4)

        pdfs_hdr_lbl = QLabel("LOADED PDFS")
        pdfs_hdr_lbl.setStyleSheet(
            f"color: {G400}; font: bold 10px 'Segoe UI'; "
            f"letter-spacing: 1px; background: transparent;"
        )
        pdfs_hdr_lay.addWidget(pdfs_hdr_lbl)
        pdfs_hdr_lay.addStretch()

        rem_all_btn = QPushButton("REMOVE ALL")
        rem_all_btn.setStyleSheet(
            f"QPushButton {{ background: transparent; color: {RED}; "
            f"font: bold 10px 'Segoe UI'; border: none; padding: 0; }}"
            f"QPushButton:hover {{ color: #DC2626; }}"
        )
        rem_all_btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        rem_all_btn.clicked.connect(self._remove_all_pdfs)
        pdfs_hdr_lay.addWidget(rem_all_btn)

        inner_lay.addWidget(pdfs_hdr_row)
        inner_lay.addSpacing(6)

        # PDF list scroll area
        pdf_scroll = QScrollArea()
        pdf_scroll.setFixedHeight(100)
        pdf_scroll.setWidgetResizable(True)
        pdf_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        pdf_scroll.setStyleSheet(
            f"QScrollArea {{ border: none; background: transparent; }}"
        )

        pdf_inner = QWidget()
        pdf_inner.setStyleSheet("background: transparent;")
        self._pdf_list_layout = QVBoxLayout(pdf_inner)
        self._pdf_list_layout.setContentsMargins(0, 0, 0, 0)
        self._pdf_list_layout.setSpacing(4)
        self._pdf_list_layout.addStretch()

        pdf_scroll.setWidget(pdf_inner)
        inner_lay.addWidget(pdf_scroll)
        inner_lay.addSpacing(16)

        # ---- CAPTURED SNIPPETS section ----
        snip_hdr_row = QWidget()
        snip_hdr_row.setStyleSheet("background: transparent;")
        snip_hdr_lay = QHBoxLayout(snip_hdr_row)
        snip_hdr_lay.setContentsMargins(0, 0, 0, 0)
        snip_hdr_lay.setSpacing(4)

        snip_hdr_lbl = QLabel("CAPTURED SNIPPETS")
        snip_hdr_lbl.setStyleSheet(
            f"color: {G400}; font: bold 10px 'Segoe UI'; "
            f"letter-spacing: 1px; background: transparent;"
        )
        snip_hdr_lay.addWidget(snip_hdr_lbl)

        self._snip_count_lbl = QLabel("(0)")
        self._snip_count_lbl.setStyleSheet(
            f"color: {G400}; font: bold 10px 'Segoe UI'; background: transparent;"
        )
        snip_hdr_lay.addWidget(self._snip_count_lbl)
        snip_hdr_lay.addStretch()

        inner_lay.addWidget(snip_hdr_row)
        inner_lay.addSpacing(8)

        # Snippet list scroll area
        snip_scroll = QScrollArea()
        snip_scroll.setWidgetResizable(True)
        snip_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        snip_scroll.setStyleSheet(
            f"QScrollArea {{ border: none; background: transparent; }}"
        )

        snip_inner = QWidget()
        snip_inner.setStyleSheet("background: transparent;")
        self._snip_layout = QVBoxLayout(snip_inner)
        self._snip_layout.setContentsMargins(0, 0, 0, 0)
        self._snip_layout.setSpacing(6)
        self._snip_layout.addStretch()

        snip_scroll.setWidget(snip_inner)
        inner_lay.addWidget(snip_scroll, 1)

        # Spacer
        inner_lay.addSpacing(8)

        # ---- Save button ----
        self._save_btn = QPushButton("Save Excerpt PDF")
        self._save_btn.setFixedHeight(48)
        self._save_btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self._save_btn.setStyleSheet(
            f"QPushButton {{ background: {GREEN}; color: {WHITE}; "
            f"font: bold 14px 'Segoe UI'; border: none; border-radius: 12px; }}"
            f"QPushButton:hover {{ background: {GREEN_HOVER}; }}"
        )
        self._save_btn.clicked.connect(self._save_excerpt)
        inner_lay.addWidget(self._save_btn)
        inner_lay.addSpacing(8)

        self._status_lbl = QLabel("")
        self._status_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._status_lbl.setStyleSheet(
            f"color: {GREEN_TXT}; font: 11px 'Segoe UI'; background: transparent;"
        )
        inner_lay.addWidget(self._status_lbl)

        lay.addWidget(inner)

    # --------------------------------------------------------------------------
    # Right panel
    # --------------------------------------------------------------------------

    def _build_right_panel(self, right: QFrame):
        lay = QVBoxLayout(right)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        # Canvas scroll area
        self._canvas_scroll = QScrollArea()
        self._canvas_scroll.setWidgetResizable(False)
        self._canvas_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self._canvas_scroll.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self._canvas_scroll.setStyleSheet(
            f"QScrollArea {{ border: none; background: transparent; }}"
        )

        self._canvas = ExcerptCanvas(self)
        self._canvas.setFixedSize(400, 400)
        self._canvas_scroll.setWidget(self._canvas)

        lay.addWidget(self._canvas_scroll, 1)

        # Navigation bar (frosted footer)
        nav = QFrame()
        nav.setFixedHeight(48)
        nav.setStyleSheet(
            f"QFrame {{ background: {SIDEBAR_BG}; border-top: 1px solid {G200}; border-left: none; border-right: none; border-bottom: none; }}"
        )
        nav_lay = QHBoxLayout(nav)
        nav_lay.setContentsMargins(16, 0, 16, 0)
        nav_lay.setSpacing(6)

        self._sidebar_toggle_btn = QPushButton()
        self._sidebar_toggle_btn.setIcon(QIcon(svg_pixmap("chevron-left", G700, 16)))
        self._sidebar_toggle_btn.setIconSize(QSize(16, 16))
        self._sidebar_toggle_btn.setFixedSize(32, 32)
        self._sidebar_toggle_btn.setToolTip("Hide sidebar")
        self._sidebar_toggle_btn.setStyleSheet(
            f"QPushButton {{ background: {G100}; border: 1px solid {G200}; border-radius: 6px; }}"
            f"QPushButton:hover {{ background: {G200}; }}"
        )
        self._sidebar_toggle_btn.clicked.connect(self._toggle_sidebar)
        nav_lay.addWidget(self._sidebar_toggle_btn)

        nav_lay.addStretch()

        self.btn_prev = QPushButton()
        self.btn_prev.setIcon(QIcon(svg_pixmap("chevron-left", G700, 14)))
        self.btn_prev.setIconSize(QSize(14, 14))
        self.btn_prev.setFixedSize(32, 32)
        self.btn_prev.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.btn_prev.setEnabled(False)
        self.btn_prev.setStyleSheet(
            f"QPushButton {{ background: {G100}; border: 1px solid {G200}; border-radius: 6px; }}"
            f"QPushButton:hover:enabled {{ background: {G200}; }}"
            f"QPushButton:disabled {{ background: {G100}; }}"
        )
        self.btn_prev.clicked.connect(self._prev_page)
        nav_lay.addWidget(self.btn_prev)

        self.page_entry = QLineEdit("–")
        self.page_entry.setFixedSize(52, 32)
        self.page_entry.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.page_entry.setStyleSheet(
            f"QLineEdit {{ background: {WHITE}; color: {G900}; "
            f"font: 12px 'Segoe UI'; border: 1px solid {G200}; border-radius: 6px; }}"
        )
        self.page_entry.returnPressed.connect(self._goto_page)
        nav_lay.addWidget(self.page_entry)

        self.total_lbl = QLabel("/ –")
        self.total_lbl.setStyleSheet(
            f"color: {G500}; font: 12px 'Segoe UI'; background: transparent;"
        )
        nav_lay.addWidget(self.total_lbl)

        self.btn_next = QPushButton()
        self.btn_next.setIcon(QIcon(svg_pixmap("chevron-right", G700, 14)))
        self.btn_next.setIconSize(QSize(14, 14))
        self.btn_next.setFixedSize(32, 32)
        self.btn_next.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.btn_next.setEnabled(False)
        self.btn_next.setStyleSheet(
            f"QPushButton {{ background: {G100}; border: 1px solid {G200}; border-radius: 6px; }}"
            f"QPushButton:hover:enabled {{ background: {G200}; }}"
            f"QPushButton:disabled {{ background: {G100}; }}"
        )
        self.btn_next.clicked.connect(self._next_page)
        nav_lay.addWidget(self.btn_next)

        nav_lay.addStretch()
        lay.addWidget(nav)


    # --------------------------------------------------------------------------
    # Bottom thumbnail strip
    # --------------------------------------------------------------------------

    def _build_thumb_strip(self, root_lay: QVBoxLayout):
        strip_container = QWidget()
        strip_container.setFixedHeight(200)
        strip_container.setStyleSheet(
            f"background: #F8FAFC; border-top: 1px solid {G200};"
        )
        strip_outer_lay = QVBoxLayout(strip_container)
        strip_outer_lay.setContentsMargins(0, 0, 0, 0)
        strip_outer_lay.setSpacing(0)

        # Header row
        strip_hdr_row = QWidget()
        strip_hdr_row.setStyleSheet("background: transparent;")
        strip_hdr_lay = QHBoxLayout(strip_hdr_row)
        strip_hdr_lay.setContentsMargins(12, 6, 12, 4)
        strip_hdr_lay.setSpacing(8)

        nav_lbl = QLabel("PAGE NAVIGATOR")
        nav_lbl.setStyleSheet(
            f"color: {G400}; font: bold 10px 'Segoe UI'; "
            f"letter-spacing: 1px; background: transparent;"
        )
        strip_hdr_lay.addWidget(nav_lbl)
        strip_hdr_lay.addStretch()

        self._page_nav_lbl = QLabel("Page 1 / 1")
        self._page_nav_lbl.setStyleSheet(
            f"color: {G700}; font: bold 12px 'Segoe UI'; background: transparent;"
        )
        strip_hdr_lay.addWidget(self._page_nav_lbl)

        strip_outer_lay.addWidget(strip_hdr_row)

        # Scroll row
        strip_lay = QHBoxLayout()
        strip_lay.setContentsMargins(0, 0, 0, 0)
        strip_lay.setSpacing(0)

        # Left arrow
        self.btn_tl = QPushButton("‹")
        self.btn_tl.setFixedSize(28, 100)
        self.btn_tl.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.btn_tl.setStyleSheet(
            f"QPushButton {{ background: transparent; color: {G400}; "
            f"font: bold 20px 'Segoe UI'; border: none; border-radius: 8px; }}"
            f"QPushButton:hover {{ background: {G200}; }}"
        )
        strip_lay.addWidget(self.btn_tl)

        # Thumbnail scroll area
        self._thumb_sa = QScrollArea()
        self._thumb_sa.setFixedHeight(160)
        self._thumb_sa.setWidgetResizable(True)
        self._thumb_sa.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._thumb_sa.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self._thumb_sa.setStyleSheet(
            f"QScrollArea {{ border: none; background: transparent; }}"
        )
        self._wheel_filter = _WheelToHScroll(self._thumb_sa)

        # Inner widget for thumbnails
        self._thumb_inner = QWidget()
        self._thumb_inner.setStyleSheet("background: transparent;")
        self._thumb_inner_lay = QHBoxLayout(self._thumb_inner)
        self._thumb_inner_lay.setContentsMargins(14, 6, 14, 6)
        self._thumb_inner_lay.setSpacing(10)
        self._thumb_inner_lay.addStretch()

        self._thumb_sa.setWidget(self._thumb_inner)
        strip_lay.addWidget(self._thumb_sa, 1)

        # Right arrow
        self.btn_tr = QPushButton("›")
        self.btn_tr.setFixedSize(28, 100)
        self.btn_tr.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.btn_tr.setStyleSheet(
            f"QPushButton {{ background: transparent; color: {G400}; "
            f"font: bold 20px 'Segoe UI'; border: none; border-radius: 8px; }}"
            f"QPushButton:hover {{ background: {G200}; }}"
        )
        strip_lay.addWidget(self.btn_tr)

        strip_outer_lay.addLayout(strip_lay)

        self.btn_tl.clicked.connect(
            lambda: self._thumb_sa.horizontalScrollBar().setValue(
                self._thumb_sa.horizontalScrollBar().value() - 3 * self.THUMB_W))
        self.btn_tr.clicked.connect(
            lambda: self._thumb_sa.horizontalScrollBar().setValue(
                self._thumb_sa.horizontalScrollBar().value() + 3 * self.THUMB_W))

        root_lay.addWidget(strip_container)

    # ==========================================================================
    # PDF LIST MANAGEMENT
    # ==========================================================================

    def _add_pdf(self):
        if fitz is None:
            QMessageBox.critical(
                self, "Error", "PyMuPDF (fitz) is not installed.")
            return
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Add PDF(s)", "", "PDF Files (*.pdf)")
        if not paths:
            return
        added = 0
        for p in paths:
            if any(d["path"] == p for d in self._pdf_list):
                continue
            try:
                doc = fitz.open(p)
                self._pdf_list.append({
                    "path": p,
                    "doc":  doc,
                    "name": Path(p).name,
                })
                added += 1
            except Exception as e:
                QMessageBox.critical(
                    self, "Error",
                    f"Could not open {Path(p).name}:\n{e}")
        if added:
            if self._active_idx == -1:
                self._set_active_pdf(0)
            else:
                self._rebuild_pdf_list()

    def _remove_active_pdf(self):
        if self._active_idx < 0:
            return
        entry = self._pdf_list.pop(self._active_idx)
        try:
            entry["doc"].close()
        except Exception:
            pass
        self._active_idx = -1
        if self._pdf_list:
            self._set_active_pdf(0)
        else:
            self._stop_thumb_timer()
            self._thumb_pixmaps.clear()
            self._thumb_frames.clear()
            self._clear_thumb_strip()
            self._preview_pixmap = None
            self._canvas.setFixedSize(400, 400)
            self._canvas.update()
            self.page_entry.setText("–")
            self.total_lbl.setText("/ –")
            self.btn_prev.setEnabled(False)
            self.btn_next.setEnabled(False)
            self._rebuild_pdf_list()

    def _remove_all_pdfs(self):
        for entry in self._pdf_list:
            try:
                entry["doc"].close()
            except Exception:
                pass
        self._pdf_list.clear()
        self._active_idx = -1
        self._stop_thumb_timer()
        self._thumb_pixmaps.clear()
        self._thumb_frames.clear()
        self._clear_thumb_strip()
        self._preview_pixmap = None
        self._canvas.setFixedSize(400, 400)
        self._canvas.update()
        self.page_entry.setText("–")
        self.total_lbl.setText("/ –")
        self.btn_prev.setEnabled(False)
        self.btn_next.setEnabled(False)
        self._rebuild_pdf_list()

    def _set_active_pdf(self, idx: int):
        if idx < 0 or idx >= len(self._pdf_list):
            return
        self._active_idx = idx
        self._current_page = 0
        self._rebuild_pdf_list()
        self._render_thumbs()
        self._show_page(0)

    def _rebuild_pdf_list(self):
        # Remove all cards (keep stretch at end: last item)
        while self._pdf_list_layout.count() > 1:
            item = self._pdf_list_layout.takeAt(0)
            if item is None:
                break
            w = item.widget()
            if w is not None:
                w.deleteLater()

        for i, entry in enumerate(self._pdf_list):
            is_active = (i == self._active_idx)
            card = QFrame()
            card.setFixedHeight(32)
            card.setStyleSheet(
                f"QFrame {{ background: {BLUE_DIM if is_active else WHITE}; "
                f"border: 1px solid {'#BFDBFE' if is_active else G200}; "
                f"border-radius: 8px; }}")
            card.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

            card_lay = QHBoxLayout(card)
            card_lay.setContentsMargins(10, 0, 8, 0)
            card_lay.setSpacing(4)

            name_lbl = QLabel(entry["name"])
            name_lbl.setStyleSheet(
                f"color: {BLUE_DARK if is_active else G700}; "
                f"font: {'bold ' if is_active else ''}12px 'Segoe UI'; "
                f"border: none; background: transparent;"
            )
            name_lbl.setSizePolicy(QSizePolicy.Policy.Expanding,
                                    QSizePolicy.Policy.Preferred)
            card_lay.addWidget(name_lbl, 1)

            pg_lbl = QLabel(f"{len(entry['doc'])}p")
            pg_lbl.setStyleSheet(
                f"color: {G400}; font: 10px 'Segoe UI'; "
                f"border: none; background: transparent;"
            )
            card_lay.addWidget(pg_lbl)

            # Click anywhere on card activates it
            card.mousePressEvent = lambda e, ii=i: self._set_active_pdf(ii)
            name_lbl.mousePressEvent = lambda e, ii=i: self._set_active_pdf(ii)
            pg_lbl.mousePressEvent = lambda e, ii=i: self._set_active_pdf(ii)

            self._pdf_list_layout.insertWidget(i, card)

        # Update active file card label
        if self._active_file_lbl is not None:
            if self._active_idx >= 0 and self._pdf_list:
                name = self._pdf_list[self._active_idx]['name']
                if len(name) > 28:
                    name = name[:25] + '...'
                self._active_file_lbl.setText(name)
            else:
                self._active_file_lbl.setText("No file loaded")

    # ==========================================================================
    # PAGE RENDERING
    # ==========================================================================

    @property
    def _active_doc(self):
        if self._active_idx < 0 or self._active_idx >= len(self._pdf_list):
            return None
        return self._pdf_list[self._active_idx]["doc"]

    def _show_page(self, idx: int):
        doc = self._active_doc
        if doc is None or idx < 0 or idx >= len(doc):
            return
        self._current_page = idx
        self._render_page()
        total = len(doc)
        self.page_entry.setText(str(idx + 1))
        self.total_lbl.setText(f"/ {total}")
        self.btn_prev.setEnabled(idx > 0)
        self.btn_next.setEnabled(idx < total - 1)
        self._hl_thumb(idx)
        if self._page_nav_lbl is not None:
            self._page_nav_lbl.setText(f"Page {idx + 1} / {total}")

    def _render_page(self):
        doc = self._active_doc
        if doc is None:
            return

        cw = max(self._canvas_scroll.viewport().width(), 300)
        ch = max(self._canvas_scroll.viewport().height(), 300)

        page  = doc[self._current_page]
        pw    = page.rect.width
        scale = max((cw - 40) / pw, 0.05) * self._zoom_factor

        mat              = fitz.Matrix(scale, scale)
        self._render_mat = mat
        self._inv_mat    = ~mat

        pix = page.get_pixmap(matrix=mat, alpha=False)
        self._preview_pixmap = _fitz_pix_to_qpixmap(pix)
        iw = self._preview_pixmap.width()
        ih = self._preview_pixmap.height()

        self._page_ox = (cw - iw) / 2
        self._page_oy = 20.0
        self._page_iw = float(iw)
        self._page_ih = float(ih)

        self._canvas.setFixedSize(cw, max(ih + 60, ch))
        self._canvas.update()
        if self._zoom_lbl is not None:
            self._zoom_lbl.setText(f"{int(self._zoom_factor * 100)}%")

    def _toggle_sidebar(self):
        if self._left_frame.isVisible():
            self._left_frame.hide()
            self._sidebar_toggle_btn.setIcon(QIcon(svg_pixmap("chevron-right", G700, 16)))
            self._sidebar_toggle_btn.setToolTip("Show sidebar")
        else:
            self._left_frame.show()
            self._sidebar_toggle_btn.setIcon(QIcon(svg_pixmap("chevron-left", G700, 16)))
            self._sidebar_toggle_btn.setToolTip("Hide sidebar")

    def _zoom_in(self):
        self._zoom_factor = min(self._zoom_factor * 1.25, 5.0)
        self._render_page()

    def _zoom_out(self):
        self._zoom_factor = max(self._zoom_factor / 1.25, 0.1)
        self._render_page()

    def keyPressEvent(self, event):
        mod = event.modifiers()
        key = event.key()
        if mod & Qt.KeyboardModifier.ControlModifier:
            if key in (Qt.Key.Key_Plus, Qt.Key.Key_Equal):
                self._zoom_in()
                return
            if key == Qt.Key.Key_Minus:
                self._zoom_out()
                return
            if key == Qt.Key.Key_0:
                self._zoom_factor = 1.0
                self._render_page()
                return
        super().keyPressEvent(event)

    def _prev_page(self):
        if self._current_page > 0:
            self._show_page(self._current_page - 1)

    def _next_page(self):
        doc = self._active_doc
        if doc and self._current_page < len(doc) - 1:
            self._show_page(self._current_page + 1)

    def _goto_page(self):
        doc = self._active_doc
        if not doc:
            return
        try:
            n = int(self.page_entry.text())
            if 1 <= n <= len(doc):
                self._show_page(n - 1)
        except ValueError:
            pass

    # ==========================================================================
    # THUMBNAIL STRIP
    # ==========================================================================

    def _clear_thumb_strip(self):
        """Remove all thumbnail frames from the strip."""
        while self._thumb_inner_lay.count() > 1:
            item = self._thumb_inner_lay.takeAt(0)
            if item is None:
                break
            w = item.widget()
            if w is not None:
                w.deleteLater()

    def _stop_thumb_timer(self):
        if self._thumb_timer is not None:
            self._thumb_timer.stop()
            self._thumb_timer = None

    def _render_thumbs(self):
        self._stop_thumb_timer()
        self._clear_thumb_strip()
        self._thumb_pixmaps.clear()
        self._thumb_frames.clear()
        self._thumb_render_next = 0
        self._highlighted_thumb = -1

        doc = self._active_doc
        if doc is None:
            return

        ph_w = self.THUMB_W
        ph_h = int(self.THUMB_W * 1.4)
        total = len(doc)

        for i in range(total):
            self._thumb_pixmaps.append(None)

            # Container frame
            frame = QFrame()
            frame.setStyleSheet(
                f"QFrame {{ background: #F9FAFB; border: 1px solid {G200}; "
                f"border-radius: 4px; padding: 4px; }}"
            )
            frame.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
            frame.setFixedSize(ph_w + 8, 100)
            frame.mousePressEvent = lambda e, ii=i: self._show_page(ii)

            frame_lay = QVBoxLayout(frame)
            frame_lay.setContentsMargins(4, 4, 4, 4)
            frame_lay.setSpacing(3)

            # Placeholder image label
            img_lbl = QLabel()
            img_lbl.setFixedSize(ph_w, 64)
            img_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            img_lbl.setStyleSheet(
                f"background: {G200}; border: none;"
            )
            img_lbl.mousePressEvent = lambda e, ii=i: self._show_page(ii)
            frame_lay.addWidget(img_lbl)

            # Page number label
            num_lbl = QLabel(str(i + 1))
            num_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            num_lbl.setStyleSheet(
                f"color: {G500}; font: 9px 'Segoe UI'; "
                f"background: transparent; border: none;"
            )
            num_lbl.mousePressEvent = lambda e, ii=i: self._show_page(ii)
            frame_lay.addWidget(num_lbl)

            # Store img_lbl reference in frame for later update
            frame_any = cast(Any, frame)
            frame_any._img_lbl = img_lbl
            frame_any._num_lbl = num_lbl

            self._thumb_frames.append(frame)
            self._thumb_inner_lay.insertWidget(i, frame)

        # Let widgetResizable handle sizing; just ensure the inner widget updates
        self._thumb_inner.updateGeometry()

        # Kick off lazy rendering
        self._thumb_timer = QTimer(self)
        self._thumb_timer.setSingleShot(True)
        self._thumb_timer.timeout.connect(self._render_thumb_batch)
        self._thumb_timer.start(0)

    def _render_thumb_batch(self, batch: int = 8):
        doc = self._active_doc
        if doc is None:
            return
        start = self._thumb_render_next
        end   = min(start + batch, len(doc))

        for i in range(start, end):
            if self._thumb_pixmaps[i] is not None:
                continue
            try:
                page  = doc[i]
                scale = self.THUMB_W / page.rect.width
                pix   = page.get_pixmap(
                    matrix=fitz.Matrix(scale, scale), alpha=False)
                qpix  = _fitz_pix_to_qpixmap(pix)
                self._thumb_pixmaps[i] = qpix
            except Exception:
                continue

            if i < len(self._thumb_frames):
                frame = self._thumb_frames[i]
                frame_any = cast(Any, frame)
                frame_any._img_lbl.setPixmap(qpix)
                frame_any._img_lbl.setStyleSheet("background: transparent; border: none;")
                # Re-apply highlight if needed
                if i == self._highlighted_thumb:
                    frame.setStyleSheet(
                        f"QFrame {{ background: {BLUE_DIM}; "
                        f"border: 1px solid {BLUE}; border-radius: 4px; padding: 4px; }}"
                    )

        self._thumb_render_next = end
        if end < len(doc):
            self._thumb_timer = QTimer(self)
            self._thumb_timer.setSingleShot(True)
            self._thumb_timer.timeout.connect(self._render_thumb_batch)
            self._thumb_timer.start(0)
        else:
            self._thumb_timer = None

    def _hl_thumb(self, idx: int):
        old = self._highlighted_thumb
        if old >= 0 and old < len(self._thumb_frames):
            old_frame = self._thumb_frames[old]
            old_frame.setStyleSheet(
                f"QFrame {{ background: #F9FAFB; border: 1px solid {G200}; "
                f"border-radius: 4px; padding: 4px; }}"
            )
            from typing import cast as _cast
            old_any = _cast(Any, old_frame)
            if hasattr(old_any, '_num_lbl'):
                old_any._num_lbl.setStyleSheet(
                    f"color: {G500}; font: 9px 'Segoe UI'; "
                    f"background: transparent; border: none;"
                )
        self._highlighted_thumb = idx
        if idx >= 0 and idx < len(self._thumb_frames):
            new_frame = self._thumb_frames[idx]
            new_frame.setStyleSheet(
                f"QFrame {{ background: {BLUE_DIM}; border: 1px solid {BLUE}; "
                f"border-radius: 4px; padding: 4px; }}"
            )
            new_any = cast(Any, new_frame)
            if hasattr(new_any, '_num_lbl'):
                new_any._num_lbl.setStyleSheet(
                    f"color: {BLUE_DARK}; font: bold 9px 'Segoe UI'; "
                    f"background: transparent; border: none;"
                )
            # Scroll to make it visible
            self._thumb_sa.ensureWidgetVisible(new_frame)

    # ==========================================================================
    # RUBBER-BAND SELECTION HELPERS
    # ==========================================================================

    def _point_on_page(self, cx: float, cy: float) -> bool:
        return (self._page_ox <= cx <= self._page_ox + self._page_iw and
                self._page_oy <= cy <= self._page_oy + self._page_ih)

    def _canvas_to_pdf_rect(self, x0, y0, x1, y1):
        """Convert clamped canvas pixel coords to a normalised fitz.Rect
        in PDF page coordinate space."""
        p0 = fitz.Point(x0 - self._page_ox,
                         y0 - self._page_oy) * self._inv_mat
        p1 = fitz.Point(x1 - self._page_ox,
                         y1 - self._page_oy) * self._inv_mat
        r = fitz.Rect(p0, p1)
        r.normalize()
        return r

    # ==========================================================================
    # FLASH FEEDBACK
    # ==========================================================================

    def _flash_feedback(self, x0, y0, x1, y1):
        self._flash_rect = (int(x0), int(y0), int(x1 - x0), int(y1 - y0))
        self._canvas.update()
        QTimer.singleShot(240, self._clear_flash)

    def _clear_flash(self):
        self._flash_rect = None
        self._canvas.update()

    # ==========================================================================
    # CAPTURE PIPELINE
    # ==========================================================================

    def _make_snippet_thumbnail(self, snip: Snippet) -> QPixmap | None:
        """Render just the crop region as a small QPixmap."""
        try:
            src_doc  = fitz.open(snip.source_path)
            clip     = snip.crop_rect
            scale    = self.SNIP_W / max(clip.width, 1)
            mat      = fitz.Matrix(scale, scale)
            pix      = src_doc[snip.page_index].get_pixmap(
                matrix=mat, clip=clip, alpha=False)
            src_doc.close()
            return _fitz_pix_to_qpixmap(pix)
        except Exception:
            return None

    _A4_W = 595.0
    _A4_H = 842.0

    def _reset_out_cursor(self):
        self._out_y_cursor = 0.0
        self._out_has_page = False

    def _append_to_output(self, snip: Snippet):
        """Pack snippet onto A4 pages vertically, preserving original x-position.
        Starts a new A4 page when the snippet does not fit in the remaining space."""
        try:
            src_doc  = fitz.open(snip.source_path)
            clip     = snip.crop_rect

            snip_w = clip.width
            snip_h = clip.height

            # Scale down if wider than A4
            if snip_w > self._A4_W:
                scale  = self._A4_W / snip_w
                snip_w = self._A4_W
                snip_h = snip_h * scale
                dest_x = 0.0
            else:
                dest_x = clip.x0
                # Keep within A4 bounds
                if dest_x + snip_w > self._A4_W:
                    dest_x = self._A4_W - snip_w

            # New A4 page if none exists or snippet won't fit on current page
            if not self._out_has_page or (self._out_y_cursor + snip_h > self._A4_H):
                self._out_doc.new_page(width=self._A4_W, height=self._A4_H)
                self._out_y_cursor = 0.0
                self._out_has_page = True

            page      = self._out_doc[-1]
            dest_rect = fitz.Rect(
                dest_x,
                self._out_y_cursor,
                dest_x + snip_w,
                self._out_y_cursor + snip_h,
            )
            page.show_pdf_page(dest_rect, src_doc, snip.page_index, clip=clip)
            self._out_y_cursor += snip_h
            src_doc.close()
        except Exception as e:
            QMessageBox.critical(self, "Capture Error",
                                  f"Failed to capture region:\n{e}")

    def _do_capture(self, crop_rect):
        doc_entry = self._pdf_list[self._active_idx]
        label = (f"{doc_entry['name']}  p.{self._current_page + 1}  "
                 f"({crop_rect.width:.0f}\u00d7{crop_rect.height:.0f} pt)")
        snip = Snippet(
            source_path=doc_entry["path"],
            page_index=self._current_page,
            crop_rect=crop_rect,
            label=label,
        )
        snip.thumbnail = self._make_snippet_thumbnail(snip)
        self._append_to_output(snip)
        self._snippets.append(snip)
        self._rebuild_snippet_list()
        self._status_lbl.setText(
            f"{len(self._snippets)} snippet(s) captured")

    # ==========================================================================
    # SNIPPET LIST UI
    # ==========================================================================

    def _rebuild_snippet_list(self):
        # Remove all snippet cards (keep stretch at end)
        while self._snip_layout.count() > 1:
            item = self._snip_layout.takeAt(0)
            if item is None:
                break
            w = item.widget()
            if w is not None:
                w.deleteLater()

        self._snip_count_lbl.setText(f"({len(self._snippets)})")

        for i, snip in enumerate(self._snippets):
            card = self._make_snippet_card(i, snip)
            self._snip_layout.insertWidget(i, card)

    def _make_snippet_card(self, idx: int, snip: Snippet) -> QFrame:
        card = QFrame()
        card.setStyleSheet(
            f"QFrame {{ background: {WHITE}; border: 1px solid {G200}; "
            f"border-radius: 12px; padding: 0px; }}"
        )
        card_lay = QHBoxLayout(card)
        card_lay.setContentsMargins(9, 9, 9, 9)
        card_lay.setSpacing(12)

        # Thumbnail container
        thumb_container = QFrame()
        thumb_container.setFixedSize(64, 64)
        thumb_container.setStyleSheet(
            f"QFrame {{ background: {G100}; border: 1px solid {G200}; "
            f"border-radius: 4px; }}"
        )
        thumb_container_lay = QVBoxLayout(thumb_container)
        thumb_container_lay.setContentsMargins(2, 2, 2, 2)
        thumb_container_lay.setSpacing(0)

        if snip.thumbnail and not snip.thumbnail.isNull():
            thumb_lbl = QLabel()
            scaled = snip.thumbnail.scaled(
                60, 60,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation
            )
            thumb_lbl.setPixmap(scaled)
            thumb_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            thumb_lbl.setStyleSheet("border: none; background: transparent;")
            thumb_container_lay.addWidget(thumb_lbl, 0, Qt.AlignmentFlag.AlignCenter)
        else:
            ph = QLabel()
            ph.setPixmap(svg_pixmap("file-text", "#9CA3AF", 28))
            ph.setAlignment(Qt.AlignmentFlag.AlignCenter)
            ph.setStyleSheet(
                f"color: {G400}; font: 20px; background: transparent; border: none;"
            )
            thumb_container_lay.addWidget(ph, 0, Qt.AlignmentFlag.AlignCenter)

        card_lay.addWidget(thumb_container)

        # Right side (text + actions)
        right_col = QWidget()
        right_col.setStyleSheet("background: transparent;")
        right_lay = QVBoxLayout(right_col)
        right_lay.setContentsMargins(0, 0, 0, 0)
        right_lay.setSpacing(2)

        # Title line
        title_lbl = QLabel(f"Page {snip.page_index + 1} Selection")
        title_lbl.setStyleSheet(
            f"color: {G700}; font: bold 11px 'Segoe UI'; "
            f"background: transparent; border: none;"
        )
        right_lay.addWidget(title_lbl)

        # Size info
        size_lbl = QLabel(
            f"{snip.crop_rect.width:.0f} × {snip.crop_rect.height:.0f} pt"
        )
        size_lbl.setStyleSheet(
            f"color: {G400}; font: 10px 'Segoe UI'; "
            f"background: transparent; border: none;"
        )
        right_lay.addWidget(size_lbl)

        right_lay.addSpacing(4)

        # Action buttons row
        action_row = QWidget()
        action_row.setStyleSheet("background: transparent;")
        action_lay = QHBoxLayout(action_row)
        action_lay.setContentsMargins(0, 0, 0, 0)
        action_lay.setSpacing(2)

        def _small_btn(label: str, color: str) -> QPushButton:
            b = QPushButton(label)
            b.setStyleSheet(
                f"QPushButton {{ background: transparent; color: {color}; "
                f"font: 10px 'Segoe UI'; border: none; padding: 0 2px; }}"
                f"QPushButton:hover {{ color: {color}; text-decoration: underline; }}"
            )
            b.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
            return b

        up_btn = _small_btn("↑", BLUE_DARK)
        up_btn.clicked.connect(lambda checked=False, ii=idx: self._move_snippet(ii, -1))
        action_lay.addWidget(up_btn)

        dn_btn = _small_btn("↓", BLUE_DARK)
        dn_btn.clicked.connect(lambda checked=False, ii=idx: self._move_snippet(ii, +1))
        action_lay.addWidget(dn_btn)

        action_lay.addStretch()

        del_btn = _small_btn("Delete", RED)
        del_btn.clicked.connect(lambda checked=False, ii=idx: self._delete_snippet(ii))
        action_lay.addWidget(del_btn)

        right_lay.addWidget(action_row)
        card_lay.addWidget(right_col, 1)

        return card

    def _delete_snippet(self, idx: int):
        if 0 <= idx < len(self._snippets):
            self._snippets.pop(idx)
            self._rebuild_output_doc()
            self._rebuild_snippet_list()
            self._status_lbl.setText(
                f"{len(self._snippets)} snippet(s) captured")

    def _move_snippet(self, idx: int, direction: int):
        new_idx = idx + direction
        if 0 <= new_idx < len(self._snippets):
            self._snippets[idx], self._snippets[new_idx] = (
                self._snippets[new_idx], self._snippets[idx])
            self._rebuild_output_doc()
            self._rebuild_snippet_list()

    # ==========================================================================
    # OUTPUT DOCUMENT
    # ==========================================================================

    def _rebuild_output_doc(self):
        """Reconstruct _out_doc from _snippets in current order."""
        try:
            self._out_doc.close()
        except Exception:
            pass
        self._out_doc = fitz.open()
        self._reset_out_cursor()
        for snip in self._snippets:
            self._append_to_output(snip)

    def _save_excerpt(self):
        if self._out_doc.page_count == 0:
            QMessageBox.warning(self, "Empty", "No snippets captured yet.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Excerpt PDF", "excerpt.pdf",
            "PDF Files (*.pdf)")
        if not path:
            return
        try:
            self._out_doc.save(path)
            QMessageBox.information(
                self, "Saved",
                f"Excerpt PDF saved:\n{path}\n"
                f"({self._out_doc.page_count} pages)")
            self._status_lbl.setText("Saved.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not save:\n{e}")

    # ==========================================================================
    # CLEANUP
    # ==========================================================================

    def cleanup(self):
        self._stop_thumb_timer()
        for entry in self._pdf_list:
            try:
                entry["doc"].close()
            except Exception:
                pass
        try:
            self._out_doc.close()
        except Exception:
            pass
