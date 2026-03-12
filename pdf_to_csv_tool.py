"""PDF to CSV Converter Tool — PySide6 port.

Two-panel layout: left settings panel, right page preview + report.
Bottom: scrollable thumbnail strip (same pattern as split_tool.py).
"""

from __future__ import annotations

import csv
import datetime
import io
import os
import re
import subprocess
import sys
import unicodedata
from typing import Optional

from PySide6.QtWidgets import (
    QWidget, QFrame, QLabel, QPushButton, QLineEdit,
    QComboBox, QScrollArea, QHBoxLayout, QVBoxLayout, QStackedWidget,
    QFileDialog, QMessageBox, QInputDialog, QProgressBar, QSizePolicy,
    QApplication,
)
from PySide6.QtCore import Qt, QTimer, QEvent, QObject, QSize
from PySide6.QtGui import QPainter, QColor, QPixmap, QPen, QFont, QCursor, QIcon
from icons import svg_pixmap, svg_icon
from colors import (
    BG, WHITE, G100, G200, G300, G400, G500, G700, G900,
    BLUE, BLUE_HOVER, GREEN, RED, THUMB_BG,
)
from utils import _fitz_pix_to_qpixmap, _WheelToHScroll

# ---------------------------------------------------------------------------
# Optional dependency check
# ---------------------------------------------------------------------------
try:
    import fitz
    _HAS_FITZ = True
except ImportError:
    _HAS_FITZ = False

try:
    import pdfplumber
    _HAS_PLUMBER = True
except ImportError:
    _HAS_PLUMBER = False

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
ENCODING_MAP = {
    "UTF-8":          "utf-8",
    "UTF-8 with BOM": "utf-8-sig",
    "UTF-16":         "utf-16",
    "ASCII":          "ascii",
    "Windows-1252":   "cp1252",
    "ISO-8859-1":     "iso-8859-1",
}

DELIMITER_MAP = {
    "Comma (,)":     ",",
    "Semicolon (;)": ";",
    "Tab":           "\t",
    "Pipe (|)":      "|",
}

LINE_ENDING_MAP = {
    "System default": os.linesep,
    "Unix (LF)":      "\n",
    "Windows (CRLF)": "\r\n",
}

def _render_page_qpixmap(doc, idx: int, max_w: int):
    page = doc[idx]
    scale = max_w / page.rect.width
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    return _fitz_pix_to_qpixmap(pix), scale


def _render_thumb_qpixmap(doc, idx: int, thumb_w: int) -> QPixmap:
    page = doc[idx]
    scale = thumb_w / page.rect.width
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    return _fitz_pix_to_qpixmap(pix)


# ---------------------------------------------------------------------------
# Custom page-preview canvas widget
# ---------------------------------------------------------------------------

class _PreviewCanvas(QWidget):
    def __init__(self, tool: "PDFtoCSVTool", parent=None):
        super().__init__(parent)
        self._t = tool
        self.setSizePolicy(QSizePolicy.Policy.Expanding,
                           QSizePolicy.Policy.Expanding)
        self.setMinimumSize(300, 300)

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        t = self._t
        p.fillRect(self.rect(), QColor(THUMB_BG))
        if t._page_pixmap is None:
            p.setPen(QColor(G400))
            p.setFont(QFont("Segoe UI", 13))
            p.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter,
                       "Open a PDF to preview it here")
            return
        p.drawPixmap(t._page_ox, t._page_oy, t._page_pixmap)
        # Draw table outlines
        pen = QPen(QColor(BLUE), 2, Qt.PenStyle.DashLine)
        p.setPen(pen)
        p.setBrush(Qt.BrushStyle.NoBrush)
        for (cx0, cy0, cx1, cy1) in t._canvas_table_rects:
            p.drawRect(int(cx0), int(cy0),
                       int(cx1 - cx0), int(cy1 - cy0))

    def resizeEvent(self, e):
        super().resizeEvent(e)
        if self._t._doc:
            self._t._render_page_canvas()


# ---------------------------------------------------------------------------
# Main tool class
# ---------------------------------------------------------------------------

class PDFtoCSVTool(QWidget):
    LEFT_W  = 460
    THUMB_W = 80

    def __init__(self, parent=None):
        super().__init__(parent)

        # ── State ─────────────────────────────────────────────────────────
        self.pdf_path: str = ""
        self.output_dir: str = ""
        self._password: str = ""
        self._doc   = None          # fitz.Document
        self._pldoc = None          # pdfplumber document
        self._total_pages = 0
        self._current_page = 0
        self._thumb_pixmaps: list = []   # list of (QPixmap|None, frame_widget, label_widget)
        self._thumb_render_next = 0
        self._thumb_timer: Optional[QTimer] = None
        self._highlighted_thumb_frame = None
        self._page_pixmap: Optional[QPixmap] = None
        self._page_scale = 1.0
        self._page_ox = 0
        self._page_oy = 0
        self._table_bboxes: list = []
        self._canvas_table_rects: list = []
        self._report_widget: Optional[QWidget] = None

        # ── Settings defaults & widget registry ───────────────────────────
        self._sv_defaults = {
            "detection":    "Auto",
            "row_tol":      "3",
            "col_tol":      "3",
            "header":       "Auto-detect",
            "linebreak":    "Replace with space",
            "custom_lb":    " | ",
            "merged":       "First column only",
            "vert_merge":   "First row only",
            "empty_marker": "",
            "delimiter":    "Comma (,)",
            "encoding":     "UTF-8 with BOM",
            "multi":        "Separate file per table",
            "range":        "all",
            "overwrite":    "Rename with suffix",
            "image_only":   "Skip with warning",
            "min_rows":     "1",
            "min_cols":     "1",
            "source_meta":  "None",
            "strip_ws":     "Enabled",
            "line_ending":  "System default",
            "unicode_norm": "NFC (recommended)",
            "type_detect":  "Disabled",
        }
        self._widgets: dict = {}

        # ── Main layout ───────────────────────────────────────────────────
        root_v = QVBoxLayout(self)
        root_v.setContentsMargins(0, 0, 0, 0)
        root_v.setSpacing(0)

        # Top area: left panel + right preview
        top = QWidget()
        top.setStyleSheet("background: transparent;")
        top_h = QHBoxLayout(top)
        top_h.setContentsMargins(0, 0, 0, 0)
        top_h.setSpacing(0)
        root_v.addWidget(top, 1)

        # Bottom thumbnail strip
        self._strip_container = QWidget()
        self._strip_container.setFixedHeight(155)
        self._strip_container.setStyleSheet(f"background: {G100};")
        root_v.addWidget(self._strip_container)

        if not _HAS_FITZ or not _HAS_PLUMBER:
            self._build_missing_deps(top)
        else:
            self._build_left(top_h)
            self._build_right(top_h)
            self._build_thumb_strip()

    # ======================================================================
    # DEPENDENCY ERROR SCREEN
    # ======================================================================

    def _build_missing_deps(self, top: QWidget):
        w = QWidget()
        w.setStyleSheet(f"background: {WHITE};")
        lay = QVBoxLayout(w)
        msg = "Missing dependencies:\n"
        if not _HAS_FITZ:
            msg += "  • pymupdf  (pip install pymupdf)\n"
        if not _HAS_PLUMBER:
            msg += "  • pdfplumber  (pip install pdfplumber)\n"
        lbl = QLabel(msg)
        lbl.setStyleSheet(f"color: {RED}; font: 14px 'Segoe UI';")
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(lbl)
        # top already has QHBoxLayout set on it; use it directly
        top_lay = top.layout()
        if top_lay is not None:
            top_lay.addWidget(w)

    # ======================================================================
    # BUILD LEFT PANEL
    # ======================================================================

    def _build_left(self, parent_h: QHBoxLayout):
        # Outer fixed-width container
        left_outer = QWidget()
        left_outer.setFixedWidth(self.LEFT_W)
        left_outer.setStyleSheet(f"background: {G100};")
        outer_v = QVBoxLayout(left_outer)
        outer_v.setContentsMargins(0, 0, 0, 0)
        outer_v.setSpacing(0)
        parent_h.addWidget(left_outer)

        # Scrollable inner area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(f"border: none; background: {G100};")
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        outer_v.addWidget(scroll)

        inner = QWidget()
        inner.setStyleSheet(f"background: {G100};")
        scroll.setWidget(inner)
        left = QVBoxLayout(inner)
        left.setContentsMargins(0, 0, 0, 0)
        left.setSpacing(0)

        # Title
        title = QLabel("PDF to CSV")
        title.setStyleSheet(f"color: {G900}; font: bold 20px 'Segoe UI';")
        title.setContentsMargins(18, 18, 18, 14)
        left.addWidget(title)

        # ── Input File ────────────────────────────────────────────────────
        self._section(left, "Input File")

        file_row = QWidget()
        file_row.setStyleSheet("background: transparent;")
        fr_h = QHBoxLayout(file_row)
        fr_h.setContentsMargins(18, 0, 18, 8)
        fr_h.setSpacing(6)
        self._file_entry = QLineEdit()
        self._file_entry.setReadOnly(True)
        self._file_entry.setPlaceholderText("No file selected…")
        self._file_entry.setFixedHeight(34)
        self._file_entry.setStyleSheet(
            f"background: {WHITE}; color: {G700}; border: 1px solid {G200}; "
            f"border-radius: 4px; font: 12px 'Segoe UI'; padding: 0 6px;")
        fr_h.addWidget(self._file_entry, 1)
        browse_btn = QPushButton("Browse")
        browse_btn.setFixedSize(72, 34)
        browse_btn.setStyleSheet(
            f"QPushButton {{background: {BLUE}; color: white; border-radius: 6px; "
            f"font: 13px 'Segoe UI';}} "
            f"QPushButton:hover {{background: {BLUE_HOVER};}}")
        browse_btn.clicked.connect(self._browse_file)
        fr_h.addWidget(browse_btn)
        left.addWidget(file_row)

        # Page range row
        range_row = QWidget()
        range_row.setStyleSheet("background: transparent;")
        rr_h = QHBoxLayout(range_row)
        rr_h.setContentsMargins(18, 0, 18, 14)
        rr_h.setSpacing(8)
        rr_lbl = QLabel("Page range:")
        rr_lbl.setStyleSheet(f"color: {G500}; font: 12px 'Segoe UI';")
        rr_h.addWidget(rr_lbl)
        range_entry = QLineEdit(self._sv_defaults["range"])
        range_entry.setFixedHeight(30)
        range_entry.setFixedWidth(110)
        range_entry.setStyleSheet(
            f"background: {WHITE}; color: {G700}; border: 1px solid {G200}; "
            f"border-radius: 4px; font: 12px 'Segoe UI'; padding: 0 6px;")
        rr_h.addWidget(range_entry)
        self._widgets["range"] = range_entry
        hint = QLabel("e.g. all  1  3-7  1,3,5-7")
        hint.setStyleSheet(f"color: {G400}; font: 10px 'Segoe UI';")
        rr_h.addWidget(hint)
        rr_h.addStretch()
        left.addWidget(range_row)

        # ── Table Detection ───────────────────────────────────────────────
        self._section(left, "Table Detection")
        self._dropdown(left, "Detection method", "detection",
                       ["Auto", "Lattice", "Stream", "Hybrid"])
        self._labeled_entry(left, "Row tolerance (pt)", "row_tol")
        self._labeled_entry(left, "Column tolerance (pt)", "col_tol")

        # ── Table Filters ─────────────────────────────────────────────────
        self._section(left, "Table Filters")
        self._labeled_entry(left, "Min rows (skip smaller)", "min_rows")
        self._labeled_entry(left, "Min columns (skip smaller)", "min_cols")
        self._dropdown(left, "Image-only pages", "image_only",
                       ["Skip with warning", "Fail entirely"])

        # ── Extraction Settings ───────────────────────────────────────────
        self._section(left, "Extraction Settings")
        self._dropdown(left, "Header row", "header",
                       ["Auto-detect", "First row is header", "No headers"])
        self._dropdown(left, "Line breaks in cells", "linebreak",
                       ["Replace with space", "Replace with custom",
                        "Preserve (\\n in cell)", "Remove entirely"],
                       on_change=self._on_linebreak_change)

        # Custom line-break row (hidden by default)
        self._custom_lb_widget = QWidget()
        self._custom_lb_widget.setStyleSheet("background: transparent;")
        clb_h = QHBoxLayout(self._custom_lb_widget)
        clb_h.setContentsMargins(18, 0, 18, 6)
        clb_lbl = QLabel("Custom replacement:")
        clb_lbl.setFixedWidth(160)
        clb_lbl.setStyleSheet(f"color: {G500}; font: 12px 'Segoe UI';")
        clb_h.addWidget(clb_lbl)
        clb_entry = QLineEdit(self._sv_defaults["custom_lb"])
        clb_entry.setFixedHeight(28)
        clb_entry.setStyleSheet(
            f"background: {WHITE}; color: {G700}; border: 1px solid {G200}; "
            f"border-radius: 4px; font: 12px 'Segoe UI'; padding: 0 6px;")
        clb_h.addWidget(clb_entry, 1)
        self._widgets["custom_lb"] = clb_entry
        left.addWidget(self._custom_lb_widget)
        self._custom_lb_widget.hide()

        self._dropdown(left, "Horizontal merged cells", "merged",
                       ["First column only", "Duplicate across columns", "Leave empty"])
        self._dropdown(left, "Vertical merged cells", "vert_merge",
                       ["First row only", "Duplicate down rows", "Leave empty"])
        self._labeled_entry(left, "Empty cell marker", "empty_marker")
        empty_hint = QLabel("  (text placed in empty merged cells; blank = leave empty)")
        empty_hint.setStyleSheet(f"color: {G400}; font: 10px 'Segoe UI';")
        empty_hint.setContentsMargins(18, 0, 18, 6)
        left.addWidget(empty_hint)

        self._dropdown(left, "Strip cell whitespace", "strip_ws",
                       ["Enabled", "Disabled"])
        self._dropdown(left, "Unicode normalization", "unicode_norm",
                       ["NFC (recommended)", "NFKC (compatibility)", "None"])
        uni_hint = QLabel("  (NFC fixes invisible combining chars from PDF fonts)")
        uni_hint.setStyleSheet(f"color: {G400}; font: 10px 'Segoe UI';")
        uni_hint.setContentsMargins(18, 0, 18, 6)
        left.addWidget(uni_hint)

        self._dropdown(left, "Type detection", "type_detect",
                       ["Disabled", "Numbers only", "Dates only", "Numbers + Dates"])
        type_warn = QLabel("  \u26a0 May alter leading zeros, zip codes, phone numbers")
        type_warn.setStyleSheet("color: #D97706; font: 10px 'Segoe UI';")
        type_warn.setContentsMargins(18, 0, 18, 6)
        left.addWidget(type_warn)

        # ── Output Settings ───────────────────────────────────────────────
        self._section(left, "Output Settings")
        self._dropdown(left, "Delimiter", "delimiter",
                       list(DELIMITER_MAP.keys()))
        self._dropdown(left, "Encoding", "encoding",
                       list(ENCODING_MAP.keys()))
        self._dropdown(left, "Line endings", "line_ending",
                       list(LINE_ENDING_MAP.keys()))
        self._dropdown(left, "Multiple tables", "multi",
                       ["Separate file per table", "Single file (concatenate)"])
        self._dropdown(left, "Source metadata column", "source_meta",
                       ["None", "Page number", "Table number", "Page + Table"])
        meta_hint = QLabel("  (adds source column(s) in concatenated output)")
        meta_hint.setStyleSheet(f"color: {G400}; font: 10px 'Segoe UI';")
        meta_hint.setContentsMargins(18, 0, 18, 4)
        left.addWidget(meta_hint)

        self._dropdown(left, "If file exists", "overwrite",
                       ["Rename with suffix", "Overwrite", "Skip"])

        # Output folder
        folder_row = QWidget()
        folder_row.setStyleSheet("background: transparent;")
        fo_h = QHBoxLayout(folder_row)
        fo_h.setContentsMargins(18, 4, 18, 14)
        fo_h.setSpacing(6)
        self._folder_entry = QLineEdit()
        self._folder_entry.setReadOnly(True)
        self._folder_entry.setPlaceholderText("Output folder…")
        self._folder_entry.setFixedHeight(34)
        self._folder_entry.setStyleSheet(
            f"background: {WHITE}; color: {G700}; border: 1px solid {G200}; "
            f"border-radius: 4px; font: 12px 'Segoe UI'; padding: 0 6px;")
        fo_h.addWidget(self._folder_entry, 1)
        folder_btn = QPushButton("Browse")
        folder_btn.setFixedSize(72, 34)
        folder_btn.setStyleSheet(
            f"QPushButton {{background: transparent; color: {G700}; "
            f"border: 1px solid {G300}; border-radius: 6px; font: 13px 'Segoe UI';}} "
            f"QPushButton:hover {{background: {G200};}}")
        folder_btn.clicked.connect(self._browse_folder)
        fo_h.addWidget(folder_btn)
        left.addWidget(folder_row)

        # ── Action ────────────────────────────────────────────────────────
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setFixedHeight(1)
        sep.setStyleSheet(f"background: {G200};")
        sep.setContentsMargins(18, 6, 18, 14)
        left.addWidget(sep)

        extract_wrap = QWidget()
        extract_wrap.setStyleSheet("background: transparent;")
        ew_v = QVBoxLayout(extract_wrap)
        ew_v.setContentsMargins(18, 0, 18, 10)
        self._extract_btn = QPushButton("Extract to CSV")
        self._extract_btn.setFixedHeight(40)
        self._extract_btn.setStyleSheet(
            f"QPushButton {{background: {BLUE}; color: white; border-radius: 6px; "
            f"font: bold 14px 'Segoe UI';}} "
            f"QPushButton:hover {{background: {BLUE_HOVER};}} "
            f"QPushButton:disabled {{background: {G300}; color: {G500};}}")
        self._extract_btn.clicked.connect(self._run_extraction)
        ew_v.addWidget(self._extract_btn)
        left.addWidget(extract_wrap)

        progress_wrap = QWidget()
        progress_wrap.setStyleSheet("background: transparent;")
        pw_v = QVBoxLayout(progress_wrap)
        pw_v.setContentsMargins(18, 0, 18, 6)
        self._progress = QProgressBar()
        self._progress.setFixedHeight(8)
        self._progress.setRange(0, 100)
        self._progress.setValue(0)
        self._progress.setTextVisible(False)
        self._progress.setStyleSheet(
            f"QProgressBar {{background: {G200}; border-radius: 4px; border: none;}} "
            f"QProgressBar::chunk {{background: {GREEN}; border-radius: 4px;}}")
        pw_v.addWidget(self._progress)
        left.addWidget(progress_wrap)

        status_wrap = QWidget()
        status_wrap.setStyleSheet("background: transparent;")
        sw_v = QVBoxLayout(status_wrap)
        sw_v.setContentsMargins(18, 0, 18, 18)
        self._status_lbl = QLabel("")
        self._status_lbl.setWordWrap(True)
        self._status_lbl.setStyleSheet(f"color: {G500}; font: 11px 'Segoe UI';")
        sw_v.addWidget(self._status_lbl)
        left.addWidget(status_wrap)

        left.addStretch()

    # ======================================================================
    # BUILD RIGHT PANEL
    # ======================================================================

    def _build_right(self, parent_h: QHBoxLayout):
        right_container = QWidget()
        right_container.setStyleSheet(f"background: {WHITE};")
        right_v = QVBoxLayout(right_container)
        right_v.setContentsMargins(0, 0, 0, 0)
        right_v.setSpacing(0)
        parent_h.addWidget(right_container, 1)

        # Stacked widget: index 0 = canvas, index 1 = report
        self._right_stack = QStackedWidget()
        right_v.addWidget(self._right_stack, 1)

        # Canvas widget (index 0)
        canvas_container = QWidget()
        canvas_container.setStyleSheet(f"background: {THUMB_BG};")
        cc_v = QVBoxLayout(canvas_container)
        cc_v.setContentsMargins(0, 0, 0, 0)
        cc_v.setSpacing(0)
        self._canvas = _PreviewCanvas(self)
        cc_v.addWidget(self._canvas)
        self._right_stack.addWidget(canvas_container)

        # Navigation bar
        nav = QFrame()
        nav.setFixedHeight(44)
        nav.setStyleSheet(f"background: {G100}; border-top: 1px solid {G200};")
        nav_h = QHBoxLayout(nav)
        nav_h.setContentsMargins(10, 7, 10, 7)
        nav_h.setSpacing(4)

        prev_btn = QPushButton()
        prev_btn.setIcon(QIcon(svg_pixmap("chevron-left", G700, 14)))
        prev_btn.setIconSize(QSize(14, 14))
        prev_btn.setFixedSize(34, 30)
        prev_btn.setStyleSheet(
            f"QPushButton {{background: transparent; border-radius: 4px;}} "
            f"QPushButton:hover {{background: {G200};}}")
        prev_btn.clicked.connect(self._prev_page)
        nav_h.addWidget(prev_btn)

        next_btn = QPushButton()
        next_btn.setIcon(QIcon(svg_pixmap("chevron-right", G700, 14)))
        next_btn.setIconSize(QSize(14, 14))
        next_btn.setFixedSize(34, 30)
        next_btn.setStyleSheet(
            f"QPushButton {{background: transparent; border-radius: 4px;}} "
            f"QPushButton:hover {{background: {G200};}}")
        next_btn.clicked.connect(self._next_page)
        nav_h.addWidget(next_btn)

        self._page_lbl = QLabel("No file loaded")
        self._page_lbl.setStyleSheet(f"color: {G500}; font: 12px 'Segoe UI';")
        nav_h.addWidget(self._page_lbl)
        nav_h.addStretch()

        right_v.addWidget(nav)

    # ======================================================================
    # BUILD THUMBNAIL STRIP
    # ======================================================================

    def _build_thumb_strip(self):
        strip_v = QVBoxLayout(self._strip_container)
        strip_v.setContentsMargins(0, 0, 0, 0)
        strip_v.setSpacing(0)

        self._thumb_scroll = QScrollArea()
        self._thumb_scroll.setWidgetResizable(True)
        self._thumb_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self._thumb_scroll.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._thumb_scroll.setStyleSheet(
            f"border: none; background: {G100};")
        strip_v.addWidget(self._thumb_scroll)

        self._thumb_inner = QWidget()
        self._thumb_inner.setStyleSheet(f"background: {G100};")
        self._thumb_h_lay = QHBoxLayout(self._thumb_inner)
        self._thumb_h_lay.setContentsMargins(4, 8, 4, 8)
        self._thumb_h_lay.setSpacing(4)
        self._thumb_h_lay.addStretch()
        self._thumb_scroll.setWidget(self._thumb_inner)

        # Install wheel-to-horizontal-scroll event filter
        _WheelToHScroll(self._thumb_scroll)

    # ======================================================================
    # UI HELPER METHODS
    # ======================================================================

    def _combo_style(self) -> str:
        return (
            f"QComboBox {{background: {WHITE}; color: {G700}; "
            f"border: 1px solid {G200}; border-radius: 4px; "
            f"font: 12px 'Segoe UI'; padding: 0 6px;}} "
            f"QComboBox::drop-down {{border: none;}} "
            f"QComboBox QAbstractItemView {{background: {WHITE}; color: {G700}; "
            f"selection-background-color: {G100};}}"
        )

    def _section(self, parent_lay: QVBoxLayout, title: str):
        row = QWidget()
        row.setStyleSheet("background: transparent;")
        h = QHBoxLayout(row)
        h.setContentsMargins(18, 12, 18, 6)
        h.setSpacing(8)
        lbl = QLabel(title)
        lbl.setStyleSheet(f"color: {G700}; font: bold 12px 'Segoe UI';")
        h.addWidget(lbl)
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet(f"color: {G300};")
        h.addWidget(sep, 1)
        parent_lay.addWidget(row)

    def _dropdown(self, parent_lay: QVBoxLayout, label: str, key: str,
                  values: list, on_change=None):
        row = QWidget()
        row.setStyleSheet("background: transparent;")
        h = QHBoxLayout(row)
        h.setContentsMargins(18, 0, 18, 6)
        h.setSpacing(0)
        lbl = QLabel(label + ":")
        lbl.setFixedWidth(160)
        lbl.setStyleSheet(f"color: {G500}; font: 12px 'Segoe UI';")
        h.addWidget(lbl)
        cb = QComboBox()
        cb.addItems(values)
        cb.setCurrentText(self._sv_defaults.get(key, values[0]) or values[0])
        cb.setStyleSheet(self._combo_style())
        cb.setFixedHeight(28)
        if on_change:
            cb.currentTextChanged.connect(on_change)
        h.addWidget(cb, 1)
        self._widgets[key] = cb
        parent_lay.addWidget(row)

    def _labeled_entry(self, parent_lay: QVBoxLayout, label: str, key: str,
                       width: int = 80):
        row = QWidget()
        row.setStyleSheet("background: transparent;")
        h = QHBoxLayout(row)
        h.setContentsMargins(18, 0, 18, 6)
        h.setSpacing(0)
        lbl = QLabel(label + ":")
        lbl.setFixedWidth(160)
        lbl.setStyleSheet(f"color: {G500}; font: 12px 'Segoe UI';")
        h.addWidget(lbl)
        entry = QLineEdit(self._sv_defaults.get(key, ""))
        entry.setFixedHeight(28)
        entry.setStyleSheet(
            f"background: {WHITE}; color: {G700}; border: 1px solid {G200}; "
            f"border-radius: 4px; font: 12px 'Segoe UI'; padding: 0 6px;")
        h.addWidget(entry, 1)
        self._widgets[key] = entry
        parent_lay.addWidget(row)

    def _get(self, key: str) -> str:
        w = self._widgets.get(key)
        if w is None:
            return self._sv_defaults.get(key, "")
        if isinstance(w, QComboBox):
            return w.currentText()
        if isinstance(w, QLineEdit):
            return w.text()
        return ""

    def _on_linebreak_change(self, value: str):
        if value == "Replace with custom":
            self._custom_lb_widget.show()
        else:
            self._custom_lb_widget.hide()

    # ======================================================================
    # FILE / FOLDER BROWSE
    # ======================================================================

    def _browse_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open PDF", "", "PDF files (*.pdf)")
        if path:
            self._open_pdf(path)

    def _browse_folder(self):
        path = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if path:
            self.output_dir = path
            self._folder_entry.setText(path)

    # ======================================================================
    # PDF OPEN / VALIDATION
    # ======================================================================

    def _open_pdf(self, path: str):
        if self._doc:
            try:
                self._doc.close()
            except Exception:
                pass
        if self._pldoc:
            try:
                self._pldoc.close()
            except Exception:
                pass
        self._doc = None
        self._pldoc = None
        self._password = ""

        try:
            doc = fitz.open(path)
        except Exception as e:
            QMessageBox.critical(self, "Cannot Open PDF",
                                 f"Failed to open file:\n{e}")
            return

        if doc.needs_pass:
            pw, ok = QInputDialog.getText(
                self, "Password Required",
                "This PDF is password-protected.\nEnter password:",
                QLineEdit.EchoMode.Password)
            if not ok:
                doc.close()
                return
            if not doc.authenticate(pw):
                QMessageBox.critical(self, "Wrong Password",
                                     "Incorrect password. Cannot open PDF.")
                doc.close()
                return
            self._password = pw

        if doc.page_count == 0:
            QMessageBox.critical(self, "Empty PDF",
                                 "This PDF contains no pages.")
            doc.close()
            return

        if not self._detect_text_layer(doc):
            QMessageBox.warning(
                self, "No Text Layer Detected",
                "This PDF appears to contain scanned images without an "
                "extractable text layer.\n\n"
                "OCR support will be added in a future release.\n"
                "Extraction may produce empty tables.",
            )

        try:
            if self._password:
                self._pldoc = pdfplumber.open(path, password=self._password)
            else:
                self._pldoc = pdfplumber.open(path)
        except Exception as e:
            QMessageBox.critical(self, "pdfplumber Error",
                                 f"Could not open PDF with pdfplumber:\n{e}")
            doc.close()
            return

        self._doc = doc
        self.pdf_path = path
        self._total_pages = doc.page_count
        self._current_page = 0

        if not self.output_dir:
            self.output_dir = os.path.dirname(path)
            self._folder_entry.setText(self.output_dir)

        self._file_entry.setText(os.path.basename(path))

        self._build_thumbnails()
        self._render_page_canvas()
        n = doc.page_count
        self._status_lbl.setText(
            f"{n} page{'s' if n != 1 else ''} loaded.")

    def _detect_text_layer(self, doc) -> bool:
        sample_pages = min(3, doc.page_count)
        for i in range(sample_pages):
            raw = doc[i].get_text()
            text = raw.strip() if isinstance(raw, str) else ""
            if len(text) > 10:
                return True
        return False

    def _page_is_image_only(self, page_idx: int) -> bool:
        """Return True if the fitz page has no extractable text."""
        if not self._doc:
            return False
        raw = self._doc[page_idx].get_text()
        text = raw.strip() if isinstance(raw, str) else ""
        return len(text) < 5

    # ======================================================================
    # THUMBNAILS
    # ======================================================================

    def _build_thumbnails(self):
        # Clear existing thumbnails
        while self._thumb_h_lay.count() > 1:
            item = self._thumb_h_lay.takeAt(0)
            if item is not None:
                w = item.widget()
                if w is not None:
                    w.deleteLater()

        self._thumb_pixmaps = []
        self._highlighted_thumb_frame = None
        self._thumb_render_next = 0

        for i in range(self._total_pages):
            frame = QWidget()
            frame.setStyleSheet(f"background: {G100};")
            frame_v = QVBoxLayout(frame)
            frame_v.setContentsMargins(0, 0, 0, 0)
            frame_v.setSpacing(2)
            frame.setFixedWidth(self.THUMB_W + 8)

            thumb_lbl = QLabel()
            thumb_lbl.setFixedSize(self.THUMB_W, 110)
            thumb_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            thumb_lbl.setStyleSheet(
                f"background: {THUMB_BG}; border: 1px solid {G300};")
            frame_v.addWidget(thumb_lbl, 0, Qt.AlignmentFlag.AlignHCenter)

            num_lbl = QLabel(str(i + 1))
            num_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            num_lbl.setStyleSheet(f"color: {G500}; font: 9px 'Segoe UI';")
            frame_v.addWidget(num_lbl)

            # Click handler
            idx_capture = i
            thumb_lbl.mousePressEvent = lambda e, idx=idx_capture: self._go_to_page(idx)
            num_lbl.mousePressEvent  = lambda e, idx=idx_capture: self._go_to_page(idx)
            frame.mousePressEvent    = lambda e, idx=idx_capture: self._go_to_page(idx)

            self._thumb_h_lay.insertWidget(self._thumb_h_lay.count() - 1, frame)
            self._thumb_pixmaps.append((None, frame, thumb_lbl))

        self._render_thumb_batch()

    def _render_thumb_batch(self, batch: int = 8):
        if not self._doc:
            return
        start = self._thumb_render_next
        end   = min(start + batch, self._total_pages)
        for i in range(start, end):
            pm_old, frame, lbl = self._thumb_pixmaps[i]
            if pm_old is not None:
                continue
            try:
                pm = _render_thumb_qpixmap(self._doc, i, self.THUMB_W)
            except Exception:
                continue
            self._thumb_pixmaps[i] = (pm, frame, lbl)
            lbl.setPixmap(pm.scaled(
                self.THUMB_W, 110,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation))
        self._thumb_render_next = end
        if end < self._total_pages:
            QTimer.singleShot(0, self._render_thumb_batch)
        else:
            self._highlight_thumb(self._current_page)

    def _highlight_thumb(self, idx: int):
        if self._highlighted_thumb_frame is not None:
            try:
                pm, frame, lbl = None, self._highlighted_thumb_frame, None
                for entry in self._thumb_pixmaps:
                    if entry[1] is self._highlighted_thumb_frame:
                        lbl = entry[2]
                        break
                if lbl:
                    lbl.setStyleSheet(
                        f"background: {THUMB_BG}; border: 1px solid {G300};")
            except Exception:
                pass
        if 0 <= idx < len(self._thumb_pixmaps):
            _, frame, lbl = self._thumb_pixmaps[idx]
            lbl.setStyleSheet(
                f"background: {THUMB_BG}; border: 2px solid {BLUE};")
            self._highlighted_thumb_frame = frame
        else:
            self._highlighted_thumb_frame = None

    # ======================================================================
    # PAGE RENDERING
    # ======================================================================

    def _render_page_canvas(self):
        if not self._doc:
            return
        cw = max(self._canvas.width(), 100)
        try:
            pm, scale = _render_page_qpixmap(self._doc, self._current_page,
                                              cw - 20)
        except Exception:
            return
        self._page_pixmap = pm
        self._page_scale = scale
        iw, ih = pm.width(), pm.height()
        self._page_ox = (cw - iw) // 2
        self._page_oy = 10
        self._page_lbl.setText(
            f"Page {self._current_page + 1} / {self._total_pages}")
        self._draw_table_outlines()
        self._highlight_thumb(self._current_page)
        self._canvas.update()

    def _draw_table_outlines(self):
        self._canvas_table_rects = []
        if not self._pldoc:
            return
        try:
            pl_page = self._pldoc.pages[self._current_page]
            tables = pl_page.find_tables(self._build_table_settings())
        except Exception:
            return
        self._table_bboxes = []
        for table in tables:
            x0, top, x1, bottom = table.bbox
            self._table_bboxes.append((x0, top, x1, bottom))
            cx0 = self._page_ox + x0 * self._page_scale
            cy0 = self._page_oy + top * self._page_scale
            cx1 = self._page_ox + x1 * self._page_scale
            cy1 = self._page_oy + bottom * self._page_scale
            self._canvas_table_rects.append((cx0, cy0, cx1, cy1))

    # ======================================================================
    # NAVIGATION
    # ======================================================================

    def _go_to_page(self, idx: int):
        if not self._doc:
            return
        idx = max(0, min(idx, self._total_pages - 1))
        self._current_page = idx
        self._render_page_canvas()

    def _prev_page(self):
        self._go_to_page(self._current_page - 1)

    def _next_page(self):
        self._go_to_page(self._current_page + 1)

    # ======================================================================
    # TABLE SETTINGS BUILDER
    # ======================================================================

    def _build_table_settings(self, method: Optional[str] = None) -> dict:
        if method is None:
            method = self._get("detection")
        try:
            row_tol = max(1, int(self._get("row_tol")))
        except ValueError:
            row_tol = 3
        try:
            col_tol = max(1, int(self._get("col_tol")))
        except ValueError:
            col_tol = 3

        if method == "Lattice":
            v_strat = "lines"
            h_strat = "lines"
        elif method == "Stream":
            v_strat = "text"
            h_strat = "text"
        elif method == "Hybrid":
            v_strat = "lines_strict"
            h_strat = "lines_strict"
        else:  # Auto — start with lines, fall back handled in extraction
            v_strat = "lines"
            h_strat = "lines"

        return {
            "vertical_strategy":         v_strat,
            "horizontal_strategy":       h_strat,
            "intersection_y_tolerance":  row_tol,
            "intersection_x_tolerance":  col_tol,
            "snap_y_tolerance":          row_tol,
            "snap_x_tolerance":          col_tol,
            "edge_min_length":           3,
            "min_words_vertical":        1,
            "min_words_horizontal":      1,
            "keep_blank_chars":          False,
            "text_tolerance":            3,
            "text_x_tolerance":          3,
            "text_y_tolerance":          3,
            "explicit_vertical_lines":   [],
            "explicit_horizontal_lines": [],
        }

    # ======================================================================
    # PAGE RANGE PARSER
    # ======================================================================

    @staticmethod
    def _parse_page_range(spec: str, total: int) -> list:
        spec = spec.strip().lower()
        if spec in ("", "all"):
            return list(range(total))
        pages: set = set()
        for part in spec.split(","):
            part = part.strip()
            if "-" in part:
                lo, _, hi = part.partition("-")
                lo_i = int(lo.strip()) - 1
                hi_i = int(hi.strip()) - 1
                if lo_i < 0 or hi_i >= total or lo_i > hi_i:
                    raise ValueError(f"Page range '{part}' is out of bounds "
                                     f"(document has {total} pages).")
                pages.update(range(lo_i, hi_i + 1))
            else:
                idx = int(part) - 1
                if idx < 0 or idx >= total:
                    raise ValueError(f"Page number {part} is out of bounds "
                                     f"(document has {total} pages).")
                pages.add(idx)
        return sorted(pages)

    # ======================================================================
    # OUTPUT PATH HELPERS (overwrite protection)
    # ======================================================================

    def _resolve_output_path(self, fpath: str) -> Optional[str]:
        """
        Apply the overwrite-protection policy to *fpath*.
        Returns the final path to write to, or None if the file should be
        skipped (policy = "Skip").
        """
        if not os.path.exists(fpath):
            return fpath

        policy = self._get("overwrite")

        if policy == "Overwrite":
            return fpath

        if policy == "Skip":
            return None  # caller interprets None as "skipped"

        # "Rename with suffix" — append _1, _2, … until free
        base, ext = os.path.splitext(fpath)
        counter = 1
        while True:
            candidate = f"{base}_{counter}{ext}"
            if not os.path.exists(candidate):
                return candidate
            counter += 1

    # ======================================================================
    # TABLE PROCESSING
    # ======================================================================

    def _process_table(self, raw: list) -> list:
        """Clean a raw pdfplumber table (list of rows, each row a list of cells)."""
        lb_mode      = self._get("linebreak")
        custom_lb    = self._get("custom_lb")
        merge_h_mode = self._get("merged")
        merge_v_mode = self._get("vert_merge")
        empty_marker = self._get("empty_marker")
        strip_ws     = self._get("strip_ws") == "Enabled"
        uni_norm     = self._get("unicode_norm")
        type_detect  = self._get("type_detect")

        rows: list = []

        # First pass: basic per-cell cleaning
        for raw_row in raw:
            if raw_row is None:
                continue
            row: list = []
            for cell in raw_row:
                if cell is None:
                    # Horizontal merged cell placeholder from pdfplumber
                    if merge_h_mode == "Duplicate across columns" and row:
                        row.append(row[-1])
                    elif empty_marker:
                        row.append(empty_marker)
                    else:
                        row.append("")
                    continue

                text = str(cell)

                # Line break handling
                if lb_mode == "Replace with space":
                    text = text.replace("\n", " ").replace("\r", " ")
                elif lb_mode == "Replace with custom":
                    text = text.replace("\n", custom_lb).replace("\r", "")
                elif lb_mode == "Remove entirely":
                    text = text.replace("\n", "").replace("\r", "")
                # else "Preserve": leave as-is

                # Strip whitespace (optional)
                if strip_ws:
                    text = text.strip()

                # Smart quotes → straight quotes
                text = (text.replace("\u2018", "'").replace("\u2019", "'")
                            .replace("\u201c", '"').replace("\u201d", '"'))

                # Ligature expansion
                text = (text.replace("\ufb01", "fi").replace("\ufb02", "fl")
                            .replace("\ufb00", "ff").replace("\ufb03", "ffi")
                            .replace("\ufb04", "ffl"))

                # Remove control characters (except standard whitespace)
                text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)

                # Collapse multiple spaces
                text = re.sub(r"  +", " ", text)

                # Unicode normalization — fixes invisible combining characters
                # that come from certain PDF font encodings and silently break
                # spreadsheet formulas / string comparisons.
                if uni_norm == "NFC (recommended)":
                    text = unicodedata.normalize("NFC", text)
                elif uni_norm == "NFKC (compatibility)":
                    text = unicodedata.normalize("NFKC", text)
                # else "None": leave as-is

                row.append(text)
            rows.append(row)

        # Second pass: vertical merge handling
        # pdfplumber returns None for cells that are vertically spanned —
        # those were already converted to "" or empty_marker above.
        # "Duplicate down rows" propagates the last non-empty value in each column.
        if merge_v_mode == "Duplicate down rows" and rows:
            n_cols = max(len(r) for r in rows)
            last_vals = [""] * n_cols
            for row in rows:
                for ci in range(len(row)):
                    if row[ci] == "" or row[ci] == empty_marker:
                        if last_vals[ci]:
                            row[ci] = last_vals[ci]
                    else:
                        last_vals[ci] = row[ci]

        # Third pass: type detection (opt-in)
        if type_detect != "Disabled":
            do_numbers = "Numbers" in type_detect
            do_dates   = "Dates"   in type_detect
            rows = [
                [self._convert_cell_type(c, do_numbers, do_dates) for c in row]
                for row in rows
            ]

        return rows

    # ── Type conversion helpers ────────────────────────────────────────────

    # Date formats attempted in order (most specific first)
    _DATE_FORMATS = [
        "%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d",
        "%d-%m-%Y", "%m-%d-%Y",
        "%d.%m.%Y", "%m.%d.%Y",
        "%d %b %Y", "%d %B %Y",
        "%b %d, %Y", "%B %d, %Y",
        "%Y/%m/%d",
    ]

    @staticmethod
    def _try_parse_date(text: str) -> Optional[str]:
        """Try to parse *text* as a date; return ISO-8601 string or None."""
        t = text.strip()
        for fmt in PDFtoCSVTool._DATE_FORMATS:
            try:
                return datetime.datetime.strptime(t, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
        return None

    @staticmethod
    def _try_parse_number(text: str) -> Optional[str]:
        """
        Try to parse *text* as a number.
        Returns a canonical numeric string (e.g. "1234.56") or None.
        Guards against strings that look numeric but are identifiers:
        leading zeros (zip codes, IDs), pure integers <= 4 digits that could
        be year values in a date column, phone-number-like strings.
        """
        t = text.strip()
        if not t:
            return None

        # Reject strings with leading zeros (zip codes, account numbers, etc.)
        # Allow "-0.5", "0.5", but not "007" or "01234"
        if re.match(r"^0\d", t):
            return None

        # Strip common currency / percent symbols and thousands separators
        cleaned = re.sub(r"[£€$¥₹,%]", "", t)
        # Handle thousands separators: 1,234,567 or 1.234.567
        # Determine which separator is the decimal point heuristically:
        # if there's exactly one comma/dot and it's followed by <= 3 digits at end → decimal
        cleaned = cleaned.replace(" ", "")  # non-breaking/normal spaces
        # Remove thousands commas (e.g. 1,234,567 → 1234567; 1,234.56 → 1234.56)
        if re.match(r"^-?\d{1,3}(,\d{3})+(\.\d+)?$", cleaned):
            cleaned = cleaned.replace(",", "")
        # European style: 1.234,56 → 1234.56
        elif re.match(r"^-?\d{1,3}(\.\d{3})+(,\d+)?$", cleaned):
            cleaned = cleaned.replace(".", "").replace(",", ".")

        try:
            val = float(cleaned)
        except ValueError:
            return None

        # Return integer representation if no fractional part
        if val == int(val) and "." not in cleaned:
            return str(int(val))
        return str(val)

    @classmethod
    def _convert_cell_type(cls, text: str,
                            do_numbers: bool, do_dates: bool) -> str:
        """Convert a cell string to its canonical type representation if possible."""
        if not text.strip():
            return text
        if do_dates:
            parsed = cls._try_parse_date(text)
            if parsed is not None:
                return parsed
        if do_numbers:
            parsed = cls._try_parse_number(text)
            if parsed is not None:
                return parsed
        return text

    def _detect_header(self, rows: list) -> tuple:
        """Return (has_header, header_row, data_rows)."""
        mode = self._get("header")
        if not rows:
            return False, [], []

        if mode == "No headers":
            return False, [], rows

        if mode == "First row is header":
            return True, rows[0], rows[1:]

        # Auto-detect
        first = rows[0]
        if not first:
            return False, [], rows

        score = 0
        if all(c != "" for c in first):
            score += 1
        if not any(re.match(r"^[\d.,%-]+$", c) for c in first if c):
            score += 1
        if len(set(c for c in first if c)) == len([c for c in first if c]):
            score += 1  # unique values

        if score >= 2:
            return True, first, rows[1:]
        return False, [], rows

    # ======================================================================
    # MINIMUM SIZE FILTER
    # ======================================================================

    def _passes_size_filter(self, rows: list) -> bool:
        """Return True if the table meets the minimum row/column thresholds."""
        try:
            min_rows = max(1, int(self._get("min_rows")))
        except ValueError:
            min_rows = 1
        try:
            min_cols = max(1, int(self._get("min_cols")))
        except ValueError:
            min_cols = 1

        if len(rows) < min_rows:
            return False
        if rows and max(len(r) for r in rows) < min_cols:
            return False
        return True

    # ======================================================================
    # CSV WRITER
    # ======================================================================

    def _write_csv(self, rows: list, path: str) -> None:
        enc_name    = self._get("encoding")
        encoding    = ENCODING_MAP.get(enc_name, "utf-8-sig")
        delim       = DELIMITER_MAP.get(self._get("delimiter"), ",")
        line_ending = LINE_ENDING_MAP.get(self._get("line_ending"), os.linesep)

        # Use io.open so we can control the line terminator precisely.
        # csv.writer's newline="" suppresses its own line endings; we add ours.
        with open(path, "w", newline="", encoding=encoding, errors="replace") as f:
            writer = csv.writer(f, delimiter=delim, quoting=csv.QUOTE_MINIMAL,
                                lineterminator=line_ending)
            writer.writerows(rows)

    # ======================================================================
    # SOURCE METADATA INJECTION
    # ======================================================================

    def _add_source_metadata(self, rows: list,
                              page_num: int, table_num: int,
                              is_header: bool) -> list:
        """
        Prepend source column(s) to every row.
        page_num and table_num are 1-based for display.
        is_header=True means the first row is a header — give it a label col.
        """
        meta_mode = self._get("source_meta")
        if meta_mode == "None":
            return rows

        result = []
        for i, row in enumerate(rows):
            is_hdr_row = (is_header and i == 0)
            if meta_mode == "Page number":
                prefix = ["Source page"] if is_hdr_row else [str(page_num)]
            elif meta_mode == "Table number":
                prefix = ["Source table"] if is_hdr_row else [str(table_num)]
            else:  # "Page + Table"
                prefix = (["Source page", "Source table"] if is_hdr_row
                           else [str(page_num), str(table_num)])
            result.append(prefix + list(row))
        return result

    # ======================================================================
    # COLUMN CONSISTENCY CHECK
    # ======================================================================

    @staticmethod
    def _check_column_consistency(rows: list) -> Optional[str]:
        """
        Return a warning string if the table has varying column counts,
        or None if all rows are consistent.
        """
        if not rows:
            return None
        col_counts = [len(r) for r in rows]
        unique_counts = set(col_counts)
        if len(unique_counts) <= 1:
            return None
        min_c = min(unique_counts)
        max_c = max(unique_counts)
        return (f"Inconsistent column count: rows range from {min_c} to {max_c} columns. "
                f"Some cells may be misaligned.")

    # ======================================================================
    # EXTRACTION PIPELINE
    # ======================================================================

    def _run_extraction(self):
        if not self._doc or not self._pldoc:
            QMessageBox.warning(self, "No File",
                                "Please open a PDF file first.")
            return
        if not self.output_dir:
            QMessageBox.warning(self, "No Output Folder",
                                "Please select an output folder.")
            return

        # Parse page range
        try:
            pages = self._parse_page_range(self._get("range"),
                                           self._total_pages)
        except ValueError as e:
            QMessageBox.critical(self, "Invalid Page Range", str(e))
            return

        if not pages:
            QMessageBox.warning(self, "Empty Selection",
                                "No pages to process.")
            return

        base_name  = os.path.splitext(os.path.basename(self.pdf_path))[0]
        multi_mode = self._get("multi")
        method     = self._get("detection")
        settings   = self._build_table_settings(method)
        meta_mode  = self._get("source_meta")
        image_only_policy = self._get("image_only")

        # Disable button during extraction
        self._extract_btn.setEnabled(False)
        self._extract_btn.setText("Extracting…")
        self._progress.setValue(0)

        report_lines: list = []
        report_lines.append("=== Extraction Complete ===\n")
        report_lines.append(f"Input:  {os.path.basename(self.pdf_path)}")
        report_lines.append(f"Output: {self.output_dir}")
        report_lines.append(f"Pages processed: {len(pages)}"
                             f"  (pages {pages[0]+1}–{pages[-1]+1})\n")

        all_table_rows: list = []    # for single-file mode
        total_tables   = 0
        total_rows     = 0
        skipped_files  = 0
        warnings: list = []
        output_files: list = []

        # Collect all tables first to compute progress
        page_table_data: list = []
        # Each entry: (page_idx, table_num_on_page, final_rows, has_header)

        for pg_idx in pages:
            self._status_lbl.setText(
                f"Detecting tables on page {pg_idx + 1}…")
            QApplication.processEvents()

            # Image-only page check
            if self._page_is_image_only(pg_idx):
                if image_only_policy == "Fail entirely":
                    QMessageBox.critical(
                        self, "Image-Only Page",
                        f"Page {pg_idx+1} contains only scanned images and "
                        "cannot be extracted.\n\nChange 'Image-only pages' to "
                        "'Skip with warning' to continue past such pages."
                    )
                    self._extract_btn.setEnabled(True)
                    self._extract_btn.setText("Extract to CSV")
                    return
                else:
                    warnings.append(
                        f"Page {pg_idx+1}: image-only page — no text layer, skipped.")
                    continue

            try:
                pl_page = self._pldoc.pages[pg_idx]
                raw_tables = pl_page.extract_tables(settings)
            except Exception as e:
                warnings.append(f"Page {pg_idx+1}: extraction error — {e}")
                continue

            # Auto fallback: if lines strategy found nothing, try text
            if method == "Auto" and not raw_tables:
                try:
                    fallback = self._build_table_settings("Stream")
                    raw_tables = pl_page.extract_tables(fallback)
                    if raw_tables:
                        warnings.append(
                            f"Page {pg_idx+1}: lattice detection found no tables, "
                            "fell back to stream mode.")
                except Exception:
                    pass

            if not raw_tables:
                warnings.append(f"Page {pg_idx+1}: no tables detected (skipped).")
                continue

            for tbl_idx, raw in enumerate(raw_tables):
                rows = self._process_table(raw)
                has_hdr, hdr, data_rows = self._detect_header(rows)

                # Minimum size filter
                if not self._passes_size_filter(rows):
                    warnings.append(
                        f"Page {pg_idx+1} table {tbl_idx+1}: "
                        f"too small ({len(rows)} rows × "
                        f"{max(len(r) for r in rows) if rows else 0} cols), skipped.")
                    continue

                if has_hdr:
                    final_rows = [hdr] + data_rows
                else:
                    final_rows = rows

                # Add source metadata columns (for concatenated mode)
                if multi_mode == "Single file (concatenate)" and meta_mode != "None":
                    final_rows = self._add_source_metadata(
                        final_rows, pg_idx + 1, tbl_idx + 1, has_hdr)

                page_table_data.append((pg_idx, tbl_idx + 1, final_rows, has_hdr))

        n_total = len(page_table_data)

        for i, (pg_idx, tbl_num, final_rows, has_hdr) in enumerate(page_table_data):
            total_tables += 1
            n_rows = len(final_rows)
            total_rows += n_rows
            n_cols = max(len(r) for r in final_rows) if final_rows else 0

            self._progress.setValue(int((i + 1) / max(n_total, 1) * 100))
            self._status_lbl.setText(
                f"Writing table {i+1}/{n_total} "
                f"(page {pg_idx+1}, table {tbl_num})…")
            QApplication.processEvents()

            # Column consistency check
            col_warn = self._check_column_consistency(final_rows)
            if col_warn:
                warnings.append(f"Page {pg_idx+1} table {tbl_num}: {col_warn}")

            if multi_mode == "Separate file per table":
                fname = f"{base_name}_page{pg_idx+1}_table{tbl_num}.csv"
                fpath_raw = os.path.join(self.output_dir, fname)
                fpath = self._resolve_output_path(fpath_raw)

                if fpath is None:
                    # Skipped due to overwrite policy
                    skipped_files += 1
                    warnings.append(
                        f"Page {pg_idx+1} table {tbl_num}: "
                        f"'{fname}' already exists — skipped.")
                    continue

                fname_actual = os.path.basename(fpath)
                try:
                    self._write_csv(final_rows, fpath)
                    output_files.append(fname_actual)
                except Exception as e:
                    warnings.append(
                        f"Page {pg_idx+1} table {tbl_num}: write error — {e}")
                    continue

                report_lines.append(f"Table {total_tables} — page {pg_idx+1}")
                report_lines.append(
                    f"  Dimensions: {n_rows} rows × {n_cols} columns")
                if fname_actual != fname:
                    report_lines.append(
                        f"  Output: {fname_actual}  (renamed — original existed)")
                else:
                    report_lines.append(f"  Output: {fname_actual}\n")

            else:  # single file — collect rows
                if all_table_rows and final_rows:
                    all_table_rows.append([])  # blank separator row
                all_table_rows.extend(final_rows)

                report_lines.append(f"Table {total_tables} — page {pg_idx+1}")
                report_lines.append(
                    f"  Dimensions: {n_rows} rows × {n_cols} columns\n")

        # Write combined file if single-file mode
        if multi_mode == "Single file (concatenate)" and all_table_rows:
            fname = f"{base_name}_all_tables.csv"
            fpath_raw = os.path.join(self.output_dir, fname)
            fpath = self._resolve_output_path(fpath_raw)

            if fpath is None:
                skipped_files += 1
                warnings.append(
                    f"'{fname}' already exists — skipped (overwrite policy).")
            else:
                fname_actual = os.path.basename(fpath)
                try:
                    self._write_csv(all_table_rows, fpath)
                    output_files.append(fname_actual)
                    if fname_actual != fname:
                        report_lines.append(
                            f"\nOutput: {fname_actual}  (renamed — original existed)")
                    else:
                        report_lines.append(f"\nOutput: {fname_actual}")
                except Exception as e:
                    warnings.append(f"Write error for combined file: {e}")

        # Summary
        report_lines.append(f"\n── Summary ──────────────────────")
        report_lines.append(f"Tables found:    {total_tables}")
        report_lines.append(f"Total rows:      {total_rows}")
        report_lines.append(f"Output files:    {len(output_files)}")
        if skipped_files:
            report_lines.append(
                f"Files skipped:   {skipped_files} (already existed)")

        if warnings:
            report_lines.append(f"\n── Warnings ─────────────────────")
            for w in warnings:
                report_lines.append(f"  \u2022 {w}")

        if not output_files:
            report_lines.append("\n\u26a0 No CSV files were created.")
            report_lines.append(
                "  The PDF may not contain extractable tables.")
            report_lines.append(
                "  Try changing the Detection Method (Stream or Hybrid).")

        self._extract_btn.setEnabled(True)
        self._extract_btn.setText("Extract to CSV")
        self._progress.setValue(100)
        n_t = total_tables
        n_f = len(output_files)
        self._status_lbl.setText(
            f"Done. {n_t} table{'s' if n_t != 1 else ''} "
            f"extracted to {n_f} file{'s' if n_f != 1 else ''}.")

        self._show_report("\n".join(report_lines), self.output_dir)

    # ======================================================================
    # REPORT PANEL
    # ======================================================================

    def _show_report(self, text: str, output_dir: str):
        if self._report_widget:
            self._report_widget.deleteLater()

        report = QWidget()
        report.setStyleSheet(f"background: {WHITE};")
        self._report_widget = report

        v = QVBoxLayout(report)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(f"border: none; background: {WHITE};")
        inner = QWidget()
        inner.setStyleSheet(f"background: {WHITE};")
        inner_lay = QVBoxLayout(inner)
        inner_lay.setContentsMargins(20, 16, 20, 16)
        lbl = QLabel(text)
        lbl.setWordWrap(True)
        lbl.setStyleSheet(
            f"color: {G700}; font: 12px 'Courier New'; background: transparent;")
        inner_lay.addWidget(lbl)
        inner_lay.addStretch()
        scroll.setWidget(inner)
        v.addWidget(scroll, 1)

        btn_bar = QFrame()
        btn_bar.setFixedHeight(50)
        btn_bar.setStyleSheet(
            f"background: {G100}; border-top: 1px solid {G200};")
        btn_bar_lay = QHBoxLayout(btn_bar)
        btn_bar_lay.setContentsMargins(16, 8, 16, 8)

        open_btn = QPushButton("Open Output Folder")
        open_btn.setFixedSize(180, 34)
        open_btn.setStyleSheet(
            f"QPushButton {{background: {BLUE}; color: white; border-radius: 6px; "
            f"font: 13px 'Segoe UI';}} "
            f"QPushButton:hover {{background: {BLUE_HOVER};}}")
        open_btn.clicked.connect(lambda: self._open_folder(output_dir))
        btn_bar_lay.addWidget(open_btn)

        back_btn = QPushButton("\u2190 Back to Preview")
        back_btn.setFixedSize(160, 34)
        back_btn.setStyleSheet(
            f"QPushButton {{background: transparent; color: {G700}; "
            f"border: 1px solid {G300}; border-radius: 6px; font: 13px 'Segoe UI';}} "
            f"QPushButton:hover {{background: {G200};}}")
        back_btn.clicked.connect(self._back_to_preview)
        btn_bar_lay.addWidget(back_btn)
        btn_bar_lay.addStretch()
        v.addWidget(btn_bar)

        self._right_stack.addWidget(report)
        self._right_stack.setCurrentWidget(report)

    def _back_to_preview(self):
        self._right_stack.setCurrentIndex(0)
        if self._report_widget:
            self._report_widget.deleteLater()
            self._report_widget = None
        if self._doc:
            self._render_page_canvas()

    # ======================================================================
    # OPEN FOLDER / CLEANUP
    # ======================================================================

    def _open_folder(self, path: str):
        try:
            if sys.platform == "darwin":
                subprocess.Popen(["open", path])
            elif sys.platform == "win32":
                os.startfile(path)
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            QMessageBox.critical(self, "Error",
                                 f"Could not open folder:\n{e}")

    def cleanup(self):
        if self._thumb_timer is not None:
            try:
                self._thumb_timer.stop()
            except Exception:
                pass
            self._thumb_timer = None
        if self._doc is not None:
            try:
                self._doc.close()
            except Exception:
                pass
        if self._pldoc is not None:
            try:
                self._pldoc.close()
            except Exception:
                pass
