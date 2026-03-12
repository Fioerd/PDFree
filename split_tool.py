"""Split Tool – PDF splitting with page preview and cut lines.

PySide6 port. Loaded by main.py when the user clicks "Split".
"""

from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QFrame, QLabel, QPushButton, QLineEdit,
    QScrollArea, QHBoxLayout, QVBoxLayout, QGridLayout,
    QFileDialog, QMessageBox, QProgressBar, QSizePolicy,
    QApplication, QComboBox,
)
from PySide6.QtCore import Qt, QEvent, QObject, QSize
from PySide6.QtGui import (
    QPainter, QColor, QPixmap, QPen, QPainterPath,
    QFont, QCursor, QKeySequence, QShortcut, QIcon,
)
from icons import svg_pixmap, svg_icon
from colors import (
    BLUE, BLUE_HOVER, GREEN, GREEN_HOVER, RED,
    G100, G200, G300, G400, G500, G700, G900, WHITE,
)
from utils import _fitz_pix_to_qpixmap, _WheelToHScroll

try:
    import fitz
except ImportError:
    fitz = None

try:
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import RectangleObject
except ImportError:
    PdfReader = PdfWriter = RectangleObject = None

# ---------------------------------------------------------------------------
# Colors (split_tool-specific)
# ---------------------------------------------------------------------------
CUT_ACTIVE   = "#DC2626"
CUT_INACTIVE = "#CBD5E1"

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _btn(text, bg, hover, text_color=WHITE, border=False, h=38, w=None) -> QPushButton:
    b = QPushButton(text)
    b.setFixedHeight(h)
    if w:
        b.setFixedWidth(w)
    border_s = f"border: 1px solid {G300};" if border else "border: none;"
    b.setStyleSheet(f"""
        QPushButton {{
            background: {bg}; color: {text_color};
            {border_s} border-radius: 6px;
            font: {'bold ' if bg == BLUE or bg == GREEN else ''}13px 'Segoe UI';
            padding: 0 12px;
        }}
        QPushButton:hover {{ background: {hover}; }}
        QPushButton:disabled {{ color: {G300}; background: {G100}; border-color: {G200}; }}
    """)
    return b


# ===========================================================================
# Preview Canvas
# ===========================================================================

class _PreviewCanvas(QWidget):
    def __init__(self, split_tool: "SplitTool", parent=None):
        super().__init__(parent)
        self._st = split_tool
        self.setMinimumSize(300, 300)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.setMouseTracking(True)

    def paintEvent(self, _event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.fillRect(self.rect(), QColor(WHITE))
        st = self._st
        if st._preview_pixmap is None:
            p.setPen(QColor(G400))
            p.setFont(QFont("Segoe UI", 16))
            p.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter,
                       "Load a PDF to see\npage preview here")
            return
        p.drawPixmap(int(st._img_l), int(st._img_t), st._preview_pixmap)
        st._paint_search_highlights(p)
        st._paint_cut(p)

    def mousePressEvent(self, event):
        self._st._start_drag(event)

    def mouseMoveEvent(self, event):
        st = self._st
        # Cursor hint when hovering near cut line
        if st._img_b > st._img_t and not st._dragging:
            ih = st._img_b - st._img_t
            cut_y = st._img_t + st._cut_ratio * ih
            near = abs(event.position().y() - cut_y) <= 15
            self.setCursor(QCursor(Qt.CursorShape.SizeVerCursor if near
                                   else Qt.CursorShape.ArrowCursor))
        st._do_drag(event)

    def mouseReleaseEvent(self, event):
        self._st._end_drag(event)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        st = self._st
        if st.doc and st.total_pages > 0:
            st._show(st.current_page)


# ===========================================================================
# SplitTool
# ===========================================================================

class SplitTool(QWidget):
    THUMB_W = 64

    def __init__(self, parent=None):
        super().__init__(parent)

        if fitz is None or PdfReader is None:
            lay = QVBoxLayout(self)
            lbl = QLabel("⚠  Missing dependencies.\n\n"
                         "Install them with:\n"
                         "  pip install pymupdf pypdf")
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl.setStyleSheet(f"color: {G500}; font: 16px 'Segoe UI';")
            lay.addWidget(lbl)
            return

        # -- State --
        self.pdf_path     = ""
        self.output_dir   = ""
        self.total_pages  = 0
        self.current_page = 0
        self.doc          = None
        self.ranges: list[tuple[int, int]] = []
        self.page_cuts: dict[int, float]   = {}
        self._cut_ratio   = 1.0
        self._dragging    = False
        self._img_l = self._img_r = self._img_t = self._img_b = 0.0
        self._preview_pixmap: QPixmap | None = None
        self._thumb_pixmaps: list[QPixmap]   = []
        self._thumb_frames: list[tuple]      = []  # (frame, img_lbl)
        self._render_scale: float            = 1.0
        self._search_matches: list           = []  # [(page_idx, fitz.Rect), ...]
        self._search_current: int            = -1

        self._build_ui()
        self._setup_shortcuts()

    # ==================================================================
    # BUILD UI
    # ==================================================================

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        body = QWidget()
        body_lay = QHBoxLayout(body)
        body_lay.setContentsMargins(0, 0, 0, 0)
        body_lay.setSpacing(0)
        body_lay.addWidget(self._build_left_panel())
        body_lay.addWidget(self._build_right_panel(), 1)
        root.addWidget(body, 1)

        # ======== HIDDEN COMPATIBILITY WIDGETS =============================
        # ======== HIDDEN COMPATIBILITY WIDGETS =============================
        self.out_entry = QLineEdit()
        self.out_entry.hide()

        self.from_entry = QLineEdit("1")
        self.from_entry.hide()

        self.to_entry = QLineEdit("1")
        self.to_entry.hide()

        self._hidden_ranges_host = QWidget()
        self._hidden_ranges_host.hide()
        self._ranges_layout = QVBoxLayout(self._hidden_ranges_host)
        self._ranges_layout.addStretch()


    def _build_left_panel(self) -> QWidget:
        # ======== LEFT ASIDE (400px, white, right-border) ==================
        left = QWidget()
        left.setFixedWidth(400)
        left.setStyleSheet(
            f"background: {WHITE};"
            f" border-right: 1px solid {G200};"
        )
        left_outer = QVBoxLayout(left)
        left_outer.setContentsMargins(0, 0, 0, 0)
        left_outer.setSpacing(0)

        # Scrollable inner area (top content)
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll_area.setStyleSheet("border: none; background: transparent;")

        scroll_inner = QWidget()
        scroll_inner.setStyleSheet(f"background: {WHITE};")
        left_lay = QVBoxLayout(scroll_inner)
        left_lay.setContentsMargins(25, 25, 25, 12)
        left_lay.setSpacing(0)

        # --- Title row ---
        title_row = QHBoxLayout()
        title_row.setSpacing(12)
        title_row.setContentsMargins(0, 0, 0, 0)

        icon_box = QLabel()
        icon_box.setFixedSize(40, 40)
        icon_box.setPixmap(svg_pixmap("scissors", "#3B82F6", 22))
        icon_box.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_box.setStyleSheet("background: #DBEAFE; border-radius: 8px;")
        title_row.addWidget(icon_box)

        title_lbl = QLabel("Split PDF")
        title_lbl.setStyleSheet(
            f"color: {G900}; font: bold 20px 'Segoe UI';"
            " background: transparent; border: none;"
        )
        title_row.addWidget(title_lbl)
        title_row.addStretch()

        left_lay.addLayout(title_row)
        left_lay.addSpacing(32)

        # --- TARGET FILE section ---
        sec_file = QLabel("TARGET FILE")
        sec_file.setStyleSheet(
            f"color: {G500}; font: bold 12px 'Segoe UI';"
            " letter-spacing: 1.2px; background: transparent; border: none;"
        )
        left_lay.addWidget(sec_file)
        left_lay.addSpacing(8)

        # Dashed drop-zone frame
        drop_zone = QFrame()
        drop_zone.setFixedHeight(56)
        drop_zone.setStyleSheet(
            "background: rgba(249,250,251,128);"
            f" border: 2px dashed {G200};"
            " border-radius: 12px;"
        )
        dz_lay = QHBoxLayout(drop_zone)
        dz_lay.setContentsMargins(10, 0, 10, 0)
        dz_lay.setSpacing(8)

        dz_icon = QLabel()
        dz_icon.setPixmap(svg_pixmap("file-text", "#6B7280", 20))
        dz_icon.setStyleSheet("border: none; background: transparent;")
        dz_lay.addWidget(dz_icon)

        self.file_entry = QLineEdit()
        self.file_entry.setReadOnly(True)
        self.file_entry.setPlaceholderText("No file selected")
        self.file_entry.setStyleSheet(
            "border: none; background: transparent;"
            f" color: {G500}; font: 13px 'Segoe UI';"
        )
        dz_lay.addWidget(self.file_entry, 1)

        browse_pdf = QPushButton("Browse")
        browse_pdf.setFixedHeight(30)
        browse_pdf.setStyleSheet(f"""
            QPushButton {{
                background: {WHITE}; color: {G700};
                border: 1px solid {G300}; border-radius: 8px;
                font: 13px 'Segoe UI'; padding: 0 10px;
            }}
            QPushButton:hover {{ background: {G100}; }}
        """)
        browse_pdf.clicked.connect(self._pick_pdf)
        dz_lay.addWidget(browse_pdf)

        left_lay.addWidget(drop_zone)
        left_lay.addSpacing(32)

        # --- SPLIT MODE section ---
        sec_mode = QLabel("SPLIT MODE")
        sec_mode.setStyleSheet(
            f"color: {G500}; font: bold 12px 'Segoe UI';"
            " letter-spacing: 1.2px; background: transparent; border: none;"
        )
        left_lay.addWidget(sec_mode)
        left_lay.addSpacing(8)

        self.split_mode_combo = QComboBox()
        self.split_mode_combo.addItems([
            "Split by Range",
            "Split Every N Pages",
            "Split in Half",
        ])
        self.split_mode_combo.setFixedHeight(44)
        self.split_mode_combo.setStyleSheet(f"""
            QComboBox {{
                background: rgba(255,255,255,180); color: {G700};
                border: 1px solid {G300}; border-radius: 8px;
                font: 13px 'Segoe UI'; padding: 0 12px;
            }}
            QComboBox::drop-down {{
                border: none; width: 28px;
            }}
            QComboBox:hover {{ border-color: {BLUE}; }}
        """)
        left_lay.addWidget(self.split_mode_combo)
        left_lay.addSpacing(24)

        # --- PAGE RANGES section ---
        sec_ranges = QLabel("PAGE RANGES")
        sec_ranges.setStyleSheet(
            f"color: {G500}; font: bold 12px 'Segoe UI';"
            " letter-spacing: 1.2px; background: transparent; border: none;"
        )
        left_lay.addWidget(sec_ranges)
        left_lay.addSpacing(8)

        # ranges text field + clear button
        ranges_row = QHBoxLayout()
        ranges_row.setSpacing(6)

        self.ranges_edit = QLineEdit()
        self.ranges_edit.setFixedHeight(43)
        self.ranges_edit.setPlaceholderText("e.g. 1-4, 7, 10-12")
        self.ranges_edit.setStyleSheet(f"""
            QLineEdit {{
                background: {WHITE}; color: {G900};
                border: 1px solid {G300}; border-radius: 8px;
                font: 14px 'Courier New', monospace; padding: 0 12px;
            }}
            QLineEdit:focus {{ border-color: {BLUE}; }}
        """)
        ranges_row.addWidget(self.ranges_edit, 1)

        clear_ranges_btn = QPushButton("✕")
        clear_ranges_btn.setFixedSize(43, 43)
        clear_ranges_btn.setToolTip("Clear all ranges")
        clear_ranges_btn.setStyleSheet(f"""
            QPushButton {{
                background: {WHITE}; color: {G500};
                border: 1px solid {G300}; border-radius: 8px;
                font: bold 14px 'Segoe UI';
            }}
            QPushButton:hover {{ background: #FEE2E2; color: {RED}; border-color: {RED}; }}
        """)
        clear_ranges_btn.clicked.connect(lambda: self.ranges_edit.clear())
        ranges_row.addWidget(clear_ranges_btn)

        left_lay.addLayout(ranges_row)
        left_lay.addSpacing(8)

        # ── Quick-add row ──────────────────────────────────────────────
        quick_row = QHBoxLayout()
        quick_row.setSpacing(6)

        _field_style = f"""
            QLineEdit {{
                background: {WHITE}; color: {G900};
                border: 1px solid {G300}; border-radius: 8px;
                font: 13px 'Segoe UI'; padding: 0 8px;
            }}
            QLineEdit:focus {{ border-color: {BLUE}; }}
        """

        from_lbl = QLabel("From")
        from_lbl.setStyleSheet(
            f"color: {G500}; font: 12px 'Segoe UI';"
            " background: transparent; border: none;"
        )
        quick_row.addWidget(from_lbl)

        self._quick_from = QLineEdit("1")
        self._quick_from.setFixedSize(56, 36)
        self._quick_from.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._quick_from.setStyleSheet(_field_style)
        quick_row.addWidget(self._quick_from)

        to_lbl = QLabel("to")
        to_lbl.setStyleSheet(
            f"color: {G500}; font: 12px 'Segoe UI';"
            " background: transparent; border: none;"
        )
        quick_row.addWidget(to_lbl)

        self._quick_to = QLineEdit("1")
        self._quick_to.setFixedSize(56, 36)
        self._quick_to.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._quick_to.setStyleSheet(_field_style)
        quick_row.addWidget(self._quick_to)

        add_range_btn = QPushButton("+ Add Range")
        add_range_btn.setFixedHeight(36)
        add_range_btn.setStyleSheet(f"""
            QPushButton {{
                background: {BLUE}; color: {WHITE};
                border: none; border-radius: 8px;
                font: bold 12px 'Segoe UI'; padding: 0 12px;
            }}
            QPushButton:hover {{ background: {BLUE_HOVER}; }}
        """)
        add_range_btn.clicked.connect(self._quick_add_range)
        quick_row.addWidget(add_range_btn)

        cur_page_btn = QPushButton("Current")
        cur_page_btn.setFixedHeight(36)
        cur_page_btn.setToolTip("Set range to the currently previewed page")
        cur_page_btn.setStyleSheet(f"""
            QPushButton {{
                background: {WHITE}; color: {G700};
                border: 1px solid {G300}; border-radius: 8px;
                font: 12px 'Segoe UI'; padding: 0 10px;
            }}
            QPushButton:hover {{ background: {G100}; }}
        """)
        cur_page_btn.clicked.connect(self._set_quick_to_current)
        quick_row.addWidget(cur_page_btn)

        left_lay.addLayout(quick_row)

        ranges_hint = QLabel("Separate ranges with commas. Use From/to to quickly add a range.")
        ranges_hint.setStyleSheet(
            f"color: {G400}; font: italic 11px 'Segoe UI';"
            " background: transparent; border: none;"
        )
        left_lay.addWidget(ranges_hint)
        left_lay.addSpacing(24)

        # --- OUTPUT FILENAME section ---
        sec_fname = QLabel("OUTPUT FILENAME")
        sec_fname.setStyleSheet(
            f"color: {G500}; font: bold 12px 'Segoe UI';"
            " letter-spacing: 1.2px; background: transparent; border: none;"
        )
        left_lay.addWidget(sec_fname)
        left_lay.addSpacing(8)

        self.filename_entry = QLineEdit("Split_%d")
        self.filename_entry.setFixedHeight(43)
        self.filename_entry.setStyleSheet(f"""
            QLineEdit {{
                background: {WHITE}; color: {G900};
                border: 1px solid {G300}; border-radius: 8px;
                font: 13px 'Segoe UI'; padding: 0 12px;
            }}
            QLineEdit:focus {{ border-color: {BLUE}; }}
        """)
        left_lay.addWidget(self.filename_entry)

        fname_hint = QLabel("%d will be replaced by part number.")
        fname_hint.setStyleSheet(
            f"color: {G400}; font: italic 11px 'Segoe UI';"
            " background: transparent; border: none;"
        )
        left_lay.addWidget(fname_hint)

        # Progress bar (hidden by default)
        left_lay.addSpacing(8)
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setFixedHeight(4)
        self.progress.setTextVisible(False)
        self.progress.setStyleSheet(f"""
            QProgressBar {{
                background: {G200}; border-radius: 2px; border: none;
            }}
            QProgressBar::chunk {{
                background: {GREEN}; border-radius: 2px;
            }}
        """)
        self.progress.hide()
        left_lay.addWidget(self.progress)

        # Status label
        self.status_lbl = QLabel("")
        self.status_lbl.setStyleSheet(
            f"color: {GREEN}; font: 13px 'Segoe UI'; background: transparent; border: none;"
        )
        left_lay.addWidget(self.status_lbl)

        left_lay.addStretch()

        scroll_area.setWidget(scroll_inner)
        left_outer.addWidget(scroll_area, 1)

        # --- Bottom section (split button) ---
        bot_section = QWidget()
        bot_section.setStyleSheet(
            f"background: {WHITE};"
            f" border-top: 1px solid {G200};"
        )
        bot_sec_lay = QVBoxLayout(bot_section)
        bot_sec_lay.setContentsMargins(25, 25, 25, 25)
        bot_sec_lay.setSpacing(0)

        self.split_btn = QPushButton("⊠  Split & Save")
        self.split_btn.setFixedHeight(56)
        self.split_btn.setStyleSheet(f"""
            QPushButton {{
                background: {GREEN}; color: {WHITE};
                border: none; border-radius: 12px;
                font: bold 16px 'Segoe UI';
            }}
            QPushButton:hover {{ background: {GREEN_HOVER}; }}
            QPushButton:disabled {{ background: {G200}; color: {G400}; }}
        """)
        self.split_btn.clicked.connect(self._split_pdf)
        bot_sec_lay.addWidget(self.split_btn)

        left_outer.addWidget(bot_section)
        return left

    def _build_right_panel(self) -> QWidget:
        right = QWidget()
        right.setStyleSheet("background: #EEF0F3; border: none;")
        right_lay = QVBoxLayout(right)
        right_lay.setContentsMargins(0, 0, 0, 0)
        right_lay.setSpacing(0)

        # Canvas top bar (48px, white, border-bottom)
        top_bar = QWidget()
        top_bar.setFixedHeight(48)
        top_bar.setStyleSheet(
            f"background: {WHITE}; border-bottom: 1px solid {G200};"
        )
        top_bar_lay = QHBoxLayout(top_bar)
        top_bar_lay.setContentsMargins(12, 0, 12, 0)
        top_bar_lay.setSpacing(4)

        zoom_out_btn = QPushButton("−")
        zoom_out_btn.setFixedSize(30, 30)
        zoom_out_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent; color: {G500};
                border: none; border-radius: 6px; font: bold 18px;
            }}
            QPushButton:hover {{ background: {G100}; }}
        """)

        self._zoom_lbl = QLabel("100%")
        self._zoom_lbl.setFixedWidth(50)
        self._zoom_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._zoom_lbl.setStyleSheet(
            f"color: {G700}; font: bold 14px 'Segoe UI';"
            " background: transparent; border: none;"
        )

        zoom_in_btn = QPushButton("+")
        zoom_in_btn.setFixedSize(30, 30)
        zoom_in_btn.setStyleSheet(zoom_out_btn.styleSheet())

        top_bar_lay.addWidget(zoom_out_btn)
        top_bar_lay.addWidget(self._zoom_lbl)
        top_bar_lay.addWidget(zoom_in_btn)
        top_bar_lay.addStretch()

        fit_btn = QPushButton("Fit to Width")
        fit_btn.setFixedHeight(30)
        fit_btn.setStyleSheet(f"""
            QPushButton {{
                background: {WHITE}; color: {G700};
                border: 1px solid {G300}; border-radius: 6px;
                font: 13px 'Segoe UI'; padding: 0 10px;
            }}
            QPushButton:hover {{ background: {G100}; }}
        """)
        top_bar_lay.addWidget(fit_btn)

        divider = QFrame()
        divider.setFixedSize(1, 20)
        divider.setStyleSheet(f"background: {G200}; border: none;")
        top_bar_lay.addWidget(divider)

        _nav_btn_style = f"""
            QPushButton {{
                background: transparent; color: {G700};
                border: 1px solid {G300}; border-radius: 6px;
                font: bold 14px 'Segoe UI';
            }}
            QPushButton:hover:enabled {{ background: {G100}; }}
            QPushButton:disabled {{ color: {G300}; border-color: {G200}; }}
        """

        self._nav_prev = QPushButton("‹")
        self._nav_prev.setFixedSize(28, 28)
        self._nav_prev.setStyleSheet(_nav_btn_style)
        self._nav_prev.setEnabled(False)
        self._nav_prev.clicked.connect(self._prev)
        top_bar_lay.addWidget(self._nav_prev)

        self.page_entry = QLineEdit("–")
        self.page_entry.setFixedSize(44, 28)
        self.page_entry.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.page_entry.setStyleSheet(f"""
            QLineEdit {{
                background: {WHITE}; color: {G900};
                border: 1px solid {G300}; border-radius: 6px;
                font: 13px 'Segoe UI';
            }}
            QLineEdit:focus {{ border-color: {BLUE}; }}
        """)
        self.page_entry.returnPressed.connect(self._goto_page)
        top_bar_lay.addWidget(self.page_entry)

        self.page_lbl = QLabel("/ –")
        self.page_lbl.setStyleSheet(
            f"color: {G500}; font: 13px 'Segoe UI';"
            " background: transparent; border: none;"
        )
        top_bar_lay.addWidget(self.page_lbl)

        self._nav_next = QPushButton("›")
        self._nav_next.setFixedSize(28, 28)
        self._nav_next.setStyleSheet(_nav_btn_style)
        self._nav_next.setEnabled(False)
        self._nav_next.clicked.connect(self._next)
        top_bar_lay.addWidget(self._nav_next)

        right_lay.addWidget(top_bar)

        # Search bar (hidden by default, shown on Ctrl+F)
        self._search_bar = QFrame()
        self._search_bar.setFixedHeight(44)
        self._search_bar.setStyleSheet(
            f"QFrame {{ background: {WHITE}; border-bottom: 1px solid {G200}; }}"
        )
        sb_lay = QHBoxLayout(self._search_bar)
        sb_lay.setContentsMargins(12, 0, 12, 0)
        sb_lay.setSpacing(6)

        _srch_icon = QLabel()
        _srch_icon.setPixmap(svg_pixmap("search", G500, 16))
        _srch_icon.setStyleSheet("background: transparent; border: none;")
        sb_lay.addWidget(_srch_icon)

        self._search_input = QLineEdit()
        self._search_input.setPlaceholderText("Search in PDF…")
        self._search_input.setFixedHeight(30)
        self._search_input.setStyleSheet(f"""
            QLineEdit {{
                background: {WHITE}; color: {G900};
                border: 1px solid {G300}; border-radius: 6px;
                font: 13px 'Segoe UI'; padding: 0 10px;
            }}
            QLineEdit:focus {{ border-color: {BLUE}; }}
        """)
        self._search_input.returnPressed.connect(self._search_next_match)
        self._search_input.textChanged.connect(self._search_do)
        sb_lay.addWidget(self._search_input, 1)

        self._search_count_lbl = QLabel("")
        self._search_count_lbl.setFixedWidth(70)
        self._search_count_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._search_count_lbl.setStyleSheet(
            f"color: {G500}; font: 12px 'Segoe UI'; background: transparent; border: none;"
        )
        sb_lay.addWidget(self._search_count_lbl)

        _sb_btn = f"""
            QPushButton {{
                background: transparent; color: {G700};
                border: 1px solid {G300}; border-radius: 6px;
                font: bold 14px 'Segoe UI';
            }}
            QPushButton:hover {{ background: {G100}; }}
        """
        sb_prev = QPushButton("‹")
        sb_prev.setFixedSize(28, 28)
        sb_prev.setStyleSheet(_sb_btn)
        sb_prev.setToolTip("Previous match")
        sb_prev.clicked.connect(self._search_prev_match)
        sb_lay.addWidget(sb_prev)

        sb_next = QPushButton("›")
        sb_next.setFixedSize(28, 28)
        sb_next.setStyleSheet(_sb_btn)
        sb_next.setToolTip("Next match")
        sb_next.clicked.connect(self._search_next_match)
        sb_lay.addWidget(sb_next)

        sb_close = QPushButton("✕")
        sb_close.setFixedSize(28, 28)
        sb_close.setStyleSheet(_sb_btn)
        sb_close.setToolTip("Close search")
        sb_close.clicked.connect(self._close_search)
        sb_lay.addWidget(sb_close)

        self._search_bar.hide()
        right_lay.addWidget(self._search_bar)

        # Preview area (flex-1)
        preview_area = QWidget()
        preview_area.setStyleSheet("background: transparent; border: none;")
        preview_lay = QHBoxLayout(preview_area)
        preview_lay.setContentsMargins(0, 0, 0, 0)
        preview_lay.setSpacing(0)

        self._preview_canvas = _PreviewCanvas(self)
        preview_lay.addWidget(self._preview_canvas, 1)

        # Right float column: prev/next circle buttons
        float_col = QWidget()
        float_col.setFixedWidth(80)
        float_col.setStyleSheet("background: transparent; border: none;")
        float_col_lay = QVBoxLayout(float_col)
        float_col_lay.setContentsMargins(0, 0, 0, 32)
        float_col_lay.setSpacing(8)
        float_col_lay.addStretch()

        circle_btn_style = f"""
            QPushButton {{
                background: {WHITE}; color: {G700};
                border: 1px solid {G200}; border-radius: 24px;
                font: bold 18px 'Segoe UI';
            }}
            QPushButton:hover {{ background: {G100}; }}
            QPushButton:disabled {{ color: {G300}; background: {G100}; }}
        """

        self.btn_prev = QPushButton("↑")
        self.btn_prev.setFixedSize(48, 48)
        self.btn_prev.setStyleSheet(circle_btn_style)
        self.btn_prev.setEnabled(False)
        self.btn_prev.clicked.connect(self._prev)
        float_col_lay.addWidget(self.btn_prev, alignment=Qt.AlignmentFlag.AlignHCenter)

        self.btn_next = QPushButton("↓")
        self.btn_next.setFixedSize(48, 48)
        self.btn_next.setStyleSheet(circle_btn_style)
        self.btn_next.setEnabled(False)
        self.btn_next.clicked.connect(self._next)
        float_col_lay.addWidget(self.btn_next, alignment=Qt.AlignmentFlag.AlignHCenter)

        preview_lay.addWidget(float_col)
        right_lay.addWidget(preview_area, 1)

        # Cut label
        self.cut_lbl = QLabel("")
        self.cut_lbl.setFixedHeight(18)
        self.cut_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.cut_lbl.setStyleSheet(
            f"color: {G500}; font: 11px 'Segoe UI';"
            " background: transparent; border: none;"
        )
        right_lay.addWidget(self.cut_lbl)

        # Thumbnail strip (144px, white, border-top)
        thumb_strip = QWidget()
        thumb_strip.setFixedHeight(144)
        thumb_strip.setStyleSheet(
            f"background: {WHITE}; border-top: 1px solid {G200};"
        )
        thumb_strip_lay = QVBoxLayout(thumb_strip)
        thumb_strip_lay.setContentsMargins(0, 0, 0, 0)
        thumb_strip_lay.setSpacing(0)

        # Header row
        thumb_header = QWidget()
        thumb_header.setStyleSheet("background: transparent; border: none;")
        thumb_header_lay = QHBoxLayout(thumb_header)
        thumb_header_lay.setContentsMargins(24, 8, 24, 4)
        thumb_header_lay.setSpacing(0)

        nav_lbl = QLabel("DOCUMENT NAVIGATION")
        nav_lbl.setStyleSheet(
            f"color: {G500}; font: bold 10px 'Segoe UI';"
            " letter-spacing: 1px; background: transparent; border: none;"
        )
        thumb_header_lay.addWidget(nav_lbl)
        thumb_header_lay.addStretch()

        self._pages_count_lbl = QLabel("— Pages")
        self._pages_count_lbl.setStyleSheet(
            f"color: {G400}; font: 10px 'Segoe UI';"
            " background: transparent; border: none;"
        )
        thumb_header_lay.addWidget(self._pages_count_lbl)

        thumb_strip_lay.addWidget(thumb_header)

        # Scroll area for thumbs
        self._thumb_scroll = QScrollArea()
        self._thumb_scroll.setWidgetResizable(True)
        self._thumb_scroll.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._thumb_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._thumb_scroll.setStyleSheet(
            f"background: transparent; border: none;")

        thumb_inner = QWidget()
        thumb_inner.setStyleSheet("background: transparent;")
        self._thumb_layout = QHBoxLayout(thumb_inner)
        self._thumb_layout.setContentsMargins(24, 0, 24, 0)
        self._thumb_layout.setSpacing(8)
        self._thumb_layout.addStretch()
        self._thumb_scroll.setWidget(thumb_inner)
        self._wheel_filter = _WheelToHScroll(self._thumb_scroll)

        thumb_strip_lay.addWidget(self._thumb_scroll, 1)
        right_lay.addWidget(thumb_strip)

        return right

    # ------------------------------------------------------------------
    def _entry_style(self, editable=False) -> str:
        bg = WHITE if editable else G100
        return f"""
            QLineEdit {{
                background: {bg}; color: {G700};
                border: 1px solid {G200}; border-radius: 6px;
                font: 13px 'Segoe UI'; padding: 0 10px;
            }}
        """

    def _setup_shortcuts(self):
        QShortcut(QKeySequence("Left"),   self).activated.connect(self._prev)
        QShortcut(QKeySequence("Right"),  self).activated.connect(self._next)
        QShortcut(QKeySequence("Ctrl+F"), self).activated.connect(self._open_search)
        QShortcut(QKeySequence("Escape"), self).activated.connect(self._close_search)

    # ==================================================================
    # FILE LOADING
    # ==================================================================

    def _pick_pdf(self):
        p, _ = QFileDialog.getOpenFileName(self, "Select PDF", "", "PDF (*.pdf)")
        if not p:
            return
        self.pdf_path = p
        self.file_entry.setText(p)
        self._load_pdf()

    def _pick_output(self):
        p = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if not p:
            return
        self.output_dir = p
        self.out_entry.setText(p)

    def _load_pdf(self):
        try:
            if fitz is None:
                QMessageBox.critical(
                    self,
                    "Missing dependency",
                    "PyMuPDF (fitz) is not installed. PDF preview is unavailable.",
                )
                return
            if self.doc:
                self.doc.close()
            self.doc = fitz.open(self.pdf_path)
            self.total_pages = len(self.doc)
            self.current_page = 0
            self.ranges.clear()
            self.page_cuts.clear()
            self._cut_ratio = 1.0

            self.from_entry.setText("1")
            self.to_entry.setText(str(self.total_pages))
            self._quick_from.setText("1")
            self._quick_to.setText(str(self.total_pages))

            self._pages_count_lbl.setText(f"{self.total_pages} Pages")
            self.page_lbl.setText(f"/ {self.total_pages}")
            self.page_entry.setText("1")
            self._zoom_lbl.setText("100%")
            self.status_lbl.setText(f"{self.total_pages} pages loaded.")
            self._rebuild_cards()
            self._render_thumbs()
            self._show(0)
        except Exception as e:
            self.total_pages = 0
            QMessageBox.critical(self, "Error", f"Could not load PDF:\n{e}")

    # ==================================================================
    # PREVIEW
    # ==================================================================

    def _show(self, idx):
        if not self.doc or idx < 0 or idx >= self.total_pages:
            return
        if fitz is None:
            return
        self.current_page = idx

        cw = max(self._preview_canvas.width(), 300)
        ch = max(self._preview_canvas.height(), 300)

        page = self.doc[idx]
        ratio = page.rect.height / page.rect.width
        fw, fh = cw - 30, ch - 30
        rw = fw if fw * ratio <= fh else int(fh / ratio)
        rw = max(rw, 100)

        s = rw / page.rect.width
        self._render_scale = s
        pix = page.get_pixmap(matrix=fitz.Matrix(s, s), alpha=False)
        self._preview_pixmap = _fitz_pix_to_qpixmap(pix)
        iw, ih = self._preview_pixmap.width(), self._preview_pixmap.height()

        cx, cy = cw // 2, ch // 2
        self._img_l = cx - iw / 2
        self._img_r = cx + iw / 2
        self._img_t = cy - ih / 2
        self._img_b = cy + ih / 2

        self._cut_ratio = self.page_cuts.get(idx, 1.0)
        self._preview_canvas.update()

        self.page_lbl.setText(f"/ {self.total_pages}")
        self.page_entry.setText(str(idx + 1))
        self.btn_prev.setEnabled(idx > 0)
        self.btn_next.setEnabled(idx < self.total_pages - 1)
        self._nav_prev.setEnabled(idx > 0)
        self._nav_next.setEnabled(idx < self.total_pages - 1)
        self._hl_thumb(idx)

    def _prev(self):
        if self.current_page > 0:
            self._show(self.current_page - 1)

    def _next(self):
        if self.current_page < self.total_pages - 1:
            self._show(self.current_page + 1)

    def _goto_page(self):
        if self.total_pages == 0:
            return
        try:
            n = int(self.page_entry.text().strip())
        except ValueError:
            self.page_entry.setText(str(self.current_page + 1))
            return
        self._show(max(0, min(n - 1, self.total_pages - 1)))

    # ==================================================================
    # SEARCH
    # ==================================================================

    def _open_search(self):
        self._search_bar.show()
        self._search_input.setFocus()
        self._search_input.selectAll()

    def _close_search(self):
        self._search_bar.hide()
        self._search_matches.clear()
        self._search_current = -1
        self._search_count_lbl.setText("")
        self._preview_canvas.update()

    def _search_do(self):
        query = self._search_input.text().strip()
        self._search_matches.clear()
        self._search_current = -1
        if not query or not self.doc:
            self._search_count_lbl.setText("")
            self._preview_canvas.update()
            return
        for pg_idx in range(self.total_pages):
            rects = self.doc[pg_idx].search_for(query)
            for r in rects:
                self._search_matches.append((pg_idx, r))
        total = len(self._search_matches)
        if total == 0:
            self._search_count_lbl.setText("No results")
            self._preview_canvas.update()
            return
        self._search_current = 0
        self._search_goto(0)

    def _search_next_match(self):
        if not self._search_matches:
            self._search_do()
            return
        self._search_current = (self._search_current + 1) % len(self._search_matches)
        self._search_goto(self._search_current)

    def _search_prev_match(self):
        if not self._search_matches:
            return
        self._search_current = (self._search_current - 1) % len(self._search_matches)
        self._search_goto(self._search_current)

    def _search_goto(self, idx: int):
        total = len(self._search_matches)
        self._search_count_lbl.setText(f"{idx + 1} / {total}")
        page_idx, _ = self._search_matches[idx]
        if page_idx != self.current_page:
            self._show(page_idx)
        else:
            self._preview_canvas.update()

    def _paint_search_highlights(self, p: QPainter):
        if not self._search_matches or self._img_b <= self._img_t:
            return
        for i, (pg_idx, r) in enumerate(self._search_matches):
            if pg_idx != self.current_page:
                continue
            x = self._img_l + r.x0 * self._render_scale
            y = self._img_t + r.y0 * self._render_scale
            w = (r.x1 - r.x0) * self._render_scale
            h = (r.y1 - r.y0) * self._render_scale
            if i == self._search_current:
                col = QColor("#FF8C00")
                col.setAlpha(180)
                p.fillRect(int(x), int(y), int(w), int(h), col)
                p.setPen(QPen(QColor("#E65C00"), 2))
                p.drawRect(int(x), int(y), int(w), int(h))
            else:
                col = QColor("#FFD700")
                col.setAlpha(130)
                p.fillRect(int(x), int(y), int(w), int(h), col)

    # ==================================================================
    # CUT LINE
    # ==================================================================

    def _paint_cut(self, p: QPainter):
        if self._img_b <= self._img_t:
            return
        has = self._cut_ratio < 0.98
        col = QColor(CUT_ACTIVE if has else CUT_INACTIVE)
        lw = 3 if has else 2

        ih = self._img_b - self._img_t
        y = self._img_t + self._cut_ratio * ih

        pen = QPen(col, lw)
        if not has:
            pen.setStyle(Qt.PenStyle.DashLine)
        p.setPen(pen)
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawLine(int(self._img_l), int(y), int(self._img_r), int(y))

        # Triangle handles at left and right edges
        hs = 9
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(col)
        for hx, d in [(self._img_l, 1), (self._img_r, -1)]:
            path = QPainterPath()
            path.moveTo(hx, y - hs)
            path.lineTo(hx + d * hs, y)
            path.lineTo(hx, y + hs)
            path.closeSubpath()
            p.fillPath(path, col)

        # Label
        p.setPen(QPen(col))
        p.setFont(QFont("Segoe UI", 10))
        txt = f"✂ {int(self._cut_ratio * 100)}%" if has else "▼ full page"
        p.drawText(int(self._img_r) + 10, int(y) + 4, txt)

        # Zone bars (blue = top part, orange = bottom part)
        if has:
            bx = int(self._img_l) - 6
            p.fillRect(bx, int(self._img_t), 4,
                       int(y - self._img_t), QColor(BLUE))
            p.fillRect(bx, int(y), 4,
                       int(self._img_b - y), QColor("#F97316"))

        self.cut_lbl.setText(
            "✂ Above → current part | Below → next part" if has
            else "Drag the line up to cut this page")

    def _start_drag(self, event):
        if self._img_b <= self._img_t:
            return
        y = event.position().y()
        ih = self._img_b - self._img_t
        cut_y = self._img_t + self._cut_ratio * ih
        if abs(y - cut_y) <= 15:
            self._dragging = True

    def _do_drag(self, event):
        if not self._dragging or self._img_b <= self._img_t:
            return
        y = max(self._img_t + 5, min(event.position().y(), self._img_b))
        self._cut_ratio = (y - self._img_t) / (self._img_b - self._img_t)
        if self._cut_ratio >= 0.98:
            self._cut_ratio = 1.0
        if self._cut_ratio < 0.98:
            self.page_cuts[self.current_page] = self._cut_ratio
        else:
            self.page_cuts.pop(self.current_page, None)
        self._preview_canvas.update()

    def _end_drag(self, event):
        if not self._dragging:
            return
        self._dragging = False
        self._preview_canvas.setCursor(QCursor(Qt.CursorShape.ArrowCursor))
        if self._cut_ratio < 0.98:
            self.page_cuts[self.current_page] = self._cut_ratio
        else:
            self.page_cuts.pop(self.current_page, None)
        self._rebuild_cards()

    # ==================================================================
    # THUMBNAILS
    # ==================================================================

    def _render_thumbs(self):
        # Clear existing thumbnails (keep the trailing stretch)
        while self._thumb_layout.count() > 1:
            item = self._thumb_layout.takeAt(0)
            if item is not None:
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()
        self._thumb_pixmaps.clear()
        self._thumb_frames.clear()

        if not self.doc:
            return

        if fitz is None:
            return

        THUMB_H = 90
        for i in range(self.total_pages):
            page = self.doc[i]
            s = self.THUMB_W / page.rect.width
            pix_raw = page.get_pixmap(matrix=fitz.Matrix(s, s), alpha=False)
            pm = _fitz_pix_to_qpixmap(pix_raw)
            # Scale to fixed 64×90 thumbnail size
            pm = pm.scaled(self.THUMB_W, THUMB_H,
                           Qt.AspectRatioMode.KeepAspectRatio,
                           Qt.TransformationMode.SmoothTransformation)
            self._thumb_pixmaps.append(pm)

            frame = QFrame()
            frame.setStyleSheet("background: transparent; border: none;")
            frame.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
            fl = QVBoxLayout(frame)
            fl.setContentsMargins(2, 2, 2, 2)
            fl.setSpacing(2)

            img_lbl = QLabel()
            img_lbl.setPixmap(pm)
            img_lbl.setFixedSize(self.THUMB_W, THUMB_H)
            img_lbl.setStyleSheet(f"border: 2px solid {G300}; background: white;")
            img_lbl.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

            num_lbl = QLabel(str(i + 1))
            num_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            num_lbl.setStyleSheet(
                f"color: {G500}; font: 9px 'Segoe UI';"
                " background: transparent; border: none;")

            fl.addWidget(img_lbl, alignment=Qt.AlignmentFlag.AlignCenter)
            fl.addWidget(num_lbl)

            frame.mousePressEvent  = lambda e, idx=i: self._show(idx)
            img_lbl.mousePressEvent = lambda e, idx=i: self._show(idx)

            # Insert before the trailing stretch
            pos = self._thumb_layout.count() - 1
            self._thumb_layout.insertWidget(pos, frame)
            self._thumb_frames.append((frame, img_lbl))

            if i % 10 == 0:
                QApplication.processEvents()

    def _hl_thumb(self, idx):
        for i, (_, img_lbl) in enumerate(self._thumb_frames):
            if i == idx:
                img_lbl.setStyleSheet(
                    f"border: 2px solid {BLUE}; background: white;")
            else:
                img_lbl.setStyleSheet(
                    f"border: 2px solid {G300}; background: white;")

    # ==================================================================
    # RANGE MANAGEMENT
    # ==================================================================

    def _add_range(self):
        try:
            s, e = int(self.from_entry.text()), int(self.to_entry.text())
        except ValueError:
            QMessageBox.critical(self, "Error", "Please enter valid page numbers.")
            return
        if s <= 0 or e <= 0:
            QMessageBox.critical(self, "Error", "Page numbers must be > 0.")
            return
        if s > e:
            QMessageBox.critical(self, "Error", "Start page must be ≤ end page.")
            return
        if e > self.total_pages:
            QMessageBox.critical(self, "Error",
                f"Page {e} doesn't exist. PDF has {self.total_pages} pages.")
            return

        self.ranges.append((s, e))
        self._rebuild_cards()
        self.from_entry.setText(str(e))
        self.to_entry.setText(str(self.total_pages))

    def _delete_range(self, idx):
        if 0 <= idx < len(self.ranges):
            self.ranges.pop(idx)
            self._rebuild_cards()

    def _rebuild_cards(self):
        # Remove all except the trailing stretch
        while self._ranges_layout.count() > 1:
            item = self._ranges_layout.takeAt(0)
            if item is not None:
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()

        for i, (s, e) in enumerate(self.ranges):
            txt = (f"Part {i + 1}: Pages {s}–{e}" if s != e
                   else f"Part {i + 1}: Page {s}")
            cuts = []
            if (e - 1) in self.page_cuts and s != e:
                cuts.append(f"p.{e} ✂{int(self.page_cuts[e-1]*100)}%↑")
            if (s - 1) in self.page_cuts and s != e:
                cuts.append(f"p.{s} ✂{int(self.page_cuts[s-1]*100)}%↓")
            if cuts:
                txt += f"  ({', '.join(cuts)})"

            card = QFrame()
            card.setFixedHeight(44)
            card.setStyleSheet(
                f"background: {WHITE}; border: 1px solid {G300};"
                " border-radius: 10px;")
            card_lay = QHBoxLayout(card)
            card_lay.setContentsMargins(14, 0, 6, 0)

            lbl = QLabel(txt)
            lbl.setStyleSheet(
                f"color: {G900}; font: 13px 'Segoe UI';"
                " background: transparent; border: none;")
            card_lay.addWidget(lbl, 1)

            del_btn = QPushButton()
            del_btn.setIcon(QIcon(svg_pixmap("trash-2", G400, 16)))
            del_btn.setIconSize(QSize(16, 16))
            del_btn.setFixedSize(34, 34)
            del_btn.setStyleSheet(f"""
                QPushButton {{
                    background: transparent;
                    border: none; border-radius: 6px;
                }}
                QPushButton:hover {{ background: {G100}; }}
            """)
            del_btn.clicked.connect(lambda checked=False, ii=i: self._delete_range(ii))
            card_lay.addWidget(del_btn)

            self._ranges_layout.insertWidget(i, card)

    def _quick_add_range(self):
        try:
            a = int(self._quick_from.text().strip())
            b = int(self._quick_to.text().strip())
        except ValueError:
            return
        if a <= 0 or b <= 0 or a > b:
            return
        token = str(a) if a == b else f"{a}-{b}"
        current = self.ranges_edit.text().strip()
        self.ranges_edit.setText(f"{current}, {token}" if current else token)
        # Old To becomes new From
        self._quick_from.setText(str(b))

    def _set_quick_to_current(self):
        if self.total_pages == 0:
            return
        self._quick_to.setText(str(self.current_page + 1))

    def _page_to_bis(self):
        if self.total_pages == 0:
            return
        self.to_entry.setText(str(self.current_page + 1))

    # ==================================================================
    # SPLIT
    # ==================================================================

    def _split_pdf(self):
        if not self.pdf_path:
            QMessageBox.critical(self, "Error", "Please select a PDF file first.")
            return
        ranges_text = self.ranges_edit.text().strip()
        if not ranges_text:
            QMessageBox.critical(self, "Error", "Please enter page ranges (e.g. 1-4, 7, 10-12).")
            return
        try:
            parsed = []
            for part in ranges_text.split(","):
                part = part.strip()
                if not part:
                    continue
                if "-" in part:
                    a, b = part.split("-", 1)
                    parsed.append((int(a.strip()), int(b.strip())))
                else:
                    n = int(part)
                    parsed.append((n, n))
        except ValueError:
            QMessageBox.critical(self, "Error", "Invalid range format. Use e.g. 1-4, 7, 10-12")
            return
        if not parsed:
            QMessageBox.critical(self, "Error", "No valid ranges entered.")
            return
        self.ranges = parsed
        if not self.output_dir:
            self.output_dir = QFileDialog.getExistingDirectory(self, "Select Output Folder")
            if not self.output_dir:
                return
        self.split_btn.setEnabled(False)
        self.progress.show()
        try:
            self._do_split()
        except Exception as ex:
            QMessageBox.critical(self, "Error", f"Split failed:\n{ex}")
            self.status_lbl.setText("Error during split.")
        finally:
            self.split_btn.setEnabled(True)

    def _do_split(self):
        if fitz is None:
            QMessageBox.critical(self, "Missing dependency",
                                 "PyMuPDF (fitz) is not installed.")
            return

        A4_W, A4_H = 595.28, 841.89
        src = fitz.open(self.pdf_path)
        base = Path(self.pdf_path).stem
        n = len(self.ranges)
        self.progress.setValue(0)

        for i, (s, e) in enumerate(self.ranges):
            out = fitz.open()
            for pn in range(s - 1, e):
                src_page = src[pn]
                new_page = out.new_page(width=A4_W, height=A4_H)
                sr = src_page.rect
                scale = A4_W / sr.width if sr.width > 0 else 1.0

                if pn in self.page_cuts and s != e:
                    cr = self.page_cuts[pn]
                    cut_y = cr * sr.height          # fitz y from top
                    first, last = pn == s - 1, pn == e - 1

                    if first and last:
                        new_page.show_pdf_page(new_page.rect, src, pn)
                    elif last:
                        # Top portion (above cut line) → placed at top of A4
                        clip = fitz.Rect(sr.x0, sr.y0, sr.x1, sr.y0 + cut_y)
                        dest_h = min(cut_y * scale, A4_H)
                        new_page.show_pdf_page(
                            fitz.Rect(0, 0, A4_W, dest_h), src, pn, clip=clip)
                    elif first:
                        # Bottom portion (below cut line) → placed at top of A4
                        clip = fitz.Rect(sr.x0, sr.y0 + cut_y, sr.x1, sr.y1)
                        dest_h = min((sr.height - cut_y) * scale, A4_H)
                        new_page.show_pdf_page(
                            fitz.Rect(0, 0, A4_W, dest_h), src, pn, clip=clip)
                    else:
                        new_page.show_pdf_page(new_page.rect, src, pn)
                else:
                    new_page.show_pdf_page(new_page.rect, src, pn)

            tmpl = self.filename_entry.text().strip() or f"{base}_part%d"
            fname = tmpl.replace("%d", str(i + 1))
            if not fname.lower().endswith(".pdf"):
                fname += ".pdf"
            out.save(str(Path(self.output_dir) / fname))
            out.close()

            self.progress.setValue(int((i + 1) / n * 100))
            self.status_lbl.setText(
                f"Part {i + 1} of {n} created (p. {s}–{e})...")
            QApplication.processEvents()

        src.close()
        self.status_lbl.setText(f"Done! {n} parts created.")

    # ==================================================================
    # CLEANUP
    # ==================================================================

    def cleanup(self):
        if self.doc:
            self.doc.close()
