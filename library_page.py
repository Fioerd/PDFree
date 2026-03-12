"""Document Library / Dashboard — PDFree.

Persistent file library with folders, favorites, trash, search, and upload.
State is stored at ~/.pdfree/library.json.
"""

from __future__ import annotations

import json
import math
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

from PySide6.QtCore import (
    QRectF, QSize, Qt, QTimer, Signal,
)
from PySide6.QtGui import (
    QBrush, QColor, QCursor, QFont, QPainter, QPainterPath, QPen, QIcon, QPixmap,
)
from icons import svg_pixmap, svg_icon
from utils import _make_back_button
import subprocess
from PySide6.QtWidgets import (
    QApplication, QDialog, QDialogButtonBox, QFileDialog, QFrame,
    QGridLayout, QHBoxLayout, QInputDialog, QLabel, QLineEdit,
    QMenu, QMessageBox, QPushButton, QScrollArea, QSizePolicy,
    QStackedWidget, QVBoxLayout, QWidget,
)

from colors import (
    BG, WHITE, G100, G200, G300, G400, G500, G600, G700, G900,
    BLUE, BLUE_ACCENT, RED, GREEN,
)

FOLDER_COLORS = ["#3B82F6", "#10B981", "#F59E0B", "#EF4444",
                 "#8B5CF6", "#EC4899", "#06B6D4", "#F97316"]

# ---------------------------------------------------------------------------
# Persistence
# ---------------------------------------------------------------------------

_STATE_PATH = Path.home() / ".pdfree" / "library.json"


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _age_str(iso: str) -> str:
    """Return a human-readable age like '5m ago', '2h ago', '3d ago'."""
    try:
        dt = datetime.fromisoformat(iso)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        delta = datetime.now(timezone.utc) - dt
        s = int(delta.total_seconds())
        if s < 60:
            return "just now"
        if s < 3600:
            return f"{s // 60}m ago"
        if s < 86400:
            return f"{s // 3600}h ago"
        return f"{s // 86400}d ago"
    except Exception:
        return ""


def _fmt_size(b: int) -> str:
    if b < 1024:
        return f"{b} B"
    if b < 1024 ** 2:
        return f"{b / 1024:.1f} KB"
    return f"{b / 1024 ** 2:.1f} MB"


def _file_size(path: str) -> int:
    try:
        return os.path.getsize(path)
    except OSError:
        return 0


class LibraryState:
    """JSON-backed library state."""

    def __init__(self, on_dirty=None):
        self._path    = _STATE_PATH
        self._data: dict = {"files": [], "folders": []}
        self._on_dirty = on_dirty  # called when data changes; defaults to immediate _save
        self._load()

    def _request_save(self):
        if self._on_dirty is not None:
            self._on_dirty()
        else:
            self._save()

    @property
    def data(self) -> dict:
        """Expose the underlying state dictionary."""
        return self._data

    # ---- IO ---------------------------------------------------------------

    def _load(self):
        try:
            if self._path.exists():
                with open(self._path, encoding="utf-8") as f:
                    loaded = json.load(f)
                    self._data = {
                        "files":   loaded.get("files",   []),
                        "folders": loaded.get("folders", []),
                    }
        except Exception:
            pass

    def _save(self):
        try:
            self._path.parent.mkdir(parents=True, exist_ok=True)
            with open(self._path, "w", encoding="utf-8") as f:
                json.dump(self._data, f, indent=2)
        except Exception:
            pass

    # ---- File tracking ----------------------------------------------------

    def track(self, path: str):
        """Add or update a file entry."""
        path = str(Path(path).resolve())
        name = Path(path).name
        size = _file_size(path)
        for entry in self._data["files"]:
            if entry["path"] == path:
                entry["last_opened"] = _now_iso()
                entry["size"] = size
                self._request_save()
                return
        self._data["files"].append({
            "path":        path,
            "name":        name,
            "last_opened": _now_iso(),
            "size":        size,
            "favorited":   False,
            "trashed":     False,
            "folder":      None,
        })
        self._request_save()

    def set_favorite(self, path: str, val: bool):
        for e in self._data["files"]:
            if e["path"] == path:
                e["favorited"] = val
                break
        self._request_save()

    def trash(self, path: str):
        for e in self._data["files"]:
            if e["path"] == path:
                e["trashed"] = True
                break
        self._request_save()

    def restore(self, path: str):
        for e in self._data["files"]:
            if e["path"] == path:
                e["trashed"] = False
                break
        self._request_save()

    def delete_permanently(self, path: str):
        self._data["files"] = [e for e in self._data["files"] if e["path"] != path]
        self._request_save()

    # ---- Folder management (real filesystem folders) ----------------------

    def add_folder(self, folder_path: str) -> bool:
        """Track a real filesystem folder. Returns False if already tracked."""
        folder_path = str(Path(folder_path).resolve())
        if any(f["path"] == folder_path for f in self._data["folders"]):
            return False
        color = FOLDER_COLORS[len(self._data["folders"]) % len(FOLDER_COLORS)]
        name  = Path(folder_path).name
        self._data["folders"].append({"path": folder_path, "name": name, "color": color})
        self._request_save()
        return True

    def delete_folder(self, folder_path: str):
        """Stop tracking a folder (does NOT delete files from disk)."""
        self._data["folders"] = [
            f for f in self._data["folders"] if f["path"] != folder_path
        ]
        self._request_save()

    def folder_color(self, folder_path: str) -> str:
        for f in self._data["folders"]:
            if f["path"] == folder_path:
                return f.get("color", BLUE)
        return BLUE

    def _scan_folder(self, folder_path: str) -> list[str]:
        """Return sorted list of .pdf paths directly inside folder_path."""
        try:
            d = Path(folder_path)
            if not d.is_dir():
                return []
            return sorted(
                str(p) for p in d.iterdir()
                if p.is_file() and p.suffix.lower() == ".pdf"
            )
        except Exception:
            return []

    # ---- Queries ----------------------------------------------------------

    def _match(self, entry: dict, q: str) -> bool:
        return not q or q in entry["name"].lower()

    def all_active(self, q: str = "") -> list[dict]:
        return [e for e in self._data["files"] if not e["trashed"] and self._match(e, q)]

    def recent(self, n: int = 20, q: str = "") -> list[dict]:
        files = self.all_active(q)
        files.sort(key=lambda e: e.get("last_opened", ""), reverse=True)
        return files[:n]

    def favorites(self, q: str = "") -> list[dict]:
        return [e for e in self.all_active(q) if e["favorited"]]

    def trashed(self) -> list[dict]:
        return [e for e in self._data["files"] if e["trashed"]]

    def in_folder(self, folder_path: str, q: str = "") -> list[dict]:
        """Return entries for all PDFs found inside folder_path on disk."""
        tracked = {e["path"]: e for e in self._data["files"]}
        results: list[dict] = []
        for p in self._scan_folder(folder_path):
            name = Path(p).name
            if q and q not in name.lower():
                continue
            if p in tracked and not tracked[p]["trashed"]:
                results.append(tracked[p])
            else:
                # Build a temporary entry for untracked files
                results.append({
                    "path":        p,
                    "name":        name,
                    "last_opened": "",
                    "size":        _file_size(p),
                    "favorited":   False,
                    "trashed":     False,
                })
        return results

    def folder_stats(self, folder_path: str) -> tuple[int, int]:
        """Return (pdf_count, total_size_bytes) by scanning the real folder."""
        pdfs = self._scan_folder(folder_path)
        total = sum(_file_size(p) for p in pdfs)
        return len(pdfs), total

    def folders(self) -> list[dict]:
        return list(self._data["folders"])


# ---------------------------------------------------------------------------
# Small helper widgets
# ---------------------------------------------------------------------------

class _SectionHdr(QLabel):
    """Small uppercase grey section label."""

    def __init__(self, text: str, parent=None):
        super().__init__(text.upper(), parent)
        self.setStyleSheet(
            f"color: {G400}; font: bold 10px 'Segoe UI'; "
            "background: transparent; letter-spacing: 0.5px;"
        )
        self.setContentsMargins(8, 0, 0, 0)


class _NavBtn(QFrame):
    """Sidebar navigation button with active / hover states."""

    clicked = Signal(str)  # emits key

    def __init__(self, key: str, icon: str, label: str, parent=None):
        super().__init__(parent)
        self._key    = key
        self._active = False
        self.setFixedHeight(38)
        self.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

        lay = QHBoxLayout(self)
        lay.setContentsMargins(10, 0, 10, 0)
        lay.setSpacing(8)

        self._icon_lbl = QLabel(icon)
        self._icon_lbl.setFixedWidth(18)
        self._icon_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(self._icon_lbl)

        self._text_lbl = QLabel(label)
        self._text_lbl.setFont(QFont("Segoe UI", 12))
        lay.addWidget(self._text_lbl, 1)

        self._apply_style()

    def set_active(self, active: bool):
        self._active = active
        self._apply_style()

    def _apply_style(self):
        if self._active:
            bg   = BLUE
            text = WHITE
        else:
            bg   = "transparent"
            text = G700
        self.setStyleSheet(f"""
            QFrame {{
                background: {bg}; border-radius: 8px;
            }}
            QFrame:hover {{
                background: {"" if self._active else G100};
            }}
        """)
        self._icon_lbl.setStyleSheet(
            f"color: {text}; background: transparent; font: 14px;"
        )
        self._text_lbl.setStyleSheet(
            f"color: {text}; background: transparent; font: 12px 'Segoe UI';"
        )

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit(self._key)


class _PdfIcon(QWidget):
    """Small painted PDF page icon."""

    def __init__(self, w: int = 40, h: int = 50, color: str = BLUE_ACCENT,
                 bg: str = G100, parent=None):
        super().__init__(parent)
        self._color = QColor(color)
        self._bg    = QColor(bg)
        self.setFixedSize(w, h)

    def paintEvent(self, _event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.fillRect(self.rect(), self._bg)
        W, H  = self.width(), self.height()
        fold  = W * 0.28
        pen   = QPen(self._color, 1.5)
        p.setPen(pen)
        p.setBrush(QColor(WHITE))

        page = QPainterPath()
        page.moveTo(2, 2)
        page.lineTo(W - fold - 2, 2)
        page.lineTo(W - 2, fold + 2)
        page.lineTo(W - 2, H - 2)
        page.lineTo(2, H - 2)
        page.closeSubpath()
        p.drawPath(page)

        p.setBrush(Qt.BrushStyle.NoBrush)
        ear = QPainterPath()
        ear.moveTo(W - fold - 2, 2)
        ear.lineTo(W - fold - 2, fold + 2)
        ear.lineTo(W - 2, fold + 2)
        p.drawPath(ear)

        lx0, lx1 = int(W * 0.20), int(W * 0.80)
        for ly in [H * 0.52, H * 0.64, H * 0.76]:
            p.drawLine(lx0, int(ly), lx1, int(ly))


# ---------------------------------------------------------------------------
# HeroBanner
# ---------------------------------------------------------------------------

class HeroBanner(QFrame):
    """Green gradient hero card showing the most recently opened PDF."""

    open_req = Signal(str)

    def __init__(self, entry: Optional[dict], parent=None):
        super().__init__(parent)
        self._entry = entry
        self.setFixedHeight(160)
        self.setObjectName("HeroBanner")
        self.setStyleSheet("""
            QFrame#HeroBanner {
                border-radius: 14px;
            }
        """)
        self._build_ui()

    def paintEvent(self, event):
        from PySide6.QtGui import QLinearGradient
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        grad = QLinearGradient(0, 0, self.width(), 0)
        grad.setColorAt(0, QColor("#059669"))
        grad.setColorAt(1, QColor("#10B981"))
        path = QPainterPath()
        path.addRoundedRect(QRectF(self.rect()), 14, 14)
        p.fillPath(path, grad)
        super().paintEvent(event)

    def _build_ui(self):
        lay = QHBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(16)

        left = QVBoxLayout()
        left.setSpacing(6)

        # Chip
        chip = QLabel("RECENTLY EDITED")
        chip.setStyleSheet(
            "color: rgba(255,255,255,0.80); background: rgba(0,0,0,0.15); "
            "border-radius: 6px; font: bold 9px 'Segoe UI'; padding: 2px 8px;"
        )
        chip.setFixedHeight(20)
        chip.setMaximumWidth(130)
        left.addWidget(chip)

        if self._entry is not None:
            entry = self._entry
            name = Path(self._entry["path"]).stem
            age  = _age_str(self._entry.get("last_opened", ""))

            name_lbl = QLabel(name)
            name_lbl.setStyleSheet(
                "color: white; font: bold 18px 'Segoe UI'; background: transparent;"
            )
            name_lbl.setMaximumWidth(400)
            left.addWidget(name_lbl)

            sub = QLabel(f"Last modified {age} by You")
            sub.setStyleSheet(
                "color: rgba(255,255,255,0.75); font: 12px 'Segoe UI'; background: transparent;"
            )
            left.addWidget(sub)
            left.addStretch()

            btns = QHBoxLayout()
            btns.setSpacing(8)

            cont_btn = QPushButton("Continue Editing →")
            cont_btn.setFixedHeight(34)
            cont_btn.setStyleSheet("""
                QPushButton {
                    background: white; color: #059669; border-radius: 8px;
                    font: bold 12px 'Segoe UI'; padding: 0 14px;
                }
                QPushButton:hover { background: #F0FDF4; }
            """)
            cont_btn.clicked.connect(lambda: self.open_req.emit(entry["path"]))
            btns.addWidget(cont_btn)
            btns.addStretch()
            left.addLayout(btns)

        else:
            welcome = QLabel("Welcome to your Document Library")
            welcome.setStyleSheet(
                "color: white; font: bold 18px 'Segoe UI'; background: transparent;"
            )
            left.addWidget(welcome)
            sub = QLabel("Upload your first PDF to get started.")
            sub.setStyleSheet(
                "color: rgba(255,255,255,0.75); font: 12px 'Segoe UI'; background: transparent;"
            )
            left.addWidget(sub)
            left.addStretch()

        lay.addLayout(left, 1)

        # Right info panel
        if self._entry is not None:
            entry = self._entry
            right = QVBoxLayout()
            right.setSpacing(4)
            right.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)

            size_lbl = QLabel(_fmt_size(entry.get("size", 0)))
            size_lbl.setStyleSheet(
                "color: rgba(255,255,255,0.90); font: bold 13px 'Segoe UI'; background: transparent;"
            )
            size_lbl.setAlignment(Qt.AlignmentFlag.AlignRight)
            right.addWidget(size_lbl)

            # Try to get page count
            try:
                import fitz
                doc   = fitz.open(entry["path"])
                pages = len(doc)
                doc.close()
                pg_lbl = QLabel(f"{pages} pages")
                pg_lbl.setStyleSheet(
                    "color: rgba(255,255,255,0.70); font: 11px 'Segoe UI'; background: transparent;"
                )
                pg_lbl.setAlignment(Qt.AlignmentFlag.AlignRight)
                right.addWidget(pg_lbl)
            except Exception:
                pass

            right.addStretch()
            lay.addLayout(right)


# ---------------------------------------------------------------------------
# FolderCard
# ---------------------------------------------------------------------------

class FolderCard(QFrame):
    """200×120 folder card with a colored top bar.

    clicked / delete_req emit the real filesystem path of the folder.
    """

    clicked    = Signal(str)   # folder_path
    delete_req = Signal(str)   # folder_path

    def __init__(self, folder_path: str, name: str, color: str,
                 file_count: int, total_size: int, parent=None):
        super().__init__(parent)
        self._folder_path = folder_path
        self._name  = name
        self._color = color
        self.setFixedSize(200, 120)
        self.setObjectName("FolderCard")
        self._style_normal  = f"""
            QFrame#FolderCard {{
                background: {WHITE}; border-radius: 12px; border: 1px solid {G200};
            }}"""
        self._style_hovered = f"""
            QFrame#FolderCard {{
                background: {G100}; border-radius: 12px; border: 1px solid {G300};
            }}"""
        self.setStyleSheet(self._style_normal)
        self.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.setAttribute(Qt.WidgetAttribute.WA_Hover, True)

        # Colored top bar
        bar = QFrame(self)
        bar.setFixedHeight(7)
        bar.setStyleSheet(
            f"background: {color}; border-radius: 0px; border-top-left-radius: 12px; "
            "border-top-right-radius: 12px; border: none;"
        )
        bar.setGeometry(0, 0, 200, 7)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 14, 12, 10)
        lay.setSpacing(4)

        top_row = QHBoxLayout()
        folder_lbl = QLabel()
        folder_lbl.setPixmap(svg_pixmap("folder", "#4a627b", 22))
        folder_lbl.setStyleSheet("background: transparent;")
        top_row.addWidget(folder_lbl)
        top_row.addStretch()

        self._del_btn = QPushButton("×")
        self._del_btn.setFixedSize(22, 22)
        self._del_btn.setStyleSheet(f"""
            QPushButton {{
                background: {G200}; color: {G600}; border-radius: 11px;
                font: bold 14px; border: none;
            }}
            QPushButton:hover {{ background: {RED}; color: white; }}
        """)
        self._del_btn.hide()
        self._del_btn.clicked.connect(lambda: self.delete_req.emit(self._folder_path))
        top_row.addWidget(self._del_btn)
        lay.addLayout(top_row)

        name_lbl = QLabel(name)
        name_lbl.setStyleSheet(
            f"color: {G900}; font: bold 13px 'Segoe UI'; background: transparent;"
        )
        lay.addWidget(name_lbl)

        sub_lbl = QLabel(
            f"{file_count} file{'s' if file_count != 1 else ''} · {_fmt_size(total_size)}"
        )
        sub_lbl.setStyleSheet(
            f"color: {G400}; font: 11px 'Segoe UI'; background: transparent;"
        )
        lay.addWidget(sub_lbl)
        lay.addStretch()

    def enterEvent(self, _event):
        self.setStyleSheet(self._style_hovered)
        self._del_btn.show()

    def leaveEvent(self, _event):
        self.setStyleSheet(self._style_normal)
        self._del_btn.hide()

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit(self._folder_path)


class _NewFolderCard(QFrame):
    """200×120 dashed card — click to pick a real folder from disk."""

    clicked = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(200, 120)
        self.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self._hovered = False
        self.setAttribute(Qt.WidgetAttribute.WA_Hover, True)

        lay = QVBoxLayout(self)
        lay.setAlignment(Qt.AlignmentFlag.AlignCenter)
        plus = QLabel("+")
        plus.setStyleSheet(
            f"color: {G400}; font: bold 28px 'Segoe UI'; background: transparent;"
        )
        plus.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(plus)
        sub = QLabel("Add Folder from PC")
        sub.setStyleSheet(
            f"color: {G400}; font: 11px 'Segoe UI'; background: transparent;"
        )
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(sub)

    def paintEvent(self, _event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        bg = QColor(G100) if self._hovered else QColor(WHITE)
        path = QPainterPath()
        path.addRoundedRect(QRectF(1, 1, self.width() - 2, self.height() - 2), 12, 12)
        p.fillPath(path, bg)
        pen = QPen(QColor(G300), 1.5, Qt.PenStyle.DashLine)
        p.setPen(pen)
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawPath(path)

    def enterEvent(self, _event):
        self._hovered = True
        self.update()

    def leaveEvent(self, _event):
        self._hovered = False
        self.update()

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()


# ---------------------------------------------------------------------------
# _PdfBadge  —  small red "PDF" label
# ---------------------------------------------------------------------------

class _PdfBadge(QLabel):
    """Red gradient rounded badge with white 'PDF' text."""

    def __init__(self, size: int = 32, parent=None):
        super().__init__("PDF", parent)
        self.setFixedSize(size, size)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        r = max(4, size // 5)
        self.setStyleSheet(f"""
            QLabel {{
                background: qlineargradient(x1:0,y1:0,x2:1,y2:1,
                    stop:0 #DC2626, stop:1 #B91C1C);
                color: white;
                font: bold {max(7, size // 4)}px 'Segoe UI';
                border-radius: {r}px;
                border: none;
            }}
        """)


# ---------------------------------------------------------------------------
# _RecentFileCard  —  horizontal card in the recent strip
# ---------------------------------------------------------------------------

class _RecentFileCard(QFrame):
    """Horizontal card for the 'Recent Files' strip above the table."""

    open_req = Signal(str)

    def __init__(self, entry: dict, parent=None):
        super().__init__(parent)
        self._entry = entry
        self.setObjectName("RFC")
        self.setFixedHeight(68)
        self.setMinimumWidth(220)
        self.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.setAttribute(Qt.WidgetAttribute.WA_Hover, True)
        self.setStyleSheet(f"""
            QFrame#RFC {{
                background: {WHITE};
                border: 1px solid {G200};
                border-radius: 10px;
            }}
            QFrame#RFC:hover {{
                background: #F9FAFB;
            }}
        """)

        lay = QHBoxLayout(self)
        lay.setContentsMargins(14, 0, 14, 0)
        lay.setSpacing(12)

        lay.addWidget(_PdfBadge(38))

        text = QVBoxLayout()
        text.setSpacing(2)
        text.setContentsMargins(0, 0, 0, 0)

        name = entry.get("name", Path(entry.get("path", "")).name)
        name_lbl = QLabel(name)
        name_lbl.setStyleSheet(
            f"color: {G900}; font: 500 13px 'Segoe UI'; background: transparent; border: none;"
        )
        name_lbl.setMaximumWidth(260)
        fm = name_lbl.fontMetrics()
        name_lbl.setText(fm.elidedText(name, Qt.TextElideMode.ElideRight, 260))
        text.addWidget(name_lbl)

        raw = entry.get("size", _file_size(entry.get("path", "")))
        size_lbl = QLabel(f"({_fmt_size(raw)})")
        size_lbl.setStyleSheet(
            f"color: {G400}; font: 12px 'Segoe UI'; background: transparent; border: none;"
        )
        text.addWidget(size_lbl)

        lay.addLayout(text, 1)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.open_req.emit(self._entry["path"])


# ---------------------------------------------------------------------------
# _FileTableRow  —  single row in the file table
# ---------------------------------------------------------------------------

class _FileTableRow(QFrame):
    """Full-width table row matching the screenshot design."""

    open_req   = Signal(str)
    toggle_sel = Signal(str, bool)
    toggle_fav = Signal(str, bool)

    ROW_H = 48

    def __init__(self, entry: dict, selected: bool = False, parent=None):
        super().__init__(parent)
        self._entry    = entry
        self._selected = selected
        self._fav      = entry.get("favorited", False)
        self._exists   = os.path.exists(entry.get("path", ""))

        self.setFixedHeight(self.ROW_H)
        self.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.setAttribute(Qt.WidgetAttribute.WA_Hover, True)
        self._apply_style()
        self._build_ui()

    def _apply_style(self, hover: bool = False):
        if self._selected:
            bg = "#EFF6FF"
        elif hover:
            bg = "#FAFAFA"
        else:
            bg = WHITE
        self.setStyleSheet(f"""
            QFrame {{
                background: {bg};
                border: none;
                border-bottom: 1px solid #F3F4F6;
            }}
        """)

    def _build_ui(self):
        lay = QHBoxLayout(self)
        lay.setContentsMargins(20, 0, 12, 0)
        lay.setSpacing(0)

        # ── Checkbox ────────────────────────────────────────
        self._chk = QPushButton()
        self._chk.setFixedSize(15, 15)
        if self._selected:
            self._chk.setStyleSheet(f"""
                QPushButton {{
                    border: 1.5px solid {BLUE_ACCENT}; border-radius: 3px;
                    background: {BLUE_ACCENT}; color: white; font: bold 10px;
                }}
            """)
            self._chk.setText("✓")
        else:
            self._chk.setStyleSheet(f"""
                QPushButton {{
                    border: 1.5px solid {G300}; border-radius: 3px;
                    background: {WHITE};
                }}
            """)
        self._chk.clicked.connect(self._on_check)
        lay.addWidget(self._chk)
        lay.addSpacing(14)

        # ── PDF badge ────────────────────────────────────────
        lay.addWidget(_PdfBadge(30))
        lay.addSpacing(10)

        # ── Name ─────────────────────────────────────────────
        name = self._entry.get("name", Path(self._entry.get("path", "")).name)
        name_color = "#1D4ED8" if self._selected else G900
        name_lbl = QLabel(name)
        name_lbl.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        name_lbl.setStyleSheet(
            f"color: {name_color}; font: 500 13px 'Segoe UI'; border: none;"
        )
        name_lbl.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        lay.addWidget(name_lbl, 1)

        # ── Last opened ───────────────────────────────────────
        age = _age_str(self._entry.get("last_opened", "")) or "—"
        age_lbl = QLabel(age)
        age_lbl.setFixedWidth(140)
        age_lbl.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        age_lbl.setStyleSheet(
            f"color: {G500}; font: 13px 'Segoe UI'; border: none;"
        )
        lay.addWidget(age_lbl)

        # ── Size ──────────────────────────────────────────────
        size_val = self._entry.get("size", 0) or _file_size(self._entry.get("path", ""))
        size_lbl = QLabel(_fmt_size(size_val))
        size_lbl.setFixedWidth(100)
        size_lbl.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        size_lbl.setStyleSheet(
            f"color: {G700}; font: 13px 'Segoe UI'; border: none;"
        )
        lay.addWidget(size_lbl)

        # ── Star ──────────────────────────────────────────────
        self._star_btn = QPushButton("★" if self._fav else "☆")
        self._star_btn.setFixedSize(28, 28)
        self._star_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent; border: none;
                color: {"#F59E0B" if self._fav else G300}; font: 15px;
            }}
            QPushButton:hover {{ color: #F59E0B; background: #FEF3C7; border-radius: 6px; }}
        """)
        self._star_btn.clicked.connect(self._on_star)
        lay.addWidget(self._star_btn)

        # ── "..." menu ────────────────────────────────────────
        self._menu_btn = QPushButton("···")
        self._menu_btn.setFixedSize(28, 28)
        self._menu_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent; border: none;
                color: {G400}; font: bold 14px 'Segoe UI'; letter-spacing: 1px;
            }}
            QPushButton:hover {{
                background: {G100}; border-radius: 6px; color: {G600};
            }}
        """)
        self._menu_btn.clicked.connect(self._show_menu)
        lay.addWidget(self._menu_btn)

    def enterEvent(self, _event):
        if not self._selected:
            self._apply_style(hover=True)

    def leaveEvent(self, _event):
        self._apply_style()

    def _on_check(self):
        self._selected = not self._selected
        self.toggle_sel.emit(self._entry["path"], self._selected)
        self._apply_style()
        # Rebuild to update checkmark and name weight
        lay = self.layout()
        if lay is not None:
            for i in reversed(range(lay.count())):
                item = lay.itemAt(i)
                if item and item.widget():
                    item.widget().setParent(None)
        self._build_ui()

    def _on_star(self):
        self._fav = not self._fav
        self.toggle_fav.emit(self._entry["path"], self._fav)
        self._star_btn.setText("★" if self._fav else "☆")
        self._star_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent; border: none;
                color: {"#F59E0B" if self._fav else G300}; font: 15px;
            }}
            QPushButton:hover {{ color: #F59E0B; }}
        """)

    def _show_menu(self):
        menu = QMenu(self)
        menu.setStyleSheet(f"""
            QMenu {{
                background: {WHITE}; border: 1px solid {G200};
                border-radius: 8px; padding: 4px;
            }}
            QMenu::item {{
                padding: 6px 20px; color: {G700};
                font: 13px 'Segoe UI'; border-radius: 4px;
            }}
            QMenu::item:selected {{ background: {G100}; }}
            QMenu::separator {{ background: {G200}; height: 1px; margin: 4px 10px; }}
        """)
        menu.addAction("Open", lambda: self.open_req.emit(self._entry["path"]))
        menu.addAction("Show in Explorer", self._show_in_explorer)
        fav_txt = "Remove from Favorites" if self._fav else "Add to Favorites"
        menu.addAction(fav_txt, self._on_star)
        menu.addSeparator()
        menu.addAction("Move to Trash", lambda: self.toggle_sel.emit(self._entry["path"], True))
        pos = self._menu_btn.mapToGlobal(self._menu_btn.rect().bottomLeft())
        menu.exec(pos)

    def _show_in_explorer(self):
        try:
            subprocess.Popen(["explorer", "/select,", str(Path(self._entry["path"]))])
        except Exception:
            pass

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            child = self.childAt(event.pos())
            if child and isinstance(child, QPushButton):
                return
            self.open_req.emit(self._entry["path"])


FileCard = _FileTableRow


# ---------------------------------------------------------------------------
# Selection bar
# ---------------------------------------------------------------------------

class SelectionBar(QFrame):
    """Bottom bar that appears when files are selected."""

    move_req   = Signal()
    open_req   = Signal()
    delete_req = Signal()
    clear_req  = Signal()

    def __init__(self, count: int, parent=None):
        super().__init__(parent)
        self.setFixedHeight(52)
        self.setStyleSheet(f"""
            QFrame {{
                background: {WHITE}; border-top: 1px solid {G200};
            }}
        """)

        lay = QHBoxLayout(self)
        lay.setContentsMargins(20, 0, 20, 0)
        lay.setSpacing(12)

        count_lbl = QLabel(f"{count} Item{'s' if count != 1 else ''} Selected")
        count_lbl.setStyleSheet(
            f"color: {G700}; font: bold 13px 'Segoe UI'; background: transparent;"
        )
        lay.addWidget(count_lbl)
        lay.addStretch()

        for label, sig in [
            ("Show in Explorer", self.move_req),
            ("Open",   self.open_req),
            ("Delete", self.delete_req),
        ]:
            btn = QPushButton(label)
            btn.setFixedHeight(34)
            is_del = label == "Delete"
            btn.setStyleSheet(f"""
                QPushButton {{
                    background: {"#FEE2E2" if is_del else G100};
                    color: {RED if is_del else G700};
                    border-radius: 8px; border: none;
                    font: 12px 'Segoe UI'; padding: 0 14px;
                }}
                QPushButton:hover {{
                    background: {"#FECACA" if is_del else G200};
                }}
            """)
            btn.clicked.connect(sig)
            lay.addWidget(btn)

        close_btn = QPushButton("×")
        close_btn.setFixedSize(32, 32)
        close_btn.setStyleSheet(f"""
            QPushButton {{
                background: {G100}; color: {G600}; border-radius: 16px;
                font: bold 16px; border: none;
            }}
            QPushButton:hover {{ background: {G200}; }}
        """)
        close_btn.clicked.connect(self.clear_req)
        lay.addWidget(close_btn)


# ---------------------------------------------------------------------------
# Trash row
# ---------------------------------------------------------------------------

class _TrashRow(QFrame):
    """Single row in the trash view."""

    restore_req = Signal(str)
    delete_req  = Signal(str)

    def __init__(self, entry: dict, parent=None):
        super().__init__(parent)
        self._path = entry["path"]
        self.setFixedHeight(52)
        self.setStyleSheet(f"""
            QFrame {{
                background: {WHITE}; border-radius: 8px;
                border: 1px solid {G200};
            }}
        """)

        lay = QHBoxLayout(self)
        lay.setContentsMargins(14, 0, 14, 0)
        lay.setSpacing(10)

        icon = QLabel()
        icon.setPixmap(svg_pixmap("file-text", "#6B7280", 18))
        icon.setStyleSheet("background: transparent;")
        lay.addWidget(icon)

        name_lbl = QLabel(entry.get("name", "Unknown"))
        name_lbl.setStyleSheet(
            f"color: {G700}; font: 12px 'Segoe UI'; background: transparent;"
        )
        lay.addWidget(name_lbl, 1)

        size_lbl = QLabel(_fmt_size(entry.get("size", 0)))
        size_lbl.setStyleSheet(
            f"color: {G400}; font: 11px 'Segoe UI'; background: transparent;"
        )
        lay.addWidget(size_lbl)

        restore_btn = QPushButton("Restore")
        restore_btn.setFixedHeight(28)
        restore_btn.setStyleSheet(f"""
            QPushButton {{
                background: {G100}; color: {G700}; border-radius: 6px;
                border: none; font: 11px 'Segoe UI'; padding: 0 10px;
            }}
            QPushButton:hover {{ background: {G200}; }}
        """)
        restore_btn.clicked.connect(lambda: self.restore_req.emit(self._path))
        lay.addWidget(restore_btn)

        del_btn = QPushButton("Delete Forever")
        del_btn.setFixedHeight(28)
        del_btn.setStyleSheet(f"""
            QPushButton {{
                background: #FEE2E2; color: {RED}; border-radius: 6px;
                border: none; font: 11px 'Segoe UI'; padding: 0 10px;
            }}
            QPushButton:hover {{ background: #FECACA; }}
        """)
        del_btn.clicked.connect(lambda: self.delete_req.emit(self._path))
        lay.addWidget(del_btn)


# ---------------------------------------------------------------------------
# LibraryPage
# ---------------------------------------------------------------------------

class LibraryPage(QWidget):
    """Full Document Library dashboard."""

    open_file      = Signal(str)
    back_requested = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._nav_key   = "all"
        self._search_q  = ""
        self._selected: set[str] = set()
        self._nav_btns: dict[str, _NavBtn] = {}
        self._folder_nav_btns: list[_NavBtn] = []

        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.timeout.connect(self._refresh_content)

        self._save_timer = QTimer(self)
        self._save_timer.setSingleShot(True)
        self._save_timer.setInterval(400)
        self._save_timer.timeout.connect(self._flush_save)
        self.state = LibraryState(on_dirty=self._schedule_save)

        self.setStyleSheet(f"background: {BG};")
        self._build_layout()
        self._refresh_content()

    def _schedule_save(self):
        """Start (or restart) the debounce timer; actual write fires 400 ms later."""
        self._save_timer.start()

    def _flush_save(self):
        self.state._save()

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------

    def _build_layout(self):
        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self._build_sidebar())

        main_col = QVBoxLayout()
        main_col.setContentsMargins(0, 0, 0, 0)
        main_col.setSpacing(0)
        main_col.addWidget(self._build_topbar())

        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setFrameShape(QFrame.Shape.NoFrame)
        self._scroll.setStyleSheet(f"""
            QScrollArea {{ background: {BG}; border: none; }}
        """)
        main_col.addWidget(self._scroll, 1)

        # Selection bar wrap
        self._sel_wrap = QWidget()
        self._sel_wrap.setFixedHeight(52)
        self._sel_wrap.hide()
        sw_lay = QVBoxLayout(self._sel_wrap)
        sw_lay.setContentsMargins(0, 0, 0, 0)
        self._sel_bar_holder = sw_lay
        main_col.addWidget(self._sel_wrap)

        root.addLayout(main_col, 1)

    def _build_sidebar(self) -> QFrame:
        sidebar = QFrame()
        sidebar.setFixedWidth(220)
        sidebar.setStyleSheet(f"""
            QFrame {{
                background: {WHITE};
                border-right: 1px solid {G200};
            }}
        """)

        lay = QVBoxLayout(sidebar)
        lay.setContentsMargins(10, 16, 10, 16)
        lay.setSpacing(2)

        # Back button
        back_btn = _make_back_button("Back to Tools", self.back_requested, color=G600)
        lay.addWidget(back_btn)
        lay.addSpacing(8)

        # Logo
        logo_row = QHBoxLayout()
        logo_row.setSpacing(8)
        badge = QLabel("E")
        badge.setFixedSize(28, 28)
        badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        badge.setStyleSheet(f"""
            background: {BLUE}; color: white; font: bold 14px 'Segoe UI';
            border-radius: 6px;
        """)
        logo_row.addWidget(badge)
        logo_txt = QLabel("PDFree")
        logo_txt.setStyleSheet(
            f"color: {G900}; font: bold 16px 'Segoe UI'; background: transparent;"
        )
        logo_row.addWidget(logo_txt, 1)
        lay.addLayout(logo_row)
        lay.addSpacing(16)

        # Static nav
        lay.addWidget(_SectionHdr("Navigation"))
        lay.addSpacing(4)

        for key, icon, label in [
            ("all",       "🗂",  "All Files"),
            ("recent",    "🕐",  "Recent"),
            ("favorites", "★",  "Favorites"),
            ("trash",     "🗑", "Trash"),
        ]:
            btn = _NavBtn(key, icon, label)
            btn.clicked.connect(self._on_nav)
            self._nav_btns[key] = btn
            lay.addWidget(btn)

        lay.addSpacing(16)

        # Folder nav (dynamic)
        self._folder_section_hdr = _SectionHdr("My Folders")
        lay.addWidget(self._folder_section_hdr)
        lay.addSpacing(4)

        self._folder_nav_container = QVBoxLayout()
        self._folder_nav_container.setSpacing(2)
        self._folder_nav_container.setContentsMargins(0, 0, 0, 0)
        lay.addLayout(self._folder_nav_container)

        lay.addStretch()
        return sidebar

    def _build_topbar(self) -> QFrame:
        bar = QFrame()
        bar.setFixedHeight(64)
        bar.setStyleSheet(f"""
            QFrame {{
                background: {WHITE};
                border-bottom: 1px solid {G200};
            }}
        """)

        lay = QHBoxLayout(bar)
        lay.setContentsMargins(20, 0, 20, 0)
        lay.setSpacing(10)

        # Search
        search_frame = QFrame()
        search_frame.setFixedSize(340, 38)
        search_frame.setStyleSheet(f"""
            QFrame {{
                background: {G100}; border-radius: 19px; border: none;
            }}
        """)
        sf_lay = QHBoxLayout(search_frame)
        sf_lay.setContentsMargins(12, 0, 12, 0)
        sf_lay.setSpacing(6)

        mag = QLabel()
        mag.setPixmap(svg_pixmap("search", G400, 15))
        mag.setStyleSheet("background: transparent;")
        sf_lay.addWidget(mag)

        self._search_edit = QLineEdit()
        self._search_edit.setPlaceholderText("Search files…")
        self._search_edit.setStyleSheet(f"""
            QLineEdit {{
                background: transparent; border: none;
                color: {G700}; font: 13px 'Segoe UI';
            }}
        """)
        self._search_edit.textChanged.connect(self._on_search)
        sf_lay.addWidget(self._search_edit, 1)
        lay.addWidget(search_frame)

        lay.addStretch()

        # Upload button
        upload_btn = QPushButton("+ Upload New")
        upload_btn.setFixedSize(140, 38)
        upload_btn.setStyleSheet(f"""
            QPushButton {{
                background: {BLUE}; color: white; border-radius: 19px;
                font: bold 12px 'Segoe UI'; border: none;
            }}
            QPushButton:hover {{ background: {BLUE_ACCENT}; }}
        """)
        upload_btn.clicked.connect(self._upload_new)
        lay.addWidget(upload_btn)

        return bar

    # ------------------------------------------------------------------
    # Content refresh
    # ------------------------------------------------------------------

    def _refresh_content(self):
        self._refresh_folder_nav()
        self._update_nav_active()

        # Build new content widget
        content = QWidget()
        content.setStyleSheet(f"background: {BG};")
        v = QVBoxLayout(content)
        v.setContentsMargins(28, 24, 28, 28)
        v.setSpacing(20)

        q = self._search_q

        if self._nav_key == "all":
            self._build_all_view(v, q)
        elif self._nav_key == "recent":
            self._build_recent_view(v, q)
        elif self._nav_key == "favorites":
            self._build_favorites_view(v, q)
        elif self._nav_key == "trash":
            self._build_trash_view(v)
        elif self._nav_key.startswith("folder:"):
            folder_path = self._nav_key[len("folder:"):]
            self._build_folder_view(v, folder_path, q)

        v.addStretch()
        self._scroll.setWidget(content)

    def _build_all_view(self, lay: QVBoxLayout, q: str):
        # Folders section
        folders = self.state.folders()
        if folders or not q:
            sec_lbl = QLabel("Folders")
            sec_lbl.setStyleSheet(
                f"color: {G700}; font: bold 15px 'Segoe UI'; background: transparent;"
            )
            lay.addWidget(sec_lbl)

            folder_row = QWidget()
            folder_row.setStyleSheet("background: transparent;")
            f_lay = QHBoxLayout(folder_row)
            f_lay.setContentsMargins(0, 0, 0, 0)
            f_lay.setSpacing(12)
            f_lay.setAlignment(Qt.AlignmentFlag.AlignLeft)

            for fd in folders:
                fp = fd["path"]
                count, size = self.state.folder_stats(fp)
                card = FolderCard(fp, fd["name"], fd.get("color", BLUE), count, size)
                card.clicked.connect(lambda p: self._on_nav(f"folder:{p}"))
                card.delete_req.connect(self._delete_folder)
                f_lay.addWidget(card)

            nfc = _NewFolderCard()
            nfc.clicked.connect(self._add_folder)
            f_lay.addWidget(nfc)
            f_lay.addStretch()
            lay.addWidget(folder_row)

        # Recent files
        files = self.state.recent(50, q)
        if files:
            sec_lbl = QLabel("Recent Files")
            sec_lbl.setStyleSheet(
                f"color: {G700}; font: bold 15px 'Segoe UI'; background: transparent;"
            )
            lay.addWidget(sec_lbl)
            lay.addWidget(self._build_recent_strip(files))
            lay.addWidget(self._build_file_table(files))

    def _build_recent_view(self, lay: QVBoxLayout, q: str):
        sec_lbl = QLabel("Recent Files")
        sec_lbl.setStyleSheet(
            f"color: {G700}; font: bold 15px 'Segoe UI'; background: transparent;"
        )
        lay.addWidget(sec_lbl)
        files = self.state.recent(50, q)
        if files:
            lay.addWidget(self._build_recent_strip(files))
            lay.addWidget(self._build_file_table(files))
        else:
            self._add_empty(lay, "No recent files yet.")

    def _build_favorites_view(self, lay: QVBoxLayout, q: str):
        files = self.state.favorites(q)
        sec_lbl = QLabel("Favorites")
        sec_lbl.setStyleSheet(
            f"color: {G700}; font: bold 15px 'Segoe UI'; background: transparent;"
        )
        lay.addWidget(sec_lbl)
        if files:
            lay.addWidget(self._build_file_table(files))
        else:
            self._add_empty(lay, "No favorites yet. Star a file to add it here.")

    def _build_trash_view(self, lay: QVBoxLayout):
        files = self.state.trashed()
        sec_lbl = QLabel("Trash")
        sec_lbl.setStyleSheet(
            f"color: {G700}; font: bold 15px 'Segoe UI'; background: transparent;"
        )
        lay.addWidget(sec_lbl)

        if files:
            for e in files:
                row = _TrashRow(e)
                row.restore_req.connect(self._restore_file)
                row.delete_req.connect(self._delete_permanently)
                lay.addWidget(row)
        else:
            self._add_empty(lay, "Trash is empty.")

    def _build_folder_view(self, lay: QVBoxLayout, folder_path: str, q: str):
        folder_name = Path(folder_path).name
        hdr_row = QWidget()
        hdr_row.setStyleSheet("background: transparent;")
        hdr_h = QHBoxLayout(hdr_row)
        hdr_h.setContentsMargins(0, 0, 0, 0)
        hdr_h.setSpacing(8)
        hdr_icon = QLabel()
        hdr_icon.setPixmap(svg_pixmap("folder", "#4a627b", 20))
        hdr_icon.setStyleSheet("background: transparent;")
        hdr_h.addWidget(hdr_icon)
        hdr = QLabel(folder_name)
        hdr.setStyleSheet(
            f"color: {G900}; font: bold 18px 'Segoe UI'; background: transparent;"
        )
        hdr_h.addWidget(hdr)
        hdr_h.addStretch()
        lay.addWidget(hdr_row)

        # Show full path as subtitle
        path_lbl = QLabel(folder_path)
        path_lbl.setStyleSheet(
            f"color: {G400}; font: 11px 'Segoe UI'; background: transparent;"
        )
        lay.addWidget(path_lbl)

        files = self.state.in_folder(folder_path, q)
        if files:
            lay.addWidget(self._build_file_table(files))
        elif not Path(folder_path).is_dir():
            self._add_empty(lay, "⚠ Folder not found on disk.")
        else:
            self._add_empty(lay, "No PDF files found in this folder.")

    def _add_empty(self, lay: QVBoxLayout, msg: str):
        lbl = QLabel(msg)
        lbl.setStyleSheet(
            f"color: {G400}; font: 14px 'Segoe UI'; background: transparent;"
        )
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(lbl)

    def _build_recent_strip(self, files: list[dict]) -> QWidget:
        strip = QWidget()
        strip.setStyleSheet("background: transparent;")
        lay = QHBoxLayout(strip)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(12)
        lay.setAlignment(Qt.AlignmentFlag.AlignLeft)
        for entry in files[:3]:
            card = _RecentFileCard(entry)
            card.open_req.connect(self._open_file)
            lay.addWidget(card, 1)
        if len(files) < 3:
            lay.addStretch()
        return strip

    def _build_file_table(self, files: list[dict]) -> QWidget:
        container = QFrame()
        container.setObjectName("FileTable")
        container.setStyleSheet(f"""
            QFrame#FileTable {{
                background: {WHITE};
                border: 1px solid {G200};
                border-radius: 12px;
            }}
        """)
        v = QVBoxLayout(container)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)

        # Header row
        hdr = QFrame()
        hdr.setFixedHeight(38)
        hdr.setObjectName("TblHdr")
        hdr.setStyleSheet(f"""
            QFrame#TblHdr {{
                background: #F9FAFB;
                border: none;
                border-bottom: 1px solid {G200};
                border-top-left-radius: 12px;
                border-top-right-radius: 12px;
            }}
        """)
        h = QHBoxLayout(hdr)
        h.setContentsMargins(20, 0, 12, 0)
        h.setSpacing(0)
        h.addSpacing(15 + 14)   # checkbox + gap
        h.addSpacing(30 + 10)   # badge + gap

        def _hdr_lbl(text: str, width: int = 0) -> QLabel:
            lbl = QLabel(text.upper())
            lbl.setStyleSheet(
                f"color: {G400}; font: bold 10px 'Segoe UI'; "
                f"background: transparent; border: none; letter-spacing: 1px;"
            )
            if width:
                lbl.setFixedWidth(width)
            return lbl

        h.addWidget(_hdr_lbl("Name"), 1)
        h.addWidget(_hdr_lbl("Last Opened", 140))
        h.addWidget(_hdr_lbl("Size", 100))
        h.addSpacing(28 + 28)   # star + menu
        v.addWidget(hdr)

        for i, entry in enumerate(files):
            row = _FileTableRow(entry, entry["path"] in self._selected)
            row.open_req.connect(self._open_file)
            row.toggle_sel.connect(self._on_toggle_sel)
            row.toggle_fav.connect(self._on_toggle_fav)
            v.addWidget(row)

        return container

    # ------------------------------------------------------------------
    # Sidebar helpers
    # ------------------------------------------------------------------

    def _refresh_folder_nav(self):
        # Remove old folder buttons
        for btn in self._folder_nav_btns:
            btn.setParent(None)
        self._folder_nav_btns.clear()

        for fd in self.state.folders():
            key = f"folder:{fd['path']}"
            btn = _NavBtn(key, "📁", fd["name"])
            btn.clicked.connect(self._on_nav)
            self._folder_nav_container.addWidget(btn)
            self._folder_nav_btns.append(btn)
            self._nav_btns[key] = btn

    def _update_nav_active(self):
        for key, btn in self._nav_btns.items():
            btn.set_active(key == self._nav_key)

    # ------------------------------------------------------------------
    # Event handlers
    # ------------------------------------------------------------------

    def _on_nav(self, key: str):
        self._nav_key = key
        self._selected.clear()
        self._update_sel_bar()
        self._refresh_content()

    def _on_search(self):
        self._search_q = self._search_edit.text().strip().lower()
        self._search_timer.start(80)

    def _upload_new(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Upload PDFs", "", "PDF files (*.pdf)"
        )
        for p in paths:
            self.state.track(p)
        if paths:
            self._refresh_content()

    def _open_file(self, path: str):
        if not os.path.exists(path):
            QMessageBox.warning(self, "File not found",
                                f"The file no longer exists:\n{path}")
            return
        self.state.track(path)
        self.open_file.emit(path)

    def _add_folder(self):
        """Pick a real folder from disk and add it to the library."""
        folder_path = QFileDialog.getExistingDirectory(
            self, "Select a Folder to Add to Library", str(Path.home())
        )
        if folder_path:
            if not self.state.add_folder(folder_path):
                QMessageBox.information(
                    self, "Already Added",
                    f"'{Path(folder_path).name}' is already in your library."
                )
            self._refresh_content()

    def _delete_folder(self, folder_path: str):
        name = Path(folder_path).name
        reply = QMessageBox.question(
            self, "Remove Folder",
            f"Remove '{name}' from the library?\n\n"
            "The folder and its files will NOT be deleted from your PC.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.state.delete_folder(folder_path)
            if self._nav_key == f"folder:{folder_path}":
                self._nav_key = "all"
            self._refresh_content()

    def _move_selected(self):
        """Open selected files' containing folder in Explorer (no-op placeholder)."""
        # Folders are real filesystem dirs — moving files between them is an
        # OS-level operation. Open the first selected file's folder instead.
        for path in self._selected:
            containing = str(Path(path).parent)
            import subprocess
            try:
                subprocess.Popen(["explorer", containing])
            except Exception:
                pass
            break

    def _trash_selected(self):
        if not self._selected:
            return
        reply = QMessageBox.question(
            self, "Move to Trash",
            f"Move {len(self._selected)} file(s) to trash?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            for path in self._selected:
                self.state.trash(path)
            self._selected.clear()
            self._update_sel_bar()
            self._refresh_content()

    def _restore_file(self, path: str):
        self.state.restore(path)
        self._refresh_content()

    def _delete_permanently(self, path: str):
        reply = QMessageBox.question(
            self, "Delete Permanently",
            "This cannot be undone. Delete permanently?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.state.delete_permanently(path)
            self._refresh_content()

    def _on_toggle_sel(self, path: str, selected: bool):
        if selected:
            self._selected.add(path)
        else:
            self._selected.discard(path)
        self._update_sel_bar()

    def _on_toggle_fav(self, path: str, fav: bool):
        self.state.set_favorite(path, fav)

    def _update_sel_bar(self):
        # Clear existing bar
        while self._sel_bar_holder.count():
            item = self._sel_bar_holder.takeAt(0)
            if item is not None:
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()

        if self._selected:
            bar = SelectionBar(len(self._selected))
            bar.move_req.connect(self._move_selected)
            bar.open_req.connect(self._open_selected)
            bar.delete_req.connect(self._trash_selected)
            bar.clear_req.connect(self._clear_selection)
            self._sel_bar_holder.addWidget(bar)
            self._sel_wrap.show()
        else:
            self._sel_wrap.hide()

    def _open_selected(self):
        for path in list(self._selected):
            self._open_file(path)

    def _clear_selection(self):
        self._selected.clear()
        self._update_sel_bar()
        self._refresh_content()
