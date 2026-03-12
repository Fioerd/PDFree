"""Lucide SVG icon library for PDF Suite.

Usage:
    from icons import svg_pixmap, svg_icon, is_svg_icon

    # QPixmap (for QLabel.setPixmap or QPainter.drawPixmap)
    px = svg_pixmap("arrow-left", "#374151", 16)

    # QIcon (for QPushButton.setIcon / QAction)
    btn.setIcon(svg_icon("search", "#6B7280", 16))

    # Check before use
    if is_svg_icon(name): ...
"""

from PySide6.QtCore import QByteArray, QSize
from PySide6.QtGui import QColor, QIcon, QPixmap, QPainter
from PySide6.QtSvg import QSvgRenderer

_PIXMAP_CACHE: dict[tuple, QPixmap] = {}

# ---------------------------------------------------------------------------
# SVG inner paths — Lucide Icons (ISC License), stroke-width independent
# ---------------------------------------------------------------------------
_SVGS: dict[str, str] = {
    # Navigation
    "arrow-left": '<path d="m12 19-7-7 7-7"/><path d="M19 12H5"/>',
    "chevron-left": '<path d="m15 18-6-6 6-6"/>',
    "chevron-right": '<path d="m9 18 6-6-6-6"/>',

    # Files & Documents
    "file-text": (
        '<path d="M6 22a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h8a2.4 2.4 0 0 1 1.704.706'
        'l3.588 3.588A2.4 2.4 0 0 1 20 8v12a2 2 0 0 1-2 2z"/>'
        '<path d="M14 2v5a1 1 0 0 0 1 1h5"/>'
        '<path d="M10 9H8"/><path d="M16 13H8"/><path d="M16 17H8"/>'
    ),
    "file-plus": (
        '<path d="M6 22a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h8a2.4 2.4 0 0 1 1.704.706'
        'l3.588 3.588A2.4 2.4 0 0 1 20 8v12a2 2 0 0 1-2 2z"/>'
        '<path d="M14 2v5a1 1 0 0 0 1 1h5"/>'
        '<path d="M9 15h6"/><path d="M12 18v-6"/>'
    ),
    "file-minus": (
        '<path d="M6 22a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h8a2.4 2.4 0 0 1 1.704.706'
        'l3.588 3.588A2.4 2.4 0 0 1 20 8v12a2 2 0 0 1-2 2z"/>'
        '<path d="M14 2v5a1 1 0 0 0 1 1h5"/>'
        '<path d="M9 15h6"/>'
    ),
    "file-search": (
        '<path d="M6 22a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h8a2.4 2.4 0 0 1 1.704.706'
        'l3.588 3.588A2.4 2.4 0 0 1 20 8v12a2 2 0 0 1-2 2z"/>'
        '<path d="M14 2v5a1 1 0 0 0 1 1h5"/>'
        '<circle cx="11.5" cy="14.5" r="2.5"/>'
        '<path d="M13.3 16.3 15 18"/>'
    ),
    "file-output": (
        '<path d="M4.226 20.925A2 2 0 0 0 6 22h12a2 2 0 0 0 2-2V8'
        'a2.4 2.4 0 0 0-.706-1.706l-3.588-3.588A2.4 2.4 0 0 0 14 2H6a2 2 0 0 0-2 2v3.127"/>'
        '<path d="M14 2v5a1 1 0 0 0 1 1h5"/>'
        '<path d="m5 11-3 3"/><path d="m5 17-3-3h10"/>'
    ),
    "book-open": (
        '<path d="M12 7v14"/>'
        '<path d="M3 18a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1h5a4 4 0 0 1 4 4'
        ' 4 4 0 0 1 4-4h5a1 1 0 0 1 1 1v13a1 1 0 0 1-1 1h-6'
        'a3 3 0 0 0-3 3 3 3 0 0 0-3-3z"/>'
    ),

    # Folders
    "folder": (
        '<path d="M20 20a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.9'
        'a2 2 0 0 1-1.69-.9L9.6 3.9A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2Z"/>'
    ),

    # Search & Zoom
    "search": '<path d="m21 21-4.34-4.34"/><circle cx="11" cy="11" r="8"/>',
    "zoom-in": (
        '<circle cx="11" cy="11" r="8"/>'
        '<line x1="21" x2="16.65" y1="21" y2="16.65"/>'
        '<line x1="11" x2="11" y1="8" y2="14"/>'
        '<line x1="8" x2="14" y1="11" y2="11"/>'
    ),
    "zoom-out": (
        '<circle cx="11" cy="11" r="8"/>'
        '<line x1="21" x2="16.65" y1="21" y2="16.65"/>'
        '<line x1="8" x2="14" y1="11" y2="11"/>'
    ),

    # Editing
    "pen-line": (
        '<path d="M13 21h8"/>'
        '<path d="M21.174 6.812a1 1 0 0 0-3.986-3.987L3.842 16.174'
        'a2 2 0 0 0-.5.83l-1.321 4.352a.5.5 0 0 0 .623.622'
        'l4.353-1.32a2 2 0 0 0 .83-.497z"/>'
    ),
    "eraser": (
        '<path d="M21 21H8a2 2 0 0 1-1.42-.587l-3.994-3.999'
        'a2 2 0 0 1 0-2.828l10-10a2 2 0 0 1 2.829 0l5.999 6'
        'a2 2 0 0 1 0 2.828L12.834 21"/>'
        '<path d="m5.082 11.09 8.828 8.828"/>'
    ),
    "rotate-cw": (
        '<path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/>'
        '<path d="M21 3v5h-5"/>'
    ),
    "copy": (
        '<rect width="14" height="14" x="8" y="8" rx="2" ry="2"/>'
        '<path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/>'
    ),

    # Toolbar tools
    "crosshair": (
        '<circle cx="12" cy="12" r="10"/>'
        '<line x1="22" x2="18" y1="12" y2="12"/>'
        '<line x1="6" x2="2" y1="12" y2="12"/>'
        '<line x1="12" x2="12" y1="6" y2="2"/>'
        '<line x1="12" x2="12" y1="22" y2="18"/>'
    ),
    "hand": (
        '<path d="M18 11V6a2 2 0 0 0-2-2a2 2 0 0 0-2 2"/>'
        '<path d="M14 10V4a2 2 0 0 0-2-2a2 2 0 0 0-2 2v2"/>'
        '<path d="M10 10.5V6a2 2 0 0 0-2-2a2 2 0 0 0-2 2v8"/>'
        '<path d="M18 8a2 2 0 1 1 4 0v6a8 8 0 0 1-8 8h-2'
        'c-2.8 0-4.5-.86-5.99-2.34l-3.6-3.6'
        'a2 2 0 0 1 2.83-2.82L7 15"/>'
    ),
    "minus-circle": '<circle cx="12" cy="12" r="10"/><path d="M8 12h8"/>',
    "plus-circle":  '<circle cx="12" cy="12" r="10"/><path d="M12 8v8M8 12h8"/>',

    # Actions
    "scissors": (
        '<circle cx="6" cy="6" r="3"/>'
        '<path d="M8.12 8.12 12 12"/>'
        '<path d="M20 4 8.12 15.88"/>'
        '<circle cx="6" cy="18" r="3"/>'
        '<path d="M14.8 14.8 20 20"/>'
    ),
    "merge": (
        '<path d="m8 6 4-4 4 4"/>'
        '<path d="M12 2v10.3a4 4 0 0 1-1.172 2.872L4 22"/>'
        '<path d="m20 22-5-5"/>'
    ),
    "upload": (
        '<path d="M12 3v12"/>'
        '<path d="m17 8-5-5-5 5"/>'
        '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>'
    ),
    "download": (
        '<path d="M12 15V3"/>'
        '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>'
        '<path d="m7 10 5 5 5-5"/>'
    ),

    # View
    "eye": (
        '<path d="M2.062 12.348a1 1 0 0 1 0-.696'
        ' 10.75 10.75 0 0 1 19.876 0'
        ' 1 1 0 0 1 0 .696'
        ' 10.75 10.75 0 0 1-19.876 0"/>'
        '<circle cx="12" cy="12" r="3"/>'
    ),
    "maximize": (
        '<path d="M8 3H5a2 2 0 0 0-2 2v3"/>'
        '<path d="M21 8V5a2 2 0 0 0-2-2h-3"/>'
        '<path d="M3 16v3a2 2 0 0 0 2 2h3"/>'
        '<path d="M16 21h3a2 2 0 0 0 2-2v-3"/>'
    ),
    "minimize": (
        '<path d="M8 3v3a2 2 0 0 1-2 2H3"/>'
        '<path d="M21 8h-3a2 2 0 0 1-2-2V3"/>'
        '<path d="M3 16h3a2 2 0 0 1 2 2v3"/>'
        '<path d="M16 21v-3a2 2 0 0 1 2-2h3"/>'
    ),
    "layout-grid": (
        '<rect width="7" height="7" x="3" y="3" rx="1"/>'
        '<rect width="7" height="7" x="14" y="3" rx="1"/>'
        '<rect width="7" height="7" x="14" y="14" rx="1"/>'
        '<rect width="7" height="7" x="3" y="14" rx="1"/>'
    ),
    "list": (
        '<path d="M3 5h.01"/><path d="M3 12h.01"/><path d="M3 19h.01"/>'
        '<path d="M8 5h13"/><path d="M8 12h13"/><path d="M8 19h13"/>'
    ),

    # Security
    "lock": (
        '<rect width="18" height="11" x="3" y="11" rx="2" ry="2"/>'
        '<path d="M7 11V7a5 5 0 0 1 10 0v4"/>'
    ),
    "unlock": (
        '<rect width="18" height="11" x="3" y="11" rx="2" ry="2"/>'
        '<path d="M7 11V7a5 5 0 0 1 9.9-1"/>'
    ),
    "shield": (
        '<path d="M20 13c0 5-3.5 7.5-7.66 8.95'
        'a1 1 0 0 1-.67-.01C7.5 20.5 4 18 4 13'
        'V6a1 1 0 0 1 1-1c2 0 4.5-1.2 6.24-2.72'
        'a1.17 1.17 0 0 1 1.52 0C14.51 3.81 17 5 19 5'
        'a1 1 0 0 1 1 1z"/>'
    ),
    "key": (
        '<path d="m15.5 7.5 2.3 2.3a1 1 0 0 0 1.4 0l2.1-2.1a1 1 0 0 0 0-1.4L19 4"/>'
        '<path d="m21 2-9.6 9.6"/>'
        '<circle cx="7.5" cy="15.5" r="5.5"/>'
    ),

    # Marking
    "ban": (
        '<circle cx="12" cy="12" r="10"/>'
        '<path d="M4.929 4.929 19.07 19.071"/>'
    ),
    "scan-line": (
        '<path d="M3 7V5a2 2 0 0 1 2-2h2"/>'
        '<path d="M17 3h2a2 2 0 0 1 2 2v2"/>'
        '<path d="M21 17v2a2 2 0 0 1-2 2h-2"/>'
        '<path d="M7 21H5a2 2 0 0 1-2-2v-2"/>'
        '<path d="M7 12h10"/>'
    ),

    # Media
    "image": (
        '<rect width="18" height="18" x="3" y="3" rx="2" ry="2"/>'
        '<circle cx="9" cy="9" r="2"/>'
        '<path d="m21 15-3.086-3.086a2 2 0 0 0-2.828 0L6 21"/>'
    ),
    "table": (
        '<path d="M12 3v18"/>'
        '<rect width="18" height="18" x="3" y="3" rx="2"/>'
        '<path d="M3 9h18"/><path d="M3 15h18"/>'
    ),

    # Decoration
    "stamp": (
        '<path d="M14 13V8.5C14 7 15 7 15 5a3 3 0 0 0-6 0c0 2 1 2 1 3.5V13"/>'
        '<path d="M20 15.5a2.5 2.5 0 0 0-2.5-2.5h-11'
        'A2.5 2.5 0 0 0 4 15.5V17a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1z"/>'
        '<path d="M5 22h14"/>'
    ),
    "droplets": (
        '<path d="M7 16.3c2.2 0 4-1.83 4-4.05 0-1.16-.57-2.26-1.71-3.19'
        'S7.29 6.75 7 5.3c-.29 1.45-1.14 2.84-2.29 3.76'
        'S3 11.1 3 12.25c0 2.22 1.8 4.05 4 4.05z"/>'
        '<path d="M12.56 6.6A10.97 10.97 0 0 0 14 3.02'
        'c.5 2.5 2 4.9 4 6.5s3 3.5 3 5.5a6.98 6.98 0 0 1-11.91 4.97"/>'
    ),

    # State & Feedback
    "check": '<path d="M20 6 9 17l-5-5"/>',
    "check-circle": (
        '<path d="M21.801 10A10 10 0 1 1 17 3.335"/>'
        '<path d="m9 11 3 3L22 4"/>'
    ),
    "x": '<path d="M18 6 6 18"/><path d="m6 6 12 12"/>',
    "alert-triangle": (
        '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14'
        'A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"/>'
        '<path d="M12 9v4"/><path d="M12 17h.01"/>'
    ),
    "info": (
        '<circle cx="12" cy="12" r="10"/>'
        '<path d="M12 16v-4"/><path d="M12 8h.01"/>'
    ),

    # Workflow
    "layers": (
        '<path d="M12.83 2.18a2 2 0 0 0-1.66 0L2.6 6.08'
        'a1 1 0 0 0 0 1.83l8.58 3.91a2 2 0 0 0 1.66 0'
        'l8.58-3.9a1 1 0 0 0 0-1.83z"/>'
        '<path d="M2 12a1 1 0 0 0 .58.91l8.6 3.91'
        'a2 2 0 0 0 1.65 0l8.58-3.9A1 1 0 0 0 22 12"/>'
        '<path d="M2 17a1 1 0 0 0 .58.91l8.6 3.91'
        'a2 2 0 0 0 1.65 0l8.58-3.9A1 1 0 0 0 22 17"/>'
    ),
    "git-branch": (
        '<path d="M15 6a9 9 0 0 0-9 9V3"/>'
        '<circle cx="18" cy="6" r="3"/>'
        '<circle cx="6" cy="18" r="3"/>'
    ),
    "settings": (
        '<path d="M9.671 4.136a2.34 2.34 0 0 1 4.659 0'
        ' 2.34 2.34 0 0 0 3.319 1.915'
        ' 2.34 2.34 0 0 1 2.33 4.033'
        ' 2.34 2.34 0 0 0 0 3.831'
        ' 2.34 2.34 0 0 1-2.33 4.033'
        ' 2.34 2.34 0 0 0-3.319 1.915'
        ' 2.34 2.34 0 0 1-4.659 0'
        ' 2.34 2.34 0 0 0-3.32-1.915'
        ' 2.34 2.34 0 0 1-2.33-4.033'
        ' 2.34 2.34 0 0 0 0-3.831'
        'A2.34 2.34 0 0 1 6.35 6.051'
        ' 2.34 2.34 0 0 0 9.671 4.136"/>'
        '<circle cx="12" cy="12" r="3"/>'
    ),
}


def is_svg_icon(name: str) -> bool:
    """Return True if *name* is a known Lucide icon."""
    return name in _SVGS


def svg_pixmap(name: str, color: str = "#374151", size: int = 20) -> QPixmap:
    """Render a Lucide icon to a QPixmap of *size* × *size* pixels (cached)."""
    key = (name, color, size)
    cached = _PIXMAP_CACHE.get(key)
    if cached is not None:
        return cached
    inner = _SVGS.get(name, "")
    svg_bytes = (
        f'<svg xmlns="http://www.w3.org/2000/svg"'
        f' width="{size}" height="{size}"'
        f' viewBox="0 0 24 24"'
        f' fill="none"'
        f' stroke="{color}"'
        f' stroke-width="2"'
        f' stroke-linecap="round"'
        f' stroke-linejoin="round">'
        f'{inner}'
        f'</svg>'
    ).encode()
    renderer = QSvgRenderer(QByteArray(svg_bytes))
    px = QPixmap(size, size)
    px.fill(QColor(0, 0, 0, 0))
    painter = QPainter(px)
    renderer.render(painter)
    painter.end()
    _PIXMAP_CACHE[key] = px
    return px


def svg_icon(name: str, color: str = "#374151", size: int = 20) -> QIcon:
    """Return a QIcon backed by a Lucide SVG rendered at *size* × *size*."""
    return QIcon(svg_pixmap(name, color, size))
