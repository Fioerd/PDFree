# -*- mode: python ; coding: utf-8 -*-
# PDFree PyInstaller spec file — macOS
# Build with: pyinstaller PDFree-mac.spec
# Produces:   dist/PDFree.app  (and optionally a .dmg via build-mac.sh)

import os

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('LOGO.svg', '.'),
    ],
    hiddenimports=[
        'pdfplumber',
        'pdfminer',
        'pdfminer.high_level',
        'pdfminer.layout',
        'pdfminer.converter',
        'pdfminer.pdfpage',
        'pypdf',
        'fitz',
        'PIL._tkinter_finder',
        'PySide6.QtSvg',
        'PySide6.QtSvgWidgets',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'numpy',
        'scipy',
        'pandas',
        'IPython',
        'jupyter',
    ],
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PDFree',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,          # UPX is unreliable on macOS
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # Use .icns if available, otherwise omit
    icon='LOGO.icns' if os.path.exists('LOGO.icns') else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='PDFree',
)

app = BUNDLE(
    coll,
    name='PDFree.app',
    icon='LOGO.icns' if os.path.exists('LOGO.icns') else None,
    bundle_identifier='com.fioerd.pdfree',
    info_plist={
        'CFBundleName': 'PDFree',
        'CFBundleDisplayName': 'PDFree',
        'CFBundleVersion': '1.0.0',
        'CFBundleShortVersionString': '1.0.0',
        'NSHighResolutionCapable': True,
        'NSRequiresAquaSystemAppearance': False,   # supports dark mode
        'CFBundleDocumentTypes': [
            {
                'CFBundleTypeName': 'PDF Document',
                'CFBundleTypeRole': 'Viewer',
                'LSItemContentTypes': ['com.adobe.pdf'],
                'CFBundleTypeExtensions': ['pdf'],
            }
        ],
        'LSMinimumSystemVersion': '11.0',          # macOS Big Sur+
    },
)
