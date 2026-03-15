#!/bin/bash
# PDFree macOS build script
# Run this on a Mac with the venv active:
#   source .venv/bin/activate
#   bash build-mac.sh
#
# Outputs:
#   dist/PDFree.app      — the app bundle
#   dist/PDFree.dmg      — drag-to-install disk image (if create-dmg is installed)

set -e

echo "==> Converting SVG icon to .icns (requires librsvg + iconutil)..."
if command -v rsvg-convert &>/dev/null; then
    mkdir -p PDFree.iconset
    for size in 16 32 128 256 512; do
        rsvg-convert -w $size -h $size LOGO.svg -o PDFree.iconset/icon_${size}x${size}.png
        rsvg-convert -w $((size*2)) -h $((size*2)) LOGO.svg -o PDFree.iconset/icon_${size}x${size}@2x.png
    done
    iconutil -c icns PDFree.iconset -o LOGO.icns
    rm -rf PDFree.iconset
    echo "    LOGO.icns created."
else
    echo "    rsvg-convert not found — skipping icon conversion (install with: brew install librsvg)"
    echo "    The app will launch without a custom icon."
fi

echo "==> Installing / upgrading PyInstaller..."
pip install --upgrade pyinstaller

echo "==> Building PDFree.app..."
pyinstaller PDFree-mac.spec --noconfirm

echo "==> Build complete: dist/PDFree.app"

# --- Optional: wrap in a .dmg ---
if command -v create-dmg &>/dev/null; then
    echo "==> Creating PDFree.dmg..."
    rm -f dist/PDFree.dmg
    create-dmg \
        --volname "PDFree" \
        --volicon "LOGO.icns" \
        --window-pos 200 120 \
        --window-size 600 400 \
        --icon-size 100 \
        --icon "PDFree.app" 150 185 \
        --hide-extension "PDFree.app" \
        --app-drop-link 450 185 \
        "dist/PDFree.dmg" \
        "dist/PDFree.app"
    echo "==> dist/PDFree.dmg ready."
else
    echo ""
    echo "Tip: install create-dmg to also produce a .dmg installer:"
    echo "     brew install create-dmg"
    echo "Then re-run this script."
fi

echo ""
echo "Done! You can now:"
echo "  • Drag dist/PDFree.app to /Applications"
echo "  • Or distribute dist/PDFree.dmg to users"
