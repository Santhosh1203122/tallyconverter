#!/usr/bin/env bash
# ============================================================
# Build TallySalesReport.app and TallySalesReport.dmg on macOS
# ============================================================
# Requirements:
#   - macOS 10.15+
#   - Python 3.9+ (install via python.org or `brew install python`)
#   - Internet access for the first run
#
# Usage:
#   1. Open Terminal in this folder
#   2. chmod +x build_macos.sh
#   3. ./build_macos.sh
#   4. The .dmg will appear in dist/TallySalesReport.dmg
# ============================================================
set -euo pipefail

cd "$(dirname "$0")"

echo "=== [1/5] Checking Python ==="
PY=python3
if ! command -v $PY >/dev/null 2>&1; then
    echo "python3 not found. Install from https://python.org or run: brew install python"
    exit 1
fi
$PY --version

echo
echo "=== [2/5] Creating virtual environment ==="
if [ ! -d .venv ]; then
    $PY -m venv .venv
fi
# shellcheck disable=SC1091
source .venv/bin/activate

echo
echo "=== [3/5] Installing dependencies ==="
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo
echo "=== [4/5] Building .app bundle with PyInstaller ==="
pyinstaller \
    --noconfirm \
    --clean \
    --windowed \
    --name TallySalesReport \
    --osx-bundle-identifier com.maya.tallysalesreport \
    --add-data "converter.py:." \
    app.py

APP_PATH="dist/TallySalesReport.app"
if [ ! -d "$APP_PATH" ]; then
    echo "Build failed: $APP_PATH not found."
    exit 1
fi
echo "App bundle built: $APP_PATH"

echo
echo "=== [5/5] Creating DMG ==="
DMG_PATH="dist/TallySalesReport.dmg"
STAGE_DIR="dist/dmg_staging"
rm -rf "$STAGE_DIR"
mkdir -p "$STAGE_DIR"
cp -R "$APP_PATH" "$STAGE_DIR/"
# Symlink to /Applications so users can drag-install
ln -s /Applications "$STAGE_DIR/Applications"

# Use macOS native hdiutil — no extra tooling required
rm -f "$DMG_PATH"
hdiutil create \
    -volname "Tally Sales Report" \
    -srcfolder "$STAGE_DIR" \
    -ov \
    -format UDZO \
    "$DMG_PATH"

rm -rf "$STAGE_DIR"

if [ -f "$DMG_PATH" ]; then
    echo
    echo "Build successful!"
    echo "  App: $APP_PATH"
    echo "  DMG: $DMG_PATH"
    echo
    echo "NOTE: This DMG is unsigned. On first launch macOS Gatekeeper may"
    echo "      block it. Right-click the .app and choose Open, then confirm."
else
    echo "DMG creation failed."
    exit 1
fi
