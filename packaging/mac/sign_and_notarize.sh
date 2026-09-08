#!/usr/bin/env bash
set -euo pipefail

# pptx-a11y: macOS Code Signing & Apple Notary Service Submission Script
# Requires environment variables:
#   DEVELOPER_ID_APPLICATION : "Developer ID Application: Your Name (TEAM_ID)"
#   APPLE_ID                 : "developer@example.com"
#   APPLE_TEAM_ID            : "TEAM_ID"
#   APPLE_APP_SPECIFIC_PASSWORD : "xxxx-xxxx-xxxx-xxxx"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="$(cd "$SCRIPT_DIR/../.." && pwd)"
APP_PATH="$ROOT_DIR/dist/pptx-a11y.app"
DMG_PATH="$ROOT_DIR/dist/pptx-a11y-macOS-arm64.dmg"
ENTITLEMENTS="$ROOT_DIR/packaging/mac/entitlements.plist"

if [ ! -d "$APP_PATH" ]; then
    echo "[ERROR] Application bundle not found at $APP_PATH. Run packaging/mac/build_mac.sh first."
    exit 1
fi

if [ -z "${DEVELOPER_ID_APPLICATION:-}" ]; then
    echo "[WARNING] DEVELOPER_ID_APPLICATION environment variable not set."
    echo "Performing ad-hoc signature for local testing..."
    codesign --deep --force --options runtime --entitlements "$ENTITLEMENTS" --sign - "$APP_PATH"
    echo "==> Ad-hoc signed successfully: $APP_PATH"
    exit 0
fi

echo "==> Signing application bundle with hardened runtime..."
codesign --deep --force --verify --verbose \
    --options runtime \
    --entitlements "$ENTITLEMENTS" \
    --sign "$DEVELOPER_ID_APPLICATION" \
    "$APP_PATH"

echo "==> Verifying signature..."
codesign --verify --deep --strict --verbose=2 "$APP_PATH"

echo "==> Packaging into DMG volume..."
rm -f "$DMG_PATH"
hdiutil create -volname "pptx-a11y" -srcfolder "$APP_PATH" -ov -format UDZO "$DMG_PATH"

echo "==> Signing DMG..."
codesign --sign "$DEVELOPER_ID_APPLICATION" "$DMG_PATH"

if [ -n "${APPLE_ID:-}" ] && [ -n "${APPLE_TEAM_ID:-}" ] && [ -n "${APPLE_APP_SPECIFIC_PASSWORD:-}" ]; then
    echo "==> Submitting DMG to Apple Notary Service..."
    xcrun notarytool submit "$DMG_PATH" \
        --apple-id "$APPLE_ID" \
        --team-id "$APPLE_TEAM_ID" \
        --password "$APPLE_APP_SPECIFIC_PASSWORD" \
        --wait

    echo "==> Stapling notarization ticket to DMG..."
    xcrun stapler staple "$DMG_PATH"
    echo "==> Notarization and stapling complete: $DMG_PATH"
else
    echo "[NOTICE] Skipping Apple Notarization (APPLE_ID or APPLE_APP_SPECIFIC_PASSWORD not configured)."
fi

echo "==> Release artifact ready for distribution: $DMG_PATH"
