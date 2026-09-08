#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="$(cd "$SCRIPT_DIR/../.." && pwd)"

echo "==> Building pptx-a11y macOS application bundle..."
cd "$ROOT_DIR"

.venv/bin/pyinstaller --clean -y packaging/pptx-a11y.spec

echo "==> Application bundle successfully built at: $ROOT_DIR/dist/pptx-a11y.app"
