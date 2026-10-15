#!/bin/bash
set -euo pipefail
cd "$(dirname "$0")"

if ! command -v brew >/dev/null 2>&1; then
  echo "Homebrew is required. Install it from https://brew.sh, then run this setup again."
  read -r -p "Press Return to close..."
  exit 1
fi

missing=()
command -v pdftoppm >/dev/null 2>&1 || missing+=(poppler)
command -v pdfinfo >/dev/null 2>&1 || missing+=(poppler)
command -v tesseract >/dev/null 2>&1 || missing+=(tesseract tesseract-lang)
if ((${#missing[@]})); then
  unique_missing=($(printf '%s\n' "${missing[@]}" | sort -u))
  echo "Installing required Homebrew packages: ${unique_missing[*]}"
  brew install "${unique_missing[@]}"
fi

if ! tesseract --list-langs 2>/dev/null | grep -qx hin; then
  echo "Installing Tesseract language data (includes Hindi)."
  brew install tesseract-lang
fi

if [[ ! -x .venv/bin/python ]]; then
  python3 -m venv .venv
fi
.venv/bin/python -m pip install --upgrade pip
.venv/bin/python -m pip install -r requirements.txt

echo
echo "Setup complete. Double-click run_local.command to start the app."
read -r -p "Press Return to close..."
