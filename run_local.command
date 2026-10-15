#!/bin/bash
set -euo pipefail
cd "$(dirname "$0")"

if [[ ! -x .venv/bin/streamlit ]]; then
  echo "Run setup_mac.command first."
  read -r -p "Press Return to close..."
  exit 1
fi
for tool in pdftoppm pdfinfo tesseract; do
  if ! command -v "$tool" >/dev/null 2>&1; then
    echo "Missing system dependency: $tool. Run setup_mac.command."
    read -r -p "Press Return to close..."
    exit 1
  fi
done
if ! tesseract --list-langs 2>/dev/null | grep -qx hin; then
  echo "Tesseract Hindi language data is missing. Run setup_mac.command."
  read -r -p "Press Return to close..."
  exit 1
fi
exec .venv/bin/streamlit run app.py --server.address=127.0.0.1 --server.port=8501
