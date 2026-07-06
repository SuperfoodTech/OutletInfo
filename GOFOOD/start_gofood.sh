#!/bin/bash

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="$(dirname "$SCRIPT_DIR")"

echo "==============================="
echo "  GoFood Merchant Scraper"
echo "==============================="
echo ""

# Gunakan venv dari root
VENV_DIR="$ROOT_DIR/.venv"

if [ ! -d "$VENV_DIR" ]; then
    echo "[!] Virtual environment (.venv) belum dibuat di root."
    echo "[*] Silakan jalankan ./start.sh di folder utama terlebih dahulu."
    exit 1
fi

echo "[*] Mengaktifkan virtual environment..."
source "$VENV_DIR/bin/activate"

cd "$SCRIPT_DIR"
echo "[*] Working directory: $SCRIPT_DIR"
echo "[*] Menjalankan GoFood scraper..."
echo ""

python3 gofood_scraper.py

echo ""
echo "[✓] Selesai."
