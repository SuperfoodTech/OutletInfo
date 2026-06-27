#!/bin/bash

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="$SCRIPT_DIR/.venv"

echo "==============================="
echo "  Grab Merchant Scraper"
echo "==============================="
echo ""

# Cek virtual environment
if [ -d "$VENV_DIR" ]; then
    echo "[*] Mengaktifkan virtual environment..."
    source "$VENV_DIR/bin/activate"
elif command -v python3 &>/dev/null; then
    echo "[*] Menggunakan python3 sistem..."
else
    echo "[!] Python3 tidak ditemukan!"
    exit 1
fi

cd "$SCRIPT_DIR"
echo "[*] Working directory: $SCRIPT_DIR"
echo "[*] Menjalankan scraper..."
echo ""

python3 grab_merchant_scraper.py

echo ""
echo "[✓] Selesai."
