#!/bin/bash

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="$(dirname "$SCRIPT_DIR")"

echo "==============================="
echo "  GoFood Merchant Scraper"
echo "==============================="
echo ""

# Cari virtual environment (cek GOFOOD dulu, lalu root)
VENV_DIR=""
for candidate in "$SCRIPT_DIR/.venv" "$ROOT_DIR/.venv"; do
    if [ -d "$candidate" ]; then
        VENV_DIR="$candidate"
        break
    fi
done

if [ -n "$VENV_DIR" ]; then
    echo "[*] Mengaktifkan virtual environment: $VENV_DIR"
    source "$VENV_DIR/bin/activate"
elif command -v python3 &>/dev/null; then
    echo "[*] Menggunakan python3 sistem..."
else
    echo "[!] Python3 tidak ditemukan!"
    exit 1
fi

cd "$SCRIPT_DIR"
echo "[*] Working directory: $SCRIPT_DIR"
echo "[*] Menjalankan GoFood scraper..."
echo ""

python3 gofood_scraper.py

echo ""
echo "[✓] Selesai."
