#!/bin/bash

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# Pastikan uv terinstall
if ! command -v uv &> /dev/null; then
    echo "[!] 'uv' tidak ditemukan di sistem."
    echo "    Install dengan: curl -LsSf https://astral.sh/uv/install.sh | sh"
    read -p "Tekan Enter untuk keluar..." _
    exit 1
fi

# Buat virtual environment jika belum ada
if [ ! -d ".venv" ]; then
    echo "[*] Membuat virtual environment baru menggunakan uv (.venv)..."
    uv venv .venv
fi

echo "[*] Mengaktifkan virtual environment..."
source .venv/bin/activate

echo "[*] Memeriksa dan menginstall dependencies (pandas, openpyxl, playwright, requests, python-dotenv)..."
uv pip install -q pandas openpyxl playwright requests python-dotenv

# Pastikan browser playwright terinstall (tanpa meminta sudo)
uv run playwright install chromium > /dev/null 2>&1

echo "[*] Menjalankan CLI Utama..."
python3 cli.py

EXIT_CODE=$?
if [ $EXIT_CODE -ne 0 ] && [ $EXIT_CODE -ne 130 ]; then
    echo ""
    echo "[!] Program keluar dengan kode: $EXIT_CODE"
    read -p "Tekan Enter untuk keluar..." _
fi
