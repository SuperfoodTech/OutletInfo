#!/bin/bash

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Cari virtual environment (cek GRAB dulu, lalu SHOPEE, lalu root)
VENV_DIR=""
for candidate in "$SCRIPT_DIR/.venv" "$SCRIPT_DIR/GRAB/.venv" "$SCRIPT_DIR/SHOPEE/.venv"; do
    if [ -d "$candidate" ]; then
        VENV_DIR="$candidate"
        break
    fi
done

# Aktifkan venv atau gunakan python sistem
if [ -n "$VENV_DIR" ]; then
    echo "[*] Mengaktifkan virtual environment: $VENV_DIR"
    source "$VENV_DIR/bin/activate"
elif command -v python3 &>/dev/null; then
    echo "[*] Menggunakan python3 sistem..."
else
    echo "[!] Python3 tidak ditemukan!"
    read -p "Tekan Enter untuk keluar..." _
    exit 1
fi

cd "$SCRIPT_DIR"

python3 cli.py

EXIT_CODE=$?
if [ $EXIT_CODE -ne 0 ] && [ $EXIT_CODE -ne 130 ]; then
    echo ""
    echo "[!] Program keluar dengan kode: $EXIT_CODE"
    read -p "Tekan Enter untuk keluar..." _
fi
