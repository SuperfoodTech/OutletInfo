#!/bin/bash
# ==============================================================================
# setup_systemd.sh
# Script Otomasi Deployment Outlet Info ke systemd (untuk VPS / Server Linux)
# ==============================================================================

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

CURRENT_USER=$(id -un)
CURRENT_GROUP=$(id -gn)
SERVICE_NAME="outlet-info"
SERVICE_FILE="/etc/systemd/system/${SERVICE_NAME}.service"
VENV_PYTHON="${SCRIPT_DIR}/.venv/bin/python"

echo "======================================================================"
echo "   SETUP PRODUCTION SYSTEMD SERVICE: ${SERVICE_NAME}"
echo "======================================================================"
echo "[*] Direktori Aplikasi : ${SCRIPT_DIR}"
echo "[*] User / Group       : ${CURRENT_USER}:${CURRENT_GROUP}"
echo ""

# 1. Cek Virtual Environment
if [ ! -f "$VENV_PYTHON" ]; then
    echo "[!] Virtual environment (.venv) belum ditemukan di ${SCRIPT_DIR}."
    echo "[*] Menjalankan start.sh untuk inisialisasi awal venv & dependency..."
    chmod +x start.sh
    ./start.sh
fi

if [ ! -f "$VENV_PYTHON" ]; then
    echo "[ERROR] Python virtual environment tetap tidak ditemukan di: $VENV_PYTHON"
    exit 1
fi

echo "[✓] Python executable ditemukan: $VENV_PYTHON"

# 2. Cek Swap Memory (Sangat Krusial untuk Server 2GB RAM)
echo ""
echo "[*] Memeriksa Swap Memory..."
SWAP_TOTAL_KB=$(grep SwapTotal /proc/meminfo | awk '{print $2}')
SWAP_TOTAL_MB=$((SWAP_TOTAL_KB / 1024))

echo "[*] Swap terdeteksi: ${SWAP_TOTAL_MB} MB"
if [ "$SWAP_TOTAL_MB" -lt 2000 ]; then
    echo "----------------------------------------------------------------------"
    echo "⚠️  PERINGATAN: Swap memory Anda kurang dari 2GB (saat ini: ${SWAP_TOTAL_MB} MB)."
    echo "   Untuk server dengan RAM fisik 2GB, Chromium membutuhkan Swap minimal 4GB"
    echo "   agar tidak terjadi Out-Of-Memory (OOM) crash."
    echo ""
    echo "   Rekomendasi perintah membuat swap 4GB (Jalankan manual jika belum):"
    echo "     sudo fallocate -l 4G /swapfile"
    echo "     sudo chmod 600 /swapfile"
    echo "     sudo mkswap /swapfile"
    echo "     sudo swapon /swapfile"
    echo "     echo '/swapfile none swap sw 0 0' | sudo tee -a /etc/fstab"
    echo "----------------------------------------------------------------------"
else
    echo "[✓] Swap memory mencukupi (>= 2GB)."
fi

# 3. Buat file systemd service dari template
echo ""
echo "[*] Menyiapkan service unit ${SERVICE_FILE}..."

# Render file konfigurasi sementara
TEMP_SERVICE="/tmp/${SERVICE_NAME}.service"
sed \
    -e "s|{{USER}}|${CURRENT_USER}|g" \
    -e "s|{{GROUP}}|${CURRENT_GROUP}|g" \
    -e "s|{{APP_DIR}}|${SCRIPT_DIR}|g" \
    -e "s|{{PYTHON_PATH}}|${VENV_PYTHON}|g" \
    "${SCRIPT_DIR}/outlet-info.service.template" > "$TEMP_SERVICE"

# Salin ke /etc/systemd/system/ (memerlukan sudo jika bukan root)
if [ "$EUID" -eq 0 ]; then
    mv "$TEMP_SERVICE" "$SERVICE_FILE"
    systemctl daemon-reload
    echo "[✓] Service file berhasil dipasang di ${SERVICE_FILE}"
else
    echo "[*] Membutuhkan hak akses sudo untuk memasang service ke ${SERVICE_FILE}..."
    sudo mv "$TEMP_SERVICE" "$SERVICE_FILE"
    sudo systemctl daemon-reload
    echo "[✓] Service file berhasil dipasang di ${SERVICE_FILE}"
fi

echo ""
echo "======================================================================"
echo "   INSTALASI SYSTEMD SELESAI!"
echo "======================================================================"
echo "Gunakan perintah berikut untuk mengelola service:"
echo ""
echo "  1. Menjalankan & mengaktifkan auto-start saat reboot:"
echo "     sudo systemctl enable --now ${SERVICE_NAME}"
echo ""
echo "  2. Memeriksa status:"
echo "     sudo systemctl status ${SERVICE_NAME}"
echo ""
echo "  3. Melihat log realtime:"
echo "     journalctl -u ${SERVICE_NAME} -f"
echo ""
echo "  4. Menghentikan atau restart service:"
echo "     sudo systemctl restart ${SERVICE_NAME}"
echo "     sudo systemctl stop ${SERVICE_NAME}"
echo "======================================================================"

