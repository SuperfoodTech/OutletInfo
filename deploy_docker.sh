#!/bin/bash
# ==============================================================================
# deploy_docker.sh
# Skrip Otomasi Deployment & Migrasi Outlet Info ke Docker & Docker Compose
# ==============================================================================

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "======================================================================"
echo "   DEPLOYMENT OUTLET INFO VIA DOCKER & DOCKER COMPOSE"
echo "======================================================================"
echo "[*] Direktori Aplikasi : ${SCRIPT_DIR}"
echo ""

# 1. Pengecekan Docker dan Docker Compose
if ! command -v docker &> /dev/null; then
    echo "[ERROR] 'docker' tidak ditemukan di sistem."
    echo "        Silakan pasang Docker terlebih dahulu: https://docs.docker.com/engine/install/"
    exit 1
fi

if ! docker compose version &> /dev/null; then
    echo "[ERROR] 'docker compose' (plugin v2) tidak ditemukan."
    echo "        Silakan pasang Docker Compose v2."
    exit 1
fi

echo "[✓] Docker & Docker Compose terdeteksi:"
docker --version
docker compose version

# 2. Pengecekan file .env
if [ ! -f ".env" ]; then
    echo "[ERROR] File .env tidak ditemukan di ${SCRIPT_DIR}."
    echo "        Harap siapkan file .env sebelum menjalankan deployment."
    exit 1
fi

# 3. Buat direktori data & volume jika belum ada
echo ""
echo "[*] Menyiapkan direktori penyimpanan volume..."
mkdir -p output_owners \
         SHOPEE/data \
         GRAB/sessions \
         GOFOOD/session \
         cronjobsot/reports \
         cache

# Setel permission agar user non-root (UID 1000) di container memiliki akses tulis
chmod -R 775 output_owners SHOPEE/data GRAB/sessions GOFOOD/session cronjobsot/reports cache 2>/dev/null || true
if [ "$EUID" -eq 0 ]; then
    chown -R 1000:1000 output_owners SHOPEE/data GRAB/sessions GOFOOD/session cronjobsot/reports cache 2>/dev/null || true
elif command -v sudo &> /dev/null && sudo -n true 2>/dev/null; then
    sudo chown -R 1000:1000 output_owners SHOPEE/data GRAB/sessions GOFOOD/session cronjobsot/reports cache 2>/dev/null || true
fi

# 4. Hentikan service systemd lama & proses bot host yang mungkin masih berjalan
echo ""
echo "[*] Memeriksa dan menghentikan proses bot lama di host..."
if [ "$EUID" -eq 0 ]; then
    systemctl stop outlet-info 2>/dev/null || true
    systemctl disable outlet-info 2>/dev/null || true
    pkill -9 -f "discord_bot.py" 2>/dev/null || true
else
    sudo systemctl stop outlet-info 2>/dev/null || true
    sudo systemctl disable outlet-info 2>/dev/null || true
    sudo pkill -9 -f "discord_bot.py" 2>/dev/null || true
    pkill -9 -f "discord_bot.py" 2>/dev/null || true
fi
echo "[✓] Lingkungan host bersih dari instance bot lama."

# 5. Build image Docker
echo ""
echo "[*] Membangun image Docker (ini membutuhkan beberapa menit pada build pertama)..."
docker compose build

# 6. Jalankan service Discord Bot
echo ""
echo "[*] Menjalankan container bot di background..."
docker compose up -d bot

echo ""
echo "======================================================================"
echo "   MIGRASI KE DOCKER SELESAI & BERHASIL!"
echo "======================================================================"
echo "Perintah berguna untuk pemantauan dan pengelolaan:"
echo ""
echo "  1. Melihat log realtime bot:"
echo "     docker compose logs -f bot"
echo ""
echo "  2. Memeriksa status container:"
echo "     docker compose ps"
echo ""
echo "  3. Menjalankan penarikan manual (CLI interaktif):"
echo "     docker compose run --rm cli"
echo ""
echo "  4. Menjalankan Weekly Batch Sync secara manual:"
echo "     docker compose run --rm weekly-sync"
echo ""
echo "  5. Merestart container bot:"
echo "     docker compose restart bot"
echo "======================================================================"
