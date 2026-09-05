"""
open_dashboard_7307.py
======================
Membuka browser Shopee Partner Portal untuk akun auto7307
menggunakan Chrome profile yang sudah tersimpan (tidak perlu login ulang).

Usage (dari folder task-weekly/):
    uv run src/shopee-omzet-automation/open_dashboard_7307.py

Browser akan tetap terbuka sampai kamu menutupnya secara manual (Ctrl+C / tutup jendela).
"""

import os
import sys
import time
from pathlib import Path

# ── Path Setup ─────────────────────────────────────────────────────────────────
SCRIPT_DIR   = Path(__file__).resolve().parent          # shopee-omzet-automation/
PROJECT_ROOT = SCRIPT_DIR.parent.parent                 # task-weekly/
sys.path.insert(0, str(SCRIPT_DIR))

# ── Config ─────────────────────────────────────────────────────────────────────
ACCOUNT_NAME = "auto7307"
HEADLESS     = False   # Selalu tampilkan browser (False = GUI)
DASHBOARD_URL = "https://partner.shopee.co.id/food/dashboard"

DATA_DIR     = SCRIPT_DIR / "data"
SESSION_FILE = DATA_DIR / f"session_{ACCOUNT_NAME}.json"


def main():
    from core import browser

    print(f"🚀 Membuka dashboard Shopee untuk akun: {ACCOUNT_NAME}")
    print(f"   Session file : {SESSION_FILE}")
    print(f"   Chrome profile: {DATA_DIR / f'chrome_profile_{ACCOUNT_NAME}'}")
    print()

    if not SESSION_FILE.exists():
        print(f"❌ File sesi tidak ditemukan: {SESSION_FILE}")
        print("   Jalankan warmer.py terlebih dahulu untuk membuat sesi.")
        sys.exit(1)

    # Arahkan browser module ke session file akun ini
    browser.set_session_file(SESSION_FILE)

    print("🌐 Memulai browser...")
    session_data = browser.get_session(
        headless=HEADLESS,
        close_browser=False,   # Biarkan browser tetap terbuka
        interactive=True,
    )

    if not session_data:
        print("❌ Gagal membuka sesi. Sesi mungkin kedaluwarsa.")
        print("   Jalankan warmer.py untuk memperbaharui sesi.")
        sys.exit(1)

    driver = session_data.get("driver")
    if not driver:
        print("❌ Driver tidak tersedia.")
        sys.exit(1)

    print(f"✅ Browser terbuka! URL saat ini: {driver.current_url}")
    print()

    # Navigasi ke dashboard jika belum di sana
    if "dashboard" not in driver.current_url.lower():
        print(f"🔄 Navigasi ke dashboard: {DASHBOARD_URL}")
        driver.get(DASHBOARD_URL)
        time.sleep(3)
        print(f"   URL sekarang: {driver.current_url}")

    print()
    print("=" * 55)
    print("  Browser sudah terbuka. Tekan Ctrl+C untuk menutup.")
    print("=" * 55)

    try:
        # Jaga proses tetap hidup selama browser terbuka
        while True:
            time.sleep(2)
            # Cek apakah browser masih hidup
            try:
                _ = driver.current_url
            except Exception:
                print("\n🔴 Browser ditutup.")
                break
    except KeyboardInterrupt:
        print("\n🛑 Dihentikan oleh pengguna.")
    finally:
        try:
            driver.quit()
        except Exception:
            pass
        print("✅ Selesai.")


if __name__ == "__main__":
    main()
