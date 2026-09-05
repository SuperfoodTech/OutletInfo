"""
Shopee Partner Dashboard Opener (Idle Mode)
===========================================
Membuka browser dan login ke Dashboard Shopee Partner,
kemudian membiarkan browser tetap aktif (idle) tanpa timeout.

Usage:
    python open_dashboard.py
"""

import os
import json
import sys
import time
from pathlib import Path

# Auto-detect and switch to .venv python if not active
SCRIPT_DIR = Path(__file__).resolve().parent
for venv_candidate in [SCRIPT_DIR / ".venv" / "bin" / "python", SCRIPT_DIR.parent / ".venv" / "bin" / "python", SCRIPT_DIR.parent.parent / ".venv" / "bin" / "python"]:
    if venv_candidate.exists() and sys.executable != str(venv_candidate):
        os.execv(str(venv_candidate), [str(venv_candidate)] + sys.argv)
        break

from selenium.webdriver.chrome.options import Options

# ──────────────────────────────────────────────────────────────
# Path Setup
# ──────────────────────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).resolve().parent
AUTOMATION_DIR = SCRIPT_DIR / "src" / "shopee-omzet-automation"
CHROME_PROFILE_DIR = SCRIPT_DIR / "data" / "chrome_profile"

# Add automation dir to sys.path for browser module import
if str(AUTOMATION_DIR) not in sys.path:
    sys.path.insert(0, str(AUTOMATION_DIR))

from core import browser

# FORCE the profile directory (same pattern as pull_outlet_info.py)
orig_add_argument = Options.add_argument
def custom_add_argument(self, argument):
    if "--user-data-dir=" in argument:
        argument = f"--user-data-dir={CHROME_PROFILE_DIR}"
        print(f"🔧 [PATCH] Mengalihkan user data dir ke: {argument}")
    orig_add_argument(self, argument)
Options.add_argument = custom_add_argument

SESSION_FILE = SCRIPT_DIR / "data" / "session.json"
CREDS_FILE = SCRIPT_DIR / "credentials.json"

DEFAULT_USERNAME = "allvbadmin"
DEFAULT_PASSWORD = "Master!00!"


def main():
    print("=" * 70)
    print("  SHOPEE PARTNER DASHBOARD - IDLE RUNNER")
    print("=" * 70)
    print()

    # Load credentials
    username = DEFAULT_USERNAME
    password = DEFAULT_PASSWORD
    if CREDS_FILE.exists():
        try:
            creds = json.loads(CREDS_FILE.read_text())
            username = creds.get("username") or creds.get("shopee_username") or username
            password = creds.get("password") or creds.get("shopee_password") or password
            print(f"  ✓ Kredensial dibaca dari {CREDS_FILE.name} (User: {username})")
        except Exception as e:
            print(f"  [!] Warning reading credentials.json: {e}")

    browser.set_session_file(SESSION_FILE)

    print("\n[*] Membuka browser (headless=False) dan login ke Shopee Dashboard...")
    session_data = browser.get_session(
        username=username,
        password=password,
        headless=False,
        close_browser=False,
        interactive=False,
    )

    if not session_data or "driver" not in session_data:
        print("\n❌ Gagal membuka browser atau login ke Shopee Dashboard.")
        return

    driver = session_data["driver"]

    print()
    print("=" * 70)
    print("  ✅ DASHBOARD SHOPEE PARTNER BERHASIL DIBUKA!")
    print("  Browser dalam kondisi IDLE (aktif tanpa timeout).")
    print("  Tekan [Ctrl + C] di terminal ini untuk menutup browser dan keluar.")
    print("=" * 70)
    print()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n[*] Menutup browser dan keluar...")
        try:
            driver.quit()
        except Exception:
            pass
        print("[✓] Browser berhasil ditutup. Selesai.")


if __name__ == "__main__":
    main()
