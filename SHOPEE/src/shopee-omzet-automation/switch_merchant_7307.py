"""
switch_merchant_7307.py
=======================
Automates launching auto7307 profile, switching merchant to 'SATE SRIWIJAYA',
refreshing the session tokens, and saving it to session_auto7307.json.
"""

import os
import sys
import time
from pathlib import Path

# Setup paths
SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent.parent
sys.path.insert(0, str(SCRIPT_DIR))

def main():
    from core import browser
    from selenium.webdriver.common.by import By

    ACCOUNT_NAME = "auto7307"
    SESSION_FILE = SCRIPT_DIR / "data" / f"session_{ACCOUNT_NAME}.json"
    
    print(f"🚀 [INIT] Target Account: {ACCOUNT_NAME}")
    print(f"💾 [INIT] Session File : {SESSION_FILE}")
    
    # Point browser module to the session file
    browser.set_session_file(SESSION_FILE)
    
    print("🌐 [BROWSER] Memulai headful browser...")
    session_data = browser.get_session(
        headless=False,
        close_browser=False,
        interactive=True
    )
    
    if not session_data:
        print("❌ [ERROR] Gagal mendapatkan session browser. Sesi mungkin kedaluwarsa.")
        sys.exit(1)
        
    driver = session_data.get("driver")
    if not driver:
        print("❌ [ERROR] Driver tidak tersedia.")
        sys.exit(1)
        
    print(f"✅ [BROWSER] Terbuka! URL: {driver.current_url}")
    
    # Navigate to dashboard if not already there
    if "dashboard" not in driver.current_url.lower():
        print("🔄 [NAVIGATE] Membuka dashboard Shopee...")
        driver.get("https://partner.shopee.co.id/food/dashboard")
        time.sleep(5)
        
    # Get current merchant name
    try:
        current_name = driver.find_element(By.CSS_SELECTOR, ".merchantName").text.strip()
        print(f"📍 [MERCHANT] Saat ini: {current_name}")
    except:
        current_name = "Tidak terdeteksi"
        print("📍 [MERCHANT] Nama merchant saat ini tidak terdeteksi di UI.")

    target_merchant = "SATE SRIWIJAYA"
    print(f"🔄 [MERCHANT] Mencoba berpindah ke: {target_merchant}...")
    
    # Attempt switch
    success = browser.auto_switch_merchant(driver, target_merchant)
    
    if success:
        print(f"🎉 [SUCCESS] Berhasil berpindah ke {target_merchant}!")
        print("🔄 [TOKEN] Merefresh session token...")
        tokens = browser.refresh_tokens(driver)
        print(f"✅ [TOKEN] Sesi baru berhasil disimpan dengan token: {tokens.get('shopee_tob_token', '')[:20]}...")
    else:
        print(f"❌ [FAILED] Gagal berpindah ke {target_merchant}.")
        # Capture screenshot for debugging
        screenshot_path = SCRIPT_DIR / f"error_{ACCOUNT_NAME}.png"
        driver.save_screenshot(str(screenshot_path))
        print(f"📸 [DEBUG] Screenshot disimpan di: {screenshot_path}")
        
    print("⏳ [WAIT] Menunggu 15 detik sebelum menutup browser...")
    time.sleep(15)
    
    try:
        driver.quit()
        print("✅ [EXIT] Browser ditutup.")
    except Exception as e:
        print(f"⚠️ [EXIT] Gagal menutup browser: {e}")

if __name__ == "__main__":
    main()
