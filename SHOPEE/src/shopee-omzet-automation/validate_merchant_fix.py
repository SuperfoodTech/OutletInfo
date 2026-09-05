"""
validate_merchant_fix.py
========================
Script validasi untuk memastikan semua fix merchant switching berfungsi:

TEST 1: Dashboard Dropdown Switch (PHASE 2)
  - Buka dashboard → klik profil → klik "Pilih Merchant Lain"
  - Verifikasi dropdown terbuka (tidak ada false negative)
  - Pilih SATE SRIWIJAYA dari dropdown

TEST 2: Deliberate Logout + Onboarding (PHASE 1)
  - Logout dari dashboard
  - Login ulang (credential login)
  - Landing di onboarding page
  - Deteksi elemen .listItem ada
  - Klik merchant pertama via js_selector_click yang baru
  - Verifikasi redirect ke dashboard

Usage:
    uv run --project src src/shopee-omzet-automation/validate_merchant_fix.py
"""

import sys
import time
import json
from pathlib import Path

SCRIPT_DIR   = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent.parent
sys.path.insert(0, str(SCRIPT_DIR))

ACCOUNT_NAME  = "auto7307"
TARGET_NAME   = "SATE SRIWIJAYA"
DATA_DIR      = SCRIPT_DIR / "data"
SESSION_FILE  = DATA_DIR / f"session_{ACCOUNT_NAME}.json"
DASHBOARD_URL = "https://partner.shopee.co.id/food/dashboard"

from core import browser
from core.logger import get_logger
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

log = get_logger("validate")

# ── Utilities ──────────────────────────────────────────────────────────────────

class TestResult:
    def __init__(self, name):
        self.name = name
        self.steps = []
        self.passed = True
    
    def step(self, desc, ok, detail=""):
        icon = "✅" if ok else "❌"
        self.steps.append((desc, ok, detail))
        if not ok:
            self.passed = False
        print(f"  {icon} {desc}" + (f" — {detail}" if detail else ""))
    
    def summary(self):
        icon = "✅ PASSED" if self.passed else "❌ FAILED"
        print(f"\n{'='*60}")
        print(f"  {self.name}: {icon}")
        print(f"{'='*60}")
        for desc, ok, detail in self.steps:
            status = "✅" if ok else "❌"
            print(f"    {status} {desc}")
        print()


def dump_onboarding_elements(driver):
    """Dump elemen .listItem dari halaman onboarding untuk debugging."""
    items = driver.execute_script("""
        var results = [];
        var listItems = document.querySelectorAll('.listItem, .merchant-item');
        for (var i = 0; i < listItems.length; i++) {
            var el = listItems[i];
            var rect = el.getBoundingClientRect();
            results.push({
                tag: el.tagName,
                className: el.className.toString().substring(0, 80),
                text: (el.innerText || '').trim().split('\\n')[0].substring(0, 60),
                visible: rect.width > 0 && rect.height > 0,
                cursor: window.getComputedStyle(el).cursor
            });
        }
        return results;
    """)
    if items:
        print(f"\n  📋 Ditemukan {len(items)} elemen .listItem:")
        for i, item in enumerate(items):
            print(f"     [{i}] text='{item['text']}' visible={item['visible']} cursor={item['cursor']}")
    else:
        print("\n  ⚠️  Tidak ada elemen .listItem ditemukan!")
        # Fallback: cek elemen lama untuk debugging
        old_items = driver.execute_script("""
            var results = [];
            var els = document.querySelectorAll('.merchantInfo, .shop-name, .ant-list-item');
            for (var i = 0; i < els.length; i++) {
                results.push({
                    className: els[i].className.toString().substring(0, 80),
                    text: (els[i].innerText || '').trim().split('\\n')[0].substring(0, 60)
                });
            }
            return results;
        """)
        if old_items:
            print(f"  📋 Elemen lama (.merchantInfo) ditemukan: {len(old_items)}")
            for item in old_items:
                print(f"     class='{item['className']}' text='{item['text']}'")
    return items


# ── TEST 1: Dashboard Dropdown Switch ──────────────────────────────────────────

def test_dashboard_dropdown_switch(driver):
    """Test apakah dropdown profil bisa dibuka dan merchant ditemukan di dalamnya."""
    test = TestResult("TEST 1: Dashboard Dropdown Switch (PHASE 2)")
    print(f"\n{'='*60}")
    print(f"  TEST 1: Dashboard Dropdown Switch")
    print(f"{'='*60}")
    
    # Step 1: Pastikan di dashboard
    if "/food/dashboard" not in driver.current_url:
        driver.get(DASHBOARD_URL)
        time.sleep(3)
    
    current_url = driver.current_url
    test.step("Berada di dashboard", "/food/dashboard" in current_url or "partner.shopee" in current_url, current_url)
    
    wait = WebDriverWait(driver, 15)
    
    # Step 2: Deteksi nama merchant saat ini
    try:
        merchant_el = driver.find_element(By.CSS_SELECTOR, ".merchantName")
        current_merchant = merchant_el.text.strip()
        test.step("Elemen .merchantName terdeteksi", True, f"Nama: '{current_merchant}'")
    except:
        test.step("Elemen .merchantName terdeteksi", False, "Tidak ditemukan")
        test.summary()
        return test
    
    # Step 3: Klik profil untuk buka dropdown
    try:
        actions = ActionChains(driver)
        actions.move_to_element(merchant_el).click().perform()
        time.sleep(1)
        test.step("Profil diklik", True)
    except Exception as e:
        test.step("Profil diklik", False, str(e))
        test.summary()
        return test
    
    # Step 4: Cek apakah dropdown terbuka (cari "Pilih Merchant Lain")
    quick_wait = WebDriverWait(driver, 3)
    dropdown_opened = False
    
    try:
        switch_trigger = quick_wait.until(EC.presence_of_element_located(
            (By.XPATH, "//span[contains(text(), 'Pilih Merchant Lain') or contains(text(), 'Switch Merchant')]")
        ))
        # Harus di-click
        ActionChains(driver).move_to_element(switch_trigger).click().perform()
        dropdown_opened = True
        test.step("Dropdown terbuka (ActionChains)", True, "'Pilih Merchant Lain' diklik")
    except:
        # Fallback JS click
        js_found = driver.execute_script("""
            var spans = document.querySelectorAll('span, p, div');
            for (var s of spans) {
                var text = (s.innerText || '').trim();
                if (text.includes('Pilih Merchant Lain') || text.includes('Switch Merchant')) {
                    s.click();
                    return true;
                }
            }
            return false;
        """)
        if js_found:
            dropdown_opened = True
            test.step("Dropdown terbuka (JS fallback)", True, "'Pilih Merchant Lain' ditemukan via JS")
        else:
            test.step("Dropdown terbuka", False, "STALE SESSION atau elemen tidak ditemukan")
    
    if not dropdown_opened:
        test.summary()
        return test
    
    # Step 5: Cek apakah TARGET_NAME ada di daftar dropdown
    time.sleep(1)
    js_switch_script = """
        var targetName = arguments[0].toLowerCase().trim();
        var items = document.querySelectorAll('li.ant-menu-item, li[role="menuitem"], .ant-dropdown-menu-item, [class*="menu-item"]');
        var allMerchants = [];
        var found = false;
        for (var i = 0; i < items.length; i++) {
            var text = (items[i].innerText || "").toLowerCase().trim();
            if (text.length > 0 && text.length < 50) {
                allMerchants.push(text);
            }
            if (text === targetName || text.includes(targetName)) {
                found = true;
            }
        }
        return {found: found, merchants: allMerchants};
    """
    
    target_found = False
    result = None
    for attempt in range(5):
        result = driver.execute_script(js_switch_script, TARGET_NAME)
        if result and result.get("found"):
            target_found = True
            break
        # Scroll down
        try:
            driver.execute_script("document.querySelectorAll('.ant-dropdown-menu, ul[role=\"menu\"], .ant-popover-inner-content').forEach(el => el.scrollTop += 600);")
        except: pass
        time.sleep(1)
    
    if result and result.get("merchants"):
        print(f"\n  📋 Daftar merchant di dropdown ({len(result['merchants'])}):")
        # Hanya print 10 terakhir untuk menghindari spam
        for m in result["merchants"][-10:]:
            marker = "  👉" if TARGET_NAME.lower() in m else "    "
            print(f"  {marker} {m}")
        if len(result["merchants"]) > 10:
            print("     ... dan lainnya")
            
    test.step(
        f"Target '{TARGET_NAME}' ditemukan di dropdown",
        target_found,
        f"{'Ditemukan' if target_found else 'TIDAK ditemukan (FALSE NEGATIVE!)'} — total merchants loaded: {len(result.get('merchants', [])) if result else 0}"
    )
    
    # Step 6: Klik target untuk switch
    if target_found:
        clicked = driver.execute_script("""
            var targetName = arguments[0].toLowerCase().trim();
            var items = document.querySelectorAll('li.ant-menu-item, li[role="menuitem"], .ant-dropdown-menu-item, [class*="menu-item"]');
            for (var i = 0; i < items.length; i++) {
                var text = (items[i].innerText || "").toLowerCase().trim();
                if (text === targetName || text.includes(targetName)) {
                    items[i].scrollIntoView({block: 'center'});
                    items[i].click();
                    return true;
                }
            }
            return false;
        """, TARGET_NAME)
        
        time.sleep(3)
        
        # Verifikasi switch berhasil
        try:
            new_name = driver.find_element(By.CSS_SELECTOR, ".merchantName").text.strip()
            switched = TARGET_NAME.lower() in new_name.lower()
            test.step(
                f"Switch ke '{TARGET_NAME}' berhasil",
                switched,
                f"Nama UI sekarang: '{new_name}'"
            )
        except:
            test.step(f"Switch ke '{TARGET_NAME}' berhasil", False, "Tidak bisa baca .merchantName")
    
    test.summary()
    return test


# ── TEST 2: Logout → Relogin → Onboarding ─────────────────────────────────────

def test_onboarding_selector(driver, username, password, phone):
    """Test logout, relogin, dan klik merchant pertama di onboarding page."""
    test = TestResult("TEST 2: Logout → Relogin → Onboarding (PHASE 1)")
    print(f"\n{'='*60}")
    print(f"  TEST 2: Logout → Relogin → Onboarding")
    print(f"{'='*60}")
    
    wait = WebDriverWait(driver, 30)
    
    # Step 1: Lakukan deliberate logout
    print("\n  [STEP 1] Melakukan deliberate logout...")
    recovered = browser._deliberate_logout_and_relogin(
        driver,
        username=username,
        password=password,
        phone=phone,
    )
    
    test.step("Deliberate logout + relogin berhasil", recovered, f"URL: {driver.current_url}")
    
    if not recovered:
        test.summary()
        return test
    
    # Step 2: Cek URL setelah relogin
    current_url = driver.current_url
    is_onboarding = "onboarding" in current_url or "merchant-selector" in current_url
    is_dashboard = "/food/dashboard" in current_url
    
    test.step(
        "Deteksi halaman pasca-relogin",
        is_onboarding or is_dashboard,
        f"{'ONBOARDING' if is_onboarding else 'DASHBOARD' if is_dashboard else 'UNKNOWN'} — {current_url}"
    )
    
    if is_onboarding:
        # Step 3: Validasi elemen .listItem ada
        time.sleep(3)  # Tunggu DOM fully rendered
        items = dump_onboarding_elements(driver)
        has_list_items = items and len(items) > 0
        test.step("Elemen .listItem ditemukan di onboarding", has_list_items, f"{len(items or [])} item")
        
        # Step 4: Test js_selector_click (PHASE 1 fix)
        js_selector_click = """
            var targetName = arguments[0].toLowerCase().trim();
            var listItems = document.querySelectorAll('.listItem, .merchant-item, li[class*="item"]');
            var firstMerchant = null;
            var foundTarget = false;

            for (var i = 0; i < listItems.length; i++) {
                var el = listItems[i];
                var text = (el.innerText || el.textContent || "").toLowerCase().trim();
                
                if (!firstMerchant && text.length > 0) {
                    firstMerchant = el;
                }

                if (text === targetName || text.includes(targetName)) {
                    el.scrollIntoView({block: 'center'});
                    el.click();
                    foundTarget = true;
                    break;
                }
            }
            
            if (!foundTarget && firstMerchant) {
                firstMerchant.scrollIntoView({block: 'center'});
                firstMerchant.click();
                foundTarget = true;
            }

            return foundTarget;
        """
        
        # Dulu klik merchant pertama (jangan target, karena kita ingin test fallback)
        clicked = driver.execute_script(js_selector_click, "NONEXISTENT_MERCHANT_FOR_TEST")
        test.step("js_selector_click klik merchant pertama (fallback)", clicked)
        
        if clicked:
            # Step 5: Tunggu redirect ke dashboard (30 detik, sesuai fix)
            print(f"\n  ⏳ Menunggu redirect ke dashboard (max 30 detik)...")
            start = time.time()
            redirected = False
            try:
                WebDriverWait(driver, 30).until(lambda d: "/food/dashboard" in d.current_url)
                elapsed = time.time() - start
                redirected = True
                test.step(
                    "Redirect ke dashboard berhasil",
                    True,
                    f"Waktu: {elapsed:.1f} detik — URL: {driver.current_url}"
                )
            except:
                elapsed = time.time() - start
                test.step(
                    "Redirect ke dashboard berhasil",
                    False,
                    f"TIMEOUT setelah {elapsed:.1f} detik — URL: {driver.current_url}"
                )
        
        # Step 6: Test bypass_js (get_session variant)
        # Kita test ini secara terpisah — kembali ke onboarding dulu
        print(f"\n  [STEP 6] Testing bypass_js (get_session variant)...")
        print(f"  Navigasi kembali ke onboarding via return_to_selector...")
        
        try:
            browser.return_to_selector(driver)
            time.sleep(3)
            
            new_url = driver.current_url
            is_back_onboarding = "onboarding" in new_url or "merchant-selector" in new_url
            
            if is_back_onboarding:
                test.step("Kembali ke onboarding untuk test bypass_js", True)
                
                bypass_js = """
                    var loaders = document.querySelectorAll('.ant-spin, [class*="loading"], .shopee-loading, .ant-spin-nested-loading');
                    loaders.forEach(el => el.remove());
                    var target = document.querySelector('.listItem, .merchant-item, .ant-list-item');
                    if (target) {
                        target.scrollIntoView({block: 'center'});
                        target.click();
                        return true;
                    }
                    return false;
                """
                
                clicked2 = driver.execute_script(bypass_js)
                test.step("bypass_js klik merchant pertama", clicked2)
                
                if clicked2:
                    try:
                        start2 = time.time()
                        WebDriverWait(driver, 30).until(lambda d: "/food/dashboard" in d.current_url)
                        elapsed2 = time.time() - start2
                        test.step(
                            "bypass_js redirect ke dashboard",
                            True,
                            f"Waktu: {elapsed2:.1f} detik"
                        )
                    except:
                        elapsed2 = time.time() - start2
                        test.step("bypass_js redirect ke dashboard", False, f"TIMEOUT setelah {elapsed2:.1f} detik")
            else:
                test.step("Kembali ke onboarding untuk test bypass_js", False, f"URL: {new_url}")
        except Exception as e:
            test.step("Test bypass_js", False, f"Error: {e}")
    
    elif is_dashboard:
        test.step("Relogin langsung masuk dashboard (tidak ada onboarding)", True, "Skip onboarding test")
    
    test.summary()
    return test


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    print(f"\n{'='*60}")
    print(f"  🔬 VALIDASI MERCHANT SWITCHING FIX")
    print(f"  Akun: {ACCOUNT_NAME} | Target: {TARGET_NAME}")
    print(f"{'='*60}\n")

    # Load credentials
    env_path = PROJECT_ROOT / "shopee-session-monitor" / ".env"
    creds = {}
    if env_path.exists():
        with open(env_path) as f:
            for line in f:
                line = line.strip()
                if '=' in line and not line.startswith('#'):
                    k, v = line.split('=', 1)
                    creds[k.strip()] = v.strip().strip('"')
    
    # Hardcode credentials dari user
    username = creds.get('SHOPEE_USERNAME_AUTO7307') or "auto7307"
    password = creds.get('SHOPEE_PASSWORD_AUTO7307') or "Auto@7307"
    phone    = creds.get('SHOPEE_PHONE_AUTO7307')    or ""
    
    # Set session file
    browser.set_session_file(SESSION_FILE)
    
    # Buka browser VISIBLE
    print("[INIT] Membuka browser (headless=False)...")
    session = browser.get_session(
        username=username, 
        password=password, 
        phone=phone,
        headless=False, 
        close_browser=False, 
        interactive=False
    )
    
    if not session or "driver" not in session:
        print("❌ Gagal membuka sesi. Pastikan session file ada dan valid.")
        sys.exit(1)
    
    driver = session["driver"]
    print(f"[INIT] Browser terbuka. URL: {driver.current_url}\n")
    
    results = []
    
    try:
        # TEST 1: Dashboard dropdown switch
        t1 = test_dashboard_dropdown_switch(driver)
        results.append(t1)
        
        input("\n⏸️  Tekan ENTER untuk lanjut ke TEST 2 (Logout → Relogin → Onboarding)...")
        
        # TEST 2: Logout + relogin + onboarding
        t2 = test_onboarding_selector(driver, username, password, phone)
        results.append(t2)
        
    except Exception as e:
        print(f"\n❌ Fatal error: {e}")
        import traceback
        traceback.print_exc()
    
    # Final summary
    print(f"\n{'='*60}")
    print(f"  📊 RINGKASAN FINAL")
    print(f"{'='*60}")
    all_passed = all(r.passed for r in results)
    for r in results:
        icon = "✅" if r.passed else "❌"
        print(f"  {icon} {r.name}")
    
    print(f"\n  {'✅ SEMUA TEST PASSED' if all_passed else '❌ ADA TEST YANG GAGAL'}")
    print(f"{'='*60}\n")
    
    input("⏸️  Tekan ENTER untuk menutup browser...")
    driver.quit()
    print("✅ Selesai.")


if __name__ == "__main__":
    main()
