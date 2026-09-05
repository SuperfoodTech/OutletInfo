import os
import json
import asyncio
import time
from datetime import datetime
from playwright.async_api import async_playwright
from core.client import ShopeeClient
from core.logger import get_logger

log = get_logger("shopee_scraper")

async def perform_shopee_login(page, username, password):
    """Handles Shopee Partner login flow."""
    log.info(f"  [Login] Navigating to Shopee login...")
    await page.goto("https://partner.shopee.co.id/login", wait_until="networkidle")
    
    if "login" not in page.url:
        return True

    try:
        user_selector = "input[name='userName'], input[placeholder*='Username'], input[placeholder*='Handphone']"
        await page.wait_for_selector(user_selector, timeout=30000)
        await page.fill("input[name='userName']" if await page.is_visible("input[name='userName']") else "input", username)
        await page.fill("input[type='password']", password)
        
        login_btn = "button.shopee-button--primary, button:has-text('Log In'), button:has-text('Masuk')"
        await page.wait_for_selector(login_btn)
        await page.click(login_btn)
        
        log.info("  [Login] Waiting for manual OTP / Redirect (No timeout)...")
        try:
            await page.wait_for_url(lambda u: "login" not in u.lower() and "authenticate" not in u.lower() or "dashboard" in u.lower(), timeout=0)
            return True
        except:
            if "login" not in page.url.lower(): return True
            return False
    except Exception as e:
        log.error(f"  ✗ [Login] Error: {e}")
        return False

async def switch_shopee_merchant(page, target_name):
    """
    Automated merchant switch matching the legacy Selenium logic EXACTLY.
    """
    try:
        log.info(f"  🔄 [SWITCH] Target: {target_name}...")
        
        # 1. Handle initial merchant selector page (right after login)
        if "merchant-selector" in page.url or "onboarding" in page.url:
            log.info("  📍 [SWITCH] On selector page. Selecting merchant...")
            success = await select_merchant_on_selector_page(page, target_name)
            if success: return True

        # 2. Natural Navigation from Dashboard
        if "/food/dashboard" not in page.url:
            await page.goto("https://partner.shopee.co.id/food/dashboard", wait_until="domcontentloaded")
            await page.wait_for_timeout(2000)

        # Check if already on correct merchant
        try:
            current_el = page.locator(".merchantName").first
            if await current_el.count() > 0:
                if target_name.lower() in (await current_el.inner_text()).lower():
                    log.info(f"  ✨ [SWITCH] Already on {target_name}")
                    return True
        except: pass

        # 3. Open Dropdown naturally
        log.info("  🖱️ [SWITCH] Opening profile menu...")
        await page.click(".merchantName", timeout=10000)
        await page.wait_for_timeout(1000)

        # 4. Click "Pilih Merchant Lain" (Natural trigger)
        selector_trigger = page.locator("//span[text()='Pilih Merchant Lain' or text()='Switch Merchant' or contains(text(), 'Ubah')]").first
        if await selector_trigger.count() > 0:
            await selector_trigger.click()
            await page.wait_for_url(lambda u: "merchant-selector" in u, timeout=15000)
            await page.wait_for_timeout(2000)
            
            # 5. Select on the selector page
            return await select_merchant_on_selector_page(page, target_name)
        else:
            log.warning("  ⚠️ [SWITCH] Could not find switch trigger in menu.")
            return False
            
    except Exception as e:
        log.error(f"  ✗ [SWITCH] Error: {e}")
        return False

async def select_merchant_on_selector_page(page, target_name):
    """Uses the exact JS logic from legacy Selenium to click a merchant."""
    js_selector_click = """
        (target) => {
            var targetName = target.toLowerCase().trim();
            var labels = document.querySelectorAll('span, div, li, p, .shop-name');
            var found = false;
            for (var i = 0; i < labels.length; i++) {
                var text = (labels[i].innerText || "").toLowerCase().trim();
                if (text === targetName || (text.includes(targetName) && labels[i].children.length < 3)) {
                    labels[i].scrollIntoView({block: 'center'});
                    labels[i].click();
                    found = true;
                    break;
                }
            }
            if (!found) return false;
            
            // Confirm Click (Masuk / Konfirmasi)
            setTimeout(() => {
                var btns = document.querySelectorAll('button');
                for (var b of btns) {
                    var bText = (b.innerText || "").toLowerCase();
                    if (bText.includes('masuk') || bText.includes('konfirmasi') || bText.includes('lanjutkan') || bText.includes('ok')) {
                        b.click();
                        break;
                    }
                }
            }, 800);
            return true;
        }
    """
    for attempt in range(1, 6):
        log.info(f"  🔎 [SCAN] Selecting merchant (Attempt {attempt}/5)...")
        if await page.evaluate(js_selector_click, target_name):
            try:
                await page.wait_for_url(lambda u: "/food/dashboard" in u, timeout=15000)
                # Check for "Unexpected Error" modal (like Selenium)
                error_modal = page.locator("//*[contains(text(), 'unexpected error') or contains(text(), 'Terjadi kesalahan')]").first
                if await error_modal.count() > 0 and await error_modal.is_visible():
                    log.warning("  ⚠️ [SWITCH] Shopee error modal detected. Refreshing...")
                    await page.reload()
                    await page.wait_for_timeout(3000)
                return True
            except: pass
        await page.evaluate("window.scrollBy(0, 300)")
        await page.wait_for_timeout(1500)
    return False

async def trigger_export_current_merchant(page, merchant_name, start_time, end_time):
    """Extracts tokens and triggers API."""
    try:
        await page.goto("https://partner.shopee.co.id/settings/shopee-food/business-hours-settings", wait_until="domcontentloaded")
        m_map = {}
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            map_path = os.path.join(base_dir, "API", "response.json")
            if os.path.exists(map_path):
                with open(map_path, "r") as f:
                    data = json.load(f)
                    for m in data.get("data", {}).get("selectMerchant", {}).get("merchantList", []):
                        m_map[m["merchantName"].lower()] = str(m["merchantId"])
        except: pass
        
        entity_id = m_map.get(merchant_name.lower())
        tob_token = None
        for _ in range(10):
            cookies = await page.context.cookies()
            tob_token = next((c['value'] for c in cookies if c['name'] == 'shopee_tob_token'), None)
            if tob_token: break
            await page.wait_for_timeout(1000)
            
        cookie_dict = {c['name']: c['value'] for c in await page.context.cookies()}
        if not tob_token or not entity_id:
            return None, f"Tokens missing (Token: {'Yes' if tob_token else 'No'}, ID: {entity_id})"

        client = ShopeeClient(tob_token=tob_token, entity_id=entity_id, extra_cookies=cookie_dict)
        res = client.export_transaction_report(merchant_ids=[entity_id], start_time=start_time, end_time=end_time)
        return ({"entity_id": entity_id, "tob_token": tob_token, "cookies": cookie_dict}, None) if res is True else (None, "API Trigger Failed")
    except Exception as e: return None, str(e)
