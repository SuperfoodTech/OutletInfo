import asyncio
import json
import logging
import os
import random
import argparse
import pandas as pd
from playwright.async_api import async_playwright

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger("GrabMerchantScraper")

auth_headers = {}

CREDENTIALS = [
    {
        "name": "F1",
        "username": "automationf1",
        "password": "Automationf1@",
        "output": "hasil_custom/F1.xlsx"
    },
    {
        "name": "F2S",
        "username": "automationf2s",
        "password": "Automationf2s@",
        "output": "hasil_custom/F2S.xlsx"
    },
    {
        "name": "W1",
        "username": "automationw1",
        "password": "Automationw1@",
        "output": "hasil_custom/W1.xlsx"
    },
    {
        "name": "L1",
        "username": "automationl1",
        "password": "Automationl1@",
        "output": "hasil_custom/L1.xlsx"
    },
    {
        "name": "L2",
        "username": "automationl2",
        "password": "Automationl2@",
        "output": "hasil_custom/L2.xlsx"
    },
    {
        "name": "DE1S",
        "username": "automationde1s",
        "password": "Automationde1s@",
        "output": "hasil_custom/DE1S.xlsx"
    },
    {
        "name": "JF1",
        "username": "automationjf1",
        "password": "Automationjf1@",
        "output": "hasil_custom/JF1.xlsx"
    },
    {
        "name": "JF1S",
        "username": "automationjf1s",
        "password": "Automationjf1s@",
        "output": "hasil_custom/JF1S.xlsx"
    },
]

BASE_URL   = "https://merchant.grab.com"
LOGIN_URL  = "https://weblogin.grab.com/merchant/login?service_id=MEXUSERS&redirect=https%3A%2F%2Fmerchant.grab.com%2Fportal"
DASHBOARD_URL = f"{BASE_URL}/dashboard"
INVENTORY_URL = f"{BASE_URL}/food/inventory"


async def human_type(locator, text):
    """Mengetik teks dengan jeda acak antar karakter agar terlihat seperti manusia."""
    await locator.click()
    for char in text:
        await locator.press(char)
        await asyncio.sleep(random.uniform(0.07, 0.18))


async def perform_login(page, username, password):
    logger.info(f"Navigating to Grab Merchant login page for {username}...")
    try:
        await page.goto(LOGIN_URL, wait_until="networkidle", timeout=60000)
        
        # Tambahan waktu tunggu untuk memastikan JavaScript selesai me-render form
        logger.info("Menunggu halaman login ter-load sepenuhnya...")
        await page.wait_for_timeout(5000)

        # ── Step 1: Pastikan tab Username aktif ────────────────────────
        logger.info("Memastikan tab 'Username' aktif...")
        username_tab = page.get_by_role("tab", name="Username")
        if await username_tab.is_visible(timeout=5000):
            await username_tab.click()
            await asyncio.sleep(random.uniform(0.4, 0.8))

        # ── Step 2: Isi username ───────────────────────────────────────
        logger.info(f"Mengetik username: {username}")
        username_input = page.locator('input[type="text"]').first
        await username_input.wait_for(state="visible", timeout=15000)
        await asyncio.sleep(random.uniform(0.3, 0.7))
        await human_type(username_input, username)
        await asyncio.sleep(random.uniform(0.5, 1.0))

        # ── Step 3: Klik Continue (pertama) ───────────────────────────
        logger.info("Klik Continue (username)...")
        continue_btn = page.get_by_role("button", name="Continue")
        await continue_btn.wait_for(state="visible", timeout=10000)
        await asyncio.sleep(random.uniform(0.3, 0.6))
        await continue_btn.click()

        # ── Step 4: Tunggu form password muncul ───────────────────────
        logger.info("Menunggu form password...")
        password_input = page.locator('input[type="password"]')
        await password_input.wait_for(state="visible", timeout=15000)
        await asyncio.sleep(random.uniform(0.4, 0.9))

        # ── Step 5: Isi password ──────────────────────────────────────
        logger.info("Mengetik password...")
        await human_type(password_input, password)
        await asyncio.sleep(random.uniform(0.5, 1.0))

        # ── Step 6: Klik Continue (kedua) ─────────────────────────────
        logger.info("Klik Continue (password)...")
        continue_btn2 = page.get_by_role("button", name="Continue")
        await continue_btn2.wait_for(state="visible", timeout=10000)
        await asyncio.sleep(random.uniform(0.2, 0.5))
        await continue_btn2.click()

        # ── Step 7: Tunggu redirect ke dashboard ─────────────────────
        logger.info("Menunggu redirect ke dashboard...")
        await page.wait_for_url(f"{BASE_URL}/**", timeout=45000)
        logger.info(f"[✓] Login berhasil untuk {username} → {page.url}")
        return True

    except Exception as e:
        logger.error(f"[✗] Login gagal untuk {username}: {e}")
        # Simpan screenshot untuk debug
        try:
            screenshot_path = f"login_error_{username}.png"
            await page.screenshot(path=screenshot_path)
            logger.info(f"Screenshot disimpan: {screenshot_path}")
        except Exception:
            pass
        return False


async def capture_auth_headers(page):
    global auth_headers

    def handle_request(request):
        headers = request.headers
        if "authorization" in headers or "x-grab-merchant" in headers:
            auth_headers.update({
                k: v for k, v in headers.items()
                if k.lower() in ["authorization", "x-grab-merchant", "x-grab-device-id",
                                  "x-grab-client-id", "content-type", "accept", "x-passkey"]
            })

    page.on("request", handle_request)


async def fetch_idmg(page):
    """Mencari Group ID (idmg) dari API merchant-selector."""
    logger.info("Mencari Group ID (idmg) dari merchant-selector...")
    try:
        url = f"{BASE_URL}/troy/user-profile/v1/merchant-selector"
        response = await page.request.get(url, headers=auth_headers)
        if response.ok:
            import re
            text = await response.text()
            match = re.search(r'(IDMG\d+)', text)
            if match:
                idmg = match.group(1)
                logger.info(f"Berhasil mendapatkan Group ID: {idmg}")
                return idmg
        logger.warning(f"Gagal mendapatkan idmg. HTTP Status: {response.status}")
    except Exception as e:
        logger.error(f"Error fetching idmg: {e}")
    return ""


async def fetch_merchant_list(page):
    """Fetch list of all merchant stores using the search API with pagination."""
    stores = []
    total_expected = 0
    logger.info("Fetching merchant store list...")

    try:
        offset = 0
        limit = 100
        
        while True:
            api_url = f"https://api.grab.com/delvplatformapi/merchant/v1/merchant-group/store/search?offset={offset}&limit={limit}&search=&includeItemsWithoutPhotosCount=true&includeInactive=true&modelType=integrated&asc=true&cityIDs[]=ALL&includeMenuGroupV2ID=false"
            logger.info(f"Fetching stores offset {offset}...")
            
            response = await page.request.get(api_url, headers=auth_headers)
            if not response.ok:
                logger.warning(f"Gagal mengambil store list pada offset {offset}: HTTP {response.status}")
                break
                
            data = await response.json()
            
            # Ambil totalCount dari response API pada halaman pertama
            if offset == 0:
                total_expected = (
                    data.get("totalCount") or
                    data.get("total") or
                    data.get("data", {}).get("totalCount") or
                    0
                )
            
            fetched_stores = (
                data.get("stores") or
                data.get("data", {}).get("stores") or
                data.get("merchantDetails") or
                data.get("data", {}).get("merchantDetails") or
                data.get("catalogStores") or
                []
            )
            
            # fallback pemetaan jika array stores ada di letak berbeda
            if not fetched_stores and isinstance(data, list):
                fetched_stores = data
            elif not fetched_stores and isinstance(data, dict):
                for k, v in data.items():
                    if isinstance(v, list) and len(v) > 0 and isinstance(v[0], dict) and ("merchantID" in v[0] or "merchantId" in v[0] or "id" in v[0]):
                        fetched_stores = v
                        break
            
            if not fetched_stores or not isinstance(fetched_stores, list):
                break
                
            stores.extend(fetched_stores)
            logger.info(f"Berhasil mengambil {len(fetched_stores)} outlet. Total terkumpul: {len(stores)}")
            
            if len(fetched_stores) < limit:
                break
                
            offset += limit

    except Exception as e:
        logger.error(f"Error fetching merchant list: {e}")

    # Bersihkan duplikat barangkali ada ID ganda
    unique_stores = {}
    for s in stores:
        mid = str(s.get("merchantID") or s.get("merchantId") or s.get("id") or "").strip()
        if mid:
            unique_stores[mid] = s

    return list(unique_stores.values()), total_expected






async def run_scraper_for_credential(playwright, cred):
    """Run scraper for a single credential set."""
    global auth_headers
    auth_headers = {}

    logger.info(f"\n{'='*50}")
    logger.info(f"Starting scraper for: {cred['name']}")
    logger.info(f"{'='*50}")

    browser = await playwright.chromium.launch(
        headless=False,
        args=["--disable-blink-features=AutomationControlled"]
    )

    context = await browser.new_context(
        viewport={"width": 1280, "height": 800},
        user_agent="Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )

    page = await context.new_page()
    await capture_auth_headers(page)

    try:
        # Login
        logged_in = await perform_login(page, cred["username"], cred["password"])
        if not logged_in:
            logger.error(f"Failed to login for {cred['name']}, skipping...")
            return

        await page.wait_for_timeout(2000)

        # Masuk ke menu dulu agar cookies/headers terkait portal ter-load penuh
        await page.goto(f"{BASE_URL}/food/menu", wait_until="networkidle", timeout=60000)
        await page.wait_for_timeout(3000)

        # Get IDMG (Group ID) from merchant-selector API
        idmg = await fetch_idmg(page)
        if not idmg:
            logger.error(f"Tidak dapat menemukan Group ID untuk {cred['name']}, melewati.")
            return

        # Get list of stores dari API
        stores, total_expected = await fetch_merchant_list(page)
        logger.info(f"Total unique stores extracted: {len(stores)} (Target expected from API: {total_expected})")

        # Log Pengecekan
        if total_expected > 0 and len(stores) < total_expected:
            logger.warning(f"[!] PERINGATAN: Jumlah outlet yang terekstrak ({len(stores)}) "
                           f"lebih kecil dari total sistem ({total_expected}). "
                           "Beberapa outlet mungkin terlewat/gagal ditarik dari API.")
        elif total_expected > 0 and len(stores) >= total_expected:
            logger.info("[✓] Pengecekan: Semua outlet berhasil ditarik sesuai dengan total di sistem Grab.")

        if not stores:
            logger.warning(f"No stores found for {cred['name']}")

        all_results = []
        
        for store in stores:
            merchant_id = (
                str(store.get("merchantID") or store.get("merchantId") or store.get("id") or "")
            ).strip()
            store_name = store.get("name") or store.get("merchantName") or store.get("storeName") or merchant_id
            status = store.get("status") or store.get("isActive") or ""
            alamat = store.get("address") or store.get("merchantAddress") or ""
            
            # Coba ambil Bank Account jika tersedia di payload API
            bank_account = (
                store.get("bankAccount") or 
                store.get("bankAccountNumber") or 
                store.get("accountNumber") or 
                ""
            )

            if not merchant_id:
                continue

            all_results.append({
                "Portal": cred["name"],
                "Nama": store_name,
                "Group ID": idmg,
                "Store ID": merchant_id,
                "Status": status,
                "Alamat": alamat,
                "Bank Account": bank_account
            })
            
        if all_results:
            os.makedirs(os.path.dirname(cred["output"]) if os.path.dirname(cred["output"]) else ".", exist_ok=True)
            df = pd.DataFrame(all_results)
            df.to_excel(cred["output"], index=False)
            logger.info(f"Results saved to: {cred['output']} ({len(df)} rows)")
        else:
            logger.warning(f"No results collected for {cred['name']}")

    except Exception as e:
        logger.error(f"Unexpected error for {cred['name']}: {e}")
    finally:
        await context.close()
        await browser.close()


async def main():
    parser = argparse.ArgumentParser(description="Grab Merchant Scraper")
    parser.add_argument("--outlet", type=str, help="Specify the outlet name to run (e.g., F1)")
    args = parser.parse_args()

    target_credentials = CREDENTIALS
    if args.outlet:
        target_credentials = [c for c in CREDENTIALS if c["name"] == args.outlet]
        if not target_credentials:
            logger.error(f"Outlet '{args.outlet}' not found in credentials.")
            return

    logger.info("Starting Grab Merchant Scraper")
    logger.info(f"Processing {len(target_credentials)} credential set(s)")

    async with async_playwright() as playwright:
        for cred in target_credentials:
            await run_scraper_for_credential(playwright, cred)

    logger.info("\nAll done!")


if __name__ == "__main__":
    asyncio.run(main())
