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
                return idmg, response.status
        logger.warning(f"Gagal mendapatkan idmg. HTTP Status: {response.status}")
        return "", response.status
    except Exception as e:
        logger.error(f"Error fetching idmg: {e}")
    return "", 0


async def fetch_merchant_list(page):
    """Fetch list of all merchant stores using the search API with pagination."""
    stores = []
    total_expected = 0
    logger.info("Fetching merchant store list...")

    try:
        for m_type in ["integrated", "non-integrated", "ALL"]:
            for is_asc in ["true", "false"]:
                logger.info(f"--- Mulai penarikan dengan metode Sorting (asc={is_asc}, modelType={m_type}) ---")
                offset = 0
                limit = 100
                
                while True:
                    m_param = "" if m_type == "ALL" else f"&modelType={m_type}"
                    api_url = f"https://api.grab.com/delvplatformapi/merchant/v1/merchant-group/store/search?offset={offset}&limit={limit}&search=&includeItemsWithoutPhotosCount=true&includeInactive=true{m_param}&asc={is_asc}&cityIDs[]=ALL&includeMenuGroupV2ID=false"
                    
                    page_retries = 3
                    fetched_stores = []
                    
                    for attempt in range(page_retries):
                        logger.info(f"Fetching stores offset {offset} (attempt {attempt+1})...")
                        response = await page.request.get(api_url, headers=auth_headers)
                        if not response.ok:
                            logger.warning(f"Gagal mengambil store list pada offset {offset}: HTTP {response.status}")
                            await asyncio.sleep(2)
                            continue
                            
                        data = await response.json()
                        
                        # Ambil totalCount dari response API pada halaman pertama
                        if offset == 0 and is_asc == "true" and m_type == "integrated":
                            total_expected = (
                                data.get("totalCount") or
                                data.get("total") or
                                data.get("data", {}).get("totalCount") or
                                total_expected
                            )
                            logger.info(f"Target total merchant dari sistem Grab untuk portal ini: {total_expected}")
                        
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
                                    
                        # Berhasil dapat data, break out of retry loop
                        break
                    
                    if not fetched_stores or not isinstance(fetched_stores, list):
                        logger.info("Tidak ada lagi data/store yang ditemukan, mengakhiri paginasi arah ini.")
                        break
                        
                    stores.extend(fetched_stores)
                    logger.info(f"Berhasil mengambil {len(fetched_stores)} outlet. Total terkumpul saat ini: {len(stores)}")
                    
                    if len(fetched_stores) < limit:
                        logger.info(f"Jumlah data kurang dari limit ({limit}), mengakhiri arah ini.")
                        break
                        
                    offset += limit
                
                # Tambahkan jeda waktu agar tidak mengambil data terlalu cepat (mencegah rate-limiting)
                logger.info("Jeda 2 detik sebelum mengambil halaman berikutnya...")
                await asyncio.sleep(2.0)

    except Exception as e:
        logger.error(f"Error fetching merchant list: {e}")

    # Mengembalikan semua data mentah tanpa membersihkan duplikat sesuai request
    return stores, total_expected






async def run_scraper_for_credential(playwright, cred, force_fresh=False):
    """Run scraper for a single credential set."""
    global auth_headers
    auth_headers = {}

    logger.info(f"\n{'='*50}")
    if force_fresh:
        logger.info(f"Starting FRESH scraper for: {cred['name']}")
    else:
        logger.info(f"Starting scraper for: {cred['name']}")
    logger.info(f"{'='*50}")

    browser = await playwright.chromium.launch(
        headless=True,
        args=["--disable-blink-features=AutomationControlled"]
    )

    os.makedirs(os.path.join(os.path.dirname(__file__), "sessions"), exist_ok=True)
    session_file = os.path.join(os.path.dirname(__file__), "sessions", f"grab_session_{cred['name']}.json")

    context_options = {
        "viewport": {"width": 1280, "height": 800},
        "user_agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }

    # Jika force_fresh, hapus file session yang ada
    if force_fresh and os.path.exists(session_file):
        os.remove(session_file)
        logger.info(f"Dihapus session lama untuk fresh login: {session_file}")

    # Load session jika file session sudah ada
    if os.path.exists(session_file) and not force_fresh:
        context_options["storage_state"] = session_file
        logger.info(f"Menggunakan session yang tersimpan: {session_file}")

    context = await browser.new_context(**context_options)

    page = await context.new_page()
    await capture_auth_headers(page)

    try:
        # Cek apakah session masih valid dengan mencoba masuk ke halaman menu
        logger.info("Mengecek validitas session...")
        await page.goto(f"{BASE_URL}/food/menu", wait_until="networkidle", timeout=60000)
        await page.wait_for_timeout(3000)

        # Jika URL redirect ke halaman login atau logout, berarti session belum ada / sudah expired
        if "login" in page.url.lower() or "logout" in page.url.lower():
            logger.info("Session tidak valid/expired. Memulai proses login...")
            logged_in = await perform_login(page, cred["username"], cred["password"])
            if not logged_in:
                if not force_fresh:
                    logger.warning(f"Gagal login dengan session saat ini untuk {cred['name']}. Akan mencoba FRESH LOGIN...")
                    return "RETRY_FRESH"
                else:
                    logger.error(f"Fresh login gagal untuk {cred['name']}, melewati (skipping)...")
                    return False

            # Simpan state/session setelah berhasil login
            await context.storage_state(path=session_file)
            logger.info(f"Session baru berhasil disimpan ke {session_file}")

            # Pindah lagi ke halaman menu setelah login berhasil
            await page.goto(f"{BASE_URL}/food/menu", wait_until="networkidle", timeout=60000)
            await page.wait_for_timeout(3000)
        else:
            logger.info("[✓] Session masih valid! Lewati proses login.")

        # Get IDMG (Group ID) from merchant-selector API
        idmg, idmg_status = await fetch_idmg(page)
        
        # Jika status 401, berarti session expired walaupun URL tidak berubah ke halaman login/logout
        if idmg_status == 401:
            logger.info("Session expired (HTTP 401 pada merchant-selector). Memulai proses login ulang...")
            logged_in = await perform_login(page, cred["username"], cred["password"])
            if not logged_in:
                if not force_fresh:
                    logger.warning(f"Gagal login ulang (401) dengan session saat ini untuk {cred['name']}. Akan mencoba FRESH LOGIN...")
                    return "RETRY_FRESH"
                else:
                    logger.error(f"Fresh login (401) gagal untuk {cred['name']}, melewati (skipping)...")
                    return False

            # Simpan state/session setelah berhasil login
            await context.storage_state(path=session_file)
            logger.info(f"Session baru berhasil disimpan ke {session_file}")

            # Pindah lagi ke halaman menu setelah login berhasil
            await page.goto(f"{BASE_URL}/food/menu", wait_until="networkidle", timeout=60000)
            await page.wait_for_timeout(3000)
            
            # Coba ambil idmg lagi
            idmg, idmg_status = await fetch_idmg(page)

        if not idmg:
            logger.error(f"Tidak dapat menemukan Group ID untuk {cred['name']}, melewati.")
            return False

        # Get list of stores dari API
        stores, total_expected = await fetch_merchant_list(page)
        
        max_retries = 3
        retry_count = 0
        while total_expected > 0 and len(stores) < total_expected and retry_count < max_retries:
            logger.warning(f"[!] PERINGATAN: Jumlah outlet yang terekstrak ({len(stores)}) "
                           f"kurang dari total sistem ({total_expected}). "
                           f"Mencoba mengambil ulang data yang terlewat... ({retry_count + 1}/{max_retries})")
            
            await page.wait_for_timeout(3000)
            stores_retry, _ = await fetch_merchant_list(page)
            
            # Gabungkan dengan data sebelumnya tanpa menghapus duplikat
            stores.extend(stores_retry)
            retry_count += 1
            
        logger.info(f"Total unique stores extracted finally: {len(stores)} (Target expected from API: {total_expected})")

        # Log Pengecekan Akhir
        if total_expected > 0 and len(stores) < total_expected:
            logger.warning(f"[!] GAGAL MENGAMBIL SEMUA: Setelah {max_retries} percobaan, jumlah outlet terekstrak ({len(stores)}) "
                           f"masih lebih kecil dari total sistem ({total_expected}).")
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
            output_dir = os.path.dirname(cred["output"]) if os.path.dirname(cred["output"]) else "."
            os.makedirs(output_dir, exist_ok=True)
            df = pd.DataFrame(all_results)
            
            # Versioning file: F1.xlsx -> F1_2.xlsx -> F1_3.xlsx
            base_path, ext = os.path.splitext(cred["output"])
            output_file = cred["output"]
            counter = 2
            while os.path.exists(output_file):
                output_file = f"{base_path}_{counter}{ext}"
                counter += 1
                
            df.to_excel(output_file, index=False)
            logger.info(f"Results saved to: {output_file} ({len(df)} rows)")
        else:
            logger.warning(f"No results collected for {cred['name']}")
            
        return True

    except Exception as e:
        logger.error(f"Unexpected error for {cred['name']}: {e}")
        return False
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
            result = await run_scraper_for_credential(playwright, cred)
            if result == "RETRY_FRESH":
                result = await run_scraper_for_credential(playwright, cred, force_fresh=True)
                
            retry_count = 0
            while result is False and retry_count < 2:
                retry_count += 1
                logger.warning(f"[!] Portal {cred['name']} gagal diproses/login. Melakukan retry ke-{retry_count} dari 2...")
                await asyncio.sleep(5)
                result = await run_scraper_for_credential(playwright, cred, force_fresh=True)

    logger.info("\n[✓] Proses scraping selesai.")
    
    # Menjalankan proses penggabungan dan pembersihan duplikat
    logger.info("\n" + "="*50)
    logger.info("Menjalankan post-processing (Combine & Remove Duplicates)")
    logger.info("="*50)
    
    try:
        from combine_custom import combine_excel_files
        from find_duplicates import find_duplicates
        from remove_dup_custom import remove_duplicates_from_files
        
        # 1. Combine all results to MASTER_ALL.xlsx
        logger.info("\n--- 1. Menggabungkan file ---")
        combine_excel_files()
        
        import shutil
        import datetime
        from format_excel import apply_formatting_and_sheets
        
        timestamp = datetime.datetime.now().strftime("%Y%m%d %H%M")
        raw_filename = f"GRAB {timestamp}.xlsx"
        try:
            apply_formatting_and_sheets("hasil_custom/0MASTER.xlsx", f"hasil_custom/{raw_filename}", create_tabs=False)
            logger.info(f"Data mentah disimpan dengan format: hasil_custom/{raw_filename}")
        except Exception as ex:
            logger.error(f"Gagal menyimpan {raw_filename}: {ex}")
        
        # 2. Find duplicates in MASTER_ALL.xlsx
        logger.info("\n--- 2. Mencari duplikat ---")
        find_duplicates()
        
        # 3. Remove duplicates from individual files
        logger.info("\n--- 3. Menghapus duplikat dari file individual ---")
        remove_duplicates_from_files()
        
        # 4. Re-combine to ensure MASTER_ALL.xlsx is clean
        logger.info("\n--- 4. Menggabungkan kembali file yang sudah bersih ---")
        combine_excel_files()
        
        processed_filename = f"GRAB PROCESSED {timestamp}.xlsx"
        try:
            apply_formatting_and_sheets("hasil_custom/0MASTER.xlsx", f"hasil_custom/{processed_filename}", create_tabs=True)
            logger.info(f"Data bersih disimpan dengan format dan tabs: hasil_custom/{processed_filename}")
        except Exception as ex:
            logger.error(f"Gagal menyimpan {processed_filename}: {ex}")
        
    except Exception as e:
        logger.error(f"Gagal menjalankan post-processing: {e}")

    logger.info("\nAll done!")


if __name__ == "__main__":
    asyncio.run(main())
