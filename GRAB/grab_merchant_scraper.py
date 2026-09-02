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

import urllib.request
import csv
import sys

import os
import time

GOOGLE_SHEET_AGENCY_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ3tLKBNXDqRgBw0mNhKZFxgvKx-JoiTDzm_s5Ix1cm7O6HCv4IvExOLR2HSRVaXSsx82V348mcr9X4/pub?output=csv"
GOOGLE_SHEET_VB_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRYSUnKOqk29LCktTxdb0wPLbWMbRaWRP3eC_UA4AwYod1FW6zDMhtLMC5ghIvot2B8upCDfBsn-TCP/pub?gid=978201567&single=true&output=csv"
LOCAL_CRED_CSV = os.path.join(os.path.dirname(__file__), '..', 'A. Credential (Outlet & Access)  - Unique Portal Gr (1).csv')


def get_credentials_from_sheet(source_type="agency", custom_url=None):
    """Mengambil daftar kredensial portal Grab berdasarkan tipe sumber (Agency atau VB)."""
    data = []
    
    target_url = custom_url
    if not target_url:
        if source_type.lower() == "vb":
            target_url = GOOGLE_SHEET_VB_URL or os.getenv("GRAB_VB_CSV_URL", "")
        else:
            target_url = GOOGLE_SHEET_AGENCY_URL

    # 1. Coba ambil secara live dari Google Sheet URL
    if target_url:
        try:
            logger.info(f"[*] Mengambil data portal Grab [{source_type.upper()}] langsung dari Google Sheet (Live URL)...")
            req = urllib.request.Request(target_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=20) as response:
                content = response.read().decode('utf-8')
                reader = csv.reader(content.splitlines())
                data = list(reader)
                logger.info(f"[✓] Berhasil memuat {len(data)} baris data [{source_type.upper()}] dari Google Sheet.")
        except Exception as e:
            logger.warning(f"⚠️ Gagal fetch live dari Google Sheet [{source_type.upper()}]: {e}.")

    # 2. Fallback ke file CSV lokal jika live fetch gagal (khusus Agency)
    if not data and source_type.lower() == "agency" and os.path.exists(LOCAL_CRED_CSV):
        try:
            logger.info(f"[*] Membaca data portal dari fallback lokal: {os.path.basename(LOCAL_CRED_CSV)}")
            with open(LOCAL_CRED_CSV, "r", encoding="utf-8") as f:
                reader = csv.reader(f)
                data = list(reader)
        except Exception as ex:
            logger.error(f"Gagal membaca fallback CSV lokal: {ex}")

    if not data:
        logger.error(f"❌ Tidak ada data portal yang dapat dimuat untuk tipe [{source_type.upper()}].")
        return []

    # Dynamic column indexing
    header = [str(h).strip().lower() for h in data[0]]
    col_app = -1
    col_owner = -1
    col_portal = -1
    col_user = -1
    col_pass = -1
    col_notes = -1

    if source_type.lower() == "agency":
        col_app = 3
        col_owner = 0
        col_portal = 2
        col_user = 26
        col_pass = 28
    else:
        for i, h in enumerate(header):
            if any(k == h or k in h for k in ['portal', 'nama portal', 'brand', 'nama brand', 'outlet']):
                col_portal = i
            elif any(k == h or k in h for k in ['username', 'nama pengguna', 'user login', 'user']):
                col_user = i
            elif any(k == h or k in h for k in ['password', 'kata sandi', 'pass login', 'pass']):
                col_pass = i
            elif any(k == h or k in h for k in ['notes', 'catatan', 'keterangan']):
                col_notes = i
            elif any(k == h or k in h for k in ['owner', 'pemilik']):
                col_owner = i
            elif any(k == h or k in h for k in ['aplikasi', 'platform', 'app']):
                col_app = i

    portals = []
    for row in data[1:]:
        # Filter status Restricted
        if col_notes != -1 and col_notes < len(row):
            note_val = str(row[col_notes]).strip().lower()
            if "restricted" in note_val:
                p_name = row[col_portal].strip() if col_portal != -1 and col_portal < len(row) else "Unknown"
                logger.info(f"🚫 Melewati portal '{p_name}' karena berstatus Restricted.")
                continue

        is_grab = True
        if col_app != -1 and col_app < len(row) and row[col_app].strip():
            is_grab = "grab" in row[col_app].strip().lower()

        if is_grab and len(row) > max(col_user, col_pass):
            owner = row[col_owner].strip() if col_owner != -1 and col_owner < len(row) else ("VB" if source_type.lower() == "vb" else "")
            portal = row[col_portal].strip() if col_portal != -1 and col_portal < len(row) else ""
            brand = portal.split(" - ")[0].strip() if " - " in portal else portal
            username = row[col_user].strip() if col_user != -1 and col_user < len(row) else ""
            password = row[col_pass].strip() if col_pass != -1 and col_pass < len(row) else ""
            
            if username and username != "-" and password and password != "-":
                safe_portal_name = "".join([c for c in portal if c.isalpha() or c.isdigit() or c == ' ']).rstrip()
                hasil_custom_dir = os.path.join(os.path.dirname(__file__), "hasil_custom")
                portals.append({
                    "owner": owner if owner else "VB",
                    "brand": brand,
                    "name": portal,
                    "username": username,
                    "password": password,
                    "source_type": source_type.upper(),
                    "output": os.path.join(hasil_custom_dir, f"{safe_portal_name}.xlsx")
                })

    logger.info(f"[✓] Terdaftar {len(portals)} portal GrabFood valid [{source_type.upper()}] untuk diproses.")
    return portals

def parse_selection(choice, max_val):
    selected = set()
    choice = choice.strip()
    if choice.lower() in ('0', 'all'):
        return list(range(max_val))
        
    parts = choice.split(',')
    for part in parts:
        part = part.strip()
        if '-' in part:
            try:
                start, end = map(int, part.split('-'))
                if 1 <= start <= end <= max_val:
                    for i in range(start, end + 1):
                        selected.add(i - 1)
            except ValueError:
                pass
        else:
            try:
                val = int(part)
                if 1 <= val <= max_val:
                    selected.add(val - 1)
            except ValueError:
                pass
    return sorted(list(selected))

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


async def handle_welcome_back_or_continue(page):
    """Mengecek dan mengklik tombol Continue jika halaman menampilkan Welcome back / Saved Account."""
    try:
        # Tunggu sampai tombol Continue muncul jika halaman sedang me-render saved account
        btn_loc = page.locator('button:has-text("Continue"), button:has-text("Lanjut"), button:has-text("Masuk")').first
        try:
            await btn_loc.wait_for(state="visible", timeout=6000)
        except Exception:
            pass

        if await btn_loc.is_visible():
            btn_text = (await btn_loc.inner_text()).strip()
            logger.info(f"👉 [Saved Account Terdeteksi] Mengklik tombol '{btn_text}' pada layar Welcome back...")
            await btn_loc.click()
            await page.wait_for_timeout(4000)
            return True
    except Exception as e:
        logger.debug(f"Check continue button exception: {e}")
    return False


async def perform_login(page, username, password):
    logger.info(f"Navigating to Grab Merchant login page for {username}...")
    try:
        # 1. Cek dulu apakah halaman saat ini sudah menampilkan tombol Continue (Welcome back)
        if await handle_welcome_back_or_continue(page):
            if "login" not in page.url.lower() and "saved-accounts" not in page.url.lower():
                logger.info(f"[✓] Berhasil masuk via tombol Continue untuk {username}!")
                return True

        await page.goto(LOGIN_URL, wait_until="networkidle", timeout=60000)
        
        # Tambahan waktu tunggu untuk memastikan JavaScript selesai me-render form
        logger.info("Menunggu halaman login ter-load sepenuhnya...")
        await page.wait_for_timeout(4000)

        # Cek sekali lagi apakah redirect login menghasilkan layar Welcome back
        if await handle_welcome_back_or_continue(page):
            if "login" not in page.url.lower() and "saved-accounts" not in page.url.lower():
                logger.info(f"[✓] Berhasil masuk via tombol Continue untuk {username}!")
                return True

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


async def fetch_group_details(page, idmg):
    """Mengambil detail profil & rekening bank grup dari endpoint troy/v1/merchant."""
    if not idmg:
        return {}
    try:
        url = f"https://merchant.grab.com/troy/v1/merchant?merchant_group_id={idmg}&isBalanceNeeded=false&currency=IDR"
        res = await page.evaluate('''async (u) => {
            try {
                const r = await fetch(u, {
                    headers: {
                        'accept': 'application/json, text/plain, */*',
                        'requestsource': 'troyPortal'
                    }
                });
                if (!r.ok) return {};
                const json = await r.json();
                return json.data || {};
            } catch (e) {
                return {};
            }
        }''', url)
        if res and isinstance(res, dict):
            bank = res.get("bank_details", {})
            if bank.get("bank_name"):
                logger.info(f"🏦 Berhasil mendapatkan data Bank Grab: {bank.get('bank_name')} a/n {bank.get('account_name')} ({bank.get('account_number')})")
            return res
    except Exception as e:
        logger.debug(f"Error fetching group details: {e}")
    return {}


async def fetch_merchant_list(page):
    """Fetch list of all merchant stores using the search API with pagination."""
    stores = []
    total_expected = 0
    logger.info("Fetching merchant store list...")

    try:
        offset = 0
        limit = 100
        
        while True:
            api_url = f"https://api.grab.com/delvplatformapi/merchant/v1/merchant-group/store/search?offset={offset}&limit={limit}&search=&includeItemsWithoutPhotosCount=true&includeInactive=true&modelType=ALL&asc=true&cityIDs[]=ALL&includeMenuGroupV2ID=false"
            
            page_retries = 3
            fetched_stores = []
            
            for attempt in range(page_retries):
                logger.info(f"Fetching stores offset {offset} (attempt {attempt+1})...")
                response = await page.request.get(api_url, headers=auth_headers)
                if not response.ok:
                    logger.warning(f"Gagal mengambil store list pada offset {offset}: HTTP {response.status}")
                    await asyncio.sleep(1.5)
                    continue
                    
                data = await response.json()
                
                # Ambil totalCount dari response API pada halaman pertama
                if offset == 0:
                    total_expected = (
                        data.get("totalCount") or
                        data.get("total") or
                        data.get("data", {}).get("totalCount") or
                        0
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
                            
                break
            
            if not fetched_stores or not isinstance(fetched_stores, list):
                break
                
            stores.extend(fetched_stores)
            logger.info(f"Berhasil mengambil {len(fetched_stores)} outlet. Total terkumpul: {len(stores)}")
            
            # Jika jumlah data yang didapat kurang dari limit (misal 1, 3, 18), berarti semua sudah terambil dalam 1 request!
            if len(fetched_stores) < limit:
                break
                
            offset += limit
            await asyncio.sleep(1.5)

    except Exception as e:
        logger.error(f"Error fetching merchant list: {e}")

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
        args=[
            "--disable-blink-features=AutomationControlled",
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--ignore-gpu-blocklist",
            "--enable-gpu-rasterization",
            "--enable-zero-copy",
            "--enable-hardware-overlays",
            "--enable-features=VaapiVideoDecoder,CanvasOopRasterization"
        ]
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
            return []

        # Ambil data profil & rekening bank dari troy/v1/merchant
        group_data = await fetch_group_details(page, idmg)
        group_bank = group_data.get("bank_details", {})
        group_bank_name = group_bank.get("bank_name", "")
        group_acc_name = group_bank.get("account_name", "")
        group_acc_no = group_bank.get("account_number", "")

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
            logger.info(f"[*] Menangani mode Single-Store / Profil Langsung untuk {cred['name']}...")
            addr = group_data.get("address", {})
            addr_str = addr.get("AddressLine1", "") if isinstance(addr, dict) else str(addr or "")
            city = addr.get("City", "") if isinstance(addr, dict) else ""
            if city and city not in addr_str:
                addr_str = f"{addr_str}, {city}".strip(", ")
                
            store_name = group_data.get("name") or cred["name"]
            
            ls_data = await page.evaluate('''() => {
                try {
                    const up = JSON.parse(localStorage.getItem('userprofileInfo') || '{}');
                    const links = up?.user_profile?.links || [];
                    const gf_link = links.find(l => l.link_entity_business_line === 'GF') || links[0];
                    return {
                        entity_id: gf_link?.link_entity_id || up?.user_profile?.grab_food_entity_id || '',
                        name: localStorage.getItem('profileInfo') ? JSON.parse(localStorage.getItem('profileInfo')).name : ''
                    };
                } catch(e) { return {}; }
            }''')
            
            merchant_id = ls_data.get("entity_id") or idmg
            if ls_data.get("name"):
                store_name = ls_data.get("name")

            fallback_store = {
                "merchantID": merchant_id,
                "merchantName": store_name,
                "address": addr_str,
                "status": "ACTIVE",
                "bankAccount": group_bank
            }
            stores = [fallback_store]
            logger.info(f"[✓] Berhasil memuat data outlet Single-Store: '{store_name}' (ID: {merchant_id})")

        all_results = []
        
        for store in stores:
            merchant_id = (
                str(store.get("merchantID") or store.get("merchantId") or store.get("id") or "")
            ).strip()
            store_name = store.get("name") or store.get("merchantName") or store.get("storeName") or merchant_id
            status = store.get("status") or store.get("isActive") or ""
            alamat = store.get("address") or store.get("merchantAddress") or ""
            
            # Coba ambil detail bank per store, fallback ke group bank details
            nama_bank = ""
            nama_pemilik = ""
            no_rekening = ""
            
            bank_obj = store.get("bankAccount") or store.get("bank_details")
            if isinstance(bank_obj, dict):
                nama_bank = bank_obj.get("bankName", "") or bank_obj.get("bank_name", "")
                nama_pemilik = bank_obj.get("accountHolderName", "") or bank_obj.get("account_name", "")
                no_rekening = bank_obj.get("accountNumber", "") or bank_obj.get("account_number", "")
            else:
                no_rekening = (
                    store.get("bankAccount") or 
                    store.get("bankAccountNumber") or 
                    store.get("accountNumber") or 
                    ""
                )

            if not nama_bank:
                nama_bank = group_bank_name
            if not nama_pemilik:
                nama_pemilik = group_acc_name
            if not no_rekening:
                no_rekening = group_acc_no

            if not merchant_id:
                continue

            link_menu = f"https://merchant.grab.com/food/menu/{merchant_id}" if merchant_id else ""

            all_results.append({
                "Nama Pemilik": cred.get("owner", ""),
                "Nama Brand": cred.get("brand", ""),
                "Aplikator": "GrabFood",
                "Nama Portal": cred["name"],
                "Group ID": idmg,
                "Nama Listing": store_name,
                "Link": link_menu,
                "Store ID": merchant_id,
                "Status Listing": status,
                "Alamat": alamat,
                "Nama Bank": nama_bank,
                "Nama Pemilik Rekening": nama_pemilik,
                "Nomor Rekening": str(no_rekening) if isinstance(no_rekening, dict) else str(no_rekening),
                "_owner": cred.get("owner", cred["name"])
            })
            
        logger.info(f"Berhasil mengekstrak {len(all_results)} outlet untuk portal {cred['name']}.")
        return all_results

    except Exception as e:
        logger.error(f"Unexpected error for {cred['name']}: {e}")
        return []
    finally:
        await context.close()
        await browser.close()


def get_template_headers():
    """Mengambil header dari template YYYY-MM-DD HH_MM Nama Pemilik.xlsx."""
    import openpyxl
    template_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "YYYY-MM-DD HH_MM Nama Pemilik.xlsx")
    if not os.path.exists(template_path):
        template_path = os.path.join(os.path.dirname(__file__), "YYYY-MM-DD HH_MM Nama Pemilik.xlsx")
    if os.path.exists(template_path):
        try:
            wb = openpyxl.load_workbook(template_path, read_only=True)
            ws = wb['Listing'] if 'Listing' in wb.sheetnames else wb.active
            return [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        except Exception:
            pass
    return ['Nama Pemilik', 'Nama Brand', 'Model', 'Tipe', 'Outlet', 'Nomor HP', 'Aplikator', 'Nama Portal', 'Group ID', 'Nama Listing', 'Link', 'Store ID', 'Status Listing', 'Alamat', 'Nama Bank', 'Nama Pemilik Rekening', 'Nomor Rekening']


def save_formatted_excel(df, file_path):
    """Menyimpan DataFrame ke file Excel menggunakan template YYYY-MM-DD HH_MM Nama Pemilik.xlsx."""
    import openpyxl
    from openpyxl.styles import Font, Alignment
    
    headers = get_template_headers()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Listing"
    
    header_font = Font(name='Calibri', size=11, bold=True)
    for c_idx, h in enumerate(headers, 1):
        if h is not None:
            cell = ws.cell(row=1, column=c_idx, value=h)
            cell.font = header_font
            
    font = Font(name='Calibri', size=11)
    alignment = Alignment(horizontal='left', vertical='center')
    
    for idx, (_, r) in enumerate(df.iterrows(), start=2):
        ws.cell(row=idx, column=1, value=str(r.get('Nama Pemilik') or '') if pd.notna(r.get('Nama Pemilik')) else '') # A: Nama Pemilik
        ws.cell(row=idx, column=2, value=str(r.get('Nama Brand') or '') if pd.notna(r.get('Nama Brand')) else '')   # B: Nama Brand
        ws.cell(row=idx, column=7, value=str(r.get('Aplikator') or '') if pd.notna(r.get('Aplikator')) else 'GrabFood') # G: Aplikator
        ws.cell(row=idx, column=8, value=str(r.get('Nama Portal') or '') if pd.notna(r.get('Nama Portal')) else '') # H: Nama Portal
        ws.cell(row=idx, column=9, value=str(r.get('Group ID') or '') if pd.notna(r.get('Group ID')) else '')       # I: Group ID
        ws.cell(row=idx, column=10, value=str(r.get('Nama Listing') or '') if pd.notna(r.get('Nama Listing')) else '') # J: Nama Listing
        ws.cell(row=idx, column=11, value=str(r.get('Link') or '') if pd.notna(r.get('Link')) else '')             # K: Link
        ws.cell(row=idx, column=12, value=str(r.get('Store ID') or '') if pd.notna(r.get('Store ID')) else '')     # L: Store ID
        ws.cell(row=idx, column=13, value=str(r.get('Status Listing') or '') if pd.notna(r.get('Status Listing')) else '') # M: Status Listing
        ws.cell(row=idx, column=14, value=str(r.get('Alamat') or '') if pd.notna(r.get('Alamat')) else '')         # N: Alamat
        ws.cell(row=idx, column=15, value=str(r.get('Nama Bank') or '') if pd.notna(r.get('Nama Bank')) else '')   # O: Nama Bank
        ws.cell(row=idx, column=16, value=str(r.get('Nama Pemilik Rekening') or '') if pd.notna(r.get('Nama Pemilik Rekening')) else '') # P: Nama Pemilik Rekening
        
        c_rek = ws.cell(row=idx, column=17, value=str(r.get('Nomor Rekening') or '') if pd.notna(r.get('Nomor Rekening')) else '') # Q: Nomor Rekening
        c_rek.number_format = '@'

        for c in range(1, 18):
            cell = ws.cell(row=idx, column=c)
            cell.font = font
            cell.alignment = alignment

    # Auto adjust column widths
    for col_idx in range(1, 18):
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        max_len = max((len(str(ws.cell(row=r, column=col_idx).value or '')) for r in range(1, len(df) + 2)), default=0)
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 60)

    wb.save(file_path)


async def main():
    import datetime
    
    parser = argparse.ArgumentParser(description="Grab Merchant Scraper")
    parser.add_argument("--type", type=str, choices=["agency", "vb"], help="Pilih tipe sumber: 'agency' atau 'vb'")
    parser.add_argument("--agency", action="store_true", help="Gunakan sumber kredensial Agency")
    parser.add_argument("--vb", action="store_true", help="Gunakan sumber kredensial VB (Virtual Brand)")
    parser.add_argument("--vb-url", type=str, help="URL Google Sheet / CSV khusus untuk VB")
    parser.add_argument("--url", type=str, help="URL Google Sheet / CSV custom")
    parser.add_argument("--outlet", type=str, help="Specify the outlet name to run (e.g., F1)")
    parser.add_argument("--all", action="store_true", help="Run for all portals without prompt")
    parser.add_argument("--fresh", action="store_true", help="Start fresh run and ignore previous progress checkpoint")
    args = parser.parse_args()

    source_type = "agency"
    custom_url = args.url or args.vb_url

    if args.vb or args.type == "vb":
        source_type = "vb"
    elif args.agency or args.type == "agency":
        source_type = "agency"
    elif not args.outlet and not args.all:
        print("\n" + "=" * 54)
        print("  PILIH SUMBER KREDENSIAL GRABFOOD")
        print("=" * 54)
        print("  [1] Agency (Master Agency Google Sheet)")
        print("  [2] VB     (Virtual Brand - Dokumen Lain)")
        print("=" * 54)
        pilihan = input("Pilih [1/2] (Default: 1 - Agency): ").strip()
        if pilihan == "2" or pilihan.lower() == "vb":
            source_type = "vb"
        else:
            source_type = "agency"

    portals = get_credentials_from_sheet(source_type=source_type, custom_url=custom_url)
    target_credentials = []

    if args.outlet:
        target_credentials = [c for c in portals if c["name"] == args.outlet]
        if not target_credentials:
            logger.error(f"Outlet '{args.outlet}' not found in credentials.")
            return
    elif args.all:
        target_credentials = portals
    else:
        if portals:
            print(f"\nDaftar Portal GrabFood [{source_type.upper()}]:")
            print("  [0] Pilih Semua (All)")
            for idx, p in enumerate(portals):
                print(f"  [{idx+1}] {p['name']} ({p['username']}) - Owner: {p.get('owner', '-')}")
            
            choice = input(f"\nPilih portal (contoh: 1, 2-3, all): ").strip().lower()
            selected_indices = parse_selection(choice, len(portals))
            for idx in selected_indices:
                target_credentials.append(portals[idx])

    if not target_credentials:
        logger.error("Tidak ada portal yang dipilih.")
        return

    output_dir = os.path.join(os.path.dirname(__file__), "sessions")
    os.makedirs(output_dir, exist_ok=True)
    progress_file = os.path.join(output_dir, f".grab_progress_{source_type}.json")
    if source_type == "agency" and not os.path.exists(progress_file) and os.path.exists(os.path.join(output_dir, ".grab_progress.json")):
        progress_file = os.path.join(output_dir, ".grab_progress.json")

    # ── Checkpoint / Resume Management ──────────────────────────────
    completed_portal_names = set()
    all_collected_stores = []
    
    if args.fresh and os.path.exists(progress_file):
        try:
            os.remove(progress_file)
            logger.info("🗑️ Checkpoint progress sebelumnya dihapus (--fresh aktif).")
        except Exception:
            pass

    if not args.fresh and os.path.exists(progress_file):
        try:
            with open(progress_file, "r", encoding="utf-8") as f:
                checkpoint_data = json.load(f)
                completed_portal_names = set(checkpoint_data.get("completed_portals", []))
                raw_stores = checkpoint_data.get("stores", [])
                all_collected_stores = [s for s in raw_stores if isinstance(s, dict)]
                logger.info(f"🔄 [RESUME AKTIF] Memuat progress checkpoint: {len(completed_portal_names)} portal sudah selesai ditarik sebelumnya.")
        except Exception as ex:
            logger.warning(f"Gagal membaca checkpoint progress: {ex}")

    logger.info("Starting Grab Merchant Scraper")
    logger.info(f"Total portal yang ditargetkan: {len(target_credentials)}")

    async with async_playwright() as playwright:
        for idx, cred in enumerate(target_credentials, 1):
            if cred["name"] in completed_portal_names:
                logger.info(f"⏩ [{idx}/{len(target_credentials)}] Portal '{cred['name']}' sudah selesai sebelumnya (Dilewati).")
                continue

            logger.info(f"\n▶️ [{idx}/{len(target_credentials)}] Memproses portal: {cred['name']} (Owner: {cred.get('owner', '-')})...")
            stores_res = await run_scraper_for_credential(playwright, cred)
            
            retry_count = 0
            while not stores_res and retry_count < 2:
                retry_count += 1
                logger.warning(f"[!] Portal {cred['name']} gagal diproses/login. Melakukan retry ke-{retry_count} dari 2...")
                await asyncio.sleep(5)
                stores_res = await run_scraper_for_credential(playwright, cred, force_fresh=True)

            if stores_res and isinstance(stores_res, list):
                valid_stores = [s for s in stores_res if isinstance(s, dict)]
                if valid_stores:
                    all_collected_stores.extend(valid_stores)
                    completed_portal_names.add(cred["name"])
                    
                    # Simpan checkpoint progress seketika (Real-Time Auto-Save)
                    try:
                        with open(progress_file, "w", encoding="utf-8") as f:
                            json.dump({
                                "completed_portals": list(completed_portal_names),
                                "stores": all_collected_stores,
                                "last_update": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            }, f, indent=2)
                    except Exception as ex:
                        logger.debug(f"Gagal menyimpan checkpoint progress: {ex}")

    logger.info("\n[✓] Proses scraping selesai.")

    # Filter ketat hanya dict objek
    all_collected_stores = [s for s in all_collected_stores if isinstance(s, dict)]

    if not all_collected_stores:
        logger.warning("Tidak ada data outlet yang berhasil dikumpulkan.")
        return

    df_all = pd.DataFrame(all_collected_stores)
    if "Store ID" in df_all.columns:
        df_all.drop_duplicates(subset=["Store ID"], keep="first", inplace=True)
    
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H_%M")

    # 1. Simpan per Owner (Nama Pemilik) sesuai template baru: YYYY-MM-DD HH_MM Nama Pemilik.xlsx
    owners = df_all["_owner"].dropna().unique()
    for owner in owners:
        df_owner = df_all[df_all["_owner"] == owner].drop(columns=["_owner"])
        safe_owner = "".join(c for c in str(owner) if c.isalnum() or c in (' ', '_', '-')).strip()
        if not safe_owner:
            safe_owner = "Unknown"
        owner_file = os.path.join(output_dir, f"{timestamp} {safe_owner}.xlsx")
        save_formatted_excel(df_owner, owner_file)
        logger.info(f"[✓] File Template Baru tersimpan: {owner_file} ({len(df_owner)} baris)")

    # 2. Simpan Master Gabungan jika lebih dari 1 owner atau opsi --all
    if len(owners) > 1 or args.all:
        master_df = df_all.drop(columns=["_owner"])
        master_file = os.path.join(output_dir, f"{timestamp} Master.xlsx")
        save_formatted_excel(master_df, master_file)
        logger.info(f"[✓] File Master tersimpan: {master_file} ({len(master_df)} baris)")

    # Bersihkan checkpoint jika semua proses selesai sukses
    if os.path.exists(progress_file):
        try:
            os.remove(progress_file)
            logger.info("🧹 Checkpoint progress sementara dibersihkan (Semua selesai).")
        except Exception:
            pass

    logger.info("\n🎉 Seluruh penarikan selesai dan tersimpan dengan template baru!")


if __name__ == "__main__":
    asyncio.run(main())
