import asyncio
import json
import logging
import os
import argparse
import pandas as pd
from playwright.async_api import async_playwright

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger("ShopeeScraper")

CREDENTIALS = [
    {
        "name": "F",
        "phone": "+6285183151531",
        "username": "superfoodapp",
        "password": "Nesia@123@",
        "output": "data/F.xlsx"
    },
    {
        "name": "W",
        "phone": "+6285182003796",
        "username": "wonderfoodapp",
        "password": "Wonder@123@@",
        "output": "data/W.xlsx"
    },
    {
        "name": "L",
        "phone": "+6285182003783",
        "username": "lokarasaapp",
        "password": "Lokar@123@",
        "output": "data/L.xlsx"
    },
    {
        "name": "D",
        "phone": "+6285136517300",
        "username": "doeatapp",
        "password": "Doeat@123@",
        "output": "data/D.xlsx"
    },
]

BASE_URL = "https://partner.shopee.co.id"
LOGIN_URL = f"{BASE_URL}/account/login"
DASHBOARD_URL = f"{BASE_URL}/food/dashboard"

auth_headers = {}


async def perform_login(page, credential):
    logger.info(f"Checking session for {credential['name']}...")
    try:
        # Cek apakah sudah login dengan mengakses dashboard
        await page.goto(DASHBOARD_URL, wait_until="networkidle", timeout=60000)
        await page.wait_for_timeout(3000)
        
        if "login" not in page.url:
            logger.info(f"✅ Sudah login menggunakan session untuk {credential['name']}!")
            return True
            
        logger.info(f"Sesi tidak valid/belum login. Mengisi form login untuk {credential['name']}...")
        
        # Try phone login
        phone_input = page.locator('input[type="tel"], input[placeholder*="Nomor" i], input[placeholder*="phone" i]').first
        if await phone_input.is_visible(timeout=5000):
            await phone_input.fill(credential.get("phone", ""))
        else:
            username_input = page.locator('input[type="text"]').first
            await username_input.fill(credential.get("username", ""))

        password_input = page.locator('input[type="password"]').first
        await password_input.fill(credential["password"])

        submit_btn = page.locator('button[type="submit"]').first
        await submit_btn.click()
        
        logger.info("Menunggu penyelesaian login (harap masukkan OTP secara manual jika diminta)...")
        # Beri waktu lebih lama agar user bisa memasukkan OTP jika diperlukan
        try:
            await page.wait_for_url(lambda url: "login" not in url, timeout=60000)
            logger.info(f"✅ Login sukses untuk {credential['name']}")
            return True
        except:
            logger.error(f"Waktu tunggu login habis untuk {credential['name']}")
            return False

    except Exception as e:
        logger.error(f"Login failed for {credential['name']}: {e}")
        return False


def handle_request(request):
    headers = request.headers
    if "authorization" in headers or "x-shopee" in headers:
        auth_headers.update({
            k: v for k, v in headers.items()
            if k.lower() in ["authorization", "x-shopee-token", "cookie", "content-type", "accept"]
        })


async def fetch_outlets(page):
    """Fetch list of all Shopee Food outlets."""
    outlets = []
    logger.info("Fetching Shopee Food outlet list...")

    try:
        async def handle_response(response):
            if "outlet" in response.url.lower() or "merchant" in response.url.lower():
                if response.status == 200:
                    try:
                        data = await response.json()
                        items = (
                            data.get("data", {}).get("outlets") or
                            data.get("data", {}).get("merchants") or
                            data.get("outlets") or
                            data.get("merchants") or
                            []
                        )
                        if items:
                            outlets.extend(items)
                            logger.info(f"Found {len(items)} outlets")
                    except Exception:
                        pass

        page.on("response", handle_response)
        await page.goto(DASHBOARD_URL, wait_until="networkidle", timeout=60000)
        await page.wait_for_timeout(3000)

    except Exception as e:
        logger.error(f"Error fetching outlets: {e}")

    return outlets


async def fetch_menu_items(page, outlet_id, outlet_name):
    """Fetch menu items for a specific outlet."""
    items = []
    logger.info(f"Fetching menu for outlet: {outlet_name} ({outlet_id})")

    try:
        menu_url = f"{BASE_URL}/food/menu?outlet_id={outlet_id}"
        await page.goto(menu_url, wait_until="domcontentloaded", timeout=60000)

        async def handle_menu_response(response):
            if "menu" in response.url.lower() and response.status == 200:
                try:
                    data = await response.json()
                    menu_items = (
                        data.get("data", {}).get("items") or
                        data.get("items") or
                        []
                    )
                    if menu_items:
                        items.extend(menu_items)
                except Exception:
                    pass

        page.on("response", handle_menu_response)
        await page.wait_for_timeout(3000)

    except Exception as e:
        logger.error(f"Error fetching menu for {outlet_name}: {e}")

    return items


def process_items(items, outlet_id, outlet_name):
    """Process raw menu items into structured rows."""
    rows = []
    jumlah_foto = 0

    for item in items:
        has_photo = bool(
            item.get("image") or item.get("images") or
            item.get("photo") or item.get("photos") or
            item.get("thumbnail")
        )
        if has_photo:
            jumlah_foto += 1

        rows.append({
            "Outlet ID": outlet_id,
            "Nama Outlet": outlet_name,
            "Item ID": item.get("id") or item.get("item_id") or "",
            "Nama Item": item.get("name") or item.get("item_name") or "",
            "Kategori": item.get("category") or item.get("category_name") or "",
            "Harga": item.get("price") or item.get("original_price") or 0,
            "Tersedia": item.get("available", True),
            "Ada Foto": has_photo,
            "Deskripsi": item.get("description") or "",
        })

    logger.info(f"  {len(rows)} items ({jumlah_foto} ada foto) untuk {outlet_name}")
    return rows


async def run_scraper_for_credential(playwright, cred):
    global auth_headers
    auth_headers = {}

    logger.info(f"\n{'='*50}")
    logger.info(f"Starting Shopee scraper for: {cred['name']}")
    logger.info(f"{'='*50}")

    user_data_dir = os.path.join(os.path.dirname(__file__), "data", "chrome_profiles", f"portal_{cred['name'].lower()}")
    os.makedirs(user_data_dir, exist_ok=True)
    
    context = await playwright.chromium.launch_persistent_context(
        user_data_dir=user_data_dir,
        headless=False,
        args=["--disable-blink-features=AutomationControlled"],
        viewport={"width": 1280, "height": 800},
        user_agent="Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    
    page = context.pages[0] if context.pages else await context.new_page()
    page.on("request", handle_request)

    try:
        logged_in = await perform_login(page, cred)
        if not logged_in:
            logger.error(f"Skipping {cred['name']} due to login failure.")
            return

        await page.wait_for_timeout(2000)
        outlets = await fetch_outlets(page)
        logger.info(f"Total outlets: {len(outlets)}")

        all_results = []

        for outlet in outlets:
            outlet_id = str(outlet.get("id") or outlet.get("outlet_id") or "")
            outlet_name = outlet.get("name") or outlet.get("outlet_name") or outlet_id

            if not outlet_id:
                continue

            items = await fetch_menu_items(page, outlet_id, outlet_name)

            if items:
                rows = process_items(items, outlet_id, outlet_name)
                all_results.extend(rows)
            else:
                all_results.append({
                    "Outlet ID": outlet_id,
                    "Nama Outlet": outlet_name,
                    "Item ID": "",
                    "Nama Item": "",
                    "Kategori": "",
                    "Harga": 0,
                    "Tersedia": True,
                    "Ada Foto": False,
                    "Deskripsi": "",
                })

        if all_results:
            output_dir = os.path.join(os.path.dirname(__file__), os.path.dirname(cred["output"]))
            os.makedirs(output_dir, exist_ok=True)

            output_path = os.path.join(os.path.dirname(__file__), cred["output"])
            df = pd.DataFrame(all_results)
            df.to_excel(output_path, index=False)
            logger.info(f"[✓] Saved {len(df)} rows to {output_path}")
        else:
            logger.warning(f"No results for {cred['name']}")

    except Exception as e:
        logger.error(f"Error for {cred['name']}: {e}")
    finally:
        await context.close()


async def main():
    parser = argparse.ArgumentParser(description="Shopee Food Merchant Scraper")
    parser.add_argument("--outlet", type=str, help="Specify the outlet name to run (e.g., F)")
    args = parser.parse_args()

    target_credentials = CREDENTIALS
    if args.outlet:
        target_credentials = [c for c in CREDENTIALS if c["name"] == args.outlet]
        if not target_credentials:
            logger.error(f"Outlet '{args.outlet}' not found in credentials.")
            return

    logger.info("Starting Shopee Food Merchant Scraper")
    logger.info(f"Processing {len(target_credentials)} credential set(s)")

    async with async_playwright() as playwright:
        for cred in target_credentials:
            await run_scraper_for_credential(playwright, cred)

    logger.info("All done!")


if __name__ == "__main__":
    asyncio.run(main())
