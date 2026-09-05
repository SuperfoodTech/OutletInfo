import os
import time
import json
import pandas as pd
from datetime import datetime, timedelta
from dotenv import load_dotenv

from core.browser import get_session, return_to_selector, refresh_tokens, auto_switch_merchant
from core.client import ShopeeClient
from core.logger import get_logger
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

# Load environment variables
load_dotenv()
log = get_logger("omzet_pipeline")

def subtract_months(dt, months):
    """Helper to subtract calendar months."""
    for _ in range(months):
        dt = (dt - timedelta(days=1)).replace(day=1)
    return dt

def download_file(url, filename, cookies=None):
    """Downloads a file from a URL with optional cookies."""
    import requests
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36"
    }
    try:
        response = requests.get(url, stream=True, cookies=cookies, headers=headers)
        response.raise_for_status()
        with open(filename, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        return True
    except Exception as e:
        log.error(f"Failed to download {filename}: {e}")
        return False

def format_rupiah(amount):
    """Formats a number as Rupiah."""
    try:
        return f"Rp {int(float(amount)):,}".replace(",", ".")
    except:
        return str(amount)

def run_pipeline():
    print("\n" + "=" * 60)
    print("  Shopee Omzet Analysis Pipeline (Batch Parallel)")
    print("=" * 60)

    phone    = os.getenv("SHOPEE_PHONE", "").strip()
    username = os.getenv("SHOPEE_USERNAME", "").strip()
    password = os.getenv("SHOPEE_PASSWORD", "").strip()
    # Load headless setting from config.json walk-up
    headless = True
    try:
        from pathlib import Path
        import json
        for parent in Path(__file__).resolve().parents:
            config_file = parent / "config.json"
            if config_file.exists():
                with open(config_file, "r") as f:
                    headless = json.load(f).get("headless_shopee", True)
                break
    except Exception:
        pass

    # ── 1. Determine Merchants to Process ───────────────────────────────
    merchants_env = os.getenv("SHOPEE_MERCHANTS", "").strip()
    target_merchants = []
    
    if merchants_env and merchants_env.upper() != "ALL":
        target_merchants = [m.strip() for m in merchants_env.split(",") if m.strip()]
        log.info(f"📋 Found {len(target_merchants)} merchants to process from .env")
    else:
        # AUTOMATIC DISCOVERY: Load all merchants from API/response.json
        try:
            from pathlib import Path
            api_response_path = Path(__file__).resolve().parent / "API" / "response.json"
            with open(api_response_path, "r") as f:
                data = json.load(f)
                merchant_list = data.get("data", {}).get("selectMerchant", {}).get("merchantList", [])
                target_merchants = [m["merchantName"] for m in merchant_list]
                log.info(f"🤖 Auto-detected {len(target_merchants)} merchants from API response.")
        except Exception as e:
            log.warning(f"⚠️ Failed to auto-detect merchants: {e}")
            target_merchants = []

    if not target_merchants:
        log.error("❌ No merchants to process. Aborting.")
        return

    # ── 2. Phase 1: Rapid Trigger (Trigger exports for all) ────────────
    log.info(f"🚀 PHASE 1: Triggering Exports for {len(target_merchants)} merchants...")
    
    # Initialize session
    session_data = get_session(username=username or None, password=password or None, phone=phone or None, 
                               headless=headless, close_browser=False, target_name=target_merchants[0])
    if not session_data: return
    driver = session_data.get("driver")

    merchants_context = {} # Store tokens/ids for each merchant
    start_time_all = int(time.time())

    for i, merchant_name in enumerate(target_merchants):
        log.info(f"  [{i+1}/{len(target_merchants)}] Triggering: {merchant_name}")
        
        # Switch if not already there
        if i > 0:
            if not auto_switch_merchant(driver, merchant_name):
                log.warning(f"  ❌ Skipping {merchant_name} due to switch failure.")
                continue
            time.sleep(3) # Wait for cookies to sync
        
        # Get tokens and VERIFY ID
        session = refresh_tokens(driver)
        active_id = str(session.get("shopee_tob_entity_id") or "")
        
        # Double check if the ID actually changed from previous
        if i > 0 and active_id == merchants_context.get(target_merchants[i-1], {}).get("entity_id"):
             log.warning("  ⚠️ ID hasn't changed yet. Retrying token refresh...")
             time.sleep(3)
             session = refresh_tokens(driver)
             active_id = str(session.get("shopee_tob_entity_id") or "")

        log.info(f"  📍 Confirmed ID for {merchant_name}: {active_id}")
        
        # Store context for polling
        merchants_context[merchant_name] = {
            "entity_id": active_id,
            "tob_token": session["shopee_tob_token"],
            "cookies": session.get("extra_cookies", {}),
            "start_trigger_time": int(time.time())
        }

        # Initialize client and trigger
        client = ShopeeClient(tob_token=session["shopee_tob_token"], entity_id=active_id, extra_cookies=session.get("extra_cookies", {}))
        
        # Define ranges
        now = datetime.now()
        month_starts = [subtract_months(now.replace(day=1, hour=0, minute=0, second=0, microsecond=0), j) for j in range(4)]
        ranges = []
        for j in range(3):
            ranges.append({"start": int(month_starts[j+1].timestamp()), "end": int(month_starts[j].timestamp()) - 1, "label": month_starts[j+1].strftime("%b %Y")})
        
        merchants_context[merchant_name]["ranges"] = ranges
        merchants_context[merchant_name]["downloaded"] = []

        # Trigger
        for r in ranges:
            client.export_transaction_report(merchant_ids=[active_id], start_time=r["start"], end_time=r["end"])
            time.sleep(1)

    # ── 3. Phase 2: Global Polling & Download ──────────────────────────
    log.info(f"⏳ PHASE 2: Global Polling for all reports...")
    os.makedirs("data/reports/merchant", exist_ok=True)
    
    total_expected = len(merchants_context) * 3
    download_count = 0
    start_poll = time.time()
    
    poll_iteration = 0
    while download_count < total_expected and (time.time() - start_poll) < 1200:
        found_new = False
        poll_iteration += 1
        
        for m_name, ctx in merchants_context.items():
            if len(ctx["downloaded"]) >= 3: continue
            
            client = ShopeeClient(tob_token=ctx["tob_token"], entity_id=ctx["entity_id"], extra_cookies=ctx["cookies"])
            reports = client.get_report_list()
            
            for rep in reports:
                # Match: status ready (2 or 3), has download URL, created after our trigger
                if rep.get("status") in [2, 3] and rep.get("download_url"):
                    if rep.get("create_time", 0) and rep["create_time"] >= ctx["start_trigger_time"]:
                        # Use report name for file naming (e.g. "Transactions_01022026_28022026_ShopeeFood.xlsx")
                        report_name = rep.get("name", f"report_{rep.get('id')}.xlsx")
                        target_path = os.path.join("data/reports/merchant", f"{m_name.replace(' ', '_')}_{report_name}")
                        
                        if target_path not in [d[0] for d in ctx["downloaded"]]:
                            if download_file(rep.get("download_url"), target_path):
                                log.info(f"  ✅ [{m_name}] Downloaded: {report_name} ({len(ctx['downloaded'])+1}/3)")
                                ctx["downloaded"].append((target_path, report_name))
                                download_count += 1
                                found_new = True
            
            # Log progress every 3 iterations (~30 seconds)
            if not found_new and poll_iteration % 3 == 0:
                 log.info(f"  ⏳ Still waiting for {m_name}... ({len(ctx['downloaded'])}/3 ready)")
            
        if download_count < total_expected:
            time.sleep(10)

    # ── Summary ──────────────────────────────────────────────────────────
    log.info("📋 DONE! Download Summary:")
    for m_name, ctx in merchants_context.items():
        log.info(f"  🏪 {m_name}: {len(ctx['downloaded'])}/3 files")
        for fpath, label in ctx["downloaded"]:
            log.info(f"     📄 {fpath}")

    driver.quit()


if __name__ == "__main__":
    run_pipeline()
