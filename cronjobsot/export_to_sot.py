#!/usr/bin/env python3
"""
Export to SOT (Google Spreadsheet Sync)
=======================================
File mandiri untuk mengekspor dan melakukan upsert data listing outlet ke
Google Spreadsheet tab 'SOT' (Spreadsheet ID: 15_Xx5ixOcxy0L90U_BrFHfVnAgjIZpiiGBEK4Ce4SyI, GID: 2135653103).

Fitur:
- Membaca data sumber dari master scraping multi-aplikator (GoFood, GrabFood, ShopeeFood)
  dan/atau data master akun dari tab 'DBR' (via URL CSV Publish to Web atau Apps Script).
- Menerapkan format 37 kolom template standar.
- Menambahkan kolom ke-38 'Terakhir Diperbaharui' secara otomatis melalui Apps Script.
- Logika upsert: menimpa baris outlet yang sama (berdasarkan Aplikator + Store ID)
  dan meng-append outlet baru.
- Dapat dijalankan secara mandiri (CLI) atau diimpor oleh modul lain.

Penggunaan CLI:
    python cronjobsot/export_to_sot.py                  # Ekspor data master gabungan ke SOT
    python cronjobsot/export_to_sot.py --from-dbr        # Ekspor data master dari Tab DBR ke SOT
    python cronjobsot/export_to_sot.py --file file.xlsx  # Ekspor file Excel tertentu ke SOT
    python cronjobsot/export_to_sot.py --owner "Nama"    # Filter hanya owner tertentu
"""

import os
import sys
import io
import time
import glob
import json
import argparse
import urllib.request
from pathlib import Path
from dotenv import load_dotenv

import pandas as pd
import requests

# Path Proyek
CRON_DIR = Path(__file__).resolve().parent
BASE_DIR = CRON_DIR.parent
CACHE_DIR = CRON_DIR / "cache"
CACHE_DIR.mkdir(parents=True, exist_ok=True)

load_dotenv(BASE_DIR / ".env")

from sheet_syncer import sync_outlets_to_google_sheet, TARGET_SPREADSHEET_ID, TARGET_SHEET_NAME, TARGET_SHEET_GID

GOOGLE_SHEET_DBR_URL = os.getenv("GOOGLE_SHEET_DBR_URL", "")
SOT_APP_SCRIPT_URL = os.getenv("SOT_APP_SCRIPT_URL", os.getenv("APP_SCRIPT_URL", ""))
GOOGLE_SHEET_VERCEL_URL = os.getenv(
    "GOOGLE_SHEET_VERCEL_URL",
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vTprbPPf_J5gAVL3PYeHbbdl5ZXQvb17HY2lJGPI2xg13Ly3AGT8eYHLYmU_m1NdtkBVg-qUGv1BoEE/pub?output=csv"
)

GOFOOD_DIR = BASE_DIR / "GOFOOD"
GRAB_DIR = BASE_DIR / "GRAB"
SHOPEE_DIR = BASE_DIR / "SHOPEE"


def load_dbr_master_data(force_live: bool = False) -> pd.DataFrame:
    """
    Memuat data master owner & kredensial dari Google Spreadsheet Tab 'DBR'.
    Prioritas sumber:
    1. GOOGLE_SHEET_DBR_URL (Publish to web CSV tab DBR jika disetel)
    2. Google Apps Script Web App (action: 'get_dbr' via SOT_APP_SCRIPT_URL)
    3. Cache lokal cronjobsot/cache/dbr_sheet_cache.csv
    4. Fallback Google Sheet Vercel CSV (GOOGLE_SHEET_VERCEL_URL) & vercel_sheet_cache.csv
    """
    dbr_cache_file = CACHE_DIR / "dbr_sheet_cache.csv"
    vercel_cache_file = BASE_DIR / "cache" / "vercel_sheet_cache.csv"
    csv_text = ""
    df = None

    # 1. Coba fetch dari GOOGLE_SHEET_DBR_URL jika ada
    if GOOGLE_SHEET_DBR_URL:
        try:
            url = GOOGLE_SHEET_DBR_URL.strip()
            if "/pubhtml" in url:
                url = url.replace("/pubhtml", "/pub?output=csv")
            sep = "&" if "?" in url else "?"
            url_busted = f"{url}{sep}_cb={int(time.time())}"
            req = urllib.request.Request(
                url_busted,
                headers={
                    'User-Agent': 'Mozilla/5.0',
                    'Cache-Control': 'no-cache, no-store, must-revalidate',
                    'Pragma': 'no-cache'
                }
            )
            with urllib.request.urlopen(req, timeout=12) as resp:
                csv_text = resp.read().decode('utf-8', errors='replace')
            if csv_text.strip():
                dbr_cache_file.write_text(csv_text, encoding="utf-8")
                print(f"   ✓ [DBR] Berhasil memuat data live tab DBR dari Google Sheet ({len(csv_text)} bytes).")
                df = pd.read_csv(io.StringIO(csv_text))
        except Exception as e:
            print(f"   ⚠️ Fetch live GOOGLE_SHEET_DBR_URL ({e}). Mencoba jalur Apps Script/Cache.")

    # 2. Coba fetch via Google Apps Script Web App jika df masih kosong
    if df is None and SOT_APP_SCRIPT_URL:
        try:
            resp = requests.post(SOT_APP_SCRIPT_URL, json={"action": "get_dbr"}, timeout=15)
            if resp.status_code == 200:
                res_json = resp.json()
                if res_json.get("status") == "success" and res_json.get("data"):
                    df = pd.DataFrame(res_json["data"])
                    df.to_csv(dbr_cache_file, index=False)
                    print(f"   ✓ [DBR] Berhasil memuat {len(df)} baris tab DBR via Apps Script Web App.")
        except Exception as e:
            print(f"   ℹ️ Fetch tab DBR via Apps Script ({e}).")

    # 3. Fallback ke cache lokal DBR
    if df is None and dbr_cache_file.exists():
        try:
            df = pd.read_csv(dbr_cache_file)
            print(f"   ℹ️ [DBR] Menggunakan cache lokal ({dbr_cache_file.name}).")
        except Exception:
            pass

    # 4. Fallback ke Vercel jika DBR belum terhubung
    if df is None:
        try:
            url = GOOGLE_SHEET_VERCEL_URL
            sep = "&" if "?" in url else "?"
            url_busted = f"{url}{sep}_cb={int(time.time())}"
            req = urllib.request.Request(
                url_busted,
                headers={'User-Agent': 'Mozilla/5.0'}
            )
            with urllib.request.urlopen(req, timeout=12) as resp:
                csv_text = resp.read().decode('utf-8', errors='replace')
            if csv_text.strip():
                df = pd.read_csv(io.StringIO(csv_text))
                print(f"   ✓ [VERCEL] Menggunakan data Vercel Sheet fallback ({len(csv_text)} bytes).")
        except Exception:
            pass

    if df is None and vercel_cache_file.exists():
        try:
            df = pd.read_csv(vercel_cache_file)
            print(f"   ℹ️ [VERCEL] Menggunakan cache lokal fallback ({vercel_cache_file.name}).")
        except Exception:
            pass

    if df is None or df.empty:
        return pd.DataFrame()

    df.columns = [str(c).strip() for c in df.columns]

    # Normalisasi kolom dasar
    if "Owner" in df.columns and "Nama Pemilik" not in df.columns:
        df["Nama Pemilik"] = df["Owner"].astype(str).str.strip()
    if "Nama Outlet" in df.columns and "Nama Brand" not in df.columns:
        df["Nama Brand"] = df["Nama Outlet"].astype(str).str.strip()
        df["Nama Listing"] = df["Nama Brand"]

    def clean_app(val):
        s = str(val).strip().lower()
        if "gofood" in s or "go" in s:
            return "GoFood"
        elif "grab" in s:
            return "GrabFood"
        elif "shopee" in s:
            return "ShopeeFood"
        return str(val).strip()

    if "Aplikasi" in df.columns and "Aplikator" not in df.columns:
        df["Aplikator"] = df["Aplikasi"].apply(clean_app)

    if "Status Listing" not in df.columns:
        df["Status Listing"] = "LIVE"

    return df


def load_all_local_masters() -> pd.DataFrame:
    """
    Memuat seluruh data master listing lokal hasil penarikan multi-aplikator:
    GoFood, GrabFood, dan ShopeeFood.
    """
    records = []

    # 1. GoFood Master
    gofood_master = GOFOOD_DIR / "master" / "0master.xlsx"
    if not gofood_master.exists():
        masters = sorted(glob.glob(str(GOFOOD_DIR / "master" / "*_master.xlsx")))
        if masters:
            gofood_master = Path(masters[-1])
    if gofood_master.exists():
        try:
            df_go = pd.read_excel(str(gofood_master), sheet_name="Listing")
            df_go["Aplikator"] = "GoFood"
            records.append(df_go)
            print(f"   ✓ Memuat GoFood master: {len(df_go)} outlet")
        except Exception as e:
            print(f"   ⚠️ Gagal membaca GoFood master: {e}")

    # 2. GrabFood Master
    grab_master = GRAB_DIR / "master" / "grab_master.xlsx"
    if grab_master.exists():
        try:
            df_gr = pd.read_excel(str(grab_master), sheet_name="Listing")
            df_gr["Aplikator"] = "GrabFood"
            records.append(df_gr)
            print(f"   ✓ Memuat GrabFood master: {len(df_gr)} outlet")
        except Exception as e:
            print(f"   ⚠️ Gagal membaca GrabFood master: {e}")

    # 3. Shopee Master
    shopee_master = SHOPEE_DIR / "master" / "shopee_master.xlsx"
    if shopee_master.exists():
        try:
            df_sh = pd.read_excel(str(shopee_master), sheet_name="Listing")
            df_sh["Aplikator"] = "ShopeeFood"
            records.append(df_sh)
            print(f"   ✓ Memuat Shopee master: {len(df_sh)} outlet")
        except Exception as e:
            print(f"   ⚠️ Gagal membaca Shopee master: {e}")

    if not records:
        return pd.DataFrame()

    combined = pd.concat(records, ignore_index=True)

    # Deduplikasi berdasarkan Aplikator dan Store ID
    if "Store ID" in combined.columns and "Aplikator" in combined.columns:
        combined = combined.drop_duplicates(subset=["Aplikator", "Store ID"], keep="last")

    return combined


def export_dataframe_to_sot(df: pd.DataFrame, owner_filter: str = None) -> tuple[bool, dict]:
    """Mengirimkan DataFrame outlet ke tab 'SOT' Google Spreadsheet via sheet_syncer."""
    if df is None or df.empty:
        print("⚠️ Data kosong, tidak ada yang diekspor ke SOT.")
        return False, {"error": "Data kosong"}

    if owner_filter:
        if "Nama Pemilik" in df.columns:
            df = df[df["Nama Pemilik"].astype(str).str.strip().str.lower() == owner_filter.strip().lower()]
            print(f"[*] Filter data untuk owner '{owner_filter}': tersisa {len(df)} baris.")

    if df.empty:
        print(f"⚠️ Tidak ada data untuk owner '{owner_filter}'.")
        return False, {"error": f"Owner '{owner_filter}' tidak ditemukan"}

    print(f"[*] Memulai ekspor {len(df)} baris ke Tab 'SOT' (GID: {TARGET_SHEET_GID})...")
    ok, res = sync_outlets_to_google_sheet(df, app_script_url=SOT_APP_SCRIPT_URL)
    return ok, res


def main():
    parser = argparse.ArgumentParser(description="Export listing data to Google Spreadsheet Tab 'SOT'")
    parser.add_argument("--from-dbr", action="store_true", help="Ambil data acuan dari Tab 'DBR'")
    parser.add_argument("--file", type=str, default=None, help="File Excel/CSV spesifik yang akan diekspor ke SOT")
    parser.add_argument("--owner", type=str, default=None, help="Filter spesifik owner yang diekspor")

    args = parser.parse_args()

    print("\n" + "=" * 65)
    print("  🚀 EXPORT TO GOOGLE SPREADSHEET TAB 'SOT'")
    print(f"  Target: {TARGET_SPREADSHEET_ID} (Tab: {TARGET_SHEET_NAME})")
    print("=" * 65)

    if args.file:
        file_path = Path(args.file)
        if not file_path.exists():
            print(f"❌ File tidak ditemukan: {file_path}")
            sys.exit(1)
        print(f"[*] Membaca file: {file_path.name}...")
        if file_path.suffix.lower() in ('.xlsx', '.xls'):
            df = pd.read_excel(str(file_path))
        else:
            df = pd.read_csv(str(file_path))
    elif args.from_dbr:
        print("[*] Memuat data dari Tab 'DBR'...")
        df = load_dbr_master_data()
    else:
        print("[*] Memuat data gabungan master lokal (GoFood, GrabFood, ShopeeFood)...")
        df = load_all_local_masters()
        if df.empty:
            print("[*] Master lokal kosong, mencoba memuat dari Tab 'DBR'...")
            df = load_dbr_master_data()

    if df.empty:
        print("❌ Tidak ada data yang dapat diekspor.")
        sys.exit(1)

    ok, res = export_dataframe_to_sot(df, owner_filter=args.owner)
    if ok:
        print(f"\n✅ Ekspor ke SOT Berhasil!")
        print(f"   Ditimpa : {res.get('total_updated', 0)}")
        print(f"   Baru    : {res.get('total_inserted', 0)}")
        sys.exit(0)
    else:
        print(f"\n❌ Gagal ekspor ke SOT: {res.get('error')}")
        sys.exit(1)


if __name__ == "__main__":
    main()

