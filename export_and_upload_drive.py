#!/usr/bin/env python3
"""
export_and_upload_drive.py
==========================
Menggabungkan data outlet per-owner dari multi-aplikator (GoFood dan GrabFood,
serta siap untuk ShopeeFood) ke dalam satu file per-owner:
    YYYY-MM-DD HH_MM [Nama Pemilik].xlsx

Spesifikasi File:
- Menggunakan schema standar 48 kolom (sheet 'Listing').
- Memiliki 2 tab (sheet) yang 100% identik:
    Tab 1: 'Listing'
    Tab 2: 'Listing 2'
- Disimpan lokal di folder `output_owners/`.
- Diunggah ke Google Drive ke dalam folder per-owner (membuat folder otomatis jika belum ada)
  melalui Google Apps Script Web App (APP_SCRIPT_URL).
"""

import os
import sys
import re
import glob
import json
import base64
import hashlib
import tempfile
import argparse
import datetime
import time
import urllib.request
import io
import requests
from pathlib import Path
from dotenv import load_dotenv

import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

from pipeline_lock import JobLock, is_pipeline_locked, get_lock_info, Timeout

# ─── Konfigurasi Direktori & URL ───────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent
GOFOOD_DIR = BASE_DIR / "GOFOOD"
GRAB_DIR = BASE_DIR / "GRAB"
SHOPEE_DIR = BASE_DIR / "SHOPEE"
CACHE_DIR = BASE_DIR / "cache"
OUTPUT_OWNERS_DIR = BASE_DIR / "output_owners"
TEMPLATE_PATH = BASE_DIR / "YYYY-MM-DD HH_MM Nama Pemilik.xlsx"

load_dotenv(BASE_DIR / ".env")
GOOGLE_SHEET_DBR_URL = os.getenv("GOOGLE_SHEET_DBR_URL", "")
GOOGLE_SHEET_VERCEL_URL = os.getenv(
    "GOOGLE_SHEET_VERCEL_URL",
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vTprbPPf_J5gAVL3PYeHbbdl5ZXQvb17HY2lJGPI2xg13Ly3AGT8eYHLYmU_m1NdtkBVg-qUGv1BoEE/pub?output=csv"
)
APP_SCRIPT_URL = os.getenv("APP_SCRIPT_URL", "")

LISTING_HEADERS = [
    'Nama Pemilik', 'Nama Brand', 'Model', 'Tipe', 'Outlet', 'Nomor HP',
    'Aplikator', 'Group ID', 'Nama Listing', 'Link', 'Store ID', 'Status Listing',
    'Alamat', 'Nama Bank', 'Nama Pemilik Rekening', 'Nomor Rekening', 'Nama Akses',
    'Email FoodMaster1', 'Email FoodMaster2', 'Nama Pengguna', 'Kata Sandi',
    'Nama Portal', 'S Nomor HP Akses Pemilik', 'S Username Akses Pemilik',
    'S Kata Sandi Akses Pemilik', 'S Allvbadmin Username Akses Staff',
    'S Allvbadmin Kata Sandi Akses Staff', 'S Bot Username Akses Staff',
    'S Bot Kata Sandi Akses Staff', 'S BD Username Akses Staff',
    'S BD Kata Sandi Akses Staff', 'BD', 'Status Internal', 'Tanggal Akuisisi',
    'Tanggal Live', 'Tanggal Churn', 'Tarif', 'Status Bot', 'Vercel Kata Sandi',
    'Paket', 'Tanggal Mulai Layanan', 'Tanggal Berakhir Layanan', 'Akses Username',
    'Akses Kata Sandi', 'Komisi ShopeeFood', 'Biaya Kemasan',
    'Nominal Biaya Kemasan', 'Invite Staff'
]
WEEKLY_RESULT_COLUMNS = [
    "Nama Pemilik", "Nama Brand", "Aplikator", "Group ID", "Nama Listing", "Link",
    "Store ID", "Status Listing", "Alamat", "Nama Bank", "Nama Pemilik Rekening", "Nomor Rekening",
]


def weekly_sot_rows(df):
    """Return the small, credential-free subset that is safe to persist for resume/SOT."""
    if df is None or df.empty:
        return []
    rows = []
    for source in df.to_dict(orient="records"):
        row = {column: source.get(column, "") for column in WEEKLY_RESULT_COLUMNS}
        if not row.get("Nama Listing"):
            row["Nama Listing"] = source.get("Nama Outlet") or source.get("Nama Brand") or ""
        if not row.get("Aplikator"):
            row["Aplikator"] = source.get("Aplikasi") or ""
        for key, value in row.items():
            if value is None or (not isinstance(value, (list, dict)) and pd.isna(value)):
                row[key] = ""
            elif isinstance(value, str) and value.strip().casefold() in ("nan", "none"):
                row[key] = ""
        rows.append(row)
    return rows


def format_grab_food_link(store_id_or_link: str) -> str:
    """
    Format ID Toko atau Link GrabFood ke URL Konsumen GrabFood:
    https://food.grab.com/id/id/restaurant/x/{clean_id}/
    Untuk store ID yang berawalan 6 (misal 6-C8DAAA4VG3E3LX atau 6C8DAAA4VG3E3LX),
    tanda strip '6-' selalu dipertahankan/disisipkan.
    """
    if not store_id_or_link:
        return ""
    val = str(store_id_or_link).strip().strip("()").strip()
    if not val or val.lower() in ('nan', 'none'):
        return ""
    if "grab.com" in val:
        val = val.split("?")[0].split("#")[0].rstrip("/").split("/")[-1].strip("()").strip()

    # Tangani store ID Grab yang berawalan 6 (misal 6-C8DAAA4VG3E3LX atau tanpa strip 6C8DAAA4VG3E3LX)
    m6 = re.match(r"^6-?([A-Za-z0-9]{5,})$", val)
    if m6:
        clean_id = f"6-{m6.group(1)}"
    else:
        clean_id = val.replace("-", "").strip()

    if not clean_id:
        return ""
    return f"https://food.grab.com/id/id/restaurant/x/{clean_id}/"


def get_template_headers():
    """Mengembalikan schema Listing resmi, terlepas dari template lama."""
    return LISTING_HEADERS.copy()


def load_gofood_data():
    """Memuat data outlet GoFood dari master file dan cache JSON terbaru."""
    records = []
    
    # 1. Dari master file
    gofood_master = GOFOOD_DIR / "master" / "0master.xlsx"
    if not gofood_master.exists():
        masters = sorted(glob.glob(str(GOFOOD_DIR / "master" / "*_master.xlsx")))
        if masters:
            gofood_master = Path(masters[-1])

    if gofood_master.exists():
        try:
            df = pd.read_excel(str(gofood_master), sheet_name="Listing")
            df["Aplikator"] = "GoFood"
            records.append(df)
        except Exception as e:
            print(f"⚠️ Gagal membaca GoFood master: {e}")

    # 2. Dari cache JSON di GOFOOD/cache/
    for jf in glob.glob(str(GOFOOD_DIR / "cache" / "*.json")):
        try:
            with open(jf, "r", encoding="utf-8") as f:
                cdata = json.load(f)
            if isinstance(cdata, dict) and "outlets" in cdata and isinstance(cdata["outlets"], list):
                if cdata["outlets"]:
                    df_c = pd.DataFrame(cdata["outlets"])
                    df_c["Aplikator"] = "GoFood"
                    records.append(df_c)
        except Exception:
            pass

    # 3. Fallback dari GOFOOD/output/*.xlsx (hanya jika master & cache kosong)
    if not records:
        for f in glob.glob(str(GOFOOD_DIR / "output" / "*.xlsx")):
            try:
                df_part = pd.read_excel(f, sheet_name="Listing")
                df_part["Aplikator"] = "GoFood"
                records.append(df_part)
            except Exception:
                pass

    if records:
        combined = pd.concat(records, ignore_index=True)
        if "Nama Portal" in combined.columns:
            combined["Nama Portal"] = ""
        if "Store ID" in combined.columns:
            combined = combined.drop_duplicates(subset=["Store ID"], keep="last")
        return combined
        
    return pd.DataFrame()


def load_grab_data():
    """Memuat data outlet GrabFood dari master file dan cache JSON terbaru."""
    records = []
    
    # 1. Dari master file
    grab_master = GRAB_DIR / "master" / "0master.xlsx"
    if not grab_master.exists():
        masters = sorted(glob.glob(str(GRAB_DIR / "master" / "*_master.xlsx")))
        if masters:
            grab_master = Path(masters[-1])

    if grab_master.exists():
        try:
            df = pd.read_excel(str(grab_master), sheet_name="Listing")
            df["Aplikator"] = "GrabFood"
            records.append(df)
        except Exception as e:
            print(f"⚠️ Gagal membaca Grab master: {e}")

    # 2. Dari cache JSON di GRAB/cache/
    for jf in glob.glob(str(GRAB_DIR / "cache" / "*.json")):
        try:
            with open(jf, "r", encoding="utf-8") as f:
                cdata = json.load(f)
            if isinstance(cdata, dict) and "outlets" in cdata and isinstance(cdata["outlets"], list):
                if cdata["outlets"]:
                    df_c = pd.DataFrame(cdata["outlets"])
                    df_c["Aplikator"] = "GrabFood"
                    records.append(df_c)
        except Exception:
            pass

    # 3. Fallback dari GRAB/output/*.xlsx (hanya jika master & cache kosong)
    if not records:
        for f in glob.glob(str(GRAB_DIR / "output" / "*.xlsx")):
            try:
                df_part = pd.read_excel(f, sheet_name="Listing")
                df_part["Aplikator"] = "GrabFood"
                records.append(df_part)
            except Exception:
                pass

    if records:
        combined = pd.concat(records, ignore_index=True)
        if "Store ID" in combined.columns:
            def _normalize_grab_sid(sid_val):
                s = str(sid_val or '').strip().strip("()").strip()
                m6 = re.match(r"^6([A-Za-z0-9]{5,})$", s)
                return f"6-{m6.group(1)}" if m6 else s
            combined["Store ID"] = combined["Store ID"].apply(_normalize_grab_sid)
            combined = combined.drop_duplicates(subset=["Store ID"], keep="last")
        if "Link" in combined.columns or "Store ID" in combined.columns:
            def _normalize_grab_link(row):
                sid = row.get("Store ID") if pd.notna(row.get("Store ID")) else ""
                link = row.get("Link") if pd.notna(row.get("Link")) else ""
                return format_grab_food_link(link or sid)
            combined["Link"] = combined.apply(_normalize_grab_link, axis=1)
        return combined

    return pd.DataFrame()


def load_shopee_data():
    """Memuat data outlet ShopeeFood dari master file atau output terbaru."""
    shopee_master = SHOPEE_DIR / "data" / "0master.xlsx"
    if not shopee_master.exists():
        masters = sorted(glob.glob(str(SHOPEE_DIR / "data" / "Shopee_*.xlsx")), key=os.path.getmtime)
        if masters:
            shopee_master = Path(masters[-1])

    if shopee_master.exists():
        try:
            df = pd.read_excel(str(shopee_master), sheet_name="Listing")
            df["Aplikator"] = "ShopeeFood"
            return df
        except Exception as e:
            print(f"⚠️ Gagal membaca Shopee master: {e}")

    return pd.DataFrame()


def get_known_shopee_merchants() -> list[str]:
    """Mengambil daftar nama merchant Shopee yang diketahui dari cache merchant_list.json."""
    try:
        p = SHOPEE_DIR / "data" / "merchant_list.json"
        if p.exists():
            data = json.loads(p.read_text(encoding="utf-8"))
            body = data.get("data") or data
            if isinstance(body, dict):
                m_list = body.get("selectMerchant", {}).get("merchantList", [])
                if not m_list and "merchantList" in body:
                    m_list = body.get("merchantList", [])
                return [m.get("merchantName", "").strip() for m in m_list if m.get("merchantName")]
    except Exception:
        pass
    return []


def load_vercel_data(force_live=False):
    """
    Memuat data outlet dari Google Sheet Vercel (Live CSV).
    Menyimpan cache lokal di cache/vercel_sheet_cache.csv untuk fallback offline.
    """
    cache_file = CACHE_DIR / "vercel_sheet_cache.csv"
    try:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass

    csv_text = ""

    # 1. Gunakan cache lokal jika tidak force_live dan file cache masih segar (< 5 menit)
    if not force_live and cache_file.exists():
        try:
            mtime = cache_file.stat().st_mtime
            if (time.time() - mtime) < 300:
                with open(cache_file, "r", encoding="utf-8") as f:
                    cached_data = f.read()
                if cached_data.strip():
                    csv_text = cached_data
        except Exception:
            pass

    # 2. Ambil data live dari Google Sheet Vercel URL
    if not csv_text:
        target_url = GOOGLE_SHEET_VERCEL_URL or "https://docs.google.com/spreadsheets/d/e/2PACX-1vTprbPPf_J5gAVL3PYeHbbdl5ZXQvb17HY2lJGPI2xg13Ly3AGT8eYHLYmU_m1NdtkBVg-qUGv1BoEE/pub?output=csv"
        try:
            url = target_url.strip()
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
            with urllib.request.urlopen(req, timeout=15) as resp:
                downloaded_text = resp.read().decode('utf-8', errors='replace')
            if downloaded_text.strip():
                csv_text = downloaded_text
                print(f"   ✓ [VERCEL] Berhasil memuat data live dari Google Sheet ({len(csv_text)} bytes).")
                try:
                    with open(cache_file, "w", encoding="utf-8") as f:
                        f.write(csv_text)
                except Exception as ce:
                    print(f"   ℹ️ [CACHE] Cache disk tidak dapat diperbarui ({ce}), tetap memakai data live.")
        except Exception as e:
            print(f"⚠️ Fetch live Vercel Sheet melewati batas waktu / offline ({e}).")

    # 3. Fallback ke file cache lokal jika live gagal
    if not csv_text and cache_file.exists():
        try:
            with open(cache_file, "r", encoding="utf-8") as f:
                csv_text = f.read()
            print(f"   ℹ️ [VERCEL] Menggunakan cache lokal ({cache_file.name}).")
        except Exception:
            pass

    if not csv_text:
        return pd.DataFrame()

    try:
        df = pd.read_csv(io.StringIO(csv_text))
    except Exception as e:
        print(f"⚠️ Gagal membaca CSV Vercel: {e}")
        return pd.DataFrame()

    # Bersihkan whitespace di nama kolom
    df.columns = [str(c).strip() for c in df.columns]

    # Map nama kolom Vercel ke nama standar internal
    if "Owner" in df.columns:
        df["Nama Pemilik"] = df["Owner"].astype(str).str.strip()
    if "Nama Outlet" in df.columns:
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

    if "Aplikasi" in df.columns:
        df["Aplikator"] = df["Aplikasi"].apply(clean_app)

    # Pastikan kolom Nama Portal ada dan bersih
    if "Nama Portal" not in df.columns:
        df["Nama Portal"] = ""

    # Untuk Shopee, fallback Nama Portal ke Merchant Name jika kosong
    if "Merchant Name" in df.columns:
        df["Nama Portal"] = df.apply(
            lambda r: r["Merchant Name"] if (pd.isna(r.get("Nama Portal")) or not str(r.get("Nama Portal")).strip() or str(r.get("Nama Portal")).strip().lower() in ("nan", "none")) and pd.notna(r.get("Merchant Name")) else r.get("Nama Portal", ""),
            axis=1
        )

    # Untuk GoFood, kosongkan kolom Nama Portal sesuai instruksi
    if "Aplikator" in df.columns:
        is_go = df["Aplikator"].astype(str).str.strip().str.lower().str.contains("go")
        df.loc[is_go, "Nama Portal"] = ""
        if "Nama Portal.1" in df.columns:
            df.loc[is_go, "Nama Portal.1"] = ""

        # Normalisasi Store ID dan Link untuk GrabFood agar format '6-' selalu lengkap
        is_grab = df["Aplikator"].astype(str).str.strip().str.lower().str.contains("grab")
        if is_grab.any():
            if "Store ID" in df.columns:
                def _norm_sid(s):
                    val = str(s or '').strip().strip("()").strip()
                    m = re.match(r"^6([A-Za-z0-9]{5,})$", val)
                    return f"6-{m.group(1)}" if m else val
                df.loc[is_grab, "Store ID"] = df.loc[is_grab, "Store ID"].apply(_norm_sid)
            if "Link" in df.columns:
                df.loc[is_grab, "Link"] = df[is_grab].apply(
                    lambda r: format_grab_food_link(r.get("Link") or r.get("Store ID")), axis=1
                )

    # Status Listing default LIVE untuk data aktif
    df["Status Listing"] = "LIVE"

    return df


def write_data_to_sheet(ws, df, headers):
    """Menulis data ke sheet yang telah memiliki header, styling, dan lebar kolom dari template."""
    body_font = Font(name='Arial', size=10)
    alignment = Alignment(horizontal='left', vertical='center')

    # Tulis ulang header agar template lama ikut menghasilkan schema terbaru.
    for c_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c_idx, value=header)
        cell.font = Font(name='Arial', size=10, bold=True)
        cell.alignment = alignment

    # Alias mapping jika nama kolom di df sedikit berbeda dengan nama header template
    col_aliases = {
        'Nama Pemilik': ['Owner', 'Nama Pemilik', 'owner', 'pemilik'],
        'Nama Brand': ['Nama Outlet', 'Nama Brand', 'Brand', 'brand'],
        'Aplikator': ['Aplikasi', 'Aplikator', 'app', 'Platform'],
        'Nama Portal': ['Nama Portal', 'Merchant Name', 'portal'],
        'Nama Listing': ['Nama Listing', 'Nama Outlet', 'Nama Brand', 'Listing'],
        'Nomor HP Akses Pemilik': ['S Nomor HP Akses Pemilik', 'Nomor HP Akses Pemilik', 'Nomor HP'],
        'Username Akses Pemilik': ['S Username Akses Pemilik', 'Username Akses Pemilik'],
        'Kata Sandi Akses Pemilik': ['S Kata Sandi Akses Pemilik', 'Kata Sandi Akses Pemilik'],
        'S BD Username Akses Staff': ['S Username Akses Staff', 'S BD Username Akses Staff', 'Username Akses Staff'],
        'S BD Kata Sandi Akses Staff': ['S Kata Sandi Akses Staff', 'S BD Kata Sandi Akses Staff', 'Kata Sandi Akses Staff'],
        'BD': ['BD', 'bd'],
    }

    text_cols = {'Nomor Rekening', 'Nomor HP', 'S Nomor HP Akses Pemilik', 'Store ID', 'Group ID'}

    # Tulis Baris Data mulai baris ke-2
    for idx, (_, r) in enumerate(df.iterrows(), start=2):
        for c_idx, h in enumerate(headers, 1):
            if not h:
                continue

            # Sesuai instruksi: untuk kolom brand ambil dari kolom Nama Outlet saja
            if h in ('Nama Brand', 'Brand'):
                val = r.get('Nama Outlet')
                if val is None or pd.isna(val) or str(val).strip() == '':
                    val = r.get('Nama Brand') or r.get('Brand')
            else:
                val = r.get(h)
                if val is None or pd.isna(val) or str(val).strip() == '':
                    # Cari via aliases
                    aliases = col_aliases.get(h, [])
                    for alias in aliases:
                        if alias in r and pd.notna(r.get(alias)) and str(r.get(alias)).strip() != '':
                            val = r.get(alias)
                            break

            val_str = str(val).strip() if pd.notna(val) and val is not None else ''
            if val_str.lower() in ('nan', 'none'):
                val_str = ''

            # Bersihkan suffix .0 pada angka/nomor panjang
            if val_str.endswith(".0") and val_str[:-2].isdigit():
                val_str = val_str[:-2]

            # Normalisasi kolom Store ID untuk GrabFood jika berawalan 6 tapi tanpa strip
            if h == 'Store ID':
                app_name = str(r.get('Aplikator') or r.get('Aplikasi') or '').strip().lower()
                if 'grab' in app_name:
                    m6 = re.match(r"^6([A-Za-z0-9]{5,})$", val_str.strip().strip("()").strip())
                    if m6:
                        val_str = f"6-{m6.group(1)}"

            # Normalisasi kolom Link untuk GrabFood ke format URL konsumen
            if h == 'Link':
                app_name = str(r.get('Aplikator') or r.get('Aplikasi') or '').strip().lower()
                if 'grab' in app_name:
                    sid = str(r.get('Store ID') or '').strip()
                    val_str = format_grab_food_link(val_str or sid)

            # Kosongkan kolom Nama Portal untuk GoFood sesuai instruksi
            if h in ('Nama Portal', 'Nama Portal.1'):
                app_name = str(r.get('Aplikator') or r.get('Aplikasi') or '').strip().lower()
                if 'go' in app_name:
                    val_str = ''

            cell = ws.cell(row=idx, column=c_idx, value=val_str)
            cell.font = body_font
            cell.alignment = alignment
            if h in text_cols or c_idx in (6, 8, 11, 16, 23):
                cell.number_format = '@'


def sort_df_by_aplikator(df):
    """
    Mengurutkan DataFrame berdasarkan urutan standar aplikator:
    1. GoFood
    2. GrabFood
    3. ShopeeFood
    Menggunakan stable sort agar urutan outlet di dalam aplikator yang sama tetap terjaga.
    """
    if df is None or df.empty or "Aplikator" not in df.columns:
        return df
    app_rank = {
        "gofood": 1,
        "grabfood": 2,
        "grab": 2,
        "shopeefood": 3,
        "shopee": 3
    }
    df_copy = df.copy()
    df_copy["_app_sort_order"] = (
        df_copy["Aplikator"]
        .astype(str)
        .str.strip()
        .str.lower()
        .map(app_rank)
        .fillna(99)
    )
    df_sorted = (
        df_copy.sort_values(by=["_app_sort_order"], kind="stable")
        .drop(columns=["_app_sort_order"])
        .reset_index(drop=True)
    )
    return df_sorted


def save_owner_workbook(owner_df, file_path, headers):
    """
    Menyimpan data per-owner ke file Excel dengan 2 tab 100% identik:
    - Tab 1: 'Listing'
    - Tab 2: 'Listing 2'
    Dibuat langsung dengan mengkloning template 'YYYY-MM-DD HH_MM Nama Pemilik.xlsx'
    sehingga pewarnaan kolom header (Merah, Pink, Hijau, Oranye), font Arial,
    dan lebar kolom asli terjaga 100% identik.
    Menggunakan penulisan atomic via folder .tmp agar terhindar dari file korup jika terjadi crash.
    Data selalu diurutkan dengan urutan standar: GoFood -> GrabFood -> ShopeeFood.
    """
    owner_df = sort_df_by_aplikator(owner_df)
    target_path = Path(file_path)
    tmp_dir = target_path.parent / ".tmp"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    tmp_file_path = tmp_dir / f"tmp_{os.getpid()}_{target_path.name}"

    if TEMPLATE_PATH.exists():
        wb = openpyxl.load_workbook(str(TEMPLATE_PATH))
        if "Listing" in wb.sheetnames:
            ws1 = wb["Listing"]
        else:
            ws1 = wb.active
            ws1.title = "Listing"
        
        # Bersihkan baris dummy template di bawah header
        if ws1.max_row > 1:
            ws1.delete_rows(2, ws1.max_row)

        if "Listing 2" in wb.sheetnames:
            del wb["Listing 2"]

        ws2 = wb.copy_worksheet(ws1)
        ws2.title = "Listing 2"
    else:
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "Listing"
        ws2 = wb.create_sheet(title="Listing 2")

    write_data_to_sheet(ws1, owner_df, headers)
    write_data_to_sheet(ws2, owner_df, headers)
    
    wb.save(str(tmp_file_path))
    os.replace(str(tmp_file_path), str(target_path))


def upload_file_to_drive(file_path, owner_name, app_script_url):
    """Mengunggah file Excel per-owner ke Google Drive via Google Apps Script Web App."""
    if not app_script_url:
        print("   ⚠️ APP_SCRIPT_URL belum disetel di .env. Lewati upload.")
        return False, {"error": "APP_SCRIPT_URL belum disetel di .env"}

    filename = os.path.basename(file_path)
    print(f"   ☁️ Mengunggah {filename} ke folder Google Drive: [{owner_name}]...")
    
    try:
        with open(file_path, "rb") as f:
            b64_data = base64.b64encode(f.read()).decode("utf-8")

        payload = {
            "fileName": filename,
            "ownerName": owner_name,
            "fileData": b64_data,
            "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }

        response = requests.post(app_script_url, json=payload, timeout=90, allow_redirects=True)
        if response.status_code in (200, 201):
            try:
                res_data = response.json()
                if res_data.get("status") == "success":
                    folder_url = res_data.get("folderUrl") or res_data.get("url", "")
                    file_url = res_data.get("fileUrl", "")
                    print(f"   ✅ Berhasil diunggah! Folder: {folder_url}")
                    return True, {
                        "folderUrl": folder_url,
                        "fileUrl": file_url,
                        "folderName": res_data.get("folderName", owner_name),
                        "fileName": filename
                    }
                else:
                    msg = res_data.get("message", response.text)
                    print(f"   ❌ Gagal: {msg}")
                    return False, {"error": msg}
            except Exception:
                print("   ✅ Berhasil diunggah (respons non-JSON 200).")
                return True, {"folderUrl": "https://drive.google.com/drive/u/0/folders/19VIrypPcBmNNbjDLGS7kxp_yIdBBwXjB", "fileName": filename}
        else:
            print(f"   ❌ HTTP Error {response.status_code}: {response.text}")
            return False, {"error": f"HTTP {response.status_code}"}

    except Exception as e:
        print(f"   ⚠️ Exception saat mengunggah: {e}")
        return False, {"error": str(e)}


def get_owners_with_metadata(source="vercel", force_live=False):
    """
    Mengambil daftar owner beserta informasi outlet dan timestamp file terbaru.
    source: 'vercel' (default) menggunakan live CSV Vercel Sheet,
            'master' menggunakan master lokal, 'dbr' menggunakan DBR.
    """
    if source == "vercel":
        combined = load_vercel_data(force_live=force_live)
    elif source == "dbr":
        from cronjobsot.export_to_sot import load_dbr_master_data
        combined = load_dbr_master_data(force_live=force_live)
    else:
        df_go = load_gofood_data()
        df_grab = load_grab_data()
        df_shopee = load_shopee_data()
        dfs = [df for df in [df_go, df_grab, df_shopee] if not df.empty]
        combined = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
    
    if source != "dbr" and (combined.empty or "Nama Pemilik" not in combined.columns):
        # Fallback jika Vercel kosong
        df_go = load_gofood_data()
        df_grab = load_grab_data()
        df_shopee = load_shopee_data()
        dfs = [df for df in [df_go, df_grab, df_shopee] if not df.empty]
        combined = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

    if combined.empty or "Nama Pemilik" not in combined.columns:
        return []

    combined = combined[combined["Nama Pemilik"].notna()]
    combined = combined[combined["Nama Pemilik"].astype(str).str.strip() != ""]
    combined = combined[combined["Nama Pemilik"].astype(str).str.strip().str.lower() != "nan"]
    
    # Pertahankan urutan unik
    seen = set()
    owners = []
    for o in combined["Nama Pemilik"]:
        o_clean = str(o).strip()
        if o_clean and o_clean.lower() not in seen:
            seen.add(o_clean.lower())
            owners.append(o_clean)

    results = []
    for owner_str in owners:
        o_df = combined[combined["Nama Pemilik"].str.strip().str.lower() == owner_str.lower()]
        go_n = len(o_df[o_df["Aplikator"] == "GoFood"]) if "Aplikator" in o_df.columns else 0
        gr_n = len(o_df[o_df["Aplikator"] == "GrabFood"]) if "Aplikator" in o_df.columns else 0
        sh_n = len(o_df[o_df["Aplikator"] == "ShopeeFood"]) if "Aplikator" in o_df.columns else 0
        
        # Cari timestamp terbaru dari file owner di output_owners/
        clean_name = "".join(c for c in owner_str if c.isalnum() or c in " ._-").strip()
        files = glob.glob(str(OUTPUT_OWNERS_DIR / f"*{clean_name}*.xlsx"))
        files.extend(glob.glob(str(GOFOOD_DIR / "output" / f"*{clean_name}*.xlsx")))
        files.extend(glob.glob(str(GRAB_DIR / "output" / f"*{clean_name}*.xlsx")))
        files.extend(glob.glob(str(SHOPEE_DIR / "data" / f"*{clean_name}*.xlsx")))
        
        latest_mtime = 0
        latest_ts_str = "Belum Ada"
        if files:
            latest_file = max(files, key=os.path.getmtime)
            latest_mtime = os.path.getmtime(latest_file)
            latest_ts_str = datetime.datetime.fromtimestamp(latest_mtime).strftime("%Y-%m-%d %H:%M")

        # Ambil nama brand / outlet
        brands = []
        for col in ["Nama Brand", "Nama Outlet", "Brand"]:
            if col in o_df.columns:
                unique_brands = [str(b).strip() for b in o_df[col].dropna().unique() if str(b).strip() and str(b).strip().lower() not in ("nan", "none", "")]
                if unique_brands:
                    brands = unique_brands
                    break
        brand_name = ", ".join(brands) if brands else ""

        results.append({
            "owner": owner_str,
            "brand": brand_name,
            "total": len(o_df),
            "gofood": go_n,
            "grab": gr_n,
            "shopee": sh_n,
            "latest_mtime": latest_mtime,
            "latest_ts": latest_ts_str,
            "is_newest": False
        })

    # Urutkan berdasarkan timestamp modifikasi jika ada file baru
    results.sort(key=lambda x: x["latest_mtime"], reverse=True)
    if results and results[0]["latest_mtime"] > 0:
        results[0]["is_newest"] = True

    return results


def _mask_account_label(value):
    value = str(value or "").strip()
    if not value:
        return "(kosong)"
    if "@" in value:
        local, domain = value.split("@", 1)
        return f"{local[:1]}***@{domain}" if local else f"***@{domain}"
    if len(value) <= 3:
        return "*" * len(value)
    return f"{value[:2]}***{value[-1:]}"


def _redact_weekly_log(line, redactions=()):
    safe_line = str(line)
    for secret, replacement in sorted(redactions, key=lambda item: len(item[0]), reverse=True):
        if secret:
            safe_line = safe_line.replace(secret, replacement)
    safe_line = re.sub(r"(?i)\b(set-cookie|cookie)\s*[:=].*$", r"\1=[REDACTED]", safe_line)
    safe_line = re.sub(r"(?i)\b(authorization)\s*[:=]\s*(?:bearer|basic)\s+\S+", r"\1: [REDACTED]", safe_line)
    safe_line = re.sub(r"(?i)\b(bearer)\s+\S+", r"\1 [REDACTED]", safe_line)
    safe_line = re.sub(r"(?i)\b(access[_ -]?token|refresh[_ -]?token|csrf[_ -]?token)\b(\s*[:=]\s*)([^\s,;]+)", r"\1\2[REDACTED]", safe_line)
    safe_line = re.sub(
        r"(?i)\b(password|kata sandi|token|cookie)\b(\s*[:=]\s*)([^\s,;]+)",
        r"\1\2[REDACTED]",
        safe_line,
    )
    return safe_line


def _weekly_owner_targets(dbr_df, owner_name, app_key):
    """Build portal groups from the one DBR snapshot; credentials stay in memory."""
    if dbr_df is None or dbr_df.empty or "Nama Pemilik" not in dbr_df.columns:
        return []
    owner_rows = dbr_df[
        dbr_df["Nama Pemilik"].fillna("").astype(str).str.strip().str.casefold()
        == str(owner_name).strip().casefold()
    ]
    if "Aplikator" not in owner_rows.columns:
        return []
    app_values = owner_rows["Aplikator"].fillna("").astype(str).str.casefold()
    if app_key == "gofood":
        owner_rows = owner_rows[app_values.str.contains("gofood")]
        portal_col = "Nama Portal"
        account_cols = ["Email FoodMaster1", "Email FoodMaster2"]
    elif app_key == "grab":
        owner_rows = owner_rows[app_values.str.contains("grab")]
        portal_col = "Nama Portal"
        account_cols = ["Nama Pengguna"]
    else:
        owner_rows = owner_rows[app_values.str.contains("shopee")]
        portal_col = "Nama Portal"
        account_cols = []

    targets_by_key = {}
    for _, row in owner_rows.iterrows():
        portal = str(row.get(portal_col) or "").strip()
        if not portal or portal.casefold() in ("nan", "none", "-"):
            continue
        accounts = [
            str(row.get(col) or "").strip()
            for col in account_cols
            if str(row.get(col) or "").strip().casefold() not in ("", "nan", "none", "-")
        ]
        has_password = any(
            str(row.get(column) or "").strip().casefold() not in ("", "nan", "none", "-")
            for column in ("Kata Sandi", "Password", "Akses Kata Sandi")
        )
        account = accounts[0] if accounts else (portal if app_key == "shopee" else "")
        key = (account.casefold(),) if account else ("portal", portal.casefold())
        if key not in targets_by_key:
            brand = str(row.get("Nama Brand") or row.get("Nama Outlet") or portal).strip()
            targets_by_key[key] = {
                "owner": owner_name,
                "brand": brand or portal,
                "portal": portal,
                "portals": [portal],
                "account": account,
                "accounts": accounts,
                "has_password": has_password,
                "app": app_key,
            }
        else:
            if portal.casefold() not in {p.casefold() for p in targets_by_key[key]["portals"]}:
                targets_by_key[key]["portals"].append(portal)
            targets_by_key[key]["has_password"] = targets_by_key[key]["has_password"] or has_password
    return list(targets_by_key.values())


def _weekly_missing_portals(dbr_df, owner_name):
    missing = {"gofood": [], "grab": [], "shopee": []}
    if dbr_df is None or dbr_df.empty or not {"Nama Pemilik", "Aplikator"}.issubset(dbr_df.columns):
        return missing
    owner_rows = dbr_df[
        dbr_df["Nama Pemilik"].fillna("").astype(str).str.strip().str.casefold()
        == str(owner_name).strip().casefold()
    ]
    for _, row in owner_rows.iterrows():
        app = str(row.get("Aplikator") or "").casefold()
        app_key = "gofood" if "gofood" in app else "grab" if "grab" in app else "shopee" if "shopee" in app else None
        portal = str(row.get("Nama Portal") or "").strip()
        if app_key and (not portal or portal.casefold() in ("nan", "none", "-")):
            brand = str(row.get("Nama Brand") or row.get("Nama Outlet") or "").strip()
            if brand not in missing[app_key]:
                missing[app_key].append(brand or "(Nama Portal kosong)")
    return missing


def _weekly_redactions(dbr_df, owner_name):
    redactions = []
    if dbr_df is None or dbr_df.empty or "Nama Pemilik" not in dbr_df.columns:
        return redactions
    owner_rows = dbr_df[
        dbr_df["Nama Pemilik"].fillna("").astype(str).str.strip().str.casefold()
        == str(owner_name).strip().casefold()
    ]
    for column in owner_rows.columns:
        normalized = str(column).casefold()
        is_password = "sandi" in normalized or "password" in normalized
        is_account = any(token in normalized for token in ("email", "username", "pengguna"))
        if not (is_password or is_account):
            continue
        for value in owner_rows[column].dropna().astype(str).unique():
            value = value.strip()
            if value.casefold() in ("", "nan", "none", "-"):
                continue
            replacement = "[REDACTED]" if is_password else _mask_account_label(value)
            redactions.append((value, replacement))
    return redactions


def _cache_fingerprints(directory, prefix):
    fingerprints = {}
    try:
        for path in Path(directory).glob(f"{prefix}*.json"):
            try:
                digest = hashlib.sha256(path.read_bytes()).hexdigest()
                fingerprints[path] = digest
            except OSError:
                continue
    except OSError:
        pass
    return fingerprints


def _read_changed_portal_caches(directory, prefix, before, owner_name, app_name):
    changed = []
    after = _cache_fingerprints(directory, prefix)
    for path, digest in after.items():
        if before.get(path) == digest:
            continue
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
            outlets = payload.get("outlets", [])
            if not isinstance(outlets, list):
                continue
            payload_owner = str(payload.get("owner") or owner_name).strip().casefold()
            if payload_owner != str(owner_name).strip().casefold():
                continue
            changed.append({
                "portal": str(payload.get("portal") or ""),
                "account": str(payload.get("email") or payload.get("username") or ""),
                "outlets": [dict(item) for item in outlets if isinstance(item, dict)],
            })
        except Exception:
            continue
    rows = []
    for entry in changed:
        for outlet in entry["outlets"]:
            row = dict(outlet)
            row.setdefault("Nama Pemilik", owner_name)
            row["Aplikator"] = app_name
            rows.append(row)
    return changed, rows


def _portal_error_context(lines, portal, return_code, allow_unscoped=False):
    norm = lambda value: re.sub(r"[^a-z0-9]", "", str(value or "").casefold())
    portal_key = norm(portal)
    related = [line for line in lines if portal_key and portal_key in norm(line)]
    if not related and allow_unscoped:
        related = list(lines)
    error_lines = [
        line for line in related
        if any(word in line.casefold() for word in ("gagal", "error", "exception", "timeout", "tidak berhasil", "kosong"))
    ]
    detail = (error_lines or related)[-1] if (error_lines or related) else f"proses scraper berhenti dengan kode {return_code}"
    lowered = detail.casefold()
    if any(word in lowered for word in ("login", "autentikasi", "otp", "password", "kata sandi", "sesi")):
        stage = "autentikasi"
    elif any(word in lowered for word in ("store", "outlet", "detail", "fetch", "data kosong", "http")):
        stage = "penarikan detail outlet"
    else:
        stage = "scraping"
    return stage, detail[:400]


def run_subprocess_stream(cmd, cwd, keywords, on_log, timeout_sec=150, cancel_event=None,
                          redactions=(), on_line=None, extra_env=None):
    """Menjalankan subprocess dengan streaming output real-time, dukungan pembatalan cepat, dan batas waktu yang aman."""
    import subprocess
    import time
    import signal
    import selectors
    env = {**os.environ, "PYTHONUNBUFFERED": "1"}
    if extra_env:
        env.update(extra_env)
    try:
        p = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            cwd=str(cwd),
            env=env,
            bufsize=1,
            preexec_fn=os.setsid
        )
        sel = selectors.DefaultSelector()
        sel.register(p.stdout, selectors.EVENT_READ)
        start_time = time.time()

        while True:
            # Periksa apakah proses dibatalkan oleh pengguna (responsif < 0.2 detik)
            if cancel_event and cancel_event.is_set():
                try:
                    os.killpg(os.getpgid(p.pid), signal.SIGKILL)
                except Exception:
                    try:
                        p.kill()
                    except Exception:
                        pass
                try:
                    p.wait(timeout=2)
                except Exception:
                    pass
                on_log("🛑 Proses dibatalkan oleh pengguna.")
                try:
                    sel.unregister(p.stdout)
                    sel.close()
                except Exception:
                    pass
                return -999

            # Cegah proses menggantung melebihi timeout
            if time.time() - start_time > timeout_sec:
                try:
                    os.killpg(os.getpgid(p.pid), signal.SIGKILL)
                except Exception:
                    try:
                        p.kill()
                    except Exception:
                        pass
                try:
                    p.wait(timeout=2)
                except Exception:
                    pass
                on_log(f"⏱️ Melewati batas waktu ({timeout_sec}s). Melanjutkan ke tahap berikutnya...")
                break

            events = sel.select(timeout=0.2)
            if events:
                line = p.stdout.readline()
                if not line:
                    if p.poll() is not None:
                        break
                    continue
                
                line_s = _redact_weekly_log(line.strip(), redactions)
                if not line_s:
                    continue

                if on_line:
                    on_line(line_s)
                    
                # Abaikan garis pembatas dekoratif
                if line_s.startswith("===") or line_s.startswith("---") or line_s.startswith("───"):
                    continue

                # Tampilkan pesan jika ada indikator status atau cocok kata kunci
                is_status = any(line_s.startswith(p) for p in ("[*]", "[✓]", "🚀", "🌐", "➡️", "📧", "⏳", "✅", "⚠️", "❌", "🎉", "⚡", "🏢", "📍", "🚨"))
                has_kw = any(k.lower() in line_s.lower() for k in keywords)
                if is_status or has_kw:
                    on_log(line_s[:500] if "[SHOPEE]" in line_s else line_s[:85])
            else:
                if p.poll() is not None:
                    break

        try:
            sel.unregister(p.stdout)
            sel.close()
        except Exception:
            pass

        p.poll()
        return p.returncode or 0
    except Exception as e:
        on_log(f"⚠️ Error proses: {e}")
        return -1


def _run_weekly_dbr_scraping(owner_name, dbr_df, progress_cb=None, cancel_event=None):
    """Scrape every DBR portal once and return only rows written by this run."""
    python_bin = sys.executable
    redactions = _weekly_redactions(dbr_df, owner_name)
    result = {
        "gofood": None,
        "grab": None,
        "shopee": None,
        "shopee_merchants": {},
        "detector_warnings": [],
        "portal_results": [],
        "live_rows": [],
    }

    def log(pct, message):
        if progress_cb:
            progress_cb(pct, message)
        print(f"[{pct:>3}%] {message}", flush=True)

    def log_child(pct, app_label, line):
        if line:
            log(pct, f"[WEEKLY][{app_label}] {line}")

    def run_platform(app_key, app_label, script_path, cache_dir, cache_prefix, args, targets, pct, timeout):
        before = _cache_fingerprints(cache_dir, cache_prefix)
        captured_lines = []
        for target in targets:
            for portal in target["portals"]:
                account_label = _mask_account_label(target.get("account"))
                log(pct, f"[WEEKLY][PORTAL][START] Owner='{owner_name}' | Aplikator={app_label} | Portal='{portal}' | Akun={account_label}.")
                log(pct, f"[WEEKLY][PORTAL][AUTH] Memulai autentikasi/pemulihan sesi untuk portal '{portal}'.")

        rc = run_subprocess_stream(
            [python_bin, "-u", str(script_path), *args],
            cwd=script_path.parent,
            keywords=("Store ID", "Berhasil", "Portal", "Login", "autentikasi", "Owner", "OTP", "Filter", "Gagal", "Error", "Exception", "Restricted", "DETEKTOR", "GAGAL_LOGIN", "DATA_KOSONG", "Target", "Sukses", "Outlet", "Merchant"),
            on_log=lambda line: log_child(pct, app_label, line),
            timeout_sec=timeout,
            cancel_event=cancel_event,
            redactions=redactions,
            on_line=captured_lines.append,
        )
        changed, all_rows = _read_changed_portal_caches(cache_dir, cache_prefix, before, owner_name, app_label)
        result[app_key] = rc

        for target in targets:
            matching_entries = []
            account_set = {str(value).strip().casefold() for value in target.get("accounts", []) if str(value).strip()}
            portal_keys = [re.sub(r"[^a-z0-9]", "", p.casefold()) for p in target["portals"]]
            for entry in changed:
                account_match = bool(account_set and entry["account"].strip().casefold() in account_set)
                entry_portal_key = re.sub(r"[^a-z0-9]", "", entry["portal"].casefold())
                portal_match = any(key and (key in entry_portal_key or entry_portal_key in key) for key in portal_keys)
                if account_match or portal_match:
                    matching_entries.append(entry)

            target_rows = [row for entry in matching_entries for row in entry["outlets"]]
            for row in target_rows:
                if not str(row.get("Nama Pemilik") or "").strip():
                    row["Nama Pemilik"] = owner_name
                if not str(row.get("Nama Brand") or "").strip():
                    row["Nama Brand"] = target["brand"]
                if not str(row.get("Nama Portal") or "").strip():
                    row["Nama Portal"] = target["portal"]
                row["Aplikator"] = app_label

            if target_rows:
                state = "success"
                stage = "detail outlet"
                message = f"{len(target_rows)} outlet berhasil ditarik."
                result["live_rows"].extend(target_rows)
                log(pct, f"[WEEKLY][PORTAL][AUTH_SUCCESS] Sesi/auth akun untuk portal '{target['portal']}' berhasil.")
                log(pct, f"[WEEKLY][PORTAL][SELECTED] Portal '{target['portal']}' berhasil diproses.")
                log(pct, f"[WEEKLY][PORTAL][STORES] Store ditemukan={len(target_rows)}.")
                log(pct, f"[WEEKLY][PORTAL][DETAIL] Berhasil={len(target_rows)} | Gagal=0.")
                log(pct, f"[WEEKLY][PORTAL][SUCCESS] Owner='{owner_name}' | Portal='{target['portal']}' | jumlah outlet={len(target_rows)}.")
            else:
                state = "failed"
                stage, message = _portal_error_context(
                    captured_lines, target["portal"], rc, allow_unscoped=len(targets) == 1
                )
                if rc == 0 and not any(target["portal"].casefold() in line.casefold() for line in captured_lines):
                    message = "Tidak ada hasil live baru untuk portal ini; hasil cache lama tidak dipakai."
                log(pct, f"[WEEKLY][PORTAL][DETAIL] Berhasil=0 | Gagal=tidak diketahui | Tahap={stage}.")
                log(pct, f"[WEEKLY][PORTAL][FAILED] Owner='{owner_name}' | Portal='{target['portal']}' | Tahap={stage} | Error={message}")

            for portal in target["portals"]:
                result["portal_results"].append({
                    "owner": owner_name,
                    "aplikator": app_label,
                    "portal": portal,
                    "account_label": _mask_account_label(target.get("account")),
                    "status": state,
                    "stage": stage,
                    "outlet_count": len(target_rows),
                    "message": message,
                })
        return all_rows

    if cancel_event and cancel_event.is_set():
        return result

    missing_portals = _weekly_missing_portals(dbr_df, owner_name)
    for app_key, app_label, pct in (("gofood", "GoFood", 18), ("grab", "GrabFood", 48), ("shopee", "ShopeeFood", 72)):
        for missing_label in missing_portals[app_key]:
            message = "Kolom Nama Portal kosong pada baris DBR."
            result["portal_results"].append({
                "owner": owner_name, "aplikator": app_label, "portal": missing_label,
                "account_label": "(tidak tersedia)", "status": "failed", "stage": "validasi DBR",
                "outlet_count": 0, "message": message,
            })
            log(pct, f"[WEEKLY][PORTAL][FAILED] Owner='{owner_name}' | Aplikator={app_label} | DBR portal='{missing_label}' | Tahap=validasi DBR | Error={message}")

    go_targets = _weekly_owner_targets(dbr_df, owner_name, "gofood")
    if go_targets:
        result["gofood"] = 0
        go_with_creds = [t for t in go_targets if t.get("accounts")]
        missing_creds = [t for t in go_targets if not t.get("accounts")]
        for target in missing_creds:
            message = "Kredensial email GoFood tidak ditemukan pada DBR."
            log(18, f"[WEEKLY][PORTAL][FAILED] Owner='{owner_name}' | Portal='{target['portal']}' | Tahap=validasi DBR | Error={message}")
            result["portal_results"].append({
                "owner": owner_name, "aplikator": "GoFood", "portal": target["portal"],
                "account_label": "(kosong)", "status": "failed", "stage": "validasi DBR",
                "outlet_count": 0, "message": message,
            })
        if go_with_creds:
            headless = os.getenv("HEADLESS_GOFOOD", os.getenv("HEADLESS", "true")).strip().lower() in ("true", "1", "yes", "y")
            if not headless and not os.getenv("DISPLAY") and sys.platform.startswith("linux"):
                headless = True
            run_platform(
                "gofood", "GoFood", GOFOOD_DIR / "gofood_scraper.py", GOFOOD_DIR / "cache", "gofood_portal_",
                ["--type", "dbr", "--owner", owner_name, "--all", "--headless" if headless else "--gui", "--no-master", "--no-upload"],
                go_with_creds, 18, 1200,
            )
    else:
        result["gofood"] = 0

    if cancel_event and cancel_event.is_set():
        return result

    grab_targets = _weekly_owner_targets(dbr_df, owner_name, "grab")
    if grab_targets:
        result["grab"] = 0
        grab_with_creds = [t for t in grab_targets if t.get("accounts") and t.get("has_password")]
        missing_creds = [t for t in grab_targets if not t.get("accounts") or not t.get("has_password")]
        for target in missing_creds:
            missing = "username" if not target.get("accounts") else "kata sandi"
            message = f"Kredensial {missing} GrabFood tidak ditemukan pada DBR."
            log(48, f"[WEEKLY][PORTAL][FAILED] Owner='{owner_name}' | Portal='{target['portal']}' | Tahap=validasi DBR | Error={message}")
            result["portal_results"].append({
                "owner": owner_name, "aplikator": "GrabFood", "portal": target["portal"],
                "account_label": "(kosong)", "status": "failed", "stage": "validasi DBR",
                "outlet_count": 0, "message": message,
            })
        if grab_with_creds:
            headless = os.getenv("HEADLESS_GRAB", os.getenv("HEADLESS", "true")).strip().lower() in ("true", "1", "yes", "y")
            if not headless and not os.getenv("DISPLAY") and sys.platform.startswith("linux"):
                headless = True
            run_platform(
                "grab", "GrabFood", GRAB_DIR / "grab_merchant_scraper.py", GRAB_DIR / "cache", "grab_portal_",
                ["--type", "dbr", "--owner", owner_name, "--fresh", "--headless" if headless else "--gui", "--no-master"],
                grab_with_creds, 48, 1200,
            )
    else:
        result["grab"] = 0

    if cancel_event and cancel_event.is_set():
        return result

    shopee_targets = _weekly_owner_targets(dbr_df, owner_name, "shopee")
    result["shopee"] = 0
    if shopee_targets:
        headless = os.getenv("HEADLESS_SHOPEE", os.getenv("HEADLESS", "true")).strip().lower() in ("true", "1", "yes", "y")
        if not headless and not os.getenv("DISPLAY") and sys.platform.startswith("linux"):
            headless = True
        with tempfile.TemporaryDirectory(prefix="weekly-shopee-") as temp_dir:
            for index, target in enumerate(shopee_targets, start=1):
                if cancel_event and cancel_event.is_set():
                    break
                portal = target["portal"]
                account_label = "allvbadmin"
                log(72, f"[WEEKLY][PORTAL][START] Owner='{owner_name}' | Aplikator=ShopeeFood | Portal='{portal}' | Akun utama={account_label} | {index}/{len(shopee_targets)}.")
                log(72, f"[WEEKLY][PORTAL][AUTH] Memulai autentikasi/pemulihan sesi Shopee untuk portal '{portal}'.")
                output_path = Path(temp_dir) / f"shopee_{index}.xlsx"
                captured_lines = []
                rc = run_subprocess_stream(
                    [
                        python_bin, "-u", str(SHOPEE_DIR / "pull_outlet_info.py"),
                        "--merchant-name", portal, "--output", str(output_path),
                        "--headless" if headless else "--gui",
                    ],
                    cwd=SHOPEE_DIR,
                    keywords=("[SHOPEE]", "login", "auth", "merchant", "store", "outlet", "detail", "gagal", "error", "berhasil", "switch", "scrape"),
                    on_log=lambda line: log_child(74, "ShopeeFood", line),
                    timeout_sec=600,
                    cancel_event=cancel_event,
                    redactions=redactions,
                    on_line=captured_lines.append,
                )
                result["shopee_merchants"][portal] = rc
                target_rows = []
                if rc == 0 and output_path.exists():
                    try:
                        df_result = pd.read_excel(output_path, sheet_name="Listing")
                        target_rows = df_result.to_dict(orient="records")
                    except Exception:
                        target_rows = []
                if target_rows:
                    for row in target_rows:
                        row["Nama Pemilik"] = owner_name
                        row["Nama Brand"] = target["brand"]
                        row["Aplikator"] = "ShopeeFood"
                        row["Nama Portal"] = portal
                    result["live_rows"].extend(target_rows)
                    log(76, f"[WEEKLY][PORTAL][AUTH_SUCCESS] Sesi akun utama allvbadmin aktif untuk '{portal}'.")
                    log(76, f"[WEEKLY][PORTAL][SELECTED] Merchant/portal '{portal}' berhasil dipilih.")
                    log(76, f"[WEEKLY][PORTAL][STORES] Store ditemukan={len(target_rows)}.")
                    log(76, f"[WEEKLY][PORTAL][DETAIL] Berhasil={len(target_rows)} | Gagal=0.")
                    result["portal_results"].append({
                        "owner": owner_name, "aplikator": "ShopeeFood", "portal": portal,
                        "account_label": account_label, "status": "success", "stage": "detail outlet",
                        "outlet_count": len(target_rows), "message": f"{len(target_rows)} outlet berhasil ditarik.",
                    })
                    log(76, f"[WEEKLY][PORTAL][SUCCESS] Owner='{owner_name}' | Portal='{portal}' | jumlah outlet={len(target_rows)}.")
                else:
                    stage, message = _portal_error_context(captured_lines, portal, rc, allow_unscoped=True)
                    if rc == 0 and not any(portal.casefold() in line.casefold() for line in captured_lines):
                        message = "Tidak ada hasil live baru untuk portal ini."
                    log(76, f"[WEEKLY][PORTAL][DETAIL] Berhasil=0 | Gagal=tidak diketahui | Tahap={stage}.")
                    result["portal_results"].append({
                        "owner": owner_name, "aplikator": "ShopeeFood", "portal": portal,
                        "account_label": account_label, "status": "failed", "stage": stage,
                        "outlet_count": 0, "message": message,
                    })
                    log(76, f"[WEEKLY][PORTAL][FAILED] Owner='{owner_name}' | Portal='{portal}' | Tahap={stage} | Error={message}")
        result["shopee"] = 1 if any(item["aplikator"] == "ShopeeFood" and item["status"] == "failed" for item in result["portal_results"]) else 0

    # Only current-run results can enter Listing and SOT. Keep output unique by platform/store.
    if result["live_rows"]:
        live_df = pd.DataFrame(result["live_rows"])
        if "Store ID" in live_df.columns and "Aplikator" in live_df.columns:
            has_id = live_df["Store ID"].fillna("").astype(str).str.strip().ne("")
            with_id = live_df[has_id].drop_duplicates(["Aplikator", "Store ID"], keep="last")
            without_id = live_df[~has_id]
            result["live_rows"] = pd.concat([with_id, without_id], ignore_index=True).to_dict(orient="records")
    failed = sum(1 for item in result["portal_results"] if item["status"] == "failed")
    log(78, f"[WEEKLY][OWNER][RESULT] Owner='{owner_name}' | Outlet live={len(result['live_rows'])} | Portal gagal={failed}.")
    return result


def run_live_scraping_for_owner(owner_name, aplikator="all", progress_cb=None, retry_targets=None, cancel_event=None, dbr_df=None):
    """
    Menjalankan live scraping untuk GoFood, GrabFood, dan ShopeeFood secara berurutan.
    progress_cb: function(percent: int, log_line: str)
    retry_targets: list of dicts (jika hanya menargetkan outlet/aplikator tertentu yang gagal)
    cancel_event: threading.Event opsional untuk membatalkan proses secara langsung
    Mengembalikan dict status hasil scraping per-aplikator.
    """
    if dbr_df is not None:
        return _run_weekly_dbr_scraping(owner_name, dbr_df, progress_cb=progress_cb, cancel_event=cancel_event)

    clean_owner = owner_name.strip()
    python_bin = sys.executable
    scrape_status = {
        "gofood": None,
        "grab": None,
        "shopee": None,
        "shopee_merchants": {},
        "detector_warnings": []
    }

    # Bersihkan file cache detektor lama
    for alert_f in [
        GOFOOD_DIR / "cache" / "gofood_detector_alerts.json",
        GRAB_DIR / "cache" / "grab_detector_alerts.json",
        SHOPEE_DIR / "data" / "shopee_detector_alerts.json",
    ]:
        if alert_f.exists():
            try:
                os.remove(alert_f)
            except Exception:
                pass

    def send_log(pct, msg):
        if progress_cb:
            progress_cb(pct, msg)
        print(f"[{pct:>3}%] {msg}")

    # Normalisasi target aplikator (mendukung 'all', list ['gofood', 'grab'], atau string 'gofood,grab')
    if isinstance(aplikator, (list, tuple, set)):
        active_apps = {str(a).strip().lower() for a in aplikator if str(a).strip()}
    elif isinstance(aplikator, str):
        active_apps = {a.strip().lower() for a in aplikator.split(",") if a.strip()}
    else:
        active_apps = {"all"}
    is_all_apps = ("all" in active_apps) or (not active_apps)

    # Cek apakah aplikator masuk dalam target eksekusi
    def should_run_app(app_key):
        if retry_targets is None:
            return is_all_apps or (app_key in active_apps)
        target_apps = [str(t.get("aplikator", "")).lower() for t in retry_targets]
        if app_key == "gofood" and any("go" in a for a in target_apps):
            return True
        if app_key == "grab" and any("grab" in a for a in target_apps):
            return True
        if app_key == "shopee" and any("shopee" in a for a in target_apps):
            return True
        return False

    # Vercel masih menjadi referensi GoFood/Grab; target merchant Shopee berasal dari DBR.
    if should_run_app("gofood") or should_run_app("grab"):
        load_vercel_data(force_live=True)

    # 1. LIVE SCRAPING GOFOOD
    if should_run_app("gofood"):
        if cancel_event and cancel_event.is_set():
            send_log(100, "🛑 Penarikan dibatalkan oleh pengguna.")
            return scrape_status
        send_log(15, f"🚀 [GoFood] Memulai penarikan live untuk '{clean_owner}'...")
        headless_go = os.getenv("HEADLESS_GOFOOD", os.getenv("HEADLESS", "true")).strip().lower() in ("true", "1", "yes", "y")
        if not headless_go and not os.getenv("DISPLAY") and sys.platform.startswith("linux"):
            headless_go = True
        cmd = [
            python_bin,
            "-u",
            str(GOFOOD_DIR / "gofood_scraper.py"),
            "--vercel",
            "--owner", clean_owner,
            "--headless" if headless_go else "--gui",
            "--no-master",
            "--no-upload"
        ]
        rc = run_subprocess_stream(
            cmd,
            cwd=GOFOOD_DIR,
            keywords=("Store ID", "Berhasil", "Portal", "Login", "Owner", "Restricted", "Memproses", "OTP", "Filter", "Gagal", "PERINGATAN", "DETEKTOR", "GAGAL_LOGIN", "DATA_KOSONG"),
            on_log=lambda m: send_log(25, f"[GoFood] {m}"),
            timeout_sec=160,
            cancel_event=cancel_event
        )
        scrape_status["gofood"] = rc
        if cancel_event and cancel_event.is_set():
            send_log(100, "🛑 Penarikan dibatalkan oleh pengguna.")
            return scrape_status
        if rc == 0:
            send_log(35, f"✅ [GoFood] Selesai memproses '{clean_owner}'.")
        else:
            send_log(35, f"⚠️ [GoFood] Selesai dengan kode status {rc}.")

    # 2. LIVE SCRAPING GRABFOOD
    if should_run_app("grab"):
        if cancel_event and cancel_event.is_set():
            send_log(100, "🛑 Penarikan dibatalkan oleh pengguna.")
            return scrape_status
        send_log(40, f"🚀 [GrabFood] Memulai penarikan live untuk '{clean_owner}'...")
        headless_grab = os.getenv("HEADLESS_GRAB", os.getenv("HEADLESS", "true")).strip().lower() in ("true", "1", "yes", "y")
        if not headless_grab and not os.getenv("DISPLAY") and sys.platform.startswith("linux"):
            headless_grab = True
        cmd = [
            python_bin,
            "-u",
            str(GRAB_DIR / "grab_merchant_scraper.py"),
            "--vercel",
            "--owner", clean_owner,
            "--headless" if headless_grab else "--gui",
            "--no-master"
        ]
        rc = run_subprocess_stream(
            cmd,
            cwd=GRAB_DIR,
            keywords=("Group ID", "Berhasil", "Target", "Portal", "Login", "Store", "Bank", "Owner", "Filter", "Gagal", "Password", "Salah", "Master@123", "fallback", "Sukses", "PERINGATAN", "DETEKTOR", "GAGAL_LOGIN", "DATA_KOSONG"),
            on_log=lambda m: send_log(55, f"[Grab] {m}"),
            timeout_sec=180,
            cancel_event=cancel_event
        )
        scrape_status["grab"] = rc
        if cancel_event and cancel_event.is_set():
            send_log(100, "🛑 Penarikan dibatalkan oleh pengguna.")
            return scrape_status
        if rc == 0:
            send_log(65, f"✅ [GrabFood] Selesai memproses '{clean_owner}'.")
        else:
            send_log(65, f"⚠️ [GrabFood] Selesai dengan kode status {rc}.")

    # 3. LIVE SCRAPING SHOPEEFOOD
    if should_run_app("shopee"):
        if cancel_event and cancel_event.is_set():
            send_log(100, "🛑 Penarikan dibatalkan oleh pengguna.")
            return scrape_status
        send_log(68, f"🚀 [ShopeeFood] Memeriksa akun Shopee untuk '{clean_owner}'...")
        try:
            from cronjobsot.export_to_sot import load_dbr_master_data

            dbr_df = load_dbr_master_data()
            if dbr_df.empty or not {"Nama Pemilik", "Aplikator", "Nama Portal"}.issubset(dbr_df.columns):
                raise ValueError("Tab DBR kosong atau tidak memiliki kolom Nama Pemilik, Aplikator, dan Nama Portal")

            owner_names = dbr_df["Nama Pemilik"].fillna("").astype(str).str.strip().str.lower()
            app_names = dbr_df["Aplikator"].fillna("").astype(str).str.strip().str.lower()
            o_shopee = dbr_df[(owner_names == clean_owner.lower()) & app_names.str.contains("shopee")]
            merchant_targets = []
            if not o_shopee.empty:
                known_shopee = get_known_shopee_merchants()
                import re

                def _norm(s):
                    return re.sub(r'[^a-z0-9]', '', re.sub(r'\[.*?\]', '', (s or '').lower()))

                for portal_value in o_shopee["Nama Portal"].dropna().astype(str):
                    portal_name = portal_value.strip()
                    if not portal_name or portal_name.lower() in ("nan", "none", "-"):
                        continue

                    best_match = None
                    if known_shopee:
                        # Canonicalize portal dari DBR ke nama merchant yang dikenal scraper.
                        c_clean = portal_name.lower().rstrip("_").strip()
                        for k in known_shopee:
                            if c_clean == k.lower().rstrip("_").strip():
                                best_match = k
                                break

                        # 2. Normalized exact match
                        if not best_match:
                            c_norm = _norm(portal_name)
                            if c_norm:
                                for k in known_shopee:
                                    if c_norm == _norm(k):
                                        best_match = k
                                        break

                        # 3. Substring match
                        if not best_match:
                            c_norm = _norm(portal_name)
                            if len(c_norm) >= 4:
                                for k in known_shopee:
                                    k_norm = _norm(k)
                                    if len(k_norm) >= 4 and (c_norm in k_norm or k_norm in c_norm):
                                        best_match = k
                                        break

                    chosen_val = best_match or portal_name

                    # Filter spesifik jika retry_targets ditentukan
                    if retry_targets is not None:
                        shopee_targets = [str(t.get("portal") or t.get("outlet") or "").strip().lower() for t in retry_targets if "shopee" in str(t.get("aplikator", "")).lower()]
                        if not any(st in chosen_val.lower() or chosen_val.lower() in st for st in shopee_targets if st):
                            continue

                    if chosen_val and chosen_val not in merchant_targets:
                        merchant_targets.append(chosen_val)
            
            if merchant_targets:
                headless_shopee = os.getenv("HEADLESS_SHOPEE", os.getenv("HEADLESS", "true")).strip().lower() in ("true", "1", "yes", "y")
                if not headless_shopee and not os.getenv("DISPLAY") and sys.platform.startswith("linux"):
                    headless_shopee = True

                for idx, merchant_name in enumerate(merchant_targets, start=1):
                    if cancel_event and cancel_event.is_set():
                        send_log(100, "🛑 Penarikan ShopeeFood dibatalkan oleh pengguna.")
                        break
                    send_log(70, f"[SHOPEE][PORTAL][START] Owner='{clean_owner}' | Portal='{merchant_name}' | Akun utama: allvbadmin | Portal {idx}/{len(merchant_targets)}.")
                    cmd = [
                        python_bin,
                        "-u",
                        str(SHOPEE_DIR / "pull_outlet_info.py"),
                        "--merchant-name", merchant_name,
                        "--headless" if headless_shopee else "--gui"
                    ]

                    rc_sh = run_subprocess_stream(
                        cmd,
                        cwd=SHOPEE_DIR,
                        keywords=("[SHOPEE]",),
                        on_log=lambda m: send_log(74, m) if "[SHOPEE]" in m or "batas waktu" in m.lower() else None,
                        timeout_sec=180,
                        cancel_event=cancel_event
                    )
                    scrape_status["shopee_merchants"][merchant_name] = rc_sh
                scrape_status["shopee"] = 0 if all(c == 0 for c in scrape_status["shopee_merchants"].values()) else 1
                send_log(76, f"✅ [ShopeeFood] Selesai memproses {len(merchant_targets)} merchant.")
            else:
                send_log(76, f"ℹ️ [ShopeeFood] Tidak ditemukan nama merchant Shopee untuk '{clean_owner}'.")
        except Exception as e:
            scrape_status["shopee"] = 1
            send_log(76, f"⚠️ [ShopeeFood] Exception scraper: {e}")

    # Kumpulkan seluruh peringatan detektor dari semua platform
    detector_warnings = []
    for alert_path in [
        GOFOOD_DIR / "cache" / "gofood_detector_alerts.json",
        GRAB_DIR / "cache" / "grab_detector_alerts.json",
        SHOPEE_DIR / "data" / "shopee_detector_alerts.json",
    ]:
        if alert_path.exists():
            try:
                with open(alert_path, "r", encoding="utf-8") as f:
                    items = json.load(f)
                    if isinstance(items, list):
                        detector_warnings.extend(items)
            except Exception:
                pass
    scrape_status["detector_warnings"] = detector_warnings
    if detector_warnings:
        send_log(77, f"🚨 Ditemukan {len(detector_warnings)} peringatan detektor dari platform!")

    return scrape_status


def verify_extraction_completeness(expected_df, owner_df, scrape_status=None):
    """
    Membandingkan daftar outlet yang diharapkan dari Vercel Sheet dengan hasil akhir di owner_df.
    Mengidentifikasi outlet yang gagal ditarik atau memiliki data Store ID kosong.
    """
    if expected_df.empty:
        return []

    scrape_status = scrape_status or {}
    missing_items = []
    matched_act_indices = set()

    def norm(s):
        import re
        return re.sub(r'[^a-z0-9]', '', str(s or '').lower())

    for _, exp_row in expected_df.iterrows():
        app = str(exp_row.get("Aplikator") or "").strip()
        outlet_name = str(exp_row.get("Nama Outlet") or exp_row.get("Nama Brand") or "").strip()
        portal_name = str(exp_row.get("Nama Portal") or exp_row.get("Nama Akses") or exp_row.get("Merchant Name") or "").strip()
        user_name = str(exp_row.get("Nama Pengguna") or "").strip()

        matched_rows = owner_df[owner_df["Aplikator"].astype(str).str.lower() == app.lower()] if "Aplikator" in owner_df.columns else pd.DataFrame()

        found_valid = False
        if not matched_rows.empty:
            norm_exp_portal = norm(portal_name)
            norm_exp_outlet = norm(outlet_name)
            norm_exp_user = norm(user_name)

            for act_idx, act_row in matched_rows.iterrows():
                if act_idx in matched_act_indices:
                    continue

                store_id = str(act_row.get("Store ID") or "").strip()
                has_store_id = bool(store_id and store_id.lower() not in ("nan", "none", "-", ""))

                act_portal = norm(act_row.get("Nama Portal") or "")
                act_outlet = norm(act_row.get("Nama Outlet") or act_row.get("Nama Brand") or act_row.get("Nama Listing") or "")
                act_user = norm(act_row.get("Nama Pengguna") or "")

                match_identity = False
                # Prioritaskan pencocokan outlet, lalu portal
                if norm_exp_outlet and (norm_exp_outlet == act_outlet or (len(norm_exp_outlet) >= 4 and (norm_exp_outlet in act_outlet or act_outlet in norm_exp_outlet))):
                    match_identity = True
                elif norm_exp_portal and (norm_exp_portal == act_portal or (len(norm_exp_portal) >= 4 and (norm_exp_portal in act_portal or act_portal in norm_exp_portal))):
                    match_identity = True
                elif norm_exp_user and norm_exp_user == act_user:
                    match_identity = True
                elif len(matched_rows) == 1 and len(expected_df[expected_df["Aplikator"].astype(str).str.lower() == app.lower()]) == 1:
                    match_identity = True

                if match_identity and has_store_id:
                    found_valid = True
                    matched_act_indices.add(act_idx)
                    break

        if not found_valid:
            reason = "Store ID kosong / data toko tidak lengkap"
            app_lower = app.lower()

            # Periksa apakah ada catatan peringatan detektor dari platform
            for w in scrape_status.get("detector_warnings", []):
                w_plat = str(w.get("platform", "")).lower()
                w_portal = str(w.get("portal", "")).lower()
                w_user = str(w.get("username", "") or w.get("email", "")).lower()
                if (w_plat in app_lower) or (app_lower in w_plat):
                    match_alert = False
                    if norm(w_portal) and (norm(w_portal) in norm(portal_name) or norm(portal_name) in norm(w_portal)):
                        match_alert = True
                    elif norm(w_user) and (norm(w_user) in norm(user_name) or norm(user_name) in norm(w_user)):
                        match_alert = True
                    elif norm(w_portal) and (norm(w_portal) in norm(outlet_name) or norm(outlet_name) in norm(w_portal)):
                        match_alert = True

                    if match_alert:
                        if w.get("tipe") == "GAGAL_LOGIN":
                            reason = f"Gagal login/auth ({w.get('pesan', 'Kredensial/OTP bermasalah')})"
                        elif w.get("tipe") == "DATA_KOSONG":
                            reason = "Data outlet kosong di portal (0 outlet)"
                        else:
                            reason = w.get("pesan", reason)
                        break

            if reason == "Store ID kosong / data toko tidak lengkap":
                if "gofood" in app_lower and scrape_status.get("gofood") not in (None, 0):
                    reason = f"Proses GoFood scraper keluar status {scrape_status.get('gofood')}"
                elif "grab" in app_lower and scrape_status.get("grab") not in (None, 0):
                    reason = f"Proses Grab scraper keluar status {scrape_status.get('grab')}"
                elif "shopee" in app_lower:
                    sh_details = scrape_status.get("shopee_merchants", {})
                    for m_name, m_rc in sh_details.items():
                        if norm(m_name) in norm(portal_name) or norm(portal_name) in norm(m_name):
                            if m_rc != 0:
                                reason = f"Gagal switch merchant Shopee '{m_name}'"
                            break

            missing_items.append({
                "aplikator": app,
                "outlet": outlet_name or portal_name or "Outlet Tanpa Nama",
                "portal": portal_name or outlet_name or "-",
                "username": user_name,
                "reason": reason
            })

    return missing_items


def generate_for_owner_pipeline(owner_name, aplikator="all", upload=True, source="vercel", live_scrape=True, progress_callback=None, lock_timeout=0, requested_by="System", retry_targets=None, cancel_event=None, dbr_df=None):
    """
    Pipeline pembuatan file per-owner dan upload Drive dengan progress callback.
    Menggunakan Vercel Sheet sebagai sumber utama dan memperkaya data dengan hasil live scraping.
    Dilindungi oleh JobLock untuk mencegah tabrakan proses / race conditions.
    
    aplikator: 'all' | 'gofood' | 'grab' | 'shopee'
    source: 'vercel' (default) | 'master'
    live_scrape: True (default) menjalankan scraping live | False (kompilasi cache saja)
    progress_callback: function(percent: int, log_line: str)
    lock_timeout: int (detik untuk menunggu lock, 0 = fail-fast)
    requested_by: str (nama user / caller)
    cancel_event: threading.Event opsional untuk pembatalan cepat
    """
    def log(pct, msg):
        if progress_callback:
            progress_callback(pct, msg)
        print(f"[{pct:>3}%] {msg}")

    try:
        with JobLock(timeout=lock_timeout, task_info={"owner": owner_name, "user": requested_by, "aplikator": aplikator, "source": source}):
            if cancel_event and cancel_event.is_set():
                log(100, "🛑 Pipeline dibatalkan sebelum memulai.")
                return {"success": False, "error": "CANCELLED", "message": "Dibatalkan oleh pengguna."}

            log(10, "Inisialisasi direktori dan template listing...")
            headers = get_template_headers()
            OUTPUT_OWNERS_DIR.mkdir(parents=True, exist_ok=True)

            # 1. LIVE SCRAPING (Jika diaktifkan dan bukan mode __ALL__)
            scrape_status = {}
            if live_scrape and owner_name != "__ALL__":
                log(12, f"Menjalankan penarikan live data toko untuk '{owner_name}'...")
                scrape_status = run_live_scraping_for_owner(
                    owner_name,
                    aplikator=aplikator,
                    progress_cb=progress_callback,
                    retry_targets=retry_targets,
                    cancel_event=cancel_event,
                    dbr_df=dbr_df if source == "dbr" else None,
                )

            if cancel_event and cancel_event.is_set():
                log(100, "🛑 Pipeline dibatalkan oleh pengguna.")
                return {"success": False, "error": "CANCELLED", "message": "Dibatalkan oleh pengguna."}

            # Normalisasi aplikator target
            if isinstance(aplikator, (list, tuple, set)):
                active_apps = {str(a).strip().lower() for a in aplikator if str(a).strip()}
            elif isinstance(aplikator, str):
                active_apps = {a.strip().lower() for a in aplikator.split(",") if a.strip()}
            else:
                active_apps = {"all"}
            is_all_apps = ("all" in active_apps) or (not active_apps)
            app_display = "ALL" if is_all_apps else "+".join(sorted(active_apps)).upper()

            log(78, f"Membaca data sumber ({source.upper()}) untuk aplikator: {app_display}...")
            if source == "dbr":
                src_df = dbr_df.copy() if isinstance(dbr_df, pd.DataFrame) else pd.DataFrame()
            elif source == "vercel":
                src_df = load_vercel_data()
            else:
                dfs = []
                if is_all_apps or any(k in active_apps for k in ("gofood", "go")):
                    df_go = load_gofood_data()
                    if not df_go.empty:
                        dfs.append(df_go)
                if is_all_apps or any(k in active_apps for k in ("grab", "grabfood")):
                    df_grab = load_grab_data()
                    if not df_grab.empty:
                        dfs.append(df_grab)
                if is_all_apps or any(k in active_apps for k in ("shopee", "shopeefood")):
                    df_shopee = load_shopee_data()
                    if not df_shopee.empty:
                        dfs.append(df_shopee)
                src_df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

            if src_df.empty:
                log(100, "❌ Data sumber kosong!")
                return {"success": False, "error": "Data sumber kosong"}

            log(80, f"Memfilter data untuk Owner: '{owner_name}'...")
            if owner_name == "__ALL__":
                owner_df = src_df.copy()
            else:
                owner_df = src_df[src_df["Nama Pemilik"].astype(str).str.strip().str.lower() == owner_name.strip().lower()].copy()

            if owner_df.empty:
                log(100, f"⚠️ Tidak ditemukan data untuk owner '{owner_name}'.")
                return {"success": False, "error": f"Owner '{owner_name}' tidak ditemukan dalam data"}

            # Filter aplikator
            if not is_all_apps:
                target_names = []
                if any(k in active_apps for k in ("gofood", "go")):
                    target_names.append("GoFood")
                if any(k in active_apps for k in ("grab", "grabfood")):
                    target_names.append("GrabFood")
                if any(k in active_apps for k in ("shopee", "shopeefood")):
                    target_names.append("ShopeeFood")
                owner_df = owner_df[owner_df["Aplikator"].isin(target_names)]

            if owner_df.empty:
                log(100, f"⚠️ Tidak ada outlet {app_display} untuk owner '{owner_name}'.")
                return {"success": False, "error": f"Tidak ada outlet {app_display} untuk owner {owner_name}"}

            expected_df = owner_df.copy()

            # Enrich data dari hasil scraping jika tersedia (Bank, Rekening, Store ID, Alamat)
            log(82, "Menggabungkan hasil penarikan data toko (Store ID, Alamat, Bank)...")
            try:
                final_rows = []
                if source == "dbr":
                    scraped_owner = pd.DataFrame(scrape_status.get("live_rows", []))
                    if not scraped_owner.empty:
                        if "Nama Pemilik" in scraped_owner.columns:
                            scraped_owner = scraped_owner[
                                scraped_owner["Nama Pemilik"].fillna("").astype(str).str.strip().str.casefold()
                                == owner_name.strip().casefold()
                            ]
                        for _, s_row in scraped_owner.iterrows():
                            row_dict = s_row.to_dict()
                            app = str(row_dict.get("Aplikator") or "").strip()
                            portal = str(row_dict.get("Nama Portal") or "").strip().casefold()
                            app_candidates = owner_df[owner_df["Aplikator"].astype(str).str.casefold() == app.casefold()]
                            dbr_target = None
                            if portal and not app_candidates.empty:
                                for _, candidate in app_candidates.iterrows():
                                    portal_values = [candidate.get(c) for c in ("Nama Portal", "Nama Akses", "Nama Outlet", "Nama Brand")]
                                    candidate_keys = [str(v or "").strip().casefold() for v in portal_values if str(v or "").strip()]
                                    if any(portal == key or portal in key or key in portal for key in candidate_keys):
                                        dbr_target = candidate
                                        break
                            if dbr_target is None and len(app_candidates) == 1:
                                dbr_target = app_candidates.iloc[0]
                            row_dict["Nama Pemilik"] = owner_name
                            if dbr_target is not None:
                                for col in ["Nama Brand", "Nama Outlet", "Outlet", "Tipe", "Model", "Nomor HP", "Nama Akses",
                                            "Email FoodMaster1", "Email FoodMaster2", "Nama Pengguna", "Kata Sandi", "Nama Portal",
                                            "S Nomor HP Akses Pemilik", "S Username Akses Pemilik", "S Kata Sandi Akses Pemilik",
                                            "S Allvbadmin Username Akses Staff", "S Allvbadmin Kata Sandi Akses Staff",
                                            "S Bot Username Akses Staff", "S Bot Kata Sandi Akses Staff",
                                            "S BD Username Akses Staff", "S BD Kata Sandi Akses Staff", "BD", "Status Internal"]:
                                    value = dbr_target.get(col)
                                    if pd.notna(value) and str(value).strip() and (col not in row_dict or not str(row_dict.get(col) or "").strip()):
                                        row_dict[col] = value
                            if not str(row_dict.get("Nama Brand") or "").strip():
                                row_dict["Nama Brand"] = row_dict.get("Nama Outlet") or ""
                            final_rows.append(row_dict)
                else:
                    df_go = load_gofood_data()
                    df_gr = load_grab_data()
                    df_sh = load_shopee_data()
                    scraped_list = [d for d in [df_go, df_gr, df_sh] if not d.empty]
                    scraped_all = pd.concat(scraped_list, ignore_index=True) if scraped_list else pd.DataFrame()

                    scraped_owner = pd.DataFrame()
                    if not scraped_all.empty and "Nama Pemilik" in scraped_all.columns:
                        scraped_owner = scraped_all[scraped_all["Nama Pemilik"].astype(str).str.strip().str.lower() == owner_name.strip().lower()]

                    app_order_std = ["GoFood", "GrabFood", "ShopeeFood"]
                    unique_apps = list(owner_df["Aplikator"].dropna().unique())
                    sorted_apps = [a for a in app_order_std if a in unique_apps] + [a for a in unique_apps if a not in app_order_std]
                    for app in sorted_apps:
                        app_vercel_rows = owner_df[owner_df["Aplikator"] == app]
                        app_scraped_rows = scraped_owner[scraped_owner["Aplikator"] == app] if not scraped_owner.empty and "Aplikator" in scraped_owner.columns else pd.DataFrame()

                        if not app_scraped_rows.empty:
                            v_first = app_vercel_rows.iloc[0]
                            for _, s_row in app_scraped_rows.iterrows():
                                row_dict = s_row.to_dict()
                                s_portal = str(s_row.get("Nama Portal") or "").strip().lower()
                                v_target = v_first
                                for _, v_candidate in app_vercel_rows.iterrows():
                                    cand_keys = [
                                        str(v_candidate.get("Merchant Name") or "").strip().lower(),
                                        str(v_candidate.get("Nama Portal") or "").strip().lower(),
                                        str(v_candidate.get("Nama Akses") or "").strip().lower(),
                                        str(v_candidate.get("Nama Outlet") or "").strip().lower()
                                    ]
                                    if s_portal in cand_keys or any(k and k in s_portal for k in cand_keys if len(k) > 3):
                                        v_target = v_candidate
                                        break

                                nama_outlet_val = str(v_target.get("Nama Outlet") or "").strip()
                                if nama_outlet_val:
                                    row_dict["Nama Outlet"] = nama_outlet_val
                                    row_dict["Nama Brand"] = nama_outlet_val
                                for col in ["Nama Akses", "Email FoodMaster1", "Email FoodMaster2", "Nama Pengguna", "Kata Sandi",
                                            "Nama Portal", "Nama Portal.1", "S Nomor HP Akses Pemilik", "S Username Akses Pemilik", "S Kata Sandi Akses Pemilik",
                                            "S Allvbadmin Username Akses Staff", "S Allvbadmin Kata Sandi Akses Staff",
                                            "S Bot Username Akses Staff", "S Bot Kata Sandi Akses Staff",
                                            "S BD Username Akses Staff", "S BD Kata Sandi Akses Staff", "BD", "Status Internal"]:
                                    if col in v_target and pd.notna(v_target[col]) and str(v_target[col]).strip() != "":
                                        if col not in row_dict or pd.isna(row_dict.get(col)) or str(row_dict.get(col)).strip() == "":
                                            row_dict[col] = v_target[col]
                                final_rows.append(row_dict)
                        else:
                            for _, v_row in app_vercel_rows.iterrows():
                                v_dict = v_row.to_dict()
                                if "Nama Outlet" in v_dict and pd.notna(v_dict["Nama Outlet"]):
                                    v_dict["Nama Brand"] = str(v_dict["Nama Outlet"]).strip()
                                final_rows.append(v_dict)

                if final_rows:
                    owner_df = pd.DataFrame(final_rows)
                    if "Nama Outlet" in owner_df.columns and source != "dbr":
                        owner_df["Nama Brand"] = owner_df["Nama Outlet"].astype(str).str.strip()
            except Exception as e:
                print(f"⚠️ Info enrich scraped: {e}")

            if owner_df.empty:
                log(100, f"❌ Tidak ada hasil live sukses untuk owner '{owner_name}'; Listing dan SOT tidak diisi dari cache lama.")
                return {
                    "success": False,
                    "error": "Tidak ada hasil live sukses",
                    "owner": owner_name,
                    "portal_results": scrape_status.get("portal_results", []),
                    "sot_rows": [],
                }

            # Deduplikasi ketat data owner untuk mencegah baris ganda (misal duplikasi input di Vercel atau multiple scrape)
            if not owner_df.empty:
                orig_len = len(owner_df)
                if "Store ID" in owner_df.columns:
                    has_id = (
                        owner_df["Store ID"].notna() 
                        & (owner_df["Store ID"].astype(str).str.strip() != "") 
                        & (owner_df["Store ID"].astype(str).str.strip() != "-")
                    )
                    df_with_id = owner_df[has_id].drop_duplicates(subset=["Aplikator", "Store ID"], keep="last")
                    
                    sub_no_id = [c for c in ["Aplikator", "Nama Outlet", "Nama Pengguna"] if c in owner_df.columns]
                    if not sub_no_id:
                        sub_no_id = ["Aplikator", "Nama Outlet"] if "Nama Outlet" in owner_df.columns else ["Aplikator"]
                    df_no_id = owner_df[~has_id].drop_duplicates(subset=sub_no_id, keep="last")
                    owner_df = pd.concat([df_with_id, df_no_id], ignore_index=True)
                else:
                    sub_cols = [c for c in ["Aplikator", "Nama Outlet", "Nama Pengguna"] if c in owner_df.columns]
                    owner_df = owner_df.drop_duplicates(subset=sub_cols, keep="last")

                if len(owner_df) < orig_len:
                    log(84, f"🧹 Berhasil membersihkan {orig_len - len(owner_df)} baris duplikat dari data akhir.")

            # Pastikan urutan selalu standar: GoFood -> GrabFood -> ShopeeFood
            owner_df = sort_df_by_aplikator(owner_df)
            expected_df = sort_df_by_aplikator(expected_df)

            # DBR mode evaluates portal outcomes directly; other pipelines keep their legacy checks.
            if source == "dbr":
                failed_portals = [item for item in scrape_status.get("portal_results", []) if item.get("status") != "success"]
                missing_items = [
                    {
                        "aplikator": item.get("aplikator", ""),
                        "outlet": item.get("portal", ""),
                        "portal": item.get("portal", ""),
                        "reason": f"{item.get('stage', 'scraping')}: {item.get('message', 'Gagal')}",
                    }
                    for item in failed_portals
                ]
            else:
                missing_items = verify_extraction_completeness(expected_df, owner_df, scrape_status)
            detector_warnings = scrape_status.get("detector_warnings", [])
            is_partial = len(missing_items) > 0 or len(detector_warnings) > 0
            completed_count = (
                sum(1 for item in scrape_status.get("portal_results", []) if item.get("status") == "success")
                if source == "dbr" else max(0, len(expected_df) - len(missing_items))
            )

            if is_partial:
                warn_note = f", {len(detector_warnings)} peringatan platform" if detector_warnings else ""
                log(86, f"⚠️ Selesai dengan catatan: {len(missing_items)} outlet belum lengkap{warn_note}.")
            else:
                log(86, f"✅ Seluruh {len(expected_df)} outlet target berhasil ditarik lengkap.")

            go_n = len(owner_df[owner_df["Aplikator"] == "GoFood"]) if "Aplikator" in owner_df.columns else 0
            gr_n = len(owner_df[owner_df["Aplikator"] == "GrabFood"]) if "Aplikator" in owner_df.columns else 0
            sh_n = len(owner_df[owner_df["Aplikator"] == "ShopeeFood"]) if "Aplikator" in owner_df.columns else 0

            log(85, f"Ditemukan {len(owner_df)} outlet (GoFood: {go_n}, Grab: {gr_n}, Shopee: {sh_n}).")

            log(80, "Menyusun file Excel dengan 2 Tab Identik ('Listing' & 'Listing 2')...")
            timestamp_str = datetime.datetime.now().strftime("%Y-%m-%d %H_%M")
            clean_owner = "".join(c for c in owner_name if c.isalnum() or c in " ._-").strip()
            filename = f"{timestamp_str} {clean_owner}.xlsx"
            file_path = OUTPUT_OWNERS_DIR / filename

            save_owner_workbook(owner_df, str(file_path), headers)
            log(85, f"File lokal berhasil dibuat: {filename}")

            if cancel_event and cancel_event.is_set():
                log(100, "🛑 Pipeline dibatalkan sebelum upload Google Drive.")
                return {"success": False, "error": "CANCELLED", "message": "Dibatalkan oleh pengguna."}

            drive_res = {}
            if upload:
                log(90, f"Mengunggah ke folder Google Drive: [{clean_owner}]...")
                ok, res = upload_file_to_drive(str(file_path), clean_owner, APP_SCRIPT_URL)
                if ok:
                    drive_res = res
                    log(100, f"✅ Sukses terunggah ke Google Drive!")
                else:
                    log(100, f"⚠️ Gagal upload ke Drive: {res.get('error', 'Unknown')}")
                    return {
                        "success": False,
                        "owner": owner_name,
                        "filename": filename,
                        "filepath": str(file_path),
                        "total": len(owner_df),
                        "completed_count": completed_count,
                        "expected_count": len(expected_df),
                        "missing_items": missing_items,
                        "detector_warnings": detector_warnings,
                        "gofood": go_n,
                        "grab": gr_n,
                        "shopee": sh_n,
                        "error": res.get("error"),
                        "portal_results": scrape_status.get("portal_results", []),
                        "sot_rows": weekly_sot_rows(owner_df) if source == "dbr" else [],
                    }
            else:
                log(100, "✅ Selesai (Mode Lokal Saja).")

            return {
                "success": True,
                "status": "PARTIAL" if is_partial else "FULL",
                "is_partial": is_partial,
                "owner": owner_name,
                "filename": filename,
                "filepath": str(file_path),
                "total": len(owner_df),
                "completed_count": completed_count,
                "expected_count": len(expected_df),
                "missing_items": missing_items,
                "detector_warnings": detector_warnings,
                "gofood": go_n,
                "grab": gr_n,
                "shopee": sh_n,
                "portal_results": scrape_status.get("portal_results", []),
                "sot_rows": weekly_sot_rows(owner_df) if source == "dbr" else [],
                "folder_url": drive_res.get("folderUrl") or "https://drive.google.com/drive/u/0/folders/19VIrypPcBmNNbjDLGS7kxp_yIdBBwXjB",
                "file_url": drive_res.get("fileUrl", "")
            }
    except Timeout:
        info = get_lock_info()
        busy_owner = info.get("owner", "Owner lain")
        busy_user = info.get("user", "Pengguna lain")
        msg = f"Pipeline sedang dikunci oleh proses lain (Owner: {busy_owner}, oleh: {busy_user})"
        log(100, f"⚠️ {msg}")
        return {
            "success": False,
            "error": "LOCKED",
            "lock_info": info,
            "message": msg
        }


def process_and_combine(owner_filter=None, upload=False, lock_timeout=0):
    """Proses utama penggabungan per-owner dan upload ke Google Drive."""
    try:
        with JobLock(timeout=lock_timeout, task_info={"task": "process_and_combine", "owner_filter": owner_filter or "__ALL__", "user": "CLI"}):
            print("=" * 65)
            print("🚀 PENGGABUNGAN DATA OUTLET PER-OWNER (GOFOOD, GRAB, SHOPEE)")
            print("=" * 65)

            headers = get_template_headers()
            OUTPUT_OWNERS_DIR.mkdir(parents=True, exist_ok=True)

            print("\n[*] Membaca data master...")
            df_go = load_gofood_data()
            df_grab = load_grab_data()
            df_shopee = load_shopee_data()

            print(f"    - GoFood outlets   : {len(df_go)} baris")
            print(f"    - GrabFood outlets : {len(df_grab)} baris")
            print(f"    - ShopeeFood outlets: {len(df_shopee)} baris")

            dfs = [df for df in [df_go, df_grab, df_shopee] if not df.empty]

            # Jika master lokal belum lengkap atau owner tertentu tidak ada di master lokal,
            # gunakan data acuan dari Google Sheet (Vercel/DBR) agar seluruh owner terjangkau
            df_sheet = load_vercel_data()
            if not df_sheet.empty:
                if not dfs:
                    print(f"    - Google Sheet master: {len(df_sheet)} baris (digunakan sebagai basis)")
                    combined_df = df_sheet
                else:
                    scraped_combined = pd.concat(dfs, ignore_index=True)
                    scraped_owners = set(scraped_combined["Nama Pemilik"].dropna().astype(str).str.strip().str.lower())
                    missing_from_scraped = df_sheet[~df_sheet["Nama Pemilik"].fillna("").astype(str).str.strip().str.lower().isin(scraped_owners)]
                    combined_df = pd.concat([scraped_combined, missing_from_scraped], ignore_index=True)
                    print(f"    - Gabungan master lokal + sheet: {len(combined_df)} baris ({len(scraped_owners)} owner dari scraping, {len(missing_from_scraped)} baris tambahan dari sheet)")
            elif dfs:
                combined_df = pd.concat(dfs, ignore_index=True)
            else:
                print("\n❌ Tidak ada data outlet yang ditemukan.")
                return False

            # Filter baris yang memiliki Nama Pemilik
            if "Nama Pemilik" not in combined_df.columns:
                print("❌ Kolom 'Nama Pemilik' tidak ditemukan dalam data.")
                return False

            combined_df = combined_df[combined_df["Nama Pemilik"].notna()]
            combined_df = combined_df[combined_df["Nama Pemilik"].astype(str).str.strip() != ""]

            all_owners = sorted(combined_df["Nama Pemilik"].unique(), key=lambda x: str(x).lower())
            print(f"\n[✓] Ditemukan {len(all_owners)} unique owner lintas platform.")

            # Filter owner jika ditentukan
            if owner_filter:
                owner_filter_lower = owner_filter.strip().lower()
                filtered = [o for o in all_owners if owner_filter_lower in str(o).lower()]
                if not filtered:
                    print(f"⚠️ Tidak ditemukan owner dengan kata kunci: '{owner_filter}'")
                    print("Daftar owner yang ada:")
                    for o in all_owners[:20]:
                        print(f"  - {o}")
                    if len(all_owners) > 20:
                        print(f"  ... dan {len(all_owners) - 20} lainnya.")
                    return False
                all_owners = filtered
                print(f"[*] Memproses {len(all_owners)} owner yang cocok dengan filter '{owner_filter}'.")

            timestamp_str = datetime.datetime.now().strftime("%Y-%m-%d %H_%M")
            success_count = 0
            upload_count = 0

            print("\n" + "─" * 65)
            for idx, owner in enumerate(all_owners, 1):
                owner_str = str(owner).strip()
                clean_owner = "".join(c for c in owner_str if c.isalnum() or c in " ._-").strip()
                filename = f"{timestamp_str} {clean_owner}.xlsx"
                file_path = OUTPUT_OWNERS_DIR / filename

                owner_df = combined_df[combined_df["Nama Pemilik"] == owner]
                # Deduplikasi per Store ID jika ada duplikat dalam aplikator yang sama
                if "Store ID" in owner_df.columns and "Aplikator" in owner_df.columns:
                    owner_df = owner_df.drop_duplicates(subset=["Aplikator", "Store ID"], keep="first")
                if "Nama Outlet" in owner_df.columns:
                    owner_df["Nama Brand"] = owner_df["Nama Outlet"].astype(str).str.strip()

                go_count = len(owner_df[owner_df["Aplikator"] == "GoFood"])
                grab_count = len(owner_df[owner_df["Aplikator"] == "GrabFood"])
                shopee_count = len(owner_df[owner_df["Aplikator"] == "ShopeeFood"])

                print(f"\n[{idx}/{len(all_owners)}] 👤 Owner: {owner_str}")
                print(f"     📊 Total Outlet: {len(owner_df)} (GoFood: {go_count}, Grab: {grab_count}, Shopee: {shopee_count})")

                # Simpan file dengan 2 tab 100% identik
                save_owner_workbook(owner_df, str(file_path), headers)
                print(f"     💾 Disimpan: {file_path.name} (Tab 'Listing' & 'Listing 2')")
                success_count += 1

                # Upload ke Google Drive jika diaktifkan
                if upload:
                    ok, _ = upload_file_to_drive(str(file_path), owner_str, APP_SCRIPT_URL)
                    if ok:
                        upload_count += 1

            print("\n" + "=" * 65)
            print("🎉 PROSES SELESAI")
            print(f"   • Total file dibuat lokal: {success_count} file di output_owners/")
            if upload:
                print(f"   • Total berhasil diunggah ke Google Drive: {upload_count}/{success_count}")
            else:
                print("   • Mode Lokal: File belum diunggah ke Google Drive (gunakan flag --upload untuk mengunggah).")
            print("=" * 65)
            return not upload or upload_count == success_count
    except Timeout:
        info = get_lock_info()
        print(f"\n⚠️ Tidak dapat menjalankan proses: Pipeline sedang dikunci oleh proses lain.")
        print(f"   Sedang memproses: {info.get('owner', 'Tidak diketahui')} (oleh {info.get('user', 'Unknown')}, PID: {info.get('pid', '-')})")
        print("   Harap tunggu hingga proses tersebut selesai.")
        return False


def list_owners():
    """Tampilkan daftar owner dan jumlah outlet per platform."""
    df_go = load_gofood_data()
    df_grab = load_grab_data()
    df_shopee = load_shopee_data()
    dfs = [df for df in [df_go, df_grab, df_shopee] if not df.empty]
    if not dfs:
        print("Data kosong.")
        return
    combined = pd.concat(dfs, ignore_index=True)
    if "Nama Pemilik" not in combined.columns:
        print("Kolom Nama Pemilik tidak ditemukan.")
        return

    owners = combined[combined["Nama Pemilik"].notna()]["Nama Pemilik"].unique()
    owners = sorted(owners, key=lambda x: str(x).lower())

    print(f"\n{'No.':<4} {'Nama Pemilik':<30} {'GoFood':<8} {'Grab':<8} {'Shopee':<8} {'Total':<8}")
    print("─" * 68)
    for i, o in enumerate(owners, 1):
        o_df = combined[combined["Nama Pemilik"] == o]
        go_n = len(o_df[o_df["Aplikator"] == "GoFood"])
        gr_n = len(o_df[o_df["Aplikator"] == "GrabFood"])
        sh_n = len(o_df[o_df["Aplikator"] == "ShopeeFood"])
        print(f"{i:<4} {str(o):<30} {go_n:<8} {gr_n:<8} {sh_n:<8} {len(o_df):<8}")
    print("─" * 68)
    print(f"Total: {len(owners)} owner terdaftar.\n")


def main():
    parser = argparse.ArgumentParser(
        description="Gabungkan data GoFood & Grab per-owner ke format YYYY-MM-DD HH_MM [Nama Pemilik].xlsx dengan 2 tab identik dan upload ke Google Drive."
    )
    parser.add_argument("--owner", type=str, default=None, help="Filter nama owner tertentu")
    parser.add_argument("--upload", action="store_true", help="Unggah file yang dibuat ke Google Drive")
    parser.add_argument("--local-only", action="store_true", help="Hanya buat file Excel lokal di output_owners/ tanpa upload")
    parser.add_argument("--list-owners", action="store_true", help="Tampilkan daftar semua owner dan jumlah outlet")
    parser.add_argument("--lock-timeout", type=int, default=0, help="Timeout menunggu JobLock (detik, default: 0 = fail-fast)")
    parser.add_argument("--weekly-dbr", action="store_true", help="Jalankan pipeline weekly live dengan snapshot DBR privat dari environment")

    args = parser.parse_args()

    if args.list_owners:
        list_owners()
        return

    do_upload = args.upload and not args.local_only
    if args.weekly_dbr:
        snapshot_path = os.getenv("WEEKLY_DBR_SNAPSHOT", "").strip()
        result_path = os.getenv("WEEKLY_OWNER_RESULT_FILE", "").strip()
        if not args.owner or not snapshot_path or not result_path or not Path(snapshot_path).is_file():
            print("❌ Mode weekly membutuhkan --owner serta snapshot dan lokasi hasil dari environment.")
            return 1
        dbr_df = pd.DataFrame()
        try:
            dbr_df = pd.read_csv(snapshot_path)
            result = generate_for_owner_pipeline(
                args.owner,
                upload=do_upload,
                source="dbr",
                live_scrape=True,
                requested_by="Weekly SOT",
                dbr_df=dbr_df,
            )
            safe_result = {
                "owner": args.owner,
                "success": bool(result.get("success")),
                "is_partial": bool(result.get("is_partial")),
                "error": _redact_weekly_log(result.get("error", ""), _weekly_redactions(dbr_df, args.owner)),
                "portal_results": result.get("portal_results", []),
                "sot_rows": result.get("sot_rows", []),
            }
            result_file = Path(result_path)
            result_file.parent.mkdir(parents=True, exist_ok=True)
            result_file.write_text(json.dumps(safe_result, ensure_ascii=False), encoding="utf-8")
            return 0 if result.get("success") else 1
        except Exception as exc:
            print(f"❌ Weekly DBR pipeline gagal ({type(exc).__name__}): {_redact_weekly_log(exc, _weekly_redactions(dbr_df, args.owner))}")
            try:
                Path(result_path).write_text(json.dumps({"owner": args.owner, "success": False, "error": type(exc).__name__, "portal_results": [], "sot_rows": []}), encoding="utf-8")
            except Exception:
                pass
            return 1

    ok = process_and_combine(owner_filter=args.owner, upload=do_upload, lock_timeout=args.lock_timeout)
    return 0 if ok is not False else 1


if __name__ == "__main__":
    sys.exit(main())

