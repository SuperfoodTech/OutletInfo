#!/usr/bin/env python3
"""
export_and_upload_drive.py
==========================
Menggabungkan data outlet per-owner dari multi-aplikator (GoFood dan GrabFood,
serta siap untuk ShopeeFood) ke dalam satu file per-owner:
    YYYY-MM-DD HH_MM [Nama Pemilik].xlsx

Spesifikasi File:
- Menggunakan template standar 37 kolom (sheet 'Listing').
- Memiliki 2 tab (sheet) yang 100% identik:
    Tab 1: 'Listing'
    Tab 2: 'Listing 2'
- Disimpan lokal di folder `output_owners/`.
- Diunggah ke Google Drive ke dalam folder per-owner (membuat folder otomatis jika belum ada)
  melalui Google Apps Script Web App (APP_SCRIPT_URL).
"""

import os
import sys
import glob
import json
import base64
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
GOOGLE_SHEET_VERCEL_URL = os.getenv(
    "GOOGLE_SHEET_VERCEL_URL",
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vTprbPPf_J5gAVL3PYeHbbdl5ZXQvb17HY2lJGPI2xg13Ly3AGT8eYHLYmU_m1NdtkBVg-qUGv1BoEE/pub?output=csv"
)
APP_SCRIPT_URL = os.getenv("APP_SCRIPT_URL", "")


def get_template_headers():
    """Mengambil 37 kolom header resmi dari file template."""
    if TEMPLATE_PATH.exists():
        try:
            wb = openpyxl.load_workbook(str(TEMPLATE_PATH), read_only=True)
            sheet = wb["Listing"] if "Listing" in wb.sheetnames else wb.active
            headers = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
            wb.close()
            return headers
        except Exception as e:
            print(f"⚠️ Gagal membaca template ({e}), menggunakan default 37 header.")
            
    # Fallback default 37 kolom jika template tidak terbaca
    return [
        'Nama Pemilik', 'Nama Brand', 'Model', 'Tipe', 'Outlet', 'Nomor HP',
        'Aplikator', 'Nama Portal', 'Group ID', 'Nama Listing', 'Link', 'Store ID',
        'Status Listing', 'Alamat', 'Nama Bank', 'Nama Pemilik Rekening', 'Nomor Rekening',
        'Nama Akses', 'Email FoodMaster1', 'Email FoodMaster2', 'Nama Pengguna',
        'Kata Sandi', 'Nama Portal.1', 'S Nomor HP Akses Pemilik', 'S Username Akses Pemilik',
        'S Kata Sandi Akses Pemilik', 'S Allvbadmin Username Akses Staff',
        'S Allvbadmin Kata Sandi Akses Staff', 'S Bot Username Akses Staff',
        'S Bot Kata Sandi Akses Staff', 'S BD Username Akses Staff',
        'S BD Kata Sandi Akses Staff', 'BD', 'Status Internal', 'Tanggal Live',
        'Tanggal Churn', 'Tarif'
    ]


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
            combined = combined.drop_duplicates(subset=["Store ID"], keep="last")
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


def load_vercel_data(force_live=False):
    """
    Memuat data outlet dari Google Sheet Vercel (Live CSV).
    Menyimpan cache lokal di cache/vercel_sheet_cache.csv untuk fallback offline.
    """
    cache_file = CACHE_DIR / "vercel_sheet_cache.csv"
    CACHE_DIR.mkdir(parents=True, exist_ok=True)

    csv_text = ""
    # 1. Coba fetch live dari Google Sheet jika memungkinkan
    try:
        url = GOOGLE_SHEET_VERCEL_URL
        sep = "&" if "?" in url else "?"
        # Cache busting timestamp untuk melewati caching edge/CDN Google Spreadsheets
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
            with open(cache_file, "w", encoding="utf-8") as f:
                f.write(csv_text)
            print(f"   ✓ [VERCEL] Berhasil memuat data live dari Google Sheet ({len(csv_text)} bytes).")
    except Exception as e:
        print(f"⚠️ Fetch live Vercel Sheet melewati batas waktu / offline ({e}). Menggunakan cache lokal.")

    # 2. Fallback ke file cache lokal jika live gagal
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

    if "Nama Akses" in df.columns:
        df["Nama Portal"] = df["Nama Akses"].astype(str).str.strip()

    # Untuk Shopee, fallback Nama Portal ke Merchant Name
    if "Merchant Name" in df.columns:
        df["Nama Portal"] = df.apply(
            lambda r: r["Merchant Name"] if (pd.isna(r.get("Nama Portal")) or not str(r.get("Nama Portal")).strip()) and pd.notna(r.get("Merchant Name")) else r.get("Nama Portal", ""),
            axis=1
        )

    # Status Listing default LIVE untuk data aktif
    df["Status Listing"] = "LIVE"

    return df


def write_data_to_sheet(ws, df, headers):
    """Menulis data ke sheet yang telah memiliki header, styling, dan lebar kolom dari template."""
    body_font = Font(name='Arial', size=10)
    alignment = Alignment(horizontal='left', vertical='center')

    # Alias mapping jika nama kolom di df sedikit berbeda dengan nama header template
    col_aliases = {
        'Nama Pemilik': ['Owner', 'Nama Pemilik', 'owner', 'pemilik'],
        'Nama Brand': ['Nama Outlet', 'Nama Brand', 'Brand', 'brand'],
        'Aplikator': ['Aplikasi', 'Aplikator', 'app', 'Platform'],
        'Nama Portal': ['Nama Akses', 'Nama Portal', 'Merchant Name', 'portal'],
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

            cell = ws.cell(row=idx, column=c_idx, value=val_str)
            cell.font = body_font
            cell.alignment = alignment
            if h in text_cols or c_idx in (6, 9, 12, 17, 24):
                cell.number_format = '@'


def save_owner_workbook(owner_df, file_path, headers):
    """
    Menyimpan data per-owner ke file Excel dengan 2 tab 100% identik:
    - Tab 1: 'Listing'
    - Tab 2: 'Listing 2'
    Dibuat langsung dengan mengkloning template 'YYYY-MM-DD HH_MM Nama Pemilik.xlsx'
    sehingga pewarnaan kolom header (Merah, Pink, Hijau, Oranye), font Arial,
    dan lebar kolom asli terjaga 100% identik.
    Menggunakan penulisan atomic via folder .tmp agar terhindar dari file korup jika terjadi crash.
    """
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
            'master' menggunakan master lokal.
    """
    if source == "vercel":
        combined = load_vercel_data(force_live=force_live)
    else:
        df_go = load_gofood_data()
        df_grab = load_grab_data()
        df_shopee = load_shopee_data()
        dfs = [df for df in [df_go, df_grab, df_shopee] if not df.empty]
        combined = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
    
    if combined.empty or "Nama Pemilik" not in combined.columns:
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


def run_subprocess_stream(cmd, cwd, keywords, on_log, timeout_sec=150):
    """Menjalankan subprocess dengan streaming output real-time dan batas waktu yang aman agar tidak stuck."""
    import subprocess
    import time
    env = {**os.environ, "PYTHONUNBUFFERED": "1"}
    try:
        p = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            cwd=str(cwd),
            env=env,
            bufsize=1
        )
        start_time = time.time()
        while True:
            # Cegah proses menggantung melebihi timeout
            if time.time() - start_time > timeout_sec:
                try:
                    p.kill()
                except Exception:
                    pass
                on_log(f"⏱️ Melewati batas waktu ({timeout_sec}s). Melanjutkan ke tahap berikutnya...")
                break
            
            line = p.stdout.readline()
            if not line:
                if p.poll() is not None:
                    break
                time.sleep(0.05)
                continue
                
            line_s = line.strip()
            if not line_s:
                continue
                
            # Abaikan garis pembatas dekoratif
            if line_s.startswith("===") or line_s.startswith("---") or line_s.startswith("───"):
                continue

            # Tampilkan pesan jika ada indikator status atau cocok kata kunci
            is_status = any(line_s.startswith(p) for p in ("[*]", "[✓]", "🚀", "🌐", "➡️", "📧", "⏳", "✅", "⚠️", "❌", "🎉", "⚡", "🏢", "📍"))
            has_kw = any(k.lower() in line_s.lower() for k in keywords)
            if is_status or has_kw:
                on_log(line_s[:85])
                
        p.poll()
        return p.returncode or 0
    except Exception as e:
        on_log(f"⚠️ Error proses: {e}")
        return -1


def run_live_scraping_for_owner(owner_name, aplikator="all", progress_cb=None, retry_targets=None):
    """
    Menjalankan live scraping untuk GoFood, GrabFood, dan ShopeeFood secara berurutan.
    progress_cb: function(percent: int, log_line: str)
    retry_targets: list of dicts (jika hanya menargetkan outlet/aplikator tertentu yang gagal)
    Mengembalikan dict status hasil scraping per-aplikator.
    """
    clean_owner = owner_name.strip()
    python_bin = sys.executable
    scrape_status = {
        "gofood": None,
        "grab": None,
        "shopee": None,
        "shopee_merchants": {}
    }

    def send_log(pct, msg):
        if progress_cb:
            progress_cb(pct, msg)
        print(f"[{pct:>3}%] {msg}")

    # Cek apakah aplikator masuk dalam target eksekusi
    def should_run_app(app_key):
        if retry_targets is None:
            return aplikator in ("all", app_key)
        target_apps = [str(t.get("aplikator", "")).lower() for t in retry_targets]
        if app_key == "gofood" and any("go" in a for a in target_apps):
            return True
        if app_key == "grab" and any("grab" in a for a in target_apps):
            return True
        if app_key == "shopee" and any("shopee" in a for a in target_apps):
            return True
        return False

    # 1. LIVE SCRAPING GOFOOD
    if should_run_app("gofood"):
        send_log(15, f"🚀 [GoFood] Memulai penarikan live untuk '{clean_owner}'...")
        headless_go = os.getenv("HEADLESS_GOFOOD", os.getenv("HEADLESS", "true")).strip().lower() in ("true", "1", "yes", "y")
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
            keywords=("Store ID", "Berhasil", "Portal", "Login", "Owner", "Restricted", "Memproses", "OTP", "Filter", "Gagal"),
            on_log=lambda m: send_log(25, f"[GoFood] {m}"),
            timeout_sec=160
        )
        scrape_status["gofood"] = rc
        if rc == 0:
            send_log(35, f"✅ [GoFood] Selesai memproses '{clean_owner}'.")
        else:
            send_log(35, f"⚠️ [GoFood] Selesai dengan kode status {rc}.")

    # 2. LIVE SCRAPING GRABFOOD
    if should_run_app("grab"):
        send_log(40, f"🚀 [GrabFood] Memulai penarikan live untuk '{clean_owner}'...")
        headless_grab = os.getenv("HEADLESS_GRAB", os.getenv("HEADLESS", "true")).strip().lower() in ("true", "1", "yes", "y")
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
            keywords=("Group ID", "Berhasil", "Target", "Portal", "Login", "Store", "Bank", "Owner", "Filter", "Gagal", "Password", "Salah", "Master@123", "fallback", "Sukses"),
            on_log=lambda m: send_log(55, f"[Grab] {m}"),
            timeout_sec=180
        )
        scrape_status["grab"] = rc
        if rc == 0:
            send_log(65, f"✅ [GrabFood] Selesai memproses '{clean_owner}'.")
        else:
            send_log(65, f"⚠️ [GrabFood] Selesai dengan kode status {rc}.")

    # 3. LIVE SCRAPING SHOPEEFOOD
    if should_run_app("shopee"):
        send_log(68, f"🚀 [ShopeeFood] Memeriksa akun Shopee untuk '{clean_owner}'...")
        try:
            v_df = load_vercel_data()
            o_shopee = v_df[(v_df["Nama Pemilik"].astype(str).str.strip().str.lower() == clean_owner.lower()) & (v_df["Aplikator"] == "ShopeeFood")]
            merchant_names = []
            if not o_shopee.empty:
                for _, r in o_shopee.iterrows():
                    for col in ["Merchant Name", "Nama Portal", "Nama Brand", "Nama Akses"]:
                        val = str(r.get(col) or "").strip()
                        if val and val.lower() not in ("nan", "none", "-") and val not in merchant_names:
                            # Filter spesifik jika retry_targets ditentukan
                            if retry_targets is not None:
                                shopee_targets = [str(t.get("portal") or t.get("outlet") or "").strip().lower() for t in retry_targets if "shopee" in str(t.get("aplikator", "")).lower()]
                                if not any(st in val.lower() or val.lower() in st for st in shopee_targets if st):
                                    continue
                            merchant_names.append(val)
                            break
            
            if merchant_names:
                headless_shopee = os.getenv("HEADLESS_SHOPEE", os.getenv("HEADLESS", "true")).strip().lower() in ("true", "1", "yes", "y")
                for idx, merchant_name in enumerate(merchant_names, start=1):
                    send_log(70, f"[ShopeeFood] Menarik data toko [{idx}/{len(merchant_names)}] '{merchant_name}'...")
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
                        keywords=("Store", "Berhasil", "Merchant", "Data", "Sukses", "Total", "Selesai"),
                        on_log=lambda m: send_log(74, f"[Shopee] {m}"),
                        timeout_sec=180
                    )
                    scrape_status["shopee_merchants"][merchant_name] = rc_sh
                scrape_status["shopee"] = 0 if all(c == 0 for c in scrape_status["shopee_merchants"].values()) else 1
                send_log(76, f"✅ [ShopeeFood] Selesai memproses {len(merchant_names)} merchant.")
            else:
                send_log(76, f"ℹ️ [ShopeeFood] Tidak ditemukan nama merchant Shopee untuk '{clean_owner}'.")
        except Exception as e:
            scrape_status["shopee"] = 1
            send_log(76, f"⚠️ [ShopeeFood] Exception scraper: {e}")

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

            for _, act_row in matched_rows.iterrows():
                store_id = str(act_row.get("Store ID") or "").strip()
                has_store_id = bool(store_id and store_id.lower() not in ("nan", "none", "-", ""))

                act_portal = norm(act_row.get("Nama Portal") or "")
                act_outlet = norm(act_row.get("Nama Outlet") or act_row.get("Nama Brand") or act_row.get("Nama Listing") or "")
                act_user = norm(act_row.get("Nama Pengguna") or "")

                match_identity = False
                if norm_exp_portal and norm_exp_portal in act_portal:
                    match_identity = True
                elif norm_exp_outlet and norm_exp_outlet in act_outlet:
                    match_identity = True
                elif norm_exp_user and norm_exp_user == act_user:
                    match_identity = True
                elif len(matched_rows) == 1 and len(expected_df[expected_df["Aplikator"].astype(str).str.lower() == app.lower()]) == 1:
                    match_identity = True

                if match_identity and has_store_id:
                    found_valid = True
                    break

        if not found_valid:
            reason = "Store ID kosong / data toko tidak lengkap"
            app_lower = app.lower()
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


def generate_for_owner_pipeline(owner_name, aplikator="all", upload=True, source="vercel", live_scrape=True, progress_callback=None, lock_timeout=0, requested_by="System", retry_targets=None):
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
    """
    def log(pct, msg):
        if progress_callback:
            progress_callback(pct, msg)
        print(f"[{pct:>3}%] {msg}")

    try:
        with JobLock(timeout=lock_timeout, task_info={"owner": owner_name, "user": requested_by, "aplikator": aplikator, "source": source}):
            log(10, "Inisialisasi direktori dan template 37 kolom...")
            headers = get_template_headers()
            OUTPUT_OWNERS_DIR.mkdir(parents=True, exist_ok=True)

            # 1. LIVE SCRAPING (Jika diaktifkan dan bukan mode __ALL__)
            scrape_status = {}
            if live_scrape and owner_name != "__ALL__":
                log(12, f"Menjalankan penarikan live data toko untuk '{owner_name}'...")
                scrape_status = run_live_scraping_for_owner(owner_name, aplikator=aplikator, progress_cb=progress_callback, retry_targets=retry_targets)

            log(78, f"Membaca data sumber ({source.upper()}) untuk aplikator: {aplikator.upper()}...")
            if source == "vercel":
                src_df = load_vercel_data()
            else:
                dfs = []
                if aplikator in ("all", "gofood"):
                    df_go = load_gofood_data()
                    if not df_go.empty:
                        dfs.append(df_go)
                if aplikator in ("all", "grab"):
                    df_grab = load_grab_data()
                    if not df_grab.empty:
                        dfs.append(df_grab)
                if aplikator in ("all", "shopee"):
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
            if aplikator == "gofood":
                owner_df = owner_df[owner_df["Aplikator"] == "GoFood"]
            elif aplikator == "grab":
                owner_df = owner_df[owner_df["Aplikator"] == "GrabFood"]
            elif aplikator == "shopee":
                owner_df = owner_df[owner_df["Aplikator"] == "ShopeeFood"]

            if owner_df.empty:
                log(100, f"⚠️ Tidak ada outlet {aplikator.upper()} untuk owner '{owner_name}'.")
                return {"success": False, "error": f"Tidak ada outlet {aplikator} untuk owner {owner_name}"}

            expected_df = owner_df.copy()

            # Enrich data dari hasil scraping jika tersedia (Bank, Rekening, Store ID, Alamat)
            log(82, "Menggabungkan hasil penarikan data toko (Store ID, Alamat, Bank)...")
            try:
                df_go = load_gofood_data()
                df_gr = load_grab_data()
                df_sh = load_shopee_data()
                scraped_list = [d for d in [df_go, df_gr, df_sh] if not d.empty]
                scraped_all = pd.concat(scraped_list, ignore_index=True) if scraped_list else pd.DataFrame()

                scraped_owner = pd.DataFrame()
                if not scraped_all.empty and "Nama Pemilik" in scraped_all.columns:
                    scraped_owner = scraped_all[scraped_all["Nama Pemilik"].astype(str).str.strip().str.lower() == owner_name.strip().lower()]

                final_rows = []
                for app in owner_df["Aplikator"].dropna().unique():
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
                            # Sesuai instruksi user: untuk kolom brand ambil dari kolom Nama Outlet saja
                            if nama_outlet_val:
                                row_dict["Nama Outlet"] = nama_outlet_val
                                row_dict["Nama Brand"] = nama_outlet_val
                            for col in ["Nama Akses", "Email FoodMaster1", "Email FoodMaster2", "Nama Pengguna", "Kata Sandi",
                                        "Nama Portal.1", "S Nomor HP Akses Pemilik", "S Username Akses Pemilik", "S Kata Sandi Akses Pemilik",
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
                    if "Nama Outlet" in owner_df.columns:
                        owner_df["Nama Brand"] = owner_df["Nama Outlet"].astype(str).str.strip()
            except Exception as e:
                print(f"⚠️ Info enrich scraped: {e}")

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

            # Verifikasi kelengkapan ekstraksi terhadap Vercel Sheet
            missing_items = verify_extraction_completeness(expected_df, owner_df, scrape_status)
            is_partial = len(missing_items) > 0
            completed_count = max(0, len(expected_df) - len(missing_items))

            if is_partial:
                log(86, f"⚠️ Selesai dengan catatan: {len(missing_items)} outlet belum lengkap / gagal ditarik live.")
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
                        "gofood": go_n,
                        "grab": gr_n,
                        "shopee": sh_n,
                        "error": res.get("error")
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
                "gofood": go_n,
                "grab": gr_n,
                "shopee": sh_n,
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
            if not dfs:
                print("\n❌ Tidak ada data outlet yang ditemukan.")
                return

            # Gabungkan data seluruh platform
            combined_df = pd.concat(dfs, ignore_index=True)
            
            # Filter baris yang memiliki Nama Pemilik
            if "Nama Pemilik" not in combined_df.columns:
                print("❌ Kolom 'Nama Pemilik' tidak ditemukan dalam data.")
                return

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
                    return
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
    except Timeout:
        info = get_lock_info()
        print(f"\n⚠️ Tidak dapat menjalankan proses: Pipeline sedang dikunci oleh proses lain.")
        print(f"   Sedang memproses: {info.get('owner', 'Tidak diketahui')} (oleh {info.get('user', 'Unknown')}, PID: {info.get('pid', '-')})")
        print("   Harap tunggu hingga proses tersebut selesai.")
        return


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

    args = parser.parse_args()

    if args.list_owners:
        list_owners()
        return

    do_upload = args.upload and not args.local_only
    process_and_combine(owner_filter=args.owner, upload=do_upload, lock_timeout=args.lock_timeout)


if __name__ == "__main__":
    main()

