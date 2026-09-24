"""
Sheet Syncer Module for Google Spreadsheet 'SOT'
================================================
Mengirimkan dan melakukan upsert data outlet ke Google Spreadsheet tab 'SOT'
(Spreadsheet ID: 15_Xx5ixOcxy0L90U_BrFHfVnAgjIZpiiGBEK4Ce4SyI, GID: 2135653103)
melalui Google Apps Script Web App (APP_SCRIPT_URL).
"""

import os
import json
import requests
import pandas as pd
from pathlib import Path
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(BASE_DIR / ".env")

SOT_APP_SCRIPT_URL = os.getenv("SOT_APP_SCRIPT_URL", os.getenv("APP_SCRIPT_URL", ""))
APP_SCRIPT_URL = SOT_APP_SCRIPT_URL
TARGET_SPREADSHEET_ID = "15_Xx5ixOcxy0L90U_BrFHfVnAgjIZpiiGBEK4Ce4SyI"
TARGET_SHEET_NAME = "SOT"
TARGET_SHEET_GID = "2135653103"

# 43 Kolom Template Standar DBR (Kolom ke-44 'Terakhir Diperbaharui' ditambahkan oleh Apps Script di tab SOT)
STANDARD_HEADERS = [
    "Nama Pemilik", "Nama Brand", "Model", "Tipe", "Outlet", "Nomor HP",
    "Aplikator", "Group ID", "Nama Listing", "Link", "Store ID",
    "Status Listing", "Alamat", "Nama Bank", "Nama Pemilik Rekening", "Nomor Rekening",
    "Nama Akses", "Email FoodMaster1", "Email FoodMaster2", "Nama Pengguna", "Kata Sandi",
    "Nama Portal", "S Nomor HP Akses Pemilik", "S Username Akses Pemilik", "S Kata Sandi Akses Pemilik",
    "S Allvbadmin Username Akses Staff", "S Allvbadmin Kata Sandi Akses Staff",
    "S Bot Username Akses Staff", "S Bot Kata Sandi Akses Staff",
    "S BD Username Akses Staff", "S BD Kata Sandi Akses Staff",
    "BD", "Status Internal", "Tanggal Live", "Tanggal Churn", "Tarif",
    "Status Bot", "Vercel Kata Sandi", "Paket", "Tanggal Mulai Layanan",
    "Tanggal Berakhir Layanan", "Akses Username", "Akses Kata Sandi"
]

COL_ALIASES = {
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


def format_grab_food_link(store_id_or_link: str) -> str:
    """
    Format ID Toko atau Link GrabFood ke URL Konsumen GrabFood:
    https://food.grab.com/id/id/restaurant/x/{clean_id}/
    di mana {clean_id} adalah ID tanpa tanda strip '-'.
    """
    if not store_id_or_link:
        return ""
    val = str(store_id_or_link).strip()
    if not val or val.lower() in ('nan', 'none'):
        return ""
    if "grab.com" in val:
        val = val.split("?")[0].split("#")[0].rstrip("/").split("/")[-1]
    clean_id = val.replace("-", "").strip()
    if not clean_id:
        return ""
    return f"https://food.grab.com/id/id/restaurant/x/{clean_id}/"


def format_dataframe_to_rows(df: pd.DataFrame, headers: list[str]) -> list[list]:
    """Mengubah DataFrame outlet ke bentuk list of lists sesuai urutan header template."""
    rows = []
    if df is None or df.empty:
        return rows

    for _, r in df.iterrows():
        row_vals = []
        for h in headers:
            if not h:
                row_vals.append("")
                continue

            if h in ('Nama Brand', 'Brand'):
                val = r.get('Nama Outlet')
                if val is None or pd.isna(val) or str(val).strip() == '':
                    val = r.get('Nama Brand') or r.get('Brand')
            else:
                val = r.get(h)
                if val is None or pd.isna(val) or str(val).strip() == '':
                    aliases = COL_ALIASES.get(h, [])
                    for alias in aliases:
                        if alias in r and pd.notna(r.get(alias)) and str(r.get(alias)).strip() != '':
                            val = r.get(alias)
                            break

            val_str = str(val).strip() if pd.notna(val) and val is not None else ''
            if val_str.lower() in ('nan', 'none'):
                val_str = ''

            # Bersihkan suffix .0 pada angka/ID panjang
            if val_str.endswith(".0") and val_str[:-2].isdigit():
                val_str = val_str[:-2]

            # Normalisasi kolom Link untuk GrabFood ke format URL konsumen
            if h == 'Link':
                app_val = str(r.get('Aplikator') or r.get('Aplikasi') or '').strip().lower()
                if 'grab' in app_val:
                    target_ref = val_str or str(r.get('Store ID') or '').strip()
                    val_str = format_grab_food_link(target_ref)

            row_vals.append(val_str)
        rows.append(row_vals)

    return rows


def sync_outlets_to_google_sheet(
    df_or_rows,
    headers: list[str] = None,
    app_script_url: str = None,
    chunk_size: int = 25
) -> tuple[bool, dict]:
    """
    Mengirimkan baris outlet ke Apps Script untuk di-upsert ke Google Spreadsheet tab 'SOT'.
    Data yang sama akan ditimpa dan kolom 'Terakhir Diperbaharui' akan diperbarui.
    """
    url = app_script_url or APP_SCRIPT_URL
    if not url:
        return False, {"error": "APP_SCRIPT_URL belum disetel di .env"}

    active_headers = headers
    if not active_headers:
        if isinstance(df_or_rows, pd.DataFrame) and len(df_or_rows.columns) >= 30:
            active_headers = [c for c in df_or_rows.columns if c != "Terakhir Diperbaharui"]
        else:
            active_headers = STANDARD_HEADERS

    if isinstance(df_or_rows, pd.DataFrame):
        rows = format_dataframe_to_rows(df_or_rows, active_headers)
    elif isinstance(df_or_rows, list):
        if df_or_rows and isinstance(df_or_rows[0], dict):
            df_temp = pd.DataFrame(df_or_rows)
            rows = format_dataframe_to_rows(df_temp, active_headers)
        else:
            rows = df_or_rows
    else:
        return False, {"error": "Format input data tidak valid"}

    if not rows:
        return True, {"status": "skipped", "message": "Tidak ada baris data untuk dikirim"}

    print(f"[*] Menyiapkan sinkronisasi {len(rows)} baris outlet ke Google Sheet '{TARGET_SHEET_NAME}'...")

    total_updated = 0
    total_inserted = 0
    final_res = {}

    # Kirim per-chunk untuk menghindari batas payload / timeout Apps Script
    for i in range(0, len(rows), chunk_size):
        chunk = rows[i : i + chunk_size]
        payload = {
            "action": "sync_outlets",
            "spreadsheetId": TARGET_SPREADSHEET_ID,
            "sheetName": TARGET_SHEET_NAME,
            "gid": TARGET_SHEET_GID,
            "headers": active_headers,
            "rows": chunk
        }

        try:
            resp = requests.post(url, json=payload, timeout=90)
            resp.raise_for_status()
            res_json = resp.json()

            if res_json.get("status") == "success":
                up = res_json.get("updated", 0)
                ins = res_json.get("inserted", 0)
                total_updated += up
                total_inserted += ins
                final_res = res_json
                print(f"  ✓ Chunk {i // chunk_size + 1}: {up} ditimpa, {ins} baru.")
            else:
                err_msg = res_json.get("message", "Unknown error from Apps Script")
                print(f"  [!] Gagal sinkronisasi chunk {i // chunk_size + 1}: {err_msg}")
                return False, {"error": err_msg, "chunk": i}
        except Exception as e:
            print(f"  [!] Exception saat mengirim ke Apps Script: {e}")
            return False, {"error": str(e), "chunk": i}

    final_res["total_updated"] = total_updated
    final_res["total_inserted"] = total_inserted
    print(f"✅ [GOOGLE SHEET] Berhasil sinkronisasi tab '{TARGET_SHEET_NAME}': {total_updated} ditimpa, {total_inserted} baru.")
    return True, final_res


if __name__ == "__main__":
    print("Testing Sheet Syncer...")
    sample_data = pd.DataFrame([{
        "Nama Pemilik": "Testing Owner",
        "Nama Outlet": "Warung Test",
        "Aplikator": "GoFood",
        "Store ID": "999888777",
        "Status Listing": "LIVE",
        "Alamat": "Jl. Testing No. 1"
    }])
    ok, res = sync_outlets_to_google_sheet(sample_data)
    print("Result:", ok, res)

