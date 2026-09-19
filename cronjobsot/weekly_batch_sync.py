#!/usr/bin/env python3
"""
Weekly Batch Listing Sync (Cron Job)
====================================
Menarik seluruh listing outlet dari semua owner di Google Sheet Tab 'DBR' secara otomatis,
mengompilasi workbook Excel per-owner, mengunggah ke Google Drive,
meng-upsert data ke Google Spreadsheet Tab 'SOT' (dengan kolom Terakhir Diperbaharui),
dan mengirimkan laporan ringkasan eksekutif ke Discord.

Jadwal Eksekusi: Setiap hari Minggu pukul 21:00 WIB (via systemd timer)
Direktori: /mnt/DATA/Proyek/Outlet Info/cronjobsot/
"""

import os
import sys
import gc
import json
import time
import random
import datetime
import argparse
import subprocess
from pathlib import Path

# Setup Path Proyek
CRON_DIR = Path(__file__).resolve().parent
BASE_DIR = CRON_DIR.parent
CACHE_DIR = CRON_DIR / "cache"
REPORTS_DIR = CRON_DIR / "reports"
CACHE_DIR.mkdir(parents=True, exist_ok=True)
REPORTS_DIR.mkdir(parents=True, exist_ok=True)

if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))
if str(CRON_DIR) not in sys.path:
    sys.path.insert(0, str(CRON_DIR))

from dotenv import load_dotenv
load_dotenv(BASE_DIR / ".env")

import pandas as pd
import requests

import export_and_upload_drive as exporter
from export_to_sot import (
    load_dbr_master_data,
    export_dataframe_to_sot,
    load_all_local_masters,
    TARGET_SPREADSHEET_ID,
    TARGET_SHEET_NAME,
    TARGET_SHEET_GID
)

DISCORD_BOT_TOKEN = os.getenv("DISCORD_BOT_TOKEN", "")
DISCORD_WEBHOOK_URL = os.getenv("DISCORD_WEBHOOK_URL", "")
DISCORD_CHANNEL_ID = os.getenv("DISCORD_CHANNEL_ID", os.getenv("OFD_CHANNEL_ID", ""))


def cleanup_zombie_processes():
    """Membersihkan proses headless Chrome atau browser zombie untuk membebaskan RAM."""
    gc.collect()
    try:
        subprocess.run(["pkill", "-f", "chrome.*--headless"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass


def send_discord_summary(summary_data: dict):
    """Mengirimkan embed laporan ringkasan mingguan ke Discord (Webhook atau Bot REST API)."""
    date_str = summary_data.get("date", datetime.date.today().isoformat())
    total_owners = summary_data.get("total_owners", 0)
    success_count = len(summary_data.get("completed_owners", []))
    failed_count = len(summary_data.get("failed_owners", []))
    sheet_res = summary_data.get("sheet_sync", {})
    duration_str = summary_data.get("duration", "0m")

    status_color = 0x10b981 if failed_count == 0 else (0xf59e0b if success_count > 0 else 0xef4444)
    status_title = "🎉 [CRON] SINKRONISASI MINGGUAN SELESAI" if failed_count == 0 else "⚠️ [CRON] SINKRONISASI MINGGUAN SELESAI DENGAN CATATAN"

    failed_desc = ""
    if failed_count > 0:
        failed_desc = "\n\n**⚠️ Owner dengan Kendala Scraping:**\n"
        for item in summary_data.get("failed_owners", [])[:10]:
            failed_desc += f"• **{item.get('owner')}**: {item.get('reason', 'Gagal')}\n"
        if failed_count > 10:
            failed_desc += f"*...dan {failed_count - 10} owner lainnya.*\n"

    sheet_status_str = "Sukses" if sheet_res.get("status") == "success" else f"Gagal ({sheet_res.get('error', '-')})"
    sheet_details = f"Ditimpa: {sheet_res.get('total_updated', sheet_res.get('updated', 0))} | Baru: {sheet_res.get('total_inserted', sheet_res.get('inserted', 0))} | Total Baris: {sheet_res.get('totalRows', '-')}"

    embed = {
        "title": status_title,
        "description": f"Penarikan data outlet otomatis mingguan untuk seluruh owner telah selesai diproses.{failed_desc}",
        "color": status_color,
        "fields": [
            {"name": "📅 Tanggal Eksekusi", "value": date_str, "inline": True},
            {"name": "⏱️ Durasi Total", "value": duration_str, "inline": True},
            {"name": "👥 Status Owner", "value": f"✅ **{success_count}** Sukses / ❌ **{failed_count}** Gagal (Total: {total_owners})", "inline": False},
            {"name": "📊 Google Sheet 'SOT'", "value": f"Status: **{sheet_status_str}**\n{sheet_details}", "inline": False},
            {"name": "🔗 Tautan Spreadsheet", "value": f"[Buka Google Spreadsheet Tab SOT](https://docs.google.com/spreadsheets/d/{TARGET_SPREADSHEET_ID}/edit#gid={TARGET_SHEET_GID})", "inline": False}
        ],
        "footer": {
            "text": "Outlet Info Automation • Weekly Cron Job System"
        },
        "timestamp": datetime.datetime.now(datetime.timezone.utc).isoformat()
    }

    # 1. Coba kirim via Webhook jika ada
    if DISCORD_WEBHOOK_URL:
        try:
            resp = requests.post(DISCORD_WEBHOOK_URL, json={"embeds": [embed]}, timeout=15)
            if resp.status_code in (200, 204):
                print("  ✓ Laporan terkirim ke Discord via Webhook.")
                return True
        except Exception as e:
            print(f"  [!] Gagal mengirim ke Discord Webhook: {e}")

    # 2. Coba kirim via Bot REST API jika channel ID dan token tersedia
    if DISCORD_BOT_TOKEN and DISCORD_CHANNEL_ID:
        try:
            url = f"https://discord.com/api/v10/channels/{DISCORD_CHANNEL_ID}/messages"
            headers = {
                "Authorization": f"Bot {DISCORD_BOT_TOKEN}",
                "Content-Type": "application/json"
            }
            resp = requests.post(url, headers=headers, json={"embeds": [embed]}, timeout=15)
            if resp.status_code in (200, 201):
                print(f"  ✓ Laporan terkirim ke Discord Channel ID {DISCORD_CHANNEL_ID} via Bot API.")
                return True
            else:
                print(f"  [!] Gagal kirim via Bot API (Status {resp.status_code}): {resp.text}")
        except Exception as e:
            print(f"  [!] Exception saat mengirim ke Discord Bot API: {e}")

    print("  ℹ️ Discord Webhook/Bot API tidak terkonfigurasi atau gagal dikirim.")
    return False


def load_checkpoint(date_str: str) -> dict:
    """Memuat data checkpoint harian jika ada."""
    cp_file = CACHE_DIR / f"weekly_checkpoint_{date_str}.json"
    if cp_file.exists():
        try:
            return json.loads(cp_file.read_text(encoding="utf-8"))
        except Exception as e:
            print(f"  [!] Gagal membaca checkpoint {cp_file.name}: {e}")
    return {
        "date": date_str,
        "completed_owners": [],
        "failed_owners": [],
        "start_time": datetime.datetime.now().isoformat(),
        "end_time": None,
        "sheet_sync": {}
    }


def save_checkpoint(date_str: str, cp_data: dict):
    """Menyimpan data checkpoint ke disk."""
    cp_file = CACHE_DIR / f"weekly_checkpoint_{date_str}.json"
    try:
        cp_file.write_text(json.dumps(cp_data, indent=2, ensure_ascii=False), encoding="utf-8")
    except Exception as e:
        print(f"  [!] Gagal menyimpan checkpoint {cp_file.name}: {e}")


def run_weekly_sync(
    dry_run: bool = False,
    resume: bool = True,
    owner_limit: int = None,
    specific_owner: str = None,
    no_sheet: bool = False,
    no_upload: bool = False
):
    """Fungsi utama pengatur eksekusi batch mingguan."""
    today_str = datetime.date.today().strftime("%Y%m%d")
    today_display = datetime.date.today().strftime("%Y-%m-%d")
    start_ts = time.time()

    print("\n" + "=" * 70)
    print(f"  🚀 WEEKLY BATCH OUTLET LISTING SYNC")
    print(f"  Tanggal: {today_display} | Mode Dry-Run: {dry_run} | Resume: {resume}")
    print("=" * 70)

    # 1. Dapatkan daftar seluruh owner dari Google Sheet Tab 'DBR'
    print("[*] Memuat daftar owner dari Tab 'DBR' Google Sheet...")
    v_df = load_dbr_master_data()
    if v_df.empty:
        print("❌ Gagal memuat data Tab DBR / fallback Sheet (koneksi offline dan cache kosong).")
        return False

    all_owners = sorted(v_df["Nama Pemilik"].dropna().astype(str).str.strip().unique())
    # Filter owner valid (bukan strip atau nan)
    all_owners = [o for o in all_owners if o and o.lower() not in ("nan", "none", "-")]

    if specific_owner:
        all_owners = [o for o in all_owners if specific_owner.lower() in o.lower()]
        print(f"[*] Filter spesifik owner: {all_owners}")

    if owner_limit and owner_limit > 0:
        all_owners = all_owners[:owner_limit]
        print(f"[*] Membatasi proses ke {owner_limit} owner pertama.")

    total_target = len(all_owners)
    print(f"[*] Terdeteksi total {total_target} owner yang akan diproses.")

    if dry_run:
        print("\n🔍 [DRY-RUN] Daftar owner yang akan diproses:")
        for idx, o in enumerate(all_owners, 1):
            o_df = v_df[v_df["Nama Pemilik"].astype(str).str.strip().str.lower() == o.lower()]
            go_c = len(o_df[o_df["Aplikator"] == "GoFood"])
            gr_c = len(o_df[o_df["Aplikator"] == "GrabFood"])
            sh_c = len(o_df[o_df["Aplikator"] == "ShopeeFood"])
            print(f"  {idx:>2}. {o:<30} (GoFood: {go_c}, Grab: {gr_c}, Shopee: {sh_c})")
        print("\n✓ Dry run selesai. Tidak ada scraping yang dijalankan.")
        return True

    # 2. Inisialisasi Checkpoint
    cp = load_checkpoint(today_str) if resume else {
        "date": today_display,
        "completed_owners": [],
        "failed_owners": [],
        "start_time": datetime.datetime.now().isoformat(),
        "end_time": None,
        "sheet_sync": {}
    }
    completed_set = set(cp.get("completed_owners", []))

    python_bin = sys.executable

    # 3. Iterasi Per-Owner secara Terisolasi
    for idx, owner in enumerate(all_owners, 1):
        if owner in completed_set:
            print(f"\n[{idx}/{total_target}] ⏩ [SKIP] Owner '{owner}' sudah selesai di checkpoint.")
            continue

        print(f"\n[{idx}/{total_target}] 👤 Memproses Owner: '{owner}'...")
        cmd = [
            python_bin,
            "-u",
            str(BASE_DIR / "export_and_upload_drive.py"),
            "--owner", owner,
        ]
        if not no_upload:
            cmd.append("--upload")

        sub_start = time.time()
        proc = subprocess.run(
            cmd,
            cwd=BASE_DIR,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace"
        )
        sub_dur = int(time.time() - sub_start)

        if proc.returncode == 0:
            print(f"  ✅ [SUKSES] Owner '{owner}' selesai diproses ({sub_dur}s).")
            cp["completed_owners"].append(owner)
            completed_set.add(owner)
        else:
            # Ambil potongan error dari output
            out_lines = proc.stdout.splitlines()
            err_snippet = out_lines[-1] if out_lines else f"Exit code {proc.returncode}"
            for line in reversed(out_lines[-10:]):
                if any(k in line.lower() for k in ("error", "exception", "gagal", "timeout")):
                    err_snippet = line.strip()
                    break
            print(f"  ❌ [GAGAL] Owner '{owner}' keluar dengan kode {proc.returncode}: {err_snippet}")
            cp["failed_owners"].append({"owner": owner, "reason": err_snippet, "code": proc.returncode})

        # Simpan progres checkpoint setelah setiap owner
        save_checkpoint(today_str, cp)

        # Pembersihan Memori & Zombie Browser
        cleanup_zombie_processes()

        # Jeda Jitter Antar-Owner (10-15 detik) untuk meredam deteksi bot & mendinginkan CPU
        if idx < total_target:
            jitter = random.randint(10, 15)
            time.sleep(jitter)

    # 4. Penggabungan Data & Sinkronisasi ke Google Spreadsheet 'SOT'
    if not no_sheet:
        print("\n" + "─" * 70)
        print(f"[*] Memulai sinkronisasi seluruh data ke Google Spreadsheet 'SOT'...")
        try:
            combined_all = load_all_local_masters()
            if not combined_all.empty:
                print(f"  ✓ Terkumpul {len(combined_all)} total baris outlet gabungan.")
                ok_sheet, res_sheet = export_dataframe_to_sot(combined_all)
                cp["sheet_sync"] = res_sheet if ok_sheet else {"status": "failed", "error": res_sheet.get("error")}
            else:
                print("  ⚠️ Tidak ada data master lokal, mencoba memuat dari Tab 'DBR'...")
                dbr_df = load_dbr_master_data()
                if not dbr_df.empty:
                    ok_sheet, res_sheet = export_dataframe_to_sot(dbr_df)
                    cp["sheet_sync"] = res_sheet if ok_sheet else {"status": "failed", "error": res_sheet.get("error")}
                else:
                    print("  ⚠️ Tidak ada data outlet yang berhasil dimuat untuk Google Sheet.")
                    cp["sheet_sync"] = {"status": "skipped", "message": "Data kosong"}
        except Exception as e:
            print(f"  [!] Exception saat sinkronisasi Google Sheet: {e}")
            cp["sheet_sync"] = {"status": "error", "error": str(e)}

    # 5. Laporan Selesai & Notifikasi Discord
    end_ts = time.time()
    total_minutes = int((end_ts - start_ts) // 60)
    total_seconds = int((end_ts - start_ts) % 60)
    dur_str = f"{total_minutes}m {total_seconds}s"

    cp["end_time"] = datetime.datetime.now().isoformat()
    cp["duration"] = dur_str
    cp["total_owners"] = total_target
    save_checkpoint(today_str, cp)

    # Simpan laporan ringkasan permanen
    report_file = REPORTS_DIR / f"weekly_report_{today_str}.json"
    report_file.write_text(json.dumps(cp, indent=2, ensure_ascii=False), encoding="utf-8")

    print("\n" + "=" * 70)
    print(f"  🏁 BATCH SELESAI ({dur_str})")
    print(f"  • Sukses: {len(cp.get('completed_owners', []))}/{total_target}")
    print(f"  • Gagal : {len(cp.get('failed_owners', []))}/{total_target}")
    print(f"  • Sheet : {cp.get('sheet_sync', {}).get('status', 'N/A')}")
    print(f"  • Laporan: {report_file.name}")
    print("=" * 70)

    # Kirim rekap ke Discord
    send_discord_summary(cp)
    return len(cp.get("failed_owners", [])) == 0


def main():
    parser = argparse.ArgumentParser(description="Weekly Scheduled Batch Sync for All Owners to Excel, Drive & Google Sheets")
    parser.add_argument("--dry-run", action="store_true", help="Tampilkan daftar owner tanpa mengeksekusi penarikan")
    parser.add_argument("--no-resume", action="store_true", help="Ulangi dari awal tanpa menggunakan checkpoint")
    parser.add_argument("--owner-limit", type=int, default=None, help="Batas jumlah owner yang diproses (untuk pengujian)")
    parser.add_argument("--owner", type=str, default=None, help="Proses hanya owner tertentu")
    parser.add_argument("--no-sheet", action="store_true", help="Lewati sinkronisasi ke Google Spreadsheet 'SOT'")
    parser.add_argument("--no-upload", action="store_true", help="Lewati upload ke Google Drive per-owner")

    args = parser.parse_args()
    success = run_weekly_sync(
        dry_run=args.dry_run,
        resume=not args.no_resume,
        owner_limit=args.owner_limit,
        specific_owner=args.owner,
        no_sheet=args.no_sheet,
        no_upload=args.no_upload
    )
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
