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
import tempfile
import uuid
import hashlib
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
    portal_failures = summary_data.get("portal_failures", [])
    sheet_res = summary_data.get("sheet_sync", {})
    sheet_failed = sheet_res.get("status") not in (None, "success", "skipped")
    duration_str = summary_data.get("duration", "0m")

    has_failure = failed_count > 0 or bool(portal_failures) or sheet_failed
    has_live_output = success_count > 0 or summary_data.get("live_outlet_count", 0) > 0
    status_color = 0x10b981 if not has_failure else (0xf59e0b if has_live_output else 0xef4444)
    if not has_failure and sheet_res.get("status") == "skipped":
        status_title = "✅ [CRON] PENARIKAN SELESAI — SOT FULL REFRESH DILEWATI"
    else:
        status_title = "🎉 [CRON] SINKRONISASI MINGGUAN SELESAI" if not has_failure else "⚠️ [CRON] SINKRONISASI MINGGUAN SELESAI DENGAN CATATAN"

    failed_desc = ""
    if failed_count > 0:
        failed_desc = "\n\n**⚠️ Owner dengan Kendala Scraping:**\n"
        for item in summary_data.get("failed_owners", [])[:6]:
            reason = str(item.get("reason", "Gagal"))[:140]
            failed_desc += f"• **{item.get('owner')}**: {reason}\n"
        if failed_count > 6:
            failed_desc += f"*...dan {failed_count - 6} owner lainnya.*\n"
    if portal_failures:
        failed_desc += "\n**⚠️ Portal gagal/parsial:**\n"
        for item in portal_failures[:8]:
            message = str(item.get("message", "Gagal"))[:140]
            failed_desc += (
                f"• **{item.get('owner', '-')}** / {item.get('aplikator', '-')} / "
                f"{item.get('portal', '-')}: {item.get('stage', 'scraping')} — {message}\n"
            )
        if len(portal_failures) > 8:
            failed_desc += f"*...dan {len(portal_failures) - 8} portal lainnya.*\n"

    if sheet_res.get("status") == "success":
        sheet_status_str = "Sukses"
    elif sheet_res.get("status") == "skipped":
        sheet_status_str = f"Dilewati aman ({sheet_res.get('message', '-')})"
    else:
        sheet_status_str = f"Gagal ({sheet_res.get('error', '-')})"
    sheet_details = f"Ditimpa: {sheet_res.get('total_updated', sheet_res.get('updated', 0))} | Baru: {sheet_res.get('total_inserted', sheet_res.get('inserted', 0))} | Total Baris: {sheet_res.get('totalRows', '-')}"

    embed = {
        "title": status_title,
        "description": f"Penarikan data outlet otomatis mingguan untuk seluruh owner telah selesai diproses.{failed_desc}",
        "color": status_color,
        "fields": [
            {"name": "📅 Tanggal Eksekusi", "value": date_str, "inline": True},
            {"name": "⏱️ Durasi Total", "value": duration_str, "inline": True},
            {"name": "👥 Status Owner", "value": f"✅ **{success_count}** Selesai / ❌ **{failed_count}** Perlu diulang (Total: {total_owners})\nPortal gagal: **{len(portal_failures)}**", "inline": False},
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


def send_discord_start(total_owners: int, date_str: str):
    """Mengirimkan pemberitahuan ke Discord bahwa batch sync mingguan telah dimulai."""
    embed = {
        "title": "🚀 [CRON] SINKRONISASI MINGGUAN DIMULAI",
        "description": "Proses penarikan data mingguan untuk seluruh owner sedang berjalan otomatis di server.",
        "color": 0x3b82f6,
        "fields": [
            {"name": "📅 Tanggal Eksekusi", "value": date_str, "inline": True},
            {"name": "👥 Total Target", "value": f"**{total_owners}** Owner", "inline": True},
            {"name": "⚙️ Status", "value": "Sedang memproses penarikan listing, kompilasi Excel, upload Drive & sinkronisasi SOT...", "inline": False}
        ],
        "footer": {
            "text": "Outlet Info Automation • Weekly Cron Job System"
        },
        "timestamp": datetime.datetime.now(datetime.timezone.utc).isoformat()
    }

    if DISCORD_WEBHOOK_URL:
        try:
            resp = requests.post(DISCORD_WEBHOOK_URL, json={"embeds": [embed]}, timeout=15)
            if resp.status_code in (200, 204):
                print("  ✓ Notifikasi mulai terkirim ke Discord via Webhook.")
                return True
        except Exception as e:
            print(f"  [!] Gagal mengirim notifikasi mulai ke Webhook: {e}")

    if DISCORD_BOT_TOKEN and DISCORD_CHANNEL_ID:
        try:
            url = f"https://discord.com/api/v10/channels/{DISCORD_CHANNEL_ID}/messages"
            headers = {
                "Authorization": f"Bot {DISCORD_BOT_TOKEN}",
                "Content-Type": "application/json"
            }
            resp = requests.post(url, headers=headers, json={"embeds": [embed]}, timeout=15)
            if resp.status_code in (200, 201):
                print(f"  ✓ Notifikasi mulai terkirim ke Discord Channel {DISCORD_CHANNEL_ID} via Bot API.")
                return True
        except Exception as e:
            print(f"  [!] Exception saat mengirim notifikasi mulai ke Discord: {e}")

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


def _owner_dbr_signature(dbr_df: pd.DataFrame, owner: str) -> str:
    """Fingerprint only non-secret DBR target fields for safe resume invalidation."""
    owner_rows = dbr_df[
        dbr_df["Nama Pemilik"].fillna("").astype(str).str.strip().str.casefold() == owner.casefold()
    ]
    columns = [column for column in ("Aplikator", "Nama Portal", "Nama Brand", "Nama Outlet", "Group ID") if column in owner_rows]
    values = []
    for _, row in owner_rows[columns].iterrows():
        values.append(tuple("" if pd.isna(row[column]) else str(row[column]).strip() for column in columns))
    values.sort()
    serialized = json.dumps(values, ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(serialized.encode("utf-8")).hexdigest()


def run_weekly_sync(
    dry_run: bool = False,
    resume: bool = True,
    owner_limit: int = None,
    specific_owner: str = None,
    no_sheet: bool = False,
    no_upload: bool = False
):
    """Scrape all DBR targets once per owner; use those fresh rows for Listing and SOT."""
    today_str = datetime.date.today().strftime("%Y%m%d")
    today_display = datetime.date.today().strftime("%Y-%m-%d")
    start_ts = time.time()
    workflow_version = "weekly-dbr-live-v1"

    print("\n" + "=" * 70)
    print("  🚀 WEEKLY LIVE PORTAL SCRAPE → LISTING + SOT")
    print(f"  Tanggal: {today_display} | Mode Dry-Run: {dry_run} | Resume: {resume}")
    print("=" * 70)

    print("[*] Memuat satu snapshot live Tab 'DBR' untuk seluruh proses weekly...")
    dbr_df = load_dbr_master_data(force_live=True)
    if dbr_df.empty or "Nama Pemilik" not in dbr_df.columns:
        cp = {
            "date": today_display,
            "workflow": workflow_version,
            "completed_owners": [],
            "failed_owners": [{"owner": "DBR", "reason": "Snapshot live DBR gagal dimuat; cache lama ditolak.", "code": 1}],
            "portal_failures": [],
            "total_owners": 0,
            "sheet_sync": {"status": "skipped", "message": "DBR live tidak tersedia; SOT tidak diubah."},
            "duration": "0m 0s",
        }
        report_file = REPORTS_DIR / f"weekly_report_{today_str}.json"
        report_file.write_text(json.dumps(cp, indent=2, ensure_ascii=False), encoding="utf-8")
        print("❌ DBR live tidak tersedia; scraping dan full refresh SOT dibatalkan.")
        send_discord_summary(cp)
        return False

    all_dbr_owners = sorted(dbr_df["Nama Pemilik"].dropna().astype(str).str.strip().unique())
    all_dbr_owners = [owner for owner in all_dbr_owners if owner and owner.casefold() not in ("nan", "none", "-")]
    if not all_dbr_owners:
        cp = {
            "date": today_display,
            "workflow": workflow_version,
            "completed_owners": [],
            "failed_owners": [{"owner": "DBR", "reason": "Snapshot DBR live tidak memiliki owner valid.", "code": 1}],
            "portal_failures": [],
            "total_owners": 0,
            "sheet_sync": {"status": "skipped", "message": "Tidak ada owner valid; SOT tidak diubah."},
            "duration": "0m 0s",
        }
        report_file = REPORTS_DIR / f"weekly_report_{today_str}.json"
        report_file.write_text(json.dumps(cp, indent=2, ensure_ascii=False), encoding="utf-8")
        print("❌ Snapshot DBR tidak memiliki owner valid; SOT tetap dipertahankan.")
        send_discord_summary(cp)
        return False

    all_owners = list(all_dbr_owners)
    if specific_owner:
        all_owners = [owner for owner in all_owners if specific_owner.casefold() in owner.casefold()]
        print(f"[*] Filter owner spesifik dari DBR: {all_owners}")
    if owner_limit and owner_limit > 0:
        all_owners = all_owners[:owner_limit]
        print(f"[*] Membatasi proses ke {owner_limit} owner pertama.")

    total_target = len(all_owners)
    print(f"[*] Snapshot DBR memuat {len(all_dbr_owners)} owner; target run ini {total_target} owner.")
    if not all_owners and not dry_run:
        cp = {
            "date": today_display,
            "workflow": workflow_version,
            "completed_owners": [],
            "failed_owners": [{"owner": specific_owner or "Target", "reason": "Tidak ada owner DBR yang cocok dengan filter run.", "code": 1}],
            "portal_failures": [],
            "total_owners": 0,
            "sheet_sync": {"status": "skipped", "message": "Target kosong; SOT tidak diubah."},
            "duration": "0m 0s",
        }
        report_file = REPORTS_DIR / f"weekly_report_{today_str}.json"
        report_file.write_text(json.dumps(cp, indent=2, ensure_ascii=False), encoding="utf-8")
        print("❌ Tidak ada owner yang cocok; SOT tidak diubah.")
        send_discord_summary(cp)
        return False
    if dry_run:
        print("\n🔍 [DRY-RUN] Portal target dari snapshot DBR:")
        for idx, owner in enumerate(all_owners, 1):
            owner_df = dbr_df[dbr_df["Nama Pemilik"].astype(str).str.strip().str.casefold() == owner.casefold()]
            app_counts = owner_df.get("Aplikator", pd.Series(dtype=str)).fillna("").astype(str).value_counts().to_dict()
            print(f"  {idx:>2}. {owner:<30} | {app_counts}")
        print("✓ Dry run selesai. Tidak ada login portal, Listing, atau SOT yang dijalankan.")
        return True

    is_full_target = all_owners == all_dbr_owners and not specific_owner and not (owner_limit and owner_limit > 0)
    cp = load_checkpoint(today_str) if resume else {}
    if not resume or cp.get("workflow") != workflow_version or cp.get("target_owners") != all_owners:
        cp = {
            "date": today_display,
            "workflow": workflow_version,
            "run_id": uuid.uuid4().hex[:12],
            "target_owners": all_owners,
            "completed_owners": [],
            "failed_owners": [],
            "portal_failures": [],
            "start_time": datetime.datetime.now().isoformat(),
            "end_time": None,
            "sheet_sync": {},
        }
    cp.setdefault("run_id", uuid.uuid4().hex[:12])

    results_name = cp.get("results_file") or f"weekly_live_results_{today_str}_{cp['run_id']}.json"
    results_path = CACHE_DIR / Path(results_name).name
    owner_results = {}
    if resume and results_path.exists():
        try:
            stored = json.loads(results_path.read_text(encoding="utf-8"))
            if stored.get("run_id") == cp.get("run_id"):
                owner_results = stored.get("owners", {})
                if not isinstance(owner_results, dict):
                    owner_results = {}
        except Exception as exc:
            print(f"⚠️ Gagal membaca hasil resume ({type(exc).__name__}); owner akan diproses ulang bila hasilnya tidak tersedia.")
    cp["results_file"] = results_path.name
    current_signatures = {owner: _owner_dbr_signature(dbr_df, owner) for owner in all_owners}
    previous_signatures = cp.get("owner_signatures", {})
    if not isinstance(previous_signatures, dict):
        previous_signatures = {}
    if not isinstance(cp.get("completed_owners"), list):
        cp["completed_owners"] = []
    changed_owners = [
        owner for owner in cp.get("completed_owners", [])
        if previous_signatures.get(owner) != current_signatures.get(owner)
    ]
    for owner in changed_owners:
        owner_results.pop(owner, None)
        print(f"[*] DBR target owner '{owner}' berubah; mengulang scraping owner ini saat resume.")
    cp["owner_signatures"] = current_signatures
    cp["completed_owners"] = [
        owner for owner in cp.get("completed_owners", [])
        if owner in owner_results and owner in all_owners
    ]
    completed_set = set(cp["completed_owners"])
    if not isinstance(cp.get("failed_owners"), list):
        cp["failed_owners"] = []
    if not isinstance(cp.get("portal_failures"), list):
        cp["portal_failures"] = []
    cp["target_owners"] = all_owners
    cp["total_owners"] = total_target

    if not completed_set:
        send_discord_start(total_target, today_display)

    python_bin = sys.executable
    # Snapshot credential file is private, short-lived, and never placed in argv/logs.
    with tempfile.TemporaryDirectory(prefix="weekly-dbr-", dir=CACHE_DIR) as private_dir:
        private_path = Path(private_dir)
        os.chmod(private_path, 0o700)
        dbr_snapshot = private_path / "dbr_snapshot.csv"
        dbr_df.to_csv(dbr_snapshot, index=False)
        os.chmod(dbr_snapshot, 0o600)

        for idx, owner in enumerate(all_owners, 1):
            if owner in completed_set:
                print(f"\n[{idx}/{total_target}] ⏩ [SKIP] Owner '{owner}' sudah selesai dengan hasil live tersimpan.")
                continue

            print(f"\n[{idx}/{total_target}] 👤 Memproses Owner: '{owner}'...")
            owner_result_path = private_path / f"owner_{idx}.json"
            env = {
                **os.environ,
                "WEEKLY_DBR_SNAPSHOT": str(dbr_snapshot),
                "WEEKLY_OWNER_RESULT_FILE": str(owner_result_path),
            }
            cmd = [
                python_bin, "-u", str(BASE_DIR / "export_and_upload_drive.py"),
                "--owner", owner, "--weekly-dbr",
            ]
            if not no_upload:
                cmd.append("--upload")

            sub_start = time.time()
            proc = subprocess.Popen(
                cmd,
                cwd=BASE_DIR,
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
            )
            out_lines = []
            redactions = exporter._weekly_redactions(dbr_df, owner)
            for line in proc.stdout:
                clean_line = exporter._redact_weekly_log(line.rstrip(), redactions)
                if not clean_line:
                    continue
                out_lines.append(clean_line)
                if len(out_lines) > 80:
                    out_lines.pop(0)
                print(clean_line[:500], flush=True)
            returncode = proc.wait()
            sub_dur = int(time.time() - sub_start)

            owner_result = {}
            if owner_result_path.exists():
                try:
                    owner_result = json.loads(owner_result_path.read_text(encoding="utf-8"))
                except Exception as exc:
                    owner_result = {"success": False, "error": f"Hasil owner tidak terbaca ({type(exc).__name__})"}
            owner_rows = owner_result.get("sot_rows", []) if isinstance(owner_result.get("sot_rows"), list) else []
            owner_results[owner] = owner_rows
            portal_outcomes = owner_result.get("portal_results", []) if isinstance(owner_result.get("portal_results"), list) else []
            portal_failed = [item for item in portal_outcomes if item.get("status") != "success"]
            cp["portal_failures"] = [item for item in cp.get("portal_failures", []) if item.get("owner") != owner]
            cp["portal_failures"].extend(portal_failed)

            is_owner_complete = bool(owner_result.get("success")) and returncode == 0 and not portal_failed
            cp["failed_owners"] = [item for item in cp.get("failed_owners", []) if item.get("owner") != owner]
            if is_owner_complete:
                print(f"  ✅ [SUKSES] Owner '{owner}' seluruh portal selesai; {len(owner_rows)} baris live disimpan ({sub_dur}s).")
                if owner not in cp["completed_owners"]:
                    cp["completed_owners"].append(owner)
                completed_set.add(owner)
            else:
                reason = owner_result.get("error") or (
                    f"{len(portal_failed)} portal gagal/parsial" if portal_failed else
                    next((line for line in reversed(out_lines) if any(k in line.casefold() for k in ("gagal", "error", "exception", "timeout"))), f"Exit code {returncode}")
                )
                reason = exporter._redact_weekly_log(reason, redactions)[:500]
                print(f"  ⚠️ [PARSIAL/GAGAL] Owner '{owner}' | Live rows={len(owner_rows)} | {reason}")
                cp["completed_owners"] = [name for name in cp.get("completed_owners", []) if name != owner]
                completed_set.discard(owner)
                cp["failed_owners"].append({"owner": owner, "reason": reason, "code": returncode})

            results_path.write_text(json.dumps({"run_id": cp["run_id"], "owners": owner_results}, ensure_ascii=False), encoding="utf-8")
            save_checkpoint(today_str, cp)
            cleanup_zombie_processes()
            if idx < total_target:
                time.sleep(random.randint(10, 15))

    # A filtered owner/limit run must never replace the complete SOT snapshot.
    if no_sheet:
        cp["sheet_sync"] = {"status": "skipped", "message": "Sinkronisasi SOT dinonaktifkan oleh opsi run."}
    elif not is_full_target:
        cp["sheet_sync"] = {"status": "skipped", "message": "Full refresh SOT dilewati karena target hanya sebagian owner."}
        print("⚠️ Full refresh SOT dilewati: run ini tidak mencakup seluruh owner DBR.")
    else:
        rows = [row for owner in all_dbr_owners for row in owner_results.get(owner, []) if isinstance(row, dict)]
        if not rows:
            cp["sheet_sync"] = {"status": "skipped", "message": "Tidak ada hasil live sukses; SOT lama dipertahankan."}
            print("⚠️ Tidak ada hasil live sukses; full refresh SOT dibatalkan agar sheet tidak kosong.")
        else:
            print("\n" + "─" * 70)
            print(f"[*] Full refresh SOT dari {len(rows)} baris hasil live sukses (snapshot DBR tunggal)...")
            try:
                ok_sheet, res_sheet = export_dataframe_to_sot(pd.DataFrame(rows))
                cp["sheet_sync"] = res_sheet if ok_sheet else {"status": "failed", "error": res_sheet.get("error", "Sinkronisasi gagal")}
            except Exception as exc:
                cp["sheet_sync"] = {"status": "error", "error": f"{type(exc).__name__}: {exc}"}
                print(f"[!] Exception saat sinkronisasi Google Sheet: {type(exc).__name__}")

    end_ts = time.time()
    dur_str = f"{int((end_ts - start_ts) // 60)}m {int((end_ts - start_ts) % 60)}s"
    cp["date"] = today_display
    cp["end_time"] = datetime.datetime.now().isoformat()
    cp["duration"] = dur_str
    cp["total_owners"] = total_target
    cp["live_outlet_count"] = sum(len(rows) for rows in owner_results.values())
    save_checkpoint(today_str, cp)

    report_file = REPORTS_DIR / f"weekly_report_{today_str}.json"
    report_file.write_text(json.dumps(cp, indent=2, ensure_ascii=False), encoding="utf-8")
    print("\n" + "=" * 70)
    print(f"  🏁 BATCH SELESAI ({dur_str})")
    print(f"  • Owner tuntas: {len(cp.get('completed_owners', []))}/{total_target}")
    print(f"  • Owner perlu ulang: {len(cp.get('failed_owners', []))}/{total_target}")
    print(f"  • Portal gagal: {len(cp.get('portal_failures', []))}")
    print(f"  • Outlet live: {sum(len(rows) for rows in owner_results.values())}")
    print(f"  • Sheet: {cp.get('sheet_sync', {}).get('status', 'N/A')}")
    print(f"  • Laporan: {report_file.name}")
    print("=" * 70)

    send_discord_summary(cp)
    sheet_status = cp.get("sheet_sync", {}).get("status")
    sheet_ok = no_sheet or sheet_status in ("success", "skipped")
    return not cp.get("failed_owners") and not cp.get("portal_failures") and sheet_ok


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
