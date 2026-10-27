#!/usr/bin/env python3
"""
bot_sot.py
==========
Discord Bot khusus untuk /sot (Sinkronisasi Data Listing ke Google Spreadsheet Tab 'SOT').
Menjalankan layanan bot mandiri untuk channel #sot.

Fitur:
- Slash Command `/sot`:
  * Mode: Full Cronjob (Scrape + Drive + SOT Sync), Sync SOT Saja, atau Dry Run
  * Filter target owner dengan Autocomplete interaktif
  * Dukungan resume checkpoint
- Slash Command `/status_sot`:
  * Cek laporan mingguan terakhir, checkpoint aktif hari ini, dan status tab SOT
- Control Panel Interaktif:
  * Tombol `🚀 Mulai Tarik Cronjob` (dengan streaming live ANSI console log & tombol pembatalan)
  * Tombol `⚡ Sync SOT Saja` (upsert cepat dari master lokal tanpa scraping ulang)
  * Tombol `🔍 Dry Run Target` (simulasi antrean target owner)
  * Tombol Link `📑 Buka Tab SOT`
- Proteksi pipeline multi-proses (JobLock & Running Status) agar tidak bentrok
"""

import os
import sys
import re
import json
import time
import glob
import asyncio
import datetime
import threading
import subprocess
from pathlib import Path
from dotenv import load_dotenv

import discord
from discord import app_commands
from discord.ext import commands

# ─── Path & Environment Setup ─────────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent
CRON_DIR = BASE_DIR / "cronjobsot"
REPORTS_DIR = CRON_DIR / "reports"
CACHE_DIR = CRON_DIR / "cache"
REPORTS_DIR.mkdir(parents=True, exist_ok=True)
CACHE_DIR.mkdir(parents=True, exist_ok=True)

if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))
if str(CRON_DIR) not in sys.path:
    sys.path.insert(0, str(CRON_DIR))

load_dotenv(BASE_DIR / ".env")

BOT_TOKEN = os.getenv("DISCORD_BOT_TOKEN_SOT", "").strip()
CHANNEL_ID = os.getenv("DISCORD_CHANNEL_ID_SOT", "").strip()
ROOT_DRIVE_URL = "https://drive.google.com/drive/u/0/folders/19VIrypPcBmNNbjDLGS7kxp_yIdBBwXjB"

# Import helper dari export_to_sot & pipeline_lock
try:
    from export_to_sot import (
        export_dataframe_to_sot,
        load_all_local_masters,
        load_dbr_master_data,
        TARGET_SPREADSHEET_ID,
        TARGET_SHEET_NAME,
        TARGET_SHEET_GID
    )
    from export_and_upload_drive import get_owners_with_metadata
    from pipeline_lock import JobLock, is_pipeline_locked, get_lock_info, Timeout
except ImportError as e:
    print(f"⚠️ Gagal mengimpor modul SOT: {e}")
    TARGET_SPREADSHEET_ID = "15_Xx5ixOcxy0L90U_BrFHfVnAgjIZpiiGBEK4Ce4SyI"
    TARGET_SHEET_NAME = "SOT"
    TARGET_SHEET_GID = "2135653103"

SOT_SPREADSHEET_URL = f"https://docs.google.com/spreadsheets/d/{TARGET_SPREADSHEET_ID}/edit#gid={TARGET_SHEET_GID}"

# ─── Visual Theme & Helpers ───────────────────────────────────────────────────
THEME_PRIMARY  = 0x5865F2   # Blurple
THEME_SUCCESS  = 0x10B981   # Emerald Green
THEME_PROGRESS = 0x3B82F6   # Blue
THEME_WARNING  = 0xF59E0B   # Amber Yellow
THEME_ERROR    = 0xEF4444   # Rose Red
THEME_ACCENT   = 0x8B5CF6   # Purple

DRIVE_ICON_URL   = "https://cdn-icons-png.flaticon.com/512/281/281761.png"
STORE_THUMBNAIL  = "https://cdn-icons-png.flaticon.com/512/869/869636.png"
SUCCESS_THUMBNAIL= "https://cdn-icons-png.flaticon.com/512/190/190411.png"

# Cache metadata owner untuk autocomplete
_owner_cache = {"data": [], "ts": 0.0}
_owner_lock = asyncio.Lock()

_sot_running_info = {
    "is_running": False,
    "user": None,
    "started_at": None,
    "process": None,
    "cancel_event": None
}


def create_progress_bar(percent: int, length: int = 14) -> str:
    """Membuat visual progress bar bergaya modern."""
    filled = int(length * percent / 100)
    empty = length - filled
    return f"{'▰' * filled}{'▱' * empty} **{percent}%**"


async def get_cached_owners(max_age: float = 180.0):
    """Mengambil daftar owner dengan caching in-memory."""
    now = time.time()
    if _owner_cache["data"] and (now - _owner_cache["ts"] < max_age):
        return _owner_cache["data"]
    async with _owner_lock:
        if _owner_cache["data"] and (now - _owner_cache["ts"] < max_age):
            return _owner_cache["data"]
        try:
            owners = await asyncio.to_thread(get_owners_with_metadata, source="vercel", force_live=False)
            if owners:
                _owner_cache["data"] = owners
                _owner_cache["ts"] = now
        except Exception:
            pass
        return _owner_cache["data"]


def get_latest_sot_report():
    """Mengambil ringkasan laporan mingguan SOT terakhir dari direktori cronjobsot/reports."""
    if not REPORTS_DIR.exists():
        return None
    report_files = sorted(REPORTS_DIR.glob("weekly_report_*.json"), reverse=True)
    if not report_files:
        return None
    try:
        data = json.loads(report_files[0].read_text(encoding="utf-8"))
        return data, report_files[0].name
    except Exception:
        return None


def get_today_sot_checkpoint():
    """Mengambil status checkpoint hari ini jika tersedia."""
    today_str = datetime.date.today().strftime("%Y%m%d")
    cp_file = CACHE_DIR / f"weekly_checkpoint_{today_str}.json"
    if cp_file.exists():
        try:
            return json.loads(cp_file.read_text(encoding="utf-8"))
        except Exception:
            pass
    return None


# ─── Interactive Views ────────────────────────────────────────────────────────

class SotRunningView(discord.ui.View):
    """View interaktif yang tampil saat cronjob SOT sedang berjalan, menyediakan tombol Batal."""

    def __init__(self, user: discord.User, cancel_event: threading.Event, proc_holder: dict = None):
        super().__init__(timeout=3600)
        self.user = user
        self.cancel_event = cancel_event
        self.proc_holder = proc_holder or {}

    @discord.ui.button(label="Batalkan Cronjob", style=discord.ButtonStyle.danger, emoji="🛑", custom_id="btn_cancel_sot_running")
    async def cancel_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        if interaction.user.id != self.user.id:
            try:
                await interaction.response.send_message("❌ Hanya pengguna yang memulai proses yang dapat membatalkannya.", ephemeral=True)
            except Exception:
                pass
            return

        button.disabled = True
        button.label = "Sedang Membatalkan..."
        self.cancel_event.set()
        proc = self.proc_holder.get("proc")
        if proc:
            try:
                proc.terminate()
            except Exception:
                pass
        try:
            await interaction.response.edit_message(view=self)
        except (discord.NotFound, discord.HTTPException):
            pass


class SotControlView(discord.ui.View):
    """Panel kontrol interaktif untuk perintah /sot."""

    def __init__(self, user: discord.User, target_owner: str = None, mode: str = "full", resume: bool = True, owners_meta: list = None):
        super().__init__(timeout=600)
        self.user = user
        self.target_owner = target_owner
        self.mode = mode
        self.resume = resume
        self.owners_meta = owners_meta or []
        self.rebuild_components()

    def rebuild_components(self):
        self.clear_items()
        self.add_item(SotStartButton(self))
        self.add_item(SotSyncOnlyButton(self))
        self.add_item(SotDryRunButton(self))
        self.add_item(discord.ui.Button(label="Buka Tab SOT", url=SOT_SPREADSHEET_URL, style=discord.ButtonStyle.link, emoji="📑"))
        self.add_item(SotCancelButton(self))

    def build_embed(self) -> discord.Embed:
        embed = discord.Embed(
            title="🔄 CRONJOB SINKRONISASI TAB SOT",
            description=(
                "Panel kontrol otomasi **Cronjob Sinkronisasi Mingguan** ke Google Spreadsheet Tab `SOT`.\n"
                "Proses ini akan menarik listing seluruh owner, menyusun workbook Excel, mengunggah ke Google Drive, "
                "dan meng-upsert data ke tab SOT."
            ),
            color=THEME_PRIMARY,
            timestamp=datetime.datetime.now()
        )
        embed.set_thumbnail(url=STORE_THUMBNAIL)

        if self.target_owner and self.target_owner != "__ALL__":
            target_str = f"👤 **{self.target_owner}** (Khusus 1 Owner)"
        else:
            total_cnt = len(self.owners_meta) if self.owners_meta else 69
            target_str = f"📦 **Semua Owner Terdaftar** (~{total_cnt} Owner)"
        embed.add_field(name="🎯 Target Pemrosesan", value=target_str, inline=True)

        mode_map = {
            "full": "🚀 Full Cronjob (Scrape + Drive + SOT)",
            "sync_only": "⚡ Sync SOT Saja (Tanpa Scraping Ulang)",
            "dry_run": "🔍 Dry Run (Simulasi Antrean)"
        }
        embed.add_field(name="⚙️ Mode Eksekusi", value=mode_map.get(self.mode, self.mode), inline=True)

        resume_str = "✅ Aktif (Lanjutkan checkpoint)" if self.resume else "❌ Nonaktif (Dari awal)"
        embed.add_field(name="🔁 Resume Checkpoint", value=resume_str, inline=True)

        # Checkpoint Hari Ini jika ada
        cp = get_today_sot_checkpoint()
        if cp:
            done_cnt = len(cp.get("completed_owners", []))
            fail_cnt = len(cp.get("failed_owners", []))
            embed.add_field(
                name="📌 Checkpoint Hari Ini",
                value=f"⚡ Progres: **{done_cnt}** Selesai / **{fail_cnt}** Gagal",
                inline=False
            )

        # Laporan Terakhir
        latest = get_latest_sot_report()
        if latest:
            rep, fname = latest
            d_str = rep.get("date", "-")
            succ = len(rep.get("completed_owners", []))
            fail = len(rep.get("failed_owners", []))
            dur = rep.get("duration", "-")
            sh_stat = rep.get("sheet_sync", {}).get("status", "-")
            tot_rows = rep.get("sheet_sync", {}).get("totalRows", "-")
            embed.add_field(
                name="🕒 Laporan Eksekusi Terakhir",
                value=f"📅 `{d_str}` • ✅ **{succ}** Sukses / ❌ **{fail}** Gagal\n⏱️ Durasi: `{dur}` • Status SOT: `{sh_stat}` ({tot_rows} baris)",
                inline=False
            )

        # Status Pipeline
        if _sot_running_info["is_running"]:
            s_user = _sot_running_info.get("user", "Pengguna lain")
            status_text = f"🟡 **Sedang Berjalan:** Penarikan SOT sedang aktif diproses (oleh `{s_user}`)"
        elif is_pipeline_locked():
            lock_info = get_lock_info()
            busy_owner = lock_info.get("owner", "Owner lain")
            busy_user = lock_info.get("user", "Pengguna lain")
            status_text = f"🟡 **Pipeline Terkunci:** Sedang memproses `{busy_owner}` (oleh `{busy_user}`)"
        else:
            status_text = "🟢 **Idle:** Siap menjalankan cronjob"
        embed.add_field(name="🛡️ Status Sistem", value=status_text, inline=False)

        embed.add_field(
            name="📑 Tautan Spreadsheet SOT",
            value=f"[Buka Google Spreadsheet Tab SOT]({SOT_SPREADSHEET_URL})",
            inline=False
        )

        embed.set_footer(text="Superfood Tech • Cronjob SOT Management", icon_url=DRIVE_ICON_URL)
        return embed


class SotStartButton(discord.ui.Button):
    def __init__(self, parent_view: SotControlView):
        super().__init__(label="Mulai Tarik Cronjob", style=discord.ButtonStyle.success, emoji="🚀", row=0)
        self.parent_view = parent_view

    async def callback(self, interaction: discord.Interaction):
        if interaction.user.id != self.parent_view.user.id:
            try:
                await interaction.response.send_message("❌ Hanya pengguna yang membuka panel ini yang dapat memulainya.", ephemeral=True)
            except Exception:
                pass
            return
        await interaction.response.defer()
        await execute_sot_cronjob(interaction, target_owner=self.parent_view.target_owner, resume=self.parent_view.resume)


class SotSyncOnlyButton(discord.ui.Button):
    def __init__(self, parent_view: SotControlView):
        super().__init__(label="Sync SOT Saja", style=discord.ButtonStyle.primary, emoji="⚡", row=0)
        self.parent_view = parent_view

    async def callback(self, interaction: discord.Interaction):
        if interaction.user.id != self.parent_view.user.id:
            try:
                await interaction.response.send_message("❌ Hanya pengguna yang membuka panel ini yang dapat memulainya.", ephemeral=True)
            except Exception:
                pass
            return
        await execute_sot_sync_only(interaction, target_owner=self.parent_view.target_owner)


class SotDryRunButton(discord.ui.Button):
    def __init__(self, parent_view: SotControlView):
        super().__init__(label="Dry Run Target", style=discord.ButtonStyle.secondary, emoji="🔍", row=0)
        self.parent_view = parent_view

    async def callback(self, interaction: discord.Interaction):
        if interaction.user.id != self.parent_view.user.id:
            try:
                await interaction.response.send_message("❌ Hanya pengguna yang membuka panel ini yang dapat memulainya.", ephemeral=True)
            except Exception:
                pass
            return
        await execute_sot_dry_run(interaction, target_owner=self.parent_view.target_owner)


class SotCancelButton(discord.ui.Button):
    def __init__(self, parent_view: SotControlView):
        super().__init__(label="Batal", style=discord.ButtonStyle.danger, emoji="❌", row=1)
        self.parent_view = parent_view

    async def callback(self, interaction: discord.Interaction):
        if interaction.user.id != self.parent_view.user.id:
            try:
                await interaction.response.send_message("❌ Hanya pengguna yang membuka panel ini yang dapat membatalkannya.", ephemeral=True)
            except Exception:
                pass
            return
        cancel_embed = discord.Embed(
            title="🛑 PANEL SOT DITUTUP",
            description="Panel kontrol cronjob SOT telah ditutup.",
            color=THEME_ERROR,
            timestamp=datetime.datetime.now()
        )
        try:
            await interaction.response.edit_message(embed=cancel_embed, view=None)
        except Exception:
            pass


# ─── Eksekusi Proses SOT ──────────────────────────────────────────────────────

async def execute_sot_cronjob(interaction: discord.Interaction, target_owner: str = None, resume: bool = True):
    """Menjalankan batch cronjob weekly_batch_sync.py di subprocess terisolasi dengan live log terminal."""
    if _sot_running_info["is_running"]:
        busy_user = _sot_running_info.get("user", "Pengguna lain")
        embed = discord.Embed(
            title="⏳ SINKRONISASI SOT SEDANG BERJALAN",
            description=f"Batch Cronjob SOT saat ini sudah sedang berjalan (dimulai oleh **{busy_user}**).\nHarap tunggu hingga proses selesai.",
            color=THEME_PROGRESS,
            timestamp=datetime.datetime.now()
        )
        try:
            await interaction.edit_original_response(embed=embed, view=None)
        except Exception:
            pass
        return

    if is_pipeline_locked():
        lock_info = get_lock_info()
        busy_owner = lock_info.get("owner", "Owner lain")
        busy_user = lock_info.get("user", "Pengguna lain")
        embed = discord.Embed(
            title="⏳ PIPELINE SEDANG DIGUNAKAN",
            description=f"Pipeline saat ini sedang memproses **{busy_owner}** (oleh **{busy_user}**).\nSilakan tunggu hingga pipeline idle.",
            color=THEME_PROGRESS,
            timestamp=datetime.datetime.now()
        )
        try:
            await interaction.edit_original_response(embed=embed, view=None)
        except Exception:
            pass
        return

    _sot_running_info["is_running"] = True
    _sot_running_info["user"] = interaction.user.display_name
    _sot_running_info["started_at"] = datetime.datetime.now()
    cancel_event = threading.Event()
    _sot_running_info["cancel_event"] = cancel_event
    proc_holder = {"proc": None}

    running_view = SotRunningView(user=interaction.user, cancel_event=cancel_event, proc_holder=proc_holder)

    owner_desc = f"**{target_owner}**" if target_owner and target_owner != "__ALL__" else "**Semua Owner**"
    progress_embed = discord.Embed(
        title="⚙️ MENARIK CRONJOB SINKRONISASI SOT...",
        description=f"Sedang menjalankan batch sync untuk {owner_desc}...\n\n{create_progress_bar(5)}",
        color=THEME_PROGRESS,
        timestamp=datetime.datetime.now()
    )
    progress_embed.set_thumbnail(url="https://i.gifer.com/ZZ5H.gif")
    progress_embed.add_field(
        name="💻 Terminal Logs",
        value="```ansi\n[00:00] \u001b[34m[INIT]\u001b[0m Memulai subprocess weekly_batch_sync.py...\n```",
        inline=False
    )

    try:
        await interaction.edit_original_response(embed=progress_embed, view=running_view)
    except Exception:
        pass

    terminal_lines = [f"[{datetime.datetime.now().strftime('%H:%M:%S')}] \u001b[34m[START]\u001b[0m Cronjob SOT dimulai."]
    last_update_time = datetime.datetime.now()
    edit_lock = asyncio.Lock()
    is_finished = False

    cmd = [sys.executable, "-u", str(BASE_DIR / "cronjobsot" / "weekly_batch_sync.py")]
    if target_owner and target_owner != "__ALL__":
        cmd.extend(["--owner", target_owner])
    if not resume:
        cmd.append("--no-resume")

    loop = asyncio.get_running_loop()

    def run_process():
        proc = subprocess.Popen(
            cmd,
            cwd=str(BASE_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1
        )
        proc_holder["proc"] = proc
        _sot_running_info["process"] = proc

        for line in iter(proc.stdout.readline, ""):
            if not line:
                break
            raw_line = line.rstrip()
            if not raw_line:
                continue

            now_str = datetime.datetime.now().strftime("%H:%M:%S")
            if "SUKSES" in raw_line or "✅" in raw_line or "Berhasil" in raw_line:
                colored = f"[{now_str}] \u001b[32m[OK]\u001b[0m {raw_line}"
            elif "GAGAL" in raw_line or "❌" in raw_line or "Error" in raw_line or "Exception" in raw_line:
                colored = f"[{now_str}] \u001b[31m[ERR]\u001b[0m {raw_line}"
            elif "Memproses" in raw_line or "👤" in raw_line:
                colored = f"[{now_str}] \u001b[36m[PROC]\u001b[0m {raw_line}"
            elif "SKIP" in raw_line:
                colored = f"[{now_str}] \u001b[33m[SKIP]\u001b[0m {raw_line}"
            else:
                colored = f"[{now_str}] \u001b[37m{raw_line}\u001b[0m"

            asyncio.run_coroutine_threadsafe(update_progress(colored), loop)

            if cancel_event.is_set():
                proc.terminate()
                try:
                    proc.wait(timeout=5)
                except Exception:
                    proc.kill()
                break

        proc.wait()
        return proc.returncode

    async def update_progress(colored_line):
        nonlocal last_update_time
        if is_finished or cancel_event.is_set():
            return
        terminal_lines.append(colored_line)
        if len(terminal_lines) > 8:
            terminal_lines.pop(0)

        if (datetime.datetime.now() - last_update_time).total_seconds() >= 2.0:
            last_update_time = datetime.datetime.now()
            log_text = "\n".join(terminal_lines)
            # Discord membatasi value embed field hingga 1024 karakter.
            log_text = log_text[-900:]
            progress_embed.set_field_at(
                0,
                name="💻 Terminal Logs",
                value=f"```ansi\n{log_text}\n```",
                inline=False
            )
            async with edit_lock:
                try:
                    await interaction.edit_original_response(embed=progress_embed, view=running_view)
                except Exception:
                    pass

    try:
        returncode = await loop.run_in_executor(None, run_process)
    except Exception as e:
        returncode = -1
        terminal_lines.append(f"❌ Exception: {e}")
    finally:
        is_finished = True
        _sot_running_info["is_running"] = False
        _sot_running_info["process"] = None
        _sot_running_info["cancel_event"] = None

    if cancel_event.is_set():
        cancel_embed = discord.Embed(
            title="🛑 CRONJOB SOT DIBATALKAN",
            description=f"Penarikan cronjob SOT untuk {owner_desc} telah dihentikan atas permintaan pengguna.",
            color=THEME_ERROR,
            timestamp=datetime.datetime.now()
        )
        try:
            await interaction.edit_original_response(embed=cancel_embed, view=None)
        except Exception:
            try:
                await interaction.channel.send(embed=cancel_embed)
            except Exception:
                pass
        return

    if returncode == 0:
        latest = get_latest_sot_report()
        succ = "-"
        fail = "-"
        dur = "-"
        sh_res = "-"
        if latest:
            rep, _ = latest
            succ = len(rep.get("completed_owners", []))
            fail = len(rep.get("failed_owners", []))
            dur = rep.get("duration", "-")
            sh_data = rep.get("sheet_sync", {})
            sh_res = f"Ditimpa: {sh_data.get('total_updated', 0)} | Baru: {sh_data.get('total_inserted', 0)} | Total: {sh_data.get('totalRows', '-')}"

        success_embed = discord.Embed(
            title="🎉 CRONJOB SINKRONISASI SOT SELESAI!",
            description=f"Penarikan data otomatis mingguan untuk {owner_desc} telah selesai diproses.",
            color=THEME_SUCCESS,
            timestamp=datetime.datetime.now()
        )
        success_embed.set_thumbnail(url=SUCCESS_THUMBNAIL)
        success_embed.add_field(name="⏱️ Total Durasi", value=f"`{dur}`", inline=True)
        success_embed.add_field(name="👥 Status Owner", value=f"✅ **{succ}** Sukses / ❌ **{fail}** Gagal", inline=True)
        success_embed.add_field(name="📊 Google Sheet Tab 'SOT'", value=f"{sh_res}", inline=False)
        success_embed.add_field(name="📑 Link Spreadsheet", value=f"[👉 Buka Google Spreadsheet Tab SOT]({SOT_SPREADSHEET_URL})", inline=False)
        success_embed.add_field(name="📁 Root Google Drive", value=f"[👉 Buka Root Google Drive]({ROOT_DRIVE_URL})", inline=False)
        success_embed.set_footer(text="Superfood Tech • SOT Weekly Batch Engine", icon_url=DRIVE_ICON_URL)

        class SotResultView(discord.ui.View):
            def __init__(self):
                super().__init__(timeout=None)
                self.add_item(discord.ui.Button(label="Buka Tab SOT", url=SOT_SPREADSHEET_URL, style=discord.ButtonStyle.link, emoji="📑"))
                self.add_item(discord.ui.Button(label="Buka Google Drive", url=ROOT_DRIVE_URL, style=discord.ButtonStyle.link, emoji="📁"))

        try:
            await interaction.edit_original_response(embed=success_embed, view=SotResultView())
        except Exception:
            try:
                await interaction.channel.send(embed=success_embed, view=SotResultView())
            except Exception:
                pass
    else:
        err_embed = discord.Embed(
            title="❌ CRONJOB SOT GAGAL / BERHENTI DENGAN ERROR",
            description=f"Proses batch sync keluar dengan kode status `{returncode}`.\nPeriksa log terminal atau laporan checkpoint.",
            color=THEME_ERROR,
            timestamp=datetime.datetime.now()
        )
        last_logs = "\n".join(terminal_lines[-6:]) if terminal_lines else "Tidak ada log tercatat."
        err_embed.add_field(name="💻 Potongan Log Terakhir", value=f"```ansi\n{last_logs}\n```", inline=False)
        try:
            await interaction.edit_original_response(embed=err_embed, view=None)
        except Exception:
            try:
                await interaction.channel.send(embed=err_embed)
            except Exception:
                pass


async def execute_sot_dry_run(interaction: discord.Interaction, target_owner: str = None):
    """Menjalankan simulasi dry-run untuk melihat daftar target antrean owner."""
    cmd = [sys.executable, "-u", str(BASE_DIR / "cronjobsot" / "weekly_batch_sync.py"), "--dry-run"]
    if target_owner and target_owner != "__ALL__":
        cmd.extend(["--owner", target_owner])

    loading_embed = discord.Embed(
        title="🔍 MEMUAT SIMULASI CRONJOB (DRY RUN)...",
        description="Sedang memuat data dari Tab DBR Google Sheet...",
        color=THEME_ACCENT,
        timestamp=datetime.datetime.now()
    )
    try:
        if interaction.response.is_done():
            await interaction.edit_original_response(embed=loading_embed, view=None)
        else:
            await interaction.response.edit_message(embed=loading_embed, view=None)
    except Exception:
        pass

    loop = asyncio.get_running_loop()
    proc = await loop.run_in_executor(
        None,
        lambda: subprocess.run(
            cmd,
            cwd=str(BASE_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace"
        )
    )

    output = proc.stdout
    lines = output.splitlines()
    target_lines = [l for l in lines if re.match(r"^\s*\d+\.", l)]

    total_target = len(target_lines)
    preview = "\n".join(target_lines[:15])
    if len(target_lines) > 15:
        preview += f"\n... dan {len(target_lines) - 15} owner lainnya."

    dry_embed = discord.Embed(
        title="🔍 HASIL SIMULASI TARGET CRONJOB (DRY RUN)",
        description=f"Terdeteksi total **{total_target}** owner yang siap diproses.",
        color=THEME_ACCENT,
        timestamp=datetime.datetime.now()
    )
    if preview:
        dry_embed.add_field(name="📋 Daftar Owner & Distribusi Platform", value=f"```\n{preview}\n```", inline=False)
    dry_embed.add_field(name="ℹ️ Catatan", value="Mode Dry Run tidak melakukan scraping, upload drive, atau pembaruan spreadsheet.", inline=False)
    dry_embed.set_footer(text="Gunakan tombol di bawah jika ingin melanjutkan penarikan.", icon_url=DRIVE_ICON_URL)

    class DryRunResultView(discord.ui.View):
        def __init__(self, parent_user: discord.User, owner: str):
            super().__init__(timeout=300)
            self.parent_user = parent_user
            self.owner = owner

        @discord.ui.button(label="Mulai Tarik Cronjob Sekarang", style=discord.ButtonStyle.success, emoji="🚀")
        async def run_btn(self, inter: discord.Interaction, btn: discord.ui.Button):
            if inter.user.id != self.parent_user.id:
                try:
                    await inter.response.send_message("❌ Hanya pemanggil awal yang dapat menjalankan.", ephemeral=True)
                except Exception:
                    pass
                return
            await inter.response.defer()
            await execute_sot_cronjob(inter, target_owner=self.owner)

    try:
        await interaction.edit_original_response(embed=dry_embed, view=DryRunResultView(interaction.user, target_owner))
    except Exception:
        pass


async def execute_sot_sync_only(interaction: discord.Interaction, target_owner: str = None):
    """Menjalankan sinkronisasi master ke tab SOT secara cepat tanpa scraping ulang."""
    cmd = [sys.executable, "-u", str(BASE_DIR / "cronjobsot" / "export_to_sot.py")]
    if target_owner and target_owner != "__ALL__":
        cmd.extend(["--owner", target_owner])

    loading_embed = discord.Embed(
        title="⚡ SINKRONISASI KE TAB SOT...",
        description="Sedang memuat data master lokal dan mengirim ke Tab 'SOT' Google Spreadsheet...",
        color=THEME_PROGRESS,
        timestamp=datetime.datetime.now()
    )
    try:
        if interaction.response.is_done():
            await interaction.edit_original_response(embed=loading_embed, view=None)
        else:
            await interaction.response.edit_message(embed=loading_embed, view=None)
    except Exception:
        pass

    loop = asyncio.get_running_loop()
    proc = await loop.run_in_executor(
        None,
        lambda: subprocess.run(
            cmd,
            cwd=str(BASE_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace"
        )
    )

    out = proc.stdout
    if proc.returncode == 0:
        res_embed = discord.Embed(
            title="✅ SINKRONISASI TAB SOT BERHASIL!",
            description="Data outlet berhasil di-upsert ke Google Spreadsheet Tab 'SOT'.",
            color=THEME_SUCCESS,
            timestamp=datetime.datetime.now()
        )
        res_embed.add_field(name="📄 Output Ringkas", value=f"```\n{out[-500:]}\n```", inline=False)
        res_embed.add_field(name="📑 Link Spreadsheet", value=f"[👉 Buka Tab SOT di Google Spreadsheet]({SOT_SPREADSHEET_URL})", inline=False)
    else:
        res_embed = discord.Embed(
            title="❌ SINKRONISASI TAB SOT GAGAL",
            description=f"Gagal saat mengeksekusi export_to_sot.py (Kode: {proc.returncode}).\n```\n{out[-500:]}\n```",
            color=THEME_ERROR,
            timestamp=datetime.datetime.now()
        )

    try:
        await interaction.edit_original_response(embed=res_embed, view=None)
    except Exception:
        pass


# ─── Bot Client & Setup ───────────────────────────────────────────────────────

class SOTBot(commands.Bot):
    def __init__(self):
        intents = discord.Intents.default()
        super().__init__(command_prefix="!", intents=intents)

    async def setup_hook(self):
        await self.tree.sync()
        print("[✓] [BOT-SOT] Slash Commands /sot dan /status_sot berhasil disinkronkan ke Discord!")

    async def on_ready(self):
        print("=" * 60)
        print(f"🤖 Bot SOT Login Berhasil sebagai: {self.user.name} ({self.user.id})")
        print(f"🌟 Status: Online & Siap Menangani Perintah /sot & /status_sot!")
        print(f"🎯 Target Sheet: Tab '{TARGET_SHEET_NAME}' ({SOT_SPREADSHEET_URL})")
        print("=" * 60)
        try:
            await self.change_presence(activity=None)
        except Exception:
            pass

        # Jika DISCORD_CHANNEL_ID_SOT disetel, kirim panel awal ke channel tersebut
        if CHANNEL_ID and CHANNEL_ID.isdigit():
            try:
                target_channel = self.get_channel(int(CHANNEL_ID))
                if target_channel:
                    owners_meta = await get_cached_owners()
                    panel_view = SotControlView(self.user, owners_meta=owners_meta)
                    embed_init = panel_view.build_embed()
                    await target_channel.send(embed=embed_init, view=panel_view)
                    print(f"[✓] [BOT-SOT] Control Panel SOT otomatis dikirim ke channel #{target_channel.name} ({CHANNEL_ID})")
            except Exception as e:
                print(f"⚠️ [BOT-SOT] Gagal mengirim pesan otomatis ke channel ID {CHANNEL_ID}: {e}")


bot_sot = SOTBot()


@bot_sot.tree.error
async def on_sot_command_error(interaction: discord.Interaction, error: app_commands.AppCommandError):
    if isinstance(error, app_commands.CommandInvokeError):
        original = error.original
        if isinstance(original, discord.NotFound) and getattr(original, "code", None) == 10062:
            return
        if isinstance(original, discord.HTTPException) and getattr(original, "code", None) == 40060:
            return
    print(f"⚠️ [BOT-SOT APP COMMAND ERROR] {error}")


@bot_sot.tree.command(name="sot", description="📊 Tarik cronjob sinkronisasi mingguan ke Google Spreadsheet SOT & Google Drive")
@app_commands.describe(
    owner="Nama Owner (opsional: proses khusus owner tertentu, default: semua owner)",
    mode="Mode eksekusi cronjob",
    resume="Lanjutkan dari checkpoint hari ini jika tersedia (default: True)"
)
@app_commands.choices(mode=[
    app_commands.Choice(name="🚀 Full Cronjob (Scrape + Drive + SOT Sync)", value="full"),
    app_commands.Choice(name="⚡ Sync SOT Saja (Tanpa Scraping Ulang)", value="sync_only"),
    app_commands.Choice(name="🔍 Dry Run (Simulasi Antrean Target)", value="dry_run")
])
async def sot_slash(
    interaction: discord.Interaction,
    owner: str = None,
    mode: str = "full",
    resume: bool = True
):
    """Menampilkan Control Panel SOT atau mengeksekusi langsung sesuai mode."""
    try:
        await interaction.response.defer()
    except (discord.NotFound, discord.HTTPException) as e:
        print(f"⚠️ [SLASH /sot] Interaksi tidak dapat di-defer ({e}).")
        return

    if mode == "dry_run":
        await execute_sot_dry_run(interaction, target_owner=owner)
        return

    if mode == "sync_only":
        await execute_sot_sync_only(interaction, target_owner=owner)
        return

    owners_meta = await get_cached_owners()
    view = SotControlView(interaction.user, target_owner=owner, mode=mode, resume=resume, owners_meta=owners_meta)
    embed = view.build_embed()
    try:
        await interaction.followup.send(embed=embed, view=view)
    except (discord.NotFound, discord.HTTPException) as e:
        print(f"⚠️ [SLASH /sot] Gagal mengirim panel followup ({e}).")


@sot_slash.autocomplete("owner")
async def sot_owner_autocomplete(interaction: discord.Interaction, current: str):
    owners = await get_cached_owners()
    choices = []
    current_lower = (current or "").strip().lower()

    if not current_lower or "semua" in current_lower or "all" in current_lower:
        choices.append(app_commands.Choice(name="📦 [Semua Owner Terdaftar]", value="__ALL__"))

    for m in owners:
        owner_name = m.get("owner", "")
        brand = m.get("brand", "")
        if not current_lower or current_lower in owner_name.lower() or (brand and current_lower in brand.lower()):
            label = f"{owner_name} ({brand})" if brand else owner_name
            if len(label) > 100:
                label = label[:97] + "..."
            choices.append(app_commands.Choice(name=label, value=owner_name))
            if len(choices) >= 25:
                break
    return choices


@bot_sot.tree.command(name="status_sot", description="📈 Cek status data tab 'SOT' dan riwayat sinkronisasi mingguan")
async def status_sot_slash(interaction: discord.Interaction):
    """Menampilkan status ringkas tab SOT, checkpoint aktif, dan riwayat laporan."""
    try:
        await interaction.response.defer()
    except Exception:
        pass

    embed = discord.Embed(
        title="📈 STATUS TAB 'SOT' & HISTORI SINKRONISASI",
        color=THEME_PRIMARY,
        timestamp=datetime.datetime.now()
    )
    embed.set_thumbnail(url=STORE_THUMBNAIL)
    embed.add_field(name="📑 Target Sheet", value=f"`{TARGET_SHEET_NAME}` (GID: `{TARGET_SHEET_GID}`)", inline=True)
    embed.add_field(name="📁 Spreadsheet ID", value=f"`{TARGET_SPREADSHEET_ID[:12]}...`", inline=True)

    # Checkpoint Hari Ini
    cp = get_today_sot_checkpoint()
    if cp:
        done_cnt = len(cp.get("completed_owners", []))
        fail_cnt = len(cp.get("failed_owners", []))
        embed.add_field(
            name="📌 Checkpoint Hari Ini",
            value=f"⚡ Progres: **{done_cnt}** Selesai / **{fail_cnt}** Gagal",
            inline=False
        )

    # Laporan Mingguan Terakhir
    report_data = get_latest_sot_report()
    if report_data:
        rep, filename = report_data
        r_date = rep.get("date", "-")
        r_dur = rep.get("duration", "-")
        r_tot = rep.get("total_owners", 0)
        r_ok = len(rep.get("completed_owners", []))
        r_fail = len(rep.get("failed_owners", []))
        sheet_sync = rep.get("sheet_sync", {})
        s_status = sheet_sync.get("status", "N/A")
        s_updated = sheet_sync.get("total_updated", "-")
        s_inserted = sheet_sync.get("total_inserted", "-")
        s_total = sheet_sync.get("totalRows", "-")

        embed.add_field(
            name="📅 Laporan Mingguan Terakhir",
            value=f"**Tanggal:** `{r_date}` ({r_dur})\n**Status Owner:** ✅ {r_ok} Sukses / ❌ {r_fail} Gagal (Total {r_tot})",
            inline=False
        )
        embed.add_field(
            name="📊 Sinkronisasi SOT Terakhir",
            value=f"• Status: `{s_status}`\n• Updated: **{s_updated}** | Inserted: **{s_inserted}**\n• Total Baris SOT: **{s_total}**\n• File: `{filename}`",
            inline=False
        )
    else:
        embed.add_field(
            name="📅 Laporan Terakhir",
            value="*Belum ada laporan mingguan yang tersimpan di server.*",
            inline=False
        )

    # Status Lock
    if _sot_running_info["is_running"]:
        s_user = _sot_running_info.get("user", "Pengguna lain")
        pipeline_status = f"🟡 **Sedang Berjalan:** Cronjob SOT aktif diproses (oleh `{s_user}`)"
    elif is_pipeline_locked():
        lock = get_lock_info()
        pipeline_status = f"🟡 **Pipeline Terkunci:** `{lock.get('owner', '-')}` oleh `{lock.get('user', '-')}`"
    else:
        pipeline_status = "🟢 **Idle:** Siap menyinkronkan"
    embed.add_field(name="⚙️ Status Pipeline", value=pipeline_status, inline=False)

    embed.add_field(
        name="🔗 Tautan Cepat",
        value=f"[👉 Buka Google Spreadsheet Tab SOT]({SOT_SPREADSHEET_URL})\n[👉 Buka Root Google Drive]({ROOT_DRIVE_URL})",
        inline=False
    )
    embed.set_footer(text="Gunakan /sot untuk membuka panel kontrol atau memulai sinkronisasi.")

    class StatusResultView(discord.ui.View):
        def __init__(self):
            super().__init__(timeout=None)
            self.add_item(discord.ui.Button(label="Buka Tab SOT", url=SOT_SPREADSHEET_URL, style=discord.ButtonStyle.link, emoji="📑"))
            self.add_item(discord.ui.Button(label="Buka Google Drive", url=ROOT_DRIVE_URL, style=discord.ButtonStyle.link, emoji="📁"))

    try:
        await interaction.followup.send(embed=embed, view=StatusResultView())
    except Exception:
        pass


def main():
    placeholder_tokens = ["masukkan_token_bot_anda_disini", "your_token_here", "token_bot_anda"]
    if not BOT_TOKEN or BOT_TOKEN.strip() in placeholder_tokens:
        print("\n" + "=" * 65)
        print("❌ DISCORD_BOT_TOKEN_SOT belum disetel di file .env!")
        print("=" * 65)
        print("Untuk menjalankan Bot SOT secara terpisah:")
        print("1. Buat Bot baru di Discord Developer Portal:")
        print("   https://discord.com/developers/applications")
        print("2. Dapatkan Bot Token dari menu 'Bot' -> 'Reset Token'")
        print("3. Tambahkan ke file .env:")
        print("   DISCORD_BOT_TOKEN_SOT=token_bot_sot_anda")
        print("   DISCORD_CHANNEL_ID_SOT=id_channel_sot (opsional)")
        print("=" * 65 + "\n")
        sys.exit(1)

    try:
        bot_sot.run(BOT_TOKEN)
    except discord.errors.LoginFailure:
        print("\n" + "=" * 65)
        print("❌ GAGAL LOGIN KE DISCORD (BOT-SOT): Token Tidak Valid!")
        print("=" * 65)
        sys.exit(1)


if __name__ == "__main__":
    main()

