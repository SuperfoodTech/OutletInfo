#!/usr/bin/env python3
"""
bot_sot.py
==========
Discord Bot khusus untuk /sot (Sinkronisasi Data Listing ke Google Spreadsheet Tab 'SOT').
Menjalankan layanan bot mandiri untuk channel #sot.

Fitur:
- Slash Command `/sot`:
  * Parameter `sumber`: Master Lokal (GoFood, Grab, Shopee) atau Tab DBR
  * Parameter `owner`: Opsional, filter nama owner spesifik (dengan Autocomplete)
- Slash Command `/status_sot`:
  * Cek laporan sinkronisasi mingguan terakhir dan status tab SOT
- Control Panel Interaktif SOT:
  * Tombol `🔄 Sinkronkan ke Tab SOT`
  * Tombol `📊 Cek Status SOT`
  * Tombol Link `🌐 Buka Spreadsheet SOT`
- Proteksi pipeline menggunakan JobLock agar tidak bentrok dengan ekspor lain
"""

import os
import sys
import time
import json
import glob
import asyncio
import datetime
from pathlib import Path
from dotenv import load_dotenv

import discord
from discord import app_commands
from discord.ext import commands

# ─── Path & Environment Setup ─────────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent
CRON_DIR = BASE_DIR / "cronjobsot"
REPORTS_DIR = CRON_DIR / "reports"

if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))
if str(CRON_DIR) not in sys.path:
    sys.path.insert(0, str(CRON_DIR))

load_dotenv(BASE_DIR / ".env")

BOT_TOKEN = (os.getenv("DISCORD_BOT_TOKEN_SOT") or os.getenv("DISCORD_BOT_TOKEN", "")).strip()
CHANNEL_ID = (os.getenv("DISCORD_CHANNEL_ID_SOT") or "").strip()

# Import SOT Modules & Lock
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

SPREADSHEET_URL = f"https://docs.google.com/spreadsheets/d/{TARGET_SPREADSHEET_ID}/edit#gid={TARGET_SHEET_GID}"

# ─── Visual Theme & Helpers ───────────────────────────────────────────────────
THEME_PRIMARY = 0x5865F2   # Blurple
THEME_SUCCESS = 0x10B981   # Emerald Green
THEME_WARNING = 0xF59E0B   # Amber Yellow
THEME_ERROR   = 0xEF4444   # Rose Red
THEME_INFO    = 0x06B6D4   # Cyan

SHEET_ICON_URL = "https://cdn-icons-png.flaticon.com/512/281/281761.png"

# Cache metadata owner untuk autocomplete
_owner_cache = {"data": [], "ts": 0.0}
_owner_lock = asyncio.Lock()


async def get_cached_owners(max_age=180.0):
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


def get_latest_sync_report():
    """Mengambil laporan sinkronisasi SOT terbaru dari folder reports/."""
    report_files = sorted(glob.glob(str(REPORTS_DIR / "weekly_report_*.json")), key=os.path.getmtime)
    if not report_files:
        return None
    try:
        with open(report_files[-1], "r", encoding="utf-8") as f:
            return json.load(f), Path(report_files[-1]).name
    except Exception:
        return None


# ─── Worker Task Sinkronisasi SOT ─────────────────────────────────────────────

def _run_sot_sync_sync(source: str = "master_lokal", owner: str = None) -> tuple[bool, dict]:
    """Fungsi sinkron untuk ekspor SOT yang dijalankan di background thread."""
    if source == "dbr":
        print("[*] [BOT-SOT] Memuat data sumber dari Tab 'DBR'...")
        df = load_dbr_master_data()
    else:
        print("[*] [BOT-SOT] Memuat data sumber dari gabungan Master Lokal...")
        df = load_all_local_masters()
        if df.empty:
            print("[*] [BOT-SOT] Master lokal kosong, fallback ke Tab 'DBR'...")
            df = load_dbr_master_data()

    if df.empty:
        return False, {"error": "Tidak ada data outlet yang dapat dimuat dari sumber."}

    ok, res = export_dataframe_to_sot(df, owner_filter=owner)
    res["source_rows"] = len(df)
    return ok, res


# ─── Interactive Discord Views ────────────────────────────────────────────────

class SOTControlView(discord.ui.View):
    """View Control Panel SOT dengan tombol aksi interaktif."""

    def __init__(self, user: discord.User = None):
        super().__init__(timeout=None)
        self.user = user

        # Tambahkan Link Button langsung ke Spreadsheet SOT
        self.add_item(discord.ui.Button(
            label="Buka Tab SOT",
            style=discord.ButtonStyle.link,
            url=SPREADSHEET_URL,
            emoji="📊"
        ))

    @discord.ui.button(label="Sinkronkan Sekarang", style=discord.ButtonStyle.primary, emoji="🔄", custom_id="sot_sync_btn")
    async def sync_button_callback(self, interaction: discord.Interaction, button: discord.ui.Button):
        await handle_sot_execution(interaction, source="master_lokal", owner=None)

    @discord.ui.button(label="Cek Status SOT", style=discord.ButtonStyle.secondary, emoji="📈", custom_id="sot_status_btn")
    async def status_button_callback(self, interaction: discord.Interaction, button: discord.ui.Button):
        await handle_status_sot_execution(interaction)


# ─── Handlers Eksekusi ────────────────────────────────────────────────────────

async def handle_sot_execution(interaction: discord.Interaction, source: str = "master_lokal", owner: str = None):
    """Menjalankan sinkronisasi SOT dengan visual progress & embed hasil."""
    if not interaction.response.is_done():
        try:
            await interaction.response.defer(thinking=True)
        except Exception:
            pass

    # Cek apakah pipeline sedang terkunci
    if is_pipeline_locked():
        lock = get_lock_info()
        busy_by = lock.get("user", "Pengguna lain")
        busy_owner = lock.get("owner", "Proses lain")
        embed_busy = discord.Embed(
            title="⚠️ PIPELINE SEDANG DIGUNAKAN",
            description=f"Saat ini pipeline sedang memproses `{busy_owner}` oleh `{busy_by}`.\nHarap tunggu sebentar hingga proses selesai sebelum menyinkronkan SOT.",
            color=THEME_WARNING,
            timestamp=datetime.datetime.now()
        )
        await interaction.followup.send(embed=embed_busy, ephemeral=True)
        return

    source_label = "Gabungan Master Lokal (GoFood, Grab, Shopee)" if source == "master_lokal" else "Tab DBR Google Sheet"
    owner_label = f"`{owner}`" if owner else "*Seluruh Owner Terdaftar*"

    embed_progress = discord.Embed(
        title="⏳ SEDANG MENYINKRONKAN KE GOOGLE SPREADSHEET 'SOT'",
        description="Mohon tunggu, data sedang dikompilasi dan dikirim ke Apps Script...",
        color=THEME_WARNING,
        timestamp=datetime.datetime.now()
    )
    embed_progress.set_thumbnail(url=SHEET_ICON_URL)
    embed_progress.add_field(name="📂 Sumber Data", value=source_label, inline=False)
    embed_progress.add_field(name="👤 Filter Owner", value=owner_label, inline=True)
    embed_progress.add_field(name="🌐 Target", value=f"[Spreadsheet SOT]({SPREADSHEET_URL})", inline=True)
    embed_progress.set_footer(text="Proses berlangsung di background • Timeout 180s per batch")

    progress_msg = await interaction.followup.send(embed=embed_progress)

    start_t = time.time()
    try:
        # Jalankan eksekusi dengan JobLock
        with JobLock(owner or "SOT Full Sync", interaction.user.name, timeout_secs=10):
            ok, res = await asyncio.to_thread(_run_sot_sync_sync, source=source, owner=owner)
    except Timeout:
        embed_to = discord.Embed(
            title="⚠️ PIPELINE TERKUNCI",
            description="Gagal mendapatkan izin eksekusi karena pipeline baru saja dipakai proses lain.",
            color=THEME_WARNING
        )
        await progress_msg.edit(embed=embed_to)
        return
    except Exception as ex:
        ok = False
        res = {"error": str(ex)}

    duration = int(time.time() - start_t)
    dur_str = f"{duration // 60}m {duration % 60}s" if duration >= 60 else f"{duration}s"

    if ok:
        updated = res.get("total_updated", 0)
        inserted = res.get("total_inserted", 0)
        total_sot = res.get("totalRows", "-")
        source_rows = res.get("source_rows", "-")

        embed_done = discord.Embed(
            title="✅ SINKRONISASI TAB 'SOT' BERHASIL",
            description=f"Data listing berhasil di-upsert ke Google Spreadsheet tab **'{TARGET_SHEET_NAME}'**.",
            color=THEME_SUCCESS,
            timestamp=datetime.datetime.now()
        )
        embed_done.set_thumbnail(url=SHEET_ICON_URL)
        embed_done.add_field(name="📊 Data Terkirim", value=f"**{source_rows}** baris", inline=True)
        embed_done.add_field(name="🔄 Ditimpa (Updated)", value=f"**{updated}** outlet", inline=True)
        embed_done.add_field(name="✨ Baru (Inserted)", value=f"**{inserted}** outlet", inline=True)
        embed_done.add_field(name="📑 Total Baris SOT", value=f"**{total_sot}** baris", inline=True)
        embed_done.add_field(name="⏱️ Durasi", value=dur_str, inline=True)
        embed_done.add_field(name="👤 Filter", value=owner_label, inline=True)
        embed_done.add_field(name="🔗 Tautan Langsung", value=f"[👉 Buka Tab SOT di Google Spreadsheet]({SPREADSHEET_URL})", inline=False)
        embed_done.set_footer(text=f"Disinkronkan oleh @{interaction.user.name} • Kolom 'Terakhir Diperbaharui' telah diupdate otomatis")

        view = SOTControlView(interaction.user)
        await progress_msg.edit(embed=embed_done, view=view)
    else:
        err_msg = res.get("error", "Terjadi kesalahan yang tidak diketahui.")
        embed_err = discord.Embed(
            title="❌ SINKRONISASI TAB 'SOT' GAGAL",
            description=f"Gagal mengirimkan data outlet ke Google Apps Script:\n```{err_msg}```",
            color=THEME_ERROR,
            timestamp=datetime.datetime.now()
        )
        embed_err.add_field(name="⏱️ Durasi", value=dur_str, inline=True)
        embed_err.add_field(name="🌐 Target", value=f"[Buka Tab SOT]({SPREADSHEET_URL})", inline=True)
        embed_err.set_footer(text="Periksa koneksi Google Apps Script atau coba lagi beberapa saat lagi.")

        view = SOTControlView(interaction.user)
        await progress_msg.edit(embed=embed_done if ok else embed_err, view=view)


async def handle_status_sot_execution(interaction: discord.Interaction):
    """Menampilkan status tab SOT dan histori laporan mingguan terakhir."""
    if not interaction.response.is_done():
        try:
            await interaction.response.defer()
        except Exception:
            pass

    report_data = get_latest_sync_report()

    embed = discord.Embed(
        title="📈 STATUS TAB 'SOT' & HISTORI SINKRONISASI",
        color=THEME_INFO,
        timestamp=datetime.datetime.now()
    )
    embed.set_thumbnail(url=SHEET_ICON_URL)
    embed.add_field(name="📑 Target Sheet", value=f"`{TARGET_SHEET_NAME}` (GID: `{TARGET_SHEET_GID}`)", inline=True)
    embed.add_field(name="📁 Spreadsheet ID", value=f"`{TARGET_SPREADSHEET_ID[:12]}...`", inline=True)

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
            value=f"• Status: `{s_status}`\n• Updated: **{s_updated}** | Inserted: **{s_inserted}**\n• Total Baris SOT: **{s_total}**\n• File Laporan: `{filename}`",
            inline=False
        )
    else:
        embed.add_field(
            name="📅 Laporan Terakhir",
            value="*Belum ada laporan mingguan yang tersimpan di server.*",
            inline=False
        )

    # Status Lock
    if is_pipeline_locked():
        lock = get_lock_info()
        embed.add_field(
            name="⚙️ Status Pipeline",
            value=f"🟡 **Sedang Berjalan:** `{lock.get('owner', '-')}` oleh `{lock.get('user', '-')}`",
            inline=False
        )
    else:
        embed.add_field(name="⚙️ Status Pipeline", value="🟢 **Idle:** Siap menyinkronkan", inline=False)

    embed.add_field(
        name="🔗 Tautan Cepat",
        value=f"[👉 Buka Google Spreadsheet Tab SOT]({SPREADSHEET_URL})",
        inline=False
    )
    embed.set_footer(text="Gunakan /sot untuk memulai sinkronisasi secara manual.")

    view = SOTControlView(interaction.user)
    await interaction.followup.send(embed=embed, view=view)


# ─── Bot Client & Slash Commands ──────────────────────────────────────────────

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
        print(f"🌟 Status: Online & Siap Menangani Perintah /sot!")
        print(f"🎯 Target Sheet: Tab '{TARGET_SHEET_NAME}' ({SPREADSHEET_URL})")
        print("=" * 60)

        await self.change_presence(
            activity=discord.Activity(
                type=discord.ActivityType.watching,
                name="Spreadsheet SOT • /sot"
            )
        )

        # Jika DISCORD_CHANNEL_ID_SOT disetel, kirim panel awal ke channel tersebut
        if CHANNEL_ID and CHANNEL_ID.isdigit():
            try:
                target_channel = self.get_channel(int(CHANNEL_ID))
                if target_channel:
                    embed_init = discord.Embed(
                        title="🟢 BOT SOT ONLINE • PANEL KONTROL SINKRONISASI",
                        description=f"Bot khusus sinkronisasi data outlet ke Google Spreadsheet Tab **'{TARGET_SHEET_NAME}'** aktif.\n\n"
                                    f"Klik tombol di bawah atau ketik `/sot` untuk menyinkronkan data.",
                        color=THEME_SUCCESS,
                        timestamp=datetime.datetime.now()
                    )
                    embed_init.set_thumbnail(url=SHEET_ICON_URL)
                    embed_init.add_field(name="📑 Spreadsheet ID", value=f"`{TARGET_SPREADSHEET_ID}`", inline=True)
                    embed_init.add_field(name="📄 Tab Target", value=f"`{TARGET_SHEET_NAME}` (GID `{TARGET_SHEET_GID}`)", inline=True)
                    embed_init.add_field(name="🌐 Tautan Langsung", value=f"[Buka Tab SOT]({SPREADSHEET_URL})", inline=False)
                    embed_init.set_footer(text="Bot SOT • Otomasi Sinkronisasi Google Spreadsheet")

                    view = SOTControlView(self.user)
                    await target_channel.send(embed=embed_init, view=view)
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


@bot_sot.tree.command(name="sot", description="🚀 Sinkronkan seluruh data outlet listing ke Google Spreadsheet Tab 'SOT'")
@app_commands.describe(
    sumber="Sumber data outlet (Master Lokal gabungan atau Tab DBR Google Sheet)",
    owner="Filter hanya owner tertentu (opsional: ketik untuk mencari nama owner)"
)
@app_commands.choices(sumber=[
    app_commands.Choice(name="📦 Gabungan Master Lokal (GoFood, Grab, Shopee) [Rekomendasi]", value="master_lokal"),
    app_commands.Choice(name="🌐 Data Tab 'DBR' Google Sheet", value="dbr")
])
async def sot_slash(
    interaction: discord.Interaction,
    sumber: app_commands.Choice[str] = None,
    owner: str = None
):
    selected_source = sumber.value if sumber else "master_lokal"
    await handle_sot_execution(interaction, source=selected_source, owner=owner)


@sot_slash.autocomplete("owner")
async def sot_owner_autocomplete(interaction: discord.Interaction, current: str):
    owners = await get_cached_owners()
    choices = []
    current_lower = (current or "").strip().lower()

    if not current_lower or "semua" in current_lower or "all" in current_lower:
        choices.append(app_commands.Choice(name="📦 [Semua Owner Terdaftar]", value=""))

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
    await handle_status_sot_execution(interaction)


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
