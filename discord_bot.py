#!/usr/bin/env python3
"""
discord_bot.py
==============
Discord Bot untuk Otomasi Generate & Upload Outlet Info ke Google Drive.
Fitur:
- Slash Command `/generate`
- Dropdown Pemilihan Aplikator (GoFood / GrabFood / Semua)
- Dropdown Pemilihan Owner (Otomatis mendeteksi dan menaruh Owner TERBARU di posisi paling atas)
- Tombol Trigger `🚀 Generate & Upload`
- Visual Progress Bar (▰▰▰▰▰▰▱▱ % ) & Live ANSI Terminal Console Logs
- Embed Hasil Akhir yang estetik dengan Link langsung ke Folder Owner di Google Drive
"""

import os
import sys
import asyncio
import datetime
from pathlib import Path
from dotenv import load_dotenv

import discord
from discord import app_commands
from discord.ext import commands

# Muat environment
BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

BOT_TOKEN = os.getenv("DISCORD_BOT_TOKEN", "")
CHANNEL_ID = os.getenv("DISCORD_CHANNEL_ID", "").strip()
ROOT_DRIVE_URL = "https://drive.google.com/drive/u/0/folders/19VIrypPcBmNNbjDLGS7kxp_yIdBBwXjB"

# Import helper dari export_and_upload_drive
try:
    from export_and_upload_drive import get_owners_with_metadata, generate_for_owner_pipeline
except ImportError:
    sys.path.append(str(BASE_DIR))
    from export_and_upload_drive import get_owners_with_metadata, generate_for_owner_pipeline


# ─── Visual Helpers & Themes ──────────────────────────────────────────────────
THEME_PRIMARY   = 0x5865F2  # Blurple
THEME_ACCENT    = 0x00E5FF  # Neon Cyan
THEME_PROGRESS  = 0xFEE75C  # Golden Yellow
THEME_SUCCESS   = 0x57F287  # Vibrant Green
THEME_ERROR     = 0xED4245  # Coral Red

STORE_THUMBNAIL = "https://cdn-icons-png.flaticon.com/512/2830/2830305.png"
SUCCESS_THUMBNAIL = "https://cdn-icons-png.flaticon.com/512/190/190411.png"
DRIVE_ICON_URL  = "https://cdn-icons-png.flaticon.com/512/2965/2965300.png"


def create_progress_bar(percent: int, length: int = 15) -> str:
    """Membuat visual progress bar blok modern."""
    filled = int(round(length * percent / 100))
    bar = "▰" * filled + "▱" * (length - filled)
    return f"`{bar}` **{percent}%**"


# ─── Discord Views & UI Components ────────────────────────────────────────────

class AplikatorSelect(discord.ui.Select):
    def __init__(self, parent_view):
        self.parent_view = parent_view
        options = [
            discord.SelectOption(
                label="Semua Platform (GoFood + Grab + Shopee)",
                value="all",
                emoji="🌐",
                default=True
            ),
            discord.SelectOption(
                label="GoFood Saja",
                value="gofood",
                emoji="🔴"
            ),
            discord.SelectOption(
                label="GrabFood Saja",
                value="grab",
                emoji="🟢"
            ),
            discord.SelectOption(
                label="ShopeeFood Saja",
                value="shopee",
                emoji="🟠"
            )
        ]
        super().__init__(
            placeholder="📌 Langkah 1: Pilih Aplikator...",
            min_values=1,
            max_values=1,
            options=options,
            row=0
        )

    async def callback(self, interaction: discord.Interaction):
        self.parent_view.selected_aplikator = self.values[0]
        # Update default state
        for opt in self.options:
            opt.default = (opt.value == self.values[0])
        await self.parent_view.update_panel(interaction)


class OwnerSelect(discord.ui.Select):
    def __init__(self, parent_view, owners_meta):
        self.parent_view = parent_view
        options = []
        
        # Opsi Batch Semua
        options.append(discord.SelectOption(
            label="[Semua Owner Terdaftar]",
            value="__ALL__",
            emoji="📦"
        ))

        # Discord batas maksimal 25 opsi per select (1 batch + 24 owner)
        for meta in owners_meta[:24]:
            owner_name = meta["owner"]
            options.append(discord.SelectOption(
                label=owner_name[:100],
                value=owner_name,
                emoji="👤",
                default=(not parent_view.selected_owner and owner_name == owners_meta[0]["owner"])
            ))

        super().__init__(
            placeholder="👤 Langkah 2: Pilih Owner...",
            min_values=1,
            max_values=1,
            options=options,
            row=1
        )

    async def callback(self, interaction: discord.Interaction):
        self.parent_view.selected_owner = self.values[0]
        for opt in self.options:
            opt.default = (opt.value == self.values[0])
        await self.parent_view.update_panel(interaction)


class GenerateButton(discord.ui.Button):
    def __init__(self, parent_view):
        self.parent_view = parent_view
        super().__init__(
            label="Mulai Generate & Upload",
            style=discord.ButtonStyle.success,
            emoji="⚡",
            row=2
        )

    async def callback(self, interaction: discord.Interaction):
        await self.parent_view.start_generation(interaction)


class RefreshButton(discord.ui.Button):
    def __init__(self, parent_view):
        self.parent_view = parent_view
        super().__init__(
            label="Refresh Data",
            style=discord.ButtonStyle.secondary,
            emoji="🔄",
            row=2
        )

    async def callback(self, interaction: discord.Interaction):
        await interaction.response.defer()
        self.parent_view.reload_metadata()
        await self.parent_view.update_panel(interaction, edit_response=True)


class ControlPanelView(discord.ui.View):
    def __init__(self, user):
        super().__init__(timeout=600)
        self.user = user
        self.selected_aplikator = "all"
        self.selected_owner = None
        self.owners_meta = []
        self.reload_metadata()

    def reload_metadata(self):
        self.clear_items()
        self.owners_meta = get_owners_with_metadata()
        if self.owners_meta and not self.selected_owner:
            # Default ke owner yang paling baru di-generate
            self.selected_owner = self.owners_meta[0]["owner"]
            
        self.aplikator_select = AplikatorSelect(self)
        self.owner_select = OwnerSelect(self, self.owners_meta)
        self.generate_btn = GenerateButton(self)
        self.refresh_btn = RefreshButton(self)

        self.add_item(self.aplikator_select)
        self.add_item(self.owner_select)
        self.add_item(self.generate_btn)
        self.add_item(self.refresh_btn)

    def build_embed(self):
        embed = discord.Embed(
            title="🏪 OUTLET INFO PIPELINE • CONTROL PANEL",
            description=(
                "Panel kontrol otomasi pembuatan laporan data merchant multi-platform.\n"
                "Data akan diformat ke dalam file Excel **2 Tab Identik** (`Listing` & `Listing 2`) "
                "dan diunggah otomatis ke **Google Drive** per-owner."
            ),
            color=THEME_PRIMARY,
            timestamp=datetime.datetime.now()
        )
        embed.set_thumbnail(url=STORE_THUMBNAIL)

        # Mapping Nama Aplikator
        app_names = {
            "all": "🌐 Semua Platform (GoFood + Grab + Shopee)",
            "gofood": "🔴 GoFood Saja",
            "grab": "🟢 GrabFood Saja",
            "shopee": "🟠 ShopeeFood Saja"
        }
        aplikator_label = app_names.get(self.selected_aplikator, self.selected_aplikator)

        # Info owner terpilih tanpa keterangan berbelit
        if self.selected_owner == "__ALL__":
            owner_display = "📦 Semua Owner Terdaftar"
        else:
            owner_display = f"👤 **{self.selected_owner}**"

        embed.add_field(name="📌 Aplikator Terpilih", value=f"**{aplikator_label}**", inline=True)
        embed.add_field(name="👤 Owner Terpilih", value=f"{owner_display}", inline=True)
        embed.add_field(
            name="📁 Tujuan Google Drive",
            value=f"[Buka Folder Induk](https://drive.google.com/drive/u/0/folders/19VIrypPcBmNNbjDLGS7kxp_yIdBBwXjB)",
            inline=True
        )
        embed.add_field(name="⚙️ Status Pipeline", value="`🟢 Siap Dijalankan`", inline=False)

        embed.set_footer(
            text=f"Diminta oleh {self.user.display_name} • Superfood Tech Engine (Vercel Sheet)",
            icon_url=self.user.display_avatar.url
        )
        return embed

    async def update_panel(self, interaction: discord.Interaction, edit_response=False):
        embed = self.build_embed()
        if edit_response:
            await interaction.followup.edit_message(message_id=interaction.message.id, embed=embed, view=self)
        else:
            await interaction.response.edit_message(embed=embed, view=self)

    async def start_generation(self, interaction: discord.Interaction):
        await interaction.response.defer()

        # Build initial progress embed
        progress_embed = discord.Embed(
            title="⚙️ MEMPROSES GENERASI DATA OUTLET...",
            description=f"Sedang memproses owner **{self.selected_owner}**...\n\n{create_progress_bar(10)}",
            color=THEME_PROGRESS,
            timestamp=datetime.datetime.now()
        )
        progress_embed.set_thumbnail(url="https://i.gifer.com/ZZ5H.gif") # Spinner
        progress_embed.add_field(
            name="💻 Terminal Logs",
            value="```ansi\n[00:00] \u001b[34m[INIT]\u001b[0m Memulai pipeline otomasi...\n```",
            inline=False
        )

        # view=None: Menghilangkan seluruh dropdown & tombol dari pesan saat generate dimulai
        message = await interaction.edit_original_response(embed=progress_embed, view=None)

        log_buffer = [f"[{datetime.datetime.now().strftime('%H:%M:%S')}] \u001b[34m[START]\u001b[0m Pipeline dimulai."]
        last_update_time = datetime.datetime.now()

        async def discord_progress_callback(percent: int, log_msg: str):
            nonlocal last_update_time
            now_str = datetime.datetime.now().strftime("%H:%M:%S")
            
            # Beri warna ANSI
            if "Sukses" in log_msg or "✅" in log_msg:
                colored = f"[{now_str}] \u001b[32m[SUCCESS]\u001b[0m {log_msg}"
            elif "⚠️" in log_msg or "Gagal" in log_msg or "❌" in log_msg:
                colored = f"[{now_str}] \u001b[31m[WARN]\u001b[0m {log_msg}"
            elif "Mengunggah" in log_msg:
                colored = f"[{now_str}] \u001b[35m[UPLOAD]\u001b[0m {log_msg}"
            else:
                colored = f"[{now_str}] \u001b[36m[STEP]\u001b[0m {log_msg}"

            log_buffer.append(colored)
            if len(log_buffer) > 8:
                log_buffer.pop(0)

            # Batasi frekuensi edit pesan agar tidak terkena Discord Rate Limit (minimal 1.2 detik sekali)
            if (datetime.datetime.now() - last_update_time).total_seconds() > 1.2 or percent >= 100:
                last_update_time = datetime.datetime.now()
                log_text = "\n".join(log_buffer)
                progress_embed.description = f"Sedang memproses owner **{self.selected_owner}**...\n\n{create_progress_bar(percent)}"
                progress_embed.set_field_at(
                    0,
                    name="💻 Terminal Logs",
                    value=f"```ansi\n{log_text}\n```",
                    inline=False
                )
                try:
                    await interaction.edit_original_response(embed=progress_embed, view=None)
                except Exception:
                    pass

        # Jalankan di ThreadPoolExecutor agar tidak memblokir event loop asyncio bot
        loop = asyncio.get_running_loop()
        
        def run_task():
            def cb(pct, msg):
                asyncio.run_coroutine_threadsafe(discord_progress_callback(pct, msg), loop)
            return generate_for_owner_pipeline(
                owner_name=self.selected_owner,
                aplikator=self.selected_aplikator,
                upload=True,
                source="vercel",
                live_scrape=True,
                progress_callback=cb
            )

        try:
            result = await loop.run_in_executor(None, run_task)
        except Exception as e:
            result = {"success": False, "error": str(e)}

        # Selesai: Tampilkan Embed Sukses / Gagal
        if result.get("success"):
            folder_url = result.get("folder_url", ROOT_DRIVE_URL)
            file_url = result.get("file_url", "")
            
            success_embed = discord.Embed(
                title="🎉 GENERASI & UPLOAD BERHASIL!",
                description=(
                    f"Data merchant untuk **{result['owner']}** berhasil digabungkan ke dalam **2 Tab Identik** "
                    f"dan telah diunggah ke folder Google Drive."
                ),
                color=THEME_SUCCESS,
                timestamp=datetime.datetime.now()
            )
            success_embed.set_thumbnail(url=SUCCESS_THUMBNAIL)
            
            success_embed.add_field(name="👤 Nama Pemilik", value=f"**{result['owner']}**", inline=True)
            success_embed.add_field(name="📊 Total Outlet", value=f"**{result['total']} Outlet**", inline=True)
            success_embed.add_field(name="📄 File Excel", value=f"`{result['filename']}`\n`✓ 2 Tab (Listing & Listing 2)`", inline=False)
            success_embed.add_field(name="📁 Link Folder Google Drive", value=f"[👉 Buka Folder `{result['owner']}` di Google Drive]({folder_url})", inline=False)
            
            success_embed.set_footer(text="Superfood Tech • Multi-Platform Engine", icon_url=DRIVE_ICON_URL)

            # Buat Link Buttons
            class ResultView(discord.ui.View):
                def __init__(self, f_url, doc_url):
                    super().__init__(timeout=None)
                    self.add_item(discord.ui.Button(label="Buka Folder Google Drive", url=f_url, style=discord.ButtonStyle.link, emoji="📁"))
                    if doc_url:
                        self.add_item(discord.ui.Button(label="Buka File Excel", url=doc_url, style=discord.ButtonStyle.link, emoji="📄"))
                    self.add_item(discord.ui.Button(label="Folder Induk Drive", url=ROOT_DRIVE_URL, style=discord.ButtonStyle.link, emoji="🌐"))

            res_view = ResultView(folder_url, file_url)
            await interaction.edit_original_response(embed=success_embed, view=res_view)

            # Kirim pesan notifikasi baru ke channel (dengan mention user) agar Discord memicu notifikasi suara/unread badge
            try:
                msg_content = (
                    f"🔔 <@{interaction.user.id}> **File Berhasil Dibuat & Diunggah ke Google Drive!**\n"
                    f"• **Owner:** {result['owner']}\n"
                    f"• **File:** `{result['filename']}`\n"
                    f"• **Total Outlet:** {result['total']}\n"
                    f"• **Folder Drive:** {folder_url}"
                )
                await interaction.followup.send(content=msg_content)
            except Exception as e:
                print(f"⚠️ Gagal mengirim pesan notifikasi lanjutan: {e}")

        else:
            error_embed = discord.Embed(
                title="❌ GAGAL MEMPROSES GENERASI",
                description=f"Terjadi kesalahan saat memproses data untuk **{self.selected_owner}**:\n```\n{result.get('error', 'Unknown Error')}\n```",
                color=THEME_ERROR,
                timestamp=datetime.datetime.now()
            )
            error_embed.set_footer(text="Periksa logs terminal atau hubungi administrator.")
            
            # Enable buttons again
            for item in self.children:
                item.disabled = False
            await interaction.edit_original_response(embed=error_embed, view=self)

            # Kirim notifikasi error ke user
            try:
                await interaction.followup.send(
                    content=f"❌ <@{interaction.user.id}> **Gagal memproses data untuk {self.selected_owner}:** {result.get('error', 'Unknown Error')}"
                )
            except Exception:
                pass


# ─── Bot Setup & Commands ─────────────────────────────────────────────────────

class OutletInfoBot(commands.Bot):
    def __init__(self):
        intents = discord.Intents.default()
        super().__init__(command_prefix="!", intents=intents)

    async def setup_hook(self):
        # Daftarkan slash command
        print("[*] Melakukan sinkronisasi Slash Commands ke Discord...")
        await self.tree.sync()
        print("[✓] Slash Commands berhasil disinkronkan!")

    async def on_ready(self):
        print("=" * 60)
        print(f"🤖 Bot Login Berhasil sebagai: {self.user.name} ({self.user.id})")
        print(f"🌟 Status: Online & Siap Digunakan!")
        print("   Gunakan perintah Slash: /generate di server Discord Anda.")
        print("=" * 60)
        await self.change_presence(
            activity=discord.Activity(
                type=discord.ActivityType.watching,
                name="Outlet Info • /generate"
            )
        )

        # Jika DISCORD_CHANNEL_ID disetel di .env, kirim Control Panel otomatis ke channel tersebut
        if CHANNEL_ID and CHANNEL_ID.isdigit():
            try:
                target_channel = self.get_channel(int(CHANNEL_ID))
                if target_channel:
                    view = ControlPanelView(self.user)
                    embed = view.build_embed()
                    embed.title = "🟢 BOT ONLINE • CONTROL PANEL"
                    await target_channel.send(embed=embed, view=view)
                    print(f"[✓] Control Panel otomatis dikirim ke channel #{target_channel.name} ({CHANNEL_ID})")
                else:
                    print(f"⚠️ Channel dengan ID {CHANNEL_ID} tidak ditemukan atau bot tidak punya akses.")
            except Exception as e:
                print(f"⚠️ Gagal mengirim pesan otomatis ke channel ID {CHANNEL_ID}: {e}")


bot = OutletInfoBot()


@bot.tree.command(name="generate", description="🚀 Buka Control Panel untuk Generate & Upload Outlet Info ke Google Drive")
async def generate_slash(interaction: discord.Interaction):
    """Menampilkan Control Panel Interaktif."""
    view = ControlPanelView(interaction.user)
    embed = view.build_embed()
    await interaction.response.send_message(embed=embed, view=view)


@bot.tree.command(name="status", description="📊 Cek status data master GoFood, Grab & Shopee saat ini")
async def status_slash(interaction: discord.Interaction):
    """Menampilkan status ringkas data outlet saat ini."""
    owners = get_owners_with_metadata()
    total_owners = len(owners)
    total_outlets = sum(o["total"] for o in owners)
    total_go = sum(o["gofood"] for o in owners)
    total_grab = sum(o["grab"] for o in owners)
    total_shopee = sum(o.get("shopee", 0) for o in owners)
    
    newest = owners[0]["owner"] if owners else "-"
    newest_ts = owners[0]["latest_ts"] if owners else "-"

    embed = discord.Embed(
        title="📊 STATUS DATA OUTLET SAAT INI",
        color=THEME_ACCENT,
        timestamp=datetime.datetime.now()
    )
    embed.add_field(name="👥 Total Owner", value=f"**{total_owners}** Owner", inline=True)
    embed.add_field(name="🏪 Total Outlet", value=f"**{total_outlets}** Toko", inline=True)
    embed.add_field(
        name="🌐 Distribusi",
        value=f"🔴 GoFood: **{total_go}**\n🟢 Grab: **{total_grab}**\n🟠 Shopee: **{total_shopee}**",
        inline=True
    )
    embed.add_field(name="🔥 Owner Terbaru", value=f"**{newest}** (`{newest_ts}`)", inline=False)
    embed.add_field(name="📁 Root Google Drive", value=f"[Buka Google Drive]({ROOT_DRIVE_URL})", inline=False)
    
    embed.set_footer(text="Gunakan /generate untuk mengekspor data.")
    await interaction.response.send_message(embed=embed)


def main():
    placeholder_tokens = ["masukkan_token_bot_anda_disini", "masukkan_token_anda_disini", "your_token_here", "token_bot_anda"]
    if not BOT_TOKEN or BOT_TOKEN.strip() in placeholder_tokens:
        print("\n" + "=" * 65)
        print("❌ DISCORD_BOT_TOKEN belum diisi dengan token asli di file .env!")
        print("=" * 65)
        print("Langkah mudah untuk mendapatkan Token Bot:")
        print("1. Buka https://discord.com/developers/applications")
        print("2. Pilih aplikasi Bot Anda (atau buat baru via 'New Application')")
        print("3. Buka menu 'Bot' di sebelah kiri, klik 'Reset Token' lalu 'Copy'")
        print("4. Buka menu 'OAuth2' -> 'URL Generator':")
        print("   - Pilih Scopes: 'bot' dan 'applications.commands'")
        print("   - Pilih Bot Permissions: 'Send Messages', 'Embed Links', 'Attach Files'")
        print("   - Salin link URL dan buka di browser untuk invite bot ke server.")
        print("5. Buka file .env di project ini dan masukkan token asli:")
        print("   DISCORD_BOT_TOKEN=token_asli_anda_disini")
        print("=" * 65 + "\n")
        sys.exit(1)

    try:
        bot.run(BOT_TOKEN.strip())
    except discord.errors.LoginFailure:
        print("\n" + "=" * 65)
        print("❌ GAGAL LOGIN KE DISCORD: Token Bot Tidak Valid!")
        print("=" * 65)
        print("Token yang ada di file .env tidak dikenali oleh Discord.")
        print("Silakan generate ulang token di Discord Developer Portal -> Bot -> Reset Token.")
        print("=" * 65 + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()
