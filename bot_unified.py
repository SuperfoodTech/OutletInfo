#!/usr/bin/env python3
"""
bot_unified.py
==============
Entrypoint tunggal untuk menjalankan Bot Listing (/listing) dan Bot SOT (/sot)
dalam satu proses Python asyncio dan satu container Docker.
Menghemat pemakaian RAM dan mencegah OOM pada server VPS 2GB.
"""
import os
import sys
import asyncio
import signal
from pathlib import Path
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

load_dotenv(BASE_DIR / ".env")

# Import instance bot dari masing-masing modul
from discord_bot import bot as bot_listing
from bot_sot import bot_sot

TOKEN_LISTING = (os.getenv("DISCORD_BOT_TOKEN_LISTING") or os.getenv("DISCORD_BOT_TOKEN", "")).strip()
TOKEN_SOT = (os.getenv("DISCORD_BOT_TOKEN_SOT") or "").strip()


async def start_bot(client, token: str, label: str):
    placeholder_tokens = ("masukkan_token_bot_anda_disini", "your_token_here", "token_bot_anda")
    if not token or token in placeholder_tokens:
        print(f"⚠️ [{label}] Token tidak ditemukan di .env. Bot ini tidak dijalankan.")
        return
    try:
        print(f"[*] [{label}] Memulai login ke Discord...")
        await client.start(token)
    except Exception as e:
        print(f"❌ [{label}] Error saat menjalankan bot: {e}")


async def main():
    print("=" * 65)
    print("🚀 MEMULAI UNIFIED BOT RUNNER (LISTING + SOT DALAM 1 PROSES)")
    print("=" * 65)

    tasks = []
    if TOKEN_LISTING:
        tasks.append(asyncio.create_task(start_bot(bot_listing, TOKEN_LISTING, "BOT-LISTING")))
    else:
        print("⚠️ [BOT-LISTING] DISCORD_BOT_TOKEN_LISTING tidak disetel.")

    if TOKEN_SOT:
        tasks.append(asyncio.create_task(start_bot(bot_sot, TOKEN_SOT, "BOT-SOT")))
    else:
        print("⚠️ [BOT-SOT] DISCORD_BOT_TOKEN_SOT tidak disetel.")

    if not tasks:
        print("❌ Tidak ada token bot yang aktif. Periksa file .env Anda!")
        return

    # Graceful shutdown handler
    loop = asyncio.get_running_loop()
    stop_event = asyncio.Event()

    def _signal_handler():
        print("\n🛑 Sinyal shutdown diterima. Menghentikan kedua bot...")
        stop_event.set()

    for sig in (signal.SIGINT, signal.SIGTERM):
        try:
            loop.add_signal_handler(sig, _signal_handler)
        except NotImplementedError:
            pass

    # Jalankan sampai ada sinyal stop atau bot selesai
    await asyncio.wait(
        [asyncio.create_task(stop_event.wait()), *tasks],
        return_when=asyncio.FIRST_COMPLETED
    )

    print("[*] Menutup sesi bot Discord...")
    try:
        if not bot_listing.is_closed():
            await bot_listing.close()
    except Exception as e:
        print(f"⚠️ Gagal menutup bot listing: {e}")

    try:
        if not bot_sot.is_closed():
            await bot_sot.close()
    except Exception as e:
        print(f"⚠️ Gagal menutup bot sot: {e}")

    print("[✓] Sesi kedua bot berhasil ditutup.")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, SystemExit):
        pass
