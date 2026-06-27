#!/usr/bin/env python3
"""
CLI utama untuk Outlet Info Scraper
Mengelola scraping data merchant dari Grab dan Shopee Food.
"""

import os
import sys
import subprocess
import asyncio

# ─── ANSI Colors ──────────────────────────────────────────────────────────────
RESET   = "\033[0m"
BOLD    = "\033[1m"
DIM     = "\033[2m"

BLACK   = "\033[30m"
RED     = "\033[31m"
GREEN   = "\033[32m"
YELLOW  = "\033[33m"
BLUE    = "\033[34m"
MAGENTA = "\033[35m"
CYAN    = "\033[36m"
WHITE   = "\033[37m"

BG_BLACK  = "\033[40m"
BG_BLUE   = "\033[44m"
BG_CYAN   = "\033[46m"

def c(text, *styles):
    return "".join(styles) + str(text) + RESET

def clear():
    os.system("clear" if os.name == "posix" else "cls")

# ─── Paths ────────────────────────────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
GRAB_DIR   = os.path.join(BASE_DIR, "GRAB")
SHOPEE_DIR = os.path.join(BASE_DIR, "SHOPEE")

GRAB_CREDENTIALS = [
    {"name": "F1",   "username": "automationf1"},
    {"name": "F2S",  "username": "automationf2s"},
    {"name": "W1",   "username": "automationw1"},
    {"name": "L1",   "username": "automationl1"},
    {"name": "L2",   "username": "automationl2"},
    {"name": "DE1S", "username": "automationde1s"},
    {"name": "JF1",  "username": "automationjf1"},
    {"name": "JF1S", "username": "automationjf1s"},
]

SHOPEE_CREDENTIALS = [
    {"name": "F",  "username": "superfoodapp"},
    {"name": "W",  "username": "wonderfoodapp"},
    {"name": "L",  "username": "lokarasaapp"},
    {"name": "D",  "username": "doeatapp"},
]

# ─── UI Helpers ───────────────────────────────────────────────────────────────

def header(title="Outlet Info CLI"):
    clear()
    width = 54
    print()
    print(c("╔" + "═" * width + "╗", CYAN, BOLD))
    print(c("║" + f"  🛒  {title}".center(width) + "║", CYAN, BOLD))
    print(c("╚" + "═" * width + "╝", CYAN, BOLD))
    print()

def section(title):
    print()
    print(c(f"  ── {title} ", YELLOW, BOLD) + c("─" * (40 - len(title)), DIM))

def menu_item(num, icon, label, desc=""):
    num_str = c(f"  [{num}]", CYAN, BOLD)
    icon_str = f" {icon} "
    label_str = c(label, WHITE, BOLD)
    desc_str = c(f"  {desc}", DIM) if desc else ""
    print(f"{num_str}{icon_str}{label_str}{desc_str}")

def divider():
    print(c("  " + "─" * 50, DIM))

def success(msg):
    print(c(f"\n  ✓ {msg}", GREEN, BOLD))

def error(msg):
    print(c(f"\n  ✗ {msg}", RED, BOLD))

def info(msg):
    print(c(f"\n  ℹ {msg}", CYAN))

def warning(msg):
    print(c(f"\n  ⚠ {msg}", YELLOW))

def prompt(msg):
    return input(c(f"\n  → {msg}: ", CYAN, BOLD)).strip()

def wait():
    input(c("\n  Tekan Enter untuk kembali...", DIM))

def run_script(script_path, cwd=None, extra_args=None):
    """Jalankan script Python dan tampilkan output-nya."""
    cmd = [sys.executable, script_path]
    if extra_args:
        cmd.extend(extra_args)
    print(c(f"\n  Menjalankan: {os.path.basename(script_path)}\n", DIM))
    print(c("  " + "─" * 50, DIM))
    try:
        proc = subprocess.run(cmd, cwd=cwd or os.path.dirname(script_path))
        if proc.returncode == 0:
            success("Script selesai.")
        else:
            warning(f"Script selesai dengan exit code {proc.returncode}.")
    except FileNotFoundError:
        error(f"File tidak ditemukan: {script_path}")
    except Exception as e:
        error(f"Error: {e}")

# ─── GRAB Menus ───────────────────────────────────────────────────────────────

def grab_select_outlets():
    """Pilih outlet Grab yang ingin di-scrape."""
    header("Grab — Pilih Outlet")
    section("Daftar Outlet")

    menu_item("0", "🔄", "Semua Outlet", "Jalankan scraping untuk semua outlet")
    divider()
    for i, cred in enumerate(GRAB_CREDENTIALS, 1):
        menu_item(str(i), "🏪", cred["name"], f"@{cred['username']}")

    print()
    menu_item("b", "↩", "Kembali")
    print()

    choice = prompt("Pilih outlet")

    if choice.lower() == "b":
        return None
    elif choice == "0":
        return "all"
    elif choice.isdigit() and 1 <= int(choice) <= len(GRAB_CREDENTIALS):
        return GRAB_CREDENTIALS[int(choice) - 1]
    else:
        error("Pilihan tidak valid.")
        wait()
        return grab_select_outlets()


def grab_run_scraper():
    """Jalankan Grab scraper."""
    selected = grab_select_outlets()
    if selected is None:
        return

    header("Grab — Menjalankan Scraper")

    if selected == "all":
        info("Menjalankan scraper untuk SEMUA outlet Grab...")
        run_script(
            os.path.join(GRAB_DIR, "grab_merchant_scraper.py"),
            cwd=GRAB_DIR
        )
    else:
        info(f"Menjalankan scraper untuk outlet: {selected['name']}")
        # Jalankan dengan argument outlet tertentu
        run_script(
            os.path.join(GRAB_DIR, "grab_merchant_scraper.py"),
            cwd=GRAB_DIR,
            extra_args=["--outlet", selected["name"]]
        )

    wait()


def grab_combine():
    header("Grab — Combine Excel")
    info("Menggabungkan semua hasil scraping menjadi MASTER_ALL.xlsx...")
    run_script(os.path.join(GRAB_DIR, "combine_custom.py"), cwd=GRAB_DIR)
    wait()


def grab_find_duplicates():
    header("Grab — Cari Duplikat")
    run_script(os.path.join(GRAB_DIR, "find_duplicates.py"), cwd=GRAB_DIR)
    wait()


def grab_remove_duplicates():
    header("Grab — Hapus Duplikat")
    warning("Ini akan MENGUBAH file Excel di hasil_custom/ secara permanen!")
    confirm = prompt("Lanjutkan? (y/N)")
    if confirm.lower() == "y":
        run_script(os.path.join(GRAB_DIR, "remove_dup_custom.py"), cwd=GRAB_DIR)
    else:
        info("Dibatalkan.")
    wait()


def grab_check_output():
    """Tampilkan ringkasan file output Grab."""
    header("Grab — Status Output")
    output_dir = os.path.join(GRAB_DIR, "hasil_custom")

    section("File di hasil_custom/")
    if not os.path.exists(output_dir):
        warning("Folder hasil_custom/ belum ada. Jalankan scraper terlebih dahulu.")
        wait()
        return

    import glob
    files = sorted(glob.glob(os.path.join(output_dir, "*.xlsx")))
    if not files:
        warning("Belum ada file hasil. Jalankan scraper terlebih dahulu.")
    else:
        try:
            import pandas as pd
            for f in files:
                name = os.path.basename(f)
                size = os.path.getsize(f)
                try:
                    df = pd.read_excel(f)
                    rows = len(df)
                    print(c(f"  ✓ {name:<20}", GREEN) + c(f"  {rows:>5} baris", CYAN) + c(f"  ({size//1024} KB)", DIM))
                except Exception:
                    print(c(f"  ? {name:<20}", YELLOW) + c(f"  ({size//1024} KB)", DIM))
        except ImportError:
            for f in files:
                name = os.path.basename(f)
                size = os.path.getsize(f)
                print(c(f"  • {name}", WHITE) + c(f"  ({size//1024} KB)", DIM))

    wait()


def menu_grab():
    while True:
        header("Grab Merchant Scraper")
        section("Menu Grab")
        menu_item("1", "▶", "Jalankan Scraper",       "Ambil data merchant dari portal Grab")
        menu_item("2", "📊", "Combine Excel",           "Gabungkan semua file hasil ke MASTER_ALL.xlsx")
        menu_item("3", "🔍", "Cari Duplikat",           "Temukan baris duplikat di MASTER_ALL.xlsx")
        menu_item("4", "🗑", "Hapus Duplikat",          "Bersihkan duplikat dari file individual")
        menu_item("5", "📁", "Status Output",           "Lihat ringkasan file hasil scraping")
        divider()
        menu_item("b", "↩", "Kembali ke Menu Utama")
        print()

        choice = prompt("Pilih menu")

        if choice == "1":
            grab_run_scraper()
        elif choice == "2":
            grab_combine()
        elif choice == "3":
            grab_find_duplicates()
        elif choice == "4":
            grab_remove_duplicates()
        elif choice == "5":
            grab_check_output()
        elif choice.lower() == "b":
            break
        else:
            error("Pilihan tidak valid.")
            wait()


# ─── SHOPEE Menus ─────────────────────────────────────────────────────────────

def shopee_select_outlets():
    """Pilih outlet Shopee yang ingin di-scrape."""
    header("Shopee — Pilih Outlet")
    section("Daftar Outlet")

    menu_item("0", "🔄", "Semua Outlet", "Jalankan scraping untuk semua outlet")
    divider()
    for i, cred in enumerate(SHOPEE_CREDENTIALS, 1):
        menu_item(str(i), "🏪", cred["name"], f"@{cred['username']}")

    print()
    menu_item("b", "↩", "Kembali")
    print()

    choice = prompt("Pilih outlet")

    if choice.lower() == "b":
        return None
    elif choice == "0":
        return "all"
    elif choice.isdigit() and 1 <= int(choice) <= len(SHOPEE_CREDENTIALS):
        return SHOPEE_CREDENTIALS[int(choice) - 1]
    else:
        error("Pilihan tidak valid.")
        wait()
        return shopee_select_outlets()


def shopee_run_scraper():
    selected = shopee_select_outlets()
    if selected is None:
        return

    header("Shopee — Menjalankan Scraper")

    if selected == "all":
        info("Menjalankan scraper untuk SEMUA outlet Shopee...")
        run_script(os.path.join(SHOPEE_DIR, "shopee_scraper.py"), cwd=SHOPEE_DIR)
    else:
        info(f"Menjalankan scraper untuk outlet: {selected['name']}")
        run_script(
            os.path.join(SHOPEE_DIR, "shopee_scraper.py"),
            cwd=SHOPEE_DIR,
            extra_args=["--outlet", selected["name"]]
        )

    wait()


def shopee_check_output():
    header("Shopee — Status Output")
    output_dir = os.path.join(SHOPEE_DIR, "data")
    section("File di data/")

    if not os.path.exists(output_dir):
        warning("Folder data/ belum ada. Jalankan scraper terlebih dahulu.")
        wait()
        return

    import glob
    files = sorted(glob.glob(os.path.join(output_dir, "*.xlsx")))
    if not files:
        warning("Belum ada file hasil.")
    else:
        try:
            import pandas as pd
            for f in files:
                name = os.path.basename(f)
                size = os.path.getsize(f)
                try:
                    df = pd.read_excel(f)
                    rows = len(df)
                    print(c(f"  ✓ {name:<20}", GREEN) + c(f"  {rows:>5} baris", CYAN) + c(f"  ({size//1024} KB)", DIM))
                except Exception:
                    print(c(f"  ? {name:<20}", YELLOW) + c(f"  ({size//1024} KB)", DIM))
        except ImportError:
            for f in files:
                name = os.path.basename(f)
                print(c(f"  • {name}", WHITE))

    wait()


def menu_shopee():
    while True:
        header("Shopee Food Scraper")
        section("Menu Shopee")
        menu_item("1", "▶", "Jalankan Scraper",   "Ambil data merchant dari Shopee Food")
        menu_item("2", "📁", "Status Output",       "Lihat ringkasan file hasil scraping")
        divider()
        menu_item("b", "↩", "Kembali ke Menu Utama")
        print()

        choice = prompt("Pilih menu")

        if choice == "1":
            shopee_run_scraper()
        elif choice == "2":
            shopee_check_output()
        elif choice.lower() == "b":
            break
        else:
            error("Pilihan tidak valid.")
            wait()


# ─── Main Menu ────────────────────────────────────────────────────────────────

def run_all():
    """Jalankan semua scraper sekaligus."""
    header("Jalankan Semua Scraper")
    warning("Ini akan menjalankan scraper Grab DAN Shopee secara berurutan.")
    confirm = prompt("Lanjutkan? (y/N)")
    if confirm.lower() != "y":
        info("Dibatalkan.")
        wait()
        return

    section("Grab Scraper")
    run_script(os.path.join(GRAB_DIR, "grab_merchant_scraper.py"), cwd=GRAB_DIR)

    section("Shopee Scraper")
    run_script(os.path.join(SHOPEE_DIR, "shopee_scraper.py"), cwd=SHOPEE_DIR)

    section("Combine Grab Excel")
    run_script(os.path.join(GRAB_DIR, "combine_custom.py"), cwd=GRAB_DIR)

    success("Semua scraper selesai dijalankan!")
    wait()


def main():
    while True:
        header()

        print(c("  Platform\n", DIM))
        menu_item("1", "🟢", "Grab",    "Scraper & tools untuk Grab Merchant")
        menu_item("2", "🟠", "Shopee",  "Scraper untuk Shopee Food Partner")
        divider()
        menu_item("3", "🚀", "Jalankan Semua", "Grab + Shopee sekaligus")
        divider()
        menu_item("q", "✖", "Keluar")
        print()

        choice = prompt("Pilih menu")

        if choice == "1":
            menu_grab()
        elif choice == "2":
            menu_shopee()
        elif choice == "3":
            run_all()
        elif choice.lower() in ("q", "exit", "quit", "0"):
            clear()
            print(c("\n  Sampai jumpa! 👋\n", CYAN, BOLD))
            break
        else:
            error("Pilihan tidak valid.")
            wait()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        clear()
        print(c("\n\n  Program dihentikan. Sampai jumpa! 👋\n", CYAN, BOLD))
        sys.exit(0)
