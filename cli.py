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
GOFOOD_DIR = os.path.join(BASE_DIR, "GOFOOD")

# Kredensial Grab sekarang diambil secara dinamis via get_credentials_from_sheet

# ─── UI Helpers ───────────────────────────────────────────────────────────────

def header(title="Outlet Info CLI"):
    clear()
    width = 54
    print()
    print(c("╔" + "═" * width + "╗", CYAN, BOLD))
    print(c("║" + f"  🛒  {title}".center(width) + "║", CYAN, BOLD))
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
    import sys
    if GRAB_DIR not in sys.path:
        sys.path.append(GRAB_DIR)
        
    header("Grab — Pilih Sumber Kredensial")
    menu_item("1", "🏢", "Agency", "Kredensial Agency (Live Google Sheet)")
    menu_item("2", "🏷️", "VB (Virtual Brand)", "Kredensial Virtual Brand / Dokumen Lain")
    menu_item("3", "☁️", "Vercel Sheet", "Kredensial Live CSV Vercel Sheet")
    divider()
    menu_item("b", "↩", "Kembali")
    print()
    src_choice = prompt("Pilih Sumber Kredensial [1/2/3]").lower()
    if src_choice == "b":
        return None, "agency"
    
    if src_choice == "2":
        source_type = "vb"
    elif src_choice == "3":
        source_type = "vercel"
    else:
        source_type = "agency"

    try:
        from grab_merchant_scraper import get_credentials_from_sheet
        credentials = get_credentials_from_sheet(source_type=source_type)
    except Exception as e:
        error(f"Gagal mengambil kredensial Grab ({source_type}): {e}")
        credentials = []

    header(f"Grab — Pilih Outlet [{source_type.upper()}]")
    section("Daftar Outlet")

    menu_item("0", "🔄", "Semua Outlet", "Jalankan scraping untuk semua outlet")
    menu_item("r", "▶", "Lanjutkan (Resume)", "Lewati yang sudah selesai ditarik")
    divider()
    for i, cred in enumerate(credentials, 1):
        menu_item(str(i), "🏪", f"{cred['name']} ({cred.get('owner', '-')})", f"@{cred['username']}")

    print()
    menu_item("b", "↩", "Kembali")
    print()

    choice = prompt("Pilih outlet")

    if choice.lower() == "b":
        return None, source_type
    elif choice == "0":
        return "all", source_type
    elif choice.lower() == "r":
        return "resume", source_type
    elif choice.isdigit() and 1 <= int(choice) <= len(credentials):
        return credentials[int(choice) - 1], source_type
    else:
        error("Pilihan tidak valid.")
        wait()
        return grab_select_outlets()

def grab_run_scraper():
    """Jalankan Grab scraper."""
    selected, source_type = grab_select_outlets()
    if selected is None:
        return

    if source_type == "vb":
        type_arg = ["--vb"]
    elif source_type == "vercel":
        type_arg = ["--vercel"]
    else:
        type_arg = ["--agency"]

    if selected == "all":
        info(f"Menjalankan scraper untuk SEMUA outlet Grab [{source_type.upper()}]...")
        run_script(
            os.path.join(GRAB_DIR, "grab_merchant_scraper.py"),
            cwd=GRAB_DIR,
            extra_args=["--all"] + type_arg
        )
    elif selected == "resume":
        info(f"Melanjutkan penarikan untuk portal [{source_type.upper()}] yang belum selesai...")
        run_script(
            os.path.join(GRAB_DIR, "grab_merchant_scraper.py"),
            cwd=GRAB_DIR,
            extra_args=["--all"] + type_arg
        )
    else:
        info(f"Menjalankan scraper untuk outlet: {selected['name']}")
        run_script(
            os.path.join(GRAB_DIR, "grab_merchant_scraper.py"),
            cwd=GRAB_DIR,
            extra_args=["--outlet", selected["name"]] + type_arg
        )

    wait()


def grab_combine():
    header("Grab — Combine Excel")
    info("Menggabungkan semua data cache JSON Grab ke Master & Output...")
    run_script(os.path.join(GRAB_DIR, "grab_merchant_scraper.py"), cwd=GRAB_DIR, extra_args=["--combine"])
    wait()


def grab_check_output():
    """Tampilkan ringkasan file output Grab."""
    header("Grab — Status Output")
    master_dir = os.path.join(GRAB_DIR, "master")
    output_dir = os.path.join(GRAB_DIR, "output")
    cache_dir = os.path.join(GRAB_DIR, "cache")
    import glob
    import pandas as pd

    # 1. Master Files
    section("File Master di master/")
    if os.path.exists(master_dir):
        m_files = sorted(glob.glob(os.path.join(master_dir, "*.xlsx")))
        if not m_files:
            print(c("  (Belum ada file master)", DIM))
        for f in m_files:
            name = os.path.basename(f)
            size = os.path.getsize(f)
            try:
                df = pd.read_excel(f)
                print(c(f"  ✓ {name:<35}", GREEN) + c(f"  {len(df):>5} baris", CYAN) + c(f"  ({size//1024} KB)", DIM))
            except Exception:
                print(c(f"  ? {name:<35}", YELLOW) + c(f"  ({size//1024} KB)", DIM))
    else:
        print(c("  (Folder master/ belum ada)", DIM))

    # 2. Output Per-Owner
    section("File Per-Owner di output/")
    if os.path.exists(output_dir):
        o_files = sorted(glob.glob(os.path.join(output_dir, "*.xlsx")))
        if not o_files:
            print(c("  (Belum ada file output per-owner)", DIM))
        for f in o_files[-15:]:  # Tampilkan 15 file terbaru
            name = os.path.basename(f)
            size = os.path.getsize(f)
            try:
                df = pd.read_excel(f)
                print(c(f"  ✓ {name:<35}", GREEN) + c(f"  {len(df):>5} baris", CYAN) + c(f"  ({size//1024} KB)", DIM))
            except Exception:
                print(c(f"  ? {name:<35}", YELLOW) + c(f"  ({size//1024} KB)", DIM))
        if len(o_files) > 15:
            print(c(f"  ... dan {len(o_files) - 15} file lainnya.", DIM))
    else:
        print(c("  (Folder output/ belum ada)", DIM))

    # 3. Cache JSON
    section("Cache JSON di cache/")
    if os.path.exists(cache_dir):
        c_files = glob.glob(os.path.join(cache_dir, "grab_portal_*.json"))
        print(c(f"  📦 Total file cache portal tersimpan: {len(c_files)} file JSON", CYAN))
    else:
        print(c("  (Folder cache/ belum ada)", DIM))

    wait()


def menu_grab():
    while True:
        header("Grab Merchant Scraper")
        section("Menu Grab")
        menu_item("1", "▶", "Jalankan Scraper",       "Ambil data merchant dari portal Grab")
        menu_item("2", "📊", "Combine Excel",           "Gabungkan cache JSON ke master & output")
        menu_item("3", "📁", "Status Output",           "Lihat ringkasan file hasil scraping")
        divider()
        menu_item("b", "↩", "Kembali ke Menu Utama")
        print()

        choice = prompt("Pilih menu")

        if choice == "1":
            grab_run_scraper()
        elif choice == "2":
            grab_combine()
        elif choice == "3":
            grab_check_output()
        elif choice.lower() == "b":
            break
        else:
            error("Pilihan tidak valid.")
            wait()


# ─── SHOPEE Menus ─────────────────────────────────────────────────────────────

def shopee_select_outlets():
    """Pilih outlet Shopee yang ingin di-scrape (Via VB Pipeline)."""
    # Ambil kredensial secara dinamis dari VB lokal
    import sys
    if SHOPEE_DIR not in sys.path:
        sys.path.append(SHOPEE_DIR)
        
    try:
        from init_sessions import get_vb_portals
        credentials = get_vb_portals()
    except Exception as e:
        error(f"Gagal mengambil kredensial dari VB: {e}")
        credentials = []
    
    header("Shopee — Pilih Outlet")
    section("Daftar Outlet")

    menu_item("0", "🔄", "Semua Outlet", "Jalankan scraping untuk semua outlet")
    divider()
    for i, cred in enumerate(credentials, 1):
        display_name = cred.get("merchant_name", cred.get("account_name", ""))
        user_info = cred.get('username') or cred.get('phone', '')
        menu_item(str(i), "🏪", display_name, f"@{user_info}")

    print()
    menu_item("b", "↩", "Kembali")
    print()

    choice = prompt("Pilih outlet")

    if choice.lower() == "b":
        return None
    elif choice == "0":
        return "all"
    
    elif choice.isdigit() and 1 <= int(choice) <= len(credentials):
        return credentials[int(choice) - 1]
    else:
        error("Pilihan tidak valid.")
        wait()
        return shopee_select_outlets()


def shopee_run_interactive():
    """Jalankan CLI interaktif Shopee untuk filter merchant."""
    header("Shopee — CLI Interaktif")
    script = os.path.join(SHOPEE_DIR, "cli.py")
    run_script(script, cwd=SHOPEE_DIR)
    wait()


def shopee_run_all():
    """Jalankan penarikan data semua merchant Shopee."""
    header("Shopee — Tarik Semua Outlet")
    script = os.path.join(SHOPEE_DIR, "pull_outlet_info.py")
    run_script(script, cwd=SHOPEE_DIR)
    wait()


def shopee_open_dashboard():
    """Buka dashboard Shopee dalam mode idle."""
    header("Shopee — Buka Dashboard")
    script = os.path.join(SHOPEE_DIR, "open_dashboard.py")
    run_script(script, cwd=SHOPEE_DIR)
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
                    print(c(f"  ✓ {name:<25}", GREEN) + c(f"  {rows:>5} baris", CYAN) + c(f"  ({size//1024} KB)", DIM))
                except Exception:
                    print(c(f"  ? {name:<25}", YELLOW) + c(f"  ({size//1024} KB)", DIM))
        except ImportError:
            for f in files:
                name = os.path.basename(f)
                print(c(f"  • {name}", WHITE))

    wait()


def menu_shopee():
    while True:
        header("Shopee Food Scraper")
        section("Menu Shopee")
        menu_item("1", "▶", "CLI Interaktif / Filter", "Pilih merchant by ID / Nama / Nomor Index")
        menu_item("2", "🚀", "Tarik Semua Outlet",       "Jalankan auto-pull & switch semua merchant")
        menu_item("3", "🌐", "Buka Dashboard (Idle)",    "Buka browser Shopee Partner tanpa timeout")
        menu_item("4", "📁", "Status Output",            "Lihat ringkasan file hasil scraping")
        divider()
        menu_item("b", "↩", "Kembali ke Menu Utama")
        print()

        choice = prompt("Pilih menu")

        if choice == "1":
            shopee_run_interactive()
        elif choice == "2":
            shopee_run_all()
        elif choice == "3":
            shopee_open_dashboard()
        elif choice == "4":
            shopee_check_output()
        elif choice.lower() == "b":
            break
        else:
            error("Pilihan tidak valid.")
            wait()


# ─── GOFOOD Menus ─────────────────────────────────────────────────────────────

def gofood_run_scraper():
    header("GoFood — Menjalankan Scraper")
    info("Menjalankan scraper untuk GoFood...")
    run_script(os.path.join(GOFOOD_DIR, "gofood_scraper.py"), cwd=GOFOOD_DIR)
    wait()


def gofood_generate_master():
    header("GoFood — Combine Excel")
    info("Menggabungkan semua data cache JSON GoFood ke Master & Output...")
    run_script(os.path.join(GOFOOD_DIR, "gofood_scraper.py"), cwd=GOFOOD_DIR, extra_args=["--combine"])
    wait()


def gofood_check_output():
    header("GoFood — Status Output")
    master_dir = os.path.join(GOFOOD_DIR, "master")
    output_dir = os.path.join(GOFOOD_DIR, "output")
    cache_dir = os.path.join(GOFOOD_DIR, "cache")
    import glob
    import pandas as pd

    # 1. Master Files
    section("File Master di master/")
    if os.path.exists(master_dir):
        m_files = sorted(glob.glob(os.path.join(master_dir, "*.xlsx")))
        if not m_files:
            print(c("  (Belum ada file master)", DIM))
        for f in m_files:
            name = os.path.basename(f)
            size = os.path.getsize(f)
            try:
                df = pd.read_excel(f)
                print(c(f"  ✓ {name:<35}", GREEN) + c(f"  {len(df):>5} baris", CYAN) + c(f"  ({size//1024} KB)", DIM))
            except Exception:
                print(c(f"  ? {name:<35}", YELLOW) + c(f"  ({size//1024} KB)", DIM))
    else:
        print(c("  (Folder master/ belum ada)", DIM))

    # 2. Output Per-Owner
    section("File Per-Owner di output/")
    if os.path.exists(output_dir):
        o_files = sorted(glob.glob(os.path.join(output_dir, "*.xlsx")))
        if not o_files:
            print(c("  (Belum ada file output per-owner)", DIM))
        for f in o_files[-15:]:
            name = os.path.basename(f)
            size = os.path.getsize(f)
            try:
                df = pd.read_excel(f)
                print(c(f"  ✓ {name:<35}", GREEN) + c(f"  {len(df):>5} baris", CYAN) + c(f"  ({size//1024} KB)", DIM))
            except Exception:
                print(c(f"  ? {name:<35}", YELLOW) + c(f"  ({size//1024} KB)", DIM))
        if len(o_files) > 15:
            print(c(f"  ... dan {len(o_files) - 15} file lainnya.", DIM))
    else:
        print(c("  (Folder output/ belum ada)", DIM))

    # 3. Cache JSON
    section("Cache JSON di cache/")
    if os.path.exists(cache_dir):
        c_files = glob.glob(os.path.join(cache_dir, "gofood_portal_*.json"))
        print(c(f"  📦 Total file cache portal tersimpan: {len(c_files)} file JSON", CYAN))
    else:
        print(c("  (Folder cache/ belum ada)", DIM))

    wait()


def menu_gofood():
    while True:
        header("GoFood Merchant Scraper")
        section("Menu GoFood")
        menu_item("1", "▶", "Jalankan Scraper",       "Ambil data merchant dari portal GoFood")
        menu_item("2", "📊", "Combine Excel",           "Gabungkan cache JSON ke master & output")
        menu_item("3", "📁", "Status Output",           "Lihat ringkasan file hasil scraping")
        divider()
        menu_item("b", "↩", "Kembali ke Menu Utama")
        print()

        choice = prompt("Pilih menu")

        if choice == "1":
            gofood_run_scraper()
        elif choice == "2":
            gofood_generate_master()
        elif choice == "3":
            gofood_check_output()
        elif choice.lower() == "b":
            break
        else:
            error("Pilihan tidak valid.")
            wait()

def run_all():
    """Jalankan semua scraper sekaligus."""
    header("Jalankan Semua Scraper")
    warning("Ini akan menjalankan scraper Grab DAN Shopee secara berurutan.")
    confirm = prompt("Lanjutkan? (y/N)")
    if confirm.lower() != "y":
        info("Dibatalkan.")
        wait()
        return
    run_script(os.path.join(GRAB_DIR, "grab_merchant_scraper.py"), cwd=GRAB_DIR)

    section("Shopee Scraper (Via VB)")
    run_script(os.path.join(SHOPEE_DIR, "run_baseline.py"), cwd=SHOPEE_DIR)

    section("GoFood Scraper")
    run_script(os.path.join(GOFOOD_DIR, "gofood_scraper.py"), cwd=GOFOOD_DIR)

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
        menu_item("3", "🔴", "GoFood",  "Scraper untuk GoFood/GoBiz")
        divider()
        menu_item("4", "🚀", "Jalankan Semua", "Grab + Shopee + GoFood sekaligus")
        divider()
        menu_item("q", "✖", "Keluar")
        print()

        choice = prompt("Pilih menu")

        if choice == "1":
            menu_grab()
        elif choice == "2":
            menu_shopee()
        elif choice == "3":
            menu_gofood()
        elif choice == "4":
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
