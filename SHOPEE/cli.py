#!/usr/bin/env python3
"""
CLI Tool for Shopee Outlet Info Puller
======================================
Alat bantu CLI untuk memfilter dan menarik data outlet dari merchant tertentu.

Penggunaan:
  1. Menu Interaktif (pilih merchant lewat nomor):
     python cli.py
     python cli.py -I

  2. Lihat Daftar Merchant:
     python cli.py --list
     python cli.py --list --search ayam

  3. Filter berdasarkan Merchant ID:
     python cli.py --id 828310
     python cli.py --id 828310 1530883

  4. Filter berdasarkan Nama / Kata Kunci:
     python cli.py --search "Ayam"
     python cli.py --name "AYAM GEPREK MAMA CHUBBY_" "DEPOT 88 CHINESE FOOD_"

  5. Filter berdasarkan Nomor Index:
     python cli.py --index 1,3,5-10

  6. Opsi Tambahan:
     --no-resume / --fresh  : Tarik ulang tanpa membaca Excel existing
     --include-excluded     : Sertakan merchant yang masuk daftar blacklist (SuperFood, dll)
     --output / -o          : Simpan output ke file Excel spesifik
     --dry-run              : Lihat merchant yang terpilih tanpa menjalankan browser
"""

import os
import sys
import argparse
from pathlib import Path

# Setup paths
SCRIPT_DIR = Path(__file__).resolve().parent

# Auto-detect and switch to .venv python if not active
for venv_candidate in [SCRIPT_DIR / ".venv" / "bin" / "python", SCRIPT_DIR.parent / ".venv" / "bin" / "python", SCRIPT_DIR.parent.parent / ".venv" / "bin" / "python"]:
    if venv_candidate.exists() and sys.executable != str(venv_candidate):
        os.execv(str(venv_candidate), [str(venv_candidate)] + sys.argv)
        break

if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from pull_outlet_info import (
    get_merchants_to_switch,
    is_excluded,
    run_pull,
    OUTPUT_DIR,
)


def parse_index_range(range_str: str, max_val: int) -> list[int]:
    """Parse string range like '1,3,5-8' into list of 1-based integer indices."""
    indices = set()
    parts = range_str.replace(" ", "").split(",")
    for part in parts:
        if not part:
            continue
        if "-" in part:
            try:
                start_s, end_s = part.split("-", 1)
                start = int(start_s)
                end = int(end_s)
                for i in range(min(start, end), max(start, end) + 1):
                    if 1 <= i <= max_val:
                        indices.add(i)
            except ValueError:
                pass
        else:
            try:
                val = int(part)
                if 1 <= val <= max_val:
                    indices.add(val)
            except ValueError:
                pass
    return sorted(list(indices))


def display_merchant_table(merchants: list[dict], search_query: str = None, show_excluded: bool = True):
    """Print a clean formatted table of merchants."""
    print("=" * 80)
    print(f" {'NO':<4} | {'MERCHANT ID':<13} | {'OCC':<4} | {'STATUS':<9} | {'NAMA MERCHANT'}")
    print("-" * 80)

    count = 0
    for idx, m in enumerate(merchants, 1):
        m_id = str(m.get("merchant_id") or "-")
        m_name = m.get("merchant_name", "")
        occ = str(m.get("occurrence_index", 0))
        excluded = is_excluded(m_name)

        if search_query and search_query.lower() not in m_name.lower() and search_query not in m_id:
            continue

        if not show_excluded and excluded:
            continue

        status_tag = "[EXCLUDE]" if excluded else "Normal"
        print(f" {idx:<4} | {m_id:<13} | {occ:<4} | {status_tag:<9} | {m_name}")
        count += 1

    print("=" * 80)
    print(f" Total ditampilkan: {count} merchant (dari total {len(merchants)} tersedia)\n")


def interactive_select(merchants: list[dict]) -> list[dict]:
    """Interactive CLI menu to select merchants."""
    while True:
        print("\n" + "=" * 80)
        print("  PILIH MERCHANT YANG AKAN DITARIK")
        print("=" * 80)
        display_merchant_table(merchants, show_excluded=True)

        print("Petunjuk Input:")
        print("  - Masukkan nomor index (contoh: 1, 3, 5-8)")
        print("  - Ketik 'all' untuk memilih SEMUA merchant yang tidak di-exclude")
        print("  - Ketik 'search <kata_kunci>' untuk memfilter tampilan")
        print("  - Ketik 'q' untuk keluar")
        print()

        try:
            choice = input("Pilihan Anda: ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\n[!] Dibatalkan.")
            sys.exit(0)

        if not choice:
            continue

        if choice.lower() in ("q", "quit", "exit"):
            print("[*] Keluar dari CLI.")
            sys.exit(0)

        if choice.lower() == "all":
            selected = [m for m in merchants if not is_excluded(m.get("merchant_name", ""))]
            print(f"✓ Memilih {len(selected)} merchant aktif (exclude dilewati).")
            return selected

        if choice.lower().startswith("search "):
            query = choice[7:].strip()
            display_merchant_table(merchants, search_query=query, show_excluded=True)
            continue

        indices = parse_index_range(choice, len(merchants))
        if not indices:
            print("[!] Nomor yang Anda masukkan tidak valid. Silakan coba lagi.")
            continue

        selected = [merchants[i - 1] for i in indices]
        print(f"\n✓ Terpilih {len(selected)} merchant:")
        for idx, m in enumerate(selected, 1):
            print(f"   {idx}. {m.get('merchant_name')} (ID: {m.get('merchant_id')})")

        confirm = input("\nLanjutkan penarikan data? (Y/n): ").strip().lower()
        if confirm in ("", "y", "yes"):
            return selected


def main():
    parser = argparse.ArgumentParser(
        description="Shopee Food Outlet Info - CLI Filter & Puller",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Contoh Penggunaan:
  python cli.py -l                              # List semua merchant
  python cli.py -l -s ayam                      # Cari merchant bernama 'ayam'
  python cli.py -i 828310 1530883               # Tarik berdasarkan ID
  python cli.py -s "Ayam" "Seblak"              # Tarik semua yang ada kata Ayam atau Seblak
  python cli.py -n "AYAM GEPREK MAMA CHUBBY_"   # Tarik nama persis
  python cli.py --index 1,3,5-8                 # Tarik nomor urut tertentu
  python cli.py -I                              # Masuk ke menu interaktif
        """,
    )

    group_filter = parser.add_argument_group("Filter Merchant")
    group_filter.add_argument("-l", "--list", action="store_true", help="Tampilkan daftar merchant yang tersedia lalu keluar")
    group_filter.add_argument("-i", "--id", nargs="+", help="Satu atau lebih Merchant ID (contoh: -i 828310 1530883)")
    group_filter.add_argument("-n", "--name", nargs="+", help="Satu atau lebih Nama Merchant persis")
    group_filter.add_argument("-s", "--search", nargs="+", help="Cari merchant berdasarkan potongan nama/keyword (case-insensitive)")
    group_filter.add_argument("--index", help="Pilih merchant berdasarkan nomor index list (contoh: --index 1,3,5-10)")
    group_filter.add_argument("-I", "--interactive", action="store_true", help="Mode interaktif pilih nomor merchant")
    group_filter.add_argument("-a", "--all", action="store_true", help="Tarik semua merchant (abaikan filter)")

    group_ops = parser.add_argument_group("Opsi Eksekusi")
    group_ops.add_argument("--no-resume", "--fresh", action="store_true", help="Tarik fresh tanpa membaca/meresume Excel existing")
    group_ops.add_argument("--include-excluded", action="store_true", help="Sertakan merchant yang biasanya di-exclude (SuperFood, WonderFood, dll)")
    group_ops.add_argument("-o", "--output", help="Path file output Excel kustom")
    group_ops.add_argument("--dry-run", action="store_true", help="Tampilkan merchant yang terpilih tanpa menjalankan browser")

    args = parser.parse_args()

    # Load semua merchant yang ada dari master list
    all_available = get_merchants_to_switch()
    if not all_available:
        print("[!] Tidak ada data merchant ditemukan di merchant_list.json.")
        sys.exit(1)

    # 1. Mode List Saja
    if args.list:
        search_query = " ".join(args.search) if args.search else None
        display_merchant_table(all_available, search_query=search_query, show_excluded=args.include_excluded)
        return

    # 2. Filter Merchant
    selected_merchants = []

    # Mode Interaktif
    if args.interactive or (len(sys.argv) == 1):
        selected_merchants = interactive_select(all_available)

    # Mode Filter Argumen CLI
    elif args.id or args.name or args.search or args.index or args.all:
        target_ids = {str(x).strip() for x in args.id} if args.id else set()
        target_names = {x.strip().lower() for x in args.name} if args.name else set()
        search_terms = [x.strip().lower() for x in args.search] if args.search else []
        target_indices = set(parse_index_range(args.index, len(all_available))) if args.index else set()

        for idx, m in enumerate(all_available, 1):
            m_id = str(m.get("merchant_id") or "")
            m_name = m.get("merchant_name", "")
            m_name_lower = m_name.lower()

            matched = False

            if args.all:
                matched = True

            if target_ids and m_id in target_ids:
                matched = True

            if target_names and m_name_lower in target_names:
                matched = True

            if search_terms and any(term in m_name_lower for term in search_terms):
                matched = True

            if target_indices and idx in target_indices:
                matched = True

            if matched:
                if not args.include_excluded and is_excluded(m_name):
                    print(f"  [EXCLUDE] Melewati merchant ter-exclude: '{m_name}' (gunakan --include-excluded jika ingin memproses)")
                    continue
                selected_merchants.append(m)

    if not selected_merchants:
        print("[!] Tidak ada merchant yang cocok dengan filter yang ditentukan.")
        sys.exit(1)

    # Tampilkan preview terpilih
    print(f"\n[*] Terpilih {len(selected_merchants)} merchant:")
    for idx, m in enumerate(selected_merchants, 1):
        print(f"    {idx}. {m.get('merchant_name')} (ID: {m.get('merchant_id')}, Occ: {m.get('occurrence_index')})")
    print()

    # Dry run mode
    if args.dry_run:
        print("[DRY-RUN] Selesai pengecekan merchant terpilih. Tidak ada browser/penarikan yang dijalankan.")
        return

    # Jalankan proses penarikan data
    run_pull(
        target_merchants=selected_merchants,
        no_resume=args.no_resume,
        include_excluded=args.include_excluded,
        output_path=args.output,
    )


if __name__ == "__main__":
    main()
